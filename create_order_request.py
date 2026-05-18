"""Submit a middle-mile shipment-request to LogiNext.

What this script does
---------------------
Sends a single ``POST`` to
``https://api.loginextsolutions.com/BookingApp/middlemile/v1/create``
with a fully-populated sample shipment request -- pickup, delivery, and
return legs, plus a single crate. Designed as a smoke-test harness for
the middle-mile booking flow.

After receiving the response the script:
    * pretty-prints the response body to stdout
    * appends a structured JSON log line to ``create_order_request.log``
      containing the returned ``shipmentRequestReferenceId`` so the
      created shipment can be followed up later (e.g. via
      ``get_order_status.py``)

Time windows
------------
All pickup / deliver / return windows are computed relative to "now"
(UTC) at import time so each run produces a valid future schedule:
    * pickup window starts 1h from now, lasts 1h
    * deliver window starts 4h after pickup ends, lasts 1h
    * return window starts 5m after deliver ends, lasts 5m
Times are emitted in ``YYYY-MM-DDTHH:MM:SS.sssZ`` form, which is what
the LogiNext booking endpoint accepts.

Authentication
--------------
``LGNX_AUTH_TOKEN`` is loaded from ``.env``. It must be the value AFTER
the literal ``BASIC `` prefix in the upstream curl example.

API reference
-------------
Mirrors ``references/curls/create_order_request.sh``.

Usage
-----
    uv run python create_order_request.py

Exit codes
----------
    0 -- 2xx response.
    1 -- missing ``LGNX_AUTH_TOKEN``.
    2 -- non-2xx response.
"""

import json
import os
import sys
import time
from datetime import datetime, timedelta, timezone

import requests
from dotenv import load_dotenv

# Documented v1 middle-mile create endpoint.
URL = "https://api.loginextsolutions.com/BookingApp/middlemile/v1/create"

# Audit-log file: one JSON object per line containing the request's
# resulting reference id and HTTP status, suitable for later grep/jq.
LOG_PATH = "create_order_request.log"

# Browser-style user-agent -- LogiNext's edge occasionally rejects the
# default python-requests UA during WAF checks.
USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36"
)


def _format_window(dt: datetime) -> str:
    """Format a UTC ``datetime`` as a LogiNext time-window string.

    LogiNext expects timestamps in the form ``2026-04-29T20:00:00.000Z``
    -- ISO-8601 with millisecond precision and a trailing ``Z`` (Zulu/UTC).
    Python's ``isoformat()`` defaults to microseconds (6 digits) and does
    not append the ``Z`` suffix, so we format manually.

    Parameters
    ----------
    dt : datetime
        Timezone-aware (or naive UTC) datetime to format.

    Returns
    -------
    str
        ``YYYY-MM-DDTHH:MM:SS.mmmZ`` representation, where ``mmm`` is
        the millisecond component derived from ``dt.microsecond``.
    """
    # ``dt.microsecond // 1000`` truncates microseconds (6 digits) to
    # milliseconds (3 digits) to match the API's expected resolution.
    return dt.strftime("%Y-%m-%dT%H:%M:%S.") + f"{dt.microsecond // 1000:03d}Z"


# --- Compute the time-window schedule, anchored to "now" at import time. ----
# Anchoring at import time means re-running the script always produces a
# fresh, future schedule without needing CLI flags. The deltas are chosen
# to be small but realistic for an end-to-end middle-mile trip.
_NOW = datetime.now(timezone.utc)
_PICKUP_START = _NOW + timedelta(hours=1)
_PICKUP_END = _PICKUP_START + timedelta(hours=1)
_DELIVER_START = _PICKUP_END + timedelta(hours=4)
_DELIVER_END = _DELIVER_START + timedelta(hours=1)
_RETURN_START = _DELIVER_END + timedelta(minutes=5)
_RETURN_END = _RETURN_START + timedelta(minutes=5)

# Unique request id, timestamp-based so re-runs don't collide on the
# server side (which rejects duplicate ``shipmentRequestNo`` values).
SHIPMENT_REQUEST_NO = f"Test-{int(time.time())}"

# Pre-format the time-window strings so the PAYLOAD literal below is easy
# to scan (no inline function calls inside a 60+ field dict).
PICKUP_START_TIME_WINDOW = _format_window(_PICKUP_START)
PICKUP_END_TIME_WINDOW = _format_window(_PICKUP_END)
DELIVER_START_TIME_WINDOW = _format_window(_DELIVER_START)
DELIVER_END_TIME_WINDOW = _format_window(_DELIVER_END)
RETURN_START_TIME_WINDOW = _format_window(_RETURN_START)
RETURN_END_TIME_WINDOW = _format_window(_RETURN_END)

# --- Sample request payload --------------------------------------------------
# The endpoint accepts a JSON ARRAY of shipment requests so multiple legs
# can be booked in one call. Here we send a single trip:
#   * Pickup from PAZ AVIATION INC in HIALEAH, FL
#   * Deliver to TEG (also in HIALEAH, FL)
#   * Return to the original pickup location
# The crate is a single 25.4 x 25.4 x 25.4 cm box weighing 8.16 kg.
PAYLOAD = [
    {
        "shipmentRequestNo": SHIPMENT_REQUEST_NO,
        "shipmentRequestType": "DELIVER",
        "returnAllowedFl": "Y",
        "cancellationAllowedFl": "Y",
        "clientCode": "BARFIELD INC",
        "distributionCenter": "HawkExpress",
        "shipmentCrateMappings": [
            {
                "crateCd": "1",
                "noOfUnits": 1,
                "crateAmount": 0,
                "crateType": "Box",
                "crateWeight": 8.1647,
                "crateVolume": 0.0000,
                "crateLength": 25.4000,
                "crateBreadth": 25.4000,
                "crateHeight": 25.4000,
            }
        ],
        "deliveryType": "Wrapping per pallet (A)",
        "numberOfItems": 1,
        # --- Pickup leg ----------------------------------------------------
        "pickupAccountName": "PAZ AVIATION INC",
        "pickupAccountCode": "PAZ AVIATION INC 1",
        "pickupEmail": None,
        "pickupPhoneNumber": None,
        "pickupCountry": "USA",
        "pickupPinCode": "33014",
        "pickupState": "FL",
        "pickupApartment": None,
        "pickupCity": "HIALEAH",
        "pickupStartTimeWindow": PICKUP_START_TIME_WINDOW,
        "pickupEndTimeWindow": PICKUP_END_TIME_WINDOW,
        "pickupStreetName": "7455 W 2nd Ct, HIALEAH, FL 33014, USA",
        "pickupNotes": None,
        # --- Delivery leg --------------------------------------------------
        "deliverAccountName": "TEG",
        "deliverAccountCode": "TEG 1",
        "deliverEmail": None,
        "deliverPhoneNumber": None,
        "deliverCountry": "USA",
        "deliverState": "FL",
        "deliverPinCode": "33013",
        "deliverApartment": None,
        "deliverCity": "HIALEAH",
        "deliverAddressTimezone": "America/New_York",
        "deliverServiceTime": 0,
        "deliverStartTimeWindow": DELIVER_START_TIME_WINDOW,
        "deliverEndTimeWindow": DELIVER_END_TIME_WINDOW,
        "deliverStreetName": "4747 E 10th Ave, HIALEAH, FL 33013, USA",
        "deliverNotes": None,
        # --- Return leg (back to pickup origin) ----------------------------
        "returnAccountName": "PAZ AVIATION INC",
        "returnAccountCode": "PAZ AVIATION INC 1",
        "returnEmail": None,
        "returnPhoneNumber": None,
        "returnCountry": "USA",
        "returnPinCode": "33014",
        "returnState": "FL",
        "returnBranch": "HawkExpress",
        "returnApartment": None,
        "returnCity": "HIALEAH",
        "returnTimezone": "America/New_York",
        "returnStartTimeWindow": RETURN_START_TIME_WINDOW,
        "returnEndTimeWindow": RETURN_END_TIME_WINDOW,
        "returnLandmark": None,
        "returnLocality": None,
        "returnStreetName": "7455 W 2nd Ct, HIALEAH, FL 33014, USA",
        "serviceType": "CARGO VAN STD",
    }
]


def main() -> int:
    """Execute the shipment-request creation flow.

    Steps performed
    ---------------
    1. Load ``.env`` and validate ``LGNX_AUTH_TOKEN`` is present.
    2. Build the request headers (content-type + auth + UA).
    3. POST ``PAYLOAD`` to the middle-mile create endpoint with a
       30-second timeout.
    4. Parse the response body. If it is not valid JSON, fall back to
       the raw text for display + logging.
    5. Print the HTTP status and pretty-printed body to stdout.
    6. Extract the first record from ``body["data"]`` (the API returns
       one result per submitted shipment) to capture the server-assigned
       ``shipmentRequestReferenceId``.
    7. Append a JSON log entry to ``LOG_PATH`` with the timestamp,
       status code, returned reference id, and the request number used.
    8. Return 0 on a 2xx response, otherwise 2.

    Returns
    -------
    int
        Process exit code.
    """
    # Step 1: configuration.
    load_dotenv()
    token = os.environ.get("LGNX_AUTH_TOKEN")
    if not token:
        print("error: LGNX_AUTH_TOKEN is missing (set it in .env)", file=sys.stderr)
        return 1

    # Step 2: assemble request headers.
    headers = {
        "content-type": "application/json",
        "user-agent": USER_AGENT,
        "www-authenticate": f"BASIC {token}",
    }

    # Step 3: send the booking request.
    response = requests.post(URL, headers=headers, json=PAYLOAD, timeout=30)

    # Step 4: best-effort JSON parse. Both pretty-printed JSON and raw
    # text fallbacks are supported so error pages remain debuggable.
    try:
        body = response.json()
        body_text = json.dumps(body, indent=2)
    except ValueError:
        body = None
        body_text = response.text

    # Step 5: display the response to the user.
    print(f"HTTP {response.status_code}")
    print(body_text)

    # Step 6: extract the first ``data`` element. LogiNext returns an
    # array even when a single shipment was submitted; we only need the
    # first record for our single-shipment payload.
    first = (body or {}).get("data", [{}])[0] if isinstance(body, dict) else {}

    # Step 7: append the audit log line. ``open(..., "a")`` is safe for
    # concurrent appends on POSIX as long as each line fits in a single
    # write (which JSON lines + newline always do here).
    log_entry = {
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "status_code": response.status_code,
        "shipmentRequestReferenceId": first.get("shipmentRequestReferenceId"),
        "shipmentRequestNo": first.get("shipmentRequestNo", SHIPMENT_REQUEST_NO),
    }
    with open(LOG_PATH, "a") as f:
        f.write(json.dumps(log_entry) + "\n")

    # Step 8: shell-friendly exit code.
    return 0 if response.ok else 2


if __name__ == "__main__":
    sys.exit(main())
