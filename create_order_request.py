"""POST a shipment-request to LogiNext middlemile/v1/create.

Mirrors references/curls/create_order_request.sh. Set LGNX_AUTH_TOKEN in .env
(the value after `BASIC ` in the curl's www-authenticate header).
"""

import json
import os
import sys
import time
from datetime import datetime, timedelta, timezone

import requests
from dotenv import load_dotenv

URL = "https://api.loginextsolutions.com/BookingApp/middlemile/v1/create"

USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36"
)

def _format_window(dt: datetime) -> str:
    # Format "2026-04-29T20:00:00.000Z"
    return dt.strftime("%Y-%m-%dT%H:%M:%S.") + f"{dt.microsecond // 1000:03d}Z"


_NOW = datetime.now(timezone.utc)
_PICKUP_START = _NOW + timedelta(hours=1)
_PICKUP_END = _PICKUP_START + timedelta(hours=1)
_DELIVER_START = _PICKUP_END + timedelta(hours=4)
_DELIVER_END = _DELIVER_START + timedelta(hours=1)
_RETURN_START = _DELIVER_END + timedelta(minutes=5)
_RETURN_END = _RETURN_START + timedelta(minutes=5)

SHIPMENT_REQUEST_NO = f"Test-{int(time.time())}"
PICKUP_START_TIME_WINDOW = _format_window(_PICKUP_START)
PICKUP_END_TIME_WINDOW = _format_window(_PICKUP_END)
DELIVER_START_TIME_WINDOW = _format_window(_DELIVER_START)
DELIVER_END_TIME_WINDOW = _format_window(_DELIVER_END)
RETURN_START_TIME_WINDOW = _format_window(_RETURN_START)
RETURN_END_TIME_WINDOW = _format_window(_RETURN_END)

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
    load_dotenv()
    token = os.environ.get("LGNX_AUTH_TOKEN")
    if not token:
        print("error: LGNX_AUTH_TOKEN is missing (set it in .env)", file=sys.stderr)
        return 1

    headers = {
        "content-type": "application/json",
        "user-agent": USER_AGENT,
        "www-authenticate": f"BASIC {token}",
    }

    response = requests.post(URL, headers=headers, json=PAYLOAD, timeout=30)

    print(f"HTTP {response.status_code}")
    try:
        print(json.dumps(response.json(), indent=2))
    except ValueError:
        print(response.text)

    return 0 if response.ok else 2


if __name__ == "__main__":
    sys.exit(main())
