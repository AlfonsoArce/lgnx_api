"""Query order status from LogiNext ShipmentApp/mile/v1/status.

What this script does
---------------------
Takes one or more shipment reference IDs on the command line and POSTs
them as a JSON array to the LogiNext "Order Status" endpoint. The full
response body is pretty-printed so the caller can inspect each order's
state, last-known location, ETA, and any error details.

Companion to
------------
``create_order_request.py`` -- after creating a shipment request, the
returned ``shipmentRequestReferenceId`` can be fed to this script to
poll for status.

Authentication
--------------
Reads ``LGNX_AUTH_TOKEN`` from ``.env``. The token is the value that
follows the literal ``BASIC `` prefix in the upstream curl's
``www-authenticate`` header; this script re-prepends ``BASIC `` itself.

Usage
-----
    uv run python get_order_status.py <reference_id> [<reference_id> ...]

Exit codes
----------
    0 -- API responded with a 2xx status.
    1 -- missing ``LGNX_AUTH_TOKEN`` or no reference ids supplied.
    2 -- API returned a non-2xx status.
"""

import json
import os
import sys

import requests
from dotenv import load_dotenv

# Documented v1 status-query endpoint.
URL = "https://api.loginextsolutions.com/ShipmentApp/mile/v1/status"

# Browser-style user-agent -- LogiNext's edge occasionally rejects the
# default python-requests UA during WAF checks.
USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36"
)


def main(reference_ids: list[str]) -> int:
    """Run the status-query flow.

    Steps performed
    ---------------
    1. Load environment variables from ``.env`` and validate that
       ``LGNX_AUTH_TOKEN`` is set.
    2. Validate that at least one reference id was provided.
    3. Build the request headers with the ``BASIC <token>`` scheme.
    4. POST the reference-id list as the JSON body to the status
       endpoint with a 30-second timeout.
    5. Print the HTTP status and the response body (pretty-printed JSON
       if parseable, otherwise raw text).
    6. Return 0 on a 2xx response, otherwise 2.

    Parameters
    ----------
    reference_ids : list[str]
        Shipment reference ids to query. Typically taken from
        ``sys.argv[1:]``.

    Returns
    -------
    int
        Process exit code.
    """
    # Step 1: load env and verify the auth token is present.
    load_dotenv()
    token = os.environ.get("LGNX_AUTH_TOKEN")
    if not token:
        print("error: LGNX_AUTH_TOKEN is missing (set it in .env)", file=sys.stderr)
        return 1

    # Step 2: ensure the caller actually passed at least one reference id;
    # without ids the request body would be ``[]`` and the API would
    # respond with an unhelpful error.
    if not reference_ids:
        print(
            "error: pass at least one order reference id as a CLI argument",
            file=sys.stderr,
        )
        return 1

    # Step 3: build the request headers. ``content-type`` is required
    # because we're sending a JSON body.
    headers = {
        "content-type": "application/json",
        "user-agent": USER_AGENT,
        "www-authenticate": f"BASIC {token}",
    }

    # Step 4: POST the list of reference ids. The API expects the body to
    # be a JSON array of strings, e.g. ``["REF-1", "REF-2"]``.
    response = requests.post(URL, headers=headers, json=reference_ids, timeout=30)

    # Step 5: print the result. JSON is pretty-printed when possible;
    # otherwise the raw body is shown so error pages remain debuggable.
    print(f"HTTP {response.status_code}")
    try:
        print(json.dumps(response.json(), indent=2))
    except ValueError:
        print(response.text)

    # Step 6: shell-friendly exit code.
    return 0 if response.ok else 2


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
