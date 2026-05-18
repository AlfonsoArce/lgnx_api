"""Create a new customer in LogiNext via ClientApp/customer/v1/create.

What this script does
---------------------
Sends a single ``POST`` request to the LogiNext "Create Customer" endpoint
with a hard-coded sample payload. The payload includes account
identification (account code, name, contact channels) and a structured
billing address. A unique ``accountCode`` is generated at import time
using the current Unix timestamp so the script can be re-run without
hitting "duplicate account" errors during testing.

API reference
-------------
Mirrors ``references/LGNX_API/loginext_mile_apis/01_customer.md``.

Authentication
--------------
The script reads ``LGNX_AUTH_TOKEN`` from ``.env`` (via python-dotenv).
That value is the raw token that follows the literal ``BASIC `` prefix in
the ``www-authenticate`` header of the upstream curl example. The
script re-prepends ``BASIC `` itself when constructing the header.

Usage
-----
    uv run python create_customer.py

Exit codes
----------
    0 -- the API responded with a 2xx status.
    1 -- ``LGNX_AUTH_TOKEN`` is missing from the environment.
    2 -- the API responded with a non-2xx status (body still printed).
"""

import json
import os
import sys
import time

import requests
from dotenv import load_dotenv

# --- Endpoint configuration --------------------------------------------------
# LogiNext production endpoint for creating a customer record. The path
# ``ClientApp/customer/v1/create`` is the documented v1 create-customer route.
URL = "https://api.loginextsolutions.com/ClientApp/customer/v1/create"

# A realistic browser user-agent is used because LogiNext's edge sometimes
# rejects requests with a generic ``python-requests/x.y`` user-agent during
# WAF checks.
USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36"
)

# Generate a unique account code per run so re-running the script doesn't
# collide with an existing customer (the API rejects duplicate accountCode).
ACCOUNT_CODE = f"cust-{int(time.time())}"

# --- Sample request payload --------------------------------------------------
# The API expects a JSON ARRAY of customer objects -- batches of one or more
# customers can be created in a single call. Here we send a single record so
# the example is easy to follow and validate.
PAYLOAD = [
    {
        "accountCode": ACCOUNT_CODE,
        "name": "James Wan",
        "mobile": "341245673212",
        "whatsappOptin": "Y",
        "email": "james_test@hawkxinc.com",
        "customerType": "Preferred",
        "billingAddress": {
            "apartment": "Suite No. 1, Milsons Towers",
            "streetName": "Michigan Avenue",
            "landmark": "Opp. Subway",
            "locality": "Downtown Chicago",
            "city": "Chicago",
            "state": "IL",
            "country": "USA",
            "pincode": "10045",
            "latitude": 40.760838,
            "longitude": 79.555,
        },
    }
]


def main() -> int:
    """Execute the create-customer flow end-to-end.

    Steps performed
    ---------------
    1. Load environment variables from ``.env`` so ``LGNX_AUTH_TOKEN`` is
       available in ``os.environ``.
    2. Validate that the auth token is present; abort early with a clear
       error and exit code 1 if it is missing.
    3. Build the request headers, including the ``www-authenticate`` header
       with the ``BASIC <token>`` scheme that LogiNext expects.
    4. POST the ``PAYLOAD`` list to the create-customer endpoint with a
       30-second timeout to avoid hanging indefinitely.
    5. Print the HTTP status and the response body. If the response is not
       valid JSON, fall back to printing the raw text.
    6. Return 0 on a successful (2xx) response, otherwise 2.

    Returns
    -------
    int
        Process exit code (see module docstring for the meaning of each
        value).
    """
    # Step 1: pull configuration out of .env into the process environment.
    load_dotenv()

    # Step 2: ensure we actually have an auth token before issuing a request.
    token = os.environ.get("LGNX_AUTH_TOKEN")
    if not token:
        print("error: LGNX_AUTH_TOKEN is missing (set it in .env)", file=sys.stderr)
        return 1

    # Step 3: build the request headers. The LogiNext API uses a non-standard
    # ``www-authenticate`` request header (normally a *response* header) with
    # the ``BASIC <token>`` scheme.
    headers = {
        "content-type": "application/json",
        "user-agent": USER_AGENT,
        "www-authenticate": f"BASIC {token}",
    }

    # Step 4: send the request. ``timeout=30`` is generous for a single-record
    # create but protects the script from hanging if the upstream stalls.
    response = requests.post(URL, headers=headers, json=PAYLOAD, timeout=30)

    # Step 5: surface the result to the user. JSON is pretty-printed when
    # possible; otherwise the raw text body is shown so failures stay
    # debuggable even when the API returns HTML or an error page.
    print(f"HTTP {response.status_code}")
    try:
        print(json.dumps(response.json(), indent=2))
    except ValueError:
        print(response.text)

    # Step 6: translate HTTP success/failure into a shell-friendly exit code.
    return 0 if response.ok else 2


if __name__ == "__main__":
    sys.exit(main())
