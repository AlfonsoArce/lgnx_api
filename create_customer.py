"""POST a customer to LogiNext ClientApp/customer/v1/create.

Mirrors references/LGNX_API/loginext_mile_apis/01_customer.md. Set
LGNX_AUTH_TOKEN in .env (the value after `BASIC ` in the curl's
www-authenticate header).
"""

import json
import os
import sys
import time

import requests
from dotenv import load_dotenv

URL = "https://api.loginextsolutions.com/ClientApp/customer/v1/create"

USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36"
)

ACCOUNT_CODE = f"cust-{int(time.time())}"

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
