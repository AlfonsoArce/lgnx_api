"""POST to LogiNext mile/v1/status to retrieve the status of one or more orders.

Mirrors the pattern of create_order_request.py. Set LGNX_AUTH_TOKEN in .env.

Usage:
    uv run python get_order_status.py <reference_id> [<reference_id> ...]
"""

import json
import os
import sys

import requests
from dotenv import load_dotenv

URL = "https://api.loginextsolutions.com/ShipmentApp/mile/v1/status"

USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36"
)


def main(reference_ids: list[str]) -> int:
    load_dotenv()
    token = os.environ.get("LGNX_AUTH_TOKEN")
    if not token:
        print("error: LGNX_AUTH_TOKEN is missing (set it in .env)", file=sys.stderr)
        return 1

    if not reference_ids:
        print(
            "error: pass at least one order reference id as a CLI argument",
            file=sys.stderr,
        )
        return 1

    headers = {
        "content-type": "application/json",
        "user-agent": USER_AGENT,
        "www-authenticate": f"BASIC {token}",
    }

    response = requests.post(URL, headers=headers, json=reference_ids, timeout=30)

    print(f"HTTP {response.status_code}")
    try:
        print(json.dumps(response.json(), indent=2))
    except ValueError:
        print(response.text)

    return 0 if response.ok else 2


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
