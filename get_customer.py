"""GET customer details from LogiNext ClientApp/customer/v1/get/list.

Mirrors references/LGNX_API/loginext_mile_apis/01_customer.md. Set
LGNX_AUTH_TOKEN in .env (the value after `BASIC ` in the curl's
www-authenticate header).

Accepts customer account codes or reference IDs (up to 20 per call) and
exports the resulting customer details to an Excel file under exports/.

Usage:
    uv run python get_customer.py <id> [<id> ...] [--out path.xlsx]
"""

import argparse
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd
import requests
from dotenv import load_dotenv

EXPORTS_DIR = Path(__file__).parent / "exports"

URL = "https://api.loginextsolutions.com/ClientApp/customer/v1/get/list"

MAX_IDS_PER_CALL = 20

USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36"
)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "ids",
        nargs="+",
        help="customer accountCode or referenceId values (up to 20)",
    )
    parser.add_argument(
        "--out",
        default=None,
        help="output Excel file path (default: exports/customer_<UTC-timestamp>.xlsx)",
    )
    args = parser.parse_args()

    load_dotenv()
    token = os.environ.get("LGNX_AUTH_TOKEN")
    if not token:
        print("error: LGNX_AUTH_TOKEN is missing (set it in .env)", file=sys.stderr)
        return 1

    if len(args.ids) > MAX_IDS_PER_CALL:
        print(
            f"error: this API accepts up to {MAX_IDS_PER_CALL} ids per call "
            f"(got {len(args.ids)})",
            file=sys.stderr,
        )
        return 1

    headers = {
        "user-agent": USER_AGENT,
        "www-authenticate": f"BASIC {token}",
    }

    response = requests.get(
        URL,
        headers=headers,
        params={"ids": ",".join(args.ids)},
        timeout=30,
    )

    print(f"HTTP {response.status_code}")
    try:
        body = response.json()
    except ValueError:
        print(response.text)
        return 2

    print(json.dumps(body, indent=2))

    if not response.ok:
        return 2

    customers = body.get("data") if isinstance(body, dict) else None
    if not customers:
        print("no customer records returned; skipping Excel export", file=sys.stderr)
        return 0

    if args.out:
        out_path = Path(args.out)
    else:
        EXPORTS_DIR.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
        out_path = EXPORTS_DIR / f"customer_{stamp}.xlsx"
    out_path.parent.mkdir(parents=True, exist_ok=True)

    df = pd.json_normalize(customers)
    df.to_excel(out_path, index=False, sheet_name="customers")

    print(f"wrote {len(df)} customers ({len(df.columns)} columns) to {out_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
