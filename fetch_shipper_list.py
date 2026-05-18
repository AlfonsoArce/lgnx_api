"""Fetch every shipper across all pages from LogiNext ClientApp/shipper/list.

Mirrors references/curls/fetch_shipper_list.sh. Set LGNX_AUTH_TOKEN in .env
(the value after `BASIC ` in the curl's www-authenticate header).

Usage:
    uv run python fetch_shipper_list.py [--page-size 50] [--delay 1.0] [--out exports/shippers.xlsx]
"""

import argparse
import json
import os
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd
import requests
from dotenv import load_dotenv

EXPORTS_DIR = Path(__file__).parent / "exports"

URL = "https://products.loginextsolutions.com/ClientApp/shipper/list"

USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36"
)


def fetch_all(token: str, page_size: int, delay: float) -> list:
    headers = {
        "accept": "application/json, text/plain, */*",
        "user-agent": USER_AGENT,
        "www-authenticate": f"BASIC {token}",
    }

    all_shippers: list = []
    page = 1
    while True:
        params = {
            "pageNumber": page,
            "pageSize": page_size,
            "dataFetchMode": "DATA",
        }
        response = requests.post(URL, headers=headers, params=params, timeout=30)
        try:
            body = response.json()
        except ValueError:
            print(
                f"page {page}: HTTP {response.status_code} non-JSON response: "
                f"{response.text[:200]}",
                file=sys.stderr,
            )
            response.raise_for_status()
            break

        if not response.ok:
            print(f"page {page}: HTTP {response.status_code} {json.dumps(body)[:300]}", file=sys.stderr)
            response.raise_for_status()

        data = body.get("data") if isinstance(body, dict) else None
        shippers = (data or {}).get("results") or []
        total_count = (data or {}).get("totalCount")
        more = bool(body.get("moreResultsExists")) if isinstance(body, dict) else False

        all_shippers.extend(shippers)
        total_str = f"/{total_count}" if total_count is not None else ""
        print(
            f"page {page}: fetched {len(shippers)} shippers "
            f"(running total {len(all_shippers)}{total_str}, more={more})"
        )

        # The API often reports moreResultsExists=False / totalCount=0 even when
        # more pages exist, so fall back to "stop on a partial or empty page".
        if not shippers or (not more and len(shippers) < page_size):
            break

        page += 1
        time.sleep(delay)

    return all_shippers


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--page-size", type=int, default=50, help="results per page (default: 50)")
    parser.add_argument(
        "--delay",
        type=float,
        default=1.0,
        help="seconds to sleep between page requests (default: 1.0)",
    )
    parser.add_argument(
        "--out",
        default=None,
        help="output Excel file path (default: exports/shippers_list_<UTC-timestamp>.xlsx)",
    )
    args = parser.parse_args()

    load_dotenv()
    token = os.environ.get("LGNX_AUTH_TOKEN")
    if not token:
        print("error: LGNX_AUTH_TOKEN is missing (set it in .env)", file=sys.stderr)
        return 1

    shippers = fetch_all(token, args.page_size, args.delay)

    if args.out:
        out_path = Path(args.out)
    else:
        EXPORTS_DIR.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
        out_path = EXPORTS_DIR / f"shippers_list_{stamp}.xlsx"
    out_path.parent.mkdir(parents=True, exist_ok=True)

    df = pd.json_normalize(shippers)
    df.to_excel(out_path, index=False, sheet_name="shippers")

    print(f"wrote {len(df)} shippers ({len(df.columns)} columns) to {out_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
