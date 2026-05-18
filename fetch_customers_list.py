"""Page through every customer in LogiNext ClientApp/customer/list/data.

What this script does
---------------------
Walks the LogiNext "customer list" endpoint page by page, collects every
record returned, normalizes the nested JSON into a flat tabular form,
and writes the result to an Excel workbook under ``output/``.

Comparison to ``get_customer.py``
---------------------------------
``get_customer.py`` fetches up to 20 specific customers by id. This
script enumerates the **entire** customer dataset by walking pages.

Pagination strategy
-------------------
Identical to the other ``fetch_*_list.py`` scripts: the API's
``moreResultsExists`` / ``totalCount`` fields are unreliable, so we keep
paginating until a page returns either zero records or a partial batch
combined with the API's own ``more=False`` signal.

Method
------
This endpoint expects a ``POST`` with the paging parameters as query
string values (the body is empty). Other ``GET``-shaped LogiNext list
endpoints exist; the difference comes straight from the upstream curl
example and is required.

Rate limits
-----------
``--delay`` (default 1.0s) introduces a small sleep between page
requests.

Authentication
--------------
``LGNX_AUTH_TOKEN`` must be set in ``.env`` (the value AFTER ``BASIC ``
in the upstream curl's ``www-authenticate`` header).

API reference
-------------
Mirrors ``references/curls/fetch_customers_list.sh``.

Usage
-----
    uv run python fetch_customers_list.py \\
        [--page-size 50] [--delay 1.0] [--out output/customers.xlsx]

Exit codes
----------
    0 -- success.
    1 -- missing ``LGNX_AUTH_TOKEN``.
    Other -- ``requests`` raises on HTTP errors; the traceback exits non-zero.
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

# Default location for Excel exports.
OUTPUT_DIR = Path(__file__).parent / "output"

# NOTE: the customer-list endpoint lives on ``products.loginextsolutions.com``
# (not ``api.``). Confirmed via the reference curl.
URL = "https://products.loginextsolutions.com/ClientApp/customer/list/data"

# Browser-style user-agent -- LogiNext's edge occasionally rejects the
# default python-requests UA during WAF checks.
USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36"
)


def fetch_all(token: str, page_size: int, delay: float) -> list:
    """Paginate through every customer and return the combined list.

    Steps performed (per page)
    --------------------------
    1. Build per-page query parameters: ``dataFetchMode=DATA``,
       ``pageNumber``, ``pageSize``.
    2. POST to the customer list endpoint with a 30s timeout. The body
       is empty; paging values are on the query string.
    3. Parse JSON. Bail loudly on parse errors.
    4. Extract ``body["data"]["results"]`` along with the paging
       metadata (``totalCount``, ``moreResultsExists``).
    5. Append records to the accumulator and print progress.
    6. Decide whether to continue: stop on an empty page, or on a
       partial page when the API also reports no more results.
    7. Sleep ``delay`` seconds before requesting the next page.

    Parameters
    ----------
    token : str
        Raw ``LGNX_AUTH_TOKEN`` value. ``BASIC `` is prepended here.
    page_size : int
        Records requested per page.
    delay : float
        Politeness sleep between page requests.

    Returns
    -------
    list
        All customer records, concatenated in page order.

    Raises
    ------
    requests.HTTPError
        Propagated from ``response.raise_for_status()`` on non-2xx
        responses or non-JSON payloads.
    """
    # Headers reused across every page request.
    headers = {
        "accept": "application/json, text/plain, */*",
        "user-agent": USER_AGENT,
        "www-authenticate": f"BASIC {token}",
    }

    all_customers: list = []
    page = 1
    while True:
        # Step 1: per-page query params.
        params = {
            "dataFetchMode": "DATA",
            "pageNumber": page,
            "pageSize": page_size,
        }

        # Step 2: POST with paging as query params (no body).
        response = requests.post(URL, headers=headers, params=params, timeout=30)

        # Step 3: parse JSON; bail loudly on parse errors.
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

        # Surface HTTP failures with a body snippet for debuggability.
        if not response.ok:
            print(f"page {page}: HTTP {response.status_code} {json.dumps(body)[:300]}", file=sys.stderr)
            response.raise_for_status()

        # Step 4: pull records + paging metadata out of the envelope.
        data = body.get("data") if isinstance(body, dict) else None
        customers = (data or {}).get("results") or []
        total_count = (data or {}).get("totalCount")
        more = bool(body.get("moreResultsExists")) if isinstance(body, dict) else False

        # Step 5: accumulate and report progress.
        all_customers.extend(customers)
        total_str = f"/{total_count}" if total_count is not None else ""
        print(
            f"page {page}: fetched {len(customers)} customers "
            f"(running total {len(all_customers)}{total_str}, more={more})"
        )

        # Step 6: termination heuristic. The API often misreports its
        # paging fields, so we treat a partial page combined with
        # ``more=False`` (or an empty page) as the end of the dataset.
        if not customers or (not more and len(customers) < page_size):
            break

        # Step 7: politeness sleep, then advance.
        page += 1
        time.sleep(delay)

    return all_customers


def main() -> int:
    """CLI entry point: parse args, fetch all customers, write Excel.

    Steps performed
    ---------------
    1. Parse ``--page-size``, ``--delay``, ``--out``.
    2. Load ``.env`` and verify ``LGNX_AUTH_TOKEN``.
    3. Drive the paginator.
    4. Resolve the output path (explicit or timestamped default).
    5. Flatten with ``pandas.json_normalize`` and write to ``.xlsx``.

    Returns
    -------
    int
        Process exit code.
    """
    # Step 1: argument parsing.
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
        help="output Excel file path (default: output/customers_list_<UTC-timestamp>.xlsx)",
    )
    args = parser.parse_args()

    # Step 2: load env + verify token.
    load_dotenv()
    token = os.environ.get("LGNX_AUTH_TOKEN")
    if not token:
        print("error: LGNX_AUTH_TOKEN is missing (set it in .env)", file=sys.stderr)
        return 1

    # Step 3: drive the paginator.
    customers = fetch_all(token, args.page_size, args.delay)

    # Step 4: resolve the output path; timestamp default names so
    # successive runs do not overwrite one another.
    if args.out:
        out_path = Path(args.out)
    else:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
        out_path = OUTPUT_DIR / f"customers_list_{stamp}.xlsx"
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # Step 5: flatten nested JSON (e.g. ``billingAddress.city``) and write.
    df = pd.json_normalize(customers)
    df.to_excel(out_path, index=False, sheet_name="customers")

    print(f"wrote {len(df)} customers ({len(df.columns)} columns) to {out_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
