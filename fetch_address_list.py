"""Page through every address in LogiNext ClientApp/clientnode/get.

What this script does
---------------------
Walks the LogiNext "client node" (address) list endpoint page by page,
accumulates every record returned, normalizes the nested JSON into a
flat tabular form, and writes the result to an Excel workbook under
``output/``.

Pagination strategy
-------------------
The API exposes ``moreResultsExists`` and ``totalCount`` fields in its
envelope, but both have been observed to misreport (e.g.
``moreResultsExists=False`` while more pages clearly exist). To stay
safe we keep paginating as long as a page returns a *full* batch
(``len(results) == page_size``) and only stop when:
    * the page is empty, OR
    * the page is a partial batch *and* the API reports no more results.

Rate limits
-----------
``--delay`` (default 1.0s) introduces a small sleep between page
requests to stay under any per-second rate limit and to be a polite
client.

Authentication
--------------
``LGNX_AUTH_TOKEN`` must be set in ``.env``. It is the value AFTER
``BASIC `` in the upstream curl's ``www-authenticate`` header; the
script re-prepends ``BASIC `` itself.

API reference
-------------
Mirrors ``references/curls/fetch_address_list.sh``.

Usage
-----
    uv run python fetch_address_list.py \\
        [--page-size 50] [--delay 1.0] [--out output/addresses.xlsx]

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

# Default location for Excel exports when ``--out`` is not provided.
OUTPUT_DIR = Path(__file__).parent / "output"

# NOTE: this endpoint lives on the ``products.loginextsolutions.com`` host
# (not the ``api.`` host used by other endpoints in this repo). The path
# ``ClientApp/clientnode/get`` is LogiNext's term for an "address" record.
URL = "https://products.loginextsolutions.com/ClientApp/clientnode/get"

# Browser-style user-agent -- LogiNext's edge occasionally rejects the
# default python-requests UA during WAF checks.
USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36"
)


def fetch_all(token: str, page_size: int, delay: float) -> list:
    """Fetch every address by walking pages until exhaustion.

    Steps performed (per page)
    --------------------------
    1. Build the ``GET`` query parameters: ``dataFetchMode=DATA`` (return
       records, not just counts), ``pageNumber`` (1-indexed), ``pageSize``.
    2. Send the request with a 30-second timeout.
    3. Parse the response JSON. If parsing fails, print a diagnostic and
       raise so the script aborts loudly rather than silently looping.
    4. Extract ``body["data"]["results"]`` -- the address array -- along
       with ``totalCount`` and ``moreResultsExists`` for logging /
       termination heuristics.
    5. Append the page's records to the accumulator and print progress.
    6. Decide whether to continue: stop on an empty page or on a partial
       page when the API also signals no more results.
    7. Sleep ``delay`` seconds between requests to stay polite.

    Parameters
    ----------
    token : str
        Raw value of the ``LGNX_AUTH_TOKEN`` env var. The ``BASIC `` prefix
        is added inside this function.
    page_size : int
        Number of records requested per page. Larger pages reduce total
        round-trips but raise the cost of a single retry.
    delay : float
        Seconds to sleep between successive page requests.

    Returns
    -------
    list
        All address records concatenated in page order.

    Raises
    ------
    requests.HTTPError
        Propagated from ``response.raise_for_status()`` on non-2xx
        responses or non-JSON payloads.
    """
    # Headers used for every page request -- declared once outside the loop.
    headers = {
        "accept": "application/json, text/plain, */*",
        "user-agent": USER_AGENT,
        "www-authenticate": f"BASIC {token}",
    }

    all_addresses: list = []
    page = 1
    while True:
        # Step 1: assemble the per-page query params.
        params = {
            "dataFetchMode": "DATA",
            "pageNumber": page,
            "pageSize": page_size,
        }

        # Step 2: issue the request.
        response = requests.get(URL, headers=headers, params=params, timeout=30)

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

        # Raise on HTTP-level failures (4xx/5xx) -- but only after we've
        # logged a snippet of the body to make the error actionable.
        if not response.ok:
            print(f"page {page}: HTTP {response.status_code} {json.dumps(body)[:300]}", file=sys.stderr)
            response.raise_for_status()

        # Step 4: pull records + paging metadata out of the envelope.
        data = body.get("data") if isinstance(body, dict) else None
        addresses = (data or {}).get("results") or []
        total_count = (data or {}).get("totalCount")
        more = bool(body.get("moreResultsExists")) if isinstance(body, dict) else False

        # Step 5: accumulate and report progress.
        all_addresses.extend(addresses)
        total_str = f"/{total_count}" if total_count is not None else ""
        print(
            f"page {page}: fetched {len(addresses)} addresses "
            f"(running total {len(all_addresses)}{total_str}, more={more})"
        )

        # Step 6: termination heuristic. The API often reports
        # ``moreResultsExists=False`` / ``totalCount=0`` even when more
        # pages exist, so fall back to "stop on a partial or empty page":
        # a full page suggests there could be more, while a partial page
        # combined with the API's own ``more=False`` is a strong signal
        # that we're at the end.
        if not addresses or (not more and len(addresses) < page_size):
            break

        # Step 7: politeness sleep, then advance.
        page += 1
        time.sleep(delay)

    return all_addresses


def main() -> int:
    """CLI entry point: parse args, fetch everything, write Excel.

    Steps performed
    ---------------
    1. Parse ``--page-size``, ``--delay``, ``--out`` from argv.
    2. Load ``.env`` and validate ``LGNX_AUTH_TOKEN``.
    3. Drive the paginator (`fetch_all`) to collect every address.
    4. Resolve the output path (explicit or timestamped default).
    5. Flatten and write to Excel via pandas.

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
        help="output Excel file path (default: output/addresses_<UTC-timestamp>.xlsx)",
    )
    args = parser.parse_args()

    # Step 2: load env + verify token.
    load_dotenv()
    token = os.environ.get("LGNX_AUTH_TOKEN")
    if not token:
        print("error: LGNX_AUTH_TOKEN is missing (set it in .env)", file=sys.stderr)
        return 1

    # Step 3: drive the paginator.
    addresses = fetch_all(token, args.page_size, args.delay)

    # Step 4: resolve the output path. Auto-generated names are stamped
    # with the UTC time so successive runs are uniquely identifiable.
    if args.out:
        out_path = Path(args.out)
    else:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
        out_path = OUTPUT_DIR / f"addresses_{stamp}.xlsx"
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # Step 5: flatten nested JSON (e.g. ``location.lat``) and write.
    df = pd.json_normalize(addresses)
    df.to_excel(out_path, index=False, sheet_name="addresses")

    print(f"wrote {len(df)} addresses ({len(df.columns)} columns) to {out_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
