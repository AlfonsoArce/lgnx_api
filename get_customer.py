"""Fetch customer details from LogiNext ClientApp/customer/v1/get/list.

What this script does
---------------------
Given one or more customer identifiers (``accountCode`` or
``referenceId``), this script calls the LogiNext "Get Customer" endpoint,
pretty-prints the JSON response, and exports the returned customer
records to an Excel workbook under the ``output/`` directory.

Up to ``MAX_IDS_PER_CALL`` (20) ids may be passed in a single invocation
to match the API's documented per-call limit.

API reference
-------------
Mirrors ``references/LGNX_API/loginext_mile_apis/01_customer.md``.

Authentication
--------------
Reads ``LGNX_AUTH_TOKEN`` from ``.env`` (the value after ``BASIC `` in
the upstream curl's ``www-authenticate`` header). The script re-prepends
``BASIC `` when constructing the request header.

Usage
-----
    uv run python get_customer.py <id> [<id> ...] [--out path.xlsx]

Examples
--------
    uv run python get_customer.py cust-1715814123
    uv run python get_customer.py CUST-A CUST-B --out output/my_customers.xlsx

Exit codes
----------
    0 -- success (records exported or no records found).
    1 -- configuration error (missing token, too many ids passed in).
    2 -- HTTP error or non-JSON response from the API.
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

# Where Excel exports are written when ``--out`` is not provided.
OUTPUT_DIR = Path(__file__).parent / "output"

# Documented v1 "get customer by list of ids" endpoint.
URL = "https://api.loginextsolutions.com/ClientApp/customer/v1/get/list"

# LogiNext caps this endpoint at 20 ids per call -- enforce client-side to
# avoid a confusing 400 from the server.
MAX_IDS_PER_CALL = 20

# Browser-style user-agent -- LogiNext's edge sometimes rejects the default
# ``python-requests`` user-agent during WAF checks.
USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36"
)


def main() -> int:
    """Run the get-customer command-line flow.

    Steps performed
    ---------------
    1. Parse CLI arguments (positional ids and optional ``--out``).
    2. Load ``LGNX_AUTH_TOKEN`` from ``.env`` and verify it is present.
    3. Reject the call early if more than ``MAX_IDS_PER_CALL`` ids were
       passed (the API would reject the batch).
    4. Build the request headers and issue a GET with ``ids`` joined as a
       comma-separated query parameter.
    5. Pretty-print the response body. Bail out with exit code 2 if the
       body is not JSON or the status is non-2xx.
    6. Extract the ``data`` array (the list of customer records). If it
       is empty, skip the Excel export and return 0.
    7. Resolve the output path -- explicit ``--out`` value or an
       auto-generated timestamped file under ``output/``.
    8. Flatten the nested JSON into a DataFrame with
       ``pandas.json_normalize`` and write to ``.xlsx``.

    Returns
    -------
    int
        Process exit code (see module docstring for the meaning of each
        value).
    """
    # Step 1: parse CLI arguments. ``__doc__`` is reused as the help banner
    # so ``--help`` mirrors this module's docstring.
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "ids",
        nargs="+",
        help="customer accountCode or referenceId values (up to 20)",
    )
    parser.add_argument(
        "--out",
        default=None,
        help="output Excel file path (default: output/customer_<UTC-timestamp>.xlsx)",
    )
    args = parser.parse_args()

    # Step 2: load credentials and abort if the token is missing.
    load_dotenv()
    token = os.environ.get("LGNX_AUTH_TOKEN")
    if not token:
        print("error: LGNX_AUTH_TOKEN is missing (set it in .env)", file=sys.stderr)
        return 1

    # Step 3: enforce the documented per-call cap client-side so the user
    # gets a clear local error rather than a generic 400.
    if len(args.ids) > MAX_IDS_PER_CALL:
        print(
            f"error: this API accepts up to {MAX_IDS_PER_CALL} ids per call "
            f"(got {len(args.ids)})",
            file=sys.stderr,
        )
        return 1

    # Step 4: build headers and send the GET. Ids are passed as a
    # comma-separated string under the ``ids`` query parameter, matching
    # the curl example in the reference docs.
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

    # Step 5: surface status, then parse JSON. Non-JSON responses are
    # printed as raw text and treated as a hard failure.
    print(f"HTTP {response.status_code}")
    try:
        body = response.json()
    except ValueError:
        print(response.text)
        return 2

    print(json.dumps(body, indent=2))

    if not response.ok:
        return 2

    # Step 6: extract the customer records. The LogiNext convention is
    # ``{"data": [...records...], ...}``. Bail out gracefully if no
    # records were returned -- no Excel file is created.
    customers = body.get("data") if isinstance(body, dict) else None
    if not customers:
        print("no customer records returned; skipping Excel export", file=sys.stderr)
        return 0

    # Step 7: resolve the output path. If the user passed ``--out`` we
    # honor it verbatim; otherwise we synthesize a timestamped filename
    # under ``output/`` so each run is uniquely identifiable.
    if args.out:
        out_path = Path(args.out)
    else:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
        out_path = OUTPUT_DIR / f"customer_{stamp}.xlsx"
    # Ensure the parent directory exists for user-supplied paths too.
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # Step 8: flatten nested JSON (e.g. ``billingAddress.city``) into a
    # tabular form and write to Excel. ``index=False`` drops the synthetic
    # RangeIndex column.
    df = pd.json_normalize(customers)
    df.to_excel(out_path, index=False, sheet_name="customers")

    print(f"wrote {len(df)} customers ({len(df.columns)} columns) to {out_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
