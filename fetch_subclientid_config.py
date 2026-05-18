"""Fetch full shipper configuration by ``subClientId`` from LogiNext.

What this script does
---------------------
Calls the LogiNext ``ClientApp/shipper/getbysubclientid`` endpoint to
download the complete configuration record (settings, geofences,
preferences, etc.) for one or more shippers, identified by their
``subClientId``.

Two modes
---------
1. **Ad-hoc mode** -- pass one or more ``subClientId`` values as
   positional arguments. The combined response is saved as a JSON file
   under ``output/`` (or wherever ``--out`` points). When exactly one
   id is provided the file contains that id's body directly; with more
   than one the file is keyed by id.

2. **Bulk mode** (``--from-xlsx``) -- read a shippers Excel produced by
   ``fetch_shipper_list.py`` and iterate every ``subClientId`` value in
   it. The combined result is flattened into a DataFrame and written
   as Excel. Two synthetic columns are added per row:
       * ``_fetchStatus`` -- HTTP status code observed for that id.
       * ``_fetchError``  -- error body snippet when the call failed.

Retries and rate limits
-----------------------
LogiNext's Tier-1 APIs are documented at 5 req/s. We sleep ``--delay``
seconds (default 0.5s -> ~2 req/s effective) between calls to leave
headroom for network jitter. Transient failures (HTTP 429 and 5xx) are
retried with exponential backoff up to 3 times before being reported.

Progress UI
-----------
Bulk mode uses ``tqdm`` to render a live progress bar with success /
error counters and the last id processed -- useful when running over
hundreds of shippers.

Authentication
--------------
``LGNX_AUTH_TOKEN`` must be set in ``.env`` (the value AFTER ``BASIC ``
in the upstream curl's ``www-authenticate`` header).

API reference
-------------
Mirrors ``references/curls/fetch_subclientId_config.sh``.

Usage
-----
    # Fetch specific subClientIds (writes JSON):
    uv run python fetch_subclientid_config.py <subclientId> [<subclientId> ...] \\
        [--out path.json]

    # Fetch every shipper from a shippers list xlsx (writes Excel):
    uv run python fetch_subclientid_config.py \\
        --from-xlsx output/shippers_list_*.xlsx \\
        [--delay 0.5] [--out path.xlsx]

Exit codes
----------
    0 -- success.
    1 -- configuration error (missing token, bad CLI flag combo).
    2 -- one or more upstream calls returned a non-2xx status.
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
from tqdm import tqdm

# Default output directory for both JSON and Excel exports.
OUTPUT_DIR = Path(__file__).parent / "output"

# Documented endpoint for "get shipper configuration by subClientId".
URL = "https://products.loginextsolutions.com/ClientApp/shipper/getbysubclientid"

# Browser-style user-agent -- LogiNext's edge occasionally rejects the
# default python-requests UA during WAF checks.
USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36"
)

# Tier 1 APIs are limited to 5 req/s; 0.5s between requests gives us
# 2 req/s effective and accommodates network jitter.
DEFAULT_DELAY = 0.5


def fetch_config(
    token: str,
    subclient_id: str,
    max_retries: int = 3,
    backoff: float = 2.0,
) -> tuple[int, dict | str]:
    """Fetch one shipper's configuration with retry-on-transient logic.

    Steps performed
    ---------------
    1. Build the request headers.
    2. Issue a ``GET`` with ``subclientId`` as a query parameter.
    3. If the status code is 429 or a 5xx *and* the retry budget is not
       exhausted, sleep with exponential backoff (``backoff * 2**attempt``)
       and try again. This pattern handles brief rate-limit blips and
       transient server hiccups gracefully.
    4. Otherwise return ``(status_code, parsed_body)``. ``parsed_body``
       is the decoded JSON when possible, or the raw text when the API
       returned a non-JSON body (e.g. an HTML error page).

    Parameters
    ----------
    token : str
        Raw ``LGNX_AUTH_TOKEN`` value. ``BASIC `` is prepended here.
    subclient_id : str
        The ``subClientId`` to query.
    max_retries : int, optional
        Maximum retry attempts on transient failures. Default 3.
    backoff : float, optional
        Base seconds for exponential backoff. The first retry waits
        ``backoff * 1``, the second ``backoff * 2``, etc.

    Returns
    -------
    tuple[int, dict | str]
        ``(status_code, body)``. ``body`` is a dict on JSON responses,
        otherwise the raw text body.
    """
    # Step 1: headers used for every request to this endpoint.
    headers = {
        "accept": "application/json, text/plain, */*",
        "user-agent": USER_AGENT,
        "www-authenticate": f"BASIC {token}",
    }
    attempt = 0
    while True:
        # Step 2: send the request.
        response = requests.get(
            URL,
            headers=headers,
            params={"subclientId": subclient_id},
            timeout=30,
        )

        # Step 3: retry on transient errors (rate limit + 5xx) with
        # exponential backoff -- but only while we still have retries.
        if response.status_code in (429, 500, 502, 503, 504) and attempt < max_retries:
            sleep_for = backoff * (2**attempt)
            print(
                f"subClientId={subclient_id}: HTTP {response.status_code}, "
                f"retrying in {sleep_for:.1f}s (attempt {attempt + 1}/{max_retries})",
                file=sys.stderr,
            )
            time.sleep(sleep_for)
            attempt += 1
            continue

        # Step 4: return the parsed response (or raw text on JSON failure).
        try:
            return response.status_code, response.json()
        except ValueError:
            return response.status_code, response.text


def fetch_all_from_xlsx(
    token: str,
    xlsx_path: Path,
    delay: float,
    id_column: str = "subClientId",
) -> pd.DataFrame:
    """Read an Excel file of shippers and fetch each shipper's config.

    Steps performed
    ---------------
    1. Read the source Excel into a DataFrame.
    2. Verify that ``id_column`` exists; raise a helpful error otherwise.
    3. Extract the ids, drop NaNs, and cast to ``int64`` (subClientId is
       always a numeric id even when Excel exports it as float).
    4. Set up a ``tqdm`` progress bar over the ids.
    5. For each id: call ``fetch_config`` and classify the result:
        * Success -- annotate the data dict with ``_fetchStatus`` and
          append to the row list.
        * Failure -- append a placeholder row containing the id, status,
          and a snippet of the error body so failures stay traceable.
    6. Update the progress bar's postfix with running ok/err totals and
       the last-processed id.
    7. Sleep ``delay`` seconds between requests (skipping the sleep on
       the final iteration so the script doesn't wait pointlessly).
    8. Flatten the accumulated rows into a DataFrame via
       ``pandas.json_normalize`` (so nested config fields become dotted
       column names).

    Parameters
    ----------
    token : str
        Raw ``LGNX_AUTH_TOKEN`` value.
    xlsx_path : Path
        Path to a shippers list xlsx.
    delay : float
        Seconds to sleep between successive API calls.
    id_column : str, optional
        Column in the xlsx that holds the subClientId. Default
        ``"subClientId"``.

    Returns
    -------
    pandas.DataFrame
        One row per shipper, flattened to a tabular shape. Includes
        ``_fetchStatus`` (always) and ``_fetchError`` (only for failed
        rows) columns.

    Raises
    ------
    ValueError
        If ``id_column`` is missing from the source workbook.
    """
    # Step 1: load the source workbook.
    src = pd.read_excel(xlsx_path)

    # Step 2: verify the id column is present.
    if id_column not in src.columns:
        raise ValueError(
            f"column {id_column!r} not found in {xlsx_path} "
            f"(columns: {list(src.columns)})"
        )

    # Step 3: extract ids -> drop missing -> coerce to int (Excel
    # sometimes turns integer ids into floats during import).
    ids = src[id_column].dropna().astype("int64").tolist()
    total = len(ids)
    print(f"fetching configuration for {total} shippers from {xlsx_path.name}")

    rows: list[dict] = []
    ok_count = 0
    err_count = 0

    # Step 4: progress bar over the ids. ``dynamic_ncols`` makes the bar
    # adapt to terminal resizing.
    bar = tqdm(ids, desc="shippers", unit="req", dynamic_ncols=True)
    for i, sid in enumerate(bar, 1):
        # Step 5: fetch and classify.
        status, body = fetch_config(token, str(sid))

        if status == 200 and isinstance(body, dict) and not body.get("hasError"):
            # Success: store the inner ``data`` dict + status marker.
            data = body.get("data") or {}
            data["_fetchStatus"] = status
            rows.append(data)
            ok_count += 1
            # Step 6: progress update.
            bar.set_postfix(ok=ok_count, err=err_count, last=str(sid))
        else:
            # Failure: store a placeholder row so the id is still
            # represented in the output and the failure is auditable.
            err = body if isinstance(body, str) else json.dumps(body)[:300]
            rows.append(
                {
                    "subClientId": sid,
                    "_fetchStatus": status,
                    "_fetchError": err,
                }
            )
            err_count += 1
            # Step 6: progress update + per-error log line.
            bar.set_postfix(ok=ok_count, err=err_count, last=str(sid))
            bar.write(f"subClientId={sid}: HTTP {status} error")

        # Step 7: politeness sleep -- skipped after the last id.
        if i < total:
            time.sleep(delay)

    bar.close()

    # Step 8: flatten nested config dicts (e.g. ``geofence.center.lat``)
    # into a wide table.
    return pd.json_normalize(rows)


def main() -> int:
    """CLI entry point dispatching between ad-hoc and bulk modes.

    Steps performed
    ---------------
    1. Parse CLI args.
    2. Validate the flag combination (exactly one of positional ids or
       ``--from-xlsx`` must be supplied).
    3. Load ``.env`` and verify ``LGNX_AUTH_TOKEN``.
    4. Prepare the output directory and a shared UTC timestamp.
    5a. Bulk mode (``--from-xlsx``): drive ``fetch_all_from_xlsx`` and
        write the resulting DataFrame to Excel. Return 2 if any rows
        had errors, else 0.
    5b. Ad-hoc mode (positional ids): call ``fetch_config`` once per id,
        print the response, accumulate into a dict, and write JSON. The
        output is the raw response when a single id was provided, or a
        ``{id: response}`` map for multiple ids.

    Returns
    -------
    int
        Process exit code.
    """
    # Step 1: argument parsing.
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "subclient_ids",
        nargs="*",
        help="one or more subClientId values to fetch configuration for "
        "(omit when using --from-xlsx)",
    )
    parser.add_argument(
        "--from-xlsx",
        default=None,
        help="path to a shippers list xlsx; iterates every subClientId column value",
    )
    parser.add_argument(
        "--id-column",
        default="subClientId",
        help="column name in --from-xlsx that holds subClientId (default: subClientId)",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=DEFAULT_DELAY,
        help=f"seconds to sleep between requests (default: {DEFAULT_DELAY})",
    )
    parser.add_argument(
        "--out",
        default=None,
        help="output path. JSON for ID args, Excel for --from-xlsx "
        "(default: output/<prefix>_<UTC-timestamp>.<ext>)",
    )
    args = parser.parse_args()

    # Step 2: enforce mode exclusivity. The two modes need different
    # output handling, so it's clearer to require an explicit choice.
    if not args.subclient_ids and not args.from_xlsx:
        parser.error("provide subclient ids or --from-xlsx")
    if args.subclient_ids and args.from_xlsx:
        parser.error("provide either subclient ids or --from-xlsx, not both")

    # Step 3: load env + verify token.
    load_dotenv()
    token = os.environ.get("LGNX_AUTH_TOKEN")
    if not token:
        print("error: LGNX_AUTH_TOKEN is missing (set it in .env)", file=sys.stderr)
        return 1

    # Step 4: prepare output dir + a shared timestamp for output names.
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")

    # Step 5a: bulk mode -- iterate every shipper in the source xlsx.
    if args.from_xlsx:
        df = fetch_all_from_xlsx(token, Path(args.from_xlsx), args.delay, args.id_column)
        out_path = Path(args.out) if args.out else OUTPUT_DIR / f"shippers_{stamp}.xlsx"
        out_path.parent.mkdir(parents=True, exist_ok=True)
        df.to_excel(out_path, index=False, sheet_name="shippers")
        print(f"wrote {len(df)} shipper configs ({len(df.columns)} columns) to {out_path}")
        # Exit 2 if any rows had a fetch error (column exists and has
        # at least one non-NaN value); otherwise 0.
        return 0 if "_fetchError" not in df.columns or df["_fetchError"].isna().all() else 2

    # Step 5b: ad-hoc mode -- one or more positional ids.
    results: dict[str, dict | str] = {}
    exit_code = 0
    for i, subclient_id in enumerate(args.subclient_ids):
        # Fetch one id at a time so each response is printed immediately.
        status, body = fetch_config(token, subclient_id)
        print(f"subClientId={subclient_id}: HTTP {status}")
        print(json.dumps(body, indent=2) if isinstance(body, dict) else body)
        results[subclient_id] = body
        # Treat any 4xx/5xx as a "partial success" that should bubble up
        # as a non-zero exit code, while still saving what we did get.
        if status >= 400:
            exit_code = 2
        # Politeness sleep between ids (skip after the final one).
        if i < len(args.subclient_ids) - 1:
            time.sleep(args.delay)

    # Resolve the JSON output path and write the result. For a single id
    # we unwrap the map so the file *is* the response body (handier when
    # piping into ``jq`` downstream); for multiple ids we keep the map
    # so each id's body is addressable by key.
    out_path = Path(args.out) if args.out else OUTPUT_DIR / f"subclient_config_{stamp}.json"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    payload = results[args.subclient_ids[0]] if len(args.subclient_ids) == 1 else results
    with out_path.open("w", encoding="utf-8") as fh:
        json.dump(payload, fh, indent=2, ensure_ascii=False)

    print(f"wrote configuration for {len(results)} subClientId(s) to {out_path}")
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
