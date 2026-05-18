"""Page through N days of LogiNext orders and write them to Excel.

What this script does
---------------------
Calls the LogiNext shipment endpoint
(``ShipmentApp/mile/v1/shipment``) repeatedly to assemble a complete
list of orders over a configurable date range (default: 30 days, all
statuses).

By default, each sub-window is queried once with ``status=ALL`` (the
catch-all used by ``fetch_current_orders.py`` / the production route
handler). Pass ``--statuses`` with a comma-separated list to split
into one query per documented status instead.

Because the API restricts each query to a date range of **less than**
7 days (per ``references/LGNX_API/loginext_mile_apis/10_order.md`` --
"the duration between the Start Date and End Date must be less than 7
days"), this script chunks the requested range into sub-windows and
paginates within each one, then dedupes by ``orderNo`` and dumps every
upstream field to a single-sheet Excel workbook.

API constraints handled
-----------------------
- **Date-range cap**: each sub-window is ``6d 23h 59m`` wide, just
  under the documented 7-day limit.
- **Single-status filter**: the ``status`` parameter accepts one value
  at a time, so each ``status x window`` is its own request (with the
  default ``status=ALL``, that collapses to one request per window).
- **Pagination**: walks ``page_no = 1..N`` per ``status x window`` with
  ``page_size = 100`` (the documented max -- larger values are silently
  capped upstream).
- **Dedup**: orders that straddle a boundary (or get returned by two
  different status filters during a status transition) are deduped by
  ``orderNo``; if missing, ``orderReferenceId`` is used as fallback.

Comparison to ``fetch_current_orders.py``
-----------------------------------------
``fetch_current_orders.py`` returns a single 7-day window for a
frontend feed, reshaping records and filtering to driver-assigned
orders. This script is its audit-shaped sibling: full upstream fields,
all matching orders (no driver filter), Excel output.

Authentication
--------------
``LGNX_AUTH_TOKEN`` must be set in ``.env`` (the value AFTER ``BASIC ``
in the upstream curl's ``www-authenticate`` header). The script
re-prepends ``BASIC `` itself.

Usage
-----
    uv run python fetch_orders_list.py \\
        [--days 30] [--statuses COMPLETED,CANCELLED] \\
        [--page-size 100] [--delay 1.0] \\
        [--api-url URL] [--out output/orders.xlsx]

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
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pandas as pd
import requests
from dotenv import load_dotenv


# Default location for Excel exports.
OUTPUT_DIR = Path(__file__).parent / "output"

# Default upstream endpoint -- overridable via --api-url or
# LOGINEXT_SHIPMENTS_API_URL.
DEFAULT_API_URL = "https://api.loginextsolutions.com/ShipmentApp/mile/v1/shipment"

DEFAULT_DAYS = 30
# Empty default = single pass per window with ``status=ALL`` (matches
# fetch_current_orders.py and the production route handler).
DEFAULT_STATUSES = ""
DEFAULT_PAGE_SIZE = 100  # documented max
DEFAULT_DELAY = 1.0

# Sub-window width: 1 minute under the documented "less than 7 days"
# cap. Matches the offset used in ``fetch_current_orders.py``.
WINDOW = timedelta(days=6, hours=23, minutes=59)

# Browser-style user-agent -- LogiNext's edge occasionally rejects the
# default python-requests UA during WAF checks.
USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36"
)


def _fmt(dt: datetime) -> str:
    """Format a datetime the way LogiNext expects: ``yyyy-MM-dd HH:mm:ss``."""
    return dt.strftime("%Y-%m-%d %H:%M:%S")


def _iter_windows(now: datetime, total_days: int):
    """Yield ``(start, end)`` sub-windows walking backwards from ``now``.

    Each window is ``WINDOW`` wide, except the oldest one which is
    clipped to ``now - total_days``. Adjacent windows are separated by
    1 second to avoid a single boundary timestamp matching in two
    requests (the API treats both bounds as inclusive in practice).
    """
    span_start = now - timedelta(days=total_days)
    end = now
    while end > span_start:
        start = max(span_start, end - WINDOW)
        yield start, end
        # Step the next end past this window with a 1s gap.
        end = start - timedelta(seconds=1)


def _fetch_page(
    token: str,
    api_url: str,
    start: datetime,
    end: datetime,
    status: str,
    page_no: int,
    page_size: int,
) -> tuple[list, dict]:
    """GET one page from the shipment endpoint.

    Returns
    -------
    (records, body)
        ``records`` is the list extracted from ``body["data"]`` or
        ``body["data"]["results"]``; ``body`` is the parsed JSON
        envelope so the caller can read paging hints
        (``moreResultsExists``).
    """
    params = {
        "start_date": _fmt(start),
        "end_date": _fmt(end),
        "status": status,
        "page_no": str(page_no),
        "page_size": str(page_size),
    }
    headers = {
        "content-type": "application/json",
        "user-agent": USER_AGENT,
        "www-authenticate": f"BASIC {token}",
    }
    response = requests.get(api_url, headers=headers, params=params, timeout=30)

    try:
        body = response.json()
    except ValueError:
        print(
            f"window [{_fmt(start)} .. {_fmt(end)}] status={status} "
            f"page={page_no}: HTTP {response.status_code} non-JSON response: "
            f"{response.text[:200]}",
            file=sys.stderr,
        )
        response.raise_for_status()
        return [], {}

    if not response.ok:
        print(
            f"window [{_fmt(start)} .. {_fmt(end)}] status={status} "
            f"page={page_no}: HTTP {response.status_code} "
            f"{json.dumps(body)[:300]}",
            file=sys.stderr,
        )
        response.raise_for_status()

    # Records live at body["data"] (array) or body["data"]["results"]
    # depending on payload shape; mirror fetch_current_orders.py.
    data = body.get("data") if isinstance(body, dict) else None
    if isinstance(data, list):
        records = data
    elif isinstance(data, dict) and isinstance(data.get("results"), list):
        records = data["results"]
    else:
        records = []
    return records, (body if isinstance(body, dict) else {})


def fetch_all(
    token: str,
    api_url: str,
    total_days: int,
    statuses: list[str],
    page_size: int,
    delay: float,
) -> list[dict]:
    """Walk every sub-window x status, paginate, dedupe by orderNo.

    Returns the deduped list of upstream order dicts in arrival order.
    """
    now = datetime.now()
    # Dict keyed by orderNo (fallback orderReferenceId) so we keep the
    # first occurrence of each order across overlapping queries.
    seen: dict[str, dict] = {}

    for window_idx, (start, end) in enumerate(_iter_windows(now, total_days), start=1):
        for status in statuses:
            page = 1
            while True:
                records, body = _fetch_page(
                    token, api_url, start, end, status, page, page_size,
                )
                more = bool(body.get("moreResultsExists"))

                added = 0
                for rec in records:
                    if not isinstance(rec, dict):
                        continue
                    key = rec.get("orderNo") or rec.get("orderReferenceId")
                    if not key:
                        continue
                    key = str(key)
                    if key in seen:
                        continue
                    seen[key] = rec
                    added += 1

                print(
                    f"win {window_idx} [{_fmt(start)} .. {_fmt(end)}] "
                    f"status={status} page={page}: fetched {len(records)} "
                    f"(+{added} new, running total {len(seen)}, more={more})"
                )

                # Stop when the API drained the window for this status:
                # zero records, or a partial page combined with the
                # API's own ``more=False`` hint.
                if not records or (not more and len(records) < page_size):
                    break

                page += 1
                time.sleep(delay)

    return list(seen.values())


def main() -> int:
    """CLI entry point: parse args, fetch all matching orders, write Excel."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--days",
        type=int,
        default=DEFAULT_DAYS,
        help=f"how many days of history to fetch (default: {DEFAULT_DAYS})",
    )
    parser.add_argument(
        "--statuses",
        default=DEFAULT_STATUSES,
        help=(
            "comma-separated statuses to query (default: empty = single pass "
            "per window with status=ALL; documented values: NOTDISPATCHED, "
            "INTRANSIT, COMPLETED, NOTCOMPLETED, PICKEDUP, CANCELLED)"
        ),
    )
    parser.add_argument(
        "--page-size",
        type=int,
        default=DEFAULT_PAGE_SIZE,
        help=f"results per page (max {DEFAULT_PAGE_SIZE}; also the default)",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=DEFAULT_DELAY,
        help=f"seconds to sleep between page requests (default: {DEFAULT_DELAY})",
    )
    parser.add_argument(
        "--api-url",
        default=None,
        help="override the LogiNext shipment endpoint",
    )
    parser.add_argument(
        "--out",
        default=None,
        help="output Excel file (default: output/orders_<statuses>_<UTC-timestamp>.xlsx)",
    )
    args = parser.parse_args()

    load_dotenv()
    token = os.environ.get("LGNX_AUTH_TOKEN")
    if not token:
        print("error: LGNX_AUTH_TOKEN is missing (set it in .env)", file=sys.stderr)
        return 1

    api_url = (
        args.api_url
        or os.environ.get("LOGINEXT_SHIPMENTS_API_URL")
        or DEFAULT_API_URL
    )
    statuses = [s.strip().upper() for s in args.statuses.split(",") if s.strip()]
    # Empty list = no per-status split; do a single pass with ``status=ALL``
    # per window, matching fetch_current_orders.py / the route handler.
    if not statuses:
        statuses = ["ALL"]

    records = fetch_all(
        token, api_url, args.days, statuses, args.page_size, args.delay,
    )

    # Resolve the output path; timestamp default names so successive
    # runs do not overwrite one another.
    if args.out:
        out_path = Path(args.out)
    else:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
        status_slug = "_".join(s.lower() for s in statuses)
        out_path = OUTPUT_DIR / f"orders_{status_slug}_{stamp}.xlsx"
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # Flatten nested JSON (e.g. ``shipperAddress.city``) and write.
    df = pd.json_normalize(records)
    df.to_excel(out_path, index=False, sheet_name="orders")
    print(f"wrote {len(df)} unique orders ({len(df.columns)} columns) to {out_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
