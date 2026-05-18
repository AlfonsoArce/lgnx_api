"""GET full shipper configuration by subClientId from LogiNext.

Mirrors references/curls/fetch_subclientId_config.sh. Set LGNX_AUTH_TOKEN in
.env (the value after `BASIC ` in the curl's WWW-Authenticate header).

Usage:
    # Fetch specific subClientIds (saves JSON):
    uv run python fetch_subclientid_config.py <subclientId> [<subclientId> ...] [--out path.json]

    # Fetch every shipper from a shippers list xlsx (saves Excel):
    uv run python fetch_subclientid_config.py --from-xlsx exports/shippers_list_*.xlsx [--delay 0.5] [--out path.xlsx]
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

EXPORTS_DIR = Path(__file__).parent / "exports"

URL = "https://products.loginextsolutions.com/ClientApp/shipper/getbysubclientid"

USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36"
)

# Tier 1 APIs are limited to 5 req/s; 0.5s between requests gives us 2 req/s
# headroom and accommodates network jitter.
DEFAULT_DELAY = 0.5


def fetch_config(
    token: str,
    subclient_id: str,
    max_retries: int = 3,
    backoff: float = 2.0,
) -> tuple[int, dict | str]:
    headers = {
        "accept": "application/json, text/plain, */*",
        "user-agent": USER_AGENT,
        "www-authenticate": f"BASIC {token}",
    }
    attempt = 0
    while True:
        response = requests.get(
            URL,
            headers=headers,
            params={"subclientId": subclient_id},
            timeout=30,
        )
        # Retry on 429 (rate limit) and 5xx
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
    src = pd.read_excel(xlsx_path)
    if id_column not in src.columns:
        raise ValueError(
            f"column {id_column!r} not found in {xlsx_path} "
            f"(columns: {list(src.columns)})"
        )
    ids = src[id_column].dropna().astype("int64").tolist()
    total = len(ids)
    print(f"fetching configuration for {total} shippers from {xlsx_path.name}")

    rows: list[dict] = []
    ok_count = 0
    err_count = 0
    bar = tqdm(ids, desc="shippers", unit="req", dynamic_ncols=True)
    for i, sid in enumerate(bar, 1):
        status, body = fetch_config(token, str(sid))
        if status == 200 and isinstance(body, dict) and not body.get("hasError"):
            data = body.get("data") or {}
            data["_fetchStatus"] = status
            rows.append(data)
            ok_count += 1
            bar.set_postfix(ok=ok_count, err=err_count, last=str(sid))
        else:
            err = body if isinstance(body, str) else json.dumps(body)[:300]
            rows.append(
                {
                    "subClientId": sid,
                    "_fetchStatus": status,
                    "_fetchError": err,
                }
            )
            err_count += 1
            bar.set_postfix(ok=ok_count, err=err_count, last=str(sid))
            bar.write(f"subClientId={sid}: HTTP {status} error")

        if i < total:
            time.sleep(delay)

    bar.close()
    return pd.json_normalize(rows)


def main() -> int:
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
        "(default: exports/<prefix>_<UTC-timestamp>.<ext>)",
    )
    args = parser.parse_args()

    if not args.subclient_ids and not args.from_xlsx:
        parser.error("provide subclient ids or --from-xlsx")
    if args.subclient_ids and args.from_xlsx:
        parser.error("provide either subclient ids or --from-xlsx, not both")

    load_dotenv()
    token = os.environ.get("LGNX_AUTH_TOKEN")
    if not token:
        print("error: LGNX_AUTH_TOKEN is missing (set it in .env)", file=sys.stderr)
        return 1

    EXPORTS_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")

    if args.from_xlsx:
        df = fetch_all_from_xlsx(token, Path(args.from_xlsx), args.delay, args.id_column)
        out_path = Path(args.out) if args.out else EXPORTS_DIR / f"shippers_{stamp}.xlsx"
        out_path.parent.mkdir(parents=True, exist_ok=True)
        df.to_excel(out_path, index=False, sheet_name="shippers")
        print(f"wrote {len(df)} shipper configs ({len(df.columns)} columns) to {out_path}")
        return 0 if "_fetchError" not in df.columns or df["_fetchError"].isna().all() else 2

    results: dict[str, dict | str] = {}
    exit_code = 0
    for i, subclient_id in enumerate(args.subclient_ids):
        status, body = fetch_config(token, subclient_id)
        print(f"subClientId={subclient_id}: HTTP {status}")
        print(json.dumps(body, indent=2) if isinstance(body, dict) else body)
        results[subclient_id] = body
        if status >= 400:
            exit_code = 2
        if i < len(args.subclient_ids) - 1:
            time.sleep(args.delay)

    out_path = Path(args.out) if args.out else EXPORTS_DIR / f"subclient_config_{stamp}.json"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    payload = results[args.subclient_ids[0]] if len(args.subclient_ids) == 1 else results
    with out_path.open("w", encoding="utf-8") as fh:
        json.dump(payload, fh, indent=2, ensure_ascii=False)

    print(f"wrote configuration for {len(results)} subClientId(s) to {out_path}")
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
