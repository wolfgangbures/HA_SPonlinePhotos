#!/usr/bin/env python3
"""Fetch historical HA logbook entries plus core/supervisor log tails."""

import argparse
import json
import os
import sys
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timedelta, timezone
from pathlib import Path

HOST = "HA1180-02.fritz.box"
BASE_URL = f"http://{HOST}:8123"
SUPPORT_LOG_ENDPOINTS = {
    "core": "/api/hassio/core/logs",
    "supervisor": "/api/hassio/supervisor/logs",
}


def build_headers():
    token = os.getenv("HA_TOKEN")
    if not token:
        print("ERROR: HA_TOKEN environment variable not set", file=sys.stderr)
        sys.exit(1)

    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }


def fetch_json(url, headers, timeout=60):
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))


def fetch_text(url, headers, timeout=60):
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=timeout) as response:
        return response.read().decode("utf-8", errors="replace")


def fetch_ha_logbook(minutes_back=10, entity_filter=None, headers=None):
    """Fetch historical logbook entries from HA API."""
    headers = headers or build_headers()
    since = (datetime.now(timezone.utc) - timedelta(minutes=minutes_back)).strftime("%Y-%m-%dT%H:%M:%S")

    url = f"{BASE_URL}/api/logbook/{since}"
    if entity_filter:
        query = urllib.parse.urlencode({"entity_id": entity_filter})
        url = f"{url}?{query}"

    print(f"[FETCHING] HA logbook from last {minutes_back} minute(s) (since {since})")

    try:
        data = fetch_json(url, headers=headers, timeout=60)

        if not data:
            print("[NO_ENTRIES] No logbook entries found for this period")
            return []

        print(f"[SUCCESS] Retrieved {len(data)} logbook entries\n")

        for entry in data:
            timestamp = entry.get("when", "unknown")
            entity_id = entry.get("entity_id", "unknown")
            name = entry.get("name", entity_id)
            state = entry.get("state", "N/A")
            message = entry.get("message", "")

            print(f"{timestamp} | {name:<40} | {state:<20}")
            if message:
                print(f"  |- {message}")

        print(f"\n[TOTAL_ENTRIES] {len(data)} logbook entries")
        return data

    except urllib.error.HTTPError as exc:
        print(f"[HTTP_ERROR] {exc.code}: {exc.reason}", file=sys.stderr)
        sys.exit(1)
    except Exception as exc:
        print(f"[ERROR] {type(exc).__name__}: {exc}", file=sys.stderr)
        sys.exit(1)


def fetch_support_logs(headers, output_dir, tail_lines=200):
    """Fetch supporting HA logs that do not appear in the logbook API."""
    saved_files = {}

    for label, endpoint in SUPPORT_LOG_ENDPOINTS.items():
        url = f"{BASE_URL}{endpoint}"
        print(f"\n[FETCHING] {label} logs from {endpoint}")
        try:
            body = fetch_text(url, headers=headers, timeout=60)
        except urllib.error.HTTPError as exc:
            print(f"[HTTP_ERROR] {label}: {exc.code} {exc.reason}", file=sys.stderr)
            continue
        except Exception as exc:
            print(f"[ERROR] {label}: {type(exc).__name__}: {exc}", file=sys.stderr)
            continue

        output_path = output_dir / f"ha_{label}_logs_latest.txt"
        output_path.write_text(body, encoding="utf-8")
        saved_files[label] = output_path

        lines = body.splitlines()
        tail = lines[-tail_lines:] if tail_lines > 0 else lines
        print(f"[SAVED] {label} logs -> {output_path} ({len(lines)} lines)")
        if tail:
            print(f"[TAIL] Last {len(tail)} {label} log lines")
            for line in tail:
                print(line)
        else:
            print(f"[EMPTY] {label} logs returned no lines")

    return saved_files


def cleanup_files(paths: dict[str, Path]) -> None:
    """Delete generated support log files."""
    for path in paths.values():
        try:
            path.unlink(missing_ok=True)
            print(f"[CLEANUP] Deleted {path}")
        except OSError as exc:
            print(f"[CLEANUP_ERROR] {path}: {exc}", file=sys.stderr)


def parse_args(argv):
    parser = argparse.ArgumentParser(
        description="Fetch Home Assistant logbook data plus core/supervisor logs."
    )
    parser.add_argument("minutes", nargs="?", type=int, default=10)
    parser.add_argument("entity", nargs="?", default=None)
    parser.add_argument(
        "--skip-support-logs",
        action="store_true",
        help="Only fetch the logbook and skip core/supervisor logs.",
    )
    parser.add_argument(
        "--support-tail-lines",
        type=int,
        default=200,
        help="How many trailing lines of each support log to print.",
    )
    parser.add_argument(
        "--output-dir",
        default=".",
        help="Directory where fetched support log files are written.",
    )
    parser.add_argument(
        "--keep-support-logs",
        action="store_true",
        help="Keep generated support log files instead of deleting them after printing.",
    )
    return parser.parse_args(argv)


if __name__ == "__main__":
    args = parse_args(sys.argv[1:])
    headers = build_headers()
    fetch_ha_logbook(args.minutes, args.entity, headers=headers)

    if not args.skip_support_logs:
        saved_files = fetch_support_logs(
            headers,
            Path(args.output_dir),
            tail_lines=args.support_tail_lines,
        )
        if saved_files and not args.keep_support_logs:
            cleanup_files(saved_files)