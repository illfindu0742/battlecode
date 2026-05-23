#!/usr/bin/env python3
"""
volunteercabinrentals.com occupancy scraper.

Queries the public date-range search endpoint night-by-night to build a
booked-percentage table for every cabin on the site. Public data only --
exactly what any guest sees while shopping for dates.

Run:
    ./run.sh --days 180 --my-cabins "Cabin A, Cabin B"

Outputs:
    occupancy.csv        -- per-cabin booked nights / total / booked%
    .cache/cabins.json   -- discovered cabin list (delete to refresh)
    .cache/nightly.json  -- per-night availability (resumable)
"""

import argparse
import csv
import json
import re
import sys
import time
from datetime import date, timedelta
from pathlib import Path
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup

BASE = "https://www.volunteercabinrentals.com"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (compatible; cabin-occupancy-research/1.0; "
        "owner self-research)"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.5",
}

LISTING_PAGES = [
    "/find-cabin-by-name.php",
    "/one-bedroom-cabins.php",
    "/two-bedroom-cabins.php",
    "/three-bedroom-cabins.php",
    "/four-bedroom-cabins.php",
    "/five-bedroom-cabins.php",
    "/large-bedroom-cabins.php",
]

# Cabin-detail links on this site follow patterns like:
#   cabin-detail.php?cid=...
#   cabin-info.php?...
#   /cabin_42.php
CABIN_HREF_RE = re.compile(r"cabin[-_]?(detail|info|view|\d)", re.I)

SKIP_PREFIXES = (
    "see ", "book ", "view ", "more", "click", "details",
    "read ", "image", "photo", "check ", "compare",
)


def fetch(session, path_or_url, params=None, retries=3):
    url = (
        path_or_url
        if path_or_url.startswith("http")
        else urljoin(BASE, path_or_url)
    )
    for attempt in range(retries):
        try:
            r = session.get(url, params=params, timeout=30)
            r.raise_for_status()
            return r.text
        except requests.RequestException as e:
            if attempt == retries - 1:
                raise
            wait = 2 ** attempt
            print(f"  retry in {wait}s ({e})", file=sys.stderr)
            time.sleep(wait)


def extract_cabin_links(html):
    """Return {name: url} pairs found on a listing or search-results page."""
    soup = BeautifulSoup(html, "html.parser")
    out = {}
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if not CABIN_HREF_RE.search(href):
            continue
        name = a.get_text(" ", strip=True)
        if not name:
            # Sometimes the visible link text is an image; try title/alt.
            img = a.find("img")
            if img:
                name = (img.get("alt") or img.get("title") or "").strip()
        if not name or len(name) > 80:
            continue
        if name.lower().startswith(SKIP_PREFIXES):
            continue
        # Prefer the first-seen URL for a given name.
        out.setdefault(name, urljoin(BASE, href))
    return out


def discover_cabins(session, delay):
    cabins = {}
    for p in LISTING_PAGES:
        try:
            html = fetch(session, p)
        except Exception as e:
            print(f"discover: {p} failed: {e}", file=sys.stderr)
            continue
        found = extract_cabin_links(html)
        for name, url in found.items():
            cabins.setdefault(name, url)
        print(
            f"  {p}: +{len(found)} cabins (total {len(cabins)})",
            file=sys.stderr,
        )
        time.sleep(delay)
    return cabins


def available_on(session, night):
    nxt = night + timedelta(days=1)
    html = fetch(
        session,
        "/cabin-search.php",
        params={"checkin": night.isoformat(), "checkout": nxt.isoformat()},
    )
    return set(extract_cabin_links(html).keys())


def main():
    ap = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument(
        "--days", type=int, default=180,
        help="Window of nights to sample (default 180)",
    )
    ap.add_argument("--start", help="ISO start date (default: today)")
    ap.add_argument(
        "--my-cabins", default="",
        help="Comma-separated names of your cabins to highlight",
    )
    ap.add_argument("--out", default="occupancy.csv")
    ap.add_argument("--cache-dir", default=".cache")
    ap.add_argument(
        "--delay", type=float, default=1.0,
        help="Seconds between requests (default 1.0 -- be polite)",
    )
    ap.add_argument(
        "--probe", action="store_true",
        help="Dump one search-page HTML to debug.html and exit so HTML "
             "selectors can be verified before a full sweep",
    )
    args = ap.parse_args()

    session = requests.Session()
    session.headers.update(HEADERS)

    if args.probe:
        n = date.today() + timedelta(days=30)
        html = fetch(
            session,
            "/cabin-search.php",
            params={
                "checkin": n.isoformat(),
                "checkout": (n + timedelta(days=1)).isoformat(),
            },
        )
        Path("debug.html").write_text(html)
        print(f"Wrote debug.html ({len(html)} bytes)", file=sys.stderr)
        found = extract_cabin_links(html)
        print(f"Detected {len(found)} cabin links:", file=sys.stderr)
        for name in sorted(found)[:25]:
            print(f"  - {name}", file=sys.stderr)
        return

    cache = Path(args.cache_dir)
    cache.mkdir(exist_ok=True)
    cabins_path = cache / "cabins.json"
    nightly_path = cache / "nightly.json"

    if cabins_path.exists():
        cabins = json.loads(cabins_path.read_text())
        print(f"Loaded {len(cabins)} cabins from cache", file=sys.stderr)
    else:
        print("Discovering cabins...", file=sys.stderr)
        cabins = discover_cabins(session, args.delay)
        cabins_path.write_text(json.dumps(cabins, indent=2))

    start = date.fromisoformat(args.start) if args.start else date.today()
    nights = [start + timedelta(days=i) for i in range(args.days)]

    nightly = {}
    if nightly_path.exists():
        nightly = json.loads(nightly_path.read_text())
        print(f"Resuming with {len(nightly)} cached nights", file=sys.stderr)

    todo = [n for n in nights if n.isoformat() not in nightly]
    print(
        f"Sweeping {len(todo)} new nights "
        f"({nights[0]} ... {nights[-1]})",
        file=sys.stderr,
    )

    for i, n in enumerate(todo, 1):
        try:
            avail = available_on(session, n)
        except Exception as e:
            print(f"[{i}/{len(todo)}] {n} FAILED: {e}", file=sys.stderr)
            continue
        nightly[n.isoformat()] = sorted(avail)
        if i % 10 == 0 or i == len(todo):
            nightly_path.write_text(json.dumps(nightly, indent=2))
            print(
                f"[{i}/{len(todo)}] {n}: {len(avail)} avail (cache saved)",
                file=sys.stderr,
            )
        time.sleep(args.delay)
    nightly_path.write_text(json.dumps(nightly, indent=2))

    # Build cabin universe: union of discovery + every night's results.
    # (Cabins booked 100% of the window won't show up in nightly data, so
    # the discovery pass is what catches them.)
    universe = set(cabins.keys())
    for names in nightly.values():
        universe.update(names)
    universe = sorted(universe)

    total = len(nightly)
    rows = []
    for name in universe:
        booked = sum(1 for ns in nightly.values() if name not in ns)
        pct = booked / total if total else 0
        rows.append((name, booked, total, pct))
    rows.sort(key=lambda r: -r[3])

    mine = {s.strip().lower() for s in args.my_cabins.split(",") if s.strip()}
    with open(args.out, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(
            ["cabin", "booked_nights", "total_nights", "booked_pct", "yours"]
        )
        for name, b, t, p in rows:
            w.writerow([name, b, t, f"{p:.1%}", name.lower() in mine])
    print(
        f"\nWrote {args.out} ({len(rows)} cabins, {total} nights sampled)",
        file=sys.stderr,
    )

    print(f"\n{'cabin':<40} booked%  nights")
    print("-" * 60)
    for name, b, t, p in rows[:25]:
        flag = " *" if name.lower() in mine else "  "
        print(f"{name[:38]:<40} {p:>6.1%}  {b}/{t}{flag}")
    if mine:
        print("\nYour cabins:")
        for idx, (name, b, t, p) in enumerate(rows, 1):
            if name.lower() in mine:
                print(f"  #{idx:>3}  {name:<40} {p:>6.1%}  {b}/{t}")


if __name__ == "__main__":
    main()
