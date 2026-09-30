#!/usr/bin/env python3
"""
download_sources.py — fetch the source files that have a direct download URL.

Driven entirely by config/sources.csv. Two columns matter:

  download_kind = direct   there is a stable file URL; this script fetches it
  download_kind = manual   the publisher only offers a landing page or an
                           interactive query builder, so a human has to do it.
                           The script prints a checklist with the URL and the
                           settings to choose.

Files land in raw_data_downloaded/ so they never overwrite the raw_data/ set the
panel was built from. Compare before swapping them in — sources revise, and a
silent change in an input is how a rebuild starts disagreeing with the published
numbers.

Usage
    python download_sources.py                    # fetch everything marked direct
    python download_sources.py --list             # show what is direct vs manual
    python download_sources.py --force            # re-fetch files already present
    python download_sources.py --dest DIR --config DIR

Only the standard library is used, so this runs before pip install.
"""
import argparse, csv, os, sys, time, urllib.request, urllib.error, urllib.parse

UA = "Mozilla/5.0 (compatible; ccdr-indicator-build/1.0)"
TIMEOUT = 120
RETRIES = 3


def filename_for(url: str) -> str:
    """Name the local file. World Bank API calls carry the indicator code in the
    path, not the filename, so reconstruct the name the portal would have given."""
    p = urllib.parse.urlparse(url)
    if "api.worldbank.org" in p.netloc and "/indicator/" in p.path:
        code = p.path.split("/indicator/")[1].split("/")[0]
        return "API_" + code.replace(".", "_") + "_DS2_en_csv_v2.zip"
    if "dataverse.harvard.edu" in p.netloc:
        return "dataverse_files.zip"
    name = os.path.basename(p.path)
    return name or (p.netloc.replace(".", "_") + ".download")


def fetch(url: str, dest: str) -> tuple[bool, str]:
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    for attempt in range(1, RETRIES + 1):
        try:
            with urllib.request.urlopen(req, timeout=TIMEOUT) as r, open(dest, "wb") as fh:
                total = 0
                while True:
                    chunk = r.read(1 << 16)
                    if not chunk:
                        break
                    fh.write(chunk)
                    total += len(chunk)
            if total < 1024:            # a 200 that returns an error page
                os.remove(dest)
                return False, f"suspiciously small ({total} bytes) - likely an error page"
            return True, f"{total/1e6:.1f} MB"
        except urllib.error.HTTPError as e:
            if attempt == RETRIES:
                return False, f"HTTP {e.code}"
        except Exception as e:
            if attempt == RETRIES:
                return False, f"{type(e).__name__}: {e}"
        time.sleep(2 * attempt)
    return False, "failed"


def main() -> int:
    here = os.path.dirname(os.path.abspath(__file__))
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", default=os.environ.get("CONFIG_DIR", os.path.join(here, "..", "config")))
    ap.add_argument("--dest",   default=os.environ.get("DL_DIR",     os.path.join(here, "..", "raw_data_downloaded")))
    ap.add_argument("--force", action="store_true", help="re-fetch files already present")
    ap.add_argument("--list",  action="store_true", help="show the plan and exit")
    a = ap.parse_args()

    cfg = os.path.join(os.path.abspath(a.config), "sources.csv")
    if not os.path.exists(cfg):
        print(f"ERROR: {cfg} not found", file=sys.stderr)
        return 1
    dest = os.path.abspath(a.dest)
    os.makedirs(dest, exist_ok=True)

    with open(cfg, newline="", encoding="utf-8") as fh:
        rows = list(csv.DictReader(fh))

    direct = [r for r in rows if (r.get("download_kind") or "").strip() == "direct"]
    manual = [r for r in rows if (r.get("download_kind") or "").strip() != "direct"]

    if a.list:
        print(f"\nDIRECT — fetched by this script ({len(direct)} entries):")
        for r in direct:
            for u in r["download_url"].split(";"):
                print(f"   {filename_for(u):44s} {u}")
        print(f"\nMANUAL — download these yourself ({len(manual)} entries):")
        for r in manual:
            print(f"   {r['file']}\n      {r['url']}\n      {r.get('download_note','')}")
        return 0

    print(f"\nDestination: {dest}\n")
    ok = skipped = failed = 0
    for r in direct:
        for url in [u for u in r["download_url"].split(";") if u.strip()]:
            name = filename_for(url)
            path = os.path.join(dest, name)
            if os.path.exists(path) and not a.force:
                print(f"  skip   {name:44s} already present (--force to refetch)")
                skipped += 1
                continue
            print(f"  get    {name:44s} ...", end=" ", flush=True)
            good, msg = fetch(url, path)
            print(msg if good else f"FAILED - {msg}")
            ok += good
            failed += (not good)

    print(f"\n  downloaded {ok}, skipped {skipped}, failed {failed}")

    print(f"\n{'='*74}\nSTILL TO DOWNLOAD BY HAND — {len(manual)} sources\n{'='*74}")
    for r in manual:
        print(f"\n  {r['file']}")
        print(f"     {r['url']}")
        if r.get("download_note"):
            print(f"     -> {r['download_note']}")
    print(f"\n{'='*74}")
    print("Files were written to raw_data_downloaded/, NOT raw_data/.")
    print("Diff them against raw_data/ before swapping any in: publishers revise")
    print("series, and an unnoticed change is how a rebuild quietly stops matching.")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
