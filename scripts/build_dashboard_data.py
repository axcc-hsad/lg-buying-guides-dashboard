#!/usr/bin/env python3

import argparse
import json
from pathlib import Path

from generate_data import build_payload_from_workbooks

ROOT = Path(__file__).resolve().parents[1]

CANONICAL = {
    "CA_EN": "CA-EN",
    "CA_FR": "CA-FR",
    "CA-EN": "CA-EN",
    "CA-FR": "CA-FR",
}


def canonical_country(raw):
    if raw is None:
        return raw
    return CANONICAL.get(str(raw).strip().replace("_", "-").upper(), str(raw).strip().replace("_", "-").upper())


def merge_pages(generated_payload, live_snapshot):
    pages = []
    seen = set()
    for page in live_snapshot.get("pages", []):
        link = page.get("link")
        if not link or not str(link).startswith("http") or link in seen:
            continue
        seen.add(link)
        pages.append(page)
    generated_payload["pages"] = pages
    return generated_payload


def main():
    parser = argparse.ArgumentParser(description="Build the final dashboard data.json from lineup + feature library workbooks.")
    parser.add_argument("--lineup-workbook", default=str(ROOT / "data" / "monthly_ga_lineupguide.xlsx"))
    parser.add_argument("--feature-workbook", default=str(ROOT / "data" / "monthly_ga_featurelibrary.xlsx"))
    parser.add_argument("--live-snapshot", default=str(ROOT / "data" / "live-data.raw.json"))
    parser.add_argument("--output", default=str(ROOT / "data" / "data.json"))
    args = parser.parse_args()

    workbook_specs = []
    lineup_path = Path(args.lineup_workbook)
    if lineup_path.exists():
        workbook_specs.append((lineup_path, "lineup"))
    feature_path = Path(args.feature_workbook)
    if feature_path.exists():
        workbook_specs.append((feature_path, "feature_library"))
    if not workbook_specs:
        raise SystemExit("No workbook files found.")

    payload = build_payload_from_workbooks(workbook_specs)

    live_snapshot_path = Path(args.live_snapshot)
    if live_snapshot_path.exists():
        with live_snapshot_path.open("r", encoding="utf-8") as stream:
            live_snapshot = json.load(stream)
        payload = merge_pages(payload, live_snapshot)

    payload["meta"]["countries"] = sorted({canonical_country(code) for code in payload["meta"]["countries"]})

    output_path = Path(args.output)
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote {output_path}")


if __name__ == "__main__":
    main()
