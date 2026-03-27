#!/usr/bin/env python3

import argparse
import json
from pathlib import Path

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
    value = str(raw).strip().replace("_", "-").upper()
    return CANONICAL.get(value, value)


def merge_metric_maps(left, right):
    merged = dict(left or {})
    for key, value in (right or {}).items():
        if key not in merged:
            merged[key] = value
            continue
        left_months = len((merged[key] or {}).get("monthly", {}))
        right_months = len((value or {}).get("monthly", {}))
        if right_months > left_months:
            merged[key] = value
    return merged


def normalize_monthly_data(source):
    output = {}
    for _, item in (source or {}).items():
        canonical = canonical_country(item.get("country"))
        key = f'{item.get("content_type")}_{item.get("category")}_{canonical}'
        normalized = dict(item)
        normalized["country"] = canonical
        if key in output:
            output[key]["metrics_2026"] = merge_metric_maps(output[key].get("metrics_2026"), normalized.get("metrics_2026"))
            output[key]["insights"] = list(dict.fromkeys((output[key].get("insights") or []) + (normalized.get("insights") or [])))
        else:
            output[key] = normalized
    return output


def average(values):
    valid = [value for value in values if isinstance(value, (int, float))]
    return sum(valid) / len(valid) if valid else None


def normalize_country_reports(source):
    output = {}
    for raw_code, report in (source or {}).items():
        canonical = canonical_country(raw_code)
        if canonical not in output:
            output[canonical] = dict(report)
            output[canonical]["links"] = list(report.get("links") or [])
            output[canonical]["analyst_notes"] = list(report.get("analyst_notes") or [])
            continue

        current = output[canonical]
        current["total_sessions"] = (current.get("total_sessions") or 0) + (report.get("total_sessions") or 0)
        current["total_clicks"] = (current.get("total_clicks") or 0) + (report.get("total_clicks") or 0)
        current["total_purchases"] = (current.get("total_purchases") or 0) + (report.get("total_purchases") or 0)
        current["avg_plp_conv"] = average([current.get("avg_plp_conv"), report.get("avg_plp_conv")])
        current["avg_purchase_conv"] = average([current.get("avg_purchase_conv"), report.get("avg_purchase_conv")])
        current["avg_duration"] = average([current.get("avg_duration"), report.get("avg_duration")])
        current["session_change_pct"] = average([current.get("session_change_pct"), report.get("session_change_pct")])
        current["links"] = list({item.get("link"): item for item in current["links"] + list(report.get("links") or []) if item.get("link")}.values())
        current["analyst_notes"] = current["analyst_notes"] + list(report.get("analyst_notes") or [])
        current["per_product"] = {**current.get("per_product", {}), **report.get("per_product", {})}
    return output


def normalize_pages(pages):
    result = []
    for page in pages or []:
        link = page.get("link")
        if not link or not str(link).startswith("http"):
            continue
        result.append(page)
    return result


def main():
    parser = argparse.ArgumentParser(description="Normalize the live dashboard snapshot while keeping feature_library data.")
    parser.add_argument("--input", default=str(ROOT / "data" / "live-data.raw.json"))
    parser.add_argument("--output", default=str(ROOT / "data" / "data.json"))
    args = parser.parse_args()

    with open(args.input, "r", encoding="utf-8") as stream:
        payload = json.load(stream)

    payload["monthly_data"] = normalize_monthly_data(payload.get("monthly_data"))
    payload["pages"] = normalize_pages(payload.get("pages"))
    payload["expert"]["country_reports"] = normalize_country_reports(payload.get("expert", {}).get("country_reports"))
    payload["meta"]["countries"] = sorted({canonical_country(code) for code in payload.get("meta", {}).get("countries", [])})
    payload["meta"]["sheets_parsed"] = list(payload["monthly_data"].keys())

    with open(args.output, "w", encoding="utf-8") as stream:
      json.dump(payload, stream, ensure_ascii=False, indent=2)

    print(f"Wrote {args.output}")


if __name__ == "__main__":
    main()
