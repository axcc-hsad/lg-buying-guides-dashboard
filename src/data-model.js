import { CONTENT_TYPE_LABELS, COUNTRY_CANONICAL_MAP, COUNTRY_NAME_MAP } from "./config.js";

function canonicalCountry(raw) {
  if (!raw) return raw;
  const normalized = String(raw).trim().replace(/_/g, "-").toUpperCase();
  return COUNTRY_CANONICAL_MAP[normalized] || normalized;
}

function parseCountryFromLabel(label) {
  const match = String(label || "").match(/\(([A-Z_-]+)\)/);
  return match ? canonicalCountry(match[1]) : null;
}

function dedupeBy(items, keyFn) {
  const seen = new Set();
  return items.filter((item) => {
    const key = keyFn(item);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function average(values) {
  const valid = values.filter((value) => typeof value === "number" && !Number.isNaN(value));
  return valid.length ? valid.reduce((sum, value) => sum + value, 0) / valid.length : null;
}

function mergeMetricMaps(base = {}, incoming = {}) {
  const merged = { ...base };
  Object.entries(incoming).forEach(([key, value]) => {
    if (!merged[key]) {
      merged[key] = value;
      return;
    }
    const existingMonths = Object.keys(merged[key].monthly || {}).length;
    const nextMonths = Object.keys(value.monthly || {}).length;
    if (nextMonths > existingMonths) merged[key] = value;
  });
  return merged;
}

function mergeSheetRecord(base, incoming) {
  return {
    category: incoming.category || base.category,
    country: canonicalCountry(incoming.country || base.country),
    content_type: incoming.content_type || base.content_type,
    metrics_2026: mergeMetricMaps(base.metrics_2026, incoming.metrics_2026),
    metrics_2025: mergeMetricMaps(base.metrics_2025, incoming.metrics_2025),
    insights: dedupeBy([...(base.insights || []), ...(incoming.insights || [])], (item) => item),
  };
}

function mergeCountryReport(base = {}, incoming = {}) {
  return {
    total_sessions: (base.total_sessions || 0) + (incoming.total_sessions || 0),
    total_clicks: (base.total_clicks || 0) + (incoming.total_clicks || 0),
    total_purchases: (base.total_purchases || 0) + (incoming.total_purchases || 0),
    avg_plp_conv: average([base.avg_plp_conv, incoming.avg_plp_conv]),
    avg_purchase_conv: average([base.avg_purchase_conv, incoming.avg_purchase_conv]),
    avg_duration: average([base.avg_duration, incoming.avg_duration]),
    session_change_pct: average([base.session_change_pct, incoming.session_change_pct]),
    status: incoming.status || base.status,
    links: dedupeBy([...(base.links || []), ...(incoming.links || [])], (item) => item.link || JSON.stringify(item)),
    analyst_notes: dedupeBy([...(base.analyst_notes || []), ...(incoming.analyst_notes || [])], (item) => JSON.stringify(item)),
    per_product: { ...(base.per_product || {}), ...(incoming.per_product || {}) },
  };
}

function normalizeCountryReports(reports = {}) {
  const result = {};
  Object.entries(reports).forEach(([countryCode, report]) => {
    const canonical = canonicalCountry(countryCode);
    result[canonical] = result[canonical]
      ? mergeCountryReport(result[canonical], report)
      : { ...report, links: [...(report.links || [])], analyst_notes: [...(report.analyst_notes || [])] };
  });
  return result;
}

function normalizePages(pages = []) {
  return pages
    .filter((page) => page && page.link && String(page.link).startsWith("http"))
    .map((page) => {
      const code = parseCountryFromLabel(page.country);
      return { ...page, country_code: code };
    });
}

export function normalizeData(raw) {
  const monthlyData = {};
  Object.entries(raw.monthly_data || {}).forEach(([rawKey, value]) => {
    const canonical = canonicalCountry(value.country);
    const normalizedKey = `${value.content_type}_${value.category}_${canonical}`;
    const normalizedValue = { ...value, country: canonical };
    monthlyData[normalizedKey] = monthlyData[normalizedKey]
      ? mergeSheetRecord(monthlyData[normalizedKey], normalizedValue)
      : normalizedValue;
  });

  const pages = normalizePages(raw.pages || []);
  const countryReports = normalizeCountryReports(raw.expert?.country_reports || {});
  const countriesFromSheets = Object.values(monthlyData).map((item) => item.country);
  const countriesFromReports = Object.keys(countryReports);

  return {
    ...raw,
    pages,
    monthly_data: monthlyData,
    meta: {
      ...(raw.meta || {}),
      countries: [...new Set([...(raw.meta?.countries || []), ...countriesFromSheets, ...countriesFromReports].map(canonicalCountry))].sort(),
      sheets_parsed: Object.keys(monthlyData),
    },
    expert: {
      ...(raw.expert || {}),
      country_reports: countryReports,
    },
  };
}

export function getCountryName(code, lang) {
  return COUNTRY_NAME_MAP[lang]?.[code] || code;
}

export function getContentTypeName(code, lang) {
  return CONTENT_TYPE_LABELS[code]?.[lang] || code;
}

export function getMonths(data) {
  return data?.meta?.months_available || [];
}

export function getLatestMonth(data) {
  const months = getMonths(data);
  return months[months.length - 1] || null;
}

export function getPrevMonth(data) {
  const months = getMonths(data);
  return months.length > 1 ? months[months.length - 2] : null;
}

export function getFilteredSheets(data, state) {
  return Object.entries(data.monthly_data || {}).filter(([, value]) => {
    if (state.selectedContentType !== "ALL" && value.content_type !== state.selectedContentType) return false;
    if (state.selectedCountry !== "ALL" && value.country !== state.selectedCountry) return false;
    if (state.selectedCategory !== "ALL" && value.category !== state.selectedCategory) return false;
    return true;
  });
}

export function findMetric(metrics = {}, ...keywords) {
  const lowered = keywords.map((keyword) => String(keyword).toLowerCase());
  return Object.entries(metrics).find(([key]) => lowered.every((keyword) => key.toLowerCase().includes(keyword)))?.[1] || null;
}

export function getMetricValue(metrics = {}, month, ...keywords) {
  const direct = findMetric(metrics, ...keywords);
  if (direct?.monthly?.[month] != null) return direct.monthly[month];

  const replacements = keywords.map((keyword) => {
    if (keyword === "lineup") return "spec_library";
    if (keyword === "tv_lineup") return "tv_spec_library";
    if (keyword === "monitor_lineup") return "monitor_spec_library";
    if (keyword === "guide_event_click") return "event_click";
    if (keyword === "external") return "external_entrance";
    if (keyword === "purchase_intent") return "purchase_intent_conversion";
    return keyword;
  });
  const fallback = findMetric(metrics, ...replacements);
  return fallback?.monthly?.[month] ?? null;
}
