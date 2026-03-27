IGNORED_SHEETS = {
    "라인업 가이드 시나리오",
    "AI",
}

MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

SHEET_LAYOUT = {
    "label_column": 3,
    "section_column": 2,
    "month_scan_rows": 450,
    "empty_row_break": 40,
}

COUNTRY_CANONICAL_MAP = {
    "CA_EN": "CA-EN",
    "CA_FR": "CA-FR",
    "CA-EN": "CA-EN",
    "CA-FR": "CA-FR",
}

METRIC_ALIASES = {
    "session": ["lineup_session", "line_up_session", "session"],
    "organic_session": ["organic"],
    "external_entrance": ["external_entrance"],
    "avg_session_duration_sec": ["avg_session_duration_sec", "avg_session_duration"],
    "contents_click": ["contents_click"],
    "plp_conversion": ["plp_conversion"],
    "purchase_conversion": ["purchase_conversion"],
    "internal": ["internal"],
    "event_click": ["event_click"],
    "engagement_rate": ["engagemet_rate", "engagement_rate"],
    "exit_rate": ["exit_rate"],
    "product_conversion": ["product_conversion"],
    "purchase_intent_conversion": ["purchase_intent_conversion", "purchase_intent"],
    "page_view_per_session": ["page_view_per_session"],
    "page_view": ["line_up", "lineup_guide", "page_view"],
    "plp": ["plp"],
    "pdp": ["pdp_pbp", "pdp"],
    "add_to_cart_checkout": ["add_to_cart", "check_out", "checkout"],
    "purchase": ["purchase"],
}

NARRATIVE_TEMPLATES = [
    {
        "title": "Top markets should drive the optimization agenda",
        "title_ko": "상위 시장 중심으로 최적화 우선순위 설정",
        "priority": "high",
    },
    {
        "title": "Low-organic pages need acquisition diversification",
        "title_ko": "오가닉 비중이 낮은 페이지의 유입 다변화 필요",
        "priority": "high",
    },
    {
        "title": "High-traffic but weak-conversion pages need CTA work",
        "title_ko": "트래픽 대비 전환이 낮은 페이지는 CTA 개선 필요",
        "priority": "medium",
    },
]
