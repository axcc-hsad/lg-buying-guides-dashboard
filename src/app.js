import { DATA_URL, COUNTRY_COLORS } from "./config.js";
import { I18N } from "./i18n.js";
import {
  normalizeData,
  getCountryName,
  getContentTypeName,
  getMonths,
  getLatestMonth,
  getPrevMonth,
  getFilteredSheets,
  getMetricValue,
} from "./data-model.js";
import { createChart, destroyCharts } from "./charts.js";

const state = {
  data: null,
  currentLang: localStorage.getItem("lg_dash_lang") || "en",
  currentTheme: localStorage.getItem("lg_dash_theme") || "dark",
  activePrimaryTab: "brief",
  activeDetailTab: "overview",
  selectedCountry: "ALL",
  selectedCategory: "ALL",
  selectedContentType: "ALL",
};

function t(key) {
  return I18N[state.currentLang]?.[key] || I18N.en[key] || key;
}

function formatNumber(value, digits = 0) {
  if (value == null || Number.isNaN(value)) return "—";
  if (typeof value === "string") return value;
  if (Math.abs(value) < 1 && value !== 0) {
    return `${(value * 100).toFixed(1)}%`;
  }
  return value.toLocaleString(state.currentLang === "ko" ? "ko-KR" : "en-US", {
    maximumFractionDigits: digits,
  });
}

function calcChange(current, previous) {
  if (current == null || previous == null || previous === 0) return null;
  return ((current - previous) / Math.abs(previous)) * 100;
}

function sum(values) {
  return values.reduce((total, value) => total + (value || 0), 0);
}

function average(values) {
  const valid = values.filter((value) => value != null && !Number.isNaN(value));
  return valid.length ? valid.reduce((total, value) => total + value, 0) / valid.length : null;
}

function countryName(code) {
  return getCountryName(code, state.currentLang);
}

function contentTypeName(code) {
  return getContentTypeName(code, state.currentLang);
}

function showToast(message) {
  const toast = document.getElementById("toast");
  toast.textContent = message;
  toast.classList.add("show");
  clearTimeout(showToast._timeout);
  showToast._timeout = setTimeout(() => toast.classList.remove("show"), 2200);
}

function kpiCard(label, value, change) {
  const changeClass = change == null ? "neutral" : change > 0 ? "up" : change < 0 ? "down" : "neutral";
  const changeText = change == null ? "" : `${change > 0 ? "+" : ""}${change.toFixed(1)}% MoM`;
  return `
    <div class="kpi-card">
      <div class="kpi-label">${label}</div>
      <div class="kpi-value">${typeof value === "number" ? formatNumber(value) : value}</div>
      <div class="kpi-change ${changeClass}">${changeText}</div>
    </div>
  `;
}

async function loadData() {
  const response = await fetch(DATA_URL);
  if (!response.ok) throw new Error(`Failed to load ${DATA_URL}`);
  const raw = await response.json();
  state.data = normalizeData(raw);
  return state.data;
}

async function callAnthropic(prompt) {
  const apiKey = localStorage.getItem("lg_dash_api_key");
  if (!apiKey) throw new Error("No API key");

  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
      "anthropic-dangerous-direct-browser-access": "true",
    },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 1600,
      messages: [{ role: "user", content: buildBriefPrompt() }],
    }),
  });

  if (!response.ok) throw new Error(`API error ${response.status}`);
  const payload = await response.json();
  return payload.content?.[0]?.text || "[]";
}

function buildBriefPrompt() {
  const expert = state.data.expert || {};
  const executive = expert.executive || {};
  const countryReports = expert.country_reports || {};
  const countries = Object.entries(countryReports)
    .sort((a, b) => (b[1].total_sessions || 0) - (a[1].total_sessions || 0))
    .slice(0, 6)
    .map(([code, report]) =>
      `${countryName(code)}: sessions ${report.total_sessions || 0}, MoM ${(report.session_change_pct || 0).toFixed(1)}%, PLP ${((report.avg_plp_conv || 0) * 100).toFixed(1)}%, status ${report.status}`
    )
    .join("\n");

  return `You are a digital analytics lead. Read the dashboard summary and return exactly 5 JSON items.
Overall sessions: ${executive.total_sessions || 0}
Overall session change: ${(executive.session_change_pct || 0).toFixed(1)}%
Average PLP conversion: ${((executive.avg_plp_conv || 0) * 100).toFixed(1)}%
Average purchase conversion: ${((executive.avg_purchase_conv || 0) * 100).toFixed(2)}%
Countries:
${countries}
Rule:
- Return only a JSON array
- Each item must include signal, text_en, text_ko, detail_en, detail_ko, data_tab
- signal must be alert, watch, or good`;
}

async function renderIntelBrief() {
  const panel = document.getElementById("intelBriefPanel");
  const apiKey = localStorage.getItem("lg_dash_api_key");

  if (!apiKey) {
    panel.innerHTML = `
      <div class="intel-brief">
        <div class="intel-brief-header"><h2>${t("intelligenceBrief")}</h2></div>
        <p>${t("briefNoKey")}</p>
        <button class="btn btn-outline" id="briefSetupButton">${t("setupApi")}</button>
      </div>
    `;
    document.getElementById("briefSetupButton")?.addEventListener("click", openSettings);
    return;
  }

  const cacheKey = "lg_dash_brief_cache_v2";
  const cached = localStorage.getItem(cacheKey);
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      if (Date.now() - parsed.timestamp < 60 * 60 * 1000 && Array.isArray(parsed.items)) {
        renderBriefItems(panel, parsed.items);
        return;
      }
    } catch {}
  }

  panel.innerHTML = `<div class="intel-brief"><h2>${t("intelligenceBrief")}</h2><p>${t("briefLoading")}</p></div>`;
  try {
    const raw = await callAnthropic(buildBriefPrompt());
    const cleaned = raw.trim().replace(/^```json\s*/i, "").replace(/```$/, "");
    const items = JSON.parse(cleaned);
    localStorage.setItem(cacheKey, JSON.stringify({ timestamp: Date.now(), items }));
    renderBriefItems(panel, items);
  } catch (error) {
    panel.innerHTML = `
      <div class="intel-brief">
        <div class="intel-brief-header">
          <h2>${t("intelligenceBrief")}</h2>
          <button class="btn btn-outline" id="briefRetryButton">${t("briefRefresh")}</button>
        </div>
        <p>${error.message}</p>
      </div>
    `;
    document.getElementById("briefRetryButton")?.addEventListener("click", renderIntelBrief);
  }
}

function renderBriefItems(panel, items) {
  panel.innerHTML = `
    <div class="intel-brief">
      <div class="intel-brief-header">
        <h2>${t("intelligenceBrief")}</h2>
        <button class="btn btn-outline" id="briefRefreshButton">${t("briefRefresh")}</button>
      </div>
      ${items.map((item, index) => {
        const signalClass = item.signal === "alert" ? "alert" : item.signal === "good" ? "good" : "watch";
        const text = state.currentLang === "ko" ? item.text_ko || item.text_en : item.text_en;
        const detail = state.currentLang === "ko" ? item.detail_ko || item.detail_en : item.detail_en;
        return `
          <div class="brief-item">
            <div class="brief-signal priority-badge ${signalClass}">${item.signal.toUpperCase()}</div>
            <div class="brief-text">${text}</div>
            <button class="see-data-btn" data-brief-index="${index}">${t("seeData")}</button>
            <div class="brief-detail" id="brief-detail-${index}">${detail}</div>
          </div>
        `;
      }).join("")}
    </div>
  `;
  document.getElementById("briefRefreshButton")?.addEventListener("click", () => {
    localStorage.removeItem("lg_dash_brief_cache_v2");
    renderIntelBrief();
  });
  panel.querySelectorAll("[data-brief-index]").forEach((button) => {
    button.addEventListener("click", () => {
      document.getElementById(`brief-detail-${button.dataset.briefIndex}`)?.classList.toggle("open");
    });
  });
}

function renderMainSummary() {
  const executive = state.data.expert?.executive || {};
  const reports = state.data.expert?.country_reports || {};
  const month = executive.latest_month || getLatestMonth(state.data) || "";
  const miniCards = Object.entries(reports)
    .sort((a, b) => (b[1].total_sessions || 0) - (a[1].total_sessions || 0))
    .map(([code, report]) => `
      <div class="country-status-mini" data-country-card="${code}">
        <div class="csm-name">${countryName(code)}</div>
        <div class="csm-stat">${formatNumber(report.total_sessions)} ${t("sessions")} · ${(report.session_change_pct || 0).toFixed(1)}%</div>
      </div>
    `)
    .join("");

  document.getElementById("mainInsightsSummary").innerHTML = `
    <div class="main-insight-hero">
      <h2>${t("intelligenceBrief")} · ${month}</h2>
      <div class="main-insight-kpis">
        <div class="main-insight-kpi"><div class="mik-value">${formatNumber(executive.total_sessions || 0)}</div><div class="mik-label">${t("totalSessions")}</div></div>
        <div class="main-insight-kpi"><div class="mik-value">${formatNumber(executive.total_clicks || 0)}</div><div class="mik-label">${t("totalClicks")}</div></div>
        <div class="main-insight-kpi"><div class="mik-value">${((executive.avg_plp_conv || 0) * 100).toFixed(1)}%</div><div class="mik-label">${t("avgPlp")}</div></div>
        <div class="main-insight-kpi"><div class="mik-value">${((executive.avg_purchase_conv || 0) * 100).toFixed(2)}%</div><div class="mik-label">${t("avgPurchase")}</div></div>
        <div class="main-insight-kpi"><div class="mik-value ${executive.session_change_pct >= 0 ? "up" : "down"}">${executive.session_change_pct >= 0 ? "+" : ""}${(executive.session_change_pct || 0).toFixed(1)}%</div><div class="mik-label">MoM</div></div>
        <div class="main-insight-kpi"><div class="mik-value">${executive.total_pages_tracked || 0}</div><div class="mik-label">${t("pagesTracked")}</div></div>
      </div>
      <div class="section-title"><span>${t("countryStatus")}</span></div>
      <div class="country-status-grid">${miniCards}</div>
    </div>
  `;

  document.querySelectorAll("[data-country-card]").forEach((button) => {
    button.addEventListener("click", () => {
      state.selectedCountry = button.dataset.countryCard;
      state.activePrimaryTab = "countries";
      syncTabUI();
      render();
    });
  });
}

function buildActionItems() {
  const latest = getLatestMonth(state.data);
  const previous = getPrevMonth(state.data);
  const sheets = getFilteredSheets(state.data, state);
  const items = [];

  sheets.forEach(([, sheet]) => {
    const currentSession = getMetricValue(sheet.metrics_2026, latest, "lineup", "session")
      ?? getMetricValue(sheet.metrics_2026, latest, "tv_lineup", "session")
      ?? getMetricValue(sheet.metrics_2026, latest, "monitor_lineup", "session");
    const previousSession = previous
      ? (getMetricValue(sheet.metrics_2026, previous, "lineup", "session")
        ?? getMetricValue(sheet.metrics_2026, previous, "tv_lineup", "session")
        ?? getMetricValue(sheet.metrics_2026, previous, "monitor_lineup", "session"))
      : null;
    const sessionChange = calcChange(currentSession, previousSession);
    const plp = getMetricValue(sheet.metrics_2026, latest, "plp_conversion");

    if (sessionChange != null && sessionChange <= -20) {
      items.push({
        title: `${sheet.category} ${countryName(sheet.country)}`,
        body: `Sessions dropped ${sessionChange.toFixed(1)}% MoM. Audit link health, SEO demand, and recent content changes.`,
      });
    } else if (plp != null && plp < 0.15) {
      items.push({
        title: `${sheet.category} ${countryName(sheet.country)}`,
        body: `PLP conversion is ${formatNumber(plp)}. Strengthen in-content product CTAs and comparison entry points.`,
      });
    }
  });
  return items.slice(0, 6);
}

function renderActions() {
  const items = buildActionItems();
  document.getElementById("actionTodosPanel").innerHTML = `
    <div class="main-actions-wrap">
      <h3>${t("actionItems")}</h3>
      <div class="todo-stack">
        ${items.length ? items.map((item) => `
          <div class="todo-item">
            <div class="todo-item-title">${item.title}</div>
            <div>${item.body}</div>
          </div>
        `).join("") : `<div class="todo-item">${t("noData")}</div>`}
      </div>
    </div>
  `;
}

function renderDirection() {
  const actions = state.data.narrative?.content_actions || [];
  document.getElementById("contentDirectionPanel").innerHTML = `
    <div class="main-actions-wrap">
      <h3>${t("contentDirection")}</h3>
      <div class="direction-stack">
        ${actions.map((action) => `
          <div class="direction-card">
            <h4>${state.currentLang === "ko" ? action.title_ko || action.title : action.title}</h4>
            <div>${state.currentLang === "ko" ? action.body_ko || action.body : action.body}</div>
          </div>
        `).join("")}
      </div>
    </div>
  `;
}

function renderCountryCards() {
  const reports = state.data.expert?.country_reports || {};
  const cards = Object.entries(reports)
    .filter(([code]) => state.selectedCountry === "ALL" || state.selectedCountry === code)
    .sort((a, b) => (b[1].total_sessions || 0) - (a[1].total_sessions || 0))
    .map(([code, report]) => `
      <div class="country-card">
        <div class="country-card-header">
          <h3>${countryName(code)}</h3>
          <span class="status-pill ${report.status || "stable"}">${t(report.status || "stable")}</span>
        </div>
        <div class="country-metrics">
          <div class="country-metric"><div class="cm-value">${formatNumber(report.total_sessions)}</div><div class="cm-label">${t("sessions")}</div></div>
          <div class="country-metric"><div class="cm-value">${formatNumber(report.total_clicks)}</div><div class="cm-label">${t("clicks")}</div></div>
          <div class="country-metric"><div class="cm-value">${((report.avg_plp_conv || 0) * 100).toFixed(1)}%</div><div class="cm-label">${t("plpConv")}</div></div>
          <div class="country-metric"><div class="cm-value">${((report.avg_purchase_conv || 0) * 100).toFixed(2)}%</div><div class="cm-label">${t("purchaseConv")}</div></div>
          <div class="country-metric"><div class="cm-value">${formatNumber(report.avg_duration, 0)}s</div><div class="cm-label">${t("avgDuration")}</div></div>
        </div>
        <div class="section-title"><span>${t("pageLinks")}</span></div>
        <div class="link-chips">
          ${(report.links || []).slice(0, 8).map((link) => `<a class="link-chip" href="${link.link}" target="_blank" rel="noreferrer">${link.category || "General"} · ${link.type}</a>`).join("")}
        </div>
        <div class="section-title"><span>${t("analystNotes")}</span></div>
        ${(report.analyst_notes || []).length ? report.analyst_notes.map((note) => `
          <div class="analyst-note">
            <strong>${note.page}</strong><br>
            ${(note.notes || []).join("<br>")}
          </div>
        `).join("") : `<div class="analyst-note">${t("noData")}</div>`}
      </div>
    `)
    .join("");

  document.getElementById("signalCardsPanel").innerHTML = `<div class="country-grid">${cards || `<div class="summary-card">${t("noData")}</div>`}</div>`;
}

function renderCategories() {
  const categories = state.data.expert?.category_reports || {};
  const cards = Object.entries(categories)
    .filter(([category]) => state.selectedCategory === "ALL" || state.selectedCategory === category)
    .map(([category, report]) => `
      <div class="category-card">
        <h3>${category}</h3>
        <div class="country-metrics">
          <div class="country-metric"><div class="cm-value">${formatNumber(report.total_sessions)}</div><div class="cm-label">${t("totalSessions")}</div></div>
          <div class="country-metric"><div class="cm-value">${((report.avg_plp_conv || 0) * 100).toFixed(1)}%</div><div class="cm-label">${t("avgPlp")}</div></div>
          <div class="country-metric"><div class="cm-value">${report.countries_count || 0}</div><div class="cm-label">${t("countries")}</div></div>
        </div>
        <div class="charts-grid">
          <div>
            <div class="section-title"><span>${t("sessionRank")}</span></div>
            <table class="rank-table"><tbody>${(report.ranked_sessions || []).slice(0, 5).map((item, index) => `<tr><td>${index + 1}</td><td>${countryName(item[0])}</td><td>${formatNumber(item[1])}</td></tr>`).join("")}</tbody></table>
          </div>
          <div>
            <div class="section-title"><span>${t("plpRank")}</span></div>
            <table class="rank-table"><tbody>${(report.ranked_plp || []).slice(0, 5).map((item, index) => `<tr><td>${index + 1}</td><td>${countryName(item[0])}</td><td>${(item[1] * 100).toFixed(1)}%</td></tr>`).join("")}</tbody></table>
          </div>
        </div>
      </div>
    `)
    .join("");

  document.getElementById("categoryPanel").innerHTML = `<div class="category-grid">${cards || `<div class="summary-card">${t("noData")}</div>`}</div>`;
}

function renderOverview() {
  const latest = getLatestMonth(state.data);
  const previous = getPrevMonth(state.data);
  const sheets = getFilteredSheets(state.data, state);

  const sessions = sum(sheets.map(([, sheet]) =>
    getMetricValue(sheet.metrics_2026, latest, "lineup", "session")
    ?? getMetricValue(sheet.metrics_2026, latest, "tv_lineup", "session")
    ?? getMetricValue(sheet.metrics_2026, latest, "monitor_lineup", "session")));
  const previousSessions = previous ? sum(sheets.map(([, sheet]) =>
    getMetricValue(sheet.metrics_2026, previous, "lineup", "session")
    ?? getMetricValue(sheet.metrics_2026, previous, "tv_lineup", "session")
    ?? getMetricValue(sheet.metrics_2026, previous, "monitor_lineup", "session"))) : null;
  const clicks = sum(sheets.map(([, sheet]) => getMetricValue(sheet.metrics_2026, latest, "event_click") ?? 0));
  const previousClicks = previous ? sum(sheets.map(([, sheet]) => getMetricValue(sheet.metrics_2026, previous, "event_click") ?? 0)) : null;
  const plpAverage = average(sheets.map(([, sheet]) => getMetricValue(sheet.metrics_2026, latest, "plp_conversion")));
  const prevPlpAverage = previous ? average(sheets.map(([, sheet]) => getMetricValue(sheet.metrics_2026, previous, "plp_conversion"))) : null;

  document.getElementById("latestMonthBadge").textContent = latest || "";
  document.getElementById("kpiGrid").innerHTML = [
    kpiCard(t("totalSessions"), sessions, calcChange(sessions, previousSessions)),
    kpiCard(t("totalClicks"), clicks, calcChange(clicks, previousClicks)),
    kpiCard(t("avgPlp"), plpAverage ?? "—", calcChange(plpAverage, prevPlpAverage)),
    kpiCard(t("pagesTracked"), sheets.length, null),
  ].join("");

  document.getElementById("overviewCharts").innerHTML = `
    <div class="chart-card"><h3>${t("totalSessions")}</h3><div style="height:280px"><canvas id="chartSessionsByCountry"></canvas></div></div>
    <div class="chart-card"><h3>${t("avgPlp")}</h3><div style="height:280px"><canvas id="chartPlpByPage"></canvas></div></div>
    <div class="chart-card"><h3>${t("eventClicks")}</h3><div style="height:280px"><canvas id="chartClicksByPage"></canvas></div></div>
    <div class="chart-card"><h3>${t("avgDuration")}</h3><div style="height:280px"><canvas id="chartDurationByPage"></canvas></div></div>
  `;

  const sessionByCountry = {};
  sheets.forEach(([, sheet]) => {
    const value = getMetricValue(sheet.metrics_2026, latest, "lineup", "session")
      ?? getMetricValue(sheet.metrics_2026, latest, "tv_lineup", "session")
      ?? getMetricValue(sheet.metrics_2026, latest, "monitor_lineup", "session")
      ?? 0;
    sessionByCountry[sheet.country] = (sessionByCountry[sheet.country] || 0) + value;
  });

  createChart("chartSessionsByCountry", "bar", {
    labels: Object.keys(sessionByCountry).map(countryName),
    datasets: [{
      data: Object.entries(sessionByCountry).map(([, value]) => value),
      backgroundColor: Object.keys(sessionByCountry).map((code) => COUNTRY_COLORS[code] || "#4969ff"),
      borderRadius: 8,
    }],
  }, { plugins: { legend: { display: false } } });

  const pageLabels = sheets.map(([, sheet]) => `${sheet.category} ${countryName(sheet.country)}`);
  createChart("chartPlpByPage", "bar", {
    labels: pageLabels,
    datasets: [{ data: sheets.map(([, sheet]) => (getMetricValue(sheet.metrics_2026, latest, "plp_conversion") || 0) * 100), backgroundColor: "rgba(73, 105, 255, 0.75)", borderRadius: 8 }],
  }, { plugins: { legend: { display: false } } });

  createChart("chartClicksByPage", "bar", {
    labels: pageLabels,
    datasets: [{ data: sheets.map(([, sheet]) => getMetricValue(sheet.metrics_2026, latest, "event_click") || 0), backgroundColor: "rgba(34, 197, 94, 0.75)", borderRadius: 8 }],
  }, { plugins: { legend: { display: false } }, indexAxis: "y" });

  createChart("chartDurationByPage", "bar", {
    labels: pageLabels,
    datasets: [{ data: sheets.map(([, sheet]) => getMetricValue(sheet.metrics_2026, latest, "avg_session_duration") ?? getMetricValue(sheet.metrics_2026, latest, "avg_session_duration_sec") ?? 0), backgroundColor: "rgba(249, 115, 22, 0.75)", borderRadius: 8 }],
  }, { plugins: { legend: { display: false } } });

  document.getElementById("comparisonTable").innerHTML = `
    <thead>
      <tr>
        <th>${t("page")}</th>
        <th>${t("sessions")}</th>
        <th>${t("clicks")}</th>
        <th>${t("plpConv")}</th>
        <th>${t("purchaseConv")}</th>
      </tr>
    </thead>
    <tbody>
      ${sheets.map(([, sheet]) => `
        <tr>
          <td>${sheet.category} ${countryName(sheet.country)} · ${contentTypeName(sheet.content_type)}</td>
          <td>${formatNumber(getMetricValue(sheet.metrics_2026, latest, "lineup", "session")
            ?? getMetricValue(sheet.metrics_2026, latest, "tv_lineup", "session")
            ?? getMetricValue(sheet.metrics_2026, latest, "monitor_lineup", "session"))}</td>
          <td>${formatNumber(getMetricValue(sheet.metrics_2026, latest, "event_click"))}</td>
          <td>${formatNumber(getMetricValue(sheet.metrics_2026, latest, "plp_conversion"))}</td>
          <td>${formatNumber(getMetricValue(sheet.metrics_2026, latest, "purchase_conversion"))}</td>
        </tr>
      `).join("")}
    </tbody>
  `;
}

function renderAcquisition() {
  const latest = getLatestMonth(state.data);
  const previous = getPrevMonth(state.data);
  const sheets = getFilteredSheets(state.data, state);

  const organic = sum(sheets.map(([, sheet]) => getMetricValue(sheet.metrics_2026, latest, "organic") || 0));
  const external = sum(sheets.map(([, sheet]) => getMetricValue(sheet.metrics_2026, latest, "external") || 0));
  const internal = sum(sheets.map(([, sheet]) => getMetricValue(sheet.metrics_2026, latest, "internal") || 0));
  const prevOrganic = previous ? sum(sheets.map(([, sheet]) => getMetricValue(sheet.metrics_2026, previous, "organic") || 0)) : null;
  const prevExternal = previous ? sum(sheets.map(([, sheet]) => getMetricValue(sheet.metrics_2026, previous, "external") || 0)) : null;
  const prevInternal = previous ? sum(sheets.map(([, sheet]) => getMetricValue(sheet.metrics_2026, previous, "internal") || 0)) : null;

  document.getElementById("acqKpis").innerHTML = [
    kpiCard(t("organic"), organic, calcChange(organic, prevOrganic)),
    kpiCard(t("external"), external, calcChange(external, prevExternal)),
    kpiCard(t("internal"), internal, calcChange(internal, prevInternal)),
    kpiCard(t("organicShare"), external ? organic / external : null, null),
  ].join("");

  document.getElementById("acqCharts").innerHTML = `
    <div class="chart-card"><h3>${t("organic")} / ${t("external")} / ${t("internal")}</h3><div style="height:280px"><canvas id="chartTrafficMix"></canvas></div></div>
    <div class="chart-card"><h3>${t("sessions")}</h3><div style="height:280px"><canvas id="chartSessionTrend"></canvas></div></div>
  `;

  createChart("chartTrafficMix", "doughnut", {
    labels: [t("organic"), t("external"), t("internal")],
    datasets: [{ data: [organic, Math.max(external - organic, 0), internal], backgroundColor: ["rgba(34,197,94,0.8)", "rgba(73,105,255,0.8)", "rgba(234,179,8,0.8)"] }],
  });

  const months = getMonths(state.data);
  createChart("chartSessionTrend", "line", {
    labels: months,
    datasets: sheets.map(([, sheet]) => ({
      label: `${sheet.category} ${countryName(sheet.country)}`,
      data: months.map((month) =>
        getMetricValue(sheet.metrics_2026, month, "lineup", "session")
        ?? getMetricValue(sheet.metrics_2026, month, "tv_lineup", "session")
        ?? getMetricValue(sheet.metrics_2026, month, "monitor_lineup", "session")),
      borderColor: COUNTRY_COLORS[sheet.country] || "#4969ff",
      tension: 0.35,
    })),
  });
}

function renderBehavior() {
  const latest = getLatestMonth(state.data);
  const previous = getPrevMonth(state.data);
  const sheets = getFilteredSheets(state.data, state);
  const clicks = sum(sheets.map(([, sheet]) => getMetricValue(sheet.metrics_2026, latest, "event_click") || 0));
  const previousClicks = previous ? sum(sheets.map(([, sheet]) => getMetricValue(sheet.metrics_2026, previous, "event_click") || 0)) : null;
  const avgDuration = average(sheets.map(([, sheet]) => getMetricValue(sheet.metrics_2026, latest, "avg_session_duration") ?? getMetricValue(sheet.metrics_2026, latest, "avg_session_duration_sec")));
  const previousDuration = previous ? average(sheets.map(([, sheet]) => getMetricValue(sheet.metrics_2026, previous, "avg_session_duration") ?? getMetricValue(sheet.metrics_2026, previous, "avg_session_duration_sec"))) : null;

  document.getElementById("behKpis").innerHTML = [
    kpiCard(t("eventClicks"), clicks, calcChange(clicks, previousClicks)),
    kpiCard(t("avgDuration"), avgDuration ? `${Math.round(avgDuration)}s` : "—", calcChange(avgDuration, previousDuration)),
    kpiCard(t("pagesTracked"), sheets.length, null),
  ].join("");

  document.getElementById("behCharts").innerHTML = `
    <div class="chart-card"><h3>${t("eventClicks")}</h3><div style="height:280px"><canvas id="chartBehaviorClicks"></canvas></div></div>
    <div class="chart-card"><h3>${t("avgDuration")}</h3><div style="height:280px"><canvas id="chartBehaviorDuration"></canvas></div></div>
  `;

  createChart("chartBehaviorClicks", "bar", {
    labels: sheets.map(([, sheet]) => `${sheet.category} ${countryName(sheet.country)}`),
    datasets: [{ data: sheets.map(([, sheet]) => getMetricValue(sheet.metrics_2026, latest, "event_click") || 0), backgroundColor: "rgba(73,105,255,0.78)", borderRadius: 8 }],
  }, { plugins: { legend: { display: false } } });

  createChart("chartBehaviorDuration", "bar", {
    labels: sheets.map(([, sheet]) => `${sheet.category} ${countryName(sheet.country)}`),
    datasets: [{ data: sheets.map(([, sheet]) => getMetricValue(sheet.metrics_2026, latest, "avg_session_duration") ?? getMetricValue(sheet.metrics_2026, latest, "avg_session_duration_sec") ?? 0), backgroundColor: "rgba(34,197,94,0.78)", borderRadius: 8 }],
  }, { plugins: { legend: { display: false } } });
}

function renderConversion() {
  const latest = getLatestMonth(state.data);
  const previous = getPrevMonth(state.data);
  const sheets = getFilteredSheets(state.data, state);
  const plp = average(sheets.map(([, sheet]) => getMetricValue(sheet.metrics_2026, latest, "plp_conversion")));
  const prevPlp = previous ? average(sheets.map(([, sheet]) => getMetricValue(sheet.metrics_2026, previous, "plp_conversion"))) : null;
  const product = average(sheets.map(([, sheet]) => getMetricValue(sheet.metrics_2026, latest, "product_conversion")));
  const prevProduct = previous ? average(sheets.map(([, sheet]) => getMetricValue(sheet.metrics_2026, previous, "product_conversion"))) : null;
  const purchaseIntent = average(sheets.map(([, sheet]) => getMetricValue(sheet.metrics_2026, latest, "purchase_intent")));
  const prevPurchaseIntent = previous ? average(sheets.map(([, sheet]) => getMetricValue(sheet.metrics_2026, previous, "purchase_intent"))) : null;
  const purchase = average(sheets.map(([, sheet]) => getMetricValue(sheet.metrics_2026, latest, "purchase_conversion")));
  const prevPurchase = previous ? average(sheets.map(([, sheet]) => getMetricValue(sheet.metrics_2026, previous, "purchase_conversion"))) : null;

  document.getElementById("convKpis").innerHTML = [
    kpiCard(t("avgPlp"), plp, calcChange(plp, prevPlp)),
    kpiCard(t("productConv"), product, calcChange(product, prevProduct)),
    kpiCard(t("purchaseIntent"), purchaseIntent, calcChange(purchaseIntent, prevPurchaseIntent)),
    kpiCard(t("avgPurchase"), purchase, calcChange(purchase, prevPurchase)),
  ].join("");

  document.getElementById("convCharts").innerHTML = `
    <div class="chart-card"><h3>${t("avgPlp")}</h3><div style="height:280px"><canvas id="chartConversionPlp"></canvas></div></div>
    <div class="chart-card"><h3>${t("avgPurchase")}</h3><div style="height:280px"><canvas id="chartConversionPurchase"></canvas></div></div>
  `;

  const labels = sheets.map(([, sheet]) => `${sheet.category} ${countryName(sheet.country)}`);
  createChart("chartConversionPlp", "bar", {
    labels,
    datasets: [{ data: sheets.map(([, sheet]) => (getMetricValue(sheet.metrics_2026, latest, "plp_conversion") || 0) * 100), backgroundColor: "rgba(73,105,255,0.78)", borderRadius: 8 }],
  }, { plugins: { legend: { display: false } } });

  createChart("chartConversionPurchase", "bar", {
    labels,
    datasets: [{ data: sheets.map(([, sheet]) => (getMetricValue(sheet.metrics_2026, latest, "purchase_conversion") || 0) * 100), backgroundColor: "rgba(34,197,94,0.78)", borderRadius: 8 }],
  }, { plugins: { legend: { display: false } } });
}

function renderFilters() {
  const data = state.data;
  const filterRoot = document.getElementById("filters");
  filterRoot.innerHTML = `
    <button class="btn btn-outline ${state.selectedContentType === "ALL" ? "active" : ""}" data-filter="content" data-value="ALL">${t("allContent")}</button>
    ${(data.meta.content_types || []).map((type) => `<button class="btn btn-outline ${state.selectedContentType === type ? "active" : ""}" data-filter="content" data-value="${type}">${contentTypeName(type)}</button>`).join("")}
    <button class="btn btn-outline ${state.selectedCountry === "ALL" ? "active" : ""}" data-filter="country" data-value="ALL">${t("allCountries")}</button>
    ${(data.meta.countries || []).map((country) => `<button class="btn btn-outline ${state.selectedCountry === country ? "active" : ""}" data-filter="country" data-value="${country}">${countryName(country)}</button>`).join("")}
    <button class="btn btn-outline ${state.selectedCategory === "ALL" ? "active" : ""}" data-filter="category" data-value="ALL">${t("allCategories")}</button>
    ${(data.meta.categories || []).map((category) => `<button class="btn btn-outline ${state.selectedCategory === category ? "active" : ""}" data-filter="category" data-value="${category}">${category}</button>`).join("")}
  `;

  filterRoot.querySelectorAll("[data-filter]").forEach((button) => {
    button.addEventListener("click", () => {
      const filter = button.dataset.filter;
      const value = button.dataset.value;
      if (filter === "content") state.selectedContentType = value;
      if (filter === "country") state.selectedCountry = value;
      if (filter === "category") state.selectedCategory = value;
      render();
    });
  });
}

function renderPageLinksDrawer() {
  const body = document.getElementById("linksDrawerBody");
  const groups = {};
  (state.data.pages || []).forEach((page) => {
    const code = page.country_code || "OTHER";
    groups[code] ??= [];
    groups[code].push(page);
  });

  body.innerHTML = Object.entries(groups).sort(([a], [b]) => a.localeCompare(b)).map(([code, items]) => `
    <div>
      <div class="section-title"><span>${code === "OTHER" ? "Other" : countryName(code)}</span></div>
      ${items.map((page) => `
        <a class="links-item" href="${page.link}" target="_blank" rel="noreferrer" data-search="${`${code} ${page.category} ${page.type}`.toLowerCase()}">
          <strong>${page.category || "General"} · ${page.type || "-"}</strong>
          <div class="li-meta">${page.feedback || ""}</div>
        </a>
      `).join("")}
    </div>
  `).join("");
}

function renderTodaysFocus() {
  const items = (state.data.insights || []).filter((item) => item.priority === "high").slice(0, 2);
  const widget = document.getElementById("todaysFocus");
  widget.innerHTML = `
    <div class="section-title"><span>${t("todayFocus")}</span></div>
    ${items.map((item) => `<div class="todays-focus-item"><strong>${item.page}</strong><br>${item.message}</div>`).join("")}
  `;
  widget.classList.toggle("hidden", items.length === 0);
}

function renderHeaderLabels() {
  document.querySelector("#refreshBtn").textContent = t("refresh");
  document.querySelector("#linksDrawerButton").textContent = t("pageLinks");
  document.querySelector('[data-ptab="brief"]').textContent = t("intelligenceBrief");
  document.querySelector('[data-ptab="countries"]').textContent = t("countryCards");
  document.querySelector('[data-ptab="categories"]').textContent = t("categoryCompare");
  document.querySelector('[data-ptab="rawdata"]').textContent = t("rawData");
  document.querySelector('[data-dtab="overview"]').textContent = t("overview");
  document.querySelector('[data-dtab="acquisition"]').textContent = t("acquisition");
  document.querySelector('[data-dtab="behavior"]').textContent = t("behavior");
  document.querySelector('[data-dtab="conversion"]').textContent = t("conversion");
}

function syncTabUI() {
  document.querySelectorAll("#primaryTabs .primary-tab").forEach((button) => {
    const active = button.dataset.ptab === state.activePrimaryTab;
    button.classList.toggle("active", active);
    document.getElementById(`ptab-${button.dataset.ptab}`).classList.toggle("active", active);
  });
  document.querySelectorAll("#detailTabs .detail-tab").forEach((button) => {
    const active = button.dataset.dtab === state.activeDetailTab;
    button.classList.toggle("active", active);
    document.getElementById(`dtab-${button.dataset.dtab}`).classList.toggle("active", active);
  });
}

function render() {
  if (!state.data) return;
  destroyCharts();
  document.documentElement.setAttribute("data-theme", state.currentTheme);
  document.documentElement.setAttribute("data-lang", state.currentLang);

  renderHeaderLabels();
  renderFilters();
  renderPageLinksDrawer();
  renderTodaysFocus();

  const updated = state.data.meta?.last_updated;
  document.getElementById("lastUpdated").textContent = updated
    ? `${t("lastUpdated")}: ${new Date(updated).toLocaleString()}`
    : `${t("lastUpdated")}: local snapshot`;

  if (state.activePrimaryTab === "brief") {
    renderIntelBrief();
    renderMainSummary();
    renderActions();
    renderDirection();
  } else if (state.activePrimaryTab === "countries") {
    renderCountryCards();
  } else if (state.activePrimaryTab === "categories") {
    renderCategories();
  } else {
    renderOverview();
    renderAcquisition();
    renderBehavior();
    renderConversion();
  }
}

function toggleDrawer(open) {
  document.getElementById("linksDrawer").classList.toggle("open", open);
  document.getElementById("linksOverlay").classList.toggle("open", open);
}

function openSettings() {
  document.getElementById("apiKeyInput").value = localStorage.getItem("lg_dash_api_key") || "";
  document.getElementById("settingsModal").classList.add("open");
}

function closeSettings() {
  document.getElementById("settingsModal").classList.remove("open");
}

function bindEvents() {
  document.querySelectorAll("#langToggle button").forEach((button) => {
    button.addEventListener("click", () => {
      state.currentLang = button.dataset.val;
      localStorage.setItem("lg_dash_lang", state.currentLang);
      document.querySelectorAll("#langToggle button").forEach((item) => item.classList.toggle("active", item === button));
      render();
    });
  });

  document.querySelectorAll("#themeToggle button").forEach((button) => {
    button.addEventListener("click", () => {
      state.currentTheme = button.dataset.val;
      localStorage.setItem("lg_dash_theme", state.currentTheme);
      document.querySelectorAll("#themeToggle button").forEach((item) => item.classList.toggle("active", item === button));
      render();
    });
  });

  document.querySelectorAll("#primaryTabs .primary-tab").forEach((button) => {
    button.addEventListener("click", () => {
      state.activePrimaryTab = button.dataset.ptab;
      syncTabUI();
      render();
    });
  });

  document.querySelectorAll("#detailTabs .detail-tab").forEach((button) => {
    button.addEventListener("click", () => {
      state.activeDetailTab = button.dataset.dtab;
      syncTabUI();
    });
  });

  document.getElementById("linksDrawerButton").addEventListener("click", () => toggleDrawer(true));
  document.getElementById("linksClose").addEventListener("click", () => toggleDrawer(false));
  document.getElementById("linksOverlay").addEventListener("click", () => toggleDrawer(false));
  document.getElementById("settingsButton").addEventListener("click", openSettings);
  document.getElementById("closeSettingsButton").addEventListener("click", closeSettings);
  document.getElementById("clearApiKeyButton").addEventListener("click", () => {
    localStorage.removeItem("lg_dash_api_key");
    closeSettings();
    render();
  });
  document.getElementById("saveApiKeyButton").addEventListener("click", () => {
    const key = document.getElementById("apiKeyInput").value.trim();
    if (key) localStorage.setItem("lg_dash_api_key", key);
    closeSettings();
    showToast("API key saved");
    render();
  });
  document.getElementById("refreshBtn").addEventListener("click", async () => {
    const button = document.getElementById("refreshBtn");
    button.textContent = t("refreshing");
    button.disabled = true;
    try {
      await loadData();
      render();
      showToast("Data refreshed");
    } catch (error) {
      showToast(error.message);
    } finally {
      button.disabled = false;
      button.textContent = t("refresh");
    }
  });
  document.getElementById("linksSearch").addEventListener("input", (event) => {
    const query = event.target.value.toLowerCase().trim();
    document.querySelectorAll(".links-item").forEach((item) => {
      item.style.display = !query || item.dataset.search.includes(query) ? "" : "none";
    });
  });
}

async function init() {
  bindEvents();
  syncTabUI();
  document.documentElement.setAttribute("data-theme", state.currentTheme);
  document.documentElement.setAttribute("data-lang", state.currentLang);
  document.querySelectorAll("#langToggle button").forEach((button) => button.classList.toggle("active", button.dataset.val === state.currentLang));
  document.querySelectorAll("#themeToggle button").forEach((button) => button.classList.toggle("active", button.dataset.val === state.currentTheme));

  try {
    await loadData();
    render();
  } catch (error) {
    showToast(error.message);
  }
}

init();
