const charts = new Map();

export function destroyCharts() {
  charts.forEach((chart) => chart.destroy());
  charts.clear();
}

export function getChartTheme() {
  const style = getComputedStyle(document.documentElement);
  return {
    grid: style.getPropertyValue("--chart-grid").trim(),
    text: style.getPropertyValue("--chart-text").trim(),
  };
}

export function createChart(id, type, data, options = {}) {
  const canvas = document.getElementById(id);
  if (!canvas || !window.Chart) return;

  charts.get(id)?.destroy();
  const theme = getChartTheme();

  const chart = new window.Chart(canvas, {
    type,
    data,
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          labels: { color: theme.text },
        },
      },
      scales: {
        x: {
          grid: { color: theme.grid },
          ticks: { color: theme.text },
        },
        y: {
          grid: { color: theme.grid },
          ticks: { color: theme.text },
        },
      },
      ...options,
    },
  });
  charts.set(id, chart);
}
