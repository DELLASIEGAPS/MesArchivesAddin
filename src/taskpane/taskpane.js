const ARCHIVE_YEARS = ["2021", "2022", "2023", "2024"];
const CONFIG = window.MES_ARCHIVES_CONFIG || {};

document.addEventListener("DOMContentLoaded", () => {
  const view = document.getElementById("agent-view");
  if (CONFIG.agentViewUrl) {
    view.href = CONFIG.agentViewUrl;
  } else {
    view.removeAttribute("href");
    view.setAttribute("aria-disabled", "true");
  }

  const container = document.getElementById("archives");
  ARCHIVE_YEARS.forEach((year) => {
    const row = document.createElement("div");
    row.className = "archive";
    const label = document.createElement("span");
    label.textContent = `Archives ${year}`;
    row.append(label);
    container.append(row);
  });
});
