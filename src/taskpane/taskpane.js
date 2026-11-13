const ARCHIVE_YEARS = ["2021", "2022", "2023", "2024"];
const STATE_KEY = "archivesMounted";
const CONFIG = window.MES_ARCHIVES_CONFIG || {};

function renderArchives(states) {
  const container = document.getElementById("archives");
  container.replaceChildren();

  ARCHIVE_YEARS.forEach((year) => {
    const mounted = states[year] === true;
    const row = document.createElement("div");
    row.className = `archive ${mounted ? "mounted" : "unmounted"}`;

    const label = document.createElement("span");
    label.textContent = `Archives ${year}`;
    const state = document.createElement("strong");
    state.textContent = mounted ? "Vert - montée" : "Rouge - démontée";
    row.append(label, state);
    container.append(row);
  });
}

async function syncSharePointStates() {
  const status = document.getElementById("sync-status");
  if (!CONFIG.webStatusEndpoint) {
    status.textContent = "État local uniquement : endpoint de statut SharePoint non configuré.";
    return;
  }

  try {
    const results = await Promise.all(ARCHIVE_YEARS.map(async (year) => {
      const url = new URL(CONFIG.webStatusEndpoint);
      if (url.protocol !== "https:") {
        throw new Error("Le point d'accès SharePoint doit utiliser HTTPS.");
      }
      url.searchParams.set("year", year);
      const response = await fetch(url);
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const data = await response.json();
      return [year, data.mounted === true];
    }));

    const states = Object.fromEntries(results);
    Office.context.roamingSettings.set(STATE_KEY, states);
    await new Promise((resolve, reject) => {
      Office.context.roamingSettings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(result.error || new Error("Impossible d'enregistrer l'état SharePoint."));
        }
      });
    });
    renderArchives(states);
    status.textContent = "État synchronisé avec SharePoint.";
  } catch (error) {
    status.textContent = `Synchronisation SharePoint impossible : ${error.message}`;
    console.error("Échec de synchronisation des archives.", error);
  }
}

document.addEventListener("DOMContentLoaded", () => {
  const upload = document.getElementById("upload-pst");
  if (CONFIG.uploadFolderUrl) {
    upload.href = CONFIG.uploadFolderUrl;
  } else {
    upload.removeAttribute("href");
    upload.setAttribute("aria-disabled", "true");
  }

  Office.onReady(() => {
    renderArchives(Office.context.roamingSettings.get(STATE_KEY) || {});
    syncSharePointStates();
  });
});
