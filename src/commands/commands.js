const STATE_KEY = "archivesMounted";
const PST_AGENT = "http://127.0.0.1:45731";
const CONFIG = window.MES_ARCHIVES_CONFIG || {};

const ARCHIVE_CONFIG = Object.freeze({
  "2021": {
    fileName: "2021.pst",
    localPath: "%OneDriveCommercial%\\BdF\\PST\\2021.pst",
    sharePointPath: "Shared Documents/PST/2021"
  },
  "2022": {
    fileName: "2022.pst",
    localPath: "%OneDriveCommercial%\\BdF\\PST\\2022.pst",
    sharePointPath: "Shared Documents/PST/2022"
  },
  "2023": {
    fileName: "2023.pst",
    localPath: "%OneDriveCommercial%\\BdF\\PST\\2023.pst",
    sharePointPath: "Shared Documents/PST/2023"
  },
  "2024": {
    fileName: "ADR.pst",
    localPath: "%OneDriveCommercial%\\Documents\\Personnel\\ADR.pst",
    sharePointPath: "Shared Documents/PST/2024"
  }
});

function readStates() {
  return Office.context.roamingSettings.get(STATE_KEY) || {};
}

function saveStates(states) {
  Office.context.roamingSettings.set(STATE_KEY, states);
  return new Promise((resolve, reject) => {
    Office.context.roamingSettings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(result.error || new Error("Impossible d'enregistrer l'état des archives."));
      }
    });
  });
}

const YEARS = ["2021", "2022", "2023", "2024"];

// Flip-flop : un seul bouton visible par année (rouge = monter, vert = démonter).
async function setRibbonState(states) {
  const build = (withVisible) => ({
    tabs: [{
      id: "TabDefault",
      groups: [{
        id: "MesArchivesGroup",
        controls: YEARS.flatMap((year) => {
          const m = states[year] === true;
          return [
            Object.assign({ id: `archive${year}_mount`, enabled: !m }, withVisible ? { visible: !m } : {}),
            Object.assign({ id: `archive${year}_unmount`, enabled: m }, withVisible ? { visible: m } : {})
          ];
        })
      }]
    }]
  });
  try {
    await Office.ribbon.requestUpdate(build(true));
  } catch (error) {
    console.warn("Show/hide non supporté, repli sur enabled.", error);
    await Office.ribbon.requestUpdate(build(false));
  }
}

function notify(message) {
  const item = Office.context.mailbox && Office.context.mailbox.item;
  if (!item || !item.notificationMessages) {
    console.info(message);
    return;
  }

  item.notificationMessages.replaceAsync("mesArchives", {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: message.substring(0, 150),
    icon: "IconVSTO.32",
    persistent: false
  }, (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("Échec de la notification Mes Archives.", result.error);
    }
  });
}

async function callAgent(route, fileName) {
  const response = await fetch(`${PST_AGENT}${route}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ fileName }),
    signal: AbortSignal.timeout(20000)
  });
  const data = await response.json();
  if (!response.ok || data.ok === false) {
    throw new Error(data.error || `Agent PST : HTTP ${response.status}`);
  }
  return data;
}

function statusUrl(year) {
  if (!CONFIG.webStatusEndpoint) {
    throw new Error("Le point d'accès de statut SharePoint n'est pas configuré.");
  }
  const url = new URL(CONFIG.webStatusEndpoint);
  if (url.protocol !== "https:") {
    throw new Error("Le point d'accès SharePoint doit utiliser HTTPS.");
  }
  url.searchParams.set("year", year);
  return url;
}

async function getWebArchiveStatus(year) {
  const response = await fetch(statusUrl(year));
  if (!response.ok) throw new Error(`Lecture SharePoint impossible : HTTP ${response.status}`);
  return response.json();
}

async function updateWebArchiveStatus(year, mounted) {
  const url = statusUrl(year);
  url.searchParams.delete("year");
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      annee: Number(year),
      Monte: mounted,
      Couleur: mounted ? "Vert" : "Rouge",
      sharePointPath: ARCHIVE_CONFIG[year].sharePointPath
    })
  });
  if (!response.ok) throw new Error(`Mise à jour SharePoint impossible : HTTP ${response.status}`);
}

async function setArchiveMounted(year, mounted) {
  const states = readStates();
  states[year] = mounted;
  await saveStates(states);
  await setRibbonState(states);
}

// Lit l'état réel (PST présent et monté) ; retourne null si la source est injoignable.
async function detectState(year) {
  try {
    if (isDesktopOutlook()) {
      const data = await callAgent("/state", ARCHIVE_CONFIG[year].fileName);
      return data.exists === true && data.mounted === true;
    }
    const status = await getWebArchiveStatus(year);
    return status.fileExists === true && status.mounted === true;
  } catch (error) {
    console.warn(`État ${year} non détectable.`, error);
    return null;
  }
}

function isDesktopOutlook() {
  return Office.context.platform !== Office.PlatformType.OfficeOnline;
}

async function runArchiveAction(year, mounted, event) {
  const archive = ARCHIVE_CONFIG[year];
  try {
    if (isDesktopOutlook()) {
      await callAgent(mounted ? "/mount-pst" : "/unmount-pst", archive.fileName);
    } else if (mounted) {
      const status = await getWebArchiveStatus(year);
      if (status.fileExists !== true) {
        throw new Error(`Le fichier PST ${year} n'est pas présent dans le partage Azure.`);
      }
      await updateWebArchiveStatus(year, true);
    } else {
      await updateWebArchiveStatus(year, false);
    }

    await setArchiveMounted(year, mounted);
    notify(mounted
      ? `${year} Montée Verte - PST monté et searchable.`
      : `${year} démontée - état Rouge.`);
  } catch (error) {
    console.error(`Action archive ${year} échouée.`, error);
    notify(`Archive ${year} : ${error.message}`);
  } finally {
    if (event) event.completed();
  }
}

Office.onReady(async () => {
  const states = readStates();
  YEARS.forEach((year) => { if (typeof states[year] !== "boolean") states[year] = false; });
  await saveStates(states);
  await setRibbonState(states);

  YEARS.forEach((year) => {
    Office.actions.associate(`mount${year}`, (event) => runArchiveAction(year, true, event));
    Office.actions.associate(`unmount${year}`, (event) => runArchiveAction(year, false, event));
  });

  // Synchronise les boutons avec la présence réelle du PST.
  for (const year of YEARS) {
    const real = await detectState(year);
    if (real !== null) states[year] = real;
  }
  await saveStates(states);
  await setRibbonState(states);
});
