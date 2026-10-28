Office.onReady(() => {});

const TOGGLE_YEARS = ["2024"];
const storageKey = (year) => `archive${year}`;

// ===== Agent local PST (Windows, Outlook Desktop) =====
const PST_AGENT = "http://127.0.0.1:45731";
const PST_FILES = { "2024": "2024.pst" }; // nom du fichier PST depose dans OneDrive (synchronise localement)

function readState(year) {
  try { return localStorage.getItem(storageKey(year)); } catch (e) { return null; }
}

function writeState(year, value) {
  try {
    if (value) localStorage.setItem(storageKey(year), value);
    else localStorage.removeItem(storageKey(year));
  } catch (e) { console.error(e); }
}

// Rouge/vert : monte => bouton "unmount" (vert) actif, "mount" (rouge) inactif ; et inversement.
async function setMounted(year, mounted) {
  if (!TOGGLE_YEARS.includes(year) || !Office.ribbon || !Office.ribbon.requestUpdate) return;
  try {
    await Office.ribbon.requestUpdate({
      tabs: [{
        id: "TabDefault",
        groups: [{
          id: "MesArchivesGroup",
          controls: [
            { id: `archive${year}_mount`, enabled: !mounted },
            { id: `archive${year}_unmount`, enabled: mounted }
          ]
        }]
      }]
    });
  } catch (e) {
    console.error("Ribbon update failed:", e);
  }
}

function notifyInfo(key, message) {
  const trimmed = (message || "").substring(0, 150);
  const item = Office && Office.context && Office.context.mailbox && Office.context.mailbox.item;
  if (!item || !item.notificationMessages) { console.warn(trimmed); return; }
  item.notificationMessages.replaceAsync(key, {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: trimmed,
    icon: "IconRed.80",
    persistent: false
  }, (res) => {
    if (res.status !== Office.AsyncResultStatus.Succeeded) console.warn("Notification failed:", res.error);
  });
}

async function callAgent(route, payload) {
  const controller = new AbortController();
  const t = setTimeout(() => controller.abort(), 20000);
  try {
    const resp = await fetch(`${PST_AGENT}${route}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload || {}),
      signal: controller.signal
    });
    let data = {};
    try { data = await resp.json(); } catch (_) {}
    return { ok: resp.ok, status: resp.status, data };
  } finally {
    clearTimeout(t);
  }
}

async function mountPst(year, event) {
  try {
    const fileName = PST_FILES[year];
    const res = await callAgent("/mount-pst", { fileName });
    if (res.ok && res.data && res.data.ok) {
      writeState(year, "mounted");
      await setMounted(year, true); // rouge -> vert
    } else if (res.status === 404 || (res.data && res.data.code === "PST_NOT_FOUND")) {
      notifyInfo("pstMount", `PST introuvable localement : ${fileName}. Verifiez la synchronisation OneDrive.`);
    } else {
      notifyInfo("pstMount", `Montage PST echoue : ${(res.data && res.data.error) || "HTTP " + res.status}`);
    }
  } catch (e) {
    notifyInfo("pstMount", "Agent PST local injoignable (127.0.0.1:45731).");
  }
  event.completed();
}

async function unmountPst(year, event) {
  try {
    const fileName = PST_FILES[year];
    const res = await callAgent("/unmount-pst", { fileName });
    if (res.ok && res.data && res.data.ok) {
      writeState(year, null);
      await setMounted(year, false); // vert -> rouge
    } else {
      notifyInfo("pstMount", `Demontage PST echoue : ${(res.data && res.data.error) || "HTTP " + res.status}`);
    }
  } catch (e) {
    notifyInfo("pstMount", "Agent PST local injoignable (127.0.0.1:45731).");
  }
  event.completed();
}

function openModop(event) {
  window.open("https://dellasiegaexternal2.github.io/MesArchivesAddin/ModOp_Mes_ARCHIVES.pdf", "_blank");
  event.completed({ allowEvent: true });
}

// 2021/2022/2023 : indisponibles
["2021", "2022", "2023"].forEach((year) => {
  Office.actions.associate(`mount${year}`, (event) => {
    notifyInfo("pstMount", `${year} indisponible.`);
    event.completed();
  });
  Office.actions.associate(`unmount${year}`, (event) => event.completed());
});

// 2024 => agent PST local
Office.actions.associate("mount2024", (event) => mountPst("2024", event));
Office.actions.associate("unmount2024", (event) => unmountPst("2024", event));
Office.actions.associate("openModop", openModop);
