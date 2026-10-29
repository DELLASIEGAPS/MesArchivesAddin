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

// 2024 => popup HTML (flip-flop rouge/vert) qui appelle l agent PST local
let dialog2024 = null;
function open2024(event) {
  const state = readState("2024") === "mounted" ? "mounted" : "unmounted";
  const url = `${location.origin}${location.pathname.replace(/[^/]*$/, "")}dialog2024.html?state=${state}`;
  Office.context.ui.displayDialogAsync(url, { height: 40, width: 25, displayInIframe: false }, (res) => {
    if (res.status !== Office.AsyncResultStatus.Succeeded) {
      notifyInfo("pstMount", "Popup 2024 impossible : " + res.error.message);
      event.completed();
      return;
    }
    dialog2024 = res.value;
    dialog2024.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
      writeState("2024", arg.message === "mounted" ? "mounted" : null);
    });
    dialog2024.addEventHandler(Office.EventType.DialogEventReceived, () => {
      dialog2024 = null;
    });
    event.completed();
  });
}
Office.actions.associate("open2024", open2024);
Office.actions.associate("openModop", openModop);

