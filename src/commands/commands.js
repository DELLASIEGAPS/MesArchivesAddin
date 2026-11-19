const STATE_KEY = "archivesMounted";
const DIALOG_URL = "https://dellasiegaps.github.io/MesArchivesAddin/src/commands/archives-dialog.html";

let dialog = null;

function saveStates(states) {
  Office.context.roamingSettings.set(STATE_KEY, states);
  Office.context.roamingSettings.saveAsync((result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("Impossible d'enregistrer l'état des archives.", result.error);
    }
  });
}

function onDialogMessage(arg) {
  let msg;
  try { msg = JSON.parse(arg.message); } catch (_) { return; }
  if (msg.type === "states") saveStates(msg.states);
  if (msg.type === "close" && dialog) {
    dialog.close();
    dialog = null;
  }
}

// Ouvre le popup unique qui gère les 4 archives (flip-flop rouge/vert).
function openArchives(event) {
  if (dialog) {
    dialog.close();
    dialog = null;
  }
  Office.context.ui.displayDialogAsync(
    DIALOG_URL,
    { height: 50, width: 45, displayInIframe: false, promptBeforeOpen: false },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, onDialogMessage);
        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => { dialog = null; });
      } else {
        console.error("Ouverture du popup impossible.", result.error);
      }
      event.completed();
    }
  );
}

Office.onReady(() => {
  Office.actions.associate("openArchives", openArchives);
});

// Requis par FunctionFile (manifeste V1_0).
window.openArchives = openArchives;
