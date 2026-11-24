const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");

function createDialog(fetch, config = {}) {
  const elements = new Map();
  function element() {
    return {
      children: [],
      listeners: {},
      textContent: "",
      append(...children) { this.children.push(...children); },
      appendChild(child) { this.children.push(child); },
      addEventListener(name, callback) { this.listeners[name] = callback; },
      click() { if (this.listeners.click) this.listeners.click(); },
      getAttribute(name) { return this[name]; },
      removeAttribute(name) { delete this[name]; },
      setAttribute(name, value) { this[name] = value; },
    };
  }
  const document = {
    listeners: {},
    visibilityState: "visible",
    addEventListener(name, callback) { this.listeners[name] = callback; },
    getElementById(id) {
      if (!elements.has(id)) elements.set(id, element());
      return elements.get(id);
    },
    createElement: element,
  };
  let onReady;
  const messages = [];
  const Office = {
    context: {
      platform: "PC",
      ui: { messageParent(message) { messages.push(JSON.parse(message)); } },
    },
    PlatformType: { OfficeOnline: "web" },
    onReady(callback) { onReady = callback; },
  };
  const html = fs.readFileSync(path.join(__dirname, "archives-dialog.html"), "utf8");
  const script = html.match(/<script>\s*([\s\S]*?)<\/script>/)[1];
  const window = { MES_ARCHIVES_CONFIG: config, listeners: {}, open() {}, addEventListener(name, callback) { this.listeners[name] = callback; } };
  const context = vm.createContext({
    document, Office, fetch, AbortSignal, URL, console, setTimeout, clearTimeout,
    window,
  });
  vm.runInContext(script, context);
  return {
    start: () => onReady(),
    toggle: (year) => vm.runInContext(`toggle("${year}")`, context),
    state: (year) => vm.runInContext(`states["${year}"]`, context),
    message: () => elements.get("msg").textContent,
    viewer: () => elements.get("agent-view"),
    viewerPanel: () => elements.get("viewer-panel"),
    viewerToggle: () => elements.get("viewer-toggle"),
    viewerFrame: () => elements.get("viewer-frame"),
    progressLabel: () => elements.get("progress-label").textContent,
    progressBar: () => elements.get("progress-bar"),
    close: () => elements.get("close").click(),
    messages,
    document,
    window,
  };
}

test("actualise au retour sur la fenêtre sans actualiser lorsqu'elle est masquée", async () => {
  let count = 0;
  const dialog = createDialog(async () => {
    count++;
    return response({ ok: true, exists: true, mounted: false });
  });
  await dialog.start();
  assert.equal(count, 4);
  dialog.document.visibilityState = "hidden";
  dialog.document.listeners.visibilitychange();
  assert.equal(count, 4);
  dialog.document.visibilityState = "visible";
  dialog.document.listeners.visibilitychange();
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(count, 8);
  dialog.window.listeners.focus();
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(count, 12);
});

test("propose le visualiseur intégré et un onglet séparé", () => {
  const rendered = createDialog(async () => {}, { agentViewUrl: "http://127.0.0.1:45731/view" });
  assert.equal(rendered.viewer().href, "http://127.0.0.1:45731/view");
  const html = fs.readFileSync(path.join(__dirname, "archives-dialog.html"), "utf8");
  assert.match(html, /id="agent-view"[^>]*target="_blank"/);
  assert.equal(rendered.viewerPanel().hidden, true);
  rendered.viewerToggle().click();
  assert.equal(rendered.viewerPanel().hidden, false);
  assert.equal(rendered.viewerToggle().getAttribute("aria-expanded"), "true");
  assert.equal(rendered.viewerFrame().src, "http://127.0.0.1:45731/view");
  rendered.viewerToggle().click();
  assert.equal(rendered.viewerPanel().hidden, true);
});

test("ferme la boîte de dialogue même lorsque le visualiseur est ouvert", () => {
  const dialog = createDialog(async () => {});
  dialog.viewerToggle().click();
  dialog.close();
  assert.equal(dialog.messages[0].type, "close");
});

test("masque le visualiseur uniquement à la demande de sa propre fenêtre locale", () => {
  const dialog = createDialog(async () => {});
  const frameWindow = {};
  dialog.viewerFrame().contentWindow = frameWindow;
  dialog.viewerToggle().click();
  dialog.window.listeners.message({ origin: "https://example.com", source: frameWindow, data: { type: "close-local-viewer" } });
  assert.equal(dialog.viewerPanel().hidden, false);
  dialog.window.listeners.message({ origin: "http://127.0.0.1:45731", source: {}, data: { type: "close-local-viewer" } });
  assert.equal(dialog.viewerPanel().hidden, false);
  dialog.window.listeners.message({ origin: "http://127.0.0.1:45731", source: frameWindow, data: { type: "close-local-viewer" } });
  assert.equal(dialog.viewerPanel().hidden, true);
});

test("affiche patienter après une seconde et termine la progression par étapes", async () => {
  let release;
  const pending = new Promise((resolve) => { release = resolve; });
  let count = 0;
  const dialog = createDialog(async () => {
    if (++count === 1) await pending;
    return response({ ok: true, exists: true, mounted: false });
  });
  const refresh = dialog.start();
  await new Promise((resolve) => setTimeout(resolve, 1100));
  assert.match(dialog.progressLabel(), /Veuillez patienter/);
  assert.equal(dialog.progressBar().value, 0);
  release();
  await refresh;
  assert.equal(dialog.progressLabel(), "Opération terminée.");
  assert.equal(dialog.progressBar().value, 4);
});

test("utilise l’agent local par défaut pour les deux modes d’affichage", () => {
  const dialog = createDialog(async () => {});
  assert.equal(dialog.viewer().href, "http://127.0.0.1:45731/view");
  assert.equal(dialog.viewerToggle().disabled, undefined);
});

function response(data, status = 200) {
  return { ok: status < 400, status, json: async () => data };
}

test("lit l Etat monte lorsque la boite de dialogue s ouvre", async () => {
  const dialog = createDialog(async (_url, options) => {
    const { fileName } = JSON.parse(options.body);
    return response({ ok: true, exists: true, mounted: fileName === "ADR.pst" });
  });
  await dialog.start();
  assert.equal(dialog.state("2024"), true);
  assert.equal(dialog.state("2021"), false);
});

test("confirme le montage avec l agent avant de signaler le succes", async () => {
  let mounted = false;
  const routes = [];
  const dialog = createDialog(async (url) => {
    const route = new URL(url).pathname;
    routes.push(route);
    if (route === "/mount-pst") {
      mounted = true;
      return response({ ok: true });
    }
    return response({ ok: true, exists: true, mounted });
  });
  await dialog.start();
  routes.length = 0;
  await dialog.toggle("2024");
  assert.deepEqual(routes, ["/state", "/mount-pst", "/state"]);
  assert.equal(dialog.state("2024"), true);
  assert.equal(dialog.messages[0].states["2024"], true);
});

test("confirme le demontage d une archive precedemment montee", async () => {
  let mounted = true;
  const routes = [];
  const dialog = createDialog(async (url) => {
    const route = new URL(url).pathname;
    routes.push(route);
    if (route === "/unmount-pst") {
      mounted = false;
      return response({ ok: true });
    }
    return response({ ok: true, exists: true, mounted });
  });
  await dialog.start();
  routes.length = 0;
  await dialog.toggle("2024");
  assert.deepEqual(routes, ["/state", "/unmount-pst", "/state"]);
  assert.equal(dialog.state("2024"), false);
  assert.equal(dialog.messages[0].states["2024"], false);
});

test("ne signale pas de succes lorsque Etat resultant n est pas confirme.", async () => {
  const dialog = createDialog(async (url) => {
    if (new URL(url).pathname === "/mount-pst") return response({ ok: true });
    return response({ ok: true, exists: true, mounted: false });
  });
  await dialog.start();
  await dialog.toggle("2024");
  assert.equal(dialog.state("2024"), undefined);
  assert.match(dialog.message(), /ne confirme pas/);
  assert.equal(dialog.messages.length, 0);
});

test("Etat inconnu et signale les erreurs si le montage ne peut etre confirme.", async () => {
  const dialog = createDialog(async (url) => {
    if (new URL(url).pathname === "/mount-pst") throw new TypeError("Erreur reseau");
    return response({ ok: true, exists: true, mounted: false });
  });
  await dialog.start();
  await dialog.toggle("2024");
  assert.equal(dialog.state("2024"), undefined);
  assert.match(dialog.message(), /injoignable.*Actualiser/);
  assert.equal(dialog.messages.length, 0);
});

test("n effectue aucune action si Etat a change en dehors du dialogue", async () => {
  let mounted = false;
  const routes = [];
  const dialog = createDialog(async (url) => {
    routes.push(new URL(url).pathname);
    return response({ ok: true, exists: true, mounted });
  });
  await dialog.start();
  mounted = true;
  routes.length = 0;
  await dialog.toggle("2024");
  assert.deepEqual(routes, ["/state"]);
  assert.match(dialog.message(), /état a changé/);
});

test("signale un agent indisponible au lieu de marquer les archives comme non montees", async () => {
  const dialog = createDialog(async () => { throw new Error("Indisponible"); });
  await dialog.start();
  assert.equal(dialog.state("2024"), undefined);
  assert.match(dialog.message(), /État indisponible/);
});
