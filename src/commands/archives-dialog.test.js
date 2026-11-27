const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");
const ARCHIVE_YEARS = ["2021", "2022", "2023", "2024"];

function createDialog(fetch, config = {}, platform = "PC", mode = "classic") {
  const elements = new Map();
  function element() {
    let textContent = "";
    return {
      children: [],
      listeners: {},
      get textContent() { return textContent; },
      set textContent(value) {
        textContent = value;
        if (value === "") this.children = [];
      },
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
      platform,
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
  if (mode === "classic") {
    elements.get("viewer-mode").value = "classic";
    elements.get("viewer-mode").listeners.change();
  }
  return {
    start: () => onReady(),
    toggle: (year) => vm.runInContext(`toggle("${year}")`, context),
    state: (year) => vm.runInContext(`states["${year}"]`, context),
    message: () => elements.get("msg").textContent,
    viewer: () => elements.get("agent-view"),
    viewerPanel: () => elements.get("viewer-panel"),
    viewerToggle: () => elements.get("viewer-toggle"),
    viewerFrame: () => elements.get("viewer-frame"),
    viewerMode: () => elements.get("viewer-mode"),
    viewerModeHint: () => elements.get("viewer-mode-hint").textContent,
    progressLabel: () => elements.get("progress-label").textContent,
    progressBar: () => elements.get("progress-bar"),
    buttonText: (year) => elements.get("list").children[ARCHIVE_YEARS.indexOf(year)]?.children[4]?.textContent,
    detailText: (year) => elements.get("list").children[ARCHIVE_YEARS.indexOf(year)]?.children[3]?.textContent,
    modeHint: () => elements.get("mode-hint").textContent,
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

test("choisit explicitement le lecteur autonome et ses actions indépendantes", () => {
  const dialog = createDialog(async () => {});
  dialog.viewerToggle().click();
  dialog.viewerMode().value = "standalone";
  dialog.viewerMode().listeners.change();
  assert.equal(dialog.viewer().href, "http://127.0.0.1:45732/view");
  assert.equal(dialog.viewerFrame().src, "http://127.0.0.1:45732/view");
  assert.match(dialog.viewerModeHint(), /sans Outlook classique/);
  assert.match(dialog.viewerModeHint(), /sans modifier le PST/);
  const frameWindow = {};
  dialog.viewerFrame().contentWindow = frameWindow;
  dialog.window.listeners.message({ origin: "http://127.0.0.1:45731", source: frameWindow, data: { type: "close-local-viewer" } });
  assert.equal(dialog.viewerPanel().hidden, false);
  dialog.window.listeners.message({ origin: "http://127.0.0.1:45732", source: frameWindow, data: { type: "close-local-viewer" } });
  assert.equal(dialog.viewerPanel().hidden, true);
  dialog.viewerMode().value = "classic";
  dialog.viewerMode().listeners.change();
  assert.equal(dialog.viewer().href, "http://127.0.0.1:45731/view");
  assert.equal(dialog.viewerFrame().getAttribute("src"), undefined);
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

test("configure le visualiseur local par défaut sans rendre le bouton dépendant d’un agent", () => {
  const dialog = createDialog(async () => {});
  assert.equal(dialog.viewer().href, "http://127.0.0.1:45731/view");
  assert.equal(dialog.viewerToggle().disabled, undefined);
});

test("sélectionne par défaut le lecteur autonome dans la version de développement", () => {
  const dialog = createDialog(async () => {}, {}, "web", "standalone");
  assert.equal(dialog.viewerMode().value, "standalone");
  assert.equal(dialog.viewer().href, "http://127.0.0.1:45732/view");
  assert.match(dialog.modeHint(), /sans montage dans Outlook/);
});

test("le mode web lit et modifie uniquement le marqueur cloud sans appeler l’agent Outlook local", async () => {
  let mounted = false;
  const calls = [];
  const dialog = createDialog(async (url, options = {}) => {
    const parsedUrl = new URL(url);
    calls.push({ url: parsedUrl, options });
    if (options.method === "POST") {
      const body = JSON.parse(options.body);
      assert.equal(body.annee, "2024");
      mounted = body.Monte;
      return response({ ok: true });
    }
    return response({ fileExists: true, mounted });
  }, { webStatusEndpoint: "https://flow.example/trigger?sig=test" }, "web");

  await dialog.start();
  assert.equal(calls.length, 4);
  assert.ok(calls.every(({ url }) => url.origin === "https://flow.example"));
  assert.equal(calls.find(({ url }) => url.searchParams.get("year") === "2024").url.searchParams.get("sig"), "test");
  assert.match(dialog.modeHint(), /le PST n'est pas monté dans Outlook Web/);
  assert.match(dialog.buttonText("2024"), /Activer côté web/);
  assert.match(dialog.detailText("2024"), /marqueur web inactif/);

  calls.length = 0;
  await dialog.toggle("2024");
  assert.deepEqual(calls.map(({ options }) => options.method || "GET"), ["GET", "POST", "GET"]);
  assert.equal(calls[1].options.headers["Content-Type"], "text/plain;charset=UTF-8");
  assert.equal(dialog.state("2024"), true);
  assert.match(dialog.detailText("2024"), /n'est pas monté dans Outlook Web/);
  assert.match(dialog.buttonText("2024"), /Retirer le marqueur web/);
});

test("un flux cloud web non configuré affiche sa propre erreur sans contacter Outlook classique", async () => {
  let calls = 0;
  const dialog = createDialog(async () => {
    calls++;
    throw new Error("L'agent local ne doit pas être appelé en mode web.");
  }, {}, "web");
  await dialog.start();
  assert.equal(calls, 0);
  assert.match(dialog.message(), /configurez le flux cloud de statut/);
  assert.equal(dialog.state("2021"), undefined);
});

test("charge et décharge via le lecteur autonome dans Outlook Web sans agent classique ni flux cloud", async () => {
  const calls = [];
  const loaded = new Set();
  const dialog = createDialog(async () => { throw new Error("Aucun appel direct attendu"); }, {}, "web", "standalone");
  const frame = {
    postMessage(message, origin) {
      assert.equal(origin, "http://127.0.0.1:45732");
      calls.push(message.route);
      if (message.route === "/load") loaded.add(message.fileName);
      if (message.route === "/unload") loaded.delete(message.fileName);
      const result = message.route === "/state"
        ? { ok: true, mode: "standalone", available: true, loaded: loaded.has(message.fileName) }
        : { ok: true, loaded: loaded.has(message.fileName) };
      dialog.window.listeners.message({
        origin: "https://example.com", source: frame,
        data: { type: "standalone-response", id: message.id, error: "Réponse falsifiée" }
      });
      queueMicrotask(() => dialog.window.listeners.message({
        origin, source: frame, data: { type: "standalone-response", id: message.id, result }
      }));
    }
  };
  dialog.viewerFrame().contentWindow = frame;
  dialog.window.listeners.message({
    origin: "http://127.0.0.1:45732", source: frame, data: { type: "standalone-ready" }
  });
  await dialog.start();
  assert.match(dialog.modeHint(), /sans montage dans Outlook/);
  calls.length = 0;
  await dialog.toggle("2024");
  assert.deepEqual(calls, ["/state", "/load", "/state"]);
  assert.equal(dialog.state("2024"), true);
  assert.equal(dialog.messages.length, 0);
  calls.length = 0;
  await dialog.toggle("2024");
  assert.deepEqual(calls, ["/state", "/unload", "/state"]);
  assert.equal(dialog.state("2024"), false);
});

test("le lecteur autonome ne signale pas un chargement si l'état final ne le confirme pas", async () => {
  const dialog = createDialog(async () => { throw new Error("Appel direct inattendu"); }, {}, "web", "standalone");
  const frame = {
    postMessage(message, origin) {
      queueMicrotask(() => dialog.window.listeners.message({
        origin, source: frame, data: {
          type: "standalone-response", id: message.id,
          result: message.route === "/state"
            ? { ok: true, mode: "standalone", available: true, loaded: false }
            : { ok: true, loaded: true }
        }
      }));
    }
  };
  dialog.viewerFrame().contentWindow = frame;
  dialog.window.listeners.message({
    origin: "http://127.0.0.1:45732", source: frame, data: { type: "standalone-ready" }
  });
  await dialog.start();
  await dialog.toggle("2024");
  assert.equal(dialog.state("2024"), undefined);
  assert.match(dialog.message(), /ne confirme pas/);
  assert.equal(dialog.messages.length, 0);
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
