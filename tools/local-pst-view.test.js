const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");

function createViewer(fetch, standalone = false) {
  const elements = new Map();
  function element() {
    let text = "";
    const classes = new Set();
    return {
      children: [], listeners: {}, style: {}, hidden: false,
      classList: { add: (value) => classes.add(value), remove: (value) => classes.delete(value), contains: (value) => classes.has(value) },
      set textContent(value) { text = value; this.children = []; },
      get textContent() { return text; },
      appendChild(child) { this.children.push(child); },
      addEventListener(name, callback) { this.listeners[name] = callback; },
      setAttribute(name, value) { this[name] = value; },
      removeAttribute(name) { delete this[name]; },
      querySelectorAll(selector) {
        return selector === ".message-button" ? this.children.map((item) => item.children[0]) : [];
      },
    };
  }
  const document = {
    querySelector: () => element(),
    getElementById(id) {
      if (!elements.has(id)) elements.set(id, element());
      return elements.get(id);
    },
    createElement: element,
    createTextNode: (text) => ({ textContent: text }),
  };
  document.getElementById("archive").value = "ADR_2021.pst";
  document.getElementById("archive").options = [];
  const html = fs.readFileSync(path.join(__dirname, "local-pst-view.html"), "utf8");
  const script = html.match(/<script>\s*([\s\S]*?)<\/script>/)[1];
  const window = { confirm: () => true, listeners: {}, addEventListener(name, callback) { this.listeners[name] = callback; } };
  window.parent = window;
  if (standalone) window.MES_ARCHIVES_STANDALONE = { token: "test-token" };
  const context = vm.createContext({ document, fetch, URLSearchParams, location: { search: "" }, setTimeout, clearTimeout, URL, window });
  vm.runInContext(script, context);
  return { elements, window, run: (code) => vm.runInContext(code, context) };
}

function jsonResponse(data, status = 200) {
  return { ok: status < 400, status, headers: { get: () => "application/json" }, json: async () => data };
}

test("lance la recherche avec le bouton ou Entrée sans soumettre de formulaire", async () => {
  let requests = 0;
  const viewer = createViewer(async () => {
    requests++;
    return jsonResponse({ ok: true, mounted: false, outlookRunning: true });
  });
  viewer.elements.get("archive").options = [{ value: "first.pst", textContent: "2021" }];
  viewer.elements.get("search-query").value = "Alpha";
  viewer.elements.get("search-submit").listeners.click();
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(requests, 1);
  let prevented = false;
  viewer.elements.get("search-query").listeners.keydown({ key: "Enter", preventDefault() { prevented = true; } });
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(prevented, true);
  assert.equal(requests, 2);
});

test("le relais autonome accepte seulement son parent de confiance et les routes de contrôle", async () => {
  const calls = [];
  const sent = [];
  const viewer = createViewer(async (route, options) => {
    calls.push({ route, headers: options.headers, body: JSON.parse(options.body) });
    return jsonResponse({ ok: true, mode: "standalone", available: true, loaded: false });
  }, true);
  const parent = { postMessage: (message, origin) => sent.push({ message, origin }) };
  viewer.window.parent = parent;
  const handler = viewer.window.listeners.message;
  const data = { type: "standalone-request", id: "1", route: "/state", fileName: "ADR_2021.pst" };
  await handler({ origin: "https://example.com", source: parent, data });
  await handler({ origin: "https://dellasiegaps.github.io", source: {}, data });
  await handler({ origin: "https://dellasiegaps.github.io", source: parent, data: { ...data, route: "/message" } });
  assert.equal(calls.length, 0);
  await handler({ origin: "https://dellasiegaps.github.io", source: parent, data });
  assert.equal(calls.length, 1);
  assert.equal(calls[0].headers["X-MesArchives-Token"], "test-token");
  assert.equal(sent[0].message.id, "1");
  assert.equal(sent[0].origin, "https://dellasiegaps.github.io");
});

test("affiche les dossiers, les messages et le détail ensemble et conserve la pagination", async () => {
  const calls = [];
  const viewer = createViewer(async (route, options) => {
    const body = JSON.parse(options.body);
    calls.push({ route, ...body });
    if (route === "/message") return jsonResponse({
      ok: true, subject: "Bonjour", sender: "Alice", to: "Bob", cc: "", date: "", body: "Contenu",
      senderContact: { name: "Alice", address: "alice@example.test" },
      recipients: [{ type: 1, contact: { name: "Bob", address: "bob@example.test" } }],
      attachments: [{ name: "test.txt", size: 12, index: 1 }],
    });
    return jsonResponse({
      ok: true, folder: { id: "root", name: "Archive" },
      folders: [{ id: "child", name: "Réception", count: 51 }],
      messages: [{ id: "m1", subject: "Bonjour", sender: "Alice", date: "" }],
      offset: body.offset, nextOffset: body.offset === 0 ? 50 : null, totalItems: 51,
    });
  });
  await viewer.run('showFolder("", 0, [])');
  const elements = viewer.elements;
  assert.equal(elements.get("folders").children.length, 1);
  const button = elements.get("messages").children[0].children[0];
  await viewer.run("showMessage('m1', document.getElementById('messages').children[0].children[0])");
  assert.equal(elements.get("message-detail").hidden, false);
  assert.equal(elements.get("reading-placeholder").hidden, true);
  assert.equal(elements.get("messages").children.length, 1);
  assert.equal(elements.get("folders").children.length, 1);
  assert.equal(button.classList.contains("selected"), true);
  assert.equal(elements.get("body").textContent, "Contenu");
  assert.equal(elements.get("from").children[0].href, "mailto:alice%40example.test");
  assert.equal(elements.get("to").children[0].href, "mailto:bob%40example.test");
  assert.equal(elements.get("attachment-list").children.length, 1);
  elements.get("next").listeners.click();
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(calls.at(-1).offset, 50);
  assert.equal(elements.get("message-detail").hidden, true);
  assert.equal(elements.get("reading-placeholder").hidden, false);
});

test("affiche une erreur sans conserver le nom du dossier précédent", async () => {
  let fail = false;
  const viewer = createViewer(async () => fail
    ? jsonResponse({ error: "PST non monté" }, 500)
    : jsonResponse({ ok: true, folder: { id: "root", name: "Ancienne archive" }, folders: [], messages: [], offset: 0, nextOffset: null, totalItems: 0 }));
  await viewer.run('showFolder("", 0, [])');
  fail = true;
  await viewer.run('showFolder("", 0, [])');
  assert.match(viewer.elements.get("status").textContent, /PST non monté/);
  assert.equal(viewer.elements.get("folder-name").textContent, "Messages");
  assert.equal(viewer.elements.get("message-detail").hidden, true);
});

test("ne transforme pas une adresse non valide en lien", () => {
  const viewer = createViewer(async () => {});
  viewer.run(`renderContacts("from", "De", [{ name: "Inconnu", address: "bad@example.test?subject=inject" }], "")`);
  assert.match(viewer.elements.get("from").children[0].textContent, /adresse indisponible/);
  assert.equal(viewer.elements.get("from").children[0].href, undefined);
});

test("recherche le corps et les métadonnées dans toutes les archives montées et leurs sous-dossiers", async () => {
  const calls = [];
  const viewer = createViewer(async (route, options) => {
    const body = JSON.parse(options.body);
    calls.push({ route, ...body });
    if (route === "/state") return jsonResponse({ ok: true, mounted: body.fileName !== "third.pst", outlookRunning: true });
    if (route === "/message") return jsonResponse({ ok: true, body: "Contenu RECHERCHÉ" });
    const child = body.folderId === "child";
    return jsonResponse({
      ok: true, folder: { id: child ? "child" : "root", name: child ? "Enfant" : "Racine" },
      folders: !child && body.fileName === "first.pst" ? [{ id: "child", name: "Enfant", count: 1 }] : [],
      messages: child ? [{ id: "mail", subject: "Sans rapport", sender: "Alice", date: "" }]
        : body.fileName === "second.pst" ? [{ id: "mail2", subject: "Recherché", sender: "Bob", date: "" }] : [],
      offset: 0, nextOffset: null, totalItems: 1,
    });
  });
  viewer.elements.get("archive").options = [
    { value: "first.pst", textContent: "2021" },
    { value: "second.pst", textContent: "2022" },
    { value: "third.pst", textContent: "2023" },
  ];
  await viewer.run('searchArchives("recherché")');
  assert.equal(viewer.elements.get("messages").children.length, 2);
  assert.match(viewer.elements.get("status").textContent, /2 résultat\(s\) dans 2 archive/);
  assert.equal(calls.some((call) => call.route === "/message" && call.fileName === "first.pst"), true);
  assert.equal(calls.some((call) => call.route === "/folder" && call.fileName === "third.pst"), false);
});

test("signale une recherche incomplète lorsque l'agent échoue", async () => {
  const viewer = createViewer(async () => jsonResponse({ error: "Agent indisponible" }, 500));
  viewer.elements.get("archive").options = [{ value: "first.pst", textContent: "2021" }];
  await viewer.run('searchArchives("test")');
  assert.match(viewer.elements.get("status").textContent, /incomplets.*Agent indisponible/);
});

test("recherche dans les destinataires, adresses, dates et noms de pièces jointes sans sensibilité aux accents", async () => {
  const viewer = createViewer(async (route) => {
    if (route === "/state") return jsonResponse({ ok: true, mounted: true, outlookRunning: true });
    if (route === "/folder") return jsonResponse({
      ok: true, folder: { id: "root", name: "Racine" }, folders: [],
      messages: [{ id: "mail", subject: "Objet", sender: "Alice", date: "2026-01-02T03:04:05" }],
      offset: 0, nextOffset: null, totalItems: 1
    });
    return jsonResponse({
      ok: true, subject: "Objet", sender: "Alice", to: "Destinataire", cc: "Copie", bcc: "Secret",
      date: "2026-01-02T03:04:05", body: "Contenu", senderContact: { name: "Alice", address: "alice@example.test" },
      recipients: [{ type: 1, contact: { name: "Destinataire", address: "bob@example.test" } }],
      attachments: [{ name: "RÉSUMÉ.pdf" }]
    });
  });
  viewer.elements.get("archive").options = [{ value: "first.pst", textContent: "2021" }];
  for (const query of ["destinataire", "copie", "secret", "alice@example.test", "bob@example.test", "resume.pdf", "02/01/2026", "contenu"]) {
    await viewer.run(`searchArchives(${JSON.stringify(query)})`);
    assert.equal(viewer.elements.get("messages").children.length, 1, query);
  }
});

test("explique les erreurs non JSON au lieu d'afficher une erreur de parsing", async () => {
  const viewer = createViewer(async () => ({
    ok: false, status: 403, headers: { get: () => "text/plain" }
  }));
  await viewer.run('showFolder("", 0, [])');
  assert.match(viewer.elements.get("status").textContent, /HTTP 403.*lecteur local/);
});

test("arrête la recherche après la requête en cours", async () => {
  let release;
  const pending = new Promise((resolve) => { release = resolve; });
  let calls = 0;
  const viewer = createViewer(async () => { calls++; await pending; return jsonResponse({ ok: true, mounted: true, outlookRunning: true }); });
  viewer.elements.get("archive").options = [{ value: "first.pst", textContent: "2021" }];
  const search = viewer.run('searchArchives("test")');
  viewer.elements.get("search-cancel").listeners.click();
  release();
  await search;
  assert.equal(calls, 1);
  assert.match(viewer.elements.get("status").textContent, /arrêtée.*incomplets/);
});

test("demande au parent de masquer le visualiseur intégré", () => {
  const viewer = createViewer(async () => {});
  viewer.run(`window.parent = { postMessage(message, origin) { window.sent = { message, origin }; } }`);
  viewer.elements.get("close-viewer").listeners.click();
  assert.equal(viewer.run("window.sent.message.type"), "close-local-viewer");
  assert.equal(viewer.run("window.sent.origin"), "https://dellasiegaps.github.io");
});

test("prépare un brouillon local avec réponse à tous, Cc et Cci éditables", () => {
  const viewer = createViewer(async () => {});
  viewer.run(`selectedMessage = {
    subject: "Sujet", sender: "Alice", to: "Bob", cc: "Claire", date: "", body: "Contenu",
    senderContact: { address: "alice@example.test" },
    recipients: [
      { type: 1, contact: { address: "bob@example.test" } },
      { type: 2, contact: { address: "claire@example.test" } },
      { type: 2, contact: { address: "" } }
    ]
  }; prepareDraft("reply-all")`);
  assert.equal(viewer.elements.get("compose-to").value, "alice@example.test; bob@example.test");
  assert.equal(viewer.elements.get("compose-cc").value, "claire@example.test");
  assert.equal(viewer.elements.get("compose-bcc").value, "");
  assert.match(viewer.elements.get("compose-body").value, /Contenu/);
  assert.match(viewer.elements.get("compose-warning").textContent, /indisponibles/);
  viewer.run('prepareDraft("forward")');
  assert.equal(viewer.elements.get("compose-to").value, "");
  assert.equal(viewer.elements.get("compose-subject").value, "TR: Sujet");
});

test("préremplit le lien Outlook Web sans envoi et bloque les liens trop longs", () => {
  const viewer = createViewer(async () => {});
  const elements = viewer.elements;
  for (const [id, value] of [
    ["compose-to", "alice@example.test"], ["compose-cc", "bob@example.test"],
    ["compose-bcc", "carol@example.test"], ["compose-subject", "Sujet & réponse"],
    ["compose-body", "Texte privé"]
  ]) {
    viewer.run(`document.getElementById(${JSON.stringify(id)})`);
    elements.get(id).value = value;
  }
  let blocked = false;
  const link = elements.get("open-compose");
  link.listeners.click({ currentTarget: link, preventDefault() { blocked = true; } });
  const url = new URL(link.href);
  assert.equal(url.origin, "https://outlook.office.com");
  assert.equal(url.searchParams.get("bcc"), "carol@example.test");
  assert.equal(url.searchParams.get("subject"), "Sujet & réponse");
  assert.equal(url.searchParams.get("body"), "Texte privé");
  assert.equal(blocked, false);
  elements.get("compose-body").value = "x".repeat(3000);
  link.listeners.click({ currentTarget: link, preventDefault() { blocked = true; } });
  assert.equal(blocked, true);
  assert.match(elements.get("compose-warning").textContent, /trop long/);
});
