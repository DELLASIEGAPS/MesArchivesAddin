const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");

function createViewer(fetch) {
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
  const context = vm.createContext({ document, fetch, URLSearchParams, location: { search: "" }, setTimeout, clearTimeout, URL });
  vm.runInContext(script, context);
  return { elements, run: (code) => vm.runInContext(code, context) };
}

function jsonResponse(data, status = 200) {
  return { ok: status < 400, status, json: async () => data };
}

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
