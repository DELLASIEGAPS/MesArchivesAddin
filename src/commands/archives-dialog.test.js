const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");

function createDialog(fetch) {
  const elements = new Map();
  function element() {
    return {
      children: [],
      textContent: "",
      append(...children) { this.children.push(...children); },
      appendChild(child) { this.children.push(child); },
      addEventListener() {},
    };
  }
  const document = {
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
  const context = vm.createContext({
    document, Office, fetch, AbortSignal, URL, console,
    window: { MES_ARCHIVES_CONFIG: {}, open() {} },
  });
  vm.runInContext(script, context);
  return {
    start: () => onReady(),
    toggle: (year) => vm.runInContext(`toggle("${year}")`, context),
    state: (year) => vm.runInContext(`states["${year}"]`, context),
    message: () => elements.get("msg").textContent,
    messages,
  };
}

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
