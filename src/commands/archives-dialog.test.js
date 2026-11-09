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

test("reads mounted state when the dialog opens", async () => {
  const dialog = createDialog(async (_url, options) => {
    const { fileName } = JSON.parse(options.body);
    return response({ ok: true, exists: true, mounted: fileName === "ADR.pst" });
  });
  await dialog.start();
  assert.equal(dialog.state("2024"), true);
  assert.equal(dialog.state("2021"), false);
});

test("confirms mounting with the agent before reporting success", async () => {
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

test("confirms unmounting a previously mounted archive", async () => {
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

test("does not report success when the resulting state is not confirmed", async () => {
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

test("keeps state unknown and reports errors if mounting cannot be confirmed", async () => {
  const dialog = createDialog(async (url) => {
    if (new URL(url).pathname === "/mount-pst") throw new TypeError("network error");
    return response({ ok: true, exists: true, mounted: false });
  });
  await dialog.start();
  await dialog.toggle("2024");
  assert.equal(dialog.state("2024"), undefined);
  assert.match(dialog.message(), /injoignable.*Actualiser/);
  assert.equal(dialog.messages.length, 0);
});

test("does not perform an action if the state changed outside the dialog", async () => {
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

test("reports an unavailable agent instead of marking archives unmounted", async () => {
  const dialog = createDialog(async () => { throw new Error("unavailable"); });
  await dialog.start();
  assert.equal(dialog.state("2024"), undefined);
  assert.match(dialog.message(), /État indisponible/);
});
