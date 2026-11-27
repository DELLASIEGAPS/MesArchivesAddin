const test = require("node:test");
const assert = require("node:assert/strict");
const { spawn } = require("node:child_process");
const { createHash } = require("node:crypto");
const fs = require("node:fs/promises");
const net = require("node:net");
const path = require("node:path");
const os = require("node:os");

const fixture = process.env.MESARCHIVES_TEST_PST;
const library = process.env.MESARCHIVES_XST_DLL;

test("lecteur autonome : PST fictif, lecture, pièce jointe et refus des requêtes non autorisées", { skip: !fixture || !library }, async (t) => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), "mesarchives-autonome-"));
  const file = path.join(root, "ADR_2021.pst");
  await fs.copyFile(fixture, file);
  await fs.writeFile(path.join(root, "ADR_2022.pst"), "PST invalide");
  const before = createHash("sha256").update(await fs.readFile(file)).digest("hex");
  const server = net.createServer();
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const port = server.address().port;
  await new Promise((resolve) => server.close(resolve));
  const base = `http://127.0.0.1:${port}`;
  const child = spawn("pwsh", ["-NoProfile", "-File", path.join(__dirname, "local-pst-lecture-autonome.ps1"),
    "-PstSourceDirectory", root, "-LibraryPath", library, "-Port", String(port)]);
  let output = "";
  child.stdout.on("data", (chunk) => { output += chunk; });
  child.stderr.on("data", (chunk) => { output += chunk; });
  t.after(async () => {
    if (child.exitCode === null) {
      child.kill();
      await new Promise((resolve) => child.once("exit", resolve));
    }
    await fs.rm(root, { recursive: true, force: true });
  });
  let html;
  const deadline = Date.now() + 60000;
  while (Date.now() < deadline) {
    if (child.exitCode !== null) throw new Error(`Lecteur arrêté : ${output}`);
    try {
      const page = await fetch(`${base}/view`);
      assert.match(page.headers.get("content-security-policy"), /frame-ancestors 'self' https:\/\/dellasiegaps\.github\.io;/);
      html = await page.text();
      break;
    } catch (error) {
      if (error instanceof assert.AssertionError) throw error;
      await new Promise((r) => setTimeout(r, 100));
    }
  }
  assert.ok(html, `Lecteur non démarré : ${output}`);
  const token = html.match(/MES_ARCHIVES_STANDALONE=\{token:'([A-F0-9]+)'\}/)?.[1];
  assert.ok(token);
  const headers = { "Content-Type": "application/json", Origin: base, "X-MesArchives-Token": token };
  const post = async (route, body, customHeaders = headers) => {
    try {
      return await fetch(`${base}${route}`, {
        method: "POST", headers: customHeaders, body: JSON.stringify({ fileName: "ADR_2021.pst", ...body }),
      });
    } catch (error) {
      throw new Error(`Erreur ${route} : ${error.message}. Lecteur : ${output}`, { cause: error });
    }
  };
  assert.equal((await post("/state", {}, { ...headers, Origin: "https://example.com" })).status, 403);
  assert.equal((await post("/state", {}, { ...headers, "X-MesArchives-Token": "invalid" })).status, 403);
  assert.equal((await post("/load", {}, { ...headers, Origin: "https://example.com" })).status, 403);
  assert.equal((await post("/folder", { fileName: "..\\ADR_2021.pst", folderId: "", offset: 0 })).status, 500);
  const initial = await (await post("/state", {})).json();
  assert.equal(initial.available, true);
  assert.equal(initial.loaded, false);
  assert.equal((await post("/folder", { folderId: "", offset: 0 })).status, 500);
  const corrupt = await post("/load", { fileName: "ADR_2022.pst" });
  assert.equal(corrupt.status, 500);
  assert.equal((await corrupt.json()).ok, false);
  assert.equal((await (await post("/state", { fileName: "ADR_2022.pst" })).json()).loaded, false);
  assert.equal((await (await post("/load", {})).json()).loaded, true);
  const pending = [""];
  let found;
  while (pending.length && !found) {
    const response = await post("/folder", { folderId: pending.shift(), offset: 0 });
    assert.equal(response.status, 200, await response.clone().text());
    const folder = await response.json();
    pending.push(...folder.folders.map((entry) => entry.id));
    found = folder.messages.find((entry) => entry.subject === "Alpha");
  }
  assert.ok(found, "Message fictif Alpha absent");
  const response = await post("/message", { itemId: found.id });
  assert.equal(response.status, 200, await response.clone().text());
  const message = await response.json();
  assert.equal(message.body.trimEnd(), "This is message alpha.");
  assert.ok(message.attachments.length);
  assert.equal(message.attachments[0].size, null);
  const attachment = await post("/attachment", { itemId: found.id, index: 1 });
  assert.equal(attachment.status, 200);
  assert.deepEqual([...new Uint8Array(await attachment.arrayBuffer()).slice(0, 8)], [137, 80, 78, 71, 13, 10, 26, 10]);
  assert.equal((await post("/message", { itemId: "old-version|0:0" })).status, 500);
  for (let i = 0; i < 10; i++) {
    assert.equal((await post("/folder", { folderId: "", offset: 0 })).status, 200);
    assert.equal((await post("/message", { itemId: found.id })).status, 200);
    assert.equal((await post("/attachment", { itemId: found.id, index: 1 })).status, 200);
  }
  await new Promise((resolve) => setTimeout(resolve, 750));
  assert.equal((await post("/state", {})).status, 200, "Le lecteur doit rester actif après les lectures.");
  assert.equal((await (await post("/unload", {})).json()).loaded, false);
  assert.equal((await post("/folder", { folderId: "", offset: 0 })).status, 500);
  assert.equal((await post("/message", { itemId: found.id })).status, 500);
  assert.equal((await post("/attachment", { itemId: found.id, index: 1 })).status, 500);
  assert.equal((await (await post("/state", {})).json()).available, true);
  assert.equal(createHash("sha256").update(await fs.readFile(file)).digest("hex"), before);
});
