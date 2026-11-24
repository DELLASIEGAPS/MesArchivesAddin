const assert = require("node:assert/strict");
const { spawn } = require("node:child_process");
const fs = require("node:fs/promises");
const net = require("node:net");
const os = require("node:os");
const path = require("node:path");
const test = require("node:test");

const script = path.join(__dirname, "local-pst-agent.ps1");

async function unusedPort() {
  const server = net.createServer();
  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(0, "127.0.0.1", resolve);
  });
  const port = server.address().port;
  await new Promise((resolve, reject) => server.close((error) => error ? reject(error) : resolve()));
  return port;
}

async function startAgent(directory, env) {
  const port = await unusedPort();
  const base = `http://127.0.0.1:${port}`;
  const child = spawn("powershell.exe", [
    "-NoProfile", "-NonInteractive", "-File", script,
    "-ListenPrefix", `${base}/`, "-PstSourceDirectory", directory,
  ], { env: { ...process.env, ...env }, stdio: ["ignore", "pipe", "pipe"] });

  let output = "";
  child.stdout.on("data", (chunk) => { output += chunk; });
  child.stderr.on("data", (chunk) => { output += chunk; });
  try {
    const deadline = Date.now() + 15000;
    while (Date.now() < deadline) {
      if (child.exitCode !== null) throw new Error(`Agent exited: ${output}`);
      try {
        const response = await fetch(`${base}/view`, { signal: AbortSignal.timeout(500) });
        if (response.ok) return { child, base };
      } catch {
        await new Promise((resolve) => setTimeout(resolve, 100));
      }
    }
    throw new Error(`Agent did not start: ${output}`);
  } catch (error) {
    child.kill();
    throw error;
  }
}

async function state(base, fileName) {
  const response = await fetch(`${base}/state`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ fileName }),
  });
  assert.equal(response.status, 200);
  return response.json();
}

test("dossier explicite : quatre PST locaux sans presence cloud inventee", async (t) => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), "mesarchives-agent-test-"));
  let child;
  t.after(async () => {
    if (child && child.exitCode === null) {
      await new Promise((resolve) => { child.once("exit", resolve); child.kill(); });
    }
    await fs.rm(root, { recursive: true, force: true });
  });
  const source = path.join(root, "source");
  const oneDrive = path.join(root, "onedrive");
  const localAppData = path.join(root, "local");
  await Promise.all([source, oneDrive, localAppData].map((dir) => fs.mkdir(dir)));
  const files = ["ADR_2021.pst", "ADR_2022.pst", "ADR_2023.pst", "ADR.pst"];
  for (const file of files) await fs.writeFile(path.join(source, file), "test");
  await fs.writeFile(path.join(oneDrive, "autre.pst"), "test");

  const agent = await startAgent(source, {
    LOCALAPPDATA: localAppData, ONEDRIVECOMMERCIAL: oneDrive, ONEDRIVE: "", ONEDRIVECONSUMER: "",
  });
  child = agent.child;
  const { base } = agent;
  for (const file of files) {
    const data = await state(base, file);
    assert.equal(data.exists, true);
    assert.equal(data.local, true);
    assert.equal(data.cloud, null);
    assert.equal(data.workingCopy, false);
  }
  assert.equal((await state(base, "autre.pst")).exists, false);
  const page = await (await fetch(`${base}/view`)).text();
  assert.match(page, /Outlook classique/);
  assert.match(page, /Lecture seule/);
  assert.match(page, /Pièces jointes/);
  const denied = await fetch(`${base}/folder`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Origin: "https://example.com" },
    body: JSON.stringify({ fileName: "ADR.pst" }),
  });
  assert.equal(denied.status, 403);
  assert.match((await denied.json()).error, /refusé/);
});

test("dossier OneDrive explicite : origine cloud et presence locale", async (t) => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), "mesarchives-onedrive-test-"));
  let child;
  t.after(async () => {
    if (child && child.exitCode === null) {
      await new Promise((resolve) => { child.once("exit", resolve); child.kill(); });
    }
    await fs.rm(root, { recursive: true, force: true });
  });
  const source = path.join(root, "onedrive", "Personnel");
  const localAppData = path.join(root, "local");
  await fs.mkdir(source, { recursive: true });
  await fs.mkdir(localAppData);
  await fs.writeFile(path.join(source, "ADR.pst"), "test");
  const agent = await startAgent(source, {
    LOCALAPPDATA: localAppData, ONEDRIVECOMMERCIAL: path.join(root, "onedrive"),
    ONEDRIVE: "", ONEDRIVECONSUMER: "",
  });
  child = agent.child;
  const { base } = agent;
  const data = await state(base, "ADR.pst");
  assert.equal(data.exists, true);
  assert.equal(data.local, true);
  assert.equal(data.cloud, true);
  assert.equal(data.onedriveLocal, true);
});
