const assert = require("assert/strict");
const fs = require("fs");
const vm = require("vm");

const source = fs.readFileSync("app.js", "utf8").split("async function startDashboard()")[0];
function setup(hostname = "localhost", stored = new Map()) {
  const elements = new Map();
  const alerts = [];
  const element = (selector) => {
    if (!elements.has(selector)) {
      const classes = new Set();
      elements.set(selector, {
        value: "", textContent: "", disabled: false,
        classList: {
          toggle(name, enabled) { enabled ? classes.add(name) : classes.delete(name); },
          contains(name) { return classes.has(name); },
        },
      });
    }
    return elements.get(selector);
  };
  const sandbox = {
    window: {
      INITIAL_WORKBOOK_DATA: { mapping: { existing: "Old local" }, weeks: [] },
      location: { protocol: "http:", hostname },
    },
    localStorage: { getItem: (key) => stored.get(key), setItem: (key, value) => stored.set(key, value) },
    sessionStorage: { getItem: () => null, setItem: () => {} },
    document: { querySelector: element },
    alert: (message) => alerts.push(message),
    btoa: (text) => Buffer.from(text, "binary").toString("base64"),
    atob: (text) => Buffer.from(text, "base64").toString("binary"),
    console,
  };
  vm.createContext(sandbox);
  vm.runInContext(source + `
    renderAll = () => renderAdminMode();
    globalThis.api = {
      state, addGameNameMapping, publishGameNameMapping, loadMappingCsv,
      parseMappingCsv, mappingEntriesFromCsv, mappingEntriesToCsv, renderAdminMode,
    };
  `, sandbox);
  return { sandbox, api: sandbox.api, element, stored, alerts };
}

async function main() {
  const env = setup();
  const { api, element, sandbox } = env;
  const add = (english, display) => {
    element("#mappingEnglishName").value = english;
    element("#mappingDisplayName").value = display;
    api.addGameNameMapping();
  };
  add("New Game", "new");
  assert.equal(api.state.mappingDrafts.length, 0, "Read-only mode rejects edits");
  api.state.adminMode = true;
  api.renderAdminMode();
  assert.equal(element("#adminPanel").classList.contains("is-open"), true);
  add("", "missing");
  assert.equal(api.state.mappingDrafts.length, 0);
  add("Foxy's Egg Heist", "original");
  add("Foxy\u2019s Egg Heist", "狐狸偷蛋");
  add('New, "Game"', "新游戏");
  assert.equal(api.state.mappingDrafts.length, 2, "Equivalent apostrophes update one draft");
  assert.equal(api.state.data.mapping["foxy's egg heist"], "狐狸偷蛋");

  const reloaded = setup("localhost", env.stored);
  reloaded.sandbox.fetch = async () => ({
    ok: true, text: async () => "english_name,display_name\nExisting,Remote latest\n",
  });
  await reloaded.api.loadMappingCsv();
  assert.equal(reloaded.api.state.mappingDrafts.length, 2);
  assert.equal(reloaded.api.state.data.mapping["foxy's egg heist"], "狐狸偷蛋");
  assert.equal(reloaded.api.state.data.mapping.existing, "Remote latest");
  reloaded.api.renderAdminMode();
  assert.match(reloaded.element("#mappingSyncStatus").textContent, /2/);

  await api.publishGameNameMapping();
  assert.match(element("#mappingSyncStatus").textContent, /Token/);
  element("#globalGithubToken").value = "test-only";
  const payloads = [];
  let reads = 0;
  sandbox.fetch = async (url, options) => {
    assert.match(url, /contents\/data\/game-name-mapping\.csv$/);
    assert.equal(options.headers.Authorization, "Bearer test-only");
    if (options.method === "PUT") {
      assert.equal(element("#addMappingButton").disabled, true);
      payloads.push(JSON.parse(options.body));
      return payloads.length === 1
        ? { ok: false, status: 409, text: async () => "Conflict" }
        : { ok: true, json: async () => ({ commit: { sha: "saved" } }) };
    }
    reads += 1;
    const csv = "english_name,display_name\nExisting,Remote latest\n"
      + (reads > 1 ? "Concurrent Game,Keep me\n" : "");
    return { ok: true, json: async () => ({ sha: "sha-" + reads, content: Buffer.from(csv).toString("base64") }) };
  };
  await api.publishGameNameMapping();
  assert.equal(payloads.length, 2);
  assert.equal(payloads[1].sha, "sha-2");
  assert.equal(payloads[1].branch, "main");
  const published = Buffer.from(payloads[1].content, "base64").toString("utf8");
  const mapping = api.parseMappingCsv(published);
  assert.equal(mapping.existing, "Remote latest", "Stale local mappings do not overwrite remote changes");
  assert.equal(mapping["concurrent game"], "Keep me", "Conflict retry preserves concurrent additions");
  assert.equal(mapping["foxy's egg heist"], "狐狸偷蛋");
  assert.equal(mapping['new, "game"'], "新游戏", "CSV quotes and Unicode round-trip");
  assert.match(published, /New,/);
  assert.equal(api.state.mappingDrafts.length, 0);
  assert.equal(element("#publishMappingButton").disabled, false);
  await api.publishGameNameMapping();
  assert.match(element("#mappingSyncStatus").textContent, /没有待同步/);

  add("Failed Game", "retain");
  sandbox.fetch = async () => ({ ok: false, status: 401, text: async () => "Bad credentials" });
  await api.publishGameNameMapping();
  assert.equal(api.state.mappingDrafts.length, 1, "Failed sync retains the draft");
  assert.match(element("#mappingSyncStatus").textContent, /401/);
  assert.equal(element("#publishMappingButton").disabled, false);

  api.state.adminMode = false;
  api.renderAdminMode();
  assert.equal(element("#adminPanel").classList.contains("is-open"), false);
  assert.equal(element("#addMappingButton").disabled, true);
  const remote = setup("dashboard.example");
  remote.api.state.adminMode = true;
  remote.api.addGameNameMapping();
  await remote.api.publishGameNameMapping();
  remote.api.renderAdminMode();
  assert.equal(remote.api.state.mappingDrafts.length, 0);
  assert.equal(remote.element(".admin-shell").classList.contains("is-hidden"), true);
  assert.equal(remote.element("#publishMappingButton").disabled, true);
  console.log("Game name mapping checks passed");
}

main().catch((error) => { console.error(error); process.exitCode = 1; });
