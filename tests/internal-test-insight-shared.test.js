const assert = require("assert/strict");
const fs = require("fs");
const vm = require("vm");

const stored = new Map();
const requests = [];
const appSource = fs.readFileSync("app.js", "utf8").split("async function startDashboard()")[0]
  + "\nglobalThis.__test = { saveSharedInternalInsight };";
const sandbox = {
  window: {
    INITIAL_WORKBOOK_DATA: { mapping: {}, weeks: [] },
    INTERNAL_TEST_DATA: { games: [] },
    location: { protocol: "http:", hostname: "localhost", origin: "http://localhost:8765" },
  },
  localStorage: {
    getItem: (key) => stored.get(key) ?? null,
    setItem: (key, value) => stored.set(key, value),
  },
  sessionStorage: { getItem: () => null, setItem: () => {} },
  document: { querySelector: () => null },
  fetch: async (url, options) => {
    requests.push({ url, options });
    return {
      ok: true,
      json: async () => ({ "id:1700131": { reporter: "共享报告人" } }),
    };
  },
  console,
};
vm.createContext(sandbox);
vm.runInContext(appSource, sandbox);

(async () => {
  const result = await sandbox.__test.saveSharedInternalInsight("id:1700131", {
    reporter: "共享报告人",
  });
  assert.equal(result.ok, true);
  assert.equal(result.reports["id:1700131"].reporter, "共享报告人");
  assert.equal(requests.length, 1);
  assert.equal(requests[0].url, "http://localhost:8765/api/internal-insights");
  assert.equal(requests[0].options.method, "POST");
  assert.deepEqual(JSON.parse(requests[0].options.body), {
    key: "id:1700131",
    report: { reporter: "共享报告人" },
  });
  console.log("Shared internal insight checks passed");
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
