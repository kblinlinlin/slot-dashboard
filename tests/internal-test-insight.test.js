const assert = require("assert/strict");
const fs = require("fs");
const vm = require("vm");

const dataContext = { window: {} };
vm.createContext(dataContext);
vm.runInContext(fs.readFileSync("data/internal-test-data.js", "utf8"), dataContext);

const stored = new Map();
const appSource = fs.readFileSync("app.js", "utf8").split("async function startDashboard()")[0]
  + "\nglobalThis.__test = { state, internalInsightKey, currentInternalInsight, saveInternalTestInsights };";
const sandbox = {
  window: {
    INITIAL_WORKBOOK_DATA: { mapping: {}, weeks: [] },
    INTERNAL_TEST_DATA: dataContext.window.INTERNAL_TEST_DATA,
    location: { protocol: "http:", hostname: "localhost" },
  },
  localStorage: {
    getItem: (key) => stored.get(key) ?? null,
    setItem: (key, value) => stored.set(key, value),
  },
  sessionStorage: { getItem: () => null, setItem: () => {} },
  document: { querySelector: () => null },
  console,
};
vm.createContext(sandbox);
vm.runInContext(appSource, sandbox);

const api = sandbox.__test;
const game = dataContext.window.INTERNAL_TEST_DATA.games.find((item) => item.name === "Amazon Gold");
const key = api.internalInsightKey(game);
const initial = api.currentInternalInsight(game);

assert.match(initial.internalData, /人均注單數/);
assert.match(initial.internalData, /当前 IGC 样本 P25/);
assert.equal(initial.reporter, "");
assert.equal(initial.mechanism, "");
assert.equal(initial.analysis.mechanics, "");

api.state.internalTestInsights[key] = {
  reporter: "测试报告人",
  mechanism: "已补充玩法说明",
  analysis: { mechanics: "已完成流程复盘" },
  conclusion: "等待复测",
};
const merged = api.currentInternalInsight(game);
assert.equal(merged.mechanism, "已补充玩法说明");
assert.equal(merged.reporter, "测试报告人");
assert.equal(merged.analysis.mechanics, "已完成流程复盘");
assert.equal(merged.analysis.pacing, "");

api.saveInternalTestInsights();
const saved = JSON.parse(stored.get("slot-dashboard-internal-test-insights-v1"));
assert.equal(saved[key].conclusion, "等待复测");
console.log("Internal test insight checks passed");
