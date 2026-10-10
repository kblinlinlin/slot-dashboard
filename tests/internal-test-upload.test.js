const assert = require("assert/strict");
const fs = require("fs");
const vm = require("vm");

const xlsxContext = {};
vm.createContext(xlsxContext);
vm.runInContext(fs.readFileSync("vendor/xlsx.full.min.js", "utf8"), xlsxContext);
const XLSX = xlsxContext.XLSX;
const appSource = fs.readFileSync("app.js", "utf8").split("async function startDashboard()")[0]
  + "\nglobalThis.__test = { parseInternalTestWorkbook, mergeInternalTestData, normalizeInternalTestData };";
const initialContext = { window: {} };
vm.createContext(initialContext);
vm.runInContext(fs.readFileSync("data/internal-test-data.js", "utf8"), initialContext);

const sandbox = {
  window: {
    INITIAL_WORKBOOK_DATA: { mapping: {}, weeks: [] },
    INTERNAL_TEST_DATA: initialContext.window.INTERNAL_TEST_DATA,
    XLSX,
    location: { protocol: "http:", hostname: "localhost" },
  },
  XLSX,
  localStorage: { getItem: () => null, setItem: () => {} },
  sessionStorage: { getItem: () => null, setItem: () => {} },
  document: { querySelector: () => null },
  console,
};
vm.createContext(sandbox);
vm.runInContext(appSource, sandbox);

const uploadedFile = "/Users/a0000/Downloads/内测数据  (1) (3).xlsx";
const workbook = XLSX.read(fs.readFileSync(uploadedFile), { type: "buffer", cellDates: false });
const uploaded = sandbox.__test.parseInternalTestWorkbook(workbook, "内测数据  (1) (3).xlsx");
assert.equal(uploaded.games.length, 12);
assert.equal(uploaded.games.every((game) => game.rows.length > 0), true);
assert.equal(uploaded.games.find((game) => game.name === "3 Pirate Ships").rows[0][0], "2026-08-11");

const previous = sandbox.__test.normalizeInternalTestData({
  sourceFile: "old.xlsx",
  games: [{
    name: "Existing",
    gameId: 1,
    columns: [{ key: "c0", label: "日期" }, { key: "c1", label: "游戏ID" }, { key: "c2", label: "游戏名称" }, { key: "c3", label: "指标" }],
    rows: [["2026-01-01", 1, "Existing", 10], ["2026-01-02", 1, "Existing", 20]],
  }],
});
const next = sandbox.__test.normalizeInternalTestData({
  sourceFile: "new.xlsx",
  games: [{
    name: "Existing",
    gameId: 1,
    columns: [{ key: "c0", label: "日期" }, { key: "c1", label: "游戏ID" }, { key: "c2", label: "游戏名称" }, { key: "c3", label: "指标" }],
    rows: [["2026-01-02", 1, "Existing", 200], ["2026-01-03", 1, "Existing", 30]],
  }, {
    name: "New",
    gameId: 2,
    columns: [{ key: "c0", label: "日期" }, { key: "c1", label: "游戏ID" }, { key: "c2", label: "游戏名称" }, { key: "c3", label: "指标" }],
    rows: [["2026-01-03", 2, "New", 40]],
  }],
});
const merged = sandbox.__test.mergeInternalTestData(previous, next);
const existing = merged.games.find((game) => game.gameId === 1);
assert.deepEqual(Array.from(existing.rows, (row) => row[3]), [10, 200, 30]);
assert.equal(merged.games.some((game) => game.gameId === 2), true);
assert.equal(merged.sourceFile, "new.xlsx");
console.log("Internal test upload checks passed");
