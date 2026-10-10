const assert = require("assert/strict");
const fs = require("fs");
const vm = require("vm");

const dataContext = { window: {} };
vm.createContext(dataContext);
vm.runInContext(fs.readFileSync("data/internal-test-data.js", "utf8"), dataContext);

const appSource = fs.readFileSync("app.js", "utf8").split("async function startDashboard()")[0]
  + "\nglobalThis.__test = { internalTestProblemObservations };";
const sandbox = {
  window: {
    INITIAL_WORKBOOK_DATA: { mapping: {}, weeks: [] },
    INTERNAL_TEST_DATA: dataContext.window.INTERNAL_TEST_DATA,
    location: { protocol: "http:", hostname: "localhost" },
  },
  localStorage: { getItem: () => null, setItem: () => {} },
  sessionStorage: { getItem: () => null, setItem: () => {} },
  document: { querySelector: () => null },
  console,
};
vm.createContext(sandbox);
vm.runInContext(appSource, sandbox);

const result = sandbox.__test.internalTestProblemObservations(dataContext.window.INTERNAL_TEST_DATA.games);
assert.equal(result.igcGames, 12);
assert.equal(result.eligibleGames, 11);
assert.deepEqual(Array.from(result.observations, (item) => item.name), ["Amazon Gold", "Canyon Beasts", "Frankenstein"]);
assert.equal(result.observations.every((item) => item.vendor === "IGC" && item.signals.length >= 2), true);
assert.equal(result.observations.find((item) => item.name === "Amazon Gold").signals.length, 3);

const nonIgcOnly = dataContext.window.INTERNAL_TEST_DATA.games.filter((game) => game.name === "Fortune Pyramid");
assert.equal(sandbox.__test.internalTestProblemObservations(nonIgcOnly).observations.length, 0);
console.log("Internal test observation checks passed");
