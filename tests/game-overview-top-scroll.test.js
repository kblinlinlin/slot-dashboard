const assert = require("assert");
const fs = require("fs");
const vm = require("vm");

const indexHtml = fs.readFileSync("index.html", "utf8");
assert.match(indexHtml, /id="gameOverviewTopScroll"/);
assert.match(indexHtml, /id="gameOverviewTableShell"/);

const appSource = fs.readFileSync("app.js", "utf8").split("async function startDashboard()")[0]
  + "\nglobalThis.__test = { syncTableScroll };";
const spacer = { style: {} };
const topScroll = {
  scrollLeft: 0,
  hidden: false,
  querySelector: () => spacer,
};
const tableShell = { scrollLeft: 0, clientWidth: 900 };
const table = { scrollWidth: 1600 };
const sandbox = {
  window: { INITIAL_WORKBOOK_DATA: { weeks: [], mapping: {} }, location: { protocol: "http:", hostname: "localhost" } },
  localStorage: { getItem: () => null, setItem: () => {} },
  sessionStorage: { getItem: () => null, setItem: () => {} },
  document: {
    querySelector: (selector) => ({
      "#gameOverviewTopScroll": topScroll,
      "#gameOverviewTableShell": tableShell,
      "#gameOverviewTable": table,
    })[selector] ?? null,
  },
  Object, String, Number, Set, Intl, console,
};
vm.createContext(sandbox);
vm.runInContext(appSource, sandbox);
sandbox.__test.syncTableScroll("#gameOverviewTopScroll", "#gameOverviewTableShell", "#gameOverviewTable");

assert.strictEqual(spacer.style.width, "1600px");
assert.strictEqual(topScroll.hidden, false);
topScroll.scrollLeft = 360;
topScroll.onscroll();
assert.strictEqual(tableShell.scrollLeft, 360);
tableShell.scrollLeft = 720;
tableShell.onscroll();
assert.strictEqual(topScroll.scrollLeft, 720);
