const assert = require("assert");
const fs = require("fs");
const vm = require("vm");

const appSource = fs.readFileSync("app.js", "utf8").split("async function startDashboard()")[0]
  + "\nglobalThis.__test = { renderGameOverview, state };";
const gameOverviewTable = { innerHTML: "" };
const data = {
  mapping: {},
  weeks: [
    {
      period: "2026-07-20_2026-07-26",
      rows: [{
        日期: "2026-07-20_2026-07-26", 游戏ID: 1, 游戏名称: "Test Game", 英文名称: "Test Game",
        游戏Key: "test game", 显示名称: "Test Game", 产商: "AA", 下注金额排名变化: 1,
        下注金额: 10, 游戏输赢: 8, 投注次数: 12, 中奖RTP: 0.95, 新增玩家: 100, 活跃玩家: 80,
      }],
    },
    {
      period: "2026-07-13_2026-07-19",
      rows: [{
        日期: "2026-07-13_2026-07-19", 游戏ID: 1, 游戏名称: "Test Game", 英文名称: "Test Game",
        游戏Key: "test game", 显示名称: "Test Game", 产商: "AA", 下注金额排名变化: 2,
        下注金额: 9, 游戏输赢: 7, 投注次数: 11, 中奖RTP: 0.9, 新增玩家: 90, 活跃玩家: 70,
      }],
    },
  ],
};
const sandbox = {
  window: { INITIAL_WORKBOOK_DATA: data, location: { protocol: "http:", hostname: "localhost" } },
  localStorage: { getItem: () => null, setItem: () => {} },
  sessionStorage: { getItem: () => null, setItem: () => {} },
  document: { querySelector: (selector) => selector === "#gameOverviewTable" ? gameOverviewTable : null },
  Object, String, Number, Set, Intl, console,
};
vm.createContext(sandbox);
vm.runInContext(appSource, sandbox);
sandbox.__test.renderGameOverview();

assert.match(gameOverviewTable.innerHTML, /中奖RTP/);
assert.match(gameOverviewTable.innerHTML, /95\.00%/);
assert.match(gameOverviewTable.innerHTML, /90\.00%/);
assert.match(gameOverviewTable.innerHTML, /↑5\.00%/);
