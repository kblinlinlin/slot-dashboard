const INITIAL_DATA = window.INITIAL_WORKBOOK_DATA;
const STORAGE_KEY = "slot-dashboard-workbook-data-v1";
const EXCLUDED_GAME_KEYS = new Set(["game lobby", "none", "secretary"]);

const VENDORS = ["AA", "IGC", "KK"];
const TOP_OPTIONS = [
  { label: "Top 10", value: "10" },
  { label: "Top 20", value: "20" },
  { label: "Top 30", value: "30" },
  { label: "Top 50", value: "50" },
  { label: "全部", value: "all" },
];
const PERIOD_OPTIONS = [
  { label: "最近 4 周", value: "4" },
  { label: "最近 8 周", value: "8" },
  { label: "最近 12 周", value: "12" },
  { label: "最近 24 周", value: "24" },
  { label: "自定义时间段", value: "custom" },
  { label: "全部周期", value: "all" },
];
const METRICS = [
  { label: "下注金额", key: "下注金额", type: "amount" },
  { label: "游戏输赢", key: "游戏输赢", type: "amount" },
  { label: "投注次数", key: "投注次数", type: "amount" },
  { label: "下注金额排名", key: "排名", type: "rank" },
  { label: "中奖RTP", key: "中奖RTP", type: "rate" },
  { label: "最大中奖倍数", key: "最大中奖倍数", type: "amount" },
  { label: "中奖率", key: "中奖率", type: "rate" },
  { label: "波动", key: "波动", type: "amount" },
  { label: "新增玩家", key: "新增玩家", type: "people" },
  { label: "活跃玩家", key: "活跃玩家", type: "people" },
];
const VENDOR_METRICS = [
  { label: "总下注金额", key: "bet", type: "amount" },
  { label: "总游戏输赢", key: "win", type: "amount" },
  { label: "总投注次数", key: "rounds", type: "amount" },
  { label: "总新增玩家", key: "newPlayers", type: "people" },
  { label: "平均新增玩家", key: "avgNewPlayers", type: "people" },
  { label: "总活跃玩家", key: "activePlayers", type: "people" },
  { label: "平均活跃玩家", key: "avgActivePlayers", type: "people" },
  { label: "平均 RTP", key: "avgRtp", type: "rate" },
  { label: "上榜游戏数", key: "count", type: "people" },
];

let state = {
  data: normalizeWorkbookData(loadStoredData() ?? INITIAL_DATA),
  activeTab: "gameOverview",
  overviewVendor: "全部",
  overviewTopN: "50",
  vendorTopN: "50",
  gameSearch: "",
  trendGames: [],
  trendGameSearch: "",
  trendGamePeriod: "12",
  trendGameStart: "",
  trendGameEnd: "",
  trendGameMetric: "下注金额",
  trendVendors: ["AA"],
  trendVendorPeriod: "12",
  trendVendorStart: "",
  trendVendorEnd: "",
  trendVendorTopN: "50",
  trendVendorMetric: "bet",
};

const $ = (selector) => document.querySelector(selector);

function loadStoredData() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

function saveStoredData() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state.data));
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function toNumber(value) {
  if (value === null || value === undefined || value === "") return null;
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function formatNumber(value, type = "amount") {
  const number = toNumber(value);
  if (number === null) return "-";
  if (type === "people") return new Intl.NumberFormat("zh-CN", { maximumFractionDigits: 0 }).format(number);
  if (type === "rate") return `${(number * 100).toFixed(2)}%`;
  if (Math.abs(number) >= 1000) return new Intl.NumberFormat("zh-CN", { maximumFractionDigits: 0 }).format(number);
  return new Intl.NumberFormat("zh-CN", { maximumFractionDigits: 2 }).format(number);
}

function normalizePeriod(value) {
  return String(value ?? "").replace(/\s+/g, "");
}

function gameKey(value) {
  return String(value ?? "").trim().toLowerCase();
}

function isExcludedGame(row) {
  return EXCLUDED_GAME_KEYS.has(row?.游戏Key ?? gameKey(row?.英文名称 ?? row?.游戏名称));
}

function periodLabel(period) {
  return normalizePeriod(period).replace("_", " 至 ");
}

function vendorFromGameId(gameId) {
  if (gameId === null || gameId === undefined || gameId === "") return "未知";
  const text = String(Number.isInteger(gameId) ? gameId : String(gameId).split(".")[0]);
  if (text.length === 7) return "IGC";
  if (text.length <= 3) return "AA";
  if (text.length === 6) return "KK";
  return "未知";
}

function getMetric(key, list = METRICS) {
  return list.find((metric) => metric.key === key) ?? list[0];
}

function rankOf(row) {
  return toNumber(row?.排名 ?? row?.下注金额排名变化);
}

function topRows(rows, topN = "50") {
  if (topN === "all") return [...rows];
  const limit = Number(topN);
  return rows.filter((row) => {
    const rank = rankOf(row);
    return rank !== null && rank > 0 && rank <= limit;
  });
}

function sortedByRank(rows) {
  return [...rows].sort((a, b) => (rankOf(a) ?? 99999) - (rankOf(b) ?? 99999));
}

function indexByEnglish(rows) {
  return new Map(rows.map((row) => [row.游戏Key ?? gameKey(row.英文名称), row]));
}

function currentRows() {
  return state.data.currentWeek?.rows?.length ? state.data.currentWeek.rows : state.data.weeks[0]?.rows ?? [];
}

function previousRows() {
  return state.data.previousWeek?.rows?.length ? state.data.previousWeek.rows : state.data.weeks[1]?.rows ?? [];
}

function currentPeriod() {
  return state.data.currentWeek?.period || state.data.weeks[0]?.period || "";
}

function previousPeriod() {
  return state.data.previousWeek?.period || state.data.weeks[1]?.period || "";
}

function sum(rows, key) {
  return rows.reduce((total, row) => total + (toNumber(row[key]) ?? 0), 0);
}

function average(rows, key) {
  const values = rows.map((row) => toNumber(row[key])).filter((value) => value !== null);
  if (!values.length) return 0;
  return values.reduce((total, value) => total + value, 0) / values.length;
}

function deltaClass(delta) {
  if (delta > 0) return "up";
  if (delta < 0) return "down";
  return "flat";
}

function changeText(current, previous, type = "amount", reverse = false) {
  const cur = toNumber(current);
  const prev = toNumber(previous);
  if (cur === null) return "-";
  if (prev === null) return "新上榜";
  const delta = reverse ? prev - cur : cur - prev;
  if (Math.abs(delta) < 0.000001) return "保持不变";
  return `${delta > 0 ? "↑" : "↓"}${formatNumber(Math.abs(delta), type)}`;
}

function changeClass(current, previous, reverse = false) {
  const cur = toNumber(current);
  const prev = toNumber(previous);
  if (cur === null || prev === null) return "flat";
  const delta = reverse ? prev - cur : cur - prev;
  return deltaClass(delta);
}

function populateSelect(selector, options, selectedValue) {
  const element = $(selector);
  element.innerHTML = options
    .map((option) => `<option value="${escapeHtml(option.value)}">${escapeHtml(option.label)}</option>`)
    .join("");
  if (Array.isArray(selectedValue)) {
    const selected = new Set(selectedValue);
    [...element.options].forEach((option) => {
      option.selected = selected.has(option.value);
    });
  } else {
    element.value = selectedValue;
  }
}

function getAllGames() {
  const map = new Map();
  for (const week of state.data.weeks) {
    for (const row of week.rows) {
      const key = row.游戏Key ?? gameKey(row.英文名称);
      if (!map.has(key)) map.set(key, row);
    }
  }
  for (const row of currentRows()) {
    const key = row.游戏Key ?? gameKey(row.英文名称);
    if (!map.has(key)) map.set(key, row);
  }
  const currentIndex = indexByEnglish(currentRows());
  return [...map.values()].sort((a, b) => {
    const rankA = rankOf(currentIndex.get(a.游戏Key ?? gameKey(a.英文名称))) ?? 999999;
    const rankB = rankOf(currentIndex.get(b.游戏Key ?? gameKey(b.英文名称))) ?? 999999;
    if (rankA !== rankB) return rankA - rankB;
    return a.显示名称.localeCompare(b.显示名称, "zh-CN");
  });
}

function vendorTotals(rows) {
  const result = {};
  for (const vendor of VENDORS) {
    const vendorRows = rows.filter((row) => row.产商 === vendor);
    result[vendor] = {
      vendor,
      count: vendorRows.length,
      top10: topRows(vendorRows, "10").length,
      top20: topRows(vendorRows, "20").length,
      top50: topRows(vendorRows, "50").length,
      bet: sum(vendorRows, "下注金额"),
      win: sum(vendorRows, "游戏输赢"),
      rounds: sum(vendorRows, "投注次数"),
      newPlayers: sum(vendorRows, "新增玩家"),
      avgNewPlayers: vendorRows.length ? sum(vendorRows, "新增玩家") / vendorRows.length : 0,
      activePlayers: sum(vendorRows, "活跃玩家"),
      avgActivePlayers: vendorRows.length ? sum(vendorRows, "活跃玩家") / vendorRows.length : 0,
      avgRtp: average(vendorRows, "中奖RTP"),
    };
  }
  return result;
}

function renderSummary() {
  $("#weeklySummary").textContent = generateWeeklySummary();
}

function formatReportDelta(value, type = "amount") {
  if (Math.abs(value) < 0.000001) return "持平";
  return `${value > 0 ? "增加" : "减少"} ${formatNumber(Math.abs(value), type)}`;
}

function rankDelta(row, previousIndex) {
  const previous = previousIndex.get(row.游戏Key ?? gameKey(row.英文名称));
  const currentRank = rankOf(row);
  const previousRank = rankOf(previous);
  if (currentRank === null || previousRank === null) return null;
  return previousRank - currentRank;
}

function generateWeeklySummary() {
  const current = currentRows();
  const previous = previousRows();
  const previousIndex = indexByEnglish(previous);
  const currentTop50 = topRows(current, "50");
  const previousTop50 = topRows(previous, "50");
  const currentTotals = vendorTotals(currentTop50);
  const previousTotals = vendorTotals(previousTop50);
  const lines = [
    `为您输出本周（${periodLabel(currentPeriod())}）对比上周（${periodLabel(previousPeriod())}）的数据变化总结：`,
    "",
    "【新游上线监测】",
  ];

  const newGames = sortedByRank(current.filter((row) => !previousIndex.has(row.游戏Key ?? gameKey(row.英文名称))));
  if (newGames.length) {
    lines.push(`本周全量榜单共发现 ${newGames.length} 款首次出现的新游戏：`);
    for (const row of newGames) {
      lines.push(`- ${row.显示名称}：首周排名第 ${formatNumber(rankOf(row), "people")}，下注金额 ${formatNumber(row.下注金额)}，新增玩家 ${formatNumber(row.新增玩家, "people")}。`);
    }
  } else {
    lines.push("本周全量榜单未发现首次出现的新游戏。");
  }

  lines.push("", "【产商大盘概况】");
  for (const vendor of VENDORS) {
    const currentVendor = currentTotals[vendor];
    const previousVendor = previousTotals[vendor];
    lines.push(
      `${vendor}：Top 50 上榜 ${currentVendor.count} 款；下注金额 ${formatNumber(currentVendor.bet)}，较上周${formatReportDelta(currentVendor.bet - previousVendor.bet)}；新增玩家 ${formatNumber(currentVendor.newPlayers, "people")}，较上周${formatReportDelta(currentVendor.newPlayers - previousVendor.newPlayers, "people")}；活跃玩家 ${formatNumber(currentVendor.activePlayers, "people")}，较上周${formatReportDelta(currentVendor.activePlayers - previousVendor.activePlayers, "people")}。`,
    );
  }

  lines.push("", "【头部游戏与 Top 10 厮杀】");
  const currentTop4 = sortedByRank(topRows(current, "10")).slice(0, 4);
  const previousTop4 = sortedByRank(topRows(previous, "10")).slice(0, 4);
  const currentTop4Names = new Set(currentTop4.map((row) => row.游戏Key ?? gameKey(row.英文名称)));
  const previousTop4Names = new Set(previousTop4.map((row) => row.游戏Key ?? gameKey(row.英文名称)));
  const enteredTop4 = currentTop4.filter((row) => !previousTop4Names.has(row.游戏Key ?? gameKey(row.英文名称)));
  const exitedTop4 = previousTop4.filter((row) => !currentTop4Names.has(row.游戏Key ?? gameKey(row.英文名称)));
  if (!enteredTop4.length && !exitedTop4.length) {
    lines.push(`本周 Top 4 与上周保持一致，头部格局相对固化，依次为：${currentTop4.map((row) => row.显示名称).join("、")}。`);
  } else {
    if (enteredTop4.length) lines.push(`本周进入 Top 4：${enteredTop4.map((row) => row.显示名称).join("、")}。`);
    if (exitedTop4.length) lines.push(`本周跌出 Top 4：${exitedTop4.map((row) => row.显示名称).join("、")}。`);
  }

  const currentTop10 = sortedByRank(topRows(current, "10"));
  const previousTop10 = sortedByRank(topRows(previous, "10"));
  const currentTop10Names = new Set(currentTop10.map((row) => row.游戏Key ?? gameKey(row.英文名称)));
  const previousTop10Names = new Set(previousTop10.map((row) => row.游戏Key ?? gameKey(row.英文名称)));
  const enteredTop10 = currentTop10.filter((row) => !previousTop10Names.has(row.游戏Key ?? gameKey(row.英文名称)));
  const exitedTop10 = previousTop10.filter((row) => !currentTop10Names.has(row.游戏Key ?? gameKey(row.英文名称)));
  const swaps = currentTop10
    .map((row) => [row, rankDelta(row, previousIndex)])
    .filter(([, delta]) => delta && Math.abs(delta) > 0)
    .slice(0, 5);
  if (enteredTop10.length) lines.push(`本周新进入 Top 10：${enteredTop10.map((row) => row.显示名称).join("、")}。`);
  if (exitedTop10.length) lines.push(`本周跌出 Top 10：${exitedTop10.map((row) => row.显示名称).join("、")}。`);
  if (swaps.length) {
    lines.push(`Top 10 内部名次变化：${swaps.map(([row, delta]) => `${row.显示名称}${delta > 0 ? "上升" : "下降"}${Math.abs(delta)}名至第${formatNumber(rankOf(row), "people")}`).join("；")}。`);
  }

  lines.push("", "【榜单异动与黑马】");
  const middleRows = sortedByRank(topRows(current, "30").filter((row) => rankOf(row) >= 11));
  const risers = [];
  const fallers = [];
  for (const row of middleRows) {
    const delta = rankDelta(row, previousIndex);
    if (delta === null) continue;
    if (delta >= 5) risers.push([row, delta]);
    if (delta <= -5) fallers.push([row, delta]);
  }
  risers.sort((a, b) => b[1] - a[1]);
  fallers.sort((a, b) => a[1] - b[1]);
  if (risers.length) {
    lines.push(`Top 11-30 区间黑马：${risers.slice(0, 5).map(([row, delta]) => `${row.显示名称}上升${delta}名至第${formatNumber(rankOf(row), "people")}`).join("；")}。`);
  }
  if (fallers.length) {
    lines.push(`Top 11-30 区间警示：${fallers.slice(0, 5).map(([row, delta]) => `${row.显示名称}下降${Math.abs(delta)}名至第${formatNumber(rankOf(row), "people")}`).join("；")}。`);
  }
  if (!risers.length && !fallers.length) {
    lines.push("Top 11-30 区间未出现 5 名及以上的大幅排名异动，整体相对稳定。");
  }

  lines.push("", "【流量池变化】");
  const currentNew = sum(currentTop50, "新增玩家");
  const previousNew = sum(previousTop50, "新增玩家");
  const currentActive = sum(currentTop50, "活跃玩家");
  const previousActive = sum(previousTop50, "活跃玩家");
  lines.push(
    `基于 Top 50 游戏，本周新增玩家合计 ${formatNumber(currentNew, "people")}，较上周${formatReportDelta(currentNew - previousNew, "people")}；活跃玩家合计 ${formatNumber(currentActive, "people")}，较上周${formatReportDelta(currentActive - previousActive, "people")}。`,
  );
  if (currentNew < previousNew && currentActive < previousActive) {
    lines.push("整体来看，本周 Top 50 流量池出现新增与活跃同步流失，需要重点关注头部游戏和主力产商的用户回落。");
  } else if (currentNew > previousNew && currentActive > previousActive) {
    lines.push("整体来看，本周 Top 50 流量池新增与活跃同步流入，平台流量表现改善。");
  } else {
    lines.push("整体来看，本周 Top 50 流量池出现结构性变化，新增与活跃玩家走势并不完全一致。");
  }
  return lines.join("\n");
}

function renderGameOverview() {
  const current = currentRows();
  const previousIndex = indexByEnglish(previousRows());
  const search = state.gameSearch.trim().toLowerCase();
  let rows = topRows(current, state.overviewTopN);
  if (state.overviewVendor !== "全部") rows = rows.filter((row) => row.产商 === state.overviewVendor);
  if (search) {
    rows = rows.filter((row) => `${row.显示名称} ${row.英文名称}`.toLowerCase().includes(search));
  }
  rows = sortedByRank(rows);

  $("#gameOverviewTable").innerHTML = `
    <thead>
      <tr class="group-head">
        <th rowspan="2">排名</th>
        <th rowspan="2">游戏</th>
        <th rowspan="2">产商</th>
        <th colspan="3">下注金额</th>
        <th colspan="3">游戏输赢</th>
        <th colspan="3">投注次数</th>
        <th colspan="3">新增玩家</th>
        <th colspan="3">活跃玩家</th>
      </tr>
      <tr>
        ${["本周", "上周", "变化"].map((label) => `<th>${label}</th>`).join("").repeat(5)}
      </tr>
    </thead>
    <tbody>
      ${rows.map((row) => {
        const previous = previousIndex.get(row.游戏Key ?? gameKey(row.英文名称));
        const cells = ["下注金额", "游戏输赢", "投注次数", "新增玩家", "活跃玩家"].map((field) => {
          const type = field.includes("玩家") ? "people" : "amount";
          return `
            <td class="num">${formatNumber(row[field], type)}</td>
            <td class="num">${formatNumber(previous?.[field], type)}</td>
            <td class="num ${changeClass(row[field], previous?.[field])}">${changeText(row[field], previous?.[field], type)}</td>
          `;
        }).join("");
        return `
          <tr>
            <td class="num">${formatNumber(rankOf(row), "people")}</td>
            <td><span class="game-title">${escapeHtml(row.显示名称)}</span><span class="game-subtitle">${escapeHtml(row.英文名称)}</span></td>
            <td><span class="pill teal">${escapeHtml(row.产商)}</span></td>
            ${cells}
          </tr>
        `;
      }).join("") || `<tr><td colspan="18" class="empty-state">没有符合条件的数据</td></tr>`}
    </tbody>
  `;
}

function renderVendorOverview() {
  const current = topRows(currentRows(), state.vendorTopN);
  const previous = topRows(previousRows(), state.vendorTopN);
  const currentTotals = vendorTotals(current);
  const previousTotals = vendorTotals(previous);

  $("#vendorMetricGrid").innerHTML = VENDORS.map((vendor) => {
    const cur = currentTotals[vendor];
    const prev = previousTotals[vendor];
    return `
      <article class="metric-card">
        <span>${vendor} 总下注金额</span>
        <strong>${formatNumber(cur.bet)}</strong>
        <em class="${deltaClass(cur.bet - prev.bet)}">较上周${formatReportDelta(cur.bet - prev.bet)}</em>
      </article>
    `;
  }).join("");

  $("#vendorOverviewTable").innerHTML = `
    <thead>
      <tr>
        <th>产商</th>
        <th>上榜数</th>
        <th>前10</th>
        <th>前20</th>
        <th>前50</th>
        <th>总下注金额</th>
        <th>变化</th>
        <th>总新增玩家</th>
        <th>变化</th>
        <th>总活跃玩家</th>
        <th>变化</th>
        <th>平均 RTP</th>
      </tr>
    </thead>
    <tbody>
      ${VENDORS.map((vendor) => {
        const cur = currentTotals[vendor];
        const prev = previousTotals[vendor];
        return `
          <tr>
            <td><span class="pill teal">${vendor}</span></td>
            <td class="num">${formatNumber(cur.count, "people")}</td>
            <td class="num">${formatNumber(cur.top10, "people")}</td>
            <td class="num">${formatNumber(cur.top20, "people")}</td>
            <td class="num">${formatNumber(cur.top50, "people")}</td>
            <td class="num">${formatNumber(cur.bet)}</td>
            <td class="num ${deltaClass(cur.bet - prev.bet)}">${formatReportDelta(cur.bet - prev.bet)}</td>
            <td class="num">${formatNumber(cur.newPlayers, "people")}</td>
            <td class="num ${deltaClass(cur.newPlayers - prev.newPlayers)}">${formatReportDelta(cur.newPlayers - prev.newPlayers, "people")}</td>
            <td class="num">${formatNumber(cur.activePlayers, "people")}</td>
            <td class="num ${deltaClass(cur.activePlayers - prev.activePlayers)}">${formatReportDelta(cur.activePlayers - prev.activePlayers, "people")}</td>
            <td class="num">${formatNumber(cur.avgRtp, "rate")}</td>
          </tr>
        `;
      }).join("")}
    </tbody>
  `;

  const previousIndex = indexByEnglish(previousRows());
  $("#vendorRankLists").innerHTML = VENDORS.map((vendor) => {
    const vendorRows = sortedByRank(topRows(currentRows(), "50").filter((row) => row.产商 === vendor)).slice(0, 12);
    return `
      <section class="vendor-list">
        <h3>${vendor} 上榜游戏</h3>
        <ol>
          ${vendorRows.map((row) => {
            const previous = previousIndex.get(row.游戏Key ?? gameKey(row.英文名称));
            return `<li><span>${escapeHtml(row.显示名称)}</span><strong class="${changeClass(rankOf(row), rankOf(previous), true)}">#${formatNumber(rankOf(row), "people")} / 上周 #${formatNumber(rankOf(previous), "people")}</strong></li>`;
          }).join("") || `<li><span>暂无数据</span><strong>-</strong></li>`}
        </ol>
      </section>
    `;
  }).join("");
}

function chronologicalWeeks() {
  return [...state.data.weeks].reverse();
}

function selectedWeeks(periodValue, startPeriod = "", endPeriod = "") {
  const weeks = [...state.data.weeks].reverse();
  if (periodValue === "custom") {
    if (!startPeriod || !endPeriod) return weeks.slice(Math.max(0, weeks.length - 12));
    const startIndex = weeks.findIndex((week) => week.period === startPeriod);
    const endIndex = weeks.findIndex((week) => week.period === endPeriod);
    if (startIndex < 0 || endIndex < 0) return weeks.slice(Math.max(0, weeks.length - 12));
    const from = Math.min(startIndex, endIndex);
    const to = Math.max(startIndex, endIndex);
    return weeks.slice(from, to + 1);
  }
  if (periodValue === "all") return weeks;
  return weeks.slice(Math.max(0, weeks.length - Number(periodValue)));
}

function renderLineChart(svgSelector, points, metric, title) {
  renderMultiLineChart(svgSelector, [{ name: title, points }], metric, title);
}

function renderMultiLineChart(svgSelector, seriesList, metric, title) {
  const svg = $(svgSelector);
  const width = 860;
  const height = 320;
  const padding = { top: 28, right: 24, bottom: 58, left: 64 };
  svg.setAttribute("viewBox", `0 0 ${width} ${height}`);

  const populatedSeries = seriesList.filter((series) => series.points.length);
  const allPoints = populatedSeries.flatMap((series) => series.points);

  if (!allPoints.length) {
    svg.innerHTML = `<text class="chart-label" x="${width / 2}" y="${height / 2}" text-anchor="middle">暂无趋势数据</text>`;
    return;
  }

  const values = allPoints.map((point) => point.value);
  const max = Math.max(...values);
  const min = Math.min(...values);
  const range = max - min || 1;
  const yFor = (value) => {
    const ratio = metric.type === "rank" ? (max - value) / range : (value - min) / range;
    return height - padding.bottom - ratio * (height - padding.top - padding.bottom);
  };
  const colors = ["#2563eb", "#0f766e", "#ad7b18", "#c43d32", "#7c3aed", "#0e7490"];
  const plottedSeries = populatedSeries.map((series, seriesIndex) => {
    const xStep = series.points.length > 1 ? (width - padding.left - padding.right) / (series.points.length - 1) : 0;
    const coords = series.points.map((point, index) => ({
      ...point,
      x: padding.left + index * xStep,
      y: yFor(point.value),
    }));
    return {
      ...series,
      color: colors[seriesIndex % colors.length],
      coords,
      line: coords.map((point, index) => `${index ? "L" : "M"} ${point.x} ${point.y}`).join(" "),
    };
  });
  const labelPoints = plottedSeries[0].coords;

  svg.innerHTML = `
    <text class="chart-label" x="${padding.left}" y="18">${escapeHtml(title)}</text>
    ${[0, 1, 2, 3].map((tick) => {
      const y = padding.top + tick * ((height - padding.top - padding.bottom) / 3);
      return `<line class="gridline" x1="${padding.left}" y1="${y}" x2="${width - padding.right}" y2="${y}"></line>`;
    }).join("")}
    ${plottedSeries.map((series) => `
      <path class="chart-line ${metric.type === "rank" ? "rank" : ""}" d="${series.line}" style="stroke:${series.color}"></path>
      ${series.coords.map((point) => `<circle class="chart-point" cx="${point.x}" cy="${point.y}" r="4" style="fill:${series.color}"></circle>`).join("")}
    `).join("")}
    ${labelPoints.map((point, index) => {
      if (labelPoints.length > 14 && index % 2 !== 0) return "";
      return `<text class="chart-label" x="${point.x}" y="${height - 28}" text-anchor="middle">${escapeHtml(point.label)}</text>`;
    }).join("")}
    ${plottedSeries.map((series, index) => {
      const x = padding.left + index * 128;
      const y = height - 8;
      return `<circle cx="${x}" cy="${y - 4}" r="4" fill="${series.color}"></circle><text class="chart-label" x="${x + 10}" y="${y}" text-anchor="start">${escapeHtml(series.name)}</text>`;
    }).join("")}
    <text class="chart-label" x="${padding.left}" y="${padding.top + 6}">${formatNumber(metric.type === "rank" ? min : max, metric.type)}</text>
    <text class="chart-label" x="${padding.left}" y="${height - padding.bottom}">${formatNumber(metric.type === "rank" ? max : min, metric.type)}</text>
  `;
}

function renderGameTrend() {
  const metric = getMetric(state.trendGameMetric);
  const weeks = selectedWeeks(state.trendGamePeriod, state.trendGameStart, state.trendGameEnd);
  const games = getAllGames();
  const selectedGames = state.trendGames.length ? state.trendGames : games.slice(0, 2).map((game) => game.游戏Key ?? gameKey(game.英文名称));
  const seriesList = selectedGames.map((selectedKey) => {
    const game = games.find((row) => (row.游戏Key ?? gameKey(row.英文名称)) === selectedKey);
    const points = weeks.map((week) => {
      const row = week.rows.find((item) => (item.游戏Key ?? gameKey(item.英文名称)) === selectedKey);
      const value = toNumber(row?.[metric.key]);
      return row && value !== null
        ? { label: week.start.slice(5), period: week.period, value, row, seriesName: game?.显示名称 ?? selectedKey }
        : null;
    }).filter(Boolean);
    return { name: game?.显示名称 ?? selectedKey, points };
  });

  renderMultiLineChart("#gameTrendChart", seriesList, metric, `${metric.label} - 多游戏对比`);
  renderTrendStats("#gameTrendStats", seriesList.flatMap((series) => series.points), metric, seriesList);
  $("#gameTrendTable").innerHTML = renderTrendTable(seriesList, metric);
}

function renderVendorTrend() {
  const metric = VENDOR_METRICS.find((item) => item.key === state.trendVendorMetric) ?? VENDOR_METRICS[0];
  const weeks = selectedWeeks(state.trendVendorPeriod, state.trendVendorStart, state.trendVendorEnd);
  const selectedVendors = state.trendVendors.length ? state.trendVendors : ["AA"];
  const seriesList = selectedVendors.map((vendor) => {
    const points = weeks.map((week) => {
      const rows = topRows(week.rows, state.trendVendorTopN).filter((row) => row.产商 === vendor);
      const totals = vendorTotals(rows)[vendor];
      return {
        label: week.start.slice(5),
        period: week.period,
        value: totals[metric.key],
        row: totals,
        seriesName: vendor,
      };
    });
    return { name: vendor, points };
  });
  renderMultiLineChart("#vendorTrendChart", seriesList, metric, `${metric.label} - 多产商对比`);
  renderTrendStats("#vendorTrendStats", seriesList.flatMap((series) => series.points), metric, seriesList);
  renderVendorBetPie(weeks);
  $("#vendorTrendTable").innerHTML = renderVendorTrendTable(seriesList, metric);
}

function renderVendorBetPie(weeks) {
  const rows = weeks.flatMap((week) => topRows(week.rows, state.trendVendorTopN));
  const totals = vendorTotals(rows);
  const data = VENDORS.map((vendor) => ({
    vendor,
    value: totals[vendor].bet,
  })).filter((item) => item.value > 0);
  const svg = $("#vendorBetPieChart");
  const stats = $("#vendorBetPieStats");
  const width = 420;
  const height = 300;
  const cx = 150;
  const cy = 150;
  const radius = 104;
  const colors = ["#c43d32", "#2563eb", "#0f766e"];
  const total = data.reduce((sumValue, item) => sumValue + item.value, 0);
  svg.setAttribute("viewBox", `0 0 ${width} ${height}`);

  if (!data.length || total <= 0) {
    svg.innerHTML = `<text class="chart-label" x="${width / 2}" y="${height / 2}" text-anchor="middle">暂无下注金额数据</text>`;
    stats.innerHTML = `<div class="empty-state">暂无占比数据</div>`;
    return;
  }

  let startAngle = -Math.PI / 2;
  const paths = data.map((item, index) => {
    const angle = (item.value / total) * Math.PI * 2;
    const endAngle = startAngle + angle;
    const x1 = cx + Math.cos(startAngle) * radius;
    const y1 = cy + Math.sin(startAngle) * radius;
    const x2 = cx + Math.cos(endAngle) * radius;
    const y2 = cy + Math.sin(endAngle) * radius;
    const largeArc = angle > Math.PI ? 1 : 0;
    const d = `M ${cx} ${cy} L ${x1} ${y1} A ${radius} ${radius} 0 ${largeArc} 1 ${x2} ${y2} Z`;
    const midAngle = startAngle + angle / 2;
    const labelX = cx + Math.cos(midAngle) * (radius + 35);
    const labelY = cy + Math.sin(midAngle) * (radius + 18);
    startAngle = endAngle;
    return `
      <path d="${d}" fill="${colors[index % colors.length]}"></path>
      <text class="chart-label" x="${labelX}" y="${labelY}" text-anchor="middle">${item.vendor}</text>
    `;
  }).join("");

  svg.innerHTML = `
    <text class="chart-label" x="18" y="22">所选周期 Top ${state.trendVendorTopN === "all" ? "全部" : state.trendVendorTopN} 总下注金额占比</text>
    ${paths}
    <circle cx="${cx}" cy="${cy}" r="48" fill="#fbfcfd"></circle>
    <text class="chart-label" x="${cx}" y="${cy - 4}" text-anchor="middle">总下注</text>
    <text class="chart-label" x="${cx}" y="${cy + 16}" text-anchor="middle">${formatNumber(total)}</text>
  `;
  stats.innerHTML = data.map((item, index) => {
    const share = item.value / total;
    return `
      <div class="stat-row">
        <span><i style="display:inline-block;width:9px;height:9px;border-radius:50%;background:${colors[index % colors.length]};margin-right:6px"></i>${item.vendor}</span>
        <strong>${formatNumber(item.value)}</strong>
        <small>占比 ${(share * 100).toFixed(1)}%</small>
      </div>
    `;
  }).join("");
}

function renderTrendStats(selector, points, metric, seriesList = []) {
  const element = $(selector);
  if (!points.length) {
    element.innerHTML = `<div class="empty-state">暂无统计数据</div>`;
    return;
  }
  if (seriesList.length > 1) {
    element.innerHTML = seriesList.map((series) => {
      const latest = series.points.at(-1);
      if (!latest) return "";
      return `
        <div class="stat-row">
          <span>${escapeHtml(series.name)} 最新值</span>
          <strong>${formatNumber(latest.value, metric.type)}</strong>
          <small>${periodLabel(latest.period)}</small>
        </div>
      `;
    }).join("");
    return;
  }
  const values = points.map((point) => point.value);
  const latest = points.at(-1);
  const best = metric.type === "rank"
    ? points.reduce((winner, point) => (point.value < winner.value ? point : winner), points[0])
    : points.reduce((winner, point) => (point.value > winner.value ? point : winner), points[0]);
  const worst = metric.type === "rank"
    ? points.reduce((winner, point) => (point.value > winner.value ? point : winner), points[0])
    : points.reduce((winner, point) => (point.value < winner.value ? point : winner), points[0]);
  const avg = values.reduce((total, value) => total + value, 0) / values.length;
  element.innerHTML = [
    ["最新值", formatNumber(latest.value, metric.type), periodLabel(latest.period)],
    ["周期均值", formatNumber(avg, metric.type), `${points.length} 个周期`],
    ["最佳表现", formatNumber(best.value, metric.type), periodLabel(best.period)],
    ["最低表现", formatNumber(worst.value, metric.type), periodLabel(worst.period)],
  ].map(([label, value, desc]) => `
    <div class="stat-row">
      <span>${label}</span>
      <strong>${value}</strong>
      <small>${desc}</small>
    </div>
  `).join("");
}

function renderTrendTable(seriesList, metric) {
  const rows = seriesList.flatMap((series) => series.points.map((point) => ({ ...point, seriesName: series.name })));
  return `
    <thead><tr><th>游戏</th><th>周期</th><th>${metric.label}</th><th>排名</th><th>下注金额</th><th>新增玩家</th><th>活跃玩家</th><th>是否有活动</th></tr></thead>
    <tbody>
      ${rows.map((point) => `
        <tr>
          <td>${escapeHtml(point.seriesName)}</td>
          <td>${periodLabel(point.period)}</td>
          <td class="num">${formatNumber(point.value, metric.type)}</td>
          <td class="num">${formatNumber(rankOf(point.row), "people")}</td>
          <td class="num">${formatNumber(point.row.下注金额)}</td>
          <td class="num">${formatNumber(point.row.新增玩家, "people")}</td>
          <td class="num">${formatNumber(point.row.活跃玩家, "people")}</td>
          <td class="num">${formatNumber(point.row.是否有活动, "people")}</td>
        </tr>
      `).join("") || `<tr><td colspan="8" class="empty-state">暂无趋势数据</td></tr>`}
    </tbody>
  `;
}

function renderVendorTrendTable(seriesList, metric) {
  const rows = seriesList.flatMap((series) => series.points.map((point) => ({ ...point, seriesName: series.name })));
  return `
    <thead><tr><th>产商</th><th>周期</th><th>${metric.label}</th><th>上榜数</th><th>总下注金额</th><th>总游戏输赢</th><th>总新增玩家</th><th>总活跃玩家</th><th>平均 RTP</th></tr></thead>
    <tbody>
      ${rows.map((point) => `
        <tr>
          <td><span class="pill teal">${escapeHtml(point.seriesName)}</span></td>
          <td>${periodLabel(point.period)}</td>
          <td class="num">${formatNumber(point.value, metric.type)}</td>
          <td class="num">${formatNumber(point.row.count, "people")}</td>
          <td class="num">${formatNumber(point.row.bet)}</td>
          <td class="num">${formatNumber(point.row.win)}</td>
          <td class="num">${formatNumber(point.row.newPlayers, "people")}</td>
          <td class="num">${formatNumber(point.row.activePlayers, "people")}</td>
          <td class="num">${formatNumber(point.row.avgRtp, "rate")}</td>
        </tr>
      `).join("")}
    </tbody>
  `;
}

function renderAll() {
  populateControls();
  renderSummary();
  renderGameOverview();
  renderVendorOverview();
  renderGameTrend();
  renderVendorTrend();
}

function populateControls() {
  populateSelect("#overviewVendor", [{ label: "全部", value: "全部" }, ...VENDORS.map((vendor) => ({ label: vendor, value: vendor }))], state.overviewVendor);
  populateSelect("#overviewTopN", TOP_OPTIONS, state.overviewTopN);
  populateSelect("#vendorTopN", TOP_OPTIONS, state.vendorTopN);
  populateSelect("#trendGamePeriod", PERIOD_OPTIONS, state.trendGamePeriod);
  populateSelect("#trendVendorPeriod", PERIOD_OPTIONS, state.trendVendorPeriod);
  const weekOptions = chronologicalWeeks().map((week) => ({ label: periodLabel(week.period), value: week.period }));
  if (!state.trendGameStart && weekOptions.length) state.trendGameStart = weekOptions[Math.max(0, weekOptions.length - 12)].value;
  if (!state.trendGameEnd && weekOptions.length) state.trendGameEnd = weekOptions.at(-1).value;
  if (!state.trendVendorStart && weekOptions.length) state.trendVendorStart = weekOptions[Math.max(0, weekOptions.length - 12)].value;
  if (!state.trendVendorEnd && weekOptions.length) state.trendVendorEnd = weekOptions.at(-1).value;
  populateSelect("#trendGameStart", weekOptions, state.trendGameStart);
  populateSelect("#trendGameEnd", weekOptions, state.trendGameEnd);
  populateSelect("#trendVendorStart", weekOptions, state.trendVendorStart);
  populateSelect("#trendVendorEnd", weekOptions, state.trendVendorEnd);
  $("#gameStartWrap").classList.toggle("is-visible", state.trendGamePeriod === "custom");
  $("#gameEndWrap").classList.toggle("is-visible", state.trendGamePeriod === "custom");
  $("#vendorStartWrap").classList.toggle("is-visible", state.trendVendorPeriod === "custom");
  $("#vendorEndWrap").classList.toggle("is-visible", state.trendVendorPeriod === "custom");
  populateSelect("#trendVendorTopN", TOP_OPTIONS, state.trendVendorTopN);
  populateSelect("#trendGameMetric", METRICS.map((metric) => ({ label: metric.label, value: metric.key })), state.trendGameMetric);
  if (!state.trendVendors.length) state.trendVendors = ["AA"];
  populateSelect("#trendVendor", VENDORS.map((vendor) => ({ label: vendor, value: vendor })), state.trendVendors);
  populateSelect("#trendVendorMetric", VENDOR_METRICS.map((metric) => ({ label: metric.label, value: metric.key })), state.trendVendorMetric);

  const allGames = getAllGames();
  state.trendGames = state.trendGames.map(gameKey).filter((name) => allGames.some((game) => (game.游戏Key ?? gameKey(game.英文名称)) === name));
  if (!state.trendGames.length) {
    state.trendGames = allGames.slice(0, 2).map((game) => game.游戏Key ?? gameKey(game.英文名称));
  }
  const query = state.trendGameSearch.trim().toLowerCase();
  const selectedSet = new Set(state.trendGames);
  const currentIndex = indexByEnglish(currentRows());
  const games = allGames.filter((game) => {
    const key = game.游戏Key ?? gameKey(game.英文名称);
    if (selectedSet.has(key)) return true;
    if (!query) return true;
    return `${game.显示名称} ${game.英文名称}`.toLowerCase().includes(query);
  });
  populateSelect("#trendGame", games.map((game) => {
    const key = game.游戏Key ?? gameKey(game.英文名称);
    const rank = rankOf(currentIndex.get(key));
    const rankLabel = rank ? `#${formatNumber(rank, "people")} ` : "";
    return { label: `${rankLabel}${game.显示名称} (${game.英文名称})`, value: key };
  }), state.trendGames);
  $("#trendGameSearch").value = state.trendGameSearch;
  $("#gameSearch").value = state.gameSearch;
}

function wireEvents() {
  document.querySelectorAll(".tab").forEach((button) => {
    button.addEventListener("click", () => {
      state.activeTab = button.dataset.tab;
      document.querySelectorAll(".tab").forEach((item) => item.classList.toggle("is-active", item === button));
      document.querySelectorAll(".tab-panel").forEach((panel) => panel.classList.toggle("is-active", panel.id === state.activeTab));
    });
  });

  const bindings = [
    ["#overviewVendor", "overviewVendor", renderGameOverview],
    ["#overviewTopN", "overviewTopN", renderGameOverview],
    ["#vendorTopN", "vendorTopN", renderVendorOverview],
    ["#trendGame", "trendGames", renderGameTrend],
    ["#trendGamePeriod", "trendGamePeriod", () => { populateControls(); renderGameTrend(); }],
    ["#trendGameStart", "trendGameStart", renderGameTrend],
    ["#trendGameEnd", "trendGameEnd", renderGameTrend],
    ["#trendGameMetric", "trendGameMetric", renderGameTrend],
    ["#trendVendor", "trendVendors", renderVendorTrend],
    ["#trendVendorPeriod", "trendVendorPeriod", () => { populateControls(); renderVendorTrend(); }],
    ["#trendVendorStart", "trendVendorStart", renderVendorTrend],
    ["#trendVendorEnd", "trendVendorEnd", renderVendorTrend],
    ["#trendVendorTopN", "trendVendorTopN", renderVendorTrend],
    ["#trendVendorMetric", "trendVendorMetric", renderVendorTrend],
  ];
  for (const [selector, key, render] of bindings) {
    $(selector).addEventListener("change", (event) => {
      state[key] = event.target.multiple
        ? [...event.target.selectedOptions].map((option) => option.value)
        : event.target.value;
      render();
    });
  }
  $("#gameSearch").addEventListener("input", (event) => {
    state.gameSearch = event.target.value;
    renderGameOverview();
  });
  $("#trendGameSearch").addEventListener("input", (event) => {
    state.trendGameSearch = event.target.value;
    populateControls();
  });
  $("#copySummaryButton").addEventListener("click", async () => {
    await navigator.clipboard.writeText($("#weeklySummary").textContent);
    $("#copySummaryButton").textContent = "已复制";
    setTimeout(() => ($("#copySummaryButton").textContent = "复制总结"), 1200);
  });
  $("#fileInput").addEventListener("change", handleFileUpload);
}

async function handleFileUpload(event) {
  const file = event.target.files?.[0];
  if (!file) return;
  if (!window.XLSX) {
    alert("Excel 解析库尚未加载，请检查网络后刷新页面。");
    return;
  }
  const buffer = await file.arrayBuffer();
  try {
    const workbook = XLSX.read(buffer, { type: "array", cellDates: false });
    const parsed = parseWorkbook(workbook, file.name);
    if (parsed.kind === "full") {
      state.data = parsed.data;
      $("#sourceLabel").textContent = `当前加载：${file.name}（完整看板）`;
    } else {
      state.data = appendSingleWeek(state.data, parsed.week, file.name);
      $("#sourceLabel").textContent = `已追加/更新单周：${periodLabel(parsed.week.period)}`;
    }
    saveStoredData();
    state.trendGames = [];
    state.trendVendors = ["AA"];
    state.trendGameStart = "";
    state.trendGameEnd = "";
    state.trendVendorStart = "";
    state.trendVendorEnd = "";
    renderAll();
  } catch (error) {
    alert(`表格解析失败：${error.message}`);
  }
}

function sheetRows(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return [];
  return XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
}

function parseWorkbook(workbook, sourceFile) {
  const mappingRows = sheetRows(workbook, "游戏名称映射");
  const mapping = {};
  for (const row of mappingRows.slice(1)) {
    if (row[0] && row[1]) mapping[gameKey(row[0])] = String(row[1]).trim();
  }
  const displayMapping = { ...(state.data?.mapping ?? {}), ...mapping };
  const loadRows = (sheetName) => {
    const rows = sheetRows(workbook, sheetName);
    if (!rows.length) return [];
    const headers = rows[0].map((item) => String(item ?? "").trim());
    return rows.slice(1).filter((row) => row[0] && row[1]).map((row) => {
      const item = {};
      headers.forEach((header, index) => {
        if (header) item[header] = row[index];
      });
      const english = String(item.游戏名称 ?? "").trim();
      item.英文名称 = english;
      item.游戏Key = gameKey(english);
      item.显示名称 = displayMapping[item.游戏Key] || english;
      item.产商 = vendorFromGameId(item.游戏ID);
      item.日期 = normalizePeriod(item.日期 || sheetName);
      item.排名 = item.下注金额排名变化;
      return item;
    }).filter((row) => !isExcludedGame(row));
  };

  const weekNames = workbook.SheetNames.filter((name) => /^\d{4}-\d{2}-\d{2}\s*_\s*\d{4}-\d{2}-\d{2}$/.test(name.trim()));
  const weeks = weekNames.map((sheetName) => {
    const rows = loadRows(sheetName);
    const period = normalizePeriod(rows[0]?.日期 || sheetName);
    const [start = period, end = period] = period.split("_");
    return { sheetName, period, start, end, rows };
  });
  const hasFullWorkbookSignals = workbook.SheetNames.includes("游戏名称映射")
    || workbook.SheetNames.includes("上周数据")
    || weekNames.length > 1
    || (workbook.SheetNames.includes("本周数据") && weekNames.length > 0);

  if (hasFullWorkbookSignals) {
    return {
      kind: "full",
      data: normalizeWorkbookData({
        sourceFile,
        generatedAt: new Date().toISOString(),
        vendors: VENDORS,
        mapping: displayMapping,
        currentWeek: { period: loadRows("本周数据")[0]?.日期 || weeks[0]?.period || "", rows: loadRows("本周数据") },
        previousWeek: { period: loadRows("上周数据")[0]?.日期 || weeks[1]?.period || "", rows: loadRows("上周数据") },
        weeks,
      }),
    };
  }

  const candidateSheet = workbook.SheetNames.find((name) => sheetRows(workbook, name).length > 1);
  const rows = candidateSheet ? loadRows(candidateSheet) : [];
  if (!rows.length) {
    throw new Error("未识别到可追加的单周数据");
  }
  const period = normalizePeriod(rows[0]?.日期 || candidateSheet);
  const [start = period, end = period] = period.split("_");
  return {
    kind: "single-week",
    week: {
      sheetName: candidateSheet,
      period,
      start,
      end,
      rows,
    },
  };
}

function normalizeWorkbookData(data) {
  const weeks = [...(data.weeks ?? [])]
    .filter((week) => week.period && week.rows?.length)
    .map((week) => ({ ...week, rows: normalizeRows(week.rows, data.mapping ?? {}) }))
    .sort((a, b) => b.period.localeCompare(a.period));
  const currentWeek = data.currentWeek?.rows?.length
    ? { ...data.currentWeek, rows: normalizeRows(data.currentWeek.rows, data.mapping ?? {}) }
    : { period: weeks[0]?.period ?? "", rows: weeks[0]?.rows ?? [] };
  const previousWeek = data.previousWeek?.rows?.length
    ? { ...data.previousWeek, rows: normalizeRows(data.previousWeek.rows, data.mapping ?? {}) }
    : { period: weeks[1]?.period ?? "", rows: weeks[1]?.rows ?? [] };
  return { ...data, mapping: normalizeMapping(data.mapping ?? {}), weeks, currentWeek, previousWeek };
}

function normalizeMapping(mapping) {
  return Object.fromEntries(Object.entries(mapping).map(([key, value]) => [gameKey(key), value]));
}

function normalizeRows(rows, mapping = {}) {
  const normalizedMapping = normalizeMapping(mapping);
  return rows
    .map((row) => {
      const english = String(row.英文名称 ?? row.游戏名称 ?? "").trim();
      const key = row.游戏Key ?? gameKey(english);
      return {
        ...row,
        英文名称: english,
        游戏Key: key,
        显示名称: normalizedMapping[key] || row.显示名称 || english,
      };
    })
    .filter((row) => !isExcludedGame(row));
}

function appendSingleWeek(data, week, sourceFile) {
  const mergedWeeks = [...data.weeks.filter((item) => item.period !== week.period), week]
    .sort((a, b) => b.period.localeCompare(a.period));
  return normalizeWorkbookData({
    ...data,
    sourceFile,
    generatedAt: new Date().toISOString(),
    weeks: mergedWeeks,
    currentWeek: { period: mergedWeeks[0]?.period ?? "", rows: mergedWeeks[0]?.rows ?? [] },
    previousWeek: { period: mergedWeeks[1]?.period ?? "", rows: mergedWeeks[1]?.rows ?? [] },
  });
}

wireEvents();
renderAll();
