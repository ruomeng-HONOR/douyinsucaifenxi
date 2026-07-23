import { analyzeRows, buildTrend, detectReportType, FUNCTION_META, normalizeExcelRow, normalizeStoredRows, REPORT_META, TIER_LABELS, validateReportHeaders } from "./analysis.js";
import { clearRows, readMaterialRows, readMonthlyRows, readRows, saveMonthlyRows, saveRows } from "./storage.js";
import { aggregateMonthlyRows, readLargeXlsx } from "./large-xlsx.js";

const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => [...document.querySelectorAll(selector)];
const money = (value) => `¥${Number(value || 0).toLocaleString("zh-CN", { maximumFractionDigits: value < 100 ? 2 : 0 })}`;
const fmt = (value, digits = 0) => Number(value || 0).toLocaleString("zh-CN", { maximumFractionDigits: digits });
const pct = (value) => `${(Number(value || 0) * 100).toFixed(1)}%`;
const esc = (value) => String(value ?? "").replace(/[&<>'"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" })[char]);
const debounce = (callback, delay = 300) => {
  let timer;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => callback(...args), delay);
  };
};
const MATERIAL_PAGE_SIZE = 100;
const ACTION_PAGE_SIZE = 100;
const ANALYSIS_CACHE_LIMIT = 12;

const state = {
  rows: [],
  analysis: null,
  filteredRows: [],
  selected: null,
  materialTier: "all",
  materialQuery: "",
  materialPage: 1,
  actionType: "全部",
  actionQuery: "",
  actionPage: 1,
  activeTab: "overview",
  renderedTabs: new Set(),
  analysisCache: new Map(),
};

const elements = {
  file: $("#excel-file"),
  notice: $("#notice"),
  filters: $("#filters"),
  mode: $("#mode-filter"),
  shop: $("#shop-filter"),
  account: $("#account-filter"),
  start: $("#start-date"),
  end: $("#end-date"),
  empty: $("#empty-state"),
  dashboard: $("#dashboard"),
};

function showNotice(message, error = false) {
  elements.notice.textContent = message;
  elements.notice.classList.toggle("error", error);
  elements.notice.hidden = !message;
  if (message) setTimeout(() => { if (elements.notice.textContent === message) elements.notice.hidden = true; }, 6000);
}

function updateBounds(resetDates = false) {
  const modes = [...new Set(state.rows.map((row) => row.deliveryMode || "推商品").filter(Boolean))].sort();
  const shops = [...new Set(state.rows.map((row) => row.shopCode).filter(Boolean))].sort();
  const accounts = [...new Set(state.rows.map((row) => row.qianchuanId).filter(Boolean))].sort();
  const months = state.rows.map((row) => row.month || row.date?.slice(0, 7)).filter(Boolean).sort();
  const minDate = months[0] || "";
  const maxDate = months.at(-1) || "";
  elements.mode.innerHTML = `<option value="all">全部类型</option>${modes.map((mode) => `<option value="${esc(mode)}">${esc(mode)}</option>`).join("")}`;
  elements.shop.innerHTML = `<option value="all">全部店铺</option>${shops.map((shop) => `<option value="${esc(shop)}">${esc(shop)}</option>`).join("")}`;
  elements.account.innerHTML = `<option value="all">全部账户</option>${accounts.map((account) => `<option value="${esc(account)}">${esc(account)}</option>`).join("")}`;
  if (resetDates || !elements.start.value) elements.start.value = minDate;
  if (resetDates || !elements.end.value) elements.end.value = maxDate;
  elements.start.min = minDate;
  elements.start.max = maxDate;
  elements.end.min = minDate;
  elements.end.max = maxDate;
  elements.filters.hidden = !state.rows.length;
  elements.empty.hidden = Boolean(state.rows.length);
  elements.dashboard.hidden = !state.rows.length;
}

function applyAnalysis() {
  const mode = elements.mode.value;
  const shop = elements.shop.value;
  const account = elements.account.value;
  const start = elements.start.value;
  const end = elements.end.value;
  const cacheKey = [mode, shop, account, start, end].join("::");
  const cached = state.analysisCache.get(cacheKey);
  const startedAt = performance.now();
  if (cached) {
    state.analysisCache.delete(cacheKey);
    state.analysisCache.set(cacheKey, cached);
    state.analysis = cached;
  } else {
    state.filteredRows = state.rows.filter((row) => (mode === "all" || row.deliveryMode === mode)
      && (shop === "all" || row.shopCode === shop)
      && (account === "all" || row.qianchuanId === account)
      && (row.month || row.date.slice(0, 7)) >= start
      && (row.month || row.date.slice(0, 7)) <= end);
    state.analysis = analyzeRows(state.filteredRows);
    state.analysisCache.set(cacheKey, state.analysis);
    while (state.analysisCache.size > ANALYSIS_CACHE_LIMIT) {
      state.analysisCache.delete(state.analysisCache.keys().next().value);
    }
  }
  state.selected = null;
  state.materialPage = 1;
  state.actionPage = 1;
  state.renderedTabs.clear();
  renderActiveTab(true);
  const elapsed = Math.round(performance.now() - startedAt);
  showNotice(cached ? `已读取分析缓存，耗时 ${elapsed}ms。` : `分析完成，耗时 ${(elapsed / 1000).toFixed(1)}秒；再次使用相同筛选将直接读取缓存。`);
}

function reportBadge(mode) {
  const meta = REPORT_META[mode] || REPORT_META["推商品"];
  return `<span class="mode-badge" style="--mode:${meta.color}">${esc(mode)}</span>`;
}

function compactDailyRow(row) {
  return {
    key: row.key,
    materialKey: row.materialKey,
    schemaVersion: row.schemaVersion,
    deliveryMode: row.deliveryMode,
    shopCode: row.shopCode,
    qianchuanId: row.qianchuanId,
    materialId: row.materialId,
    date: row.date,
    spend: row.spend,
    grossAmount: row.grossAmount,
    netAmount: row.netAmount,
    netOrders: row.netOrders,
  };
}

function renderMetrics() {
  const a = state.analysis;
  const s = a.summary;
  const good = a.materials.filter((item) => item.tier === "good");
  const goodSpend = good.reduce((sum, item) => sum + item.spend, 0);
  const goodNet = good.reduce((sum, item) => sum + item.netAmount, 0);
  return `<section class="metric-grid">
    <article><span>投放消耗</span><strong>${money(s.spend)}</strong><small>${a.materials.length} 条素材</small></article>
    <article><span>整体成交额 / ROI</span><strong>${money(s.grossAmount)}</strong><small>整体ROI ${s.grossRoi.toFixed(2)}</small></article>
    <article><span>净成交额 / ROI</span><strong class="good-text">${money(s.netAmount)}</strong><small>净ROI ${s.netRoi.toFixed(2)}</small></article>
    <article><span>退款损失</span><strong class="poor-text">${money(s.refundLoss)}</strong><small>金额损失率 ${pct(s.refundRate)}</small></article>
    <article><span>整体 / 净订单</span><strong>${fmt(s.grossOrders)}</strong><small>净成交 ${fmt(s.netOrders)} 单</small></article>
    <article><span>达到评分门槛</span><strong>${a.thresholds.eligible}</strong><small>消耗 ≥ ¥50</small></article>
    <article><span>优质素材贡献</span><strong>${a.tiers.good} 条</strong><small>消耗${pct(goodSpend / (s.spend || 1))} · 净成交${pct(goodNet / (s.netAmount || 1))}</small></article>
    <article><span>视频播放 / 完播率</span><strong>${fmt(s.videoPlays)}</strong><small>完播 ${pct(s.videoCompletionRate)} · 3秒 ${pct(s.threeSecondRate)}</small></article>
  </section>`;
}

function monthlyChart(data) {
  if (!data.length) return `<div class="chart-empty">所选范围暂无趋势数据</div>`;
  const width = 760, height = 292, left = 48, right = 44, top = 18, bottom = 34;
  const plotWidth = width - left - right, plotHeight = height - top - bottom;
  const maxSpend = Math.max(...data.map((item) => item.spend), 1);
  const maxRoi = Math.max(...data.map((item) => item.grossRoi), ...data.map((item) => item.netRoi), 1);
  const step = plotWidth / data.length, barWidth = Math.min(34, step * 0.48);
  const points = (key) => data.map((item, index) => `${left + (index + 0.5) * step},${top + plotHeight * (1 - item[key] / maxRoi)}`).join(" ");
  return `<svg class="trend-chart" viewBox="0 0 ${width} ${height}" role="img" aria-label="月度消耗、整体ROI和净ROI趋势">
    ${[0, 0.25, 0.5, 0.75, 1].map((t) => `<line class="grid-line" x1="${left}" x2="${width - right}" y1="${top + plotHeight * t}" y2="${top + plotHeight * t}"/><text class="axis-label" x="${left - 7}" y="${top + plotHeight * t + 4}" text-anchor="end">${fmt(maxSpend * (1 - t) / 10000, 1)}万</text>`).join("")}
    ${data.map((item, index) => { const h = plotHeight * item.spend / maxSpend; return `<rect class="spend-bar" x="${left + (index + 0.5) * step - barWidth / 2}" y="${top + plotHeight - h}" width="${barWidth}" height="${h}" rx="3"><title>${esc(item.month)} 消耗${money(item.spend)} 净ROI ${item.netRoi.toFixed(2)}</title></rect>`; }).join("")}
    <polyline points="${points("grossRoi")}" class="gross-line"/><polyline points="${points("netRoi")}" class="roi-line"/>
    ${data.map((item, index) => `<text class="axis-label" x="${left + (index + 0.5) * step}" y="${height - 8}" text-anchor="middle">${esc(item.month.slice(5))}</text>`).join("")}
  </svg>`;
}

function renderOverview() {
  const a = state.analysis;
  const s = a.summary;
  const maxCategory = Math.max(...a.categories.map((item) => item.spend), 1);
  $("#view-overview").innerHTML = `${renderMetrics()}
    <section class="insight-grid">
      <article><b>双模型评分</b><p>推商品更侧重商品成交与退款质量；推直播提高点击、完播和前5秒留存权重，两类素材不共用排名基准。</p></article>
      <article><b>退款侵蚀</b><p>整体ROI ${s.grossRoi.toFixed(2)} 降至净ROI ${s.netRoi.toFixed(2)}，退款造成 ${money(s.refundLoss)} 成交损失。</p></article>
      <article><b>评分可信度</b><p>采用¥${a.thresholds.priorWeight}先验权重收缩小样本ROI，并纳入完播、3秒和5秒播放质量。</p></article>
    </section>
    <section class="two-column">
      <article class="panel"><div class="panel-head"><div><h2>月度经营趋势</h2><p>消耗柱 · 整体ROI线 · 净ROI线</p></div><div class="chart-legend"><i class="spend"></i>消耗<i class="gross"></i>整体ROI<i class="net"></i>净ROI</div></div>${monthlyChart(a.monthly)}</article>
      <article class="panel"><div class="panel-head"><div><h2>等级分布</h2><p>按当前筛选范围重新计算</p></div></div><table><thead><tr><th>等级</th><th>素材数</th><th>占比</th></tr></thead><tbody>${["good", "normal", "poor"].map((tier) => `<tr><td><span class="tier ${tier}">${TIER_LABELS[tier]}</span></td><td>${a.tiers[tier]}</td><td>${pct(a.tiers[tier] / (a.materials.length || 1))}</td></tr>`).join("")}</tbody></table></article>
    </section>
    <section class="panel"><div class="panel-head"><div><h2>投放类型表现</h2><p>两类报表分别评分，经营结果在此汇总对比</p></div></div><div class="mode-summary">${a.modeSummaries.map((item) => `<article style="--mode:${REPORT_META[item.mode].color}"><div>${reportBadge(item.mode)}<b>${item.count}条素材</b></div><span>消耗<strong>${money(item.spend)}</strong></span><span>净ROI<strong>${item.netRoi.toFixed(2)}</strong></span><span>完播率<strong>${pct(item.videoCompletionRate)}</strong></span><span>有效评分<strong>${item.eligible}</strong></span></article>`).join("")}</div></section>
    <section class="panel"><div class="panel-head"><div><h2>品类经营表现</h2><p>根据素材名称关键词自动归类</p></div></div><div class="category-list">${a.categories.map((item) => `<div class="category-row"><div class="category-name"><b>${esc(item.name)}</b><small>${item.count}条</small></div><div class="bar-track"><i style="width:${item.spend / maxCategory * 100}%"></i></div><span>${money(item.spend)}</span><span>净ROI ${item.netRoi.toFixed(2)}</span><span>退款 ${pct(item.refundRate)}</span></div>`).join("")}</div></section>`;
}

function filteredMaterials() {
  const query = state.materialQuery.trim().toLowerCase();
  return state.analysis.materials.filter((item) => (state.materialTier === "all" || item.tier === state.materialTier) && (!query || item.id.toLowerCase().includes(query) || item.name.toLowerCase().includes(query)));
}

function paginationMarkup(page, total, pageSize, target) {
  const pages = Math.max(1, Math.ceil(total / pageSize));
  if (pages <= 1) return `<div class="pagination"><span>共 ${total.toLocaleString("zh-CN")} 条</span></div>`;
  const candidates = [1, page - 1, page, page + 1, pages]
    .filter((value) => value >= 1 && value <= pages)
    .filter((value, index, array) => array.indexOf(value) === index)
    .sort((a, b) => a - b);
  let previous = 0;
  const buttons = candidates.map((value) => {
    const gap = previous && value - previous > 1 ? "<i>…</i>" : "";
    previous = value;
    return `${gap}<button type="button" data-page-target="${target}" data-page="${value}" class="${value === page ? "active" : ""}">${value}</button>`;
  }).join("");
  return `<div class="pagination"><span>共 ${total.toLocaleString("zh-CN")} 条 · 第 ${page}/${pages} 页</span><div><button type="button" data-page-target="${target}" data-page="${page - 1}" ${page <= 1 ? "disabled" : ""}>上一页</button>${buttons}<button type="button" data-page-target="${target}" data-page="${page + 1}" ${page >= pages ? "disabled" : ""}>下一页</button></div></div>`;
}

function lifecycleChart(data) {
  if (!data.length) return `<div class="chart-empty">所选时间暂无投放记录</div>`;
  const width = 820, height = 305, left = 50, right = 44, top = 38, bottom = 36;
  const plotWidth = width - left - right, plotHeight = height - top - bottom;
  const maxSpend = Math.max(...data.map((item) => item.spend), 1), maxRoi = Math.max(...data.map((item) => item.rollRoi), 1);
  const step = plotWidth / data.length, barWidth = Math.max(2, Math.min(12, step * 0.58));
  const colors = { "测试期": "#a7b1c2", "放量期": "#4b78f0", "成熟期": "#18a689", "稳定期": "#73b9d2", "衰退期": "#e69a2b", "停投期": "#d94c55" };
  const path = data.map((item, index) => `${left + (index + 0.5) * step},${top + plotHeight * (1 - item.rollRoi / maxRoi)}`).join(" ");
  return `<svg class="trend-chart" viewBox="0 0 ${width} ${height}" role="img" aria-label="素材生命周期趋势">
    ${data.map((item, index) => `<rect x="${left + index * step}" y="8" width="${Math.max(1, step)}" height="13" fill="${colors[item.stage] || colors["稳定期"]}"><title>${esc(item.date)} ${esc(item.stage)}</title></rect>`).join("")}
    ${[0, 0.5, 1].map((t) => `<line class="grid-line" x1="${left}" x2="${width - right}" y1="${top + plotHeight * t}" y2="${top + plotHeight * t}"/>`).join("")}
    ${data.map((item, index) => { const h = plotHeight * item.spend / maxSpend; return `<rect class="spend-bar" x="${left + (index + 0.5) * step - barWidth / 2}" y="${top + plotHeight - h}" width="${barWidth}" height="${h}" rx="2"><title>${esc(item.date)} 消耗${money(item.spend)} 7日净ROI ${item.rollRoi.toFixed(2)} ${esc(item.stage)}</title></rect>`; }).join("")}
    <polyline points="${path}" class="roi-line"/>
    ${[0, Math.floor((data.length - 1) / 2), data.length - 1].filter((value, index, array) => array.indexOf(value) === index).map((index) => `<text class="axis-label" x="${left + (index + 0.5) * step}" y="${height - 8}" text-anchor="middle">${esc(data[index].date.slice(5))}</text>`).join("")}
  </svg>`;
}

async function renderMaterialDetail() {
  const container = $("#material-detail");
  if (!container) return;
  const item = state.selected;
  if (!item) {
    container.innerHTML = `<div class="detail-empty"><strong>选择一条素材</strong><span>查看评分理由、经营指标和投放生命周期</span></div>`;
    return;
  }
  container.innerHTML = `<div class="detail-empty"><strong>正在读取日趋势…</strong><span>仅加载当前素材，避免大型报表占满内存</span></div>`;
  const selectedKey = item.materialKey;
  let dailyRows = [];
  try {
    dailyRows = await readMaterialRows(item, `${elements.start.value}-01`, `${elements.end.value}-31`);
  } catch (error) {
    showNotice("无法读取该素材的日趋势数据", true);
  }
  if (state.selected?.materialKey !== selectedKey) return;
  const trend = buildTrend(dailyRows, item);
  const settlementRows = [[7, item.settlement7Roi, item.settlement7Rate], [14, item.settlement14Roi, item.settlement14Rate], [30, item.settlement30Roi, item.settlement30Rate], [90, item.settlement90Roi, item.settlement90Rate]];
  container.innerHTML = `<div class="detail-top"><span class="tier ${item.tier}">${TIER_LABELS[item.tier]} · ${item.score}分</span><span>${reportBadge(item.deliveryMode)} <span class="function-badge" style="--func:${FUNCTION_META[item.functionalType].color}">${esc(item.functionalType)}</span></span></div>
    <h2>${esc(item.name)}</h2><div class="material-id">店铺 ${esc(item.shopCode)} · 千川ID ${esc(item.qianchuanId || "-")} · 素材ID ${esc(item.id)}${item.videoType ? ` · ${esc(item.videoType)}` : ""}</div>
    <div class="decision-note"><b>分析理由：</b>${esc(item.reason)}<br><b>建议：</b>${esc(item.action)}</div>
    <div class="detail-metrics"><div><span>消耗</span><strong>${money(item.spend)}</strong></div><div><span>净ROI</span><strong>${item.netRoi.toFixed(2)}</strong></div><div><span>净成交</span><strong>${money(item.netAmount)}</strong></div><div><span>退款损失率</span><strong>${pct(item.refund)}</strong></div><div><span>视频播放</span><strong>${fmt(item.videoPlays)}</strong></div><div><span>完播 / 3秒</span><strong>${pct(item.videoCompletionRate)} / ${pct(item.threeSecondRate)}</strong></div><div><span>平均观看</span><strong>${fmt(item.avgWatchTime, 2)}秒</strong></div><div><span>内容质量分</span><strong>${item.contentScore.toFixed(0)}</strong></div></div>
    <div class="settlement-grid">${settlementRows.map(([days, roi, rate]) => `<div><span>${days}日结算</span><b>ROI ${roi.toFixed(2)}</b><small>GMV结算率 ${pct(rate)}</small></div>`).join("")}</div>
    <div class="panel-head"><div><h2>素材投放趋势</h2><p>日消耗柱 · 7日滚动净ROI线 · 顶部生命周期阶段</p></div></div>${lifecycleChart(trend)}`;
}

function renderMaterials() {
  const allItems = filteredMaterials();
  const pages = Math.max(1, Math.ceil(allItems.length / MATERIAL_PAGE_SIZE));
  state.materialPage = Math.min(Math.max(1, state.materialPage), pages);
  const offset = (state.materialPage - 1) * MATERIAL_PAGE_SIZE;
  const items = allItems.slice(offset, offset + MATERIAL_PAGE_SIZE);
  $("#view-materials").innerHTML = `<section class="tier-grid">
    <article><span>全部素材</span><strong>${state.analysis.materials.length}</strong><small>当前筛选范围</small></article>
    ${["good", "normal", "poor"].map((tier) => `<article><span>${TIER_LABELS[tier]}素材</span><strong>${state.analysis.tiers[tier]}</strong><small>${pct(state.analysis.tiers[tier] / (state.analysis.materials.length || 1))}</small></article>`).join("")}
    </section>
    <section class="workspace"><div class="table-panel"><div class="table-head"><div class="toolbar"><div><div class="score-title"><b>素材评分明细</b><span class="score-help"><button type="button" class="score-help-button" aria-describedby="score-definition">评分定义 <i>?</i></button><span id="score-definition" class="score-tooltip" role="tooltip"><strong>素材评分口径（0–100分）</strong><span><b>分级：</b>优质 ≥ 70分；普通 40–69分；劣质 ＜ 40分。</span><span><b>样本门槛：</b>所选周期消耗不足 ¥50 时记为0分，不进入稳定评分。</span><span><b>相对评价：</b>推商品与推直播分别在同类素材中排名，日期、店铺或账户改变后会重新计算。</span><span><b>核心指标：</b>贝叶斯整体/净ROI、点击率、转化率、净成交、消耗、退款质量、完播率及3秒/5秒留存；使用 ¥200 先验权重降低小样本波动。</span></span></span></div><p>点击素材查看投放趋势</p></div><div class="toolbar-group"><label class="field">等级<select id="tier-filter"><option value="all">全部等级</option><option value="good">优质</option><option value="normal">普通</option><option value="poor">劣质</option></select></label><label class="field">素材ID或名称<input id="material-search" placeholder="输入素材ID或名称"></label></div></div></div>
      <div class="table-wrap"><table><thead><tr><th>素材</th><th>投放类型</th><th>功能画像</th><th>评分</th><th>内容分</th><th>消耗</th><th>净ROI</th><th>退款</th></tr></thead><tbody>${items.map((item, index) => `<tr data-material="${index}" class="${state.selected?.id === item.id && state.selected?.shopCode === item.shopCode && state.selected?.deliveryMode === item.deliveryMode && state.selected?.qianchuanId === item.qianchuanId ? "selected" : ""}"><td class="material-name">${esc(item.name)}<small>${esc(item.id)} · ${esc(item.shopCode)} · ${esc(item.qianchuanId || "-")}</small></td><td>${reportBadge(item.deliveryMode)}</td><td><span class="function-badge" style="--func:${FUNCTION_META[item.functionalType].color}">${esc(item.functionalType)}</span></td><td><span class="tier ${item.tier}">${item.score} · ${TIER_LABELS[item.tier]}</span></td><td>${item.contentScore.toFixed(0)}</td><td>${money(item.spend)}</td><td>${item.netRoi.toFixed(2)}</td><td>${pct(item.refund)}</td></tr>`).join("")}</tbody></table></div>${paginationMarkup(state.materialPage, allItems.length, MATERIAL_PAGE_SIZE, "materials")}</div>
      <aside id="material-detail" class="detail-panel"></aside></section>`;
  $("#tier-filter").value = state.materialTier;
  $("#material-search").value = state.materialQuery;
  $("#tier-filter").addEventListener("change", (event) => { state.materialTier = event.target.value; state.materialPage = 1; state.selected = null; renderMaterials(); });
  const applyMaterialSearch = debounce((value) => { state.materialQuery = value; state.materialPage = 1; state.selected = null; renderMaterials(); });
  $("#material-search").addEventListener("input", (event) => applyMaterialSearch(event.target.value));
  $$("tr[data-material]").forEach((row) => row.addEventListener("click", () => {
    state.selected = items[Number(row.dataset.material)];
    $$("tr[data-material]").forEach((item) => item.classList.toggle("selected", item === row));
    void renderMaterialDetail();
  }));
  $$('[data-page-target="materials"]').forEach((button) => button.addEventListener("click", () => {
    if (button.disabled) return;
    state.materialPage = Number(button.dataset.page);
    state.selected = null;
    renderMaterials();
  }));
  void renderMaterialDetail();
  state.renderedTabs.add("materials");
}

function renderFunctional() {
  $("#view-functional").innerHTML = `<section class="function-grid">${state.analysis.functions.map((item) => { const meta = FUNCTION_META[item.type]; return `<article style="--func:${meta.color}"><div class="function-head"><i></i><h2>${esc(item.type)}</h2><b>${item.count}条</b></div><p>${esc(meta.description)}</p><div class="function-metrics"><span>素材占比<b>${pct(item.share)}</b></span><span>消耗占比<b>${pct(item.spendShare)}</b></span><span>净ROI<b>${item.netRoi.toFixed(2)}</b></span><span>净成交占比<b>${pct(item.netShare)}</b></span><span>退款损失率<b>${pct(item.refundRate)}</b></span></div></article>`; }).join("")}</section>
    <section class="panel"><div class="panel-head"><div><h2>功能画像判定逻辑</h2><p>两类报表的阈值分别计算，画像名称保持统一</p></div></div><p>爆款型：点击、转化、净ROI同时较高；精品型：低消耗、高净ROI；引流型：点击较高但转化偏低；损耗型：消耗较高但净ROI偏低；误导型：点击较高且退款损失率达到15%以上。综合评分额外纳入完播率、3秒和5秒播放率。</p></section>`;
}

function renderActions() {
  const types = ["全部", ...Object.keys(FUNCTION_META)];
  const query = state.actionQuery.trim().toLowerCase();
  const eligible = state.analysis.materials.filter((item) => item.spend >= 50);
  const counts = new Map([["全部", eligible.length]]);
  for (const item of eligible) counts.set(item.functionalType, (counts.get(item.functionalType) || 0) + 1);
  const allItems = eligible.filter((item) => (state.actionType === "全部" || item.functionalType === state.actionType) && (!query || item.id.toLowerCase().includes(query) || item.name.toLowerCase().includes(query)));
  const pages = Math.max(1, Math.ceil(allItems.length / ACTION_PAGE_SIZE));
  state.actionPage = Math.min(Math.max(1, state.actionPage), pages);
  const offset = (state.actionPage - 1) * ACTION_PAGE_SIZE;
  const items = allItems.slice(offset, offset + ACTION_PAGE_SIZE);
  $("#view-actions").innerHTML = `<section class="panel"><div class="panel-head"><div><h2>功能画像行动清单</h2><p>按统一的素材运营角色查看建议，共${allItems.length}条</p></div></div>
    <div class="toolbar"><div class="action-filter">${types.map((type) => `<button type="button" data-action-type="${esc(type)}" class="${state.actionType === type ? "active" : ""}" style="--dot:${type === "全部" ? "#19213a" : FUNCTION_META[type].color}"><i></i>${esc(type)} <b>${counts.get(type) || 0}</b></button>`).join("")}</div><label class="field">素材ID或名称<input id="action-search" placeholder="输入素材ID或名称" value="${esc(state.actionQuery)}"></label></div>
    <div class="table-wrap"><table><thead><tr><th>素材名称</th><th>素材ID</th><th>投放类型</th><th>功能画像</th><th>品类</th><th>消耗</th><th>净ROI</th><th>退款</th><th>建议动作</th></tr></thead><tbody>${items.map((item) => `<tr><td class="material-name">${esc(item.name)}<small>${esc(item.qianchuanId || "-")}</small></td><td>${esc(item.id)}</td><td>${reportBadge(item.deliveryMode)}</td><td><span class="function-badge" style="--func:${FUNCTION_META[item.functionalType].color}">${esc(item.functionalType)}</span></td><td>${esc(item.category)}</td><td>${money(item.spend)}</td><td>${item.netRoi.toFixed(2)}</td><td>${pct(item.refund)}</td><td class="action-copy">${esc(item.action)}</td></tr>`).join("")}</tbody></table></div>${paginationMarkup(state.actionPage, allItems.length, ACTION_PAGE_SIZE, "actions")}</section>`;
  $$('[data-action-type]').forEach((button) => button.addEventListener("click", () => { state.actionType = button.dataset.actionType; state.actionPage = 1; renderActions(); }));
  const applyActionSearch = debounce((value) => { state.actionQuery = value; state.actionPage = 1; renderActions(); });
  $("#action-search").addEventListener("input", (event) => applyActionSearch(event.target.value));
  $$('[data-page-target="actions"]').forEach((button) => button.addEventListener("click", () => {
    if (button.disabled) return;
    state.actionPage = Number(button.dataset.page);
    renderActions();
  }));
  state.renderedTabs.add("actions");
}

function renderActiveTab(force = false) {
  if (!state.analysis || (!force && state.renderedTabs.has(state.activeTab))) return;
  const renderers = {
    overview: renderOverview,
    materials: renderMaterials,
    functional: renderFunctional,
    actions: renderActions,
  };
  renderers[state.activeTab]?.();
  state.renderedTabs.add(state.activeTab);
}

async function handleUpload(file) {
  if (!file) return;
  showNotice("正在解析 Excel，请稍候…");
  try {
    let rows = [];
    let monthlyRows = [];
    let counts = {};
    if (file.size >= 50 * 1024 * 1024) {
      ({ monthlyRows, counts } = await readLargeXlsx(file, {
        collectRows: false,
        onProgress: (message) => showNotice(message),
        onBatch: (batch) => saveRows(batch.map(compactDailyRow), (saved, total) => showNotice(`正在保存当前批次：${saved.toLocaleString("zh-CN")} / ${total.toLocaleString("zh-CN")} 行…`)),
      }));
    } else {
      const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true, dense: true });
      for (const sheetName of workbook.SheetNames) {
        const sheet = workbook.Sheets[sheetName];
        const raw = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: true });
        if (!raw.length) continue;
        const headers = Object.keys(raw[0]);
        const reportType = detectReportType(headers, sheetName);
        if (!reportType) continue;
        const missing = validateReportHeaders(headers);
        if (missing.length) throw new Error(`${sheetName}缺少字段：${missing.join("、")}`);
        const normalized = raw.map((row) => normalizeExcelRow(row, reportType)).filter((row) => !row.isSummary && row.materialId && /^\d{4}-\d{2}-\d{2}$/.test(row.date));
        rows.push(...normalized);
        counts[reportType] = (counts[reportType] || 0) + normalized.length;
        await new Promise((resolve) => setTimeout(resolve, 0));
      }
      monthlyRows = aggregateMonthlyRows(rows);
      await saveRows(rows.map(compactDailyRow), (saved, total) => showNotice(`正在保存到浏览器：${saved.toLocaleString("zh-CN")} / ${total.toLocaleString("zh-CN")} 行…`));
    }
    const importedCount = Object.values(counts).reduce((sum, count) => sum + count, 0);
    if (!importedCount || !monthlyRows.length) throw new Error("未识别到有效的素材ID与日期");
    await saveMonthlyRows(monthlyRows, (saved, total) => showNotice(`正在保存月度汇总：${saved.toLocaleString("zh-CN")} / ${total.toLocaleString("zh-CN")} 条…`));
    rows = [];
    state.rows = normalizeStoredRows(await readMonthlyRows());
    state.analysisCache.clear();
    updateBounds(true);
    applyAnalysis();
    const detail = Object.entries(counts).map(([mode, count]) => `${mode} ${count.toLocaleString("zh-CN")}行`).join("，");
    showNotice(`已导入${detail}；生成月度素材记录 ${state.rows.length.toLocaleString("zh-CN")} 条。`);
  } catch (error) {
    showNotice(error instanceof Error ? error.message : "Excel 导入失败", true);
  } finally {
    elements.file.value = "";
  }
}

elements.file.addEventListener("change", (event) => handleUpload(event.target.files?.[0]));
$("#reanalyze").addEventListener("click", applyAnalysis);
$("#clear-data").addEventListener("click", async () => {
  if (!state.rows.length || !confirm("确定清空当前浏览器保存的全部素材数据吗？")) return;
  await clearRows();
  state.rows = [];
  state.analysis = null;
  state.selected = null;
  state.analysisCache.clear();
  state.renderedTabs.clear();
  updateBounds(true);
  showNotice("本地数据已清空。");
});
$$('[data-tab]').forEach((button) => button.addEventListener("click", () => {
  state.activeTab = button.dataset.tab;
  $$('[data-tab]').forEach((item) => item.classList.toggle("active", item === button));
  $$(".tab-view").forEach((view) => { view.hidden = view.id !== `view-${button.dataset.tab}`; });
  renderActiveTab();
}));

try {
  state.rows = normalizeStoredRows(await readMonthlyRows());
  if (!state.rows.length) {
    const legacyRows = normalizeStoredRows(await readRows());
    if (legacyRows.length) {
      state.rows = aggregateMonthlyRows(legacyRows);
      await saveMonthlyRows(state.rows);
    }
  }
  updateBounds(true);
  if (state.rows.length) applyAnalysis();
} catch (error) {
  showNotice("无法读取浏览器本地数据，请确认未禁用 IndexedDB。", true);
}
