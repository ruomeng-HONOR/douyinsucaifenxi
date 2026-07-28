import { analyzeRows, buildTrend, detectReportType, FUNCTION_META, normalizeExcelRow, normalizeStoredRows, REPORT_META, TIER_LABELS, validateReportHeaders } from "./analysis.js";
import { clearPlatformRows, readMaterialRows, readMonthlyRows, readRows, saveMonthlyRows, saveRows } from "./storage.js";
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
const SYSTEMS = {
  douyin: {
    platform: "抖音",
    eyebrow: "DOUYIN MATERIAL OPERATING SYSTEM",
    title: "抖音视频素材分析系统",
    description: "独立分析千川推商品与推直播素材 · 数据仅保存在当前浏览器",
    upload: "＋ 上传抖音报表",
    emptyTitle: "上传第一份抖音素材数据",
    emptyCopy: "支持千川“推商品”和“推直播”短视频素材报表，重点分析净ROI、退款质量、转化与视频留存。",
    accountLabel: "千川账户",
    materialLabel: "素材ID",
    rangeMeta: "抖音评分基于净ROI、退款质量与视频留存",
  },
  taobao: {
    platform: "淘系",
    eyebrow: "TAOBAO CONTENT MARKETING SYSTEM",
    title: "淘宝视频素材分析系统",
    description: "独立分析万相台内容营销短视频 · 数据仅保存在当前浏览器",
    upload: "＋ 上传淘宝报表",
    emptyTitle: "上传第一份淘宝素材数据",
    emptyCopy: "支持万相台内容营销“短视频推广效果”报表，重点分析成交ROI、有效观看、收藏加购、访问与新客。",
    accountLabel: "投放计划",
    materialLabel: "主体ID",
    rangeMeta: "淘宝评分基于成交、有效观看、收藏加购与新客效率",
  },
};

const state = {
  allRows: [],
  rows: [],
  activeSystem: null,
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
  systemFilters: { douyin: null, taobao: null },
};

const elements = {
  home: $("#system-home"),
  workspace: $("#system-workspace"),
  file: $("#excel-file"),
  notice: $("#notice"),
  filters: $("#filters"),
  mode: $("#mode-filter"),
  shop: $("#shop-filter"),
  account: $("#account-filter"),
  material: $("#material-filter"),
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

function rememberFilters() {
  if (!state.activeSystem) return;
  state.systemFilters[state.activeSystem] = {
    mode: elements.mode.value,
    shop: elements.shop.value,
    account: elements.account.value,
    material: elements.material.value,
    start: elements.start.value,
    end: elements.end.value,
  };
}

function restoreFilters(resetDates) {
  const saved = state.systemFilters[state.activeSystem];
  if (!saved || resetDates) return;
  for (const [key, value] of Object.entries(saved)) {
    if (!elements[key]) continue;
    if (elements[key].tagName === "SELECT") {
      if ([...elements[key].options].some((option) => option.value === value)) elements[key].value = value;
    } else {
      elements[key].value = value;
    }
  }
}

function configureWorkspace() {
  const config = SYSTEMS[state.activeSystem];
  $("#system-eyebrow").textContent = config.eyebrow;
  $("#system-title").textContent = config.title;
  $("#system-description").textContent = config.description;
  $("#upload-label").textContent = config.upload;
  $("#empty-title").textContent = config.emptyTitle;
  $("#empty-copy").textContent = config.emptyCopy;
  $("#account-label").textContent = config.accountLabel;
  $("#material-label").textContent = config.materialLabel;
  $("#range-meta").textContent = config.rangeMeta;
  $("#mode-field").hidden = state.activeSystem === "taobao";
  $("#shop-field").hidden = state.activeSystem === "taobao";
  elements.workspace.dataset.system = state.activeSystem;
}

function enterSystem(system) {
  if (!SYSTEMS[system]) return;
  rememberFilters();
  const hasSavedFilters = Boolean(state.systemFilters[system]);
  state.activeSystem = system;
  const platform = SYSTEMS[system].platform;
  state.rows = state.allRows.filter((row) => (row.platform || REPORT_META[row.deliveryMode]?.platform || "抖音") === platform);
  state.analysis = null;
  state.selected = null;
  state.materialTier = "all";
  state.materialQuery = "";
  state.actionType = "全部";
  state.actionQuery = "";
  state.activeTab = "overview";
  state.analysisCache.clear();
  state.renderedTabs.clear();
  configureWorkspace();
  elements.home.hidden = true;
  elements.workspace.hidden = false;
  $$("[data-tab]").forEach((button) => button.classList.toggle("active", button.dataset.tab === "overview"));
  $$(".tab-view").forEach((view) => { view.hidden = view.id !== "view-overview"; });
  updateBounds(!hasSavedFilters);
  restoreFilters(false);
  if (state.rows.length) applyAnalysis();
}

function updateBounds(resetDates = false) {
  const modes = [...new Set(state.rows.map((row) => row.deliveryMode || "推商品").filter(Boolean))].sort();
  const shops = [...new Set(state.rows.map((row) => row.shopCode).filter(Boolean))].sort();
  const accounts = [...new Set(state.rows.map((row) => row.qianchuanId).filter(Boolean))].sort();
  const materials = [...new Map(state.rows.filter((row) => row.materialId).map((row) => [row.materialId, row.name || row.materialId])).entries()];
  const months = state.rows.map((row) => row.month || row.date?.slice(0, 7)).filter(Boolean).sort();
  const minDate = months[0] || "";
  const maxDate = months.at(-1) || "";
  elements.mode.innerHTML = `<option value="all">全部类型</option>${modes.map((mode) => `<option value="${esc(mode)}">${esc(mode)}</option>`).join("")}`;
  elements.shop.innerHTML = `<option value="all">全部店铺</option>${shops.map((shop) => `<option value="${esc(shop)}">${esc(shop)}</option>`).join("")}`;
  elements.account.innerHTML = `<option value="all">全部${state.activeSystem === "taobao" ? "计划" : "账户"}</option>${accounts.map((account) => `<option value="${esc(account)}">${esc(account)}</option>`).join("")}`;
  elements.material.innerHTML = `<option value="all">全部${state.activeSystem === "taobao" ? "主体" : "素材"}</option>${materials.map(([id, name]) => `<option value="${esc(id)}">${esc(id)} · ${esc(name).slice(0, 22)}</option>`).join("")}`;
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
  const material = elements.material.value;
  const start = elements.start.value;
  const end = elements.end.value;
  const cacheKey = [state.activeSystem, mode, shop, account, material, start, end].join("::");
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
      && (material === "all" || row.materialId === material)
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
  rememberFilters();
  const elapsed = Math.round(performance.now() - startedAt);
  showNotice(cached ? `已读取分析缓存，耗时 ${elapsed}ms。` : `分析完成，耗时 ${(elapsed / 1000).toFixed(1)}秒；再次使用相同筛选将直接读取缓存。`);
}

function reportBadge(mode) {
  const meta = REPORT_META[mode] || REPORT_META["推商品"];
  return `<span class="mode-badge" style="--mode:${meta.color}">${esc(mode)}</span>`;
}

const isTmallItem = (item) => item.platform === "淘系" || item.deliveryMode === "淘系短视频" || item.mode === "淘系短视频";
const performanceRoi = (item) => isTmallItem(item) ? item.grossRoi : item.netRoi;
const performanceRoiLabel = (item) => isTmallItem(item) ? "成交ROI" : "净ROI";
const videoQualityLabel = (item) => isTmallItem(item) ? "有效观看率" : "完播率";

function compactDailyRow(row) {
  return {
    key: row.key,
    materialKey: row.materialKey,
    schemaVersion: row.schemaVersion,
    platform: row.platform,
    deliveryMode: row.deliveryMode,
    shopCode: row.shopCode,
    qianchuanId: row.qianchuanId,
    materialId: row.materialId,
    date: row.date,
    spend: row.spend,
    grossOrders: row.grossOrders,
    grossAmount: row.grossAmount,
    netAmount: row.netAmount,
    netOrders: row.netOrders,
    videoPlays: row.videoPlays,
    videoCompletes: row.videoCompletes,
    effectiveViews: row.effectiveViews,
    engagements: row.engagements,
    watchTimeWeighted: row.watchTimeWeighted,
    favoriteCart: row.favoriteCart,
    addCart: row.addCart,
    favorites: row.favorites,
    guidedVisits: row.guidedVisits,
    deepVisits: row.deepVisits,
    newCustomerReach: row.newCustomerReach,
    newCustomerOrders: row.newCustomerOrders,
    newCustomerAmount: row.newCustomerAmount,
  };
}

function renderMetrics() {
  const a = state.analysis;
  const s = a.summary;
  const hasTmall = a.modeSummaries.some((item) => item.mode === "淘系短视频");
  const hasDouyin = a.modeSummaries.some((item) => item.mode !== "淘系短视频");
  const tmallOnly = hasTmall && !hasDouyin;
  const good = a.materials.filter((item) => item.tier === "good");
  const goodSpend = good.reduce((sum, item) => sum + item.spend, 0);
  const goodNet = good.reduce((sum, item) => sum + item.netAmount, 0);
  return `<section class="metric-grid">
    <article><span>投放消耗</span><strong>${money(s.spend)}</strong><small>${a.materials.length} 条素材</small></article>
    <article><span>成交金额 / ROI</span><strong>${money(s.grossAmount)}</strong><small>成交ROI ${s.grossRoi.toFixed(2)}</small></article>
    ${tmallOnly
      ? `<article><span>收藏加购 / 加购率</span><strong class="good-text">${fmt(s.favoriteCart)}</strong><small>收藏加购率 ${pct(s.favoriteCartRate)}</small></article>
         <article><span>引导访问 / 访问率</span><strong>${fmt(s.guidedVisits)}</strong><small>引导访问率 ${pct(s.guidedVisitRate)}</small></article>`
      : `<article><span>${hasTmall ? "综合归因成交" : "净成交额 / ROI"}</span><strong class="good-text">${money(s.netAmount)}</strong><small>${hasTmall ? "淘系按总成交口径" : `净ROI ${s.netRoi.toFixed(2)}`}</small></article>
         <article><span>抖音退款损失</span><strong class="poor-text">${money(s.refundLoss)}</strong><small>${hasDouyin ? `金额损失率 ${pct(s.refundRate)}` : "淘系报表不含退款字段"}</small></article>`}
    <article><span>${tmallOnly ? "成交订单 / 新客订单" : "整体 / 净订单"}</span><strong>${fmt(s.grossOrders)}</strong><small>${tmallOnly ? `新客成交 ${fmt(s.newCustomerOrders)} 笔` : `净成交 ${fmt(s.netOrders)} 单`}</small></article>
    <article><span>达到评分门槛</span><strong>${a.thresholds.eligible}</strong><small>消耗 ≥ ¥50</small></article>
    <article><span>优质素材贡献</span><strong>${a.tiers.good} 条</strong><small>消耗${pct(goodSpend / (s.spend || 1))} · 净成交${pct(goodNet / (s.netAmount || 1))}</small></article>
    <article><span>${tmallOnly ? "观看 / 有效观看率" : "视频播放 / 完播率"}</span><strong>${fmt(s.videoPlays)}</strong><small>${tmallOnly ? `有效观看 ${pct(s.effectiveViewRate)} · 互动 ${pct(s.interactionRate)}` : `完播 ${pct(s.videoCompletionRate)} · 3秒 ${pct(s.threeSecondRate)}`}</small></article>
  </section>`;
}

function monthlyChart(data, tmallOnly = false) {
  if (!data.length) return `<div class="chart-empty">所选范围暂无趋势数据</div>`;
  const width = 760, height = 292, left = 48, right = 44, top = 18, bottom = 34;
  const plotWidth = width - left - right, plotHeight = height - top - bottom;
  const maxSpend = Math.max(...data.map((item) => item.spend), 1);
  const maxRoi = Math.max(...data.map((item) => item.grossRoi), ...data.map((item) => item.netRoi), 1);
  const step = plotWidth / data.length, barWidth = Math.min(34, step * 0.48);
  const points = (key) => data.map((item, index) => `${left + (index + 0.5) * step},${top + plotHeight * (1 - item[key] / maxRoi)}`).join(" ");
  return `<svg class="trend-chart" viewBox="0 0 ${width} ${height}" role="img" aria-label="${tmallOnly ? "月度消耗和成交ROI趋势" : "月度消耗、整体ROI和净ROI趋势"}">
    ${[0, 0.25, 0.5, 0.75, 1].map((t) => `<line class="grid-line" x1="${left}" x2="${width - right}" y1="${top + plotHeight * t}" y2="${top + plotHeight * t}"/><text class="axis-label" x="${left - 7}" y="${top + plotHeight * t + 4}" text-anchor="end">${fmt(maxSpend * (1 - t) / 10000, 1)}万</text>`).join("")}
    ${data.map((item, index) => { const h = plotHeight * item.spend / maxSpend; return `<rect class="spend-bar" x="${left + (index + 0.5) * step - barWidth / 2}" y="${top + plotHeight - h}" width="${barWidth}" height="${h}" rx="3"><title>${esc(item.month)} 消耗${money(item.spend)} 净ROI ${item.netRoi.toFixed(2)}</title></rect>`; }).join("")}
    <polyline points="${points("grossRoi")}" class="${tmallOnly ? "roi-line" : "gross-line"}"/>${tmallOnly ? "" : `<polyline points="${points("netRoi")}" class="roi-line"/>`}
    ${data.map((item, index) => `<text class="axis-label" x="${left + (index + 0.5) * step}" y="${height - 8}" text-anchor="middle">${esc(item.month.slice(5))}</text>`).join("")}
  </svg>`;
}

function renderOverview() {
  const a = state.analysis;
  const s = a.summary;
  const hasTmall = a.modeSummaries.some((item) => item.mode === "淘系短视频");
  const hasDouyin = a.modeSummaries.some((item) => item.mode !== "淘系短视频");
  const tmallOnly = hasTmall && !hasDouyin;
  const tmallSummary = a.modeSummaries.find((item) => item.mode === "淘系短视频") || s;
  const maxCategory = Math.max(...a.categories.map((item) => item.spend), 1);
  $("#view-overview").innerHTML = `${renderMetrics()}
    <section class="insight-grid">
      <article><b>平台独立评分</b><p>${tmallOnly ? "淘系以成交ROI为核心，结合有效观看、收藏加购、访问深度和新客效率。" : "抖音推商品、推直播与淘系短视频分别建模，不共用排名基准。"}</p></article>
      <article><b>${tmallOnly ? "内容到成交" : "退款与成交质量"}</b><p>${tmallOnly ? `有效观看率 ${pct(s.effectiveViewRate)}，收藏加购率 ${pct(s.favoriteCartRate)}，引导访问率 ${pct(s.guidedVisitRate)}。` : `整体ROI ${s.grossRoi.toFixed(2)}，净ROI ${s.netRoi.toFixed(2)}，退款造成 ${money(s.refundLoss)} 成交损失。`}</p></article>
      <article><b>评分可信度</b><p>采用¥${a.thresholds.priorWeight}先验权重收缩小样本ROI；${tmallOnly ? "淘系内容分由有效观看、观看时长、互动和收藏加购组成。" : "抖音内容分纳入完播、3秒和5秒播放质量。"}</p></article>
    </section>
    ${hasTmall ? `<section class="commerce-funnel panel"><div class="panel-head"><div><h2>淘系内容成交漏斗</h2><p>从观看质量到访问、收藏加购与成交</p></div></div><div class="funnel-flow"><article><span>观看量</span><strong>${fmt(tmallSummary.videoPlays)}</strong></article><i>→</i><article><span>有效观看</span><strong>${fmt(tmallSummary.effectiveViews)}</strong><small>${pct(tmallSummary.effectiveViewRate)}</small></article><i>→</i><article><span>引导访问</span><strong>${fmt(tmallSummary.guidedVisits)}</strong><small>${pct(tmallSummary.guidedVisitRate)}</small></article><i>→</i><article><span>收藏加购</span><strong>${fmt(tmallSummary.favoriteCart)}</strong><small>${pct(tmallSummary.favoriteCartRate)}</small></article><i>→</i><article><span>成交订单</span><strong>${fmt(tmallSummary.grossOrders)}</strong><small>ROI ${tmallSummary.grossRoi.toFixed(2)}</small></article></div></section>` : ""}
    <section class="two-column">
      <article class="panel"><div class="panel-head"><div><h2>月度经营趋势</h2><p>${tmallOnly ? "消耗柱 · 成交ROI线" : "消耗柱 · 整体ROI线 · 净ROI线"}</p></div><div class="chart-legend"><i class="spend"></i>消耗${tmallOnly ? `<i class="net"></i>成交ROI` : `<i class="gross"></i>整体ROI<i class="net"></i>净ROI`}</div></div>${monthlyChart(a.monthly, tmallOnly)}</article>
      <article class="panel"><div class="panel-head"><div><h2>等级分布</h2><p>按当前筛选范围重新计算</p></div></div><table><thead><tr><th>等级</th><th>素材数</th><th>占比</th></tr></thead><tbody>${["good", "normal", "poor"].map((tier) => `<tr><td><span class="tier ${tier}">${TIER_LABELS[tier]}</span></td><td>${a.tiers[tier]}</td><td>${pct(a.tiers[tier] / (a.materials.length || 1))}</td></tr>`).join("")}</tbody></table></article>
    </section>
    <section class="panel"><div class="panel-head"><div><h2>投放类型表现</h2><p>各平台分别评分，经营结果在此汇总对比</p></div></div><div class="mode-summary">${a.modeSummaries.map((item) => `<article style="--mode:${REPORT_META[item.mode].color}"><div>${reportBadge(item.mode)}<b>${item.count}条素材</b></div><span>消耗<strong>${money(item.spend)}</strong></span><span>${performanceRoiLabel(item)}<strong>${performanceRoi(item).toFixed(2)}</strong></span><span>${videoQualityLabel(item)}<strong>${pct(isTmallItem(item) ? item.effectiveViewRate : item.videoCompletionRate)}</strong></span><span>有效评分<strong>${item.eligible}</strong></span></article>`).join("")}</div></section>
    <section class="panel"><div class="panel-head"><div><h2>品类经营表现</h2><p>根据素材名称关键词自动归类</p></div></div><div class="category-list">${a.categories.map((item) => `<div class="category-row"><div class="category-name"><b>${esc(item.name)}</b><small>${item.count}条</small></div><div class="bar-track"><i style="width:${item.spend / maxCategory * 100}%"></i></div><span>${money(item.spend)}</span><span>${tmallOnly ? "成交ROI" : "净ROI"} ${(tmallOnly ? item.grossRoi : item.netRoi).toFixed(2)}</span><span>${tmallOnly ? `收藏加购 ${pct(item.favoriteCartRate)}` : `退款 ${pct(item.refundRate)}`}</span></div>`).join("")}</div></section>`;
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
  const tmall = isTmallItem(item);
  const settlementRows = [[7, item.settlement7Roi, item.settlement7Rate], [14, item.settlement14Roi, item.settlement14Rate], [30, item.settlement30Roi, item.settlement30Rate], [90, item.settlement90Roi, item.settlement90Rate]];
  container.innerHTML = `<div class="detail-top"><span class="tier ${item.tier}">${TIER_LABELS[item.tier]} · ${item.score}分</span><span>${reportBadge(item.deliveryMode)} <span class="function-badge" style="--func:${FUNCTION_META[item.functionalType].color}">${esc(item.functionalType)}</span></span></div>
    <h2>${esc(item.name)}</h2><div class="material-id">${tmall ? `计划 ${esc(item.planName || item.qianchuanId || "-")} · 主体ID ${esc(item.id)}${item.itemId ? ` · 宝贝ID ${esc(item.itemId)}` : ""}` : `店铺 ${esc(item.shopCode)} · 千川ID ${esc(item.qianchuanId || "-")} · 素材ID ${esc(item.id)}`}${item.videoType ? ` · ${esc(item.videoType)}` : ""}</div>
    <div class="decision-note"><b>分析理由：</b>${esc(item.reason)}<br><b>建议：</b>${esc(item.action)}</div>
    <div class="detail-metrics">${tmall
      ? `<div><span>消耗</span><strong>${money(item.spend)}</strong></div><div><span>成交ROI</span><strong>${item.grossRoi.toFixed(2)}</strong></div><div><span>成交金额</span><strong>${money(item.grossAmount)}</strong></div><div><span>成交订单</span><strong>${fmt(item.grossOrders)}</strong></div><div><span>有效观看率</span><strong>${pct(item.effectiveViewRate)}</strong></div><div><span>收藏加购率</span><strong>${pct(item.favoriteCartRate)}</strong></div><div><span>引导访问率</span><strong>${pct(item.guidedVisitRate)}</strong></div><div><span>内容质量分</span><strong>${item.contentScore.toFixed(0)}</strong></div>`
      : `<div><span>消耗</span><strong>${money(item.spend)}</strong></div><div><span>净ROI</span><strong>${item.netRoi.toFixed(2)}</strong></div><div><span>净成交</span><strong>${money(item.netAmount)}</strong></div><div><span>退款损失率</span><strong>${pct(item.refund)}</strong></div><div><span>视频播放</span><strong>${fmt(item.videoPlays)}</strong></div><div><span>完播 / 3秒</span><strong>${pct(item.videoCompletionRate)} / ${pct(item.threeSecondRate)}</strong></div><div><span>平均观看</span><strong>${fmt(item.avgWatchTime, 2)}秒</strong></div><div><span>内容质量分</span><strong>${item.contentScore.toFixed(0)}</strong></div>`}</div>
    ${tmall
      ? `<div class="settlement-grid"><div><span>直接成交</span><b>${money(item.directGrossAmount)}</b><small>占比 ${pct(item.directAmountShare)}</small></div><div><span>间接成交</span><b>${money(item.indirectGrossAmount)}</b><small>内容归因延展</small></div><div><span>新客成交</span><b>${fmt(item.newCustomerOrders)} 笔</b><small>${money(item.newCustomerAmount)}</small></div><div><span>平均有效观看</span><b>${fmt(item.avgWatchTime, 1)}秒</b><small>互动率 ${pct(item.interactionRate)}</small></div></div>`
      : `<div class="settlement-grid">${settlementRows.map(([days, roi, rate]) => `<div><span>${days}日结算</span><b>ROI ${roi.toFixed(2)}</b><small>GMV结算率 ${pct(rate)}</small></div>`).join("")}</div>`}
    <div class="panel-head"><div><h2>素材投放趋势</h2><p>日消耗柱 · 7日滚动${tmall ? "成交" : "净"}ROI线 · 顶部生命周期阶段</p></div></div>${lifecycleChart(trend)}`;
}

function renderMaterials() {
  const allItems = filteredMaterials();
  const pages = Math.max(1, Math.ceil(allItems.length / MATERIAL_PAGE_SIZE));
  state.materialPage = Math.min(Math.max(1, state.materialPage), pages);
  const offset = (state.materialPage - 1) * MATERIAL_PAGE_SIZE;
  const items = allItems.slice(offset, offset + MATERIAL_PAGE_SIZE);
  $("#view-materials").innerHTML = `<section class="tier-grid" aria-label="按素材等级筛选">
    <button type="button" data-tier-card="all" class="${state.materialTier === "all" ? "active" : ""}" aria-pressed="${state.materialTier === "all"}"><span>全部素材</span><strong>${state.analysis.materials.length}</strong><small>点击查看全部</small></button>
    ${["good", "normal", "poor"].map((tier) => `<button type="button" data-tier-card="${tier}" class="${state.materialTier === tier ? `active ${tier}` : tier}" aria-pressed="${state.materialTier === tier}"><span>${TIER_LABELS[tier]}素材</span><strong>${state.analysis.tiers[tier]}</strong><small>${pct(state.analysis.tiers[tier] / (state.analysis.materials.length || 1))} · 点击筛选</small></button>`).join("")}
    </section>
    <section class="workspace"><div class="table-panel"><div class="table-head"><div class="toolbar"><div><div class="score-title"><b>素材评分明细</b><span class="score-help"><button type="button" class="score-help-button" aria-describedby="score-definition">评分定义 <i>?</i></button><span id="score-definition" class="score-tooltip" role="tooltip"><strong>素材评分口径（0–100分）</strong><span><b>分级：</b>优质 ≥ 70分；普通 40–69分；劣质 ＜ 40分。</span><span><b>样本门槛：</b>所选周期消耗不足 ¥50 时记为0分，不进入稳定评分。</span><span><b>平台隔离：</b>抖音推商品、推直播和淘系短视频分别在同类素材中排名，筛选范围改变后重新计算。</span><span><b>抖音模型：</b>以贝叶斯整体/净ROI、转化、点击、退款质量和视频留存综合评分。</span><span><b>淘系模型：</b>成交ROI 24%、转化14%、点击8%、净成交规模8%、消耗5%，有效观看12%、平均有效观看6%、互动5%、收藏加购8%、访问5%、新客5%。</span><span><b>内容质量分：</b>抖音使用完播、3秒和5秒留存；淘系使用有效观看40%、平均有效观看20%、互动15%、收藏加购25%。均为同类素材百分位综合分，不等于原始比率。</span></span></span></div><p>点击素材查看投放趋势</p></div><div class="toolbar-group"><label class="field">等级<select id="tier-filter"><option value="all">全部等级</option><option value="good">优质</option><option value="normal">普通</option><option value="poor">劣质</option></select></label><label class="field">素材ID或名称<input id="material-search" placeholder="输入素材ID或名称"></label></div></div></div>
      <div class="table-wrap"><table><thead><tr><th>素材</th><th>投放类型</th><th>功能画像</th><th>评分</th><th>内容分</th><th>消耗</th><th>ROI口径</th><th>质量信号</th></tr></thead><tbody>${items.map((item, index) => `<tr data-material="${index}" class="${state.selected?.id === item.id && state.selected?.shopCode === item.shopCode && state.selected?.deliveryMode === item.deliveryMode && state.selected?.qianchuanId === item.qianchuanId ? "selected" : ""}"><td class="material-name">${esc(item.name)}<small>${esc(item.id)} · ${esc(item.shopCode)} · ${esc(item.qianchuanId || "-")}</small></td><td>${reportBadge(item.deliveryMode)}</td><td><span class="function-badge" style="--func:${FUNCTION_META[item.functionalType].color}">${esc(item.functionalType)}</span></td><td><span class="tier ${item.tier}">${item.score} · ${TIER_LABELS[item.tier]}</span></td><td>${item.contentScore.toFixed(0)}</td><td>${money(item.spend)}</td><td>${performanceRoiLabel(item)} ${performanceRoi(item).toFixed(2)}</td><td>${isTmallItem(item) ? `收藏加购 ${pct(item.favoriteCartRate)}` : `退款 ${pct(item.refund)}`}</td></tr>`).join("")}</tbody></table></div>${paginationMarkup(state.materialPage, allItems.length, MATERIAL_PAGE_SIZE, "materials")}</div>
      <aside id="material-detail" class="detail-panel"></aside></section>`;
  $("#tier-filter").value = state.materialTier;
  $("#material-search").value = state.materialQuery;
  $$("[data-tier-card]").forEach((button) => button.addEventListener("click", () => {
    state.materialTier = button.dataset.tierCard;
    state.materialPage = 1;
    state.selected = null;
    renderMaterials();
  }));
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
  $("#view-functional").innerHTML = `<section class="function-grid">${state.analysis.functions.map((item) => { const meta = FUNCTION_META[item.type] || FUNCTION_META["其他"]; return `<article style="--func:${meta.color}"><div class="function-head"><i></i><h2>${esc(item.type)}</h2><b>${item.count}条</b></div><p>${esc(meta.description)}</p><div class="function-metrics"><span>素材占比<b>${pct(item.share)}</b></span><span>消耗占比<b>${pct(item.spendShare)}</b></span><span>成交ROI<b>${Number(item.grossRoi || 0).toFixed(2)}</b></span><span>成交占比<b>${pct(item.netShare)}</b></span><span>${item.favoriteCart ? "收藏加购率" : "退款损失率"}<b>${pct(item.favoriteCart ? item.favoriteCartRate : item.refundRate)}</b></span></div></article>`; }).join("")}</section>
    <section class="panel"><div class="panel-head"><div><h2>功能画像判定逻辑</h2><p>各平台阈值分别计算，画像名称保持统一</p></div></div><p>爆款型：点击、转化和ROI同时较高；精品型：低消耗、高ROI；引流型：点击或访问较高但转化偏低；损耗型：消耗较高但ROI偏低。抖音可根据退款识别误导型；淘系报表不含退款字段，因此不强行判断误导型。</p></section>`;
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
    <div class="table-wrap"><table><thead><tr><th>素材名称</th><th>素材ID</th><th>投放类型</th><th>功能画像</th><th>品类</th><th>消耗</th><th>ROI口径</th><th>质量信号</th><th>建议动作</th></tr></thead><tbody>${items.map((item) => `<tr><td class="material-name">${esc(item.name)}<small>${esc(item.qianchuanId || "-")}</small></td><td>${esc(item.id)}</td><td>${reportBadge(item.deliveryMode)}</td><td><span class="function-badge" style="--func:${FUNCTION_META[item.functionalType].color}">${esc(item.functionalType)}</span></td><td>${esc(item.category)}</td><td>${money(item.spend)}</td><td>${performanceRoiLabel(item)} ${performanceRoi(item).toFixed(2)}</td><td>${isTmallItem(item) ? `收藏加购 ${pct(item.favoriteCartRate)}` : `退款 ${pct(item.refund)}`}</td><td class="action-copy">${esc(item.action)}</td></tr>`).join("")}</tbody></table></div>${paginationMarkup(state.actionPage, allItems.length, ACTION_PAGE_SIZE, "actions")}</section>`;
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
  try {
    renderers[state.activeTab]?.();
    state.renderedTabs.add(state.activeTab);
  } catch (error) {
    const container = $(`#view-${state.activeTab}`);
    if (container) container.innerHTML = `<section class="panel render-error"><h2>该模块暂时无法展示</h2><p>数据仍然保留，请刷新页面后重试。</p></section>`;
    console.error(`模块渲染失败：${state.activeTab}`, error);
  }
}

async function handleUpload(file) {
  if (!file) return;
  const expectedPlatform = SYSTEMS[state.activeSystem].platform;
  showNotice("正在解析 Excel，请稍候…");
  try {
    let rows = [];
    let monthlyRows = [];
    let counts = {};
    if (file.size >= 50 * 1024 * 1024) {
      ({ monthlyRows, counts } = await readLargeXlsx(file, {
        collectRows: false,
        allowedPlatform: expectedPlatform,
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
        const reportPlatform = reportType === "淘系短视频" ? "淘系" : "抖音";
        if (reportPlatform !== expectedPlatform) {
          throw new Error(`当前是${state.activeSystem === "taobao" ? "淘宝" : "抖音"}系统，请上传对应平台的报表`);
        }
        const missing = validateReportHeaders(headers, reportType);
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
    state.allRows = normalizeStoredRows(await readMonthlyRows());
    state.rows = state.allRows.filter((row) => (row.platform || REPORT_META[row.deliveryMode]?.platform || "抖音") === expectedPlatform);
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
  const systemName = state.activeSystem === "taobao" ? "淘宝" : "抖音";
  if (!state.rows.length || !confirm(`确定只清空${systemName}系统保存的素材数据吗？另一系统的数据不会受影响。`)) return;
  await clearPlatformRows(SYSTEMS[state.activeSystem].platform);
  state.allRows = state.allRows.filter((row) => (row.platform || REPORT_META[row.deliveryMode]?.platform || "抖音") !== SYSTEMS[state.activeSystem].platform);
  state.rows = [];
  state.analysis = null;
  state.selected = null;
  state.analysisCache.clear();
  state.renderedTabs.clear();
  updateBounds(true);
  showNotice(`${systemName}系统的本地数据已清空，另一系统数据保持不变。`);
});
$$("[data-enter-system]").forEach((button) => button.addEventListener("click", () => enterSystem(button.dataset.enterSystem)));
$("#back-home").addEventListener("click", () => {
  rememberFilters();
  elements.workspace.hidden = true;
  elements.home.hidden = false;
});
$$('[data-tab]').forEach((button) => button.addEventListener("click", () => {
  state.activeTab = button.dataset.tab;
  $$('[data-tab]').forEach((item) => item.classList.toggle("active", item === button));
  $$(".tab-view").forEach((view) => { view.hidden = view.id !== `view-${button.dataset.tab}`; });
  renderActiveTab();
}));

try {
  state.allRows = normalizeStoredRows(await readMonthlyRows());
  if (!state.allRows.length) {
    const legacyRows = normalizeStoredRows(await readRows());
    if (legacyRows.length) {
      state.allRows = aggregateMonthlyRows(legacyRows);
      await saveMonthlyRows(state.allRows);
    }
  }
} catch (error) {
  showNotice("无法读取浏览器本地数据，请确认未禁用 IndexedDB。", true);
}
