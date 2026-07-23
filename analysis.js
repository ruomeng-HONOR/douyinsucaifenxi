export const TIER_LABELS = { good: "优质", normal: "普通", poor: "劣质" };

export const REPORT_META = {
  "推商品": { color: "#2f6bff", description: "以商品成交效率、净ROI和退款质量为主，兼顾视频留存" },
  "推直播": { color: "#9b5de5", description: "以直播成交效率为主，提高点击与视频留存的评分权重" },
};

export const FUNCTION_META = {
  "爆款型": { color: "#e36d55", description: "点击、转化和净ROI均较优，具备稳定放量基础" },
  "精品型": { color: "#16a394", description: "小预算高回报，适合阶梯式验证放量" },
  "引流型": { color: "#2f6bff", description: "点击能力较强但成交承接不足" },
  "损耗型": { color: "#8a93a3", description: "消耗较高但净ROI偏低，需要控制损耗" },
  "误导型": { color: "#dd5a61", description: "点击较高但退款偏高，应检查内容承诺" },
  "其他": { color: "#b7bfca", description: "样本或特征不足，暂未形成明确功能" },
};

const FIELD_ALIASES = {
  shopCode: ["店铺", "店铺编码"],
  materialId: ["素材ID"],
  date: ["日期"],
  impressions: ["整体展示次数", "整体展现次数"],
  clicks: ["整体点击次数"],
  spend: ["整体消耗"],
  grossAmount: ["整体成交金额"],
  netAmount: ["净成交金额"],
  netOrders: ["净成交订单数"],
};

const SCORE_WEIGHTS = {
  "推商品": { gross: 0.20, net: 0.20, cvr: 0.12, ctr: 0.08, amount: 0.08, spend: 0.07, refund: 0.10, complete: 0.05, threeSecond: 0.05, fiveSecond: 0.05 },
  "推直播": { gross: 0.17, net: 0.18, cvr: 0.10, ctr: 0.10, amount: 0.07, spend: 0.08, refund: 0.10, complete: 0.08, threeSecond: 0.06, fiveSecond: 0.06 },
};

const n = (value) => Number(value) || 0;
const div = (a, b) => (b ? a / b : 0);
const pick = (row, names, fallback = "") => {
  for (const name of names) {
    if (row[name] !== undefined && row[name] !== null && row[name] !== "") return row[name];
  }
  return fallback;
};

export function normalizeDate(value) {
  if (value instanceof Date && !Number.isNaN(value.valueOf())) {
    return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, "0")}-${String(value.getDate()).padStart(2, "0")}`;
  }
  const raw = String(value ?? "").trim();
  const match = raw.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (match) return `${match[1]}-${match[2].padStart(2, "0")}-${match[3].padStart(2, "0")}`;
  if (/^\d{5}(\.\d+)?$/.test(raw)) {
    const epoch = new Date(Date.UTC(1899, 11, 30));
    epoch.setUTCDate(epoch.getUTCDate() + Math.floor(Number(raw)));
    return epoch.toISOString().slice(0, 10);
  }
  return raw;
}

export function numberValue(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const parsed = Number(String(value ?? "").replace(/[%¥￥,，\s]/g, ""));
  return Number.isFinite(parsed) ? parsed : 0;
}

export function percentValue(value) {
  const text = String(value ?? "");
  const parsed = numberValue(value);
  return text.includes("%") || parsed > 1 ? parsed / 100 : parsed;
}

export function detectReportType(headers = [], sheetName = "") {
  const names = new Set(headers);
  if (/直播/.test(sheetName) || names.has("全域素材视频类型") || names.has("基础消耗") || names.has("整体消耗占比")) return "推直播";
  if (/商品/.test(sheetName) || names.has("整体展示次数") || names.has("整体未完结预售订单预估金额")) return "推商品";
  if (names.has("素材ID") && names.has("日期") && (names.has("整体展示次数") || names.has("整体展现次数"))) return "推商品";
  return "";
}

export function validateReportHeaders(headers = []) {
  const names = new Set(headers);
  return Object.entries(FIELD_ALIASES)
    .filter(([, aliases]) => !aliases.some((alias) => names.has(alias)))
    .map(([field]) => ({
      shopCode: "店铺/店铺编码",
      materialId: "素材ID",
      date: "日期",
      impressions: "整体展示次数/整体展现次数",
      clicks: "整体点击次数",
      spend: "整体消耗",
      grossAmount: "整体成交金额",
      netAmount: "净成交金额",
      netOrders: "净成交订单数",
    })[field]);
}

export function normalizeExcelRow(row, reportType = "") {
  const explicitReportType = typeof reportType === "string" ? reportType : "";
  const deliveryMode = explicitReportType || row.__reportType || detectReportType(Object.keys(row), "") || "推商品";
  const shopCode = String(pick(row, FIELD_ALIASES.shopCode)).trim();
  const qianchuanId = String(row["千川ID"] ?? "").trim();
  const materialId = String(row["素材ID"] ?? "").trim();
  const materialName = String(row["素材视频名称"] ?? "未命名素材").trim() || "未命名素材";
  const videoType = String(row["全域素材视频类型"] ?? "").trim();
  const date = normalizeDate(row["日期"]);
  const impressions = numberValue(pick(row, FIELD_ALIASES.impressions));
  const clicks = numberValue(row["整体点击次数"]);
  const spend = numberValue(row["整体消耗"]);
  const grossAmount = numberValue(row["整体成交金额"]);
  const grossOrders = numberValue(row["整体成交订单数"]);
  const netAmount = numberValue(row["净成交金额"]);
  const netOrders = numberValue(row["净成交订单数"]);
  const videoPlays = numberValue(row["视频播放数"]);
  const videoCompletionRate = percentValue(row["视频完播率"]);
  const videoCompletes = numberValue(row["视频完播数"]) || videoPlays * videoCompletionRate;
  const avgWatchTime = numberValue(row["平均观看时长"]);
  const refund1hAmount = numberValue(row["1小时内退款金额"]);
  const refund1hRate = percentValue(row["1小时内退款率"]);
  const isSummary = ["-", "全部"].includes(materialId) || (videoType === "全部" && materialName === "全部");
  const keyParts = ["v2", deliveryMode, shopCode, qianchuanId || "-", materialId, date];
  const materialKey = [deliveryMode, shopCode, qianchuanId || "-", materialId].join("::");
  return {
    schemaVersion: 2,
    key: keyParts.join("::"),
    materialKey,
    deliveryMode,
    shopCode,
    qianchuanId,
    materialId,
    materialName,
    materialCreatedAt: String(row["素材创建时间"] ?? ""),
    videoType,
    date,
    isSummary,
    impressions,
    clicks,
    spend,
    baseSpend: numberValue(row["基础消耗"]),
    grossOrders,
    grossAmount,
    netAmount,
    netOrders,
    userPaidAmount: numberValue(row["用户实际支付金额"]),
    userPaidNetAmount: numberValue(row["用户实际支付净成交金额"]),
    couponAmount: numberValue(row["智能优惠券金额"]),
    subsidyAmount: numberValue(row["电商平台补贴金额"]),
    unrefundedCouponAmount: numberValue(row["智能优惠券未退款金额"]),
    unrefundedSubsidyAmount: numberValue(row["电商平台补贴未退款金额"]),
    presaleOrders: numberValue(row["整体预售订单数"]),
    presaleAmount: numberValue(row["整体预售订单金额"]),
    unfinishedPresaleAmount: numberValue(row["整体未完结预售订单预估金额"]),
    refund1hOrders: numberValue(row["1小时内退款订单数"]),
    refund1hAmount,
    refundWeighted: refund1hAmount || grossAmount * refund1hRate || Math.max(0, grossAmount - netAmount),
    settlement7Amount: numberValue(row["7日结算金额"]),
    settlement7Orders: numberValue(row["7日结算订单数"]),
    settlement14Amount: numberValue(row["14日结算金额"]),
    settlement14Orders: numberValue(row["14日结算订单数"]),
    settlement30Amount: numberValue(row["30日结算金额"]),
    settlement30Orders: numberValue(row["30日结算订单数"]),
    settlement90Amount: numberValue(row["90日结算金额"]),
    settlement90Orders: numberValue(row["90日结算订单数"]),
    videoPlays,
    videoCompletes,
    likes: numberValue(row["视频点赞数"]),
    comments: numberValue(row["视频评论数"]),
    newFollowers: numberValue(row["新增粉丝数"]),
    watchTimeWeighted: videoPlays * avgWatchTime,
    twoSecondPlays: videoPlays * percentValue(row["2秒播放率"]),
    threeSecondPlays: videoPlays * percentValue(row["3秒播放率"]),
    fiveSecondPlays: videoPlays * percentValue(row["5秒播放率"]),
    tenSecondPlays: videoPlays * percentValue(row["10秒播放率"]),
    grossAmountShare: percentValue(row["整体成交金额占比"]),
    spendShare: percentValue(row["整体消耗占比"]),
  };
}

export function normalizeStoredRows(rows) {
  const currentIdentities = new Set(rows.filter((row) => row.schemaVersion === 2).map((row) => `${row.deliveryMode}::${row.shopCode}::${row.materialId}::${row.date}`));
  return rows.filter((row) => row.schemaVersion === 2 || !currentIdentities.has(`${row.deliveryMode || "推商品"}::${row.shopCode}::${row.materialId}::${row.date}`)).map((row) => ({
    deliveryMode: "推商品",
    qianchuanId: "",
    videoType: "",
    ...row,
  }));
}

export function quantile(values, p) {
  const sorted = values.filter(Number.isFinite).sort((a, b) => a - b);
  if (!sorted.length) return 0;
  const index = (sorted.length - 1) * p;
  const low = Math.floor(index);
  const high = Math.ceil(index);
  return sorted[low] + (sorted[high] - sorted[low]) * (index - low);
}

function ranker(values, inverse = false) {
  const sorted = values.filter(Number.isFinite).sort((a, b) => a - b);
  return (value) => {
    let low = 0;
    let high = sorted.length;
    while (low < high) {
      const mid = (low + high) >> 1;
      if (sorted[mid] <= value) low = mid + 1;
      else high = mid;
    }
    const rank = sorted.length ? low / sorted.length : 0;
    return inverse ? 1 - rank : rank;
  };
}

function categoryOf(name) {
  const sets = [
    ["毛绒", /毛绒|公仔|玩偶|小猪|猫|熊|挂件|水豚|史迪仔|阿柴/i],
    ["潮玩", /盲盒|三丽鸥|kitty|萌豆|yoyo|冰箱贴|腕表|玩具/i],
    ["出行个护", /气囊梳|梳子|风扇|随手杯|保温杯|旅行|u型枕|个护/i],
    ["香氛护肤彩妆", /香水|香氛|护肤|彩妆|气垫|口红|唇|面膜|精华/i],
  ];
  const matched = sets.filter(([, re]) => re.test(name)).map(([label]) => label);
  return matched.length > 1 ? "分类冲突" : matched[0] || "其他";
}

const SUM_FIELDS = [
  "impressions", "clicks", "spend", "baseSpend", "grossOrders", "grossAmount", "netAmount", "netOrders",
  "userPaidAmount", "userPaidNetAmount", "couponAmount", "subsidyAmount", "unrefundedCouponAmount", "unrefundedSubsidyAmount",
  "presaleOrders", "presaleAmount", "unfinishedPresaleAmount", "refund1hOrders", "refund1hAmount", "refundWeighted",
  "settlement7Amount", "settlement7Orders", "settlement14Amount", "settlement14Orders", "settlement30Amount", "settlement30Orders",
  "settlement90Amount", "settlement90Orders", "videoPlays", "videoCompletes", "likes", "comments", "newFollowers", "watchTimeWeighted",
  "twoSecondPlays", "threeSecondPlays", "fiveSecondPlays", "tenSecondPlays",
];

function aggregateRows(rows) {
  const map = new Map();
  for (const row of rows) {
    const deliveryMode = row.deliveryMode || "推商品";
    const key = `${deliveryMode}::${row.shopCode}::${row.qianchuanId || "-"}::${row.materialId}`;
    const current = map.get(key) || {
      deliveryMode,
      shopCode: row.shopCode,
      qianchuanId: row.qianchuanId || "",
      id: row.materialId,
      name: row.materialName,
      materialCreatedAt: row.materialCreatedAt || "",
      videoType: row.videoType || "",
      dates: new Set(),
      activeDaysTotal: 0,
      rowCount: 0,
      grossAmountShareTotal: 0,
      spendShareTotal: 0,
      ...Object.fromEntries(SUM_FIELDS.map((field) => [field, 0])),
    };
    current.name = row.materialName || current.name;
    current.videoType = row.videoType || current.videoType;
    current.materialCreatedAt = row.materialCreatedAt || current.materialCreatedAt;
    current.dates.add(row.date);
    current.activeDaysTotal += row.isMonthlyAggregate ? n(row.activeDays) : 0;
    current.rowCount += 1;
    current.grossAmountShareTotal += n(row.grossAmountShare);
    current.spendShareTotal += n(row.spendShare);
    for (const field of SUM_FIELDS) current[field] += n(row[field]);
    map.set(key, current);
  }
  return [...map.values()].map((item) => ({
    ...item,
    days: item.activeDaysTotal || item.dates.size,
    avgGrossAmountShare: div(item.grossAmountShareTotal, item.rowCount),
    avgSpendShare: div(item.spendShareTotal, item.rowCount),
    dates: undefined,
  }));
}

function enrich(item) {
  return {
    ...item,
    grossRoi: div(item.grossAmount, item.spend),
    netRoi: div(item.netAmount, item.spend),
    ctr: div(item.clicks, item.impressions),
    cvr: div(item.grossOrders, item.clicks),
    cpc: div(item.spend, item.clicks),
    cpa: item.netOrders ? item.spend / item.netOrders : 0,
    settlement: div(item.netAmount, item.grossAmount),
    refund: div(Math.max(0, item.grossAmount - item.netAmount), item.grossAmount),
    refund1hRate: div(item.refund1hAmount, item.grossAmount),
    videoCompletionRate: div(item.videoCompletes, item.videoPlays),
    avgWatchTime: div(item.watchTimeWeighted, item.videoPlays),
    twoSecondRate: div(item.twoSecondPlays, item.videoPlays),
    threeSecondRate: div(item.threeSecondPlays, item.videoPlays),
    fiveSecondRate: div(item.fiveSecondPlays, item.videoPlays),
    tenSecondRate: div(item.tenSecondPlays, item.videoPlays),
    settlement7Roi: div(item.settlement7Amount, item.spend),
    settlement14Roi: div(item.settlement14Amount, item.spend),
    settlement30Roi: div(item.settlement30Amount, item.spend),
    settlement90Roi: div(item.settlement90Amount, item.spend),
    settlement7Rate: div(item.settlement7Amount, item.grossAmount),
    settlement14Rate: div(item.settlement14Amount, item.grossAmount),
    settlement30Rate: div(item.settlement30Amount, item.grossAmount),
    settlement90Rate: div(item.settlement90Amount, item.grossAmount),
    category: categoryOf(item.name),
  };
}

function summarize(materials) {
  const total = materials.reduce((acc, item) => {
    for (const field of SUM_FIELDS) acc[field] += n(item[field]);
    return acc;
  }, Object.fromEntries(SUM_FIELDS.map((field) => [field, 0])));
  return {
    ...total,
    grossRoi: div(total.grossAmount, total.spend),
    netRoi: div(total.netAmount, total.spend),
    ctr: div(total.clicks, total.impressions),
    cvr: div(total.grossOrders, total.clicks),
    refundLoss: Math.max(0, total.grossAmount - total.netAmount),
    refundRate: div(total.grossAmount - total.netAmount, total.grossAmount),
    videoCompletionRate: div(total.videoCompletes, total.videoPlays),
    avgWatchTime: div(total.watchTimeWeighted, total.videoPlays),
    threeSecondRate: div(total.threeSecondPlays, total.videoPlays),
    fiveSecondRate: div(total.fiveSecondPlays, total.videoPlays),
  };
}

function classifyFunctional(item, thresholds) {
  if (item.spend < 50) return "其他";
  const highCtr = item.ctr >= thresholds.ctrMedian;
  const highCvr = item.cvr >= thresholds.cvrMedian;
  const highRoi = item.netRoi >= thresholds.roiMedian;
  if (highCtr && item.refund >= thresholds.refundRisk) return "误导型";
  if (highCtr && highCvr && highRoi) return "爆款型";
  if (highCtr && !highCvr) return "引流型";
  if (item.spend >= thresholds.spend75 && !highRoi) return "损耗型";
  if (item.spend <= thresholds.spendMedian && item.netRoi >= thresholds.roi75) return "精品型";
  return "其他";
}

function actionFor(item) {
  const landing = item.deliveryMode === "推直播" ? "直播间承接、开场话术和人货匹配" : "商品承接、利益点和成交路径";
  if (item.grossOrders >= 10 && item.grossAmount >= 500 && item.refund >= 0.2) return `暂停扩量，核对内容承诺与实际成交体验，并复盘${landing}。`;
  if (item.spend >= 50 && item.netOrders === 0) return `停止继续试量；如需重启，先更换前3秒内容并优化${landing}。`;
  if (item.functionalType === "损耗型") return "降低出价或缩小受众，连续观察净ROI；无改善则暂停。";
  if (item.functionalType === "引流型") return `保留引流价值，重点优化${landing}。`;
  if (item.functionalType === "精品型") return "逐级增加预算，验证放量后的净ROI、结算率和退款率是否稳定。";
  if (item.functionalType === "爆款型") return "守住预算并按周监测ROI与视频留存衰退，提前准备同题材替代素材。";
  if (item.functionalType === "误导型") return "检查内容承诺与实际体验落差，降低冲动表达并验证退款是否回落。";
  return "小预算验证，观察净ROI、视频留存与退款稳定性后再决定是否放量。";
}

function reasonFor(item) {
  const reasons = [];
  if (item.spend < 50) reasons.push("消耗不足¥50，暂不进入稳定评分");
  else reasons.push(`${item.deliveryMode}同类贝叶斯净ROI ${item.bayesNetRoi.toFixed(2)}`);
  reasons.push(`内容质量分 ${item.contentScore.toFixed(0)}`);
  if (item.refund >= 0.15) reasons.push(`退款损失率 ${(item.refund * 100).toFixed(1)}%`);
  if (item.netOrders) reasons.push(`净订单成本 ¥${item.cpa.toFixed(2)}`);
  else if (item.spend >= 50) reasons.push("有消耗但无净订单");
  return reasons.join("；");
}

function modelCohort(items, deliveryMode) {
  const eligible = items.filter((item) => item.spend >= 50);
  const priorGross = quantile(eligible.map((item) => item.grossRoi), 0.5);
  const priorNet = quantile(eligible.map((item) => item.netRoi), 0.5);
  const priorWeight = 200;
  const modeled = items.map((item) => ({
    ...item,
    bayesRoi: div(item.grossAmount + priorWeight * priorGross, item.spend + priorWeight),
    bayesNetRoi: div(item.netAmount + priorWeight * priorNet, item.spend + priorWeight),
    confidence: 100 * item.spend / (item.spend + priorWeight),
  }));
  const thresholds = {
    deliveryMode,
    eligible: eligible.length,
    total: modeled.length,
    priorWeight,
    priorGross,
    priorNet,
    ctrMedian: quantile(eligible.map((item) => item.ctr), 0.5),
    cvrMedian: quantile(eligible.map((item) => item.cvr), 0.5),
    roiMedian: quantile(eligible.map((item) => item.netRoi), 0.5),
    roi75: quantile(eligible.map((item) => item.netRoi), 0.75),
    spendMedian: quantile(eligible.map((item) => item.spend), 0.5),
    spend75: quantile(eligible.map((item) => item.spend), 0.75),
    refundRisk: 0.15,
  };
  const source = modeled.filter((item) => item.spend >= 50);
  const ranks = {
    gross: ranker(source.map((item) => item.bayesRoi)),
    net: ranker(source.map((item) => item.bayesNetRoi)),
    cvr: ranker(source.map((item) => item.cvr)),
    ctr: ranker(source.map((item) => item.ctr)),
    amount: ranker(source.map((item) => Math.log1p(item.netAmount))),
    spend: ranker(source.map((item) => Math.log1p(item.spend))),
    refund: ranker(source.map((item) => item.refund), true),
    complete: ranker(source.map((item) => item.videoCompletionRate)),
    threeSecond: ranker(source.map((item) => item.threeSecondRate)),
    fiveSecond: ranker(source.map((item) => item.fiveSecondRate)),
  };
  const weights = SCORE_WEIGHTS[deliveryMode] || SCORE_WEIGHTS["推商品"];
  const materials = modeled.map((item) => {
    const values = {
      gross: ranks.gross(item.bayesRoi),
      net: ranks.net(item.bayesNetRoi),
      cvr: ranks.cvr(item.cvr),
      ctr: ranks.ctr(item.ctr),
      amount: ranks.amount(Math.log1p(item.netAmount)),
      spend: ranks.spend(Math.log1p(item.spend)),
      refund: ranks.refund(item.refund),
      complete: ranks.complete(item.videoCompletionRate),
      threeSecond: ranks.threeSecond(item.threeSecondRate),
      fiveSecond: ranks.fiveSecond(item.fiveSecondRate),
    };
    const score = item.spend < 50 ? 0 : Math.round(100 * Object.entries(weights).reduce((sum, [key, weight]) => sum + weight * values[key], 0));
    const contentWeight = weights.complete + weights.threeSecond + weights.fiveSecond;
    const contentScore = item.spend < 50 ? 0 : 100 * (weights.complete * values.complete + weights.threeSecond * values.threeSecond + weights.fiveSecond * values.fiveSecond) / contentWeight;
    const tier = score >= 70 ? "good" : score >= 40 ? "normal" : "poor";
    const functionalType = classifyFunctional(item, thresholds);
    const result = { ...item, score, contentScore, tier, functionalType };
    return { ...result, reason: reasonFor(result), action: actionFor(result) };
  });
  return { materials, thresholds };
}

export function analyzeRows(rows) {
  const base = aggregateRows(rows).map(enrich);
  const byMode = {};
  const materials = [];
  for (const mode of Object.keys(REPORT_META)) {
    const modeled = modelCohort(base.filter((item) => item.deliveryMode === mode), mode);
    byMode[mode] = modeled.thresholds;
    materials.push(...modeled.materials);
  }
  materials.sort((a, b) => b.score - a.score || b.spend - a.spend);
  const summary = summarize(materials);
  const eligible = materials.filter((item) => item.spend >= 50);
  const thresholds = {
    eligible: eligible.length,
    total: materials.length,
    priorWeight: 200,
    byMode,
  };
  const tiers = {
    good: materials.filter((item) => item.tier === "good").length,
    normal: materials.filter((item) => item.tier === "normal").length,
    poor: materials.filter((item) => item.tier === "poor").length,
  };
  const categories = ["毛绒", "潮玩", "出行个护", "香氛护肤彩妆", "其他", "分类冲突"]
    .map((name) => {
      const items = materials.filter((item) => item.category === name);
      return { name, count: items.length, ...summarize(items) };
    }).filter((item) => item.count).sort((a, b) => b.spend - a.spend);
  const functions = Object.keys(FUNCTION_META).map((type) => {
    const items = eligible.filter((item) => item.functionalType === type);
    const totals = summarize(items);
    return {
      type,
      count: items.length,
      share: div(items.length, eligible.length),
      spendShare: div(totals.spend, summary.spend),
      netShare: div(totals.netAmount, summary.netAmount),
      netRoi: totals.netRoi,
      refundRate: totals.refundRate,
    };
  });
  const modeSummaries = Object.keys(REPORT_META).map((mode) => {
    const items = materials.filter((item) => item.deliveryMode === mode);
    return { mode, count: items.length, eligible: items.filter((item) => item.spend >= 50).length, ...summarize(items) };
  }).filter((item) => item.count);
  const months = new Map();
  for (const row of rows) {
    const month = row.date.slice(0, 7);
    const current = months.get(month) || { month, spend: 0, grossAmount: 0, netAmount: 0, ids: new Set() };
    current.spend += n(row.spend);
    current.grossAmount += n(row.grossAmount);
    current.netAmount += n(row.netAmount);
    current.ids.add(`${row.deliveryMode}::${row.shopCode}::${row.qianchuanId || "-"}::${row.materialId}`);
    months.set(month, current);
  }
  const monthly = [...months.values()].sort((a, b) => a.month.localeCompare(b.month)).map((item) => ({
    month: item.month,
    spend: item.spend,
    grossAmount: item.grossAmount,
    netAmount: item.netAmount,
    grossRoi: div(item.grossAmount, item.spend),
    netRoi: div(item.netAmount, item.spend),
    materials: item.ids.size,
  }));
  return { materials, summary, tiers, thresholds, categories, functions, modeSummaries, monthly };
}

export function buildTrend(rows, material) {
  const daily = new Map();
  for (const row of rows.filter((item) => item.materialId === material.id
    && item.shopCode === material.shopCode
    && (item.qianchuanId || "") === (material.qianchuanId || "")
    && (item.deliveryMode || "推商品") === material.deliveryMode)) {
    const current = daily.get(row.date) || { date: row.date, spend: 0, grossAmount: 0, netAmount: 0, netOrders: 0, videoPlays: 0 };
    current.spend += n(row.spend);
    current.grossAmount += n(row.grossAmount);
    current.netAmount += n(row.netAmount);
    current.netOrders += n(row.netOrders);
    current.videoPlays += n(row.videoPlays);
    daily.set(row.date, current);
  }
  const base = [...daily.values()].sort((a, b) => a.date.localeCompare(b.date));
  const roiMedian = quantile(base.map((item) => div(item.netAmount, item.spend)), 0.5);
  return base.map((item, index) => {
    const win = base.slice(Math.max(0, index - 6), index + 1);
    const prev = base.slice(Math.max(0, index - 13), Math.max(0, index - 6));
    const rollSpend = win.reduce((sum, row) => sum + row.spend, 0) / win.length;
    const rollRoi = div(win.reduce((sum, row) => sum + row.netAmount, 0), win.reduce((sum, row) => sum + row.spend, 0));
    const prevSpend = prev.length ? prev.reduce((sum, row) => sum + row.spend, 0) / prev.length : rollSpend;
    const prevRoi = div(prev.reduce((sum, row) => sum + row.netAmount, 0), prev.reduce((sum, row) => sum + row.spend, 0));
    let stage = "稳定期";
    if (index < 7) stage = "测试期";
    else if (win.every((row) => row.spend === 0)) stage = "停投期";
    else if (rollSpend > prevSpend * 1.25 && rollRoi >= roiMedian) stage = "放量期";
    else if (rollSpend < prevSpend * 0.7 || rollRoi < prevRoi * 0.7) stage = "衰退期";
    else if (rollRoi >= roiMedian) stage = "成熟期";
    return {
      ...item,
      grossRoi: div(item.grossAmount, item.spend),
      netRoi: div(item.netAmount, item.spend),
      rollSpend,
      rollRoi,
      stage,
    };
  });
}
