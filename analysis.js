export const TIER_LABELS = { good: "优质", normal: "普通", poor: "劣质" };

export const FUNCTION_META = {
  "爆款型": { color: "#e36d55", description: "点击、转化和净ROI均较优，具备稳定放量基础" },
  "精品型": { color: "#16a394", description: "小预算高回报，适合阶梯式验证放量" },
  "引流型": { color: "#2f6bff", description: "点击能力较强但成交承接不足" },
  "损耗型": { color: "#8a93a3", description: "消耗较高但净ROI偏低，需要控制损耗" },
  "误导型": { color: "#dd5a61", description: "点击较高但退款偏高，应检查内容承诺" },
  "其他": { color: "#b7bfca", description: "样本或特征不足，暂未形成明确功能" },
};

const n = (value) => Number(value) || 0;
const div = (a, b) => (b ? a / b : 0);

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
  return text.includes("%") ? parsed / 100 : parsed;
}

export function normalizeExcelRow(row) {
  const shopCode = String(row["店铺编码"] ?? "").trim();
  const materialId = String(row["素材ID"] ?? "").trim();
  const date = normalizeDate(row["日期"]);
  const grossAmount = numberValue(row["整体成交金额"]);
  const netAmount = numberValue(row["净成交金额"]);
  const refundRate = percentValue(row["1小时内退款率"]);
  return {
    key: `${shopCode}::${materialId}::${date}`,
    shopCode,
    materialId,
    materialName: String(row["素材视频名称"] ?? "未命名素材").trim() || "未命名素材",
    materialCreatedAt: String(row["素材创建时间"] ?? ""),
    date,
    impressions: numberValue(row["整体展示次数"]),
    clicks: numberValue(row["整体点击次数"]),
    spend: numberValue(row["整体消耗"]),
    grossOrders: numberValue(row["整体成交订单数"]),
    grossAmount,
    netAmount,
    netOrders: numberValue(row["净成交订单数"]),
    refundWeighted: grossAmount * refundRate || Math.max(0, grossAmount - netAmount),
  };
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

function aggregateRows(rows) {
  const map = new Map();
  for (const row of rows) {
    const key = `${row.shopCode}::${row.materialId}`;
    const current = map.get(key) || {
      shopCode: row.shopCode,
      id: row.materialId,
      name: row.materialName,
      dates: new Set(),
      impressions: 0,
      clicks: 0,
      spend: 0,
      grossOrders: 0,
      grossAmount: 0,
      netAmount: 0,
      netOrders: 0,
      refundWeighted: 0,
    };
    current.name = row.materialName || current.name;
    current.dates.add(row.date);
    for (const field of ["impressions", "clicks", "spend", "grossOrders", "grossAmount", "netAmount", "netOrders", "refundWeighted"]) {
      current[field] += n(row[field]);
    }
    map.set(key, current);
  }
  return [...map.values()].map((item) => ({ ...item, days: item.dates.size, dates: undefined }));
}

function summarize(materials) {
  const total = materials.reduce((acc, item) => {
    for (const field of ["spend", "grossAmount", "netAmount", "grossOrders", "netOrders", "impressions", "clicks", "refundWeighted"]) {
      acc[field] += n(item[field]);
    }
    return acc;
  }, { spend: 0, grossAmount: 0, netAmount: 0, grossOrders: 0, netOrders: 0, impressions: 0, clicks: 0, refundWeighted: 0 });
  return {
    ...total,
    grossRoi: div(total.grossAmount, total.spend),
    netRoi: div(total.netAmount, total.spend),
    ctr: div(total.clicks, total.impressions),
    refundLoss: Math.max(0, total.grossAmount - total.netAmount),
    refundRate: div(total.grossAmount - total.netAmount, total.grossAmount),
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
  if (item.grossOrders >= 10 && item.grossAmount >= 500 && item.refund >= 0.2) return "暂停扩量，核对规格、价格、赠品与实物展示，并制作低冲动表达版本对照测试。";
  if (item.spend >= 50 && item.netOrders === 0) return "停止继续试量；如需重启，先更换前3秒内容假设和商品承接方案。";
  if (item.functionalType === "损耗型") return "降低出价或缩小受众，连续观察净ROI；无改善则暂停。";
  if (item.functionalType === "引流型") return "保留引流价值，重点优化商品承接、利益点和成交路径。";
  if (item.functionalType === "精品型") return "逐级增加预算，验证放量后的净ROI和退款率是否稳定。";
  if (item.functionalType === "爆款型") return "守住预算并按周监测衰退信号，提前准备同题材替代素材。";
  if (item.functionalType === "误导型") return "检查内容承诺与实物落差，降低冲动表达并验证退款是否回落。";
  return "小预算验证，观察净ROI与退款率稳定性后再决定是否放量。";
}

function reasonFor(item) {
  const reasons = [];
  if (item.spend < 50) reasons.push("消耗不足¥50，暂不进入稳定评分");
  else reasons.push(`贝叶斯净ROI ${item.bayesNetRoi.toFixed(2)}`);
  if (item.refund >= 0.15) reasons.push(`退款损失率 ${(item.refund * 100).toFixed(1)}%`);
  if (item.netOrders) reasons.push(`净订单成本 ¥${item.cpa.toFixed(2)}`);
  else if (item.spend >= 50) reasons.push("有消耗但无净订单");
  return reasons.join("；");
}

export function analyzeRows(rows) {
  const base = aggregateRows(rows).map((item) => ({
    ...item,
    grossRoi: div(item.grossAmount, item.spend),
    netRoi: div(item.netAmount, item.spend),
    ctr: div(item.clicks, item.impressions),
    cvr: div(item.grossOrders, item.clicks),
    cpc: div(item.spend, item.clicks),
    cpa: item.netOrders ? item.spend / item.netOrders : 0,
    settlement: div(item.netAmount, item.grossAmount),
    refund: div(item.grossAmount - item.netAmount, item.grossAmount),
    category: categoryOf(item.name),
  }));
  const eligible = base.filter((item) => item.spend >= 50);
  const priorGross = quantile(eligible.map((item) => item.grossRoi), 0.5);
  const priorNet = quantile(eligible.map((item) => item.netRoi), 0.5);
  const priorWeight = 200;
  const modeled = base.map((item) => ({
    ...item,
    bayesRoi: div(item.grossAmount + priorWeight * priorGross, item.spend + priorWeight),
    bayesNetRoi: div(item.netAmount + priorWeight * priorNet, item.spend + priorWeight),
    confidence: 100 * item.spend / (item.spend + priorWeight),
  }));
  const thresholds = {
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
  const ranks = {
    gross: ranker(modeled.filter((m) => m.spend >= 50).map((m) => m.bayesRoi)),
    net: ranker(modeled.filter((m) => m.spend >= 50).map((m) => m.bayesNetRoi)),
    cvr: ranker(eligible.map((m) => m.cvr)),
    ctr: ranker(eligible.map((m) => m.ctr)),
    amount: ranker(eligible.map((m) => Math.log1p(m.netAmount))),
    spend: ranker(eligible.map((m) => Math.log1p(m.spend))),
    refund: ranker(eligible.map((m) => m.refund), true),
  };
  const materials = modeled.map((item) => {
    const score = item.spend < 50 ? 0 : Math.round(100 * (
      0.25 * ranks.gross(item.bayesRoi) +
      0.20 * ranks.net(item.bayesNetRoi) +
      0.15 * ranks.cvr(item.cvr) +
      0.10 * ranks.ctr(item.ctr) +
      0.10 * ranks.amount(Math.log1p(item.netAmount)) +
      0.10 * ranks.spend(Math.log1p(item.spend)) +
      0.10 * ranks.refund(item.refund)
    ));
    const tier = score >= 70 ? "good" : score >= 40 ? "normal" : "poor";
    const functionalType = classifyFunctional(item, thresholds);
    const result = { ...item, score, tier, functionalType };
    return { ...result, reason: reasonFor(result), action: actionFor(result) };
  }).sort((a, b) => b.score - a.score || b.spend - a.spend);

  const summary = summarize(materials);
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
    const items = materials.filter((item) => item.spend >= 50 && item.functionalType === type);
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
  const months = new Map();
  for (const row of rows) {
    const month = row.date.slice(0, 7);
    const current = months.get(month) || { month, spend: 0, grossAmount: 0, netAmount: 0, ids: new Set() };
    current.spend += n(row.spend);
    current.grossAmount += n(row.grossAmount);
    current.netAmount += n(row.netAmount);
    current.ids.add(`${row.shopCode}::${row.materialId}`);
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
  return { materials, summary, tiers, thresholds, categories, functions, monthly };
}

export function buildTrend(rows, material) {
  const daily = new Map();
  for (const row of rows.filter((r) => r.materialId === material.id && r.shopCode === material.shopCode)) {
    const current = daily.get(row.date) || { date: row.date, spend: 0, grossAmount: 0, netAmount: 0, netOrders: 0 };
    current.spend += n(row.spend);
    current.grossAmount += n(row.grossAmount);
    current.netAmount += n(row.netAmount);
    current.netOrders += n(row.netOrders);
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
