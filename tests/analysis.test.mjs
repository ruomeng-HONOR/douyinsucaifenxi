import assert from "node:assert/strict";
import test from "node:test";
import { analyzeRows, buildTrend, detectReportType, normalizeExcelRow, validateReportHeaders } from "../analysis.js";

const source = [
  { "店铺编码": "A01", "素材ID": "V1", "素材视频名称": "毛绒小猪", "日期": "2026-07-01", "整体展示次数": 10000, "整体点击次数": 500, "整体消耗": 300, "整体成交订单数": 40, "整体成交金额": 1800, "净成交金额": 1500, "净成交订单数": 35, "1小时内退款率": "5%" },
  { "店铺编码": "A01", "素材ID": "V1", "素材视频名称": "毛绒小猪", "日期": "2026-07-02", "整体展示次数": 9000, "整体点击次数": 450, "整体消耗": 280, "整体成交订单数": 35, "整体成交金额": 1500, "净成交金额": 1300, "净成交订单数": 31, "1小时内退款率": "4%" },
  { "店铺编码": "B02", "素材ID": "V1", "素材视频名称": "同ID不同店铺", "日期": "2026-07-01", "整体展示次数": 5000, "整体点击次数": 100, "整体消耗": 120, "整体成交订单数": 1, "整体成交金额": 60, "净成交金额": 40, "净成交订单数": 1, "1小时内退款率": "20%" },
  { "店铺编码": "A01", "素材ID": "V2", "素材视频名称": "旅行随手杯", "日期": "2026-07-01", "整体展示次数": 4000, "整体点击次数": 300, "整体消耗": 90, "整体成交订单数": 0, "整体成交金额": 0, "净成交金额": 0, "净成交订单数": 0, "1小时内退款率": 0 },
];

const rows = source.map(normalizeExcelRow);

test("normalizes Excel fields", () => {
  assert.equal(rows[0].shopCode, "A01");
  assert.equal(rows[0].date, "2026-07-01");
  assert.equal(rows[0].refundWeighted, 90);
});

test("keeps identical material IDs separate by shop", () => {
  const analysis = analyzeRows(rows);
  assert.equal(analysis.materials.length, 3);
  assert.equal(analysis.materials.filter((item) => item.id === "V1").length, 2);
  assert.equal(analysis.summary.spend, 790);
});

test("builds trend for only the selected shop and material", () => {
  const selected = analyzeRows(rows).materials.find((item) => item.shopCode === "A01" && item.id === "V1");
  const trend = buildTrend(rows, selected);
  assert.equal(trend.length, 2);
  assert.equal(trend[0].spend, 300);
  assert.equal(trend[0].stage, "测试期");
});

test("detects and normalizes product and live report variants", () => {
  const productHeaders = ["店铺", "素材ID", "日期", "整体展示次数", "整体点击次数", "整体消耗", "整体成交金额", "净成交金额", "净成交订单数"];
  const liveHeaders = ["店铺", "素材ID", "日期", "整体展现次数", "整体点击次数", "整体消耗", "整体成交金额", "净成交金额", "净成交订单数", "全域素材视频类型", "基础消耗"];
  assert.equal(detectReportType(productHeaders, "推商品"), "推商品");
  assert.equal(detectReportType(liveHeaders, "推直播"), "推直播");
  assert.deepEqual(validateReportHeaders(productHeaders), []);
  assert.deepEqual(validateReportHeaders(liveHeaders), []);

  const live = normalizeExcelRow({
    "店铺": "013",
    "千川ID": "1698082011424845",
    "素材ID": "7595518584757682230",
    "素材视频名称": "直播素材",
    "日期": "2026-01-23",
    "整体展现次数": "1,000",
    "整体点击次数": "50",
    "整体消耗": "100",
    "整体成交订单数": "2",
    "整体成交金额": "300",
    "净成交金额": "250",
    "净成交订单数": "2",
    "视频播放数": "800",
    "视频完播数": "80",
    "3秒播放率": "35%",
  }, "推直播");
  assert.equal(live.deliveryMode, "推直播");
  assert.equal(live.shopCode, "013");
  assert.equal(live.impressions, 1000);
  assert.equal(live.videoCompletes, 80);
  assert.equal(live.threeSecondPlays, 280);
});

test("scores product and live materials in separate cohorts", () => {
  const product = normalizeExcelRow({ "店铺": "013", "素材ID": "P1", "素材视频名称": "商品素材", "日期": "2026-01-23", "整体展示次数": 1000, "整体点击次数": 100, "整体消耗": 100, "整体成交订单数": 5, "整体成交金额": 500, "净成交金额": 400, "净成交订单数": 4, "视频播放数": 900, "视频完播率": "10%", "3秒播放率": "40%", "5秒播放率": "20%" }, "推商品");
  const live = normalizeExcelRow({ "店铺": "013", "素材ID": "L1", "素材视频名称": "直播素材", "日期": "2026-01-23", "整体展现次数": 1000, "整体点击次数": 100, "整体消耗": 100, "整体成交订单数": 5, "整体成交金额": 500, "净成交金额": 400, "净成交订单数": 4, "视频播放数": 900, "视频完播数": 90, "3秒播放率": "40%", "5秒播放率": "20%" }, "推直播");
  const analysis = analyzeRows([product, live]);
  assert.equal(analysis.materials.length, 2);
  assert.equal(analysis.thresholds.byMode["推商品"].eligible, 1);
  assert.equal(analysis.thresholds.byMode["推直播"].eligible, 1);
  assert.equal(analysis.modeSummaries.length, 2);
});

test("detects, normalizes and scores Wanxiangtai short-video rows", () => {
  const headers = ["日期", "计划ID", "计划名字", "主体ID", "主体类型", "主体名称", "展现量", "点击量", "花费", "观看量", "有效观看量", "平均有效观看时长", "互动量", "总成交金额", "总成交笔数", "宝贝收藏加购数", "引导访问量", "新客成交笔数", "新客成交金额", "宝贝ID"];
  assert.equal(detectReportType(headers, "Sheet1"), "淘系短视频");
  assert.deepEqual(validateReportHeaders(headers, "淘系短视频"), []);

  const row = normalizeExcelRow({
    "日期": "2026-07-26",
    "计划ID": "14670816987",
    "计划名字": "出行个护_干发喷雾",
    "主体ID": "1278019715722414",
    "主体类型": "短视频",
    "主体名称": "油头救星，一喷蓬松",
    "展现量": 1000,
    "点击量": 80,
    "花费": 100,
    "观看量": 950,
    "有效观看量": 190,
    "平均有效观看时长": 20,
    "互动量": 10,
    "总成交金额": 500,
    "总成交笔数": 12,
    "直接成交金额": 400,
    "间接成交金额": 100,
    "宝贝收藏加购数": 20,
    "宝贝加购数": 15,
    "引导访问量": 300,
    "新客触达数": 60,
    "新客成交笔数": 8,
    "新客成交金额": 360,
    "宝贝ID": "741040714529",
  }, "淘系短视频");

  assert.equal(row.platform, "淘系");
  assert.equal(row.deliveryMode, "淘系短视频");
  assert.equal(row.materialId, "1278019715722414");
  assert.equal(row.netAmount, 500);
  assert.equal(row.effectiveViews, 190);
  assert.equal(row.favoriteCart, 20);

  const analysis = analyzeRows([row]);
  const material = analysis.materials[0];
  assert.equal(material.platform, "淘系");
  assert.equal(material.grossRoi, 5);
  assert.equal(material.effectiveViewRate, 0.2);
  assert.equal(material.favoriteCartRate, 0.25);
  assert.equal(material.contentScore, 100);
  assert.equal(analysis.modeSummaries[0].mode, "淘系短视频");
  const functional = analysis.functions.find((item) => item.type === material.functionalType);
  assert.equal(functional.grossRoi, 5);
  assert.equal(functional.favoriteCart, 20);
  assert.equal(functional.favoriteCartRate, 0.25);
});
