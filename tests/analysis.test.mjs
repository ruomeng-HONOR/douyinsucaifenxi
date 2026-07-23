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
