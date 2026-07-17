import assert from "node:assert/strict";
import test from "node:test";
import { analyzeRows, buildTrend, normalizeExcelRow } from "../analysis.js";

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
