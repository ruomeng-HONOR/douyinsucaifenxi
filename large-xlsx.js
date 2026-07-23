import { detectReportType, normalizeExcelRow, validateReportHeaders } from "./analysis.js";

const decoder = new TextDecoder("utf-8");

function xmlText(value) {
  return String(value || "")
    .replace(/<[^>]+>/g, "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&")
    .replace(/_x000D_/g, "\n");
}

function columnIndex(reference) {
  const letters = String(reference || "").match(/^[A-Z]+/)?.[0] || "A";
  let value = 0;
  for (const letter of letters) value = value * 26 + letter.charCodeAt(0) - 64;
  return value - 1;
}

async function zipEntries(file) {
  const tailSize = Math.min(file.size, 66000);
  const tailOffset = file.size - tailSize;
  const tail = new Uint8Array(await file.slice(tailOffset).arrayBuffer());
  let eocd = -1;
  for (let index = tail.length - 22; index >= 0; index -= 1) {
    if (tail[index] === 0x50 && tail[index + 1] === 0x4b && tail[index + 2] === 0x05 && tail[index + 3] === 0x06) {
      eocd = index;
      break;
    }
  }
  if (eocd < 0) throw new Error("无法读取Excel压缩目录");
  const eocdView = new DataView(tail.buffer, tail.byteOffset + eocd);
  const directorySize = eocdView.getUint32(12, true);
  const directoryOffset = eocdView.getUint32(16, true);
  if (directorySize === 0xffffffff || directoryOffset === 0xffffffff) throw new Error("暂不支持ZIP64格式的Excel文件");
  const bytes = new Uint8Array(await file.slice(directoryOffset, directoryOffset + directorySize).arrayBuffer());
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const entries = new Map();
  let offset = 0;
  while (offset + 46 <= bytes.length && view.getUint32(offset, true) === 0x02014b50) {
    const method = view.getUint16(offset + 10, true);
    const compressedSize = view.getUint32(offset + 20, true);
    const uncompressedSize = view.getUint32(offset + 24, true);
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const localOffset = view.getUint32(offset + 42, true);
    const name = decoder.decode(bytes.subarray(offset + 46, offset + 46 + nameLength)).replace(/^\//, "");
    entries.set(name, { name, method, compressedSize, uncompressedSize, localOffset });
    offset += 46 + nameLength + extraLength + commentLength;
  }
  return entries;
}

async function entryStream(file, entry) {
  if (!entry) throw new Error("Excel内部文件缺失");
  const header = new Uint8Array(await file.slice(entry.localOffset, entry.localOffset + 30).arrayBuffer());
  const view = new DataView(header.buffer, header.byteOffset, header.byteLength);
  if (view.getUint32(0, true) !== 0x04034b50) throw new Error("Excel内部文件头无效");
  const nameLength = view.getUint16(26, true);
  const extraLength = view.getUint16(28, true);
  const start = entry.localOffset + 30 + nameLength + extraLength;
  const compressed = file.slice(start, start + entry.compressedSize).stream();
  if (entry.method === 0) return compressed;
  if (entry.method !== 8 || typeof DecompressionStream === "undefined") throw new Error("当前浏览器不支持大文件流式解压，请更新Chrome或Edge");
  return compressed.pipeThrough(new DecompressionStream("deflate-raw"));
}

async function entryText(file, entries, name) {
  return new Response(await entryStream(file, entries.get(name))).text();
}

function sharedStringValues(xml) {
  const values = [];
  for (const match of xml.matchAll(/<si(?:\s[^>]*)?>([\s\S]*?)<\/si>/g)) {
    const fragments = [...match[1].matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map((part) => xmlText(part[1]));
    values.push(fragments.join(""));
  }
  return values;
}

function workbookSheets(xml, relationships) {
  const targets = new Map();
  for (const match of relationships.matchAll(/<Relationship\b([^>]*)\/?\s*>/g)) {
    const id = match[1].match(/\bId="([^"]+)"/)?.[1];
    const target = match[1].match(/\bTarget="([^"]+)"/)?.[1];
    if (id && target) targets.set(id, target.replace(/^\//, ""));
  }
  const sheets = [];
  for (const match of xml.matchAll(/<sheet\b([^>]*)\/?\s*>/g)) {
    const name = xmlText(match[1].match(/\bname="([^"]+)"/)?.[1]);
    const id = match[1].match(/\br:id="([^"]+)"/)?.[1];
    let target = targets.get(id) || "";
    if (!target.startsWith("xl/")) target = `xl/${target.replace(/^\.\//, "")}`;
    if (name && target) sheets.push({ name, target });
  }
  return sheets;
}

function parseXmlRow(xml, sharedStrings) {
  const cells = [];
  for (const match of xml.matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
    const reference = match[1].match(/\br="([A-Z]+\d+)"/)?.[1];
    const type = match[1].match(/\bt="([^"]+)"/)?.[1] || "n";
    const raw = match[2].match(/<v>([\s\S]*?)<\/v>/)?.[1] ?? match[2].match(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/)?.[1] ?? "";
    let value = xmlText(raw);
    if (type === "s") value = sharedStrings[Number(value)] ?? "";
    else if (type === "n" && value !== "") value = Number(value);
    else if (type === "b") value = value === "1";
    cells[columnIndex(reference)] = value;
  }
  return cells;
}

function addMonthlyRow(map, row) {
  const month = row.date.slice(0, 7);
  const key = `${month}::${row.materialKey}`;
  let current = map.get(key);
  if (!current) {
    current = {
      ...row,
      key: `month::${key}`,
      date: `${month}-01`,
      month,
      isMonthlyAggregate: true,
      activeDates: new Set(),
      sourceRowCount: 0,
    };
    for (const [field, value] of Object.entries(row)) {
      if (typeof value === "number" && field !== "schemaVersion") current[field] = 0;
    }
    map.set(key, current);
  }
  current.activeDates.add(row.date);
  current.sourceRowCount += 1;
  for (const [field, value] of Object.entries(row)) {
    if (typeof value === "number" && field !== "schemaVersion") current[field] += value;
  }
}

function finishMonthlyRows(map) {
  return [...map.values()].map((row) => {
    const result = { ...row, activeDays: row.activeDates.size };
    delete result.activeDates;
    return result;
  });
}

export function aggregateMonthlyRows(rows) {
  const map = new Map();
  rows.forEach((row) => addMonthlyRow(map, row));
  return finishMonthlyRows(map);
}

async function streamSheet(file, entry, sharedStrings, sheetName, onRow, onProgress, maxRows = Infinity) {
  const reader = (await entryStream(file, entry)).pipeThrough(new TextDecoderStream()).getReader();
  let buffer = "";
  let processed = 0;
  let done = false;
  while (!done && processed < maxRows) {
    const part = await reader.read();
    done = part.done;
    buffer += part.value || "";
    while (processed < maxRows) {
      const start = buffer.search(/<row\b/);
      if (start < 0) {
        if (buffer.length > 512) buffer = buffer.slice(-512);
        break;
      }
      const end = buffer.indexOf("</row>", start);
      if (end < 0) {
        if (start > 0) buffer = buffer.slice(start);
        break;
      }
      const rowXml = buffer.slice(start, end + 6);
      buffer = buffer.slice(end + 6);
      processed += 1;
      await onRow(parseXmlRow(rowXml, sharedStrings), processed);
      if (processed % 5000 === 0) onProgress?.(`${sheetName} 已读取 ${processed.toLocaleString("zh-CN")} 行…`);
    }
  }
  await reader.cancel();
}

export async function readLargeXlsx(file, { onProgress, onBatch, collectRows = true, maxRowsPerSheet = Infinity } = {}) {
  onProgress?.("正在读取大型Excel目录…");
  const entries = await zipEntries(file);
  const sharedEntry = entries.get("xl/sharedStrings.xml");
  const sharedStrings = sharedEntry ? sharedStringValues(await entryText(file, entries, "xl/sharedStrings.xml")) : [];
  const workbookXml = await entryText(file, entries, "xl/workbook.xml");
  const relationships = await entryText(file, entries, "xl/_rels/workbook.xml.rels");
  const sheets = workbookSheets(workbookXml, relationships);
  const rows = [];
  const monthlyMap = new Map();
  let batch = [];
  const counts = {};
  for (const sheet of sheets) {
    if (!/推商品|推直播/.test(sheet.name)) continue;
    let headers = [];
    let reportType = "";
    await streamSheet(file, entries.get(sheet.target), sharedStrings, sheet.name, async (cells, rowNumber) => {
      if (rowNumber === 1) {
        headers = cells.map((value) => String(value ?? "").trim());
        reportType = detectReportType(headers, sheet.name);
        const missing = validateReportHeaders(headers);
        if (missing.length) throw new Error(`${sheet.name}缺少字段：${missing.join("、")}`);
        return;
      }
      const raw = {};
      headers.forEach((header, index) => { if (header) raw[header] = cells[index] ?? ""; });
      const normalized = normalizeExcelRow(raw, reportType);
      if (!normalized.isSummary && normalized.materialId && /^\d{4}-\d{2}-\d{2}$/.test(normalized.date)) {
        if (collectRows) rows.push(normalized);
        addMonthlyRow(monthlyMap, normalized);
        if (onBatch) {
          batch.push(normalized);
          if (batch.length >= 2000) {
            await onBatch(batch);
            batch = [];
          }
        }
        counts[reportType] = (counts[reportType] || 0) + 1;
      }
    }, onProgress, maxRowsPerSheet);
    if (onBatch && batch.length) {
      await onBatch(batch);
      batch = [];
    }
  }
  return { rows, monthlyRows: finishMonthlyRows(monthlyMap), counts };
}
