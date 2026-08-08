/*
 * Holiday Records — in-browser comparison core.
 *
 * Reads an .xlsx entirely client-side (no dependencies, no network): unzips the
 * workbook with the browser's built-in DecompressionStream and parses the XML.
 * Each sheet/tab is one system; the tab name IS the system name (an in-sheet
 * "System" column is ignored). Compares every system against a chosen source.
 *
 * Runs in the browser and in Node (for tests) — uses only TextDecoder,
 * DecompressionStream, DataView and RegExp; no DOM APIs.
 */
(function (root) {
  "use strict";

  const CURRENCY_HEADERS = new Set([
    "currency", "ccy", "curr", "currency code", "holiday ccy", "holiday currency",
  ]);
  const DATE_HEADERS = new Set([
    "date", "dates", "holiday", "holiday date", "holidays", "holiday dates", "value",
  ]);
  const BUILTIN_DATE_FMT_IDS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47]);
  const CURRENCY_RE = /^[A-Za-z]{3}$/;
  const FALLBACK_FORMATS = ["MM/DD/YYYY", "DD/MM/YYYY", "YYYY-MM-DD", "MM-DD-YYYY"];

  // ---------- ZIP + inflate ----------

  async function inflateRaw(bytes) {
    const ds = new DecompressionStream("deflate-raw");
    const writer = ds.writable.getWriter();
    writer.write(bytes);
    writer.close();
    const reader = ds.readable.getReader();
    const chunks = [];
    let total = 0;
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      chunks.push(value);
      total += value.length;
    }
    const out = new Uint8Array(total);
    let o = 0;
    for (const c of chunks) { out.set(c, o); o += c.length; }
    return out;
  }

  // Parse a ZIP via its central directory; returns {name: Uint8Array (inflated)}.
  async function readZip(buf) {
    const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
    let eocd = -1;
    for (let i = buf.length - 22; i >= 0; i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error("Not a valid .xlsx file (no ZIP end record found).");
    const cdCount = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    const decoder = new TextDecoder("utf-8");
    const files = {};
    for (let n = 0; n < cdCount; n++) {
      if (dv.getUint32(p, true) !== 0x02014b50) break;
      const method = dv.getUint16(p + 10, true);
      const compSize = dv.getUint32(p + 20, true);
      const fnLen = dv.getUint16(p + 28, true);
      const extraLen = dv.getUint16(p + 30, true);
      const commentLen = dv.getUint16(p + 32, true);
      const localOffset = dv.getUint32(p + 42, true);
      const name = decoder.decode(buf.subarray(p + 46, p + 46 + fnLen));
      const lhFnLen = dv.getUint16(localOffset + 26, true);
      const lhExtraLen = dv.getUint16(localOffset + 28, true);
      const dataStart = localOffset + 30 + lhFnLen + lhExtraLen;
      const comp = buf.subarray(dataStart, dataStart + compSize);
      files[name] = method === 0 ? comp : await inflateRaw(comp);
      p += 46 + fnLen + extraLen + commentLen;
    }
    return files;
  }

  function text(u8) {
    return u8 ? new TextDecoder("utf-8").decode(u8) : "";
  }

  // ---------- XML helpers ----------

  function unescapeXml(s) {
    return s.replace(/&lt;/g, "<").replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
      .replace(/&#(\d+);/g, (_, d) => String.fromCharCode(parseInt(d, 10)))
      .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCharCode(parseInt(h, 16)))
      .replace(/&amp;/g, "&");
  }

  function attr(tag, name) {
    const m = new RegExp('\\b' + name + '="([^"]*)"').exec(tag);
    return m ? m[1] : null;
  }

  function parseSharedStrings(xml) {
    const items = [];
    if (!xml) return items;
    const siRe = /<si\b[^>]*>([\s\S]*?)<\/si>/g;
    let m;
    while ((m = siRe.exec(xml))) {
      const tRe = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
      let s = "", tm;
      while ((tm = tRe.exec(m[1]))) s += unescapeXml(tm[1]);
      items.push(s);
    }
    return items;
  }

  // Returns a Set of cellXfs indices whose number format is a date format.
  function parseDateStyles(xml) {
    const dateStyles = new Set();
    if (!xml) return dateStyles;
    const customDateIds = new Set();
    const nfRe = /<numFmt\b[^>]*\/>/g;
    let m;
    while ((m = nfRe.exec(xml))) {
      const id = parseInt(attr(m[0], "numFmtId"), 10);
      const code = (attr(m[0], "formatCode") || "").toLowerCase();
      // A date format contains day/year tokens (avoid matching plain numbers).
      if (/[dy]/.test(code) && !/[#0]\s*[eE]/.test(code)) customDateIds.add(id);
    }
    const block = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml);
    if (!block) return dateStyles;
    const xfRe = /<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g;
    let idx = 0;
    while ((m = xfRe.exec(block[1]))) {
      const id = parseInt(attr(m[0], "numFmtId") || "0", 10);
      if (BUILTIN_DATE_FMT_IDS.has(id) || customDateIds.has(id)) dateStyles.add(idx);
      idx++;
    }
    return dateStyles;
  }

  function parseWorkbookSheets(files) {
    const wbXml = text(files["xl/workbook.xml"]);
    const relsXml = text(files["xl/_rels/workbook.xml.rels"]);
    const rels = {};
    let m;
    const relRe = /<Relationship\b[^>]*\/>/g;
    while ((m = relRe.exec(relsXml))) {
      const id = attr(m[0], "Id");
      let target = attr(m[0], "Target");
      if (id && target) {
        if (!target.startsWith("/")) target = "xl/" + target.replace(/^\.\//, "");
        else target = target.slice(1);
        rels[id] = target;
      }
    }
    const sheets = [];
    const sheetRe = /<sheet\b[^>]*\/>/g;
    while ((m = sheetRe.exec(wbXml))) {
      const name = unescapeXml(attr(m[0], "name") || "");
      const rid = attr(m[0], "r:id") || attr(m[0], "id");
      const path = rels[rid];
      if (name && path && files[path]) sheets.push({ name, path });
    }
    return sheets;
  }

  // ---------- cell parsing ----------

  function colIndex(ref) {
    const m = /^([A-Z]+)\d+$/.exec(ref);
    if (!m) return -1;
    let n = 0;
    for (const ch of m[1]) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }
  function rowNumber(ref) {
    const m = /(\d+)$/.exec(ref);
    return m ? parseInt(m[1], 10) : -1;
  }

  function serialToYMD(serial) {
    const days = Math.floor(serial);
    const d = new Date(Date.UTC(1899, 11, 30) + days * 86400000);
    return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate() };
  }

  // Parse one worksheet into rows: array indexed by row -> {colIdx: cellObj}.
  function parseSheetRows(xml, shared, dateStyles) {
    const rows = [];
    const cRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
    let m;
    while ((m = cRe.exec(xml))) {
      const attrs = m[1];
      const inner = m[2] || "";
      const ref = attr("<c " + attrs + ">", "r");
      if (!ref) continue;
      const ci = colIndex(ref);
      const ri = rowNumber(ref);
      if (ci < 0 || ri < 0) continue;
      const t = attr("<c " + attrs + ">", "t");
      const sAttr = attr("<c " + attrs + ">", "s");
      const style = sAttr != null ? parseInt(sAttr, 10) : -1;
      const vMatch = /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(inner);
      const vRaw = vMatch ? vMatch[1] : "";
      let cell;
      if (t === "s") {
        cell = { type: "s", text: shared[parseInt(vRaw, 10)] || "" };
      } else if (t === "inlineStr") {
        let s = "", tm;
        const tRe = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
        while ((tm = tRe.exec(inner))) s += unescapeXml(tm[1]);
        cell = { type: "s", text: s };
      } else if (t === "str") {
        cell = { type: "s", text: unescapeXml(vRaw) };
      } else if (t === "b") {
        cell = { type: "b", text: vRaw };
      } else if (vRaw !== "") {
        const num = parseFloat(vRaw);
        if (style >= 0 && dateStyles.has(style)) {
          cell = { type: "d", ymd: serialToYMD(num) };
        } else {
          cell = { type: "n", text: vRaw };
        }
      } else {
        continue;
      }
      if (!rows[ri]) rows[ri] = {};
      rows[ri][ci] = cell;
    }
    return rows;
  }

  // ---------- date format tokens ----------

  function parseTextDate(str, fmt) {
    const tokens = [];
    let pattern = "";
    const re = /(YYYY|YY|MM|DD)|([^A-Za-z]+)/g;
    let m;
    while ((m = re.exec(fmt))) {
      if (m[1]) { tokens.push(m[1]); pattern += "(\\d{1,4})"; }
      else pattern += m[2].replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    }
    const rx = new RegExp("^" + pattern + "$");
    const mm = rx.exec(String(str).trim());
    if (!mm) return null;
    let y, mo, d;
    tokens.forEach((tk, i) => {
      const v = parseInt(mm[i + 1], 10);
      if (tk[0] === "Y") y = tk === "YY" ? 2000 + v : v;
      else if (tk === "MM") mo = v;
      else d = v;
    });
    if (!y || !mo || !d || mo > 12 || d > 31) return null;
    return { y, m: mo, d };
  }

  function resolveDate(cell, fmt) {
    if (!cell) return null;
    if (cell.type === "d") return cell.ymd;
    if (cell.type === "s") {
      for (const f of [fmt, ...FALLBACK_FORMATS]) {
        const r = parseTextDate(cell.text, f);
        if (r) return r;
      }
    }
    return null;
  }

  function cellText(cell) {
    return cell && cell.text != null ? String(cell.text).trim() : "";
  }
  function normCurrency(s) {
    const t = String(s || "").trim();
    return CURRENCY_RE.test(t) ? t.toUpperCase() : null;
  }
  function ymdKey(ymd) {
    return String(ymd.y).padStart(4, "0") + "-" + String(ymd.m).padStart(2, "0") + "-" + String(ymd.d).padStart(2, "0");
  }
  function formatYMD(ymd, fmt) {
    return fmt.replace("YYYY", String(ymd.y).padStart(4, "0"))
      .replace("YY", String(ymd.y % 100).padStart(2, "0"))
      .replace("MM", String(ymd.m).padStart(2, "0"))
      .replace("DD", String(ymd.d).padStart(2, "0"));
  }

  // Extract {key -> {cur, ymd}} pairs from one sheet's rows (LONG or WIDE).
  function extractPairs(rows, fmt) {
    const pairs = new Map();
    const maxRow = rows.length;
    // LONG: find header row with currency + date headers.
    let curCol = null, dateCol = null, headerRow = -1;
    for (let r = 0; r < Math.min(maxRow, 6); r++) {
      const row = rows[r];
      if (!row) continue;
      let cc = null, dc = null;
      for (const k of Object.keys(row)) {
        const label = cellText(row[k]).toLowerCase();
        if (CURRENCY_HEADERS.has(label) && cc === null) cc = parseInt(k, 10);
        else if (DATE_HEADERS.has(label) && dc === null) dc = parseInt(k, 10);
      }
      if (cc !== null && dc !== null) { curCol = cc; dateCol = dc; headerRow = r; break; }
    }
    if (curCol !== null) {
      for (let r = headerRow + 1; r < maxRow; r++) {
        const row = rows[r];
        if (!row) continue;
        const cur = normCurrency(cellText(row[curCol]));
        const date = resolveDate(row[dateCol], fmt);
        if (cur && date) pairs.set(cur + "|" + ymdKey(date), { cur, ymd: date });
      }
      return pairs;
    }
    // WIDE: a row whose cells are mostly 3-letter currency headers.
    for (let r = 0; r < Math.min(maxRow, 6); r++) {
      const row = rows[r];
      if (!row) continue;
      const keys = Object.keys(row);
      const curCols = keys.filter((k) => normCurrency(cellText(row[k])));
      const nonEmpty = keys.filter((k) => cellText(row[k]));
      if (curCols.length >= 1 && curCols.length >= nonEmpty.length / 2) {
        const map = {};
        curCols.forEach((k) => { map[k] = normCurrency(cellText(row[k])); });
        for (let rr = r + 1; rr < maxRow; rr++) {
          const drow = rows[rr];
          if (!drow) continue;
          for (const k of Object.keys(map)) {
            const date = resolveDate(drow[k], fmt);
            if (date) pairs.set(map[k] + "|" + ymdKey(date), { cur: map[k], ymd: date });
          }
        }
        return pairs;
      }
    }
    return pairs;
  }

  // ---------- public API ----------

  async function loadWorkbook(buf, dateFormat) {
    const files = await readZip(buf);
    const shared = parseSharedStrings(text(files["xl/sharedStrings.xml"]));
    const dateStyles = parseDateStyles(text(files["xl/styles.xml"]));
    const sheets = parseWorkbookSheets(files);
    const order = [];
    const systems = new Map();
    for (const sheet of sheets) {
      const rows = parseSheetRows(text(files[sheet.path]), shared, dateStyles);
      systems.set(sheet.name, extractPairs(rows, dateFormat));
      order.push(sheet.name);
    }
    return { order, systems, dateFormat };
  }

  function resolveSource(order, requested) {
    if (order.includes(requested)) return requested;
    const lower = requested.toLowerCase();
    return order.find((n) => n.toLowerCase() === lower) || null;
  }

  function buildReport(wb, source, outFormat) {
    const others = wb.order.filter((s) => s !== source);
    const systemCols = [source, ...others];
    const all = new Map();
    for (const pairs of wb.systems.values()) {
      for (const [k, v] of pairs) all.set(k, v);
    }
    const header = ["Currency", "Date", ...systemCols, "Status", "Missing_In"];
    const keys = [...all.keys()].sort((a, b) => {
      const A = all.get(a), B = all.get(b);
      if (A.cur !== B.cur) return A.cur < B.cur ? -1 : 1;
      return ymdKey(A.ymd) < ymdKey(B.ymd) ? -1 : ymdKey(A.ymd) > ymdKey(B.ymd) ? 1 : 0;
    });
    const rows = [];
    let matches = 0;
    for (const key of keys) {
      const { cur, ymd } = all.get(key);
      const flags = systemCols.map((s) => (wb.systems.get(s).has(key) ? "Y" : "N"));
      const missing = systemCols.filter((s, i) => flags[i] === "N");
      const status = missing.length === 0 ? "MATCH" : "MISMATCH";
      if (status === "MATCH") matches++;
      rows.push([cur, formatYMD(ymd, outFormat), ...flags, status, missing.join("; ")]);
    }
    return {
      header, rows, systemCols,
      summary: { total: rows.length, matches, mismatches: rows.length - matches },
    };
  }

  function toCSV(header, rows) {
    const esc = (v) => {
      const s = String(v);
      return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    };
    return [header, ...rows].map((r) => r.map(esc).join(",")).join("\r\n");
  }

  const api = { loadWorkbook, buildReport, resolveSource, toCSV, formatYMD, parseTextDate };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.HolidayCore = api;
})(typeof window !== "undefined" ? window : globalThis);
