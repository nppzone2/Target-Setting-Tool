/* Target Setting Tool – engine chia Target M+1 bottom-up từ Outlet.
   Dùng chung cho trình duyệt (index.html) và kiểm thử bằng Node.
   Developed by Ta Thi Thu Minh */
"use strict";
(function (root) {

/* ============================================================ tiện ích */
const D = "\u0001";
const num = v => { if (v == null || v === "") return 0; if (typeof v === "number") return isFinite(v) ? v : 0; const n = Number(String(v).replace(/,/g, "").trim()); return isFinite(n) ? n : 0; };
const str = v => v == null ? "" : String(v).trim();
const add = (map, k, v) => map.set(k, (map.get(k) || 0) + v);
const nextMonth = ym => { const y = Math.floor(ym / 100), m = ym % 100; return m === 12 ? (y + 1) * 100 + 1 : y * 100 + m + 1; };
const prevMonth = ym => { const y = Math.floor(ym / 100), m = ym % 100; return m === 1 ? (y - 1) * 100 + 12 : y * 100 + m - 1; };
const ymLabel = ym => `T${ym % 100}/${Math.floor(ym / 100)}`;
const distCode = name => str(name).split("-")[0].trim().toUpperCase();

/* ============================================================ đọc ZIP / XLSX (không dùng thư viện ngoài) */
async function inflateRaw(u8) {
  const ds = new DecompressionStream("deflate-raw");
  const out = await new Response(new Blob([u8]).stream().pipeThrough(ds)).arrayBuffer();
  return new Uint8Array(out);
}
async function unzip(buf) {
  const u8 = new Uint8Array(buf), dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  let eocd = -1;
  for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  if (eocd < 0) throw new Error("File không phải định dạng .xlsx hợp lệ.");
  const count = dv.getUint16(eocd + 10, true);
  let p = dv.getUint32(eocd + 16, true);
  const files = {}, dec = new TextDecoder();
  for (let k = 0; k < count; k++) {
    if (dv.getUint32(p, true) !== 0x02014b50) break;
    const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
    const nlen = dv.getUint16(p + 28, true), elen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
    const loc = dv.getUint32(p + 42, true);
    const name = dec.decode(u8.subarray(p + 46, p + 46 + nlen));
    const lnl = dv.getUint16(loc + 26, true), lel = dv.getUint16(loc + 28, true);
    const start = loc + 30 + lnl + lel;
    files[name] = { method, data: u8.subarray(start, start + csize) };
    p += 46 + nlen + elen + clen;
  }
  return {
    text: async n => { const f = files[n]; if (!f) return null; const raw = f.method === 8 ? await inflateRaw(f.data) : f.data; return dec.decode(raw); }
  };
}
const ENT = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" };
const unxml = s => s.indexOf("&") < 0 ? s : s.replace(/&(#x[0-9a-f]+|#\d+|\w+);/gi, (m, e) => e[0] === "#" ? String.fromCodePoint(e[1] === "x" || e[1] === "X" ? parseInt(e.slice(2), 16) : +e.slice(1)) : (ENT[e] ?? m));
const attr = (s, a) => { const m = s.match(new RegExp(`\\b${a}="([^"]*)"`)); return m ? m[1] : null; };
function colIndex(ref) { let n = 0; for (let i = 0; i < ref.length; i++) { const c = ref.charCodeAt(i); if (c >= 65 && c <= 90) n = n * 26 + (c - 64); else break; } return n - 1; }
function textOf(inner) { let s = ""; const re = /<t\b[^>]*>([\s\S]*?)<\/t>/g; let m; while ((m = re.exec(inner))) s += m[1]; return unxml(s); }
// Đọc sheet bằng regex để xử lý nhanh file lớn (trên 100 nghìn dòng)
function parseSheet(xml, ss) {
  const rows = [];
  const reRow = /<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g, reC = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
  let m, rAuto = 0;
  while ((m = reRow.exec(xml))) {
    const rAttr = attr(m[1], "r"), ri = rAttr ? +rAttr - 1 : rAuto; rAuto = ri + 1;
    const arr = [], body = m[2] || "";
    let c, cAuto = 0; reC.lastIndex = 0;
    while ((c = reC.exec(body))) {
      const ref = attr(c[1], "r"), ci = ref ? colIndex(ref) : cAuto; cAuto = ci + 1;
      const t = attr(c[1], "t"), inner = c[2] || "";
      const vm = inner.match(/<v>([\s\S]*?)<\/v>/);
      let v = null;
      if (t === "s") v = vm ? ss[+vm[1]] : null;
      else if (t === "inlineStr") v = textOf(inner);
      else if (t === "str" || t === "e") v = vm ? unxml(vm[1]) : null;
      else if (t === "b") v = vm ? vm[1] === "1" : null;
      else if (vm && vm[1] !== "") v = Number(vm[1]);
      arr[ci] = v;
    }
    rows[ri] = arr;
  }
  for (let i = 0; i < rows.length; i++) if (!rows[i]) rows[i] = [];
  return rows;
}
async function readWorkbook(buf) {
  const z = await unzip(buf);
  const ss = [], sst = await z.text("xl/sharedStrings.xml");
  if (sst) { const re = /<si>([\s\S]*?)<\/si>/g; let m; while ((m = re.exec(sst))) ss.push(textOf(m[1])); }
  const wb = await z.text("xl/workbook.xml"), rels = await z.text("xl/_rels/workbook.xml.rels") || "";
  const relMap = {}; { const re = /<Relationship\b([^>]*)\/?>/g; let m; while ((m = re.exec(rels))) relMap[attr(m[1], "Id")] = attr(m[1], "Target"); }
  const sheets = []; const re = /<sheet\b([^>]*)\/?>/g; let m;
  while ((m = re.exec(wb))) {
    const rid = attr(m[1], "r:id"); let target = relMap[rid] || "";
    target = target.startsWith("/") ? target.slice(1) : "xl/" + target.replace(/^\.\//, "");
    const x = await z.text(target); if (!x) continue;
    sheets.push({ name: unxml(attr(m[1], "name") || ""), rows: parseSheet(x, ss) });
  }
  return sheets;
}

/* ============================================================ đọc dữ liệu nghiệp vụ */
const lc = s => str(s).toLowerCase();
function findCols(rows, must, maxScan = 30) {
  for (let h = 0; h < Math.min(rows.length, maxScan); h++) {
    const n = (rows[h] || []).map(lc);
    if (must.every(alts => alts.some(a => n.includes(a.toLowerCase())))) return { h, n };
  }
  return null;
}
const pick = (n, names) => { for (const x of names) { const i = n.indexOf(x.toLowerCase()); if (i >= 0) return i; } return -1; };
function serialToYM(v) { const d = new Date(Math.round((v - 25569) * 86400000)); return d.getUTCFullYear() * 100 + d.getUTCMonth() + 1; }
function headerYM(v) {
  if (v == null || v === "") return null;
  if (typeof v === "number") {
    if (v > 20000 && v < 80000) return serialToYM(v);
    if (v >= 190001 && v <= 299912 && v % 100 >= 1 && v % 100 <= 12) return v;
    return null;
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{4})[-\/.](\d{1,2})(?:[-\/.]\d{1,2})?(?:[ T].*)?$/); if (m) return +m[1] * 100 + +m[2];
  m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})$/); if (m) return +m[3] * 100 + +m[2];
  m = s.match(/^(\d{1,2})[-\/.](\d{4})$/); if (m) return +m[2] * 100 + +m[1];
  m = s.match(/^(\d{4})(\d{2})$/); if (m && +m[2] >= 1 && +m[2] <= 12) return +s;
  return null;
}

// SI-SO cấp NPP: Region, Area, Distributor_Name, ShortCode, calendar_month, Target (Case), Sales Out (Case)
function parseSI(sheets, label) {
  for (const s of sheets) {
    const f = findCols(s.rows, [["Distributor_Name"], ["ShortCode"], ["calendar_month"], ["Target (Case)"]]); if (!f) continue;
    const { h, n } = f, I = c => n.indexOf(c.toLowerCase());
    const iR = I("Region"), iA = I("Area"), iD = I("Distributor_Name"), iS = I("ShortCode"), iM = I("calendar_month");
    const iT = I("Target (Case)"), iTH = I("Target (Htls)"), iSO = I("Sales Out (Case)"), iSI = I("Sales In (Case)"), iSOH = I("Sales Out (Htls)");
    const out = [];
    for (let i = h + 1; i < s.rows.length; i++) {
      const r = s.rows[i]; const dist = str(r[iD]), sc = str(r[iS]), m = Math.round(num(r[iM]));
      if (!dist || !sc || !m) continue;
      out.push({ region: str(r[iR]), area: str(r[iA]), dist, sc, m, tC: num(r[iT]), tH: iTH >= 0 ? num(r[iTH]) : 0, so: iSO >= 0 ? num(r[iSO]) : 0, soH: iSOH >= 0 ? num(r[iSOH]) : 0, si: iSI >= 0 ? num(r[iSI]) : 0 });
    }
    if (out.length) return out;
  }
  throw new Error(`File SI-SO ${label} cần các cột Region, Area, Distributor_Name, ShortCode, calendar_month, Target (Case), Sales Out (Case).`);
}
// SO Detail by Outlet: Area_Name, Seller_Name, Outlet_ID, Outlet_Name, Outlet_Segment, Calendar_Month, ShortCode, UOM (Case)
function parseOutlet(sheets, label) {
  for (const s of sheets) {
    const f = findCols(s.rows, [["Outlet_ID", "Outlet_Code", "Outlet code"], ["Seller_Name", "Distributor_Name"], ["Calendar_Month"], ["ShortCode", "Search_Text"]]); if (!f) continue;
    const { h, n } = f;
    const iO = pick(n, ["Outlet_ID", "Outlet_Code", "Outlet code"]), iN = pick(n, ["Outlet_Name"]), iSeg = pick(n, ["Outlet_Segment"]);
    const iD = pick(n, ["Seller_Name", "Distributor_Name"]), iA = pick(n, ["Area_Name", "Area"]), iR = pick(n, ["Region_Name", "Region"]);
    const iM = pick(n, ["Calendar_Month"]), iS = pick(n, ["ShortCode", "Search_Text"]);
    let iV = pick(n, ["UOM", "Sales Out (Case)", "Sales_Out", "Value", "Case"]);
    if (iV < 0) iV = n.length - 1;
    const out = [];
    for (let i = h + 1; i < s.rows.length; i++) {
      const r = s.rows[i]; const o = str(r[iO]), d = str(r[iD]), m = Math.round(num(r[iM])), sc = str(r[iS]);
      if (!o || !d || !m || !sc) continue;
      out.push({ oid: o, name: iN >= 0 ? str(r[iN]) : "", seg: iSeg >= 0 ? str(r[iSeg]) : "", dist: d, area: iA >= 0 ? str(r[iA]) : "", region: iR >= 0 ? str(r[iR]) : "", m, sc, v: num(r[iV]) });
    }
    if (out.length) return out;
  }
  throw new Error(`File SO theo Outlet ${label} cần các cột Outlet_ID, Seller_Name, Calendar_Month, ShortCode và cột số (UOM, đơn vị Case).`);
}
// Item Detail: BrandGroup, BrandFamily, ShortCode 2026, PackSize, Cl per Can, Hl
function parseItem(sheets) {
  for (const s of sheets) {
    const f = findCols(s.rows, [["BrandFamily", "Brand Family"], ["ShortCode 2026", "ShortCode"]]); if (!f) continue;
    const { h, n } = f;
    const iG = pick(n, ["BrandGroup", "Brand Group"]), iF = pick(n, ["BrandFamily", "Brand Family"]), iS = pick(n, ["ShortCode 2026", "ShortCode"]);
    const iP = pick(n, ["PackSize"]), iC = pick(n, ["Cl per Can"]), iH = pick(n, ["Hl"]);
    const out = {};
    for (let i = h + 1; i < s.rows.length; i++) {
      const r = s.rows[i]; const sc = str(r[iS]); if (!sc) continue;
      let hl = iH >= 0 ? num(r[iH]) : 0;
      if (!hl && iP >= 0 && iC >= 0) hl = num(r[iP]) * num(r[iC]) / 100000;
      out[sc] = { fam: normFam(str(r[iF])), grp: str(r[iG]) || "Khác", hl };
    }
    return out;
  }
  throw new Error("File Item Detail cần các cột BrandGroup, BrandFamily, ShortCode 2026.");
}
const normFam = f => /^hen?e?i?n?e?ken$|^heneiken$/i.test(f) ? "Heineken" : (f || "Khác");
// Target by Month (Htls): Area × ShortCode, cột tháng
function parseTargetMonth(sheets, label) {
  for (const s of sheets) {
    const f = findCols(s.rows, [["Area"], ["ShortCode"]]); if (!f) continue;
    const { h } = f, hdr = s.rows[h], n = f.n;
    const iA = n.indexOf("area"), iS = n.indexOf("shortcode"), iK = n.indexOf("sku"), iR = n.indexOf("region");
    const months = {};
    hdr.forEach((v, i) => { const ym = headerYM(v); if (ym && ![iA, iS, iK, iR].includes(i)) months[ym] = i; });
    if (!Object.keys(months).length) return null;
    const out = [];
    for (let i = h + 1; i < s.rows.length; i++) {
      const r = s.rows[i]; const area = str(r[iA]), sc = str(r[iS]); if (!area || !sc) continue;
      const v = {}; for (const [ym, ci] of Object.entries(months)) v[ym] = Math.max(num(r[ci]), 0);
      out.push({ region: iR >= 0 ? str(r[iR]) : "", area, sc, sku: iK >= 0 ? str(r[iK]) : "", htls: v });
    }
    return { kind: "month", rows: out, months: Object.keys(months).map(Number).sort((a, b) => a - b) };
  }
  return null;
}
// Target tháng M+1 theo Case: Item_Code × các cột Area
function parseTargetCase(sheets) {
  const skip = /^(sku|region|total|grand total|tổng|tong|item_name|item name|brand|group|group repertoire|brandfamily|brand family)$/i;
  for (const s of sheets) for (let h = 0; h < Math.min(s.rows.length, 30); h++) {
    const n = (s.rows[h] || []).map(str);
    const iS = n.findIndex(v => /^(item_code|item code|shortcode|short code|search_text)$/i.test(v));
    if (iS < 0) continue;
    if (n.some(v => /^area$/i.test(v)) || n.some(v => headerYM(v) != null)) return null;
    const cols = n.map((v, i) => [v, i]).filter(([v, i]) => i !== iS && v && !skip.test(v));
    if (!cols.length) return null;
    const out = [];
    for (let i = h + 1; i < s.rows.length; i++) {
      const r = s.rows[i]; if (!r) continue; const sc = str(r[iS]); if (!sc) continue;
      for (const [a, ci] of cols) out.push({ area: a, sc, cases: Math.max(num(r[ci]), 0) });
    }
    return { kind: "case", rows: out };
  }
  return null;
}
function parseTarget(sheets, label) {
  const t = parseTargetMonth(sheets, label) || parseTargetCase(sheets);
  if (!t) throw new Error(`File Target ${label} không đúng dạng. Cần Target by Month (Area, ShortCode, các cột tháng, Htls) hoặc Target M+1 theo Case (Item_Code và mỗi Area một cột).`);
  return t;
}
// File CP: Distributor_Name, ShortCode, cột số CP (Case)
function parseCP(sheets, label) {
  for (const s of sheets) for (let h = 0; h < Math.min(s.rows.length, 30); h++) {
    const n = (s.rows[h] || []).map(lc);
    const iD = pick(n, ["Distributor_Name", "Distributor", "Customer_Name", "NPP"]), iS = pick(n, ["ShortCode", "Search_Text"]);
    const iV = n.findIndex(v => /\bcp\b/i.test(v));
    if (iD < 0 || iS < 0 || iV < 0) continue;
    const out = [];
    for (let i = h + 1; i < s.rows.length; i++) {
      const r = s.rows[i]; const d = str(r[iD]), sc = str(r[iS]), v = num(r[iV]);
      if (d && sc && v) out.push({ dist: d, sc, v });
    }
    return out;
  }
  throw new Error(`File CP ${label} cần các cột Distributor_Name, ShortCode và cột số CP (Case).`);
}

// Danh mục SKU mặc định (theo file Item Detail), dùng khi không nạp file Item Detail
const ITEM_DEFAULT = {"BIB":["Bivina","Khác",0.071],"BIX":["Bivina","Khác",0.0792],"BV12":["Bia Viet","Group BB",0.0396],"BVC":["Bia Viet","Group BB",0.0792],"BVP":["Bia Viet","Group BB",0.0792],"EDB":["Edelweiss","Group BB",0.0792],"EDD":["Edelweiss","Group BB",0.2],"ESC4":["Edelweiss","Group BB",0.0792],"H025":["Heineken","Group BB",0.0792],"H0M":["Heineken","Group BB",0.0792],"HA12":["Heineken","Group AA",0.0396],"HBB":["Heineken","Group BB",0.045],"HS":["Heineken","Group BB",0.0792],"HP":["Heineken","Group BB",0.0792],"HSB":["Heineken","Group AA",0.0792],"HS2":["Heineken","Group AA",0.0792],"HSC":["Heineken","Group AA",0.0792],"HSD":["Heineken","Group BB",0.2],"HZM":["Heineken","Group BB",0.0792],"L35":["Larue","Khác",0.071],"LC":["Larue","Khác",0.0792],"LSC":["Larue","Khác",0.0792],"LMC":["Larue","Khác",0.0792],"LGC":["Larue","Khác",0.0792],"SB6":["Strongbow","Group BB",0.0768],"SPB":["Strongbow","Group BB",0.0792],"SD6":["Strongbow","Group BB",0.0768],"SG6":["Strongbow","Group BB",0.0768],"SGC6":["Strongbow","Group BB",0.0768],"SKB":["Strongbow","Group BB",0.0792],"SKC6":["Strongbow","Group BB",0.0768],"SM32":["Strongbow","Group BB",0.0768],"SPC6":["Strongbow","Group BB",0.0768],"SBP":["Strongbow","Group BB",0.0768],"TC25":["Tiger","Group BB",0.0792],"TCPC":["Tiger","Group BB",0.0792],"TCPP":["Tiger","Group BB",0.0792],"TCS":["Tiger","Group BB",0.0792],"TD":["Tiger","Group BB",0.2],"TMC":["Tiger","Group AA",0.0792],"TP":["Tiger","Group BB",0.0792],"TP2C":["Tiger","Group BB",0.066],"TS":["Tiger","Group AA",0.0792],"TS25":["Tiger","Group AA",0.0792],"TSD":["Tiger","Khác",0.2],"TSS":["Tiger","Group AA",0.0792],"TWMC":["Tiger","Group BB",0.0792],"TWMP":["Tiger","Group BB",0.0792]};
const itemDefault = () => Object.fromEntries(Object.entries(ITEM_DEFAULT).map(([k, [fam, grp, hl]]) => [k, { fam, grp, hl }]));

/* ============================================================ làm tròn */
// Làm tròn theo phần dư lớn nhất: tổng các dòng nguyên luôn bằng tổng (đã làm tròn)
function lrRound(total, items) {
  const T = Math.round(total), out = new Map();
  if (!items.length) return out;
  const pos = items.filter(x => x[1] > 0), S = pos.reduce((a, x) => a + x[1], 0);
  if (S <= 0 || T <= 0) { for (const [k] of items) out.set(k, 0); return out; }
  let used = 0; const rem = [];
  for (const [k, v] of items) {
    if (!(v > 0)) { out.set(k, 0); continue; }
    const x = v * T / S, f = Math.floor(x); out.set(k, f); used += f; rem.push([k, x - f, v]);
  }
  rem.sort((a, b) => b[1] - a[1] || b[2] - a[2]);
  for (let i = 0; i < T - used && i < rem.length; i++) out.set(rem[i][0], out.get(rem[i][0]) + 1);
  return out;
}

/* ============================================================ chuẩn bị dữ liệu (chạy một lần khi nạp file) */
function parseTransfer(text) {
  const map = new Map();
  for (const part of str(text).split(/[,;\n]+/)) {
    const m = part.split(/=|→|->|>/).map(s => s.trim().toUpperCase()).filter(Boolean);
    if (m.length === 2) map.set(m[0], m[1]);
  }
  return map;
}
function prepare(inp, cfg) {
  const warn = [];
  const item = inp.item || itemDefault();
  const si = inp.si;
  const months = [...new Set(si.map(r => r.m))].sort((a, b) => a - b);
  const M = cfg.monthM || months[months.length - 1], T = nextMonth(M);
  const L = []; { let x = M; for (let i = 0; i < cfg.l3m; i++) { x = prevMonth(x); L.unshift(x); } }
  const tr = parseTransfer(cfg.transfer);

  // NPP hiện hành: tên đầy đủ theo mã, Area và Region của NPP (lấy tháng gần nhất)
  const nppName = new Map(), nppArea = new Map(), nppRegion = new Map(), areaRegion = new Map();
  for (const r of [...si].sort((a, b) => a.m - b.m)) {
    const c = distCode(r.dist); nppName.set(c, r.dist); nppArea.set(c, r.area); nppRegion.set(c, r.region); areaRegion.set(r.area, r.region);
  }
  const mapNpp = name => { let c = distCode(name); const seen = new Set(); while (tr.has(c) && !seen.has(c)) { seen.add(c); c = tr.get(c); } return c; };
  for (const [a, b] of tr) { if (!nppName.has(b)) warn.push(`Mã NPP nhận chuyển ${b} (từ ${a}) không có trong SI-SO.`); }
  const activeNpp = new Set(); for (const r of si) if (r.m === M && r.tC > 0) activeNpp.add(mapNpp(r.dist));
  if (!activeNpp.size) for (const r of si) if (r.m === M) activeNpp.add(mapNpp(r.dist));

  // Hệ số Case ↔ Htls theo SKU
  const hlOf = sc => {
    if (item[sc] && item[sc].hl > 0) return item[sc].hl;
    let c = 0, h = 0; for (const r of si) if (r.sc === sc) { c += r.tC; h += r.tH; }
    return c > 0 && h > 0 ? h / c : 0.0792;
  };
  const info = sc => item[sc] || { fam: "Khác", grp: "Khác", hl: 0 };
  const skuMissing = new Set();

  // Target tháng M (ưu tiên file Target by Month, thiếu thì lấy SI-SO) và M+1 theo Region × SKU, Area × SKU
  const tM = new Map(), tM1 = new Map(), tMArea = new Map(), tM1Area = new Map();
  let srcM = "SI-SO", srcM1 = null;
  const monthFiles = inp.targets.filter(t => t.kind === "month"), caseFiles = inp.targets.filter(t => t.kind === "case");
  const regOfArea = (a, r) => r || areaRegion.get(a) || "";
  const hasM = monthFiles.some(t => t.months.includes(M));
  if (hasM) {
    srcM = "File Target";
    for (const t of monthFiles) for (const r of t.rows) if (r.htls[M] != null) { const rg = regOfArea(r.area, r.region), c = r.htls[M] / hlOf(r.sc); add(tM, rg + D + r.sc, c); add(tMArea, r.area + D + r.sc, c); }
  } else {
    for (const r of si) if (r.m === M) { add(tM, r.region + D + r.sc, r.tC); add(tMArea, r.area + D + r.sc, r.tC); }
    warn.push(`File Target không có cột tháng ${ymLabel(M)}, Target tháng M lấy theo SI-SO.`);
  }
  if (caseFiles.length) {
    srcM1 = "File Target theo Case";
    for (const t of caseFiles) for (const r of t.rows) { const rg = regOfArea(r.area, ""); if (!rg) { warn.push(`Area ${r.area} trong file Target theo Case không có trong SI-SO.`); continue; } add(tM1, rg + D + r.sc, r.cases); add(tM1Area, r.area + D + r.sc, r.cases); }
  } else if (monthFiles.some(t => t.months.includes(T))) {
    srcM1 = "File Target by Month";
    for (const t of monthFiles) for (const r of t.rows) if (r.htls[T] != null) { const rg = regOfArea(r.area, r.region), c = r.htls[T] / hlOf(r.sc); add(tM1, rg + D + r.sc, c); add(tM1Area, r.area + D + r.sc, c); }
  } else throw new Error(`File Target chưa có số tháng ${ymLabel(T)} (tháng M+1). Nạp file Target by Month có cột ${ymLabel(T)} hoặc file Target theo Case.`);
  for (const k of [...tM1.keys(), ...tM.keys()]) { const sc = k.split(D)[1]; if (!item[sc]) skuMissing.add(sc); }

  // CP theo NPP × SKU
  const cp = new Map();
  for (const r of inp.cp || []) { const c = mapNpp(r.dist); if (!nppName.has(c)) { warn.push(`NPP ${r.dist} trong file CP không có trong SI-SO.`); continue; } add(cp, c + D + r.sc, r.v); }

  // Gốc L3M theo Outlet (đơn vị = NPP hiện hành × Outlet)
  const units = new Map(); // key -> {npp, area, region, oid, name, seg, base: Map(sc->sum L3M), any L3M, early}
  const Lset = new Set(L), firstL = L[0];
  let skippedNpp = new Set(), rowsUsed = 0;
  for (const r of inp.outlets) {
    if (r.m >= M && !Lset.has(r.m)) continue;
    const c = mapNpp(r.dist);
    if (!nppName.has(c)) { skippedNpp.add(r.dist); continue; }
    const k = c + D + r.oid;
    let u = units.get(k);
    if (!u) { u = { key: k, npp: c, area: nppArea.get(c), region: nppRegion.get(c), oid: r.oid, name: r.name, seg: r.seg, base: new Map(), early: false, lastM: 0 }; units.set(k, u); }
    if (r.m >= u.lastM) { u.lastM = r.m; if (r.name) u.name = r.name; if (r.seg) u.seg = r.seg; }
    if (Lset.has(r.m)) { add(u.base, r.sc, r.v); rowsUsed++; }
    else if (r.m < firstL) u.early = true;
  }
  if (skippedNpp.size) warn.push(`Bỏ qua dữ liệu Outlet của NPP không có trong SI-SO: ${[...skippedNpp].join(", ")}.`);
  const monthsOutlet = new Set(inp.outlets.map(r => r.m));
  const missL = L.filter(m => !monthsOutlet.has(m));
  if (missL.length) warn.push(`File SO theo Outlet thiếu tháng ${missL.map(ymLabel).join(", ")} trong L3M.`);
  let negUnits = 0;
  for (const u of units.values()) {
    let tot = 0;
    for (const [sc, v] of u.base) { if (v < 0) { negUnits++; u.base.set(sc, 0); } else tot += v; }
    u.l3m = tot;
    u.dormant = tot <= 0 && u.early && activeNpp.has(u.npp);
  }
  if (negUnits) warn.push(`${negUnits} dòng Outlet × SKU có Sales Out L3M âm, tính bằng 0.`);

  // Đối chiếu SO Outlet với SO SI-SO trong L3M
  let soSI = 0, soOut = 0;
  for (const r of si) if (Lset.has(r.m)) soSI += r.so;
  for (const r of inp.outlets) if (Lset.has(r.m)) soOut += r.v;

  return { siRows: si, M, T, L, item, info, hlOf, skuMissing, tM, tM1, tMArea, tM1Area, srcM, srcM1, cp, units, nppName, nppArea, nppRegion, areaRegion, activeNpp, warn, soSI, soOut, rowsUsed, mapNpp };
}

/* ============================================================ chia số (chạy lại mỗi lần Adjust) */
// adj: { area: Map("region|area|fam" -> delta), npp: Map("npp|sc" -> case) }
function allocate(P, cfg, adj) {
  adj = adj || {};
  const adjArea = adj.area || new Map(), adjNpp = adj.npp || new Map();
  const main = new Set(cfg.mainSkus);
  const regions = [...new Set([...P.tM1.keys(), ...P.tM.keys()].map(k => k.split(D)[0]))].filter(Boolean).sort();
  const skuAll = new Set([...P.tM1.keys(), ...P.tM.keys()].map(k => k.split(D)[1]));
  const unitsBy = new Map(); for (const u of P.units.values()) { if (!unitsBy.has(u.region)) unitsBy.set(u.region, []); unitsBy.get(u.region).push(u); }
  const areasOf = rg => [...new Set([...P.nppArea.entries()].filter(([c]) => P.activeNpp.has(c) && P.nppRegion.get(c) === rg).map(([, a]) => a))].sort();
  const checks = [], regionSku = [];
  const alloc = new Map();   // unitKey|sc -> {raw, gap, val, type}
  const pending = new Map(); // area|sc -> case chờ adjust
  const cpBy = (rg, sc) => { let t = 0; for (const [k, v] of P.cp) { const [c, s] = k.split(D); if (s === sc && P.nppRegion.get(c) === rg) t += v; } return t; };

  // Bước 1–4: g, gốc L3M, tối thiểu, phần chênh theo contribution
  for (const rg of regions) {
    const units = unitsBy.get(rg) || [];
    const dormant = units.filter(u => u.dormant);
    for (const sc of [...skuAll].sort()) {
      const TM = P.tM.get(rg + D + sc) || 0, T1 = P.tM1.get(rg + D + sc) || 0;
      if (!TM && !T1) continue;
      const CP = cpBy(rg, sc), pool = Math.max(Math.round(T1 - CP), 0);
      const g = TM > 0 ? T1 / TM - 1 : null;
      const inf = P.info(sc);
      let B = 0; const based = [];
      for (const u of units) { const b = (u.base.get(sc) || 0) / cfg.l3m; if (b > 0) { based.push([u, b]); B += b; } }
      let reserved = 0, nRes = 0;
      if (main.has(sc) && dormant.length && cfg.minCase > 0 && B > 0) {
        if (dormant.length * cfg.minCase <= pool) { nRes = dormant.length; reserved = nRes * cfg.minCase; for (const u of dormant) alloc.set(u.key + D + sc, { raw: 0, gap: 0, val: cfg.minCase, type: "Tối thiểu", fixed: true }); }
        else checks.push({ lvl: "warn", msg: `${sc} (${rg}): Target sau CP (${pool}) nhỏ hơn số case tối thiểu cho ${dormant.length} Outlet chưa mua L3M, bỏ qua tối thiểu.` });
      }
      const rest = pool - reserved;
      let status = "Chia từ Outlet", sumRaw = 0;
      if (B > 0) {
        for (const [u, b] of based) sumRaw += b * (1 + (g ?? 0));
        const gap = rest - sumRaw;
        for (const [u, b] of based) { const raw = b * (1 + (g ?? 0)); alloc.set(u.key + D + sc, { raw, gap: gap * b / B, val: raw + gap * b / B, type: "L3M", fixed: false, base: b }); }
      } else if (rest > 0) {
        status = "Chờ adjust";
        // Như rule cũ: giữ số Area trong file Target, NPP adjust tay
        const areas = areasOf(rg), items = areas.map(a => [a, P.tM1Area.get(a + D + sc) || 0]);
        const tot = items.reduce((a, x) => a + x[1], 0);
        const split = tot > 0 ? lrRound(rest, items) : lrRound(rest, areas.map(a => [a, 1]));
        for (const [a, v] of split) if (v) pending.set(rg + D + a + D + sc, v);
        checks.push({ lvl: "adj", msg: `${sc} (${rg}): chưa có Sales Out L3M ở Outlet nào, giữ số theo Area trong file Target (${rest.toLocaleString("vi-VN")} case), chờ adjust xuống NPP.` });
      }
      regionSku.push({ region: rg, sc, fam: inf.fam, grp: inf.grp, TM, T1, g, CP, pool, reserved, nRes, base: B, nOutlet: based.length, gImplied: B > 0 ? (rest / B - 1) : null, gap: B > 0 ? rest - sumRaw : 0, status });
    }
  }

  // Tổng hợp đề xuất theo Region × Area × Brand family
  const famOf = sc => P.info(sc).fam;
  const prop = new Map(), fixedPart = new Map();
  const addP = (map, k, v) => add(map, k, v);
  for (const [k, a] of alloc) { const [c, oid, sc] = k.split(D); const ar = P.nppArea.get(c), rg = P.nppRegion.get(c); const kk = rg + D + ar + D + famOf(sc); addP(prop, kk, a.val); if (a.fixed) addP(fixedPart, kk, a.val); }
  for (const [k, v] of pending) { const [rg, ar, sc] = k.split(D); const kk = rg + D + ar + D + famOf(sc); addP(prop, kk, v); addP(fixedPart, kk, v); }

  // Bước 6: Adjust % tăng trưởng Area theo Brand family, các Area còn lại tự cân để giữ tổng Region
  const areaFam = [], factor = new Map();
  const famList = [...new Set([...prop.keys()].map(k => k.split(D)[2]))];
  const tMAreaFam = new Map(); for (const [k, v] of P.tMArea) { const [ar, sc] = k.split(D); add(tMAreaFam, (P.areaRegion.get(ar) || "") + D + ar + D + famOf(sc), v); }
  for (const rg of regions) for (const fam of famList) {
    const areas = areasOf(rg).filter(a => prop.has(rg + D + a + D + fam) || tMAreaFam.has(rg + D + a + D + fam));
    if (!areas.length) continue;
    const Pv = a => prop.get(rg + D + a + D + fam) || 0;
    const total = areas.reduce((s, a) => s + Pv(a), 0);
    const want = new Map(); let fixedSum = 0, freeP = 0;
    for (const a of areas) {
      const k = rg + D + a + D + fam;
      if (adjArea.has(k)) { let d = adjArea.get(k); d = Math.max(cfg.floor, Math.min(cfg.cap, d)); want.set(a, Pv(a) * (1 + d)); fixedSum += Pv(a) * (1 + d); }
      else freeP += Pv(a);
    }
    const remain = total - fixedSum;
    let imbalance = 0;
    for (const a of areas) if (!want.has(a)) want.set(a, freeP > 0 ? Pv(a) * remain / freeP : 0);
    if (freeP <= 0 && Math.abs(remain) > 0.5) imbalance = -remain;
    // RAS: giữ tổng từng SKU của Region và đạt tổng mong muốn của từng Area trong Brand family
    const X = new Map(), colT = new Map(), rowT = new Map();
    for (const [k, a] of alloc) { if (a.fixed) continue; const [c, , sc] = k.split(D); if (P.nppRegion.get(c) !== rg || famOf(sc) !== fam) continue; const ar = P.nppArea.get(c); add(X, ar + D + sc, a.val); add(colT, sc, a.val); }
    for (const a of areas) rowT.set(a, Math.max(want.get(a) - (fixedPart.get(rg + D + a + D + fam) || 0), 0));
    const Y = new Map(X);
    if (adjArea.size) for (let it = 0; it < 200; it++) {
      const rs = new Map(); for (const [k, v] of Y) add(rs, k.split(D)[0], v);
      for (const [k, v] of Y) { const a = k.split(D)[0], r = rs.get(a); Y.set(k, r > 0 ? v * rowT.get(a) / r : v); }
      const cs = new Map(); for (const [k, v] of Y) add(cs, k.split(D)[1], v);
      let err = 0;
      for (const [k, v] of Y) { const sc = k.split(D)[1], c = cs.get(sc); const nv = c > 0 ? v * colT.get(sc) / c : v; Y.set(k, nv); }
      const rs2 = new Map(); for (const [k, v] of Y) add(rs2, k.split(D)[0], v);
      for (const a of areas) err = Math.max(err, Math.abs((rs2.get(a) || 0) - rowT.get(a)));
      if (err < 0.01) break;
    }
    for (const [k, v] of X) factor.set(rg + D + k, v > 0 ? Y.get(k) / v : 1);
    const finalOf = new Map(); for (const [k, v] of Y) add(finalOf, k.split(D)[0], v);
    for (const a of areas) {
      const k = rg + D + a + D + fam, p = Pv(a), fx = fixedPart.get(k) || 0, w = (finalOf.get(a) || 0) + fx;
      const tm = tMAreaFam.get(k) || 0, delta = p > 0 ? w / p - 1 : 0;
      const edited = adjArea.has(k);
      const flags = [];
      if (!edited && (delta < cfg.floor - 1e-9 || delta > cfg.cap + 1e-9)) flags.push("Vượt giới hạn khi tự cân");
      if (tm > 0 && w / tm - 1 > cfg.capVsM + 1e-9) flags.push(`Tăng trên ${Math.round(cfg.capVsM * 100)}% so với Target M`);
      areaFam.push({ region: rg, area: a, fam, TM: tm, proposal: p, final: w, delta, edited, input: edited ? adjArea.get(k) : null, gProp: tm > 0 ? p / tm - 1 : null, gFinal: tm > 0 ? w / tm - 1 : null, flags, imbalance: a === areas[areas.length - 1] ? imbalance : 0 });
    }
    if (imbalance) checks.push({ lvl: "warn", msg: `${fam} (${rg}): mọi Area đều đã adjust, tổng lệch ${Math.round(imbalance).toLocaleString("vi-VN")} case so với Target Region. Bỏ adjust ở một Area để tự cân.` });
  }
  for (const [k, a] of alloc) { if (a.fixed) { a.adj = a.val; continue; } const [c, , sc] = k.split(D); const kk = P.nppRegion.get(c) + D + P.nppArea.get(c) + D + sc; a.adj = a.val * (factor.get(kk) ?? 1); }
  const pendAdj = new Map(); for (const [k, v] of pending) pendAdj.set(k, v); // phần chờ adjust giữ nguyên (đã tính vào phần cố định)

  // Bước 5: làm tròn theo tầng Region → Area → NPP → Outlet, cộng dồn
  const areaSku = new Map(), nppSku = new Map(), outletRows = [];
  const byRS = new Map();
  for (const [k, a] of alloc) { const [c, oid, sc] = k.split(D); const rg = P.nppRegion.get(c); const kk = rg + D + sc; if (!byRS.has(kk)) byRS.set(kk, []); byRS.get(kk).push([k, a, c, oid, sc]); }
  for (const [kk, list] of byRS) {
    const [rg, sc] = kk.split(D);
    const fixedL = list.filter(x => x[1].fixed), varL = list.filter(x => !x[1].fixed);
    const fixedTot = fixedL.reduce((s, x) => s + x[1].adj, 0);
    const varTot = varL.reduce((s, x) => s + x[1].adj, 0);
    for (const x of fixedL) x[1].fin = Math.round(x[1].adj);
    // Area
    const aMap = new Map(); for (const x of varL) add(aMap, P.nppArea.get(x[2]), x[1].adj);
    const aR = lrRound(varTot, [...aMap]);
    for (const [ar, av] of aR) {
      const inA = varL.filter(x => P.nppArea.get(x[2]) === ar);
      const nMap = new Map(); for (const x of inA) add(nMap, x[2], x[1].adj);
      const nR = lrRound(av, [...nMap]);
      for (const [c, nv] of nR) {
        const inN = inA.filter(x => x[2] === c);
        const oR = lrRound(nv, inN.map(x => [x[0], x[1].adj]));
        for (const x of inN) x[1].fin = oR.get(x[0]) || 0;
      }
    }
  }
  for (const [k, a] of alloc) {
    const [c, oid, sc] = k.split(D); const u = P.units.get(c + D + oid);
    const rg = P.nppRegion.get(c), ar = P.nppArea.get(c);
    add(nppSku, c + D + sc, a.fin); add(areaSku, rg + D + ar + D + sc, a.fin);
    outletRows.push({ region: rg, area: ar, npp: c, nppName: P.nppName.get(c), oid, name: u.name, seg: u.seg, sc, fam: famOf(sc), grp: P.info(sc).grp, base: a.base || 0, raw: a.raw, gap: a.gap, prop: a.val, final: a.fin, type: a.type });
  }
  // Chờ adjust: số Area giữ nguyên, NPP nhập tay
  const pendRows = [];
  for (const [k, v] of pendAdj) {
    const [rg, ar, sc] = k.split(D);
    add(areaSku, rg + D + ar + D + sc, v);
    const npps = [...P.activeNpp].filter(c => P.nppArea.get(c) === ar).sort();
    let given = 0;
    for (const c of npps) { const x = adjNpp.get(c + D + sc) || 0; if (x) { add(nppSku, c + D + sc, x); given += x; } }
    if (npps.length === 1 && !adjNpp.has(npps[0] + D + sc)) { add(nppSku, npps[0] + D + sc, v); given = v; }
    pendRows.push({ region: rg, area: ar, sc, fam: famOf(sc), areaCase: v, given, left: v - given, npps });
    if (v - given) checks.push({ lvl: "adj", msg: `${sc} – ${ar}: còn ${(v - given).toLocaleString("vi-VN")} case chưa giao cho NPP.` });
  }
  // CP cộng thẳng vào NPP
  for (const [k, v] of P.cp) { const [c, sc] = k.split(D); add(nppSku, c + D + sc, 0); }

  // Bảng NPP × SKU
  const nppRows = [];
  const tMNpp = new Map(), soL = new Map();
  for (const u of P.units.values()) for (const [sc, b] of u.base) add(soL, u.npp + D + sc, b / cfg.l3m);
  const keysN = new Set([...nppSku.keys(), ...P.cp.keys()]);
  for (const k of keysN) {
    const [c, sc] = k.split(D); const cpv = P.cp.get(k) || 0, out = nppSku.get(k) || 0;
    nppRows.push({ region: P.nppRegion.get(c), area: P.nppArea.get(c), npp: c, nppName: P.nppName.get(c), sc, fam: famOf(sc), grp: P.info(sc).grp, soL3M: soL.get(k) || 0, outlet: out, cp: cpv, final: out + cpv });
  }
  // Target M của NPP (SI-SO, sau chuyển NPP)
  const tmN = new Map(); for (const r of P.siRows || []) if (r.m === P.M) add(tmN, P.mapNpp(r.dist) + D + r.sc, r.tC);
  for (const r of nppRows) r.TM = tmN.get(r.npp + D + r.sc) || 0;

  // Đối chiếu tổng
  for (const rs of regionSku) {
    let s = 0; for (const r of nppRows) if (r.region === rs.region && r.sc === rs.sc) s += r.final;
    rs.final = s;
    const pendLeft = pendRows.filter(p => p.region === rs.region && p.sc === rs.sc).reduce((a, p) => a + p.left, 0);
    rs.diff = Math.round(rs.T1) - s - pendLeft;
  }
  return { regions, regionSku, areaFam, nppRows, outletRows, pendRows, areaSku, checks };
}

const api = { readWorkbook, parseSI, parseOutlet, parseItem, parseTarget, parseCP, itemDefault, prepare, allocate, lrRound, nextMonth, prevMonth, ymLabel, distCode, D, num, str, add };
if (typeof module !== "undefined" && module.exports) module.exports = api; else root.TSE = api;
})(typeof window !== "undefined" ? window : globalThis);
