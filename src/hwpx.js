import headerXml from "./hwpx-header.xml?raw";
import { cellText, downloadBlob } from "./excel.js";

const enc = new TextEncoder();
const PAGE = { w: 59528, h: 84188 };
const MARGIN = 2268;
const CHAR_BASE = headerXml.match(/<hh:charPr id="2"[\s\S]*?<\/hh:charPr>/)?.[0] || "";
const PARA_LEFT = headerXml.match(/<hh:paraPr id="11"[\s\S]*?<\/hh:paraPr>/)?.[0] || "";

const CRC_TABLE = new Uint32Array(256);
for (let n = 0; n < 256; n++) {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  CRC_TABLE[n] = c;
}

function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 255] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}
function u16(n) {
  return Uint8Array.of(n & 255, (n >>> 8) & 255);
}
function u32(n) {
  return Uint8Array.of(n & 255, (n >>> 8) & 255, (n >>> 16) & 255, (n >>> 24) & 255);
}
function cat(parts) {
  const n = parts.reduce((s, p) => s + p.length, 0);
  const out = new Uint8Array(n);
  let o = 0;
  for (const p of parts) {
    out.set(p, o);
    o += p.length;
  }
  return out;
}
function zipStore(files) {
  const locals = [];
  const centrals = [];
  let offset = 0;
  for (const { name, data } of files) {
    const nameB = enc.encode(name);
    const crc = crc32(data);
    const local = cat([
      enc.encode("PK\x03\x04"),
      u16(20),
      u16(0),
      u16(0),
      u16(0),
      u16(0),
      u32(crc),
      u32(data.length),
      u32(data.length),
      u16(nameB.length),
      u16(0),
      nameB,
      data,
    ]);
    locals.push(local);
    centrals.push(
      cat([
        enc.encode("PK\x01\x02"),
        u16(20),
        u16(20),
        u16(0),
        u16(0),
        u16(0),
        u16(0),
        u32(crc),
        u32(data.length),
        u32(data.length),
        u16(nameB.length),
        u16(0),
        u16(0),
        u16(0),
        u16(0),
        u32(0),
        u32(offset),
        nameB,
      ]),
    );
    offset += local.length;
  }
  const central = cat(centrals);
  const end = cat([
    enc.encode("PK\x05\x06"),
    u16(0),
    u16(0),
    u16(files.length),
    u16(files.length),
    u32(central.length),
    u32(offset),
    u16(0),
  ]);
  return cat([...locals, central, end]);
}

function xesc(s) {
  return String(s ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}
function rgbHex(value) {
  const s = String(value || "");
  if (!s || s === "transparent") return "FFFFFF";
  const m = s.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([0-9.]+))?\)/i);
  if (!m) return "FFFFFF";
  if (m[4] !== undefined && Number(m[4]) === 0) return "FFFFFF";
  return [m[1], m[2], m[3]].map((n) => Number(n).toString(16).padStart(2, "0")).join("").toUpperCase();
}
function pxHwp(px) {
  return Math.max(200, Math.round(Number(px) * 75));
}
function fit(nums, total) {
  const out = nums.map((n) => Math.max(400, Math.round(n)));
  if (out.length) out[out.length - 1] += total - out.reduce((a, b) => a + b, 0);
  return out;
}
function contentW(landscape) {
  return (landscape ? PAGE.h : PAGE.w) - MARGIN * 2;
}

function makeStyles() {
  const fills = new Map([["FFFFFF", "3"]]);
  const chars = new Map();
  let fillN = 3;
  let charN = 6;
  return {
    fillId(hex) {
      const h = hex || "FFFFFF";
      if (!fills.has(h)) fills.set(h, String(++fillN));
      return fills.get(h);
    },
    charId(height, color) {
      const h = Math.max(600, Math.min(2400, Number(height) || 900));
      const c = color || "000000";
      const key = `${h}|${c}`;
      if (!chars.has(key)) chars.set(key, String(++charN));
      return chars.get(key);
    },
    fills,
    chars,
  };
}

function elStyle(el) {
  const cs = getComputedStyle(el);
  return {
    fill: rgbHex(cs.backgroundColor),
    color: rgbHex(cs.color),
    size: Math.round(parseFloat(cs.fontSize) * 75),
    center: cs.textAlign === "center" || el.tagName === "TH",
    middle: cs.verticalAlign === "middle" || el.tagName === "TH",
  };
}

function cellLines(el) {
  const bits = [...el.querySelectorAll(":scope > .n, :scope > .md, .act, .hol")];
  if (bits.length) {
    return bits
      .map((n) => ({ text: n.textContent.replace(/\s+/g, " ").trim(), ...elStyle(n) }))
      .filter((n) => n.text);
  }
  const st = elStyle(el);
  return String(cellText(el) || "")
    .split(/\n/)
    .map((t) => t.trim())
    .filter(Boolean)
    .map((text) => ({ text, ...st }));
}

function tableGrid(table) {
  const taken = new Set();
  const cells = [];
  let rows = 0;
  let cols = 0;
  [...table.rows].forEach((tr, ri) => {
    let c = 0;
    [...tr.cells].forEach((el) => {
      while (taken.has(`${ri},${c}`)) c += 1;
      const rs = el.rowSpan || 1;
      const cs = el.colSpan || 1;
      cells.push({ r: ri, c, rs, cs, el, lines: cellLines(el), header: el.tagName === "TH" });
      for (let i = 0; i < rs; i++) for (let j = 0; j < cs; j++) taken.add(`${ri + i},${c + j}`);
      rows = Math.max(rows, ri + rs);
      cols = Math.max(cols, c + cs);
      c += cs;
    });
  });
  return { cells, rows, cols };
}

function colWidths(table, cols, width) {
  const row = table.tHead?.rows[0] || table.rows[0];
  const raw = Array(cols).fill(width / Math.max(cols, 1));
  if (row) {
    let c = 0;
    [...row.cells].forEach((el) => {
      const cs = el.colSpan || 1;
      const w = Math.max(1, el.getBoundingClientRect().width || 1);
      for (let i = 0; i < cs && c + i < cols; i++) raw[c + i] = pxHwp(w / cs);
      c += cs;
    });
  }
  return fit(raw, width);
}

function rowHeights(table, rows, width) {
  const tablePx = table.getBoundingClientRect().width || 1;
  const scale = width / pxHwp(tablePx);
  const raw = Array(rows).fill(1200);
  [...table.rows].forEach((tr, ri) => {
    if (ri < rows) raw[ri] = Math.max(400, Math.round(pxHwp(tr.getBoundingClientRect().height) * scale));
  });
  return raw;
}

function paraXml(id, text, charPr, paraPr = "0") {
  const t = xesc(text);
  const body = t ? `<hp:t>${t}</hp:t>` : "<hp:t/>";
  return `<hp:p id="${id}" paraPrIDRef="${paraPr}" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="${charPr}">${body}</hp:run></hp:p>`;
}

function linesXml(lines, startId, styles, fallback) {
  const used = lines.length ? lines : [{ text: "", ...fallback }];
  return used
    .map((line, i) =>
      paraXml(
        startId + i,
        line.text,
        styles.charId(line.size || fallback.size, line.color || fallback.color),
        line.center || fallback.center ? "20" : "0",
      ),
    )
    .join("");
}

function tcXml(cell, colW, rowH, styles, pid) {
  let w = 0;
  let h = 0;
  for (let j = 0; j < cell.cs; j++) w += colW[cell.c + j] || 400;
  for (let i = 0; i < cell.rs; i++) h += rowH[cell.r + i] || 400;
  const st = elStyle(cell.el);
  const body = linesXml(cell.lines, pid, styles, st);
  const n = Math.max(1, cell.lines.length || 1);
  return {
    xml: `<hp:tc name="" header="${cell.header ? 1 : 0}" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="${styles.fillId(st.fill)}"><hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="${st.middle ? "CENTER" : "TOP"}" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">${body}</hp:subList><hp:cellAddr colAddr="${cell.c}" rowAddr="${cell.r}"/><hp:cellSpan colSpan="${cell.cs}" rowSpan="${cell.rs}"/><hp:cellSz width="${w}" height="${h}"/><hp:cellMargin left="80" right="80" top="40" bottom="40"/></hp:tc>`,
    next: pid + n,
  };
}

function tblXml(table, pid, width, styles) {
  const { cells, rows, cols } = tableGrid(table);
  if (!rows || !cols) return { xml: "", nextId: pid };
  const colW = colWidths(table, cols, width);
  const rowH = rowHeights(table, rows, width);
  const height = rowH.reduce((a, b) => a + b, 0);
  let next = pid + 1;
  const trs = [];
  for (let r = 0; r < rows; r++) {
    const tcs = [];
    for (const cell of cells.filter((c) => c.r === r)) {
      const placed = tcXml(cell, colW, rowH, styles, next);
      tcs.push(placed.xml);
      next = placed.next;
    }
    trs.push(`<hp:tr>${tcs.join("")}</hp:tr>`);
  }
  return {
    xml: `<hp:p id="${pid}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="1"><hp:tbl id="${2000000000 + pid}" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="NONE" repeatHeader="1" rowCnt="${rows}" colCnt="${cols}" cellSpacing="0" borderFillIDRef="3" noAdjust="0"><hp:sz width="${width}" widthRelTo="ABSOLUTE" height="${height}" heightRelTo="ABSOLUTE" protect="0"/><hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/><hp:outMargin left="0" right="0" top="0" bottom="0"/><hp:inMargin left="80" right="80" top="40" bottom="40"/>${trs.join("")}</hp:tbl></hp:run></hp:p>`,
    nextId: next,
  };
}

function tcWrap(inner, fill, w, h, col, row) {
  return `<hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="${fill}"><hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">${inner}</hp:subList><hp:cellAddr colAddr="${col}" rowAddr="${row}"/><hp:cellSpan colSpan="1" rowSpan="1"/><hp:cellSz width="${w}" height="${h}"/><hp:cellMargin left="80" right="80" top="40" bottom="40"/></hp:tc>`;
}

function cardInner(card, width, pid, styles) {
  const parts = [];
  let next = pid;
  const heads = [...card.querySelectorAll("h3, .month-top")].map((n) => n.innerText.trim()).filter(Boolean);
  if (heads.length) {
    const st = elStyle(card.querySelector("h3, .month-top"));
    parts.push(paraXml(next, heads.join("  "), styles.charId(Math.max(st.size, 1100), st.color), "20"));
    next += 1;
  }
  const table = card.querySelector("table");
  if (table) {
    const placed = tblXml(table, next, Math.max(1200, width - 200), styles);
    parts.push(placed.xml);
    next = placed.nextId;
  }
  const extra = card.querySelector(".month-events, .daylist");
  if (extra) {
    const st = elStyle(extra);
    for (const line of String(cellText(extra) || "")
      .split(/\n/)
      .map((t) => t.trim())
      .filter(Boolean)) {
      parts.push(paraXml(next, line, styles.charId(st.size, st.color), "0"));
      next += 1;
    }
  }
  if (!parts.length) parts.push(paraXml(next++, "", styles.charId(800, "000000")));
  return { xml: parts.join(""), nextId: next };
}

function cardsXml(cards, cols, width, pid, styles) {
  const rows = Math.ceil(cards.length / cols) || 1;
  const colW = fit(Array(cols).fill(width / cols), width);
  const rowH = Array.from({ length: rows }, (_, r) => {
    let h = 4000;
    for (let c = 0; c < cols; c++) {
      const card = cards[r * cols + c];
      if (card) h = Math.max(h, pxHwp(card.getBoundingClientRect().height));
    }
    return h;
  });
  let next = pid + 1;
  const trs = [];
  const emptyFill = styles.fillId("FFFFFF");
  for (let r = 0; r < rows; r++) {
    const tcs = [];
    for (let c = 0; c < cols; c++) {
      const card = cards[r * cols + c];
      const w = colW[c];
      const h = rowH[r];
      if (!card) {
        tcs.push(tcWrap(paraXml(next++, "", styles.charId(800, "000000")), emptyFill, w, h, c, r));
        continue;
      }
      const inner = cardInner(card, w, next, styles);
      next = inner.nextId;
      tcs.push(tcWrap(inner.xml, emptyFill, w, h, c, r));
    }
    trs.push(`<hp:tr>${tcs.join("")}</hp:tr>`);
  }
  const height = rowH.reduce((a, b) => a + b, 0);
  return {
    xml: `<hp:p id="${pid}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="1"><hp:tbl id="${2000000000 + pid}" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="NONE" repeatHeader="0" rowCnt="${rows}" colCnt="${cols}" cellSpacing="0" borderFillIDRef="3" noAdjust="0"><hp:sz width="${width}" widthRelTo="ABSOLUTE" height="${height}" heightRelTo="ABSOLUTE" protect="0"/><hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/><hp:outMargin left="0" right="0" top="0" bottom="0"/><hp:inMargin left="40" right="40" top="40" bottom="40"/>${trs.join("")}</hp:tbl></hp:run></hp:p>`,
    nextId: next,
  };
}

function paperBlocks(paper) {
  const blocks = [];
  const title = paper.querySelector("h1, .ym-box")?.innerText.trim();
  const school = paper.querySelector(".school")?.innerText.trim();
  if (title) blocks.push({ type: "p", text: title, title: true });
  if (school && school !== title) blocks.push({ type: "p", text: school });
  const grid = paper.querySelector(".month-grid-12, .sem-cals");
  const cards = [...paper.querySelectorAll(".month-card, .sem-col")];
  if (grid && cards.length) {
    const cols =
      getComputedStyle(grid)
        .gridTemplateColumns.split(" ")
        .filter(Boolean).length || (grid.classList.contains("sem-cals") ? cards.length : 4);
    blocks.push({ type: "cards", cards, cols });
  } else {
    for (const table of paper.querySelectorAll("table")) blocks.push({ type: "table", table });
  }
  const note = paper.querySelector(".note, .sheet-foot")?.innerText.trim();
  if (note && note !== school) blocks.push({ type: "p", text: note });
  return blocks;
}

function headerWithStyles(styles) {
  const fillXml = [...styles.fills.entries()]
    .map(
      ([hex, id]) =>
        `<hh:borderFill id="${id}" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0"><hh:slash type="NONE" Crooked="0" isCounter="0"/><hh:backSlash type="NONE" Crooked="0" isCounter="0"/><hh:leftBorder type="SOLID" width="0.12 mm" color="#4A5568"/><hh:rightBorder type="SOLID" width="0.12 mm" color="#4A5568"/><hh:topBorder type="SOLID" width="0.12 mm" color="#4A5568"/><hh:bottomBorder type="SOLID" width="0.12 mm" color="#4A5568"/><hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/><hc:fillBrush><hc:winBrush faceColor="#${hex}" hatchColor="#999999" alpha="0"/></hc:fillBrush></hh:borderFill>`,
    )
    .join("");
  const charXml = [...styles.chars.entries()]
    .map(([key, id]) => {
      const [height, color] = key.split("|");
      return CHAR_BASE.replace('id="2"', `id="${id}"`)
        .replace('height="900"', `height="${height}"`)
        .replace('textColor="#000000"', `textColor="#${color}"`);
    })
    .join("");
  const center = PARA_LEFT.replace('id="11"', 'id="20"').replace('horizontal="LEFT"', 'horizontal="CENTER"');
  return headerXml
    .replace('hh:borderFills itemCnt="2"', `hh:borderFills itemCnt="${2 + styles.fills.size}"`)
    .replace("</hh:borderFills>", `${fillXml}</hh:borderFills>`)
    .replace('hh:charProperties itemCnt="7"', `hh:charProperties itemCnt="${7 + styles.chars.size}"`)
    .replace("</hh:charProperties>", `${charXml}</hh:charProperties>`)
    .replace('hh:paraProperties itemCnt="20"', 'hh:paraProperties itemCnt="21"')
    .replace("</hh:paraProperties>", `${center}</hh:paraProperties>`);
}

function secPr(landscape) {
  const w = landscape ? PAGE.h : PAGE.w;
  const h = landscape ? PAGE.w : PAGE.h;
  const land = landscape ? "NARROWLY" : "WIDELY";
  return `<hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134" tabStop="8000" tabStopVal="4000" tabStopUnit="HWPUNIT" outlineShapeIDRef="1" memoShapeIDRef="0" textVerticalWidthHead="0" masterPageCnt="0"><hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/><hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/><hp:visibility hideFirstHeader="0" hideFirstFooter="0" hideFirstMasterPage="0" border="SHOW_ALL" fill="SHOW_ALL" hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0"/><hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/><hp:pagePr landscape="${land}" width="${w}" height="${h}" gutterType="LEFT_ONLY"><hp:margin header="0" footer="0" gutter="0" left="${MARGIN}" right="${MARGIN}" top="${MARGIN}" bottom="${MARGIN}"/></hp:pagePr><hp:footNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/><hp:noteLine length="-1" type="SOLID" width="0.12 mm" color="#000000"/><hp:noteSpacing betweenNotes="283" belowLine="567" aboveLine="850"/><hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="EACH_COLUMN" beneathText="0"/></hp:footNotePr><hp:endNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/><hp:noteLine length="14692344" type="SOLID" width="0.12 mm" color="#000000"/><hp:noteSpacing betweenNotes="0" belowLine="567" aboveLine="850"/><hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="END_OF_DOCUMENT" beneathText="0"/></hp:endNotePr><hp:pageBorderFill type="BOTH" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="0" right="0" top="0" bottom="0"/></hp:pageBorderFill><hp:pageBorderFill type="EVEN" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="0" right="0" top="0" bottom="0"/></hp:pageBorderFill><hp:pageBorderFill type="ODD" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="0" right="0" top="0" bottom="0"/></hp:pageBorderFill></hp:secPr><hp:ctrl><hp:colPr id="" type="NEWSPAPER" layout="LEFT" colCount="1" sameSz="1" sameGap="0"/></hp:ctrl>`;
}

const NS =
  'xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph" xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history" xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:opf="http://www.idpf.org/2007/opf/" xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart" xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar" xmlns:epub="http://www.idpf.org/2007/ops" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"';

function buildSection(blocks, landscape, styles) {
  const width = contentW(landscape);
  const first = blocks.find((b) => b.type === "p") || { text: "학사일정표", title: true };
  const rest = blocks.filter((b) => b !== first);
  const titlePr = styles.charId(first.title ? 1800 : 1000, first.title ? "1B365D" : "000000");
  let pid = 1;
  const parts = [
    `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hs:sec ${NS}><hp:p id="0" paraPrIDRef="20" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0">${secPr(landscape)}</hp:run><hp:run charPrIDRef="${titlePr}"><hp:t>${xesc(first.text || "학사일정표")}</hp:t></hp:run></hp:p>`,
  ];
  for (const block of rest) {
    if (block.type === "p") {
      parts.push(paraXml(pid, block.text, styles.charId(block.title ? 1400 : 900, "000000"), block.title ? "20" : "0"));
      pid += 1;
    } else if (block.type === "table") {
      const placed = tblXml(block.table, pid, width, styles);
      parts.push(placed.xml);
      pid = placed.nextId;
    } else if (block.type === "cards") {
      const placed = cardsXml(block.cards, block.cols, width, pid, styles);
      parts.push(placed.xml);
      pid = placed.nextId;
    }
  }
  parts.push("</hs:sec>");
  return parts.join("");
}

function contentHpf() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><opf:package ${NS} version="" unique-identifier="" id=""><opf:metadata><opf:title/><opf:language>ko</opf:language></opf:metadata><opf:manifest><opf:item id="header" href="Contents/header.xml" media-type="application/xml"/><opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/><opf:item id="settings" href="settings.xml" media-type="application/xml"/></opf:manifest><opf:spine><opf:itemref idref="header" linear="yes"/><opf:itemref idref="section0" linear="yes"/></opf:spine></opf:package>`;
}

function packHwpx(sectionXml, header) {
  return zipStore([
    { name: "mimetype", data: enc.encode("application/hwp+zip") },
    {
      name: "version.xml",
      data: enc.encode(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hv:HCFVersion xmlns:hv="http://www.hancom.co.kr/hwpml/2011/version" tagetApplication="WORDPROCESSOR" major="5" minor="1" micro="1" buildNumber="0" os="1" xmlVersion="1.5" application="Hancom Office Hangul" appVersion="13, 0, 0, 1408 WIN32LEWindows_10"/>',
      ),
    },
    {
      name: "settings.xml",
      data: enc.encode(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><ha:HWPApplicationSetting xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"><ha:CaretPosition listIDRef="0" paraIDRef="0" pos="16"/></ha:HWPApplicationSetting>',
      ),
    },
    {
      name: "META-INF/container.xml",
      data: enc.encode(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><ocf:container xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf"><ocf:rootfiles><ocf:rootfile full-path="Contents/content.hpf" media-type="application/hwpml-package+xml"/></ocf:rootfiles></ocf:container>',
      ),
    },
    {
      name: "META-INF/manifest.xml",
      data: enc.encode(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><odf:manifest xmlns:odf="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>',
      ),
    },
    { name: "Contents/content.hpf", data: enc.encode(contentHpf()) },
    { name: "Contents/header.xml", data: enc.encode(header) },
    { name: "Contents/section0.xml", data: enc.encode(sectionXml) },
    { name: "Preview/PrvText.txt", data: enc.encode("학사일정표") },
  ]);
}

export function downloadHwpx(paper, filename) {
  const landscape = paper.classList.contains("landscape");
  const styles = makeStyles();
  const xml = buildSection(paperBlocks(paper), landscape, styles);
  downloadBlob(packHwpx(xml, headerWithStyles(styles)), filename, "application/hwp+zip");
}

export function selfCheck() {
  if (!CHAR_BASE || !PARA_LEFT) throw new Error("hwpx header templates");
  const table = document.createElement("table");
  table.innerHTML = "<tr><th>A</th><th>B</th></tr><tr><td class='exam'>1</td><td>2</td></tr>";
  document.body.appendChild(table);
  try {
    const styles = makeStyles();
    const xml = buildSection(
      [
        { type: "p", text: "check", title: true },
        { type: "table", table },
      ],
      false,
      styles,
    );
    if (!xml.includes("<hp:tbl") || xml.includes("<hp:pic") || xml.includes("\n")) throw new Error("hwpx section");
    const header = headerWithStyles(styles);
    if (!header.includes('hh:borderFill id="3"') || !header.includes('id="20"')) throw new Error("hwpx header styles");
    const zip = packHwpx(xml, header);
    if (zip[0] !== 0x50 || zip[1] !== 0x4b) throw new Error("hwpx zip");
    const nameLen = zip[26] | (zip[27] << 8);
    const name = new TextDecoder().decode(zip.slice(30, 30 + nameLen));
    if (name !== "mimetype") throw new Error("hwpx mimetype first");
  } finally {
    table.remove();
  }
}
