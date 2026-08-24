import headerXml from "./hwpx-header.xml?raw";
import { cellText, downloadBlob } from "./excel.js";

const enc = new TextEncoder();

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

function tableGrid(table) {
  const taken = new Set();
  const cells = [];
  let maxR = 0;
  let maxC = 0;
  [...table.rows].forEach((tr, ri) => {
    let c = 0;
    [...tr.cells].forEach((el) => {
      while (taken.has(`${ri},${c}`)) c += 1;
      const rs = el.rowSpan || 1;
      const cs = el.colSpan || 1;
      cells.push({ r: ri, c, rs, cs, text: cellText(el), header: el.tagName === "TH" });
      for (let i = 0; i < rs; i++) {
        for (let j = 0; j < cs; j++) taken.add(`${ri + i},${c + j}`);
      }
      maxR = Math.max(maxR, ri + rs);
      maxC = Math.max(maxC, c + cs);
      c += cs;
    });
  });
  return { cells, rows: maxR, cols: maxC };
}

function colRatios(table, cols) {
  const row = table.tHead?.rows[0] || table.rows[0];
  const widths = Array(cols).fill(1);
  if (row) {
    let c = 0;
    [...row.cells].forEach((el) => {
      const cs = el.colSpan || 1;
      const w = Math.max(1, el.getBoundingClientRect().width || el.offsetWidth || 1);
      for (let i = 0; i < cs && c + i < cols; i++) widths[c + i] = w / cs;
      c += cs;
    });
  }
  const sum = widths.reduce((a, b) => a + b, 0) || cols;
  return widths.map((w) => w / sum);
}

function paperBlocks(paper) {
  const blocks = [];
  const title = paper.querySelector("h1, .ym-box")?.innerText.trim();
  const school = paper.querySelector(".school")?.innerText.trim();
  if (title) blocks.push({ type: "p", text: title, title: true });
  if (school) blocks.push({ type: "p", text: school });
  const cards = [...paper.querySelectorAll(".month-card, .sem-col")];
  if (cards.length) {
    for (const card of cards) {
      const heading = card.querySelector("h3")?.innerText.trim();
      if (heading) blocks.push({ type: "p", text: heading });
      const table = card.querySelector("table");
      if (table) blocks.push({ type: "table", table });
      const extra = card.querySelector(".month-events, .daylist");
      if (extra) {
        const text = cellText(extra);
        if (text) blocks.push({ type: "p", text });
      }
    }
  } else {
    for (const table of paper.querySelectorAll("table")) blocks.push({ type: "table", table });
    const note = paper.querySelector(".note")?.innerText.trim();
    if (note) blocks.push({ type: "p", text: note });
  }
  return blocks;
}

function headerWithBorders() {
  const fill =
    '<hh:borderFill id="3" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0"><hh:slash type="NONE" Crooked="0" isCounter="0"/><hh:backSlash type="NONE" Crooked="0" isCounter="0"/><hh:leftBorder type="SOLID" width="0.12 mm" color="#4A5568"/><hh:rightBorder type="SOLID" width="0.12 mm" color="#4A5568"/><hh:topBorder type="SOLID" width="0.12 mm" color="#4A5568"/><hh:bottomBorder type="SOLID" width="0.12 mm" color="#4A5568"/><hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/></hh:borderFill>';
  return headerXml
    .replace('hh:borderFills itemCnt="2"', 'hh:borderFills itemCnt="3"')
    .replace("</hh:borderFills>", `${fill}</hh:borderFills>`);
}

function paraXml(id, text, charPr = "1") {
  return String(text || "")
    .split(/\n/)
    .map((line, i) => {
      const t = xesc(line);
      const body = t ? `<hp:t>${t}</hp:t>` : `<hp:t/>`;
      return `<hp:p id="${id + i}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="${charPr}">${body}</hp:run></hp:p>`;
    })
    .join("");
}

function tblXml(table, pid, landscape) {
  const { cells, rows, cols } = tableGrid(table);
  if (!rows || !cols) return { xml: "", nextId: pid };
  const pageW = landscape ? 67178 : 42520;
  const ratios = colRatios(table, cols);
  const colW = ratios.map((r) => Math.max(800, Math.round(r * pageW)));
  colW[colW.length - 1] += pageW - colW.reduce((a, b) => a + b, 0);
  const rowH = Array(rows).fill(1400);
  for (const cell of cells) {
    const lines = Math.max(1, String(cell.text || "").split(/\n/).length);
    const h = Math.max(1200, Math.min(3600, 900 + lines * 400));
    for (let i = 0; i < cell.rs; i++) rowH[cell.r + i] = Math.max(rowH[cell.r + i], Math.round(h / cell.rs));
  }
  const height = rowH.reduce((a, b) => a + b, 0);
  let next = pid;
  const trs = [];
  for (let r = 0; r < rows; r++) {
    const tcs = cells
      .filter((c) => c.r === r)
      .map((cell) => {
        let w = 0;
        for (let j = 0; j < cell.cs; j++) w += colW[cell.c + j] || 800;
        let h = 0;
        for (let i = 0; i < cell.rs; i++) h += rowH[cell.r + i] || 1400;
        const body = paraXml(next, cell.text, cell.header ? "6" : "2");
        next += Math.max(1, String(cell.text || "").split(/\n/).length);
        return `<hp:tc name="" header="${cell.header ? 1 : 0}" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="3"><hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">${body}</hp:subList><hp:cellAddr colAddr="${cell.c}" rowAddr="${cell.r}"/><hp:cellSpan colSpan="${cell.cs}" rowSpan="${cell.rs}"/><hp:cellSz width="${w}" height="${h}"/><hp:cellMargin left="140" right="140" top="80" bottom="80"/></hp:tc>`;
      })
      .join("");
    trs.push(`<hp:tr>${tcs}</hp:tr>`);
  }
  const xml = `<hp:p id="${next}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="1"><hp:tbl id="${2000000000 + pid}" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="NONE" repeatHeader="1" rowCnt="${rows}" colCnt="${cols}" cellSpacing="0" borderFillIDRef="3" noAdjust="0"><hp:sz width="${pageW}" widthRelTo="ABSOLUTE" height="${height}" heightRelTo="ABSOLUTE" protect="0"/><hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/><hp:outMargin left="0" right="0" top="0" bottom="0"/><hp:inMargin left="140" right="140" top="80" bottom="80"/>${trs.join("")}</hp:tbl></hp:run></hp:p>`;
  return { xml, nextId: next + 1 };
}

function secPr(landscape) {
  const width = landscape ? 84186 : 59528;
  const height = landscape ? 59528 : 84186;
  const land = landscape ? "NARROWLY" : "WIDELY";
  return `<hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134" tabStop="8000" tabStopVal="4000" tabStopUnit="HWPUNIT" outlineShapeIDRef="1" memoShapeIDRef="0" textVerticalWidthHead="0" masterPageCnt="0"><hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/><hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/><hp:visibility hideFirstHeader="0" hideFirstFooter="0" hideFirstMasterPage="0" border="SHOW_ALL" fill="SHOW_ALL" hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0"/><hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/><hp:pagePr landscape="${land}" width="${width}" height="${height}" gutterType="LEFT_ONLY"><hp:margin header="2834" footer="2834" gutter="0" left="4252" right="4252" top="4252" bottom="4252"/></hp:pagePr><hp:footNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/><hp:noteLine length="-1" type="SOLID" width="0.12 mm" color="#000000"/><hp:noteSpacing betweenNotes="283" belowLine="567" aboveLine="850"/><hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="EACH_COLUMN" beneathText="0"/></hp:footNotePr><hp:endNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/><hp:noteLine length="14692344" type="SOLID" width="0.12 mm" color="#000000"/><hp:noteSpacing betweenNotes="0" belowLine="567" aboveLine="850"/><hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="END_OF_DOCUMENT" beneathText="0"/></hp:endNotePr><hp:pageBorderFill type="BOTH" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill><hp:pageBorderFill type="EVEN" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill><hp:pageBorderFill type="ODD" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill></hp:secPr><hp:ctrl><hp:colPr id="" type="NEWSPAPER" layout="LEFT" colCount="1" sameSz="1" sameGap="0"/></hp:ctrl>`;
}

const NS =
  'xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph" xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history" xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:opf="http://www.idpf.org/2007/opf/" xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart" xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar" xmlns:epub="http://www.idpf.org/2007/ops" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"';

function buildSection(blocks, landscape) {
  const first = blocks.find((b) => b.type === "p") || { text: "학사일정표", title: true };
  const rest = blocks.filter((b) => b !== first);
  const title = xesc(first.text || "학사일정표");
  let pid = 1;
  const parts = [
    `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hs:sec ${NS}><hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0">${secPr(landscape)}</hp:run><hp:run charPrIDRef="${first.title ? "5" : "1"}"><hp:t>${title}</hp:t></hp:run></hp:p>`,
  ];
  for (const block of rest) {
    if (block.type === "p") {
      parts.push(paraXml(pid, block.text, block.title ? "5" : "1"));
      pid += Math.max(1, String(block.text || "").split(/\n/).length);
    } else if (block.type === "table") {
      const placed = tblXml(block.table, pid, landscape);
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

function packHwpx(sectionXml) {
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
    { name: "Contents/header.xml", data: enc.encode(headerWithBorders()) },
    { name: "Contents/section0.xml", data: enc.encode(sectionXml) },
    { name: "Preview/PrvText.txt", data: enc.encode("학사일정표") },
  ]);
}

export function downloadHwpx(paper, filename) {
  const landscape = paper.classList.contains("landscape");
  downloadBlob(packHwpx(buildSection(paperBlocks(paper), landscape)), filename, "application/hwp+zip");
}

export function selfCheck() {
  const table = document.createElement("table");
  table.innerHTML = "<tr><th>A</th><th>B</th></tr><tr><td>1</td><td>2</td></tr>";
  const xml = buildSection(
    [
      { type: "p", text: "check", title: true },
      { type: "table", table },
    ],
    false,
  );
  if (!xml.includes("<hp:tbl") || xml.includes("\n")) throw new Error("hwpx section");
  const zip = packHwpx(xml);
  if (zip[0] !== 0x50 || zip[1] !== 0x4b) throw new Error("hwpx zip");
  const nameLen = zip[26] | (zip[27] << 8);
  const name = new TextDecoder().decode(zip.slice(30, 30 + nameLen));
  if (name !== "mimetype") throw new Error("hwpx mimetype first");
}
