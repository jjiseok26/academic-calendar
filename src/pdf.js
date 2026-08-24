import html2canvas from "html2canvas";
import { downloadBlob } from "./excel.js";

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

function jpegPdf(jpeg, imgW, imgH, landscape) {
  const pageW = landscape ? 841.89 : 595.28;
  const pageH = landscape ? 595.28 : 841.89;
  const enc = new TextEncoder();
  const chunks = [];
  const offsets = [];
  const push = (data) => {
    chunks.push(typeof data === "string" ? enc.encode(data) : data);
  };
  const obj = (n, body, stream) => {
    offsets[n] = chunks.reduce((a, c) => a + c.length, 0);
    if (stream) {
      push(`${n} 0 obj\n${body}\nstream\n`);
      push(stream);
      push("\nendstream\nendobj\n");
    } else {
      push(`${n} 0 obj\n${body}\nendobj\n`);
    }
  };
  push("%PDF-1.4\n");
  obj(1, "<< /Type /Catalog /Pages 2 0 R >>");
  obj(2, "<< /Type /Pages /Kids [3 0 R] /Count 1 >>");
  obj(
    3,
    `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${pageW} ${pageH}] /Resources << /XObject << /Im0 4 0 R >> >> /Contents 5 0 R >>`,
  );
  obj(
    4,
    `<< /Type /XObject /Subtype /Image /Width ${imgW} /Height ${imgH} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${jpeg.length} >>`,
    jpeg,
  );
  const content = enc.encode(`q ${pageW} 0 0 ${pageH} 0 0 cm /Im0 Do Q`);
  obj(5, `<< /Length ${content.length} >>`, content);
  const xrefAt = chunks.reduce((a, c) => a + c.length, 0);
  let xref = `xref\n0 6\n0000000000 65535 f \n`;
  for (let i = 1; i <= 5; i++) xref += `${String(offsets[i]).padStart(10, "0")} 00000 n \n`;
  push(xref);
  push(`trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n${xrefAt}\n%%EOF`);
  return cat(chunks);
}

export async function rasterPaper(paper) {
  const landscape = paper.classList.contains("landscape");
  const host = document.createElement("div");
  host.style.cssText = "position:fixed;left:-9999px;top:0;background:#fff;";
  const clone = paper.cloneNode(true);
  clone.style.zoom = "1";
  clone.style.boxShadow = "none";
  host.appendChild(clone);
  document.body.appendChild(host);
  try {
    const canvas = await html2canvas(clone, {
      scale: 2,
      backgroundColor: "#ffffff",
      useCORS: true,
      logging: false,
    });
    return { canvas, landscape };
  } finally {
    host.remove();
  }
}

export async function downloadPdf(paper, filename) {
  const { canvas, landscape } = await rasterPaper(paper);
  const blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/jpeg", 0.92));
  const jpeg = new Uint8Array(await blob.arrayBuffer());
  downloadBlob(jpegPdf(jpeg, canvas.width, canvas.height, landscape), filename, "application/pdf");
}
