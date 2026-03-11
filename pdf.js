const fs = require("fs");
const path = require("path");
const PDFDocument = require("pdfkit");

// A4 dimensions in points
const PAGE_W = 595.28;
const PAGE_H = 841.89;
const MARGIN_LEFT = 28;
const MARGIN_TOP = 20;
const MARGIN_RIGHT = 28;

// Total content width
const CONTENT_W = PAGE_W - MARGIN_LEFT - MARGIN_RIGHT;

// Column widths proportional to Excel (total Excel width ≈ 92.42 units)
const COL_SCALE = CONTENT_W / 92.42;
const COL_W = {
  A: 6.55 * COL_SCALE,   // ~38
  B: 10.89 * COL_SCALE,  // ~63
  C: 17.89 * COL_SCALE,  // ~104
  D: 8.00 * COL_SCALE,   // ~47
  E: 8.11 * COL_SCALE,   // ~47
  F: 7.33 * COL_SCALE,   // ~43
  G: 7.33 * COL_SCALE,   // ~43
  H: 7.11 * COL_SCALE,   // ~41
  I: 6.55 * COL_SCALE,   // ~38
  J: 6.33 * COL_SCALE,   // ~37
  K: 6.33 * COL_SCALE,   // ~37
};

// Column start positions
const COL_X = {};
{
  let x = MARGIN_LEFT;
  for (const col of "ABCDEFGHIJK".split("")) {
    COL_X[col] = x;
    x += COL_W[col];
  }
}

function colSpanWidth(startCol, endCol) {
  const cols = "ABCDEFGHIJK";
  const si = cols.indexOf(startCol);
  const ei = cols.indexOf(endCol);
  let w = 0;
  for (let i = si; i <= ei; i++) w += COL_W[cols[i]];
  return w;
}

// Colors
const COLOR_BLUE_BG = "#b4c6e7";       // Light blue (theme 3, tint 0.8)
const COLOR_BORDER = "#000000";         // Black thin borders
const COLOR_DOTTED = "#000000";

// Arial font paths (macOS system fonts)
const ARIAL_REGULAR = "/System/Library/Fonts/Supplemental/Arial.ttf";
const ARIAL_BOLD = "/System/Library/Fonts/Supplemental/Arial Bold.ttf";
// Fallback to Helvetica if Arial not available
const HAS_ARIAL = fs.existsSync(ARIAL_REGULAR) && fs.existsSync(ARIAL_BOLD);
const FONT_REGULAR = HAS_ARIAL ? ARIAL_REGULAR : "Helvetica";
const FONT_BOLD = HAS_ARIAL ? ARIAL_BOLD : "Helvetica-Bold";

const LOGO_PATH = path.join(__dirname, "Vorlagen", "logo.png");

function formatDate(d) {
  if (!d) d = new Date();
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yyyy = d.getFullYear();
  return `${dd}.${mm}.${yyyy}`;
}

// Draw a cell with optional fill, border, text
function drawCell(doc, x, y, w, h, opts = {}) {
  // Fill
  if (opts.fill) {
    doc.save().rect(x, y, w, h).fill(opts.fill).restore();
  }
  // Borders
  doc.save().lineWidth(0.5).strokeColor(COLOR_BORDER);
  if (opts.borderTop) { doc.moveTo(x, y).lineTo(x + w, y).stroke(); }
  if (opts.borderBottom) { doc.moveTo(x, y + h).lineTo(x + w, y + h).stroke(); }
  if (opts.borderLeft) { doc.moveTo(x, y).lineTo(x, y + h).stroke(); }
  if (opts.borderRight) { doc.moveTo(x + w, y).lineTo(x + w, y + h).stroke(); }
  if (opts.borderDottedBottom) {
    doc.save().dash(2, { space: 2 });
    doc.moveTo(x, y + h).lineTo(x + w, y + h).stroke();
    doc.undash().restore();
  }
  if (opts.borderAll) {
    doc.rect(x, y, w, h).stroke();
  }
  doc.restore();

  // Text
  if (opts.text !== undefined && opts.text !== null && opts.text !== "") {
    const pad = opts.padding || 3;
    const fontSize = opts.fontSize || 10;
    const font = opts.bold ? FONT_BOLD : FONT_REGULAR;
    const align = opts.align || "left";
    const vAlign = opts.vAlign || "middle";
    doc.font(font).fontSize(fontSize).fillColor(opts.color || "#000000");
    const textH = doc.heightOfString(String(opts.text), { width: w - pad * 2 });
    let textY = y + pad;
    if (vAlign === "middle") textY = y + (h - textH) / 2;
    else if (vAlign === "top") textY = y + pad;
    else if (vAlign === "bottom") textY = y + h - textH - pad;
    if (textY < y + 1) textY = y + 1;
    doc.text(String(opts.text), x + pad, textY, {
      width: w - pad * 2,
      height: h - 2,
      align,
      lineBreak: opts.wrap !== false,
    });
  }
}

/**
 * Generate a PDF report matching the Excel Regiebericht template 1:1.
 */
async function generateReportPDF({ project, reportNumber, sections, photos, outFile }) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ size: "A4", margins: { top: MARGIN_TOP, bottom: 20, left: MARGIN_LEFT, right: MARGIN_RIGHT } });
    const stream = fs.createWriteStream(outFile);
    stream.on("finish", resolve);
    stream.on("error", reject);
    doc.pipe(stream);

    let y = MARGIN_TOP;

    // === ROW 1-2: Title "Regiebericht" (A1:C2) + Logo (G-K, rows 1-2) ===
    const titleW = colSpanWidth("A", "C");
    const titleH = 22.5 + 29.4; // rows 1+2
    drawCell(doc, COL_X.A, y, titleW, titleH, {
      text: "Regiebericht",
      fontSize: 20,
      bold: true,
      vAlign: "middle",
      borderLeft: true,
      borderTop: true,
      borderBottom: true,
    });

    // Logo on the right (G-K)
    if (fs.existsSync(LOGO_PATH)) {
      const logoX = COL_X.G;
      const logoW = colSpanWidth("G", "K");
      try {
        doc.image(LOGO_PATH, logoX, y + 2, {
          width: logoW,
          height: titleH - 4,
          fit: [logoW, titleH - 4],
          align: "right",
          valign: "center",
        });
      } catch (e) { /* ignore logo errors */ }
    }

    // Border around D-K for rows 1-2
    const headerRightX = COL_X.D;
    const headerRightW = colSpanWidth("D", "K");
    doc.save().lineWidth(0.5).strokeColor(COLOR_BORDER);
    doc.moveTo(headerRightX, y).lineTo(headerRightX + headerRightW, y).stroke();
    doc.moveTo(headerRightX + headerRightW, y).lineTo(headerRightX + headerRightW, y + titleH).stroke();
    doc.moveTo(headerRightX, y + titleH).lineTo(headerRightX + headerRightW, y + titleH).stroke();
    doc.restore();

    y += titleH;

    // === ROW 3-6: Company address right-aligned (B3:K6) ===
    const addrH = 17.25 + 15 + 15 + 3.6; // rows 3-6
    const addrW = colSpanWidth("B", "K");
    drawCell(doc, COL_X.A, y, COL_W.A, addrH, {
      borderLeft: true,
    });
    drawCell(doc, COL_X.B, y, addrW, addrH, {
      text: "inntro Mintura GmbH\nInhausen 1a\n85778 Haimhausen",
      fontSize: 10,
      align: "right",
      vAlign: "top",
      padding: 4,
      wrap: true,
      borderRight: true,
    });

    y += addrH;

    // === ROW 7: Labels ===
    const row7H = 15;
    // "AUFTRAGGEBER:" (A7:E7)
    drawCell(doc, COL_X.A, y, colSpanWidth("A", "E"), row7H, {
      text: "AUFTRAGGEBER:",
      fontSize: 10,
      vAlign: "middle",
      borderTop: true,
      borderLeft: true,
    });
    // "BAUVORHABEN:" (F7:I7)
    drawCell(doc, COL_X.F, y, colSpanWidth("F", "I"), row7H, {
      text: "BAUVORHABEN:",
      fontSize: 10,
      vAlign: "middle",
      borderTop: true,
    });
    // "NR.:" (J7)
    drawCell(doc, COL_X.J, y, COL_W.J, row7H, {
      text: "NR.:",
      fontSize: 10,
      vAlign: "middle",
      borderTop: true,
    });
    // Report number (K7:K8) - spans 2 rows
    const nrH = row7H + 15; // row7 + half of row 8-10 area

    y += row7H;

    // === ROW 8-10: Values (3 rows, ~15pt each = 45pt) ===
    const valuesH = 45;
    const projectLabel = project?.name && project?.code
      ? `${project.code}\n${project.name}`
      : project?.name || project?.code || "";

    // Auftraggeber (A8:E10)
    drawCell(doc, COL_X.A, y, colSpanWidth("A", "E"), valuesH, {
      text: project?.client || "",
      fontSize: 11,
      bold: true,
      vAlign: "top",
      padding: 4,
      wrap: true,
      borderLeft: true,
      borderBottom: true,
    });
    // Bauvorhaben (F8:I10)
    drawCell(doc, COL_X.F, y, colSpanWidth("F", "I"), valuesH, {
      text: projectLabel,
      fontSize: 11,
      bold: true,
      vAlign: "top",
      padding: 4,
      wrap: true,
      borderBottom: true,
    });
    // Nr. value (K7:K8) - big number
    drawCell(doc, COL_X.K, y - row7H, COL_W.K, nrH, {
      text: String(reportNumber || ""),
      fontSize: 16,
      bold: true,
      align: "center",
      vAlign: "middle",
      borderRight: true,
    });
    // DATUM label (J9)
    drawCell(doc, COL_X.J, y + 15, COL_W.J, 15, {
      text: "DATUM:",
      fontSize: 10,
      vAlign: "middle",
    });
    // Date value (J10:K10)
    drawCell(doc, COL_X.J, y + 30, colSpanWidth("J", "K"), 15, {
      text: formatDate(new Date()),
      fontSize: 10,
      bold: true,
      align: "center",
      vAlign: "middle",
      borderBottom: true,
      borderRight: true,
    });

    y += valuesH;

    // === ROW 11: "ERBRACHTE LEISTUNGEN:" section header ===
    const sectionH = 15.75;
    drawCell(doc, COL_X.A, y, CONTENT_W, sectionH, {
      text: "ERBRACHTE LEISTUNGEN:",
      fontSize: 10,
      vAlign: "middle",
      fill: COLOR_BLUE_BG,
      borderTop: true,
      borderBottom: true,
      borderLeft: true,
      borderRight: true,
    });

    y += sectionH;

    // === ROWS 12-19: Leistungen (8 rows, two-column layout A:F | G:K) ===
    const leistRowH = 18;
    const leistungen = sections?.leistungen || [];
    const maxLeistRows = 8;
    const halfLen = Math.ceil(Math.min(leistungen.length, maxLeistRows * 2) / 2);
    const leftCol = leistungen.slice(0, Math.min(halfLen, maxLeistRows));
    const rightCol = leistungen.slice(halfLen, halfLen + maxLeistRows);
    const leftW = colSpanWidth("A", "F");
    const rightW = colSpanWidth("G", "K");

    for (let i = 0; i < maxLeistRows; i++) {
      drawCell(doc, COL_X.A, y, leftW, leistRowH, {
        text: leftCol[i] || "",
        fontSize: 10,
        vAlign: "middle",
        borderLeft: true,
        borderDottedBottom: true,
      });
      drawCell(doc, COL_X.G, y, rightW, leistRowH, {
        text: rightCol[i] || "",
        fontSize: 10,
        vAlign: "middle",
        borderDottedBottom: true,
        borderRight: true,
      });
      y += leistRowH;
    }

    // === ROW 20-21: "EINGESETZTE ARBEITSKRÄFTE UND GERÄTE." header ===
    const akHeaderH1 = 15;
    drawCell(doc, COL_X.A, y, colSpanWidth("A", "C"), akHeaderH1, {
      text: "EINGESETZTE ARBEITSKRÄFTE",
      fontSize: 8,
      vAlign: "middle",
      fill: COLOR_BLUE_BG,
      borderTop: true,
      borderLeft: true,
    });
    drawCell(doc, COL_X.D, y, colSpanWidth("D", "K"), akHeaderH1, {
      fill: COLOR_BLUE_BG,
      borderTop: true,
      borderRight: true,
    });

    y += akHeaderH1;

    const akHeaderH2 = 12.75;
    drawCell(doc, COL_X.A, y, colSpanWidth("A", "C"), akHeaderH2, {
      text: "UND GERÄTE.",
      fontSize: 8,
      vAlign: "middle",
      fill: COLOR_BLUE_BG,
      borderLeft: true,
    });
    drawCell(doc, COL_X.D, y, colSpanWidth("D", "I"), akHeaderH2, {
      text: "ARBEITSZEIT",
      fontSize: 8,
      align: "center",
      vAlign: "middle",
      fill: COLOR_BLUE_BG,
    });
    drawCell(doc, COL_X.J, y, colSpanWidth("J", "K"), akHeaderH2, {
      fill: COLOR_BLUE_BG,
      borderRight: true,
    });

    y += akHeaderH2;

    // === ROW 22: AK column headers ===
    const akColHeaderH = 33;
    const akCols = [
      { col: "A", text: "STD.", fill: COLOR_BLUE_BG },
      { col: "B", text: "BESCHÄFTI-\nGUNGSGRUPPE\n(Pos. LV)", fill: COLOR_BLUE_BG },
      { col: "C", text: "NAME", fill: COLOR_BLUE_BG },
      { col: "D", text: "VON:" },
      { col: "E", text: "BIS:" },
      { col: "F", text: "VON:" },
      { col: "G", text: "BIS:" },
      { col: "H", text: "VON:" },
      { col: "I", text: "BIS:" },
    ];
    for (const c of akCols) {
      drawCell(doc, COL_X[c.col], y, COL_W[c.col], akColHeaderH, {
        text: c.text,
        fontSize: 6,
        align: "center",
        vAlign: "middle",
        wrap: true,
        fill: c.fill,
        borderAll: true,
      });
    }
    drawCell(doc, COL_X.J, y, colSpanWidth("J", "K"), akColHeaderH, {
      text: "AUFZAHLUNGEN\nÜBERSTD.\nERSCHWERN.\nUsw.",
      fontSize: 6,
      align: "center",
      vAlign: "middle",
      wrap: true,
      fill: COLOR_BLUE_BG,
      borderAll: true,
    });

    y += akColHeaderH;

    // === ROWS 23-28: Worker data (6 rows) ===
    const workerRowH = 18;
    const workers = sections?.arbeitskraefte || [];
    for (let i = 0; i < 6; i++) {
      const w = workers[i];
      // STD
      drawCell(doc, COL_X.A, y, COL_W.A, workerRowH, {
        text: w && w.hasHours ? String(w.hours) : "",
        fontSize: 10,
        align: "center",
        vAlign: "middle",
        borderLeft: true,
        borderRight: true,
        borderBottom: true,
      });
      // Gruppe
      drawCell(doc, COL_X.B, y, COL_W.B, workerRowH, {
        text: w?.group || "",
        fontSize: 10,
        align: "center",
        vAlign: "middle",
        borderRight: true,
        borderBottom: true,
      });
      // Name
      drawCell(doc, COL_X.C, y, COL_W.C, workerRowH, {
        text: w?.name || "",
        fontSize: 10,
        align: "center",
        vAlign: "middle",
        borderRight: true,
        borderBottom: true,
      });
      // Time slots D-I (empty)
      for (const col of "DEFGHI".split("")) {
        drawCell(doc, COL_X[col], y, COL_W[col], workerRowH, {
          borderRight: true,
          borderBottom: true,
        });
      }
      // J-K (empty)
      drawCell(doc, COL_X.J, y, colSpanWidth("J", "K"), workerRowH, {
        borderRight: true,
        borderBottom: true,
      });
      y += workerRowH;
    }

    // === ROW 29: empty spacer row ===
    y += 18;

    // === ROW 30: Material section header ===
    const matHeaderH = 27.75;
    drawCell(doc, COL_X.A, y, CONTENT_W, matHeaderH, {
      text: "BEISTELLUNG VON BAU-, HILFS-, NEBENSTOFFEN, GERÄTEN U. BETRIEBSSTOFFEN, FREMDLEISTUNGEN u. SONSTIGE KOSTEN",
      fontSize: 8,
      align: "center",
      vAlign: "middle",
      fill: COLOR_BLUE_BG,
      borderTop: true,
      borderBottom: true,
      borderLeft: true,
      borderRight: true,
    });

    y += matHeaderH;

    // === ROW 31: Material column headers (two halves) ===
    const matColHeaderH = 21;
    // Left half: A=MENGE, B=EINHEIT, C-D=BEZEICHNUNG
    drawCell(doc, COL_X.A, y, COL_W.A, matColHeaderH, {
      text: "MENGE", fontSize: 6, align: "center", vAlign: "middle",
      fill: COLOR_BLUE_BG, borderAll: true,
    });
    drawCell(doc, COL_X.B, y, COL_W.B, matColHeaderH, {
      text: "EINHEIT", fontSize: 6, align: "center", vAlign: "middle",
      fill: COLOR_BLUE_BG, borderAll: true,
    });
    drawCell(doc, COL_X.C, y, colSpanWidth("C", "D"), matColHeaderH, {
      text: "BEZEICHNUNG", fontSize: 6, align: "center", vAlign: "middle",
      fill: COLOR_BLUE_BG, borderAll: true,
    });
    // Right half: E=MENGE, F-G=EINHEIT, H-K=BEZEICHNUNG
    drawCell(doc, COL_X.E, y, COL_W.E, matColHeaderH, {
      text: "MENGE", fontSize: 6, align: "center", vAlign: "middle",
      fill: COLOR_BLUE_BG, borderAll: true,
    });
    drawCell(doc, COL_X.F, y, colSpanWidth("F", "G"), matColHeaderH, {
      text: "EINHEIT", fontSize: 6, align: "center", vAlign: "middle",
      fill: COLOR_BLUE_BG, borderAll: true,
    });
    drawCell(doc, COL_X.H, y, colSpanWidth("H", "K"), matColHeaderH, {
      text: "BEZEICHNUNG", fontSize: 6, align: "center", vAlign: "middle",
      fill: COLOR_BLUE_BG, borderAll: true,
    });

    y += matColHeaderH;

    // === ROWS 32-38: Material data (7 rows, two halves) ===
    const matRowH = 18;
    const materials = sections?.material || [];
    for (let i = 0; i < 7; i++) {
      const m = materials[i];
      const m2 = materials[i + 7]; // second half (right side)
      // Left: A=qty, B=unit, C-D=desc
      drawCell(doc, COL_X.A, y, COL_W.A, matRowH, {
        text: m?.qty || "", fontSize: 10, align: "center", vAlign: "middle",
        borderLeft: true, borderRight: true, borderBottom: true,
      });
      drawCell(doc, COL_X.B, y, COL_W.B, matRowH, {
        text: m?.unit || "", fontSize: 10, align: "center", vAlign: "middle",
        borderRight: true, borderBottom: true,
      });
      drawCell(doc, COL_X.C, y, colSpanWidth("C", "D"), matRowH, {
        text: m?.desc || "", fontSize: 10, align: "center", vAlign: "middle",
        borderRight: true, borderBottom: true,
      });
      // Right: E=qty, F-G=unit, H-K=desc
      drawCell(doc, COL_X.E, y, COL_W.E, matRowH, {
        text: m2?.qty || "", fontSize: 10, align: "center", vAlign: "middle",
        borderRight: true, borderBottom: true,
      });
      drawCell(doc, COL_X.F, y, colSpanWidth("F", "G"), matRowH, {
        text: m2?.unit || "", fontSize: 10, align: "center", vAlign: "middle",
        borderRight: true, borderBottom: true,
      });
      drawCell(doc, COL_X.H, y, colSpanWidth("H", "K"), matRowH, {
        text: m2?.desc || "", fontSize: 10, align: "center", vAlign: "middle",
        borderRight: true, borderBottom: true,
      });
      y += matRowH;
    }

    // === ROW 39: Blue separator bar ===
    const sepH = 32.25;
    drawCell(doc, COL_X.A, y, CONTENT_W, sepH, {
      fill: COLOR_BLUE_BG,
      borderTop: true,
      borderBottom: true,
      borderLeft: true,
      borderRight: true,
    });

    y += sepH;

    // === ROW 40: Signature area ===
    const sigH = 20.25;
    drawCell(doc, COL_X.A, y, colSpanWidth("A", "D"), sigH, {
      text: "AUFTRAGNEHMER:",
      fontSize: 10,
      vAlign: "top",
      padding: 4,
      borderLeft: true,
      borderTop: true,
    });
    drawCell(doc, COL_X.E, y, colSpanWidth("E", "K"), sigH, {
      text: "ÖBA:",
      fontSize: 10,
      vAlign: "top",
      padding: 4,
      borderTop: true,
      borderRight: true,
    });

    // === PHOTO PAGES ===
    const photoFiles = (photos || []).filter((p) => fs.existsSync(p));
    for (let i = 0; i < photoFiles.length; i++) {
      doc.addPage({ size: "A4", margin: 40 });
      const photoMargin = 40;

      doc
        .font(FONT_BOLD)
        .fontSize(14)
        .fillColor("#000000")
        .text(`Foto ${i + 1}`, photoMargin, photoMargin, {
          align: "center",
          width: PAGE_W - photoMargin * 2,
        });

      const imgY = photoMargin + 30;
      const maxW = PAGE_W - photoMargin * 2;
      const maxH = PAGE_H - photoMargin * 2 - 40;

      try {
        const img = doc.openImage(photoFiles[i]);
        const scale = Math.min(maxW / img.width, maxH / img.height, 1);
        const w = img.width * scale;
        const h = img.height * scale;
        const x = photoMargin + (maxW - w) / 2;
        doc.image(img, x, imgY, { width: w, height: h });
      } catch (imgErr) {
        doc
          .font(FONT_REGULAR)
          .fontSize(10)
          .text(
            `Foto konnte nicht geladen werden: ${path.basename(photoFiles[i])}`,
            photoMargin,
            imgY
          );
      }
    }

    doc.end();
  });
}

module.exports = { generateReportPDF };
