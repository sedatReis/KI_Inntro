const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

const DATA_ROOT = path.join(__dirname, "Bauvorhaben");
const META_DIR = path.join(DATA_ROOT, "_meta");
const PROJECTS_JSON = path.join(META_DIR, "projects.json");
const PROJECTS_XLSX = path.join(META_DIR, "Bauvorhaben.xlsx");
const TEMPLATE_RB = path.join(__dirname, "Vorlagen", "RB Vorlage.xlsx");
const TEMPLATE_BTB = path.join(__dirname, "Vorlagen", "BTB Vorlage.xlsx");

function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true });
}

function ensureBaseStructure() {
  ensureDir(DATA_ROOT);
  ensureDir(META_DIR);
}

function safeDirName(value) {
  return String(value || "")
    .trim()
    .replace(/[\\/:"*?<>|]+/g, "_")
    .replace(/\s+/g, "_");
}

function loadProjects() {
  if (!fs.existsSync(PROJECTS_JSON)) return [];
  try {
    const raw = fs.readFileSync(PROJECTS_JSON, "utf8");
    const data = JSON.parse(raw);
    return Array.isArray(data) ? data : [];
  } catch (err) {
    return [];
  }
}

async function writeProjectsWorkbook(projects) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Bauvorhaben");
  ws.addRow([
    "kuerzel",
    "projekt",
    "ansprechperson",
    "email",
    "cc",
    "auftraggeber",
    "created_at",
  ]);
  for (const p of projects) {
    ws.addRow([
      p.code || "",
      p.name || "",
      p.contactName || "",
      p.contactEmail || "",
      p.cc || "",
      p.client || "",
      p.createdAt || "",
    ]);
  }
  await wb.xlsx.writeFile(PROJECTS_XLSX);
}

async function saveProjects(projects) {
  ensureBaseStructure();
  fs.writeFileSync(PROJECTS_JSON, JSON.stringify(projects, null, 2), "utf8");
  await writeProjectsWorkbook(projects);
}

async function upsertProject(project) {
  ensureBaseStructure();
  const projects = loadProjects();
  const key = (project.code || "").trim().toLowerCase();
  const dirName = safeDirName(project.code);
  const nextProject = { ...project, dirName };
  const existingIdx = projects.findIndex((p) => (p.code || "").toLowerCase() === key);
  if (existingIdx >= 0) {
    projects[existingIdx] = { ...projects[existingIdx], ...nextProject };
  } else {
    projects.push(nextProject);
  }
  await saveProjects(projects);
  ensureProjectStructure(nextProject);
  return nextProject;
}

function getProjectByCode(code) {
  const projects = loadProjects();
  const key = (code || "").trim().toLowerCase();
  return projects.find((p) => (p.code || "").toLowerCase() === key);
}

function ensureProjectStructure(projectOrCode) {
  ensureBaseStructure();
  const code =
    typeof projectOrCode === "string" ? projectOrCode : projectOrCode?.code;
  const dirName =
    typeof projectOrCode === "object" && projectOrCode?.dirName
      ? projectOrCode.dirName
      : safeDirName(code);
  const projectDir = path.join(DATA_ROOT, dirName);
  ensureDir(projectDir);
  ensureDir(path.join(projectDir, "Regieberichte"));
  ensureDir(path.join(projectDir, "Bautageberichte"));
  return projectDir;
}

function getNextReportNumber(dir, prefix) {
  if (!fs.existsSync(dir)) return 1;
  const files = fs.readdirSync(dir);
  const re = new RegExp(`^${prefix}(\\d+)\\.xlsx$`, "i");
  const numbers = files
    .map((f) => {
      const match = f.match(re);
      return match ? Number(match[1]) : null;
    })
    .filter((n) => Number.isFinite(n));
  return numbers.length ? Math.max(...numbers) + 1 : 1;
}

function splitReportLine(raw) {
  let line = String(raw || "");
  if (!line.trim()) return [];

  line = line.replace(
    /(AK:|MAT:|MATERIAL:|LEISTUNG:|LEISTUNGEN:|ERGEBNIS:|ERGEBNISSE:)/gi,
    "\n$1"
  );
  line = line.replace(/\bAK\b(?!\s*:)/gi, "\nAK:");
  line = line.replace(/\bMAT\b(?!\s*:)/gi, "\nMAT:");
  line = line.replace(/\bMATERIAL\b(?!\s*:)/gi, "\nMAT:");
  line = line.replace(/\bLEISTUNG\b(?!\s*:)/gi, "\nLeistung:");
  line = line.replace(/\bLEISTUNGEN\b(?!\s*:)/gi, "\nLeistung:");
  line = line.replace(/\bERGEBNIS\b(?!\s*:)/gi, "\nErgebnis:");
  line = line.replace(/\bERGEBNISSE\b(?!\s*:)/gi, "\nErgebnis:");

  return line
    .split(/\r?\n/)
    .map((part) => part.trim())
    .filter(Boolean);
}

function parseReportLines(lines) {
  const sections = {
    leistungen: [],
    arbeitskraefte: [],
    material: [],
  };

  for (const raw of lines) {
    const fragments = splitReportLine(raw);
    for (const line of fragments) {
    const lower = line.toLowerCase();
    if (lower.startsWith("ak:") || lower.startsWith("ak ")) {
      sections.arbeitskraefte.push(line.replace(/^ak[:\s]*/i, ""));
      continue;
    }
    if (
      lower.startsWith("mat:") ||
      lower.startsWith("material:") ||
      lower.startsWith("material ")
    ) {
      sections.material.push(line.replace(/^(mat|material)[:\s]*/i, ""));
      continue;
    }
    if (
      lower.startsWith("lei:") ||
      lower.startsWith("leistung:") ||
      lower.startsWith("leistungen:") ||
      lower.startsWith("ergebnis:") ||
      lower.startsWith("ergebnisse:")
    ) {
      sections.leistungen.push(line.replace(/^(lei|leistung|leistungen|ergebnis|ergebnisse)[:\s]*/i, ""));
      continue;
    }
      sections.leistungen.push(line);
    }
  }

  return sections;
}

function parseWorkerLine(line) {
  const parts = line.split(/[;,-]/).map((p) => p.trim()).filter(Boolean);
  if (parts.length >= 2) {
    return { group: parts[0], name: parts.slice(1).join(" ") };
  }
  return { group: "", name: line.trim() };
}

function parseMaterialLine(line) {
  const parts = line.split(";").map((p) => p.trim()).filter(Boolean);
  if (parts.length >= 3) {
    return { qty: parts[0], unit: parts[1], desc: parts.slice(2).join(" ") };
  }
  const m = line.match(/^(\d+(?:[.,]\d+)?)\s+([A-Za-z]+)\s+(.+)$/);
  if (m) {
    return { qty: m[1], unit: m[2], desc: m[3] };
  }
  return { qty: "", unit: "", desc: line.trim() };
}

function normalizeTextForCompare(text) {
  return String(text || "")
    .toLowerCase()
    .replace(/[.,;:()]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function isServiceLine(text) {
  return !/\b(arbeitszeit|pause|stunden|monteur|mitarbeiter|vor\s+ort|uhr|material|geliefert|lieferung|mat\.?)\b/i.test(
    text || ""
  );
}

function isMaterialLine(text) {
  return /\b(material|geliefert|lieferung|mat\.?)\b/i.test(text || "");
}

function normalizeHoursValue(value) {
  if (value === undefined || value === null || value === "") return null;
  const num = Number(String(value).replace(",", "."));
  return Number.isFinite(num) ? num : null;
}

function extractNameHourPairs(text) {
  const pairs = [];
  const regex =
    /(\b\p{Lu}[\p{L}.'-]+(?:\s+\p{Lu}[\p{L}.'-]+)+)\s*(?:\(|,|:)?\s*(\d+(?:[.,]\d+)?)\s*(?:h|std\.?|stunden)\b/giu;
  let match;
  while ((match = regex.exec(text)) !== null) {
    pairs.push({
      name: match[1].trim(),
      hours: normalizeHoursValue(match[2]),
    });
  }
  return pairs;
}

function extractNames(text) {
  let subject = String(text || "");
  if (subject.includes(":")) {
    subject = subject.split(":").slice(1).join(":");
  }
  const vorOrtIdx = subject.search(/\bvor\s+ort\b/i);
  if (vorOrtIdx >= 0) {
    subject = subject.slice(vorOrtIdx + 7);
  }
  const cutIdx = subject.search(/\bjeweils\b|\bstunden\b|\bstd\.?\b|\bh\b/i);
  if (cutIdx >= 0) {
    subject = subject.slice(0, cutIdx);
  }
  subject = subject.replace(/\bund\b/gi, ",");
  subject = subject.replace(/[\/&]/g, ",");
  const tokens = subject
    .split(",")
    .map((t) => t.trim())
    .filter(Boolean);

  const nameRegex =
    /^[A-ZÄÖÜ][A-Za-zÄÖÜäöüß.'-]+(?:\s+[A-ZÄÖÜ][A-Za-zÄÖÜäöüß.'-]+)+$/;

  return tokens.filter((t) => nameRegex.test(t));
}

function extractHoursInfo(text) {
  const hoursMatch = String(text || "").match(
    /(\d+(?:[.,]\d+)?)\s*(?:h|std\.?|stunden)\b/i
  );
  if (!hoursMatch) return { hours: null, applyToAll: false };
  const hours = normalizeHoursValue(hoursMatch[1]);
  const applyToAll = /\bjeweils\b|\bpro\s+person\b|\bje\s*weils\b/i.test(
    text || ""
  );
  return { hours, applyToAll };
}

function parseWorkerEntries(line) {
  if (!line || !String(line).trim()) return [];
  const parts = line.split(";").map((p) => p.trim()).filter(Boolean);
  let group = "";
  let text = String(line || "");
  if (parts.length >= 2) {
    group = parts[0];
    text = parts.slice(1).join(" ");
  }

  const pairs = extractNameHourPairs(text);
  if (pairs.length) {
    return pairs.map((p) => ({ group, name: p.name, hours: p.hours }));
  }

  const names = extractNames(text);
  const hoursInfo = extractHoursInfo(text);
  if (names.length) {
    const shouldApply = hoursInfo.applyToAll || names.length === 1;
    return names.map((name) => ({
      group,
      name,
      hours: shouldApply ? hoursInfo.hours : null,
    }));
  }

  return [
    {
      group,
      name: text.trim(),
      hours: hoursInfo.hours,
    },
  ];
}

function aggregateWorkers(entries) {
  const map = new Map();
  for (const entry of entries || []) {
    const name = String(entry.name || "").trim();
    if (!name) continue;
    const key = name.toLowerCase();
    const existing = map.get(key) || {
      name,
      group: String(entry.group || "").trim(),
      hours: 0,
      hasHours: false,
    };
    const hours = normalizeHoursValue(entry.hours);
    if (hours !== null) {
      existing.hours += hours;
      existing.hasHours = true;
    }
    if (!existing.group && entry.group) {
      existing.group = String(entry.group || "").trim();
    }
    map.set(key, existing);
  }
  return Array.from(map.values());
}

function splitSentences(text) {
  return String(text || "")
    .replace(/\r?\n/g, ". ")
    .split(/[.!?]+/g)
    .map((s) => s.trim())
    .filter(Boolean);
}

function extractWorkerLinesFromText(lines) {
  const sentences = splitSentences((lines || []).join(" "));
  return sentences.filter((s) =>
    /\bstunden\b|\barbeitszeit\b|\bmonteur\b|\bmitarbeiter\b|\bvor\s+ort\b/i.test(
      s
    )
  );
}

function extractLeistungenFromLines(lines) {
  const sentences = splitSentences((lines || []).join(" "));
  return expandLeistungLines(sentences);
}

function expandLeistungLines(lines) {
  const results = [];
  for (const line of lines || []) {
    if (!isServiceLine(line)) continue;
    let cleaned = String(line || "")
      .replace(/^(zusatzleistung|leistung|leistungen|ergebnis|ergebnisse)\s*:\s*/i, "")
      .trim();
    if (!cleaned) continue;
    const parts = cleaned
      .split(",")
      .map((p) => p.trim())
      .filter(Boolean);
    for (const part of parts) {
      if (part) results.push(part);
    }
  }
  return results;
}

function extractMaterialsFromLines(lines) {
  const sentences = splitSentences((lines || []).join(" "));
  const results = [];
  for (const sentence of sentences) {
    if (!isMaterialLine(sentence)) continue;
    const cleaned = sentence.replace(/^(material\s+geliefert|material)\s*:\s*/i, "");
    const regex =
      /(\d+(?:[.,]\d+)?)\s*([A-Za-zÄÖÜäöüßµ²³mM]+)\s+([^,;]+)/g;
    let match;
    while ((match = regex.exec(cleaned)) !== null) {
      results.push({
        qty: match[1],
        unit: match[2],
        desc: match[3].trim(),
      });
    }
  }
  return results;
}

function expandMaterialLines(lines) {
  const results = [];
  for (const line of lines || []) {
    const entries = extractMaterialsFromLines([line]);
    if (entries.length) {
      results.push(
        ...entries.map((m) => `${m.qty}; ${m.unit}; ${m.desc}`)
      );
    } else if (String(line || "").trim()) {
      results.push(String(line || "").trim());
    }
  }
  return results;
}

function mergeUnique(listA, listB, normalizer = normalizeTextForCompare) {
  const seen = new Set();
  const out = [];
  for (const item of [...(listA || []), ...(listB || [])]) {
    const text = String(item || "").trim();
    if (!text) continue;
    const key = normalizer(text);
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(text);
  }
  return out;
}

function normalizeReportSections(lines, sections) {
  const base = sections || parseReportLines(lines);
  const extraLeistungen = extractLeistungenFromLines(lines);
  const extraMaterial = extractMaterialsFromLines(lines);
  const baseLeistungen = (base.leistungen || []).filter(isServiceLine);
  const baseMaterial = expandMaterialLines(base.material || []);
  const mergedLeistungen = mergeUnique(baseLeistungen, extraLeistungen);
  const expandedLeistungen = expandLeistungLines(mergedLeistungen);
  return {
    leistungen: mergeUnique(expandedLeistungen, []),
    arbeitskraefte: Array.isArray(base.arbeitskraefte) ? base.arbeitskraefte : [],
    material: mergeUnique(
      baseMaterial,
      extraMaterial.map((m) => `${m.qty}; ${m.unit}; ${m.desc}`),
      (text) => normalizeTextForCompare(text)
    ),
  };
}

function fillRange(sheet, startRow, endRow, col, values) {
  let idx = 0;
  for (let row = startRow; row <= endRow; row += 1) {
    const cell = sheet.getCell(`${col}${row}`);
    cell.value = values[idx] || "";
    idx += 1;
  }
}

function ensureLeistungenColumns(sheet, startRow, endRow, leftRange, rightRange) {
  for (let row = startRow; row <= endRow; row += 1) {
    const mergedAddress = `${leftRange.startCol}${row}:${rightRange.endCol}${row}`;
    try {
      sheet.unMergeCells(mergedAddress);
    } catch (err) {
      // ignore if not merged
    }
    sheet.mergeCells(`${leftRange.startCol}${row}:${leftRange.endCol}${row}`);
    sheet.mergeCells(`${rightRange.startCol}${row}:${rightRange.endCol}${row}`);
  }
}

function fillLeistungenTwoColumns(sheet, startRow, endRow, values) {
  const maxRows = endRow - startRow + 1;
  const leftValues = values.slice(0, maxRows);
  const rightValues = values.slice(maxRows, maxRows * 2);

  ensureLeistungenColumns(
    sheet,
    startRow,
    endRow,
    { startCol: "A", endCol: "F" },
    { startCol: "G", endCol: "K" }
  );

  fillRange(sheet, startRow, endRow, "A", leftValues);
  fillRange(sheet, startRow, endRow, "G", rightValues);
}

async function fillRegiebericht({ project, reportNumber, lines, sections, outFile }) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(TEMPLATE_RB);
  const sheet = wb.getWorksheet("Regiebericht") || wb.worksheets[0];
  if (!sheet) {
    throw new Error("RB Vorlage: kein Worksheet gefunden.");
  }

  if (reportNumber) sheet.getCell("K7").value = Number(reportNumber);
  sheet.getCell("J10").value = new Date();
  if (project?.client) sheet.getCell("A8").value = project.client;
  if (project?.name || project?.code) {
    const label =
      project?.name && project?.code
        ? `${project.code} - ${project.name}`
        : project?.name || project?.code;
    sheet.getCell("F8").value = label;
  }

  const finalSections = normalizeReportSections(lines, sections);
  fillLeistungenTwoColumns(sheet, 12, 19, finalSections.leistungen);

  const rawWorkerLines = finalSections.arbeitskraefte.length
    ? finalSections.arbeitskraefte
    : extractWorkerLinesFromText(lines);
  const workerEntries = rawWorkerLines.flatMap(parseWorkerEntries);
  const workers = aggregateWorkers(workerEntries);

  for (let i = 0; i < 6; i += 1) {
    const row = 23 + i;
    const item = workers[i];
    sheet.getCell(`A${row}`).value = item && item.hasHours ? item.hours : "";
    sheet.getCell(`B${row}`).value = item ? item.group : "";
    sheet.getCell(`C${row}`).value = item ? item.name : "";
    sheet.getCell(`D${row}`).value = "";
    sheet.getCell(`E${row}`).value = "";
    sheet.getCell(`F${row}`).value = "";
    sheet.getCell(`G${row}`).value = "";
    sheet.getCell(`H${row}`).value = "";
    sheet.getCell(`I${row}`).value = "";
  }

  const materials = finalSections.material.map(parseMaterialLine);
  for (let i = 0; i < 7; i += 1) {
    const row = 32 + i;
    const item = materials[i];
    sheet.getCell(`A${row}`).value = item ? item.qty : "";
    sheet.getCell(`B${row}`).value = item ? item.unit : "";
    sheet.getCell(`C${row}`).value = item ? item.desc : "";
  }

  await wb.xlsx.writeFile(outFile);
}

async function fillBautagesbericht({ project, reportNumber, lines, sections, outFile }) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(TEMPLATE_BTB);
  const sheet = wb.worksheets[0];
  if (!sheet) {
    throw new Error("BTB Vorlage: kein Worksheet gefunden.");
  }

  if (reportNumber) sheet.getCell("J2").value = String(reportNumber);
  sheet.getCell("Q2").value = new Date();
  if (project?.name || project?.code) {
    const label =
      project?.name && project?.code
        ? `${project.code} - ${project.name}`
        : project?.name || project?.code;
    sheet.getCell("C3").value = label;
  }

  const finalSections = normalizeReportSections(lines, sections);
  fillRange(sheet, 19, 25, "A", finalSections.leistungen);
  sheet.getCell("L45").value = new Date();

  await wb.xlsx.writeFile(outFile);
}

async function generateReportFile({ type, project, lines, sections }) {
  const projectDir = ensureProjectStructure(project);
  const reportDir =
    type === "RB"
      ? path.join(projectDir, "Regieberichte")
      : path.join(projectDir, "Bautageberichte");
  const prefix = type === "RB" ? "RB" : "BTB";
  const number = getNextReportNumber(reportDir, prefix);
  const outFile = path.join(reportDir, `${prefix}${number}.xlsx`);

  if (type === "RB") {
    await fillRegiebericht({ project, reportNumber: number, lines, sections, outFile });
  } else {
    await fillBautagesbericht({ project, reportNumber: number, lines, sections, outFile });
  }

  return { outFile, number };
}

module.exports = {
  DATA_ROOT,
  ensureBaseStructure,
  ensureProjectStructure,
  loadProjects,
  upsertProject,
  getProjectByCode,
  generateReportFile,
};
