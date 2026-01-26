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

function fillRange(sheet, startRow, endRow, col, values) {
  let idx = 0;
  for (let row = startRow; row <= endRow; row += 1) {
    const cell = sheet.getCell(`${col}${row}`);
    cell.value = values[idx] || "";
    idx += 1;
  }
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
    sheet.getCell("F8").value = project.name || project.code;
  }

  const finalSections = sections || parseReportLines(lines);
  fillRange(sheet, 12, 19, "A", finalSections.leistungen);

  const workers = finalSections.arbeitskraefte.map(parseWorkerLine);
  for (let i = 0; i < 6; i += 1) {
    const row = 23 + i;
    const item = workers[i];
    sheet.getCell(`B${row}`).value = item ? item.group : "";
    sheet.getCell(`C${row}`).value = item ? item.name : "";
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
    sheet.getCell("C3").value = project.name || project.code;
  }

  const finalSections = sections || parseReportLines(lines);
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
