const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

const DEFAULT_DATA_ROOT = path.join(__dirname, "Bauvorhaben");
const DATA_ROOT = (process.env.BAUVORHABEN_ROOT || "").trim() || DEFAULT_DATA_ROOT;
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
  const entries = fs.readdirSync(dir, { withFileTypes: true });
  const fileRe = new RegExp(`^${prefix}(\\d+)\\.xlsx$`, "i");
  const dirRe = new RegExp(`^${prefix}(\\d+)$`, "i");
  const numbers = entries
    .map((entry) => {
      const name = entry.name;
      let match = name.match(fileRe);
      if (!match && entry.isDirectory()) match = name.match(dirRe);
      return match ? Number(match[1]) : null;
    })
    .filter((n) => Number.isFinite(n));
  return numbers.length ? Math.max(...numbers) + 1 : 1;
}

function splitReportLine(raw) {
  let line = String(raw || "");
  if (!line.trim()) return [];

  line = line.replace(
    /(AK:|MAT:|MATERIAL:|MATERIALIEN:|DIENSTLEISTUNG:|DIENSTLEISTUNGEN:|LEISTUNG:|LEISTUNGEN:|MITARBEITER:|ARBEITSKRAEFTE:|ARBEITSKRÄFTE:|ERGEBNIS:|ERGEBNISSE:)/gi,
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

function stripLeadingBullet(text) {
  return String(text || "")
    .replace(/^\s*(?:[-*•]\s*)+/g, "")
    .trim();
}

function normalizeHeaderToken(text) {
  return stripLeadingBullet(text)
    .toLowerCase()
    .replace(/[:.\s]+$/g, "")
    .trim();
}

const SECTION_HEADER_MAP = {
  ak: "arbeitskraefte",
  arbeitskraefte: "arbeitskraefte",
  "arbeitskräfte": "arbeitskraefte",
  mitarbeiter: "arbeitskraefte",
  personal: "arbeitskraefte",
  team: "arbeitskraefte",
  dienstleistung: "leistungen",
  dienstleistungen: "leistungen",
  leistung: "leistungen",
  leistungen: "leistungen",
  ergebnis: "leistungen",
  ergebnisse: "leistungen",
  material: "material",
  materialien: "material",
  mat: "material",
};

function detectSectionHeader(text) {
  const token = normalizeHeaderToken(text);
  return SECTION_HEADER_MAP[token] || null;
}

function stripHeaderPrefix(text) {
  return stripLeadingBullet(text).replace(
    /^(ak|arbeitskraefte|arbeitskräfte|mitarbeiter|personal|team|dienstleistung(?:en)?|leistung(?:en)?|ergebnis(?:se)?|material(?:ien)?|mat)\s*:?\s*/i,
    ""
  );
}

function extractSectionPrefix(text) {
  const value = stripLeadingBullet(text);
  const match = value.match(
    /^(ak|arbeitskraefte|arbeitskräfte|mitarbeiter|personal|team|dienstleistung(?:en)?|leistung(?:en)?|ergebnis(?:se)?|material(?:ien)?|mat)\s*:\s*(.*)$/i
  );
  if (!match) return null;
  const section = detectSectionHeader(match[1]);
  if (!section) return null;
  return { section, content: String(match[2] || "").trim() };
}

function isSectionHeaderLine(text) {
  return Boolean(detectSectionHeader(text));
}

function parseReportLines(lines) {
  const sections = {
    leistungen: [],
    arbeitskraefte: [],
    material: [],
  };
  let currentSection = null;

  for (const raw of lines) {
    const fragments = splitReportLine(raw);
    for (const line of fragments) {
      const trimmed = String(line || "").trim();
      if (!trimmed) continue;
      const prefixed = extractSectionPrefix(trimmed);
      if (prefixed) {
        currentSection = prefixed.section;
        if (prefixed.content) {
          sections[prefixed.section].push(stripLeadingBullet(prefixed.content));
        }
        continue;
      }
      const headerSection = detectSectionHeader(trimmed);
      if (headerSection) {
        currentSection = headerSection;
        const content = stripHeaderPrefix(trimmed).trim();
        if (content) sections[headerSection].push(content);
        continue;
      }

      const cleaned = stripLeadingBullet(trimmed);
      if (isWorkerLine(cleaned)) {
        sections.arbeitskraefte.push(cleaned);
        currentSection = "arbeitskraefte";
        continue;
      }
      if (isMaterialLine(cleaned)) {
        sections.material.push(cleaned);
        currentSection = "material";
        continue;
      }

      if (currentSection === "arbeitskraefte") {
        sections.arbeitskraefte.push(cleaned);
      } else if (currentSection === "material") {
        sections.material.push(cleaned);
      } else {
        sections.leistungen.push(cleaned);
        currentSection = "leistungen";
      }
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

const MATERIAL_UNIT_ALIASES = {
  m: "m",
  mm: "mm",
  cm: "cm",
  dm: "dm",
  km: "km",
  lfm: "lfm",
  qm: "m²",
  "m2": "m²",
  "m²": "m²",
  "m^2": "m²",
  cbm: "m³",
  "m3": "m³",
  "m³": "m³",
  "m^3": "m³",
  l: "l",
  lt: "l",
  liter: "l",
  ml: "ml",
  cl: "cl",
  kg: "kg",
  g: "g",
  t: "t",
  stk: "Stk",
  st: "Stk",
  "stück": "Stk",
  stueck: "Stk",
  pcs: "Stk",
  pc: "Stk",
  paar: "Paar",
  set: "Set",
  satz: "Satz",
  rolle: "Rolle",
  rollen: "Rolle",
  beutel: "Beutel",
  sack: "Sack",
  kiste: "Kiste",
  karton: "Karton",
  palette: "Palette",
  pal: "Palette",
  dose: "Dose",
  paket: "Paket",
  pack: "Pack",
};

const MATERIAL_UNIT_CANONICAL = new Set(
  Object.values(MATERIAL_UNIT_ALIASES).map((u) => u.toLowerCase())
);

function normalizeMaterialUnit(rawUnit) {
  const value = String(rawUnit || "").trim();
  if (!value) return "";
  const key = value
    .toLowerCase()
    .replace(/\./g, "")
    .replace(/\s+/g, "");
  return MATERIAL_UNIT_ALIASES[key] || value;
}

function isKnownMaterialUnit(rawUnit) {
  const normalized = normalizeMaterialUnit(rawUnit);
  if (!normalized) return false;
  return MATERIAL_UNIT_CANONICAL.has(normalized.toLowerCase());
}

function isTimeUnitToken(rawUnit) {
  return /^(h|std\.?|stunden)$/i.test(String(rawUnit || "").trim());
}

function parseMaterialEntries(text) {
  const value = String(text || "").trim();
  if (!value) return [];

  const parts = value.split(";").map((p) => p.trim()).filter(Boolean);
  if (parts.length === 3 && parts[0] && parts[1]) {
    return [
      {
        qty: parts[0],
        unit: normalizeMaterialUnit(parts[1]),
        desc: parts[2],
      },
    ];
  }

  const results = [];
  const itemRegex =
    /(\d+(?:[.,]\d+)?)\s*([A-Za-zÄÖÜäöüßµ²³]+)\s+([^,;]+)/gi;
  let match;
  while ((match = itemRegex.exec(value)) !== null) {
    const unitRaw = match[2];
    if (!isKnownMaterialUnit(unitRaw) || isTimeUnitToken(unitRaw)) continue;
    results.push({
      qty: match[1],
      unit: normalizeMaterialUnit(unitRaw),
      desc: match[3].trim(),
    });
  }

  if (results.length) return results;

  const xRegex = /(\d+(?:[.,]\d+)?)\s*[x×]\s*([^,;]+)/gi;
  while ((match = xRegex.exec(value)) !== null) {
    results.push({
      qty: match[1],
      unit: "Stk",
      desc: match[2].trim(),
    });
  }

  return results;
}

function parseMaterialLine(line) {
  const entries = parseMaterialEntries(line);
  if (entries.length) return entries[0];
  return { qty: "", unit: "", desc: String(line || "").trim() };
}

function normalizeTextForCompare(text) {
  return String(text || "")
    .toLowerCase()
    .replace(/[.,;:()•·]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function isDerivedFromLines(item, lines) {
  const needle = normalizeTextForCompare(item);
  if (!needle) return false;
  return (lines || []).some((line) => {
    const hay = normalizeTextForCompare(line);
    if (!hay) return false;
    return hay.includes(needle) || needle.includes(hay);
  });
}

function sanitizeSections(sections, lines) {
  if (!sections || typeof sections !== "object") return sections;
  const filterItems = (items) =>
    (items || []).filter(
      (item) => isDerivedFromLines(item, lines) && !isSectionHeaderLine(item)
    );
  return {
    leistungen: filterItems(sections.leistungen),
    arbeitskraefte: filterItems(sections.arbeitskraefte),
    material: filterItems(sections.material),
  };
}

function isWorkerLine(text) {
  const value = String(text || "");
  if (!value.trim()) return false;
  const normalizedValue = replaceNumberWords(value);
  const hasWorkerWord =
    /\b(mann|mitarbeiter|arbeiter|monteur|helfer|personal|kolonne|team)\b/i.test(
      value
    );
  const hasHours =
    /\b\d+(?:[.,]\d+)?\s*(?:h|std\.?|stunden)\b/i.test(normalizedValue);
  const hasTime = /\b\d{1,2}[:.]\d{2}\b|\buhr\b/i.test(value);
  const hasArbeitszeit = /\barbeitszeit\b/i.test(value);
  const hasVorOrt = /\bvor\s+ort\b/i.test(value);
  const hasCountRoleJeHours = /\b\d+(?:[.,]\d+)?\s+(?!m[²2]\b)(?:[a-zäöüß][\p{L}äöüß.-]*)(?:\s+[a-zäöüß][\p{L}äöüß.-]*){0,2}\s+je(?:weils)?\s+\d+(?:[.,]\d+)?(?:\s*(?:h|std\.?|stunden))?\b/iu.test(
    normalizedValue
  );
  const hasRoleCountHours = /\b\d+(?:[.,]\d+)?\s+(?!m[²2]\b)(?:[a-zäöüß][\p{L}äöüß.-]*)(?:\s+[a-zäöüß][\p{L}äöüß.-]*){0,2}\b/i.test(
    normalizedValue
  ) && hasHours;
  const hasNamedPeopleWithHours = extractNameHourPairs(value).length > 0;

  if (hasArbeitszeit) return true;
  if (hasWorkerWord && (hasHours || hasTime || hasVorOrt)) return true;
  if (hasCountRoleJeHours) return true;
  if (hasNamedPeopleWithHours) return true;
  if (hasVorOrt && hasRoleCountHours) return true;
  return false;
}

function isServiceLine(text) {
  if (isWorkerLine(text)) return false;
  if (isMaterialLine(text)) return false;
  return true;
}

function isMaterialLine(text) {
  const value = String(text || "");
  if (!value.trim()) return false;
  if (isWorkerLine(value)) return false;
  if (/\b(material|geliefert|lieferung|mat\.?)\b/i.test(value)) return true;
  return parseMaterialEntries(value).length > 0;
}

function normalizeHoursValue(value) {
  if (value === undefined || value === null || value === "") return null;
  const num = Number(String(value).replace(",", "."));
  return Number.isFinite(num) ? num : null;
}

const NUMBER_WORDS = {
  eins: 1,
  eine: 1,
  ein: 1,
  einen: 1,
  einem: 1,
  einer: 1,
  zwei: 2,
  drei: 3,
  vier: 4,
  fuenf: 5,
  sechs: 6,
  sieben: 7,
  acht: 8,
  neun: 9,
  zehn: 10,
  elf: 11,
  zwoelf: 12,
};

const NUMBER_WORD_REGEX =
  /\b(eins|eine|ein|einen|einem|einer|zwei|drei|vier|fuenf|fünf|sechs|sieben|acht|neun|zehn|elf|zwoelf|zwölf)\b/gi;

function normalizeNumberWord(word) {
  return String(word || "")
    .toLowerCase()
    .replace(/ä/g, "ae")
    .replace(/ö/g, "oe")
    .replace(/ü/g, "ue")
    .replace(/ß/g, "ss");
}

function replaceNumberWords(text) {
  return String(text || "").replace(NUMBER_WORD_REGEX, (match) => {
    const key = normalizeNumberWord(match);
    const value = NUMBER_WORDS[key];
    return value === undefined ? match : String(value);
  });
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
  const normalizedText = replaceNumberWords(text);
  let hoursMatch = String(normalizedText || "").match(
    /(\d+(?:[.,]\d+)?)\s*(?:h|std\.?|stunden)\b/i
  );
  let inferredFromJe = false;
  if (!hoursMatch) {
    // Accept shorthand like "2 Installateure je 4" in worker lines.
    hoursMatch = String(normalizedText || "").match(
      /\b(?:je|jeweils)\s+(\d+(?:[.,]\d+)?)\b/i
    );
    inferredFromJe = Boolean(hoursMatch);
  }
  if (!hoursMatch) return { hours: null, applyToAll: false };
  const hours = normalizeHoursValue(hoursMatch[1]);
  const applyToAll =
    inferredFromJe ||
    /\bjeweils\b|\bpro\s+person\b|\bje\s*weils\b/i.test(text || "") ||
    /\bje\s+\d+(?:[.,]\d+)?(?:\s*(?:h|std\.?|stunden))?\b/i.test(
      normalizedText || ""
    );
  return { hours, applyToAll };
}

function normalizeWorkerRoleLabel(roleText) {
  const cleaned = String(roleText || "")
    .trim()
    .replace(/[.,;:]+$/g, "");
  if (!cleaned) return "Mitarbeiter";

  const roleKey = normalizeNumberWord(cleaned)
    .toLowerCase()
    .replace(/\s+/g, " ");

  const roleMap = {
    "männer": "Mann",
    maenner: "Mann",
    mann: "Mann",
    mitarbeiter: "Mitarbeiter",
    arbeiter: "Arbeiter",
    monteur: "Monteur",
    monteure: "Monteur",
    installateur: "Installateur",
    installateure: "Installateur",
    helfer: "Helfer",
    personal: "Personal",
    kolonne: "Kolonne",
    team: "Team",
    ma: "Mitarbeiter",
  };

  const genericRoles = new Set([
    "Mann",
    "Mitarbeiter",
    "Arbeiter",
    "Helfer",
    "Personal",
    "Kolonne",
    "Team",
  ]);

  let roleLabel = roleMap[roleKey];
  if (!roleLabel) {
    roleLabel = cleaned
      .split(/\s+/)
      .map((part) => {
        let token = String(part || "");
        if (/eure$/i.test(token)) {
          token = token.replace(/eure$/i, (m) =>
            m[0] === "E" ? "Eur" : "eur"
          );
        }
        if (/^[a-zäöüß]/.test(token)) {
          token = token.charAt(0).toUpperCase() + token.slice(1);
        }
        return token;
      })
      .join(" ");
  }

  return genericRoles.has(roleLabel) ? "Mitarbeiter" : roleLabel;
}

function sanitizeWorkerDisplayName(name) {
  let value = String(name || "").trim();
  if (!value) return "";

  value = value.replace(
    /\s*(?:,|;|:)?\s*(?:je|jeweils)\s+\d+(?:[.,]\d+)?\s*(?:h|std\.?|stunden)\b\.?\s*$/i,
    ""
  );
  value = value.replace(
    /\s*(?:,|;|:)?\s*\d+(?:[.,]\d+)?\s*(?:h|std\.?|stunden)\b\.?\s*$/i,
    ""
  );

  return value.trim();
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

  const hoursInfo = extractHoursInfo(text);
  const pairs = extractNameHourPairs(text).filter((p) => {
    const name = String(p.name || "").toLowerCase();
    if (!name) return false;
    if (/\bje\b|\bjeweils\b/.test(name)) return false;
    if (
      /\b(mann|männer|maenner|mitarbeiter|arbeiter|monteur|monteure|helfer|personal|kolonne|team)\b/i.test(
        name
      )
    )
      return false;
    return true;
  });
  if (pairs.length) {
    return pairs.map((p) => ({ group, name: p.name, hours: p.hours }));
  }

  const names = extractNames(text);
  if (names.length) {
    const shouldApply = hoursInfo.applyToAll || names.length === 1;
    return names.map((name) => ({
      group,
      name,
      hours: shouldApply ? hoursInfo.hours : null,
    }));
  }

  const normalizedText = replaceNumberWords(text);
  let countMatch = normalizedText.match(
    /(\d+(?:[.,]\d+)?)\s*(?:x|×)?\s*(mann|männer|maenner|mitarbeiter|arbeiter|monteur(?:e)?|helfer|personal|kolonne|team)\b/i
  );
  if (!countMatch) {
    countMatch = normalizedText.match(
      /(\d+(?:[.,]\d+)?)\s*(?:x|×)?\s*([a-zäöüß][\p{L}äöüß.-]*(?:\s+[a-zäöüß][\p{L}äöüß.-]*){0,2})\s*(?=je(?:weils)?\b)/iu
    );
  }
  if (countMatch && hoursInfo.hours !== null) {
    const count = normalizeHoursValue(countMatch[1]);
    if (count !== null && Number.isInteger(count) && count > 0) {
      const baseLabel = normalizeWorkerRoleLabel(countMatch[2]);
      return Array.from({ length: count }, (_, idx) => ({
        group,
        name: `${baseLabel} ${idx + 1}`,
        hours: hoursInfo.hours,
        _noAggregate: true,
      }));
    }
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
  const list = Array.isArray(entries) ? entries : [];
  for (let idx = 0; idx < list.length; idx += 1) {
    const entry = list[idx];
    const name = sanitizeWorkerDisplayName(entry.name);
    if (!name) continue;
    const key = entry._noAggregate
      ? `${name.toLowerCase()}__${idx}`
      : name.toLowerCase();
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
  return sentences.filter((s) => isWorkerLine(s));
}

function extractMaterialLinesFromText(lines) {
  const sentences = splitSentences((lines || []).join(" "));
  return sentences.filter((s) => isMaterialLine(s));
}

function extractLeistungenFromLines(lines) {
  const sentences = splitSentences((lines || []).join(" "));
  return expandLeistungLines(sentences.filter((s) => !isSectionHeaderLine(s)));
}

function expandLeistungLines(lines) {
  const results = [];
  for (const line of lines || []) {
    if (!isServiceLine(line)) continue;
    let cleaned = stripLeadingBullet(String(line || ""))
      .replace(
        /^(dienstleistung(?:en)?|zusatzleistung(?:en)?|leistung(?:en)?|ergebnis(?:se)?)\s*:?\s*/i,
        ""
      )
      .trim();
    cleaned = stripLeadingBullet(cleaned)
      .trim();
    if (!cleaned || isSectionHeaderLine(cleaned)) continue;
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
    const entries = parseMaterialEntries(cleaned);
    if (entries.length) results.push(...entries);
  }
  return results;
}

function expandMaterialLines(lines) {
  const results = [];
  for (const line of lines || []) {
    const entries = parseMaterialEntries(line);
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
  const parsed = parseReportLines(lines);
  const parsedClean = sanitizeSections(parsed, lines) || parsed;
  const parsedIndex = new Set(
    [
      ...(parsedClean.leistungen || []),
      ...(parsedClean.arbeitskraefte || []),
      ...(parsedClean.material || []),
    ].map((item) => normalizeTextForCompare(item))
  );
  let base = parsedClean;
  if (sections) {
    const aiClean = sanitizeSections(sections, lines);
    if (aiClean) {
      const addIfNew = (items) =>
        (items || []).filter(
          (item) => !parsedIndex.has(normalizeTextForCompare(item))
        );
      base = {
        leistungen: mergeUnique(
          parsedClean.leistungen || [],
          addIfNew(aiClean.leistungen)
        ),
        arbeitskraefte: mergeUnique(
          parsedClean.arbeitskraefte || [],
          addIfNew(aiClean.arbeitskraefte)
        ),
        material: mergeUnique(
          parsedClean.material || [],
          addIfNew(aiClean.material)
        ),
      };
    }
  }
  const extraLeistungen = extractLeistungenFromLines(lines);
  const extraMaterial = extractMaterialsFromLines(lines);
  const baseLeistungenRaw = Array.isArray(base.leistungen) ? base.leistungen : [];
  const workerLinesFromLeistungen = baseLeistungenRaw.filter(isWorkerLine);
  const materialLinesFromLeistungen = baseLeistungenRaw.filter(isMaterialLine);
  const baseLeistungen = baseLeistungenRaw.filter(isServiceLine);
  const baseWorkerLines = [
    ...(Array.isArray(base.arbeitskraefte) ? base.arbeitskraefte : []),
    ...workerLinesFromLeistungen,
  ];
  const fallbackWorkerLines = baseWorkerLines.length
    ? []
    : extractWorkerLinesFromText(lines);
  const baseMaterialLines = [
    ...(Array.isArray(base.material) ? base.material : []),
    ...materialLinesFromLeistungen,
  ];
  const fallbackMaterialLines = baseMaterialLines.length
    ? []
    : extractMaterialLinesFromText(lines);
  const rawMaterialLines = mergeUnique(
    [...baseMaterialLines, ...fallbackMaterialLines],
    [],
    normalizeTextForCompare
  );
  const baseMaterial = expandMaterialLines(rawMaterialLines);
  const mergedLeistungen = mergeUnique(baseLeistungen, extraLeistungen);
  const expandedLeistungen = expandLeistungLines(mergedLeistungen);
  const workerLines = mergeUnique(
    [...baseWorkerLines, ...fallbackWorkerLines],
    [],
    normalizeTextForCompare
  );
  const workerLineKeys = new Set(
    workerLines.map((line) => normalizeTextForCompare(line)).filter(Boolean)
  );
  const materialLineKeys = new Set(
    rawMaterialLines.map((line) => normalizeTextForCompare(line)).filter(Boolean)
  );
  const cleanedLeistungen = expandedLeistungen.filter((item) => {
    const key = normalizeTextForCompare(item);
    if (!key) return false;
    if (workerLineKeys.has(key)) return false;
    if (materialLineKeys.has(key)) return false;
    return true;
  });
  return {
    leistungen: mergeUnique(cleanedLeistungen, []),
    arbeitskraefte: workerLines,
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

function cloneStyle(style) {
  if (!style || typeof style !== "object") return {};
  return JSON.parse(JSON.stringify(style));
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
  const leftBaseStyle = cloneStyle(sheet.getCell(`A${startRow}`).style);
  const rightBaseStyle = cloneStyle(sheet.getCell(`G${startRow}`).style);

  ensureLeistungenColumns(
    sheet,
    startRow,
    endRow,
    { startCol: "A", endCol: "F" },
    { startCol: "G", endCol: "K" }
  );

  for (let row = startRow; row <= endRow; row += 1) {
    sheet.getCell(`A${row}`).style = cloneStyle(leftBaseStyle);
    sheet.getCell(`G${row}`).style = cloneStyle(rightBaseStyle);
  }

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
  const reportFolder = path.join(reportDir, `${prefix}${number}`);
  ensureDir(reportFolder);
  const outFile = path.join(reportFolder, `${prefix}${number}.xlsx`);
  const photosDir = path.join(reportFolder, "Fotos");

  if (type === "RB") {
    await fillRegiebericht({ project, reportNumber: number, lines, sections, outFile });
  } else {
    await fillBautagesbericht({ project, reportNumber: number, lines, sections, outFile });
  }

  return { outFile, number, reportFolder, photosDir };
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
