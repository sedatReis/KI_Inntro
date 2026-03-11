require("dotenv").config();
const { Telegraf } = require("telegraf");
const Database = require("better-sqlite3");
const nodemailer = require("nodemailer");
require("isomorphic-fetch");
const { Client } = require("@microsoft/microsoft-graph-client");
const { ClientSecretCredential } = require("@azure/identity");
const fs = require("fs");
const path = require("path");
const { generateReport, chatProjectAI, generateEmailDraft, classifyReportLines, chatReportAI, extractRuleFromCorrection, buildReportSystemPrompt, detectIntent } = require("./ai");

const fetch = (...args) =>
  import("node-fetch").then(({ default: fetchImpl }) => fetchImpl(...args));
const {
  DATA_ROOT,
  ensureBaseStructure,
  upsertProject,
  getProjectByCode,
  generateReportFile,
  buildReportSectionsFromAIData,
  parseWorkerEntries,
  aggregateWorkers,
  parseMaterialLine,
  normalizeReportSections,
} = require("./reports");
const { generateReportPDF } = require("./pdf");

const BOT_TOKEN = process.env.BOT_TOKEN;
const OWNER_USER_ID = Number(process.env.OWNER_USER_ID);
const SMTP_HOST = process.env.SMTP_HOST;
const SMTP_PORT = Number(process.env.SMTP_PORT || 587);
const SMTP_USER = process.env.SMTP_USER;
const SMTP_PASS = process.env.SMTP_PASS;
const SMTP_FROM = process.env.SMTP_FROM || SMTP_USER;
const DEFAULT_EMAIL_TO = process.env.EMAIL_TO;
const AZURE_TENANT_ID = process.env.AZURE_TENANT_ID;
const AZURE_CLIENT_ID = process.env.AZURE_CLIENT_ID;
const AZURE_CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;
const NAS_SHARE = process.env.NAS_SHARE || "//admin@192.168.2.126/server";
const NAS_PASS = process.env.NAS_PASS || "";
const NAS_MOUNT = "/tmp/nas_server";
const NAS_REPORTS_PATH = "Projekte/Aufträge";
const REPORT_EMAIL_SUBJECT = process.env.REPORT_EMAIL_SUBJECT || "Regiebericht {reportNumber} – {projectName}";
const REPORT_EMAIL_BODY = process.env.REPORT_EMAIL_BODY || "Sehr geehrte/r {contactName},\\n\\nanbei erhalten Sie den Regiebericht...";
const META_DIR = path.join(DATA_ROOT, "_meta");
const CONTACTS_FILE = path.join(META_DIR, "contacts.csv");
const EMAIL_LOG_FILE = path.join(META_DIR, "email_logs.csv");

const ALLOWED_GROUP_IDS = (process.env.ALLOWED_GROUP_IDS || "")
  .split(",")
  .map((s) => s.trim())
  .filter(Boolean)
  .map(Number);

const ALLOWED_USER_IDS = (process.env.ALLOWED_USER_IDS || "")
  .split(",")
  .map((s) => s.trim())
  .filter(Boolean)
  .map(Number);

if (!BOT_TOKEN) throw new Error("BOT_TOKEN fehlt");
if (!OWNER_USER_ID) throw new Error("OWNER_USER_ID fehlt");

ensureBaseStructure();
const bot = new Telegraf(BOT_TOKEN, { handlerTimeout: 300_000 });
bot.botInfo = { username: "KIInntro_bot" };
const db = new Database(path.join(META_DIR, "messages.db"));
console.log(`[BOOT] DATA_ROOT: ${DATA_ROOT}`);

// DB init
db.exec(`
CREATE TABLE IF NOT EXISTS messages (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  chat_id INTEGER,
  chat_title TEXT,
  user TEXT,
  text TEXT,
  ts DATETIME DEFAULT CURRENT_TIMESTAMP
)
`);

const insertMsg = db.prepare(
  "INSERT INTO messages (chat_id, chat_title, user, text) VALUES (?, ?, ?, ?)"
);

const selectRecent = db.prepare(
  `SELECT chat_id, chat_title, user, text, ts
   FROM messages
   WHERE ts >= datetime('now', ?)`
);

const selectRecentByChat = db.prepare(
  `SELECT chat_id, chat_title, user, text, ts
   FROM messages
   WHERE chat_id = ?
     AND ts >= datetime('now', ?)
   ORDER BY ts ASC`
);

function isOwner(ctx) {
  return ctx.from?.id === OWNER_USER_ID;
}

function isAllowedUser(ctx) {
  return ALLOWED_USER_IDS.includes(ctx.from?.id);
}

function summarize(rows) {
  if (!rows.length) return "Keine Nachrichten im Zeitraum gefunden.";

  const counts = {};
  for (const r of rows) {
    counts[r.chat_title] = (counts[r.chat_title] || 0) + 1;
  }

  let text = "📊 Zusammenfassung:\n";
  for (const [chat, count] of Object.entries(counts)) {
    text += `- ${chat}: ${count} Nachrichten\n`;
  }
  return text;
}

// ======= KI-Chat State =======
const projectNameByUser = new Map(); // userId -> string
const contextHoursByUser = new Map(); // userId -> number (optional, default 24)
const pendingEmailByUser = new Map(); // userId -> { to, subject, body, project }
const pendingProjectByUser = new Map(); // userId -> { fields, step }
const activeReportByChat = new Map(); // chatId -> ReportSession (see below)

// ======= Group Links =======
const GROUP_LINKS_FILE = path.join(META_DIR, "group_links.json");

function loadGroupLinks() {
  if (!fs.existsSync(GROUP_LINKS_FILE)) return {};
  try {
    return JSON.parse(fs.readFileSync(GROUP_LINKS_FILE, "utf8"));
  } catch (e) { return {}; }
}

function saveGroupLink(chatId, projectCode, username) {
  const links = loadGroupLinks();
  links[String(chatId)] = {
    projectCode,
    linkedAt: new Date().toISOString(),
    linkedBy: username,
  };
  fs.writeFileSync(GROUP_LINKS_FILE, JSON.stringify(links, null, 2), "utf8");
}

function getLinkedProject(chatId) {
  const links = loadGroupLinks();
  return links[String(chatId)]?.projectCode || null;
}

function isGroupAllowed(chatId) {
  if (ALLOWED_GROUP_IDS.includes(chatId)) return true;
  return !!getLinkedProject(chatId);
}

// ======= Separate Rules (RB / BTB) =======
const RB_REGELN_FILE = path.join(DATA_ROOT, "_meta", "rb_regeln.json");
const BTB_REGELN_FILE = path.join(DATA_ROOT, "_meta", "btb_regeln.json");
const OLD_REGELN_FILE = path.join(DATA_ROOT, "_meta", "regeln.json");

function migrateRules() {
  if (fs.existsSync(OLD_REGELN_FILE) && !fs.existsSync(RB_REGELN_FILE)) {
    fs.copyFileSync(OLD_REGELN_FILE, RB_REGELN_FILE);
    console.log("[MIGRATE] Copied regeln.json -> rb_regeln.json");
  }
  if (!fs.existsSync(RB_REGELN_FILE)) {
    fs.writeFileSync(RB_REGELN_FILE, JSON.stringify({ rules: [] }, null, 2), "utf8");
  }
  if (!fs.existsSync(BTB_REGELN_FILE)) {
    fs.writeFileSync(BTB_REGELN_FILE, JSON.stringify({ rules: [] }, null, 2), "utf8");
    console.log("[MIGRATE] Created empty btb_regeln.json");
  }
}

function loadRules(reportType) {
  const file = reportType === "BTB" ? BTB_REGELN_FILE : RB_REGELN_FILE;
  if (!fs.existsSync(file)) return [];
  try {
    const data = JSON.parse(fs.readFileSync(file, "utf8"));
    return Array.isArray(data.rules) ? data.rules : [];
  } catch (e) {
    return [];
  }
}

function saveRule(rule, reportType) {
  const file = reportType === "BTB" ? BTB_REGELN_FILE : RB_REGELN_FILE;
  const rules = loadRules(reportType);
  rules.push({
    id: `r_${Date.now()}`,
    created: new Date().toISOString(),
    ...rule,
  });
  const trimmed = rules.slice(-50);
  fs.writeFileSync(file, JSON.stringify({ rules: trimmed }, null, 2), "utf8");
}

function buildRulesBlock(reportType) {
  const rules = loadRules(reportType);
  if (!rules.length) return "";
  let block = "\nGelernte Regeln (IMMER anwenden):\n";
  for (const r of rules) {
    if (r.rule) block += `- ${r.rule}\n`;
  }
  return block;
}

// Run migration at boot
migrateRules();

function cleanupPreviewFile(session) {
  if (session.previewTimer) {
    clearTimeout(session.previewTimer);
    session.previewTimer = null;
  }
  if (session.previewFile) {
    try {
      if (fs.existsSync(session.previewFile)) fs.unlinkSync(session.previewFile);
    } catch (e) { /* ignore */ }
    session.previewFile = null;
  }
}

function reportDataChanged(oldData, newData) {
  if (!oldData || !newData) return false;
  return JSON.stringify(oldData) !== JSON.stringify(newData);
}

function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true });
}

// Cleanup idle sessions every 60 seconds
setInterval(() => {
  const now = Date.now();
  for (const [chatId, session] of activeReportByChat.entries()) {
    if (now - session.lastActivity > 30 * 60 * 1000) {
      cleanupPreviewFile(session);
      activeReportByChat.delete(chatId);
      console.log(`[CLEANUP] Idle report session removed for chat ${chatId}`);
    }
  }
}, 60000);

function getPhotoExtension(filePath) {
  const ext = path.extname(filePath || "").toLowerCase();
  if (ext) return ext;
  return ".jpg";
}

async function downloadTelegramFile(fileUrl, destPath) {
  const res = await fetch(fileUrl);
  if (!res.ok) {
    throw new Error(`Download failed: ${res.status} ${res.statusText}`);
  }
  const arrayBuffer = await res.arrayBuffer();
  fs.writeFileSync(destPath, Buffer.from(arrayBuffer));
}

async function saveReportPhotos({ ctx, state, photosDir }) {
  if (!state?.photos?.length) return [];
  ensureDir(photosDir);
  const saved = [];

  for (let i = 0; i < state.photos.length; i += 1) {
    const fileId = state.photos[i];
    const file = await ctx.telegram.getFile(fileId);
    const filePath = file?.file_path || "";
    const ext = getPhotoExtension(filePath);
    const fileUrl = await ctx.telegram.getFileLink(fileId);
    const dest = path.join(photosDir, `Foto ${i + 1}${ext}`);
    await downloadTelegramFile(fileUrl.toString(), dest);
    saved.push(dest);
  }
  return saved;
}

function getProjectName(userId) {
  return projectNameByUser.get(userId) || "Allgemein";
}
function getContextHours(userId) {
  return contextHoursByUser.get(userId) || 24;
}

function buildContextTextForHours(hours) {
  // Kontext aus ALLEN erlaubten/verknuepften Gruppen
  const rows = selectRecent.all(`-${hours} hours`).filter((r) =>
    isGroupAllowed(r.chat_id)
  );

  if (!rows.length) return "";

  // nach Gruppe sortieren, aber kompakt
  const byChat = new Map();
  for (const r of rows) {
    if (!byChat.has(r.chat_title)) byChat.set(r.chat_title, []);
    byChat.get(r.chat_title).push(`${r.user}: ${r.text}`);
  }

  let out = "";
  for (const [title, msgs] of byChat.entries()) {
    out += `\n[${title}]\n`;
    // nicht zu lang machen, sonst Prompt explodiert:
    const last = msgs.slice(-80); // letzte 80 Messages pro Gruppe
    out += last.map((m) => `- ${m}`).join("\n") + "\n";
  }
  return out.trim();
}

function ensureContactsFile() {
  if (fs.existsSync(CONTACTS_FILE)) return;
  const header = "project;email;phone;name;source_user;source_text;ts\n";
  fs.writeFileSync(CONTACTS_FILE, header, "utf8");
}

function ensureEmailLogFile() {
  if (fs.existsSync(EMAIL_LOG_FILE)) return;
  const header = "ts;action;project;to;subject;status;error;user;context\n";
  fs.writeFileSync(EMAIL_LOG_FILE, header, "utf8");
}

function logEmailEvent({ action, project, to, subject, status, error, user, context }) {
  ensureEmailLogFile();
  const ts = new Date().toISOString();
  const row = [
    ts,
    action || "",
    project || "",
    to || "",
    subject || "",
    status || "",
    error || "",
    user || "",
    context || "",
  ]
    .map(sanitizeCsvValue)
    .join(";");
  fs.appendFileSync(EMAIL_LOG_FILE, row + "\n", "utf8");
}

function extractEmails(text) {
  const matches = text.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi);
  return matches ? Array.from(new Set(matches)) : [];
}

function extractPhones(text) {
  const matches = text.match(/(?:\+?\d[\d\s().-]{7,}\d)/g);
  return matches ? Array.from(new Set(matches)) : [];
}

function sanitizeCsvValue(value) {
  const v = String(value || "").replace(/\r?\n/g, " ").trim();
  return v.replace(/;/g, ",");
}

function appendContactRows(project, emails, phones, user, text) {
  if (!emails.length && !phones.length) return;
  ensureContactsFile();
  const ts = new Date().toISOString();
  const rows = [];
  const emailList = emails.length ? emails : [""];
  const phoneList = phones.length ? phones : [""];
  for (const email of emailList) {
    for (const phone of phoneList) {
      rows.push(
        [
          project,
          email,
          phone,
          "",
          user,
          text,
          ts,
        ].map(sanitizeCsvValue).join(";")
      );
    }
  }
  fs.appendFileSync(CONTACTS_FILE, rows.join("\n") + "\n", "utf8");
}

function loadContacts() {
  if (!fs.existsSync(CONTACTS_FILE)) return [];
  const raw = fs.readFileSync(CONTACTS_FILE, "utf8");
  const lines = raw.split(/\r?\n/).filter(Boolean);
  if (!lines.length) return [];
  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    const parts = lines[i].split(";");
    const [project, email, phone, name] = parts;
    if (!project) continue;
    rows.push({
      project: project.trim(),
      email: (email || "").trim(),
      phone: (phone || "").trim(),
      name: (name || "").trim(),
    });
  }
  return rows;
}

function getRecipientsForProject(projectName) {
  const contacts = loadContacts();
  const target = (projectName || "").trim().toLowerCase();
  if (!target) return [];
  const matches = contacts.filter(
    (c) => c.project.toLowerCase() === target || target.includes(c.project.toLowerCase())
  );
  const emails = matches.map((m) => m.email).filter(Boolean);
  return Array.from(new Set(emails));
}

function canSendEmail() {
  return SMTP_HOST && SMTP_USER && SMTP_PASS && SMTP_FROM;
}

function getTransporter() {
  return nodemailer.createTransport({
    host: SMTP_HOST,
    port: SMTP_PORT,
    secure: SMTP_PORT === 465,
    auth: { user: SMTP_USER, pass: SMTP_PASS },
  });
}

const { execSync } = require("child_process");

function ensureNasMounted() {
  try {
    fs.accessSync(NAS_MOUNT, fs.constants.R_OK);
    const contents = fs.readdirSync(NAS_MOUNT);
    if (contents.length > 0) return true;
  } catch (e) { /* not mounted */ }
  try {
    fs.mkdirSync(NAS_MOUNT, { recursive: true });
    const pass = NAS_PASS ? `'${NAS_PASS}'` : "";
    execSync(`mount_smbfs //admin:${pass}@192.168.2.126/server ${NAS_MOUNT}`, { timeout: 10000 });
    return true;
  } catch (e) {
    console.error("[NAS-MOUNT] error:", e.message);
    return false;
  }
}

async function copyReportToNAS({ outFile, pdfFile, photosDir, projectName, reportNumber }) {
  if (!ensureNasMounted()) {
    console.error("[NAS-COPY] NAS nicht erreichbar, Kopie uebersprungen.");
    return false;
  }
  const destDir = path.join(NAS_MOUNT, NAS_REPORTS_PATH, projectName, "Regieberichte", `RB${reportNumber}`);
  fs.mkdirSync(destDir, { recursive: true });
  // Copy report file
  const destFile = path.join(destDir, path.basename(outFile));
  fs.copyFileSync(outFile, destFile);
  console.log(`[NAS-COPY] Report kopiert: ${destFile}`);
  // Copy PDF file
  if (pdfFile && fs.existsSync(pdfFile)) {
    const destPdf = path.join(destDir, path.basename(pdfFile));
    fs.copyFileSync(pdfFile, destPdf);
    console.log(`[NAS-COPY] PDF kopiert: ${destPdf}`);
  }
  // Copy photos if they exist
  if (photosDir && fs.existsSync(photosDir)) {
    const photosDestDir = path.join(destDir, "Fotos");
    fs.mkdirSync(photosDestDir, { recursive: true });
    const photos = fs.readdirSync(photosDir);
    for (const photo of photos) {
      fs.copyFileSync(path.join(photosDir, photo), path.join(photosDestDir, photo));
    }
    console.log(`[NAS-COPY] ${photos.length} Fotos kopiert.`);
  }
  return true;
}

async function saveReportEmailDraft({ to, cc, contactName, attachmentPath, reportNumber, projectName }) {
  const subject = REPORT_EMAIL_SUBJECT
    .replace(/\{reportNumber\}/g, reportNumber)
    .replace(/\{projectName\}/g, projectName);

  const credential = new ClientSecretCredential(AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET);
  const client = Client.initWithMiddleware({
    authProvider: {
      getAccessToken: async () => {
        const token = await credential.getToken("https://graph.microsoft.com/.default");
        return token.token;
      },
    },
  });

  const fileBuffer = fs.readFileSync(attachmentPath);
  const contentBytes = fileBuffer.toString("base64");

  const toRecipients = to.split(",").map((e) => e.trim()).filter(Boolean).map((email) => ({
    emailAddress: { address: email },
  }));
  const ccRecipients = (cc || "").split(",").map((e) => e.trim()).filter(Boolean).map((email) => ({
    emailAddress: { address: email },
  }));

  const bodyContent = "";
  const contentType = "Text";

  const message = {
    subject: subject,
    body: { contentType, content: bodyContent },
    toRecipients: toRecipients,
    ccRecipients: ccRecipients.length > 0 ? ccRecipients : undefined,
    attachments: [
      {
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: path.basename(attachmentPath),
        contentBytes: contentBytes,
      },
    ],
  };

  await client.api("/users/office@inntro.de/messages").post(message);
}

function extractEmail(text) {
  const match = text.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return match ? match[0] : "";
}

function hasEmailAddress(text) {
  return extractEmails(text).length > 0;
}

function isSendIntent(text) {
  const lower = text.toLowerCase();
  return (
    lower.includes("schick") ||
    lower.includes("sende") ||
    lower.includes("send") ||
    lower.includes("verschick") ||
    lower.includes("versend") ||
    lower.includes("abschick")
  );
}

function isMailRequest(text) {
  const lower = text.toLowerCase();
  const hasMailWord = /(?:e-?mail|e mail|email|mail)/i.test(text);
  const mailPattern = /(?:mail|email)\s+an\s+/i.test(text) || /per\s+(mail|email)/i.test(text);
  return (isSendIntent(text) && (hasMailWord || hasEmailAddress(text))) || mailPattern;
}

function detectProjectFromText(text) {
  const contacts = loadContacts();
  const names = Array.from(new Set(contacts.map((c) => c.project).filter(Boolean)));
  const lower = text.toLowerCase();
  return names.find((name) => lower.includes(name.toLowerCase())) || "";
}

function parseReportStart(text, chatId) {
  const trimmed = String(text || "").trim();
  // "RB SB-NORD-01" or "BTB SB-NORD-01"
  let match = trimmed.match(/^(RB|BTB)\s+(.+)$/i);
  if (match) return { type: match[1].toUpperCase(), projectCode: match[2].trim() };
  // Bare "RB" or "BTB" - use linked project if available
  match = trimmed.match(/^(RB|BTB)\s*$/i);
  if (match && chatId) {
    const linkedCode = getLinkedProject(chatId);
    if (linkedCode) return { type: match[1].toUpperCase(), projectCode: linkedCode };
  }
  return null;
}

// ========== Middleware: /verknuepfen Fallback ==========
bot.use(async (ctx, next) => {
  const text = ctx.message?.text || "";
  const match = text.match(/^\/verknuepfen(?:@\w+)?\s+(.*)/i);
  if (match && ctx.chat && ["group", "supergroup"].includes(ctx.chat.type)) {
    const code = match[1].trim();
    console.log(`[VERKNUEPFEN-MW] User ${ctx.from?.id}, Chat ${ctx.chat.id}, Code: ${code}`);
    if (!code) {
      return ctx.reply("Bitte Projekt-Kuerzel angeben: /verknuepfen SB-NORD-01");
    }
    const project = getProjectByCode(code);
    if (!project) {
      return ctx.reply(`Projekt '${code}' nicht gefunden. Erstelle es zuerst mit /newproject.`);
    }
    const user = ctx.from?.username || ctx.from?.first_name || "unknown";
    saveGroupLink(ctx.chat.id, project.code, user);
    return ctx.reply(
      `Gruppe verknuepft mit Projekt: ${project.code} - ${project.name}\n` +
        `Du kannst jetzt einfach 'RB' oder 'BTB' schreiben um einen Bericht zu starten.`
    );
  }
  return next();
});

// ========== COMMANDS ==========
bot.start((ctx) => {
  ctx.reply(
    "Bau-KI Bot ist aktiv.\n\n" +
      "Befehle:\n" +
      "/who - zeigt deine User-ID\n" +
      "/chatid - zeigt Chat-/Gruppen-ID\n" +
      "/verknuepfen <Projekt> - Gruppe mit Projekt verknuepfen\n" +
      "/summary - Zusammenfassung (24h)\n" +
      "/summarygroup - Zusammenfassung nur diese Gruppe\n" +
      "/report - KI Baustellenbericht (24h)\n\n" +
      "Projekte:\n" +
      "/newproject <Kuerzel> | <Name> | <Person> | <Email> | <CC> | <Auftraggeber>\n" +
      "Oder schreib mir privat z.B. 'Neues Projekt anlegen'\n\n" +
      "Berichte (in Gruppen):\n" +
      "RB <Projekt> oder RB - Regiebericht starten\n" +
      "BTB <Projekt> oder BTB - Bautagesbericht starten\n" +
      "Ende - Vorschau erstellen\n" +
      "freigeben - Bericht speichern\n\n" +
      "Privater Chat:\n" +
      "Schreib mir einfach - ich bin immer aktiv!\n" +
      "/project <name> - Projekt setzen\n" +
      "/ask <frage> - einmalige Frage"
  );
});

bot.command("who", (ctx) => {
  ctx.reply(`Deine Telegram User-ID:\n${ctx.from.id}`);
});

// ✅ Chat/Gruppen-ID anzeigen
bot.command("chatid", (ctx) => {
  if (!isAllowedUser(ctx)) return;
  const chatId = ctx.chat?.id;
  const title = ctx.chat?.title || ctx.chat?.username || "private";
  ctx.reply(`📌 Chat: ${title}\n🆔 Chat-ID: ${chatId}`);
});

// ✅ Prüfen ob aktuelle Gruppe erlaubt ist
bot.command("allowthis", (ctx) => {
  if (!isAllowedUser(ctx)) return;
  const chatId = ctx.chat?.id;
  const title = ctx.chat?.title || ctx.chat?.username || "private";
  const allowed = isGroupAllowed(chatId);
  ctx.reply(
    allowed
      ? `Erlaubt: ${title}\nID: ${chatId}`
      : `Nicht erlaubt: ${title}\nID: ${chatId}\n\nNutze /verknuepfen <Projekt> um diese Gruppe zu verknuepfen.`
  );
});

// ✅ Zusammenfassung über ALLE Gruppen (Owner only)
bot.command("summary", (ctx) => {
  if (!isAllowedUser(ctx)) return ctx.reply("⛔ Keine Berechtigung");

  const arg = (ctx.message.text.split(" ")[1] || "").trim();
  const hours = arg && /^\d+$/.test(arg) ? Number(arg) : 24;

  const rows = selectRecent.all(`-${hours} hours`).filter((r) =>
    isGroupAllowed(r.chat_id)
  );

  if (!rows.length) {
    return ctx.reply(
      "Keine Nachrichten gefunden.\n" +
        "Check: Bot ist in der Gruppe? /setprivacy in BotFather = Disable?"
    );
  }

  ctx.reply(summarize(rows));
});

// ✅ Zusammenfassung NUR für die aktuelle Gruppe
bot.command("summarygroup", (ctx) => {
  if (!isAllowedUser(ctx)) return ctx.reply("⛔ Keine Berechtigung");

  const chat = ctx.chat;
  if (!chat || !["group", "supergroup"].includes(chat.type)) {
    return ctx.reply("Dieser Befehl funktioniert nur in einer Gruppe.");
  }

  if (!isGroupAllowed(chat.id)) {
    return ctx.reply(
      `Diese Gruppe ist nicht verknuepft.\nNutze /verknuepfen <Projekt> um sie zu verknuepfen.`
    );
  }

  const arg = (ctx.message.text.split(" ")[1] || "").trim();
  const hours = arg && /^\d+$/.test(arg) ? Number(arg) : 24;

  const rows = selectRecentByChat.all(chat.id, `-${hours} hours`);
  ctx.reply(summarize(rows));
});

// ✅ KI Report (nur diese Gruppe, letzte 24h)
bot.command("report", async (ctx) => {
  if (!isAllowedUser(ctx)) return ctx.reply("⛔ Keine Berechtigung");

  const chat = ctx.chat;
  if (!chat || !["group", "supergroup"].includes(chat.type)) {
    return ctx.reply("Bitte diesen Befehl in einer Gruppe ausführen.");
  }

  if (!isGroupAllowed(chat.id)) {
    return ctx.reply("Diese Gruppe ist nicht verknuepft. Nutze /verknuepfen <Projekt>.");
  }

  const rows = selectRecentByChat.all(chat.id, "-24 hours");
  if (!rows.length) return ctx.reply("Keine Nachrichten für den Bericht gefunden.");

  await ctx.reply("🧠 KI erstellt Bericht… ⏳");

  try {
    console.log("[REPORT] calling KI for chat", chat.id, chat.title, "rows:", rows.length);
    const report = await generateReport(rows, chat.title || String(chat.id));
    console.log("[REPORT] ok, len:", report.length);

    const chunks = report.match(/[\s\S]{1,3500}/g) || [];
    for (const part of chunks) await ctx.reply(part);
  } catch (err) {
    console.error("[REPORT] error:", err);
    await ctx.reply("❌ Fehler bei der KI-Berichtserstellung. Schau ins Terminal-Log.");
  }
});

// ======= Projekte =======
bot.command("newproject", async (ctx) => {
  if (!isAllowedUser(ctx)) return ctx.reply("⛔ Keine Berechtigung");
  const raw = ctx.message.text.replace(/^\/newproject\s*/i, "").trim();
  if (!raw) {
    return ctx.reply(
      "Bitte angeben:\n" +
        "/newproject <Kuerzel> | <Projektname> | <Ansprechperson> | <Email> | <CC> | <Auftraggeber>"
    );
  }

  const parts = raw.split("|").map((p) => p.trim());
  const [code, name, contactName, contactEmail, cc, client] = parts;
  if (!code || !contactName || !contactEmail) {
    return ctx.reply(
      "Fehlende Pflichtfelder. Format:\n" +
        "/newproject <Kuerzel> | <Projektname> | <Ansprechperson> | <Email> | <CC> | <Auftraggeber>"
    );
  }

  const project = {
    code,
    name: name || code,
    contactName,
    contactEmail,
    cc: cc || "",
    client: client || "",
    createdAt: new Date().toISOString(),
  };

  const savedProject = await upsertProject(project);
  const projectPath = path.join(DATA_ROOT, savedProject.dirName);
  return ctx.reply(
    `✅ Projekt gespeichert: ${savedProject.code}\n` +
      `Ordnerstruktur erstellt in ${projectPath}`
  );
});

// ======= Privater Projekt-KI Chat =======

// Projekt setzen
bot.command("project", (ctx) => {
  if (!isAllowedUser(ctx)) return ctx.reply("⛔ Keine Berechtigung");
  const name = ctx.message.text.replace(/^\/project\s*/i, "").trim();
  if (!name) return ctx.reply("Bitte Projektname angeben: /project Baustelle Nord");
  projectNameByUser.set(ctx.from.id, name);
  ctx.reply(`✅ Projekt gesetzt: ${name}`);
});

// Kontext-Zeitraum (in Stunden)
bot.command("context", (ctx) => {
  if (!isAllowedUser(ctx)) return ctx.reply("⛔ Keine Berechtigung");
  const arg = (ctx.message.text.split(" ")[1] || "").trim();
  const hours = arg && /^\d+$/.test(arg) ? Number(arg) : 24;
  contextHoursByUser.set(ctx.from.id, hours);
  ctx.reply(`✅ Kontext-Zeitraum gesetzt: letzte ${hours} Stunden`);
});

// ======= Gruppenverknuepfung =======
bot.command("verknuepfen", async (ctx) => {
  console.log(`[VERKNUEPFEN] User ${ctx.from?.id}, Chat ${ctx.chat?.id}, Text: ${ctx.message?.text}`);
  const chat = ctx.chat;
  if (!chat || !["group", "supergroup"].includes(chat.type)) {
    return ctx.reply("Dieser Befehl funktioniert nur in einer Gruppe.");
  }
  const code = ctx.message.text.replace(/^\/verknuepfen\s*/i, "").trim();
  if (!code) {
    return ctx.reply("Bitte Projekt-Kuerzel angeben: /verknuepfen SB-NORD-01");
  }
  const project = getProjectByCode(code);
  if (!project) {
    return ctx.reply(`Projekt '${code}' nicht gefunden. Erstelle es zuerst mit /newproject.`);
  }
  const user = ctx.from?.username || ctx.from?.first_name || "unknown";
  saveGroupLink(chat.id, project.code, user);
  ctx.reply(
    `Gruppe verknuepft mit Projekt: ${project.code} - ${project.name}\n` +
      `Du kannst jetzt einfach 'RB' oder 'BTB' schreiben um einen Bericht zu starten.`
  );
});

// Einmalige Frage (privat empfohlen)
bot.command("ask", async (ctx) => {
  if (!isAllowedUser(ctx)) return ctx.reply("⛔ Keine Berechtigung");

  const question = ctx.message.text.replace(/^\/ask\s*/i, "").trim();
  if (!question) return ctx.reply("Bitte Frage angeben: /ask Was sind die Risiken heute?");

  await ctx.reply("🧠 Denke nach… ⏳");

  try {
    const pn = getProjectName(ctx.from.id);
    const hours = getContextHours(ctx.from.id);
    const contextText = buildContextTextForHours(hours);

    const answer = await chatProjectAI({
      userText: question,
      projectName: pn,
      contextText,
    });

    const chunks = answer.match(/[\s\S]{1,3500}/g) || [];
    for (const part of chunks) await ctx.reply(part);
  } catch (err) {
    console.error("[ASK] error:", err);
    await ctx.reply("❌ KI-Fehler. Schau ins Terminal.");
  }
});

// ========== ONE MESSAGE HANDLER ==========
bot.on("photo", async (ctx) => {
  const chat = ctx.chat;
  if (!chat || !["group", "supergroup"].includes(chat.type)) return;
  if (!isGroupAllowed(chat.id)) return;
  if (!activeReportByChat.has(chat.id)) return;

  const photos = ctx.message?.photo || [];
  if (!photos.length) return;
  const largest = photos[photos.length - 1];
  const state = activeReportByChat.get(chat.id);
  if (largest?.file_id) state.photos.push(largest.file_id);
});

bot.on("document", async (ctx) => {
  const chat = ctx.chat;
  if (!chat || !["group", "supergroup"].includes(chat.type)) return;
  if (!isGroupAllowed(chat.id)) return;
  if (!activeReportByChat.has(chat.id)) return;

  const doc = ctx.message?.document;
  if (!doc?.file_id) return;
  if (!String(doc.mime_type || "").startsWith("image/")) return;
  const state = activeReportByChat.get(chat.id);
  state.photos.push(doc.file_id);
});

bot.on("text", async (ctx) => {
  console.log(`[DEBUG] Nachricht von User ${ctx.from?.id} (${ctx.from?.first_name}) in Chat ${ctx.chat?.id} (${ctx.chat?.type}): ${ctx.message?.text?.substring(0, 50)}`);
  const chat = ctx.chat;

  // 1) Gruppen -> speichern
  if (chat && ["group", "supergroup"].includes(chat.type)) {
    if (!isGroupAllowed(chat.id)) return;

    const user =
      ctx.from?.username ||
      [ctx.from?.first_name, ctx.from?.last_name].filter(Boolean).join(" ") ||
      "unknown";

    const text = ctx.message?.text || "";
    const project = chat.title || String(chat.id);

    const startInfo = parseReportStart(text, chat.id);
    if (startInfo) {
      const existing = activeReportByChat.get(chat.id);
      if (existing) {
        await ctx.reply(
          `Es laeuft bereits ein ${existing.type}-Bericht fuer ${existing.projectCode}. Schreibe 'abbrechen' zum Beenden oder chatte weiter.`
        );
      } else {
        const projectEntry = getProjectByCode(startInfo.projectCode);
        if (!projectEntry) {
          await ctx.reply(
            `Projekt '${startInfo.projectCode}' nicht gefunden. Lege es mit /newproject an.`
          );
        } else {
          const rulesBlock = buildRulesBlock(startInfo.type);
          const systemPrompt = buildReportSystemPrompt({
            type: startInfo.type,
            projectCode: startInfo.projectCode,
            projectName: projectEntry.name,
            client: projectEntry.client,
            rulesBlock,
          });
          activeReportByChat.set(chat.id, {
            type: startInfo.type,
            projectCode: startInfo.projectCode,
            project: projectEntry,
            conversationHistory: [],
            reportData: { leistungen: [], arbeitskraefte: [], material: [] },
            photos: [],
            previewFile: null,
            previewTimer: null,
            previewNumber: null,
            systemPrompt,
            startedAt: Date.now(),
            lastActivity: Date.now(),
          });
          await ctx.reply(
            `${startInfo.type}-Bericht gestartet fuer ${startInfo.projectCode}.\n` +
              "Erzaehl mir was heute auf der Baustelle gemacht wurde.\n" +
              "Du kannst Fotos senden, Korrekturen machen, und mit 'Ende' eine Vorschau erstellen.\n" +
              "'freigeben' zum endgueltigen Speichern | 'abbrechen' zum Verwerfen."
          );
        }
      }
    } else if (activeReportByChat.has(chat.id)) {
      const session = activeReportByChat.get(chat.id);
      session.lastActivity = Date.now();
      const lower = text.trim().toLowerCase();

      // === ABBRECHEN ===
      if (/^(abbrechen|cancel|stop)$/i.test(lower)) {
        cleanupPreviewFile(session);
        activeReportByChat.delete(chat.id);
        await ctx.reply("Bericht abgebrochen.");
        insertMsg.run(chat.id, chat.title || String(chat.id), user, text);
        return;
      }

      // === FREIGEBEN ===
      if (/^(freigeben|freigabe|speichern)$/i.test(lower)) {
        if (!session.reportData || (
          !session.reportData.leistungen.length &&
          !session.reportData.arbeitskraefte.length &&
          !session.reportData.material.length
        )) {
          await ctx.reply("Noch keine Daten vorhanden. Erzaehl mir zuerst was auf der Baustelle gemacht wurde.");
          insertMsg.run(chat.id, chat.title || String(chat.id), user, text);
          return;
        }
        try {
          await ctx.reply("Bericht wird gespeichert...");
          const sections = buildReportSectionsFromAIData(session.reportData);
          const lines = [
            ...sections.leistungen,
            ...sections.arbeitskraefte,
            ...sections.material,
          ];
          const result = await generateReportFile({
            type: session.type,
            project: session.project,
            lines,
            sections,
            preview: false,
          });
          let savedPhotos = [];
          try {
            savedPhotos = await saveReportPhotos({ ctx, state: session, photosDir: result.photosDir });
          } catch (err) {
            console.error("[REPORT-PHOTOS] error:", err);
          }
          // PDF generieren
          let pdfFile = null;
          if (session.type === "RB") {
            try {
              const finalSections = normalizeReportSections(lines, sections);
              const workerEntries = finalSections.arbeitskraefte.flatMap(parseWorkerEntries);
              const workers = aggregateWorkers(workerEntries);
              const materials = finalSections.material.map(parseMaterialLine);
              pdfFile = result.outFile.replace(/\.xlsx$/i, ".pdf");
              await generateReportPDF({
                project: session.project,
                reportNumber: result.number,
                sections: {
                  leistungen: finalSections.leistungen,
                  arbeitskraefte: workers,
                  material: materials,
                },
                photos: savedPhotos || [],
                outFile: pdfFile,
              });
              console.log(`[REPORT-PDF] PDF erstellt: ${pdfFile}`);
            } catch (pdfErr) {
              console.error("[REPORT-PDF] error:", pdfErr);
              pdfFile = null;
            }
          }
          // Telegram: PDF + Excel senden
          try {
            if (pdfFile && fs.existsSync(pdfFile)) {
              await ctx.replyWithDocument({ source: pdfFile });
            }
            await ctx.replyWithDocument({ source: result.outFile });
          } catch (err) {
            console.error("[REPORT-UPLOAD] error:", err);
          }
          await ctx.reply(`Bericht gespeichert: ${result.reportFolder}`);
          // Auf NAS kopieren
          try {
            const nasCopied = await copyReportToNAS({
              outFile: result.outFile,
              pdfFile: pdfFile,
              photosDir: result.photosDir,
              projectName: session.project.name,
              reportNumber: result.number,
            });
            if (nasCopied) {
              await ctx.reply("Bericht auf NAS-Server kopiert.");
            }
          } catch (nasErr) {
            console.error("[NAS-COPY] error:", nasErr);
            await ctx.reply("Hinweis: Kopie auf NAS-Server fehlgeschlagen.");
          }
          // E-Mail-Entwurf speichern (PDF als Anhang, Fallback auf Excel)
          if (session.project.contactEmail && AZURE_CLIENT_ID && AZURE_TENANT_ID && AZURE_CLIENT_SECRET) {
            try {
              const emailAttachment = (pdfFile && fs.existsSync(pdfFile)) ? pdfFile : result.outFile;
              await saveReportEmailDraft({
                to: session.project.contactEmail,
                cc: session.project.cc || "",
                contactName: session.project.contactName || "",
                attachmentPath: emailAttachment,
                reportNumber: result.number,
                projectName: session.project.name,
              });
              await ctx.reply("E-Mail-Entwurf gespeichert.");
            } catch (emailErr) {
              console.error("[EMAIL-DRAFT] error:", emailErr);
              await ctx.reply("Hinweis: E-Mail-Entwurf konnte nicht gespeichert werden.");
            }
          }
        } catch (err) {
          console.error("[REPORT-SAVE] error:", err);
          await ctx.reply("Fehler beim Speichern. Schau ins Terminal.");
        } finally {
          cleanupPreviewFile(session);
          activeReportByChat.delete(chat.id);
        }
        insertMsg.run(chat.id, chat.title || String(chat.id), user, text);
        return;
      }

      // === ENDE (Vorschau) ===
      if (/^ende$/i.test(lower)) {
        if (!session.reportData || (
          !session.reportData.leistungen.length &&
          !session.reportData.arbeitskraefte.length &&
          !session.reportData.material.length
        )) {
          await ctx.reply("Noch keine Daten vorhanden. Erzaehl mir zuerst was auf der Baustelle gemacht wurde.");
          insertMsg.run(chat.id, chat.title || String(chat.id), user, text);
          return;
        }
        try {
          await ctx.reply("Vorschau wird erstellt...");
          const sections = buildReportSectionsFromAIData(session.reportData);
          const lines = [
            ...sections.leistungen,
            ...sections.arbeitskraefte,
            ...sections.material,
          ];
          // Clean up old preview if exists
          cleanupPreviewFile(session);
          const result = await generateReportFile({
            type: session.type,
            project: session.project,
            lines,
            sections,
            preview: true,
          });
          session.previewFile = result.outFile;
          session.previewNumber = result.number;
          // 10 min cleanup timer
          session.previewTimer = setTimeout(() => {
            if (session.previewFile && fs.existsSync(session.previewFile)) {
              try { fs.unlinkSync(session.previewFile); } catch (e) { /* ignore */ }
              session.previewFile = null;
            }
          }, 10 * 60 * 1000);
          try {
            await ctx.replyWithDocument({ source: result.outFile });
          } catch (err) {
            console.error("[REPORT-PREVIEW-UPLOAD] error:", err);
          }
          await ctx.reply(
            "Vorschau gesendet. Pruefe den Bericht.\n" +
              "- Beschreibe Korrekturen wenn etwas nicht stimmt\n" +
              "- 'freigeben' zum endgueltigen Speichern\n" +
              "- 'Ende' fuer eine neue Vorschau nach Korrekturen\n" +
              "- 'abbrechen' zum Verwerfen"
          );
        } catch (err) {
          console.error("[REPORT-PREVIEW] error:", err);
          await ctx.reply("Fehler bei der Vorschau. Schau ins Terminal.");
        }
        insertMsg.run(chat.id, chat.title || String(chat.id), user, text);
        return;
      }

      // === NORMAL CHAT MESSAGE -> AI ===
      try {
        await ctx.sendChatAction("typing");
        const previousData = session.reportData
          ? JSON.parse(JSON.stringify(session.reportData))
          : null;

        const { response, reportData } = await chatReportAI({
          systemPrompt: session.systemPrompt,
          conversationHistory: session.conversationHistory,
          userMessage: text,
        });

        // Update conversation history (keep last 20 messages max)
        session.conversationHistory.push({ role: "user", content: text });
        session.conversationHistory.push({ role: "assistant", content: response });
        if (session.conversationHistory.length > 40) {
          session.conversationHistory = session.conversationHistory.slice(-40);
        }

        // Update report data if AI returned valid data
        if (reportData) {
          session.reportData = reportData;
        }

        // Send AI response
        const chunks = response.match(/[\s\S]{1,3500}/g) || [];
        for (const part of chunks) await ctx.reply(part);

        // Check if correction happened -> extract rule in background
        if (reportData && previousData && reportDataChanged(previousData, reportData) &&
            session.conversationHistory.length > 4) {
          // Don't await - run in background
          extractRuleFromCorrection({
            userMessage: text,
            previousData,
            correctedData: reportData,
          }).then((ruleResult) => {
            if (ruleResult && ruleResult.rule) {
              saveRule({
                trigger: ruleResult.trigger,
                rule: ruleResult.rule,
                source: { userMessage: text, projectCode: session.projectCode },
              }, session.type);
              console.log(`[RULE] New ${session.type} rule saved:`, ruleResult.rule);
            }
          }).catch((err) => {
            console.error("[RULE-EXTRACT] error:", err);
          });
        }
      } catch (err) {
        console.error("[REPORT-CHAT] error:", err);
        await ctx.reply("KI-Fehler. Versuche es nochmal oder schau ins Terminal.");
      }
    }

    const emails = extractEmails(text);
    const phones = extractPhones(text);
    appendContactRows(project, emails, phones, user, text);
    insertMsg.run(chat.id, chat.title || String(chat.id), user, text);
    return;
  }

  // 2) Private -> KI Chat (fuer erlaubte User)
  if (chat && chat.type === "private") {
    if (!isAllowedUser(ctx)) return;

    const userText = ctx.message.text;
    const lower = userText.trim().toLowerCase();

    // === Projekt-Erstellung per Chat ===
    const pendingProject = pendingProjectByUser.get(ctx.from.id);
    if (pendingProject) {
      if (lower === "abbrechen" || lower === "cancel") {
        pendingProjectByUser.delete(ctx.from.id);
        return ctx.reply("Projekterstellung abgebrochen.");
      }
      // Fill in missing fields step by step
      const fields = pendingProject.fields;
      if (!fields.code) {
        fields.code = userText.trim();
        if (!fields.name) {
          return ctx.reply("Projektname?");
        }
      } else if (!fields.name) {
        fields.name = userText.trim();
        return ctx.reply("Ansprechperson (Name)?");
      } else if (!fields.contactName) {
        fields.contactName = userText.trim();
        return ctx.reply("E-Mail der Ansprechperson?");
      } else if (!fields.contactEmail) {
        fields.contactEmail = userText.trim();
        return ctx.reply("CC E-Mail? (oder 'nein' zum Ueberspringen)");
      } else if (fields.cc === undefined) {
        fields.cc = lower === "nein" || lower === "no" || lower === "-" ? "" : userText.trim();
        return ctx.reply("Auftraggeber?");
      } else if (!fields.client) {
        fields.client = userText.trim();
        // All fields collected - create project
        try {
          const project = {
            code: fields.code,
            name: fields.name || fields.code,
            contactName: fields.contactName,
            contactEmail: fields.contactEmail,
            cc: fields.cc || "",
            client: fields.client || "",
            createdAt: new Date().toISOString(),
          };
          const saved = await upsertProject(project);
          pendingProjectByUser.delete(ctx.from.id);
          return ctx.reply(
            `Projekt erstellt: ${saved.code} - ${saved.name}\n` +
              `Ordner: ${path.join(DATA_ROOT, saved.dirName)}\n\n` +
              "Erstelle jetzt eine Telegram-Gruppe, fuege mich hinzu, und verknuepfe sie mit:\n" +
              `/verknuepfen ${saved.code}`
          );
        } catch (err) {
          console.error("[PROJECT-CREATE] error:", err);
          pendingProjectByUser.delete(ctx.from.id);
          return ctx.reply("Fehler beim Erstellen. Schau ins Terminal.");
        }
      }
      return;
    }

    const pending = pendingEmailByUser.get(ctx.from.id);
    const sendIntent = isSendIntent(userText);

    if (
      pending &&
      (lower === "go" ||
        lower === "senden" ||
        lower === "send" ||
        lower === "verschicke" ||
        lower === "schicke" ||
        lower === "versenden" ||
        lower === "abschicken" ||
        lower === "los")
    ) {
      if (!canSendEmail()) {
        logEmailEvent({
          action: "send",
          project: pending.project,
          to: pending.to,
          subject: pending.subject,
          status: "error",
          error: "SMTP_NOT_CONFIGURED",
          user: String(ctx.from?.id || ""),
          context: "missing_smtp",
        });
        return ctx.reply(
          "❌ SMTP nicht konfiguriert. Bitte SMTP_HOST/SMTP_USER/SMTP_PASS/SMTP_FROM setzen."
        );
      }
      try {
        const transporter = getTransporter();
        await transporter.sendMail({
          from: SMTP_FROM,
          to: pending.to,
          subject: pending.subject,
          text: pending.body,
        });
        logEmailEvent({
          action: "send",
          project: pending.project,
          to: pending.to,
          subject: pending.subject,
          status: "ok",
          user: String(ctx.from?.id || ""),
          context: "sent",
        });
        pendingEmailByUser.delete(ctx.from.id);
        return ctx.reply("✅ E-Mail gesendet.");
      } catch (err) {
        console.error("[EMAIL] send error:", err);
        logEmailEvent({
          action: "send",
          project: pending.project,
          to: pending.to,
          subject: pending.subject,
          status: "error",
          error: String(err?.message || err),
          user: String(ctx.from?.id || ""),
          context: "send_failed",
        });
        return ctx.reply("❌ Fehler beim Senden der E-Mail. Schau ins Terminal.");
      }
    }

    if (pending && (lower === "abbrechen" || lower === "cancel" || lower === "stop")) {
      pendingEmailByUser.delete(ctx.from.id);
      return ctx.reply("🛑 E-Mail-Entwurf verworfen.");
    }

    const detectedProject = detectProjectFromText(userText);
    const pn = getProjectName(ctx.from.id);
    const projectForRecipients = detectedProject || pn;
    const projectRecipients = getRecipientsForProject(projectForRecipients);
    const shouldDraft = isMailRequest(userText) || (sendIntent && projectRecipients.length > 0);

    if (!pending && sendIntent && !shouldDraft) {
      logEmailEvent({
        action: "send",
        project: projectForRecipients,
        to: "",
        subject: "",
        status: "error",
        error: "NO_DRAFT",
        user: String(ctx.from?.id || ""),
        context: "send_without_draft",
      });
      return ctx.reply(
        "Ich habe keinen Entwurf gefunden. Bitte schreibe zuerst eine Mail-Anfrage (z.B. 'schicke eine Mail an name@firma.at ...')."
      );
    }

    if (shouldDraft) {
      const recipient =
        extractEmail(userText) ||
        (projectRecipients.length ? projectRecipients.join(",") : "") ||
        DEFAULT_EMAIL_TO;
      if (!recipient) {
        logEmailEvent({
          action: "draft",
          project: projectForRecipients,
          to: "",
          subject: "",
          status: "error",
          error: "NO_RECIPIENT",
          user: String(ctx.from?.id || ""),
          context: "missing_recipient",
        });
        return ctx.reply(
          "Bitte Empfaenger-E-Mail angeben (z.B. name@firma.at) oder Kontakte in der Gruppen-Chat speichern."
        );
      }

      await ctx.reply("✉️ Erstelle E-Mail-Entwurf… ⏳");

      try {
        const hours = getContextHours(ctx.from.id);
        const contextText = buildContextTextForHours(hours);
        const { subject, body } = await generateEmailDraft({
          userText,
          projectName: projectForRecipients,
          contextText,
        });

        pendingEmailByUser.set(ctx.from.id, {
          to: recipient,
          subject,
          body,
          project: projectForRecipients,
        });
        logEmailEvent({
          action: "draft",
          project: projectForRecipients,
          to: recipient,
          subject,
          status: "ok",
          user: String(ctx.from?.id || ""),
          context: "draft_created",
        });

        const hint = canSendEmail()
          ? "Antworte mit 'go' zum Senden oder 'abbrechen' zum Verwerfen."
          : "SMTP ist nicht konfiguriert. Entwurf gespeichert, aber Versand erst nach SMTP-Setup.";

        return ctx.reply(
          `📧 Entwurf:\nAn: ${recipient}\nBetreff: ${subject}\n\n${body}\n\n${hint}`
        );
      } catch (err) {
        console.error("[EMAIL] draft error:", err);
        return ctx.reply("❌ Fehler bei der E-Mail-Erstellung. Schau ins Terminal.");
      }
    }

    // Check for project creation intent
    try {
      const intentResult = await detectIntent(userText);
      if (intentResult && intentResult.intent === "project_creation") {
        const f = intentResult.fields || {};
        pendingProjectByUser.set(ctx.from.id, {
          fields: {
            code: f.kuerzel || "",
            name: f.name || "",
            contactName: f.contactName || "",
            contactEmail: f.contactEmail || "",
            cc: f.cc || undefined,
            client: f.client || "",
          },
        });
        const pf = pendingProjectByUser.get(ctx.from.id).fields;
        // Ask for first missing field
        if (!pf.code) return ctx.reply("Neues Projekt! Wie lautet das Kuerzel (z.B. SB-NORD-01)?");
        if (!pf.name) return ctx.reply(`Kuerzel: ${pf.code}\nWie heisst das Projekt?`);
        if (!pf.contactName) return ctx.reply(`Projekt: ${pf.code} - ${pf.name}\nAnsprechperson (Name)?`);
        if (!pf.contactEmail) return ctx.reply("E-Mail der Ansprechperson?");
        if (pf.cc === undefined) return ctx.reply("CC E-Mail? (oder 'nein')");
        if (!pf.client) return ctx.reply("Auftraggeber?");
        // All fields already extracted by AI
        const project = {
          code: pf.code,
          name: pf.name || pf.code,
          contactName: pf.contactName,
          contactEmail: pf.contactEmail,
          cc: pf.cc || "",
          client: pf.client || "",
          createdAt: new Date().toISOString(),
        };
        const saved = await upsertProject(project);
        pendingProjectByUser.delete(ctx.from.id);
        return ctx.reply(
          `Projekt erstellt: ${saved.code} - ${saved.name}\n` +
            `Ordner: ${path.join(DATA_ROOT, saved.dirName)}\n\n` +
            "Erstelle eine Telegram-Gruppe, fuege mich hinzu, und verknuepfe mit:\n" +
            `/verknuepfen ${saved.code}`
        );
      }
    } catch (err) {
      console.error("[INTENT] error:", err);
      // Fall through to general chat
    }

    await ctx.sendChatAction("typing");

    try {
      const pn = getProjectName(ctx.from.id);
      const hours = getContextHours(ctx.from.id);
      const contextText = buildContextTextForHours(hours);

      const answer = await chatProjectAI({
        userText,
        projectName: pn,
        contextText,
      });

      const chunks = answer.match(/[\s\S]{1,3500}/g) || [];
      for (const part of chunks) await ctx.reply(part);
    } catch (err) {
      console.error("[AI-CHAT] error:", err);
      await ctx.reply("KI-Fehler. Schau ins Terminal.");
    }
  }
});

// ========== START ==========
bot.launch();
console.log("✅ Bau-KKI Bot läuft");

// graceful shutdown
process.once("SIGINT", () => bot.stop("SIGINT"));
process.once("SIGTERM", () => bot.stop("SIGTERM"));
