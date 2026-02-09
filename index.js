require("dotenv").config();
const { Telegraf } = require("telegraf");
const Database = require("better-sqlite3");
const nodemailer = require("nodemailer");
const fs = require("fs");
const path = require("path");
const { generateReport, chatProjectAI, generateEmailDraft, classifyReportLines } = require("./ai");

const fetch = (...args) =>
  import("node-fetch").then(({ default: fetchImpl }) => fetchImpl(...args));
const {
  DATA_ROOT,
  ensureBaseStructure,
  upsertProject,
  getProjectByCode,
  generateReportFile,
} = require("./reports");

const BOT_TOKEN = process.env.BOT_TOKEN;
const OWNER_USER_ID = Number(process.env.OWNER_USER_ID);
const SMTP_HOST = process.env.SMTP_HOST;
const SMTP_PORT = Number(process.env.SMTP_PORT || 587);
const SMTP_USER = process.env.SMTP_USER;
const SMTP_PASS = process.env.SMTP_PASS;
const SMTP_FROM = process.env.SMTP_FROM || SMTP_USER;
const DEFAULT_EMAIL_TO = process.env.EMAIL_TO;
const META_DIR = path.join(DATA_ROOT, "_meta");
const CONTACTS_FILE = path.join(META_DIR, "contacts.csv");
const EMAIL_LOG_FILE = path.join(META_DIR, "email_logs.csv");

const ALLOWED_GROUP_IDS = (process.env.ALLOWED_GROUP_IDS || "")
  .split(",")
  .map((s) => s.trim())
  .filter(Boolean)
  .map(Number);

if (!BOT_TOKEN) throw new Error("BOT_TOKEN fehlt");
if (!OWNER_USER_ID) throw new Error("OWNER_USER_ID fehlt");

ensureBaseStructure();
const bot = new Telegraf(BOT_TOKEN);
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

// ======= KI-Chat State (nur Owner nutzt) =======
const aiMode = new Map(); // userId -> boolean
const projectNameByUser = new Map(); // userId -> string
const contextHoursByUser = new Map(); // userId -> number (optional, default 24)
const pendingEmailByUser = new Map(); // userId -> { to, subject, body, project }
const activeReportByChat = new Map(); // chatId -> { type, projectCode, lines, photos }

function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true });
}

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
  // Kontext aus ALLEN erlaubten Gruppen
  const rows = selectRecent.all(`-${hours} hours`).filter((r) =>
    ALLOWED_GROUP_IDS.includes(r.chat_id)
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

function parseReportStart(text) {
  const match = String(text || "").trim().match(/^(RB|BTB)\s+(.+)$/i);
  if (!match) return null;
  return { type: match[1].toUpperCase(), projectCode: match[2].trim() };
}

// ========== COMMANDS ==========
bot.start((ctx) => {
  ctx.reply(
    "👋 Bau-KI Bot ist aktiv.\n\n" +
      "Befehle:\n" +
      "/who – zeigt deine User-ID\n" +
      "/chatid – zeigt Chat-/Gruppen-ID\n" +
      "/allowthis – zeigt ob diese Gruppe erlaubt ist\n" +
      "/summary – Zusammenfassung (24h, alle erlaubten Gruppen)\n" +
      "/summary 6 – letzte 6 Stunden\n" +
      "/summarygroup – Zusammenfassung nur für diese Gruppe (24h)\n" +
      "/summarygroup 6 – nur diese Gruppe, letzte 6 Stunden\n" +
      "/report – KI Baustellenbericht (letzte 24h, nur diese Gruppe)\n\n" +
      "Projekte:\n" +
      "/newproject <Kuerzel> | <Projektname> | <Ansprechperson> | <Email> | <CC> | <Auftraggeber>\n\n" +
      "Privater Projekt-KI Chat:\n" +
      "/project <name> – Projekt setzen\n" +
      "/context <stunden> – Kontextzeitraum für KI-Chat (Default 24)\n" +
      "/ai – KI-Modus AN (privat chatten)\n" +
      "/ai off – KI-Modus AUS\n" +
      "/ask <frage> – einmalige Frage an KI (privat)\n\n" +
      "E-Mail Entwurf (privat):\n" +
      "Schreibe z.B.: KI schreib bitte eine Mail ueber den aktuellen Stand ..."
  );
});

bot.command("who", (ctx) => {
  ctx.reply(`Deine Telegram User-ID:\n${ctx.from.id}`);
});

// ✅ Chat/Gruppen-ID anzeigen
bot.command("chatid", (ctx) => {
  if (!isOwner(ctx)) return;
  const chatId = ctx.chat?.id;
  const title = ctx.chat?.title || ctx.chat?.username || "private";
  ctx.reply(`📌 Chat: ${title}\n🆔 Chat-ID: ${chatId}`);
});

// ✅ Prüfen ob aktuelle Gruppe erlaubt ist
bot.command("allowthis", (ctx) => {
  if (!isOwner(ctx)) return;
  const chatId = ctx.chat?.id;
  const title = ctx.chat?.title || ctx.chat?.username || "private";
  const allowed = ALLOWED_GROUP_IDS.includes(chatId);
  ctx.reply(
    allowed
      ? `✅ Erlaubt: ${title}\n🆔 ${chatId}`
      : `❌ Nicht erlaubt: ${title}\n🆔 ${chatId}\n\nTrage die ID in ALLOWED_GROUP_IDS ein und starte den Bot neu.`
  );
});

// ✅ Zusammenfassung über ALLE Gruppen (Owner only)
bot.command("summary", (ctx) => {
  if (!isOwner(ctx)) return ctx.reply("⛔ Keine Berechtigung");

  const arg = (ctx.message.text.split(" ")[1] || "").trim();
  const hours = arg && /^\d+$/.test(arg) ? Number(arg) : 24;

  const rows = selectRecent.all(`-${hours} hours`).filter((r) =>
    ALLOWED_GROUP_IDS.includes(r.chat_id)
  );

  if (!rows.length) {
    return ctx.reply(
      "Keine Nachrichten gefunden.\n" +
        "👉 Check: Bot ist in der Gruppe? /setprivacy in BotFather = Disable? Gruppe steht in ALLOWED_GROUP_IDS?"
    );
  }

  ctx.reply(summarize(rows));
});

// ✅ Zusammenfassung NUR für die aktuelle Gruppe
bot.command("summarygroup", (ctx) => {
  if (!isOwner(ctx)) return ctx.reply("⛔ Keine Berechtigung");

  const chat = ctx.chat;
  if (!chat || !["group", "supergroup"].includes(chat.type)) {
    return ctx.reply("Dieser Befehl funktioniert nur in einer Gruppe.");
  }

  if (!ALLOWED_GROUP_IDS.includes(chat.id)) {
    return ctx.reply(
      `❌ Diese Gruppe ist nicht erlaubt.\n` +
        `Nutze /chatid und trage die ID in ALLOWED_GROUP_IDS ein, dann Bot neu starten.`
    );
  }

  const arg = (ctx.message.text.split(" ")[1] || "").trim();
  const hours = arg && /^\d+$/.test(arg) ? Number(arg) : 24;

  const rows = selectRecentByChat.all(chat.id, `-${hours} hours`);
  ctx.reply(summarize(rows));
});

// ✅ KI Report (nur diese Gruppe, letzte 24h)
bot.command("report", async (ctx) => {
  if (!isOwner(ctx)) return ctx.reply("⛔ Keine Berechtigung");

  const chat = ctx.chat;
  if (!chat || !["group", "supergroup"].includes(chat.type)) {
    return ctx.reply("Bitte diesen Befehl in einer Gruppe ausführen.");
  }

  if (!ALLOWED_GROUP_IDS.includes(chat.id)) {
    return ctx.reply("❌ Diese Gruppe ist nicht erlaubt.");
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
  if (!isOwner(ctx)) return ctx.reply("⛔ Keine Berechtigung");
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
  if (!isOwner(ctx)) return ctx.reply("⛔ Keine Berechtigung");
  const name = ctx.message.text.replace(/^\/project\s*/i, "").trim();
  if (!name) return ctx.reply("Bitte Projektname angeben: /project Baustelle Nord");
  projectNameByUser.set(ctx.from.id, name);
  ctx.reply(`✅ Projekt gesetzt: ${name}`);
});

// Kontext-Zeitraum (in Stunden)
bot.command("context", (ctx) => {
  if (!isOwner(ctx)) return ctx.reply("⛔ Keine Berechtigung");
  const arg = (ctx.message.text.split(" ")[1] || "").trim();
  const hours = arg && /^\d+$/.test(arg) ? Number(arg) : 24;
  contextHoursByUser.set(ctx.from.id, hours);
  ctx.reply(`✅ Kontext-Zeitraum gesetzt: letzte ${hours} Stunden`);
});

// KI-Modus togglen
bot.command("ai", (ctx) => {
  if (!isOwner(ctx)) return ctx.reply("⛔ Keine Berechtigung");

  const arg = (ctx.message.text.split(" ")[1] || "").trim().toLowerCase();
  if (arg === "off") {
    aiMode.set(ctx.from.id, false);
    return ctx.reply("🧠 Projekt-KI: AUS");
  }

  aiMode.set(ctx.from.id, true);
  const pn = getProjectName(ctx.from.id);
  const h = getContextHours(ctx.from.id);
  return ctx.reply(
    `🧠 Projekt-KI: AN\nProjekt: ${pn}\nKontext: letzte ${h} Stunden aus erlaubten Gruppen\n\n` +
      `Schreib mir jetzt privat deine Frage.\n` +
      `Beenden: /ai off`
  );
});

// Einmalige Frage (privat empfohlen)
bot.command("ask", async (ctx) => {
  if (!isOwner(ctx)) return ctx.reply("⛔ Keine Berechtigung");

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
  if (!ALLOWED_GROUP_IDS.includes(chat.id)) return;
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
  if (!ALLOWED_GROUP_IDS.includes(chat.id)) return;
  if (!activeReportByChat.has(chat.id)) return;

  const doc = ctx.message?.document;
  if (!doc?.file_id) return;
  if (!String(doc.mime_type || "").startsWith("image/")) return;
  const state = activeReportByChat.get(chat.id);
  state.photos.push(doc.file_id);
});

bot.on("text", async (ctx) => {
  const chat = ctx.chat;

  // 1) Gruppen -> speichern
  if (chat && ["group", "supergroup"].includes(chat.type)) {
    if (!ALLOWED_GROUP_IDS.includes(chat.id)) return;

    const user =
      ctx.from?.username ||
      [ctx.from?.first_name, ctx.from?.last_name].filter(Boolean).join(" ") ||
      "unknown";

    const text = ctx.message?.text || "";
    const project = chat.title || String(chat.id);

    const startInfo = parseReportStart(text);
    if (startInfo) {
      const existing = activeReportByChat.get(chat.id);
      if (existing) {
        await ctx.reply(
          `⚠️ Es läuft bereits ein ${existing.type}-Bericht für ${existing.projectCode}. Schreibe 'Ende' um ihn zu beenden.`
        );
      } else {
        const projectEntry = getProjectByCode(startInfo.projectCode);
        if (!projectEntry) {
          await ctx.reply(
            `❌ Projekt '${startInfo.projectCode}' nicht gefunden. Lege es mit /newproject an.`
          );
        } else {
          activeReportByChat.set(chat.id, {
            type: startInfo.type,
            projectCode: startInfo.projectCode,
            lines: [],
            photos: [],
          });
          await ctx.reply(
            `✅ ${startInfo.type}-Bericht gestartet für ${startInfo.projectCode}.\n` +
              "Schreibe die Inhalte in die Gruppe.\n" +
              "Optional: AK: <Gruppe; Name> | MAT: <Menge; Einheit; Bezeichnung>\n" +
              "Beenden mit 'Ende'."
          );
        }
      }
    } else if (activeReportByChat.has(chat.id)) {
      const state = activeReportByChat.get(chat.id);
      if (text.trim().toLowerCase() === "ende") {
        const projectEntry = getProjectByCode(state.projectCode);
        if (!projectEntry) {
          activeReportByChat.delete(chat.id);
          await ctx.reply(
            `❌ Projekt '${state.projectCode}' nicht gefunden. Lege es mit /newproject an.`
          );
          return;
        }
        try {
          await ctx.reply("📄 Bericht wird erstellt…");
          let sections = null;
          try {
            sections = await classifyReportLines(state.lines);
          } catch (err) {
            console.error("[REPORT-CLASSIFY] error:", err);
          }
          const result = await generateReportFile({
            type: state.type,
            project: projectEntry,
            lines: state.lines,
            sections: sections || undefined,
          });
          try {
            await saveReportPhotos({
              ctx,
              state,
              photosDir: result.photosDir,
            });
          } catch (err) {
            console.error("[REPORT-PHOTOS] error:", err);
          }
          try {
            await ctx.replyWithDocument({ source: result.outFile });
          } catch (err) {
            console.error("[REPORT-UPLOAD] error:", err);
          }
          const savedMsg = [
            `✅ Bericht gespeichert in: ${result.reportFolder}`,
            `Excel: ${result.outFile}`,
          ].join("\n");
          await ctx.reply(savedMsg);
        } catch (err) {
          console.error("[REPORT-FILE] error:", err);
          await ctx.reply("❌ Fehler beim Erstellen des Berichts. Schau ins Terminal.");
        } finally {
          activeReportByChat.delete(chat.id);
        }
      } else {
        state.lines.push(text);
      }
    }

    const emails = extractEmails(text);
    const phones = extractPhones(text);
    appendContactRows(project, emails, phones, user, text);
    insertMsg.run(chat.id, chat.title || String(chat.id), user, text);
    return;
  }

  // 2) Private -> KI Chat (nur Owner + wenn aiMode aktiv)
  if (chat && chat.type === "private") {
    if (!isOwner(ctx)) return;
    const enabled = aiMode.get(ctx.from.id) === true;
    if (!enabled) return;

    const userText = ctx.message.text;
    const pending = pendingEmailByUser.get(ctx.from.id);
    const lower = userText.trim().toLowerCase();
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

    await ctx.reply("🧠 …");

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
      await ctx.reply("❌ KI-Fehler. Schau ins Terminal.");
    }
  }
});

// ========== START ==========
bot.launch();
console.log("✅ Bau-KKI Bot läuft");

// graceful shutdown
process.once("SIGINT", () => bot.stop("SIGINT"));
process.once("SIGTERM", () => bot.stop("SIGTERM"));
