require("dotenv").config();
const { Telegraf } = require("telegraf");
const Database = require("better-sqlite3");
const { generateReport, chatProjectAI } = require("./ai");

const BOT_TOKEN = process.env.BOT_TOKEN;
const OWNER_USER_ID = Number(process.env.OWNER_USER_ID);

const ALLOWED_GROUP_IDS = (process.env.ALLOWED_GROUP_IDS || "")
  .split(",")
  .map((s) => s.trim())
  .filter(Boolean)
  .map(Number);

if (!BOT_TOKEN) throw new Error("BOT_TOKEN fehlt");
if (!OWNER_USER_ID) throw new Error("OWNER_USER_ID fehlt");

const bot = new Telegraf(BOT_TOKEN);
const db = new Database("messages.db");

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
      "Privater Projekt-KI Chat:\n" +
      "/project <name> – Projekt setzen\n" +
      "/context <stunden> – Kontextzeitraum für KI-Chat (Default 24)\n" +
      "/ai – KI-Modus AN (privat chatten)\n" +
      "/ai off – KI-Modus AUS\n" +
      "/ask <frage> – einmalige Frage an KI (privat)"
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
    insertMsg.run(chat.id, chat.title || String(chat.id), user, text);
    return;
  }

  // 2) Private -> KI Chat (nur Owner + wenn aiMode aktiv)
  if (chat && chat.type === "private") {
    if (!isOwner(ctx)) return;
    const enabled = aiMode.get(ctx.from.id) === true;
    if (!enabled) return;

    const userText = ctx.message.text;
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
