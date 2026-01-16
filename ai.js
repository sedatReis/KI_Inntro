const axios = require("axios");

async function ollama(prompt) {
  const { data } = await axios.post(
    "http://localhost:11434/api/generate",
    { model: "llama3", prompt, stream: false },
    { timeout: 120000 }
  );
  return (data.response || "").trim();
}

async function chatProjectAI({ userText, projectName, contextText }) {
  const prompt = `
Du bist ein KI-Assistent für eine Baufirma.
Du beantwortest NUR projektbezogene Fragen (Bauablauf, Termine, Lieferungen, Mängel, To-Dos, Risiken, Kommunikation).
Wenn eine Frage nicht zum Projekt passt, frage nach dem Projektbezug oder lenke zurück: "Welche Baustelle / welches Gewerk / welcher Abschnitt?"

Antwort-Stil:
- kurz & klar
- wenn sinnvoll: Bulletpoints
- immer mit: "Nächste Schritte" (max. 3 Punkte)

Projekt: ${projectName || "Allgemein"}
${contextText ? `Kontext aus Telegram-Gruppen:\n${contextText}\n` : ""}

User: ${userText}
Antwort:
`.trim();

  return await ollama(prompt);
}

async function generateReport(messages, groupName) {
  const content = messages.map(m => `- ${m.user}: ${m.text}`).join("\n");
  const prompt = `
Du bist ein erfahrener Bauleiter.
Erstelle aus folgenden Chat-Nachrichten einen strukturierten Baustellenbericht.

Format:
1) Kurzlage (1-2 Sätze)
2) Lieferungen
3) Termine
4) Probleme/Risiken
5) Maßnahmen / Nächste Schritte

Nachrichten:
${content}

Bericht:
`.trim();

  const text = await ollama(prompt);
  return `📋 KI-Bericht – ${groupName}\n\n${text}`;
}

module.exports = { chatProjectAI, generateReport };
