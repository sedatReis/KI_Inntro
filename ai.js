const axios = require("axios");

const OLLAMA_URL = process.env.OLLAMA_URL || "http://localhost:11434/api/generate";
const OLLAMA_MODEL = process.env.OLLAMA_MODEL || "llama3";

async function ollama(prompt) {
  const { data } = await axios.post(
    OLLAMA_URL,
    { model: OLLAMA_MODEL, prompt, stream: false },
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

async function generateEmailDraft({ userText, projectName, contextText }) {
  const prompt = `
Du bist ein Assistent, der professionelle E-Mails fuer eine Baufirma schreibt.
Schreibe eine kurze, klare E-Mail auf Deutsch.
Gib exakt dieses Format aus:
Subject: <Betreffzeile>
Body:
<E-Mail-Text als Klartext, ohne Markdown>

Anforderungen:
- konkret, sachlich, freundlich
- wenn sinnvoll: kurze Aufzaehlung im Text (mit "-" pro Zeile)
- beende mit "Naechste Schritte:" und maximal 3 Punkten

Projekt: ${projectName || "Allgemein"}
${contextText ? `Kontext aus Telegram-Gruppen:\n${contextText}\n` : ""}

User-Anfrage: ${userText}
`.trim();

  const raw = await ollama(prompt);
  const subjectMatch = raw.match(/^Subject:\s*(.+)$/im);
  const bodyMatch = raw.match(/^Body:\s*([\s\S]*)$/im);

  const subject = subjectMatch ? subjectMatch[1].trim() : `Baustellen-Update: ${projectName || "Allgemein"}`;
  const body = bodyMatch ? bodyMatch[1].trim() : raw.trim();

  return { subject, body };
}

async function classifyReportLines(lines) {
  const prompt = `
Du bist ein Assistent fuer Baustellenberichte.
Ordne die folgenden Zeilen in drei Kategorien ein:
1) leistungen
2) arbeitskraefte
3) material

Gib NUR JSON zurueck, ohne Markdown:
{"leistungen":[],"arbeitskraefte":[],"material":[]}

Regeln:
- Teile Inhalte, wenn mehrere Themen in einer Zeile stehen.
- Gib nur Inhalte zurueck, die in den Zeilen vorkommen. Erfinde nichts.
- Die Zeilen kommen ab jetzt in dieser Reihenfolge: zuerst leistungen, dann arbeitskraefte, dann material.
  Wenn keine Labels vorhanden sind, nutze die Reihenfolge zur Zuordnung.
- arbeitskraefte: Mitarbeiter/Berufsgruppe (z.B. Monteur, Installateur, Elektriker, Helfer, Team) + Stunden/Arbeitszeit/Schicht/Zeiten.
  Normalisiere das Format:
  - Anzahl: "<Anzahl> <Rolle> je <Stunden>h" (z.B. "2 Monteure je 12h", "2 Installateure je 4h")
  - Person: "<Name> <Stunden>h" (z.B. "Mueller 8h")
  - Optional Gruppe/FA: "<Gruppe>; <Name> <Stunden>h"
  - Wenn in der Quelle "je 4" / "je 8" ohne Einheit steht und es um arbeitskraefte geht, als Stunden verstehen und mit "h" ausgeben.
  Teile mehrere Personen in einzelne Eintraege.
- material: Menge + Einheit + Bezeichnung. Normalisiere zu "Menge; Einheit; Bezeichnung".
  Beispiele: "10; m; Rohr", "3; Stk; Duebel", "1,5; l; Farbe".
  Wenn "Material" oder "Lieferung" erwaehnt wird, als material.
  Teile mehrere Materialien in einzelne Eintraege.
- leistungen: ausgefuehrte Taetigkeiten ohne Personal- oder Materialangaben.
- Wenn unklar, ordne als leistungen ein.

Zeilen:
${(lines || []).map((l) => `- ${l}`).join("\n")}
`.trim();

  const raw = await ollama(prompt);
  try {
    const jsonText = raw.trim().replace(/^```json|```$/g, "").trim();
    const data = JSON.parse(jsonText);
    return {
      leistungen: Array.isArray(data.leistungen) ? data.leistungen : [],
      arbeitskraefte: Array.isArray(data.arbeitskraefte) ? data.arbeitskraefte : [],
      material: Array.isArray(data.material) ? data.material : [],
    };
  } catch (err) {
    return null;
  }
}

module.exports = { chatProjectAI, generateReport, generateEmailDraft, classifyReportLines };
