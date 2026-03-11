const axios = require("axios");

const OLLAMA_URL = process.env.OLLAMA_URL || "http://localhost:11434/api/generate";
const OLLAMA_CHAT_URL = process.env.OLLAMA_CHAT_URL || "http://localhost:11434/api/chat";
const OLLAMA_MODEL = process.env.OLLAMA_MODEL || "qwen2.5:3b";

async function ollama(prompt) {
  const { data } = await axios.post(
    OLLAMA_URL,
    { model: OLLAMA_MODEL, prompt, stream: false },
    { timeout: 300000 }
  );
  return (data.response || "").trim();
}

async function ollamaChat(messages) {
  const { data } = await axios.post(
    OLLAMA_CHAT_URL,
    { model: OLLAMA_MODEL, messages, stream: false },
    { timeout: 300000 }
  );
  return (data.message?.content || "").trim();
}

function buildRBSystemPrompt({ projectCode, projectName, client, rulesBlock }) {
  return `Du bist ein Baustellenassistent fuer Regieberichte. Du hilfst dem Benutzer, einen Regiebericht zu erstellen.

Deine Aufgabe:
1. Verstehe natuerliche Sprache und extrahiere daraus Berichtsdaten
2. Ordne Informationen in drei Kategorien ein:
   - leistungen: ausgefuehrte Taetigkeiten/Arbeiten (kurz und praegnant)
   - arbeitskraefte: Personal mit Stunden (Format: "<Anzahl> <Rolle> je <Stunden>h" z.B. "2 Monteure je 8h" oder "<Name> <Stunden>h")
   - material: verwendetes Material (Format: "<Menge>; <Einheit>; <Bezeichnung>" z.B. "10; m; Kupferrohr")
3. Bestaetige was du verstanden hast und zeige den aktuellen Stand
4. Stelle Rueckfragen wenn wichtige Informationen fehlen (z.B. Stunden, Anzahl Mitarbeiter)
5. Wenn der Benutzer Korrekturen macht, passe die Daten entsprechend an

Antworte IMMER in der Sprache des Benutzers.

Antworte IMMER exakt in diesem Format:
---RESPONSE---
<Deine natuerliche Antwort an den Benutzer - kurz, klar, mit aktuellem Stand>
---DATA---
{"leistungen":[...],"arbeitskraefte":[...],"material":[...]}

WICHTIG: Der DATA-Block muss IMMER den VOLLSTAENDIGEN aktuellen Stand aller drei Kategorien enthalten (nicht nur neue Eintraege).
Wenn nichts geaendert wurde, gib den bisherigen Stand zurueck.
Gib im DATA-Block NUR valides JSON zurueck, ohne Markdown-Formatierung.
${rulesBlock ? "\n" + rulesBlock : ""}

Projekt: ${projectCode}${projectName ? " - " + projectName : ""}
${client ? "Auftraggeber: " + client : ""}`.trim();
}

function buildBTBSystemPrompt({ projectCode, projectName, client, rulesBlock }) {
  return `Du bist ein Baustellenassistent fuer Bautagesberichte (BTB). Du hilfst dem Benutzer, einen Bautagesbericht zu erstellen.

Deine Aufgabe:
1. Verstehe natuerliche Sprache und extrahiere daraus Berichtsdaten
2. Ordne Informationen in drei Kategorien ein:
   - leistungen: ausgefuehrte Taetigkeiten/Arbeiten des Tages (kurz und praegnant, max 10 Eintraege)
   - arbeitskraefte: Personal mit Anzahl und Rolle (Format: "<Anzahl> <Rolle>" z.B. "5 Monteure", "2 Helfer", "1 Polier", "3 Bau-FA")
     Bekannte Rollen: Projektleiter, Bauleiter, Polier, gehob.Bau-FA, Spezialbau-FA, Werkpolier, Bauvorarbeiter, Bau-FA, Baufachwerker, Bauwerker, Maschinist, Kranfahrer, Fremdfirmen, Monteur, Installateur, Helfer
   - material: verwendetes/geliefertes Material (Format: "<Menge>; <Einheit>; <Bezeichnung>" z.B. "10; m; Kupferrohr")
3. Bestaetige was du verstanden hast und zeige den aktuellen Stand
4. Stelle Rueckfragen wenn wichtige Informationen fehlen
5. Wenn der Benutzer Korrekturen macht, passe die Daten entsprechend an

Ein Bautagesbericht dokumentiert den gesamten Tagesablauf auf der Baustelle:
- Alle ausgefuehrten Leistungen/Arbeiten
- Alle anwesenden Arbeitskraefte nach Rolle/Gewerk
- Alles verwendete/angelieferte Material

Antworte IMMER in der Sprache des Benutzers.

Antworte IMMER exakt in diesem Format:
---RESPONSE---
<Deine natuerliche Antwort an den Benutzer - kurz, klar, mit aktuellem Stand>
---DATA---
{"leistungen":[...],"arbeitskraefte":[...],"material":[...]}

WICHTIG: Der DATA-Block muss IMMER den VOLLSTAENDIGEN aktuellen Stand aller drei Kategorien enthalten.
Gib im DATA-Block NUR valides JSON zurueck, ohne Markdown-Formatierung.
${rulesBlock ? "\n" + rulesBlock : ""}

Projekt: ${projectCode}${projectName ? " - " + projectName : ""}
${client ? "Auftraggeber: " + client : ""}`.trim();
}

function buildReportSystemPrompt({ type, projectCode, projectName, client, rulesBlock }) {
  if (type === "BTB") {
    return buildBTBSystemPrompt({ projectCode, projectName, client, rulesBlock });
  }
  return buildRBSystemPrompt({ projectCode, projectName, client, rulesBlock });
}

async function detectIntent(userText) {
  const prompt = `Analysiere die folgende Nachricht und bestimme die Absicht des Benutzers.

Moegliche Absichten:
1. "project_creation" - Der Benutzer will ein neues Bauprojekt/Bauvorhaben erstellen
2. "general_chat" - Allgemeine Frage oder Unterhaltung

Wenn "project_creation": Extrahiere alle genannten Felder:
- kuerzel: Projektkuerzel/Code
- name: Projektname
- contactName: Ansprechperson
- contactEmail: E-Mail
- cc: CC E-Mail
- client: Auftraggeber

Gib NUR JSON zurueck, ohne Markdown:
{"intent":"project_creation|general_chat","fields":{"kuerzel":"","name":"","contactName":"","contactEmail":"","cc":"","client":""}}

Nachricht: "${userText}"`;

  const raw = await ollama(prompt);
  try {
    const jsonText = raw.trim().replace(/^```json\s*|```\s*$/g, "").trim();
    return JSON.parse(jsonText);
  } catch (e) {
    return { intent: "general_chat", fields: {} };
  }
}

async function chatReportAI({ systemPrompt, conversationHistory, userMessage }) {
  const messages = [
    { role: "system", content: systemPrompt },
    ...conversationHistory,
    { role: "user", content: userMessage }
  ];

  const fullResponse = await ollamaChat(messages);

  // Parse structured response
  let responsePart = fullResponse;
  let reportData = null;

  const responseIdx = fullResponse.indexOf("---RESPONSE---");
  const dataIdx = fullResponse.indexOf("---DATA---");

  function extractJSON(dataPart) {
    // Remove markdown code fences
    let jsonText = dataPart.replace(/^```json\s*/g, "").replace(/```\s*$/g, "").trim();
    // Extract first valid JSON object using brace matching
    const start = jsonText.indexOf("{");
    if (start < 0) return null;
    let depth = 0;
    let end = -1;
    for (let i = start; i < jsonText.length; i++) {
      if (jsonText[i] === "{") depth++;
      else if (jsonText[i] === "}") depth--;
      if (depth === 0) { end = i; break; }
    }
    if (end < 0) return null;
    return jsonText.substring(start, end + 1);
  }

  function flattenToString(item) {
    if (typeof item === "string") return item;
    if (item && typeof item === "object") {
      // Handle objects like {qty: "4", unit: "lfm", desc: "UW50"}
      const values = Object.values(item).filter(v => v !== null && v !== undefined && v !== "");
      return values.join("; ");
    }
    return String(item || "");
  }

  function parseReportJSON(dataPart) {
    try {
      const jsonText = extractJSON(dataPart);
      if (!jsonText) return null;
      const parsed = JSON.parse(jsonText);
      const toStringArray = (arr) => (Array.isArray(arr) ? arr.map(flattenToString).filter(Boolean) : []);
      return {
        leistungen: toStringArray(parsed.leistungen),
        arbeitskraefte: toStringArray(parsed.arbeitskraefte),
        material: toStringArray(parsed.material),
      };
    } catch (e) {
      console.error("[CHAT-REPORT-AI] JSON parse error:", e.message);
      return null;
    }
  }

  if (responseIdx >= 0 && dataIdx >= 0) {
    responsePart = fullResponse.substring(responseIdx + "---RESPONSE---".length, dataIdx).trim();
    const dataPart = fullResponse.substring(dataIdx + "---DATA---".length).trim();
    reportData = parseReportJSON(dataPart);
  } else if (dataIdx >= 0) {
    responsePart = fullResponse.substring(0, dataIdx).trim();
    const dataPart = fullResponse.substring(dataIdx + "---DATA---".length).trim();
    reportData = parseReportJSON(dataPart);
  }

  return { response: responsePart, reportData };
}

async function extractRuleFromCorrection({ userMessage, previousData, correctedData }) {
  const prompt = `Der Benutzer hat eine Korrektur an einem Baustellenbericht vorgenommen.

Vorherige Daten: ${JSON.stringify(previousData)}
Korrigierte Daten: ${JSON.stringify(correctedData)}
Benutzer sagte: "${userMessage}"

Extrahiere eine allgemeine Regel, die in Zukunft bei ALLEN Regieberichten angewendet werden soll.
Die Regel soll verallgemeinert sein (nicht spezifisch fuer diesen einen Bericht).

Gib NUR JSON zurueck, ohne Markdown:
{"trigger":"<Wann gilt diese Regel - kurz>","rule":"<Die Regel als klare Anweisung>"}

Wenn keine sinnvolle allgemeine Regel extrahiert werden kann (z.B. nur ein Tippfehler korrigiert wurde), gib zurueck:
{"trigger":"","rule":""}`;

  const raw = await ollama(prompt);
  try {
    const jsonText = raw.trim().replace(/^```json\s*|```\s*$/g, "").trim();
    const parsed = JSON.parse(jsonText);
    if (parsed.trigger && parsed.rule) {
      return { trigger: parsed.trigger, rule: parsed.rule };
    }
    return null;
  } catch (e) {
    console.error("[EXTRACT-RULE] parse error:", e.message);
    return null;
  }
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
    const toStrArr = (arr) => (Array.isArray(arr) ? arr.map(v => typeof v === "string" ? v : (v && typeof v === "object" ? Object.values(v).filter(Boolean).join("; ") : String(v || ""))).filter(Boolean) : []);
    return {
      leistungen: toStrArr(data.leistungen),
      arbeitskraefte: toStrArr(data.arbeitskraefte),
      material: toStrArr(data.material),
    };
  } catch (err) {
    return null;
  }
}

module.exports = {
  chatProjectAI,
  generateReport,
  generateEmailDraft,
  classifyReportLines,
  chatReportAI,
  extractRuleFromCorrection,
  buildReportSystemPrompt,
  detectIntent,
};
