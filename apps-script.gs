/**
 * Domácí pomocník — Apps Script backend
 * v4 — přidán OpenAI TTS endpoint (lidský hlas pro Rikiho)
 * Vygenerováno: 2026-06-23 18:59
 *
 * SETUP (jednorázově):
 *   1) Project Settings (⚙️) → Script Properties → "+ Add property"
 *      Property: OPENAI_API_KEY
 *      Value:    sk-proj-... (tvůj OpenAI API klíč)
 *   2) Save → Deploy → Manage deployments → Edit → New version → Deploy
 *
 * ENDPOINTY:
 *   GET  ?press=KEY              → zaznamenat stisk tlačítka (Hue/Apple Home)
 *   GET  ?action=events          → vrátí seznam stisků
 *   GET  ?action=events&since=ts → vrátí stisky od daného času
 *   GET  ?action=clear           → smaže historii stisků
 *   GET  ?action=clearstate      → smaže sync state
 *   GET                          → vrátí aktuální sync state
 *   GET  ?action=chat&prompt=...&systemPrompt=...&history=[...] → AI chat (Riki)
 *   GET  ?action=tts&text=...&voice=nova → vrátí base64 MP3 (lidský hlas)
 *   POST { action: 'pull' }      → vrátí sync state
 *   POST { action: 'push', state: ... } → uloží sync state
 */

const SHEET_ID = '1kDE1Vnjp-hpEJG6Y8NII2V1JOMNmDFdZnCg2E2u9lsU';
const SHEET_NAME = 'data';

function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange('A1').setValue('state');
    sheet.getRange('B1').setValue('updated');
  }
  return sheet;
}

function readState() {
  const sheet = getSheet();
  const val = sheet.getRange('A2').getValue();
  if (!val) return null;
  const s = String(val);
  if (s.startsWith('{') && !s.startsWith('{"') && s.includes('=')) {
    return null;
  }
  return s;
}

function writeState(stateValue) {
  const sheet = getSheet();
  const stateStr = typeof stateValue === 'string' ? stateValue : JSON.stringify(stateValue);
  sheet.getRange('A2').setValue(stateStr);
  sheet.getRange('B2').setValue(new Date().toISOString());
}

function getEvents() {
  const s = PropertiesService.getScriptProperties().getProperty('hue_events');
  return s ? JSON.parse(s) : [];
}

function setEvents(events) {
  PropertiesService.getScriptProperties().setProperty('hue_events', JSON.stringify(events));
}

function doGet(e) {
  try {
    const p = (e && e.parameter) || {};

    // v3: AI CHAT — Riki přes OpenAI GPT-4o-mini
    if (p.action === 'chat') {
      return handleChat_(p);
    }

    // v4: TTS — lidský hlas přes OpenAI Audio
    if (p.action === 'tts') {
      return handleTTS_(p);
    }

    if (p.press) {
      const key = String(p.press).slice(0, 50);
      const events = getEvents();
      events.unshift({ ts: new Date().toISOString(), key: key });
      if (events.length > 50) events.length = 50;
      setEvents(events);
      return ContentService.createTextOutput('OK ' + key).setMimeType(ContentService.MimeType.TEXT);
    }
    if (p.action === 'events') {
      const events = getEvents();
      const since = p.since;
      const filtered = since ? events.filter(function(ev) { return ev.ts > since; }) : events;
      return ContentService.createTextOutput(JSON.stringify(filtered)).setMimeType(ContentService.MimeType.JSON);
    }
    if (p.action === 'clear') {
      setEvents([]);
      return ContentService.createTextOutput('Historie smazána').setMimeType(ContentService.MimeType.TEXT);
    }
    if (p.action === 'clearstate') {
      writeState('');
      return ContentService.createTextOutput('State smazán').setMimeType(ContentService.MimeType.TEXT);
    }
    const sheet = getSheet();
    const state = readState();
    const updated = sheet.getRange('B2').getValue();
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, state: state, updated: updated ? String(updated) : null }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    if (body.action === 'pull') {
      const sheet = getSheet();
      const state = readState();
      const updated = sheet.getRange('B2').getValue();
      return ContentService.createTextOutput(JSON.stringify({ ok: true, state: state, updated: updated ? String(updated) : null })).setMimeType(ContentService.MimeType.JSON);
    }
    if (body.action === 'push') {
      writeState(body.state);
      return ContentService.createTextOutput(JSON.stringify({ ok: true, updated: new Date().toISOString() })).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'unknown action' })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  }
}

// =====================================================================
// v3: AI CHAT — Riki přes OpenAI GPT-4o-mini
// =====================================================================
function handleChat_(p) {
  const apiKey = PropertiesService.getScriptProperties()
                   .getProperty('OPENAI_API_KEY');

  if (!apiKey) {
    return ContentService
      .createTextOutput(JSON.stringify({
        error: 'OPENAI_API_KEY není v Script Properties. Project Settings → Script Properties → Add property.'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const systemPrompt = p.systemPrompt || 'Jsi Riki, kamarádský mascot.';
  let history = [];
  try { history = JSON.parse(p.history || '[]'); } catch (err) {}

  const messages = [{ role: 'system', content: systemPrompt }].concat(history);

  try {
    const resp = UrlFetchApp.fetch(
      'https://api.openai.com/v1/chat/completions',
      {
        method: 'post',
        contentType: 'application/json',
        headers: { 'Authorization': 'Bearer ' + apiKey },
        payload: JSON.stringify({
          model: 'gpt-4o-mini',
          messages: messages,
          max_tokens: 200,
          temperature: 0.8
        }),
        muteHttpExceptions: true
      }
    );

    const data = JSON.parse(resp.getContentText());

    if (data.choices && data.choices[0] && data.choices[0].message) {
      const text = data.choices[0].message.content;
      return ContentService
        .createTextOutput(JSON.stringify({ text: text }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(JSON.stringify({
        error: 'OpenAI nevrátil odpověď',
        raw: data
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({
        error: 'Apps Script chyba: ' + err.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// =====================================================================
// v4: TTS — Lidský hlas přes OpenAI Audio API
// Voice options: alloy, echo, fable, onyx, nova, shimmer
// Model: tts-1 (rychlejší, levnější) nebo tts-1-hd (vyšší kvalita)
// =====================================================================
function handleTTS_(p) {
  const apiKey = PropertiesService.getScriptProperties()
                   .getProperty('OPENAI_API_KEY');

  if (!apiKey) {
    return ContentService
      .createTextOutput(JSON.stringify({
        error: 'OPENAI_API_KEY není v Script Properties.'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const text = (p.text || '').slice(0, 4000); // OpenAI limit 4096
  const voice = p.voice || 'nova';
  const model = p.model || 'tts-1';

  if (!text) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'Prázdný text' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    const resp = UrlFetchApp.fetch(
      'https://api.openai.com/v1/audio/speech',
      {
        method: 'post',
        contentType: 'application/json',
        headers: { 'Authorization': 'Bearer ' + apiKey },
        payload: JSON.stringify({
          model: model,
          voice: voice,
          input: text,
          response_format: 'mp3'
        }),
        muteHttpExceptions: true
      }
    );

    if (resp.getResponseCode() !== 200) {
      var errText = resp.getContentText();
      return ContentService
        .createTextOutput(JSON.stringify({
          error: 'OpenAI TTS chyba: ' + errText.substring(0, 500)
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Binary MP3 → base64
    const blob = resp.getBlob();
    const base64Audio = Utilities.base64Encode(blob.getBytes());

    return ContentService
      .createTextOutput(JSON.stringify({
        audio: base64Audio,
        format: 'mp3',
        voice: voice
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({
        error: 'Apps Script TTS chyba: ' + err.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// =====================================================================
// TEST RUNNERS — spusť v editoru aby se ověřil setup
// =====================================================================
function testChatLocally() {
  const apiKey = PropertiesService.getScriptProperties()
                   .getProperty('OPENAI_API_KEY');
  if (!apiKey) {
    Logger.log('❌ OPENAI_API_KEY není v Script Properties!');
    return;
  }
  Logger.log('✓ Klíč je nastaven (' + apiKey.substring(0, 10) + '...)');

  const result = handleChat_({
    prompt: 'Ahoj Riki!',
    systemPrompt: 'Jsi Riki. Odpověz česky, jednou krátkou větou.',
    history: '[]'
  });
  Logger.log('Odpověď: ' + result.getContent());
}

function testTTSLocally() {
  const apiKey = PropertiesService.getScriptProperties()
                   .getProperty('OPENAI_API_KEY');
  if (!apiKey) {
    Logger.log('❌ OPENAI_API_KEY není v Script Properties!');
    return;
  }
  Logger.log('✓ Klíč je nastaven, testuji TTS...');

  const result = handleTTS_({
    text: 'Ahoj, já jsem Riki. Slyšíš mě?',
    voice: 'nova'
  });
  const json = JSON.parse(result.getContent());
  if (json.audio) {
    Logger.log('✓ TTS funguje! Audio délka (base64): ' + json.audio.length + ' znaků');
  } else {
    Logger.log('❌ TTS chyba: ' + result.getContent());
  }
}
