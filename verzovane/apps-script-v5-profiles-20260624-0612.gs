/**
 * Domácí pomocník — Apps Script backend
 * v5 — sync per-profile (Alex a Albie mají oddělená data v Sheet)
 * Vygenerováno: 2026-06-24 06:12
 *
 * SETUP (jednorázově):
 *   1) Project Settings (⚙️) → Script Properties → "+ Add property"
 *      Property: OPENAI_API_KEY
 *      Value:    sk-proj-... (tvůj OpenAI API klíč)
 *   2) Save → Deploy → Manage deployments → Edit → New version → Deploy
 *
 * SHEET STRUKTURA (data sheet):
 *   A1: 'profile'    B1: 'state'    C1: 'updated'
 *   A2: 'alex'       B2: {...}      C2: timestamp
 *   A3: 'albert'     B3: {...}      C3: timestamp
 *
 * MIGRACE: starý formát (1 řádek bez profile) se automaticky přesune do "alex"
 *
 * ENDPOINTY:
 *   GET  ?press=KEY              → zaznamenat stisk tlačítka (Hue/Apple Home)
 *   GET  ?action=events          → vrátí seznam stisků
 *   GET  ?action=events&since=ts → vrátí stisky od daného času
 *   GET  ?action=clear           → smaže historii stisků
 *   GET  ?action=clearstate&profile=X → smaže sync state daného profilu
 *   GET                          → vrátí aktuální sync state (default alex)
 *   GET  ?action=chat&prompt=...&systemPrompt=...&history=[...] → AI chat (Riki)
 *   GET  ?action=tts&text=...&voice=nova → vrátí base64 MP3 (lidský hlas)
 *   POST { action: 'pull', profile: 'alex' }  → vrátí sync state daného profilu
 *   POST { action: 'push', profile: 'alex', state: ... } → uloží sync state profilu
 */

const SHEET_ID = '1kDE1Vnjp-hpEJG6Y8NII2V1JOMNmDFdZnCg2E2u9lsU';
const SHEET_NAME = 'data';
const PROFILES = ['alex', 'albert'];

function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange('A1').setValue('profile');
    sheet.getRange('B1').setValue('state');
    sheet.getRange('C1').setValue('updated');
  }
  return sheet;
}

// v5: najde řádek pro daný profil (nebo vytvoří). Vrátí řádek 1-based.
function findOrCreateProfileRow_(sheet, profile) {
  const lastRow = Math.max(2, sheet.getLastRow());
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues(); // sloupec A
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === profile) {
      return i + 2; // řádek (1-based)
    }
  }
  // Nový řádek
  const newRow = lastRow + 1;
  sheet.getRange(newRow, 1).setValue(profile);
  return newRow;
}

// v5: čte state pro daný profil. Migrace: pokud profile='alex' a A2 vypadá jako starý formát (B2 obsahuje state ale A2 ne 'alex'), přidělí ho Alexovi.
function readStateForProfile_(profile) {
  const sheet = getSheet();
  // Migrace starého formátu
  migrateLegacyFormat_(sheet);
  const row = findRowForProfile_(sheet, profile);
  if (!row) return { state: null, updated: null };
  const state = sheet.getRange(row, 2).getValue();
  const updated = sheet.getRange(row, 3).getValue();
  return {
    state: state ? String(state) : null,
    updated: updated ? String(updated) : null
  };
}

function findRowForProfile_(sheet, profile) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === profile) {
      return i + 2;
    }
  }
  return null;
}

function writeStateForProfile_(profile, stateValue) {
  const sheet = getSheet();
  // Migrace
  migrateLegacyFormat_(sheet);
  const row = findOrCreateProfileRow_(sheet, profile);
  const stateStr = typeof stateValue === 'string' ? stateValue : JSON.stringify(stateValue);
  sheet.getRange(row, 2).setValue(stateStr);
  sheet.getRange(row, 3).setValue(new Date().toISOString());
}

// v5: migrace starého formátu (A1=state, A2={...}) → nový (A1=profile, B1=state, C1=updated)
function migrateLegacyFormat_(sheet) {
  const a1 = String(sheet.getRange('A1').getValue() || '').toLowerCase();
  // Pokud A1 už je 'profile', migrace proběhla
  if (a1 === 'profile') return;
  // Starý formát: A1='state', B1='updated', A2={...}, B2=timestamp
  if (a1 === 'state' || !a1) {
    const oldState = sheet.getRange('A2').getValue();
    const oldUpdated = sheet.getRange('B2').getValue();
    // Vyčistit a přenastavit headery
    sheet.clear();
    sheet.getRange('A1').setValue('profile');
    sheet.getRange('B1').setValue('state');
    sheet.getRange('C1').setValue('updated');
    if (oldState) {
      // Stará data → Alex
      sheet.getRange('A2').setValue('alex');
      sheet.getRange('B2').setValue(oldState);
      sheet.getRange('C2').setValue(oldUpdated || new Date().toISOString());
    }
  }
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

    if (p.action === 'chat') {
      return handleChat_(p);
    }

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
      const prof = String(p.profile || 'alex').toLowerCase();
      writeStateForProfile_(prof, '');
      return ContentService.createTextOutput('State smazán pro ' + prof).setMimeType(ContentService.MimeType.TEXT);
    }
    // Default: vrátí state Alexe (pro zpětnou kompatibilitu)
    const result = readStateForProfile_(String(p.profile || 'alex').toLowerCase());
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, state: result.state, updated: result.updated }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    // v5: profile parametr (default alex pro zpětnou kompatibilitu)
    const profile = String(body.profile || 'alex').toLowerCase();
    if (PROFILES.indexOf(profile) < 0) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'unknown profile: ' + profile })).setMimeType(ContentService.MimeType.JSON);
    }

    if (body.action === 'pull') {
      const result = readStateForProfile_(profile);
      return ContentService.createTextOutput(JSON.stringify({
        ok: true,
        profile: profile,
        state: result.state,
        updated: result.updated
      })).setMimeType(ContentService.MimeType.JSON);
    }
    if (body.action === 'push') {
      writeStateForProfile_(profile, body.state);
      return ContentService.createTextOutput(JSON.stringify({
        ok: true,
        profile: profile,
        updated: new Date().toISOString()
      })).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'unknown action' })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  }
}

// =====================================================================
// AI CHAT — Riki přes OpenAI GPT-4o-mini
// =====================================================================
function handleChat_(p) {
  const apiKey = PropertiesService.getScriptProperties()
                   .getProperty('OPENAI_API_KEY');

  if (!apiKey) {
    return ContentService
      .createTextOutput(JSON.stringify({
        error: 'OPENAI_API_KEY není v Script Properties.'
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
// TTS — Lidský hlas přes OpenAI Audio API
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

  const text = (p.text || '').slice(0, 4000);
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
      const errText = resp.getContentText();
      return ContentService
        .createTextOutput(JSON.stringify({
          error: 'OpenAI TTS chyba: ' + errText.substring(0, 500)
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

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
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) { Logger.log('❌ OPENAI_API_KEY není v Script Properties!'); return; }
  Logger.log('✓ Klíč je nastaven');
  const result = handleChat_({ prompt: 'Ahoj!', systemPrompt: 'Jsi Riki. Odpověz krátce.', history: '[]' });
  Logger.log('Odpověď: ' + result.getContent());
}

function testTTSLocally() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) { Logger.log('❌ OPENAI_API_KEY není v Script Properties!'); return; }
  Logger.log('✓ Klíč nastaven, testuji TTS...');
  const result = handleTTS_({ text: 'Ahoj!', voice: 'nova' });
  const json = JSON.parse(result.getContent());
  if (json.audio) Logger.log('✓ TTS funguje! Audio: ' + json.audio.length + ' znaků base64');
  else Logger.log('❌ ' + result.getContent());
}

// v5: Test že per-profile sync funguje
function testProfileSync() {
  Logger.log('--- Testuji per-profile sync ---');

  // Test pull pro Alexe
  const alexState = readStateForProfile_('alex');
  Logger.log('Alex state: ' + (alexState.state ? alexState.state.length + ' znaků (' + alexState.updated + ')' : 'NIC'));

  // Test pull pro Alberta/Albieho
  const albertState = readStateForProfile_('albert');
  Logger.log('Albie state: ' + (albertState.state ? albertState.state.length + ' znaků (' + albertState.updated + ')' : 'NIC'));

  Logger.log('--- Sheet struktura ---');
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  Logger.log('Řádků: ' + lastRow);
  if (lastRow >= 1) {
    const headers = sheet.getRange(1, 1, 1, 3).getValues()[0];
    Logger.log('Headers: ' + JSON.stringify(headers));
  }
  for (let i = 2; i <= lastRow; i++) {
    const row = sheet.getRange(i, 1, 1, 3).getValues()[0];
    Logger.log('Řádek ' + i + ': profile=' + row[0] + ', state=' + (row[1] ? String(row[1]).substring(0, 50) + '...' : 'NIC') + ', updated=' + row[2]);
  }
}
