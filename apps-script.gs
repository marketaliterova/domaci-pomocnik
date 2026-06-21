/**
 * Domácí pomocník — Apps Script backend (sync + webhook)
 *
 * Dělá DVĚ věci:
 *  1. SYNC: ukládá/načítá stav appky (jako dřív)
 *  2. WEBHOOK: přijímá GET /?press=KEY z Apple Home Shortcut, loguje event,
 *              app polluje /?action=events pro nové stisky
 *
 * INSTALACE (jednorázově, ~5 min):
 *
 * 1. Otevři SVOJI EXISTUJÍCÍ Apps Script (přes Sheet → Rozšíření → Apps Script
 *    nebo přímo script.google.com s tvojim projektem)
 * 2. Smaž celý kód a nahraď tímto souborem (Ctrl+A, Ctrl+V)
 * 3. Nahraď SHEET_ID_SEM níže za tvoje ID (pokud tam už máš, neměň)
 * 4. Cmd+S (uložit)
 * 5. Klikni Nasadit (Deploy) → SPRAVOVAT NASAZENÍ (Manage deployments)
 * 6. Vyber existující deployment → klikni tužku ✏️ → Verze → "New version"
 * 7. Deploy → autorizuj (pokud je to potřeba)
 * 8. URL ZŮSTÁVÁ STEJNÁ - příkazy se rozšířily, neztrácíš stávající sync
 *
 * Hotovo. URL kterou už máš (https://script.google.com/macros/s/.../exec)
 * teď zvládá oboje — vlož ji v appce do "Synchronizace URL" I do "Webhook URL".
 */

const SHEET_ID = 'SHEET_ID_SEM';   // <--- TADY NAHRAĎ (pokud máš starý script, ID zachovej)
const SHEET_NAME = 'data';          // sync sheet (vytvoří se sám)

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

// ============= WEBHOOK STORAGE (in-memory přes PropertiesService) =============

function getEvents() {
  const s = PropertiesService.getScriptProperties().getProperty('hue_events');
  return s ? JSON.parse(s) : [];
}

function setEvents(events) {
  PropertiesService.getScriptProperties().setProperty('hue_events', JSON.stringify(events));
}

// ============= HTTP HANDLERS =============

function doGet(e) {
  try {
    const p = (e && e.parameter) || {};

    // 1. WEBHOOK: Apple Home Shortcut → ?press=KEY
    if (p.press) {
      const key = String(p.press).slice(0, 50);
      const events = getEvents();
      events.unshift({ ts: new Date().toISOString(), key: key });
      if (events.length > 50) events.length = 50;
      setEvents(events);
      return ContentService.createTextOutput('OK ' + key).setMimeType(ContentService.MimeType.TEXT);
    }

    // 2. APP POLLING: ?action=events&since=TIMESTAMP
    if (p.action === 'events') {
      const events = getEvents();
      const since = p.since;
      const filtered = since ? events.filter(function(ev) { return ev.ts > since; }) : events;
      return ContentService.createTextOutput(JSON.stringify(filtered)).setMimeType(ContentService.MimeType.JSON);
    }

    // 3. ADMIN: smazat history ?action=clear
    if (p.action === 'clear') {
      setEvents([]);
      return ContentService.createTextOutput('Historie smazána').setMimeType(ContentService.MimeType.TEXT);
    }

    // 4. SYNC PULL: vrátí současný state (default)
    const sheet = getSheet();
    const state = sheet.getRange('A2').getValue();
    const updated = sheet.getRange('B2').getValue();
    return ContentService
      .createTextOutput(JSON.stringify({
        ok: true,
        state: state || null,
        updated: updated || null
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const sheet = getSheet();

    if (body.action === 'pull') {
      const state = sheet.getRange('A2').getValue();
      const updated = sheet.getRange('B2').getValue();
      return ContentService
        .createTextOutput(JSON.stringify({
          ok: true,
          state: state || null,
          updated: updated || null
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (body.action === 'push') {
      sheet.getRange('A2').setValue(body.state || '');
      sheet.getRange('B2').setValue(new Date().toISOString());
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: 'unknown action' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============= TEST FUNCTIONS (spustitelné v editoru) =============

function testWebhookPress() {
  const result = doGet({ parameter: { press: 'test-button' } });
  Logger.log(result.getContent());
}

function testWebhookEvents() {
  const result = doGet({ parameter: { action: 'events' } });
  Logger.log(result.getContent());
}
