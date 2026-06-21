/**
 * Domácí pomocník — Apps Script backend (sync + webhook)
 * v2 — opravený bug s ukládáním (objekt → string)
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
