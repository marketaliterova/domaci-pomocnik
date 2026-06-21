/**
 * Domácí Pomocník - Hue / Apple Home Webhook
 *
 * Setup:
 * 1. https://script.google.com → Nový projekt
 * 2. Smaž ukázkový kód, vlož TENHLE celý
 * 3. Uloz (Cmd+S)
 * 4. Klepni "Deploy" → "New deployment"
 * 5. Type: Web app
 * 6. Execute as: Me
 * 7. Who has access: Anyone
 * 8. Deploy → autorizuj svým Google účtem
 * 9. Zkopíruj URL kterou ti to dá (končí na /exec)
 * 10. V appce v admin → Webhook tlačítka → vlož URL do "Webhook URL"
 *
 * Apple Home setup pro každé tlačítko:
 * - V app klepni "📋 Kopír" vedle tlačítka → zkopíruje URL s ?press=KEY
 * - iPhone → Domácnost → Automatizace → "+" → "Když je stisknuto tlačítko"
 * - Vyber Hue Dimmer Switch a typ stisku (ON/OFF/UP/DOWN)
 * - Pokračovat → vyber libovolnou akci (např. zapni libovolné světlo)
 * - DALŠÍ obrazovka → klepni dolů "Convert to Shortcut"
 * - Přidej akci "Načíst obsah URL"
 * - Vlož URL kterou jsi zkopírovala
 * - Hotovo
 *
 * Apple TV musí být doma a zapnutý (slouží jako home hub).
 */

// Otestuj v Apps Script editoru - zavolá doGet s press=test
function testWebhook() {
  const result = doGet({ parameter: { press: 'test' } });
  Logger.log(result.getContent());
}

function doGet(e) {
  const p = (e && e.parameter) || {};

  // Webhook: Apple Home Shortcut volá ?press=KEY
  if (p.press) {
    const key = String(p.press).slice(0, 50);
    const events = getEvents();
    events.unshift({ ts: new Date().toISOString(), key: key });
    if (events.length > 50) events.length = 50;
    setEvents(events);
    return text('OK ' + key);
  }

  // App polluje ?action=events&since=ISO_TIMESTAMP
  if (p.action === 'events') {
    const events = getEvents();
    const since = p.since;
    const filtered = since ? events.filter(function(ev) { return ev.ts > since; }) : events;
    return json(filtered);
  }

  // Manuální reset (smaž historii) - GET ?action=clear
  if (p.action === 'clear') {
    setEvents([]);
    return text('Historie smazána');
  }

  // Health check
  return text('Domácí Pomocník Webhook OK. Použij ?press=KEY pro zápis nebo ?action=events pro čtení.');
}

function doPost(e) {
  // Podporujeme i POST pro případ že Apple Home pošle POST místo GET
  return doGet(e);
}

function getEvents() {
  const s = PropertiesService.getScriptProperties().getProperty('hue_events');
  return s ? JSON.parse(s) : [];
}

function setEvents(events) {
  PropertiesService.getScriptProperties().setProperty('hue_events', JSON.stringify(events));
}

function text(t) {
  return ContentService.createTextOutput(t).setMimeType(ContentService.MimeType.TEXT);
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
