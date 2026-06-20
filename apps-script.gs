/**
 * Domácí pomocník v7 — Apps Script backend pro sync
 *
 * INSTALACE (jednorázově, ~5 min):
 *
 * 1. Vytvoř nový Google Sheet (libovolný název, např. "Domácí pomocník sync")
 *    https://sheets.new
 * 2. Z URL si vykopíruj SHEET_ID (mezi /d/ a /edit):
 *    https://docs.google.com/spreadsheets/d/SHEET_ID/edit
 * 3. V Sheetu: Rozšíření → Apps Script
 * 4. Smaž defaultní kód a nahraď tímto souborem (Ctrl+A, Ctrl+V)
 * 5. Nahraď SHEET_ID_SEM níže za tvoje ID
 * 6. Klikni na Nasadit (Deploy) → Nasazení (New deployment) → ozubené kolečko → Web app
 * 7. Nastav:
 *    - Description: domaci-pomocnik
 *    - Execute as: Me (tvoje@gmail.com)
 *    - Who has access: Anyone
 * 8. Klikni Deploy → autorizuj přes Google
 * 9. Zkopíruj Web App URL (končí /exec)
 * 10. V appce: Rodičovský režim → Nastavení → Synchronizace → vlož URL
 *
 * Hotovo. Tablet a mobil se teď synchronizují.
 *
 * --- BEZPEČNOST ---
 * URL je veřejná, ale neobsahuje nic citlivého — útočník by maximálně mohl číst
 * tvoje data o rutinách dítěte. Pokud to vadí, můžeš později přidat jednoduchý
 * auth token jako další parametr (viz komentář dole).
 */

const SHEET_ID = 'SHEET_ID_SEM';   // <--- TADY NAHRAĎ
const SHEET_NAME = 'data';          // název listu (vytvoří se sám)

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

function doGet(e) {
  try {
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

    // Volitelné: kontrola tokenu (pro mírnou bezpečnost)
    // const EXPECTED_TOKEN = 'tvuj-tajny-token';
    // if (body.token !== EXPECTED_TOKEN) {
    //   return ContentService
    //     .createTextOutput(JSON.stringify({ ok: false, error: 'unauthorized' }))
    //     .setMimeType(ContentService.MimeType.JSON);
    // }

    if (body.action === 'pull') {
      // Stejné jako doGet
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
      const stateStr = JSON.stringify(body.state || {});
      const now = new Date().toISOString();
      sheet.getRange('A2').setValue(stateStr);
      sheet.getRange('B2').setValue(now);
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, updated: now }))
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
