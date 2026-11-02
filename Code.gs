/**
 * Synaptic Excel Academy — Google Apps Script
 *
 * INSTRUCCIONES DE DESPLIEGUE:
 * 1. Ve a script.google.com → Nuevo proyecto
 * 2. Pega este código completo
 * 3. Guarda con Ctrl+S
 * 4. Clic en "Desplegar" → "Nueva implementación"
 * 5. Tipo: Aplicación web
 * 6. Ejecutar como: Yo (tu cuenta)
 * 7. Quién tiene acceso: Cualquier persona
 * 8. Clic en "Desplegar" → copia la URL
 * 9. Pega esa URL en index.html donde dice TU_URL_APPS_SCRIPT_AQUI
 */

const SHEET_ID  = '1mEwajewLkehcLVWXungGTC-v9qYMSNgJTc6h6Pt3lUg';
const SHEET_NAME = 'Progreso Excel Academy';

function getOrCreateSheet() {
  let ss;
  if (SHEET_ID) {
    ss = SpreadsheetApp.openById(SHEET_ID);
  } else {
    // Busca o crea el spreadsheet
    const files = DriveApp.getFilesByName('Synaptic Excel Academy — Registros');
    if (files.hasNext()) {
      ss = SpreadsheetApp.open(files.next());
    } else {
      ss = SpreadsheetApp.create('Synaptic Excel Academy — Registros');
    }
  }

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // Encabezados
    sheet.getRange(1, 1, 1, 8).setValues([[
      'Timestamp', 'Evento', 'Nombre', 'Email', 'Nivel', 'Puntaje', 'Puntaje Máx', '% Logro'
    ]]);
    sheet.getRange(1, 1, 1, 8)
      .setBackground('#1a73e8')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, 8, 160);
  }
  return sheet;
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet();

    sheet.appendRow([
      data.timestamp   || new Date().toISOString(),
      data.evento      || '',
      data.nombre      || '',
      data.email       || '',
      data.nivel       || '',
      data.puntaje     || 0,
      data.puntaje_max || '',
      data.porcentaje  ? data.porcentaje + '%' : ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ok: true}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ok: false, error: err.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  // Para pruebas — retorna conteo de registros
  const sheet = getOrCreateSheet();
  const count = sheet.getLastRow() - 1;
  return ContentService
    .createTextOutput(JSON.stringify({ok: true, registros: count, sheet: sheet.getParent().getUrl()}))
    .setMimeType(ContentService.MimeType.JSON);
}
