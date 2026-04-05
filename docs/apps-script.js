// ============================================
// Google Apps Script - Wedding RSVP API
// ============================================
// Instructions:
// 1. Open your Google Sheet
// 2. Go to Extensions > Apps Script
// 3. Paste this code in Code.gs
// 4. Deploy > New deployment > Web app
//    - Execute as: Me
//    - Who has access: Anyone
// 5. Copy the deployment URL and paste it in rsvp.html

const SHEET_NAME_GUESTS = 'Invitados';
const SHEET_NAME_CODES = 'Codigos';

// GET endpoint: fetch family data by code
function doGet(e) {
  const code = (e.parameter.code || '').trim().toLowerCase();

  if (!code) {
    return jsonResponse({ error: 'no_code' });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_GUESTS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const colIndex = {};
  headers.forEach((h, i) => colIndex[h] = i);

  const members = [];
  let familyName = '';

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[colIndex['code']]).trim().toLowerCase() === code) {
      familyName = row[colIndex['familia']];
      members.push({
        row: i + 1, // 1-indexed for Sheets
        nombre: row[colIndex['nombre']],
        es_nino: row[colIndex['es_nino']] === true || String(row[colIndex['es_nino']]).toLowerCase() === 'true',
        confirmado: row[colIndex['confirmado']],
        menu: row[colIndex['menu']] || '',
        notas: row[colIndex['notas']] || ''
      });
    }
  }

  if (members.length === 0) {
    return jsonResponse({ error: 'not_found' });
  }

  return jsonResponse({
    familia: familyName,
    miembros: members
  });
}

// POST endpoint: save RSVP responses
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const code = (body.code || '').trim().toLowerCase();
    const responses = body.responses || [];

    if (!code || responses.length === 0) {
      return jsonResponse({ error: 'invalid_data' });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME_GUESTS);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    const colIndex = {};
    headers.forEach((h, i) => colIndex[h] = i);

    const now = new Date().toISOString();

    responses.forEach(resp => {
      const rowNum = resp.row;
      if (rowNum && rowNum > 1 && rowNum <= data.length) {
        // Verify the row belongs to the same code
        const rowCode = String(data[rowNum - 1][colIndex['code']]).trim().toLowerCase();
        if (rowCode === code) {
          sheet.getRange(rowNum, colIndex['confirmado'] + 1).setValue(resp.confirmado);
          sheet.getRange(rowNum, colIndex['menu'] + 1).setValue(resp.menu || '');
          sheet.getRange(rowNum, colIndex['notas'] + 1).setValue(resp.notas || '');
          sheet.getRange(rowNum, colIndex['actualizado'] + 1).setValue(now);
        }
      }
    });

    return jsonResponse({ success: true });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// Generate unique codes for families that don't have one
function generateCodes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_GUESTS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const colIndex = {};
  headers.forEach((h, i) => colIndex[h] = i);

  const codeCol = colIndex['code'];
  const familiaCol = colIndex['familia'];

  // Collect existing codes
  const existingCodes = new Set();
  const familyCodes = {};

  for (let i = 1; i < data.length; i++) {
    const code = String(data[i][codeCol]).trim();
    const familia = String(data[i][familiaCol]).trim();
    if (code && code !== '' && code !== 'undefined') {
      existingCodes.add(code);
      familyCodes[familia] = code;
    }
  }

  // Assign codes to families without one
  for (let i = 1; i < data.length; i++) {
    const code = String(data[i][codeCol]).trim();
    const familia = String(data[i][familiaCol]).trim();

    if (!code || code === '' || code === 'undefined') {
      if (familyCodes[familia]) {
        // Use existing code for this family
        sheet.getRange(i + 1, codeCol + 1).setValue(familyCodes[familia]);
      } else {
        // Generate new code
        let newCode;
        do {
          newCode = randomCode(6);
        } while (existingCodes.has(newCode));

        existingCodes.add(newCode);
        familyCodes[familia] = newCode;
        sheet.getRange(i + 1, codeCol + 1).setValue(newCode);
      }
    }
  }

  // Update Codigos tab
  const codesSheet = ss.getSheetByName(SHEET_NAME_CODES);
  if (codesSheet) {
    // Clear existing data (keep headers)
    const lastRow = codesSheet.getLastRow();
    if (lastRow > 1) {
      codesSheet.getRange(2, 1, lastRow - 1, 4).clearContent();
    }

    // Write codes
    const families = Object.entries(familyCodes);
    families.forEach((entry, i) => {
      codesSheet.getRange(i + 2, 1).setValue(entry[1]); // code
      codesSheet.getRange(i + 2, 2).setValue(entry[0]); // familia
      codesSheet.getRange(i + 2, 3).setValue(false);     // enviado
      codesSheet.getRange(i + 2, 4).setValue('');         // fecha_envio
    });
  }

  SpreadsheetApp.getUi().alert('Codigos generados: ' + Object.keys(familyCodes).length + ' familias.');
}

// Generate WhatsApp messages
function generateWhatsAppMessages() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const codesSheet = ss.getSheetByName(SHEET_NAME_CODES);
  const data = codesSheet.getDataRange().getValues();

  const baseUrl = 'https://maoaiz.github.io/wedding/rsvp.html?code=';
  let messages = 'MENSAJES DE WHATSAPP\n\n';

  for (let i = 1; i < data.length; i++) {
    const code = data[i][0];
    const familia = data[i][1];
    const url = baseUrl + code;

    messages += '--- ' + familia + ' ---\n';
    messages += 'Hola ' + familia + '! Estan invitados a nuestra boda el 6 de junio de 2026.\n';
    messages += 'Confirma tu asistencia aqui: ' + url + '\n\n';
  }

  // Create a new sheet with the messages
  let msgSheet = ss.getSheetByName('Mensajes');
  if (!msgSheet) {
    msgSheet = ss.insertSheet('Mensajes');
  }
  msgSheet.getRange(1, 1).setValue(messages);

  SpreadsheetApp.getUi().alert('Mensajes generados en la pestaña "Mensajes".');
}

// Utility: generate random alphanumeric code
function randomCode(length) {
  const chars = 'abcdefghjkmnpqrstuvwxyz23456789'; // no ambiguous chars
  let code = '';
  for (let i = 0; i < length; i++) {
    code += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return code;
}

// Utility: return JSON response
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// Add custom menu to the spreadsheet
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('RSVP')
    .addItem('Generar codigos', 'generateCodes')
    .addItem('Generar mensajes WhatsApp', 'generateWhatsAppMessages')
    .addToUi();
}
