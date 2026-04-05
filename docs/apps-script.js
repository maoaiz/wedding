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
//
// Sheet columns (Invitados tab):
// First Name | Last Name | Tags | Phone | Group | code | es_nino | confirmado | menu | notas | actualizado

const SHEET_NAME = 'Invitados';

// GET endpoint: fetch family data by code
function doGet(e) {
  const code = (e.parameter.code || '').trim().toLowerCase();

  if (!code) {
    return jsonResponse({ error: 'no_code' });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const col = {};
  headers.forEach((h, i) => col[h] = i);

  const members = [];
  let groupName = '';

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[col['code']]).trim().toLowerCase() === code) {
      groupName = row[col['Group']];
      const esNino = String(row[col['es_nino']]).trim().toLowerCase() === 'si';
      members.push({
        row: i + 1,
        nombre: (row[col['First Name']] + ' ' + row[col['Last Name']]).trim(),
        es_nino: esNino,
        confirmado: row[col['confirmado']],
        menu: row[col['menu']] || '',
        notas: row[col['notas']] || ''
      });
    }
  }

  if (members.length === 0) {
    return jsonResponse({ error: 'not_found' });
  }

  return jsonResponse({
    familia: groupName,
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
    const sheet = ss.getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    const col = {};
    headers.forEach((h, i) => col[h] = i);

    const now = new Date().toISOString();

    responses.forEach(resp => {
      const rowNum = resp.row;
      if (rowNum && rowNum > 1 && rowNum <= data.length) {
        const rowCode = String(data[rowNum - 1][col['code']]).trim().toLowerCase();
        if (rowCode === code) {
          sheet.getRange(rowNum, col['confirmado'] + 1).setValue(resp.confirmado);
          sheet.getRange(rowNum, col['menu'] + 1).setValue(resp.menu || '');
          sheet.getRange(rowNum, col['notas'] + 1).setValue(resp.notas || '');
          sheet.getRange(rowNum, col['actualizado'] + 1).setValue(now);
        }
      }
    });

    return jsonResponse({ success: true });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// Generate unique codes per Group
function generateCodes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const col = {};
  headers.forEach((h, i) => col[h] = i);

  const codeCol = col['code'];
  const groupCol = col['Group'];

  const existingCodes = new Set();
  const groupCodes = {};

  // Collect existing codes
  for (let i = 1; i < data.length; i++) {
    const code = String(data[i][codeCol]).trim();
    const group = String(data[i][groupCol]).trim();
    if (code && code !== '' && code !== 'undefined') {
      existingCodes.add(code);
      groupCodes[group] = code;
    }
  }

  // Assign codes to groups without one
  let generated = 0;
  for (let i = 1; i < data.length; i++) {
    const code = String(data[i][codeCol]).trim();
    const group = String(data[i][groupCol]).trim();

    if (!code || code === '' || code === 'undefined') {
      if (groupCodes[group]) {
        sheet.getRange(i + 1, codeCol + 1).setValue(groupCodes[group]);
      } else {
        let newCode;
        do {
          newCode = randomCode(6);
        } while (existingCodes.has(newCode));

        existingCodes.add(newCode);
        groupCodes[group] = newCode;
        sheet.getRange(i + 1, codeCol + 1).setValue(newCode);
        generated++;
      }
    }
  }

  SpreadsheetApp.getUi().alert(
    'Codigos generados: ' + generated + ' nuevos.\n' +
    'Total grupos: ' + Object.keys(groupCodes).length
  );
}

// Generate WhatsApp links per group
// Creates a "Mensajes" tab with: Group | Phone | Code | URL | WhatsApp Link | Message
function generateWhatsAppMessages() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const col = {};
  headers.forEach((h, i) => col[h] = i);

  const baseUrl = 'https://maoaiz.github.io/wedding/rsvp.html?code=';
  const groups = {};

  // Collect unique groups with their code and phone
  for (let i = 1; i < data.length; i++) {
    const group = String(data[i][col['Group']]).trim();
    const code = String(data[i][col['code']]).trim();
    const phone = String(data[i][col['Phone']]).trim();

    if (group && !groups[group]) {
      groups[group] = { code: code, phone: phone };
    }
    // If a row has a phone, use it (some rows might have it, others not)
    if (phone && phone !== '' && phone !== 'undefined') {
      groups[group].phone = phone;
    }
  }

  // Create or clear Mensajes tab
  let msgSheet = ss.getSheetByName('Mensajes');
  if (!msgSheet) {
    msgSheet = ss.insertSheet('Mensajes');
  } else {
    msgSheet.clearContents();
  }

  // Headers
  msgSheet.getRange(1, 1, 1, 6).setValues([[
    'Group', 'Phone', 'Code', 'RSVP URL', 'WhatsApp Link', 'Mensaje'
  ]]);

  // Data rows
  const entries = Object.entries(groups);
  entries.forEach((entry, i) => {
    const group = entry[0];
    const info = entry[1];
    const url = baseUrl + info.code;
    const message = 'Hola ' + group + '! Estan invitados a nuestra boda el 6 de junio de 2026. Confirma tu asistencia aqui: ' + url;
    const encodedMsg = encodeURIComponent(message);
    const phone = info.phone || '';
    const waLink = phone ? 'https://wa.me/' + phone.replace(/[^0-9]/g, '') + '?text=' + encodedMsg : '';

    msgSheet.getRange(i + 2, 1, 1, 6).setValues([[
      group, phone, info.code, url, waLink, message
    ]]);
  });

  // Auto-resize columns
  msgSheet.autoResizeColumns(1, 6);

  SpreadsheetApp.getUi().alert(
    'Mensajes generados: ' + entries.length + ' grupos.\n' +
    'Revisa la pestana "Mensajes".\n\n' +
    'Los grupos con telefono tendran un enlace de WhatsApp directo.'
  );
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

// Add custom menu
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('RSVP')
    .addItem('Generar codigos', 'generateCodes')
    .addItem('Generar mensajes WhatsApp', 'generateWhatsAppMessages')
    .addToUi();
}
