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
// Columnas del tab "Invitados":
// Nombre | Apellido | Etiqueta | Telefono | Grupo | codigo | es_nino | confirmado | menu | notas | actualizado | visitas

var SHEET_NAME = 'Invitados';

// GET endpoint
function doGet(e) {
  var code = (e.parameter.code || '').trim().toLowerCase();
  var callback = e.parameter.callback || '';
  var action = e.parameter.action || 'get';

if (!code) {
    return respondGet({ error: 'no_code' }, callback);
  }

  // Save action
  if (action === 'save') {
    try {
      var responses = JSON.parse(e.parameter.data || '[]');
      return respondGet(saveResponses(code, responses), callback);
    } catch (err) {
      return respondGet({ error: err.message }, callback);
    }
  }

  // Get family data
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var col = {};
  headers.forEach(function(h, i) { col[h] = i; });

  var members = [];
  var groupName = '';
  var visitRows = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[col['codigo']]).trim().toLowerCase() === code) {
      groupName = row[col['Grupo']];
      var esNino = String(row[col['es_nino']]).trim().toLowerCase() === 'si';
      var confirmado = row[col['confirmado']];

      // Parse confirmado values
      if (confirmado === true || String(confirmado).toLowerCase() === 'si' || String(confirmado).toLowerCase() === 'true') {
        confirmado = true;
      } else if (confirmado === false || String(confirmado).toLowerCase() === 'no' || String(confirmado).toLowerCase() === 'false') {
        confirmado = false;
      } else if (String(confirmado).toLowerCase() === 'tal_vez') {
        confirmado = 'tal_vez';
      } else {
        confirmado = '';
      }

      members.push({
        row: i + 1,
        nombre: (row[col['Nombre']] + ' ' + row[col['Apellido']]).trim(),
        es_nino: esNino,
        confirmado: confirmado,
        menu: row[col['menu']] || '',
        notas: row[col['notas']] || ''
      });
      visitRows.push(i + 1);
    }
  }

  if (members.length === 0) {
    return respondGet({ error: 'not_found' }, callback);
  }

  // Increment visit counter
  if (col['visitas'] !== undefined) {
    visitRows.forEach(function(rowNum) {
      var currentVisits = Number(data[rowNum - 1][col['visitas']]) || 0;
      sheet.getRange(rowNum, col['visitas'] + 1).setValue(currentVisits + 1);
    });
  }

  return respondGet({
    familia: groupName,
    miembros: members
  }, callback);
}

// Return JSONP if callback is provided, otherwise plain JSON
function respondGet(data, callback) {
  if (callback) {
    var js = callback + '(' + JSON.stringify(data) + ')';
    return ContentService.createTextOutput(js).setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return jsonResponse(data);
}

// Save responses
function saveResponses(code, responses) {
  if (!code || responses.length === 0) {
    return { error: 'invalid_data' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var col = {};
  headers.forEach(function(h, i) { col[h] = i; });

  var now = new Date().toISOString();

  responses.forEach(function(resp) {
    var rowNum = resp.row;
    if (rowNum && rowNum > 1 && rowNum <= data.length) {
      var rowCode = String(data[rowNum - 1][col['codigo']]).trim().toLowerCase();
      if (rowCode === code) {
        // Store confirmado as: si / no / tal_vez
        var confirmadoVal = '';
        if (resp.confirmado === true) confirmadoVal = 'si';
        else if (resp.confirmado === false) confirmadoVal = 'no';
        else if (resp.confirmado === 'tal_vez') confirmadoVal = 'tal_vez';

        sheet.getRange(rowNum, col['confirmado'] + 1).setValue(confirmadoVal);
        sheet.getRange(rowNum, col['menu'] + 1).setValue(resp.menu || '');
        sheet.getRange(rowNum, col['notas'] + 1).setValue(resp.notas || '');
        sheet.getRange(rowNum, col['actualizado'] + 1).setValue(now);
      }
    }
  });

  return { success: true };
}

// POST endpoint (fallback)
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var code = (body.code || '').trim().toLowerCase();
    return jsonResponse(saveResponses(code, body.responses || []));
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// Generate unique codes per Grupo
function generateCodes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var col = {};
  headers.forEach(function(h, i) { col[h] = i; });

  var codeCol = col['codigo'];
  var groupCol = col['Grupo'];

  var existingCodes = {};
  var groupCodes = {};

  // Collect existing codes
  for (var i = 1; i < data.length; i++) {
    var code = String(data[i][codeCol]).trim();
    var group = String(data[i][groupCol]).trim();
    if (code && code !== '' && code !== 'undefined') {
      existingCodes[code] = true;
      groupCodes[group] = code;
    }
  }

  // Assign codes to groups without one
  var generated = 0;
  for (var i = 1; i < data.length; i++) {
    var code = String(data[i][codeCol]).trim();
    var group = String(data[i][groupCol]).trim();

    if (!code || code === '' || code === 'undefined') {
      if (groupCodes[group]) {
        sheet.getRange(i + 1, codeCol + 1).setValue(groupCodes[group]);
      } else {
        var newCode;
        do {
          newCode = randomCode(6);
        } while (existingCodes[newCode]);

        existingCodes[newCode] = true;
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

// Generate WhatsApp messages
function generateWhatsAppMessages() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var col = {};
  headers.forEach(function(h, i) { col[h] = i; });

  var baseUrl = 'https://maoaiz.github.io/wedding/?code=';
  var groups = {};

  for (var i = 1; i < data.length; i++) {
    var group = String(data[i][col['Grupo']]).trim();
    var code = String(data[i][col['codigo']]).trim();
    var phone = String(data[i][col['Telefono']]).trim();

    if (group && !groups[group]) {
      groups[group] = { code: code, phone: '' };
    }
    if (!groups[group].phone && phone && phone !== '' && phone !== 'undefined') {
      groups[group].phone = phone;
    }
  }

  // Create or clear Mensajes tab
  var msgSheet = ss.getSheetByName('Mensajes');
  if (!msgSheet) {
    msgSheet = ss.insertSheet('Mensajes');
  } else {
    msgSheet.clearContents();
  }

  // Headers
  msgSheet.getRange(1, 1, 1, 5).setValues([[
    'Grupo', 'Telefono', 'Codigo', 'Enlace RSVP', 'Enlace WhatsApp'
  ]]);

  // Data rows
  var entries = Object.entries(groups);
  entries.forEach(function(entry, i) {
    var group = entry[0];
    var info = entry[1];
    var url = baseUrl + info.code;
    var ring = String.fromCodePoint(0x1F48D);
    var bride = String.fromCodePoint(0x1F470);
    var groom = String.fromCodePoint(0x1F935);
    var party = String.fromCodePoint(0x1F389);
    var heart = String.fromCodePoint(0x2764) + String.fromCodePoint(0xFE0F);
    var down = String.fromCodePoint(0x1F447);
    var pray = String.fromCodePoint(0x1F64F);
    var message = 'Hola ' + group + '! ' + ring + '\n\n' +
      'Queremos compartir contigo una noticia que nos llena de alegr\u00eda: \u00a1nos casamos! ' + bride + groom + party + '\n\n' +
      'Nos encantar\u00eda que nos acompa\u00f1aras en este d\u00eda tan especial. ' + heart + '\n\n' +
      'Por favor confirma tu asistencia antes del 6 de mayo ' + down + '\n' + url + '\n\n' +
      'Un abrazo, Eyla y Mauricio ' + pray;
    var encodedMsg = encodeURIComponent(message);
    var phone = info.phone || '';
    var cleanPhone = phone ? phone.replace(/[^0-9]/g, '') : '';
    var waLink = cleanPhone ? 'https://wa.me/' + cleanPhone + '?text=' + encodedMsg : '';

    var row = i + 2;
    msgSheet.getRange(row, 1, 1, 3).setValues([[group, phone, info.code]]);
    msgSheet.getRange(row, 4).setFormula('=HYPERLINK("' + url + '", "Ver invitaci\u00f3n")');
    if (waLink) {
      msgSheet.getRange(row, 5).setFormula('=HYPERLINK("' + waLink.replace(/"/g, '""') + '", "Enviar WhatsApp")');
    }
  });

  msgSheet.autoResizeColumns(1, 5);

  SpreadsheetApp.getUi().alert(
    'Mensajes generados: ' + entries.length + ' grupos.\n' +
    'Revisa la pestana "Mensajes".'
  );
}

// Utility: random code
function randomCode(length) {
  var chars = 'abcdefghjkmnpqrstuvwxyz23456789';
  var code = '';
  for (var i = 0; i < length; i++) {
    code += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return code;
}

// Utility: JSON response
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// Create or update Resumen tab with summary formulas
function createResumen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Find column letters
  var colLetter = {};
  headers.forEach(function(h, i) {
    colLetter[h] = String.fromCharCode(65 + i); // A, B, C...
    if (i >= 26) colLetter[h] = 'A' + String.fromCharCode(65 + i - 26); // AA, AB...
  });

  var conf = colLetter['confirmado'];
  var menu = colLetter['menu'];
  var nino = colLetter['es_nino'];
  var nombre = colLetter['Nombre'];
  var etiqueta = colLetter['Etiqueta'];
  var visitas = colLetter['visitas'];

  var resSheet = ss.getSheetByName('Resumen');
  if (!resSheet) {
    resSheet = ss.insertSheet('Resumen');
  } else {
    resSheet.clearContents();
  }

  var data = [
    ['RESUMEN GENERAL', ''],
    ['', ''],
    ['Total invitados', '=COUNTA(Invitados!' + nombre + '2:' + nombre + ')'],
    ['Confirmados (si)', '=COUNTIF(Invitados!' + conf + '2:' + conf + ', "si")'],
    ['No asisten', '=COUNTIF(Invitados!' + conf + '2:' + conf + ', "no")'],
    ['A\u00fan no saben', '=COUNTIF(Invitados!' + conf + '2:' + conf + ', "tal_vez")'],
    ['Sin responder', '=COUNTBLANK(Invitados!' + conf + '2:' + conf + ')'],
    ['', ''],
    ['MENUS', ''],
    ['Normal', '=COUNTIF(Invitados!' + menu + '2:' + menu + ', "normal")'],
    ['Vegetariano', '=COUNTIF(Invitados!' + menu + '2:' + menu + ', "vegetariano")'],
    ['Infantil', '=COUNTIF(Invitados!' + menu + '2:' + menu + ', "infantil")'],
    ['', ''],
    ['NI\u00d1OS', ''],
    ['Total ni\u00f1os', '=COUNTIF(Invitados!' + nino + '2:' + nino + ', "Si")'],
    ['', ''],
    ['POR ETIQUETA', ''],
    ['Brice\u00f1os', '=COUNTIF(Invitados!' + etiqueta + '2:' + etiqueta + ', "*Brice\u00f1os*")'],
    ['Amigos Eyla', '=COUNTIF(Invitados!' + etiqueta + '2:' + etiqueta + ', "*Amigos Eyla*")'],
    ['Amigos Mauricio', '=COUNTIF(Invitados!' + etiqueta + '2:' + etiqueta + ', "*Amigos Mauricio*")'],
    ['Pekin', '=COUNTIF(Invitados!' + etiqueta + '2:' + etiqueta + ', "*Pekin*")'],
    ['Astros', '=COUNTIF(Invitados!' + etiqueta + '2:' + etiqueta + ', "*Astros*")'],
    ['Aizagas', '=COUNTIF(Invitados!' + etiqueta + '2:' + etiqueta + ', "*Aizagas*")'],
    ['iKono', '=COUNTIF(Invitados!' + etiqueta + '2:' + etiqueta + ', "*iKono*")'],
    ['UTP', '=COUNTIF(Invitados!' + etiqueta + '2:' + etiqueta + ', "*UTP*")'],
    ['Pasto', '=COUNTIF(Invitados!' + etiqueta + '2:' + etiqueta + ', "*Pasto*")'],
    ['', ''],
    ['ENGAGEMENT', ''],
    ['Invitados que han abierto el enlace', '=COUNTIF(Invitados!' + visitas + '2:' + visitas + ', ">"&0)'],
    ['Total visitas', '=SUM(Invitados!' + visitas + '2:' + visitas + ')'],
  ];

  resSheet.getRange(1, 1, data.length, 2).setValues(data);

  // Format headers
  var boldRows = [1, 9, 14, 17, 28];
  boldRows.forEach(function(r) {
    resSheet.getRange(r, 1, 1, 2).setFontWeight('bold');
  });

  resSheet.setColumnWidth(1, 280);
  resSheet.setColumnWidth(2, 100);

  SpreadsheetApp.getUi().alert('Pesta\u00f1a Resumen creada.');
}

// Menu
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('RSVP')
    .addItem('Generar codigos', 'generateCodes')
    .addItem('Generar mensajes WhatsApp', 'generateWhatsAppMessages')
    .addItem('Crear Resumen', 'createResumen')
    .addToUi();
}
