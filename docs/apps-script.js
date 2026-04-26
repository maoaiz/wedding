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
// Nombre | Apellido | Etiqueta | Telefono | Grupo | codigo | es_nino | confirmado | menu | notas | actualizado | visitas | ultima_visita
// (ultima_visita es opcional; si la columna existe, doGet la actualiza al timestamp de cada visita)

var SHEET_NAME = 'Invitados';

// GET endpoint
function doGet(e) {
  var code = (e.parameter.code || '').trim().toLowerCase();
  var callback = e.parameter.callback || '';
  var action = e.parameter.action || 'get';

  // Refresh dropdown only (safe, doesn't reset assignments)
  if (action === 'refresh_dropdown') {
    try { refreshMesaDropdown(); return respondGet({ success: true }, callback); }
    catch (err) { return respondGet({ error: err.message }, callback); }
  }

  // Admin actions (no code needed)
  if (action === 'generar_codigos') {
    try { generateCodesRemote(); return respondGet({ success: true }, callback); }
    catch (err) { return respondGet({ error: err.message }, callback); }
  }
  if (action === 'generar_mensajes') {
    try { generateWhatsAppMessagesRemote(); return respondGet({ success: true }, callback); }
    catch (err) { return respondGet({ error: err.message }, callback); }
  }
  if (action === 'setup_mesas') {
    try { setupMesasRemote(); return respondGet({ success: true }, callback); }
    catch (err) { return respondGet({ error: err.message }, callback); }
  }
  if (action === 'crear_resumen') {
    try { createResumenRemote(); return respondGet({ success: true }, callback); }
    catch (err) { return respondGet({ error: err.message }, callback); }
  }
  if (action === 'generar_mapa') {
    try { generateMapaMesasRemote(); return respondGet({ success: true }, callback); }
    catch (err) { return respondGet({ error: err.message }, callback); }
  }

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

  // Increment visit counter and stamp last-visit date (per row in the group)
  var nowVisit = new Date();
  visitRows.forEach(function(rowNum) {
    if (col['visitas'] !== undefined) {
      var currentVisits = Number(data[rowNum - 1][col['visitas']]) || 0;
      sheet.getRange(rowNum, col['visitas'] + 1).setValue(currentVisits + 1);
    }
    if (col['ultima_visita'] !== undefined) {
      sheet.getRange(rowNum, col['ultima_visita'] + 1).setValue(nowVisit);
    }
  });

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

  // Send notification email
  try {
    var names = responses.map(function(r) {
      var nombre = '';
      if (r.row && r.row > 1 && r.row <= data.length) {
        nombre = (data[r.row - 1][col['Nombre']] + ' ' + data[r.row - 1][col['Apellido']]).trim();
      }
      var estado = r.confirmado === true ? 'Si' : r.confirmado === false ? 'No' : r.confirmado === 'tal_vez' ? 'A\u00fan no sabe' : '-';
      var menu = r.menu || '-';
      return nombre + ' -> ' + estado + (r.confirmado === true ? ' (' + menu + ')' : '');
    }).join('\n');

    var grupo = data[responses[0].row - 1][col['Grupo']];
    var subject = 'RSVP Boda: ' + grupo;
    var body = 'El grupo "' + grupo + '" acaba de confirmar:\n\n' + names + '\n\nVer sheet: https://docs.google.com/spreadsheets/d/1rTGn5WOs_b3OFi1n3PVdrZZ8wA4Qm6ARYTQvcTfWgVw/edit';

    MailApp.sendEmail('mauricioaizaga@gmail.com,eylabriceno15@gmail.com', subject, body);
  } catch (e) {
    // Don't fail the save if email fails
  }

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
  generateWhatsAppMessagesRemote();
  SpreadsheetApp.getUi().alert('Mensajes generados. Revisa la pesta\u00f1a "Mensajes".');
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
  createResumenRemote();
  SpreadsheetApp.getUi().alert('Pesta\u00f1a Resumen actualizada.');
}

// Setup table management: creates "Mesas" config tab, dropdown validation, and conditional formatting
function setupMesas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  var col = {};
  headers.forEach(function(h, i) { col[h] = i; });

  // Check if "mesa" column exists
  if (col['mesa'] === undefined) {
    SpreadsheetApp.getUi().alert('No se encontr\u00f3 la columna "mesa" en Invitados. Cr\u00e9ala primero.');
    return;
  }

  var mesaColIndex = col['mesa'] + 1; // 1-based for getRange

  // Create or reset Mesas config tab
  var mesasSheet = ss.getSheetByName('Mesas');
  if (!mesasSheet) {
    mesasSheet = ss.insertSheet('Mesas');
  } else {
    mesasSheet.clearContents();
    mesasSheet.clearFormats();
  }

  // Default 8 table names
  var tableNames = [
    'Mesa 1', 'Mesa 2', 'Mesa 3', 'Mesa 4',
    'Mesa 5', 'Mesa 6', 'Mesa 7', 'Mesa 8'
  ];

  var mesaCol = String.fromCharCode(65 + col['mesa']); // Column letter for mesa
  var lastRow = sheet.getLastRow();

  mesasSheet.getRange(1, 1, 1, 4).setValues([['Mesa', 'Capacidad', 'Ocupados', 'Disponibles']]);
  mesasSheet.getRange(1, 1, 1, 4).setFontWeight('bold');

  for (var i = 0; i < tableNames.length; i++) {
    var row = i + 2;
    mesasSheet.getRange(row, 1).setValue(tableNames[i]);
    mesasSheet.getRange(row, 2).setValue(10);
    mesasSheet.getRange(row, 3).setFormula('=COUNTIF(Invitados!' + mesaCol + '2:' + mesaCol + lastRow + ', A' + row + ')');
    mesasSheet.getRange(row, 4).setFormula('=B' + row + '-C' + row);
  }

  // Totals row
  var totalRow = tableNames.length + 2;
  mesasSheet.getRange(totalRow, 1).setValue('TOTAL');
  mesasSheet.getRange(totalRow, 1).setFontWeight('bold');
  mesasSheet.getRange(totalRow, 2).setFormula('=SUM(B2:B' + (totalRow - 1) + ')');
  mesasSheet.getRange(totalRow, 3).setFormula('=SUM(C2:C' + (totalRow - 1) + ')');
  mesasSheet.getRange(totalRow, 4).setFormula('=SUM(D2:D' + (totalRow - 1) + ')');

  // Sin mesa row
  var sinMesaRow = totalRow + 1;
  mesasSheet.getRange(sinMesaRow, 1).setValue('Sin mesa (confirmados)');
  mesasSheet.getRange(sinMesaRow, 1).setFontWeight('bold');
  var confCol = String.fromCharCode(65 + col['confirmado']);
  mesasSheet.getRange(sinMesaRow, 3).setFormula(
    '=COUNTIFS(Invitados!' + confCol + '2:' + confCol + lastRow + ',"si",Invitados!' + mesaCol + '2:' + mesaCol + lastRow + ',"")'
  );

  // Conditional formatting on Mesas tab: red if disponibles < 0, yellow if = 0, green if > 0
  var rangeDisp = mesasSheet.getRange('D2:D' + (totalRow - 1));
  var rules = mesasSheet.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setBackground('#f4cccc')
    .setRanges([rangeDisp])
    .build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(0)
    .setBackground('#fff2cc')
    .setRanges([rangeDisp])
    .build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground('#d9ead3')
    .setRanges([rangeDisp])
    .build());
  mesasSheet.setConditionalFormatRules(rules);

  mesasSheet.setColumnWidth(1, 200);
  mesasSheet.setColumnWidth(2, 100);
  mesasSheet.setColumnWidth(3, 100);
  mesasSheet.setColumnWidth(4, 100);

  // Add dropdown validation to "mesa" column in Invitados
  var mesaNamesRange = mesasSheet.getRange('A2:A' + (tableNames.length + 1));
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(mesaNamesRange, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, mesaColIndex, lastRow - 1, 1).setDataValidation(rule);

  // Conditional formatting on mesa column in Invitados
  // Green if has a mesa assigned, no color if empty
  var mesaRange = sheet.getRange(2, mesaColIndex, lastRow - 1, 1);
  var invRules = sheet.getConditionalFormatRules();
  invRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('Mesa')
    .setBackground('#d9ead3')
    .setRanges([mesaRange])
    .build());
  sheet.setConditionalFormatRules(invRules);

  SpreadsheetApp.getUi().alert(
    'Mesas configuradas.\n\n' +
    '- Columna "mesa" con dropdown en Invitados\n' +
    '- Pesta\u00f1a "Mesas" con conteo en tiempo real\n' +
    '- Puedes renombrar las mesas en la pesta\u00f1a "Mesas" columna A\n\n' +
    'Usa "Generar mapa de mesas" para ver el mapa visual.'
  );
}

// Generate visual table map
function generateMapaMesasMenu() {
  generateMapaMesasRemote();
  SpreadsheetApp.getUi().alert('Mapa de mesas generado.');
}

// Remote versions (no getUi)
function setupMesasRemote() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  var col = {};
  headers.forEach(function(h, i) { col[h] = i; });

  if (col['mesa'] === undefined) throw new Error('Columna "mesa" no encontrada');

  var mesaColIndex = col['mesa'] + 1;
  var mesasSheet = ss.getSheetByName('Mesas');
  if (!mesasSheet) mesasSheet = ss.insertSheet('Mesas');
  else { mesasSheet.clearContents(); mesasSheet.clearFormats(); }

  var tableNames = ['Principal','Brice\u00f1o','C\u00e1rdenas','Chamos','iKono','Aizaga 1','Aizaga 2','Pais','UTP'];
  var mesaCol = String.fromCharCode(65 + col['mesa']);
  var lastRow = sheet.getLastRow();

  mesasSheet.getRange(1, 1, 1, 4).setValues([['Mesa', 'Capacidad', 'Ocupados', 'Disponibles']]);
  mesasSheet.getRange(1, 1, 1, 4).setFontWeight('bold');

  for (var i = 0; i < tableNames.length; i++) {
    var row = i + 2;
    mesasSheet.getRange(row, 1).setValue(tableNames[i]);
    mesasSheet.getRange(row, 2).setValue(10);
    mesasSheet.getRange(row, 3).setFormula('=COUNTIF(Invitados!' + mesaCol + '2:' + mesaCol + lastRow + ', A' + row + ')');
    mesasSheet.getRange(row, 4).setFormula('=B' + row + '-C' + row);
  }

  var totalRow = tableNames.length + 2;
  mesasSheet.getRange(totalRow, 1).setValue('TOTAL');
  mesasSheet.getRange(totalRow, 1).setFontWeight('bold');
  mesasSheet.getRange(totalRow, 2).setFormula('=SUM(B2:B' + (totalRow - 1) + ')');
  mesasSheet.getRange(totalRow, 3).setFormula('=SUM(C2:C' + (totalRow - 1) + ')');
  mesasSheet.getRange(totalRow, 4).setFormula('=SUM(D2:D' + (totalRow - 1) + ')');

  var sinMesaRow = totalRow + 1;
  mesasSheet.getRange(sinMesaRow, 1).setValue('Sin mesa (confirmados)');
  mesasSheet.getRange(sinMesaRow, 1).setFontWeight('bold');
  var confCol = String.fromCharCode(65 + col['confirmado']);
  mesasSheet.getRange(sinMesaRow, 3).setFormula(
    '=COUNTIFS(Invitados!' + confCol + '2:' + confCol + lastRow + ',"si",Invitados!' + mesaCol + '2:' + mesaCol + lastRow + ',"")'
  );

  var rangeDisp = mesasSheet.getRange('D2:D' + (totalRow - 1));
  var rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setBackground('#f4cccc').setRanges([rangeDisp]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(0).setBackground('#fff2cc').setRanges([rangeDisp]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground('#d9ead3').setRanges([rangeDisp]).build());
  mesasSheet.setConditionalFormatRules(rules);

  mesasSheet.setColumnWidth(1, 200);
  mesasSheet.setColumnWidth(2, 100);
  mesasSheet.setColumnWidth(3, 100);
  mesasSheet.setColumnWidth(4, 100);

  var mesaNamesRange = mesasSheet.getRange('A2:A' + (tableNames.length + 1));
  var rule = SpreadsheetApp.newDataValidation().requireValueInRange(mesaNamesRange, true).setAllowInvalid(false).build();
  sheet.getRange(2, mesaColIndex, lastRow - 1, 1).setDataValidation(rule);

  var mesaRange = sheet.getRange(2, mesaColIndex, lastRow - 1, 1);
  var invRules = sheet.getConditionalFormatRules();
  invRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('Mesa').setBackground('#d9ead3').setRanges([mesaRange]).build());
  sheet.setConditionalFormatRules(invRules);
}

function generateCodesRemote() {
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

  for (var i = 1; i < data.length; i++) {
    var code = String(data[i][codeCol]).trim();
    var group = String(data[i][groupCol]).trim();
    if (code && code !== '' && code !== 'undefined') {
      existingCodes[code] = true;
      groupCodes[group] = code;
    }
  }

  for (var i = 1; i < data.length; i++) {
    var code = String(data[i][codeCol]).trim();
    var group = String(data[i][groupCol]).trim();
    if (!code || code === '' || code === 'undefined') {
      if (groupCodes[group]) {
        sheet.getRange(i + 1, codeCol + 1).setValue(groupCodes[group]);
      } else {
        var newCode;
        do { newCode = randomCode(6); } while (existingCodes[newCode]);
        existingCodes[newCode] = true;
        groupCodes[group] = newCode;
        sheet.getRange(i + 1, codeCol + 1).setValue(newCode);
      }
    }
  }
}

function generateWhatsAppMessagesRemote() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var col = {};
  headers.forEach(function(h, i) { col[h] = i; });

  var baseUrl = 'https://maoaiz.github.io/wedding/invitation/?code=';
  var saveTheDateUrl = 'https://maoaiz.github.io/wedding/?code=';
  var groups = {};

  for (var i = 1; i < data.length; i++) {
    var group = String(data[i][col['Grupo']]).trim();
    var code = String(data[i][col['codigo']]).trim();
    var phone = String(data[i][col['Telefono']]).trim();
    var confirmado = String(data[i][col['confirmado']]).trim().toLowerCase();
    if (group && !groups[group]) {
      groups[group] = { code: code, phone: '', hasSiNo: false };
    }
    if (!groups[group].phone && phone && phone !== '' && phone !== 'undefined') {
      groups[group].phone = phone;
    }
    if (confirmado === 'si' || confirmado === 'no') {
      groups[group].hasSiNo = true;
    }
  }

  var msgSheet = ss.getSheetByName('Mensajes');
  if (!msgSheet) msgSheet = ss.insertSheet('Mensajes');
  else msgSheet.clearContents();

  msgSheet.getRange(1, 1, 1, 3).setValues([['Grupo', 'Invitaci\u00f3n', 'Save the date']]);
  msgSheet.getRange(1, 1, 1, 3).setFontWeight('bold');

  var ring = String.fromCodePoint(0x1F48D);
  var partying = String.fromCodePoint(0x1F973);
  var manDance = String.fromCodePoint(0x1F57A) + String.fromCodePoint(0x1F3FD);
  var womanDance = String.fromCodePoint(0x1F483) + String.fromCodePoint(0x1F3FC);
  var heart = String.fromCodePoint(0x2764) + String.fromCodePoint(0xFE0F);
  var down = String.fromCodePoint(0x1F447);
  var smiley = String.fromCodePoint(0x263A) + String.fromCodePoint(0xFE0F);

  var entries = Object.entries(groups);
  entries.forEach(function(entry, i) {
    var group = entry[0];
    var info = entry[1];
    var url = baseUrl + info.code;
    var message = 'Hola ' + group + '! ' + heart + '\n\n' +
      '\u00a1Ya tenemos todos los detalles!' + ring + partying + manDance + womanDance + '\n\n' +
      'Abre aqu\u00ed para verlos ' + down + '\n' + url;
    if (!info.hasSiNo) {
      message += '\n\nEn caso de no obtener respuesta entenderemos que no podr\u00e1s acompa\u00f1arnos presencialmente, no te preocupes por eso, no es ning\u00fan problema ' + smiley;
    }
    var encodedMsg = encodeURIComponent(message);
    var cleanPhone = info.phone ? info.phone.replace(/[^0-9]/g, '') : '';
    var waLink = cleanPhone ? 'https://wa.me/' + cleanPhone + '?text=' + encodedMsg : '';

    var row = i + 2;
    if (waLink) {
      msgSheet.getRange(row, 1).setFormula('=HYPERLINK("' + waLink.replace(/"/g, '""') + '", "' + group.replace(/"/g, '""') + '")');
    } else {
      msgSheet.getRange(row, 1).setValue(group);
      msgSheet.getRange(row, 1).setFontColor('#cccccc');
    }
    msgSheet.getRange(row, 2).setFormula('=HYPERLINK("' + url + '", "Ver invitaci\u00f3n")');
    var stdUrl = saveTheDateUrl + info.code;
    msgSheet.getRange(row, 3).setFormula('=HYPERLINK("' + stdUrl + '", "Ver save the date")');
  });

  msgSheet.setColumnWidth(1, 280);
  msgSheet.setColumnWidth(2, 140);
  msgSheet.setColumnWidth(3, 150);
}

function createResumenRemote() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  var colLetter = {};
  headers.forEach(function(h, i) {
    colLetter[h] = String.fromCharCode(65 + i);
    if (i >= 26) colLetter[h] = 'A' + String.fromCharCode(65 + i - 26);
  });

  var conf = colLetter['confirmado'];
  var menu = colLetter['menu'];
  var nino = colLetter['es_nino'];
  var nombre = colLetter['Nombre'];
  var etiqueta = colLetter['Etiqueta'];
  var visitas = colLetter['visitas'];

  var resSheet = ss.getSheetByName('Resumen');
  if (!resSheet) resSheet = ss.insertSheet('Resumen');
  else { resSheet.clearContents(); resSheet.clearFormats(); }

  // Helper to build range strings (e.g. "Invitados!H2:H")
  var range = function(letter) { return 'Invitados!' + letter + '2:' + letter; };
  var total = 'COUNTA(' + range(nombre) + ')';
  var nSi = 'COUNTIF(' + range(conf) + ',"si")';
  var nNo = 'COUNTIF(' + range(conf) + ',"no")';
  var nTalVez = 'COUNTIF(' + range(conf) + ',"tal_vez")';
  var nResp = '(' + nSi + '+' + nNo + '+' + nTalVez + ')';
  var nPend = '(' + total + '-' + nResp + ')';
  var nNinos = 'COUNTIF(' + range(nino) + ',"Si")';
  var nNinosSi = 'COUNTIFS(' + range(nino) + ',"Si",' + range(conf) + ',"si")';
  var nVisitados = 'COUNTIF(' + range(visitas) + ',">0")';

  var tagRow = function(label, pattern, base) {
    var formula = 'COUNTIF(' + range(etiqueta) + ',"*' + pattern + '*")';
    return [label, '=' + formula, '=IFERROR(' + formula + '/' + base + ',0)'];
  };

  var tagConfirmedRow = function(label, pattern) {
    var match = 'COUNTIFS(' + range(etiqueta) + ',"*' + pattern + '*",' + range(conf) + ',"si")';
    var totalTag = 'COUNTIF(' + range(etiqueta) + ',"*' + pattern + '*")';
    return [label, '=' + match, '=IFERROR(' + match + '/' + totalTag + ',0)'];
  };

  var data = [
    ['RESUMEN', 'Cantidad', '%'],
    ['', '', ''],
    ['Total invitados', '=' + total, ''],
    ['Confirmados (s\u00ed)', '=' + nSi, '=' + nSi + '/' + total],
    ['No asisten', '=' + nNo, '=' + nNo + '/' + total],
    ['A\u00fan no saben', '=' + nTalVez, '=' + nTalVez + '/' + total],
    ['Sin responder', '=' + nPend, '=' + nPend + '/' + total],
    ['', '', ''],
    ['TASA DE RESPUESTA', '', ''],
    ['Han respondido', '=' + nResp, '=' + nResp + '/' + total],
    ['Pendientes', '=' + nPend, '=' + nPend + '/' + total],
    ['', '', ''],
    ['ASISTENCIA ESPERADA', '', ''],
    ['M\u00ednima (solo confirmados)', '=' + nSi, '=' + nSi + '/' + total],
    ['Probable (confirmados + 50% tal vez)', '=' + nSi + '+ROUND(' + nTalVez + '*0.5)', '=(' + nSi + '+ROUND(' + nTalVez + '*0.5))/' + total],
    ['M\u00e1xima (confirmados + tal vez)', '=' + nSi + '+' + nTalVez, '=(' + nSi + '+' + nTalVez + ')/' + total],
    ['', '', ''],
    ['MEN\u00daS (de confirmados)', '', ''],
    ['Normal', '=COUNTIF(' + range(menu) + ',"normal")', '=IFERROR(COUNTIF(' + range(menu) + ',"normal")/' + nSi + ',0)'],
    ['Vegetariano', '=COUNTIF(' + range(menu) + ',"vegetariano")', '=IFERROR(COUNTIF(' + range(menu) + ',"vegetariano")/' + nSi + ',0)'],
    ['Infantil', '=COUNTIF(' + range(menu) + ',"infantil")', '=IFERROR(COUNTIF(' + range(menu) + ',"infantil")/' + nSi + ',0)'],
    ['Sin definir', '=' + nSi + '-COUNTIF(' + range(menu) + ',"normal")-COUNTIF(' + range(menu) + ',"vegetariano")-COUNTIF(' + range(menu) + ',"infantil")', '=IFERROR((' + nSi + '-COUNTIF(' + range(menu) + ',"normal")-COUNTIF(' + range(menu) + ',"vegetariano")-COUNTIF(' + range(menu) + ',"infantil"))/' + nSi + ',0)'],
    ['', '', ''],
    ['DEMOGRAF\u00cdA', '', ''],
    ['Ni\u00f1os invitados', '=' + nNinos, '=' + nNinos + '/' + total],
    ['Adultos invitados', '=' + total + '-' + nNinos, '=(' + total + '-' + nNinos + ')/' + total],
    ['Ni\u00f1os confirmados', '=' + nNinosSi, '=IFERROR(' + nNinosSi + '/' + nNinos + ',0)'],
    ['Adultos confirmados', '=' + nSi + '-' + nNinosSi, '=IFERROR((' + nSi + '-' + nNinosSi + ')/(' + total + '-' + nNinos + '),0)'],
    ['', '', ''],
    ['POR ETIQUETA (total / % confirmaci\u00f3n)', '', ''],
    tagConfirmedRow('Brice\u00f1os', 'Brice\u00f1os'),
    tagConfirmedRow('Amigos Eyla', 'Amigos Eyla'),
    tagConfirmedRow('Amigos Mauricio', 'Amigos Mauricio'),
    tagConfirmedRow('Pekin', 'Pekin'),
    tagConfirmedRow('Astros', 'Astros'),
    tagConfirmedRow('Aizagas', 'Aizagas'),
    tagConfirmedRow('iKono', 'iKono'),
    tagConfirmedRow('UTP', 'UTP'),
    tagConfirmedRow('Pasto', 'Pasto'),
    ['', '', ''],
    ['ENGAGEMENT', '', ''],
    ['Han abierto el enlace', '=' + nVisitados, '=' + nVisitados + '/' + total],
    ['Abrieron pero no respondieron', '=COUNTIFS(' + range(visitas) + ',">0",' + range(conf) + ',"")', '=IFERROR(COUNTIFS(' + range(visitas) + ',">0",' + range(conf) + ',"")/' + nVisitados + ',0)'],
    ['Total visitas', '=SUM(' + range(visitas) + ')', ''],
    ['Promedio visitas / invitado que abri\u00f3', '=IFERROR(SUM(' + range(visitas) + ')/' + nVisitados + ',0)', ''],
    ['', '', ''],
    ['FECHAS', '', ''],
    ['D\u00edas para la boda', '=DATE(2026,6,6)-TODAY()', ''],
    ['D\u00edas para deadline RSVP (6 may)', '=DATE(2026,5,6)-TODAY()', '']
  ];

  resSheet.getRange(1, 1, data.length, 3).setValues(data);

  // Bold section headers
  data.forEach(function(row, i) {
    if (row[0] && (row[0] === row[0].toUpperCase() || row[0].indexOf('POR ETIQUETA') === 0 || row[0].indexOf('MEN') === 0) && row[1] === '' && row[2] === '') {
      resSheet.getRange(i + 1, 1, 1, 3).setFontWeight('bold').setBackground('#f0ede5');
    }
  });

  // Header row formatting
  resSheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#5c6b4f').setFontColor('#ffffff');

  // Format percentage column
  resSheet.getRange(1, 3, data.length, 1).setNumberFormat('0%');

  // Column widths
  resSheet.setColumnWidth(1, 320);
  resSheet.setColumnWidth(2, 100);
  resSheet.setColumnWidth(3, 90);
  resSheet.setFrozenRows(1);
}

function refreshMesaDropdownMenu() {
  refreshMesaDropdown();
  SpreadsheetApp.getUi().alert('Dropdown actualizado.');
}

function refreshMesaDropdown() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var mesasSheet = ss.getSheetByName('Mesas');
  if (!mesasSheet) throw new Error('No existe la pesta\u00f1a Mesas');

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var col = {};
  headers.forEach(function(h, i) { col[h] = i; });

  var mesaColIndex = col['mesa'] + 1;
  var lastRow = sheet.getLastRow();

  // Find last mesa row (skip TOTAL and Sin mesa rows)
  var mesasData = mesasSheet.getDataRange().getValues();
  var lastMesaRow = 1;
  for (var i = 1; i < mesasData.length; i++) {
    var name = String(mesasData[i][0]).trim();
    if (name && name !== 'TOTAL' && name !== 'Sin mesa (confirmados)') {
      lastMesaRow = i + 1;
    }
  }

  var mesaNamesRange = mesasSheet.getRange('A2:A' + lastMesaRow);
  var rule = SpreadsheetApp.newDataValidation().requireValueInRange(mesaNamesRange, true).setAllowInvalid(false).build();
  sheet.getRange(2, mesaColIndex, lastRow - 1, 1).setDataValidation(rule);
}

function confLabel(conf) {
  if (conf === 'si') return { text: 'Confirmado', color: '#d9ead3' };
  if (conf === 'no') return { text: 'No asiste', color: '#f4cccc' };
  if (conf === 'tal_vez') return { text: 'No sabe', color: '#fff2cc' };
  return { text: 'Pendiente', color: '#eeeeee' };
}

function generateMapaMesasRemote() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var mesasSheet = ss.getSheetByName('Mesas');
  if (!mesasSheet) throw new Error('Primero ejecuta setupMesas');

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var col = {};
  headers.forEach(function(h, i) { col[h] = i; });

  var data = sheet.getDataRange().getValues();
  var mesasData = mesasSheet.getDataRange().getValues();
  var tables = {};
  for (var i = 1; i < mesasData.length; i++) {
    var name = String(mesasData[i][0]).trim();
    if (name && name !== 'TOTAL' && name !== 'Sin mesa (confirmados)') {
      tables[name] = { capacity: mesasData[i][1] || 10, members: [] };
    }
  }

  for (var i = 1; i < data.length; i++) {
    var mesa = String(data[i][col['mesa']]).trim();
    var nombre = (data[i][col['Nombre']] + ' ' + data[i][col['Apellido']]).trim();
    var conf = String(data[i][col['confirmado']]).trim().toLowerCase();
    if (mesa && tables[mesa]) {
      tables[mesa].members.push({ nombre: nombre, confirmado: conf });
    }
  }

  var mapaSheet = ss.getSheetByName('Mapa de Mesas');
  if (mapaSheet) ss.deleteSheet(mapaSheet);
  mapaSheet = ss.insertSheet('Mapa de Mesas');

  var tableNames = Object.keys(tables);
  var colors = ['#e8d5c4','#d5c4b3','#c4b3a2','#dcc8b8','#e8d0c0','#d0b8a8','#c8b0a0','#e0c8b8'];
  var border = SpreadsheetApp.BorderStyle.SOLID;
  var borderColor = '#c0b0a0';

  // Title
  var r = 1;
  mapaSheet.getRange(r, 1).setValue('MAPA DE MESAS - Boda Eyla & Mauricio');
  mapaSheet.getRange(r, 1).setFontWeight('bold').setFontSize(14);
  r += 2;

  // Draw a single table at position
  function drawTable(name, table, row, colStart, color) {
    mapaSheet.getRange(row, colStart, 1, 2).setValues([[name + ' (' + table.members.length + '/' + table.capacity + ')', '']]);
    mapaSheet.getRange(row, colStart, 1, 2).setBackground(color).setFontWeight('bold').setHorizontalAlignment('center');
    mapaSheet.getRange(row, colStart, 1, 2).setBorder(true, true, true, true, false, false, borderColor, border);
    row++;

    for (var m = 0; m < table.capacity; m++) {
      if (m < table.members.length) {
        var p = table.members[m];
        var e = confLabel(p.confirmado);
        mapaSheet.getRange(row, colStart).setValue(p.nombre).setBackground('#f5f0eb');
        mapaSheet.getRange(row, colStart + 1).setValue(e.text).setBackground(e.color).setFontSize(9).setHorizontalAlignment('center');
      } else {
        mapaSheet.getRange(row, colStart).setValue('- vac\u00edo -').setFontColor('#cccccc').setHorizontalAlignment('center');
        mapaSheet.getRange(row, colStart + 1).setBackground('#fafafa');
      }
      mapaSheet.getRange(row, colStart, 1, 2).setBorder(false, true, false, true, false, false, borderColor, border);
      row++;
    }
    mapaSheet.getRange(row - 1, colStart, 1, 2).setBorder(false, true, true, true, false, false, borderColor, border);
    return row;
  }

  // First table (Principal) alone
  var firstEnd = drawTable(tableNames[0], tables[tableNames[0]], r, 1, colors[0]);
  r = firstEnd + 2;

  // Rest in pairs
  for (var t = 1; t < tableNames.length; t += 2) {
    var startRow = r;

    // LEFT TABLE
    var leftName = tableNames[t];
    var leftTable = tables[leftName];
    var leftEnd = drawTable(leftName, leftTable, r, 1, colors[t % colors.length]);

    // RIGHT TABLE
    if (t + 1 < tableNames.length) {
      var rightName = tableNames[t + 1];
      var rightTable = tables[rightName];
      drawTable(rightName, rightTable, startRow, 4, colors[(t+1) % colors.length]);
    }

    r = leftEnd + 2;
  }

  // Sin mesa
  var sinMesa = [];
  for (var i = 1; i < data.length; i++) {
    var mesa = String(data[i][col['mesa']]).trim();
    var nombre = (data[i][col['Nombre']] + ' ' + data[i][col['Apellido']]).trim();
    var conf = String(data[i][col['confirmado']]).trim().toLowerCase();
    if (!mesa && nombre) {
      sinMesa.push({ nombre: nombre, confirmado: conf });
    }
  }

  if (sinMesa.length > 0) {
    r += 1;
    mapaSheet.getRange(r, 1, 1, 5).setValues([['SIN MESA ASIGNADA (' + sinMesa.length + ')', '', '', '', '']]);
    mapaSheet.getRange(r, 1, 1, 5).setFontWeight('bold').setFontSize(12).setBackground('#f4cccc');
    r++;

    mapaSheet.getRange(r, 1).setValue('Nombre').setFontWeight('bold');
    mapaSheet.getRange(r, 2).setValue('Estado').setFontWeight('bold');
    r++;

    sinMesa.forEach(function(p) {
      var estado = p.confirmado === 'si' ? 'Confirmado' : p.confirmado === 'no' ? 'No asiste' : p.confirmado === 'tal_vez' ? 'A\u00fan no sabe' : 'Sin responder';
      mapaSheet.getRange(r, 1).setValue(p.nombre);
      mapaSheet.getRange(r, 2).setValue(estado);
      if (p.confirmado === 'si') mapaSheet.getRange(r, 2).setBackground('#d9ead3');
      else if (p.confirmado === 'no') mapaSheet.getRange(r, 2).setBackground('#f4cccc');
      r++;
    });
  }

  // Column widths
  mapaSheet.setColumnWidth(1, 180);
  mapaSheet.setColumnWidth(2, 90);
  mapaSheet.setColumnWidth(3, 30);
  mapaSheet.setColumnWidth(4, 180);
  mapaSheet.setColumnWidth(5, 90);
}

// Menu
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('RSVP')
    .addItem('Generar codigos', 'generateCodes')
    .addItem('Generar mensajes WhatsApp', 'generateWhatsAppMessages')
    .addItem('Crear Resumen', 'createResumen')
    .addSeparator()
    .addItem('Configurar mesas (resetea todo)', 'setupMesas')
    .addItem('Actualizar dropdown de mesas', 'refreshMesaDropdownMenu')
    .addItem('Generar mapa de mesas', 'generateMapaMesasMenu')
    .addToUi();
}
