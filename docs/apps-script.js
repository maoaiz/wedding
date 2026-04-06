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
    ['Sin responder', '=COUNTA(Invitados!' + nombre + '2:' + nombre + ')-COUNTIF(Invitados!' + conf + '2:' + conf + ',"si")-COUNTIF(Invitados!' + conf + '2:' + conf + ',"no")-COUNTIF(Invitados!' + conf + '2:' + conf + ',"tal_vez")'],
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

  var msgSheet = ss.getSheetByName('Mensajes');
  if (!msgSheet) msgSheet = ss.insertSheet('Mensajes');
  else msgSheet.clearContents();

  msgSheet.getRange(1, 1, 1, 2).setValues([['Grupo', 'Invitaci\u00f3n']]);
  msgSheet.getRange(1, 1, 1, 2).setFontWeight('bold');

  var ring = String.fromCodePoint(0x1F48D);
  var bride = String.fromCodePoint(0x1F470);
  var groom = String.fromCodePoint(0x1F935);
  var party = String.fromCodePoint(0x1F389);
  var heart = String.fromCodePoint(0x2764) + String.fromCodePoint(0xFE0F);
  var down = String.fromCodePoint(0x1F447);
  var pray = String.fromCodePoint(0x1F64F);

  var entries = Object.entries(groups);
  entries.forEach(function(entry, i) {
    var group = entry[0];
    var info = entry[1];
    var url = baseUrl + info.code;
    var message = 'Hola ' + group + '! ' + heart + '\n\n' +
      'Tenemos una invitaci\u00f3n muy especial para ti. ' + ring + '\n\n' +
      'Abre aqu\u00ed para ver todos los detalles ' + down + '\n' + url;
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
  });

  msgSheet.setColumnWidth(1, 280);
  msgSheet.setColumnWidth(2, 120);
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
  else resSheet.clearContents();

  var data = [
    ['RESUMEN GENERAL', ''],
    ['', ''],
    ['Total invitados', '=COUNTA(Invitados!' + nombre + '2:' + nombre + ')'],
    ['Confirmados (si)', '=COUNTIF(Invitados!' + conf + '2:' + conf + ', "si")'],
    ['No asisten', '=COUNTIF(Invitados!' + conf + '2:' + conf + ', "no")'],
    ['A\u00fan no saben', '=COUNTIF(Invitados!' + conf + '2:' + conf + ', "tal_vez")'],
    ['Sin responder', '=COUNTA(Invitados!' + nombre + '2:' + nombre + ')-COUNTIF(Invitados!' + conf + '2:' + conf + ',"si")-COUNTIF(Invitados!' + conf + '2:' + conf + ',"no")-COUNTIF(Invitados!' + conf + '2:' + conf + ',"tal_vez")'],
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
    ['Han abierto el enlace', '=COUNTIF(Invitados!' + visitas + '2:' + visitas + ', ">"&0)'],
    ['Total visitas', '=SUM(Invitados!' + visitas + '2:' + visitas + ')'],
  ];

  resSheet.getRange(1, 1, data.length, 2).setValues(data);
  var boldRows = [1, 9, 14, 17, 28];
  boldRows.forEach(function(r) { resSheet.getRange(r, 1, 1, 2).setFontWeight('bold'); });
  resSheet.setColumnWidth(1, 280);
  resSheet.setColumnWidth(2, 100);
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

  // Build all data as a 2D array first, then write once (no merge needed)
  var tableNames = Object.keys(tables);
  var rows = [];
  var formats = []; // {row, col, bg, bold, fontSize, fontColor, align}

  // Title
  rows.push(['MAPA DE MESAS - Boda Eyla & Mauricio', '', '', '', '', '']);
  formats.push({r:0, c:0, bold:true, fontSize:14, align:'left'});
  rows.push(['', '', '', '', '', '']);

  var colors = ['#e8d5c4','#d5c4b3','#c4b3a2','#dcc8b8','#e8d0c0','#d0b8a8','#c8b0a0','#e0c8b8'];

  for (var t = 0; t < tableNames.length; t += 2) {
    var leftName = tableNames[t];
    var leftTable = tables[leftName];
    var rightName = (t + 1 < tableNames.length) ? tableNames[t + 1] : null;
    var rightTable = rightName ? tables[rightName] : null;
    var maxCap = Math.max(leftTable.capacity, rightTable ? rightTable.capacity : 0);

    // Header row
    var headerRow = rows.length;
    var leftHeader = leftName + ' (' + leftTable.members.length + '/' + leftTable.capacity + ')';
    var rightHeader = rightTable ? rightName + ' (' + rightTable.members.length + '/' + rightTable.capacity + ')' : '';
    rows.push([leftHeader, '', '', rightHeader, '', '']);
    formats.push({r:headerRow, c:0, bold:true, bg:colors[t % colors.length]});
    formats.push({r:headerRow, c:1, bg:colors[t % colors.length]});
    formats.push({r:headerRow, c:2, bg:colors[t % colors.length]});
    if (rightTable) {
      formats.push({r:headerRow, c:3, bold:true, bg:colors[(t+1) % colors.length]});
      formats.push({r:headerRow, c:4, bg:colors[(t+1) % colors.length]});
      formats.push({r:headerRow, c:5, bg:colors[(t+1) % colors.length]});
    }

    // Member rows
    for (var m = 0; m < maxCap; m++) {
      var rowData = ['', '', '', '', '', ''];
      var rowIdx = rows.length;

      if (m < leftTable.members.length) {
        var lp = leftTable.members[m];
        var le = confLabel(lp.confirmado);
        rowData[0] = lp.nombre;
        rowData[1] = le.text;
        formats.push({r:rowIdx, c:0, bg:'#f5f0eb'});
        formats.push({r:rowIdx, c:1, bg:le.color, fontSize:9, align:'center'});
      } else if (m < leftTable.capacity) {
        rowData[0] = '- vac\u00edo -';
        formats.push({r:rowIdx, c:0, fontColor:'#cccccc', align:'center'});
      }

      if (rightTable) {
        if (m < rightTable.members.length) {
          var rp = rightTable.members[m];
          var re = confLabel(rp.confirmado);
          rowData[3] = rp.nombre;
          rowData[4] = re.text;
          formats.push({r:rowIdx, c:3, bg:'#f5f0eb'});
          formats.push({r:rowIdx, c:4, bg:re.color, fontSize:9, align:'center'});
        } else if (m < rightTable.capacity) {
          rowData[3] = '- vac\u00edo -';
          formats.push({r:rowIdx, c:3, fontColor:'#cccccc', align:'center'});
        }
      }

      rows.push(rowData);
    }

    // Spacing
    rows.push(['', '', '', '', '', '']);
    rows.push(['', '', '', '', '', '']);
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
    var sinRow = rows.length;
    rows.push(['SIN MESA ASIGNADA (' + sinMesa.length + ')', '', '', '', '', '']);
    formats.push({r:sinRow, c:0, bold:true, fontSize:12, bg:'#f4cccc'});
    formats.push({r:sinRow, c:1, bg:'#f4cccc'});
    formats.push({r:sinRow, c:2, bg:'#f4cccc'});
    formats.push({r:sinRow, c:3, bg:'#f4cccc'});
    formats.push({r:sinRow, c:4, bg:'#f4cccc'});
    formats.push({r:sinRow, c:5, bg:'#f4cccc'});

    var hdrRow = rows.length;
    rows.push(['Nombre', '', '', 'Estado', '', '']);
    formats.push({r:hdrRow, c:0, bold:true});
    formats.push({r:hdrRow, c:3, bold:true});

    sinMesa.forEach(function(p) {
      var idx = rows.length;
      var estado = p.confirmado === 'si' ? 'Confirmado' : p.confirmado === 'no' ? 'No asiste' : p.confirmado === 'tal_vez' ? 'A\u00fan no sabe' : 'Sin responder';
      var bg = p.confirmado === 'si' ? '#d9ead3' : p.confirmado === 'no' ? '#f4cccc' : '';
      rows.push([p.nombre, '', '', estado, '', '']);
      if (bg) formats.push({r:idx, c:3, bg:bg});
    });
  }

  // Write all data at once
  if (rows.length > 0) {
    mapaSheet.getRange(1, 1, rows.length, 6).setValues(rows);
  }

  // Apply formats
  formats.forEach(function(f) {
    var cell = mapaSheet.getRange(f.r + 1, f.c + 1);
    if (f.bold) cell.setFontWeight('bold');
    if (f.fontSize) cell.setFontSize(f.fontSize);
    if (f.bg) cell.setBackground(f.bg);
    if (f.fontColor) cell.setFontColor(f.fontColor);
    if (f.align) cell.setHorizontalAlignment(f.align);
  });

  // Column widths
  mapaSheet.setColumnWidth(1, 180);
  mapaSheet.setColumnWidth(2, 80);
  mapaSheet.setColumnWidth(3, 30);
  mapaSheet.setColumnWidth(4, 180);
  mapaSheet.setColumnWidth(5, 80);
  mapaSheet.setColumnWidth(6, 30);
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
