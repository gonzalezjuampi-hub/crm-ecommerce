/**
 * CRM Ecommerce — Backend en Google Apps Script
 * Guarda y sirve fichas de clientes desde una hoja de Google Sheets llamada "Clientes".
 *
 * INSTALACIÓN:
 * 1. Crear un Google Sheet nuevo (vacío).
 * 2. Extensiones → Apps Script. Borrar el contenido de Code.gs y pegar este archivo.
 * 3. Guardar. Implementar → Nueva implementación → Tipo: Aplicación web.
 *    - Ejecutar como: Yo
 *    - Quién tiene acceso: Cualquier usuario (o "Cualquier usuario de AWS GROUP" si usan Workspace)
 * 4. Copiar la URL de la app web (termina en /exec). Esa es la APPS_SCRIPT_URL que va en el proxy de Netlify.
 * 5. (Opcional) Para las alarmas por email: Triggers (reloj) → agregar un trigger que corra
 *    checkAlarmsAndNotify() una vez por día.
 */

const SHEET_NAME = 'Clientes';
const HEADERS = [
  'ID', 'NumeroCliente', 'NombreCliente', 'Nacionalidad', 'FechaPrimerContacto',
  'DetallePedido', 'BodegasPreferencia', 'VinosPreferencia', 'RangoPresupuesto',
  'FrecuenciaContacto', 'FechaUltimoContacto', 'FechaProximoContacto', 'FechaCompra',
  'DuracionNegociacionDias', 'PuntosDebiles', 'PuntosInteres', 'Estado', 'FechaCreacion'
];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function doGet(e) {
  const action = (e.parameter && e.parameter.action) || 'list';
  if (action === 'list') return json_(listClients_());
  return json_({ error: 'Acción no reconocida' });
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ error: 'JSON inválido' });
  }
  const action = body.action;
  let result;
  switch (action) {
    case 'create': result = createClient_(body.data); break;
    case 'update': result = updateClient_(body.data); break;
    case 'delete': result = deleteClient_(body.id); break;
    case 'logContact': result = logContact_(body.id, body.frecuencia); break;
    default: result = { error: 'Acción no reconocida' };
  }
  return json_(result);
}

function listClients_() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  return values.slice(1)
    .filter(row => row[0])
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
}

function findRowById_(sheet, id) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) return i + 1; // fila 1-indexada
  }
  return -1;
}

function createClient_(data) {
  const sheet = getSheet_();
  const id = Utilities.getUuid();
  const now = new Date();
  const frecuencia = data.FrecuenciaContacto || 30;
  const next = addDays_(now, frecuencia);
  const row = HEADERS.map(h => {
    if (h === 'ID') return id;
    if (h === 'FechaCreacion') return now;
    if (h === 'FechaProximoContacto') return next;
    if (h === 'Estado') return data.Estado || 'Activo';
    return data[h] !== undefined ? data[h] : '';
  });
  sheet.appendRow(row);
  return { success: true, id: id };
}

function updateClient_(data) {
  const sheet = getSheet_();
  const rowIndex = findRowById_(sheet, data.ID);
  if (rowIndex === -1) return { error: 'Cliente no encontrado' };
  HEADERS.forEach((h, i) => {
    if (h !== 'ID' && h !== 'FechaCreacion' && data[h] !== undefined) {
      sheet.getRange(rowIndex, i + 1).setValue(data[h]);
    }
  });
  return { success: true };
}

function deleteClient_(id) {
  const sheet = getSheet_();
  const rowIndex = findRowById_(sheet, id);
  if (rowIndex === -1) return { error: 'Cliente no encontrado' };
  sheet.deleteRow(rowIndex);
  return { success: true };
}

function logContact_(id, frecuencia) {
  const sheet = getSheet_();
  const rowIndex = findRowById_(sheet, id);
  if (rowIndex === -1) return { error: 'Cliente no encontrado' };
  const now = new Date();
  const next = addDays_(now, frecuencia);
  sheet.getRange(rowIndex, HEADERS.indexOf('FechaUltimoContacto') + 1).setValue(now);
  sheet.getRange(rowIndex, HEADERS.indexOf('FechaProximoContacto') + 1).setValue(next);
  sheet.getRange(rowIndex, HEADERS.indexOf('FrecuenciaContacto') + 1).setValue(frecuencia);
  return { success: true, fechaProximoContacto: next.toISOString() };
}

function addDays_(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + parseInt(days, 10));
  return d;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Trigger diario opcional: manda un mail con los clientes que vencen hoy o están vencidos.
 * Configurar en Triggers → Agregar trigger → checkAlarmsAndNotify → basado en tiempo → diario.
 */
function checkAlarmsAndNotify() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idxNombre = headers.indexOf('NombreCliente');
  const idxNumero = headers.indexOf('NumeroCliente');
  const idxProximo = headers.indexOf('FechaProximoContacto');
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const due = [];
  for (let i = 1; i < values.length; i++) {
    const fecha = values[i][idxProximo];
    if (fecha && new Date(fecha) <= today) {
      due.push('#' + values[i][idxNumero] + ' — ' + values[i][idxNombre]);
    }
  }
  if (due.length > 0) {
    MailApp.sendEmail(
      Session.getActiveUser().getEmail(),
      'CRM Ecommerce: clientes para contactar hoy (' + due.length + ')',
      due.join('\n')
    );
  }
}
