/* ════════════════════════════════════════════════════════════════
   SALA COMPARTIDA — Auditoría de Ambientes 2026-II
   Servidor de la página direccionac.github.io/aulaschiolecca/

   Qué hace: guarda en esta hoja de cálculo las propuestas de cambio
   y de eventos que la gente registra desde la web, y devuelve la
   lista para que todos vean lo mismo.

   No toques nada de aquí abajo salvo que sepas lo que haces. Las
   claves NO van en este archivo: se ponen en Configuración del
   proyecto → Propiedades del script.
   ════════════════════════════════════════════════════════════════ */

var HOJA = 'CAMBIOS';
/* 'json' se queda en la columna 7 a propósito: las filas que ya existen
   lo tienen ahí. Lo nuevo (alcance, desde, hasta) se agrega detrás, así
   nada de lo registrado hasta hoy se corre de sitio. */
var CABECERAS = ['cid', 'ts', 'estado', 'por', 'accion', 'resumen', 'json',
                 'alcance', 'desde', 'hasta'];

/* ---------- utilidades ---------- */

function hoja_() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var h = libro.getSheetByName(HOJA);
  if (!h) {
    h = libro.insertSheet(HOJA);
    h.getRange(1, 1, 1, CABECERAS.length).setValues([CABECERAS]).setFontWeight('bold');
    h.setFrozenRows(1);
    return h;
  }
  /* Hoja que viene de la versión anterior: se le agregan las cabeceras
     que le faltan sin tocar ni una fila de datos. */
  var ancho = Math.max(h.getLastColumn(), 1);
  var actuales = h.getRange(1, 1, 1, ancho).getValues()[0];
  var falta = false;
  for (var i = 0; i < CABECERAS.length; i++) {
    if (String(actuales[i] || '') !== CABECERAS[i]) falta = true;
  }
  if (falta) {
    if (h.getMaxColumns() < CABECERAS.length) {
      h.insertColumnsAfter(h.getMaxColumns(), CABECERAS.length - h.getMaxColumns());
    }
    h.getRange(1, 1, 1, CABECERAS.length).setValues([CABECERAS]).setFontWeight('bold');
    h.setFrozenRows(1);
  }
  /* Las fechas deben quedarse como texto: si Sheets las convierte a fecha,
     '2026-09-07' vuelve como un objeto Date y el portal no lo reconoce. */
  var cD = CABECERAS.indexOf('desde') + 1, cH = CABECERAS.indexOf('hasta') + 1;
  h.getRange(2, cD, Math.max(h.getMaxRows() - 1, 1), 1).setNumberFormat('@');
  h.getRange(2, cH, Math.max(h.getMaxRows() - 1, 1), 1).setNumberFormat('@');
  return h;
}

/* Normaliza una fecha a 'YYYY-MM-DD'. Acepta el texto que manda el portal
   y también un Date, por si alguien la escribió a mano en la hoja. */
function fechaTexto_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(v).trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
}

function responder_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function leerCambios_() {
  var h = hoja_();
  if (h.getLastRow() < 2) return [];
  var col = CABECERAS.indexOf('json') + 1;
  var celdas = h.getRange(2, col, h.getLastRow() - 1, 1).getValues();
  var salida = [];
  for (var i = 0; i < celdas.length; i++) {
    var txt = celdas[i][0];
    if (!txt) continue;
    try { salida.push(JSON.parse(txt)); } catch (e) { /* fila corrupta: se ignora */ }
  }
  return salida;
}

/* Valida quién escribe. La clave de admin sirve para todo; la de
   editor solo para proponer. Se guardan en Propiedades del script. */
function rol_(clave) {
  var p = PropertiesService.getScriptProperties();
  var admin  = p.getProperty('CLAVE_ADMIN');
  var editor = p.getProperty('CLAVE_EDITOR');
  if (!clave) return null;
  if (admin && clave === admin) return 'admin';
  if (editor && clave === editor) return 'editor';
  return null;
}

/* ---------- lectura: la página consulta con GET ---------- */

function doGet() {
  try {
    return responder_({ ok: true, cambios: leerCambios_() });
  } catch (err) {
    return responder_({ ok: false, error: String(err) });
  }
}

/* ---------- escritura: proponer y resolver ---------- */

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return responder_({ ok: false, error: 'La hoja está ocupada, vuelve a intentar en unos segundos.' });
  }

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return responder_({ ok: false, error: 'No llegó ningún dato.' });
    }

    var p;
    try { p = JSON.parse(e.postData.contents); }
    catch (err) { return responder_({ ok: false, error: 'El dato enviado no se pudo leer.' }); }

    var rol = rol_(p.clave);
    if (!rol) return responder_({ ok: false, error: 'Clave incorrecta. Vuelve a identificarte.' });

    var h = hoja_();

    /* --- proponer: entra una reserva, un movimiento o un evento --- */
    if (p.accion === 'proponer') {
      var c = p.cambio;
      if (!c || !c.cid) return responder_({ ok: false, error: 'La propuesta llegó incompleta.' });

      c.correo = p.correo || '';
      if (!c.ts) c.ts = new Date().toISOString();
      if (!c.estado) c.estado = 'pendiente';

      /* Sin fechas se lee como permanente: así se comportaban las
         propuestas antes de que existiera el alcance. */
      c.desde = fechaTexto_(c.desde);
      c.hasta = fechaTexto_(c.hasta);
      if (c.desde && !c.hasta) c.hasta = c.desde;
      if (!c.alcance) c.alcance = c.desde ? (c.desde === c.hasta ? 'una' : 'rango') : 'ciclo';

      h.appendRow([
        c.cid,
        c.ts,
        c.estado,
        c.por || c.correo,
        c.accion || '',
        c.resumen || '',
        JSON.stringify(c),
        c.alcance,
        c.desde,
        c.hasta
      ]);

      return responder_({ ok: true, cambios: leerCambios_() });
    }

    /* --- resolver: aprobar o rechazar. Solo Dirección Académica. --- */
    if (p.accion === 'resolver') {
      if (rol !== 'admin') {
        return responder_({ ok: false, error: 'Solo Dirección Académica puede aprobar o rechazar.' });
      }
      if (!p.cid) return responder_({ ok: false, error: 'Falta indicar cuál propuesta.' });

      var filas = h.getLastRow() - 1;
      if (filas < 1) return responder_({ ok: false, error: 'No hay propuestas registradas.' });

      var datos = h.getRange(2, 1, filas, CABECERAS.length).getValues();
      var colEstado = CABECERAS.indexOf('estado') + 1;
      var colJson   = CABECERAS.indexOf('json') + 1;

      for (var i = 0; i < datos.length; i++) {
        if (String(datos[i][0]) !== String(p.cid)) continue;

        var obj = {};
        try { obj = JSON.parse(datos[i][colJson - 1]); } catch (err) {}
        obj.estado      = p.estado;
        obj.resueltoPor = p.correo || '';
        obj.resueltoTs  = new Date().toISOString();

        h.getRange(i + 2, colEstado).setValue(p.estado);
        h.getRange(i + 2, colJson).setValue(JSON.stringify(obj));

        return responder_({ ok: true, cambios: leerCambios_() });
      }
      return responder_({ ok: false, error: 'No se encontró esa propuesta.' });
    }

    return responder_({ ok: false, error: 'Acción no reconocida: ' + p.accion });

  } catch (err) {
    return responder_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- prueba rápida desde el propio editor ----------
   Ejecuta prepararHoja() una vez para crear la pestaña CAMBIOS
   antes de publicar. No es obligatorio, la crea sola al primer uso. */
function prepararHoja() {
  hoja_();
  SpreadsheetApp.getActiveSpreadsheet().toast('Hoja CAMBIOS lista.');
}

/* Ejecutar UNA vez tras subir esta versión si la hoja ya tenía datos:
   rellena alcance/desde/hasta de las filas antiguas leyéndolas del json.
   Es opcional; sin ella esas filas siguen valiendo para todo el ciclo. */
function completarAlcances() {
  var h = hoja_();
  var filas = h.getLastRow() - 1;
  if (filas < 1) return;
  var cJson = CABECERAS.indexOf('json') + 1;
  var cAlc  = CABECERAS.indexOf('alcance') + 1;
  var datos = h.getRange(2, cJson, filas, 1).getValues();
  var salida = [];
  for (var i = 0; i < filas; i++) {
    var o = {};
    try { o = JSON.parse(datos[i][0]); } catch (e) {}
    salida.push([o.alcance || 'ciclo', fechaTexto_(o.desde), fechaTexto_(o.hasta)]);
  }
  h.getRange(2, cAlc, filas, 3).setValues(salida);
  SpreadsheetApp.getActiveSpreadsheet().toast(filas + ' fila(s) al día.');
}