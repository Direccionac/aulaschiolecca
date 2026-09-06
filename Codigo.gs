/**
 * Sala compartida — Auditoría de Ambientes, Instituto Chio Lecca
 *
 * Guarda las propuestas de cambio en una hoja de Google y valida las claves
 * del lado del servidor, que es lo único que una página estática no puede hacer.
 *
 * INSTALACIÓN (una sola vez, ver README):
 *   1. Hoja de cálculo nueva → Extensiones → Apps Script → pega este archivo.
 *   2. Ejecuta la función  preparar()  una vez y autoriza los permisos.
 *   3. Configuración del proyecto → Propiedades del script → añade:
 *        CLAVE_ADMIN    (la que aprueba y rechaza)
 *        CLAVE_EDITOR   (la que solo propone)
 *   4. Implementar → Nueva implementación → Aplicación web
 *        Ejecutar como: Yo
 *        Quién tiene acceso: Cualquier usuario
 *      Copia la URL que termina en /exec y pégala en la pestaña Datos del portal.
 */

var HOJA = 'Cambios';
var COLS = ['cid','ts','por','correo','did','accion','aula','dia','ini','fin',
            'motivo','resumen','estado','resueltoPor','tsResuelto',
            'titulo','docente','alumnos','area'];

function preparar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var h = ss.getSheetByName(HOJA) || ss.insertSheet(HOJA);
  // Escribe (o repara) la fila de encabezados. Si el archivo venía de una versión
  // anterior con menos columnas, esto la extiende sin tocar los datos existentes.
  h.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
  h.setFrozenRows(1);
  // Sin esto, Google convierte "13:00" en una fecha de 1899 y el portal ya no la entiende.
  h.getRange(1, 1, h.getMaxRows(), COLS.length).setNumberFormat('@');
  return 'Listo. Ahora define CLAVE_ADMIN y CLAVE_EDITOR en las propiedades del script.';
}

function doGet(e) {
  try {
    return json({ ok: true, cambios: leer() });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var p = JSON.parse(e.postData.contents);
    var props = PropertiesService.getScriptProperties();
    var cAdmin = props.getProperty('CLAVE_ADMIN');
    var cEditor = props.getProperty('CLAVE_EDITOR');

    if (!cAdmin && !cEditor) {
      return json({ ok: false, error: 'El servidor no tiene claves configuradas todavía.' });
    }
    if (!p.clave || (p.clave !== cAdmin && p.clave !== cEditor)) {
      return json({ ok: false, error: 'Clave de edición incorrecta.' });
    }
    var esAdmin = (p.clave === cAdmin);

    if (p.accion === 'proponer') {
      var c = p.cambio || {};
      if (!c.did) return json({ ok: false, error: 'Falta el dictado.' });
      var fila = COLS.map(function (k) {
        if (k === 'correo') return p.correo || '';
        if (k === 'estado') return 'pendiente';
        return c[k] != null ? c[k] : '';
      });
      var hj = hoja();
      hj.appendRow(fila);
      hj.getRange(hj.getLastRow(), 1, 1, COLS.length).setNumberFormat('@');

    } else if (p.accion === 'resolver') {
      if (!esAdmin) return json({ ok: false, error: 'Solo el administrador aprueba o rechaza.' });
      if (['aprobado', 'rechazado'].indexOf(p.estado) < 0) {
        return json({ ok: false, error: 'Estado no válido.' });
      }
      var h = hoja();
      var datos = h.getDataRange().getValues();
      var iCid = COLS.indexOf('cid'), iEst = COLS.indexOf('estado');
      var iQui = COLS.indexOf('resueltoPor'), iTs = COLS.indexOf('tsResuelto');
      var encontrado = false;
      for (var i = 1; i < datos.length; i++) {
        if (String(datos[i][iCid]) === String(p.cid)) {
          h.getRange(i + 1, iEst + 1).setValue(p.estado);
          h.getRange(i + 1, iQui + 1).setValue(p.correo || '');
          h.getRange(i + 1, iTs + 1).setValue(new Date().toISOString());
          encontrado = true;
          break;
        }
      }
      if (!encontrado) return json({ ok: false, error: 'No se encontró esa propuesta.' });

    } else {
      return json({ ok: false, error: 'Acción desconocida: ' + p.accion });
    }

    return json({ ok: true, cambios: leer() });

  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function hoja() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(HOJA) || (function () { preparar(); return ss.getSheetByName(HOJA); })();
}

function leer() {
  var h = hoja();
  if (h.getLastRow() < 2) return [];
  var datos = h.getRange(2, 1, h.getLastRow() - 1, COLS.length).getValues();
  return datos.map(function (r) {
    var o = {};
    COLS.forEach(function (k, i) { o[k] = normaliza(k, r[i]); });
    return o;
  }).filter(function (o) { return o.cid; });
}

/** Devuelve siempre texto. Si Google guardó una hora como fecha, la recupera como HH:mm. */
function normaliza(campo, v) {
  if (v === '' || v === null || v === undefined) return '';
  if (v instanceof Date) {
    if (campo === 'ini' || campo === 'fin') {
      return Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm');
    }
    return v.toISOString();
  }
  return String(v);
}

function json(o) {
  return ContentService
    .createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
