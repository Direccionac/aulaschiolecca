/* ════════════════════════════════════════════════════════════════
   SALA COMPARTIDA — Auditoría de Ambientes 2026-II
   Servidor de la página direccionac.github.io/aulaschiolecca/

   Qué hace: guarda en esta hoja de cálculo las propuestas de cambio
   y de eventos que la gente registra desde la web, y devuelve la
   lista para que todos vean lo mismo. Desde la fase 1 también sirve
   la programación (pestaña PROGRAMACION) y desde la fase 2 permite
   editarla desde el portal.

   Versión 8 (13-10-2026): APROBACIÓN POR EXCEPCIÓN.
   - Los editores ya pueden crear y editar horarios desde el portal.
   - El servidor revisa cada cambio. Si no tiene riesgo (ambiente
     libre, sin cruce de docente, dentro del aforo, fuera de bloqueos,
     sin tocar carga docente ni el horario madre de carrera) se aplica
     solo y Dirección Académica se entera en el resumen diario. Si
     tiene riesgo, queda pendiente y llega el correo de siempre, con
     el motivo por el que llegó.
   - Dirección Académica puede revertir cualquier cambio aplicado solo.
   - Claves personales por usuario (CLAVE_PLOPINTA, CLAVE_MVILCA...).
   - Lo que mueve carga docente se cruza con el padrón del tablero de
     Gestión Docente.
   Versión 7 (02-10-2026): los cruces miran la ocupación real (base +
   cambios aprobados). Versión 6: formato de columnas de la versión 3.

   No toques nada de aquí abajo salvo que sepas lo que haces. Las
   claves NO van en este archivo: se ponen en Configuración del
   proyecto → Propiedades del script.

     CLAVE_ADMIN        la clave de Dirección Académica
     CLAVE_EDITOR       la clave compartida de editores (Leslie,
                        Annie, Orientación al Alumno)
     CLAVE_<USUARIO>    clave personal: lo que va antes de la @ del
                        correo, en mayúsculas. Ej.: CLAVE_PLOPINTA,
                        CLAVE_MVILCA. Si una persona tiene clave
                        personal, la compartida ya no vale para ella.
     DIAS_PERMANENTE    (opcional, 28) a partir de cuántos días un
                        cambio en carrera cuenta como permanente.
     DOCENTES_URL       (opcional) dirección del padrón de Gestión
                        Docente. Ya viene puesta.
   ════════════════════════════════════════════════════════════════ */

var HOJA = 'Cambios';
/* Mismo orden que la versión 3: así están guardadas las filas existentes.
   Lo nuevo va siempre detrás, para que ninguna fila registrada hasta hoy
   se corra de sitio. Versión 8 agrega: nivel (auto | revision | admin),
   razon (por qué fue a revisión), datos (lo que se escribe en la base al
   aprobar) y antes (lo que decía la base, para poder revertir). */
var COLS = ['cid', 'ts', 'por', 'correo', 'did', 'accion', 'aula', 'dia', 'ini', 'fin',
            'motivo', 'resumen', 'estado', 'resueltoPor', 'tsResuelto',
            'titulo', 'docente', 'alumnos', 'area',
            'alcance', 'desde', 'hasta', 'clases',
            'nivel', 'razon', 'datos', 'antes'];

var AUTO = 'Aprobación automática';
var PREFIJO_LOTE = 'Importación · ';

/* ---------- utilidades ---------- */

/* Devuelve la pestaña Cambios y deja la fila 1 con los encabezados correctos.
   Solo toca la fila de títulos: los datos no se mueven. */
function hoja_() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var h = libro.getSheetByName(HOJA);
  if (!h) h = libro.insertSheet(HOJA);
  if (h.getMaxColumns() < COLS.length) {
    h.insertColumnsAfter(h.getMaxColumns(), COLS.length - h.getMaxColumns());
  }
  var actuales = h.getRange(1, 1, 1, COLS.length).getValues()[0];
  var distinto = false;
  for (var i = 0; i < COLS.length; i++) {
    if (String(actuales[i] || '').trim() !== COLS[i]) distinto = true;
  }
  if (distinto) {
    h.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
    h.setFrozenRows(1);
  }
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

/* Una celda de Cambios a texto. Si Google guardó una hora o una fecha como
   objeto Date, se recupera en el formato que entiende el portal. */
function celdaCambio_(campo, v) {
  if (v === '' || v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    if (campo === 'ini' || campo === 'fin') return Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm');
    if (campo === 'desde' || campo === 'hasta') return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    return v.toISOString();
  }
  return String(v);
}

function filaACambio_(r) {
  var o = {};
  for (var j = 0; j < COLS.length; j++) o[COLS[j]] = celdaCambio_(COLS[j], r[j]);
  o.resueltoTs = o.tsResuelto;   /* nombre que usan los avisos por correo */
  return o;
}

/* Al portal no le hace falta el JSON interno de las filas (datos, antes):
   se queda en la hoja y no viaja en cada sincronización. */
function leerCambios_(completo) {
  var h = hoja_();
  if (h.getLastRow() < 2) return [];
  var datos = h.getRange(2, 1, h.getLastRow() - 1, COLS.length).getValues();
  var salida = [];
  for (var i = 0; i < datos.length; i++) {
    var o = filaACambio_(datos[i]);
    if (!o.cid) continue;
    if (!completo) { delete o.datos; delete o.antes; }
    salida.push(o);
  }
  return salida;
}

/* Escribe una fila nueva en Cambios. La fila se marca como texto ANTES de
   escribir: si no, Google convierte "13:00" en una hora de 1899. */
function agregarFila_(c) {
  var fila = COLS.map(function (k) {
    var v = c[k];
    return (v === null || v === undefined) ? '' : String(v);
  });
  var h = hoja_();
  var n = h.getLastRow() + 1;
  h.getRange(n, 1, 1, COLS.length).setNumberFormat('@').setValues([fila]);
}

/* Cambia campos de una fila de Cambios por su cid. Devuelve la fila ya
   actualizada, o null si no existe. */
function actualizarFila_(cid, campos) {
  var h = hoja_();
  var filas = h.getLastRow() - 1;
  if (filas < 1) return null;
  var datos = h.getRange(2, 1, filas, COLS.length).getValues();
  var iCid = COLS.indexOf('cid');
  for (var i = 0; i < datos.length; i++) {
    if (String(datos[i][iCid]) !== String(cid)) continue;
    for (var k in campos) {
      var j = COLS.indexOf(k);
      if (j < 0) continue;
      h.getRange(i + 2, j + 1).setNumberFormat('@').setValue(campos[k] === undefined ? '' : String(campos[k]));
      datos[i][j] = campos[k];
    }
    return filaACambio_(datos[i]);
  }
  return null;
}

function buscarCambio_(cid) {
  var l = leerCambios_(true);
  for (var i = 0; i < l.length; i++) if (String(l[i].cid) === String(cid)) return l[i];
  return null;
}

/* Valida quién escribe.
   - CLAVE_ADMIN sirve para todo.
   - Si la persona tiene clave personal (CLAVE_PLOPINTA para
     plopinta@...), vale solo esa: así nadie firma con su nombre.
   - Si no, vale la clave compartida CLAVE_EDITOR. */
function rol_(clave, correo) {
  var p = PropertiesService.getScriptProperties();
  if (!clave) return null;
  var admin = p.getProperty('CLAVE_ADMIN');
  if (admin && clave === admin) return 'admin';
  var usuario = String(correo || '').split('@')[0].replace(/[^A-Za-z0-9]/g, '').toUpperCase();
  var personal = usuario ? p.getProperty('CLAVE_' + usuario) : null;
  if (personal) return clave === personal ? 'editor' : null;
  var editor = p.getProperty('CLAVE_EDITOR');
  if (editor && clave === editor) return 'editor';
  return null;
}

function uid_() { return 'x' + Utilities.getUuid().replace(/-/g, '').slice(0, 12); }
function ahora_() { return new Date().toISOString(); }

/* ---------- lectura: la página consulta con GET ---------- */

function doGet(e) {
  try {
    var r = { ok: true, cambios: leerCambios_() };
    /* La base solo se manda cuando el portal la pide (?base=1): son más de
       300 filas y no hace falta releerlas en cada sincronización. */
    if (e && e.parameter && e.parameter.base) r.base = leerBase_();
    return responder_(r);
  } catch (err) {
    return responder_({ ok: false, error: String(err) });
  }
}

/* ---------- escritura ---------- */

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

    var rol = rol_(p.clave, p.correo);
    if (!rol) return responder_({ ok: false, error: 'Clave incorrecta. Vuelve a identificarte.' });

    if (p.accion === 'proponer') return responder_(proponer_(p, rol));

    if (p.accion === 'resolver') {
      if (rol !== 'admin') {
        return responder_({ ok: false, error: 'Solo Dirección Académica puede aprobar, rechazar o revertir.' });
      }
      return responder_(resolver_(p));
    }

    if (p.accion === 'leer-plan') return responder_(leerPlan_(p));
    if (p.accion === 'aviso-lote') return responder_(avisarLote_(p.lote, !!p.resuelto));

    if (p.accion === 'base-crear' || p.accion === 'base-editar' || p.accion === 'base-baja') {
      if (rol === 'admin') return responder_(escribirBase_(p));
      return responder_(baseEditor_(p));
    }

    return responder_({ ok: false, error: 'Acción no reconocida: ' + p.accion });

  } catch (err) {
    return responder_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- proponer: reserva, reubicación, baja o evento ---------- */

function firma_(x) {
  var nuevo = !x.did || String(x.did).indexOf('nuevo-') === 0;
  return [x.accion, nuevo ? (x.resumen || '') : x.did, x.aula || '', x.dia || '',
          horaTexto_(x.ini), horaTexto_(x.fin), x.desde || '', x.hasta || ''].join('|');
}

function proponer_(p, rol) {
  var c = p.cambio || {};
  if (!c.cid) return { ok: false, error: 'La propuesta llegó incompleta.' };

  c.correo = p.correo || '';
  if (!c.ts) c.ts = ahora_();

  /* Sin fechas se lee como permanente, igual que antes. */
  c.desde = fechaTexto_(c.desde);
  c.hasta = fechaTexto_(c.hasta);
  if (c.desde && !c.hasta) c.hasta = c.desde;
  if (!c.alcance) c.alcance = c.desde ? (c.desde === c.hasta ? 'una' : 'rango') : 'ciclo';

  var todos = leerCambios_();

  /* El mismo cid dos veces es un reintento del navegador: no se duplica. */
  if (todos.some(function (x) { return String(x.cid) === String(c.cid); })) {
    return { ok: true, cambios: todos, estadoNuevo: 'repetido' };
  }

  /* Una solicitud idéntica ya registrada solo entra si quien la manda
     confirmó en el portal que de verdad quiere otra igual. */
  if (!c.confirmaDuplicado) {
    var fc = firma_(c);
    var igual = todos.filter(function (x) {
      return (x.estado === 'pendiente' || x.estado === 'aprobado') && x.accion === c.accion && firma_(x) === fc;
    })[0];
    if (igual) {
      return { ok: false, duplicado: true,
               error: 'Ya hay una solicitud idéntica ' + (igual.estado === 'aprobado' ? 'aprobada' : 'pendiente') +
                      ' (' + (igual.por || 'sin nombre') + '): ' + (igual.resumen || '') };
    }
  }
  delete c.confirmaDuplicado;

  if (rol === 'admin') {
    /* Lo que registra Dirección Académica ya está aprobado: no tiene
       sentido que se lo apruebe a sí misma. */
    c.estado = 'aprobado'; c.nivel = 'admin';
    c.resueltoPor = c.correo; c.tsResuelto = ahora_();
    agregarFila_(c);
    return { ok: true, cambios: leerCambios_(), estadoNuevo: 'aprobado' };
  }

  var ev = evaluarPropuesta_(c);
  if (!ev.razones.length) {
    c.estado = 'aprobado'; c.nivel = 'auto';
    c.resueltoPor = AUTO; c.tsResuelto = ahora_();
    agregarFila_(c);
    return { ok: true, cambios: leerCambios_(), estadoNuevo: 'aprobado' };
  }

  c.estado = 'pendiente'; c.nivel = 'revision';
  c.razon = ev.razones.join(' ');
  agregarFila_(c);
  /* El correo va al final y aislado: si el envío falla, la solicitud ya
     quedó guardada y el usuario no ve ningún error. */
  avisarSolicitud_(c);
  return { ok: true, cambios: leerCambios_(), estadoNuevo: 'pendiente', razon: c.razon };
}

/* ---------- resolver: aprobar, rechazar o revertir ---------- */

function resolver_(p) {
  if (['aprobado', 'rechazado', 'revertido'].indexOf(p.estado) < 0) {
    return { ok: false, error: 'Estado no válido.' };
  }
  if (!p.cid) return { ok: false, error: 'Falta indicar cuál propuesta.' };
  var c = buscarCambio_(p.cid);
  if (!c) return { ok: false, error: 'No se encontró esa propuesta.' };

  var esBase = String(c.accion).indexOf('base-') === 0;
  var base = null;

  if (p.estado === 'revertido') {
    if (c.estado !== 'aprobado') return { ok: false, error: 'Solo se revierte un cambio aprobado.' };
    if (esBase) {
      var r = revertirBase_(c, p.correo);
      if (!r.ok) return r;
      base = r.base;
    }
    /* Reubicaciones, bajas y eventos dejan de contar en cuanto dejan de
       estar «aprobado»: la ocupación real solo mira los aprobados. */
  } else {
    if (c.estado !== 'pendiente') return { ok: false, error: 'Esa solicitud ya estaba resuelta (' + c.estado + ').' };
    if (p.estado === 'aprobado' && esBase) {
      var d = {};
      try { d = JSON.parse(c.datos || '{}'); } catch (err) {}
      d.correo = c.correo;      /* en la bitácora queda quien lo pidió */
      d.forzar = true;          /* aprobarlo es autorizar lo que se avisó */
      var w = escribirBase_(d);
      if (!w.ok) return { ok: false, error: 'No se pudo aplicar en la programación: ' + w.error };
      base = w.base;
    }
  }

  var fila = actualizarFila_(c.cid, { estado: p.estado, resueltoPor: p.correo || '', tsResuelto: ahora_() });
  /* Las filas de una importación se avisan juntas (aviso-lote), no una por una. */
  if (String(c.motivo || '').indexOf(PREFIJO_LOTE) !== 0) avisarResolucion_(fila);
  var out = { ok: true, cambios: leerCambios_() };
  if (base) out.base = base;
  return out;
}

/* ---------- desde el editor ----------
   Ejecuta prepararHoja() una vez tras pegar esta versión: devuelve los
   encabezados correctos a la fila 1 de Cambios y deja la pestaña en texto
   para que Google no convierta horas ni fechas. No mueve ningún dato. */
function prepararHoja() {
  var h = hoja_();
  h.getRange(1, 1, h.getMaxRows(), COLS.length).setNumberFormat('@');
  var n = Math.max(h.getLastRow() - 1, 0);
  SpreadsheetApp.getActiveSpreadsheet().toast('Cambios lista: ' + n + ' fila(s), encabezados al día.');
  Logger.log('Cambios: %s fila(s). Primera: %s', n, n ? JSON.stringify(leerCambios_()[0]) : '—');
}


/* ════════════════════════════════════════════════════════════════
   LA BASE ACADÉMICA (fase 1)

   La programación vive en cuatro pestañas de esta misma hoja y el
   portal la lee en vivo. El Excel pasa a ser respaldo, no fuente.

   Pestañas: PROGRAMACION (un dictado por fila), AMBIENTES (el
   catálogo de aulas), SECCIONES (padrón y vigencias) y META (la
   ficha del corte). Ejecuta prepararBase() una vez para crearlas.
   ════════════════════════════════════════════════════════════════ */

var H_PROG = 'PROGRAMACION', H_AMB = 'AMBIENTES', H_SEC = 'SECCIONES', H_META = 'META';

var COL_PROG = ['dia', 'turno', 'tipo', 'aula', 'curso', 'seccion', 'docente', 'horario',
                'activo', 'nota', 'ini', 'fin', 'tipoAula', 'aforo', 'cruce', 'programa',
                'grupo', 'alumnos', 'estado', 'id', 'vigDesde', 'vigHasta', 'cadencia',
                'fechas', 'crucePermitido', 'docentes', 'vigProvisional', 'esEvento'];
/* bloqueoTurnos, bloqueoDesde y bloqueoHasta son opcionales: si la pestaña
   AMBIENTES los tiene, el servidor y el portal los usan para saber
   exactamente cuándo rige un bloqueo (ej. SUM: «Mañana,Tarde», 2026-09-08,
   2026-10-31). Si no están, se lee el texto de «bloqueada». */
var COL_AMB  = ['aula', 'tipo', 'aforo', 'capacidad', 'bloques', 'obs', 'maxAlumnos',
                'estado', 'bloqueada', 'bloqueoTurnos', 'bloqueoDesde', 'bloqueoHasta'];
var COL_SEC  = ['clave', 'programa', 'ciclo', 'turno', 'grupo', 'campana', 'alumnos',
                'estado', 'nota', 'vigDesde', 'vigHasta'];

/* Columnas que no son texto plano y hay que devolver con su forma. */
var PROG_HORA  = ['ini', 'fin'];
var PROG_FECHA = ['vigDesde', 'vigHasta'];
var PROG_LISTA = ['fechas', 'docentes'];
var PROG_BOOL  = ['vigProvisional', 'esEvento'];

/* Si Sheets convirtió '08:00' en un objeto hora, vuelve a texto. */
function horaTexto_(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm');
  }
  var m = String(v).trim().match(/^(\d{1,2}):(\d{2})/);
  return m ? ('0' + m[1]).slice(-2) + ':' + m[2] : String(v).trim();
}

/* Lee una pestaña emparejando por nombre de encabezado, no por posición:
   así se puede reordenar o agregar columnas en la hoja sin romper nada. */
function leerTabla_(nombre, cols) {
  var h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre);
  if (!h || h.getLastRow() < 2) return [];
  var ancho = Math.max(h.getLastColumn(), 1);
  var cab = h.getRange(1, 1, 1, ancho).getValues()[0];
  var pos = {};
  for (var c = 0; c < cols.length; c++) pos[cols[c]] = cab.indexOf(cols[c]);
  var datos = h.getRange(2, 1, h.getLastRow() - 1, ancho).getValues();
  var salida = [];
  for (var i = 0; i < datos.length; i++) {
    var o = {}, vacia = true;
    for (var j = 0; j < cols.length; j++) {
      var k = pos[cols[j]];
      if (k < 0) continue;
      var v = datos[i][k];
      if (v !== '' && v !== null && v !== undefined) vacia = false;
      o[cols[j]] = v;
    }
    if (!vacia) salida.push(o);
  }
  return salida;
}

/* Todo lo que no sea lista ni booleano viaja como texto, igual que en el
   JSON que traía el archivo: el portal compara cadenas, no números. */
function aTexto_(o, saltar) {
  for (var k in o) {
    if (saltar.indexOf(k) >= 0) continue;
    o[k] = (o[k] === null || o[k] === undefined) ? '' : String(o[k]).trim();
  }
  return o;
}

function leerProg_() {
  return leerTabla_(H_PROG, COL_PROG).map(function (o) {
    var i, c;
    for (i = 0; i < PROG_HORA.length; i++) o[PROG_HORA[i]] = horaTexto_(o[PROG_HORA[i]]);
    for (i = 0; i < PROG_FECHA.length; i++) {
      c = PROG_FECHA[i];
      var f = fechaTexto_(o[c]);
      if (f) o[c] = f; else delete o[c];
    }
    for (i = 0; i < PROG_LISTA.length; i++) {
      c = PROG_LISTA[i];
      var s = String(o[c] === null || o[c] === undefined ? '' : o[c]).trim();
      if (!s) { delete o[c]; continue; }
      o[c] = s.split(',').map(function (x) { return x.trim(); }).filter(String);
    }
    for (i = 0; i < PROG_BOOL.length; i++) {
      c = PROG_BOOL[i];
      var b = String(o[c] === null || o[c] === undefined ? '' : o[c]).trim().toUpperCase();
      if (b === 'SI' || b === 'SÍ' || b === 'TRUE' || b === 'VERDADERO') o[c] = true;
      else if (b === 'NO' || b === 'FALSE' || b === 'FALSO') o[c] = false;
      else delete o[c];
    }
    aTexto_(o, PROG_LISTA.concat(PROG_BOOL));
    if (!o.horario && o.ini && o.fin) o.horario = o.ini + '-' + o.fin;
    return o;
  });
}

function leerAmbientes_() {
  return leerTabla_(H_AMB, COL_AMB).map(function (o) {
    o.bloqueoDesde = fechaTexto_(o.bloqueoDesde);
    o.bloqueoHasta = fechaTexto_(o.bloqueoHasta);
    aTexto_(o, []);
    /* vacío = disponible; sin columna = sin dato */
    ['bloqueada', 'bloqueoTurnos', 'bloqueoDesde', 'bloqueoHasta'].forEach(function (k) {
      if (!o[k]) delete o[k];
    });
    return o;
  });
}

function leerSecciones_() {
  return leerTabla_(H_SEC, COL_SEC).map(function (o) {
    var i, c;
    for (i = 0; i < PROG_FECHA.length; i++) {
      c = PROG_FECHA[i];
      var f = fechaTexto_(o[c]);
      if (f) o[c] = f; else delete o[c];
    }
    return aTexto_(o, []);
  });
}

function leerMeta_() {
  var m = {};
  leerTabla_(H_META, ['clave', 'valor']).forEach(function (o) {
    m[String(o.clave).trim()] = o.valor;
  });
  return m;
}

/* Devuelve null si la pestaña PROGRAMACION no está cargada: así el portal
   sabe que debe seguir con la base de respaldo del archivo en vez de
   quedarse en blanco. */
function leerBase_() {
  var prog = leerProg_();
  if (!prog.length) return null;
  var m = leerMeta_();
  var provis = String(m.fechasProvisionales || '').split(',')
                 .map(function (s) { return s.trim(); }).filter(String);
  return {
    prog: prog,
    aulas: leerAmbientes_(),
    secciones: leerSecciones_(),
    generado: String(m.generado || ''),
    fuenteQ10: String(m.fuenteQ10 || ''),
    fuentePEC: String(m.fuentePEC || ''),
    libro: String(m.libro || ''),
    aforoEstandar: Number(m.aforoEstandar || 23),
    fechasProvisionales: provis,
    provisionalNota: String(m.provisionalNota || ''),
    filas: prog.length,
    leidoTs: new Date().toISOString()
  };
}

/* Ejecutar UNA vez: crea las pestañas de la base con sus encabezados y deja
   en texto las columnas de hora y de fecha. Si una pestaña ya existe, no se
   toca su contenido. */
function prepararBase() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var tablas = [[H_PROG, COL_PROG], [H_AMB, COL_AMB], [H_SEC, COL_SEC], [H_META, ['clave', 'valor']],
                [H_BITACORA, COL_BITACORA]];
  var hechas = [];
  for (var i = 0; i < tablas.length; i++) {
    var nombre = tablas[i][0], cols = tablas[i][1];
    var h = libro.getSheetByName(nombre);
    if (!h) { h = libro.insertSheet(nombre); hechas.push(nombre); }
    if (h.getMaxColumns() < cols.length) {
      h.insertColumnsAfter(h.getMaxColumns(), cols.length - h.getMaxColumns());
    }
    h.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    h.setFrozenRows(1);
    var texto = PROG_HORA.concat(PROG_FECHA);
    for (var j = 0; j < cols.length; j++) {
      if (texto.indexOf(cols[j]) < 0) continue;
      h.getRange(2, j + 1, Math.max(h.getMaxRows() - 1, 1), 1).setNumberFormat('@');
    }
  }
  libro.toast(hechas.length ? 'Pestañas creadas: ' + hechas.join(', ') : 'Las pestañas ya existían; encabezados al día.');
}

/* Comprobación rápida desde el editor: cuántas filas está leyendo el
   portal y si algo quedó mal parseado. Mira el registro de ejecución. */
function probarBase() {
  var b = leerBase_();
  if (!b) { Logger.log('PROGRAMACION está vacía: el portal usaría la base del archivo.'); return; }
  Logger.log('prog: %s filas · ambientes: %s · secciones: %s · corte: %s',
             b.prog.length, b.aulas.length, b.secciones.length, b.generado);
  Logger.log('primera fila: %s', JSON.stringify(b.prog[0]));
  var malos = b.prog.filter(function (p) { return !/^\d{2}:\d{2}$/.test(p.ini) || !/^\d{2}:\d{2}$/.test(p.fin); });
  Logger.log('filas con hora ilegible: %s', malos.length);
  if (malos.length) Logger.log(JSON.stringify(malos.slice(0, 3)));
  var sinId = b.prog.filter(function (p) { return !p.id; });
  Logger.log('filas sin id: %s', sinId.length);
}


/* ════════════════════════════════════════════════════════════════
   AVISOS AL CORREO

   Desde la versión 8 solo llega correo en el momento cuando algo
   necesita la decisión de Dirección Académica, con el motivo por el
   que no se aplicó solo. Lo que se aplicó solo llega en el resumen
   diario, con un enlace al portal para revertirlo si hiciera falta.

   A quién le llega se configura en Configuración del proyecto →
   Propiedades del script, sin tocar este código:

     AVISOS_PARA           correos separados por coma. Si está
                           vacío, va al dueño del script.
     AVISOS_INSTANTANEOS   'NO' apaga los correos de cada
                           solicitud y deja solo el resumen diario.
     PORTAL_URL            dirección del portal para el enlace.

   Ningún fallo de correo puede tumbar el registro de un cambio:
   todo va dentro de try/catch y siempre después de escribir la fila.
   ════════════════════════════════════════════════════════════════ */

var PORTAL_POR_DEFECTO = 'https://direccionac.github.io/aulaschiolecca/';

function prop_(clave, porDefecto) {
  var v = PropertiesService.getScriptProperties().getProperty(clave);
  return (v === null || v === undefined || String(v).trim() === '') ? porDefecto : String(v).trim();
}

function destinatarios_() {
  var lista = prop_('AVISOS_PARA', '').split(',')
                .map(function (s) { return s.trim(); })
                .filter(function (s) { return s.indexOf('@') > 0; });
  if (lista.length) return lista.join(',');
  try { return Session.getEffectiveUser().getEmail(); } catch (e) { return ''; }
}

function portal_() { return prop_('PORTAL_URL', PORTAL_POR_DEFECTO); }

/* Nombres legibles de cada tipo de solicitud. */
function tituloAccion_(a) {
  if (a === 'mover') return 'Reubicación de un dictado';
  if (a === 'desactivar') return 'Baja de clases de un dictado';
  if (a === 'evento') return 'Reserva de ambiente para un evento';
  if (a === 'base-crear') return 'Dictado nuevo en la programación';
  if (a === 'base-editar') return 'Edición de un dictado';
  if (a === 'base-baja') return 'Baja de un dictado';
  return 'Solicitud de cambio';
}

/* La vigencia en una línea, como la lee una persona. */
function vigenciaTexto_(c) {
  if (!c.desde) return String(c.accion || '').indexOf('base-') === 0 ? 'La del dictado' : 'Todo el ciclo';
  if (c.desde === c.hasta) return 'Solo el ' + c.desde;
  return 'Del ' + c.desde + ' al ' + c.hasta +
         (c.clases ? ' · ' + c.clases + ' clase(s)' : '');
}

function escapaHtml_(v) {
  return String(v === null || v === undefined ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/* Tabla de dos columnas con los datos que traiga la solicitud. */
function filasHtml_(pares) {
  var h = '';
  for (var i = 0; i < pares.length; i++) {
    var k = pares[i][0], v = pares[i][1];
    if (v === '' || v === null || v === undefined) continue;
    h += '<tr>' +
         '<td style="padding:6px 14px 6px 0;color:#666;font-size:13px;vertical-align:top;white-space:nowrap">' + escapaHtml_(k) + '</td>' +
         '<td style="padding:6px 0;font-size:14px;vertical-align:top">' + escapaHtml_(v) + '</td>' +
         '</tr>';
  }
  return h;
}

function cuerpoHtml_(encabezado, color, c, pie) {
  var lugar = [];
  if (c.aula) lugar.push('Ambiente ' + c.aula);
  if (c.dia) lugar.push(c.dia);
  if (c.ini && c.fin) lugar.push(c.ini + ' – ' + c.fin);

  var pares = [
    ['Solicitud', tituloAccion_(c.accion)],
    ['Qué pide', c.resumen || ''],
    ['Por qué te llega', c.razon || ''],
    ['Dónde', lugar.join(' · ')],
    ['Docente', c.docente || ''],
    ['Vigencia', vigenciaTexto_(c)],
    ['Motivo', c.motivo || ''],
    ['Pidió', (c.por || '') + (c.correo ? ' (' + c.correo + ')' : '')],
    ['Registrado', c.ts ? fechaHoraTexto_(c.ts) : ''],
    ['Código', c.cid || '']
  ];

  return '' +
    '<div style="font-family:Helvetica,Arial,sans-serif;max-width:560px;color:#1a1a1a">' +
      '<div style="border-left:3px solid ' + color + ';padding:2px 0 2px 14px;margin-bottom:18px">' +
        '<div style="font-size:12px;letter-spacing:.08em;text-transform:uppercase;color:#888">Auditoría de Ambientes 2026-II</div>' +
        '<div style="font-size:19px;font-weight:600;margin-top:3px">' + escapaHtml_(encabezado) + '</div>' +
      '</div>' +
      '<table cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%">' + filasHtml_(pares) + '</table>' +
      '<p style="margin:22px 0 0"><a href="' + portal_() + '" style="background:#1a1a1a;color:#fff;text-decoration:none;padding:10px 18px;border-radius:6px;font-size:14px;display:inline-block">Abrir el portal</a></p>' +
      '<p style="font-size:12.5px;color:#888;margin:16px 0 0;line-height:1.5">' + pie + '</p>' +
    '</div>';
}

/* Fecha y hora en la zona del script, legible. */
function fechaHoraTexto_(iso) {
  try {
    return Utilities.formatDate(new Date(iso), Session.getScriptTimeZone(), "dd/MM/yyyy 'a las' HH:mm");
  } catch (e) { return String(iso); }
}

/* Recorta el resumen para que el asunto quepa en la bandeja. */
function corto_(s, n) {
  s = String(s || '').replace(/\s+/g, ' ').trim();
  return s.length > n ? s.slice(0, n - 1) + '…' : s;
}

function avisarSolicitud_(c) {
  try {
    if (prop_('AVISOS_INSTANTANEOS', 'SI').toUpperCase() === 'NO') return;
    var para = destinatarios_();
    if (!para) return;
    MailApp.sendEmail({
      to: para,
      name: 'Auditoría de Ambientes',
      subject: '[Ambientes] Pendiente de aprobar · ' + corto_(c.resumen || tituloAccion_(c.accion), 70),
      htmlBody: cuerpoHtml_('Hay una solicitud esperando tu aprobación', '#c0392b', c,
        'No se aplicó sola por lo que dice «Por qué te llega». Entra al portal, pestaña Cambios, para aprobarla o rechazarla. Mientras siga pendiente no altera la programación.')
    });
  } catch (err) {
    /* El cambio ya está guardado: un fallo de correo no lo invalida. */
    Logger.log('Aviso de solicitud no enviado: %s', err);
  }
}

function avisarResolucion_(c) {
  try {
    if (!c) return;
    var est = String(c.estado).toLowerCase();
    var aprobado = est === 'aprobado', revertido = est === 'revertido';
    /* Quien pidió, más Dirección Académica, sin repetir a nadie. */
    var para = [], vistos = {};
    var candidatos = [];
    if (c.correo && c.correo.indexOf('@') > 0) candidatos.push(c.correo);
    candidatos = candidatos.concat(destinatarios_().split(','));
    for (var i = 0; i < candidatos.length; i++) {
      var e = String(candidatos[i]).trim(), k = e.toLowerCase();
      if (e.indexOf('@') < 1 || vistos[k]) continue;
      vistos[k] = 1; para.push(e);
    }
    if (!para.length) return;
    var palabra = aprobado ? 'Aprobado' : (revertido ? 'Revertido' : 'Rechazado');
    MailApp.sendEmail({
      to: para.join(','),
      name: 'Auditoría de Ambientes',
      subject: '[Ambientes] ' + palabra + ' · ' + corto_(c.resumen || '', 70),
      htmlBody: cuerpoHtml_(
        aprobado ? 'Dirección Académica aprobó la solicitud'
                 : (revertido ? 'Dirección Académica revirtió un cambio que se había aplicado'
                              : 'Dirección Académica rechazó la solicitud'),
        aprobado ? '#1e8449' : '#888', c,
        (aprobado ? 'El cambio ya está aplicado en el portal. '
                  : (revertido ? 'La programación volvió a como estaba antes de este cambio. '
                               : 'La programación se queda como estaba. ')) +
        'Resolvió ' + escapaHtml_(c.resueltoPor || 'Dirección Académica') +
        (c.resueltoTs ? ' el ' + escapaHtml_(fechaHoraTexto_(c.resueltoTs)) : '') + '.')
    });
  } catch (err) {
    Logger.log('Aviso de resolución no enviado: %s', err);
  }
}

/* ---------- resumen diario ----------
   Tres bloques: lo que espera tu decisión, lo que se aplicó solo en las
   últimas 24 horas (para que sepas qué se está haciendo sin tener que
   aprobarlo) y lo que resolviste tú. Solo manda correo si hay algo que
   contar. El activador se crea con configurarResumen(). */
function resumenDiario() {
  var cambios = leerCambios_();
  var ayer = new Date(Date.now() - 24 * 3600 * 1000).toISOString();
  var pend = cambios.filter(function (c) { return c.estado === 'pendiente'; });
  var solos = cambios.filter(function (c) {
    return c.nivel === 'auto' && String(c.tsResuelto || c.ts || '') > ayer;
  });
  var resueltos = cambios.filter(function (c) {
    return c.nivel !== 'auto' && c.estado !== 'pendiente' && c.resueltoTs && c.resueltoTs > ayer;
  });
  if (!pend.length && !solos.length && !resueltos.length) return;   /* sin ruido */

  var para = destinatarios_();
  if (!para) return;

  function lista(arr, vacio, conRazon) {
    if (!arr.length) return '<p style="font-size:14px;color:#888;margin:4px 0 0">' + vacio + '</p>';
    return '<ul style="font-size:14px;padding-left:20px;margin:6px 0 0;line-height:1.6">' +
      arr.map(function (c) {
        var extra = (c.por || 'sin nombre') + ' · ' + vigenciaTexto_(c);
        if (c.estado === 'revertido') extra += ' · revertido';
        if (c.estado === 'rechazado') extra += ' · rechazado';
        return '<li>' + escapaHtml_(c.resumen || tituloAccion_(c.accion)) +
               '<br><span style="font-size:12.5px;color:#888">' + escapaHtml_(extra) + '</span>' +
               (conRazon && c.razon ? '<br><span style="font-size:12.5px;color:#c0392b">' + escapaHtml_(c.razon) + '</span>' : '') +
               '</li>';
      }).join('') + '</ul>';
  }

  var html = '<div style="font-family:Helvetica,Arial,sans-serif;max-width:560px;color:#1a1a1a">' +
    '<div style="font-size:12px;letter-spacing:.08em;text-transform:uppercase;color:#888">Auditoría de Ambientes 2026-II</div>' +
    '<div style="font-size:19px;font-weight:600;margin:3px 0 20px">Resumen del día</div>' +
    '<div style="font-size:15px;font-weight:600">Esperando tu decisión (' + pend.length + ')</div>' +
    lista(pend, 'Nada pendiente.', true) +
    '<div style="font-size:15px;font-weight:600;margin-top:22px">Se aplicó solo en las últimas 24 horas (' + solos.length + ')</div>' +
    '<p style="font-size:12.5px;color:#888;margin:2px 0 0">Pasaron todas las revisiones. Si algo no corresponde, en la pestaña Cambios tienes el botón Revertir.</p>' +
    lista(solos, 'Nada.', false) +
    '<div style="font-size:15px;font-weight:600;margin-top:22px">Resuelto por ti en las últimas 24 horas (' + resueltos.length + ')</div>' +
    lista(resueltos, 'Sin movimientos.', false) +
    '<p style="margin:24px 0 0"><a href="' + portal_() + '" style="background:#1a1a1a;color:#fff;text-decoration:none;padding:10px 18px;border-radius:6px;font-size:14px;display:inline-block">Abrir el portal</a></p>' +
    '</div>';

  try {
    MailApp.sendEmail({
      to: para, name: 'Auditoría de Ambientes',
      subject: '[Ambientes] Resumen: ' + pend.length + ' por decidir · ' + solos.length + ' aplicado(s) solo(s)',
      htmlBody: html
    });
  } catch (err) { Logger.log('Resumen diario no enviado: %s', err); }
}

/* Ejecutar UNA vez desde el editor: deja un solo activador diario del
   resumen, a las 6 de la tarde (cambia HORA si prefieres otra). Borra
   los activadores anteriores de resumenDiario para que no lleguen dos. */
function configurarResumen() {
  var HORA = 18;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'resumenDiario') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('resumenDiario').timeBased().everyDays(1).atHour(HORA).create();
  Logger.log('Resumen diario programado a las %s:00 (zona %s).', HORA, Session.getScriptTimeZone());
}

/* Prueba: manda un correo de ejemplo a los destinatarios configurados
   sin registrar nada en la hoja. Ejecutar desde el editor. */
function probarAviso() {
  var para = destinatarios_();
  Logger.log('Destinatarios: %s', para || '(ninguno)');
  Logger.log('Envíos que te quedan hoy: %s', MailApp.getRemainingDailyQuota());
  if (!para) { Logger.log('Configura AVISOS_PARA en Propiedades del script.'); return; }
  avisarSolicitud_({
    cid: 'prueba-' + Date.now(), accion: 'mover', aula: '106', dia: 'Miércoles',
    ini: '15:30', fin: '17:50', desde: '', hasta: '',
    por: 'Prueba de configuración', correo: para.split(',')[0],
    motivo: 'Correo de prueba: no corresponde a ninguna solicitud real.',
    razon: 'Ejemplo: cambio de todo el ciclo en un dictado de carrera.',
    resumen: 'DM-6B-T — mover del 106 al 204, miércoles 13:00-17:45',
    ts: new Date().toISOString()
  });
  Logger.log('Correo de prueba enviado a %s', para);
}


/* ════════════════════════════════════════════════════════════════
   ESCRITURA EN LA BASE (fase 2)

   Se edita la programación desde el portal y esto la escribe en la
   pestaña PROGRAMACION. Tres reglas que no se negocian:

   1. Nada se borra. Dar de baja es poner activo = NO; la fila se
      queda con su historia.
   2. Todo queda en la BITACORA: qué fila, qué decía antes, qué dice
      ahora, quién y cuándo.
   3. Si el cambio genera un cruce, se devuelve como advertencia y no
      se escribe. El portal la muestra y, si aun así corresponde
      (hay solapes autorizados), se reenvía con forzar = true.

   Dirección Académica escribe directo. Los editores pasan por
   baseEditor_ (más abajo): lo que no tiene riesgo se escribe solo y
   lo demás queda pendiente.
   ════════════════════════════════════════════════════════════════ */

var H_BITACORA = 'BITACORA';
var COL_BITACORA = ['ts', 'quien', 'accion', 'id', 'resumen', 'antes', 'despues'];

var DIAS_VALIDOS = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];

/* Campos que el portal puede escribir. Lo que no esté aquí, se ignora:
   así un campo calculado no se corrompe desde el navegador. */
var CAMPOS_EDITABLES = ['dia', 'tipo', 'aula', 'curso', 'seccion', 'docente', 'activo', 'nota',
                        'ini', 'fin', 'programa', 'grupo', 'alumnos', 'vigDesde', 'vigHasta',
                        'cadencia', 'fechas', 'crucePermitido'];

function minutos_(h) {
  var m = String(h || '').match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return NaN;
  var n = (+m[1]) * 60 + (+m[2]);
  return (n >= 0 && n < 24 * 60) ? n : NaN;
}

function indiceProg_() {
  var h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(H_PROG);
  if (!h) return null;
  var ancho = Math.max(h.getLastColumn(), 1);
  var cab = h.getRange(1, 1, 1, ancho).getValues()[0];
  var idx = {};
  for (var i = 0; i < cab.length; i++) {
    var nombre = String(cab[i]).trim();
    if (nombre) idx[nombre] = i;
  }
  return { hoja: h, ancho: ancho, idx: idx };
}

/* Localiza una fila por su id. Devuelve el número de fila real de la hoja. */
function filaDeId_(tabla, id) {
  if (tabla.idx.id === undefined || tabla.hoja.getLastRow() < 2) return 0;
  var col = tabla.hoja.getRange(2, tabla.idx.id + 1, tabla.hoja.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < col.length; i++) {
    if (String(col[i][0]).trim() === String(id).trim()) return i + 2;
  }
  return 0;
}

/* Un dictado a texto plano, con el turno y los derivados al día. */
function normalizarDictado_(d) {
  var o = {};
  for (var i = 0; i < CAMPOS_EDITABLES.length; i++) {
    var c = CAMPOS_EDITABLES[i];
    if (d[c] === undefined || d[c] === null) continue;
    o[c] = (Object.prototype.toString.call(d[c]) === '[object Array]')
      ? d[c].join(', ') : String(d[c]).trim();
  }
  if (o.ini && o.fin) {
    o.horario = o.ini + '-' + o.fin;
    var m = minutos_(o.ini);
    o.turno = m < 13 * 60 ? 'Mañana' : (m < 18 * 60 ? 'Tarde' : 'Noche');
  }
  /* tipoAula y aforo salen del catálogo, no del navegador. */
  if (o.aula) {
    var amb = leerAmbientes_().filter(function (a) { return a.aula === o.aula; })[0];
    if (amb) { o.tipoAula = amb.tipo || ''; o.aforo = amb.aforo || ''; }
    else { o.tipoAula = '—'; o.aforo = '—'; }
  }
  if (!o.activo) o.activo = 'SÍ';
  if (!o.tipo) o.tipo = 'CARRERA';
  if (!o.cruce) o.cruce = 'OK';
  return o;
}

/* Estado de aforo, con las mismas reglas que muestra el portal. */
function estadoAforo_(o) {
  if (String(o.aula).toUpperCase() === 'VIRTUAL') return 'VIRTUAL';
  var n = parseInt(o.alumnos, 10), a = parseInt(o.aforo, 10);
  if (isNaN(n) || isNaN(a)) return 'OK';
  if (n > a) return '⚠ SOBRE AFORO +' + (n - a);
  if (n >= a - 1) return 'AL LÍMITE';
  return 'OK';
}

function validarDictado_(o, esNuevo) {
  var e = [];
  if (!o.aula) e.push('Falta el ambiente.');
  if (!o.seccion) e.push('Falta la sección.');
  if (!o.curso) e.push('Falta el curso.');
  if (!o.dia) e.push('Falta el día.');
  else if (DIAS_VALIDOS.indexOf(o.dia) < 0) e.push('El día debe ser uno de: ' + DIAS_VALIDOS.join(', ') + '.');
  var a = minutos_(o.ini), b = minutos_(o.fin);
  if (isNaN(a) || isNaN(b)) e.push('Las horas deben ir como HH:MM en formato de 24 horas.');
  else if (b <= a) e.push('La hora de fin debe ser posterior a la de inicio.');
  if (o.activo && ['SÍ', 'SI', 'NO'].indexOf(o.activo.toUpperCase()) < 0) e.push('«Activo» debe ser SÍ o NO.');
  ['vigDesde', 'vigHasta'].forEach(function (c) {
    if (o[c] && !/^\d{4}-\d{2}-\d{2}$/.test(o[c])) e.push('La fecha de ' + c + ' debe ir como AAAA-MM-DD.');
  });
  if (o.vigDesde && o.vigHasta && o.vigHasta < o.vigDesde) e.push('La vigencia termina antes de empezar.');
  return e;
}

/* ════════════════════════════════════════════════════════════════
   OCUPACIÓN REAL DE LOS AMBIENTES
   La base dice dónde se programó cada dictado; los cambios aprobados
   dicen dónde está de verdad. Aquí se combinan las dos cosas:
   - una reubicación aprobada saca el dictado de su aula original en
     sus fechas y lo pone en la nueva;
   - una baja aprobada lo saca en sus fechas;
   - un evento aprobado ocupa su ambiente en sus fechas.
   Si un dictado tiene varios cambios, manda el último aprobado.
   El mismo código corre en el portal y en el servidor, para que los
   dos digan lo mismo. (Versión 8: cada ocupación lleva también su
   docente, para revisar cruces de docente en el servidor.)
   ════════════════════════════════════════════════════════════════ */
var OCUP = (function () {
  var ABIERTO_D = '0000-01-01', ABIERTO_H = '9999-12-31';
  var MESES = { ene: 1, feb: 2, mar: 3, abr: 4, may: 5, jun: 6, jul: 7, ago: 8, set: 9, sep: 9, oct: 10, nov: 11, dic: 12 };
  var NUMDIA = { 'Domingo': 0, 'Lunes': 1, 'Martes': 2, 'Miércoles': 3, 'Jueves': 4, 'Viernes': 5, 'Sábado': 6 };

  function iso(v) { var s = String(v || '').trim(); return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : ''; }
  function dosDig(n) { return ('0' + n).slice(-2); }
  function hoy() { var d = new Date(); return d.getFullYear() + '-' + dosDig(d.getMonth() + 1) + '-' + dosDig(d.getDate()); }
  function masDias(f, n) { var p = f.split('-'); var d = new Date(+p[0], +p[1] - 1, +p[2] + n); return d.getFullYear() + '-' + dosDig(d.getMonth() + 1) + '-' + dosDig(d.getDate()); }
  function semana(f) { var p = f.split('-'); return new Date(+p[0], +p[1] - 1, +p[2]).getDay(); }
  function hhmm(v) { var m = String(v || '').match(/(\d{1,2}):(\d{2})/); return m ? dosDig(+m[1]) + ':' + m[2] : ''; }
  function min(v) { var h = hhmm(v); if (!h) return NaN; return (+h.slice(0, 2)) * 60 + (+h.slice(3)); }
  function corta(d, mes, anio) { var m = MESES[String(mes).toLowerCase().slice(0, 3)]; return m ? anio + '-' + dosDig(m) + '-' + dosDig(d) : ''; }

  /* Fechas de un cambio. Las filas antiguas no las guardaban: se leen del
     resumen ("desde el 28 set hasta fin de ciclo", "solo el 5 oct"...).
     Sin nada legible, vale todo el ciclo. */
  function alcance(c) {
    var d = iso(c.desde), h = iso(c.hasta);
    if (d || h) return { desde: d || ABIERTO_D, hasta: h || d };
    var t = String(c.resumen || ''), y = hoy().slice(0, 4), m, a, b;
    if ((m = /solo el\s+(\d{1,2})\s+([a-zA-Z]{3})/i.exec(t))) { a = corta(m[1], m[2], y); if (a) return { desde: a, hasta: a }; }
    if ((m = /(\d{1,2})\s+([a-zA-Z]{3})\s*(?:\u2192|->)\s*(\d{1,2})\s+([a-zA-Z]{3})/.exec(t))) {
      a = corta(m[1], m[2], y); b = corta(m[3], m[4], y); if (a && b) return { desde: a, hasta: b };
    }
    if ((m = /desde el\s+(\d{1,2})\s+([a-zA-Z]{3})\s+hasta fin de ciclo/i.exec(t))) { a = corta(m[1], m[2], y); if (a) return { desde: a, hasta: ABIERTO_H }; }
    if ((m = /\u00b7\s*(\d{1,2})\s+([a-zA-Z]{3})\s*$/.exec(t))) { a = corta(m[1], m[2], y); if (a) return { desde: a, hasta: a }; }
    return { desde: ABIERTO_D, hasta: ABIERTO_H };
  }

  /* Reemplaza en la línea de tiempo el tramo [d, h] por "slot" (null = libre). */
  function cortar(segs, d, h, slot) {
    var r = [];
    segs.forEach(function (s) {
      if (s.hasta < d || s.desde > h) { r.push(s); return; }
      if (s.desde < d) r.push({ desde: s.desde, hasta: masDias(d, -1), slot: s.slot });
      r.push({ desde: s.desde > d ? s.desde : d, hasta: s.hasta < h ? s.hasta : h, slot: slot });
      if (s.hasta > h) r.push({ desde: masDias(h, 1), hasta: s.hasta, slot: s.slot });
    });
    return r;
  }

  /* Lista de ocupaciones efectivas: {id, etiqueta, aula, dia, ini, fin,
     docente, vigDesde, vigHasta, fechas}. vigDesde/vigHasta vacíos = sin límite. */
  function ocupaciones(prog, cambios) {
    var aprob = (cambios || []).filter(function (c) { return String(c.estado) === 'aprobado'; })
      .sort(function (a, b) { return String(a.tsResuelto || a.ts || '').localeCompare(String(b.tsResuelto || b.ts || '')); });
    var porDid = {};
    aprob.forEach(function (c) { var k = String(c.did || ''); (porDid[k] = porDid[k] || []).push(c); });
    var out = [];

    function aplicar(base, clave) {
      var segs = [{ desde: ABIERTO_D, hasta: ABIERTO_H, slot: base }];
      (porDid[clave] || []).forEach(function (c) {
        if (c.accion !== 'mover' && c.accion !== 'desactivar') return;
        var a = alcance(c), nuevo = null;
        if (c.accion === 'mover') {
          nuevo = {};
          for (var k in base) nuevo[k] = base[k];
          if (c.aula) nuevo.aula = String(c.aula);
          if (c.dia) nuevo.dia = String(c.dia);
          if (hhmm(c.ini)) nuevo.ini = hhmm(c.ini);
          if (hhmm(c.fin)) nuevo.fin = hhmm(c.fin);
          nuevo.movido = true;
        }
        segs = cortar(segs, a.desde, a.hasta, nuevo);
      });
      segs.forEach(function (s) {
        if (!s.slot) return;
        var vd = iso(s.slot.vigDesde), vh = iso(s.slot.vigHasta);
        var d = (vd && vd > s.desde) ? vd : s.desde, h = (vh && vh < s.hasta) ? vh : s.hasta;
        if (d > h) return;
        var o = {};
        for (var k in s.slot) o[k] = s.slot[k];
        o.vigDesde = d === ABIERTO_D ? '' : d;
        o.vigHasta = h === ABIERTO_H ? '' : h;
        out.push(o);
      });
    }

    (prog || []).forEach(function (q) {
      if (String(q.activo).toUpperCase() === 'NO') return;
      aplicar({ id: String(q.id), etiqueta: q.seccion + ' (' + hhmm(q.ini) + '-' + hhmm(q.fin) + ', ' + q.curso + ')',
                aula: String(q.aula), dia: q.dia, ini: hhmm(q.ini), fin: hhmm(q.fin), docente: q.docente || '',
                vigDesde: q.vigDesde || '', vigHasta: q.vigHasta || '', fechas: q.fechas || '' }, String(q.id));
    });
    aprob.forEach(function (c) {
      if (c.accion !== 'evento' || !c.aula || !c.dia) return;
      var a = alcance(c);
      aplicar({ id: 'ev-' + c.cid, etiqueta: 'Evento «' + (c.titulo || c.resumen || c.cid) + '» (' + hhmm(c.ini) + '-' + hhmm(c.fin) + ')',
                aula: String(c.aula), dia: c.dia, ini: hhmm(c.ini), fin: hhmm(c.fin), docente: '',
                vigDesde: a.desde === ABIERTO_D ? '' : a.desde, vigHasta: a.hasta === ABIERTO_H ? '' : a.hasta, fechas: '' }, 'ev-' + c.cid);
    });
    return out;
  }

  function listaFechas(v) {
    if (!v) return [];
    return (Object.prototype.toString.call(v) === '[object Array]' ? v : String(v).split(',')).map(iso).filter(String);
  }

  /* Primera fecha de clase que comparten x e y desde hoy, o ''. */
  function fechaComun(x, y) {
    var desde = [hoy(), iso(x.vigDesde), iso(y.vigDesde)].filter(String).sort().pop();
    var hs = [iso(x.vigHasta), iso(y.vigHasta)].filter(String).sort();
    var hasta = hs.length ? hs[0] : ABIERTO_H;
    if (desde > hasta) return '';
    var fx = listaFechas(x.fechas), fy = listaFechas(y.fechas);
    var dentro = function (f) { return f >= desde && f <= hasta; };
    if (fx.length && fy.length) return fx.filter(function (f) { return fy.indexOf(f) >= 0 && dentro(f); }).sort()[0] || '';
    if (fx.length || fy.length) return (fx.length ? fx : fy).filter(dentro).sort()[0] || '';
    var obj = NUMDIA[x.dia]; if (obj === undefined) return desde;
    var f = desde;
    for (var i = 0; i < 7 && f <= hasta; i++) { if (semana(f) === obj) return f; f = masDias(f, 1); }
    return '';
  }

  /* Cruces de un dictado "o" contra la ocupación real. idPropio se ignora
     (al editar, el dictado no choca consigo mismo). */
  function cruces(o, prog, cambios, idPropio) { return crucesEn(o, ocupaciones(prog, cambios), idPropio); }
  function crucesEn(o, lista, idPropio) {
    if (String(o.activo).toUpperCase() === 'NO' || String(o.aula).toUpperCase() === 'VIRTUAL') return [];
    var a = min(o.ini), b = min(o.fin);
    if (isNaN(a) || isNaN(b) || b <= a) return [];
    var r = [];
    lista.forEach(function (q) {
      if (idPropio && q.id === String(idPropio)) return;
      if (q.aula !== String(o.aula) || q.dia !== o.dia) return;
      var qa = min(q.ini), qb = min(q.fin);
      if (isNaN(qa) || isNaN(qb) || !(qa < b && a < qb)) return;
      var f = fechaComun(o, q);
      if (f) r.push(q.etiqueta + ', desde el ' + f.slice(8, 10) + '-' + f.slice(5, 7) + '-' + f.slice(0, 4) + (q.movido ? ' — reubicado aquí' : ''));
    });
    return r;
  }

  return { ocupaciones: ocupaciones, cruces: cruces, crucesEn: crucesEn, fechaComun: fechaComun, alcance: alcance, min: min };
})();

/* Cruces contra la ocupación REAL: la base más los cambios aprobados
   (reubicaciones, bajas y eventos). Ver el bloque OCUPACIÓN de arriba. */
function crucesDe_(o, idPropio) {
  return OCUP.cruces(o, leerProg_(), leerCambios_(), idPropio);
}

function bitacora_(accion, id, resumen, antes, despues, quien) {
  try {
    var libro = SpreadsheetApp.getActiveSpreadsheet();
    var h = libro.getSheetByName(H_BITACORA);
    if (!h) {
      h = libro.insertSheet(H_BITACORA);
      h.getRange(1, 1, 1, COL_BITACORA.length).setValues([COL_BITACORA]).setFontWeight('bold');
      h.setFrozenRows(1);
    }
    h.appendRow([new Date().toISOString(), quien || '', accion, id || '', resumen || '',
                 antes ? JSON.stringify(antes) : '', despues ? JSON.stringify(despues) : '']);
  } catch (err) {
    Logger.log('Bitácora no escrita: %s', err);
  }
}

function escribirBase_(p) {
  var tabla = indiceProg_();
  if (!tabla) return { ok: false, error: 'No existe la pestaña ' + H_PROG + '. Ejecuta prepararBase() primero.' };
  if (tabla.idx.id === undefined) return { ok: false, error: 'La pestaña ' + H_PROG + ' no tiene columna «id».' };

  var quien = p.correo || '';

  /* ---- dar de baja: activo = NO, la fila se queda ---- */
  if (p.accion === 'base-baja') {
    var filaB = filaDeId_(tabla, p.id);
    if (!filaB) return { ok: false, error: 'No se encontró el dictado ' + p.id + ' en la base.' };
    if (tabla.idx.activo === undefined) return { ok: false, error: 'La pestaña no tiene columna «activo».' };
    var antesB = tabla.hoja.getRange(filaB, 1, 1, tabla.ancho).getValues()[0];
    tabla.hoja.getRange(filaB, tabla.idx.activo + 1).setValue('NO');
    if (p.nota && tabla.idx.nota !== undefined) {
      var previa = String(tabla.hoja.getRange(filaB, tabla.idx.nota + 1).getValue() || '').trim();
      tabla.hoja.getRange(filaB, tabla.idx.nota + 1).setValue(
        (String(p.nota).trim() + (previa ? ' ' + previa : '')).slice(0, 4000));
    }
    bitacora_('baja', p.id, p.resumen || '', antesB, ['activo → NO'], quien);
    return { ok: true, base: leerBase_(), mensaje: 'Dictado dado de baja. La fila se conserva en la base con activo = NO.' };
  }

  /* ---- crear o editar ---- */
  var o = normalizarDictado_(p.dictado || {});
  var esNuevo = (p.accion === 'base-crear');
  var errores = validarDictado_(o, esNuevo);
  if (errores.length) return { ok: false, error: errores.join(' ') };

  var id = esNuevo ? (String(p.dictado && p.dictado.id || '').trim() || 'w' + Utilities.getUuid().slice(0, 7))
                   : String(p.id || '').trim();
  if (!id) return { ok: false, error: 'Falta el id del dictado.' };

  var fila = filaDeId_(tabla, id);
  if (esNuevo && fila) return { ok: false, error: 'Ya existe un dictado con el id ' + id + '.' };
  if (!esNuevo && !fila) return { ok: false, error: 'No se encontró el dictado ' + id + ' en la base.' };

  var choques = crucesDe_(o, esNuevo ? '' : id);
  if (choques.length && !p.forzar) {
    return { ok: false, cruce: choques,
             error: 'Choca con ' + choques.join('; ') + ' en el ambiente ' + o.aula + ' el ' + o.dia + '.' };
  }
  if (choques.length && p.forzar) {
    o.cruce = 'SOLAPE AUTORIZADO';
    if (!o.crucePermitido) {
      o.crucePermitido = 'Solape registrado desde el portal por ' + (quien || 'Dirección Académica') +
                         ' el ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy') + '.';
    }
  }
  o.estado = estadoAforo_(o);
  o.id = id;

  var antes = fila ? tabla.hoja.getRange(fila, 1, 1, tabla.ancho).getValues()[0] : null;

  if (esNuevo) {
    var nueva = [];
    for (var i = 0; i < tabla.ancho; i++) nueva.push('');
    for (var k in o) {
      if (tabla.idx[k] !== undefined) nueva[tabla.idx[k]] = o[k];
    }
    tabla.hoja.appendRow(nueva);
    bitacora_('crear', id, p.resumen || (o.seccion + ' ' + o.dia + ' ' + o.horario + ' en ' + o.aula), null, o, quien);
  } else {
    /* Solo se tocan las celdas que vienen: lo que no se manda, no se pisa. */
    for (var c in o) {
      if (tabla.idx[c] === undefined) continue;
      tabla.hoja.getRange(fila, tabla.idx[c] + 1).setValue(o[c]);
    }
    bitacora_('editar', id, p.resumen || (o.seccion + ' ' + o.dia + ' ' + o.horario + ' en ' + o.aula), antes, o, quien);
  }

  return { ok: true, base: leerBase_(), id: id,
           mensaje: (esNuevo ? 'Dictado creado' : 'Dictado actualizado') +
                    (choques.length ? ', con el solape registrado en la nota.' : '.') };
}


/* ════════════════════════════════════════════════════════════════
   APROBACIÓN POR EXCEPCIÓN (versión 8)

   Lo que hace un editor se aplica solo si pasa TODAS estas pruebas:
   - el ambiente existe, no está bloqueado en esas fechas y el grupo
     cabe en su aforo;
   - no choca con nada en el ambiente (ocupación real, desde hoy);
   - el docente no está dictando en otro lado a esa hora;
   - no cambia la duración de la clase (eso mueve horas de pago);
   - no es un dictado nuevo, un cambio de docente, de vigencia o de
     alumnos, ni una baja (todo eso mueve carga en Gestión Docente);
   - si es carrera o GT15, no es permanente: dura como mucho
     DIAS_PERMANENTE días (28 por defecto). Lo permanente toca el
     horario madre y Q10.
   Si falla una sola, queda pendiente con el motivo por escrito.
   ════════════════════════════════════════════════════════════════ */

var TIPOS_CARRERA = ['CARRERA', 'GT15'];
var SIN_DOCENTE = ['', 'ASINCRONO', 'PROGRAMA EXTERNO', 'X', '—', '-', 'POR DEFINIR', 'ACADEMICO SUPERVISA', 'S/D'];
var DOCENTES_URL_DEF = 'https://script.google.com/macros/s/AKfycbzkplgTRwDINbHUvHTWr0xXnz189-qwAbcfzdQb8lmUTuw6Nmyiy06FB5ArTeTWUsez/exec?a=docentes';

var ETIQUETA_CAMPO = {
  docente: 'el docente', curso: 'el curso', seccion: 'la sección', tipo: 'el tipo',
  alumnos: 'los alumnos', vigDesde: 'el inicio', vigHasta: 'el término', fechas: 'las fechas',
  cadencia: 'la cadencia', activo: 'el estado activo', programa: 'el programa', grupo: 'el grupo',
  crucePermitido: 'el solape autorizado'
};

function hoyIso_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }

function normNombre_(s) {
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toUpperCase().replace(/\s+/g, ' ').trim();
}

function diasEntre_(a, b) {
  var pa = a.split('-'), pb = b.split('-');
  return Math.round((new Date(+pb[0], +pb[1] - 1, +pb[2]) - new Date(+pa[0], +pa[1] - 1, +pa[2])) / 86400000);
}

function duracion_(ini, fin) { return minutos_(horaTexto_(fin)) - minutos_(horaTexto_(ini)); }

function turno_(ini) {
  var m = minutos_(horaTexto_(ini));
  return m < 13 * 60 ? 'Mañana' : (m < 17 * 60 + 50 ? 'Tarde' : 'Noche');
}

/* Lo que hace falta para revisar: base, cambios, catálogo y ocupación real. */
function contexto_() {
  var prog = leerProg_(), cambios = leerCambios_();
  return { prog: prog, cambios: cambios, ambs: leerAmbientes_(), ocup: OCUP.ocupaciones(prog, cambios) };
}

/* ¿Rige el bloqueo del ambiente para este dictado? Devuelve el texto del
   bloqueo, o '' si no rige. Usa las columnas bloqueoTurnos / bloqueoDesde /
   bloqueoHasta si existen; si no, lee el texto («hasta el 31-10-2026»,
   «mañana y tarde»). Ante la duda, rige: mejor que lo veas tú. */
function bloqueo_(amb, o) {
  if (!amb || !amb.bloqueada) return '';
  var desde = o.vigDesde || hoyIso_(), hasta = o.vigHasta || '9999-12-31';
  var t = String(amb.bloqueada);
  if (amb.bloqueoHasta && amb.bloqueoHasta < desde) return '';
  if (amb.bloqueoDesde && amb.bloqueoDesde > hasta) return '';
  if (amb.bloqueoTurnos) {
    var turnos = String(amb.bloqueoTurnos).split(',').map(function (s) { return s.trim(); });
    return turnos.indexOf(turno_(o.ini)) >= 0 ? t : '';
  }
  var m = /hasta(?: el)?\s+(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/i.exec(t);
  if (m) {
    var f = m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
    if (f < desde) return '';
  }
  var fm = /hasta\s+(?:el\s+)?fin(?:al(?:es)?)?\s+de\s+([a-záéíóú]+)(?:\s+(?:de\s+)?(\d{4}))?/i.exec(t);
  var MESES_ = { ENERO: 1, FEBRERO: 2, MARZO: 3, ABRIL: 4, MAYO: 5, JUNIO: 6, JULIO: 7, AGOSTO: 8, SETIEMBRE: 9, SEPTIEMBRE: 9, OCTUBRE: 10, NOVIEMBRE: 11, DICIEMBRE: 12 };
  if (fm && MESES_[normNombre_(fm[1])]) {
    var y = +(fm[2] || desde.slice(0, 4)), mm = MESES_[normNombre_(fm[1])];
    var ult = Utilities.formatDate(new Date(y, mm, 0, 12), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (ult < desde) return '';
  }
  if (/ma[nñ]ana y tarde/i.test(t) && turno_(o.ini) === 'Noche') return '';
  return t;
}

/* Cruces del docente: misma persona, mismo día, horas que se pisan y al
   menos una fecha de clase en común desde hoy. */
function crucesDocente_(o, lista, idPropio) {
  var d = normNombre_(o.docente);
  if (SIN_DOCENTE.indexOf(d) >= 0) return [];
  var a = OCUP.min(o.ini), b = OCUP.min(o.fin);
  if (isNaN(a) || isNaN(b)) return [];
  var r = [];
  lista.forEach(function (q) {
    if (idPropio && q.id === String(idPropio)) return;
    if (normNombre_(q.docente) !== d || q.dia !== o.dia) return;
    var qa = OCUP.min(q.ini), qb = OCUP.min(q.fin);
    if (isNaN(qa) || isNaN(qb) || !(qa < b && a < qb)) return;
    var f = OCUP.fechaComun(o, q);
    if (f) r.push(q.etiqueta + ' en el ' + q.aula + ', desde el ' + f.slice(8, 10) + '-' + f.slice(5, 7));
  });
  return r;
}

/* Las pruebas del ambiente y del docente, comunes a todo. aulaOriginal:
   si el dictado se queda en su mismo ambiente, el aforo no se vuelve a
   juzgar (si ya estaba sobre aforo, eso ya lo habías autorizado). */
function revisarEspacio_(o, idPropio, ctx, razones, aulaOriginal) {
  o.activo = o.activo || 'SÍ';
  if (String(o.activo).toUpperCase() === 'NO') return;
  if (String(o.aula).toUpperCase() !== 'VIRTUAL') {
    var amb = ctx.ambs.filter(function (x) { return x.aula === String(o.aula); })[0];
    if (!amb) {
      razones.push('El ambiente «' + o.aula + '» no está en el catálogo.');
    } else {
      var b = bloqueo_(amb, o);
      if (b) razones.push('El ' + o.aula + ' está bloqueado: ' + b);
      var n = parseInt(o.alumnos, 10), af = parseInt(amb.aforo, 10);
      if (String(aulaOriginal || '') !== String(o.aula) && !isNaN(n) && !isNaN(af) && n > af) razones.push('Sobre aforo: ' + n + ' alumnos en un ambiente de ' + af + '.');
    }
    var ch = OCUP.crucesEn(o, ctx.ocup, idPropio);
    if (ch.length) razones.push('Choca en el ' + o.aula + ' con ' + ch.join('; ') + '.');
  }
  var dch = crucesDocente_(o, ctx.ocup, idPropio);
  if (dch.length) razones.push('El docente ya dicta a esa hora: ' + dch.join('; ') + '.');
}

function esLargo_(desde, hasta, alcance) {
  if (!desde || alcance === 'ciclo') return true;
  return diasEntre_(desde, hasta || desde) > Number(prop_('DIAS_PERMANENTE', '28'));
}

function mayor_(a, b) { return !a ? (b || '') : (!b ? a : (a > b ? a : b)); }
function menor_(a, b) { return !a ? (b || '') : (!b ? a : (a < b ? a : b)); }

/* ---- el padrón de Gestión Docente ----
   Solo nombres, igual que el que usa el portal. Se guarda 15 minutos para
   no consultarlo en cada cambio. Si no responde, no se frena nada. */
function padron_() {
  var cache = CacheService.getScriptCache();
  var v = cache.get('padron');
  if (v) { try { return JSON.parse(v); } catch (e) {} }
  try {
    var r = UrlFetchApp.fetch(prop_('DOCENTES_URL', DOCENTES_URL_DEF), { muteHttpExceptions: true, followRedirects: true });
    var j = JSON.parse(r.getContentText());
    if (j && j.ok && j.docentes) {
      var l = j.docentes.map(normNombre_);
      cache.put('padron', JSON.stringify(l), 900);
      return l;
    }
  } catch (err) { Logger.log('Padrón no disponible: %s', err); }
  return null;
}

function avisoPadron_(nombre, razones) {
  var n = normNombre_(nombre);
  if (SIN_DOCENTE.indexOf(n) >= 0) return;
  var l = padron_();
  if (l && l.indexOf(n) < 0) {
    razones.push('«' + nombre + '» no figura con ese nombre en el padrón de Gestión Docente: dalo de alta o corrige el nombre antes de aprobar, para que su carga se costee.');
  }
}

/* ---- reubicaciones, bajas de clases y eventos (Reprogramar) ---- */
function evaluarPropuesta_(c) {
  var r = [];
  if (c.accion === 'desactivar') {
    r.push('Es una baja de clases: cambia las horas que se le pagan al docente en Gestión Docente.');
    return { razones: r };
  }
  if (c.accion !== 'evento' && c.accion !== 'mover') {
    r.push('Este tipo de solicitud siempre lo revisa Dirección Académica.');
    return { razones: r };
  }
  var ctx = contexto_();

  if (c.accion === 'evento') {
    revisarEspacio_({ aula: c.aula, dia: c.dia, ini: horaTexto_(c.ini), fin: horaTexto_(c.fin),
                      vigDesde: c.desde, vigHasta: c.hasta, alumnos: c.alumnos, docente: '' }, '', ctx, r);
    return { razones: r };
  }

  /* mover: un dictado de la base o un evento ya aprobado */
  var base = null, did = String(c.did || '');
  if (did.indexOf('ev-') === 0) {
    var ev = ctx.cambios.filter(function (x) { return 'ev-' + x.cid === did; })[0];
    if (ev) base = { aula: ev.aula, dia: ev.dia, ini: horaTexto_(ev.ini), fin: horaTexto_(ev.fin), tipo: 'EVENTO',
                     vigDesde: ev.desde, vigHasta: ev.hasta, alumnos: ev.alumnos, docente: '' };
  } else {
    base = ctx.prog.filter(function (q) { return String(q.id) === did; })[0];
  }
  if (!base) { r.push('No se encontró el dictado original en la base.'); return { razones: r }; }

  var tipo = String(base.tipo || '').toUpperCase();
  var ni = horaTexto_(c.ini) || base.ini, nf = horaTexto_(c.fin) || base.fin;
  if (duracion_(ni, nf) !== duracion_(base.ini, base.fin)) {
    r.push('Cambia la duración de la clase (' + base.ini + '-' + base.fin + ' → ' + ni + '-' + nf + '): mueve horas de pago en Gestión Docente.');
  }
  if (TIPOS_CARRERA.indexOf(tipo) >= 0 && esLargo_(c.desde, c.hasta, c.alcance)) {
    r.push('Es un cambio ' + (c.desde ? 'de más de ' + prop_('DIAS_PERMANENTE', '28') + ' días' : 'de todo el ciclo') +
           ' en un dictado de ' + tipo + ': toca el horario madre y Q10.');
  }
  var o = { aula: c.aula || base.aula, dia: c.dia || base.dia, ini: ni, fin: nf,
            vigDesde: mayor_(c.desde, base.vigDesde), vigHasta: menor_(c.hasta, base.vigHasta),
            fechas: base.fechas || '', docente: base.docente || '', alumnos: base.alumnos || '' };
  if (!(o.vigDesde && o.vigHasta && o.vigDesde > o.vigHasta)) revisarEspacio_(o, did, ctx, r, base.aula);
  return { razones: r };
}

/* ---- crear, editar o dar de baja desde Horarios ---- */

function valorCampo_(k, v) {
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Array]') return v.join(', ');
  var s = String(v).trim();
  if (k === 'ini' || k === 'fin') return horaTexto_(s);
  if (k === 'activo') return s.toUpperCase().replace('SI', 'SÍ');
  return s;
}

function camposCambiados_(antes, o) {
  return CAMPOS_EDITABLES.filter(function (k) {
    if (!(k in o)) return false;
    return valorCampo_(k, antes[k]) !== valorCampo_(k, o[k]);
  });
}

/* Lo que se guarda para poder revertir: los campos editables tal como
   estaban antes del cambio. */
function copiaEditable_(x) {
  var r = {};
  CAMPOS_EDITABLES.forEach(function (k) { if (x[k] !== undefined) r[k] = x[k]; });
  return r;
}

function firmaBase_(x) {
  return [x.accion, x.did, x.aula, x.dia, horaTexto_(x.ini), horaTexto_(x.fin), x.docente || '', x.titulo || ''].join('|');
}

function filaBase_(p, o, id, razones, auto, antes) {
  return {
    cid: uid_(), ts: ahora_(), por: p.nombre || p.correo || '', correo: p.correo || '',
    did: 'base:' + (id || 'nuevo'), accion: p.accion,
    aula: o.aula || '', dia: o.dia || '', ini: o.ini || '', fin: o.fin || '',
    motivo: p.motivo || '', resumen: p.resumen || ((o.seccion || '') + ' — ' + (o.dia || '') + ' ' + (o.horario || '') + ' en ' + (o.aula || '')),
    estado: auto ? 'aprobado' : 'pendiente', resueltoPor: auto ? AUTO : '', tsResuelto: auto ? ahora_() : '',
    titulo: o.curso || '', docente: o.docente || '', alumnos: o.alumnos || '', area: '',
    alcance: '', desde: '', hasta: '', clases: '',
    nivel: auto ? 'auto' : 'revision', razon: (razones || []).join(' '),
    datos: JSON.stringify({ accion: p.accion, id: id || '', dictado: p.dictado || null, nota: p.nota || '', resumen: p.resumen || '' }),
    antes: antes ? JSON.stringify(antes) : ''
  };
}

function pendienteBase_(p, o, id, razones, antes) {
  var c = filaBase_(p, o, id, razones, false, antes);
  var ya = leerCambios_().filter(function (x) { return x.estado === 'pendiente' && firmaBase_(x) === firmaBase_(c); })[0];
  if (ya) return { ok: false, error: 'Esta misma solicitud ya está esperando a Dirección Académica (registrada por ' + (ya.por || 'alguien') + ').' };
  agregarFila_(c);
  if (!p.lote) avisarSolicitud_(c);   /* una importación avisa una sola vez, al final */
  return { ok: true, pendiente: true, razon: c.razon, cambios: leerCambios_(),
           mensaje: 'Quedó pendiente de Dirección Académica. ' + c.razon };
}

function baseEditor_(p) {
  var prog = leerProg_();

  if (p.accion === 'base-baja') {
    var fila = prog.filter(function (q) { return String(q.id) === String(p.id); })[0];
    if (!fila) return { ok: false, error: 'No se encontró el dictado ' + p.id + ' en la base.' };
    return pendienteBase_(p, fila, p.id, ['Dar de baja un dictado quita su carga en Gestión Docente.'], null);
  }

  var o = normalizarDictado_(p.dictado || {});
  var esNuevo = (p.accion === 'base-crear');
  var errores = validarDictado_(o, esNuevo);
  if (errores.length) return { ok: false, error: errores.join(' ') };

  var antes = null;
  if (!esNuevo) {
    antes = prog.filter(function (q) { return String(q.id) === String(p.id); })[0];
    if (!antes) return { ok: false, error: 'No se encontró el dictado ' + p.id + ' en la base.' };
  }
  var idPropio = esNuevo ? '' : String(p.id);
  var ctx = contexto_();
  var cambiados = esNuevo ? [] : camposCambiados_(antes, o);
  /* Si no cambia ni ambiente, ni día, ni hora, ni vigencia, ni fechas, el
     lugar que ocupa sigue igual: no se vuelven a juzgar cruces viejos que
     ya habías autorizado. */
  var mueve = esNuevo || cambiados.some(function (k) {
    return ['aula', 'dia', 'ini', 'fin', 'vigDesde', 'vigHasta', 'fechas', 'activo', 'docente'].indexOf(k) >= 0;
  });

  /* El cruce de ambiente se avisa primero, igual que a Dirección
     Académica: el portal pregunta si se manda igual (a revisión). */
  var choques = mueve ? OCUP.crucesEn(o, ctx.ocup, idPropio) : [];
  if (choques.length && !p.forzar) {
    return { ok: false, cruce: choques,
             error: 'Choca con ' + choques.join('; ') + ' en el ambiente ' + o.aula + ' el ' + o.dia + '.' };
  }

  var r = [];
  var tipo = String(o.tipo || (antes && antes.tipo) || '').toUpperCase();
  if (esNuevo) {
    if (tipo !== 'EVENTO') {
      r.push('Es un dictado nuevo de ' + tipo + ': suma carga docente y entra al tablero de Gestión Docente.');
      if (o.docente) avisoPadron_(o.docente, r);
    }
  } else {
    var sensibles = cambiados.filter(function (k) { return ['aula', 'dia', 'ini', 'fin', 'nota'].indexOf(k) < 0; });
    if (sensibles.length) {
      var otros = sensibles.filter(function (k) { return k !== 'docente'; });
      var t = otros.length ? 'Cambia ' + otros.map(function (k) { return ETIQUETA_CAMPO[k] || k; }).join(', ') + '.' : '';
      if (sensibles.indexOf('docente') >= 0) {
        t += (t ? ' ' : '') + 'El docente pasa de «' + (antes.docente || '—') + '» a «' + (o.docente || '—') + '»: mueve la carga en Gestión Docente.';
      } else if (sensibles.some(function (k) { return ['alumnos', 'vigDesde', 'vigHasta', 'fechas', 'activo'].indexOf(k) >= 0; })) {
        t += ' Cambia las horas o los alumnos que ve Gestión Docente.';
      }
      r.push(t);
      if (sensibles.indexOf('docente') >= 0 && o.docente) avisoPadron_(o.docente, r);
    }
    if (duracion_(o.ini, o.fin) !== duracion_(antes.ini, antes.fin)) {
      r.push('Cambia la duración de la clase (' + antes.ini + '-' + antes.fin + ' → ' + o.ini + '-' + o.fin + '): mueve horas de pago.');
    }
    var traslada = cambiados.some(function (k) { return ['aula', 'dia', 'ini', 'fin'].indexOf(k) >= 0; });
    if (traslada && TIPOS_CARRERA.indexOf(tipo) >= 0) {
      r.push('Editar la base de un dictado de ' + tipo + ' es un cambio permanente: toca el horario madre y Q10. Si es por unas semanas, usa Reprogramar con fechas y se aplica solo.');
    }
  }
  if (mueve) revisarEspacio_(o, idPropio, ctx, r, antes ? antes.aula : '');

  if (r.length) return pendienteBase_(p, o, idPropio, r, antes ? copiaEditable_(antes) : null);

  /* Pasó todo: se escribe y queda en Cambios como aplicado solo. */
  var w = escribirBase_(p);
  if (!w.ok) return w;
  agregarFila_(filaBase_(p, o, w.id || idPropio, [], true, antes ? copiaEditable_(antes) : null));
  w.cambios = leerCambios_();
  w.mensaje += ' Se aplicó sin esperar aprobación; Dirección Académica lo verá en su resumen.';
  return w;
}

/* ---- revertir un cambio de la base ---- */
function revertirBase_(c, quien) {
  var d = {}, antes = null;
  try { d = JSON.parse(c.datos || '{}'); } catch (e) {}
  try { antes = c.antes ? JSON.parse(c.antes) : null; } catch (e) {}
  var nota = 'Revertido por Dirección Académica el ' + hoyIso_() + '.';

  if (c.accion === 'base-crear') {
    if (!d.id) return { ok: false, error: 'No se sabe qué dictado se creó; dalo de baja desde Horarios.' };
    return escribirBase_({ accion: 'base-baja', id: d.id, nota: nota, correo: quien, resumen: 'Reversión: ' + (c.resumen || '') });
  }
  if (c.accion === 'base-editar') {
    if (!antes || !d.id) return { ok: false, error: 'No quedó copia de cómo estaba; corrígelo a mano desde Horarios.' };
    return escribirBase_({ accion: 'base-editar', id: d.id, dictado: antes, forzar: true, correo: quien,
                           resumen: 'Reversión: ' + (c.resumen || '') });
  }
  if (c.accion === 'base-baja') {
    if (!d.id) return { ok: false, error: 'No se sabe qué dictado se dio de baja.' };
    return escribirBase_({ accion: 'base-editar', id: d.id, dictado: { activo: 'SÍ' }, forzar: true, correo: quien,
                           resumen: 'Reversión de baja: ' + (c.resumen || '') });
  }
  return { ok: false, error: 'Ese cambio no se puede revertir desde aquí.' };
}

/* ---- prueba desde el editor ----
   Muestra en el registro qué haría el servidor con una reubicación de un
   solo día, sin escribir nada. Cambia el id por uno real de PROGRAMACION. */
function probarRevision() {
  var prog = leerProg_();
  var q = prog.filter(function (x) { return x.activo !== 'NO' && x.aula !== 'VIRTUAL'; })[0];
  var hoy = hoyIso_();
  var c = { accion: 'mover', did: q.id, aula: q.aula, dia: q.dia, ini: q.ini, fin: q.fin, desde: hoy, hasta: hoy, alcance: 'una' };
  Logger.log('Dictado de prueba: %s', JSON.stringify(q));
  Logger.log('Razones (vacío = se aplicaría solo): %s', JSON.stringify(evaluarPropuesta_(c).razones));
  Logger.log('Padrón leído: %s nombres', (padron_() || []).length);
}



/* ════════════════════════════════════════════════════════════════
   IMPORTACIÓN DEL PLAN DE INICIO (versión 8)

   El portal arma los dictados de una campaña y los crea uno por uno con
   el mismo «lote». Si los manda un editor, quedan pendientes y aquí se
   manda UN correo con todo el lote. Al aprobar o rechazar el lote, el
   portal vuelve a llamar con resuelto = true y quien lo envió recibe
   un solo resumen.

   leerPlan_ lee la pestaña «Plan inicio AAAA» de la hoja Nueva Oferta
   Académica para que no haga falta descargar el CSV. Configurar en
   Propiedades del script:
     PLAN_INICIO_URL   el enlace de la hoja
     PLAN_PESTANA      (opcional) el nombre exacto de la pestaña; si no
                       está, toma la «Plan inicio» del año más alto.
   Solo devuelve columnas de programación: nada de precios ni montos.
   ════════════════════════════════════════════════════════════════ */

var PLAN_COLUMNAS = ['CAMPANA', 'MODALIDAD', 'PROGRAMA', 'TIPO DE PROGRAMA', 'HORAS', 'SESIONES', 'FRECUENCIA',
                     'FECHA DE INICIO', 'FECHA DE TERMINO', 'DIAS', 'HORARIO', 'DOCENTE', 'GRUPOS', 'META Q'];

function leerPlan_(p) {
  var url = prop_('PLAN_INICIO_URL', '');
  if (!url) return { ok: false, error: 'Falta configurar la hoja: en Propiedades del script agrega PLAN_INICIO_URL con el enlace de «Nueva Oferta Académica». Mientras tanto, sube el CSV.' };
  var ss;
  try { ss = SpreadsheetApp.openByUrl(url); }
  catch (err) { return { ok: false, error: 'No pude abrir la hoja del Plan de inicio: ' + err }; }
  var hojas = ss.getSheets().filter(function (h) { return /plan\s*(de\s*)?inicio/i.test(h.getName()); });
  var nombre = p.pestana || prop_('PLAN_PESTANA', '');
  var h = nombre ? ss.getSheetByName(nombre) : hojas.sort(function (a, b) {
    var ya = +((a.getName().match(/\d{4}/) || ['0'])[0]), yb = +((b.getName().match(/\d{4}/) || ['0'])[0]);
    return yb - ya;
  })[0];
  if (!h) return { ok: false, error: 'No encontré la pestaña del Plan de inicio' + (nombre ? ' «' + nombre + '»' : '') + '.' };
  var v = h.getDataRange().getDisplayValues();
  var cab = -1;
  for (var i = 0; i < v.length; i++) {
    var u = v[i].map(normNombre_);
    if (u.indexOf('CAMPANA') >= 0 && u.indexOf('PROGRAMA') >= 0) { cab = i; break; }
  }
  if (cab < 0) return { ok: false, error: 'La pestaña «' + h.getName() + '» no tiene la fila de encabezados (CAMPAÑA, PROGRAMA…).' };
  var hdr = v[cab].map(normNombre_);
  var idx = PLAN_COLUMNAS.map(function (c) { return hdr.indexOf(c); });
  var filas = v.slice(cab).map(function (f) { return idx.map(function (j) { return j >= 0 ? f[j] : ''; }); })
               .filter(function (f) { return f.some(function (x) { return String(x).trim(); }); });
  filas[0] = PLAN_COLUMNAS.map(function (c, k) { return idx[k] >= 0 ? v[cab][idx[k]] : c; });
  return { ok: true, pestana: h.getName(), pestanas: hojas.map(function (x) { return x.getName(); }), filas: filas };
}

function avisarLote_(lote, resuelto) {
  try {
    if (!lote) return { ok: false, error: 'Falta el lote.' };
    var filas = leerCambios_().filter(function (c) { return c.motivo === PREFIJO_LOTE + lote; });
    if (!filas.length) return { ok: true, enviados: 0 };
    var pend = filas.filter(function (c) { return c.estado === 'pendiente'; });
    var hechas = filas.filter(function (c) { return c.estado !== 'pendiente'; });
    var lista = function (arr) {
      return '<ul style="font-size:14px;padding-left:20px;line-height:1.6">' + arr.map(function (c) {
        return '<li>' + escapaHtml_(c.resumen) + '<br><span style="font-size:12.5px;color:#888">' + escapaHtml_(c.docente || 'sin docente') +
               (resuelto ? ' · ' + escapaHtml_(c.estado) : '') + '</span>' +
               (!resuelto && c.razon ? '<br><span style="font-size:12.5px;color:#c0392b">' + escapaHtml_(c.razon) + '</span>' : '') + '</li>';
      }).join('') + '</ul>';
    };
    var cab = function (titulo) {
      return '<div style="font-family:Helvetica,Arial,sans-serif;max-width:600px;color:#1a1a1a">' +
        '<div style="font-size:12px;letter-spacing:.08em;text-transform:uppercase;color:#888">Auditoría de Ambientes 2026-II</div>' +
        '<div style="font-size:19px;font-weight:600;margin:3px 0 14px">' + escapaHtml_(titulo) + '</div>';
    };
    var boton = '<p style="margin:22px 0 0"><a href="' + portal_() + '" style="background:#1a1a1a;color:#fff;text-decoration:none;padding:10px 18px;border-radius:6px;font-size:14px;display:inline-block">Abrir el portal</a></p></div>';
    if (!resuelto) {
      if (!pend.length) return { ok: true, enviados: 0 };
      if (prop_('AVISOS_INSTANTANEOS', 'SI').toUpperCase() === 'NO') return { ok: true, enviados: 0 };
      MailApp.sendEmail({
        to: destinatarios_(), name: 'Auditoría de Ambientes',
        subject: '[Ambientes] Lote por aprobar · ' + pend.length + ' bloque(s) · ' + corto_(lote, 60),
        htmlBody: cab(pend.length + ' bloque(s) nuevos esperan tu aprobación') +
          '<p style="font-size:14px">Los envió ' + escapaHtml_(pend[0].por || '') + ' desde la importación del Plan de inicio (' + escapaHtml_(lote) + '). En la pestaña Cambios los apruebas juntos con «Aprobar el lote».</p>' +
          lista(pend) + boton
      });
    } else {
      var para = {};
      hechas.forEach(function (c) { if (c.correo && c.correo.indexOf('@') > 0) para[c.correo.toLowerCase()] = 1; });
      var dest = Object.keys(para);
      if (!dest.length) return { ok: true, enviados: 0 };
      MailApp.sendEmail({
        to: dest.join(','), name: 'Auditoría de Ambientes',
        subject: '[Ambientes] Lote resuelto · ' + corto_(lote, 60),
        htmlBody: cab('Dirección Académica resolvió el lote') + lista(hechas) + boton
      });
    }
    return { ok: true, enviados: 1 };
  } catch (err) {
    Logger.log('Aviso de lote no enviado: %s', err);
    return { ok: true, enviados: 0 };
  }
}
