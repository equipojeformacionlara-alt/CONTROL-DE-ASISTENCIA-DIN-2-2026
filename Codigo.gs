/**
 * CONTROL DE ASISTENCIA DIN – Reporte de clases / Microclases conversatorio
 * ---------------------------------------------------------------------------
 * Script VINCULADO A LA HOJA DE RESPUESTAS del formulario
 * (Formulario > Respuestas > Vincular con Hojas > Extensiones > Apps Script).
 *
 * Qué hace:
 *  1. Mantiene la lista de docentes en una hoja ("Docentes") y la sincroniza con el formulario.
 *  2. Corrige validaciones del formulario (campos numéricos, "no aparece en el listado" opcional, etc.).
 *  3. Valida cada envío: horas, fecha, duplicados, docente, foto, totales.
 *  4. Calcula columnas: total, % asistencia, duración, cédula, condición, estado, alertas.
 *  5. Envía confirmación al docente y alerta al coordinador si hay algo por revisar.
 *  6. Genera hoja "Resumen" (por docente, núcleo, localidad, semestre) y un correo semanal
 *     con los docentes que NO reportaron.
 *
 * Instalación: pegar en el editor, ajustar CONFIG, recargar la hoja y usar el menú
 * "Control DIN" > "1. Instalar / actualizar todo".
 */

// ============================== CONFIGURACIÓN ==============================
const CONFIG = {
  FORM_ID: 'https://script.google.com/macros/s/AKfycbxuUDTGZxL3sgybwiB2SByxMOgTP-YqrTovHBF7GwREifWCKWSL6FiA2w3x6X1DcYpb/exec',          // <-- CAMBIAR (está en la URL del formulario)
  HOJA_RESPUESTAS: '',                 // vacío = detecta "Respuestas de formulario"
  HOJA_DOCENTES: 'Docentes',
  HOJA_RESUMEN: 'Resumen',
  COORDINADOR_EMAIL: 'coordinador@ejemplo.com',   // <-- CAMBIAR
  ADMINS: [],                          // correos adicionales que pueden gestionar docentes en la interfaz web
  DIAS_MAX_ANTIGUEDAD: 30,             // fechas más viejas se marcan para revisión
  DIAS_SIN_REPORTE: 7,                 // ventana del correo semanal
  TZ: 'America/Caracas',
  ENVIAR_CONFIRMACION: true,
  COLOR_REVISAR: '#fde2e1',
  COLOR_OK: '#ffffff'
};

// Fragmentos (sin tildes, en mayúsculas) para ubicar columnas por el título de la pregunta
const CAMPOS = {
  EMAIL: ['CORREO', 'EMAIL'],
  DOCENTE: ['NOMBRE Y APELLIDO'],
  OTRO: ['SI NO APARECE'],
  NUCLEO: ['NUCLEO'],
  LOCALIDAD: ['LOCALIDAD'],
  FECHA: ['FECHA'],
  INICIO: ['HORA DE INICIO'],
  SALIDA: ['HORA DE SALIDA'],
  SEMESTRE: ['SEMESTRE'],
  CARRERA: ['CODIGO DE CARRERA'],
  SECCION: ['SECCION'],
  ASIST: ['ESTUDIANTES ASISTENTES'],
  FALT: ['ESTUDIANTES FALTANTES'],
  DESER: ['DESERCION'],
  TEMA: ['TEMA A TRATAR'],
  UNIDAD: ['UNIDAD'],
  OBS: ['OBSERVACION'],
  IMPACTO: ['IMPACTO'],
  FOTO: ['RESENA FOTOGRAFICA']
};

// Columnas calculadas que el script agrega al final de la hoja de respuestas
const CALC = {
  TOTAL: 'Total estudiantes',
  PCT: '% Asistencia',
  DUR: 'Duración (min)',
  CED: 'Cédula (normalizada)',
  COND: 'Condición',
  ESTADO: 'Estado',
  ALERTA: 'Alertas'
};

// Lista inicial de docentes (se vuelca una sola vez a la hoja "Docentes")
const SEMILLA_DOCENTES = [
  // Ejemplos ficticios. Los docentes reales se cargan en la hoja "Docentes", no en el repositorio.
  'Nombre Apellido 12345678 (TV)',
  'Otro Docente 87654321 (FIJO)'
];

// ================================ MENÚ ====================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Control DIN')
    .addItem('1. Instalar / actualizar todo', 'instalarTodo')
    .addItem('2. Sincronizar docentes con el formulario', 'actualizarFormulario')
    .addItem('3. Revalidar todas las respuestas', 'revalidarTodo')
    .addItem('4. Generar hoja Resumen', 'generarResumen')
    .addItem('5. Enviar resumen semanal ahora', 'enviarResumenSemanal')
    .addToUi();
}

function instalarTodo() {
  crearHojaDocentes_();
  prepararHoja_(hojaRespuestas_());
  actualizarFormulario();
  instalarTriggers_();
  SpreadsheetApp.getUi().alert('Listo: hoja Docentes, columnas calculadas, formulario y disparadores configurados.');
}

function instalarTriggers_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const f = t.getHandlerFunction();
    if (f === 'alEnviarFormulario' || f === 'enviarResumenSemanal') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('alEnviarFormulario')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onFormSubmit().create();
  ScriptApp.newTrigger('enviarResumenSemanal')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).inTimezone(CONFIG.TZ).create();
}

// ========================= UTILIDADES GENERALES ===========================
function norm_(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[°º]/g, '').toUpperCase().replace(/\s+/g, ' ').trim();
}
function num_(x) {
  const n = parseInt(String(x == null ? '' : x).replace(/\D/g, ''), 10);
  return isNaN(n) ? null : n;
}
function esc_(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
function aMinutos_(s) {
  const m = String(s == null ? '' : s).trim().match(/^(\d{1,2})[:.](\d{2})(?::\d{2})?\s*([apAP])?/);
  if (!m) return null;
  let h = +m[1];
  const ap = m[3] && m[3].toLowerCase();
  if (ap === 'p' && h < 12) h += 12;
  if (ap === 'a' && h === 12) h = 0;
  return h * 60 + (+m[2]);
}
function extraerCedula_(texto) {
  const m = String(texto || '').match(/\d[\d.,]{5,}/);
  return m ? m[0].replace(/\D/g, '') : '';
}
function hojaRespuestas_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (CONFIG.HOJA_RESPUESTAS) {
    const s = ss.getSheetByName(CONFIG.HOJA_RESPUESTAS);
    if (!s) throw new Error('No existe la hoja: ' + CONFIG.HOJA_RESPUESTAS);
    return s;
  }
  return ss.getSheets().find(s => /respuestas|responses/i.test(s.getName())) || ss.getSheets()[0];
}

// ============================ HOJA DOCENTES ===============================
function crearHojaDocentes_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CONFIG.HOJA_DOCENTES);
  if (sh) return sh;
  sh = ss.insertSheet(CONFIG.HOJA_DOCENTES);
  const filas = SEMILLA_DOCENTES.map(t => {
    const m = t.match(/^(.+?)\s*(\d[\d.]*)\s*\((.+)\)\s*$/);
    return m ? [m[1].trim(), m[2].replace(/\D/g, ''), m[3].trim(), 'SI', 'Inicial'] : [t, '', '', 'SI', 'Inicial'];
  });
  sh.getRange(1, 1, 1, 5).setValues([['Nombre y apellido', 'Cédula', 'Condición', 'Activo (SI/NO)', 'Alta']])
    .setFontWeight('bold').setBackground('#d9ead3');
  sh.getRange(2, 1, filas.length, 5).setValues(filas);
  sh.getRange(2, 2, filas.length, 1).setNumberFormat('@');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 5);
  return sh;
}

function leerDocentes_() {
  const sh = crearHojaDocentes_();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues()
    .filter(r => String(r[0]).trim() && String(r[3]).trim().toUpperCase() !== 'NO')
    .map(r => ({
      nombre: String(r[0]).trim(),
      cedula: String(r[1]).replace(/\D/g, ''),
      condicion: String(r[2]).trim()
    }))
    .sort((a, b) => a.nombre.localeCompare(b.nombre, 'es'));
}

function mapaDocentes_() {
  const map = {};
  leerDocentes_().forEach(d => { if (d.cedula) map[d.cedula] = d; });
  return map;
}

// ======================= MEJORAS AL FORMULARIO ============================
function actualizarFormulario() {
  const form = FormApp.openById(CONFIG.FORM_ID);
  const docentes = leerDocentes_();
  if (!docentes.length) throw new Error('La hoja Docentes está vacía.');
  const items = form.getItems();
  const buscar = frag => items.find(it => norm_(it.getTitle()).indexOf(frag) !== -1);

  // 1) Lista de docentes sincronizada con la hoja
  const itDoc = buscar('NOMBRE Y APELLIDO');
  if (itDoc) {
    const etiquetas = docentes.map(d => d.nombre + ' ' + d.cedula + (d.condicion ? ' (' + d.condicion + ')' : ''));
    if (itDoc.getType() === FormApp.ItemType.MULTIPLE_CHOICE) {
      itDoc.asMultipleChoiceItem().setChoiceValues(etiquetas).showOtherOption(false);
    } else if (itDoc.getType() === FormApp.ItemType.LIST) {
      itDoc.asListItem().setChoiceValues(etiquetas);
    }
  }

  // 2) "Si no aparece en el listado": opcional y con validación de cédula
  const itOtro = buscar('SI NO APARECE');
  if (itOtro) {
    const v = FormApp.createTextValidation()
      .setHelpText('Escriba nombre completo y cédula (solo números). Ej.: María Pérez 12345678')
      .requireTextMatchesPattern('^(?=.*\\d{6,9}).{8,}$').build();
    itOtro.asTextItem().setRequired(false).setValidation(v);
  }

  // 3) Campos numéricos: enteros de 0 a 9999
  ['ESTUDIANTES ASISTENTES', 'ESTUDIANTES FALTANTES', 'DESERCION'].forEach(frag => {
    const it = buscar(frag);
    if (it && it.getType() === FormApp.ItemType.TEXT) {
      const v = FormApp.createTextValidation()
        .setHelpText('Ingrese solo un número entero (0 si no aplica).')
        .requireTextMatchesPattern('^\\d{1,4}$').build();
      it.asTextItem().setValidation(v).setRequired(true);
    }
  });

  // 4) Mensaje de confirmación
  form.setConfirmationMessage('Reporte registrado correctamente. Si algún dato requiere revisión, la coordinación se comunicará con usted.');
}

// ===================== PREPARACIÓN DE LA HOJA DE RESPUESTAS ================
function prepararHoja_(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  Object.keys(CALC).forEach(k => {
    if (headers.indexOf(CALC[k]) === -1) {
      sh.getRange(1, headers.length + 1).setValue(CALC[k]).setFontWeight('bold').setBackground('#d9ead3');
      headers.push(CALC[k]);
    }
  });
  const c = {};
  Object.keys(CAMPOS).forEach(k => {
    const i = headers.findIndex(h => CAMPOS[k].some(f => norm_(h).indexOf(f) !== -1));
    c[k] = i + 1; // 0 = no encontrada
  });
  Object.keys(CALC).forEach(k => { c[k] = headers.indexOf(CALC[k]) + 1; });
  return c;
}

// ========================== PROCESO AL ENVIAR ==============================
function alEnviarFormulario(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = e && e.range ? e.range.getSheet() : hojaRespuestas_();
    const row = e && e.range ? e.range.getRow() : sh.getLastRow();
    const ctx = construirContexto_(sh);
    procesarFila_(sh, row, ctx, true);
  } finally {
    lock.releaseLock();
  }
}

function revalidarTodo() {
  const sh = hojaRespuestas_();
  const ctx = construirContexto_(sh);
  for (let r = 2; r <= sh.getLastRow(); r++) procesarFila_(sh, r, ctx, false);
  SpreadsheetApp.getActiveSpreadsheet().toast('Revalidación terminada', 'Control DIN', 5);
}

function construirContexto_(sh) {
  const c = prepararHoja_(sh);
  const ctx = { c: c, docMap: mapaDocentes_(), indice: {} };
  const n = sh.getLastRow() - 1;
  if (n < 1) return ctx;
  const rango = sh.getRange(2, 1, n, sh.getLastColumn());
  const vals = rango.getValues();
  const disp = rango.getDisplayValues();
  vals.forEach((v, i) => {
    const k = claveFila_(v, disp[i], c);
    if (k) (ctx.indice[k] = ctx.indice[k] || []).push(i + 2);
  });
  return ctx;
}

function claveFila_(vals, disp, c) {
  const g = (a, k) => (c[k] ? a[c[k] - 1] : '');
  const ced = extraerCedula_(g(vals, 'DOCENTE')) || extraerCedula_(g(vals, 'OTRO'));
  if (!ced) return '';
  return [ced, g(disp, 'FECHA'), g(disp, 'INICIO'), g(disp, 'CARRERA'), g(disp, 'SECCION')].join('|');
}

function procesarFila_(sh, row, ctx, notificar) {
  const c = ctx.c;
  const lastCol = sh.getLastColumn();
  const rango = sh.getRange(row, 1, 1, lastCol);
  const vals = rango.getValues()[0];
  const disp = rango.getDisplayValues()[0];
  const v = k => (c[k] ? vals[c[k] - 1] : '');
  const d = k => (c[k] ? disp[c[k] - 1] : '');
  const alertas = [];

  // --- Docente ---
  const sel = String(v('DOCENTE') || '').trim();
  const otro = String(v('OTRO') || '').trim();
  const texto = sel || otro;
  const cedula = extraerCedula_(sel) || extraerCedula_(otro);
  const enLista = cedula && ctx.docMap[cedula];
  let condicion = enLista ? enLista.condicion : '';
  if (!condicion) {
    const m = texto.match(/\(([^)]+)\)/);
    condicion = m ? m[1].trim().toUpperCase() : '';
  }
  if (!texto) alertas.push('Sin docente');
  else if (!cedula) alertas.push('Docente sin cédula válida');
  else if (!enLista) alertas.push('Docente NO registrado en la hoja Docentes');
  if (sel && otro) alertas.push('Eligió docente de la lista y además escribió otro');

  // --- Fecha ---
  const fecha = v('FECHA');
  if (fecha instanceof Date) {
    const hoy = new Date();
    const dias = Math.floor((hoy - fecha) / 86400000);
    if (fecha.getTime() > hoy.getTime() + 86400000) alertas.push('Fecha futura');
    else if (dias > CONFIG.DIAS_MAX_ANTIGUEDAD) alertas.push('Fecha con más de ' + CONFIG.DIAS_MAX_ANTIGUEDAD + ' días de antigüedad');
  } else {
    alertas.push('Fecha inválida');
  }

  // --- Horas ---
  const ini = aMinutos_(d('INICIO'));
  const fin = aMinutos_(d('SALIDA'));
  let dur = '';
  if (ini == null || fin == null) alertas.push('Hora inválida');
  else if (fin <= ini) alertas.push('Hora de salida menor o igual a la de inicio');
  else {
    dur = fin - ini;
    if (dur < 20) alertas.push('Duración menor a 20 min');
    if (dur > 360) alertas.push('Duración mayor a 6 horas');
  }

  // --- Estudiantes ---
  const asist = num_(v('ASIST'));
  const falt = num_(v('FALT'));
  const deser = num_(v('DESER'));
  let total = '', pct = '';
  if (asist == null || falt == null) alertas.push('Cantidades de estudiantes inválidas');
  else {
    total = asist + falt;
    if (total === 0) alertas.push('Total de estudiantes = 0');
    else pct = asist / total;
  }
  if (deser != null && falt != null && deser > falt) alertas.push('Deserción mayor a inasistentes');

  // --- Foto ---
  if (c.FOTO && !String(v('FOTO')).trim()) alertas.push('Sin reseña fotográfica');

  // --- Duplicados ---
  const clave = claveFila_(vals, disp, c);
  if (clave && ctx.indice[clave] && ctx.indice[clave][0] !== row) {
    alertas.push('Posible duplicado de la fila ' + ctx.indice[clave][0]);
  }

  // --- Escritura ---
  const estado = alertas.length ? 'REVISAR' : 'OK';
  sh.getRange(row, c.TOTAL).setValue(total);
  sh.getRange(row, c.PCT).setValue(pct).setNumberFormat('0.0%');
  sh.getRange(row, c.DUR).setValue(dur);
  sh.getRange(row, c.CED).setNumberFormat('@').setValue(cedula);
  sh.getRange(row, c.COND).setValue(condicion);
  sh.getRange(row, c.ESTADO).setValue(estado);
  sh.getRange(row, c.ALERTA).setValue(alertas.join(' | '));
  sh.getRange(row, 1, 1, c.ALERTA).setBackground(estado === 'OK' ? CONFIG.COLOR_OK : CONFIG.COLOR_REVISAR);

  if (notificar) {
    const info = {
      docente: enLista ? enLista.nombre : texto, fecha: d('FECHA'), inicio: d('INICIO'), salida: d('SALIDA'),
      carrera: d('CARRERA'), seccion: d('SECCION'), asist: asist, falt: falt, tema: v('TEMA'), unidad: v('UNIDAD'),
      estado: estado, alertas: alertas
    };
    try {
      if (CONFIG.ENVIAR_CONFIRMACION) enviarConfirmacion_(String(v('EMAIL') || '').trim(), info);
      if (estado === 'REVISAR') alertarCoordinador_(sh, row, info);
    } catch (err) {
      console.error('Error enviando correo: ' + err);
    }
  }
}

// ============================== CORREOS ===================================
function enviarConfirmacion_(email, i) {
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return;
  const html = '<p>Su reporte de clase fue recibido.</p><ul>' +
    '<li><b>Docente:</b> ' + esc_(i.docente) + '</li>' +
    '<li><b>Fecha:</b> ' + esc_(i.fecha) + ' (' + esc_(i.inicio) + ' – ' + esc_(i.salida) + ')</li>' +
    '<li><b>Carrera / Sección:</b> ' + esc_(i.carrera) + ' / ' + esc_(i.seccion) + '</li>' +
    '<li><b>Asistentes / Faltantes:</b> ' + esc_(i.asist) + ' / ' + esc_(i.falt) + '</li>' +
    '<li><b>Tema:</b> ' + esc_(i.tema) + ' (' + esc_(i.unidad) + ')</li></ul>' +
    (i.estado === 'REVISAR'
      ? '<p><b>Atención:</b> el sistema detectó puntos por revisar: ' + esc_(i.alertas.join('; ')) + '.</p>'
      : '<p>Validación automática: OK.</p>');
  MailApp.sendEmail({ to: email, subject: 'Confirmación de reporte de clase DIN', htmlBody: html });
}

function alertarCoordinador_(sh, row, i) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const url = ss.getUrl() + '#gid=' + sh.getSheetId() + '&range=A' + row;
  MailApp.sendEmail({
    to: CONFIG.COORDINADOR_EMAIL,
    subject: 'Reporte DIN por revisar – ' + i.docente,
    htmlBody: '<p>Fila <a href="' + url + '">' + row + '</a> con alertas:</p><ul>' +
      i.alertas.map(a => '<li>' + esc_(a) + '</li>').join('') + '</ul>'
  });
}

// ============================ RESUMEN / REPORTES ==========================
function leerRegistros_(sh) {
  const c = prepararHoja_(sh);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const vals = sh.getRange(2, 1, n, sh.getLastColumn()).getValues();
  const g = (r, k) => (c[k] ? r[c[k] - 1] : '');
  return vals.filter(r => g(r, 'FECHA') instanceof Date).map(r => ({
    docente: String(g(r, 'DOCENTE') || g(r, 'OTRO')).replace(/\s*\d[\d.,]{5,}.*$/, '').trim(),
    ced: String(g(r, 'CED')),
    cond: g(r, 'COND'),
    nucleo: g(r, 'NUCLEO'),
    localidad: g(r, 'LOCALIDAD'),
    semestre: g(r, 'SEMESTRE'),
    fecha: g(r, 'FECHA'),
    asist: num_(g(r, 'ASIST')) || 0,
    falt: num_(g(r, 'FALT')) || 0,
    deser: num_(g(r, 'DESER')) || 0,
    dur: Number(g(r, 'DUR')) || 0,
    estado: g(r, 'ESTADO')
  }));
}

function agrupar_(regs, fnClave) {
  const m = {};
  regs.forEach(r => {
    const k = fnClave(r) || '(sin dato)';
    const a = m[k] = m[k] || { clases: 0, asist: 0, falt: 0, deser: 0, min: 0, rev: 0 };
    a.clases++; a.asist += r.asist; a.falt += r.falt; a.deser += r.deser; a.min += r.dur;
    if (r.estado === 'REVISAR') a.rev++;
  });
  return Object.keys(m).sort().map(k => {
    const a = m[k], t = a.asist + a.falt;
    return [k, a.clases, a.asist, a.falt, t ? a.asist / t : 0, a.deser, Math.round(a.min / 6) / 10, a.rev];
  });
}

function generarResumen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regs = leerRegistros_(hojaRespuestas_());
  let sh = ss.getSheetByName(CONFIG.HOJA_RESUMEN) || ss.insertSheet(CONFIG.HOJA_RESUMEN);
  sh.clear();
  const enc = ['Grupo', 'Clases', 'Asistentes', 'Faltantes', '% Asistencia', 'Deserción', 'Horas dictadas', 'Por revisar'];
  let fila = 1;
  sh.getRange(fila, 1).setValue('RESUMEN CONTROL DE ASISTENCIA DIN').setFontSize(14).setFontWeight('bold');
  sh.getRange(fila + 1, 1).setValue('Generado: ' + Utilities.formatDate(new Date(), CONFIG.TZ, 'dd/MM/yyyy HH:mm') +
    ' · Registros: ' + regs.length);
  fila += 3;

  const tablas = [
    ['Por docente', r => r.docente + (r.cond ? ' (' + r.cond + ')' : '')],
    ['Por núcleo', r => r.nucleo],
    ['Por localidad', r => r.localidad],
    ['Por semestre', r => r.semestre],
    ['Por mes', r => Utilities.formatDate(r.fecha, CONFIG.TZ, 'yyyy-MM')]
  ];
  tablas.forEach(t => {
    const filas = agrupar_(regs, t[1]);
    sh.getRange(fila, 1).setValue(t[0]).setFontWeight('bold').setFontSize(12);
    sh.getRange(fila + 1, 1, 1, enc.length).setValues([enc]).setFontWeight('bold').setBackground('#d9ead3');
    if (filas.length) {
      sh.getRange(fila + 2, 1, filas.length, enc.length).setValues(filas);
      sh.getRange(fila + 2, 5, filas.length, 1).setNumberFormat('0.0%');
    }
    fila += filas.length + 4;
  });
  sh.autoResizeColumns(1, enc.length);
  sh.setColumnWidth(1, 300);
  ss.setActiveSheet(sh);
}

function enviarResumenSemanal() {
  const regs = leerRegistros_(hojaRespuestas_());
  const desde = new Date(Date.now() - CONFIG.DIAS_SIN_REPORTE * 86400000);
  const recientes = regs.filter(r => r.fecha >= desde);
  const reportaron = {};
  recientes.forEach(r => { reportaron[r.ced] = (reportaron[r.ced] || 0) + 1; });

  const sinReporte = leerDocentes_().filter(dc => !reportaron[dc.cedula]);
  const asist = recientes.reduce((s, r) => s + r.asist, 0);
  const falt = recientes.reduce((s, r) => s + r.falt, 0);
  const pct = asist + falt ? (100 * asist / (asist + falt)).toFixed(1) + '%' : 'N/D';
  const rev = recientes.filter(r => r.estado === 'REVISAR').length;

  const html = '<h3>Resumen semanal – Control de asistencia DIN</h3>' +
    '<p>Período: últimos ' + CONFIG.DIAS_SIN_REPORTE + ' días.</p><ul>' +
    '<li>Clases reportadas: <b>' + recientes.length + '</b></li>' +
    '<li>Asistentes: <b>' + asist + '</b> · Faltantes: <b>' + falt + '</b> · Asistencia: <b>' + pct + '</b></li>' +
    '<li>Reportes por revisar: <b>' + rev + '</b></li></ul>' +
    '<h4>Docentes sin reporte (' + sinReporte.length + ')</h4>' +
    (sinReporte.length
      ? '<ul>' + sinReporte.map(dc => '<li>' + esc_(dc.nombre) + ' – ' + esc_(dc.cedula) + ' (' + esc_(dc.condicion) + ')</li>').join('') + '</ul>'
      : '<p>Todos los docentes activos reportaron.</p>') +
    '<p><i>Nota: quienes notificaron inasistencia por el otro formulario pueden aparecer aquí.</i></p>';

  MailApp.sendEmail({
    to: CONFIG.COORDINADOR_EMAIL,
    subject: 'Resumen semanal DIN – ' + Utilities.formatDate(new Date(), CONFIG.TZ, 'dd/MM/yyyy'),
    htmlBody: html
  });
}

// ===================== WEB APP (INTERFAZ HTML) =====================
/**
 * WEB APP – Interfaz HTML del Control de asistencia DIN
 * Archivo del MISMO proyecto de Apps Script (junto a ControlAsistenciaDIN.gs e Index.html).
 *
 * Incluye:
 *  - Registro de clases.
 *  - Auto-registro de docentes nuevos (se dan de alta al reportar su primera clase).
 *  - Gestión de docentes (agregar, editar, activar/desactivar) solo para administradores.
 */

const CATALOGOS = {
  nucleos: ['LARA'],
  localidades: ['UNEFA LARA DE LA 25', 'UNEFA LARA DE LA 51', 'UNEFA LARA DE LA 57'],
  semestres: ['00S', '01S', '02S', '03S', '04S', '05S', '06S', '07S', '08S'],
  carreras: ['2613', '2213', '0313', '0913', '1013', '2013'],
  secciones: ['D1', 'D2', 'D3', 'D4', 'D5', 'D6', 'D7', 'N1', 'N2', 'N3', 'N4', 'N5'],
  unidades: ['UNIDAD 1', 'UNIDAD 2', 'UNIDAD 3', 'UNIDAD 4', 'UNIDAD 5', 'UNIDAD 6', 'UNIDAD 7'],
  condiciones: ['TV', 'FIJO', 'ORDINARIO', 'CONTRATADO', 'HP']
};
const CARPETA_FOTOS = 'Reseñas fotográficas DIN';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Control de asistencia DIN')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function obtenerOpciones() {
  return {
    docentes: leerDocentes_().map(d => ({
      id: d.cedula,
      label: d.nombre + ' · ' + d.cedula + (d.condicion ? ' (' + d.condicion + ')' : '')
    })),
    catalogos: CATALOGOS,
    usuario: Session.getActiveUser().getEmail(),
    esAdmin: esAdmin_(),
    urlHoja: SpreadsheetApp.getActiveSpreadsheet().getUrl()
  };
}

// ============================ PERMISOS ====================================
function esAdmin_() {
  const me = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!me) return false;
  return [CONFIG.COORDINADOR_EMAIL].concat(CONFIG.ADMINS || [])
    .some(e => String(e).toLowerCase() === me);
}
function exigirAdmin_() {
  if (!esAdmin_()) throw new Error('Solo la coordinación puede realizar esta acción.');
}

// ============================ DOCENTES ====================================
function titulo_(s) {
  return String(s || '').trim().replace(/\s+/g, ' ').toLowerCase()
    .replace(/(^|\s)(\S)/g, (m, a, b) => a + b.toUpperCase());
}

/** Hoja Docentes con la columna "Alta" garantizada (A:Nombre B:Cédula C:Condición D:Activo E:Alta). */
function hojaDocentes_() {
  const sh = crearHojaDocentes_();
  if (sh.getMaxColumns() < 5) sh.insertColumnsAfter(sh.getMaxColumns(), 5 - sh.getMaxColumns());
  if (!sh.getRange(1, 5).getValue()) {
    sh.getRange(1, 5).setValue('Alta').setFontWeight('bold').setBackground('#d9ead3');
  }
  return sh;
}

function buscarDocente_(ced) {
  const sh = hojaDocentes_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return null;
  const vals = sh.getRange(2, 1, n, 5).getValues();
  for (let i = 0; i < n; i++) {
    if (String(vals[i][1]).replace(/\D/g, '') === ced) {
      return {
        fila: i + 2, nombre: String(vals[i][0]).trim(), cedula: ced,
        condicion: String(vals[i][2]).trim(),
        activo: String(vals[i][3]).trim().toUpperCase() !== 'NO'
      };
    }
  }
  return null;
}

function validarDocente_(d) {
  const ced = String(d.cedula || '').replace(/\D/g, '');
  if (!/^\d{6,9}$/.test(ced)) throw new Error('La cédula debe tener entre 6 y 9 dígitos.');
  const nombre = titulo_(d.nombre);
  if (nombre.split(' ').length < 2) throw new Error('Escriba nombres y apellidos.');
  const cond = String(d.condicion || '').trim().toUpperCase();
  if (cond && CATALOGOS.condiciones.indexOf(cond) === -1) throw new Error('Condición no válida.');
  return { ced: ced, nombre: nombre, cond: cond };
}

/** Agrega un docente nuevo (sin bloqueo propio: el que llama debe tener el LockService). */
function agregarDocente_(d, origen) {
  const v = validarDocente_(d);
  if (buscarDocente_(v.ced)) throw new Error('Ya existe un docente con esa cédula.');
  const sh = hojaDocentes_();
  const fila = sh.getLastRow() + 1;
  sh.getRange(fila, 2).setNumberFormat('@');
  sh.getRange(fila, 1, 1, 5).setValues([[
    v.nombre, v.ced, v.cond, 'SI',
    origen + ' ' + Utilities.formatDate(new Date(), CONFIG.TZ, 'dd/MM/yyyy')
  ]]);
  return { fila: fila, nombre: v.nombre, cedula: v.ced, condicion: v.cond, activo: true };
}

/** Mantiene la lista del formulario de Google igual a la hoja Docentes. Devuelve '' o un aviso. */
function sincronizarFormulario_() {
  try {
    actualizarFormulario();
    return '';
  } catch (e) {
    return 'Guardado, pero no se pudo sincronizar el formulario (' + e.message +
      '). Use el menú Control DIN → Sincronizar docentes.';
  }
}

// ---------- Gestión (solo administradores) ----------
function listarDocentesAdmin() {
  exigirAdmin_();
  const sh = hojaDocentes_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 5).getDisplayValues()
    .map(r => ({
      nombre: r[0], cedula: String(r[1]).replace(/\D/g, ''), condicion: r[2],
      activo: String(r[3]).trim().toUpperCase() !== 'NO', alta: r[4]
    }))
    .filter(d => d.nombre)
    .sort((a, b) => a.nombre.localeCompare(b.nombre, 'es'));
}

/** Agrega un docente o, si la cédula ya existe, actualiza nombre y condición. */
function guardarDocente(d) {
  exigirAdmin_();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let mensaje;
  try {
    const v = validarDocente_(d);
    const ex = buscarDocente_(v.ced);
    if (ex) {
      const sh = hojaDocentes_();
      sh.getRange(ex.fila, 1).setValue(v.nombre);
      sh.getRange(ex.fila, 3).setValue(v.cond);
      mensaje = 'Docente actualizado.';
    } else {
      agregarDocente_(d, 'Admin');
      mensaje = 'Docente agregado.';
    }
  } finally {
    lock.releaseLock();
  }
  return { mensaje: mensaje, aviso: sincronizarFormulario_() };
}

function cambiarEstadoDocente(cedula, activo) {
  exigirAdmin_();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ex = buscarDocente_(String(cedula).replace(/\D/g, ''));
    if (!ex) throw new Error('Docente no encontrado.');
    hojaDocentes_().getRange(ex.fila, 4).setValue(activo ? 'SI' : 'NO');
  } finally {
    lock.releaseLock();
  }
  return { mensaje: activo ? 'Docente activado.' : 'Docente desactivado.', aviso: sincronizarFormulario_() };
}

// ============================ GUARDAR REPORTE =============================
function guardarReporte(d) {
  const req = ['fecha', 'inicio', 'salida', 'nucleo', 'localidad', 'semestre', 'carrera',
    'seccion', 'asist', 'falt', 'tema', 'unidad'];
  req.forEach(k => {
    if (d[k] === undefined || String(d[k]).trim() === '') throw new Error('Falta el campo: ' + k);
  });
  if (!d.docenteId && !d.nuevo) throw new Error('Seleccione un docente o regístrese.');
  if (!d.docenteId && d.nuevo && !String(d.nuevo.condicion || '').trim()) {
    throw new Error('Seleccione su condición.');
  }

  let registrado = false, sincroniza = false;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let resultado;
  try {
    // 1) Docente: de la lista o auto-registro
    let doc = d.docenteId ? mapaDocentes_()[d.docenteId] : null;
    if (!doc && d.nuevo) {
      const v = validarDocente_({
        nombre: String(d.nuevo.nombres || '') + ' ' + String(d.nuevo.apellidos || ''),
        cedula: d.nuevo.cedula, condicion: d.nuevo.condicion
      });
      const ex = buscarDocente_(v.ced);
      if (ex) {
        if (!ex.activo) throw new Error('Esta cédula figura como inactiva. Comuníquese con la coordinación.');
        doc = ex; // ya estaba registrado: se usa el existente
      } else {
        doc = agregarDocente_({ nombre: v.nombre, cedula: v.ced, condicion: v.cond }, 'Auto-registro');
        registrado = true;
      }
    }
    if (!doc) throw new Error('Docente no encontrado en la lista.');

    // 2) Fila del reporte
    const sh = hojaRespuestas_();
    const c = prepararHoja_(sh);
    const fila = new Array(sh.getLastColumn()).fill('');
    const set = (k, val) => { if (c[k]) fila[c[k] - 1] = val; };
    const f = String(d.fecha).split('-').map(Number);
    fila[0] = new Date();
    set('EMAIL', Session.getActiveUser().getEmail());
    set('DOCENTE', doc.nombre + ' ' + doc.cedula + (doc.condicion ? ' (' + doc.condicion + ')' : ''));
    set('OTRO', '');
    set('NUCLEO', d.nucleo);
    set('LOCALIDAD', d.localidad);
    set('FECHA', new Date(f[0], f[1] - 1, f[2], 12, 0, 0));
    set('INICIO', d.inicio);
    set('SALIDA', d.salida);
    set('SEMESTRE', d.semestre);
    set('CARRERA', d.carrera);
    set('SECCION', d.seccion);
    set('ASIST', num_(d.asist));
    set('FALT', num_(d.falt));
    set('DESER', num_(d.deser) || 0);
    set('IMPACTO', d.impacto || '');
    set('TEMA', d.tema);
    set('UNIDAD', d.unidad);
    set('OBS', d.obs || '');
    if (d.foto && d.foto.data) set('FOTO', guardarFoto_(d.foto));

    sh.appendRow(fila);
    const row = sh.getLastRow();
    procesarFila_(sh, row, construirContexto_(sh), true);
    resultado = {
      fila: row,
      estado: sh.getRange(row, c.ESTADO).getValue(),
      alertas: String(sh.getRange(row, c.ALERTA).getValue() || ''),
      registrado: registrado,
      docente: doc.nombre
    };
    sincroniza = registrado;
  } finally {
    lock.releaseLock();
  }

  // 3) Si hubo alta nueva: actualizar el formulario de Google y avisar a la coordinación
  if (sincroniza) {
    sincronizarFormulario_();
    try {
      MailApp.sendEmail({
        to: CONFIG.COORDINADOR_EMAIL,
        subject: 'Nuevo docente registrado: ' + resultado.docente,
        htmlBody: '<p>' + esc_(resultado.docente) + ' se registró desde la interfaz web al reportar su primera clase.</p>' +
          '<p>Puede revisarlo, editarlo o desactivarlo en la pestaña <b>Docentes</b> de la aplicación.</p>'
      });
    } catch (err) {
      console.error('Aviso de nuevo docente: ' + err);
    }
  }
  return resultado;
}

function guardarFoto_(foto) {
  const it = DriveApp.getFoldersByName(CARPETA_FOTOS);
  const carpeta = it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA_FOTOS);
  const bytes = Utilities.base64Decode(String(foto.data).split(',').pop());
  const nombre = Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyyMMdd_HHmmss') + '_' +
    String(foto.nombre || 'foto').replace(/[^\w.\-]/g, '_');
  const archivo = carpeta.createFile(Utilities.newBlob(bytes, foto.tipo || 'image/jpeg', nombre));
  return archivo.getUrl();
}

// ============================== CONSULTAS =================================
function listarRegistros() {
  const sh = hojaRespuestas_();
  const c = prepararHoja_(sh);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const desde = Math.max(2, last - 499); // últimos 500
  const disp = sh.getRange(desde, 1, last - desde + 1, sh.getLastColumn()).getDisplayValues();
  const g = (r, k) => (c[k] ? r[c[k] - 1] : '');
  return disp.map((r, i) => ({
    fila: desde + i,
    fecha: g(r, 'FECHA'),
    docente: g(r, 'DOCENTE') || g(r, 'OTRO'),
    localidad: g(r, 'LOCALIDAD'),
    carrera: g(r, 'CARRERA'),
    seccion: g(r, 'SECCION'),
    asist: g(r, 'ASIST'),
    falt: g(r, 'FALT'),
    pct: g(r, 'PCT'),
    estado: g(r, 'ESTADO'),
    alertas: g(r, 'ALERTA')
  })).reverse();
}

function obtenerResumen(dias) {
  let regs = leerRegistros_(hojaRespuestas_());
  if (dias) {
    const desde = new Date(Date.now() - dias * 86400000);
    regs = regs.filter(r => r.fecha >= desde);
  }
  const asist = regs.reduce((s, r) => s + r.asist, 0);
  const falt = regs.reduce((s, r) => s + r.falt, 0);
  return {
    kpi: {
      clases: regs.length,
      asist: asist,
      falt: falt,
      pct: asist + falt ? asist / (asist + falt) : 0,
      rev: regs.filter(r => r.estado === 'REVISAR').length,
      horas: Math.round(regs.reduce((s, r) => s + r.dur, 0) / 6) / 10
    },
    docentes: agrupar_(regs, r => r.docente),
    localidades: agrupar_(regs, r => r.localidad)
  };
}

// ================================ EXCEL ===================================
/** Exporta TODO el libro (respuestas, docentes, resumen) como .xlsx y lo devuelve en base64. */
function exportarExcel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  DriveApp.getFileById(ss.getId()).getName(); // fuerza el permiso de Drive necesario para exportar
  const resp = UrlFetchApp.fetch(
    'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx',
    { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true }
  );
  if (resp.getResponseCode() !== 200) throw new Error('No se pudo exportar (HTTP ' + resp.getResponseCode() + ')');
  return {
    nombre: 'Control_Asistencia_DIN_' + Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyyMMdd_HHmm') + '.xlsx',
    base64: Utilities.base64Encode(resp.getBlob().getBytes())
  };
}
