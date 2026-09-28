/**
 * Programa de Atención a la Diversidad — Backend
 * CEIP Carlos III · La Carlota
 *
 * Hoja maestra (registro): 11bkLpUZKKkSbEPZkmCPqI23LBWreishKmIYL1yRJS74
 *
 * Arquitectura:
 *   - SS_ID = Spreadsheet maestro: pestañas Config (centro/localidad) + Cursos (registro de años académicos).
 *   - Por cada curso académico, un Spreadsheet propio en la misma carpeta de Drive,
 *     con pestañas Config (cursoEscolar + listas de cursos/docentes) + Índice + una pestaña por alumno.
 */

const SS_ID = '11bkLpUZKKkSbEPZkmCPqI23LBWreishKmIYL1yRJS74';
const INDICE_TAB = 'Índice';
const CONFIG_TAB = 'Config';
const CURSOS_TAB = 'Cursos';
const LOCKS_TAB = 'Locks';
const PRESENCE_TTL_MS = 5 * 60 * 1000;

/* ───────── Web App entry point ───────── */

function doGet() {
  if (!isAuthorized_()) {
    const email = getCurrentUserEmail_();
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;max-width:520px;margin:15vh auto;padding:0 16px;color:#1b4332">' +
      '<h2>Acceso no autorizado</h2>' +
      '<p>La cuenta <b>' + escapeHtml_(email || 'desconocida') + '</b> no está en la lista de usuarios autorizados.</p>' +
      '<p>Pide al administrador de la aplicación que te añada desde Ajustes.</p></div>')
      .setTitle('Acceso no autorizado')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Programas de Atención a la Diversidad')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ───────── Autorización (lista de usuarios) ───────── */
//  - Administrador: la cuenta que despliega el script (Session.getEffectiveUser).
//  - Usuarios autorizados: claves "usuario" en la pestaña Config del maestro.
//  - Si la lista está vacía, se permite el acceso a todo el dominio (comportamiento previo).

const USERS_KEY = 'usuario';
const USERS_CACHE_KEY = 'allowedUsers.v1';

function escapeHtml_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function normEmail_(e) {
  return String(e || '').trim().toLowerCase();
}

function getAdminEmail_() {
  try { return normEmail_(Session.getEffectiveUser().getEmail()); } catch (e) { return ''; }
}

function getAllowedUsers_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(USERS_CACHE_KEY);
  if (cached !== null) {
    try { return JSON.parse(cached); } catch (e) { /* recalcular */ }
  }
  const ss = getMasterSS_();
  const cfg = readKeyValueSheet_(ss.getSheetByName(CONFIG_TAB));
  const raw = cfg[USERS_KEY];
  const list = (Array.isArray(raw) ? raw : (raw ? [raw] : [])).map(normEmail_).filter(function(e) { return e; });
  cache.put(USERS_CACHE_KEY, JSON.stringify(list), 300);
  return list;
}

function isAdmin_() {
  const admin = getAdminEmail_();
  return !!admin && normEmail_(getCurrentUserEmail_()) === admin;
}

function isAuthorized_() {
  if (isAdmin_()) return true;
  const list = getAllowedUsers_();
  if (!list.length) return true;
  const me = normEmail_(getCurrentUserEmail_());
  return !!me && list.indexOf(me) !== -1;
}

function assertAuthorized_() {
  if (!isAuthorized_()) {
    throw new Error('ACCESO_DENEGADO: tu cuenta no está autorizada para usar esta aplicación.');
  }
}

function assertAdmin_() {
  if (!isAdmin_()) throw new Error('Solo el administrador puede realizar esta acción.');
}

function getAllowedUsers() {
  assertAdmin_();
  return { admin: getAdminEmail_(), users: getAllowedUsers_() };
}

function saveAllowedUsers(payload) {
  assertAdmin_();
  const data = JSON.parse(payload);
  const seen = {};
  const users = [];
  (data.users || []).forEach(function(u) {
    const e = normEmail_(u);
    if (!e) return;
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e)) throw new Error('Correo no válido: ' + e);
    if (!seen[e]) { seen[e] = true; users.push(e); }
  });
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('Otro usuario está guardando. Reintenta.');
  try {
    const ss = getMasterSS_();
    const cfg = getOrCreateConfigIn_(ss);
    setMultiKV_(cfg, USERS_KEY, users);
    CacheService.getScriptCache().remove(USERS_CACHE_KEY);
  } finally {
    lock.releaseLock();
  }
  return { success: true, users: users };
}

/* ───────── Helpers: Spreadsheets ───────── */

function getMasterSS_() {
  return SpreadsheetApp.openById(SS_ID);
}

function getMasterParentFolder_() {
  const file = DriveApp.getFileById(SS_ID);
  const parents = file.getParents();
  return parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
}

function autoCursoEscolar_() {
  const now = new Date();
  const y = now.getFullYear();
  const m = now.getMonth();
  return m >= 8 ? (y + '/' + (y + 1)) : ((y - 1) + '/' + y);
}

function readKeyValueSheet_(sheet) {
  const out = {};
  if (!sheet) return out;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0] || '').trim();
    if (!key) continue;
    if (out[key] === undefined) {
      out[key] = String(data[i][1] || '').trim();
    } else {
      // Multi-valor (e.g. courses, docentes): convertir en array
      if (!Array.isArray(out[key])) out[key] = [out[key]];
      out[key].push(String(data[i][1] || '').trim());
    }
  }
  return out;
}

// Pestaña Config (CLAVE/VALOR) de un spreadsheet; la crea con cabecera si no existe.
function getOrCreateConfigIn_(ss) {
  let sheet = ss.getSheetByName(CONFIG_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG_TAB);
    sheet.appendRow(['CLAVE', 'VALOR']);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/* ───────── Cursos registry (en maestro) ───────── */

function getOrCreateCursos_() {
  const ss = getMasterSS_();
  let sheet = ss.getSheetByName(CURSOS_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(CURSOS_TAB);
    sheet.appendRow(['CURSO_ESCOLAR', 'SPREADSHEET_ID', 'URL', 'CREADO_EN', 'ARCHIVADO']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 140);
    sheet.setColumnWidth(2, 360);
    sheet.setColumnWidth(3, 360);
    sheet.setColumnWidth(4, 180);
    sheet.setColumnWidth(5, 100);
  }
  return sheet;
}

// Registro de Cursos cacheado (se lee en casi todas las llamadas). Se invalida al
// duplicar, archivar o borrar un curso; ediciones manuales de la hoja tardan ≤10 min.
const CURSOS_CACHE_KEY = 'cursosRows.v1';

function invalidateCursosCache_() {
  CacheService.getScriptCache().remove(CURSOS_CACHE_KEY);
}

function readCursosRows_(fresh) {
  const cache = CacheService.getScriptCache();
  if (!fresh) {
    const cached = cache.get(CURSOS_CACHE_KEY);
    if (cached !== null) {
      try { return JSON.parse(cached); } catch (e) { /* recalcular */ }
    }
  }
  const rows = readCursosRowsFromSheet_();
  try { cache.put(CURSOS_CACHE_KEY, JSON.stringify(rows), 600); } catch (e) { /* demasiado grande: sin caché */ }
  return rows;
}

function readCursosRowsFromSheet_() {
  const sheet = getOrCreateCursos_();
  const data = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][1]) continue;
    rows.push({
      label: String(data[i][0] || '').trim(),
      id: String(data[i][1] || '').trim(),
      url: String(data[i][2] || '').trim(),
      createdAt: String(data[i][3] || '').trim(),
      archived: String(data[i][4] || '').toUpperCase() === 'TRUE'
    });
  }
  return rows;
}

function getActiveYearId_(rows) {
  rows = rows || readCursosRows_();
  let best = null;
  rows.forEach(function(r) {
    if (r.archived) return;
    if (!best) { best = r; return; }
    const a = new Date(r.createdAt).getTime() || 0;
    const b = new Date(best.createdAt).getTime() || 0;
    if (a > b) best = r;
  });
  if (!best && rows.length) {
    // Si todos están archivados, usar el más reciente
    rows.forEach(function(r) {
      if (!best) { best = r; return; }
      const a = new Date(r.createdAt).getTime() || 0;
      const b = new Date(best.createdAt).getTime() || 0;
      if (a > b) best = r;
    });
  }
  return best ? best.id : null;
}

function getYearSS_(yearId) {
  if (yearId) {
    assertKnownYear_(yearId);
  } else {
    yearId = getActiveYearId_();
  }
  if (!yearId) throw new Error('No hay curso académico activo. Crea uno desde Ajustes.');
  return SpreadsheetApp.openById(yearId);
}

// Solo se permite operar sobre hojas registradas en Cursos (evita abrir IDs arbitrarios).
function assertKnownYear_(yearId) {
  const id = String(yearId || '').trim();
  if (!id || !readCursosRows_().some(function(r) { return r.id === id; })) {
    throw new Error('Curso académico no válido.');
  }
}

/* ───────── Helpers: validación de datos ───────── */

// Pestañas internas que no pueden usarse como nombre de alumno.
function assertValidStudentName_(name) {
  const n = String(name || '').trim().toLowerCase();
  if ([CONFIG_TAB, INDICE_TAB, LOCKS_TAB].some(function(t) { return t.toLowerCase() === n; })) {
    throw new Error('"' + String(name).trim() + '" es un nombre reservado. Usa otro nombre para el alumno.');
  }
}

// Evita que un texto que empieza por = + - @ se interprete como fórmula en la hoja.
// El apóstrofo inicial no forma parte del valor: getValues() devuelve el texto original.
function safeCell_(v) {
  return (typeof v === 'string' && /^[=+\-@]/.test(v)) ? "'" + v : v;
}

/* ───────── Migración del maestro (una vez) ───────── */

const MIGRATED_PROP = 'masterMigrated.v1';

function migrateMasterIfNeeded_() {
  // Marca en ScriptProperties para no abrir el maestro en cada llamada.
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(MIGRATED_PROP) === '1') return;
  const ss = getMasterSS_();
  if (ss.getSheetByName(CURSOS_TAB)) { props.setProperty(MIGRATED_PROP, '1'); return; } // ya migrado

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('Migración en curso por otro usuario, reintenta en unos segundos.');
  }
  try {
    if (ss.getSheetByName(CURSOS_TAB)) return;

    const masterCfgSheet = ss.getSheetByName(CONFIG_TAB);
    const masterCfg = readKeyValueSheet_(masterCfgSheet);
    const cursoEscolar = (masterCfg.cursoEscolar || autoCursoEscolar_()).trim();
    const safeYearLabel = cursoEscolar.replace(/\//g, '-');
    const fileName = 'Programa Atención a la Diversidad — ' + safeYearLabel;

    // 1) Copia el maestro entero al mismo directorio
    const parentFolder = getMasterParentFolder_();
    const masterFile = DriveApp.getFileById(SS_ID);
    const newFile = masterFile.makeCopy(fileName, parentFolder);
    const newId = newFile.getId();
    const newUrl = newFile.getUrl();

    // 2) En la copia: dejar Config con sólo cursoEscolar (eliminar centro/localidad)
    const newSS = SpreadsheetApp.openById(newId);
    let newCfg = newSS.getSheetByName(CONFIG_TAB);
    if (!newCfg) newCfg = newSS.insertSheet(CONFIG_TAB);
    newCfg.clear();
    newCfg.appendRow(['CLAVE', 'VALOR']);
    newCfg.appendRow(['cursoEscolar', cursoEscolar]);
    newCfg.getRange(1, 1, 1, 2).setFontWeight('bold');
    newCfg.setFrozenRows(1);

    // 3) En el maestro: insertar Cursos, registrar el año, y borrar todo lo que no sea Config/Cursos
    const cursos = ss.insertSheet(CURSOS_TAB);
    cursos.appendRow(['CURSO_ESCOLAR', 'SPREADSHEET_ID', 'URL', 'CREADO_EN', 'ARCHIVADO']);
    cursos.appendRow([cursoEscolar, newId, newUrl, new Date().toISOString(), 'FALSE']);
    cursos.getRange(1, 1, 1, 5).setFontWeight('bold');
    cursos.setFrozenRows(1);
    cursos.setColumnWidth(1, 140);
    cursos.setColumnWidth(2, 360);
    cursos.setColumnWidth(3, 360);
    cursos.setColumnWidth(4, 180);
    cursos.setColumnWidth(5, 100);

    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      const name = sheets[i].getName();
      if (name !== CONFIG_TAB && name !== CURSOS_TAB) {
        ss.deleteSheet(sheets[i]);
      }
    }

    // 4) Reducir Config del maestro a centro/localidad
    let mCfg = ss.getSheetByName(CONFIG_TAB);
    if (!mCfg) mCfg = ss.insertSheet(CONFIG_TAB);
    mCfg.clear();
    mCfg.appendRow(['CLAVE', 'VALOR']);
    mCfg.appendRow(['centro', masterCfg.centro || '']);
    mCfg.appendRow(['localidad', masterCfg.localidad || '']);
    mCfg.getRange(1, 1, 1, 2).setFontWeight('bold');
    mCfg.setFrozenRows(1);
    invalidateCursosCache_();
    PropertiesService.getScriptProperties().setProperty(MIGRATED_PROP, '1');
  } finally {
    lock.releaseLock();
  }
}

/* ───────── CONFIG: datos del centro y del año ───────── */

function getConfig(yearId) {
  assertAuthorized_();
  migrateMasterIfNeeded_();
  return buildConfig_(yearId).config;
}

// Carga inicial en una sola llamada: configuración + lista de alumnos del año activo.
function getBootstrap(yearId) {
  assertAuthorized_();
  migrateMasterIfNeeded_();
  const built = buildConfig_(yearId);
  return {
    config: built.config,
    students: built.yearSS ? listStudentsIn_(built.yearSS) : []
  };
}

function buildConfig_(yearId) {
  // Maestro: centro/localidad + lista de años
  const ss = getMasterSS_();
  const masterCfg = readKeyValueSheet_(ss.getSheetByName(CONFIG_TAB));
  const years = readCursosRows_();

  // Año activo (resuelto a partir de yearId o más reciente no archivado)
  const resolvedId = (yearId && years.some(function(y) { return y.id === yearId; }))
    ? yearId
    : getActiveYearId_(years);

  let cursoEscolar = '';
  let courses = [];
  let docentes = [];
  let yearSS = null;

  if (resolvedId) {
    yearSS = SpreadsheetApp.openById(resolvedId);
    const yearCfg = readKeyValueSheet_(yearSS.getSheetByName(CONFIG_TAB));
    cursoEscolar = yearCfg.cursoEscolar || autoCursoEscolar_();
    courses = Array.isArray(yearCfg.course) ? yearCfg.course : (yearCfg.course ? [yearCfg.course] : []);
    docentes = Array.isArray(yearCfg.docente) ? yearCfg.docente : (yearCfg.docente ? [yearCfg.docente] : []);
  } else {
    cursoEscolar = autoCursoEscolar_();
  }

  return {
    yearSS: yearSS,
    config: {
      centro: masterCfg.centro || '',
      localidad: masterCfg.localidad || '',
      cursoEscolar: cursoEscolar,
      activeYearId: resolvedId || '',
      years: years,
      courses: courses,
      docentes: docentes,
      isAdmin: isAdmin_()
    }
  };
}

function saveConfig(payload) {
  assertAuthorized_();
  migrateMasterIfNeeded_();
  const data = JSON.parse(payload);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('Otro usuario está guardando. Reintenta.');
  try {
    const ss = getMasterSS_();
    const mCfg = getOrCreateConfigIn_(ss);
    upsertKV_(mCfg, 'centro', data.centro != null ? String(data.centro) : '');
    upsertKV_(mCfg, 'localidad', data.localidad != null ? String(data.localidad) : '');

    // cursoEscolar va al año activo (o al indicado)
    if (data.cursoEscolar !== undefined) {
      const yearId = data.yearId || getActiveYearId_();
      if (data.yearId) assertKnownYear_(data.yearId);
      if (yearId) {
        const yearSS = SpreadsheetApp.openById(yearId);
        const yCfg = getOrCreateConfigIn_(yearSS);
        upsertKV_(yCfg, 'cursoEscolar', String(data.cursoEscolar).trim());
      }
    }
  } finally {
    lock.releaseLock();
  }
  return { success: true };
}

function upsertKV_(sheet, key, value) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) {
      sheet.getRange(i + 1, 2).setValue(safeCell_(value));
      return;
    }
  }
  sheet.appendRow([key, safeCell_(value)]);
}

function setMultiKV_(sheet, key, values) {
  // Elimina todas las filas con esa clave y reescribe una fila por valor (en bloque).
  const data = sheet.getDataRange().getValues();
  const header = data[0] || ['CLAVE', 'VALOR'];
  const kept = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== key) kept.push([data[i][0], data[i][1]].map(safeCell_));
  }
  (values || []).forEach(function(v) {
    const s = String(v == null ? '' : v).trim();
    if (s) kept.push([key, safeCell_(s)]);
  });
  const oldRows = data.length;
  const newRows = kept.length + 1;
  sheet.getRange(1, 1, 1, 2).setValues([[header[0], header[1]]]);
  if (kept.length) sheet.getRange(2, 1, kept.length, 2).setValues(kept);
  if (oldRows > newRows) sheet.getRange(newRows + 1, 1, oldRows - newRows, sheet.getMaxColumns()).clearContent();
}

function saveYearLists(payload) {
  assertAuthorized_();
  migrateMasterIfNeeded_();
  const data = JSON.parse(payload);
  const yearId = data.yearId || getActiveYearId_();
  if (!yearId) throw new Error('No hay curso académico activo.');
  if (data.yearId) assertKnownYear_(data.yearId);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('Otro usuario está guardando. Reintenta.');
  try {
    const yearSS = SpreadsheetApp.openById(yearId);
    const cfg = getOrCreateConfigIn_(yearSS);
    if (Array.isArray(data.courses)) setMultiKV_(cfg, 'course', data.courses);
    if (Array.isArray(data.docentes)) setMultiKV_(cfg, 'docente', data.docentes);
  } finally {
    lock.releaseLock();
  }
  return { success: true };
}

/* ───────── Índice y READ: lista de alumnos ───────── */

function getOrCreateIndiceIn_(ss) {
  let sheet = ss.getSheetByName(INDICE_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(INDICE_TAB, 0);
    sheet.appendRow(['ALUMNO/A', 'CURSO', 'PROGRAMA', 'ÁREA/ÁMBITO', 'DOCENTES']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    sheet.setFrozenRows(1);
    return sheet;
  }
  // Migración: si la cabecera no tiene DOCENTES, añadirla
  const lastCol = sheet.getLastColumn();
  if (lastCol < 5) {
    sheet.getRange(1, 5).setValue('DOCENTES').setFontWeight('bold');
  }
  return sheet;
}

function getStudentList(yearId) {
  assertAuthorized_();
  migrateMasterIfNeeded_();
  return listStudentsIn_(getYearSS_(yearId));
}

function listStudentsIn_(ss) {
  const sheet = getOrCreateIndiceIn_(ss);
  const data = sheet.getDataRange().getValues();
  const students = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] && String(row[0]).trim()) {
      students.push({
        name: String(row[0]).trim(),
        course: String(row[1] || '').trim(),
        program: String(row[2] || '').trim(),
        area: String(row[3] || '').trim(),
        docentes: parseDocentesString_(String(row[4] || ''))
      });
    }
  }
  return students;
}

function parseDocentesString_(s) {
  if (!s) return [];
  return s.split('|').map(function(x) { return x.trim(); }).filter(function(x) { return x; });
}

function serializeDocentes_(arr) {
  if (!arr || !arr.length) return '';
  return arr.map(function(x) { return String(x || '').trim(); }).filter(function(x) { return x; }).join('|');
}

/* ───────── READ: datos de un alumno ───────── */

function getStudentData(studentName, yearId) {
  assertAuthorized_();
  migrateMasterIfNeeded_();
  assertValidStudentName_(studentName);
  const ss = getYearSS_(yearId);
  const sheet = ss.getSheetByName(studentName);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;

  // Row 1: metadata
  const meta = data[0];
  let docentesStr = '';
  if (String(meta[8] || '').trim().toUpperCase() === 'DOCENTES') {
    docentesStr = String(meta[9] || '').trim();
  }
  let updatedAt = '';
  let updatedBy = '';
  if (String(meta[10] || '').trim().toUpperCase() === 'UPDATED_AT') {
    updatedAt = String(meta[11] || '').trim();
  }
  if (String(meta[12] || '').trim().toUpperCase() === 'UPDATED_BY') {
    updatedBy = String(meta[13] || '').trim();
  }
  const result = {
    studentName: String(meta[1] || '').trim(),
    course: String(meta[3] || '').trim(),
    docentes: parseDocentesString_(docentesStr),
    updatedAt: updatedAt,
    updatedBy: updatedBy,
    valoracionInicial: '',
    seguimiento1T: '',
    seguimiento2T: '',
    seguimiento3T: '',
    areas: [],
    informeFamilias: { indicadores: [], evaluaciones: [] }
  };

  // Find the header row (contains 'TIPO') then read data after it
  let headerIndex = -1;
  for (let h = 1; h < data.length; h++) {
    if (String(data[h][1] || '').trim().toUpperCase() === 'TIPO') {
      headerIndex = h;
      break;
    }
  }
  if (headerIndex < 0) return result;

  let currentArea = null;
  let currentObj = null;

  for (let i = headerIndex + 1; i < data.length; i++) {
    const row = data[i];
    const tipo = String(row[1] || '').trim().toUpperCase();
    const texto = String(row[2] || '').trim();
    const eval1T = String(row[3] || '').trim();
    const eval2T = String(row[4] || '').trim();
    const eval3T = String(row[5] || '').trim();
    const col6 = String(row[6] || '').trim();

    if (!tipo && !texto) continue;

    if (tipo === 'ÁREA' || tipo === 'AREA') {
      currentArea = { name: texto, objectives: [] };
      result.areas.push(currentArea);
      currentObj = null;
    } else if (tipo === 'OBJETIVO' && currentArea) {
      currentObj = {
        title: texto,
        indicators: [],
        contents: [],
        activities: ''
      };
      currentArea.objectives.push(currentObj);
    } else if (tipo === 'VALORACIÓN INICIAL') {
      result.valoracionInicial = texto;
    } else if (tipo === 'SEGUIMIENTO 1T') {
      result.seguimiento1T = texto;
    } else if (tipo === 'SEGUIMIENTO 2T') {
      result.seguimiento2T = texto;
    } else if (tipo === 'SEGUIMIENTO 3T') {
      result.seguimiento3T = texto;
    } else if (tipo === 'INFORME_INDICADOR') {
      result.informeFamilias.indicadores.push({ text: texto });
      result.informeFamilias.evaluaciones.push({ eval1T, eval2T, eval3T });
    } else if (currentObj) {
      const item = { text: texto, eval1T, eval2T, eval3T, observaciones: col6 };
      if (tipo === 'INDICADOR') {
        currentObj.indicators.push(item);
      } else if (tipo === 'CONTENIDO') {
        if (texto) currentObj.contents.push({ text: texto });
      } else if (tipo === 'ACTIVIDAD') {
        currentObj.activities = currentObj.activities
          ? (currentObj.activities + (texto ? '<br>' + texto : ''))
          : texto;
      }
    }
  }

  return result;
}

/* ───────── WRITE: guardar datos de un alumno ───────── */

function saveStudentData(payload) {
  assertAuthorized_();
  migrateMasterIfNeeded_();
  const data = JSON.parse(payload);
  const yearId = data.yearId || getActiveYearId_();
  const ss = getYearSS_(yearId);
  const tabName = String(data.studentName || '').trim();
  if (!tabName) throw new Error('El nombre del alumno no puede estar vacío.');
  const originalName = String(data.originalStudentName || '').trim();
  assertValidStudentName_(tabName);
  if (originalName) assertValidStudentName_(originalName);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Otro usuario está guardando. Reintenta.');
  try {
    const isRename = !!(originalName && originalName !== tabName);
    let sheet = ss.getSheetByName(isRename ? originalName : tabName);
    if (isRename && ss.getSheetByName(tabName)) {
      throw new Error('Ya existe un alumno con el nombre "' + tabName + '".');
    }
    // Comprobación de versión optimista (antes de renombrar nada): si la pestaña ya
    // existe y el cliente envía expectedUpdatedAt, debe coincidir con el actual de la hoja.
    if (sheet && data.expectedUpdatedAt !== undefined && data.expectedUpdatedAt !== null) {
      const last = sheet.getLastColumn() >= 12 ? sheet.getRange(1, 1, 1, 14).getValues()[0] : [];
      let currentUpdatedAt = '';
      if (String(last[10] || '').trim().toUpperCase() === 'UPDATED_AT') {
        currentUpdatedAt = String(last[11] || '').trim();
      }
      if (currentUpdatedAt && String(data.expectedUpdatedAt) !== currentUpdatedAt) {
        const byRaw = (String(last[12] || '').trim().toUpperCase() === 'UPDATED_BY') ? String(last[13] || '').trim() : '';
        const by = byRaw ? (' por ' + byRaw) : '';
        throw new Error('CONFLICT: otra persona ha guardado cambios' + by + ' mientras editabas. Recarga el alumno y vuelve a aplicar tus cambios.');
      }
    }
    // Renombrado: renombrar la pestaña y quitar la fila antigua del Índice
    if (isRename && sheet) {
      sheet.setName(tabName);
      const indice = getOrCreateIndiceIn_(ss);
      const indData = indice.getDataRange().getValues();
      for (let i = indData.length - 1; i >= 1; i--) {
        if (String(indData[i][0]).trim() === originalName) {
          indice.deleteRow(i + 1);
          break;
        }
      }
    }
    const isNewSheet = !sheet;
    if (sheet) {
      sheet.clear();
    } else {
      sheet = ss.insertSheet(tabName);
    }

    const NUM_COLS = 7;
    const pad = function(row) {
      while (row.length < NUM_COLS) row.push('');
      return row;
    };

    const areaNames = (data.areas || []).map(function(a) { return a.name; }).join(', ');
    const rows = [];

    rows.push(['ALUMNO/A', data.studentName, 'CURSO', data.course, 'PROGRAMA', 'PE', 'ÁMBITOS']);
    rows.push(['', 'TIPO', 'TEXTO', '1T', '2T', '3T', 'OBSERVACIONES']);

    const areas = data.areas || [];
    for (let a = 0; a < areas.length; a++) {
      const area = areas[a];
      rows.push(pad(['', 'ÁREA', area.name]));

      const objectives = area.objectives || [];
      for (let i = 0; i < objectives.length; i++) {
        const obj = objectives[i];
        const objLabel = 'Obj. ' + (i + 1);

        rows.push(pad([objLabel, 'OBJETIVO', obj.title || '']));

        (obj.indicators || []).forEach(function(ind) {
          if (ind.text && ind.text.trim()) {
            rows.push([objLabel, 'INDICADOR', ind.text.trim(),
              ind.eval1T || '', ind.eval2T || '', ind.eval3T || '', ind.observaciones || '']);
          }
        });

        const contentsArr = Array.isArray(obj.contents)
          ? obj.contents
          : (typeof obj.contents === 'string' && obj.contents.trim() ? [{ text: obj.contents }] : []);
        contentsArr.forEach(function(cnt) {
          if (cnt && cnt.text && String(cnt.text).trim()) {
            rows.push(pad([objLabel, 'CONTENIDO', cnt.text]));
          }
        });

        const activitiesHtml = typeof obj.activities === 'string'
          ? obj.activities
          : (Array.isArray(obj.activities) ? obj.activities.map(function(a) { return a && a.text ? a.text : ''; }).filter(function(t) { return t; }).join('<br>') : '');
        if (activitiesHtml && activitiesHtml.trim()) {
          rows.push(pad([objLabel, 'ACTIVIDAD', activitiesHtml]));
        }
      }

      if (a < areas.length - 1) {
        rows.push(pad(['']));
      }
    }

    rows.push(pad(['']));
    if (data.valoracionInicial) rows.push(pad(['', 'VALORACIÓN INICIAL', data.valoracionInicial]));
    if (data.seguimiento1T) rows.push(pad(['', 'SEGUIMIENTO 1T', data.seguimiento1T]));
    if (data.seguimiento2T) rows.push(pad(['', 'SEGUIMIENTO 2T', data.seguimiento2T]));
    if (data.seguimiento3T) rows.push(pad(['', 'SEGUIMIENTO 3T', data.seguimiento3T]));

    var informe = data.informeFamilias;
    if (informe && informe.indicadores && informe.indicadores.length > 0) {
      rows.push(pad(['']));
      for (var fi = 0; fi < informe.indicadores.length; fi++) {
        var indText = informe.indicadores[fi].text || '';
        var ev = (informe.evaluaciones && informe.evaluaciones[fi]) || {};
        if (indText.trim()) {
          rows.push(['', 'INFORME_INDICADOR', indText.trim(),
            ev.eval1T || '', ev.eval2T || '', ev.eval3T || '', '']);
        }
      }
    }

    const docentesStr = serializeDocentes_(data.docentes || []);
    const newUpdatedAt = new Date().toISOString();
    const userEmail = getCurrentUserEmail_();
    if (rows.length > 0) {
      const safeRows = rows.map(function(r) { return r.map(safeCell_); });
      sheet.getRange(1, 1, safeRows.length, NUM_COLS).setValues(safeRows);
      sheet.getRange(1, 8, 1, 7).setValues([[
        safeCell_(areaNames), 'DOCENTES', safeCell_(docentesStr),
        'UPDATED_AT', newUpdatedAt, 'UPDATED_BY', userEmail
      ]]);
    }

    formatStudentSheet_(sheet, rows, isNewSheet);
    updateIndiceIn_(ss, data.studentName, data.course, 'PE', areaNames, docentesStr);
    return { success: true, message: 'Datos guardados correctamente', updatedAt: newUpdatedAt, updatedBy: userEmail };
  } finally {
    lock.releaseLock();
  }
}

function getCurrentUserEmail_() {
  try {
    var e = Session.getActiveUser().getEmail();
    return e || '';
  } catch (err) {
    return '';
  }
}

/* ───────── DELETE: eliminar alumno ───────── */

function deleteStudent(studentName, yearId) {
  assertAuthorized_();
  migrateMasterIfNeeded_();
  studentName = String(studentName || '').trim();
  if (!studentName) throw new Error('Falta el nombre del alumno.');
  assertValidStudentName_(studentName);
  const ss = getYearSS_(yearId);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('Otro usuario está guardando. Reintenta.');
  try {
    const sheet = ss.getSheetByName(studentName);
    if (sheet) ss.deleteSheet(sheet);

    const indice = getOrCreateIndiceIn_(ss);
    const data = indice.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]).trim() === studentName) {
        indice.deleteRow(i + 1);
        break;
      }
    }
  } finally {
    lock.releaseLock();
  }
  return { success: true };
}

/* ───────── Helpers: format & index ───────── */

function formatStudentSheet_(sheet, rowsData, isNewSheet) {
  const NUM_COLS = 7;

  // clear() no restablece anchos de columna: solo se fijan al crear la pestaña.
  if (isNewSheet) {
    sheet.setColumnWidth(1, 80);
    sheet.setColumnWidth(2, 150);
    sheet.setColumnWidth(3, 450);
    sheet.setColumnWidths(4, 3, 50);
    sheet.setColumnWidth(7, 300);
  }

  sheet.setFrozenRows(2);

  const totalRows = rowsData.length;
  if (totalRows < 1) return;

  const backgrounds = [];
  const fontColors = [];
  const fontWeights = [];
  const wraps = [];

  const fillRow = function(bg, color, weight, wrap) {
    const bgRow = [], fcRow = [], fwRow = [], wrRow = [];
    for (let c = 0; c < NUM_COLS; c++) {
      bgRow.push(bg);
      fcRow.push(color);
      fwRow.push(weight);
      wrRow.push(wrap);
    }
    return { bg: bgRow, fc: fcRow, fw: fwRow, wr: wrRow };
  };

  const source = rowsData;

  for (let i = 0; i < totalRows; i++) {
    let bg = null, fc = null, fw = 'normal', wrap = false;
    if (i === 0) {
      bg = null; fc = null; fw = 'bold';
    } else if (i === 1) {
      bg = '#2d6a4f'; fc = '#ffffff'; fw = 'bold';
    } else {
      const tipo = String((source[i] && source[i][1]) || '').trim().toUpperCase();
      if (tipo === 'ÁREA' || tipo === 'AREA') {
        bg = '#1b4332'; fc = '#ffffff'; fw = 'bold';
      } else if (tipo === 'OBJETIVO') {
        bg = '#d1fae5'; fw = 'bold';
      } else if (tipo === 'INDICADOR') {
        bg = '#fef3c7';
      } else if (tipo === 'CONTENIDO') {
        bg = '#ede9fe';
      } else if (tipo === 'ACTIVIDAD') {
        bg = '#dbeafe';
      } else if (tipo.indexOf('VALORACIÓN') === 0 || tipo.indexOf('SEGUIMIENTO') === 0) {
        bg = '#f3f4f6'; fw = 'bold'; wrap = true;
      } else if (tipo === 'INFORME_INDICADOR') {
        bg = '#fce7f3';
      }
    }
    const f = fillRow(bg, fc, fw, wrap);
    backgrounds.push(f.bg);
    fontColors.push(f.fc);
    fontWeights.push(f.fw);
    wraps.push(f.wr);
  }

  const fullRange = sheet.getRange(1, 1, totalRows, NUM_COLS);
  fullRange.setBackgrounds(backgrounds);
  fullRange.setFontColors(fontColors);
  fullRange.setFontWeights(fontWeights);
  fullRange.setWraps(wraps);

  const metaWeights = [['bold', 'normal', 'bold', 'normal', 'bold', 'normal', 'bold', 'normal', 'bold', 'normal', 'bold', 'normal', 'bold', 'normal']];
  sheet.getRange(1, 1, 1, 14).setFontWeights(metaWeights);
}

/* ───────── Gestión de cursos académicos ───────── */

function cloneSchoolYear(payload) {
  assertAuthorized_();
  migrateMasterIfNeeded_();
  const data = JSON.parse(payload);
  const sourceId = data.sourceYearId || getActiveYearId_();
  const newLabel = String(data.newLabel || '').trim();
  if (!sourceId) throw new Error('No hay curso académico de origen.');
  if (!newLabel) throw new Error('Indica el nombre del nuevo curso académico.');

  assertKnownYear_(sourceId);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(60000)) throw new Error('Otra operación en curso. Reintenta.');
  try {
    // Validar (dentro del lock) que no existe ya un año con esa etiqueta
    const existing = readCursosRows_(true);
    if (existing.some(function(y) { return y.label.toLowerCase() === newLabel.toLowerCase(); })) {
      throw new Error('Ya existe un curso académico con ese nombre.');
    }
    const safeLabel = newLabel.replace(/\//g, '-');
    const fileName = 'Programa Atención a la Diversidad — ' + safeLabel;

    const sourceFile = DriveApp.getFileById(sourceId);
    const parentFolder = getMasterParentFolder_();
    const newFile = sourceFile.makeCopy(fileName, parentFolder);
    const newId = newFile.getId();
    const newUrl = newFile.getUrl();

    const newSS = SpreadsheetApp.openById(newId);

    // Actualizar cursoEscolar en la Config del nuevo año (mantiene courses/docentes heredados)
    const newCfg = getOrCreateConfigIn_(newSS);
    upsertKV_(newCfg, 'cursoEscolar', newLabel);

    // Limpiar evaluaciones, observaciones, valoración, seguimientos e informes
    cleanYearSpreadsheet_(newSS);

    // Registrar en Cursos del maestro
    const cursos = getOrCreateCursos_();
    cursos.appendRow([safeCell_(newLabel), newId, newUrl, new Date().toISOString(), 'FALSE']);
    invalidateCursosCache_();

    return { id: newId, url: newUrl, label: newLabel };
  } finally {
    lock.releaseLock();
  }
}

function cleanYearSpreadsheet_(ss) {
  const sheets = ss.getSheets();
  sheets.forEach(function(sh) {
    const name = sh.getName();
    if (name === CONFIG_TAB || name === INDICE_TAB) return;
    cleanStudentSheet_(sh);
  });
}

function cleanStudentSheet_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  const range = sh.getRange(1, 1, lastRow, 7);
  const values = range.getValues();
  let headerIdx = -1;
  for (let h = 1; h < values.length; h++) {
    if (String(values[h][1] || '').trim().toUpperCase() === 'TIPO') {
      headerIdx = h;
      break;
    }
  }
  if (headerIdx < 0) return;
  for (let i = headerIdx + 1; i < values.length; i++) {
    const tipo = String(values[i][1] || '').trim().toUpperCase();
    if (tipo === 'INDICADOR' || tipo === 'INFORME_INDICADOR') {
      values[i][3] = '';
      values[i][4] = '';
      values[i][5] = '';
      values[i][6] = '';
    } else if (tipo === 'VALORACIÓN INICIAL' || tipo.indexOf('SEGUIMIENTO') === 0) {
      values[i][2] = '';
    }
  }
  range.setValues(values);
}

function archiveYear(yearId, archived) {
  assertAuthorized_();
  migrateMasterIfNeeded_();
  if (!yearId) throw new Error('Falta el id del curso académico.');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('Otro usuario está guardando. Reintenta.');
  try {
    const sheet = getOrCreateCursos_();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === String(yearId).trim()) {
        sheet.getRange(i + 1, 5).setValue(archived ? 'TRUE' : 'FALSE');
        invalidateCursosCache_();
        return { success: true };
      }
    }
    throw new Error('No se ha encontrado el curso académico.');
  } finally {
    lock.releaseLock();
  }
}

function deleteYear(yearId, confirmLabel) {
  assertAuthorized_();
  migrateMasterIfNeeded_();
  if (!yearId) throw new Error('Falta el id del curso académico.');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Otro usuario está guardando. Reintenta.');
  try {
    const sheet = getOrCreateCursos_();
    const data = sheet.getDataRange().getValues();
    let foundRow = -1;
    let label = '';
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === String(yearId).trim()) {
        foundRow = i + 1;
        label = String(data[i][0]).trim();
        break;
      }
    }
    if (foundRow < 0) throw new Error('No se ha encontrado el curso académico.');
    if (String(confirmLabel || '').trim() !== label) {
      throw new Error('El nombre escrito no coincide con el del curso académico.');
    }
    // Validar que no es el único año
    let total = 0;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim()) total++;
    }
    if (total <= 1) {
      throw new Error('No puedes borrar el único curso académico. Crea otro antes.');
    }
    // Mover a papelera de Drive (recuperable durante 30 días)
    try {
      DriveApp.getFileById(yearId).setTrashed(true);
    } catch (e) {
      // El archivo ya pudo ser borrado manualmente — seguimos para limpiar el registro
    }
    sheet.deleteRow(foundRow);
    invalidateCursosCache_();
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

/* ───────── Presencia (avisar de ediciones simultáneas) ───────── */

// Presencia en CacheService (sin hoja ni LockService): cada alumno tiene una entrada
// { email: timestamp } que caduca sola. Una escritura concurrente puede perder una
// entrada, pero se restablece en el siguiente latido (cada 60 s).
function presenceKey_(yearId, tabName) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, yearId + '|' + tabName, Utilities.Charset.UTF_8);
  return 'presence:' + Utilities.base64EncodeWebSafe(digest);
}

function updatePresence_(yearId, tabName, doRegister) {
  if (!yearId || !tabName) return [];
  const cache = CacheService.getScriptCache();
  const key = presenceKey_(String(yearId), String(tabName));
  const myEmail = getCurrentUserEmail_();
  const nowMs = Date.now();
  let entries = {};
  try { entries = JSON.parse(cache.get(key) || '{}') || {}; } catch (e) { entries = {}; }

  const others = [];
  const next = {};
  Object.keys(entries).forEach(function(email) {
    const ts = Number(entries[email] || 0);
    if (!ts || (nowMs - ts) > PRESENCE_TTL_MS) return; // caducada
    if (email === myEmail) return;                       // la propia se reescribe abajo
    next[email] = ts;
    others.push({ email: email || 'Usuario sin identificar', ts: ts });
  });
  if (doRegister) next[myEmail] = nowMs;

  if (Object.keys(next).length) {
    cache.put(key, JSON.stringify(next), Math.ceil(PRESENCE_TTL_MS / 1000));
  } else {
    cache.remove(key);
  }
  return others;
}

function acquirePresence(yearId, tabName) {
  assertAuthorized_();
  return updatePresence_(yearId, tabName, true);
}

function heartbeatPresence(yearId, tabName) {
  assertAuthorized_();
  return updatePresence_(yearId, tabName, true);
}

function releasePresence(yearId, tabName) {
  assertAuthorized_();
  return updatePresence_(yearId, tabName, false);
}

function updateIndiceIn_(ss, name, course, program, area, docentesStr) {
  const sheet = getOrCreateIndiceIn_(ss);
  const data = sheet.getDataRange().getValues();
  const dStr = docentesStr || '';
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === name.trim()) {
      sheet.getRange(i + 1, 1, 1, 5).setValues([[name, course, program, area, dStr].map(safeCell_)]);
      return;
    }
  }
  sheet.appendRow([name, course, program, area, dStr].map(safeCell_));
}
