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

// Icono de la pestaña: pieza de puzle 🧩 (Twemoji, CC-BY 4.0)
const FAVICON_URL = 'https://cdn.jsdelivr.net/gh/twitter/twemoji@14.0.2/assets/72x72/1f9e9.png';

function doGet() {
  if (!isAuthorized_()) {
    const email = getCurrentUserEmail_();
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;max-width:520px;margin:15vh auto;padding:0 16px;color:#1b4332">' +
      '<h2>Acceso no autorizado</h2>' +
      '<p>La cuenta <b>' + escapeHtml_(email || 'desconocida') + '</b> no está en la lista de usuarios autorizados.</p>' +
      '<p>Pide al administrador de la aplicación que te añada desde Ajustes.</p></div>')
      .setTitle('Acceso no autorizado')
      .setFaviconUrl(FAVICON_URL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Programas de Atención a la Diversidad')
    .setFaviconUrl(FAVICON_URL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ───────── Autorización (lista de usuarios) ───────── */
//  - Administrador: la cuenta que despliega el script (Session.getEffectiveUser).
//  - Usuarios autorizados: claves "usuario" en la pestaña Config del maestro.
//  - Si la lista está vacía, se permite el acceso a todo el dominio (comportamiento previo).

const USERS_KEY = 'usuario';

/* ───────── Equipo (acceso + docentes asignables) ─────────
 * Pestaña «Equipo» de la hoja maestra: NOMBRE | ESPECIALIDAD | CORREO | ADMIN.
 *  - Solo entran los correos del equipo (y los administradores del script).
 *    Si el equipo no tiene ningún correo, entra cualquier cuenta del dominio.
 *  - Quien tenga ADMIN marcado gestiona Ajustes, el equipo y los cursos académicos,
 *    así la app sigue funcionando aunque falte quien la desplegó.
 *  - Los docentes asignables a una ficha son los miembros con nombre: «Nombre (Especialidad)».
 */
const TEAM_TAB = 'Equipo';
const TEAM_TAB_ALT = 'Equipo (app)';
const TEAM_CACHE_KEY = 'team.v1';

function teamLabel_(m) {
  return m.nombre + (m.rol ? ' (' + m.rol + ')' : '');
}

function readTeamRows_(sheet) {
  const data = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const nombre = String(data[i][0] || '').trim();
    const rol = String(data[i][1] || '').trim();
    const email = normEmail_(data[i][2]);
    if (!nombre && !email) continue;
    const adm = String(data[i][3] === undefined ? '' : data[i][3]).trim().toUpperCase();
    out.push({ nombre: nombre, rol: rol, email: email, admin: !(adm === 'FALSE' || adm === 'NO' || adm === '0') });
  }
  return out;
}

function writeTeamRows_(sheet, team) {
  const rows = [['NOMBRE', 'ESPECIALIDAD', 'CORREO', 'ADMIN']].concat(team.map(function(m) {
    return [safeCell_(m.nombre), safeCell_(m.rol), m.email, m.admin ? 'SÍ' : 'NO'];
  }));
  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, 4).setValues(rows);
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

// Crea la pestaña Equipo la primera vez, a partir de los docentes del año activo
// («Nombre (PT)») y de la antigua lista de usuarios autorizados (solo correo).
// Busca la pestaña del equipo sin distinguir mayúsculas/espacios. Una pestaña «Equipo» que
// no sea de la app (otra cabecera) no se toca nunca: en ese caso se usa «Equipo (app)».
function findTeamSheet_(ss) {
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const n = sheets[i].getName().trim().toLowerCase();
    if (n !== TEAM_TAB.toLowerCase() && n !== TEAM_TAB_ALT.toLowerCase()) continue;
    const a1 = sheets[i].getLastRow() ? String(sheets[i].getRange(1, 1).getValue()).trim().toUpperCase() : '';
    if (!a1 || a1 === 'NOMBRE') return sheets[i];
  }
  return null;
}

function getOrCreateTeamSheet_() {
  const ss = getMasterSS_();
  let sheet = findTeamSheet_(ss);
  if (sheet) return sheet;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('Otro usuario está guardando. Reintenta.');
  try {
    SpreadsheetApp.flush();
    sheet = findTeamSheet_(ss);
    if (sheet) return sheet;
    const team = [];
    try {
      const yearId = getActiveYearId_();
      if (yearId) {
        const yCfg = readKeyValueSheet_(SpreadsheetApp.openById(yearId).getSheetByName(CONFIG_TAB));
        const docs = Array.isArray(yCfg.docente) ? yCfg.docente : (yCfg.docente ? [yCfg.docente] : []);
        docs.forEach(function(d) {
          const m = String(d).match(/^(.*?)\s*\(([^)]*)\)\s*$/);
          team.push({ nombre: (m ? m[1] : String(d)).trim(), rol: m ? m[2].trim() : '', email: '', admin: true });
        });
      }
    } catch (e) { /* sin año activo */ }
    const cfg = readKeyValueSheet_(ss.getSheetByName(CONFIG_TAB));
    const raw = cfg[USERS_KEY];
    (Array.isArray(raw) ? raw : (raw ? [raw] : [])).map(normEmail_).filter(function(e) { return e; })
      .forEach(function(e) { team.push({ nombre: '', rol: '', email: e, admin: true }); });
    sheet = insertTeamSheet_(ss);
    // Solo se escribe en una pestaña recién creada y vacía: nunca sobre datos existentes
    if (!sheet.getLastRow()) writeTeamRows_(sheet, team);
    return sheet;
  } finally {
    lock.releaseLock();
  }
}

function insertTeamSheet_(ss) {
  const names = [TEAM_TAB, TEAM_TAB_ALT];
  for (let i = 0; i < names.length; i++) {
    try {
      return ss.insertSheet(names[i]);
    } catch (e) {
      // El nombre ya existe (otra ejecución acaba de crearla o es una pestaña ajena): reintentar
      SpreadsheetApp.flush();
      const found = findTeamSheet_(ss);
      if (found) return found;
    }
  }
  throw new Error('No se pudo crear la pestaña del equipo en la hoja maestra.');
}

function getTeam_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(TEAM_CACHE_KEY);
  if (cached !== null) {
    try { return JSON.parse(cached); } catch (e) { /* recalcular */ }
  }
  const team = readTeamRows_(getOrCreateTeamSheet_());
  cache.put(TEAM_CACHE_KEY, JSON.stringify(team), 300);
  return team;
}

function escapeHtml_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function normEmail_(e) {
  return String(e || '').trim().toLowerCase();
}

// Administradores (siempre tienen acceso y gestionan la lista):
//  1) Propiedad del script ADMIN_EMAILS (correos separados por comas), si existe.
//  2) La cuenta que despliega (Session.getEffectiveUser), si Apps Script la devuelve.
//  3) El propietario de la hoja maestra en Drive (respaldo: getEffectiveUser puede venir vacío).
// Se cachea 6 h para no consultar Drive en cada llamada.
const ADMINS_CACHE_KEY = 'adminEmails.v1';

function getAdminEmails_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(ADMINS_CACHE_KEY);
  if (cached !== null) {
    try { return JSON.parse(cached); } catch (e) { /* recalcular */ }
  }
  const set = {};
  const add = function(e) { e = normEmail_(e); if (e) set[e] = true; };
  try {
    String(PropertiesService.getScriptProperties().getProperty('ADMIN_EMAILS') || '')
      .split(',').forEach(add);
  } catch (e) { /* sin propiedad */ }
  try { add(Session.getEffectiveUser().getEmail()); } catch (e) { /* vacío */ }
  try {
    const owner = DriveApp.getFileById(SS_ID).getOwner();
    if (owner) add(owner.getEmail());
  } catch (e) { /* unidad compartida o sin permiso */ }
  const list = Object.keys(set);
  if (list.length) cache.put(ADMINS_CACHE_KEY, JSON.stringify(list), 21600);
  return list;
}

function getAdminEmail_() {
  return getAdminEmails_().join(', ');
}

function getAllowedUsers_() {
  return getTeam_().map(function(m) { return m.email; }).filter(function(e) { return e; });
}

function isScriptAdmin_() {
  const me = normEmail_(getCurrentUserEmail_());
  return !!me && getAdminEmails_().indexOf(me) !== -1;
}

function isAdmin_() {
  if (isScriptAdmin_()) return true;
  const me = normEmail_(getCurrentUserEmail_());
  const team = getTeam_();
  // Mientras ningún admin del equipo tenga correo (p. ej. justo tras la migración), cualquier
  // usuario autorizado puede administrar: así nadie se queda sin poder configurar el equipo.
  if (!team.some(function(m) { return m.admin && m.email; })) {
    const emails = team.map(function(m) { return m.email; }).filter(function(e) { return e; });
    return !emails.length || (!!me && emails.indexOf(me) !== -1);
  }
  return !!me && team.some(function(m) { return m.admin && m.email === me; });
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

function getTeam() {
  assertAuthorized_();
  return { owner: getAdminEmail_(), me: normEmail_(getCurrentUserEmail_()), team: getTeam_(), canEdit: isAdmin_() };
}

function saveTeam(payload) {
  assertAdmin_();
  const data = JSON.parse(payload);
  const seen = {};
  const team = [];
  (data.team || []).forEach(function(m) {
    const nombre = String(m.nombre || '').trim();
    const rol = String(m.rol || '').trim();
    const email = normEmail_(m.email);
    if (!nombre && !email) return;
    if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Correo no válido: ' + email);
    if (email && seen[email]) throw new Error('Correo repetido: ' + email);
    if (email) seen[email] = true;
    team.push({ nombre: nombre, rol: rol, email: email, admin: m.admin !== false });
  });
  // Evita que alguien se quede fuera a sí mismo (los administradores del script siempre entran)
  const me = normEmail_(getCurrentUserEmail_());
  if (!isScriptAdmin_() && !team.some(function(m) { return m.email === me && m.admin; })) {
    throw new Error('No puedes quitarte a ti mismo del equipo ni tu permiso de administrador.');
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('Otro usuario está guardando. Reintenta.');
  try {
    writeTeamRows_(getOrCreateTeamSheet_(), team);
    CacheService.getScriptCache().remove(TEAM_CACHE_KEY);
  } finally {
    lock.releaseLock();
  }
  return { success: true, team: team };
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
  built.config.appUrl = ScriptApp.getService().getUrl();
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
      docentes: teamDocentes_(docentes),
      isAdmin: isAdmin_(),
      adminEmail: getAdminEmail_(),
      currentUser: normEmail_(getCurrentUserEmail_())
    }
  };
}

// Docentes asignables: los del Equipo con nombre; si aún no hay ninguno, la lista antigua del año
function teamDocentes_(legacy) {
  const fromTeam = getTeam_().filter(function(m) { return m.nombre; }).map(teamLabel_);
  return fromTeam.length ? fromTeam : legacy;
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
        const newLabel = String(data.cursoEscolar).trim();
        const yearSS = SpreadsheetApp.openById(yearId);
        const yCfg = getOrCreateConfigIn_(yearSS);
        upsertKV_(yCfg, 'cursoEscolar', newLabel);
        if (newLabel) renameYearLabel_(yearId, newLabel);
      }
    }
  } finally {
    lock.releaseLock();
  }
  return { success: true };
}

// Renombra el curso académico en la lista de cursos (selector) y el archivo de Drive
function renameYearLabel_(yearId, newLabel) {
  const sheet = getOrCreateCursos_();
  const data = sheet.getDataRange().getValues();
  let row = -1;
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][1]).trim();
    if (id === String(yearId).trim()) row = i + 1;
    else if (id && String(data[i][0]).trim().toLowerCase() === newLabel.toLowerCase()) {
      throw new Error('Ya existe un curso académico llamado "' + newLabel + '".');
    }
  }
  if (row < 0 || String(data[row - 1][0]).trim() === newLabel) return;
  sheet.getRange(row, 1).setValue(safeCell_(newLabel));
  invalidateCursosCache_();
  try {
    DriveApp.getFileById(yearId).setName('Programa Atención a la Diversidad — ' + newLabel.replace(/\//g, '-'));
  } catch (e) { /* sin permiso para renombrar el archivo: no es imprescindible */ }
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
    sheet.appendRow(['ALUMNO/A', 'CURSO', 'PROGRAMA', 'ÁREA/ÁMBITO', 'DOCENTES', 'UNIDAD', 'FIN VIGENCIA']);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
    sheet.setFrozenRows(1);
    return sheet;
  }
  // Migración: si la cabecera no tiene DOCENTES, añadirla
  const lastCol = sheet.getLastColumn();
  if (lastCol < 5) {
    sheet.getRange(1, 5).setValue('DOCENTES').setFontWeight('bold');
  }
  if (lastCol < 7) {
    sheet.getRange(1, 6, 1, 2).setValues([['UNIDAD', 'FIN VIGENCIA']]).setFontWeight('bold');
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
        docentes: parseDocentesString_(String(row[4] || '')),
        unidad: String(row[5] || '').trim(),
        finVigencia: dateCell_(row[6])
      });
    }
  }
  return students;
}

// Sheets convierte "2026-08-31" en fecha: se devuelve siempre como texto yyyy-MM-dd
function dateCell_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '').trim();
}

// Datos generales del PE que se guardan como filas TIPO/TEXTO al final de la pestaña
const PE_FIELDS = [
  ['UNIDAD', 'unidad'],
  ['FECHA ELABORACIÓN', 'fechaElaboracion'],
  ['FIN VIGENCIA', 'fechaFinVigencia'],
  ['PROFESIONALES', 'profesionales']
];

// Organización y recursos de cada ámbito, los mismos dos campos que tiene Séneca
// (filas TIPO/TEXTO tras sus objetivos)
const ORG_FIELDS = [
  ['ORGANIZACIÓN', 'organizacion'],
  ['RECURSOS', 'recursos']
];

function emptyOrg_() {
  const o = {};
  ORG_FIELDS.forEach(function(f) { o[f[1]] = ''; });
  return o;
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

// withPresence: además registra la presencia del usuario en esa ficha y devuelve quién
// más la está editando (ahorra la llamada acquirePresence al abrir un alumno).
function getStudentData(studentName, yearId, withPresence) {
  assertAuthorized_();
  migrateMasterIfNeeded_();
  assertValidStudentName_(studentName);
  const ss = getYearSS_(yearId);
  const data = readStudentData_(ss, studentName);
  if (data && withPresence) {
    data.presence = updatePresence_(String(yearId || ss.getId()), String(studentName), true);
  }
  return data;
}

function readStudentData_(ss, studentName) {
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
    pe: { unidad: '', fechaElaboracion: '', fechaFinVigencia: '', profesionales: '' },
    informeFamilias: { indicadores: [], evaluaciones: [] }
  };
  const peByTipo = {};
  PE_FIELDS.forEach(function(f) { peByTipo[f[0]] = f[1]; });
  const orgByTipo = {};
  ORG_FIELDS.forEach(function(f) { orgByTipo[f[0]] = f[1]; });
  let legacySesiones = '';

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
      currentArea = { name: texto, programa: col6, org: emptyOrg_(), objectives: [] };
      result.areas.push(currentArea);
      currentObj = null;
    } else if (tipo === 'OBJETIVO' && currentArea) {
      currentObj = {
        title: texto,
        indicators: [],
        contents: [],
        activities: '',
        competencias: []
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
    } else if (orgByTipo[tipo] && currentArea) {
      currentArea.org[orgByTipo[tipo]] = texto;
    } else if (tipo === 'SESIONES') {
      legacySesiones = texto; // campo global antiguo: se reparte por ámbitos más abajo
    } else if (peByTipo[tipo]) {
      result.pe[peByTipo[tipo]] = dateCell_(row[2]);
    } else if (tipo === 'INFORME_INDICADOR') {
      result.informeFamilias.indicadores.push({ text: texto });
      result.informeFamilias.evaluaciones.push({ eval1T, eval2T, eval3T });
    } else if (currentObj) {
      const item = { text: texto, eval1T, eval2T, eval3T, observaciones: col6 };
      if (tipo === 'INDICADOR') {
        currentObj.indicators.push(item);
      } else if (tipo === 'CONTENIDO') {
        if (texto) currentObj.contents.push({ text: texto });
      } else if (tipo === 'COMPETENCIAS') {
        currentObj.competencias = texto.split('|').map(function(c) { return c.trim(); }).filter(function(c) { return c; });
      } else if (tipo === 'ACTIVIDAD') {
        currentObj.activities = currentObj.activities
          ? (currentObj.activities + (texto ? '<br>' + texto : ''))
          : texto;
      }
    }
  }

  // Migración: el antiguo «Sesiones y agrupamiento» global pasa a la organización de cada ámbito
  if (legacySesiones && !result.areas.some(function(a) { return a.org.organizacion; })) {
    result.areas.forEach(function(a) { a.org.organizacion = legacySesiones; });
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
      rows.push(pad(['', 'ÁREA', area.name, '', '', '', area.programa || '']));

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

        const comps = (obj.competencias || []).map(function(c) { return String(c || '').trim(); }).filter(function(c) { return c; });
        if (comps.length) rows.push(pad([objLabel, 'COMPETENCIAS', comps.join(' | ')]));
      }

      const org = area.org || {};
      ORG_FIELDS.forEach(function(f) {
        if (org[f[1]] && String(org[f[1]]).trim()) rows.push(pad(['', f[0], String(org[f[1]])]));
      });

      if (a < areas.length - 1) {
        rows.push(pad(['']));
      }
    }

    rows.push(pad(['']));
    if (data.valoracionInicial) rows.push(pad(['', 'VALORACIÓN INICIAL', data.valoracionInicial]));
    if (data.seguimiento1T) rows.push(pad(['', 'SEGUIMIENTO 1T', data.seguimiento1T]));
    if (data.seguimiento2T) rows.push(pad(['', 'SEGUIMIENTO 2T', data.seguimiento2T]));
    if (data.seguimiento3T) rows.push(pad(['', 'SEGUIMIENTO 3T', data.seguimiento3T]));

    const pe = data.pe || {};
    if (PE_FIELDS.some(function(f) { return pe[f[1]]; })) {
      rows.push(pad(['']));
      PE_FIELDS.forEach(function(f) {
        // El apóstrofo evita que Sheets convierta las fechas: se guardan como texto
        if (pe[f[1]]) rows.push(pad(['', f[0], (/^fecha/.test(f[1]) ? "'" : '') + String(pe[f[1]])]));
      });
    }

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
    updateIndiceIn_(ss, data.studentName, data.course, 'PE', areaNames, docentesStr, pe.unidad, pe.fechaFinVigencia);
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
  let trashName = '';
  try {
    // No se borra nada: la pestaña pasa a la papelera (oculta y renombrada) y se puede restaurar
    const sheet = ss.getSheetByName(studentName);
    if (sheet) {
      trashName = uniqueTrashName_(ss, studentName);
      sheet.setName(trashName);
      sheet.hideSheet();
    }

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
  return { success: true, trashName: trashName };
}

/* ───────── Papelera de alumnos ───────── */
const TRASH_PREFIX = '🗑 ';

function isTrashName_(name) { return String(name).indexOf(TRASH_PREFIX) === 0; }

// "🗑 Nombre · 2026-10-11 18:05" (único y dentro del límite de 100 caracteres de Sheets)
function uniqueTrashName_(ss, studentName) {
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  const base = TRASH_PREFIX + String(studentName).slice(0, 70) + ' · ' + stamp;
  let name = base, n = 2;
  while (ss.getSheetByName(name)) name = base + ' (' + (n++) + ')';
  return name;
}

function trashOriginalName_(trashName) {
  return String(trashName).slice(TRASH_PREFIX.length).replace(/ · \d{4}-\d{2}-\d{2} \d{2}:\d{2}( \(\d+\))?$/, '');
}

function listTrash(yearId) {
  assertAuthorized_();
  const ss = getYearSS_(yearId);
  return ss.getSheets().map(function(sh) { return sh.getName(); }).filter(isTrashName_).map(function(n) {
    return { trashName: n, name: trashOriginalName_(n), deletedAt: n.slice(n.lastIndexOf(' · ') + 3) };
  }).reverse();
}

function restoreStudent(trashName, yearId) {
  assertAuthorized_();
  if (!isTrashName_(trashName)) throw new Error('No es un elemento de la papelera.');
  const ss = getYearSS_(yearId);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('Otro usuario está guardando. Reintenta.');
  try {
    const sheet = ss.getSheetByName(trashName);
    if (!sheet) throw new Error('Ya no está en la papelera.');
    const name = trashOriginalName_(trashName);
    if (ss.getSheetByName(name)) throw new Error('Ya existe un alumno llamado "' + name + '". Cámbiale el nombre antes de restaurar.');
    sheet.setName(name);
    sheet.showSheet();
    const data = readStudentData_(ss, name) || { course: '', areas: [], docentes: [], pe: {} };
    const areaNames = (data.areas || []).map(function(a) { return a.name; }).join(', ');
    const pe = data.pe || {};
    updateIndiceIn_(ss, name, data.course, 'PE', areaNames, serializeDocentes_(data.docentes), pe.unidad, pe.fechaFinVigencia);
    return { name: name, course: data.course || '', program: 'PE', area: areaNames, docentes: data.docentes || [],
             unidad: pe.unidad || '', finVigencia: pe.fechaFinVigencia || '' };
  } finally {
    lock.releaseLock();
  }
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
      } else if (PE_FIELDS.some(function(f) { return f[0] === tipo; })) {
        bg = '#e0f2fe'; wrap = true;
      } else if (ORG_FIELDS.some(function(f) { return f[0] === tipo; })) {
        bg = '#f1f5f9'; wrap = true;
      } else if (tipo === 'COMPETENCIAS') {
        bg = '#ecfccb';
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
    if (isTrashName_(name)) { ss.deleteSheet(sh); return; } // la papelera no pasa al nuevo curso (sigue en el original)
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

// Comprobación ligera de sesión (el cliente la llama periódicamente y al volver a la pestaña).
function ping() {
  assertAuthorized_();
  return { ok: true, user: normEmail_(getCurrentUserEmail_()), ts: Date.now() };
}

// Guarda todos los Ajustes en una sola llamada (antes eran hasta tres).
function saveSettings(payload) {
  assertAuthorized_();
  const data = JSON.parse(payload);
  if (data.config) saveConfig(JSON.stringify(data.config));
  if (data.lists) saveYearLists(JSON.stringify(data.lists));
  if (data.team) saveTeam(JSON.stringify(data.team));
  return { success: true };
}

// Duplica la ficha de un alumno en el servidor (antes: leer en el cliente y volver a enviar).
function duplicateStudent(studentName, newName, yearId) {
  assertAuthorized_();
  migrateMasterIfNeeded_();
  assertValidStudentName_(studentName);
  const ss = getYearSS_(yearId);
  const data = readStudentData_(ss, studentName);
  if (!data) throw new Error('No se encontró la ficha original.');
  data.studentName = String(newName || '').trim();
  if (!data.studentName) throw new Error('Indica el nombre del nuevo alumno.');
  if (ss.getSheetByName(data.studentName)) throw new Error('Ya existe un alumno con el nombre "' + data.studentName + '".');
  data.yearId = yearId;
  delete data.updatedAt;
  delete data.updatedBy;
  saveStudentData(JSON.stringify(data));
  return {
    name: data.studentName,
    course: data.course || '',
    area: (data.areas || []).map(function(a) { return a.name; }).join(', '),
    docentes: data.docentes || [],
    unidad: (data.pe && data.pe.unidad) || '',
    finVigencia: (data.pe && data.pe.fechaFinVigencia) || ''
  };
}

function heartbeatPresence(yearId, tabName) {
  assertAuthorized_();
  return updatePresence_(yearId, tabName, true);
}

function releasePresence(yearId, tabName) {
  assertAuthorized_();
  return updatePresence_(yearId, tabName, false);
}

function updateIndiceIn_(ss, name, course, program, area, docentesStr, unidad, finVigencia) {
  const sheet = getOrCreateIndiceIn_(ss);
  const data = sheet.getDataRange().getValues();
  const row = [name, course, program, area, docentesStr || '', unidad || ''].map(safeCell_).concat([finVigencia ? "'" + finVigencia : '']);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === name.trim()) {
      sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return;
    }
  }
  sheet.appendRow(row);
}
