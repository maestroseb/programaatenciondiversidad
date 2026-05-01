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

/* ───────── Web App entry point ───────── */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Programas de Atención a la Diversidad')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
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

function readCursosRows_() {
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

function getActiveYearId_() {
  const rows = readCursosRows_();
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
  if (!yearId) yearId = getActiveYearId_();
  if (!yearId) throw new Error('No hay curso académico activo. Crea uno desde Ajustes.');
  return SpreadsheetApp.openById(yearId);
}

/* ───────── Migración del maestro (una vez) ───────── */

function migrateMasterIfNeeded_() {
  const ss = getMasterSS_();
  if (ss.getSheetByName(CURSOS_TAB)) return; // ya migrado

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
  } finally {
    lock.releaseLock();
  }
}

/* ───────── CONFIG: datos del centro y del año ───────── */

function getConfig(yearId) {
  migrateMasterIfNeeded_();

  // Maestro: centro/localidad + lista de años
  const ss = getMasterSS_();
  const masterCfg = readKeyValueSheet_(ss.getSheetByName(CONFIG_TAB));
  const years = readCursosRows_();

  // Año activo (resuelto a partir de yearId o más reciente no archivado)
  const resolvedId = (yearId && years.some(function(y) { return y.id === yearId; }))
    ? yearId
    : getActiveYearId_();

  let cursoEscolar = '';
  let courses = [];
  let docentes = [];

  if (resolvedId) {
    const yearSS = SpreadsheetApp.openById(resolvedId);
    const yearCfg = readKeyValueSheet_(yearSS.getSheetByName(CONFIG_TAB));
    cursoEscolar = yearCfg.cursoEscolar || autoCursoEscolar_();
    courses = Array.isArray(yearCfg.course) ? yearCfg.course : (yearCfg.course ? [yearCfg.course] : []);
    docentes = Array.isArray(yearCfg.docente) ? yearCfg.docente : (yearCfg.docente ? [yearCfg.docente] : []);
  } else {
    cursoEscolar = autoCursoEscolar_();
  }

  return {
    centro: masterCfg.centro || '',
    localidad: masterCfg.localidad || '',
    cursoEscolar: cursoEscolar,
    activeYearId: resolvedId || '',
    years: years,
    courses: courses,
    docentes: docentes
  };
}

function saveConfig(payload) {
  migrateMasterIfNeeded_();
  const data = JSON.parse(payload);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('Otro usuario está guardando. Reintenta.');
  try {
    const ss = getMasterSS_();
    let mCfg = ss.getSheetByName(CONFIG_TAB);
    if (!mCfg) {
      mCfg = ss.insertSheet(CONFIG_TAB);
      mCfg.appendRow(['CLAVE', 'VALOR']);
      mCfg.getRange(1, 1, 1, 2).setFontWeight('bold');
      mCfg.setFrozenRows(1);
    }
    upsertKV_(mCfg, 'centro', data.centro != null ? data.centro : '');
    upsertKV_(mCfg, 'localidad', data.localidad != null ? data.localidad : '');

    // cursoEscolar va al año activo (o al indicado)
    if (data.cursoEscolar !== undefined) {
      const yearId = data.yearId || getActiveYearId_();
      if (yearId) {
        const yearSS = SpreadsheetApp.openById(yearId);
        let yCfg = yearSS.getSheetByName(CONFIG_TAB);
        if (!yCfg) {
          yCfg = yearSS.insertSheet(CONFIG_TAB);
          yCfg.appendRow(['CLAVE', 'VALOR']);
          yCfg.getRange(1, 1, 1, 2).setFontWeight('bold');
          yCfg.setFrozenRows(1);
        }
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
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

function setMultiKV_(sheet, key, values) {
  // Elimina todas las filas con esa clave y reescribe una fila por valor.
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() === key) {
      sheet.deleteRow(i + 1);
    }
  }
  (values || []).forEach(function(v) {
    var s = String(v == null ? '' : v).trim();
    if (s) sheet.appendRow([key, s]);
  });
}

function saveYearLists(payload) {
  migrateMasterIfNeeded_();
  const data = JSON.parse(payload);
  const yearId = data.yearId || getActiveYearId_();
  if (!yearId) throw new Error('No hay curso académico activo.');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('Otro usuario está guardando. Reintenta.');
  try {
    const yearSS = SpreadsheetApp.openById(yearId);
    let cfg = yearSS.getSheetByName(CONFIG_TAB);
    if (!cfg) {
      cfg = yearSS.insertSheet(CONFIG_TAB);
      cfg.appendRow(['CLAVE', 'VALOR']);
      cfg.getRange(1, 1, 1, 2).setFontWeight('bold');
      cfg.setFrozenRows(1);
    }
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
  migrateMasterIfNeeded_();
  const ss = getYearSS_(yearId);
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
  migrateMasterIfNeeded_();
  const ss = getYearSS_(yearId);
  const sheet = ss.getSheetByName(studentName);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;

  // Row 1: metadata
  const meta = data[0];
  // Docentes pueden estar en col 9 (label) + col 10 (valor) si la pestaña ya se guardó con la versión nueva
  let docentesStr = '';
  if (String(meta[8] || '').trim().toUpperCase() === 'DOCENTES') {
    docentesStr = String(meta[9] || '').trim();
  }
  const result = {
    studentName: String(meta[1] || '').trim(),
    course: String(meta[3] || '').trim(),
    docentes: parseDocentesString_(docentesStr),
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
  migrateMasterIfNeeded_();
  const data = JSON.parse(payload);
  const yearId = data.yearId || getActiveYearId_();
  const ss = getYearSS_(yearId);
  const tabName = data.studentName.trim();

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Otro usuario está guardando. Reintenta.');
  try {
    let sheet = ss.getSheetByName(tabName);
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
    if (rows.length > 0) {
      sheet.getRange(1, 1, rows.length, NUM_COLS).setValues(rows);
      sheet.getRange(1, 8).setValue(areaNames);
      sheet.getRange(1, 9).setValue('DOCENTES');
      sheet.getRange(1, 10).setValue(docentesStr);
    }

    formatStudentSheet_(sheet, rows);
    updateIndiceIn_(ss, data.studentName, data.course, 'PE', areaNames, docentesStr);
  } finally {
    lock.releaseLock();
  }

  return { success: true, message: 'Datos guardados correctamente' };
}

/* ───────── DELETE: eliminar alumno ───────── */

function deleteStudent(studentName, yearId) {
  migrateMasterIfNeeded_();
  const ss = getYearSS_(yearId);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('Otro usuario está guardando. Reintenta.');
  try {
    const sheet = ss.getSheetByName(studentName);
    if (sheet) ss.deleteSheet(sheet);

    const indice = getOrCreateIndiceIn_(ss);
    const data = indice.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]).trim() === studentName.trim()) {
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

function formatStudentSheet_(sheet, rowsData) {
  const NUM_COLS = 7;

  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 450);
  sheet.setColumnWidth(4, 50);
  sheet.setColumnWidth(5, 50);
  sheet.setColumnWidth(6, 50);
  sheet.setColumnWidth(7, 300);

  sheet.setFrozenRows(2);

  const totalRows = rowsData ? rowsData.length : sheet.getLastRow();
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

  const source = rowsData || sheet.getRange(1, 1, totalRows, NUM_COLS).getValues();

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

  const metaWeights = [['bold', 'normal', 'bold', 'normal', 'bold', 'normal', 'bold', 'normal', 'bold', 'normal']];
  sheet.getRange(1, 1, 1, 10).setFontWeights(metaWeights);
}

function updateIndiceIn_(ss, name, course, program, area, docentesStr) {
  const sheet = getOrCreateIndiceIn_(ss);
  const data = sheet.getDataRange().getValues();
  const dStr = docentesStr || '';
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === name.trim()) {
      sheet.getRange(i + 1, 1, 1, 5).setValues([[name, course, program, area, dStr]]);
      return;
    }
  }
  sheet.appendRow([name, course, program, area, dStr]);
}
