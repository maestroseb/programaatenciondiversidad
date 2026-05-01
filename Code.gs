/**
 * Programa de Atención a la Diversidad — Backend
 * CEIP Carlos III · La Carlota
 *
 * Hoja vinculada: 11bkLpUZKKkSbEPZkmCPqI23LBWreishKmIYL1yRJS74
 */

const SS_ID = '11bkLpUZKKkSbEPZkmCPqI23LBWreishKmIYL1yRJS74';
const INDICE_TAB = 'Índice';
const CONFIG_TAB = 'Config';

/* ───────── Web App entry point ───────── */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Programas de Atención a la Diversidad')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ───────── Helpers ───────── */

function getSS_() {
  return SpreadsheetApp.openById(SS_ID);
}

function getOrCreateIndice_() {
  const ss = getSS_();
  let sheet = ss.getSheetByName(INDICE_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(INDICE_TAB, 0);
    sheet.appendRow(['ALUMNO/A', 'CURSO', 'PROGRAMA', 'ÁREA/ÁMBITO']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/* ───────── CONFIG: datos del centro ───────── */

function getOrCreateConfig_() {
  const ss = getSS_();
  let sheet = ss.getSheetByName(CONFIG_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG_TAB, 0);
    sheet.appendRow(['CLAVE', 'VALOR']);
    sheet.appendRow(['centro', 'CEIP Carlos III']);
    sheet.appendRow(['localidad', 'La Carlota']);
    sheet.appendRow(['cursoEscolar', '']);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getConfig() {
  const sheet = getOrCreateConfig_();
  const data = sheet.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0]).trim();
    const val = String(data[i][1]).trim();
    if (key) config[key] = val;
  }
  // Auto-calculate cursoEscolar if empty
  if (!config.cursoEscolar) {
    const now = new Date();
    const y = now.getFullYear();
    const m = now.getMonth();
    config.cursoEscolar = m >= 8 ? (y + '/' + (y + 1)) : ((y - 1) + '/' + y);
  }
  return config;
}

function saveConfig(payload) {
  const data = JSON.parse(payload);
  const sheet = getOrCreateConfig_();
  const rows = sheet.getDataRange().getValues();

  for (const key in data) {
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === key) {
        sheet.getRange(i + 1, 2).setValue(data[key]);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([key, data[key]]);
    }
  }
  return { success: true };
}

/* ───────── READ: lista de alumnos ───────── */

function getStudentList() {
  const sheet = getOrCreateIndice_();
  const data = sheet.getDataRange().getValues();
  const students = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] && String(row[0]).trim()) {
      students.push({
        name: String(row[0]).trim(),
        course: String(row[1]).trim(),
        program: String(row[2]).trim(),
        area: String(row[3]).trim()
      });
    }
  }
  return students;
}

/* ───────── READ: datos de un alumno ───────── */

function getStudentData(studentName) {
  const ss = getSS_();
  const sheet = ss.getSheetByName(studentName);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;

  // Row 1: metadata
  const meta = data[0];
  const result = {
    studentName: String(meta[1] || '').trim(),
    course: String(meta[3] || '').trim(),
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
    const col0 = String(row[0] || '').trim();
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
  const data = JSON.parse(payload);
  const ss = getSS_();
  const tabName = data.studentName.trim();

  // Create or clear student sheet
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

  // Row 1: metadata (8 cols, but pad to NUM_COLS for setValues width)
  rows.push(['ALUMNO/A', data.studentName, 'CURSO', data.course, 'PROGRAMA', 'PE', 'ÁMBITOS']);
  // The 8th metadata cell (areaNames) will be written separately — keep the data range at 7 cols.

  // Row 2: headers
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

  // Follow-up fields
  rows.push(pad(['']));
  if (data.valoracionInicial) rows.push(pad(['', 'VALORACIÓN INICIAL', data.valoracionInicial]));
  if (data.seguimiento1T) rows.push(pad(['', 'SEGUIMIENTO 1T', data.seguimiento1T]));
  if (data.seguimiento2T) rows.push(pad(['', 'SEGUIMIENTO 2T', data.seguimiento2T]));
  if (data.seguimiento3T) rows.push(pad(['', 'SEGUIMIENTO 3T', data.seguimiento3T]));

  // Informe a las familias
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

  // Single batched write
  if (rows.length > 0) {
    sheet.getRange(1, 1, rows.length, NUM_COLS).setValues(rows);
    // Write the 8th metadata cell (ÁMBITOS value) separately
    sheet.getRange(1, 8).setValue(areaNames);
  }

  // Format the sheet (also batched)
  formatStudentSheet_(sheet, rows);

  // Update Índice
  updateIndice_(data.studentName, data.course, 'PE', areaNames);

  return { success: true, message: 'Datos guardados correctamente' };
}

/* ───────── DELETE: eliminar alumno ───────── */

function deleteStudent(studentName) {
  const ss = getSS_();
  const sheet = ss.getSheetByName(studentName);
  if (sheet) {
    ss.deleteSheet(sheet);
  }

  // Remove from Índice
  const indice = getOrCreateIndice_();
  const data = indice.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() === studentName.trim()) {
      indice.deleteRow(i + 1);
      break;
    }
  }

  return { success: true };
}

/* ───────── Helpers: format & index ───────── */

function formatStudentSheet_(sheet, rowsData) {
  const NUM_COLS = 7;

  // Column widths
  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 450);
  sheet.setColumnWidth(4, 50);
  sheet.setColumnWidth(5, 50);
  sheet.setColumnWidth(6, 50);
  sheet.setColumnWidth(7, 300);

  // Freeze header rows
  sheet.setFrozenRows(2);

  const totalRows = rowsData ? rowsData.length : sheet.getLastRow();
  if (totalRows < 1) return;

  // Build full formatting matrices for the entire data range (rows 1..totalRows, cols 1..7)
  const backgrounds = [];
  const fontColors = [];
  const fontWeights = [];
  const wraps = [];

  // Helpers
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
      // Metadata row: bold labels at cols 1,3,5,7 (handled below as bold whole row, then unbold values)
      bg = null; fc = null; fw = 'bold';
    } else if (i === 1) {
      // Headers
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

  // Metadata row: also format col 8 + unbold value cells
  const metaWeights = [['bold', 'normal', 'bold', 'normal', 'bold', 'normal', 'bold', 'normal']];
  sheet.getRange(1, 1, 1, 8).setFontWeights(metaWeights);
}

function updateIndice_(name, course, program, area) {
  const sheet = getOrCreateIndice_();
  const data = sheet.getDataRange().getValues();

  // Check if student exists
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === name.trim()) {
      // Update existing row
      sheet.getRange(i + 1, 1, 1, 4).setValues([[name, course, program, area]]);
      return;
    }
  }

  // Add new row
  sheet.appendRow([name, course, program, area]);
}