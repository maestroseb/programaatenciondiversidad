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
    programType: String(meta[5] || '').trim(),
    area: String(meta[7] || '').trim(),
    objectives: []
  };

  // Find the header row (contains 'TIPO') then read data after it
  let headerIndex = -1;
  for (let h = 1; h < data.length; h++) {
    if (String(data[h][1] || '').trim().toUpperCase() === 'TIPO') {
      headerIndex = h;
      break;
    }
  }
  if (headerIndex < 0) return result; // no data rows

  let currentObj = null;

  for (let i = headerIndex + 1; i < data.length; i++) {
    const row = data[i];
    const objNum = String(row[0] || '').trim();
    const tipo = String(row[1] || '').trim().toUpperCase();
    const texto = String(row[2] || '').trim();
    const eval1T = String(row[3] || '').trim();
    const eval2T = String(row[4] || '').trim();
    const eval3T = String(row[5] || '').trim();

    if (!tipo && !texto) continue;

    if (tipo === 'OBJETIVO') {
      currentObj = {
        title: texto,
        indicators: [],
        contents: [],
        activities: []
      };
      result.objectives.push(currentObj);
    } else if (currentObj) {
      const item = { text: texto, eval1T, eval2T, eval3T };
      if (tipo === 'INDICADOR') {
        currentObj.indicators.push(item);
      } else if (tipo === 'CONTENIDO') {
        currentObj.contents.push(item);
      } else if (tipo === 'ACTIVIDAD') {
        currentObj.activities.push(item);
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

  // Row 1: metadata
  sheet.appendRow([
    'ALUMNO/A', data.studentName,
    'CURSO', data.course,
    'PROGRAMA', data.programType,
    'ÁREA', data.area
  ]);

  // Row 2: headers
  sheet.appendRow(['Nº OBJ', 'TIPO', 'TEXTO', '1T', '2T', '3T']);

  // Row 4+: objectives data
  const objectives = data.objectives || [];
  for (let i = 0; i < objectives.length; i++) {
    const obj = objectives[i];
    const objLabel = 'Obj. ' + (i + 1);

    // Objective description
    sheet.appendRow([objLabel, 'OBJETIVO', obj.title || '', '', '', '']);

    // Indicators (with evaluation data)
    (obj.indicators || []).forEach(function(ind) {
      if (ind.text && ind.text.trim()) {
        sheet.appendRow([objLabel, 'INDICADOR', ind.text.trim(),
          ind.eval1T || '', ind.eval2T || '', ind.eval3T || '']);
      }
    });

    // Contents
    (obj.contents || []).forEach(function(cnt) {
      if (cnt.text && cnt.text.trim()) {
        sheet.appendRow([objLabel, 'CONTENIDO', cnt.text.trim(), '', '', '']);
      }
    });

    // Activities
    (obj.activities || []).forEach(function(act) {
      if (act.text && act.text.trim()) {
        sheet.appendRow([objLabel, 'ACTIVIDAD', act.text.trim(), '', '', '']);
      }
    });

    // Empty row between objectives
    if (i < objectives.length - 1) {
      sheet.appendRow(['']);
    }
  }

  // Format the sheet
  formatStudentSheet_(sheet);

  // Update Índice
  updateIndice_(data.studentName, data.course, data.programType, data.area);

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

function formatStudentSheet_(sheet) {
  // Header row 1 (metadata)
  const metaRange = sheet.getRange(1, 1, 1, 8);
  metaRange.setFontWeight('bold');
  sheet.getRange(1, 1).setFontWeight('bold');
  sheet.getRange(1, 3).setFontWeight('bold');
  sheet.getRange(1, 5).setFontWeight('bold');
  sheet.getRange(1, 7).setFontWeight('bold');
  sheet.getRange(1, 2).setFontWeight('normal');
  sheet.getRange(1, 4).setFontWeight('normal');
  sheet.getRange(1, 6).setFontWeight('normal');
  sheet.getRange(1, 8).setFontWeight('normal');

  // Headers row 2
  const headerRange = sheet.getRange(2, 1, 1, 6);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#2d6a4f');
  headerRange.setFontColor('#ffffff');

  // Column widths
  sheet.setColumnWidth(1, 80);   // Nº OBJ
  sheet.setColumnWidth(2, 100);  // TIPO
  sheet.setColumnWidth(3, 450);  // TEXTO
  sheet.setColumnWidth(4, 50);   // 1T
  sheet.setColumnWidth(5, 50);   // 2T
  sheet.setColumnWidth(6, 50);   // 3T

  // Freeze header rows
  sheet.setFrozenRows(2);

  // Color-code rows by type
  const lastRow = sheet.getLastRow();
  if (lastRow > 2) {
    const dataRange = sheet.getRange(3, 1, lastRow - 2, 6);
    const values = dataRange.getValues();

    for (let i = 0; i < values.length; i++) {
      const tipo = String(values[i][1]).trim().toUpperCase();
      const row = i + 3;
      const range = sheet.getRange(row, 1, 1, 6);

      if (tipo === 'OBJETIVO') {
        range.setBackground('#d1fae5').setFontWeight('bold');
      } else if (tipo === 'INDICADOR') {
        range.setBackground('#fef3c7');
      } else if (tipo === 'CONTENIDO') {
        range.setBackground('#ede9fe');
      } else if (tipo === 'ACTIVIDAD') {
        range.setBackground('#dbeafe');
      }
    }
  }
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