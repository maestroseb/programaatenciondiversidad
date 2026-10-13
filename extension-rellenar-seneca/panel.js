// Panel lateral: lee la pantalla de Séneca, la compara con los datos copiados en la app
// y rellena los campos de lo que falta. Nunca pulsa Aplicar ni Aceptar.

const $ = (id) => document.getElementById(id);
let datos = null; // formato 'pae-seneca' copiado desde la app
let lectura = null; // { tabId, frameId, info }
let ultimaFirma = '';
let ocupado = false;
let ambitoElegido = '';
let objetivoElegido = '';

const MAX_TEXTO = 1000;

// ─── Utilidades ───
const norm = (s) => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
  .replace(/\s+/g, ' ').replace(/[\s.;:,]+$/, '').trim();
const sinNumero = (s) => String(s || '').replace(/^\s*\d+\s*\.\s*/, '').trim();

function el(tag, cls, texto) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (texto != null) e.textContent = texto;
  return e;
}

// Igualdad tolerando textos recortados por Séneca con «...»
function iguales(app, sen) {
  const a = norm(app);
  const b = norm(sen);
  if (a === b) return true;
  const corte = String(sen).trim().match(/^(.*?)(\.\.\.|…)$/);
  return !!corte && norm(corte[1]).length >= 20 && a.startsWith(norm(corte[1]));
}

function comparar(app, seneca) {
  const enSeneca = (t) => seneca.some((s) => iguales(t, s));
  const enApp = (s) => app.some((t) => iguales(t, s));
  return {
    items: app.map((t) => ({ texto: t, estado: enSeneca(t) ? 'ok' : 'falta' }))
      .concat(seneca.filter((s) => !enApp(s)).map((s) => ({ texto: s, estado: 'sobra' }))),
    faltan: app.filter((t) => !enSeneca(t)),
    sobran: seneca.filter((s) => !enApp(s)),
  };
}

const ETIQUETA = { ok: '✓ Igual', falta: '+ Falta en Séneca', sobra: '− Sobra en Séneca' };

function lista(diff) {
  const ul = el('ul');
  if (!diff.items.length) ul.append(el('li', '', '(vacío)'));
  diff.items.forEach((it) => {
    const li = el('li', it.estado);
    li.append(el('span', 'etq', ETIQUETA[it.estado]), document.createTextNode(it.texto));
    if (it.estado === 'falta' && it.texto.length > MAX_TEXTO) li.append(el('div', 'nota', `⚠ ${it.texto.length} caracteres: Séneca admite ${MAX_TEXTO}.`));
    ul.append(li);
  });
  return ul;
}

function notaSobran(diff, que) {
  return diff.sobran.length
    ? el('p', 'nota', `Si sobra ${que}, bórralo tú en Séneca: pulsa sobre él → Borrar. Si solo ha cambiado el texto, usa Detalle para editarlo.`)
    : null;
}

let ultimoEstado = null;
function estado(msg, error) {
  ultimoEstado = { msg, error };
  let e = $('estado');
  if (!e) { e = el('div'); e.id = 'estado'; $('pantalla').append(e); }
  e.textContent = msg;
  e.style.color = error ? 'var(--sobra)' : 'var(--ok)';
}

// ─── Datos de la app ───
function cargarDatos(texto) {
  let d;
  try { d = JSON.parse(texto); } catch { return 'Lo pegado no es el formato de la app. Pulsa «Copiar para Séneca» en la ficha.'; }
  if (!d || d.formato !== 'pae-seneca' || !Array.isArray(d.ambitos)) return 'Lo pegado no viene del botón «Copiar para Séneca» de la app.';
  datos = d;
  chrome.storage.local.set({ datos: d });
  ambitoElegido = '';
  objetivoElegido = '';
  pintarResumen();
  ultimaFirma = '';
  leer();
  return null;
}

function pintarResumen() {
  const r = $('resumenDatos');
  if (!datos) return;
  const nObj = datos.ambitos.reduce((n, a) => n + a.objetivos.length, 0);
  r.className = '';
  r.textContent = `${datos.alumno || 'Alumno/a'}${datos.curso ? ' · ' + datos.curso : ''} — ${datos.ambitos.length} ámbitos, ${nObj} objetivos`;
}

$('pegar').addEventListener('click', async () => {
  try {
    const err = cargarDatos(await navigator.clipboard.readText());
    if (err) { $('resumenDatos').textContent = err; $('resumenDatos').className = 'muted'; }
  } catch {
    $('manual').hidden = false;
    $('manual').focus();
  }
});
$('verPegado').addEventListener('click', () => { $('manual').hidden = !$('manual').hidden; $('manual').focus(); });
$('manual').addEventListener('change', () => {
  const err = cargarDatos($('manual').value);
  if (err) { $('resumenDatos').textContent = err; $('resumenDatos').className = 'muted'; } else { $('manual').value = ''; $('manual').hidden = true; }
});

// ─── Lectura de Séneca (se ejecuta dentro de cada marco; autocontenida) ───
function leerPantalla() {
  const g = (k) => document.getElementById(k) || document.querySelector(`[name="${k}"]`);
  const filas = (tabla) => {
    const out = [];
    document.querySelectorAll(`td[onclick*="showOpciones('${tabla}'"]`).forEach((td) => {
      const m = td.getAttribute('onclick').match(/'\s*,\s*(\d+)\s*,\s*(\d+)/);
      if (!m || m[2] !== '1') return;
      const conTitulo = td.querySelector('[title]'); // Séneca recorta con «...» y deja el texto completo en title
      out.push(((conTitulo && conTitulo.getAttribute('title')) || td.innerText).replace(/\s+/g, ' ').trim());
    });
    return out;
  };
  if (g('N_OBJ') && g('D_TITULO')) {
    return {
      pantalla: 'area',
      titulo: g('D_TITULO').value,
      organizacion: g('D_ORGANIZACION') ? g('D_ORGANIZACION').value : '',
      recursos: g('D_RECURSOS') ? g('D_RECURSOS').value : '',
      objetivos: filas('TABLA_OBJETIVOS'),
      texto: document.body.innerText.slice(0, 4000),
    };
  }
  if (g('X_OBJPROESP') && g('N_IND')) {
    const s = g('X_OBJPROESP');
    const o = s.options[s.selectedIndex];
    return {
      pantalla: 'mapa',
      objetivo: o && o.value !== '-1' ? (o.title || o.text).trim() : '',
      indicadores: filas('TABLA_INDICADORES'),
      contenidos: filas('TABLA_CONTENIDOS'),
      actividades: filas('TABLA_ACTIVIDADES'),
      competencias: filas('TABLA_COMPETENCIAS'),
      compDisponibles: g('X_COMBAS') ? [...g('X_COMBAS').options].map((x) => x.text.trim()).filter(Boolean) : [],
    };
  }
  if (g('cAreasDisp') && g('cAreasAsig')) {
    return {
      pantalla: 'datos',
      asignadas: [...g('cAreasAsig').options].map((x) => x.text.trim()),
      disponibles: [...g('cAreasDisp').options].map((x) => x.text.trim()),
    };
  }
  return null;
}

// ─── Escritura en Séneca (se ejecuta en el marco detectado; autocontenida) ───
async function accionEnPagina(op) {
  const g = (k) => document.getElementById(k) || document.querySelector(`[name="${k}"]`);
  const espera = (ms) => new Promise((r) => setTimeout(r, ms));
  const fijar = (e, v) => {
    e.value = v;
    e.dispatchEvent(new Event('input', { bubbles: true }));
    e.dispatchEvent(new Event('change', { bubbles: true }));
  };
  const resaltar = (e) => {
    if (!e) return;
    e.scrollIntoView({ block: 'center', behavior: 'smooth' });
    e.style.outline = '4px solid #f0b35a';
    e.style.outlineOffset = '2px';
    setTimeout(() => { e.style.outline = ''; }, 10000);
  };
  const pestanna = async (txt) => {
    if (!txt) return;
    const a = [...document.querySelectorAll('a.nav-link')].find((x) => x.textContent.trim() === txt);
    if (a && !a.classList.contains('active')) { a.click(); await espera(300); }
  };

  if (op.tipo === 'lista') {
    await pestanna(op.pestanna);
    const num = g(op.num);
    if (!num) return { error: 'No encuentro el campo ' + op.num };
    fijar(num, String(op.textos.length));
    const btn = document.querySelector(`button[onclick*="'${op.td}'"]`);
    if (!btn) return { error: 'No encuentro el botón Mostrar' };
    btn.click();
    let cuadros = [];
    for (let i = 0; i < 50; i++) {
      cuadros = [...document.querySelectorAll('#_DHTML_TABLA_DINAMICA_' + op.td + ' textarea')];
      if (cuadros.length >= op.textos.length) break;
      await espera(100);
    }
    if (cuadros.length < op.textos.length) return { error: 'No han aparecido los cuadros de texto al pulsar Mostrar' };
    op.textos.forEach((t, i) => fijar(cuadros[i], t));
    resaltar(g(op.aplicar));
    return { ok: true };
  }
  if (op.tipo === 'competencia') {
    await pestanna(op.pestanna);
    const s = g('X_COMBAS');
    const o = s && [...s.options].find((x) => x.text.trim() === op.texto);
    if (!o) return { error: 'Esa competencia no está en la lista de Séneca' };
    fijar(s, o.value);
    resaltar(g('BOTON_APLICAR_COM'));
    return { ok: true };
  }
  if (op.tipo === 'campos') {
    let primero = null;
    for (const [id, v] of Object.entries(op.valores)) {
      const e = g(id);
      if (!e) continue;
      fijar(e, v);
      e.style.background = '#fff4d6';
      primero = primero || e;
    }
    resaltar(primero);
    return primero ? { ok: true } : { error: 'No encuentro los campos' };
  }
  if (op.tipo === 'areas') {
    const d = g('cAreasDisp');
    let n = 0;
    [...d.options].forEach((o) => { o.selected = op.textos.includes(o.text.trim()); if (o.selected) n++; });
    if (!n) return { error: 'Esas áreas no están en la lista de disponibles' };
    g('rightcAreas').click();
    return { ok: true };
  }
  return { error: 'Acción desconocida' };
}

async function pestañaActiva() {
  const [tab] = await chrome.tabs.query({ active: true, lastFocusedWindow: true });
  return tab;
}

async function leer() {
  if (ocupado) return;
  const tab = await pestañaActiva();
  let info = null;
  let frameId = 0;
  if (tab && /juntadeandalucia\.es/.test(tab.url || '')) {
    try {
      const res = await chrome.scripting.executeScript({ target: { tabId: tab.id, allFrames: true }, func: leerPantalla });
      const r = res.find((x) => x.result);
      if (r) { info = r.result; frameId = r.frameId; }
    } catch { /* pestaña sin permiso o recargándose */ }
  }
  const firma = JSON.stringify([info, !!datos, ambitoElegido, objetivoElegido]);
  if (firma === ultimaFirma) return;
  ultimaFirma = firma;
  lectura = tab && info ? { tabId: tab.id, frameId, info } : null;
  pintar();
}

async function ejecutar(op, siguiente) {
  if (!lectura) return;
  ocupado = true;
  try {
    const [r] = await chrome.scripting.executeScript({
      target: { tabId: lectura.tabId, frameIds: [lectura.frameId] },
      func: accionEnPagina,
      args: [op],
    });
    const res = r && r.result;
    if (res && res.ok) estado(siguiente);
    else estado((res && res.error) || 'No se pudo completar', true);
  } catch (e) {
    estado('Error: ' + e.message, true);
  } finally {
    ocupado = false;
  }
}

function boton(texto, onClick, sec) {
  const b = el('button', sec ? 'sec' : '', texto);
  b.addEventListener('click', onClick);
  return b;
}

// ─── Pintado según la pantalla ───
function pintar() {
  const p = $('pantalla');
  p.replaceChildren();
  p.className = '';
  if (!lectura) {
    $('tituloPantalla').textContent = 'Pantalla de Séneca';
    p.className = 'muted';
    p.textContent = 'Abre en Séneca una de estas pantallas del programa específico: Datos del PE, Propuesta de área (Detalle) o Mapa de relaciones.';
    return;
  }
  if (!datos) {
    p.className = 'muted';
    p.textContent = 'Pega primero los datos de la app.';
    return;
  }
  const info = lectura.info;
  if (ultimoEstado && ultimoEstado.pantalla && ultimoEstado.pantalla !== info.pantalla) ultimoEstado = null;
  if (info.pantalla === 'datos') pintarDatosPE(p, info);
  else if (info.pantalla === 'area') pintarArea(p, info);
  else if (info.pantalla === 'mapa') pintarMapa(p, info);
  if (ultimoEstado) { ultimoEstado.pantalla = info.pantalla; estado(ultimoEstado.msg, ultimoEstado.error); }
}

function pintarDatosPE(p, info) {
  $('tituloPantalla').textContent = 'Datos del PE: áreas';
  const diff = comparar(datos.ambitos.map((a) => a.ambito), info.asignadas);
  p.append(lista(diff));
  const posibles = diff.faltan.filter((t) => info.disponibles.some((d) => norm(d) === norm(t)));
  const exactos = info.disponibles.filter((d) => posibles.some((t) => norm(t) === norm(d)));
  if (exactos.length) {
    p.append(boton(`Pasar a asignadas ${exactos.length} área(s)`, () => ejecutar({ tipo: 'areas', textos: exactos }, 'Hecho. Revisa y pulsa Aceptar en Séneca.')));
  }
  if (diff.faltan.length > exactos.length) p.append(el('p', 'nota', 'Alguna área de la app no aparece con ese nombre en Séneca: pásala tú a mano.'));
  const n = notaSobran(diff, 'un área');
  if (n) p.append(n);
}

function elegirAmbito(info) {
  if (ambitoElegido && datos.ambitos.some((a) => a.ambito === ambitoElegido)) return ambitoElegido;
  const texto = norm(info.texto);
  const porNombre = datos.ambitos.find((a) => texto.includes(norm(a.ambito)));
  if (porNombre) return porNombre.ambito;
  // Si no, el que más objetivos comparte con la tabla de Séneca
  const sen = new Set(info.objetivos.map(norm));
  let mejor = datos.ambitos[0];
  let max = -1;
  datos.ambitos.forEach((a) => {
    const n = a.objetivos.filter((o) => sen.has(norm(o.objetivo))).length;
    if (n > max) { max = n; mejor = a; }
  });
  return mejor ? mejor.ambito : '';
}

function selector(opciones, valor, onChange) {
  const s = el('select');
  opciones.forEach(([v, t]) => { const o = el('option', '', t); o.value = v; o.selected = v === valor; s.append(o); });
  s.addEventListener('change', () => onChange(s.value));
  return s;
}

function pintarArea(p, info) {
  $('tituloPantalla').textContent = 'Propuesta de área';
  const nombre = elegirAmbito(info);
  const amb = datos.ambitos.find((a) => a.ambito === nombre);
  p.append(el('div', 'nota', 'Ámbito de la app:'), selector(datos.ambitos.map((a) => [a.ambito, a.ambito]), nombre, (v) => { ambitoElegido = v; ultimaFirma = ''; leer(); }));
  if (!amb) return;

  // Título, organización y recursos
  const org = amb.organizacion || {};
  const campos = [
    ['D_TITULO', 'Título', amb.programa, info.titulo],
    ['D_ORGANIZACION', 'Organización', org.organizacion, info.organizacion],
    ['D_RECURSOS', 'Recursos', org.recursos, info.recursos],
  ];
  const cambian = campos.filter(([, , app, sen]) => String(app || '').trim() && norm(app) !== norm(sen));
  p.append(el('h2', '', 'Título, organización y recursos'));
  const ul = el('ul');
  campos.forEach(([, etq, app, sen]) => {
    const vacio = !String(app || '').trim();
    const igual = vacio || norm(app) === norm(sen);
    const li = el('li', igual ? 'ok' : 'falta');
    li.append(el('span', 'etq', etq + (vacio ? ': vacío en la app' : igual ? ': igual' : ': distinto')));
    if (!igual) li.append(el('div', 'nota', 'App: ' + String(app).slice(0, 200)));
    ul.append(li);
  });
  p.append(ul);
  if (cambian.length) {
    const valores = {};
    cambian.forEach(([id, , app]) => { valores[id] = String(app).trim(); });
    p.append(boton(`Escribir lo de la app (${cambian.length})`, () => ejecutar({ tipo: 'campos', valores }, 'Escrito (fondo amarillo). Revisa y pulsa Aceptar en Séneca.')));
  }

  // Objetivos
  p.append(el('h2', '', 'Objetivos'));
  const diff = comparar(amb.objetivos.map((o) => o.objetivo), info.objetivos);
  p.append(lista(diff));
  if (diff.faltan.length) {
    p.append(boton(`Añadir ${diff.faltan.length} objetivo(s) que faltan`, () => ejecutar(
      { tipo: 'lista', num: 'N_OBJ', td: 'TD_OBJ', aplicar: 'BOTON_APLICAR_OBJ', textos: diff.faltan },
      'Escritos. Revisa y pulsa Aplicar en Séneca.')));
  }
  const n = notaSobran(diff, 'un objetivo');
  if (n) p.append(n);
  p.append(el('p', 'nota', 'Después, abre cada objetivo con «Mapa de relaciones» para sus indicadores, contenidos, actividades y competencias.'));
}

function todosLosObjetivos() {
  const out = [];
  datos.ambitos.forEach((a) => a.objetivos.forEach((o, i) => out.push({ clave: a.ambito + '|' + i, etiqueta: `${a.ambito} · ${i + 1}. ${o.objetivo.slice(0, 80)}`, obj: o })));
  return out;
}

function pintarMapa(p, info) {
  $('tituloPantalla').textContent = 'Mapa de relaciones';
  const todos = todosLosObjetivos();
  const delSeneca = norm(sinNumero(info.objetivo));
  let actual = todos.find((x) => x.clave === objetivoElegido) || todos.find((x) => norm(x.obj.objetivo) === delSeneca);
  if (!info.objetivo) { p.className = 'muted'; p.textContent = 'Elige un objetivo en Séneca.'; return; }
  if (!todos.find((x) => x.clave === objetivoElegido) && !actual) {
    p.append(el('p', 'nota', 'El objetivo de Séneca no coincide con ninguno de la app. Elige cuál es:'));
  }
  p.append(el('div', 'nota', 'Objetivo de la app:'), selector([['', '—']].concat(todos.map((x) => [x.clave, x.etiqueta])), actual ? actual.clave : '', (v) => { objetivoElegido = v; ultimaFirma = ''; leer(); }));
  if (!actual) return;
  const o = actual.obj;

  const secciones = [
    ['Indicadores', o.indicadores, info.indicadores, { num: 'N_IND', td: 'TD_IND', aplicar: 'BOTON_APLICAR_IND', pestanna: 'Indicadores evaluación' }, 'un indicador'],
    ['Contenidos', o.contenidos, info.contenidos, { num: 'N_CONT', td: 'TD_CONT', aplicar: 'BOTON_APLICAR_CONT', pestanna: 'Contenidos' }, 'un contenido'],
    ['Actividades y tareas', o.actividades, info.actividades, { num: 'N_ACT', td: 'TD_ACT', aplicar: 'BOTON_APLICAR_ACT', pestanna: 'Actividades y Tareas' }, 'una actividad'],
  ];
  secciones.forEach(([titulo, app, sen, cfg, que]) => {
    p.append(el('h2', '', titulo));
    const diff = comparar(app || [], sen);
    p.append(lista(diff));
    if (diff.faltan.length) {
      p.append(boton(`Añadir ${diff.faltan.length} que faltan`, () => ejecutar(
        Object.assign({ tipo: 'lista', textos: diff.faltan }, cfg),
        `Escritos en «${cfg.pestanna}». Revisa y pulsa Aplicar en Séneca.`)));
    }
    const n = notaSobran(diff, que);
    if (n) p.append(n);
  });

  p.append(el('h2', '', 'Competencias'));
  const diffC = comparar(o.competencias || [], info.competencias);
  p.append(lista(diffC));
  const siguiente = diffC.faltan.find((c) => info.compDisponibles.some((d) => norm(d) === norm(c)));
  if (siguiente) {
    const exacto = info.compDisponibles.find((d) => norm(d) === norm(siguiente));
    p.append(boton(`Seleccionar «${exacto}»`, () => ejecutar(
      { tipo: 'competencia', texto: exacto, pestanna: 'Contribución competencias' },
      'Seleccionada. Pulsa Aplicar en Séneca y vuelve aquí para la siguiente.')));
  }
  if (diffC.faltan.length && !siguiente) p.append(el('p', 'nota', 'Las competencias que faltan no aparecen con ese nombre en la lista de Séneca.'));
  const n = notaSobran(diffC, 'una competencia');
  if (n) p.append(n);
}

// ─── Arranque ───
$('leer').addEventListener('click', () => { ultimaFirma = ''; leer(); });
chrome.tabs.onActivated.addListener(() => leer());
chrome.tabs.onUpdated.addListener((id, cambio) => { if (cambio.status === 'complete') leer(); });
setInterval(() => { if (document.visibilityState === 'visible') leer(); }, 3000);

chrome.storage.local.get('datos', (d) => {
  if (d.datos) { datos = d.datos; pintarResumen(); }
  leer();
});
