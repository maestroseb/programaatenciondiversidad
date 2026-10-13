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
      objetivoValor: o ? o.value : '',
      opcionesObjetivo: [...s.options].filter((x) => x.value !== '-1').map((x) => ({ valor: x.value, texto: (x.title || x.text).trim() })),
      indicadores: filas('TABLA_INDICADORES'),
      contenidos: filas('TABLA_CONTENIDOS'),
      actividades: filas('TABLA_ACTIVIDADES'),
      competencias: filas('TABLA_COMPETENCIAS'),
      compDisponibles: g('X_COMBAS') ? [...g('X_COMBAS').options].map((x) => x.text.trim()).filter(Boolean) : [],
    };
  }
  const apartados = [...document.querySelectorAll('td[onclick*="showOpciones("]')]
    .map((td) => ({ texto: td.innerText.replace(/\s+/g, ' ').trim(), m: td.getAttribute('onclick').match(/showOpciones\((\d+)\)/) }))
    .filter((x) => x.m && /^Propuesta de .reas/i.test(x.texto));
  if (apartados.length) {
    return {
      pantalla: 'apartados',
      areas: apartados.map((x) => ({ fila: x.m[1], area: ((x.texto.match(/\(([^()]*)\)\s*$/) || [])[1] || '').trim() })),
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
    if (op.pulsar) { await espera(200); g(op.aplicar).click(); } else resaltar(g(op.aplicar));
    return { ok: true };
  }
  if (op.tipo === 'competencia') {
    await pestanna(op.pestanna);
    const s = g('X_COMBAS');
    const o = s && [...s.options].find((x) => x.text.trim() === op.texto);
    if (!o) return { error: 'Esa competencia no está en la lista de Séneca' };
    fijar(s, o.value);
    if (op.pulsar) { await espera(200); g('BOTON_APLICAR_COM').click(); } else resaltar(g('BOTON_APLICAR_COM'));
    return { ok: true };
  }
  if (op.tipo === 'menu') {
    // Pulsa una fila (abre el menú emergente de Séneca) y luego la opción indicada
    const td = document.querySelector(`td[onclick*="${op.fila}"]`);
    if (!td) return { error: 'No encuentro la fila en la tabla' };
    td.click();
    for (let i = 0; i < 30; i++) {
      await espera(100);
      const item = [...document.querySelectorAll('[id^="menuItem"]')].find((x) => x.offsetParent && x.textContent.replace(/\s+/g, ' ').includes(op.opcion));
      if (item) { item.click(); return { ok: true }; }
    }
    return { error: 'No aparece la opción «' + op.opcion + '» en el menú' };
  }
  if (op.tipo === 'boton') {
    // Botones de la barra de Séneca: aceptar() / volver()
    const a = [...document.querySelectorAll('a[href]')].find((x) => x.getAttribute('href').includes('cuerpo.' + op.funcion + '()'));
    if (!a) return { error: 'No encuentro el botón ' + op.funcion };
    a.click();
    return { ok: true };
  }
  if (op.tipo === 'objetivo') {
    const s = g('X_OBJPROESP');
    if (!s) return { error: 'No encuentro el selector de objetivo' };
    fijar(s, op.valor);
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

async function leerAhora(tabId) {
  const tab = tabId ? await chrome.tabs.get(tabId).catch(() => null) : await pestañaActiva();
  let info = null;
  let frameId = 0;
  if (tab && /juntadeandalucia\.es/.test(tab.url || '')) {
    try {
      const res = await chrome.scripting.executeScript({ target: { tabId: tab.id, allFrames: true }, func: leerPantalla });
      const r = res.find((x) => x.result);
      if (r) { info = r.result; frameId = r.frameId; }
    } catch { /* pestaña sin permiso o recargándose */ }
  }
  return { tab, info, frameId };
}

async function leer() {
  if (ocupado) return;
  const { tab, info, frameId } = await leerAhora();
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

// ─── Modo automático: el panel dirige cada paso y espera a ver el resultado en la tabla ───
let cancelar = false;
const dormir = (ms) => new Promise((r) => setTimeout(r, ms));

async function accionAuto(op) {
  const { tab, info, frameId } = await leerAhora(lectura.tabId);
  if (!tab || !info) throw new Error('No encuentro la pantalla de Séneca');
  const [r] = await chrome.scripting.executeScript({ target: { tabId: tab.id, frameIds: [frameId] }, func: accionEnPagina, args: [op] });
  const res = r && r.result;
  if (!res || !res.ok) throw new Error((res && res.error) || 'No se pudo completar');
}

// Relee la pantalla hasta que se cumpla la condición (tolera recargas de Séneca)
async function esperar(cond, que, ms = 25000) {
  const fin = Date.now() + ms;
  while (Date.now() < fin) {
    if (cancelar) throw new Error('Detenido');
    await dormir(700);
    const { info } = await leerAhora(lectura.tabId);
    if (info && cond(info)) return info;
  }
  throw new Error('Séneca no ha mostrado el resultado esperado: ' + que);
}

async function automatico(pasos) {
  ocupado = true;
  cancelar = false;
  pintarProgreso('Empezando…');
  try {
    await pasos((msg) => pintarProgreso(msg));
  } catch (e) {
    ocupado = false;
    ultimaFirma = '';
    await leer();
    estado((e.message === 'Detenido' ? 'Detenido. ' : 'Parado: ') + e.message.replace('Detenido', '') + ' Lo ya aplicado está guardado; puedes volver a pulsar el botón para seguir.', e.message !== 'Detenido');
    return;
  }
  ocupado = false;
  ultimaFirma = '';
  await leer();
}

function pintarProgreso(msg) {
  const p = $('pantalla');
  p.replaceChildren(el('p', '', '⏳ ' + msg), boton('Detener', () => { cancelar = true; }, true));
}

const SECCIONES = [
  ['indicadores', 'Indicadores', { num: 'N_IND', td: 'TD_IND', aplicar: 'BOTON_APLICAR_IND', pestanna: 'Indicadores evaluación' }],
  ['contenidos', 'Contenidos', { num: 'N_CONT', td: 'TD_CONT', aplicar: 'BOTON_APLICAR_CONT', pestanna: 'Contenidos' }],
  ['actividades', 'Actividades', { num: 'N_ACT', td: 'TD_ACT', aplicar: 'BOTON_APLICAR_ACT', pestanna: 'Actividades y Tareas' }],
];

function objetivoDeApp(textoSeneca) {
  return todosLosObjetivos().find((x) => iguales(x.obj.objetivo, sinNumero(textoSeneca)));
}

// Recorre todos los objetivos del desplegable del mapa de relaciones y aplica lo que falta
async function hacerMapa(avance) {
  let info = (await leerAhora(lectura.tabId)).info;
  const lista = info.opcionesObjetivo.map((op) => ({ op, app: objetivoDeApp(op.texto) })).filter((x) => x.app);
  if (!lista.length) throw new Error('Ningún objetivo de Séneca coincide con los de la app.');
  for (let k = 0; k < lista.length; k++) {
    const { op, app } = lista[k];
    const etiqueta = `Objetivo ${k + 1}/${lista.length}`;
    if (info.objetivoValor !== op.valor) {
      avance(`${etiqueta}: abriendo…`);
      await accionAuto({ tipo: 'objetivo', valor: op.valor });
      info = await esperar((i) => i.pantalla === 'mapa' && i.objetivoValor === op.valor, 'cambio de objetivo');
    }
    for (const [clave, nombre, cfg] of SECCIONES) {
      const faltan = comparar(app.obj[clave] || [], info[clave]).faltan;
      if (!faltan.length) continue;
      avance(`${etiqueta}: ${nombre.toLowerCase()} (${faltan.length})…`);
      await accionAuto(Object.assign({ tipo: 'lista', textos: faltan, pulsar: true }, cfg));
      info = await esperar((i) => i.pantalla === 'mapa' && i.objetivoValor === op.valor && !comparar(faltan, i[clave]).faltan.length, nombre.toLowerCase() + ' añadidos');
    }
    for (const c of comparar(app.obj.competencias || [], info.competencias).faltan) {
      const exacto = info.compDisponibles.find((d) => norm(d) === norm(c));
      if (!exacto) continue;
      avance(`${etiqueta}: competencia «${exacto}»…`);
      await accionAuto({ tipo: 'competencia', texto: exacto, pestanna: 'Contribución competencias', pulsar: true });
      info = await esperar((i) => i.pantalla === 'mapa' && i.objetivoValor === op.valor && i.competencias.some((x) => norm(x) === norm(exacto)), 'competencia añadida');
    }
  }
  const sinApp = info.opcionesObjetivo.length - lista.length;
  return `${lista.length} objetivo(s) completados.` + (sinApp ? ` ${sinApp} de Séneca no están en la app y no se han tocado.` : '');
}

function autoMapa() {
  automatico(async (avance) => {
    ultimoEstado = { msg: 'Listo: ' + await hacerMapa(avance) + ' Revisa y lo que sobre bórralo tú.' };
  });
}

// Recorrido completo: Apartados → cada área → objetivos → mapa de relaciones → volver → aceptar → siguiente
let pulsarAceptar = true;

function camposArea(amb, info) {
  const org = amb.organizacion || {};
  const valores = {};
  [['D_TITULO', amb.programa, info.titulo], ['D_ORGANIZACION', org.organizacion, info.organizacion], ['D_RECURSOS', org.recursos, info.recursos]]
    .forEach(([id, app, sen]) => { if (String(app || '').trim() && norm(app) !== norm(sen)) valores[id] = String(app).trim(); });
  return valores;
}

function autoPrograma() {
  automatico(async (avance) => {
    const hechos = new Set();
    const resumen = [];
    for (let vuelta = 0; vuelta < 40; vuelta++) {
      let info = (await leerAhora(lectura.tabId)).info;
      if (!info) throw new Error('No reconozco la pantalla de Séneca.');

      if (info.pantalla === 'apartados') {
        const sig = datos.ambitos.find((a) => !hechos.has(a.ambito) && info.areas.some((x) => norm(x.area) === norm(a.ambito)));
        if (!sig) break;
        const fila = info.areas.find((x) => norm(x.area) === norm(sig.ambito)).fila;
        avance(`${sig.ambito}: abriendo…`);
        ambitoElegido = sig.ambito;
        await accionAuto({ tipo: 'menu', fila: `showOpciones(${fila})`, opcion: 'Detalle' });
        await esperar((i) => i.pantalla === 'area', 'abrir la propuesta de área');
        continue;
      }
      if (info.pantalla !== 'area') throw new Error('Empieza desde «Apartados del documento» o desde una Propuesta de área.');

      const amb = datos.ambitos.find((a) => a.ambito === elegirAmbito(info));
      if (!amb || hechos.has(amb.ambito)) break;
      hechos.add(amb.ambito);

      const faltan = comparar(amb.objetivos.map((o) => o.objetivo), info.objetivos).faltan;
      if (faltan.length) {
        avance(`${amb.ambito}: añadiendo ${faltan.length} objetivo(s)…`);
        await accionAuto({ tipo: 'lista', num: 'N_OBJ', td: 'TD_OBJ', aplicar: 'BOTON_APLICAR_OBJ', textos: faltan, pulsar: true });
        info = await esperar((i) => i.pantalla === 'area' && !comparar(faltan, i.objetivos).faltan.length, 'objetivos añadidos');
      }
      if (info.objetivos.length) {
        avance(`${amb.ambito}: abriendo el mapa de relaciones…`);
        await accionAuto({ tipo: 'menu', fila: "showOpciones('TABLA_OBJETIVOS', 1, 1", opcion: 'Mapa de relaciones' });
        await esperar((i) => i.pantalla === 'mapa', 'abrir el mapa de relaciones');
        resumen.push(amb.ambito + ': ' + await hacerMapa((m) => avance(amb.ambito + ' · ' + m)));
        avance(`${amb.ambito}: volviendo…`);
        await accionAuto({ tipo: 'boton', funcion: 'volver' });
        info = await esperar((i) => i.pantalla === 'area', 'volver a la propuesta de área');
      }

      const valores = camposArea(amb, info);
      if (Object.keys(valores).length) {
        avance(`${amb.ambito}: título, organización y recursos…`);
        await accionAuto({ tipo: 'campos', valores });
      }
      if (!pulsarAceptar) {
        ultimoEstado = { msg: resumen.join(' ') + ` Revisa «${amb.ambito}» y pulsa Aceptar en Séneca. Para la siguiente área, vuelve a pulsar el botón.` };
        return;
      }
      avance(`${amb.ambito}: aceptando…`);
      await accionAuto({ tipo: 'boton', funcion: 'aceptar' });
      info = await esperar((i) => i.pantalla === 'apartados' || (i.pantalla === 'area' && !Object.keys(camposArea(amb, i)).length), 'guardar la propuesta de área');
      if (info.pantalla === 'area') {
        await accionAuto({ tipo: 'boton', funcion: 'volver' });
        await esperar((i) => i.pantalla === 'apartados', 'volver a Apartados');
      }
    }
    const sinArea = datos.ambitos.filter((a) => !hechos.has(a.ambito)).map((a) => a.ambito);
    ultimoEstado = { msg: 'Programa completado. ' + resumen.join(' ') + (sinArea.length ? ` No están en Séneca: ${sinArea.join(', ')} (asígnalas en Datos del PE y vuelve a pulsar).` : '') + ' Lo que sobre, bórralo tú.' };
  });
}

function controlesPrograma(p) {
  p.append(boton('▶ Rellenar todo el programa', () => autoPrograma()));
  const lbl = el('label', 'nota');
  const chk = el('input');
  chk.type = 'checkbox';
  chk.checked = pulsarAceptar;
  chk.addEventListener('change', () => { pulsarAceptar = chk.checked; });
  lbl.append(chk, document.createTextNode(' Pulsar Aceptar al terminar cada área y seguir con la siguiente'));
  p.append(lbl);
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
  else if (info.pantalla === 'apartados') pintarApartados(p, info);
  if (ultimoEstado) { ultimoEstado.pantalla = info.pantalla; estado(ultimoEstado.msg, ultimoEstado.error); }
}

function pintarApartados(p, info) {
  $('tituloPantalla').textContent = 'Apartados del documento';
  const diff = comparar(datos.ambitos.map((a) => a.ambito), info.areas.map((x) => x.area));
  p.append(lista(diff));
  controlesPrograma(p);
  p.append(el('p', 'nota', 'Entra en cada área, añade los objetivos, completa el mapa de relaciones de cada uno y vuelve, sin que toques nada. No borra nada. Se para con «Detener».'));
  if (diff.faltan.length) p.append(el('p', 'nota', 'Las áreas que faltan hay que asignarlas antes en Detalle del PE.'));
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

  controlesPrograma(p);
  p.append(el('p', 'nota', 'Hace esta área entera (objetivos, mapa de relaciones de cada uno, título, organización y recursos) y sigue con las demás. No borra nada. O paso a paso:'));

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
  if (!info.objetivo) { p.append(boton('▶ Rellenar todos los objetivos', () => autoMapa())); return; }
  if (!todos.find((x) => x.clave === objetivoElegido) && !actual) {
    p.append(el('p', 'nota', 'El objetivo de Séneca no coincide con ninguno de la app. Elige cuál es:'));
  }
  p.append(boton('▶ Rellenar todos los objetivos', () => autoMapa()), el('p', 'nota', 'Recorre cada objetivo del desplegable y aplica los indicadores, contenidos, actividades y competencias que faltan. No borra nada. O este objetivo paso a paso:'));
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
