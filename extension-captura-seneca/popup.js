const $ = (id) => document.getElementById(id);

chrome.storage.local.get('nombres', (d) => {
  $('nombres').value = d.nombres || '';
});

$('nombres').addEventListener('change', () => {
  chrome.storage.local.set({ nombres: $('nombres').value });
});

$('capturar').addEventListener('click', async () => {
  const estado = $('estado');
  estado.textContent = 'Capturando…';
  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    const nombres = $('nombres').value.split('\n').map((s) => s.trim()).filter((s) => s.length > 1);
    chrome.storage.local.set({ nombres: $('nombres').value });

    const resultados = await chrome.scripting.executeScript({
      target: { tabId: tab.id, allFrames: true },
      func: capturarMarco,
      args: [nombres],
    });

    const { contador = 0 } = await chrome.storage.local.get('contador');
    const n = contador + 1;
    await chrome.storage.local.set({ contador: n });

    const salida = {
      captura: n,
      fecha: new Date().toISOString(),
      nota: $('nota').value.trim(),
      marcos: resultados.map((r) => r.result).filter(Boolean),
    };
    const base = 'seneca-pantalla-' + String(n).padStart(2, '0');
    const json = 'data:application/json;charset=utf-8,' + encodeURIComponent(JSON.stringify(salida, null, 2));
    await chrome.downloads.download({ url: json, filename: base + '.json' });

    if ($('imagen').checked) {
      const png = await chrome.tabs.captureVisibleTab(tab.windowId, { format: 'png' });
      await chrome.downloads.download({ url: png, filename: base + '.png' });
    }

    const campos = salida.marcos.reduce((t, m) => t + m.campos.length, 0);
    estado.textContent = `Hecho: ${base}.json\n${salida.marcos.length} marcos, ${campos} campos.`;
    $('nota').value = '';
  } catch (e) {
    estado.textContent = 'Error: ' + e.message;
  }
});

// Se ejecuta dentro de cada marco de la página. Debe ser autocontenida.
function capturarMarco(nombres) {
  const MAX_HTML = 400000;
  const escapar = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const patrones = nombres.map((n) => new RegExp(escapar(n), 'gi'));
  const ocultar = (s) => {
    if (!s) return s;
    let r = s;
    for (const p of patrones) r = r.replace(p, '[NOMBRE]');
    // Automático: «Apellidos, Nombre - NIE: 123» (alumnado) y «Apellidos, Nombre (Profesor…)» (usuario).
    r = r.replace(/[A-ZÁÉÍÓÚÑ][^<>",;:\n]{1,60}, [A-ZÁÉÍÓÚÑ][^<>",;:\n(]{1,40}?(?=\s*(?:-\s*NIE|\())/g, '[NOMBRE]');
    r = r.replace(/NIE:\s*\w+/g, 'NIE: [oculto]');
    return r;
  };
  const texto = (el) => ocultar((el?.innerText || el?.textContent || '').replace(/\s+/g, ' ').trim()).slice(0, 200);

  // Quita los valores de la URL (pueden llevar identificadores del alumno), deja los nombres de parámetro.
  const limpiarUrl = (u) => {
    try {
      const url = new URL(u, location.href);
      for (const k of [...url.searchParams.keys()]) url.searchParams.set(k, '');
      return url.origin + url.pathname + (url.search ? url.search : '');
    } catch {
      return '';
    }
  };

  const etiquetaDe = (el) => {
    if (el.id) {
      const l = document.querySelector(`label[for="${CSS.escape(el.id)}"]`);
      if (l) return texto(l);
    }
    const lp = el.closest('label');
    if (lp) return texto(lp);
    // Séneca suele maquetar con tablas: texto de la celda anterior.
    const td = el.closest('td');
    if (td) {
      let prev = td.previousElementSibling;
      while (prev && !texto(prev)) prev = prev.previousElementSibling;
      if (prev) return texto(prev);
    }
    return el.getAttribute('title') || el.getAttribute('placeholder') || '';
  };

  const visible = (el) => !!(el.offsetWidth || el.offsetHeight || el.getClientRects().length);

  const campos = [...document.querySelectorAll('input, select, textarea')].map((el) => {
    const c = {
      etiqueta: etiquetaDe(el),
      tag: el.tagName.toLowerCase(),
      type: el.type || '',
      name: el.name || '',
      id: el.id || '',
      form: el.form ? el.form.name || el.form.id || '' : '',
      visible: visible(el),
      disabled: el.disabled,
      readOnly: !!el.readOnly,
      maxlength: el.maxLength > 0 ? el.maxLength : undefined,
      clase: el.className || undefined,
      eventos: {},
    };
    for (const ev of ['onchange', 'onclick', 'onblur', 'onkeyup', 'onfocus']) {
      const v = el.getAttribute(ev);
      if (v) c.eventos[ev] = v.slice(0, 300);
    }
    if (el.tagName === 'SELECT') {
      c.opciones = [...el.options].slice(0, 300).map((o) => ({ value: o.value, texto: ocultar(o.text.trim()), selected: o.selected }));
    } else if (['checkbox', 'radio'].includes(el.type)) {
      c.value = el.value;
      c.checked = el.checked;
    } else if (el.type === 'hidden') {
      c.value = '[oculto]';
    } else {
      c.longitudValor = (el.value || '').length; // nunca se guarda el contenido
    }
    return c;
  });

  const botones = [...document.querySelectorAll('button, input[type=button], input[type=submit], input[type=image], a, [onclick]')]
    .filter((el) => !['INPUT', 'SELECT', 'TEXTAREA'].includes(el.tagName) || /button|submit|image/.test(el.type))
    .map((el) => {
      const href = el.getAttribute('href') || '';
      return {
        texto: texto(el) || el.value || el.getAttribute('alt') || el.getAttribute('title') || '',
        tag: el.tagName.toLowerCase(),
        id: el.id || '',
        name: el.getAttribute('name') || '',
        visible: visible(el),
        onclick: (el.getAttribute('onclick') || '').slice(0, 300),
        href: href.startsWith('javascript:') ? href.slice(0, 300) : href ? limpiarUrl(href) : '',
        imagen: el.tagName === 'IMG' || el.type === 'image' ? limpiarUrl(el.src) : undefined,
      };
    })
    .filter((b) => b.texto || b.onclick || b.href);

  const formularios = [...document.forms].map((f) => ({
    name: f.name || '',
    id: f.id || '',
    action: limpiarUrl(f.getAttribute('action') || ''),
    method: f.method,
    onsubmit: (f.getAttribute('onsubmit') || '').slice(0, 300),
  }));

  // HTML anonimizado: sin valores de campos y con los nombres indicados sustituidos.
  const copia = document.documentElement.cloneNode(true);
  copia.querySelectorAll('input').forEach((el) => {
    if (!['checkbox', 'radio', 'button', 'submit', 'image'].includes(el.type)) el.setAttribute('value', '');
  });
  copia.querySelectorAll('textarea').forEach((el) => (el.textContent = ''));
  copia.querySelectorAll('style, link[rel=stylesheet]').forEach((el) => el.remove());
  let html = ocultar(copia.outerHTML).replace(/(\?|&amp;|&)([\w.-]+)=[^&"'\s]*/g, '$1$2=');
  const truncado = html.length > MAX_HTML;
  if (truncado) html = html.slice(0, MAX_HTML);

  return {
    url: limpiarUrl(location.href),
    nombreMarco: window.name || '',
    esPrincipal: window === window.top,
    titulo: ocultar(document.title),
    encabezados: [...document.querySelectorAll('h1, h2, h3, h4, legend, th, caption')].map(texto).filter(Boolean).slice(0, 100),
    formularios,
    campos,
    botones,
    html,
    htmlTruncado: truncado,
  };
}
