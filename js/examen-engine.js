/* Motor del panel de exámenes y retos (configurable con window.EXAMEN_CFG): cronómetro, subtareas con casilla, entrega, tablero (1000 pts) y archivo de entrega con el tiempo. */
(function () {
  const CFG = window.EXAMEN_CFG, KEY = CFG.key, $ = id => document.getElementById(id);
  let st = JSON.parse(localStorage.getItem(KEY) || 'null');
  if (st) { st.hechas = st.hechas || {}; st.marcas = st.marcas || {}; }
  const N = CFG.tareas.length, ENTREGA = N + 1; // 0 = Sumario, 1..N = tareas/áreas, N+1 = Entrega
  let tab = 0;
  const guardar = () => localStorage.setItem(KEY, JSON.stringify(st));
  const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const hms = ms => { const s = Math.max(0, Math.floor(ms / 1000)); return [Math.floor(s / 3600), Math.floor(s / 60) % 60, s % 60].map(n => String(n).padStart(2, '0')).join(':'); };
  const trans = () => (st.fin || Date.now()) - st.inicio;
  const TOTAL = CFG.total, MIN = CFG.aprobado, nPasos = CFG.tareas.reduce((a, t) => a + t.pasos.length, 0);
  const ptsTarea = t => t.pasos.reduce((a, p) => a + CFG.puntos[p.item].total, 0);
  const hechasTarea = t => t.pasos.filter(p => st.hechas[p.item]).length;
  const zonaArch = (tipo, num, tit, ext) => `<div class="drop${ARCH[tipo] ? ' listo' : ''}" data-tipo="${tipo}"><iconify-icon icon="fluent:document-arrow-up-24-regular"></iconify-icon><div class="drop-t"><b>${num}. ${tit} (${ext})</b><span id="n${cap(tipo)}">Ningún archivo seleccionado</span><small>Arrástralo aquí o usa el botón</small></div><div class="drop-b"><button type="button" class="btn-sel" data-sel="${tipo}">Seleccionar archivo…</button>${window.showOpenFilePicker ? `<button type="button" class="lnk" data-alt="${tipo}">Selector alternativo</button>` : ''}</div><input type="file" id="f${cap(tipo)}" accept="${ext}" class="sr-only" tabindex="-1" aria-label="${tit}"></div>`;
  const nivelTag = n => `<span class="nivel n${n}">${CFG.nombreNivel[n]}</span>`;

  // ---------- archivos adjuntos: selector nativo + selector alternativo + arrastrar y soltar ----------
  const ARCH = { docx: null, pdf: null };
  const TIPOS = { docx: { ext: /\.docx$/i, nom: 'Word (.docx)', mime: { 'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'] } }, pdf: { ext: /\.pdf$/i, nom: 'PDF (.pdf)', mime: { 'application/pdf': ['.pdf'] } } };
  const cap = t => t.charAt(0).toUpperCase() + t.slice(1);
  const tamano = b => b > 1048576 ? (b / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(b / 1024)) + ' KB';
  function avisarEntrega(txt, error) { const m = $('evalMsg'); if (m) { m.textContent = txt || ''; m.className = 'res' + (error ? ' err-ent' : ''); } }
  function pintarArch(tipo) { const n = $('n' + cap(tipo)), z = document.querySelector('.drop[data-tipo="' + tipo + '"]'); if (!n || !z) return; const f = ARCH[tipo]; n.textContent = f ? f.name + ' · ' + tamano(f.size) : 'Ningún archivo seleccionado'; z.classList.toggle('listo', !!f); }
  function asignar(tipo, f) {
    if (!f) return false;
    if (!TIPOS[tipo].ext.test(f.name)) { avisarEntrega('«' + f.name + '» no es un archivo ' + TIPOS[tipo].nom + '. Selecciona el archivo correcto.', true); return false; }
    ARCH[tipo] = f; pintarArch(tipo); avisarEntrega(''); return true;
  }
  async function selectorAlterno(tipo) { // File System Access API (Edge/Chrome): no depende del <input type="file">
    if (!window.showOpenFilePicker) return false;
    try { const [h] = await window.showOpenFilePicker({ multiple: false, types: [{ description: TIPOS[tipo].nom, accept: TIPOS[tipo].mime }] }); return asignar(tipo, await h.getFile()); }
    catch (e) { return !!(e && e.name === 'AbortError'); }
  }
  function elegirArchivo(tipo) {
    const inp = $('f' + cap(tipo)); let abrio = false; const alBlur = () => { abrio = true; }; window.addEventListener('blur', alBlur, { once: true });
    try { inp.click(); } catch (e) { }
    setTimeout(() => { // si la ventana de Windows no tomó el foco, se usa el selector alterno
      window.removeEventListener('blur', alBlur); if (abrio || !document.hasFocus()) return;
      if (window.showOpenFilePicker) selectorAlterno(tipo).then(ok => { if (!ok) avisarEntrega('No se pudo abrir la ventana de selección. Arrastra tu archivo hasta el recuadro punteado.', true); });
      else avisarEntrega('Si no se abrió la ventana de selección, arrastra tu archivo hasta el recuadro punteado.', true);
    }, 900);
  }
  function conectarEntrega() {
    ['docx', 'pdf'].forEach(tipo => {
      const z = document.querySelector('.drop[data-tipo="' + tipo + '"]'); if (!z) return; const inp = $('f' + cap(tipo));
      z.querySelector('[data-sel]').onclick = () => elegirArchivo(tipo);
      const alt = z.querySelector('[data-alt]'); if (alt) alt.onclick = () => selectorAlterno(tipo);
      inp.onchange = () => { if (inp.files[0]) asignar(tipo, inp.files[0]); inp.value = ''; };
      z.ondragover = e => { e.preventDefault(); z.classList.add('sobre'); }; z.ondragleave = () => z.classList.remove('sobre');
      z.ondrop = e => { e.preventDefault(); e.stopPropagation(); z.classList.remove('sobre'); asignar(tipo, e.dataTransfer.files[0]); };
      pintarArch(tipo);
    });
  }
  // Soltar un archivo en cualquier parte del panel lo asigna según su extensión (evita que el navegador lo abra y se pierda el panel)
  window.addEventListener('dragover', e => e.preventDefault());
  window.addEventListener('drop', e => {
    e.preventDefault(); if (!st || $('app').hidden) return; const fs = [...(e.dataTransfer ? e.dataTransfer.files : [])]; if (!fs.length) return;
    if (tab !== ENTREGA) { tab = ENTREGA; pintarTabs(); pintar(); }
    fs.forEach(f => { const t = /\.docx$/i.test(f.name) ? 'docx' : (/\.pdf$/i.test(f.name) && CFG.pdf ? 'pdf' : null); if (t) asignar(t, f); else avisarEntrega('«' + f.name + '» no es un archivo válido para este examen.', true); });
  });

  // ---------- descarga de archivos sin abandonar el panel ----------
  function descargar(nombre) {
    const r = CFG.recursosB64 && CFG.recursosB64[nombre];
    if (!r) { window.open('assets/docs/final/' + nombre, '_blank'); return; }
    const bin = atob(r.b64), u8 = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
    const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([u8], { type: r.mime })); a.download = nombre;
    document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }
  const AVISO = `<div class="aviso"><iconify-icon icon="fluent:document-arrow-right-24-regular"></iconify-icon><div><b>Antes de empezar, abre el documento base en Word.</b> Si usaste el lanzador, ya se abrió <code>${CFG.docTrabajo}</code> (carpeta ${CFG.carpeta}). Si no, descarga <code>${CFG.docBase}</code> en el Sumario y ábrelo. Realiza <b>todas las tareas en ese documento</b> y guarda con frecuencia (Ctrl+G).</div></div>`;

  // ---------- inicio ----------
  function arrancar() {
    if (!st) { $('start').hidden = false; $('app').hidden = true; return; }
    $('start').hidden = true; $('app').hidden = false; $('quien').textContent = st.nombre;
    $('btnResultado').hidden = !st.ultimo; pintarTabs(); pintar(); avance(); tick(); etiquetaDock();
  }
  $('btnComenzar').onclick = () => {
    const n = $('nombre').value.trim().replace(/\s+/g, ' ');
    if (n.split(' ').filter(w => w.length > 1).length < 2) { $('nombreErr').textContent = 'Escribe tu nombre completo (al menos nombre y apellido).'; return; }
    st = { nombre: n, inicio: Date.now(), fin: null, marcas: {}, hechas: {}, intentos: 0, ultimo: null }; guardar(); tab = 0; arrancar();
  };
  $('nombre').addEventListener('keydown', e => { if (e.key === 'Enter') $('btnComenzar').click(); });
  // ---------- ventana: acoplar al 25 % (siempre disponible) ----------
  const MODO_KEY = 'tbox_dock_modo';
  let modo = new URLSearchParams(location.search).get('modo') || localStorage.getItem(MODO_KEY) || 'abajo';
  if (!['abajo', 'derecha'].includes(modo)) modo = 'abajo';
  const area = () => ({ x: screen.availLeft || 0, y: screen.availTop || 0, w: screen.availWidth, h: screen.availHeight });
  const rectPanel = m => { const a = area(), k = m === 'derecha' ? [Math.round(a.w * .75), 0] : [0, Math.round(a.h * .75)];
    return m === 'derecha' ? { x: a.x + k[0], y: a.y, w: a.w - k[0], h: a.h } : { x: a.x, y: a.y + k[1], w: a.w, h: a.h - k[1] }; };
  const colocar = r => { try { window.moveTo(r.x, r.y); window.resizeTo(r.w, r.h); } catch (e) { /* algunos navegadores lo restringen */ } };
  const cerca = r => Math.abs(outerWidth - r.w) <= 30 && Math.abs(outerHeight - r.h) <= 30;
  let toastT;
  function aviso(txt) { const t = $('toast'); t.textContent = txt; t.hidden = false; clearTimeout(toastT); toastT = setTimeout(() => { t.hidden = true; }, 5000); }
  function etiquetaDock() {
    $('dockLbl').textContent = (modo === 'abajo' ? 'Acoplar 25 % ↓' : 'Acoplar 25 % →');
    $('dockIcon').setAttribute('icon', modo === 'abajo' ? 'fluent:panel-bottom-24-regular' : 'fluent:panel-right-24-regular');
    const a = area(), grande = outerWidth > a.w * .6 && outerHeight > a.h * .6; // ventana grande/maximizada: perdió la proporción
    $('btnDock').classList.toggle('llamar', grande);
    if (grande) $('dockLbl').textContent = 'Restaurar 25 %';
  }
  function acoplar(m) {
    if (m) { modo = m; localStorage.setItem(MODO_KEY, modo); }
    const r = rectPanel(modo); colocar(r); etiquetaDock();
    setTimeout(() => {
      etiquetaDock(); if (cerca(r)) { aviso('Panel acoplado al 25 % de la pantalla.'); return; }
      // Alternativa: el navegador no dejó mover esta ventana → se abre una ventana emergente con el tamaño exacto
      const extra = { w: Math.max(0, outerWidth - innerWidth), h: Math.max(0, outerHeight - innerHeight) };
      const url = location.href.split('#')[0].replace(/[?&]modo=\w+/, '') + (location.search ? '&' : '?') + 'modo=' + modo;
      const w = window.open(url, 'panel_examen', `popup=yes,left=${r.x},top=${r.y},width=${Math.max(300, r.w - extra.w)},height=${Math.max(200, r.h - extra.h)}`);
      if (!w) { aviso('Permite las ventanas emergentes de este sitio y vuelve a pulsar el botón, o ejecuta Reacomodar_Ventanas.bat.'); return; }
      setTimeout(() => { try { w.moveTo(r.x, r.y); w.resizeTo(r.w, r.h); } catch (e) { } }, 500);
      aviso('Se abrió el panel acoplado. Tu avance y cronómetro continúan allí; puedes cerrar esta ventana.');
    }, 450);
  }
  function ampliar() {
    const a = area(); colocar({ x: a.x, y: a.y, w: a.w, h: a.h });
    setTimeout(() => { etiquetaDock(); if (!cerca({ w: a.w, h: a.h })) aviso('Usa el botón Maximizar (□) de la barra de título de esta ventana.'); }, 450);
  }
  window.addEventListener('resize', () => { if (!$('app').hidden) etiquetaDock(); });
  const cerrarMenu = () => { $('dockMenu').hidden = true; $('btnDockMenu').setAttribute('aria-expanded', 'false'); };
  $('btnDock').onclick = () => acoplar();
  $('btnDockMenu').onclick = e => { e.stopPropagation(); const abrir = $('dockMenu').hidden; $('dockMenu').hidden = !abrir; $('btnDockMenu').setAttribute('aria-expanded', String(abrir)); };
  $('dockMenu').querySelectorAll('[data-acc]').forEach(b => b.onclick = () => { cerrarMenu(); b.dataset.acc === 'ampliar' ? ampliar() : acoplar(b.dataset.acc); });
  document.addEventListener('click', e => { if (!$('dock').contains(e.target)) cerrarMenu(); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape') cerrarMenu(); });
  $('btnPanel').onclick = () => acoplar();
  $('btnReiniciar').onclick = () => { if (confirm('¿Reiniciar el examen? Se borrarán tu avance, tus marcas y el cronómetro.')) { localStorage.removeItem(KEY); st = null; location.reload(); } };
  $('btnResultado').onclick = () => mostrarTablero(st.ultimo);

  // ---------- cronómetro ----------
  function tick() { $('reloj').textContent = hms(trans()); }
  setInterval(() => { if (st && !st.fin) tick(); }, 500);

  // ---------- avance, pestañas y contenido ----------
  function avance() {
    const n = Object.keys(st.hechas).filter(k => st.hechas[k]).length, tot = CFG.tareas.reduce((a, t) => a + t.pasos.length, 0);
    $('avTxt').textContent = `${n}/${tot} subtareas`; $('avBar').style.width = (n / tot * 100) + '%';
  }
  function pintarTabs() {
    const nombres = [{ n: 'Sumario' }, ...CFG.tareas.map(t => ({ n: CFG.etiqueta + ' ' + t.id, t })), { n: 'Entrega' }];
    $('tabs').innerHTML = nombres.map((o, i) => {
      let cls = i === tab ? 'on ' : '', extra = '';
      if (o.t) { const h = hechasTarea(o.t), tot = o.t.pasos.length; cls += h === tot ? 'ok ' : h > 0 ? 'part ' : ''; cls += st.marcas[i] ? 'rev' : ''; extra = ` <small>${h}/${tot}</small>`; }
      return `<button data-i="${i}" class="${cls}">${o.n}${extra}</button>`;
    }).join('');
    $('tabs').querySelectorAll('button').forEach(b => b.onclick = () => { tab = +b.dataset.i; pintarTabs(); pintar(); });
    const t = CFG.tareas[tab - 1];
    $('btnOk').lastChild.textContent = t && hechasTarea(t) === t.pasos.length ? ' Desmarcar todas' : ' Marcar todas como hechas';
    $('btnRev').classList.toggle('activo', !!(t && st.marcas[tab]));
  }
  function pintar() {
    const el = $('instr'); el.scrollTop = 0; let h = '';
    if (tab === 0) {
      h = `<h2>Sumario del proyecto</h2>${AVISO}<p class="res">${CFG.descripcionDoc} Son ${CFG.tareas.length} ${CFG.etiquetaPlural} con ${nPasos} subtareas que suman <b>${TOTAL} puntos</b>; cada subtarea vale según su dificultad: ${nivelTag(1)} ${nivelTag(2)} ${nivelTag(3)}. Marca cada subtarea con su casilla cuando la termines y usa <b>Ayuda</b> para ver la ruta en la cinta de opciones.</p>
        <b>Archivos de trabajo</b><div class="recursos">${CFG.recursos.map(r => `<a href="${r.archivo}" data-r="${r.nombre}"><iconify-icon icon="${r.icono}"></iconify-icon><span><b>${r.nombre}</b><small>${r.desc}</small></span></a>`).join('')}</div>
        <ol class="resumen">${CFG.tareas.map(t => `<li><span class="num">${t.id}</span><span><b>${CFG.etiqueta} ${t.id}: ${t.titulo}.</b> ${t.resumen} <span class="pts">${ptsTarea(t)} pts</span></span></li>`).join('')}</ol>${CFG.extraSumario || ''}`;
      el.classList.remove('cols');
    } else if (tab === ENTREGA) {
      h = `<h2>Entrega y calificación</h2><p class="res">Sube tu documento de Word terminado${CFG.pdf ? ' y el PDF exportado' : ''}. La calificación (sobre ${TOTAL} puntos) se calcula en tu equipo; ningún archivo sale de él.</p>
        <div class="subida">${zonaArch('docx', 1, 'Documento Word', '.docx')}${CFG.pdf ? zonaArch('pdf', 2, 'Exportación a PDF', '.pdf') : ''}</div>
        <p class="res nota-ent">¿No se abre la ventana de selección? Arrastra tu archivo hasta el recuadro punteado o usa «Selector alternativo». Si la ventana aparece detrás de Word, búscala en la barra de tareas.</p>
        <button class="btn btn-tb" style="width:auto" id="btnEval"><iconify-icon icon="fluent:play-24-filled"></iconify-icon> Evaluar mi trabajo</button> <span id="evalMsg" class="res"></span>`;
      el.classList.remove('cols');
    } else {
      const t = CFG.tareas[tab - 1];
      h = `<h2>${CFG.etiqueta} ${t.id}: ${t.titulo} <span class="pts">${ptsTarea(t)} pts</span></h2>${tab === 1 ? AVISO : '<p class="res" style="margin-top:-6px"><iconify-icon icon="fluent:document-24-regular"></iconify-icon> Trabaja en el documento base abierto en Word.</p>'}
        <ol class="steps">${t.pasos.map((p, k) => { const P = CFG.puntos[p.item]; return `<li class="${st.hechas[p.item] ? 'hecho' : ''}"><span class="num">${t.id}.${k + 1}</span><label class="paso"><input type="checkbox" data-item="${p.item}" ${st.hechas[p.item] ? 'checked' : ''}><span><span class="txt">${p.t}</span> <span class="pts">${P.total} pts</span> ${nivelTag(P.nivel)}<span class="ayuda"><b>Ruta:</b> ${p.ayuda}</span></span></label></li>`; }).join('')}</ol>`;
      el.classList.add('cols');
    }
    el.innerHTML = h;
    el.querySelectorAll('a[data-r]').forEach(a => a.onclick = e => { e.preventDefault(); descargar(a.dataset.r); });
    el.querySelectorAll('input[data-item]').forEach(c => c.onchange = () => { st.hechas[c.dataset.item] = c.checked; c.closest('li').classList.toggle('hecho', c.checked); guardar(); pintarTabs(); avance(); });
    if (tab === ENTREGA) { $('btnEval').onclick = evaluar; conectarEntrega(); }
    $('btnPrev').disabled = tab === 0; $('btnNext').disabled = tab === ENTREGA;
    $('btnOk').disabled = $('btnRev').disabled = tab < 1 || tab > N;
  }
  $('btnOk').onclick = () => { const t = CFG.tareas[tab - 1]; if (!t) return; const todas = hechasTarea(t) !== t.pasos.length; t.pasos.forEach(p => st.hechas[p.item] = todas); guardar(); pintarTabs(); pintar(); avance(); };
  $('btnRev').onclick = () => { if (tab < 1 || tab > N) return; st.marcas[tab] = !st.marcas[tab]; guardar(); pintarTabs(); };
  $('btnPrev').onclick = () => { if (tab > 0) { tab--; pintarTabs(); pintar(); } };
  $('btnNext').onclick = () => { if (tab < ENTREGA) { tab++; pintarTabs(); pintar(); } };
  $('btnAyuda').onclick = () => $('instr').classList.toggle('ver-ayuda');

  // ---------- evaluación ----------
  async function evaluar() {
    const fd = ARCH.docx, fp = ARCH.pdf, msg = $('evalMsg');
    if (!fd) { avisarEntrega('Primero selecciona tu archivo .docx (botón «Seleccionar archivo…» o arrástralo al recuadro).', true); return; }
    try {
      msg.textContent = 'Analizando…';
      const zip = await JSZip.loadAsync(fd), files = {};
      for (const n of Object.keys(zip.files)) if (!zip.files[n].dir) files[n] = /\.(xml|rels|svg)$/i.test(n) ? await zip.file(n).async('text') : '';
      let pdf = null;
      if (fp) { const b = new Uint8Array(await fp.slice(0, 5).arrayBuffer()); pdf = { name: fp.name, size: fp.size, header: String.fromCharCode(...b).startsWith('%PDF') }; }
      const r = CFG.evaluar(files, { nombre: st.nombre, pdf });
      st.intentos++; const ms = trans();
      const tareas = CFG.tareas.map(t => { const its = r.items.filter(i => i.tarea === t.id); return { id: t.id, titulo: t.titulo, pts: its.reduce((a, i) => a + i.obtenidos, 0), max: its.reduce((a, i) => a + i.pts, 0) }; });
      st.ultimo = { r, tareas, ms, tiempo: hms(ms), intento: st.intentos, fecha: new Date().toLocaleString('es'), archivo: fd.name, pdf: pdf ? pdf.name : '—', hechas: { ...st.hechas }, codigo: codigo(st.nombre, r.total, hms(ms)) };
      guardar(); $('btnResultado').hidden = false; msg.textContent = ''; mostrarTablero(st.ultimo);
    } catch (e) { msg.textContent = 'No se pudo leer el archivo: ' + e.message; }
  }
  function codigo(n, t, tiempo) { let h = 2166136261; for (const c of `${n}|${t}|${tiempo}|TBOX`) { h ^= c.charCodeAt(0); h = Math.imul(h, 16777619) >>> 0; } return h.toString(36).toUpperCase().padStart(7, '0'); }

  // ---------- tablero ----------
  function mostrarTablero(u) {
    const r = u.r, ap = r.aprobado, d = $('dash'); d.hidden = false;
    const fallas = []; r.items.forEach(i => i.subs.forEach(s => { if (!s.ok) fallas.push({ i, s }); })); fallas.sort((a, b) => b.s.pts - a.s.pts);
    const perdidos = fallas.reduce((a, f) => a + f.s.pts, 0);
    d.innerHTML = `<div class="dash-in"><div class="dash-head"><h2><iconify-icon icon="fluent:data-bar-vertical-24-regular"></iconify-icon> Resultado · ${CFG.titulo}</h2>
      <div class="dash-actions" style="margin:0"><button class="btn btn-tb-o" id="dAmpliar" title="Agranda esta ventana para leer el reporte"><iconify-icon icon="fluent:full-screen-maximize-24-regular"></iconify-icon> Ampliar ventana</button><button class="btn btn-tb-o" id="dCerrar">${st.fin ? 'Cerrar' : 'Seguir trabajando'}</button></div></div>
      <div class="kpis"><div class="kpi ${ap ? 'ok' : 'bad'}"><b>${r.total} / ${r.max}</b><span>Puntaje (${Math.round(r.total / r.max * 100)} %) · ${ap ? 'APROBADO' : 'NO APROBADO'} (mín. ${r.minimo})</span></div>
      <div class="kpi"><b>${u.tiempo}</b><span>Tiempo trabajado</span></div>
      <div class="kpi"><b>${esc(st.nombre)}</b><span>Estudiante registrado</span></div>
      <div class="kpi"><b>${u.intento}</b><span>Intento de evaluación</span></div>
      <div class="kpi ${perdidos ? 'bad' : 'ok'}"><b>${perdidos}</b><span>Puntos por recuperar</span></div></div>
      <p class="res" style="font-size:.84rem;color:#616161">Nombre detectado en el documento: <b>${esc(r.nombreDetectado || 'No detectado')}</b> · Archivo: ${esc(u.archivo)} · PDF: ${esc(u.pdf)} · ${u.fecha} · Código de verificación: <b class="codigo">${u.codigo}</b></p>
      <h4>Puntaje por tarea</h4>
      <div class="barras">${u.tareas.map(t => `<div class="barra"><span>${CFG.etiqueta} ${t.id}: ${esc(t.titulo)}</span><i><s style="width:${t.pts / t.max * 100}%"></s></i><b>${t.pts}/${t.max}</b></div>`).join('')}</div>
      <h4>Qué reforzar (ordenado por puntos perdidos)</h4>
      ${fallas.length ? `<ul class="refuerzo">${fallas.map(f => `<li><b>−${f.s.pts} pts</b> ${nivelTag(f.s.nivel)} <span>${CFG.etiqueta} ${f.i.tarea} · criterio ${f.i.id}: ${esc(f.s.label)}</span><small>Pista: ${esc(f.s.pista)}</small></li>`).join('')}</ul>` : '<p class="si">¡Excelente! Todos los criterios se cumplieron.</p>'}
      <h4>Detalle por criterio</h4>
      <div class="tabla-wrap"><table><thead><tr><th style="width:36px">#</th><th>Criterio y subcriterios</th><th style="width:96px">Marcada por ti</th><th style="width:96px">Puntos</th></tr></thead><tbody>${r.items.map(i => {
        const m = !!(u.hechas && u.hechas[i.id]), incompleto = i.obtenidos < i.pts;
        return `<tr><td>${i.id}</td><td><b>${esc(i.titulo)}</b> ${nivelTag(i.nivel)}${i.subs.map(s => `<div>${s.ok ? '<span class="si">✓</span>' : '<span class="no">✗</span>'} ${esc(s.label)} <small class="gris">(${s.nivelNombre} · ${s.pts} pts)</small></div>`).join('')}</td>
        <td>${m ? '<span class="si">✓ Sí</span>' : '<span class="gris">—</span>'}${m && incompleto ? '<span class="alerta">Revisa: la marcaste, pero no está completa</span>' : ''}${!m && !incompleto ? '<span class="alerta ok2">Hecha sin marcar</span>' : ''}</td>
        <td><b>${i.obtenidos}</b> / ${i.pts}</td></tr>`; }).join('')}</tbody></table></div>
      <div class="dash-actions"><button class="btn btn-tb" id="dFin"><iconify-icon icon="fluent:arrow-download-24-regular"></iconify-icon> Finalizar y descargar entrega (.docx con tiempo)</button>
      <button class="btn btn-tb-o" id="dPrint"><iconify-icon icon="fluent:print-24-regular"></iconify-icon> Imprimir / guardar reporte PDF</button></div></div>`;
    $('dCerrar').onclick = () => { d.hidden = true; };
    $('dAmpliar').onclick = ampliar;
    $('dPrint').onclick = () => window.print();
    $('dFin').onclick = finalizar;
  }

  // ---------- archivo de entrega: incrusta tiempo y puntaje en las propiedades personalizadas del .docx ----------
  async function finalizar() {
    const f = ARCH.docx;
    if (!f) { alert('Para generar el archivo de entrega, vuelve a la pestaña Entrega y selecciona tu .docx.'); $('dash').hidden = true; tab = ENTREGA; pintarTabs(); pintar(); return; }
    if (!confirm('Al finalizar se detiene el cronómetro y se descarga tu entrega. ¿Continuar?')) return;
    const u = st.ultimo; if (!st.fin) { st.fin = Date.now(); guardar(); tick(); }
    const zip = await JSZip.loadAsync(f);
    const props = { TBox_Estudiante: st.nombre, TBox_TiempoTrabajado: u.tiempo, TBox_Puntaje: `${u.r.total}/${u.r.max}`, TBox_Intentos: String(u.intento), TBox_Fecha: u.fecha, TBox_Codigo: u.codigo };
    const body = Object.entries(props).map(([k, v], i) => `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${i + 2}" name="${k}"><vt:lpwstr>${esc(v)}</vt:lpwstr></property>`).join('');
    zip.file('docProps/custom.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">${body}</Properties>`);
    let ct = await zip.file('[Content_Types].xml').async('text');
    if (!/custom\.xml/.test(ct)) zip.file('[Content_Types].xml', ct.replace('</Types>', '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/></Types>'));
    let rl = await zip.file('_rels/.rels').async('text');
    if (!/custom-properties/.test(rl)) zip.file('_rels/.rels', rl.replace('</Relationships>', '<Relationship Id="rIdTBoxCustom" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>'));
    const blob = await zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob);
    a.download = 'Entrega_' + st.nombre.replace(/[^\wÁ-ú]+/g, '_') + '.docx'; a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 4000);
    mostrarTablero(u);
  }
  arrancar();
})();
