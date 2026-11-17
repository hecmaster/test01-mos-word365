/* Reproductor de cuestionarios teóricos: una pregunta a la vez, racha, XP, medallas, repaso de fallos y mejor resultado guardado. */
(function () {
  const KEY = 'tbox_quiz_v1', $ = (s, r) => (r || document).querySelector(s);
  const leer = () => { try { return JSON.parse(localStorage.getItem(KEY) || '{}'); } catch (e) { return {}; } };
  const guardar = o => { try { localStorage.setItem(KEY, JSON.stringify(o)); } catch (e) { } };
  const mezclar = a => { const b = a.slice(); for (let i = b.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [b[i], b[j]] = [b[j], b[i]]; } return b; };
  const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const OK = ['¡Correcto!', '¡Muy bien!', '¡Excelente!', '¡Así se hace!', '¡Perfecto!'], NO = ['Casi…', 'Buen intento.', 'Esta vez no, pero ya lo sabes.', 'Cada error enseña.'];
  const pick = a => a[Math.floor(Math.random() * a.length)], LETRAS = ['A', 'B', 'C', 'D', 'E'];
  const limpiarFb = t => String(t).replace(/^¡[^!]{1,30}!\s*/, '');
  const medalla = p => p >= 90 ? { n: 'Maestro MOS', e: '🏆', c: 'oro', m: '¡Dominas este tema!' } : p >= 70 ? { n: 'Plata', e: '🥈', c: 'plata', m: '¡Aprobado! Un repaso más y llegas al oro.' } : p >= 50 ? { n: 'Bronce', e: '🥉', c: 'bronce', m: 'Vas por buen camino. Repasa los fallos y vuelve a intentarlo.' } : { n: 'En entrenamiento', e: '💪', c: 'base', m: 'Todo experto empezó aquí. Repasa y reintenta.' };
  let S = null, root = null;

  function preparar(id, soloIdx) {
    const m = bancoTeorico[id], base = m.preguntas.map((p, i) => ({ p, i })).filter(x => !soloIdx || soloIdx.includes(x.i));
    const qs = mezclar(base).map(({ p, i }) => ({ i, tema: p.tema, pregunta: p.pregunta, fb: limpiarFb(p.retroalimentacion), opts: mezclar(p.opciones.map((t, k) => ({ t: t.replace(/^[A-E]\)\s*/, ''), ok: k === p.correcta }))) }));
    S = { id, titulo: m.titulo.replace(/\s*\(\d+ Preguntas\)/i, ''), qs, i: 0, aciertos: 0, xp: 0, racha: 0, mejorRacha: 0, res: [], t0: Date.now(), fase: 'pregunta', sel: null, repaso: !!soloIdx };
  }
  function abrir(id, soloIdx) {
    if (!bancoTeorico[id]) return; preparar(id, soloIdx);
    if (!root) { root = document.createElement('div'); root.id = 'quiz'; root.className = 'qz'; root.setAttribute('role', 'dialog'); root.setAttribute('aria-modal', 'true'); document.body.appendChild(root); document.addEventListener('keydown', teclas); }
    root.hidden = false; document.body.style.overflow = 'hidden'; render();
  }
  function cerrar(forzar) {
    if (!forzar && S && S.fase !== 'fin' && S.res.length > 0 && !confirm('¿Salir del cuestionario? Perderás el avance de este intento.')) return;
    root.hidden = true; document.body.style.overflow = ''; S = null; if (window.actualizarProgresoQuiz) window.actualizarProgresoQuiz();
  }
  function render() { S.fase === 'fin' ? renderFin() : renderPregunta(); }

  function renderPregunta() {
    const q = S.qs[S.i], n = S.qs.length;
    root.innerHTML = `<div class="qz-shell">
      <header class="qz-top"><button class="qz-x" id="qzX" aria-label="Cerrar cuestionario">&times;</button><div class="qz-title">${esc(S.titulo)}${S.repaso ? ' · <b>Repaso</b>' : ''}</div>
        <div class="qz-stats"><span class="qz-chip" title="Racha de aciertos seguidos">🔥 <b id="qzRacha">${S.racha}</b></span><span class="qz-chip" title="Puntos de experiencia">⭐ <b id="qzXp">${S.xp}</b> XP</span></div></header>
      <div class="qz-seg" role="progressbar" aria-valuemin="0" aria-valuemax="${n}" aria-valuenow="${S.i}">${S.qs.map((_, k) => `<i class="${k < S.i ? (S.res[k] && S.res[k].ok ? 'ok' : 'no') : k === S.i ? 'on' : ''}"></i>`).join('')}</div>
      <main class="qz-main"><div class="qz-count">Pregunta ${S.i + 1} de ${n}</div><span class="qz-tema">${esc(q.tema)}</span><h2 class="qz-q">${esc(q.pregunta)}</h2>
        <div class="qz-opts" role="radiogroup" aria-label="Opciones de respuesta">${q.opts.map((o, k) => `<button type="button" class="qz-opt" role="radio" aria-checked="false" data-k="${k}"><span class="qz-let">${LETRAS[k]}</span><span class="qz-txt">${esc(o.t)}</span><span class="qz-ico" aria-hidden="true"></span></button>`).join('')}</div>
        <div class="qz-fb" id="qzFb" aria-live="polite"></div></main>
      <footer class="qz-foot"><span class="qz-hint">Elige una respuesta · teclas A–D y Enter</span><button class="qz-btn" id="qzMain" disabled>Comprobar</button></footer></div>`;
    $('#qzX').onclick = () => cerrar(); $('#qzMain').onclick = accion;
    root.querySelectorAll('.qz-opt').forEach(b => b.onclick = () => elegir(+b.dataset.k));
    S.fase = 'pregunta'; S.sel = null;
  }
  function elegir(k) {
    if (S.fase !== 'pregunta') return; S.sel = k;
    root.querySelectorAll('.qz-opt').forEach((b, j) => { b.classList.toggle('sel', j === k); b.setAttribute('aria-checked', String(j === k)); });
    $('#qzMain').disabled = false;
  }
  function accion() { if (S.fase === 'pregunta') comprobar(); else if (S.fase === 'feedback') siguiente(); }
  function comprobar() {
    if (S.sel === null) return; const q = S.qs[S.i], ok = q.opts[S.sel].ok; S.fase = 'feedback';
    const ganado = ok ? 10 + Math.min(S.racha, 5) * 2 : 0;
    if (ok) { S.aciertos++; S.racha++; S.mejorRacha = Math.max(S.mejorRacha, S.racha); S.xp += ganado; } else S.racha = 0;
    S.res[S.i] = { ok, i: q.i, tema: q.tema, pregunta: q.pregunta, correcta: q.opts.find(o => o.ok).t, elegida: q.opts[S.sel].t, fb: q.fb };
    root.querySelectorAll('.qz-opt').forEach((b, j) => { b.disabled = true; b.classList.remove('sel'); if (q.opts[j].ok) b.classList.add('bien'); else if (j === S.sel) b.classList.add('mal'); else b.classList.add('off'); });
    const fb = $('#qzFb'); fb.className = 'qz-fb ' + (ok ? 'ok' : 'no') + ' show';
    const extra = ok ? (S.racha >= 3 ? `<span class="qz-bonus">🔥 Racha de ${S.racha}</span>` : '') + `<span class="qz-bonus">+${ganado} XP</span>` : '';
    fb.innerHTML = `<div class="qz-fb-h"><b>${ok ? pick(OK) : pick(NO)}</b>${extra}</div>` + (ok ? '' : `<p class="qz-ans">Respuesta correcta: <b>${esc(S.res[S.i].correcta)}</b></p>`) + `<p>${esc(q.fb)}</p>`;
    $('#qzRacha').textContent = S.racha; $('#qzXp').textContent = S.xp;
    const seg = root.querySelectorAll('.qz-seg i')[S.i]; seg.className = ok ? 'ok' : 'no';
    const m = $('#qzMain'); m.textContent = S.i === S.qs.length - 1 ? 'Ver mi resultado' : 'Siguiente'; m.focus(); $('.qz-hint').textContent = 'Enter para continuar';
    fb.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
  }
  function siguiente() { if (S.i < S.qs.length - 1) { S.i++; renderPregunta(); } else { S.fase = 'fin'; renderFin(); } }

  function renderFin() {
    const n = S.qs.length, pct = Math.round(S.aciertos / n * 100), med = medalla(pct), seg = Math.round((Date.now() - S.t0) / 1000), tiempo = Math.floor(seg / 60) + ':' + String(seg % 60).padStart(2, '0');
    const all = leer(); const prev = all[S.id] || { intentos: 0, best: 0 }; let nuevo = false;
    if (!S.repaso) { nuevo = pct > prev.best; all[S.id] = { best: Math.max(prev.best, pct), intentos: prev.intentos + 1, fecha: new Date().toLocaleDateString('es'), medalla: medalla(Math.max(prev.best, pct)).c }; guardar(all); }
    const fallos = S.res.filter(r => !r.ok), temas = {}; S.res.forEach(r => { temas[r.tema] = temas[r.tema] || { a: 0, t: 0 }; temas[r.tema].t++; if (r.ok) temas[r.tema].a++; });
    root.innerHTML = `<div class="qz-shell qz-fin"><header class="qz-top"><button class="qz-x" id="qzX" aria-label="Cerrar">&times;</button><div class="qz-title">${esc(S.titulo)} · Resultado</div></header>
      <main class="qz-main"><div class="qz-hero ${med.c}"><div class="qz-ring" style="--p:${pct}"><span>${pct}%</span></div><div><div class="qz-medal">${med.e} ${med.n}</div><p>${med.m}</p>${nuevo ? '<span class="qz-bonus">★ ¡Nuevo mejor resultado!</span>' : ''}</div></div>
      <div class="qz-kpis"><div><b>${S.aciertos}/${n}</b><span>Aciertos</span></div><div><b>${S.xp}</b><span>XP ganados</span></div><div><b>🔥 ${S.mejorRacha}</b><span>Mejor racha</span></div><div><b>${tiempo}</b><span>Tiempo</span></div></div>
      <h3 class="qz-h">Resultado por tema</h3><div class="qz-temas">${Object.entries(temas).map(([t, v]) => `<div class="qz-tm"><span>${esc(t)}</span><i><s class="${v.a === v.t ? 'full' : ''}" style="width:${v.a / v.t * 100}%"></s></i><b>${v.a}/${v.t}</b></div>`).join('')}</div>
      ${fallos.length ? `<h3 class="qz-h">Para repasar (${fallos.length})</h3><div class="qz-rev">${fallos.map(f => `<details><summary>${esc(f.pregunta)}</summary><p><span class="no">Tu respuesta:</span> ${esc(f.elegida)}</p><p><span class="si">Correcta:</span> <b>${esc(f.correcta)}</b></p><p class="qz-exp">${esc(f.fb)}</p></details>`).join('')}</div>` : '<p class="qz-perfect">¡Sin errores! Respuestas perfectas. 🎉</p>'}</main>
      <footer class="qz-foot qz-foot-fin">${fallos.length ? '<button class="qz-btn sec" id="qzRep">Repasar solo mis fallos</button>' : ''}<button class="qz-btn sec" id="qzRe">Intentar de nuevo</button><button class="qz-btn" id="qzCerrar">Terminar</button></footer></div>`;
    $('#qzX').onclick = () => cerrar(true); $('#qzCerrar').onclick = () => cerrar(true); $('#qzRe').onclick = () => abrir(S.id);
    if ($('#qzRep')) $('#qzRep').onclick = () => abrir(S.id, fallos.map(f => f.i));
    if (pct >= 70 && !matchMedia('(prefers-reduced-motion: reduce)').matches) confeti();
  }
  function confeti() { const c = document.createElement('div'); c.className = 'qz-confetti'; const col = ['#B5CC03', '#089BD9', '#FFB900', '#0B3552', '#F25F00']; for (let i = 0; i < 46; i++) { const s = document.createElement('i'); s.style.cssText = `left:${Math.random() * 100}%;background:${col[i % 5]};animation-delay:${Math.random() * .6}s;animation-duration:${1.6 + Math.random() * 1.4}s;transform:rotate(${Math.random() * 360}deg)`; c.appendChild(s); } root.appendChild(c); setTimeout(() => c.remove(), 3600); }
  function teclas(e) {
    if (!root || root.hidden || !S) return;
    if (e.key === 'Escape') { cerrar(); return; }
    if (S.fase === 'pregunta') { const k = LETRAS.indexOf(e.key.toUpperCase()), d = parseInt(e.key, 10) - 1; const idx = k >= 0 && k < S.qs[S.i].opts.length ? k : (d >= 0 && d < S.qs[S.i].opts.length ? d : -1); if (idx >= 0) { elegir(idx); e.preventDefault(); } else if (e.key === 'Enter' && S.sel !== null) { comprobar(); e.preventDefault(); } }
    else if (S.fase === 'feedback' && e.key === 'Enter') { siguiente(); e.preventDefault(); }
  }
  window.abrirCuestionarioTeorico = id => abrir(id);
  window.progresoQuiz = leer;
})();
