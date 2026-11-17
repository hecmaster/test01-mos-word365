/* Evaluador Reto 2 · Estilos, Imágenes y Elementos Especiales (100 pts, reparto por dificultad). */
(function (root) {
  const L = root.DocxLib, { norm } = L;
  const NIVELES = { 1: [1, 1, 1, 1], 2: [2, 2, 1], 3: [2], 4: [1, 1], 5: [1, 2], 6: [1, 2, 1], 7: [3], 8: [2, 1], 9: [1, 3, 1], 10: [1, 2, 1, 1] };
  const EV = L.crearEvaluador(NIVELES, 100, 70);
  const GRISES = ['C0C0C0', 'BFBFBF', 'D9D9D9', 'D0CECE', 'D8D8D8', 'CCCCCC'];

  function evaluar(files, ctx) {
    ctx = ctx || {}; const M = L.parse(files, ctx), E = L.crearEvaluador(NIVELES, 100, 70), S = E.S, add = E.add, P = M.paras;
    const realRuns = p => p.runs.filter(r => r.text.trim());
    const find = pre => P.find(p => p.n.startsWith(pre));

    // ===== ÁREA 1: ESTILOS Y DISEÑO =====
    const heads = ['1. elementos clave', '2. la regla de oro', '3. tipos principales', '4. metricas clave'].map(find);
    add(1, 1, 'Estilo Título 2 en los cuatro subtítulos', heads.map((h, i) => S(h && M.isHeadingStyle(h, 2), 'Título 2 en «' + ['1. Elementos clave de un bono', '2. La regla de oro del mercado de renta fija', '3. Tipos principales de bonos', '4. Métricas clave para analizar antes de invertir'][i] + '»', 'Selecciona el subtítulo › Inicio › Estilos › Título 2.')));
    const h1 = M.styleByName('heading 1'), h2 = M.styleByName('heading 2'), ti = M.styleByName('title');
    const shaded = s => !!s && /<w:shd\b[^>]*w:fill="(?!auto|FFFFFF)[0-9A-Fa-f]{6}"/.test(s.x);
    add(2, 1, 'Diseño del documento «Sombreado»', [
      S(shaded(h1), 'Títulos de nivel 1 con sombreado (conjunto de estilos Sombreado)', 'Diseño › Formato del documento › Sombreado.'),
      S(shaded(h2), 'Títulos de nivel 2 con sombreado', 'Aplica el conjunto «Sombreado» completo desde Diseño › Formato del documento.'),
      S(ti && /<w:caps\b(?![^>]*w:val="(?:0|false)")/.test(ti.x), 'Estilo Título en mayúsculas propio de «Sombreado»', 'Elige exactamente el diseño llamado «Sombreado».')]);
    const ph = 'desde investment grade hasta high yield', pp = P.find(p => p.n.includes(ph));
    const phRuns = pp ? pp.runs.filter(r => { const t = norm(r.text).trim(); return t.length > 2 && ph.includes(t); }) : [];
    const baseEf = pp ? M.efRun({ style: pp.style, pprNoRpr: '' }, null) : null;
    add(3, 1, 'Borrar el formato de la frase «desde Investment Grade hasta High Yield»', [
      S(pp && phRuns.length > 0 && phRuns.every(r => { const e = M.efRun(pp, r); return !e.b && !e.i && !e.u && !e.highlight && !e.shd && (!e.color || e.color === '000000') && e.sz === baseEf.sz && (e.font || '') === (baseEf.font || ''); }), 'La frase quedó sin formato especial (sin negrita, cursiva ni color)', 'Selecciona la frase › Inicio › Borrar todo el formato (A con borrador).')]);

    // ===== ÁREA 2: TEXTO Y SÍMBOLOS =====
    add(4, 2, 'Símbolo de Copyright antes de Moody’s', [
      S(/©/.test(M.text), 'Existe el símbolo ©', 'Insertar › Símbolo › Más símbolos › ©.'),
      S(/©\s*moody/.test(M.text), 'El símbolo © está justo antes de «Moody’s»', 'Cursor a la izquierda de «Moody’s» antes de insertar el símbolo.')]);
    const fo = P.find(p => /<m:oMath/.test(p.raw) && p.n.includes('precio del bono')), fr = fo ? fo.runs.filter(r => r.text.trim()) : [];
    const hl = r => (r.rpr.match(/<w:highlight w:val="(\w+)"/) || [])[1], sh = r => (r.rpr.match(/<w:shd\b[^>]*w:fill="([0-9A-Fa-f]{6})"/) || [])[1];
    add(5, 2, 'Resaltado gris del 25 % en la fórmula', [
      S(fr.length > 0 && fr.every(r => (hl(r) && hl(r) !== 'none') || (sh(r) && !/^FFFFFF$/i.test(sh(r)))), 'Toda la fórmula está resaltada', 'Selecciona toda la fórmula › Inicio › Color de resaltado de texto.'),
      S(fr.length > 0 && fr.every(r => hl(r) === 'lightGray' || (sh(r) && GRISES.includes(sh(r).toUpperCase()))), 'El tono es Gris 25 %', 'Elige «Gris 25 %» en la paleta de resaltado.')]);

    // ===== ÁREA 3: IMÁGENES =====
    const img = M.drawings.find(d => !d.nested && d.isPic && /^(jpe?g|png)$/.test(d.ext) && d.cy && Math.abs(d.cx / d.cy - 1584 / 672) < 0.2) || null;
    const p1 = P.find(p => p.n.startsWith('un bono es un instrumento')), p2 = P.find(p => p.n.startsWith('para operar e invertir'));
    add(6, 3, 'Insertar bono.jpg después del primer párrafo y centrarlo', [
      S(!!img, 'Imagen bono.jpg insertada', 'Insertar › Imágenes › Este dispositivo › bono.jpg.'),
      S(img && p1 && p2 && img.para >= p1.idx && img.para < p2.idx && !(img.para === p1.idx && img.kind === 'inline' && P[img.para].n.length > 0 && P[img.para].raw.indexOf('<w:drawing') < P[img.para].raw.indexOf('<w:t')), 'La imagen está después del primer párrafo', 'Cursor al final del primer párrafo › Enter › inserta la imagen en la línea nueva.'),
      S(img && ((img.kind === 'inline' && M.efPara(P[img.para]).jc === 'center') || (img.kind === 'anchor' && img.hAlign === 'center')), 'La imagen está centrada', 'Selecciona la imagen › Inicio › Centrar.')]);
    add(7, 3, 'Saturación de color del 66 % en la imagen', [S(img && img.sat === 66000, 'Saturación 66 % aplicada a la imagen', 'Formato de imagen › Color › Saturación de color › 66 %.')]);

    // ===== ÁREA 4: TABLAS Y LISTAS =====
    const tb = M.tables.find(t => t.text.includes('soberano')) || null;
    add(8, 4, 'Estilo «Tabla de cuadrícula 2»', [
      S(tb && /^grid table 2\b/.test(tb.styleName), 'Estilo de la familia «Tabla de cuadrícula 2»', 'Diseño de tabla › Estilos de tabla › Tabla de cuadrícula 2.'),
      S(tb && tb.styleName === 'grid table 2', 'Es «Tabla de cuadrícula 2» (sin énfasis de color)', 'Elige la variante sin color de énfasis.')]);
    const fin = P.filter(p => p.text.trim() && !M.inTable(p)).slice(-3), ni = fin.map(p => M.numInfo(p));
    const picAll = P.filter(p => { const n = M.numInfo(p); return n && n.pic !== null; });
    add(9, 4, 'Viñetas con la imagen dinero.png en los tres últimos párrafos', [
      S(fin.length === 3 && ni.every(n => n && n.fmt === 'bullet'), 'Los tres últimos párrafos tienen viñetas', 'Selecciona solo los tres últimos párrafos › Inicio › Viñetas.'),
      S(fin.length === 3 && ni.every(n => n && n.pic !== null), 'Las viñetas son de imagen (dinero.png)', 'Viñetas › Definir nueva viñeta › Imagen › dinero.png.'),
      S(picAll.length === 3 && picAll.every(p => fin.includes(p)), 'Solo esos tres párrafos usan la viñeta de imagen', 'No cambies las viñetas del resto del documento.')]);

    // ===== ÁREA 5: REVISIÓN =====
    const title = find('bono inversiones inteligentes') || P[0], cm = M.comments.find(c => c.nAnchor.includes('inteligentes')) || M.comments[0];
    add(10, 5, 'Comentario con el nombre completo en «Inteligentes»', [
      S(M.comments.length > 0, 'Existe un comentario en el documento', 'Revisar › Nuevo comentario.'),
      S(M.comments.some(c => c.nAnchor.includes('inteligentes')), 'El comentario está sobre la palabra «Inteligentes»', 'Selecciona la palabra «Inteligentes» del título antes de comentar.'),
      S(cm && L.nombreValido(cm.text), 'El comentario contiene nombre y apellido completos', 'Escribe tu nombre y apellido en el comentario.'),
      S(cm && L.coincideNombre(ctx.nombre || '', cm.text), 'El nombre coincide con el registrado al iniciar el reto', 'Usa exactamente el nombre que escribiste al comenzar.')]);
    return E.cerrar(cm && L.nombreValido(cm.text) ? cm.text.trim() : '');
  }
  root.CHECKER_RETO2 = { evaluar, PUNTOS: EV.PUNTOS, NOMBRE_NIVEL: L.NOMBRE_NIVEL };
  if (typeof module !== 'undefined') module.exports = root.CHECKER_RETO2;
})(typeof window !== 'undefined' ? window : globalThis);
