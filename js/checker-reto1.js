/* Evaluador Reto 1 · Formato y Estilos (100 pts, reparto por dificultad). Usa DocxLib. */
(function (root) {
  const L = root.DocxLib, { norm, near, isBlue } = L;
  const NIVELES = { 1: [1, 1], 2: [2, 2, 1], 3: [2, 1, 1, 1, 1, 2, 2, 1], 4: [1, 2, 1, 1, 2, 2], 5: [1, 2, 3], 6: [1, 1, 2, 2], 7: [1, 1, 2], 8: [1, 2, 3], 9: [2, 2, 3, 1], 10: [1, 2, 3], 11: [2, 2, 2] };
  const EV = L.crearEvaluador(NIVELES, 100, 70);
  const PLAZOS_BASE = 7;                       // apariciones de "plazo" en el documento base
  const LOGO_BASES = [4194048, 4615200];       // ancho natural del logo a 300 ppp y ancho de inserción de Word (EMU)

  function evaluar(files, ctx) {
    ctx = ctx || {}; const M = L.parse(files, ctx), S = EV.S; const E = L.crearEvaluador(NIVELES, 100, 70);
    const add = E.add, P = M.paras;
    const textRuns = p => p.runs.filter(r => r.text.trim());
    const allRuns = (p, f) => { const rs = textRuns(p); return rs.length > 0 && rs.every(r => f(M.efRun(p, r), r)); };
    const find = pre => P.find(p => p.n.startsWith(pre));

    // ===== ÁREA 1: PÁGINA Y SECCIONES =====
    const pres = P.find(p => /^presentado por\s*:/.test(p.n));
    let nombre = ''; if (pres) { nombre = pres.text.replace(/^\s*presentado\s+por\s*:/i, '').split(/fecha\s*:/i)[0].trim(); }
    add(1, 1, 'Identificación del estudiante (portada)', [
      S(pres && L.nombreValido(nombre), 'Se reemplazó «Estudiante» por tu nombre y apellido completos', 'En la portada escribe tu nombre completo después de «Presentado por:».'),
      S(pres && L.coincideNombre(ctx.nombre || '', nombre), 'El nombre coincide con el registrado al iniciar el reto', 'Usa exactamente el nombre que escribiste al comenzar.')]);
    const contAnexo = P.find(p => p.n.startsWith('en esta seccion final')), secA = contAnexo ? M.sections[contAnexo.sec] : null, h5 = find('5. anexos'), h1 = find('1. introduccion');
    add(2, 1, 'Salto de sección y orientación horizontal en Anexos', [
      S(secA && secA.index > 0 && secA.type === 'nextPage' && (!h5 || h5.sec === secA.index || h5.sec === secA.index - 1), 'Salto de sección (página siguiente) antes de «5. Anexos»', 'Cursor al inicio de «5. Anexos…» › Disposición › Saltos › Página siguiente.'),
      S(secA && secA.landscape, 'La sección de Anexos tiene orientación horizontal', 'Dentro de esa sección: Disposición › Orientación › Horizontal.'),
      S(secA && secA.landscape && M.sections.filter(s => s.index < secA.index).every(s => !s.landscape) && h1 && !M.sections[h1.sec].landscape, 'El resto del documento permanece en vertical', 'Aplica la orientación solo a la sección de Anexos, no a todo el documento.')]);
    const hdrPs = M.allHeaders.flatMap(h => M.parasOf(h.xml)).filter(p => p.text.trim());
    const hp = hdrPs[0], ftrPs = M.allFooters.flatMap(f => M.parasOf(f.xml).map(p => ({ p, field: /<w:instrText[^>]*>\s*PAGE\b|w:instr="\s*PAGE\b/.test(f.xml) })));
    const wm = M.allHeaders.map(h => h.xml).join('\n'), wmTag = (wm.match(/<v:shape\b[^>]*PowerPlusWaterMarkObject[^>]*>/) || wm.match(/<v:shape\b[^>]*type="#_x0000_t136"[^>]*>/) || [''])[0];
    const wmText = norm(wm).includes('confidencial') && /<v:textpath\b[^>]*string="[^"]*CONFIDENCIAL/i.test(wm);
    add(3, 1, 'Encabezado, pie de página y marca de agua', [
      S(M.sections.some(s => s.titlePg), 'Primera página diferente activada', 'Disposición › clic en el iniciador de Configurar página › Diseño › «Primera página diferente».'),
      S(!!hp, 'El encabezado contiene texto', 'Insertar › Encabezado y escribe un texto (por ejemplo, el título del informe).'),
      S(hp && M.efPara(hp).jc === 'right', 'Encabezado alineado a la derecha', 'Inicio › Alinear a la derecha en el encabezado.'),
      S(hp && allRuns(hp, e => e.i), 'Encabezado en cursiva', 'Selecciona el texto del encabezado › Cursiva (Ctrl+K).'),
      S(hp && allRuns(hp, e => isBlue(e.color)), 'Encabezado en color azul', 'Selecciona el texto del encabezado › Color de fuente azul.'),
      S(ftrPs.some(f => (f.p.text.trim() || f.field) && M.efPara(f.p).jc === 'center'), 'Pie de página con contenido centrado', 'Insertar › Número de página (o escribe un texto) y usa Centrar en el pie.'),
      S(wmText, 'Marca de agua «CONFIDENCIAL»', 'Diseño › Marca de agua › CONFIDENCIAL 1.'),
      S(wmText && /rotation:\s*(315|-45)/.test(wmTag), 'Marca de agua en diagonal (CONFIDENCIAL 1)', 'Elige «CONFIDENCIAL 1» (diagonal), no «CONFIDENCIAL 2».')]);

    // ===== ÁREA 2: ESTILOS Y EFECTOS =====
    const heads = ['1. introduccion', '2. fases', '3. comparativa', '4. proyeccion', '5. anexos'].map(find), H = heads.every(Boolean);
    const hAll = f => H && heads.every(f);
    add(4, 2, 'Modificar el estilo Título 1', [
      S(hAll(p => M.isHeadingStyle(p, 1)), 'Estilo «Título 1» aplicado a los 5 encabezados numerados', 'Aplica Inicio › Estilos › Título 1 a cada uno de los 5 encabezados.'),
      S(hAll(p => allRuns(p, e => (e.font || '').toLowerCase() === 'arial')), 'Fuente Arial', 'Clic derecho en Título 1 › Modificar › Fuente: Arial.'),
      S(hAll(p => allRuns(p, e => e.sz === 32)), 'Tamaño 16 pt', 'En Modificar estilo, tamaño 16.'),
      S(hAll(p => allRuns(p, e => e.b)), 'Negrita', 'Activa Negrita en el estilo Título 1.'),
      S(hAll(p => allRuns(p, e => isBlue(e.color))), 'Color azul', 'Elige un azul de la paleta de colores de fuente.'),
      S(hAll(p => M.efPara(p).after === 120), 'Espacio posterior de 6 pt', 'Modificar › Formato › Párrafo › Espaciado posterior: 6 pt.')]);
    const tit = find('informe de evaluacion y gestion de riesgos'), shadow = tit ? (tit.runs.map(r => (r.rpr.match(/<w14:shadow\b[^>]*>/) || [])[0]).find(Boolean)) : null;
    add(5, 2, 'Efecto de sombra y mayúsculas en el título', [
      S(tit && (tit.text === tit.text.toUpperCase() || allRuns(tit, e => e.caps)), 'Título en MAYÚSCULAS', 'Inicio › Cambiar mayúsculas y minúsculas › MAYÚSCULAS (o Fuente › Mayúsculas).'),
      S(!!shadow, 'Efecto de sombra aplicado al título', 'Inicio › Efectos de texto y tipografía › Sombra › Sombra externa.'),
      S(shadow && (d => d >= 4500000 && d <= 6300000)(Number((shadow.match(/w14:dir="(\d+)"/) || [])[1])), 'Sombra con desplazamiento hacia abajo', 'Sombra externa › «Desplazamiento abajo».')]);
    const cntT = f => P.reduce((a, p) => a + p.runs.reduce((b, r) => b + (f(M.efRun(p, r)) ? (norm(r.text).match(/\btiempo\b/g) || []).length : 0), 0), 0);
    add(6, 2, 'Buscar y reemplazar «plazo» por «tiempo»', [
      S(!/\bplazo\b/.test(M.text), 'Ya no existe la palabra «plazo»', 'Inicio › Reemplazar (Ctrl+L) › Reemplazar todos.'),
      S((M.text.match(/\btiempo\b/g) || []).length >= PLAZOS_BASE + 2, '«tiempo» reemplazó cada «plazo» (sin borrar texto)', 'Reemplaza la palabra, no la frase completa.'),
      S(cntT(e => e.b) >= PLAZOS_BASE, 'Todas las sustituciones en negrita', 'En Reemplazar › Más › Formato › Fuente › Negrita.'),
      S(cntT(e => isBlue(e.color)) >= PLAZOS_BASE, 'Todas las sustituciones en color azul', 'En Reemplazar › Más › Formato › Fuente › Color azul.')]);

    // ===== ÁREA 3: PÁRRAFOS Y LISTAS =====
    const cuerpo = P.filter(p => p.text.trim().length >= 50 && !M.numInfo(p) && !M.inTable(p) && !/\|/.test(p.text) && !heads.includes(p) && !M.isHeadingStyle(p, 1) && p !== tit && !p.n.startsWith('modulo'));
    const cAll = f => cuerpo.length >= 4 && cuerpo.every(p => f(M.efPara(p)));
    add(7, 3, 'Formato de párrafo (justificado, 1.15 y sangría)', [
      S(cAll(e => e.jc === 'both'), 'Párrafos de texto justificados', 'Inicio › Justificar (Ctrl+J). Usa «Copiar formato» para replicar.'),
      S(cAll(e => near(e.line, 276, 1) && (!e.lineRule || e.lineRule === 'auto')), 'Interlineado 1.15', 'Inicio › Espaciado de línea y párrafo › 1.15.'),
      S(cAll(e => near(e.firstLine, 709, 15)), 'Sangría de primera línea de 1.25 cm', 'Párrafo › Sangría especial › Primera línea › 1.25 cm.')]);
    const hF = find('2. fases'), hC = find('3. comparativa'), lista = (hF && hC) ? P.filter(p => p.idx > hF.idx && p.idx < hC.idx && p.text.trim() && !p.n.startsWith('a continuacion')) : [];
    const info = lista.map(p => ({ p, ni: M.numInfo(p) })), fases = info.filter(x => x.p.n.startsWith('fase de'));
    add(8, 3, 'Lista multinivel (1. / 1.1. / 1.1.1.)', [
      S(lista.length >= 9 && info.every(x => x.ni && x.ni.fmt === 'decimal'), 'Las 9 líneas están numeradas', 'Selecciona las líneas › Inicio › Lista multinivel.'),
      S(fases.length === 3 && fases.every(f => { const k = info.indexOf(f), nxt = info.slice(k + 1, k + 3); return f.ni && nxt.length === 2 && nxt.every(n => n.ni && n.ni.ilvl > f.ni.ilvl && !n.p.n.startsWith('fase de')); }), 'Cada fase tiene jerarquía superior a sus actividades (niveles con Tab)', 'Usa la tecla Tab para bajar un nivel en las actividades de cada fase.'),
      S(info.length > 0 && info.every(x => x.ni && x.ni.abs && x.ni.abs.multi === 'multilevel') && info.some(x => x.ni && /%1\.?.*%2/.test(x.ni.text)), 'Formato de numeración multinivel (1.1., 1.1.1.)', 'Elige la lista multinivel con numeración 1. / 1.1. / 1.1.1. de la biblioteca.')]);

    // ===== ÁREA 4: TABLAS E ILUSTRACIONES =====
    const tb = M.tables.find(t => t.firstRow.includes('instrumento')) || null, last = tb && tb.rows[tb.rows.length - 1];
    add(9, 4, 'Tabla con estilo y fila final combinada', [
      S(tb && tb.nCols >= 4 && tb.nRows >= 6 && tb.text.includes('renta fija') && tb.text.includes('fondos mutuos'), 'Texto convertido en tabla (4 columnas + fila final)', 'Selecciona las líneas › Insertar › Tabla › Convertir texto en tabla (separador: |).'),
      S(tb && tb.styleName === 'grid table 4 accent 1', 'Estilo «Tabla de cuadrícula 4 – Énfasis 1»', 'Diseño de tabla › Estilos de tabla.'),
      S(tb && last && last.length === 1 && last[0].span >= tb.nCols && norm(last[0].text) === 'informacion recopilada', 'Fila final combinada con «INFORMACIÓN RECOPILADA»', 'Agrega una fila al final › selecciona sus celdas › Disposición › Combinar celdas.'),
      S(tb && last && last.length === 1 && /<w:jc w:val="center"/.test(last[0].xml), 'Texto de la fila final centrado', 'Inicio › Centrar.')]);
    const img = M.drawings.find(d => !d.nested && d.isPic && /^(jpe?g|png)$/.test(d.ext)) || null, pw = M.sections[0] ? M.sections[0].w : 12240;
    add(10, 4, 'Logo corporativo (esquina superior derecha, escala 50 %)', [
      S(!!img, 'Imagen del logo insertada en el documento', 'Insertar › Imágenes › Este dispositivo › logo_Corporativo.jpg.'),
      S(img && ((img.kind === 'anchor' && (img.hAlign === 'right' || (img.hOff !== null && img.hOff > pw * 635 * 0.5)) && (img.vAlign === 'top' || (img.vOff !== null && img.vOff < 1500000))) || (img.kind === 'inline' && img.para <= 6 && M.efPara(P[img.para]).jc === 'right')), 'Ubicado en la parte superior derecha', 'Formato de imagen › Posición › Parte superior derecha con ajuste de texto cuadrado.'),
      S(img && LOGO_BASES.some(b => near(img.cx / b, 0.5, 0.05)) && near(img.cy / img.cx, 768 / 1376, 0.03), 'Escala del 50 % con proporciones conservadas', 'Formato de imagen › Tamaño › Escala alto y ancho: 50 % (bloquear relación de aspecto).')]);

    // ===== ÁREA 5: REFERENCIAS =====
    const fnTxt = 'es el interes calculado sobre el capital inicial y tambien sobre los intereses ya ganados', fn = M.footnotes.find(f => f.n.includes(fnTxt)), ref = fn && M.fnRefs.find(r => r.id === fn.id);
    add(11, 5, 'Nota al pie en «Interés Compuesto»', [
      S(M.footnotes.length > 0 && M.fnRefs.length > 0, 'Existe una nota al pie', 'Referencias › Insertar nota al pie.'),
      S(!!fn, 'Texto literal de la nota al pie', 'Escribe: «Es el interés calculado sobre el capital inicial y también sobre los intereses ya ganados.»'),
      S(ref && P[ref.para] && P[ref.para].n.includes('interes compuesto'), 'La nota está en el párrafo de «Interés Compuesto»', 'Coloca el cursor después de «Interés Compuesto» antes de insertar la nota.')]);
    return E.cerrar(L.nombreValido(nombre) ? nombre : '');
  }
  root.CHECKER_RETO1 = { evaluar, PUNTOS: EV.PUNTOS, NOMBRE_NIVEL: L.NOMBRE_NIVEL };
  if (typeof module !== 'undefined') module.exports = root.CHECKER_RETO1;
})(typeof window !== 'undefined' ? window : globalThis);
