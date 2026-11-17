/* Evaluador Reto 3 · Formatos Avanzados y Elementos Visuales (100 pts, reparto por dificultad). */
(function (root) {
  const L = root.DocxLib, { norm, near } = L;
  const NIVELES = { 1: [1, 3], 2: [1, 3], 3: [2], 4: [1, 2], 5: [2, 2], 6: [1, 2, 2], 7: [1, 2], 8: [2, 3, 2], 9: [2, 1, 2], 10: [2, 2, 1] };
  const EV = L.crearEvaluador(NIVELES, 100, 70);
  const FIN = /dinero|moneda|finanz|banco|billete|efectivo|cartera|ahorro|inversi|cuenta|pago|dolar|euro|econom|capital|credito|tarjeta|cajero|bolsa|grafic|tendencia|cifra|cheque|oro|lingote|tesoro|money|cash|coin|bank|financ|dollar|wallet|saving|invest|payment|credit|chart|graph|gold|treasure|currency/;

  function evaluar(files, ctx) {
    ctx = ctx || {}; const M = L.parse(files, ctx), E = L.crearEvaluador(NIVELES, 100, 70), S = E.S, add = E.add, P = M.paras;
    const find = pre => P.find(p => p.n.startsWith(pre));
    const h1 = find('1. fundamentos'), h2 = find('2. comparativa'), h3 = find('3. estrategias');
    const first = P.find(p => p.text.trim().length >= 60 && !M.isHeadingStyle(p, 1) && !M.isHeadingStyle(p, 2) && !M.inTable(p) && !/^\d\. /.test(p.text));
    const hShapes = M.allHeaders.flatMap(h => M.shapesOf(h.xml)), fShapes = M.allFooters.flatMap(f => M.shapesOf(f.xml));
    const bodyTxt = P.filter(p => p.text.trim()).length;

    // ===== ÁREA 1: ENCABEZADO Y PIE =====
    add(1, 1, 'Encabezado «Con bandas» con el título', [
      S(M.allHeaders.some(h => norm(h.text).includes('inversiones inteligentes')), 'El encabezado contiene «INVERSIONES INTELIGENTES»', 'Insertar › Encabezado y escribe el texto solicitado.'),
      S(hShapes.some(s => s.prst.includes('rect') && s.fill && s.txbx.some(t => t.includes('inversiones inteligentes'))), 'Estilo de encabezado «Con bandas» (banda de color con el título)', 'Insertar › Encabezado › Con bandas.')]);
    add(2, 1, 'Número de página «Círculo» en la parte inferior', [
      S(M.allFooters.some(f => /<w:instrText[^>]*>\s*PAGE\b|w:instr="\s*PAGE\b/.test(f.xml)), 'Número de página en el pie del documento', 'Insertar › Número de página › Final de página.'),
      S(fShapes.some(s => s.prst.includes('ellipse') && s.pageField), 'Formato «Círculo» (número dentro de un círculo)', 'Insertar › Número de página › Final de página › Círculo.')]);

    // ===== ÁREA 2: DISEÑO DE PÁGINA =====
    const mod = s => near(s.mar.top, 1440, 10) && near(s.mar.bottom, 1440, 10) && near(s.mar.left, 1080, 10) && near(s.mar.right, 1080, 10);
    const cover = M.sections.filter(mod).reduce((a, s) => a + s.paras.filter(p => p.text.trim()).length, 0);
    add(3, 2, 'Márgenes «Moderado»', [S(mod(M.sections[0]) && bodyTxt > 0 && cover / bodyTxt >= 0.6, 'Márgenes Moderado (2.54 cm arriba/abajo y 1.91 cm a los lados)', 'Disposición › Márgenes › Moderado.')]);
    const bg = M.background, bgc = (L.wat(bg, 'color') || '').toUpperCase();
    add(4, 2, 'Color de página «Gris claro, Fondo 2»', [
      S(bg && bgc && bgc !== 'FFFFFF' && bgc !== 'AUTO', 'Se aplicó un color de página', 'Diseño › Color de página.'),
      S(bg && ((L.wat(bg, 'themeColor') === 'background2' && !L.wat(bg, 'themeShade') && !L.wat(bg, 'themeTint')) || bgc === (M.themeColors.lt2 || 'E8E8E8')), 'El color es «Gris claro, Fondo 2» de los colores del tema', 'Diseño › Color de página › Colores del tema › Gris claro, Fondo 2.')]);
    const t3 = h3 ? P.filter(p => p.idx > h3.idx && p.text.trim()) : [], t12 = P.filter(p => p.text.trim() && h3 && p.idx < h3.idx);
    const colOf = p => M.sections[p.sec].cols, col3 = t3.length >= 5 && t3.every(p => colOf(p) === 2);
    add(5, 2, 'Dos columnas en el Tema 3', [
      S(col3, 'Todos los párrafos del Tema 3 están en dos columnas', 'Selecciona los párrafos bajo el título del Tema 3 › Disposición › Columnas › Dos.'),
      S(col3 && t12.length > 0 && t12.every(p => colOf(p) === 1), 'Los temas 1 y 2 permanecen en una columna', 'Aplica las columnas solo a la selección del Tema 3.')]);

    // ===== ÁREA 3: REFERENCIAS Y PROPIEDADES =====
    const fnT = 'es importante diferenciar entre el rendimiento nominal y el rendimiento real', fn = M.footnotes.find(f => f.n.includes(fnT)), ref = fn && M.fnRefs.find(r => r.id === fn.id);
    const atEnd = ref && first && ref.para === first.idx && first.raw.lastIndexOf('<w:footnoteReference') > first.raw.lastIndexOf('<w:t>') && first.raw.lastIndexOf('<w:footnoteReference') > first.raw.lastIndexOf('<w:t ');
    add(6, 3, 'Nota al pie en el primer párrafo', [
      S(M.footnotes.length > 0 && M.fnRefs.length > 0, 'Existe una nota al pie', 'Referencias › Insertar nota al pie.'),
      S(!!fn, 'Texto exacto de la nota al pie', 'Escribe: «Es importante diferenciar entre el rendimiento nominal y el rendimiento real.»'),
      S(!!atEnd, 'La nota está al final del primer párrafo', 'Cursor al final del primer párrafo antes de insertar la nota.')]);
    const desc = M.core.description;
    add(7, 3, 'Propiedad «Comentarios» con tu nombre completo', [
      S(L.nombreValido(desc), 'La propiedad Comentarios tiene nombre y apellido completos', 'Archivo › Información › Mostrar todas las propiedades › Comentarios.'),
      S(L.coincideNombre(ctx.nombre || '', desc), 'El nombre coincide con el registrado al iniciar el reto', 'Usa exactamente el nombre que escribiste al comenzar.')]);

    // ===== ÁREA 4: ILUSTRACIONES Y ELEMENTOS =====
    const q = M.drawings.find(d => !d.nested && d.txbx.some(t => t.n.includes('a mayor liquidez, menor rendimiento esperado'))), qTxt = q ? q.txbx.map(t => t.n).join(' ') : M.legacyTx.join(' ');
    const pLast1 = P.find(p => p.n.startsWith('a mayor inmovilizacion'));
    add(8, 4, 'Cuadro de texto «Cotización de filigrana»', [
      S(qTxt.includes('a mayor liquidez, menor rendimiento esperado') && qTxt.includes('mayor compensacion exigida por el mercado') && qTxt.includes('principio fundamental de las finanzas'), 'Cuadro de texto con la cita y su atribución', 'Escribe la cita completa y «Principio fundamental de las finanzas».'),
      S(q && q.txbx.some(t => t.hasPic), 'Formato «Cotización de filigrana» (incluye las comillas gráficas)', 'Insertar › Cuadro de texto › Cotización de filigrana.'),
      S(q && pLast1 && h2 && q.para >= pLast1.idx && q.para < h2.idx, 'Ubicado al final del último párrafo del Tema 1', 'Cursor al final del último párrafo del Tema 1 antes de insertar el cuadro.')]);
    const ic = M.drawings.find(d => !d.nested && d.svg);
    add(9, 4, 'Ícono de dinero o finanzas al inicio del primer párrafo', [
      S(!!ic, 'Ícono insertado desde Insertar › Íconos', 'Insertar › Íconos y busca «dinero» o «finanzas».'),
      S(ic && (!ic.descr || FIN.test(norm(ic.descr))), 'El ícono está relacionado con dinero o finanzas', 'Elige un ícono de dinero, banco, gráficos o finanzas.'),
      S(ic && first && ic.para === first.idx && ic.hostPos < first.raw.search(/<w:t[ >]/), 'Ubicado al inicio del primer párrafo', 'Cursor antes de la primera palabra del primer párrafo.')]);
    const sc = M.drawings.find(d => !d.nested && d.prst.includes('horizontalScroll'));
    add(10, 4, 'Forma «Pergamino: Horizontal» con el texto FINANZAS', [
      S(!!sc, 'Forma «Pergamino: Horizontal» insertada', 'Insertar › Formas › Cintas y estrellas › Pergamino: Horizontal.'),
      S(sc && sc.txbx.some(t => t.n.includes('finanzas')), 'La forma contiene el texto FINANZAS', 'Clic derecho en la forma › Agregar texto › FINANZAS.'),
      S(sc && sc.para >= P.length - 4 && (!h3 || sc.para > h3.idx), 'Ubicada al final del documento', 'Inserta la forma al final del documento.')]);
    return E.cerrar(L.nombreValido(desc) ? desc : '');
  }
  root.CHECKER_RETO3 = { evaluar, PUNTOS: EV.PUNTOS, NOMBRE_NIVEL: L.NOMBRE_NIVEL };
  if (typeof module !== 'undefined') module.exports = root.CHECKER_RETO3;
})(typeof window !== 'undefined' ? window : globalThis);
