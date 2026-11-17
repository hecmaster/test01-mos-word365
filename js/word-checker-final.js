/* Evaluador del Proyecto Final (Word 365) - lee el XML interno del .docx.
   evaluarProyectoFinal(files, ctx) -> { items, total, nombreDetectado }
   files: { 'word/document.xml': '...', ... } (texto) | ctx: { nombre, pdf:{name,size,header} } */
(function (root) {
  const TOTAL_PTS = 1000, APROBADO = 700;
  // Nivel de dificultad de cada subcriterio (1 = Básico, 2 = Intermedio, 3 = Avanzado). El orden coincide con los criterios de cada item.
  const NIVELES = { 1: [1, 1, 1, 2], 2: [2, 3], 3: [2, 2], 4: [1, 1, 2], 5: [2, 2, 1, 2], 6: [1, 1, 1], 7: [1, 1], 8: [2, 2], 9: [1, 2, 2], 10: [1, 2],
    11: [2, 2], 12: [1, 2], 13: [3, 2], 14: [2, 1, 1], 15: [1, 2], 16: [2, 3], 17: [2, 2], 18: [3], 19: [2, 2], 20: [3, 2] };
  const NOMBRE_NIVEL = { 1: 'Básico', 2: 'Intermedio', 3: 'Avanzado' };
  // Reparto por mayor residuo: la suma de todos los subcriterios es exactamente 1000
  const PUNTOS = (() => {
    const ids = Object.keys(NIVELES).map(Number), W = ids.reduce((a, i) => a + NIVELES[i].reduce((x, y) => x + y, 0), 0);
    const flat = []; ids.forEach(i => NIVELES[i].forEach((w, k) => flat.push({ i, k, w, ex: TOTAL_PTS * w / W })));
    flat.forEach(f => f.p = Math.floor(f.ex)); let rest = TOTAL_PTS - flat.reduce((a, f) => a + f.p, 0);
    [...flat].sort((a, b) => (b.ex - b.p) - (a.ex - a.p) || a.i - b.i || a.k - b.k).slice(0, rest).forEach(f => f.p++);
    const out = {}; ids.forEach(i => { out[i] = { pts: [], niveles: NIVELES[i] }; }); flat.forEach(f => { out[f.i].pts[f.k] = f.p; });
    ids.forEach(i => { out[i].total = out[i].pts.reduce((a, b) => a + b, 0); const m = NIVELES[i].reduce((a, b) => a + b, 0) / NIVELES[i].length; out[i].nivel = m >= 2.5 ? 3 : m >= 1.5 ? 2 : 1; });
    return out;
  })();

  const norm = s => (s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[“”«»]/g, '"').replace(/[‘’]/g, "'").replace(/\s+/g, ' ').trim().toLowerCase();
  const dec = s => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
  const textOf = x => dec([...x.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map(m => m[1]).join(''));
  const attr = (tag, n) => { const m = (tag || '').match(new RegExp('w:' + n + '="([^"]*)"')); return m ? m[1] : null; };
  const all = (re, s) => [...(s || '').matchAll(re)].map(m => m[0]);
  const near = (v, t, tol) => v !== null && Math.abs(Number(v) - t) <= tol;

  function evaluarProyectoFinal(files, ctx) {
    ctx = ctx || {};
    const F = n => files[n] || '';
    const docXml = F('word/document.xml'), stylesXml = F('word/styles.xml'), numXml = F('word/numbering.xml');
    const names = Object.keys(files);
    const group = re => names.filter(n => re.test(n)).map(n => files[n]).join('\n');
    const headers = names.filter(n => /^word\/header\d*\.xml$/.test(n)).map(n => files[n]);
    const footers = names.filter(n => /^word\/footer\d*\.xml$/.test(n)).map(n => files[n]);

    // ---- cuerpo, párrafos y secciones ----
    const body = docXml.replace(/<mc:Fallback>[\s\S]*?<\/mc:Fallback>/g, '').replace(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g, '');
    const paras = [];
    for (const m of body.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)) {
      const xml = m[0], ppr = (xml.match(/<w:pPr>[\s\S]*?<\/w:pPr>/) || [''])[0];
      paras.push({ xml, ppr, text: textOf(xml), n: norm(textOf(xml)), style: (ppr.match(/<w:pStyle w:val="([^"]+)"/) || [])[1] || '', sect: (ppr.match(/<w:sectPr[ >][\s\S]*?<\/w:sectPr>/) || [null])[0] });
    }
    const sections = []; let cur = [];
    for (const p of paras) { cur.push(p); p.sec = sections.length; if (p.sect) { sections.push({ paras: cur, sect: p.sect }); cur = []; } }
    const fin = body.match(/(<w:sectPr[ >](?:(?!<w:sectPr)[\s\S])*?<\/w:sectPr>)\s*<\/w:body>/);
    cur.forEach(p => p.sec = sections.length);
    sections.push({ paras: cur, sect: fin ? fin[1] : '' });
    const sects = sections.map(s => s.sect).filter(Boolean);
    const pg = s => { const t = (s.match(/<w:pgSz[^>]*>/) || [''])[0]; return { w: attr(t, 'w'), h: attr(t, 'h'), o: attr(t, 'orient') }; };
    const mar = s => { const t = (s.match(/<w:pgMar[^>]*>/) || [''])[0]; return ['top', 'bottom', 'left', 'right'].map(k => attr(t, k)); };

    // ---- estilos ----
    const styles = all(/<w:style\b[^>]*>[\s\S]*?<\/w:style>/g, stylesXml).map(x => ({ x, id: attr(x.match(/<w:style\b[^>]*>/)[0], 'styleId') || (x.match(/w:styleId="([^"]*)"/) || [])[1], name: norm((x.match(/<w:name w:val="([^"]*)"/) || [])[1]) }));
    const styleByName = n => styles.find(s => s.name === n);
    const styleById = i => styles.find(s => s.id === i);
    const h1 = styleByName('heading 1'), h2 = styleByName('heading 2');

    // ---- listas ----
    const numFmt = id => {
      const num = (numXml.match(new RegExp('<w:num w:numId="' + id + '"[^>]*>[\\s\\S]*?</w:num>')) || [''])[0];
      const a = (num.match(/<w:abstractNumId w:val="(\d+)"/) || [])[1]; if (a === undefined) return null;
      const abs = (numXml.match(new RegExp('<w:abstractNum [^>]*w:abstractNumId="' + a + '"[\\s\\S]*?</w:abstractNum>')) || [''])[0];
      const lvl = (abs.match(/<w:lvl w:ilvl="0"[\s\S]*?<\/w:lvl>/) || [''])[0];
      return (lvl.match(/<w:numFmt w:val="([^"]+)"/) || [])[1] || null;
    };
    const numIdOf = p => {
      let m = p.ppr.match(/<w:numId w:val="(\d+)"/);
      if (!m && p.style) { const s = styleById(p.style); m = s && s.x.match(/<w:numId w:val="(\d+)"/); }
      return m && m[1] !== '0' ? m[1] : null;
    };

    // ---- propiedades ----
    const core = F('docProps/core.xml');
    const creator = dec((core.match(/<dc:creator>([^<]*)<\/dc:creator>/) || [, ''])[1]).trim();
    const desc = dec((core.match(/<dc:description>([^<]*)<\/dc:description>/) || [, ''])[1]).trim();
    const nombreOk = n => n.split(/\s+/).filter(w => w.length > 1).length >= 2 && !/escriba|nombre completo|estudiante|usuario|autor/i.test(n);
    const tokens = s => norm(s).split(' ').filter(w => w.length > 2);
    const nombreDetectado = nombreOk(creator) ? creator : (nombreOk(desc) ? desc : (creator || desc || ''));

    // ---- tema 1..4 ----
    const hi = n => paras.findIndex(p => p.n.startsWith('tema ' + n));
    const rango = (a, b) => { const i = hi(a); const j = b === 'anexo' ? paras.findIndex(p => p.n.startsWith('anexo')) : hi(b); return i < 0 ? [] : paras.slice(i + 1, j < 0 ? undefined : j); };
    const tema1 = rango(1, 2).filter(p => p.text.trim()), tema3 = rango(3, 4).filter(p => p.text.trim()), tema4 = rango(4, 'anexo');
    const secCols = p => { const s = sections[p.sec]; return s ? attr((s.sect.match(/<w:cols[^>]*>/) || [''])[0], 'num') : null; };
    const allText = norm(paras.map(p => p.text).join(' '));

    // ---- tablas, imágenes, formas ----
    const tbls = all(/<w:tbl>[\s\S]*?<\/w:tbl>/g, body).map(t => {
      const rows = all(/<w:tr[ >][\s\S]*?<\/w:tr>/g, t);
      return { x: t, rows: rows.length, cols: Math.max(0, ...rows.map(r => all(/<w:tc[ >]/g, r).length)), first: norm(textOf(rows[0] || '')), t: norm(textOf(t)), style: (t.match(/<w:tblStyle w:val="([^"]+)"/) || [])[1] };
    });
    const tbl = tbls.sort((a, b) => (b.rows * b.cols) - (a.rows * a.cols))[0];
    const rels = {}; for (const m of F('word/_rels/document.xml.rels').matchAll(/<Relationship [^>]*>/g)) { const id = (m[0].match(/Id="([^"]+)"/) || [])[1], tg = (m[0].match(/Target="([^"]+)"/) || [])[1]; if (id) rels[id] = tg; }
    const fotos = all(/<w:drawing>[\s\S]*?<\/w:drawing>/g, docXml).map(d => ({ tg: rels[(d.match(/r:embed="([^"]+)"/) || [])[1]] || '', cx: Number((d.match(/<wp:extent cx="(\d+)"/) || [, 0])[1]) })).filter(f => /\.jpe?g$/i.test(f.tg));
    const txbx = all(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g, docXml).map(x => norm(textOf(x)));
    const formas = all(/<wps:wsp[\s\S]*?<\/wps:wsp>/g, docXml);

    const S = (ok, label, pista) => ({ ok: !!ok, label, pista });
    const items = [];
    const add = (id, tarea, titulo, subs) => {
      const P = PUNTOS[id]; subs.forEach((s, k) => { s.pts = P.pts[k]; s.nivel = P.niveles[k]; s.nivelNombre = NOMBRE_NIVEL[P.niveles[k]]; });
      items.push({ id, tarea, titulo, pts: P.total, nivel: P.nivel, nivelNombre: NOMBRE_NIVEL[P.nivel], subs });
    };

    // ===== TAREA 1 =====
    add(1, 1, 'Papel A4, márgenes Estrecho, color de página y tema Faceta', [
      S(sects.length && sects.every(s => { const g = pg(s), a = Math.min(g.w, g.h), b = Math.max(g.w, g.h); return near(a, 11906, 60) && near(b, 16838, 60); }), 'Tamaño de papel A4', 'Disposición > Tamaño > A4.'),
      S(sects.length && sects.every(s => mar(s).every(v => near(v, 720, 20))), 'Márgenes Estrecho (1.27 cm)', 'Disposición > Márgenes > Estrecho.'),
      S((() => { const t = (docXml.match(/<w:background[^>]*>/) || [''])[0]; const c = (attr(t, 'color') || '').toUpperCase(); return t && c !== 'FFFFFF' && c !== 'AUTO' || (t && attr(t, 'themeColor')); })(), 'Color de página distinto de blanco', 'Diseño > Color de página y elige un color del tema.'),
      S(/<a:theme\b[^>]*name="[^"]*facet/i.test(F('word/theme/theme1.xml')), 'Tema "Faceta" aplicado', 'Diseño > Temas > Faceta.')]);
    const land = sections.findIndex(s => { const g = pg(s.sect); return g.o === 'landscape' || Number(g.w) > Number(g.h); });
    const anexo = paras.find(p => p.n.startsWith('anexo'));
    add(2, 1, 'Salto de sección y orientación horizontal en el Anexo', [
      S(sections.length >= 2 && land >= 0, 'Existe una sección con orientación horizontal', 'Disposición > Saltos > Página siguiente, y luego Orientación > Horizontal.'),
      S(anexo && land >= 0 && anexo.sec === land && tema1.length && tema1[0].sec !== land, 'Solo el Anexo está en horizontal (el resto vertical)', 'Coloca el salto de sección justo antes del título "Anexo".')]);
    const col3 = tema3.length && tema3.every(p => secCols(p) === '2');
    add(3, 1, 'Dos columnas en el Tema 3', [
      S(col3, 'Todos los párrafos del Tema 3 en dos columnas', 'Selecciona los párrafos del Tema 3 y usa Disposición > Columnas > Dos.'),
      S(col3 && tema1.length && tema1.every(p => secCols(p) !== '2') && tema4.every(p => secCols(p) !== '2'), 'Los demás temas siguen en una columna', 'Aplica columnas solo a la selección del Tema 3.')]);
    add(4, 1, 'Propiedades del documento con tu nombre', [
      S(nombreOk(creator), 'Autor con nombre completo', 'Archivo > Información > Mostrar todas las propiedades > Autor.'),
      S(nombreOk(desc) && (norm(desc) === norm(creator) || norm(desc).includes(norm(creator)) || norm(creator).includes(norm(desc))), 'Comentarios con el mismo nombre', 'Escribe tu nombre completo en la propiedad Comentarios.'),
      S((() => { const t = tokens(ctx.nombre || ''); const c = norm(creator); return t.length >= 2 && t.filter(w => c.includes(w)).length >= Math.max(2, Math.ceil(t.length * .6)); })(), 'Coincide con el nombre registrado al iniciar el examen', 'Usa exactamente el nombre que escribiste al comenzar.')]);

    // ===== TAREA 2 =====
    const rp = h1 ? h1.x : '';
    add(5, 2, 'Modificar el estilo Título 1', [
      S(/<w:rFonts[^>]*(ascii|hAnsi)="Georgia"/i.test(rp), 'Fuente Georgia', 'Clic derecho en Título 1 > Modificar > Fuente.'),
      S(/<w:sz w:val="32"/.test(rp), 'Tamaño 16 pt', 'En Modificar estilo, tamaño 16.'),
      S(/<w:b(?:\s+w:val="(?:1|true)")?\/>/.test(rp), 'Negrita', 'Activa Negrita en el estilo.'),
      S((() => { const c = (rp.match(/<w:color w:val="([0-9A-Fa-f]{6})"/) || [])[1]; if (!c) return false; const r = parseInt(c.slice(0, 2), 16), g = parseInt(c.slice(2, 4), 16), b = parseInt(c.slice(4, 6), 16); return b > r + 20 && b >= g; })(), 'Color azul', 'Elige un azul de la paleta de colores de fuente.')]);
    const subt = ['instrumentos de renta fija', 'instrumentos de renta variable', 'fondos de inversion'];
    add(6, 2, 'Estilo Título 2 en los tres subtítulos', subt.map(t => S((() => { const p = paras.find(q => q.n === t); return p && h2 && p.style === h2.id; })(), 'Título 2: ' + t.charAt(0).toUpperCase() + t.slice(1), 'Inicio > Estilos > Título 2.')));
    const p1 = tema1[0] || { ppr: '' };
    add(7, 2, 'Formato de párrafo del Tema 1', [
      S(/<w:jc w:val="both"/.test(p1.ppr), 'Primer párrafo justificado', 'Inicio > Párrafo > Justificar.'),
      S((() => { const l = attr((p1.ppr.match(/<w:spacing[^>]*>/) || [''])[0], 'line'); return l && near(l, 276, 8); })(), 'Interlineado 1.15', 'Inicio > Interlineado > 1.15.')]);
    const ph = 'rendimiento esperado: alto, medio y bajo';
    add(8, 2, 'Borrar formato y efecto de texto', [
      S((() => { const p = paras.find(q => q.n.includes(ph)); if (!p) return false; const runs = all(/<w:r[ >][\s\S]*?<\/w:r>/g, p.xml).filter(r => { const t = norm(textOf(r)); return t.length > 2 && ph.includes(t); }); return runs.length > 0 && runs.every(r => !/<w:(b|i|u|color|highlight|sz|shd)[ \/>]/.test((r.match(/<w:rPr>[\s\S]*?<\/w:rPr>/) || [''])[0])); })(), 'Frase "Rendimiento esperado…" sin formato', 'Selecciona la frase y usa Inicio > Borrar todo el formato.'),
      S((() => { const p = paras.find(q => q.style.toLowerCase() === 'title' || q.style.toLowerCase().startsWith('ttulo') && !/\d/.test(q.style) || q.n.startsWith('informe de inversiones')); return p && /w14:(shadow|glow|reflection|textOutline)|<w:shadow\b/.test(p.xml); })(), 'Efecto de texto (sombra) en el título', 'Inicio > Efectos de texto y tipografía > Sombra.')]);

    // ===== TAREA 3 =====
    add(9, 3, 'Reemplazar texto e insertar información adicional', [
      S(!/ahorrista/.test(allText), 'Ya no existe la palabra "ahorrista"', 'Inicio > Reemplazar.'),
      S((allText.match(/\binversor\b/g) || []).length >= 3, '"inversor" aparece 3 veces o más', 'Reemplazar todos: ahorrista → inversor.'),
      S(allText.includes(norm('La diversificación de la cartera reduce el impacto de un mal resultado individual')), 'Párrafo complementario insertado en el documento', 'Copia el texto de Texto_Complementario.docx al final del Tema 4.')]);
    const riesgos = ['riesgo de mercado', 'riesgo de credito', 'riesgo de liquidez'].map(t => paras.find(p => p.n === t));
    add(10, 3, 'Lista numerada de riesgos', [
      S(riesgos.every(p => p && numIdOf(p)), 'Los tres riesgos forman una lista', 'Inicio > Numeración.'),
      S(riesgos.every(p => { const i = p && numIdOf(p); const f = i && numFmt(i); return f && f !== 'bullet' && f !== 'none'; }), 'La lista es numerada (no viñetas)', 'Usa Numeración, no Viñetas.')]);
    add(11, 3, 'Símbolos ® y ©', [
      S(allText.includes('bloomberg®') || allText.includes('bloomberg ®'), '® después de Bloomberg', 'Insertar > Símbolo > Más símbolos.'),
      S(/©\s?moody/.test(allText), '© antes de Moody’s', 'Insertar > Símbolo > Más símbolos.')]);
    const cm = group(/^word\/comments\.xml$/), cms = all(/<w:comment\b[\s\S]*?<\/w:comment>/g, cm);
    add(12, 3, 'Comentario con tu nombre', [
      S(cms.length > 0, 'Existe al menos un comentario', 'Revisar > Nuevo comentario sobre la palabra "Invertir".'),
      S(cms.length && tokens(nombreDetectado).filter(w => norm(textOf(cm)).includes(w)).length >= 2, 'El comentario incluye tu nombre completo', 'Escribe tu nombre completo en el comentario.')]);

    // ===== TAREA 4 =====
    add(13, 4, 'Convertir texto en tabla', [
      S(tbl && tbl.rows >= 5 && tbl.cols >= 4, 'Tabla de al menos 5 filas y 4 columnas', 'Pega Datos_Portafolio.txt y usa Insertar > Tabla > Convertir texto en tabla.'),
      S(tbl && tbl.first.includes('instrumento') && tbl.t.includes('bonos del tesoro'), 'Contiene los datos del portafolio', 'Usa tabuladores como separador.')]);
    add(14, 4, 'Estilo de tabla y ampliación', [
      S((() => { if (!tbl || !tbl.style) return false; const s = styleById(tbl.style); const n = (s ? s.name : norm(tbl.style)).replace(/[^a-z0-9]/g, ''); return /gridtable4accent1|tabladecuadricula4enfasis1/.test(n) || /gridtable4.*accent1/.test(n); })(), 'Estilo "Tabla de cuadrícula 4 - Énfasis 1"', 'Diseño de tabla > Estilos de tabla.'),
      S(tbl && tbl.rows >= 6, 'Se agregó una fila', 'Disposición de tabla > Insertar abajo.'),
      S(tbl && tbl.cols >= 5, 'Se agregó una columna', 'Disposición de tabla > Insertar a la derecha.')]);
    add(15, 4, 'Imagen insertada y ajustada', [
      S(fotos.length > 0, 'Imagen Grafico_Financiero.jpg insertada', 'Insertar > Imágenes > Este dispositivo.'),
      S(fotos.some(f => near(f.cx, 3600000, 270000)), 'Ancho aproximado de 10 cm', 'Formato de imagen > Ancho: 10 cm.')]);
    add(16, 4, 'Ícono y forma', [
      S(names.some(n => /^word\/media\/.*\.svg$/i.test(n)) || /svgBlip/.test(docXml), 'Ícono de finanzas insertado', 'Insertar > Íconos.'),
      S(formas.some(f => /prst="horizontalScroll"/.test(f) && norm(textOf(f)).includes('finanzas')), 'Forma Pergamino: Horizontal con el texto FINANZAS', 'Insertar > Formas > Cintas y estrellas.')]);

    // ===== TAREA 5 =====
    add(17, 5, 'Encabezado y número de página', [
      S(headers.some(h => norm(textOf(h)).includes('inversiones inteligentes')), 'Encabezado "INVERSIONES INTELIGENTES"', 'Insertar > Encabezado.'),
      S(footers.some(f => /<w:instrText[^>]*>\s*PAGE\b|w:instr="\s*PAGE\b/.test(f)), 'Número de página en el pie', 'Insertar > Número de página > Final de página.')]);
    add(18, 5, 'Marca de agua CONFIDENCIAL', [
      S(headers.some(h => norm(h).includes('confidencial') && /PowerPlusWaterMarkObject|WordArt|_x0000_t136/i.test(h)), 'Marca de agua CONFIDENCIAL', 'Diseño > Marca de agua > Confidencial.')]);
    const fn = all(/<w:footnote\b(?![^>]*w:type=)[\s\S]*?<\/w:footnote>/g, F('word/footnotes.xml'));
    add(19, 5, 'Nota al pie', [
      S(/<w:footnoteReference\b/.test(docXml) && fn.length > 0, 'Existe una nota al pie', 'Referencias > Insertar nota al pie.'),
      S(fn.some(f => norm(textOf(f)).includes(norm('Es importante diferenciar entre el rendimiento nominal y el rendimiento real'))), 'Texto de la nota correcto', 'Copia el texto indicado sin cambios.')]);
    const pdf = ctx.pdf || {};
    add(20, 5, 'Cuadro de texto y PDF', [
      S(txbx.some(t => t.includes('a mayor liquidez, menor rendimiento esperado') && t.includes('principio fundamental')), 'Cuadro de texto con la cita', 'Insertar > Cuadro de texto > Cotización de filigrana.'),
      S(pdf.header && pdf.size > 3000 && /\.pdf$/i.test(pdf.name || ''), 'Archivo PDF cargado y válido', 'Archivo > Guardar como > PDF, y súbelo en la pestaña Entrega.')]);

    items.forEach(it => { it.obtenidos = it.subs.reduce((a, s) => a + (s.ok ? s.pts : 0), 0); });
    const total = items.reduce((a, i) => a + i.obtenidos, 0);
    return { items, total, max: TOTAL_PTS, aprobado: total >= APROBADO, minimo: APROBADO, nombreDetectado };
  }
  root.evaluarProyectoFinal = evaluarProyectoFinal;
  root.PUNTOS_FINAL = PUNTOS; root.NOMBRE_NIVEL_FINAL = NOMBRE_NIVEL; root.PUNTAJE_TOTAL_FINAL = TOTAL_PTS; root.PUNTAJE_APROBADO_FINAL = APROBADO;
  if (typeof module !== 'undefined') module.exports = { evaluarProyectoFinal, PUNTOS_FINAL: PUNTOS };
})(typeof window !== 'undefined' ? window : globalThis);
