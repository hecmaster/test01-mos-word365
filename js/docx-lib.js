/* Librería de lectura estructurada de .docx (solo regex; funciona en navegador y Node).
   DocxLib.parse(files, ctx) -> modelo con párrafos, secciones, estilos efectivos, listas, tablas, dibujos, encabezados, notas, comentarios. */
(function (root) {
  const norm = s => (s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[“”«»„]/g, '"').replace(/[‘’´`]/g, "'").replace(/\s+/g, ' ').trim().toLowerCase();
  const dec = s => (s || '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
  const all = (re, s) => [...(s || '').matchAll(re)].map(m => m[0]);
  const open = (x, name) => { const m = (x || '').match(new RegExp('<' + name + '(?=[\\s/>])[^>]*>')); return m ? m[0] : null; };
  const at = (tag, n) => { if (!tag) return null; const m = tag.match(new RegExp('(?:^|[\\s:])' + n.replace(':', ':') + '="([^"]*)"')); return m ? m[1] : null; };
  const wat = (tag, n) => at(tag, 'w:' + n);
  const textOf = x => dec([...(x || '').matchAll(/<(?:w:t|m:t)(?:\s[^>]*)?>([^<]*)<\/(?:w:t|m:t)>/g)].map(m => m[1]).join(''));
  const toggle = (rpr, name) => { const t = open(rpr, name); if (!t) return undefined; const v = wat(t, 'val'); return !(v === '0' || v === 'false' || v === 'off'); };
  const near = (v, t, tol) => v !== null && v !== undefined && !isNaN(v) && Math.abs(Number(v) - t) <= tol;

  /* ---- reparto de puntos por dificultad (suma exacta) ---- */
  function repartir(NIVELES, total) {
    const ids = Object.keys(NIVELES).map(Number), W = ids.reduce((a, i) => a + NIVELES[i].reduce((x, y) => x + y, 0), 0), flat = [];
    ids.forEach(i => NIVELES[i].forEach((w, k) => flat.push({ i, k, w, ex: total * w / W })));
    flat.forEach(f => { f.p = Math.floor(f.ex); }); const rest = total - flat.reduce((a, f) => a + f.p, 0);
    [...flat].sort((a, b) => (b.ex - b.p) - (a.ex - a.p) || a.i - b.i || a.k - b.k).slice(0, rest).forEach(f => f.p++);
    const out = {}; ids.forEach(i => { out[i] = { pts: [], niveles: NIVELES[i] }; }); flat.forEach(f => { out[f.i].pts[f.k] = f.p; });
    ids.forEach(i => { const o = out[i]; o.total = o.pts.reduce((a, b) => a + b, 0); const m = o.niveles.reduce((a, b) => a + b, 0) / o.niveles.length; o.nivel = m >= 2.5 ? 3 : m >= 1.5 ? 2 : 1; });
    return out;
  }
  const NOMBRE_NIVEL = { 1: 'Básico', 2: 'Intermedio', 3: 'Avanzado' };
  function crearEvaluador(NIVELES, total, aprobado) {
    const PUNTOS = repartir(NIVELES, total), items = [];
    const add = (id, area, titulo, subs) => { const P = PUNTOS[id]; if (!P || P.pts.length !== subs.length) throw new Error('Niveles del criterio ' + id + ' no coinciden'); subs.forEach((s, k) => { s.pts = P.pts[k]; s.nivel = P.niveles[k]; s.nivelNombre = NOMBRE_NIVEL[P.niveles[k]]; }); items.push({ id, tarea: area, titulo, pts: P.total, nivel: P.nivel, nivelNombre: NOMBRE_NIVEL[P.nivel], subs }); };
    const cerrar = nombreDetectado => { items.forEach(it => { it.obtenidos = it.subs.reduce((a, s) => a + (s.ok ? s.pts : 0), 0); }); const t = items.reduce((a, i) => a + i.obtenidos, 0); return { items, total: t, max: total, aprobado: t >= aprobado, minimo: aprobado, nombreDetectado }; };
    return { PUNTOS, add, cerrar, S: (ok, label, pista) => ({ ok: !!ok, label, pista }) };
  }

  /* ---- exploración con profundidad (párrafos y tablas anidados en cuadros de texto) ---- */
  function scan(xml, openRe, closeStr) {
    const out = []; const re = new RegExp(openRe.source + '|' + closeStr.replace(/[/]/g, '\\/'), 'g'); let m, depth = 0, start = -1;
    while ((m = re.exec(xml))) {
      if (m[0] === closeStr) { depth--; if (depth === 0) out.push({ start, end: m.index + closeStr.length, xml: xml.slice(start, m.index + closeStr.length) }); }
      else { const c = xml.indexOf('>', m.index); if (xml[c - 1] === '/') continue; if (depth === 0) start = m.index; depth++; }
    }
    return out;
  }

  function parse(files, ctx) {
    ctx = ctx || {}; const F = n => files[n] || '', names = Object.keys(files);
    const M = { files, names, ctx };
    const raw = F('word/document.xml').replace(/<mc:Fallback>[\s\S]*?<\/mc:Fallback>/g, '');
    M.raw = raw;
    // ---- relaciones, tema, estilos ----
    M.rels = {}; for (const t of all(/<Relationship\s[^>]*>/g, F('word/_rels/document.xml.rels'))) { const id = at(t, 'Id'), tg = at(t, 'Target'), ty = at(t, 'Type'); if (id) M.rels[id] = { target: tg, type: ty }; }
    const themeXml = F('word/theme/theme1.xml'); M.themeName = (themeXml.match(/<a:theme\b[^>]*\sname="([^"]*)"/) || [, ''])[1];
    M.themeColors = {}; for (const n of ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink']) { const m = themeXml.match(new RegExp('<a:' + n + '>\\s*<a:(?:srgbClr val|sysClr val="[^"]*" lastClr)="([0-9A-Fa-f]{6})"')); if (m) M.themeColors[n] = m[1].toUpperCase(); }
    M.themeFonts = { major: (themeXml.match(/<a:majorFont>\s*<a:latin typeface="([^"]*)"/) || [, ''])[1], minor: (themeXml.match(/<a:minorFont>\s*<a:latin typeface="([^"]*)"/) || [, ''])[1] };
    const stylesXml = F('word/styles.xml'); M.stylesXml = stylesXml;
    M.styles = all(/<w:style\b[^>]*>[\s\S]*?<\/w:style>/g, stylesXml).map(x => ({ x, id: wat(open(x, 'w:style'), 'styleId'), type: wat(open(x, 'w:style'), 'type'), name: norm(wat(open(x, 'w:name'), 'val')), basedOn: wat(open(x, 'w:basedOn'), 'val'), nameRaw: wat(open(x, 'w:name'), 'val') }));
    M.styleById = i => M.styles.find(s => s.id === i);
    M.styleByName = n => M.styles.find(s => s.name === norm(n));
    const dd = (stylesXml.match(/<w:docDefaults>[\s\S]*?<\/w:docDefaults>/) || [''])[0];
    M.defRpr = (dd.match(/<w:rPrDefault>[\s\S]*?<\/w:rPrDefault>/) || [''])[0]; M.defPpr = (dd.match(/<w:pPrDefault>[\s\S]*?<\/w:pPrDefault>/) || [''])[0];
    const chain = id => { const out = []; let s = M.styleById(id), g = 0; while (s && g++ < 12) { out.push(s); s = s.basedOn ? M.styleById(s.basedOn) : null; } return out; };
    M.chain = chain;
    const sRpr = s => ((s.x.replace(/<w:pPr>[\s\S]*?<\/w:pPr>/, '').match(/<w:rPr>[\s\S]*?<\/w:rPr>/)) || [''])[0];
    const sPpr = s => (s.x.match(/<w:pPr>[\s\S]*?<\/w:pPr>/) || [''])[0];
    // ---- propiedades del documento ----
    const core = F('docProps/core.xml'); const cget = n => dec((core.match(new RegExp('<' + n + '>([^<]*)</' + n + '>')) || [, ''])[1]).trim();
    M.core = { creator: cget('dc:creator'), description: cget('dc:description'), title: cget('dc:title'), lastModifiedBy: cget('cp:lastModifiedBy') };
    const settings = F('word/settings.xml'); M.evenAndOdd = /<w:evenAndOddHeaders\b(?![^>]*w:val="(?:0|false)")/.test(settings); M.displayBg = /<w:displayBackgroundShape\b/.test(settings);
    M.background = open(raw, 'w:background');

    // ---- párrafos de nivel superior ----
    const txbxRanges = all(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g, raw).length ? [] : [];
    for (const m of raw.matchAll(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g)) txbxRanges.push([m.index, m.index + m[0].length]);
    const inTx = pos => txbxRanges.some(r => pos > r[0] && pos < r[1]);
    const buildPara = (p, idx) => {
      const xml = p.xml.replace(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g, '');
      const ppr = (xml.match(/<w:pPr>[\s\S]*?<\/w:pPr>/) || [''])[0], pprNoRpr = ppr.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/g, '');
      const runs = all(/<(w:r|m:r)(?=[\s>])[\s\S]*?<\/\1>/g, xml).map(r => ({ xml: r, rpr: (r.match(/<w:rPr>[\s\S]*?<\/w:rPr>/) || [''])[0], text: textOf(r) }));
      return { idx, pos: p.start, end: p.end, raw: p.xml, xml, ppr, pprNoRpr, style: wat(open(ppr, 'w:pStyle'), 'val') || '', text: textOf(xml), n: norm(textOf(xml)), runs,
        sect: (ppr.match(/<w:sectPr[ >][\s\S]*?<\/w:sectPr>/) || [null])[0], numPr: (ppr.match(/<w:numPr>[\s\S]*?<\/w:numPr>/) || [''])[0] };
    };
    M.paras = scan(raw, /<w:p[ >\/]/, '</w:p>').map((p, idx) => buildPara(p, idx));
    M.parasOf = xml => scan(xml, /<w:p[ >\/]/, '</w:p>').map((p, i) => buildPara(p, i));
    M.paras.forEach(p => { p.styleName = (M.styleById(p.style) || {}).name || (p.style ? norm(p.style) : 'normal'); });
    // ---- secciones ----
    M.sections = []; let cur = [];
    for (const p of M.paras) { cur.push(p); p.sec = M.sections.length; if (p.sect) { M.sections.push({ paras: cur, sect: p.sect }); cur = []; } }
    const fin = raw.match(/(<w:sectPr[ >](?:(?!<w:sectPr)[\s\S])*?<\/w:sectPr>)\s*<\/w:body>/);
    cur.forEach(p => { p.sec = M.sections.length; }); M.sections.push({ paras: cur, sect: fin ? fin[1] : '' });
    M.sections.forEach((s, i) => {
      const pg = open(s.sect, 'w:pgSz'), mg = open(s.sect, 'w:pgMar'), cols = open(s.sect, 'w:cols'), ty = open(s.sect, 'w:type');
      s.w = Number(wat(pg, 'w')) || 0; s.h = Number(wat(pg, 'h')) || 0; s.landscape = wat(pg, 'orient') === 'landscape' || (s.w > s.h && s.w > 0);
      s.mar = { top: wat(mg, 'top'), bottom: wat(mg, 'bottom'), left: wat(mg, 'left'), right: wat(mg, 'right') };
      s.cols = Number(wat(cols, 'num')) || 1; s.type = wat(ty, 'val') || 'nextPage'; s.titlePg = /<w:titlePg\b(?![^>]*w:val="(?:0|false)")/.test(s.sect); s.index = i;
      s.refs = {}; for (const t of all(/<w:(?:header|footer)Reference\s[^>]*>/g, s.sect)) { const kind = /headerReference/.test(t) ? 'header' : 'footer', ty2 = wat(t, 'type'), rid = at(t, 'r:id'); const tg = M.rels[rid] && M.rels[rid].target; if (tg) s.refs[kind + ':' + ty2] = 'word/' + tg.replace(/^\//, '').replace(/^word\//, ''); }
    });
    M.sections.forEach((s, i) => { ['header', 'footer'].forEach(k => ['default', 'first', 'even'].forEach(t => { const key = k + ':' + t; if (!s.refs[key] && i > 0 && M.sections[i - 1].refs[key]) s.refs[key] = M.sections[i - 1].refs[key]; })); });
    M.hf = (sec, kind, type) => { const f = M.sections[sec] && M.sections[sec].refs[kind + ':' + type]; return f && files[f] ? { name: f, xml: files[f], text: textOf(files[f]) } : null; };
    M.allHeaders = names.filter(n => /^word\/header\d*\.xml$/.test(n)).map(n => ({ name: n, xml: files[n], text: textOf(files[n]) }));
    M.allFooters = names.filter(n => /^word\/footer\d*\.xml$/.test(n)).map(n => ({ name: n, xml: files[n], text: textOf(files[n]) }));

    // ---- numeración ----
    const numXml = F('word/numbering.xml'); M.nums = {}; M.abs = {}; M.picBullets = [...numXml.matchAll(/<w:numPicBullet\s[^>]*w:numPicBulletId="(\d+)"/g)].map(m => m[1]);
    for (const a of all(/<w:abstractNum\b[\s\S]*?<\/w:abstractNum>/g, numXml)) { const id = wat(open(a, 'w:abstractNum'), 'abstractNumId'), lv = {}; for (const l of all(/<w:lvl\b[\s\S]*?<\/w:lvl>/g, a)) lv[wat(open(l, 'w:lvl'), 'ilvl')] = { fmt: wat(open(l, 'w:numFmt'), 'val'), text: wat(open(l, 'w:lvlText'), 'val') || '', pic: wat(open(l, 'w:lvlPicBulletId'), 'val') }; M.abs[id] = { multi: wat(open(a, 'w:multiLevelType'), 'val'), lvls: lv }; }
    for (const n of all(/<w:num\b[\s\S]*?<\/w:num>/g, numXml)) M.nums[wat(open(n, 'w:num'), 'numId')] = wat(open(n, 'w:abstractNumId'), 'val');
    M.numInfo = p => {
      let np = p.numPr, id = wat(open(np, 'w:numId'), 'val'), lv = wat(open(np, 'w:ilvl'), 'val');
      if ((id === null || id === undefined) && p.style) for (const s of chain(p.style)) { const sp = sPpr(s), npp = (sp.match(/<w:numPr>[\s\S]*?<\/w:numPr>/) || [''])[0]; if (npp) { id = wat(open(npp, 'w:numId'), 'val'); if (lv === null) lv = wat(open(npp, 'w:ilvl'), 'val'); break; } }
      if (id === null || id === undefined || id === '0') return null; lv = lv === null ? '0' : lv; const a = M.abs[M.nums[id]]; const l = a && a.lvls[lv] || {};
      return { numId: id, ilvl: Number(lv), multi: a ? a.multi : null, fmt: l.fmt || null, text: l.text || '', pic: l.pic !== undefined && l.pic !== null ? l.pic : null, abs: a };
    };

    // ---- formato efectivo ----
    const colorOf = rpr => { const t = open(rpr, 'w:color'); if (!t) return undefined; const v = wat(t, 'val'), th = wat(t, 'themeColor'); if (v && v !== 'auto' && /^[0-9A-Fa-f]{6}$/.test(v)) return v.toUpperCase(); if (th) { const k = { text1: 'dk1', background1: 'lt1', text2: 'dk2', background2: 'lt2' }[th] || th; return M.themeColors[k] || null; } return null; };
    const fontOf = rpr => { const t = open(rpr, 'w:rFonts'); if (!t) return undefined; const a = wat(t, 'ascii') || wat(t, 'hAnsi'); if (a) return a; const th = wat(t, 'asciiTheme') || wat(t, 'hAnsiTheme'); if (th) return /major/i.test(th) ? M.themeFonts.major : M.themeFonts.minor; return undefined; };
    const parseR = rpr => ({ font: fontOf(rpr), sz: (v => v === null ? undefined : Number(v))(wat(open(rpr, 'w:sz'), 'val')), b: toggle(rpr, 'w:b'), i: toggle(rpr, 'w:i'), caps: toggle(rpr, 'w:caps'), smallCaps: toggle(rpr, 'w:smallCaps'), u: (t => t ? wat(t, 'val') !== 'none' : undefined)(open(rpr, 'w:u')), color: colorOf(rpr), highlight: (t => t ? wat(t, 'val') : undefined)(open(rpr, 'w:highlight')), shd: (t => t ? wat(t, 'fill') : undefined)(open(rpr, 'w:shd')), shadow: /<w14:shadow\b|<w:shadow\b/.test(rpr || '') || undefined });
    M.efRun = (p, run) => {
      const layers = [parseR(run ? run.rpr : '')]; const rs = wat(open(run ? run.rpr : '', 'w:rStyle'), 'val'); if (rs) for (const s of chain(rs)) layers.push(parseR(sRpr(s)));
      if (p && p.style) for (const s of chain(p.style)) layers.push(parseR(sRpr(s)));
      layers.push(parseR(M.defRpr)); const out = {}; for (const k of ['font', 'sz', 'b', 'i', 'caps', 'smallCaps', 'u', 'color', 'highlight', 'shd', 'shadow']) { for (const L of layers) if (L[k] !== undefined) { out[k] = L[k]; break; } }
      out.b = !!out.b; out.i = !!out.i; out.caps = !!out.caps; out.u = !!out.u; if (out.sz === undefined) out.sz = 22; return out;
    };
    const parseP = ppr => { const sp = open(ppr, 'w:spacing'), ind = open(ppr, 'w:ind'); const n = v => v === null ? undefined : Number(v); return { jc: wat(open(ppr, 'w:jc'), 'val') || undefined, before: n(wat(sp, 'before')), after: n(wat(sp, 'after')), line: n(wat(sp, 'line')), lineRule: wat(sp, 'lineRule') || undefined, firstLine: n(wat(ind, 'firstLine')), hanging: n(wat(ind, 'hanging')), left: n(wat(ind, 'left') || wat(ind, 'start')) }; };
    M.efPara = p => { const layers = [parseP(p.pprNoRpr)]; if (p.style) for (const s of chain(p.style)) layers.push(parseP(sPpr(s))); layers.push(parseP(M.defPpr)); const o = {}; for (const k of ['jc', 'before', 'after', 'line', 'lineRule', 'firstLine', 'hanging', 'left']) { for (const L of layers) if (L[k] !== undefined) { o[k] = L[k]; break; } } if (o.hanging) o.firstLine = o.firstLine || -o.hanging; return o; };
    M.isHeadingStyle = (p, n) => { const s = M.styleByName('heading ' + n); return !!(s && p.style === s.id); };
    M.styleRpr = id => { const s = M.styleById(id); return s ? s.x : ''; };
    M.styleEf = id => { const fake = { style: id, pprNoRpr: '' }; return { run: M.efRun(fake, null), para: M.efPara(fake) }; };

    // ---- tablas ----
    const rawNoTx = raw.replace(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g, m => ' '.repeat(m.length));
    M.tables = scan(rawNoTx, /<w:tbl>/, '</w:tbl>').map(t => {
      const x = raw.slice(t.start, t.end), rows = all(/<w:tr[ >][\s\S]*?<\/w:tr>/g, x).map(r => all(/<w:tc>[\s\S]*?<\/w:tc>/g, r).map(c => ({ xml: c, span: Number(wat(open(c, 'w:gridSpan'), 'val')) || 1, vmerge: !!open(c, 'w:vMerge'), text: textOf(c), jc: wat(open(c, 'w:jc'), 'val') })));
      const sid = wat(open(x, 'w:tblStyle'), 'val'), st = sid && M.styleById(sid);
      return { pos: t.start, end: t.end, xml: x, rows, nRows: rows.length, nCols: Math.max(0, ...rows.map(r => r.reduce((a, c) => a + c.span, 0))), style: sid, styleName: st ? st.name : (sid ? norm(sid) : ''), text: norm(textOf(x)), firstRow: norm(rows[0] ? rows[0].map(c => c.text).join(' ') : '') };
    });
    // ---- dibujos, formas y cuadros de texto ----
    M.drawings = [];
    const mkDraw = (x, pos, nested) => {
      const own = x.replace(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g, ''), kind = /^<w:drawing>\s*<wp:anchor/.test(x) ? 'anchor' : 'inline', ext = open(own, 'wp:extent'), docPr = open(own, 'wp:docPr');
      const emb = at(open(own, 'a:blip'), 'r:embed'), tg = emb && M.rels[emb] ? M.rels[emb].target : '', svgRid = (own.match(/<asvg:svgBlip[^>]*r:embed="([^"]+)"/) || [])[1], svgT = svgRid && M.rels[svgRid] ? M.rels[svgRid].target : '';
      const posH = own.match(/<wp:positionH relativeFrom="(\w+)">\s*(?:<wp:align>(\w+)<\/wp:align>|<wp:posOffset>(-?\d+)<\/wp:posOffset>)/), posV = own.match(/<wp:positionV relativeFrom="(\w+)">\s*(?:<wp:align>(\w+)<\/wp:align>|<wp:posOffset>(-?\d+)<\/wp:posOffset>)/);
      const host = M.paras.find(p => pos >= p.pos && pos < p.end);
      return { xml: x, pos, kind, nested, cx: Number(at(ext, 'cx')) || 0, cy: Number(at(ext, 'cy')) || 0, name: at(docPr, 'name') || '', descr: at(docPr, 'descr') || '', embed: emb, target: tg, ext: (tg.match(/\.(\w+)$/) || [, ''])[1].toLowerCase(),
        svg: !!svgT || /\.svg$/i.test(tg), isPic: /<pic:pic\b/.test(own), isShape: /<wps:wsp\b/.test(own), prst: all(/prst="(\w+)"/g, own).map(t => t.slice(6, -1)).filter(t => t !== 'textNoShape'),
        sat: Number((own.match(/<a14:saturation sat="(\d+)"/) || [])[1]) || null, hAlign: posH ? posH[2] || null : null, hOff: posH && posH[3] ? Number(posH[3]) : null, hFrom: posH ? posH[1] : null, vAlign: posV ? posV[2] || null : null, vOff: posV && posV[3] ? Number(posV[3]) : null, vFrom: posV ? posV[1] : null,
        txbx: all(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g, x).map(t => ({ text: textOf(t), n: norm(textOf(t)), hasPic: /<w:drawing>/.test(t) })), para: host ? host.idx : -1, hostPos: host ? pos - host.pos : 0 };
    };
    const drawingsIn = (xml, base, nested) => { const out = []; for (const d of scan(xml, /<w:drawing>/, '</w:drawing>')) { out.push(mkDraw(d.xml, base + d.start, nested)); out.push(...drawingsIn(d.xml.slice(11, -12), base + d.start + 11, true)); } return out; };
    M.drawings = drawingsIn(raw, 0, false);
    M.inTable = p => M.tables.some(t => p.pos >= t.pos && p.pos < t.end);
    M.shapesOf = xml => scan(xml, /<w:drawing>/, '</w:drawing>').map(d => d.xml).map(x => ({ xml: x, prst: all(/prst="(\w+)"/g, x).map(s => s.slice(6, -1)).filter(s => s !== 'textNoShape'), txbx: all(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g, x).map(t => norm(textOf(t))), fill: /<a:solidFill>/.test((x.match(/<wps:spPr[\s\S]*?<\/wps:spPr>/) || [''])[0]), hasPic: /<pic:pic\b/.test(x), pageField: /<w:instrText[^>]*>\s*PAGE\b|w:instr="\s*PAGE\b/.test(x) }));
    M.legacyTx = all(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g, raw).map(t => norm(textOf(t)));
    // ---- notas, comentarios ----
    M.footnotes = all(/<w:footnote\b[\s\S]*?<\/w:footnote>/g, F('word/footnotes.xml')).filter(f => !/w:type="(separator|continuationSeparator|continuationNotice)"/.test(f)).map(f => ({ id: wat(open(f, 'w:footnote'), 'id'), text: textOf(f), n: norm(textOf(f)) }));
    M.fnRefs = []; for (const p of M.paras) for (const m of p.xml.matchAll(/<w:footnoteReference\b[^>]*w:id="(\d+)"/g)) M.fnRefs.push({ id: m[1], para: p.idx });
    M.comments = all(/<w:comment\b[\s\S]*?<\/w:comment>/g, F('word/comments.xml')).map(c => { const id = wat(open(c, 'w:comment'), 'id'); let anchor = '', apara = -1; for (const p of M.paras) { const m = p.xml.match(new RegExp('<w:commentRangeStart w:id="' + id + '"/>([\\s\\S]*?)<w:commentRangeEnd w:id="' + id + '"/>')); if (m) { anchor = textOf(m[1]); apara = p.idx; break; } } return { id, author: wat(open(c, 'w:comment'), 'author') || '', text: textOf(c), n: norm(textOf(c)), anchor, nAnchor: norm(anchor), para: apara }; });
    M.mediaNames = names.filter(n => /^word\/media\//.test(n));
    M.text = norm(M.paras.map(p => p.text).join(' '));
    return M;
  }

  /* ---- utilidades de color ---- */
  const hexToRgb = h => [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
  function hue(h) { const [r, g, b] = hexToRgb(h).map(v => v / 255), mx = Math.max(r, g, b), mn = Math.min(r, g, b), d = mx - mn; if (!d) return -1; let x = mx === r ? ((g - b) / d) % 6 : mx === g ? (b - r) / d + 2 : (r - g) / d + 4; x *= 60; return x < 0 ? x + 360 : x; }
  const isBlue = h => { if (!h) return false; const hh = hue(h), [r, g, b] = hexToRgb(h); return hh >= 185 && hh <= 255 && b >= 90 && b > r + 20; };
  const isGray = h => { if (!h) return false; const [r, g, b] = hexToRgb(h); return Math.max(r, g, b) - Math.min(r, g, b) <= 18; };

  const tokens = x => norm(x).split(' ').filter(w => w.length > 2);
  const nombreValido = n => { n = (n || '').trim(); return n.split(/\s+/).filter(w => w.length > 1).length >= 2 && !/escriba|nombre completo|estudiante|usuario|autor|apellido|nombre y/i.test(n) && !/\d{3,}/.test(n); };
  const coincideNombre = (reg, cand) => { const t = tokens(reg), c = norm(cand); return t.length >= 2 && t.filter(w => c.includes(w)).length >= Math.max(2, Math.ceil(t.length * 0.6)); };
  const API = { tokens, nombreValido, coincideNombre, norm, dec, all, open, at, wat, textOf, near, parse, repartir, crearEvaluador, NOMBRE_NIVEL, isBlue, isGray, hexToRgb, hue };
  root.DocxLib = API; if (typeof module !== 'undefined') module.exports = API;
})(typeof window !== 'undefined' ? window : globalThis);
