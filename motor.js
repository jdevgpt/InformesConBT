/* Motor de generación del INFORME BT: rellena la plantilla Word y crea el PDF.
   Requiere JSZip y PDFLib (cargados desde CDN en index.html). */
(function (global) {
  'use strict';
  const NS = {
    w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
    pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
    ct: 'http://schemas.openxmlformats.org/package/2006/content-types'
  };
  const W = NS.w;
  const TXT_SIN = 'NO SE HAN DETECTADO DEFECTOS EN LA INSTALACIÓN.';

  // Orden de los hijos según el esquema de Word (si no se respeta, Word puede dar el archivo por dañado)
  const PPR_ORDER = ['pStyle','keepNext','keepLines','pageBreakBefore','framePr','widowControl','numPr','suppressLineNumbers','pBdr','shd','tabs','suppressAutoHyphens','kinsoku','wordWrap','overflowPunct','topLinePunct','autoSpaceDE','autoSpaceDN','bidi','adjustRightInd','snapToGrid','spacing','ind','contextualSpacing','mirrorIndents','suppressOverlap','jc','textDirection','textAlignment','textboxTightWrap','outlineLvl','divId','cnfStyle','rPr','sectPr','pPrChange'];
  const TRPR_ORDER = ['cnfStyle','divId','gridBefore','gridAfter','wBefore','wAfter','cantSplit','trHeight','tblHeader','tblCellSpacing','jc','hidden','ins','del','trPrChange'];

  const kids = (el, name) => Array.from(el.childNodes).filter(n => n.nodeType === 1 && n.localName === name && n.namespaceURI === W);
  const kid = (el, name) => kids(el, name)[0] || null;
  const all = (el, name, ns = W) => Array.from(el.getElementsByTagNameNS(ns, name));
  const ptext = el => all(el, 't').map(t => t.textContent).join('');
  const wAttr = (el, name) => el ? el.getAttributeNS(W, name) || el.getAttribute('w:' + name) : null;

  function insertOrdered(parent, el, order) {
    const idx = order.indexOf(el.localName);
    for (const c of Array.from(parent.childNodes)) {
      if (c.nodeType === 1 && order.indexOf(c.localName) > idx) { parent.insertBefore(el, c); return el; }
    }
    parent.appendChild(el); return el;
  }
  function getPPr(p) {
    let pPr = kid(p, 'pPr');
    if (!pPr) { pPr = p.ownerDocument.createElementNS(W, 'w:pPr'); p.insertBefore(pPr, p.firstChild); }
    return pPr;
  }
  function setFlag(p, name) {
    const pPr = getPPr(p);
    if (kid(pPr, name)) return;
    insertOrdered(pPr, p.ownerDocument.createElementNS(W, 'w:' + name), PPR_ORDER);
  }
  function setRowFlag(tr, name) {
    let trPr = kid(tr, 'trPr');
    if (!trPr) {
      trPr = tr.ownerDocument.createElementNS(W, 'w:trPr');
      const after = kid(tr, 'tblPrEx');
      tr.insertBefore(trPr, after ? after.nextSibling : tr.firstChild);
    }
    if (!kid(trPr, name)) insertOrdered(trPr, tr.ownerDocument.createElementNS(W, 'w:' + name), TRPR_ORDER);
  }

  // Sustituye [CAMPO] aunque Word haya partido el texto en varios fragmentos
  function replaceInPara(p, map) {
    const ts = all(p, 't');
    if (!ts.length) return;
    const full = ts.map(t => t.textContent).join('');
    const re = /\[([A-ZÁÉÍÓÚÑ_ ]+)\]/g; const hits = []; let m;
    while ((m = re.exec(full))) {
      const key = m[1].trim();
      if (Object.prototype.hasOwnProperty.call(map, key)) hits.push({ s: m.index, e: m.index + m[0].length, v: String(map[key]) });
    }
    for (const h of hits.reverse()) {
      let pos = 0;
      for (const t of ts) {
        const txt = t.textContent, a = pos, b = pos + txt.length; pos = b;
        if (b <= h.s || a >= h.e || txt.length === 0) continue;
        const ls = Math.max(h.s - a, 0), le = Math.min(h.e - a, txt.length);
        t.textContent = (a <= h.s && h.s < b) ? txt.slice(0, ls) + h.v + txt.slice(le) : txt.slice(0, ls) + txt.slice(le);
        t.setAttribute('xml:space', 'preserve');
      }
    }
  }

  function bodyOf(doc) { return all(doc, 'body')[0]; }
  function isEmptyPara(p) {
    return p.localName === 'p' && ptext(p).trim() === '' && !all(p, 'drawing').length && !all(p, 'pict').length && !all(p, 'sectPr').length && !hasPageBreak(p);
  }
  function hasPageBreak(p) {
    return all(p, 'br').some(b => wAttr(b, 'type') === 'page') || all(p, 'pageBreakBefore').length > 0 || all(p, 'lastRenderedPageBreak').length > 0 && false;
  }

  // ---------- imágenes ----------
  let imgCounter = 0;
  function emuSize(wPx, hPx, maxWcm, maxHcm) {
    const cm = 360000; let w = maxWcm * cm, h = w * hPx / wPx;
    if (h > maxHcm * cm) { h = maxHcm * cm; w = h * wPx / hPx; }
    return { cx: Math.round(w), cy: Math.round(h) };
  }
  function fitBox(wPx, hPx, boxCx, boxCy) {
    const k = Math.min(boxCx / wPx, boxCy / hPx);
    return { cx: Math.round(wPx * k), cy: Math.round(hPx * k) };
  }
  // Cambia la imagen de un dibujo ya existente en la plantilla, ajustada a la caja indicada (EMU)
  async function swapImage(p, img, zip, relsDoc, ctDoc, base, boxCx, boxCy) {
    const dr = all(p, 'drawing')[0]; if (!dr) return false;
    const blip = all(dr, 'blip', NS.a)[0]; if (!blip) return false;
    const ext0 = all(dr, 'extent', NS.wp)[0];
    const bx = boxCx || +ext0.getAttribute('cx'), by = boxCy || +ext0.getAttribute('cy');
    const rid = await addMedia(zip, relsDoc, ctDoc, img, base);
    blip.setAttributeNS(NS.r, 'r:embed', rid);
    const sz = fitBox(img.w, img.h, bx, by);
    all(dr, 'extent', NS.wp).forEach(e => { e.setAttribute('cx', sz.cx); e.setAttribute('cy', sz.cy); });
    all(dr, 'ext', NS.a).forEach(e => { if (e.hasAttribute('cx')) { e.setAttribute('cx', sz.cx); e.setAttribute('cy', sz.cy); } });
    return true;
  }
  function removeDrawings(p) {
    for (const dr of all(p, 'drawing')) { const r = dr.parentNode; r.parentNode.removeChild(r); }
    // quita el salto de línea que quedaba antes de la imagen
    const runs = kids(p, 'r'); const last = runs[runs.length - 1];
    if (last && kids(last, 'br').length && !kids(last, 't').length) p.removeChild(last);
  }

  function drawingRun(doc, rid, cx, cy, name) {
    imgCounter++;
    const id = 9000 + imgCounter;
    const xml = `<w:r xmlns:w="${NS.w}" xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" xmlns:pic="${NS.pic}" xmlns:r="${NS.r}"><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="${name} ${id}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="${NS.pic}"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="${name}.img"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
    const frag = new DOMParser().parseFromString(xml, 'application/xml').documentElement;
    return doc.importNode(frag, true);
  }
  async function addMedia(zip, relsDoc, ctDoc, img, base) {
    const ext = img.type === 'image/png' ? 'png' : 'jpeg';
    const fname = `${base}_${Date.now()}_${imgCounter}.${ext}`;
    zip.file('word/media/' + fname, img.bytes);
    const root = relsDoc.documentElement;
    const ids = Array.from(root.getElementsByTagNameNS(NS.rel, 'Relationship')).map(r => r.getAttribute('Id'));
    let n = 1000; while (ids.includes('rIdIBT' + n)) n++;
    const rid = 'rIdIBT' + n;
    const rel = relsDoc.createElementNS(NS.rel, 'Relationship');
    rel.setAttribute('Id', rid);
    rel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image');
    rel.setAttribute('Target', 'media/' + fname);
    root.appendChild(rel);
    const ctRoot = ctDoc.documentElement;
    const has = Array.from(ctRoot.getElementsByTagNameNS(NS.ct, 'Default')).some(d => (d.getAttribute('Extension') || '').toLowerCase() === ext);
    if (!has) {
      const d = ctDoc.createElementNS(NS.ct, 'Default');
      d.setAttribute('Extension', ext); d.setAttribute('ContentType', 'image/' + ext);
      ctRoot.insertBefore(d, ctRoot.firstChild);
    }
    return rid;
  }

  function pageGeom(doc) {
    const sect = all(doc, 'sectPr').pop();
    const pgSz = sect && kid(sect, 'pgSz'), pgMar = sect && kid(sect, 'pgMar');
    const num = (el, a, d) => { const v = el && wAttr(el, a); return v ? parseInt(v, 10) : d; };
    const g = { w: num(pgSz, 'w', 11906), h: num(pgSz, 'h', 16838), top: num(pgMar, 'top', 1417), bottom: num(pgMar, 'bottom', 1417), left: num(pgMar, 'left', 1701), right: num(pgMar, 'right', 1701), header: num(pgMar, 'header', 709), footer: num(pgMar, 'footer', 709) };
    g.textWcm = (g.w - g.left - g.right) / 567; g.textHcm = (g.h - g.top - g.bottom) / 567;
    return g;
  }

  function baseRunInfo(stylesDoc, refPara) {
    // tamaño y fuente del texto normal, para la frase de "sin defectos" (+1 punto)
    let sz = null, fonts = null;
    if (refPara) {
      for (const r of all(refPara, 'r')) {
        const rPr = kid(r, 'rPr'); if (!rPr) continue;
        if (!sz && kid(rPr, 'sz')) sz = parseInt(wAttr(kid(rPr, 'sz'), 'val'), 10);
        if (!fonts && kid(rPr, 'rFonts')) fonts = kid(rPr, 'rFonts');
      }
    }
    if (!sz && stylesDoc) {
      const st = resolveStyles(stylesDoc);
      const pStyle = refPara && kid(getPPr(refPara), 'pStyle');
      const sid = pStyle ? wAttr(pStyle, 'val') : st.defaultParaStyle;
      sz = st.runSize(sid);
    }
    return { sz: sz || 22, fonts };
  }

  function sinDefectosPara(doc, info, refPara) {
    const p = doc.createElementNS(W, 'w:p');
    const pPr = doc.createElementNS(W, 'w:pPr');
    const refStyle = refPara && kid(refPara, 'pPr') && kid(kid(refPara, 'pPr'), 'pStyle');
    if (refStyle) pPr.appendChild(refStyle.cloneNode(true)); // misma fuente que el texto del informe
    const sp = doc.createElementNS(W, 'w:spacing'); sp.setAttributeNS(W, 'w:before', '240'); sp.setAttributeNS(W, 'w:beforeAutospacing', '0'); sp.setAttributeNS(W, 'w:after', '240'); sp.setAttributeNS(W, 'w:afterAutospacing', '0');
    const ind = doc.createElementNS(W, 'w:ind'); ind.setAttributeNS(W, 'w:left', '0'); ind.setAttributeNS(W, 'w:right', '0'); ind.setAttributeNS(W, 'w:firstLine', '0');
    const jc = doc.createElementNS(W, 'w:jc'); jc.setAttributeNS(W, 'w:val', 'center');
    pPr.appendChild(sp); pPr.appendChild(ind); pPr.appendChild(jc); p.appendChild(pPr);
    const r = doc.createElementNS(W, 'w:r'); const rPr = doc.createElementNS(W, 'w:rPr');
    if (info.fonts) rPr.appendChild(info.fonts.cloneNode(true));
    rPr.appendChild(doc.createElementNS(W, 'w:b')); rPr.appendChild(doc.createElementNS(W, 'w:bCs'));
    const sz = doc.createElementNS(W, 'w:sz'); sz.setAttributeNS(W, 'w:val', String(info.sz + 2));
    const szCs = doc.createElementNS(W, 'w:szCs'); szCs.setAttributeNS(W, 'w:val', String(info.sz + 2));
    rPr.appendChild(sz); rPr.appendChild(szCs); r.appendChild(rPr);
    const t = doc.createElementNS(W, 'w:t'); t.textContent = TXT_SIN; r.appendChild(t); p.appendChild(r);
    return p;
  }

  /**
   * data = { nombre, fecha, centro, direccion, poblacion, sinDefectos, defectos:[{apoyo,linea,cod,des,obs}],
   *          firma:{bytes,type,w,h}|null, foto:{bytes,type,w,h}|null }
   */
  async function generarDocx(templateBytes, data) {
    const zip = await JSZip.loadAsync(templateBytes);
    const parser = new DOMParser(), ser = new XMLSerializer();
    const doc = parser.parseFromString(await zip.file('word/document.xml').async('string'), 'application/xml');
    const relsPath = 'word/_rels/document.xml.rels';
    const relsDoc = parser.parseFromString(await zip.file(relsPath).async('string'), 'application/xml');
    const ctDoc = parser.parseFromString(await zip.file('[Content_Types].xml').async('string'), 'application/xml');
    const stylesFile = zip.file('word/styles.xml');
    const stylesDoc = stylesFile ? parser.parseFromString(await stylesFile.async('string'), 'application/xml') : null;
    const body = bodyOf(doc);
    const geom = pageGeom(doc);

    const map = {
      NOMBRE: data.nombre, FECHA: data.fecha, CENTRO: data.centro, DIRECCION: data.direccion,
      'DIRECCIÓN': data.direccion, POBLACION: data.poblacion, 'POBLACIÓN': data.poblacion, PROVINCIA: data.provincia || 'Badajoz'
    };

    // 1) Tabla de defectos
    const tbl = all(body, 'tbl').find(t => /\[(COD_DEF|CODIGO_APOYO|DES_DEF)\]/.test(ptext(t)));
    let tableAnchor = null; // elemento que queda en el lugar de la tabla
    if (tbl) {
      const rows = kids(tbl, 'tr');
      const tplRow = rows.find(r => /\[(COD_DEF|CODIGO_APOYO|DES_DEF)\]/.test(ptext(r)));
      if (data.sinDefectos || !data.defectos.length) {
        // referencia de tamaño: el párrafo anterior a la tabla
        let prev = tbl.previousSibling; while (prev && (prev.nodeType !== 1 || prev.localName !== 'p')) prev = prev.previousSibling;
        const para = sinDefectosPara(doc, baseRunInfo(stylesDoc, prev), prev);
        body.replaceChild(para, tbl);
        tableAnchor = para;
      } else {
        rows.forEach(r => { if (r !== tplRow && rows.indexOf(r) < rows.indexOf(tplRow)) { setRowFlag(r, 'cantSplit'); setRowFlag(r, 'tblHeader'); } });
        for (const d of data.defectos) {
          const row = tplRow.cloneNode(true);
          const m = {
            CODIGO_APOYO: up(d.apoyo), CENTRO: up(d.linea || data.centro), LINEA: up(d.linea || data.centro),
            COD_DEF: up(d.cod), DES_DEF: up(d.des), OBSERV: up(d.obs), OBSERVACIONES: up(d.obs)
          };
          all(row, 'p').forEach(p => replaceInPara(p, m));
          setRowFlag(row, 'cantSplit');
          tbl.insertBefore(row, tplRow);
        }
        tbl.removeChild(tplRow);
        tableAnchor = tbl;
      }
    }

    // 2) Campos sueltos
    all(body, 'p').forEach(p => replaceInPara(p, map));
    for (const f of Object.keys(zip.files)) {
      if (/^word\/(header|footer)\d*\.xml$/.test(f)) {
        const hd = parser.parseFromString(await zip.file(f).async('string'), 'application/xml');
        all(hd, 'p').forEach(p => replaceInPara(p, map));
        zip.file(f, ser.serializeToString(hd));
      }
    }

    // 3) Firma y bloque final (último punto + firma juntos en la misma página)
    const top = () => Array.from(body.childNodes).filter(n => n.nodeType === 1);
    const fdo = top().find(n => n.localName === 'p' && /Fdo\.?:/i.test(ptext(n)));
    let anexo1 = top().find(n => n.localName === 'p' && /^\s*ANEXO\s+I\b(?!I)/i.test(ptext(n)));

    if (fdo) {
      // la plantilla trae un hueco de imagen para la firma junto a "Fdo."
      let sigPara = all(fdo, 'drawing').length ? fdo : null;
      if (!sigPara) { let pv = fdo.previousSibling; while (pv && pv.nodeType !== 1) pv = pv.previousSibling; if (pv && pv.localName === 'p' && all(pv, 'drawing').length && !ptext(pv).trim()) sigPara = pv; }
      if (!data.firma && sigPara) removeDrawings(sigPara); // sin firma guardada: el informe sale sin firmar
      if (data.firma) {
        if (sigPara) {
          const ext0 = all(sigPara, 'extent', NS.wp)[0];
          await swapImage(sigPara, data.firma, zip, relsDoc, ctDoc, 'firma', +ext0.getAttribute('cx'), +ext0.getAttribute('cy'));
        } else {
          const rid = await addMedia(zip, relsDoc, ctDoc, data.firma, 'firma');
          const sz = emuSize(data.firma.w, data.firma.h, 4.5, 2.2);
          const fp = doc.createElementNS(W, 'w:p');
          const fdoPPr = kid(fdo, 'pPr'); if (fdoPPr) fp.appendChild(fdoPPr.cloneNode(true));
          fp.appendChild(drawingRun(doc, rid, sz.cx, sz.cy, 'Firma'));
          body.insertBefore(fp, fdo);
        }
      }
      const list = top();
      const start = tableAnchor ? list.indexOf(tableAnchor) + 1 : Math.max(0, list.indexOf(fdo) - 4);
      const end = list.indexOf(fdo);
      for (let i = start; i <= end; i++) {
        const n = list[i]; if (n.localName !== 'p') continue;
        setFlag(n, 'keepLines');
        if (i < end) setFlag(n, 'keepNext');
      }
      if (tableAnchor) {
        let prev = tableAnchor.previousSibling; while (prev && (prev.nodeType !== 1 || prev.localName !== 'p')) prev = prev.previousSibling;
        if (prev) setFlag(prev, 'keepNext');
      }
    }

    // 4) Anexos siempre en hoja nueva, sin riesgo de página en blanco
    if (anexo1) {
      let firstAfter = anexo1;
      if (fdo) {
        // se quitan los saltos de página manuales entre la firma y los anexos
        let n = fdo.nextSibling; const toRemove = [];
        while (n && n !== anexo1) {
          if (n.nodeType === 1 && n.localName === 'p' && !all(n, 'sectPr').length) {
            const onlyBreak = ptext(n).trim() === '' && !all(n, 'drawing').length;
            if (onlyBreak && all(n, 'br').some(b => wAttr(b, 'type') === 'page')) toRemove.push(n);
          }
          n = n.nextSibling;
        }
        toRemove.forEach(x => body.removeChild(x));
        // primer párrafo tras la firma: empieza la hoja de anexos
        n = fdo.nextSibling; while (n && n.nodeType !== 1) n = n.nextSibling;
        if (n && n.localName === 'p') firstAfter = n;
      }
      setFlag(firstAfter, 'pageBreakBefore');

      // 5) Foto del ANEXO I
      const geomCx = Math.round(geom.textWcm * 360000);
      if (all(anexo1, 'drawing').length) {
        if (data.foto) {
          const ext0 = all(anexo1, 'extent', NS.wp)[0];
          const bx = Math.max(+ext0.getAttribute('cx'), Math.min(geomCx - 567 * 635, 5400040));
          await swapImage(anexo1, data.foto, zip, relsDoc, ctDoc, 'anexo1', bx, 11.5 * 360000);
        } else removeDrawings(anexo1);
      } else if (data.foto) {
        const rid = await addMedia(zip, relsDoc, ctDoc, data.foto, 'anexo1');
        const sz = fitBox(data.foto.w, data.foto.h, Math.min(5400040, geomCx), 11.5 * 360000);
        const fp = doc.createElementNS(W, 'w:p');
        const a1PPr = kid(anexo1, 'pPr'); if (a1PPr) { const c = a1PPr.cloneNode(true); const pb = kid(c, 'pageBreakBefore'); if (pb) c.removeChild(pb); fp.appendChild(c); }
        fp.appendChild(drawingRun(doc, rid, sz.cx, sz.cy, 'AnexoI'));
        body.insertBefore(fp, anexo1.nextSibling);
      }
    }

    zip.file('word/document.xml', ser.serializeToString(doc));
    zip.file(relsPath, ser.serializeToString(relsDoc));
    zip.file('[Content_Types].xml', ser.serializeToString(ctDoc));
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  }
  const up = s => String(s || '').toLocaleUpperCase('es-ES').trim();

  // ---------- estilos (para el PDF) ----------
  function resolveStyles(stylesDoc) {
    const styles = {}; let defaultParaStyle = null;
    let defSz = 20, defAfter = 0, defLine = 240, defFont = null;
    if (stylesDoc) {
      const dd = all(stylesDoc, 'docDefaults')[0];
      if (dd) {
        const sz = all(dd, 'sz')[0]; if (sz) defSz = parseInt(wAttr(sz, 'val'), 10);
        const rf = all(dd, 'rFonts')[0]; if (rf && wAttr(rf, 'ascii')) defFont = wAttr(rf, 'ascii');
        const sp = all(dd, 'spacing')[0];
        if (sp) { if (wAttr(sp, 'after')) defAfter = +wAttr(sp, 'after'); if (wAttr(sp, 'line')) defLine = +wAttr(sp, 'line'); }
      }
      for (const s of all(stylesDoc, 'style')) {
        const id = wAttr(s, 'styleId'); const type = wAttr(s, 'type');
        if (type === 'paragraph' && (wAttr(s, 'default') === '1' || wAttr(s, 'default') === 'true')) defaultParaStyle = id;
        styles[id] = s;
      }
    }
    const chain = id => { const out = []; let g = 0; while (id && styles[id] && g++ < 10) { out.unshift(styles[id]); const b = kid(styles[id], 'basedOn'); id = b ? wAttr(b, 'val') : null; } return out; };
    function pProps(id, tblStyle) {
      const o = { after: defAfter, before: 0, line: defLine, jc: null, indL: 0, indR: 0, hang: 0, first: 0, numId: null, keepNext: false, sz: defSz, b: false, font: defFont, caps: false, u: false };
      // dentro de una tabla, el estilo de tabla va antes que el del párrafo (p. ej. «Tabla con cuadrícula» quita el espacio después)
      if (tblStyle) for (const ts of chain(tblStyle)) { const pPr = kid(ts, 'pPr'), rPr = kid(ts, 'rPr'); if (pPr) applyPPr(o, pPr); if (rPr) applyRPr(o, rPr); }
      const list = chain(id || defaultParaStyle);
      if (id && defaultParaStyle && id !== defaultParaStyle && !list.some(s => wAttr(s, 'styleId') === defaultParaStyle)) list.unshift(...chain(defaultParaStyle));
      for (const s of list) {
        const pPr = kid(s, 'pPr'), rPr = kid(s, 'rPr');
        if (pPr) applyPPr(o, pPr);
        if (rPr) applyRPr(o, rPr);
      }
      return o;
    }
    // estilos de carácter (p. ej. «Fuerte» = negrita)
    function charStyle(o, id) { for (const st of chain(id)) { const rPr = kid(st, 'rPr'); if (rPr) applyRPr(o, rPr); } }
    return { pProps, charStyle, defaultParaStyle, runSize: id => pProps(id).sz };
  }
  function applyPPr(o, pPr) {
    const sp = kid(pPr, 'spacing');
    if (sp) {
      if (wAttr(sp, 'after') != null && wAttr(sp, 'after') !== '') o.after = +wAttr(sp, 'after');
      if (wAttr(sp, 'before')) o.before = +wAttr(sp, 'before');
      if (wAttr(sp, 'line') && (wAttr(sp, 'lineRule') || 'auto') === 'auto') o.line = +wAttr(sp, 'line');
      const ba = wAttr(sp, 'beforeAutospacing'), aa = wAttr(sp, 'afterAutospacing');
      if (ba) o.beforeAuto = !['0', 'false'].includes(ba);
      if (aa) o.afterAuto = !['0', 'false'].includes(aa);
    }
    const jc = kid(pPr, 'jc'); if (jc) o.jc = wAttr(jc, 'val');
    const ind = kid(pPr, 'ind');
    if (ind) {
      const l = wAttr(ind, 'left') || wAttr(ind, 'start'); if (l) o.indL = +l;
      const rt = wAttr(ind, 'right') || wAttr(ind, 'end'); if (rt) o.indR = +rt; // sangría derecha
      if (wAttr(ind, 'hanging')) { o.hang = +wAttr(ind, 'hanging'); o.first = 0; }
      if (wAttr(ind, 'firstLine')) { o.first = +wAttr(ind, 'firstLine'); o.hang = 0; }
    }
    const np = kid(pPr, 'numPr');
    if (np) { const id = kid(np, 'numId'); o.numId = id ? wAttr(id, 'val') : o.numId; const lv = kid(np, 'ilvl'); o.ilvl = lv ? +wAttr(lv, 'val') : 0; }
    if (kid(pPr, 'keepNext')) o.keepNext = true;
  }
  function applyRPr(o, rPr) {
    const sz = kid(rPr, 'sz'); if (sz) o.sz = parseInt(wAttr(sz, 'val'), 10);
    const b = kid(rPr, 'b'); if (b) o.b = !['0', 'false'].includes(wAttr(b, 'val'));
    const f = kid(rPr, 'rFonts'); if (f && wAttr(f, 'ascii')) o.font = wAttr(f, 'ascii');
    const c = kid(rPr, 'caps'); if (c) o.caps = !['0', 'false'].includes(wAttr(c, 'val'));
    const u = kid(rPr, 'u'); if (u) o.u = !!wAttr(u, 'val') && wAttr(u, 'val') !== 'none';
  }

  // ---------- PDF ----------
  async function generarPdf(docxBytes) {
    const { PDFDocument, StandardFonts, rgb } = PDFLib;
    const zip = await JSZip.loadAsync(docxBytes);
    const parser = new DOMParser();
    const doc = parser.parseFromString(await zip.file('word/document.xml').async('string'), 'application/xml');
    const relsDoc = parser.parseFromString(await zip.file('word/_rels/document.xml.rels').async('string'), 'application/xml');
    const stylesFile = zip.file('word/styles.xml');
    const st = resolveStyles(stylesFile ? parser.parseFromString(await stylesFile.async('string'), 'application/xml') : null);
    const numFile = zip.file('word/numbering.xml');
    const numDoc = numFile ? parser.parseFromString(await numFile.async('string'), 'application/xml') : null;
    const geom = pageGeom(doc);
    const T = v => v / 20; // twips -> pt

    const pdf = await PDFDocument.create();
    const F = {
      sans: [await pdf.embedFont(StandardFonts.Helvetica), await pdf.embedFont(StandardFonts.HelveticaBold)],
      serif: [await pdf.embedFont(StandardFonts.TimesRoman), await pdf.embedFont(StandardFonts.TimesRomanBold)],
      mono: [await pdf.embedFont(StandardFonts.Courier), await pdf.embedFont(StandardFonts.CourierBold)]
    };
    const pickFont = (name, bold) => {
      const fam = /times|georgia|garamond|cambria|book|serif/i.test(name || '') ? 'serif' : /courier|consolas|mono/i.test(name || '') ? 'mono' : 'sans';
      return F[fam][bold ? 1 : 0];
    };
    const fR = F.sans[0];
    const numCount = {};
    const PW = T(geom.w), PH = T(geom.h), ML = T(geom.left), MR = T(geom.right), MT = T(geom.top), MB = T(geom.bottom);
    const CW = PW - ML - MR;
    const safeCache = {}; const fontIds = new Map();
    const safe = (s, f) => Array.from(s.replace(/ /g, ' ').replace(/[‐-―]/g, '-')).map(ch => {
      if (!fontIds.has(f)) fontIds.set(f, fontIds.size);
      const k = ch + '|' + fontIds.get(f);
      if (!(k in safeCache)) { try { f.widthOfTextAtSize(ch, 10); safeCache[k] = ch; } catch (e) { safeCache[k] = ch === '' ? '•' : '?'; } }
      return safeCache[k];
    }).join('');

    const relTarget = (rels, rid) => { const r = Array.from(rels.documentElement.getElementsByTagNameNS(NS.rel, 'Relationship')).find(x => x.getAttribute('Id') === rid); return r ? r.getAttribute('Target') : null; };
    const imgCache = {};
    async function loadImg(rels, baseDir, rid) {
      const tgt = relTarget(rels, rid); if (!tgt) return null;
      const path = (baseDir + tgt).replace(/[^/]+\/\.\.\//g, '');
      if (path in imgCache) return imgCache[path];
      const f = zip.file(path); if (!f) return (imgCache[path] = null);
      const bytes = await f.async('uint8array');
      let img = null;
      try { img = /\.png$/i.test(path) ? await pdf.embedPng(bytes) : /\.jpe?g$/i.test(path) ? await pdf.embedJpg(bytes) : null; } catch (e) { img = null; }
      return (imgCache[path] = img);
    }

    const roman = n => { const m = [[10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']]; let r = ''; for (const [v, c] of m) while (n >= v) { r += c; n -= v; } return r; };
    const indOf = lvl => { const o = {}; const pPr = kid(lvl, 'pPr'); const ind = pPr && kid(pPr, 'ind'); if (ind) { if (wAttr(ind, 'left')) o.indL = +wAttr(ind, 'left'); if (wAttr(ind, 'hanging')) o.hang = +wAttr(ind, 'hanging'); } return o; };
    function bulletFor(numId, ilvl) {
      const res = { text: '•', indL: 720, hang: 360 };
      if (!numDoc || !numId || numId === '0') return res;
      const num = all(numDoc, 'num').find(n => wAttr(n, 'numId') === numId);
      const absId = num && kid(num, 'abstractNumId') ? wAttr(kid(num, 'abstractNumId'), 'val') : null;
      const abs = all(numDoc, 'abstractNum').find(a => wAttr(a, 'abstractNumId') === absId);
      const lvl = abs && all(abs, 'lvl').find(l => +wAttr(l, 'ilvl') === (ilvl || 0));
      if (lvl) {
        const lt = kid(lvl, 'lvlText'); const fmt = kid(lvl, 'numFmt');
        let t = lt ? wAttr(lt, 'val') : '•';
        if (fmt && wAttr(fmt, 'val') !== 'bullet') {
          const key = absId + ':' + (ilvl || 0);
          numCount[key] = (numCount[key] || 0) + 1;
          const n = numCount[key], f = wAttr(fmt, 'val');
          const v = f === 'lowerLetter' ? String.fromCharCode(96 + n) : f === 'upperLetter' ? String.fromCharCode(64 + n) : f === 'lowerRoman' ? roman(n).toLowerCase() : f === 'upperRoman' ? roman(n) : String(n);
          return { ...res, text: t.replace(/%\d/, v), ...(indOf(lvl)) };
        }
        if (!t || /[-]/.test(t) || t === 'o' || t === '§') t = t === '-' ? '-' : '•';
        res.text = t;
        const pPr = kid(lvl, 'pPr'); const ind = pPr && kid(pPr, 'ind');
        if (ind) { if (wAttr(ind, 'left')) res.indL = +wAttr(ind, 'left'); if (wAttr(ind, 'hanging')) res.hang = +wAttr(ind, 'hanging'); }
      }
      return res;
    }

    // Convierte un párrafo Word a un bloque medible
    async function paraBlock(p, rels, baseDir, cellW, tblStyle) {
      const pPr = kid(p, 'pPr');
      const sid = pPr && kid(pPr, 'pStyle') ? wAttr(kid(pPr, 'pStyle'), 'val') : null;
      const o = st.pProps(sid, tblStyle);
      if (pPr) applyPPr(o, pPr);
      const pRPr = pPr && kid(pPr, 'rPr');
      const anchors = [];
      const items = []; // {type:'text', t, font, size} | {type:'img', img, w, h} | {type:'tab'}
      let leadTabs = 0, sawContent = false;
      const runs = all(p, 'r');
      for (const r of runs) {
        if (r.parentNode.localName === 'del') continue;
        const rp = { sz: o.sz, b: o.b, font: o.font, caps: o.caps, u: false };
        const rPr = kid(r, 'rPr');
        const rSt = rPr && kid(rPr, 'rStyle'); if (rSt) st.charStyle(rp, wAttr(rSt, 'val'));
        if (rPr) applyRPr(rp, rPr);
        for (const c of Array.from(r.childNodes)) {
          if (c.nodeType !== 1) continue;
          if (c.localName === 't') { if (c.textContent.length) { const tx = rp.caps ? c.textContent.toLocaleUpperCase('es-ES') : c.textContent; items.push({ type: 'text', t: tx, font: pickFont(rp.font, rp.b), u: rp.u, size: rp.sz / 2 }); if (c.textContent.trim()) sawContent = true; } }
          else if (c.localName === 'tab') { if (!sawContent && !items.some(i => i.type === 'img')) leadTabs++; else items.push({ type: 'text', t: '    ', font: pickFont(rp.font, rp.b), size: rp.sz / 2 }); }
          else if (c.localName === 'br' && wAttr(c, 'type') !== 'page') items.push({ type: 'br' });
          else if (c.localName === 'drawing') {
            const blip = all(c, 'blip', NS.a)[0]; const ext = all(c, 'extent', NS.wp)[0];
            const rid = blip && (blip.getAttributeNS(NS.r, 'embed') || blip.getAttribute('r:embed'));
            const img = rid ? await loadImg(rels, baseDir, rid) : null;
            const anc = all(c, 'anchor', NS.wp)[0];
            if (img && ext && anc) {
              const ph = all(anc, 'positionH', NS.wp)[0], pv = all(anc, 'positionV', NS.wp)[0];
              const offOf = el => { const o2 = el && el.getElementsByTagNameNS(NS.wp, 'posOffset')[0]; return o2 ? +o2.textContent / 12700 : 0; };
              anchors.push({ img, w: +ext.getAttribute('cx') / 12700, h: +ext.getAttribute('cy') / 12700, offX: offOf(ph), offY: offOf(pv), relH: ph ? ph.getAttribute('relativeFrom') : 'column' });
            } else if (img && ext) { items.push({ type: 'img', img, w: +ext.getAttribute('cx') / 12700, h: +ext.getAttribute('cy') / 12700 }); sawContent = true; }
          }
        }
      }
      let indL = T(o.indL), first = T(o.first) - T(o.hang), bullet = null;
      if (o.numId && o.numId !== '0') {
        const bf = bulletFor(o.numId, o.ilvl);
        if (!(pPr && kid(pPr, 'ind'))) { indL = T(bf.indL); first = -T(bf.hang); }
        bullet = bf.text;
      }
      const tabStop = 35.4; // 1,25 cm
      const leadOffset = leadTabs * tabStop;
      const width = cellW - indL - T(o.indR || 0);
      let markSz = o.sz;
      if (pRPr) { const mk = { sz: o.sz }; const rs = kid(pRPr, 'rStyle'); if (rs) st.charStyle(mk, wAttr(rs, 'val')); applyRPr(mk, pRPr); markSz = mk.sz; }
      const fsize = items.find(i => i.type === 'text') ? items.find(i => i.type === 'text').size : markSz / 2;
      const lineH = fsize * 1.15 * (o.line / 240); // interlineado sencillo de Times/Arial ≈ 1,15 × tamaño
      // trocear en palabras
      const words = [];
      for (const it of items) {
        if (it.type === 'img') { words.push(it); continue; }
        if (it.type === 'br') { words.push({ type: 'br' }); continue; }
        const font = it.font;
        const t = safe(it.t, font);
        const parts = t.split(/( +)/);
        for (const part of parts) {
          if (part === '') continue;
          if (/^ +$/.test(part)) { words.push({ type: 'space', w: font.widthOfTextAtSize(part, it.size), size: it.size }); continue; }
          words.push({ type: 'word', t: part, font, u: it.u, size: it.size, w: font.widthOfTextAtSize(part, it.size) });
        }
      }
      // líneas
      let endsBr = false;
      const lines = []; let cur = { items: [], w: 0, h: lineH, x0: 0 };
      const startX = (i) => (i === 0 ? first + leadOffset : leadOffset);
      const maxW = i => width - startX(i);
      const push = () => { while (cur.items.length && cur.items[cur.items.length - 1].type === 'space') { cur.w -= cur.items.pop().w; } lines.push(cur); cur = { items: [], w: 0, h: lineH }; };
      for (const wd of words) {
        if (wd.type === 'br') { push(); endsBr = true; continue; }
        endsBr = false;
        if (wd.type === 'space') { if (cur.items.length) { cur.items.push(wd); cur.w += wd.w; } continue; }
        const ww = wd.w;
        if (cur.w + ww > maxW(lines.length) && cur.items.length) push();
        if (wd.type === 'img') { let w = wd.w, h = wd.h; const mw = maxW(lines.length); if (w > mw) { h = h * mw / w; w = mw; } cur.items.push({ ...wd, w, h }); cur.w += w; cur.h = Math.max(cur.h, h + 2); continue; }
        // palabra más larga que la línea: partir por caracteres
        if (ww > maxW(lines.length)) {
          let chunk = '';
          for (const ch of wd.t) {
            const tw = wd.font.widthOfTextAtSize(chunk + ch, wd.size);
            if (tw > maxW(lines.length) - cur.w && chunk) { cur.items.push({ ...wd, t: chunk, w: wd.font.widthOfTextAtSize(chunk, wd.size) }); cur.w += wd.font.widthOfTextAtSize(chunk, wd.size); push(); chunk = ''; }
            chunk += ch;
          }
          if (chunk) { const w = wd.font.widthOfTextAtSize(chunk, wd.size); cur.items.push({ ...wd, t: chunk, w }); cur.w += w; }
          continue;
        }
        cur.items.push(wd); cur.w += ww;
      }
      if (cur.items.length || !lines.length || endsBr) push(); // un salto de línea final deja una línea en blanco, como en Word
      const before = o.beforeAuto ? 14 : T(o.before), after = o.afterAuto ? 14 : T(o.after);
      const h = before + lines.reduce((s, l) => s + l.h, 0) + after;
      return {
        kind: 'p', h, before, after, beforeAuto: !!o.beforeAuto, afterAuto: !!o.afterAuto, numId: o.numId && o.numId !== '0' ? o.numId : null, lines, indL, first, leadOffset, width, jc: o.jc, bullet, fsize, keepNext: o.keepNext,
        pageBreakBefore: !!(pPr && kid(pPr, 'pageBreakBefore')) || all(p, 'br').some(b => wAttr(b, 'type') === 'page'),
        draw(page, x, yTop, spaceBefore) {
          const sb = spaceBefore == null ? before : spaceBefore;
          for (const an of anchors) {
            const ax = an.relH === 'page' ? an.offX : x + an.offX;
            page.drawImage(an.img, { x: ax, y: yTop - sb - an.offY - an.h, width: an.w, height: an.h });
          }
          let y = yTop - sb;
          const fw = words.find(w2 => w2.type === 'word'); const bf = fw ? fw.font : fR;
          lines.forEach((ln, i) => {
            const sx = x + indL + (i === 0 ? first : 0) + leadOffset;
            const avail = width - (i === 0 ? first : 0) - leadOffset;
            let dx = 0, extra = 0;
            if (this.jc === 'center') dx = (avail - ln.w) / 2;
            else if (this.jc === 'right' || this.jc === 'end') dx = avail - ln.w;
            else if ((this.jc === 'both' || this.jc === 'distribute') && i < lines.length - 1) {
              const spaces = ln.items.filter(it => it.type === 'space').length;
              if (spaces) extra = (avail - ln.w) / spaces;
            }
            if (i === 0 && bullet) page.drawText(safe(bullet, bf), { x: x + indL + first, y: y - ln.h + (ln.h - fsize) * 0.35 + 1, size: fsize, font: bf, color: rgb(0, 0, 0) });
            let cx = sx + (i === 0 && bullet ? -first : 0) + dx;
            if (i === 0 && bullet) cx = x + indL + dx + leadOffset;
            const base = y - ln.h + (ln.h - fsize) * 0.35 + 1;
            for (const it of ln.items) {
              if (it.type === 'space') { cx += it.w + extra; continue; }
              if (it.type === 'img') { page.drawImage(it.img, { x: cx, y: y - it.h - 1, width: it.w, height: it.h }); cx += it.w; continue; }
              page.drawText(it.t, { x: cx, y: base, size: it.size, font: it.font, color: rgb(0, 0, 0) });
              if (it.u) page.drawLine({ start: { x: cx, y: base - 1.5 }, end: { x: cx + it.w, y: base - 1.5 }, thickness: 0.5, color: rgb(0, 0, 0) });
              cx += it.w;
            }
            y -= ln.h;
          });
        }
      };
    }

    async function tableBlock(tbl, rels, baseDir) {
      const grid = all(tbl, 'gridCol').map(g => T(+wAttr(g, 'w') || 0));
      let total = grid.reduce((a, b) => a + b, 0);
      const maxW = PW - 2 * 28;
      const scale = total > maxW ? maxW / total : 1;
      const tblPr = kid(tbl, 'tblPr');
      const tblStyle = tblPr && kid(tblPr, 'tblStyle') ? wAttr(kid(tblPr, 'tblStyle'), 'val') : null;
      const tjc = tblPr && kid(tblPr, 'jc') ? wAttr(kid(tblPr, 'jc'), 'val') : 'left';
      const cols = grid.map(g => g * scale);
      total = cols.reduce((a, b) => a + b, 0);
      const rows = [];
      for (const tr of kids(tbl, 'tr')) {
        const trPr = kid(tr, 'trPr');
        const header = !!(trPr && kid(trPr, 'tblHeader'));
        const cells = []; let ci = 0;
        for (const tc of kids(tr, 'tc')) {
          const tcPr = kid(tc, 'tcPr');
          const span = tcPr && kid(tcPr, 'gridSpan') ? +wAttr(kid(tcPr, 'gridSpan'), 'val') : 1;
          const w = cols.slice(ci, ci + span).reduce((a, b) => a + b, 0) || total / Math.max(1, kids(tr, 'tc').length);
          const shd = tcPr && kid(tcPr, 'shd'); const fill = shd ? wAttr(shd, 'fill') : null;
          const vAlignEl = tcPr && kid(tcPr, 'vAlign');
          const blocks = [];
          for (const p of kids(tc, 'p')) blocks.push(await paraBlock(p, rels, baseDir, w - 10.8, tblStyle));
          const h = blocks.reduce((s, b) => s + b.h, 0) + 0.5;
          cells.push({ x: cols.slice(0, ci).reduce((a, b) => a + b, 0), w, fill: fill && fill !== 'auto' ? fill : null, blocks, h, vAlign: vAlignEl ? wAttr(vAlignEl, 'val') : 'top' });
          ci += span;
        }
        const trH = trPr && kid(trPr, 'trHeight');
        const minH = trH ? T(+wAttr(trH, 'val') || 0) : 0; // altura mínima de la fila
        const h = trH && wAttr(trH, 'hRule') === 'exact' ? minH : Math.max(minH, 12, ...cells.map(c => c.h));
        rows.push({ header, cells, h });
      }
      const offX = tjc === 'center' ? (CW - total) / 2 : (tjc === 'right' || tjc === 'end') ? CW - total : 0;
      return { kind: 'tbl', rows, total, offX };
    }

    function hex(c) { const n = parseInt(c, 16); return rgb(((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255); }
    function drawRow(page, row, x, yTop) {
      for (const c of row.cells) {
        if (c.fill && /^[0-9a-f]{6}$/i.test(c.fill)) page.drawRectangle({ x: x + c.x, y: yTop - row.h, width: c.w, height: row.h, color: hex(c.fill) });
        page.drawRectangle({ x: x + c.x, y: yTop - row.h, width: c.w, height: row.h, borderColor: rgb(0, 0, 0), borderWidth: 0.5 });
        let y = yTop; // márgenes de celda arriba/abajo = 0 (tabla normal de Word)
        if (c.vAlign === 'center') y -= (row.h - c.h) / 2;
        for (const b of c.blocks) { b.draw(page, x + c.x + 5.4, y); y -= b.h; }
      }
    }

    // cabecera y pie de la plantilla
    const sect = all(doc, 'sectPr').pop();
    async function hfBlocks(kind) {
      if (!sect) return [];
      const ref = all(sect, kind + 'Reference').find(r => wAttr(r, 'type') === 'default') || all(sect, kind + 'Reference')[0];
      if (!ref) return [];
      const rid = ref.getAttributeNS(NS.r, 'id') || ref.getAttribute('r:id');
      const tgt = relTarget(relsDoc, rid); if (!tgt) return [];
      const f = zip.file('word/' + tgt); if (!f) return [];
      const hdoc = parser.parseFromString(await f.async('string'), 'application/xml');
      const relF = zip.file('word/_rels/' + tgt + '.rels');
      const hrels = relF ? parser.parseFromString(await relF.async('string'), 'application/xml') : parser.parseFromString(`<Relationships xmlns="${NS.rel}"/>`, 'application/xml');
      const out = [];
      for (const el of Array.from(hdoc.documentElement.childNodes)) {
        if (el.nodeType !== 1) continue;
        if (el.localName === 'p') {
          // imágenes ancladas: se tratan como en línea
          for (const anc of all(el, 'anchor', NS.wp)) {
            const blip = all(anc, 'blip', NS.a)[0]; const ext = all(anc, 'extent', NS.wp)[0];
            const rid2 = blip && (blip.getAttributeNS(NS.r, 'embed') || blip.getAttribute('r:embed'));
            const img = rid2 ? await loadImg(hrels, 'word/', rid2) : null;
            if (img && ext) {
              const posH = all(anc, 'positionH', NS.wp)[0]; const off = posH && posH.getElementsByTagNameNS(NS.wp, 'posOffset')[0];
              const al = posH && posH.getElementsByTagNameNS(NS.wp, 'align')[0];
              out.push({ kind: 'anchor', img, w: +ext.getAttribute('cx') / 12700, h: +ext.getAttribute('cy') / 12700, offX: off ? +off.textContent / 12700 : null, align: al ? al.textContent : null, rel: posH ? posH.getAttribute('relativeFrom') : 'column' });
            }
          }
          if (ptext(el).trim() || all(el, 'inline', NS.wp).length) out.push(await paraBlock(el, hrels, 'word/', CW));
        } else if (el.localName === 'tbl') out.push(await tableBlock(el, hrels, 'word/'));
      }
      return out;
    }
    const header = await hfBlocks('header');
    const footer = await hfBlocks('footer');
    function drawHF(page, blocks, yTop) {
      let y = yTop;
      for (const b of blocks) {
        if (b.kind === 'anchor') {
          let x = ML + (b.offX || 0);
          if (b.rel === 'page') x = b.offX || 0;
          if (b.align === 'right') x = ML + CW - b.w; else if (b.align === 'center') x = ML + (CW - b.w) / 2;
          page.drawImage(b.img, { x, y: yTop - b.h, width: b.w, height: b.h });
        } else if (b.kind === 'tbl') { for (const r of b.rows) { drawRow(page, r, ML, y); y -= r.h; } }
        else { b.draw(page, ML, y); y -= b.h; }
      }
    }
    const footerH = footer.reduce((s, b) => s + (b.kind === 'tbl' ? b.rows.reduce((a, r) => a + r.h, 0) : b.kind === 'anchor' ? 0 : b.h), 0);

    let page, y, prev = null;
    const newPage = () => {
      page = pdf.addPage([PW, PH]); y = PH - MT; prev = null;
      drawHF(page, header, PH - T(geom.header));
      drawHF(page, footer, T(geom.footer) + footerH);
    };
    newPage();
    const bottom = MB;

    // bloques del cuerpo
    const body = bodyOf(doc);
    const blocks = [];
    for (const el of Array.from(body.childNodes)) {
      if (el.nodeType !== 1) continue;
      if (el.localName === 'p') blocks.push(await paraBlock(el, relsDoc, 'word/', CW));
      else if (el.localName === 'tbl') blocks.push(await tableBlock(el, relsDoc, 'word/'));
    }
    const blockH = b => b.kind === 'tbl' ? b.rows.reduce((s, r) => s + r.h, 0) : b.h;
    const pageCap = PH - MT - bottom;

    // Espaciado entre párrafos como Word: con espaciado automático (HTML) no se suman los espacios,
    // entre elementos de la misma lista no hay espacio y arriba de página no se aplica el automático.
    const atTop = () => y >= PH - MT - 0.5;
    const gapFor = (b, top) => {
      const bef = top && b.beforeAuto ? 0 : b.before;
      if (!prev || top) return bef;
      if (prev.numId && b.numId && prev.numId === b.numId && (prev.afterAuto || b.beforeAuto)) return 0;
      if (prev.afterAuto || b.beforeAuto) return Math.max(prev.after, bef);
      return prev.after + bef;
    };
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i];
      if (b.kind === 'p' && b.pageBreakBefore && !atTop()) newPage();
      // mantener con el siguiente: calcular el grupo encadenado
      if (b.kind === 'p' && b.keepNext) {
        let gh = 0, j = i;
        while (j < blocks.length) {
          const bj = blocks[j];
          if (bj.kind === 'tbl') { const hdr = bj.rows.filter(r => r.header); gh += hdr.reduce((s, r) => s + r.h, 0) + (bj.rows.find(r => !r.header) || { h: 0 }).h; break; }
          gh += bj.h; if (!bj.keepNext) break; j++;
        }
        if (gh <= pageCap && y - gh < bottom) newPage();
      }
      if (b.kind === 'p') {
        const lh = b.h - b.before - b.after;
        let g = gapFor(b, atTop());
        if (y - g - lh < bottom && !atTop()) { newPage(); g = gapFor(b, true); }
        b.draw(page, ML, y, g); y -= g + lh; prev = b;
      } else {
        if (prev) y -= prev.after;
        const hdrRows = b.rows.filter(r => r.header);
        let first = true;
        for (const r of b.rows) {
          if (y - r.h < bottom) {
            newPage();
            if (!r.header) for (const hr of hdrRows) { drawRow(page, hr, ML + b.offX, y); y -= hr.h; }
          }
          drawRow(page, r, ML + b.offX, y); y -= r.h; first = false;
        }
        prev = { after: 0, before: 0 };
      }
    }
    return pdf.save();
  }

  global.MotorBT = { generarDocx, generarPdf, TXT_SIN };
})(window);
