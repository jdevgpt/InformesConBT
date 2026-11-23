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

  function sinDefectosPara(doc, info, refPara, texto) {
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
    const t = doc.createElementNS(W, 'w:t'); t.textContent = texto || TXT_SIN; r.appendChild(t); p.appendChild(r);
    return p;
  }

  /**
   * data = { nombre, fecha, centro, direccion, poblacion, sinDefectos, defectos:[{apoyo,linea,cod,des,obs}],
   *          firma:{bytes,type,w,h}|null, foto:{bytes,type,w,h}|null, tecnico (Boletín) }
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
      'DIRECCIÓN': data.direccion, POBLACION: data.poblacion, 'POBLACIÓN': data.poblacion, PROVINCIA: data.provincia || 'Badajoz',
      // Boletín BT (Málaga)
      NUM_CD: data.centro, TECNICO: data.tecnico || data.nombre, 'TÉCNICO': data.tecnico || data.nombre
    };
    // Párrafo [FOTO] (Boletín BT): la foto del ANEXO I va justo ahí
    const fotoPara = all(body, 'p').find(p => /\[FOTO\]/.test(ptext(p)));
    if (fotoPara) {
      all(fotoPara, 'r').forEach(r => r.parentNode.removeChild(r));
      if (data.foto) {
        const rid = await addMedia(zip, relsDoc, ctDoc, data.foto, 'anexo1');
        const sz = fitBox(data.foto.w, data.foto.h, Math.min(5400040, Math.round(geom.textWcm * 360000)), 11.5 * 360000);
        fotoPara.appendChild(drawingRun(doc, rid, sz.cx, sz.cy, 'AnexoI'));
      }
    }

    // 1) Tabla de defectos
    const tbl = all(body, 'tbl').find(t => /\[(COD_DEF|CODIGO_APOYO|DES_DEF)\]/.test(ptext(t)));
    let tableAnchor = null; // elemento que queda en el lugar de la tabla
    if (tbl) {
      const rows = kids(tbl, 'tr');
      const tplRow = rows.find(r => /\[(COD_DEF|CODIGO_APOYO|DES_DEF)\]/.test(ptext(r)));
      if (data.sinDefectos || !data.defectos.length) {
        // referencia de tamaño: el párrafo anterior a la tabla
        let prev = tbl.previousSibling; while (prev && (prev.nodeType !== 1 || prev.localName !== 'p')) prev = prev.previousSibling;
        const para = sinDefectosPara(doc, baseRunInfo(stylesDoc, prev), prev, data.textoSinDefectos);
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

    // 2) Campos sueltos. El número del CD del Boletín ([NUM_CD]) va en negrita y 1 punto más grande.
    all(body, 'p').forEach(p => {
      all(p, 'r').forEach(r => {
        if (!/\[NUM_CD\]/.test(all(r, 't').map(t => t.textContent).join(''))) return;
        const base = baseRunInfo(stylesDoc, p).sz;
        let rPr = kid(r, 'rPr'); if (!rPr) { rPr = doc.createElementNS(W, 'w:rPr'); r.insertBefore(rPr, r.firstChild); }
        ['b', 'bCs', 'sz', 'szCs'].forEach(n => { const o = kid(rPr, n); if (o) rPr.removeChild(o); });
        rPr.appendChild(doc.createElementNS(W, 'w:b')); rPr.appendChild(doc.createElementNS(W, 'w:bCs'));
        ['sz', 'szCs'].forEach(n => { const e = doc.createElementNS(W, 'w:' + n); e.setAttributeNS(W, 'w:val', String(base + 2)); rPr.appendChild(e); });
      });
      replaceInPara(p, map);
    });
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
      // Se puede cortar de página antes de un bloque nuevo («Nota aclaratoria…», «Y para que conste…»),
      // pero el último bloque siempre va junto con «Fdo.» y nunca se parte un párrafo.
      const empiezaBloque = n => n && n.localName === 'p' && /^\s*(Nota aclaratoria|Y para que conste)/i.test(ptext(n));
      for (let i = start; i <= end; i++) {
        const n = list[i]; if (n.localName !== 'p') continue;
        setFlag(n, 'keepLines');
        if (i < end && !empiezaBloque(list.slice(i + 1).find(x => x.localName === 'p' && ptext(x).trim()))) setFlag(n, 'keepNext');
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

      // 5) Foto del ANEXO I (si la plantilla no trae [FOTO])
      const geomCx = Math.round(geom.textWcm * 360000);
      if (fotoPara) { /* ya colocada */ }
      else if (all(anexo1, 'drawing').length) {
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
    let defSz = 20, defAfter = 0, defLine = 240, defFont = null, defKern = 0;
    if (stylesDoc) {
      const dd = all(stylesDoc, 'docDefaults')[0];
      if (dd) {
        const sz = all(dd, 'sz')[0]; if (sz) defSz = parseInt(wAttr(sz, 'val'), 10);
        const rf = all(dd, 'rFonts')[0]; if (rf && wAttr(rf, 'ascii')) defFont = wAttr(rf, 'ascii'); else if (rf && wAttr(rf, 'asciiTheme')) defFont = '@' + wAttr(rf, 'asciiTheme');
        const kn = all(dd, 'kern')[0]; if (kn) defKern = +wAttr(kn, 'val') || 0;
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
      const o = { after: defAfter, before: 0, line: defLine, jc: null, indL: 0, indR: 0, hang: 0, first: 0, numId: null, keepNext: false, sz: defSz, b: false, font: defFont, caps: false, u: false, kern: defKern };
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
    const f = kid(rPr, 'rFonts'); if (f && wAttr(f, 'ascii')) o.font = wAttr(f, 'ascii'); else if (f && wAttr(f, 'asciiTheme')) o.font = '@' + wAttr(f, 'asciiTheme');
    const c = kid(rPr, 'caps'); if (c) o.caps = !['0', 'false'].includes(wAttr(c, 'val'));
    const u = kid(rPr, 'u'); if (u) o.u = !!wAttr(u, 'val') && wAttr(u, 'val') !== 'none';
    const k = kid(rPr, 'kern'); if (k) o.kern = +wAttr(k, 'val') || 0;
  }

  // ---------- PDF ----------
  // Anchos de letra (milésimas de em) de Times New Roman, Calibri y Arial, sacados de sus fuentes
  // equivalentes (Liberation Serif, Carlito, Liberation Sans). Las fuentes estándar del PDF son algo
  // distintas; midiendo con estas, las líneas cortan donde en Word y se ajusta cada palabra a su ancho.
  const METRICAS = { codigos: [32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95,96,97,98,99,100,101,102,103,104,105,106,107,108,109,110,111,112,113,114,115,116,117,118,119,120,121,122,123,124,125,126,160,161,162,163,164,165,166,167,168,169,170,171,172,173,174,175,176,177,178,179,180,181,182,183,184,185,186,187,188,189,190,191,192,193,194,195,196,197,198,199,200,201,202,203,204,205,206,207,208,209,210,211,212,213,214,215,216,217,218,219,220,221,222,223,224,225,226,227,228,229,230,231,232,233,234,235,236,237,238,239,240,241,242,243,244,245,246,247,248,249,250,251,252,253,254,255,8211,8212,8216,8217,8220,8221,8226,8230,8364,160], tnr: [[250,333,408,500,500,833,778,180,333,333,500,564,250,333,250,278,500,500,500,500,500,500,500,500,500,500,278,278,564,564,564,444,921,722,667,667,722,611,556,722,722,333,389,722,611,889,722,722,556,722,667,556,611,722,722,944,722,722,611,333,278,333,469,500,333,444,500,444,500,444,333,500,500,278,278,500,278,778,500,500,500,500,333,389,278,500,500,722,500,500,444,480,200,480,541,250,333,500,500,500,500,200,500,333,760,276,500,564,333,760,500,400,549,300,300,333,576,453,333,333,300,310,500,750,750,750,444,722,722,722,722,722,722,889,667,611,611,611,611,333,333,333,333,722,722,722,722,722,722,722,564,722,722,722,722,722,722,556,500,444,444,444,444,444,444,667,444,444,444,444,444,278,278,278,278,500,500,500,500,500,500,500,549,500,500,500,500,500,500,500,500,500,1000,333,333,444,444,350,1000,500,250], [250,333,555,500,500,1000,833,278,333,333,500,570,250,333,250,278,500,500,500,500,500,500,500,500,500,500,333,333,570,570,570,500,930,722,667,722,722,667,611,778,778,389,500,778,667,944,722,778,611,778,722,556,667,722,722,1000,722,722,667,333,278,333,581,500,333,500,556,444,556,444,333,500,556,278,333,556,278,833,556,500,556,556,444,389,333,556,500,722,500,500,444,394,220,394,520,250,333,500,500,500,500,220,500,333,747,300,500,570,333,747,500,400,549,300,300,333,576,540,333,333,300,330,500,750,750,750,500,722,722,722,722,722,722,1000,722,667,667,667,667,389,389,389,389,722,722,778,778,778,778,778,570,778,722,722,722,722,722,611,556,500,500,500,500,500,500,722,444,444,444,444,444,278,278,278,278,500,556,500,500,500,500,500,549,500,556,556,556,556,500,556,500,500,1000,333,333,500,500,350,1000,500,250]], calibri: [[226,326,401,498,507,715,682,221,303,303,498,498,250,306,252,386,507,507,507,507,507,507,507,507,507,507,268,268,498,498,498,463,894,579,544,533,615,488,459,631,623,252,319,520,420,855,646,662,517,673,543,459,487,642,567,890,519,487,468,307,386,307,498,498,291,479,525,423,525,498,305,471,525,229,239,455,229,799,525,527,525,525,349,391,335,525,452,715,433,453,395,314,460,314,498,226,326,498,507,498,507,498,498,393,834,402,512,498,0,507,394,339,498,336,334,292,550,586,252,307,246,422,512,636,671,675,463,579,579,579,579,579,579,763,533,488,488,488,488,252,252,252,252,625,646,662,662,662,662,662,498,664,642,642,642,642,487,517,527,479,479,479,479,479,479,773,423,498,498,498,498,229,229,229,229,525,525,527,527,527,527,527,498,529,525,525,525,525,453,525,453,498,905,250,250,418,418,498,690,507,226], [226,326,438,498,507,729,705,233,312,312,498,498,258,306,267,430,507,507,507,507,507,507,507,507,507,507,276,276,498,498,498,463,898,606,561,529,630,488,459,637,631,267,331,547,423,874,659,676,532,686,563,473,495,653,591,906,551,520,478,325,430,325,498,498,300,494,537,418,537,503,316,474,537,246,255,480,246,813,537,538,537,537,355,399,347,537,473,745,459,474,397,344,475,344,498,226,326,498,507,498,507,498,498,415,834,416,539,498,0,507,390,342,498,338,336,301,563,598,268,303,252,435,539,658,691,702,463,606,606,606,606,606,606,775,529,488,488,488,488,267,267,267,267,639,659,676,676,676,676,676,498,681,653,653,653,653,520,532,555,494,494,494,494,494,494,775,418,503,503,503,503,246,246,246,246,537,537,538,538,538,538,538,498,544,537,537,537,537,474,537,474,498,905,258,258,435,435,498,711,507,226]], arial: [[278,278,355,556,556,889,667,191,333,333,389,584,278,333,278,278,556,556,556,556,556,556,556,556,556,556,278,278,584,584,584,556,1015,667,667,722,722,667,611,778,722,278,500,667,556,833,722,778,667,778,722,667,611,722,667,944,667,667,611,278,278,278,469,556,333,556,556,500,556,556,278,556,556,222,222,500,222,833,556,556,556,556,333,500,278,556,500,722,500,500,500,334,260,334,584,278,333,556,556,556,556,260,556,333,737,370,556,584,333,737,552,400,549,333,333,333,576,537,333,333,333,365,556,834,834,834,611,667,667,667,667,667,667,1000,722,667,667,667,667,278,278,278,278,722,722,778,778,778,778,778,584,778,722,722,722,722,667,667,611,556,556,556,556,556,556,889,500,556,556,556,556,278,278,278,278,556,556,556,556,556,556,556,549,611,556,556,556,556,500,556,500,556,1000,222,222,333,333,350,1000,556,278], [278,333,474,556,556,889,722,238,333,333,389,584,278,333,278,278,556,556,556,556,556,556,556,556,556,556,333,333,584,584,584,611,975,722,722,722,722,667,611,778,722,278,556,722,611,833,722,778,667,778,722,667,611,722,667,944,667,667,611,333,278,333,584,556,333,556,611,556,611,556,333,611,611,278,278,556,278,889,611,611,611,611,389,556,333,611,556,778,556,556,500,389,280,389,584,278,333,556,556,556,556,280,556,333,737,370,556,584,333,737,552,400,549,333,333,333,576,556,333,333,333,365,556,834,834,834,611,722,722,722,722,722,722,1000,722,667,667,667,667,278,278,278,278,722,722,778,778,778,778,778,584,778,722,722,722,722,667,667,611,556,556,556,556,556,556,889,556,556,556,556,556,278,278,278,278,611,611,611,611,611,611,611,549,611,611,611,611,611,556,611,556,556,1000,278,278,500,500,350,1000,556,278]] };
  // Kerning (pares de letras que Word acerca, p. ej. «VA») cuando el documento lo tiene activado (w:kern)
  const KERNING = {"tnr":[{" A":-55," T":-18," V":-18," W":-18," Y":-37,"11":-37,"A ":-55,"AT":-111,"AV":-129,"AW":-80,"AY":-92,"Av":-74,"Aw":-92,"Ay":-92,"A’":-111,"F,":-80,"F.":-80,"FA":-74,"L ":-37,"LT":-92,"LV":-92,"LW":-74,"LY":-100,"Ly":-55,"L’":-92,"P ":-37,"P,":-111,"P.":-111,"PA":-92,"RT":-60,"RV":-80,"RW":-55,"RY":-55,"Ry":-40,"T ":-18,"T,":-74,"T-":-92,"T.":-74,"T:":-50,"T;":-55,"TA":-80,"TO":-18,"Ta":-70,"Tc":-70,"Te":-70,"Ti":-35,"To":-70,"Tr":-35,"Ts":-70,"Tu":-35,"Tw":-70,"Ty":-70,"V ":-18,"V,":-129,"V-":-92,"V.":-129,"V:":-74,"V;":-74,"VA":-129,"Va":-111,"Ve":-111,"Vi":-60,"Vo":-129,"Vr":-60,"Vu":-60,"Vy":-111,"W ":-18,"W,":-92,"W-":-55,"W.":-92,"W:":-37,"W;":-37,"WA":-111,"Wa":-80,"We":-80,"Wi":-40,"Wo":-80,"Wr":-40,"Wu":-40,"Wy":-60,"Y ":-37,"Y,":-129,"Y-":-111,"Y.":-129,"Y:":-92,"Y;":-92,"YA":-111,"Ya":-100,"Ye":-100,"Yi":-55,"Yo":-100,"Yp":-92,"Yq":-111,"Yu":-111,"Yv":-100,"ff":-18,"f’":55,"r,":-40,"r-":-20,"r.":-55,"rg":-18,"r’":37,"v,":-65,"v.":-65,"w,":-65,"w.":-65,"y,":-65,"y.":-65,"‘‘":-74,"’ ":-74,"’s":-55,"’t":-18,"’’":-74},{" A":-55," T":-18," V":-18," W":-18," Y":-37,"11":-55,"A ":-55,"AT":-74,"AV":-129,"AW":-111,"AY":-92,"Av":-74,"Aw":-74,"Ay":-74,"A’":-74,"F ":-37,"F,":-92,"F.":-92,"FA":-74,"L ":-55,"LT":-92,"LV":-92,"LW":-92,"LY":-92,"Ly":-55,"L’":-92,"P ":-55,"P,":-92,"P.":-92,"PA":-74,"RT":-35,"RV":-35,"RW":-35,"RY":-35,"Ry":-35,"T ":-18,"T,":-74,"T-":-92,"T.":-74,"T:":-74,"T;":-74,"TA":-74,"TO":-18,"Ta":-92,"Tc":-92,"Te":-92,"Ti":-18,"To":-92,"Tr":-74,"Ts":-92,"Tu":-92,"Tw":-74,"Ty":-74,"V ":-18,"V,":-129,"V-":-74,"V.":-129,"V:":-92,"V;":-92,"VA":-129,"VO":-20,"Va":-92,"Ve":-92,"Vi":-37,"Vo":-92,"Vr":-74,"Vu":-92,"Vy":-92,"W ":-18,"W,":-92,"W-":-37,"W.":-92,"W:":-55,"W;":-55,"WA":-111,"Wa":-55,"We":-55,"Wi":-18,"Wo":-55,"Wr":-18,"Wu":-18,"Wy":-37,"Y ":-37,"Y,":-92,"Y-":-92,"Y.":-92,"Y:":-92,"Y;":-92,"YA":-92,"Ya":-111,"Ye":-111,"Yi":-37,"Yo":-111,"Yp":-92,"Yq":-111,"Yu":-92,"Yv":-111,"f’":55,"r ":-18,"r,":-92,"r-":-37,"r.":-92,"rc":-18,"re":-18,"ro":-18,"rq":-18,"r’":18,"v,":-55,"v.":-55,"w,":-55,"w.":-55,"y,":-55,"y.":-55,"‘‘":-74,"’ ":-74,"’s":-37,"’’":-74}],"calibri":[{"Q)":15,"Q,":51,"Q/":66,"Q;":29,"QJ":20,"QT":-23,"QV":-12,"QW":-6,"QX":6,"QY":-22,"Q]":16,"Qg":29,"Qj":39,"Qx":15,"Q}":15,"QÆ":36,"QÝ":-22,"Q’":4,"Q”":4,"ßf":-15,"ßt":-15,"ßu":-8,"ßv":-15,"ßw":-19,"ßx":-15,"ßy":-21,"ßz":-13,"ßý":-21,"ßÿ":-21,"ß‘":-23,"ß“":-23,"v,":-78,"v-":-23,"v.":-82,"va":-15,"vc":-12,"vd":-10,"ve":-10,"vf":5,"vg":-14,"vo":-9,"vq":-10,"vs":-4,"vt":5,"vv":6,"vw":6,"vy":6,"vz":-13,"và":-15,"vá":-15,"vâ":-15,"vã":-15,"vä":-15,"vå":-15,"væ":-15,"vç":-12,"vè":-10,"vé":-10,"vê":-10,"vë":-10,"vò":-9,"vó":-9,"vô":-9,"võ":-9,"vö":-9,"vø":-9,"vý":6,"vÿ":6,"v‘":7,"v’":29,"v“":7,"v”":29,",A":8,",J":11,",T":-65,",V":-68,",W":-73,",X":5,",Y":-72,",t":-21,",À":8,",Á":8,",Â":8,",Ã":8,",Ä":8,",Å":8,",Ý":-72,",–":-26,",‘":-88,",’":-68,",“":-88,",”":-68,".-":-55,".A":6,".C":-19,".G":-19,".J":9,".O":-11,".Q":-11,".T":-75,".U":-9,".V":-74,".W":-70,".Y":-92,".f":-20,".g":11,".t":-21,".v":-38,".w":-33,".y":-38,".z":5,".À":6,".Á":6,".Â":6,".Ã":6,".Ä":6,".Å":6,".Ç":-19,".Ò":-11,".Ó":-11,".Ô":-11,".Õ":-11,".Ö":-11,".Ù":-9,".Ú":-9,".Û":-9,".Ü":-9,".Ý":-92,".ý":-38,".ÿ":-38,".‘":-95,".’":-73,".“":-95,".”":-73,"/A":-35,"/J":-18,"/a":-18,"/c":-18,"/e":-18,"/g":-17,"/o":-17,"/s":-19,"/²":-122,"/³":-129,"/¹":-93,"/À":-35,"/Á":-35,"/Â":-35,"/Ã":-35,"/Ä":-35,"/Å":-35,"/Æ":-34,"/à":-18,"/á":-18,"/â":-18,"/ã":-18,"/ä":-18,"/å":-18,"/æ":-18,"/ç":-18,"/è":-18,"/é":-18,"/ê":-18,"/ë":-18,"/ò":-17,"/ó":-17,"/ô":-17,"/õ":-17,"/ö":-17,"/ø":-17,"-Y":-49,"-f":-9,"-v":-9,"-x":-19,"-y":-10,"-Ý":-49,"-ý":-10,"-ÿ":-10,"(g":5,"(j":31,"++":7,"²/":-142,"³/":-156,"¹/":-134,"¿T":-50,"¿V":-47,"¿Y":-54,"¿f":-23,"¿j":11,"¿v":-32,"¿w":-36,"¿y":-33,"¿Ý":-54,"¿ý":-33,"¿ÿ":-33,"í‘":4,"í’":36,"í“":4,"í”":36,"A,":7,"A-":-7,"A.":6,"A?":-33,"AC":-7,"AG":-7,"AJ":11,"AO":-11,"AQ":-11,"AT":-78,"AU":-16,"AV":-43,"AW":-39,"AY":-73,"At":-25,"Av":-19,"Ay":-20,"AÇ":-7,"AÒ":-11,"AÓ":-11,"AÔ":-11,"AÕ":-11,"AÖ":-11,"AÙ":-16,"AÚ":-16,"AÛ":-16,"AÜ":-16,"AÝ":-73,"Aý":-20,"Aÿ":-20,"A‘":-88,"A’":-38,"A“":-88,"A”":-38,"Æ-":-13,"ÆA":-11,"ÆC":-12,"ÆG":-12,"ÆO":-16,"ÆQ":-16,"ÆS":-10,"ÆZ":-5,"Æa":-17,"Æc":-14,"Æd":-15,"Æe":-18,"Æf":-31,"Æo":-18,"Æq":-15,"Æt":-12,"Æv":-23,"Æw":-17,"Æy":-23,"ÆÀ":-11,"ÆÁ":-11,"ÆÂ":-11,"ÆÃ":-11,"ÆÄ":-11,"ÆÅ":-11,"ÆÇ":-12,"ÆÒ":-16,"ÆÓ":-16,"ÆÔ":-16,"ÆÕ":-16,"ÆÖ":-16,"Æà":-17,"Æá":-17,"Æâ":-17,"Æã":-17,"Æä":-17,"Æå":-17,"Ææ":-17,"Æç":-14,"Æè":-18,"Æé":-18,"Æê":-18,"Æë":-18,"Æò":-18,"Æó":-18,"Æô":-18,"Æõ":-18,"Æö":-18,"Æø":-18,"Æý":-23,"Æÿ":-23,"B,":-14,"BA":-10,"BT":-23,"BV":-12,"BW":-12,"BX":-21,"BY":-28,"BZ":-10,"Bf":-10,"Bt":-10,"Bv":-10,"Bx":-7,"By":-10,"BÀ":-10,"BÁ":-10,"BÂ":-10,"BÃ":-10,"BÄ":-10,"BÅ":-10,"BÝ":-28,"Bý":-10,"Bÿ":-10,"B‘":-14,"B“":-14,"C,":-4,"CG":-9,"CJ":6,"CO":-9,"CQ":-9,"CT":5,"CÆ":10,"CÒ":-9,"CÓ":-9,"CÔ":-9,"CÕ":-9,"CÖ":-9,"C‘":6,"C’":43,"C“":6,"C”":43,"Ç,":-4,"ÇG":-9,"ÇJ":6,"ÇO":-9,"ÇQ":-9,"ÇT":5,"ÇÆ":10,"ÇÒ":-9,"ÇÓ":-9,"ÇÔ":-9,"ÇÕ":-9,"ÇÖ":-9,"Ç‘":6,"Ç’":43,"Ç“":6,"Ç”":43,"D,":-27,"D.":-23,"DA":-15,"DJ":-11,"DT":-11,"DV":-12,"DW":-7,"DX":-15,"DY":-19,"DZ":-11,"DÀ":-15,"DÁ":-15,"DÂ":-15,"DÃ":-15,"DÄ":-15,"DÅ":-15,"DÆ":-13,"DÝ":-19,"D‘":-8,"D’":1,"D“":-8,"D”":1,"Ð,":-27,"Ð.":-23,"ÐA":-15,"ÐJ":-11,"ÐT":-11,"ÐV":-12,"ÐW":-7,"ÐX":-15,"ÐY":-19,"ÐZ":-11,"ÐÀ":-15,"ÐÁ":-15,"ÐÂ":-15,"ÐÃ":-15,"ÐÄ":-15,"ÐÅ":-15,"ÐÆ":-13,"ÐÝ":-19,"Ð‘":-8,"Ð’":1,"Ð“":-8,"Ð”":1,"E-":-13,"EA":-11,"EC":-12,"EG":-12,"EO":-16,"EQ":-16,"ES":-10,"EZ":-5,"Ea":-17,"Ec":-14,"Ed":-15,"Ee":-18,"Ef":-31,"Eo":-18,"Eq":-15,"Et":-12,"Ev":-23,"Ew":-17,"Ey":-23,"EÀ":-11,"EÁ":-11,"EÂ":-11,"EÃ":-11,"EÄ":-11,"EÅ":-11,"EÇ":-12,"EÒ":-16,"EÓ":-16,"EÔ":-16,"EÕ":-16,"EÖ":-16,"Eà":-17,"Eá":-17,"Eâ":-17,"Eã":-17,"Eä":-17,"Eå":-17,"Eæ":-17,"Eç":-14,"Eè":-18,"Eé":-18,"Eê":-18,"Eë":-18,"Eò":-18,"Eó":-18,"Eô":-18,"Eõ":-18,"Eö":-18,"Eø":-18,"Eý":-23,"Eÿ":-23,"F,":-100,"F.":-94,"F/":-32,"FA":-56,"FC":-9,"FG":-9,"FJ":-53,"FO":-9,"FQ":-9,"FS":-14,"FX":-11,"FZ":-5,"Fa":-27,"Fc":-14,"Fd":-10,"Fe":-15,"Fo":-14,"Fq":-10,"Fs":-17,"FÀ":-56,"FÁ":-56,"FÂ":-56,"FÃ":-56,"FÄ":-56,"FÅ":-56,"FÆ":-46,"FÇ":-9,"FÒ":-9,"FÓ":-9,"FÔ":-9,"FÕ":-9,"FÖ":-9,"Fà":-27,"Fá":-27,"Fâ":-27,"Fã":-27,"Fä":-27,"Få":-27,"Fæ":-27,"Fç":-14,"Fè":-15,"Fé":-15,"Fê":-15,"Fë":-15,"Fî":9,"Fï":12,"Fò":-14,"Fó":-14,"Fô":-14,"Fõ":-14,"Fö":-14,"Fø":-14,"F’":1,"F”":1,"GT":-5,"GV":-5,"GW":-4,"GY":-15,"Gv":-14,"Gw":-11,"Gx":-7,"Gy":-15,"GÆ":-5,"GÝ":-15,"Gý":-15,"Gÿ":-15,"Hî":6,"Hï":6,"H‘":-8,"H“":-8,"J,":-13,"J.":-9,"JA":-17,"JX":-10,"JÀ":-17,"JÁ":-17,"JÂ":-17,"JÃ":-17,"JÄ":-17,"JÅ":-17,"JÆ":-12,"K-":-17,"KC":-38,"KG":-39,"KO":-47,"KQ":-47,"KS":-9,"KU":-14,"KW":-17,"Ka":-17,"Kc":-20,"Kd":-16,"Ke":-18,"Kf":-12,"Km":-16,"Kn":-16,"Ko":-18,"Kp":-16,"Kq":-16,"Kr":-16,"Ks":-9,"Kt":-19,"Ku":-16,"Kv":-49,"Kw":-46,"Ky":-42,"KÇ":-38,"KÒ":-47,"KÓ":-47,"KÔ":-47,"KÕ":-47,"KÖ":-47,"KÙ":-14,"KÚ":-14,"KÛ":-14,"KÜ":-14,"Kà":-17,"Ká":-17,"Kâ":-17,"Kã":-17,"Kä":-17,"Kå":-17,"Kæ":-17,"Kç":-20,"Kè":-18,"Ké":-18,"Kê":-18,"Kë":-18,"Kì":14,"Kî":17,"Kï":14,"Kñ":-16,"Kò":-18,"Kó":-18,"Kô":-18,"Kõ":-18,"Kö":-18,"Kø":-18,"Kù":-16,"Kú":-16,"Kû":-16,"Kü":-16,"Ký":-42,"Kÿ":-42,"L,":8,"LC":-11,"LG":-23,"LJ":12,"LO":-22,"LQ":-22,"LT":-73,"LU":-21,"LV":-72,"LW":-58,"LY":-82,"Lf":-11,"Lt":-19,"Lv":-38,"Lw":-35,"Ly":-39,"L·":-27,"LÆ":31,"LÇ":-11,"LÒ":-22,"LÓ":-22,"LÔ":-22,"LÕ":-22,"LÖ":-22,"LÙ":-21,"LÚ":-21,"LÛ":-21,"LÜ":-21,"LÝ":-82,"Lý":-39,"Lÿ":-39,"L‘":-96,"L’":-72,"L“":-96,"L”":-72,"O,":-23,"O.":-15,"OA":-11,"OJ":-13,"OT":-27,"OV":-12,"OW":-11,"OX":-31,"OY":-27,"OZ":-19,"Ox":-6,"Oz":-5,"OÀ":-11,"OÁ":-11,"OÂ":-11,"OÃ":-11,"OÄ":-11,"OÅ":-11,"OÆ":-14,"OÝ":-27,"O‘":-9,"O’":6,"O“":-9,"O”":6,"P,":-124,"P-":-30,"P.":-127,"P/":-53,"PA":-74,"PJ":-68,"PT":-4,"PV":-5,"PX":-17,"PY":-5,"PZ":-14,"Pa":-21,"Pc":-21,"Pd":-17,"Pe":-20,"Pf":6,"Po":-20,"Pq":-17,"Ps":-16,"Pt":6,"Py":6,"PÀ":-74,"PÁ":-74,"PÂ":-74,"PÃ":-74,"PÄ":-74,"PÅ":-74,"PÆ":-68,"PÝ":-5,"Pà":-21,"Pá":-21,"Pâ":-21,"Pã":-21,"Pä":-21,"På":-21,"Pæ":-21,"Pç":-21,"Pè":-20,"Pé":-20,"Pê":-20,"Pë":-20,"Pò":-20,"Pó":-20,"Pô":-20,"Põ":-20,"Pö":-20,"Pø":-20,"Pý":6,"Pÿ":6,"P‘":6,"P’":23,"P“":6,"P”":23,"R.":5,"RC":-9,"RG":-9,"RO":-10,"RQ":-10,"RS":-13,"RT":-10,"RV":-14,"RW":-9,"RY":-15,"Re":-18,"Ro":-21,"Rv":-13,"Rw":-16,"Ry":-16,"RÇ":-9,"RÒ":-10,"RÓ":-10,"RÔ":-10,"RÕ":-10,"RÖ":-10,"RÝ":-15,"Rè":-18,"Ré":-18,"Rê":-18,"Rë":-18,"Rò":-21,"Ró":-21,"Rô":-21,"Rõ":-21,"Rö":-21,"Rø":-21,"Rý":-16,"Rÿ":-16,"S-":-4,"SA":-7,"SJ":-4,"ST":-7,"SV":-7,"SW":-7,"SX":-6,"SY":-10,"Sv":-11,"Sw":-8,"Sy":-12,"SÀ":-7,"SÁ":-7,"SÂ":-7,"SÃ":-7,"SÄ":-7,"SÅ":-7,"SÆ":-2,"SÝ":-10,"Sý":-12,"Sÿ":-12,"S‘":-10,"S’":-7,"S“":-10,"S”":-7,"T,":-103,"T-":-78,"T.":-99,"T/":-52,"T:":-65,"T;":-56,"TA":-78,"TC":-21,"TG":-29,"TJ":-32,"TO":-28,"TQ":-28,"TS":-5,"TT":14,"Ta":-78,"Tc":-86,"Td":-72,"Te":-89,"Tg":-74,"Tm":-62,"Tn":-62,"To":-89,"Tp":-62,"Tq":-72,"Tr":-62,"Ts":-75,"Tu":-62,"Tv":-45,"Tw":-42,"Tx":-44,"Ty":-45,"Tz":-69,"TÀ":-78,"TÁ":-78,"TÂ":-78,"TÃ":-78,"TÄ":-78,"TÅ":-78,"TÆ":-83,"TÇ":-21,"TÒ":-28,"TÓ":-28,"TÔ":-28,"TÕ":-28,"TÖ":-28,"Tà":-78,"Tá":-78,"Tâ":-78,"Tã":-78,"Tä":-78,"Tå":-78,"Tæ":-78,"Tç":-86,"Tè":-89,"Té":-89,"Tê":-89,"Të":-89,"Tì":19,"Tî":32,"Tï":41,"Tñ":-62,"Tò":-89,"Tó":-89,"Tô":-89,"Tõ":-89,"Tö":-89,"Tø":-89,"Tù":-62,"Tú":-62,"Tû":-62,"Tü":-62,"Tý":-45,"Tÿ":-45,"T‘":11,"T’":23,"T“":11,"T”":23,"U,":-14,"U.":-18,"UA":-22,"UJ":-20,"UÀ":-22,"UÁ":-22,"UÂ":-22,"UÃ":-22,"UÄ":-22,"UÅ":-22,"UÆ":-17,"V,":-81,"V-":-49,"V.":-100,"V/":-52,"V:":-36,"V;":-53,"VA":-47,"VC":-9,"VG":-12,"VJ":-39,"VO":-13,"VQ":-13,"VS":-6,"VV":4,"Va":-56,"Vc":-50,"Vd":-42,"Ve":-50,"Vg":-49,"Vm":-24,"Vn":-24,"Vo":-42,"Vp":-24,"Vq":-42,"Vr":-24,"Vs":-44,"Vu":-24,"Vy":-17,"Vz":-40,"VÀ":-47,"VÁ":-47,"VÂ":-47,"VÃ":-47,"VÄ":-47,"VÅ":-47,"VÆ":-72,"VÇ":-9,"VÒ":-13,"VÓ":-13,"VÔ":-13,"VÕ":-13,"VÖ":-13,"Và":-56,"Vá":-56,"Vâ":-56,"Vã":-56,"Vä":-56,"Vå":-56,"Væ":-56,"Vç":-50,"Vè":-50,"Vé":-50,"Vê":-50,"Vë":-50,"Vì":23,"Vî":32,"Vï":35,"Vñ":-24,"Vò":-42,"Vó":-42,"Vô":-42,"Võ":-42,"Vö":-42,"Vø":-42,"Vù":-24,"Vú":-24,"Vû":-24,"Vü":-24,"Vý":-17,"Vÿ":-17,"V’":26,"V”":26,"W,":-105,"W-":-43,"W.":-101,"W;":-76,"WA":-45,"WC":-11,"WG":-11,"WJ":-43,"WO":-11,"WQ":-11,"WS":-5,"WX":-6,"Wa":-35,"Wc":-38,"Wd":-35,"We":-37,"Wg":-26,"Wm":-29,"Wn":-29,"Wo":-42,"Wp":-29,"Wq":-35,"Wr":-29,"Ws":-36,"Wu":-29,"Wv":-17,"Wy":-26,"WÀ":-45,"WÁ":-45,"WÂ":-45,"WÃ":-45,"WÄ":-45,"WÅ":-45,"WÆ":-60,"WÇ":-11,"WÒ":-11,"WÓ":-11,"WÔ":-11,"WÕ":-11,"WÖ":-11,"Wà":-35,"Wá":-35,"Wâ":-35,"Wã":-35,"Wä":-35,"Wå":-35,"Wæ":-35,"Wç":-38,"Wè":-37,"Wé":-37,"Wê":-37,"Wë":-37,"Wì":8,"Wî":19,"Wï":23,"Wñ":-29,"Wò":-42,"Wó":-42,"Wô":-42,"Wõ":-42,"Wö":-42,"Wø":-42,"Wù":-29,"Wú":-29,"Wû":-29,"Wü":-29,"Wý":-26,"Wÿ":-26,"W’":7,"W”":7,"X-":-40,"XC":-28,"XG":-32,"XO":-28,"XQ":-28,"XS":-10,"Xd":-21,"Xe":-19,"Xg":-4,"Xo":-19,"Xq":-21,"Xt":-15,"Xu":-19,"Xv":-27,"Xw":-24,"Xy":-21,"XÇ":-28,"XÒ":-28,"XÓ":-28,"XÔ":-28,"XÕ":-28,"XÖ":-28,"Xè":-19,"Xé":-19,"Xê":-19,"Xë":-19,"Xò":-19,"Xó":-19,"Xô":-19,"Xõ":-19,"Xö":-19,"Xø":-19,"Xý":-21,"Xÿ":-21,"X’":-7,"X”":-7,"Y,":-123,"Y-":-57,"Y.":-116,"Y/":-61,"Y:":-75,"Y;":-67,"YA":-74,"YC":-33,"YG":-33,"YJ":-55,"YO":-32,"YQ":-32,"YS":-8,"YZ":-5,"Ya":-65,"Yc":-78,"Yd":-64,"Ye":-72,"Yf":-30,"Yg":-69,"Yi":-16,"Yj":-24,"Ym":-46,"Yn":-46,"Yo":-75,"Yp":-46,"Yq":-64,"Yr":-46,"Ys":-56,"Yt":-21,"Yu":-46,"Yv":-34,"Yw":-30,"Yx":-34,"Yy":-32,"Yz":-49,"YÀ":-74,"YÁ":-74,"YÂ":-74,"YÃ":-74,"YÄ":-74,"YÅ":-74,"YÆ":-86,"YÇ":-33,"YÒ":-32,"YÓ":-32,"YÔ":-32,"YÕ":-32,"YÖ":-32,"Yß":-35,"Yà":-65,"Yá":-65,"Yâ":-65,"Yã":-65,"Yä":-65,"Yå":-65,"Yæ":-65,"Yç":-78,"Yè":-72,"Yé":-72,"Yê":-72,"Yë":-72,"Yì":20,"Yí":-34,"Yî":23,"Yï":25,"Yñ":-46,"Yò":-75,"Yó":-75,"Yô":-75,"Yõ":-75,"Yö":-75,"Yø":-75,"Yù":-46,"Yú":-46,"Yû":-46,"Yü":-46,"Yý":-32,"Yÿ":-32,"Y’":13,"Y”":13,"Z-":-22,"ZA":-5,"ZC":-12,"ZG":-12,"ZO":-12,"ZQ":-12,"ZW":-3,"ZY":-3,"Za":-5,"Zc":-6,"Zd":-9,"Ze":-15,"Zo":-14,"Zq":-9,"Zv":-22,"Zw":-19,"Zy":-18,"ZÀ":-5,"ZÁ":-5,"ZÂ":-5,"ZÃ":-5,"ZÄ":-5,"ZÅ":-5,"ZÇ":-12,"ZÒ":-12,"ZÓ":-12,"ZÔ":-12,"ZÕ":-12,"ZÖ":-12,"ZÝ":-3,"Zà":-5,"Zá":-5,"Zâ":-5,"Zã":-5,"Zä":-5,"Zå":-5,"Zæ":-5,"Zç":-6,"Zè":-15,"Zé":-15,"Zê":-15,"Zë":-15,"Zì":8,"Zî":21,"Zï":27,"Zò":-14,"Zó":-14,"Zô":-14,"Zõ":-14,"Zö":-14,"Zø":-14,"Zý":-18,"Zÿ":-18,"Z’":22,"Z”":22,"af":-6,"at":-9,"av":-17,"aw":-7,"ax":-9,"ay":-19,"aý":-19,"aÿ":-19,"a‘":-18,"a“":-18,"æf":-9,"æt":-5,"æv":-5,"æw":-5,"æx":-15,"æy":-6,"æz":-10,"æý":-6,"æÿ":-6,"æ‘":-9,"æ“":-9,"bf":-8,"bs":-5,"bt":-4,"bv":-5,"bw":-5,"bx":-20,"by":-5,"bz":-14,"bý":-5,"bÿ":-5,"b‘":-11,"b“":-11,"ca":-8,"co":-8,"cà":-8,"cá":-8,"câ":-8,"cã":-8,"cä":-8,"cå":-8,"cæ":-8,"cò":-8,"có":-8,"cô":-8,"cõ":-8,"cö":-8,"cø":-8,"ça":-8,"ço":-8,"çà":-8,"çá":-8,"çâ":-8,"çã":-8,"çä":-8,"çå":-8,"çæ":-8,"çò":-8,"çó":-8,"çô":-8,"çõ":-8,"çö":-8,"çø":-8,"ef":-9,"et":-5,"ev":-5,"ew":-5,"ex":-15,"ey":-6,"ez":-10,"eý":-6,"eÿ":-6,"e‘":-9,"e“":-9,"f)":15,"f,":-62,"f-":-27,"f.":-66,"fa":-20,"fc":-22,"fd":-26,"fe":-25,"ff":-10,"fg":-29,"fo":-21,"fq":-26,"fs":-13,"fv":6,"fw":3,"fy":5,"fz":-10,"f®":11,"fà":-20,"fá":-20,"fâ":-20,"fã":-20,"fä":-20,"få":-20,"fæ":-20,"fç":-22,"fè":-25,"fé":-25,"fê":-25,"fë":-25,"fò":-21,"fó":-21,"fô":-21,"fõ":-21,"fö":-21,"fø":-21,"fý":5,"fÿ":5,"f‘":12,"f’":53,"f“":12,"f”":53,"g,":11,"g/":36,"ga":-19,"gc":-6,"gd":-9,"ge":-8,"gg":9,"go":-7,"gq":-9,"gt":-15,"gà":-19,"gá":-19,"gâ":-19,"gã":-19,"gä":-19,"gå":-19,"gæ":-19,"gç":-6,"gè":-8,"gé":-8,"gê":-8,"gë":-8,"gò":-7,"gó":-7,"gô":-7,"gõ":-7,"gö":-7,"gø":-7,"g‘":6,"g’":41,"g“":6,"g”":41,"hf":-6,"ht":-9,"hv":-17,"hw":-7,"hx":-9,"hy":-19,"hý":-19,"hÿ":-19,"h‘":-18,"h“":-18,"i’":11,"i”":11,"j’":11,"j”":11,"k-":-38,"ka":-17,"kc":-23,"kd":-27,"ke":-32,"ko":-34,"kq":-27,"ks":-9,"kt":-5,"ku":-13,"kà":-17,"ká":-17,"kâ":-17,"kã":-17,"kä":-17,"kå":-17,"kæ":-17,"kç":-23,"kè":-32,"ké":-32,"kê":-32,"kë":-32,"kò":-34,"kó":-34,"kô":-34,"kõ":-34,"kö":-34,"kø":-34,"mf":-6,"mt":-9,"mv":-17,"mw":-7,"mx":-9,"my":-19,"mý":-19,"mÿ":-19,"m‘":-18,"m“":-18,"nf":-6,"nt":-9,"nv":-17,"nw":-7,"nx":-9,"ny":-19,"ný":-19,"nÿ":-19,"n‘":-18,"n“":-18,"o,":-18,"ov":-4,"ow":-4,"ox":-20,"oy":-5,"oz":-13,"oý":-5,"oÿ":-5,"ø,":-18,"øv":-4,"øw":-4,"øx":-20,"øy":-5,"øz":-13,"øý":-5,"øÿ":-5,"pf":-8,"ps":-5,"pt":-4,"pv":-5,"pw":-5,"px":-20,"py":-5,"pz":-14,"pý":-5,"pÿ":-5,"p‘":-11,"p“":-11,"þf":-8,"þs":-5,"þt":-4,"þv":-5,"þw":-5,"þx":-20,"þy":-5,"þz":-14,"þý":-5,"þÿ":-5,"þ‘":-11,"þ“":-11,"qg":5,"r,":-86,"r-":-29,"r.":-101,"ra":-21,"rc":-15,"rd":-14,"re":-13,"rg":-14,"ro":-16,"rq":-14,"rs":-17,"rv":9,"rw":5,"ry":5,"rà":-21,"rá":-21,"râ":-21,"rã":-21,"rä":-21,"rå":-21,"ræ":-21,"rç":-15,"rè":-13,"ré":-13,"rê":-13,"rë":-13,"rò":-16,"ró":-16,"rô":-16,"rõ":-16,"rö":-16,"rø":-16,"rý":5,"rÿ":5,"r‘":14,"r’":44,"r“":14,"r”":44,"sf":-9,"st":-11,"sv":-15,"sw":-5,"sx":-11,"sy":-18,"sz":-9,"sý":-18,"sÿ":-18,"s‘":-13,"s“":-13,"t-":-31,"ta":-12,"tc":-12,"td":-11,"te":-11,"to":-10,"tq":-11,"tt":-14,"tà":-12,"tá":-12,"tâ":-12,"tã":-12,"tä":-12,"tå":-12,"tæ":-12,"tç":-12,"tè":-11,"té":-11,"tê":-11,"të":-11,"tò":-10,"tó":-10,"tô":-10,"tõ":-10,"tö":-10,"tø":-10,"t‘":5,"t’":35,"t“":5,"t”":35,"w,":-86,"w-":-21,"w.":-64,"wa":-11,"wc":-10,"wd":-9,"we":-9,"wf":3,"wg":-9,"wo":-9,"wq":-9,"ws":-9,"wt":2,"wv":6,"ww":4,"wy":6,"wz":-8,"wà":-11,"wá":-11,"wâ":-11,"wã":-11,"wä":-11,"wå":-11,"wæ":-11,"wç":-10,"wè":-9,"wé":-9,"wê":-9,"wë":-9,"wò":-9,"wó":-9,"wô":-9,"wõ":-9,"wö":-9,"wø":-9,"wý":6,"wÿ":6,"w‘":7,"w’":29,"w“":7,"w”":29,"x-":-30,"xa":-18,"xc":-22,"xd":-21,"xe":-26,"xo":-27,"xq":-21,"xs":-6,"xt":3,"xu":-10,"xà":-18,"xá":-18,"xâ":-18,"xã":-18,"xä":-18,"xå":-18,"xæ":-18,"xç":-22,"xè":-26,"xé":-26,"xê":-26,"xë":-26,"xò":-27,"xó":-27,"xô":-27,"xõ":-27,"xö":-27,"xø":-27,"x’":25,"x”":25,"y,":-71,"y-":-22,"y.":-66,"ya":-15,"yc":-13,"yd":-12,"ye":-12,"yf":5,"yg":-13,"yo":-12,"yq":-12,"ys":-9,"yt":5,"yv":6,"yw":4,"yy":5,"yz":-8,"yà":-15,"yá":-15,"yâ":-15,"yã":-15,"yä":-15,"yå":-15,"yæ":-15,"yç":-13,"yè":-12,"yé":-12,"yê":-12,"yë":-12,"yò":-12,"yó":-12,"yô":-12,"yõ":-12,"yö":-12,"yø":-12,"yý":5,"yÿ":5,"y‘":7,"y’":30,"y“":7,"y”":30,"za":-17,"zc":-22,"zd":-22,"ze":-22,"zf":-5,"zg":-8,"zo":-22,"zq":-22,"zs":-11,"zu":-5,"zv":-9,"zw":-11,"zy":-9,"zà":-17,"zá":-17,"zâ":-17,"zã":-17,"zä":-17,"zå":-17,"zæ":-17,"zç":-22,"zè":-22,"zé":-22,"zê":-22,"zë":-22,"zò":-22,"zó":-22,"zô":-22,"zõ":-22,"zö":-22,"zø":-22,"zý":-9,"zÿ":-9,"z’":14,"z”":14,"?’":17,"?”":17,"‘,":-93,"‘.":-93,"‘A":-84,"‘G":-18,"‘J":-44,"‘O":-11,"‘Q":-11,"‘T":44,"‘V":27,"‘W":5,"‘Y":21,"‘a":-24,"‘c":-46,"‘d":-45,"‘e":-42,"‘g":-40,"‘o":-34,"‘q":-45,"‘s":-23,"‘À":-84,"‘Á":-84,"‘Â":-84,"‘Ã":-84,"‘Ä":-84,"‘Å":-84,"‘Æ":-96,"‘Ò":-11,"‘Ó":-11,"‘Ô":-11,"‘Õ":-11,"‘Ö":-11,"‘Ý":21,"‘à":-24,"‘á":-24,"‘â":-24,"‘ã":-24,"‘ä":-24,"‘å":-24,"‘æ":-24,"‘ç":-46,"‘è":-42,"‘é":-42,"‘ê":-42,"‘ë":-42,"‘ì":38,"‘î":41,"‘ï":34,"‘ò":-34,"‘ó":-34,"‘ô":-34,"‘õ":-34,"‘ö":-34,"‘ø":-34,"’,":-97,"’.":-95,"’A":-129,"’C":-33,"’G":-38,"’J":-63,"’O":-38,"’Q":-38,"’T":7,"’Y":6,"’a":-73,"’c":-72,"’d":-71,"’e":-71,"’g":-69,"’i":4,"’j":-5,"’m":-22,"’n":-22,"’o":-72,"’p":-22,"’q":-71,"’r":-22,"’s":-61,"’u":-22,"’v":-13,"’w":-13,"’x":-23,"’y":-21,"’z":-12,"’À":-129,"’Á":-129,"’Â":-129,"’Ã":-129,"’Ä":-129,"’Å":-129,"’Æ":-147,"’Ç":-33,"’Ò":-38,"’Ó":-38,"’Ô":-38,"’Õ":-38,"’Ö":-38,"’Ý":6,"’à":-73,"’á":-73,"’â":-73,"’ã":-73,"’ä":-73,"’å":-73,"’æ":-73,"’ç":-72,"’è":-71,"’é":-71,"’ê":-71,"’ë":-71,"’ì":30,"’î":19,"’ï":21,"’ñ":-22,"’ò":-72,"’ó":-72,"’ô":-72,"’õ":-72,"’ö":-72,"’ø":-72,"’ù":-22,"’ú":-22,"’û":-22,"’ü":-22,"’ý":-21,"’ÿ":-21,"“,":-93,"“.":-93,"“A":-84,"“G":-18,"“J":-44,"“O":-11,"“Q":-11,"“T":44,"“V":27,"“W":5,"“Y":21,"“a":-24,"“c":-46,"“d":-45,"“e":-42,"“g":-40,"“o":-34,"“q":-45,"“s":-23,"“À":-84,"“Á":-84,"“Â":-84,"“Ã":-84,"“Ä":-84,"“Å":-84,"“Æ":-96,"“Ò":-11,"“Ó":-11,"“Ô":-11,"“Õ":-11,"“Ö":-11,"“Ý":21,"“à":-24,"“á":-24,"“â":-24,"“ã":-24,"“ä":-24,"“å":-24,"“æ":-24,"“ç":-46,"“è":-42,"“é":-42,"“ê":-42,"“ë":-42,"“ì":38,"“î":41,"“ï":34,"“ò":-34,"“ó":-34,"“ô":-34,"“õ":-34,"“ö":-34,"“ø":-34,"”,":-97,"”.":-95,"”A":-129,"”C":-33,"”G":-38,"”J":-63,"”O":-38,"”Q":-38,"”T":7,"”Y":6,"”a":-73,"”c":-72,"”d":-71,"”e":-71,"”g":-69,"”i":4,"”j":-5,"”m":-22,"”n":-22,"”o":-72,"”p":-22,"”q":-71,"”r":-22,"”s":-61,"”u":-22,"”v":-13,"”w":-13,"”x":-23,"”y":-21,"”z":-12,"”À":-129,"”Á":-129,"”Â":-129,"”Ã":-129,"”Ä":-129,"”Å":-129,"”Æ":-147,"”Ç":-33,"”Ò":-38,"”Ó":-38,"”Ô":-38,"”Õ":-38,"”Ö":-38,"”Ý":6,"”à":-73,"”á":-73,"”â":-73,"”ã":-73,"”ä":-73,"”å":-73,"”æ":-73,"”ç":-72,"”è":-71,"”é":-71,"”ê":-71,"”ë":-71,"”ì":30,"”î":19,"”ï":21,"”ñ":-22,"”ò":-72,"”ó":-72,"”ô":-72,"”õ":-72,"”ö":-72,"”ø":-72,"”ù":-22,"”ú":-22,"”û":-22,"”ü":-22,"”ý":-21,"”ÿ":-21,"À,":7,"À-":-7,"À.":6,"À?":-33,"ÀC":-7,"ÀG":-7,"ÀJ":11,"ÀO":-11,"ÀQ":-11,"ÀT":-78,"ÀU":-16,"ÀV":-43,"ÀW":-39,"ÀY":-73,"Àt":-25,"Àv":-19,"Ày":-20,"ÀÇ":-7,"ÀÒ":-11,"ÀÓ":-11,"ÀÔ":-11,"ÀÕ":-11,"ÀÖ":-11,"ÀÙ":-16,"ÀÚ":-16,"ÀÛ":-16,"ÀÜ":-16,"ÀÝ":-73,"Àý":-20,"Àÿ":-20,"À‘":-88,"À’":-38,"À“":-88,"À”":-38,"Á,":7,"Á-":-7,"Á.":6,"Á?":-33,"ÁC":-7,"ÁG":-7,"ÁJ":11,"ÁO":-11,"ÁQ":-11,"ÁT":-78,"ÁU":-16,"ÁV":-43,"ÁW":-39,"ÁY":-73,"Át":-25,"Áv":-19,"Áy":-20,"ÁÇ":-7,"ÁÒ":-11,"ÁÓ":-11,"ÁÔ":-11,"ÁÕ":-11,"ÁÖ":-11,"ÁÙ":-16,"ÁÚ":-16,"ÁÛ":-16,"ÁÜ":-16,"ÁÝ":-73,"Áý":-20,"Áÿ":-20,"Á‘":-88,"Á’":-38,"Á“":-88,"Á”":-38,"Â,":7,"Â-":-7,"Â.":6,"Â?":-33,"ÂC":-7,"ÂG":-7,"ÂJ":11,"ÂO":-11,"ÂQ":-11,"ÂT":-78,"ÂU":-16,"ÂV":-43,"ÂW":-39,"ÂY":-73,"Ât":-25,"Âv":-19,"Ây":-20,"ÂÇ":-7,"ÂÒ":-11,"ÂÓ":-11,"ÂÔ":-11,"ÂÕ":-11,"ÂÖ":-11,"ÂÙ":-16,"ÂÚ":-16,"ÂÛ":-16,"ÂÜ":-16,"ÂÝ":-73,"Âý":-20,"Âÿ":-20,"Â‘":-88,"Â’":-38,"Â“":-88,"Â”":-38,"Ã,":7,"Ã-":-7,"Ã.":6,"Ã?":-33,"ÃC":-7,"ÃG":-7,"ÃJ":11,"ÃO":-11,"ÃQ":-11,"ÃT":-78,"ÃU":-16,"ÃV":-43,"ÃW":-39,"ÃY":-73,"Ãt":-25,"Ãv":-19,"Ãy":-20,"ÃÇ":-7,"ÃÒ":-11,"ÃÓ":-11,"ÃÔ":-11,"ÃÕ":-11,"ÃÖ":-11,"ÃÙ":-16,"ÃÚ":-16,"ÃÛ":-16,"ÃÜ":-16,"ÃÝ":-73,"Ãý":-20,"Ãÿ":-20,"Ã‘":-88,"Ã’":-38,"Ã“":-88,"Ã”":-38,"Ä,":7,"Ä-":-7,"Ä.":6,"Ä?":-33,"ÄC":-7,"ÄG":-7,"ÄJ":11,"ÄO":-11,"ÄQ":-11,"ÄT":-78,"ÄU":-16,"ÄV":-43,"ÄW":-39,"ÄY":-73,"Ät":-25,"Äv":-19,"Äy":-20,"ÄÇ":-7,"ÄÒ":-11,"ÄÓ":-11,"ÄÔ":-11,"ÄÕ":-11,"ÄÖ":-11,"ÄÙ":-16,"ÄÚ":-16,"ÄÛ":-16,"ÄÜ":-16,"ÄÝ":-73,"Äý":-20,"Äÿ":-20,"Ä‘":-88,"Ä’":-38,"Ä“":-88,"Ä”":-38,"Å,":7,"Å-":-7,"Å.":6,"Å?":-33,"ÅC":-7,"ÅG":-7,"ÅJ":11,"ÅO":-11,"ÅQ":-11,"ÅT":-78,"ÅU":-16,"ÅV":-43,"ÅW":-39,"ÅY":-73,"Åt":-25,"Åv":-19,"Åy":-20,"ÅÇ":-7,"ÅÒ":-11,"ÅÓ":-11,"ÅÔ":-11,"ÅÕ":-11,"ÅÖ":-11,"ÅÙ":-16,"ÅÚ":-16,"ÅÛ":-16,"ÅÜ":-16,"ÅÝ":-73,"Åý":-20,"Åÿ":-20,"Å‘":-88,"Å’":-38,"Å“":-88,"Å”":-38,"È-":-13,"ÈA":-11,"ÈC":-12,"ÈG":-12,"ÈO":-16,"ÈQ":-16,"ÈS":-10,"ÈZ":-5,"Èa":-17,"Èc":-14,"Èd":-15,"Èe":-18,"Èf":-31,"Èo":-18,"Èq":-15,"Èt":-12,"Èv":-23,"Èw":-17,"Èy":-23,"ÈÀ":-11,"ÈÁ":-11,"ÈÂ":-11,"ÈÃ":-11,"ÈÄ":-11,"ÈÅ":-11,"ÈÇ":-12,"ÈÒ":-16,"ÈÓ":-16,"ÈÔ":-16,"ÈÕ":-16,"ÈÖ":-16,"Èà":-17,"Èá":-17,"Èâ":-17,"Èã":-17,"Èä":-17,"Èå":-17,"Èæ":-17,"Èç":-14,"Èè":-18,"Èé":-18,"Èê":-18,"Èë":-18,"Èò":-18,"Èó":-18,"Èô":-18,"Èõ":-18,"Èö":-18,"Èø":-18,"Èý":-23,"Èÿ":-23,"É-":-13,"ÉA":-11,"ÉC":-12,"ÉG":-12,"ÉO":-16,"ÉQ":-16,"ÉS":-10,"ÉZ":-5,"Éa":-17,"Éc":-14,"Éd":-15,"Ée":-18,"Éf":-31,"Éo":-18,"Éq":-15,"Ét":-12,"Év":-23,"Éw":-17,"Éy":-23,"ÉÀ":-11,"ÉÁ":-11,"ÉÂ":-11,"ÉÃ":-11,"ÉÄ":-11,"ÉÅ":-11,"ÉÇ":-12,"ÉÒ":-16,"ÉÓ":-16,"ÉÔ":-16,"ÉÕ":-16,"ÉÖ":-16,"Éà":-17,"Éá":-17,"Éâ":-17,"Éã":-17,"Éä":-17,"Éå":-17,"Éæ":-17,"Éç":-14,"Éè":-18,"Éé":-18,"Éê":-18,"Éë":-18,"Éò":-18,"Éó":-18,"Éô":-18,"Éõ":-18,"Éö":-18,"Éø":-18,"Éý":-23,"Éÿ":-23,"Ê-":-13,"ÊA":-11,"ÊC":-12,"ÊG":-12,"ÊO":-16,"ÊQ":-16,"ÊS":-10,"ÊZ":-5,"Êa":-17,"Êc":-14,"Êd":-15,"Êe":-18,"Êf":-31,"Êo":-18,"Êq":-15,"Êt":-12,"Êv":-23,"Êw":-17,"Êy":-23,"ÊÀ":-11,"ÊÁ":-11,"ÊÂ":-11,"ÊÃ":-11,"ÊÄ":-11,"ÊÅ":-11,"ÊÇ":-12,"ÊÒ":-16,"ÊÓ":-16,"ÊÔ":-16,"ÊÕ":-16,"ÊÖ":-16,"Êà":-17,"Êá":-17,"Êâ":-17,"Êã":-17,"Êä":-17,"Êå":-17,"Êæ":-17,"Êç":-14,"Êè":-18,"Êé":-18,"Êê":-18,"Êë":-18,"Êò":-18,"Êó":-18,"Êô":-18,"Êõ":-18,"Êö":-18,"Êø":-18,"Êý":-23,"Êÿ":-23,"Ë-":-13,"ËA":-11,"ËC":-12,"ËG":-12,"ËO":-16,"ËQ":-16,"ËS":-10,"ËZ":-5,"Ëa":-17,"Ëc":-14,"Ëd":-15,"Ëe":-18,"Ëf":-31,"Ëo":-18,"Ëq":-15,"Ët":-12,"Ëv":-23,"Ëw":-17,"Ëy":-23,"ËÀ":-11,"ËÁ":-11,"ËÂ":-11,"ËÃ":-11,"ËÄ":-11,"ËÅ":-11,"ËÇ":-12,"ËÒ":-16,"ËÓ":-16,"ËÔ":-16,"ËÕ":-16,"ËÖ":-16,"Ëà":-17,"Ëá":-17,"Ëâ":-17,"Ëã":-17,"Ëä":-17,"Ëå":-17,"Ëæ":-17,"Ëç":-14,"Ëè":-18,"Ëé":-18,"Ëê":-18,"Ëë":-18,"Ëò":-18,"Ëó":-18,"Ëô":-18,"Ëõ":-18,"Ëö":-18,"Ëø":-18,"Ëý":-23,"Ëÿ":-23,"Ò,":-23,"Ò.":-15,"ÒA":-11,"ÒJ":-13,"ÒT":-27,"ÒV":-12,"ÒW":-11,"ÒX":-31,"ÒY":-27,"ÒZ":-19,"Òx":-6,"Òz":-5,"ÒÀ":-11,"ÒÁ":-11,"ÒÂ":-11,"ÒÃ":-11,"ÒÄ":-11,"ÒÅ":-11,"ÒÆ":-14,"ÒÝ":-27,"Ò‘":-9,"Ò’":6,"Ò“":-9,"Ò”":6,"Ó,":-23,"Ó.":-15,"ÓA":-11,"ÓJ":-13,"ÓT":-27,"ÓV":-12,"ÓW":-11,"ÓX":-31,"ÓY":-27,"ÓZ":-19,"Óx":-6,"Óz":-5,"ÓÀ":-11,"ÓÁ":-11,"ÓÂ":-11,"ÓÃ":-11,"ÓÄ":-11,"ÓÅ":-11,"ÓÆ":-14,"ÓÝ":-27,"Ó‘":-9,"Ó’":6,"Ó“":-9,"Ó”":6,"Ô,":-23,"Ô.":-15,"ÔA":-11,"ÔJ":-13,"ÔT":-27,"ÔV":-12,"ÔW":-11,"ÔX":-31,"ÔY":-27,"ÔZ":-19,"Ôx":-6,"Ôz":-5,"ÔÀ":-11,"ÔÁ":-11,"ÔÂ":-11,"ÔÃ":-11,"ÔÄ":-11,"ÔÅ":-11,"ÔÆ":-14,"ÔÝ":-27,"Ô‘":-9,"Ô’":6,"Ô“":-9,"Ô”":6,"Õ,":-23,"Õ.":-15,"ÕA":-11,"ÕJ":-13,"ÕT":-27,"ÕV":-12,"ÕW":-11,"ÕX":-31,"ÕY":-27,"ÕZ":-19,"Õx":-6,"Õz":-5,"ÕÀ":-11,"ÕÁ":-11,"ÕÂ":-11,"ÕÃ":-11,"ÕÄ":-11,"ÕÅ":-11,"ÕÆ":-14,"ÕÝ":-27,"Õ‘":-9,"Õ’":6,"Õ“":-9,"Õ”":6,"Ö,":-23,"Ö.":-15,"ÖA":-11,"ÖJ":-13,"ÖT":-27,"ÖV":-12,"ÖW":-11,"ÖX":-31,"ÖY":-27,"ÖZ":-19,"Öx":-6,"Öz":-5,"ÖÀ":-11,"ÖÁ":-11,"ÖÂ":-11,"ÖÃ":-11,"ÖÄ":-11,"ÖÅ":-11,"ÖÆ":-14,"ÖÝ":-27,"Ö‘":-9,"Ö’":6,"Ö“":-9,"Ö”":6,"Ù,":-14,"Ù.":-18,"ÙA":-22,"ÙJ":-20,"ÙÀ":-22,"ÙÁ":-22,"ÙÂ":-22,"ÙÃ":-22,"ÙÄ":-22,"ÙÅ":-22,"ÙÆ":-17,"Ú,":-14,"Ú.":-18,"ÚA":-22,"ÚJ":-20,"ÚÀ":-22,"ÚÁ":-22,"ÚÂ":-22,"ÚÃ":-22,"ÚÄ":-22,"ÚÅ":-22,"ÚÆ":-17,"Û,":-14,"Û.":-18,"ÛA":-22,"ÛJ":-20,"ÛÀ":-22,"ÛÁ":-22,"ÛÂ":-22,"ÛÃ":-22,"ÛÄ":-22,"ÛÅ":-22,"ÛÆ":-17,"Ü,":-14,"Ü.":-18,"ÜA":-22,"ÜJ":-20,"ÜÀ":-22,"ÜÁ":-22,"ÜÂ":-22,"ÜÃ":-22,"ÜÄ":-22,"ÜÅ":-22,"ÜÆ":-17,"Ý,":-123,"Ý-":-57,"Ý.":-116,"Ý/":-61,"Ý:":-75,"Ý;":-67,"ÝA":-74,"ÝC":-33,"ÝG":-33,"ÝJ":-55,"ÝO":-32,"ÝQ":-32,"ÝS":-8,"ÝZ":-5,"Ýa":-65,"Ýc":-78,"Ýd":-64,"Ýe":-72,"Ýf":-30,"Ýg":-69,"Ýi":-16,"Ýj":-24,"Ým":-46,"Ýn":-46,"Ýo":-75,"Ýp":-46,"Ýq":-64,"Ýr":-46,"Ýs":-56,"Ýt":-21,"Ýu":-46,"Ýv":-34,"Ýw":-30,"Ýx":-34,"Ýy":-32,"Ýz":-49,"ÝÀ":-74,"ÝÁ":-74,"ÝÂ":-74,"ÝÃ":-74,"ÝÄ":-74,"ÝÅ":-74,"ÝÆ":-86,"ÝÇ":-33,"ÝÒ":-32,"ÝÓ":-32,"ÝÔ":-32,"ÝÕ":-32,"ÝÖ":-32,"Ýß":-35,"Ýà":-65,"Ýá":-65,"Ýâ":-65,"Ýã":-65,"Ýä":-65,"Ýå":-65,"Ýæ":-65,"Ýç":-78,"Ýè":-72,"Ýé":-72,"Ýê":-72,"Ýë":-72,"Ýì":20,"Ýí":-34,"Ýî":23,"Ýï":25,"Ýñ":-46,"Ýò":-75,"Ýó":-75,"Ýô":-75,"Ýõ":-75,"Ýö":-75,"Ýø":-75,"Ýù":-46,"Ýú":-46,"Ýû":-46,"Ýü":-46,"Ýý":-32,"Ýÿ":-32,"Ý’":13,"Ý”":13,"àf":-6,"àt":-9,"àv":-17,"àw":-7,"àx":-9,"ày":-19,"àý":-19,"àÿ":-19,"à‘":-18,"à“":-18,"áf":-6,"át":-9,"áv":-17,"áw":-7,"áx":-9,"áy":-19,"áý":-19,"áÿ":-19,"á‘":-18,"á“":-18,"âf":-6,"ât":-9,"âv":-17,"âw":-7,"âx":-9,"ây":-19,"âý":-19,"âÿ":-19,"â‘":-18,"â“":-18,"ãf":-6,"ãt":-9,"ãv":-17,"ãw":-7,"ãx":-9,"ãy":-19,"ãý":-19,"ãÿ":-19,"ã‘":-18,"ã“":-18,"äf":-6,"ät":-9,"äv":-17,"äw":-7,"äx":-9,"äy":-19,"äý":-19,"äÿ":-19,"ä‘":-18,"ä“":-18,"åf":-6,"åt":-9,"åv":-17,"åw":-7,"åx":-9,"åy":-19,"åý":-19,"åÿ":-19,"å‘":-18,"å“":-18,"èf":-9,"èt":-5,"èv":-5,"èw":-5,"èx":-15,"èy":-6,"èz":-10,"èý":-6,"èÿ":-6,"è‘":-9,"è“":-9,"éf":-9,"ét":-5,"év":-5,"éw":-5,"éx":-15,"éy":-6,"éz":-10,"éý":-6,"éÿ":-6,"é‘":-9,"é“":-9,"êf":-9,"êt":-5,"êv":-5,"êw":-5,"êx":-15,"êy":-6,"êz":-10,"êý":-6,"êÿ":-6,"ê‘":-9,"ê“":-9,"ëf":-9,"ët":-5,"ëv":-5,"ëw":-5,"ëx":-15,"ëy":-6,"ëz":-10,"ëý":-6,"ëÿ":-6,"ë‘":-9,"ë“":-9,"îb":6,"îh":6,"îk":6,"îl":6,"îþ":6,"î‘":19,"î’":40,"î“":19,"î”":40,"ïb":6,"ïh":6,"ïk":6,"ïl":6,"ïþ":6,"ï‘":19,"ï’":24,"ï“":19,"ï”":24,"ñf":-6,"ñt":-9,"ñv":-17,"ñw":-7,"ñx":-9,"ñy":-19,"ñý":-19,"ñÿ":-19,"ñ‘":-18,"ñ“":-18,"ò,":-18,"òv":-4,"òw":-4,"òx":-20,"òy":-5,"òz":-13,"òý":-5,"òÿ":-5,"ó,":-18,"óv":-4,"ów":-4,"óx":-20,"óy":-5,"óz":-13,"óý":-5,"óÿ":-5,"ô,":-18,"ôv":-4,"ôw":-4,"ôx":-20,"ôy":-5,"ôz":-13,"ôý":-5,"ôÿ":-5,"õ,":-18,"õv":-4,"õw":-4,"õx":-20,"õy":-5,"õz":-13,"õý":-5,"õÿ":-5,"ö,":-18,"öv":-4,"öw":-4,"öx":-20,"öy":-5,"öz":-13,"öý":-5,"öÿ":-5,"ý,":-71,"ý-":-22,"ý.":-66,"ýa":-15,"ýc":-13,"ýd":-12,"ýe":-12,"ýf":5,"ýg":-13,"ýo":-12,"ýq":-12,"ýs":-9,"ýt":5,"ýv":6,"ýw":4,"ýy":5,"ýz":-8,"ýà":-15,"ýá":-15,"ýâ":-15,"ýã":-15,"ýä":-15,"ýå":-15,"ýæ":-15,"ýç":-13,"ýè":-12,"ýé":-12,"ýê":-12,"ýë":-12,"ýò":-12,"ýó":-12,"ýô":-12,"ýõ":-12,"ýö":-12,"ýø":-12,"ýý":5,"ýÿ":5,"ý‘":7,"ý’":30,"ý“":7,"ý”":30,"ÿ,":-71,"ÿ-":-22,"ÿ.":-66,"ÿa":-15,"ÿc":-13,"ÿd":-12,"ÿe":-12,"ÿf":5,"ÿg":-13,"ÿo":-12,"ÿq":-12,"ÿs":-9,"ÿt":5,"ÿv":6,"ÿw":4,"ÿy":5,"ÿz":-8,"ÿà":-15,"ÿá":-15,"ÿâ":-15,"ÿã":-15,"ÿä":-15,"ÿå":-15,"ÿæ":-15,"ÿç":-13,"ÿè":-12,"ÿé":-12,"ÿê":-12,"ÿë":-12,"ÿò":-12,"ÿó":-12,"ÿô":-12,"ÿõ":-12,"ÿö":-12,"ÿø":-12,"ÿý":5,"ÿÿ":5,"ÿ‘":7,"ÿ’":30,"ÿ“":7,"ÿ”":30},{"Q)":15,"Q,":46,"Q/":54,"Q;":29,"QJ":16,"QT":-20,"QV":-17,"QW":-7,"QX":2,"QY":-29,"Q]":17,"Qg":27,"Qj":39,"Qx":11,"Q}":15,"QÆ":26,"QÝ":-29,"Q’":6,"Q”":6,"V'":16,"V,":-77,"V-":-42,"V.":-95,"V/":-47,"V:":-29,"V;":-42,"V?":12,"VA":-53,"VC":-16,"VG":-17,"VJ":-39,"VO":-20,"VQ":-20,"VS":-8,"VV":4,"Va":-55,"Vc":-53,"Vd":-45,"Ve":-51,"Vf":-12,"Vg":-49,"Vm":-28,"Vn":-28,"Vo":-46,"Vp":-28,"Vq":-45,"Vr":-28,"Vs":-42,"Vt":-20,"Vu":-28,"Vv":-27,"Vx":-32,"Vy":-14,"Vz":-38,"VÀ":-53,"VÁ":-53,"VÂ":-53,"VÃ":-53,"VÄ":-53,"VÅ":-53,"VÆ":-71,"VÇ":-16,"VÒ":-20,"VÓ":-20,"VÔ":-20,"VÕ":-20,"VÖ":-20,"Và":-55,"Vá":-55,"Vâ":-55,"Vã":-55,"Vä":-55,"Vå":-55,"Væ":-55,"Vç":-53,"Vè":-51,"Vé":-51,"Vê":-51,"Vë":-51,"Vì":29,"Vî":37,"Vï":46,"Vñ":-28,"Vò":-46,"Vó":-46,"Vô":-46,"Võ":-46,"Vö":-46,"Vø":-46,"Vù":-28,"Vú":-28,"Vû":-28,"Vü":-28,"Vý":-14,"Vÿ":-14,"V’":18,"V”":18,"f'":23,"f)":15,"f,":-49,"f-":-22,"f.":-54,"fa":-16,"fc":-17,"fd":-18,"fe":-18,"fg":-20,"fo":-15,"fq":-18,"fs":-10,"fv":10,"fw":3,"fy":5,"fz":-10,"f®":15,"fà":-16,"fá":-16,"fâ":-16,"fã":-16,"fä":-16,"få":-16,"fæ":-16,"fç":-17,"fè":-18,"fé":-18,"fê":-18,"fë":-18,"fò":-15,"fó":-15,"fô":-15,"fõ":-15,"fö":-15,"fø":-15,"fý":5,"fÿ":5,"f‘":17,"f’":44,"f“":17,"f”":44,"n'":-32,"nf":-5,"nt":-9,"nv":-14,"nw":-8,"nx":-9,"ny":-18,"ný":-18,"nÿ":-18,"n‘":-15,"n“":-15,"qg":4,"ßf":-14,"ßt":-15,"ßu":-6,"ßv":-15,"ßw":-17,"ßx":-15,"ßy":-21,"ßz":-10,"ßý":-21,"ßÿ":-21,"ß‘":-22,"ß“":-22,"v,":-59,"v-":-20,"v.":-71,"vA":-40,"va":-15,"vc":-12,"vd":-10,"ve":-9,"vf":7,"vg":-11,"vo":-9,"vq":-10,"vs":-3,"vt":5,"vv":7,"vw":7,"vy":7,"vz":-9,"và":-15,"vá":-15,"vâ":-15,"vã":-15,"vä":-15,"vå":-15,"væ":-15,"vç":-12,"vè":-9,"vé":-9,"vê":-9,"vë":-9,"vò":-9,"vó":-9,"vô":-9,"võ":-9,"vö":-9,"vø":-9,"vý":7,"vÿ":7,"v‘":12,"v’":29,"v“":12,"v”":29,"w,":-73,"w-":-15,"w.":-58,"wa":-11,"wc":-9,"wd":-7,"we":-8,"wf":3,"wg":-8,"wo":-8,"wq":-7,"ws":-6,"wt":3,"wv":7,"ww":6,"wy":7,"wz":-6,"wà":-11,"wá":-11,"wâ":-11,"wã":-11,"wä":-11,"wå":-11,"wæ":-11,"wç":-9,"wè":-8,"wé":-8,"wê":-8,"wë":-8,"wò":-8,"wó":-8,"wô":-8,"wõ":-8,"wö":-8,"wø":-8,"wý":7,"wÿ":7,"w‘":12,"w’":29,"w“":12,"w”":29,",A":15,",J":22,",T":-49,",V":-49,",W":-47,",X":9,",Y":-61,",t":-15,",À":15,",Á":15,",Â":15,",Ã":15,",Ä":15,",Å":15,",Ý":-61,",–":-20,",‘":-71,",’":-59,",“":-71,",”":-59,".'":-79,".-":-40,".A":7,".C":-17,".G":-17,".J":17,".O":-12,".Q":-12,".T":-68,".U":-7,".V":-66,".W":-47,".Y":-90,".f":-20,".g":12,".t":-24,".v":-34,".w":-32,".y":-34,".z":5,".À":7,".Á":7,".Â":7,".Ã":7,".Ä":7,".Å":7,".Ç":-17,".Ò":-12,".Ó":-12,".Ô":-12,".Õ":-12,".Ö":-12,".Ù":-7,".Ú":-7,".Û":-7,".Ü":-7,".Ý":-90,".ý":-34,".ÿ":-34,".‘":-90,".’":-73,".“":-90,".”":-73,"/A":-44,"/J":-24,"/a":-24,"/c":-24,"/e":-24,"/g":-22,"/o":-22,"/s":-27,"/²":-142,"/³":-147,"/¹":-115,"/À":-44,"/Á":-44,"/Â":-44,"/Ã":-44,"/Ä":-44,"/Å":-44,"/Æ":-42,"/à":-24,"/á":-24,"/â":-24,"/ã":-24,"/ä":-24,"/å":-24,"/æ":-24,"/ç":-24,"/è":-24,"/é":-24,"/ê":-24,"/ë":-24,"/ò":-22,"/ó":-22,"/ô":-22,"/õ":-22,"/ö":-22,"/ø":-22,"-'":-63,"-.":-51,"-A":-32,"-T":-71,"-V":-47,"-W":-16,"-X":-27,"-Y":-42,"-Z":-20,"-f":-7,"-v":-7,"-x":-17,"-y":-10,"-Ý":-42,"-ý":-10,"-ÿ":-10,"(O":-16,"(g":5,"(j":24,"(o":-16,"'-":-63,"'.":-79,"'A":-79,"'O":-16,"'V":16,"'W":16,"'Y":16,"'a":-27,"'o":-40,"++":12,"²/":-145,"³/":-165,"¹/":-141,"¿T":-44,"¿V":-43,"¿Y":-55,"¿f":-22,"¿j":15,"¿v":-27,"¿w":-29,"¿y":-29,"¿Ý":-55,"¿ý":-29,"¿ÿ":-29,"í‘":6,"í’":39,"í“":6,"í”":39,"A'":-79,"A,":12,"A-":-12,"A.":7,"A?":-41,"AC":-12,"AG":-12,"AJ":14,"AO":-14,"AQ":-14,"AT":-78,"AU":-17,"AV":-51,"AW":-39,"AY":-82,"At":-31,"Av":-24,"Ay":-30,"AÇ":-12,"AÒ":-14,"AÓ":-14,"AÔ":-14,"AÕ":-14,"AÖ":-14,"AÙ":-17,"AÚ":-17,"AÛ":-17,"AÜ":-17,"AÝ":-82,"Aý":-30,"Aÿ":-30,"A‘":-81,"A’":-46,"A“":-81,"A”":-46,"B,":-12,"BA":-10,"BT":-22,"BV":-17,"BW":-15,"BX":-24,"BY":-34,"BZ":-10,"Bf":-10,"Bt":-10,"Bv":-10,"Bx":-12,"By":-10,"BÀ":-10,"BÁ":-10,"BÂ":-10,"BÃ":-10,"BÄ":-10,"BÅ":-10,"BÝ":-34,"Bý":-10,"Bÿ":-10,"B‘":-12,"B“":-12,"C,":-3,"C-":-55,"CG":-7,"CJ":7,"CO":-7,"CQ":-7,"CT":5,"CÆ":10,"CÒ":-7,"CÓ":-7,"CÔ":-7,"CÕ":-7,"CÖ":-7,"C‘":7,"C’":30,"C“":7,"C”":30,"Ç,":-3,"ÇG":-7,"ÇJ":7,"ÇO":-7,"ÇQ":-7,"ÇT":5,"ÇÆ":10,"ÇÒ":-7,"ÇÓ":-7,"ÇÔ":-7,"ÇÕ":-7,"ÇÖ":-7,"Ç‘":7,"Ç’":30,"Ç“":7,"Ç”":30,"D,":-22,"D.":-22,"DA":-23,"DJ":-12,"DT":-14,"DV":-15,"DW":-10,"DX":-25,"DY":-27,"DZ":-12,"DÀ":-23,"DÁ":-23,"DÂ":-23,"DÃ":-23,"DÄ":-23,"DÅ":-23,"DÆ":-20,"DÝ":-27,"D‘":-6,"D’":5,"D“":-6,"D”":5,"Ð,":-22,"Ð.":-22,"ÐA":-23,"ÐJ":-12,"ÐT":-14,"ÐV":-15,"ÐW":-10,"ÐX":-25,"ÐY":-27,"ÐZ":-12,"ÐÀ":-23,"ÐÁ":-23,"ÐÂ":-23,"ÐÃ":-23,"ÐÄ":-23,"ÐÅ":-23,"ÐÆ":-20,"ÐÝ":-27,"Ð‘":-6,"Ð’":5,"Ð“":-6,"Ð”":5,"F,":-85,"F.":-79,"F/":-27,"F?":12,"FA":-57,"FC":-7,"FG":-7,"FJ":-53,"FO":-7,"FQ":-7,"FS":-13,"FX":-12,"FZ":-7,"Fa":-27,"Fc":-12,"Fd":-10,"Fe":-15,"Fn":-23,"Fo":-12,"Fq":-10,"Fs":-12,"FÀ":-57,"FÁ":-57,"FÂ":-57,"FÃ":-57,"FÄ":-57,"FÅ":-57,"FÆ":-51,"FÇ":-7,"FÒ":-7,"FÓ":-7,"FÔ":-7,"FÕ":-7,"FÖ":-7,"Fà":-27,"Fá":-27,"Fâ":-27,"Fã":-27,"Fä":-27,"Få":-27,"Fæ":-27,"Fç":-12,"Fè":-15,"Fé":-15,"Fê":-15,"Fë":-15,"Fî":17,"Fï":17,"Fò":-12,"Fó":-12,"Fô":-12,"Fõ":-12,"Fö":-12,"Fø":-12,"F’":5,"F”":5,"GT":-5,"GV":-5,"GW":-4,"GY":-15,"Gv":-13,"Gw":-11,"Gx":-10,"Gy":-15,"GÆ":-5,"GÝ":-15,"Gý":-15,"Gÿ":-15,"Hî":7,"Hï":7,"H‘":-5,"H“":-5,"J,":-10,"J.":-7,"JA":-22,"JX":-10,"JÀ":-22,"JÁ":-22,"JÂ":-22,"JÃ":-22,"JÄ":-22,"JÅ":-22,"JÆ":-17,"K-":-20,"KC":-37,"KG":-39,"KO":-44,"KQ":-44,"KS":-7,"KU":-13,"KW":-11,"Ka":-13,"Kc":-20,"Kd":-18,"Ke":-18,"Kf":-17,"Km":-11,"Kn":-11,"Ko":-16,"Kp":-11,"Kq":-18,"Kr":-11,"Ks":-7,"Kt":-25,"Ku":-11,"Kv":-43,"Kw":-42,"Ky":-40,"KÇ":-37,"KÒ":-44,"KÓ":-44,"KÔ":-44,"KÕ":-44,"KÖ":-44,"KÙ":-13,"KÚ":-13,"KÛ":-13,"KÜ":-13,"Kà":-13,"Ká":-13,"Kâ":-13,"Kã":-13,"Kä":-13,"Kå":-13,"Kæ":-13,"Kç":-20,"Kè":-18,"Ké":-18,"Kê":-18,"Kë":-18,"Kì":13,"Kî":22,"Kï":23,"Kñ":-11,"Kò":-16,"Kó":-16,"Kô":-16,"Kõ":-16,"Kö":-16,"Kø":-16,"Kù":-11,"Kú":-11,"Kû":-11,"Kü":-11,"Ký":-40,"Kÿ":-40,"L'":-110,"L,":15,"L-":-63,"L.":23,"L?":-20,"LC":-12,"LG":-20,"LJ":16,"LO":-18,"LQ":-18,"LT":-73,"LU":-24,"LV":-78,"LW":-56,"LY":-88,"Lf":-14,"Lo":-16,"Lt":-19,"Lv":-38,"Lw":-35,"Ly":-41,"L·":-31,"LÆ":24,"LÇ":-12,"LÒ":-18,"LÓ":-18,"LÔ":-18,"LÕ":-18,"LÖ":-18,"LÙ":-24,"LÚ":-24,"LÛ":-24,"LÜ":-24,"LÝ":-88,"Lý":-41,"Lÿ":-41,"L‘":-93,"L’":-72,"L“":-93,"L”":-72,"O'":-16,"O)":-16,"O,":-22,"O.":-15,"OA":-14,"OJ":-10,"OT":-22,"OV":-17,"OW":-12,"OX":-34,"OY":-32,"OZ":-16,"Ox":-7,"Oz":-5,"OÀ":-14,"OÁ":-14,"OÂ":-14,"OÃ":-14,"OÄ":-14,"OÅ":-14,"OÆ":-22,"OÝ":-32,"O‘":-7,"O’":7,"O“":-7,"O”":7,"P,":-110,"P-":-22,"P.":-110,"P/":-42,"PA":-68,"PJ":-60,"PT":-3,"PV":-5,"PX":-22,"PY":-6,"PZ":-14,"Pa":-20,"Pc":-19,"Pd":-12,"Pe":-16,"Pf":7,"Po":-16,"Pq":-12,"Ps":-11,"Pt":7,"Py":7,"PÀ":-68,"PÁ":-68,"PÂ":-68,"PÃ":-68,"PÄ":-68,"PÅ":-68,"PÆ":-68,"PÝ":-6,"Pà":-20,"Pá":-20,"Pâ":-20,"Pã":-20,"Pä":-20,"På":-20,"Pæ":-20,"Pç":-19,"Pè":-16,"Pé":-16,"Pê":-16,"Pë":-16,"Pò":-16,"Pó":-16,"Pô":-16,"Põ":-16,"Pö":-16,"Pø":-16,"Pý":7,"Pÿ":7,"P‘":10,"P’":22,"P“":10,"P”":22,"R.":5,"RC":-7,"RG":-8,"RO":-10,"RQ":-10,"RS":-10,"RT":-10,"RU":-23,"RV":-12,"RW":-7,"RY":-19,"Re":-15,"Ro":-18,"Rv":-9,"Rw":-11,"Ry":-11,"RÇ":-7,"RÒ":-10,"RÓ":-10,"RÔ":-10,"RÕ":-10,"RÖ":-10,"RÝ":-19,"Rè":-15,"Ré":-15,"Rê":-15,"Rë":-15,"Rò":-18,"Ró":-18,"Rô":-18,"Rõ":-18,"Rö":-18,"Rø":-18,"Rý":-11,"Rÿ":-11,"S-":-4,"SA":-12,"SJ":-4,"ST":-10,"SV":-10,"SW":-7,"SX":-9,"SY":-19,"Sv":-11,"Sw":-9,"Sy":-14,"SÀ":-12,"SÁ":-12,"SÂ":-12,"SÃ":-12,"SÄ":-12,"SÅ":-12,"SÆ":-5,"SÝ":-19,"Sý":-14,"Sÿ":-14,"S‘":-10,"S’":-10,"S“":-10,"S”":-10,"T,":-93,"T-":-68,"T.":-83,"T/":-59,"T:":-42,"T;":-42,"TA":-79,"TC":-22,"TG":-27,"TJ":-36,"TO":-27,"TQ":-27,"TS":-5,"TT":12,"Ta":-76,"Tc":-87,"Td":-72,"Te":-86,"Tg":-72,"Tm":-51,"Tn":-51,"To":-86,"Tp":-51,"Tq":-72,"Tr":-51,"Ts":-66,"Tu":-51,"Tv":-31,"Tw":-29,"Tx":-37,"Ty":-31,"Tz":-52,"TÀ":-79,"TÁ":-79,"TÂ":-79,"TÃ":-79,"TÄ":-79,"TÅ":-79,"TÆ":-83,"TÇ":-22,"TÒ":-27,"TÓ":-27,"TÔ":-27,"TÕ":-27,"TÖ":-27,"Tà":-76,"Tá":-76,"Tâ":-76,"Tã":-76,"Tä":-76,"Tå":-76,"Tæ":-76,"Tç":-87,"Tè":-86,"Té":-86,"Tê":-86,"Të":-86,"Tì":17,"Tî":37,"Tï":44,"Tñ":-51,"Tò":-86,"Tó":-86,"Tô":-86,"Tõ":-86,"Tö":-86,"Tø":-86,"Tù":-51,"Tú":-51,"Tû":-51,"Tü":-51,"Tý":-31,"Tÿ":-31,"T‘":12,"T’":18,"T“":12,"T”":18,"W'":16,"W,":-84,"W-":-32,"W.":-86,"W;":-54,"W?":12,"WA":-47,"WC":-13,"WG":-12,"WJ":-42,"WO":-12,"WQ":-12,"WS":-5,"WX":-7,"Wa":-36,"Wc":-38,"Wd":-33,"We":-36,"Wg":-28,"Wm":-27,"Wn":-27,"Wo":-39,"Wp":-27,"Wq":-33,"Wr":-27,"Ws":-33,"Wu":-27,"Wv":-12,"Wy":-18,"WÀ":-47,"WÁ":-47,"WÂ":-47,"WÃ":-47,"WÄ":-47,"WÅ":-47,"WÆ":-61,"WÇ":-13,"WÒ":-12,"WÓ":-12,"WÔ":-12,"WÕ":-12,"WÖ":-12,"Wà":-36,"Wá":-36,"Wâ":-36,"Wã":-36,"Wä":-36,"Wå":-36,"Wæ":-36,"Wç":-38,"Wè":-36,"Wé":-36,"Wê":-36,"Wë":-36,"Wì":5,"Wî":27,"Wï":29,"Wñ":-27,"Wò":-39,"Wó":-39,"Wô":-39,"Wõ":-39,"Wö":-39,"Wø":-39,"Wù":-27,"Wú":-27,"Wû":-27,"Wü":-27,"Wý":-18,"Wÿ":-18,"W’":6,"W”":6,"X-":-32,"XC":-34,"XG":-37,"XO":-34,"XQ":-34,"XS":-10,"Xd":-22,"Xe":-19,"Xg":-3,"Xo":-19,"Xq":-22,"Xt":-22,"Xu":-17,"Xv":-36,"Xw":-32,"Xy":-32,"XÇ":-34,"XÒ":-34,"XÓ":-34,"XÔ":-34,"XÕ":-34,"XÖ":-34,"Xè":-19,"Xé":-19,"Xê":-19,"Xë":-19,"Xò":-19,"Xó":-19,"Xô":-19,"Xõ":-19,"Xö":-19,"Xø":-19,"Xý":-32,"Xÿ":-32,"X’":-10,"X”":-10,"Y'":16,"Y,":-116,"Y-":-54,"Y.":-112,"Y/":-66,"Y:":-69,"Y;":-66,"Y?":12,"YA":-85,"YC":-39,"YG":-39,"YJ":-56,"YO":-38,"YQ":-38,"YS":-15,"YZ":-5,"Ya":-68,"Yc":-85,"Yd":-71,"Ye":-80,"Yf":-32,"Yg":-75,"Yi":-15,"Yj":-22,"Ym":-49,"Yn":-49,"Yo":-81,"Yp":-49,"Yq":-71,"Yr":-49,"Ys":-62,"Yt":-24,"Yu":-49,"Yv":-34,"Yw":-31,"Yx":-42,"Yy":-37,"Yz":-55,"YÀ":-85,"YÁ":-85,"YÂ":-85,"YÃ":-85,"YÄ":-85,"YÅ":-85,"YÆ":-92,"YÇ":-39,"YÒ":-38,"YÓ":-38,"YÔ":-38,"YÕ":-38,"YÖ":-38,"Yß":-37,"Yà":-68,"Yá":-68,"Yâ":-68,"Yã":-68,"Yä":-68,"Yå":-68,"Yæ":-68,"Yç":-85,"Yè":-80,"Yé":-80,"Yê":-80,"Yë":-80,"Yì":20,"Yí":-34,"Yî":29,"Yï":37,"Yñ":-49,"Yò":-81,"Yó":-81,"Yô":-81,"Yõ":-81,"Yö":-81,"Yø":-81,"Yù":-49,"Yú":-49,"Yû":-49,"Yü":-49,"Yý":-37,"Yÿ":-37,"Y’":10,"Y”":10,"Z-":-17,"Z?":12,"ZA":-7,"ZC":-17,"ZG":-15,"ZO":-15,"ZQ":-15,"ZW":-6,"ZY":-6,"Za":-5,"Zc":-7,"Zd":-7,"Ze":-15,"Zo":-13,"Zq":-7,"Zs":-5,"Zv":-21,"Zw":-19,"Zy":-17,"ZÀ":-7,"ZÁ":-7,"ZÂ":-7,"ZÃ":-7,"ZÄ":-7,"ZÅ":-7,"ZÇ":-17,"ZÒ":-15,"ZÓ":-15,"ZÔ":-15,"ZÕ":-15,"ZÖ":-15,"ZÝ":-6,"Zà":-5,"Zá":-5,"Zâ":-5,"Zã":-5,"Zä":-5,"Zå":-5,"Zæ":-5,"Zç":-7,"Zè":-15,"Zé":-15,"Zê":-15,"Zë":-15,"Zì":6,"Zî":22,"Zï":22,"Zò":-13,"Zó":-13,"Zô":-13,"Zõ":-13,"Zö":-13,"Zø":-13,"Zý":-17,"Zÿ":-17,"Z’":16,"Z”":16,"ca":-6,"co":-5,"cà":-6,"cá":-6,"câ":-6,"cã":-6,"cä":-6,"cå":-6,"cæ":-6,"cò":-5,"có":-5,"cô":-5,"cõ":-5,"cö":-5,"cø":-5,"ça":-6,"ço":-5,"çà":-6,"çá":-6,"çâ":-6,"çã":-6,"çä":-6,"çå":-6,"çæ":-6,"çò":-5,"çó":-5,"çô":-5,"çõ":-5,"çö":-5,"çø":-5,"g,":15,"g/":29,"ga":-17,"gc":-7,"gd":-11,"ge":-11,"gg":8,"go":-10,"gq":-11,"gt":-7,"gà":-17,"gá":-17,"gâ":-17,"gã":-17,"gä":-17,"gå":-17,"gæ":-17,"gç":-7,"gè":-11,"gé":-11,"gê":-11,"gë":-11,"gò":-10,"gó":-10,"gô":-10,"gõ":-10,"gö":-10,"gø":-10,"g‘":9,"g’":34,"g“":9,"g”":34,"k-":-34,"ka":-12,"kc":-22,"kd":-22,"ke":-27,"ko":-28,"kq":-22,"ks":-8,"kt":-5,"ku":-12,"kà":-12,"ká":-12,"kâ":-12,"kã":-12,"kä":-12,"kå":-12,"kæ":-12,"kç":-22,"kè":-27,"ké":-27,"kê":-27,"kë":-27,"kò":-28,"kó":-28,"kô":-28,"kõ":-28,"kö":-28,"kø":-28,"o'":-40,"o)":-16,"o,":-15,"oV":-51,"oW":-20,"ov":-3,"ow":-2,"ox":-20,"oy":-6,"oz":-10,"oý":-6,"oÿ":-6,"r,":-73,"r-":-20,"r.":-88,"ra":-22,"rc":-14,"rd":-12,"re":-11,"rg":-12,"ro":-12,"rq":-12,"rs":-12,"rv":9,"rw":6,"ry":4,"rà":-22,"rá":-22,"râ":-22,"rã":-22,"rä":-22,"rå":-22,"ræ":-22,"rç":-14,"rè":-11,"ré":-11,"rê":-11,"rë":-11,"rò":-12,"ró":-12,"rô":-12,"rõ":-12,"rö":-12,"rø":-12,"rý":4,"rÿ":4,"r‘":12,"r’":37,"r“":12,"r”":37,"sf":-8,"st":-11,"sv":-14,"sw":-5,"sx":-11,"sy":-15,"sz":-7,"sý":-15,"sÿ":-15,"s‘":-10,"s“":-10,"x-":-32,"xa":-16,"xc":-23,"xd":-19,"xe":-24,"xo":-23,"xq":-19,"xs":-7,"xu":-10,"xà":-16,"xá":-16,"xâ":-16,"xã":-16,"xä":-16,"xå":-16,"xæ":-16,"xç":-23,"xè":-24,"xé":-24,"xê":-24,"xë":-24,"xò":-23,"xó":-23,"xô":-23,"xõ":-23,"xö":-23,"xø":-23,"x’":17,"x”":17,"‘,":-83,"‘.":-85,"‘A":-80,"‘G":-15,"‘J":-42,"‘O":-12,"‘Q":-12,"‘T":37,"‘V":22,"‘W":7,"‘Y":19,"‘a":-28,"‘c":-42,"‘d":-39,"‘e":-39,"‘g":-32,"‘o":-34,"‘q":-39,"‘s":-20,"‘À":-80,"‘Á":-80,"‘Â":-80,"‘Ã":-80,"‘Ä":-80,"‘Å":-80,"‘Æ":-93,"‘Ò":-12,"‘Ó":-12,"‘Ô":-12,"‘Õ":-12,"‘Ö":-12,"‘Ý":19,"‘à":-28,"‘á":-28,"‘â":-28,"‘ã":-28,"‘ä":-28,"‘å":-28,"‘æ":-28,"‘ç":-42,"‘è":-39,"‘é":-39,"‘ê":-39,"‘ë":-39,"‘ì":44,"‘î":44,"‘ï":47,"‘ò":-34,"‘ó":-34,"‘ô":-34,"‘õ":-34,"‘ö":-34,"‘ø":-34,"’,":-95,"’.":-90,"’A":-119,"’C":-32,"’G":-34,"’J":-56,"’O":-34,"’Q":-34,"’T":12,"’Y":10,"’a":-62,"’c":-71,"’d":-66,"’e":-66,"’g":-61,"’i":2,"’j":-5,"’m":-17,"’n":-17,"’o":-71,"’p":-17,"’q":-66,"’r":-17,"’s":-56,"’u":-17,"’v":-10,"’w":-10,"’x":-20,"’y":-15,"’z":-17,"’À":-119,"’Á":-119,"’Â":-119,"’Ã":-119,"’Ä":-119,"’Å":-119,"’Æ":-132,"’Ç":-32,"’Ò":-34,"’Ó":-34,"’Ô":-34,"’Õ":-34,"’Ö":-34,"’Ý":10,"’à":-62,"’á":-62,"’â":-62,"’ã":-62,"’ä":-62,"’å":-62,"’æ":-62,"’ç":-71,"’è":-66,"’é":-66,"’ê":-66,"’ë":-66,"’ì":30,"’î":27,"’ï":32,"’ñ":-17,"’ò":-71,"’ó":-71,"’ô":-71,"’õ":-71,"’ö":-71,"’ø":-71,"’ù":-17,"’ú":-17,"’û":-17,"’ü":-17,"’ý":-15,"’ÿ":-15,"“,":-83,"“.":-85,"“A":-80,"“G":-15,"“J":-42,"“O":-12,"“Q":-12,"“T":37,"“V":22,"“W":7,"“Y":19,"“a":-28,"“c":-42,"“d":-39,"“e":-39,"“g":-32,"“o":-34,"“q":-39,"“s":-20,"“À":-80,"“Á":-80,"“Â":-80,"“Ã":-80,"“Ä":-80,"“Å":-80,"“Æ":-93,"“Ò":-12,"“Ó":-12,"“Ô":-12,"“Õ":-12,"“Ö":-12,"“Ý":19,"“à":-28,"“á":-28,"“â":-28,"“ã":-28,"“ä":-28,"“å":-28,"“æ":-28,"“ç":-42,"“è":-39,"“é":-39,"“ê":-39,"“ë":-39,"“ì":44,"“î":44,"“ï":47,"“ò":-34,"“ó":-34,"“ô":-34,"“õ":-34,"“ö":-34,"“ø":-34,"”,":-95,"”.":-90,"”A":-119,"”C":-32,"”G":-34,"”J":-56,"”O":-34,"”Q":-34,"”T":12,"”Y":10,"”a":-62,"”c":-71,"”d":-66,"”e":-66,"”g":-61,"”i":2,"”j":-5,"”m":-17,"”n":-17,"”o":-71,"”p":-17,"”q":-66,"”r":-17,"”s":-56,"”u":-17,"”v":-10,"”w":-10,"”x":-20,"”y":-15,"”z":-17,"”À":-119,"”Á":-119,"”Â":-119,"”Ã":-119,"”Ä":-119,"”Å":-119,"”Æ":-132,"”Ç":-32,"”Ò":-34,"”Ó":-34,"”Ô":-34,"”Õ":-34,"”Ö":-34,"”Ý":10,"”à":-62,"”á":-62,"”â":-62,"”ã":-62,"”ä":-62,"”å":-62,"”æ":-62,"”ç":-71,"”è":-66,"”é":-66,"”ê":-66,"”ë":-66,"”ì":30,"”î":27,"”ï":32,"”ñ":-17,"”ò":-71,"”ó":-71,"”ô":-71,"”õ":-71,"”ö":-71,"”ø":-71,"”ù":-17,"”ú":-17,"”û":-17,"”ü":-17,"”ý":-15,"”ÿ":-15,"Ý,":-116,"Ý-":-54,"Ý.":-112,"Ý/":-66,"Ý:":-69,"Ý;":-66,"ÝA":-85,"ÝC":-39,"ÝG":-39,"ÝJ":-56,"ÝO":-38,"ÝQ":-38,"ÝS":-15,"ÝZ":-5,"Ýa":-68,"Ýc":-85,"Ýd":-71,"Ýe":-80,"Ýf":-32,"Ýg":-75,"Ýi":-15,"Ýj":-22,"Ým":-49,"Ýn":-49,"Ýo":-81,"Ýp":-49,"Ýq":-71,"Ýr":-49,"Ýs":-62,"Ýt":-24,"Ýu":-49,"Ýv":-34,"Ýw":-31,"Ýx":-42,"Ýy":-37,"Ýz":-55,"ÝÀ":-85,"ÝÁ":-85,"ÝÂ":-85,"ÝÃ":-85,"ÝÄ":-85,"ÝÅ":-85,"ÝÆ":-92,"ÝÇ":-39,"ÝÒ":-38,"ÝÓ":-38,"ÝÔ":-38,"ÝÕ":-38,"ÝÖ":-38,"Ýß":-37,"Ýà":-68,"Ýá":-68,"Ýâ":-68,"Ýã":-68,"Ýä":-68,"Ýå":-68,"Ýæ":-68,"Ýç":-85,"Ýè":-80,"Ýé":-80,"Ýê":-80,"Ýë":-80,"Ýì":20,"Ýí":-34,"Ýî":29,"Ýï":37,"Ýñ":-49,"Ýò":-81,"Ýó":-81,"Ýô":-81,"Ýõ":-81,"Ýö":-81,"Ýø":-81,"Ýù":-49,"Ýú":-49,"Ýû":-49,"Ýü":-49,"Ýý":-37,"Ýÿ":-37,"Ý’":10,"Ý”":10,"ïb":10,"ïh":10,"ïk":10,"ïl":10,"ïþ":10,"ï‘":24,"ï’":32,"ï“":24,"ï”":32,"Æ-":-20,"ÆA":-12,"ÆC":-15,"ÆG":-15,"ÆO":-17,"ÆQ":-17,"ÆS":-10,"ÆZ":-5,"Æa":-13,"Æc":-12,"Æd":-15,"Æe":-17,"Æf":-27,"Æo":-17,"Æq":-15,"Æt":-15,"Æv":-26,"Æw":-20,"Æy":-26,"ÆÀ":-12,"ÆÁ":-12,"ÆÂ":-12,"ÆÃ":-12,"ÆÄ":-12,"ÆÅ":-12,"ÆÇ":-15,"ÆÒ":-17,"ÆÓ":-17,"ÆÔ":-17,"ÆÕ":-17,"ÆÖ":-17,"Æà":-13,"Æá":-13,"Æâ":-13,"Æã":-13,"Æä":-13,"Æå":-13,"Ææ":-13,"Æç":-12,"Æè":-17,"Æé":-17,"Æê":-17,"Æë":-17,"Æò":-17,"Æó":-17,"Æô":-17,"Æõ":-17,"Æö":-17,"Æø":-17,"Æý":-26,"Æÿ":-26,"E-":-20,"EA":-12,"EC":-15,"EG":-15,"EO":-17,"EQ":-17,"ES":-10,"EZ":-5,"Ea":-13,"Ec":-12,"Ed":-15,"Ee":-17,"Ef":-27,"Eo":-17,"Eq":-15,"Et":-15,"Ev":-26,"Ew":-20,"Ey":-26,"EÀ":-12,"EÁ":-12,"EÂ":-12,"EÃ":-12,"EÄ":-12,"EÅ":-12,"EÇ":-15,"EÒ":-17,"EÓ":-17,"EÔ":-17,"EÕ":-17,"EÖ":-17,"Eà":-13,"Eá":-13,"Eâ":-13,"Eã":-13,"Eä":-13,"Eå":-13,"Eæ":-13,"Eç":-12,"Eè":-17,"Eé":-17,"Eê":-17,"Eë":-17,"Eò":-17,"Eó":-17,"Eô":-17,"Eõ":-17,"Eö":-17,"Eø":-17,"Eý":-26,"Eÿ":-26,"U,":-12,"U.":-13,"UA":-27,"UJ":-16,"UÀ":-27,"UÁ":-27,"UÂ":-27,"UÃ":-27,"UÄ":-27,"UÅ":-27,"UÆ":-20,"af":-5,"at":-9,"av":-14,"aw":-8,"ax":-9,"ay":-18,"aý":-18,"aÿ":-18,"a‘":-15,"a“":-15,"æf":-7,"æt":-7,"æv":-5,"æw":-5,"æx":-16,"æy":-9,"æz":-10,"æý":-9,"æÿ":-9,"æ‘":-7,"æ“":-7,"bf":-5,"bs":-5,"bt":-4,"bv":-5,"bw":-5,"bx":-18,"by":-5,"bz":-12,"bý":-5,"bÿ":-5,"b‘":-12,"b“":-12,"ef":-7,"et":-7,"ev":-5,"ew":-5,"ex":-16,"ey":-9,"ez":-10,"eý":-9,"eÿ":-9,"e‘":-7,"e“":-7,"hf":-5,"ht":-9,"hv":-14,"hw":-8,"hx":-9,"hy":-18,"hý":-18,"hÿ":-18,"h‘":-15,"h“":-15,"i’":12,"i”":12,"j’":12,"j”":12,"mf":-5,"mt":-9,"mv":-14,"mw":-8,"mx":-9,"my":-18,"mý":-18,"mÿ":-18,"m‘":-15,"m“":-15,"ø,":-15,"øv":-3,"øw":-2,"øx":-20,"øy":-6,"øz":-10,"øý":-6,"øÿ":-6,"pf":-5,"ps":-5,"pt":-4,"pv":-5,"pw":-5,"px":-18,"py":-5,"pz":-12,"pý":-5,"pÿ":-5,"p‘":-12,"p“":-12,"þf":-5,"þs":-5,"þt":-4,"þv":-5,"þw":-5,"þx":-18,"þy":-5,"þz":-12,"þý":-5,"þÿ":-5,"þ‘":-12,"þ“":-12,"t-":-24,"ta":-10,"tc":-13,"td":-11,"te":-12,"to":-10,"tq":-11,"tt":-12,"tà":-10,"tá":-10,"tâ":-10,"tã":-10,"tä":-10,"tå":-10,"tæ":-10,"tç":-13,"tè":-12,"té":-12,"tê":-12,"të":-12,"tò":-10,"tó":-10,"tô":-10,"tõ":-10,"tö":-10,"tø":-10,"t‘":6,"t’":27,"t“":6,"t”":27,"y,":-61,"y-":-17,"y.":-61,"ya":-15,"yc":-13,"yd":-11,"ye":-11,"yf":5,"yg":-10,"yo":-11,"yq":-11,"ys":-6,"yt":5,"yv":7,"yw":6,"yy":5,"yz":-6,"yà":-15,"yá":-15,"yâ":-15,"yã":-15,"yä":-15,"yå":-15,"yæ":-15,"yç":-13,"yè":-11,"yé":-11,"yê":-11,"yë":-11,"yò":-11,"yó":-11,"yô":-11,"yõ":-11,"yö":-11,"yø":-11,"yý":5,"yÿ":5,"y‘":12,"y’":32,"y“":12,"y”":32,"za":-13,"zc":-18,"zd":-19,"ze":-19,"zf":-5,"zg":-5,"zo":-17,"zq":-19,"zs":-9,"zu":-5,"zv":-7,"zw":-8,"zy":-7,"zà":-13,"zá":-13,"zâ":-13,"zã":-13,"zä":-13,"zå":-13,"zæ":-13,"zç":-18,"zè":-19,"zé":-19,"zê":-19,"zë":-19,"zò":-17,"zó":-17,"zô":-17,"zõ":-17,"zö":-17,"zø":-17,"zý":-7,"zÿ":-7,"z’":12,"z”":12,"?’":12,"?”":12,"À,":12,"À-":-12,"À.":7,"À?":-41,"ÀC":-12,"ÀG":-12,"ÀJ":14,"ÀO":-14,"ÀQ":-14,"ÀT":-78,"ÀU":-17,"ÀV":-51,"ÀW":-39,"ÀY":-82,"Àt":-31,"Àv":-24,"Ày":-30,"ÀÇ":-12,"ÀÒ":-14,"ÀÓ":-14,"ÀÔ":-14,"ÀÕ":-14,"ÀÖ":-14,"ÀÙ":-17,"ÀÚ":-17,"ÀÛ":-17,"ÀÜ":-17,"ÀÝ":-82,"Àý":-30,"Àÿ":-30,"À‘":-81,"À’":-46,"À“":-81,"À”":-46,"Á,":12,"Á-":-12,"Á.":7,"Á?":-41,"ÁC":-12,"ÁG":-12,"ÁJ":14,"ÁO":-14,"ÁQ":-14,"ÁT":-78,"ÁU":-17,"ÁV":-51,"ÁW":-39,"ÁY":-82,"Át":-31,"Áv":-24,"Áy":-30,"ÁÇ":-12,"ÁÒ":-14,"ÁÓ":-14,"ÁÔ":-14,"ÁÕ":-14,"ÁÖ":-14,"ÁÙ":-17,"ÁÚ":-17,"ÁÛ":-17,"ÁÜ":-17,"ÁÝ":-82,"Áý":-30,"Áÿ":-30,"Á‘":-81,"Á’":-46,"Á“":-81,"Á”":-46,"Â,":12,"Â-":-12,"Â.":7,"Â?":-41,"ÂC":-12,"ÂG":-12,"ÂJ":14,"ÂO":-14,"ÂQ":-14,"ÂT":-78,"ÂU":-17,"ÂV":-51,"ÂW":-39,"ÂY":-82,"Ât":-31,"Âv":-24,"Ây":-30,"ÂÇ":-12,"ÂÒ":-14,"ÂÓ":-14,"ÂÔ":-14,"ÂÕ":-14,"ÂÖ":-14,"ÂÙ":-17,"ÂÚ":-17,"ÂÛ":-17,"ÂÜ":-17,"ÂÝ":-82,"Âý":-30,"Âÿ":-30,"Â‘":-81,"Â’":-46,"Â“":-81,"Â”":-46,"Ã,":12,"Ã-":-12,"Ã.":7,"Ã?":-41,"ÃC":-12,"ÃG":-12,"ÃJ":14,"ÃO":-14,"ÃQ":-14,"ÃT":-78,"ÃU":-17,"ÃV":-51,"ÃW":-39,"ÃY":-82,"Ãt":-31,"Ãv":-24,"Ãy":-30,"ÃÇ":-12,"ÃÒ":-14,"ÃÓ":-14,"ÃÔ":-14,"ÃÕ":-14,"ÃÖ":-14,"ÃÙ":-17,"ÃÚ":-17,"ÃÛ":-17,"ÃÜ":-17,"ÃÝ":-82,"Ãý":-30,"Ãÿ":-30,"Ã‘":-81,"Ã’":-46,"Ã“":-81,"Ã”":-46,"Ä,":12,"Ä-":-12,"Ä.":7,"Ä?":-41,"ÄC":-12,"ÄG":-12,"ÄJ":14,"ÄO":-14,"ÄQ":-14,"ÄT":-78,"ÄU":-17,"ÄV":-51,"ÄW":-39,"ÄY":-82,"Ät":-31,"Äv":-24,"Äy":-30,"ÄÇ":-12,"ÄÒ":-14,"ÄÓ":-14,"ÄÔ":-14,"ÄÕ":-14,"ÄÖ":-14,"ÄÙ":-17,"ÄÚ":-17,"ÄÛ":-17,"ÄÜ":-17,"ÄÝ":-82,"Äý":-30,"Äÿ":-30,"Ä‘":-81,"Ä’":-46,"Ä“":-81,"Ä”":-46,"Å,":12,"Å-":-12,"Å.":7,"Å?":-41,"ÅC":-12,"ÅG":-12,"ÅJ":14,"ÅO":-14,"ÅQ":-14,"ÅT":-78,"ÅU":-17,"ÅV":-51,"ÅW":-39,"ÅY":-82,"Åt":-31,"Åv":-24,"Åy":-30,"ÅÇ":-12,"ÅÒ":-14,"ÅÓ":-14,"ÅÔ":-14,"ÅÕ":-14,"ÅÖ":-14,"ÅÙ":-17,"ÅÚ":-17,"ÅÛ":-17,"ÅÜ":-17,"ÅÝ":-82,"Åý":-30,"Åÿ":-30,"Å‘":-81,"Å’":-46,"Å“":-81,"Å”":-46,"È-":-20,"ÈA":-12,"ÈC":-15,"ÈG":-15,"ÈO":-17,"ÈQ":-17,"ÈS":-10,"ÈZ":-5,"Èa":-13,"Èc":-12,"Èd":-15,"Èe":-17,"Èf":-27,"Èo":-17,"Èq":-15,"Èt":-15,"Èv":-26,"Èw":-20,"Èy":-26,"ÈÀ":-12,"ÈÁ":-12,"ÈÂ":-12,"ÈÃ":-12,"ÈÄ":-12,"ÈÅ":-12,"ÈÇ":-15,"ÈÒ":-17,"ÈÓ":-17,"ÈÔ":-17,"ÈÕ":-17,"ÈÖ":-17,"Èà":-13,"Èá":-13,"Èâ":-13,"Èã":-13,"Èä":-13,"Èå":-13,"Èæ":-13,"Èç":-12,"Èè":-17,"Èé":-17,"Èê":-17,"Èë":-17,"Èò":-17,"Èó":-17,"Èô":-17,"Èõ":-17,"Èö":-17,"Èø":-17,"Èý":-26,"Èÿ":-26,"É-":-20,"ÉA":-12,"ÉC":-15,"ÉG":-15,"ÉO":-17,"ÉQ":-17,"ÉS":-10,"ÉZ":-5,"Éa":-13,"Éc":-12,"Éd":-15,"Ée":-17,"Éf":-27,"Éo":-17,"Éq":-15,"Ét":-15,"Év":-26,"Éw":-20,"Éy":-26,"ÉÀ":-12,"ÉÁ":-12,"ÉÂ":-12,"ÉÃ":-12,"ÉÄ":-12,"ÉÅ":-12,"ÉÇ":-15,"ÉÒ":-17,"ÉÓ":-17,"ÉÔ":-17,"ÉÕ":-17,"ÉÖ":-17,"Éà":-13,"Éá":-13,"Éâ":-13,"Éã":-13,"Éä":-13,"Éå":-13,"Éæ":-13,"Éç":-12,"Éè":-17,"Éé":-17,"Éê":-17,"Éë":-17,"Éò":-17,"Éó":-17,"Éô":-17,"Éõ":-17,"Éö":-17,"Éø":-17,"Éý":-26,"Éÿ":-26,"Ê-":-20,"ÊA":-12,"ÊC":-15,"ÊG":-15,"ÊO":-17,"ÊQ":-17,"ÊS":-10,"ÊZ":-5,"Êa":-13,"Êc":-12,"Êd":-15,"Êe":-17,"Êf":-27,"Êo":-17,"Êq":-15,"Êt":-15,"Êv":-26,"Êw":-20,"Êy":-26,"ÊÀ":-12,"ÊÁ":-12,"ÊÂ":-12,"ÊÃ":-12,"ÊÄ":-12,"ÊÅ":-12,"ÊÇ":-15,"ÊÒ":-17,"ÊÓ":-17,"ÊÔ":-17,"ÊÕ":-17,"ÊÖ":-17,"Êà":-13,"Êá":-13,"Êâ":-13,"Êã":-13,"Êä":-13,"Êå":-13,"Êæ":-13,"Êç":-12,"Êè":-17,"Êé":-17,"Êê":-17,"Êë":-17,"Êò":-17,"Êó":-17,"Êô":-17,"Êõ":-17,"Êö":-17,"Êø":-17,"Êý":-26,"Êÿ":-26,"Ë-":-20,"ËA":-12,"ËC":-15,"ËG":-15,"ËO":-17,"ËQ":-17,"ËS":-10,"ËZ":-5,"Ëa":-13,"Ëc":-12,"Ëd":-15,"Ëe":-17,"Ëf":-27,"Ëo":-17,"Ëq":-15,"Ët":-15,"Ëv":-26,"Ëw":-20,"Ëy":-26,"ËÀ":-12,"ËÁ":-12,"ËÂ":-12,"ËÃ":-12,"ËÄ":-12,"ËÅ":-12,"ËÇ":-15,"ËÒ":-17,"ËÓ":-17,"ËÔ":-17,"ËÕ":-17,"ËÖ":-17,"Ëà":-13,"Ëá":-13,"Ëâ":-13,"Ëã":-13,"Ëä":-13,"Ëå":-13,"Ëæ":-13,"Ëç":-12,"Ëè":-17,"Ëé":-17,"Ëê":-17,"Ëë":-17,"Ëò":-17,"Ëó":-17,"Ëô":-17,"Ëõ":-17,"Ëö":-17,"Ëø":-17,"Ëý":-26,"Ëÿ":-26,"Ò,":-22,"Ò.":-15,"ÒA":-14,"ÒJ":-10,"ÒT":-22,"ÒV":-17,"ÒW":-12,"ÒX":-34,"ÒY":-32,"ÒZ":-16,"Òx":-7,"Òz":-5,"ÒÀ":-14,"ÒÁ":-14,"ÒÂ":-14,"ÒÃ":-14,"ÒÄ":-14,"ÒÅ":-14,"ÒÆ":-22,"ÒÝ":-32,"Ò‘":-7,"Ò’":7,"Ò“":-7,"Ò”":7,"Ó,":-22,"Ó.":-15,"ÓA":-14,"ÓJ":-10,"ÓT":-22,"ÓV":-17,"ÓW":-12,"ÓX":-34,"ÓY":-32,"ÓZ":-16,"Óx":-7,"Óz":-5,"ÓÀ":-14,"ÓÁ":-14,"ÓÂ":-14,"ÓÃ":-14,"ÓÄ":-14,"ÓÅ":-14,"ÓÆ":-22,"ÓÝ":-32,"Ó‘":-7,"Ó’":7,"Ó“":-7,"Ó”":7,"Ô,":-22,"Ô.":-15,"ÔA":-14,"ÔJ":-10,"ÔT":-22,"ÔV":-17,"ÔW":-12,"ÔX":-34,"ÔY":-32,"ÔZ":-16,"Ôx":-7,"Ôz":-5,"ÔÀ":-14,"ÔÁ":-14,"ÔÂ":-14,"ÔÃ":-14,"ÔÄ":-14,"ÔÅ":-14,"ÔÆ":-22,"ÔÝ":-32,"Ô‘":-7,"Ô’":7,"Ô“":-7,"Ô”":7,"Õ,":-22,"Õ.":-15,"ÕA":-14,"ÕJ":-10,"ÕT":-22,"ÕV":-17,"ÕW":-12,"ÕX":-34,"ÕY":-32,"ÕZ":-16,"Õx":-7,"Õz":-5,"ÕÀ":-14,"ÕÁ":-14,"ÕÂ":-14,"ÕÃ":-14,"ÕÄ":-14,"ÕÅ":-14,"ÕÆ":-22,"ÕÝ":-32,"Õ‘":-7,"Õ’":7,"Õ“":-7,"Õ”":7,"Ö,":-22,"Ö.":-15,"ÖA":-14,"ÖJ":-10,"ÖT":-22,"ÖV":-17,"ÖW":-12,"ÖX":-34,"ÖY":-32,"ÖZ":-16,"Öx":-7,"Öz":-5,"ÖÀ":-14,"ÖÁ":-14,"ÖÂ":-14,"ÖÃ":-14,"ÖÄ":-14,"ÖÅ":-14,"ÖÆ":-22,"ÖÝ":-32,"Ö‘":-7,"Ö’":7,"Ö“":-7,"Ö”":7,"Ù,":-12,"Ù.":-13,"ÙA":-27,"ÙJ":-16,"ÙÀ":-27,"ÙÁ":-27,"ÙÂ":-27,"ÙÃ":-27,"ÙÄ":-27,"ÙÅ":-27,"ÙÆ":-20,"Ú,":-12,"Ú.":-13,"ÚA":-27,"ÚJ":-16,"ÚÀ":-27,"ÚÁ":-27,"ÚÂ":-27,"ÚÃ":-27,"ÚÄ":-27,"ÚÅ":-27,"ÚÆ":-20,"Û,":-12,"Û.":-13,"ÛA":-27,"ÛJ":-16,"ÛÀ":-27,"ÛÁ":-27,"ÛÂ":-27,"ÛÃ":-27,"ÛÄ":-27,"ÛÅ":-27,"ÛÆ":-20,"Ü,":-12,"Ü.":-13,"ÜA":-27,"ÜJ":-16,"ÜÀ":-27,"ÜÁ":-27,"ÜÂ":-27,"ÜÃ":-27,"ÜÄ":-27,"ÜÅ":-27,"ÜÆ":-20,"àf":-5,"àt":-9,"àv":-14,"àw":-8,"àx":-9,"ày":-18,"àý":-18,"àÿ":-18,"à‘":-15,"à“":-15,"áf":-5,"át":-9,"áv":-14,"áw":-8,"áx":-9,"áy":-18,"áý":-18,"áÿ":-18,"á‘":-15,"á“":-15,"âf":-5,"ât":-9,"âv":-14,"âw":-8,"âx":-9,"ây":-18,"âý":-18,"âÿ":-18,"â‘":-15,"â“":-15,"ãf":-5,"ãt":-9,"ãv":-14,"ãw":-8,"ãx":-9,"ãy":-18,"ãý":-18,"ãÿ":-18,"ã‘":-15,"ã“":-15,"äf":-5,"ät":-9,"äv":-14,"äw":-8,"äx":-9,"äy":-18,"äý":-18,"äÿ":-18,"ä‘":-15,"ä“":-15,"åf":-5,"åt":-9,"åv":-14,"åw":-8,"åx":-9,"åy":-18,"åý":-18,"åÿ":-18,"å‘":-15,"å“":-15,"èf":-7,"èt":-7,"èv":-5,"èw":-5,"èx":-16,"èy":-9,"èz":-10,"èý":-9,"èÿ":-9,"è‘":-7,"è“":-7,"éf":-7,"ét":-7,"év":-5,"éw":-5,"éx":-16,"éy":-9,"éz":-10,"éý":-9,"éÿ":-9,"é‘":-7,"é“":-7,"êf":-7,"êt":-7,"êv":-5,"êw":-5,"êx":-16,"êy":-9,"êz":-10,"êý":-9,"êÿ":-9,"ê‘":-7,"ê“":-7,"ëf":-7,"ët":-7,"ëv":-5,"ëw":-5,"ëx":-16,"ëy":-9,"ëz":-10,"ëý":-9,"ëÿ":-9,"ë‘":-7,"ë“":-7,"îb":7,"îh":7,"îk":7,"îl":7,"îþ":7,"î‘":27,"î’":42,"î“":27,"î”":42,"ñf":-5,"ñt":-9,"ñv":-14,"ñw":-8,"ñx":-9,"ñy":-18,"ñý":-18,"ñÿ":-18,"ñ‘":-15,"ñ“":-15,"ò,":-15,"òv":-3,"òw":-2,"òx":-20,"òy":-6,"òz":-10,"òý":-6,"òÿ":-6,"ó,":-15,"óv":-3,"ów":-2,"óx":-20,"óy":-6,"óz":-10,"óý":-6,"óÿ":-6,"ô,":-15,"ôv":-3,"ôw":-2,"ôx":-20,"ôy":-6,"ôz":-10,"ôý":-6,"ôÿ":-6,"õ,":-15,"õv":-3,"õw":-2,"õx":-20,"õy":-6,"õz":-10,"õý":-6,"õÿ":-6,"ö,":-15,"öv":-3,"öw":-2,"öx":-20,"öy":-6,"öz":-10,"öý":-6,"öÿ":-6,"ý,":-61,"ý-":-17,"ý.":-61,"ýa":-15,"ýc":-13,"ýd":-11,"ýe":-11,"ýf":5,"ýg":-10,"ýo":-11,"ýq":-11,"ýs":-6,"ýt":5,"ýv":7,"ýw":6,"ýy":5,"ýz":-6,"ýà":-15,"ýá":-15,"ýâ":-15,"ýã":-15,"ýä":-15,"ýå":-15,"ýæ":-15,"ýç":-13,"ýè":-11,"ýé":-11,"ýê":-11,"ýë":-11,"ýò":-11,"ýó":-11,"ýô":-11,"ýõ":-11,"ýö":-11,"ýø":-11,"ýý":5,"ýÿ":5,"ý‘":12,"ý’":32,"ý“":12,"ý”":32,"ÿ,":-61,"ÿ-":-17,"ÿ.":-61,"ÿa":-15,"ÿc":-13,"ÿd":-11,"ÿe":-11,"ÿf":5,"ÿg":-10,"ÿo":-11,"ÿq":-11,"ÿs":-6,"ÿt":5,"ÿv":7,"ÿw":6,"ÿy":5,"ÿz":-6,"ÿà":-15,"ÿá":-15,"ÿâ":-15,"ÿã":-15,"ÿä":-15,"ÿå":-15,"ÿæ":-15,"ÿç":-13,"ÿè":-11,"ÿé":-11,"ÿê":-11,"ÿë":-11,"ÿò":-11,"ÿó":-11,"ÿô":-11,"ÿõ":-11,"ÿö":-11,"ÿø":-11,"ÿý":5,"ÿÿ":5,"ÿ‘":12,"ÿ’":32,"ÿ“":12,"ÿ”":32}],"arial":[{" A":-55," T":-18," Y":-18,"11":-74,"A ":-55,"AT":-74,"AV":-74,"AW":-37,"AY":-74,"Av":-18,"Aw":-18,"Ay":-18,"A’":-74,"F,":-111,"F.":-111,"FA":-55,"L ":-37,"LT":-74,"LV":-74,"LW":-74,"LY":-74,"Ly":-37,"L’":-55,"P ":-18,"P,":-129,"P.":-129,"PA":-74,"RT":-18,"RV":-18,"RW":-18,"RY":-18,"T ":-18,"T,":-111,"T-":-55,"T.":-111,"T:":-111,"T;":-111,"TA":-74,"TO":-18,"Ta":-111,"Tc":-111,"Te":-111,"Ti":-37,"To":-111,"Tr":-37,"Ts":-111,"Tu":-37,"Tw":-55,"Ty":-55,"V,":-92,"V-":-55,"V.":-92,"V:":-37,"V;":-37,"VA":-74,"Va":-74,"Ve":-55,"Vi":-18,"Vo":-55,"Vr":-37,"Vu":-37,"Vy":-37,"W,":-55,"W-":-18,"W.":-55,"W:":-18,"W;":-18,"WA":-37,"Wa":-37,"We":-18,"Wo":-18,"Wr":-18,"Wu":-18,"Wy":-9,"Y ":-18,"Y,":-129,"Y-":-92,"Y.":-129,"Y:":-55,"Y;":-65,"YA":-74,"Ya":-74,"Ye":-92,"Yi":-37,"Yo":-92,"Yp":-74,"Yq":-92,"Yu":-55,"Yv":-55,"ff":-18,"f’":18,"r,":-55,"r.":-55,"r’":37,"v,":-74,"v.":-74,"w,":-55,"w.":-55,"y,":-74,"y.":-74,"‘‘":-18,"’ ":-37,"’s":-18,"’’":-18},{" A":-37," Y":-18,"11":-55,"A ":-37,"AT":-74,"AV":-74,"AW":-55,"AY":-92,"Av":-37,"Aw":-18,"Ay":-37,"A’":-55,"F,":-111,"F.":-111,"FA":-55,"L ":-18,"LT":-74,"LV":-74,"LW":-55,"LY":-92,"Ly":-37,"L’":-55,"P ":-18,"P,":-129,"P.":-129,"PA":-74,"RV":-18,"RW":-18,"RY":-37,"T,":-111,"T-":-55,"T.":-111,"T:":-111,"T;":-111,"TA":-74,"TO":-18,"Ta":-74,"Tc":-74,"Te":-74,"Ti":-18,"To":-74,"Tr":-55,"Ts":-74,"Tu":-74,"Tw":-74,"Ty":-74,"V,":-92,"V-":-55,"V.":-92,"V:":-55,"V;":-55,"VA":-74,"Va":-55,"Ve":-55,"Vi":-18,"Vo":-74,"Vr":-55,"Vu":-37,"Vy":-37,"W,":-55,"W-":-20,"W.":-55,"W:":-18,"W;":-18,"WA":-55,"Wa":-37,"We":-18,"Wi":-9,"Wo":-18,"Wr":-18,"Wu":-18,"Wy":-18,"Y ":-18,"Y,":-111,"Y-":-55,"Y.":-111,"Y:":-74,"Y;":-74,"YA":-92,"Ya":-55,"Ye":-55,"Yi":-37,"Yo":-74,"Yp":-55,"Yq":-74,"Yu":-55,"Yv":-55,"f’":18,"r,":-55,"r.":-55,"r’":37,"v,":-74,"v.":-74,"w,":-37,"w.":-37,"y,":-74,"y.":-74,"‘‘":-37,"’ ":-55,"’s":-37,"’’":-37}]};
  const MET_IDX = new Map(METRICAS.codigos.map((c, i) => [c, i]));
  async function generarPdf(docxBytes) {
    const { PDFDocument, StandardFonts, rgb, pushGraphicsState, popGraphicsState, setCharacterSpacing, rectangle, clip, endPath } = PDFLib;
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
    // fuentes del tema (Calibri / Cambria…)
    const tema = { minor: null, major: null };
    const thf = Object.keys(zip.files).find(n => /^word\/theme\/theme\d*\.xml$/.test(n));
    if (thf) {
      const td = await zip.file(thf).async('string');
      const mi = td.match(/<a:minorFont>\s*<a:latin typeface="([^"]*)"/), ma = td.match(/<a:majorFont>\s*<a:latin typeface="([^"]*)"/);
      tema.minor = mi ? mi[1] : null; tema.major = ma ? ma[1] : null;
    }
    F.calibri = [await pdf.embedFont(StandardFonts.Helvetica), await pdf.embedFont(StandardFonts.HelveticaBold)];
    const metricaDe = new Map([[F.serif[0], METRICAS.tnr[0]], [F.serif[1], METRICAS.tnr[1]], [F.sans[0], METRICAS.arial[0]], [F.sans[1], METRICAS.arial[1]], [F.calibri[0], METRICAS.calibri[0]], [F.calibri[1], METRICAS.calibri[1]]]);
    const pickFont = (name, bold) => {
      if (name && name[0] === '@') name = /major/i.test(name) ? tema.major : tema.minor;
      const fam = /times|georgia|garamond|cambria|book|serif/i.test(name || '') ? 'serif' : /courier|consolas|mono/i.test(name || '') ? 'mono' : /calibri|carlito/i.test(name || '') ? 'calibri' : 'sans';
      return F[fam][bold ? 1 : 0];
    };
    // ancho como en Word
    const kernDe = new Map([[F.serif[0], KERNING.tnr[0]], [F.serif[1], KERNING.tnr[1]], [F.sans[0], KERNING.arial[0]], [F.sans[1], KERNING.arial[1]], [F.calibri[0], KERNING.calibri[0]], [F.calibri[1], KERNING.calibri[1]]]);
    const medir = (font, t, size, kern) => {
      const m = metricaDe.get(font); if (!m) return font.widthOfTextAtSize(t, size);
      const kp = kern ? kernDe.get(font) : null;
      let u = 0, prevCh = '';
      for (const ch of t) {
        const i = MET_IDX.get(ch.codePointAt(0)); u += (i != null && m[i]) ? m[i] : font.widthOfTextAtSize(ch, 1000);
        if (kp && prevCh) u += kp[prevCh + ch] || 0;
        prevCh = ch;
      }
      return u * size / 1000;
    };
    // dibuja ajustando el espacio entre letras para que la palabra ocupe lo mismo que en Word
    const dibujarTexto = (page, t, o) => {
      const real = o.font.widthOfTextAtSize(t, o.size), obj = medir(o.font, t, o.size, o.kern), n = Array.from(t).length;
      const tc = n ? (obj - real) / n : 0;
      const op = { x: o.x, y: o.y, size: o.size, font: o.font, color: o.color };
      if (Math.abs(tc) < 0.001) return page.drawText(t, op);
      page.pushOperators(pushGraphicsState(), setCharacterSpacing(tc));
      page.drawText(t, op);
      page.pushOperators(popGraphicsState());
    };
    // imagen con recorte (srcRect de Word)
    const dibujarImagen = (page, img, x, y, w, h, crop) => {
      if (!crop) return page.drawImage(img, { x, y, width: w, height: h });
      const fw = w / Math.max(0.01, 1 - (crop.l + crop.r)), fh = h / Math.max(0.01, 1 - (crop.t + crop.b));
      page.pushOperators(pushGraphicsState(), rectangle(x, y, w, h), clip(), endPath());
      page.drawImage(img, { x: x - fw * crop.l, y: y - fh * crop.b, width: fw, height: fh });
      page.pushOperators(popGraphicsState());
    };
    const recorteDe = c => {
      const sr = all(c, 'srcRect', NS.a)[0]; if (!sr) return null;
      const g = k => (+sr.getAttribute(k) || 0) / 100000;
      const r = { l: g('l'), t: g('t'), r: g('r'), b: g('b') };
      return (r.l || r.t || r.r || r.b) ? r : null;
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
        const rp = { sz: o.sz, b: o.b, font: o.font, caps: o.caps, u: false, kern: o.kern };
        const rPr = kid(r, 'rPr');
        const rSt = rPr && kid(rPr, 'rStyle'); if (rSt) st.charStyle(rp, wAttr(rSt, 'val'));
        if (rPr) applyRPr(rp, rPr);
        for (const c of Array.from(r.childNodes)) {
          if (c.nodeType !== 1) continue;
          if (c.localName === 't') { if (c.textContent.length) { const tx = rp.caps ? c.textContent.toLocaleUpperCase('es-ES') : c.textContent; items.push({ type: 'text', t: tx, font: pickFont(rp.font, rp.b), u: rp.u, size: rp.sz / 2, kern: !!rp.kern && rp.sz >= rp.kern }); if (c.textContent.trim()) sawContent = true; } }
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
              anchors.push({ img, w: +ext.getAttribute('cx') / 12700, h: +ext.getAttribute('cy') / 12700, offX: offOf(ph), offY: offOf(pv), relH: ph ? ph.getAttribute('relativeFrom') : 'column', crop: recorteDe(c) });
            } else if (img && ext) { items.push({ type: 'img', img, w: +ext.getAttribute('cx') / 12700, h: +ext.getAttribute('cy') / 12700, crop: recorteDe(c) }); sawContent = true; }
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
      // tabuladores al principio: Word salta a las paradas por defecto (cada 1,25 cm desde el margen), no desde la sangría
      const leadOffset = leadTabs ? (Math.floor((indL + 0.01) / tabStop) + leadTabs) * tabStop - indL : 0;
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
        for (let pi = 0; pi < parts.length; pi++) {
          const part = parts[pi];
          if (part === '') continue;
          if (/^ +$/.test(part)) {
            // el espacio incluye el kerning con la letra anterior y la siguiente (p. ej. «Y C»), como Word
            let w = medir(font, part, it.size, it.kern);
            if (it.kern) {
              const a = Array.from(parts[pi - 1] || '').pop() || '', b = Array.from(parts[pi + 1] || '')[0] || '';
              w = medir(font, a + part + b, it.size, true) - medir(font, a, it.size) - medir(font, b, it.size);
            }
            words.push({ type: 'space', w, size: it.size }); continue;
          }
          words.push({ type: 'word', t: part, font, u: it.u, size: it.size, kern: it.kern, w: medir(font, part, it.size, it.kern) });
        }
      }
      // líneas
      let endsBr = false;
      const lines = []; let cur = { items: [], w: 0, h: lineH, x0: 0 };
      // con número o viñeta el texto de la 1ª línea empieza en la sangría (el número va en la sangría francesa)
      const startX = (i) => (i === 0 && !bullet ? first + leadOffset : leadOffset);
      const maxW = i => width - startX(i);
      const push = () => { while (cur.items.length && cur.items[cur.items.length - 1].type === 'space') { cur.w -= cur.items.pop().w; } lines.push(cur); cur = { items: [], w: 0, h: lineH }; };
      for (const wd of words) {
        if (wd.type === 'br') { push(); endsBr = true; continue; }
        endsBr = false;
        if (wd.type === 'space') { if (cur.items.length || lines.length === 0) { cur.items.push(wd); cur.w += wd.w; } continue; }
        const ww = wd.w;
        if (cur.w + ww > maxW(lines.length) + 0.25 && cur.items.length) push(); // 0,25 pt de tolerancia por redondeo, como Word
        if (wd.type === 'img') { let w = wd.w, h = wd.h; const mw = maxW(lines.length); if (w > mw) { h = h * mw / w; w = mw; } cur.items.push({ ...wd, w, h }); cur.w += w; cur.h = Math.max(cur.h, h + 2); continue; }
        // palabra más larga que la línea: partir por caracteres
        if (ww > maxW(lines.length) + 0.25) {
          let chunk = '';
          for (const ch of wd.t) {
            const tw = medir(wd.font, chunk + ch, wd.size, wd.kern);
            if (tw > maxW(lines.length) - cur.w && chunk) { const cw = medir(wd.font, chunk, wd.size, wd.kern); cur.items.push({ ...wd, t: chunk, w: cw }); cur.w += cw; push(); chunk = ''; }
            chunk += ch;
          }
          if (chunk) { const w = medir(wd.font, chunk, wd.size, wd.kern); cur.items.push({ ...wd, t: chunk, w }); cur.w += w; }
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
            dibujarImagen(page, an.img, ax, yTop - sb - an.offY - an.h, an.w, an.h, an.crop);
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
            if (i === 0 && bullet) dibujarTexto(page, safe(bullet, bf), { x: x + indL + first, y: y - ln.h + (ln.h - fsize) * 0.35 + 1, size: fsize, font: bf, color: rgb(0, 0, 0) });
            let cx = sx + (i === 0 && bullet ? -first : 0) + dx;
            if (i === 0 && bullet) cx = x + indL + dx + leadOffset;
            const base = y - ln.h + (ln.h - fsize) * 0.35 + 1;
            for (const it of ln.items) {
              if (it.type === 'space') { cx += it.w + extra; continue; }
              if (it.type === 'img') { dibujarImagen(page, it.img, cx, y - it.h - 1, it.w, it.h, it.crop); cx += it.w; continue; }
              dibujarTexto(page, it.t, { x: cx, y: base, size: it.size, font: it.font, color: rgb(0, 0, 0), kern: it.kern });
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
      const hijos = Array.from(hdoc.documentElement.childNodes).filter(n => n.nodeType === 1);
      for (const el of hijos) {
        if (el.localName === 'p') {
          // imágenes ancladas: se tratan como en línea
          for (const anc of all(el, 'anchor', NS.wp)) {
            const blip = all(anc, 'blip', NS.a)[0]; const ext = all(anc, 'extent', NS.wp)[0];
            const rid2 = blip && (blip.getAttributeNS(NS.r, 'embed') || blip.getAttribute('r:embed'));
            const img = rid2 ? await loadImg(hrels, 'word/', rid2) : null;
            if (img && ext) {
              const posH = all(anc, 'positionH', NS.wp)[0]; const off = posH && posH.getElementsByTagNameNS(NS.wp, 'posOffset')[0];
              const al = posH && posH.getElementsByTagNameNS(NS.wp, 'align')[0];
              const posV = all(anc, 'positionV', NS.wp)[0]; const offV = posV && posV.getElementsByTagNameNS(NS.wp, 'posOffset')[0];
              // en el pie, la imagen se coloca respecto a su párrafo: se mide lo que ocupan ese párrafo y los de debajo (aunque estén vacíos)
              let debajo = 0;
              for (const p2 of hijos.slice(hijos.indexOf(el))) if (p2.localName === 'p') debajo += (await paraBlock(p2, hrels, 'word/', CW)).h;
              out.push({ kind: 'anchor', debajo, img, w: +ext.getAttribute('cx') / 12700, h: +ext.getAttribute('cy') / 12700, offX: off ? +off.textContent / 12700 : null, offY: offV && (!posV.getAttribute('relativeFrom') || /paragraph|line/.test(posV.getAttribute('relativeFrom'))) ? +offV.textContent / 12700 : 0, align: al ? al.textContent : null, rel: posH ? posH.getAttribute('relativeFrom') : 'column', crop: recorteDe(anc) });
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
          const topAnc = blocks === footer ? T(geom.footer) + Math.max(b.debajo || 0, footerH) : yTop;
          let x = ML + (b.offX || 0);
          if (b.rel === 'page') x = b.offX || 0;
          if (b.align === 'right') x = ML + CW - b.w; else if (b.align === 'center') x = ML + (CW - b.w) / 2;
          dibujarImagen(page, b.img, x, topAnc - (b.offY || 0) - b.h, b.w, b.h, b.crop);
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
          gh += bj.h; if (!bj.keepNext) { gh -= bj.after || 0; break; } j++;
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
