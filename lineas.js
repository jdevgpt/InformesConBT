/* JDev · Base de centros: número de CD/PT -> nombre de la línea (lineas.json) */
(function () {
  'use strict';
  var base = null, carga = null;

  function cargar() {
    if (!carga) {
      carga = fetch('lineas.json')
        .then(function (r) { if (!r.ok) throw new Error(r.status); return r.json(); })
        .then(function (j) { base = j; return j; })
        .catch(function () { carga = null; return null; });
    }
    return carga;
  }
  function clave(n) {
    return String(n == null ? '' : n).toUpperCase().replace(/\s+/g, '')
      .replace(/^(CD|CT|PT|ST)[-_.]?/, '').replace(/^0+(?=\d)/, '');
  }
  // Devuelve el nombre de la línea, null si el CD no está, o undefined si la base aún no ha cargado.
  function buscar(n) {
    if (!base) return undefined;
    var k = clave(n);
    if (!k) return undefined;
    var i = base.c[k];
    return i === undefined ? null : base.l[i];
  }
  // Rellena la línea al escribir el número del centro.
  // opciones: { pista: elemento para el mensaje, alRellenar: function(linea), forzar: cambiar la línea aunque ya haya una escrita }
  function vincular(inCD, inLinea, op) {
    op = op || {};
    var auto = '';
    function poner(v) {
      inLinea.value = v; auto = v;
      inLinea.dispatchEvent(new Event('input', { bubbles: true }));
      inLinea.dispatchEvent(new Event('change', { bubbles: true }));
      if (op.alRellenar) op.alRellenar(v);
    }
    function pista(txt, ok) {
      if (!op.pista) return;
      op.pista.textContent = txt;
      op.pista.style.color = ok ? 'var(--naranja-osc, #D9650A)' : '';
    }
    function revisar() {
      var n = inCD.value.trim();
      if (!n) { pista('', false); return; }
      var l = buscar(n);
      if (l === undefined) { pista('Cargando base de centros…', false); cargar().then(function (j) { if (j) revisar(); else pista('', false); }); return; }
      var libre = op.forzar || !inLinea.value.trim() || inLinea.value === auto;
      if (l) {
        if (libre) { if (inLinea.value !== l) poner(l); pista('✓ Línea: ' + l, true); }
        else if (inLinea.value.trim().toUpperCase() !== l.toUpperCase()) pista('Según la base, este centro es de la línea ' + l, false);
        else pista('✓ Línea: ' + l, true);
      } else {
        if (auto && inLinea.value === auto) poner('');
        pista('Centro no encontrado en la base: escribe la línea a mano.', false);
      }
    }
    inCD.addEventListener('input', revisar);
    inLinea.addEventListener('input', function (e) { if (e.isTrusted) auto = ''; });
    return { revisar: revisar, olvidar: function () { auto = ''; pista('', false); } };
  }

  window.LineasCD = { cargar: cargar, buscar: buscar, vincular: vincular };
  cargar();
})();
