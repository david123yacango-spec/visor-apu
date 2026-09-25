/*
 * Lector de los Excel de exportación de Power Cost / vidBIM Costos (mismo formato):
 *   · «Análisis de Costos Unitarios» (hoja SP1): un bloque por partida.
 *   · «Presupuesto» (hoja SP1): ítems desde la fila 10.
 * Recibe el libro ya abierto con SheetJS (XLSX.read). Todo pasa en el navegador.
 * Mismas reglas que vidBIM Costos (Nucleo/PowerCost/ImportadorPowerCost.cs y el motor).
 */
(function (raiz) {
  'use strict';

  var CATEGORIAS = {
    'Mano de Obra': 'mo', 'Materiales': 'mat', 'Equipo': 'eq', 'Subcontratos': 'sc', 'Servicios': 'sc'
  };
  var NOMBRE_CAT = { mo: 'Mano de obra', mat: 'Materiales', eq: 'Equipo', sc: 'Subcontratos' };
  var RX_REND = /^\s*([\d,]+(?:\.\d+)?)?\s*-?\s*(.*?)\s*\/\s*DIA\s*$/i;
  var RX_CU = /Costo Unitario por\s+(.*?)\s*:/i;
  var RX_ITEM = /^\d+(\.\d+)*$/;

  function r2(x) { return Math.round((x + (x >= 0 ? 1e-9 : -1e-9)) * 100) / 100; }
  function r4(x) { return Math.round((x + (x >= 0 ? 1e-11 : -1e-11)) * 10000) / 10000; }

  /** Filas de una hoja como arreglos (columna A = índice 0); las celdas con solo espacios cuentan como vacías. */
  function filasDe(XLSX, ws) {
    var filas = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
    for (var i = 0; i < filas.length; i++) {
      var f = filas[i] || [];
      for (var j = 0; j < f.length; j++) if (typeof f[j] === 'string' && f[j].trim() === '') f[j] = null;
      filas[i] = f;
    }
    return filas;
  }
  function txt(f, c) { var v = f[c]; return v == null ? null : String(v); }
  function num(f, c) {
    var v = f[c];
    if (typeof v === 'number') return v;
    if (typeof v === 'string') { var d = parseFloat(v.replace(/,/g, '')); return isNaN(d) ? null : d; }
    return null;
  }
  function usadas(f) { var n = []; for (var j = 0; j < f.length; j++) if (f[j] != null) n.push(j); return n; }

  function encabezado(filas, hasta) {
    var e = {};
    for (var r = 0; r < Math.min(hasta, filas.length); r++) {
      var f = filas[r];
      for (var c = 0; c < f.length; c++) {
        if (typeof f[c] !== 'string') continue;
        var k = f[c].trim().replace(/:$/, '').trim();
        if (['Proyecto', 'Sub Presupuesto', 'Cliente', 'Ubicación', 'Costo a'].indexOf(k) < 0) continue;
        for (var d = c + 1; d < f.length; d++) if (f[d] != null) { e[k] = String(f[d]).trim(); break; }
      }
    }
    return e;
  }

  /** ¿Qué es este libro? 'apu', 'presupuesto' o null. */
  function tipoDe(XLSX, wb) {
    var ws = wb.Sheets.SP1;
    if (!ws) return null;
    var filas = filasDe(XLSX, ws).slice(0, 15);
    for (var r = 0; r < filas.length; r++) {
      var t = (filas[r] || []).filter(function (v) { return typeof v === 'string'; }).join(' ');
      if (/An.lisis de Costos Unitarios/i.test(t)) return 'apu';
      if (/^\s*Presupuesto\b/i.test(t)) return 'presupuesto';
    }
    return null;
  }

  /** APU: lista de { item, descripcion, unidad, rendimiento, rendTexto, lineas[], subtotales{cat}, cu } */
  function leerApu(XLSX, wb) {
    var filas = filasDe(XLSX, wb.Sheets.SP1);
    var apus = [], cur = null, cat = null;
    for (var r = 0; r < filas.length; r++) {
      var f = filas[r], a0 = txt(f, 0);
      if (a0 === 'Partida') {
        cur = { fila: r + 1, item: (txt(f, 1) || '').trim(), descripcion: (txt(f, 3) || '').trim(), rendTexto: (txt(f, 9) || '').trim(),
                unidad: '', rendimiento: null, lineas: [], subtotales: {}, cu: null };
        var m = RX_REND.exec(cur.rendTexto);
        if (m) {
          if (m[1]) cur.rendimiento = parseFloat(m[1].replace(/,/g, ''));
          cur.unidad = (m[2] || '').trim();
        }
        apus.push(cur); cat = null; continue;
      }
      if (!cur) continue;
      if (txt(f, 1) === 'Código') continue;
      var u = usadas(f);
      if (u.length === 1 && u[0] === 2 && CATEGORIAS[(txt(f, 2) || '').trim()]) { cat = CATEGORIAS[txt(f, 2).trim()]; continue; }
      var t8 = txt(f, 8), mc = t8 ? RX_CU.exec(t8) : null;
      if (mc) { cur.cu = num(f, 9); if (!cur.unidad) cur.unidad = mc[1].trim(); continue; }
      if (u.length === 1 && u[0] === 9 && cat) { cur.subtotales[cat] = num(f, 9) || 0; continue; }
      if (f[1] != null && cat) {
        cur.lineas.push({
          fila: r + 1, categoria: cat, codigo: String(f[1]).trim(), descripcion: (txt(f, 2) || '').trim(), unidad: (txt(f, 4) || '').trim(),
          cuadrilla: num(f, 5), cantidad: num(f, 6), precio: num(f, 7), parcial: num(f, 9)
        });
      }
    }
    return { encabezado: encabezado(filas, 10), apus: apus };
  }

  /** Presupuesto: lista de { item, descripcion, unidad, metrado, precio, parcial, total, nivel, esPartida } */
  function leerPresupuesto(XLSX, wb) {
    var filas = filasDe(XLSX, wb.Sheets.SP1), lista = [];
    for (var r = 9; r < filas.length; r++) {
      var f = filas[r], item = (txt(f, 0) || '').trim();
      if (!RX_ITEM.test(item)) continue;
      var und = txt(f, 3);
      var fila = {
        fila: r + 1, item: item, descripcion: (txt(f, 1) || '').trim(), unidad: und == null ? null : und.trim(),
        metrado: num(f, 4), precio: num(f, 5), parcial: num(f, 6), subtotal: num(f, 7), total: num(f, 8),
        nivel: item.split('.').length, esPartida: und != null
      };
      if (!fila.esPartida) fila.total = fila.total != null ? fila.total : fila.subtotal;
      lista.push(fila);
    }
    return { encabezado: encabezado(filas, 8), filas: lista };
  }

  var TIEMPO = { hh: 1, hm: 1, he: 1 };

  /**
   * Explica cada línea del APU como lo calcula Power Cost, y comprueba lo que trae el Excel.
   * tipo: 'tiempo' (cuadrilla × 8 ÷ rendimiento), 'dia' (cuadrilla ÷ rendimiento), 'pmo' (% de la mano de obra),
   *       'directa' (cantidad escrita).
   */
  function explicar(apu) {
    var mo = apu.subtotales.mo || 0, difs = [];
    apu.lineas.forEach(function (l) {
      var u = l.unidad.toLowerCase(), rend = apu.rendimiento;
      if (u === '%mo') {
        l.tipo = 'pmo';
        l.calcPrecio = mo;
        l.calcParcial = r2((l.cantidad || 0) * mo / 100);
      } else {
        if (TIEMPO[u] && rend && l.cuadrilla != null) { l.tipo = 'tiempo'; l.calcCantidad = r4(l.cuadrilla * 8 / rend); }
        else if (u === 'dia' && rend && l.cuadrilla != null) { l.tipo = 'dia'; l.calcCantidad = r4(l.cuadrilla / rend); }
        else l.tipo = 'directa';
        l.calcParcial = r2((l.cantidad || 0) * (l.precio || 0));
      }
      if (l.calcCantidad != null && Math.abs(l.calcCantidad - (l.cantidad || 0)) > 0.00005) difs.push(l.descripcion + ': cantidad');
      if (Math.abs(l.calcParcial - (l.parcial || 0)) > 0.005) difs.push(l.descripcion + ': parcial');
    });
    apu.calcSubtotales = {};
    Object.keys(NOMBRE_CAT).forEach(function (c) {
      var ls = apu.lineas.filter(function (l) { return l.categoria === c; });
      if (ls.length || apu.subtotales[c] != null)
        apu.calcSubtotales[c] = r2(ls.reduce(function (s, l) { return s + (l.parcial || 0); }, 0));
    });
    apu.calcCu = r2(Object.keys(apu.calcSubtotales).reduce(function (s, c) { return s + apu.calcSubtotales[c]; }, 0));
    Object.keys(apu.calcSubtotales).forEach(function (c) {
      if (Math.abs(apu.calcSubtotales[c] - (apu.subtotales[c] || 0)) > 0.005) difs.push('subtotal de ' + NOMBRE_CAT[c]);
    });
    if (apu.cu != null && Math.abs(apu.calcCu - apu.cu) > 0.005) difs.push('costo unitario');
    apu.diferencias = difs;
    return apu;
  }

  /** Junta presupuesto y APU por ítem; arma el árbol y los totales. */
  function armar(pre, apu) {
    var apus = apu ? apu.apus.map(explicar) : [];
    var porItem = {};
    apus.forEach(function (a) { porItem[a.item] = a; });
    var filas = pre ? pre.filas : apus.map(function (a) {
      return { item: a.item, descripcion: a.descripcion, unidad: a.unidad, metrado: null, precio: a.cu, parcial: null, nivel: 1, esPartida: true };
    });
    var avisos = [];
    // Títulos que el Excel del presupuesto trae de otros ítems (sin APU) y partidas sin APU.
    filas.forEach(function (f) {
      if (!f.esPartida) return;
      f.apu = porItem[f.item] || null;
      if (apu && !f.apu) avisos.push('Sin APU: ' + f.item + ' ' + f.descripcion);
      else if (f.apu && f.precio != null && Math.abs(f.apu.cu - f.precio) > 0.005)
        avisos.push('P.U. distinto: ' + f.item + ' presupuesto ' + f.precio + ' / APU ' + f.apu.cu);
    });
    // Insumos: dónde se usa cada uno.
    var insumos = {};
    apus.forEach(function (a) {
      a.lineas.forEach(function (l) {
        var k = l.codigo + '|' + l.descripcion;
        var i = insumos[k] || (insumos[k] = { codigo: l.codigo, descripcion: l.descripcion, unidad: l.unidad, categoria: l.categoria, usos: [] });
        i.usos.push({ item: a.item, cantidad: l.cantidad, precio: l.precio, parcial: l.parcial });
      });
    });
    var cd = pre ? r2(filas.filter(function (f) { return f.esPartida; }).reduce(function (s, f) { return s + (f.parcial || 0); }, 0)) : null;
    var cdTitulos = pre ? r2(filas.filter(function (f) { return !f.esPartida && f.nivel === 1; }).reduce(function (s, f) { return s + (f.total || 0); }, 0)) : null;
    return {
      encabezado: Object.assign({}, apu ? apu.encabezado : {}, pre ? pre.encabezado : {}),
      filas: filas, apus: apus, porItem: porItem, insumos: insumos, cd: cd, cdTitulos: cdTitulos, avisos: avisos,
      conDiferencias: apus.filter(function (a) { return a.diferencias.length; })
    };
  }

  var Lector = { tipoDe: tipoDe, leerApu: leerApu, leerPresupuesto: leerPresupuesto, explicar: explicar, armar: armar, r2: r2, r4: r4, NOMBRE_CAT: NOMBRE_CAT };
  if (typeof module !== 'undefined' && module.exports) module.exports = Lector; else raiz.Lector = Lector;
})(this);
