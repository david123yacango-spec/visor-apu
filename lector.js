/*
 * Lector de los Excel de exportación de Power Cost / vidBIM Costos (mismo formato):
 *   · «Análisis de Costos Unitarios» (hoja SP1, SP4…): un bloque por partida y, si las hay, los bloques
 *     «Sub Partida» con el APU de cada subpartida.
 *   · «Presupuesto» (hoja SP1, SP4…): ítems desde la fila 10.
 * También los reportes de Delphin Express («PRESUPUESTO DE OBRA» y «Analisis de Costos Unitarios», .xls o .xlsx):
 * sus APU pueden venir con la numeración de otra versión del presupuesto, así que se cruzan por nombre.
 * Recibe el libro ya abierto con SheetJS (XLSX.read). Todo pasa en el navegador.
 * Mismas reglas que vidBIM Costos (Nucleo/PowerCost/ImportadorPowerCost.cs y el motor).
 */
(function (raiz) {
  'use strict';

  var CATEGORIAS = {
    'Mano de Obra': 'mo', 'Materiales': 'mat', 'Equipo': 'eq', 'Subcontratos': 'sc', 'Servicios': 'sc', 'Sub partidas': 'sp', 'Subpartidas': 'sp'
  };
  var NOMBRE_CAT = { mo: 'Mano de obra', mat: 'Materiales', eq: 'Equipo', sc: 'Subcontratos', sp: 'Subpartidas' };
  var RX_SUB = /^Sub\s*Partida\s+(\S+)/i;
  var RX_REND = /^\s*([\d,]+(?:\.\d+)?)?\s*-?\s*(.*?)\s*\/\s*DIA\s*$/i;
  var RX_CU = /Costo Unitario por\s+(.*?)\s*:/i;
  var RX_ITEM = /^\d+(\.\d+)*$/;
  // Delphin Express
  var RX_REND_D = /^\s*Rendimiento\s*:\s*([\d,]+(?:\.\d+)?)?\s*:?\s*(.*?)\s*(?:\/\s*d[ií]a)?\s*$/i;
  var RX_CU_D = /^\s*Costo\s+unit\.?\s+por\s+(.*?)\s*:?\s*$/i;
  function categoriaDelphin(s) {
    var k = String(s || '').toUpperCase().replace(/[^A-Z]/g, '');
    return { MANODEOBRA: 'mo', MATERIALES: 'mat', MATERIAL: 'mat', EQUIPOS: 'eq', EQUIPO: 'eq', EQUIPOSYHERRAMIENTAS: 'eq', HERRAMIENTAS: 'eq',
             SUBCONTRATOS: 'sc', SUBCONTRATO: 'sc', SERVICIOS: 'sc', SUBPARTIDAS: 'sp' }[k] || null;
  }
  /** ¿Reporte de Delphin? Título «PRESUPUESTO DE OBRA» o bloques «Partida: ítem» en la columna A. */
  function esDelphin(filas) {
    for (var r = 0; r < Math.min(filas.length, 40); r++) {
      var a = txt(filas[r], 0), f = filas[r];
      if (a && /^\s*Partida\s*:/i.test(a)) return true;
      if (f.some(function (v) { return typeof v === 'string' && /^\s*PRESUPUESTO DE OBRA\s*$/i.test(v); })) return true;
    }
    return false;
  }
  /** Nombre para cruzar partida y APU: mayúsculas, sin tildes ni signos. */
  function nombre(s) {
    return String(s || '').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^A-Z0-9Ñ]+/g, ' ').trim();
  }

  function r2(x) { return Math.round((x + (x >= 0 ? 1e-9 : -1e-9)) * 100) / 100; }
  function r4(x) { return Math.round((x + (x >= 0 ? 1e-11 : -1e-11)) * 10000) / 10000; }

  /** Filas de una hoja como arreglos (columna A = índice 0); las celdas con solo espacios cuentan como vacías. */
  function filasDe(XLSX, ws) {
    // Siempre desde A1: si la hoja empieza más abajo o más a la derecha, las filas y columnas no se corren.
    var rango = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : null;
    if (rango) { rango.s.r = 0; rango.s.c = 0; }
    var filas = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true, range: rango ? XLSX.utils.encode_range(rango) : undefined });
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

  var ETIQUETAS = ['Proyecto', 'Sub Presupuesto', 'Cliente', 'Ubicación', 'Costo a'];
  function encabezado(filas, hasta) {
    var e = {}, ultima = null, colValor = -1;
    for (var r = 0; r < Math.min(hasta, filas.length); r++) {
      var f = filas[r], hubo = false;
      for (var c = 0; c < f.length; c++) {
        if (typeof f[c] !== 'string') continue;
        var k = f[c].trim().replace(/:$/, '').trim();
        if (ETIQUETAS.indexOf(k) < 0) continue;
        hubo = true;
        for (var d = c + 1; d < f.length; d++) if (f[d] != null) { e[k] = String(f[d]).trim(); if (c === 0) { ultima = k; colValor = d; } break; }
      }
      // Un nombre largo sigue en la fila de abajo, en la misma columna y sin etiqueta.
      if (!hubo && ultima && colValor >= 0 && typeof f[colValor] === 'string' && f.filter(function (v) { return v != null; }).length === 1)
        e[ultima] += ' ' + f[colValor].trim();
      else if (hubo || f.some(function (v) { return v != null; })) ultima = hubo ? ultima : null;
    }
    return e;
  }

  /** La hoja de datos: SP1, SP4… (la primera «SPn» que traiga el título), o la primera hoja. */
  function hojaDe(XLSX, wb) {
    var nombres = wb.SheetNames.filter(function (n) { return /^SP\d+$/i.test(n); });
    for (var i = 0; i < nombres.length; i++) {
      var filas = filasDe(XLSX, wb.Sheets[nombres[i]]).slice(0, 15);
      if (filas.some(function (f) { return f.some(function (v) { return typeof v === 'string' && /(An.lisis de Costos Unitarios|^\s*Presupuesto\b)/i.test(v); }); })) return wb.Sheets[nombres[i]];
    }
    return nombres.length ? wb.Sheets[nombres[0]] : wb.Sheets[wb.SheetNames[0]];
  }

  /** ¿Qué es este libro? 'apu', 'presupuesto' o null. */
  function tipoDe(XLSX, wb) {
    var ws = hojaDe(XLSX, wb);
    if (!ws) return null;
    var filas = filasDe(XLSX, ws).slice(0, 15);
    for (var r = 0; r < filas.length; r++) {
      var t = (filas[r] || []).filter(function (v) { return typeof v === 'string'; }).join(' ');
      if (/^\s*PRESUPUESTO DE OBRA\b/i.test(t)) return 'presupuesto';   // Delphin
      if (/An.lisis de Costos Unitarios/i.test(t)) return 'apu';
      if (/^\s*Presupuesto\b/i.test(t)) return 'presupuesto';
    }
    return null;
  }

  /**
   * APU: lista de { item, descripcion, unidad, rendimiento, rendTexto, lineas[], subtotales{cat}, cu } y las
   * subpartidas por código («SP 11884» → su APU, con el mismo formato).
   */
  function leerApu(XLSX, wb) {
    var filas = filasDe(XLSX, hojaDe(XLSX, wb));
    if (esDelphin(filas)) return leerApuDelphin(filas);
    var apus = [], subpartidas = {}, cur = null, cat = null;
    function nuevo(r, item, descripcion, rendTexto) {
      var a = { fila: r + 1, item: item, descripcion: descripcion, rendTexto: rendTexto, unidad: '', rendimiento: null, lineas: [], subtotales: {}, cu: null };
      var m = RX_REND.exec(rendTexto);
      if (m) { if (m[1]) a.rendimiento = parseFloat(m[1].replace(/,/g, '')); a.unidad = (m[2] || '').trim(); }
      return a;
    }
    for (var r = 0; r < filas.length; r++) {
      var f = filas[r], a0 = txt(f, 0), b1 = txt(f, 1);
      if (a0 === 'Partida') {
        cur = nuevo(r, (b1 || '').trim(), (txt(f, 3) || '').trim(), (txt(f, 9) || '').trim());
        apus.push(cur); cat = null; continue;
      }
      var ms = b1 && !a0 ? RX_SUB.exec(b1.trim()) : null;
      if (ms) {
        // Bloque «Sub Partida 11884»: el APU de la línea «SP 11884» de las partidas. Se repite por cada partida que la usa.
        cur = nuevo(r, 'SP ' + ms[1], (txt(f, 3) || '').trim(), (txt(f, 9) || '').trim());
        cur.esSubpartida = true;
        if (!subpartidas[cur.item]) subpartidas[cur.item] = cur;
        cat = null; continue;
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
    return { encabezado: encabezado(filas, 10), apus: apus, subpartidas: subpartidas };
  }

  /**
   * APU de Delphin: «Partida: ítem» y «Rendimiento: …» (N), el nombre debajo, «Costo unit. por und» (N) con el CU en Q,
   * categorías en A (MANO DE OBRA, MATERIALES, EQUIPOS, SUB-CONTRATOS) y líneas: código A, descripción B, unidad J,
   * recursos K, cantidad M, precio P, parcial Q.
   */
  function leerApuDelphin(filas) {
    var apus = [], cur = null, cat = null, enLineas = false;
    for (var r = 0; r < filas.length; r++) {
      var f = filas[r], a = txt(f, 0), b = txt(f, 1);
      a = a == null ? null : a.trim();
      if (a && /^Partida\s*:/i.test(a)) {
        cur = { fila: r + 1, item: a.replace(/^Partida\s*:/i, '').trim(), descripcion: '', rendTexto: (txt(f, 13) || '').trim(), unidad: '', rendimiento: null,
                lineas: [], subtotales: {}, cu: null };
        var m = RX_REND_D.exec(cur.rendTexto);
        if (m) { if (m[1]) cur.rendimiento = parseFloat(m[1].replace(/,/g, '')); cur.unidad = (m[2] || '').trim(); }
        cur.rendTexto = cur.rendTexto.replace(/^Rendimiento\s*:\s*/i, '');
        apus.push(cur); cat = null; enLineas = false; continue;
      }
      if (!cur) continue;
      var mc = RX_CU_D.exec(txt(f, 13) || '');
      if (mc) { cur.cu = num(f, 16); cur.unidad = mc[1].trim(); }
      if (!enLineas) {
        if (a && /^C/.test(a) && b && /^Descrip/i.test(b.trim())) { enLineas = true; continue; }
        if (a && !cur.descripcion) cur.descripcion = a;   // el nombre va debajo de «Partida:»
        continue;
      }
      if (!a) continue;
      var c = b == null ? categoriaDelphin(a) : null;
      if (c) { cat = c; cur.subtotales[c] = num(f, 16) || 0; continue; }
      if (cat && b != null) {
        var k = txt(f, 10);
        cur.lineas.push({
          fila: r + 1, categoria: cat, codigo: a, descripcion: b.trim(), unidad: (txt(f, 9) || '').trim(),
          cuadrilla: k != null && /\d/.test(k) ? num(f, 10) : null, cantidad: num(f, 12), precio: num(f, 15), parcial: num(f, 16)
        });
      }
    }
    var e = {}, ENC = { PROYECTO: 'Proyecto', PRESUPUESTO: 'Sub Presupuesto', UBICACION: 'Ubicación' };
    for (var i = 0; i < Math.min(filas.length, 12); i++) { var k0 = (txt(filas[i], 0) || '').trim().toUpperCase(); if (ENC[k0] && txt(filas[i], 4)) e[ENC[k0]] = txt(filas[i], 4).trim(); }
    return { encabezado: e, apus: apus, subpartidas: {}, formato: 'delphin' };
  }

  /** Presupuesto de Delphin: ítem A, descripción B, unidad L, metrado M, precio O, total P; la fila del subpresupuesto se salta. */
  function leerPresupuestoDelphin(filas) {
    var e = {}, ENC = { PROYECTO: 'Proyecto', PRESUPUESTO: 'Sub Presupuesto', UBICACION: 'Ubicación' }, desde = -1, lista = [];
    for (var r = 0; r < filas.length; r++) {
      var a = (txt(filas[r], 0) || '').trim();
      if (/^Item$/i.test(a)) { desde = r + 1; break; }
      if (ENC[a.toUpperCase()] && txt(filas[r], 4)) e[ENC[a.toUpperCase()]] = txt(filas[r], 4).trim();
    }
    for (r = Math.max(desde, 0); r < filas.length; r++) {
      var f = filas[r], item = (txt(f, 0) || '').trim();
      if (!RX_ITEM.test(item)) continue;
      var und = txt(f, 11);
      lista.push({
        fila: r + 1, item: item, descripcion: (txt(f, 1) || '').trim(), unidad: und == null ? null : und.trim(),
        metrado: num(f, 12), precio: num(f, 14), parcial: und != null ? num(f, 15) : null, subtotal: null, total: und == null ? num(f, 15) : null,
        nivel: item.split('.').length, esPartida: und != null
      });
    }
    // Subpresupuesto: fila sin unidad seguida de otra con el mismo ítem («01 MANTENIMIENTO VIAL…» y «01 TRABAJOS…»).
    lista = lista.filter(function (x, i) {
      var sig = lista[i + 1];
      if (!x.esPartida && sig && sig.item === x.item) { if (!e['Sub Presupuesto']) e['Sub Presupuesto'] = x.descripcion; return false; }
      return true;
    });
    return { encabezado: e, filas: lista, formato: 'delphin' };
  }

  /** Presupuesto: lista de { item, descripcion, unidad, metrado, precio, parcial, total, nivel, esPartida } */
  function leerPresupuesto(XLSX, wb) {
    var filas = filasDe(XLSX, hojaDe(XLSX, wb)), lista = [];
    if (esDelphin(filas)) return leerPresupuestoDelphin(filas);
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
    return { encabezado: encabezado(filas, 9), filas: lista };
  }

  var TIEMPO = { hh: 1, hm: 1, he: 1 };

  /**
   * Explica cada línea del APU como lo calcula Power Cost, y comprueba lo que trae el Excel.
   * tipo: 'tiempo' (cuadrilla × 8 ÷ rendimiento), 'dia' (cuadrilla ÷ rendimiento), 'pmo' (% de la mano de obra),
   *       'pmt' (% de los demás materiales), 'directa' (cantidad escrita).
   */
  function explicar(apu) {
    var mo = apu.subtotales.mo || 0, difs = [];
    // Base del %MT: los materiales que no son porcentaje.
    var mt = r2(apu.lineas.filter(function (l) { return l.categoria === 'mat' && l.unidad.toLowerCase() !== '%mt'; })
                          .reduce(function (s, l) { return s + (l.parcial || 0); }, 0));
    apu.lineas.forEach(function (l) {
      var u = l.unidad.toLowerCase(), rend = apu.rendimiento;
      if (u === '%mo' || u === '%mt') {
        l.tipo = u === '%mo' ? 'pmo' : 'pmt';
        l.calcPrecio = u === '%mo' ? mo : mt;
        l.calcParcial = r2((l.cantidad || 0) * l.calcPrecio / 100);
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
    var subpartidas = {};
    if (apu && apu.subpartidas) Object.keys(apu.subpartidas).forEach(function (k) { subpartidas[k] = explicar(apu.subpartidas[k]); });
    apus.forEach(function (a) { a.lineas.forEach(function (l) { if (l.categoria === 'sp') l.subpartida = subpartidas[l.codigo] || null; }); });
    var porItem = {};
    apus.forEach(function (a) { porItem[a.item] = a; });
    var avisos = [];
    if (pre && apu && (apu.formato === 'delphin' || pre.formato === 'delphin')) {
      // Delphin: la numeración de los APU puede ser de otra versión del presupuesto → cada partida toma el APU de su
      // nombre (si hay varios, el de igual P.U.), con su propio ítem.
      var porNombre = {}, usados = [];
      apus.forEach(function (a) { (porNombre[nombre(a.descripcion)] = porNombre[nombre(a.descripcion)] || []).push(a); });
      porItem = {};
      var propios = [];
      pre.filas.forEach(function (f) {
        if (!f.esPartida) return;
        var c = porNombre[nombre(f.descripcion)] || [];
        var igualPu = function (a) { return a.cu != null && f.precio != null && Math.abs(a.cu - f.precio) < 0.005; };
        var a = c.filter(function (x) { return igualPu(x) && x.item === f.item; })[0] || c.filter(igualPu)[0] || c.filter(function (x) { return x.item === f.item; })[0] || c[0];
        if (!a) return;
        if (usados.indexOf(a) < 0) usados.push(a);
        var copia = Object.assign({}, a, { item: f.item, itemOrigen: a.item });
        porItem[f.item] = copia; propios.push(copia);
      });
      var sobran = apus.filter(function (a) { return usados.indexOf(a) < 0 && !usados.some(function (u) { return nombre(u.descripcion) === nombre(a.descripcion); }); });
      if (sobran.length) avisos.push(sobran.length + ' APU no están en el presupuesto: ' + sobran.slice(0, 6).map(function (a) { return a.item + ' ' + a.descripcion; }).join(', ') + (sobran.length > 6 ? '…' : ''));
      apus = propios;
    }
    var filas = pre ? pre.filas : apus.map(function (a) {
      return { item: a.item, descripcion: a.descripcion, unidad: a.unidad, metrado: null, precio: a.cu, parcial: null, nivel: 1, esPartida: true };
    });
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
      formato: (pre && pre.formato) || (apu && apu.formato) || 'powercost',
      filas: filas, apus: apus, porItem: porItem, insumos: insumos, subpartidas: subpartidas, cd: cd, cdTitulos: cdTitulos, avisos: avisos,
      conDiferencias: apus.filter(function (a) { return a.diferencias.length; })
    };
  }

  var Lector = { tipoDe: tipoDe, leerApu: leerApu, leerPresupuesto: leerPresupuesto, explicar: explicar, armar: armar, r2: r2, r4: r4, NOMBRE_CAT: NOMBRE_CAT };
  if (typeof module !== 'undefined' && module.exports) module.exports = Lector; else raiz.Lector = Lector;
})(this);
