/** @OnlyCurrentDoc */

/**
 * Banco de Semillas: conecta la página pública con esta planilla.
 *
 * - La página lee las pestañas Variedades, Movimientos y Config.
 * - Los alumnos registran salidas y devoluciones sin clave.
 * - Ingresos, ensayos, bajas, variedades y ajustes piden la clave de docente
 *   (pestaña Config, fila "Clave docente").
 * - La pestaña Inventario se recalcula sola con cada cambio.
 */

var HOJA_VAR = 'Variedades';
var HOJA_MOV = 'Movimientos';
var HOJA_INV = 'Inventario';
var HOJA_CFG = 'Config';

var COLS_VAR = [
  ['id', 'ID'], ['especie', 'Especie'], ['variedad', 'Variedad'], ['cientifico', 'Nombre científico'],
  ['familia', 'Familia'], ['grupo', 'Grupo'], ['unidad', 'Unidad'], ['vida', 'Vida útil (años)'],
  ['envases', 'Envases'], ['notas', 'Notas']
];
var COLS_MOV = [
  ['t', 'Fecha'], ['tipo', 'Tipo'], ['alumno', 'Alumno/a'], ['curso', 'Curso'], ['semilla', 'Semilla'],
  ['anio', 'Lote (año)'], ['cant', 'Cantidad'], ['unidad', 'Unidad'], ['pg', 'PG (%)'], ['pgFuente', 'Origen del PG'],
  ['destino', 'Destino'], ['obs', 'Observaciones'], ['tipoSem', 'Semilla devuelta'], ['es', 'Semillas puestas'],
  ['ge', 'Semillas germinadas'], ['sinPesar', 'Sin pesar'], ['cerrada', 'Cerrada'], ['id', 'ID'],
  ['vid', 'ID variedad'], ['salidaId', 'ID salida']
];
var COLS_INV = ['Especie', 'Variedad', 'Lote (año)', 'Cantidad', 'Unidad', 'PG (%)', 'Estado', 'Envases', 'Préstamos pendientes'];

var CONFIG_INICIAL = [['Dato', 'Valor'], ['Escuela', 'Escuela Agropecuaria'], ['WhatsApp del docente', "'5492344400874"], ['Clave docente', '']];

var TIPOS = { salida: 'Salida', devolucion: 'Devolución', ingreso: 'Ingreso', ensayo: 'Ensayo', baja: 'Baja' };
var FUENTES = { ensayo: 'Ensayo propio', dato: 'Dato del envase', estimado: 'Estimado', sin: 'Sin dato' };
var SEMILLA_DEV = { cosecha: 'Cosecha propia', sinusar: 'Sin usar' };
var SOLO_ALUMNOS = { salida: 1, devolucion: 1 };

/* ===== Entradas web ===== */

function doGet(e) {
  var p = (e && e.parameter) || {};
  var listo = conCaptura(function () { prepararPlanilla(); return { ok: true }; });
  if (!listo.ok) return responder(listo);
  if (p.accion === 'guardar' && p.datos) return responder(escribir(JSON.parse(p.datos)));
  return responder(conCaptura(function () { return { ok: true, estado: leerEstado() }; }));
}

function doPost(e) {
  var cuerpo = (e && e.postData && e.postData.contents) || '{}';
  var pedido;
  try { pedido = JSON.parse(cuerpo); } catch (err) { return responder({ ok: false, error: 'Pedido ilegible.' }); }
  var listo = conCaptura(function () { prepararPlanilla(); return { ok: true }; });
  if (!listo.ok) return responder(listo);
  return responder(escribir(pedido));
}

/* Crea las pestañas que falten. Se puede correr a mano desde el editor y es seguro repetirla. */
function prepararPlanilla() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  if (libro.getSheetByName(HOJA_VAR) && libro.getSheetByName(HOJA_MOV) && libro.getSheetByName(HOJA_INV) && libro.getSheetByName(HOJA_CFG)) return;
  if (!libro.getSheetByName(HOJA_VAR)) {
    var primera = libro.getSheets()[0];
    var cab = primera.getRange(1, 1, 1, Math.max(1, primera.getLastColumn())).getValues()[0];
    if (cab.map(norm).indexOf('especie') >= 0) primera.setName(HOJA_VAR);
    else crearHoja(libro, HOJA_VAR, [COLS_VAR.map(function (c) { return c[1]; })]);
  }
  if (!libro.getSheetByName(HOJA_MOV)) crearHoja(libro, HOJA_MOV, [COLS_MOV.map(function (c) { return c[1]; })]);
  if (!libro.getSheetByName(HOJA_INV)) crearHoja(libro, HOJA_INV, [COLS_INV]);
  if (!libro.getSheetByName(HOJA_CFG)) {
    var cfg = crearHoja(libro, HOJA_CFG, CONFIG_INICIAL);
    cfg.getRange(4, 2).setBackground('#FFF2CC').setNote('Escribí acá la clave para entrar al modo docente de la página. No la compartas con los alumnos.');
  }
  [HOJA_MOV, HOJA_INV, HOJA_VAR, HOJA_CFG].forEach(function (n, i) {
    var h = libro.getSheetByName(n);
    formatearEncabezado(h);
    libro.setActiveSheet(h);
    libro.moveActiveSheet(i + 1);
  });
  recalcularInventario();
}

function crearHoja(libro, nombre, filas) {
  var h = libro.insertSheet(nombre);
  h.getRange(1, 1, filas.length, filas[0].length).setValues(filas);
  return h;
}

function formatearEncabezado(h) {
  var ancho = Math.max(1, h.getLastColumn());
  h.getRange(1, 1, 1, ancho).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#2E5E3E');
  h.setFrozenRows(1);
  h.autoResizeColumns(1, ancho);
}

function responder(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function conCaptura(fn) {
  try { return fn(); } catch (err) { return { ok: false, error: String(err && err.message || err) }; }
}

/* ===== Menú y edición a mano ===== */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Banco de semillas')
    .addItem('Recalcular inventario', 'recalcularInventario')
    .addToUi();
}

function onEdit(e) {
  var nombre = e && e.range && e.range.getSheet().getName();
  if (nombre === HOJA_MOV || nombre === HOJA_VAR) recalcularInventario();
}

/* ===== Escritura ===== */

function escribir(pedido) {
  return conCaptura(function () {
    var op = pedido.op;
    var cfg = leerConfig();
    var esDocente = !!cfg.clave && String(pedido.clave || '') === cfg.clave;

    if (op === 'verificar') {
      if (!cfg.clave) return { ok: false, error: 'Todavía no hay clave de docente: escribila en la pestaña Config de la planilla.' };
      return esDocente ? { ok: true } : { ok: false, error: 'La clave no es correcta.' };
    }

    var soloAlumno = op === 'addMovs' && (pedido.movs || []).every(function (m) { return SOLO_ALUMNOS[m.tipo]; });
    if (!soloAlumno && !esDocente) return { ok: false, error: 'Esta acción necesita la clave de docente.' };

    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      if (op === 'addMovs') agregarMovs(pedido.movs || [], esDocente);
      else if (op === 'delMov') borrarFila(HOJA_MOV, pedido.id);
      else if (op === 'updMov') actualizarMov(pedido.id, pedido.campos || {});
      else if (op === 'setVar') guardarVariedad(pedido.v || {});
      else if (op === 'delVar') borrarVariedad(pedido.id);
      else if (op === 'config') guardarConfig(pedido.c || {});
      else return { ok: false, error: 'Acción desconocida.' };
      recalcularInventario();
    } finally {
      lock.releaseLock();
    }
    return { ok: true, estado: leerEstado() };
  });
}

function agregarMovs(movs, esDocente) {
  var hoja = hojaPorNombre(HOJA_MOV);
  var est = leerEstado();
  var ids = {};
  est.movs.forEach(function (m) { ids[m.id] = 1; });
  var vars = indicePorId(est.variedades);
  movs.forEach(function (m) {
    if (!m || !m.id || ids[m.id]) return;
    if (!TIPOS[m.tipo]) throw new Error('Tipo de movimiento desconocido.');
    if (!vars[m.vid]) throw new Error('La semilla ya no está en el banco.');
    if (!esDocente) {
      if (!String(m.alumno || '').trim()) throw new Error('Falta el nombre del alumno.');
      if (!(num(m.cant) > 0)) throw new Error('Falta la cantidad.');
      m.cerrada = false;
    }
    hoja.appendRow(filaMov(m, vars[m.vid], hoja));
    hoja.getRange(hoja.getLastRow(), 1).setNumberFormat('dd/mm/yyyy hh:mm');
    ids[m.id] = 1;
  });
}

function actualizarMov(id, campos) {
  var hoja = hojaPorNombre(HOJA_MOV);
  var datos = hoja.getDataRange().getValues();
  var col = mapaColumnas(datos[0], COLS_MOV);
  for (var r = 1; r < datos.length; r++) {
    if (String(datos[r][col.id]) === String(id)) {
      if ('cerrada' in campos && col.cerrada != null) hoja.getRange(r + 1, col.cerrada + 1).setValue(campos.cerrada ? 'Sí' : '');
      return;
    }
  }
  throw new Error('No encontré ese movimiento.');
}

function guardarVariedad(v) {
  if (!String(v.especie || '').trim()) throw new Error('Falta el nombre de la especie.');
  var hoja = hojaPorNombre(HOJA_VAR);
  var datos = hoja.getDataRange().getValues();
  var col = mapaColumnas(datos[0], COLS_VAR);
  var fila = datos[0].map(function () { return ''; });
  COLS_VAR.forEach(function (c) {
    if (col[c[0]] == null) return;
    var val = v[c[0]];
    fila[col[c[0]]] = val == null ? '' : (typeof val === 'number' ? val : texto(val));
  });
  for (var r = 1; r < datos.length; r++) {
    if (String(datos[r][col.id]) === String(v.id)) {
      hoja.getRange(r + 1, 1, 1, fila.length).setValues([fila]);
      return;
    }
  }
  hoja.appendRow(fila);
}

function borrarVariedad(id) {
  var usada = leerEstado().movs.some(function (m) { return m.vid === id; });
  if (usada) throw new Error('Esta variedad tiene movimientos: anulalos primero.');
  borrarFila(HOJA_VAR, id);
}

function borrarFila(nombreHoja, id) {
  var hoja = hojaPorNombre(nombreHoja);
  var datos = hoja.getDataRange().getValues();
  var col = mapaColumnas(datos[0], nombreHoja === HOJA_VAR ? COLS_VAR : COLS_MOV);
  for (var r = datos.length - 1; r >= 1; r--) {
    if (String(datos[r][col.id]) === String(id)) { hoja.deleteRow(r + 1); return; }
  }
  throw new Error('No encontré esa fila.');
}

function guardarConfig(c) {
  var hoja = hojaPorNombre(HOJA_CFG);
  var datos = hoja.getDataRange().getValues();
  var claves = { escuela: 'Escuela', wa: 'WhatsApp del docente' };
  Object.keys(claves).forEach(function (k) {
    if (!(k in c)) return;
    var valor = k === 'wa' ? String(c[k] || '').replace(/\D/g, '') : String(c[k] || '');
    for (var r = 0; r < datos.length; r++) {
      if (norm(datos[r][0]) === norm(claves[k])) {
        hoja.getRange(r + 1, 2).setValue(valor ? "'" + valor : '');
        return;
      }
    }
    hoja.appendRow([claves[k], valor ? "'" + valor : '']);
  });
}

/* ===== Lectura ===== */

function leerEstado() {
  var variedades = leerVariedades();
  var cfg = leerConfig();
  return {
    config: { escuela: cfg.escuela, wa: cfg.wa },
    variedades: variedades,
    movs: leerMovs(variedades)
  };
}

function leerConfig() {
  var datos = hojaPorNombre(HOJA_CFG).getDataRange().getValues();
  var cfg = { escuela: '', wa: '', clave: '' };
  datos.forEach(function (fila) {
    var k = norm(fila[0]), v = fila[1] == null ? '' : String(fila[1]).trim();
    if (k === 'escuela') cfg.escuela = v;
    else if (k === 'whatsapp del docente') cfg.wa = v.replace(/\D/g, '');
    else if (k === 'clave docente') cfg.clave = v;
  });
  return cfg;
}

function leerVariedades() {
  var datos = hojaPorNombre(HOJA_VAR).getDataRange().getValues();
  if (datos.length < 2) return [];
  var col = mapaColumnas(datos[0], COLS_VAR);
  var out = [];
  for (var r = 1; r < datos.length; r++) {
    var f = datos[r];
    var especie = celda(f, col.especie);
    if (!especie) continue;
    out.push({
      id: celda(f, col.id) || 'v-fila' + (r + 1),
      especie: especie,
      variedad: celda(f, col.variedad),
      cientifico: celda(f, col.cientifico),
      familia: celda(f, col.familia),
      grupo: celda(f, col.grupo),
      unidad: celda(f, col.unidad) || 'g',
      vida: num(f[col.vida]) || null,
      envases: celda(f, col.envases),
      notas: celda(f, col.notas)
    });
  }
  return out;
}

function leerMovs(variedades) {
  var datos = hojaPorNombre(HOJA_MOV).getDataRange().getValues();
  if (datos.length < 2) return [];
  var col = mapaColumnas(datos[0], COLS_MOV);
  var porNombre = {};
  variedades.forEach(function (v) { porNombre[norm(nombreVar(v))] = v.id; });
  var tipos = invertir(TIPOS), fuentes = invertir(FUENTES), sems = invertir(SEMILLA_DEV);
  var out = [];
  for (var r = 1; r < datos.length; r++) {
    var f = datos[r];
    var tipo = tipos[norm(f[col.tipo])];
    if (!tipo) continue;
    var vid = celda(f, col.vid) || porNombre[norm(f[col.semilla])];
    if (!vid) continue;
    var t = f[col.t];
    var fecha = t instanceof Date ? t : (t ? new Date(t) : null);
    var m = {
      id: celda(f, col.id) || 'fila-' + hash([fecha && fecha.getTime(), tipo, vid, f[col.cant], f[col.alumno]].join('|')),
      t: fecha && !isNaN(fecha.getTime()) ? fecha.toISOString() : '1970-01-01T00:00:00.000Z',
      tipo: tipo,
      vid: vid,
      anio: num(f[col.anio]) || 0,
      cant: num(f[col.cant]) || 0,
      pg: f[col.pg] === '' || f[col.pg] == null || isNaN(num(f[col.pg])) ? null : num(f[col.pg]),
      alumno: celda(f, col.alumno),
      curso: celda(f, col.curso),
      destino: celda(f, col.destino),
      obs: celda(f, col.obs)
    };
    var fuente = fuentes[norm(f[col.pgFuente])];
    if (fuente) m.pgFuente = fuente;
    var sem = sems[norm(f[col.tipoSem])];
    if (sem) m.tipoSem = sem;
    if (num(f[col.es]) > 0) m.es = num(f[col.es]);
    if (f[col.ge] !== '' && f[col.ge] != null && !isNaN(num(f[col.ge]))) m.ge = num(f[col.ge]);
    if (esSi(f[col.sinPesar])) m.sinPesar = true;
    if (esSi(f[col.cerrada])) m.cerrada = true;
    var sal = celda(f, col.salidaId);
    if (sal) m.salidaId = sal;
    out.push(m);
  }
  return out;
}

/* ===== Inventario ===== */

function recalcularInventario() {
  var est = leerEstado();
  var vars = indicePorId(est.variedades);
  var calc = calcular(est.movs);
  var filas = [];
  est.variedades.slice().sort(compararVar).forEach(function (v) {
    var lotes = calc.lotes.filter(function (L) { return L.vid === v.id && L.stock > 0; })
      .sort(function (a, b) { return b.anio - a.anio; });
    var pend = calc.pendientes.filter(function (s) { return s.vid === v.id; }).length;
    if (!lotes.length) {
      var sinPesar = !!v.envases && !calc.pesadas[v.id];
      filas.push([v.especie, v.variedad, '', '', v.unidad, '', sinPesar ? 'Sin pesar · se puede llevar' : 'Sin existencias', v.envases, pend || '']);
      return;
    }
    lotes.forEach(function (L, i) {
      filas.push([v.especie, v.variedad, L.anio || 's/f', redondear(L.stock), v.unidad, L.pg == null ? '' : redondear(L.pg),
        nivel(L.pg), i === 0 ? v.envases : '', i === 0 ? (pend || '') : '']);
    });
  });
  var hoja = hojaPorNombre(HOJA_INV);
  hoja.clearContents();
  hoja.getRange(1, 1, 1, COLS_INV.length).setValues([COLS_INV]);
  if (filas.length) hoja.getRange(2, 1, filas.length, COLS_INV.length).setValues(filas);
  hoja.getRange(1, 1).setNote('Esta pestaña se calcula sola a partir de Variedades y Movimientos. No escribas acá: se borra en el próximo cambio.');
  return filas.length;
}

/* Misma cuenta que la página: lotes por variedad y año, PG ponderado, salidas sin pesar sin descontar. */
function calcular(movs) {
  var lotes = {}, salidas = {}, pesadas = {}, salSP = {};
  var ms = movs.slice().sort(function (a, b) { return a.t < b.t ? -1 : a.t > b.t ? 1 : (a.id < b.id ? -1 : 1); });
  ms.forEach(function (m) { if (m.tipo === 'salida' && m.sinPesar) salSP[m.id] = 1; });
  ms.forEach(function (m) {
    if (m.tipo === 'salida' && m.sinPesar) { salidas[m.id] = { m: m, devuelto: 0 }; return; }
    if (m.tipo === 'devolucion' && m.tipoSem === 'sinusar' && salSP[m.salidaId] && !pesadas[m.vid]) return;
    if (m.tipo === 'ingreso') pesadas[m.vid] = true;
    var k = m.vid + '|' + m.anio;
    var L = lotes[k] || (lotes[k] = { vid: m.vid, anio: +m.anio || 0, stock: 0, pg: null });
    var c = +m.cant || 0;
    if (m.tipo === 'ingreso' || m.tipo === 'devolucion') {
      var pin = m.pg == null || m.pg === '' ? null : +m.pg;
      var prev = Math.max(0, L.stock);
      if (pin !== null) L.pg = (prev <= 0 || L.pg === null) ? pin : (prev * L.pg + c * pin) / (prev + c);
      L.stock += c;
    } else if (m.tipo === 'salida' || m.tipo === 'baja') {
      L.stock -= c;
    } else if (m.tipo === 'ensayo') {
      L.pg = +m.pg;
    }
    if (m.tipo === 'salida') salidas[m.id] = { m: m, devuelto: 0 };
  });
  ms.forEach(function (m) {
    if (m.tipo === 'devolucion' && m.salidaId && salidas[m.salidaId]) salidas[m.salidaId].devuelto += +m.cant || 0;
  });
  var pendientes = [];
  Object.keys(salidas).forEach(function (k) {
    var s = salidas[k];
    if (!s.m.cerrada && s.devuelto < (+s.m.cant || 0)) pendientes.push({ vid: s.m.vid, id: s.m.id });
  });
  return { lotes: Object.keys(lotes).map(function (k) { return lotes[k]; }), pendientes: pendientes, pesadas: pesadas };
}

/* ===== Utilidades ===== */

function hojaPorNombre(nombre) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre);
  if (!hoja) throw new Error('Falta la pestaña "' + nombre + '" en la planilla.');
  return hoja;
}

function filaMov(m, v, hoja) {
  var encabezado = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  var col = mapaColumnas(encabezado, COLS_MOV);
  var fila = encabezado.map(function () { return ''; });
  var valores = {
    t: new Date(m.t || Date.now()),
    tipo: TIPOS[m.tipo],
    alumno: texto(m.alumno), curso: texto(m.curso),
    semilla: nombreVar(v),
    anio: +m.anio ? +m.anio : '',
    cant: num(m.cant) || 0,
    unidad: v.unidad || '',
    pg: m.pg == null || m.pg === '' ? '' : num(m.pg),
    pgFuente: FUENTES[m.pgFuente] || '',
    destino: texto(m.destino), obs: texto(m.obs),
    tipoSem: SEMILLA_DEV[m.tipoSem] || '',
    es: m.es == null ? '' : num(m.es),
    ge: m.ge == null ? '' : num(m.ge),
    sinPesar: m.sinPesar ? 'Sí' : '',
    cerrada: m.cerrada ? 'Sí' : '',
    id: texto(m.id), vid: texto(m.vid), salidaId: texto(m.salidaId)
  };
  Object.keys(valores).forEach(function (k) { if (col[k] != null) fila[col[k]] = valores[k]; });
  return fila;
}

function mapaColumnas(encabezado, cols) {
  var porTitulo = {};
  encabezado.forEach(function (t, i) { porTitulo[norm(t)] = i; });
  var col = {};
  cols.forEach(function (c) { if (porTitulo[norm(c[1])] != null) col[c[0]] = porTitulo[norm(c[1])]; });
  if (col.id == null) throw new Error('Falta la columna "ID" en la planilla.');
  return col;
}

function celda(fila, i) { return i == null || fila[i] == null ? '' : String(fila[i]).trim(); }
function num(v) { if (typeof v === 'number') return v; var n = parseFloat(String(v == null ? '' : v).replace(',', '.')); return isNaN(n) ? NaN : n; }
function texto(v) { var s = v == null ? '' : String(v); return /^[=+\-@]/.test(s) ? "'" + s : s; }
function esSi(v) { var s = norm(v); return s === 'si' || s === 'true' || s === 'x'; }
function norm(s) { return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim(); }
function invertir(o) { var r = {}; Object.keys(o).forEach(function (k) { r[norm(o[k])] = k; r[norm(k)] = k; }); return r; }
function indicePorId(lista) { var r = {}; lista.forEach(function (x) { r[x.id] = x; }); return r; }
function nombreVar(v) { return v ? v.especie + (v.variedad ? ' ' + v.variedad : '') : ''; }
function compararVar(a, b) { return nombreVar(a).localeCompare(nombreVar(b), 'es', { sensitivity: 'base' }); }
function redondear(x) { return Math.round(x * 10) / 10; }
function nivel(pg) { return pg == null ? 'Sin dato de PG' : pg >= 80 ? 'Bueno' : pg >= 60 ? 'Regular' : 'Bajo'; }
function hash(s) { var h = 0; for (var i = 0; i < s.length; i++) { h = ((h << 5) - h + s.charCodeAt(i)) | 0; } return (h >>> 0).toString(36); }
