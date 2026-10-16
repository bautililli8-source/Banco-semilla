/** @OnlyCurrentDoc */

/**
 * Sillas, tablones y caballetes: conecta la página pública con esta planilla.
 *
 * - La página lee y guarda acá. Cada retiro, devolución, rotura o reparación
 *   queda registrado al instante, sin cargar nada a mano.
 * - Cada registro manda un mail de aviso a la dirección de la pestaña Config.
 * - Altas, muebles, borrar registros y ajustes piden la clave de docente
 *   (pestaña Config, fila "Clave docente").
 * - Las pestañas Inventario y Retirado ahora se recalculan solas.
 * - La columna oculta "Datos" de Movimientos es la que usa la página: no la edites.
 */

var ZONA = 'America/Argentina/Buenos_Aires';
var H_INV = 'Inventario', H_RET = 'Retirado ahora', H_MOV = 'Movimientos', H_MUE = 'Muebles', H_CFG = 'Config';
var CAB_MOV = ['Fecha', 'Tipo', 'Quién', 'Curso, área o institución', 'Qué', 'Dónde', 'Teléfono', 'Vuelve', 'Falla u origen', 'Observaciones', 'ID', 'Datos'];
var CAB_MUE = ['ID', 'Nombre', 'Plural', 'Clase', 'Dónde se guardan', 'Si se rompe', 'Notas'];
var CAB_INV = ['Mueble', 'Total', 'En el depósito', 'Retirados', 'Para reparar', 'Descartados por rotura'];
var CAB_RET = ['Desde', 'Quién', 'Curso, área o institución', 'Qué', 'Dónde', 'Teléfono', 'Vuelve', 'Vencido'];
var MUEBLES_INICIALES = [
  ['t-sillas', 'Silla', 'Sillas', 'silla', 'Depósito', 'Se descarta', 'De plástico: si se rompen, se descartan.'],
  ['t-tablones', 'Tablón', 'Tablones', 'tablon', 'Depósito', 'Se repara', 'De madera: se reparan.'],
  ['t-cab-madera', 'Caballete de madera', 'Caballetes de madera', 'caballete', 'Depósito', 'Se repara', 'Se reparan.'],
  ['t-cab-hierro', 'Caballete de hierro', 'Caballetes de hierro', 'caballete', 'Depósito', 'Se repara', 'Se reparan.']
];
var CONFIG_INICIAL = [
  ['Dato', 'Valor'],
  ['Escuela', 'CFR Saladillo'],
  ['WhatsApp del docente', "'5492344400874"],
  ['Clave docente', ''],
  ['Mail para avisos', ''],
  ['Avisar por mail', 'Sí'],
  ['Dirección de la página', 'https://bautililli8-source.github.io/Banco-semilla/mobiliario/']
];
var TIPO_TXT = { prestamo: 'Retiro', devolucion: 'Devolución', rotura: 'Rotura', reparacion: 'Reparación', baja: 'Baja', ingreso: 'Alta', ajuste: 'Recuento' };
var PUBLICOS = { prestamo: 1, devolucion: 1, rotura: 1, reparacion: 1 };
var CLASES = ['silla', 'tablon', 'caballete', 'otro'];

/* ===== Entradas web ===== */

function doGet(e) {
  return responder(conCaptura(function () {
    prepararPlanilla();
    return { ok: true, estado: leerEstado(false) };
  }));
}

function doPost(e) {
  var pedido;
  try { pedido = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return responder({ ok: false, error: 'Pedido ilegible.' }); }
  return responder(conCaptura(function () { prepararPlanilla(); return atender(pedido || {}); }));
}

function responder(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function conCaptura(fn) {
  try { return fn(); } catch (err) { return { ok: false, error: String((err && err.message) || err) }; }
}

/* ===== Pedidos de la página ===== */

function atender(p) {
  var cfg = leerConfig();
  var esDoc = !!cfg.clave && String(p.clave || '') === cfg.clave;

  if (p.op === 'verificar') {
    if (!cfg.clave) return { ok: false, error: 'Todavía no hay clave de docente: escribila en la pestaña Config de la planilla.' };
    return esDoc ? { ok: true } : { ok: false, error: 'La clave no es correcta.' };
  }
  if (p.op === 'estado') return { ok: true, estado: leerEstado(esDoc) };

  var publico = p.op === 'addMovs' && Array.isArray(p.movs) && p.movs.every(function (m) { return m && PUBLICOS[m.tipo]; });
  if (!publico && !esDoc) return { ok: false, error: 'Esta acción necesita la clave de docente.' };

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  var r;
  try {
    if (p.op === 'addMovs') r = agregarMovs(p.movs || [], esDoc);
    else if (p.op === 'delMov') r = borrarMov(String(p.id || ''));
    else if (p.op === 'setTipo') r = guardarMueble(p.t || {});
    else if (p.op === 'delTipo') r = borrarMueble(String(p.id || ''));
    else if (p.op === 'config') r = guardarConfig(p.c || {});
    else if (p.op === 'importar') r = importar(p.estado || {}, !!p.reemplazar);
    else r = { ok: false, error: 'Acción desconocida.' };
    if (r.ok) recalcular();
  } finally {
    lock.releaseLock();
  }
  if (!r.ok) return r;
  if (!esDoc && r.avisar && r.avisar.length) avisar(r.avisar, cfg);
  return { ok: true, estado: leerEstado(esDoc) };
}

/* ===== Lectura ===== */

function hoja(nombre) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre); }

function leerConfig() {
  var c = { avisar: true };
  hoja(H_CFG).getDataRange().getValues().forEach(function (f) {
    var k = norm(f[0]), v = String(f[1] == null ? '' : f[1]).trim();
    if (k === 'escuela') c.escuela = v;
    else if (k.indexOf('whatsapp') === 0) c.wa = v.replace(/\D/g, '');
    else if (k === 'clave docente') c.clave = v;
    else if (k.indexOf('mail para avisos') === 0) c.mail = v;
    else if (k.indexOf('avisar por mail') === 0) c.avisar = norm(v) !== 'no';
    else if (k.indexOf('direccion de la pagina') === 0) c.url = v;
  });
  return c;
}

function leerMuebles() {
  var h = hoja(H_MUE), n = h.getLastRow();
  if (n < 2) return [];
  return h.getRange(2, 1, n - 1, CAB_MUE.length).getValues().filter(function (f) {
    return String(f[0]).trim() && String(f[2]).trim();
  }).map(function (f) {
    var cl = norm(f[3]);
    return {
      id: String(f[0]).trim(), nombre: String(f[1]).trim() || String(f[2]).trim(), plural: String(f[2]).trim(),
      clase: CLASES.indexOf(cl) >= 0 ? cl : 'otro', guardado: String(f[4]).trim(),
      reparable: norm(f[5]).indexOf('descarta') < 0, notas: String(f[6]).trim()
    };
  });
}

function leerMovs() {
  var h = hoja(H_MOV), n = h.getLastRow(), col = CAB_MOV.length, out = [];
  if (n < 2) return out;
  h.getRange(2, 1, n - 1, col).getValues().forEach(function (f, i) {
    var raw = f[col - 1];
    if (!raw) return;
    try { var m = JSON.parse(raw); if (m && m.id && m.tipo) out.push({ fila: i + 2, m: m }); } catch (e) {}
  });
  return out;
}

function leerEstado(conTel) {
  var cfg = leerConfig();
  var movs = leerMovs().map(function (x) {
    if (conTel || x.m.tel == null) return x.m;
    var c = JSON.parse(JSON.stringify(x.m)); delete c.tel; return c;
  });
  return { tipos: leerMuebles(), movs: movs, config: { escuela: cfg.escuela || '', wa: cfg.wa || '', url: cfg.url || '' } };
}

/* ===== Escritura ===== */

function agregarMovs(nuevos, esDoc) {
  var tipos = leerMuebles(), TP = mapaTipos(tipos);
  var movs = leerMovs().map(function (x) { return x.m; }), ids = {};
  movs.forEach(function (m) { ids[m.id] = 1; });
  var aceptados = [], ahora = new Date().toISOString(), dia = hoy();
  var falla = function (t) { return { ok: false, error: t }; };

  for (var k = 0; k < nuevos.length; k++) {
    var m = limpiar(nuevos[k]);
    if (!m) return falla('Hay un registro con datos incompletos.');
    if (ids[m.id]) continue;
    if (!TIPO_TXT[m.tipo]) return falla('Tipo de registro desconocido.');
    if (!PUBLICOS[m.tipo] && !esDoc) return falla('Esta acción necesita la clave de docente.');
    if (!esDoc || !m.t) m.t = ahora;

    var C = calc({ tipos: tipos, movs: movs }, dia), x, i;
    if (m.tipo === 'reparacion' && !C.R[m.roturaId]) return falla('Esa rotura ya está resuelta o no existe.');
    if ((m.tipo === 'reparacion' || m.tipo === 'baja') && m.roturaId && C.R[m.roturaId]) m.tid = C.R[m.roturaId].m.tid;
    var its = itemsDe(m);
    if (m.tipo !== 'ajuste') {
      if (!its.length) return falla('Indicá qué muebles y cuántos.');
      for (var j = 0; j < its.length; j++) {
        if (!TP[its[j].tid]) return falla('Uno de los muebles ya no está en la lista.');
        if (!(its[j].cant > 0)) return falla('Las cantidades tienen que ser números enteros mayores que cero.');
      }
    }
    if (m.tipo === 'prestamo') {
      if (!m.quien) return falla('Falta el nombre de quien retira.');
      if (!m.lugar) return falla('Falta adónde se lleva.');
      for (i = 0; i < its.length; i++) {
        x = Math.max(0, C.T[its[i].tid].disp);
        if (its[i].cant > x) return falla('No alcanza: en el depósito hay ' + x + ' ' + nomT(TP, its[i].tid, x) + '.');
      }
    } else if (m.tipo === 'devolucion') {
      var P = C.P[m.prestamoId];
      if (!P || !P.abierto) return falla('Ese retiro ya fue devuelto o no existe.');
      for (i = 0; i < its.length; i++) {
        var pend = Math.max(0, P.pend[its[i].tid] || 0);
        if (its[i].cant > pend) return falla('De ese retiro faltan volver ' + pend + ' ' + nomT(TP, its[i].tid, pend) + ', no más.');
        if ((its[i].rotas || 0) > its[i].cant) return falla('Los rotos no pueden ser más que los que se devuelven.');
      }
    } else if (m.tipo === 'rotura') {
      if (TP[m.tid].reparable === false) m.descarte = true; else delete m.descarte;
      x = Math.max(0, C.T[m.tid].disp);
      if (m.cant > x) return falla('En el depósito hay ' + x + ' ' + nomT(TP, m.tid, x) + '. Si lo que se rompió está retirado, avisalo al devolverlo.');
    } else if (m.tipo === 'reparacion') {
      var R = C.R[m.roturaId];
      if (!R) return falla('Esa rotura ya está resuelta o no existe.');
      if (m.cant > R.abiertas) return falla('De esa rotura quedan ' + R.abiertas + ' sin arreglar.');
      m.tid = R.m.tid;
    } else if (m.tipo === 'baja') {
      if (m.roturaId) {
        var R2 = C.R[m.roturaId];
        if (!R2 || m.cant > R2.abiertas) return falla('Esa rotura ya está resuelta.');
        m.tid = R2.m.tid;
      } else if (m.cant > Math.max(0, C.T[m.tid].disp)) {
        return falla('En el depósito no hay tantos para dar de baja.');
      }
    }
    movs.push(m); aceptados.push(m); ids[m.id] = 1;
  }

  if (aceptados.length) {
    var h = hoja(H_MOV), filas = aceptados.map(function (m) { return filaMov(m, TP); });
    h.getRange(h.getLastRow() + 1, 1, filas.length, CAB_MOV.length).setValues(filas);
  }
  return { ok: true, avisar: aceptados };
}

function borrarMov(id) {
  var todos = leerMovs(), ids = {}, cambio = true;
  ids[id] = 1;
  while (cambio) {
    cambio = false;
    todos.forEach(function (x) {
      var m = x.m;
      if (!ids[m.id] && (ids[m.prestamoId] || ids[m.roturaId] || ids[m.origen])) { ids[m.id] = 1; cambio = true; }
    });
  }
  var h = hoja(H_MOV);
  todos.filter(function (x) { return ids[x.m.id]; }).map(function (x) { return x.fila; })
    .sort(function (a, b) { return b - a; }).forEach(function (f) { h.deleteRow(f); });
  return { ok: true };
}

function filaMueble(t) {
  return [
    String(t.id || '').trim() || 't-' + Utilities.getUuid().slice(0, 8),
    texto(t.nombre, 60) || texto(t.plural, 60), texto(t.plural, 60),
    CLASES.indexOf(t.clase) >= 0 ? t.clase : 'otro', texto(t.guardado, 80),
    t.reparable === false ? 'Se descarta' : 'Se repara', texto(t.notas, 200)
  ];
}

function guardarMueble(t) {
  if (!texto(t.plural, 60)) return { ok: false, error: 'Falta el nombre del mueble.' };
  var fila = filaMueble(t), h = hoja(H_MUE), n = h.getLastRow();
  if (n >= 2) {
    var ids = h.getRange(2, 1, n - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === fila[0]) { h.getRange(i + 2, 1, 1, fila.length).setValues([fila]); return { ok: true }; }
    }
  }
  h.appendRow(fila);
  return { ok: true };
}

function borrarMueble(id) {
  var usado = leerMovs().some(function (x) { return itemsDe(x.m).some(function (i) { return i.tid === id; }); });
  if (usado) return { ok: false, error: 'Ese mueble tiene movimientos anotados.' };
  var h = hoja(H_MUE), n = h.getLastRow();
  if (n >= 2) {
    var ids = h.getRange(2, 1, n - 1, 1).getValues();
    for (var i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]).trim() === id) h.deleteRow(i + 2);
  }
  return { ok: true };
}

function guardarConfig(c) {
  var h = hoja(H_CFG), v = h.getDataRange().getValues();
  var poner = function (prefijo, etiqueta, valor) {
    for (var i = 0; i < v.length; i++) {
      if (norm(v[i][0]).indexOf(prefijo) === 0) { h.getRange(i + 1, 2).setValue(valor); return; }
    }
    h.appendRow([etiqueta, valor]);
  };
  if (c.escuela != null) poner('escuela', 'Escuela', texto(c.escuela, 120));
  if (c.wa != null) poner('whatsapp', 'WhatsApp del docente', "'" + String(c.wa).replace(/\D/g, ''));
  if (c.url != null) poner('direccion de la pagina', 'Dirección de la página', texto(c.url, 300));
  return { ok: true };
}

/* Pasa a la planilla los datos que tenía la página. Solo con la planilla vacía. */
function importar(est, reemplazar) {
  var h = hoja(H_MOV);
  if (leerMovs().length && !reemplazar) return { ok: false, error: 'La planilla ya tiene movimientos.' };
  if (Array.isArray(est.tipos) && est.tipos.length) {
    var hm = hoja(H_MUE), filasM = est.tipos.map(filaMueble);
    hm.clearContents();
    hm.getRange(1, 1, 1, CAB_MUE.length).setValues([CAB_MUE]);
    hm.getRange(2, 1, filasM.length, CAB_MUE.length).setValues(filasM);
    formatear(hm);
  }
  var TP = mapaTipos(leerMuebles());
  var movs = (est.movs || []).map(limpiar).filter(function (m) {
    return m && TIPO_TXT[m.tipo] && itemsDe(m).every(function (i) { return TP[i.tid]; });
  });
  movs.forEach(function (m) { if (m.tipo === 'rotura' && TP[m.tid] && TP[m.tid].reparable === false) m.descarte = true; });
  if (reemplazar && h.getLastRow() > 1) h.deleteRows(2, h.getLastRow() - 1);
  if (movs.length) {
    var filas = movs.map(function (m) { return filaMov(m, TP); });
    h.getRange(h.getLastRow() + 1, 1, filas.length, CAB_MOV.length).setValues(filas);
  }
  var c = est.config || {};
  guardarConfig({ escuela: c.escuela || null, wa: c.wa || null, url: c.url || null });
  return { ok: true };
}

/* ===== Planilla ===== */

/* Crea las pestañas que falten. Se puede correr a mano desde el editor y es seguro repetirla. */
function prepararPlanilla() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var orden = [H_INV, H_RET, H_MOV, H_MUE, H_CFG];
  var falta = orden.filter(function (n) { return !libro.getSheetByName(n); });
  if (!falta.length) return;
  var libre = libro.getSheets().filter(function (h) { return h.getLastRow() === 0; })[0] || null;
  falta.forEach(function (n) {
    var h;
    if (libre) { h = libre; h.setName(n); libre = null; } else h = libro.insertSheet(n);
    if (n === H_MOV) {
      h.getRange(1, 1, 1, CAB_MOV.length).setValues([CAB_MOV]);
      h.hideColumns(CAB_MOV.length);
    } else if (n === H_MUE) {
      h.getRange(1, 1, 1, CAB_MUE.length).setValues([CAB_MUE]);
      h.getRange(2, 1, MUEBLES_INICIALES.length, CAB_MUE.length).setValues(MUEBLES_INICIALES);
    } else if (n === H_CFG) {
      var filas = CONFIG_INICIAL.map(function (f) { return f.slice(); });
      try { filas[4][1] = Session.getEffectiveUser().getEmail() || ''; } catch (e) {}
      h.getRange(1, 1, filas.length, 2).setValues(filas);
      h.getRange(4, 2).setBackground('#FFF2CC').setNote('Escribí acá la clave para entrar al modo docente de la página. No la compartas.');
    } else if (n === H_INV) {
      h.getRange(1, 1, 1, CAB_INV.length).setValues([CAB_INV]);
    } else if (n === H_RET) {
      h.getRange(1, 1, 1, CAB_RET.length).setValues([CAB_RET]);
    }
    formatear(h);
  });
  orden.forEach(function (n, i) { libro.setActiveSheet(libro.getSheetByName(n)); libro.moveActiveSheet(i + 1); });
  recalcular();
}

function formatear(h) {
  var ancho = Math.max(1, h.getLastColumn());
  h.getRange(1, 1, 1, ancho).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#24506C');
  h.setFrozenRows(1);
}

function escribirTabla(h, cab, filas) {
  h.clearContents();
  h.getRange(1, 1, 1, cab.length).setValues([cab]);
  if (filas.length) h.getRange(2, 1, filas.length, cab.length).setValues(filas);
  formatear(h);
}

function recalcular() {
  var tipos = leerMuebles(), TP = mapaTipos(tipos);
  var C = calc({ tipos: tipos, movs: leerMovs().map(function (x) { return x.m; }) }, hoy());
  escribirTabla(hoja(H_INV), CAB_INV, tipos.map(function (t) {
    var x = C.T[t.id]; return [t.plural, x.total, x.disp, x.prestado, x.rotas, x.desc];
  }));
  escribirTabla(hoja(H_RET), CAB_RET, C.abiertos.map(function (p) {
    return [new Date(p.m.t), p.m.quien || '', p.m.rol || '', itemsTxt(pendItems(p), TP), p.m.lugar || '',
      p.m.tel ? "'" + p.m.tel : '', p.m.hasta ? fechaCorta(p.m.hasta) : '', p.vencido ? 'Sí' : ''];
  }));
}

function filaMov(m, TP) {
  var tipo = m.descarte ? 'Rotura (se descarta)' : (TIPO_TXT[m.tipo] || m.tipo);
  var que = itemsTxt(itemsDe(m), TP);
  if (m.tipo === 'devolucion') {
    var r = (m.items || []).filter(function (i) { return i.rotas > 0; }).map(function (i) { return { tid: i.tid, cant: i.rotas }; });
    if (r.length) que += ' (vienen rotos: ' + itemsTxt(r, TP) + ')';
  }
  var origen = m.tipo === 'ingreso' ? (m.origen || '') : (m.falla || m.motivo || '');
  return [new Date(m.t), tipo, m.quien || '', m.rol || '', que, m.lugar || '', m.tel ? "'" + m.tel : '',
    m.hasta ? fechaCorta(m.hasta) : '', origen, m.obs || m.detalle || '', m.id, JSON.stringify(m)];
}

/* ===== Avisos por mail ===== */

function avisar(movs, cfg) {
  if (cfg.avisar === false || !cfg.mail) return;
  try {
    var TP = mapaTipos(leerMuebles());
    var prin = movs.filter(function (m) { return !(m.tipo === 'rotura' && m.origen); })[0] || movs[0];
    var asunto = 'Mobiliario: ' + (TIPO_TXT[prin.tipo] || prin.tipo).toLowerCase() + ' de ' + itemsTxt(itemsDe(prin), TP) + (prin.quien ? ' (' + prin.quien + ')' : '');
    var l = [];
    movs.forEach(function (m) { l = l.concat(textoMov(m, TP)); l.push(''); });
    l.push('Ya quedó guardado en la planilla: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl());
    if (cfg.url) l.push('Página: ' + cfg.url);
    MailApp.sendEmail(cfg.mail, asunto, l.join('\n'));
  } catch (e) {}
}

function textoMov(m, TP) {
  var l = [], quien = (m.quien || 'Sin nombre') + (m.rol ? ' (' + m.rol + ')' : '');
  if (m.tipo === 'prestamo') {
    l.push('RETIRO: ' + itemsTxt(m.items, TP), 'Quién: ' + quien, 'Adónde: ' + (m.lugar || ''));
    if (m.tel) l.push('Teléfono: ' + m.tel);
    if (m.hasta) l.push('Devuelve el ' + fechaCorta(m.hasta));
  } else if (m.tipo === 'devolucion') {
    l.push('DEVOLUCIÓN: ' + itemsTxt(m.items, TP), 'Quién: ' + quien);
  } else if (m.tipo === 'rotura') {
    l.push((m.origen ? 'Vino roto: ' : 'ROTURA: ') + m.cant + ' ' + nomT(TP, m.tid, m.cant) + ' · ' + (m.falla || '') + (m.descarte ? ' (se descarta)' : ' (pasa a reparar)'));
    if (!m.origen) l.push('Quién avisa: ' + quien, 'Dónde está: ' + (m.lugar || ''));
  } else if (m.tipo === 'reparacion') {
    l.push('REPARACIÓN: ' + m.cant + ' ' + nomT(TP, m.tid, m.cant) + ' quedaron arreglados', 'Quién: ' + quien);
  }
  if (m.obs) l.push('Observaciones: ' + m.obs);
  if (m.detalle) l.push('Detalle: ' + m.detalle);
  return l;
}

/* ===== Menú ===== */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Mobiliario').addItem('Recalcular inventario', 'recalcular').addToUi();
}

function onEdit(e) {
  var n = e && e.range && e.range.getSheet().getName();
  if (n === H_MUE) recalcular();
}

/* ===== Cálculo (el mismo que usa la página) ===== */

function calc(st, dia) {
  var tp = {}, T = {};
  st.tipos.forEach(function (t) { tp[t.id] = t; T[t.id] = { total: 0, prestado: 0, rotas: 0, desc: 0 }; });
  function S(tid) { return T[tid] || (T[tid] = { total: 0, prestado: 0, rotas: 0, desc: 0, huerfano: true }); }
  var ms = st.movs.slice().sort(cmpMov), P = {}, R = {};
  ms.forEach(function (m) {
    var c, k = +m.cant || 0;
    if (m.tipo === 'ingreso') { (m.items || []).forEach(function (i) { S(i.tid).total += +i.cant || 0; }); }
    else if (m.tipo === 'ajuste') { S(m.tid).total += k; }
    else if (m.tipo === 'prestamo') {
      var p = P[m.id] = { m: m, pend: {}, devs: [] };
      (m.items || []).forEach(function (i) { var q = +i.cant || 0; p.pend[i.tid] = (p.pend[i.tid] || 0) + q; S(i.tid).prestado += q; });
    } else if (m.tipo === 'devolucion') {
      var p2 = P[m.prestamoId]; if (!p2) return;
      (m.items || []).forEach(function (i) { c = Math.min(+i.cant || 0, p2.pend[i.tid] || 0); p2.pend[i.tid] = (p2.pend[i.tid] || 0) - c; S(i.tid).prestado -= c; });
      p2.devs.push(m);
    } else if (m.tipo === 'rotura') {
      if (m.descarte) { S(m.tid).total -= k; S(m.tid).desc += k; return; }
      R[m.id] = { m: m, abiertas: k }; S(m.tid).rotas += k;
    } else if (m.tipo === 'reparacion') {
      var r = R[m.roturaId]; if (!r) return;
      c = Math.min(k, r.abiertas); r.abiertas -= c; S(r.m.tid).rotas -= c;
    } else if (m.tipo === 'baja') {
      if (m.roturaId && R[m.roturaId]) { var r2 = R[m.roturaId]; c = Math.min(k, r2.abiertas); r2.abiertas -= c; S(m.tid).rotas -= c; S(m.tid).total -= c; }
      else if (m.prestamoId && P[m.prestamoId]) { var p3 = P[m.prestamoId]; c = Math.min(k, p3.pend[m.tid] || 0); p3.pend[m.tid] = (p3.pend[m.tid] || 0) - c; S(m.tid).prestado -= c; S(m.tid).total -= c; }
      else S(m.tid).total -= k;
    }
  });
  Object.keys(T).forEach(function (tid) { var x = T[tid]; x.disp = x.total - x.prestado - x.rotas; });
  var abiertos = [];
  Object.keys(P).forEach(function (id) {
    var p = P[id], tot = 0;
    Object.keys(p.pend).forEach(function (tid) { tot += Math.max(0, p.pend[tid]); });
    p.abierto = tot > 0 && !p.m.cerrado;
    p.vencido = p.abierto && !!p.m.hasta && p.m.hasta < dia;
    if (p.abierto) abiertos.push(p);
  });
  abiertos.sort(function (a, b) { return cmpMov(a.m, b.m); });
  Object.keys(R).forEach(function (id) { if (R[id].abiertas <= 0) delete R[id]; });
  return { T: T, P: P, R: R, abiertos: abiertos };
}

/* ===== Utilidades ===== */

function mapaTipos(tipos) { var o = {}; tipos.forEach(function (t) { o[t.id] = t; }); return o; }
function itemsDe(m) { return m.items ? m.items : (m.tid ? [{ tid: m.tid, cant: m.cant }] : []); }
function pendItems(p) { return Object.keys(p.pend).filter(function (k) { return p.pend[k] > 0; }).map(function (k) { return { tid: k, cant: p.pend[k] }; }); }
function cmpMov(a, b) { return a.t < b.t ? -1 : a.t > b.t ? 1 : (a.id < b.id ? -1 : a.id > b.id ? 1 : 0); }
function nomT(TP, tid, n) { var t = TP[tid]; if (!t) return 'muebles'; return (n === 1 ? t.nombre : t.plural).toLowerCase(); }
function itemsTxt(items, TP) {
  var a = (items || []).filter(function (i) { return +i.cant > 0; }).map(function (i) { return i.cant + ' ' + nomT(TP, i.tid, +i.cant); });
  if (a.length < 2) return a.join('');
  return a.slice(0, -1).join(', ') + ' y ' + a[a.length - 1];
}
function norm(s) { return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim(); }
function texto(v, n) { return v == null ? '' : String(v).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, n || 200); }
function entero(v, conSigno) { var n = Math.round(Number(v)); if (!isFinite(n)) return 0; return conSigno ? n : Math.max(0, n); }
function hoy() { return Utilities.formatDate(new Date(), ZONA, 'yyyy-MM-dd'); }
function fechaCorta(s) { var p = String(s).split('-'); return p.length === 3 ? p[2] + '/' + p[1] + '/' + p[0] : String(s); }

function limpiar(x) {
  if (!x || typeof x !== 'object') return null;
  var m = { id: texto(x.id, 60) || 'm' + Utilities.getUuid().replace(/-/g, '').slice(0, 14), tipo: texto(x.tipo, 20) };
  if (x.t && !isNaN(Date.parse(x.t))) m.t = new Date(x.t).toISOString();
  ['quien', 'rol', 'lugar', 'tel', 'falla', 'origen', 'motivo', 'prestamoId', 'roturaId', 'tid'].forEach(function (k) {
    if (x[k] != null && x[k] !== '') m[k] = texto(x[k], 120);
  });
  ['obs', 'detalle'].forEach(function (k) { if (x[k] != null && x[k] !== '') m[k] = texto(x[k], 300); });
  if (m.tipo === 'prestamo') m.hasta = /^\d{4}-\d{2}-\d{2}$/.test(String(x.hasta || '')) ? String(x.hasta) : '';
  if (Array.isArray(x.items)) {
    m.items = x.items.map(function (i) {
      var o = { tid: texto(i && i.tid, 60), cant: entero(i && i.cant) };
      if (i && i.rotas != null) o.rotas = entero(i.rotas);
      return o;
    }).filter(function (i) { return i.tid; });
  }
  if (x.cant != null) m.cant = entero(x.cant, m.tipo === 'ajuste');
  if (x.descarte) m.descarte = true;
  if (x.cerrado) m.cerrado = true;
  if (x.contado != null) m.contado = entero(x.contado);
  if (x.esperado != null) m.esperado = entero(x.esperado, true);
  return m;
}
