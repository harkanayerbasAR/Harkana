/**
 * PEDIDOS HARKANA — motor que guarda los pedidos en esta planilla.
 * 1) Completá las dos claves y los mails de aviso acá abajo.
 * 2) Guardá, y después: Implementar > Administrar implementaciones > lápiz > Nueva versión > Implementar.
 */
const CLAVE = "CAMBIAR-ESTA-CLAVE";               // clave del DUEÑO: ve y puede todo
const CLAVE_SOCIO = "CAMBIAR-CLAVE-SOCIO";         // clave del SOCIO: entrega pedidos y carga retiros, nada más
const CLAVE_CIERRE = "CAMBIAR-CLAVE-CIERRE";       // clave aparte, solo para CERRAR DÍA en la feria (la sabe solo el dueño)
const NOMBRE_SOCIO = "Felipe";
const AVISAR_A = "harkanayerbas@gmail.com";        // mails de aviso, separados por coma
const TARIFA = { q250: 300, q500: 500, q1: 800, q3: 2000 };   // lo que cobra el socio por paquete elaborado
const SALDO_INICIAL = 373100;                      // a favor del socio al arrancar el registro
const PRECIOS_FERIA = { q250: 3900, q500: 5500, q1: 8000, q3: 23000 };   // precios de venta al público en ferias
const WEB = "https://harkana.com.ar/pedidos";      // dirección de la página de pedidos
const ZONA = "America/Argentina/Cordoba";

const PRES = [["q250", 0.25, "1/4 kg"], ["q500", 0.5, "1/2 kg"], ["q1", 1, "1 kg"], ["q3", 3, "3 kg"]];
const COLS = ["id", "fecha", "codigo", "comercio", "localidad", "telefono", "q250", "q500", "q1", "q3", "kg", "nota", "estado", "entregado_el", "cliente_conocido", "pago_produccion"];
const COMERCIOS = ["Salvador Supermercado", "Regionales Don Valente", "Kwik E Mart", "Shop Puma", "IKEI", "German Dos Alejandro", "German Dos Reducción", "Mini Super Mari", "Supermercado Renacer", "Verdulería Los Tres Hermanos", "El Inglés"];

function configurar() {
  const ss = SpreadsheetApp.getActive();
  let p = ss.getSheetByName("Pedidos") || ss.insertSheet("Pedidos");
  if (p.getLastRow() === 0) { p.appendRow(COLS); p.setFrozenRows(1); p.getRange(1, 1, 1, COLS.length).setFontWeight("bold"); }
  let s = ss.getSheetByName("Stock") || ss.insertSheet("Stock");
  if (s.getLastRow() === 0) { s.appendRow(["fecha", "kg", "nota"]); s.setFrozenRows(1); s.getRange(1, 1, 1, 3).setFontWeight("bold"); }
  let c = ss.getSheetByName("Comercios") || ss.insertSheet("Comercios");
  if (c.getLastRow() === 0) {
    c.appendRow(["codigo", "nombre", "localidad", "link_para_enviarle"]);
    c.setFrozenRows(1); c.getRange(1, 1, 1, 4).setFontWeight("bold");
    COMERCIOS.forEach(function (n) { const cod = nuevoCodigo_(); c.appendRow([cod, n, "", WEB + "?c=" + cod]); });
    c.autoResizeColumns(1, 4);
  }
}

/** Para sumar un comercio: escribí el nombre en la hoja Comercios y ejecutá "completarCodigos". */
function completarCodigos() {
  const c = SpreadsheetApp.getActive().getSheetByName("Comercios");
  const v = c.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (v[i][1] && !v[i][0]) { const cod = nuevoCodigo_(); c.getRange(i + 1, 1).setValue(cod); c.getRange(i + 1, 4).setValue(WEB + "?c=" + cod); }
  }
}

function nuevoCodigo_() {
  const abc = "ABCDEFGHJKMNPQRSTUVWXYZ23456789"; let s = "";
  for (let i = 0; i < 6; i++) s += abc.charAt(Math.floor(Math.random() * abc.length));
  return s;
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function txt_(v, max) { let s = String(v == null ? "" : v).trim().slice(0, max); if (/^[=+\-@]/.test(s)) s = "'" + s; return s; }
function num_(v) { const n = parseInt(v, 10); return isNaN(n) ? 0 : Math.max(0, Math.min(999, n)); }

function buscarComercio_(codigo) {
  codigo = String(codigo || "").trim().toUpperCase();
  if (!codigo) return null;
  const c = SpreadsheetApp.getActive().getSheetByName("Comercios");
  if (!c) return null;
  const v = c.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim().toUpperCase() === codigo) return { codigo: codigo, nombre: String(v[i][1]), localidad: String(v[i][2] || "") };
  }
  return null;
}

function doGet(e) {
  const com = buscarComercio_(e && e.parameter ? e.parameter.cod : "");
  if (!com) return json_({ ok: false });
  // Último pedido de ese comercio, para el botón "Repetir mi último pedido".
  let ultimo = null;
  const v = SpreadsheetApp.getActive().getSheetByName("Pedidos").getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) {
    if (String(v[i][2]).trim().toUpperCase() === com.codigo) {
      ultimo = { q250: Number(v[i][6]) || 0, q500: Number(v[i][7]) || 0, q1: Number(v[i][8]) || 0, q3: Number(v[i][9]) || 0 };
      break;
    }
  }
  return json_({ ok: true, nombre: com.nombre, ultimo: ultimo });
}

/** Devuelve "admin", "socio" o null. Una clave que quedó sin cambiar nunca sirve para entrar. */
function rol_(clave) {
  clave = String(clave || "");
  if (!clave || clave.indexOf("CAMBIAR") === 0) return null;
  if (clave === CLAVE) return "admin";
  if (clave === CLAVE_SOCIO) return "socio";
  return null;
}

/** La clave de cierre del día de feria. Si quedó sin cambiar, nadie puede cerrar. */
function cierreOk_(c) { c = String(c || ""); return !!c && CLAVE_CIERRE.indexOf("CAMBIAR") !== 0 && c === CLAVE_CIERRE; }

function pago_(q250, q500, q1, q3) {
  return (Number(q250) || 0) * TARIFA.q250 + (Number(q500) || 0) * TARIFA.q500 + (Number(q1) || 0) * TARIFA.q1 + (Number(q3) || 0) * TARIFA.q3;
}

/** Crea la hoja Produccion (con el saldo inicial) y la columna pago_produccion la primera vez. */
function hojaProduccion_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName("Produccion");
  if (!sh) {
    sh = ss.insertSheet("Produccion");
    sh.appendRow(["fecha", "tipo", "monto", "nota", "cargado_por"]);
    sh.appendRow([new Date(), "saldo inicial", SALDO_INICIAL, "A favor de " + NOMBRE_SOCIO + " al arrancar", "sistema"]);
    const p = ss.getSheetByName("Pedidos");
    if (p) p.getRange(1, 16).setValue("pago_produccion");
  }
  return sh;
}

/** tipo "retiro": el socio saca plata (monto positivo). tipo "ajuste": corrección del dueño (positivo o negativo). */
function movimiento_(tipo, monto, nota, rol) {
  monto = Math.round(Number(String(monto).replace(/\./g, "").replace(",", ".")));
  if (!monto || isNaN(monto) || Math.abs(monto) > 100000000) return { ok: false, error: "monto" };
  if (tipo === "retiro" && monto < 0) return { ok: false, error: "monto" };
  const quien = rol === "admin" ? "dueño" : NOMBRE_SOCIO;
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { hojaProduccion_().appendRow([new Date(), tipo, monto, txt_(nota, 120), quien]); } finally { lock.releaseLock(); }
  if (tipo === "retiro") {
    try {
      const r = lista_();
      MailApp.sendEmail(AVISAR_A, "Retiro de " + NOMBRE_SOCIO + ": $" + monto,
        "Retiro cargado por " + quien + ": $" + monto + (nota ? "\nNota: " + txt_(nota, 120) : "") + "\nSaldo a favor que queda: $" + r.produccion.saldo);
      return r;
    } catch (err) { /* el retiro ya quedó guardado aunque falle el mail */ }
  }
  return lista_();
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: "datos" }); }
  try {
    if (d.accion === "pedido") return json_(guardarPedido_(d));
    if (d.accion === "contacto") return json_(guardarContacto_(d));
    if (d.accion === "compra") return json_(guardarCompra_(d));
    const rol = rol_(d.clave);
    if (!rol) return json_({ ok: false, error: "clave" });
    let r;
    if (d.accion === "lista") r = lista_();
    else if (d.accion === "estado") r = cambiarEstado_(d.id, d.estado);
    else if (d.accion === "retiro") r = movimiento_("retiro", d.monto, d.nota, rol);
    else if (d.accion === "feriaVentas") r = feriaVentas_(d);
    else if ((d.accion === "feriaResumen" || d.accion === "feriaCerrar") && !cierreOk_(d.claveCierre)) return json_({ ok: false, error: "clave_cierre" });
    else if (d.accion === "feriaResumen") r = feriaResumen_(d.origen, d.consultas);
    else if (d.accion === "feriaCerrar") r = feriaCerrar_(d);
    else if (rol !== "admin") return json_({ ok: false, error: "permiso" });
    else if (d.accion === "stock") r = cargarStock_(d.kg, d.nota);
    else if (d.accion === "ajuste") r = movimiento_("ajuste", d.monto, d.nota, rol);
    else if (d.accion === "borrar") r = borrarPedido_(d.id);
    else if (d.accion === "editar") r = editarPedido_(d.id, d);
    if (r) { if (r.ok) r.rol = rol; return json_(r); }
    return json_({ ok: false, error: "accion" });
  } catch (err) { return json_({ ok: false, error: "servidor" }); }
}

function guardarPedido_(d) {
  const com = buscarComercio_(d.codigo);
  const nombre = com ? com.nombre : txt_(d.comercio, 60);
  const localidad = com && com.localidad ? com.localidad : txt_(d.localidad, 40);
  const q = {}; let kg = 0;
  PRES.forEach(function (p) { q[p[0]] = num_(d[p[0]]); kg += q[p[0]] * p[1]; });
  if (!nombre) return { ok: false, error: "nombre" };
  if (kg <= 0) return { ok: false, error: "cantidad" };
  const id = Utilities.getUuid().slice(0, 8);
  const fila = [id, new Date(), com ? com.codigo : "", nombre, localidad, txt_(d.telefono, 20), q.q250, q.q500, q.q1, q.q3, kg, txt_(d.nota, 200), "pendiente", "", com ? "sí" : "NO — confirmar"];
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { SpreadsheetApp.getActive().getSheetByName("Pedidos").appendRow(fila); } finally { lock.releaseLock(); }
  try {
    const det = PRES.filter(function (p) { return q[p[0]] > 0; }).map(function (p) { return q[p[0]] + " × " + p[2]; }).join(" · ");
    MailApp.sendEmail(AVISAR_A, "Pedido Harkana: " + nombre + " — " + kg + " kg",
      nombre + (localidad ? " (" + localidad + ")" : "") + "\n" + det + "\nTotal: " + kg + " kg" +
      (fila[11] ? "\nNota: " + fila[11] : "") + (fila[5] ? "\nTeléfono: " + fila[5] : "") +
      (com ? "" : "\n\nATENCIÓN: pidió sin link propio. Confirmá que sea un cliente real."));
  } catch (err) { /* el pedido ya quedó guardado aunque falle el mail */ }
  return { ok: true, id: id, comercio: nombre, kg: kg };
}

function fmt_(v) { return v instanceof Date ? Utilities.formatDate(v, ZONA, "dd/MM HH:mm") : String(v || ""); }

function lista_() {
  const v = SpreadsheetApp.getActive().getSheetByName("Pedidos").getDataRange().getValues();
  const pend = [], entr = []; const tot = { kg: 0, q250: 0, q500: 0, q1: 0, q3: 0 }; let kgEntregado = 0;
  const mes = Utilities.formatDate(new Date(), ZONA, "yyyy-MM"); let ganado = 0, ganadoMes = 0;
  const porCodigo = {}, conPendiente = {};
  for (let i = 1; i < v.length; i++) {
    const r = v[i]; if (!r[0]) continue;
    const cod = String(r[2] || "").trim().toUpperCase();
    if (cod) { (porCodigo[cod] = porCodigo[cod] || []).push(r); if (String(r[12]) !== "entregado") conPendiente[cod] = true; }
    const o = { id: String(r[0]), fecha: fmt_(r[1]), comercio: String(r[3]), localidad: String(r[4]), telefono: String(r[5]),
      q250: Number(r[6]) || 0, q500: Number(r[7]) || 0, q1: Number(r[8]) || 0, q3: Number(r[9]) || 0, kg: Number(r[10]) || 0,
      nota: String(r[11] || ""), conocido: String(r[14]) === "sí", entregado: fmt_(r[13]) };
    if (String(r[12]) === "entregado") {
      entr.push(o); kgEntregado += o.kg;
      o.pago = Number(r[15]) || 0; ganado += o.pago;
      if (r[13] instanceof Date && Utilities.formatDate(r[13], ZONA, "yyyy-MM") === mes) ganadoMes += o.pago;
    }
    else { o.pago = pago_(o.q250, o.q500, o.q1, o.q3); pend.push(o); tot.kg += o.kg; tot.q250 += o.q250; tot.q500 += o.q500; tot.q1 += o.q1; tot.q3 += o.q3; }
  }
  // Stock: lo comprado (hoja Stock) menos todo lo pedido. "deposito" todavía incluye lo pendiente de entregar.
  let comprado = 0; const mov = [];
  const sh = SpreadsheetApp.getActive().getSheetByName("Stock");
  if (sh) {
    const s = sh.getDataRange().getValues();
    for (let i = 1; i < s.length; i++) {
      const k = Number(s[i][1]) || 0; if (!k) continue;
      comprado += k; mov.push({ fecha: fmt_(s[i][0]), kg: k, nota: String(s[i][2] || "") });
    }
  }
  const stock = { cargado: mov.length > 0, deposito: comprado - kgEntregado, restante: comprado - kgEntregado - tot.kg, movimientos: mov.slice(-5).reverse() };
  // Producción del socio: saldo inicial y ajustes + lo ganado por pedidos entregados - retiros.
  let base = 0, retirado = 0, retiradoMes = 0; const movs = [];
  const pv = hojaProduccion_().getDataRange().getValues();
  for (let i = 1; i < pv.length; i++) {
    const m = Number(pv[i][2]) || 0; const tipo = String(pv[i][1]); if (!m) continue;
    if (tipo === "retiro") { retirado += m; if (pv[i][0] instanceof Date && Utilities.formatDate(pv[i][0], ZONA, "yyyy-MM") === mes) retiradoMes += m; }
    else base += m;
    movs.push({ fecha: fmt_(pv[i][0]), tipo: tipo, monto: m, nota: String(pv[i][3] || ""), quien: String(pv[i][4] || "") });
  }
  const produccion = { socio: NOMBRE_SOCIO, saldo: base + ganado - retirado, ganadoMes: ganadoMes, retiradoMes: retiradoMes, tarifa: TARIFA, movimientos: movs.slice(-8).reverse() };
  return { ok: true, pendientes: pend, entregados: entr.slice(-10).reverse(), totales: tot, stock: stock, produccion: produccion, recordar: recordar_(porCodigo, conPendiente) };
}

/** Suma una compra de yerba (kg positivo) o un ajuste (kg negativo: ventas sueltas, mermas, corrección). */
function cargarStock_(kg, nota) {
  kg = Math.round(Number(String(kg).replace(",", ".")) * 100) / 100;
  if (!kg || isNaN(kg) || Math.abs(kg) > 100000) return { ok: false, error: "kg" };
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName("Stock") || ss.insertSheet("Stock");
    if (sh.getLastRow() === 0) sh.appendRow(["fecha", "kg", "nota"]);
    sh.appendRow([new Date(), kg, txt_(nota, 120)]);
  } finally { lock.releaseLock(); }
  return lista_();
}

function cambiarEstado_(id, estado) {
  estado = estado === "pendiente" ? "pendiente" : "entregado";
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName("Pedidos");
    const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
    for (let i = 1; i < ids.length; i++) {
      if (String(ids[i][0]) === String(id)) {
        sh.getRange(i + 1, 13, 1, 2).setValues([[estado, estado === "entregado" ? new Date() : ""]]);
        // Al entregar se acredita la producción al socio con la tarifa de hoy; al deshacer se quita.
        hojaProduccion_();
        const q = sh.getRange(i + 1, 7, 1, 4).getValues()[0];
        sh.getRange(i + 1, 16).setValue(estado === "entregado" ? pago_(q[0], q[1], q[2], q[3]) : "");
        return lista_();
      }
    }
  } finally { lock.releaseLock(); }
  return { ok: false, error: "no_encontrado" };
}

/**
 * Comercios a los que conviene escribirles: ya pasó su tiempo habitual entre pedidos y no tienen nada pendiente.
 * Con un solo pedido cargado se toma 21 días. El teléfono sale de la columna E de la hoja Comercios (opcional).
 */
function recordar_(porCodigo, conPendiente) {
  const sh = SpreadsheetApp.getActive().getSheetByName("Comercios");
  if (!sh) return [];
  const c = sh.getDataRange().getValues(); const hoy = new Date().getTime(); const out = [];
  for (let i = 1; i < c.length; i++) {
    const cod = String(c[i][0] || "").trim().toUpperCase(); if (!cod || conPendiente[cod]) continue;
    const filas = porCodigo[cod]; if (!filas || !filas.length) continue;
    const t = filas.map(function (r) { return r[1] instanceof Date ? r[1].getTime() : 0; }).filter(function (x) { return x > 0; }).sort(function (a, b) { return a - b; });
    if (!t.length) continue;
    const dias = Math.floor((hoy - t[t.length - 1]) / 86400000);
    const cada = t.length > 1 ? Math.round((t[t.length - 1] - t[0]) / 86400000 / (t.length - 1)) : 0;
    const limite = cada ? Math.max(7, cada) : 21;
    if (dias < limite) continue;
    out.push({ nombre: String(c[i][1]), dias: dias, cada: cada, link: String(c[i][3] || ""), telefono: String(c[i][4] || "").replace(/\D/g, "") });
  }
  out.sort(function (a, b) { return b.dias - a.dias; });
  return out;
}

function filaPedido_(sh, id) {
  const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = 1; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 1;
  return 0;
}

/** Solo el dueño: borra un pedido cargado por error. */
function borrarPedido_(id) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName("Pedidos");
    const f = filaPedido_(sh, id); if (!f) return { ok: false, error: "no_encontrado" };
    sh.deleteRow(f);
  } finally { lock.releaseLock(); }
  return lista_();
}

/** Solo el dueño: corrige las cantidades de un pedido. Si ya estaba entregado, recalcula la producción. */
function editarPedido_(id, d) {
  const q = PRES.map(function (p) { return num_(d[p[0]]); });
  const kg = q[0] * 0.25 + q[1] * 0.5 + q[2] * 1 + q[3] * 3;
  if (kg <= 0) return { ok: false, error: "cantidad" };
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName("Pedidos");
    const f = filaPedido_(sh, id); if (!f) return { ok: false, error: "no_encontrado" };
    sh.getRange(f, 7, 1, 5).setValues([[q[0], q[1], q[2], q[3], kg]]);
    if (String(sh.getRange(f, 13).getValue()) === "entregado") { hojaProduccion_(); sh.getRange(f, 16).setValue(pago_(q[0], q[1], q[2], q[3])); }
  } finally { lock.releaseLock(); }
  return lista_();
}

/**
 * Público: alguien deja sus datos desde el QR (tipo "muestra") o quiere revender (tipo "revendedor").
 * Una fila por celular y tipo. No manda mail por cada muestra; sí avisa cuando alguien quiere revender.
 */
function guardarContacto_(d) {
  const nombre = txt_(d.nombre, 60), localidad = txt_(d.localidad, 40);
  const cel = String(d.celular || "").replace(/\D/g, "").slice(0, 20);
  const tipo = d.tipo === "revendedor" ? "revendedor" : "muestra";
  if (!nombre) return { ok: false, error: "nombre" };
  if (cel.length < 8) return { ok: false, error: "celular" };
  let numero = 0, repetido = false;
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActive();
    const hoja = tipo === "revendedor" ? "Reventa" : "Muestras";
    let sh = ss.getSheetByName(hoja);
    if (!sh) { sh = ss.insertSheet(hoja); sh.appendRow(["fecha", "nombre", "celular", "localidad", "tipo", "origen", "numero"]); sh.setFrozenRows(1); }
    const v = sh.getDataRange().getValues();
    for (let i = 1; i < v.length; i++) {
      if (String(v[i][4]) !== tipo) continue;
      numero++;
      if (celKey_(v[i][2]) === celKey_(cel)) { repetido = true; numero = Number(v[i][6]) || numero; break; }
    }
    if (!repetido) { numero++; sh.appendRow([new Date(), nombre, "'" + cel, localidad, tipo, txt_(d.origen, 40), numero]); }
  } finally { lock.releaseLock(); }
  if (tipo === "revendedor" && !repetido) {
    try { MailApp.sendEmail(AVISAR_A, "Quiere revender Harkana: " + nombre, nombre + (localidad ? " (" + localidad + ")" : "") + "\nCelular: " + cel); } catch (err) { }
  }
  return { ok: true, tipo: tipo, numero: numero, repetido: repetido, nombre: nombre };
}

/* ======================= FERIAS: venta en el stand, QR de compra y cierre del día ======================= */

/**
 * Celular comparable: característica + número, 10 dígitos (ej. 3584388479).
 * Entiende +54 9, 54, 0, el 15 después de la característica y espacios o guiones:
 * +5490358154388479, 03584388479, 358154388479 y 3584388479 dan lo mismo.
 */
function celKey_(v) {
  let s = String(v || "").replace(/\D/g, "");
  if (s.indexOf("00") === 0) s = s.slice(2);
  if (s.length > 10 && s.indexOf("54") === 0) s = s.slice(2);
  if (s.length > 10 && s.charAt(0) === "9") s = s.slice(1);
  if (s.charAt(0) === "0") s = s.slice(1);
  if (s.length === 12) {
    const pos = [3, 4, 2];   // característica de 3 (Río Cuarto, Córdoba), 4 o 2 dígitos (Buenos Aires)
    for (let i = 0; i < pos.length; i++) if (s.substr(pos[i], 2) === "15") { s = s.slice(0, pos[i]) + s.slice(pos[i] + 2); break; }
  }
  return s.length > 10 ? s.slice(-10) : s;
}

function hoja_(nombre, cols) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(nombre);
  if (!sh) { sh = ss.insertSheet(nombre); sh.appendRow(cols); sh.setFrozenRows(1); sh.getRange(1, 1, 1, cols.length).setFontWeight("bold"); }
  return sh;
}
const COLS_VENTAS = ["id", "fecha", "origen", "numero", "q250", "q500", "q1", "q3", "kg", "total", "pago", "estado", "subida_el"];
const COLS_COMPRADORES = ["fecha", "venta", "nombre", "celular", "localidad", "origen", "datos"];

/** Busca a una persona por celular en Compradores, Muestras y Reventa. Devuelve {nombre, localidad} o null. */
function buscarPersona_(cel) {
  const k = celKey_(cel); if (k.length < 8) return null;
  const ss = SpreadsheetApp.getActive(); let hallado = null;
  [["Compradores", 2, 3, 4], ["Reventa", 1, 2, 3], ["Muestras", 1, 2, 3]].forEach(function (h) {
    if (hallado) return;
    const sh = ss.getSheetByName(h[0]); if (!sh) return;
    const v = sh.getDataRange().getValues();
    for (let i = v.length - 1; i >= 1; i--) {
      if (celKey_(v[i][h[2]]) === k && String(v[i][h[1]]).trim()) { hallado = { nombre: String(v[i][h[1]]).trim(), localidad: String(v[i][h[3]] || "") }; return; }
    }
  });
  return hallado;
}

/**
 * Público: el cliente escaneó el QR de su compra.
 * Solo celular → si ya lo conocemos se une la compra; si no, responde nuevo:true para pedir nombre y localidad.
 * Una compra queda a nombre de una sola persona (la primera que la registra).
 */
function guardarCompra_(d) {
  const venta = String(d.venta || "").replace(/[^A-Za-z0-9]/g, "").slice(0, 12);
  const cel = String(d.celular || "").replace(/\D/g, "").slice(0, 20);
  if (venta.length < 4) return { ok: false, error: "venta" };
  if (cel.length < 8) return { ok: false, error: "celular" };
  let nombre = txt_(d.nombre, 60), localidad = txt_(d.localidad, 40), datos = "nuevo";
  if (!nombre) {
    const p = buscarPersona_(cel);
    if (!p) return { ok: false, nuevo: true };
    nombre = p.nombre; localidad = p.localidad; datos = "conocido";
  }
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = hoja_("Compradores", COLS_COMPRADORES);
    const v = sh.getDataRange().getValues();
    for (let i = 1; i < v.length; i++) {
      if (String(v[i][1]) === venta) return { ok: true, nombre: nombre.split(" ")[0], yaEstaba: true, mismo: celKey_(v[i][3]) === celKey_(cel) };
    }
    sh.appendRow([new Date(), venta, nombre, "'" + cel, localidad, txt_(d.origen, 40), datos]);
  } finally { lock.releaseLock(); }
  return { ok: true, nombre: nombre.split(" ")[0], localidad: localidad };
}

/** El celular del stand sube sus ventas (también las anuladas). Se guardan por id: subir dos veces no duplica. */
function feriaVentas_(d) {
  const lista = Array.isArray(d.ventas) ? d.ventas.slice(0, 500) : [];
  const origen = txt_(d.origen, 40) || "feria";
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  const subidas = [];
  try {
    const sh = hoja_("Ventas", COLS_VENTAS);
    const v = sh.getDataRange().getValues(); const fila = {};
    for (let i = 1; i < v.length; i++) fila[String(v[i][0])] = i + 1;
    lista.forEach(function (x) {
      const id = String(x.id || "").replace(/[^A-Za-z0-9]/g, "").slice(0, 12); if (id.length < 4) return;
      const q = PRES.map(function (p) { return num_(x[p[0]]); });
      const kg = q[0] * 0.25 + q[1] * 0.5 + q[2] + q[3] * 3;
      const total = q[0] * PRECIOS_FERIA.q250 + q[1] * PRECIOS_FERIA.q500 + q[2] * PRECIOS_FERIA.q1 + q[3] * PRECIOS_FERIA.q3;
      const pago = ["efectivo", "transferencia", "mercadopago"].indexOf(x.pago) >= 0 ? x.pago : "efectivo";
      const t = new Date(Number(x.t) || Date.now());
      const r = [id, t, origen, num_(x.n), q[0], q[1], q[2], q[3], kg, total, pago, x.anulada ? "anulada" : "ok", new Date()];
      if (fila[id]) sh.getRange(fila[id], 1, 1, r.length).setValues([r]);
      else { sh.appendRow(r); fila[id] = sh.getLastRow(); }
      subidas.push(id);
    });
    if (d.consultas != null) guardarConsultas_(origen, d.consultas);
  } finally { lock.releaseLock(); }
  return { ok: true, subidas: subidas };
}

function guardarConsultas_(origen, n) {
  const sh = hoja_("Consultas", ["origen", "cantidad", "actualizado"]);
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === origen) { sh.getRange(i + 1, 2, 1, 2).setValues([[num_(n), new Date()]]); return; }
  sh.appendRow([origen, num_(n), new Date()]);
}

/** Resumen del día de feria: ventas, plata por forma de pago, muestras, compradores, revendedores, consultas. */
function feriaResumen_(origen, consultas) {
  origen = txt_(origen, 40) || "feria";
  const ss = SpreadsheetApp.getActive();
  const r = { ok: true, origen: origen, ventas: 0, anuladas: 0, q250: 0, q500: 0, q1: 0, q3: 0, kg: 0, total: 0,
    pagos: { efectivo: 0, transferencia: 0, mercadopago: 0 }, muestras: 0, revendedores: 0,
    conDatos: 0, sinDatos: 0, volvieron: 0, consultas: 0, cerrado: false };
  const ids = {};
  const vs = ss.getSheetByName("Ventas");
  if (vs) vs.getDataRange().getValues().slice(1).forEach(function (x) {
    if (String(x[2]) !== origen) return;
    if (String(x[11]) === "anulada") { r.anuladas++; return; }
    ids[String(x[0])] = x[1] instanceof Date ? x[1].getTime() : 0;
    r.ventas++; r.q250 += Number(x[4]) || 0; r.q500 += Number(x[5]) || 0; r.q1 += Number(x[6]) || 0; r.q3 += Number(x[7]) || 0;
    r.kg += Number(x[8]) || 0; r.total += Number(x[9]) || 0;
    const pg = String(x[10]); r.pagos[pg] = (r.pagos[pg] || 0) + (Number(x[9]) || 0);
  });
  const muestraCel = {};
  const ms = ss.getSheetByName("Muestras");
  if (ms) ms.getDataRange().getValues().slice(1).forEach(function (x) {
    if (String(x[5]) === origen) r.muestras++;
    muestraCel[celKey_(x[2])] = true;
  });
  const rs = ss.getSheetByName("Reventa");
  if (rs) rs.getDataRange().getValues().slice(1).forEach(function (x) { if (String(x[5]) === origen) r.revendedores++; });
  const compr = {}, volvio = {};
  const cs = ss.getSheetByName("Compradores");
  if (cs) cs.getDataRange().getValues().slice(1).forEach(function (x) {
    const id = String(x[1]); if (!(id in ids) || compr[id]) return;
    compr[id] = true; r.conDatos++;
    const k = celKey_(x[3]); if (muestraCel[k]) volvio[k] = true;
  });
  r.sinDatos = r.ventas - r.conDatos;
  r.volvieron = Object.keys(volvio).length;
  const cq = ss.getSheetByName("Consultas");
  if (cq) cq.getDataRange().getValues().slice(1).forEach(function (x) { if (String(x[0]) === origen) r.consultas = Number(x[1]) || 0; });
  if (consultas != null) r.consultas = Math.max(r.consultas, num_(consultas));
  const p = ss.getSheetByName("Pedidos");
  if (p) r.cerrado = filaPedido_(p, "F-" + origen) > 0;
  return r;
}

/**
 * Cierra el día: crea (o actualiza, si se cierra de nuevo) un pedido "entregado" con todo lo vendido.
 * Así descuenta el stock y suma la producción del socio igual que cualquier pedido entregado.
 * También arma la hoja "Cruce" (una fila por celular) y manda el resumen por mail.
 */
function feriaCerrar_(d) {
  if (d.ventas) feriaVentas_(d);
  const origen = txt_(d.origen, 40) || "feria";
  const nombre = txt_(d.nombre, 60) || "Feria";
  const r = feriaResumen_(origen, d.consultas);
  if (r.kg <= 0) return { ok: false, error: "sin_ventas" };
  const id = "F-" + origen;
  const money = function (n) { return "$" + Math.round(n).toLocaleString("es-AR"); };
  const nota = r.ventas + " ventas · " + money(r.total) + " (efectivo " + money(r.pagos.efectivo) + ", transferencia " + money(r.pagos.transferencia) +
    ", Mercado Pago " + money(r.pagos.mercadopago) + ") · " + r.consultas + " consultas · " + r.muestras + " muestras";
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName("Pedidos");
    hojaProduccion_();
    const fila = [id, new Date(), "", nombre, txt_(d.localidad, 40), "", r.q250, r.q500, r.q1, r.q3, r.kg, nota, "entregado", new Date(), "sí", pago_(r.q250, r.q500, r.q1, r.q3)];
    const f = filaPedido_(sh, id);
    if (f) { const prev = sh.getRange(f, 2).getValue(); fila[1] = prev || fila[1]; sh.getRange(f, 1, 1, fila.length).setValues([fila]); }
    else sh.appendRow(fila);
  } finally { lock.releaseLock(); }
  try { armarCruce(); } catch (err) { }
  try {
    MailApp.sendEmail(AVISAR_A, nombre + ": " + r.kg + " kg vendidos — " + money(r.total),
      nombre + "\n\nPaquetes: " + r.q250 + " × 1/4 · " + r.q500 + " × 1/2 · " + r.q1 + " × 1 kg · " + r.q3 + " × 3 kg  (" + r.kg + " kg)" +
      "\nVentas: " + r.ventas + " — " + money(r.total) +
      "\n  Efectivo " + money(r.pagos.efectivo) + "\n  Transferencia " + money(r.pagos.transferencia) + "\n  Mercado Pago " + money(r.pagos.mercadopago) +
      "\n\nMuestras: " + r.muestras + "\nCompradores con datos: " + r.conDatos + " · sin datos: " + r.sinDatos +
      "\nProbaron la muestra y compraron: " + r.volvieron + "\nConsultas: " + r.consultas + "\nQuieren revender: " + r.revendedores);
  } catch (err) { }
  r.cerrado = true; r.pedido = id;
  return r;
}

/** Arma la hoja "Cruce": una fila por celular uniendo muestras, compras y revendedores. Se puede ejecutar a mano. */
function armarCruce() {
  const ss = SpreadsheetApp.getActive(); const gente = {};
  const persona = function (cel, nombre, loc) {
    const k = celKey_(cel); if (k.length < 8) return null;
    const g = gente[k] = gente[k] || { cel: k, nombre: "", loc: "", muestra: "", compras: 0, gastado: 0, revende: "", origenes: {} };
    if (nombre && !g.nombre) g.nombre = String(nombre); if (loc && !g.loc) g.loc = String(loc);
    return g;
  };
  const totales = {};
  const vs = ss.getSheetByName("Ventas");
  if (vs) vs.getDataRange().getValues().slice(1).forEach(function (x) { if (String(x[11]) !== "anulada") totales[String(x[0])] = Number(x[9]) || 0; });
  const ms = ss.getSheetByName("Muestras");
  if (ms) ms.getDataRange().getValues().slice(1).forEach(function (x) { const g = persona(x[2], x[1], x[3]); if (g) { g.muestra = fmt_(x[0]); g.origenes[String(x[5])] = 1; } });
  const rs = ss.getSheetByName("Reventa");
  if (rs) rs.getDataRange().getValues().slice(1).forEach(function (x) { const g = persona(x[2], x[1], x[3]); if (g) { g.revende = fmt_(x[0]); g.origenes[String(x[5])] = 1; } });
  const vistas = {};
  const cs = ss.getSheetByName("Compradores");
  if (cs) cs.getDataRange().getValues().slice(1).forEach(function (x) {
    const id = String(x[1]); if (vistas[id]) return; vistas[id] = 1;
    const g = persona(x[3], x[2], x[4]); if (!g) return;
    g.compras++; g.gastado += totales[id] || 0; g.origenes[String(x[5])] = 1;
  });
  const filas = Object.keys(gente).map(function (k) {
    const g = gente[k];
    return ["'" + g.cel, g.nombre, g.loc, g.muestra ? "sí" : "", g.compras || "", g.gastado || "", g.revende ? "sí" : "", (g.muestra && g.compras) ? "sí" : "", Object.keys(g.origenes).filter(String).join(", ")];
  });
  let sh = ss.getSheetByName("Cruce");
  if (sh) sh.clear(); else sh = ss.insertSheet("Cruce");
  const cab = ["celular", "nombre", "localidad", "muestra", "compras", "gastado", "revendedor", "probó y compró", "dónde"];
  sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight("bold"); sh.setFrozenRows(1);
  if (filas.length) sh.getRange(2, 1, filas.length, cab.length).setValues(filas);
}
