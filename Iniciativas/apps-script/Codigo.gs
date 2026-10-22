// API de la votacion "Que priorizamos en Riesgos para 2027"
// Se pega en Extensiones > Apps Script de la Google Sheet que guarda los votos.
// Antes de implementar, cambia la clave del facilitador (no la subas a ningun repo).

const CLAVE_FACILITADOR = "cambiar-esta-clave";

const IMPULSORES = ["NIIF 9", "Finanzas abiertas", "Negocio", "Riesgos", "Cobranzas", "Eficiencia"];
const MAX_AGREGADAS = 200;
const HOJAS = {
  Iniciativas: ["id", "frente", "texto", "impulsores", "autor", "orden", "creado", "borrada"],
  Votos: ["uid", "item", "prioridad", "actualizado"],
  Acuerdos: ["item", "prioridad", "actualizado"],
  Config: ["clave", "valor"]
};

function doGet() {
  return salida({ ok: true, mensaje: "API de priorizacion activa" });
}

function doPost(e) {
  let b;
  try { b = JSON.parse((e && e.postData && e.postData.contents) || "{}"); }
  catch (err) { return salida({ ok: false, error: "json" }); }

  const uid = String(b.uid || "");
  if (!/^[a-z0-9-]{8,40}$/i.test(uid)) return salida({ ok: false, error: "uid" });
  const fac = CLAVE_FACILITADOR !== "cambiar-esta-clave" && String(b.clave || "") === CLAVE_FACILITADOR;

  if (b.accion !== "estado") {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) return salida({ ok: false, error: "ocupado" });
    try {
      switch (b.accion) {
        case "votar": votar(uid, b.item, b.p); break;
        case "agregar": agregar(uid, b); break;
        case "editar": editar(uid, fac, b); break;
        case "borrar": borrar(uid, fac, b.id); break;
        case "acordar": acordar(fac, b.item, b.p); break;
        case "ocultar": ocultar(fac, b.valor); break;
        default: throw new Error("accion");
      }
      SpreadsheetApp.flush();
    } catch (err) {
      return salida({ ok: false, error: String(err.message || err) });
    } finally {
      lock.releaseLock();
    }
  }
  return salida({ ok: true, estado: estado(uid, fac) });
}

// ---------- acciones ----------
function votar(uid, item, p) {
  item = validarId(item);
  if (["A", "M", "B", ""].indexOf(p) < 0) throw new Error("prioridad");
  upsert("Votos", r => r[0] === uid && r[1] === item, p ? [uid, item, p, new Date()] : null);
}

function agregar(uid, b) {
  const d = validarIniciativa(b);
  const filas = leer("Iniciativas");
  if (filas.length >= MAX_AGREGADAS) throw new Error("limite");
  const id = "t" + Date.now().toString(36) + Math.floor(Math.random() * 1296).toString(36);
  hoja("Iniciativas").appendRow([id, d.frente, d.texto, d.impulsores, uid, 100 + filas.length, new Date(), false]);
}

function editar(uid, fac, b) {
  const d = validarIniciativa(b);
  const { fila, r } = buscarIniciativa(b.id);
  if (!fac && r[4] !== uid) throw new Error("permiso");
  hoja("Iniciativas").getRange(fila, 2, 1, 3).setValues([[d.frente, d.texto, d.impulsores]]);
}

function borrar(uid, fac, id) {
  const { fila, r } = buscarIniciativa(id);
  if (!fac && r[4] !== uid) throw new Error("permiso");
  hoja("Iniciativas").getRange(fila, 8).setValue(true);
}

function acordar(fac, item, p) {
  if (!fac) throw new Error("permiso");
  item = validarId(item);
  if (["A", "M", "B", ""].indexOf(p) < 0) throw new Error("prioridad");
  upsert("Acuerdos", r => r[0] === item, p ? [item, p, new Date()] : null);
}

function ocultar(fac, valor) {
  if (!fac) throw new Error("permiso");
  upsert("Config", r => r[0] === "ocultar", ["ocultar", !!valor]);
}

// ---------- lectura ----------
function estado(uid, fac) {
  const ocultos = leer("Config").some(r => r[0] === "ocultar" && (r[1] === true || r[1] === "TRUE"));
  const iniciativas = leer("Iniciativas")
    .filter(r => r[0] && r[7] !== true && r[7] !== "TRUE")
    .map(r => ({
      id: String(r[0]), frente: String(r[1]), texto: String(r[2]),
      impulsores: r[3] ? String(r[3]).split(" + ") : [],
      orden: Number(r[5]) || 0, mia: r[4] === uid
    }));

  const votos = {};
  leer("Votos").forEach(r => {
    if (!r[0] || !r[2]) return;
    (votos[r[0]] = votos[r[0]] || {})[String(r[1])] = String(r[2]);
  });
  const votantes = Object.keys(votos).length;

  const acuerdos = {};
  leer("Acuerdos").forEach(r => { if (r[0] && r[1]) acuerdos[String(r[0])] = String(r[1]); });

  return {
    iniciativas, acuerdos, votantes, ocultos, facilitador: fac,
    votos: ocultos && !fac ? { [uid]: votos[uid] || {} } : votos
  };
}

// ---------- helpers ----------
function hoja(nombre) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(nombre);
  if (!sh) {
    sh = ss.insertSheet(nombre);
    sh.getRange(1, 1, 1, HOJAS[nombre].length).setValues([HOJAS[nombre]]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

function leer(nombre) {
  const sh = hoja(nombre);
  const n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, HOJAS[nombre].length).getValues() : [];
}

// Reemplaza la primera fila que cumple `match` por `valores`; si `valores` es null la borra.
function upsert(nombre, match, valores) {
  const sh = hoja(nombre);
  const i = leer(nombre).findIndex(match);
  if (i >= 0 && valores) sh.getRange(i + 2, 1, 1, valores.length).setValues([valores]);
  else if (i >= 0) sh.deleteRow(i + 2);
  else if (valores) sh.appendRow(valores);
}

function buscarIniciativa(id) {
  id = validarId(id);
  const filas = leer("Iniciativas");
  const i = filas.findIndex(r => r[0] === id && r[7] !== true && r[7] !== "TRUE");
  if (i < 0) throw new Error("no_existe");
  return { fila: i + 2, r: filas[i] };
}

function validarId(id) {
  id = String(id || "");
  if (!/^[a-z0-9]{1,24}$/i.test(id)) throw new Error("id");
  return id;
}

function validarIniciativa(b) {
  const frente = texto(b.frente, 60), txt = texto(b.texto, 500);
  if (!frente || !txt) throw new Error("incompleta");
  const imp = (Array.isArray(b.impulsores) ? b.impulsores : []).filter(t => IMPULSORES.indexOf(t) >= 0);
  return { frente, texto: txt, impulsores: imp.join(" + ") };
}

// Recorta y evita que la planilla interprete el texto como formula.
function texto(s, max) {
  s = String(s == null ? "" : s).trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function salida(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
