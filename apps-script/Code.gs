/**
 * Canje Mayorista AJE · Backend (Google Apps Script)
 *
 * Recibe los registros de la app de canjistas (asistencia, canjes, incidencias),
 * los guarda en este Google Sheet y las fotos en una carpeta de Drive.
 * El drop y el premio se recalculan aquí con la hoja Drops (no se confía en el celular).
 *
 * Instalación: crear un proyecto de Apps Script, pegar este archivo, correr setup() una vez
 * y publicar como aplicación web (ejecutar como: yo · acceso: cualquier persona).
 * Copiar la URL /exec en config.js → API_URL.
 */

const TZ = 'America/Lima';
const P = PropertiesService.getScriptProperties();
const MERCADOS = ['Productores', 'Caquetá', 'Unicachi Norte', 'Huamantanga'];
const USUARIO_MERCADO = { 'Productores': 'productores', 'Caquetá': 'caqueta', 'Unicachi Norte': 'unicachi', 'Huamantanga': 'huamantanga' };

const ESQUEMA = {
  Canjes: ['id', 'registrado', 'fecha', 'hora', 'canjista', 'mercado', 'cliente', 'formato', 'bebida', 'paquetes',
    'drop', 'premio', 'foto_boleta', 'foto_premio', 'nota', 'lat', 'lng'],
  Asistencia: ['id', 'registrado', 'fecha', 'hora', 'canjista', 'mercado', 'marca', 'lat', 'lng', 'precision_m', 'foto'],
  Incidencias: ['id', 'registrado', 'fecha', 'hora', 'canjista', 'mercado', 'tipo_incidencia', 'detalle', 'foto'],
  Equipo: ['nombre', 'rol', 'mercado', 'activo', 'usuario', 'contrasena', 'celular'],
  // Una fila por mercado y drop. mercado = "otros" aplica a los mercados que no tengan filas propias.
  Drops: ['mercado', 'drop', 'paquetes_min', 'premio_familiar', 'premio_personal'],
  Config: ['clave', 'valor', 'nota'],
};
const COLUMNAS_TEXTO = ['id', 'fecha', 'hora', 'usuario', 'contrasena', 'celular', 'valor'];

/* ------------------------------------------------------------------ */
/* Instalación                                                         */
/* ------------------------------------------------------------------ */

function setup() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss && P.getProperty('SS_ID')) ss = SpreadsheetApp.openById(P.getProperty('SS_ID'));
  if (!ss) ss = SpreadsheetApp.create('Canje Mayorista AJE · Base de datos');
  P.setProperty('SS_ID', ss.getId());

  Object.keys(ESQUEMA).forEach(function (nombre) {
    const cols = ESQUEMA[nombre];
    const sh = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    cols.forEach(function (c, i) {
      if (COLUMNAS_TEXTO.indexOf(c) >= 0) sh.getRange(1, i + 1, sh.getMaxRows(), 1).setNumberFormat('@');
    });
  });
  ['Hoja 1', 'Sheet1', 'Hoja1'].forEach(function (n) {
    const sh = ss.getSheetByName(n);
    if (sh && ss.getSheets().length > 1 && sh.getLastRow() <= 1) ss.deleteSheet(sh);
  });

  const dr = ss.getSheetByName('Drops');
  if (dr.getLastRow() < 2) {
    const prod = [[24, 3, 8], [50, 6, 15], [150, 9, 30], [500, 15, 45], [800, 20, 60]];
    const otros = [[12, 3, 6], [25, 6, 12], [50, 9, 18], [80, 15, 24], [150, 20, 30]];
    const filas = prod.map(function (r, i) { return ['Productores', i + 1].concat(r); })
      .concat(otros.map(function (r, i) { return ['otros', i + 1].concat(r); }));
    dr.getRange(2, 1, filas.length, 5).setValues(filas);
  }

  const eq = ss.getSheetByName('Equipo');
  if (eq.getLastRow() < 2) {
    const filas = MERCADOS.map(function (m) { return ['Canjista ' + m, 'canjista', m, 'SI', USUARIO_MERCADO[m], claveAleatoria_(), '']; });
    filas.push(['Supervisor', 'supervisor', '', 'SI', 'supervisor', claveAleatoria_(), '']);
    filas.push(['Coordinador Free', 'coordinador', '', 'SI', 'coordinador', claveAleatoria_(), '']);
    eq.getRange(2, 1, filas.length, 7).setValues(filas);
  }

  const cfg = ss.getSheetByName('Config');
  if (cfg.getLastRow() < 2) {
    cfg.getRange(2, 1, 3, 3).setValues([
      ['fecha_inicio', '', 'AAAA-MM-DD · primer día de esta etapa'],
      ['cliente_usuario', 'aje', 'Usuario del dashboard del cliente'],
      ['cliente_clave', claveAleatoria_(), 'Contraseña del dashboard del cliente'],
    ]);
  }

  if (!P.getProperty('FOLDER_ID')) {
    const folder = DriveApp.createFolder('Canje Mayorista AJE · Evidencias');
    try { folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { }
    P.setProperty('FOLDER_ID', folder.getId());
  }
  if (!P.getProperty('DASH_TOKEN')) P.setProperty('DASH_TOKEN', Utilities.getUuid().replace(/-/g, '').slice(0, 20));
  verAccesos();
}

function verAccesos() {
  Logger.log('Google Sheet: ' + ss_().getUrl());
  Logger.log('Fotos: https://drive.google.com/drive/folders/' + P.getProperty('FOLDER_ID'));
  Logger.log('Accesos de canjistas: pestaña Equipo (usuario y contrasena)');
  Logger.log('Token del dashboard: ' + P.getProperty('DASH_TOKEN'));
}

/* ------------------------------------------------------------------ */
/* Endpoints                                                           */
/* ------------------------------------------------------------------ */

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === 'data') {
      if (!p.k || p.k !== P.getProperty('DASH_TOKEN')) return out_({ ok: false, code: 'auth', error: 'Link no autorizado' });
      return out_(datos_(false));
    }
    return out_({ ok: true, servicio: 'Canje Mayorista AJE' });
  } catch (err) {
    return out_({ ok: false, code: 'server', error: String(err && err.message || err) });
  }
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, code: 'val', error: 'Datos inválidos' }); }
  try {
    if (b.tipo === 'login_cliente') {
      const c = config_();
      const ok = normUsuario_(b.usuario) === normUsuario_(c.cliente_usuario) && normClave_(b.clave) !== '' && normClave_(b.clave) === normClave_(c.cliente_clave);
      return out_(ok ? { ok: true, k: P.getProperty('DASH_TOKEN') } : { ok: false, code: 'auth', error: 'Usuario o contraseña incorrectos' });
    }
    const u = auth_(b);
    if (!u) return out_({ ok: false, code: 'auth', error: 'Usuario o contraseña incorrectos' });
    if (b.tipo === 'login') return out_({ ok: true, usuario: u, drops: drops_() });
    if (b.tipo === 'datos_interno') {
      if (!esJefe_(u)) return out_({ ok: false, code: 'auth', error: 'El acceso interno es solo para el supervisor y el coordinador' });
      return out_(Object.assign(datos_(true), { usuario: u }));
    }
    const fn = { asistencia: asistencia_, canje: canje_, incidencia: incidencia_ }[b.tipo];
    if (!fn) return out_({ ok: false, code: 'val', error: 'Tipo de registro desconocido' });
    if (!b.id) return out_({ ok: false, code: 'val', error: 'Falta el id del registro' });
    return out_(fn(b, u));
  } catch (err) {
    return out_({ ok: false, code: err && err.val ? 'val' : 'server', error: String(err && err.message || err) });
  }
}

/* ------------------------------------------------------------------ */
/* Registros                                                           */
/* ------------------------------------------------------------------ */

function mercadoDe_(b, u) {
  const m = esJefe_(u) ? limpio_(b.mercado) : u.mercado;
  if (MERCADOS.indexOf(m) < 0) throw val_('Mercado inválido: ' + m);
  return m;
}

function asistencia_(b, u) {
  if (existeId_('Asistencia', b.id)) return { ok: true, duplicado: true };
  const m = mercadoDe_(b, u);
  if (['Entrada', 'Salida'].indexOf(b.marca) < 0) throw val_('Marca inválida');
  if (!b.foto) throw val_('Falta la selfie');
  const t = cuando_(b.capturado), subidas = [];
  const foto = foto_(b.foto, t.fecha + '_' + m + '_' + b.marca, t.fecha, subidas);
  return conLock_(function () {
    if (existeId_('Asistencia', b.id)) { borrar_(subidas); return { ok: true, duplicado: true }; }
    agregar_('Asistencia', { id: b.id, registrado: ahora_(), fecha: t.fecha, hora: t.hora, canjista: u.nombre, mercado: m,
      marca: b.marca, lat: num_(b.lat), lng: num_(b.lng), precision_m: num_(b.acc), foto: foto });
    return { ok: true };
  });
}

function canje_(b, u) {
  if (existeId_('Canjes', b.id)) return { ok: true, duplicado: true };
  const m = mercadoDe_(b, u);
  const formato = String(b.formato || '').toUpperCase();
  if (['FAMILIAR', 'PERSONAL'].indexOf(formato) < 0) throw val_('Formato inválido');
  const bebida = limpio_(b.bebida).toUpperCase();
  if (!bebida) throw val_('Falta la bebida');
  const paq = Math.floor(Number(b.paquetes));
  if (!(paq > 0)) throw val_('Paquetes inválidos');
  const r = calcular_(m, formato, paq);
  if (!r) throw val_('Con ' + paq + ' paquetes no alcanza el drop 1');
  if (!b.foto_boleta || !b.foto_premio) throw val_('Faltan las fotos del canje');
  const t = cuando_(b.capturado), subidas = [];
  const base = t.fecha + '_' + t.hora.replace(':', '') + '_' + m;
  const fb = foto_(b.foto_boleta, base + '_boleta', t.fecha, subidas);
  const fp = foto_(b.foto_premio, base + '_premio', t.fecha, subidas);
  return conLock_(function () {
    if (existeId_('Canjes', b.id)) { borrar_(subidas); return { ok: true, duplicado: true }; }
    agregar_('Canjes', { id: b.id, registrado: ahora_(), fecha: t.fecha, hora: t.hora, canjista: u.nombre, mercado: m,
      cliente: limpio_(b.cliente), formato: formato, bebida: bebida, paquetes: paq, drop: r.drop, premio: r.premio,
      foto_boleta: fb, foto_premio: fp, nota: limpio_(b.nota), lat: num_(b.lat), lng: num_(b.lng) });
    return { ok: true, drop: r.drop, premio: r.premio };
  });
}

function incidencia_(b, u) {
  if (existeId_('Incidencias', b.id)) return { ok: true, duplicado: true };
  const m = mercadoDe_(b, u);
  const tipo = limpio_(b.tipo_incidencia);
  if (!tipo) throw val_('Falta el tipo de incidencia');
  const t = cuando_(b.capturado), subidas = [];
  const foto = b.foto ? foto_(b.foto, t.fecha + '_' + m + '_incidencia', t.fecha, subidas) : '';
  return conLock_(function () {
    if (existeId_('Incidencias', b.id)) { borrar_(subidas); return { ok: true, duplicado: true }; }
    agregar_('Incidencias', { id: b.id, registrado: ahora_(), fecha: t.fecha, hora: t.hora, canjista: u.nombre, mercado: m,
      tipo_incidencia: tipo, detalle: limpio_(b.detalle), foto: foto });
    return { ok: true };
  });
}

/**
 * Datos para el dashboard. Cliente (AJE): canjes sin ubicación e incidencias sin detalle ni foto.
 * Interno (Free): todo, más la asistencia.
 */
function datos_(interno) {
  const ini = config_().fecha_inicio || '';
  const f = function (r) { return !ini || r.fecha >= ini; };
  const canjes = leer_('Canjes').filter(f).map(function (r) { if (!interno) { delete r.lat; delete r.lng; } return r; });
  const incidencias = leer_('Incidencias').filter(f).map(function (r) {
    return interno ? r : { fecha: r.fecha, hora: r.hora, mercado: r.mercado, tipo_incidencia: r.tipo_incidencia };
  });
  const o = { ok: true, generado: ahora_(), canjes: canjes, incidencias: incidencias, drops: drops_() };
  if (interno) o.asistencia = leer_('Asistencia').filter(f);
  return o;
}

function esJefe_(u) { return /^(supervisor|coordinador)/.test(String(u && u.rol || '')); }

/* ------------------------------------------------------------------ */
/* Mecánica                                                            */
/* ------------------------------------------------------------------ */

/** { Productores: {paq:[..], FAMILIAR:[..], PERSONAL:[..]}, otros: {...} } desde la hoja Drops. */
function drops_() {
  const o = {};
  leer_('Drops').forEach(function (r) {
    const k = limpio_(r.mercado) || 'otros', i = Number(r.drop) - 1;
    if (!(i >= 0 && i < 10)) return;
    o[k] = o[k] || { paq: [], FAMILIAR: [], PERSONAL: [] };
    o[k].paq[i] = Number(r.paquetes_min); o[k].FAMILIAR[i] = Number(r.premio_familiar); o[k].PERSONAL[i] = Number(r.premio_personal);
  });
  return o;
}

function calcular_(mercado, formato, paq) {
  const d = drops_(), t = d[mercado] || d.otros;
  if (!t) throw new Error('La hoja Drops está vacía');
  let i = -1;
  t.paq.forEach(function (min, n) { if (paq >= min) i = n; });
  return i < 0 ? null : { drop: i + 1, premio: t[formato][i] };
}

/* ------------------------------------------------------------------ */
/* Utilidades                                                          */
/* ------------------------------------------------------------------ */

function ss_() { return SpreadsheetApp.openById(P.getProperty('SS_ID')); }
function hoja_(n) { return ss_().getSheetByName(n); }

function leer_(nombre) {
  const sh = hoja_(nombre);
  if (!sh || sh.getLastRow() < 2) return [];
  const v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getDisplayValues();
  const h = v.shift();
  return v.filter(function (r) { return r[0] !== ''; }).map(function (r) {
    const o = {}; h.forEach(function (k, i) { o[k] = r[i]; }); return o;
  });
}

function agregar_(nombre, obj) {
  hoja_(nombre).appendRow(ESQUEMA[nombre].map(function (k) { return obj[k] === undefined || obj[k] === null ? '' : obj[k]; }));
}

function existeId_(nombre, id) {
  const sh = hoja_(nombre);
  if (sh.getLastRow() < 2 || !id) return false;
  return !!sh.getRange(2, 1, sh.getLastRow() - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
}

function auth_(b) {
  const usuario = normUsuario_(b.usuario), clave = normClave_(b.clave);
  if (!usuario || !clave) return null;
  const r = leer_('Equipo').find(function (x) {
    return normUsuario_(x.usuario) === usuario && normClave_(x.contrasena) === clave && !/^no$/i.test(x.activo);
  });
  return r ? { nombre: r.nombre, rol: String(r.rol || 'canjista').toLowerCase(), mercado: r.mercado, usuario: r.usuario } : null;
}

function config_() { const o = {}; leer_('Config').forEach(function (r) { if (r.clave) o[r.clave] = r.valor; }); return o; }

function foto_(dataUrl, nombre, fecha, subidas) {
  if (!dataUrl) return '';
  const m = String(dataUrl).match(/^data:(image\/[a-z+]+);base64,(.+)$/);
  if (!m) throw val_('Formato de foto inválido');
  const ext = m[1].split('/')[1].replace('jpeg', 'jpg');
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], String(nombre).replace(/[^\w\-áéíóúñÁÉÍÓÚÑ]+/g, '_') + '.' + ext);
  const f = carpetaDia_(fecha).createFile(blob);
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { }
  subidas.push(f);
  return 'https://drive.google.com/file/d/' + f.getId() + '/view';
}

function carpetaDia_(fecha) {
  const raiz = DriveApp.getFolderById(P.getProperty('FOLDER_ID'));
  const it = raiz.getFoldersByName(fecha);
  return it.hasNext() ? it.next() : raiz.createFolder(fecha);
}

function borrar_(archivos) { archivos.forEach(function (f) { try { f.setTrashed(true); } catch (e) { } }); }
function conLock_(fn) { const l = LockService.getScriptLock(); l.waitLock(30000); try { return fn(); } finally { l.releaseLock(); } }
function cuando_(iso) {
  let d = iso ? new Date(iso) : new Date();
  if (isNaN(d.getTime()) || d.getTime() - Date.now() > 3600 * 1000) d = new Date();
  return { fecha: Utilities.formatDate(d, TZ, 'yyyy-MM-dd'), hora: Utilities.formatDate(d, TZ, 'HH:mm') };
}
function ahora_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function limpio_(s) { return String(s === undefined || s === null ? '' : s).trim().slice(0, 500); }
function normUsuario_(s) { return limpio_(s).toLowerCase().normalize('NFD').replace(/[^a-z0-9]/g, ''); }
function normClave_(s) { return limpio_(s).toLowerCase().replace(/\s+/g, ''); }
function num_(x) { const n = Number(x); return isFinite(n) && x !== '' && x !== null && x !== undefined ? n : ''; }
function val_(msg) { const e = new Error(msg); e.val = true; return e; }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function claveAleatoria_() {
  const abc = 'abcdefghjkmnpqrstuvwxyz23456789'; let s = '';
  for (let i = 0; i < 8; i++) s += abc.charAt(Math.floor(Math.random() * abc.length));
  return s;
}
