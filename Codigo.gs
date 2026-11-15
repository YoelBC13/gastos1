// Script vinculado a la hoja (Extensiones > Apps Script).
var API_TOKEN = ''; // Opcional: si lo definís, la PWA externa debe enviar el mismo token.

function doGet(e) {
  if (e && e.parameter && e.parameter.fn) return handle_(e.parameter, false);
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Gastos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  var p;
  try { p = JSON.parse(e.postData.contents); } catch (err) { p = {}; }
  return handle_(p, true);
}

function handle_(p, isPost) {
  var out;
  try {
    if (API_TOKEN && p.token !== API_TOKEN) throw new Error('No autorizado');
    if (p.fn === 'getInit') out = { data: getInit() };
    else if (p.fn === 'saveExpense' && isPost) out = { data: saveExpense(p) };
    else throw new Error('Operación inválida');
  } catch (err) {
    out = { error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function getInit() {
  return { cats: cats_(), resumen: resumen_() };
}

function cats_() {
  var sh = ss_().getSheetByName('Config');
  var n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 1).getValues()
    .map(function (r) { return String(r[0]).trim(); })
    .filter(function (s) { return s; });
}

function saveExpense(d) {
  var m = Number(String(d.monto).replace(',', '.'));
  if (!(m > 0)) throw new Error('Monto inválido');
  var det = String(d.detalle || '').slice(0, 120);
  if (/^[=+\-@]/.test(det)) det = "'" + det; // evita fórmulas inyectadas
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    ss_().getSheetByName('Gastos').appendRow([new Date(), det, m, String(d.categoria || '')]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return { resumen: resumen_() };
}

function resumen_() {
  var sh = ss_().getSheetByName('Gastos');
  var n = sh.getLastRow() - 1;
  var rows = n > 0 ? sh.getRange(2, 1, n, 4).getValues() : []; // lectura masiva
  var tz = Session.getScriptTimeZone();
  var now = new Date();
  var ym = Utilities.formatDate(now, tz, 'yyyy-MM');
  var hoy = Number(Utilities.formatDate(now, tz, 'd'));
  var total = 0, cat = {}, dia = [], all = [], i;
  for (i = 0; i < hoy; i++) dia.push(0);

  rows.forEach(function (r, idx) {
    var d = r[0] instanceof Date ? r[0] : new Date(r[0]);
    var m = Number(r[2]);
    if (isNaN(d.getTime()) || isNaN(m)) return;
    var c = String(r[3]).trim() || 'Sin categoría';
    all.push({ t: d.getTime(), i: idx, d: String(r[1]), c: c, f: Utilities.formatDate(d, tz, 'dd/MM'), m: m });
    if (Utilities.formatDate(d, tz, 'yyyy-MM') === ym) {
      total += m;
      cat[c] = (cat[c] || 0) + m;
      var k = Number(Utilities.formatDate(d, tz, 'd'));
      if (k <= hoy) dia[k - 1] += m;
    }
  });

  var evo = [], acc = 0;
  dia.forEach(function (v) { acc += v; evo.push(acc); });

  var lista = Object.keys(cat).map(function (c) { return { c: c, m: cat[c] }; })
    .sort(function (a, b) { return b.m - a.m; });
  var top = lista.slice(0, 4).map(function (x) {
    return { c: x.c, m: x.m, p: total ? Math.round(x.m / total * 100) : 0 };
  });
  var dona = lista.slice(0, 4);
  var otras = lista.slice(4).reduce(function (s, x) { return s + x.m; }, 0);
  if (otras > 0) dona.push({ c: 'Otras', m: otras });

  all.sort(function (a, b) { return b.t - a.t || b.i - a.i; });
  return { total: total, top: top, dona: dona, evo: evo, last: all.slice(0, 8) };
}
