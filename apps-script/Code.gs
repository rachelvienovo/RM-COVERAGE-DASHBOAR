/**
 * RM Coverage Dashboards: live sync backend (read-only, key protected).
 *
 * Reads the two source spreadsheets and returns the same JSON the dashboards embed,
 * so the "Sync now" button can refresh a page without a rebuild.
 *
 *   GET ?action=status&key=KEY   -> sync timestamps of both sheets (cheap)
 *   GET ?action=data&key=KEY     -> full dataset for all four dashboards
 *
 * Setup: see SETUP-LIVE-SYNC.md.
 */
var WH_ID = '1wFwiKAfhgkTKgL3QvKj8yn2jWyxGw55infzXfJtZXsI';   // RM_WH_Mgt
var PBA_ID = '1uavWZvYKFIivSr4Sz8NXDUdDZuQsSNH-PI6wH3LF4kQ';  // SBO Purchase Blanket Agreement
var AUDIT_ID = '17dwpOMvXaS3Rx1jmpX9Hf18X-74JoOLty5EapFxnbFY'; // Inventory Price: IMPORTRANGE copy of Inventory (Ageing), tab Sheet1

var PLANTS = ['AC', 'PFMIS', 'BUKID', 'SOUTH', 'HOREB', 'ARGAO'];
// plant -> its MIN_ sheet (consumption = AV W Fcasts consump x WEEKS_PER_MONTH)
var MIN_SHEETS = { MIN_BULACAN: 'AC', MIN_ISA: 'PFMIS', MIN_HOREB: 'HOREB', MIN_ARGAO: 'ARGAO', MIN_BUKID: 'BUKID', MIN_SOUTH: 'SOUTH', MIN_CCPC: 'BUKID' };
var WEEKS_PER_MONTH = 4;
// class, label, unit, divisor from sheet units, decimals, stock decimals
var GROUPS = [
  ['MACRO', 'Macro', 't', 1000, 0, 1],
  ['PACKAGING', 'Packaging', 'pcs', 1, 0, 0],
  ['MICRO', 'Micro', 'kg', 1, 0, 0],
  ['FEEDTAGS', 'Feedtags', 'pcs', 1, 0, 0]
];

function isAdmin_(key) {
  var real = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY') || '';
  if (!real || real.length < 16 || !key || key.length !== real.length) return false;
  var diff = 0;
  for (var i = 0; i < real.length; i++) diff |= real.charCodeAt(i) ^ key.charCodeAt(i);
  return diff === 0;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (!isAdmin_(p.key)) return json_({ ok: false, error: 'unauthorized' });
  try {
    if (p.action === 'status') return json_({ ok: true, status: status_() });
    return json_({ ok: true, data: build_() });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function code_(x) {
  var n = Number(x);
  return (x !== '' && x !== null && !isNaN(n)) ? String(Math.round(n)) : String(x);
}
function round3_(v) { return Math.round(v * 1000) / 1000; }
function ymd_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function dateIn_(txt) {
  var m = /(\d{2})\/(\d{2})\/(\d{4})/.exec(String(txt));
  return m ? new Date(Number(m[3]), Number(m[1]) - 1, Number(m[2])) : null;
}

/** Warehouse code -> warehouse, as given by the business. Quarantine (-QA) and in-transit (-IT) codes are not counted. */
var WH_MAP = {};
('AC AC|ACEXT AC|ACEXT-MK AC|AC-MK AC|ACOW AC|ACOW-MK AC|AC-PD AC|ALAE BUKID|ARGAO ARGAO|ARGAO-PD ARGAO|BAC HOREB|' +
 'BUKID BUKID|BUKID-MK BUKID|BUKID-PD BUKID|CCPC BUKID|HBEXT HOREB|HOREB HOREB|HOREB-MK HOREB|HOREB-PD HOREB|' +
 'PFMIS PFMIS|PFMIS-PD PFMIS|SOUTH SOUTH|SOUTH-PD SOUTH').split('|').forEach(function (p) { var a = p.split(' '); WH_MAP[a[0]] = a[1]; });
function plantOfWh_(w) { return w ? (WH_MAP[String(w)] || null) : null; }

function costSync_(a2, d2) {
  return String(a2) + (d2 instanceof Date ? ' | as of ' + Utilities.formatDate(d2, Session.getScriptTimeZone(), 'MM/dd/yyyy') : '');
}

function status_() {
  var inv = SpreadsheetApp.openById(WH_ID).getSheetByName('INV').getRange('A2').getValue();
  var pba = SpreadsheetApp.openById(PBA_ID).getSheetByName('OPEN PBA + PO').getRange('A3').getValue();
  var cs = SpreadsheetApp.openById(AUDIT_ID).getSheetByName('Sheet1');
  return { invSyncFull: String(inv), pbaSyncFull: String(pba), costSyncFull: costSync_(cs.getRange('A2').getValue(), cs.getRange('D2').getValue()) };
}

function zeroes_() { var o = {}; PLANTS.forEach(function (p) { o[p] = 0; }); return o; }

function build_() {
  var wh = SpreadsheetApp.openById(WH_ID);
  var pbaSs = SpreadsheetApp.openById(PBA_ID);
  var values = function (ss, name) { return ss.getSheetByName(name).getDataRange().getValues(); };

  var invRows = values(wh, 'INV');
  var invSync = String(invRows[1][0]);
  var asof = dateIn_(invSync) || new Date();
  var openRows = values(pbaSs, 'OPEN PBA + PO');
  var pbaSync = String(openRows[2][0]);

  // end dates of each PBA line (RM_WH_Mgt > PBA)
  var ends = {};
  values(wh, 'PBA').slice(3).forEach(function (r) {
    if (typeof r[1] === 'number' && r[3] instanceof Date) ends[Math.round(r[1]) + '|' + code_(r[6]) + '|' + r[8]] = r[3];
  });

  // Cost per warehouse from Inventory (Ageing): Cost (column M) weighted by QTY in KG (column J).
  // Stock quantities stay as in RM_WH_Mgt.
  var auditRows = values(SpreadsheetApp.openById(AUDIT_ID), 'Sheet1');
  var costSums = {};
  auditRows.slice(4).forEach(function (r) {
    if (!r[0]) return;
    var p = plantOfWh_(r[4]);
    if (p && typeof r[9] === 'number' && typeof r[12] === 'number' && r[9] > 0 && r[12] > 0) {
      var k = r[0] + '|' + code_(r[3]) + '|' + p;
      var a = costSums[k] || (costSums[k] = [0, 0]);
      a[0] += r[9] * r[12]; a[1] += r[9];
    }
  });

  // Warehouse for pending lines whose Warehouse is blank in OPEN PBA + PO. From the same workbook, tab PBA Monitoring:
  // the agreement's own Warehouse if filled, else the warehouse (PO_Whs) of its POs that are still undelivered.
  // If that tab has nothing, RM_WH_Mgt > OPEN_PO is used (comment 'Origin: Blanket Agreement N.').
  // poWh[pba|item][warehouse] = [weight, unit price]
  var own = {}, viaPo = {}, viaOpo = {}, poWh = {};
  values(pbaSs, 'PBA Monitoring').slice(5).forEach(function (r) {
    if (typeof r[0] !== 'number' || r[6] === '') return;
    var k = Math.round(r[0]) + '|' + code_(r[6]);
    if (r[3] && plantOfWh_(r[3])) { own[k] = {}; own[k][r[3]] = [1, r[12]]; }
    else if (typeof r[19] === 'number' && r[19] > 0 && plantOfWh_(r[18])) {
      var d = viaPo[k] || (viaPo[k] = {}), a = d[r[18]] || (d[r[18]] = [0, r[12]]);
      a[0] += r[19];
    }
  });
  values(wh, 'OPEN_PO').slice(3).forEach(function (r) {
    var m = /Blanket Agreement (\d+)/.exec(String(r[12] || ''));
    if (m && r[5] !== '' && typeof r[9] === 'number' && r[9] > 0 && plantOfWh_(r[7])) {
      var k = m[1] + '|' + code_(r[5]), d = viaOpo[k] || (viaOpo[k] = {}), a = d[r[7]] || (d[r[7]] = [0, r[10]]);
      a[0] += r[9];
    }
  });
  [viaOpo, viaPo, own].forEach(function (src) {   // later sources override earlier ones
    Object.keys(src).forEach(function (k) {
      poWh[k] = {};
      Object.keys(src[k]).forEach(function (w) { poWh[k][w] = [src[k][w][0], src[k][w][1]]; });
    });
  });
  Object.keys(poWh).forEach(function (k) {         // a missing price falls back to the PO's unit price in OPEN_PO
    var pr = 0;
    if (viaOpo[k]) Object.keys(viaOpo[k]).forEach(function (w) { if (!pr && viaOpo[k][w][1]) pr = viaOpo[k][w][1]; });
    Object.keys(poWh[k]).forEach(function (w) { if (!poWh[k][w][1]) poWh[k][w][1] = pr || 0; });
  });

  var minData = {};
  Object.keys(MIN_SHEETS).forEach(function (n) { minData[n] = values(wh, n); });

  var out = { asof: ymd_(asof), invSync: Utilities.formatDate(asof, Session.getScriptTimeZone(), 'MMM dd'),
    invSyncFull: invSync, pbaSyncFull: pbaSync, costSyncFull: costSync_(auditRows[1][0], auditRows[1][3]), plants: PLANTS, groups: {} };
  var pbaDate = dateIn_(pbaSync);
  out.pbaSync = Utilities.formatDate(pbaDate || asof, Session.getScriptTimeZone(), 'MMM dd');

  GROUPS.forEach(function (g) {
    var cls = g[0], div = g[3], items = {};
    var item = function (c, name) {
      if (!items[c]) items[c] = { code: c, name: name, cons: zeroes_(), stock: zeroes_(), pend: zeroes_(), pos: [], wh: {} };
      return items[c];
    };
    // consumption: average weekly forecast x 4, per plant MIN_ sheet
    Object.keys(MIN_SHEETS).forEach(function (n) {
      var rows = minData[n], hdr = -1, ix = -1;
      for (var i = 0; i < Math.min(12, rows.length); i++) {
        if (rows[i][0] === 'INV_GROUP') { hdr = i; ix = rows[i].indexOf('AV W Fcasts consump'); break; }
      }
      if (hdr < 0 || ix < 0) throw new Error('Header not found in ' + n);
      for (var j = hdr + 1; j < rows.length; j++) {
        var r = rows[j];
        if (r[0] === cls && r[1] !== '') {
          var w = r[ix];
          item(code_(r[1]), r[2]).cons[MIN_SHEETS[n]] += (typeof w === 'number' ? w : 0) * WEEKS_PER_MONTH;
        }
      }
    });
    // stock on hand
    invRows.slice(3).forEach(function (r) {
      if (r[2] === cls && r[3] !== '') {
        var p = plantOfWh_(r[1]);
        if (p) {
          var it = item(code_(r[3]), r[4]);
          var q = Number(r[5]) || 0;
          it.stock[p] += q;
          it.wh[r[1]] = (it.wh[r[1]] || 0) + q;
        }
      }
    });
    // pending PBA + PO balance
    openRows.slice(5).forEach(function (r) {
      if (!r[1] || r[5] === '') return;
      var c = code_(r[5]);
      if (!items[c]) return;
      var bal = Number(r[9]) || 0;
      if (bal <= 0) return;
      var pba = Math.round(Number(r[1])), e = ends[pba + '|' + c + '|' + r[7]];
      var targets = [];
      if (plantOfWh_(r[7])) targets.push([r[7], 1, r[11]]);
      else if (poWh[pba + '|' + c]) {
        var d = poWh[pba + '|' + c], tot = 0, w;
        for (w in d) tot += d[w][0];
        for (w in d) targets.push([w, d[w][0] / tot, d[w][1]]);
      }
      targets.forEach(function (t) {
        var p = plantOfWh_(t[0]);
        if (!p) return;
        items[c].pend[p] += bal * t[1];
        items[c].pos.push({ pba: pba, vendor: r[4], wh: t[0], plant: p,
          created: r[2] instanceof Date ? ymd_(r[2]) : String(r[2]), end: e ? ymd_(e) : null,
          t: round3_(bal * t[1] / div), price: t[2] || r[11], expired: !!(e && e < asof) });
      });
    });

    var round = function (o) { var x = {}; PLANTS.forEach(function (p) { x[p] = round3_(o[p] / div); }); return x; };
    var list = Object.keys(items).map(function (c) {
      var it = items[c], wh = {};
      Object.keys(it.wh).forEach(function (k) { if (it.wh[k]) wh[k] = round3_(it.wh[k] / div); });
      var cost = {};
      PLANTS.forEach(function (p) { var a = costSums[cls + '|' + c + '|' + p]; if (a) cost[p] = Math.round(a[0] / a[1] * 10000) / 10000; });
      return { code: c, name: it.name, cons: round(it.cons), stock: round(it.stock), pend: round(it.pend), stockWh: wh, cost: cost,
        pos: it.pos.sort(function (a, b) { return b.t - a.t; }) };
    }).filter(function (it) {
      var s = 0; PLANTS.forEach(function (p) { s += it.cons[p] + it.stock[p] + it.pend[p]; });
      return s > 0;
    }).sort(function (a, b) {
      var an = /^\d+$/.test(a.code), bn = /^\d+$/.test(b.code);
      if (an !== bn) return an ? -1 : 1;
      return an ? Number(a.code) - Number(b.code) : (a.code < b.code ? -1 : 1);
    });
    out.groups[cls] = { label: g[1], unit: g[2], dec: g[4], stockDec: g[5], div: div, items: list };
  });
  return out;
}
