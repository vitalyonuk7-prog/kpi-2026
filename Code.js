/*************************************************
 * Code.gs — KPI backend + меню + веб-апка + PDF
 *************************************************/

/******************** MENU: вкладка KPI в документі ********************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('KPI')
    .addItem('Відкрити панель KPI', 'openKpiDialog')
    .addItem('Веб-апка (повний екран)', 'openKpiInNewTab_')
    .addSeparator()
    .addItem('Показати URL веб-апки', 'showWebAppUrl')
    .addToUi();
}

/** Діалог у Google Sheets */
function openKpiDialog() {
  var html = HtmlService.createHtmlOutputFromFile('kpi_dash')
    .setWidth(1280).setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(html, 'KPI');
}
/** Alias (якщо десь було прив’язано стару назву) */
function openKpiDialogDetails() { openKpiDialog(); }

/** Надійне відкриття деплою в новій вкладці, без «Drive viewer» */
function openKpiInNewTab_() {
  var url = getWebAppUrl();
  var html = HtmlService.createHtmlOutput(
    '<!doctype html><html><head><meta charset="utf-8"><base target="_top">' +
    '<style>body{font:14px system-ui;margin:18px} .btn{display:inline-block;padding:10px 14px;border-radius:999px;border:1px solid #e5e7eb;background:#fff;cursor:pointer;text-decoration:none} .hint{color:#64748b;margin-top:8px}</style>' +
    '</head><body>' +
    '<a class="btn" href="'+url+'">Open KPI</a>' +
    '<div class="hint">If nothing happens, click the button above.</div>' +
    '<script>' +
    'try{var w=window.open('+JSON.stringify(url)+', "_blank", "noopener");' +
    'if(!w||w.closed||typeof w.closed==="undefined"){ top.location.href='+JSON.stringify(url)+'; }}catch(e){ top.location.href='+JSON.stringify(url)+'; }' +
    'setTimeout(function(){ window.close && window.close(); }, 4000);' +
    '</script></body></html>'
  ).setWidth(360).setHeight(120);
  SpreadsheetApp.getUi().showModalDialog(html, 'Opening KPI…');
}

/** Показати URL */
function showWebAppUrl(){
  SpreadsheetApp.getUi().alert('Full Screen URL', getWebAppUrl() || '(no URL)', SpreadsheetApp.getUi().ButtonSet.OK);
}

/******************** Web entrypoint (для Full screen) ********************/
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('kpi_dash')
    .setTitle('KPI')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Отримати URL деплою (з пріоритетом на ScriptApp.getService().getUrl) */
function getWebAppUrl(){
  return kpiGetWebAppUrl();
}
/** URL веб-апки з fallback-ами (твій exec як запасний) */
function kpiGetWebAppUrl() {
  var hardcoded = 'https://script.google.com/a/macros/bets.io/s/AKfycbwm6zh3qw00OXpug7hqiLc79Szs5oaKlsj9rqqT-tTbINUho57yqolUb51fdSmcVlhP/exec';
  try {
    var url = ScriptApp.getService().getUrl();
    if (url) return url;
  } catch (e) {}
  var prop = PropertiesService.getScriptProperties().getProperty('KPI_WEBAPP_URL');
  if (prop) return prop;
  return hardcoded;
}
/** Якщо треба — один раз збережи URL вручну */
function kpiSetWebAppUrl(url) {
  PropertiesService.getScriptProperties().setProperty('KPI_WEBAPP_URL', String(url || '').trim());
  return true;
}

/******************** ПРОЄКТИ / АВТО-ДЕТЕКТ АРКУШІВ ********************/
/** Префікс імен аркушів за проектом */
function kpiSheetsFor(project) {
  var p = String(project || 'BETS.IO').toUpperCase();
  var prefix = (p === 'BETSIO.COM') ? '[BETSIO.COM] ' : '[BETS.IO] ';
  return { prefix: prefix };
}
/** Знаходимо всі MONTHLY/WEEKLY аркуші з відповідним префіксом */
function kpiFindSheetsForProject_(project) {
  var prefix = kpiSheetsFor(project).prefix;
  var sheets = SpreadsheetApp.getActive().getSheets();
  var monthly = [], weekly = [];
  for (var i=0;i<sheets.length;i++){
    var name = sheets[i].getName();
    if (!name || name.indexOf(prefix)!==0) continue;
    var up = String(name).toUpperCase();
    if (up.indexOf('MONTHLY')!==-1) monthly.push(name);
    if (up.indexOf('WEEKLY') !==-1) weekly.push(name);
  }
  return { monthly: monthly, weekly: weekly };
}

/** Основний вхід з фронту: ігноруємо segment, просто агрегуємо все MONTHLY/WEEKLY */
function kpiLoadProject(project /*, segment */) {
  var found = kpiFindSheetsForProject_(project);
  var mon = readMonthlyDynamicMulti_(found.monthly);
  var wk  = readWeeklyDynamicMulti_ (found.weekly);

  var months = [];
  var ymKeys = {};
  Object.keys(mon.map).forEach(function(k){ ymKeys[k]=1; });
  Object.keys(wk.map ).forEach(function(k){ ymKeys[k]=1; });

  Object.keys(ymKeys).sort().forEach(function(ym){
    var label = mon.map[ym] ? mon.map[ym].label : monthLabelFromYm_(ym);
    var monthly = emptyFromNames_(mergeNameSets_(mon.metrics));
    if (mon.map[ym]) copyAdd_(monthly, mon.map[ym].monthly);

    var weeksArr = [];
    if (wk.map[ym]) {
      var weekKeys = Object.keys(wk.map[ym]).map(function(s){return +s;}).sort(function(a,b){return a-b;});
      for (var i=0; i<weekKeys.length; i++){
        var ts  = weekKeys[i];
        var row = wk.map[ym][ts];
        var dIso = row && row.date ? String(row.date) : null;
        row.label = 'Week ' + (i+1) + (dIso ? (' (' + dIso + ')') : '');
        row.weekNo = (i+1);
        weeksArr.push(row);
      }
    }

    var sumWeeks = emptyFromNames_(mergeNameSets_(wk.metrics));
    weeksArr.forEach(function(w){ copyAdd_(sumWeeks, w.weekly); });

    months.push({ ym: ym, label: label, monthly: monthly, weeks: weeksArr, sumWeeks: sumWeeks });
  });

  var metricsObj = mergeNameSets_(mergeNameSets_(mon.metrics), wk.metrics);
  var metricsArr = Object.keys(metricsObj || {});
  var nonZero = {};
  metricsArr.forEach(function (m) { nonZero[m] = hasNonZero_(months, m); });

  return { metrics: metricsArr, months: months, nonZero: nonZero };
}

/******************** АГРЕГАЦІЯ З КІЛЬКОХ АРКУШІВ ********************/
function readMonthlyDynamicMulti_(sheetNames){
  var total = { map: {}, metrics: {} };
  (sheetNames || []).forEach(function(name){
    var one = readMonthlyDynamic_(name);
    Object.keys(one.metrics||{}).forEach(function(k){ total.metrics[k]=1; });
    Object.keys(one.map||{}).forEach(function(ym){
      var box = total.map[ym] || (total.map[ym] = { label: one.map[ym].label, monthly: {} });
      copyAdd_(box.monthly, one.map[ym].monthly);
    });
  });
  return total;
}
function readWeeklyDynamicMulti_(sheetNames){
  var total = { map: {}, metrics: {} };
  (sheetNames || []).forEach(function(name){
    var one = readWeeklyDynamic_(name);
    Object.keys(one.metrics||{}).forEach(function(k){ total.metrics[k]=1; });
    Object.keys(one.map||{}).forEach(function(ym){
      var tMonth = total.map[ym] || (total.map[ym] = {});
      var sMonth = one.map[ym] || {};
      Object.keys(sMonth).forEach(function(ts){
        var src = sMonth[ts];
        var dst = tMonth[ts] || (tMonth[ts] = { label: src.label||'', ts: ts, weekly: {} });
        if (src.date) dst.date = src.date;
        copyAdd_(dst.weekly, src.weekly);
      });
    });
  });
  return total;
}

/******************** DYNAMIC READERS (MONTHLY / WEEKLY) ********************/
function readMonthlyDynamic_(sheetName){
  var sh = getSheetSafe_(sheetName); if (!sh) return {map:{}, metrics:{}};
  var values = sh.getDataRange().getValues(); if (!values.length) return {map:{}, metrics:{}};

  var head = values[0].map(function(h){ return String(h||'').trim(); });
  var norm = head.map(norm_);
  var rows = values.slice(1);

  var dateIdx   = firstIndex_(norm, ['month','period','date']);
  var metricIdx = detectMetricColumns_(rows, head, norm);

  var out = {}; var metrics = {};
  rows.forEach(function(r){
    var ym = ymFromAny_(r[dateIdx]); if (!ym) return;
    var box = out[ym] || (out[ym] = { label: monthLabelFromYm_(ym), monthly: {} });
    metricIdx.forEach(function(i){
      var name = head[i]; var val = toNumber_(r[i]); if (val == null) return;
      metrics[name]=1; box.monthly[name] = (box.monthly[name]||0) + val;
    });
  });
  return { map: out, metrics: metrics };
}

function readWeeklyDynamic_(sheetName){
  var sh = getSheetSafe_(sheetName); if (!sh) return {map:{}, metrics:{}};
  var values = sh.getDataRange().getValues(); if (!values.length) return {map:{}, metrics:{}};

  var head = values[0].map(function(h){ return String(h||'').trim(); });
  var norm = head.map(norm_);
  var rows = values.slice(1);

  var monthIdx = firstIndex_(norm, ['month','period','date']);
  var weekIdx  = findWeekDateColumn_(head, norm, rows);
  var metricIdx = detectMetricColumns_(rows, head, norm, true);

  var out = {}; var metrics = {}; var seqByYm = {};

  rows.forEach(function(r){
    var ym = ymFromAny_(r[monthIdx]); if (!ym) return;

    var d = dateFromAny_(r[weekIdx]);
    var ts;
    if (d) {
      d = new Date(d.getFullYear(), d.getMonth(), d.getDate());
      ts = Math.floor(d.getTime()/1000);
    } else {
      var seq=(seqByYm[ym]||0)+1; seqByYm[ym]=seq;
      ts = pseudoTs_(ym+':'+seq);
    }

    var mbox = out[ym] || (out[ym] = {});
    var w = mbox[ts] || (mbox[ts] = { label: '', ts: ts, weekly: {} });
    if (d) w.date = isoDate_(d);

    metricIdx.forEach(function(i){
      var name = head[i]; var val = toNumber_(r[i]); if (val == null) return;
      metrics[name]=1; w.weekly[name] = (w.weekly[name]||0) + val;
    });
  });

  return { map: out, metrics: metrics };
}

/** знайти колонку з датою старту тижня */
function findWeekDateColumn_(head, norm, rows){
  var prefer = ['week start','start_of_week','week date','start date','monday','start'];
  for (var i=0;i<norm.length;i++){
    for (var j=0;j<prefer.length;j++){
      if (norm[i].indexOf(prefer[j]) !== -1 && columnLooksLikeDate_(rows, i)) return i;
    }
  }
  var candidates = [];
  for (var k=0;k<norm.length;k++){
    if (norm[k]==='week') continue;
    if (columnLooksLikeDate_(rows, k)) candidates.push(k);
  }
  return candidates.length ? candidates[0] : -1;
}
function columnLooksLikeDate_(rows, idx){
  if (idx<0) return false;
  var hits=0, total=0;
  for (var i=0;i<rows.length;i++){
    var v = rows[i][idx]; if (v==='' || v==null) continue;
    total++; if (dateFromAny_(v)) hits++;
  }
  return total>0 && hits/Math.max(total,1) >= 0.5;
}

/******************** helpers ********************/
function getSheetSafe_(name){ try { return SpreadsheetApp.getActive().getSheetByName(name) || null; } catch(e){ return null; } }
function norm_(s){ return String(s||'').toLowerCase().trim(); }

function firstIndex_(normHeaders, needlesArr){
  var needles = needlesArr.map(norm_);
  for (var i=0;i<normHeaders.length;i++){
    for (var j=0;j<needles.length;j++){
      if (normHeaders[i].indexOf(needles[j]) !== -1) return i;
    }
  } return -1;
}

function detectMetricColumns_(rows, head, norm, excludeWeekCols){
  var idx = [];
  for (var i=0;i<head.length;i++){
    var h = norm[i];
    if (h.indexOf('month')!==-1 || h.indexOf('period')!==-1 || h.indexOf('date')!==-1) continue;
    if (excludeWeekCols && (h.indexOf('week')!==-1 || h.indexOf('start')!==-1)) continue;
    if (h==='label' || h==='id' || h==='ym' || h==='ts') continue;

    var isMetric=false;
    for (var r=0;r<rows.length;r++){
      var v = rows[r][i];
      if (isNumberLike_(v)) { isMetric=true; break; }
    }
    if (isMetric) idx.push(i);
  }
  return idx;
}

function isNumberLike_(v){ if (v==null || v==='') return false; if (typeof v==='number') return isFinite(v);
  var s=String(v).replace(/[€,%\s,]/g,'').trim(); return s!=='' && !isNaN(+s); }
function toNumber_(v){ if (!isNumberLike_(v)) return null; if (typeof v==='number') return v;
  var s=String(v).replace(/[€,%\s,]/g,'').trim(); return Number(s)||0; }

function ymFromAny_(v){
  var d=dateFromAny_(v); if (d) return d.getFullYear()+'-'+('0'+(d.getMonth()+1)).slice(-2);
  var s=String(v||'').trim(); var m=s.match(/^(\d{4})[-\/](\d{1,2})/);
  return m?(m[1]+'-'+('0'+m[2]).slice(-2)):null;
}

function dateFromAny_(v){
  if (Object.prototype.toString.call(v)==='[object Date]' && !isNaN(v)) return v;
  var s=String(v||'').trim();
  var m;
  m=s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/); if (m) return new Date(+m[3], +m[2]-1, +m[1]);
  m=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/); if (м = m){} // no-op
  m=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/); if (m) return new Date(+m[3], +m[2]-1, +m[1]);
  m=s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/); if (m) return new Date(+m[1], +m[2]-1, +m[3]);
  return null;
}

function isoDate_(d){ return d.getFullYear()+'-'+('0'+(d.getMonth()+1)).slice(-2)+'-'+('0'+d.getDate()).slice(-2); }
function monthLabelFromYm_(ym){
  var p=ym.split('-'); var y=+p[0], m=+p[1];
  var names=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return names[m-1]+' '+y;
}
function pseudoTs_(s){ var h=0; for (var i=0;i<s.length;i++) h=((h<<5)-h)+s.charCodeAt(i)|0; return 1700000000 + (Math.abs(h)%1000000); }

function mergeNameSets_(objA, objB){
  var args = Array.prototype.slice.call(arguments), out = {};
  args.forEach(function(o){ if (!o) return; Object.keys(o).forEach(function(k){ out[k]=1; }); });
  return out;
}
function emptyFromNames_(namesObj){ var o={}; Object.keys(namesObj||{}).forEach(function(k){ o[k]=0; }); return o; }
function copyAdd_(dst, src){ Object.keys(src||{}).forEach(function(k){ dst[k]=(dst[k]||0)+Number(src[k]||0); }); }
function hasNonZero_(months, metric){
  for (var i=0;i<(months||[]).length;i++){
    var m=months[i];
    if (Number(m.monthly && m.monthly[metric])>0) return true;
    if (Number(m.sumWeeks && m.sumWeeks[metric])>0) return true;
    for (var j=0;j<(m.weeks||[]).length;j++){
      var w=m.weeks[j]; if (Number(w.weekly && w.weekly[metric])>0) return true;
    }
  } return false;
}

/******************** Збереження PDF на Диск ********************/
function savePdfBase64(fileName, dataUrl) {
  var m = String(dataUrl||'').match(/^data:.*?;base64,(.*)$/);
  if (!m) throw new Error('Bad data URL');
  var bytes = Utilities.base64Decode(m[1]);
  var blob  = Utilities.newBlob(bytes, 'application/pdf', fileName || 'kpi.pdf');
  var file  = DriveApp.createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e){}
  return { id:file.getId(), name:file.getName(), url:file.getUrl() };
}

/** Альтернативний сейв за base64 (без data: префікса) + в папку */
function kpiSavePdf(base64, filename, folderId) {
  if (!base64) throw new Error('Empty PDF data');
  var raw = base64.indexOf(',') > -1 ? base64.split(',')[1] : base64;
  var blob = Utilities.newBlob(Utilities.base64Decode(raw), 'application/pdf', filename || 'export.pdf');
  var file = folderId ? DriveApp.getFolderById(folderId).createFile(blob) : DriveApp.createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return { id: file.getId(), url: file.getUrl(), name: file.getName() };
}

/******************** Локаль клієнту ********************/
function getLocale(){
  try { return Session.getActiveUserLocale() || 'uk'; } catch(e){ return 'uk'; }
}
