/**
 * ビイサイド（福祉事業部）実績ダッシュボード — 広告コスト・求人別データの自動更新 Apps Script
 *  convert.py と同じ data.json スキーマ（{meta, adCost, jobs}）を生成し、
 *  GitHub（lc-rm/b-side-dashboard）の data.json に5分ごと自動push。
 *  ★氏名・電話・年齢などのPIIは一切出力しない（広告指標と求人掲載実績のみ＝非PII）。
 *
 * 【初回セットアップ（対象スプシを開ける会社アカウント r_murai で実行）】
 *  A. GitHubで fine-grained トークン（Resource owner=lc-rm / Repository=lc-rm/b-side-dashboard /
 *     Permissions: Contents = Read and write）を作成
 *  B. このコードを script.google.com のプロジェクトに丸ごと貼り付け
 *  C. 「プロジェクトの設定（歯車）> スクリプト プロパティ」で GITHUB_TOKEN = （作ったトークン）を追加
 *  D. 関数 setupTrigger を1回だけ実行（▶）→ 5分ごとの自動更新が有効に
 */

var SHEET_ID      = '1yoQ38egYzZD-hg2uDiIEmzxLOyk9xPcMrlaAYUh-gxE';      // 実績元データ（Indeed結果 ＝広告コスト）
var SHEET_ID_JOBS = '1BqN8GKY4rv07pNzDBZI6_MpwqpXVrD_Ro1vBkzQH_5M';      // 求人別データ（福祉事業部）
var GITHUB_REPO   = 'lc-rm/b-side-dashboard';
var GITHUB_PATH   = 'data.json';

var COST_COLS = ['imp', 'click', 'CTR', '応募開始数', 'CV', '応募完了率', 'CVR', 'CPC', 'CPA', 'COST', '予算残'];

// ===== データ集計（非PII）=====
function buildData() {
  var out = { meta: { source: 'ビイサイド（福祉事業部）Indeed広告 / 求人別', generatedFrom: 'AppsScript', fetchedAt: new Date().toISOString() }, adCost: [], jobs: [] };

  // --- 1) 広告コスト：indeedﾃﾞｰﾀ(YYYY年度) シート ---
  var ssI = SpreadsheetApp.openById(SHEET_ID);
  var seen = {};
  ssI.getSheets().forEach(function (sh) {
    if (sh.getName().indexOf('indeedﾃﾞｰﾀ') !== 0) return;
    var rows = sh.getDataRange().getValues();
    for (var i = 0; i < rows.length;) {
      var r = rows[i];
      if (r && t(r[0]) === 'indeed応募' && t(r[1]) === '予算') {
        var idx = {};
        for (var ci = 0; ci < r.length; ci++) { if (COST_COLS.indexOf(r[ci]) >= 0) idx[r[ci]] = ci; }
        var tot = rows[i + 1], ym = '', camps = [], j = i + 2;
        while (j < rows.length && !(rows[j] && t(rows[j][0]) === 'indeed応募')) {
          var cr = rows[j], name = cr ? t(cr[0]) : '';
          if (name && name !== '合計') {
            var p = periodYm(cr[2]); if (p && !ym) ym = p;
            camps.push({
              type: name, period: t(cr[2]), budget: n(cr[1]), cost: gv(cr, idx, 'COST'),
              imp: Math.round(gv(cr, idx, 'imp')), click: Math.round(gv(cr, idx, 'click')), ctr: rv(cr, idx, 'CTR'),
              applyStart: Math.round(gv(cr, idx, '応募開始数')), cv: Math.round(gv(cr, idx, 'CV')), cvr: rv(cr, idx, 'CVR'),
              cpc: rv(cr, idx, 'CPC'), cpa: rv(cr, idx, 'CPA')
            });
          }
          j++;
        }
        if (tot && t(tot[0]) === '合計' && ym && !seen[ym]) {
          seen[ym] = 1;
          out.adCost.push({
            ym: ym, budget: n(tot[1]), cost: gv(tot, idx, 'COST'),
            imp: Math.round(gv(tot, idx, 'imp')), click: Math.round(gv(tot, idx, 'click')), ctr: rv(tot, idx, 'CTR'),
            applyStart: Math.round(gv(tot, idx, '応募開始数')), cv: Math.round(gv(tot, idx, 'CV')), cvr: rv(tot, idx, 'CVR'),
            cpc: rv(tot, idx, 'CPC'), cpa: rv(tot, idx, 'CPA'),
            budgetLeft: ('予算残' in idx) ? n(tot[idx['予算残']]) : null, campaigns: camps
          });
        }
        i = j;
      } else i++;
    }
  });
  out.adCost.sort(function (a, b) { return a.ym < b.ym ? -1 : 1; });

  // --- 2) 求人別：月別シート「YYYY.M」。(ym, 求人名) で合算。応募>0 の求人のみ。 ---
  var ssJ = SpreadsheetApp.openById(SHEET_ID_JOBS);
  var jmap = {};
  ssJ.getSheets().forEach(function (sh) {
    var m = sh.getName().match(/^(\d{4})\.(\d{1,2})$/); if (!m) return;
    var ym = m[1] + ('0' + m[2]).slice(-2);
    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i]; if (!r) continue;
      var title = t(r[3]); if (!title) continue;
      var cp = cpNorm(r[1]), loc = t(r[5]).split(',')[0];
      var key = ym + '' + title;
      var a = jmap[key];
      if (!a) { a = { ym: ym, title: title, cp: cp, loc: loc, imp: 0, click: 0, applyStart: 0, applies: 0, cost: 0 }; jmap[key] = a; }
      a.imp += Math.round(n(r[13])); a.click += Math.round(n(r[14])); a.applyStart += Math.round(n(r[15]));
      a.applies += Math.round(n(r[16])); a.cost += Math.round(n(r[14]) * n(r[21]));
      if (!a.cp && cp) a.cp = cp;
    }
  });
  Object.keys(jmap).forEach(function (k) {
    var a = jmap[k]; if (a.applies <= 0) return;
    a.cpc = a.click ? Math.round(a.cost / a.click) : 0;
    a.cpa = a.applies ? Math.round(a.cost / a.applies) : 0;
    out.jobs.push(a);
  });
  out.jobs.sort(function (a, b) { return a.ym !== b.ym ? (a.ym < b.ym ? -1 : 1) : (b.applies - a.applies); });

  out.meta.adCostMonths = out.adCost.length;
  out.meta.jobRows = out.jobs.length;
  out.meta.latestAdYm = out.adCost.length ? out.adCost[out.adCost.length - 1].ym : '';
  return out;
}

// ===== GitHubへ自動push（5分ごと）=====
function autoPush() {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('GITHUB_TOKEN');
  if (!token) { Logger.log('GITHUB_TOKEN が未設定です'); return; }
  var json = JSON.stringify(buildData());
  var hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, json));
  if (props.getProperty('LAST_HASH') === hash) { Logger.log('変更なし → スキップ'); return; }

  var base = 'https://api.github.com/repos/' + GITHUB_REPO + '/contents/' + GITHUB_PATH;
  var headers = { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28' };
  var sha = null;
  var g = UrlFetchApp.fetch(base + '?ref=main', { headers: headers, muteHttpExceptions: true });
  if (g.getResponseCode() === 200) sha = JSON.parse(g.getContentText()).sha;

  var payload = { message: 'auto: update data.json ' + new Date().toISOString(),
    content: Utilities.base64Encode(json, Utilities.Charset.UTF_8), branch: 'main' };
  if (sha) payload.sha = sha;
  var p = UrlFetchApp.fetch(base, { method: 'put', headers: headers, contentType: 'application/json',
    payload: JSON.stringify(payload), muteHttpExceptions: true });
  var code = p.getResponseCode();
  if (code === 200 || code === 201) { props.setProperty('LAST_HASH', hash); Logger.log('更新成功 ' + code); }
  else Logger.log('失敗 ' + code + ' ' + p.getContentText().slice(0, 300));
}

// 5分ごとの自動更新を設定（1回だけ実行）
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (tr) {
    if (tr.getHandlerFunction() === 'autoPush') ScriptApp.deleteTrigger(tr);
  });
  ScriptApp.newTrigger('autoPush').timeBased().everyMinutes(5).create();
  autoPush();
  Logger.log('5分ごとの自動更新を設定しました');
}

// ===== helpers =====
function t(x) { return (x == null ? '' : String(x)).trim(); }
function n(x) { if (x === '' || x == null) return 0; var v = (typeof x === 'number') ? x : parseFloat(String(x).replace(/[^0-9.\-]/g, '')); return isNaN(v) ? 0 : v; }
function rate(x) { if (x === '' || x == null) return null; var v = (typeof x === 'number') ? x : parseFloat(String(x).replace(/[^0-9.\-]/g, '')); return isNaN(v) ? null : v; }
function gv(row, idx, k) { return (k in idx) ? n(row[idx[k]]) : 0; }
function rv(row, idx, k) { return (k in idx) ? rate(row[idx[k]]) : null; }
function cpNorm(v) { return t(v).replace(/^\d{4}年\d{1,2}月[:：]\s*/, ''); }
function periodYm(s) { var m = String(s == null ? '' : s).match(/(\d{4})\/(\d{1,2})/); return m ? m[1] + ('0' + m[2]).slice(-2) : ''; }
