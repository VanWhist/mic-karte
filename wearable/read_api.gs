/**
 * read_api.gs — コンディション可視化アプリ 閲覧用API（読み取り専用）
 *
 * 【設置場所】
 *   既存のApps Scriptプロジェクト「ウェアラブル取り込み用」
 *   （スクリプトID: 1FcftU72wx3PO4HcIxm0He1Z-VHKEU_Sb_DMeDriC0TfB9oz1nk4ANr37）
 *   に、このファイルを新規ファイルとして追加する。
 *
 * 【重要：新しいプロジェクトを作らないこと】
 *   Apps Scriptはプロジェクト内の全ファイルがグローバル名前空間を共有する。
 *   別ファイルに doGet / doPost を作ると後から読まれた定義で上書きされ、
 *   エラーが出ないまま既存の取り込みアプリが壊れる。
 *   → このファイルには doGet / doPost を定義していない。
 *      既存の doGet に「差し込み手順」の1行を足して使う。
 *
 * 【安全のための命名】
 *   このファイルが定義するグローバル名はすべて view で始まる。
 *   既存コードに同名のものが無いことだけ確認すればよい。
 *     VIEW_CFG / viewSummary_ / viewJson_ / viewYmd_ / viewCell_ /
 *     viewBaselineFor_ / viewSafeEquals_ / viewSelfTest
 *
 * 【書き込みは一切しない】
 *   このファイルはシートを読むだけ。appendRow / setValue の類は使わない。
 */


/* ==========================================================================
   差し込み手順：既存の doGet の先頭に、次の2行を足すだけ
   ==========================================================================

   function doGet(e) {
     var p = (e && e.parameter) || {};                       // ← 既存に無ければ足す
     if (p.action === 'summary') return viewSummary_(p);     // ← この1行を足す

     // ── ここから下は既存のコードをそのまま残す ──
     ...
   }

   既存の doGet が e.parameter を別の変数名で受けている場合は、
   その変数を使って `if (<既存の変数>.action === 'summary') return viewSummary_(<既存の変数>);`
   と書き換えるだけでよい。既存の分岐は消さないこと。

   ========================================================================== */


/* --------------------------------------------------------------------------
   設定
   -------------------------------------------------------------------------- */
var VIEW_CFG = {
  SHEET_NAME:    'wearable',      // 読むのはこのシートだけ
  TZ:            'Asia/Tokyo',
  DEFAULT_DAYS:  60,              // days 未指定のときに返す日数
  MAX_DAYS:      400,

  // スクリプトプロパティのキー名。値そのものはコードに書かない。
  // 取り込み用の INGEST_KEY は流用しない（用途と権限が違うため）。
  TOKEN_PROP:    'VIEW_TOKEN',
  BASELINE_PROP: 'VIEW_BASELINE_JSON'
};

// フロントに返す列。wearable シートのヘッダー名と一致させる。
var VIEW_COLUMNS = [
  'date', 'athlete_id',
  'sleep_start', 'sleep_end',
  'sleep_hours', 'in_bed_hours', 'sleep_efficiency',
  'deep_hours', 'rem_hours', 'core_hours', 'awake_hours',
  'resting_hr', 'hr_min', 'hr_avg', 'hrv_ms',
  'resp_rate', 'spo2', 'wrist_temp',
  'steps', 'source', 'updated_at'
];

// 日付として扱う列（Date型で入っていたら yyyy-MM-dd の文字列に戻す）
var VIEW_DATE_COLUMNS = { date: 1 };
// 日時として扱う列（Date型で入っていたら yyyy-MM-dd HH:mm:ss の文字列に戻す）
var VIEW_DATETIME_COLUMNS = { sleep_start: 1, sleep_end: 1, updated_at: 1 };


/* --------------------------------------------------------------------------
   本体
   -------------------------------------------------------------------------- */
function viewSummary_(p) {
  try {
    // --- 閲覧用トークンの照合（読み取り専用） -----------------------------
    var expected = PropertiesService.getScriptProperties().getProperty(VIEW_CFG.TOKEN_PROP);
    if (!expected) {
      return viewJson_({ ok: false, error: '閲覧用トークンが未設定です（スクリプトプロパティを確認してください）' });
    }
    if (!viewSafeEquals_(String(p.token || ''), String(expected))) {
      // 何が違うかは返さない
      return viewJson_({ ok: false, error: 'アクセストークンが正しくありません' });
    }

    var athleteId = String(p.athlete_id || '').trim();
    if (!athleteId) {
      return viewJson_({ ok: false, error: 'athlete_id が指定されていません' });
    }

    var days = parseInt(p.days, 10);
    if (!isFinite(days) || days <= 0) days = VIEW_CFG.DEFAULT_DAYS;
    days = Math.min(days, VIEW_CFG.MAX_DAYS);

    // --- シート読み込み ---------------------------------------------------
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VIEW_CFG.SHEET_NAME);
    if (!sh) {
      return viewJson_({ ok: false, error: 'シート「' + VIEW_CFG.SHEET_NAME + '」が見つかりません' });
    }
    var values = sh.getDataRange().getValues();
    if (values.length < 2) {
      return viewJson_(viewEmptyPayload_(athleteId));
    }

    var header = values.shift().map(function (h) { return String(h).trim(); });
    var idx = {};
    header.forEach(function (h, i) { idx[h] = i; });

    if (!(('date') in idx) || !(('athlete_id') in idx)) {
      return viewJson_({ ok: false, error: 'wearable シートに date / athlete_id 列が見つかりません' });
    }

    // --- 対象期間（日付はすべて文字列のまま比較する） ---------------------
    // ※Date型に変換すると突き合わせが壊れる（このプロジェクトの既知バグ類型）
    var today  = Utilities.formatDate(new Date(), VIEW_CFG.TZ, 'yyyy-MM-dd');
    var cutoff = Utilities.formatDate(
      new Date(new Date().getTime() - (days - 1) * 86400000), VIEW_CFG.TZ, 'yyyy-MM-dd');

    var rows = [];
    for (var r = 0; r < values.length; r++) {
      var row = values[r];
      if (String(row[idx.athlete_id]).trim() !== athleteId) continue;

      var ymd = viewYmd_(row[idx.date]);
      if (!ymd) continue;                 // 日付が読めない行は捨てる（0扱いしない）
      if (ymd < cutoff || ymd > today) continue;

      var obj = {};
      for (var c = 0; c < VIEW_COLUMNS.length; c++) {
        var name = VIEW_COLUMNS[c];
        obj[name] = (name in idx) ? viewCell_(name, row[idx[name]]) : '';
      }
      obj.date = ymd;
      rows.push(obj);
    }

    // 日付の昇順（文字列比較でよい形式なのでそのまま比較する）
    rows.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });

    var bl = viewBaselineFor_(athleteId);

    return viewJson_({
      ok: true,
      athlete_id: athleteId,
      athlete_name: (bl && bl.name) ? bl.name : athleteId,
      baseline: bl ? {
        start:       bl.start || '',
        end:         bl.end || '',
        approved_by: bl.approved_by || '',
        approved_at: bl.approved_at || ''
      } : null,
      days: rows,
      generated_at: Utilities.formatDate(new Date(), VIEW_CFG.TZ, 'yyyy-MM-dd HH:mm:ss'),
      range: { from: cutoff, to: today }
    });

  } catch (err) {
    // 例外の中身に設定値が混ざる可能性があるため、種別だけ返す
    return viewJson_({ ok: false, error: 'サーバー側でエラーが発生しました（' + err.name + '）' });
  }
}


/* --------------------------------------------------------------------------
   補助
   -------------------------------------------------------------------------- */
function viewEmptyPayload_(athleteId) {
  var bl = viewBaselineFor_(athleteId);
  return {
    ok: true,
    athlete_id: athleteId,
    athlete_name: (bl && bl.name) ? bl.name : athleteId,
    baseline: bl ? {
      start: bl.start || '', end: bl.end || '',
      approved_by: bl.approved_by || '', approved_at: bl.approved_at || ''
    } : null,
    days: [],
    generated_at: Utilities.formatDate(new Date(), VIEW_CFG.TZ, 'yyyy-MM-dd HH:mm:ss')
  };
}

/**
 * 個人基準（ベースライン）の期間設定を返す。
 *
 * ★自動更新はしない★
 *   基準を毎日自動で更新すると、徐々に悪化した状態まで「新しい平常」として
 *   吸収してしまい、悪化が検出できなくなる。復帰期の選手では特に危険。
 *   そのため、ここは固定値を読むだけにしてある。
 *   期間を変えるのは、PTが承認したときに人がプロパティを書き換えるときだけ。
 *
 * スクリプトプロパティ VIEW_BASELINE_JSON の形式：
 *   {
 *     "van-test": {
 *       "name": "van（テスト）",
 *       "start": "2026-07-28",
 *       "end": "2026-08-03",
 *       "approved_by": "担当PT",
 *       "approved_at": "2026-08-04"
 *     }
 *   }
 */
function viewBaselineFor_(athleteId) {
  var raw = PropertiesService.getScriptProperties().getProperty(VIEW_CFG.BASELINE_PROP);
  if (!raw) return null;
  var map;
  try {
    map = JSON.parse(raw);
  } catch (e) {
    return null;   // 壊れていたら「基準未設定」として扱う（画面側は「データ不足」を出す）
  }
  return map[athleteId] || null;
}

/** セル値を、その列の型に合わせた文字列／数値に整える。空欄は空文字のまま返す（0にしない）。 */
function viewCell_(name, v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) {
    if (VIEW_DATE_COLUMNS[name])     return Utilities.formatDate(v, VIEW_CFG.TZ, 'yyyy-MM-dd');
    if (VIEW_DATETIME_COLUMNS[name]) return Utilities.formatDate(v, VIEW_CFG.TZ, 'yyyy-MM-dd HH:mm:ss');
    return Utilities.formatDate(v, VIEW_CFG.TZ, 'yyyy-MM-dd HH:mm:ss');
  }
  if (typeof v === 'number') return v;
  return String(v).trim();
}

/**
 * date 列の値を yyyy-MM-dd の文字列にして返す。読めなければ '' を返す。
 * シート側で日付が文字列として入っていても、Sheetsに勝手にDate化されていても、
 * どちらでも同じ文字列になるようにしている。
 */
function viewYmd_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, VIEW_CFG.TZ, 'yyyy-MM-dd');
  var s = String(v == null ? '' : v).trim();
  if (!s) return '';
  var m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (!m) return '';
  return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
}

/** 長さの違いで早期に抜けない比較。トークン照合に使う。 */
function viewSafeEquals_(a, b) {
  if (a.length !== b.length) return false;
  var diff = 0;
  for (var i = 0; i < a.length; i++) diff |= (a.charCodeAt(i) ^ b.charCodeAt(i));
  return diff === 0;
}

function viewJson_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


/* --------------------------------------------------------------------------
   動作確認用（エディタから実行してログを見る）
   ブラウザからは呼べない。トークンの値はログにも出さない。
   -------------------------------------------------------------------------- */
function viewSelfTest() {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty(VIEW_CFG.TOKEN_PROP);
  Logger.log('VIEW_TOKEN 設定: %s', token ? 'あり' : 'なし ← 未設定です');
  Logger.log('VIEW_BASELINE_JSON 設定: %s', props.getProperty(VIEW_CFG.BASELINE_PROP) ? 'あり' : 'なし ← 未設定です');

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VIEW_CFG.SHEET_NAME);
  Logger.log('シート「%s」: %s', VIEW_CFG.SHEET_NAME, sh ? ('あり（' + sh.getLastRow() + '行）') : 'なし ← 名前を確認してください');

  if (!token) return;
  var res = viewSummary_({ action: 'summary', token: token, athlete_id: 'van-test', days: 60 });
  var json = JSON.parse(res.getContent());
  Logger.log('ok: %s', json.ok);
  Logger.log('error: %s', json.error || '(なし)');
  Logger.log('取得件数: %s 日分', json.days ? json.days.length : 0);
  Logger.log('基準期間: %s', json.baseline ? (json.baseline.start + ' 〜 ' + json.baseline.end) : '(未設定)');
  if (json.days && json.days.length) {
    Logger.log('最新行: %s', JSON.stringify(json.days[json.days.length - 1]));
    // date が文字列のままか確認する（Date型に化けていないこと）
    Logger.log('date の型: %s', typeof json.days[json.days.length - 1].date);
  }
}
