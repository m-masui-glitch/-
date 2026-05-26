// =====================================================================
// 回数券集計スクリプト — CSVファイル: 91cf8b6a
// 顧客数: 227名（参考）  列レイアウト: タイプC（氏名列=F列, オフセット+2）
//
// 【使い方】
// 1. スプレッドシートを開く → メニュー「拡張機能」→「Apps Script」
// 2. このコードを丸ごと貼り付けて保存（Ctrl+S）
// 3. ページ上部メニューに「回数券管理」が表示されるのでクリック
//    →「集計表を今すぐ更新」で集計が実行されます
// =====================================================================

// ===== 設定（このファイル専用） =====
const SOURCE_SHEET_NAME  = 'カルテ';  // ← カルテが入っているシートのタブ名
const SUMMARY_SHEET_NAME = '集計';    // ← 集計結果を書き出すシートのタブ名

// 列番号（0始まり：A=0, B=1, C=2, D=3, E=4, F=5, G=6 …）
const COL_NAME = 6;  // 氏名の列
const COL_T    = 14;  // 3回券完了マーカーの列
const COL_3K   = 15;  // 3回券セッション開始列（1/3 の列）
const COL_GRP  = 18;  // 最初の7列グループ開始列（チケット購入履歴）
// ====================================

function updateCouponSummary() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!src) {
    SpreadsheetApp.getUi().alert(
      '「' + SOURCE_SHEET_NAME + '」シートが見つかりません。\n' +
      'シート名（タブ名）を確認してください。'
    );
    return;
  }

  const data    = src.getDataRange().getValues();
  const results = [];
  const totals  = { '3': 0, '4': 0, '8': 0, '12': 0 };

  for (let r = 0; r < data.length; r++) {
    const row  = data[r];
    const name = String(row[COL_NAME] || '').trim();

    if (!name || _isLabelRow(name)) continue;

    const counts = { '3': 0, '4': 0, '8': 0, '12': 0 };

    // 3回券：3セッション列のいずれかに日付があれば1件
    const has3k = [COL_3K, COL_3K + 1, COL_3K + 2]
      .some(c => _isSessionDate(row[c]));
    if (has3k) {
      counts['3'] = 1;
      totals['3']++;
    }

    // 4回券・8回券・12回券：COL_GRP から右へ7列ずつスキャン
    // グループ構造：[マーカー][種別][購入日][施術1][施術2][施術3][施術4]
    //   〇 = 新規購入  ● = 継続（8・12回券後半）
    for (let c = COL_GRP; c + 3 < row.length; c += 7) {
      const marker = String(row[c] || '').trim();
      if (!marker && !row[c+1] && !row[c+2] && !row[c+3]) break;
      if (marker === '〇') {
        const type = _normalizeType(row[c + 1]);
        if (type) {
          counts[type]++;
          totals[type]++;
        }
      }
    }

    if (Object.values(counts).some(v => v > 0)) {
      results.push([name, counts['3'], counts['4'], counts['8'], counts['12']]);
    }
  }

  let dst = ss.getSheetByName(SUMMARY_SHEET_NAME);
  if (!dst) dst = ss.insertSheet(SUMMARY_SHEET_NAME);
  dst.clearContents();

  dst.getRange(1, 1, 1, 5).setValues(
    [['氏名', '3回券', '4回券', '8回券', '12回券']]
  );
  if (results.length > 0) {
    dst.getRange(2, 1, results.length, 5).setValues(results);
  }
  dst.getRange(results.length + 2, 1, 1, 5).setValues(
    [['【合計】', totals['3'], totals['4'], totals['8'], totals['12']]]
  );

  SpreadsheetApp.getUi().alert(
    '集計完了！ ' + results.length + '名を集計しました。'
  );
}

function _isLabelRow(name) {
  const LABELS = [
    '周期空き', '有効期限２か月前', '要注意（1か月前）', '期限切れ'
  ];
  return LABELS.includes(name) || /^Column\d/.test(name) || name.length > 40;
}

function _isSessionDate(val) {
  if (!val && val !== 0) return false;
  const s = String(val).trim();
  if (!s) return false;
  const SKIP = new Set(['●', '〇', '×', '✕', '△', '‐', '-',
                        '都度', 'FALSE', 'TRUE', 'NG', '機械のみ 当日のみ']);
  if (SKIP.has(s)) return false;
  return /\d/.test(s);
}

function _normalizeType(val) {
  if (!val) return null;
  const s = String(val).trim()
    .replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
    .replace(/[^\d]/g, '');
  return ['4', '8', '12'].includes(s) ? s : null;
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('回数券管理')
    .addItem('集計表を今すぐ更新', 'updateCouponSummary')
    .addItem('自動更新を設定する（1時間ごと）', 'setupTrigger')
    .addItem('自動更新を止める', 'removeTrigger')
    .addToUi();
}

function setupTrigger() {
  removeTrigger();
  ScriptApp.newTrigger('updateCouponSummary')
    .timeBased()
    .everyHours(1)
    .create();
  SpreadsheetApp.getUi().alert('1時間ごとの自動更新を設定しました。');
}

function removeTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'updateCouponSummary')
    .forEach(t => ScriptApp.deleteTrigger(t));
}
