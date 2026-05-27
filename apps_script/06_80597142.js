// =====================================================================
// 回数券集計スクリプト — スタンダード版
// CSVファイル: 80597142　顧客数（参考）: 1338名
//
// 【使い方】
// 1. スプレッドシートを開く → メニュー「拡張機能」→「Apps Script」
// 2. このコードを丸ごと貼り付けて保存（Ctrl+S）
// 3. 上部メニュー「回数券管理」→「集計表を今すぐ更新」をクリック
// =====================================================================

// ===== 設定 =====
const SOURCE_SHEET_NAME  = 'カルテ';  // ← シートのタブ名が違う場合は変更
const SUMMARY_SHEET_NAME = '集計';

// 列番号（0始まり：A=0, B=1, C=2...）　ない列は -1 に設定
const COL_CUSTOMER_ID = 1;   // 顧客番号（ない場合は -1）
const COL_NAME        = 4;   // 氏名
const COL_GENDER      = -1;  // 性別（ある場合は列番号を設定）
const COL_AGE         = -1;  // 年代（ある場合は列番号を設定）
const COL_CYCLE       = 3;   // 周期
const COL_EXPIRY      = 5;   // 有効期限
const COL_LAST_VISIT  = 6;   // 最終来店日
const COL_T           = 12;  // 3回券完了マーカーの列
const COL_3K          = 13;  // 3回券セッション開始列（1/3）
const COL_GRP         = 16;  // 最初の7列グループ開始列
// ================

function updateCouponSummary() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!src) {
    SpreadsheetApp.getUi().alert(
      '「' + SOURCE_SHEET_NAME + '」シートが見つかりません。\nシート名（タブ名）を確認してください。'
    );
    return;
  }

  const data      = src.getDataRange().getValues();
  const buyers12  = [];
  const buyers8   = [];
  const buyers4   = [];
  const only3k    = [];
  const firstOnly = [];

  for (let r = 0; r < data.length; r++) {
    const row  = data[r];
    const name = String(row[COL_NAME] || '').trim();
    if (!name || _isLabelRow(name)) continue;

    const tVal       = String(row[COL_T] || '').trim();
    const sessions3k = [COL_3K, COL_3K + 1, COL_3K + 2]
      .filter(c => _isSessionDate(row[c])).length;
    const ticket     = _getCurrentTicket(row);

    if (ticket) {
      const t3kStr = sessions3k > 0
        ? (sessions3k + '/3 ' + (tVal === '〇' ? '完了' : '使用中'))
        : (tVal === '〇' ? '完了' : '-');
      const entry = [
        ..._baseInfo(row),
        t3kStr,
        ticket.used,
        ticket.remaining,
        _fmtDate(ticket.purchaseDate)
      ];
      if      (ticket.type === '12') buyers12.push(entry);
      else if (ticket.type === '8')  buyers8.push(entry);
      else                           buyers4.push(entry);

    } else if (sessions3k > 0) {
      const status = (tVal === '〇') ? '完了・継続なし' : (sessions3k + '/3 使用中');
      only3k.push([..._baseInfo(row), sessions3k, status]);

    } else {
      firstOnly.push([..._baseInfoNoExpiry(row), tVal || '-']);
    }
  }

  let dst = ss.getSheetByName(SUMMARY_SHEET_NAME);
  if (!dst) dst = ss.insertSheet(SUMMARY_SHEET_NAME);
  dst.clearContents();

  const rows = [];
  rows.push(['【12回券 購入者】（' + buyers12.length + '名）']);
  rows.push(_buyerHeader('12'));
  buyers12.forEach(r => rows.push(r));
  rows.push([]);

  rows.push(['【8回券 購入者】（' + buyers8.length + '名）']);
  rows.push(_buyerHeader('8'));
  buyers8.forEach(r => rows.push(r));
  rows.push([]);

  rows.push(['【4回券 購入者】（' + buyers4.length + '名）']);
  rows.push(_buyerHeader('4'));
  buyers4.forEach(r => rows.push(r));
  rows.push([]);

  rows.push(['【3回券のみ（' + only3k.length + '名 - 使用中または完了・追加購入なし）】']);
  rows.push(_only3kHeader());
  only3k.forEach(r => rows.push(r));
  rows.push([]);

  rows.push(['【初回のみ・回数券未購入（' + firstOnly.length + '名）】']);
  rows.push(_firstOnlyHeader());
  firstOnly.forEach(r => rows.push(r));

  const maxCols = rows.reduce((m, r) => Math.max(m, r.length), 0);
  const padded  = rows.map(r => {
    const a = r.slice();
    while (a.length < maxCols) a.push('');
    return a;
  });
  if (padded.length > 0) {
    dst.getRange(1, 1, padded.length, maxCols).setValues(padded);
  }

  SpreadsheetApp.getUi().alert(
    '集計完了！\n' +
    '12回券：' + buyers12.length   + '名\n' +
    '8回券：'  + buyers8.length    + '名\n' +
    '4回券：'  + buyers4.length    + '名\n' +
    '3回券のみ：' + only3k.length  + '名\n' +
    '初回のみ：'  + firstOnly.length + '名'
  );
}

// ===== 内部関数（変更不要） =====

function _baseInfo(row) {
  const a = [];
  if (COL_CUSTOMER_ID >= 0) a.push(row[COL_CUSTOMER_ID] || '');
  a.push(String(row[COL_NAME] || '').trim());
  if (COL_GENDER >= 0)     a.push(row[COL_GENDER]    || '');
  if (COL_AGE >= 0)        a.push(row[COL_AGE]       || '');
  if (COL_CYCLE >= 0)      a.push(row[COL_CYCLE]     || '');
  if (COL_EXPIRY >= 0)     a.push(_fmtDate(row[COL_EXPIRY]));
  if (COL_LAST_VISIT >= 0) a.push(_fmtDate(row[COL_LAST_VISIT]));
  return a;
}

function _baseInfoNoExpiry(row) {
  const a = [];
  if (COL_CUSTOMER_ID >= 0) a.push(row[COL_CUSTOMER_ID] || '');
  a.push(String(row[COL_NAME] || '').trim());
  if (COL_GENDER >= 0)     a.push(row[COL_GENDER]    || '');
  if (COL_AGE >= 0)        a.push(row[COL_AGE]       || '');
  if (COL_CYCLE >= 0)      a.push(row[COL_CYCLE]     || '');
  if (COL_LAST_VISIT >= 0) a.push(_fmtDate(row[COL_LAST_VISIT]));
  return a;
}

function _buyerHeader(type) {
  const h = [];
  if (COL_CUSTOMER_ID >= 0) h.push('顧客番号');
  h.push('氏名');
  if (COL_GENDER >= 0) h.push('性別');
  if (COL_AGE >= 0)    h.push('年代');
  if (COL_CYCLE >= 0)  h.push('周期');
  if (COL_EXPIRY >= 0) h.push('有効期限');
  if (COL_LAST_VISIT >= 0) h.push('最終来店日');
  h.push('3回券');
  h.push(type + '回券(使用)');
  h.push(type + '回券(残り)');
  h.push('購入日');
  return h;
}

function _only3kHeader() {
  const h = [];
  if (COL_CUSTOMER_ID >= 0) h.push('顧客番号');
  h.push('氏名');
  if (COL_GENDER >= 0) h.push('性別');
  if (COL_AGE >= 0)    h.push('年代');
  if (COL_CYCLE >= 0)  h.push('周期');
  if (COL_EXPIRY >= 0) h.push('有効期限');
  if (COL_LAST_VISIT >= 0) h.push('最終来店日');
  h.push('3回券(使用)');
  h.push('状態');
  return h;
}

function _firstOnlyHeader() {
  const h = [];
  if (COL_CUSTOMER_ID >= 0) h.push('顧客番号');
  h.push('氏名');
  if (COL_GENDER >= 0) h.push('性別');
  if (COL_AGE >= 0)    h.push('年代');
  if (COL_CYCLE >= 0)  h.push('周期');
  if (COL_LAST_VISIT >= 0) h.push('最終来店日');
  h.push('更新状況');
  return h;
}

function _getCurrentTicket(row) {
  let lastIdx  = -1;
  let lastType = null;
  let lastDate = null;

  for (let c = COL_GRP; c + 3 < row.length; c += 7) {
    const marker = String(row[c] || '').trim();
    if (!marker && !row[c+1] && !row[c+2] && !row[c+3]) break;
    if (marker === '〇') {
      const type = _normalizeType(row[c + 1]);
      if (type) {
        lastIdx  = c;
        lastType = type;
        lastDate = row[c + 2];
      }
    }
  }
  if (lastIdx === -1) return null;

  const typeNum      = parseInt(lastType);
  const groupsNeeded = Math.ceil(typeNum / 4);
  let   used         = 0;

  for (let g = 0; g < groupsNeeded; g++) {
    const startCol = lastIdx + g * 7;
    if (startCol >= row.length) break;
    const marker = String(row[startCol] || '').trim();
    if (g > 0 && marker === '〇') break;
    for (let s = 3; s <= 6; s++) {
      if (startCol + s < row.length && _isSessionDate(row[startCol + s])) used++;
    }
  }

  return { type: lastType, used, remaining: typeNum - used, purchaseDate: lastDate };
}

function _fmtDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return val.getFullYear() + '/' + (val.getMonth() + 1) + '/' + val.getDate();
  }
  return String(val).trim();
}

function _isLabelRow(name) {
  const LABELS = ['周期空き', '有効期限２か月前', '要注意（1か月前）', '期限切れ'];
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
