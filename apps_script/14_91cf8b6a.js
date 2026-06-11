// =====================================================================
// 回数券集計スクリプト — スタンダード版
// CSVファイル: 91cf8b6a　顧客数（参考）: 227名
//
// 【使い方】
// 1. スプレッドシートを開く → メニュー「拡張機能」→「Apps Script」
// 2. このコードを丸ごと貼り付けて保存（Ctrl+S）
// 3. 上部メニュー「回数券管理」→「集計表を今すぐ更新」をクリック
// =====================================================================
// ===== 設定 =====
const SOURCE_SHEET_NAME  = '顧客管理';  // ← シートのタブ名が違う場合は変更
const SUMMARY_SHEET_NAME = '集計';

// 列番号（0始まり：A=0, B=1, C=2...）　ない列は -1 に設定
const COL_CUSTOMER_ID    = 1;   // B列：顧客番号
const COL_NAME           = 4;   // E列：氏名
const COL_GENDER         = -1;
const COL_AGE            = -1;
const COL_CYCLE          = 3;   // D列：周期
const COL_EXPIRY         = 5;   // F列：有効期限（自動入力）
const COL_LAST_VISIT     = 6;   // G列：最終来店日（自動入力）
const COL_FIRST_VISIT_3K = 11;  // L列：初回来店日（3回券の有効期限起算日）
const COL_T              = 12;  // M列：更新列（分類用・表には出力しない）
const COL_3K             = 13;  // N列：3回券セッション開始
const COL_GRP            = 16;  // Q列：7列グループ開始
// ================

// ===== F・G列の自動更新 =====

function updateDates() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!src) {
    SpreadsheetApp.getUi().alert('「' + SOURCE_SHEET_NAME + '」シートが見つかりません。');
    return;
  }
  _updateDatesInSheet(src);
  SpreadsheetApp.getUi().alert('F列（有効期限）・G列（最終来店日）を更新しました。');
}

function _updateDatesInSheet(src) {
  const range      = src.getDataRange();
  const numRows    = range.getNumRows();
  const data       = range.getValues();
  const fontColors = range.getFontColors();

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const expiryVals    = [];
  const expiryBgs     = [];
  const lastVisitVals = [];
  const lastVisitBgs  = [];

  for (let r = 0; r < numRows; r++) {
    const row  = data[r];
    const name = String(row[COL_NAME] || '').trim();

    if (!name || _isLabelRow(name)) {
      expiryVals.push([row[COL_EXPIRY] !== undefined ? row[COL_EXPIRY] : '']);
      expiryBgs.push([null]);
      lastVisitVals.push([row[COL_LAST_VISIT] !== undefined ? row[COL_LAST_VISIT] : '']);
      lastVisitBgs.push([null]);
      continue;
    }

    // F列：有効期限を計算
    const expiry = _calcExpiry(row, fontColors, r);
    if (expiry) {
      expiryVals.push([expiry]);
      // 有効期限から1ヶ月超えたら赤
      const threshold = new Date(expiry.getFullYear(), expiry.getMonth() + 1, expiry.getDate());
      expiryBgs.push([today > threshold ? '#ea4335' : null]);
    } else {
      expiryVals.push([row[COL_EXPIRY] !== undefined ? row[COL_EXPIRY] : '']);
      expiryBgs.push([null]);
    }

    // G列：最終来店日（黒文字の最新日付）
    const lastVisit = _calcLastVisit(row, fontColors, r);
    if (lastVisit) {
      lastVisitVals.push([lastVisit]);
      // 1ヶ月来店なしは黄色
      const oneMonthAgo = new Date(today.getFullYear(), today.getMonth() - 1, today.getDate());
      lastVisitBgs.push([lastVisit < oneMonthAgo ? '#fff2cc' : null]);
    } else {
      lastVisitVals.push([row[COL_LAST_VISIT] !== undefined ? row[COL_LAST_VISIT] : '']);
      lastVisitBgs.push([null]);
    }
  }

  const fRange = src.getRange(1, COL_EXPIRY + 1, numRows, 1);
  const gRange = src.getRange(1, COL_LAST_VISIT + 1, numRows, 1);
  fRange.setValues(expiryVals);
  fRange.setBackgrounds(expiryBgs);
  fRange.setNumberFormat('yyyy/mm/dd');
  gRange.setValues(lastVisitVals);
  gRange.setBackgrounds(lastVisitBgs);
  gRange.setNumberFormat('yyyy/mm/dd');
}

// 有効期限を計算
// - 4/8/12回券あり → 最新〇の購入日（col+2）+ 6ヶ月の月末
// - 3回券のみ     → L列（初回日）+ 3ヶ月の月末
function _calcExpiry(row, fontColors, r) {
  let lastPurchaseDate = null;

  for (let c = COL_GRP; c + 3 < row.length; c += 7) {
    const marker = String(row[c] || '').trim();
    if (!marker && !row[c+1] && !row[c+2] && !row[c+3]) break;
    if (marker === '〇') {
      const type = _normalizeType(row[c + 1]);
      if (type) {
        const d = _toDate(row[c + 2]);
        if (d) lastPurchaseDate = d;
      }
    }
  }

  if (lastPurchaseDate) return _monthEnd(lastPurchaseDate, 6);

  const firstVisit = _toDate(row[COL_FIRST_VISIT_3K]);
  if (firstVisit) return _monthEnd(firstVisit, 3);

  return null;
}

// 最終来店日を計算（黒文字の日付の中で最新）
function _calcLastVisit(row, fontColors, r) {
  let latest = null;

  const check = (val, c) => {
    if (_isSessionDate(val) && _isBlackText(_fc(fontColors, r, c))) {
      const d = _toDate(val);
      if (d && (!latest || d > latest)) latest = d;
    }
  };

  // 3回券エリア
  for (let c = COL_3K; c <= COL_3K + 2 && c < row.length; c++) check(row[c], c);

  // 7列グループのセッションスロット（col+3〜col+6）
  for (let c = COL_GRP; c + 3 < row.length; c += 7) {
    const marker = String(row[c] || '').trim();
    if (!marker && !row[c+1] && !row[c+2] && !row[c+3]) break;
    for (let s = 3; s <= 6; s++) {
      const col = c + s;
      if (col < row.length) check(row[col], col);
    }
  }

  return latest;
}

// N ヶ月後の月末日（例: 6月1日 + 3 → 8月31日、6月1日 + 6 → 11月30日）
// new Date(y, m, 0) は月 m の前月末日を返す
function _monthEnd(date, months) {
  return new Date(date.getFullYear(), date.getMonth() + months, 0);
}

// 値をDateオブジェクトに変換
function _toDate(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  const s = String(val).trim();
  if (!s) return null;
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d;
  const m = s.match(/(\d{4})[年\/\-](\d{1,2})[月\/\-](\d{1,2})/);
  if (m) {
    const d2 = new Date(+m[1], +m[2] - 1, +m[3]);
    return isNaN(d2.getTime()) ? null : d2;
  }
  return null;
}

// ===== 集計表の更新 =====

function updateCouponSummary() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!src) {
    SpreadsheetApp.getUi().alert(
      '「' + SOURCE_SHEET_NAME + '」シートが見つかりません。\nシート名（タブ名）を確認してください。'
    );
    return;
  }

  // F・G列を先に自動更新してから集計
  _updateDatesInSheet(src);

  // 更新後のデータを再読み込み
  const range       = src.getDataRange();
  const data        = range.getValues();
  const fontColors  = range.getFontColors();
  const backgrounds = range.getBackgrounds();

  const act12 = [], act8 = [], act4 = [], act3k = [];
  const chu12 = [], chu8 = [], chu4 = [], chu3k = [];
  const expired = [];

  for (let r = 0; r < data.length; r++) {
    const row  = data[r];
    const name = String(row[COL_NAME] || '').trim();
    if (!name || _isLabelRow(name)) continue;

    const tVal      = String(row[COL_T] || '').trim();
    const isExpired = _checkExpired(row, r, backgrounds);
    const isChurned = (tVal === '×' || tVal === '✕');

    const sessions3k = [COL_3K, COL_3K + 1, COL_3K + 2].filter(c =>
      _isSessionDate(row[c]) && _isBlackText(_fc(fontColors, r, c))
    ).length;

    const has3kYellow = sessions3k === 0 && [COL_3K, COL_3K + 1, COL_3K + 2].some(c =>
      _isYellow(_bg(backgrounds, r, c))
    );

    const ticket = _getCurrentTicket(row, r, fontColors, backgrounds);
    const base   = _baseInfo(row);

    if (ticket) {
      const t3k = has3kYellow ? '0/3 未使用'
                : sessions3k === 3 ? '3/3 完了'
                : sessions3k > 0   ? sessions3k + '/3 使用中'
                : '-';
      const entry = [...base, t3k, ticket.used, ticket.remaining];

      if      (isExpired)  expired.push([...base, ticket.type + '回券']);
      else if (isChurned)  _push(chu12, chu8, chu4, ticket.type, entry);
      else                 _push(act12, act8, act4, ticket.type, entry);

    } else if (sessions3k > 0 || has3kYellow) {
      const used3k      = has3kYellow ? 0 : sessions3k;
      const remaining3k = 3 - used3k;
      const entry = [...base, used3k, remaining3k];

      if      (isExpired)  expired.push([...base, '3回券']);
      else if (isChurned)  chu3k.push(entry);
      else                 act3k.push(entry);

    } else {
      if (isExpired) expired.push([...base, '初回のみ']);
    }
  }

  let dst = ss.getSheetByName(SUMMARY_SHEET_NAME);
  if (!dst) dst = ss.insertSheet(SUMMARY_SHEET_NAME);
  dst.clearContents();

  const out = [];
  const nAct = act12.length + act8.length + act4.length + act3k.length;
  const nChu = chu12.length + chu8.length + chu4.length + chu3k.length;

  out.push(['▼ 保有顧客（' + nAct + '名）']);
  out.push([]);
  _appendSection(out, '12回券 購入者', act12, _buyerHeader('12'));
  _appendSection(out, '8回券 購入者',  act8,  _buyerHeader('8'));
  _appendSection(out, '4回券 購入者',  act4,  _buyerHeader('4'));
  _appendSection(out, '3回券 購入者',  act3k, _only3kHeader());
  out.push([]);

  out.push(['▼ 離客（' + nChu + '名）']);
  out.push([]);
  _appendSection(out, '12回券（離客）', chu12, _buyerHeader('12'));
  _appendSection(out, '8回券（離客）',  chu8,  _buyerHeader('8'));
  _appendSection(out, '4回券（離客）',  chu4,  _buyerHeader('4'));
  _appendSection(out, '3回券（離客）',  chu3k, _only3kHeader());
  out.push([]);

  out.push(['▼ 有効期限なし・期限切れ（' + expired.length + '名）']);
  out.push(_expiredHeader());
  expired.forEach(r => out.push(r));

  const maxCols = out.reduce((m, r) => Math.max(m, r.length), 0);
  if (maxCols > 0) {
    const padded = out.map(r => {
      const a = r.slice();
      while (a.length < maxCols) a.push('');
      return a;
    });
    dst.getRange(1, 1, padded.length, maxCols).setValues(padded);
  }

  SpreadsheetApp.getUi().alert(
    '集計完了！\n' +
    '保有顧客 ' + nAct + '名（12回券:' + act12.length + ' 8回券:' + act8.length +
    ' 4回券:' + act4.length + ' 3回券:' + act3k.length + '）\n' +
    '離客 ' + nChu + '名\n' +
    '期限切れ ' + expired.length + '名'
  );
}

// ===== 内部関数（変更不要） =====

function _push(list12, list8, list4, type, entry) {
  if      (type === '12') list12.push(entry);
  else if (type === '8')  list8.push(entry);
  else                    list4.push(entry);
}

function _appendSection(out, title, data, header) {
  out.push(['【' + title + '】（' + data.length + '名）']);
  if (data.length > 0) {
    out.push(header);
    data.forEach(r => out.push(r));
  }
  out.push([]);
}

function _checkExpired(row, r, backgrounds) {
  if (COL_EXPIRY < 0) return false;
  const val = row[COL_EXPIRY];
  if (!val || String(val).trim() === '') return true;
  if (_isRed(_bg(backgrounds, r, COL_EXPIRY))) return true;
  // 有効期限日が今日より前なら期限切れ
  const d = _toDate(val);
  if (d) {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    return d < today;
  }
  return false;
}

// 現在のチケット（最新購入分）を返す
// ─ 7列グループ構造: [〇/●マーカー][回数券種別][購入日][s1][s2][s3][s4]
// ─ 回数券の種類は〇グループの青3列の真ん中（col+1）から読み取る
// ─ 途中で種類が変わっても「最後の〇」＝最新購入を採用する
function _getCurrentTicket(row, r, fontColors, backgrounds) {
  let lastIdx = -1, lastType = null;

  for (let c = COL_GRP; c + 3 < row.length; c += 7) {
    const marker = String(row[c] || '').trim();
    if (!marker && !row[c+1] && !row[c+2] && !row[c+3]) break;
    if (marker === '〇') {
      const type = _normalizeType(row[c + 1]);
      if (type) { lastIdx = c; lastType = type; }
    }
  }
  if (lastIdx === -1) return null;

  const typeNum      = parseInt(lastType);
  const groupsNeeded = Math.ceil(typeNum / 4);

  const firstSessCol = lastIdx + 3;
  if (!_isSessionDate(row[firstSessCol]) && _isYellow(_bg(backgrounds, r, firstSessCol))) {
    return { type: lastType, used: 0, remaining: typeNum };
  }

  let used = 0;
  for (let g = 0; g < groupsNeeded; g++) {
    const startCol = lastIdx + g * 7;
    if (startCol >= row.length) break;
    if (g > 0 && String(row[startCol] || '').trim() === '〇') break;
    for (let s = 3; s <= 6; s++) {
      const col = startCol + s;
      if (col < row.length && _isSessionDate(row[col]) && _isBlackText(_fc(fontColors, r, col))) {
        used++;
      }
    }
  }

  return { type: lastType, used, remaining: typeNum - used };
}

// セルの値・色を安全に取得
function _fc(fontColors, r, c) {
  return (r < fontColors.length && c < fontColors[r].length) ? fontColors[r][c] : '';
}
function _bg(backgrounds, r, c) {
  return (r < backgrounds.length && c < backgrounds[r].length) ? backgrounds[r][c] : '';
}

// 色判定
function _hexToRgb(hex) {
  if (!hex || hex.length < 7) return null;
  return {
    r: parseInt(hex.slice(1, 3), 16),
    g: parseInt(hex.slice(3, 5), 16),
    b: parseInt(hex.slice(5, 7), 16)
  };
}
function _isBlackText(fc) {
  const c = (fc || '').toLowerCase();
  return !c || c === '#000000' || c === '#1f1f1f' || c === '#202124';
}
function _isYellow(bg) {
  const rgb = _hexToRgb(bg);
  return rgb ? (rgb.r > 200 && rgb.g > 170 && rgb.b < 120) : false;
}
function _isRed(bg) {
  const rgb = _hexToRgb(bg);
  return rgb ? (rgb.r > 160 && rgb.g < 120 && rgb.b < 120) : false;
}

// 行データ構築
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

// ヘッダー構築
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
  h.push('3回券(残り)');
  return h;
}
function _expiredHeader() {
  const h = [];
  if (COL_CUSTOMER_ID >= 0) h.push('顧客番号');
  h.push('氏名');
  if (COL_GENDER >= 0) h.push('性別');
  if (COL_AGE >= 0)    h.push('年代');
  if (COL_CYCLE >= 0)  h.push('周期');
  if (COL_EXPIRY >= 0) h.push('有効期限');
  if (COL_LAST_VISIT >= 0) h.push('最終来店日');
  h.push('最後のチケット');
  return h;
}

// ユーティリティ
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

// ===== メニュー・トリガー =====
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('回数券管理')
    .addItem('集計表を今すぐ更新', 'updateCouponSummary')
    .addItem('F・G列の日付のみ更新する', 'updateDates')
    .addItem('自動更新を設定する（1時間ごと）', 'setupTrigger')
    .addItem('自動更新を止める', 'removeTrigger')
    .addToUi();
}
function setupTrigger() {
  removeTrigger();
  ScriptApp.newTrigger('updateCouponSummary').timeBased().everyHours(1).create();
  SpreadsheetApp.getUi().alert('1時間ごとの自動更新を設定しました。');
}
function removeTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'updateCouponSummary')
    .forEach(t => ScriptApp.deleteTrigger(t));
}
