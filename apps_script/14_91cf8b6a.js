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

  const range       = src.getDataRange();
  const data        = range.getValues();
  const fontColors  = range.getFontColors();
  const backgrounds = range.getBackgrounds();

  // 保有顧客（T=〇 or △、有効期限あり）
  const act12 = [], act8 = [], act4 = [], act3k = [], actFirst = [];
  // 離客（T=×）
  const chu12 = [], chu8 = [], chu4 = [], chu3k = [], chuFirst = [];
  // 期限切れ（有効期限なし or 赤セル）→ 最下部
  const expired = [];

  for (let r = 0; r < data.length; r++) {
    const row  = data[r];
    const name = String(row[COL_NAME] || '').trim();
    if (!name || _isLabelRow(name)) continue;

    const tVal      = String(row[COL_T] || '').trim();
    const isExpired = _checkExpired(row, r, backgrounds);
    const isChurned = (tVal === '×' || tVal === '✕');

    // 3回券：黒文字の日付のみカウント
    const sessions3k = [COL_3K, COL_3K + 1, COL_3K + 2].filter(c =>
      _isSessionDate(row[c]) && _isBlackText(_fc(fontColors, r, c))
    ).length;

    // 3回券エリアが黄色（チケット保有・未使用）
    const has3kYellow = sessions3k === 0 && [COL_3K, COL_3K + 1, COL_3K + 2].some(c =>
      _isYellow(_bg(backgrounds, r, c))
    );

    const ticket = _getCurrentTicket(row, r, fontColors, backgrounds);
    const base   = _baseInfo(row);

    if (ticket) {
      // 4・8・12回券 購入者
      const t3k = has3kYellow ? '0/3 未使用'
                : sessions3k > 0 ? (sessions3k + '/3 ' + (tVal === '〇' ? '完了' : '使用中'))
                : (tVal === '〇' ? '完了' : '-');
      const entry = [...base, t3k, ticket.used, ticket.remaining];

      if      (isExpired)  expired.push([...base, tVal || '-', ticket.type + '回券']);
      else if (isChurned)  _push(chu12, chu8, chu4, ticket.type, entry);
      else                 _push(act12, act8, act4, ticket.type, entry);

    } else if (sessions3k > 0 || has3kYellow) {
      // 3回券のみ
      const used3k      = has3kYellow ? 0 : sessions3k;
      const remaining3k = 3 - used3k;
      const entry = [...base, used3k, remaining3k];

      if      (isExpired)  expired.push([...base, tVal || '-', '3回券']);
      else if (isChurned)  chu3k.push(entry);
      else                 act3k.push(entry);

    } else {
      // 初回のみ
      const entry = [..._baseNoExpiry(row), tVal || '-'];

      if      (isExpired)  expired.push([...base, tVal || '-', '初回のみ']);
      else if (isChurned)  chuFirst.push(entry);
      else                 actFirst.push(entry);
    }
  }

  // ===== 集計シートへ書き出し =====
  let dst = ss.getSheetByName(SUMMARY_SHEET_NAME);
  if (!dst) dst = ss.insertSheet(SUMMARY_SHEET_NAME);
  dst.clearContents();

  const out = [];
  const nAct = act12.length + act8.length + act4.length + act3k.length + actFirst.length;
  const nChu = chu12.length + chu8.length + chu4.length + chu3k.length + chuFirst.length;

  // ── 保有顧客 ──
  out.push(['▼ 保有顧客（' + nAct + '名）']);
  out.push([]);
  _appendSection(out, '12回券 購入者',       act12,    _buyerHeader('12'));
  _appendSection(out, '8回券 購入者',        act8,     _buyerHeader('8'));
  _appendSection(out, '4回券 購入者',        act4,     _buyerHeader('4'));
  _appendSection(out, '3回券 購入者',          act3k,    _only3kHeader());
  _appendSection(out, '初回のみ・回数券未購入', actFirst, _firstOnlyHeader());
  out.push([]);

  // ── 離客 ──
  out.push(['▼ 離客（' + nChu + '名）']);
  out.push([]);
  _appendSection(out, '12回券（離客）',          chu12,    _buyerHeader('12'));
  _appendSection(out, '8回券（離客）',           chu8,     _buyerHeader('8'));
  _appendSection(out, '4回券（離客）',           chu4,     _buyerHeader('4'));
  _appendSection(out, '3回券（離客）',           chu3k,    _only3kHeader());
  _appendSection(out, '初回のみ（離客）',         chuFirst, _firstOnlyHeader());
  out.push([]);

  // ── 期限切れ（最下部）──
  out.push(['▼ 有効期限なし・期限切れ（' + expired.length + '名）']);
  out.push(_expiredHeader());
  expired.forEach(r => out.push(r));

  // シートに書き込み
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
    ' 4回券:' + act4.length + ' 3回券のみ:' + act3k.length + ' 初回:' + actFirst.length + '）\n' +
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
  return _isRed(_bg(backgrounds, r, COL_EXPIRY));
}

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
  const groupsNeeded = Math.ceil(typeNum / 4); // 4→1, 8→2, 12→3

  // 最初のセッションスロットが黄色 → 保有・未使用
  const firstSessCol = lastIdx + 3;
  if (!_isSessionDate(row[firstSessCol]) && _isYellow(_bg(backgrounds, r, firstSessCol))) {
    return { type: lastType, used: 0, remaining: typeNum };
  }

  // 黒文字の日付のみ「使用」としてカウント
  let used = 0;
  for (let g = 0; g < groupsNeeded; g++) {
    const startCol = lastIdx + g * 7;
    if (startCol >= row.length) break;
    if (g > 0 && String(row[startCol] || '').trim() === '〇') break; // 次の新規購入で停止
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
function _baseNoExpiry(row) {
  const a = [];
  if (COL_CUSTOMER_ID >= 0) a.push(row[COL_CUSTOMER_ID] || '');
  a.push(String(row[COL_NAME] || '').trim());
  if (COL_GENDER >= 0)     a.push(row[COL_GENDER]    || '');
  if (COL_AGE >= 0)        a.push(row[COL_AGE]       || '');
  if (COL_CYCLE >= 0)      a.push(row[COL_CYCLE]     || '');
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
function _expiredHeader() {
  const h = [];
  if (COL_CUSTOMER_ID >= 0) h.push('顧客番号');
  h.push('氏名');
  if (COL_GENDER >= 0) h.push('性別');
  if (COL_AGE >= 0)    h.push('年代');
  if (COL_CYCLE >= 0)  h.push('周期');
  if (COL_EXPIRY >= 0) h.push('有効期限');
  if (COL_LAST_VISIT >= 0) h.push('最終来店日');
  h.push('更新状況');
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
