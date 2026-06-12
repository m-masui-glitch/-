// =====================================================================
// 回数券集計スクリプト — スタンダード版（全店舗共通）
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

// ===== レポート設定 =====
const VISIT_CYCLE = 10;  // 来店周期（日）

const COLORS = {
  HEADER_BG   : '#1F4E79',
  SECTION_BG  : '#2E75B6',
  TOTAL_BG    : '#1F4E79',
  SUBTOTAL_BG : '#DEEAF1',
  NEW_SUB_BG  : '#E2EFDA',
  INPUT_BG    : '#FFFFC0',
  AUTO_BG     : '#F2F2F2',
  TYPE_12     : '#BDD7EE',
  TYPE_8      : '#C6EFCE',
  TYPE_4      : '#FFEB9C',
  TYPE_3      : '#FCE4D6',
  OVERDUE     : '#FF0000',
  URGENT      : '#FFC7CE',
};
// =======================

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
      // F列色分け：残り日数に応じて3段階
      const daysLeft = (expiry - today) / 86400000;
      let expiryBg = null;
      if (daysLeft < 0)          expiryBg = '#b7b7b7'; // 期限切れ：グレー
      else if (daysLeft <= 30)   expiryBg = '#fce4ec'; // 残り0〜30日：ピンク
      else if (daysLeft <= 60)   expiryBg = '#fffde7'; // 残り31〜60日：クリーム
      expiryBgs.push([expiryBg]);
    } else {
      expiryVals.push([row[COL_EXPIRY] !== undefined ? row[COL_EXPIRY] : '']);
      expiryBgs.push([null]);
    }

    // G列：最終来店日（黒文字の最新日付）
    const lastVisit = _calcLastVisit(row, fontColors, r);
    if (lastVisit) {
      lastVisitVals.push([lastVisit]);
      const oneMonthAgo  = new Date(today.getFullYear(), today.getMonth() - 1, today.getDate());
      const sixMonthsAgo = new Date(today.getFullYear(), today.getMonth() - 6, today.getDate());
      let lastVisitBg = null;
      if (lastVisit < sixMonthsAgo)     lastVisitBg = '#70ad47'; // 6ヶ月以上来店なし：緑
      else if (lastVisit < oneMonthAgo) lastVisitBg = '#fff2cc'; // 30日以上来店なし：薄黄色
      lastVisitBgs.push([lastVisitBg]);
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
  // G2 は日付ではなく「最終来店日」ラベルとして固定
  if (numRows >= 2) {
    src.getRange(2, COL_LAST_VISIT + 1).setValue('最終来店日').setNumberFormat('@').setBackground(null);
  }
}

// 有効期限を計算
// - 4/8/12回券あり → 最新〇の購入日（col+2）+ 6ヶ月の月末
// - 3回券のみ     → L列（初回日）+ 3ヶ月の月末
function _calcExpiry(row, fontColors, r) {
  let lastPurchaseDate = null;

  for (let c = COL_GRP; c + 3 < row.length; c += 7) {
    const marker = String(row[c] || '').trim();
    if (marker === '〇') {
      const type = _normalizeType(row[c + 1]);
      if (type) {
        const d = _toDate(row[c + 2]);
        if (d) lastPurchaseDate = d;
      }
    }
  }

  if (lastPurchaseDate) return _expiryDate(lastPurchaseDate, 6);

  const firstVisit = _toDate(row[COL_FIRST_VISIT_3K]);
  if (firstVisit) return _expiryDate(firstVisit, 3);

  return null;
}

// 最終来店日を計算（黒文字の日付の中で最新）
function _calcLastVisit(row, fontColors, r) {
  let latest = null;
  let currentPurchaseDate = null;

  const check = (val, c, refDate) => {
    if (_isSessionDate(val) && _isBlackText(_fc(fontColors, r, c))) {
      const d = _toDate(val, refDate);
      if (d && (!latest || d > latest)) latest = d;
    }
  };

  // 3回券エリア（初回来店日を参照日として使用）
  const ref3k = _toDate(row[COL_FIRST_VISIT_3K]);
  for (let c = COL_3K; c <= COL_3K + 2 && c < row.length; c++) check(row[c], c, ref3k);

  // 7列グループのセッションスロット（〇の購入日を参照日として使用）
  for (let c = COL_GRP; c + 3 < row.length; c += 7) {
    const marker = String(row[c] || '').trim();
    if (marker === '〇') {
      const pd = _toDate(row[c + 2]);
      if (pd) currentPurchaseDate = pd;
    }
    for (let s = 3; s <= 6; s++) {
      const col = c + s;
      if (col < row.length) check(row[col], col, currentPurchaseDate);
    }
  }

  return latest;
}

// 購入日 + N ヶ月 - 1日（例: 1月10日 + 6 → 7月9日、3月7日 + 6 → 9月6日）
function _expiryDate(date, months) {
  return new Date(date.getFullYear(), date.getMonth() + months, date.getDate() - 1);
}

// スプレッドシートのタイムゾーンをキャッシュして返す
let _TIMEZONE = null;
function _tz() {
  if (!_TIMEZONE) {
    try { _TIMEZONE = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(); }
    catch(e) { _TIMEZONE = Session.getScriptTimeZone(); }
  }
  return _TIMEZONE;
}

// 値をDateオブジェクトに変換
// refDate を渡すと、M/D形式の年を「refDate以降になる年」で補完する
function _toDate(val, refDate) {
  if (!val) return null;
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return null;
    // スプレッドシートのタイムゾーンで正しく日付を読む（UTC→ローカル変換のズレを防ぐ）
    const s = Utilities.formatDate(val, _tz(), 'yyyy/M/d');
    const p = s.split('/');
    return new Date(+p[0], +p[1] - 1, +p[2]);
  }
  const s = String(val).trim();
  if (!s) return null;
  // 年付きフォーマット（2026/4/8, 2026-04-08, 2026年4月8日）
  const full = s.match(/(\d{4})[年\/\-](\d{1,2})[月\/\-](\d{1,2})/);
  if (full) return new Date(+full[1], +full[2] - 1, +full[3]);
  // 月/日 または 月月日日（年なし）
  const md = s.match(/^(\d{1,2})[\/月](\d{1,2})日?$/);
  if (md) {
    const month    = +md[1] - 1;
    const day      = +md[2];
    const baseYear = refDate ? refDate.getFullYear() : new Date().getFullYear();
    const d        = new Date(baseYear, month, day);
    // 購入日より前になる場合は翌年として補完
    if (refDate && d < refDate) return new Date(baseYear + 1, month, day);
    return d;
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

// ============================================================
//  月次レポート生成
// ============================================================

function generateNextMonthReport() {
  const today  = new Date();
  const target = new Date(today.getFullYear(), today.getMonth() + 1, 1);
  _runMonthlyReport(target.getFullYear(), target.getMonth() + 1);
}

function generateReportWithPrompt() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.prompt('対象月を入力してください', '例: 2026/7　または　202607', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const text  = res.getResponseText().replace(/[\/\-\s]/g, '');
  const year  = parseInt(text.slice(0, 4));
  const month = parseInt(text.slice(4, 6));
  if (isNaN(year) || isNaN(month) || month < 1 || month > 12) {
    ui.alert('年月の形式が正しくありません。例: 2026/7');
    return;
  }
  _runMonthlyReport(year, month);
}

function setupMonthlyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'scheduledMonthlyRun')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('scheduledMonthlyRun')
    .timeBased().onMonthDay(25).atHour(9).create();
  SpreadsheetApp.getUi().alert('✅ 毎月25日 AM9時に翌月レポートを自動生成するよう設定しました。');
}
function scheduledMonthlyRun() { generateNextMonthReport(); }

// ── レポート出力先スプレッドシート管理 ──

const REPORT_SS_KEY = 'REPORT_SPREADSHEET_ID';

function _getReportSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  const id    = props.getProperty(REPORT_SS_KEY);
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch(e) {}
  }
  const sourceName = SpreadsheetApp.getActiveSpreadsheet().getName();
  const newSS = SpreadsheetApp.create(sourceName + '【月次レポート】');
  const sheets = newSS.getSheets();
  if (sheets.length > 0) {
    newSS.insertSheet('_dummy');
    newSS.deleteSheet(sheets[0]);
  }
  props.setProperty(REPORT_SS_KEY, newSS.getId());
  return newSS;
}

function showReportSpreadsheetInfo() {
  const props = PropertiesService.getScriptProperties();
  const id    = props.getProperty(REPORT_SS_KEY);
  const ui    = SpreadsheetApp.getUi();
  if (!id) {
    ui.alert('まだレポート出力先が設定されていません。\n「翌月レポートを生成」を実行すると自動で作成されます。');
    return;
  }
  try {
    const ss  = SpreadsheetApp.openById(id);
    ui.alert('📂 現在のレポート出力先\n\n' + ss.getName() + '\n\n' + ss.getUrl() + '\n\n※URLをコピーして共有してください。');
  } catch(e) {
    ui.alert('登録済みのIDが無効です。「レポート出力先を新規作成」で再設定してください。');
  }
}

function resetReportSpreadsheet() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.alert('新しいレポート用スプレッドシートを作成します。\n（既存のレポートはそのまま残ります）\n\nよろしいですか？', ui.ButtonSet.OK_CANCEL);
  if (res !== ui.Button.OK) return;
  PropertiesService.getScriptProperties().deleteProperty(REPORT_SS_KEY);
  const ss = _getReportSpreadsheet();
  ui.alert('✅ 新しいレポート出力先を作成しました。\n\n' + ss.getUrl() + '\n\n※このURLを関係者に共有してください。');
}

// ── コアロジック ──
function _runMonthlyReport(year, month) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!src) {
    SpreadsheetApp.getUi().alert('「' + SOURCE_SHEET_NAME + '」シートが見つかりません。');
    return;
  }

  _updateDatesInSheet(src);

  const range       = src.getDataRange();
  const data        = range.getValues();
  const fontColors  = range.getFontColors();
  const backgrounds = range.getBackgrounds();

  const targetStart = new Date(year, month - 1, 1);
  const targetEnd   = new Date(year, month, 0);

  const customers = [];

  for (let r = 0; r < data.length; r++) {
    const row  = data[r];
    const name = String(row[COL_NAME] || '').trim();
    if (!name || _isLabelRow(name)) continue;
    if (_checkExpired(row, r, backgrounds)) continue;
    const tVal = String(row[COL_T] || '').trim();
    if (tVal === '×' || tVal === '✕') continue;

    const ticket = _getCurrentTicket(row, r, fontColors, backgrounds);
    const sessions3k = [COL_3K, COL_3K + 1, COL_3K + 2].filter(c =>
      _isSessionDate(row[c]) && _isBlackText(_fc(fontColors, r, c))
    ).length;
    const has3kYellow = sessions3k === 0 && [COL_3K, COL_3K + 1, COL_3K + 2].some(c =>
      _isYellow(_bg(backgrounds, r, c))
    );

    let couponType, remaining;
    if (ticket) {
      couponType = ticket.type + '回券';
      remaining  = ticket.remaining;
    } else if (sessions3k > 0 || has3kYellow) {
      couponType = '3回券';
      remaining  = 3 - (has3kYellow ? 0 : sessions3k);
    } else {
      continue;
    }

    const lastVisitRaw = row[COL_LAST_VISIT];
    const lastVisit    = lastVisitRaw instanceof Date ? lastVisitRaw : null;
    let renewalDate    = null;
    if (lastVisit) {
      renewalDate = new Date(lastVisit.getTime());
      renewalDate.setDate(renewalDate.getDate() + remaining * VISIT_CYCLE);
    }

    let status;
    if (!renewalDate) {
      status = '算出不可（都度）';
    } else if (remaining === 0 && renewalDate < targetStart) {
      status = '要確認（残0・期限超過）';
    } else if (renewalDate < targetStart) {
      status = '要確認（更新期限超過）';
    } else if (remaining === 0) {
      status = '当月更新必要（残0）';
    } else if (renewalDate <= targetEnd) {
      status = '当月更新見込み';
    } else {
      status = '翌月以降';
    }

    customers.push({ id: row[COL_CUSTOMER_ID] || '', name, couponType, lastVisit, remaining, renewalDate, status });
  }

  const monthStr = year + '年' + month + '月';
  const sheetId  = String(year) + String(month).padStart(2, '0');
  const reportSS = _getReportSpreadsheet();
  _createListSheet(reportSS, customers, monthStr, sheetId);
  _createSalesSheet(reportSS, customers, monthStr, sheetId);

  SpreadsheetApp.getUi().alert(
    '✅ ' + monthStr + 'のレポートを生成しました！\n\n' +
    '出力先スプレッドシート：\n' + reportSS.getUrl() + '\n\n' +
    '※メニュー「レポート出力先を確認・変更」からURLを再確認できます。'
  );
}

// ── 更新見込みリストシート ──
function _createListSheet(ss, customers, monthStr, sheetId) {
  const name = '更新リスト_' + sheetId;
  const old  = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet(name);

  const typeColorMap = { '12回券': COLORS.TYPE_12, '8回券': COLORS.TYPE_8, '4回券': COLORS.TYPE_4, '3回券': COLORS.TYPE_3 };

  let row = 1;
  _mergeSet(sh, row, 1, 7, monthStr + ' 回数券更新見込み客リスト（来店周期' + VISIT_CYCLE + '日ベース）',
    { fontSize: 13, bold: true, color: '#FFFFFF', bg: COLORS.HEADER_BG, align: 'center', height: 30 });
  row++;
  _mergeSet(sh, row, 1, 7, '更新予定日 = 最終来店日 + 残り回数 × ' + VISIT_CYCLE + '日　／　赤＝期限超過（要フォロー）',
    { fontSize: 9, color: '#595959', align: 'left', height: 16 });
  row++;

  const sections = [
    { title: '▼ 当月更新見込み（残0含む）',      filter: c => c.status.includes('当月') },
    { title: '▼ 要確認：更新期限超過（前月以前）', filter: c => c.status.includes('要確認') },
    { title: '▼ 翌月以降',                       filter: c => c.status === '翌月以降' },
    { title: '▼ 算出不可（都度来院）',             filter: c => c.status.includes('算出不可') },
  ];

  for (const sec of sections) {
    const group = customers.filter(sec.filter);
    if (!group.length) continue;

    row++;
    _mergeSet(sh, row, 1, 7, sec.title + '　（' + group.length + '名）',
      { bold: true, fontSize: 10, bg: COLORS.SECTION_BG, color: '#FFFFFF', align: 'left', height: 22 });
    row++;

    const hdrRange = sh.getRange(row, 1, 1, 7);
    hdrRange.setValues([['顧客番号', '氏名', '券種', '最終来店日', '残り回数', '更新予定日', 'ステータス']]);
    _styleRange(hdrRange, { bg: COLORS.HEADER_BG, color: '#FFFFFF', bold: true, align: 'center' });
    sh.setRowHeight(row, 20);
    row++;

    for (const ct of ['12回券', '8回券', '4回券', '3回券']) {
      const typeGroup = group.filter(c => c.couponType === ct)
        .sort((a, b) => (a.renewalDate || new Date(9999, 0)) - (b.renewalDate || new Date(9999, 0)));
      if (!typeGroup.length) continue;

      _mergeSet(sh, row, 1, 7, '  【' + ct + '】　' + typeGroup.length + '名',
        { bold: true, bg: typeColorMap[ct], align: 'left', height: 18 });
      row++;

      for (const c of typeGroup) {
        const isOverdue = c.status.includes('要確認');
        const isUrgent  = c.status.includes('残0') && !isOverdue;
        const bg = isOverdue ? COLORS.OVERDUE : isUrgent ? COLORS.URGENT : (c.status.includes('当月') ? typeColorMap[ct] : null);
        const fc = isOverdue ? '#FFFFFF' : null;
        const lv = c.lastVisit   ? _fmtDate(c.lastVisit)   : '—';
        const rv = c.renewalDate ? _fmtDate(c.renewalDate) : '—';
        const rng = sh.getRange(row, 1, 1, 7);
        rng.setValues([[c.id, c.name, c.couponType, lv, c.remaining, rv, c.status]]);
        if (bg) rng.setBackground(bg);
        if (fc) rng.setFontColor(fc);
        sh.setRowHeight(row, 17);
        row++;
      }
    }
  }

  [10, 18, 8, 12, 8, 12, 22].forEach((w, i) => sh.setColumnWidth(i + 1, w * 7));
  sh.setFrozenRows(1);
}

// ── 売上目標シート ──
function _createSalesSheet(ss, customers, monthStr, sheetId) {
  const name = '売上目標_' + sheetId;
  const old  = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet(name);

  [28, 16, 16, 14, 16, 14].forEach((w, i) => sh.setColumnWidth(i + 1, w * 7));

  const types       = ['12回券', '8回券', '4回券', '3回券'];
  const typeColorMap = { '12回券': COLORS.TYPE_12, '8回券': COLORS.TYPE_8, '4回券': COLORS.TYPE_4, '3回券': COLORS.TYPE_3 };
  const PRICES      = { '12回券': 80780, '8回券': 55440, '4回券': 28510, '3回券': 19250 };
  const autoCnt     = {};
  types.forEach(t => { autoCnt[t] = customers.filter(c => c.couponType === t && c.status.includes('当月')).length; });

  let r = 1;

  _mergeSet(sh, r, 1, 6, '売上目標設定シート　／　' + monthStr,
    { fontSize: 15, bold: true, color: '#1F4E79', align: 'center', height: 36 });
  r++;
  _mergeSet(sh, r, 1, 6, '※ 黄色セルを入力してください　　灰色セル＝自動算出（変更不要）',
    { fontSize: 9, color: '#C00000', italic: true, align: 'left', height: 18 });
  r++;
  r++;

  const hdrRange = sh.getRange(r, 1, 1, 6);
  hdrRange.setValues([['項目', '更新見込み件数\n（自動算出）', '実際の顧客数\n【手入力】', '単価（円）', '小計（円）', '備考']]);
  _styleRange(hdrRange, { bg: COLORS.HEADER_BG, color: '#FFFFFF', bold: true, align: 'center', wrap: true });
  sh.setRowHeight(r, 40);
  r++;

  _mergeSet(sh, r, 1, 6, '■ 既存顧客　回数券更新見込み（' + monthStr + '）',
    { bold: true, fontSize: 11, bg: COLORS.SECTION_BG, color: '#FFFFFF', align: 'left', height: 22 });
  r++;

  const ticketStartRow = r;
  for (const ct of types) {
    const cnt   = autoCnt[ct];
    const bg    = typeColorMap[ct];
    const price = PRICES[ct];
    _setCell(sh, r, 1, ct + ' 更新見込み', { bold: true, bg });
    _setCell(sh, r, 2, cnt,   { bg: COLORS.AUTO_BG, align: 'center', format: '0' });
    _setCell(sh, r, 3, cnt,   { bold: true, color: '#C00000', bg: COLORS.INPUT_BG, align: 'center', format: '0' });
    _setCell(sh, r, 4, price, { bold: true, bg: COLORS.INPUT_BG, align: 'center', format: '#,##0' });
    sh.getRange(r, 5).setFormula('=C' + r + '*D' + r);
    _setCell(sh, r, 5, null,  { bold: true, bg: COLORS.AUTO_BG, align: 'right', format: '#,##0' });
    sh.getRange(r, 5).setFormula('=C' + r + '*D' + r);
    _setCell(sh, r, 6, ct.replace('回券', '') + '枚綴り', {});
    sh.setRowHeight(r, 22);
    r++;
  }
  const ticketEndRow = r - 1;

  sh.getRange(r, 1, 1, 4).merge();
  _setCell(sh, r, 1, '既存顧客　小計', { bold: true, fontSize: 11, color: '#1F4E79', bg: COLORS.SUBTOTAL_BG, align: 'right' });
  sh.getRange(r, 5).setFormula('=SUM(E' + ticketStartRow + ':E' + ticketEndRow + ')');
  _setCell(sh, r, 5, null, { bold: true, fontSize: 11, color: '#1F4E79', bg: COLORS.SUBTOTAL_BG, align: 'right', format: '#,##0' });
  sh.getRange(r, 5).setFormula('=SUM(E' + ticketStartRow + ':E' + ticketEndRow + ')');
  _setCell(sh, r, 6, '', { bg: COLORS.SUBTOTAL_BG });
  sh.setRowHeight(r, 24);
  const existSubRow = r;
  r++;

  r++;
  _mergeSet(sh, r, 1, 6, '■ 新規顧客',
    { bold: true, fontSize: 11, bg: COLORS.SECTION_BG, color: '#FFFFFF', align: 'left', height: 22 });
  r++;

  sh.getRange(r, 1, 1, 2).merge();
  _setCell(sh, r, 1, '新規顧客数', { bold: true, bg: '#FFF2CC' });
  _setCell(sh, r, 3, 0,    { bold: true, color: '#C00000', bg: COLORS.INPUT_BG, align: 'center', format: '0' });
  _setCell(sh, r, 4, 3000, { bold: true, bg: COLORS.INPUT_BG, align: 'center', format: '#,##0' });
  sh.getRange(r, 5).setFormula('=C' + r + '*D' + r);
  _setCell(sh, r, 5, null, { bold: true, bg: COLORS.AUTO_BG, align: 'right', format: '#,##0' });
  sh.getRange(r, 5).setFormula('=C' + r + '*D' + r);
  _setCell(sh, r, 6, '初回チケット平均単価', {});
  sh.setRowHeight(r, 22);
  const newRow = r;
  r++;

  sh.getRange(r, 1, 1, 4).merge();
  _setCell(sh, r, 1, '新規顧客　小計', { bold: true, fontSize: 11, color: '#1F4E79', bg: COLORS.NEW_SUB_BG, align: 'right' });
  sh.getRange(r, 5).setFormula('=E' + newRow);
  _setCell(sh, r, 5, null, { bold: true, fontSize: 11, color: '#1F4E79', bg: COLORS.NEW_SUB_BG, align: 'right', format: '#,##0' });
  sh.getRange(r, 5).setFormula('=E' + newRow);
  _setCell(sh, r, 6, '', { bg: COLORS.NEW_SUB_BG });
  sh.setRowHeight(r, 24);
  const newSubRow = r;
  r++;

  r++;
  sh.getRange(r, 1, 1, 4).merge();
  _setCell(sh, r, 1, '▶ 当月　売上目標（最低見込み）',
    { bold: true, fontSize: 12, color: '#FFFFFF', bg: COLORS.TOTAL_BG, align: 'right' });
  sh.getRange(r, 5).setFormula('=E' + existSubRow + '+E' + newSubRow);
  _setCell(sh, r, 5, null, { bold: true, fontSize: 14, color: '#FFFFFF', bg: COLORS.TOTAL_BG, align: 'right', format: '#,##0' });
  sh.getRange(r, 5).setFormula('=E' + existSubRow + '+E' + newSubRow);
  _setCell(sh, r, 6, '', { bg: COLORS.TOTAL_BG });
  sh.setRowHeight(r, 34);
  r++;

  _mergeSet(sh, r, 1, 6, '※ 都度来院客・アップセル分は含まれません（最低限の見込み額）',
    { fontSize: 9, color: '#595959', italic: true, align: 'left', height: 16 });
  r += 3;

  _mergeSet(sh, r, 1, 6, '■ 件数サマリー（参考）',
    { bold: true, fontSize: 11, color: '#1F4E79', align: 'left', height: 22 });
  r++;

  const smHdr = sh.getRange(r, 1, 1, 6);
  smHdr.setValues([['券種', '当月更新見込み', '要確認（期限超過）', '翌月以降', '算出不可（都度）', '合計']]);
  _styleRange(smHdr, { bg: COLORS.HEADER_BG, color: '#FFFFFF', bold: true, align: 'center' });
  sh.setRowHeight(r, 20);
  r++;

  for (const ct of types) {
    const may  = customers.filter(c => c.couponType === ct && c.status.includes('当月')).length;
    const over = customers.filter(c => c.couponType === ct && c.status.includes('要確認')).length;
    const fut  = customers.filter(c => c.couponType === ct && c.status === '翌月以降').length;
    const noc  = customers.filter(c => c.couponType === ct && c.status.includes('算出不可')).length;
    sh.getRange(r, 1, 1, 6)
      .setValues([[ct, may, over, fut, noc, may + over + fut + noc]])
      .setHorizontalAlignment('center');
    sh.getRange(r, 1).setBackground(typeColorMap[ct]).setFontWeight('bold');
    sh.getRange(r, 2).setBackground(COLORS.URGENT).setFontWeight('bold');
    sh.setRowHeight(r, 20);
    r++;
  }

  r += 2;
  _mergeSet(sh, r, 1, 6, '■ 入力手順',
    { bold: true, fontSize: 11, color: '#1F4E79', align: 'left', height: 22 });
  r++;
  const guides = [
    '① C列「実際の顧客数」を実績に合わせて入力（初期値は自動算出値）',
    '② D列「単価」は固定値が入力済みです。変更が必要な場合は上書きしてください',
    '③ 新規顧客数（C列 新規顧客行）を入力してください',
    '④ B列「更新見込み件数」はシステム自動算出の参照値です（変更不要）',
    '⑤ 売上目標（E列 合計行）に当月の最低見込み売上が自動表示されます',
  ];
  for (const g of guides) {
    _mergeSet(sh, r, 1, 6, g, { fontSize: 10, align: 'left', height: 20 });
    r++;
  }
}

// ============================================================
//  スタイルヘルパー
// ============================================================

function _setCell(sh, row, col, value, opts) {
  const c = sh.getRange(row, col);
  if (value !== null && value !== undefined) c.setValue(value);
  if (opts.bold)     c.setFontWeight('bold');
  if (opts.fontSize) c.setFontSize(opts.fontSize);
  if (opts.color)    c.setFontColor(opts.color);
  if (opts.italic)   c.setFontStyle('italic');
  if (opts.bg)       c.setBackground(opts.bg);
  if (opts.align)    c.setHorizontalAlignment(opts.align);
  if (opts.format)   c.setNumberFormat(opts.format);
  if (opts.wrap)     c.setWrap(true);
  c.setBorder(true, true, true, true, null, null);
}

function _mergeSet(sh, row, col, span, value, opts) {
  const range = sh.getRange(row, col, 1, span);
  range.merge().setValue(value);
  if (opts.bold)     range.setFontWeight('bold');
  if (opts.fontSize) range.setFontSize(opts.fontSize);
  if (opts.color)    range.setFontColor(opts.color);
  if (opts.italic)   range.setFontStyle('italic');
  if (opts.bg)       range.setBackground(opts.bg);
  if (opts.align)    range.setHorizontalAlignment(opts.align === 'right' ? 'right' : opts.align === 'center' ? 'center' : 'left');
  range.setVerticalAlignment('middle');
  if (opts.wrap)     range.setWrap(true);
  if (opts.height)   sh.setRowHeight(row, opts.height);
  range.setBorder(true, true, true, true, null, null);
}

function _styleRange(range, opts) {
  if (opts.bg)    range.setBackground(opts.bg);
  if (opts.color) range.setFontColor(opts.color);
  if (opts.bold)  range.setFontWeight('bold');
  if (opts.align) range.setHorizontalAlignment(opts.align);
  if (opts.wrap)  range.setWrap(true);
  range.setBorder(true, true, true, true, null, null);
}

// ===== メニュー・トリガー =====
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('回数券管理');
  menu.addItem('集計表を今すぐ更新', 'updateCouponSummary');
  menu.addItem('F・G列の日付のみ更新する', 'updateDates');
  menu.addItem('自動更新を設定する（1時間ごと）', 'setupTrigger');
  menu.addItem('自動更新を止める', 'removeTrigger');
  menu.addSeparator();
  menu.addItem('翌月レポートを今すぐ生成', 'generateNextMonthReport');
  menu.addItem('対象月を手動指定してレポート生成', 'generateReportWithPrompt');
  menu.addItem('月次自動生成を設定（毎月25日）', 'setupMonthlyTrigger');
  menu.addSeparator();
  menu.addItem('レポート出力先を確認・変更', 'showReportSpreadsheetInfo');
  menu.addItem('レポート出力先を新規作成', 'resetReportSpreadsheet');
  menu.addToUi();
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
