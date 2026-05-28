// ============================================================
//  回数券管理システム（統合版）
//  ① 集計表更新        → updateCouponSummary
//  ② 月次売上目標シート → generateNextMonthReport
//
//  両機能とも「顧客管理」シートを共通データソースとして使用
// ============================================================

// ─────────────────────────────────────────
//  ★ 設定値
// ─────────────────────────────────────────
const SOURCE_SHEET_NAME  = '顧客管理';
const SUMMARY_SHEET_NAME = '集計';
const VISIT_CYCLE        = 10;   // 平均来店周期（日）

// 顧客管理シートの列番号（0始まり）
const COL_CUSTOMER_ID = 1;
const COL_NAME        = 4;
const COL_GENDER      = -1;
const COL_AGE         = -1;
const COL_CYCLE       = 3;
const COL_EXPIRY      = 5;
const COL_LAST_VISIT  = 6;
const COL_T           = 12;
const COL_3K          = 13;
const COL_GRP         = 16;

// 月次レポート用カラーコード
const COLORS = {
  HEADER_BG   : '#1F4E79',
  SECTION_BG  : '#2E75B6',
  TOTAL_BG    : '#1F4E79',
  SUBTOTAL_BG : '#DEEAF1',
  NEW_SUB_BG  : '#E2EFDA',
  INPUT_BG    : '#FFFFC0',   // 黄色：手入力
  AUTO_BG     : '#F2F2F2',   // グレー：自動算出
  TYPE_12     : '#BDD7EE',
  TYPE_8      : '#C6EFCE',
  TYPE_4      : '#FFEB9C',
  TYPE_3      : '#FCE4D6',
  OVERDUE     : '#FF0000',
  URGENT      : '#FFC7CE',
};

// ─────────────────────────────────────────
//  メニュー（集計 ＋ 月次レポートを統合）
// ─────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('回数券管理')
    .addItem('集計表を今すぐ更新',              'updateCouponSummary')
    .addItem('自動更新を設定する（1時間ごと）',  'setupTrigger')
    .addItem('自動更新を止める',                'removeTrigger')
    .addSeparator()
    .addItem('▶ 翌月レポートを今すぐ生成',      'generateNextMonthReport')
    .addItem('⚙ 対象月を手動指定して生成',      'generateReportWithPrompt')
    .addItem('📅 月次自動生成を設定（毎月25日）','setupMonthlyTrigger')
    .addToUi();
}

// ============================================================
//  ① 集計表更新（既存機能・変更なし）
// ============================================================
function updateCouponSummary() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!src) {
    SpreadsheetApp.getUi().alert('「' + SOURCE_SHEET_NAME + '」シートが見つかりません。\nシート名（タブ名）を確認してください。');
    return;
  }

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
    const padded = out.map(r => { const a = r.slice(); while (a.length < maxCols) a.push(''); return a; });
    dst.getRange(1, 1, padded.length, maxCols).setValues(padded);
  }

  SpreadsheetApp.getUi().alert(
    '集計完了！\n保有顧客 ' + nAct + '名（12回券:' + act12.length + ' 8回券:' + act8.length +
    ' 4回券:' + act4.length + ' 3回券:' + act3k.length + '）\n離客 ' + nChu + '名\n期限切れ ' + expired.length + '名'
  );
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

// ============================================================
//  ② 月次売上目標シート（新規追加）
// ============================================================

// 翌月レポートを自動生成
function generateNextMonthReport() {
  const today  = new Date();
  const target = new Date(today.getFullYear(), today.getMonth() + 1, 1);
  _runMonthlyReport(target.getFullYear(), target.getMonth() + 1);
}

// 対象月を手動入力して生成
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

// 毎月25日 AM9時に自動実行するトリガーを設定
function setupMonthlyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'scheduledMonthlyRun')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('scheduledMonthlyRun')
    .timeBased().onMonthDay(25).atHour(9).create();
  SpreadsheetApp.getUi().alert('✅ 毎月25日 AM9時に翌月レポートを自動生成するよう設定しました。');
}
function scheduledMonthlyRun() { generateNextMonthReport(); }

// ── コアロジック ──
function _runMonthlyReport(year, month) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!src) {
    SpreadsheetApp.getUi().alert('「' + SOURCE_SHEET_NAME + '」シートが見つかりません。');
    return;
  }

  const range       = src.getDataRange();
  const data        = range.getValues();
  const fontColors  = range.getFontColors();
  const backgrounds = range.getBackgrounds();

  const targetStart = new Date(year, month - 1, 1);
  const targetEnd   = new Date(year, month, 0);   // 月末日

  const customers = [];

  for (let r = 0; r < data.length; r++) {
    const row  = data[r];
    const name = String(row[COL_NAME] || '').trim();
    if (!name || _isLabelRow(name)) continue;

    // 期限切れ・離客はスキップ
    if (_checkExpired(row, r, backgrounds)) continue;
    const tVal = String(row[COL_T] || '').trim();
    if (tVal === '×' || tVal === '✕') continue;

    // チケット情報（既存ロジック再利用）
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

    // 更新予定日 = 最終来店日 + 残り回数 × 来店周期
    const lastVisitRaw = row[COL_LAST_VISIT];
    const lastVisit    = lastVisitRaw instanceof Date ? lastVisitRaw : null;
    let renewalDate    = null;
    if (lastVisit) {
      renewalDate = new Date(lastVisit.getTime());
      renewalDate.setDate(renewalDate.getDate() + remaining * VISIT_CYCLE);
    }

    // ステータス分類
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

    customers.push({
      id: row[COL_CUSTOMER_ID] || '',
      name, couponType, lastVisit, remaining, renewalDate, status,
    });
  }

  const monthStr = year + '年' + month + '月';
  const sheetId  = String(year) + String(month).padStart(2, '0');

  _createListSheet(ss, customers, monthStr, sheetId);
  _createSalesSheet(ss, customers, monthStr, sheetId);

  ss.setActiveSheet(ss.getSheetByName('更新リスト_' + sheetId));
  SpreadsheetApp.getUi().alert('✅ ' + monthStr + 'のレポートを生成しました！\n・更新リスト_' + sheetId + '\n・売上目標_' + sheetId);
}

// ── 更新見込みリストシート ──
function _createListSheet(ss, customers, monthStr, sheetId) {
  const name = '更新リスト_' + sheetId;
  const old  = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet(name);

  let row = 1;

  // タイトル
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

  const typeColorMap = { '12回券': COLORS.TYPE_12, '8回券': COLORS.TYPE_8, '4回券': COLORS.TYPE_4, '3回券': COLORS.TYPE_3 };

  for (const sec of sections) {
    const group = customers.filter(sec.filter);
    if (!group.length) continue;

    row++;
    _mergeSet(sh, row, 1, 7, sec.title + '　（' + group.length + '名）',
      { bold: true, fontSize: 10, bg: COLORS.SECTION_BG, color: '#FFFFFF', align: 'left', height: 22 });
    row++;

    // ヘッダー行
    const hRow = sh.getRange(row, 1, 1, 7);
    hRow.setValues([['顧客番号', '氏名', '券種', '最終来店日', '残り回数', '更新予定日', 'ステータス']]);
    _styleRange(hRow, { bg: COLORS.HEADER_BG, color: '#FFFFFF', bold: true, align: 'center' });
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
        const r  = sh.getRange(row, 1, 1, 7);
        r.setValues([[c.id, c.name, c.couponType, lv, c.remaining, rv, c.status]]);
        if (bg) r.setBackground(bg);
        if (fc) r.setFontColor(fc);
        sh.setRowHeight(row, 17);
        row++;
      }
    }
  }

  [10, 18, 8, 12, 8, 12, 22].forEach((w, i) => sh.setColumnWidth(i + 1, w * 7));
  sh.setFrozenRows(1);
}

// ── 売上目標シート（d390399準拠：B=自動算出 C=手入力 D=単価 E=小計 F=備考） ──
function _createSalesSheet(ss, customers, monthStr, sheetId) {
  const name = '売上目標_' + sheetId;
  const old  = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet(name);

  // 列幅設定
  [28, 16, 16, 14, 16, 14].forEach((w, i) => sh.setColumnWidth(i + 1, w * 7));

  // 当月更新見込み件数カウント
  const types    = ['12回券', '8回券', '4回券', '3回券'];
  const autoCnt  = {};
  types.forEach(t => { autoCnt[t] = customers.filter(c => c.couponType === t && c.status.includes('当月')).length; });

  const typeColorMap = { '12回券': COLORS.TYPE_12, '8回券': COLORS.TYPE_8, '4回券': COLORS.TYPE_4, '3回券': COLORS.TYPE_3 };

  let r = 1;

  // ── タイトル ──
  _mergeSet(sh, r, 1, 6, '売上目標設定シート　／　' + monthStr,
    { fontSize: 15, bold: true, color: '#1F4E79', align: 'center', height: 36 });
  r++;
  _mergeSet(sh, r, 1, 6, '※ 黄色セルを入力してください　　灰色セル＝自動算出（変更不要）',
    { fontSize: 9, color: '#C00000', italic: true, align: 'left', height: 18 });
  r++;
  r++;  // 空行

  // ── 列ヘッダー ──
  const colHdr = [['項目', '更新見込み件数\n（自動算出）', '実際の顧客数\n【手入力】', '単価（円）\n【要入力】', '小計（円）', '備考']];
  const hdrRange = sh.getRange(r, 1, 1, 6);
  hdrRange.setValues(colHdr);
  _styleRange(hdrRange, { bg: COLORS.HEADER_BG, color: '#FFFFFF', bold: true, align: 'center', wrap: true });
  sh.setRowHeight(r, 40);
  r++;

  // ── 既存顧客セクション ──
  _mergeSet(sh, r, 1, 6, '■ 既存顧客　回数券更新見込み（' + monthStr + '）',
    { bold: true, fontSize: 11, bg: COLORS.SECTION_BG, color: '#FFFFFF', align: 'left', height: 22 });
  r++;

  const ticketStartRow = r;
  const typeRows = {};

  for (const ct of types) {
    typeRows[ct] = r;
    const cnt = autoCnt[ct];
    const bg  = typeColorMap[ct];

    // A: 項目名
    _setCell(sh, r, 1, ct + ' 更新見込み', { bold: true, bg });
    // B: 自動算出件数（グレー・参照用）
    _setCell(sh, r, 2, cnt, { bg: COLORS.AUTO_BG, align: 'center', format: '0' });
    // C: 実際の顧客数（黄色・手入力、初期値は自動算出値）
    _setCell(sh, r, 3, cnt, { bold: true, color: '#C00000', bg: COLORS.INPUT_BG, align: 'center', format: '0' });
    // D: 単価（黄色・要入力）
    _setCell(sh, r, 4, 0,   { bold: true, color: '#C00000', bg: COLORS.INPUT_BG, align: 'center', format: '#,##0' });
    // E: 小計 = C × D（グレー・自動）
    sh.getRange(r, 5).setFormula('=C' + r + '*D' + r);
    _setCell(sh, r, 5, null, { bold: true, bg: COLORS.AUTO_BG, align: 'right', format: '#,##0' });
    sh.getRange(r, 5).setFormula('=C' + r + '*D' + r);
    // F: 備考
    _setCell(sh, r, 6, ct.replace('回券', '') + '枚綴り', {});
    sh.setRowHeight(r, 22);
    r++;
  }

  const ticketEndRow = r - 1;

  // 既存小計行
  sh.getRange(r, 1, 1, 4).merge();
  _setCell(sh, r, 1, '既存顧客　小計', { bold: true, fontSize: 11, color: '#1F4E79', bg: COLORS.SUBTOTAL_BG, align: 'right' });
  sh.getRange(r, 5).setFormula('=SUM(E' + ticketStartRow + ':E' + ticketEndRow + ')');
  _setCell(sh, r, 5, null, { bold: true, fontSize: 11, color: '#1F4E79', bg: COLORS.SUBTOTAL_BG, align: 'right', format: '#,##0' });
  sh.getRange(r, 5).setFormula('=SUM(E' + ticketStartRow + ':E' + ticketEndRow + ')');
  _setCell(sh, r, 6, '', { bg: COLORS.SUBTOTAL_BG });
  sh.setRowHeight(r, 24);
  const existSubRow = r;
  r++;

  // ── 新規顧客セクション ──
  r++;
  _mergeSet(sh, r, 1, 6, '■ 新規顧客',
    { bold: true, fontSize: 11, bg: COLORS.SECTION_BG, color: '#FFFFFF', align: 'left', height: 22 });
  r++;

  // A-B: ラベル（結合）
  sh.getRange(r, 1, 1, 2).merge();
  _setCell(sh, r, 1, '新規顧客数', { bold: true, bg: '#FFF2CC' });
  // C: 新規顧客数（黄色・手入力）
  _setCell(sh, r, 3, 0, { bold: true, color: '#C00000', bg: COLORS.INPUT_BG, align: 'center', format: '0' });
  // D: 単価（黄色）
  _setCell(sh, r, 4, 0, { bold: true, color: '#C00000', bg: COLORS.INPUT_BG, align: 'center', format: '#,##0' });
  // E: 小計 = C × D
  sh.getRange(r, 5).setFormula('=C' + r + '*D' + r);
  _setCell(sh, r, 5, null, { bold: true, bg: COLORS.AUTO_BG, align: 'right', format: '#,##0' });
  sh.getRange(r, 5).setFormula('=C' + r + '*D' + r);
  // F: 備考
  _setCell(sh, r, 6, '初回チケット平均単価', {});
  sh.setRowHeight(r, 22);
  const newRow = r;
  r++;

  // 新規小計行
  sh.getRange(r, 1, 1, 4).merge();
  _setCell(sh, r, 1, '新規顧客　小計', { bold: true, fontSize: 11, color: '#1F4E79', bg: COLORS.NEW_SUB_BG, align: 'right' });
  sh.getRange(r, 5).setFormula('=E' + newRow);
  _setCell(sh, r, 5, null, { bold: true, fontSize: 11, color: '#1F4E79', bg: COLORS.NEW_SUB_BG, align: 'right', format: '#,##0' });
  sh.getRange(r, 5).setFormula('=E' + newRow);
  _setCell(sh, r, 6, '', { bg: COLORS.NEW_SUB_BG });
  sh.setRowHeight(r, 24);
  const newSubRow = r;
  r++;

  // ── 合計行 ──
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
  r++;

  // ── 件数サマリー表 ──
  r += 2;
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
    const row  = sh.getRange(r, 1, 1, 6);
    row.setValues([[ct, may, over, fut, noc, may + over + fut + noc]]).setHorizontalAlignment('center');
    sh.getRange(r, 1).setBackground(typeColorMap[ct]).setFontWeight('bold');
    sh.getRange(r, 2).setBackground(COLORS.URGENT).setFontWeight('bold');
    sh.setRowHeight(r, 20);
    r++;
  }

  // ── 入力ガイド ──
  r += 2;
  _mergeSet(sh, r, 1, 6, '■ 入力手順',
    { bold: true, fontSize: 11, color: '#1F4E79', align: 'left', height: 22 });
  r++;

  const guides = [
    '① C列「実際の顧客数」（C列 各回数券行）を実績に合わせて入力（初期値は自動算出値）',
    '② D列「単価」を各回数券の金額に入力　→ E列「小計」が自動計算されます',
    '③ 新規顧客数（C列 新規顧客行）と初回チケット平均単価（D列）を入力',
    '④ B列「更新見込み件数」はシステム自動算出の参照値です（変更不要）',
    '⑤ 売上目標（E列 合計行）に当月の最低見込み売上が自動表示されます',
  ];
  for (const g of guides) {
    _mergeSet(sh, r, 1, 6, g, { fontSize: 10, align: 'left', height: 20 });
    r++;
  }
}

// ============================================================
//  スタイルヘルパー関数
// ============================================================
function _setCell(sh, row, col, value, opts) {
  const c = sh.getRange(row, col);
  if (value !== null && value !== undefined) c.setValue(value);
  if (opts.bold)    c.setFontWeight('bold');
  if (opts.fontSize) c.setFontSize(opts.fontSize);
  if (opts.color)   c.setFontColor(opts.color);
  if (opts.italic)  c.setFontStyle('italic');
  if (opts.bg)      c.setBackground(opts.bg);
  if (opts.align)   c.setHorizontalAlignment(opts.align);
  if (opts.format)  c.setNumberFormat(opts.format);
  if (opts.wrap)    c.setWrap(true);
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
  if (opts.align)    range.setHorizontalAlignment(opts.align === 'center' ? 'center' : opts.align === 'right' ? 'right' : 'left');
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

// ============================================================
//  共通ヘルパー関数（既存コードより・変更なし）
// ============================================================
function _push(list12, list8, list4, type, entry) {
  if (type === '12') list12.push(entry); else if (type === '8') list8.push(entry); else list4.push(entry);
}
function _appendSection(out, title, data, header) {
  out.push(['【' + title + '】（' + data.length + '名）']);
  if (data.length > 0) { out.push(header); data.forEach(r => out.push(r)); }
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
    if (marker === '〇') { const type = _normalizeType(row[c+1]); if (type) { lastIdx = c; lastType = type; } }
  }
  if (lastIdx === -1) return null;
  const typeNum = parseInt(lastType);
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
      if (col < row.length && _isSessionDate(row[col]) && _isBlackText(_fc(fontColors, r, col))) used++;
    }
  }
  return { type: lastType, used, remaining: typeNum - used };
}
function _fc(fontColors, r, c) { return (r < fontColors.length && c < fontColors[r].length) ? fontColors[r][c] : ''; }
function _bg(backgrounds, r, c) { return (r < backgrounds.length && c < backgrounds[r].length) ? backgrounds[r][c] : ''; }
function _hexToRgb(hex) {
  if (!hex || hex.length < 7) return null;
  return { r: parseInt(hex.slice(1,3),16), g: parseInt(hex.slice(3,5),16), b: parseInt(hex.slice(5,7),16) };
}
function _isBlackText(fc) { const c=(fc||'').toLowerCase(); return !c||c==='#000000'||c==='#1f1f1f'||c==='#202124'; }
function _isYellow(bg) { const rgb=_hexToRgb(bg); return rgb?(rgb.r>200&&rgb.g>170&&rgb.b<120):false; }
function _isRed(bg) { const rgb=_hexToRgb(bg); return rgb?(rgb.r>160&&rgb.g<120&&rgb.b<120):false; }
function _baseInfo(row) {
  const a = [];
  if (COL_CUSTOMER_ID >= 0) a.push(row[COL_CUSTOMER_ID] || '');
  a.push(String(row[COL_NAME] || '').trim());
  if (COL_GENDER >= 0)     a.push(row[COL_GENDER] || '');
  if (COL_AGE >= 0)        a.push(row[COL_AGE] || '');
  if (COL_CYCLE >= 0)      a.push(row[COL_CYCLE] || '');
  if (COL_EXPIRY >= 0)     a.push(_fmtDate(row[COL_EXPIRY]));
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
  h.push('3回券'); h.push(type + '回券(使用)'); h.push(type + '回券(残り)');
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
  h.push('3回券(使用)'); h.push('3回券(残り)');
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
function _fmtDate(val) {
  if (!val) return '';
  if (val instanceof Date) return val.getFullYear() + '/' + (val.getMonth()+1) + '/' + val.getDate();
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
  const SKIP = new Set(['●','〇','×','✕','△','‐','-','都度','FALSE','TRUE','NG','機械のみ 当日のみ']);
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
