// ============================================================
//  回数券管理システム（統合版）
//  ① 集計表更新        : updateCouponSummary
//  ② 月次売上目標レポート: generateNextMonthReport
//
//  データソース：「顧客管理」シート（共通）
// ============================================================

// ─────────────────────────────────────────
//  ★ 設定値（必要に応じて変更してください）
// ─────────────────────────────────────────
const SOURCE_SHEET_NAME  = '顧客管理';
const SUMMARY_SHEET_NAME = '集計';

// 顧客管理シートの列番号（0始まり）
const COL_CUSTOMER_ID = 1;
const COL_NAME        = 4;
const COL_GENDER      = -1;   // 列がなければ -1
const COL_AGE         = -1;
const COL_CYCLE       = 3;
const COL_EXPIRY      = 5;
const COL_LAST_VISIT  = 6;
const COL_T           = 12;
const COL_3K          = 13;
const COL_GRP         = 16;

// 月次レポート設定
const VISIT_CYCLE = 10;        // 平均来店周期（日）
const TYPE_COLORS = {
  '12回券': '#BDD7EE',
  '8回券' : '#C6EFCE',
  '4回券' : '#FFEB9C',
  '3回券' : '#FCE4D6',
};

// ─────────────────────────────────────────
//  メニュー（集計 ＋ 月次レポートを統合）
// ─────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('回数券管理')
    // ── 既存機能 ──
    .addItem('集計表を今すぐ更新',           'updateCouponSummary')
    .addItem('自動更新を設定する（1時間ごと）', 'setupTrigger')
    .addItem('自動更新を止める',              'removeTrigger')
    .addSeparator()
    // ── 月次レポート機能（新規追加） ──
    .addItem('▶ 翌月レポートを今すぐ生成',    'generateNextMonthReport')
    .addItem('⚙ 対象月を手動指定して生成',    'generateReportWithPrompt')
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
//  ② 月次売上目標レポート（新規追加）
//     「顧客管理」シートを直接読んで翌月の更新見込みを生成
// ============================================================

// 翌月レポートを自動生成（メニューから呼ぶ）
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
    ui.alert('年月の形式が正しくありません。\n例: 2026/7　または　202607');
    return;
  }
  _runMonthlyReport(year, month);
}

// 毎月25日AM9時に自動実行するトリガーを設定
function setupMonthlyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'generateNextMonthReport')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('generateNextMonthReport')
    .timeBased().onMonthDay(25).atHour(9).create();
  SpreadsheetApp.getUi().alert('✅ 毎月25日 AM9時に翌月レポートを自動生成するよう設定しました。');
}

// スケジュール実行用（トリガーから呼ばれる）
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

    // 期限切れ・離客は対象外
    if (_checkExpired(row, r, backgrounds)) continue;
    const tVal = String(row[COL_T] || '').trim();
    if (tVal === '×' || tVal === '✕') continue;

    // チケット情報を取得（既存ロジックを再利用）
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

    // 最終来店日
    const lastVisitRaw = row[COL_LAST_VISIT];
    const lastVisit    = lastVisitRaw instanceof Date ? lastVisitRaw : null;

    // 更新予定日 = 最終来店日 + 残り回数 × 来店周期
    let renewalDate = null;
    if (lastVisit) {
      renewalDate = new Date(lastVisit);
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
      id:          row[COL_CUSTOMER_ID] || '',
      name:        name,
      type:        couponType,
      lastVisit:   lastVisit,
      remaining:   remaining,
      renewalDate: renewalDate,
      status:      status,
    });
  }

  const monthStr = year + '年' + month + '月';
  const sheetId  = String(year) + String(month).padStart(2, '0');

  _createListSheet(ss, customers, monthStr, sheetId);
  _createSalesSheet(ss, customers, monthStr, sheetId);

  ss.setActiveSheet(ss.getSheetByName('更新リスト_' + sheetId));
  SpreadsheetApp.getUi().alert(
    '✅ ' + monthStr + 'のレポートを生成しました！\n' +
    '・更新リスト_' + sheetId + '\n' +
    '・売上目標_'   + sheetId
  );
}

// 更新見込みリストシートを生成
function _createListSheet(ss, customers, monthStr, sheetId) {
  const sheetName = '更新リスト_' + sheetId;
  const old = ss.getSheetByName(sheetName);
  if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet(sheetName);

  let row = 1;

  // タイトル
  sh.getRange(row, 1, 1, 7).merge()
    .setValue(monthStr + ' 回数券更新見込み客リスト（来店周期' + VISIT_CYCLE + '日ベース）')
    .setFontSize(13).setFontWeight('bold').setFontColor('#1F4E79')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(row, 32); row++;

  sh.getRange(row, 1, 1, 7).merge()
    .setValue('更新予定日 = 最終来店日 + 残り回数 × ' + VISIT_CYCLE + '日　／　赤＝期限超過（要フォロー）')
    .setFontSize(9).setFontColor('#595959');
  sh.setRowHeight(row, 16); row++;

  const sections = [
    { title: '▼ 当月更新見込み（残0含む）',     filter: c => c.status.includes('当月') },
    { title: '▼ 要確認：更新期限超過（前月以前）', filter: c => c.status.includes('要確認') },
    { title: '▼ 翌月以降',                      filter: c => c.status === '翌月以降' },
    { title: '▼ 算出不可（都度来院）',            filter: c => c.status.includes('算出不可') },
  ];

  for (const sec of sections) {
    const group = customers.filter(sec.filter);
    if (group.length === 0) continue;

    row++;
    sh.getRange(row, 1, 1, 7).merge()
      .setValue(sec.title + '　（' + group.length + '名）')
      .setBackground('#2E75B6').setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(10).setVerticalAlignment('middle');
    sh.setRowHeight(row, 22); row++;

    // ヘッダー
    sh.getRange(row, 1, 1, 7)
      .setValues([['顧客番号','氏名','券種','最終来店日','残り回数','更新予定日','ステータス']])
      .setBackground('#1F4E79').setFontColor('#FFFFFF')
      .setFontWeight('bold').setHorizontalAlignment('center');
    sh.setRowHeight(row, 20); row++;

    const couponOrder = ['12回券','8回券','4回券','3回券'];
    for (const ct of couponOrder) {
      const typeGroup = group.filter(c => c.type === ct)
        .sort((a, b) => (a.renewalDate || new Date(9999,0)) - (b.renewalDate || new Date(9999,0)));
      if (typeGroup.length === 0) continue;

      // 券種サブ見出し
      sh.getRange(row, 1, 1, 7).merge()
        .setValue('  【' + ct + '】　' + typeGroup.length + '名')
        .setBackground(TYPE_COLORS[ct] || '#EEEEEE')
        .setFontWeight('bold').setFontSize(10);
      sh.setRowHeight(row, 18); row++;

      for (const c of typeGroup) {
        const isOverdue = c.status.includes('要確認');
        const isUrgent  = c.status.includes('残0') && !isOverdue;
        const bg = isOverdue ? '#FF0000'
                 : isUrgent  ? '#FFC7CE'
                 : c.status.includes('当月') ? TYPE_COLORS[ct] : null;
        const fc = isOverdue ? '#FFFFFF' : '#000000';

        const lv = c.lastVisit   ? Utilities.formatDate(c.lastVisit,   'Asia/Tokyo', 'yyyy/MM/dd') : '—';
        const rv = c.renewalDate ? Utilities.formatDate(c.renewalDate, 'Asia/Tokyo', 'yyyy/MM/dd') : '—';
        const range = sh.getRange(row, 1, 1, 7);
        range.setValues([[c.id, c.name, c.type, lv, c.remaining, rv, c.status]]);
        if (bg) range.setBackground(bg).setFontColor(fc);
        sh.setRowHeight(row, 17); row++;
      }
    }
  }

  [10,18,8,12,8,12,22].forEach((w, i) => sh.setColumnWidth(i + 1, w * 7));
  sh.setFrozenRows(1);
}

// 売上目標シートを生成
function _createSalesSheet(ss, customers, monthStr, sheetId) {
  const sheetName = '売上目標_' + sheetId;
  const old = ss.getSheetByName(sheetName);
  if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet(sheetName);

  const couponTypes = ['12回券','8回券','4回券','3回券'];
  const mayCounts   = {};
  couponTypes.forEach(ct => {
    mayCounts[ct] = customers.filter(c => c.type === ct && c.status.includes('当月')).length;
  });

  let r = 1;

  sh.getRange(r,1,1,5).merge()
    .setValue('売上目標設定シート　／　' + monthStr)
    .setFontSize(15).setFontWeight('bold').setFontColor('#1F4E79').setHorizontalAlignment('center');
  sh.setRowHeight(r, 36); r++;

  sh.getRange(r,1,1,5).merge()
    .setValue('※ 黄色セルに数値を入力すると合計が自動計算されます')
    .setFontSize(9).setFontColor('#C00000').setFontStyle('italic');
  sh.setRowHeight(r, 16); r++;
  r++;

  // 列ヘッダー
  sh.getRange(r,1,1,5).setValues([['項目','件数（変更可）','単価（円）【要入力】','小計（円）','備考']])
    .setBackground('#1F4E79').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setWrap(true);
  sh.setRowHeight(r, 34); r++;

  // 既存顧客セクション
  sh.getRange(r,1,1,5).merge()
    .setValue('■ 既存顧客　回数券更新見込み（' + monthStr + '）')
    .setBackground('#2E75B6').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(11);
  sh.setRowHeight(r, 22); r++;

  const ticketStart = r;
  for (const ct of couponTypes) {
    sh.getRange(r,1).setValue(ct + ' 更新見込み').setBackground(TYPE_COLORS[ct]).setFontWeight('bold');
    sh.getRange(r,2).setValue(mayCounts[ct])
      .setBackground('#FFFFC0').setFontWeight('bold').setHorizontalAlignment('center').setNumberFormat('0');
    sh.getRange(r,3).setValue(0)
      .setBackground('#FFFFC0').setFontColor('#C00000').setFontWeight('bold')
      .setHorizontalAlignment('center').setNumberFormat('#,##0');
    sh.getRange(r,4).setFormula('=B'+r+'*C'+r)
      .setBackground('#F2F2F2').setFontWeight('bold').setHorizontalAlignment('right').setNumberFormat('#,##0');
    sh.getRange(r,5).setValue(ct.replace('回券','') + '枚綴り');
    sh.setRowHeight(r, 22); r++;
  }
  const ticketEnd = r - 1;

  // 既存小計
  sh.getRange(r,1,1,3).merge()
    .setValue('既存顧客　小計').setBackground('#DEEAF1').setFontWeight('bold')
    .setFontColor('#1F4E79').setHorizontalAlignment('right').setFontSize(11);
  sh.getRange(r,4).setFormula('=SUM(D'+ticketStart+':D'+ticketEnd+')')
    .setBackground('#DEEAF1').setFontWeight('bold').setFontColor('#1F4E79')
    .setHorizontalAlignment('right').setNumberFormat('#,##0').setFontSize(11);
  sh.setRowHeight(r, 24);
  const existSubRow = r; r++;

  // 新規顧客セクション
  r++;
  sh.getRange(r,1,1,5).merge()
    .setValue('■ 新規顧客')
    .setBackground('#2E75B6').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(11);
  sh.setRowHeight(r, 22); r++;

  sh.getRange(r,1).setValue('新規顧客数').setBackground('#FFF2CC').setFontWeight('bold');
  sh.getRange(r,2).setValue(0)
    .setBackground('#FFFFC0').setFontColor('#C00000').setFontWeight('bold')
    .setHorizontalAlignment('center').setNumberFormat('0');
  sh.getRange(r,3).setValue(0)
    .setBackground('#FFFFC0').setFontColor('#C00000').setFontWeight('bold')
    .setHorizontalAlignment('center').setNumberFormat('#,##0');
  sh.getRange(r,4).setFormula('=B'+r+'*C'+r)
    .setBackground('#F2F2F2').setFontWeight('bold').setHorizontalAlignment('right').setNumberFormat('#,##0');
  sh.getRange(r,5).setValue('初回チケット平均単価');
  sh.setRowHeight(r, 22);
  const newRow = r; r++;

  sh.getRange(r,1,1,3).merge()
    .setValue('新規顧客　小計').setBackground('#E2EFDA').setFontWeight('bold')
    .setFontColor('#1F4E79').setHorizontalAlignment('right').setFontSize(11);
  sh.getRange(r,4).setFormula('=D'+newRow)
    .setBackground('#E2EFDA').setFontWeight('bold').setFontColor('#1F4E79')
    .setHorizontalAlignment('right').setNumberFormat('#,##0').setFontSize(11);
  sh.setRowHeight(r, 24);
  const newSubRow = r; r++;

  // 合計
  r++;
  sh.getRange(r,1,1,3).merge()
    .setValue('▶ 当月　売上目標（最低見込み）')
    .setBackground('#1F4E79').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(12).setHorizontalAlignment('right');
  sh.getRange(r,4).setFormula('=D'+existSubRow+'+D'+newSubRow)
    .setBackground('#1F4E79').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(14).setHorizontalAlignment('right').setNumberFormat('#,##0');
  sh.setRowHeight(r, 34); r++;

  sh.getRange(r,1,1,5).merge()
    .setValue('※ 都度来院客・アップセル分は含まれません（最低限の見込み額）')
    .setFontSize(9).setFontColor('#595959').setFontStyle('italic');
  sh.setRowHeight(r, 16); r++;

  // サマリー表
  r += 2;
  sh.getRange(r,1,1,5).merge()
    .setValue('■ 件数サマリー（参考）').setFontSize(11).setFontWeight('bold').setFontColor('#1F4E79');
  sh.setRowHeight(r, 22); r++;

  sh.getRange(r,1,1,5).setValues([['券種','当月更新見込み','要確認（期限超過）','翌月以降','算出不可（都度）']])
    .setBackground('#1F4E79').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sh.setRowHeight(r, 20); r++;

  for (const ct of couponTypes) {
    const may  = customers.filter(c => c.type === ct && c.status.includes('当月')).length;
    const over = customers.filter(c => c.type === ct && c.status.includes('要確認')).length;
    const fut  = customers.filter(c => c.type === ct && c.status === '翌月以降').length;
    const noc  = customers.filter(c => c.type === ct && c.status.includes('算出不可')).length;
    sh.getRange(r,1,1,5).setValues([[ct, may, over, fut, noc]]).setHorizontalAlignment('center');
    sh.getRange(r,1).setBackground(TYPE_COLORS[ct]).setFontWeight('bold');
    sh.getRange(r,2).setBackground('#FFC7CE').setFontWeight('bold');
    sh.setRowHeight(r, 20); r++;
  }

  [22,14,15,14,18].forEach((w,i) => sh.setColumnWidth(i+1, w*7));
}

// ============================================================
//  共通ヘルパー関数（集計・月次レポート両方で使用）
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
