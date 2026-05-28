// ============================================================
//  回数券 月次更新見込みレポート 自動生成スクリプト
//  Google Apps Script (スプレッドシートのスクリプトエディタに貼り付けて使用)
// ============================================================

// ─────────────────────────────────────────
//  ★ 設定値（必要に応じて変更してください）
// ─────────────────────────────────────────
const CFG = {
  VISIT_CYCLE   : 10,           // 平均来店周期（日）
  DATA_SHEET    : '顧客データ', // 元データシート名
  COUPON_TYPES  : ['12回券', '8回券', '4回券', '3回券'],
  // 券種ごとの背景色
  TYPE_COLORS   : {
    '12回券': '#BDD7EE',
    '8回券' : '#C6EFCE',
    '4回券' : '#FFEB9C',
    '3回券' : '#FCE4D6',
  },
};

// ─────────────────────────────────────────
//  メニュー追加（スプレッドシートを開いたとき）
// ─────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 月次レポート')
    .addItem('▶ 翌月レポートを今すぐ生成', 'generateNextMonthReport')
    .addSeparator()
    .addItem('⚙ 対象月を手動指定して生成', 'generateReportWithPrompt')
    .addSeparator()
    .addItem('📋 顧客データシートを初期化（初回のみ）', 'initDataSheet')
    .addToUi();
}

// ─────────────────────────────────────────
//  翌月レポート自動生成（メインエントリ）
// ─────────────────────────────────────────
function generateNextMonthReport() {
  const today    = new Date();
  const target   = new Date(today.getFullYear(), today.getMonth() + 1, 1); // 翌月1日
  _runGenerate(target.getFullYear(), target.getMonth() + 1); // getMonth は 0始まりなので +1
}

// ─────────────────────────────────────────
//  手動で年月を指定して生成
// ─────────────────────────────────────────
function generateReportWithPrompt() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.prompt('対象月を入力', '例: 2026/7  または  202607', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const text = res.getResponseText().replace(/[\/\-\s]/g, '');
  const year  = parseInt(text.slice(0, 4));
  const month = parseInt(text.slice(4, 6));
  if (isNaN(year) || isNaN(month) || month < 1 || month > 12) {
    ui.alert('年月の形式が正しくありません。例: 2026/7');
    return;
  }
  _runGenerate(year, month);
}

// ─────────────────────────────────────────
//  月次スケジュール実行用（トリガー設定で使用）
//  「毎月末日」に自動実行するトリガーを設定してください
// ─────────────────────────────────────────
function scheduledMonthlyRun() {
  generateNextMonthReport();
}

// ─────────────────────────────────────────
//  コアロジック
// ─────────────────────────────────────────
function _runGenerate(year, month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── 1. データ読み込み ──
  const dataSheet = ss.getSheetByName(CFG.DATA_SHEET);
  if (!dataSheet) {
    SpreadsheetApp.getUi().alert('「' + CFG.DATA_SHEET + '」シートが見つかりません。\n先に「顧客データシートを初期化」を実行してください。');
    return;
  }
  const raw     = dataSheet.getDataRange().getValues();
  const headers = raw[0];

  // カラム位置を名前で取得
  const col = {};
  ['顧客番号','氏名','券種','最終来店日','残り回数','備考'].forEach(h => {
    col[h] = headers.indexOf(h);
  });

  // ── 2. 顧客データ解析 ──
  const targetStart = new Date(year, month - 1, 1);
  const targetEnd   = new Date(year, month, 0);   // 月末日

  const customers = [];
  for (let i = 1; i < raw.length; i++) {
    const r = raw[i];
    const name = r[col['氏名']];
    if (!name || name === '') continue;

    const lastVisit  = r[col['最終来店日']] instanceof Date ? r[col['最終来店日']] : null;
    const remaining  = parseInt(r[col['残り回数']]) || 0;
    const couponType = String(r[col['券種']] || '');

    // 更新予定日計算
    let renewalDate = null;
    if (lastVisit) {
      renewalDate = new Date(lastVisit);
      renewalDate.setDate(renewalDate.getDate() + remaining * CFG.VISIT_CYCLE);
    }

    // ステータス分類
    let status;
    if (!renewalDate) {
      status = '算出不可（都度）';
    } else if (renewalDate < targetStart) {
      status = remaining === 0 ? '要確認（残0・期限超過）' : '要確認（更新期限超過）';
    } else if (renewalDate <= targetEnd) {
      status = remaining === 0 ? '当月更新必要（残0）' : '当月更新見込み';
    } else {
      status = '翌月以降';
    }

    customers.push({
      id         : r[col['顧客番号']],
      name       : name,
      type       : couponType,
      lastVisit  : lastVisit,
      remaining  : remaining,
      renewalDate: renewalDate,
      status     : status,
      notes      : r[col['備考']] || '',
    });
  }

  // ── 3. シート生成 ──
  const monthStr = year + '年' + month + '月';
  const sheetId  = String(year) + String(month).padStart(2, '0');

  _createListSheet(ss, customers, year, month, monthStr, sheetId);
  _createSalesSheet(ss, customers, year, month, monthStr, sheetId);

  // シートをアクティブにする
  ss.setActiveSheet(ss.getSheetByName('更新リスト_' + sheetId));

  SpreadsheetApp.getUi().alert('✅ ' + monthStr + 'のレポートを生成しました！\n・更新リスト_' + sheetId + '\n・売上目標_' + sheetId);
}

// ─────────────────────────────────────────
//  更新見込みリストシート生成
// ─────────────────────────────────────────
function _createListSheet(ss, customers, year, month, monthStr, sheetId) {
  const name = '更新リスト_' + sheetId;
  const old  = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet(name);

  const targetStart = new Date(year, month - 1, 1);
  const targetEnd   = new Date(year, month, 0);

  let row = 1;

  // タイトル
  sh.getRange(row, 1, 1, 7).merge()
    .setValue(monthStr + ' 回数券更新見込み客リスト（来店周期' + CFG.VISIT_CYCLE + '日ベース）')
    .setFontSize(14).setFontWeight('bold').setFontColor('#1F4E79')
    .setVerticalAlignment('middle').setHorizontalAlignment('center')
    .setRowHeight && sh.setRowHeight(row, 36);
  sh.setRowHeight(row, 36);
  row++;

  sh.getRange(row, 1, 1, 7).merge()
    .setValue('更新予定日 = 最終来店日 + 残り回数 × ' + CFG.VISIT_CYCLE + '日 ／ 対象: ' + monthStr + '1日〜月末')
    .setFontSize(9).setFontColor('#595959');
  sh.setRowHeight(row, 16);
  row++;

  // 各セクション
  const sections = [
    { title: '▼ 当月更新見込み（残0含む）', filter: c => c.status.includes('当月') },
    { title: '▼ 要確認：更新期限超過（前月以前）', filter: c => c.status.includes('要確認') },
    { title: '▼ 翌月以降', filter: c => c.status === '翌月以降' },
  ];

  for (const section of sections) {
    const sectionCustomers = customers.filter(section.filter);
    if (sectionCustomers.length === 0) continue;

    // セクション見出し
    row++;
    sh.getRange(row, 1, 1, 7).merge()
      .setValue(section.title + '　（' + sectionCustomers.length + '名）')
      .setBackground('#2E75B6').setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(10)
      .setVerticalAlignment('middle');
    sh.setRowHeight(row, 22);
    row++;

    // ヘッダー行
    const hdr = ['顧客番号', '氏名', '券種', '最終来店日', '残り回数', '更新予定日', 'ステータス'];
    sh.getRange(row, 1, 1, hdr.length)
      .setValues([hdr])
      .setBackground('#1F4E79').setFontColor('#FFFFFF')
      .setFontWeight('bold').setHorizontalAlignment('center');
    sh.setRowHeight(row, 20);
    row++;

    for (const ct of CFG.COUPON_TYPES) {
      const group = sectionCustomers
        .filter(c => c.type === ct)
        .sort((a, b) => (a.renewalDate || new Date(9999,0,1)) - (b.renewalDate || new Date(9999,0,1)));
      if (group.length === 0) continue;

      // 券種サブ見出し
      sh.getRange(row, 1, 1, 7).merge()
        .setValue('  【' + ct + '】　' + group.length + '名')
        .setBackground(CFG.TYPE_COLORS[ct] || '#EEEEEE')
        .setFontWeight('bold').setFontSize(10);
      sh.setRowHeight(row, 18);
      row++;

      for (const c of group) {
        const isOverdue  = c.status.includes('要確認');
        const isUrgent   = c.status.includes('残0');
        const isMay      = c.status.includes('当月');
        const bg = isOverdue ? '#FF0000'
                 : isUrgent  ? '#FFC7CE'
                 : isMay     ? CFG.TYPE_COLORS[ct] || '#FFFFFF'
                 : null;
        const fc = isOverdue ? '#FFFFFF' : '#000000';

        const lv  = c.lastVisit   ? Utilities.formatDate(c.lastVisit,   'Asia/Tokyo', 'yyyy/MM/dd') : '—';
        const rv  = c.renewalDate ? Utilities.formatDate(c.renewalDate, 'Asia/Tokyo', 'yyyy/MM/dd') : '—';
        const rowData = [[c.id, c.name, c.type, lv, c.remaining, rv, c.status]];
        const range = sh.getRange(row, 1, 1, 7);
        range.setValues(rowData);
        if (bg) range.setBackground(bg).setFontColor(fc);
        sh.setRowHeight(row, 17);
        row++;
      }
    }
  }

  // 列幅調整
  [10, 18, 8, 12, 8, 12, 22].forEach((w, i) => sh.setColumnWidth(i + 1, w * 7));
  sh.setFrozenRows(1);
}

// ─────────────────────────────────────────
//  売上目標シート生成
// ─────────────────────────────────────────
function _createSalesSheet(ss, customers, year, month, monthStr, sheetId) {
  const name = '売上目標_' + sheetId;
  const old  = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet(name);

  // 当月更新見込み件数カウント
  const mayCounts = {};
  for (const ct of CFG.COUPON_TYPES) {
    mayCounts[ct] = customers.filter(c => c.type === ct && c.status.includes('当月')).length;
  }

  let r = 1;

  // ── タイトル
  sh.getRange(r, 1, 1, 5).merge()
    .setValue('売上目標設定シート　／　' + monthStr)
    .setFontSize(16).setFontWeight('bold').setFontColor('#1F4E79')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(r, 38); r++;

  sh.getRange(r, 1, 1, 5).merge()
    .setValue('※ 黄色セルに数値を入力すると合計が自動計算されます')
    .setFontSize(9).setFontColor('#C00000').setFontStyle('italic');
  sh.setRowHeight(r, 16); r++;

  // ── 列ヘッダー
  r++;
  const colHdr = ['項目', '件数\n（変更可）', '単価（円）\n【要入力】', '小計（円）', '備考'];
  sh.getRange(r, 1, 1, 5).setValues([colHdr])
    .setBackground('#1F4E79').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setWrap(true);
  sh.setRowHeight(r, 36); r++;

  // ── 既存顧客セクション
  sh.getRange(r, 1, 1, 5).merge()
    .setValue('■ 既存顧客　回数券更新見込み（' + monthStr + '）')
    .setBackground('#2E75B6').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(11);
  sh.setRowHeight(r, 22); r++;

  const ticketStartRow = r;
  const ticketRows = {};
  for (const ct of CFG.COUPON_TYPES) {
    ticketRows[ct] = r;
    const bg = CFG.TYPE_COLORS[ct] || '#EEEEEE';
    sh.getRange(r, 1).setValue(ct + ' 更新見込み').setBackground(bg).setFontWeight('bold');
    // 件数（黄・変更可）
    sh.getRange(r, 2).setValue(mayCounts[ct])
      .setBackground('#FFFFC0').setFontWeight('bold').setHorizontalAlignment('center')
      .setNumberFormat('0');
    // 単価（黄・要入力）
    sh.getRange(r, 3).setValue(0)
      .setBackground('#FFFFC0').setFontColor('#C00000').setFontWeight('bold')
      .setHorizontalAlignment('center').setNumberFormat('#,##0');
    // 小計（自動計算）
    sh.getRange(r, 4).setFormula('=B' + r + '*C' + r)
      .setBackground('#F2F2F2').setFontWeight('bold')
      .setHorizontalAlignment('right').setNumberFormat('#,##0');
    sh.getRange(r, 5).setValue(ct.replace('回券', '') + '枚綴り');
    sh.setRowHeight(r, 22); r++;
  }

  const ticketEndRow = r - 1;

  // 既存小計
  sh.getRange(r, 1, 1, 3).merge()
    .setValue('既存顧客　小計').setBackground('#DEEAF1')
    .setFontWeight('bold').setFontColor('#1F4E79').setHorizontalAlignment('right').setFontSize(11);
  sh.getRange(r, 4).setFormula('=SUM(D' + ticketStartRow + ':D' + ticketEndRow + ')')
    .setBackground('#DEEAF1').setFontWeight('bold').setFontColor('#1F4E79')
    .setHorizontalAlignment('right').setNumberFormat('#,##0').setFontSize(11);
  sh.setRowHeight(r, 24);
  const existingSubRow = r; r++;

  // ── 新規顧客セクション
  r++;
  sh.getRange(r, 1, 1, 5).merge()
    .setValue('■ 新規顧客')
    .setBackground('#2E75B6').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(11);
  sh.setRowHeight(r, 22); r++;

  sh.getRange(r, 1).setValue('新規顧客数').setBackground('#FFF2CC').setFontWeight('bold');
  sh.getRange(r, 2).setValue(0)
    .setBackground('#FFFFC0').setFontColor('#C00000').setFontWeight('bold')
    .setHorizontalAlignment('center').setNumberFormat('0');
  sh.getRange(r, 3).setValue(0)
    .setBackground('#FFFFC0').setFontColor('#C00000').setFontWeight('bold')
    .setHorizontalAlignment('center').setNumberFormat('#,##0');
  sh.getRange(r, 4).setFormula('=B' + r + '*C' + r)
    .setBackground('#F2F2F2').setFontWeight('bold')
    .setHorizontalAlignment('right').setNumberFormat('#,##0');
  sh.getRange(r, 5).setValue('初回チケット平均単価');
  sh.setRowHeight(r, 22);
  const newRow = r; r++;

  sh.getRange(r, 1, 1, 3).merge()
    .setValue('新規顧客　小計').setBackground('#E2EFDA')
    .setFontWeight('bold').setFontColor('#1F4E79').setHorizontalAlignment('right').setFontSize(11);
  sh.getRange(r, 4).setFormula('=D' + newRow)
    .setBackground('#E2EFDA').setFontWeight('bold').setFontColor('#1F4E79')
    .setHorizontalAlignment('right').setNumberFormat('#,##0').setFontSize(11);
  sh.setRowHeight(r, 24);
  const newSubRow = r; r++;

  // ── 合計
  r++;
  sh.getRange(r, 1, 1, 3).merge()
    .setValue('▶ 当月　売上目標（最低見込み）')
    .setBackground('#1F4E79').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(12).setHorizontalAlignment('right');
  sh.getRange(r, 4).setFormula('=D' + existingSubRow + '+D' + newSubRow)
    .setBackground('#1F4E79').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(14).setHorizontalAlignment('right').setNumberFormat('#,##0');
  sh.setRowHeight(r, 34); r++;

  sh.getRange(r, 1, 1, 5).merge()
    .setValue('※ この数値は最低限の見込み売上です。都度来院客・アップセル分は含まれません。')
    .setFontSize(9).setFontColor('#595959').setFontStyle('italic');
  sh.setRowHeight(r, 16); r++;

  // ── 件数サマリー表
  r += 2;
  sh.getRange(r, 1, 1, 5).merge()
    .setValue('■ 件数サマリー（参考）')
    .setFontSize(12).setFontWeight('bold').setFontColor('#1F4E79');
  sh.setRowHeight(r, 22); r++;

  const smHdr = ['券種', '当月更新見込み', '要確認（期限超過）', '翌月以降', '都度（算出不可）'];
  sh.getRange(r, 1, 1, 5).setValues([smHdr])
    .setBackground('#1F4E79').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sh.setRowHeight(r, 20); r++;

  for (const ct of CFG.COUPON_TYPES) {
    const may  = customers.filter(c => c.type === ct && c.status.includes('当月')).length;
    const over = customers.filter(c => c.type === ct && c.status.includes('要確認')).length;
    const fut  = customers.filter(c => c.type === ct && c.status === '翌月以降').length;
    const noc  = customers.filter(c => c.type === ct && c.status.includes('算出不可')).length;
    sh.getRange(r, 1, 1, 5).setValues([[ct, may, over, fut, noc]])
      .setHorizontalAlignment('center');
    sh.getRange(r, 1).setBackground(CFG.TYPE_COLORS[ct] || '#EEEEEE').setFontWeight('bold');
    sh.getRange(r, 2).setBackground('#FFC7CE').setFontWeight('bold'); // 当月強調
    sh.setRowHeight(r, 20); r++;
  }

  // ── 入力ガイド
  r += 2;
  sh.getRange(r, 1, 1, 5).merge()
    .setValue('■ 入力手順').setFontSize(12).setFontWeight('bold').setFontColor('#1F4E79');
  sh.setRowHeight(r, 22); r++;

  const guides = [
    '① C列（単価）を各回数券の金額に入力 → 既存顧客の小計が自動計算',
    '② B列（件数）は実態に合わせて上書き可能',
    '③ 新規顧客数・初回単価を入力 → 合計が自動更新',
    '④「顧客データ」シートの最終来店日・残り回数を更新するたびに\n　 メニュー「翌月レポートを今すぐ生成」を実行すると最新データで再生成されます',
  ];
  for (const g of guides) {
    sh.getRange(r, 1, 1, 5).merge().setValue(g).setFontSize(10).setWrap(true);
    sh.setRowHeight(r, g.includes('\n') ? 36 : 20); r++;
  }

  // 列幅
  [22, 14, 15, 14, 20].forEach((w, i) => sh.setColumnWidth(i + 1, w * 7));
}

// ─────────────────────────────────────────
//  顧客データシート初期化（初回セットアップ用）
// ─────────────────────────────────────────
function initDataSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  if (ss.getSheetByName(CFG.DATA_SHEET)) {
    const res = ui.alert('「' + CFG.DATA_SHEET + '」シートは既に存在します。\n初期化すると内容が消えますがよろしいですか？', ui.ButtonSet.YES_NO);
    if (res !== ui.Button.YES) return;
    ss.deleteSheet(ss.getSheetByName(CFG.DATA_SHEET));
  }

  const sh = ss.insertSheet(CFG.DATA_SHEET, 0);
  const headers = ['顧客番号', '氏名', '券種', '最終来店日', '残り回数', '備考'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#1F4E79').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setFontSize(10);
  sh.setFrozenRows(1);

  // 列幅
  [10, 18, 8, 13, 10, 20].forEach((w, i) => sh.setColumnWidth(i + 1, w * 7));

  // 「最終来店日」列を日付フォーマットに設定
  sh.getRange(2, 4, 500, 1).setNumberFormat('yyyy/mm/dd');

  // サンプルデータ（3行）
  const sample = [
    [2348, 'コダマ　ショウ', '12回券', new Date(2026, 4, 24), 4, '周期B'],
    [2228, 'イトウ　トシアキ', '8回券', new Date(2026, 4, 18), 1, ''],
    [2291, 'ササキ　クミ', '4回券', new Date(2026, 4, 17), 1, ''],
  ];
  sh.getRange(2, 1, sample.length, headers.length).setValues(sample);

  ui.alert('✅「' + CFG.DATA_SHEET + '」シートを初期化しました。\nサンプルデータ3行を追加しています。実際の顧客データに差し替えてください。\n\n【データの貼り付け方】\n・顧客番号・氏名・券種・最終来店日・残り回数 の5列を入力\n・「都度」来院の方は「最終来店日」を空欄にしてください');
}

// ─────────────────────────────────────────
//  月次自動実行トリガー設定（初回のみ手動実行）
// ─────────────────────────────────────────
function setupMonthlyTrigger() {
  // 既存トリガーを削除
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'scheduledMonthlyRun') ScriptApp.deleteTrigger(t);
  });
  // 毎月25日 AM9:00 に自動実行（翌月分を月末前に生成）
  ScriptApp.newTrigger('scheduledMonthlyRun')
    .timeBased()
    .onMonthDay(25)
    .atHour(9)
    .create();
  SpreadsheetApp.getUi().alert('✅ 毎月25日 AM9時に自動生成するトリガーを設定しました。');
}
