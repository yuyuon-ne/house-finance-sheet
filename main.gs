function onOpen() {
  SpreadsheetApp.getUi().createMenu('🏠 家計簿メニュー').addItem('入力フォームを開く', 'showSidebar').addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('index').setTitle('家計簿入力').setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

function addRows(data) {
  if (!data) return;
  try {
    const ss = SpreadsheetApp.getActive();
    const ledgerSheet = ss.getSheetByName(data.sheetName);
    if (!ledgerSheet) throw new Error("シートが見つかりません");

    const dateObj = new Date(data.date);
    const category = data.category || "";
    
    // 台帳へ書き込み
    ledgerSheet.appendRow([dateObj, category, (data.type === 'income' ? '収入' : '支出'), Number(data.amount), data.itemName, data.shopName || ""]);
    
    // 「給与」の場合の連動処理
    if (category.indexOf('給与') !== -1) {
      const tz = ss.getSpreadsheetTimeZone();
      // 入力日の翌月を「対象月」とする（例：1/28入力なら2026/02分）
      const targetMonthDate = new Date(dateObj.getFullYear(), dateObj.getMonth() + 1, 1);
      const targetMonthStr = Utilities.formatDate(targetMonthDate, tz, "yyyy/MM");
      
      updateSalaryList(ss, dateObj, targetMonthStr, tz);
      
      // 分析用シートのG7を更新
      const analysisSheet = ss.getSheetByName('分析用');
      if (analysisSheet) {
        analysisSheet.getRange('G7').setValue(targetMonthStr);
      }
    }
    return "success";
  } catch (e) { throw new Error(e.message); }
}

function updateSalaryList(ss, salaryDate, targetMonthStr, tz) {
  const listSheet = ss.getSheetByName('給与日リスト');
  if (!listSheet) return;

  const data = listSheet.getDataRange().getValues();
  let targetRow = -1;
  let prevRow = -1;

  // 前月（2026/01）の特定用
  const prevMonthDate = new Date(salaryDate.getFullYear(), salaryDate.getMonth(), 1);
  const prevMonthStr = Utilities.formatDate(prevMonthDate, tz, "yyyy/MM");

  for (let i = 1; i < data.length; i++) {
    let m = (data[i][0] instanceof Date) ? Utilities.formatDate(data[i][0], tz, "yyyy/MM") : String(data[i][0]);
    if (m === targetMonthStr) targetRow = i + 1;
    if (m === prevMonthStr) prevRow = i + 1;
  }

  // 今回の終了日の計算（給与日の1ヶ月後の前日：例 1/28 → 2/27）
  let nextEndDay = new Date(salaryDate.getFullYear(), salaryDate.getMonth() + 1, salaryDate.getDate() - 1);

  if (targetRow !== -1) {
    // 既存の2026/02行があれば更新
    listSheet.getRange(targetRow, 2).setValue(salaryDate);
    listSheet.getRange(targetRow, 3).setValue(nextEndDay);
  } else {
    // なければ新しく追加
    listSheet.appendRow([targetMonthStr, salaryDate, nextEndDay]); 
  }

  // 前月（2026/01）の終了日を今回の給与日の前日（1/27）で確定させる
  if (prevRow !== -1) {
    let lastMonthEnd = new Date(salaryDate);
    lastMonthEnd.setDate(lastMonthEnd.getDate() - 1);
    listSheet.getRange(prevRow, 3).setValue(lastMonthEnd);
  }
  
  // セルの書式を日付形式に統一
  listSheet.getRange("B2:C100").setNumberFormat('yyyy/MM/dd');
}
