function createCalendarSheetsForYearMonth(year, month) {
  const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);

  const yyyyMM = `${year}${month.toString().padStart(2, "0")}`;
  const bSheetName = `b_calendar_${yyyyMM}`;
  const dSheetName = `d_calendar_${yyyyMM}`;

  // 既存のシートがあれば削除
  const existingBSheet = spreadsheet.getSheetByName(bSheetName);
  const existingDSheet = spreadsheet.getSheetByName(dSheetName);

  if (existingBSheet) {
    spreadsheet.deleteSheet(existingBSheet);
  }
  if (existingDSheet) {
    spreadsheet.deleteSheet(existingDSheet);
  }

  // 朝食カレンダーシートの作成
  const bSheet = spreadsheet.insertSheet(bSheetName);
  bSheet.getRange("A1:C1").setValues([["b_calendar_id", "date", "b_menu_id"]]);
  bSheet.getRange("A1:C1").setFontWeight("bold");

  // 夕食カレンダーシートの作成
  const dSheet = spreadsheet.insertSheet(dSheetName);
  dSheet.getRange("A1:C1").setValues([["d_calendar_id", "date", "d_menu_id"]]);
  dSheet.getRange("A1:C1").setFontWeight("bold");

  // 月の最終日を取得
  const lastDay = new Date(year, month, 0).getDate();

  let bCalendarIdCounter = 1;
  let dCalendarIdCounter = 1;

  // 全ての日のデータを生成
  for (let day = 1; day <= lastDay; day++) {
    const date = new Date(year, month - 1, day);
    const dayOfWeek = date.getDay(); // 0=日曜日, 6=土曜日

    // 日曜日を除外（朝食、夕食共通）
    if (dayOfWeek === 0) {
      continue;
    }

    // 朝食カレンダーに行を追加（日曜日以外）
    bSheet.appendRow([bCalendarIdCounter, date, 0]);
    bCalendarIdCounter++;

    // 夕食カレンダーに行を追加（日曜日と土曜日以外）
    if (dayOfWeek !== 6) {
      // 土曜日は夕食を除外
      dSheet.appendRow([dCalendarIdCounter, date, 0]);
      dCalendarIdCounter++;
    }
  }

  // 日付の形式を設定
  if (bCalendarIdCounter > 1) {
    const dateRange1 = bSheet.getRange(2, 2, bCalendarIdCounter - 1, 1);
    dateRange1.setNumberFormat("yyyy-MM-dd");
  }

  if (dCalendarIdCounter > 1) {
    const dateRange2 = dSheet.getRange(2, 2, dCalendarIdCounter - 1, 1);
    dateRange2.setNumberFormat("yyyy-MM-dd");
  }

  // 列の幅を自動調整
  bSheet.autoResizeColumns(1, 3);
  dSheet.autoResizeColumns(1, 3);

  return {
    success: true,
    message: `${year}年${month}月のカレンダーシート ${bSheetName} と ${dSheetName} を作成しました。
朝食: ${bCalendarIdCounter - 1}日分（日曜日を除く）
夕食: ${dCalendarIdCounter - 1}日分（土曜日と日曜日を除く）`,
    sheets: [bSheetName, dSheetName],
  };
}
