function getUserReservationCalendar(userId, year, month) {
  const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
  const ss = SpreadsheetApp.openById(spreadsheetId);

  const yyyyMM = `${year}${month.toString().padStart(2, "0")}`;
  const bCalendarSheetName = `b_calendar_${yyyyMM}`;
  const dCalendarSheetName = `d_calendar_${yyyyMM}`;
  const bReservationSheetName = `b_reservations_${yyyyMM}`;
  const dReservationSheetName = `d_reservations_${yyyyMM}`;

  const bCalendarSheet = ss.getSheetByName(bCalendarSheetName);
  const dCalendarSheet = ss.getSheetByName(dCalendarSheetName);
  const bReservationSheet = ss.getSheetByName(bReservationSheetName);
  const dReservationSheet = ss.getSheetByName(dReservationSheetName);
  const bMenuSheet = ss.getSheetByName("b_menus");
  const dMenuSheet = ss.getSheetByName("d_menus");

  if (
    !bCalendarSheet ||
    !dCalendarSheet ||
    !bReservationSheet ||
    !dReservationSheet
  ) {
    return {
      success: false,
      message: `必要なシートが見つかりません。カレンダーと予約シートが存在するか確認してください。`,
    };
  }

  // データの取得
  const bCalendarData = bCalendarSheet.getDataRange().getValues();
  const dCalendarData = dCalendarSheet.getDataRange().getValues();
  const bReservationData = bReservationSheet.getDataRange().getValues();
  const dReservationData = dReservationSheet.getDataRange().getValues();

  // メニューデータの取得（存在する場合）
  let bMenuData = [];
  let dMenuData = [];

  if (bMenuSheet) {
    bMenuData = bMenuSheet.getDataRange().getValues();
  }

  if (dMenuSheet) {
    dMenuData = dMenuSheet.getDataRange().getValues();
  }

  // ヘッダー行の列インデックスを取得
  const bCalendarHeaders = bCalendarData[0];
  const dCalendarHeaders = dCalendarData[0];
  const bReservationHeaders = bReservationData[0];
  const dReservationHeaders = dReservationData[0];

  const bCalendarIdIndex = bCalendarHeaders.indexOf("b_calendar_id");
  const bCalendarDateIndex = bCalendarHeaders.indexOf("date");
  const bCalendarMenuIdIndex = bCalendarHeaders.indexOf("b_menu_id");
  const bCalendarIsActiveIndex = bCalendarHeaders.indexOf("is_active");

  const dCalendarIdIndex = dCalendarHeaders.indexOf("d_calendar_id");
  const dCalendarDateIndex = dCalendarHeaders.indexOf("date");
  const dCalendarMenuIdIndex = dCalendarHeaders.indexOf("d_menu_id");
  const dCalendarIsActiveIndex = dCalendarHeaders.indexOf("is_active");

  const bReservationCalendarIdIndex =
    bReservationHeaders.indexOf("b_calendar_id");
  const bReservationUserIdIndex = bReservationHeaders.indexOf("user_id");
  const bReservationStatusIndex = bReservationHeaders.indexOf("is_reserved");

  const dReservationCalendarIdIndex =
    dReservationHeaders.indexOf("d_calendar_id");
  const dReservationUserIdIndex = dReservationHeaders.indexOf("user_id");
  const dReservationStatusIndex = dReservationHeaders.indexOf("is_reserved");

  // 朝食メニューマップの作成
  const bMenuMap = {};
  if (bMenuData.length > 1) {
    const bMenuIdIndex = bMenuData[0].indexOf("b_menu_id");
    const bMenuNameIndex = bMenuData[0].indexOf("breakfast_menu");
    const bMenuCalorieIndex = bMenuData[0].indexOf("calorie");

    if (bMenuIdIndex !== -1 && bMenuNameIndex !== -1) {
      for (let i = 1; i < bMenuData.length; i++) {
        const menuId = bMenuData[i][bMenuIdIndex];
        const menuName = bMenuData[i][bMenuNameIndex];
        const calorie = bMenuCalorieIndex !== -1 ? bMenuData[i][bMenuCalorieIndex] : 0;
        bMenuMap[menuId] = {
          name: menuName,
          calorie: calorie
        };
      }
    }
  }

  // 夕食メニューマップの作成
  const dMenuMap = {};
  if (dMenuData.length > 1) {
    const dMenuIdIndex = dMenuData[0].indexOf("d_menu_id");
    const dMenuNameIndex = dMenuData[0].indexOf("dinner_menu");
    const dMenuCalorieIndex = dMenuData[0].indexOf("calorie");

    if (dMenuIdIndex !== -1 && dMenuNameIndex !== -1) {
      for (let i = 1; i < dMenuData.length; i++) {
        const menuId = dMenuData[i][dMenuIdIndex];
        const menuName = dMenuData[i][dMenuNameIndex];
        const calorie = dMenuCalorieIndex !== -1 ? dMenuData[i][dMenuCalorieIndex] : 0;
        dMenuMap[menuId] = {
          name: menuName,
          calorie: calorie
        };
      }
    }
  }

  // ユーザーの朝食予約を取得
  const userBreakfastReservations = [];
  for (let i = 1; i < bReservationData.length; i++) {
    const row = bReservationData[i];
    if (row[bReservationUserIdIndex] === userId) {
      userBreakfastReservations.push({
        calendarId: row[bReservationCalendarIdIndex],
        isReserved: row[bReservationStatusIndex],
      });
    }
  }

  // ユーザーの夕食予約を取得
  const userDinnerReservations = [];
  for (let i = 1; i < dReservationData.length; i++) {
    const row = dReservationData[i];
    if (row[dReservationUserIdIndex] === userId) {
      userDinnerReservations.push({
        calendarId: row[dReservationCalendarIdIndex],
        isReserved: row[dReservationStatusIndex],
      });
    }
  }

  // 朝食カレンダーマップの作成
  const breakfastCalendar = {};
  for (let i = 1; i < bCalendarData.length; i++) {
    const row = bCalendarData[i];
    const calendarId = row[bCalendarIdIndex];
    const date = row[bCalendarDateIndex];
    const menuId = row[bCalendarMenuIdIndex];
    const isActive = row[bCalendarIsActiveIndex];

    if (date instanceof Date) {
      const dateStr = formatDate(date);
      const menuInfo = bMenuMap[menuId] || { name: "未設定", calorie: 0 };
      breakfastCalendar[calendarId] = {
        id: calendarId,
        date: dateStr,
        menu_id: menuId,
        menu_name: menuInfo.name,
        menu_calorie: menuInfo.calorie,
        is_active: isActive
      };
    }
  }

  // 夕食カレンダーマップの作成
  const dinnerCalendar = {};
  for (let i = 1; i < dCalendarData.length; i++) {
    const row = dCalendarData[i];
    const calendarId = row[dCalendarIdIndex];
    const date = row[dCalendarDateIndex];
    const menuId = row[dCalendarMenuIdIndex];
    const isActive = row[dCalendarIsActiveIndex];

    if (date instanceof Date) {
      const dateStr = formatDate(date);
      const menuInfo = dMenuMap[menuId] || { name: "未設定", calorie: 0 };
      dinnerCalendar[calendarId] = {
        id: calendarId,
        date: dateStr,
        menu_id: menuId,
        menu_name: menuInfo.name,
        menu_calorie: menuInfo.calorie,
        is_active: isActive
      };
    }
  }

  // 日付でグループ化された朝食データを作成
  const groupedBreakfast = {};
  for (const reservation of userBreakfastReservations) {
    const calendarData = breakfastCalendar[reservation.calendarId];
    if (calendarData) {
      const dateStr = calendarData.date;
      groupedBreakfast[dateStr] = {
        ...calendarData,
        is_reserved: reservation.isReserved,
      };
    }
  }

  // 日付でグループ化された夕食データを作成
  const groupedDinner = {};
  for (const reservation of userDinnerReservations) {
    const calendarData = dinnerCalendar[reservation.calendarId];
    if (calendarData) {
      const dateStr = calendarData.date;
      groupedDinner[dateStr] = {
        ...calendarData,
        is_reserved: reservation.isReserved,
      };
    }
  }

  // 月のすべての日付を取得
  const monthDates = [];
  const daysInMonth = new Date(year, month, 0).getDate();
  const firstDayOfMonth = new Date(year, month - 1, 1);
  const offset = firstDayOfMonth.getDay(); // 0 (日曜日) から 6 (土曜日)

  for (let day = 1; day <= daysInMonth; day++) {
    const date = new Date(year, month - 1, day);
    const dateStr = formatDate(date);
    monthDates.push(dateStr);
  }

  // メニューマップをレスポンスに含める
  const menuMap = {
    breakfast: bMenuMap,
    dinner: dMenuMap,
  };

  return {
    success: true,
    userId: userId,
    year: year,
    month: month,
    offset: offset,
    monthDates: monthDates,
    weekdays: ["日", "月", "火", "水", "木", "金", "土"],
    groupedBreakfast: groupedBreakfast,
    groupedDinner: groupedDinner,
    menuMap: menuMap,
  };
}

function updateReservation(
  userId,
  calendarId,
  mealType,
  isReserved,
  year,
  month
) {
  const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
  const ss = SpreadsheetApp.openById(spreadsheetId);

  const yyyyMM = `${year}${month.toString().padStart(2, "0")}`;
  const prefix = mealType === "breakfast" ? "b" : "d";
  const calendarSheetName = `${prefix}_calendar_${yyyyMM}`;
  const reservationSheetName = `${prefix}_reservations_${yyyyMM}`;

  const reservationSheet = ss.getSheetByName(reservationSheetName);
  if (!reservationSheet) {
    return {
      success: false,
      message: `予約シート ${reservationSheetName} が見つかりません。`,
    };
  }

  const reservationData = reservationSheet.getDataRange().getValues();
  const headers = reservationData[0];

  const calendarIdIndex = headers.indexOf(`${prefix}_calendar_id`);
  const userIdIndex = headers.indexOf("user_id");
  const isReservedIndex = headers.indexOf("is_reserved");

  if (calendarIdIndex === -1 || userIdIndex === -1 || isReservedIndex === -1) {
    return {
      success: false,
      message: "必要なカラムが予約シートに見つかりません。",
    };
  }

  let rowIndex = -1;
  for (let i = 1; i < reservationData.length; i++) {
    const row = reservationData[i];
    if (row[calendarIdIndex] == calendarId && row[userIdIndex] == userId) {
      rowIndex = i;
      break;
    }
  }

  if (rowIndex === -1) {
    return {
      success: false,
      message: `ユーザーID ${userId} とカレンダーID ${calendarId} に一致する予約が見つかりません。`,
    };
  }

  reservationSheet
    .getRange(rowIndex + 1, isReservedIndex + 1)
    .setValue(isReserved);

  // 食事原紙シートをリアルタイムで更新
  const calendarSheet = ss.getSheetByName(calendarSheetName);
  if (calendarSheet) {
    const calendarData = calendarSheet.getDataRange().getValues();
    const calendarHeaders = calendarData[0];
    const calIdIndex = calendarHeaders.indexOf(`${prefix}_calendar_id`);
    const calDateIndex = calendarHeaders.indexOf("date");
    if (calIdIndex !== -1 && calDateIndex !== -1) {
      for (let i = 1; i < calendarData.length; i++) {
        if (calendarData[i][calIdIndex] == calendarId) {
          const calDate = calendarData[i][calDateIndex];
          const dateStr = calDate instanceof Date ? formatDate(calDate) : String(calDate);
          updateMealSheetForUser(userId, dateStr, mealType, isReserved);
          break;
        }
      }
    }
  }

  return {
    success: true,
  };
}

function createAllUsersReservations(year, month) { 

  // シート名を生成
  const yyyyMM = `${year}${month.toString().padStart(2, "0")}`;
  const bCalendarSheetName = `b_calendar_${yyyyMM}`;
  const dCalendarSheetName = `d_calendar_${yyyyMM}`;
  const bReservationSheetName = `b_reservations_${yyyyMM}`;
  const dReservationSheetName = `d_reservations_${yyyyMM}`;

  // カレンダーシートの存在確認
  const bCalendarSheet = ss.getSheetByName(bCalendarSheetName);
  const dCalendarSheet = ss.getSheetByName(dCalendarSheetName);

  if (!bCalendarSheet || !dCalendarSheet) {
    return {
      success: false,
      message: `カレンダーシートが存在しません。先にcreateCalendarSheetsForYearMonth関数を実行してください。`,
    };
  }

  // ユーザー一覧を取得
  const usersSheet = ss.getSheetByName("users");
  if (!usersSheet) {
    return {
      success: false,
      message: `ユーザーシートが存在しません。`,
    };
  }

  const usersData = usersSheet.getDataRange().getValues();
  const userIdIndex = usersData[0].indexOf("user_id");

  if (userIdIndex === -1) {
    return {
      success: false,
      message: `ユーザーシートにuser_idカラムが見つかりません。`,
    };
  }

  // ヘッダー行を除外したユーザーIDの配列
  const userIds = [];
  for (let i = 1; i < usersData.length; i++) {
    const userId = usersData[i][userIdIndex];
    if (userId) {
      // 空でないIDのみ追加
      userIds.push(userId);
    }
  }

  // 予約シートの取得または作成
  let bReservationSheet = ss.getSheetByName(bReservationSheetName);
  let dReservationSheet = ss.getSheetByName(dReservationSheetName);

  // 朝食予約シートがなければ作成
  if (!bReservationSheet) {
    bReservationSheet = ss.insertSheet(bReservationSheetName);
    bReservationSheet
      .getRange("A1:D1")
      .setValues([
        ["b_reservation_id", "b_calendar_id", "user_id", "is_reserved"],
      ]);
    bReservationSheet.getRange("A1:D1").setFontWeight("bold");
    bReservationSheet.autoResizeColumns(1, 4);
  }

  // 夕食予約シートがなければ作成
  if (!dReservationSheet) {
    dReservationSheet = ss.insertSheet(dReservationSheetName);
    dReservationSheet
      .getRange("A1:D1")
      .setValues([
        ["d_reservation_id", "d_calendar_id", "user_id", "is_reserved"],
      ]);
    dReservationSheet.getRange("A1:D1").setFontWeight("bold");
    dReservationSheet.autoResizeColumns(1, 4);
  }

  // カレンダーデータを取得
  const bCalendarData = bCalendarSheet.getDataRange().getValues();
  const dCalendarData = dCalendarSheet.getDataRange().getValues();

  // ヘッダー行を除外したカレンダーIDの配列
  const bCalendarIds = [];
  const dCalendarIds = [];

  for (let i = 1; i < bCalendarData.length; i++) {
    bCalendarIds.push(bCalendarData[i][0]);
  }

  for (let i = 1; i < dCalendarData.length; i++) {
    dCalendarIds.push(dCalendarData[i][0]);
  }

  // 既存の予約データを取得
  const bReservationData = bReservationSheet.getDataRange().getValues();
  const dReservationData = dReservationSheet.getDataRange().getValues();

  // 各予約シートの次のIDを決定
  let bReservationIdCounter = 1;
  let dReservationIdCounter = 1;

  if (bReservationData.length > 1) {
    const lastBReservationId = bReservationData[bReservationData.length - 1][0];
    bReservationIdCounter = lastBReservationId + 1;
  }

  if (dReservationData.length > 1) {
    const lastDReservationId = dReservationData[dReservationData.length - 1][0];
    dReservationIdCounter = lastDReservationId + 1;
  }

  // 既存の予約組み合わせ（カレンダーID + ユーザーID）をセットで管理
  const existingBReservations = new Set();
  const existingDReservations = new Set();

  for (let i = 1; i < bReservationData.length; i++) {
    const row = bReservationData[i];
    const key = `${row[1]}_${row[2]}`; // b_calendar_id_user_id
    existingBReservations.add(key);
  }

  for (let i = 1; i < dReservationData.length; i++) {
    const row = dReservationData[i];
    const key = `${row[1]}_${row[2]}`; // d_calendar_id_user_id
    existingDReservations.add(key);
  }

  // 朝食予約データ作成用の配列
  const bReservationsToAdd = [];
  let bRowsAdded = 0;

  // 夕食予約データ作成用の配列
  const dReservationsToAdd = [];
  let dRowsAdded = 0;

  // 全ユーザー×全カレンダーの組み合わせループ
  for (const userId of userIds) {
    // 朝食予約
    for (const bCalendarId of bCalendarIds) {
      const key = `${bCalendarId}_${userId}`;

      // この組み合わせの予約がまだ存在しない場合
      if (!existingBReservations.has(key)) {
        bReservationsToAdd.push([
          bReservationIdCounter,
          bCalendarId,
          userId,
          false, // デフォルトは予約なし
        ]);
        bReservationIdCounter++;
        bRowsAdded++;
      }
    }

    // 夕食予約
    for (const dCalendarId of dCalendarIds) {
      const key = `${dCalendarId}_${userId}`;

      // この組み合わせの予約がまだ存在しない場合
      if (!existingDReservations.has(key)) {
        dReservationsToAdd.push([
          dReservationIdCounter,
          dCalendarId,
          userId,
          false, // デフォルトは予約なし
        ]);
        dReservationIdCounter++;
        dRowsAdded++;
      }
    }
  }

  // 一括で予約データを追加（効率的な書き込み）
  if (bReservationsToAdd.length > 0) {
    const lastRow = bReservationSheet.getLastRow();
    bReservationSheet
      .getRange(lastRow + 1, 1, bReservationsToAdd.length, 4)
      .setValues(bReservationsToAdd);
  }

  if (dReservationsToAdd.length > 0) {
    const lastRow = dReservationSheet.getLastRow();
    dReservationSheet
      .getRange(lastRow + 1, 1, dReservationsToAdd.length, 4)
      .setValues(dReservationsToAdd);
  }

  return {
    success: true,
    message: `${year}年${month}月の予約データを${userIds.length}人分作成しました。朝食: ${bRowsAdded}件, 夕食: ${dRowsAdded}件追加されました。`,
    totalUsers: userIds.length,
    bRowsAdded: bRowsAdded,
    dRowsAdded: dRowsAdded,
  };
}

function bulkUpdateReservationsWithIds(userId, mealType, isReserved, year, month, calendarIds) {
  const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
  const ss = SpreadsheetApp.openById(spreadsheetId);
  
  // シート名を生成
  const yyyyMM = `${year}${month.toString().padStart(2, "0")}`;
  const prefix = mealType === "breakfast" ? "b" : "d";
  const reservationSheetName = `${prefix}_reservations_${yyyyMM}`;
  
  // シートの存在確認
  const reservationSheet = ss.getSheetByName(reservationSheetName);
  if (!reservationSheet) {
    return {
      success: false,
      message: `予約シート ${reservationSheetName} が見つかりません。`
    };
  }
  
  // 予約データを取得
  const reservationData = reservationSheet.getDataRange().getValues();
  const headers = reservationData[0];
  
  // カラムインデックスを取得
  const calendarIdIndex = headers.indexOf(`${prefix}_calendar_id`);
  const userIdIndex = headers.indexOf("user_id");
  const isReservedIndex = headers.indexOf("is_reserved");
  
  if (calendarIdIndex === -1 || userIdIndex === -1 || isReservedIndex === -1) {
    return {
      success: false,
      message: "必要なカラムが予約シートに見つかりません。"
    };
  }
  
  // 更新対象の行を収集
  const rowsToUpdate = [];
  for (let i = 1; i < reservationData.length; i++) {
    const row = reservationData[i];
    // 該当ユーザーかつ指定されたカレンダーIDのみを対象にする
    if (row[userIdIndex] == userId && calendarIds.includes(row[calendarIdIndex])) {
      rowsToUpdate.push(i + 1); // スプレッドシートの行番号は1始まり、ヘッダー行があるので+1
    }
  }
  
  // カレンダーシートから calendarId -> 日付 のマップを作成
  const calendarSheetName = `${prefix}_calendar_${yyyyMM}`;
  const calendarDateMap = {};
  const calendarSheet = ss.getSheetByName(calendarSheetName);
  if (calendarSheet) {
    const calendarData = calendarSheet.getDataRange().getValues();
    const calendarHeaders = calendarData[0];
    const calIdIndex = calendarHeaders.indexOf(`${prefix}_calendar_id`);
    const calDateIndex = calendarHeaders.indexOf("date");
    if (calIdIndex !== -1 && calDateIndex !== -1) {
      for (let i = 1; i < calendarData.length; i++) {
        const calDate = calendarData[i][calDateIndex];
        calendarDateMap[calendarData[i][calIdIndex]] =
          calDate instanceof Date ? formatDate(calDate) : String(calDate);
      }
    }
  }

  // 一括で更新
  let updatedCount = 0;
  for (const rowIndex of rowsToUpdate) {
    reservationSheet.getRange(rowIndex, isReservedIndex + 1).setValue(isReserved);
    updatedCount++;

    // 食事原紙シートをリアルタイムで更新
    const row = reservationData[rowIndex - 1];
    const calId = row[calendarIdIndex];
    const dateStr = calendarDateMap[calId];
    if (dateStr) {
      updateMealSheetForUser(userId, dateStr, mealType, isReserved);
    }
  }
  
  return {
    success: true,
    message: `${updatedCount}件の${mealType === "breakfast" ? "朝食" : "夕食"}予約を一括更新しました。`,
    updatedCount: updatedCount
  };
}

function updateMealSheetForUser(userId, date, mealType, isReserved) {
  const mealSpreadsheetId = "17iuUzC-fx8lfMA8M5HrLwMlzvCpS9TCRcoCDzMrHjE4";
  const mealSs = SpreadsheetApp.openById(mealSpreadsheetId);

  const parts = date.split("-");
  const year = parseInt(parts[0]);
  const month = parseInt(parts[1]);
  const day = parseInt(parts[2]);

  const dateObj = new Date(year, month - 1, day);
  const dayOfWeek = dateObj.getDay(); // 0=Sunday, 6=Saturday

  // 土曜日の夕食は提供なし
  if (mealType === "dinner" && dayOfWeek === 6) {
    return;
  }

  const yyyyMM = `${year}${month.toString().padStart(2, "0")}`;
  const sheetName = `食事原紙_${yyyyMM}`;
  const sheet = mealSs.getSheetByName(sheetName);

  if (!sheet) {
    return;
  }

  // 列オフセット: 朝食=3列目(C), 夕食=4列目(D) から2列/日
  const mealOffset = mealType === "dinner" ? 4 : 3;

  let col;
  let targetRowStart;
  let targetRowEnd;

  if (day <= 16) {
    // 前半（1〜16日）: ヘッダー行2, ユーザー行5〜37
    col = (day - 1) * 2 + mealOffset;
    targetRowStart = 5;
    targetRowEnd = 37;
  } else {
    // 後半（17〜31日）: ヘッダー行42, ユーザー行45〜77
    col = (day - 17) * 2 + mealOffset;
    targetRowStart = 45;
    targetRowEnd = 77;
  }

  // A列を走査してuserIdに対応する行を取得
  const aValues = sheet
    .getRange(targetRowStart, 1, targetRowEnd - targetRowStart + 1, 1)
    .getValues();
  let targetRow = -1;
  for (let i = 0; i < aValues.length; i++) {
    if (aValues[i][0] == userId) {
      targetRow = targetRowStart + i;
      break;
    }
  }

  if (targetRow === -1) {
    return;
  }

  sheet.getRange(targetRow, col).setValue(isReserved ? 1 : "");
}

function syncMealSheetFromReservations() {
  const today = new Date();
  const tomorrow = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 1);
  const baseDateStr = formatDate(tomorrow);

  console.log(`[syncMealSheetFromReservations] 開始: ${baseDateStr} 以降のデータを同期します`);

  const thisYear = today.getFullYear();
  const thisMonth = today.getMonth() + 1;
  let nextMonth = thisMonth + 1;
  let nextYear = thisYear;
  if (nextMonth > 12) { nextMonth = 1; nextYear++; }

  const targetMonths = [
    { year: thisYear, month: thisMonth },
    { year: nextYear, month: nextMonth }
  ];

  let totalUpdated = 0;
  let totalSkipped = 0;

  for (const { year, month } of targetMonths) {
    const yyyyMM = `${year}${month.toString().padStart(2, "0")}`;
    console.log(`--- ${year}年${month}月 処理開始 ---`);

    const bCalendarSheet    = ss.getSheetByName(`b_calendar_${yyyyMM}`);
    const dCalendarSheet    = ss.getSheetByName(`d_calendar_${yyyyMM}`);
    const bReservationSheet = ss.getSheetByName(`b_reservations_${yyyyMM}`);
    const dReservationSheet = ss.getSheetByName(`d_reservations_${yyyyMM}`);

    if (!bCalendarSheet || !dCalendarSheet || !bReservationSheet || !dReservationSheet) {
      console.warn(`${year}年${month}月: 必要なシートが見つかりません。スキップします。`);
      continue;
    }

    const buildDateMap = (calSheet, idCol) => {
      const data = calSheet.getDataRange().getValues();
      const headers = data[0];
      const idIdx = headers.indexOf(idCol);
      const dtIdx = headers.indexOf("date");
      const map = {};
      for (let i = 1; i < data.length; i++) {
        const d = data[i][dtIdx];
        map[data[i][idIdx]] = d instanceof Date ? formatDate(d) : String(d);
      }
      return map;
    };

    const bDateMap = buildDateMap(bCalendarSheet, "b_calendar_id");
    const dDateMap = buildDateMap(dCalendarSheet, "d_calendar_id");

    const mealConfigs = [
      { sheet: bReservationSheet, dateMap: bDateMap, mealType: "breakfast", idCol: "b_calendar_id" },
      { sheet: dReservationSheet, dateMap: dDateMap, mealType: "dinner",    idCol: "d_calendar_id" }
    ];

    for (const { sheet, dateMap, mealType, idCol } of mealConfigs) {
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const calIdx = headers.indexOf(idCol);
      const uidIdx = headers.indexOf("user_id");
      const resIdx = headers.indexOf("is_reserved");

      let updated = 0;
      let skipped = 0;

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const calId = row[calIdx];
        const userId = row[uidIdx];
        const isReserved = Boolean(row[resIdx]);
        const dateStr = dateMap[calId];

        if (!dateStr || dateStr < baseDateStr) {
          skipped++;
          totalSkipped++;
          continue;
        }

        updateMealSheetForUser(userId, dateStr, mealType, isReserved);
        updated++;
        totalUpdated++;
      }

      console.log(`  ${mealType}: 更新=${updated}件, スキップ=${skipped}件`);
    }

    console.log(`--- ${year}年${month}月 処理完了 ---`);
  }

  console.log(`[syncMealSheetFromReservations] 完了: 更新=${totalUpdated}件, スキップ=${totalSkipped}件`);
}

function syncAprilDinnerOnly() {
  const baseDateStr = "2026-03-27";
  const year = 2026;
  const month = 4;
  const yyyyMM = "202604";

  console.log(`[syncAprilDinnerOnly] 開始`);

  const dCalendarSheet    = ss.getSheetByName(`d_calendar_${yyyyMM}`);
  const dReservationSheet = ss.getSheetByName(`d_reservations_${yyyyMM}`);

  if (!dCalendarSheet || !dReservationSheet) {
    console.warn("4月の夕食シートが見つかりません。");
    return;
  }

  const calData = dCalendarSheet.getDataRange().getValues();
  const calHeaders = calData[0];
  const calIdIdx = calHeaders.indexOf("d_calendar_id");
  const calDtIdx = calHeaders.indexOf("date");
  const dateMap = {};
  for (let i = 1; i < calData.length; i++) {
    const d = calData[i][calDtIdx];
    dateMap[calData[i][calIdIdx]] = d instanceof Date ? formatDate(d) : String(d);
  }

  const resData = dReservationSheet.getDataRange().getValues();
  const resHeaders = resData[0];
  const calIdx = resHeaders.indexOf("d_calendar_id");
  const uidIdx = resHeaders.indexOf("user_id");
  const resIdx = resHeaders.indexOf("is_reserved");

  let updated = 0;
  for (let i = 1; i < resData.length; i++) {
    const row = resData[i];
    const dateStr = dateMap[row[calIdx]];
    if (!dateStr || dateStr < baseDateStr) continue;
    updateMealSheetForUser(row[uidIdx], dateStr, "dinner", Boolean(row[resIdx]));
    updated++;
  }

  console.log(`[syncAprilDinnerOnly] 完了: 更新=${updated}件`);
}
