function getUserReservationCalendar(userId, year, month) {
  const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
  const ss = SpreadsheetApp.openById(spreadsheetId);
    
  const yyyyMM = `${year}${month.toString().padStart(2, '0')}`;
  const bCalendarSheetName = `b_calendar_${yyyyMM}`;
  const dCalendarSheetName = `d_calendar_${yyyyMM}`;
  const bReservationSheetName = `b_reservation_${yyyyMM}`;
  const dReservationSheetName = `d_reservation_${yyyyMM}`;
    
  const bCalendarSheet = ss.getSheetByName(bCalendarSheetName);
  const dCalendarSheet = ss.getSheetByName(dCalendarSheetName);
  const bReservationSheet = ss.getSheetByName(bReservationSheetName);
  const dReservationSheet = ss.getSheetByName(dReservationSheetName);
  const bMenuSheet = ss.getSheetByName('b_menu');
  const dMenuSheet = ss.getSheetByName('d_menu');
  
  if (!bCalendarSheet || !dCalendarSheet || !bReservationSheet || !dReservationSheet) {
    return {
      success: false,
      message: `必要なシートが見つかりません。カレンダーと予約シートが存在するか確認してください。`
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
  
  const bCalendarIdIndex = bCalendarHeaders.indexOf('b_calendar_id');
  const bCalendarDateIndex = bCalendarHeaders.indexOf('date');
  const bCalendarMenuIdIndex = bCalendarHeaders.indexOf('b_menu_id');
  
  const dCalendarIdIndex = dCalendarHeaders.indexOf('d_calendar_id');
  const dCalendarDateIndex = dCalendarHeaders.indexOf('date');
  const dCalendarMenuIdIndex = dCalendarHeaders.indexOf('d_menu_id');
  
  const bReservationCalendarIdIndex = bReservationHeaders.indexOf('b_calendar_id');
  const bReservationUserIdIndex = bReservationHeaders.indexOf('user_id');
  const bReservationStatusIndex = bReservationHeaders.indexOf('is_reserved');
  
  const dReservationCalendarIdIndex = dReservationHeaders.indexOf('d_calendar_id');
  const dReservationUserIdIndex = dReservationHeaders.indexOf('user_id');
  const dReservationStatusIndex = dReservationHeaders.indexOf('is_reserved');
  
  // 朝食メニューマップの作成
  const bMenuMap = {};
  if (bMenuData.length > 1) {
    const bMenuIdIndex = bMenuData[0].indexOf('b_menu_id');
    const bMenuNameIndex = bMenuData[0].indexOf('breakfast_menu');
    
    if (bMenuIdIndex !== -1 && bMenuNameIndex !== -1) {
      for (let i = 1; i < bMenuData.length; i++) {
        const menuId = bMenuData[i][bMenuIdIndex];
        const menuName = bMenuData[i][bMenuNameIndex];
        bMenuMap[menuId] = menuName;
      }
    }
  }
  
  // 夕食メニューマップの作成
  const dMenuMap = {};
  if (dMenuData.length > 1) {
    const dMenuIdIndex = dMenuData[0].indexOf('d_menu_id');
    const dMenuNameIndex = dMenuData[0].indexOf('dinner_menu');
    
    if (dMenuIdIndex !== -1 && dMenuNameIndex !== -1) {
      for (let i = 1; i < dMenuData.length; i++) {
        const menuId = dMenuData[i][dMenuIdIndex];
        const menuName = dMenuData[i][dMenuNameIndex];
        dMenuMap[menuId] = menuName;
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
        isReserved: row[bReservationStatusIndex]
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
        isReserved: row[dReservationStatusIndex]
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
    
    if (date instanceof Date) {
      const dateStr = formatDate(date);
      breakfastCalendar[calendarId] = {
        id: calendarId,
        date: dateStr,
        menu_id: menuId,
        menu_name: bMenuMap[menuId] || '未設定'
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
    
    if (date instanceof Date) {
      const dateStr = formatDate(date);
      dinnerCalendar[calendarId] = {
        id: calendarId,
        date: dateStr,
        menu_id: menuId,
        menu_name: dMenuMap[menuId] || '未設定'
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
        is_reserved: reservation.isReserved
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
        is_reserved: reservation.isReserved
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
    dinner: dMenuMap
  };
  
  return {
    success: true,
    userId: userId,
    year: year,
    month: month,
    offset: offset,
    monthDates: monthDates,
    weekdays: ['日', '月', '火', '水', '木', '金', '土'],
    groupedBreakfast: groupedBreakfast,
    groupedDinner: groupedDinner,
    menuMap: menuMap
  };
}



function updateReservation(userId, calendarId, mealType, isReserved, year, month) {
  const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
  const ss = SpreadsheetApp.openById(spreadsheetId);
    
  const yyyyMM = `${year}${month.toString().padStart(2, '0')}`;
  const prefix = mealType === 'breakfast' ? 'b' : 'd';
  const calendarSheetName = `${prefix}_calendar_${yyyyMM}`;
  const reservationSheetName = `${prefix}_reservation_${yyyyMM}`;
    
  const reservationSheet = ss.getSheetByName(reservationSheetName);
  if (!reservationSheet) {
    return {
      success: false,
      message: `予約シート ${reservationSheetName} が見つかりません。`
    };
  }
    
  const reservationData = reservationSheet.getDataRange().getValues();
  const headers = reservationData[0];
   
  const calendarIdIndex = headers.indexOf(`${prefix}_calendar_id`);
  const userIdIndex = headers.indexOf('user_id');
  const isReservedIndex = headers.indexOf('is_reserved');
  
  if (calendarIdIndex === -1 || userIdIndex === -1 || isReservedIndex === -1) {
    return {
      success: false,
      message: '必要なカラムが予約シートに見つかりません。'
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
      message: `ユーザーID ${userId} とカレンダーID ${calendarId} に一致する予約が見つかりません。`
    };
  }
    
  reservationSheet.getRange(rowIndex + 1, isReservedIndex + 1).setValue(isReserved);
  
  return {
    success: true
  };
}


function createAllUsersReservations(year, month) {
  const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
  const ss = SpreadsheetApp.openById(spreadsheetId);
  
  // シート名を生成
  const yyyyMM = `${year}${month.toString().padStart(2, '0')}`;
  const bCalendarSheetName = `b_calendar_${yyyyMM}`;
  const dCalendarSheetName = `d_calendar_${yyyyMM}`;
  const bReservationSheetName = `b_reservation_${yyyyMM}`;
  const dReservationSheetName = `d_reservation_${yyyyMM}`;
  
  // カレンダーシートの存在確認
  const bCalendarSheet = ss.getSheetByName(bCalendarSheetName);
  const dCalendarSheet = ss.getSheetByName(dCalendarSheetName);
  
  if (!bCalendarSheet || !dCalendarSheet) {
    return {
      success: false,
      message: `カレンダーシートが存在しません。先にcreateCalendarSheetsForYearMonth関数を実行してください。`
    };
  }
  
  // ユーザー一覧を取得
  const usersSheet = ss.getSheetByName('users');
  if (!usersSheet) {
    return {
      success: false,
      message: `ユーザーシートが存在しません。`
    };
  }
  
  const usersData = usersSheet.getDataRange().getValues();
  const userIdIndex = usersData[0].indexOf('user_id');
  
  if (userIdIndex === -1) {
    return {
      success: false,
      message: `ユーザーシートにuser_idカラムが見つかりません。`
    };
  }
  
  // ヘッダー行を除外したユーザーIDの配列
  const userIds = [];
  for (let i = 1; i < usersData.length; i++) {
    const userId = usersData[i][userIdIndex];
    if (userId) { // 空でないIDのみ追加
      userIds.push(userId);
    }
  }
  
  // 予約シートの取得または作成
  let bReservationSheet = ss.getSheetByName(bReservationSheetName);
  let dReservationSheet = ss.getSheetByName(dReservationSheetName);
  
  // 朝食予約シートがなければ作成
  if (!bReservationSheet) {
    bReservationSheet = ss.insertSheet(bReservationSheetName);
    bReservationSheet.getRange('A1:D1').setValues([['b_reservation_id', 'b_calendar_id', 'user_id', 'is_reserved']]);
    bReservationSheet.getRange('A1:D1').setFontWeight('bold');
    bReservationSheet.autoResizeColumns(1, 4);
  }
  
  // 夕食予約シートがなければ作成
  if (!dReservationSheet) {
    dReservationSheet = ss.insertSheet(dReservationSheetName);
    dReservationSheet.getRange('A1:D1').setValues([['d_reservation_id', 'd_calendar_id', 'user_id', 'is_reserved']]);
    dReservationSheet.getRange('A1:D1').setFontWeight('bold');
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
          false // デフォルトは予約なし
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
          false // デフォルトは予約なし
        ]);
        dReservationIdCounter++;
        dRowsAdded++;
      }
    }
  }
  
  // 一括で予約データを追加（効率的な書き込み）
  if (bReservationsToAdd.length > 0) {
    const lastRow = bReservationSheet.getLastRow();
    bReservationSheet.getRange(lastRow + 1, 1, bReservationsToAdd.length, 4).setValues(bReservationsToAdd);
  }
  
  if (dReservationsToAdd.length > 0) {
    const lastRow = dReservationSheet.getLastRow();
    dReservationSheet.getRange(lastRow + 1, 1, dReservationsToAdd.length, 4).setValues(dReservationsToAdd);
  }
  
  return {
    success: true,
    message: `${year}年${month}月の予約データを${userIds.length}人分作成しました。朝食: ${bRowsAdded}件, 夕食: ${dRowsAdded}件追加されました。`,
    totalUsers: userIds.length,
    bRowsAdded: bRowsAdded,
    dRowsAdded: dRowsAdded
  };
}