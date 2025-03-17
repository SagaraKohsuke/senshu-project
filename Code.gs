function doGet() {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("泉州会館 食事予約");
}

function getAllSheetsData() {
  var spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
  var sheetNames = [
    "users",
    "b_menus",
    "d_menus",
    "b_calendar_202503",
    "d_calendar_202503",
    "b_reservation_202503",
    "d_reservation_202503",
  ];

  var result = {};

  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);

  for (var i = 0; i < sheetNames.length; i++) {
    var sheetName = sheetNames[i];

    try {
      var sheet = spreadsheet.getSheetByName(sheetName);
      if (sheet) {
        var range = sheet.getDataRange();
        var values = range.getValues();
        var headers = values[0];
        var data = [];
        for (var j = 1; j < values.length; j++) {
          var rowData = {};
          for (var k = 0; k < headers.length; k++) {
            rowData[headers[k]] = values[j][k];
          }
          data.push(rowData);
        }
        result[sheetName] = {
          headers: headers,
          data: data,
        };
      } else {
        Logger.log("シート「" + sheetName + "」が見つかりません");
        result[sheetName] = {
          error: "Sheet not found",
          headers: [],
          data: [],
        };
      }
    } catch (error) {
      Logger.log("シート「" + sheetName + "」の処理中にエラー: " + error);
      result[sheetName] = {
        error: error.toString(),
        headers: [],
        data: [],
      };
    }
  }

  Logger.log(JSON.stringify(result, null, 2));

  return result;
}

function getInitialData() {
  const allData = getAllSheetsData();
  if (!allData.users || !allData.b_menus || !allData.d_menus) {
    Logger.log(
      "初期データ取得エラー: users, b_menus, d_menus シートのいずれかが存在しません"
    );
    throw new Error(
      "初期データ取得エラー: users, b_menus, d_menus シートのいずれかが存在しません"
    );
  }
  const users = allData.users.data.map((user) => ({
    id: user.user_id,
    name: user.name,
  }));
  const breakfastMenus = allData.b_menus.data.map((menu) => ({
    id: `B${menu.b_menu_id}`,
    name: menu.breakfast_menu,
  }));
  const dinnerMenus = allData.d_menus.data.map((menu) => ({
    id: `D${menu.d_menu_id}`,
    name: menu.dinner_menu,
  }));
  const menus = [...breakfastMenus, ...dinnerMenus];

  return {
    users: users,
    menus: menus,
  };
}

function getCalendarData(userId) {
  const allData = getAllSheetsData();
  if (
    !allData.b_calendar_202503 ||
    !allData.d_calendar_202503 ||
    !allData.b_reservation_202503 ||
    !allData.d_reservation_202503
  ) {
    Logger.log(
      "カレンダーデータ取得エラー: カレンダーまたは予約シートが存在しません"
    );
    throw new Error(
      "カレンダーデータ取得エラー: カレンダーまたは予約シートが存在しません"
    );
  }

  const breakfastCalendar = allData.b_calendar_202503.data.map((entry) => {
    const date = new Date(entry.date);
    const formattedDate = date.toISOString().split("T")[0];

    return {
      id: entry.b_calendar_id,
      date: formattedDate,
      menu_id: `B${entry.b_menu_id}`,
    };
  });

  const dinnerCalendar = allData.d_calendar_202503.data.map((entry) => {
    const date = new Date(entry.date);
    const formattedDate = date.toISOString().split("T")[0];

    return {
      id: entry.d_calendar_id,
      date: formattedDate,
      menu_id: `D${entry.d_menu_id}`,
    };
  });

  const breakfastReservations = allData.b_reservation_202503.data
    .filter((reservation) => String(reservation.user_id) === String(userId))
    .map((reservation) => ({
      id: reservation.b_reservation_id,
      calendar_id: reservation.b_calendar_id,
      is_reserved: reservation.is_reserved === "TRUE",
    }));

  const dinnerReservations = allData.d_reservation_202503.data
    .filter((reservation) => String(reservation.user_id) === String(userId))
    .map((reservation) => ({
      id: reservation.d_reservation_id,
      calendar_id: reservation.d_calendar_id,
      is_reserved: reservation.is_reserved === "TRUE",
    }));

  return {
    calendar: [...breakfastCalendar, ...dinnerCalendar],
    reservations: [...breakfastReservations, ...dinnerReservations],
  };
}

function updateReservation(userId, calendarId, mealType, isReserved) {
  var spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheetName = "";
  var idColumnName = "";
  var calendarIdColumnName = "";
  var isReservedColumnName = "is_reserved";

  if (mealType === "breakfast") {
    sheetName = "b_reservation_202503";
    idColumnName = "b_reservation_id";
    calendarIdColumnName = "b_calendar_id";
  } else if (mealType === "dinner") {
    sheetName = "d_reservation_202503";
    idColumnName = "d_reservation_id";
    calendarIdColumnName = "d_calendar_id";
  } else {
    Logger.log("mealType が不正です: " + mealType);
    return { success: false, error: "Invalid mealType" };
  }

  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log("シート「" + sheetName + "」が見つかりません");
    return { success: false, error: "Sheet not found: " + sheetName };
  }

  var headers = sheet.getDataRange().getValues()[0];
  var idColumnIndex = headers.indexOf(idColumnName) + 1;
  var calendarIdColumnIndex = headers.indexOf(calendarIdColumnName) + 1;
  var isReservedColumnIndex = headers.indexOf(isReservedColumnName) + 1;

  if (
    idColumnIndex === 0 ||
    calendarIdColumnIndex === 0 ||
    isReservedColumnIndex === 0
  ) {
    Logger.log("必要な列が見つかりません");
    return {
      success: false,
      error: "Required column not found in sheet: " + sheetName,
    };
  }

  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  var lastRow = dataRange.getLastRow();

  for (var i = 1; i < lastRow; i++) {
    if (
      values[i][calendarIdColumnIndex - 1] === calendarId &&
      String(values[i][1]) === String(userId)
    ) {
      // user_id は2列目と仮定
      sheet
        .getRange(i + 1, isReservedColumnIndex)
        .setValue(isReserved ? "TRUE" : "FALSE");
      return { success: true };
    }
  }

  if (mealType === "breakfast") {
    sheet.appendRow([
      Utilities.getUuid(),
      userId,
      calendarId,
      isReserved ? "TRUE" : "FALSE",
    ]);
  } else if (mealType === "dinner") {
    sheet.appendRow([
      Utilities.getUuid(),
      userId,
      calendarId,
      isReserved ? "TRUE" : "FALSE",
    ]);
  }

  return { success: true };
}
