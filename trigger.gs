/**
 * 自動トリガー管理システム
 * 
 * このファイルは以下の機能を提供します:
 * - 毎月のシート管理 (1日の午前3時に実行)
 * - 月末リマインダーメール送信 (29日の午前10時に実行)
 * 
 * 注意: monthlyReminderManagement は email.gs の sendMonthlyReminderEmail 関数に依存します
 */

function tmpSheetManagement(){
  createCalendarSheetsForYearMonth(2025, 3);
  createAllUsersReservations(2025, 3);      
}

function monthlySheetManagement() {
  // 現在の日付を取得
  const today = new Date();
  
  // 翌月の年月を計算
  let nextMonth = today.getMonth() + 2; // JavaScriptの月は0始まりなので+2で翌月
  let nextMonthYear = today.getFullYear();
  
  if (nextMonth > 12) {
    nextMonth = nextMonth - 12;
    nextMonthYear++;
  }
  
  // 2ヶ月前の年月を計算
  let twoMonthsAgo = today.getMonth() - 1; // 現在の月から2ヶ月前
  let twoMonthsAgoYear = today.getFullYear();
  
  if (twoMonthsAgo < 1) {
    twoMonthsAgo = twoMonthsAgo + 12;
    twoMonthsAgoYear--;
  }
  
  // 翌月のシートを作成
  console.log(`Creating sheets for ${nextMonthYear}/${nextMonth}`);
  const createResult = createCalendarSheetsForYearMonth(nextMonthYear, nextMonth);
  
  if (createResult.success) {
    // 予約データを作成
    console.log(`Creating reservations for ${nextMonthYear}/${nextMonth}`);
    const reservationResult = createAllUsersReservations(nextMonthYear, nextMonth);
    console.log(reservationResult.message);
  } else {
    console.error(createResult.message);
    return;
  }
  
  // 2ヶ月前のシートを削除
  deleteOldSheets(twoMonthsAgoYear, twoMonthsAgo);
  
  console.log('Monthly sheet management completed successfully.');
}

/**
 * 月末リマインダーメール送信処理
 */
function monthlyReminderManagement() {
  try {
    console.log('Monthly reminder email sending started');
    
    const result = sendMonthlyReminderEmail();
    
    if (result.success) {
      console.log('Monthly reminder email sent successfully:', result.message);
      console.log(`Success: ${result.successCount}, Failures: ${result.failureCount}`);
      
      if (result.errors && result.errors.length > 0) {
        console.warn('Email sending errors:', result.errors);
      }
    } else {
      console.error('Monthly reminder email sending failed:', result.message);
    }
    
    return result;
    
  } catch (error) {
    console.error('Monthly reminder management error:', error);
    return {
      success: false,
      message: 'リマインダーメール送信中にエラーが発生しました: ' + error.message
    };
  }
}


function deleteOldSheets(year, month) {
  const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
  const ss = SpreadsheetApp.openById(spreadsheetId);
  
  // シート名を生成
  const yyyyMM = `${year}${month.toString().padStart(2, '0')}`;
  const bCalendarSheetName = `b_calendar_${yyyyMM}`;
  const dCalendarSheetName = `d_calendar_${yyyyMM}`;
  const bReservationSheetName = `b_${yyyyMM}reservation`;
  const dReservationSheetName = `d_${yyyyMM}reservation`;
  
  const sheetsToDelete = [
    bCalendarSheetName,
    dCalendarSheetName, 
    bReservationSheetName, 
    dReservationSheetName
  ];
  

  let deletedCount = 0;
  for (const sheetName of sheetsToDelete) {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      ss.deleteSheet(sheet);
      deletedCount++;
      console.log(`Deleted sheet: ${sheetName}`);
    }
  }
  
  console.log(`Deleted ${deletedCount} old sheets for ${year}/${month}`);
}


function setMonthlyTrigger() {
  // 既存のトリガーをすべて削除
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'monthlySheetManagement') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  
  // 新しいトリガーを設定（毎月1日の午前3時に実行）
  ScriptApp.newTrigger('monthlySheetManagement')
    .timeBased()
    .onMonthDay(1)
    .atHour(3)
    .create();
  
  console.log('Monthly trigger has been set successfully.');
}

/**
 * 月末リマインダーメール送信のトリガーを設定
 */
function setReminderTrigger() {
  // 既存のリマインダートリガーをすべて削除
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'monthlyReminderManagement') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  
  // 新しいリマインダートリガーを設定（毎月29日の午前10時に実行）
  // 29日を選択することで、30日・31日のない月でも確実に月末2日前に実行される
  ScriptApp.newTrigger('monthlyReminderManagement')
    .timeBased()
    .onMonthDay(29)
    .atHour(10)
    .create();
  
  console.log('Monthly reminder trigger has been set successfully (29th of every month at 10:00 AM).');
}

/**
 * 全てのトリガーを一括設定
 */
function setAllTriggers() {
  setMonthlyTrigger();
  setReminderTrigger();
  console.log('All triggers have been set successfully.');
}