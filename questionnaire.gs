function submitQuestionnaire(type, detail) {
  const spreadsheetId = "1-FWRtkpGNKHrLtc2V-f66ygsAI67wPaaWkt26R_Pg1c";
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName("questionnaire");
  
  if (!sheet) {
    return {
      success: false,
      message: "お問い合わせシートが見つかりません。"
    };
  }
  
  try {
    const timestamp = new Date();
    sheet.appendRow([timestamp, type, detail]);
    
    return {
      success: true,
      message: "お問い合わせを送信しました。"
    };
  } catch (error) {
    return {
      success: false,
      message: "送信中にエラーが発生しました: " + error.toString()
    };
  }
}
