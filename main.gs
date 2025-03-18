function doGet() {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("泉州会館 食事予約");
}

const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
const ss = SpreadsheetApp.openById(spreadsheetId);
