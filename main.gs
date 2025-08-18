function doGet(e) {
  const template = HtmlService.createTemplateFromFile("index");
  
  // デバッグログ
  console.log('doGet parameter:', e.parameter);
  
  // URL parameters を template に渡す
  template.hasPathParam = e && e.parameter && Object.keys(e.parameter).length > 0;
  template.pathParams = e && e.parameter ? JSON.stringify(e.parameter) : '{}';
  
  console.log('hasPathParam:', template.hasPathParam);
  console.log('pathParams:', template.pathParams);
  
  return template.evaluate().setTitle("泉州会館 食事予約");
}

const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
const ss = SpreadsheetApp.openById(spreadsheetId);
