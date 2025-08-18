function doGet(e) {
  const template = HtmlService.createTemplateFromFile("index");
  
  // デバッグログ
  console.log('=== doGet called ===');
  console.log('e:', e);
  console.log('e.parameter:', e.parameter);
  
  // URL parameters を template に渡す
  const hasParam = e && e.parameter && Object.keys(e.parameter).length > 0;
  const params = e && e.parameter ? e.parameter : {};
  
  template.hasPathParam = hasParam;
  template.pathParams = JSON.stringify(params);
  
  console.log('hasPathParam:', hasParam);
  console.log('pathParams:', JSON.stringify(params));
  console.log('room parameter:', params.room);
  console.log('==================');
  
  return template.evaluate().setTitle("泉州会館 食事予約");
}

const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
const ss = SpreadsheetApp.openById(spreadsheetId);
