function getUsers() {  
  const usersSheet = ss.getSheetByName('users');
  
  // データ範囲を一度に取得（行ごとの個別取得を避ける）
  const dataRange = usersSheet.getDataRange();
  const values = dataRange.getValues();
   
  const headers = values[0];
  const userIdIndex = headers.indexOf('user_id');
  const nameIndex = headers.indexOf('name');
    
  const usersArray = [];
  
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const userId = row[userIdIndex];
    
    if (userId) {
      usersArray.push({
        userId: userId,
        name: row[nameIndex]
      });
    }
  }  
  return usersArray;
}