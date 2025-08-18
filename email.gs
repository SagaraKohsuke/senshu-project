/**
 * メール一斉送信機能
 */

/**
 * 全ユーザーに個別の予約リンク付きメールを一斉送信
 * @param {string} subject - メールの件名
 * @param {string} bodyTemplate - メール本文のテンプレート（{name}と{link}がプレースホルダー）
 * @return {Object} 送信結果
 */
function sendBulkEmailToUsers(subject, bodyTemplate) {
  try {
    const userSheet = ss.getSheetByName("users");
    if (!userSheet) {
      return {
        success: false,
        message: "usersシートが見つかりません"
      };
    }
    
    // データを取得（1行目はヘッダー）
    const userData = userSheet.getDataRange().getValues();
    if (userData.length <= 1) {
      return {
        success: false,
        message: "ユーザーデータが見つかりません"
      };
    }
    
    // ベースURL
    const baseUrl = "https://script.google.com/macros/s/AKfycbyV0jDcsGHIRAY79IRsDMVEGa7RPlrpwt_Bu-Xn8BEp6LQabxhedrKbPExuaNSZjlrPJw/exec";
    
    let successCount = 0;
    let failureCount = 0;
    const errors = [];
    
    // 2行目から処理（1行目はヘッダー）
    for (let i = 1; i < userData.length; i++) {
      const row = userData[i];
      const userId = row[0];  // A列: user_id
      const name = row[1];    // B列: name
      const email = row[2];   // C列: email
      
      // 必要なデータがない場合はスキップ
      if (!userId || !name || !email) {
        failureCount++;
        errors.push(`行${i + 1}: 必要なデータが不足 (userId: ${userId}, name: ${name}, email: ${email})`);
        continue;
      }
      
      // 個別の予約リンクを作成
      const personalLink = `${baseUrl}?room=${userId}`;
      
      // メール本文を個人用にカスタマイズ
      const personalizedBody = bodyTemplate
        .replace(/\{name\}/g, name)
        .replace(/\{link\}/g, personalLink);
      
      try {
        // メール送信
        MailApp.sendEmail({
          to: email,
          subject: subject,
          htmlBody: personalizedBody
        });
        
        successCount++;
        
        // Gmail API制限対策のため少し待機
        Utilities.sleep(100);
        
      } catch (error) {
        failureCount++;
        errors.push(`${name} (${email}): ${error.message}`);
      }
    }
    
    return {
      success: true,
      message: `メール送信完了。成功: ${successCount}件、失敗: ${failureCount}件`,
      successCount: successCount,
      failureCount: failureCount,
      errors: errors
    };
    
  } catch (error) {
    console.error("メール一斉送信エラー:", error);
    return {
      success: false,
      message: "メール送信中にエラーが発生しました: " + error.message
    };
  }
}
