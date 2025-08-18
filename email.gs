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

/**
 * 予約システムのURL付きメールを全ユーザーに送信
 * @return {Object} 送信結果
 */
function sendReservationURLEmail() {
  const subject = "【泉州会館】朝夕食予約システムのご案内";
  
  const bodyTemplate = `
    <div style="font-family: 'Helvetica Neue', Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px;">
      <h2 style="color: #333; border-bottom: 2px solid #007bff; padding-bottom: 10px;">
        朝夕食予約システムのご案内
      </h2>
      
      <p style="font-size: 16px; line-height: 1.6;">
        {name} 様
      </p>
      
      <p style="font-size: 16px; line-height: 1.6;">
        いつもお世話になっております。<br>
        朝夕食の予約システムをご利用いただき、ありがとうございます。
      </p>
      
      <div style="background: #f8f9fa; border-left: 4px solid #007bff; padding: 15px; margin: 20px 0;">
        <h3 style="color: #007bff; margin-top: 0;">🔗 あなた専用の予約ページ</h3>
        <p style="margin: 10px 0;">
          下記のリンクをクリックすると、お客様専用の予約ページが開きます。
        </p>
        <div style="text-align: center; margin: 15px 0;">
          <a href="{link}" style="display: inline-block; background: #007bff; color: white; padding: 15px 30px; text-decoration: none; border-radius: 8px; font-weight: bold; font-size: 16px;">
            📅 予約ページを開く
          </a>
        </div>
        <p style="font-size: 14px; color: #666; margin: 10px 0 0; word-break: break-all;">
          URL: <a href="{link}" style="color: #007bff;">{link}</a>
        </p>
      </div>
      
      <div style="background: #fff3cd; border: 1px solid #ffeaa7; padding: 15px; margin: 20px 0; border-radius: 5px;">
        <h4 style="color: #856404; margin-top: 0;">⏰ 予約締切時間のお知らせ</h4>
        <ul style="margin: 10px 0; color: #856404;">
          <li><strong>朝食</strong>：前日の12:00まで</li>
          <li><strong>夕食</strong>：当日の12:00まで</li>
        </ul>
        <p style="font-size: 14px; color: #856404; margin: 10px 0 0;">
          締切時間を過ぎると予約・キャンセルができませんのでご注意ください。
        </p>
      </div>
      
      <h4 style="color: #333; margin-top: 30px;">📱 システムの使い方</h4>
      <ol style="line-height: 1.8; color: #555;">
        <li>上記のリンクをクリックして予約ページを開く</li>
        <li>カレンダーから希望の日付の朝食・夕食にチェックを入れる</li>
        <li>予約内容は自動的に保存されます</li>
        <li>一括予約ボタンで月全体の予約も可能です</li>
      </ol>
      
      <div style="background: #e7f3ff; border: 1px solid #b3d9ff; padding: 15px; margin: 20px 0; border-radius: 5px;">
        <h4 style="color: #0056b3; margin-top: 0;">💡 便利な機能</h4>
        <ul style="margin: 10px 0; color: #0056b3;">
          <li><strong>スマートフォン対応</strong>：どのデバイスからでも利用可能</li>
          <li><strong>一括予約</strong>：月全体の朝食・夕食をまとめて予約</li>
          <li><strong>自動更新</strong>：締切時間になると自動的に画面が更新</li>
        </ul>
      </div>
      
      <div style="border-top: 1px solid #eee; padding-top: 20px; margin-top: 30px; font-size: 14px; color: #666;">
        <p>
          ご不明な点がございましたら、予約ページの「お問い合わせ」ボタンからお気軽にお尋ねください。
        </p>
        <p style="margin-top: 20px;">
          泉州会館<br>
          朝夕食予約システム
        </p>
      </div>
    </div>
  `;
  
  return sendBulkEmailToUsers(subject, bodyTemplate);
}
