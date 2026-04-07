# 朝夕食予約システム

Google Spreadsheet と Google Apps Script (GAS)を使用した朝夕食予約管理 Web アプリケーションです。ユーザーは自分の朝食と夕食の予約状況をカレンダー形式で確認・変更することができます。

## 機能概要

- ユーザー認証（ユーザー ID 選択方式）
- 月単位のカレンダー表示
- 朝食・夕食の予約状況表示
- チェックボックスによる予約の追加・削除
- メニュー情報の表示
- 月間ナビゲーション（前月・次月）
- 月末リマインダーメールの自動送信（毎月 29 日午前 10 時）
- 全ユーザーへの予約システム URL メールの一斉送信
- お問い合わせフォーム

## システム構成

### データベース (Google Spreadsheet)

以下のシートで構成されています：

1. **users**: ユーザー情報

   - user_id: ユーザー ID（部屋番号）
   - name: ユーザー名
   - email: メールアドレス（リマインダーメール送信に使用）

2. **b_calendar_YYYYMM**: 朝食カレンダー（年月ごと）

   - b_calendar_id: カレンダー ID
   - date: 日付
   - b_menu_id: 朝食メニュー ID

3. **d_calendar_YYYYMM**: 夕食カレンダー（年月ごと）

   - d_calendar_id: カレンダー ID
   - date: 日付
   - d_menu_id: 夕食メニュー ID

4. **b_menu**: 朝食メニューマスタ

   - b_menu_id: 朝食メニュー ID
   - breakfast_menu: 朝食メニュー名

5. **d_menu**: 夕食メニューマスタ

   - d_menu_id: 夕食メニュー ID
   - dinner_menu: 夕食メニュー名

6. **b_reservations_YYYYMM**: 朝食予約情報（年月ごと）

   - b_reservation_id: 予約 ID
   - b_calendar_id: 朝食カレンダー ID
   - user_id: ユーザー ID
   - is_reserved: 予約状態（true/false）

7. **d_reservations_YYYYMM**: 夕食予約情報（年月ごと）
   - d_reservation_id: 予約 ID
   - d_calendar_id: 夕食カレンダー ID
   - user_id: ユーザー ID
   - is_reserved: 予約状態（true/false）

### バックエンド (Google Apps Script)

1. **getUsers()**: ユーザー一覧を取得
2. **createCalendarSheetsForYearMonth(year, month)**: 指定年月のカレンダーシートを作成
3. **createAllUsersReservations(year, month)**: 全ユーザーの予約データを一括生成
4. **getUserReservationCalendar(userId, year, month)**: ユーザーの予約カレンダーデータを取得
5. **updateReservation(userId, calendarId, mealType, isReserved, year, month)**: 予約状態を更新
6. **sendBulkEmailToUsers(subject, bodyTemplate)**: 全ユーザーに個別リンク付きメールを一斉送信
7. **sendReservationURLEmail()**: 予約システム URL を全ユーザーにメール通知
8. **sendMonthlyReminderEmail()**: 翌月予約を促すリマインダーメールを全ユーザーに送信
9. **checkGmailSettings()**: Gmail 送信権限と残り送信可能件数を確認
10. **submitQuestionnaire(type, detail)**: お問い合わせ内容をスプレッドシートに保存

### フロントエンド (HTML/CSS/JavaScript with Vue.js)

1. **ユーザー選択画面**:

   - ドロップダウンリストからユーザー（部屋番号）を選択

2. **カレンダー表示**:

   - 月間カレンダー形式で朝食と夕食の予約状況を表示
   - 各日付のセルに朝食と夕食の情報を表示
   - チェックボックスで予約状態を切り替え
   - 土曜日の夕食は予約不可
   - メニュー情報の表示

3. **ナビゲーション**:
   - 前月・次月ボタンで月の移動が可能

4. **お問い合わせフォーム** (contact.html):
   - お問い合わせ種別の選択（要望・クレーム・その他）
   - 詳細内容の入力と送信
   - 送信内容は専用スプレッドシートに自動保存

## 設計上の特徴

- **月ごとのテーブル分割**: データ管理を容易にするため、カレンダーと予約情報は月ごとに別テーブルで管理
- **データの効率的な取得**: 一度に全データを取得してクライアント側でフィルタリング
- **リアルタイム更新**: チェックボックス変更時に即座にサーバーサイドのデータが更新
- **メール一斉送信**: 全ユーザーに個別リンク付きのリマインダーメールを自動送信
- **自動トリガー管理**: 月次シート作成（毎月 1 日 午前 3 時）と月末リマインダー送信（毎月 29 日 午前 10 時）を自動実行

## 使用技術

- **Google Spreadsheet**: データベースとして使用
- **Google Apps Script (GAS)**: サーバーサイドスクリプト
- **Gmail (MailApp)**: リマインダーメールおよび URL 通知メールの送信
- **HTML/CSS**: UI の構築
- **Vue.js**: フロントエンドフレームワーク
- **JavaScript**: クライアントサイド処理

## セットアップ

1. Google Spreadsheet を作成
2. 必要なシート（users, b_menu, d_menu）を作成し、初期データを入力（users シートには user_id・name・email 列を用意）
3. Google Apps Script エディタを開き、スクリプトファイルを作成
4. 提供された GAS コードをスクリプトに貼り付け
5. HTML ファイルを作成し、提供された HTML コードを貼り付け
6. Web アプリケーションとして公開
7. GAS エディタで `setAllTriggers()` を実行してトリガーを設定（月次シート管理と月末リマインダーが自動設定されます）
8. 初回メール送信テストは `checkGmailSettings()` で権限確認後、`sendReservationURLEmail()` を実行

## 使用方法

1. アプリケーションにアクセス
2. ドロップダウンリストからユーザーを選択
3. 表示されるカレンダーで予約状況を確認
4. チェックボックスをクリック/クリック解除して予約を管理
5. 前月/次月ボタンで別の月の予約状況を確認・変更

## 注意点

- メニューが「未設定」と表示される場合は、メニューマスターのデータが未入力か、カレンダーテーブルのメニュー ID との紐付けが正しくない可能性があります
- 日曜日はカレンダーに表示されますが、カレンダーテーブルに日曜日のデータがない場合は「データなし」と表示されます
- 土曜日の夕食は予約できません
- メール送信機能を使用するには users シートに email 列（C 列）が必要です
- Gmail の 1 日の送信上限は通常 100 通です。ユーザー数が多い場合は複数日に分けて送信してください
- 月末リマインダーは毎月 29 日に実行されるため、30 日・31 日のない月でも確実に月末 2 日前に通知されます
