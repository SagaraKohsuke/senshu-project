# 朝夕食予約システム

Google SpreadsheetとGoogle Apps Script (GAS)を使用した朝夕食予約管理Webアプリケーションです。ユーザーは自分の朝食と夕食の予約状況をカレンダー形式で確認・変更することができます。

## 機能概要

- ユーザー認証（ユーザーID選択方式）
- 月単位のカレンダー表示
- 朝食・夕食の予約状況表示
- チェックボックスによる予約の追加・削除
- メニュー情報の表示
- 月間ナビゲーション（前月・次月）

## システム構成

### データベース (Google Spreadsheet)

以下のシートで構成されています：

1. **users**: ユーザー情報
   - user_id: ユーザーID（部屋番号）
   - name: ユーザー名

2. **b_calendar_YYYYMM**: 朝食カレンダー（年月ごと）
   - b_calendar_id: カレンダーID
   - date: 日付
   - b_menu_id: 朝食メニューID

3. **d_calendar_YYYYMM**: 夕食カレンダー（年月ごと）
   - d_calendar_id: カレンダーID
   - date: 日付
   - d_menu_id: 夕食メニューID

4. **b_menu**: 朝食メニューマスタ
   - b_menu_id: 朝食メニューID
   - breakfast_menu: 朝食メニュー名

5. **d_menu**: 夕食メニューマスタ
   - d_menu_id: 夕食メニューID
   - dinner_menu: 夕食メニュー名

6. **b_YYYYMMreservation**: 朝食予約情報（年月ごと）
   - b_reservation_id: 予約ID
   - b_calendar_id: 朝食カレンダーID
   - user_id: ユーザーID
   - is_reserved: 予約状態（true/false）

7. **d_YYYYMMreservation**: 夕食予約情報（年月ごと）
   - d_reservation_id: 予約ID
   - d_calendar_id: 夕食カレンダーID
   - user_id: ユーザーID
   - is_reserved: 予約状態（true/false）

### バックエンド (Google Apps Script)

1. **getUsers()**: ユーザー一覧を取得
2. **createCalendarSheetsForYearMonth(year, month)**: 指定年月のカレンダーシートを作成
3. **createAllUsersReservations(year, month)**: 全ユーザーの予約データを一括生成
4. **getUserReservationCalendar(userId, year, month)**: ユーザーの予約カレンダーデータを取得
5. **updateReservation(userId, calendarId, mealType, isReserved, year, month)**: 予約状態を更新

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

## 設計上の特徴

- **月ごとのテーブル分割**: データ管理を容易にするため、カレンダーと予約情報は月ごとに別テーブルで管理
- **データの効率的な取得**: 一度に全データを取得してクライアント側でフィルタリング
- **リアルタイム更新**: チェックボックス変更時に即座にサーバーサイドのデータが更新

## 使用技術

- **Google Spreadsheet**: データベースとして使用
- **Google Apps Script (GAS)**: サーバーサイドスクリプト
- **HTML/CSS**: UIの構築
- **Vue.js**: フロントエンドフレームワーク
- **JavaScript**: クライアントサイド処理

## セットアップ

1. Google Spreadsheetを作成
2. 必要なシート（users, b_menu, d_menu）を作成し、初期データを入力
3. Google Apps Script エディタを開き、スクリプトファイルを作成
4. 提供されたGASコードをスクリプトに貼り付け
5. HTMLファイルを作成し、提供されたHTMLコードを貼り付け
6. Webアプリケーションとして公開

## 使用方法

1. アプリケーションにアクセス
2. ドロップダウンリストからユーザーを選択
3. 表示されるカレンダーで予約状況を確認
4. チェックボックスをクリック/クリック解除して予約を管理
5. 前月/次月ボタンで別の月の予約状況を確認・変更

## 注意点

- メニューが「未設定」と表示される場合は、メニューマスターのデータが未入力か、カレンダーテーブルのメニューIDとの紐付けが正しくない可能性があります
- 日曜日はカレンダーに表示されますが、カレンダーテーブルに日曜日のデータがない場合は「データなし」と表示されます
- 土曜日の夕食は予約できません
