# info.md

## プロジェクト概要

このプロジェクトは、Googleスプレッドシートの **「毎日のタスク」** シートと **Google Tasks** を同期するための Google Apps Script です。  
主な目的は、**日付ごとの学習タスク管理**をしやすくすることです。

### 実現したいこと
- スプレッドシートでタスクを管理（チェックボックスで完了）
- Google Tasks でも同じタスクを見られる / 完了できる
- 双方向同期（シート ⇄ Google Tasks）
- 日付をまたいだ未完了タスクの繰り越しを自動化
- 深夜の学習運用（0:00〜2:00）に合わせた挙動
- 遅延タスクを色で可視化（赤・橙）

---

## 対象シート構成（重要）

シート名: **`毎日のタスク`**

このプロジェクトは **横持ちレイアウト（2列1セット）** 前提です。

- `A1` = 日付（例: `2026/2/22/日`）
- `A2:A` = タスク名
- `B2:B` = 完了チェック（チェックボックス）
- `C1` = 次の日付
- `C2:C` = タスク名
- `D2:D` = 完了チェック
- 以降、`E/F`, `G/H`... と続く

### 例
- A列/B列 = 2/22のタスク
- C列/D列 = 2/23のタスク
- E列/F列 = 2/24のタスク

---

## 主要機能

### 1) 双方向同期（シート ⇄ Google Tasks）
- **シート編集時（onEdit）**:
  - タスク名変更
  - チェックON/OFF
  を即座に Google Tasks に反映
- **定期ポーリング**:
  - Google Tasks 側で完了にした変更をシートに戻す
  - Google Tasks 側で新規作成したタスクをシートに取り込む（`due`があるタスクのみ）

### 2) 繰り越し（未完了タスクの翌日移動）
- 毎日 **2:00** ごろ（トリガー）に実行
- 当日分の未完了タスクを翌日にコピー
- **元のタスクはシートで未完了のまま残す**
- Google Tasks 側では、前日分は **完了扱い** にして見えにくくし、当日分だけを表示状態にする

#### 繰り越しタイトルの表記
- `数学β青チャ51-55（繰越1日目）`
- `数学β青チャ51-55（繰越2日目）`

### 3) 深夜運用（2:00境界）
通常の日付切り替え（0:00）ではなく、**運用上の1日の区切りを 2:00** にしています。

- `0:00〜1:59` は **前日扱い**
- `2:00〜` は当日扱い

この「運用日」は以下に使われます:
- 色付け判定（赤/橙）
- Google Tasks → シートの完了状態反映の対象日判定

### 4) 色付けルール（シート側）
タスク名セル（チェック列ではない）を色分けします。

#### 赤（`COLOR_LATE_OPEN`）
- 繰越タスク（`（繰越N日目）`）で未完了

#### 橙（`COLOR_LATE_DONE`）
- 繰越が発生した系列（元タスク + 繰越タスク群）で系列内のどれかが完了
- 元タスク・繰越タスク両方を橙
- 赤（未完了繰越）があるセルは赤優先

---

## Google Tasks 側の仕様

### 対象タスクリスト
`TASK_SYNC_CONFIG.TASKLIST_TITLE`（例: `マイタスク`）に一致するリストを使用。  
見つからない場合は最初のリストを使用。

### 期日（due）の扱い
Google Tasks の `due` はタイムゾーンずれが起きやすいため、**JSTの正午固定**で設定します。  
例: `2026-02-22T12:00:00+09:00`

### Tasks側で新規作成したタスクの取り込み
- `due` があるタスクのみ取り込み
- 対応日付の列ペアがなければ自動追加
- 同日・同名（または同ベース名）なら重複作成せず紐付け

---

## データ管理（隠しシート）

Task ID は見える列には置かず、隠しシート **`_TaskSyncMap`** に保存します。

### `_TaskSyncMap` の構造
- `A:key`
- `B:taskId`
- `C:updated`
- `D:sheetName`
- `E:row`
- `F:titleCol`
- `G:doneCol`

### key例
`毎日のタスク::r12::t1::d2`

---

## コード構造（TaskSync.gs）

### 公開関数
- `syncDailyTasksBidirectional()`: メイン同期
- `onEditDailyTasksSync(e)`: シート編集時即時同期
- `pollGoogleTasksCompletionToSheet()`: Tasks側変更の定期反映
- `rolloverUncompletedTasksToTomorrow()`: 未完了繰越（2:00）
- `setupDailyTaskSyncTriggers()`: トリガー一括設定
- `removeTaskSyncTriggers_()`: TaskSync系トリガー削除
- `refreshTaskColors()`: 手動色更新

### 主な内部関数
- シート→Tasks: `syncSheetDayToGoogleTasks_`, `syncSingleTaskCellRow_`
- Tasks→シート: `syncGoogleTasksToSheet_`, `importNewGoogleTasksToSheet_`
- 繰越補助: `markTaskAsCompletedInGoogleTasksByCell_`
- 色更新: `refreshTaskColors_`
- 日付列管理: `findPairByDate_`, `createDayPairColumn_`, `getTaskColumnPairs_`
- 文字列/日付 helper: `makeCarryoverTitle_`, `normalizeTaskBaseTitle_`, `getOperationalToday_` など

---

## トリガー構成（推奨）
`setupDailyTaskSyncTriggers()` で以下を作成します。

1. 毎時: `syncDailyTasksBidirectional`
2. 15分ごと: `pollGoogleTasksCompletionToSheet`
3. 毎日2:00ごろ: `rolloverUncompletedTasksToTomorrow`
4. onEdit: `onEditDailyTasksSync`

---

## 設定値（TASK_SYNC_CONFIG）でよく調整する項目
- `SHEET_NAME`
- `TASKLIST_TITLE`
- `TOMORROW_RELEASE_HOUR`
- `DAY_BOUNDARY_HOUR`
- `COLOR_LATE_DONE`, `COLOR_LATE_OPEN`

---

## 運用上の注意
- `due` なしのGoogle Tasksはシートに取り込まれない
- 同名タスクが多いと、色付けで同系列扱いになる可能性あり（ベースタイトル判定）
- Apps Script の高度なGoogleサービス `Tasks API` を有効化しておく必要あり

---

## VSCodeでの編集メモ
### 推奨ファイル分割
- `TaskSync.gs`
- `CalendarSync.gs`
- `Menu.gs`

### 初見で読む順番
1. `syncDailyTasksBidirectional()`
2. `rolloverUncompletedTasksToTomorrow()`
3. `syncSingleTaskCellRow_()`
4. `syncGoogleTasksToSheet_()`
5. `refreshTaskColors_()`

---

## 初期セットアップ
コード更新後、Apps Scriptで以下を一度実行:
1. `syncDailyTasksBidirectional()`
2. `setupDailyTaskSyncTriggers()`

（権限許可あり）
