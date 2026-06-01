# tasks_gas

Google スプレッドシート ⇄ Google Tasks ⇄ Google カレンダー の同期 + Claude による予定生成。
プロジェクトの仕様詳細は [`info.md`](./info.md) を参照。

---

## 構成ファイル

| ファイル | 役割 |
| --- | --- |
| `appsscript.json` | Apps Script マニフェスト（タイムゾーン・有効化サービス） |
| `menu.js` | スプレッドシートのメニューバー定義 |
| `calendarsync.js` | 「週予定」シート → Google カレンダー同期 |
| `taskssync.js` | 「毎日のタスク」シート ⇄ Google Tasks 双方向同期 |
| `planner.js` | AI 週次プランナー — `全タスク` を読み、自然言語指示から `毎日のタスク` と `週予定` を一括生成 |
| `plannerDialog.html` | AI 週次プランナーの入力ダイアログ |
| `info.md` | 設計ドキュメント |
| `README.md` | このファイル（セットアップ手順） |

---

## 初回セットアップ（clasp の導入）

ローカル（VSCode / GitHub）と Apps Script を `clasp` で同期します。

### 1. Node.js と clasp をインストール

```bash
# Node.js（未インストールの場合は https://nodejs.org/ から LTS）
node --version   # v18 以上推奨

# clasp をプロジェクトに導入（package.json に記載済み）
npm install
```

> グローバルインストールでも可: `npm install -g @google/clasp`

### 2. Google アカウントで clasp にログイン

```bash
npm run login
# ブラウザが開いて Google ログイン → 許可
```

ログイン情報は `~/.clasprc.json` に保存されます（gitignore 済み）。

### 3. Apps Script API を有効化

ブラウザで以下を開き、トグルを ON：
<https://script.google.com/home/usersettings>

### 4. `.clasp.json` を作成（Script ID を設定）

`.clasp.json.example` をコピーして `.clasp.json` を作る:

```bash
cp .clasp.json.example .clasp.json
```

Script ID の確認方法：
1. 対象のスプレッドシートを開く
2. 「拡張機能」→「Apps Script」
3. プロジェクトの URL `https://script.google.com/d/【ここがScriptID】/edit` の中央部分

`.clasp.json` の `scriptId` に貼り付け：

```json
{
  "scriptId": "1abcDEFghIJklmNOpqRStuvwxYZ_0123456789-abcdEFGhij",
  "rootDir": "."
}
```

### 5. GAS から最新版をプル

**重要**: ローカルファイルが上書きされるので、未コミットの変更があれば事前にコミットまたは退避してください。

```bash
npm run pull
```

差分を確認:

```bash
git status
git diff
```

### 6. （Claude プランナーを使う場合）API キーを登録

GAS エディタで以下を一度だけ実行:

```javascript
function setAnthropicApiKey() {
  PropertiesService.getScriptProperties().setProperty(
    'ANTHROPIC_API_KEY',
    'sk-ant-api03-...'  // ここに実際のキー
  );
}
```

実行後、関数本体からキー文字列を削除して保存。
キーは GAS の Script Properties に保存され、ソースコードには残りません。

---

## 日常運用

### コードを変更したとき

```bash
# ローカル → GAS に反映
npm run push

# GAS で誰かが直接編集 → ローカルに取り込み
npm run pull
```

### GAS エディタを開く

```bash
npm run open
```

### ログをリアルタイムで見る

```bash
npm run logs
```

### Git に保存

```bash
git add .
git commit -m "feat: ..."
git push
```

---

## トリガー設定

GAS エディタで以下を一度だけ実行：

- `setupDailyTaskSyncTriggers()` — Tasks 同期系トリガー
- （カレンダー側は手動 / メニュー実行を想定）

---

## トラブルシューティング

| 症状 | 対処 |
| --- | --- |
| `clasp push` で "User has not enabled the Apps Script API" | 手順 3 を再確認 |
| `clasp pull` 後にファイル名が変わる | GAS 側のファイル名と一致させているため正常。`.gs` は `.js` で扱われます |
| 同じ関数が複数定義されている警告 | JS は最後の定義が優先される。古い定義を削除 |
| Claude API がタイムアウト | `planner.js` の `MAX_TOKENS` を減らす、または `MODEL` を `claude-haiku-4-5` に変更 |
