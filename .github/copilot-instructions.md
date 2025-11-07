# Gunslinger - Pokémon Card Tournament Matching System

## アーキテクチャ概要

Google Apps Script (GAS) ベースのポケモンカードバトル大会マッチングシステム。スプレッドシートを UI 兼データベースとして使用。

### コアコンポーネント

- **constants.js**: システム全体の定数定義（シート名、ステータス、卓設定）
- **setup.js**: 初期化とカスタムメニュー（`onOpen()` トリガー）
- **config.js**: システム設定の管理（最大卓数の取得・設定）
- **match-manager.js**: マッチングロジックと対戦結果記録
- **player-operations.js**: プレイヤーのUI操作（登録・休憩・復帰・ドロップアウト）
- **player-queries.js**: プレイヤーデータの取得と検索
- **player-stats.js**: プレイヤー統計の更新処理
- **player-state.js**: 状態遷移の共通処理（`updatePlayerState()` が中核）
- **ui-utils.js**: UI操作の共通ユーティリティ
- **sheet-utils.js**: スプレッドシート操作の抽象化
- **lock-utils.js**: 排他制御（`acquireLock()` / `releaseLock()`）
- **test-utils.js**: テストデータ生成

### 3つのシート構造

1. **プレイヤー**: プレイヤーマスタ（ID、名前、勝敗数、参加状況、最終対戦日時）
2. **対戦履歴**: 完了した対戦の記録（日時、卓番号、両者ID、勝者名、対戦ID）
3. **マッチング**: 進行中の対戦（卓番号、両者ID・名前） - **卓番号は削除せず再利用**

## 重要な設計パターン

### 状態遷移フロー

プレイヤーは 3 つの状態を遷移:
- `待機` → `対戦中` → `待機` (通常フロー)
- `待機`/`対戦中` → `終了` (ドロップアウト)

**すべての状態変更は `updatePlayerState()` 経由で実行**:
```javascript
updatePlayerState({
  targetPlayerId: "P001",
  newStatus: PLAYER_STATUS.WAITING,
  opponentNewStatus: PLAYER_STATUS.WAITING,
  recordResult: true,  // 対戦結果記録フラグ
  isTargetWinner: true
});
```

### ロック機構の必須パターン

複数ユーザーの同時操作を防ぐため、**すべての書き込み処理でロック取得が必須**:

```javascript
let lock = null;
try {
  lock = acquireLock('操作名');
  // 処理内容
} catch (e) {
  Logger.log("エラー: " + e.message);
} finally {
  releaseLock(lock);
}
```

**デッドロック回避**: `player-state.js` では複数ロックを固定順序で取得（状態変更→対戦結果）。

### マッチングアルゴリズム

1. **待機中プレイヤーのソート優先順位**:
   - 勝数が多い順（降順）
   - 最終対戦日時が新しい順（直近の勝者を優先）
   
2. **再戦回避**: `getPastOpponents()` で過去対戦相手を取得し、未対戦相手のみマッチング
   - 全員が過去対戦者の場合は**マッチングを成立させず待機継続**

3. **卓番号の割り当て**:
   - 勝者の前回使用卓が空いていれば再利用（`getLastTableNumber()`）
   - 空きがなければ新規卓を `getNextAvailableTableNumber()` で取得
   - 対戦終了後も卓番号行は削除せず、ID/名前のみクリア

### 自動マッチングトリガー

以下のタイミングで待機者が 2 人以上いると自動実行:
- プレイヤー登録完了後（`registerPlayer()`）
- 対戦結果記録後（`updatePlayerState()` → 最後に `matchPlayers()` 呼び出し）
- テストプレイヤー登録後（`registerTestPlayers()`）

### データ構造の検証パターン

すべてのシート操作は `getSheetStructure()` 経由でヘッダー検証:
```javascript
const { indices, data } = getSheetStructure(sheet, SHEET_PLAYERS);
const playerId = row[indices["プレイヤーID"]];
```
- `REQUIRED_HEADERS` 定数で必須列を定義
- 列インデックスを動的に取得（列順変更に対応）

## 開発ワークフロー

### ローカル開発環境

```bash
# Clasp インストール & ログイン
npm install -g @google/clasp
clasp login

# プロジェクトクローン
clasp clone "スクリプトID"

# 自動アップロード（推奨）
clasp push --watch

# 手動アップロード
clasp push
```

### Git + Clasp 並行運用

- `.clasp.json` は `.gitignore` に追加（各開発者が個別に clasp clone）
- コミットメッセージは**日本語**で記述
- `clasp push` と `git push` を併用して GAS とリポジトリを同期

### テスト実行

スプレッドシート上でカスタムメニューから実行:
1. 「シートの初期設定」でシート構造を作成
2. スクリプトエディタから `registerTestPlayers()` を実行してテストデータ生成

## プロジェクト固有の規約

### 命名規則

- プレイヤーID: `P` + 3桁数字（例: `P001`）- `ID_DIGITS` 定数で制御
- 対戦ID: `T` + 4桁数字（例: `T0001`）
- 卓番号: 1～50（`TABLE_CONFIG.MAX_TABLES` で制御）

### UI 入力規則

ユーザーインターフェースでは**数字部分のみ**入力を要求:
```javascript
// ユーザーが「1」と入力 → システムで「P001」に整形
const playerId = PLAYER_ID_PREFIX + 
  Utilities.formatString(`%0${ID_DIGITS}d`, parseInt(rawId, 10));
```

### ロケール設定

- タイムゾーン: `Asia/Tokyo`（`appsscript.json`）
- 日時フォーマット: `yyyy/MM/dd HH:mm:ss`

### ログとエラーハンドリング

- すべての catch ブロックで `Logger.log()` にエラー詳細を記録
- ユーザーには `ui.alert()` で簡潔なメッセージを表示
- データ不整合検出時は警告ログを出力しつつ処理継続

## 重要な注意事項

1. **卓番号行は削除禁止**: `cleanUpInProgressSheet()` は意図的に何もしない
2. **ロック取得順序の厳守**: 複数ロック取得時はデッドロック防止のため順序固定
3. **状態遷移は `updatePlayerState()` 経由**: 直接シート更新しない
4. **カスタムメニュー関数はロック管理不要**: 内部の実処理関数でロック取得
5. **数値の型変換**: `parseInt()` 使用時は必ず基数 `10` を指定

## 外部依存関係

- Google Apps Script V8 ランタイム
- SpreadsheetApp サービス
- LockService（排他制御）
- Utilities（日時フォーマット、文字列整形）

## 参考リソース

- 詳細な使用方法: `README.md`
- テスト関数例: `test-utils.js`
- GAS API ドキュメント: https://developers.google.com/apps-script
