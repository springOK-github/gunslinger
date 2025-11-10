# Gunslinger - Google Apps Script マッチングシステム

ガンスリンガー方式トーナメントのマッチングシステム。Google Apps Script (GAS) とスプレッドシートで動作するサーバーレスアプリケーション。

## アーキテクチャ概要

**重要**: このシステムはスプレッドシートを UI 兼データベースとして使用する GAS アプリケーション。すべてのコードはスクリプトエディタで実行され、ローカル実行は不可。

### レイヤー構造

- **app.js**: 初期化・設定・排他制御（`onOpen()`, `acquireLock()`, `getMaxTables()`）
- **player-domain.js**: プレイヤー操作・状態管理（`updatePlayerState()`, `registerPlayer()`）
- **match-domain.js**: マッチングロジック（`matchPlayers()`, `recordResult()`）
- **shared.js**: シート操作共通処理（`getSheetStructure()`, `getPlayerName()`）
- **constants.js**: システム定数（シート名、ステータス、ID 規則）
- **test-utils.js**: テストデータ生成（`registerTestPlayers()`）

### データモデル（3 シート構造）

1. **プレイヤー**: マスタデータ（ID, 名前, 勝敗数, 参加状況, 最終対戦日時）
2. **対戦履歴**: 完了した対戦記録（対戦 ID, 日時, 卓番号, 両者 ID, 勝者名）
3. **マッチング**: 進行中の対戦（卓番号, 両者 ID・名前）

## 重要パターンと制約

### 1. 排他制御（最重要）

**全ての書き込み処理でロック必須**。複数ユーザーの同時操作を防ぐ:

```javascript
let lock = null;
try {
  lock = acquireLock("操作名"); // 30秒タイムアウト
  // シート書き込み処理
} catch (e) {
  Logger.log("エラー: " + e.message);
} finally {
  releaseLock(lock); // 必ず解放
}
```

カスタムメニューから呼ばれる UI 関数は直接ロックを取得。内部の実処理関数（`updatePlayerState()` など）は呼び出し元がロックを管理。

### 2. 状態遷移の一元管理

**プレイヤー状態の変更は必ず `updatePlayerState()` 経由**。直接シート更新禁止:

```javascript
updatePlayerState({
  targetPlayerId: "P001",
  newStatus: PLAYER_STATUS.WAITING, // 対象プレイヤーの新状態
  opponentNewStatus: PLAYER_STATUS.WAITING, // 対戦相手の新状態
  recordResult: true, // 対戦結果を履歴に記録
  isTargetWinner: true, // 勝者判定
});
```

状態: `待機` → `対戦中` → `待機` → ... → `終了`（ドロップアウト）

### 3. シート操作の必須パターン

**全てのシート読み込みは `getSheetStructure()` 経由**。列順変更に対応:

```javascript
const { indices, data } = getSheetStructure(sheet, SHEET_PLAYERS);
const playerId = row[indices["プレイヤーID"]]; // 動的インデックス取得
const status = row[indices["参加状況"]];
```

`REQUIRED_HEADERS` 定数で各シートの必須列を定義済み。

### 4. マッチングアルゴリズム（パフォーマンス最適化済み）

`matchPlayers()` の処理フロー:

1. **全データを一括取得**してメモリキャッシュ（シートアクセス最小化）
2. **対戦履歴を `Map<PlayerId, Set<OpponentId>>` で構築**（O(1) 検索）:
   ```javascript
   const opponentsMap = new Map();
   for (let i = 1; i < historyData.length; i++) {
     const p1 = historyData[i][historyIndices["ID1"]];
     const p2 = historyData[i][historyIndices["ID2"]];
     if (!p1 || !p2) continue;
     if (!opponentsMap.has(p1)) opponentsMap.set(p1, new Set());
     if (!opponentsMap.has(p2)) opponentsMap.set(p2, new Set());
     opponentsMap.get(p1).add(p2);
     opponentsMap.get(p2).add(p1);
   }
   ```
3. **待機プレイヤーをソート**: 勝数降順 → 最終対戦日時昇順（先着優先）
4. **再戦回避**: 過去対戦相手を Set でチェック、未対戦者のみマッチング:
   ```javascript
   const p1Opponents = opponentsMap.get(p1Id) || new Set();
   for (let i = 0; i < availablePlayers.length; i++) {
     const p2Id = availablePlayers[i][playerIndices["プレイヤーID"]];
     if (!p1Opponents.has(p2Id)) {
       // マッチング成立
       break;
     }
   }
   ```
5. **卓番号割り当て**: 勝者の前回卓を再利用、なければ新規卓（`getMaxTables()` まで）

**重要**: 過去対戦相手チェックは `matchPlayers()` 内で完結。外部ヘルパー関数なし。

### 5. 自動マッチングトリガー

待機者 2 人以上で以下のタイミングで自動実行:

- `registerPlayer()` 完了後
- `updatePlayerState()` で対戦終了後
- `registerTestPlayers()` 完了後

## 開発ワークフロー

### Clasp セットアップ

```bash
npm install -g @google/clasp
clasp login
clasp clone "スクリプトID"  # GAS プロジェクトから取得
clasp push --watch  # 自動アップロード推奨
```

**注意**: `.clasp.json` は `.gitignore` に追加。各開発者が個別にクローン。

### テスト実行

1. スプレッドシートのカスタムメニュー「⚙️ シートの初期設定」実行
2. スクリプトエディタから `registerTestPlayers()` 実行（`test-utils.js`）
3. `TEST_CONFIG.NUM_PLAYERS` でプレイヤー数調整（デフォルト: 8 人）

**テストシナリオ例**:

```javascript
// 1. 基本的なマッチングテスト
registerTestPlayers(); // 8人登録 → 自動で4組マッチング

// 2. 休憩・復帰のテスト
restPlayer(); // カスタムメニューから実行
returnPlayerFromResting(); // 復帰 → 待機者がいれば自動マッチング

// 3. ドロップアウトのテスト
dropoutPlayer(); // カスタムメニューから実行、対戦中なら相手も待機に戻る

// 4. 対戦結果記録のテスト
promptAndRecordResult(); // カスタムメニューから実行、両者が待機に戻り自動マッチング
```

### デプロイ & デバッグ

- **ログ確認**: Apps Script エディタの「実行ログ」タブ（`Logger.log()` 使用）
- **デプロイ**: `clasp push` で自動反映（Web アプリデプロイ不要）
- **エラー**: `ui.alert()` でユーザー通知、`Logger.log()` で詳細記録

## プロジェクト固有規約

### ID・命名規則

- プレイヤー ID: `P` + 3 桁数字（例: `P001`）← `PLAYER_ID_PREFIX` + `ID_DIGITS`
- 対戦 ID: `T` + 4 桁数字（例: `T0001`）
- 卓番号: 1 ～ 200（`getMaxTables()` で動的取得、デフォルト: 50）
- UI 入力: 数字のみ（例: ユーザーが「1」入力 → システムで「P001」整形）

### コミット規約

[Conventional Commits](https://www.conventionalcommits.org/ja/) 準拠、**日本語メッセージ**:

- `feat:` 新機能
- `fix:` バグ修正
- `refactor:` リファクタリング
- `perf:` パフォーマンス改善

### ロケール設定

- タイムゾーン: `Asia/Tokyo`（`appsscript.json`）
- 日時フォーマット: `yyyy/MM/dd HH:mm:ss`（`Utilities.formatDate()`）
- `parseInt()` 使用時は必ず基数 `10` 指定

## コードレビューチェックリスト

- [ ] ロック取得・解放が適切（try-finally）
- [ ] 状態変更は `updatePlayerState()` 経由
- [ ] シート操作は `getSheetStructure()` 使用
- [ ] `Logger.log()` でエラー詳細記録
- [ ] パフォーマンス影響考慮（シートアクセス最小化）

## 参考リンク

- 使用方法: [README.md](../README.md)
- GAS API: https://developers.google.com/apps-script
