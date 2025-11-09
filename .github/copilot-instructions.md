# Gunslinger - Pokémon Card Tournament Matching System

**開発者向けドキュメント**

このドキュメントはシステムの内部構造と開発ワークフローを説明します。
使用方法については [README.md](../README.md) を参照してください。

---

## アーキテクチャ概要

Google Apps Script (GAS) ベースのガンスリンガー方式マッチングシステム。スプレッドシートを UI 兼データベースとして使用。

### コアコンポーネント

#### ドメイン層
- **player-domain.js**: プレイヤードメイン
  - プレイヤー操作: 登録・休憩・復帰・ドロップアウト
  - データ取得・検索: 待機者リスト、過去対戦相手、プレイヤー名取得
  - 統計更新: 勝敗数・試合数の更新
  - 状態遷移: `updatePlayerState()` による一元管理
  
- **match-domain.js**: 対戦ドメイン
  - マッチング管理: `matchPlayers()` - 再戦回避、勝者優先、パフォーマンス最適化済み
  - 対戦結果記録: `recordResult()` - 履歴記録、統計更新
  - 対戦結果修正: `correctMatchResult()` - 誤記録の修正

#### 共通層
- **shared.js**: 共有ユーティリティ
  - シート操作: `getSheetStructure()`, `getPlayerName()`, 卓番号管理
  - UI共通処理: `promptPlayerId()`, `changePlayerStatus()`

#### アプリケーション層
- **app.js**: アプリケーション層
  - システム初期化: `onOpen()` - カスタムメニュー、`setupSheets()` - シート作成
  - システム設定: `getMaxTables()`, `setMaxTables()` - 最大卓数管理
  - 排他制御: `acquireLock()`, `releaseLock()` - 同時操作防止

#### その他
- **constants.js**: システム定数（シート名、ステータス、卓設定）
- **test-utils.js**: テストデータ生成

### 3つのシート構造

1. **プレイヤー**: プレイヤーマスタ（ID、名前、勝敗数、参加状況、最終対戦日時）
2. **対戦履歴**: 完了した対戦の記録（日時、卓番号、両者ID、勝者名、対戦ID）
3. **マッチング**: 進行中の対戦（卓番号、両者ID・名前）

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

**デッドロック回避**: `player-domain.js` では複数ロックを固定順序で取得（状態変更→対戦結果）。

### マッチングアルゴリズム

1. **待機中プレイヤーのソート優先順位**:
   - 勝数が多い順（降順）
   - 最終対戦日時が古い順（昇順 = 先着優先）
   
2. **再戦回避**:
   - 対戦履歴を Map/Set でキャッシュし、O(1) で過去対戦相手をチェック
   - 未対戦相手のみマッチング
   - 全員が過去対戦者の場合は**マッチングを成立させず待機継続**

3. **卓番号の割り当て**:
   - 勝者の前回使用卓が空いていれば再利用（`getLastTableNumber()`）
   - 空きがなければ新規卓を `getNextAvailableTableNumber()` で取得

4. **パフォーマンス最適化**:
   - 全データを一括取得してキャッシュ（シートアクセス最小化）
   - プレイヤー名を Map でキャッシュ（O(1) 取得）
   - 対戦履歴を Map<PlayerId, Set<OpponentId>> で構築（O(1) 検索）
   - インラインソートで中間関数呼び出しを削減

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
- 卓番号: 1～200（`getMaxTables()` で動的取得、デフォルト: 50）

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

1. **ロック取得順序の厳守**: 複数ロック取得時はデッドロック防止のため順序固定
2. **状態遷移は `updatePlayerState()` 経由**: 直接シート更新しない
3. **カスタムメニュー関数はロック管理不要**: 内部の実処理関数でロック取得
4. **数値の型変換**: `parseInt()` 使用時は必ず基数 `10` を指定
5. **パフォーマンス最適化の維持**: `matchPlayers()` は Map/Set キャッシュで最適化済み。過去対戦相手のチェックは関数内で完結しており、外部ヘルパー関数は使用しない。

## 外部依存関係

- Google Apps Script V8 ランタイム
- SpreadsheetApp サービス
- LockService（排他制御）
- Utilities（日時フォーマット、文字列整形）

## 貢献ガイド

### Pull Request の提出

1. このリポジトリをフォーク
2. 機能ブランチを作成 (`git checkout -b feature/amazing-feature`)
3. 変更をコミット (`git commit -m 'feat: 素晴らしい機能を追加'`)
4. ブランチにプッシュ (`git push origin feature/amazing-feature`)
5. Pull Request を作成

### コミットメッセージ規約

[Conventional Commits](https://www.conventionalcommits.org/ja/) に準拠:

- `feat:` 新機能
- `fix:` バグ修正
- `docs:` ドキュメントのみの変更
- `refactor:` リファクタリング
- `perf:` パフォーマンス改善
- `test:` テスト追加・修正

### コードレビューのポイント

- ✅ ロック機構が適切に使用されているか
- ✅ エラーハンドリングが適切か
- ✅ ログ出力が適切か
- ✅ パフォーマンスへの影響を考慮しているか
- ✅ ドキュメントが更新されているか

## 参考リソース

- 詳細な使用方法: `README.md`
- テスト関数例: `test-utils.js`
- GAS API ドキュメント: https://developers.google.com/apps-script
