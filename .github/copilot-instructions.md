# Gunslinger - Tournament Matching System

**開発者向けドキュメント**

このドキュメントはシステムの内部構造と開発ワークフローを説明します。
使用方法については [README.md](../README.md) を参照してください。

---

## アーキテクチャ概要

Google Apps Script (GAS) ベースのガンスリンガー方式マッチングシステム。スプレッドシートを UI 兼データベースとして使用。

### コアコンポーネント（AI エージェント視点の補足付き）

#### ドメイン層

- **player-domain.js**: プレイヤードメイン

  - プレイヤー操作: 登録・休憩・復帰・ドロップアウト
  - データ取得・検索: 待機者リスト、過去対戦相手、プレイヤー名取得
  - 統計更新: 勝敗数・試合数の更新
  - 状態遷移: `updatePlayerState()` による一元管理（ロック取得と履歴更新までカバーするため直接シートを触らない）

- **match-domain.js**: 対戦ドメイン
  - マッチング管理: `matchPlayers()` - 再戦回避、勝者優先、パフォーマンス最適化済み
  - 対戦結果記録: `recordResult()` - 履歴記録、統計更新
  - 対戦結果修正: `correctMatchResult()` - 誤記録の修正

#### 共通層

- **shared.js**: 共有ユーティリティ
  - シート操作: `getSheetStructure()`, `getPlayerName()`, 卓番号管理
  - UI 共通処理: `promptPlayerId()`, `changePlayerStatus()`
  - 新規ヘルパー（実装差分）:
    - `formatElapsedMs(elapsedMs)` — 経過ミリ秒を HH:mm:ss 形式で返すフォーマッタ。GAS の Date/Utilities.formatDate を経由した表示で生じるタイムゾーンズレを避ける目的で導入しています（経過時間はミリ秒差から直接算出することを基本にしています）。
    - `getMaxUsedTableNumber()` — マッチングシート上で現在使用中の最大卓番号を返します。`configureMaxTables()` の検証で利用されています。
    - `getNextAvailableTableNumber(inProgressSheet)` — 使用中の卓番号をスキャンして空き番号を返す既存ロジックの補助。

#### アプリケーション層

- **app.js**: アプリケーション層
  - システム初期化: `onOpen()` - カスタムメニュー、`setupSheets()` - シート作成
  - システム設定: `getMaxTables()`, `setMaxTables()` - 最大卓数管理
  - 排他制御: `acquireLock()`, `releaseLock()` - 同時操作防止

#### その他

- **constants.js**: システム定数（シート名、ステータス、卓設定）
- **test-utils.js**: テストデータ生成

### 3 つのシート構造（列ヘッダーは `constants.js` の `REQUIRED_HEADERS` で定義）

1. **プレイヤー**: プレイヤーマスタ（ID、名前、勝敗数、参加状況、最終対戦時刻）
2. **対戦履歴**: 完了した対戦の記録（時刻、卓番号、両者 ID、対戦 ID、対戦時間）。`勝者名` 列に勝者、`敗者名` 列に敗者を保存します。
3. **マッチング**: 進行中の対戦（卓番号、両者 ID・名前、対戦開始時刻、経過時間）

## 重要な設計パターン

### 状態遷移フロー

プレイヤーは 3 つの状態を遷移:

- `待機` → `対戦中` → `待機` (通常フロー)
- `待機`/`対戦中` → `終了` (ドロップアウト)

**すべての状態変更は `updatePlayerState()` 経由で実行（AI エージェントでの手動シート更新禁止）**:

```javascript
updatePlayerState({
  targetPlayerId: "P001",
  newStatus: PLAYER_STATUS.WAITING,
  opponentNewStatus: PLAYER_STATUS.WAITING,
  recordResult: true, // 対戦結果記録フラグ
  isTargetWinner: true,
});
```

### ロック機構の必須パターン（AI 向け補足）

複数ユーザーの同時操作を防ぐため、**すべての書き込み処理でロック取得が必須**:

```javascript
let lock = null;
try {
  lock = acquireLock("操作名");
  // シート書き込み処理（ここでのみ SpreadsheetApp を更新）
} catch (e) {
  Logger.log("エラー: " + e.message);
} finally {
  releaseLock(lock);
}
```

**デッドロック回避**: `player-domain.js` では複数ロックを固定順序で取得（状態変更 → 対戦結果）。

### マッチングアルゴリズム（詳細フロー）

1. **待機中プレイヤーのソート優先順位**:
   - 勝数が多い順（降順）
   - 最終対戦時刻が古い順（昇順 = 先着優先）
2. **再戦回避**:

- 対戦履歴を Map/Set でキャッシュし、O(1) で過去対戦相手をチェック（`opponentsMap`）
- 未対戦相手のみマッチング
- 全員が過去対戦者の場合は**マッチングを成立させず待機継続**

3. **卓番号の割り当て**:

   - 勝者の前回使用卓が空いていれば再利用（`getLastTableNumber()`）

- 空きがなければ新規卓を `getNextAvailableTableNumber()` で取得（`getMaxTables()` の範囲内で調整）

4. **パフォーマンス最適化**:
   - 全データを一括取得してキャッシュ（シートアクセス最小化）
   - プレイヤー名を Map でキャッシュ（O(1) 取得）
   - 対戦履歴を Map<PlayerId, Set<OpponentId>> で構築（O(1) 検索）
   - インラインソートで中間関数呼び出しを削減

### 自動マッチングトリガー

以下のタイミングで待機者が 2 人以上いると自動実行（AI が新処理を追加する場合もこのトリガーを維持）:

- プレイヤー登録完了後（`registerPlayer()`）
- 対戦結果記録後（`updatePlayerState()` → 最後に `matchPlayers()` 呼び出し）
- テストプレイヤー登録後（`registerTestPlayers()`）

追加の実装差分（経過時間更新）:

- `updateAllMatchTimes()` — マッチングシート上の対戦開始時刻からミリ秒差を算出し、`formatElapsedMs()` を使って `経過時間` 列を HH:mm:ss 表示で更新します。従来の Date を経由したフォーマットで時刻ズレが出ていたため、ミリ秒 → 文字列のヘルパーを導入しました。
- メニュー/トリガー関係: `setupMatchTimeUpdaterTrigger()` と `deleteMatchTimeUpdaterTrigger()` を `app.js` に追加し、1 分毎に `updateAllMatchTimes()` を実行するトリガーの開始/停止をサポートしています（手動で開始/停止する運用を想定）。

### 大会開始 / 終了と運用フラグ

- `app.js` に `startTournament()`（以前の setupSheets 相当）と `endTournament()` を追加しました。`startTournament()` はシート初期化とタイムゾーン設定を行い、`endTournament()` は対戦履歴の日時付きバックアップを作成します。
- `endTournament()` 実行時は進行中の対戦があれば `endAllActiveMatches()` を呼んで強制終了し、プレイヤーを待機状態に戻します（進行中の対戦は履歴に記録されません）。
- 保守／強制終了中に自動マッチングが走らないよう、`MAINTENANCE_MODE` フラグを PropertiesService に保存しておき、`matchPlayers()` の先頭でチェックしてスキップする仕組みを導入しました。

### データ構造の検証パターン

すべてのシート操作は `getSheetStructure()` 経由でヘッダー検証（列追加時は `REQUIRED_HEADERS` を先に更新）:

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

### テスト実行（自動マッチング確認手順）

スプレッドシート上でカスタムメニューから実行:

1. 「シートの初期設定」でシート構造を作成
2. スクリプトエディタから `registerTestPlayers()` を実行してテストデータ生成

## プロジェクト固有の規約と運用 Tips

### 命名規則

- プレイヤー ID: `P` + 3 桁数字（例: `P001`）- `ID_DIGITS` 定数で制御
- 対戦 ID: `T` + 4 桁数字（例: `T0001`）
- 卓番号: 1 ～ 200（`getMaxTables()` で動的取得、デフォルト: 50）
- ID 採番時は既存最大値を走査し `Utilities.formatString` で埋め桁

### UI 入力規則

ユーザーインターフェースでは**数字部分のみ**入力を要求:

```javascript
// ユーザーが「1」と入力 → システムで「P001」に整形
const playerId = PLAYER_ID_PREFIX + Utilities.formatString(`%0${ID_DIGITS}d`, parseInt(rawId, 10));
```

### ロケール設定

- タイムゾーン: `Asia/Tokyo`（`appsscript.json`）
- 時刻フォーマット: `HH:mm:ss`

### ログとエラーハンドリング

- すべての catch ブロックで `Logger.log()` にエラー詳細を記録
- ユーザーには `ui.alert()` で簡潔なメッセージを表示（AI の自動処理でも UI メッセージ整合性を保つ）
- データ不整合検出時は警告ログを出力しつつ処理継続

## 重要な注意事項

1. **ロック取得順序の厳守**: 複数ロック取得時はデッドロック防止のため順序固定（`updatePlayerState` では状態 → 対戦結果）。
2. **状態遷移は `updatePlayerState()` 経由**: 直接シート更新しない。AI がバッチ処理を追加する場合もこの関数を経由。
3. **カスタムメニュー関数はロック管理不要**: 内部の実処理関数でロック取得。AI が新しいメニュー処理を追加する際も同じ責務分割を守る。
4. **数値の型変換**: `parseInt()` 使用時は必ず基数 `10` を指定。ユーザー入力を扱うときは正規表現チェックも追加。
5. **パフォーマンス最適化の維持**: `matchPlayers()` は Map/Set キャッシュで最適化済み。過去対戦相手のチェックは関数内で完結しており、外部ヘルパー関数は使用しない。
6. **Apps Script 制限の考慮**: 1 回の処理でアクセスするシート回数を最小化。ループ内で `getRange`/`setValue` を繰り返さない。

### 実装時の静的解析と注意点（開発者向け補足）

- 一部の変更でエディタや静的解析（型チェック）が "Sheet | null" や JSDoc の型不一致を指摘するケースが増えています。これは `getSheetByName()` の戻り値が `null` になる可能性や、LockService 型が JSDoc で正しく解釈されないためです。
- 開発時の推奨対応:
  - `getSheetStructure()` を呼ぶ前提でシート存在を検証するか、呼び出し側で適切に null チェックを行ってください。現状の実装は「シートが存在する前提」で例外を投げることで早期検出する設計になっています。
  - JSDoc の LockService 型や外部 API の型はエディタ設定により警告が出ることがあるため、必要に応じてコメントで型を明示してください（例: `/** @type {LockService.Lock} */`）。
  - `formatElapsedMs()` のように表示ロジックは共通化しておくと、Date/タイムゾーンの問題を局所化できます。

## 外部依存関係

- Google Apps Script V8 ランタイム
- SpreadsheetApp サービス
- LockService（排他制御）
- Utilities（時刻フォーマット、文字列整形）

## 貢献ガイド

### Pull Request の提出

1. このリポジトリをフォーク
2. 機能ブランチを作成 (`git checkout -b feature/amazing-feature`)
3. 変更をコミット (`git commit -m 'feat: 素晴らしい機能を追加'`)
4. ブランチにプッシュ (`git push origin feature/amazing-feature`)
5. Pull Request を作成

### コミットメッセージ規約

Conventional Commits に準拠:

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
- GAS API ドキュメント: <https://developers.google.com/apps-script>
