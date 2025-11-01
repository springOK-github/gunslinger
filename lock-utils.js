/**
 * ポケモンカード・ガンスリンガーバトル用マッチングシステム
 * @fileoverview スプレッドシートのロック機構
 * @author SpringOK
 */

// ロックの最大待機時間（ミリ秒）
const LOCK_TIMEOUT = 30000; // 30秒

/**
 * スプレッドシートの排他ロックを取得します。
 * @param {string} lockName - ロックの名前（操作の種類を識別）
 * @returns {LockService.Lock} 取得したロック
 * @throws {Error} ロックが取得できない場合
 */
function acquireLock(lockName) {
  const lock = LockService.getScriptLock();
  const success = lock.tryLock(LOCK_TIMEOUT);
  
  if (!success) {
    throw new Error(
      '他のユーザーが操作中です。\n' +
      'しばらく待ってから再度お試しください。\n' +
      `(${lockName})`
    );
  }
  
  return lock;
}

/**
 * ロックを解放します。
 * @param {LockService.Lock} lock - 解放するロック
 */
function releaseLock(lock) {
  if (lock) {
    try {
      lock.releaseLock();
    } catch (e) {
      Logger.log('ロックの解放に失敗: ' + e.toString());
    }
  }
}