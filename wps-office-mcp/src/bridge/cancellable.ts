/**
 * Input: 动作名
 * Output: 该动作能否被安全取消
 * Pos: 取消机制的安全红线。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 *
 * 取消的两种时机（差别很大，别混为一谈）
 * --------------------------------------
 *   A. 【排队中】请求还没写进 PowerShell —— 任何动作都能取消，删掉队列项即可。
 *      请求从未执行，既不产生副作用，也不需要重放。这是最有用的一档。
 *   B. 【已在途】请求已经写进去，WPS 正在跑 —— **取消不了**。
 *      COM 是 STA，PowerShell 那边正在等一个 COM 调用返回，我们没有安全的中断手段
 *      （杀进程会让 WPS 处于半完成状态，那正是 wps_common_repair 要收拾的烂摊子）。
 *      所以这一档能做的只有【放弃等待】：调用方立刻返回，子进程继续跑完，
 *      回话到了直接丢掉。
 *
 * B 档为什么还要分只读 / 写入
 * ---------------------------
 *   只读动作放弃等待：没有副作用，丢掉回话零风险。
 *   写入动作放弃等待：调用方会以为"没做成"，实际它【可能已经做成了一半】
 *   —— 于是要么重放（变成执行两次），要么以为失败却已经生效。两种都比等着更糟。
 *   所以写入动作【拒绝放弃】，让它老实跑完、如实返回结果。
 *
 * 一句话：能取消的只有"还没开始"的；已经开始的一律不打断，只读的可以不等，写入的不许不等。
 */

/** 取消动作时抛这个。与业务失败区分开 —— 它不是 WPS 出错，是我们主动不干了。 */
export class CancelError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'CancelError';
  }
}

/** 一眼就是只读的动作前缀 */
const READ_ONLY_PREFIXES = ['get', 'list', 'read', 'find', 'query', 'probe', 'ping', 'check', 'caps', 'diagnose'];

/** 名字看不出来但实际只读的（探测类内建动作 + 少数历史命名） */
const READ_ONLY_EXTRA = new Set([
  '__probeParams',
  '__probeDoc',
  '__probeCaps',
  '__probeDialogs',
  '__probeStyleAdd',
  '__probeClose',
  // 时序诊断动作：只是占住通道，无任何副作用（配合 __sleep 对照，见 scripts/wps-com.ps1）
  '__sleepRead',
  '__bridgeStatus',
  'wpsStatus',
  'getAppInfo',
  'getOpenDocuments',
  'getOpenWorkbooks',
  'getOpenPresentations',
  'proofreadBasic',
  'evaluateFormula',
  'diagnoseFormula',
]);

/** 这些前缀【一定】是写入，优先级高于上面的猜测（防止 getXxx 里混进写操作） */
const WRITE_PREFIXES = ['set', 'add', 'insert', 'delete', 'remove', 'update', 'modify', 'save', 'create', 'close', 'quit', 'convert', 'apply', 'merge', 'split', 'move', 'copy', 'paste', 'replace', 'clear', 'hide', 'show', 'protect', 'unprotect', 'sort', 'fill', 'group', 'mark', 'accept', 'reject', 'enable', 'disable', 'start', 'switch', 'open', 'export', 'import', 'rename', 'rotate', 'duplicate', 'wrap', 'lock'];

/**
 * 该动作是不是只读。
 * 判定顺序：显式白名单 → 写入前缀（一票否决）→ 只读前缀 → 默认【否】。
 * 默认否是刻意的：判不出来就按"不能安全放弃等待"处理，宁可让调用方多等一会儿。
 */
export function isReadOnlyAction(action: string): boolean {
  if (!action) return false;
  if (READ_ONLY_EXTRA.has(action)) return true;

  const a = action.charAt(0).toLowerCase() + action.slice(1);
  for (const p of WRITE_PREFIXES) {
    if (a.startsWith(p)) return false;
  }
  for (const p of READ_ONLY_PREFIXES) {
    if (a.startsWith(p)) return true;
  }
  return false;
}
