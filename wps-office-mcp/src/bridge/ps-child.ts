/**
 * Input: 应用键（excel / word / ppt / common）与动作名、参数
 * Output: 该动作的 COM 结果（{ success, data?, error? }）
 * Pos: 桥进程内部的 PowerShell 子进程管理器 —— 一个应用一个实例。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么按应用拆进程（per-app 隔离）
 * ----------------------------------
 * COM 是 STA：一个进程同一时刻只能跑一个动作。改造前所有应用共用一个 powershell，
 * 于是"WPS 文字卡在一个慢动作上"会把"表格"的调用一起堵死。
 * 拆成一应用一进程后，某一路卡死只影响它自己；桥还能【只杀那一路】并重建，
 * 另外两路完全不受影响。对标 JulianPoleszczuk/office-mcp 的 per-app 连接模型。
 *
 * 超时语义
 * --------
 * 超时即判定【该应用的这一路已死】：杀掉该子进程、拒绝它在途的请求，下次调用重建。
 * 绝不能"超时了还继续往这个进程塞请求"。
 * 另外：超时后【不自动重放】—— 写动作（保存/导出）可能已经执行了一半，
 * 重放会变成执行两次。由调用方决定怎么办。
 *
 * 子进程生命周期
 * --------------
 * wps-comd.ps1 自带父进程看门狗，本进程（桥）一死，子进程会自己走 —— 不需要额外清理。
 */

import { ChildProcess, spawn } from 'child_process';
import * as path from 'path';
import { v4 as uuidv4 } from 'uuid';

/** 常驻分发脚本（随包分发，见 .npmignore） */
const PS_SCRIPT = path.join(__dirname, '../../scripts/wps-comd.ps1');

/** 起进程 + 握手就绪的上限。冷启动实测约 5 秒，给足余量 */
const READY_TIMEOUT_MS = 30000;

/**
 * 单次动作上限。COM 动作可能很慢（大文档导出、首次插图表要拉起 Excel），默认 120 秒。
 * 每次调用时重读环境变量，便于按场景临时收紧/放宽，不用重启桥。
 */
function callTimeoutMs(): number {
  const v = parseInt(process.env.WPS_DAEMON_TIMEOUT || '120000', 10);
  return Number.isFinite(v) && v > 0 ? v : 120000;
}

/**
 * 每应用的排队上限。超了立刻回 busy —— 把"无声地等几分钟"换成"明确被拒"。
 * 注意：这不等于并发度（COM 是 STA，在途恒为 1），它限的是"等待位"的数量。
 */
export function queueMax(): number {
  const v = parseInt(process.env.WPS_QUEUE_MAX || '4', 10);
  return Number.isFinite(v) && v > 0 ? v : 4;
}

/** 排队上限。排太久不如早拒 —— 与动作超时是两个独立维度。 */
export function queueTimeoutMs(): number {
  const v = parseInt(process.env.WPS_QUEUE_TIMEOUT || '10000', 10);
  return Number.isFinite(v) && v > 0 ? v : 10000;
}

/**
 * 需要「插队」的动作：会开/关/另存文档、切目标、退出 WPS，或弹出对话框。
 * 这些是用户主动发起的，等待位里已经排了一堆读操作时，让它们先走。
 * 饥饿风险可控 —— 这类动作很稀疏，且跳过后队列恢复 FIFO。
 * 可用 WPS_EXCLUSIVE_ACTIONS="a,b,c" 覆盖。
 */
const DEFAULT_EXCLUSIVE_ACTIONS = [
  'openFile',
  'openDocument',
  'openWorkbook',
  'openPresentation',
  'closeDocument',
  'closeWorkbook',
  'closePresentation',
  'switchDocument',
  'switchWorkbook',
  'switchPresentation',
  'createDocument',
  'createWorkbook',
  'createPresentation',
  'save',
  'saveAs',
  'convertToPDF',
  'convertFormat',
  'quitWps',
];

export function exclusiveSet(): Set<string> {
  const override = process.env.WPS_EXCLUSIVE_ACTIONS;
  const list = override
    ? override.split(',').map((s) => s.trim()).filter(Boolean)
    : DEFAULT_EXCLUSIVE_ACTIONS;
  return new Set(list);
}

/** 子进程压根没起来（请求从未送达）时抛这个，调用方可以安全重建。 */
export class ChildUnavailableError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ChildUnavailableError';
  }
}

/** 队列满 / 排队超时：请求【从未送达】PowerShell，属于可安全重试的基础设施问题。 */
export class QueueBusyError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'QueueBusyError';
  }
}

interface PendingCall {
  resolve: (value: Record<string, unknown>) => void;
  reject: (err: Error) => void;
  timer: NodeJS.Timeout;
}

/** 队列里等位的一个请求 */
interface QueuedTask {
  action: string;
  params: Record<string, unknown>;
  exclusive: boolean;
  enqueuedAt: number;
  resolve: (value: Record<string, unknown>) => void;
  reject: (err: Error) => void;
  /** 排队超时定时器 */
  timer?: NodeJS.Timeout;
}

export interface ChildStats {
  app: string;
  ready: boolean;
  pid: number | null;
  uptimeMs: number;
  restarts: number;
  /** 在途（已送出、等 WPS 回话） */
  pending: number;
  /** 队列里等位的数量 */
  queued: number;
  /** 是否已有一个动作在途（STA：恒为 0 或 1） */
  inFlight: boolean;
  /** 队列上限 */
  queueMax: number;
  /** 排队时长上限 */
  queueTimeoutMs: number;
  /** 最近一次排队等待的毫秒数（无则 0） */
  lastQueueWaitMs: number;
  /** 累计因队列满/排队超时被拒的次数 */
  rejected: number;
}

export class PsChildDaemon {
  private proc: ChildProcess | null = null;
  private stdoutBuf = '';
  private pending = new Map<string, PendingCall>();
  private queue: QueuedTask[] = [];
  private starting: Promise<void> | null = null;
  private ready = false;
  private disposed = false;
  private startedAt: number | null = null;
  private restarts = 0;
  private lastQueueWaitMs = 0;
  private rejected = 0;
  /**
   * 是否已有一个动作被 drain 送出、还没回话。
   *
   * 不能用 pending.size 判断：pending 是【送出之后】才登记的，
   * 而 ensureStarted() 是异步的 —— 冷启动那几秒里 pending 一直是 0，
   * 结果一串并发请求会全部"穿过"队列被同时写进子进程（既破了并发上限，
   * 也让 STA 约束失效）。所以这里用一个同步置位的标志位。
   */
  private inFlight = false;

  constructor(private readonly appKey: string) {}

  get isReady(): boolean {
    return this.proc !== null && this.ready;
  }

  get stats(): ChildStats {
    return {
      app: this.appKey,
      ready: this.ready,
      pid: this.proc?.pid ?? null,
      uptimeMs: this.startedAt === null ? 0 : Date.now() - this.startedAt,
      restarts: this.restarts,
      pending: this.pending.size,
      queued: this.queue.length,
      inFlight: this.inFlight,
      queueMax: queueMax(),
      queueTimeoutMs: queueTimeoutMs(),
      lastQueueWaitMs: this.lastQueueWaitMs,
      rejected: this.rejected,
    };
  }

  /**
   * 调用一个动作。业务成功/失败都以响应对象 resolve；
   * 只有【基础设施问题】（起不来、超时、进程退出、队列满）才 reject。
   *
   * 并发语义（重要）
   * ----------------
   * COM 是 STA：一个应用同一时刻【只能有一个动作在跑】。所以这里不是"提高并发度"，
   * 而是把"无声排队"变成"有界的、看得见的等待"：
   *   - 队列上限 queueMax（默认 4），超了立刻回 QueueBusyError，不让调用方干等；
   *   - 排队超时 queueTimeoutMs（默认 10s），与动作超时独立；
   *   - 排他的动作（open/save/quit 等）插队，避免用户主动操作被一堆读操作压住。
   */
  async call(action: string, params: Record<string, unknown> = {}): Promise<Record<string, unknown>> {
    if (this.disposed) {
      throw new ChildUnavailableError(`[${this.appKey}] 子进程已关闭`);
    }

    const max = queueMax();
    if (this.queue.length >= max) {
      this.rejected += 1;
      throw new QueueBusyError(
        `WPS ${this.appKey} 队列已满（${max} 个在等，动作=${action}）。` +
          '该应用同一时刻只能跑一个动作（COM 是 STA），请稍后重试或先降低并发。'
      );
    }

    const exclusive = exclusiveSet().has(action);

    return new Promise<Record<string, unknown>>((resolve, reject) => {
      const task: QueuedTask = {
        action,
        params,
        exclusive,
        enqueuedAt: Date.now(),
        resolve,
        reject,
      };
      const qTimeout = queueTimeoutMs();
      task.timer = setTimeout(() => {
        const i = this.queue.indexOf(task);
        if (i >= 0) this.queue.splice(i, 1);
        this.rejected += 1;
        reject(
          new QueueBusyError(
            `WPS ${this.appKey} 排队超时（${qTimeout}ms，动作=${action}）。` +
              '前面还有动作在跑；这个请求从未送出，可以安全重试。'
          )
        );
      }, qTimeout);

      if (exclusive) this.queue.unshift(task);
      else this.queue.push(task);

      void this.drain();
    });
  }

  /**
   * 队列泵：把队首送进子进程，回话后继续送下一个。
   * 在途只允许 1 个 —— 这是 STA 的硬约束，不是策略选择。
   */
  private async drain(): Promise<void> {
    if (this.disposed || this.inFlight) return;
    const task = this.queue.shift();
    if (!task) return;
    if (task.timer) {
      clearTimeout(task.timer);
      task.timer = undefined;
    }
    this.inFlight = true;
    this.lastQueueWaitMs = Date.now() - task.enqueuedAt;

    try {
      const msg = await this.dispatch(task.action, task.params);
      task.resolve(msg);
    } catch (err) {
      task.reject(err as Error);
    } finally {
      this.inFlight = false;
      // 无论成败都继续放下一个，否则一次失败会把整个队列堵死
      void this.drain();
    }
  }

  /** 真正把一行 JSON 写进子进程并等它的响应 */
  private dispatch(
    action: string,
    params: Record<string, unknown>
  ): Promise<Record<string, unknown>> {
    return new Promise<Record<string, unknown>>((resolve, reject) => {
      // 冷启动放在真正送出时才做：排队中的请求不必陪着等握手
      this.ensureStarted()
        .then(() => {
          const id = uuidv4();
          const line = JSON.stringify({ id, action, params });
          const timeoutMs = callTimeoutMs();

          const timer = setTimeout(() => {
            this.pending.delete(id);
            // 卡死就是这一路死了：只丢弃重建【本应用】的进程，其它应用不受影响
            this.kill(`动作超时（${action}）`);
            reject(
              new Error(
                `WPS 动作超时（${timeoutMs}ms）: ${action}（应用=${this.appKey}）。` +
                  '该应用的常驻进程已重建；因原请求可能执行了一半，这里不自动重放。'
              )
            );
          }, timeoutMs);

          this.pending.set(id, { resolve, reject, timer });
          try {
            if (!this.proc?.stdin || this.proc.stdin.destroyed) {
              throw new Error('stdin 不可写');
            }
            this.proc.stdin.write(line + '\n');
          } catch (err) {
            clearTimeout(timer);
            this.pending.delete(id);
            reject(new Error(`写入常驻进程失败: ${(err as Error).message}`));
          }
        })
        .catch((err: unknown) => reject(err as Error));
    });
  }

  /** 杀掉当前子进程并清理在途请求（下次调用重建）。reason 仅用于诊断。 */
  kill(reason = '手动'): void {
    const p = this.proc;
    this.proc = null;
    this.ready = false;
    if (p) {
      try {
        p.stdin?.end();
      } catch {
        /* stdin 可能已关，忽略 */
      }
      try {
        p.kill();
      } catch {
        /* 进程可能已没，忽略 */
      }
    }
    this.failAll(new Error(`[${this.appKey}] WPS 常驻进程已退出（${reason}），请重试`));
  }

  dispose(): void {
    this.disposed = true;
    this.failQueued(new QueueBusyError(`[${this.appKey}] 桥正在关闭，队列中的请求已取消`));
    this.kill('dispose');
  }

  private failAll(err: Error): void {
    for (const [, p] of this.pending) {
      clearTimeout(p.timer);
      p.reject(err);
    }
    this.pending.clear();
  }

  /** 子进程没了，队列里等位的也没意义了 —— 立刻拒掉，不要让它们干等到排队超时 */
  private failQueued(err: Error): void {
    const tasks = this.queue.splice(0, this.queue.length);
    for (const t of tasks) {
      if (t.timer) clearTimeout(t.timer);
      this.rejected += 1;
      t.reject(err);
    }
  }

  private ensureStarted(): Promise<void> {
    if (this.proc && this.ready) return Promise.resolve();
    if (this.starting) return this.starting;
    this.starting = this.start().finally(() => {
      this.starting = null;
    });
    return this.starting;
  }

  private start(): Promise<void> {
    return new Promise<void>((resolve, reject) => {
      let proc: ChildProcess;
      try {
        proc = spawn(
          'powershell',
          ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', PS_SCRIPT],
          { windowsHide: true, stdio: ['pipe', 'pipe', 'pipe'] }
        );
      } catch (err) {
        reject(new ChildUnavailableError(`无法启动 PowerShell（${this.appKey}）: ${(err as Error).message}`));
        return;
      }

      this.proc = proc;
      this.ready = false;
      this.stdoutBuf = '';
      if (this.startedAt !== null) this.restarts += 1;
      this.startedAt = Date.now();

      let settled = false;
      const readyTimer = setTimeout(() => {
        if (settled) return;
        settled = true;
        this.kill('握手超时');
        reject(new ChildUnavailableError(`[${this.appKey}] 常驻进程 ${READY_TIMEOUT_MS}ms 内未完成握手`));
      }, READY_TIMEOUT_MS);

      const onReady = (): void => {
        if (settled) return;
        settled = true;
        clearTimeout(readyTimer);
        resolve();
      };

      proc.stdout?.on('data', (chunk: Buffer) => this.onStdout(chunk, onReady));

      proc.stderr?.on('data', (chunk: Buffer) => {
        const s = chunk.toString('utf8').trim();
        if (s) {
          // 桥进程 stdio 被忽略，只能靠文件日志
          this.logStderr(s);
        }
      });

      proc.on('error', (err: Error) => {
        clearTimeout(readyTimer);
        this.proc = null;
        this.ready = false;
        this.failAll(new Error(`[${this.appKey}] 常驻进程出错: ${err.message}`));
        if (!settled) {
          settled = true;
          reject(new ChildUnavailableError(`无法启动 PowerShell 常驻进程（${this.appKey}）: ${err.message}`));
        }
      });

      proc.on('exit', (code, signal) => {
        clearTimeout(readyTimer);
        if (this.proc === proc) {
          this.proc = null;
          this.ready = false;
        }
        this.failAll(new Error(`[${this.appKey}] 常驻进程已退出（code=${code} signal=${signal}），请重试`));
        if (!settled) {
          settled = true;
          reject(new ChildUnavailableError(`[${this.appKey}] 常驻进程启动后立即退出（code=${code}）`));
        }
      });
    });
  }

  /** 按行切分 stdout，各自当一条响应处理 */
  private onStdout(chunk: Buffer, onReady: () => void): void {
    this.stdoutBuf += chunk.toString('utf8');

    let idx = this.stdoutBuf.indexOf('\n');
    while (idx >= 0) {
      const line = this.stdoutBuf.slice(0, idx).trim();
      this.stdoutBuf = this.stdoutBuf.slice(idx + 1);
      if (line) this.handleLine(line, onReady);
      idx = this.stdoutBuf.indexOf('\n');
    }
  }

  private handleLine(line: string, onReady: () => void): void {
    let msg: Record<string, unknown>;
    try {
      msg = JSON.parse(line) as Record<string, unknown>;
    } catch {
      this.logStderr(`常驻进程输出非 JSON，已忽略: ${line.substring(0, 200)}`);
      return;
    }

    const data = msg.data as Record<string, unknown> | undefined;
    if (!this.ready && data && data.ready === true) {
      this.ready = true;
      onReady();
      return;
    }

    const id = typeof msg.id === 'string' ? msg.id : undefined;
    const pending = id ? this.pending.get(id) : undefined;
    if (!pending || !id) {
      this.logStderr(`收到无对应请求的响应（${this.appKey}）: ${line.substring(0, 200)}`);
      return;
    }
    clearTimeout(pending.timer);
    this.pending.delete(id);
    pending.resolve(msg);
  }

  /** 桥是 detached 进程，stdio 被忽略 —— 落文件日志，出问题有据可查 */
  private logStderr(text: string): void {
    try {
      // 延迟 require，避免在不需要日志时也加载 winston
      // eslint-disable-next-line @typescript-eslint/no-var-requires
      const { log } = require('../utils/logger') as { log: { warn: (m: string, ctx?: Record<string, unknown>) => void } };
      log.warn(`[bridge:${this.appKey}] ${text}`);
    } catch {
      /* 日志失败不能影响主流程 */
    }
  }
}
