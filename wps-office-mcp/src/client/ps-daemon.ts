/**
 * Input: 动作名与参数对象
 * Output: WPS 动作结果（与直跑模式同形：{ success, data? , error? }）
 * Pos: PowerShell 常驻守护进程客户端。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 *
 * 为什么要有这个文件
 * ------------------
 * 改造前每调一次工具就 spawn 一个 powershell.exe，重新解析 5000 行 wps-com.ps1，再连一次 COM，
 * 单次白花 0.3~0.6 秒；多步任务里这是主要瓶颈。
 * 这里改成常驻一个 powershell 进程（scripts/wps-comd.ps1），按【一行一个 JSON】收发：
 * 进程只起一次、脚本只解析一次、COM 附着关系也留在同一个 STA 线程上。
 *
 * 两条路都在
 * ----------
 * wps-client.ts 优先走守护进程；守护进程【起不来】时自动退回原来的直跑方式，
 * 不让一个进程启动问题把整条链路堵死。设 WPS_USE_DAEMON=0 可强制走直跑。
 *
 * 超时语义（对标 office-mcp 的 bridge/connection_manager.py）
 * ---------------------------------------------------------
 * COM 是 STA，守护进程同一时刻只能跑一个动作。某个动作卡死会让整个进程失去响应，
 * 所以【超时即判定该连接已死】：杀掉进程、拒绝所有在途请求，下次调用重建。
 * 绝不能"超时了但继续往这个进程里塞请求"——那只会一个接一个地接着超时。
 *
 * 另一条红线：超时后【不自动重试】。写动作（保存/导出/退出）可能已经执行了一半，
 * 自动重放会变成"执行两次"。超时一律如实抛错，由调用方决定怎么办。
 */

import { ChildProcess, spawn } from 'child_process';
import * as path from 'path';
import { v4 as uuidv4 } from 'uuid';
import { log } from '../utils/logger';

/** 守护进程脚本（与本文件同时随包分发，见 .npmignore / package.json） */
const PS_SCRIPT = path.join(__dirname, '../../scripts/wps-comd.ps1');

/** 起进程 + 握手就绪的上限 */
const READY_TIMEOUT_MS = 20000;
/** 单次动作上限。COM 动作可能很慢（大文档导出、第一次插图表要拉起 Excel），默认 120 秒。 */
const CALL_TIMEOUT_MS = parseInt(process.env.WPS_DAEMON_TIMEOUT || '120000', 10);

/**
 * 守护进程【根本没跑起来】时抛这个 —— 调用方看到它可以安全地退回直跑，
 * 因为请求从未被送达。与之相对，超时错误的请求可能已经执行了一半，绝不能自动重放。
 */
export class DaemonUnavailableError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'DaemonUnavailableError';
  }
}

interface PendingCall {
  resolve: (value: unknown) => void;
  reject: (err: Error) => void;
  timer: NodeJS.Timeout;
}

export class PsDaemon {
  private proc: ChildProcess | null = null;
  private stdoutBuf = '';
  private pending = new Map<string, PendingCall>();
  private starting: Promise<void> | null = null;
  private ready = false;
  private disposed = false;
  /** 当前这个进程是什么时候起来的（诊断用） */
  private startedAt: number | null = null;
  /** 重建过几次（超时/异常导致的），诊断用 */
  private restarts = 0;

  /** 当前是否有一个已就绪的守护进程 */
  get isReady(): boolean {
    return this.proc !== null && this.ready;
  }

  /** 在途请求数（诊断用） */
  get pendingCount(): number {
    return this.pending.size;
  }

  /** 桥这一层的状态，给 wps_status 用 */
  get stats(): {
    ready: boolean;
    pid: number | null;
    uptimeMs: number;
    restarts: number;
    pending: number;
    disposed: boolean;
  } {
    return {
      ready: this.ready,
      pid: this.proc?.pid ?? null,
      uptimeMs: this.startedAt === null ? 0 : Date.now() - this.startedAt,
      restarts: this.restarts,
      pending: this.pending.size,
      disposed: this.disposed,
    };
  }

  /**
   * 调用一个动作。成功/业务失败都以响应对象 resolve；
   * 只有【基础设施问题】（起不来、超时、进程退出）才 reject。
   */
  async call(action: string, params: Record<string, unknown> = {}): Promise<unknown> {
    if (this.disposed) {
      throw new DaemonUnavailableError('守护进程已关闭（dispose 之后不再接受调用）');
    }
    await this.ensureStarted();

    const id = uuidv4();
    const line = JSON.stringify({ id, action, params });

    return new Promise<unknown>((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id);
        // 卡死就是整条连接死了：丢弃重建。不重试，见文件头说明。
        log.error('守护进程动作超时，丢弃并重建', undefined, {
          action,
          id,
          timeoutMs: CALL_TIMEOUT_MS,
        });
        this.kill();
        reject(
          new Error(
            `WPS 动作超时（${CALL_TIMEOUT_MS}ms）: ${action}。` +
              '守护进程已重建，可重试；但因为原请求可能执行了一半，这里不自动重放。'
          )
        );
      }, CALL_TIMEOUT_MS);

      this.pending.set(id, { resolve, reject, timer });
      try {
        if (!this.proc?.stdin || this.proc.stdin.destroyed) {
          throw new Error('stdin 不可写');
        }
        this.proc.stdin.write(line + '\n');
      } catch (err) {
        clearTimeout(timer);
        this.pending.delete(id);
        reject(new Error(`写入守护进程失败: ${(err as Error).message}`));
      }
    });
  }

  /** 关掉守护进程，用于进程退出时收尾 */
  dispose(): void {
    this.disposed = true;
    this.kill();
  }

  /** 杀掉当前进程并清理在途请求（下次调用会重建） */
  kill(): void {
    const p = this.proc;
    this.proc = null;
    this.ready = false;
    if (p) {
      try {
        p.stdin?.end();
      } catch {
        /* stdin 可能已经关了，忽略 */
      }
      try {
        p.kill();
      } catch {
        /* 进程可能已经没了，忽略 */
      }
    }
    this.failAll(new Error('WPS 守护进程已退出，请重试'));
  }

  private failAll(err: Error): void {
    for (const [, p] of this.pending) {
      clearTimeout(p.timer);
      p.reject(err);
    }
    this.pending.clear();
  }

  private ensureStarted(): Promise<void> {
    if (this.proc && this.ready) {
      return Promise.resolve();
    }
    if (this.starting) {
      return this.starting;
    }
    this.starting = this.start().finally(() => {
      this.starting = null;
    });
    return this.starting;
  }

  private start(): Promise<void> {
    return new Promise<void>((resolve, reject) => {
      log.info('启动 WPS COM 守护进程', { script: PS_SCRIPT });

      let proc: ChildProcess;
      try {
        proc = spawn(
          'powershell',
          ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', PS_SCRIPT],
          { windowsHide: true, stdio: ['pipe', 'pipe', 'pipe'] }
        );
      } catch (err) {
        reject(new DaemonUnavailableError(`无法启动 PowerShell: ${(err as Error).message}`));
        return;
      }

      this.proc = proc;
      this.ready = false;
      this.stdoutBuf = '';
      if (this.startedAt !== null) {
        this.restarts += 1;
      }
      this.startedAt = Date.now();

      let settled = false;
      const readyTimer = setTimeout(() => {
        if (settled) return;
        settled = true;
        log.error('守护进程握手超时', undefined, { timeoutMs: READY_TIMEOUT_MS });
        this.kill();
        reject(
          new DaemonUnavailableError(`WPS 守护进程 ${READY_TIMEOUT_MS}ms 内未完成握手`)
        );
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
          log.warn('守护进程 stderr', { stderr: s.substring(0, 400) });
        }
      });

      proc.on('error', (err: Error) => {
        clearTimeout(readyTimer);
        this.proc = null;
        this.ready = false;
        log.error('守护进程启动出错', err);
        this.failAll(new Error(`WPS 守护进程出错: ${err.message}`));
        if (!settled) {
          settled = true;
          reject(new DaemonUnavailableError(`无法启动 PowerShell 守护进程: ${err.message}`));
        }
      });

      proc.on('exit', (code, signal) => {
        clearTimeout(readyTimer);
        if (this.proc === proc) {
          this.proc = null;
          this.ready = false;
        }
        log.warn('守护进程退出', { code, signal });
        this.failAll(new Error(`WPS 守护进程已退出（code=${code} signal=${signal}），请重试`));
        if (!settled) {
          settled = true;
          reject(new DaemonUnavailableError(`WPS 守护进程启动后立即退出（code=${code}）`));
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

      if (line) {
        this.handleLine(line, onReady);
      }
      idx = this.stdoutBuf.indexOf('\n');
    }
  }

  private handleLine(line: string, onReady: () => void): void {
    let msg: Record<string, unknown>;
    try {
      msg = JSON.parse(line) as Record<string, unknown>;
    } catch {
      // 守护进程只应输出 JSON；出现别的说明有东西污染了 stdout，如实报出来别吞掉
      log.warn('守护进程输出非 JSON，已忽略', { line: line.substring(0, 200) });
      return;
    }

    const data = msg.data as Record<string, unknown> | undefined;
    if (!this.ready && data && data.ready === true) {
      this.ready = true;
      log.info('守护进程就绪', { pid: data.pid });
      onReady();
      return;
    }

    const id = typeof msg.id === 'string' ? msg.id : undefined;
    const pending = id ? this.pending.get(id) : undefined;
    if (!pending || !id) {
      log.warn('收到无对应请求的响应', { id, body: line.substring(0, 200) });
      return;
    }
    clearTimeout(pending.timer);
    this.pending.delete(id);
    pending.resolve(msg);
  }
}

/** 全局单例：一个 MCP server 进程只维护一个守护进程 */
export const psDaemon = new PsDaemon();

export default psDaemon;
