/**
 * Input: 平台信息与WPS调用参数
 * Output: WPS API 调用结果
 * Pos: 跨平台 WPS 客户端。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * WPS通信客户端 - 仅 Windows
 * 通过 PowerShell 调用 WPS 的 COM 接口；优先走常驻桥，桥起不来退回直跑。
 *
 * 历史：曾经有一条 macOS/Linux 的"反向轮询"路线（WPS Mac 加载项在沙箱里
 * 起不了 HTTP server，只能让 MCP 当服务端、加载项来轮询）。该路线**已移除** ——
 * 只有代码没有配套加载项，实际跑不起来，留着只会让人误以为支持。
 * 相关代码（mac-poll-server.ts）与加载项（wps-claude-assistant/）已删除。
 */

import { spawn } from 'child_process';
import * as path from 'path';
import {
  WpsEndpointConfig,
  WpsApiRequest,
  WpsApiResponse,
  WpsAppType,
  WpsClientStatus,
  DocumentInfo,
  WorkbookInfo,
  PresentationInfo,
} from '../types/wps';
import { log, logRequest, logResponse } from '../utils/logger';
import { errorUtils } from '../utils/error';
import { psDaemon, DaemonUnavailableError } from './ps-daemon';
import { AppKey, resolveAppKey } from './action-routes';

// 仅 Windows：优先走常驻桥（省掉每次 spawn powershell + 解析 5000 行脚本）。
// 设 WPS_USE_DAEMON=0 可强制退回直跑模式。
const USE_DAEMON = process.env.WPS_USE_DAEMON !== '0';
// 守护进程一旦起不来就本次全程直跑，不再每次调用都白等 20 秒握手
let daemonBroken = false;

// PowerShell脚本路径 (Windows)
const PS_SCRIPT_PATH = path.join(__dirname, '../../scripts/wps-com.ps1');


// PowerShell 默认超时（毫秒）
const PS_TIMEOUT = 30000;

// ==================== 目标文稿锁定（避免多文档打开时 Active* 漂移）====================
// 通过各 app 的 set*Target 设置后，所有该 app 的调用自动注入目标名，
// 由 wps-com.ps1 的 Resolve-Doc / Resolve-Book / Get-TargetPres 精确定位；
// 单次调用显式传 documentName / workbookName / presentationName 可覆盖。
//
// 为什么三个都要有：ActiveDocument 是【弱引用】，实测会在 Documents.Count > 0 时
// 返回一个取不到内容的空壳（见 wps_common_repair 的诊断）。锁定目标名是唯一
// 不依赖 UI 焦点、也不受空壳影响的定位方式。
//
// 用法：
//   wps_word_switch_document / wps_excel_switch_workbook 已经在做这件事（它们会调 set*Target）
//   这里只是让"锁定"这件事对所有工具生效，而不只是切换类工具。
let pptTargetName: string | undefined;
let wordTargetName: string | undefined;
let excelTargetName: string | undefined;

export function setPptTarget(name?: string): void {
  pptTargetName = name && name.trim() ? name.trim() : undefined;
}
export function getPptTarget(): string | undefined {
  return pptTargetName;
}

export function setWordTarget(name?: string): void {
  wordTargetName = name && name.trim() ? name.trim() : undefined;
}
export function getWordTarget(): string | undefined {
  return wordTargetName;
}

export function setExcelTarget(name?: string): void {
  excelTargetName = name && name.trim() ? name.trim() : undefined;
}
export function getExcelTarget(): string | undefined {
  return excelTargetName;
}

/**
 * 执行PowerShell命令 (Windows)
 */
async function execPowerShell(action: string, params: Record<string, unknown> = {}): Promise<unknown> {
  return new Promise((resolve, reject) => {
    // JSON参数通过spawn args数组传递，Node自动处理Windows引号转义
    const paramsJson = JSON.stringify(params);
    const args = [
      // -NoProfile/-NonInteractive：直跑是回退路径，不该再花时间读用户 profile
      '-NoProfile',
      '-NonInteractive',
      '-ExecutionPolicy', 'Bypass',
      '-File', PS_SCRIPT_PATH,
      '-Action', action,
      '-Params', paramsJson
    ];

    log.debug('Executing PowerShell', { action, params });

    const ps = spawn('powershell', args, {
      windowsHide: true,
      stdio: ['pipe', 'pipe', 'pipe']
    });

    let stdout = '';
    let stderr = '';
    let killed = false;

    // 超时保护：防止PowerShell进程挂起，超时时主动 SIGTERM 并记录 PID
    const timeoutHandle = setTimeout(() => {
      killed = true;
      log.warn('PowerShell timeout, killing process', { pid: ps.pid, action, timeoutMs: PS_TIMEOUT });
      ps.kill('SIGTERM');
      reject(new Error(`PowerShell 执行超时（${PS_TIMEOUT}ms）: ${action} (PID: ${ps.pid})`));
    }, PS_TIMEOUT);

    ps.stdout.on('data', (data) => {
      stdout += data.toString();
    });

    ps.stderr.on('data', (data) => {
      stderr += data.toString();
    });

    ps.on('close', (code) => {
      clearTimeout(timeoutHandle);
      if (killed) return; // 已超时处理，忽略后续事件

      if (code !== 0) {
        // 非零退出码：分别处理 stderr 有内容 / 为空两种路径，避免空 stderr 误入 JSON.parse
        if (stderr) {
          log.error('PowerShell error', { stderr, code, pid: ps.pid, action });
          reject(new Error(stderr));
        } else {
          log.error('PowerShell exited with non-zero code', { code, stdout: stdout.substring(0, 200), pid: ps.pid, action });
          reject(new Error(`PowerShell 退出码 ${code}: ${stdout.substring(0, 200) || '(无输出)'}`));
        }
        return;
      }

      try {
        const result = JSON.parse(stdout.trim());
        resolve(result);
      } catch (_e) {
        log.error('Failed to parse PowerShell output', { stdout: stdout.substring(0, 200), pid: ps.pid, action });
        reject(new Error(`PowerShell 输出解析失败（非有效JSON）: ${stdout.substring(0, 200)}`));
      }
    });

    ps.on('error', (err) => {
      clearTimeout(timeoutHandle);
      if (killed) return;
      log.error('PowerShell spawn error', { error: err.message, pid: ps.pid, action });
      reject(new Error(`无法启动 PowerShell 进程: ${err.message}`));
    });
  });
}

/**
 * 统一执行接口（仅 Windows）
 *   常驻桥优先；桥起不来则退回"一调用一进程"的直跑方式
 */
async function execWpsAction(
  action: string,
  params: Record<string, unknown> = {},
  app: AppKey = 'common'
): Promise<unknown> {
  if (USE_DAEMON && !daemonBroken) {
    try {
      return await psDaemon.call(action, params, app);
    } catch (err) {
      // 只有"请求压根没送出去"（桥起不来）才允许退回直跑。
      // 超时那类错误请求可能已经在 WPS 里执行了一半，自动重放会变成执行两次 —— 一律如实抛出。
      if (err instanceof DaemonUnavailableError) {
        daemonBroken = true;
        log.warn('桥不可用，本进程后续改用直跑模式', { reason: err.message });
      } else {
        throw err;
      }
    }
  }

  return execPowerShell(action, params);
}

/**
 * WPS客户端类
 * 仅 Windows：PowerShell + COM 桥接
 */
export class WpsClient {
  private status: WpsClientStatus;

  constructor(_config?: Partial<WpsEndpointConfig>) {
    this.status = { connected: false };
    log.info('WPS Client initialized', { method: 'COM bridge (Windows only)' });
  }

  /**
   * 调用WPS接口（跨平台）
   * @param app 显式指定走哪一路子进程；不传则按动作名自动解析（见 action-routes）
   */
  async invokeAction<T = unknown>(
    action: string,
    params: Record<string, unknown> = {},
    app?: AppKey
  ): Promise<WpsApiResponse<T>> {
    const startTime = Date.now();
    logRequest(action, params);

    try {
      const result = (await execWpsAction(action, params, app ?? resolveAppKey(action))) as WpsApiResponse<T>;
      const duration = Date.now() - startTime;
      logResponse(action, result.success, duration);

      if (result.success) {
        this.status.connected = true;
        this.status.lastHeartbeat = new Date();
      }

      return result;
    } catch (error) {
      const duration = Date.now() - startTime;
      logResponse(action, false, duration);
      this.status.connected = false;
      throw errorUtils.wrap(error, `WPS 调用失败（COM 桥）: ${action}`);
    }
  }

  /**
   * 兼容旧API
   */
  async callApi<T = unknown>(request: WpsApiRequest): Promise<WpsApiResponse<T>> {
    const actionMap: Record<string, string> = {
      // Excel 旧API
      'workbook.getActive': 'getActiveWorkbook',
      'cell.getValue': 'getCellValue',
      'cell.setValue': 'setCellValue',
      'range.getData': 'getRangeData',
      'range.setData': 'setRangeData',
      // Word 旧API
      'document.getActive': 'getActiveDocument',
      'document.getText': 'getDocumentText',
      'document.insertText': 'insertText',
      // PPT 旧API
      'presentation.getActive': 'getActivePresentation',
      'presentation.addSlide': 'addSlide',
      // 通用旧API
      'file.save': 'save',
      'file.saveAs': 'saveAs',
      'file.open': 'openFile',
      'ping': 'ping',
    };
    const action = actionMap[request.method] || request.method;
    return this.invokeAction<T>(action, request.params || {});
  }

  /**
   * 检查WPS连接状态
   */
  async checkConnection(): Promise<boolean> {
    try {
      const result = await this.invokeAction('ping');
      this.status.connected = result.success;
      return result.success;
    } catch {
      this.status.connected = false;
      this.status.error = 'Connection check failed';
      return false;
    }
  }

  /**
   * 获取客户端状态
   */
  getStatus(): WpsClientStatus {
    return { ...this.status };
  }

  // ==================== 表格操作 (WPS表格) ====================

  async getActiveWorkbook(): Promise<WorkbookInfo | null> {
    const response = await this.invokeAction<WorkbookInfo>('getActiveWorkbook');
    return response.success ? response.data || null : null;
  }

  async getCellValue(sheet: string | number, row: number, col: number): Promise<unknown> {
    const response = await this.invokeAction<{ value: unknown }>('getCellValue', { sheet, row, col });
    return response.data?.value;
  }

  async setCellValue(sheet: string | number, row: number, col: number, value: unknown): Promise<boolean> {
    const response = await this.invokeAction('setCellValue', { sheet, row, col, value });
    return response.success;
  }

  async getRangeData(sheet: string | number, range: string): Promise<unknown[][]> {
    const response = await this.invokeAction<{ data: unknown[][] }>('getRangeData', { sheet, range });
    return response.data?.data || [];
  }

  async setRangeData(sheet: string | number, range: string, data: unknown[][]): Promise<boolean> {
    const response = await this.invokeAction('setRangeData', { sheet, range, data });
    return response.success;
  }

  async setFormula(sheet: string | number, row: number, col: number, formula: string): Promise<boolean> {
    const response = await this.invokeAction('setFormula', { sheet, row, col, formula });
    return response.success;
  }

  // ==================== 文档操作 (WPS文字) ====================

  async getActiveDocument(): Promise<DocumentInfo | null> {
    const response = await this.invokeAction<DocumentInfo>('getActiveDocument');
    return response.success ? response.data || null : null;
  }

  async createDocument(): Promise<boolean> {
    const response = await this.invokeAction('createDocument');
    return response.success;
  }

  async insertText(text: string, position?: number): Promise<boolean> {
    const response = await this.invokeAction('insertText', { text, position });
    return response.success;
  }

  async getDocumentText(): Promise<string> {
    const response = await this.invokeAction<{ text: string }>('getDocumentText');
    return response.data?.text || '';
  }

  // ==================== 演示操作 (WPS演示) ====================

  async getActivePresentation(): Promise<PresentationInfo | null> {
    const response = await this.invokeAction<PresentationInfo>('getActivePresentation');
    return response.success ? response.data || null : null;
  }

  async createPresentation(): Promise<boolean> {
    const response = await this.invokeAction('createPresentation');
    return response.success;
  }

  async addSlide(layout?: string): Promise<boolean> {
    const response = await this.invokeAction('addSlide', { layout });
    return response.success;
  }

  // ==================== 通用操作 ====================

  /**
   * 通用方法调用入口。
   * appType 是调用点声明的目标应用，会被翻译成桥的应用键（excel/word/ppt），
   * 从而落到对应的常驻子进程上 —— 这是 per-app 隔离的关键，别把它丢掉。
   */
  async executeMethod<T = unknown>(
    method: string,
    params?: Record<string, unknown>,
    appType?: WpsAppType
  ): Promise<WpsApiResponse<T>> {
    let finalParams: Record<string, unknown> = params || {};
    // 锁定目标文稿后自动注入目标名（已显式传入则不覆盖）。
    // 三个 app 各有自己的参数名：presentationName / documentName / workbookName ——
    // 名字与 wps-com.ps1 里 Resolve-Doc / Resolve-Book / Get-TargetPres 的读取一致。
    if (appType === WpsAppType.PRESENTATION && pptTargetName && finalParams.presentationName === undefined) {
      finalParams = { ...finalParams, presentationName: pptTargetName };
    }
    if (appType === WpsAppType.WRITER && wordTargetName && finalParams.documentName === undefined) {
      finalParams = { ...finalParams, documentName: wordTargetName };
    }
    if (appType === WpsAppType.SPREADSHEET && excelTargetName && finalParams.workbookName === undefined) {
      finalParams = { ...finalParams, workbookName: excelTargetName };
    }
    return this.invokeAction<T>(method, finalParams, resolveAppKey(method, appType));
  }

  async openFile(filePath: string, _appType?: WpsAppType): Promise<boolean> {
    const response = await this.invokeAction('openFile', { path: filePath });
    return response.success;
  }

  async saveFile(_appType?: WpsAppType): Promise<boolean> {
    const response = await this.invokeAction('save');
    return response.success;
  }

  async saveFileAs(filePath: string, _appType?: WpsAppType): Promise<boolean> {
    const response = await this.invokeAction('saveAs', { path: filePath });
    return response.success;
  }
}

// 导出单例
export const wpsClient = new WpsClient();

export default WpsClient;
