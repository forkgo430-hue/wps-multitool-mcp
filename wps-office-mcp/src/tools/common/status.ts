/**
 * Input: 无
 * Output: 桥/守护进程 + WPS 三件套的运行状态
 * Pos: 通用状态探针工具。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 *
 * 为什么需要它
 * ------------
 * 改造前没有任何地方能回答"现在到底连上没有"。出了问题只能看日志翻，
 * 或者盲猜。对标 office-mcp 的 office_status() 与 ghidra-mcp 的 mcp_health。
 *
 * 报告两层：
 *   1) 桥这一层（本进程内）：守护进程是否就绪、PID、存活时长、重建次数、在途请求数
 *   2) WPS 这一层（PowerShell 侧）：三件套分别有没有【已附着】的实例
 *
 * 要点：本工具【只读】且【不产生副作用】。
 *   PowerShell 侧走 wpsStatus 动作，内部只用 Get-ActiveWpsApp（只取已运行实例），
 *   不会像 getAppInfo 那样顺手 New-Object 把 WPS 拉起来。
 *   WPS 没开时也要能正常返回 —— 状态探针不该因为目标没运行就报错。
 */

import {
  ToolDefinition,
  ToolHandler,
  ToolCallResult,
  ToolCategory,
  RegisteredTool,
} from '../../types/tools';
import { wpsClient } from '../../client/wps-client';
import { psDaemon } from '../../client/ps-daemon';

/** PowerShell 侧 wpsStatus 动作返回的数据形状 */
interface AppState {
  running: boolean;
  name?: string;
  version?: string;
  documents?: number;
}

interface WpsStatusData {
  apps?: Record<string, AppState>;
  psPid?: number;
  psVersion?: string;
  daemon?: boolean;
}

export const statusDefinition: ToolDefinition = {
  name: 'wps_common_status',
  description: `查看 WPS MCP 的运行状态：桥（守护进程）这一层 + WPS 三件套这一层。

报告内容：
- 守护进程：是否就绪、进程号、已存活多久、重建过几次、当前在途请求数
- WPS 三件套：表格/文字/演示分别有没有已连接的实例，以及实例名、版本、打开文档数
- PowerShell 侧进程号与版本

使用场景：
- "WPS 连上了吗"
- "为什么刚才那个工具没反应"
- "现在开了几个文档"
- 排查问题时先跑这个，比翻日志快

本工具只读，且不会顺手把 WPS 启动起来。`,
  category: ToolCategory.COMMON,
  inputSchema: {
    type: 'object',
    properties: {},
  },
};

export const statusHandler: ToolHandler = async (): Promise<ToolCallResult> => {
  // 桥这一层：永远拿得到，本进程内的
  const bridge = {
    daemonEnabled: process.env.WPS_USE_DAEMON !== '0',
    ...psDaemon.stats,
  };

  // WPS 这一层：拿不到就如实标注，不要把失败伪装成"没运行"
  let wps: WpsStatusData | null = null;
  let wpsError: string | null = null;
  try {
    const response = await wpsClient.executeMethod<WpsStatusData>('wpsStatus', {});
    if (response.success) {
      wps = response.data ?? null;
    } else {
      wpsError = String(response.error ?? '未知错误');
    }
  } catch (err) {
    wpsError = err instanceof Error ? err.message : String(err);
  }

  const payload = {
    bridge,
    wps: wps
      ? {
          apps: wps.apps ?? {},
          powershell: { pid: wps.psPid, version: wps.psVersion, daemonMode: wps.daemon },
        }
      : null,
    wpsError,
  };

  return {
    id: '',
    success: true,
    content: [{ type: 'text', text: JSON.stringify(payload, null, 2) }],
  };
};

/** 状态类工具集合 */
export const statusTools: RegisteredTool[] = [
  { definition: statusDefinition, handler: statusHandler },
];

export default statusTools;
