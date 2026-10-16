/**
 * Input: MCP Tool 类型与Schema
 * Output: Tool 类型定义
 * Pos: MCP Tool 类型定义模块。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Tool类型定义 - 老王出品，必属精品
 * 定义MCP Tool的所有相关类型，别tm乱改
 */

/**
 * Tool参数的JSON Schema类型定义
 */
export interface ToolParameterSchema {
  type: 'string' | 'number' | 'boolean' | 'object' | 'array';
  description?: string;
  /**
   * 枚举可选值。声明成只读数组，是为了允许直接复用 `as const` 的 schema 常量
   * （否则 TS 会因为 readonly 数组不能赋给可变数组而报错 —— 抽常量复用时必踩）。
   */
  enum?: readonly string[];
  items?: ToolParameterSchema;
  properties?: Record<string, ToolParameterSchema>;
  required?: string[];
  default?: unknown;
}

/**
 * Tool输入Schema定义
 */
export interface ToolInputSchema {
  type: 'object';
  properties: Record<string, ToolParameterSchema>;
  required?: string[];
}

/**
 * Tool定义 - MCP协议规范的Tool结构
 */
export interface ToolDefinition {
  /** Tool名称，唯一标识符 */
  name: string;
  /** Tool描述，告诉AI这玩意儿能干啥 */
  description: string;
  /** 输入参数Schema */
  inputSchema: ToolInputSchema;
  /** Tool分类，方便管理 */
  category?: ToolCategory;
  /**
   * 工具组名。不填则按 category 推导（document→word / spreadsheet→excel /
   * presentation→ppt / common→common）。
   * 组的意义：未加载的组【不会出现在 tools/list 里】，以节省上下文。
   * 仅 `core` 组永远加载且不允许卸载（工具组管理工具自己在里面）。
   */
  group?: ToolGroupName;
}

/** 工具组名 —— 与三件套对应，另加 core（常驻管理类） */
export type ToolGroupName = 'core' | 'common' | 'word' | 'excel' | 'ppt';

/** 默认加载哪些组：ppt（115 个）默认不加载，用到再拉 */
export const DEFAULT_LOADED_GROUPS: ToolGroupName[] = ['core', 'common', 'word', 'excel'];

/** core 组永远加载，不允许卸载 */
export const ALWAYS_LOADED_GROUP: ToolGroupName = 'core';

/** 由 ToolDefinition 推导所属组（显式 group 优先） */
export function groupOfTool(def: ToolDefinition): ToolGroupName {
  if (def.group) return def.group;
  switch (def.category) {
    case ToolCategory.DOCUMENT:
      return 'word';
    case ToolCategory.SPREADSHEET:
      return 'excel';
    case ToolCategory.PRESENTATION:
      return 'ppt';
    default:
      return 'common';
  }
}

/**
 * Tool分类枚举 - WPS三件套对应的分类
 */
export enum ToolCategory {
  /** 文档操作 - Word那一套 */
  DOCUMENT = 'document',
  /** 表格操作 - Excel那一套 */
  SPREADSHEET = 'spreadsheet',
  /** 演示操作 - PPT那一套 */
  PRESENTATION = 'presentation',
  /** 通用操作 - 跨应用的通用功能 */
  COMMON = 'common',
}

/**
 * Tool调用请求
 */
export interface ToolCallRequest {
  /** 请求ID，用于追踪 */
  id: string;
  /** Tool名称 */
  name: string;
  /** 调用参数 */
  arguments: Record<string, unknown>;
}

/**
 * Tool调用结果
 */
export interface ToolCallResult {
  /** 对应的请求ID */
  id: string;
  /** 是否成功 */
  success: boolean;
  /** 返回的内容 */
  content: ToolContent[];
  /** 错误信息（失败时） */
  error?: string;
  /** 稳定的错误类型标识（失败时），取值见 WpsErrorType */
  errorType?: string;
}

/**
 * Tool返回内容类型
 */
export interface ToolContent {
  type: 'text' | 'image' | 'resource';
  text?: string;
  data?: string;
  mimeType?: string;
}

/**
 * Tool处理函数类型
 */
export type ToolHandler = (
  args: Record<string, unknown>
) => Promise<ToolCallResult>;

/**
 * 注册的Tool完整信息
 */
export interface RegisteredTool {
  definition: ToolDefinition;
  handler: ToolHandler;
}

/**
 * Tool列表响应
 */
export interface ListToolsResponse {
  tools: ToolDefinition[];
}

/**
 * Tool调用响应
 */
export interface CallToolResponse {
  content: ToolContent[];
  isError?: boolean;
}
