/**
 * Input: MCP 工具名
 * Output: MCP 标准工具注解（readOnlyHint / destructiveHint / idempotentHint / openWorldHint）
 * Pos: 让客户端【从 tools/list 直接读到】每个工具的读写属性，而不是让模型靠描述去猜。
 *      一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么必须是注解，不能靠文档教
 * ------------------------------
 * 实测 273 个工具里，靠名字能一眼确认只读的只有 37 个（见 cancellable.ts 的判定逻辑），
 * 其余 236 个是写入或判不出。指望模型读描述去判断"这个能不能动我的文档"，
 * 就是让它在 236 个模糊项里赌。注解是机器可读的，客户端可以据此决定要不要弹确认、
 * 要不要允许自动执行 —— 这比在 skill 里写一句"注意只读"可靠得多。
 *
 * 数据来源
 * --------
 * 动作名映射由 scripts/manual/gen-tool-annotations.js 运行时观测生成：
 * 把 wpsClient 的出口换成假实现，逐个跑 handler 记下它发出的动作名。
 * 注解本身在这里【运行时算】，不写死进生成文件 —— 判定规则改了不必重新生成。
 */

import {
  isReadOnlyAction,
  isDestructiveAction,
  isIdempotentAction,
} from '../bridge/cancellable';
import { TOOL_ACTIONS } from './tool-annotations.generated';

/** MCP 规范里的工具注解（字段名与规范一致，别改） */
export interface ToolAnnotations {
  /** 人类可读的短标题，给不支持富渲染的客户端用 */
  title?: string;
  /** 只读：不改任何状态。客户端可放心自动执行 */
  readOnlyHint?: boolean;
  /** 可能造成破坏性更新：会删内容、丢未保存改动、原地改格式 */
  destructiveHint?: boolean;
  /** 重复执行结果一致（可安全重放） */
  idempotentHint?: boolean;
  /** 会与外部系统交互。本 MCP 只碰本机 WPS，一律 false */
  openWorldHint?: boolean;
}

const cache = new Map<string, ToolAnnotations | undefined>();

/**
 * 自动生成覆盖不到的工具 —— 手写在这里，并说明为什么。
 *
 * wps_execute_method 在 mcp-server.ts 里直接 this.registry.register()，
 * 不走 allTools，所以 gen 脚本观测不到它。它能执行【任意】WPS 动作
 * （参数是方法名字符串），意味着调用方可以让它删东西、关文档、原地转格式，
 * 所以破坏性按 true 给，不是 false。
 */
const EXTRA_ANNOTATIONS: Record<string, ToolAnnotations> = {
  wps_execute_method: {
    readOnlyHint: false,
    destructiveHint: true,
    idempotentHint: false,
    openWorldHint: false,
  },
};

/**
 * 算一个工具的注解。
 * 拿不到动作名（工具没在映射表里、也不在 EXTRA 里）就返回 undefined ——
 * 宁可没有注解，也不要给一个猜出来的错注解（客户端可能拿它做安全决策）。
 */
export function getToolAnnotations(toolName: string): ToolAnnotations | undefined {
  if (cache.has(toolName)) return cache.get(toolName);

  const extra = EXTRA_ANNOTATIONS[toolName];
  if (extra) {
    cache.set(toolName, extra);
    return extra;
  }

  const actions = TOOL_ACTIONS[toolName];
  if (!actions || !actions.length) {
    cache.set(toolName, undefined);
    return undefined;
  }

  // 一个工具可能发多条命令，只要有一条会写，它就不能算只读
  const readOnly = actions.every(isReadOnlyAction);
  const destructive = actions.some(isDestructiveAction);

  const ann: ToolAnnotations = {
    readOnlyHint: readOnly,
    // 只读的按定义不破坏；写入里只有删/丢数据/原地转换那批算破坏
    destructiveHint: readOnly ? false : destructive,
    idempotentHint: actions.every(isIdempotentAction),
    // 本 MCP 的全部副作用都在本机 WPS，不访问外部系统
    openWorldHint: false,
  };

  cache.set(toolName, ann);
  return ann;
}

/** 注解覆盖了多少工具（诊断用） */
export function annotationCoverage(toolNames: string[]): {
  total: number;
  annotated: number;
} {
  const annotated = toolNames.filter((n) => getToolAnnotations(n) !== undefined).length;
  return { total: toolNames.length, annotated };
}

/** 汇总分布（诊断/文档用，数字必须实测别手写） */
export function annotationSummary(toolNames: string[]): {
  total: number;
  annotated: number;
  readOnly: number;
  destructive: number;
  nonIdempotent: number;
} {
  let readOnly = 0;
  let destructive = 0;
  let nonIdempotent = 0;
  let annotated = 0;
  for (const n of toolNames) {
    const a = getToolAnnotations(n);
    if (!a) continue;
    annotated++;
    if (a.readOnlyHint) readOnly++;
    if (a.destructiveHint) destructive++;
    if (a.idempotentHint === false) nonIdempotent++;
  }
  return { total: toolNames.length, annotated, readOnly, destructive, nonIdempotent };
}