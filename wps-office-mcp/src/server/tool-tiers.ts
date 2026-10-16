/**
 * Input: 工具名
 * Output: 是否属于"基础档"
 * Pos: 工具档位（基础 40 / 全量）白名单。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么要这一个档
 * ----------------
 * 全量 270 个工具一次性塞进上下文既贵又稀释注意力；但只按组切（word/excel/ppt）粒度又太粗
 * —— 只做一份报告也要背上 45 个 word 工具。所以再叠一层"基础档"：默认只列 40 个
 * （覆盖三件套最常用的读写 + 工具组管理 + 诊断），其余靠 search_tools 搜 + load_tool_group 拉。
 *
 * 语义（要说清，否则容易误判）
 * ----------------------------
 *   默认（WPS_TOOL_TIER=essential）：列出的 = 基础档 40 个 ∪ 已"完整加载"的组
 *   WPS_TOOL_TIER=full：忽略基础档，按组来（列出 = 已加载组的全部工具）
 *   加载某个组 → 该组全部工具开始列出；卸载 → 回到基础档（基础档里属于该组的仍会列出，
 *   因为基础档是跨组的常用集，本来就该一直在）
 *
 * 清单怎么来：按"改现有文档"这条主干挑的 —— 先能定位/打开目标，再能读写，最后能导出。
 */

import { ToolGroupName } from '../types/tools';

/**
 * 基础档：core 之外的 35 个常用工具。
 * core 组自己有 5 个（list/load/unload/search + route_intent）且永远列出，合计正好 40。
 * 拆开算的原因：core 不需要白名单（见 ALWAYS_LOADED_GROUP），所以这份名单只放非 core 的。
 */
export const ESSENTIAL_TOOLS: readonly string[] = [
  // 通用（7）：状态、存盘、导出、逃生舱
  // 注意 wps_common_ping 不在名单里 —— liveness 由 wps_common_status 覆盖；
  // 它仍可直接调用，只是默认不列出。这样能空出一个名额给 wps_execute_method。
  'wps_common_status',
  'wps_common_save',
  'wps_common_save_as',
  'wps_common_show_window',
  'wps_convert_to_pdf',
  'wps_convert_format',
  'wps_execute_method',

  // Word（10）：定位 → 读 → 写 → 标题/大纲 → 字体 → 表格
  'wps_word_get_open_documents',
  'wps_word_open_document',
  'wps_word_switch_document',
  'wps_word_get_document_text',
  'wps_word_insert_text',
  'wps_word_find_replace',
  'wps_word_add_heading',
  'wps_word_get_outline',
  'wps_word_set_default_font',
  'wps_word_insert_table',

  // Excel（10）：定位 → 表 → 单元格读写 → 区域 → 公式 → 图表
  'wps_excel_get_open_workbooks',
  'wps_excel_open_workbook',
  'wps_excel_switch_workbook',
  'wps_excel_get_sheet_list',
  'wps_excel_get_cell_value',
  'wps_excel_set_cell_value',
  'wps_excel_read_range',
  'wps_excel_write_range',
  'wps_excel_set_formula',
  'wps_excel_create_chart',

  // PPT（8）：定位 → 新建 → 加页 → 标题 → 图 → 备注 → 导出图片
  'wps_ppt_get_open_presentations',
  'wps_ppt_create_presentation',
  'wps_ppt_add_slide',
  'wps_ppt_get_slide_count',
  'wps_ppt_set_slide_title',
  'wps_ppt_insert_image',
  'wps_ppt_export_slide_as_image',
  'wps_ppt_add_speaker_notes',
];

export const ESSENTIAL_TOOL_NAMES: ReadonlySet<string> = new Set(ESSENTIAL_TOOLS);

/** 基础档里各占多少（core 另算 4 个） */
export const ESSENTIAL_BY_GROUP: Readonly<Record<string, number>> = ESSENTIAL_TOOLS.reduce(
  (acc, name) => {
    const g = name.startsWith('wps_word_')
      ? 'word'
      : name.startsWith('wps_excel_')
        ? 'excel'
        : name.startsWith('wps_ppt_')
          ? 'ppt'
          : 'common';
    acc[g] = (acc[g] ?? 0) + 1;
    return acc;
  },
  {} as Record<string, number>
);

/** core 组的工具数（在 toolgroups.ts 里定义，这里只标注预期值，用于自检） */
export const CORE_GROUP_TOOL_COUNT = 4;

/** 当前档位：默认 essential（40），WPS_TOOL_TIER=full 关闭基础档过滤 */
export function isEssentialTier(): boolean {
  return (process.env.WPS_TOOL_TIER || 'essential').toLowerCase() !== 'full';
}

/** 诊断：把基础档清单与实际注册名对一遍，找出"写了但不存在"的名字（防手误） */
export function findUnknownEssentialTools(registeredNames: ReadonlySet<string>): string[] {
  return ESSENTIAL_TOOLS.filter((n) => !registeredNames.has(n));
}

/** 诊断用：某个工具组在基础档里占几个 */
export function essentialCountOf(group: ToolGroupName): number {
  return ESSENTIAL_BY_GROUP[group] ?? 0;
}
