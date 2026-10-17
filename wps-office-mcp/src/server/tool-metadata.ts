/**
 * Input: 工具定义（名字 / 描述）
 * Output: 该工具的中文别名、正例、反例（给 wps_search_tools 用）
 * Pos: 工具检索元数据。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 *
 * 为什么要有这一层
 * ----------------
 * 先说清楚它【不是】在补什么：工具描述本身就是中文，所以中文关键词原本就能搜到
 * （实测 66 个常见中文关键词，去掉别名后仍有绝大多数命中）。别名不是救命稻草。
 *
 * 真正的问题在别处——【近名工具分不出来】。本仓有 6 个 background、5 个 animation、
 * 4 个 transition、3 个 image 插入、2 个 toc、2 个 zoom，名字和描述高度重合，
 * 模型搜出来一堆同名候选只能乱猜。所以这一层主要补两样：
 *   1) 反例 notFor：显式写"这个不干 X，要干 X 用 Y" —— 这是主力价值
 *   2) 别名：按英文词根表自动拼中文说法（285 个工具全覆盖，不用逐个手写），
 *      实测能给 66 个关键词里的 17 个多捞到命中（共 +34 对），是锦上添花
 *
 * 打分为什么别名最低
 * ------------------
 * 别名是机器按词根拼的，会串味（搜"域"会命中"区域"，搜"组合"会命中 load_tool_group）。
 * 所以 searchTools 里 name(100) > description(40) > alias(25)，别让派生词压过人写的描述。
 *
 * 设计取舍
 * --------
 * 不搞 embedding / 向量检索：近名工具的区别是"哪个管哪一页/哪种背景"这种事实差异，
 * 靠相似度算不出来，只能显式写；引入模型预分类属于过度设计。
 */

import { ToolDefinition } from '../types/tools';

/** 应用前缀，别名里不重复出现 */
const APP_TOKENS = new Set(['wps', 'word', 'excel', 'ppt', 'common', 'core']);

/**
 * 英文词根 → 中文同义词。
 * 只在"自动拼出来的说法像人话"时才填；拼不出来的交给下面的 CURATED 手工补。
 */
const TOKEN_ZH: Record<string, string[]> = {
  // 动词
  set: ['设置', '设定'],
  get: ['获取', '读取', '查看'],
  add: ['添加', '新增', '加'],
  insert: ['插入', '新增', '加入'],
  delete: ['删除', '删掉', '去掉'],
  remove: ['删除', '移除', '去掉'],
  create: ['创建', '新建', '生成'],
  generate: ['生成', '创建'],
  update: ['更新', '刷新'],
  modify: ['修改', '改'],
  apply: ['应用', '套用'],
  export: ['导出', '输出'],
  open: ['打开'],
  close: ['关闭'],
  switch: ['切换'],
  save: ['保存'],
  convert: ['转换', '转'],
  move: ['移动'],
  copy: ['复制', '拷贝'],
  paste: ['粘贴'],
  find: ['查找', '搜索'],
  replace: ['替换'],
  read: ['读取'],
  write: ['写入'],
  show: ['显示', '展示'],
  hide: ['隐藏'],
  group: ['组合', '分组'],
  ungroup: ['取消组合'],
  merge: ['合并'],
  unmerge: ['取消合并'],
  split: ['拆分', '分割'],
  fill: ['填充'],
  sort: ['排序'],
  filter: ['筛选', '过滤'],
  freeze: ['冻结'],
  protect: ['保护'],
  unprotect: ['取消保护'],
  lock: ['锁定'],
  zoom: ['缩放', '显示比例'],
  print: ['打印'],
  evaluate: ['求值', '计算'],
  diagnose: ['诊断', '排查'],
  transpose: ['转置'],
  subtotal: ['分类汇总'],
  distribute: ['分布', '等距分布'],
  align: ['对齐'],
  duplicate: ['复制', '副本'],
  beautify: ['美化'],
  unify: ['统一'],
  start: ['开始', '启动'],
  list: ['列出', '列举'],
  search: ['搜索', '查找'],
  enable: ['开启', '启用'],
  clear: ['清除', '清空'],
  clean: ['清洗', '清理'],
  rename: ['重命名', '改名'],
  auto: ['自动'],
  smart: ['智能'],
  route: ['路由', '识别'],
  load: ['加载'],
  unload: ['卸载'],
  repair: ['修复', '自救'],
  ping: ['连通性', '探活'],

  // 名词
  formula: ['公式'],
  range: ['区域', '范围'],
  cell: ['单元格'],
  cells: ['单元格'],
  row: ['行'],
  rows: ['行'],
  column: ['列'],
  columns: ['列', '分栏'],
  sheet: ['工作表', 'sheet'],
  sheets: ['工作表'],
  workbook: ['工作簿', '表格文件'],
  workbooks: ['工作簿'],
  chart: ['图表'],
  charts: ['图表'],
  image: ['图片'],
  images: ['图片'],
  comment: ['批注'],
  comments: ['批注'],
  style: ['样式'],
  styles: ['样式'],
  text: ['文本', '文字'],
  textbox: ['文本框'],
  textboxes: ['文本框'],
  font: ['字体'],
  color: ['颜色'],
  paragraph: ['段落'],
  paragraphs: ['段落'],
  table: ['表格'],
  tables: ['表格'],
  heading: ['标题'],
  toc: ['目录'],
  bookmark: ['书签'],
  page: ['页面', '页'],
  section: ['节', '分节'],
  header: ['页眉'],
  footer: ['页脚'],
  footnote: ['脚注'],
  caption: ['题注'],
  numbering: ['编号'],
  border: ['边框'],
  number: ['数字', '编号'],
  format: ['格式'],
  data: ['数据'],
  validation: ['数据验证', '有效性'],
  named: ['命名'],
  hyperlink: ['超链接'],
  pivot: ['数据透视表'],
  selection: ['选区', '选中区域'],
  window: ['窗口'],
  document: ['文档'],
  documents: ['文档'],
  slide: ['幻灯片', '页面'],
  slides: ['幻灯片'],
  presentation: ['演示文稿', 'PPT'],
  presentations: ['演示文稿'],
  shape: ['形状', '图形'],
  shapes: ['形状'],
  animation: ['动画'],
  animations: ['动画'],
  transition: ['切换效果', '切换动画'],
  background: ['背景'],
  master: ['母版'],
  theme: ['主题'],
  layout: ['版式', '布局'],
  notes: ['备注'],
  title: ['标题'],
  subtitle: ['副标题'],
  content: ['内容'],
  speaker: ['演讲者'],
  date: ['日期'],
  time: ['时间'],
  field: ['域', '字段'],
  fields: ['域'],
  outline: ['大纲'],
  revision: ['修订'],
  revisions: ['修订'],
  index: ['索引'],
  cross: ['交叉'],
  reference: ['引用'],
  track: ['跟踪'],
  changes: ['更改'],
  proofread: ['校对'],
  spacing: ['间距'],
  line: ['行距', '行'],
  setup: ['设置'],
  break: ['分隔符', '分页'],
  area: ['区域'],
  width: ['宽度', '列宽'],
  height: ['高度', '行高'],
  sum: ['求和'],
  array: ['数组'],
  conditional: ['条件'],
  info: ['信息'],
  status: ['状态'],
  value: ['值'],
  position: ['位置'],
  shadow: ['阴影'],
  gradient: ['渐变'],
  transparency: ['透明度'],
  depth: ['深度'],
  material: ['材质'],
  rotation: ['旋转'],
  order: ['顺序', '层级'],
  series: ['序列'],
  grid: ['网格'],
  gauge: ['仪表盘'],
  timeline: ['时间轴'],
  donut: ['环形图'],
  kpi: ['KPI', '指标卡'],
  cards: ['卡片'],
  progress: ['进度'],
  bar: ['条', '进度条'],
  indicator: ['指示器', '页码'],
  decoration: ['装饰'],
  method: ['方法'],
  pdf: ['PDF'],
  app: ['应用', '程序'],
  quit: ['退出'],
  wire: ['连通性'],
  tool: ['工具'],
  groups: ['工具组'],
  intent: ['意图'],
  emphasis: ['强调'],
  preset: ['预设'],
  active: ['当前', '活动'],
  target: ['目标'],
  default: ['默认'],
  basic: ['基础'],
  selected: ['选中'],
  level: ['层级'],
  flow: ['流程'],
  org: ['组织'],
  mini: ['迷你'],
};

/**
 * 手工补充：自动拼不出来的口语说法，以及【本仓真的会混】的那些近名工具的正例/反例。
 * notFor 写"这个工具不干 X，要干 X 用 Y" —— 这是模型最容易踩的坑。
 */
const CURATED: Record<
  string,
  { aliases?: string[]; examples?: string[]; notFor?: string[] }
> = {
  // ===== core =====
  wps_search_tools: {
    aliases: ['找工具', '搜工具', '有没有什么工具'],
    examples: ['有没有导出图片的工具', '找一下和批注有关的工具'],
    notFor: ['不是执行工具，只负责找；找到后要么直接调用，要么先 wps_load_tool_group'],
  },
  wps_load_tool_group: {
    aliases: ['加载组', '拉起组', '把PPT工具弄出来'],
    notFor: ['不是单个工具加载，是整组；只要一个工具也用这个'],
  },
  wps_route_intent: {
    aliases: ['意图识别', '该用哪个组', '这句话是干嘛'],
    notFor: ['不是执行工具，只返回"该加载哪个组"的建议'],
  },

  // ===== Word 近名/易混 =====
  wps_word_generate_toc: {
    aliases: ['生成目录', '插入目录', '自动目录'],
    notFor: ['与 wps_word_generate_doc_toc 近名：本工具走【段落大纲级别】生成目录'],
  },
  wps_word_generate_doc_toc: {
    aliases: ['生成目录', '文档目录'],
    notFor: ['与 wps_word_generate_toc 近名：本工具按【整篇文档】重建目录，会覆盖已有目录'],
  },
  wps_word_insert_table: {
    aliases: ['插入表格', '加表格', '建表'],
    examples: ['在文末插入 3 行 4 列表格'],
    notFor: ['改已有表格的样式/边框用 wps_word_format_table / wps_word_set_cell_borders；合并单元格用 wps_word_merge_table_cells'],
  },
  wps_word_revisions: {
    aliases: ['修订', '批注', '审阅', '接受修订', '拒绝修订'],
    notFor: ['只管"修订"；插入新批注用 wps_word_insert_comment'],
  },
  wps_word_insert_comment: {
    aliases: ['加批注', '插入批注', '评论'],
    notFor: ['查看/接受/拒绝已有修订用 wps_word_revisions'],
  },
  wps_word_add_heading: {
    aliases: ['加标题', '插入标题', '几级标题'],
    notFor: ['给已有段落套"标题样式"用 wps_word_apply_style'],
  },
  wps_word_apply_style: {
    aliases: ['套样式', '应用样式', '改成标题样式'],
    notFor: ['新建/改/删样式定义本身用 wps_word_add_style / modify_style / delete_style'],
  },
  wps_word_cross_reference: {
    aliases: ['交叉引用', '引用图表', '见图'],
    notFor: ['不是参考文献；插入索引条目用 wps_word_index'],
  },
  wps_word_index: {
    aliases: ['索引', '关键词索引', '索引表'],
    notFor: ['不是目录（目录用 generate_toc）；不是交叉引用（用 cross_reference）'],
  },

  // ===== Excel 近名/易混 =====
  wps_excel_zoom: {
    aliases: ['缩放', '显示比例'],
    notFor: ['与 wps_excel_set_zoom 同名能力（两者都设显示比例），优先用 set_zoom'],
  },
  wps_excel_set_zoom: {
    aliases: ['设置缩放', '显示比例'],
    notFor: ['与 wps_excel_zoom 重复，二选一即可'],
  },
  wps_excel_insert_row: {
    aliases: ['插入行', '插入一行'],
    notFor: ['一次插多行用 wps_excel_insert_rows'],
  },
  wps_excel_insert_rows: {
    aliases: ['插入多行', '批量插入行'],
    notFor: ['只插一行用 wps_excel_insert_row'],
  },
  wps_excel_delete_row: {
    aliases: ['删除行', '删一行'],
    notFor: ['批量删行用 wps_excel_delete_rows'],
  },
  wps_excel_delete_rows: {
    aliases: ['删除多行', '批量删行'],
  },
  wps_excel_set_cell_value: {
    aliases: ['写单元格', '填值', '设置单元格内容'],
    notFor: ['一次写一大片用 wps_excel_write_range；写公式用 wps_excel_set_formula'],
  },
  wps_excel_write_range: {
    aliases: ['批量写', '写区域', '一次写一片'],
    notFor: ['只写单个格子用 wps_excel_set_cell_value'],
  },
  wps_excel_set_formula: {
    aliases: ['写公式', '设置公式', '填公式'],
    notFor: ['算一次就走（不求值）用 wps_excel_evaluate_formula；公式报错排查用 wps_excel_diagnose_formula'],
  },
  wps_excel_merge_cells: {
    aliases: ['合并单元格'],
    notFor: ['这是【Excel】的；Word 表格合并用 wps_word_merge_table_cells'],
  },
  wps_excel_export_range_as_image: {
    aliases: ['区域转图片', '导出图片'],
    notFor: ['整份转 PDF 用 wps_convert_to_pdf'],
  },
  wps_excel_hide_column: {
    aliases: ['隐藏列'],
    notFor: ['隐藏行用 wps_excel_hide_rows；恢复显示用 wps_excel_show_columns'],
  },

  // ===== Word 表格 vs Excel 表格 =====
  wps_word_merge_table_cells: {
    aliases: ['合并单元格', '合并表格单元格'],
    notFor: ['这是【Word 表格】的；Excel 用 wps_excel_merge_cells'],
  },

  // ===== PPT：6 个 background 是重灾区 =====
  wps_ppt_set_background: {
    aliases: ['设置背景', '换背景'],
    notFor: ['纯色用 set_background_color；图片用 set_background_image；渐变用 set_background_gradient；当前页用 set_slide_background；母版用 set_master_background'],
  },
  wps_ppt_set_slide_background: {
    aliases: ['设置当前页背景'],
    notFor: ['所有页统一背景用 wps_ppt_set_background 或改母版 wps_ppt_set_master_background'],
  },
  wps_ppt_set_background_color: {
    aliases: ['背景色', '背景颜色', '纯色背景'],
  },
  wps_ppt_set_background_image: {
    aliases: ['背景图', '背景图片'],
  },
  wps_ppt_set_background_gradient: {
    aliases: ['渐变背景'],
  },
  wps_ppt_set_master_background: {
    aliases: ['母版背景'],
    notFor: ['只改一页用 wps_ppt_set_slide_background'],
  },
  // PPT：5 个 animation
  wps_ppt_set_animation: {
    aliases: ['设置动画'],
    notFor: ['新增动画用 add_animation / add_animation_preset；强调动画用 add_emphasis_animation；调整播放顺序用 set_animation_order；查看用 get_animations'],
  },
  wps_ppt_add_animation: {
    aliases: ['添加动画', '加动画'],
    notFor: ['改已有动画用 wps_ppt_set_animation'],
  },
  wps_ppt_add_emphasis_animation: {
    aliases: ['强调动画'],
  },
  // PPT：4 个 transition
  wps_ppt_set_transition: {
    aliases: ['切换效果', '切换动画', '翻页效果'],
    notFor: ['只改当前页用 set_slide_transition；一键套全部用 apply_transition_to_all；删掉用 remove_slide_transition'],
  },
  wps_ppt_set_slide_transition: {
    aliases: ['当前页切换效果'],
  },
  wps_ppt_apply_transition_to_all: {
    aliases: ['全部应用切换', '所有页切换效果'],
  },
  // PPT：3 个 image 插入
  wps_ppt_insert_image: {
    aliases: ['插入图片', '加图片'],
    notFor: ['按页插用 insert_slide_image / insert_ppt_image（同族）；替换已有图片用 replace_ppt_image'],
  },
  wps_ppt_replace_ppt_image: {
    aliases: ['替换图片', '换图'],
  },
  // PPT：3 个 chart
  wps_ppt_add_chart: {
    aliases: ['插入图表', '加图表'],
    notFor: ['同族还有 create_chart / insert_ppt_chart；改数据用 set_ppt_chart_data，改样式用 set_ppt_chart_style'],
  },
  wps_ppt_set_ppt_chart_data: {
    aliases: ['改图表数据'],
  },
  // PPT：3 个 beautify
  wps_ppt_beautify: {
    aliases: ['美化', '美化单页'],
    notFor: ['整篇美化用 beautify_all_slides / auto_beautify_slide'],
  },
  wps_ppt_beautify_all_slides: {
    aliases: ['美化全部', '整篇美化'],
  },
  // PPT 表格
  wps_ppt_insert_table: {
    aliases: ['插入表格', '加表格'],
    notFor: ['这是【PPT 里的表格】；Word 用 wps_word_insert_table，Excel 本来就是表格'],
  },
  // PPT 文本框
  wps_ppt_add_textbox: {
    aliases: ['加文本框', '插入文本框'],
    notFor: ['改已有文本框内容用 set_textbox_text；删掉用 delete_textbox'],
  },
  wps_ppt_set_slide_title: {
    aliases: ['设置标题', '改标题'],
    notFor: ['读取标题用 wps_ppt_get_slide_title；副标题用 set_slide_subtitle'],
  },

  // ===== 通用 =====
  wps_convert_to_pdf: {
    aliases: ['转PDF', '导出PDF', '转成PDF'],
    notFor: ['转成别的 Office 格式（docx/xlsx 互转）用 wps_convert_format'],
  },
  wps_convert_format: {
    aliases: ['格式转换', '转格式', 'docx转xlsx'],
    notFor: ['转 PDF 用 wps_convert_to_pdf'],
  },
  wps_common_status: {
    aliases: ['状态', '桥状态', '看看还活着吗', '诊断'],
    notFor: ['只是"看"；真要修复卡死用 wps_common_repair'],
  },
  wps_common_repair: {
    aliases: ['修复', '卡死了', '自救', '清场'],
    notFor: ['会关掉未保存的文档（自救动作），只想看状态用 wps_common_status'],
  },
  wps_common_save: {
    aliases: ['保存', '存盘'],
    notFor: ['另存到新路径用 wps_common_save_as'],
  },
  wps_common_save_as: {
    aliases: ['另存为', '存到'],
  },
  wps_execute_method: {
    aliases: ['直接调COM', '万能方法', '没有对应工具'],
    notFor: ['逃生舱：优先用专用工具，只有确实没有对应工具时才用它裸调 COM'],
  },

  // ===== 内置工具（mcp-server.ts 的 registerBuiltinTools）=====
  // 这一批在 server 里手写，不在 src/tools 下，机器拼不出中文别名 ——
  // wps_check_connection 之前别名是 0，搜「检查连接」命中 0 个。
  wps_check_connection: {
    aliases: ['检查连接', '连接状态', '通不通', '能连上吗'],
    examples: ['WPS 连上了吗', '连接断了吗'],
    notFor: ['只测连接；看桥进程与队列状态用 wps_common_status'],
  },
  wps_get_active_document: {
    aliases: ['当前文档', '活动文档', '打开的是哪个文档'],
    notFor: ['只查是哪个；切换目标文档用 wps_word_switch_document'],
  },
  wps_get_active_workbook: {
    aliases: ['当前工作簿', '活动工作簿', '打开的是哪个表格'],
    notFor: ['只查是哪个；切换目标工作簿用 wps_excel_switch_workbook'],
  },
  wps_get_active_presentation: {
    aliases: ['当前演示文稿', '活动演示', '打开的是哪个PPT'],
    notFor: ['只查是哪个；切换目标演示用 wps_ppt_switch_presentation'],
  },
  wps_insert_text: {
    aliases: ['插入文本', '插字', '在文档里打字'],
    notFor: ['往光标处插文本；替换已有内容用 wps_word_find_replace'],
  },
  wps_get_cell_value: {
    aliases: ['读单元格', '单元格值', '取格子里的值'],
    notFor: ['取单个格子；读整片区域用 wps_excel_read_range'],
  },
  wps_set_cell_value: {
    aliases: ['写单元格', '填格子', '设置单元格内容'],
    notFor: ['写单个格子；写整片区域用 wps_excel_write_range'],
  },
  wps_cache_data: {
    aliases: ['暂存数据', '缓存数据', '存中间结果'],
    examples: ['先把算出来的结果暂存起来'],
    notFor: ['只在本 MCP 进程内的内存暂存，不落盘、不跨会话'],
  },
  wps_get_cached_data: {
    aliases: ['取暂存数据', '读缓存', '取暂存结果'],
    notFor: ['读之前 wps_cache_data 暂存的东西；读表格内容用 wps_excel_read_range'],
  },
  wps_list_cache: {
    aliases: ['看缓存列表', '暂存了哪些'],
    notFor: ['只列 key；要值用 wps_get_cached_data'],
  },
  wps_clear_cache: {
    aliases: ['清缓存', '清空暂存', '删掉暂存数据'],
    notFor: ['只清 MCP 进程的内存暂存，不动你的文档'],
  },

  // ===== 第 4 期补齐的 Excel 高频缺口 =====
  wps_excel_get_sheet_info: {
    aliases: ['表详情', '表多大', '用到哪', '隐藏行列', '数据范围'],
    examples: ['这张表有多少行', 'A 列藏起来了吗'],
    notFor: ['只给表名与可见性用 wps_excel_get_sheet_list；读具体内容用 wps_excel_read_range'],
  },
  wps_excel_auto_fit: {
    aliases: ['自适应', '自动调整列宽', '自动列宽', '双击列宽', '合适的宽度'],
    examples: ['列宽按内容调一下', '把列宽自适应一下'],
    notFor: ['只设固定宽度用 wps_excel_set_column_width / set_row_height'],
  },
  wps_excel_set_wrap_text: {
    aliases: ['自动换行', '换行显示', '文字折行'],
    notFor: ['只改显示不改内容；合并单元格用 wps_excel_merge_cells'],
  },
  wps_excel_find_in_sheet: {
    aliases: ['表内查找', '在哪', '这个数在哪', '找一下这个值'],
    examples: ['"张伟"出现在哪几行', '编号 1001 在哪'],
    notFor: ['只读不改；要替换内容用 wps_excel_find_replace'],
  },
  wps_excel_delete_chart: {
    aliases: ['删图表', '删除图表', '把图删了', '图表删掉'],
    notFor: ['只改图表样式/数据用 wps_excel_update_chart'],
  },
  wps_excel_get_pivot_tables: {
    aliases: ['透视表列表', '有哪些透视表', '透视表现在什么样'],
    notFor: ['只查表名与区域；改透视表用 wps_excel_update_pivot_table'],
  },
  wps_excel_refresh_pivot_tables: {
    aliases: ['刷新透视表', '透视表更新', '数据没变'],
    examples: ['源数据改了，透视表数字没变'],
    notFor: ['只改透视表布局用 wps_excel_update_pivot_table'],
  },
  wps_excel_clear_pivot_table: {
    aliases: ['清空透视表', '透视表重置', '把透视表内容清掉'],
    notFor: ['保留位置与格式只清数据；彻底删掉透视表目前没有工具'],
  },
  wps_excel_set_sheet_page_setup: {
    aliases: ['页面设置', '打印设置', '横向打印', '缩放到一页', '纸张大小'],
    examples: ['横向打印', 'A4 纸', '缩放到一页宽'],
    notFor: ['只划打印区域用 wps_excel_set_print_area；每页重复表头用 wps_excel_set_sheet_print_titles'],
  },
  wps_excel_set_sheet_print_titles: {
    aliases: ['打印标题', '每页重复表头', '重复表头行', '每页都有表头'],
    examples: ['第二页开始也要有表头'],
    notFor: ['横向/纸张/缩放用 wps_excel_set_sheet_page_setup'],
  },
};

/** 把 wps_word_insert_table 拆成词根（丢掉应用前缀） */
function tokenize(name: string): string[] {
  return name
    .toLowerCase()
    .split(/[^a-z0-9]+/)
    .filter((t) => t && !APP_TOKENS.has(t));
}

/** 取词根的中文说法；没有就原样返回（英文原词仍可被搜到） */
function zh(tok: string): string[] {
  return TOKEN_ZH[tok] ?? [tok];
}

/**
 * 自动拼别名：动词各种说法 × 后面名词的规范说法。
 * 组合数封顶 6，避免长名字（6~7 个词根）炸出几十条。
 */
function deriveAliases(name: string): string[] {
  const toks = tokenize(name);
  if (!toks.length) return [];

  const head = zh(toks[0]);
  const tail = toks.slice(1).map((t) => zh(t)[0]);
  const out = new Set<string>();

  for (const h of head) {
    out.add([h, ...tail].join(''));
    if (out.size >= 6) break;
  }
  // 名词单独成条，让"表格""批注"这类裸名词也能命中
  for (const t of toks.slice(1)) {
    for (const z of zh(t)) out.add(z);
  }
  return [...out].filter((s) => s && !/^[a-z0-9]+$/.test(s));
}

export interface ToolMeta {
  /** 中文别名（自动拼 + 手工补） */
  aliases: string[];
  /** 正例：什么样的诉求该用它 */
  examples?: string[];
  /** 反例：它【不】干什么，以及该用谁 */
  notFor?: string[];
  /** 检索用的大杂烩（名字 + 描述 + 别名，全小写） */
  haystack: string;
}

const cache = new Map<string, ToolMeta>();

/** 取（并缓存）某个工具的检索元数据 */
export function getToolMeta(def: ToolDefinition): ToolMeta {
  const hit = cache.get(def.name);
  if (hit) return hit;

  const auto = deriveAliases(def.name);
  const manual = CURATED[def.name]?.aliases ?? [];
  const aliases = [...new Set([...manual, ...auto])].slice(0, 10);
  const c = CURATED[def.name];

  const meta: ToolMeta = {
    aliases,
    ...(c?.examples?.length ? { examples: c.examples } : {}),
    ...(c?.notFor?.length ? { notFor: c.notFor } : {}),
    haystack: [def.name, def.description ?? '', ...aliases].join(' ').toLowerCase(),
  };
  cache.set(def.name, meta);
  return meta;
}

/** 清空缓存（测试用；生产里工具在启动时注册完就不再变） */
export function __clearToolMetaCache(): void {
  cache.clear();
}
