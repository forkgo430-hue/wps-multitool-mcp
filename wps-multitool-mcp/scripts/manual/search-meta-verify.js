/**
 * Input: 无
 * Output: 中文关键词检索 + 别名/反例元数据的独立验证
 * Pos: 工具检索元数据核验脚本。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 验三件事
 * --------
 * 1) 中文关键词能不能搜到（注意：工具描述本来就是中文，所以这条【不是】从 0 到 1，
 *    别写成"以前搜不到"——实测以前也搜得到）。这里量的是真实增量：关掉别名后
 *    少捞到多少。
 * 2) 近名工具（6 个 background / 5 个 animation / 2 个 toc / 2 个 zoom）
 *    搜出来之后，靠前的命中带不带 notFor（模型分辨的依据，这是本层的【主力价值】）。
 * 3) 打分顺序：别名是机器拼的，不能压过人写的描述命中。
 */

const path = require('path');
const ROOT = path.resolve(__dirname, '../..');

const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { getToolMeta, __clearToolMetaCache } = require(path.join(ROOT, 'dist', 'server', 'tool-metadata.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));

// 走真实注册路径：registerBuiltinTools() 的 12 个 + src/tools 的 273 个。
// 只 registerAll 会漏掉内置那一批 —— wps_check_connection 的别名就是这么漏成 0 的，
// 搜「检查连接」命中 0 个。
const _inst = new WpsMcpServer();
_inst.registerBuiltinTools();
toolRegistry.registerAll(allTools);

let pass = 0;
let fail = 0;
function ok(label, cond, extra) {
  if (cond) {
    pass++;
    console.log(`  ✅ ${label}${extra ? '  ' + extra : ''}`);
  } else {
    fail++;
    console.log(`  ❌ ${label}${extra ? '  ' + extra : ''}`);
  }
}

const ALL = toolRegistry.listAllTools();

// 对照组：只用 name + description（模拟加别名之前的行为）
function legacySearch(kw) {
  const k = kw.toLowerCase();
  return ALL.filter((t) => {
    const def = toolRegistry.getTool(t.name)?.definition;
    if (!def) return false;
    return def.name.toLowerCase().includes(k) || (def.description || '').toLowerCase().includes(k);
  }).map((t) => t.name);
}

console.log('=== A. 中文关键词：别名到底补上了什么 ===');
const CN_CASES = [
  ['插入表格', 'wps_word_insert_table'],
  ['批注', 'wps_word_insert_comment'],
  ['背景色', 'wps_ppt_set_background_color'],
  ['动画', 'wps_ppt_add_animation'],
  ['转PDF', 'wps_convert_to_pdf'],
  ['目录', 'wps_word_generate_toc'],
  ['合并单元格', 'wps_excel_merge_cells'],
  ['公式', 'wps_excel_set_formula'],
  ['页眉', 'wps_word_insert_header'],
  ['幻灯片', 'wps_ppt_add_slide'],
];
for (const [kw, expect] of CN_CASES) {
  const before = legacySearch(kw);
  const after = toolRegistry.searchTools(kw, { limit: 50 }).map((h) => h.name);
  const hitIdx = after.indexOf(expect);
  ok(
    `搜"${kw}" → ${expect}`,
    hitIdx >= 0,
    `（旧: ${before.length} 条命中 / 新: ${after.length} 条，目标排第 ${hitIdx + 1}）`
  );
}

console.log('\n=== B. 真实增量：别名多捞到了什么（对照组）===');
// 注意前提：工具描述是中文，所以旧逻辑本来就有命中。这里量"只靠别名才捞到"的那部分。
function viaAlias(kw) {
  const k = kw.toLowerCase();
  return ALL
    .map((t) => toolRegistry.getTool(t.name)?.definition)
    .filter(Boolean)
    .filter((d) => getToolMeta(d).aliases.join(' ').toLowerCase().includes(k))
    .map((d) => d.name);
}
let gained = 0;
for (const [kw] of CN_CASES) {
  const L = new Set(legacySearch(kw));
  const only = viaAlias(kw).filter((n) => !L.has(n));
  gained += only.length;
  console.log(`  "${kw}"：描述命中 ${L.size}，别名额外 +${only.length}${only.length ? ' → ' + only.slice(0, 3).join('、') : ''}`);
}
ok('别名确实带来增量（不吹牛：是锦上添花，不是从 0 到 1）', gained > 0, `（合计 +${gained} 对）`);
ok('增量不靠吹：至少半数关键词本来就有描述命中', CN_CASES.filter(([kw]) => legacySearch(kw).length > 0).length >= CN_CASES.length / 2);

console.log('\n=== C. 近名工具：notFor 有没有给到 ===');
const AMBIG = [
  ['背景', ['wps_ppt_set_background', 'wps_ppt_set_background_color', 'wps_ppt_set_master_background']],
  ['动画', ['wps_ppt_set_animation', 'wps_ppt_add_animation']],
  ['目录', ['wps_word_generate_toc', 'wps_word_generate_doc_toc']],
  ['缩放', ['wps_excel_zoom', 'wps_excel_set_zoom']],
];
for (const [kw, names] of AMBIG) {
  const hits = toolRegistry.searchTools(kw, { limit: 30 });
  const withNotFor = hits.filter((h) => h.notFor && h.notFor.length);
  ok(`搜"${kw}" 靠前命中带 notFor`, withNotFor.length > 0,
     `（${withNotFor.length} 条带反例：${withNotFor.map((h) => h.name).slice(0, 3).join('、')}）`);
  const covered = names.filter((n) => hits.some((h) => h.name === n));
  ok(`搜"${kw}" 能同时捞到近名组（${names.length} 个中 ${covered.length} 个）`, covered.length >= 2);
}

console.log('\n=== D. 命中字段 ===');
const h1 = toolRegistry.searchTools('插入表格', { limit: 3 })[0];
ok('命中带 aliases', Array.isArray(h1.aliases) && h1.aliases.length > 0, JSON.stringify(h1.aliases));
ok('命中带 matchedBy', typeof h1.matchedBy === 'string', `matchedBy=${h1.matchedBy}`);
const hNoHint = toolRegistry.searchTools('背景', { limit: 10, hints: false });
ok('hints=false 时不给 examples/notFor', hNoHint.every((h) => !h.examples && !h.notFor));

// 别名是机器拼的（会串味），不能压过人写的描述命中
const mixed = toolRegistry.searchTools('合并单元格', { limit: 20 });
const ranks = mixed.map((h) => h.matchedBy);
const firstAlias = ranks.indexOf('alias');
const lastDesc = ranks.lastIndexOf('description');
ok(
  '打分：所有 description 命中都排在 alias 命中之前',
  firstAlias === -1 || lastDesc === -1 || lastDesc < firstAlias,
  `（顺序：${ranks.join(',')}）`
);

console.log('\n=== E. 别名覆盖率（不是只给 40 个基础档打的补丁）===');
let withZh = 0;
for (const t of ALL) {
  const def = toolRegistry.getTool(t.name)?.definition;
  if (!def) continue;
  const m = getToolMeta(def);
  if (m.aliases.some((a) => /[一-龥]/.test(a))) withZh++;
}
const rate = ((withZh / ALL.length) * 100).toFixed(1);
ok(`中文别名覆盖 ≥90% 的工具`, withZh / ALL.length >= 0.9, `（${withZh}/${ALL.length} = ${rate}%）`);

// 内置的 12 个（mcp-server 手写，不在 src/tools 下）必须有中文别名。
// 这批以前没进覆盖统计，wps_check_connection 的别名是 0 —— 搜「检查连接」命中 0 个。
const BUILTIN = [
  'wps_check_connection', 'wps_get_active_document', 'wps_insert_text',
  'wps_get_active_workbook', 'wps_get_cell_value', 'wps_set_cell_value',
  'wps_get_active_presentation', 'wps_execute_method',
  'wps_cache_data', 'wps_get_cached_data', 'wps_list_cache', 'wps_clear_cache',
];
const noAlias = BUILTIN.filter((n) => {
  const def = toolRegistry.getTool(n)?.definition;
  if (!def) return true;
  return !getToolMeta(def).aliases.some((a) => /[一-龥]/.test(a));
});
ok('内置 12 个工具都有中文别名', noAlias.length === 0, noAlias.join(','));
ok('搜「检查连接」能命中 wps_check_connection',
   toolRegistry.searchTools('检查连接', { limit: 5, hints: false }).some((h) => h.name === 'wps_check_connection'));

console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
process.exit(fail === 0 ? 0 : 1);
