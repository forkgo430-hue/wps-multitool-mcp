/**
 * 工具注解验证
 * -----------
 * 验四件事：
 *   1) 285 个工具是不是全覆盖（拿不到动作名的不能给错注解，宁可没有）
 *   2) 抽查关键工具的注解值对不对（只读的别标成写，删除类的要标破坏）
 *   3) 两个判定函数的已知陷阱有没有被修掉：
 *      find_replace 曾经因 find 前缀被误判成【只读】，但它真会改文档；
 *      evaluateFormula（PS 侧驼峰）与 evaluate_formula（工具名下划线）必须同判
 *   4) 分布数字（文档里写的都以这里为准，别手写）
 */

const path = require('path');
const ROOT = path.resolve(__dirname, '..', '..');

let pass = 0;
let fail = 0;
const eq = (label, got, want) => {
  if (JSON.stringify(got) === JSON.stringify(want)) {
    pass++;
    console.log(`  ✓ ${label} = ${JSON.stringify(got)}`);
  } else {
    fail++;
    console.log(`  ✗ ${label}：得到 ${JSON.stringify(got)}，期望 ${JSON.stringify(want)}`);
  }
};
const ok = (label, cond, extra) => {
  if (cond) {
    pass++;
    console.log(`  ✓ ${label}${extra ? ' — ' + extra : ''}`);
  } else {
    fail++;
    console.log(`  ✗ ${label}${extra ? ' — ' + extra : ''}`);
  }
};

const { getToolAnnotations, annotationSummary, annotationCoverage } = require(
  path.join(ROOT, 'dist', 'server', 'tool-annotations.js')
);
const {
  isReadOnlyAction,
  isDestructiveAction,
  isIdempotentAction,
} = require(path.join(ROOT, 'dist', 'bridge', 'cancellable.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { TOOL_ACTIONS } = require(
  path.join(ROOT, 'dist', 'server', 'tool-annotations.generated.js')
);

// 必须走 server.start() 的真实注册路径：registerBuiltinTools() 的 12 个
// + registerAll(allTools) 的 273 个。只 registerAll 会少一截，
// 那样验证出来的"全覆盖"是假的（这坑踩过一次）。
const inst = new WpsMcpServer();
inst.registerBuiltinTools();
toolRegistry.registerAll(allTools);
const names = toolRegistry.listAllTools().map((t) => t.name);

console.log('\n=== A. 覆盖率与三方对账 ===');
const cov = annotationCoverage(names);
eq('真实注册的工具总数', cov.total, 285);
eq('拿到注解的个数', cov.annotated, cov.total);
eq('注解映射表条目数', Object.keys(TOOL_ACTIONS).length, cov.total);
const mapOnly = Object.keys(TOOL_ACTIONS).filter((n) => !names.includes(n));
ok('注解表里没有已下线的工具', mapOnly.length === 0, mapOnly.slice(0, 5).join(','));
// 这条是防漂移的关键：注册表加了工具却没重跑生成器，这里立刻红
eq('每个工具都有动作映射（无遗漏）', names.filter((n) => !(n in TOOL_ACTIONS)).length, 0);

console.log('\n=== B. 判定函数的已知陷阱 ===');
eq("isReadOnlyAction('findReplace')", isReadOnlyAction('findReplace'), false);
eq("isReadOnlyAction('find_replace')", isReadOnlyAction('find_replace'), false);
eq("isReadOnlyAction('evaluateFormula')", isReadOnlyAction('evaluateFormula'), true);
eq("isReadOnlyAction('evaluate_formula')", isReadOnlyAction('evaluate_formula'), true);
eq("isReadOnlyAction('diagnoseFormula')", isReadOnlyAction('diagnoseFormula'), true);
eq("isReadOnlyAction('diagnose_formula')", isReadOnlyAction('diagnose_formula'), true);
ok(
  '驼峰与下划线两种拼法判定一致',
  isReadOnlyAction('evaluateFormula') === isReadOnlyAction('evaluate_formula') &&
    isReadOnlyAction('getCellText') === isReadOnlyAction('get_cell_text')
);

console.log('\n=== C. 抽查工具注解 ===');
const cases = [
  ['wps_word_get_cell_text', { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false }],
  ['wps_excel_get_cell_value', { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false }],
  ['wps_common_status', { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false }],
  ['wps_excel_write_range', { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: false }],
  ['wps_excel_insert_row', { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: false }],
  ['wps_common_save_as', { readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: false }],
  ['wps_convert_to_pdf', { readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: false }],
];
for (const [name, want] of cases) {
  if (!names.includes(name)) {
    // 名字对不上就直接判失败，别静默跳过（工具改名了这里要跟着改）
    fail++;
    console.log(`  ✗ 工具不存在：${name}`);
    continue;
  }
  eq(name, getToolAnnotations(name), want);
}

// 内置的 12 个（registerBuiltinTools）也必须带注解 —— 之前正是漏了这一批
const builtin = names.filter((n) =>
  ['wps_check_connection', 'wps_get_active_document', 'wps_insert_text',
   'wps_get_active_workbook', 'wps_get_cell_value', 'wps_set_cell_value',
   'wps_get_active_presentation', 'wps_execute_method', 'wps_cache_data',
   'wps_get_cached_data', 'wps_list_cache', 'wps_clear_cache'].includes(n)
);
const builtinNoAnn = builtin.filter((n) => !getToolAnnotations(n));
eq('server 直注册的 12 个工具都有注解', builtinNoAnn.length, 0);
eq('wps_execute_method 标为破坏性（能执行任意动作）',
   getToolAnnotations('wps_execute_method')?.destructiveHint, true);

const delTools = names.filter((n) => /delete|remove|clear/.test(n));
const delBad = delTools.filter((n) => getToolAnnotations(n)?.destructiveHint !== true);
ok(
  '删除类工具全部标了 destructiveHint',
  delBad.length === 0,
  `检查 ${delTools.length} 个，不合格 ${delBad.length}${delBad.length ? '：' + delBad.slice(0, 5).join(',') : ''}`
);

const roTools = names.filter((n) => getToolAnnotations(n)?.readOnlyHint === true);
const roBad = roTools.filter((n) => /insert|delete|write_|set_/.test(n));
ok('只读工具里没有名字像写入的', roBad.length === 0, roBad.slice(0, 5).join(','));

console.log('\n=== D. 分布（文档数字以这里为准） ===');
const s = annotationSummary(names);
console.log(`  总数 ${s.total} / 已注解 ${s.annotated}`);
console.log(`  只读 ${s.readOnly} | 破坏性 ${s.destructive} | 非幂等 ${s.nonIdempotent}`);
console.log(`  写入类 ${s.total - s.readOnly}`);

console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
process.exit(fail === 0 ? 0 : 1);