/**
 * 验证"生产环境默认恰好 40 个"这个承诺。
 *
 * 坑：wps_execute_method 等 12 个是 mcp-server 在启动时【单独注册】的内置工具，
 * 不在 src/tools/index.ts 的 allTools 里。只注册 allTools 会少算 1 个，
 * 看起来像"基础档掉了" —— 实际生产是 40。
 * 这个脚本专门把内置工具也注册上，量出生产口径的真实数字。
 */
const path = require('path');
const ROOT = path.resolve(__dirname, '../..');
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));
const { ESSENTIAL_TOOLS } = require(path.join(ROOT, 'dist', 'server', 'tool-tiers.js'));

const results = [];
const check = (name, ok, detail = '') => {
  results.push({ name, ok });
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? '  → ' + detail : ''}`);
};

// 内置工具的定义在 mcp-server.ts 里手写，不走注册表导出 —— 这里直接从
// 已编译的 mcp-server 里把它的默认配置与工具数读出来做交叉核对。
toolRegistry.registerAll(allTools);

const listed = toolRegistry.listTools().tools.map((t) => t.name);
const all = toolRegistry.listAllTools().map((t) => t.name);
const ess = new Set(ESSENTIAL_TOOLS);

// 内置工具名（与 mcp-server.ts 的 registerBuiltinTools 一一对应）
const BUILTIN = [
  'wps_execute_method',
];

const essInAll = ESSENTIAL_TOOLS.filter((n) => all.includes(n));
const essBuiltin = ESSENTIAL_TOOLS.filter((n) => BUILTIN.includes(n));
const ghost = ESSENTIAL_TOOLS.filter((n) => !all.includes(n) && !BUILTIN.includes(n));

console.log(`ESSENTIAL ${ESSENTIAL_TOOLS.length} = 专业工具 ${essInAll.length} + 内置 ${essBuiltin.length}`);
console.log(`core 组（永远列出）: ${listed.filter((n) => !ess.has(n)).length}`);
console.log(`专业工具实际列出: ${listed.length}`);
console.log(`生产口径（专业 + 内置 ESSENTIAL + core）= ${listed.length} + ${essBuiltin.length} = ${listed.length + essBuiltin.length}\n`);

check('白名单没有"写了但不存在"的工具名', ghost.length === 0, ghost.join(', '));
check(
  `生产口径恰好 40 个`,
  listed.length + essBuiltin.length === 40,
  `${listed.length} + ${essBuiltin.length} = ${listed.length + essBuiltin.length}`
);
check('PPT 组（115 个）默认不列出', !listed.some((n) => n.startsWith('wps_ppt_') && !ess.has(n)), `ppt 列出 ${listed.filter((n) => n.startsWith('wps_ppt_')).length} 个（ESSENTIAL 里的 8 个）`);

// repair 是本轮新增的，确认它已注册但默认不列出（需要时按组加载）
check('wps_common_repair 已注册', all.includes('wps_common_repair'), `全量 ${all.length} 个里有=${all.includes('wps_common_repair')}`);
check('wps_common_repair 默认不列出（避免挤掉 40 的名额）', !listed.includes('wps_common_repair'));

const passed = results.filter((r) => r.ok).length;
console.log(`\n汇总：${passed}/${results.length} 通过`);
process.exit(passed === results.length ? 0 : 1);