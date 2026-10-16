/**
 * Input: 无
 * Output: 工具组（按需加载）的断言结果
 * Pos: 工具组的手工验证脚本。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 用法（先 npm run build）：node scripts/manual/tool-groups-test.js
 *
 * 说明：直接使用注册表单例 + allTools，不起 MCP 传输层。
 * 通知（notifications/tools/list_changed）不在这里测 —— 那需要真客户端，
 * 这里只断言"变更事件确实被触发"。
 */

const path = require('path');
const ROOT = path.resolve(__dirname, '../..');

const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));
const { onToolGroupsChanged } = require(path.join(ROOT, 'dist', 'tools', 'common', 'toolgroups.js'));

const results = [];
function check(name, cond, detail) {
  results.push({ name, ok: !!cond });
  console.log(`  ${cond ? '✅' : '❌'} ${name}${detail ? '  → ' + detail : ''}`);
}

toolRegistry.registerAll(allTools);

console.log('=== 工具组默认状态 ===');
const groups0 = toolRegistry.getGroups();
console.log('  ' + groups0.map((g) => `${g.group}(${g.loaded ? '已加载' : '未加载'},${g.total})`).join('  '));

const listed0 = toolRegistry.listTools().tools.length;
const total0 = toolRegistry.listAllTools().length;
console.log(`  列出 ${listed0} / 共 ${total0}`);

check('总工具数 > 250（此处未注册 mcp-server 的 12 个内置工具）', total0 > 250, `${total0}`);
check('core 已加载', groups0.find((g) => g.group === 'core')?.loaded === true);
check('ppt 默认未加载', groups0.find((g) => g.group === 'ppt')?.loaded === false);
check(
  '列出的数量 = 已加载组的工具数（即少于总数）',
  listed0 < total0,
  `${listed0} < ${total0}`
);

const listedNames = new Set(toolRegistry.listTools().tools.map((t) => t.name));
check('未列出的工具里不含任何 wps_ppt_ 前缀', ![...listedNames].some((n) => n.startsWith('wps_ppt_')));
check('工具组管理工具本身始终列出', listedNames.has('wps_list_tool_groups') && listedNames.has('wps_load_tool_group'));

console.log('\n=== 搜索（含未加载组）===');
const hits = toolRegistry.searchTools('ppt', { limit: 5 });
console.log(`  搜 "ppt" 命中 ${hits.length} 条（前 5）: ` + hits.map((h) => `${h.name}[${h.group}${h.listed ? '' : ',未列出'}]`).join(', '));
check('能搜到 ppt 组的工具', hits.length > 0 && hits.some((h) => h.group === 'ppt'));
check('搜到的 ppt 工具标记为未列出', hits.some((h) => h.group === 'ppt' && h.listed === false));

const hits2 = toolRegistry.searchTools('导出图片');
check('搜中文关键词不崩', Array.isArray(hits2), `命中 ${hits2.length}`);

console.log('\n=== 加载 ppt 组 ===');
let notified = 0;
onToolGroupsChanged(() => {
  notified += 1;
});

const load = toolRegistry.loadGroup('ppt');
check('loadGroup 返回 ok', load.ok === true);
check('loadGroup 报告发生了变化', load.changed === true);
check('变更事件被触发一次', notified === 1, `notified=${notified}`);

const listed1 = toolRegistry.listTools().tools.length;
check('列出数量增加', listed1 > listed0, `${listed0} → ${listed1}`);
check('加载后 ppt 工具出现在列表里', [...toolRegistry.listTools().tools.map((t) => t.name)].some((n) => n.startsWith('wps_ppt_')));

console.log('\n=== 重复加载不应再触发通知 ===');
const load2 = toolRegistry.loadGroup('ppt');
check('重复 load 报告未变化', load2.changed === false);
check('通知次数没有增加', notified === 1, `notified=${notified}`);

console.log('\n=== 保护 core 组 ===');
const unloadCore = toolRegistry.unloadGroup('core');
check('core 不允许卸载', unloadCore.ok === false, String(unloadCore.error || ''));

console.log('\n=== 卸载 ppt 组 ===');
const un = toolRegistry.unloadGroup('ppt');
check('unload 返回 ok', un.ok === true);
check('卸载触发通知', notified === 2, `notified=${notified}`);
check('列出数量回到初始', toolRegistry.listTools().tools.length === listed0, `${toolRegistry.listTools().tools.length}`);

console.log('\n=== 不存在的组 ===');
check('加载不存在的组失败', toolRegistry.loadGroup('nope').ok === false);
check('卸载不存在的组失败', toolRegistry.unloadGroup('nope').ok === false);

const passed = results.filter((r) => r.ok).length;
console.log(`\n汇总：${passed}/${results.length} 通过`);
process.exit(passed === results.length ? 0 : 1);
