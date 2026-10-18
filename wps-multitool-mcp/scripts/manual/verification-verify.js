/**
 * 验证"每个工具都有验证状态、默认档里没有未验证的工具"这两个承诺。
 *
 * 为什么需要门禁而不是靠自觉
 * ------------------------
 * 276 个工具时没人管验证状态也没出事；到 306 个时 PPT 已经占了一整组，
 * 却一个都没在真机上跑过。这种"沉默的缺口"不会自己报警 —— 必须有人因为数字对不上而失败。
 *
 * 查五件事：
 *   1) 覆盖率：注册表里每个工具都能查到状态（307 个里漏 1 个就红）
 *   2) 生成物与 manifest 一致（改了 json 忘了跑生成器）
 *   3) 默认档里没有 static 工具（默认档不能替未验证的东西背书）
 *   4) 默认档数量 = core 组 + 基础档白名单里 live 的那批（数字不许手写）
 *   5) ppt 组里的 live 必须在 manifest.tools 里逐条点名，不能靠组默认浑过去
 */
const path = require('path');
const fs = require('fs');
const ROOT = path.resolve(__dirname, '../..');

const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));
const { groupOfTool, ALWAYS_LOADED_GROUP } = require(path.join(ROOT, 'dist', 'types', 'tools.js'));
const { TOOL_VERIFICATION } = require(path.join(ROOT, 'dist', 'server', 'tool-verification.generated.js'));
const { verificationOf, isLiveVerified } = require(path.join(ROOT, 'dist', 'server', 'tool-verification.js'));
const { ESSENTIAL_TOOLS } = require(path.join(ROOT, 'dist', 'server', 'tool-tiers.js'));

const results = [];
const check = (name, ok, detail = '') => {
  results.push({ name, ok });
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? '  → ' + detail : ''}`);
};

// 真实注册路径：漏 registerBuiltinTools 就少 12 个（老坑）
const inst = new WpsMcpServer();
inst.registerBuiltinTools();
toolRegistry.registerAll(allTools);

const all = toolRegistry.listAllTools();
const names = all.map((d) => d.name);
const manifest = JSON.parse(fs.readFileSync(path.join(__dirname, 'tool-verification.json'), 'utf8'));

// 1) 覆盖率
const missing = names.filter((n) => TOOL_VERIFICATION[n] === undefined);
check(`验证状态覆盖 ${names.length - missing.length}/${names.length}`, missing.length === 0, missing.slice(0, 8).join(', '));

// 2) 生成物与 manifest 一致
const expected = {};
for (const d of all) {
  const g = groupOfTool(d);
  const s = manifest.tools[d.name] || manifest.groups[g];
  if (s) expected[d.name] = s;
}
const drift = names.filter((n) => expected[n] && TOOL_VERIFICATION[n] !== expected[n]);
check(
  'tool-verification.generated.ts 与 manifest 一致',
  drift.length === 0,
  drift.length ? `${drift.slice(0, 5).join(', ')}（跑 node scripts/manual/gen-tool-verification.js）` : ''
);

// 3) 默认档里不许有 static
const listed = toolRegistry.listTools().tools.map((d) => d.name);
const staticListed = listed.filter((n) => !isLiveVerified(n));
check('默认档里没有未验证的工具', staticListed.length === 0, staticListed.join(', '));

// 4) 默认档数量 = core 组 + 基础档 live
const coreNames = all.filter((d) => groupOfTool(d) === ALWAYS_LOADED_GROUP).map((d) => d.name);
const essentialLive = ESSENTIAL_TOOLS.filter((n) => isLiveVerified(n) && names.includes(n));
const expectedListed = new Set([...coreNames, ...essentialLive]);
check(
  `默认列出 ${listed.length} = core ${coreNames.length} + 基础档 live ${essentialLive.length}`,
  listed.length === expectedListed.size && listed.every((n) => expectedListed.has(n)),
  `实际 ${listed.length}，期望 ${expectedListed.size}`
);

// 5) ppt 组里的 live 必须逐条点名
const pptLive = all
  .filter((d) => groupOfTool(d) === 'ppt' && isLiveVerified(d.name))
  .map((d) => d.name);
const pptUnnamed = pptLive.filter((n) => !manifest.tools[n]);
check(
  'PPT 的 live 全部逐条写进 manifest.tools（不能靠组默认）',
  pptUnnamed.length === 0,
  pptUnnamed.join(', ')
);

// 6) manifest 里的死条目（写了但工具不存在）
const registered = new Set(names);
const dead = Object.keys(manifest.tools || {}).filter((n) => !registered.has(n));
check('manifest.tools 没有死条目', dead.length === 0, dead.join(', '));

const live = names.filter((n) => verificationOf(n) === 'live').length;
const passed = results.filter((r) => r.ok).length;
console.log(`\n真机验证 ${live}/${names.length}（static ${names.length - live}）`);
console.log(`汇总：${passed}/${results.length} 通过`);
process.exit(passed === results.length ? 0 : 1);
