/**
 * 工具注解生成器（一次性，跑完生成 src/server/tool-annotations.generated.ts）
 *
 * 要解决的问题
 * ------------
 * 273 个工具一个 MCP 注解都没有（readOnlyHint / destructiveHint / idempotentHint /
 * openWorldHint 全空）。而靠名字判断读写不可靠：isReadOnlyAction 只能确认 37 个只读。
 * 模型只能靠读描述去猜哪个能动文档。
 *
 * 怎么拿到"这个工具到底会发什么命令"
 * ----------------------------------
 * 静态扫源码不行 —— 工具的 definition 与 handler 是分开 export 的，动作名写在
 * handler 内部靠手敲字符串，顺序也对不上（一个工具还可能发多条命令）。
 *
 * 所以用运行时观测：把 wpsClient 的最终出口 invokeAction 换成假实现，
 * 逐个跑一遍 handler，记下它发了哪些动作名，然后一律返回 success=false，
 * 让 handler 走它自己的失败分支。
 * **全程不碰桥、不碰 WPS、不碰任何真实文档。**
 *
 * 产物只存「工具名 → 动作名列表」，注解本身在运行时由 cancellable.ts 的判定函数算，
 * 这样改了判定规则不必重新生成。
 *
 * 用法：node scripts/manual/gen-tool-annotations.js
 */

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
const OUT = path.join(ROOT, 'src', 'server', 'tool-annotations.generated.ts');

const { wpsClient } = require(path.join(ROOT, 'dist', 'client', 'wps-client.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));

// ---- 1. 换掉真正的出口，装上记录器 ----------------------------------------
let recording = null;
const origInvoke = wpsClient.invokeAction.bind(wpsClient);
wpsClient.invokeAction = async (action, params, appKey) => {
  if (recording) recording.actions.add(action);
  return { success: false, error: '__annotation_probe__' };
};
// openFile/saveFile 走 invokeAction，但如果将来改成直连，记得同步 patch

(async () => {
  toolRegistry.registerAll(allTools);
  const tools = toolRegistry.listAllTools();

  const map = {};
  const missed = [];

  for (const t of tools) {
    const name = t.name;
    const rec = { actions: new Set() };
    recording = rec;
    try {
      const tool = toolRegistry.getTool(name);
      // 加超时：有些 handler 会做重循环或长 sleep，观测脚本不该被它们拖住。
      // 超时不算失败 —— 已经记到的动作名照样算数。
      await Promise.race([
        tool.handler({}),
        new Promise((_, rej) => setTimeout(() => rej(new Error('__probe_timeout__')), 150)),
      ]);
    } catch (e) {
      // 参数校验就抛错的（没走到调用出口），走下面的名字推断兜底
    }
    recording = null;
    if (rec.actions.size) {
      map[name] = [...rec.actions];
    } else {
      missed.push(name);
    }
    if ((tools.indexOf(t) + 1) % 50 === 0) {
      console.log(`  …${tools.indexOf(t) + 1}/${tools.length}`);
    }
  }

  // ---- 2. 没被观测到的，用工具名推断动作名（兜底，标注来源） ----------------
  // wps_excel_get_cell_value -> getCellValue
  const infer = (toolName) => {
    const act = toolName
      .replace(/^wps_/, '')
      .replace(/^(word|excel|ppt|common)_/, '');
    return act.replace(/_([a-z0-9])/g, (_m, c) => c.toUpperCase());
  };
  for (const name of missed) {
    map[name] = [infer(name)];
  }

  // ---- 3. 落盘 ------------------------------------------------------------
  const lines = [];
  lines.push('/**');
  lines.push(' * 由 scripts/manual/gen-tool-annotations.js 生成，请勿手改。');
  lines.push(' * 改法：改那个脚本再跑一次。');
  lines.push(' *');
  lines.push(' * 内容：MCP 工具名 → 它在 wps-com.ps1 里实际会触发的动作名。');
  lines.push(' * 注解本身不在这里写死，由 cancellable.ts 的判定函数在运行时算，');
  lines.push(' * 这样改了判定规则不必重新生成本文件。');
  lines.push(' */');
  lines.push('');
  lines.push('export const TOOL_ACTIONS: Readonly<Record<string, readonly string[]>> = {');
  for (const name of Object.keys(map).sort()) {
    const acts = map[name].map((a) => `'${a}'`).join(', ');
    lines.push(`  ${JSON.stringify(name)}: [${acts}],`);
  }
  lines.push('};');
  lines.push('');
  fs.writeFileSync(OUT, lines.join('\n'), 'utf8');

  // ---- 4. 报告 ------------------------------------------------------------
  const multi = Object.entries(map).filter(([, a]) => a.length > 1);
  console.log(`工具总数        : ${tools.length}`);
  console.log(`运行时观测到动作: ${tools.length - missed.length}`);
  console.log(`名字推断兜底    : ${missed.length}${missed.length ? ' → ' + missed.slice(0, 8).join(', ') : ''}`);
  console.log(`一工具多动作    : ${multi.length}`);
  for (const [n, a] of multi.slice(0, 6)) console.log(`   ${n} → ${a.join(', ')}`);
  console.log(`\n已写入 ${path.relative(ROOT, OUT)}`);
  process.exit(0);
})().catch((e) => {
  console.error('生成失败:', e);
  process.exit(1);
});