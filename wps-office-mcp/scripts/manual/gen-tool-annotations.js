/**
 * 工具注解生成器（一次性，跑完生成 src/server/tool-annotations.generated.ts）
 *
 * 要解决的问题
 * ------------
 * 285 个工具一个 MCP 注解都没有（readOnlyHint / destructiveHint / idempotentHint /
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
 * 覆盖范围：必须走【真实的注册路径】
 * ------------------------------------
 * 只 registerAll(allTools) 只拿到 273 个，但 mcp-server.ts 的 registerBuiltinTools()
 * 还 this.registry.register() 了 12 个（wps_execute_method / wps_cache_data /
 * wps_get_cell_value…），start() 里两步都跑，真实运行时是 285 个。
 * 少做一步，注解表就会跟线上注册表对不上 —— 这种"两处维护"的漂移已经吃过一次亏了。
 *
 * 产物只存「工具名 → 动作名列表」，注解本身在运行时由 cancellable.ts 的判定函数算，
 * 这样改了判定规则不必重新生成。
 *
 * 用法：node scripts/manual/gen-tool-annotations.js
 */

const fs = require('fs');
const path = require('path');
const { AsyncLocalStorage } = require('async_hooks');

const ROOT = path.resolve(__dirname, '..', '..');
const OUT = path.join(ROOT, 'src', 'server', 'tool-annotations.generated.ts');

const { wpsClient } = require(path.join(ROOT, 'dist', 'client', 'wps-client.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));

/**
 * 观测上下文用 AsyncLocalStorage，不用全局变量。
 *
 * 为什么：一开始用的是全局 `recording`，结果生成器不幂等 ——
 * 同一个 wps_common_status，这次观测到 'status'、那次 'wpsStatus'。
 * 因为加了超时保护后，超时的 handler 【并没有真的停下】，它还在异步跑，
 * 全局 recording 早就被下一个工具换成别人了，它的回话就串到别人头上去了。
 * AsyncLocalStorage 跟着各自的调用链走，超时也不串。
 */
const als = new AsyncLocalStorage();

/** 动作名推不出来、注解手写的工具。名单真源在 handwritten-annotations.js，
 *  与 tool-annotations.ts 的 EXTRA_ANNOTATIONS 共用（下面有断言卡住同步）。 */
const { EXTRA_ANNOTATIONS_KEYS } = require(path.join(ROOT, 'scripts', 'manual', 'handwritten-annotations.js'));
const HANDWRITTEN = new Set(EXTRA_ANNOTATIONS_KEYS);

// ---- 1. 换掉真正的出口，装上记录器 ----------------------------------------
wpsClient.invokeAction = async (action, params, appKey) => {
  // 只记合法的动作名。wps_execute_method 的动作名是【调用时传入的参数】，
  // 空参探测时拿到的是 undefined —— 记进去就是脏数据，会误导后来人。
  const ctx = als.getStore();
  if (ctx) {
    if (typeof action === 'string' && /^[A-Za-z_][A-Za-z0-9_]*$/.test(action)) {
      ctx.actions.add(action);
    } else {
      ctx.unresolved = true; // 走到了出口，但动作名推不出来
    }
  }
  return { success: false, error: '__annotation_probe__' };
};
// openFile/saveFile 走 invokeAction，但如果将来改成直连，记得同步 patch

(async () => {
  // 按 server.start() 的真实顺序走一遍：
  //   1) registerBuiltinTools() —— mcp-server.ts 里手写的 12 个
  //   2) registerAll(allTools)  —— src/tools 下的 273 个
  // 少做第 1 步，注解表就会比线上少一截（273 vs 285）。这类"两处维护"的漂移
  // 已经吃过一次亏，所以这里不写死数字，让它跟真实注册路径跑。
  const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
  const inst = new WpsMcpServer();
  inst.registerBuiltinTools();
  toolRegistry.registerAll(allTools);
  const tools = toolRegistry.listAllTools();

  const map = {};
  /** [{name, unresolved}] —— 没记到合法动作名的工具 */
  const missed = [];

  for (const t of tools) {
    const name = t.name;
    const rec = { actions: new Set(), unresolved: false };
    try {
      const tool = toolRegistry.getTool(name);
      // 加超时：有些 handler 会做重循环或长 sleep，观测脚本不该被它们拖住。
      // 超时不算失败 —— 已经记到的动作名照样算数（ALS 保证不会串到别的工具）。
      await Promise.race([
        als.run(rec, () => tool.handler({})),
        new Promise((_, rej) => setTimeout(() => rej(new Error('__probe_timeout__')), 150)),
      ]);
    } catch (e) {
      // 参数校验就抛错的（没走到调用出口），走下面的名字推断兜底
    }
    if (HANDWRITTEN.has(name)) {
      // 编排类/逃生舱类：动作名不由它自己决定，映射记空，注解走手写表
      map[name] = [];
    } else if (rec.actions.size) {
      map[name] = [...rec.actions];
    } else {
      missed.push({ name, unresolved: rec.unresolved });
    }
    if ((tools.indexOf(t) + 1) % 50 === 0) {
      console.log(`  …${tools.indexOf(t) + 1}/${tools.length}`);
    }
  }

  // ---- 2. 没记到合法动作名的，分两种处理 --------------------------------------
  // wps_excel_get_cell_value -> getCellValue
  const infer = (toolName) => {
    const act = toolName
      .replace(/^wps_/, '')
      .replace(/^(word|excel|ppt|common)_/, '');
    return act.replace(/_([a-z0-9])/g, (_m, c) => c.toUpperCase());
  };
  // unresolved 保留作调试线索：这些工具到了出口但动作名是运行期参数
  const unresolved = [];
  const inferred = [];
  for (const { name, unresolved: u } of missed) {
    if (u) {
      // 走到了调用出口，但动作名是【运行期参数】，静态推不出来。
      // 这种必须走 tool-annotations.ts 里手写的 EXTRA_ANNOTATIONS，
      // 不能拿猜出来的名字判读写 —— 客户端可能拿这个注解做安全决策。
      map[name] = [];
      unresolved.push(name);
    } else {
      // 压根没走到出口（参数校验就抛错了），名字推断基本等价于真实动作名
      map[name] = [infer(name)];
      inferred.push(name);
    }
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
  lines.push(' *');
  lines.push(' * 空数组 = 动作名是运行期参数，静态推不出来（目前只有 wps_execute_method，');
  lines.push(' * 它能执行任意 WPS 动作）。这类工具的注解写在 tool-annotations.ts 的');
  lines.push(' * EXTRA_ANNOTANCES 里手写，别拿猜出来的名字去判读写。');
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
  console.log(`工具总数（真实注册路径）: ${tools.length}`);
  console.log(`  运行时观测到动作      : ${tools.length - missed.length - HANDWRITTEN.size}`);
  console.log(`  名字推断兜底          : ${inferred.length}${inferred.length ? ' → ' + inferred.slice(0, 6).join(', ') : ''}`);
  console.log(`  手写注解（不推断）    : ${HANDWRITTEN.size} → ${[...HANDWRITTEN].join(', ')}`);
  console.log(`一工具多动作    : ${multi.length}`);
  for (const [n, a] of multi.slice(0, 6)) console.log(`   ${n} → ${a.join(', ')}`);

  // 自检：注解表条目必须与真实注册表一致，否则就是漂移
  if (Object.keys(map).length !== tools.length) {
    console.error(`✗ 注解表 ${Object.keys(map).length} 条 ≠ 注册表 ${tools.length} 个工具`);
    process.exit(1);
  }
  // HANDWRITTEN 与 EXTRA_ANNOTATIONS 必须同步（名单真源已抽到
  // handwritten-annotations.js）。这里反向验：名单里的工具必须真能拿到注解 ——
  // 它们在映射表里被记成空数组，注解只能来自 tool-annotations.ts 的
  // EXTRA_ANNOTATIONS。忘了写就会退回"名字推断"，而推断对编排类是猜错的。
  const { getToolAnnotations } = require(path.join(ROOT, 'dist', 'server', 'tool-annotations.js'));
  const noHand = EXTRA_ANNOTATIONS_KEYS.filter((k) => !getToolAnnotations(k));
  if (noHand.length) {
    console.error(
      `✗ 名单里的工具拿不到手写注解：${noHand.join(', ')}\n  ` +
        '查 tool-annotations.ts 的 EXTRA_ANNOTATIONS 是否已写；忘了写就会静默退回名字推断。'
    );
    process.exit(1);
  }
  console.log(`\n已写入 ${path.relative(ROOT, OUT)}（条目 ${Object.keys(map).length} = 注册表 ${tools.length}）`);
  process.exit(0);
})().catch((e) => {
  console.error('生成失败:', e);
  process.exit(1);
});