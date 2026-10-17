# 贡献指南

## 环境

- Windows 10/11 + WPS Office（真机验证绕不开，很多接口只有 WPS 见过世面）
- Node.js ≥ 18
- 首次：`cd wps-multitool-mcp && npm ci && npm run build`

## 常用命令

```powershell
cd wps-multitool-mcp
npm run build          # TypeScript 构建
npm run verify         # 快测门禁（9 项，秒级）
npm run verify:all     # 全量门禁（14 项，含真机探测，约 1 分钟）
npm run gen:annotations  # 重新生成「工具 → PS 动作」映射表
node scripts/manual/gen-doc-stats.js   # 重新生成 README / 面板里的统计数字
```

## 三条硬规矩

1. **生成物不许手改**
   - `src/server/tool-annotations.generated.ts`（`npm run gen:annotations` 产出）
   - `src/client/action-routes.generated.ts`（`node scripts/gen-action-routes.js` 产出）
   - README / README.en.md 的 `<!-- AUTO:… -->` 区间、panel.html 的数字（`gen-doc-stats.js` 产出）
   CI 会跑生成器再 `git diff --exit-code`，手改过或者忘了重跑都会红。

2. **统计数字只许来自注册表**
   总数、按组分布、注解分布，一律走真实注册路径（`registerBuiltinTools()` + `registerAll(allTools)`）。
   少做第一步就少算 12 个内置工具——这种"两处维护"的漂移吃过亏。

3. **PS 动作改动要重跑路由表**
   改完 `scripts/wps-com.ps1` 就跑 `node scripts/gen-action-routes.js`（它按 case 体推断动作属于哪个应用）。
   同一个动作名出现两个 case 时 PowerShell 只认第一个，后者是死代码——加动作前先 `findstr` 确认没重名。

## 新增一个工具的流程

1. 在 `src/tools/<app>/` 里加 definition + handler（参考 `word/gaps.ts` 与 `excel/sheet-gaps.ts` 的写法）
2. 如果是新 PS 动作：在 `scripts/wps-com.ps1` 加 case，跑 `node scripts/gen-action-routes.js`
3. 在 `src/server/tool-metadata.ts` 的 `CURATED` 补别名与 `notFor`（近名工具必写反例）
4. 需要真机验证的接口：加进对应 `__probeWordCaps` / `__probeExcelCaps` 的探测项
5. 跑 `npm run gen:annotations` + `gen-doc-stats.js`，再 `npm run verify:all`
6. 提交

## 提交信息

沿用 `type(scope): 摘要`，例如：

```
feat(word): 补 get_tables / get_table_data
fix(ci): 修正文档统计门禁的 pathspec
docs: 双语 README 门面化
```

**说清"为什么"比说清"改了什么"更重要**——这个仓库的注释和提交信息都以"踩过的坑"为主，
便于下一个人不重踩。

## 关于 LICENSE

根目录 MIT 许可与版权声明必须保留（协议要求）。文档与代码里不再引用其它项目的署名，
但许可文本本身不可删除或改写。
