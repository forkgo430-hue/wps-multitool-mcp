# WPS Office 多面手 MCP

**简体中文** | [English](README.en.md)

<!-- AUTO:badges:start -->
[![CI](https://github.com/forkgo430/wps-multitool-mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/forkgo430/wps-multitool-mcp/actions/workflows/ci.yml) [![license](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE) [![npm](https://img.shields.io/npm/v/wps-multitool-mcp)](https://www.npmjs.com/package/wps-multitool-mcp) [![release](https://img.shields.io/github/v/release/forkgo430/wps-multitool-mcp)](https://github.com/forkgo430/wps-multitool-mcp/releases) [![tools](https://img.shields.io/badge/tools-296-brightgreen.svg)](#能力一览) [![platform](https://img.shields.io/badge/platform-Windows%20only-lightgrey.svg)](#安装)
<!-- AUTO:badges:end -->

> 让 AI 直接操控你**正开着**的 WPS 文字 / 表格 / 演示。

<sub>作者 / 维护者：**赢容** · [npm](https://www.npmjs.com/package/wps-multitool-mcp) · [Release](https://github.com/forkgo430/wps-multitool-mcp/releases)</sub>

## 30 秒上手

```powershell
npx -y wps-multitool-mcp
```

```json
{
  "mcpServers": {
    "wps-multitool-mcp": {
      "command": "npx",
      "args": ["-y", "wps-multitool-mcp"]
    }
  }
}
```

重启客户端，打开一个 WPS 文档，然后直接说：

> 把这份合同的金额段落加粗，并在末尾补一段风险提示

想要侧边面板（快照 / 撤销 / 状态），再装加载项：

```powershell
irm https://cdn.jsdelivr.net/gh/forkgo430/wps-multitool-mcp@main/scripts/install.ps1 | iex
```

## 为什么不是"又一个 WPS 插件"

| 坑 | 机制 |
|---|---|
| 几百个工具塞爆上下文 | 默认只列基础档，其余 `wps_search_tools` 搜、`wps_load_tool_group` 拉 |
| 写到一半失败，说不清到底改没改 | `warnings` 第三态：做完了，但有 N 处没按预期 |
| 近名工具选错，不报错，改到别处 | 每个工具带 `notFor` 反例，搜索命中必读 |
| 改坏了退不回去 | 写入前自动快照（`wps_common_batch` 内置）+ 操作日志，可回滚到磁盘那一版 |
| 不知道哪些工具真能用 | 每个工具标 `真机验证 / 仅静态验证`，默认档只收验证过的 |

COM 不是事务，所以**失败带重试标签**：`[可重试]` 直接重发，`[勿重试·可能已执行一半]` 先读回来再决定。

## 适用对象

**该用**：Windows 上用 WPS 干活、并且想直接改**已经开着的**文档的人。

- WPS 文字/表格/演示是主力工具（公文、合同、学术论文、汇报材料、数据报表）
- 要动的是"眼前这份"——带着批注、修订、页眉页脚、还没保存的改动
- 需要中文场景的细节：学术排版、修订留痕、模板填写、目录与交叉引用
- 接受"AI 会犯错"这个前提，因此愿意让它先快照、再动手

**不该用**：下面这些场景它帮不上忙，换工具更省事。

- macOS / Linux —— 这条路走的是 Windows COM，没有移植计划
- 无人值守的 CI / 服务器 —— 它需要交互式桌面和一个开着 WPS 的登录会话
- 成千上万个文件的批处理 —— 那是文件解析类库的活（`python-docx` / `openpyxl`），不是 COM 的强项
- 只想读文档内容做检索 —— 不需要开 Office 的场景，用解析库更快更稳

**只做 WPS，不做 MS Office** —— 这是取舍，不是遗漏。WPS 与 Office 是两套 COM 对象模型（`Ket/Kwps/Kwpp` vs `Excel/Word/PowerPoint`），双栈意味着每个动作都要写两遍、在两种环境各验一遍，维护成本翻倍而收益只覆盖一部分用户。这套工具的细节（中文公文格式、修订、学术排版）本来就是冲着 WPS 用户去的 —— 与其两头都半桶水，不如一头做到底。用 MS Office 的话，社区里有专门的 Excel/PPT MCP，各取所需。

## 规模

<!-- AUTO:tool-stats:start -->
- **296 个工具**（8 个内置 + 288 个专业/组工具）；**默认只列 40 个**（基础档 ∩ 真机验证，按需加载 —— 口径与校验见 `scripts/manual/basic-tier-verify.js`）
- 分布（按工具组，内置计入对应组）：通用 27（含 core）/ Word 67 / Excel 89 / PPT 113
- 注解（实测）：只读 51 / 破坏性 34 / 非幂等 110，覆盖 296/296
- 真机验证 201/296（Word/Excel 为组级真机探测；PPT 仅静态校验 —— 明细见 [TOOLS.md](TOOLS.md)）
<!-- AUTO:tool-stats:end -->

## 架构

```mermaid
flowchart TB
    S["Skills 层<br/>探活 → 定组 → 加载组 → 选型 → 写前确认"]
    M["MCP Server<br/>stdio，默认 40 个基础工具"]
    B["常驻桥（Node）<br/>按应用拆子进程"]
    P["常驻 PowerShell<br/>excel / word / ppt / common"]
    A["wps-comd.ps1<br/>stdin JSON 分发"]
    C["wps-com.ps1<br/>COM 动作"]
    W["Ket / Kwps / Kwpp.Application"]

    S -->|MCP 协议| M
    M -->|"executeMethod()<br/>命名管道"| B
    B --> P
    P --> A
    A --> C
    C --> W
```

冷启动约 5 秒，之后每次调用 2~3ms。一应用一进程，WPS 文字卡死不会堵住表格。

## 能力一览

计数由 `scripts/manual/gen-doc-stats.js` 从注册表生成，勿手改（CI 有 diff 门禁）。
口径与上方统计区一致：内置工具计入对应应用组（core 计入通用），故“内置”行不叠加：

<!-- AUTO:tool-table:start -->
| 应用 | 工具数 | 主要能力 |
|------|--------|---------|
| Excel | 89 | 公式/数据/图表/透视表/工作表/格式/工作簿/行列/批注保护/图片导出 |
| Word | 67 | 格式/内容/文档管理/页眉页脚/批注/书签/表格/模板填写/学术排版/修订/索引/交叉引用 |
| PPT | 113 | 幻灯片/形状/图片/表格/美化/动画/图表/3D/数据可视化/图片导出 |
| 通用 | 27 | 保存/连接检测/文本选取/格式转换/工具组管理/意图路由/批量执行 |
| 内置 | 8（已计入上方各组） | 连接检查/万能方法调用/数据缓存 |
<!-- AUTO:tool-table:end -->

全部 306 个工具的完整清单（名称 / 默认档 / 注解 / 真机验证状态）见 **[TOOLS.md](TOOLS.md)**。它由注册表生成，与代码永远同步。

## 安装

<details>
<summary><b>环境要求</b></summary>

| 要求 | 说明 |
|------|------|
| Windows 10 / 11 | WPS 的 COM 自动化只在 Windows 可用 |
| WPS Office 2019+（实测 12.1） | 装哪个盘都行 |
| Node.js ≥ 18 | 构建与运行 MCP Server |
| 一个 MCP 客户端 | Claude Code / Cursor / Codex CLI / CodeBuddy / WorkBuddy 等 |

</details>

<details>
<summary><b>完整安装（一键脚本 + 加载项 + skill）</b></summary>

```powershell
irm https://cdn.jsdelivr.net/gh/forkgo430/wps-multitool-mcp@main/scripts/install.ps1 | iex
```

自动完成：检查环境 → 拉取项目 → 装依赖并编译 → 部署 WPS 加载项 → 安装 skill → 打印 MCP 配置。
装完重启 WPS 和客户端，把配置粘进去即可。需要能访问 GitHub；否则按下面手动装。

</details>

<details>
<summary><b>手动安装</b></summary>

```bash
git clone https://github.com/forkgo430/wps-multitool-mcp.git
cd wps-multitool-mcp
powershell -ExecutionPolicy Bypass -File scripts/install.ps1
```

脚本支持 `-DryRun`（只检查不改动）、`-Dev`（加载项用符号链接，改源码即时生效）、`-InstallDir`（指定目录）。

</details>

<details>
<summary><b>只配 MCP，不装加载项</b></summary>

不用手抄路径，让程序自己找：

```powershell
node wps-multitool-mcp/dist/index.js install
```

它会自动识别 CodeBuddy / WorkBuddy / Codex / Claude Code / Cursor（只认已经装过的，不会给你新建目录），
写前备份原文件、写后回读校验，失败自动还原。常用的两个开关：

```powershell
node dist/index.js install --dry-run    # 只看会改什么
node dist/index.js install --print      # 只打印配置片段，一个文件都不动
```

没识别到的客户端，它会打印可直接粘贴的片段。标准 stdio 实现（spec 2025-11-25），任何 MCP 客户端都能接。

</details>

## 隐私

- **文档内容不出本机**：没有遥测、没有崩溃上报、没有使用统计，`src/` 里没有任何出站请求
- 唯一的本地网络端点是桥的只读状态接口（`127.0.0.1:27731/status`，只给 WPS 任务窗格轮询）
- 快照按文档留在 `~/.wps-multitool-mcp/snapshots/`，等同文档副本，属于敏感目录，处理完敏感文档建议清空

完整威胁模型与建议见 [SECURITY.md](SECURITY.md)。

## 故障排除

| 问题 | 解决方案 |
|------|---------|
| MCP 连接失败 | 确认 `dist/index.js` 存在，客户端路径正确 |
| WPS 未响应 | 重启 WPS，确认 `%APPDATA%\kingsoft\wps\jsaddons\wps-claude-addon_` 存在 |
| "arguments error" | 重跑 `scripts/install.ps1`，重启 WPS |
| 一键安装卡住 | 访问不了 GitHub，改用「手动安装」 |
| 其他问题 | [提 issue](https://github.com/forkgo430/wps-multitool-mcp/issues) |

## 许可

MIT
