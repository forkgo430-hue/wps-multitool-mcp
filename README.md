# WPS Office 多面手 MCP

AI编程助手通过自然语言操控WPS Office的MCP工具集。

> **WPS Office 多面手 MCP**（wps-multitool-mcp）
> 当前维护者：**赢容**。完整变更见 [CHANGELOG.md](CHANGELOG.md)。

## 项目定位

本项目是MCP Server + Skills框架，让AI助手（Claude Code/Cursor/Augment等）能操控WPS Office。

<!-- AUTO:tool-stats:start -->
- **306 个工具**（12 个内置 + 294 个专业/组工具）；**默认只列 40 个**（基础档，按需加载 —— 口径与校验见 `scripts/manual/basic-tier-verify.js`）
- 分布（按工具组，内置计入对应组）：通用 25（含 core）/ Word 70 / Excel 95 / PPT 116
- 注解（实测）：只读 53 / 破坏性 34 / 非幂等 115，覆盖 306/306
<!-- AUTO:tool-stats:end -->
- 每个工具都带 MCP 注解（`readOnlyHint` / `destructiveHint` / `idempotentHint` / `openWorldHint`），
  客户端从 `tools/list` 直接读到读写与破坏性，不必靠描述猜
- 配套 skill `wps-multitool`（`wps-office-mcp/skills/`）：把「探活→定组→加载组→
  选型读反例→写前确认」固化成默认链路，安装到 `~/.workbuddy/skills/`
- 仅支持 Windows（WPS 的 COM 自动化只在 Windows 可用）
- 走**常驻桥进程**：冷启动约 5 秒，之后每次调用 2~3ms（一应用一进程，互不阻塞）

## 前提条件

- 已安装 WPS Office（https://www.wps.cn/ 或 https://www.wps.com/）
- Node.js >= 18.0.0
- Git

## 自动安装（AI执行）

以下步骤由AI助手自动执行。需要人类交互的步骤标注 ⚠️。

### 步骤1: 克隆项目

```bash
git clone https://github.com/forkgo430/wps-multitool-mcp.git
cd wps-multitool-mcp
```

如果项目已存在，跳过此步。

### 步骤2: 安装依赖并编译

```bash
cd wps-office-mcp
npm install
npm run build
cd ..
```

### 步骤3: 配置MCP Server

根据使用的AI工具，将以下配置写入对应文件。注意将路径替换为实际的项目绝对路径。

**Claude Code** — 写入 `~/.claude/settings.json`：
```json
{
  "mcpServers": {
    "wps-office": {
      "command": "node",
      "args": ["/你的路径/wps-multitool-mcp/wps-office-mcp/dist/index.js"]
    }
  }
}
```

**Cursor** — 写入项目根目录 `.cursor/mcp.json`：
```json
{
  "mcpServers": {
    "wps-office": {
      "command": "node",
      "args": ["/你的路径/wps-multitool-mcp/wps-office-mcp/dist/index.js"]
    }
  }
}
```

**OpenAI Codex CLI** — 写入 `~/.codex/config.toml`：
```toml
[mcp_servers.wps-office]
command = "node"
args = ["/你的路径/wps-multitool-mcp/wps-office-mcp/dist/index.js"]
```
或命令行注册：`codex mcp add wps-office -- node /你的路径/wps-multitool-mcp/wps-office-mcp/dist/index.js`

**Augment / 其他MCP兼容IDE** — 参考各IDE的MCP Server配置文档，使用相同的command和args。本项目 MCP Server 为标准 stdio 实现（spec 2025-11-25），与所有 MCP 一等客户端（Claude Code / Cursor / Codex CLI / GitHub Copilot CLI / Windsurf 等）兼容。

### 步骤4: 安装WPS加载项

⚠️ 需要人工操作（AI无法直接操作WPS应用）：

```powershell
powershell -ExecutionPolicy Bypass -File scripts/install.ps1
```

⚠️ 安装后必须重启WPS Office才能生效。

### 步骤5: 安装Skills（仅Claude Code需要）

```bash
# 创建skills目录（如不存在）
mkdir -p ~/.claude/skills

# 创建符号链接
ln -sf "$(pwd)/wps-office-mcp/skills/wps-multitool" ~/.claude/skills/wps-multitool
```

### 步骤6: 验证安装

```powershell
cd wps-office-mcp
npm run verify
# 应看到 14 项验证门禁全绿（含 Word/Excel 真机能力探测）
```

## 架构

```
Skills 层：SKILL.md（探活 → 定组 → 加载组 → 选型 → 写前确认）
  ↓ MCP 协议（stdio）
MCP Server（默认只列 40 个基础工具，其余按需加载）
  ↓ wpsClient.executeMethod() —— 命名管道
常驻桥（Node：dist/bridge/index.js；按应用拆子进程，互不阻塞）
  ↓ 每个应用一个常驻 PowerShell（excel / word / ppt / common）
wps-comd.ps1（stdin JSON 分发）→ wps-com.ps1（COM 动作）
  ↓
Ket / Kwps / Kwpp.Application（WPS 表格 / 文字 / 演示）
```

## 工具清单

计数由 `scripts/manual/gen-doc-stats.js` 从注册表生成，勿手改（CI 有 diff 门禁）。
口径与上方统计区一致：内置工具计入对应应用组（core 计入通用），故“内置”行不叠加：

<!-- AUTO:tool-table:start -->
| 应用 | 工具数 | 主要能力 |
|------|--------|---------|
| Excel | 95 | 公式/数据/图表/透视表/工作表/格式/工作簿/行列/批注保护/图片导出 |
| Word | 70 | 格式/内容/文档管理/页眉页脚/批注/书签/表格/模板填写/学术排版/修订/索引/交叉引用 |
| PPT | 116 | 幻灯片/形状/图片/表格/美化/动画/图表/3D/数据可视化/图片导出 |
| 通用 | 25 | 保存/连接检测/文本选取/格式转换/工具组管理/意图路由/批量执行 |
| 内置 | 12（已计入上方各组） | 连接检查/万能方法调用/数据缓存 |
<!-- AUTO:tool-table:end -->

## 故障排除

| 问题 | 解决方案 |
|------|---------|
| MCP连接失败 | 确认 `npm install && npm run build` 已执行，检查dist/index.js存在 |
| WPS未响应 | 重启WPS Office，确认加载项已安装 |
| "arguments error" | 重新运行安装脚本，重启WPS |
| Linux找不到插件 | 本项目仅支持 Windows（WPS 的 COM 自动化只在 Windows 可用） |
| 工具调用返回null | 确认WPS中已打开对应类型的文档 |
| 其他问题 | 到 GitHub 仓库提 issue：https://github.com/forkgo430/wps-multitool-mcp/issues |

## 许可证

MIT
