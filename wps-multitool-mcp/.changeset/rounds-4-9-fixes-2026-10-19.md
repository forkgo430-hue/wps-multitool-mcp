---
"wps-multitool-mcp": patch
---

第 4~9 轮的真 bug 修复（10 个）。这批此前一个字都没进 changeset，本次补记。

## 四个"从来就没成功过"的工具

以下四个工具**在此之前每一次调用都没有真正生效**，但都返回 `success`。如果你以前用过它们，
看到的结果不是"参数写错了"，而是"这功能实际上一直没工作"。

- **`wps_ppt_insert_slide_image`** —— 它调用的动作名是 `insertImage`，而该动作属于 **Word** 段
  （`Get-WpsWord` + `InlineShapes.AddPicture`）。PPT 工具去调它必然拿不到文档对象，
  于是每次调用都报"不能对 Null 值表达式调用方法"。现在改走 PPT 段的 `insertPptImage`。
- **`wps_ppt_create_kpi_cards`** —— 回 `success`，但**一张卡片都没建**。
  PowerShell 把单元素数组在管道里摊平，`$cards.Count` 变成空值，`for` 循环一次都不跑。
- **`wps_ppt_create_mini_charts`** —— 同一个数组摊平问题导致 0 个图表；
  且多项时 `data[].value` 从 JSON 过来是 `Int32`，直接赋给 `TextRange.Text` 抛"指定的转换无效"。
- **`wps_ppt_align_objects`** —— 回"对象对齐完成"，实际把**整页所有形状**一起左对齐
  （同页的表格、图片、占位符都被搬走）。原因是没给形状清单时，`Shapes.Range($null)`
  在 WPS 上不报错，含义是"整页"。现在 `shapeIndices` 改为必填，底层对空清单**响亮报错**而不是替用户选。

## 四个"读回形同虚设"的工具

以下工具返回成功，但回包里的关键字段恒为 `undefined`，等于读回功能不存在：

- `wps_ppt_get_textboxes` —— 读 `data.textboxes`，底层回的是 `textBoxes`（大写 B），
  永远打印 `第N页文本框: []`
- `wps_ppt_get_animations` —— 按 `effect` / `trigger`（字符串）读，底层回的是
  `effectType`（MsoAnimEffect 枚举数字）与 `duration`，列表一直是 `1. [矩形 1] undefined - undefined`
- `wps_ppt_open_presentation` —— 读 `data.filePath`，底层回的是 `path`，永远打印 `路径: undefined`
- `wps_ppt_create_kpi_cards` / `wps_ppt_create_mini_charts` —— 读 `cardCount` / `chartCount`，
  底层回的是 `count`，永远打印"数量: undefined 个"

## 两个占位符编号错的

- `wps_ppt_set_slide_subtitle` 与 `wps_ppt_set_slide_content` **必然**报
  `placeholder not found`。原因是占位符编号是查表抄的而不是实测的：
  实测 WPS 的 `layout='title'` 页里副标题是 `PlaceholderFormat.Type=4`（居中标题=3），
  `layout='title_content'` 页里正文是 **2**；而代码里副标题找 2、正文找 7，两个都不对。
  现在按实测编号匹配，报错信息也改成说明"哪种版式能写"。

## 一个加一次生成两条的

- `addAnimation` 加**一个**动画实际生成**两条**一模一样的效果，删一条还剩一条，看起来像删不掉。
  原因是 `AddEffect(shape, id, level)` 的第三个参数是 **Level（文本级别）**而不是触发时机，
  对没有文本的形状必须传 `msoAnimateLevelNone(0)`。

## 最狠的一条：切换效果被静默改写

- `applyTransitionToAll` 读的是 `$p.transition`，而工具一直传的是 `effect` →
  **整份演示文稿的切换效果被静默设成 fade**。你要 push，得到的是 fade。
  参数名不匹配不会报错，门禁和用户都没发现。同批还修了 `setSlideTransition`
  （把效果名字符串直接赋给 `PpEntryEffect` 枚举会抛错、且不认 `none`，导致"取消切换"必炸）。

## 表现最唬人的一条：像 69 个工具一起坏了

- `addAnimation` 曾用 `Shapes.Item($p.shapeName)` 取形状，而工具传的是 `shapeIndex` →
  拿到空对象后调 `AddEffect`，于是该页的 `Shapes` / `Background` 全部报 Null，
  而文稿列表仍然正常。**症状是"大量 PPT 工具同时失效"，根因只有一个。**
  现在同时认 `shapeName` 与 `shapeIndex`。

## 排版类

- `auto_layout` / `smart_distribute` 的形状收集与赋值改为逐个 `try/catch`：
  页面里有 3D / 表格 / 图片时读 `Width` / `Height` 会抛"指定的转换无效"，此前会让整页排版失败；
  现在跳过读不到尺寸的形状，并在结果里回报 `skipped` 计数。

## 测试口径同步修正

- PPT 冒烟从顺序执行改为**每条重建全新靶子**（`ppt-isolated-verify.js`）。顺序执行时前一条留下的
  状态会让"导出图无变化"分不清是工具坏了还是靶子已经变了 —— 实测过同一批工具顺序跑大面积失败、
  单跑全绿。
- 21 条"需外部素材 / 前置不足"的用例改判据后救回，其中 `remove_animation`、`set_animation_order`
  这类所谓"靶子没有可操作对象"的，本质是**用例设计没跟上**，不是工具问题。
- 真机覆盖矩阵 100 行 → 103 行，执行 74 条 → **98 条**，全部通过；跳过项由 21 条收敛到 **5 条**
  （5 条跳过各有实测原因，不是漏跑）。