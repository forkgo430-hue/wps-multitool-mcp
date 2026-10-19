---
"wps-multitool-mcp": patch
---

PPT 工具族合并（见 `ppt-merge-2026-10-18.md`）的两个副作用点名。

那条 changeset 的定级是 `minor` 且**保持不动**（破坏面窄、且旧参数名已被参数映射兜住）。
但副作用必须在这里写明，否则升级的人看不到。

## 副作用一：7 个工具名不再出现在 `tools/list`

PPT 动画族与切换族合并后，这 7 个工具名降为**别名**：

```
wps_ppt_add_animation        wps_ppt_add_animation_preset
wps_ppt_add_emphasis_animation  wps_ppt_set_animation_order
wps_ppt_set_slide_transition wps_ppt_apply_transition_to_all
wps_ppt_remove_slide_transition
```

- **按名字直接调用仍然可用** —— 别名在入口处解析到规范名，旧名 + 旧参数都能跑。
- **它们不再出现在 `tools/list` 里。** 任何按 `tools/list` 枚举、或按清单硬编码工具名的脚本，
  升级后会"找不到"这 7 个。
- 工具总数因此从 291 → 284。

> 如果你的代码枚举 `tools/list` 来生成 UI 或做校验，请把工具总数与工具名清单改成动态获取，
> 别硬编码。

## 副作用二：`set_transition` 的 `sound` 参数被移除

- WPS 的 COM 接口**根本没有"切换音效"这个能力**（那是 PowerPoint 的 `AddMediaObject2`）。
- 该参数此前**永远不生效** —— 传了它和没传一样。
- 现在从 schema 里删掉。留着它只会让人以为这个功能存在。
- 传过 `sound` 的调用不会报错，但该参数不再有意义。

## 一句话

以上两点都不是"功能变少了"，是**把两个一直存在的假象拿掉**：一个假清单、一个假参数。