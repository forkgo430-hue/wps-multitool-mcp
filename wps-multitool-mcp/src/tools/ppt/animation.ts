/**
 * Input: PPT 动画列表与动画删除操作参数
 * Output: 动画查询/删除结果
 * Pos: PPT 动画工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * PPT动画Tools - 动画查询与删除模块
 *
 * 只留两个不能合并的工具：
 * - wps_ppt_get_animations: 查询动画列表（只读，没东西可并入别的工具）
 * - wps_ppt_remove_animation: 删除单个动画（破坏性，混进"设置"工具会污染注解）
 *
 * 第 8 轮合并掉了本文件原先的 7 个工具，它们都是别名，能力没丢：
 * - wps_ppt_add_animation / add_animation_preset / add_emphasis_animation / set_animation_order
 *   → wps_ppt_set_animation（op: 'add' | 'reorder'）
 * - wps_ppt_set_slide_transition / apply_transition_to_all / remove_slide_transition
 *   → wps_ppt_set_transition（target: 'slide' | 'all'，effect: 'none' 即取消切换）
 * 旧名走 scripts/manual/tool-aliases.json 的别名表调用，参数映射也在那张表里。
 */

import { v4 as uuidv4 } from 'uuid';
import {
  ToolDefinition,
  ToolHandler,
  ToolCallResult,
  ToolCategory,
  RegisteredTool,
} from '../../types/tools';
import { wpsClient } from '../../client/wps-client';
import { WpsAppType } from '../../types/wps';

const TRIGGER_NAME_MAP: Record<string, string> = {
  onClick: '单击时',
  withPrevious: '与上一动画同时',
  afterPrevious: '上一动画之后',
};

// ==================== 1. 移除动画效果 ====================

/**
 * 移除动画效果
 * 从指定幻灯片中移除指定的动画
 */
export const removeAnimationDefinition: ToolDefinition = {
  name: 'wps_ppt_remove_animation',
  description: `移除幻灯片中指定的动画效果。

使用场景：
- "删除第1页的第2个动画"
- "移除这个动画效果"

先想清楚删的是哪一个：wps_ppt_get_animations 能列出序号，确认序号再删。
序号是本页动画窗格里的顺序（1 开始），加完动画会变。`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      animationIndex: {
        type: 'number',
        description: '动画索引（从1开始）',
      },
    },
    required: ['slideIndex', 'animationIndex'],
  },
};

export const removeAnimationHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, animationIndex } = args as {
    slideIndex: number;
    animationIndex: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'removeAnimation',
      { slideIndex, animationIndex },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `动画移除成功！\n幻灯片: 第 ${slideIndex} 页\n已移除第 ${animationIndex} 个动画`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `移除动画失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `移除动画出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ==================== 2. 获取动画列表 ====================

/**
 * 获取动画列表
 * 查看指定幻灯片上所有动画效果
 */
export const getAnimationsDefinition: ToolDefinition = {
  name: 'wps_ppt_get_animations',
  description: `获取幻灯片上所有动画效果的列表。

使用场景：
- "查看第1页有哪些动画"
- "列出这页的所有动画"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
    },
    required: ['slideIndex'],
  },
};

export const getAnimationsHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex } = args as { slideIndex: number };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      animations: Array<{
        index: number;
        effect: string;
        trigger: string;
        shapeName: string;
      }>;
    }>(
      'getAnimations',
      { slideIndex },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      const animations = response.data.animations;
      if (!animations || animations.length === 0) {
        return {
          id: uuidv4(),
          success: true,
          content: [
            {
              type: 'text',
              text: `第 ${slideIndex} 页没有动画效果。`,
            },
          ],
        };
      }

      let output = `第 ${slideIndex} 页的动画效果（共 ${animations.length} 个）：\n`;
      animations.forEach((anim) => {
        output += `  ${anim.index}. [${anim.shapeName}] ${anim.effect} - ${TRIGGER_NAME_MAP[anim.trigger] || anim.trigger}\n`;
      });
      output += `要调顺序用 wps_ppt_set_animation(op='reorder')，要删某一条用 wps_ppt_remove_animation。`;

      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: output }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `获取动画列表失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `获取动画列表出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 导出动画查询/删除相关的Tools
 */
export const animationTools: RegisteredTool[] = [
  { definition: removeAnimationDefinition, handler: removeAnimationHandler },
  { definition: getAnimationsDefinition, handler: getAnimationsHandler },
];

export default animationTools;