import type { DeckRequest, GenerationStyle, SlideContent, TemplateProfile } from "@/types/ppt";

const styleGuidance: Record<GenerationStyle, string> = {
  professional: "专业简洁：结论先行，正文短句化，视觉保持克制但每页都有图形证据。",
  minimal: "极简高效：减少装饰，使用大留白、单一焦点和明确层级，不做纯文字页。",
  visual: "视觉叙事：优先用图片、流程、对比和大数字承载信息，文字只解释关键判断。",
  academic: "学术严谨：用研究问题、方法、证据、结论组织页面，图表和引用说明要清楚。",
  executive: "高管汇报：先给判断和影响，再给证据、风险、资源与行动建议。",
  storytelling: "故事化表达：每页承担一个叙事动作，使用场景、冲突、转折和下一步推动。",
};

export const PPT_SKILL_DESIGN_BRIEF = [
  "参考优秀 PPT 制作 skill 的要求：",
  "- 每一页都必须有视觉元素，不能只给标题和项目符号；视觉元素可以是图表、指标卡、流程、对比矩阵、真实场景图建议或图标化结构。",
  "- 调色应有主次关系：一个主色占 60%-70%，1-2 个辅助色，一个强调色；不要默认使用泛蓝色，也不要让所有颜色权重相同。",
  "- 使用“深色封面/结束页 + 浅色内容页”的 sandwich 结构，除非模板本身明确要求全深色。",
  "- 正文左对齐，少用居中正文；每页只表达一个主判断，正文控制在 2-4 条，每条尽量不超过 24 个中文字符。",
  "- 页面建议必须具体到可执行版式，例如指标卡、双栏证据、时间线、2x2 对比、流程泳道、半屏图片，而不是泛泛写“美观排版”。",
  "- 图表建议必须说明图表类型、比较维度和读者应看到的结论，不输出占位符或 lorem ipsum。",
  "- 避免模板味很重的标题下划线，使用留白、色块、侧边栏、编号圆点或图形节奏建立统一视觉母题。",
  "- 生成内容时同时考虑 QA：不让文字过长、避免低对比、避免元素拥挤，页边距至少 0.5 英寸。",
].join("\n");

export function buildDeckCraftBrief(request: DeckRequest) {
  return [
    PPT_SKILL_DESIGN_BRIEF,
    `当前风格策略：${styleGuidance[request.style]}`,
    `模板适配：沿用 ${request.templateProfile?.name ?? request.templateId} 的主色、字体、封面/目录/结束页结构和已检测版式。`,
  ].join("\n");
}

function compactText(value: string, max = 42) {
  const cleaned = value.replace(/\s+/g, " ").trim();
  return cleaned.length > max ? `${cleaned.slice(0, max - 1)}…` : cleaned;
}

function ensureVisualSuggestion(slide: SlideContent) {
  if (slide.templateRole === "agenda") {
    return {
      chartSuggestion: slide.chartSuggestion || "使用纵向时间线展示本次汇报路径，每个节点对应一个章节判断。",
      imageSuggestion: slide.imageSuggestion || "使用小图标或编号圆点强化章节节奏，不使用抽象装饰图。",
      layoutSuggestion: slide.layoutSuggestion || "左侧目录节点，右侧保留主题视觉色块和一条进度轨道。",
    };
  }

  if (slide.templateRole === "data") {
    return {
      chartSuggestion: slide.chartSuggestion || "使用 3 个关键指标卡 + 横向条形图，突出最高影响项和变化方向。",
      imageSuggestion: slide.imageSuggestion || "使用数据仪表盘、产品界面或业务场景图作为真实语境。",
      layoutSuggestion: slide.layoutSuggestion || "上方结论标题，下方左侧指标卡，右侧条形图和一句洞察。",
    };
  }

  if (slide.templateRole === "closing") {
    return {
      chartSuggestion: slide.chartSuggestion || "无需复杂图表，使用下一步行动清单和优先级标记。",
      imageSuggestion: slide.imageSuggestion || "使用深色背景上的高对比行动卡片或路线视觉。",
      layoutSuggestion: slide.layoutSuggestion || "左侧总结判断，右侧 3 个下一步行动卡片。",
    };
  }

  if (slide.templateRole === "cover") {
    return {
      chartSuggestion: slide.chartSuggestion || "无需图表，使用强主题视觉和 1-2 个关键信息标签。",
      imageSuggestion: slide.imageSuggestion || "使用与主题直接相关的真实场景图、产品图或抽象度低的图形母题。",
      layoutSuggestion: slide.layoutSuggestion || "大标题 + 副标题 + 右侧视觉焦点，深色背景承载品牌感。",
    };
  }

  return {
    chartSuggestion: slide.chartSuggestion || "使用小型对比图或流程图支撑本页结论，标出最重要的差异或路径。",
    imageSuggestion: slide.imageSuggestion || "使用与本页观点直接相关的真实图像、界面截图或图标组。",
    layoutSuggestion: slide.layoutSuggestion || "左侧结论与要点，右侧证据卡片、流程图或图像占位。",
  };
}

export function enforcePptSkillOnSlide(
  slide: SlideContent,
  index: number,
  total: number,
  template: TemplateProfile,
): SlideContent {
  const isCover = index === 0;
  const isClosing = index === total - 1;
  const role = isCover ? "cover" : isClosing ? "closing" : slide.templateRole;
  const body = (slide.body.length ? slide.body : ["明确核心判断", "补充关键证据", "给出下一步动作"])
    .slice(0, role === "cover" ? 2 : 4)
    .map((point) => compactText(point, role === "cover" ? 34 : 46));
  const suggestions = ensureVisualSuggestion({ ...slide, templateRole: role });

  return {
    ...slide,
    title: compactText(slide.title || `第 ${index + 1} 页`, role === "cover" ? 38 : 34),
    body,
    colorPalette: slide.colorPalette?.length ? slide.colorPalette : template.dominantColors,
    templateRole: role,
    chartSuggestion: compactText(suggestions.chartSuggestion, 96),
    imageSuggestion: compactText(suggestions.imageSuggestion, 88),
    layoutSuggestion: compactText(suggestions.layoutSuggestion || template.fallbackStrategy, 88),
    speakerNotes:
      slide.speakerNotes ||
      `本页先讲清「${slide.title}」的结论，再用图形区域补充证据，最后自然过渡到下一页。`,
  };
}
