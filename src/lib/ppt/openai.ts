import type { DeckRequest, OutlineSection, SlideContent } from "@/types/ppt";
import { buildDeckCraftBrief } from "./pptSkillGuidelines";

const DEFAULT_OPENAI_BASE_URL = "https://api.openai.com";
export const DEFAULT_MODEL = "gpt-5.5";

type JsonSchema = Record<string, unknown>;

export class MissingApiKeyError extends Error {
  constructor() {
    super("服务器未配置 OPENAI_API_KEY。请在部署平台的 Secrets / Environment Variables 中配置。");
    this.name = "MissingApiKeyError";
  }
}

export class OpenAIRequestError extends Error {
  status: number;

  constructor(status: number, message: string) {
    super(message);
    this.name = "OpenAIRequestError";
    this.status = status;
  }
}

function sanitizeUpstreamError(details: string) {
  const redacted = details.replace(/sk-[A-Za-z0-9_-]+/g, "sk-***");

  try {
    const parsed = JSON.parse(redacted) as { error?: { message?: string; code?: string; type?: string } };
    const code = parsed.error?.code ? ` code=${parsed.error.code}` : "";
    const type = parsed.error?.type ? ` type=${parsed.error.type}` : "";
    const message = parsed.error?.message ? ` ${parsed.error.message}` : "";
    return `${type}${code}${message}`.trim() || "上游服务返回错误。";
  } catch {
    return redacted.slice(0, 220);
  }
}

function getModel() {
  return process.env.OPENAI_MODEL?.trim() || DEFAULT_MODEL;
}

function getResponsesApiUrl() {
  const baseUrl = (process.env.OPENAI_BASE_URL?.trim() || DEFAULT_OPENAI_BASE_URL).replace(/\/+$/, "");
  return baseUrl.endsWith("/v1") ? `${baseUrl}/responses` : `${baseUrl}/v1/responses`;
}

async function createStructuredResponse<T>(input: string, schema: JsonSchema, mock: T): Promise<T> {
  if (process.env.USE_MOCK_AI === "true") {
    return mock;
  }

  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    throw new MissingApiKeyError();
  }

  const response = await fetch(getResponsesApiUrl(), {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      model: getModel(),
      input,
      text: {
        format: {
          type: "json_schema",
          name: "ai_ppt_payload",
          strict: true,
          schema,
        },
      },
    }),
  });

  if (!response.ok) {
    const details = await response.text();
    throw new OpenAIRequestError(
      response.status,
      `OpenAI 模型调用失败：${response.status} ${sanitizeUpstreamError(details)}`,
    );
  }

  const data = (await response.json()) as {
    output_text?: string;
    output?: Array<{
      content?: Array<{ type?: string; text?: string }>;
    }>;
  };

  const text =
    data.output_text ??
    data.output
      ?.flatMap((item) => item.content ?? [])
      .find((content) => content.type === "output_text" || typeof content.text === "string")?.text;

  if (!text) {
    throw new Error("OpenAI 返回为空，未能解析结构化 JSON。");
  }

  return JSON.parse(text) as T;
}

function requestSummary(request: DeckRequest) {
  return [
    `主题：${request.topic}`,
    `用途：${request.purpose}`,
    `受众：${request.audience}`,
    `页数：${request.slideCount}`,
    `语言：${request.language}`,
    `风格：${request.style}`,
    `模板：${request.templateProfile?.name ?? request.templateId} / ${request.templateSource}`,
    `模板信息：${request.templateProfile ? JSON.stringify(request.templateProfile) : "未提供"}`,
    `参考资料：${request.referenceText?.slice(0, 8000) || "无"}`,
  ].join("\n");
}

function mockOutline(request: DeckRequest): OutlineSection[] {
  const base = [
    "封面与核心观点",
    "背景与问题定义",
    "关键洞察",
    "方案框架",
    "实施路径",
    "风险与资源",
    "总结与下一步",
  ];

  return base.slice(0, Math.min(base.length, Math.max(4, request.slideCount))).map((title, index) => ({
    id: `section-${index + 1}`,
    title,
    intent: index === 0 ? "快速建立主题、场景和价值判断。" : `围绕「${request.topic}」展开第 ${index + 1} 个叙事模块。`,
    slideEstimate: index === 0 || index === base.length - 1 ? 1 : Math.max(1, Math.round(request.slideCount / base.length)),
    keyPoints: [
      `${request.audience} 最关心的判断依据`,
      `与 ${request.purpose} 直接相关的可执行信息`,
      "可视化呈现建议和演讲推进逻辑",
    ],
  }));
}

function mockSlides(request: DeckRequest): SlideContent[] {
  const outline = request.outline?.length ? request.outline : mockOutline(request);
  const colors = request.templateProfile?.dominantColors ?? ["#102820", "#F8FAFC", "#C2A83E"];

  return Array.from({ length: request.slideCount }, (_, index) => {
    const section = outline[Math.min(index, outline.length - 1)];
    const isCover = index === 0;
    const isClosing = index === request.slideCount - 1;
    return {
      id: `slide-${index + 1}`,
      slideNumber: index + 1,
      sectionId: section.id,
      title: isCover ? request.topic : isClosing ? "总结与下一步" : section.title,
      body: isCover
        ? [`面向 ${request.audience}`, `${request.purpose} | ${request.language}`]
        : [
            section.keyPoints[0] ?? "关键观点",
            section.keyPoints[1] ?? "证据与说明",
            section.keyPoints[2] ?? "行动建议",
          ],
      chartSuggestion: isCover
        ? "无需图表，使用强主题视觉和 1-2 个关键信息标签。"
        : index % 4 === 0
          ? "使用 3 个指标卡 + 横向条形图突出变化方向和最高影响项。"
          : "使用小型对比图或流程图支撑本页核心判断。",
      imageSuggestion: "选择与主题直接相关的真实场景图、产品图或图标组，不使用装饰性抽象图。",
      layoutSuggestion: isCover
        ? "大标题 + 副标题 + 右侧视觉焦点，深色背景承载品牌感。"
        : index === 1
          ? "左侧目录节点，右侧进度轨道和主题色块。"
          : "左侧结论与要点，右侧证据卡片、流程图或图像占位。",
      speakerNotes: `本页用 45-60 秒说明「${section.title}」，先讲结论，再补充关键依据。`,
      colorPalette: colors,
      templateRole: isCover ? "cover" : isClosing ? "closing" : index === 1 ? "agenda" : index % 4 === 0 ? "data" : "content",
    };
  });
}

const outlineSchema: JsonSchema = {
  type: "object",
  additionalProperties: false,
  required: ["outline"],
  properties: {
    outline: {
      type: "array",
      minItems: 3,
      items: {
        type: "object",
        additionalProperties: false,
        required: ["id", "title", "intent", "slideEstimate", "keyPoints"],
        properties: {
          id: { type: "string" },
          title: { type: "string" },
          intent: { type: "string" },
          slideEstimate: { type: "integer" },
          keyPoints: { type: "array", items: { type: "string" } },
        },
      },
    },
  },
};

const slidesSchema: JsonSchema = {
  type: "object",
  additionalProperties: false,
  required: ["slides"],
  properties: {
    slides: {
      type: "array",
      items: {
        type: "object",
        additionalProperties: false,
        required: [
          "id",
          "slideNumber",
          "sectionId",
          "title",
          "body",
          "chartSuggestion",
          "imageSuggestion",
          "layoutSuggestion",
          "speakerNotes",
          "colorPalette",
          "templateRole",
        ],
        properties: {
          id: { type: "string" },
          slideNumber: { type: "integer" },
          sectionId: { type: "string" },
          title: { type: "string" },
          body: { type: "array", items: { type: "string" } },
          chartSuggestion: { type: "string" },
          imageSuggestion: { type: "string" },
          layoutSuggestion: { type: "string" },
          speakerNotes: { type: "string" },
          colorPalette: { type: "array", items: { type: "string" } },
          templateRole: {
            type: "string",
            enum: ["cover", "agenda", "content", "data", "closing"],
          },
        },
      },
    },
  },
};

export async function generateOutline(request: DeckRequest): Promise<OutlineSection[]> {
  const prompt = `你是资深 PPT 策划师。请为一个 AI 自动生成 PPT 的平台生成可编辑大纲。
要求：
- 输出语言必须是 ${request.language}
- 大纲总页数要接近 ${request.slideCount}
- 结构需要适合 ${request.purpose}，受众是 ${request.audience}
- 如果有模板信息，要沿用其封面、目录、内容页、结束页等视觉结构
- 大纲要服务于可视化表达：每个章节都应能自然落成图表、流程、对比、指标或场景图
- 不要输出 Markdown，只按 JSON schema 返回

${buildDeckCraftBrief(request)}

${requestSummary(request)}`;

  const result = await createStructuredResponse<{ outline: OutlineSection[] }>(prompt, outlineSchema, {
    outline: mockOutline(request),
  });

  return result.outline;
}

export async function generateSlides(request: DeckRequest): Promise<SlideContent[]> {
  const prompt = `你是 AI PPT 内容生成与版式建议引擎。请基于已确认大纲生成逐页内容。
每页必须包含：标题、正文要点、图表建议、图片建议、页面布局建议、演讲备注。
要求：
- 正文要点简洁，不堆砌长段落
- 页面建议应尽量匹配模板视觉风格、主色、字体和版式角色
- 每页都要能导出为近似成品页，不要把图表建议/版式建议写成给设计师看的空泛说明
- chartSuggestion 必须写出图表类型、比较维度和本页要强调的结论
- imageSuggestion 必须说明真实图片/截图/图标组的内容，不要写“配图即可”
- layoutSuggestion 必须指向明确布局，如指标卡、双栏证据、时间线、2x2 对比、流程泳道、半屏图片
- 输出页数必须为 ${request.slideCount}
- 输出语言必须是 ${request.language}
- 不要输出 Markdown，只按 JSON schema 返回

${buildDeckCraftBrief(request)}

${requestSummary(request)}

已确认大纲：
${JSON.stringify(request.outline ?? [], null, 2)}`;

  const result = await createStructuredResponse<{ slides: SlideContent[] }>(prompt, slidesSchema, {
    slides: mockSlides(request),
  });

  return result.slides.slice(0, request.slideCount).map((slide, index) => ({
    ...slide,
    slideNumber: index + 1,
    id: slide.id || `slide-${index + 1}`,
  }));
}

export async function rewriteSlide(
  slide: SlideContent,
  mode: "regenerate" | "polish" | "shorten" | "expand",
  request: DeckRequest,
): Promise<SlideContent> {
  const prompt = `你是 PPT 单页编辑助手。请根据模式改写当前页，保持 JSON schema。
模式：${mode}
语言：${request.language}
主题：${request.topic}
模板：${request.templateProfile?.name ?? request.templateId}
改写时继续遵守以下 PPT 制作规则：
${buildDeckCraftBrief(request)}
当前页：
${JSON.stringify(slide, null, 2)}`;

  const result = await createStructuredResponse<{ slides: SlideContent[] }>(prompt, slidesSchema, {
    slides: [
      {
        ...slide,
        title: mode === "shorten" ? slide.title : `${slide.title}`,
        body:
          mode === "shorten"
            ? slide.body.slice(0, 2).map((item) => item.replace(/[，,].*/, ""))
            : mode === "expand"
              ? [...slide.body, "补充关键证据、行动建议和落地条件。"]
              : slide.body.map((item) => `${item}`),
        speakerNotes:
          mode === "polish"
            ? `${slide.speakerNotes} 讲述时保持结论先行，并用一个具体例子承接。`
            : slide.speakerNotes,
      },
    ],
  });

  return { ...result.slides[0], id: slide.id, slideNumber: slide.slideNumber };
}
