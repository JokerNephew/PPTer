"use client";

import {
  AlertCircle,
  ArrowDownToLine,
  ArrowUpRight,
  BookOpen,
  Bot,
  Check,
  ChevronRight,
  FileText,
  ImagePlus,
  Layers,
  Loader2,
  PanelLeft,
  Palette,
  Pencil,
  Plus,
  RefreshCw,
  Search,
  Scissors,
  Sparkles,
  Trash2,
  Upload,
  Wand2,
  X,
} from "lucide-react";
import { type FormEvent, useMemo, useRef, useState } from "react";
import { builtInTemplates, getCategoryLabel } from "@/lib/ppt/templates";
import type {
  ApiErrorBody,
  DeckExportPayload,
  DeckRequest,
  ExternalTemplateResult,
  GenerationStage,
  GenerationStyle,
  OutlineSection,
  SlideContent,
  TemplateSearchResponse,
  TemplateProfile,
} from "@/types/ppt";

const modelName = "GPT-5.5";

const generationStages: Array<{ id: GenerationStage; label: string }> = [
  { id: "reading-template", label: "读取模板" },
  { id: "analyzing-topic", label: "分析主题" },
  { id: "generating-outline", label: "生成大纲" },
  { id: "generating-slides", label: "生成页面内容" },
  { id: "applying-template", label: "应用模板样式" },
  { id: "preparing-export", label: "准备导出文件" },
];

const styleOptions: Array<{ value: GenerationStyle; label: string }> = [
  { value: "professional", label: "专业简洁" },
  { value: "minimal", label: "极简高效" },
  { value: "visual", label: "视觉叙事" },
  { value: "academic", label: "学术严谨" },
  { value: "executive", label: "高管汇报" },
  { value: "storytelling", label: "故事化表达" },
];

interface DeckFormState {
  topic: string;
  purpose: string;
  audience: string;
  slideCount: number;
  language: string;
  style: GenerationStyle;
}

const initialFormState: DeckFormState = {
  topic: "AI 生成 PPT 平台产品方案",
  purpose: "产品汇报",
  audience: "管理层和产品团队",
  slideCount: 10,
  language: "中文",
  style: "professional",
};

function createRequest(
  formState: DeckFormState,
  template: TemplateProfile,
  referenceText: string,
  outline?: OutlineSection[],
): DeckRequest {
  return {
    topic: formState.topic,
    purpose: formState.purpose,
    audience: formState.audience,
    slideCount: formState.slideCount,
    language: formState.language,
    style: formState.style,
    templateId: template.id,
    templateSource: template.source,
    templateProfile: template,
    referenceText,
    outline,
  };
}

function stageIndex(stage: GenerationStage) {
  return generationStages.findIndex((item) => item.id === stage);
}

async function readApiError(response: Response, fallback: string) {
  try {
    const body = (await response.json()) as ApiErrorBody;
    return body.error || fallback;
  } catch {
    return fallback;
  }
}

function createExternalTemplateProfile(result: ExternalTemplateResult): TemplateProfile {
  const tagText = result.tags.length ? result.tags.slice(0, 4).join(" / ") : "external template";

  return {
    id: `external-${result.id}`,
    name: result.title,
    source: "built-in",
    category: "business",
    pageCount: result.slideCount ?? 12,
    aspectRatio: result.aspectRatio === "4:3" ? "4:3" : result.aspectRatio === "A4" ? "unknown" : "16:9",
    dominantColors: ["#111827", "#F8FAFC", "#2563EB", "#F59E0B"],
    fonts: ["Aptos Display", "Aptos"],
    titleHierarchy: ["source template title", "section heading", "body placeholder"],
    detectedLayouts: ["cover", "agenda", "content", "image", "closing"],
    hasCover: true,
    hasAgenda: true,
    hasClosing: true,
    thumbnail: result.thumbnail,
    warnings: [
      "外部模板尚未下载解析，当前只作为生成风格参考。下载 PPTX 后上传，可提取真实配色、字体和页数。",
    ],
    fallbackStrategy: `Use the visual direction from ${result.source}: ${tagText}. Keep layouts close to ${result.title}.`,
    externalUrl: result.url,
    externalSource: result.source,
    externalLicenseNote: result.licenseNote,
  };
}

function downloadBlob(blob: Blob, fileName: string) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function splitLines(value: string) {
  return value
    .split("\n")
    .map((item) => item.trim())
    .filter(Boolean);
}

function SlidePreview({ slide, template }: { slide?: SlideContent; template: TemplateProfile }) {
  const palette = slide?.colorPalette?.length ? slide.colorPalette : template.dominantColors;
  const [dark = "#102820", light = "#F8FAFC", accent = "#C2A83E"] = palette;

  if (!slide) {
    return (
      <div className="preview-shell min-h-[420px] items-center justify-center border border-dashed border-slate-300 bg-slate-50 text-slate-500">
        <div className="text-center">
          <Layers className="mx-auto mb-3 h-9 w-9" />
          <p className="text-sm font-medium">确认大纲后生成页面预览</p>
        </div>
      </div>
    );
  }

  const isCover = slide.templateRole === "cover";
  const isClosing = slide.templateRole === "closing";

  return (
    <div
      className="preview-shell shadow-[0_24px_70px_rgba(15,23,42,0.18)]"
      style={{
        background: isCover || isClosing ? dark : light,
        color: isCover || isClosing ? "#fff" : "#172033",
      }}
    >
      <div
        className="absolute right-0 top-0 h-full w-[31%]"
        style={{ background: isCover || isClosing ? accent : "rgba(255,255,255,0.78)" }}
      />
      <div className="relative z-10 flex h-full flex-col justify-between p-8">
        <div>
          <div className="mb-4 flex items-center gap-2 text-[11px] font-semibold uppercase tracking-[0.18em] opacity-75">
            <span className="h-2 w-2 rounded-full" style={{ background: accent }} />
            {template.name} · {slide.templateRole}
          </div>
          <h2 className="max-w-[72%] text-3xl font-semibold leading-tight">{slide.title}</h2>
        </div>

        <div className="grid grid-cols-[1.05fr_0.95fr] gap-5">
          <div className="rounded-md bg-white/88 p-5 text-slate-900 shadow-sm">
            <div className="mb-3 text-xs font-semibold uppercase tracking-[0.14em] text-slate-500">正文要点</div>
            <ul className="space-y-2 text-sm leading-6">
              {slide.body.map((point) => (
                <li key={point} className="flex gap-2">
                  <span className="mt-2 h-1.5 w-1.5 shrink-0 rounded-full" style={{ background: accent }} />
                  <span>{point}</span>
                </li>
              ))}
            </ul>
          </div>
          <div className="space-y-3 text-sm">
            <div className="rounded-md bg-black/12 p-4 backdrop-blur">
              <p className="mb-1 text-xs font-semibold opacity-70">图表建议</p>
              <p className="leading-5">{slide.chartSuggestion}</p>
            </div>
            <div className="rounded-md bg-black/12 p-4 backdrop-blur">
              <p className="mb-1 text-xs font-semibold opacity-70">版式建议</p>
              <p className="leading-5">{slide.layoutSuggestion}</p>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

function ExternalTemplateCard({
  template,
  density = "comfortable",
  onApply,
}: {
  template: ExternalTemplateResult;
  density?: "comfortable" | "compact";
  onApply: (template: ExternalTemplateResult) => void;
}) {
  return (
    <article className={`external-template-card ${density === "compact" ? "external-template-card-compact" : ""}`}>
      <a href={template.url} target="_blank" rel="noreferrer" className="external-template-thumb">
        {template.thumbnail ? (
          // eslint-disable-next-line @next/next/no-img-element
          <img src={template.thumbnail} alt={template.title} />
        ) : (
          <span>{template.source}</span>
        )}
      </a>
      <div className="min-w-0 flex-1">
        <div className="mb-1 flex items-start justify-between gap-2">
          <h4 className="line-clamp-2 text-sm font-semibold leading-snug text-slate-950">{template.title}</h4>
          <span className="external-source-badge">{template.source}</span>
        </div>
        <p className="line-clamp-2 text-xs leading-5 text-slate-500">{template.description || template.licenseNote}</p>
        <div className="mt-2 flex flex-wrap gap-1.5 text-[11px] font-semibold text-slate-500">
          {template.formats.slice(0, 3).map((format) => (
            <span key={format} className="external-meta-pill">
              {format}
            </span>
          ))}
          {template.aspectRatio ? <span className="external-meta-pill">{template.aspectRatio}</span> : null}
          {template.slideCount ? <span className="external-meta-pill">{template.slideCount} 页</span> : null}
        </div>
        <div className="mt-3 grid grid-cols-2 gap-2">
          <button type="button" className="tool-button px-2 py-1.5 text-xs" onClick={() => onApply(template)}>
            <Palette className="h-3.5 w-3.5" />
            借用风格
          </button>
          <a href={template.downloadUrl ?? template.url} target="_blank" rel="noreferrer" className="tool-button px-2 py-1.5 text-xs">
            <ArrowUpRight className="h-3.5 w-3.5" />
            打开下载
          </a>
        </div>
      </div>
    </article>
  );
}

export function AiPptWorkbench() {
  const [formState, setFormState] = useState<DeckFormState>(initialFormState);
  const [selectedTemplate, setSelectedTemplate] = useState<TemplateProfile>(builtInTemplates[0]);
  const [uploadedTemplate, setUploadedTemplate] = useState<TemplateProfile | null>(null);
  const [referenceText, setReferenceText] = useState("");
  const [outline, setOutline] = useState<OutlineSection[]>([]);
  const [slides, setSlides] = useState<SlideContent[]>([]);
  const [activeSlideId, setActiveSlideId] = useState<string | null>(null);
  const [stage, setStage] = useState<GenerationStage>("idle");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [exportName, setExportName] = useState("AI-PPT-Deck");
  const [exportFormat, setExportFormat] = useState<"pptx" | "pdf">("pptx");
  const [exportStatus, setExportStatus] = useState("待导出");
  const [templateQuery, setTemplateQuery] = useState("business report");
  const [externalTemplates, setExternalTemplates] = useState<ExternalTemplateResult[]>([]);
  const [templateSearchStatus, setTemplateSearchStatus] = useState("可搜索 SlidesCarnival / PresentationGO / PPTMON / Slidesgo");
  const [templateSearchBusy, setTemplateSearchBusy] = useState(false);
  const [templateLibraryOpen, setTemplateLibraryOpen] = useState(false);
  const templateInputRef = useRef<HTMLInputElement>(null);
  const referenceInputRef = useRef<HTMLInputElement>(null);

  const activeSlide = useMemo(
    () => slides.find((slide) => slide.id === activeSlideId) ?? slides[0],
    [activeSlideId, slides],
  );

  const currentRequest = useMemo(
    () => createRequest(formState, selectedTemplate, referenceText, outline),
    [formState, selectedTemplate, referenceText, outline],
  );

  const completedStageIndex = stage === "done" ? generationStages.length - 1 : stageIndex(stage);

  async function handleTemplateSearch(event?: FormEvent) {
    event?.preventDefault();
    const query = templateQuery.trim();
    if (!query) {
      setTemplateSearchStatus("请输入关键词，例如 business report、pitch deck、education。");
      return;
    }

    setTemplateSearchBusy(true);
    setTemplateSearchStatus("正在连接免费模板站点...");
    setError("");

    try {
      const response = await fetch(`/api/search-templates?q=${encodeURIComponent(query)}`);
      if (!response.ok) throw new Error(await readApiError(response, "模板搜索失败。"));
      const body = (await response.json()) as TemplateSearchResponse;
      setExternalTemplates(body.results);
      const sourceSummary = body.sources
        .map((source) => `${source.name}${source.status === "ok" ? "✓" : "未响应"}`)
        .join(" · ");
      setTemplateSearchStatus(body.results.length ? `找到 ${body.results.length} 个免费模板 · ${sourceSummary}` : `没有找到结果 · ${sourceSummary}`);
    } catch (event) {
      setTemplateSearchStatus(event instanceof Error ? event.message : "模板搜索失败。");
    } finally {
      setTemplateSearchBusy(false);
    }
  }

  function applyExternalTemplate(result: ExternalTemplateResult) {
    const template = createExternalTemplateProfile(result);
    setSelectedTemplate(template);
    setTemplateSearchStatus(`已借用 ${result.source} 的「${result.title}」作为风格参考。下载 PPTX 后上传可完整解析。`);
    setTemplateLibraryOpen(false);
  }

  async function handleTemplateUpload(file?: File) {
    if (!file) return;
    setError("");
    setBusy(true);
    setStage("reading-template");
    const formData = new FormData();
    formData.append("template", file);

    try {
      const response = await fetch("/api/analyze-template", { method: "POST", body: formData });
      if (!response.ok) throw new Error(await readApiError(response, "模板解析失败。"));
      const body = (await response.json()) as { template: TemplateProfile };
      setUploadedTemplate(body.template);
      setSelectedTemplate(body.template);
      setStage("idle");
    } catch (event) {
      setStage("error");
      setError(event instanceof Error ? event.message : "模板解析失败。");
    } finally {
      setBusy(false);
    }
  }

  async function handleReferenceUpload(file?: File) {
    if (!file) return;
    setError("");
    const formData = new FormData();
    formData.append("reference", file);

    try {
      const response = await fetch("/api/extract-reference", { method: "POST", body: formData });
      if (!response.ok) throw new Error(await readApiError(response, "参考资料读取失败。"));
      const body = (await response.json()) as { text: string; file: { name: string } };
      setReferenceText((current) => `${current}\n\n[${body.file.name}]\n${body.text}`.trim());
    } catch (event) {
      setError(event instanceof Error ? event.message : "参考资料读取失败。");
    }
  }

  async function handleGenerateOutline() {
    setBusy(true);
    setError("");
    setSlides([]);
    setActiveSlideId(null);
    setStage("analyzing-topic");

    try {
      await new Promise((resolve) => setTimeout(resolve, 260));
      setStage("generating-outline");
      const request = createRequest(formState, selectedTemplate, referenceText);
      const response = await fetch("/api/generate-outline", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(request),
      });
      if (!response.ok) throw new Error(await readApiError(response, "生成大纲失败。"));
      const body = (await response.json()) as { outline: OutlineSection[] };
      setOutline(body.outline);
      setStage("done");
    } catch (event) {
      setStage("error");
      setError(event instanceof Error ? event.message : "生成大纲失败。");
    } finally {
      setBusy(false);
    }
  }

  async function handleGenerateSlides() {
    setBusy(true);
    setError("");
    setStage("generating-slides");

    try {
      const request = createRequest(formState, selectedTemplate, referenceText, outline);
      const response = await fetch("/api/generate-slides", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(request),
      });
      if (!response.ok) throw new Error(await readApiError(response, "生成页面内容失败。"));
      setStage("applying-template");
      const body = (await response.json()) as { slides: SlideContent[] };
      await new Promise((resolve) => setTimeout(resolve, 320));
      setSlides(body.slides);
      setActiveSlideId(body.slides[0]?.id ?? null);
      setExportName(`${request.topic || "AI-PPT-Deck"}`.slice(0, 48));
      setStage("done");
    } catch (event) {
      setStage("error");
      setError(event instanceof Error ? event.message : "生成页面内容失败。");
    } finally {
      setBusy(false);
    }
  }

  async function rewriteCurrentSlide(mode: "regenerate" | "polish" | "shorten" | "expand") {
    if (!activeSlide) return;
    setBusy(true);
    setError("");
    setStage("generating-slides");

    try {
      const response = await fetch("/api/rewrite-slide", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ slide: activeSlide, mode, request: currentRequest }),
      });
      if (!response.ok) throw new Error(await readApiError(response, "单页改写失败。"));
      const body = (await response.json()) as { slide: SlideContent };
      setSlides((current) => current.map((slide) => (slide.id === activeSlide.id ? body.slide : slide)));
      setStage("done");
    } catch (event) {
      setStage("error");
      setError(event instanceof Error ? event.message : "单页改写失败。");
    } finally {
      setBusy(false);
    }
  }

  function updateOutlineSection(id: string, patch: Partial<OutlineSection>) {
    setOutline((current) => current.map((section) => (section.id === id ? { ...section, ...patch } : section)));
  }

  function addOutlineSection() {
    setOutline((current) => [
      ...current,
      {
        id: `section-${Date.now()}`,
        title: "新增章节",
        intent: "补充本段叙事目的。",
        slideEstimate: 1,
        keyPoints: ["关键观点", "证据或案例", "行动建议"],
      },
    ]);
  }

  function updateActiveSlide(patch: Partial<SlideContent>) {
    if (!activeSlide) return;
    setSlides((current) => current.map((slide) => (slide.id === activeSlide.id ? { ...slide, ...patch } : slide)));
  }

  function reapplyTemplateStyle() {
    setSlides((current) =>
      current.map((slide) => ({
        ...slide,
        colorPalette: selectedTemplate.dominantColors,
        layoutSuggestion: `${selectedTemplate.detectedLayouts.slice(0, 3).join(" / ")}：${slide.layoutSuggestion}`,
      })),
    );
    setStage("done");
  }

  async function exportDeck() {
    if (!slides.length) {
      setError("请先生成页面内容再导出。");
      return;
    }

    setBusy(true);
    setStage("preparing-export");
    setExportStatus("导出中");
    setError("");

    try {
      const payload: DeckExportPayload = {
        fileName: exportName,
        format: exportFormat,
        request: currentRequest,
        slides,
        templateProfile: selectedTemplate,
      };
      const response = await fetch("/api/export", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      if (!response.ok) throw new Error(await readApiError(response, "导出失败。"));
      const blob = await response.blob();
      downloadBlob(blob, `${exportName || "ai-ppt-deck"}.${exportFormat}`);
      setExportStatus("已导出");
      setStage("done");
    } catch (event) {
      setStage("error");
      setExportStatus("导出失败");
      setError(event instanceof Error ? event.message : "导出失败。");
    } finally {
      setBusy(false);
    }
  }

  return (
    <main className="min-h-screen bg-[#eef2f5] text-slate-950">
      <div className="flex min-h-screen flex-col">
        <header className="border-b border-slate-200 bg-white/92 px-4 py-3 backdrop-blur md:px-6">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div className="flex items-center gap-3">
              <div className="flex h-9 w-9 items-center justify-center rounded-md bg-[#102820] text-white">
                <Sparkles className="h-5 w-5" />
              </div>
              <div>
                <h1 className="text-lg font-semibold tracking-tight">AI PPT Studio</h1>
                <p className="text-xs text-slate-500">输入需求 → 上传模板 → 生成大纲 → 编辑页面 → 导出文件</p>
              </div>
            </div>
            <div className="flex flex-wrap items-center gap-2 text-xs">
              <span className="inline-flex items-center gap-2 rounded-md border border-slate-200 bg-white px-3 py-2 font-medium text-slate-700">
                <Bot className="h-4 w-4 text-[#0f766e]" />
                当前模型：{modelName}
              </span>
              <span className="inline-flex items-center gap-2 rounded-md border border-slate-200 bg-white px-3 py-2 font-medium text-slate-700">
                <Palette className="h-4 w-4 text-[#c28f2c]" />
                模板来源：
                {selectedTemplate.externalSource ?? (selectedTemplate.source === "uploaded" ? "用户上传" : "内置模板")}
              </span>
              <span className="inline-flex items-center gap-2 rounded-md border border-slate-200 bg-white px-3 py-2 font-medium text-slate-700">
                <ArrowDownToLine className="h-4 w-4 text-[#334155]" />
                导出状态：{exportStatus}
              </span>
            </div>
          </div>
        </header>

        <section className="border-b border-slate-200 bg-[#f7f9fb] px-4 py-3 md:px-6">
          <div className="grid gap-2 md:grid-cols-6">
            {generationStages.map((item, index) => {
              const active = item.id === stage;
              const done = completedStageIndex >= index && stage !== "error" && stage !== "idle";
              return (
                <div
                  key={item.id}
                  className={`flex items-center gap-2 rounded-md border px-3 py-2 text-xs font-medium ${
                    active
                      ? "border-[#0f766e] bg-white text-[#0f766e]"
                      : done
                        ? "border-slate-200 bg-white text-slate-700"
                        : "border-slate-200 bg-transparent text-slate-400"
                  }`}
                >
                  {active && busy ? <Loader2 className="h-3.5 w-3.5 animate-spin" /> : done ? <Check className="h-3.5 w-3.5" /> : <span className="h-3.5 w-3.5 rounded-full border border-current" />}
                  {item.label}
                </div>
              );
            })}
          </div>
          {error ? (
            <div className="mt-3 flex items-start gap-2 rounded-md border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-700">
              <AlertCircle className="mt-0.5 h-4 w-4 shrink-0" />
              <span>{error}</span>
            </div>
          ) : null}
        </section>

        <div className="grid flex-1 grid-cols-1 lg:grid-cols-[360px_minmax(460px,1fr)_380px]">
          <aside className="border-r border-slate-200 bg-white">
            <div className="border-b border-slate-200 p-4">
              <form className="space-y-3">
                <label className="block text-xs font-semibold text-slate-500">主题</label>
                <input
                  name="topic"
                  value={formState.topic}
                  onChange={(event) => setFormState((current) => ({ ...current, topic: event.target.value }))}
                  className="field"
                  placeholder="例如：AI 生成 PPT 平台商业计划"
                />
                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="block text-xs font-semibold text-slate-500">用途</label>
                    <input
                      name="purpose"
                      value={formState.purpose}
                      onChange={(event) => setFormState((current) => ({ ...current, purpose: event.target.value }))}
                      className="field"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-500">受众</label>
                    <input
                      name="audience"
                      value={formState.audience}
                      onChange={(event) => setFormState((current) => ({ ...current, audience: event.target.value }))}
                      className="field"
                    />
                  </div>
                </div>
                <div className="grid grid-cols-3 gap-3">
                  <div>
                    <label className="block text-xs font-semibold text-slate-500">页数</label>
                    <input
                      name="slideCount"
                      type="number"
                      min={4}
                      max={40}
                      value={formState.slideCount}
                      onChange={(event) =>
                        setFormState((current) => ({
                          ...current,
                          slideCount: Math.min(40, Math.max(4, Number(event.target.value) || 4)),
                        }))
                      }
                      className="field"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-500">语言</label>
                    <select
                      name="language"
                      value={formState.language}
                      onChange={(event) => setFormState((current) => ({ ...current, language: event.target.value }))}
                      className="field"
                    >
                      <option>中文</option>
                      <option>English</option>
                      <option>日本語</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-500">风格</label>
                    <select
                      name="style"
                      value={formState.style}
                      onChange={(event) => setFormState((current) => ({ ...current, style: event.target.value as GenerationStyle }))}
                      className="field"
                    >
                      {styleOptions.map((item) => (
                        <option key={item.value} value={item.value}>
                          {item.label}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>
              </form>

              <div className="mt-4 grid grid-cols-2 gap-2">
                <button className="tool-button" type="button" onClick={() => templateInputRef.current?.click()}>
                  <Upload className="h-4 w-4" />
                  上传模板
                </button>
                <button className="tool-button" type="button" onClick={() => referenceInputRef.current?.click()}>
                  <FileText className="h-4 w-4" />
                  上传资料
                </button>
                <input
                  ref={templateInputRef}
                  type="file"
                  accept=".pptx"
                  className="hidden"
                  onChange={(event) => handleTemplateUpload(event.target.files?.[0])}
                />
                <input
                  ref={referenceInputRef}
                  type="file"
                  accept=".doc,.docx,.pdf,.txt,.md,.pptx"
                  className="hidden"
                  onChange={(event) => handleReferenceUpload(event.target.files?.[0])}
                />
              </div>

              <label className="mt-4 block text-xs font-semibold text-slate-500">参考文本</label>
              <textarea
                value={referenceText}
                onChange={(event) => setReferenceText(event.target.value)}
                className="field min-h-[92px] resize-y"
                placeholder="粘贴产品文档、调研资料、会议纪要或演讲重点"
              />

              <div className="mt-4 flex gap-2">
                <button className="primary-button flex-1" type="button" disabled={busy} onClick={handleGenerateOutline}>
                  {busy && stage === "generating-outline" ? <Loader2 className="h-4 w-4 animate-spin" /> : <Wand2 className="h-4 w-4" />}
                  生成大纲
                </button>
                <button className="secondary-button" type="button" disabled={busy || outline.length === 0} onClick={handleGenerateSlides}>
                  <ChevronRight className="h-4 w-4" />
                  生成页面
                </button>
              </div>
            </div>

            <div className="p-4">
              <div className="mb-3 flex items-center justify-between">
                <h2 className="text-sm font-semibold">模板系统</h2>
                <button className="icon-button" type="button" title="重新应用模板样式" onClick={reapplyTemplateStyle}>
                  <RefreshCw className="h-4 w-4" />
                </button>
              </div>
              <button type="button" className="template-library-button" onClick={() => setTemplateLibraryOpen(true)}>
                <span className="flex h-10 w-10 shrink-0 items-center justify-center rounded-md bg-[#102820] text-white">
                  <Search className="h-5 w-5" />
                </span>
                <span className="min-w-0 flex-1 text-left">
                  <span className="block text-sm font-semibold text-slate-950">打开免费模板库</span>
                  <span className="block truncate text-xs text-slate-500">{templateSearchStatus}</span>
                </span>
                <ChevronRight className="h-4 w-4 text-slate-400" />
              </button>

              {uploadedTemplate ? (
                <button
                  type="button"
                  className={`template-row ${selectedTemplate.id === uploadedTemplate.id ? "template-row-active" : ""}`}
                  onClick={() => setSelectedTemplate(uploadedTemplate)}
                >
                  <span className="template-preview" style={{ background: uploadedTemplate.thumbnail }} />
                  <span className="min-w-0">
                    <span className="block truncate text-sm font-semibold">{uploadedTemplate.name}</span>
                    <span className="block text-xs text-slate-500">{uploadedTemplate.pageCount} 页 · 上传模板</span>
                  </span>
                </button>
              ) : null}
              {selectedTemplate.externalUrl ? (
                <div className="mb-3 rounded-md border border-blue-100 bg-blue-50 p-3 text-xs leading-5 text-blue-900">
                  当前使用外部模板风格参考。
                  <a href={selectedTemplate.externalUrl} target="_blank" rel="noreferrer" className="ml-1 font-semibold underline">
                    打开来源
                  </a>
                  <div className="mt-1 text-blue-800">{selectedTemplate.externalLicenseNote}</div>
                </div>
              ) : null}
              <div className="space-y-2">
                {builtInTemplates.map((template) => (
                  <button
                    type="button"
                    key={template.id}
                    className={`template-row ${selectedTemplate.id === template.id ? "template-row-active" : ""}`}
                    onClick={() => setSelectedTemplate(template)}
                  >
                    <span className="template-preview" style={{ background: template.thumbnail }} />
                    <span className="min-w-0">
                      <span className="block truncate text-sm font-semibold">{template.name}</span>
                      <span className="block text-xs text-slate-500">{getCategoryLabel(template.category)} · {template.pageCount} 页</span>
                    </span>
                  </button>
                ))}
              </div>
            </div>
          </aside>

          {templateLibraryOpen ? (
            <div className="template-library-overlay" role="dialog" aria-modal="true" aria-label="免费模板库">
              <div className="template-library-modal">
                <div className="template-library-header">
                  <div>
                    <div className="mb-1 text-xs font-semibold uppercase text-[#0f766e]">Free Template Library</div>
                    <h2 className="text-2xl font-semibold tracking-tight text-slate-950">免费 PPT 模板库</h2>
                    <p className="mt-1 text-sm text-slate-500">搜索多个免费模板站点，挑选后可借用风格，下载 PPTX 后上传可完整解析。</p>
                  </div>
                  <button className="icon-button" type="button" title="关闭模板库" onClick={() => setTemplateLibraryOpen(false)}>
                    <X className="h-4 w-4" />
                  </button>
                </div>

                <form className="template-library-search" onSubmit={handleTemplateSearch}>
                  <div className="relative min-w-0 flex-1">
                    <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
                    <input
                      value={templateQuery}
                      onChange={(event) => setTemplateQuery(event.target.value)}
                      className="template-library-search-input"
                      placeholder="搜索 business report、pitch deck、education、medical、minimal..."
                      autoFocus
                    />
                  </div>
                  <button className="primary-button px-5" type="submit" disabled={templateSearchBusy}>
                    {templateSearchBusy ? <Loader2 className="h-4 w-4 animate-spin" /> : <Search className="h-4 w-4" />}
                    搜索模板
                  </button>
                </form>

                <div className="template-library-toolbar">
                  <span>{templateSearchStatus}</span>
                  <span>{externalTemplates.length ? `当前展示 ${externalTemplates.length} 个模板` : "输入关键词后开始搜索"}</span>
                </div>

                <div className="template-library-results">
                  {externalTemplates.length ? (
                    <div className="template-library-grid">
                      {externalTemplates.map((template) => (
                        <ExternalTemplateCard key={template.id} template={template} onApply={applyExternalTemplate} />
                      ))}
                    </div>
                  ) : (
                    <div className="template-library-empty">
                      <Search className="mx-auto mb-3 h-8 w-8 text-slate-400" />
                      <p className="text-sm font-semibold text-slate-700">搜索关键词后，这里会展示外部免费模板。</p>
                      <p className="mt-1 text-xs text-slate-500">可以试试 business report、pitch deck、education 或 minimal。</p>
                    </div>
                  )}
                </div>
              </div>
            </div>
          ) : null}

          <section className="min-w-0 bg-[#eef2f5] p-4 md:p-6">
            <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
              <div>
                <h2 className="text-base font-semibold">页面预览</h2>
                <p className="text-xs text-slate-500">
                  {selectedTemplate.name} · {selectedTemplate.externalSource ?? selectedTemplate.aspectRatio} · {selectedTemplate.fonts.join(" / ")}
                </p>
              </div>
              <div className="flex items-center gap-2">
                {selectedTemplate.dominantColors.slice(0, 5).map((color) => (
                  <span key={color} className="h-5 w-5 rounded-full border border-white shadow-sm" style={{ background: color }} />
                ))}
              </div>
            </div>

            <SlidePreview slide={activeSlide} template={selectedTemplate} />

            <div className="mt-5 grid gap-4 2xl:grid-cols-[190px_minmax(0,1fr)]">
              <div className="rounded-md border border-slate-200 bg-white p-3">
                <div className="mb-3 flex items-center gap-2 text-xs font-semibold text-slate-500">
                  <PanelLeft className="h-4 w-4" />
                  页面缩略图
                </div>
                <div className="grid grid-cols-2 gap-2 md:grid-cols-3 2xl:grid-cols-1">
                  {(slides.length ? slides : outline).map((item, index) => {
                    const isSlide = "slideNumber" in item;
                    const id = isSlide ? item.id : item.id;
                    const title = isSlide ? item.title : item.title;
                    return (
                      <button
                        type="button"
                        key={id}
                        onClick={() => isSlide && setActiveSlideId(id)}
                        className={`thumb ${activeSlide?.id === id ? "thumb-active" : ""}`}
                      >
                        <span className="text-[11px] font-semibold text-slate-400">{index + 1}</span>
                        <span className="line-clamp-2 text-left text-xs font-medium">{title}</span>
                      </button>
                    );
                  })}
                </div>
              </div>

              <div className="rounded-md border border-slate-200 bg-white p-4 shadow-sm">
                <div className="mb-4 flex flex-wrap items-center justify-between gap-3 border-b border-slate-100 pb-3">
                  <div>
                    <h3 className="text-base font-semibold text-slate-950">大纲编辑</h3>
                    <p className="text-xs font-medium text-slate-500">{outline.length ? `${outline.length} 个章节` : "等待生成"}</p>
                  </div>
                  <button className="tool-button px-3 py-2 text-xs" type="button" onClick={addOutlineSection}>
                    <Plus className="h-3.5 w-3.5" />
                    增加章节
                  </button>
                </div>
                <div className="outline-grid">
                  {outline.map((section, index) => (
                    <div key={section.id} className="outline-card">
                      <div className="mb-4 flex items-start justify-between gap-3">
                        <span className="outline-index">{index + 1}</span>
                        <div className="min-w-0 flex-1">
                          <label className="sr-only" htmlFor={`${section.id}-title`}>
                            章节标题
                          </label>
                          <input
                            id={`${section.id}-title`}
                            className="outline-title-input"
                            value={section.title}
                            onChange={(event) => updateOutlineSection(section.id, { title: event.target.value })}
                          />
                          <div className="mt-1 text-xs font-medium text-slate-500">预计 {section.slideEstimate} 页</div>
                        </div>
                        <button
                          className="icon-button h-8 w-8 shrink-0"
                          type="button"
                          title="删除章节"
                          onClick={() => setOutline((current) => current.filter((item) => item.id !== section.id))}
                        >
                          <Trash2 className="h-3.5 w-3.5" />
                        </button>
                      </div>
                      <div className="space-y-3">
                        <label className="block">
                          <span className="outline-field-label">章节目标</span>
                          <textarea
                            value={section.intent}
                            onChange={(event) => updateOutlineSection(section.id, { intent: event.target.value })}
                            className="outline-textarea min-h-[90px]"
                            rows={3}
                          />
                        </label>
                        <label className="block">
                          <span className="outline-field-label">关键要点</span>
                          <textarea
                            value={section.keyPoints.join("\n")}
                            onChange={(event) => updateOutlineSection(section.id, { keyPoints: splitLines(event.target.value) })}
                            className="outline-textarea min-h-[128px]"
                            rows={5}
                          />
                        </label>
                      </div>
                    </div>
                  ))}
                  {outline.length === 0 ? (
                    <div className="outline-empty">
                      <BookOpen className="mx-auto mb-2 h-5 w-5" />
                      等待生成可编辑大纲
                    </div>
                  ) : null}
                </div>
              </div>
            </div>
          </section>

          <aside className="border-l border-slate-200 bg-white">
            <div className="border-b border-slate-200 p-4">
              <h2 className="text-sm font-semibold">编辑面板</h2>
              <p className="text-xs text-slate-500">标题、正文、备注、配色、版式、图片建议</p>
            </div>

            <div className="space-y-4 p-4">
              {activeSlide ? (
                <>
                  <div>
                    <label className="block text-xs font-semibold text-slate-500">页面标题</label>
                    <input className="field" value={activeSlide.title} onChange={(event) => updateActiveSlide({ title: event.target.value })} />
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-500">正文要点</label>
                    <textarea
                      className="field min-h-[120px]"
                      value={activeSlide.body.join("\n")}
                      onChange={(event) => updateActiveSlide({ body: splitLines(event.target.value) })}
                    />
                  </div>
                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block text-xs font-semibold text-slate-500">页面角色</label>
                      <select
                        className="field"
                        value={activeSlide.templateRole}
                        onChange={(event) => updateActiveSlide({ templateRole: event.target.value as SlideContent["templateRole"] })}
                      >
                        <option value="cover">封面</option>
                        <option value="agenda">目录</option>
                        <option value="content">内容页</option>
                        <option value="data">数据页</option>
                        <option value="closing">结束页</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-xs font-semibold text-slate-500">主色</label>
                      <input
                        type="color"
                        className="h-10 w-full rounded-md border border-slate-200 bg-white p-1"
                        value={activeSlide.colorPalette[0] ?? "#102820"}
                        onChange={(event) => updateActiveSlide({ colorPalette: [event.target.value, ...activeSlide.colorPalette.slice(1)] })}
                      />
                    </div>
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-500">图表建议</label>
                    <textarea
                      className="field min-h-[70px]"
                      value={activeSlide.chartSuggestion}
                      onChange={(event) => updateActiveSlide({ chartSuggestion: event.target.value })}
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-500">图片建议</label>
                    <textarea
                      className="field min-h-[70px]"
                      value={activeSlide.imageSuggestion}
                      onChange={(event) => updateActiveSlide({ imageSuggestion: event.target.value })}
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-500">版式建议</label>
                    <textarea
                      className="field min-h-[70px]"
                      value={activeSlide.layoutSuggestion}
                      onChange={(event) => updateActiveSlide({ layoutSuggestion: event.target.value })}
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-500">演讲备注</label>
                    <textarea
                      className="field min-h-[110px]"
                      value={activeSlide.speakerNotes}
                      onChange={(event) => updateActiveSlide({ speakerNotes: event.target.value })}
                    />
                  </div>

                  <div className="grid grid-cols-2 gap-2">
                    <button className="tool-button" type="button" disabled={busy} onClick={() => rewriteCurrentSlide("regenerate")}>
                      <RefreshCw className="h-4 w-4" />
                      重新生成
                    </button>
                    <button className="tool-button" type="button" disabled={busy} onClick={() => rewriteCurrentSlide("polish")}>
                      <Sparkles className="h-4 w-4" />
                      优化文案
                    </button>
                    <button className="tool-button" type="button" disabled={busy} onClick={() => rewriteCurrentSlide("shorten")}>
                      <Scissors className="h-4 w-4" />
                      缩短内容
                    </button>
                    <button className="tool-button" type="button" disabled={busy} onClick={() => rewriteCurrentSlide("expand")}>
                      <Pencil className="h-4 w-4" />
                      扩展内容
                    </button>
                    <button className="tool-button col-span-2" type="button" onClick={() => templateInputRef.current?.click()}>
                      <ImagePlus className="h-4 w-4" />
                      切换模板
                    </button>
                  </div>
                </>
              ) : (
                <div className="rounded-md border border-dashed border-slate-300 p-6 text-center text-sm text-slate-500">
                  生成页面后可编辑每一页内容
                </div>
              )}

              <div className="rounded-md border border-slate-200 p-3">
                <h3 className="mb-3 text-sm font-semibold">导出预览</h3>
                <label className="block text-xs font-semibold text-slate-500">文件名</label>
                <input className="field" value={exportName} onChange={(event) => setExportName(event.target.value)} />
                <div className="mt-3 grid grid-cols-2 gap-2">
                  <button
                    type="button"
                    className={`segmented ${exportFormat === "pptx" ? "segmented-active" : ""}`}
                    onClick={() => setExportFormat("pptx")}
                  >
                    PPTX
                  </button>
                  <button
                    type="button"
                    className={`segmented ${exportFormat === "pdf" ? "segmented-active" : ""}`}
                    onClick={() => setExportFormat("pdf")}
                  >
                    PDF
                  </button>
                </div>
                <button className="primary-button mt-3 w-full" type="button" disabled={busy || !slides.length} onClick={exportDeck}>
                  {busy && stage === "preparing-export" ? <Loader2 className="h-4 w-4 animate-spin" /> : <ArrowDownToLine className="h-4 w-4" />}
                  导出文件
                </button>
              </div>
            </div>
          </aside>
        </div>
      </div>
    </main>
  );
}
