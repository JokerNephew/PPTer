import JSZip from "jszip";
import type { TemplateProfile } from "@/types/ppt";

const MAX_TEMPLATE_SIZE = 25 * 1024 * 1024;

function inferAspectRatio(width?: number, height?: number): TemplateProfile["aspectRatio"] {
  if (!width || !height) return "unknown";
  const ratio = width / height;
  if (Math.abs(ratio - 16 / 9) < 0.08) return "16:9";
  if (Math.abs(ratio - 4 / 3) < 0.08) return "4:3";
  if (ratio > 1.7) return "wide";
  return "unknown";
}

function parseXmlNumber(value: string | undefined) {
  return value ? Number.parseInt(value, 10) : undefined;
}

function extractThemeColors(themeXml: string) {
  const colors = new Set<string>();
  const colorMatches = themeXml.matchAll(/<a:srgbClr val="([0-9A-Fa-f]{6})"/g);
  for (const match of colorMatches) colors.add(`#${match[1].toUpperCase()}`);
  return Array.from(colors).slice(0, 8);
}

function extractFonts(themeXml: string) {
  const fonts = new Set<string>();
  const latinMatches = themeXml.matchAll(/<a:latin typeface="([^"]+)"/g);
  for (const match of latinMatches) {
    if (match[1] && match[1] !== "+mj-lt" && match[1] !== "+mn-lt") fonts.add(match[1]);
  }
  return Array.from(fonts).slice(0, 5);
}

function createPreview(colors: string[]) {
  const [first = "#202124", second = "#FFFFFF", accent = "#00A6A6"] = colors;
  return `linear-gradient(135deg, ${first} 0 48%, ${accent} 48% 55%, ${second} 55% 100%)`;
}

export async function parseTemplateFile(file: File): Promise<TemplateProfile> {
  if (!file.name.toLowerCase().endsWith(".pptx")) {
    throw new Error("仅支持上传 .pptx 模板文件。");
  }

  if (file.size > MAX_TEMPLATE_SIZE) {
    throw new Error("模板文件超过 25MB，请压缩后再上传。");
  }

  const arrayBuffer = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(arrayBuffer);
  const presentationXml = await zip.file("ppt/presentation.xml")?.async("text");
  const themeXml =
    (await zip.file("ppt/theme/theme1.xml")?.async("text")) ??
    (await zip.file("ppt/theme/theme2.xml")?.async("text")) ??
    "";

  if (!presentationXml) {
    throw new Error("无法读取 PPTX 的 presentation.xml，文件可能已损坏或不是有效模板。");
  }

  const slideCount = zip.file(/^ppt\/slides\/slide\d+\.xml$/).length;
  const sizeMatch = presentationXml.match(/<p:sldSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"/);
  const width = parseXmlNumber(sizeMatch?.[1]);
  const height = parseXmlNumber(sizeMatch?.[2]);
  const aspectRatio = inferAspectRatio(width, height);
  const dominantColors = extractThemeColors(themeXml);
  const fonts = extractFonts(themeXml);

  const slideXml = await Promise.all(
    zip
      .file(/^ppt\/slides\/slide\d+\.xml$/)
      .slice(0, 8)
      .map((entry) => entry.async("text")),
  );
  const combinedSlides = slideXml.join("\n").toLowerCase();
  const detectedLayouts = [
    slideCount > 0 ? "cover" : "",
    combinedSlides.includes("title") || combinedSlides.includes("ctrtitle") ? "title" : "",
    combinedSlides.includes("tbl") ? "table" : "",
    combinedSlides.includes("chart") || combinedSlides.includes("graphicframe") ? "chart/data" : "",
    combinedSlides.includes("pic") ? "image" : "",
    "content",
  ].filter(Boolean);

  const warnings: string[] = [];
  if (dominantColors.length === 0) warnings.push("未能从主题 XML 提取主色，导出时会使用安全默认配色。");
  if (fonts.length === 0) warnings.push("未能识别模板字体，导出时会使用 Aptos/Calibri fallback。");
  if (slideCount === 0) warnings.push("未检测到模板页面，将只复用主题色和比例信息。");

  return {
    id: `uploaded-${Date.now()}`,
    name: file.name,
    source: "uploaded",
    pageCount: slideCount,
    aspectRatio,
    dominantColors: dominantColors.length ? dominantColors : ["#202124", "#FFFFFF", "#00A6A6"],
    fonts: fonts.length ? fonts : ["Aptos", "Calibri"],
    titleHierarchy: ["cover/title placeholder", "section heading", "body placeholder"],
    detectedLayouts: Array.from(new Set(detectedLayouts)),
    hasCover: slideCount >= 1,
    hasAgenda: slideCount >= 3,
    hasClosing: slideCount >= 4,
    thumbnail: createPreview(dominantColors),
    warnings,
    fallbackStrategy:
      "Template reuse module extracted page ratio, theme colors, fonts, and available slide roles. Full XML layout cloning can be attached behind templateParser/themeExtractor later.",
  };
}
