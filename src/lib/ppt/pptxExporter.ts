import pptxgen from "pptxgenjs";
import type { DeckExportPayload, SlideContent } from "@/types/ppt";
import { type ExportTheme, extractTheme } from "./themeExtractor";

interface SlideGeometry {
  width: number;
  height: number;
  margin: number;
}

const WHITE = "FFFFFF";
const SLATE_50 = "F8FAFC";
const SLATE_100 = "F1F5F9";
const SLATE_200 = "E2E8F0";
const SLATE_500 = "64748B";
const SLATE_600 = "475569";
const SLATE_900 = "0F172A";

function cleanColor(color: string | undefined, fallback = SLATE_900) {
  const value = (color ?? fallback).replace("#", "").trim().toUpperCase();
  return /^[0-9A-F]{6}$/.test(value) ? value : fallback;
}

function geometryFor(theme: ExportTheme): SlideGeometry {
  return {
    width: theme.aspectRatio === "LAYOUT_4X3" ? 10 : 13.333,
    height: 7.5,
    margin: 0.58,
  };
}

function compactText(text: string, max = 90) {
  const cleaned = text.replace(/\s+/g, " ").trim();
  return cleaned.length > max ? `${cleaned.slice(0, max - 1)}…` : cleaned;
}

function labelFromPoint(point: string, max = 22) {
  return compactText(point.split(/[，。；;:：,.]/)[0] || point, max);
}

function fontSizeFor(text: string, base: number, min: number) {
  const length = Array.from(text).length;
  if (length > 48) return Math.max(min, base - 4);
  if (length > 34) return Math.max(min, base - 3);
  if (length > 24) return Math.max(min, base - 2);
  return base;
}

function addFooter(
  slide: pptxgen.Slide,
  geometry: SlideGeometry,
  slideNumber: number,
  total: number,
  themeName: string,
  isDark = false,
) {
  slide.addText(`${themeName} · ${slideNumber}/${total}`, {
    x: geometry.margin,
    y: geometry.height - 0.36,
    w: geometry.width - geometry.margin * 2,
    h: 0.16,
    fontSize: 7,
    color: isDark ? "CBD5E1" : SLATE_500,
    margin: 0,
  });
}

function addKicker(
  slide: pptxgen.Slide,
  pptx: pptxgen,
  text: string,
  x: number,
  y: number,
  accent: string,
  color = SLATE_500,
) {
  slide.addShape(pptx.ShapeType.ellipse, {
    x,
    y: y + 0.03,
    w: 0.09,
    h: 0.09,
    fill: { color: accent },
    line: { color: accent },
  });
  slide.addText(text, {
    x: x + 0.17,
    y,
    w: 4.8,
    h: 0.17,
    fontSize: 7.5,
    bold: true,
    color,
    charSpacing: 1.1,
    margin: 0,
  });
}

function addBodyPoints(
  slide: pptxgen.Slide,
  pptx: pptxgen,
  slideContent: SlideContent,
  theme: ExportTheme,
  x: number,
  y: number,
  w: number,
  maxItems = 4,
) {
  const accent = cleanColor(theme.colors.accent);
  slideContent.body.slice(0, maxItems).forEach((point, index) => {
    const itemY = y + index * 0.76;
    slide.addShape(pptx.ShapeType.ellipse, {
      x,
      y: itemY + 0.09,
      w: 0.14,
      h: 0.14,
      fill: { color: accent },
      line: { color: accent },
    });
    slide.addText(compactText(point, 54), {
      x: x + 0.3,
      y: itemY,
      w,
      h: 0.42,
      fontFace: theme.fonts.body,
      fontSize: fontSizeFor(point, 13, 10.5),
      color: cleanColor(theme.colors.text),
      breakLine: false,
      margin: 0,
      valign: "middle",
    });
  });
}

function addInsightBars(
  slide: pptxgen.Slide,
  pptx: pptxgen,
  points: string[],
  theme: ExportTheme,
  x: number,
  y: number,
  w: number,
) {
  const accent = cleanColor(theme.colors.accent);
  const alt = cleanColor(theme.colors.accentAlt);
  slide.addText("重点分布", {
    x,
    y,
    w,
    h: 0.24,
    fontFace: theme.fonts.heading,
    fontSize: 12,
    bold: true,
    color: alt,
    margin: 0,
  });

  points.slice(0, 4).forEach((point, index) => {
    const barY = y + 0.52 + index * 0.52;
    const widthRatio = [0.9, 0.74, 0.62, 0.5][index] ?? 0.48;
    slide.addText(labelFromPoint(point, 18), {
      x,
      y: barY - 0.02,
      w: Math.min(2.05, w * 0.42),
      h: 0.2,
      fontFace: theme.fonts.body,
      fontSize: 8.6,
      color: SLATE_600,
      margin: 0,
    });
    slide.addShape(pptx.ShapeType.roundRect, {
      x: x + Math.min(2.2, w * 0.45),
      y: barY,
      w: w - Math.min(2.2, w * 0.45),
      h: 0.16,
      fill: { color: SLATE_200 },
      line: { color: SLATE_200, transparency: 100 },
    });
    slide.addShape(pptx.ShapeType.roundRect, {
      x: x + Math.min(2.2, w * 0.45),
      y: barY,
      w: (w - Math.min(2.2, w * 0.45)) * widthRatio,
      h: 0.16,
      fill: { color: index === 0 ? accent : alt, transparency: index === 0 ? 0 : 22 },
      line: { color: index === 0 ? accent : alt, transparency: 100 },
    });
  });
}

function addAgendaVisual(
  slide: pptxgen.Slide,
  pptx: pptxgen,
  slideContent: SlideContent,
  theme: ExportTheme,
  x: number,
  y: number,
  w: number,
  h: number,
) {
  const accent = cleanColor(theme.colors.accent);
  const points = slideContent.body.length ? slideContent.body : ["背景判断", "关键洞察", "行动路径"];
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    fill: { color: WHITE },
    line: { color: SLATE_200 },
  });
  slide.addText("汇报路径", {
    x: x + 0.36,
    y: y + 0.32,
    w: w - 0.72,
    h: 0.26,
    fontFace: theme.fonts.heading,
    fontSize: 13,
    bold: true,
    color: cleanColor(theme.colors.accentAlt),
    margin: 0,
  });
  slide.addShape(pptx.ShapeType.line, {
    x: x + 0.52,
    y: y + 0.95,
    w: 0,
    h: Math.max(2.4, h - 1.45),
    line: { color: accent, width: 1.4, transparency: 8 },
  });

  points.slice(0, 5).forEach((point, index) => {
    const itemY = y + 0.86 + index * Math.min(0.72, (h - 1.35) / Math.max(points.length, 1));
    slide.addShape(pptx.ShapeType.ellipse, {
      x: x + 0.39,
      y: itemY,
      w: 0.26,
      h: 0.26,
      fill: { color: index === 0 ? accent : WHITE },
      line: { color: accent, width: 1.2 },
    });
    slide.addText(`0${index + 1}`, {
      x: x + 0.86,
      y: itemY - 0.01,
      w: 0.36,
      h: 0.18,
      fontFace: theme.fonts.heading,
      fontSize: 8,
      bold: true,
      color: accent,
      margin: 0,
    });
    slide.addText(compactText(point, 32), {
      x: x + 1.25,
      y: itemY - 0.05,
      w: w - 1.62,
      h: 0.28,
      fontFace: theme.fonts.body,
      fontSize: 10.8,
      color: SLATE_900,
      margin: 0,
    });
  });
}

function addDataVisual(
  slide: pptxgen.Slide,
  pptx: pptxgen,
  slideContent: SlideContent,
  theme: ExportTheme,
  x: number,
  y: number,
  w: number,
  h: number,
) {
  const accent = cleanColor(theme.colors.accent);
  const alt = cleanColor(theme.colors.accentAlt);
  const points = slideContent.body.length ? slideContent.body : ["关键指标", "变化方向", "行动优先级"];
  const cardW = (w - 0.36) / 3;

  points.slice(0, 3).forEach((point, index) => {
    const cardX = x + index * (cardW + 0.18);
    slide.addShape(pptx.ShapeType.roundRect, {
      x: cardX,
      y,
      w: cardW,
      h: 1.14,
      fill: { color: index === 0 ? accent : WHITE, transparency: 0 },
      line: { color: index === 0 ? accent : SLATE_200 },
    });
    slide.addText(`0${index + 1}`, {
      x: cardX + 0.18,
      y: y + 0.18,
      w: 0.58,
      h: 0.28,
      fontFace: theme.fonts.heading,
      fontSize: 18,
      bold: true,
      color: index === 0 ? WHITE : alt,
      margin: 0,
    });
    slide.addText(labelFromPoint(point, 24), {
      x: cardX + 0.18,
      y: y + 0.68,
      w: cardW - 0.36,
      h: 0.26,
      fontFace: theme.fonts.body,
      fontSize: 8.6,
      color: index === 0 ? WHITE : SLATE_600,
      margin: 0,
    });
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y: y + 1.52,
    w,
    h: h - 1.52,
    fill: { color: WHITE },
    line: { color: SLATE_200 },
  });
  slide.addText(compactText(slideContent.chartSuggestion, 68), {
    x: x + 0.34,
    y: y + 1.84,
    w: w - 0.68,
    h: 0.32,
    fontFace: theme.fonts.heading,
    fontSize: 11,
    bold: true,
    color: alt,
    margin: 0,
  });
  addInsightBars(slide, pptx, points, theme, x + 0.34, y + 2.38, w - 0.68);
}

function addContentVisual(
  slide: pptxgen.Slide,
  pptx: pptxgen,
  slideContent: SlideContent,
  theme: ExportTheme,
  x: number,
  y: number,
  w: number,
  h: number,
) {
  const accent = cleanColor(theme.colors.accent);
  const alt = cleanColor(theme.colors.accentAlt);
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    fill: { color: WHITE },
    line: { color: SLATE_200 },
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: x + 0.35,
    y: y + 0.34,
    w: w - 0.7,
    h: 1.25,
    fill: { color: accent, transparency: 7 },
    line: { color: accent, transparency: 100 },
  });
  slide.addText("视觉焦点", {
    x: x + 0.65,
    y: y + 0.58,
    w: 1.45,
    h: 0.2,
    fontFace: theme.fonts.heading,
    fontSize: 10,
    bold: true,
    color: WHITE,
    margin: 0,
  });
  slide.addText(compactText(slideContent.imageSuggestion, 72), {
    x: x + 0.65,
    y: y + 0.94,
    w: w - 1.3,
    h: 0.35,
    fontFace: theme.fonts.body,
    fontSize: 9.6,
    color: WHITE,
    margin: 0,
  });

  const cardY = y + 1.92;
  const cardH = Math.max(1.0, (h - 2.25) / 2);
  [
    ["图表呈现", slideContent.chartSuggestion],
    ["版式结构", slideContent.layoutSuggestion],
  ].forEach(([label, value], index) => {
    const itemY = cardY + index * (cardH + 0.22);
    slide.addShape(pptx.ShapeType.roundRect, {
      x: x + 0.35,
      y: itemY,
      w: w - 0.7,
      h: cardH,
      fill: { color: index === 0 ? SLATE_50 : SLATE_100 },
      line: { color: SLATE_200 },
    });
    slide.addShape(pptx.ShapeType.rect, {
      x: x + 0.35,
      y: itemY,
      w: 0.09,
      h: cardH,
      fill: { color: index === 0 ? accent : alt },
      line: { color: index === 0 ? accent : alt },
    });
    slide.addText(label, {
      x: x + 0.65,
      y: itemY + 0.2,
      w: 1.6,
      h: 0.2,
      fontFace: theme.fonts.heading,
      fontSize: 9,
      bold: true,
      color: alt,
      margin: 0,
    });
    slide.addText(compactText(value, 78), {
      x: x + 0.65,
      y: itemY + 0.54,
      w: w - 1.3,
      h: 0.36,
      fontFace: theme.fonts.body,
      fontSize: 9,
      color: SLATE_600,
      margin: 0,
    });
  });
}

function addClosingActions(
  slide: pptxgen.Slide,
  pptx: pptxgen,
  slideContent: SlideContent,
  theme: ExportTheme,
  x: number,
  y: number,
  w: number,
  h: number,
) {
  const accent = cleanColor(theme.colors.accent);
  const actions = slideContent.body.length ? slideContent.body : ["复盘关键判断", "确认资源与责任人", "推进下一步行动"];
  actions.slice(0, 3).forEach((action, index) => {
    const itemY = y + index * ((h - 0.4) / 3);
    slide.addShape(pptx.ShapeType.roundRect, {
      x,
      y: itemY,
      w,
      h: 1.28,
      fill: { color: WHITE, transparency: 4 },
      line: { color: WHITE, transparency: 70 },
    });
    slide.addText(`0${index + 1}`, {
      x: x + 0.28,
      y: itemY + 0.25,
      w: 0.55,
      h: 0.28,
      fontFace: theme.fonts.heading,
      fontSize: 18,
      bold: true,
      color: accent,
      margin: 0,
    });
    slide.addText(compactText(action, 36), {
      x: x + 1.0,
      y: itemY + 0.31,
      w: w - 1.28,
      h: 0.36,
      fontFace: theme.fonts.body,
      fontSize: 10.8,
      color: WHITE,
      margin: 0,
    });
  });
}

function addCoverSlide(pptx: pptxgen, slideContent: SlideContent, payload: DeckExportPayload, theme: ExportTheme) {
  const geometry = geometryFor(theme);
  const slide = pptx.addSlide();
  const accent = cleanColor(theme.colors.accent);
  const dark = cleanColor(theme.colors.accentAlt);
  slide.background = { color: dark };
  slide.addShape(pptx.ShapeType.rect, {
    x: geometry.width * 0.66,
    y: 0,
    w: geometry.width * 0.34,
    h: geometry.height,
    fill: { color: accent, transparency: 3 },
    line: { color: accent },
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: geometry.width * 0.71,
    y: 1.05,
    w: geometry.width * 0.22,
    h: 4.98,
    fill: { color: WHITE, transparency: 88 },
    line: { color: WHITE, transparency: 76 },
  });
  slide.addShape(pptx.ShapeType.line, {
    x: geometry.width * 0.74,
    y: 1.55,
    w: geometry.width * 0.14,
    h: 0,
    line: { color: WHITE, width: 1.5, transparency: 12 },
  });
  slide.addShape(pptx.ShapeType.line, {
    x: geometry.width * 0.74,
    y: 5.48,
    w: geometry.width * 0.14,
    h: 0,
    line: { color: WHITE, width: 1.5, transparency: 12 },
  });
  addKicker(slide, pptx, "AI PPT STUDIO", geometry.margin, 0.72, accent, "CBD5E1");
  slide.addText(compactText(slideContent.title, 42), {
    x: geometry.margin,
    y: 1.42,
    w: geometry.width * 0.58,
    h: 1.5,
    fontFace: theme.fonts.heading,
    fontSize: fontSizeFor(slideContent.title, 37, 28),
    bold: true,
    color: WHITE,
    margin: 0,
    breakLine: false,
  });
  slide.addText(slideContent.body.slice(0, 2).join("\n"), {
    x: geometry.margin + 0.02,
    y: 3.25,
    w: geometry.width * 0.5,
    h: 0.72,
    fontFace: theme.fonts.body,
    fontSize: 14,
    color: "E2E8F0",
    breakLine: false,
    margin: 0,
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: geometry.margin,
    y: 5.52,
    w: Math.min(4.8, geometry.width * 0.46),
    h: 0.38,
    fill: { color: accent, transparency: 0 },
    line: { color: accent },
  });
  slide.addText(compactText(payload.request.purpose || "AI generated deck", 42), {
    x: geometry.margin + 0.18,
    y: 5.64,
    w: Math.min(4.45, geometry.width * 0.42),
    h: 0.12,
    fontFace: theme.fonts.body,
    fontSize: 8.5,
    bold: true,
    color: WHITE,
    margin: 0,
  });
  addFooter(slide, geometry, slideContent.slideNumber, payload.slides.length, theme.name, true);
  slide.addNotes(slideContent.speakerNotes);
}

function addContentSlide(pptx: pptxgen, slideContent: SlideContent, payload: DeckExportPayload, theme: ExportTheme) {
  const geometry = geometryFor(theme);
  const slide = pptx.addSlide();
  const accent = cleanColor(theme.colors.accent);
  const alt = cleanColor(theme.colors.accentAlt);
  slide.background = { color: cleanColor(theme.colors.background, SLATE_50) };

  addKicker(slide, pptx, `${theme.name.toUpperCase()} · ${slideContent.templateRole.toUpperCase()}`, geometry.margin, 0.4, accent);
  slide.addText(compactText(slideContent.title, 42), {
    x: geometry.margin,
    y: 0.68,
    w: geometry.width - geometry.margin * 2,
    h: 0.45,
    fontFace: theme.fonts.heading,
    fontSize: fontSizeFor(slideContent.title, 25, 19),
    bold: true,
    color: cleanColor(theme.colors.text),
    margin: 0,
  });

  const leftX = geometry.margin;
  const leftW = geometry.width < 11 ? 4.12 : 5.42;
  const gap = 0.42;
  const rightX = leftX + leftW + gap;
  const rightW = geometry.width - rightX - geometry.margin;
  const panelY = 1.38;
  const panelH = 5.35;

  slide.addShape(pptx.ShapeType.roundRect, {
    x: leftX,
    y: panelY,
    w: leftW,
    h: panelH,
    fill: { color: cleanColor(theme.colors.surface, WHITE), transparency: 0 },
    line: { color: SLATE_200, transparency: 5 },
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: leftX,
    y: panelY,
    w: 0.1,
    h: panelH,
    fill: { color: accent },
    line: { color: accent },
  });
  slide.addText("核心要点", {
    x: leftX + 0.35,
    y: panelY + 0.34,
    w: leftW - 0.7,
    h: 0.24,
    fontFace: theme.fonts.heading,
    fontSize: 12,
    bold: true,
    color: alt,
    margin: 0,
  });
  addBodyPoints(slide, pptx, slideContent, theme, leftX + 0.38, panelY + 0.93, leftW - 0.78);
  slide.addText(compactText(slideContent.speakerNotes, 100), {
    x: leftX + 0.35,
    y: panelY + panelH - 0.72,
    w: leftW - 0.7,
    h: 0.32,
    fontFace: theme.fonts.body,
    fontSize: 8.3,
    italic: true,
    color: SLATE_500,
    margin: 0,
  });

  if (slideContent.templateRole === "agenda") {
    addAgendaVisual(slide, pptx, slideContent, theme, rightX, panelY, rightW, panelH);
  } else if (slideContent.templateRole === "data") {
    addDataVisual(slide, pptx, slideContent, theme, rightX, panelY, rightW, panelH);
  } else {
    addContentVisual(slide, pptx, slideContent, theme, rightX, panelY, rightW, panelH);
  }

  addFooter(slide, geometry, slideContent.slideNumber, payload.slides.length, theme.name);
  slide.addNotes(slideContent.speakerNotes);
}

function addClosingSlide(pptx: pptxgen, slideContent: SlideContent, payload: DeckExportPayload, theme: ExportTheme) {
  const geometry = geometryFor(theme);
  const slide = pptx.addSlide();
  const dark = cleanColor(theme.colors.accentAlt);
  const accent = cleanColor(theme.colors.accent);
  slide.background = { color: dark };
  addKicker(slide, pptx, "SUMMARY · NEXT STEPS", geometry.margin, 0.78, accent, "CBD5E1");
  slide.addText(compactText(slideContent.title, 38), {
    x: geometry.margin,
    y: 1.22,
    w: geometry.width * 0.58,
    h: 0.74,
    fontFace: theme.fonts.heading,
    fontSize: fontSizeFor(slideContent.title, 30, 23),
    bold: true,
    color: WHITE,
    margin: 0,
  });
  slide.addText(compactText(slideContent.layoutSuggestion, 100), {
    x: geometry.margin + 0.02,
    y: 2.28,
    w: geometry.width * 0.48,
    h: 0.54,
    fontFace: theme.fonts.body,
    fontSize: 12,
    color: "CBD5E1",
    margin: 0,
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: geometry.margin,
    y: 5.2,
    w: Math.min(4.6, geometry.width * 0.44),
    h: 0.42,
    fill: { color: accent },
    line: { color: accent },
  });
  slide.addText("Review · Refine · Export", {
    x: geometry.margin + 0.22,
    y: 5.33,
    w: Math.min(4.1, geometry.width * 0.4),
    h: 0.14,
    fontFace: theme.fonts.body,
    fontSize: 8.8,
    bold: true,
    color: WHITE,
    margin: 0,
  });

  const actionsX = geometry.width < 11 ? 5.35 : 8.05;
  addClosingActions(slide, pptx, slideContent, theme, actionsX, 1.18, geometry.width - actionsX - geometry.margin, 4.85);
  addFooter(slide, geometry, slideContent.slideNumber, payload.slides.length, theme.name, true);
  slide.addNotes(slideContent.speakerNotes);
}

export async function exportPptx(payload: DeckExportPayload): Promise<Buffer> {
  const pptx = new pptxgen();
  const theme = extractTheme(payload.templateProfile, payload.request.templateId);
  pptx.layout = theme.aspectRatio;
  pptx.author = "AI PPT Studio";
  pptx.subject = payload.request.topic;
  pptx.title = payload.fileName;
  pptx.company = "AI PPT Studio";
  pptx.theme = {
    headFontFace: theme.fonts.heading,
    bodyFontFace: theme.fonts.body,
  };

  payload.slides.forEach((slideContent, index) => {
    if (index === 0 || slideContent.templateRole === "cover") {
      addCoverSlide(pptx, slideContent, payload, theme);
    } else if (index === payload.slides.length - 1 || slideContent.templateRole === "closing") {
      addClosingSlide(pptx, slideContent, payload, theme);
    } else {
      addContentSlide(pptx, slideContent, payload, theme);
    }
  });

  const arrayBuffer = (await pptx.write({ outputType: "arraybuffer" })) as ArrayBuffer;
  return Buffer.from(new Uint8Array(arrayBuffer));
}
