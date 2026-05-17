import { PDFDocument, StandardFonts, rgb } from "pdf-lib";
import type { DeckExportPayload } from "@/types/ppt";
import { extractTheme } from "./themeExtractor";

function hexToRgb(hex: string) {
  const normalized = hex.replace("#", "");
  const value = normalized.length === 6 ? normalized : "111827";
  return rgb(
    Number.parseInt(value.slice(0, 2), 16) / 255,
    Number.parseInt(value.slice(2, 4), 16) / 255,
    Number.parseInt(value.slice(4, 6), 16) / 255,
  );
}

function wrapText(text: string, max = 72) {
  const chunks: string[] = [];
  let current = "";
  for (const word of text.split(/\s+/)) {
    if ((current + word).length > max) {
      chunks.push(current.trim());
      current = word;
    } else {
      current += ` ${word}`;
    }
  }
  if (current.trim()) chunks.push(current.trim());
  return chunks;
}

export async function exportPdf(payload: DeckExportPayload): Promise<Buffer> {
  const pdf = await PDFDocument.create();
  const headingFont = await pdf.embedFont(StandardFonts.HelveticaBold);
  const bodyFont = await pdf.embedFont(StandardFonts.Helvetica);
  const theme = extractTheme(payload.templateProfile, payload.request.templateId);
  const pageSize: [number, number] = theme.aspectRatio === "LAYOUT_4X3" ? [720, 540] : [960, 540];

  payload.slides.forEach((slide, index) => {
    const page = pdf.addPage(pageSize);
    const { width, height } = page.getSize();
    const isCover = index === 0 || slide.templateRole === "cover";
    const background = isCover ? theme.colors.accentAlt : theme.colors.background;

    page.drawRectangle({ x: 0, y: 0, width, height, color: hexToRgb(background) });
    page.drawText(slide.title.slice(0, 92), {
      x: 42,
      y: height - 92,
      size: isCover ? 34 : 24,
      font: headingFont,
      color: hexToRgb(isCover ? "#FFFFFF" : theme.colors.text),
      maxWidth: width - 84,
    });

    let y = height - (isCover ? 158 : 142);
    slide.body.forEach((point) => {
      wrapText(`• ${point}`, 80).forEach((line) => {
        page.drawText(line, {
          x: 54,
          y,
          size: 13,
          font: bodyFont,
          color: hexToRgb(isCover ? "#E2E8F0" : theme.colors.text),
          maxWidth: width - 108,
        });
        y -= 22;
      });
      y -= 8;
    });

    if (!isCover) {
      page.drawRectangle({
        x: 54,
        y: 62,
        width: width - 108,
        height: 92,
        color: hexToRgb("#FFFFFF"),
        borderColor: hexToRgb(theme.colors.accent),
        borderWidth: 1.5,
      });
      page.drawText(`Chart: ${slide.chartSuggestion}`.slice(0, 135), {
        x: 76,
        y: 118,
        size: 10,
        font: bodyFont,
        color: hexToRgb("#334155"),
        maxWidth: width - 152,
      });
      page.drawText(`Layout: ${slide.layoutSuggestion}`.slice(0, 135), {
        x: 76,
        y: 92,
        size: 10,
        font: bodyFont,
        color: hexToRgb("#334155"),
        maxWidth: width - 152,
      });
    }

    page.drawText(`${payload.fileName} · ${index + 1}/${payload.slides.length}`, {
      x: 42,
      y: 24,
      size: 8,
      font: bodyFont,
      color: hexToRgb(isCover ? "#CBD5E1" : "#64748B"),
    });
  });

  const bytes = await pdf.save();
  return Buffer.from(bytes);
}
