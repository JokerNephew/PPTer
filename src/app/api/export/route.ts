import { NextResponse } from "next/server";
import { exportPdf } from "@/lib/ppt/pdfExporter";
import { exportPptx } from "@/lib/ppt/pptxExporter";
import type { ApiErrorBody, DeckExportPayload } from "@/types/ppt";

export const runtime = "nodejs";

function safeFileName(fileName: string, format: "pptx" | "pdf") {
  const base = (fileName || "ai-ppt-deck")
    .replace(/[\\/:*?"<>|]/g, "-")
    .replace(/\.(pptx|pdf)$/i, "")
    .slice(0, 80);
  return `${base || "ai-ppt-deck"}.${format}`;
}

export async function POST(request: Request) {
  try {
    const payload = (await request.json()) as DeckExportPayload;
    if (!payload.slides?.length) {
      return NextResponse.json<ApiErrorBody>({ error: "没有可导出的页面内容。" }, { status: 400 });
    }

    const fileName = safeFileName(payload.fileName, payload.format);
    const buffer = payload.format === "pdf" ? await exportPdf(payload) : await exportPptx(payload);

    return new Response(new Uint8Array(buffer), {
      status: 200,
      headers: {
        "Content-Type":
          payload.format === "pdf"
            ? "application/pdf"
            : "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "Content-Disposition": `attachment; filename="${encodeURIComponent(fileName)}"`,
      },
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "导出失败。";
    return NextResponse.json<ApiErrorBody>({ error: message }, { status: 500 });
  }
}
