import { NextResponse } from "next/server";
import { generateSlides, MissingApiKeyError, OpenAIRequestError } from "@/lib/ppt/openai";
import { applyTemplateStyleToSlides, createSlideGenerationContext } from "@/lib/ppt/slideGenerator";
import type { ApiErrorBody, DeckRequest } from "@/types/ppt";

export const runtime = "nodejs";

export async function POST(request: Request) {
  try {
    const payload = (await request.json()) as DeckRequest;
    if (!payload.outline?.length) {
      return NextResponse.json<ApiErrorBody>({ error: "请先确认或编辑 PPT 大纲。" }, { status: 400 });
    }

    const context = createSlideGenerationContext(payload);
    const slides = applyTemplateStyleToSlides(await generateSlides(payload), context.template);
    return NextResponse.json({ slides, model: process.env.OPENAI_MODEL || "gpt-5.5" });
  } catch (error) {
    const message = error instanceof Error ? error.message : "生成页面内容失败。";
    const status = error instanceof MissingApiKeyError ? 500 : error instanceof OpenAIRequestError ? error.status : 502;
    return NextResponse.json<ApiErrorBody>({ error: message }, { status });
  }
}
