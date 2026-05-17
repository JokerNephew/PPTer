import { NextResponse } from "next/server";
import { MissingApiKeyError, OpenAIRequestError, rewriteSlide } from "@/lib/ppt/openai";
import type { ApiErrorBody, DeckRequest, SlideContent } from "@/types/ppt";

export const runtime = "nodejs";

interface RewritePayload {
  slide: SlideContent;
  mode: "regenerate" | "polish" | "shorten" | "expand";
  request: DeckRequest;
}

export async function POST(request: Request) {
  try {
    const payload = (await request.json()) as RewritePayload;
    if (!payload.slide || !payload.mode) {
      return NextResponse.json<ApiErrorBody>({ error: "缺少当前页或改写模式。" }, { status: 400 });
    }

    const slide = await rewriteSlide(payload.slide, payload.mode, payload.request);
    return NextResponse.json({ slide, model: process.env.OPENAI_MODEL || "gpt-5.5" });
  } catch (error) {
    const message = error instanceof Error ? error.message : "单页改写失败。";
    const status = error instanceof MissingApiKeyError ? 500 : error instanceof OpenAIRequestError ? error.status : 502;
    return NextResponse.json<ApiErrorBody>({ error: message }, { status });
  }
}
