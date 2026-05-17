import { NextResponse } from "next/server";
import { generateOutline, MissingApiKeyError, OpenAIRequestError } from "@/lib/ppt/openai";
import type { ApiErrorBody, DeckRequest } from "@/types/ppt";

export const runtime = "nodejs";

export async function POST(request: Request) {
  try {
    const payload = (await request.json()) as DeckRequest;
    if (!payload.topic?.trim()) {
      return NextResponse.json<ApiErrorBody>({ error: "请输入 PPT 主题。" }, { status: 400 });
    }

    const outline = await generateOutline(payload);
    return NextResponse.json({ outline, model: process.env.OPENAI_MODEL || "gpt-5.5" });
  } catch (error) {
    const message = error instanceof Error ? error.message : "生成大纲失败。";
    const status = error instanceof MissingApiKeyError ? 500 : error instanceof OpenAIRequestError ? error.status : 502;
    return NextResponse.json<ApiErrorBody>({ error: message }, { status });
  }
}
