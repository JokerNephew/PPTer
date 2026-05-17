import { NextResponse } from "next/server";
import { searchExternalTemplates } from "@/lib/ppt/externalTemplates";
import type { ApiErrorBody, TemplateSearchResponse } from "@/types/ppt";

export const runtime = "nodejs";

export async function GET(request: Request) {
  const { searchParams } = new URL(request.url);
  const query = searchParams.get("q")?.trim() ?? "";

  if (query.length > 80) {
    return NextResponse.json<ApiErrorBody>({ error: "搜索词不能超过 80 个字符。" }, { status: 400 });
  }

  try {
    const response: TemplateSearchResponse = await searchExternalTemplates(query);
    return NextResponse.json(response);
  } catch (error) {
    const message = error instanceof Error ? error.message : "模板搜索失败。";
    return NextResponse.json<ApiErrorBody>({ error: message }, { status: 500 });
  }
}
