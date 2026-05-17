import { NextResponse } from "next/server";
import { searchExternalTemplatesPage } from "@/lib/ppt/externalTemplates";
import type { ApiErrorBody, TemplateSearchResponse } from "@/types/ppt";

export const runtime = "nodejs";
const DEFAULT_PAGE_SIZE = 12;
const MAX_PAGE_SIZE = 24;

function readPositiveInt(value: string | null, fallback: number) {
  const parsed = Number.parseInt(value ?? "", 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
}

export async function GET(request: Request) {
  const { searchParams } = new URL(request.url);
  const query = searchParams.get("q")?.trim() ?? "";
  const page = readPositiveInt(searchParams.get("page"), 1);
  const pageSize = Math.min(MAX_PAGE_SIZE, readPositiveInt(searchParams.get("pageSize"), DEFAULT_PAGE_SIZE));

  if (query.length > 80) {
    return NextResponse.json<ApiErrorBody>({ error: "搜索词不能超过 80 个字符。" }, { status: 400 });
  }

  try {
    const response: TemplateSearchResponse = await searchExternalTemplatesPage(query, page, pageSize);
    return NextResponse.json(response);
  } catch (error) {
    const message = error instanceof Error ? error.message : "模板搜索失败。";
    return NextResponse.json<ApiErrorBody>({ error: message }, { status: 500 });
  }
}
