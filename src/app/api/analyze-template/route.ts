import { NextResponse } from "next/server";
import { parseTemplateFile } from "@/lib/ppt/templateParser";
import type { ApiErrorBody } from "@/types/ppt";

export const runtime = "nodejs";

export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    const file = formData.get("template");

    if (!(file instanceof File)) {
      return NextResponse.json<ApiErrorBody>({ error: "请上传 .pptx 模板文件。" }, { status: 400 });
    }

    const template = await parseTemplateFile(file);
    return NextResponse.json({ template });
  } catch (error) {
    const message = error instanceof Error ? error.message : "模板解析失败。";
    return NextResponse.json<ApiErrorBody>({ error: message }, { status: 400 });
  }
}
