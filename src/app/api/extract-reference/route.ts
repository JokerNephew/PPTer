import { NextResponse } from "next/server";
import { extractReferenceText } from "@/lib/ppt/referenceExtractor";
import type { ApiErrorBody } from "@/types/ppt";

export const runtime = "nodejs";

export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    const file = formData.get("reference");

    if (!(file instanceof File)) {
      return NextResponse.json<ApiErrorBody>({ error: "请上传参考资料文件。" }, { status: 400 });
    }

    const text = await extractReferenceText(file);
    return NextResponse.json({
      file: {
        name: file.name,
        size: file.size,
        type: file.type || "unknown",
      },
      text,
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "参考资料读取失败。";
    return NextResponse.json<ApiErrorBody>({ error: message }, { status: 400 });
  }
}
