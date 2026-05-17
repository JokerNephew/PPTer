export async function extractReferenceText(file: File): Promise<string> {
  const maxSize = 15 * 1024 * 1024;
  if (file.size > maxSize) {
    throw new Error("参考资料超过 15MB，请拆分或压缩后上传。");
  }

  const extension = file.name.split(".").pop()?.toLowerCase();

  if (extension === "txt" || extension === "md") {
    return file.text();
  }

  if (extension === "docx" || extension === "pptx") {
    return `${file.name} 已上传。当前版本保留文件名、大小和格式元数据，后续可接入 markitdown 或 Office XML 解析以抽取全文。`;
  }

  if (extension === "pdf") {
    return `${file.name} 已上传。当前版本保留 PDF 元数据，后续可接入 pdf-parse 或 OCR 服务提取全文。`;
  }

  throw new Error("参考资料仅支持 Word、PDF、TXT、Markdown 或 PPTX 文件。");
}
