# AI PPT Studio

一个用于 AI 生成 PPT 的 Next.js 工作台。首屏就是创建和编辑界面，支持需求输入、参考资料、上传 PPTX 模板、生成大纲、生成页面内容、编辑单页、重新生成、优化文案、导出 PPTX/PDF。

## 运行

```bash
npm install
npm run dev
```

打开 `http://localhost:3000`。

## 环境变量

复制 `.env.example` 为 `.env.local`，只在服务器环境中配置真实 Key：

```bash
OPENAI_API_KEY=你的服务器端 Key
OPENAI_MODEL=gpt-5.5
OPENAI_BASE_URL=https://api.openai.com
```

不要把真实 Key 写进前端代码、后端源码、Git 仓库或页面。部署到 Vercel、Render、Railway 等平台时，请在 Secrets / Environment Variables 中配置 `OPENAI_API_KEY`。前端只请求本项目的 `/api/*` 接口，OpenAI 调用都在服务端完成。

`OPENAI_BASE_URL` 默认是 OpenAI 官方接口 `https://api.openai.com`。只有在你明确使用兼容 Responses API 的代理或网关时才改它，并确认该服务可信。

本地没有 Key 时可临时设置 `USE_MOCK_AI=true` 测试界面流程，但生产环境应关闭。

## 主要模块

- `src/app/api/generate-outline/route.ts`：服务端生成 PPT 大纲。
- `src/app/api/generate-slides/route.ts`：服务端生成逐页内容。
- `src/app/api/rewrite-slide/route.ts`：重新生成、优化、缩短、扩展单页。
- `src/app/api/analyze-template/route.ts`：解析上传 PPTX 模板基础信息。
- `src/app/api/export/route.ts`：导出 PPTX 或 PDF。
- `src/lib/ppt/templateParser.ts`：模板解析预留模块。
- `src/lib/ppt/themeExtractor.ts`：模板主题提取模块。
- `src/lib/ppt/slideGenerator.ts`：页面内容与模板样式合成模块。
- `src/lib/ppt/pptxExporter.ts`：PPTX 生成模块。
- `src/lib/ppt/pdfExporter.ts`：PDF fallback 导出模块。

## 当前模板能力

上传 `.pptx` 后，系统会读取页面数量、页面比例、主题色、字体和部分布局信号，并在生成内容和导出时尽量复用。复杂模板的完整 XML 级复刻已通过 `templateParser`、`themeExtractor`、`slideGenerator/pptxExporter` 的结构预留，可后续替换为更深入的 Office Open XML 处理。

## 部署说明

1. 推送代码到部署平台。
2. 在平台 Secrets / Environment Variables 中配置 `OPENAI_API_KEY`。
3. 可选配置 `OPENAI_MODEL=gpt-5.5`。
4. 执行构建命令 `npm run build`，启动命令 `npm run start`。
