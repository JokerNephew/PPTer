export type BuiltInTemplateCategory =
  | "business"
  | "academic"
  | "startup"
  | "product"
  | "education"
  | "data";

export type GenerationStyle =
  | "professional"
  | "minimal"
  | "visual"
  | "academic"
  | "executive"
  | "storytelling";

export type GenerationStage =
  | "idle"
  | "reading-template"
  | "analyzing-topic"
  | "generating-outline"
  | "generating-slides"
  | "applying-template"
  | "preparing-export"
  | "done"
  | "error";

export type TemplateSource = "built-in" | "uploaded";
export type ExternalTemplateSource =
  | "SlidesCarnival"
  | "PresentationGO"
  | "PPTMON"
  | "Slidesgo";

export interface DeckRequest {
  topic: string;
  purpose: string;
  audience: string;
  slideCount: number;
  language: string;
  style: GenerationStyle;
  templateId: string;
  templateSource: TemplateSource;
  referenceText?: string;
  outline?: OutlineSection[];
  templateProfile?: TemplateProfile | null;
}

export interface TemplateProfile {
  id: string;
  name: string;
  source: TemplateSource;
  category?: BuiltInTemplateCategory;
  pageCount: number;
  aspectRatio: "16:9" | "4:3" | "wide" | "unknown";
  dominantColors: string[];
  fonts: string[];
  titleHierarchy: string[];
  detectedLayouts: string[];
  hasCover: boolean;
  hasAgenda: boolean;
  hasClosing: boolean;
  thumbnail?: string;
  warnings: string[];
  fallbackStrategy: string;
  externalUrl?: string;
  externalSource?: ExternalTemplateSource;
  externalLicenseNote?: string;
}

export interface ExternalTemplateResult {
  id: string;
  title: string;
  source: ExternalTemplateSource;
  url: string;
  thumbnail?: string;
  description?: string;
  formats: string[];
  aspectRatio?: string;
  slideCount?: number;
  tags: string[];
  isFree: boolean;
  downloadUrl?: string;
  licenseNote: string;
}

export interface TemplateSearchResponse {
  query: string;
  results: ExternalTemplateResult[];
  sources: Array<{
    name: ExternalTemplateSource;
    status: "ok" | "error";
    message?: string;
  }>;
}

export interface OutlineSection {
  id: string;
  title: string;
  intent: string;
  slideEstimate: number;
  keyPoints: string[];
}

export interface SlideContent {
  id: string;
  slideNumber: number;
  sectionId: string;
  title: string;
  body: string[];
  chartSuggestion: string;
  imageSuggestion: string;
  layoutSuggestion: string;
  speakerNotes: string;
  colorPalette: string[];
  templateRole: "cover" | "agenda" | "content" | "data" | "closing";
}

export interface DeckExportPayload {
  fileName: string;
  format: "pptx" | "pdf";
  request: DeckRequest;
  slides: SlideContent[];
  templateProfile?: TemplateProfile | null;
}

export interface ApiErrorBody {
  error: string;
  details?: string;
}
