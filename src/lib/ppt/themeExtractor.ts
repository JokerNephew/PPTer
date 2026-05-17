import type { TemplateProfile } from "@/types/ppt";
import { getBuiltInTemplate } from "./templates";

export interface ExportTheme {
  name: string;
  colors: {
    background: string;
    surface: string;
    text: string;
    muted: string;
    accent: string;
    accentAlt: string;
  };
  fonts: {
    heading: string;
    body: string;
  };
  aspectRatio: "LAYOUT_WIDE" | "LAYOUT_4X3";
}

function readableText(background: string) {
  const hex = background.replace("#", "");
  if (hex.length !== 6) return "#111827";
  const r = Number.parseInt(hex.slice(0, 2), 16);
  const g = Number.parseInt(hex.slice(2, 4), 16);
  const b = Number.parseInt(hex.slice(4, 6), 16);
  const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
  return luminance > 0.55 ? "#111827" : "#F8FAFC";
}

export function extractTheme(templateProfile?: TemplateProfile | null, templateId?: string): ExportTheme {
  const profile = templateProfile ?? getBuiltInTemplate(templateId ?? "boardroom-brief");
  const palette = profile.dominantColors.length ? profile.dominantColors : ["#102820", "#F8FAFC", "#C2A83E"];
  const background = palette[1] ?? "#F8FAFC";
  const accent = palette[2] ?? palette[0] ?? "#0F766E";
  const text = readableText(background) === "#111827" ? "#111827" : "#F8FAFC";

  return {
    name: profile.name,
    colors: {
      background,
      surface: text === "#111827" ? "#FFFFFF" : "#18212F",
      text,
      muted: text === "#111827" ? "#475569" : "#CBD5E1",
      accent,
      accentAlt: palette[0] ?? "#102820",
    },
    fonts: {
      heading: profile.fonts[0] ?? "Aptos Display",
      body: profile.fonts[1] ?? profile.fonts[0] ?? "Aptos",
    },
    aspectRatio: profile.aspectRatio === "4:3" ? "LAYOUT_4X3" : "LAYOUT_WIDE",
  };
}
