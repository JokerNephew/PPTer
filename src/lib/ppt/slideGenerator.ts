import type { DeckRequest, SlideContent, TemplateProfile } from "@/types/ppt";
import { enforcePptSkillOnSlide } from "./pptSkillGuidelines";
import { getBuiltInTemplate } from "./templates";

export interface SlideGenerationContext {
  request: DeckRequest;
  template: TemplateProfile;
}

export function createSlideGenerationContext(request: DeckRequest): SlideGenerationContext {
  return {
    request,
    template: request.templateProfile ?? getBuiltInTemplate(request.templateId),
  };
}

export function applyTemplateStyleToSlides(
  slides: SlideContent[],
  template: TemplateProfile,
): SlideContent[] {
  return slides.map((slide, index) => {
    const isCover = index === 0;
    const isClosing = index === slides.length - 1;

    return enforcePptSkillOnSlide(
      {
        ...slide,
        colorPalette: slide.colorPalette?.length ? slide.colorPalette : template.dominantColors,
        templateRole: isCover
          ? "cover"
          : isClosing
            ? "closing"
            : slide.templateRole,
        layoutSuggestion: slide.layoutSuggestion || template.fallbackStrategy,
      },
      index,
      slides.length,
      template,
    );
  });
}
