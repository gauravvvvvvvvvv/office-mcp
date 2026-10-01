import { createPresentation, type PowerPointElement, type PowerPointSlideSpec } from "./powerpoint.js";

type DesignStyle = "editorial" | "cinematic";

type DesignedSlide = (
  | { layout: "cover"; title: string; subtitle?: string; imagePath?: string; imageAltText?: string }
  | { layout: "imageText"; title: string; body: string; imagePath: string; imageAltText?: string; imageSide?: "left" | "right" }
  | { layout: "statement"; title: string; body?: string }
  | { layout: "comparison"; title: string; left: { heading: string; body: string }; right: { heading: string; body: string }; takeaway?: string }
  | { layout: "panorama"; title: string; body?: string; imagePath: string; imageAltText?: string }
) & { speakerNotes?: string };

export interface DesignedPresentationSpec {
  path: string;
  title?: string;
  style?: DesignStyle;
  slides: DesignedSlide[];
  overwrite?: boolean;
}

interface Palette {
  background: string;
  foreground: string;
  secondary: string;
  accent: string;
  headFont: string;
  bodyFont: string;
}

const palettes: Record<DesignStyle, Palette> = {
  editorial: {
    background: "F7F5F0",
    foreground: "1F2830",
    secondary: "48545F",
    accent: "9B4B31",
    headFont: "Georgia",
    bodyFont: "Aptos"
  },
  cinematic: {
    background: "10151F",
    foreground: "F8F7F3",
    secondary: "C9CDD3",
    accent: "F0B663",
    headFont: "Aptos Display",
    bodyFont: "Aptos"
  }
};

function text(
  value: string,
  x: number,
  y: number,
  w: number,
  h: number,
  fontSize: number,
  color: string,
  fontFace: string,
  bold = false
): PowerPointElement {
  return { type: "text", text: value, x, y, w, h, fontSize, color, fontFace, bold, margin: 0, valign: "middle" };
}

function image(path: string, x: number, y: number, w: number, h: number, altText?: string): PowerPointElement {
  return { type: "image", path, x, y, w, h, fit: "cover", altText };
}

function composeSlide(slide: DesignedSlide, palette: Palette): PowerPointSlideSpec {
  const elements: PowerPointElement[] = [];

  if (slide.layout === "cover") {
    if (slide.imagePath) {
      elements.push(image(slide.imagePath, 7.28, 0, 6.05, 7.5, slide.imageAltText));
      elements.push(text(slide.title, 0.78, 1.35, 5.75, 2.55, 46, palette.foreground, palette.headFont, true));
      if (slide.subtitle) elements.push(text(slide.subtitle, 0.82, 4.47, 5.58, 1.65, 21, palette.secondary, palette.bodyFont));
    } else {
      elements.push(text(slide.title, 0.84, 1.45, 11.6, 2.9, 56, palette.foreground, palette.headFont, true));
      if (slide.subtitle) elements.push(text(slide.subtitle, 0.88, 4.72, 10.65, 1.25, 23, palette.secondary, palette.bodyFont));
    }
  } else if (slide.layout === "imageText") {
    const imageLeft = slide.imageSide === "left";
    const imageX = imageLeft ? 0 : 6.48;
    const textX = imageLeft ? 7.18 : 0.82;
    elements.push(image(slide.imagePath, imageX, 0, 6.85, 7.5, slide.imageAltText));
    elements.push(text(slide.title, textX, 1.18, 5.28, 1.92, 36, palette.foreground, palette.headFont, true));
    elements.push(text(slide.body, textX + 0.03, 3.49, 5.05, 2.35, 21, palette.secondary, palette.bodyFont));
  } else if (slide.layout === "statement") {
    elements.push(text(slide.title, 0.86, 1.48, 11.55, 2.6, 49, palette.foreground, palette.headFont, true));
    if (slide.body) elements.push(text(slide.body, 0.9, 4.64, 10.8, 1.54, 23, palette.secondary, palette.bodyFont));
  } else if (slide.layout === "comparison") {
    elements.push(text(slide.title, 0.8, 0.61, 11.8, 1.27, 35, palette.foreground, palette.headFont, true));
    elements.push(text(slide.left.heading, 0.83, 2.35, 5.3, 0.65, 21, palette.accent, palette.bodyFont, true));
    elements.push(text(slide.left.body, 0.86, 3.23, 5.25, 2.85, 22, palette.foreground, palette.bodyFont));
    elements.push(text(slide.right.heading, 7.08, 2.35, 5.25, 0.65, 21, palette.accent, palette.bodyFont, true));
    elements.push(text(slide.right.body, 7.11, 3.23, 5.2, 2.85, 22, palette.foreground, palette.bodyFont));
    if (slide.takeaway) elements.push(text(slide.takeaway, 0.86, 6.34, 11.6, 0.64, 20, palette.accent, palette.bodyFont, true));
  } else {
    elements.push(image(slide.imagePath, 0, 0, 13.333, 4.73, slide.imageAltText));
    elements.push(text(slide.title, 0.82, 5.06, slide.body ? 7.1 : 11.7, 1.58, 36, palette.foreground, palette.headFont, true));
    if (slide.body) elements.push(text(slide.body, 8.45, 5.19, 4.0, 1.37, 19, palette.secondary, palette.bodyFont));
  }

  return { backgroundColor: palette.background, speakerNotes: slide.speakerNotes, elements };
}

function designWarnings(slides: DesignedSlide[]): string[] {
  const warnings: string[] = [];
  if (slides.length >= 3 && slides.every((slide) => slide.layout === slides[0].layout)) {
    warnings.push("Every slide uses the same composition. Vary layouts when the content warrants it.");
  }
  if (slides.length >= 3 && !slides.some((slide) => "imagePath" in slide && slide.imagePath)) {
    warnings.push("No slide contains an image. Consider relevant visual assets rather than decorative shapes.");
  }
  const imagePaths = slides.flatMap((slide) => "imagePath" in slide && slide.imagePath ? [slide.imagePath] : []);
  if (new Set(imagePaths).size < imagePaths.length) {
    warnings.push("An image appears on multiple slides. Reuse it only when the repetition is deliberate.");
  }
  return warnings;
}

export async function createDesignedPresentation(spec: DesignedPresentationSpec) {
  const palette = palettes[spec.style ?? "editorial"];
  const created = await createPresentation({
    path: spec.path,
    title: spec.title ?? ("title" in spec.slides[0] ? spec.slides[0].title : ""),
    layout: "wide",
    theme: { headFontFace: palette.headFont, bodyFontFace: palette.bodyFont },
    slides: spec.slides.map((slide) => composeSlide(slide, palette)),
    overwrite: spec.overwrite
  });
  return { ...created, designWarnings: designWarnings(spec.slides) };
}
