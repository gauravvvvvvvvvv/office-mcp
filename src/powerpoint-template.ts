import path from "node:path";
import { OfficeMcpError } from "./errors.js";
import { runNativeBatch } from "./native.js";
import { resolveAllowedPath, resolveReadablePath } from "./paths.js";
import { inspectPresentation } from "./powerpoint.js";

type TemplateContent = {
  placeholderText?: Array<{ index: number; text: string }>;
  shapeText?: Array<{ name: string; text: string }>;
  textBoxes?: Array<{
    text: string; left: number; top: number; width: number; height: number;
    fontSize?: number; fontFace?: string; color?: string; bold?: boolean;
  }>;
  speakerNotes?: string;
};

export type TemplateSlide = TemplateContent & (
  | { source: "layout"; designIndex: number; layoutIndex: number }
  | { source: "duplicate"; sourceSlide: number }
);

export interface CreateFromTemplateSpec {
  templatePath: string;
  outputPath: string;
  slides: TemplateSlide[];
  keepTemplateSlides?: boolean;
  overwrite?: boolean;
}

export async function inspectTemplate(templatePath: string) {
  const path = await resolveReadablePath(templatePath, [".pptx"]);
  const result = await runNativeBatch("powerpoint", { target: { path }, operations: [{ op: "inspect" }] }) as {
    results: Array<{ designs: unknown[]; slides: unknown[]; width: number; height: number; presentationWindowCount: number }>;
  };
  const inspection = result.results[0];
  return {
    path,
    sizePoints: { width: inspection.width, height: inspection.height },
    designs: inspection.designs,
    slides: inspection.slides,
    presentationWindowCount: inspection.presentationWindowCount,
    note: "Use design/layout indexes to add new slides, or duplicate a source slide and update its named shapes. Native inspection stays hidden by default."
  };
}

export async function createFromTemplate(spec: CreateFromTemplateSpec) {
  const templatePath = await resolveReadablePath(spec.templatePath, [".pptx"]);
  const outputPath = resolveAllowedPath(spec.outputPath, [".pptx"]);
  const samePath = process.platform === "win32"
    ? templatePath.toLowerCase() === outputPath.toLowerCase()
    : templatePath === outputPath;
  if (samePath) throw new OfficeMcpError("Template and output paths must differ", "INVALID_TEMPLATE_OUTPUT");
  const template = await inspectPresentation(templatePath);
  const originalCount = template.slideCount;
  const operations: Array<Record<string, unknown> & { op: string }> = [];

  for (const [index, slide] of spec.slides.entries()) {
    const slideNumber = originalCount + index + 1;
    if (slide.source === "layout") {
      operations.push({ op: "add_slide_from_layout", designIndex: slide.designIndex, layoutIndex: slide.layoutIndex });
    } else {
      if (slide.sourceSlide > originalCount) {
        throw new OfficeMcpError(`Source slide ${slide.sourceSlide} does not exist in the template`, "INVALID_TEMPLATE_SLIDE");
      }
      operations.push({ op: "duplicate_slide", slide: slide.sourceSlide, toIndex: slideNumber });
    }
    for (const placeholder of slide.placeholderText ?? []) {
      operations.push({ op: "set_placeholder_text", slide: slideNumber, placeholderIndex: placeholder.index, text: placeholder.text });
    }
    for (const shape of slide.shapeText ?? []) {
      operations.push({ op: "update_shape", slide: slideNumber, name: shape.name, text: shape.text });
    }
    for (const box of slide.textBoxes ?? []) {
      operations.push({ op: "add_text", slide: slideNumber, ...box });
    }
    if (slide.speakerNotes !== undefined) {
      operations.push({ op: "set_speaker_notes", slide: slideNumber, text: slide.speakerNotes });
    }
  }
  if (!spec.keepTemplateSlides) {
    for (let number = originalCount; number >= 1; number--) operations.push({ op: "delete_slide", slide: number });
  }
  operations.push({ op: "save", outputPath });
  await runNativeBatch("powerpoint", {
    target: { path: templatePath },
    operations,
    overwrite: spec.overwrite
  });
  return {
    path: outputPath,
    slideCount: spec.slides.length + (spec.keepTemplateSlides ? originalCount : 0),
    keptTemplateSlides: Boolean(spec.keepTemplateSlides),
    createdFrom: path.basename(templatePath)
  };
}
