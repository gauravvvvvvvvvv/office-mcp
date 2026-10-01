import JSZip from "jszip";
import path from "node:path";
import { createRequire } from "node:module";
import { readFile, writeFile } from "node:fs/promises";
import { OfficeMcpError } from "./errors.js";
import { prepareOutputPath, resolveInputOutput, resolveReadablePath } from "./paths.js";

const POWERPOINT_EXTENSIONS = [".pptx"];
const IMAGE_EXTENSIONS = [".png", ".jpg", ".jpeg", ".gif", ".svg"];

interface SlideLike {
  background: { color: string };
  addText(text: string, options: Record<string, unknown>): unknown;
  addShape(shape: string, options: Record<string, unknown>): unknown;
  addImage(options: Record<string, unknown>): unknown;
  addTable(rows: unknown[], options: Record<string, unknown>): unknown;
  addChart(type: string, data: unknown[], options: Record<string, unknown>): unknown;
  addNotes(notes: string): unknown;
}

interface PresentationLike {
  layout: string;
  author: string;
  company: string;
  subject: string;
  title: string;
  lang: string;
  theme: Record<string, unknown>;
  ShapeType: Record<string, string>;
  ChartType: Record<string, string>;
  addSlide(): SlideLike;
  writeFile(options: { fileName: string; compression: boolean }): Promise<string>;
}

const require = createRequire(import.meta.url);
const PptxConstructor = require("pptxgenjs") as new () => PresentationLike;

interface Position {
  x: number;
  y: number;
  w: number;
  h: number;
}

export type PowerPointElement =
  | (Position & {
      type: "text";
      text: string;
      fontSize?: number;
      fontFace?: string;
      color?: string;
      bold?: boolean;
      italic?: boolean;
      align?: "left" | "center" | "right" | "justify";
      valign?: "top" | "middle" | "bottom";
      margin?: number;
    })
  | (Position & {
      type: "shape";
      shape?: "rect" | "roundRect" | "ellipse" | "line";
      fillColor?: string;
      lineColor?: string;
      lineWidth?: number;
      text?: string;
    })
  | (Position & {
      type: "image";
      path: string;
      altText?: string;
      transparency?: number;
      fit?: "cover" | "contain" | "stretch";
    })
  | (Position & {
      type: "table";
      rows: string[][];
      fontSize?: number;
      headerFill?: string;
      borderColor?: string;
    })
  | (Position & {
      type: "chart";
      chartType: "bar" | "line" | "pie" | "doughnut" | "area";
      series: Array<{ name: string; labels: string[]; values: number[] }>;
      title?: string;
      showLegend?: boolean;
      showValue?: boolean;
      showCategoryName?: boolean;
    });

export interface PowerPointSlideSpec {
  backgroundColor?: string;
  speakerNotes?: string;
  elements: PowerPointElement[];
}

export interface PowerPointSpec {
  path: string;
  layout?: "wide" | "standard";
  title?: string;
  subject?: string;
  author?: string;
  company?: string;
  theme?: {
    headFontFace?: string;
    bodyFontFace?: string;
    language?: string;
  };
  slides: PowerPointSlideSpec[];
  overwrite?: boolean;
}

function cleanColor(color: string | undefined, fallback: string): string {
  return (color ?? fallback).replace(/^#/, "").toUpperCase();
}

function decodeXml(text: string): string {
  return text
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function encodeXml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function slideText(xml: string): string[] {
  return [...xml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((match) => decodeXml(match[1]));
}

async function addElement(slide: SlideLike, element: PowerPointElement, pptx: PresentationLike) {
  const position = { x: element.x, y: element.y, w: element.w, h: element.h };

  if (element.type === "text") {
    slide.addText(element.text, {
      ...position,
      fontSize: element.fontSize ?? 18,
      fontFace: element.fontFace,
      color: cleanColor(element.color, "1F2937"),
      bold: element.bold,
      italic: element.italic,
      align: element.align,
      valign: element.valign,
      margin: element.margin ?? 0.08,
      breakLine: false,
      fit: "shrink"
    });
    return;
  }

  if (element.type === "shape") {
    const shapeMap = {
      rect: pptx.ShapeType.rect,
      roundRect: pptx.ShapeType.roundRect,
      ellipse: pptx.ShapeType.ellipse,
      line: pptx.ShapeType.line
    };
    slide.addShape(shapeMap[element.shape ?? "rect"], {
      ...position,
      fill: element.shape === "line" ? undefined : { color: cleanColor(element.fillColor, "E8EEF4") },
      line: {
        color: cleanColor(element.lineColor, "5B7083"),
        width: element.lineWidth ?? 1
      }
    });
    if (element.text) {
      slide.addText(element.text, {
        ...position,
        fontSize: 16,
        align: "center",
        valign: "middle",
        margin: 0.08,
        fit: "shrink"
      });
    }
    return;
  }

  if (element.type === "image") {
    const imagePath = await resolveReadablePath(element.path, IMAGE_EXTENSIONS);
    slide.addImage({
      ...position,
      path: imagePath,
      altText: element.altText ?? "Description not provided",
      transparency: element.transparency,
      sizing: element.fit && element.fit !== "stretch"
        ? { type: element.fit, w: element.w, h: element.h }
        : undefined
    });
    return;
  }

  if (element.type === "table") {
    const rows = element.rows.map((row, rowIndex) =>
      row.map((text) => ({
        text,
        options: rowIndex === 0
          ? { bold: true, color: "FFFFFF", fill: cleanColor(element.headerFill, "1F4E78") }
          : {}
      }))
    );
    slide.addTable(rows, {
      ...position,
      fontSize: element.fontSize ?? 12,
      border: { color: cleanColor(element.borderColor, "B7C9D6"), width: 1 },
      margin: 0.06,
      valign: "middle",
      autoFit: false
    });
    return;
  }

  const chartTypeMap = {
    bar: pptx.ChartType.bar,
    line: pptx.ChartType.line,
    pie: pptx.ChartType.pie,
    doughnut: pptx.ChartType.doughnut,
    area: pptx.ChartType.area
  };
  slide.addChart(chartTypeMap[element.chartType], element.series, {
    ...position,
    showTitle: Boolean(element.title),
    title: element.title,
    showLegend: element.showLegend ?? element.series.length > 1,
    showValue: element.showValue,
    showCategoryName: element.showCategoryName,
    showCatName: element.showCategoryName,
    catAxisLabelFontSize: 10,
    valAxisLabelFontSize: 10,
    chartColors: ["2F75B5", "70AD47", "ED7D31", "A5A5A5", "FFC000"],
    showBorder: false
  });
}

export async function createPresentation(spec: PowerPointSpec) {
  const outputPath = await prepareOutputPath(spec.path, POWERPOINT_EXTENSIONS, spec.overwrite);
  const pptx = new PptxConstructor();
  pptx.layout = spec.layout === "standard" ? "LAYOUT_4X3" : "LAYOUT_WIDE";
  pptx.author = spec.author ?? "Office MCP";
  pptx.company = spec.company ?? "";
  pptx.subject = spec.subject ?? "";
  pptx.title = spec.title ?? "";
  pptx.lang = spec.theme?.language ?? "en-US";
  pptx.theme = {
    headFontFace: spec.theme?.headFontFace ?? "Aptos Display",
    bodyFontFace: spec.theme?.bodyFontFace ?? "Aptos",
    lang: spec.theme?.language ?? "en-US"
  };

  for (const slideSpec of spec.slides) {
    const slide = pptx.addSlide();
    if (slideSpec.backgroundColor) {
      slide.background = { color: cleanColor(slideSpec.backgroundColor, "FFFFFF") };
    }
    for (const element of slideSpec.elements) {
      await addElement(slide, element, pptx);
    }
    if (slideSpec.speakerNotes) slide.addNotes(slideSpec.speakerNotes);
  }

  await pptx.writeFile({ fileName: outputPath, compression: true });
  return { path: outputPath, slideCount: spec.slides.length };
}

async function loadPresentation(filePath: string) {
  const resolvedPath = await resolveReadablePath(filePath, POWERPOINT_EXTENSIONS);
  const zip = await JSZip.loadAsync(await readFile(resolvedPath));
  return { resolvedPath, zip };
}

function numberedEntries(zip: JSZip, expression: RegExp) {
  return Object.keys(zip.files)
    .map((name) => ({ name, number: Number(expression.exec(name)?.[1] ?? Number.NaN) }))
    .filter((entry) => Number.isFinite(entry.number))
    .sort((a, b) => a.number - b.number);
}

export async function inspectPresentation(filePath: string, maxSlides = 200) {
  const { resolvedPath, zip } = await loadPresentation(filePath);
  const slideEntries = numberedEntries(zip, /^ppt\/slides\/slide(\d+)\.xml$/).slice(0, maxSlides);
  const slides = [];

  for (const entry of slideEntries) {
    const xml = await zip.file(entry.name)!.async("string");
    const texts = slideText(xml);
    const notesEntry = zip.file(`ppt/notesSlides/notesSlide${entry.number}.xml`);
    const notes = notesEntry ? slideText(await notesEntry.async("string")).filter((text) => text.trim()).join(" ") : null;
    slides.push({
      number: entry.number,
      title: texts[0] ?? null,
      text: texts,
      speakerNotes: notes,
      shapeCount: (xml.match(/<p:sp\b/g) ?? []).length,
      pictureCount: (xml.match(/<p:pic\b/g) ?? []).length,
      chartCount: (xml.match(/<c:chart\b/g) ?? []).length
    });
  }

  const presentationXml = await zip.file("ppt/presentation.xml")?.async("string");
  const sizeMatch = presentationXml && /<p:sldSz\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(presentationXml);
  return {
    kind: "powerpoint",
    path: resolvedPath,
    filename: path.basename(resolvedPath),
    slideCount: numberedEntries(zip, /^ppt\/slides\/slide(\d+)\.xml$/).length,
    sizeInches: sizeMatch
      ? { width: Number(sizeMatch[1]) / 914400, height: Number(sizeMatch[2]) / 914400 }
      : null,
    slides,
    truncated: slideEntries.length < numberedEntries(zip, /^ppt\/slides\/slide(\d+)\.xml$/).length
  };
}

function replaceLiteral(text: string, find: string, replacement: string, matchCase: boolean) {
  if (!find) throw new OfficeMcpError("The find text cannot be empty", "EMPTY_FIND_TEXT");
  if (matchCase) {
    const parts = text.split(find);
    return { text: parts.join(replacement), count: parts.length - 1 };
  }
  const escaped = find.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  let count = 0;
  return {
    text: text.replace(new RegExp(escaped, "gi"), () => {
      count += 1;
      return replacement;
    }),
    count
  };
}

export async function replacePresentationText(
  filePath: string,
  find: string,
  replacement: string,
  outputPath?: string,
  overwrite = false,
  matchCase = true
) {
  const { input, output } = resolveInputOutput(filePath, outputPath, POWERPOINT_EXTENSIONS);
  if (input !== output) await prepareOutputPath(output, POWERPOINT_EXTENSIONS, overwrite);
  const { zip } = await loadPresentation(input);
  const slides = numberedEntries(zip, /^ppt\/slides\/slide(\d+)\.xml$/);
  let replacements = 0;

  for (const entry of slides) {
    const xml = await zip.file(entry.name)!.async("string");
    const updated = xml.replace(/<a:p\b[\s\S]*?<\/a:p>/g, (paragraphXml) => {
      const tags = [...paragraphXml.matchAll(/<a:t>[\s\S]*?<\/a:t>/g)];
      if (!tags.length) return paragraphXml;
      const currentText = slideText(paragraphXml).join("");
      const result = replaceLiteral(currentText, find, replacement, matchCase);
      if (!result.count) return paragraphXml;
      replacements += result.count;
      let first = true;
      return paragraphXml.replace(/<a:t>[\s\S]*?<\/a:t>/g, () => {
        if (!first) return "<a:t></a:t>";
        first = false;
        return `<a:t>${encodeXml(result.text)}</a:t>`;
      });
    });
    zip.file(entry.name, updated);
  }

  if (!replacements) throw new OfficeMcpError(`Text was not found: ${find}`, "TEXT_NOT_FOUND");
  await writeFile(output, await zip.generateAsync({ type: "nodebuffer" }));
  return { path: output, replacements };
}
