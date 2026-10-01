import { readFile } from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import { OfficeMcpError } from "./errors.js";
import { resolveReadablePath } from "./paths.js";

const EMU_PER_INCH = 914400;
const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });

type XmlNode = Record<string, unknown>;
type Severity = "warning" | "review";

export interface PowerPointAuditIssue {
  slide: number;
  code: "OFF_SLIDE" | "SMALL_TEXT" | "DENSE_SLIDE" | "MISSING_ALT_TEXT" | "EMPTY_SLIDE" | "UNMEASURED_GROUP" | "LOW_CONTRAST" | "TEXT_OVERLAP";
  severity: Severity;
  message: string;
  object?: string;
}

function node(value: unknown): XmlNode {
  return value && typeof value === "object" && !Array.isArray(value) ? value as XmlNode : {};
}

function list(value: unknown): unknown[] {
  return value === undefined ? [] : Array.isArray(value) ? value : [value];
}

function attribute(value: unknown, name: string): string | undefined {
  const result = node(value)[`@_${name}`];
  return result === undefined ? undefined : String(result);
}

function collect(nodeValue: unknown, key: string): unknown[] {
  if (Array.isArray(nodeValue)) return nodeValue.flatMap((item) => collect(item, key));
  if (!nodeValue || typeof nodeValue !== "object") return [];
  const current = nodeValue as XmlNode;
  return [...list(current[key]), ...Object.values(current).flatMap((child) => collect(child, key))];
}

function textContent(shape: unknown): string {
  return collect(shape, "a:t").map((value) => String(value)).join(" ").trim();
}

function shapeName(shape: XmlNode, type: string, index: number): string {
  const nonVisual = type === "p:pic" ? node(shape["p:nvPicPr"]) : node(shape["p:nvSpPr"]);
  return attribute(nonVisual["p:cNvPr"], "name") ?? `${type} ${index}`;
}

function dimensions(shape: XmlNode, type: string) {
  const transform = type === "p:graphicFrame"
    ? node(shape["p:xfrm"])
    : node(node(shape["p:spPr"])["a:xfrm"]);
  const offset = node(transform["a:off"] ?? transform["p:off"]);
  const extent = node(transform["a:ext"] ?? transform["p:ext"]);
  const x = Number(attribute(offset, "x"));
  const y = Number(attribute(offset, "y"));
  const width = Number(attribute(extent, "cx"));
  const height = Number(attribute(extent, "cy"));
  if (![x, y, width, height].every(Number.isFinite) || width <= 0 || height <= 0) return null;
  return { x, y, width, height };
}

function solidColor(value: unknown): string | null {
  const solidFill = node(node(value)["a:solidFill"]);
  if (Object.keys(solidFill).length === 0 || collect(solidFill, "a:alpha").length > 0) return null;
  const color = attribute(solidFill["a:srgbClr"], "val");
  return color && /^[\da-f]{6}$/i.test(color) ? color.toUpperCase() : null;
}

function luminance(hex: string): number {
  const channels = [0, 2, 4].map((offset) => Number.parseInt(hex.slice(offset, offset + 2), 16) / 255)
    .map((value) => value <= 0.04045 ? value / 12.92 : ((value + 0.055) / 1.055) ** 2.4);
  return channels[0] * 0.2126 + channels[1] * 0.7152 + channels[2] * 0.0722;
}

function contrastRatio(left: string, right: string): number {
  const first = luminance(left);
  const second = luminance(right);
  return (Math.max(first, second) + 0.05) / (Math.min(first, second) + 0.05);
}

function intersectionShare(first: NonNullable<ReturnType<typeof dimensions>>, second: NonNullable<ReturnType<typeof dimensions>>) {
  const width = Math.max(0, Math.min(first.x + first.width, second.x + second.width) - Math.max(first.x, second.x));
  const height = Math.max(0, Math.min(first.y + first.height, second.y + second.height) - Math.max(first.y, second.y));
  return (width * height) / Math.min(first.width * first.height, second.width * second.height);
}

function slideEntries(zip: JSZip, presentation: XmlNode, relationships: XmlNode) {
  const available = Object.keys(zip.files)
    .map((name) => ({ name, number: Number(/^ppt\/slides\/slide(\d+)\.xml$/.exec(name)?.[1]) }))
    .filter((entry) => Number.isInteger(entry.number) && entry.number > 0)
    .sort((left, right) => left.number - right.number);
  const relationshipTargets = new Map(
    list(relationships["Relationship"]).map((relationship) => {
      const identifier = attribute(relationship, "Id") ?? "";
      const target = attribute(relationship, "Target") ?? "";
      const name = target.startsWith("/")
        ? target.slice(1)
        : path.posix.normalize(path.posix.join("ppt", target.replaceAll("\\", "/")));
      return [identifier, name];
    })
  );
  const ordered = list(node(presentation["p:sldIdLst"])["p:sldId"])
    .map((slideId, index) => ({ name: relationshipTargets.get(attribute(slideId, "r:id") ?? "") ?? "", number: index + 1 }))
    .filter((entry) => Boolean(zip.file(entry.name)));
  return ordered.length === available.length ? ordered : available.map((entry, index) => ({ name: entry.name, number: index + 1 }));
}

export async function auditPresentation(filePath: string, minimumFontPoints = 14, denseSlideCharacters = 600) {
  const resolvedPath = await resolveReadablePath(filePath, [".pptx"]);
  const zip = await JSZip.loadAsync(await readFile(resolvedPath));
  const presentationFile = zip.file("ppt/presentation.xml");
  if (!presentationFile) throw new OfficeMcpError("PPTX is missing presentation.xml", "INVALID_PRESENTATION");
  const presentation = node(parser.parse(await presentationFile.async("string"))["p:presentation"]);
  const relationshipsFile = zip.file("ppt/_rels/presentation.xml.rels");
  const relationships = relationshipsFile
    ? node(parser.parse(await relationshipsFile.async("string"))["Relationships"])
    : {};
  const slideSize = node(presentation["p:sldSz"]);
  const slideWidth = Number(attribute(slideSize, "cx"));
  const slideHeight = Number(attribute(slideSize, "cy"));
  if (!Number.isFinite(slideWidth) || !Number.isFinite(slideHeight) || slideWidth <= 0 || slideHeight <= 0) {
    throw new OfficeMcpError("PPTX has no valid slide size", "INVALID_PRESENTATION");
  }

  const issues: PowerPointAuditIssue[] = [];
  const slides = slideEntries(zip, presentation, relationships);
  for (const entry of slides) {
    const slide = node(parser.parse(await zip.file(entry.name)!.async("string"))["p:sld"]);
    const tree = node(node(slide["p:cSld"])["p:spTree"]);
    const background = solidColor(node(node(node(slide["p:cSld"])["p:bg"])["p:bgPr"]));
    const objects = ["p:sp", "p:pic", "p:graphicFrame", "p:cxnSp"].flatMap((type) =>
      list(tree[type]).map((value) => ({ type, value: node(value) }))
    );
    const allText = objects.map(({ value }) => textContent(value)).filter(Boolean).join(" ");
    if (objects.length === 0) {
      issues.push({ slide: entry.number, code: "EMPTY_SLIDE", severity: "review", message: "Slide has no visible objects." });
    }
    if (allText.length > denseSlideCharacters) {
      issues.push({
        slide: entry.number,
        code: "DENSE_SLIDE",
        severity: "review",
        message: `Slide contains ${allText.length} characters of visible text; review readability and consider splitting it.`
      });
    }
    if (list(tree["p:grpSp"]).length > 0) {
      issues.push({
        slide: entry.number,
        code: "UNMEASURED_GROUP",
        severity: "review",
        message: "Grouped objects are present; this audit does not transform their child coordinates. Inspect the rendered slide."
      });
    }

    const textBoxes: Array<{ name: string; box: NonNullable<ReturnType<typeof dimensions>> }> = [];
    for (const [index, { type, value }] of objects.entries()) {
      const name = shapeName(value, type, index + 1);
      const box = dimensions(value, type);
      const tolerance = EMU_PER_INCH * 0.05;
      if (box && (box.x < -tolerance || box.y < -tolerance ||
        box.x + box.width > slideWidth + tolerance || box.y + box.height > slideHeight + tolerance)) {
        issues.push({
          slide: entry.number,
          code: "OFF_SLIDE",
          severity: "review",
          object: name,
          message: "Object extends beyond the slide bounds. This may be intentional bleed; inspect the rendered slide."
        });
      }
      if (type === "p:sp" && textContent(value)) {
        if (box) textBoxes.push({ name, box });
        const sizes = [...collect(value, "a:rPr"), ...collect(value, "a:defRPr"), ...collect(value, "a:endParaRPr")]
          .map((item) => Number(attribute(item, "sz")) / 100)
          .filter((size) => Number.isFinite(size) && size > 0);
        if (sizes.length && Math.min(...sizes) < minimumFontPoints) {
          issues.push({
            slide: entry.number,
            code: "SMALL_TEXT",
            severity: "review",
            object: name,
            message: `Explicit text size reaches ${Math.min(...sizes)} pt, below the ${minimumFontPoints} pt review threshold.`
          });
        }
        const fill = solidColor(value["p:spPr"]) ?? background;
        const textColors = [...collect(value, "a:rPr"), ...collect(value, "a:defRPr")]
          .map(solidColor).filter((color): color is string => Boolean(color));
        if (fill && textColors.length) {
          const minimumContrast = Math.min(...textColors.map((color) => contrastRatio(color, fill)));
          const threshold = sizes.length && Math.min(...sizes) >= 18 ? 3 : 4.5;
          if (minimumContrast < threshold) {
            issues.push({
              slide: entry.number, code: "LOW_CONTRAST", severity: "review", object: name,
              message: `Explicit text/background contrast is ${minimumContrast.toFixed(2)}:1, below the ${threshold}:1 review threshold.`
            });
          }
        }
      }
      if (type === "p:pic") {
        const properties = node(node(value["p:nvPicPr"])["p:cNvPr"]);
        const description = attribute(properties, "descr") ?? attribute(properties, "title") ?? "";
        if (!description || description === "Description not provided" || /\.(png|jpe?g|gif|svg|webp)$/i.test(description)) {
          issues.push({
            slide: entry.number,
            code: "MISSING_ALT_TEXT",
            severity: "review",
            object: name,
            message: "Picture has no meaningful alt text (a file path is not a description). Describe it if informative, or mark it decorative."
          });
        }
      }
    }
    for (let left = 0; left < textBoxes.length; left++) {
      for (let right = left + 1; right < textBoxes.length; right++) {
        if (intersectionShare(textBoxes[left].box, textBoxes[right].box) > 0.2) {
          issues.push({
            slide: entry.number, code: "TEXT_OVERLAP", severity: "review",
            object: `${textBoxes[left].name} / ${textBoxes[right].name}`,
            message: "Two text boxes overlap substantially. This may be intentional; inspect the rendered slide."
          });
        }
      }
    }
  }

  return {
    kind: "powerpoint_audit",
    path: resolvedPath,
    slideCount: slides.length,
    slideSizeInches: { width: slideWidth / EMU_PER_INCH, height: slideHeight / EMU_PER_INCH },
    summary: {
      issueCount: issues.length,
      byCode: Object.fromEntries([...new Set(issues.map((issue) => issue.code))]
        .map((code) => [code, issues.filter((issue) => issue.code === code).length]))
    },
    issues,
    limitations: [
      "This is a structural and heuristic audit, not visual proof of quality or text fit. Contrast checks only apply to explicit solid colors.",
      "Font sizes inherited from a master or theme may not be visible in slide XML.",
      "Render and inspect every slide; animation playback requires separate verification."
    ]
  };
}
