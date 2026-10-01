import {
  AlignmentType,
  BorderStyle,
  Document,
  ExternalHyperlink,
  Footer,
  FootnoteReferenceRun,
  Header,
  HeadingLevel,
  ImageRun,
  LevelFormat,
  PageNumber,
  PageOrientation,
  PageBreak,
  Packer,
  Paragraph,
  SectionType,
  Table,
  TableCell,
  TableOfContents,
  TableRow,
  TextRun,
  WidthType
} from "docx";
import JSZip from "jszip";
import path from "node:path";
import { readFile, writeFile } from "node:fs/promises";
import { OfficeMcpError } from "./errors.js";
import { prepareOutputPath, resolveInputOutput, resolveReadablePath } from "./paths.js";

const WORD_EXTENSIONS = [".docx"];

export type WordBlock =
  | { type: "heading"; text: string; level?: number }
  | { type: "paragraph"; text: string; bold?: boolean; italic?: boolean; align?: "left" | "center" | "right" | "justify" }
  | { type: "bullet"; text: string; level?: number }
  | { type: "number"; text: string; level?: number }
  | { type: "table"; rows: string[][]; headerRow?: boolean }
  | { type: "toc"; maxLevel?: number }
  | { type: "cited_paragraph"; text: string; sourceIds: string[] }
  | { type: "bibliography"; title?: string }
  | { type: "image"; path: string; widthPx: number; heightPx: number; altText?: string; decorative?: boolean; caption?: string }
  | { type: "hyperlink"; text: string; url: string }
  | { type: "page_break" }
  | { type: "footnote_paragraph"; text: string; footnoteText: string };

export interface WordSource {
  id: string;
  author: string;
  title: string;
  year: string;
  url?: string;
}

export interface WordSectionSpec {
  blocks: WordBlock[];
  pageSize?: "letter" | "a4";
  orientation?: "portrait" | "landscape";
  marginInches?: { top?: number; right?: number; bottom?: number; left?: number };
  headerText?: string;
  footerText?: string;
  pageNumbers?: boolean;
  breakType?: "nextPage" | "continuous";
}

export interface WordDocumentSpec {
  path: string;
  title?: string;
  author?: string;
  blocks?: WordBlock[];
  sections?: WordSectionSpec[];
  sources?: WordSource[];
  overwrite?: boolean;
}

function decodeXml(text: string): string {
  return text
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&")
    .replace(/&#(\d+);/g, (_, value: string) => String.fromCodePoint(Number(value)))
    .replace(/&#x([\da-f]+);/gi, (_, value: string) => String.fromCodePoint(Number.parseInt(value, 16)));
}

function encodeXml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function textFromXmlFragment(xml: string): string {
  const parts: string[] = [];
  const tokenPattern = /<w:t\b[^>]*>([\s\S]*?)<\/w:t>|<w:tab\b[^>]*\/>|<w:br\b[^>]*\/>/g;
  let match: RegExpExecArray | null;
  while ((match = tokenPattern.exec(xml))) {
    if (match[1] !== undefined) parts.push(decodeXml(match[1]));
    else if (match[0].startsWith("<w:tab")) parts.push("\t");
    else parts.push("\n");
  }
  return parts.join("");
}

function alignment(value?: "left" | "center" | "right" | "justify") {
  const map = {
    left: AlignmentType.LEFT,
    center: AlignmentType.CENTER,
    right: AlignmentType.RIGHT,
    justify: AlignmentType.JUSTIFIED
  };
  return value ? map[value] : undefined;
}

function headingLevel(level = 1) {
  const levels = [
    HeadingLevel.HEADING_1,
    HeadingLevel.HEADING_2,
    HeadingLevel.HEADING_3,
    HeadingLevel.HEADING_4,
    HeadingLevel.HEADING_5,
    HeadingLevel.HEADING_6
  ];
  return levels[level - 1] ?? HeadingLevel.HEADING_1;
}

async function blockToDocx(
  block: WordBlock,
  sources: Map<string, WordSource>,
  footnotes: Record<string, { children: Paragraph[] }>
): Promise<Array<Paragraph | Table | TableOfContents>> {
  if (block.type === "heading") {
    return [new Paragraph({ text: block.text, heading: headingLevel(block.level) })];
  }

  if (block.type === "paragraph") {
    return [new Paragraph({
      children: [new TextRun({ text: block.text, bold: block.bold, italics: block.italic })],
      alignment: alignment(block.align),
      spacing: { after: 160 }
    })];
  }

  if (block.type === "bullet") {
    return [new Paragraph({
      text: block.text,
      bullet: { level: block.level ?? 0 },
      spacing: { after: 80 }
    })];
  }

  if (block.type === "number") {
    return [new Paragraph({
      text: block.text,
      numbering: { reference: "office-numbering", level: block.level ?? 0 },
      spacing: { after: 80 }
    })];
  }

  if (block.type === "toc") {
    return [new Paragraph({ children: [new TextRun({ text: "Table of Contents", bold: true, size: 32 })], spacing: { after: 160 } }),
      new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: `1-${block.maxLevel ?? 3}`, beginDirty: true })];
  }

  if (block.type === "cited_paragraph") {
    const citations = block.sourceIds.map((id) => {
      const source = sources.get(id);
      if (!source) throw new OfficeMcpError(`Unknown citation source: ${id}`, "UNKNOWN_CITATION_SOURCE");
      const surname = source.author.trim().split(/\s+/).at(-1) ?? source.author;
      return `${surname}, ${source.year}`;
    });
    return [new Paragraph({
      children: [new TextRun(block.text), new TextRun({ text: ` (${citations.join("; ")})`, italics: false })],
      spacing: { after: 160 }
    })];
  }

  if (block.type === "bibliography") {
    return [
      new Paragraph({ text: block.title ?? "References", heading: HeadingLevel.HEADING_1 }),
      ...[...sources.values()].map((source) => new Paragraph({
        text: `${source.author} (${source.year}). ${source.title}.${source.url ? ` ${source.url}` : ""}`,
        spacing: { after: 120 }
      }))
    ];
  }

  if (block.type === "image") {
    if (!block.altText && !block.decorative) {
      throw new OfficeMcpError("Image blocks require altText or decorative=true", "IMAGE_DESCRIPTION_REQUIRED");
    }
    const imagePath = await resolveReadablePath(block.path, [".png", ".jpg", ".jpeg", ".gif", ".bmp"]);
    const extension = path.extname(imagePath).slice(1).toLowerCase();
    const image = new ImageRun({
      type: extension === "jpeg" ? "jpg" : extension as "png" | "jpg" | "gif" | "bmp",
      data: await readFile(imagePath),
      transformation: { width: block.widthPx, height: block.heightPx },
      altText: { name: path.basename(imagePath), description: block.altText },
      decorative: block.decorative
    });
    return [new Paragraph({ children: [image] }),
      ...(block.caption ? [new Paragraph({ text: block.caption, style: "Caption" })] : [])];
  }

  if (block.type === "hyperlink") {
    return [new Paragraph({ children: [new ExternalHyperlink({
      children: [new TextRun({ text: block.text, style: "Hyperlink" })], link: block.url
    })] })];
  }

  if (block.type === "page_break") {
    return [new Paragraph({ children: [new PageBreak()] })];
  }

  if (block.type === "footnote_paragraph") {
    const id = Object.keys(footnotes).length + 1;
    footnotes[String(id)] = { children: [new Paragraph(block.footnoteText)] };
    return [new Paragraph({ children: [new TextRun(block.text), new FootnoteReferenceRun(id)] })];
  }

  return [new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 1, color: "B7C9D6" },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: "B7C9D6" },
      left: { style: BorderStyle.SINGLE, size: 1, color: "B7C9D6" },
      right: { style: BorderStyle.SINGLE, size: 1, color: "B7C9D6" },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "D9E2F3" },
      insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "D9E2F3" }
    },
    rows: block.rows.map(
      (row, rowIndex) =>
        new TableRow({
          tableHeader: Boolean(block.headerRow && rowIndex === 0),
          children: row.map(
            (cell) =>
              new TableCell({
                children: [
                  new Paragraph({
                    children: [new TextRun({ text: cell, bold: block.headerRow && rowIndex === 0 })]
                  })
                ],
                shading: block.headerRow && rowIndex === 0 ? { fill: "D9EAF7" } : undefined
              })
          )
        })
    )
  })];
}

export async function createDocument(spec: WordDocumentSpec) {
  const outputPath = await prepareOutputPath(spec.path, WORD_EXTENSIONS, spec.overwrite);
  if (spec.sections && spec.blocks) {
    throw new OfficeMcpError("Use either blocks or sections, not both", "INVALID_DOCUMENT_SPEC");
  }
  const sections = spec.sections ?? [{ blocks: spec.blocks ?? [] }];
  if (!sections.length) throw new OfficeMcpError("At least one section is required", "INVALID_DOCUMENT_SPEC");
  const sources = new Map<string, WordSource>();
  for (const source of spec.sources ?? []) {
    if (sources.has(source.id)) throw new OfficeMcpError(`Duplicate citation source: ${source.id}`, "DUPLICATE_CITATION_SOURCE");
    sources.set(source.id, source);
  }
  const twips = (inches: number) => Math.round(inches * 1440);
  const footnotes: Record<string, { children: Paragraph[] }> = {};
  const renderedSections = [];
  for (const [index, section] of sections.entries()) {
    const paper = section.pageSize === "a4" ? [11907, 16839] : [12240, 15840];
    const landscape = section.orientation === "landscape";
    const margins = section.marginInches;
    const children = (await Promise.all(section.blocks.map((block) => blockToDocx(block, sources, footnotes)))).flat();
    renderedSections.push({
      properties: {
        type: index === 0 ? undefined : section.breakType === "continuous" ? SectionType.CONTINUOUS : SectionType.NEXT_PAGE,
        page: {
          size: { width: paper[0], height: paper[1], orientation: landscape ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT },
          margin: margins ? {
            top: twips(margins.top ?? 1), right: twips(margins.right ?? 1),
            bottom: twips(margins.bottom ?? 1), left: twips(margins.left ?? 1)
          } : undefined
        }
      },
      headers: section.headerText ? { default: new Header({ children: [new Paragraph(section.headerText)] }) } : undefined,
      footers: section.footerText || section.pageNumbers ? { default: new Footer({
        children: [new Paragraph({ children: [
          ...(section.footerText ? [new TextRun(section.footerText)] : []),
          ...(section.footerText && section.pageNumbers ? [new TextRun("  •  ")] : []),
          ...(section.pageNumbers ? [new TextRun("Page "), new TextRun({ children: [PageNumber.CURRENT] })] : [])
        ] })]
      }) } : undefined,
      children
    });
  }
  const document = new Document({
    creator: spec.author ?? "Office MCP",
    title: spec.title,
    description: "Created with Office MCP",
    features: { updateFields: true },
    footnotes,
    numbering: {
      config: [
        {
          reference: "office-numbering",
          levels: Array.from({ length: 6 }, (_, level) => ({
            level,
            format: LevelFormat.DECIMAL,
            text: `%${level + 1}.`,
            alignment: AlignmentType.START,
            style: { paragraph: { indent: { left: 720 * (level + 1), hanging: 360 } } }
          }))
        }
      ]
    },
    sections: renderedSections
  });

  await writeFile(outputPath, await Packer.toBuffer(document));
  return { path: outputPath, blockCount: sections.reduce((total, section) => total + section.blocks.length, 0), sectionCount: sections.length };
}

async function loadDocumentXml(filePath: string) {
  const resolvedPath = await resolveReadablePath(filePath, WORD_EXTENSIONS);
  const zip = await JSZip.loadAsync(await readFile(resolvedPath));
  const documentEntry = zip.file("word/document.xml");
  if (!documentEntry) throw new OfficeMcpError("The DOCX package has no word/document.xml", "INVALID_DOCX");
  return { resolvedPath, zip, xml: await documentEntry.async("string") };
}

export async function inspectDocument(filePath: string, maxParagraphs = 200) {
  const { resolvedPath, zip, xml } = await loadDocumentXml(filePath);
  const paragraphMatches = [...xml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)];
  const paragraphs = paragraphMatches
    .map((match, index) => {
      const paragraphXml = match[0];
      const style = /<w:pStyle\b[^>]*w:val="([^"]+)"/.exec(paragraphXml)?.[1] ?? null;
      return { index, style, text: textFromXmlFragment(paragraphXml) };
    })
    .filter((paragraph) => paragraph.text.length > 0)
    .slice(0, maxParagraphs);

  const tables = [...xml.matchAll(/<w:tbl\b[\s\S]*?<\/w:tbl>/g)].map((tableMatch, tableIndex) => {
    const rows = [...tableMatch[0].matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map((rowMatch) =>
      [...rowMatch[0].matchAll(/<w:tc\b[\s\S]*?<\/w:tc>/g)].map((cellMatch) =>
        textFromXmlFragment(cellMatch[0])
      )
    );
    return { index: tableIndex, rows };
  });

  const coreXml = await zip.file("docProps/core.xml")?.async("string");
  const property = (name: string) =>
    coreXml ? decodeXml(new RegExp(`<${name}[^>]*>([\\s\\S]*?)<\\/${name}>`).exec(coreXml)?.[1] ?? "") || null : null;

  return {
    kind: "word",
    path: resolvedPath,
    filename: path.basename(resolvedPath),
    title: property("dc:title"),
    creator: property("dc:creator"),
    modified: property("dcterms:modified"),
    paragraphCount: paragraphMatches.length,
    tableCount: tables.length,
    paragraphs,
    tables,
    truncated: paragraphs.length < paragraphMatches.filter((match) => textFromXmlFragment(match[0])).length
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
  const result = text.replace(new RegExp(escaped, "gi"), () => {
    count += 1;
    return replacement;
  });
  return { text: result, count };
}

export async function replaceDocumentText(
  filePath: string,
  find: string,
  replacement: string,
  outputPath?: string,
  overwrite = false,
  matchCase = true
) {
  const { input, output } = resolveInputOutput(filePath, outputPath, WORD_EXTENSIONS);
  if (input !== output) await prepareOutputPath(output, WORD_EXTENSIONS, overwrite);
  const { zip, xml } = await loadDocumentXml(input);
  let replacements = 0;

  const updated = xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (paragraphXml) => {
    const textTags = [...paragraphXml.matchAll(/<w:t\b[^>]*>[\s\S]*?<\/w:t>/g)];
    if (!textTags.length) return paragraphXml;
    const currentText = textFromXmlFragment(paragraphXml);
    const result = replaceLiteral(currentText, find, replacement, matchCase);
    if (!result.count) return paragraphXml;
    replacements += result.count;
    let first = true;
    return paragraphXml.replace(/<w:t\b([^>]*)>[\s\S]*?<\/w:t>/g, (_, attributes: string) => {
      if (!first) return `<w:t${attributes}></w:t>`;
      first = false;
      const preserved = attributes.includes("xml:space") ? attributes : `${attributes} xml:space="preserve"`;
      return `<w:t${preserved}>${encodeXml(result.text)}</w:t>`;
    });
  });

  if (!replacements) {
    throw new OfficeMcpError(`Text was not found: ${find}`, "TEXT_NOT_FOUND");
  }

  zip.file("word/document.xml", updated);
  await writeFile(output, await zip.generateAsync({ type: "nodebuffer" }));
  return { path: output, replacements };
}

export async function appendDocumentParagraph(
  filePath: string,
  text: string,
  style?: string,
  outputPath?: string,
  overwrite = false
) {
  const { input, output } = resolveInputOutput(filePath, outputPath, WORD_EXTENSIONS);
  if (input !== output) await prepareOutputPath(output, WORD_EXTENSIONS, overwrite);
  const { zip, xml } = await loadDocumentXml(input);
  const styleXml = style ? `<w:pPr><w:pStyle w:val="${encodeXml(style)}"/></w:pPr>` : "";
  const paragraphXml = `<w:p>${styleXml}<w:r><w:t xml:space="preserve">${encodeXml(text)}</w:t></w:r></w:p>`;
  const insertionPoint = xml.lastIndexOf("<w:sectPr");
  const bodyEnd = xml.lastIndexOf("</w:body>");
  const index = insertionPoint >= 0 ? insertionPoint : bodyEnd;
  if (index < 0) throw new OfficeMcpError("The DOCX document body is malformed", "INVALID_DOCX");
  const updated = `${xml.slice(0, index)}${paragraphXml}${xml.slice(index)}`;

  zip.file("word/document.xml", updated);
  await writeFile(output, await zip.generateAsync({ type: "nodebuffer" }));
  return { path: output, appended: true };
}
