#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/server";
import { serveStdio } from "@modelcontextprotocol/server/stdio";
import { readFile } from "node:fs/promises";
import * as z from "zod/v4";
import { jsonResult, toErrorResult } from "./errors.js";
import { createWorkbook, inspectWorkbook, readRange, setFormulas, writeRange } from "./excel.js";
import { analyzeDataset, queryDataset } from "./excel-analysis.js";
import {
  appendDocumentParagraph,
  createDocument,
  inspectDocument,
  replaceDocumentText
} from "./word.js";
import {
  createPresentation,
  inspectPresentation,
  replacePresentationText
} from "./powerpoint.js";
import { createDesignedPresentation } from "./powerpoint-design.js";
import { auditPresentation } from "./powerpoint-audit.js";
import { createFromTemplate, inspectTemplate } from "./powerpoint-template.js";
import { capabilities, convertToPdf, exportOfficePdf, inspectOfficeFile, searchOfficeFiles } from "./office.js";
import { listOpenOfficeFiles, nativeStatus, renderPowerPointNative, runNativeBatch } from "./native.js";
import { checkOfficeQuality, finalizeOfficeFile, prepareOfficeReview, readReviewImage } from "./quality.js";
import { compareOfficeFiles } from "./compare.js";
import { OFFICE_MCP_VERSION } from "./version.js";

const filePath = z.string().min(1).describe("Absolute path or path relative to the server working directory");
const outputOptions = {
  outputPath: filePath.optional().describe("Optional output path. Omit to update the input file in place."),
  overwrite: z.boolean().optional().default(false)
};
const readOnlyAnnotations = { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false };
const writeAnnotations = { readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: false };

const qualityContractSchema = z.object({
  objective: z.string().min(10).describe("The actual user request summarized without changing its intent"),
  criteria: z.array(z.object({ id: z.string().regex(/^[A-Za-z0-9_-]+$/), description: z.string().min(8) })).min(1)
    .refine((items) => new Set(items.map((item) => item.id)).size === items.length, "Criterion IDs must be unique"),
  requiredText: z.array(z.string().min(1)).optional(),
  forbiddenText: z.array(z.string().min(1)).optional(),
  baseline: z.object({
    path: filePath,
    maxRemovedSheets: z.number().int().nonnegative().optional(),
    maxRemovedSlides: z.number().int().nonnegative().optional(),
    maxRemovedParagraphs: z.number().int().nonnegative().optional(),
    maxChangedCells: z.number().int().nonnegative().optional()
  }).optional(),
  excel: z.object({
    requiredSheets: z.array(z.string().min(1)).optional(),
    requiredHeaders: z.array(z.object({ sheet: z.string().min(1), row: z.number().int().positive().optional(), headers: z.array(z.string().min(1)).min(1) })).optional(),
    expectedCells: z.array(z.object({ sheet: z.string().min(1), cell: z.string().regex(/^[A-Za-z]+[1-9]\d*$/), value: z.union([z.string(), z.number(), z.boolean(), z.null()]), tolerance: z.number().nonnegative().optional() })).optional(),
    maxFormulaErrors: z.number().int().nonnegative().optional(),
    maxUncachedFormulas: z.number().int().nonnegative().optional()
  }).optional(),
  word: z.object({
    requiredHeadings: z.array(z.string().min(1)).optional(),
    minTables: z.number().int().nonnegative().optional(),
    requireImageAltText: z.boolean().optional()
  }).optional(),
  powerpoint: z.object({
    minSlides: z.number().int().positive().optional(),
    maxSlides: z.number().int().positive().optional(),
    minimumFontPoints: z.number().min(6).max(36).optional(),
    requireSpeakerNotes: z.boolean().optional()
  }).optional()
});

const qualityReviewSchema = z.object({
  fileSha256: z.string().regex(/^[a-f0-9]{64}$/i),
  evidenceManifestPath: filePath,
  units: z.array(z.object({ unit: z.string().min(1), verdict: z.enum(["pass", "fail"]), note: z.string().min(12) })),
  criteria: z.array(z.object({ id: z.string().min(1), verdict: z.enum(["pass", "fail"]), note: z.string().min(12) })),
  acceptedWarnings: z.array(z.object({ code: z.string().min(1), reason: z.string().min(12) })).optional()
});

function tool<T>(handler: (input: T) => Promise<unknown>) {
  return async (input: T) => {
    try {
      return jsonResult(await handler(input));
    } catch (error) {
      console.error(error);
      return toErrorResult(error);
    }
  };
}

const wordBlockSchema = z.discriminatedUnion("type", [
  z.object({ type: z.literal("heading"), text: z.string(), level: z.number().int().min(1).max(6).optional() }),
  z.object({
    type: z.literal("paragraph"),
    text: z.string(),
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    align: z.enum(["left", "center", "right", "justify"]).optional()
  }),
  z.object({ type: z.literal("bullet"), text: z.string(), level: z.number().int().min(0).max(8).optional() }),
  z.object({ type: z.literal("number"), text: z.string(), level: z.number().int().min(0).max(5).optional() }),
  z.object({ type: z.literal("table"), rows: z.array(z.array(z.string())).min(1), headerRow: z.boolean().optional() }),
  z.object({ type: z.literal("toc"), maxLevel: z.number().int().min(1).max(6).optional() }),
  z.object({ type: z.literal("cited_paragraph"), text: z.string(), sourceIds: z.array(z.string().min(1)).min(1) }),
  z.object({ type: z.literal("bibliography"), title: z.string().optional() }),
  z.object({ type: z.literal("image"), path: filePath, widthPx: z.number().positive(), heightPx: z.number().positive(), altText: z.string().min(1).optional(), decorative: z.boolean().optional(), caption: z.string().optional() }),
  z.object({ type: z.literal("hyperlink"), text: z.string().min(1), url: z.string().url() }),
  z.object({ type: z.literal("page_break") }),
  z.object({ type: z.literal("footnote_paragraph"), text: z.string(), footnoteText: z.string().min(1) })
]);

const wordSectionSchema = z.object({
  blocks: z.array(wordBlockSchema),
  pageSize: z.enum(["letter", "a4"]).optional(),
  orientation: z.enum(["portrait", "landscape"]).optional(),
  marginInches: z.object({
    top: z.number().min(0.25).max(3).optional(),
    right: z.number().min(0.25).max(3).optional(),
    bottom: z.number().min(0.25).max(3).optional(),
    left: z.number().min(0.25).max(3).optional()
  }).optional(),
  headerText: z.string().optional(),
  footerText: z.string().optional(),
  pageNumbers: z.boolean().optional(),
  breakType: z.enum(["nextPage", "continuous"]).optional()
});

const positionSchema = {
  x: z.number().nonnegative(),
  y: z.number().nonnegative(),
  w: z.number().positive(),
  h: z.number().positive()
};

const presentationElementSchema = z.discriminatedUnion("type", [
  z.object({
    type: z.literal("text"),
    ...positionSchema,
    text: z.string(),
    fontSize: z.number().positive().optional(),
    fontFace: z.string().optional(),
    color: z.string().optional(),
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    align: z.enum(["left", "center", "right", "justify"]).optional(),
    valign: z.enum(["top", "middle", "bottom"]).optional(),
    margin: z.number().nonnegative().optional()
  }),
  z.object({
    type: z.literal("shape"),
    ...positionSchema,
    shape: z.enum(["rect", "roundRect", "ellipse", "line"]).optional(),
    fillColor: z.string().optional(),
    lineColor: z.string().optional(),
    lineWidth: z.number().nonnegative().optional(),
    text: z.string().optional()
  }),
  z.object({
    type: z.literal("image"),
    ...positionSchema,
    path: filePath,
    altText: z.string().min(1).optional().describe("Meaningful description for an informative image; omit for an undecided/decorative image and review the audit warning"),
    transparency: z.number().min(0).max(100).optional(),
    fit: z.enum(["cover", "contain", "stretch"]).optional().describe("cover crops to the frame; contain preserves the full image; stretch may distort it")
  }),
  z.object({
    type: z.literal("table"),
    ...positionSchema,
    rows: z.array(z.array(z.string())).min(1),
    fontSize: z.number().positive().optional(),
    headerFill: z.string().optional(),
    borderColor: z.string().optional()
  }),
  z.object({
    type: z.literal("chart"),
    ...positionSchema,
    chartType: z.enum(["bar", "line", "pie", "doughnut", "area"]),
    series: z.array(z.object({
      name: z.string(),
      labels: z.array(z.string()),
      values: z.array(z.number())
    })).min(1),
    title: z.string().optional(),
    showLegend: z.boolean().optional(),
    showValue: z.boolean().optional(),
    showCategoryName: z.boolean().optional()
  })
]);

const nativeTargetSchema = z.object({
  path: filePath.optional().describe("Existing Office file to open or attach to"),
  active: z.boolean().optional().default(false).describe("Opt in to controlling the user's running Office application or active document; may affect its visible window"),
  create: z.boolean().optional().default(false).describe("Create a new blank file in the native Office application"),
  visible: z.boolean().optional().default(false).describe("Show a native Office window when creating or opening a file; false keeps file-based PowerPoint work hidden")
});

const excelNativeOperationSchema = z.discriminatedUnion("op", [
  z.object({ op: z.literal("inspect") }),
  z.object({ op: z.literal("read_range"), sheet: z.string(), range: z.string() }),
  z.object({ op: z.literal("set_values"), sheet: z.string(), startCell: z.string(), values: z.array(z.array(z.unknown())) }),
  z.object({ op: z.literal("set_formula"), sheet: z.string(), range: z.string(), formula: z.string() }),
  z.object({
    op: z.literal("format_range"),
    sheet: z.string(),
    range: z.string(),
    format: z.object({
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
      fontSize: z.number().positive().optional(),
      fontColor: z.string().optional(),
      fillColor: z.string().optional(),
      numberFormat: z.string().optional(),
      wrapText: z.boolean().optional(),
      horizontalAlignment: z.enum(["left", "center", "right", "general"]).optional(),
      merge: z.boolean().optional(),
      autoFitColumns: z.boolean().optional(),
      autoFitRows: z.boolean().optional()
    })
  }),
  z.object({ op: z.literal("add_sheet"), name: z.string() }),
  z.object({ op: z.literal("define_name"), name: z.string().regex(/^[A-Za-z_][A-Za-z0-9_.]*$/), sheet: z.string(), range: z.string() }),
  z.object({
    op: z.literal("set_conditional_format"), sheet: z.string(), range: z.string(),
    kind: z.enum(["cellValue", "formula"]),
    operator: z.enum(["between", "notBetween", "equal", "notEqual", "greater", "less", "greaterEqual", "lessEqual"]).optional(),
    formula1: z.string(), formula2: z.string().optional(),
    fillColor: z.string().optional(), fontColor: z.string().optional()
  }),
  z.object({
    op: z.literal("set_validation"), sheet: z.string(), range: z.string(),
    kind: z.enum(["list", "whole", "decimal", "date", "time", "textLength", "custom"]),
    operator: z.enum(["between", "notBetween", "equal", "notEqual", "greater", "less", "greaterEqual", "lessEqual"]).optional(),
    formula1: z.string().max(255), formula2: z.string().max(255).optional(), errorMessage: z.string().max(255).optional()
  }),
  z.object({ op: z.literal("rename_sheet"), sheet: z.string(), newName: z.string() }),
  z.object({ op: z.literal("delete_sheet"), sheet: z.string() }),
  z.object({ op: z.literal("add_table"), sheet: z.string(), range: z.string(), name: z.string(), style: z.string().optional() }),
  z.object({
    op: z.literal("add_chart"),
    sheet: z.string(),
    sourceRange: z.string(),
    chartType: z.enum(["column", "bar", "line", "pie", "doughnut", "area", "scatter"]),
    title: z.string().optional(),
    left: z.number(), top: z.number(), width: z.number().positive(), height: z.number().positive()
  }),
  z.object({
    op: z.literal("create_pivot"),
    sourceSheet: z.string(),
    sourceRange: z.string(),
    destinationSheet: z.string(),
    destinationCell: z.string(),
    name: z.string(),
    rows: z.array(z.string()).optional(),
    columns: z.array(z.string()).optional(),
    filters: z.array(z.string()).optional(),
    values: z.array(z.object({
      field: z.string(),
      aggregation: z.enum(["sum", "count", "average", "max", "min"]).optional(),
      caption: z.string().optional(),
      numberFormat: z.string().optional()
    })).min(1)
  }),
  z.object({ op: z.literal("sort"), sheet: z.string(), range: z.string(), key: z.string(), ascending: z.boolean().optional() }),
  z.object({ op: z.literal("autofilter"), sheet: z.string(), range: z.string(), field: z.number().int().positive(), criteria: z.unknown().optional() }),
  z.object({ op: z.literal("recalculate") }),
  z.object({ op: z.literal("save"), outputPath: filePath.optional() }),
  z.object({ op: z.literal("export_pdf"), outputPath: filePath })
]);

const wordNativeOperationSchema = z.discriminatedUnion("op", [
  z.object({ op: z.literal("inspect") }),
  z.object({ op: z.literal("read_text"), maxCharacters: z.number().int().positive().optional() }),
  z.object({ op: z.literal("apply_style"), paragraphIndex: z.number().int().positive(), style: z.string().min(1) }),
  z.object({ op: z.literal("insert_section"), breakType: z.enum(["nextPage", "continuous"]).optional().default("nextPage"), orientation: z.enum(["portrait", "landscape"]).optional() }),
  z.object({ op: z.literal("create_toc"), position: z.enum(["start", "end"]).optional().default("start"), maxLevel: z.number().int().min(1).max(9).optional().default(3) }),
  z.object({ op: z.literal("update_toc") }),
  z.object({ op: z.literal("add_comment"), paragraphIndex: z.number().int().positive(), text: z.string().min(1) }),
  z.object({ op: z.literal("append_text"), text: z.string(), newParagraph: z.boolean().optional(), style: z.string().optional() }),
  z.object({ op: z.literal("add_heading"), text: z.string(), level: z.number().int().min(1).max(9).optional() }),
  z.object({ op: z.literal("replace_text"), find: z.string().min(1), replacement: z.string(), matchCase: z.boolean().optional() }),
  z.object({ op: z.literal("add_table"), rows: z.array(z.array(z.string())).min(1), headerRow: z.boolean().optional(), style: z.string().optional() }),
  z.object({ op: z.literal("set_track_changes"), enabled: z.boolean() }),
  z.object({ op: z.literal("save"), outputPath: filePath.optional() }),
  z.object({ op: z.literal("export_pdf"), outputPath: filePath })
]);

const powerpointNativeOperationSchema = z.discriminatedUnion("op", [
  z.object({ op: z.literal("inspect") }),
  z.object({ op: z.literal("list_shapes"), slide: z.number().int().positive() }),
  z.object({ op: z.literal("list_placeholders"), slide: z.number().int().positive() }),
  z.object({ op: z.literal("set_placeholder_text"), slide: z.number().int().positive(), placeholderIndex: z.number().int().positive(), text: z.string() }),
  z.object({ op: z.literal("list_animations"), slide: z.number().int().positive() }),
  z.object({ op: z.literal("add_slide"), layout: z.number().int().optional().describe("PowerPoint PpSlideLayout value; 12 is blank") }),
  z.object({ op: z.literal("add_slide_from_layout"), designIndex: z.number().int().positive(), layoutIndex: z.number().int().positive() }),
  z.object({ op: z.literal("duplicate_slide"), slide: z.number().int().positive(), toIndex: z.number().int().positive().optional() }),
  z.object({ op: z.literal("move_slide"), slide: z.number().int().positive(), toIndex: z.number().int().positive() }),
  z.object({
    op: z.literal("add_text"), slide: z.number().int().positive(), text: z.string(),
    left: z.number(), top: z.number(), width: z.number().positive(), height: z.number().positive(),
    fontSize: z.number().positive().optional(), fontFace: z.string().optional(), color: z.string().optional(), bold: z.boolean().optional()
  }),
  z.object({
    op: z.literal("add_shape"), slide: z.number().int().positive(),
    shape: z.enum(["rectangle", "roundRectangle", "ellipse", "triangle", "diamond", "chevron"]),
    left: z.number(), top: z.number(), width: z.number().positive(), height: z.number().positive(),
    fillColor: z.string().optional(), lineColor: z.string().optional(), text: z.string().optional()
  }),
  z.object({
    op: z.literal("add_picture"), slide: z.number().int().positive(), path: filePath,
    left: z.number(), top: z.number(), width: z.number().positive(), height: z.number().positive(),
    altText: z.string().optional()
  }),
  z.object({
    op: z.literal("add_media"), slide: z.number().int().positive(), path: filePath,
    left: z.number(), top: z.number(), width: z.number().positive(), height: z.number().positive()
  }),
  z.object({
    op: z.literal("rename_shape"), slide: z.number().int().positive(),
    shape: z.union([z.string().min(1), z.number().int().positive()]).describe("Current shape name or 1-based index from list_shapes"),
    name: z.string().min(1).max(255).describe("New unique name. Prefix with !! to force Morph matching across slides.")
  }),
  z.object({
    op: z.literal("group_shapes"), slide: z.number().int().positive(),
    shapes: z.array(z.union([z.string().min(1), z.number().int().positive()])).min(2),
    name: z.string().min(1).max(255).optional()
  }),
  z.object({
    op: z.literal("ungroup_shape"), slide: z.number().int().positive(),
    shape: z.union([z.string().min(1), z.number().int().positive()])
  }),
  z.object({
    op: z.literal("set_z_order"), slide: z.number().int().positive(),
    shape: z.union([z.string().min(1), z.number().int().positive()]),
    action: z.enum(["bringToFront", "bringForward", "sendBackward", "sendToBack"])
  }),
  z.object({
    op: z.literal("update_shape"), slide: z.number().int().positive(), name: z.string(),
    left: z.number().optional(), top: z.number().optional(), width: z.number().positive().optional(), height: z.number().positive().optional(),
    text: z.string().optional(), fillColor: z.string().optional(), lineColor: z.string().optional(),
    fontSize: z.number().positive().optional(), fontColor: z.string().optional(), altText: z.string().optional(),
    rotation: z.number().min(-360).max(360).optional(),
    fillTransparency: z.number().min(0).max(1).optional(), lineTransparency: z.number().min(0).max(1).optional()
  }),
  z.object({ op: z.literal("set_speaker_notes"), slide: z.number().int().positive(), text: z.string() }),
  z.object({
    op: z.literal("add_animation"),
    slide: z.number().int().positive(),
    shape: z.union([z.string().min(1), z.number().int().positive()]).describe("Shape name or 1-based index from list_shapes"),
    phase: z.enum(["entrance", "emphasis", "exit"]).optional().default("entrance"),
    effect: z.enum(["appear", "fade", "fly", "wipe", "zoom", "spin", "growShrink"]),
    trigger: z.enum(["onClick", "withPrevious", "afterPrevious"]).optional().default("onClick"),
    durationSeconds: z.number().min(0.1).max(30).optional().default(0.6),
    delaySeconds: z.number().min(0).max(60).optional().default(0),
    repeatCount: z.number().int().min(1).max(100).optional(),
    autoReverse: z.boolean().optional(),
    position: z.number().int().positive().optional()
  }),
  z.object({
    op: z.literal("update_animation"), slide: z.number().int().positive(), animationIndex: z.number().int().positive(),
    trigger: z.enum(["onClick", "withPrevious", "afterPrevious"]).optional(),
    durationSeconds: z.number().min(0.1).max(30).optional(),
    delaySeconds: z.number().min(0).max(60).optional(),
    repeatCount: z.number().int().min(1).max(100).optional(),
    autoReverse: z.boolean().optional()
  }),
  z.object({ op: z.literal("move_animation"), slide: z.number().int().positive(), animationIndex: z.number().int().positive(), toIndex: z.number().int().positive() }),
  z.object({ op: z.literal("delete_animation"), slide: z.number().int().positive(), animationIndex: z.number().int().positive() }),
  z.object({ op: z.literal("clear_animations"), slide: z.number().int().positive() }),
  z.object({
    op: z.literal("set_transition"),
    slide: z.number().int().positive(),
    effect: z.enum(["none", "cut", "fade", "pushLeft", "pushRight", "wipeLeft", "wipeRight", "zoomIn", "morph", "morphWords", "morphCharacters"]),
    durationSeconds: z.number().min(0.1).max(10).optional(),
    advanceOnClick: z.boolean().optional().default(true),
    advanceAfterSeconds: z.number().min(0.1).max(3600).optional()
  }),
  z.object({ op: z.literal("replace_text"), find: z.string().min(1), replacement: z.string(), matchCase: z.boolean().optional() }),
  z.object({ op: z.literal("delete_slide"), slide: z.number().int().positive() }),
  z.object({ op: z.literal("save"), outputPath: filePath.optional() }),
  z.object({ op: z.literal("export_pdf"), outputPath: filePath }),
  z.object({
    op: z.literal("render_slides"), outputDirectory: filePath,
    width: z.number().int().positive().optional(), height: z.number().int().positive().optional()
  }),
  z.object({
    op: z.literal("export_video"), outputPath: filePath,
    useTimingsAndNarrations: z.boolean().optional().default(true),
    defaultSlideDurationSeconds: z.number().int().min(1).max(60).optional().default(5),
    verticalResolution: z.enum(["480", "720", "1080", "2160"]).optional().default("1080"),
    framesPerSecond: z.number().int().min(12).max(60).optional().default(30),
    quality: z.number().int().min(1).max(100).optional().default(85),
    timeoutSeconds: z.number().int().min(30).max(900).optional().default(600)
  })
]);

const designedSlideSchema = z.discriminatedUnion("layout", [
  z.object({
    layout: z.literal("cover"),
    title: z.string().min(1).max(65),
    subtitle: z.string().max(180).optional(),
    imagePath: filePath.optional(),
    imageAltText: z.string().min(1).optional(),
    speakerNotes: z.string().optional()
  }),
  z.object({
    layout: z.literal("imageText"),
    title: z.string().min(1).max(60),
    body: z.string().min(1).max(170),
    imagePath: filePath,
    imageAltText: z.string().min(1).optional(),
    imageSide: z.enum(["left", "right"]).optional(),
    speakerNotes: z.string().optional()
  }),
  z.object({
    layout: z.literal("statement"),
    title: z.string().min(1).max(80),
    body: z.string().max(180).optional(),
    speakerNotes: z.string().optional()
  }),
  z.object({
    layout: z.literal("comparison"),
    title: z.string().min(1).max(70),
    left: z.object({ heading: z.string().min(1).max(55), body: z.string().min(1).max(180) }),
    right: z.object({ heading: z.string().min(1).max(55), body: z.string().min(1).max(180) }),
    takeaway: z.string().max(130).optional(),
    speakerNotes: z.string().optional()
  }),
  z.object({
    layout: z.literal("panorama"),
    title: z.string().min(1).max(70),
    body: z.string().max(120).optional(),
    imagePath: filePath,
    imageAltText: z.string().min(1).optional(),
    speakerNotes: z.string().optional()
  })
]);

serveStdio(() => {
  const server = new McpServer(
    { name: "office-mcp", version: OFFICE_MCP_VERSION },
    {
      instructions:
        "Use office_capabilities when unsure what is supported. Discover and inspect relevant files before editing; use office_compare_files and a quality contract baseline to detect accidental losses. For a new narrative deck, plan one point per slide, prefer powerpoint_create_designed_presentation, and supply relevant, distinct images for image-led layouts. Use panorama for a wide scene and comparison.takeaway for the conclusion of a contrast; avoid repetitive text-only slides. Use powerpoint_inspect_template then powerpoint_create_from_template for an existing brand deck; use powerpoint_create_presentation for precise custom layouts. File-based native PowerPoint work uses hidden presentation windows by default; never set target.active=true or target.visible=true merely to bypass a background-session refusal without user consent. Before delivery, call office_quality_check with prompt-specific assertions, office_prepare_review, inspect every rendered slide or the complete PDF, repair issues, then office_finalize with honest per-unit and per-criterion review notes. Static renders do not show motion. Use excel_analyze_dataset for deterministic profiling and excel_query_dataset for read-only filters and aggregations, but independently reason about the business question and do not infer causation from correlation. Prefer a new outputPath unless the user explicitly wants an in-place edit. All paths must be inside OFFICE_MCP_ROOTS."
    }
  );

  server.registerTool(
    "office_capabilities",
    {
      description: "List supported Office operations, path roots, converters, and known limitations.",
      annotations: readOnlyAnnotations
    },
    tool(async () => capabilities())
  );

  server.registerTool(
    "office_inspect_file",
    {
      description: "Inspect an XLSX, DOCX, or PPTX file and return a semantic summary.",
      inputSchema: z.object({ path: filePath, sampleLimit: z.number().int().min(1).max(500).optional().default(20) }),
      annotations: readOnlyAnnotations
    },
    tool(async ({ path, sampleLimit }) => inspectOfficeFile(path, sampleLimit))
  );

  server.registerTool(
    "office_search_files",
    {
      description: "Find XLSX, DOCX, and PPTX files inside an allowed directory so an orchestrator can discover relevant Office documents.",
      inputSchema: z.object({
        directory: filePath,
        recursive: z.boolean().optional().default(true),
        extensions: z.array(z.enum([".xlsx", ".docx", ".pptx"])).optional(),
        maxResults: z.number().int().min(1).max(5000).optional().default(500)
      }),
      annotations: readOnlyAnnotations
    },
    tool(async ({ directory, recursive, extensions, maxResults }) =>
      searchOfficeFiles(directory, recursive, extensions, maxResults)
    )
  );

  server.registerTool(
    "office_convert_to_pdf",
    {
      description: "Convert an XLSX, DOCX, or PPTX file to PDF using LibreOffice when available.",
      inputSchema: z.object({ path: filePath, outputPath: filePath.optional(), overwrite: z.boolean().optional().default(false) }),
      annotations: writeAnnotations
    },
    tool(async ({ path, outputPath, overwrite }) => convertToPdf(path, outputPath, overwrite))
  );

  server.registerTool(
    "office_export_pdf",
    {
      description: "Export XLSX, DOCX, or PPTX to a reviewable PDF using hidden native Microsoft Office on Windows, with optional LibreOffice fallback. This creates a separate file and does not open a visible window.",
      inputSchema: z.object({
        path: filePath,
        outputPath: filePath.optional(),
        engine: z.enum(["auto", "native", "libreoffice"]).optional().default("auto"),
        overwrite: z.boolean().optional().default(false)
      }),
      annotations: writeAnnotations
    },
    tool(async ({ path, outputPath, engine, overwrite }) => exportOfficePdf(path, outputPath, engine, overwrite))
  );

  server.registerTool(
    "office_quality_check",
    {
      description: "Validate an XLSX, DOCX, or PPTX draft against prompt-specific requirements and Office package integrity. Returns the file hash, blocking issues, warnings, and exact units requiring visual or analytical review. Machine checks alone never certify quality.",
      inputSchema: z.object({ path: filePath, contract: qualityContractSchema }),
      annotations: readOnlyAnnotations
    },
    tool(async ({ path, contract }) => checkOfficeQuality(path, contract))
  );

  server.registerTool(
    "office_compare_files",
    {
      description: "Compare two Office files of the same format semantically. Reports removed Word paragraphs, PowerPoint slide content changes, or Excel sheet and cell changes without modifying either file.",
      inputSchema: z.object({
        beforePath: filePath,
        afterPath: filePath,
        maxChanges: z.number().int().min(1).max(1000).optional().default(100)
      }),
      annotations: readOnlyAnnotations
    },
    tool(async ({ beforePath, afterPath, maxChanges }) => compareOfficeFiles(beforePath, afterPath, maxChanges))
  );

  server.registerTool(
    "office_prepare_review",
    {
      description: "Create hashed review evidence for a quality-checked draft: PowerPoint slide PNGs or a Word/Excel PDF, with a source hash and artifact manifest. PowerPoint renders use hidden windows. Inspect every slide or the complete PDF before attesting quality.",
      inputSchema: z.object({
        path: filePath,
        contract: qualityContractSchema,
        outputDirectory: filePath,
        maxInlineImages: z.number().int().min(0).max(50).optional().default(12)
      }),
      annotations: writeAnnotations
    },
    async ({ path, contract, outputDirectory, maxInlineImages }) => {
      try {
        const prepared = await prepareOfficeReview(path, contract, outputDirectory);
        const images = prepared.kind === "powerpoint"
          ? await Promise.all(prepared.artifacts.slice(0, maxInlineImages).map(async (artifact) => ({
            type: "image" as const,
            data: (await readFile(artifact.path)).toString("base64"),
            mimeType: "image/png" as const
          })))
          : [];
        return {
          content: [
            { type: "text" as const, text: JSON.stringify({
              ...prepared, returnedImages: images.length,
              truncatedImages: prepared.kind === "powerpoint" && images.length < prepared.artifacts.length
            }, null, 2) },
            ...images
          ],
          structuredContent: prepared
        };
      } catch (error) {
        console.error(error);
        return toErrorResult(error);
      }
    }
  );

  server.registerTool(
    "office_finalize",
    {
      description: "Publish a separate XLSX, DOCX, or PPTX only after recomputing machine checks and verifying a complete review of every slide/sheet/document and every prompt criterion against the unchanged file hash. Never overwrites an existing output.",
      inputSchema: z.object({ path: filePath, outputPath: filePath, contract: qualityContractSchema, review: qualityReviewSchema }),
      annotations: writeAnnotations
    },
    tool(async ({ path, outputPath, contract, review }) => finalizeOfficeFile(path, outputPath, contract, review, true))
  );

  server.registerTool(
    "office_get_review_image",
    {
      description: "Return one previously rendered PowerPoint slide as MCP image content, including slides omitted from the initial inline preview. Rejects stale or changed evidence.",
      inputSchema: z.object({ manifestPath: filePath, unit: z.string().regex(/^slide:[1-9]\d*$/) }),
      annotations: readOnlyAnnotations
    },
    async ({ manifestPath, unit }) => {
      try {
        const image = await readReviewImage(manifestPath, unit);
        return {
          content: [
            { type: "text" as const, text: JSON.stringify({ unit: image.unit, imagePath: image.imagePath }) },
            { type: "image" as const, data: image.data.toString("base64"), mimeType: "image/png" as const }
          ]
        };
      } catch (error) {
        console.error(error);
        return toErrorResult(error);
      }
    }
  );

  server.registerTool(
    "excel_create_workbook",
    {
      description: "Create a new XLSX workbook with one or more worksheets and tabular data.",
      inputSchema: z.object({
        path: filePath,
        overwrite: z.boolean().optional().default(false),
        creator: z.string().optional(),
        sheets: z.array(z.object({
          name: z.string().min(1).max(31),
          data: z.array(z.array(z.unknown())).optional(),
          headerRow: z.boolean().optional(),
          freezeRows: z.number().int().nonnegative().optional(),
          columnWidths: z.array(z.number().positive()).optional(),
          autoFilter: z.boolean().optional()
        })).min(1)
      }),
      annotations: writeAnnotations
    },
    tool(createWorkbook)
  );

  server.registerTool(
    "excel_inspect_workbook",
    {
      description: "Inspect workbook metadata, worksheets, dimensions, formulas, and sample rows.",
      inputSchema: z.object({ path: filePath, sampleRows: z.number().int().min(0).max(100).optional().default(5) }),
      annotations: readOnlyAnnotations
    },
    tool(async ({ path, sampleRows }) => inspectWorkbook(path, sampleRows))
  );

  server.registerTool(
    "excel_read_range",
    {
      description: "Read cell values, text, formulas, and optional styles from an XLSX range.",
      inputSchema: z.object({
        path: filePath,
        sheet: z.string().min(1),
        range: z.string().regex(/^[A-Za-z]+[1-9]\d*(?::[A-Za-z]+[1-9]\d*)?$/),
        includeStyles: z.boolean().optional().default(false)
      }),
      annotations: readOnlyAnnotations
    },
    tool(async ({ path, sheet, range, includeStyles }) => readRange(path, sheet, range, includeStyles))
  );

  server.registerTool(
    "excel_analyze_dataset",
    {
      description: "Profile an XLSX dataset without changing it: column types, missing values, numeric summaries, duplicate rows, IQR outliers, top categories, and Pearson correlations. Descriptive analysis only; formulas need cached results or native recalculation.",
      inputSchema: z.object({
        path: filePath,
        sheet: z.string().min(1),
        range: z.string().regex(/^[A-Za-z]+[1-9]\d*:[A-Za-z]+[1-9]\d*$/).optional(),
        hasHeader: z.boolean().optional().default(true),
        maxRows: z.number().int().min(1).max(50000).optional().default(10000),
        topValues: z.number().int().min(1).max(20).optional().default(5)
      }),
      annotations: readOnlyAnnotations
    },
    tool(async ({ path, sheet, range, hasHeader, maxRows, topValues }) =>
      analyzeDataset(path, sheet, range, hasHeader, maxRows, topValues)
    )
  );

  server.registerTool(
    "excel_query_dataset",
    {
      description: "Read-only Excel dataset query with typed filters, group-by summaries, aggregations, sorting, and bounded output. Use for focused analytical questions without changing the workbook.",
      inputSchema: z.object({
        path: filePath,
        sheet: z.string().min(1),
        range: z.string().regex(/^[A-Za-z]+[1-9]\d*:[A-Za-z]+[1-9]\d*$/).optional(),
        hasHeader: z.boolean().optional().default(true),
        filters: z.array(z.object({
          column: z.string().min(1),
          operator: z.enum(["eq", "neq", "gt", "gte", "lt", "lte", "contains", "in", "isBlank", "notBlank"]),
          value: z.union([z.string(), z.number(), z.boolean(), z.null()]).optional(),
          values: z.array(z.union([z.string(), z.number(), z.boolean(), z.null()])).optional()
        })).optional(),
        groupBy: z.array(z.string().min(1)).max(10).optional(),
        metrics: z.array(z.object({
          column: z.string().min(1),
          aggregation: z.enum(["count", "countDistinct", "sum", "average", "min", "max"]),
          as: z.string().min(1).optional()
        })).max(20).optional(),
        sort: z.array(z.object({ column: z.string().min(1), direction: z.enum(["asc", "desc"]) })).optional(),
        maxRows: z.number().int().min(1).max(50000).optional().default(10000),
        limit: z.number().int().min(1).max(1000).optional().default(100)
      }),
      annotations: readOnlyAnnotations
    },
    tool(queryDataset)
  );

  server.registerTool(
    "excel_write_range",
    {
      description: "Write a rectangular matrix of values to an existing XLSX worksheet.",
      inputSchema: z.object({
        path: filePath,
        sheet: z.string().min(1),
        startCell: z.string().regex(/^[A-Za-z]+[1-9]\d*$/),
        values: z.array(z.array(z.unknown())).min(1),
        ...outputOptions,
        numberFormat: z.string().optional(),
        bold: z.boolean().optional(),
        fillColor: z.string().optional()
      }),
      annotations: writeAnnotations
    },
    tool(writeRange)
  );

  server.registerTool(
    "excel_set_formulas",
    {
      description: "Set formulas in specific XLSX cells and request recalculation when opened in Excel.",
      inputSchema: z.object({
        path: filePath,
        sheet: z.string().min(1),
        formulas: z.array(z.object({
          cell: z.string().regex(/^[A-Za-z]+[1-9]\d*$/),
          formula: z.string().min(1),
          result: z.union([z.string(), z.number(), z.boolean(), z.null()]).optional()
        })).min(1),
        ...outputOptions
      }),
      annotations: writeAnnotations
    },
    tool(async ({ path, sheet, formulas, outputPath, overwrite }) =>
      setFormulas(path, sheet, formulas, outputPath, overwrite)
    )
  );

  server.registerTool(
    "word_create_document",
    {
      description: "Create a DOCX from sections, headings, paragraphs, lists, tables, a Word-updateable TOC, and source-backed citation text with a bibliography.",
      inputSchema: z.object({
        path: filePath,
        title: z.string().optional(),
        author: z.string().optional(),
        blocks: z.array(wordBlockSchema).optional().describe("Simple single-section document; use sections instead for page layout and headers/footers"),
        sections: z.array(wordSectionSchema).min(1).optional(),
        sources: z.array(z.object({
          id: z.string().min(1), author: z.string().min(1), title: z.string().min(1),
          year: z.string().min(1), url: z.string().url().optional()
        })).optional().describe("Source metadata for cited_paragraph and bibliography blocks"),
        overwrite: z.boolean().optional().default(false)
      }),
      annotations: writeAnnotations
    },
    tool(createDocument)
  );

  server.registerTool(
    "word_inspect_document",
    {
      description: "Inspect DOCX metadata, paragraphs, paragraph styles, and tables.",
      inputSchema: z.object({ path: filePath, maxParagraphs: z.number().int().min(1).max(1000).optional().default(200) }),
      annotations: readOnlyAnnotations
    },
    tool(async ({ path, maxParagraphs }) => inspectDocument(path, maxParagraphs))
  );

  server.registerTool(
    "word_append_paragraph",
    {
      description: "Append a paragraph to an existing DOCX document while preserving its package.",
      inputSchema: z.object({ path: filePath, text: z.string(), style: z.string().optional(), ...outputOptions }),
      annotations: writeAnnotations
    },
    tool(async ({ path, text, style, outputPath, overwrite }) =>
      appendDocumentParagraph(path, text, style, outputPath, overwrite)
    )
  );

  server.registerTool(
    "word_replace_text",
    {
      description: "Replace literal text in DOCX paragraphs, including text split across multiple runs.",
      inputSchema: z.object({
        path: filePath,
        find: z.string().min(1),
        replacement: z.string(),
        matchCase: z.boolean().optional().default(true),
        ...outputOptions
      }),
      annotations: writeAnnotations
    },
    tool(async ({ path, find, replacement, matchCase, outputPath, overwrite }) =>
      replaceDocumentText(path, find, replacement, outputPath, overwrite, matchCase)
    )
  );

  server.registerTool(
    "powerpoint_create_presentation",
    {
      description: "Create a PPTX presentation with positioned text, shapes, images, tables, charts, and notes.",
      inputSchema: z.object({
        path: filePath,
        layout: z.enum(["wide", "standard"]).optional().default("wide"),
        title: z.string().optional(),
        subject: z.string().optional(),
        author: z.string().optional(),
        company: z.string().optional(),
        theme: z.object({
          headFontFace: z.string().optional(),
          bodyFontFace: z.string().optional(),
          language: z.string().optional()
        }).optional(),
        slides: z.array(z.object({
          backgroundColor: z.string().optional(),
          speakerNotes: z.string().optional(),
          elements: z.array(presentationElementSchema)
        })).min(1),
        overwrite: z.boolean().optional().default(false)
      }),
      annotations: writeAnnotations
    },
    tool(createPresentation)
  );

  server.registerTool(
    "powerpoint_create_designed_presentation",
    {
      description: "Create a widescreen narrative deck from editorial compositions with readable typography and optional cropped imagery. Prefer this for new story decks; render and critique the result before delivery.",
      inputSchema: z.object({
        path: filePath,
        title: z.string().optional(),
        style: z.enum(["editorial", "cinematic"]).optional().default("editorial"),
        slides: z.array(designedSlideSchema).min(1),
        overwrite: z.boolean().optional().default(false)
      }),
      annotations: writeAnnotations
    },
    tool(createDesignedPresentation)
  );

  server.registerTool(
    "powerpoint_inspect_presentation",
    {
      description: "Inspect PPTX slide size, text, notes, and object counts.",
      inputSchema: z.object({ path: filePath, maxSlides: z.number().int().min(1).max(1000).optional().default(200) }),
      annotations: readOnlyAnnotations
    },
    tool(async ({ path, maxSlides }) => inspectPresentation(path, maxSlides))
  );

  server.registerTool(
    "powerpoint_audit_presentation",
    {
      description: "Read-only structural PPTX audit for off-slide objects, small explicit text, dense slides, pictures without alt text, and other review flags. This is heuristic; render and visually inspect the deck as well.",
      inputSchema: z.object({
        path: filePath,
        minimumFontPoints: z.number().min(6).max(36).optional().default(14),
        denseSlideCharacters: z.number().int().min(100).max(3000).optional().default(600)
      }),
      annotations: readOnlyAnnotations
    },
    tool(async ({ path, minimumFontPoints, denseSlideCharacters }) =>
      auditPresentation(path, minimumFontPoints, denseSlideCharacters)
    )
  );

  server.registerTool(
    "powerpoint_inspect_template",
    {
      description: "Inspect an existing PPTX through hidden desktop PowerPoint to list designs, custom layouts and placeholder indexes for template-based generation. Requires Windows PowerPoint and refuses a running user session by default.",
      inputSchema: z.object({ templatePath: filePath }),
      annotations: readOnlyAnnotations
    },
    tool(async ({ templatePath }) => inspectTemplate(templatePath))
  );

  server.registerTool(
    "powerpoint_create_from_template",
    {
      description: "Create a new PPTX from an existing deck's masters/layouts or duplicated slides, filling placeholders or named text shapes. Native PowerPoint runs with hidden presentation windows and leaves the source template unchanged.",
      inputSchema: z.object({
        templatePath: filePath,
        outputPath: filePath,
        slides: z.array(z.discriminatedUnion("source", [
          z.object({
            source: z.literal("layout"), designIndex: z.number().int().positive(), layoutIndex: z.number().int().positive(),
            placeholderText: z.array(z.object({ index: z.number().int().positive(), text: z.string() })).optional(),
            shapeText: z.array(z.object({ name: z.string().min(1), text: z.string() })).optional(),
            textBoxes: z.array(z.object({
              text: z.string(), left: z.number(), top: z.number(), width: z.number().positive(), height: z.number().positive(),
              fontSize: z.number().positive().optional(), fontFace: z.string().optional(), color: z.string().optional(), bold: z.boolean().optional()
            })).optional(),
            speakerNotes: z.string().optional()
          }),
          z.object({
            source: z.literal("duplicate"), sourceSlide: z.number().int().positive(),
            placeholderText: z.array(z.object({ index: z.number().int().positive(), text: z.string() })).optional(),
            shapeText: z.array(z.object({ name: z.string().min(1), text: z.string() })).optional(),
            textBoxes: z.array(z.object({
              text: z.string(), left: z.number(), top: z.number(), width: z.number().positive(), height: z.number().positive(),
              fontSize: z.number().positive().optional(), fontFace: z.string().optional(), color: z.string().optional(), bold: z.boolean().optional()
            })).optional(),
            speakerNotes: z.string().optional()
          })
        ])).min(1),
        keepTemplateSlides: z.boolean().optional().default(false),
        overwrite: z.boolean().optional().default(false)
      }),
      annotations: writeAnnotations
    },
    tool(createFromTemplate)
  );

  server.registerTool(
    "powerpoint_replace_text",
    {
      description: "Replace literal text in PPTX slides, including text split across multiple runs.",
      inputSchema: z.object({
        path: filePath,
        find: z.string().min(1),
        replacement: z.string(),
        matchCase: z.boolean().optional().default(true),
        ...outputOptions
      }),
      annotations: writeAnnotations
    },
    tool(async ({ path, find, replacement, matchCase, outputPath, overwrite }) =>
      replacePresentationText(path, find, replacement, outputPath, overwrite, matchCase)
    )
  );

  server.registerTool(
    "office_native_status",
    {
      description: "Detect installed and currently running native Microsoft Excel, Word, and PowerPoint applications on Windows.",
      annotations: readOnlyAnnotations
    },
    tool(async () => nativeStatus())
  );

  server.registerTool(
    "office_list_open_files",
    {
      description: "List files currently open in running native Excel, Word, and PowerPoint applications.",
      annotations: readOnlyAnnotations
    },
    tool(async () => listOpenOfficeFiles())
  );

  server.registerTool(
    "excel_native_batch",
    {
      description: "Execute an ordered batch through real Microsoft Excel. Supports active or closed workbooks, native calculation, formatting, tables, charts, filtering, sorting, PDF export, and save/save-as.",
      inputSchema: z.object({
        target: nativeTargetSchema,
        operations: z.array(excelNativeOperationSchema).min(1),
        overwrite: z.boolean().optional().default(false)
      }),
      annotations: writeAnnotations
    },
    tool(async (input) => runNativeBatch("excel", input))
  );

  server.registerTool(
    "word_native_batch",
    {
      description: "Execute an ordered batch through real Microsoft Word. Supports active or closed documents, inspection, editing, tables, styles, tracked changes, PDF export, and save/save-as.",
      inputSchema: z.object({
        target: nativeTargetSchema,
        operations: z.array(wordNativeOperationSchema).min(1),
        overwrite: z.boolean().optional().default(false)
      }),
      annotations: writeAnnotations
    },
    tool(async (input) => runNativeBatch("word", input))
  );

  server.registerTool(
    "powerpoint_native_batch",
    {
      description: "Execute an ordered batch through real Microsoft PowerPoint. Supports active or closed presentations, slide and shape editing, grouping and z-order, object animations, true Morph transitions, PDF/video export, and PNG rendering. Use list_shapes before targeting shapes. Prefix paired object names with !! for deterministic Morph matching. Native coordinates are points.",
      inputSchema: z.object({
        target: nativeTargetSchema,
        operations: z.array(powerpointNativeOperationSchema).min(1),
        overwrite: z.boolean().optional().default(false)
      }),
      annotations: writeAnnotations
    },
    tool(async (input) => runNativeBatch("powerpoint", input))
  );

  server.registerTool(
    "powerpoint_render_presentation",
    {
      description: "Render a PPTX with real Microsoft PowerPoint and return slide PNGs as MCP images for visual inspection and iterative correction.",
      inputSchema: z.object({
        path: filePath,
        outputDirectory: filePath,
        width: z.number().int().min(320).max(4096).optional().default(1600),
        height: z.number().int().min(180).max(4096).optional().default(900),
        maxSlides: z.number().int().min(1).max(50).optional().default(10),
        overwrite: z.boolean().optional().default(false)
      }),
      annotations: writeAnnotations
    },
    async ({ path, outputDirectory, width, height, maxSlides, overwrite }) => {
      try {
        const rendered = await renderPowerPointNative(path, outputDirectory, width, height, overwrite);
        const returned = rendered.images.slice(0, maxSlides);
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify({
                outputDirectory: rendered.outputDirectory,
                slideCount: rendered.images.length,
                returnedImages: returned.length,
                truncated: returned.length < rendered.images.length,
                paths: rendered.images.map((item) => item.path)
              }, null, 2)
            },
            ...returned.map((item) => ({ type: "image" as const, data: item.data, mimeType: "image/png" as const }))
          ],
          structuredContent: {
            outputDirectory: rendered.outputDirectory,
            slideCount: rendered.images.length,
            paths: rendered.images.map((item) => item.path)
          }
        };
      } catch (error) {
        console.error(error);
        return toErrorResult(error);
      }
    }
  );

  return server;
});
