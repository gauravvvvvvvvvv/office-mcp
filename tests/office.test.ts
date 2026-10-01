import assert from "node:assert/strict";
import path from "node:path";
import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import JSZip from "jszip";
import { after, before, test } from "node:test";
import { createWorkbook, inspectWorkbook, readRange, setFormulas, writeRange } from "../src/excel.js";
import { analyzeDataset, queryDataset } from "../src/excel-analysis.js";
import {
  appendDocumentParagraph,
  createDocument,
  inspectDocument,
  replaceDocumentText
} from "../src/word.js";
import {
  createPresentation,
  inspectPresentation,
  replacePresentationText
} from "../src/powerpoint.js";
import { createDesignedPresentation } from "../src/powerpoint-design.js";
import { auditPresentation } from "../src/powerpoint-audit.js";
import { resolveAllowedPath } from "../src/paths.js";
import { searchOfficeFiles } from "../src/office.js";

let testDirectory: string;

before(async () => {
  const workDirectory = path.join(process.cwd(), "work");
  await mkdir(workDirectory, { recursive: true });
  testDirectory = await mkdtemp(path.join(workDirectory, "office-mcp-test-"));
});

after(async () => {
  if (testDirectory && testDirectory.startsWith(path.join(process.cwd(), "work"))) {
    await rm(testDirectory, { recursive: true, force: true });
  }
});

test("path guard rejects files outside the configured root", () => {
  const outside = path.resolve(process.cwd(), "..", "outside.xlsx");
  assert.throws(() => resolveAllowedPath(outside, [".xlsx"]), /outside the configured/);
});

test("Excel create, inspect, write, formula, and read flow", async () => {
  const workbookPath = path.join(testDirectory, "sample.xlsx");
  await createWorkbook({
    path: workbookPath,
    sheets: [{
      name: "Revenue",
      data: [["Month", "Revenue"], ["Jan", 100], ["Feb", 120]],
      headerRow: true,
      freezeRows: 1,
      autoFilter: true
    }]
  });

  const inspection = await inspectWorkbook(workbookPath);
  assert.equal(inspection.sheetCount, 1);
  assert.equal(inspection.sheets[0].name, "Revenue");

  await writeRange({ path: workbookPath, sheet: "Revenue", startCell: "B4", values: [[140]] });
  await setFormulas(workbookPath, "Revenue", [{ cell: "B5", formula: "SUM(B2:B4)", result: 360 }]);
  const range = await readRange(workbookPath, "Revenue", "A1:B5");
  assert.equal(range.cells[3][1].value, 140);
  assert.equal(range.cells[4][1].formula, "SUM(B2:B4)");
});

test("Excel dataset analysis computes quality and descriptive statistics", async () => {
  const workbookPath = path.join(testDirectory, "analysis.xlsx");
  await createWorkbook({
    path: workbookPath,
    sheets: [{ name: "Data", data: [
      ["Region", "Revenue", "Cost"],
      ["North", 10, 5],
      ["South", 20, 10],
      ["South", 20, 10],
      ["West", 100, 50],
      [null, null, 5]
    ] }]
  });
  const analysis = await analyzeDataset(workbookPath, "Data");
  assert.equal(analysis.rowsAnalyzed, 5);
  assert.equal(analysis.duplicateRows, 1);
  assert.equal(analysis.columns[0].missing, 1);
  assert.equal(analysis.columns[1].numeric?.outlierCount, 1);
  assert.equal(analysis.columns[1].numeric?.median, 20);
  assert.equal(analysis.strongestCorrelations[0].coefficient, 1);
  const query = await queryDataset({
    path: workbookPath, sheet: "Data",
    filters: [{ column: "Revenue", operator: "notBlank" }],
    groupBy: ["Region"],
    metrics: [{ column: "Revenue", aggregation: "sum", as: "totalRevenue" }],
    sort: [{ column: "totalRevenue", direction: "desc" }]
  });
  assert.equal(query.matchedRows, 4);
  assert.equal(query.rows[0].Region, "West");
  assert.equal(query.rows[0].totalRevenue, 100);
  assert.equal(query.rows.find((row) => row.Region === "South")?.totalRevenue, 40);
});

test("Word create, inspect, append, and replace flow", async () => {
  const documentPath = path.join(testDirectory, "sample.docx");
  await createDocument({
    path: documentPath,
    title: "Quarterly Review",
    blocks: [
      { type: "heading", text: "Quarterly Review", level: 1 },
      { type: "paragraph", text: "Revenue increased during the quarter." },
      { type: "table", headerRow: true, rows: [["Month", "Revenue"], ["Jan", "100"]] }
    ]
  });

  await appendDocumentParagraph(documentPath, "Prepared by the finance team.");
  await replaceDocumentText(documentPath, "finance", "strategy");
  const inspection = await inspectDocument(documentPath);
  assert.equal(inspection.title, "Quarterly Review");
  assert.equal(inspection.tableCount, 1);
  assert.ok(inspection.paragraphs.some((paragraph) => paragraph.text.includes("strategy team")));
});

test("Word creates sections, a TOC field, and source-backed citations", async () => {
  const documentPath = path.join(testDirectory, "professional.docx");
  const created = await createDocument({
    path: documentPath,
    sources: [{ id: "smith2025", author: "Jane Smith", title: "Market Outlook", year: "2025", url: "https://example.com/outlook" }],
    sections: [
      {
        pageSize: "a4", headerText: "Strategy Group", footerText: "Confidential", pageNumbers: true,
        blocks: [{ type: "toc" }, { type: "heading", text: "Overview", level: 1 }]
      },
      {
        orientation: "landscape",
        blocks: [
          { type: "heading", text: "Evidence", level: 1 },
          { type: "cited_paragraph", text: "The market is growing", sourceIds: ["smith2025"] },
          { type: "bibliography" }
        ]
      }
    ]
  });
  assert.equal(created.sectionCount, 2);
  const inspected = await inspectDocument(documentPath);
  assert.ok(inspected.paragraphs.some((paragraph) => paragraph.text.includes("Smith, 2025")));
  assert.ok(inspected.paragraphs.some((paragraph) => paragraph.text.includes("Market Outlook")));
  const zip = await JSZip.loadAsync(await readFile(documentPath));
  const documentXml = await zip.file("word/document.xml")!.async("string");
  const settingsXml = await zip.file("word/settings.xml")!.async("string");
  assert.ok(documentXml.includes("TOC \\h \\o"));
  assert.match(documentXml, /w:orient="landscape"/);
  assert.match(documentXml, /w:w="15840" w:h="12240" w:orient="landscape"/);
  assert.match(settingsXml, /w:updateFields/);
});

test("Word embeds described images, links, page breaks, and footnotes", async () => {
  const documentPath = path.join(testDirectory, "rich.docx");
  const imagePath = path.join(testDirectory, "word-pixel.png");
  await writeFile(imagePath, Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAF/gL+AvzZAAAAAElFTkSuQmCC", "base64"));
  await createDocument({
    path: documentPath,
    blocks: [
      { type: "image", path: imagePath, widthPx: 100, heightPx: 100, altText: "A sample data point", caption: "Figure 1" },
      { type: "hyperlink", text: "Source", url: "https://example.com/source" },
      { type: "page_break" },
      { type: "footnote_paragraph", text: "A supported claim", footnoteText: "Supporting note" }
    ]
  });
  const zip = await JSZip.loadAsync(await readFile(documentPath));
  const documentXml = await zip.file("word/document.xml")!.async("string");
  const footnoteXml = await zip.file("word/footnotes.xml")!.async("string");
  assert.match(documentXml, /A sample data point/);
  assert.match(documentXml, /w:br w:type="page"/);
  assert.match(documentXml, /w:footnoteReference/);
  assert.match(footnoteXml, /Supporting note/);
  assert.ok(Object.keys(zip.files).some((file) => file.startsWith("word/media/")));
});

test("PowerPoint create, inspect, and replace flow", async () => {
  const presentationPath = path.join(testDirectory, "sample.pptx");
  const imagePath = path.join(testDirectory, "pixel.png");
  await writeFile(
    imagePath,
    Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAF/gL+AvzZAAAAAElFTkSuQmCC", "base64")
  );
  await createPresentation({
    path: presentationPath,
    title: "Quarterly Review",
    slides: [
      {
        speakerNotes: "Explain the quarter's highlights.",
        elements: [
          { type: "text", text: "Quarterly Review", x: 0.7, y: 0.5, w: 8, h: 0.6, fontSize: 28, bold: true },
          { type: "image", path: imagePath, x: 11.5, y: 0.4, w: 0.4, h: 0.4 },
          {
            type: "chart",
            chartType: "bar",
            x: 0.7,
            y: 1.4,
            w: 6,
            h: 4,
            series: [{ name: "Revenue", labels: ["Jan", "Feb"], values: [100, 120] }]
          }
        ]
      }
    ]
  });

  let inspection = await inspectPresentation(presentationPath);
  assert.equal(inspection.slideCount, 1);
  assert.equal(inspection.slides[0].title, "Quarterly Review");
  await replacePresentationText(presentationPath, "Quarterly", "Annual");
  inspection = await inspectPresentation(presentationPath);
  assert.equal(inspection.slides[0].title, "Annual Review");
});

test("Designed PowerPoint composes readable narrative layouts with image cropping", async () => {
  const presentationPath = path.join(testDirectory, "designed.pptx");
  const imagePath = path.join(testDirectory, "design-pixel.png");
  await writeFile(
    imagePath,
    Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAF/gL+AvzZAAAAAElFTkSuQmCC", "base64")
  );
  const panoramaImagePath = path.join(testDirectory, "panorama-pixel.png");
  await writeFile(panoramaImagePath, await readFile(imagePath));
  const result = await createDesignedPresentation({
    path: presentationPath,
    style: "editorial",
    slides: [
      { layout: "cover", title: "A clear idea", subtitle: "A short opening statement", imagePath },
      { layout: "statement", title: "One decisive point", body: "A concise explanation." },
      { layout: "comparison", title: "Two approaches", left: { heading: "Current", body: "The current state" }, right: { heading: "Proposed", body: "The proposed state" }, takeaway: "The proposed approach is clearer" },
      { layout: "panorama", title: "A wider view", body: "The narrative continues", imagePath: panoramaImagePath, imageAltText: "A test image" }
    ]
  });
  assert.equal(result.slideCount, 4);
  assert.deepEqual(result.designWarnings, []);
  const inspection = await inspectPresentation(presentationPath);
  assert.equal(inspection.slideCount, 4);
  assert.equal(inspection.slides[0].title, "A clear idea");
  assert.equal(inspection.slides[2].title, "Two approaches");
  assert.equal(inspection.slides[3].title, "A wider view");
});

test("PowerPoint audit flags reviewable slide defects without opening Office", async () => {
  const presentationPath = path.join(testDirectory, "audit.pptx");
  const imagePath = path.join(testDirectory, "audit-pixel.png");
  await writeFile(
    imagePath,
    Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAF/gL+AvzZAAAAAElFTkSuQmCC", "base64")
  );
  await createPresentation({
    path: presentationPath,
    slides: [
      { elements: [
        { type: "text", text: "Small text", x: 0.5, y: 0.5, w: 2, h: 0.5, fontSize: 10 },
        { type: "text", text: "Outside", x: 12.8, y: 1, w: 2, h: 0.5, fontSize: 24 },
        { type: "image", path: imagePath, x: 2, y: 2, w: 1, h: 1 }
      ] },
      { backgroundColor: "101010", elements: [
        { type: "text", text: "Hard to see", x: 1, y: 1, w: 4, h: 1, fontSize: 24, color: "202020" },
        { type: "text", text: "Overlapping text", x: 2, y: 1.2, w: 4, h: 1, fontSize: 24, color: "FFFFFF" }
      ] }
    ]
  });
  const audit = await auditPresentation(presentationPath);
  assert.equal(audit.slideCount, 2);
  assert.ok(audit.issues.some((issue) => issue.code === "SMALL_TEXT"));
  assert.ok(audit.issues.some((issue) => issue.code === "OFF_SLIDE"));
  assert.ok(audit.issues.some((issue) => issue.code === "MISSING_ALT_TEXT"));
  assert.ok(audit.issues.some((issue) => issue.code === "LOW_CONTRAST"));
  assert.ok(audit.issues.some((issue) => issue.code === "TEXT_OVERLAP"));
  assert.ok(audit.limitations.some((item) => item.includes("visual")));
});

test("Office file discovery finds supported formats", async () => {
  const result = await searchOfficeFiles(testDirectory, true);
  assert.ok(result.files.some((file) => file.extension === ".xlsx"));
  assert.ok(result.files.some((file) => file.extension === ".docx"));
  assert.ok(result.files.some((file) => file.extension === ".pptx"));
});
