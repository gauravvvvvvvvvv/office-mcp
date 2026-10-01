import assert from "node:assert/strict";
import path from "node:path";
import { mkdir, mkdtemp, readFile, rm } from "node:fs/promises";
import { after, before, test } from "node:test";
import { createWorkbook, setFormulas } from "../src/excel.js";
import { createDocument, appendDocumentParagraph } from "../src/word.js";
import { createPresentation } from "../src/powerpoint.js";
import { checkOfficeQuality, finalizeOfficeFile, type QualityContract } from "../src/quality.js";
import { compareOfficeFiles } from "../src/compare.js";

let directory: string;

before(async () => {
  const work = path.join(process.cwd(), "work");
  await mkdir(work, { recursive: true });
  directory = await mkdtemp(path.join(work, "quality-test-"));
});

after(async () => {
  if (directory?.startsWith(path.join(process.cwd(), "work"))) await rm(directory, { recursive: true, force: true });
});

test("Excel quality checks prompt assertions and uncached formulas", async () => {
  const file = path.join(directory, "analysis.xlsx");
  await createWorkbook({ path: file, sheets: [{ name: "Data", data: [["Region", "Revenue"], ["North", 25]] }] });
  const contract: QualityContract = {
    objective: "Summarize regional revenue without losing the source figures.",
    criteria: [{ id: "totals", description: "The revenue values match the supplied source." }],
    requiredText: ["Revenue"],
    excel: {
      requiredSheets: ["Data"],
      requiredHeaders: [{ sheet: "Data", headers: ["Region", "Revenue"] }],
      expectedCells: [{ sheet: "Data", cell: "B2", value: 25 }]
    }
  };
  const good = await checkOfficeQuality(file, contract);
  assert.equal(good.machineStatus, "needs_review");
  assert.deepEqual(good.reviewUnits, ["sheet:Data"]);
  assert.equal(good.issues.length, 0);

  await setFormulas(file, "Data", [{ cell: "B3", formula: "SUM(B2:B2)" }]);
  const stale = await checkOfficeQuality(file, contract);
  assert.equal(stale.machineStatus, "blocked");
  assert.ok(stale.issues.some((item) => item.code === "UNCACHED_FORMULAS"));
});

test("Word finalization requires current hash, every review unit, and prompt criteria", async () => {
  const draft = path.join(directory, "report.docx");
  const final = path.join(directory, "report-final.docx");
  await createDocument({ path: draft, blocks: [
    { type: "heading", text: "Findings", level: 1 },
    { type: "paragraph", text: "Revenue rose in the example data." }
  ] });
  const contract: QualityContract = {
    objective: "Create a concise report with a findings heading and the supplied conclusion.",
    criteria: [{ id: "meaning", description: "The conclusion matches the source evidence." }],
    requiredText: ["Revenue rose"],
    word: { requiredHeadings: ["Findings"] }
  };
  const report = await checkOfficeQuality(draft, contract);
  assert.equal(report.machineStatus, "needs_review");
  await assert.rejects(() => finalizeOfficeFile(draft, final, contract, {
    fileSha256: report.fileSha256, units: [], criteria: []
  }), /Review must cover/);
  const review = {
    fileSha256: report.fileSha256,
    units: [{ unit: "document", verdict: "pass" as const, note: "Reviewed the complete report layout and text." }],
    criteria: [{ id: "meaning", verdict: "pass" as const, note: "The stated conclusion matches the supplied example." }]
  };
  await appendDocumentParagraph(draft, "A later edit changes the draft.");
  await assert.rejects(() => finalizeOfficeFile(draft, final, contract, review), /changed after review/);
  const current = await checkOfficeQuality(draft, contract);
  const result = await finalizeOfficeFile(draft, final, contract, { ...review, fileSha256: current.fileSha256 });
  assert.equal(result.reviewStatus, "attested");
  assert.deepEqual(await readFile(final), await readFile(draft));
  await assert.rejects(() => finalizeOfficeFile(draft, final, contract, { ...review, fileSha256: current.fileSha256 }), /already exists/);
});

test("PowerPoint quality checks structure and each slide review", async () => {
  const draft = path.join(directory, "slides.pptx");
  const final = path.join(directory, "slides-final.pptx");
  await createPresentation({ path: draft, slides: [
    { elements: [{ type: "text", text: "Plan", x: 0.8, y: 0.8, w: 8, h: 1, fontSize: 36 }] },
    { elements: [{ type: "text", text: "Execution", x: 0.8, y: 0.8, w: 8, h: 1, fontSize: 36 }] }
  ] });
  const contract: QualityContract = {
    objective: "Create a two-slide plan with clear titles and no overflow.",
    criteria: [{ id: "story", description: "The two slides form a coherent short narrative." }],
    powerpoint: { minSlides: 2, maxSlides: 2, minimumFontPoints: 17 }
  };
  const report = await checkOfficeQuality(draft, contract);
  assert.equal(report.machineStatus, "needs_review");
  assert.deepEqual(report.reviewUnits, ["slide:1", "slide:2"]);
  await assert.rejects(() => finalizeOfficeFile(draft, final, contract, {
    fileSha256: report.fileSha256,
    units: [{ unit: "slide:1", verdict: "pass", note: "The slide was inspected at full presentation size." }],
    criteria: [{ id: "story", verdict: "pass", note: "Both slides tell the requested short story." }]
  }), /Review must cover/);
  const result = await finalizeOfficeFile(draft, final, contract, {
    fileSha256: report.fileSha256,
    units: [
      { unit: "slide:1", verdict: "pass", note: "The slide was inspected at full presentation size." },
      { unit: "slide:2", verdict: "pass", note: "The slide was inspected at full presentation size." }
    ],
    criteria: [{ id: "story", verdict: "pass", note: "Both slides tell the requested short story." }]
  });
  assert.equal(result.kind, "powerpoint");
});

test("Baseline comparison catches accidental Word content loss", async () => {
  const before = path.join(directory, "source.docx");
  const afterPath = path.join(directory, "edited.docx");
  await createDocument({ path: before, blocks: [
    { type: "heading", text: "Key findings", level: 1 },
    { type: "paragraph", text: "Preserve this important observation." }
  ] });
  await createDocument({ path: afterPath, blocks: [
    { type: "heading", text: "Key findings", level: 1 }
  ] });
  const comparison = await compareOfficeFiles(before, afterPath);
  assert.equal(comparison.kind, "word");
  assert.equal(comparison.summary.removedParagraphs, 1);
  const report = await checkOfficeQuality(afterPath, {
    objective: "Improve this report without losing the existing observations.",
    criteria: [{ id: "preserve", description: "All existing observations remain in the document." }],
    baseline: { path: before }
  });
  assert.equal(report.machineStatus, "blocked");
  assert.ok(report.issues.some((item) => item.code === "BASELINE_PARAGRAPHS_REMOVED"));
});

test("Baseline comparison catches removed Excel sheets", async () => {
  const before = path.join(directory, "source.xlsx");
  const afterPath = path.join(directory, "edited.xlsx");
  await createWorkbook({ path: before, sheets: [
    { name: "Data", data: [["Value"], [3]] },
    { name: "Assumptions", data: [["Rate"], [0.2]] }
  ] });
  await createWorkbook({ path: afterPath, sheets: [{ name: "Data", data: [["Value"], [5]] }] });
  const comparison = await compareOfficeFiles(before, afterPath);
  assert.equal(comparison.kind, "excel");
  assert.deepEqual(comparison.summary.removedSheets, ["Assumptions"]);
  const report = await checkOfficeQuality(afterPath, {
    objective: "Update workbook values while keeping its original sheets.",
    criteria: [{ id: "preserve", description: "All original sheets remain available." }],
    baseline: { path: before }
  });
  assert.equal(report.machineStatus, "blocked");
  assert.ok(report.issues.some((item) => item.code === "BASELINE_SHEETS_REMOVED"));
});
