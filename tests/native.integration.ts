import assert from "node:assert/strict";
import path from "node:path";
import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import JSZip from "jszip";
import { after, before, test } from "node:test";
import { nativeStatus, renderPowerPointNative, runNativeBatch } from "../src/native.js";
import { inspectWorkbook } from "../src/excel.js";
import { inspectDocument } from "../src/word.js";
import { inspectPresentation } from "../src/powerpoint.js";
import { createDesignedPresentation } from "../src/powerpoint-design.js";
import { createFromTemplate, inspectTemplate } from "../src/powerpoint-template.js";
import { exportOfficePdf } from "../src/office.js";
import { checkOfficeQuality, finalizeOfficeFile, prepareOfficeReview, readReviewImage, type QualityContract } from "../src/quality.js";

let testDirectory: string;

before(async () => {
  const workDirectory = path.join(process.cwd(), "work");
  await mkdir(workDirectory, { recursive: true });
  testDirectory = await mkdtemp(path.join(workDirectory, "office-native-test-"));
});

after(async () => {
  if (testDirectory && testDirectory.startsWith(path.join(process.cwd(), "work"))) {
    await rm(testDirectory, { recursive: true, force: true });
  }
});

test("native Microsoft Office automation creates and renders real Office files", { timeout: 240000 }, async (context) => {
  if (process.platform !== "win32") {
    context.skip("Native Office automation is Windows-only");
    return;
  }

  const status = await nativeStatus() as { applications: Array<{ application: string; installed: boolean }> };
  if (!status.applications.every((application) => application.installed)) {
    context.skip("Excel, Word, and PowerPoint are not all installed");
    return;
  }

  const workbookPath = path.join(testDirectory, "native.xlsx");
  await runNativeBatch("excel", {
    target: { create: true },
    operations: [
      { op: "add_sheet", name: "Data" },
      { op: "set_values", sheet: "Data", startCell: "A1", values: [["Month", "Revenue"], ["Jan", 100], ["Feb", 125]] },
      { op: "set_formula", sheet: "Data", range: "B4", formula: "=SUM(B2:B3)" },
      { op: "format_range", sheet: "Data", range: "A1:B1", format: { bold: true, fillColor: "1F4E78", fontColor: "FFFFFF", autoFitColumns: true } },
      { op: "add_table", sheet: "Data", range: "A1:B3", name: "RevenueTable", style: "TableStyleMedium2" },
      { op: "define_name", name: "RevenueValues", sheet: "Data", range: "B2:B3" },
      { op: "set_conditional_format", sheet: "Data", range: "B2:B3", kind: "cellValue", operator: "greater", formula1: "110", fillColor: "C6EFCE" },
      { op: "set_validation", sheet: "Data", range: "A2:A3", kind: "list", formula1: "Jan,Feb,Mar", errorMessage: "Choose a known month" },
      { op: "add_chart", sheet: "Data", sourceRange: "A1:B3", chartType: "column", title: "Revenue", left: 250, top: 20, width: 450, height: 260 },
      { op: "add_sheet", name: "Pivot" },
      {
        op: "create_pivot",
        sourceSheet: "Data",
        sourceRange: "A1:B3",
        destinationSheet: "Pivot",
        destinationCell: "A1",
        name: "RevenuePivot",
        rows: ["Month"],
        values: [{ field: "Revenue", aggregation: "sum", numberFormat: "#,##0" }]
      },
      { op: "recalculate" },
      { op: "save", outputPath: workbookPath }
    ]
  });
  const workbook = await inspectWorkbook(workbookPath);
  assert.ok(workbook.sheets.some((sheet) => sheet.name === "Data"));
  const workbookZip = await JSZip.loadAsync(await readFile(workbookPath));
  const workbookXml = await workbookZip.file("xl/workbook.xml")!.async("string");
  const worksheetXml = await Promise.all(Object.keys(workbookZip.files)
    .filter((name) => /^xl\/worksheets\/sheet\d+\.xml$/.test(name))
    .map((name) => workbookZip.file(name)!.async("string")));
  assert.match(workbookXml, /RevenueValues/);
  assert.ok(worksheetXml.some((xml) => xml.includes("conditionalFormatting")));
  assert.ok(worksheetXml.some((xml) => xml.includes("dataValidation")));
  const workbookPdf = await exportOfficePdf(workbookPath, path.join(testDirectory, "native-workbook.pdf"), "native");
  assert.equal(workbookPdf.engine, "native");
  assert.equal((await readFile(workbookPdf.path)).subarray(0, 4).toString(), "%PDF");

  const documentPath = path.join(testDirectory, "native.docx");
  await runNativeBatch("word", {
    target: { create: true },
    operations: [
      { op: "add_heading", text: "Native Office Report", level: 1 },
      { op: "append_text", text: "This document was created through Microsoft Word automation." },
      { op: "add_table", headerRow: true, rows: [["Metric", "Value"], ["Revenue", "225"]] },
      { op: "apply_style", paragraphIndex: 1, style: "Heading 1" },
      { op: "add_comment", paragraphIndex: 1, text: "Review the title." },
      { op: "insert_section", breakType: "nextPage", orientation: "landscape" },
      { op: "add_heading", text: "Appendix", level: 1 },
      { op: "create_toc", position: "start", maxLevel: 2 },
      { op: "update_toc" },
      { op: "save", outputPath: documentPath }
    ]
  });
  const document = await inspectDocument(documentPath);
  assert.ok(document.paragraphs.some((paragraph) => paragraph.text.includes("Native Office Report")));
  const wordState = await runNativeBatch("word", {
    target: { path: documentPath }, operations: [{ op: "inspect" }]
  }) as { results: Array<{ sections: number; tablesOfContents: number; comments: number }> };
  assert.ok(wordState.results[0].sections >= 2);
  assert.equal(wordState.results[0].tablesOfContents, 1);
  assert.equal(wordState.results[0].comments, 1);
  const documentPdf = await exportOfficePdf(documentPath, path.join(testDirectory, "native-document.pdf"), "native");
  assert.equal((await readFile(documentPdf.path)).subarray(0, 4).toString(), "%PDF");

  const presentationPath = path.join(testDirectory, "native.pptx");
  const renderDirectory = path.join(testDirectory, "rendered");
  await runNativeBatch("powerpoint", {
    target: { create: true },
    operations: [
      { op: "add_slide", layout: 12 },
      { op: "add_text", slide: 1, text: "Native Office Presentation", left: 40, top: 35, width: 600, height: 60, fontSize: 28, bold: true, color: "1F4E78" },
      { op: "add_shape", slide: 1, shape: "roundRectangle", left: 40, top: 130, width: 300, height: 100, fillColor: "D9EAF7", text: "Rendered by PowerPoint" },
      { op: "set_speaker_notes", slide: 1, text: "Explain the native rendering workflow." },
      { op: "save", outputPath: presentationPath }
    ]
  });
  const rendered = await renderPowerPointNative(presentationPath, renderDirectory, 1280, 720);
  assert.equal(rendered.images.length, 1);
  assert.ok(rendered.images[0].data.length > 1000);
  const presentation = await inspectPresentation(presentationPath);
  assert.equal(presentation.slideCount, 1);
  assert.equal(presentation.slides[0].title, "Native Office Presentation");
  const presentationPdf = await exportOfficePdf(presentationPath, path.join(testDirectory, "native-presentation.pdf"), "native");
  assert.equal((await readFile(presentationPdf.path)).subarray(0, 4).toString(), "%PDF");

  const audioPath = path.join(testDirectory, "sample.wav");
  const audio = Buffer.alloc(8044);
  audio.write("RIFF", 0);
  audio.writeUInt32LE(8036, 4);
  audio.write("WAVEfmt ", 8);
  audio.writeUInt32LE(16, 16);
  audio.writeUInt16LE(1, 20);
  audio.writeUInt16LE(1, 22);
  audio.writeUInt32LE(8000, 24);
  audio.writeUInt32LE(8000, 28);
  audio.writeUInt16LE(1, 32);
  audio.writeUInt16LE(8, 34);
  audio.write("data", 36);
  audio.writeUInt32LE(8000, 40);
  audio.fill(128, 44);
  await writeFile(audioPath, audio);
  const mediaDeckPath = path.join(testDirectory, "with-audio.pptx");
  await runNativeBatch("powerpoint", {
    target: { path: presentationPath }, operations: [
      { op: "add_media", slide: 1, path: audioPath, left: 50, top: 250, width: 80, height: 40 },
      { op: "save", outputPath: mediaDeckPath }
    ]
  });
  const mediaZip = await JSZip.loadAsync(await readFile(mediaDeckPath));
  assert.ok(Object.keys(mediaZip.files).some((name) => name.startsWith("ppt/media/") && name.endsWith(".wav")));

  const designedPath = path.join(testDirectory, "designed-native.pptx");
  const designedRenderDirectory = path.join(testDirectory, "designed-rendered");
  const imagePath = path.join(testDirectory, "design-pixel.png");
  await writeFile(
    imagePath,
    Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAF/gL+AvzZAAAAAElFTkSuQmCC", "base64")
  );
  await createDesignedPresentation({
    path: designedPath,
    slides: [
      { layout: "cover", title: "A clearer opening", subtitle: "Testing the editorial composition", imagePath },
      { layout: "statement", title: "One main idea", body: "The layout remains readable when rendered by PowerPoint." }
    ]
  });
  const designedRender = await renderPowerPointNative(designedPath, designedRenderDirectory, 1280, 720);
  assert.equal(designedRender.images.length, 2);
  assert.ok(designedRender.images.every((item) => item.data.length > 1000));
  const hiddenInspection = await runNativeBatch("powerpoint", {
    target: { path: designedPath },
    operations: [{ op: "inspect" }]
  }) as { results: Array<{ presentationWindowCount: number }> };
  assert.equal(hiddenInspection.results[0].presentationWindowCount, 0);

  const templateInspection = await inspectTemplate(designedPath) as {
    designs: Array<{ index: number; layouts: Array<{ index: number }> }>;
    presentationWindowCount: number;
  };
  assert.equal(templateInspection.presentationWindowCount, 0);
  assert.ok(templateInspection.designs[0]?.layouts.length);
  const templateShapes = await runNativeBatch("powerpoint", {
    target: { path: designedPath }, operations: [{ op: "list_shapes", slide: 1 }]
  }) as { results: Array<{ shapes: Array<{ name: string; text: string }> }> };
  const titleShape = templateShapes.results[0].shapes.find((shape) => shape.text.includes("A clearer opening"));
  assert.ok(titleShape);
  const templatedPath = path.join(testDirectory, "from-template.pptx");
  await createFromTemplate({
    templatePath: designedPath,
    outputPath: templatedPath,
    slides: [
      { source: "duplicate", sourceSlide: 1, shapeText: [{ name: titleShape.name, text: "Template-based title" }] },
      {
        source: "layout", designIndex: templateInspection.designs[0].index,
        layoutIndex: templateInspection.designs[0].layouts[0].index,
        textBoxes: [{ text: "New slide from master layout", left: 50, top: 50, width: 500, height: 60, fontSize: 28 }]
      }
    ]
  });
  const templated = await inspectPresentation(templatedPath);
  assert.equal(templated.slideCount, 2);
  assert.equal(templated.slides[0].title, "Template-based title");
  assert.equal((await inspectPresentation(designedPath)).slideCount, 2);
  const placeholderListing = await runNativeBatch("powerpoint", {
    target: { path: templatedPath }, operations: [{ op: "list_placeholders", slide: 2 }]
  }) as { results: Array<{ placeholders: Array<{ index: number; hasTextFrame: boolean }> }> };
  assert.ok(Array.isArray(placeholderListing.results[0].placeholders));
  const fillable = placeholderListing.results[0].placeholders.find((item) => item.hasTextFrame);
  if (fillable) {
    const filledPath = path.join(testDirectory, "from-template-filled.pptx");
    await runNativeBatch("powerpoint", {
      target: { path: templatedPath },
      operations: [
        { op: "set_placeholder_text", slide: 2, placeholderIndex: fillable.index, text: "Filled template placeholder" },
        { op: "save", outputPath: filledPath }
      ]
    });
    assert.ok((await inspectPresentation(filledPath)).slides[1].text.includes("Filled template placeholder"));
  }

  const morphPath = path.join(testDirectory, "morph-native.pptx");
  const morphBatch = await runNativeBatch("powerpoint", {
    target: { path: designedPath },
    operations: [
      { op: "list_shapes", slide: 1 },
      { op: "rename_shape", slide: 1, shape: 1, name: "!!MorphHero" },
      { op: "duplicate_slide", slide: 1, toIndex: 2 },
      { op: "update_shape", slide: 2, name: "!!MorphHero", left: 180, top: 90, width: 420, height: 120, fillTransparency: 0.1, lineTransparency: 0.2 },
      { op: "set_z_order", slide: 2, shape: "!!MorphHero", action: "bringToFront" },
      { op: "set_transition", slide: 2, effect: "morph", durationSeconds: 1.25, advanceOnClick: true },
      { op: "save", outputPath: morphPath }
    ]
  }) as { results: Array<Record<string, unknown>> };
  assert.equal(morphBatch.results[1].name, "!!MorphHero");
  assert.equal(morphBatch.results[5].effect, "morph");
  const reopenedMorph = await runNativeBatch("powerpoint", {
    target: { path: morphPath },
    operations: [{ op: "list_shapes", slide: 1 }, { op: "list_shapes", slide: 2 }, { op: "list_animations", slide: 2 }]
  }) as { results: Array<Record<string, unknown>> };
  assert.ok((reopenedMorph.results[0].shapes as Array<{ name: string }>).some((shape) => shape.name === "!!MorphHero"));
  assert.ok((reopenedMorph.results[1].shapes as Array<{ name: string }>).some((shape) => shape.name === "!!MorphHero"));
  assert.equal((reopenedMorph.results[2].transition as { effectId: number }).effectId, 3954);

  const groupingBatch = await runNativeBatch("powerpoint", {
    target: { path: designedPath },
    operations: [
      { op: "group_shapes", slide: 1, shapes: [1, 2], name: "TemporaryGroup" },
      { op: "set_z_order", slide: 1, shape: "TemporaryGroup", action: "sendToBack" },
      { op: "ungroup_shape", slide: 1, shape: "TemporaryGroup" }
    ]
  }) as { results: Array<Record<string, unknown>> };
  assert.equal(groupingBatch.results[0].shape, "TemporaryGroup");
  assert.equal(groupingBatch.results[0].itemCount, 2);
  assert.equal((groupingBatch.results[2].shapes as Array<unknown>).length, 2);

  const animatedPath = path.join(testDirectory, "animated-native.pptx");
  const animationBatch = await runNativeBatch("powerpoint", {
    target: { path: designedPath },
    operations: [
      { op: "list_shapes", slide: 1 },
      { op: "add_animation", slide: 1, shape: 1, phase: "entrance", effect: "fade", trigger: "afterPrevious", durationSeconds: 0.8, delaySeconds: 0.2 },
      { op: "add_animation", slide: 1, shape: 2, phase: "emphasis", effect: "spin", trigger: "onClick", durationSeconds: 1 },
      { op: "add_animation", slide: 1, shape: 3, phase: "exit", effect: "fade", trigger: "afterPrevious", durationSeconds: 0.5 },
      { op: "set_transition", slide: 2, effect: "fade", durationSeconds: 0.7, advanceOnClick: true, advanceAfterSeconds: 5 },
      { op: "list_animations", slide: 1 },
      { op: "move_animation", slide: 1, animationIndex: 3, toIndex: 1 },
      { op: "update_animation", slide: 1, animationIndex: 1, trigger: "afterPrevious", durationSeconds: 1.2, delaySeconds: 0.3, repeatCount: 2, autoReverse: true },
      { op: "list_animations", slide: 1 },
      { op: "delete_animation", slide: 1, animationIndex: 2 },
      { op: "list_animations", slide: 1 },
      { op: "save", outputPath: animatedPath }
    ]
  }) as { results: Array<Record<string, unknown>> };
  const shapes = animationBatch.results[0].shapes as Array<{ name: string }>;
  assert.ok(shapes.length > 0 && shapes[0].name);
  const animationInfo = animationBatch.results[5].animations as Array<{ effectId: number; shape: string; exit: boolean }>;
  assert.equal(animationInfo.length, 3);
  assert.equal(animationInfo[0].effectId, 10);
  assert.equal(animationInfo[0].shape, shapes[0].name);
  assert.equal(animationInfo[1].effectId, 61);
  assert.equal(animationInfo[2].exit, true);
  const reordered = animationBatch.results[8].animations as Array<{ exit: boolean; durationSeconds: number; delaySeconds: number; trigger: number }>;
  assert.equal(reordered[0].exit, true);
  assert.ok(Math.abs(reordered[0].durationSeconds - 1.2) < 0.001);
  assert.ok(Math.abs(reordered[0].delaySeconds - 0.3) < 0.001);
  assert.equal(reordered[0].trigger, 3);
  assert.equal((animationBatch.results[10].animations as Array<unknown>).length, 2);

  const reopened = await runNativeBatch("powerpoint", {
    target: { path: animatedPath },
    operations: [{ op: "list_animations", slide: 1 }, { op: "list_animations", slide: 2 }, { op: "clear_animations", slide: 1 }, { op: "list_animations", slide: 1 }]
  }) as { results: Array<Record<string, unknown>> };
  assert.equal((reopened.results[0].animations as Array<unknown>).length, 2);
  assert.equal((reopened.results[1].transition as { effectId: number }).effectId, 1793);
  assert.equal((reopened.results[1].transition as { advanceOnTime: boolean }).advanceOnTime, true);
  assert.equal((reopened.results[1].transition as { advanceAfterSeconds: number }).advanceAfterSeconds, 5);
  assert.equal((reopened.results[3].animations as Array<unknown>).length, 0);

  const qualityCases: Array<{ path: string; contract: QualityContract; unit: string }> = [
    {
      path: workbookPath,
      contract: { objective: "Keep the native revenue workbook complete and readable.", criteria: [{ id: "data", description: "The revenue source data remains intact." }], excel: { requiredSheets: ["Data"] } },
      unit: "sheet:Data"
    },
    {
      path: documentPath,
      contract: { objective: "Keep the report heading, appendix, and table intact.", criteria: [{ id: "structure", description: "The report retains its original structure." }], word: { requiredHeadings: ["Native Office Report"], minTables: 1 } },
      unit: "document"
    },
    {
      path: presentationPath,
      contract: { objective: "Keep the native presentation readable and complete.", criteria: [{ id: "slide", description: "The slide is readable and complete." }], powerpoint: { minSlides: 1, maxSlides: 1 } },
      unit: "slide:1"
    }
  ];
  for (const [index, item] of qualityCases.entries()) {
    const report = await checkOfficeQuality(item.path, item.contract);
    assert.equal(report.machineStatus, "needs_review", JSON.stringify(report.issues));
    const prepared = await prepareOfficeReview(item.path, item.contract, path.join(testDirectory, `review-${index}`));
    assert.ok(prepared.artifacts.some((artifact) => artifact.unit === item.unit));
    if (index === 2) {
      const slide = await readReviewImage(prepared.manifestPath, item.unit);
      assert.equal(slide.data.subarray(0, 4).toString("hex"), "89504e47");
      await assert.rejects(() => finalizeOfficeFile(item.path, path.join(testDirectory, "changed-contract.pptx"), {
        ...item.contract, requiredText: ["Native Office Presentation"]
      }, {
        fileSha256: report.fileSha256,
        evidenceManifestPath: prepared.manifestPath,
        units: [{ unit: item.unit, verdict: "pass", note: "Synthetic review note for a changed contract." }],
        criteria: [{ id: "slide", verdict: "pass", note: "Synthetic criterion note for a changed contract." }]
      }, true), /Review evidence does not match/);
      const final = await finalizeOfficeFile(item.path, path.join(testDirectory, "native-reviewed.pptx"), item.contract, {
        fileSha256: report.fileSha256,
        evidenceManifestPath: prepared.manifestPath,
        units: [{ unit: item.unit, verdict: "pass", note: "A native render was generated for this test slide." }],
        criteria: [{ id: "slide", verdict: "pass", note: "This is a synthetic protocol review for testing." }]
      }, true);
      assert.equal(final.fileSha256, report.fileSha256);
    }
  }
});
