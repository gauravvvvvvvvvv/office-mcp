import { Client } from "@modelcontextprotocol/client";
import { StdioClientTransport } from "@modelcontextprotocol/client/stdio";
import { mkdir, mkdtemp, rm, stat } from "node:fs/promises";
import path from "node:path";

const client = new Client({ name: "office-mcp-smoke-test", version: "1.0.0" });
const transport = new StdioClientTransport({
  command: process.execPath,
  args: ["dist/index.js"],
  cwd: process.cwd()
});
const smokeRoot = path.join(process.cwd(), "work");
await mkdir(smokeRoot, { recursive: true });
const smokeDirectory = await mkdtemp(path.join(smokeRoot, "office-mcp-smoke-"));

try {
  await client.connect(transport);
  const tools = await client.listTools();
  console.log(`Discovered ${tools.tools.length} tools:`);
  for (const tool of tools.tools) console.log(`- ${tool.name}`);

  const result = await client.callTool({ name: "office_capabilities", arguments: {} });
  if (result.isError) throw new Error(`office_capabilities failed: ${JSON.stringify(result.content)}`);
  console.log("\noffice_capabilities succeeded.");

  const deckPath = path.join(smokeDirectory, "designed.pptx");
  const designed = await client.callTool({
    name: "powerpoint_create_designed_presentation",
    arguments: {
      path: deckPath,
      style: "editorial",
      slides: [{ layout: "cover", title: "Smoke test", subtitle: "A readable test deck" }]
    }
  });
  if (designed.isError) throw new Error(`powerpoint_create_designed_presentation failed: ${JSON.stringify(designed.content)}`);
  if ((await stat(deckPath)).size < 1000) throw new Error("Designed presentation output is unexpectedly small.");
  console.log("powerpoint_create_designed_presentation succeeded.");

  const audit = await client.callTool({ name: "powerpoint_audit_presentation", arguments: { path: deckPath } });
  if (audit.isError) throw new Error(`powerpoint_audit_presentation failed: ${JSON.stringify(audit.content)}`);
  console.log("powerpoint_audit_presentation succeeded.");

  const workbookPath = path.join(smokeDirectory, "data.xlsx");
  const workbook = await client.callTool({
    name: "excel_create_workbook",
    arguments: { path: workbookPath, sheets: [{ name: "Data", data: [["Month", "Revenue"], ["Jan", 10], ["Feb", 20]] }] }
  });
  if (workbook.isError) throw new Error(`excel_create_workbook failed: ${JSON.stringify(workbook.content)}`);
  const analysis = await client.callTool({ name: "excel_analyze_dataset", arguments: { path: workbookPath, sheet: "Data" } });
  if (analysis.isError) throw new Error(`excel_analyze_dataset failed: ${JSON.stringify(analysis.content)}`);
  console.log("excel_analyze_dataset succeeded.");
  const query = await client.callTool({
    name: "excel_query_dataset",
    arguments: { path: workbookPath, sheet: "Data", metrics: [{ column: "Revenue", aggregation: "sum" }] }
  });
  if (query.isError) throw new Error(`excel_query_dataset failed: ${JSON.stringify(query.content)}`);
  console.log("excel_query_dataset succeeded.");

  const documentPath = path.join(smokeDirectory, "report.docx");
  const document = await client.callTool({
    name: "word_create_document",
    arguments: {
      path: documentPath,
      sources: [{ id: "source-1", author: "A. Researcher", title: "Example Report", year: "2025" }],
      sections: [{ pageSize: "a4", pageNumbers: true, blocks: [
        { type: "toc" },
        { type: "heading", text: "Findings", level: 1 },
        { type: "cited_paragraph", text: "Evidence was reviewed", sourceIds: ["source-1"] },
        { type: "bibliography" }
      ] }]
    }
  });
  if (document.isError) throw new Error(`word_create_document failed: ${JSON.stringify(document.content)}`);
  console.log("word_create_document with sections and citations succeeded.");

  const qualityCases = [
    {
      path: workbookPath,
      contract: {
        objective: "Create a workbook containing the source revenue figures.",
        criteria: [{ id: "source", description: "The source numbers remain correct." }],
        excel: { requiredSheets: ["Data"], expectedCells: [{ sheet: "Data", cell: "B3", value: 20 }] }
      }
    },
    {
      path: documentPath,
      contract: {
        objective: "Create a report with the requested findings and bibliography.",
        criteria: [{ id: "structure", description: "The report has a findings section." }],
        word: { requiredHeadings: ["Findings"] }
      }
    },
    {
      path: deckPath,
      contract: {
        objective: "Create a readable single-slide presentation for review.",
        criteria: [{ id: "slide", description: "The cover is readable at presentation size." }],
        powerpoint: { minSlides: 1, maxSlides: 1 }
      }
    }
  ];
  for (const item of qualityCases) {
    const quality = await client.callTool({ name: "office_quality_check", arguments: item });
    if (quality.isError) throw new Error(`office_quality_check failed: ${JSON.stringify(quality.content)}`);
    const report = quality.structuredContent as { machineStatus: string; issues: unknown[] };
    if (report.machineStatus !== "needs_review") throw new Error(`Quality check blocked: ${JSON.stringify(report.issues)}`);
  }
  console.log("office_quality_check passed for Excel, Word, and PowerPoint.");

  const deckContract = qualityCases[2].contract;
  const deckQuality = await client.callTool({
    name: "office_quality_check", arguments: { path: deckPath, contract: deckContract }
  });
  const report = deckQuality.structuredContent as { fileSha256: string };
  const native = process.platform === "win32"
    ? await client.callTool({ name: "office_native_status", arguments: {} })
    : null;
  if (native?.isError) throw new Error(`office_native_status failed: ${JSON.stringify(native.content)}`);
  const status = native?.structuredContent as { applications?: Array<{ application: string; installed: boolean }> } | undefined;
  if (status?.applications?.some((application) => application.application === "powerpoint" && application.installed)) {
    const prepared = await client.callTool({
      name: "office_prepare_review",
      arguments: { path: deckPath, contract: deckContract, outputDirectory: path.join(smokeDirectory, "review"), maxInlineImages: 0 }
    });
    if (prepared.isError) throw new Error(`office_prepare_review failed: ${JSON.stringify(prepared.content)}`);
    const evidence = prepared.structuredContent as { manifestPath: string };
    const slideImage = await client.callTool({
      name: "office_get_review_image", arguments: { manifestPath: evidence.manifestPath, unit: "slide:1" }
    });
    if (slideImage.isError || !slideImage.content.some((item) => item.type === "image")) {
      throw new Error("office_get_review_image did not return the rendered slide.");
    }
    const finished = await client.callTool({
      name: "office_finalize",
      arguments: {
        path: deckPath, outputPath: path.join(smokeDirectory, "designed-final.pptx"), contract: deckContract,
        review: {
          fileSha256: report.fileSha256, evidenceManifestPath: evidence.manifestPath,
          units: [{ unit: "slide:1", verdict: "pass", note: "Synthetic protocol test review, not a design assessment." }],
          criteria: [{ id: "slide", verdict: "pass", note: "Synthetic protocol test criterion, not a design assessment." }]
        }
      }
    });
    if (finished.isError) throw new Error(`office_finalize failed: ${JSON.stringify(finished.content)}`);
    console.log("office_finalize accepted rendered evidence and a synthetic protocol review.");
  } else {
    console.log("Native PowerPoint unavailable; skipped render/finalize protocol smoke.");
  }
} finally {
  await client.close();
  await rm(smokeDirectory, { recursive: true, force: true });
}
