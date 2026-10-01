import assert from "node:assert/strict";
import { mkdtemp, readFile, stat } from "node:fs/promises";
import path from "node:path";
import { Client } from "@modelcontextprotocol/client";
import { StdioClientTransport } from "@modelcontextprotocol/client/stdio";
import JSZip from "jszip";

const packageRoot = process.env.OFFICE_PACKAGE_SMOKE_ROOT ?? path.join(process.cwd(), "work", "package-smoke");
const serverPath = path.join(packageRoot, "node_modules", "office-mcp", "dist", "index.js");
const installedRoot = path.join(packageRoot, "node_modules", "office-mcp");
const expectedVersion = (JSON.parse(await readFile(path.join(process.cwd(), "package.json"), "utf8")) as { version: string }).version;
const testDirectory = await mkdtemp(path.join(packageRoot, "test-files-"));
const client = new Client({ name: "packaged-office-mcp-test", version: "1.0.0" });
const transport = new StdioClientTransport({ command: process.execPath, args: [serverPath], cwd: packageRoot });

async function call(name: string, args: Record<string, unknown>) {
  const result = await client.callTool({ name, arguments: args });
  assert.notEqual(result.isError, true, `${name}: ${JSON.stringify(result.content)}`);
  return result;
}

try {
  const uuid = JSON.parse(await readFile(path.join(installedRoot, "node_modules", "uuid", "package.json"), "utf8")) as { version: string };
  const imageSize = JSON.parse(await readFile(path.join(installedRoot, "node_modules", "image-size", "package.json"), "utf8")) as { version: string };
  assert.ok(Number(uuid.version.split(".")[0]) >= 11, "The packaged Excel dependency must include the patched UUID release.");
  assert.ok(Number(imageSize.version.split(".")[0]) > 2 || (Number(imageSize.version.split(".")[0]) === 2 && Number(imageSize.version.split(".")[2]) >= 3), "The packaged PowerPoint dependency must include a patched image-size release.");
  await client.connect(transport);
  const listed = await client.listTools();
  assert.equal(listed.tools.length, 34);
  console.log(`Packaged server exposed ${listed.tools.length} tools.`);

  const capabilities = await call("office_capabilities", {});
  assert.equal((capabilities.structuredContent as { version: string }).version, expectedVersion);
  let nativePowerPoint = false;
  if (process.platform === "win32") {
    const native = await call("office_native_status", {});
    const status = native.structuredContent as { applications?: Array<{ application: string; installed: boolean }> };
    nativePowerPoint = Boolean(status.applications?.some((application) => application.application === "powerpoint" && application.installed));
  }
  console.log("Capabilities and packaged server loaded.");

  const workbookPath = path.join(testDirectory, "data.xlsx");
  await call("excel_create_workbook", {
    path: workbookPath,
    sheets: [{ name: "Data", data: [["Region", "Revenue"], ["North", 10], ["South", 20], ["North", 15]] }]
  });
  await call("excel_query_dataset", {
    path: workbookPath, sheet: "Data", groupBy: ["Region"],
    metrics: [{ column: "Revenue", aggregation: "sum", as: "total" }],
    sort: [{ column: "total", direction: "desc" }]
  });
  await call("excel_inspect_workbook", { path: workbookPath });
  await call("office_quality_check", {
    path: workbookPath,
    contract: {
      objective: "Preserve the packaged workbook data and worksheet structure.",
      criteria: [{ id: "data", description: "The workbook contains the expected source figures." }],
      excel: { requiredSheets: ["Data"], expectedCells: [{ sheet: "Data", cell: "B2", value: 10 }] }
    }
  });
  assert.ok((await stat(workbookPath)).size > 1000);
  const duplicate = await client.callTool({
    name: "excel_create_workbook",
    arguments: { path: workbookPath, sheets: [{ name: "Data", data: [["Value"], [1]] }] }
  });
  assert.equal(duplicate.isError, true, "Existing workbooks must not be overwritten by default.");
  const outside = await client.callTool({
    name: "excel_create_workbook",
    arguments: { path: path.join(process.cwd(), "work", "outside.xlsx"), sheets: [{ name: "Data", data: [["Value"], [1]] }] }
  });
  assert.equal(outside.isError, true, "The installed server must reject paths outside its allowed root.");
  console.log("Packaged Excel create, query, and inspect passed.");

  const documentPath = path.join(testDirectory, "report.docx");
  await call("word_create_document", {
    path: documentPath,
    blocks: [{ type: "heading", text: "Packaged report", level: 1 }, { type: "paragraph", text: "Hello from the package." }]
  });
  await call("word_inspect_document", { path: documentPath });
  const documentZip = await JSZip.loadAsync(await readFile(documentPath));
  assert.ok((await documentZip.file("word/document.xml")!.async("string")).includes("Packaged report"));
  console.log("Packaged Word create and inspect passed.");

  const presentationPath = path.join(testDirectory, "slides.pptx");
  await call("powerpoint_create_designed_presentation", {
    path: presentationPath, style: "editorial",
    slides: [{ layout: "cover", title: "Packaged presentation", subtitle: "Release test" }]
  });
  await call("powerpoint_inspect_presentation", { path: presentationPath });
  await call("powerpoint_audit_presentation", { path: presentationPath });
  if (nativePowerPoint) {
    const renderDirectory = path.join(testDirectory, "rendered");
    await call("powerpoint_render_presentation", {
      path: presentationPath, outputDirectory: renderDirectory, width: 1280, height: 720
    });
    assert.ok((await stat(path.join(renderDirectory, "Slide1.png"))).size > 1000);
  }
  const presentationZip = await JSZip.loadAsync(await readFile(presentationPath));
  assert.ok((await presentationZip.file("ppt/slides/slide1.xml")!.async("string")).includes("Packaged presentation"));
  console.log(`Packaged PowerPoint create, inspect, and audit passed${nativePowerPoint ? ", including native render" : ""}.`);
  console.log(`Test files: ${testDirectory}`);
} finally {
  await client.close();
}
