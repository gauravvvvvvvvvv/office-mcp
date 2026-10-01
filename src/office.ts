import path from "node:path";
import { access, copyFile, lstat, mkdtemp, readdir, rm } from "node:fs/promises";
import { constants } from "node:fs";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { inspectWorkbook } from "./excel.js";
import { inspectDocument } from "./word.js";
import { inspectPresentation } from "./powerpoint.js";
import { OfficeMcpError } from "./errors.js";
import { getAllowedRoots, prepareOutputPath, resolveAllowedPath, resolveReadablePath } from "./paths.js";
import { runNativeBatch } from "./native.js";
import { OFFICE_MCP_VERSION } from "./version.js";

const execFileAsync = promisify(execFile);
const OFFICE_EXTENSIONS = [".xlsx", ".docx", ".pptx"];

const converterCandidates = [
  process.env.OFFICE_MCP_SOFFICE,
  process.platform === "win32" ? "C:\\Program Files\\LibreOffice\\program\\soffice.exe" : undefined,
  process.platform === "win32" ? "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe" : undefined,
  "soffice"
].filter((candidate): candidate is string => Boolean(candidate));

export async function findLibreOffice(): Promise<string | null> {
  for (const candidate of converterCandidates) {
    try {
      if (path.isAbsolute(candidate)) await access(candidate, constants.X_OK);
      await execFileAsync(candidate, ["--version"], { timeout: 5000, windowsHide: true });
      return candidate;
    } catch {
      // Try the next known location.
    }
  }
  return null;
}

export async function capabilities() {
  const libreOffice = await findLibreOffice();
  return {
    name: "office-mcp",
    version: OFFICE_MCP_VERSION,
    allowedRoots: getAllowedRoots(),
    applications: {
      excel: {
        extensions: [".xlsx"],
        operations: ["create", "inspect", "read_range", "analyze_dataset", "query_dataset", "write_range", "set_formulas"],
        nativeOperations: ["define_name", "set_conditional_format", "set_validation", "create_chart", "pivot_table", "sort", "autofilter", "recalculate", "export_pdf"]
      },
      word: {
        extensions: [".docx"],
        operations: ["create", "create_sections", "create_toc", "create_citations", "create_images", "create_hyperlinks", "create_footnotes", "inspect", "append_paragraph", "replace_text"],
        nativeOperations: ["apply_style", "insert_section", "create_toc", "update_toc", "add_comment", "add_table", "set_track_changes", "export_pdf"]
      },
      powerpoint: {
        extensions: [".pptx"],
        operations: ["create", "create_designed", "inspect", "audit", "inspect_template", "create_from_template", "replace_text"],
        createElements: ["text", "shape", "image", "table", "chart", "speaker_notes"],
        designedLayouts: ["cover", "imageText", "statement", "comparison", "panorama"],
        designedStyles: ["editorial", "cinematic"],
        nativeAnimation: {
          operations: ["list_shapes", "list_animations", "add_animation", "update_animation", "move_animation", "delete_animation", "clear_animations", "set_transition", "add_media"],
          effects: ["appear", "fade", "fly", "wipe", "zoom", "spin", "growShrink"],
          triggers: ["onClick", "withPrevious", "afterPrevious"]
        }
      }
    },
    pdfExport: {
      available: Boolean(libreOffice),
      converter: libreOffice
    },
    engines: {
      portable: "Direct XLSX, DOCX, and PPTX package operations without requiring Microsoft Office",
      nativeWindows: "Full-fidelity automation through installed Excel, Word, and PowerPoint",
      visualVerification: "PowerPoint slide rendering to MCP image content through the native application"
    },
    orchestration: {
      discovery: "office_search_files",
      activeFiles: "office_list_open_files",
      nativeStatus: "office_native_status",
      powerpointRenderLoop: "powerpoint_render_presentation",
      qualityWorkflow: ["office_quality_check", "office_compare_files", "office_prepare_review", "office_get_review_image", "office_finalize"]
    },
    limitations: [
      "Designed PowerPoint layouts are editorial starting points, not a substitute for relevant imagery, content judgment, or visual review.",
      "The quality gate checks explicit assertions and requires a render-backed review attestation. It cannot independently prove factual accuracy, visual taste, or that a human actually inspected the evidence.",
      "The PowerPoint audit is heuristic. It checks contrast only for explicit solid colors and cannot prove text fit, overall layout quality, or animation playback; render and review every slide.",
      "Word citation markers and bibliography are formatted text, not Microsoft Word citation-manager fields. TOC entries populate when Word updates the field.",
      "Excel dataset analysis is descriptive and bounded to 50,000 rows and 100 columns per call; formula cells need cached results or native recalculation.",
      "PowerPoint object animations and slide transitions require installed desktop PowerPoint on Windows; static PNG renders do not show motion.",
      "File-based native PowerPoint work keeps presentation windows hidden. If PowerPoint is already running, the server refuses background native work unless target.active=true is explicitly requested; portable PPTX creation still works.",
      "Editing an existing Excel file through ExcelJS may not preserve unsupported Excel features such as embedded charts or macros.",
      "Word and PowerPoint text replacement preserves package structure but consolidates replaced multi-run text into the first run.",
      "Cross-platform live control requires the planned Office.js bridge; Windows live control is available through the native batch tools."
    ]
  };
}

export async function inspectOfficeFile(filePath: string, sampleLimit = 20) {
  const resolved = await resolveReadablePath(filePath, OFFICE_EXTENSIONS);
  const extension = path.extname(resolved).toLowerCase();
  if (extension === ".xlsx") return inspectWorkbook(resolved, sampleLimit);
  if (extension === ".docx") return inspectDocument(resolved, sampleLimit);
  return inspectPresentation(resolved, sampleLimit);
}

export async function searchOfficeFiles(
  directory: string,
  recursive = true,
  extensions: string[] = OFFICE_EXTENSIONS,
  maxResults = 500
) {
  const root = await resolveReadablePath(directory);
  const requestedExtensions = extensions.map((extension) =>
    (extension.startsWith(".") ? extension : `.${extension}`).toLowerCase()
  );
  const unsupported = requestedExtensions.filter((extension) => !OFFICE_EXTENSIONS.includes(extension));
  if (unsupported.length) {
    throw new OfficeMcpError(`Unsupported search extensions: ${unsupported.join(", ")}`, "UNSUPPORTED_EXTENSION");
  }

  const files: Array<{ path: string; relativePath: string; extension: string; size: number; modified: Date }> = [];
  const pending = [root];
  while (pending.length && files.length < maxResults) {
    const current = pending.shift()!;
    for (const entry of await readdir(current, { withFileTypes: true })) {
      if (files.length >= maxResults) break;
      const candidate = path.join(current, entry.name);
      if (entry.isSymbolicLink()) continue;
      if (entry.isDirectory()) {
        if (recursive) pending.push(candidate);
        continue;
      }
      const extension = path.extname(entry.name).toLowerCase();
      if (!requestedExtensions.includes(extension)) continue;
      const metadata = await lstat(candidate);
      files.push({
        path: candidate,
        relativePath: path.relative(root, candidate),
        extension,
        size: metadata.size,
        modified: metadata.mtime
      });
    }
  }
  return { directory: root, recursive, files, truncated: pending.length > 0 || files.length >= maxResults };
}

export async function convertToPdf(filePath: string, outputPath?: string, overwrite = false) {
  const input = await resolveReadablePath(filePath, OFFICE_EXTENSIONS);
  const desired = await prepareOutputPath(
    outputPath ?? path.join(path.dirname(input), `${path.basename(input, path.extname(input))}.pdf`),
    [".pdf"],
    overwrite
  );
  const converter = await findLibreOffice();
  if (!converter) {
    throw new OfficeMcpError(
      "PDF export requires LibreOffice. Install it or set OFFICE_MCP_SOFFICE to the soffice executable.",
      "CONVERTER_UNAVAILABLE"
    );
  }

  const temporaryDirectory = await mkdtemp(path.join(path.dirname(desired), ".office-mcp-convert-"));
  try {
    await execFileAsync(
      converter,
      ["--headless", "--convert-to", "pdf", "--outdir", temporaryDirectory, input],
      { timeout: 120000, windowsHide: true, maxBuffer: 1024 * 1024 }
    );
    const generated = resolveAllowedPath(
      path.join(temporaryDirectory, `${path.basename(input, path.extname(input))}.pdf`),
      [".pdf"]
    );
    try {
      await access(generated, constants.R_OK);
    } catch {
      throw new OfficeMcpError("LibreOffice completed without producing the expected PDF", "CONVERSION_FAILED");
    }
    await copyFile(generated, desired, overwrite ? 0 : constants.COPYFILE_EXCL);
    return { input, path: desired, format: "pdf" };
  } finally {
    await rm(temporaryDirectory, { recursive: true, force: true });
  }
}

export async function exportOfficePdf(
  filePath: string,
  outputPath?: string,
  engine: "auto" | "native" | "libreoffice" = "auto",
  overwrite = false
) {
  const input = await resolveReadablePath(filePath, OFFICE_EXTENSIONS);
  const desired = outputPath ?? path.join(path.dirname(input), `${path.basename(input, path.extname(input))}.pdf`);
  const application = { ".xlsx": "excel", ".docx": "word", ".pptx": "powerpoint" }[path.extname(input).toLowerCase()] as "excel" | "word" | "powerpoint";
  if (engine !== "libreoffice" && process.platform === "win32") {
    try {
      await runNativeBatch(application, {
        target: { path: input },
        operations: [{ op: "export_pdf", outputPath: desired }],
        overwrite
      });
      const result = await resolveReadablePath(desired, [".pdf"]);
      return { input, path: result, format: "pdf", engine: "native" };
    } catch (error) {
      if (engine === "native" || !(await findLibreOffice())) throw error;
    }
  } else if (engine === "native") {
    throw new OfficeMcpError("Native PDF export requires Windows desktop Office.", "NATIVE_UNAVAILABLE");
  }
  const result = await convertToPdf(input, desired, overwrite);
  return { ...result, engine: "libreoffice" };
}
