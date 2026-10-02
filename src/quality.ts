import path from "node:path";
import { createHash, randomUUID } from "node:crypto";
import { readFile, copyFile, mkdir, unlink, writeFile } from "node:fs/promises";
import { constants } from "node:fs";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import { OfficeMcpError } from "./errors.js";
import { prepareOutputPath, resolveAllowedPath, resolveReadablePath } from "./paths.js";
import { inspectDocument } from "./word.js";
import { inspectPresentation } from "./powerpoint.js";
import { auditPresentation } from "./powerpoint-audit.js";
import { renderPowerPointNative } from "./native.js";
import { exportOfficePdf } from "./office.js";
import { compareOfficeFiles } from "./compare.js";

type Scalar = string | number | boolean | null;
type Kind = "excel" | "word" | "powerpoint";
type Severity = "error" | "warning";

export interface QualityContract {
  objective: string;
  criteria: Array<{ id: string; description: string }>;
  requiredText?: string[];
  forbiddenText?: string[];
  baseline?: {
    path: string;
    maxRemovedSheets?: number;
    maxRemovedSlides?: number;
    maxRemovedParagraphs?: number;
    maxChangedCells?: number;
  };
  excel?: {
    requiredSheets?: string[];
    requiredHeaders?: Array<{ sheet: string; row?: number; headers: string[] }>;
    expectedCells?: Array<{ sheet: string; cell: string; value: Scalar; tolerance?: number }>;
    maxFormulaErrors?: number;
    maxUncachedFormulas?: number;
  };
  word?: {
    requiredHeadings?: string[];
    minTables?: number;
    requireImageAltText?: boolean;
  };
  powerpoint?: {
    minSlides?: number;
    maxSlides?: number;
    minimumFontPoints?: number;
    requireSpeakerNotes?: boolean;
  };
}

export interface QualityIssue {
  code: string;
  severity: Severity;
  location: string;
  message: string;
}

export interface QualityReview {
  fileSha256: string;
  evidenceManifestPath?: string;
  units: Array<{ unit: string; verdict: "pass" | "fail"; note: string }>;
  criteria: Array<{ id: string; verdict: "pass" | "fail"; note: string }>;
  acceptedWarnings?: Array<{ code: string; reason: string }>;
}

interface ReviewManifest {
  version: 1;
  kind: Kind;
  sourcePath: string;
  sourceSha256: string;
  contractSha256: string;
  baselineSha256?: string;
  reviewUnits: string[];
  artifacts: Array<{ unit: string; path: string; sha256: string; mediaType: "image/png" | "application/pdf" }>;
}

const extensions: Record<string, Kind> = { ".xlsx": "excel", ".docx": "word", ".pptx": "powerpoint" };
const requiredParts: Record<Kind, string> = {
  excel: "xl/workbook.xml",
  word: "word/document.xml",
  powerpoint: "ppt/presentation.xml"
};
const relationshipParser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });

function issue(code: string, severity: Severity, location: string, message: string): QualityIssue {
  return { code, severity, location, message };
}

function normalize(value: string): string {
  return value.replace(/\s+/g, " ").trim().toLocaleLowerCase();
}

function contractHash(contract: QualityContract): string {
  const sorted = (value: unknown): unknown => Array.isArray(value)
    ? value.map(sorted)
    : value && typeof value === "object"
      ? Object.fromEntries(Object.entries(value).sort(([a], [b]) => a.localeCompare(b)).map(([key, item]) => [key, sorted(item)]))
      : value;
  return createHash("sha256").update(JSON.stringify(sorted(contract))).digest("hex");
}

function list<T>(value: T | T[] | undefined): T[] {
  return value === undefined ? [] : Array.isArray(value) ? value : [value];
}

async function checkRelationships(zip: JSZip): Promise<QualityIssue[]> {
  const issues: QualityIssue[] = [];
  for (const relationshipPath of Object.keys(zip.files).filter((name) => name.endsWith(".rels"))) {
    const xml = await zip.file(relationshipPath)!.async("string");
    let parsed: { Relationships?: { Relationship?: unknown } };
    try {
      parsed = relationshipParser.parse(xml) as typeof parsed;
    } catch {
      issues.push(issue("INVALID_RELATIONSHIPS", "error", relationshipPath, "Relationship XML cannot be parsed."));
      continue;
    }
    const relationships = list(parsed.Relationships?.Relationship as Record<string, unknown> | Record<string, unknown>[] | undefined);
    const sourceDirectory = relationshipPath === "_rels/.rels"
      ? ""
      : path.posix.dirname(path.posix.dirname(relationshipPath));
    for (const relationship of relationships) {
      if (relationship["@_TargetMode"] === "External") continue;
      const rawTarget = String(relationship["@_Target"] ?? "");
      if (!rawTarget) {
        issues.push(issue("EMPTY_RELATIONSHIP", "error", relationshipPath, "An internal relationship has no target."));
        continue;
      }
      const target = rawTarget.startsWith("/")
        ? rawTarget.slice(1)
        : path.posix.normalize(path.posix.join(sourceDirectory, rawTarget.replaceAll("\\", "/")));
      if (!zip.file(target)) {
        issues.push(issue("BROKEN_RELATIONSHIP", "error", relationshipPath, `Missing package target: ${target}`));
      }
    }
  }
  return issues;
}

function checkText(content: string, contract: QualityContract, issues: QualityIssue[]): void {
  const normalized = normalize(content);
  for (const text of contract.requiredText ?? []) {
    if (!normalized.includes(normalize(text))) {
      issues.push(issue("REQUIRED_TEXT_MISSING", "error", "document", `Required text is absent: ${text}`));
    }
  }
  for (const text of contract.forbiddenText ?? []) {
    if (normalized.includes(normalize(text))) {
      issues.push(issue("FORBIDDEN_TEXT_PRESENT", "error", "document", `Forbidden text is present: ${text}`));
    }
  }
}

async function checkExcel(buffer: Buffer, contract: QualityContract, issues: QualityIssue[]) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer as never);
  const checks = contract.excel ?? {};
  const reviewUnits = workbook.worksheets.map((sheet) => `sheet:${sheet.name}`);
  const allText: string[] = [];
  let formulaErrors = 0;
  let uncachedFormulas = 0;
  let cellsScanned = 0;

  for (const sheet of workbook.worksheets) {
    sheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        cellsScanned += 1;
        if (cellsScanned > 1_000_000) throw new OfficeMcpError("Quality check exceeds one million cells", "QUALITY_LIMIT_EXCEEDED");
        allText.push(cell.text);
        const value = cell.value;
        if (value && typeof value === "object" && ("formula" in value || "sharedFormula" in value)) {
          const result = "result" in value ? value.result : undefined;
          if (result === undefined || result === null) uncachedFormulas += 1;
          if (result && typeof result === "object" && "error" in result) formulaErrors += 1;
        } else if (value && typeof value === "object" && "error" in value) {
          formulaErrors += 1;
        }
      });
    });
  }

  for (const name of checks.requiredSheets ?? []) {
    if (!workbook.getWorksheet(name)) issues.push(issue("SHEET_MISSING", "error", `sheet:${name}`, `Required worksheet is absent: ${name}`));
  }
  for (const requirement of checks.requiredHeaders ?? []) {
    const sheet = workbook.getWorksheet(requirement.sheet);
    if (!sheet) {
      issues.push(issue("SHEET_MISSING", "error", `sheet:${requirement.sheet}`, "Cannot check headers because the worksheet is absent."));
      continue;
    }
    const rowValues = sheet.getRow(requirement.row ?? 1).values;
    const headers = new Set((Array.isArray(rowValues) ? rowValues.slice(1) : []).map((value) => normalize(String(value ?? ""))));
    for (const header of requirement.headers) {
      if (!headers.has(normalize(header))) issues.push(issue("HEADER_MISSING", "error", `sheet:${requirement.sheet}`, `Required header is absent: ${header}`));
    }
  }
  for (const expected of checks.expectedCells ?? []) {
    const sheet = workbook.getWorksheet(expected.sheet);
    if (!sheet) {
      issues.push(issue("SHEET_MISSING", "error", `sheet:${expected.sheet}`, "Cannot check a cell because the worksheet is absent."));
      continue;
    }
    const cell = sheet.getCell(expected.cell);
    const raw = cell.value;
    const actual = raw && typeof raw === "object" && "result" in raw ? raw.result : raw;
    const matches = typeof actual === "number" && typeof expected.value === "number"
      ? Math.abs(actual - expected.value) <= (expected.tolerance ?? 0)
      : actual === expected.value;
    if (!matches) issues.push(issue("CELL_MISMATCH", "error", `${expected.sheet}!${expected.cell}`, `Expected ${JSON.stringify(expected.value)}, found ${JSON.stringify(actual)}.`));
  }
  if (formulaErrors > (checks.maxFormulaErrors ?? 0)) {
    issues.push(issue("FORMULA_ERRORS", "error", "workbook", `${formulaErrors} formula or error cells exceed the allowed count.`));
  }
  if (uncachedFormulas > (checks.maxUncachedFormulas ?? 0)) {
    issues.push(issue("UNCACHED_FORMULAS", "error", "workbook", `${uncachedFormulas} formulas have no cached result; recalculate in Excel before finalization.`));
  }
  if (workbook.worksheets.length === 0) issues.push(issue("EMPTY_WORKBOOK", "error", "workbook", "Workbook has no worksheets."));
  checkText(allText.join(" "), contract, issues);
  return { reviewUnits, evidence: { sheets: workbook.worksheets.length, cellsScanned, formulaErrors, uncachedFormulas } };
}

async function checkWord(filePath: string, zip: JSZip, contract: QualityContract, issues: QualityIssue[]) {
  const document = await inspectDocument(filePath, 100_000);
  const checks = contract.word ?? {};
  const documentXml = await zip.file("word/document.xml")!.async("string");
  const content = document.paragraphs.map((paragraph) => paragraph.text).join(" ");
  checkText(content, contract, issues);
  for (const heading of checks.requiredHeadings ?? []) {
    if (!document.paragraphs.some((paragraph) => /^Heading[1-6]$/i.test(paragraph.style ?? "") && normalize(paragraph.text) === normalize(heading))) {
      issues.push(issue("HEADING_MISSING", "error", "document", `Required heading is absent: ${heading}`));
    }
  }
  if (document.tableCount < (checks.minTables ?? 0)) {
    issues.push(issue("TABLE_COUNT", "error", "document", `Expected at least ${checks.minTables} tables, found ${document.tableCount}.`));
  }
  if (checks.requireImageAltText) {
    for (const [index, match] of [...documentXml.matchAll(/<wp:docPr\b[^>]*>/g)].entries()) {
      const description = /\bdescr="([^"]*)"/.exec(match[0])?.[1] ?? /\btitle="([^"]*)"/.exec(match[0])?.[1] ?? "";
      if (!description.trim()) issues.push(issue("IMAGE_ALT_TEXT_MISSING", "error", `image:${index + 1}`, "An embedded image lacks alternative text."));
    }
  }
  let lastHeading = 0;
  for (const paragraph of document.paragraphs) {
    const level = Number(/^Heading([1-6])$/i.exec(paragraph.style ?? "")?.[1]);
    if (!level) continue;
    if (lastHeading && level > lastHeading + 1) {
      issues.push(issue("HEADING_LEVEL_JUMP", "warning", `paragraph:${paragraph.index + 1}`, `Heading level jumps from ${lastHeading} to ${level}.`));
    }
    lastHeading = level;
  }
  return { reviewUnits: ["document"], evidence: { paragraphs: document.paragraphCount, tables: document.tableCount } };
}

async function checkPowerPoint(filePath: string, contract: QualityContract, issues: QualityIssue[]) {
  const presentation = await inspectPresentation(filePath, 10_000);
  const checks = contract.powerpoint ?? {};
  const audit = await auditPresentation(filePath, checks.minimumFontPoints ?? 17);
  if (presentation.slideCount < (checks.minSlides ?? 1)) {
    issues.push(issue("TOO_FEW_SLIDES", "error", "presentation", `Expected at least ${checks.minSlides} slides.`));
  }
  if (checks.maxSlides !== undefined && presentation.slideCount > checks.maxSlides) {
    issues.push(issue("TOO_MANY_SLIDES", "error", "presentation", `Expected at most ${checks.maxSlides} slides.`));
  }
  for (const slide of presentation.slides) {
    if (checks.requireSpeakerNotes && (!slide.speakerNotes || slide.speakerNotes.trim() === String(slide.number))) {
      issues.push(issue("SPEAKER_NOTES_MISSING", "error", `slide:${slide.number}`, "Required speaker notes are absent."));
    }
  }
  for (const finding of audit.issues) {
    const reviewDependent = new Set(["UNMEASURED_GROUP", "TEXT_OVERLAP", "OFF_SLIDE", "LOW_CONTRAST"]);
    const severity: Severity = reviewDependent.has(finding.code) ? "warning" : "error";
    issues.push(issue(finding.code, severity, `slide:${finding.slide}`, finding.message));
  }
  checkText(presentation.slides.flatMap((slide) => slide.text).join(" "), contract, issues);
  return { reviewUnits: presentation.slides.map((slide) => `slide:${slide.number}`), evidence: { slides: presentation.slideCount, auditIssues: audit.issues.length } };
}

export async function checkOfficeQuality(filePath: string, contract: QualityContract) {
  const resolvedPath = await resolveReadablePath(filePath, Object.keys(extensions));
  const kind = extensions[path.extname(resolvedPath).toLowerCase()];
  if (!kind) throw new OfficeMcpError("Unsupported Office format", "UNSUPPORTED_EXTENSION");
  const contractSha256 = contractHash(contract);
  const buffer = await readFile(resolvedPath);
  if (buffer.length > 250_000_000) throw new OfficeMcpError("Quality check supports files up to 250 MB", "QUALITY_LIMIT_EXCEEDED");
  const fileSha256 = createHash("sha256").update(buffer).digest("hex");
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(buffer);
  } catch {
    throw new OfficeMcpError("File is not a valid Office ZIP package", "INVALID_OFFICE_FILE");
  }
  const issues: QualityIssue[] = [];
  if (!zip.file("[Content_Types].xml")) issues.push(issue("CONTENT_TYPES_MISSING", "error", "package", "Office content types are absent."));
  if (!zip.file(requiredParts[kind])) issues.push(issue("MAIN_PART_MISSING", "error", "package", `Missing ${requiredParts[kind]}.`));
  issues.push(...await checkRelationships(zip));
  if (issues.some((finding) => finding.severity === "error")) {
    return { kind, path: resolvedPath, fileSha256, contractSha256, objective: contract.objective, machineStatus: "blocked" as const, issues, reviewUnits: [], criteria: contract.criteria, evidence: {}, baselineSha256: undefined as string | undefined };
  }
  const result = kind === "excel"
    ? await checkExcel(buffer, contract, issues)
    : kind === "word"
      ? await checkWord(resolvedPath, zip, contract, issues)
      : await checkPowerPoint(resolvedPath, contract, issues);
  let baselineSha256: string | undefined;
  let baselineEvidence: unknown;
  if (contract.baseline) {
    const compared = await compareOfficeFiles(contract.baseline.path, resolvedPath, 20);
    if (compared.beforePath === compared.afterPath) {
      throw new OfficeMcpError("The baseline must be a separate source file.", "INVALID_BASELINE");
    }
    baselineSha256 = compared.beforeSha256;
    baselineEvidence = compared;
    if (compared.kind === "excel") {
      if (compared.summary.removedSheets.length > (contract.baseline.maxRemovedSheets ?? 0)) {
        issues.push(issue("BASELINE_SHEETS_REMOVED", "error", "workbook", "Sheets were removed relative to the source workbook."));
      }
      if (contract.baseline.maxChangedCells !== undefined && compared.summary.changedCells > contract.baseline.maxChangedCells) {
        issues.push(issue("BASELINE_TOO_MANY_CHANGED_CELLS", "error", "workbook", "More cells changed than the contract allows."));
      }
    } else if (compared.kind === "word" && compared.summary.removedParagraphs > (contract.baseline.maxRemovedParagraphs ?? 0)) {
      issues.push(issue("BASELINE_PARAGRAPHS_REMOVED", "error", "document", "Paragraphs were removed or replaced relative to the source document."));
    } else if (compared.kind === "powerpoint" && compared.summary.removedSlides > (contract.baseline.maxRemovedSlides ?? 0)) {
      issues.push(issue("BASELINE_SLIDES_REMOVED", "error", "presentation", "Slides were removed relative to the source presentation."));
    }
  }
  return {
    kind, path: resolvedPath, fileSha256, contractSha256, baselineSha256, objective: contract.objective,
    machineStatus: issues.some((finding) => finding.severity === "error") ? "blocked" as const : "needs_review" as const,
    issues, reviewUnits: result.reviewUnits, criteria: contract.criteria, evidence: { ...result.evidence, baseline: baselineEvidence }
  };
}

export async function prepareOfficeReview(filePath: string, contract: QualityContract, outputDirectory: string) {
  const report = await checkOfficeQuality(filePath, contract);
  if (report.machineStatus === "blocked") {
    throw new OfficeMcpError("Repair blocking quality issues before rendering the review copy.", "QUALITY_CHECK_FAILED", { issues: report.issues });
  }
  const root = resolveAllowedPath(outputDirectory);
  await mkdir(root, { recursive: true });
  const reviewDirectory = resolveAllowedPath(path.join(root, `review-${randomUUID()}`));
  await mkdir(reviewDirectory);
  const artifacts: ReviewManifest["artifacts"] = [];
  if (report.kind === "powerpoint") {
    const rendered = await renderPowerPointNative(report.path, reviewDirectory, 1600, 900);
    if (rendered.images.length !== report.reviewUnits.length) {
      throw new OfficeMcpError("Rendered slide count does not match the draft.", "RENDER_COUNT_MISMATCH");
    }
    for (const [index, image] of rendered.images.entries()) {
      artifacts.push({
        unit: report.reviewUnits[index], path: image.path,
        sha256: createHash("sha256").update(Buffer.from(image.data, "base64")).digest("hex"),
        mediaType: "image/png"
      });
    }
  } else {
    const pdfPath = path.join(reviewDirectory, "review.pdf");
    const exported = await exportOfficePdf(report.path, pdfPath, "auto");
    const pdf = await readFile(exported.path);
    if (pdf.subarray(0, 4).toString() !== "%PDF") throw new OfficeMcpError("Office export did not produce a PDF.", "INVALID_REVIEW_PDF");
    const sha256 = createHash("sha256").update(pdf).digest("hex");
    for (const unit of report.reviewUnits) artifacts.push({ unit, path: exported.path, sha256, mediaType: "application/pdf" });
  }
  const manifest: ReviewManifest = {
    version: 1, kind: report.kind, sourcePath: report.path, sourceSha256: report.fileSha256,
    contractSha256: report.contractSha256, baselineSha256: report.baselineSha256,
    reviewUnits: report.reviewUnits, artifacts
  };
  const manifestPath = await prepareOutputPath(path.join(reviewDirectory, "review-manifest.json"), [".json"]);
  await writeFile(manifestPath, JSON.stringify(manifest, null, 2), "utf8");
  return { manifestPath, reviewDirectory, fileSha256: report.fileSha256, kind: report.kind, reviewUnits: report.reviewUnits, artifacts };
}

async function verifyEvidence(report: Awaited<ReturnType<typeof checkOfficeQuality>>, manifestPath: string): Promise<void> {
  const resolved = await resolveReadablePath(manifestPath, [".json"]);
  let manifest: ReviewManifest;
  try {
    manifest = JSON.parse(await readFile(resolved, "utf8")) as ReviewManifest;
  } catch {
    throw new OfficeMcpError("Review evidence manifest is invalid JSON.", "INVALID_REVIEW_EVIDENCE");
  }
  if (manifest.version !== 1 || manifest.kind !== report.kind || manifest.sourcePath !== report.path || manifest.sourceSha256 !== report.fileSha256 || manifest.contractSha256 !== report.contractSha256 || manifest.baselineSha256 !== report.baselineSha256) {
    throw new OfficeMcpError("Review evidence does not match the current draft.", "STALE_REVIEW_EVIDENCE");
  }
  if (JSON.stringify(manifest.reviewUnits) !== JSON.stringify(report.reviewUnits) || manifest.artifacts.length !== report.reviewUnits.length) {
    throw new OfficeMcpError("Review evidence does not cover every unit.", "REVIEW_EVIDENCE_INCOMPLETE");
  }
  for (const [index, item] of manifest.artifacts.entries()) {
    if (item.unit !== report.reviewUnits[index]) throw new OfficeMcpError("Review evidence units are out of order.", "REVIEW_EVIDENCE_INCOMPLETE");
    const extension = item.mediaType === "image/png" ? [".png"] : [".pdf"];
    const artifactPath = await resolveReadablePath(item.path, extension);
    const data = await readFile(artifactPath);
    const signature = item.mediaType === "image/png" ? "89504e470d0a1a0a" : "25504446";
    if (!data.subarray(0, signature.length / 2).equals(Buffer.from(signature, "hex"))) {
      throw new OfficeMcpError("Review artifact has an invalid file signature.", "INVALID_REVIEW_EVIDENCE");
    }
    const sha256 = createHash("sha256").update(data).digest("hex");
    if (sha256 !== item.sha256) throw new OfficeMcpError("Review artifact changed after preparation.", "STALE_REVIEW_EVIDENCE");
  }
}

export async function readReviewImage(manifestPath: string, unit: string) {
  const resolved = await resolveReadablePath(manifestPath, [".json"]);
  let manifest: ReviewManifest;
  try {
    manifest = JSON.parse(await readFile(resolved, "utf8")) as ReviewManifest;
  } catch {
    throw new OfficeMcpError("Review evidence manifest is invalid JSON.", "INVALID_REVIEW_EVIDENCE");
  }
  if (manifest.version !== 1 || manifest.kind !== "powerpoint" || !manifest.reviewUnits.includes(unit)) {
    throw new OfficeMcpError("Requested slide is not in this PowerPoint review.", "REVIEW_UNIT_NOT_FOUND");
  }
  const source = await resolveReadablePath(manifest.sourcePath, [".pptx"]);
  const sourceSha256 = createHash("sha256").update(await readFile(source)).digest("hex");
  if (sourceSha256 !== manifest.sourceSha256) {
    throw new OfficeMcpError("The draft changed after the slides were rendered.", "STALE_REVIEW_EVIDENCE");
  }
  const artifact = manifest.artifacts.find((item) => item.unit === unit);
  if (!artifact || artifact.mediaType !== "image/png") {
    throw new OfficeMcpError("The rendered slide image is unavailable.", "REVIEW_UNIT_NOT_FOUND");
  }
  const imagePath = await resolveReadablePath(artifact.path, [".png"]);
  const data = await readFile(imagePath);
  const signature = Buffer.from("89504e470d0a1a0a", "hex");
  if (!data.subarray(0, signature.length).equals(signature) || createHash("sha256").update(data).digest("hex") !== artifact.sha256) {
    throw new OfficeMcpError("The rendered slide changed after preparation.", "STALE_REVIEW_EVIDENCE");
  }
  return { unit, imagePath, data };
}

function verifyReview(
  report: Awaited<ReturnType<typeof checkOfficeQuality>>,
  review: QualityReview
): void {
  if (review.fileSha256 !== report.fileSha256) {
    throw new OfficeMcpError("The draft changed after review; run quality checks and review it again.", "STALE_QUALITY_REVIEW");
  }
  const checkedUnits = new Map(review.units.map((entry) => [entry.unit, entry]));
  if (checkedUnits.size !== review.units.length || checkedUnits.size !== report.reviewUnits.length) {
    throw new OfficeMcpError("Review must cover each slide, sheet, or document exactly once.", "REVIEW_INCOMPLETE");
  }
  for (const unit of report.reviewUnits) {
    const entry = checkedUnits.get(unit);
    if (!entry || entry.verdict !== "pass" || entry.note.trim().length < 12) {
      throw new OfficeMcpError(`Review is missing or failed for ${unit}.`, "REVIEW_INCOMPLETE");
    }
  }
  const checkedCriteria = new Map(review.criteria.map((entry) => [entry.id, entry]));
  if (checkedCriteria.size !== review.criteria.length || checkedCriteria.size !== report.criteria.length) {
    throw new OfficeMcpError("Review must address every prompt criterion exactly once.", "REVIEW_INCOMPLETE");
  }
  for (const criterion of report.criteria) {
    const entry = checkedCriteria.get(criterion.id);
    if (!entry || entry.verdict !== "pass" || entry.note.trim().length < 12) {
      throw new OfficeMcpError(`Prompt criterion is missing or failed: ${criterion.id}.`, "REVIEW_INCOMPLETE");
    }
  }
  const accepted = new Map((review.acceptedWarnings ?? []).map((item) => [item.code, item.reason]));
  for (const warning of report.issues.filter((item) => item.severity === "warning")) {
    if ((accepted.get(warning.code) ?? "").trim().length < 12) {
      throw new OfficeMcpError(`Warning requires a reasoned acceptance: ${warning.code}.`, "WARNING_NOT_ACCEPTED");
    }
  }
}

export async function finalizeOfficeFile(filePath: string, outputPath: string, contract: QualityContract, review: QualityReview, enforceEvidence = false) {
  const report = await checkOfficeQuality(filePath, contract);
  if (report.machineStatus === "blocked") {
    throw new OfficeMcpError("Machine quality checks failed; repair the draft before finalization.", "QUALITY_CHECK_FAILED", {
      issues: report.issues
    });
  }
  verifyReview(report, review);
  if (enforceEvidence) {
    if (!review.evidenceManifestPath) throw new OfficeMcpError("A rendered review manifest is required before publication.", "REVIEW_EVIDENCE_REQUIRED");
    await verifyEvidence(report, review.evidenceManifestPath);
  }
  const extension = path.extname(report.path).toLowerCase();
  const output = await prepareOutputPath(outputPath, [extension], false);
  if (output === report.path) throw new OfficeMcpError("Finalize to a separate file, leaving the draft unchanged.", "SEPARATE_OUTPUT_REQUIRED");
  try {
    await copyFile(report.path, output, constants.COPYFILE_EXCL);
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "EEXIST") throw new OfficeMcpError("Final output already exists.", "FILE_EXISTS");
    throw error;
  }
  const finalHash = createHash("sha256").update(await readFile(output)).digest("hex");
  const currentDraftHash = createHash("sha256").update(await readFile(report.path)).digest("hex");
  if (finalHash !== report.fileSha256 || currentDraftHash !== report.fileSha256) {
    await unlink(output);
    throw new OfficeMcpError("The draft changed during finalization; the attempted final copy was removed.", "FINAL_COPY_MISMATCH");
  }
  if (contract.baseline) {
    const baselinePath = await resolveReadablePath(contract.baseline.path, [extension]);
    const currentBaselineHash = createHash("sha256").update(await readFile(baselinePath)).digest("hex");
    if (currentBaselineHash !== report.baselineSha256) {
      await unlink(output);
      throw new OfficeMcpError("The baseline changed during finalization; the attempted final copy was removed.", "STALE_BASELINE");
    }
  }
  return { path: output, fileSha256: finalHash, kind: report.kind, machineStatus: "pass", reviewStatus: "attested", reviewedUnits: report.reviewUnits, criteria: report.criteria };
}
