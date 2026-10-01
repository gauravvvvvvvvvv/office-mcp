import path from "node:path";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import ExcelJS from "exceljs";
import { OfficeMcpError } from "./errors.js";
import { resolveReadablePath } from "./paths.js";
import { inspectDocument } from "./word.js";
import { inspectPresentation } from "./powerpoint.js";

const extensions = [".xlsx", ".docx", ".pptx"];

async function hashFile(filePath: string): Promise<string> {
  return createHash("sha256").update(await readFile(filePath)).digest("hex");
}

function counts(values: string[]): Map<string, number> {
  const result = new Map<string, number>();
  for (const value of values) result.set(value, (result.get(value) ?? 0) + 1);
  return result;
}

function multisetDifference(before: string[], after: string[]) {
  const beforeCounts = counts(before);
  const afterCounts = counts(after);
  const removed: string[] = [];
  const added: string[] = [];
  for (const [value, count] of beforeCounts) {
    for (let index = 0; index < count - (afterCounts.get(value) ?? 0); index++) removed.push(value);
  }
  for (const [value, count] of afterCounts) {
    for (let index = 0; index < count - (beforeCounts.get(value) ?? 0); index++) added.push(value);
  }
  return { removed, added };
}

async function compareWord(beforePath: string, afterPath: string, maxChanges: number) {
  const before = await inspectDocument(beforePath, 100_000);
  const after = await inspectDocument(afterPath, 100_000);
  if (before.truncated || after.truncated) throw new OfficeMcpError("Document comparison exceeds 100,000 paragraphs", "COMPARE_LIMIT_EXCEEDED");
  const paragraphKey = (paragraph: { style: string | null; text: string }) => `${paragraph.style ?? "Normal"}: ${paragraph.text}`;
  const difference = multisetDifference(before.paragraphs.map(paragraphKey), after.paragraphs.map(paragraphKey));
  return {
    kind: "word" as const,
    summary: {
      beforeParagraphs: before.paragraphCount, afterParagraphs: after.paragraphCount,
      removedParagraphs: difference.removed.length, addedParagraphs: difference.added.length,
      beforeTables: before.tableCount, afterTables: after.tableCount
    },
    changes: { removed: difference.removed.slice(0, maxChanges), added: difference.added.slice(0, maxChanges) },
    truncatedChanges: difference.removed.length > maxChanges || difference.added.length > maxChanges
  };
}

async function comparePowerPoint(beforePath: string, afterPath: string, maxChanges: number) {
  const before = await inspectPresentation(beforePath, 10_000);
  const after = await inspectPresentation(afterPath, 10_000);
  if (before.truncated || after.truncated) throw new OfficeMcpError("Presentation comparison exceeds 10,000 slides", "COMPARE_LIMIT_EXCEEDED");
  const changed: Array<{ slide: number; beforeText: string[]; afterText: string[]; beforePictures: number; afterPictures: number; beforeCharts: number; afterCharts: number }> = [];
  for (let index = 0; index < Math.min(before.slides.length, after.slides.length); index++) {
    const first = before.slides[index];
    const second = after.slides[index];
    if (JSON.stringify(first.text) !== JSON.stringify(second.text) || first.pictureCount !== second.pictureCount || first.chartCount !== second.chartCount) {
      changed.push({
        slide: index + 1, beforeText: first.text, afterText: second.text,
        beforePictures: first.pictureCount, afterPictures: second.pictureCount,
        beforeCharts: first.chartCount, afterCharts: second.chartCount
      });
    }
  }
  return {
    kind: "powerpoint" as const,
    summary: {
      beforeSlides: before.slideCount, afterSlides: after.slideCount,
      removedSlides: Math.max(0, before.slideCount - after.slideCount),
      addedSlides: Math.max(0, after.slideCount - before.slideCount),
      changedSlides: changed.length
    },
    changes: changed.slice(0, maxChanges),
    truncatedChanges: changed.length > maxChanges
  };
}

function sheetCells(sheet: ExcelJS.Worksheet): Map<string, string> {
  const cells = new Map<string, string>();
  sheet.eachRow({ includeEmpty: false }, (row) => row.eachCell({ includeEmpty: false }, (cell) => {
    if (cells.size >= 1_000_000) throw new OfficeMcpError("Workbook comparison exceeds one million cells per sheet", "COMPARE_LIMIT_EXCEEDED");
    cells.set(cell.address, JSON.stringify({ value: cell.value, numberFormat: cell.numFmt }));
  }));
  return cells;
}

async function compareExcel(beforePath: string, afterPath: string, maxChanges: number) {
  const before = new ExcelJS.Workbook();
  const after = new ExcelJS.Workbook();
  await before.xlsx.readFile(beforePath);
  await after.xlsx.readFile(afterPath);
  const beforeNames = before.worksheets.map((sheet) => sheet.name);
  const afterNames = after.worksheets.map((sheet) => sheet.name);
  const removedSheets = beforeNames.filter((name) => !afterNames.includes(name));
  const addedSheets = afterNames.filter((name) => !beforeNames.includes(name));
  const changes: Array<{ sheet: string; cell: string; before: string | null; after: string | null }> = [];
  let changedCells = 0;
  for (const name of beforeNames.filter((candidate) => afterNames.includes(candidate))) {
    const first = sheetCells(before.getWorksheet(name)!);
    const second = sheetCells(after.getWorksheet(name)!);
    for (const address of new Set([...first.keys(), ...second.keys()])) {
      const original = first.get(address) ?? null;
      const updated = second.get(address) ?? null;
      if (original !== updated) {
        changedCells += 1;
        if (changes.length < maxChanges) changes.push({ sheet: name, cell: address, before: original, after: updated });
      }
    }
  }
  return {
    kind: "excel" as const,
    summary: { beforeSheets: beforeNames.length, afterSheets: afterNames.length, removedSheets, addedSheets, changedCells },
    changes,
    truncatedChanges: changedCells > maxChanges
  };
}

export async function compareOfficeFiles(beforePath: string, afterPath: string, maxChanges = 100) {
  const before = await resolveReadablePath(beforePath, extensions);
  const after = await resolveReadablePath(afterPath, extensions);
  const extension = path.extname(before).toLowerCase();
  if (extension !== path.extname(after).toLowerCase()) {
    throw new OfficeMcpError("Office comparison requires matching file formats", "FORMAT_MISMATCH");
  }
  const compared = extension === ".xlsx"
    ? await compareExcel(before, after, maxChanges)
    : extension === ".docx"
      ? await compareWord(before, after, maxChanges)
      : await comparePowerPoint(before, after, maxChanges);
  return {
    beforePath: before, afterPath: after,
    beforeSha256: await hashFile(before), afterSha256: await hashFile(after),
    ...compared
  };
}
