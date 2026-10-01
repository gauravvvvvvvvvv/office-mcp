import ExcelJS from "exceljs";
import path from "node:path";
import { OfficeMcpError } from "./errors.js";
import { prepareOutputPath, resolveInputOutput, resolveReadablePath } from "./paths.js";

const EXCEL_EXTENSIONS = [".xlsx"];

export interface ExcelSheetSpec {
  name: string;
  data?: unknown[][];
  headerRow?: boolean;
  freezeRows?: number;
  columnWidths?: number[];
  autoFilter?: boolean;
}

export interface ExcelWorkbookSpec {
  path: string;
  sheets: ExcelSheetSpec[];
  overwrite?: boolean;
  creator?: string;
}

export interface ExcelWriteSpec {
  path: string;
  sheet: string;
  startCell: string;
  values: unknown[][];
  outputPath?: string;
  overwrite?: boolean;
  numberFormat?: string;
  bold?: boolean;
  fillColor?: string;
}

export interface ExcelFormulaSpec {
  cell: string;
  formula: string;
  result?: string | number | boolean | null;
}

function columnToNumber(column: string): number {
  let result = 0;
  for (const char of column.toUpperCase()) {
    result = result * 26 + char.charCodeAt(0) - 64;
  }
  return result;
}

function parseCellAddress(address: string): { row: number; column: number } {
  const match = /^([A-Za-z]+)([1-9]\d*)$/.exec(address.trim());
  if (!match) {
    throw new OfficeMcpError(`Invalid cell address: ${address}`, "INVALID_CELL_ADDRESS");
  }
  return { column: columnToNumber(match[1]), row: Number(match[2]) };
}

function parseRange(address: string): { startRow: number; startColumn: number; endRow: number; endColumn: number } {
  const [startText, endText = startText] = address.split(":");
  const start = parseCellAddress(startText);
  const end = parseCellAddress(endText);
  if (end.row < start.row || end.column < start.column) {
    throw new OfficeMcpError(`Range is reversed: ${address}`, "INVALID_RANGE");
  }
  return {
    startRow: start.row,
    startColumn: start.column,
    endRow: end.row,
    endColumn: end.column
  };
}

function serializeCell(cell: ExcelJS.Cell, includeStyles: boolean) {
  const value = cell.value;
  const formula = value && typeof value === "object" && "formula" in value
    ? String(value.formula)
    : undefined;
  const result = formula && value && typeof value === "object" && "result" in value
    ? value.result
    : undefined;

  return {
    address: cell.address,
    value: formula ? result ?? null : value ?? null,
    text: cell.text,
    ...(formula ? { formula } : {}),
    ...(includeStyles
      ? {
          style: {
            numberFormat: cell.numFmt || undefined,
            bold: cell.font?.bold || undefined,
            italic: cell.font?.italic || undefined,
            fill: cell.fill,
            alignment: cell.alignment
          }
        }
      : {})
  };
}

async function loadWorkbook(filePath: string): Promise<{ workbook: ExcelJS.Workbook; resolvedPath: string }> {
  const resolvedPath = await resolveReadablePath(filePath, EXCEL_EXTENSIONS);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(resolvedPath);
  return { workbook, resolvedPath };
}

export async function createWorkbook(spec: ExcelWorkbookSpec) {
  const outputPath = await prepareOutputPath(spec.path, EXCEL_EXTENSIONS, spec.overwrite);
  const workbook = new ExcelJS.Workbook();
  workbook.creator = spec.creator ?? "Office MCP";
  workbook.created = new Date();
  workbook.modified = new Date();

  for (const sheetSpec of spec.sheets) {
    const worksheet = workbook.addWorksheet(sheetSpec.name, {
      views: sheetSpec.freezeRows
        ? [{ state: "frozen", ySplit: sheetSpec.freezeRows }]
        : undefined
    });

    for (const row of sheetSpec.data ?? []) {
      worksheet.addRow(row as ExcelJS.CellValue[]);
    }

    sheetSpec.columnWidths?.forEach((width, index) => {
      worksheet.getColumn(index + 1).width = width;
    });

    if (sheetSpec.headerRow && worksheet.rowCount > 0) {
      const header = worksheet.getRow(1);
      header.font = { bold: true, color: { argb: "FFFFFFFF" } };
      header.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF1F4E78" }
      };
      header.alignment = { vertical: "middle" };
    }

    if (sheetSpec.autoFilter && worksheet.rowCount > 0 && worksheet.columnCount > 0) {
      worksheet.autoFilter = {
        from: { row: 1, column: 1 },
        to: { row: worksheet.rowCount, column: worksheet.columnCount }
      };
    }
  }

  await workbook.xlsx.writeFile(outputPath);
  return {
    path: outputPath,
    sheets: workbook.worksheets.map((sheet) => sheet.name)
  };
}

export async function inspectWorkbook(filePath: string, sampleRows = 5) {
  const { workbook, resolvedPath } = await loadWorkbook(filePath);
  const sheets = workbook.worksheets.map((worksheet) => {
    let formulas = 0;
    let nonEmptyCells = 0;

    worksheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        nonEmptyCells += 1;
        if (cell.value && typeof cell.value === "object" && "formula" in cell.value) formulas += 1;
      });
    });

    const sample: unknown[][] = [];
    for (let rowNumber = 1; rowNumber <= Math.min(worksheet.rowCount, sampleRows); rowNumber += 1) {
      const row: unknown[] = [];
      for (let column = 1; column <= worksheet.columnCount; column += 1) {
        row.push(worksheet.getCell(rowNumber, column).text);
      }
      sample.push(row);
    }

    return {
      name: worksheet.name,
      state: worksheet.state,
      rowCount: worksheet.rowCount,
      columnCount: worksheet.columnCount,
      usedRange: worksheet.rowCount && worksheet.columnCount
        ? `A1:${worksheet.getCell(worksheet.rowCount, worksheet.columnCount).address}`
        : null,
      nonEmptyCells,
      formulaCount: formulas,
      mergedRanges: (worksheet.model as { merges?: string[] }).merges ?? [],
      sample
    };
  });

  return {
    kind: "excel",
    path: resolvedPath,
    filename: path.basename(resolvedPath),
    creator: workbook.creator || null,
    modified: workbook.modified ?? null,
    sheetCount: sheets.length,
    sheets
  };
}

export async function readRange(
  filePath: string,
  sheetName: string,
  rangeAddress: string,
  includeStyles = false
) {
  const { workbook, resolvedPath } = await loadWorkbook(filePath);
  const worksheet = workbook.getWorksheet(sheetName);
  if (!worksheet) throw new OfficeMcpError(`Worksheet not found: ${sheetName}`, "SHEET_NOT_FOUND");

  const range = parseRange(rangeAddress);
  const cells = [];
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    const outputRow = [];
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      outputRow.push(serializeCell(worksheet.getCell(row, column), includeStyles));
    }
    cells.push(outputRow);
  }

  return { path: resolvedPath, sheet: sheetName, range: rangeAddress, cells };
}

export async function writeRange(spec: ExcelWriteSpec) {
  const { input, output } = resolveInputOutput(spec.path, spec.outputPath, EXCEL_EXTENSIONS);
  await resolveReadablePath(input, EXCEL_EXTENSIONS);
  if (input !== output) await prepareOutputPath(output, EXCEL_EXTENSIONS, spec.overwrite);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(input);
  const worksheet = workbook.getWorksheet(spec.sheet);
  if (!worksheet) throw new OfficeMcpError(`Worksheet not found: ${spec.sheet}`, "SHEET_NOT_FOUND");
  const start = parseCellAddress(spec.startCell);

  spec.values.forEach((values, rowOffset) => {
    values.forEach((value, columnOffset) => {
      const cell = worksheet.getCell(start.row + rowOffset, start.column + columnOffset);
      cell.value = value as ExcelJS.CellValue;
      if (spec.numberFormat) cell.numFmt = spec.numberFormat;
      if (spec.bold !== undefined) cell.font = { ...cell.font, bold: spec.bold };
      if (spec.fillColor) {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: spec.fillColor.replace(/^#/, "").padStart(8, "F") }
        };
      }
    });
  });

  workbook.modified = new Date();
  await workbook.xlsx.writeFile(output);
  return { path: output, sheet: spec.sheet, startCell: spec.startCell, rowsWritten: spec.values.length };
}

export async function setFormulas(
  filePath: string,
  sheetName: string,
  formulas: ExcelFormulaSpec[],
  outputPath?: string,
  overwrite = false
) {
  const { input, output } = resolveInputOutput(filePath, outputPath, EXCEL_EXTENSIONS);
  await resolveReadablePath(input, EXCEL_EXTENSIONS);
  if (input !== output) await prepareOutputPath(output, EXCEL_EXTENSIONS, overwrite);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(input);
  const worksheet = workbook.getWorksheet(sheetName);
  if (!worksheet) throw new OfficeMcpError(`Worksheet not found: ${sheetName}`, "SHEET_NOT_FOUND");

  for (const item of formulas) {
    parseCellAddress(item.cell);
    worksheet.getCell(item.cell).value = {
      formula: item.formula.replace(/^=/, ""),
      ...(item.result !== undefined ? { result: item.result } : {})
    } as ExcelJS.CellFormulaValue;
  }

  workbook.calcProperties.fullCalcOnLoad = true;
  workbook.modified = new Date();
  await workbook.xlsx.writeFile(output);
  return { path: output, sheet: sheetName, formulasWritten: formulas.length };
}
