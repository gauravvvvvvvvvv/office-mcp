import ExcelJS from "exceljs";
import { OfficeMcpError } from "./errors.js";
import { resolveReadablePath } from "./paths.js";

type Scalar = string | number | boolean | Date | null;

function columnNumber(label: string): number {
  return [...label.toUpperCase()].reduce((value, letter) => value * 26 + letter.charCodeAt(0) - 64, 0);
}

function parseRange(range: string) {
  const match = /^([A-Za-z]+)([1-9]\d*):([A-Za-z]+)([1-9]\d*)$/.exec(range);
  if (!match) throw new OfficeMcpError(`Invalid dataset range: ${range}`, "INVALID_RANGE");
  const firstColumn = columnNumber(match[1]);
  const firstRow = Number(match[2]);
  const lastColumn = columnNumber(match[3]);
  const lastRow = Number(match[4]);
  if (lastColumn < firstColumn || lastRow < firstRow) throw new OfficeMcpError(`Reversed dataset range: ${range}`, "INVALID_RANGE");
  return { firstColumn, firstRow, lastColumn, lastRow };
}

function cellScalar(cell: ExcelJS.Cell): { value: Scalar; formula: boolean; error: boolean } {
  let value: unknown = cell.value;
  let formula = false;
  if (value && typeof value === "object" && "formula" in value) {
    formula = true;
    value = "result" in value ? value.result : null;
  }
  if (value instanceof Date) return { value, formula, error: false };
  if (value && typeof value === "object" && "error" in value) return { value: null, formula, error: true };
  if (value === undefined || value === null || value === "") return { value: null, formula, error: false };
  if (typeof value === "number" || typeof value === "string" || typeof value === "boolean") {
    return { value, formula, error: false };
  }
  return { value: cell.text || null, formula, error: false };
}

function percentile(sorted: number[], fraction: number): number | null {
  if (!sorted.length) return null;
  const position = (sorted.length - 1) * fraction;
  const lower = Math.floor(position);
  const upper = Math.ceil(position);
  return sorted[lower] + (sorted[upper] - sorted[lower]) * (position - lower);
}

function correlation(rows: Scalar[][], leftIndex: number, rightIndex: number) {
  let count = 0;
  let sumX = 0;
  let sumY = 0;
  let sumXX = 0;
  let sumYY = 0;
  let sumXY = 0;
  for (const row of rows) {
    const x = row[leftIndex];
    const y = row[rightIndex];
    if (typeof x !== "number" || typeof y !== "number" || !Number.isFinite(x) || !Number.isFinite(y)) continue;
    count++;
    sumX += x;
    sumY += y;
    sumXX += x * x;
    sumYY += y * y;
    sumXY += x * y;
  }
  if (count < 3) return null;
  const covariance = sumXY - (sumX * sumY) / count;
  const varianceX = sumXX - (sumX * sumX) / count;
  const varianceY = sumYY - (sumY * sumY) / count;
  const coefficient = varianceX > 0 && varianceY > 0 ? covariance / Math.sqrt(varianceX * varianceY) : null;
  return coefficient === null ? null : { coefficient, count };
}

export async function analyzeDataset(
  filePath: string,
  sheetName: string,
  rangeAddress?: string,
  hasHeader = true,
  maxRows = 10000,
  topValues = 5
) {
  const resolvedPath = await resolveReadablePath(filePath, [".xlsx"]);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(resolvedPath);
  const sheet = workbook.getWorksheet(sheetName);
  if (!sheet) throw new OfficeMcpError(`Worksheet not found: ${sheetName}`, "SHEET_NOT_FOUND");
  const range = rangeAddress
    ? parseRange(rangeAddress)
    : { firstColumn: 1, firstRow: 1, lastColumn: sheet.columnCount, lastRow: sheet.rowCount };
  if (range.lastColumn < range.firstColumn || range.lastRow < range.firstRow) {
    throw new OfficeMcpError("The worksheet is empty", "EMPTY_DATASET");
  }
  const columnCount = range.lastColumn - range.firstColumn + 1;
  if (columnCount > 100) throw new OfficeMcpError("Analyze at most 100 columns at a time; specify a narrower range", "DATASET_TOO_WIDE");
  const firstDataRow = range.firstRow + (hasHeader ? 1 : 0);
  const availableRows = Math.max(0, range.lastRow - firstDataRow + 1);
  const rowsAnalyzed = Math.min(availableRows, maxRows);
  const names = Array.from({ length: columnCount }, (_, index) => {
    const column = range.firstColumn + index;
    const header = hasHeader ? sheet.getCell(range.firstRow, column).text.trim() : "";
    return header || `Column ${sheet.getCell(1, column).address.replace(/\d+$/, "")}`;
  });
  const uniqueNames = names.map((name, index) => {
    const previous = names.slice(0, index).filter((candidate) => candidate === name).length;
    return previous ? `${name} (${previous + 1})` : name;
  });
  const valuesByColumn: Scalar[][] = uniqueNames.map(() => []);
  const formulaCounts = uniqueNames.map(() => 0);
  const errorCounts = uniqueNames.map(() => 0);
  const rowValues: Scalar[][] = [];
  const fingerprints = new Set<string>();
  let duplicateRows = 0;
  let emptyRows = 0;

  for (let rowNumber = firstDataRow; rowNumber < firstDataRow + rowsAnalyzed; rowNumber++) {
    const values = uniqueNames.map((_, index) => {
      const scalar = cellScalar(sheet.getCell(rowNumber, range.firstColumn + index));
      if (scalar.formula) formulaCounts[index]++;
      if (scalar.error) errorCounts[index]++;
      valuesByColumn[index].push(scalar.value);
      return scalar.value;
    });
    rowValues.push(values);
    if (values.every((value) => value === null)) {
      emptyRows++;
      continue;
    }
    const fingerprint = JSON.stringify(values.map((value) => value instanceof Date ? value.toISOString() : value));
    if (fingerprints.has(fingerprint)) duplicateRows++;
    else fingerprints.add(fingerprint);
  }

  const columns = uniqueNames.map((name, index) => {
    const values = valuesByColumn[index];
    const numeric = values.filter((value): value is number => typeof value === "number" && Number.isFinite(value)).sort((a, b) => a - b);
    const dates = values.filter((value): value is Date => value instanceof Date).sort((a, b) => a.getTime() - b.getTime());
    const text = values.filter((value): value is string => typeof value === "string");
    const booleans = values.filter((value): value is boolean => typeof value === "boolean");
    const missing = values.filter((value) => value === null).length;
    const kinds = [numeric.length && "number", dates.length && "date", text.length && "text", booleans.length && "boolean"].filter(Boolean);
    const counts = new Map<string, number>();
    for (const value of text) counts.set(value, (counts.get(value) ?? 0) + 1);
    const q1 = percentile(numeric, 0.25);
    const median = percentile(numeric, 0.5);
    const q3 = percentile(numeric, 0.75);
    const lowerFence = q1 !== null && q3 !== null ? q1 - 1.5 * (q3 - q1) : null;
    const upperFence = q1 !== null && q3 !== null ? q3 + 1.5 * (q3 - q1) : null;
    const outlierCount = lowerFence === null || upperFence === null
      ? 0 : numeric.filter((value) => value < lowerFence || value > upperFence).length;
    const sum = numeric.reduce((total, value) => total + value, 0);
    const mean = numeric.length ? sum / numeric.length : null;
    const standardDeviation = numeric.length > 1 && mean !== null
      ? Math.sqrt(numeric.reduce((total, value) => total + (value - mean) ** 2, 0) / (numeric.length - 1))
      : null;
    return {
      name,
      type: kinds.length === 0 ? "empty" : kinds.length === 1 ? kinds[0] : "mixed",
      count: values.length - missing,
      missing,
      missingPercent: values.length ? Math.round((10000 * missing) / values.length) / 100 : 0,
      formulaCount: formulaCounts[index],
      errorCount: errorCounts[index],
      numeric: numeric.length ? { count: numeric.length, min: numeric[0], max: numeric.at(-1)!, sum, mean, median, q1, q3, standardDeviation, outlierCount } : null,
      dates: dates.length ? { count: dates.length, earliest: dates[0].toISOString(), latest: dates.at(-1)!.toISOString() } : null,
      text: text.length ? {
        count: text.length,
        distinctCount: counts.size,
        topValues: [...counts.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
          .slice(0, topValues).map(([value, count]) => ({ value, count }))
      } : null
    };
  });

  const numericIndices = columns.map((column, index) => column.type === "number" ? index : -1).filter((index) => index >= 0).slice(0, 20);
  const correlations: Array<{ left: string; right: string; coefficient: number; pairedRows: number }> = [];
  for (let left = 0; left < numericIndices.length; left++) {
    for (let right = left + 1; right < numericIndices.length; right++) {
      const leftIndex = numericIndices[left];
      const rightIndex = numericIndices[right];
      const result = correlation(rowValues, leftIndex, rightIndex);
      if (result !== null) correlations.push({
        left: uniqueNames[leftIndex], right: uniqueNames[rightIndex],
        coefficient: Math.round(result.coefficient * 10000) / 10000,
        pairedRows: result.count
      });
    }
  }
  correlations.sort((left, right) => Math.abs(right.coefficient) - Math.abs(left.coefficient));

  return {
    kind: "excel_dataset_analysis",
    path: resolvedPath,
    sheet: sheetName,
    range: rangeAddress ?? (sheet.rowCount && sheet.columnCount ? `A1:${sheet.getCell(sheet.rowCount, sheet.columnCount).address}` : null),
    hasHeader,
    rowsAnalyzed,
    availableRows,
    truncated: rowsAnalyzed < availableRows,
    emptyRows,
    duplicateRows,
    columns,
    strongestCorrelations: correlations.slice(0, 10),
    cautions: [
      "Correlations are descriptive, not evidence of causation.",
      "Outliers use the 1.5×IQR rule; they are review candidates, not automatically errors.",
      "Formula cells without cached results count as missing until recalculated in Excel."
    ]
  };
}

type QueryValue = string | number | boolean | null;
type QueryFilter = {
  column: string;
  operator: "eq" | "neq" | "gt" | "gte" | "lt" | "lte" | "contains" | "in" | "isBlank" | "notBlank";
  value?: QueryValue;
  values?: QueryValue[];
};
type QueryMetric = { column: string; aggregation: "count" | "countDistinct" | "sum" | "average" | "min" | "max"; as?: string };

export interface DatasetQuerySpec {
  path: string;
  sheet: string;
  range?: string;
  hasHeader?: boolean;
  filters?: QueryFilter[];
  groupBy?: string[];
  metrics?: QueryMetric[];
  sort?: Array<{ column: string; direction: "asc" | "desc" }>;
  maxRows?: number;
  limit?: number;
}

function queryValue(value: Scalar): QueryValue {
  return value instanceof Date ? value.toISOString() : value;
}

function compareValues(left: QueryValue, right: QueryValue): number {
  if (left === right) return 0;
  if (left === null) return -1;
  if (right === null) return 1;
  if (typeof left === "number" && typeof right === "number") return left - right;
  return String(left).localeCompare(String(right), undefined, { numeric: true });
}

function passesFilter(value: QueryValue, filter: QueryFilter): boolean {
  switch (filter.operator) {
    case "isBlank": return value === null;
    case "notBlank": return value !== null;
    case "eq": return value === filter.value;
    case "neq": return value !== filter.value;
    case "gt": return value !== null && filter.value != null && compareValues(value, filter.value) > 0;
    case "gte": return value !== null && filter.value != null && compareValues(value, filter.value) >= 0;
    case "lt": return value !== null && filter.value != null && compareValues(value, filter.value) < 0;
    case "lte": return value !== null && filter.value != null && compareValues(value, filter.value) <= 0;
    case "contains": return value !== null && String(value).toLowerCase().includes(String(filter.value ?? "").toLowerCase());
    case "in": return (filter.values ?? []).includes(value);
  }
}

export async function queryDataset(spec: DatasetQuerySpec) {
  const resolvedPath = await resolveReadablePath(spec.path, [".xlsx"]);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(resolvedPath);
  const sheet = workbook.getWorksheet(spec.sheet);
  if (!sheet) throw new OfficeMcpError(`Worksheet not found: ${spec.sheet}`, "SHEET_NOT_FOUND");
  const range = spec.range
    ? parseRange(spec.range)
    : { firstColumn: 1, firstRow: 1, lastColumn: sheet.columnCount, lastRow: sheet.rowCount };
  const columnCount = range.lastColumn - range.firstColumn + 1;
  if (columnCount <= 0 || range.lastRow < range.firstRow) throw new OfficeMcpError("The worksheet is empty", "EMPTY_DATASET");
  if (columnCount > 100) throw new OfficeMcpError("Query at most 100 columns at a time", "DATASET_TOO_WIDE");
  const hasHeader = spec.hasHeader ?? true;
  const names = Array.from({ length: columnCount }, (_, index) => {
    const column = range.firstColumn + index;
    const header = hasHeader ? sheet.getCell(range.firstRow, column).text.trim() : "";
    return header || `Column ${sheet.getCell(1, column).address.replace(/\d+$/, "")}`;
  });
  const headers = names.map((name, index) => {
    const previous = names.slice(0, index).filter((candidate) => candidate === name).length;
    return previous ? `${name} (${previous + 1})` : name;
  });
  const indexByHeader = new Map(headers.map((name, index) => [name, index]));
  for (const name of [
    ...(spec.filters ?? []).map((filter) => filter.column),
    ...(spec.groupBy ?? []),
    ...(spec.metrics ?? []).filter((metric) => metric.column !== "*").map((metric) => metric.column),
  ]) {
    if (!indexByHeader.has(name)) throw new OfficeMcpError(`Dataset column not found: ${name}`, "COLUMN_NOT_FOUND");
  }
  const firstDataRow = range.firstRow + (hasHeader ? 1 : 0);
  const availableRows = Math.max(0, range.lastRow - firstDataRow + 1);
  const rowsScanned = Math.min(availableRows, spec.maxRows ?? 10000);
  const matched: Array<Record<string, QueryValue>> = [];
  for (let row = firstDataRow; row < firstDataRow + rowsScanned; row++) {
    const record = Object.fromEntries(headers.map((name, index) => [
      name, queryValue(cellScalar(sheet.getCell(row, range.firstColumn + index)).value)
    ])) as Record<string, QueryValue>;
    if ((spec.filters ?? []).every((filter) => passesFilter(record[filter.column], filter))) matched.push(record);
  }

  const groupBy = spec.groupBy ?? [];
  const metrics = spec.metrics?.length ? spec.metrics : groupBy.length ? [{ column: "*", aggregation: "count" as const, as: "count" }] : [];
  let results: Array<Record<string, QueryValue>>;
  if (groupBy.length || metrics.length) {
    const groups = new Map<string, Array<Record<string, QueryValue>>>();
    if (!groupBy.length) groups.set("[]", []);
    for (const record of matched) {
      const key = JSON.stringify(groupBy.map((name) => record[name]));
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key)!.push(record);
      if (groups.size > 5000) throw new OfficeMcpError("Query produced over 5,000 groups; narrow the filter", "TOO_MANY_GROUPS");
    }
    results = [...groups.values()].map((records) => {
      const result: Record<string, QueryValue> = {};
      for (const name of groupBy) result[name] = records[0][name];
      for (const metric of metrics) {
        const alias = metric.as ?? `${metric.aggregation}_${metric.column}`;
        const values = metric.column === "*" ? [] : records.map((record) => record[metric.column]).filter((value) => value !== null);
        const numbers = values.filter((value): value is number => typeof value === "number" && Number.isFinite(value));
        if (metric.aggregation === "count") result[alias] = metric.column === "*" ? records.length : values.length;
        else if (metric.aggregation === "countDistinct") result[alias] = new Set(values).size;
        else if (metric.aggregation === "sum") result[alias] = numbers.length ? numbers.reduce((sum, value) => sum + value, 0) : null;
        else if (metric.aggregation === "average") result[alias] = numbers.length ? numbers.reduce((sum, value) => sum + value, 0) / numbers.length : null;
        else if (metric.aggregation === "min") result[alias] = numbers.length ? numbers.reduce((low, value) => Math.min(low, value), Infinity) : null;
        else result[alias] = numbers.length ? numbers.reduce((high, value) => Math.max(high, value), -Infinity) : null;
      }
      return result;
    });
  } else {
    results = matched;
  }
  for (const sort of spec.sort ?? []) {
    if (results.length && !(sort.column in results[0])) {
      throw new OfficeMcpError(`Sort column not found in results: ${sort.column}`, "COLUMN_NOT_FOUND");
    }
  }
  results.sort((left, right) => {
    for (const sort of spec.sort ?? []) {
      const order = compareValues(left[sort.column], right[sort.column]);
      if (order) return sort.direction === "desc" ? -order : order;
    }
    return 0;
  });
  const limit = spec.limit ?? 100;
  return {
    kind: "excel_dataset_query",
    path: resolvedPath,
    sheet: spec.sheet,
    columns: results.length ? Object.keys(results[0]) : [...groupBy, ...metrics.map((metric) => metric.as ?? `${metric.aggregation}_${metric.column}`)],
    availableRows,
    rowsScanned,
    matchedRows: matched.length,
    resultCount: results.length,
    truncatedInput: rowsScanned < availableRows,
    truncatedOutput: results.length > limit,
    rows: results.slice(0, limit)
  };
}
