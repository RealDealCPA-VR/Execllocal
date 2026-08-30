/**
 * Office.js implementations of the workbook tools. Every tool runs inside a
 * single Excel.run batch and returns JSON-able results for the model plus a
 * short human-readable summary for the chat UI.
 */
/* global Excel */

const MAX_CELLS = 1500;
/** autofitColumns() on a huge range stalls Excel; skip it past this width. */
const MAX_AUTOFIT_COLUMNS = 200;

export interface ToolOutcome {
  result: unknown;
  summary: string;
}

/** Human-readable text for an Office.js / JS error, keeping the Excel error code. */
export function describeError(e: unknown): string {
  const err = e as { message?: string; code?: string; debugInfo?: { errorLocation?: string } };
  const base = err?.message ? String(err.message) : String(e);
  const code = err?.code ? " [" + err.code + "]" : "";
  const where = err?.debugInfo?.errorLocation ? " at " + err.debugInfo.errorLocation : "";
  return base + code + where;
}

function errorCode(e: unknown): string {
  return String((e as { code?: string })?.code ?? "");
}

/** Normalize a sheet name the way Excel does when addressing ranges. */
export function cleanSheet(sheet: string): string {
  return String(sheet ?? "")
    .trim()
    .replace(/^'+|'+$/g, "");
}

/**
 * Models frequently pass fully-qualified addresses ("Sheet1!A1:B2" or
 * "'Q1 Sales'!A1") even though the sheet is a separate argument. Strip the
 * sheet part so ws.getRange() never sees it.
 */
export function cleanAddress(address: unknown): string {
  const s = String(address ?? "").trim();
  const bang = s.lastIndexOf("!");
  return (bang >= 0 ? s.slice(bang + 1) : s).trim();
}

/** Rows/columns implied by a value grid, or null when it is not rectangular. */
export function gridShape(input: unknown): { rows: number; cols: number } | null {
  if (!Array.isArray(input) || input.length === 0) {
    return null;
  }
  const rows = input as unknown[];
  if (rows.every((r) => Array.isArray(r))) {
    const widths = (rows as unknown[][]).map((r) => r.length);
    const cols = widths[0];
    if (!cols || widths.some((w) => w !== cols)) {
      return null; // ragged: fall back to flat reshaping
    }
    return { rows: rows.length, cols };
  }
  if (rows.some((r) => Array.isArray(r))) {
    return null; // mixed rows and scalars
  }
  return { rows: 1, cols: rows.length };
}

/** Reshape a flat or 2D array into rows x cols (row-major), or null if sizes mismatch. */
export function reshape(input: unknown, rows: number, cols: number): unknown[][] | null {
  const flat: unknown[] = Array.isArray(input)
    ? (input as unknown[]).flatMap((row) => (Array.isArray(row) ? (row as unknown[]) : [row]))
    : [input];
  if (flat.length !== rows * cols) {
    return null;
  }
  const out: unknown[][] = [];
  for (let r = 0; r < rows; r++) {
    out.push(flat.slice(r * cols, r * cols + cols));
  }
  return out;
}

const CSS_HEX = /^#?([0-9a-f]{3}|[0-9a-f]{6})$/i;

/**
 * Accepts "#FF0000", "FF0000", "f00" and plain CSS color names ("red").
 * Bare names must NOT get a "#" glued on: Excel rejects "#red".
 */
export function normalizeColor(c: unknown): string | undefined {
  if (typeof c !== "string" || !c.trim()) {
    return undefined;
  }
  const s = c.trim();
  if (CSS_HEX.test(s)) {
    return s.startsWith("#") ? s : "#" + s;
  }
  return /^[a-z]+$/i.test(s) ? s : undefined; // color name, or unusable
}

export function alignment(value: unknown): string | undefined {
  const s = String(value ?? "").toLowerCase();
  if (s === "left") return "Left";
  if (s === "center" || s === "centre") return "Center";
  if (s === "right") return "Right";
  return undefined;
}

export function chartType(name: unknown): string | undefined {
  switch (String(name ?? "").toLowerCase()) {
    case "column":
    case "columnclustered":
      return "ColumnClustered";
    case "bar":
    case "barclustered":
      return "BarClustered";
    case "line":
      return "Line";
    case "pie":
      return "Pie";
    default:
      return undefined;
  }
}

/** Excel table names allow no spaces and cannot look like a cell reference. */
export function sanitizeTableName(raw: string): string | undefined {
  const name = raw.trim().replace(/[^\w.]/g, "_");
  if (!name || /^[0-9.]/.test(name) || /^[A-Z]{1,3}[0-9]{1,7}$/i.test(name)) {
    return undefined;
  }
  return name.slice(0, 255);
}

interface Dim {
  address: string;
  rows: number;
  cols: number;
  rowIndex: number;
  colIndex: number;
}

async function loadDims(ws: Excel.Worksheet, address: string): Promise<Dim> {
  const range = ws.getRange(address);
  range.load("address,rowCount,columnCount,rowIndex,columnIndex");
  await range.context.sync();
  return {
    address: range.address,
    rows: range.rowCount,
    cols: range.columnCount,
    rowIndex: range.rowIndex,
    colIndex: range.columnIndex,
  };
}

/** Strip the "Sheet1!" prefix Excel puts on returned addresses. */
function localAddress(address: string): string {
  return address.replace(/^[^!]*!/, "");
}

export const excelExecutor = {
  async execute(name: string, args: Record<string, unknown>): Promise<ToolOutcome> {
    try {
      return await Excel.run(async (ctx) => {
        switch (name) {
          case "get_workbook_info":
            return getWorkbookInfo(ctx);
          case "get_selection":
            return getSelection(ctx);
          case "read_range":
            return readRange(ctx, args);
          case "write_range":
            return writeGrid(ctx, args, "values");
          case "write_formulas":
            return writeGrid(ctx, args, "formulas");
          case "format_range":
            return formatRange(ctx, args);
          case "create_sheet":
            return createSheet(ctx, args);
          case "delete_sheet":
            return deleteSheet(ctx, args);
          case "create_table":
            return createTable(ctx, args);
          case "create_chart":
            return createChart(ctx, args);
          default:
            return { result: { error: "Unknown tool: " + name }, summary: "Unknown tool" };
        }
      });
    } catch (e) {
      // Rethrow with a message the model can act on: Excel's own text is terse,
      // and "ItemNotFound" is almost always a mistyped sheet name.
      throw new Error(await explainToolError(e, args));
    }
  },
};

/** Best-effort enrichment of an Office.js failure with actionable context. */
async function explainToolError(e: unknown, args: Record<string, unknown>): Promise<string> {
  const base = describeError(e);
  if (errorCode(e) !== "ItemNotFound") {
    return base;
  }
  const wanted = cleanSheet(String(args.sheet ?? args.name ?? ""));
  try {
    const names = await Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      sheets.load("name");
      await ctx.sync();
      return sheets.items.map((w) => w.name);
    });
    return (
      base +
      (wanted ? " - no worksheet named " + JSON.stringify(wanted) + "." : "") +
      " Existing sheets: " +
      names.map((n) => JSON.stringify(n)).join(", ") +
      "."
    );
  } catch {
    return base;
  }
}

/**
 * Top-left sample of a sheet's data. The used range does NOT necessarily start
 * at A1, so the sample must start at its own row/column index or it reads a
 * block of empty cells and the model is told the sheet looks blank.
 */
function sampleOfUsedRange(
  ws: Excel.Worksheet,
  used: Excel.Range,
  maxRows: number,
  maxCols: number
): Excel.Range {
  return ws.getRangeByIndexes(
    used.rowIndex,
    used.columnIndex,
    Math.max(1, Math.min(used.rowCount, maxRows)),
    Math.max(1, Math.min(used.columnCount, maxCols))
  );
}

async function getWorkbookInfo(ctx: Excel.RequestContext): Promise<ToolOutcome> {
  const wb = ctx.workbook;
  wb.load("name");
  const sheets = wb.worksheets;
  sheets.load("name");
  const tables = wb.tables;
  tables.load("name,worksheet/name");
  const selected = wb.getSelectedRange();
  selected.load("address");
  await ctx.sync();

  const sheetInfo: Array<Record<string, unknown>> = [];
  for (const ws of sheets.items) {
    const used = ws.getUsedRangeOrNullObject(true);
    used.load("isNullObject,address,rowCount,columnCount,rowIndex,columnIndex");
    let sample: Excel.Range | null = null;
    await ctx.sync();
    if (!used.isNullObject) {
      sample = sampleOfUsedRange(ws, used, 5, 8);
      sample.load("address,values");
      await ctx.sync();
    }
    sheetInfo.push({
      name: ws.name,
      usedRange: used.isNullObject ? null : localAddress(used.address),
      rows: used.isNullObject ? 0 : used.rowCount,
      columns: used.isNullObject ? 0 : used.columnCount,
      sampleAddress: sample ? localAddress(sample.address) : null,
      sampleTopLeft: sample ? sample.values : null,
    });
  }

  return {
    result: {
      workbook: wb.name,
      selection: selected.address,
      sheets: sheetInfo,
      tables: tables.items.map((t) => ({ name: t.name, sheet: t.worksheet.name })),
    },
    summary: "Workbook overview read",
  };
}

async function getSelection(ctx: Excel.RequestContext): Promise<ToolOutcome> {
  const range = ctx.workbook.getSelectedRange();
  // Load dimensions first: a whole-column selection must not materialize 1M rows.
  range.load("address,rowCount,columnCount");
  await ctx.sync();
  if (range.rowCount * range.columnCount > MAX_CELLS) {
    return {
      result: {
        address: range.address,
        rowCount: range.rowCount,
        columnCount: range.columnCount,
        error: "Selection too large to return; ask the user or use read_range on a smaller range.",
      },
      summary: "Selection too large",
    };
  }
  range.load("values,formulas");
  await ctx.sync();
  return {
    result: {
      address: range.address,
      rowCount: range.rowCount,
      columnCount: range.columnCount,
      values: range.values,
      formulas: hasFormulas(range.formulas) ? range.formulas : undefined,
    },
    summary: "Selection " + range.address + " read",
  };
}

function hasFormulas(formulas: unknown[][]): boolean {
  return formulas.some((row) => row.some((c) => typeof c === "string" && c.startsWith("=")));
}

async function readRange(
  ctx: Excel.RequestContext,
  args: Record<string, unknown>
): Promise<ToolOutcome> {
  const sheetName = cleanSheet(String(args.sheet ?? ""));
  const address = cleanAddress(args.address);
  if (!address) {
    return { result: { error: "address is required, e.g. A1:D20" }, summary: "Invalid input" };
  }
  const ws = ctx.workbook.worksheets.getItem(sheetName);
  const dims = await loadDims(ws, address);
  if (dims.rows * dims.cols > MAX_CELLS) {
    return {
      result: {
        error:
          "Range " +
          dims.address +
          " is " +
          dims.rows +
          "x" +
          dims.cols +
          " (" +
          dims.rows * dims.cols +
          " cells); limit is " +
          MAX_CELLS +
          ". Use a narrower range or aggregate the data first.",
      },
      summary: "Read refused: too large",
    };
  }
  const range = ws.getRange(address);
  range.load("values,formulas");
  await ctx.sync();
  const includeFormulas =
    args.include_formulas === undefined ? true : Boolean(args.include_formulas);
  return {
    result: {
      sheet: sheetName,
      address: localAddress(dims.address),
      rowCount: dims.rows,
      columnCount: dims.cols,
      values: range.values,
      formulas: includeFormulas && hasFormulas(range.formulas) ? range.formulas : undefined,
    },
    summary:
      sheetName + "!" + localAddress(dims.address) + " (" + dims.rows + "x" + dims.cols + ") read",
  };
}

async function writeGrid(
  ctx: Excel.RequestContext,
  args: Record<string, unknown>,
  kind: "values" | "formulas"
): Promise<ToolOutcome> {
  const sheetName = cleanSheet(String(args.sheet ?? ""));
  const address = cleanAddress(args.address);
  const grid = args[kind];
  if (!address) {
    return {
      result: { error: "address is required, e.g. A1 or B2:D10" },
      summary: "Invalid input",
    };
  }
  if (!Array.isArray(grid) || grid.length === 0) {
    return { result: { error: kind + " must be a non-empty 2D array" }, summary: "Invalid input" };
  }

  const ws = ctx.workbook.worksheets.getItem(sheetName);
  let dims = await loadDims(ws, address);

  // Models routinely pass only the top-left cell together with a full block of
  // data. Grow a 1x1 target to fit the grid instead of failing on shape.
  const shape = gridShape(grid);
  let targetAddress = address;
  if (dims.rows === 1 && dims.cols === 1 && shape && shape.rows * shape.cols > 1) {
    const grown = ws.getRangeByIndexes(dims.rowIndex, dims.colIndex, shape.rows, shape.cols);
    grown.load("address,rowCount,columnCount,rowIndex,columnIndex");
    await ctx.sync();
    dims = {
      address: grown.address,
      rows: grown.rowCount,
      cols: grown.columnCount,
      rowIndex: grown.rowIndex,
      colIndex: grown.columnIndex,
    };
    targetAddress = localAddress(grown.address);
  }

  if (dims.rows * dims.cols > MAX_CELLS) {
    return {
      result: {
        error:
          "Target " +
          dims.address +
          " is " +
          dims.rows * dims.cols +
          " cells; writes are capped at " +
          MAX_CELLS +
          " cells per call. Split the write into smaller ranges.",
      },
      summary: "Write refused: too large",
    };
  }

  const shaped = reshape(grid, dims.rows, dims.cols);
  if (!shaped) {
    return {
      result: {
        error:
          "Shape mismatch: target " +
          dims.address +
          " is " +
          dims.rows +
          " rows x " +
          dims.cols +
          " columns (" +
          dims.rows * dims.cols +
          " cells) but " +
          kind +
          " described " +
          (shape ? shape.rows + " rows x " + shape.cols + " columns" : "a ragged array") +
          ". Send a rectangular 2D array matching the target, or pass just the top-left cell as the address.",
      },
      summary: "Shape mismatch",
    };
  }

  const range = ws.getRange(targetAddress);
  if (kind === "values") {
    range.values = shaped;
  } else {
    // Excel requires formula cells to be strings; coerce so one stray number
    // does not fail the whole batch.
    range.formulas = shaped.map((row) =>
      row.map((c) => (c === null || c === undefined ? "" : String(c)))
    );
  }
  await ctx.sync();

  // Read back computed values so the model can verify formulas.
  let computed: unknown[][] | undefined;
  if (kind === "formulas") {
    range.load("values");
    await ctx.sync();
    computed = range.values;
  }
  return {
    result: {
      written: kind,
      sheet: sheetName,
      address: localAddress(dims.address),
      rows: dims.rows,
      columns: dims.cols,
      computedValues: computed,
    },
    summary:
      "Wrote " +
      kind +
      " to " +
      sheetName +
      "!" +
      localAddress(dims.address) +
      " (" +
      dims.rows +
      "x" +
      dims.cols +
      ")",
  };
}

async function formatRange(
  ctx: Excel.RequestContext,
  args: Record<string, unknown>
): Promise<ToolOutcome> {
  const sheetName = cleanSheet(String(args.sheet ?? ""));
  const address = cleanAddress(args.address);
  if (!address) {
    return { result: { error: "address is required, e.g. A1:D1" }, summary: "Invalid input" };
  }
  const ws = ctx.workbook.worksheets.getItem(sheetName);
  const dims = await loadDims(ws, address);

  if (dims.rows * dims.cols > MAX_CELLS && args.number_format !== undefined) {
    return {
      result: {
        error:
          "Range is " +
          dims.rows +
          "x" +
          dims.cols +
          " cells; number_format is applied per cell and is capped at " +
          MAX_CELLS +
          " cells. Narrow the range or apply the format manually.",
      },
      summary: "number_format refused: range too large",
    };
  }

  const range = ws.getRange(address);
  const applied: Record<string, unknown> = {};
  const ignored: string[] = [];

  if (args.bold !== undefined) {
    range.format.font.bold = Boolean(args.bold);
    applied.bold = Boolean(args.bold);
  }
  if (args.italic !== undefined) {
    range.format.font.italic = Boolean(args.italic);
    applied.italic = Boolean(args.italic);
  }
  const size = Number(args.font_size);
  if (Number.isFinite(size) && size > 0) {
    range.format.font.size = size;
    applied.font_size = size;
  }
  const fontColor = normalizeColor(args.font_color);
  if (fontColor) {
    range.format.font.color = fontColor;
    applied.font_color = fontColor;
  } else if (args.font_color !== undefined) {
    ignored.push("font_color");
  }
  const fillColor = normalizeColor(args.fill_color);
  if (fillColor) {
    range.format.fill.color = fillColor;
    applied.fill_color = fillColor;
  } else if (args.fill_color !== undefined) {
    ignored.push("fill_color");
  }
  const align = alignment(args.horizontal_alignment);
  if (align) {
    range.format.horizontalAlignment = align as Excel.HorizontalAlignment;
    applied.horizontal_alignment = align;
  } else if (args.horizontal_alignment !== undefined) {
    ignored.push("horizontal_alignment");
  }
  if (typeof args.number_format === "string" && args.number_format.trim()) {
    const fmt = args.number_format.trim();
    const grid: string[][] = [];
    for (let r = 0; r < dims.rows; r++) {
      grid.push(new Array(dims.cols).fill(fmt));
    }
    range.numberFormat = grid;
    applied.number_format = fmt;
  }
  // autofitColumns() is O(cells) inside Excel; skip it on very wide ranges.
  const wantsAutofit = args.autofit_columns === undefined || Boolean(args.autofit_columns);
  if (wantsAutofit && dims.cols <= MAX_AUTOFIT_COLUMNS) {
    range.format.autofitColumns();
    applied.autofit_columns = true;
  }
  await ctx.sync();
  return {
    result: {
      sheet: sheetName,
      address: localAddress(dims.address),
      applied,
      ...(ignored.length
        ? {
            ignored,
            note: "Unrecognized values ignored; use hex like #FF0000 or a plain color name such as red.",
          }
        : {}),
    },
    summary: "Formatted " + sheetName + "!" + localAddress(dims.address),
  };
}

async function createSheet(
  ctx: Excel.RequestContext,
  args: Record<string, unknown>
): Promise<ToolOutcome> {
  const name = String(args.name ?? "").trim();
  if (!name) {
    return { result: { error: "Sheet name is required" }, summary: "Invalid input" };
  }
  const ws = ctx.workbook.worksheets.add(name);
  ws.load("name");
  await ctx.sync();
  return { result: { created: ws.name }, summary: "Created sheet " + ws.name };
}

async function deleteSheet(
  ctx: Excel.RequestContext,
  args: Record<string, unknown>
): Promise<ToolOutcome> {
  const name = cleanSheet(String(args.name ?? ""));
  if (!name) {
    return { result: { error: "Sheet name is required" }, summary: "Invalid input" };
  }
  const sheets = ctx.workbook.worksheets;
  sheets.load("name");
  await ctx.sync();
  if (sheets.items.length <= 1) {
    return {
      result: {
        error:
          "Cannot delete the last remaining worksheet; a workbook must keep at least one sheet.",
      },
      summary: "Delete refused",
    };
  }
  const ws = sheets.getItem(name);
  ws.load("name");
  await ctx.sync();
  const actual = ws.name;
  ws.delete();
  await ctx.sync();
  return { result: { deleted: actual }, summary: "Deleted sheet " + actual };
}

async function createTable(
  ctx: Excel.RequestContext,
  args: Record<string, unknown>
): Promise<ToolOutcome> {
  const sheetName = cleanSheet(String(args.sheet ?? ""));
  const address = cleanAddress(args.address);
  if (!address) {
    return { result: { error: "address is required, e.g. A1:D20" }, summary: "Invalid input" };
  }
  const hasHeaders = args.has_headers === undefined ? true : Boolean(args.has_headers);
  const ws = ctx.workbook.worksheets.getItem(sheetName);
  const table = ctx.workbook.tables.add(ws.getRange(address), hasHeaders);
  table.style = "TableStyleMedium2";
  table.load("name");
  await ctx.sync();

  // A rejected name must not throw away the table that was already created.
  let renamed: string | undefined;
  let nameNote: string | undefined;
  if (typeof args.name === "string" && args.name.trim()) {
    const wanted = sanitizeTableName(args.name);
    if (!wanted) {
      nameNote =
        "Ignored table name " +
        JSON.stringify(args.name) +
        ": it must not start with a digit or look like a cell reference.";
    } else {
      try {
        table.name = wanted;
        await ctx.sync();
        renamed = wanted;
      } catch (e) {
        nameNote = "Kept the default table name (" + describeError(e) + ").";
        table.load("name");
        await ctx.sync();
      }
    }
  }

  return {
    result: {
      name: renamed ?? table.name,
      sheet: sheetName,
      address,
      hasHeaders,
      ...(nameNote ? { note: nameNote } : {}),
    },
    summary: "Created table " + (renamed ?? table.name) + " on " + sheetName,
  };
}

async function createChart(
  ctx: Excel.RequestContext,
  args: Record<string, unknown>
): Promise<ToolOutcome> {
  const sheetName = cleanSheet(String(args.sheet ?? ""));
  const type = chartType(args.chart_type);
  if (!type) {
    return {
      result: { error: "chart_type must be one of: column, bar, line, pie" },
      summary: "Invalid chart type",
    };
  }
  const sourceAddress = cleanAddress(args.source_address);
  const targetCell = cleanAddress(args.target_cell);
  if (!sourceAddress) {
    return {
      result: { error: "source_address is required, e.g. A1:B10" },
      summary: "Invalid input",
    };
  }
  const ws = ctx.workbook.worksheets.getItem(sheetName);
  const source = ws.getRange(sourceAddress);
  const chart = ws.charts.add(type as Excel.ChartType, source, "Auto" as Excel.ChartSeriesBy);
  // Anchor only: passing the same cell as both start AND end shrinks the chart
  // down to one cell, which looks to the user like nothing happened.
  if (targetCell) {
    chart.setPosition(targetCell);
  }
  chart.load("name");
  await ctx.sync();
  return {
    result: {
      chart: chart.name,
      sheet: sheetName,
      type: String(args.chart_type),
      source: sourceAddress,
      at: targetCell || null,
    },
    summary:
      "Created " +
      String(args.chart_type) +
      " chart on " +
      sheetName +
      (targetCell ? " at " + targetCell : ""),
  };
}
