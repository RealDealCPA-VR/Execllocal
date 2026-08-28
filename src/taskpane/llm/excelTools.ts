/**
 * Office.js implementations of the workbook tools. Every tool runs inside a
 * single Excel.run batch and returns JSON-able results for the model plus a
 * short human-readable summary for the chat UI.
 */
/* global Excel */

const MAX_CELLS = 1500;

export interface ToolOutcome {
  result: unknown;
  summary: string;
}

function describeError(e: unknown): string {
  const err = e as { message?: string; code?: string };
  return err?.message ?? String(e);
}

/** Normalize a sheet name the way Excel does when addressing ranges. */
function cleanSheet(sheet: string): string {
  return String(sheet ?? "").replace(/^'+|'+$/g, "");
}

/** Reshape a flat or 2D array into rows x cols (row-major), or null if sizes mismatch. */
function reshape(input: unknown, rows: number, cols: number): unknown[][] | null {
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

function normalizeColor(c: unknown): string | undefined {
  if (typeof c !== "string" || !c.trim()) {
    return undefined;
  }
  const s = c.trim();
  return s.startsWith("#") ? s : "#" + s;
}

function alignment(value: unknown): string | undefined {
  const s = String(value ?? "").toLowerCase();
  if (s === "left") return "Left";
  if (s === "center" || s === "centre") return "Center";
  if (s === "right") return "Right";
  return undefined;
}

function chartType(name: unknown): Excel.ChartType | undefined {
  switch (String(name ?? "").toLowerCase()) {
    case "column":
      return "ColumnClustered" as Excel.ChartType;
    case "bar":
      return "BarClustered" as Excel.ChartType;
    case "line":
      return "Line" as Excel.ChartType;
    case "pie":
      return "Pie" as Excel.ChartType;
    default:
      return undefined;
  }
}

interface Dim { address: string; rows: number; cols: number }

async function loadDims(ws: Excel.Worksheet, address: string): Promise<Dim> {
  const range = ws.getRange(address);
  range.load("address,rowCount,columnCount");
  await range.context.sync();
  return { address: range.address, rows: range.rowCount, cols: range.columnCount };
}

export const excelExecutor = {
  async execute(name: string, args: Record<string, unknown>): Promise<ToolOutcome> {
    return Excel.run(async (ctx) => {
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
  },
};

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
    used.load("isNullObject,address,rowCount,columnCount");
    let sample: Excel.Range | null = null;
    await ctx.sync();
    if (!used.isNullObject) {
      sample = ws.getRangeByIndexes(0, 0, Math.min(used.rowCount, 5), Math.min(used.columnCount, 8));
      sample.load("values");
      await ctx.sync();
    }
    sheetInfo.push({
      name: ws.name,
      usedRange: used.isNullObject ? null : used.address.replace(/^[^!]*!/, ""),
      rows: used.isNullObject ? 0 : used.rowCount,
      columns: used.isNullObject ? 0 : used.columnCount,
      sampleTopLeft: used.isNullObject ? null : sample!.values,
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
  range.load("address,rowCount,columnCount,values,formulas");
  await ctx.sync();
  if (range.rowCount * range.columnCount > MAX_CELLS) {
    return {
      result: { address: range.address, error: "Selection too large to return; ask the user or use read_range on a smaller range." },
      summary: "Selection too large",
    };
  }
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

async function readRange(ctx: Excel.RequestContext, args: Record<string, unknown>): Promise<ToolOutcome> {
  const sheetName = cleanSheet(String(args.sheet ?? ""));
  const address = String(args.address ?? "");
  const ws = ctx.workbook.worksheets.getItem(sheetName);
  const dims = await loadDims(ws, address);
  if (dims.rows * dims.cols > MAX_CELLS) {
    return {
      result: { error: "Range " + dims.address + " is " + dims.rows + "x" + dims.cols + " (" + dims.rows * dims.cols + " cells); limit is " + MAX_CELLS + ". Use a narrower range or aggregate the data first." },
      summary: "Read refused: too large",
    };
  }
  const range = ws.getRange(address);
  range.load("values,formulas");
  await ctx.sync();
  const includeFormulas = args.include_formulas === undefined ? true : Boolean(args.include_formulas);
  return {
    result: {
      sheet: sheetName,
      address: dims.address.replace(/^[^!]*!/, ""),
      rowCount: dims.rows,
      columnCount: dims.cols,
      values: range.values,
      formulas: includeFormulas && hasFormulas(range.formulas) ? range.formulas : undefined,
    },
    summary: sheetName + "!" + dims.address.replace(/^[^!]*!/, "") + " (" + dims.rows + "x" + dims.cols + ") read",
  };
}

async function writeGrid(
  ctx: Excel.RequestContext,
  args: Record<string, unknown>,
  kind: "values" | "formulas"
): Promise<ToolOutcome> {
  const sheetName = cleanSheet(String(args.sheet ?? ""));
  const address = String(args.address ?? "");
  const grid = args[kind];
  if (!Array.isArray(grid) || grid.length === 0) {
    return { result: { error: kind + " must be a non-empty 2D array" }, summary: "Invalid input" };
  }
  const ws = ctx.workbook.worksheets.getItem(sheetName);
  const dims = await loadDims(ws, address);
  const shaped = reshape(grid, dims.rows, dims.cols);
  if (!shaped) {
    return {
      result: {
        error:
          "Shape mismatch: target " + dims.address + " is " + dims.rows + " rows x " + dims.cols + " columns (" +
          dims.rows * dims.cols + " cells) but " + kind + " contained " +
          (Array.isArray(grid) ? grid.length : 0) + " top-level entries. Send a 2D array matching the target shape.",
      },
      summary: "Shape mismatch",
    };
  }
  const range = ws.getRange(address);
  if (kind === "values") {
    range.values = shaped;
  } else {
    range.formulas = shaped;
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
      address: dims.address.replace(/^[^!]*!/, ""),
      rows: dims.rows,
      columns: dims.cols,
      computedValues: computed,
    },
    summary: "Wrote " + kind + " to " + sheetName + "!" + dims.address.replace(/^[^!]*!/, "") + " (" + dims.rows + "x" + dims.cols + ")",
  };
}

async function formatRange(ctx: Excel.RequestContext, args: Record<string, unknown>): Promise<ToolOutcome> {
  const sheetName = cleanSheet(String(args.sheet ?? ""));
  const address = String(args.address ?? "");
  const ws = ctx.workbook.worksheets.getItem(sheetName);
  const dims = await loadDims(ws, address);
  const range = ws.getRange(address);
  const applied: Record<string, unknown> = {};

  if (args.bold !== undefined) { range.format.font.bold = Boolean(args.bold); applied.bold = Boolean(args.bold); }
  if (args.italic !== undefined) { range.format.font.italic = Boolean(args.italic); applied.italic = Boolean(args.italic); }
  const size = Number(args.font_size);
  if (size > 0) { range.format.font.size = size; applied.font_size = size; }
  const fontColor = normalizeColor(args.font_color);
  if (fontColor) { range.format.font.color = fontColor; applied.font_color = fontColor; }
  const fillColor = normalizeColor(args.fill_color);
  if (fillColor) { range.format.fill.color = fillColor; applied.fill_color = fillColor; }
  const align = alignment(args.horizontal_alignment);
  if (align) { range.format.horizontalAlignment = align as Excel.HorizontalAlignment; applied.horizontal_alignment = align; }
  if (typeof args.number_format === "string" && args.number_format.trim()) {
    const fmt = args.number_format.trim();
    const grid: string[][] = [];
    for (let r = 0; r < dims.rows; r++) {
      grid.push(new Array(dims.cols).fill(fmt));
    }
    range.numberFormat = grid;
    applied.number_format = fmt;
  }
  if (args.autofit_columns === undefined || args.autofit_columns) {
    range.format.autofitColumns();
  }
  await ctx.sync();
  return {
    result: { sheet: sheetName, address: dims.address.replace(/^[^!]*!/, ""), applied },
    summary: "Formatted " + sheetName + "!" + dims.address.replace(/^[^!]*!/, ""),
  };
}

async function createSheet(ctx: Excel.RequestContext, args: Record<string, unknown>): Promise<ToolOutcome> {
  const name = String(args.name ?? "").trim();
  if (!name) {
    return { result: { error: "Sheet name is required" }, summary: "Invalid input" };
  }
  const ws = ctx.workbook.worksheets.add(name);
  ws.load("name");
  await ctx.sync();
  return { result: { created: ws.name }, summary: "Created sheet " + ws.name };
}

async function deleteSheet(ctx: Excel.RequestContext, args: Record<string, unknown>): Promise<ToolOutcome> {
  const name = String(args.name ?? "").trim();
  const ws = ctx.workbook.worksheets.getItem(name);
  ws.load("name");
  await ctx.sync();
  ws.delete();
  await ctx.sync();
  return { result: { deleted: name }, summary: "Deleted sheet " + name };
}

async function createTable(ctx: Excel.RequestContext, args: Record<string, unknown>): Promise<ToolOutcome> {
  const sheetName = cleanSheet(String(args.sheet ?? ""));
  const address = String(args.address ?? "");
  const hasHeaders = args.has_headers === undefined ? true : Boolean(args.has_headers);
  const ws = ctx.workbook.worksheets.getItem(sheetName);
  const table = ctx.workbook.tables.add(ws.getRange(address), hasHeaders);
  table.load("name");
  await ctx.sync();
  if (typeof args.name === "string" && args.name.trim()) {
    table.name = args.name.trim();
    await ctx.sync();
  }
  table.style = "TableStyleMedium2";
  await ctx.sync();
  return { result: { name: table.name, sheet: sheetName, address: address }, summary: "Created table " + table.name + " on " + sheetName };
}

async function createChart(ctx: Excel.RequestContext, args: Record<string, unknown>): Promise<ToolOutcome> {
  const sheetName = cleanSheet(String(args.sheet ?? ""));
  const type = chartType(args.chart_type);
  if (!type) {
    return { result: { error: "chart_type must be one of: column, bar, line, pie" }, summary: "Invalid chart type" };
  }
  const ws = ctx.workbook.worksheets.getItem(sheetName);
  const source = ws.getRange(String(args.source_address ?? ""));
  const chart = ws.charts.add(type, source, "Auto");
  chart.setPosition(String(args.target_cell ?? ""), String(args.target_cell ?? ""));
  chart.load("id");
  await ctx.sync();
  return {
    result: { chartId: chart.id, sheet: sheetName, type: String(args.chart_type) },
    summary: "Created " + String(args.chart_type) + " chart on " + sheetName + " at " + String(args.target_cell),
  };
}
