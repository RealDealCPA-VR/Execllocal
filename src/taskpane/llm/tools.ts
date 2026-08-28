/**
 * Workbook tool definitions exposed to the model (OpenAI function-calling
 * format) plus the registry of which tools mutate the workbook.
 */

function objProps(props: Record<string, unknown>, required: string[]) {
  return { type: "object", properties: props, required };
}

const RANGE_PROPS = {
  sheet: { type: "string", description: "Exact worksheet name" },
  address: { type: "string", description: "A1-style range, e.g. B2:D10" },
};

export const TOOL_SCHEMAS: unknown[] = [
  {
    type: "function",
    function: {
      name: "get_workbook_info",
      description:
        "Overview of the workbook: sheet names with used-range dimensions, first rows sample, tables, and the current selection address.",
      parameters: objProps({}, []),
    },
  },
  {
    type: "function",
    function: {
      name: "get_selection",
      description: "Values and formulas of the range the user currently has selected (capped in size).",
      parameters: objProps({}, []),
    },
  },
  {
    type: "function",
    function: {
      name: "read_range",
      description:
        "Read values (and formulas, if any) from a rectangular A1 range. Always inspect data before answering questions about it, and verify after writing.",
      parameters: objProps(
        {
          ...RANGE_PROPS,
          include_formulas: { type: "boolean", description: "Include formulas when present (default true)" },
        },
        ["sheet", "address"]
      ),
    },
  },
  {
    type: "function",
    function: {
      name: "write_range",
      description: "Write a 2D array of literal values into a range. Rows x columns must fit the target shape (a single row or column may be given as a 1D array).",
      parameters: objProps(
        {
          ...RANGE_PROPS,
          values: {
            type: "array",
            description: '2D array of values, e.g. [["Name","Qty"],["Widget",3]]',
            items: { type: "array" },
          },
        },
        ["sheet", "address", "values"]
      ),
    },
  },
  {
    type: "function",
    function: {
      name: "write_formulas",
      description: 'Write a 2D array of formulas (each a string starting with "=") into a range. Returns the computed values after recalculation.',
      parameters: objProps(
        {
          ...RANGE_PROPS,
          formulas: {
            type: "array",
            description: '2D array of formula strings, e.g. ["=SUM(B2:B10)"]',
            items: { type: "array" },
          },
        },
        ["sheet", "address", "formulas"]
      ),
    },
  },
  {
    type: "function",
    function: {
      name: "format_range",
      description: "Apply formatting to a range. Only provided properties are changed.",
      parameters: objProps(
        {
          ...RANGE_PROPS,
          bold: { type: "boolean" },
          italic: { type: "boolean" },
          font_size: { type: "number" },
          font_color: { type: "string", description: "Hex color like #FF0000 or a color name" },
          fill_color: { type: "string", description: "Hex background color like #FFFF00" },
          number_format: { type: "string", description: 'Excel number format code, e.g. "#,##0.00" or "0%"' },
          horizontal_alignment: { type: "string", description: "Left | Center | Right" },
          autofit_columns: { type: "boolean", description: "Auto-fit the column widths (default true when formatting)" },
        },
        ["sheet", "address"]
      ),
    },
  },
  {
    type: "function",
    function: {
      name: "create_sheet",
      description: "Create a new worksheet.",
      parameters: objProps({ name: { type: "string" } }, ["name"]),
    },
  },
  {
    type: "function",
    function: {
      name: "delete_sheet",
      description: "Delete a worksheet by name.",
      parameters: objProps({ name: { type: "string" } }, ["name"]),
    },
  },
  {
    type: "function",
    function: {
      name: "create_table",
      description: "Convert a range into a styled Excel table (adds filter buttons).",
      parameters: objProps(
        {
          ...RANGE_PROPS,
          name: { type: "string", description: "Optional table name (must be unique, no spaces)" },
          has_headers: { type: "boolean", description: "Whether the first row is a header row (default true)" },
        },
        ["sheet", "address"]
      ),
    },
  },
  {
    type: "function",
    function: {
      name: "create_chart",
      description: "Create a chart from a source range, positioned at a cell.",
      parameters: objProps(
        {
          sheet: { type: "string", description: "Worksheet that holds the data" },
          source_address: { type: "string", description: "A1 range with the data (include headers)" },
          chart_type: { type: "string", description: "column | bar | line | pie" },
          target_cell: { type: "string", description: "Top-left cell where the chart is placed, e.g. F2" },
        },
        ["sheet", "source_address", "chart_type", "target_cell"]
      ),
    },
  },
];

/** Tools that mutate the workbook — gated by the confirm-before-write setting. */
export const WRITE_TOOLS = new Set<string>([
  "write_range",
  "write_formulas",
  "format_range",
  "create_sheet",
  "delete_sheet",
  "create_table",
  "create_chart",
]);
