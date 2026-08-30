/**
 * Builds the workbook snapshot that is injected into the system prompt each
 * turn, so the model always has fresh context without dumping whole sheets.
 */
/* global Excel */

/** Sheets past this many are summarized by name only, to bound prompt size. */
const MAX_DETAILED_SHEETS = 12;

export async function buildWorkbookSummary(): Promise<string> {
  try {
    return await Excel.run(async (ctx) => {
      const wb = ctx.workbook;
      wb.load("name");
      const sheets = wb.worksheets;
      sheets.load("name");
      const selected = wb.getSelectedRange();
      selected.load("address");
      await ctx.sync();

      const lines: string[] = [];
      lines.push("Workbook: " + (wb.name || "(unsaved)"));
      lines.push("Selection: " + selected.address);

      const detailed = sheets.items.slice(0, MAX_DETAILED_SHEETS);
      for (const ws of detailed) {
        const used = ws.getUsedRangeOrNullObject(true);
        used.load("isNullObject,address,rowCount,columnCount,rowIndex,columnIndex");
        const tables = ws.tables;
        tables.load("name");
        await ctx.sync();

        if (used.isNullObject) {
          lines.push("- Sheet '" + ws.name + "': empty");
          continue;
        }
        const addr = used.address.replace(/^[^!]*!/, "");
        lines.push(
          "- Sheet '" +
            ws.name +
            "': used range " +
            addr +
            " (" +
            used.rowCount +
            " rows x " +
            used.columnCount +
            " cols)"
        );
        for (const t of tables.items) {
          lines.push("  - table '" + t.name + "'");
        }
        // The used range does not necessarily start at A1, so the sample must
        // be anchored to its own row/column index or it reads empty cells.
        const sample = ws.getRangeByIndexes(
          used.rowIndex,
          used.columnIndex,
          Math.max(1, Math.min(used.rowCount, 4)),
          Math.max(1, Math.min(used.columnCount, 8))
        );
        sample.load("address,values");
        await ctx.sync();
        lines.push(
          "  - " + sample.address.replace(/^[^!]*!/, "") + ": " + JSON.stringify(sample.values)
        );
      }
      if (sheets.items.length > detailed.length) {
        lines.push(
          "- " +
            (sheets.items.length - detailed.length) +
            " more sheet(s): " +
            sheets.items
              .slice(MAX_DETAILED_SHEETS)
              .map((w) => "'" + w.name + "'")
              .join(", ") +
            " (use get_workbook_info for details)"
        );
      }

      return lines.join("\n");
    });
  } catch (e) {
    return "(Workbook snapshot unavailable: " + String((e as Error)?.message ?? e) + ")";
  }
}

export function buildSystemPrompt(workbookSummary: string, confirmWrites: boolean): string {
  return [
    "You are ExcelLocal, a capable assistant embedded in an Excel task pane, similar to an AI copilot for spreadsheets.",
    "You work directly with the user's live workbook through tools. Current time: " +
      new Date().toString(),
    "",
    "# Current workbook snapshot",
    workbookSummary,
    "",
    "# Rules",
    "- Use tools to inspect real data before answering questions about the workbook; never guess cell contents.",
    "- Keep reads small and targeted; read_range is capped at 1500 cells and will refuse larger ranges.",
    "- After writing values or formulas, verify the result with read_range before telling the user it is done.",
    "- Only change what the user asked for; keep formatting minimal unless requested.",
    "- When writing formulas, use valid Excel formulas with real A1 references that match the sheet layout.",
    "- Sheets in formulas need quotes when the name contains spaces: ='Q1 Sales'!B2",
    "- Pass the sheet name in the 'sheet' argument and a bare A1 address (B2:D10) in 'address' - do not repeat the sheet inside the address.",
    "- For write_range/write_formulas you may pass just the top-left cell as the address and a full 2D array; the range is sized to the data.",
    confirmWrites
      ? "- Writing tools require user confirmation; if the user declined a write, do not retry it without asking."
      : "- Write tools execute immediately; double-check addresses before writing.",
    "- If a tool returns an error, read it: it usually names the exact problem (wrong sheet name, wrong shape) and how to fix it.",
    "- If a request is ambiguous, state your assumption and proceed; ask only when the risk is high (e.g. deleting data).",
    "- Answer in the user's language, be concise, and show results (values, formulas, addresses) rather than describing them vaguely.",
  ].join("\n");
}
