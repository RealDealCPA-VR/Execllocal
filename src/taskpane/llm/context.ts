/**
 * Builds the workbook snapshot that is injected into the system prompt each
 * turn, so the model always has fresh context without dumping whole sheets.
 */
/* global Excel */

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

      for (const ws of sheets.items) {
        const used = ws.getUsedRangeOrNullObject(true);
        used.load("isNullObject,address,rowCount,columnCount");
        let sample: Excel.Range | null = null;
        await ctx.sync();
        if (used.isNullObject) {
          lines.push("- Sheet '" + ws.name + "': empty");
          continue;
        }
        const addr = used.address.replace(/^[^!]*!/, "");
        lines.push("- Sheet '" + ws.name + "': used range " + addr + " (" + used.rowCount + " rows x " + used.columnCount + " cols)");
        sample = ws.getRangeByIndexes(0, 0, Math.min(used.rowCount, 4), Math.min(used.columnCount, 8));
        sample.load("values");
        const tables = ws.tables;
        tables.load("name");
        await ctx.sync();
        for (const t of tables.items) {
          lines.push("  - table '" + t.name + "'");
        }
        lines.push("  - first rows (up to 4x8): " + JSON.stringify(sample.values));
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
    "You work directly with the user's live workbook through tools. Current time: " + new Date().toString(),
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
    confirmWrites
      ? "- Writing tools require user confirmation; if the user declined a write, do not retry it without asking."
      : "- Write tools execute immediately; double-check addresses before writing.",
    "- If a request is ambiguous, state your assumption and proceed; ask only when the risk is high (e.g. deleting data).",
    "- Answer in the user's language, be concise, and show results (values, formulas, addresses) rather than describing them vaguely.",
  ].join("\n");
}
