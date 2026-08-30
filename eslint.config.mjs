// ESLint 9 flat config. `office-addin-lint check` picks this up automatically
// when it exists at the project root (otherwise it falls back to its own
// bundled config, which is why the old .eslintrc.json was never applied).
import officeAddins from "eslint-plugin-office-addins";
import tsParser from "@typescript-eslint/parser";

export default [
  ...officeAddins.configs.recommended,
  {
    plugins: {
      "office-addins": officeAddins,
    },
    languageOptions: {
      parser: tsParser,
    },
  },
  {
    // Generated/vendored output and third-party code.
    ignores: ["dist/**", "tools/build-test/**", "node_modules/**"],
  },
  {
    // The office-addins sync rules do a shallow, single-pass scan: they cannot
    // see an `await ctx.sync()` that happens inside a loop, behind a helper, or
    // via `range.context.sync()`. Both files sync before every read (covered by
    // the integration tests), so the rules only produce false positives here.
    files: ["src/taskpane/llm/excelTools.ts", "src/taskpane/llm/context.ts"],
    rules: {
      "office-addins/call-sync-after-load": "off",
      "office-addins/call-sync-before-read": "off",
      "office-addins/no-navigational-load": "off",
    },
  },
];
