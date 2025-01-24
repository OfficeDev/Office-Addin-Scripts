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
    rules: {
      "office-addins/call-sync-before-read": "off",
      "office-addins/load-object-before-read": "off",
    },
  },
];
