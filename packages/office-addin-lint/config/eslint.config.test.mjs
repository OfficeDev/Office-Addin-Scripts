import officeAddins from "eslint-plugin-office-addins";
import tsParser from "@typescript-eslint/parser";

export default [
  ...officeAddins.configs.test,
  {
    plugins: {
      "office-addins": officeAddins,
    },
    languageOptions: {
      parser: tsParser,
    },
  },
];