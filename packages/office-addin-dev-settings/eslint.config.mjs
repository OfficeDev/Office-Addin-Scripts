import officeAddins from "eslint-plugin-office-addins";
import tsParser from "@typescript-eslint/parser";
import sdl from "@microsoft/eslint-plugin-sdl";

export default [
  ...officeAddins.configs.recommended,
  ...sdl.configs.recommended,
  {
    plugins: {
      "office-addins": officeAddins,
    },
    languageOptions: {
      parser: tsParser,
      parserOptions: {
        projectService: true,
        tsconfigRootDir: "__dirname"
      }
    },
  },
];
