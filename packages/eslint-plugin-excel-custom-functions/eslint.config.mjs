import tsParser from "@typescript-eslint/parser";
import globals from "globals";
import typescriptplugin from "@typescript-eslint/eslint-plugin";
import prettierplugin from "eslint-plugin-prettier";
import eslintjs from "@eslint/js";
import eslintts from "typescript-eslint";
import eslintConfigPrettier from "eslint-config-prettier";

export default [
  eslintjs.configs.recommended,
  eslintConfigPrettier,
  {
    files: ["**/*.{js,mjs,cjs,ts,cts,mts}"],
    ignores: ["**/node_modules/**", "**/lib/**"],
    plugins: {
      "@typescript-eslint": typescriptplugin,
      prettier: prettierplugin,
    },

    languageOptions: {
      parser: tsParser,
      ecmaVersion: 6,
      sourceType: "module",
      globals: { ...globals.browser, ...globals.node },

      parserOptions: {
        ecmaFeatures: {
          jsx: true,
        },
      },
    },

    rules: {
      "@typescript-eslint/no-unused-vars": "error",
      "no-delete-var": "warn",
      "no-eval": "error",
      "no-inner-declarations": "warn",
      "no-octal": "warn",
      "no-unused-vars": "off",

      "prettier/prettier": [
        "error",
        {
          endOfLine: "auto",
        },
      ],
    },
  },
];
