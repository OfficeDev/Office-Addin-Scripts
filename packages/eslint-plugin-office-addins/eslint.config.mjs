import eslintjs from "@eslint/js";
import tseslint from "typescript-eslint";
import prettierplugin from "eslint-plugin-prettier";
import eslintConfigPrettier from "eslint-config-prettier";

export default [
  eslintjs.configs.recommended,
  ...tseslint.configs.recommended,
  eslintConfigPrettier,
  {
    files: ["**/*.{js,mjs,cjs,ts,cts,mts}"],
    ignores: ["**/node_modules/**", "**/lib/**"],
    plugins: {
      "@typescript-eslint": tseslint.plugin,
      prettier: prettierplugin,
    },

    languageOptions: {
      parser: tseslint.parser,
      ecmaVersion: 6,
      sourceType: "module",
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
