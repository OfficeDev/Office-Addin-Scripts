import rules from "./rules";
import tsParser from "@typescript-eslint/parser";
import typescriptplugin from "@typescript-eslint/eslint-plugin";
import prettierplugin from "eslint-plugin-prettier";
import reactplugin from "eslint-plugin-react";
import eslintjs from "@eslint/js";
import eslintConfigPrettier from "eslint-config-prettier";

// eslint-disable-next-line @typescript-eslint/no-require-imports
const reactnativeplugin = require("eslint-plugin-react-native");

const plugin = {
  meta: {
    name: "eslint-plugin-office-addins",
    version: "5.0.0",
  },
  rules,
  configs: {},
};

const recommended = [
  eslintjs.configs.recommended,
  eslintConfigPrettier,
  {
    files: ["**/*.{js,mjs,cjs,ts,cts,mts}"],
    plugins: {
      "@typescript-eslint": typescriptplugin,
      "office-addins": plugin,
      prettier: prettierplugin,
    },
    languageOptions: {
      parser: tsParser,
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
      "office-addins/call-sync-after-load": "error",
      "office-addins/call-sync-before-read": "error",
      "office-addins/load-object-before-read": "error",
      "office-addins/no-context-sync-in-loop": "warn",
      "office-addins/no-empty-load": "warn",
      "office-addins/no-navigational-load": "warn",
      "office-addins/no-office-initialize": "warn",
      "office-addins/test-for-null-using-isNullObject": "error",
      "prettier/prettier": ["error", { endOfLine: "auto" }],
    },
  },
];

const react = [
  reactplugin.configs.flat.recommended,
  ...recommended,
  {
    plugins: {
      "office-addins": plugin,
      react: reactplugin,
    },
    settings: {
      react: {
        version: "detect",
      },
    },
  },
];

const reactnative = [
  reactnativeplugin.configs.all,
  ...recommended,
  {
    plugins: {
      "office-addins": plugin,
      react: reactnativeplugin,
    },
    settings: {
      react: {
        version: "detect",
      },
    },
  },
];

const test = [
  ...recommended,
  {
    plugins: {
      "office-addins": plugin,
    },
    rules: {},
  },
];

Object.assign(plugin.configs, {
  recommended,
  react,
  reactnative,
  test,
});

module.exports = plugin;
