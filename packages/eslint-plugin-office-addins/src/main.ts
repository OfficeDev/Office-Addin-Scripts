/* global module */

import rules from "./rules";

module.exports = {
  rules,
  configs: {
    recommended: {
      parser: "@typescript-eslint/parser",
      plugins: ["@typescript-eslint", "office-addins", "prettier"],
      parserOptions: {
        ecmaVersion: 6,
        sourceType: "module",
        ecmaFeatures: {
          jsx: true,
        },
        project: "./tsconfig.json",
      },
      extends: ["eslint:recommended", "prettier"],
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
    react: {
      plugins: ["office-addins", "react"],
      extends: ["plugin:office-addins/recommended", "plugin:react/recommended"],
      settings: {
        react: {
          version: "detect",
        },
      },
    },
    reactnative: {
      plugins: ["office-addins", "react-native"],
      extends: ["plugin:office-addins/recommended", "plugin:react-native/all"],
      settings: {
        react: {
          version: "detect",
        },
      },
    },
    test: {
      plugins: ["office-addins"],
      extends: ["plugin:office-addins/recommended"],
      rules: {},
    },
  },
};
