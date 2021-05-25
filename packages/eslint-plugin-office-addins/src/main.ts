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
      extends: ["eslint:recommended"],
      rules: {
        "prettier/prettier": ["error", { endOfLine: "auto" }],
        "no-eval": "error",
        "no-delete-var": "warn",
        "no-octal": "warn",
        "no-inner-declarations": "warn",
      },
    },
    react: {
      plugins: ["office-addins", "react"],
      extends: ["plugin:office-addins/recommended", "plugin:react/recommended"],
      settings: {
        react: {
          "version": "detect"
        },
      },
    },
    reactnative: {
      plugins: ["office-addins", "react-native"],
      extends: ["plugin:office-addins/recommended", "plugin:react-native/all"],
      settings: {
        react: {
          "version": "detect"
        },
      },
    },
    test: {
      plugins: ["office-addins"],
      extends: ["plugin:office-addins/recommended"],
      rules: {
        "office-addins/no-context-sync-in-loop": "warn",
      },
    },
  },
};
