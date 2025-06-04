import rules from "./rules";
import eslintjs from "@eslint/js";
import eslintConfigPrettier from "eslint-config-prettier";
import prettierplugin from "eslint-plugin-prettier";
import tseslint from "typescript-eslint";

// eslint-disable-next-line @typescript-eslint/no-require-imports
const reactplugin = require("eslint-plugin-react");

// eslint-disable-next-line @typescript-eslint/no-require-imports
const reactnativeplugin = require("eslint-plugin-react-native");

const plugin = {
  meta: {
    name: "eslint-plugin-office-addins",
    version: "5.1.0",
  },
  rules,
  configs: {},
};

const recommended = tseslint.config(
  eslintjs.configs.recommended,
  eslintConfigPrettier,
  {
    files: ["**/*.{js,mjs,cjs,ts,cts,mts}"],
    plugins: {
      "@typescript-eslint": tseslint.plugin,
      "office-addins": plugin,
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
  }
);

const react = tseslint.config(
  ...reactplugin.configs.flat.recommended,
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
  }
);

const reactnative = tseslint.config(
  ...reactnativeplugin.configs.all,
  ...recommended,
  {
    plugins: {
      "office-addins": plugin,
      "react-native": reactnativeplugin,
    },
    settings: {
      react: {
        version: "detect",
      },
    },
  }
);

const test = tseslint.config(
  ...recommended,
  {
    plugins: {
      "office-addins": plugin,
    },
    rules: {},
  }
);

Object.assign(plugin.configs, {
  recommended,
  react,
  reactnative,
  test,
});

export default plugin;
