module.exports = {
    rules: {
        'no-context-sync-in-loop': require('./rules/no-context-sync-in-loop'),
    },
    configs: {
        recommended: {
            parser: "@typescript-eslint/parser",
            plugins: [
                '@typescript-eslint',
                'office-addins',
                'prettier',
            ],
            parserOptions: {
                ecmaVersion: 6,
                sourceType: "module",
                ecmaFeatures: {
                    jsx: true
                },
                project: "./tsconfig.json"
            },
            extends: ['eslint:recommended'],
            rules: {
                'prettier/prettier': ['error', { 'endOfLine': 'auto' }],
                'no-eval': 'error',
                'no-delete-var': 'warn',
                'no-octal': 'warn',
                'no-inner-declarations': 'warn',
            }
        },
        react: {
            extends: [
                'plugin:office-addins/recommended',
                'plugin:react/recommended',
            ]
        },
        reactnative: {
            extends: [
                'plugin:office-addins/recommended',
                'plugin:react-native/all',
            ],
        }
    }
  };
  