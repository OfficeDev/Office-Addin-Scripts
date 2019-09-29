module.exports = {
    configs: {
        recommended: {
            plugins: ['office-addins'],
            parserOptions: {
                ecmaFeatures: {
                    jsx: true
                  }
            },
            extends: ['eslint:recommended'],
            rules: {}
        },
        react: {
            plugins: ['office-addins'],
            parserOptions: {
                ecmaFeatures: {
                    jsx: true
                  }
            },
            extends: ['plugin:react/recommended'],
            rules: {}
        },
        reactnative: {
            plugins: ['office-addins'],
            parserOptions: {
                ecmaFeatures: {
                    jsx: true
                  }
            },
            extends: ['plugin:react-native/all'],
            rules: {}
        }
    }
  };