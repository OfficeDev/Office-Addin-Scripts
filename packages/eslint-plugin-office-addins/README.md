# eslint-plugin-office-addins

eslint plugin for office-addins

## Installation

This plugin is designed to work with the office-addin-lint package.

Install `office-addin-lint`

``` shell
npm i office-addin-lint --save-dev
```

Next, install `eslint-plugin-office-addins`:

``` shell
npm install eslint-plugin-office-addins --save-dev

```

## Usage

Add `office-addins` to the plugins section of your `eslint.config.js` configuration file. You can omit the `eslint-plugin-` prefix:

```json
{
    "plugins": [
        "office-addins"
    ]
}
```

Then configure the default property by choosing one of the configurations for the plugin.

- configs
-- "recommended"
-- "react"
-- "reactnative"

``` js
// eslint.config.js
import officeAddins from "eslint-plugin-office-addins";

export default [
  ...officeAddins.configs.recommended,
  {
    // Additional project-specific configuration
  }
];
```

### Legacy eslintrc format (Eslint version < 9)

Then configure the extended property by choosing one of the configurations for the plugin.

- plugin
-- office-addins/recommended
-- office-addins/react
-- office-addins/reactnative

``` json
// .eslintrc
{
    "extended": [
        "plugin:office-addins/recommended"
    ]
}
```
