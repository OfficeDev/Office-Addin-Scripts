# eslint-plugin-office-addins

eslint plugin for office-addins

## Installation

This plugin is designed to work with the office-addin-lint package. 

Install `office-addin-lint`

```
$ npm i office-addin-lint --save-dev
```

Next, install `eslint-plugin-office-addins`:

```
$ npm install eslint-plugin-office-addins --save-dev
```

## Usage

Add `office-addins` to the plugins section of your `.eslintrc` configuration file. You can omit the `eslint-plugin-` prefix:

```json
{
    "plugins": [
        "office-addins"
    ]
}
```

Then configure the extended property by choosing one of the configurations for the plugin. 
```json
{
    "extended": [
        "plugin:office-addins/recommended"
    ]
}

Other configurations available:
    "plugin:office-addins/react",
    "plugin:office-addins/reactnative"
```
