# eslint-plugin-excel-custom-functions

This eslint plugin checks for Office.js api calls within custom functions. It throws a linting error on write operations and a linting warning on read operations.

## Installation

You'll first need to install [ESLint](http://eslint.org):

```
$ npm i eslint --save-dev
```

Next, install `eslint-plugin-excel-custom-functions`:

```
$ npm install eslint-plugin-excel-custom-functions --save-dev
```

**Note:** If you installed ESLint globally (using the `-g` flag) then you must also install `eslint-plugin-excel-custom-functions` globally.

## Usage

Add `excel-custom-functions` to the plugins section of your `eslint.config.js` configuration file. You can omit the `eslint-plugin-` prefix:

```json
{
    "plugins": [
        "excel-custom-functions"
    ]
}
```


Then configure the rules you want to use under the rules section.

```json
{
    "rules": {
        "excel-custom-functions/no-office-read-calls": "warn",
        "excel-custom-functions/no-office-write-calls": "error"
    }
}
```

## Supported Rules

* no-office-read-calls
* no-office-write-calls





