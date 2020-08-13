# eslint-plugin-excel-custom-functions

eslint plugin to check against Office.js api calls within the shared app

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

Add `excel-custom-functions` to the plugins section of your `.eslintrc` configuration file. You can omit the `eslint-plugin-` prefix:

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
        "excel-custom-functions/rule-name": 2
    }
}
```

## Supported Rules

* Fill in provided rules here





