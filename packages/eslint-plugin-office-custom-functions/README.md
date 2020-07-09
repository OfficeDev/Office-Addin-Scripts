# eslint-plugin-office-custom-functions

eslint plugin for custom functions

## Installation

You'll first need to install [ESLint](http://eslint.org):

```
$ npm i eslint --save-dev
```

Next, install `eslint-plugin-office-custom-functions`:

```
$ npm install eslint-plugin-office-custom-functions --save-dev
```

**Note:** If you installed ESLint globally (using the `-g` flag) then you must also install `eslint-plugin-office-custom-functions` globally.

## Usage

Add `office-custom-functions` to the plugins section of your `.eslintrc` configuration file. You can omit the `eslint-plugin-` prefix:

```json
{
    "plugins": [
        "office-custom-functions"
    ]
}
```


Then configure the rules you want to use under the rules section.

```json
{
    "rules": {
        "office-custom-functions/rule-name": 2
    }
}
```

## Supported Rules

* Fill in provided rules here





