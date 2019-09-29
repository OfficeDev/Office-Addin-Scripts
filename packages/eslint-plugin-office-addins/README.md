# eslint-plugin-office-addins

eslint plugin for office-addins

## Installation

You'll first need to install [ESLint](http://eslint.org):

```
$ npm i eslint --save-dev
```

Next, install `eslint-plugin-office-addins`:

```
$ npm install eslint-plugin-office-addins --save-dev
```

**Note:** If you installed ESLint globally (using the `-g` flag) then you must also install `eslint-plugin-office-addins` globally.

## Usage

Add `office-addins` to the plugins section of your `.eslintrc` configuration file. You can omit the `eslint-plugin-` prefix:

```json
{
    "plugins": [
        "office-addins"
    ]
}
```

Then configure the extends property.
```json
{
    "extended": [
        "plugin:office-addins/recommended",
        "plugin:office-addins/react",
        "plugin:office-addins/reactnative"
    ]
}
```

Then configure the rules you want to use under the rules section.

```json
{
    "rules": {
        "office-addins/rule-name": 2
    }
}
```
