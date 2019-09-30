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
