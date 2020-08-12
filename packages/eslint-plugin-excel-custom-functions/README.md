# eslint-plugin-office-shared-app

eslint plugin to check against Office.js api calls within the shared app

## Installation

You'll first need to install [ESLint](http://eslint.org):

```
$ npm i eslint --save-dev
```

Next, install `eslint-plugin-office-shared-app`:

```
$ npm install eslint-plugin-office-shared-app --save-dev
```

**Note:** If you installed ESLint globally (using the `-g` flag) then you must also install `eslint-plugin-office-shared-app` globally.

## Usage

Add `office-shared-app` to the plugins section of your `.eslintrc` configuration file. You can omit the `eslint-plugin-` prefix:

```json
{
    "plugins": [
        "office-shared-app"
    ]
}
```


Then configure the rules you want to use under the rules section.

```json
{
    "rules": {
        "office-shared-app/rule-name": 2
    }
}
```

## Supported Rules

* Fill in provided rules here





