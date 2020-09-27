# Prevents office api read calls (no-office-read-calls)

This rule is designed to throw a linting warning on Office.js API read calls within custom functions. This is to prevent resource mismanagement and unhelpful loops.


## Rule Details

This rule aims to...

Examples of **incorrect** code for this rule:

```js
/**
 * Custom Function for Testing
 * @customfunction
 */
function myCustomFunction() {
    let context = new Excel.RequestContext();
    context.workbook.worksheets.getActiveWorksheet();
}
```

```js
/**
 * Custom Function for Testing
 * @customfunction
 */
function myCustomFunction() {
    Excel.run((context) => {
        var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
        var expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
        expensesTable = currentWorksheet.tables.add("A2:D2", true /*hasHeaders*/);
        return context.sync();
    });
}
```

Examples of **correct** code for this rule:

```js
/**
 * Custom Function for Testing
 * @customfunction
 */
function myCustomFunction() {
    Excel.createWorkbook(undefined);
}
```

```js
/**
 * Custom Function for Testing
 * @customfunction
 */
function myCustomFunction() {
    console.log("Hello World!");
}

function readOperations() {
    Excel.run(function (context) {
        var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
        var expensesTable = currentWorksheet.tables.getItemAt(0);
        var expenseValues = expensesTable.getHeaderRowRange().values;
        return context.sync();
    });
}
```