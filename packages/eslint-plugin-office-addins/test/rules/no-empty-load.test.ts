import { RuleTester } from '@typescript-eslint/rule-tester'
import rule from '../../src/rules/no-empty-load';

const ruleTester = new RuleTester({
  parser: '@typescript-eslint/parser',
});

ruleTester.run('no-empty-load', rule, {
  valid: [ 
    {
      code: `
        var sheetName = 'Sheet1';
        var rangeAddress = 'A1:B2';
        var myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);  
        myRange.load('address');
        context.sync()
          .then(function () {
            console.log (myRange.address);   // ok
          });`
    },
    {
      code: `
        var property = worksheet.getItem("sheet");
        property.load('G2');
        var variableName = property.G2;`
    },
    {
      code: `
        const notProxyObject = anotherObject.thisIsNotAGetFunction();
        notProxyObject.load();`
    },
    {
      code: `
        var selectedRange = context.workbook.getSelectedRange();
        selectedRange = "new variable";
        selectedRange.load()`
    },
    {
      code: `
        var selectedRange = context.workbook.getSelectedRange();
        selectedRange.load("*")`
    },
  ],
  invalid: [
    {
      code: `
        var selectedRange = context.workbook.getSelectedRange();
        selectedRange.load();`,
      errors: [{ messageId: "emptyLoad"}]
    },
    {
      code: `
        var selectedRange = context.workbook.getSelectedRange();
        selectedRange.load("");`,
      errors: [{ messageId: "emptyLoad"}]
    },
    {
      code: `
        var myRange;
        myRange = context.workbook.worksheets.getSelectedRange();
        myRange.load();
        console.log(myRange.values);`,
      errors: [{ messageId: "emptyLoad"}]
    },
    {
      code: `
        var myRange;
        myRange = context.workbook.worksheets.getSelectedRange();
        myRange.load(["address", "values", ""]);
        console.log(myRange.values);`,
      errors: [{ messageId: "emptyLoad"}]
    },
    {
      code: `
        var myRange;
        myRange = context.workbook.worksheets.getSelectedRange();
        myRange.load("address, values, ");
        console.log(myRange.values);`,
      errors: [{ messageId: "emptyLoad"}]
    },
  ]
});
