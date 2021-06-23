import { ESLintUtils } from '@typescript-eslint/experimental-utils'
import rule from '../../src/rules/no-empty-load';

const ruleTester = new ESLintUtils.RuleTester({
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
  ],
  invalid: [
    {
      code: `
        var selectedRange = context.workbook.getSelectedRange();
        selectedRange.load()`,
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
        context.workbook.worksheets.getSelectedRange().load()`,
      errors: [{ messageId: "emptyLoad"}]
    },
  ]
});