import { ESLintUtils } from '@typescript-eslint/experimental-utils'
import rule from '../../src/rules/load-object-before-read';

const ruleTester = new ESLintUtils.RuleTester({
  parser: '@typescript-eslint/parser',
});

ruleTester.run('load-object-before-read', rule, {
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
        var selectedRange = context.workbook.getSelectedRange();
        selectedRange.load('values');
        if(selectedRange.values === [2]){}`
    }
  ],
  invalid: [
    {
      code: `
        var sheetName = 'Sheet1';
        var rangeAddress = 'A1:B2';
        var myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);  
        myRange.load('address');
        context.sync()
          .then(function () {
            console.log (myRange.values);  // not ok as it was not loaded
          });`,
      errors: [{ messageId: "loadBeforeRead" }]
    },
    {
      code: `
        var selectedRange = context.workbook.getSelectedRange();
        if(selectedRange.values === ["sampleText"]){}`,
      errors: [{ messageId: "loadBeforeRead" }]
    },
    {
      code: `
        var myRange = context.workbook.worksheets.getItem("sheet").getRange("A1");
        console.log (myRange.adress);`,
      errors: [{ messageId: "loadBeforeRead" }]
    },
    {
      code: `
        var selectedRange = context.workbook.getSelectedRange();
        console.log(selectedRange.values);`,
      errors: [{ messageId: "loadBeforeRead" }]
    },
    {
      code: `
        var selectedRange = context.workbook.getSelectedRange();
        var test = selectedRange.values;`,
      errors: [{ messageId: "loadBeforeRead" }]
    }
  ]
});