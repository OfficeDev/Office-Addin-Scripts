import { ESLintUtils } from '@typescript-eslint/experimental-utils'
import rule from '../../src/rules/call-sync-before-read';

const ruleTester = new ESLintUtils.RuleTester({
  parser: '@typescript-eslint/parser',
});

ruleTester.run('call-sync-before-read', rule, {
  valid: [ 
    {
      code: `
      Excel.run(function (context) {
        var selectedRange = context.workbook.getSelectedRange();
        selectedRange.load('address');
        await context.sync()
        console.log('The selected range is: ' + selectedRange.address);
      });`
    },
    {
      code: `
      Excel.run(function (context) {
        var selectedRange;
        selectedRange = context.workbook.getSelectedRange();
        selectedRange.load('address');
        await context.sync()
        console.log('The selected range is: ' + selectedRange.address);
      });`
    },
    {
      code: `
      Excel.run(function (context) {
        var selectedRange = context.workbook.getSelectedRange();
        selectedRange.load('address');
        return context.sync()
          .then(function () {
            console.log('The selected range is: ' + selectedRange.address);
        });
      })`
    },
    {
      code: `
      Excel.run(function (context) {
        var dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");
        await context.sync();
        if (dataSheet.isNullObject) {
          dataSheet.position = 1;
        }
      })`
    },
    {
      code: `
      Excel.run(function (context) {
        var dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");
        if (true) {
          await context.sync();
          dataSheet.position = 1;
        }
      })`
    },
  ],
  invalid: [
    {
      code: `
      Excel.run(function (context) {
        var selectedRange = context.workbook.getSelectedRange();
        selectedRange.load('address');
        console.log('The selected range is: ' + selectedRange.address);
      });`,
      errors: [{ messageId: "callSync", data: { name: "selectedRange" } }]
    },
    {
      code: `
      Excel.run(function (context) {
        var selectedRange = context.workbook.getSelectedRange();
        selectedRange.load('address');
        console.log('The selected range is: ' + selectedRange.address);
        await context.sync();
      });`,
      errors: [{ messageId: "callSync", data: { name: "selectedRange" } }]
    },
    {
      code: `
      Excel.run(function (context) {
        var selectedRange = context.workbook.getSelectedRange();
        var selectedRange2 = context.workbook.getSelectedRange();
        selectedRange.load('address');
        selectedRange2.load('address');
        console.log('The selected range is: ' + selectedRange.address);
        console.log('This should be the same: ' + selectedRange2.address);
      });`,
      errors: [
        { messageId: "callSync", data: { name: "selectedRange" } }, 
        { messageId: "callSync", data: { name: "selectedRange2" } }
      ]
    },
    {
      code: `
      Excel.run(function (context) {
        var dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");
        if (dataSheet.isNullObject) {
          dataSheet.position = 1;
        }
      });`,
      errors: [
        { messageId: "callSync", data: { name: "dataSheet" } },
        { messageId: "callSync", data: { name: "dataSheet" } },
      ]
    },
    {
      code: `
      Excel.run(function (context) {
        var dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");
        if (true) {
          dataSheet.position = 1;
        }
        await context.sync();
      });`,
      errors: [ { messageId: "callSync", data: { name: "dataSheet" } } ]
    },
  ]
});