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
    {
      code: `
      Excel.run(function (context) {
        var dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");
        if (true) {
          dataSheet.position = 1;  // Write is OK
        }
        await context.sync();
      });`
    },
    {
      code: `
      var range = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
      range.format.fill.color = "yellow";`
    },
    {
      code: `
      var range = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
      range.format.font.load('color');`
    },
    {
      code: `
        var table = worksheet.getTables();
        return context.sync().then(function () {
          table.delete();
        });`
    },
    {
      code: `
        var range = worksheet.getSelectedRange();
        range.getCell(0,0);`
    },
    {
      code: `
        var range = worksheet.getSelectedRange();
        range.load("font");
        context.sync();
        range.font.getColor();`
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
      ]
    },
    {
      code: `
      Excel.run(function (context) {
        var dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");
        if (true) {
          var pos = dataSheet.position;
        }
        await context.sync();
      });`,
      errors: [ { messageId: "callSync", data: { name: "dataSheet" } } ]
    },
    {
      code: `
      var range = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
      range.format.fill.color == "yellow";`,
      errors: [ { messageId: "callSync", data: { name: "range" } } ]
    },
    {
      code: `
      var range = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
      var value = range.format.fill.color;
      value = range.format.fill.color;`,
      errors: [ 
        { messageId: "callSync", data: { name: "range" } },
        { messageId: "callSync", data: { name: "range" } } 
      ]
    },
  ]
});
