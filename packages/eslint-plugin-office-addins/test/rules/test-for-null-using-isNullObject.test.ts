import { RuleTester, TestCaseError } from '@typescript-eslint/rule-tester'
import rule from '../../src/rules/test-for-null-using-isNullObject';

const ruleTester = new RuleTester({
  parser: '@typescript-eslint/parser',
});

const errors: TestCaseError<"useIsNullObject">[] = [{ messageId: "useIsNullObject", data: { name: "dataSheet" } }];

ruleTester.run('test-for-null-using-isNullObject', rule, {
  valid: [ 
    {
      code: `
      await Excel.run(async (context) =>  {
        var dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");
        return context.sync().then(function () {
            if (dataSheet.isNullObject) {
              dataSheet = context.workbook.worksheets.add("Data");
            }
            dataSheet.position = 1;
         });
      });`,
    },
    {
      code: `
      var dataSheet;
      dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");
      return context.sync().then(function () {
        if (dataSheet.isNullObject) {
          dataSheet = context.workbook.worksheets.add("Data");
        }
        dataSheet.position = 1;
      });`,
    },
    {
      code: `
      var dataSheet;
      dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");
      return context.sync().then(function () {
        if (dataSheet.isNullObject) {
          dataSheet = context.workbook.worksheets.add("Data");
        }
        dataSheet.position = 1;
      });`,
    },
    {
      code: `
      var dataSheet;
      dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");
      if (dataSheet.isNullObject) {
        dataSheet = context.workbook.worksheets.add("Data");
      }`,
    },
  ],
  invalid: [
    {
      code: `
      await Excel.run(async (context) =>  {
        var dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");
        return context.sync().then(function () {
          if (dataSheet) {
            dataSheet = context.workbook.worksheets.add("Data");
          }
          dataSheet.position = 1;
        });
      });`,
      errors,
      output: `
      await Excel.run(async (context) =>  {
        var dataSheet = context.workbook.worksheets.getItemOrNullObject(\"Data\");
        return context.sync().then(function () {
          if (dataSheet.isNullObject) {
            dataSheet = context.workbook.worksheets.add(\"Data\");
          }
          dataSheet.position = 1;
        });
      });`,
    },
    {
      code: `
      var dataSheet;
      dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");
      return context.sync().then(function () {
        if (dataSheet) {
          dataSheet = context.workbook.worksheets.add("Data");
        }
        dataSheet.position = 1;
      });`,
      errors,
      output: `
      var dataSheet;
      dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");
      return context.sync().then(function () {
        if (dataSheet.isNullObject) {
          dataSheet = context.workbook.worksheets.add("Data");
        }
        dataSheet.position = 1;
      });`,
    },
    {
      code: `
      var dataSheet;
      dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");
      if (dataSheet) {
        dataSheet = context.workbook.worksheets.add("Data");
      }`,
      errors,
      output: `
      var dataSheet;
      dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");
      if (dataSheet.isNullObject) {
        dataSheet = context.workbook.worksheets.add("Data");
      }`,
    },
    {
      code: `
      var dataSheet;
      dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");
      return context.sync().then(function () {
        if (!dataSheet) {
          dataSheet = context.workbook.worksheets.add("Data");
        }
        dataSheet.position = 1;
      });`,
      errors,
      output: `
      var dataSheet;
      dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");
      return context.sync().then(function () {
        if (!dataSheet.isNullObject) {
          dataSheet = context.workbook.worksheets.add("Data");
        }
        dataSheet.position = 1;
      });`,
    },
    {
      code: `
      var dataSheet;
      dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");
      return context.sync().then(function () {
        if (true && dataSheet) {
          dataSheet = context.workbook.worksheets.add("Data");
        }
        dataSheet.position = 1;
      });`,
      errors,
      output: `
      var dataSheet;
      dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");
      return context.sync().then(function () {
        if (true && dataSheet.isNullObject) {
          dataSheet = context.workbook.worksheets.add("Data");
        }
        dataSheet.position = 1;
      });`,
    },
    {
      code: `
      var dataSheet;
      dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");
      return context.sync().then(function () {
        if (null != dataSheet) {
          dataSheet = context.workbook.worksheets.add("Data");
        }
        dataSheet.position = 1;
      });`,
      errors,
      output: `
      var dataSheet;
      dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");
      return context.sync().then(function () {
        if (dataSheet.isNullObject) {
          dataSheet = context.workbook.worksheets.add("Data");
        }
        dataSheet.position = 1;
      });`,
    },
    {
      code: `
      var dataSheet;
      dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");
      return context.sync().then(function () {
        for (var i = 0; i < 5; i++) {
          if (dataSheet) {
            dataSheet = context.workbook.worksheets.add("Data");
          }
        }
        dataSheet.position = 1;
      });`,
      errors,
      output: `
      var dataSheet;
      dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");
      return context.sync().then(function () {
        for (var i = 0; i < 5; i++) {
          if (dataSheet.isNullObject) {
            dataSheet = context.workbook.worksheets.add("Data");
          }
        }
        dataSheet.position = 1;
      });`,
    },
    {
      code: `
      var dataSheet;
      dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");
      return context.sync().then(function () {
        do {
          console.log("test case");
        } while (dataSheet);
        dataSheet.position = 1;
      });`,
      errors,
      output: `
      var dataSheet;
      dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");
      return context.sync().then(function () {
        do {
          console.log("test case");
        } while (dataSheet.isNullObject);
        dataSheet.position = 1;
      });`,
    },
  ]
});
