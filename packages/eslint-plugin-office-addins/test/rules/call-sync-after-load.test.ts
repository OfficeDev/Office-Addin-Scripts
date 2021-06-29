import { ESLintUtils } from '@typescript-eslint/experimental-utils'
import rule from '../../src/rules/call-sync-after-load';

const ruleTester = new ESLintUtils.RuleTester({
  parser: '@typescript-eslint/parser',
});

ruleTester.run('call-sync-after-load', rule, {
  valid: [ 
    {
      code: `
        var property = worksheet.getItem("sheet");
        property.load("values");
        await context.sync();
        console.log(property.values);`
    },
    {
      code: `
        var property = worksheet.getItem("sheet");
        property.load("values");
        await context.sync();
        console.log(property.values);
        property.load("length");
        console.log(property.length);`
    },
  ],
  invalid: [
    {
      code: `
        var property = worksheet.getItem("sheet");
        await context.sync();
        property.load("values");
        console.log(property.values);`,
      errors: [{ messageId: "callSyncAfterLoad", data: { name: "property", loadValue: "values" }}]
    },
  ]
});