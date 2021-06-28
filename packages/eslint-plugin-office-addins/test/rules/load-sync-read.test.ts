import { ESLintUtils } from '@typescript-eslint/experimental-utils'
import rule from '../../src/rules/load-sync-read';

const ruleTester = new ESLintUtils.RuleTester({
  parser: '@typescript-eslint/parser',
});

ruleTester.run('load-sync-read', rule, {
  valid: [ 
    {
      code: `
        var property = worksheet.getItem("sheet");
        property.load("values");
        await context.sync();
        console.log(property.values);`
    },
  ],
  invalid: [
    {
      code: `
        var property = worksheet.getItem("sheet");
        await context.sync();
        property.load("values");
        console.log(property.values);`,
      errors: [{ messageId: "loadSyncRead", data: { name: "property" }}]
    },
  ]
});