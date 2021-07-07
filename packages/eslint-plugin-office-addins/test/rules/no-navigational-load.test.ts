import { ESLintUtils } from '@typescript-eslint/experimental-utils'
import rule from '../../src/rules/no-navigational-load';

const ruleTester = new ESLintUtils.RuleTester({
  parser: '@typescript-eslint/parser',
});

ruleTester.run('no-navigational-load', rule, {
  valid: [ 
    {
      code: `
        var range = worksheet.getRange("A1");
        range.load("format/fill/size");
        console.log(range.format.fill.size);`
    },
    {
      code: `
        `
    },
    {
      code: ``
    },
    {
      code: ``
    },
  ],
  invalid: [
    {
      code: ``,
      errors: [{ messageId: "navigationalLoad"}]
    },
    {
      code: ``,
      errors: [{ messageId: "navigationalLoad"}]
    },
  ]
});