import { ESLintUtils } from '@typescript-eslint/experimental-utils'
import rule from '../../src/rules/load-object-before-read';

const ruleTester = new ESLintUtils.RuleTester({
  parser: '@typescript-eslint/parser',
});

ruleTester.run('load-object-before-read', rule, {
  valid: [ 
    {
      code: "context.sync()"
    },
    {
      code: `Excel.run(async (context) => { 
        context.sync(); 
      });`
    }
  ],
  invalid: [
    {
      code: `Word.run(async (context) => { 
          for(i = 0; i < 5; i++) { context.sync(); } 
        });`,
      errors: [{ messageId: "loadBeforeRead" }]
    },
    {
      code: `Excel.run(async (context) => { 
          var person = { fname:\"John\", lname:\"Doe\", age:25 }; 
          var x; 
          for(x in person) { context.sync(); } 
        });`,
      errors: [{ messageId: "loadBeforeRead" }]
    },
    {
      code: `PowerPoint.run(async (context) => { 
          var cars = ['BMW', 'Volvo', 'Mini']; 
          var x; 
          for(x of cars) { context.sync(); } 
        });`,
      errors: [{ messageId: "loadBeforeRead" }]
    },
    {
      code: `Word.run(async (context) => { 
          while(true) { context.sync() } 
        });`,
      errors: [{ messageId: "loadBeforeRead" }]
    },
    {
      code: `Excel.run(async (context) => { 
          do { context.sync() } while(true); 
        });`,
      errors: [{ messageId: "loadBeforeRead" }]
    }
  ]
});