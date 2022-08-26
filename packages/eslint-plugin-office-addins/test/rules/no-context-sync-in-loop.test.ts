import { ESLintUtils } from '@typescript-eslint/utils'
import rule from '../../src/rules/no-context-sync-in-loop';

const ruleTester = new ESLintUtils.RuleTester({
  parser: '@typescript-eslint/parser',
});

ruleTester.run('no-context-sync-in-loop', rule, {
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
      errors: [{ messageId: "loopedSync" }]
    },
    {
      code: `Excel.run(async (context) => { 
          var person = { fname:\"John\", lname:\"Doe\", age:25 }; 
          var x; 
          for(x in person) { context.sync(); } 
        });`,
      errors: [{ messageId: "loopedSync" }]
    },
    {
      code: `PowerPoint.run(async (context) => { 
          var cars = ['BMW', 'Volvo', 'Mini']; 
          var x; 
          for(x of cars) { context.sync(); } 
        });`,
      errors: [{ messageId: "loopedSync" }]
    },
    {
      code: `Word.run(async (context) => { 
          while(true) { context.sync() } 
        });`,
      errors: [{ messageId: "loopedSync" }]
    },
    {
      code: `Excel.run(async (context) => { 
          do { context.sync() } while(true); 
        });`,
      errors: [{ messageId: "loopedSync" }]
    }
  ]
});
