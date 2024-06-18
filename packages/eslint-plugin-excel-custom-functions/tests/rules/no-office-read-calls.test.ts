import { TSESLint } from '@typescript-eslint/utils';
import { RuleTester } from '@typescript-eslint/rule-tester';
import rule from '../../src/rules/no-office-read-calls';
import * as path from 'path';

type Options = unknown[];
type MessageIds = 'officeReadCall';

const ruleTester = new RuleTester({
  parser: '@typescript-eslint/parser',
  parserOptions: {
    ecmaVersion: 2018,
    sourceType: 'module',
    tsconfigRootDir: path.resolve(__dirname, '..'),
    project: './tsconfig.test.json', // relative to tsconfigRootDir
  },
});

ruleTester.run('no-office-read-calls', rule, {
  valid: [
    // Multi-file scenarios are not supported
    getValidTestCase( `
    import { readOperations } from '../fixtures/secondFile';

    /**
     * Custom Function for Testing
     * @customfunction
     */
    function myCustomFunction() {
      readOperations();
    }
    `),

    // Should not throw a read error on write operations
    getValidTestCase( `
    /**
     * Custom Function for Testing
     * @customfunction
     */
    function myCustomFunction() {
      Excel.createWorkbook(undefined);
    }
    `),

    // Should not throw a read error when Office calls aren't used in a custom function
    getValidTestCase( `
    /**
     * Custom Function for Testing
     * @customfunction
     */
    function myCustomFunction() {
      console.log("Hello World!");
    }

    function readOperations() {
      Excel.run(function (context) {
          var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
          var expensesTable = currentWorksheet.tables.getItemAt(0);
          var expenseValues = expensesTable.getHeaderRowRange().values;
          return context.sync();
      });
    }
    `),
  ],
  // Warning cases. `// WARN: x` marks the spot where the error occurs.
  invalid: [
    // Testing passing in a context object
    getInvalidTestCase( `
    /**
     * Custom Function for Testing
     * @customfunction
     */
    function myCustomFunction() {
      let context = new Excel.RequestContext();                                            //WARN: context = new Excel.RequestContext()
      readOperations(context);                                                             //WARN: readOperations                  
    }

    function readOperations(context: Excel.RequestContext) {
      var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      var expensesTable = currentWorksheet.tables.getItemAt(0);
      var expenseValues = expensesTable.getHeaderRowRange().values;
      return context.sync();
    }
    `),

    // Testing helper functions
    getInvalidTestCase( `
    /**
     * Custom Function for Testing
     * @customfunction
     */
    function myCustomFunction() {
      helper1();                                                                            //WARN: helper1
    }
    
    function helper2() {
      helper3();
    }

    function readOperations() {
      Excel.run(function (context) {
          var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
          var expensesTable = currentWorksheet.tables.getItemAt(0);
          var expenseValues = expensesTable.getHeaderRowRange().values;
          return context.sync();
      });
    }
    
    function helper1() {
      helper2();
    }
    
    function helper3() {
      readOperations();
    }
    `),

    // Testing helper functions in different order
    getInvalidTestCase( `
    function helper2() {
      helper3();
    }

    function readOperations() {
      Excel.run(function (context) {
          var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
          var expensesTable = currentWorksheet.tables.getItemAt(0);
          var expenseValues = expensesTable.getHeaderRowRange().values;
          return context.sync();
      });
    }
    
    function helper1() {
      helper2();
    }

    /**
     * Custom Function for Testing
     * @customfunction
     */
    function myCustomFunction() {
      helper1();                                                                            //WARN: helper1
    }
    
    function helper3() {
      readOperations();
    }
    `),

    // testing creating context object in helper func
    getInvalidTestCase( `
    /**
     * Custom Function for Testing
     * @customfunction
     */
    function myCustomFunction() {
      readOperations("helloWorld");                                                          //WARN: readOperations
    }
    
    function readOperations(text: string) {
      console.log(text);
      let context = new Excel.RequestContext()
      context.sync()
    }
      `),
  ]
});

function getValidTestCase(code: string): TSESLint.ValidTestCase<Options> {
  return {
    code,
    filename: 'fixtures/file.ts',
  };
}

/**
 * Instead of hardcoding the line and column numbers of errors, calculate them
 * based on the position of "ERROR: someName" markers in the code.
 */
function getInvalidTestCase(
  code: string,
): TSESLint.InvalidTestCase<MessageIds, Options> {
  const lines = code.split(/\r?\n/g);
  const errors = [] as TSESLint.TestCaseError<MessageIds>[];

  lines.forEach((line, i) => {
    const errorInfo = /WARN: (\w+)/.exec(line);
    if (errorInfo) {
      errors.push({
        line: i + 1,
        column: line.indexOf(errorInfo[1]) + 1,
        messageId: 'officeReadCall',
      });
    }
  });

  if (!errors.length) {
    throw new Error(
      'No WARN: indications found in supposedly invalid code:\n' + code,
    );
  }

  return {
    code,
    errors,
    filename: 'fixtures/file.ts',
  };
}
