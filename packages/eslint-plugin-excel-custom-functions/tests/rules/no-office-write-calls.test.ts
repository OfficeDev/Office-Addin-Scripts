import { TSESLint, ESLintUtils } from '@typescript-eslint/utils';
import rule from '../../src/rules/no-office-write-calls';
import * as path from 'path';

type Options = unknown[];
type MessageIds = 'officeWriteCall';

const ruleTester = new ESLintUtils.RuleTester({
  parser: '@typescript-eslint/parser',
  parserOptions: {
    ecmaVersion: 2018,
    sourceType: 'module',
    tsconfigRootDir: path.resolve(__dirname, '..'),
    project: './tsconfig.test.json', // relative to tsconfigRootDir
  },
});

ruleTester.run('no-office-write-calls', rule, {
  valid: [
    // Multi-file scenarios are not supported
    getValidTestCase( `
    import { writeOperations } from '../fixtures/secondFile';

    /**
     * Custom Function for Testing
     * @customfunction
     */
    function myCustomFunction() {
      writeOperations();
    }
    `),

    // Should not throw a write error on read operations
    getValidTestCase( `
    /**
     * Custom Function for Testing
     * @customfunction
     */
    function myCustomFunction() {
      let context = new Excel.RequestContext();
      context.workbook.worksheets.getActiveWorksheet();
    }
    `),

    // Should not throw a write error when Office calls aren't used in a custom function
    getValidTestCase( `
    /**
     * Custom Function for Testing
     * @customfunction
     */
    function myCustomFunction() {
      console.log("Hello World!");
    }

    function writeOperations(context: Excel.RequestContext) {
      var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      var expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
      expensesTable.name = "ExpensesTable";
      return context.sync();
    }
    `),
  ],
  // Error cases. `// ERROR: x` marks the spot where the error occurs.
  invalid: [
    // Testing write operations erroring out with a test sample that has no read operations
    getInvalidTestCase( `
    /**
     * Custom Function for Testing
     * @customfunction
     */
    function myCustomFunction() {
      Excel.createWorkbook(undefined);                                                      //ERROR: Excel.createWorkbook(undefined)
    }
    `),
    // Testing passing in a context object
    getInvalidTestCase( `
    /**
     * Custom Function for Testing
     * @customfunction
     */
    function myCustomFunction() {
      let context = new Excel.RequestContext();
      writeOperations(context);                                                             //ERROR: writeOperations                  
    }

    function writeOperations(context: Excel.RequestContext) {
      var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      var expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
      expensesTable.name = "ExpensesTable";
      return context.sync();
    }
      `),

    // Functions that pass in an unused context object should be ok
    //This throws an error
    getInvalidTestCase( `
    /**
     * Custom Function for Testing
     * @customfunction
     */
    function myCustomFunction() {
      helper1();                                                                            //ERROR: helper1
    }
    
    function helper2() {
      helper3();
    }
    
    function writeOperations() {
      Excel.run((context) => {
        var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
        var expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
        expensesTable.name = "ExpensesTable";
        return context.sync();
      });
    }
    
    function helper1() {
      helper2();
    }
    
    function helper3() {
      writeOperations();
    }
      `),

    // Functions that pass in an unused context object should be ok
    //This throws an error
    getInvalidTestCase( `
    /**
     * Custom Function for Testing
     * @customfunction
     */
    function myCustomFunction() {
      writeOperations("helloWorld");                                                          //ERROR: writeOperations
    }
    
    function writeOperations(text: string) {
      console.log(text);
      let context = new Excel.RequestContext()
      context.sync()
    }
      `),

    // Functions that pass in an unused context object should be ok
    //This throws an error
    getInvalidTestCase( `
    /**
     * Custom Function for Testing
     * @customfunction
     */
    function myCustomFunction() {
      let context = new Excel.RequestContext();
      context.workbook.worksheets.getActiveWorksheet();
      writeOperations(context);                                                              //ERROR: writeOperations
    }

    function writeOperations(context: Excel.RequestContext) {
      console.log("hello world!");
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
    const errorInfo = /ERROR: (\w+)/.exec(line);
    if (errorInfo) {
      errors.push({
        line: i + 1,
        column: line.indexOf(errorInfo[1]) + 1,
        messageId: 'officeWriteCall',
      });
    }
  });

  if (!errors.length) {
    throw new Error(
      'No ERROR: indications found in supposedly invalid code:\n' + code,
    );
  }

  return {
    code,
    errors,
    filename: 'fixtures/file.ts',
  };
}
