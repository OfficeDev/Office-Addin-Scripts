import { TSESLint, ESLintUtils } from '@typescript-eslint/experimental-utils';
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
  // Don't warn at the spot where the deprecated thing is declared
  valid: [
    // Variables (var/const/let are the same from ESTree perspective)
    getValidTestCase( `
      /**
       * Displays the current time once a second.
       * @customfunction
       * @param invocation Custom function handler
       */
      export function clock(invocation: CustomFunctions.StreamingInvocation<string>): void {
      const timer = setInterval(() => {
          const time = currentTime();
          invocation.setResult(time);
      }, 1000);
      
      invocation.onCanceled = () => {
          clearInterval(timer);
      };
      }
      `),
      getValidTestCase(`
        /**
         * Adds two numbers.
         * @customfunction
         * @param first First number
         * @param second Second number
         * @returns The sum of the two numbers.
         */
        /* global clearInterval, console, setInterval */
        
        export function add(first: number, second: number): number {
          try {
            Excel.run(function (context) {
              /**
               * Insert your Excel code here
               */
              context.workbook.worksheets.add();                              // ERROR: context.workbook.worksheets.add()
              var sheetFunc = context.workbook.worksheets.getItem;
              var sheet = sheetFunc("Sheet1");                                // ERROR: sheet.showOutlineLevels(1,1)
              sheet.showOutlineLevels(1,1);
              var sheet2 = sheetFunc("Sheet2");
              const range = sheet.getRange("A1:C3");
          
        
              let myExcel = Excel;
        
              // Update the fill color
              range.format.fill.color = "yellow";                             // ERROR: range.format.fill.color = "yellow"
        
              let wow = myExcel.Range.length;
              
              console.log(wow);
              return context.sync();
            });
          } catch (error) {
            return 12;
          }
          return first + second;
        }
      `)
  ],
  // Error cases. `// ERROR: x` marks the spot where the error occurs.
  invalid: [
    getInvalidTestCase(`
      /**
       * Adds two numbers.
       * @customfunction
       * @param first First number
       * @param second Second number
       * @returns The sum of the two numbers.
       */
      /* global clearInterval, console, setInterval */
      
      export function add(first: number, second: number): number {
        try {
          Excel.run(function (context) {
            /**
             * Insert your Excel code here
             */
            context.workbook.worksheets.add();                              // ERROR: context.workbook.worksheets.add()
            var sheetFunc = context.workbook.worksheets.getItem;
            var sheet = sheetFunc("Sheet1");                                // ERROR: sheet.showOutlineLevels(1,1)
            sheet.showOutlineLevels(1,1);
            var sheet2 = sheetFunc("Sheet2");
            const range = sheet.getRange("A1:C3");
        
      
            let myExcel = Excel;
      
            // Update the fill color
            range.format.fill.color = "yellow";                             // ERROR: range.format.fill.color = "yellow"
      
            let wow = myExcel.Range.length;
            
            console.log(wow);
            return context.sync();
          });
        } catch (error) {
          return 12;
        }
        return first + second;
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
