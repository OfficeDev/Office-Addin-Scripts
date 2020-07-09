import { TSESLint } from '@typescript-eslint/experimental-utils';
import resolveFrom from 'resolve-from';

import rule from '../no-office-api-calls';

const ruleTester = new TSESLint.RuleTester({
    parser: resolveFrom(require.resolve('typescript-eslint'), 'parser'),
    parserOptions: { ecmaVersion: 2015 },
  });

ruleTester.run('no-office-api-calls', rule, {
    valid: [

        {
            code: `
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
            `
        }
    ],
    invalid: [
        {
        code: `
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
              var sheet = context.workbook.worksheets.getItem("Sheet1");
              const range = sheet.getRange("A1:C3");
        
              // Update the fill color
              range.format.fill.color = "yellow"; 
        
              return context.sync();
            });
          } catch (error) {
            return 69;
          }
          return first + second;
        }
        `,
        errors: [{ messageId: 'contextSync', column: 9, line: 3 }],
        },

    ],
});