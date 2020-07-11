"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
exports.__esModule = true;
var experimental_utils_1 = require("@typescript-eslint/experimental-utils");
var resolve_from_1 = __importDefault(require("resolve-from"));
var no_office_api_calls_1 = __importDefault(require("../no-office-api-calls"));
var ruleTester = new experimental_utils_1.TSESLint.RuleTester({
    parser: resolve_from_1["default"](require.resolve('typescript-eslint'), 'parser'),
    parserOptions: { ecmaVersion: 2015 }
});
ruleTester.run('no-office-api-calls', no_office_api_calls_1["default"], {
    valid: [
        {
            code: "\n                /**\n                 * Displays the current time once a second.\n                 * @customfunction\n                 * @param invocation Custom function handler\n                 */\n                export function clock(invocation: CustomFunctions.StreamingInvocation<string>): void {\n                const timer = setInterval(() => {\n                    const time = currentTime();\n                    invocation.setResult(time);\n                }, 1000);\n                \n                invocation.onCanceled = () => {\n                    clearInterval(timer);\n                };\n                }\n            "
        }
    ],
    invalid: [
        {
            code: "\n        /**\n         * Adds two numbers.\n         * @customfunction\n         * @param first First number\n         * @param second Second number\n         * @returns The sum of the two numbers.\n         */\n        /* global clearInterval, console, setInterval */\n        \n        export function add(first: number, second: number): number {\n          try {\n            Excel.run(function (context) {\n              /**\n               * Insert your Excel code here\n               */\n              var sheet = context.workbook.worksheets.getItem(\"Sheet1\");\n              const range = sheet.getRange(\"A1:C3\");\n        \n              // Update the fill color\n              range.format.fill.color = \"yellow\"; \n        \n              return context.sync();\n            });\n          } catch (error) {\n            return 69;\n          }\n          return first + second;\n        }\n        ",
            errors: [{ messageId: 'contextSync', column: 9, line: 3 }]
        },
    ]
});
