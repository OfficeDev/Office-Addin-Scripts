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
     * Adds two numbers.
     * @customfunction
     * @param first First number
     * @param second Second number
     * @returns The sum of the two numbers.
     */
    /* global clearInterval, console, setInterval */
    
    function createTable() {
        Excel.run(function (context) {
            var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
            var expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
            expensesTable.name = "ExpensesTable";
            expensesTable.getHeaderRowRange().values =
            [["Date", "Merchant", "Category", "Amount"]];
            expensesTable.rows.add(undefined /*add at the end*/, [
                ["1/1/2017", "The Phone Company", "Communications", "120"],
                ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
                ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
                ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
                ["1/11/2017", "Bellows College", "Education", "350.1"],
                ["1/15/2017", "Trey Research", "Other", "135"],
                ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
            ]);
    
        expensesTable.columns.getItemAt(3).getRange().numberFormat = [['\u20AC#,##0.00']];
        expensesTable.getRange().format.autofitColumns();
        expensesTable.getRange().format.autofitRows();
    
            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
      `)
  ],
  // Error cases. `// ERROR: x` marks the spot where the error occurs.
  invalid: [
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
