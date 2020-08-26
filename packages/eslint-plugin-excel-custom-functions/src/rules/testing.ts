
// /**
//  * Adds two numbers.
//  * @customfunction
//  * @param first First number
//  * @param second Second number
//  * @returns The sum of the two numbers.
//  */
// /* global clearInterval, console, setInterval */

// function createTable() {
//     Excel.run(function (context) {
//         var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
//         var expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
//         expensesTable.name = "ExpensesTable";
//         expensesTable.getHeaderRowRange().values =
//         [["Date", "Merchant", "Category", "Amount"]];
//         expensesTable.rows.add(undefined /*add at the end*/, [
//             ["1/1/2017", "The Phone Company", "Communications", "120"],
//             ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
//             ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
//             ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
//             ["1/11/2017", "Bellows College", "Education", "350.1"],
//             ["1/15/2017", "Trey Research", "Other", "135"],
//             ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
//         ]);

//     expensesTable.columns.getItemAt(3).getRange().numberFormat = [['\u20AC#,##0.00']];
//     expensesTable.getRange().format.autofitColumns();
//     expensesTable.getRange().format.autofitRows();

//         return context.sync();
//     })
//     .catch(function (error) {
//         console.log("Error: " + error);
//         if (error instanceof OfficeExtension.Error) {
//             console.log("Debug info: " + JSON.stringify(error.debugInfo));
//         }
//     });
// }