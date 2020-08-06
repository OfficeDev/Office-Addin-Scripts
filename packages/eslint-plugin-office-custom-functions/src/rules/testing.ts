
// /**
//  * Adds two numbers.
//  * @customfunction
//  * @param first First number
//  * @param second Second number
//  * @returns The sum of the two numbers.
//  */
// /* global clearInterval, console, setInterval */

// export function add(first: number, second: number): number {
//   try {
//     Excel.run(function (context) {
//       /**
//        * Insert your Excel code here
//        */
//       var sheet = context.workbook.worksheets.getItem("Sheet1");
//       const range = sheet.getRange("A1:C3");
  

//       let myExcel = Excel;

//       // Update the fill color
//       range.format.fill.color = "yellow";                           // ERROR: range.format.fill.color = "yellow"

//       let wow = myExcel.Range.length;

//       console.log(wow);
//       return context.sync();
//     });
//   } catch (error) {
//     return 69;
//   }
//   return first + second;
// }