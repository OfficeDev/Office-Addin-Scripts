
// /**
//      * Adds two numbers.
//      * @customfunction
//      * @param first First number
//      * @param second Second number
//      * @returns The sum of the two numbers.
//      */
//     /* global clearInterval, console, setInterval */
    
//     export function add(first: number, second: number): number {
//     try {
//         Excel.run(function (context) {                                    // WARN: Excel.run()
//         /**
//          * Insert your Excel code here
//          */
//         context.workbook.worksheets.add();                              
//         var sheetFunc = context.workbook.worksheets.getItem;            // WARN: sheetFunc = context.workbook.worksheets.getItem
//         var sheet = sheetFunc("Sheet1");                                // WARN: sheet = sheetFunc("Sheet1")
//         sheet.showOutlineLevels(1,1);
//         var sheet2 = sheetFunc("Sheet2");                               // WARN: sheet2 = sheetFunc("Sheet2")
//         const range = sheet.getRange("A1:C3");                          // WARN: range = sheet.getRange("A1:C3")
    
    
//         let myExcel = Excel;                                            // WARN: myExcel = Excel
    
//         // Update the fill color
//         range.format.fill.color = "yellow";
    
//         let wow = myExcel.Range.length;
        
//         console.log(wow);
//     return context.sync();                                              // WARN: context.sync()
//         });
//     } catch (error) {
//         return 12;
//     }
//     return first + second;
//     }