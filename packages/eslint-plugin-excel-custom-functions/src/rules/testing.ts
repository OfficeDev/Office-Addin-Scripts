// async function abc() {
//     await Excel.run(async ctx => {
//         ctx.application.calculationMode = "Manual";
//     });
// }
  
//   /**
//  * Adds two numbers.
//  * @customfunction
//  * @param first First number
//  * @param second Second number
//  * @returns The sum of the two numbers.
//  */
// /* global clearInterval, console, setInterval */
// export async function add(first: number, second: number): Promise<number> {
//     await abc(); //This line would not have been caught previously
//     return first + second;
// }