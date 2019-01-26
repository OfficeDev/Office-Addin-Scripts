/**
 * Test comments
 * @param {number} first - the first number
 * @param {number} second 
 * @helpUrl https://dev.office.com
 * @CustomFunction
 * @notfound test123
 * @volatile
 * @streaming cancelable
 * @return {returntypetest}
 */
function add(first: number, second: number): number {
    return first + second;
}

/**
 * @param {string} one - onetest
 * @param {number} x - x arraynumber
 * @CustomFunction
 */
function complexFunction(one: string[][], x: Array<Array<number>>): string[][] {
    return [""][""];
}

function notAdded() {
}

/**
 * Testing boolean
 * @CustomFunction
 */
function testBool(abc: boolean): boolean {
    return true;
}

/**
 * Test function for number type
 * @param one - A number
 * @CustomFunction
 */
function testNumber(one: number): number {
    return 0;
}

/**
 * Test function for string type
 * @param word - Some string
 * @CustomFunction
 */
function testString(word: string): string {
    return "";
}

/**
 * Test function for void type
 * @CustomFunction
 */
function voidTest(): void {
}

/**
 * Test function for object type
 * @param obj - Some object
 * @CustomFunction
 */
function objectTest(obj: object): object {
    let o;
    return o;
}

/**
 * @CustomFunction
 */
function testdatetime(d: number): string {
    return "";
}

enum Color {Red,Green,Blue};

/**
 * Test function for enum type
 * @param e - enum type
 * @CustomFunction
 */
function enumTest(e: Color) : Color {
    let r : Color;
    return r;
}

/**
 * Test function for tuple type
 * @param t 
 * @CustomFunction
 */
function tupleTest(t:[string,number]){
}

/**
 * Test function for streaming type
 * @param x - Test string
 * @param sf - Streaming function type return type should be number
 * @CustomFunction 
 */
function streamingTest(x: string, sf: CustomFunctions.StreamingHandler<number>) {
}

/**
 * Test function for optional parameter
 * @param x - Optional string
 * @CustomFunction
 */
function testOptional(x?: string){
}

/**
 * Test any type
 * @CustomFunction
 */
function testAny(a: any): any {}

/**
 * Test support for the CustomFunctions.CancelableHandler
 * @param x string parameter
 * @param cf Cancelable Handler parameter
 * @CustomFunction
 */
async function testCancelableFunction(x: string, cf: CustomFunctions.CancelableHandler ): Promise<number> {
    return 1;
}

/**
 * Test the custom function id and name
 * @param x test string
 * @CustomFunction updateId updateName
 */
function customFunctionIdNameTest(x:string){
}

/**
 * Test the requiresAddress tag
 * @param x string
 * @param handler my handler
 * @CustomFunction
 * @requiresAddress
 */
function requiresAddressTest(x: string, handler: CustomFunctions.StreamingHandler<number>){}

CustomFunctionMappings.ADD=add;

