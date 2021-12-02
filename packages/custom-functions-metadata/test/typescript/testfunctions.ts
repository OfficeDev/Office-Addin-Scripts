// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

/**
 * Test comments
 * @param {number} first - the first number
 * @param {number} second 
 * @helpUrl https://docs.microsoft.com/office/dev/add-ins
 * @customfunction
 * @notfound test123
 * @volatile
 * @streaming
 * @cancelable
 * @return {number}
 */
function add(first: number, second: number): number {
    return first + second;
}

/**
 * @param {string[][]} one - onetest
 * @param {Array<Array<number>>} x - x arraynumber
 * @customfunction
 */
function complexFunction(one: string[][], x: Array<Array<number>>): string[][] {
    return [""][""];
}

function notAdded() {
}

/**
 * Testing boolean
 * @customfunction
 */
function testBool(abc: boolean): boolean {
    return true;
}

/**
 * Test function for number type
 * @param one - A number
 * @customfunction
 */
function testNumber(one: number): number {
    return 0;
}

/**
 * Test function for string type
 * @param word - Some string
 * @customfunction
 */
function testString(word: string): string {
    return "";
}

/**
 * Test function for void type
 * @customfunction
 */
function voidTest(): void {
}

/**
 * Test function for object type
 * @param obj - Some object
 * @customfunction
 */
function objectTest(obj: object): object {
    let o;
    return o;
}

/**
 * @customfunction
 */
function testdatetime(d: number): string {
    return "";
}

enum Color {Red,Green,Blue};

/**
 * Test function for enum type
 * @param e - enum type
 * @customfunction
 */
function enumTest(e: Color) : Color {
    let r : Color;
    return r;
}

/**
 * Test function for tuple type
 * @param t 
 * @customfunction
 */
function tupleTest(t:[string,number]){
}

/**
 * Test function for streaming type
 * @param x - Test string
 * @param sf - Streaming function type return type should be number
 * @customfunction 
 */
function streamingTest(x: string, sf: CustomFunctions.StreamingHandler<number>) {
}

/**
 * Test function for optional parameter
 * @param x - Optional string
 * @customfunction
 */
function testOptional(x?: string){
}

/**
 * Test any type
 * @customfunction
 */
function testAny(a: any): any {}

/**
 * Test support for the CustomFunctions.CancelableHandler
 * @param x string parameter
 * @param cf Cancelable Handler parameter
 * @customfunction
 */
async function testCancelableFunction(x: string, cf: CustomFunctions.CancelableHandler ): Promise<number> {
    return 1;
}

/**
 * Test the custom function id and name
 * @param x test string
 * @customfunction updateId updateName
 */
function customFunctionIdNameTest(x:string){
}

/**
 * Test the requiresAddress tag
 * @param x string
 * @param handler my handler
 * @customfunction
 * @requiresAddress
 */
function requiresAddressTest(x: string, handler: CustomFunctions.Invocation){}

/**
 * Test the CustomFunctions.Invocation type
 * @customfunction
 * @param invocation Invocation parameter
 * @requiresAddress
 */
function customFunctionInvocationTest(x: string, invocation: CustomFunctions.Invocation){}

/**
 * Test the new cancelable type
 * @param x string
 * @param cancel CustomFunctions.CancelableInvocation type
 * @customfunction
 * @requiresAddress
 */
function customFunctionCancelableInvocationTest(x: string, cancel: CustomFunctions.CancelableInvocation){}

/**
 * Test the new streaming invocation type
 * @param x string
 * @param stream StreamingInvocation type
 * @customfunction
 */
function customFunctionStreamingInvocationTest(x: string, stream: CustomFunctions.StreamingInvocation<string>){}

/**
 * @customfunction
 */
function UPPERCASE(){}

CustomFunctionMappings.ADD=add;

