// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

/**
 * This function is testing add
 * @customfunction
 * @param {number} number1 - first number
 * @return {number} - return number
 */
function testAdd(number1){
}

/**
 * Test function for boolean type
 * @param {boolean} bool - boolean parameter
 * @return {boolean}
 * @customfunction
 */
function testBoolean(bool){
}

/**
 * Test the optional parameter
 * @param {string} [opt] - this parameter should be optional
 * @customfunction 
 */
function testOptional(opt){}

/**
 * Test the string type
 * @param {string} str - string type parameter
 * @return {string}
 * @customfunction
 */
function testString(str){}

/**
 * Test the any type
 * @param {any} a - any type parameter
 * @return {any}
 * @customfunction
 */
function testAny(a){}

/**
 * Test streaming handler function
 * @param {string} x 
 * @param {CustomFunctions.StreamingHandler<string>} handler
 * @customfunction 
 */
function testStreaming(x, handler){}

/**
 * Test the cancelable handler
 * @param {string} x 
 * @param {CustomFunctions.CancelableHandler} chandler
 * @customfunction 
 */
function testCancel(x, chandler){}

/**
 * Test the custom function id
 * @param {string} x
 * @customfunction newIdTest 
 */
function customIdTest(x){}

/**
 * Test the custom function id and name
 * @param {string} x
 * @customfunction newId newName 
 */
function customIdNameTest(x){}

/**
 * Test the new invocation type
 * @param {CustomFunctions.Invocation} inv
 * @customfunction
 * @requiresAddress 
 */
function customInvocationTest(inv){}

/**
 * Test streaming handler function
 * @param {string} x 
 * @param {CustomFunctions.StreamingInvocation<string>} handler
 * @customfunction 
 */
function testStreamingInvocation(x, handler){}

/**
 * Test the cancelable handler
 * @param {string} x 
 * @param {CustomFunctions.CancelableInvocation} chandler
 * @customfunction 
 */
function testCancelInvocation(x, chandler){}