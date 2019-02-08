/**
 * This function is testing add
 * @CustomFunction
 * @param {number} number1 - first number
 * @return {number} - return number
 */
function testAdd(number1){
}

/**
 * Test function for boolean type
 * @param {boolean} bool - boolean parameter
 * @return {boolean}
 * @CustomFunction
 */
function testBoolean(bool){
}

/**
 * Test the optional parameter
 * @param {string} [opt] - this parameter should be optional
 * @CustomFunction 
 */
function testOptional(opt){}

/**
 * Test the string type
 * @param {string} str - string type parameter
 * @return {string}
 * @CustomFunction
 */
function testString(str){}

/**
 * Test the any type
 * @param {any} a - any type parameter
 * @return {any}
 * @CustomFunction
 */
function testAny(a){}

/**
 * Test streaming handler function
 * @param {string} x 
 * @param {CustomFunctions.StreamingHandler<string>} handler
 * @CustomFunction 
 */
function testStreaming(x, handler){}

/**
 * Test the cancelable handler
 * @param {string} x 
 * @param {CustomFunctions.CancelableHandler} chandler
 * @CustomFunction 
 */
function testCancel(x, chandler){}

/**
 * Test the custom function id
 * @param {string} x
 * @CustomFunction newId 
 */
function customIdTest(x){}

/**
 * Test the custom function id and name
 * @param {string} x
 * @CustomFunction newId newName 
 */
function customIdNameTest(x){}

/**
 * Test the new invocation type
 * @param {CustomFunctions.Invocation} inv
 * @CustomFunction
 * @requiresAddress 
 */
function customInvocationTest(inv){}

/**
 * Test streaming handler function
 * @param {string} x 
 * @param {CustomFunctions.StreamingInvocation<string>} handler
 * @CustomFunction 
 */
function testStreamingInvocation(x, handler){}

/**
 * Test the cancelable handler
 * @param {string} x 
 * @param {CustomFunctions.CancelableInvocation} chandler
 * @CustomFunction 
 */
function testCancelInvocation(x, chandler){}