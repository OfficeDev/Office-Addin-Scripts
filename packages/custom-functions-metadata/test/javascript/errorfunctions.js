/**
 * This function is testing add
 * @CustomFunction
 * @param {badtype} num1 - testing add
 * @return {number} - return number
 */
function testadd(num1){
}

/**
 * Testing bad result type
 * @CustomFunction
 * @return {badreturn} - not a return type
 */
function badResult(){
}

/**
 * Function contains a bad id
 * @CustomFunction id-bad
 */
function badId(){}

/**
 * This funciton has an invalid name
 * @CustomFunction id 1invalidname
 */
function badName(){}

/**
 * requiresAddress tag requires parmeter to be of type Invocation
 * @param {string} x
 * @CustomFunction
 * @requiresAddress 
 */
function missingInvocationType(x){}

/** @customfunction */
function привет() {
}

/**
 * Duplicate function name test
 */
function testadd(){
}

/**
 * Duplicate function name test set in CustomFunction tag
 * @CustomFunction id testadd
 */
function customNameTest(){
}

/**
 * First custom function with name in CustomFunction tag
 * @CustomFunction myid1 myName1
 */
function customNameTest2(){
}

/**
 * Custom function with duplicate name in CustomFunction tag
 * @CustomFunction myid2 myName1
 */
function customIdTest(){
}

/**
 * Custom function with duplicate id in CustomFunction tag
 * @CustomFunction myid2 myName3
 */
function customIdTest2(){
}

/**
 * Custom function name with duplicate id in CustomFunction tag
 * @CustomFunction
 */
function myid2(){
}

