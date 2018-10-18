/**
 * Test comments
 * @param {number} first - the first number
 * @param {number} [second]
 * @param {number} [optional] - an optional number 
 * @helpUrl https://dev.office.com
 * @CustomFunction
 * @notfound test123
 * @volatile
 * @streaming cancelable
 * @return {returntypetest}
 */
function add(first: number, second: number, optional?: number): number {
    return first + second;
}

/**
 * @param {string} one - onetest
 * @param {number} x - x arraynumber
 * @CustomFunction
 */
function bad(one: string[][], x: Array<Array<number>>): string[][] {
    return [""][""];
}

function notadded() {
}

/**
 * Testing boolean
 * @CustomFunction
 */
function testbool(abc: boolean): string {
    return "";
}

/**
 * @CustomFunction
 */
function testdatetime(d?: number): string {
    return "";
}