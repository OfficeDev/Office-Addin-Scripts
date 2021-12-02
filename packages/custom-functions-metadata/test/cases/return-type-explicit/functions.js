// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

/**
 * @customfunction
 * @returns {boolean}
 */
function returnsBoolean() {
  return true;
}

/**
 * @customfunction
 * @returns {number}
 */
function returnsNumber() {
  return 5;
}

/**
 * @customfunction
 * @returns {string}
 */
function returnsString() {
  return "abc";
}

/**
 * @customfunction
 * @returns {object}
 */
function returnsObject() {
  return {};
}

/**
 * @customfunction
 * @returns {boolean[][]}
 */
function returnsMatrixBoolean() {
  return [[true]];
}

/**
 * @customfunction
 * @returns {number[][]}
 */
function returnsMatrixNumber() {
  return [[5]];
}

/**
 * @customfunction
 * @returns {string[][]}
 */
function returnsMatrixString() {
  return [["abc"]];
}

/**
 * @customfunction
 * @returns {Promise<boolean>}
 */
function returnsBooleanPromise() {
  return Promise.resolve(true);
}

/**
 * @customfunction
 * @returns {Promise<number>}
 */
function returnsNumberPromise() {
  return Promise.resolve(5);
}

/**
 * @customfunction
 * @returns {Promise<string>}
 */
function returnsStringPromise() {
  return Promise.resolve("abc");
}

/**
 * @customfunction
 * @returns {Promise<object>}
 */
function returnsObjectPromise() {
  return Promise.resolve({});
}

/**
 * @customfunction
 * @returns {Promise<boolean[][]>}
 */
function returnsMatrixBooleanPromise() {
  return Promise.resolve([[true]]);
}

/**
 * @customfunction
 * @returns {Promise<number[][]>}
 */
function returnsMatrixNumberPromise() {
  return Promise.resolve([[5]]);
}

/**
 * @customfunction
 * @returns {Promise<string[][]>}
 */
function returnsMatrixStringPromise() {
  return Promise.resolve([["abc"]]);
}
