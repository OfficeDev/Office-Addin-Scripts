// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

/**
 * @customfunction
 */
function returnsBoolean() {
  return true;
}

/**
 * @customfunction
 */
function returnsNumber() {
  return 5;
}

/**
 * @customfunction
 */
function returnsString() {
  return "abc";
}

/**
 * @customfunction
 */
function returnsObject() {
  return {};
}

/**
 * @customfunction
 */
function returnsMatrixBoolean() {
  return [[true]];
}

/**
 * @customfunction
 */
function returnsMatrixNumber() {
  return [[5]];
}

/**
 * @customfunction
 */
function returnsMatrixString() {
  return [["abc"]];
}

/**
 * @customfunction
 */
function returnsBooleanPromise() {
  return Promise.resolve(true);
}

/**
 * @customfunction
 */
function returnsNumberPromise() {
  return Promise.resolve(5);
}

/**
 * @customfunction
 */
function returnsStringPromise() {
  return Promise.resolve("abc");
}

/**
 * @customfunction
 */
function returnsObjectPromise() {
  return Promise.resolve({});
}

/**
 * @customfunction
 */
function returnsMatrixBooleanPromise() {
  return Promise.resolve([[true]]);
}

/**
 * @customfunction
 */
function returnsMatrixNumberPromise() {
  return Promise.resolve([[5]]);
}

/**
 * @customfunction
 */
function returnsMatrixStringPromise() {
  return Promise.resolve([["abc"]]);
}
