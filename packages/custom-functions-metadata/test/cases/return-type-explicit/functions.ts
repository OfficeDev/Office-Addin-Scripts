// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

/**
 * @customfunction
 */
function returnsBoolean(): boolean {
  return true;
}

/**
 * @customfunction
 */
function returnsNumber(): number {
  return 5;
}

/**
 * @customfunction
 */
function returnsString(): string {
  return "abc";
}

/**
 * @customfunction
 */
function returnsObject(): object {
  return {};
}

/**
 * @customfunction
 */
function returnsMatrixBoolean(): boolean[][] {
  return [[true]];
}

/**
 * @customfunction
 */
function returnsMatrixNumber(): number[][] {
  return [[5]];
}

/**
 * @customfunction
 */
function returnsMatrixString(): string[][] {
  return [["abc"]];
}

/**
 * @customfunction
 */
function returnsBooleanPromise(): Promise<boolean> {
  return Promise.resolve(true);
}

/**
 * @customfunction
 */
function returnsNumberPromise(): Promise<number> {
  return Promise.resolve(5);
}

/**
 * @customfunction
 */
function returnsStringPromise(): Promise<string> {
  return Promise.resolve("abc");
}

/**
 * @customfunction
 */
function returnsObjectPromise(): Promise<object> {
  return Promise.resolve({});
}

/**
 * @customfunction
 */
function returnsMatrixBooleanPromise(): Promise<boolean[][]> {
  return Promise.resolve([[true]]);
}

/**
 * @customfunction
 */
function returnsMatrixNumberPromise(): Promise<number[][]> {
  return Promise.resolve([[5]]);
}

/**
 * @customfunction
 */
function returnsMatrixStringPromise(): Promise<string[][]> {
  return Promise.resolve([["abc"]]);
}
