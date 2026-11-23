// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

/**
 * Adds two numbers. See {@link https://example.com/add} for details.
 * @customfunction
 * @param first First number, see {@link https://example.com/first}
 * @param second Second number
 * @returns The sum
 */
function addWithLink(first: number, second: number): number {
  return first + second;
}

/**
 * Multiplies two numbers.
 * @customfunction
 * @param first - First number, see {@link https://example.com/first | the docs}
 * @param second Second number
 */
function multiplyWithParamLink(first: number, second: number): number {
  return first * second;
}
