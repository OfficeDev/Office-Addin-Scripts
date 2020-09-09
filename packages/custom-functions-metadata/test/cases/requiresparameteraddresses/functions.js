// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

/**
 * Test the requiresParameterAddresses tag
 * @param x {string} string
 * @param handler {CustomFunctions.Invocation} my handler
 * @customfunction
 * @requiresParameterAddresses
 */
function requiresParameterAddressesTest(x, handler) {
    // Empty
}

/**
 * Test the requiresParameterAddresses tag with multiple parameters
 * @param x {string} string
 * @param y {string} string
 * @param handler {CustomFunctions.Invocation} my handler
 * @customfunction
 * @requiresParameterAddresses
 */
function requiresParameterAddressesWithMultipleParameterTest(x, y, handler) {
    // Empty
}

/**
 * Test having both requireAddress and requiresParameterAddresses tag
 * @param x {string} string
 * @param handler {CustomFunctions.Invocation} my handler
 * @customfunction
 * @requiresAddress
 * @requiresParameterAddresses
 */
function requiresBothAddressesAndParameterAddressesTest(x, handler) {
    // Empty
}