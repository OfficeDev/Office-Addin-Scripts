// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

/**
 * Test the requiresStreamParameterAddresses tag
 * @param x {string} string
 * @param handler {CustomFunctions.StreamingInvocation<string>} my handler
 * @customfunction
 * @requiresStreamParameterAddresses
 * @streaming
 */
function requiresStreamParameterAddressesTest(x, handler) {
    // Empty
}

/**
 * Test the requiresStreamParameterAddresses tag with multiple parameters
 * @param x {string} string
 * @param y {string} string
 * @param handler {CustomFunctions.StreamingInvocation<string>} my handler
 * @customfunction
 * @requiresStreamParameterAddresses
 * @streaming
 */
function requiresStreamParameterAddressesWithMultipleParameterTest(x, y, handler) {
    // Empty
}

/**
 * Test having both requiresStreamAddress and requiresStreamParameterAddresses tag
 * @param x {string} string
 * @param handler {CustomFunctions.StreamingInvocation<string>} my handler
 * @customfunction
 * @requiresStreamAddress
 * @requiresStreamParameterAddresses
 * @streaming
 */
function requiresStreamBothAddressesAndParameterAddressesTest(x, handler) {
    // Empty
}