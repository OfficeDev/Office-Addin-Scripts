// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

/**
 * Test the requiresParameterAddresses and streaming tag
 * @param x {string} string
 * @param invocation {CustomFunctions.StreamingInvocation<string>} my invocation
 * @customfunction
 * @requiresParameterAddresses
 * @streaming
 */
function requiresStreamParameterAddressesTest(x, invocation) {
    // Empty
}

/**
 * Test the requiresParameterAddresses and streaming tag with multiple parameters
 * @param x {string} string
 * @param y {string} string
 * @param invocation {CustomFunctions.StreamingInvocation<string>} my invocation
 * @customfunction
 * @requiresParameterAddresses
 * @streaming
 */
function requiresStreamParameterAddressesWithMultipleParameterTest(x, y, invocation) {
    // Empty
}

/**
 * Test having requiresAddress, requiresParameterAddresses and streaming tag
 * @param x {string} string
 * @param invocation {CustomFunctions.StreamingInvocation<string>} my invocation
 * @customfunction
 * @requiresAddress
 * @requiresParameterAddresses
 * @streaming
 */
function requiresStreamBothAddressesAndParameterAddressesTest(x, invocation) {
    // Empty
}