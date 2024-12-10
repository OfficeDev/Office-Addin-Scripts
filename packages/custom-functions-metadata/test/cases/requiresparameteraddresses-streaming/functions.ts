// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

/**
 * Test the requiresParameterAddresses and streaming tag
 * @param x string
 * @param invocation my invocation
 * @customfunction
 * @requiresParameterAddresses
 * @streaming
 */
function requiresStreamParameterAddressesTest(x: string, invocation: CustomFunctions.StreamingInvocation<string>) {
    // Empty
}

/**
 * Test the requiresParameterAddresses and streaming tag with multiple parameters
 * @param x string
 * @param y string
 * @param invocation my invocation
 * @customfunction
 * @requiresParameterAddresses
 * @streaming
 */
function requiresStreamParameterAddressesWithMultipleParameterTest(x: string, y: string, invocation: CustomFunctions.StreamingInvocation<string>) {
    // Empty
}

/**
 * Test having requiresAddress, requiresParameterAddresses and streaming tag
 * @param x string
 * @param invocation my invocation
 * @customfunction
 * @requiresAddress
 * @requiresParameterAddresses
 * @streaming
 */
function requiresStreamBothAddressesAndParameterAddressesTest(x: string, invocation: CustomFunctions.StreamingInvocation<string>) {
    // Empty
}