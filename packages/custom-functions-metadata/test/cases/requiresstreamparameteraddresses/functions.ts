// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

/**
 * Test the requiresStreamParameterAddresses tag
 * @param x string
 * @param handler my handler
 * @customfunction
 * @requiresStreamParameterAddresses
 * @streaming
 */
function requiresStreamParameterAddressesTest(x: string, handler: CustomFunctions.StreamingInvocation<string>) {
    // Empty
}

/**
 * Test the requiresStreamParameterAddresses tag with multiple parameters
 * @param x string
 * @param y string
 * @param handler my handler
 * @customfunction
 * @requiresStreamParameterAddresses
 * @streaming
 */
function requiresStreamParameterAddressesWithMultipleParameterTest(x: string, y: string, handler: CustomFunctions.StreamingInvocation<string>) {
    // Empty
}

/**
 * Test having both requiresStreamAddress and requiresStreamParameterAddresses tag
 * @param x string
 * @param handler my handler
 * @customfunction
 * @requiresStreamAddress
 * @requiresStreamParameterAddresses
 * @streaming
 */
function requiresStreamBothAddressesAndParameterAddressesTest(x: string, handler: CustomFunctions.StreamingInvocation<string>) {
    // Empty
}