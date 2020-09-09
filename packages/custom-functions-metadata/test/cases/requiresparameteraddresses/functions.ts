// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

/**
 * Test the requiresParameterAddresses tag
 * @param x string
 * @param handler my handler
 * @customfunction
 * @requiresParameterAddresses
 */
function requiresParameterAddressesTest(x: string, handler: CustomFunctions.Invocation) {
    // Empty
}

/**
 * Test the requiresParameterAddresses tag with multiple parameters
 * @param x string
 * @param y string
 * @param handler my handler
 * @customfunction
 * @requiresParameterAddresses
 */
function requiresParameterAddressesWithMultipleParameterTest(x: string, y: string, handler: CustomFunctions.Invocation) {
    // Empty
}

/**
 * Test having both requireAddress and requiresParameterAddresses tag
 * @param x string
 * @param handler my handler
 * @customfunction
 * @requiresAddress
 * @requiresParameterAddresses
 */
function requiresBothAddressesAndParameterAddressesTest(x: string, handler: CustomFunctions.Invocation) {
    // Empty
}