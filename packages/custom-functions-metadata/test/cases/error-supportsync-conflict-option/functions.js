// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

/**
 * Test supportSync can't coexist with stream
 * @param x string
 * @customfunction
 * @supportSync
 * @stream
 */
async function customFunctionSyncStream(x) {
}

/**
 * Test supportSync can't coexist with stream
 * @param x string
 * @customfunction
 * @supportSync
 * @volatile
 */
async function customFunctionSyncVolatile(x) {
}
