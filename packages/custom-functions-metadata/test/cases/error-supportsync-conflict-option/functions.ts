// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

/**
 * Test supportSync can't coexist with streaming
 * @param x string
 * @customfunction
 * @supportSync
 * @streaming
 */
async function customFunctionSyncStream(x: string) {
}

/**
 * Test supportSync can't coexist with volatile
 * @param x string
 * @customfunction
 * @supportSync
 * @volatile
 */
async function customFunctionSyncVolatile(x: string) {
}
