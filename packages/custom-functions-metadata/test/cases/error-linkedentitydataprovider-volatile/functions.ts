// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

/**
 * Test the linkedEntityDataProvider tag in combination with the volatile tag
 * @param {any} linkedEntityId Unique `LinkedEntityId` of the `LinkedEntityCellValue`s which is being requested for resolution/refresh.
 * @customfunction
 * @linkedEntityDataProvider
 * @volatile
 * @returns {Promise<any>} Resolved/Updated `LinkedEntityCellValue` that was requested by the passed-in id.
 */
async function linkedEntityDataProviderTest(linkedEntityId: any): Promise<any> {
    // Empty
}
