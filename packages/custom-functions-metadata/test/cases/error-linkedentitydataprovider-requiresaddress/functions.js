// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

/**
 * Test the linkedEntityDataProvider tag in combination with the requiresAddress tag
 * @param linkedEntityId Unique `LinkedEntityId` of the `LinkedEntityCellValue`s which is being requested for resolution/refresh.
 * @param handler {CustomFunctions.Invocation} my handler
 * @customfunction
 * @linkedEntityDataProvider
 * @requiresAddress
 * @returns {Promise<any>} Resolved/Updated `LinkedEntityCellValue` that was requested by the passed-in id.
 */
function linkedEntityDataProviderTest(linkedEntityId, handler) {
    // Empty
}
