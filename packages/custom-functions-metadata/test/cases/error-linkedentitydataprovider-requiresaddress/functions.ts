// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

/**
 * Test the linkedEntityLoadService tag in combination with the requiresAddress tag
 * @param linkedEntityId Unique `LinkedEntityId` of the `LinkedEntityCellValue`s which is being requested for resolution/refresh.
 * @param handler my handler
 * @customfunction
 * @linkedEntityLoadService
 * @requiresAddress
 * @returns Resolved/Updated `LinkedEntityCellValue` that was requested by the passed-in id.
 */
async function linkedEntityLoadServiceTest(linkedEntityId: unknown, handler: CustomFunctions.Invocation): Promise<any> {
    // Empty
}
