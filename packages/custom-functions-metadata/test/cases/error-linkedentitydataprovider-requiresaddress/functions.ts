// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

/**
 * Test the linkedEntityLoadService tag in combination with the requiresAddress tag
 * @param request Represents a request to the `@linkedEntityLoadService` custom function to load `LinkedEntityCellValue` objects.
 * @param handler my handler
 * @customfunction
 * @linkedEntityLoadService
 * @requiresAddress
 * @returns Resolved/Updated `LinkedEntityCellValue` objects that were requested by the passed-in request.
 */
async function linkedEntityLoadServiceTest(request: unknown, handler: CustomFunctions.Invocation): Promise<any> {
    // Empty
}
