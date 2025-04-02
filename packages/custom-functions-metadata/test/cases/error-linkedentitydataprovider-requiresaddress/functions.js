// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

/**
 * Test the linkedEntityLoadService tag in combination with the requiresAddress tag
 * @param request Represents a request to the `@linkedEntityLoadService` custom function to load `LinkedEntityCellValue` objects.
 * @param handler {CustomFunctions.Invocation} my handler
 * @customfunction
 * @linkedEntityLoadService
 * @requiresAddress
 * @returns {Promise<any>} Resolved/Updated `LinkedEntityCellValue` objects that were requested by the passed-in request.
 */
function linkedEntityLoadServiceTest(request, handler) {
    // Empty
}
