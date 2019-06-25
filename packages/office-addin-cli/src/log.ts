// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

export function logErrorMessage(err: any) {
    console.error(`Error: ${err instanceof Error ? err.message : err}`);
}
