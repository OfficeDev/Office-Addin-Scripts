// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

/* global console */

/**
 * Logs an error message
 * @param err The error to be logged
 */
export function logErrorMessage(err: any) {
  console.error(`Error: ${err instanceof Error ? err.message : err}`);
}
