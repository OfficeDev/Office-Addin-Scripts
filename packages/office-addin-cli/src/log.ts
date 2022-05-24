// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as OfficeAddinUsageData from "office-addin-usage-data";

/**
 * Logs an error message
 * @param err The error to be logged
 * @deprecated scriptName The name of the script
 */
export function logErrorMessage(err: any) {
  OfficeAddinUsageData.logErrorMessage(err);
}
