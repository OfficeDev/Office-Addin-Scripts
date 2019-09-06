// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as os from "os";
import * as path from "path";

export const usageDataJsonFilePath: string = path.join(os.homedir(), "/office-addin-usage-data.json");
export const groupName = "office-addin-usage-data";
export const instrumentationKey = "de0d9e7c-1f46-4552-bc21-4e43e489a015";
