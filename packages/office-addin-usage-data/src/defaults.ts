// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import os from "os";
import path from "path";

export const usageDataJsonFilePath: string = path.join(
  os.homedir(),
  "/office-addin-usage-data.json"
);
export const groupName: string = "office-addin-usage-data";
export const instrumentationKeyForOfficeAddinCLITools: string =
  "de0d9e7c-1f46-4552-bc21-4e43e489a015";
export const connectionStringForOfficeAddinCLITools: string =
  "InstrumentationKey=de0d9e7c-1f46-4552-bc21-4e43e489a015;IngestionEndpoint=https://westus2-4.in.applicationinsights.azure.com/;LiveEndpoint=https://westus2.livediagnostics.monitor.azure.com/;ApplicationId=8877f7e6-d54b-4ce8-a5c3-0393f788f174";
export const generatorOffice: string = "generator-office";
