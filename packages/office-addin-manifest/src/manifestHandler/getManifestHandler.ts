// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { ManifestHandler } from "./manifestHandler";
import { ManifestHandlerJson } from "./manifestHandlerJson";
import { ManifestHandlerXml } from "./manifestHandlerXml";

function isJsonObject(file: any) {
  try {
    JSON.parse(file);
  } catch (e) {
    return false;
  }
  return true;
}

export function getManifestHandler(fileData: string): ManifestHandler {
  let manifestHandler: ManifestHandler;
  if (isJsonObject(fileData)) {
    manifestHandler = new ManifestHandlerJson();
  } else {
    manifestHandler = new ManifestHandlerXml();
  }
  return manifestHandler;
}
