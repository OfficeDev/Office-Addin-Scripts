// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { ManifestHandler } from "./manifestHandler";
import { ManifestHandlerJson } from "./manifestHandlerJson";
import { ManifestHandlerXml } from "./manifestHandlerXml";

function isJsonFile(path: string) {
  return path.endsWith(".json");
}

export function getManifestHandler(manifestPath: string): ManifestHandler {
  let manifestHandler: ManifestHandler;
  if (isJsonFile(manifestPath)) {
    manifestHandler = new ManifestHandlerJson();
  } else {
    manifestHandler = new ManifestHandlerXml();
  }
  return manifestHandler;
}
