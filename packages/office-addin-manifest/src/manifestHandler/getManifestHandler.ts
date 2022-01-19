// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { ManifestHandler } from "./manifestHandler";
import { ManifestHandlerJson } from "./manifestHandlerJson";
import { ManifestHandlerXml } from "./manifestHandlerXml";

export async function getManifestHandler(manifestPath: string): Promise<ManifestHandler> {
  let manifestHandler: ManifestHandler;
  if (manifestPath.endsWith(".json")) {
    manifestHandler = new ManifestHandlerJson();
  } else {
    manifestHandler = new ManifestHandlerXml();
  }
  await manifestHandler.readFromManifestFile(manifestPath);
  return manifestHandler;
}
