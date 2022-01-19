// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { ManifestHandler } from "./manifestHandler";
import { ManifestHandlerJson } from "./manifestHandlerJson";
import { ManifestHandlerXml } from "./manifestHandlerXml";

export function getManifestHandler(manifestPath: string): ManifestHandler {
  if (manifestPath.endsWith(".json")) {
    return new ManifestHandlerJson();
  } else {
    return new ManifestHandlerXml();
  }
}
