// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { ManifestHandler } from "./manifestHandler";
import { ManifestHandlerJson } from "./manifestHandlerJson";
import { ManifestHandlerXml } from "./manifestHandlerXml";

export async function getManifestHandler(manifestPath: string): Promise<ManifestHandler> {
  let manifestHandler: ManifestHandler;
  if (manifestPath.endsWith(".json")) {
    manifestHandler = new ManifestHandlerJson();
  } else if (manifestPath.endsWith(".xml")) {
    manifestHandler = new ManifestHandlerXml();
  } else {
    const extension: string = manifestPath.split(".").pop() ?? "<no extension>";
    throw new Error(
      `Manifest operations are not supported in .${extension}.\nThey are only supported in .xml and in .json.`
    );
  }
  await manifestHandler.readFromManifestFile(manifestPath);
  return manifestHandler;
}
