// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commnder from "commander";
import * as manifestInfo from "./manifestInfo";

export async function info(path: string) {
  try {
    const manifest = await manifestInfo.readManifestFile(path);
    logManifestInfo(path, manifest);
  } catch (err) {
    console.error(`Error: ${err}`);
  }
}

export async function modify(path: string, command: commnder.Command) {
  try {
    const guid: string | undefined = (command.guid) ? command.guid : undefined;
    const displayName: string | undefined = (command.displayName) ? command.displayName : undefined;

    if (guid === undefined && displayName === undefined) {
      throw new Error("You need to specify something to change in the manifest.");
    }

    const manifest = await manifestInfo.modifyManifestFile(path, guid, displayName);
    logManifestInfo(path, manifest);
  } catch (err) {
    console.error(`Error: ${err}`);
  }
}

function logManifestInfo(path: string, manifest: manifestInfo.ManifestInfo) {
  console.log(`Manifest: ${path}`);
  console.log(`  Id: ${manifest.id || ""}`);
  console.log(`  Name: ${manifest.displayName || ""}`);
  console.log(`  Provider: ${manifest.providerName || ""}`);
  console.log(`  Type: ${manifest.officeAppType || ""}`);
  console.log(`  Version: ${manifest.version || ""}`);
  console.log(`  Default Locale: ${manifest.defaultLocale || ""}`);
  console.log(`  Description: ${manifest.description || ""}`);
}
