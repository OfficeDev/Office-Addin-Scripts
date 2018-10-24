// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commnder from "commander";
import * as manifestInfo from "./manifestInfo";

export async function info(manifestPath: string) {
  try {
    const manifest = await manifestInfo.readManifestFile(manifestPath);
    logManifestInfo(manifestPath, manifest);
  } catch (err) {
    logErrorMessage(err);
  }
}

function logErrorMessage(err: any) {
  console.error(`Error: ${err instanceof Error ? err.message : err}`);
}

function logManifestInfo(manifestPath: string, manifest: manifestInfo.ManifestInfo) {
  console.log(`Manifest: ${manifestPath}`);
  console.log(`  Id: ${manifest.id || ""}`);
  console.log(`  Name: ${manifest.displayName || ""}`);
  console.log(`  Provider: ${manifest.providerName || ""}`);
  console.log(`  Type: ${manifest.officeAppType || ""}`);
  console.log(`  Version: ${manifest.version || ""}`);
  console.log(`  Default Locale: ${manifest.defaultLocale || ""}`);
  console.log(`  Description: ${manifest.description || ""}`);
}

export async function modify(manifestPath: string, command: commnder.Command) {
  try {
    const guid: string | undefined = (command.guid) ? command.guid : undefined;
    const displayName: string | undefined = (command.displayName) ? command.displayName : undefined;

    if (guid === undefined && displayName === undefined) {
      throw new Error("You need to specify something to change in the manifest.");
    }

    const manifest = await manifestInfo.modifyManifestFile(manifestPath, guid, displayName);
    logManifestInfo(manifestPath, manifest);
  } catch (err) {
    logErrorMessage(err);
  }
}