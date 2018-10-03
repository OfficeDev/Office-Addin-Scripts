// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commnder from "commander";
import { readManifestFile, personalizeManifestFile, ManifestInfo } from "./manifestInfo";

export async function info(path: string) {
  try {
    const manifest = await readManifestFile(path);

    console.log(`Manifest: ${path}`);
    console.log(`  Id: ${manifest.id || ""}`);
    console.log(`  Name: ${manifest.displayName || ""}`);
    console.log(`  Provider: ${manifest.providerName || ""}`);
    console.log(`  Type: ${manifest.officeAppType || ""}`);
    console.log(`  Version: ${manifest.version || ""}`);
    console.log(`  Default Locale: ${manifest.defaultLocale || ""}`);
    console.log(`  Description: ${manifest.description || ""}`);
  } catch (err) {
    console.error(`Error: ${err}`);
  }
}


export async function personalize(path: string, command: commnder.Command) {
  try {
    const guid: string | undefined = (command.guid) ? command.guid : undefined;
    const displayName: string | undefined = (command.displayName) ? command.displayName: undefined;

    if (guid == undefined && displayName == undefined) {
      throw('No parameters passed with personalize command - exiting command.')
    }

    await personalizeManifestFile(path, guid, displayName);
    const manifest = await readManifestFile(path);

    console.log(`Manifest: ${path}`);
    console.log(`  Id: ${manifest.id || ""}`);
    console.log(`  Name: ${manifest.displayName || ""}`);
    console.log(`  Provider: ${manifest.providerName || ""}`);
    console.log(`  Type: ${manifest.officeAppType || ""}`);
    console.log(`  Version: ${manifest.version || ""}`);
    console.log(`  Default Locale: ${manifest.defaultLocale || ""}`);
    console.log(`  Description: ${manifest.description || ""}`);
  } catch (err) {
    console.error(`Error: ${err}`);
  }
}
