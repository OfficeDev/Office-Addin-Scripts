// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commnder from "commander";
import * as manifestInfo from "./manifestInfo";

function getCommandOptionString(option: string | boolean, defaultValue?: string): string | undefined {
  // For a command option defined with an optional value, e.g. "--option [value]",
  // when the option is provided with a value, it will be of type "string", return the specified value;
  // when the option is provided without a value, it will be of type "boolean", return undefined.
  return (typeof(option) === "boolean") ? defaultValue : option;
}

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
    // if the --guid command option is provided without a value, use "" to specify to change to a random guid value.
    const guid: string | undefined = getCommandOptionString(command.guid, "");
    const displayName: string | undefined = getCommandOptionString(command.displayName);

    const manifest = await manifestInfo.modifyManifestFile(manifestPath, guid, displayName);
    logManifestInfo(manifestPath, manifest);
  } catch (err) {
    logErrorMessage(err);
  }
}
