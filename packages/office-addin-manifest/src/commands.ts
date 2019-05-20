// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commnder from "commander";
import { logErrorMessage } from "office-addin-cli";
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

function logManifestInfo(manifestPath: string, manifest: manifestInfo.ManifestInfo) {
  console.log(`Manifest: ${manifestPath}`);
  console.log(`  Id: ${manifest.id || ""}`);
  console.log(`  Name: ${manifest.displayName || ""}`);
  console.log(`  Provider: ${manifest.providerName || ""}`);
  console.log(`  Type: ${manifest.officeAppType || ""}`);
  console.log(`  Version: ${manifest.version || ""}`);
  if (manifest.alternateId) {
    console.log(`  AlternateId: ${manifest.alternateId}`);
  }
  console.log(`  AppDomains: ${manifest.appDomains ? manifest.appDomains.join(", ") : ""}`);
  console.log(`  Default Locale: ${manifest.defaultLocale || ""}`);
  console.log(`  Description: ${manifest.description || ""}`);
  console.log(`  High Resolution Icon Url: ${manifest.highResolutionIconUrl || ""}`);
  console.log(`  Hosts: ${manifest.hosts ? manifest.hosts.join(", ") : ""}`);
  console.log(`  Icon Url: ${manifest.iconUrl || ""}`);
  console.log(`  Permissions: ${manifest.permissions || ""}`);
  console.log(`  Support Url: ${manifest.supportUrl || ""}`);

  if (manifest.defaultSettings) {
    console.log("  Default Settings:");
    console.log(`    Requested Height: ${manifest.defaultSettings.requestedHeight || ""}`);
    console.log(`    Requested Width: ${manifest.defaultSettings.requestedWidth || ""}`);
    console.log(`    Source Location: ${manifest.defaultSettings.sourceLocation || ""}`);
  }
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
