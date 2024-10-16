// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import chalk from "chalk";
import commander from "commander";
import { logErrorMessage } from "office-addin-usage-data";
import { ManifestInfo } from "./manifestInfo";
import { OfficeAddinManifest } from "./manifestOperations";
import { ManifestValidation, ManifestValidationIssue, ManifestValidationProduct, validateManifest } from "./validate";
import { usageDataObject } from "./defaults";
import { exportMetadataPackage } from "./export";

/* global console, process */

function getCommandOptionString(option: string | boolean, defaultValue?: string): string | undefined {
  // For a command option defined with an optional value, e.g. "--option [value]",
  // when the option is provided with a value, it will be of type "string", return the specified value;
  // when the option is provided without a value, it will be of type "boolean", return undefined.
  return typeof option === "boolean" ? defaultValue : option;
}

export async function info(manifestPath: string) {
  try {
    const manifest = await OfficeAddinManifest.readManifestFile(manifestPath);
    logManifestInfo(manifestPath, manifest);
    usageDataObject.reportSuccess("info");
  } catch (err: any) {
    usageDataObject.reportException("info", err);
    logErrorMessage(err);
  }
}

function logManifestInfo(manifestPath: string, manifest: ManifestInfo) {
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

function logManifestValidationErrors(errors: ManifestValidationIssue[] | undefined) {
  if (errors) {
    let errorNumber = 1;
    for (const currentError of errors) {
      console.log(chalk.bold.red(`Error #${errorNumber}: `));
      logManifestValidationIssue(currentError);
      console.log();
      ++errorNumber;
    }
  }
}

function logManifestValidationInfos(infos: ManifestValidationIssue[] | undefined) {
  if (infos) {
    console.log(chalk.bold.blue(`Validation Information: `));
    for (const currentInfo of infos) {
      logManifestValidationIssue(currentInfo);
      console.log();
    }
  }
}

function logManifestValidationWarnings(warnings: ManifestValidationIssue[] | undefined) {
  if (warnings) {
    let warningNumber = 1;
    for (const currentWarning of warnings) {
      console.log(chalk.bold.yellow(`Warning #${warningNumber}: `));
      logManifestValidationIssue(currentWarning);
      console.log();
      ++warningNumber;
    }
  }
}

function logManifestValidationIssue(issue: ManifestValidationIssue) {
  console.log(`${issue.title}: ${issue.content}` + (issue.helpUrl ? ` (link: ${issue.helpUrl})` : ``));

  if (issue.code) {
    console.log(`  - Details: ${issue.code}`);
  }
  if (issue.line) {
    console.log(`  - Line: ${issue.line}`);
  }
  if (issue.column) {
    console.log(`  - Column: ${issue.column}`);
  }
}

function logManifestValidationSupportedProducts(products: ManifestValidationProduct[] | undefined) {
  if (products) {
    const productTitles = new Set(products.filter((product) => product.title).map((product) => product.title));

    if (productTitles.size > 0) {
      console.log(
        `\nBased on the requirements specified in your manifest, your add-in can run on the following platforms; your add-in will be tested on these platforms when you submit it to the Office Store:`
      );
      for (const productTitle of productTitles) {
        console.log(`  - ${productTitle}`);
      }
      console.log(
        `Important: This analysis is based on the requirements specified in your manifest and does not account for any runtime JavaScript calls within your add-in. For information about which API sets and features are supported on each platform, see Office Add-in host and platform availability. (https://docs.microsoft.com/office/dev/add-ins/overview/office-add-in-availability).\n`
      );
      console.log(
        `*This does not include mobile apps. You can opt-in to support mobile apps when you submit your add-in.`
      );
    }
  }
}

export async function modify(manifestPath: string, command: commander.Command) {
  try {
    // if the --guid command option is provided without a value, use "" to specify to change to a random guid value.
    const guid: string | undefined = getCommandOptionString(command.guid, "");
    const displayName: string | undefined = getCommandOptionString(command.displayName);

    const manifest = await OfficeAddinManifest.modifyManifestFile(manifestPath, guid, displayName);
    logManifestInfo(manifestPath, manifest);
    usageDataObject.reportSuccess("modify");
  } catch (err: any) {
    usageDataObject.reportException("modify", err);
    logErrorMessage(err);
  }
}

export async function validate(
  manifestPath: string,
  command: commander.Command /* eslint-disable-line @typescript-eslint/no-unused-vars */
) {
  try {
    const verifyProduction: boolean = command.production;
    const validation: ManifestValidation = await validateManifest(manifestPath, verifyProduction);
    if (validation.status && validation.status != 200) {
      console.log(`Unable to validate the manifest.\n${validation.status}\n${validation.statusText}`);
    } else if (validation.report) {
      logManifestValidationInfos(validation.report.notes);
      logManifestValidationErrors(validation.report.errors);
      logManifestValidationWarnings(validation.report.warnings);

      if (validation.isValid) {
        if (validation.report.addInDetails) {
          logManifestValidationSupportedProducts(validation.report.addInDetails.supportedProducts);
          console.log();
        }
        console.log(chalk.bold.green("The manifest is valid.\n"));
      } else {
        console.log(chalk.bold.red("The manifest is not valid.\n"));
      }
    }

    process.exitCode = validation.isValid ? 0 : 1;
    usageDataObject.reportSuccess("validate");
  } catch (err: any) {
    usageDataObject.reportException("validate", err);
    logErrorMessage(err);
  }
}

export async function exportManifest(command: commander.Command) {
  try {
    const outputPath: string = command.output ?? "";
    const manifestPath: string = command.manifest ?? "./manifest.json";

    await exportMetadataPackage(outputPath, manifestPath);
    usageDataObject.reportSuccess("export");
  } catch (err: any) {
    usageDataObject.reportException("export", err);
    logErrorMessage(err);
  }
}
