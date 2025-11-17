// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import path from "path";
import {
  connectionStringForOfficeAddinCLITools,
  OfficeAddinUsageData,
} from "office-addin-usage-data";

/* global process __dirname */

// File path defaults
export const azCliInstallCommandPath: string = path.resolve(
  `${__dirname}/scripts/azCliInstallCmd.ps1`
);
export const azRestAddTenantOutlookReplyUrlsCommandPath = path.resolve(
  `${__dirname}/scripts/azRestAddTenantOutlookReplyUrls.txt`
);
export const azRestAddTenantReplyUrlsCommandPath = path.resolve(
  `${__dirname}/scripts/azRestAddTenantReplyUrls.txt`
);
export const azRestAddSecretCommandPath = path.resolve(`${__dirname}/scripts/azAddSecretCmd.txt`);
export const azRestAppCreateCommandPath: string = path.resolve(
  `${__dirname}/scripts/azRestAppCreateCmd.txt`
);
export const azRestGetOrganizationDetailsCommandPath: string = path.resolve(
  `${__dirname}/scripts/azGetOrganizationDetails.txt`
);
export const azRestGetTenantAdminMembershipCommandPath: string = path.resolve(
  `${__dirname}/scripts/azRestGetTenantAdminMembership.txt`
);
export const azRestGetTenantRolesPath: string = path.resolve(
  `${__dirname}/scripts/azRestGetTenantRoles.txt`
);
export const azRestSetIdentifierUriCommmandPath: string = path.resolve(
  `${__dirname}/scripts/azRestSetIdentifierUri.txt`
);
export const azRestSetSigninAudienceCommandPath: string = path.resolve(
  `${__dirname}/scripts/azSetSignInAudienceCmd.txt`
);
export const envDataFilePath = path.resolve(`${process.cwd()}/.ENV`);
export const fallbackAuthDialogTypescriptFilePath = path.resolve(
  `${process.cwd()}/src/helpers/fallbackAuthDialog.ts`
);
export const fallbackAuthDialogJavascriptFilePath = path.resolve(
  `${process.cwd()}/src/helpers/fallbackAuthDialog.js`
);
export const getInstalledAppsPath: string = path.resolve(
  `${__dirname}/scripts/getInstalledApps.ps1`
);
export const manifestFilePath = path.resolve(`${process.cwd()}/manifest.xml`);
export const addSecretCommandPath: string = path.resolve(`${__dirname}/scripts/addAppSecret.ps1`);
export const getSecretCommandPath: string = path.resolve(`${__dirname}/scripts/getAppSecret.ps1`);
export const testEnvDataFilePath = path.resolve(`${process.cwd()}/test/test-env`);
export const testFallbackAuthDialogFilePath: string = path.resolve(
  `${process.cwd()}/test/test-fallbackauthdialog`
);
export const testManifestFilePath = path.resolve(`${process.cwd()}/test/test-manifest.xml`);

// Usage data defaults
export const usageDataObject: OfficeAddinUsageData = new OfficeAddinUsageData({
  projectName: "office-addin-sso",
  connectionString: connectionStringForOfficeAddinCLITools,
  raisePrompt: false,
});
