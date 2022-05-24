// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import { parseNumber } from "office-addin-cli";
import { logErrorMessage } from "office-addin-usage-data";
import * as defaults from "./defaults";
import { ensureCertificatesAreInstalled } from "./install";
import { deleteCertificateFiles, uninstallCaCertificate } from "./uninstall";
import { verifyCertificates } from "./verify";
import { usageDataObject } from "./defaults";
import { ExpectedError } from "office-addin-usage-data";

/* global console */

function parseDays(optionValue: any): number | undefined {
  const days = parseNumber(optionValue, "--days should specify a number.");

  if (days !== undefined) {
    if (!Number.isInteger(days)) {
      throw new ExpectedError("--days should be integer.");
    }
    if (days <= 0) {
      throw new ExpectedError("--days should be greater than zero.");
    }
  }
  return days;
}

export async function install(command: commander.Command) {
  try {
    const days = parseDays(command.days);

    await ensureCertificatesAreInstalled(days, command.machine);
    usageDataObject.reportSuccess("install");
  } catch (err: any) {
    usageDataObject.reportException("install", err);
    logErrorMessage(err);
  }
}

export async function verify(command: commander.Command /* eslint-disable-line @typescript-eslint/no-unused-vars */) {
  try {
    if (await verifyCertificates()) {
      console.log(
        `You have trusted access to https://localhost.\nCertificate: ${defaults.localhostCertificatePath}\nKey: ${defaults.localhostKeyPath}`
      );
    } else {
      console.log(`You need to install certificates for trusted access to https://localhost.`);
    }
    usageDataObject.reportSuccess("verify");
  } catch (err: any) {
    usageDataObject.reportException("verify", err);
    logErrorMessage(err);
  }
}

export async function uninstall(command: commander.Command) {
  try {
    await uninstallCaCertificate(command.machine);
    deleteCertificateFiles(defaults.certificateDirectory);
    usageDataObject.reportSuccess("uninstall");
  } catch (err: any) {
    usageDataObject.reportException("uninstall", err);
    logErrorMessage(err);
  }
}
