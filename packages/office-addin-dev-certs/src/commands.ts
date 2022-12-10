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

export async function install(command: commander.Command) {
  try {
    const days = parseDays(command.days);
    const domains = parseDomains(command.domains);

    await ensureCertificatesAreInstalled(days, domains, command.machine);
    usageDataObject.reportSuccess("install");
  } catch (err: any) {
    usageDataObject.reportException("install", err);
    logErrorMessage(err);
  }
}

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

function parseDomains(optionValue: any): string[] | undefined {
  switch (typeof optionValue) {
    case "string": {
      try {
        return optionValue.split(",");
      } catch (err) {
        throw new Error("string value not in the correct format");
      }
    }
    case "undefined": {
      return undefined;
    }
    default: {
      throw new Error("--domains value should be a sting.");
    }
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
