// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { execSync } from "child_process";
import * as path from "path";
import * as defaults from "./defaults";
import { generateCertificates } from "./generate";
import { deleteCertificateFiles, uninstallCaCertificate } from "./uninstall";
import { isCaCertificateInstalled, verifyCertificates } from "./verify";
import { usageDataObject } from "./defaults";
import { ExpectedError } from "office-addin-usage-data";

/* global process, console, __dirname */

function getInstallCommand(caCertificatePath: string, machine: boolean = false): string {
  switch (process.platform) {
    case "win32": {
      const script = path.resolve(__dirname, "..\\scripts\\install.ps1");
      return `powershell -ExecutionPolicy Bypass -File "${script}" ${
        machine ? "LocalMachine" : "CurrentUser"
      } "${caCertificatePath}"`;
    }
    case "darwin": // macOS
      const prefix = machine ? "sudo " : ""
      const keychainFile = machine ? "/Library/Keychains/System.keychain" : "~/Library/Keychains/login.keychain-db"
      return `${prefix}security add-trusted-cert -d -r trustRoot -k ${keychainFile} '${caCertificatePath}'`;
    case "linux":
      const script = path.resolve(__dirname, "../scripts/install_linux.sh");
      return `sudo sh '${script}' '${caCertificatePath}'`;
    default:
      throw new ExpectedError(`Platform not supported: ${process.platform}`);
  }
}

export async function ensureCertificatesAreInstalled(
  daysUntilCertificateExpires: number = defaults.daysUntilCertificateExpires,
  domains: string[] = defaults.domain,
  machine: boolean = false
) {
  try {
    const areCertificatesValid = verifyCertificates();

    if (areCertificatesValid) {
      console.log(
        `You already have trusted access to https://localhost.\nCertificate: ${defaults.localhostCertificatePath}\nKey: ${defaults.localhostKeyPath}`
      );
    } else {
      await uninstallCaCertificate(false, false);
      deleteCertificateFiles(defaults.certificateDirectory);
      await generateCertificates(
        defaults.caCertificatePath,
        defaults.localhostCertificatePath,
        defaults.localhostKeyPath,
        daysUntilCertificateExpires,
        domains
      );
      await installCaCertificate(defaults.caCertificatePath, machine);
    }

    usageDataObject.reportSuccess("ensureCertificatesAreInstalled()");
  } catch (err: any) {
    usageDataObject.reportException("ensureCertificatesAreInstalled()", err);
    throw err;
  }
}

export async function installCaCertificate(
  caCertificatePath: string = defaults.caCertificatePath,
  machine: boolean = false
) {
  const command = getInstallCommand(caCertificatePath, machine);

  try {
    console.log(`Installing CA certificate "Developer CA for Microsoft Office Add-ins"...`);
    // If the certificate is already installed by another instance skip it.
    if (!isCaCertificateInstalled()) {
      execSync(command, { stdio: "pipe" });
    }
    console.log(
      `You now have trusted access to https://localhost.\nCertificate: ${defaults.localhostCertificatePath}\nKey: ${defaults.localhostKeyPath}`
    );
  } catch (error: any) {
    throw new Error(`Unable to install the CA certificate. ${error.stderr.toString()}`);
  }
}
