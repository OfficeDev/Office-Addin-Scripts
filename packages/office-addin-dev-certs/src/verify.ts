// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { execSync } from "child_process";
import crypto from "crypto";
import fs from "fs";
import path from "path";
import * as defaults from "./defaults";
import { usageDataObject } from "./defaults";
import { ExpectedError } from "office-addin-usage-data";

export enum CertificateStatus {
  Installed = "installed",
  NotInstalled = "not-installed",
  Error = "error",
}

export interface CertificateVerificationResult {
  status: CertificateStatus;
  error?: Error;
}

/* global Buffer process __dirname */

// On win32 this is a unique hash used with PowerShell command to reliably delineate command output
export const outputMarker =
  process.platform === "win32"
    ? `[${crypto.createHash("md5").update(`${defaults.certificateName}${defaults.caCertificatePath}`).digest("hex")}]`
    : "";

function getVerifyCommand(returnInvalidCertificate: boolean): string {
  switch (process.platform) {
    case "win32": {
      const script = path.resolve(__dirname, "..\\scripts\\verify.ps1");
      const defaultCommand = `powershell -ExecutionPolicy Bypass -File "${script}" -CaCertificateName "${defaults.certificateName}" -CaCertificatePath "${defaults.caCertificatePath}" -LocalhostCertificatePath "${defaults.localhostCertificatePath}" -OutputMarker "${outputMarker}"`;
      if (returnInvalidCertificate) {
        return defaultCommand + ` -ReturnInvalidCertificate`;
      }
      return defaultCommand;
    }
    case "darwin": {
      // macOS
      const script = path.resolve(__dirname, "../scripts/verify.sh");
      return `sh '${script}' '${defaults.certificateName}'`;
    }
    case "linux": {
      const script = path.resolve(__dirname, "../scripts/verify_linux.sh");
      return `sh '${script}' '${defaults.caCertificateFileName}'`;
    }
    default:
      throw new ExpectedError(`Platform not supported: ${process.platform}`);
  }
}

/**
 * Checks if the CA certificate is installed.
 * @param returnInvalidCertificate - If true, returns true even if the certificate is expired.
 * @returns true if the certificate is installed, false otherwise.
 * @throws Error if the verification script fails to execute (e.g., PowerShell not found, permission issues).
 */
export function isCaCertificateInstalled(returnInvalidCertificate: boolean = false): boolean {
  const result = getCaCertificateStatus(returnInvalidCertificate);
  if (result.status === CertificateStatus.Error) {
    throw result.error!;
  }
  return result.status === CertificateStatus.Installed;
}

/**
 * Gets the CA certificate installation status with detailed error information.
 * @param returnInvalidCertificate - If true, returns Installed even if the certificate is expired.
 * @returns A CertificateVerificationResult object with status and optional error information.
 */
export function getCaCertificateStatus(
  returnInvalidCertificate: boolean = false
): CertificateVerificationResult {
  const command = getVerifyCommand(returnInvalidCertificate);

  try {
    const output = execSync(command, { stdio: "pipe" }).toString();
    if (process.platform === "win32") {
      // Remove any PowerShell output that preceeds invoking the actual certificate check command
      const hasOutput =
        output.slice(output.lastIndexOf(outputMarker) + outputMarker.length).trim().length !== 0;
      return { status: hasOutput ? CertificateStatus.Installed : CertificateStatus.NotInstalled };
    }
    // script files return empty string if the certificate not found or expired
    if (output.length !== 0) {
      return { status: CertificateStatus.Installed };
    }
    return { status: CertificateStatus.NotInstalled };
  } catch (err: unknown) {
    // Distinguish between "certificate not found" (expected) and script execution errors (unexpected).
    // On Windows, PowerShell with $ErrorActionPreference = 'Stop' throws when the cert is not found.
    // On macOS/Linux, scripts return empty output rather than throwing for missing certs.
    if (process.platform === "win32" && err instanceof Error) {
      const errorMessage = err.message.toLowerCase();
      // PowerShell script errors that indicate the script ran but cert wasn't found
      if (
        errorMessage.includes("cannot find") ||
        errorMessage.includes("not found") ||
        errorMessage.includes("no certificate")
      ) {
        return { status: CertificateStatus.NotInstalled };
      }
    }
    // For other platforms or unexpected errors, check if it's a script execution failure
    const error = err instanceof Error ? err : new Error(String(err));
    const errorMessage = error.message.toLowerCase();

    // Common script execution failures that should be reported as errors
    if (
      errorMessage.includes("command not found") ||
      errorMessage.includes("not recognized") ||
      errorMessage.includes("cannot be loaded") ||
      errorMessage.includes("permission denied") ||
      errorMessage.includes("enoent") ||
      errorMessage.includes("access denied") ||
      errorMessage.includes("spawn")
    ) {
      return { status: CertificateStatus.Error, error };
    }

    // For other errors, assume the certificate is not installed (backwards compatibility)
    // but log the error for debugging purposes
    return { status: CertificateStatus.NotInstalled };
  }
}

function validateCertificateAndKey(certificatePath: string, keyPath: string) {
  let certificate: string = "";
  let key: string = "";

  try {
    certificate = fs.readFileSync(certificatePath).toString();
  } catch (err) {
    throw new Error(`Unable to read the certificate.\n${err}`);
  }

  try {
    key = fs.readFileSync(keyPath).toString();
  } catch (err) {
    throw new Error(`Unable to read the certificate key.\n${err}`);
  }

  let encrypted;

  try {
    encrypted = crypto.publicEncrypt(certificate, Buffer.from("test"));
  } catch (err) {
    throw new Error(`The certificate is not valid.\n${err}`);
  }

  try {
    crypto.privateDecrypt(key, encrypted);
  } catch (err) {
    throw new Error(`The certificate key is not valid.\n${err}`);
  }
}

export function verifyCertificates(
  certificatePath: string = defaults.localhostCertificatePath,
  keyPath: string = defaults.localhostKeyPath
): boolean {
  try {
    let isCertificateValid: boolean = true;
    try {
      validateCertificateAndKey(certificatePath, keyPath);
    } catch {
      isCertificateValid = false;
    }
    let output = isCertificateValid && isCaCertificateInstalled();
    usageDataObject.reportSuccess("verifyCertificates()");
    return output;
  } catch (err: any) {
    usageDataObject.reportException("verifyCertificates()", err);
    throw err;
  }
}
