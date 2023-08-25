// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { execSync } from "child_process";
import * as crypto from "crypto";
import * as fs from "fs";
import * as path from "path";
import * as defaults from "./defaults";
import { usageDataObject } from "./defaults";
import { ExpectedError } from "office-addin-usage-data";

// On win32 this is a unique hash used with PowerShell command to reliably delineate command output
export const outputMarker = process.platform === "win32" ? `[${crypto.createHash("md5").update(`${defaults.certificateName}${defaults.caCertificatePath}`).digest("hex")}]` : "";

/* global process, Buffer, __dirname */

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
    case "linux":
      const script = path.resolve(__dirname, "../scripts/verify_linux.sh");
      return `sh '${script}' '${defaults.certificateName}'`;
    default:
      throw new ExpectedError(`Platform not supported: ${process.platform}`);
  }
}

export function isCaCertificateInstalled(returnInvalidCertificate: boolean = false): boolean {
  const command = getVerifyCommand(returnInvalidCertificate);

  try {
    const output = execSync(command, { stdio: "pipe" }).toString();
    if (process.platform === "win32") {
      // Remove any PowerShell output that preceeds invoking the actual certificate check command
      return output.slice(output.lastIndexOf(outputMarker) + outputMarker.length).trim().length !== 0;
    }
    // script files return empty string if the certificate not found or expired
    if (output.length !== 0) {
      return true;
    }
  } catch (error) {
    // Some commands throw errors if the certifcate is not found or expired
  }

  return false;
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
    } catch (err) {
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
