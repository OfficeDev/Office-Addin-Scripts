// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { execSync } from "child_process";
import * as crypto from "crypto";
import * as fs from "fs";
import * as path from "path";
import * as defaults from "./defaults";
import { deleteCertificateFiles, uninstallCaCertificate } from './uninstall'

function getVerifyCommand(): string {
    switch (process.platform) {
       case "win32":
          const script = path.resolve(__dirname, "..\\scripts\\verify.ps1");
          return `powershell -ExecutionPolicy Bypass -File "${script}" "${defaults.certificateName}"`;
       case "darwin": // macOS
          return `security find-certificate -c '${defaults.certificateName}' -p | openssl x509 -noout`;
       default:
          throw new Error(`Platform not supported: ${process.platform}`);
    }
}

export function isCaCertificateInstalled(): boolean {
    const command = getVerifyCommand();

    try {
        const output = execSync(command, {stdio : "pipe" }).toString();
        if (process.platform === "darwin") {
            // Ceritificate exists. Now check and see if it's expired and remove it if is.
            try {
                const command = `security find-certificate -c '${defaults.certificateName}' -p | openssl x509 -checkend 86400 -noout`;
                execSync(command, {stdio : "pipe" }).toString();
                return true;
            }
            catch {
                uninstallCaCertificate(false /* machine */, true /* verbose */, true /* expiredCert */);
                deleteCertificateFiles(defaults.certificateDirectory);
                return false;
            }

        } else if (output.length !== 0) {
            return true; // powershell command return empty string if the certificate not-found/expired
        }
    } catch (error) {
        // Mac security command throws error if the certifcate is not-found/expired
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

export function verifyCertificates(certificatePath: string = defaults.localhostCertificatePath, keyPath: string = defaults.localhostKeyPath): boolean {
    let isCertificateValid: boolean = true;
    try {
        validateCertificateAndKey(certificatePath, keyPath);
    } catch (err) {
        isCertificateValid = false;
    }
    return isCertificateValid && isCaCertificateInstalled();
}
