import * as crypto from "crypto";
import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import * as defaults from "./defaults";
import { generateCertificates } from "./generate";
import { installCaCertificate } from "./install";
import { uninstallCaCertificate } from "./uninstall";
import { isCaCertificateInstalled } from "./verify";
import {logErrorMessage } from "office-addin-cli";
interface IHttpsServerOptions {
    cert: Buffer;
    key: Buffer;
}

function getCertificateDirectory(): string {
    const userHomeDirectory = os.homedir();
    return path.join(userHomeDirectory, defaults.certificateDirectory);
}

function validateCertificateAndKey(certificate: string, key: string) {
    let encrypted;
    try {
        encrypted = crypto.publicEncrypt(certificate, Buffer.from("test"));
    } catch (err) {
        throw new Error(`The certificate is not valid.\n${err}`);
    }

    try {
        crypto.privateDecrypt(key, encrypted);
    } catch (err) {
        throw new Error(`The localhost certificate key is invalid.\n${err}`);
    }
}

async function generateAndInstallCertificate(caCertificatePath: string, localhostCertificatePath: string, localhostKeyPath: string) {
    await generateCertificates(caCertificatePath, localhostCertificatePath, localhostKeyPath);
    await uninstallCaCertificate();
    await installCaCertificate(caCertificatePath);
}

export async function gethttpsServerOptions(): Promise<IHttpsServerOptions> {
    const certificateDirectory = getCertificateDirectory();
    const caCertificatePath = path.join(certificateDirectory, defaults.caCertificateFileName);
    const localhostCertificatePath = path.join(certificateDirectory, defaults.localhostCertificateFileName);
    const localhostKeyPath = path.join(certificateDirectory, defaults.localhostKeyFileName);

    let localhostCertificate: Buffer = Buffer.alloc(0);
    let localhostKey: Buffer = Buffer.alloc(0);
    let needToGenerateCertificates: boolean = false;

    try {
        localhostCertificate = fs.readFileSync(localhostCertificatePath);
        localhostKey = fs.readFileSync(localhostKeyPath);
        validateCertificateAndKey(localhostCertificate.toString(), localhostKey.toString());
        isCaCertificateInstalled();
    } catch (err) {
        logErrorMessage(err);
        needToGenerateCertificates = true;
    }

    if (needToGenerateCertificates) {
        await generateAndInstallCertificate(caCertificatePath, localhostCertificatePath, localhostKeyPath);
        localhostCertificate = fs.readFileSync(localhostCertificatePath);
        localhostKey = fs.readFileSync(localhostKeyPath);
    }

    const httpsServerOptions = {} as IHttpsServerOptions;
    try {
        httpsServerOptions.cert = localhostCertificate;
        httpsServerOptions.key = localhostKey;
    } catch (err) {
        throw new Error(`Error occured while reading certificate files.\n${err}`);
    }

    return httpsServerOptions;
}
