import * as crypto from "crypto";
import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import * as defaults from "./defaults";
import { generateCertificates } from "./generate";
import { installCaCertificate } from "./install";
import { uninstallCaCertificate } from "./uninstall";
import { verifyCaCertificate } from "./verify";

interface IHttpsServerOptions {
    ca: Buffer;
    cert: Buffer;
    key: Buffer;
}

function getCertificateDirectory(): string {
    const userHomeDirectory = os.homedir();
    return path.join(userHomeDirectory, defaults.certificateDirectory);
}

function checkIfCertificatesExistsOnDisk(caCertificatePath: string, localhostCertificatePath: string, localhostKeyPath: string): boolean {
    const ret = fs.existsSync(caCertificatePath) && fs.existsSync(localhostCertificatePath) && fs.existsSync(localhostKeyPath);
    if (!ret) {
        console.log(`Certificate's not present on the disk.`);
    }
    return ret;
}

function validateKeyAndCerts(localhostCertificate: string, localhostKey: string): boolean {
    let encrypted;
    try {
        encrypted = crypto.publicEncrypt(localhostCertificate, Buffer.from("test"));
    } catch (err) {
        console.log(`The localhost certificate is invalid.`);
        return false;
    }

    try {
        crypto.privateDecrypt(localhostKey, encrypted);
    } catch (err) {
        console.log(`The localhost certificate key is invalid.`);
        return false;
    }
    return true;
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
    const certificatesExistsOnDisk = checkIfCertificatesExistsOnDisk(caCertificatePath, localhostCertificatePath, localhostKeyPath);
    const isCertificateInstalled  = verifyCaCertificate();
    let localhostCertificate: Buffer = Buffer.alloc(0);
    let localhostKey: Buffer = Buffer.alloc(0);
    let isCertificateValid: boolean;
    try {
        localhostCertificate = fs.readFileSync(localhostCertificatePath);
        localhostKey = fs.readFileSync(localhostKeyPath);
        isCertificateValid = validateKeyAndCerts(localhostCertificate.toString(), localhostKey.toString());
    } catch (err) {
        isCertificateValid = false;
    }

    let newCertificateGenerated = false;
    if (!isCertificateInstalled || !certificatesExistsOnDisk || !isCertificateValid) {
        await generateAndInstallCertificate(caCertificatePath, localhostCertificatePath, localhostKeyPath);
        newCertificateGenerated = true;
    }

    const httpsServerOptions = {} as IHttpsServerOptions;
    try {
        httpsServerOptions.ca = fs.readFileSync(caCertificatePath);
        httpsServerOptions.cert = newCertificateGenerated ? fs.readFileSync(localhostCertificatePath) : localhostCertificate;
        httpsServerOptions.key = newCertificateGenerated ? fs.readFileSync(localhostKeyPath) : localhostKey;
    } catch (err) {
        throw new Error(`Error occured while reading certificate files.\n${err}`);
    }

    return httpsServerOptions;
}
