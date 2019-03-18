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

export async function gethttpsServerOptions(): Promise<IHttpsServerOptions> {
    const certificateDirectory = getCertificateDirectory();
    const caCertificatePath = path.join(certificateDirectory, defaults.caCertificateFileName);
    const localhostCertificatePath = path.join(certificateDirectory, defaults.localhostCertificateFileName);
    const localhostKeyPath = path.join(certificateDirectory, defaults.localhostKeyFileName);

    const isCertificateValid  = verifyCaCertificate();
    if (!isCertificateValid) {
        await generateCertificates(caCertificatePath, localhostCertificatePath, localhostKeyPath);
        await uninstallCaCertificate();
        await installCaCertificate(caCertificatePath);
    }

    const httpsServerOptions = {} as IHttpsServerOptions;
    try {
        httpsServerOptions.ca = fs.readFileSync(caCertificatePath);
        httpsServerOptions.cert = fs.readFileSync(localhostCertificatePath);
        httpsServerOptions.key = fs.readFileSync(localhostKeyPath);
    } catch (err) {
        throw new Error(`Error occured while reading certificate files\n${err}`);
    }

    return httpsServerOptions;
}
