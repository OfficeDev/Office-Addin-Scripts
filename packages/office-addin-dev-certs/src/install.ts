import * as fs from "fs";
import * as fsExtra from "fs-extra";
import * as mkcert from "mkcert";
import * as path from "path";
import * as defaults from "./defaults";
import * as verify from "./verify";

async function generateCertificates(caCertificatePath: string, localhostCertificatePath: string, localhostKeyPath: string,  daysUntilCertificateExpires: number) {
    try {
        fsExtra.ensureDirSync(path.dirname(caCertificatePath));
        fsExtra.ensureDirSync(path.dirname(localhostCertificatePath));
        fsExtra.ensureDirSync(path.dirname(localhostKeyPath));
    } catch (err) {
        throw new Error(`Unable to create directory\n${err}`);
    }

    const cACertificateInfo: mkcert.CACertificateInfo = {
                                                            countryCode: defaults.countryCode,
                                                            locality: defaults.locality,
                                                            organization: defaults.certificateName,
                                                            state: defaults.state,
                                                            validityDays: daysUntilCertificateExpires,
                                                        };
    let caCertificate: mkcert.Certificate;
    try {
        caCertificate = await mkcert.createCA(cACertificateInfo);
    } catch (err) {
        throw new Error(`Unable to generate CA certificate\n${err}`);
    }

    const localhostCertificateInfo: mkcert.CertificateInfo = {
                                                                caCert: caCertificate.cert,
                                                                caKey: caCertificate.key,
                                                                domains: defaults.domain,
                                                                validityDays: daysUntilCertificateExpires,
                                                            };
    let localhostCertificate: mkcert.Certificate;
    try {
        localhostCertificate = await mkcert.createCert(localhostCertificateInfo);
    } catch (err) {
        throw new Error(`Unable to generate localhost certificate\n${err}`);
    }

    try {
        fs.writeFileSync(`${caCertificatePath}`, caCertificate.cert);
        fs.writeFileSync(`${localhostCertificatePath}`, localhostCertificate.cert);
        fs.writeFileSync(`${localhostKeyPath}`, localhostCertificate.key);
    } catch (err) {
        throw new Error(`Unable to write generated certificates\n${err}`);
    }

    if (caCertificatePath === defaults.caCertificatePath) {
        console.log("The developer certificates have been generated in " + process.cwd());
    } else {
        console.log("The developer certificates have been generated");
    }
}

async function installCACertificate(caCertPath: string) {
    let command = "powershell Import-Certificate -CertStoreLocation cert:\\CurrentUser\\Root ";
    if (process.platform !== "win32") {
       command = "sudo security add-trusted-cert -d -r trustRoot -k /Library/Keychains/System.keychain";
    }
    const finalCommand = `${command} ${caCertPath}`;
    const execSync = require("child_process").execSync;
    try {
       console.log(`Installing CA certificate "Developer CA for Microsoft Office Add-ins"`);
       execSync(finalCommand, {stdio : "pipe" });
       console.log("Successfully installed certificate to trusted store");
    } catch (error) {
       throw new Error(`Failed to install CA certificate\n${error.stdout.toString()}`);
    }
 }

/* Install operation will check if there is already valid certificate installed.
   if yes, then this operation will be no op.
   else, new certificates are generated and installed.
*/
export async function generateAndInstallCertificate(caCertificatePath: string = defaults.caCertificatePath, localhostCertificatePath: string = defaults.localhostCertificatePath, localhostKeyPath: string = defaults.localhostKeyPath,
                                        daysUntilCertificateExpires: number = defaults.daysUntilCertificateExpires) {

    const isCertificateValid  = verify.verifyCaCertificate();
    if (isCertificateValid) {
        console.log("A valid CA certificate is already installed.");
        return;
    }

    await generateCertificates(caCertificatePath, localhostCertificatePath, localhostKeyPath, daysUntilCertificateExpires);

    await installCACertificate(caCertificatePath);
}
