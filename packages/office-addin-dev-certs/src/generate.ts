import * as fs from "fs";
import * as mkcert from "mkcert";
import * as mkdirp from "mkdirp";
import * as path from "path";
import * as defaults from "./defaults";
import * as install from "./install";
import * as verify from "./verify";

function createDirectoryIfNotExists(certificatePath: string) {
    const certificateDir = path.dirname(certificatePath);
    if (!fs.existsSync(certificateDir)) {
        mkdirp(certificateDir, function(err) {
        if (err) {
            console.log("Failed to create directory " + err);
        } else {
            console.log("Sucessfully created a directory " + certificateDir);
        }
        });
    }
  }

/* Generate operation will check if there is already valid certificate installed.
   if yes, then this operation will be no op.
   else, new certificates are generated and installed if --install was provided.
*/
export async function generateCertificates(caCertificatePath: string = defaults.caCertificatePath, localhostCertificatePath: string = defaults.localhostCertificatePath, localhostKeyPath: string = defaults.localhostKeyPath, installCaCertificate: boolean = false) {

    const isCertificateValid  = verify.verifyCaCertificate();
    if (isCertificateValid) {
        console.log("A valid CA certificate is already installed.");
        return;
    }
    createDirectoryIfNotExists(caCertificatePath);
    createDirectoryIfNotExists(localhostCertificatePath);
    createDirectoryIfNotExists(localhostKeyPath);
    const caCertificate = await mkcert.createCA({
                            countryCode: defaults.countryCode,
                            locality: defaults.locality,
                            organization: defaults.certificateName,
                            state: defaults.state,
                            validityDays: defaults.daysUntilCertificateExpires,
                        });
    const localhostCertificate = await mkcert.createCert({
                                        caCert: caCertificate.cert,
                                        caKey: caCertificate.key,
                                        domains: defaults.domain,
                                        validityDays: defaults.daysUntilCertificateExpires,
                                    });
    fs.writeFileSync(`${caCertificatePath}`, caCertificate.cert);
    fs.writeFileSync(`${localhostCertificatePath}`, localhostCertificate.cert);
    fs.writeFileSync(`${localhostKeyPath}`, localhostCertificate.key);
    if (caCertificatePath === defaults.caCertificatePath) {
        console.log("The developer certificates have been generated in " + process.cwd());
    } else {
        console.log("The developer certificates have been generated");
    }
    if (installCaCertificate) {
        install.installCaCertificate(caCertificatePath);
    }
}
