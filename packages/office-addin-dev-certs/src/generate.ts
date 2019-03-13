import * as mkcert from "mkcert";
import * as defaults from "./defaults";
import {installCaCertificate} from "./install";
import {verifyCaCertificate} from "./verify";

/* Generate operation will check if there is already valid certificate installed.
   if yes, then this operation will be no op.
   else, new certificates are generated and installed if --install was provided.
*/
export async function generateCertificates(caCertificatePath: string = defaults.caCertificatePath, certPath: string = defaults.localhostCertificatePath, keyPath: string = defaults.localhostKeyPath, install: boolean = false) {

    const isCertificateValid  = verifyCaCertificate();
    if (isCertificateValid) {
        console.log("A valid CA certificate is already installed.");
        return;
    }

    const fs = require("fs-path");
    try {
        const caCertificate= await mkcert.createCA({
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
        fs.writeFileSync(`${certPath}`, localhostCertificate.cert);
        fs.writeFileSync(`${keyPath}`, localhostCertificate.key);
        if (caCertificatePath === defaults.caCertificatePath) {
            console.log("The developer certificates have been generated in " + process.cwd());
        } else {
            console.log("The developer certificates have been generated");
        }
        if (install) {
            installCaCertificate(caCertificatePath);
        }
     } catch (err) {
        console.error(err);
     }
}
