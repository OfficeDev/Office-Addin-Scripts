import {Certificate, createCA, createCert} from "mkcert";
import * as defaults from "./default";
import {installCaCertificate} from "./install";
import {verifyCaCertificate} from "./verify";

/* Generate operation will check if there is already valid certificate installed.
   if yes, then this operation will be no op.
   else, new certificates are generated and installed if --install was provided.
*/
export function generateCertificates(caCertPath: string | undefined, certPath: string | undefined, keyPath: string | undefined, install: boolean = false): void {
    if (!caCertPath) { caCertPath = defaults.defaultCaCertPath; }
    if (!certPath) { certPath = defaults.defaultCertPath; }
    if (!keyPath) { keyPath = defaults.defaultKeyPath; }

    const isCertificateValid  = verifyCaCertificate();
    if (isCertificateValid) {
        console.log("A valid CA certificate is already installed.");
        return;
    }

    const fs = require("fs");
    createCA({
        countryCode: defaults.countryCode,
        locality: defaults.locality,
        organization: defaults.certificateName,
        state: defaults.state,
        validityDays: defaults.daysUntilCertificateExpires,
    })
    .then((caCertificateInfo: Certificate) => {
        createCert({
            caCert: caCertificateInfo.cert,
            caKey: caCertificateInfo.key,
            domains: defaults.domain,
            validityDays: defaults.daysUntilCertificateExpires,
        })
        .then((localhost: Certificate) => {
            fs.writeFileSync(`${caCertPath}`, caCertificateInfo.cert);
            fs.writeFileSync(`${certPath}`, localhost.cert);
            fs.writeFileSync(`${keyPath}`, localhost.key);
            console.log("The developer certificates have been generated.");
            if (install) {
                installCaCertificate(caCertPath);
            }
        })
        .catch((err: any) => console.error(err));
    })
    .catch((err: any) => console.error(err));
}
