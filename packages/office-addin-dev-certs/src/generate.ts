import {certificateName, certificateValidity, defaultCaCertPath, defaultCertPath, defaultKeyPath} from "./default";
import {installCaCertificate} from "./install";
import {verifyCaCertificate} from "./verify";

/* Generate operation will check if there is already valid certificate installed.
   if yes, then this operation will be no op.
   else, new certificates are generated and installed if --install was provided.
*/
export function generateCertificates(caCertPath: string | undefined, certPath: string | undefined, keyPath: string | undefined, install: boolean = false): void {
    if (!caCertPath) { caCertPath = defaultCaCertPath; }
    if (!certPath) { certPath = defaultCertPath; }
    if (!keyPath) { keyPath = defaultKeyPath; }
    interface ICertificateInfo {
        cert: string;
        key: string;
    }

    const isCertificateInstalled = verifyCaCertificate();
    if (isCertificateInstalled) {
        console.log("A valid CA certificate already exists in trusted store.");
        return;
    }

    const createCA = require("mkcert").createCA;
    const createCert = require("mkcert").createCert;
    const fs = require("fs");
    createCA({
        countryCode: "US",
        locality: "Redmond",
        organization: certificateName,
        state: "WA",
        validityDays: certificateValidity,
    })
    .then((caCertificateInfo: ICertificateInfo) => {
        createCert({
            caCert: caCertificateInfo.cert,
            caKey: caCertificateInfo.key,
            domains: ["127.0.0.1", "localhost"],
            validityDays: certificateValidity,
        })
        .then((localhost: ICertificateInfo) => {
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
