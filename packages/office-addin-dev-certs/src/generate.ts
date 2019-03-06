import {installCertificates} from "./install";
import {verifyCertificates} from "./verify";

export function generateCertificates(caCertPath: string, certPath: string , keyPath: string, overwriteCert: any, installCert: any): void {
    if (!overwriteCert) {
        const isCertificateInstalled = verifyCertificates();
        if (isCertificateInstalled) {
            console.log("Valid certificate already exists. Run with --overwrite to overwrite the existing certificates");
            return;
        }
    }
    const createCA = require("mkcert").createCA;
    const createCert = require("mkcert").createCert;
    const fs = require("fs");
    createCA({
        countryCode: "US",
        locality: "Redmond",
        organization: "Developer CA for Microsoft Office Add-ins",
        state: "WA",
        validityDays: 30,
    })
    .then((ca: any) => {
        createCert({
            caCert: ca,
            caKey: ca,
            domains: ["127.0.0.1", "localhost"],
            validityDays: 30,
        })
        .then((localhost: any) => {
            fs.writeFileSync(`${caCertPath}`, ca.cert);
            fs.writeFileSync(`${certPath}`, localhost.cert);
            fs.writeFileSync(`${keyPath}`, localhost.key);
            console.log("Certificates is generated successfully");
            if (installCert) {
                installCertificates(caCertPath);
            }
        })
        .catch((err: any) => console.error(err));
    })
    .catch((err: any) => console.error(err));
}
