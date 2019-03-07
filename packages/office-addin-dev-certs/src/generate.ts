import {installCaCertificate} from "./install";
import {verifyCaCertificate} from "./verify";

export const defaultCaCertPath = "ca.crt"
export const defaultCertPath = "localhost.crt"
export const defaultKeyPath = "localhost.key"

export function generateCertificates(caCertPath: string | undefined, certPath: string | undefined, keyPath: string | undefined, overwriteCert: any, installCert: any): void {
    if(!caCertPath) caCertPath = defaultCaCertPath;
    if(!certPath) certPath = defaultCertPath;
    if(!keyPath) keyPath = defaultKeyPath;
    interface certificateInfo{
        cert: string;
        key: string;
    }
    if (!overwriteCert) {
        const isCertificateInstalled = verifyCaCertificate();
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
    .then((caCertificateInfo: certificateInfo) => {
        createCert({
            caCert: caCertificateInfo.cert,
            caKey: caCertificateInfo.key,
            domains: ["127.0.0.1", "localhost"],
            validityDays: 30,
        })
        .then((localhost: certificateInfo) => {
            fs.writeFileSync(`${caCertPath}`, caCertificateInfo.cert);
            fs.writeFileSync(`${certPath}`, localhost.cert);
            fs.writeFileSync(`${keyPath}`, localhost.key);
            console.log("The developer certificates have been generated.");
            if (installCert) {
                installCaCertificate(caCertPath);
            }
        })
        .catch((err: any) => console.error(err));
    })
    .catch((err: any) => console.error(err));
}
