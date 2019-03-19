import * as path from "path";

// Default certificate names
export const certificateDirectory = ".office-addin-dev-certs";
export const caCertificateFileName = "ca.crt";
export const caCertificatePath = path.join(".", caCertificateFileName);
export const localhostCertificateFileName = "localhost.crt";
export const localhostCertificatePath = path.join(".", localhostCertificateFileName);
export const localhostKeyFileName = "localhost.key";
export const localhostKeyPath = path.join(".", localhostKeyFileName);

// Default certificate details
export const certificateName = "Developer CA for Microsoft Office Add-ins";
export const countryCode = "US";
export const daysUntilCertificateExpires = 30;
export const domain = ["127.0.0.1", "localhost"];
export const locality = "Redmond";
export const state = "WA";
