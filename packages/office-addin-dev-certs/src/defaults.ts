import * as os from "os";
import * as path from "path";

// Default certificate names
export const userHomeDirectory = os.homedir();
export const certificateDirectoryName = ".office-addin-dev-certs";
export const certificateDirectory =  path.join(userHomeDirectory, certificateDirectoryName);
export const caCertificateFileName = "ca.crt";
export const caCertificatePath = path.join(certificateDirectory, caCertificateFileName);
export const localhostCertificateFileName = "localhost.crt";
export const localhostCertificatePath = path.join(certificateDirectory, localhostCertificateFileName);
export const localhostKeyFileName = "localhost.key";
export const localhostKeyPath = path.join(certificateDirectory, localhostKeyFileName);

// Default certificate details
export const certificateName = "Developer CA for Microsoft Office Add-ins";
export const countryCode = "US";
export const daysUntilCertificateExpires = 30;
export const domain = ["127.0.0.1", "localhost"];
export const locality = "Redmond";
export const state = "WA";
