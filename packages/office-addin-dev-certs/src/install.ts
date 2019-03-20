import { execSync } from "child_process";
import * as defaults from "./defaults";
import { generateCertificates } from "./generate";
import { uninstallCaCertificate } from "./uninstall";
import { verifyCertificates } from "./verify";

function getInstallCommand(caCertificatePath: string): string {
   let command: string;
   switch (process.platform) {
      case "win32":
         command = `powershell Import-Certificate -CertStoreLocation cert:\\CurrentUser\\Root ${caCertificatePath}`;
         break;
      case "darwin": // macOS
         command = `sudo security add-trusted-cert -d -r trustRoot -k /Library/Keychains/System.keychain ${caCertificatePath}`;
         break;
      default:
         throw new Error(`Platform not supported: ${process.platform}`);
   }
   return command;
}

export async function ensureCertificatesAreInstalled(caCertificatePath: string = defaults.caCertificatePath,
   localhostCertificatePath: string = defaults.localhostCertificatePath,
   localhostKeyPath: string = defaults.localhostKeyPath,
   daysUntilCertificateExpires: number = defaults.daysUntilCertificateExpires) {

   const areCertificatesValid = verifyCertificates();

   if (!areCertificatesValid) {
      await generateCertificates();
      await uninstallCaCertificate();
      await installCaCertificate();
   }
}

export async function installCaCertificate(caCertificatePath: string = defaults.caCertificatePath) {
    const command = getInstallCommand(caCertificatePath);

    try {
        console.log(`Installing CA certificate "Developer CA for Microsoft Office Add-ins"...`);
        execSync(command, {stdio : "pipe" });
        console.log(`The CA certificate was installed.`);
    } catch (error) {
        throw new Error(`Unable to install the CA certificate. ${error.stderr.toString()}`);
    }
}
