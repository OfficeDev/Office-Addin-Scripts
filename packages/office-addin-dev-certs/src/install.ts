import { execSync } from "child_process";
import * as defaults from "./defaults";
import { generateCertificates } from "./generate";
import { uninstallCaCertificate } from "./uninstall";
import { verifyCertificates } from "./verify";

function getInstallCommand(caCertificatePath: string, machine: boolean = false): string {
   let command: string;

   switch (process.platform) {
      case "win32":
         command = `powershell Import-Certificate -CertStoreLocation cert:\\${machine ? "LocalMachine" : "CurrentUser"}\\Root ${caCertificatePath}`;
         break;
      case "darwin": // macOS
         command = `sudo security add-trusted-cert -d -r trustRoot -k /Library/Keychains/System.keychain ${caCertificatePath}`;
         break;
      default:
         throw new Error(`Platform not supported: ${process.platform}`);
   }
   return command;
}

export async function ensureCertificatesAreInstalled(machine: boolean = false, daysUntilCertificateExpires: number = defaults.daysUntilCertificateExpires) {
   const areCertificatesValid = verifyCertificates();

   if (areCertificatesValid) {
      console.log(`You already have trusted access to https://localhost.\nCertificate: ${defaults.localhostCertificatePath}\nKey: ${defaults.localhostKeyPath}`);
   } else {
      await generateCertificates(defaults.caCertificatePath, defaults.localhostCertificatePath, defaults.localhostKeyPath, daysUntilCertificateExpires);
      await uninstallCaCertificate(false);
      await installCaCertificate(machine);
   }
}

export async function installCaCertificate(machine: boolean = false, caCertificatePath: string = defaults.caCertificatePath) {
    const command = getInstallCommand(caCertificatePath, machine);

    try {
        console.log(`Installing CA certificate "Developer CA for Microsoft Office Add-ins"...`);
        execSync(command, {stdio : "pipe" });
        console.log(`You now have trusted access to https://localhost.\nCertificate: ${defaults.localhostCertificatePath}\nKey: ${defaults.localhostKeyPath}`);
    } catch (error) {
        throw new Error(`Unable to install the CA certificate. ${error.stderr.toString()}`);
    }
}
