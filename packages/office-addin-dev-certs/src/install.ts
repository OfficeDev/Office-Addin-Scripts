import { execSync } from "child_process";
import * as defaults from "./defaults";
import { generateCertificates } from "./generate";
import { uninstallCaCertificate } from "./uninstall";
import { verifyCertificates } from "./verify";

function getInstallCommand(installToLocalMachine: boolean = false, caCertificatePath: string): string {
   let installLocation: string;
   let command: string;

   if (installToLocalMachine) {
      installLocation = "cert:\\LocalMachine\\Root";
   } else {
      installLocation = "cert:\\CurrentUser\\Root";
   }

   switch (process.platform) {
      case "win32":
         command = `powershell Import-Certificate -CertStoreLocation ${installLocation} ${caCertificatePath}`;
         break;
      case "darwin": // macOS
         command = `sudo security add-trusted-cert -d -r trustRoot -k /Library/Keychains/System.keychain ${caCertificatePath}`;
         break;
      default:
         throw new Error(`Platform not supported: ${process.platform}`);
   }
   return command;
}

export async function ensureCertificatesAreInstalled(installToLocalMachine: boolean = false, daysUntilCertificateExpires: number = defaults.daysUntilCertificateExpires) {
   const areCertificatesValid = verifyCertificates();

   if (areCertificatesValid) {
      console.log(`You already have trusted access to https://localhost.\nCertificate: ${defaults.localhostCertificatePath}\nKey: ${defaults.localhostKeyPath}`);
   } else {
      await generateCertificates(defaults.caCertificatePath, defaults.localhostCertificatePath, defaults.localhostKeyPath, daysUntilCertificateExpires);
      await uninstallCaCertificate(false);
      await installCaCertificate(installToLocalMachine);
   }
}

export async function installCaCertificate(installToLocalMachine: boolean = false, caCertificatePath: string = defaults.caCertificatePath) {
    const command = getInstallCommand(installToLocalMachine, caCertificatePath);

    try {
        console.log(`Installing CA certificate "Developer CA for Microsoft Office Add-ins"...`);
        execSync(command, {stdio : "pipe" });
        console.log(`You now have trusted access to https://localhost.\nCertificate: ${defaults.localhostCertificatePath}\nKey: ${defaults.localhostKeyPath}`);
    } catch (error) {
        throw new Error(`Unable to install the CA certificate. ${error.stderr.toString()}`);
    }
}
