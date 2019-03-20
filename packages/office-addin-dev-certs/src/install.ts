import { execSync } from "child_process";
import * as defaults from "./defaults";

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
