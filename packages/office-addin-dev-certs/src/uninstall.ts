import { execSync } from "child_process";
import * as fsExtra from "fs-extra";
import * as defaults from "./defaults";
import { isCaCertificateInstalled } from "./verify";

function getUninstallCommand(): string {
   switch (process.platform) {
      case "win32":
         return `powershell -command "Get-ChildItem cert:\\CurrentUser\\Root | where { $_.IssuerName.Name -like '*CN=${defaults.certificateName}*' } |  Remove-Item"`;
      case "darwin": // macOS
         return `sudo security delete-certificate -c "${defaults.certificateName}"`;
      default:
         throw new Error(`Platform not supported: ${process.platform}`);
   }
}

// Deletes the generated certificate files and delete the certificate directory if its empty
function cleanup(): void {
   fsExtra.removeSync(defaults.localhostCertificatePath);
   fsExtra.removeSync(defaults.localhostKeyPath);
   fsExtra.removeSync(defaults.caCertificatePath);

   if (fsExtra.readdirSync(defaults.certificateDirectory).length === 0) {
      fsExtra.removeSync(defaults.certificateDirectory);
   }
}

export function uninstallCaCertificate(verbose: boolean = true): void {
   if (!isCaCertificateInstalled()) {
      if (verbose) {
          console.log(`The CA certificate is not installed.`);
      }
      return;
   }
   const command = getUninstallCommand();
   try {
      console.log(`Uninstalling CA certificate "Developer CA for Microsoft Office Add-ins"...`);
      execSync(command, {stdio : "pipe" });
      cleanup();
      console.log(`You no longer have trusted access to https://localhost.`);
   } catch (error) {
      throw new Error(`Unable to uninstall the CA certificate.\n${error.stderr.toString()}`);
   }
}
