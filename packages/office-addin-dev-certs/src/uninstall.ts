import { execSync } from "child_process";
import * as fsExtra from "fs-extra";
import * as path from "path";
import * as defaults from "./defaults";
import * as lockFile from "./lockfile";
import { isCaCertificateInstalled } from "./verify";

function getUninstallCommand(machine: boolean = false): string {
   switch (process.platform) {
      case "win32":
         return `powershell -command "Get-ChildItem  cert:\\${machine ? "LocalMachine" : "CurrentUser"}\\Root | where { $_.IssuerName.Name -like '*CN=${defaults.certificateName}*' } |  Remove-Item"`;
      case "darwin": // macOS
         return `sudo security delete-certificate -c "${defaults.certificateName}"`;
      default:
         throw new Error(`Platform not supported: ${process.platform}`);
   }
}

// Deletes the generated certificate files and delete the certificate directory if its empty
export function deleteCertificateFiles(certificateDirectory: string = defaults.certificateDirectory): void {
   if (fsExtra.existsSync(certificateDirectory)) {
      fsExtra.removeSync(path.join(certificateDirectory, defaults.localhostCertificateFileName));
      fsExtra.removeSync(path.join(certificateDirectory, defaults.localhostKeyFileName));
      fsExtra.removeSync(path.join(certificateDirectory, defaults.caCertificateFileName));

      if (fsExtra.readdirSync(certificateDirectory).length === 0) {
         fsExtra.removeSync(certificateDirectory);
      }
   }
}

export async function uninstallCaCertificate(enableLocking: boolean = true, machine: boolean = false, verbose: boolean = true, maxWaitTimeToAcquireLock: number = defaults.maxWaitTimeToAcquireLock) {
   if (enableLocking) {
      try {
         await lockFile.lock(defaults.devCertsLockPath, maxWaitTimeToAcquireLock);
      } catch (err) {
         throw new Error(`Another process is using office-addin-dev-certs, please wait for it complete.\n${err}`);
      }
   }
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
      console.log(`You no longer have trusted access to https://localhost.`);
   } catch (error) {
      throw new Error(`Unable to uninstall the CA certificate.\n${error.stderr.toString()}`);
   }
   if (enableLocking) {
      lockFile.unlockSync(defaults.devCertsLockPath);
   }
}
