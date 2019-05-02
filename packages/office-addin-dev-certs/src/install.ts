import { execSync } from "child_process";
import * as defaults from "./defaults";
import { generateCertificates } from "./generate";
import * as lockFile from "./lockfile";
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

export async function ensureCertificatesAreInstalled(daysUntilCertificateExpires: number = defaults.daysUntilCertificateExpires, machine: boolean = false, maxWaitTimeToAcquireLock: number = defaults.maxWaitTimeToAcquireLock) {
   try {
      await lockFile.lock(defaults.devCertsLockPath, maxWaitTimeToAcquireLock);
   } catch (err) {
      throw new Error(`Another process is using office-addin-dev-certs, please wait for it complete.\n${err}`);
   }
  
   const areCertificatesValid = verifyCertificates();

   if (areCertificatesValid) {
      console.log(`You already have trusted access to https://localhost.\nCertificate: ${defaults.localhostCertificatePath}\nKey: ${defaults.localhostKeyPath}`);
   } else {
      await uninstallCaCertificate(false, false, false);
      await generateCertificates(defaults.caCertificatePath, defaults.localhostCertificatePath, defaults.localhostKeyPath, daysUntilCertificateExpires);
      await installCaCertificate(defaults.caCertificatePath, machine);
   }
   lockFile.unlockSync(defaults.devCertsLockPath);
}

export async function installCaCertificate(caCertificatePath: string = defaults.caCertificatePath, machine: boolean = false) {
   const command = getInstallCommand(caCertificatePath, machine);

   try {
      console.log(`Installing CA certificate "Developer CA for Microsoft Office Add-ins"...`);
      execSync(command, {stdio : "pipe" });
      console.log(`You now have trusted access to https://localhost.\nCertificate: ${defaults.localhostCertificatePath}\nKey: ${defaults.localhostKeyPath}`);
   } catch (error) {
      throw new Error(`Unable to install the CA certificate. ${error.stderr.toString()}`);
   }
}
