import { execSync } from "child_process";
import * as defaults from "./defaults";
import * as logmessages from "./logmessages";
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
      console.log(logmessages.UNINSTALL_SUCESSS_MSG);
   } catch (error) {
      throw new Error(`Unable to uninstall the CA certificate.\n${error.stderr.toString()}`);
   }
}
