import * as defaults from "./defaults";

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

export function uninstallCaCertificate(): void {
   const command = getUninstallCommand();
   const execSync = require("child_process").execSync;
   try {
      execSync(command, {stdio : "pipe" });
      console.log("Successfully removed certificate from trusted store");
   } catch (error) {
      throw new Error(error.stderr.toString());
   }
}
