import * as defaults from "./defaults";

function getInstallCommand(caCertPath: string): string {
   let command: string;
   switch (process.platform) {
      case "win32":
         command = `powershell Import-Certificate -CertStoreLocation cert:\\CurrentUser\\Root ${caCertPath}`;
         break;
      case "darwin": // macOS
         command = `sudo security add-trusted-cert -d -r trustRoot -k /Library/Keychains/System.keychain ${caCertPath}`;
         break;
      default:
         throw new Error(`Platform not supported: ${process.platform}`);
   }
   return command;
}

export function installCaCertificate(caCertPath: string = defaults.caCertificatePath): void {
   const command = getInstallCommand(caCertPath);
   const execSync = require("child_process").execSync;
   try {
      execSync(command, {stdio : "pipe" });
      console.log("Successfully installed certificate to trusted store");
   } catch (error) {
      throw new Error(error.stderr.toString());
   }
}
