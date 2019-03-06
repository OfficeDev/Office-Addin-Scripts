export function installCertificates(caCertPath: string): void {
   let command = "powershell Import-Certificate -CertStoreLocation cert:\\CurrentUser\\Root ";
   if (process.platform !== "win32") {
      command = "sudo security add-trusted-cert -d -r trustRoot -k /Library/Keychains/System.keychain";
   }
   const finalCommand = `${command} ${caCertPath}`;
   const execSync = require("child_process").execSync;
   try {
      execSync(finalCommand, {stdio : "pipe" });
      console.log("Successfully installed certificate to trusted store");
   } catch (error) {
      throw new Error(error.stderr.toString());
   }
}
