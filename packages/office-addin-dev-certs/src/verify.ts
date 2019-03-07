import {certificateName} from "./default";

function getVerifyCommand(): string {
    switch (process.platform) {
       case "win32":
          return `powershell -command "dir cert:\\CurrentUser\\Root | Where-Object Issuer -like '*CN=${certificateName}*' | Format-List"`;
       case "darwin": // macOS
          return `sudo security find-certificate -c "${certificateName}"`;
       default:
          throw new Error(`Platform not supported: ${process.platform}`);
    }
 }

function isCaCertificateValid(): boolean{
    const command = getVerifyCommand();
    const execSync = require("child_process").execSync;
    try {
        const output = execSync(command, {stdio : "pipe" }).toString();
        if (output.length !== 0) {
            console.log("Certificate found in trusted store");
            return true;
        } else {
            console.log("Certificate not found in trusted store");
        }
    } catch (error) {
        console.log("Error occured while verifying certificate" + error);
    }
    return false;
}

export function verifyCaCertificate(): boolean {
   return isCaCertificateValid();
   // todo add more verfication
}
