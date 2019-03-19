import * as defaults from "./defaults";

function getVerifyCommand(): string {
    switch (process.platform) {
       case "win32":
          return `powershell -command "dir cert:\\CurrentUser\\Root | Where-Object Issuer -like '*CN=${defaults.certificateName}*' | Where-Object { $_.NotAfter -gt [datetime]::today } | Format-List"`;
       case "darwin": // macOS
          return `security find-certificate -c "${defaults.certificateName}"`;
       default:
          throw new Error(`Platform not supported: ${process.platform}`);
    }
 }

export function isCaCertificateInstalled() {
    const command = getVerifyCommand();
    const execSync = require("child_process").execSync;
    try {
        const output = execSync(command, {stdio : "pipe" }).toString();
        if (output.length !== 0) {
            console.log("CA Certificate is installed.");
        } else {
            throw new Error("CA Certificate is not installed.");
        }
    } catch (error) {
        throw new Error("CA Certificate is not installed."); // Mac security command throws error if certifcate is not found
    }
}

export function verifyCaCertificate(): boolean {
    try {
        isCaCertificateInstalled();
    } catch (err) {
        return false;
    }
    return true;
}
