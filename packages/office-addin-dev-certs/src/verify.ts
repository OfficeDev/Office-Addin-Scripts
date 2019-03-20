import { execSync } from "child_process";
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

export function isCaCertificateInstalled(): boolean {
    const command = getVerifyCommand();

    try {
        const output = execSync(command, {stdio : "pipe" }).toString();
        if (output.length !== 0) {
            return true;
        }
    } catch (error) {
        // Mac security command throws error if certifcate is not found
    }

    return false;
}

export function verifyCaCertificate(): boolean {
    return isCaCertificateInstalled();
}
