export function verifyCertificates(): boolean {
    let command = `powershell -command "dir cert:\\CurrentUser\\Root | Where-Object Issuer -like '*CN=Developer CA for Microsoft Office Add-ins*' | Format-List"`;
    if (process.platform !== "win32") {
        command = `sudo security find-certificate -c "Developer CA for Microsoft Office Add-ins"`;
    }
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
