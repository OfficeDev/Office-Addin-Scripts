export function uninstallCertificates(): void {
   let command = `powershell -command "Get-ChildItem cert:\\CurrentUser\\Root | where { $_.IssuerName.Name -like '*CN=Developer CA for Microsoft Office Add-ins*' } |  Remove-Item"`;
   if (process.platform !== "win32") {
      command = `sudo security delete-certificate –c "Developer CA for Microsoft Office Add-ins"`;
   }
   const execSync = require("child_process").execSync;
   try {
      execSync(command, {stdio : "pipe" })
      console.log("Successfully removed certificate from trusted store");
   } catch (error) {
      throw new Error(error.stderr.toString());
   }
}
