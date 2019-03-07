function getUninstallCommand(): string {
   let command: string;
   switch (process.platform) {
      case "win32":
         command = `powershell -command "Get-ChildItem cert:\\CurrentUser\\Root | where { $_.IssuerName.Name -like '*CN=Developer CA for Microsoft Office Add-ins*' } |  Remove-Item"`;
      case "darwin": // macOS
         command = `sudo security delete-certificate –c "Developer CA for Microsoft Office Add-ins"`;
      default:
         throw new Error(`Platform not supported: ${process.platform}`);
   }
   return command;
}

export function uninstallCaCertificate(): void {
   const command = getUninstallCommand();
   const execSync = require("child_process").execSync;
   try {
      execSync(command, {stdio : "pipe" })
      console.log("Successfully removed certificate from trusted store");
   } catch (error) {
      throw new Error(error.stderr.toString());
   }
}
