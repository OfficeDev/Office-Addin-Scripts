if($args.Count -ne 2){
    throw "Usage: uninstall.ps1 <LocalMachine | CurrentUser> <CA-certficate-name>"
}

# Without this, the script always succeeds (exit code = 0)
$ErrorActionPreference = 'Stop'

$machine = $args[0]
$caCertificateName=$args[1]
if(Get-Command -name Import-Certificate -ErrorAction SilentlyContinue){
    if ($PSVersionTable.PSVersion.Major -le 5) {
        # The following line is required in case pwsh is one of the parent callers
        # because the changes it makes to PSModulePath are not backward compatible with Windows powershell.
        $env:PSModulePath = [Environment]::GetEnvironmentVariable('PSModulePath', 'Machine')
    }
    Get-ChildItem  cert:\\$machine\\Root | Where-Object { $_.IssuerName.Name -like "*CN=$caCertificateName*" } |  Remove-Item
}
else{
    # Legacy system support
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("root", $machine)
    $store.Open("MaxAllowed")
    $certs = $store.Certificates.Find("FindBySubjectName", $caCertificateName, $false)
    foreach ($cert in $certs){
        $store.Remove($cert)
    }
    $store.close()
}
