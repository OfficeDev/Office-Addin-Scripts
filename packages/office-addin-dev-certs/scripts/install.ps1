if($args.Count -ne 2){
    throw "Usage: install.ps1 <LocalMachine | CurrentUser> <CA-certificate-path>"
}

# Without this, the script always succeeds (exit code = 0)
$ErrorActionPreference = 'Stop'

$machine = $args[0]
$caCertificatePath=$args[1]
if(Get-Command -name Import-Certificate -ErrorAction SilentlyContinue){
    if ($PSVersionTable.PSVersion.Major -le 5) {
        # The following line is required in case pwsh is one of the parent callers
        # because the changes it makes to PSModulePath are not backward compatible with Windows powershell.
        $env:PSModulePath = [Environment]::GetEnvironmentVariable('PSModulePath', 'Machine')
    }
    Import-Certificate -CertStoreLocation cert:\\$machine\\Root ${caCertificatePath}
}
else{
    # Legacy system support
    $pfx = new-object System.Security.Cryptography.X509Certificates.X509Certificate2
    $pfx.import($caCertificatePath)

    $store = new-object System.Security.Cryptography.X509Certificates.X509Store("Root", $machine)
    $store.open("MaxAllowed")
    $store.add($pfx)
    $store.close()
}
