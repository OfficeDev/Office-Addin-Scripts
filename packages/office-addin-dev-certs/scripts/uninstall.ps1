if($args.Count -ne 2){
    throw "Usage: uninstall.ps1 <LocalMachine | CurrentUser> <CA-certficate-name>"
}

$machine = $args[0]
$caCertificateName=$args[1]
if(Get-Command -name Import-Certificate -ErrorAction SilentlyContinue){
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
