if($args.Count -ne 3){
    throw "Usage uninstall.ps1 <LocalMachine/CurrentUser> <CA certificate path> <CA certficate Name>"
}

$machine = $args[0]
$caCertificatePath=$args[1]
$caCertificateName=$args[2]
if(Get-Command -name Import-Certificate --ErrorAction SilentltyContinue){
    Get-ChildItem  cert:\\$machine\\Root | Where-Object { $_.IssuerName.Name -like "*CN=$caCertificateName*" } |  Remove-Item
}
else{
    # Legacy system support
    $pfx = new-object System.Security.Cryptography.X509Certificates.X509Certificate2
    $pfx.import($caCertificatePath)

    $store = new-object System.Security.Cryptography.X509Certificates.X509Store("Root", $machine)
    $store.open("MaxAllowed")
    $store.remove($pfx)
    $store.close()
}
