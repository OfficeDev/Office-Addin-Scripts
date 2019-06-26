if($args.Count -ne 2){
    throw "Usage install.ps1 <LocalMachine/CurrentUser> <CA certificate path>"
}

$machine = $args[0]
$caCertificatePath=$args[1]
if(Get-Command -name Import-Certificate --ErrorAction SilentltyContinue){
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