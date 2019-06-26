if($args.Count -ne 1){
    throw "Usage verify.ps1 <CA certficate Name>"
}

$caCertificateName=$args[0]
if(Get-Command -name Import-Certificate --ErrorAction SilentltyContinue){
    Get-ChildItem cert:\\CurrentUser\\Root | Where-Object Issuer -like "*CN=$caCertificateName*" | Where-Object { $_.NotAfter -gt [datetime]::today } | Format-List
}
else{
    # Legacy system support
    Get-ChildItem cert:\\CurrentUser\\Root | Where-Object { $_.Subject -like "*CN=$caCertificateName*"} | Where-Object { $_.NotAfter -gt [datetime]::today } | Format-List
}
