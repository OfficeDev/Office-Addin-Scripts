($apps32bit= Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate) | Out-Null
($apps64bit= Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate) | Out-Null
($allapps = $($apps32bit; $apps64bit) | ConvertTo-Json) | Out-Null
$allapps