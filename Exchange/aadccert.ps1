[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$installPath = "$env:PUBLIC\Desktop\Hybrid"
$SSLPath = "C:\CentralSSL"

If(!(test-path $installPath))
{
      New-Item -ItemType Directory -Force -Path $installPath
}

If(!(test-path $SSLPath))
{
      New-Item -ItemType Directory -Force -Path $SSLPath
}

$outpath = "$installPath\1.AADConnect.exe"
$url = "https://download.microsoft.com/download/B/0/0/B00291D0-5A83-4DE7-86F5-980BC00DE05A/AzureADConnect.msi"
Invoke-WebRequest -Uri $url -OutFile $outpath

$outpath = "$installPath\ExchangeCertificatesTemp.zip"
$url = "https://github.com/win-acme/win-acme/releases/download/v2.1.13.1/win-acme.v2.1.13.978.x64.pluggable.zip"
Invoke-WebRequest -Uri $url -OutFile $outpath

$path = "$installPath\2.ExchangeCertificates"
If(!(test-path $path))
{
      New-Item -ItemType Directory -Force -Path $path
}

Expand-Archive -Path $outpath -DestinationPath $path

Remove-Item -Path $outpath

$txtOutput = @"
DOWNLOAD SUCCESSFUL!`n
Please follow the guide on OneNote to continue:`n
https://microsoft.sharepoint.com/teams/OfficePeople/_layouts/OneNote.aspx?id=%2Fteams%2FOfficePeople%2FSiteAssets%2FOffice%20People%20Notebook&wd=target%28Projects%2FTest%20Lab.one%7CA70B945C-4284-48F6-983A-30A0C55E78C1%2FExchange%202016%20Setup%20%2B%20Hybrid%20%26%20AADC%7C97D18BAD-BFC4-4C9E-92C5-AF81E1A13042%2F%29 `
`n The Configuration/Installation order should be: `n
1. Change DNS (If you haven't installed on Msfastlabs.com) `n
2. Set up Exchange Certificates `n
3. Set up Azure AD Connect `n
4. Set up Exchange Hybrid `n
Test `n
"@

$txtOutput | out-file -filepath $installPath\README.txt -append -width 200

# disable IE Enhanced Security Configuration for administrators
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}" -Name "IsInstalled" -Value 0 -ErrorAction Stop

$outpath = "$installPath\ndp48-x86-x64-allos-enu.exe"
$url = "https://download.visualstudio.microsoft.com/download/pr/014120d7-d689-4305-befd-3cb711108212/0fd66638cde16859462a6243a4629a50/ndp48-x86-x64-allos-enu.exe"
Invoke-WebRequest -Uri $url -OutFile $outpath
$arguments="/q /norestart"
Start-Process -FilePath $outpath -ArgumentList $arguments -Wait -PassThru
Remove-Item -Path $outpath

$outpath = "$installPath\vcredist_x64.exe"
$url = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=30679&6B49FDFB-8E5B-4B07-BC31-15695C5A2143=1"
Invoke-WebRequest -Uri $url -OutFile $outpath
$arguments="-q"
Start-Process -FilePath $outpath -ArgumentList $arguments -Wait -PassThru
Remove-Item -Path $outpath

$outpath = "$installPath\vcredists_x64.exe"
$url = "https://aka.ms/highdpimfc2013x64enu"
Invoke-WebRequest -Uri $url -OutFile $outpath
$arguments="-q"
Start-Process -FilePath $outpath -ArgumentList $arguments -Wait -PassThru
Remove-Item -Path $outpath

Install-WindowsFeature Server-Media-Foundation

$outpath = "$installPath\UcmaRuntimeSetup.exe"
$url = "https://download.microsoft.com/download/2/C/4/2C47A5C1-A1F3-4843-B9FE-84C0032C61EC/UcmaRuntimeSetup.exe"
Invoke-WebRequest -Uri $url -OutFile $outpath
$arguments="-q"
Start-Process -FilePath $outpath -ArgumentList $arguments -Wait -PassThru

Install-WindowsFeature Server-Media-Foundation, NET-Framework-45-Features, RPC-over-HTTP-proxy, RSAT-Clustering, RSAT-Clustering-CmdInterface, RSAT-Clustering-Mgmt, RSAT-Clustering-PowerShell, WAS-Process-Model, Web-Asp-Net45, Web-Basic-Auth, Web-Client-Auth, Web-Digest-Auth, Web-Dir-Browsing, Web-Dyn-Compression, Web-Http-Errors, Web-Http-Logging, Web-Http-Redirect, Web-Http-Tracing, Web-ISAPI-Ext, Web-ISAPI-Filter, Web-Lgcy-Mgmt-Console, Web-Metabase, Web-Mgmt-Console, Web-Mgmt-Service, Web-Net-Ext45, Web-Request-Monitor, Web-Server, Web-Stat-Compression, Web-Static-Content, Web-Windows-Auth, Web-WMI, Windows-Identity-Foundation, RSAT-ADDS

Unblock-File -Path $path\Scripts\ImportExchange.ps1
