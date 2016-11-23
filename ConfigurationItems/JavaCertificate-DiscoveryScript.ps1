<#
.SYNOPSIS
 Installs certificates to the Java RTE certificate store of Windows client workstations

 .DESCRIPTION
 Jave keytool is used to import certificates into the Java RTE certificate store for all users

.NOTES
 Author: Mark Allen
 Created: 22-11-2016
 References: https://docs.microsoft.com/en-us/azure/java-add-certificate-ca-store
 Credit to Steve Renard for the Get-JavaHomeLocation function: http://powershell-guru.com/author/powershellgu/
#>

function Get-JavaHomeLocation
{
    $architecture = Get-WmiObject Win32_OperatingSystem | Select -ExpandProperty OSArchitecture

    switch ($architecture)
    {
        '32-bit' { $javaPath = 'HKLM:\SOFTWARE\JavaSoft' }
        '64-bit' { $javaPath = 'HKLM:\SOFTWARE\Wow6432Node\JavaSoft' }
    }

    [string]$currentVersion = Get-ChildItem $javaPath -Recurse | %{($_ | Get-ItemProperty).Psobject.Properties | ?{$_.Name -eq 'CurrentVersion'} } | Select-Object -ExpandProperty Value -Unique -First 1
    Get-ItemProperty "$javaPath\Java Runtime Environment\$currentVersion" -Name JavaHome | Select-Object -ExpandProperty JavaHome
}

<#
 *** Customise here only ***
#>
# create  array for the certificates that should be imported
# The item is the certificate alias
# eg $Certificates = @('my-alias-1','my-alias-2')
$Certificates = @('comodo-certauth','comodo-codesign')
<#
 *** End customisation ***
#>

<#
 Form the relevant file and folder paths
#>
$JavaHome = Get-JavaHomeLocation
$KeyTool = $JavaHome + '\bin\keytool.exe'
$CaCerts = $JavaHome + '\lib\security\cacerts'

<#
 Test that all the relevant paths have been formed correctly
#>
if (!(Test-Path $JavaHome)) {exit}
if (!(Test-Path $KeyTool)) {exit}
if (!(Test-Path $CaCerts)) {exit}

<#
 Iterate through the array of certificates check if they already exist in the certificate store
#>
# if any of the certificates is missing the script will return 'Non-Compliant' and exit immediately, this is a requirement for the ConfigMgr compliance failure status
$Certificates | ForEach-Object { if( (& $KeyTool -list -keystore $CaCerts -storepass changeit -alias $_ -noprompt) -like "keytool error: java.lang.Exception: Alias <*> does not exist" ) { Write-Host 'Non-Compliant';exit } }
Write-Host 'Compliant'