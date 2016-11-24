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
    $OSArchitecture = Get-WmiObject -Class Win32_OperatingSystem | Select-Object -ExpandProperty OSArchitecture
    
    switch ($OSArchitecture)
    {
        '32-bit' { $javaPath = 'HKLM:\SOFTWARE\JavaSoft' }
        '64-bit' { $javaPath = 'HKLM:\SOFTWARE\Wow6432Node\JavaSoft' }
        Default  { return 'Unable to determine OS architecture'}
    }
    
    if (Test-Path -Path $javaPath)
    {
        try
        {
            $javaPathRegedit =  Get-ChildItem -Path $javaPath -Recurse -ErrorAction Stop
            [bool]$foundCurrentVersion = ($javaPathRegedit| ForEach-Object {($_ | Get-ItemProperty).PSObject.Properties} | Select-Object -ExpandProperty Name -Unique).Contains('CurrentVersion')
        }
        catch
        {
            return $_.Exception.Message
        }

        if ($foundCurrentVersion)
        {
            [string]$currentVersion = $javaPathRegedit | ForEach-Object {($_ | Get-ItemProperty).PSObject.Properties | Where-Object {$_.Name -eq 'CurrentVersion'} } | Select-Object -ExpandProperty Value -Unique -First 1
            Get-ItemProperty -Path "$javaPath\Java Runtime Environment\$currentVersion" -Name JavaHome | Select-Object -ExpandProperty JavaHome
        }
        else
        {
            return "Unable to retrieve CurrentVersion"
        }
    }
    else
    {
        return "$env:PROCESSOR_ARCHITECTURE : $javaPath not found"
    }
}

<#
 *** Customise here only ***
#>
# create  array for the certificates that should be imported
# The item is the certificate alias
# eg $Certificates = @('my-alias-1','my-alias-2')
$Certificates = @('my-alias-1','my-alias-2')
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
if (!(Test-Path $JavaHome)) {Write-Host "JavaHome error: $JavaHome";exit}
if (!(Test-Path $KeyTool)) {Write-Host "Can't find: $JavaHome";exit}
if (!(Test-Path $CaCerts)) {Write-Host "Can't find: $JavaHome";exit}

<#
 Iterate through the array of certificates check if they already exist in the certificate store
#>
# if any of the certificates is missing the script will return 'Non-Compliant' and exit immediately
$Certificates | ForEach-Object { if( (& $KeyTool -list -keystore $CaCerts -storepass changeit -alias $_ -noprompt) -like "keytool error: java.lang.Exception: Alias <*> does not exist" ) { Write-Host 'Non-Compliant';exit } }
# Returning 'Compliant' is a requirement for a ConfigMgr compliance setting that uses PowerShell
Write-Host 'Compliant'