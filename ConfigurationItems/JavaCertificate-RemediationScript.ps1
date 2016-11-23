<#
.SYNOPSIS
 Installs certificates to the Java RTE certificate store of Windows client workstations

 .DESCRIPTION
 The Java keytool.exe is used to import certificates into the Java RTE certificate store for all users

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
# define the publicly acccessible folder that will store certificate files (*.cer)
$ExternalFileStore = '\\WMOGVASCCM01\PkgSrc$\Apps\WMO_COMODO_Certificates\'
# create  a hash table for the certificates that should be imported
# The key is the certificate alias and the item is the certificate file name
# eg $Certificates = @{'my-alias-1' = 'MyCertificate-1.cer';'my-alias-2' = 'MyCertificate-2.cer'}
$Certificates = @{'comodo-certauth' = 'COMODO-RSA-CertificationAuthority.cer';'comodo-codesign' = 'COMODO-RSA-CodeSigningCA.cer'}
<#
 *** End customisation ***
#>

<#
 Form the relevant file and folder paths
#>
$JavaHome = Get-JavaHomeLocation
$KeyTool = $JavaHome + '\bin\keytool.exe'
$CaCerts = $JavaHome + '\lib\security\cacerts'
$CaCertsBak = $CaCerts + '.bak'
$LogFile = $JavaHome + '\lib\security\import-certificates.log'

<#
 Test that all the relevant paths have been formed correctly
#>
if (!(Test-Path $ExternalFileStore)) {"Can't access $ExternalFileStore" | Add-Content $LogFile; exit 1}
if (!(Test-Path $JavaHome)) {"Can't find $JavaHome" | Add-Content $LogFile; exit 1}
if (!(Test-Path $KeyTool)) {"Can't find $KeyTool" | Add-Content $LogFile; exit 1}
if (!(Test-Path $CaCerts)) {"Can't find $CaCerts" | Add-Content $LogFile; exit 1}

<#
 Iterate through the collection of certificates and import if missing from the certificate store
#>
foreach ($Certificate in $Certificates.Keys) {
    # test if the certificate is is already present in the certificate store
    if( (& $KeyTool -list -keystore $CaCerts -storepass changeit -alias $Certificate -noprompt) -like "keytool error: java.lang.Exception: Alias <*> does not exist" )
    {
        $CertificateFile = $ExternalFileStore + $Certificates.Item($Certificate)
        # test that the certificate file exists and can be accessed
        if (!(Test-Path $CertificateFile)) { "Can't access $CertificateFile" | Add-Content $LogFile }
        "Found $CertificateFile, importing $Certificate into $CaCerts." | Add-Content $LogFile
        # execute the import of the certificate file
        & $KeyTool -keystore $CaCerts -storepass changeit -importcert -alias $Certificate -file $CertificateFile -noprompt
    }
}

<#
 Confirm that all of the certificates have been correctly added to the certificate store
#>
"Checking that certificates were correctly added..." | Add-Content $LogFile
# if any of the certificates is missing the script will exit with code 1, this is a requirement for the ConfigMgr remediation failure status
$Certificates.Keys | ForEach-Object { if( (& $KeyTool -list -keystore $CaCerts -storepass changeit -alias $_ -noprompt) -like "keytool error: java.lang.Exception: Alias <*> does not exist" ) { 'Non-Compliant' | Add-Content $LogFile ; Exit 1 } }
'Compliant' | Add-Content $LogFile
Write-Host 'Compliant'