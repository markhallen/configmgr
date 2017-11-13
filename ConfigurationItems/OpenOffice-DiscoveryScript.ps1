<#
  This script will check if automatic updates are disabled and return a Compliant/Non-Compliant string.

  Created:     09.11.2017
  Version:     1.0
  Author:      Mark Allen
  Homepage:    https://markallenit.com/
    
  References:
  - unopkg - Apache OpenOfficce Extension Manager
    https://wiki.openoffice.org/wiki/Documentation/DevGuide/Extensions/unopkg
#>

$OSArchitecture = Get-WmiObject -Class Win32_OperatingSystem | Select-Object OSArchitecture

$Path = "$Env:ProgramFiles"
If($OSArchitecture.OSArchitecture -ne "32-bit")
{
    $Path = "$Env:ProgramFiles(x86)"
}
$Path = "$Path\OpenOffice.org 2.1\program\quickstart.exe"
if(!(Test-Path -Path $Path))
{
    # OpenOffice 2.1 is not installed
    Write-Host 'Compliant'
    Exit
}

$Path = "$env:APPDATA\OpenOffice.org2\user\registry\data\org\openoffice\Office\Jobs.xcu"
if(!(Test-Path -Path $Path))
{
    Write-Host "Can't find $Path"
    Exit
}

try {
    $Settings = Get-Content -Path $Path
}
catch {
    Write-Host "Can't read $Path"
    Exit
}

if($Settings.Item(($Settings | Select-String -Pattern '<prop oor:name="AutoCheckEnabled" oor:type="xs:boolean">').LineNumber) | Select-String -Pattern 'false')
{
    Write-Host 'Compliant'
    Exit
}
Write-Host "AutoCheckEnabled is not set to 'false' in $Path"
Exit