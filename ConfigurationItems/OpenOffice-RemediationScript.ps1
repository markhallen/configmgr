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
    Exit
}

$Path = "$env:APPDATA\OpenOffice.org2\user\registry\data\org\openoffice\Office"
if(!(Test-Path -Path $Path))
{
    Exit
}

$Path = "$Path\Jobs.xcu"
if(!(Test-Path -Path $Path))
{
    # create the Job.xcu file
    $Content  = '<?xml version="1.0" encoding="UTF-8"?>
    <oor:component-data xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" oor:name="Jobs" oor:package="org.openoffice.Office">
     <node oor:name="Jobs">
      <node oor:name="UpdateCheck">
       <node oor:name="Arguments">
        <prop oor:name="AutoCheckEnabled" oor:type="xs:boolean">
         <value>false</value>
        </prop>
       </node>
      </node>
     </node>
    </oor:component-data>'
    $Content | Set-Content $Path
    Exit
}

try {
    $Settings = Get-Content -Path $Path
}
catch {
    Exit
}

$LineNumber = ($Settings | Select-String -Pattern '<prop oor:name="AutoCheckEnabled" oor:type="xs:boolean">').LineNumber
if($Settings.Item($LineNumber) | Select-String -Pattern 'false')
{
    Exit
}
$Settings[$LineNumber] = $Settings[$LineNumber] -replace "true", "false"
$Settings | Set-Content $Path
Exit