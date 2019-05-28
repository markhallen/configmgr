<#
  This script will check if automatic updates are disabled and return a Compliant/Non-Compliant string.
  Created:     28.05.2019
  Version:     1.0
  Author:      Mark Allen
  Homepage:    https://markallenit.com/
#>
$ConfigFile = $Env:AppData + '\Notepad++\config.xml'
if(Test-Path $ConfigFile)
{
	[xml]$Config = Get-Content $ConfigFile
	$UpdatesDisabled = $Config.SelectNodes('//GUIConfig') | Where-Object { $_.Name -eq 'noUpdate' } | ForEach-Object { $_.'#text' }
	if('yes' -ne $UpdatesDisabled)
	{
		'Non-compliant'
		Exit
	}
}
'Compliant'