<#
  This script will disable automatic updates.
  Created:     28.05.2019
  Version:     1.0
  Author:      Mark Allen
  Homepage:    https://markallenit.com/
#>
$ConfigFile = $Env:AppData + '\Notepad++\config.xml'
if(Test-Path $ConfigFile)
{
	[xml]$Config = Get-Content $ConfigFile
	$Config.SelectNodes('//GUIConfig') | Where-Object { $_.Name -eq 'noUpdate' } | ForEach-Object { $_.'#text' = 'yes' }
	$Config.Save($ConfigFile)
}