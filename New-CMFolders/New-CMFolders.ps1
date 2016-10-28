# Description:  Create a default set of folders in Configuration Manager
# Created:      27/10/2016
# Author:       Mark Allen
# Usage:        .\New-CMFolders.ps1 -SiteCode PR1 -Path \\%PathToCSVFile% -CustDirectory "PST"

[CmdletBinding( SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "" ) ]
param(
    [parameter(Mandatory=$true)]
    [string]$SiteCode,
    [ValidateScript({Test-Path $(Split-Path $_) -PathType 'Container'})] 
    [string]$Path,
    [parameter(Mandatory=$false)]
    [string]$CustDirectory
)

$CurrentDrive = "$($pwd.Drive.Name):"
$FolderStructure = Import-Csv -Path $Path

Import-Module ($env:SMS_ADMIN_UI_PATH.Substring(0,$env:SMS_ADMIN_UI_PATH.Length – 5) + '\ConfigurationManager.psd1') | Out-Null

#CM12 cmdlets need to be run from the CM12 drive
Set-Location "$($SiteCode):" | Out-Null
if (-not (Get-PSDrive -Name $SiteCode))
{
    Write-Error "There was a problem loading the Configuration Manager powershell module and accessing the site's PSDrive."
    exit 1
}

function New-CMFolder
{
    param (
        [parameter(Mandatory=$true)]
        [String]
            $Path
    )
    
    # catch NULL or empty top level directories
    if($Path.substring($Path.length -1, 1) -eq '\') { return $false }

    if (!(Test-Path $Path))
    {
        Write-Host -ForegroundColor Yellow ("Creating folder $Path")
        New-Item -Path (Split-Path $Path) -Name (Split-Path $Path -Leaf) -ItemType Directory

        if (!(Test-Path $Path))
        {
            Write-Host -ForegroundColor Red ("Failed to create $Path")
            return $false
        }
    }

    Write-Host -ForegroundColor Green ("$Path exists")
    return $true
}

foreach ($Folder in ($FolderStructure.ParentRoot | Get-Unique))
{
    $ThisRoot = $SiteCode + ':\' + $Folder
    if($CustDirectory) 
    {
        $ThisRoot = $ThisRoot + '\' + $CustDirectory
        if(!(New-CMFolder $ThisRoot)) {continue}
    }

    foreach ($ChildOne in (($FolderStructure | Where {$_.ParentRoot -eq $Folder}).ChildOne) | Get-Unique)
    {
        $ChildOnePath = $ThisRoot + '\' + $ChildOne

        if(!(New-CMFolder $ChildOnePath)) {continue}

        foreach ($ChildTwo in (($FolderStructure | Where {$_.ChildOne -eq $ChildOne -and $_.ParentRoot -eq $Folder}).ChildTwo) | Get-Unique)
        {
            $ChildTwoPath = $ChildOnePath + '\' + $ChildTwo
            
            if(!(New-CMFolder $ChildTwoPath)) {continue}

            foreach ($ChildThree in (($FolderStructure | Where {$_.ChildTwo -eq $ChildTwo -and $_.ChildOne -eq $ChildOne -and $_.ParentRoot -eq $Folder}).ChildThree) | Get-Unique)
            {
                $ChildThreePath = $ChildTwoPath + '\' + $ChildThree
                
                if(!(New-CMFolder $ChildThreePath)) {continue}
            }
        }
    }
}

# reset the current directory
Set-Location $CurrentDrive
start-sleep 3
Get-Module -Name ConfigurationManager | remove-module -force