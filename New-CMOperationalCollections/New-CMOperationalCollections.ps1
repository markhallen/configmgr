# Description:  
# Created:      18/10/2016
# Author:       Mark Allen
# Usage:        .\New-CMOperationalCollections.ps1 -SiteCode PR1 -Path \\%PathToXMLFile% -LimitingCollection "All Systems" -Organization "PST"
# References:
#       Benoit Lecours: Set Operational SCCM Collections https://gallery.technet.microsoft.com/Set-of-Operational-SCCM-19fa8178

[CmdletBinding( SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "" ) ]
param(
 [Parameter(Mandatory=$true)]
 [string]$SiteCode,
 [ValidateScript({Test-Path $(Split-Path $_) -PathType 'Container'})] 
 [string]$Path,
 [string]$LimitingCollection = 'All Systems',
 [String]$Organization = $null,
 [String]$FolderName = 'Operational',
 [String]$RecurInterval = 'Days',
 [Int]$RecurCount = 7
)
$RootFolder = $SiteCode + ':\' + 'DeviceCollection'

# record the starting location
$CurrentDrive = "$($pwd.Drive.Name):"

# Import the XML data
[xml]$OperationalCollections = Get-Content $Path

# Import the Configuration Manager PowerShell module
Import-Module ($env:SMS_ADMIN_UI_PATH.Substring(0,$env:SMS_ADMIN_UI_PATH.Length – 5) + '\ConfigurationManager.psd1') | Out-Null

#CM12 cmdlets need to be run from the CM12 drive
Set-Location "$($SiteCode):" | Out-Null
if (-not (Get-PSDrive -Name $SiteCode))
    {
        Write-Error "There was a problem loading the Configuration Manager powershell module and accessing the site's PSDrive."
        exit 1
    }

Function New-CMOperationalCollection
{
    param (
        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]
            $CollectionName,
        [parameter(Mandatory=$true)]
        [String]
            $Query,
        [String]
            $Description,
        [parameter(Mandatory=$true)]
        [Microsoft.ConfigurationManagement.ManagementProvider.WqlQueryEngine.WqlResultObjectBase]
            $Schedule,
        [String]
            $FolderPath,
        [String]
            $LimitingCollection
    )

    Write-Host -ForegroundColor Yellow (" - Creating collection $CollectionName")
    New-CMDeviceCollection -Name $CollectionName -Comment $Description -LimitingCollectionName $LimitingCollection -RefreshSchedule $Schedule -RefreshType 2 | Out-Null
        
    Write-Host -ForegroundColor Yellow (" - Adding query to the collection")
    $RuleName = $CollectionName -replace '\s',''
    Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionName -QueryExpression $Query -RuleName $RuleName -Verbose

    #Move the collection to the correct folder
    if (Test-Path $FolderPath) {
        Write-Host -ForegroundColor Yellow (" - Moving $CollectionName to $FolderPath")
        Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionName)
    }
}

Function Exit-Script
{
    # reset the current directory
    Set-Location $CurrentDrive
    start-sleep 3
    Get-Module -Name ConfigurationManager | remove-module -force
    exit
}

#############################################
# Check that the limiting collection exists #
#############################################
$ThisLimitingCollection = Get-CMDeviceCollection -Name $LimitingCollection
if($ThisLimitingCollection.Name -ne $LimitingCollection)
{
    Write-Host -ForegroundColor Red (" - The limiting collection $LimitingCollection does not exist.")
    Exit-Script
}
Write-Host -ForegroundColor Green (" - The limiting collection $LimitingCollection exists.")

########################################
# Check or the create folder structure #
########################################
$ParentPath = $RootFolder

# create an organization folder if required
if($Organization -ne $null)
{
    $ParentPath = $ParentPath + "\" + $Organization
    if (!(Test-Path $ParentPath))
    {
        Write-Host -ForegroundColor Yellow ("Creating folder $ParentPath")
        New-Item -Name $ParentPath -ItemType Directory
    }
}

# create the folder
$FolderPath = $ParentPath + "\" + $FolderName
if (!(Test-Path $FolderPath))
{
    Write-Host -ForegroundColor Yellow ("Creating folder $FolderPath")
    New-Item -Name $FolderName -Path $ParentPath -ItemType Directory
}

######################################
# Create the operational collections #
######################################

# Create refresh Schedule
$Schedule = New-CMSchedule –RecurInterval $RecurInterval –RecurCount $RecurCount

foreach ($Coll in $OperationalCollections.Collections.Collection)
    {
        $CollectionName = "$($Coll.Name)"
        if($Organization -ne $null) { $CollectionName = "$Organization $CollectionName" }

        $CollectionTest = Get-CMDeviceCollection -Name $CollectionName
        # Check install collection exists
        if($CollectionTest.Name -eq $CollectionName)
        {
            Write-Host -ForegroundColor Yellow (" - $CollectionName already exists.")
            Continue
        }
        New-CMOperationalCollection -CollectionName $CollectionName -Query $Coll.Query -Description $Coll.Description -Schedule $Schedule -FolderPath $FolderPath -LimitingCollection $LimitingCollection
    }
Write-Host -ForegroundColor Green ("All collections have been added from $Path")

Exit-Script