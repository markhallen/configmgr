<#
.SYNOPSIS
 Create collections in Configuration Manager that are useful for operational monitoring.

 .DESCRIPTION
 Will accept an XML file of the desired collections and queries and create the collections in Configuration Manager.

.PARAMETER SiteCode
 3-character site code

.PARAMETER Path
 Full or relative path to the XML file containing the collections and queries

.PARAMETER LimitingCollection
 [Optional] Sets the LimitingCollection used for new collections.
 [Default] "All Systems"

.PARAMETER Organization
 [Optional] allows a default top level directory to be defined. This is useful 
 for a hierarchy that is shared between distinct departments or organizational units.
 The value is also prepended to the collections names to allow each oraganisation to
 create collections with unique names.
 [Default] ""

.PARAMETER FolderName
 [Optional] This is the folder that will contain the operational collections. It will
 be in the root of device collections or within the Organization folder if defined.
 [Default] "Operational"

.PARAMETER RecurInterval
 [Optional] Used in conjunction with RecurCount to set a collection update schedule.
 Acceptable values are 'Minutes','Hours' or 'Days'.
 [Default] "Days"

.PARAMETER RecurCount
 [Optional] Used in conjunction with RecurInterval to set a collection update schedule.
 [Default] 7

.NOTES
 Author: Mark Allen
 Created: 18/10/2016
 References: Benoit Lecours: Set Operational SCCM Collections 
  https://gallery.technet.microsoft.com/Set-of-Operational-SCCM-19fa8178

.EXAMPLE
 .\New-CMOperationalCollections.ps1 -SiteCode PR1 -Path .\MyCollections.xml
 Will create the folder 'Operational' in the root node Device Collections.
 The collections will be created in the Operational Folder.
 EG Device Collections > Operational > <Collections from XML>

.EXAMPLE
 .\New-CMOperationalCollections.ps1 -SiteCode PR1 -Path .\MyCollections.xml -LimitingCollection "All Systems" -Organization "MyOrg"
 Will create a "Test" folder in the root node of each object type; subfolders will be
 created within the relevant folder. Collection names will be prepended by <Organization>
 EG Device Collections > MyOrg > Operational > <Collections from XML>
    Collections will be prepended by "MyOrg ..."

.EXAMPLE
 .\New-CMOperationalCollections.ps1 -SiteCode PR1 -Path .\MyCollections.xml -RecurInterval "Days" -RecurCount "14"
 A custom refesh interval will be set for all new collections.
#>
[CmdletBinding( SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "" ) ]
param(
 [Parameter(Mandatory=$true)]
 [string]$SiteCode,
 [Parameter(Mandatory=$true)]
 [ValidateScript({Test-Path $(Split-Path $_) -PathType 'Container'})] 
 [string]$Path,
 [string]$LimitingCollection = 'All Systems',
 [String]$Organization = '',
 [String]$FolderName = 'Operational',
 [Parameter()]
 [ValidateSet('Minutes','Hours','Days')]
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

#CM cmdlets need to be run from the CM drive
if (-not (Get-PSDrive -Name $SiteCode))
{
    Write-Error "There was a problem loading the Configuration Manager powershell module and accessing the site's PSDrive."
    exit 1
}
Set-Location "$($SiteCode):" | Out-Null

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

    Write-Host -ForegroundColor Green (" - Creating collection $CollectionName")
    New-CMDeviceCollection -Name $CollectionName -Comment $Description -LimitingCollectionName $LimitingCollection -RefreshSchedule $Schedule -RefreshType 2 | Out-Null
        
    Write-Host -ForegroundColor Yellow (" - Adding query to the collection")
    $RuleName = $CollectionName -replace '\s',''
    Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionName -QueryExpression $Query -RuleName $RuleName | Out-Null

    #Move the collection to the correct folder
    if (Test-Path $FolderPath) {
        Write-Host -ForegroundColor White (" - Moving $CollectionName to $FolderPath")
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
if($Organization -ne '')
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
            Write-Host -ForegroundColor Green (" - $CollectionName already exists.")
            Continue
        }
        New-CMOperationalCollection -CollectionName $CollectionName -Query $Coll.Query -Description $Coll.Description -Schedule $Schedule -FolderPath $FolderPath -LimitingCollection $LimitingCollection
    }
Write-Host -ForegroundColor Green ("All collections have been added from $Path")

Exit-Script
