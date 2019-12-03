<#
This was written because I've had issues with duplicate devices appearing in our ConfigMgr site and the workarounds
I found online were not working for me. The following post explains the timing issue causing duplicates very well:
https://techcommunity.microsoft.com/t5/Configuration-Manager-Archive/Known-Issue-and-Workaround-Duplicate-Records-When-You-Use/ba-p/272965.

This script will delete duplicate records from a specified collection and is intended to be pointed at a collection containing
devices of the same name, excluding copmputers named "unknown". Criteria for deletion is that the ConfigMgr
client is not installed, and that the device does not have SMBIOSGUID data in ConfigMgr. As always... Use at your own risk :)

I commented out the final few lines that called for an AD system discovery because I'm not sure how to limit that yet. To work around this,
I just increased the frequency of our full discovery cycles.

All actions will be logged in the directory the script is run from and named Remove-DuplicateDeviceRecords.log

Manditory Parameters:
-SiteCode
-ProviderMachine (primary site server)
-Device Collection (this collection will be queried for devices with no client or SMBIOSGUID in ConfigMgr
    for deletion)

Optional Parameters:
-$Testing (please, please use this first)

Example Run Script
.\Remove-DuplicateDeviceRecords.ps1 -SiteCode XXX -ProviderMachine HostName.Domain.com -DeviceCollection CollectionWithYourDuplicates -Testing = $true

Written by John Bart
Inspired by https://configmgr.nl/2017/06/06/sccm-duplicate-device-records/
Logging with CMTrace https://adamtheautomator.com/building-logs-for-cmtrace-powershell/
#>


###### SHIT TO DO !!!######################
    # consider adding a sleep timer if triggering by CM rules
    # add full AD discovery to this?
    # should probably add checking for SMBIOSGUID to not exist


##################
# Set Parameters #
##################

Param (
    [Parameter (Mandatory = $true)]
    [string]$SiteCode,

    [Parameter (Mandatory = $true)]
    [string]$ProviderMachineName,

    [Parameter (Mandatory = $true)]
    [string]$DeviceCollection,

    [Parameter (Mandatory = $false)]
    [bool]$Testing
)


#####################
# Declare Functions #
#####################

# The log file will need a home
function Get-ScriptDirectory {
        $global:ScriptDirectory = Get-Location | Convert-Path
}

# Creates a log file if one isn't there and sets the path
function Start-Log {
    [CmdletBinding()]
    param (
        [ValidateScript({ Split-Path $_ -Parent | Test-Path })]
        [string]$FilePath
    )
	
    try {
        if (!(Test-Path $FilePath))	{
	    ## Create the log file
	    New-Item $FilePath -Type File | Out-Null
	}
		
	## Set the global variable to be used as the FilePath for all subsequent Write-Log
	## calls in this session
	$global:ScriptLogFilePath = $FilePath
    }
    catch {
        Write-Error $_.Exception.Message
    }
}

# Appends the log with data
function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter()]
        [ValidateSet(1, 2, 3)]
        [int]$LogLevel = 1
    )

    $TimeGenerated = "$(Get-Date -Format HH:mm:ss).$((Get-Date).Millisecond)+000"
    $Line = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="" type="{4}" thread="" file="">'
    $LineFormat = $Message, $TimeGenerated, (Get-Date -Format MM-dd-yyyy), "$($MyInvocation.ScriptName | Split-Path -Leaf):$($MyInvocation.ScriptLineNumber)", $LogLevel
    $Line = $Line -f $LineFormat
    Add-Content -Value $Line -Path $ScriptLogFilePath
}

# Deletes duplicated devices from the specified collection (only if no client is installed and SMBIOSGUID data is not in ConfigMgr)
function Remove-CMDuplicateDevices {
    try {
        $Duplicates = Get-CMCollectionMember -CollectionName $DeviceCollection | Where-Object {($_.IsClient -ne "True") -and ($_.SMBIOSGUID -eq $null)}
        $TotalToDelete = $Duplicates.Count
        Write-Log "$TotalToDelete duplicate devices detected"
        $Duplicates | ForEach-Object {
            try {
                $ToLog = ("Deleting device - ResourceID: " + $_.ResourceID + "name: " + $_.name)
                Write-Log -Message "$ToLog"
                if ($Testing) {
                    $ToLog = ("TEST MODE! WhatIf: Deleting device - ResourceID: " + $_.ResourceID + "name: " + $_.name)
                    Write-Log -Message "$ToLog"
                    Remove-CMResource -ResourceId $_.ResourceID -WhatIf
                }
                else {
                    $ToLog = ("Deleting device - ResourceID: " + $_.ResourceID + "name: " + $_.name)
                    Write-Log -Message "$ToLog"
                    Remove-CMResource -ResourceId $_.ResourceID -Force
                }

                Remove-CMResource -ResourceId $_.ResourceID -WhatIf
                # Comment out the above -WhatIf to run deletion
            }
            catch {
                $ToLog = ( "Unable to delete ResourceID: " + $_.ResourceID + "name: " + $_.name)
                Write-Log -Message "$ToLog" -LogLevel 2
                Write-Log $_.Exception.Message -LogLevel 2
            }
            
        }
    }
    catch {
        Write-Log $_.Exception.Message -LogLevel 2 
    }
}

###################################
# Start logging and connect to CM #
###################################

# Start logging
try {
    Get-ScriptDirectory
    Start-Log -FilePath "$ScriptDirectory\Remove-DuplicateDevicesFromCollection.log"
}
catch {
    Write-Output "Unable to get working directory for script path" 
}

# Connect to CM
try {
    Write-Log 'Connecting to ConfigMgr'

    # Customizations
    $initParams = @{}
    #$initParams.Add("Verbose", $true) # Uncomment this line to enable verbose logging
    #$initParams.Add("ErrorAction", "Stop") # Uncomment this line to stop the script on any errors

    # Do not change anything below this line

    # Import the ConfigurationManager.psd1 module 
    if((Get-Module ConfigurationManager) -eq $null) {
        Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
    }

    # Connect to the site's drive if it is not already present
    if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
    }

    # Set the current location to be the site code.
    Set-Location "$($SiteCode):\" @initParams
    Write-Log 'Connected to ConfigMgr'
}
catch {
    Write-Log 'Unable to connect to ConfigMgr' -LogLevel 3
    Write-Log $_.Exception.Message -LogLevel 3
    Break
}


##################################################################
# Delete duplicate device records and invoke AD System Discovery #
##################################################################

# Wait 10 minutes for devices to register with ConfigMgr
Start-Sleep -Seconds 300

# Delete duplicates
try {
    Remove-CMDuplicateDevices
}
catch {
    Write-Log $_.Exception.Message
    Break
}

<#
# Call AD System Discovery
if ($Testing) {
    Write-Log "In testing mode. Will not call full AD system discovery"
}
else {
    try {
        Write-Log -Message "Sleeping for 120 seconds"
        Start-Sleep -Seconds 120
        Write-Log -Message "Starting full AD system discovery"
        Invoke-CMSystemDiscovery -SiteCode $SiteCode -ErrorAction SilentlyContinue 
    }
    catch {
        Write-Log $_.Exception.Message
        Break
    }
}
#>