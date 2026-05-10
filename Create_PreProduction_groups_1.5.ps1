<#
.SYNOPSIS
Create Intune preproduction pilot groups using Microsoft Graph.

.DESCRIPTION
This script automates the creation and maintenance of Intune preproduction pilot
and nested roll-up groups using Microsoft Graph.

It performs the following actions:
- discovers source Autopilot device groups based on provided name filters;
- evaluates each group member for compliance, management state, and recent activity;
- optionally logs inactive or non-compliant devices to CSV files;
- creates or updates corresponding "-PreProductionDevices" pilot groups containing
  a randomized 10% sample of active, compliant devices;
- removes outdated or inactive devices from existing pilot groups;
- creates or updates zone-level roll-up groups and nests relevant pilot groups;
- creates or updates region-level roll-up groups and nests relevant zone groups.

The script relies on Microsoft Graph PowerShell cmdlets and requires valid
Azure AD application credentials with appropriate directory permissions.

.PARAMETER ClientId
Azure AD application (client) ID. 

.PARAMETER ClientSecret
Path to a file containing the Azure AD application client secret, or the secret
value itself if the file path is not used.

.PARAMETER TenantId
Azure AD tenant ID.

.PARAMETER Filters
Comma-separated list of group name patterns used to locate source Autopilot device
groups. Default is 'Autopilot-Desktops-Devices','Autopilot-Notebooks-Devices'.

.PARAMETER LogInactiveDevices
Optional path to store CSV files with inactive device details. If empty, no CSV
files are created.

.PARAMETER Log
Optional path to store the script log file. If empty, logging to file is disabled.

.NOTES
Version:          1.5
Author:           Viktor Kravtsov
Creation Date:    2025-03-05
Purpose/Change:   final design with bug fixes

Prerequisites:    Microsoft.Graph PowerShell module

.EXAMPLE
.\Create_PreProduction_groups_1.5.ps1 \
    -ClientId '<ClientId>' \
    -ClientSecret 'D:\path\to\secret.txt' \
    -TenantId '<tenantId>' \
    -Filters "'Autopilot-Desktops-Devices','Autopilot-Notebooks-Devices'" \
    -LogInactiveDevices 'C:\TEMP\PilotGroupsCreation\InactiveDevices' \
    -Log 'C:\Temp\PilotGroupsCreation\Log'
#>

Param(
    [Parameter(Mandatory = $false)][string] $ClientId,
    [Parameter(Mandatory = $false)][string] $ClientSecret,
    [Parameter(Mandatory = $false)][string] $tenantid,
    [Parameter(Mandatory = $false)][string] $Filters = "'Autopilot-Desktops-Devices','Autopilot-Notebooks-Devices'",
    [Parameter(Mandatory = $false)][string] $LogInactiveDevices = '.\log\InactiveDevices',
    [Parameter(Mandatory = $False)][string] $Log = ".\Log"


)

####################################################
#region functions
####################################################
Function Get-Groups {
    param
    (
        [Parameter(Mandatory = $true)]$Filters,
        [Parameter(Mandatory = $false)]$Groups,
        [Parameter(Mandatory = $false)][switch]$Dynamic = $false
    )

    #Convert filter string to array
    $Filters = ($Filters.Replace('"', '')).Replace("'", '') -split ","

    If ($Filters.Count -gt 1) {
        $modifiedFilters = $Filters | ForEach-Object { '($_.displayName -like "*' + $_ + '*")' }
        $condition = $modifiedFilters -Join " -or"
    }
    else { $condition = '($_.displayName -like "*' + $Filters + '*")' }


    $group = $null
    If ($Groups) { $group = $Groups }
    else {
        #Dirt native pre-search
        Foreach ($Filter in $Filters) {
            $search = "'" + '"DisplayName:' + $filter + '"' + "'"
            $group += (get-mggroup -Search (Invoke-Expression $search) -ConsistencyLevel eventual -All)
        }
    }
    #precise search
    if ($Dynamic) {
        $Foundgroups = ($group | Where-Object { (Invoke-Expression $condition) -and ($_.grouptypes -eq 'DynamicMembership') })
    }
    else {
        $Foundgroups = ($group | Where-Object { (Invoke-Expression $condition) })
    }
    return $Foundgroups
}

####################################################
Function Create-Group {
    param
    (
        [Parameter(Mandatory = $false)]$adminUnitDisplayName,
        [Parameter(Mandatory = $true)]$groupName,
        [Parameter(Mandatory = $false)]$groupDescription,
        [Parameter(Mandatory = $false)]$MailNickname
    )

    $filter = "DisplayName eq '" + $adminUnitDisplayName + "'"
    $adminUnitObj = Get-MgDirectoryAdministrativeUnit -Filter $filter

    $params = @{
        "@odata.type"   = "#microsoft.graph.group"
        Description     = $groupDescription
        DisplayName     = $groupName
        mailEnabled     = $false
        MailNickname    = $MailNickname
        securityEnabled = $true
    }
    $created = New-MgDirectoryAdministrativeUnitMember -AdministrativeUnitId $adminUnitObj.Id -BodyParameter $params
    return $created
}
####################################################
Function DeviceIsActive {
    param
    (
        [Parameter(Mandatory = $true)]$device,
        [Parameter(Mandatory = $true)]$DaysActive
    )
    # Get device details
    #$device = Get-MgDevice -DeviceId $Id
    # Check if the device is compliant and active
    if ($device.IsCompliant -eq $true -and $device.IsManaged -eq $true) {
        # Check last logon time (assuming it's available in the device properties)
        $lastLogonTime = $device.ApproximateLastSignInDateTime
        if ($lastLogonTime -and ($lastLogonTime -gt (Get-Date).AddDays(-$DaysActive))) { return $true }
        else { return $false }
    }
    else { return $false }
}

#endregion

####################################################
################ MAIN EXECUTION ####################
####################################################

#Set logging
If ($Log) {
    $Log = ($Log + '\').replace('\\', '\')
    If (!(Test-Path -LiteralPath $Log)) { New-Item ($Log) -Force -ItemType Directory }
    $Logfile = ($Log + 'Set-PreProduction_Groups.log')
    $Logfile_old = ($Log + 'Set-PreProduction_Groups_old.log')
    If (Test-Path -LiteralPath $Logfile) {
        $Size = [math]::round(((Get-Item $Logfile).Length) / 1MB, 5)
        If ($Size -ge 50.0) { Move-Item -Path $Logfile -Destination $Logfile_old -Force -ErrorAction Stop }
    }
}

$msg = ("$env:computername : " + (Get-Date -Format "dd-MM-yyyy HH:mm:ss K") + "(UTC) : START   : " + $env:USERNAME )
Write-host $msg  -ForegroundColor Black -BackgroundColor Cyan
If ($Log) { Write-Output $msg | Out-File -LiteralPath $Logfile -Append }#-Encoding utf8


# checking prerequisites
if (Get-Module -ListAvailable -Name Microsoft.Graph) { Write-Host "Microsoft.Graph Module exists" } 
else {
    $msg = ("$env:computername : " + (Get-Date -Format "dd-MM-yyyy HH:mm:ss K") + "(UTC) : WARNING : Microsoft.Graph Module does not exist, installing")
    Write-host $msg  -ForegroundColor Yellow
    If ($Log) { Write-Output $msg | Out-File -LiteralPath $Logfile -Append }#-Encoding utf8
    Install-Module Microsoft.Graph -Scope CurrentUser
}


try {
    ################################################
    ##variables
    $exclusionList = @()
    $ActiveParentGroupCounter = 0
    $NewPilotGroupCounter = 0
    $ExistingPilotGroupCounter = 0
    $NewRegionZoneGroupCounter = 0
    $ExistingRegionZoneGroupCounter = 0
    $NewRegionGroupCounter = 0
    $ExistingRegionGroupCounter = 0
    ################################################


    #set MS Graph connection authentification
    if (Test-Path -LiteralPath $ClientSecret) {
        try {
            $ClientSecret = get-content $ClientSecret -erroraction stop
            if ($null -eq $ClientSecret) {
                $msg = ("$env:computername : " + (Get-Date -Format "dd-MM-yyyy HH:mm:ss K") + "(UTC) : ERROR   : Please supply Azure Application credentials with proper permissions via txt files: ClientSecret.txt")
                Write-host $msg  -ForegroundColor RED
                If ($Log) { Write-Output $msg | Out-File -LiteralPath $Logfile -Append }
                Exit 999
            }
        }
        catch {
            $msg = ("$env:computername : " + (Get-Date -Format "dd-MM-yyyy HH:mm:ss K") + "(UTC) : ERROR   : Unable to obtain credentials/tenant info: $($_.Exception.Message)")
            Write-host $msg  -ForegroundColor RED
            If ($Log) { Write-Output $msg | Out-File -LiteralPath $Logfile -Append }
            Exit 999
        }
    }


    # Connect to Microsoft Graph
    $secureAppSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential ($ClientId, $secureAppSecret)
    Connect-MgGraph -ClientSecretCredential $credential -TenantId $tenantId 

    $msg = ("$env:computername : " + (Get-Date -Format "dd-MM-yyyy HH:mm:ss K") + "(UTC) : INFO    : Start group search by fiter: " + $Filters)
    Write-host $msg  -ForegroundColor WHITE
    If ($Log) { Write-Output $msg | Out-File -LiteralPath $Logfile -Append }



    $devicegroups = Get-Groups -Filters $Filters -Dynamic
    $devicegroups = $devicegroups | Where-Object { $_.DisplayName -notin $exclusionList }

    $msg = ("$env:computername : " + (Get-Date -Format "dd-MM-yyyy HH:mm:ss K") + "(UTC) : INFO    : Found groups: " + $devicegroups.Count + ", evaluation started.")
    Write-host $msg  -ForegroundColor WHITE
    If ($Log) { Write-Output $msg | Out-File -LiteralPath $Logfile -Append }

    Foreach ($group in $devicegroups) {
        # Get all members of the device group
        $groupMembers = Get-MgGroupMember -GroupId $group.id -All

        If ($LogInactiveDevices) {
            $InactiveDevices = @()
            $LogInactiveDevices = ($LogInactiveDevices + '\').replace('\\', '\')
            If (!(Test-Path -LiteralPath $LogInactiveDevices)) { New-Item ($LogInactiveDevices) -Force -ItemType Directory }
        }

        # Initialize an array to store compliant and active devices
        $compliantActiveDevices = @()

        # Loop through each member to check compliance and activity status
        foreach ($member in $groupMembers) {

            # Get device details
            $device = Get-MgDevice -DeviceId $member.Id -ea SilentlyContinue

            # Check if the device is compliant and active within the last 30 days
            If ($device) {
                if (DeviceIsActive -device $device -DaysActive 30) { $compliantActiveDevices += $device }
                elseIf ($LogInactiveDevices) { $InactiveDevices += $device | Select-Object -Property Id, MdmAppId, model, DisplayName, RegisteredOwners, RegisteredUsers, ApproximateLastSignInDateTime, IsCompliant, IsManaged }
            }

        }
        # Export inactive devices to csv
        If ($InactiveDevices -and ($LogInactiveDevices)) {
            # Prepare the data for CSV export
            $deviceData = $InactiveDevices | ForEach-Object {
                [PSCustomObject]@{
                    DeviceId         = $_.Id
                    IntuneId         = $_.MdmAppId
                    DeviceModel      = $_.Model
                    DisplayName      = $_.DisplayName
                    IsCompliant      = $_.IsCompliant
                    IsManaged        = $_.IsManaged
                    LastActivityDate = $_.ApproximateLastSignInDateTime
                    #RegisteredOwners = ($_.RegisteredOwners | ForEach-Object { $_.DisplayName }) -join "; "
                    #RegisteredUsers  = ($_.RegisteredUsers | ForEach-Object { $_.DisplayName }) -join "; "
                    RegisteredOwners = $_.RegisteredOwners
                    RegisteredUsers  = $_.RegisteredUsers
                }
            }

            # Export the data to a CSV file
            $csvPath = $LogInactiveDevices + $group.DisplayName + "_InactiveDevices.csv"
            $deviceData | Export-Csv -Path $csvPath -NoTypeInformation -Delimiter ','

        }
        # Output the compliant and active devices
        $Info = ( $group.displayname + ' active devices: ' + $compliantActiveDevices.Count.ToString() + "/" + $groupMembers.count.ToString())            
        $msg = ("$env:computername : " + (Get-Date -Format "dd-MM-yyyy HH:mm:ss K") + "(UTC) : INFO    : " + $Info)
        Write-host $msg  -ForegroundColor WHITE
        If ($Log) { Write-Output $msg | Out-File -LiteralPath $Logfile -Append }

        If ($compliantActiveDevices.Count -ne 0) {
            $ActiveParentGroupCounter++
            # Get 10% of the compliant and active devices
            $sampleSize = [math]::Ceiling($compliantActiveDevices.Count * 0.1)
            $pilotDevices = $compliantActiveDevices | Get-Random -Count $sampleSize

            # Create or update the 'Pilot-devices' group
            $pilotGroupName = $group.displayname.Replace('-Devices', '-PreProductionDevices')
            $pilotGroupName = $pilotGroupName.Replace('-D-', '-S-')

            # Check if the 'Pilot-devices' group already exists
            $existingPilotGroup = Get-MgGroup -Filter "displayName eq '$pilotGroupName'" -All

            if (!($existingPilotGroup)) {
                # Create the 'Pilot-devices' group
                $pilotGroup = Create-Group -adminUnitDisplayName "<ADMUName>" -groupName $pilotGroupName -groupDescription "This is created automatically group containing 10% of active devices of Autopilot-Desktops/Notebook-Devices groups." -MailNickname 'PilotWindows'
                If ($pilotGroup) {
                    $NewPilotGroupCounter++
                    $pilotGroupId = $pilotGroup.Id
                }

                # Add devices to the 'Pilot-devices' group
                ForEach ($pilotDevice in $pilotDevices) {
                    New-MgGroupMember -GroupId $pilotGroupId -DirectoryObjectId $pilotDevice.Id
                }
                $members = Get-MgGroupMember -GroupId $pilotGroupId -All
                $Count = $members.Count
                $Info = "$pilotGroupName group has been created with " + $Count + " compliant and active devices."
                $msg = ("$env:computername : " + (Get-Date -Format "dd-MM-yyyy HH:mm:ss K") + "(UTC) : SUCCESS : " + $Info)
                Write-host $msg  -ForegroundColor Green
                If ($Log) { Write-Output $msg | Out-File -LiteralPath $Logfile -Append }
            } 
            else {
                # Get the existing group's ID
                $pilotGroupId = $existingPilotGroup.Id
                # Get the members of the group
                $members = Get-MgGroupMember -GroupId $pilotGroupId -All

                $ExistingPilotGroupCounter++

                foreach ($member in $members) {
                    # Get device details
                    $device = Get-MgDevice -DeviceId $member.Id
                    # Check if the device is compliant and active
                    If ($device) {
                        if (!(DeviceIsActive -device $device -DaysActive 30)) { Remove-MgGroupMemberByRef -GroupId $pilotGroupId -DirectoryObjectId $device.Id }
                    }
                }
                # Get the members of the group after cleanup
                $members = Get-MgGroupMember -GroupId $pilotGroupId -All
                If ($members.Count -lt $pilotDevices.Count) {
                    $memberdevices = @()
                    Foreach ($member in $members) {
                        $memberdevices += Get-MgDevice -DeviceId $member.Id -ea SilentlyContinue
                    }

                    # Create a new list to store devices that are not in the group
                    $updatedPilotDevices = @()
                    # Loop through each device in the pilotDevices list
                    foreach ($pilotDevice in $compliantActiveDevices) {
                        # Check if the device is already a member of the group
                        $isMember = $memberdevices | Where-Object { $_.Id -eq $pilotDevice.Id }
                        if (-not $isMember) { $updatedPilotDevices += $pilotDevice }
                    }
                    $AddPilotDevices = $null
                    $AddPilotDevices = $updatedPilotDevices | Get-Random -Count ($pilotDevices.Count - $members.Count)
                    # Add devices to the 'Pilot-devices' group
                    ForEach ($addPilotDevice in $AddPilotDevices) {
                        New-MgGroupMember -GroupId $pilotGroupId -DirectoryObjectId $addPilotDevice.Id
                    }
                    $members = Get-MgGroupMember -GroupId $pilotGroupId -All
                    $Count = $members.Count

                    $Info = "$pilotGroupName group has been updated to " + $Count + " compliant and active devices."
                    $msg = ("$env:computername : " + (Get-Date -Format "dd-MM-yyyy HH:mm:ss K") + "(UTC) : SUCCESS : " + $Info)
                    Write-host $msg  -ForegroundColor Green
                    If ($Log) { Write-Output $msg | Out-File -LiteralPath $Logfile -Append }
                }
                else {
                    $Count = $members.Count
                    $Info = "$pilotGroupName group already contains " + $Count + " compliant and active devices."
                    $msg = ("$env:computername : " + (Get-Date -Format "dd-MM-yyyy HH:mm:ss K") + "(UTC) : SUCCESS : " + $Info)
                    Write-host $msg  -ForegroundColor Green
                    If ($Log) { Write-Output $msg | Out-File -LiteralPath $Logfile -Append }
                }
            }
        }
    }

    $Info = "Active Parent Groups processed: $ActiveParentGroupCounter / New Pilot Groups created:$NewPilotGroupCounter / Existing Pilot Groups checked/updated: $ExistingPilotGroupCounter."
    $msg = ("$env:computername : " + (Get-Date -Format "dd-MM-yyyy HH:mm:ss K") + "(UTC) : FINISH  : " + $Info)
    Write-host $msg  -ForegroundColor Black -BackgroundColor Cyan
    If ($Log) { Write-Output $msg | Out-File -LiteralPath $Logfile -Append }



    ############################################################################################
    # Zone PreProduction Groups creation
    ############################################################################################

    $Info = "CREATION/UPDATE Zone PREPRODUCTION GROUPS STARTED"
    $msg = ("$env:computername : " + (Get-Date -Format "dd-MM-yyyy HH:mm:ss K") + "(UTC) : START   : " + $Info)
    Write-host $msg  -ForegroundColor Black -BackgroundColor Cyan
    If ($Log) { Write-Output $msg | Out-File -LiteralPath $Logfile -Append }



    # Zone group creation
    $PilotGroupFilters = "'Autopilot-Desktops-PreProductionDevices','Autopilot-Notebooks-PreProductionDevices'"
    $RegionZone = @()
    $devicegroups = Get-Groups -Filters $PilotGroupFilters

    $devicegroups | Select-Object -Property DisplayName | ForEach-Object { $RegionZone += (($_.DisplayName -split '-')[3] + '-' + ($_.DisplayName -split '-')[4]) }
    $RegionZoneNames = $RegionZone | Select-Object -Unique

    Foreach ($RegionZoneName in $RegionZoneNames) {

        $ZoneNameDeviceGroups = Get-Groups -Filters $RegionZoneName -Groups $devicegroups

        # Define the Zone group details
        $RegionZoneGroupName = "Prefix-" + $RegionZoneName + "-Autopilot-PreProduction"

        # Check if the RegionZone group exists
        $RegionZoneGroup = Get-MgGroup -Filter "displayName eq '$RegionZoneGroupName'"
        if (-not $RegionZoneGroup) {
            # Create the RegionZone group if it doesn't exist
            $RegionZoneGroup = Create-Group -adminUnitDisplayName "ADMUName" -groupName $RegionZoneGroupName -groupDescription "This is created automatically Zone group containing per-country Autopilot-Desktops/Notebook-PilotDevices groups." -MailNickname 'ZonePreProductionWindows'
            If ($RegionZoneGroup) {
                $NewRegionZoneGroupCounter++
                $Info = "Zone group '$RegionZoneGroupName' created."
                $msg = ("$env:computername : " + (Get-Date -Format "dd-MM-yyyy HH:mm:ss K") + "(UTC) : SUCCESS : " + $Info)
                Write-host $msg  -ForegroundColor GREEN
                If ($Log) { Write-Output $msg | Out-File -LiteralPath $Logfile -Append }
            }
        } 
        else {
            $ExistingRegionZoneGroupCounter++
            $Info = "Zone group '$RegionZoneGroupName' already exists."
            $msg = ("$env:computername : " + (Get-Date -Format "dd-MM-yyyy HH:mm:ss K") + "(UTC) : SUCCESS : " + $Info)
            Write-host $msg  -ForegroundColor GREEN
            If ($Log) { Write-Output $msg | Out-File -LiteralPath $Logfile -Append }
        }

        # Get all groups matching the child group patterns
        $childGroups = Get-Groups -Filters $RegionZoneName -Groups $devicegroups

        # Add each child group to the RegionZone group only if it is not empty
        foreach ($childGroup in $childGroups) {
            # Check if the child group is not empty
            $childGroupMembers = Get-MgGroupMember -GroupId $childGroup.Id
            if ($childGroupMembers.Count -gt 0) {
                $isMember = Get-MgGroupMember -GroupId $RegionZoneGroup.Id | Where-Object { $_.Id -eq $childGroup.Id }
                if (-not $isMember) {
                    New-MgGroupMember -GroupId $RegionZoneGroup.Id -DirectoryObjectId $childGroup.Id
                    $Info = "Pilot group '$($childGroup.DisplayName)' added to Zone group '$RegionZoneGroupName'."
                    $msg = ("$env:computername : " + (Get-Date -Format "dd-MM-yyyy HH:mm:ss K") + "(UTC) : INFO    : " + $Info)
                    Write-host $msg  -ForegroundColor WHITE
                    If ($Log) { Write-Output $msg | Out-File -LiteralPath $Logfile -Append }

                } 
                else {
                    $Info = "Pilot group '$($childGroup.DisplayName)' is already a member of the Zone group."
                    $msg = ("$env:computername : " + (Get-Date -Format "dd-MM-yyyy HH:mm:ss K") + "(UTC) : INFO    : " + $Info)
                    Write-host $msg  -ForegroundColor WHITE
                    If ($Log) { Write-Output $msg | Out-File -LiteralPath $Logfile -Append }

                }
            } 
            else {
                $Info = "Pilot group '$($childGroup.DisplayName)' is empty and will not be added to the Zone group."
                $msg = ("$env:computername : " + (Get-Date -Format "dd-MM-yyyy HH:mm:ss K") + "(UTC) : INFO    : " + $Info)
                Write-host $msg  -ForegroundColor YELLOW
                If ($Log) { Write-Output $msg | Out-File -LiteralPath $Logfile -Append }
            }
        }
    }


    $Info = "Zone PreProduction group check and creation process completed. New Zone PreProduction groups created: $NewRegionZoneGroupCounter / Existing Zone PreProduction groups checked/updated: $ExistingRegionZoneGroupCounter"
    $msg = ("$env:computername : " + (Get-Date -Format "dd-MM-yyyy HH:mm:ss K") + "(UTC) : FINISH  : " + $Info)
    Write-host $msg  -ForegroundColor Black -BackgroundColor Cyan
    If ($Log) { Write-Output $msg | Out-File -LiteralPath $Logfile -Append }

    # Region group creation
    $PreprodZoneGroupFilters = '-Autopilot-PreProduction'
    $Region = @()
    $devicegroups = Get-Groups -Filters $PreprodZoneGroupFilters

    $devicegroups | Select-Object -Property DisplayName | ForEach-Object { $Region += (($_.DisplayName -split '-')[3]) }
    $RegionNames = $Region | Select-Object -Unique

    Foreach ($RegionName in $RegionNames) {

        $RegionNameDeviceGroups = Get-Groups -Filters $RegionName -Groups $devicegroups

        # Define the Zone group details
        $RegionGroupName = "Prefix-" + $RegionName + "-Autopilot-PreProduction"

        # Check if the Region group exists
        $RegionGroup = Get-MgGroup -Filter "displayName eq '$RegionGroupName'"
        if (-not $RegionGroup) {
            # Create the Region group if it doesn't exist
            $RegionGroup = Create-Group -adminUnitDisplayName "ADMUName" -groupName $RegionGroupName -groupDescription "This is created automatically Region group containing per-country Autopilot-Desktops/Notebook-PilotDevices groups." -MailNickname 'RegionPreProductionWindows'
            If ($RegionGroup) {
                $NewRegionGroupCounter++
                $Info = "Region group '$RegionGroupName' created."
                $msg = ("$env:computername : " + (Get-Date -Format "dd-MM-yyyy HH:mm:ss K") + "(UTC) : SUCCESS : " + $Info)
                Write-host $msg  -ForegroundColor GREEN
                If ($Log) { Write-Output $msg | Out-File -LiteralPath $Logfile -Append }
            }
        } 
        else {
            $ExistingRegionGroupCounter++
            $Info = "Region group '$RegionGroupName' already exists."
            $msg = ("$env:computername : " + (Get-Date -Format "dd-MM-yyyy HH:mm:ss K") + "(UTC) : SUCCESS : " + $Info)
            Write-host $msg  -ForegroundColor GREEN
            If ($Log) { Write-Output $msg | Out-File -LiteralPath $Logfile -Append }
        }

        # Get all groups matching the child group patterns
        $childGroups = Get-Groups -Filters $RegionName -Groups $devicegroups | Where-Object { ($_.MailNickname -eq 'ZonePreProductionWindows') -and $_.Description -eq 'This is created automatically Zone group containing per-country Autopilot-Desktops/Notebook-PilotDevices groups.' }

        # Add each child group to the Region group only if it is not empty
        foreach ($childGroup in $childGroups) {
            # Check if the child group is not empty
            $childGroupMembers = Get-MgGroupMember -GroupId $childGroup.Id
            if ($childGroupMembers.Count -gt 0) {
                $isMember = Get-MgGroupMember -GroupId $RegionGroup.Id | Where-Object { $_.Id -eq $childGroup.Id }
                if (-not $isMember) {
                    New-MgGroupMember -GroupId $RegionGroup.Id -DirectoryObjectId $childGroup.Id
                    $Info = "Zone PreProduction group '$($childGroup.DisplayName)' added to Region group '$RegionGroupName'."
                    $msg = ("$env:computername : " + (Get-Date -Format "dd-MM-yyyy HH:mm:ss K") + "(UTC) : INFO    : " + $Info)
                    Write-host $msg  -ForegroundColor WHITE
                    If ($Log) { Write-Output $msg | Out-File -LiteralPath $Logfile -Append }

                } 
                else {
                    $Info = "Zone PreProduction group '$($childGroup.DisplayName)' is already a member of the Region group '$RegionGroupName'."
                    $msg = ("$env:computername : " + (Get-Date -Format "dd-MM-yyyy HH:mm:ss K") + "(UTC) : INFO    : " + $Info)
                    Write-host $msg  -ForegroundColor WHITE
                    If ($Log) { Write-Output $msg | Out-File -LiteralPath $Logfile -Append }

                }
            } 
            else {
                $Info = "Zone PreProduction group '$($childGroup.DisplayName)' is empty and will not be added to the Region group."
                $msg = ("$env:computername : " + (Get-Date -Format "dd-MM-yyyy HH:mm:ss K") + "(UTC) : INFO    : " + $Info)
                Write-host $msg  -ForegroundColor YELLOW
                If ($Log) { Write-Output $msg | Out-File -LiteralPath $Logfile -Append }
            }
        }
    }

    Write-Output "Nested group check and creation process completed."
    $Info = "Region PreProduction group check and creation process completed. New Region PreProduction groups created: $NewRegionGroupCounter / Existing Zone PreProduction groups checked/updated: $ExistingRegionGroupCounter"
    $msg = ("$env:computername : " + (Get-Date -Format "dd-MM-yyyy HH:mm:ss K") + "(UTC) : FINISH  : " + $Info)
    Write-host $msg  -ForegroundColor Black -BackgroundColor Cyan
    If ($Log) { Write-Output $msg | Out-File -LiteralPath $Logfile -Append }


}
catch {
    $msg = ("$env:computername : " + (Get-Date -Format "dd-MM-yyyy HH:mm:ss K") + '(UTC) : ERROR   : process failed, execution stopped.')
    Write-host $msg -ForegroundColor Red
    If ($Log) { Write-Output $msg | Out-File -LiteralPath $Logfile -Append }
    $msg = $_
    Write-host $msg -ForegroundColor Red
    If ($Log) { Write-Output $msg | Out-File -LiteralPath $Logfile -Append }
}