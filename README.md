# Create Intune PreProduction Groups

Automated PowerShell script to create and manage Intune preproduction pilot and nested roll-up groups using Microsoft Graph API.

## Overview

This script automates the creation and maintenance of pilot and roll-up groups for Intune preproduction environments. It discovers source Autopilot device groups, evaluates devices for compliance and activity, and creates pilot groups containing a randomized 10% sample of active, compliant devices.

## Features

- **Autopilot Device Discovery**: Automatically finds source Autopilot device groups based on configurable name filters
- **Compliance Evaluation**: Assesses each device for compliance status, management state, and recent activity
- **Pilot Group Creation**: Creates "-PreProductionDevices" pilot groups with randomized 10% device samples
- **Group Maintenance**: Removes outdated or inactive devices from existing pilot groups
- **Roll-up Groups**: Creates and manages zone-level and region-level nested group hierarchies
- **Audit Logging**: Optional CSV logging of inactive or non-compliant devices
- **Script Logging**: Full execution logging for troubleshooting and audit trails

## Prerequisites

- **PowerShell 7.0+** (or PowerShell 5.1 with compatible modules)
- **Microsoft.Graph PowerShell Module**
- **Azure AD Application** with the following permissions:
  - `Group.Create`
  - `Group.ReadWrite.All`
  - `Device.Read.All`
  - `Directory.Read.All`
  - `Directory.ReadWrite.All`

### Install Microsoft.Graph Module

```powershell
Install-Module -Name Microsoft.Graph -Scope CurrentUser
```

## Parameters

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `ClientId` | Yes | - | Azure AD application (client) ID |
| `ClientSecret` | Yes | - | Path to file containing client secret or the secret value itself |
| `TenantId` | Yes | - | Azure AD tenant ID |
| `Filters` | No | `'Autopilot-Desktops-Devices','Autopilot-Notebooks-Devices'` | Comma-separated group name patterns to search for |
| `LogInactiveDevices` | No | `.\log\InactiveDevices` | Path to store CSV files with inactive device details (leave empty to disable) |
| `Log` | No | `.\Log` | Path to store script log file (leave empty to disable) |

## Usage

### Basic Example

```powershell
.\Create_PreProduction_groups_1.5.ps1 `
    -ClientId 'your-client-id' `
    -ClientSecret 'D:\path\to\secret.txt' `
    -TenantId 'your-tenant-id' `
    -Filters "'Autopilot-Desktops-Devices','Autopilot-Notebooks-Devices'" `
    -LogInactiveDevices 'C:\TEMP\PilotGroupsCreation\InactiveDevices' `
    -Log 'C:\Temp\PilotGroupsCreation\Log'
```

### Minimal Example

```powershell
.\Create_PreProduction_groups_1.5.ps1 `
    -ClientId 'your-client-id' `
    -ClientSecret 'your-secret' `
    -TenantId 'your-tenant-id'
```

## How It Works

1. **Authentication**: Connects to Microsoft Graph using provided Azure AD credentials
2. **Group Discovery**: Searches for source groups matching specified filters
3. **Device Evaluation**: Analyzes each device for:
   - Compliance status
   - Management state
   - Last sign-in activity
4. **Sampling**: Selects randomized 10% of active, compliant devices
5. **Group Management**:
   - Creates or updates pilot groups (suffixed with "-PreProductionDevices")
   - Removes inactive devices from existing groups
   - Creates/updates zone and region-level roll-up groups
   - Maintains group nesting hierarchy
6. **Logging**: Records results and device details (if enabled)

## Output

### Groups Created

- **Pilot Groups**: `{SourceGroupName}-PreProductionDevices`
  - Contains 10% randomized sample of compliant, active devices
  
- **Zone Groups**: `{ZoneName}-PreProduction`
  - Nested pilot groups grouped by zone
  
- **Region Groups**: `{RegionName}-PreProduction`
  - Nested zone groups grouped by region

### Log Files

- **Script Log**: Records all operations, errors, and warnings
- **Inactive Devices CSV**: Lists devices that don't meet activity requirements
- **Non-Compliant Devices CSV**: Lists devices that don't meet compliance requirements

## Scheduling

To run this script automatically, schedule it as a Windows Task Scheduler task:

```powershell
# Example: Run daily at 2 AM
$trigger = New-ScheduledTaskTrigger -Daily -At 2am
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "D:\path\to\Create_PreProduction_groups_1.5.ps1 -ClientId '...' -ClientSecret '...' -TenantId '...'"
Register-ScheduledTask -TaskName "Create Intune PreProduction Groups" -Trigger $trigger -Action $action
```

## Troubleshooting

### Authentication Fails
- Verify ClientId, ClientSecret, and TenantId are correct
- Ensure the Azure AD application has required Graph permissions
- Check that the application is not disabled

### Groups Not Found
- Verify source group names match the `Filters` parameter
- Check that you have permissions to read groups
- Review the script log for details

### No Devices in Pilot Groups
- Verify source groups contain devices
- Check device compliance status in Intune
- Review inactive device logs

## Notes

- Version: 1.5
- Author: Viktor Kravtsov
- Creation Date: 2025-03-05
- Last Updated: Final design with bug fixes

## License

Specify your license here (e.g., MIT, Apache 2.0, etc.)

## Support

For issues or questions, please create an issue in this repository.
