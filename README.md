# Conditional Access Vacation Creator

A PowerShell-based GUI tool for creating Microsoft Entra ID (Azure AD) Conditional Access policies for users traveling internationally.

![Application Screenshot](Resources/screenshot.png)

## Overview

The **Conditional Access Vacation Creator** creates geofencing Conditional Access policies that block access from all locations *except* the vacation destination. When users travel internationally, this tool helps maintain security by restricting their access to only their travel location.

## Features

- **Multi-User Selection**: Select one or multiple users for group vacations
- **Manual User Entry**: Add users by typing their email address
- **Named Location Integration**: Choose from your configured geofencing locations
- **Smart Filtering**: Automatically excludes admin and service accounts
- **Automatic Naming**: Policies follow format: `GEO-username-country-ticket-date-VACATIONMODE`
- **Main Policy Exclusion**: Automatically excludes users from your main geofencing policy
- **Modern GUI**: User-friendly Windows interface with status updates

## Prerequisites

- Windows 10/11 or Windows Server 2016+
- PowerShell 5.1 or higher
- Microsoft Entra ID (Azure AD) with Conditional Access licensing (P1 or higher)
- Named Locations configured in Entra ID
- Global Administrator or Conditional Access Administrator role

**Required PowerShell Modules** (auto-installed if missing):
- `Microsoft.Graph.Authentication`
- `Microsoft.Graph.Users`
- `Microsoft.Graph.Identity.SignIns`

## Installation

```powershell
git clone https://github.com/Cavanite/Conditional-Access-Vacation-Creator.git
cd "Conditional Access Vacation Creator"
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
.\Run-VacationMode.ps1
```

## Usage

1. **Sign In**: Click "Sign In to Microsoft Graph" and authenticate
2. **Select Users**: Either click "Refresh Users" and select from the list, or manually add users by typing their email
3. **Select Destination**: Choose the vacation country from the dropdown
4. **Select Main Policy** (optional): Choose your main geofencing policy to exclude users from
5. **Enter Details**: 
   - Ticket Number (for tracking)
   - End Date (format: dd-mm-yyyy)
6. **Create Policy**: Click "Create CA Policy"
7. **Enable**: Go to Entra ID portal and enable the policy after review

## ⚙️ How It Works

### Policy Architecture

```
┌─────────────────────────────────────────────────┐
│       Main Geofencing Policy                    │
│  (Blocks all countries in blocklist)            │
│                                                  │
│  Excluded Users:                                │
│    - Break-glass accounts                       │
│    - Service accounts                           │
│    - Users on vacation (auto-added) ✓          │
└─────────────────────────────────────────────────┘
                      │
                      │
                      ▼
┌─────────────────────────────────────────────────┐
│    Vacation Mode Policy (New)                   │
│  GEO-username-COUNTRY-TICKET-DATE-VACATIONMODE  │
│                                                  │
│  Conditions:                                     │
│    - Users: Selected vacation users             │
│    - Locations: All EXCEPT vacation country     │
│    - Applications: All cloud apps               │
│                                                  │
│  Control:                                        │
│    - Action: BLOCK access                       │
└─────────────────────────────────────────────────┘
```

### Technical Flow

1. **Authentication**: Connects to Microsoft Graph using OAuth 2.0
2. **Data Retrieval**: 
   - Fetches all non-admin users from Entra ID
   - Loads Named Locations (geofencing zones)
   - Retrieves existing Conditional Access policies
3. **User Selection**: Maps display names to user GUIDs
4. **Location Mapping**: Resolves location names to location GUIDs
5. **Policy Creation**:
   - Builds JSON policy object
   - Sets location condition (all locations EXCEPT vacation destination)
   - Configures block control
   - Creates policy via Graph API
6. **Exclusion Update**: Adds users to exclusion list of main policy

### Policy JSON Structure
How It Works

The tool creates a Conditional Access policy that:
- **Blocks** access from all locations **except** the selected vacation destination
- Is created in **disabled** state for review
- Uses Named Locations configured in your Entra ID tenant
- Optionally excludes users from your main geofencing policy

**Policy Naming:**
- Single user: `GEO-jdoe-Spain-INC123456-31-12-2026-VACATIONMODE`
- Multiple users: `GEO-jdoe-Plus2-Spain-INC123456-31-12-2026-VACATIONMODE`
---

#### 🔴 Modules installation fails

**Cause**: Insufficient permissions or network issues

**Solution**:
1. Run PowerShell as Administrator
2. Set execution policy: `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser`
3. Check internet connectivity
4. Manually install: `Install-Module Microsoft.Graph -Scope CurrentUser`

---

### Debug Mode

Run the script with verbose logging:

```powershell
.\Run-VacationMode.ps1 -LogLevel Debug
```

### Support Resources

- **Microsoft Entra ID Documentation**: [Conditional Access Documentation](https://learn.microsoft.com/entra/identity/conditional-access/)
- **Microsoft Graph API**: [Graph API Reference](https://learn.microsoft.com/graph/api/resources/conditionalaccesspolicy)
- **Contact Author**: b.dezeeuw@bizway.nl

## Author

**Bert de Zeeuw**  
Bizway BV  
📧 b.dezeeuw@bizway.nl

---

**Note:** Policies are created in disabled state. Always review and test before enabling in production.


Troubleshooting

**"No geofencing policy found"**  
Create a main geofencing policy with keywords like "geofence", "country", or "blocklist" in the name.

**"Failed to connect to Microsoft Graph"**  
Ensure you have Conditional Access Administrator role and complete MFA if required.

**"Could not find GUID for user"**  
Click "Refresh Users" to reload the list or verify the user exists in Entra ID.

**Date format errors**  
Use format `dd-mm-yyyy` (e.g., `31-12-2026`) with dashes, not slashes.

**Module installation fails**  
Run as Administrator and manually install: `Install-Module Microsoft.Graph -Scope CurrentUser`