# Conditional Access Vacation Creator

A PowerShell-based GUI tool for creating Microsoft Entra ID (Azure AD) Conditional Access policies for users traveling internationally.

![Application Screenshot](Conditional%20Access%20Vacation%20Creator/Resources/screenshot.png)

## Overview

The **Conditional Access Vacation Creator** creates geofencing Conditional Access policies that block access from all locations *except* the vacation destination. When users travel internationally, this tool helps maintain security by restricting their access to only their travel location.

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
(or create new locations with the upcoming country creation script)
4. **Select Main Policy**: Choose your main geofencing policy to exclude users from
5. **Enter Details**:
   - Ticket Number (for tracking)
   - End Date (format: dd-mm-yyyy)
6. **Create Policy**: Click "Create CA Policy"
7. **Enable**: Go to Entra ID portal and enable the policy after review

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