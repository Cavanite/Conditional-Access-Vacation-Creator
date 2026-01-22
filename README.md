# Conditional Access Vacation Creator

A PowerShell-based GUI tool for creating Microsoft Entra ID (Azure AD) Conditional Access policies for users traveling internationally.

![Application Screenshot](Conditional%20Access%20Vacation%20Creator/Resources/screenshot.png)

## Overview

The **Conditional Access Vacation Creator** creates geofencing Conditional Access policies that block access from all locations *except* the vacation destination. When users travel internationally, this tool helps maintain security by restricting their access to only their travel location.

![Application Screenshot 2](Conditional%20Access%20Vacation%20Creator/Resources/screenshot2.png)


## Installation

```powershell
git clone https://github.com/Cavanite/Conditional-Access-Vacation-Creator.git
cd "Conditional Access Vacation Creator"
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
.\Run-VacationMode.ps1
```

## Usage

1. **Sign In**: Click "Sign In to Microsoft Graph" and authenticate
2. **Select Users**: The users will be shown automatically or either click "Refresh Users" and select from the list
3. **Select Destination**: Choose the vacation country from the dropdown
(or create new locations using the create countries button)
4. **Select Main Policy**: Choose your main geofencing policy to exclude users from
5. **Select the user their home country**: Choose the user there home location this will prevent user lock-out.
6. **Enter Details**:
   - Ticket Number (for tracking)
   - End Date (format: dd-mm-yyyy)
7. **Create Policy**: Click "Create CA Policy"
8. **Enable**: Go to Entra ID portal and enable the policy after review

   **Fix graph modules** : I recently came across and issue when my computer has been updates or got some new patches my graph modules were corrupted. The button Fix Graph Modules will solve this issue.
   The button wil re-install the Graph Modules needed for this application to run.

### Policy JSON Structure
How It Works

The tool creates a Conditional Access policy that:
- **Blocks** access from all locations **except** the selected vacation destination
- Is created in **disabled** state for review
- Uses Named Locations configured in your Entra ID tenant
- Optionally excludes users from your main geofencing policy


**Policy Naming:**
- Single user: `GEO-jdoe-Spain-INC123456-31-12-2026-VACATIONMODE`
---

## Author

**Bert de Zeeuw**  
Bizway BV  
📧 b.dezeeuw@bizway.nl

---

**Note:** Policies are created in disabled state. Always review and test before enabling in production.


### Troubleshooting

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

**"Authentication failed"**
Please use the button "Fix Graph Modules" this will solve most of the Authentication related issues.

### Feature updates
- Also be able to revert back the vacation mode action.
   this will include the user from the main policy and delete the Conditional Access policy created for this user.

-  The user of the script will be able to get a calendar item so they can import this in there own calendar.
   This way you will be notified when you need to revert back the changes.
