# Conditional Access Vacation Creator

A PowerShell-based GUI tool for creating Microsoft Entra ID (Azure AD) Conditional Access policies to secure users during international travel.

## 📋 Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
- [How It Works](#how-it-works)
- [Policy Configuration](#policy-configuration)
- [Security Considerations](#security-considerations)
- [Troubleshooting](#troubleshooting)
- [Author](#author)

## 🎯 Overview

The **Conditional Access Vacation Creator** is a security automation tool designed to protect organizational resources when users travel internationally. It creates geofencing Conditional Access policies that block access from all locations *except* the vacation destination, ensuring that user credentials remain secure even if compromised.

### Key Benefits

- **Enhanced Security**: Restricts user access to only their vacation location
- **Automated Policy Creation**: Eliminates manual policy configuration errors
- **User-Friendly GUI**: No command-line expertise required
- **Integration with Main Geofencing**: Automatically excludes users from main blocklist policies

## ✨ Features

### 🔐 Authentication & Authorization

- **Microsoft Graph Integration**: Secure authentication using Microsoft Graph API
- **Required Permissions**: 
  - `User.Read.All` - Read user directory
  - `Policy.Read.All` - Read Conditional Access policies
  - `Policy.ReadWrite.ConditionalAccess` - Create and modify CA policies

### 👥 User Management

- **Multi-User Selection**: Select multiple users for group vacations
- **Manual User Entry**: Add users by typing their UPN (User Principal Name)
- **Smart Filtering**: Automatically excludes:
  - Administrative accounts
  - Break-glass accounts
  - Service accounts
  - External users (guests)
- **Live User Lookup**: Fetches users from Entra ID with real-time validation

### 🌍 Location-Based Controls

- **Named Location Selection**: Choose from configured geofencing locations
- **Country-Specific Policies**: Restricts access to selected vacation destination
- **Inverse Geofencing**: Blocks all locations *except* the vacation country

### 📝 Policy Details

- **Automatic Naming Convention**: Policies follow standardized format:
  ```
  GEO-<username>-<country>-<ticket>-<enddate>-VACATIONMODE
  ```
- **Policy States**: Created in **disabled** state for review before activation
- **Exclusion Management**: Automatically excludes users from main geofencing policy

### 🎨 User Interface

- **Modern WPF Interface**: Clean, professional Windows application
- **Real-Time Status Updates**: Console-style status messages
- **Connection Status Indicator**: Visual feedback for Graph API connection
- **Input Validation**: Prevents invalid configurations

## 📦 Prerequisites

### System Requirements

- **Operating System**: Windows 10/11 or Windows Server 2016+
- **PowerShell**: Version 5.1 or higher
- **.NET Framework**: 4.7.2 or higher (for WPF support)

### Azure/Microsoft 365 Requirements

- **Microsoft Entra ID** (Azure AD) tenant
- **Conditional Access** licensing (Azure AD Premium P1 or higher)
- **Named Locations** configured in Entra ID
- **Existing Geofencing Policy** (recommended)

### Required PowerShell Modules

The script will automatically install these modules if not present:

- `Microsoft.Graph.Authentication`
- `Microsoft.Graph.Users`
- `Microsoft.Graph.Identity.SignIns`

### Permissions Required

The account running the script needs:

- **Global Administrator** or **Conditional Access Administrator** role
- Permissions to create and modify Conditional Access policies
- Permissions to read user directory

## 🚀 Installation

### 1. Download the Repository

```powershell
git clone https://github.com/Cavanite/Conditional-Access-Vacation-Creator.git
cd Conditional-Access-Vacation-Creator
```

### 2. Verify File Structure

Ensure the following structure exists:

```
Conditional Access Vacation Creator/
├── Run-VacationMode.ps1          # Main script
├── Modules/                       # PowerShell modules (optional)
└── Resources/                     # Resources folder
    └── Exported/                  # Policy export location
```

### 3. Review Script Execution Policy

```powershell
# Check current execution policy
Get-ExecutionPolicy

# If restricted, set to RemoteSigned (requires admin)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## 💻 Usage

### Starting the Application

```powershell
cd "Conditional Access Vacation Creator"
.\Run-VacationMode.ps1
```

### Step-by-Step Workflow

#### 1️⃣ **Sign In to Microsoft Graph**

1. Click **"Sign In to Microsoft Graph"**
2. Authenticate with an account that has Conditional Access Administrator permissions
3. Consent to the required permissions if prompted
4. Wait for the connection status to change to **"Connected"** (green)

#### 2️⃣ **Select Users**

**Option A: From Directory**
1. Click **"Refresh Users"** to load users from Entra ID
2. Select one or multiple users from the list
3. Use **"Select All"** or **"Clear Selection"** as needed

**Option B: Manual Entry**
1. Type or paste the user's UPN (e.g., `user@domain.com`) in the input box
2. Click **"Add"** or press Enter
3. Repeat for additional users

#### 3️⃣ **Select Vacation Destination**

1. Choose the country/location from the **"Vacation Destination"** dropdown
2. This should match a Named Location configured in your Entra ID tenant

#### 4️⃣ **Select Main Geofencing Policy** (Optional but Recommended)

1. Select your organization's main geofencing/blocklist policy
2. Users will be automatically excluded from this policy to allow travel

#### 5️⃣ **Enter Policy Details**

1. **Ticket Number**: Reference number for tracking (e.g., ServiceNow ticket)
2. **End Date**: When the vacation ends (format: `dd-mm-yyyy`, e.g., `31-12-2026`)
3. **Policy Name**: Auto-generated based on inputs (read-only)

#### 6️⃣ **Create the Policy**

1. Click **"Create CA Policy"**
2. Review the confirmation dialog carefully
3. Click **"Yes"** to create the policy
4. The policy will be created in **disabled** state

#### 7️⃣ **Review and Enable**

1. Navigate to **Entra ID > Security > Conditional Access**
2. Locate the newly created policy
3. Review the configuration
4. Change state from **Disabled** to **Enabled**

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

```json
{
  "displayName": "GEO-jdoe-Spain-INC123456-31-12-2026-VACATIONMODE",
  "state": "disabled",
  "conditions": {
    "applications": {
      "includeApplications": ["All"]
    },
    "users": {
      "includeUsers": ["<user-guid>"]
    },
    "locations": {
      "includeLocations": ["All"],
      "excludeLocations": ["<spain-location-guid>"]
    }
  },
  "grantControls": {
    "operator": "OR",
    "builtInControls": ["block"]
  }
}
```

## 🔧 Policy Configuration

### Naming Convention

Policies follow this standardized format:

```
GEO-<username>-<country>-<ticket>-<enddate>-VACATIONMODE
```

**Examples:**

- Single user: `GEO-jdoe-Spain-INC123456-31-12-2026-VACATIONMODE`
- Multiple users: `GEO-jdoe-Plus2-Spain-INC123456-31-12-2026-VACATIONMODE`

**Components:**

| Component | Description | Example |
|-----------|-------------|---------|
| `GEO` | Prefix indicating geofencing policy | `GEO` |
| `username` | First user's username (before @) | `jdoe` |
| `Plus#` | Number of additional users (multi-user) | `Plus2` |
| `country` | Vacation destination | `Spain` |
| `ticket` | Reference ticket number | `INC123456` |
| `enddate` | End date (dd-mm-yyyy) | `31-12-2026` |
| `VACATIONMODE` | Policy type identifier | `VACATIONMODE` |

### Default Settings

- **State**: `disabled` (requires manual review before activation)
- **Included Applications**: All cloud applications
- **Included Locations**: All locations
- **Excluded Locations**: Selected vacation destination
- **Access Control**: Block access
- **Operator**: OR (any control satisfies requirement)

## 🔒 Security Considerations

### ⚠️ Important Security Notes

1. **Review Before Enabling**: Always review policies in the Azure Portal before enabling
2. **Break-Glass Accounts**: Ensure break-glass accounts are never included in vacation policies
3. **Policy Expiration**: Manually disable or delete policies after the vacation ends
4. **Location Accuracy**: Named Locations rely on IP geolocation, which isn't perfect
5. **VPN Consideration**: Users should not use VPNs that exit in other countries

### Best Practices

✅ **Do:**

- Create policies for users traveling internationally
- Set realistic end dates with buffer time
- Document policy creation with ticket numbers
- Review all policies weekly for expired vacations
- Test policies with a test user first
- Communicate policy activation to users

❌ **Don't:**

- Create policies without user notification
- Use for permanent location restrictions
- Include administrative or break-glass accounts
- Leave policies enabled after vacation ends
- Rely solely on geofencing for security

### Compliance Considerations

- **GDPR**: Location tracking may require user consent
- **Data Residency**: Ensure Named Locations comply with data residency requirements
- **Audit Logs**: All policy changes are logged in Entra ID audit logs
- **Documentation**: Maintain records of why policies were created

## 🛠️ Troubleshooting

### Common Issues

#### 🔴 "No geofencing policy found"

**Cause**: No existing Conditional Access policy with geofencing keywords detected

**Solution**: 
1. Create a main geofencing policy with keywords like "geofence", "country", "blocklist" in the name
2. Or contact the script author to configure your tenant

---

#### 🔴 "Failed to connect to Microsoft Graph"

**Cause**: Insufficient permissions or authentication failure

**Solution**:
1. Ensure you're using an account with Conditional Access Administrator role
2. Check if MFA is required and complete it
3. Verify network connectivity
4. Try signing out and back in

---

#### 🔴 "Could not find GUID for user"

**Cause**: User not found in Entra ID or caching issue

**Solution**:
1. Click **"Refresh Users"** to reload the list
2. Verify the user exists in Entra ID
3. Check spelling if entering manually
4. Disconnect and reconnect to Microsoft Graph

---

#### 🔴 "Date format validation error"

**Cause**: Incorrect date format entered

**Solution**: 
- Use format `dd-mm-yyyy` (e.g., `31-12-2026`)
- Ensure day, month, and year are all numeric
- Use dashes (`-`), not slashes (`/`)

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

## 👤 Author

**Bert de Zeeuw**  
Bizway BV  
📧 b.dezeeuw@bizway.nl

---

## 📄 License

This project is provided as-is for use within organizations. Please review your organization's policies before deployment.

## 🤝 Contributing

For bug reports, feature requests, or contributions, please contact the author directly.

---

**Version**: 1.0  
**Last Updated**: January 2026  
**PowerShell Version**: 5.1+


