#Requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter()]
    [ValidateSet('Debug', 'Information', 'Warning', 'Error')]
    [string]$LogLevel = 'Information',

    [Parameter()]
    [string]$ConfigPath
)

<#
Purpose: This script will create a Conditional Access policy in Entra ID (Azure AD) to block access for users who are on vacation.
the script will ask the following:
- Which users should be included in the policy (multi-select)
- Which Country the users will be traveling to (single select)
the inputs will be exported to a JSON file so the settings can be imported in Conditional Access policies.
#>

#Script information


#######################################################################################################

#######################################################################################################
# Hide PowerShell console window
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'

$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0) | Out-Null

# Set up PowerShell runspace for MSAL (required for MSAL.NET to work properly)
if (-not [System.Management.Automation.Runspaces.Runspace]::DefaultRunspace) {
    $runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $runspace.ApartmentState = 'STA'
    $runspace.ThreadOptions = 'ReuseThread'
    $runspace.Open()
    [System.Management.Automation.Runspaces.Runspace]::DefaultRunspace = $runspace
}

$ErrorActionPreference = 'Stop'

#region Script Initialization
$ScriptRoot = $PSScriptRoot
$ModulesPath = Join-Path -Path $ScriptRoot -ChildPath 'Modules'
$ResourcesPath = Join-Path -Path $ScriptRoot -ChildPath 'Resources'


# Add required assemblies for WPF
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# Import modules
$modules = @(
    'Initialize-Named-Location'
)

foreach ($module in $modules) {
    $modulePath = Join-Path -Path $ModulesPath -ChildPath "$module.psm1"
    if (Test-Path -Path $modulePath) {
        Import-Module $modulePath -Force -DisableNameChecking
        Write-Verbose "Imported module: $module"
    }
    else {
        throw "Required module not found: $modulePath"
    }
}

# Create the main window
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Conditional Access Vacation Creator" Height="800" Width="1200"
    WindowStartupLocation="CenterScreen" Topmost="False">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <TextBlock Grid.Row="0" Text="Conditional Access Vacation Creator" 
                   FontSize="20" FontWeight="Bold" 
                   HorizontalAlignment="Left" Margin="10"/>
        
        <!-- Main Content Area -->
        <Grid Grid.Row="1" Margin="10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            
            <!-- Left Panel - User Selection and Status -->
            <GroupBox Grid.Column="0" Header="Select Users on Vacation" 
                      FontSize="14" FontWeight="Bold">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="260"/>
                    </Grid.RowDefinitions>
                    
                    <TextBlock Grid.Row="0" Text="Select Users:" 
                               Margin="5,5,5,2" FontSize="11" FontWeight="Bold"/>
                    
                    <!-- Search Box -->
                    <Grid Grid.Row="1" Margin="5,0,5,5">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox Grid.Column="0" Name="UserSearchBox" 
                                 Height="30" 
                                 Padding="8"
                                 VerticalContentAlignment="Center"
                                 FontSize="11"
                                 Text="Search users..."/>
                        <Button Grid.Column="1" Name="ClearSearchBtn" 
                                Content="X" Width="30" Height="30" 
                                Margin="5,0,0,0"
                                FontSize="14"
                                Background="#E0E0E0"
                                FontWeight="Bold"/>
                    </Grid>
                    
                    <ListBox Grid.Row="2" Name="UsersListBox" 
                             SelectionMode="Multiple"
                             Margin="5"
                             VerticalAlignment="Stretch"/>
                    
                    <Grid Grid.Row="3" Margin="5">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        
                        <StackPanel Grid.Row="0" Orientation="Horizontal">
                            <Button Name="RefreshUsersBtn" Content="Refresh Users" 
                                    Width="120" Height="30" Margin="5"/>
                            <Button Name="SelectAllUsersBtn" Content="Select All" 
                                    Width="100" Height="30" Margin="5"/>
                            <Button Name="ClearUsersBtn" Content="Clear Selection" 
                                    Width="120" Height="30" Margin="5"/>
                            <Border Name="ConnectionStatusBorder" 
                                    BorderBrush="Gray" BorderThickness="1" 
                                    CornerRadius="3" Padding="10,5" Margin="5,5,5,5"
                                    Background="#F0F0F0">
                                <TextBlock Name="ConnectionStatusText" 
                                           Text="Not Connected" 
                                           FontSize="11" FontWeight="Bold"
                                           Foreground="Gray" 
                                           VerticalAlignment="Center"/>
                            </Border>
                        </StackPanel>
                        
                        <Button Grid.Row="1" Name="SignInBtn" Content="Sign In to Microsoft Graph" 
                                Height="35" Margin="5" FontWeight="Bold"
                                Background="#0078D4" Foreground="White"/>
                        
                        <Button Grid.Row="1" Name="DisconnectBtn" Content="Disconnect" 
                                Height="35" Margin="5" FontWeight="Bold"
                                Background="#D13438" Foreground="White"
                                Visibility="Collapsed"/>
                    </Grid>
                    
                    <!-- Status Messages -->
                    <GroupBox Grid.Row="5" Header="Status" 
                              FontSize="12" FontWeight="Bold" Margin="0,5,0,0">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBox Grid.Column="0" Grid.ColumnSpan="2" Name="StatusTextBox" 
                                     IsReadOnly="True" 
                                     VerticalScrollBarVisibility="Auto"
                                     HorizontalScrollBarVisibility="Auto"
                                     FontFamily="Consolas" FontSize="10"
                                     Margin="5"
                                     TextWrapping="Wrap"
                                     Height="220"/>
                            <Button Grid.Column="1" Name="ClearStatusBtn" Content="Clear" 
                                    Width="60" Height="25" FontSize="10"
                                    Background="#E0E0E0" Margin="5,5,5,0"
                                    VerticalAlignment="Top"/>
                        </Grid>
                    </GroupBox>
                </Grid>
            </GroupBox>
            
            <!-- Right Panel - Country and Policy Details -->
            <Grid Grid.Column="2">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                
                <!-- Country Selection -->
                <GroupBox Grid.Row="0" Header="Vacation Destination" 
                          FontSize="14" FontWeight="Bold" Margin="0,0,0,10">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        
                        <TextBlock Grid.Row="0" Text="Select the country the user(s) will be traveling to:" 
                                   Margin="5" TextWrapping="Wrap"/>
                        
                        <Grid Grid.Row="1">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            
                            <ComboBox Grid.Column="0" Name="CountryComboBox" 
                                      Margin="5" Height="30"
                                      IsEditable="True"
                                      IsTextSearchEnabled="True"/>
                            
                            <Button Grid.Column="1" Name="RefreshCountriesBtn" 
                                    Content="RF" Width="35" Height="30" Margin="0,5,5,5"
                                    ToolTip="Refresh country list"
                                    FontSize="16" Padding="0"/>
                        </Grid>
                    </Grid>
                </GroupBox>
                
                <!-- User's Current Location -->
                <GroupBox Grid.Row="1" Header="User's Current Location (Home Country)" 
                          FontSize="14" FontWeight="Bold" Margin="0,0,0,10">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        
                        <TextBlock Grid.Row="0" Text="Select the user's current location (to avoid blocking them before they leave):" 
                                   Margin="5" TextWrapping="Wrap" FontSize="11"/>
                        
                        <ComboBox Grid.Row="1" Name="UserCurrentLocationComboBox" 
                                  Margin="5" Height="30"
                                  IsEditable="True"
                                  IsTextSearchEnabled="True"/>
                    </Grid>
                </GroupBox>
                
                <!-- Existing Geofencing Policy -->
                <GroupBox Grid.Row="2" Header="Main Geofencing Policy" 
                          FontSize="14" FontWeight="Bold" Margin="0,0,0,10">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        
                        <TextBlock Grid.Row="0" Text="Select the main CA policy to exclude users from:" 
                                   Margin="5" TextWrapping="Wrap"/>
                        
                        <ComboBox Grid.Row="1" Name="ExistingPolicyComboBox" 
                                  Margin="5" Height="30"
                                  IsEditable="True"
                                  IsTextSearchEnabled="True"/>
                    </Grid>
                </GroupBox>
                
                <!-- Policy Details -->
                <GroupBox Grid.Row="3" Header="Policy Details" 
                          FontSize="14" FontWeight="Bold" Margin="0,0,0,10">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        
                        <TextBlock Grid.Row="0" Text="Ticket Number: *" Margin="5,5,5,2" Foreground="#D13438"/>
                        <TextBox Grid.Row="1" Name="TicketNumberTextBox" 
                                 Margin="5,0,5,5" Height="25"/>
                        
                        <TextBlock Grid.Row="2" Text="End Date (dd-mm-yyyy): *" Margin="5,5,5,2" Foreground="#D13438"/>
                        <TextBox Grid.Row="3" Name="EndDateTextBox" 
                                 Margin="5,0,5,5" Height="25"/>
                        
                        <TextBlock Grid.Row="4" Text="Policy Name:" Margin="5,5,5,2"/>
                        <TextBox Grid.Row="5" Name="PolicyNameTextBox" 
                                 Margin="5,0,5,5" Height="25" IsReadOnly="True"
                                 Background="#F0F0F0"/>
                    </Grid>
                </GroupBox>
            </Grid>
        </Grid>
        
        <!-- Action Buttons and Contact Info -->
        <Grid Grid.Row="2" Margin="10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            
            <!-- Contact Information -->
            <StackPanel Grid.Column="0" VerticalAlignment="Center" Orientation="Horizontal">
                <StackPanel VerticalAlignment="Center" Margin="0,0,5,0">
                    <TextBlock Text="Written by: Bert de Zeeuw - Bizway BV" FontSize="11" FontWeight="Bold"/>
                    <Button Name="EmailBtn" Background="Transparent" BorderThickness="0" 
                            Padding="0" Cursor="Hand" HorizontalAlignment="Left"
                            ToolTip="Send email to b.dezeeuw@bizway.nl">
                        <TextBlock Text="Email: b.dezeeuw@bizway.nl" FontSize="10" Foreground="#0078D4" TextDecorations="Underline"/>
                    </Button>
                </StackPanel>
                <Button Name="GitHubBtn" Content="GitHub" ToolTip="Visit GitHub Profile"
                        Width="75" Height="30" Margin="5,0,10,0"
                        FontSize="10" Background="#24292e" Foreground="White"
                        BorderBrush="#444d56" Cursor="Hand"/>
                <Button Name="ExcludeCountriesBtn" Content="Create Countries" 
                        Width="140" Height="35" Margin="5"
                        FontWeight="Bold" Background="#0078D4" Foreground="White"/>
            </StackPanel>
            
            <!-- Action Buttons -->
            <StackPanel Grid.Column="1" Orientation="Horizontal">
                <Button Name="FixGraphModulesBtn" Content="Fix Graph Modules" 
                        Width="150" Height="35" Margin="5"
                        FontWeight="Bold" Background="#FF6B00" Foreground="White"
                        ToolTip="Uninstall and reinstall Microsoft Graph modules"/>
                <Button Name="CreatePolicyBtn" Content="Create CA Policy" 
                        Width="140" Height="35" Margin="5"
                        FontWeight="Bold"/>
                <Button Name="CloseBtn" Content="Close" 
                        Width="100" Height="35" Margin="5"/>
            </StackPanel>
        </Grid>
    </Grid>
</Window>
"@

# Parse XAML and create window
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# Get references to UI elements
$UsersListBox = $window.FindName("UsersListBox")
$UserSearchBox = $window.FindName("UserSearchBox")
$ClearSearchBtn = $window.FindName("ClearSearchBtn")
$CountryComboBox = $window.FindName("CountryComboBox")
$UserCurrentLocationComboBox = $window.FindName("UserCurrentLocationComboBox")
$ExistingPolicyComboBox = $window.FindName("ExistingPolicyComboBox")
$TicketNumberTextBox = $window.FindName("TicketNumberTextBox")
$EndDateTextBox = $window.FindName("EndDateTextBox")
$PolicyNameTextBox = $window.FindName("PolicyNameTextBox")
$PolicyDescriptionTextBox = $window.FindName("PolicyDescriptionTextBox")
$StatusTextBox = $window.FindName("StatusTextBox")
$RefreshUsersBtn = $window.FindName("RefreshUsersBtn")
$SelectAllUsersBtn = $window.FindName("SelectAllUsersBtn")
$ClearUsersBtn = $window.FindName("ClearUsersBtn")
$CreatePolicyBtn = $window.FindName("CreatePolicyBtn")
$CloseBtn = $window.FindName("CloseBtn")
$SignInBtn = $window.FindName("SignInBtn")
$DisconnectBtn = $window.FindName("DisconnectBtn")
$ConnectionStatusBorder = $window.FindName("ConnectionStatusBorder")
$ConnectionStatusText = $window.FindName("ConnectionStatusText")
$ExcludeCountriesBtn = $window.FindName("ExcludeCountriesBtn")
$RefreshCountriesBtn = $window.FindName("RefreshCountriesBtn")
$ClearStatusBtn = $window.FindName("ClearStatusBtn")
$GitHubBtn = $window.FindName("GitHubBtn")
$EmailBtn = $window.FindName("EmailBtn")
$FixGraphModulesBtn = $window.FindName("FixGraphModulesBtn")

# Global variable to track Graph connection status
$script:GraphConnected = $false
$script:UserCache = @{}
$script:NamedLocationsCache = @{}
$script:CAPoliciesCache = @{}

# Disable country combobox until signed in
$CountryComboBox.IsEnabled = $false
$ExistingPolicyComboBox.IsEnabled = $false

# Function to check and install Microsoft.Graph module
function Ensure-GraphModule {
    Add-StatusMessage "Checking Microsoft Graph modules..."
    
    # Check if we can load the modules
    $loadError = $false
    try {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        Import-Module Microsoft.Graph.Users -ErrorAction Stop
        Import-Module Microsoft.Graph.Identity.SignIns -ErrorAction Stop
        Add-StatusMessage "Graph modules loaded successfully"
        return
    } catch {
        $loadError = $true
        Add-StatusMessage "ERROR: Cannot load Graph modules - $($_.Exception.Message)"
    }
    
    if ($loadError) {
        Add-StatusMessage "Attempting to fix module installation..."
        Add-StatusMessage "This may take a few minutes..."
        
        try {
            # Uninstall all Graph modules to clean up
            Add-StatusMessage "Removing old Graph modules..."
            $graphModules = Get-Module -ListAvailable -Name Microsoft.Graph.* | Select-Object -ExpandProperty Name -Unique
            foreach ($mod in $graphModules) {
                try {
                    Uninstall-Module -Name $mod -AllVersions -Force -ErrorAction SilentlyContinue
                } catch { }
            }
            
            # Install fresh versions
            Add-StatusMessage "Installing fresh Graph modules..."
            $requiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Users', 'Microsoft.Graph.Identity.SignIns')
            
            foreach ($moduleName in $requiredModules) {
                Add-StatusMessage "  Installing $moduleName..."
                Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber -Repository PSGallery -SkipPublisherCheck
            }
            
            # Import the newly installed modules
            Add-StatusMessage "Loading newly installed modules..."
            Import-Module Microsoft.Graph.Authentication -Force
            Import-Module Microsoft.Graph.Users -Force
            Import-Module Microsoft.Graph.Identity.SignIns -Force
            
            Add-StatusMessage "SUCCESS: Graph modules installed and loaded"
        } catch {
            Add-StatusMessage "ERROR: Failed to fix modules - $($_.Exception.Message)"
            Add-StatusMessage ""
            Add-StatusMessage "MANUAL FIX REQUIRED:"
            Add-StatusMessage "1. Close this application"
            Add-StatusMessage "2. Open PowerShell as Administrator"
            Add-StatusMessage "3. Run these commands:"
            Add-StatusMessage "   Get-Module -ListAvailable Microsoft.Graph.* | Uninstall-Module -Force"
            Add-StatusMessage "   Install-Module Microsoft.Graph -Scope CurrentUser -Force"
            Add-StatusMessage "4. Restart this application"
            throw
        }
    }
}

# Add status message helper
function Add-StatusMessage {
    param([string]$Message)
    $timestamp = Get-Date -Format "HH:mm:ss"
    $StatusTextBox.AppendText("[$timestamp] $Message`r`n")
    $StatusTextBox.ScrollToEnd()
}

# Function to update policy name based on selected users
function Update-PolicyName {
    $selectedUsers = $UsersListBox.SelectedItems
    $selectedCountry = $CountryComboBox.SelectedItem
    $ticketNumber = $TicketNumberTextBox.Text.Trim()
    $endDate = $EndDateTextBox.Text.Trim()
    
    # Set placeholders if empty
    if ([string]::IsNullOrWhiteSpace($ticketNumber)) { $ticketNumber = "TICKETNUMBER" }
    if ([string]::IsNullOrWhiteSpace($endDate)) { $endDate = "ENDDATE" }
    if ([string]::IsNullOrWhiteSpace($selectedCountry)) { $selectedCountry = "COUNTRY" }
    
    if ($selectedUsers.Count -eq 0) {
        $PolicyNameTextBox.Text = "GEO-USERNAME-$selectedCountry-$ticketNumber-$endDate-VACATIONMODE"
    }
    elseif ($selectedUsers.Count -eq 1) {
        # Extract username from display format "Name (upn)"
        $userText = $selectedUsers[0]
        if ($userText -match '\((.+?)\)') {
            $username = ($matches[1] -split '@')[0]
        } else {
            $username = ($userText -split '@')[0]
        }
        $PolicyNameTextBox.Text = "GEO-$username-$selectedCountry-$ticketNumber-$endDate-VACATIONMODE"
    }
    else {
        # Multiple users - use first username
        $userText = $selectedUsers[0]
        if ($userText -match '\((.+?)\)') {
            $username = ($matches[1] -split '@')[0]
        } else {
            $username = ($userText -split '@')[0]
        }
        $PolicyNameTextBox.Text = "GEO-$username-Plus$($selectedUsers.Count - 1)-$selectedCountry-$ticketNumber-$endDate-VACATIONMODE"
    }
}

Add-StatusMessage "Application started. Please select users and destination country."

# Set default policy name
$PolicyNameTextBox.Text = "GEO-USERNAME-COUNTRY-TICKETNUMBER-dd-mm-yyyy-VACATIONMODE"

# Disable Refresh Users button until signed in
$RefreshUsersBtn.IsEnabled = $false

# Button Event Handlers
# Update policy name when selection changes
$UsersListBox.Add_SelectionChanged({
    Update-PolicyName
})

# Update policy name when country changes
$CountryComboBox.Add_SelectionChanged({
    Update-PolicyName
})

# Update policy name when ticket number changes
$TicketNumberTextBox.Add_TextChanged({
    Update-PolicyName
})

# Update policy name when end date changes
$EndDateTextBox.Add_TextChanged({
    Update-PolicyName
})

$SignInBtn.Add_Click({
    try {
        Add-StatusMessage "Checking Microsoft Graph modules..."
        Ensure-GraphModule
        
        Add-StatusMessage "Connecting to Microsoft Graph..."
        Connect-MgGraph -Scopes "User.Read.All", "Policy.Read.All", "Policy.ReadWrite.ConditionalAccess" -ContextScope Process -NoWelcome
        
        $context = Get-MgContext
        if ($context) {
            $script:GraphConnected = $true
            Add-StatusMessage "SUCCESS: Connected as $($context.Account)"
            
            # Hide Sign In button and show Disconnect button
            $SignInBtn.Visibility = "Collapsed"
            $DisconnectBtn.Visibility = "Visible"
            
            # Update connection status
            $ConnectionStatusText.Text = "Connected"
            $ConnectionStatusText.Foreground = "White"
            $ConnectionStatusBorder.Background = "#107C10"
            $ConnectionStatusBorder.BorderBrush = "#107C10"
            
            $RefreshUsersBtn.IsEnabled = $true
            
            # Fetch Named Locations
            Add-StatusMessage "Loading named locations..."
            try {
                $namedLocations = Get-MgIdentityConditionalAccessNamedLocation -All -ErrorAction Stop
                
                # Clear the cache and repopulate
                $script:NamedLocationsCache = @{}
                
                $CountryComboBox.Items.Clear()
                $UserCurrentLocationComboBox.Items.Clear()
                foreach ($location in $namedLocations) {
                    # Store location ID in cache with display name as key
                    $script:NamedLocationsCache[$location.DisplayName] = $location.Id
                    $CountryComboBox.Items.Add($location.DisplayName) | Out-Null
                    $UserCurrentLocationComboBox.Items.Add($location.DisplayName) | Out-Null
                }
                
                $CountryComboBox.IsEnabled = $true
                $UserCurrentLocationComboBox.IsEnabled = $true
                Add-StatusMessage "SUCCESS: Loaded $($namedLocations.Count) named locations"
            } catch {
                $errorMsg = $_.Exception.Message
                Add-StatusMessage "WARNING: Failed to load named locations - $errorMsg"
            }
            
            # Fetch Conditional Access Policies
            Add-StatusMessage "Loading Conditional Access policies..."
            try {
                $caPolicies = Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop
                
                # Check for geofencing policies - strict filter: must contain "Geofencing" or "Countrywhitelist"
                $geofencingPolicies = $caPolicies | Where-Object {
                    $_.DisplayName -like '*Geofenc*' -or 
                    $_.DisplayName -like '*Countrywhitelist*'
                }
                
                if ($geofencingPolicies.Count -eq 0) {
                    Add-StatusMessage "WARNING: No geofencing policy found!"
                } else {
                    # Clear the cache and repopulate with only geofencing policies
                    $script:CAPoliciesCache = @{}
                    
                    $ExistingPolicyComboBox.Items.Clear()
                    foreach ($policy in $geofencingPolicies) {
                        # Store policy ID in cache with display name as key
                        $script:CAPoliciesCache[$policy.DisplayName] = $policy.Id
                        $ExistingPolicyComboBox.Items.Add($policy.DisplayName) | Out-Null
                    }
                    
                    $ExistingPolicyComboBox.IsEnabled = $true
                    Add-StatusMessage "SUCCESS: Loaded $($geofencingPolicies.Count) geofencing policies"
                }
                
                # Automatically load users after successful connection
                Add-StatusMessage "Automatically loading users..."
                $RefreshUsersBtn.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
            } catch {
                $errorMsg = $_.Exception.Message
                $statusCode = $_.Exception.Response.StatusCode.value__
                $innerMsg = $_.Exception.InnerException.Message
                
                Add-StatusMessage "ERROR: Failed to load CA policies"
                Add-StatusMessage "  Error: $errorMsg"
                if ($innerMsg) { Add-StatusMessage "  Details: $innerMsg" }
                if ($statusCode) { Add-StatusMessage "  Status Code: $statusCode" }
                
                # Check if it's a licensing issue
                if ($errorMsg -match "Premium|license|subscription|does not have") {
                    Add-StatusMessage "  >>> This tenant requires Azure AD Premium P1 or P2 for Conditional Access"
                }
            }
        }
    } catch {
        Add-StatusMessage "ERROR: Sign in failed - $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Failed to connect to Microsoft Graph: $($_.Exception.Message)", "Sign In Error", "OK", "Error")
    }
})

$RefreshUsersBtn.Add_Click({
    if (-not $script:GraphConnected) {
        Add-StatusMessage "ERROR: Please sign in to Microsoft Graph first."
        [System.Windows.MessageBox]::Show("Please sign in to Microsoft Graph first.", "Not Connected", "OK", "Warning")
        return
    }
    
    try {
        Add-StatusMessage "Fetching users from Entra ID..."
        $users = Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,AssignedLicenses | Select-Object Id,DisplayName,UserPrincipalName,AssignedLicenses
        
        # Filter patterns for admin and breakglass accounts
        $excludePatterns = @(
            '*admin*',
            '*administrator*',
            '*breakglass*',
            '*break-glass*',
            '*break.glass*',
            '*emergency*',
            '*emerg*',
            '*privileged*',
            '*service*',
            '*svc*',
            '*system*'
        )
        
        # Admin role patterns to exclude
        $adminRolePatterns = @(
            '*admin*',
            '*administrator*',
            '*privileged*',
            '*global*',
            '*security*'
        )
        
        # Clear and populate the cache as a hashtable
        $script:UserCache = @{}
        $UsersListBox.Items.Clear()
        
        $filteredCount = 0
        foreach ($user in $users) {
            # Check if user is external (contains #EXT# in UPN)
            $isExternal = $user.UserPrincipalName -like '*#EXT#*'
            
            # Check if user has a valid license assigned
            $hasLicense = $null -ne $user.AssignedLicenses -and $user.AssignedLicenses.Count -gt 0
            
            # Check if user matches any exclusion pattern
            $shouldExclude = $false
            foreach ($pattern in $excludePatterns) {
                if (($user.DisplayName -like $pattern) -or ($user.UserPrincipalName -like $pattern)) {
                    $shouldExclude = $true
                    break
                }
            }
            
            # Check if user has admin roles
            $hasAdminRole = $false
            if (-not $shouldExclude) {
                try {
                    $userRoles = Get-MgUserMemberOf -UserId $user.Id -All -ErrorAction SilentlyContinue
                    foreach ($role in $userRoles) {
                        $roleName = $role.AdditionalProperties.displayName
                        foreach ($pattern in $adminRolePatterns) {
                            if ($roleName -like $pattern) {
                                $hasAdminRole = $true
                                break
                            }
                        }
                        if ($hasAdminRole) { break }
                    }
                } catch {
                    # Silent fail if role retrieval fails
                }
            }
            
            # Exclude if external, no license, matches exclusion pattern, or has admin role
            if ($isExternal -or -not $hasLicense -or $shouldExclude -or $hasAdminRole) {
                $filteredCount++
            } else {
                # Only add internal, licensed, non-admin/non-breakglass users
                $displayText = "$($user.DisplayName) ($($user.UserPrincipalName))"
                # Store GUID with display text as key
                $script:UserCache[$displayText] = $user.Id
                $UsersListBox.Items.Add($displayText) | Out-Null
            }
        }
        
        Add-StatusMessage "SUCCESS: Loaded $($UsersListBox.Items.Count) users from Entra ID ($filteredCount admin/service/external accounts filtered)"
    } catch {
        Add-StatusMessage "ERROR: Failed to fetch users - $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Failed to fetch users: $($_.Exception.Message)", "Error", "OK", "Error")
    }
})

# User Search Functionality
$UserSearchBox.Add_TextChanged({
    $searchTerm = $UserSearchBox.Text.Trim()
    
    if ([string]::IsNullOrWhiteSpace($searchTerm) -or $searchTerm -eq "Search users...") {
        # Show all users
        $UsersListBox.Items.Clear()
        foreach ($user in ($script:UserCache.Keys | Sort-Object)) {
            $UsersListBox.Items.Add($user) | Out-Null
        }
    } else {
        # Filter users based on search term (case-insensitive)
        $UsersListBox.Items.Clear()
        foreach ($user in ($script:UserCache.Keys | Sort-Object)) {
            if ($user -like "*$searchTerm*") {
                $UsersListBox.Items.Add($user) | Out-Null
            }
        }
    }
})

# Clear Search Button
$ClearSearchBtn.Add_Click({
    $UserSearchBox.Text = "Search users..."
    $UsersListBox.Items.Clear()
    foreach ($user in ($script:UserCache.Keys | Sort-Object)) {
        $UsersListBox.Items.Add($user) | Out-Null
    }
})

# Handle focus for search box placeholder
$UserSearchBox.Add_GotFocus({
    if ($UserSearchBox.Text -eq "Search users...") {
        $UserSearchBox.Text = ""
        $UserSearchBox.Foreground = "Black"
        $UserSearchBox.FontStyle = "Normal"
    }
})

$UserSearchBox.Add_LostFocus({
    if ([string]::IsNullOrWhiteSpace($UserSearchBox.Text)) {
        $UserSearchBox.Text = "Search users..."
        $UserSearchBox.Foreground = "#999999"
        $UserSearchBox.FontStyle = "Italic"
    }
})


$SelectAllUsersBtn.Add_Click({
    $UsersListBox.SelectAll()
    Add-StatusMessage "All users selected."
})

$ClearUsersBtn.Add_Click({
    $UsersListBox.UnselectAll()
    Add-StatusMessage "Selection cleared."
})

$RefreshCountriesBtn.Add_Click({
    if (-not $script:GraphConnected) {
        Add-StatusMessage "ERROR: Please sign in to Microsoft Graph first."
        [System.Windows.MessageBox]::Show("Please sign in to Microsoft Graph first.", "Not Connected", "OK", "Warning")
        return
    }
    
    try {
        Add-StatusMessage "Refreshing country list..."
        $namedLocations = Get-MgIdentityConditionalAccessNamedLocation -All -ErrorAction Stop
        
        # Clear the cache and repopulate
        $script:NamedLocationsCache = @{}
        $CountryComboBox.Items.Clear()
        
        foreach ($location in $namedLocations) {
            $script:NamedLocationsCache[$location.DisplayName] = $location.Id
            $CountryComboBox.Items.Add($location.DisplayName) | Out-Null
        }
        
        Add-StatusMessage "SUCCESS: Refreshed country list ($($namedLocations.Count) countries available)"
    } catch {
        Add-StatusMessage "ERROR: Failed to refresh country list - $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Failed to refresh country list: $($_.Exception.Message)", "Error", "OK", "Error")
    }
})

$CreatePolicyBtn.Add_Click({
    try {
        # Validate required fields
        $selectedUsers = $UsersListBox.SelectedItems
        $selectedCountry = $CountryComboBox.SelectedItem
        $ticketNumber = $TicketNumberTextBox.Text.Trim()
        $endDate = $EndDateTextBox.Text.Trim()
        $policyName = $PolicyNameTextBox.Text.Trim()
        
        # Validate user selection
        if ($selectedUsers.Count -eq 0) {
            [System.Windows.MessageBox]::Show("Please select at least one user.", "Validation Error", "OK", "Warning")
            return
        }
        
        # Validate country selection
        if ([string]::IsNullOrWhiteSpace($selectedCountry)) {
            [System.Windows.MessageBox]::Show("Please select a vacation destination (Named Location).", "Validation Error", "OK", "Warning")
            return
        }
        
        # Validate ticket number
        if ([string]::IsNullOrWhiteSpace($ticketNumber)) {
            [System.Windows.MessageBox]::Show("Please enter a ticket number.", "Validation Error", "OK", "Warning")
            return
        }
        
        # Validate end date
        if ([string]::IsNullOrWhiteSpace($endDate)) {
            [System.Windows.MessageBox]::Show("Please enter an end date (dd-mm-yyyy).", "Validation Error", "OK", "Warning")
            return
        }
        
        # Validate date format (dd-mm-yyyy)
        if ($endDate -notmatch '^\d{2}-\d{2}-\d{4}$') {
            [System.Windows.MessageBox]::Show("Invalid date format. Please use dd-mm-yyyy format (e.g., 31-12-2026).", "Validation Error", "OK", "Warning")
            return
        }
        
        # Validate existing policy selection
        $selectedExistingPolicy = $ExistingPolicyComboBox.SelectedItem
        if ([string]::IsNullOrWhiteSpace($selectedExistingPolicy)) {
            $result = [System.Windows.MessageBox]::Show("No main geofencing policy selected. Users will NOT be excluded from any existing policy.`n`nDo you want to continue?", "Warning", "YesNo", "Warning")
            if ($result -ne "Yes") {
                return
            }
        }
        
        # Validate user's current location selection
        $selectedUserCurrentLocation = $UserCurrentLocationComboBox.SelectedItem
        if ([string]::IsNullOrWhiteSpace($selectedUserCurrentLocation)) {
            [System.Windows.MessageBox]::Show("Please select the user's current location to prevent blocking them in their home country.", "Validation Error", "OK", "Warning")
            return
        }
        
        # Check Graph connection
        if (-not $script:graphConnected) {
            [System.Windows.MessageBox]::Show("Please sign in to Microsoft Graph first.", "Authentication Required", "OK", "Warning")
            return
        }
        
        # Get location IDs from cache
        $vacationLocationId = $null
        if ($script:namedLocationsCache.ContainsKey($selectedCountry)) {
            $vacationLocationId = $script:namedLocationsCache[$selectedCountry]
        }
        
        if ([string]::IsNullOrWhiteSpace($vacationLocationId)) {
            Add-StatusMessage "ERROR: Could not find location ID for: $selectedCountry"
            [System.Windows.MessageBox]::Show("Could not find location ID for selected vacation country.", "Error", "OK", "Error")
            return
        }
        
        # Get user's current location ID
        $userLocationId = $null
        if ($script:namedLocationsCache.ContainsKey($selectedUserCurrentLocation)) {
            $userLocationId = $script:namedLocationsCache[$selectedUserCurrentLocation]
        }
        
        if ([string]::IsNullOrWhiteSpace($userLocationId)) {
            Add-StatusMessage "ERROR: Could not find location ID for: $selectedUserCurrentLocation"
            [System.Windows.MessageBox]::Show("Could not find location ID for user's current location.", "Error", "OK", "Error")
            return
        }
        
        # Build confirmation message
        $userList = $selectedUsers -join "`n  - "
        $confirmMessage = @"
Are you sure you want to create this Conditional Access Policy?

Policy Name: $policyName
Ticket Number: $ticketNumber
End Date: $endDate

Users ($($selectedUsers.Count)):
  - $userList

Vacation Location: $selectedCountry
User's Current Location: $selectedUserCurrentLocation

This policy will:
- BLOCK access from all locations EXCEPT:
  * The vacation destination ($selectedCountry)
  * The user's current location ($selectedUserCurrentLocation)
- Be created in DISABLED state for review
- Require manual enablement after verification

Do you want to proceed?
"@
        
        # Show confirmation dialog
        $result = [System.Windows.MessageBox]::Show($confirmMessage, "Confirm Policy Creation", "YesNo", "Question")
        
        if ($result -ne "Yes") {
            Add-StatusMessage "Policy creation cancelled by user."
            return
        }
        
        Add-StatusMessage "Creating Conditional Access policy..."
        
        # Map user display names to GUIDs
        $userGuids = @()
        foreach ($userDisplay in $selectedUsers) {
            if ($script:UserCache.ContainsKey($userDisplay)) {
                $userGuids += $script:UserCache[$userDisplay]
            } else {
                Add-StatusMessage "WARNING: Could not find GUID for user: $userDisplay"
            }
        }
        
        if ($userGuids.Count -eq 0) {
            Add-StatusMessage "ERROR: No valid user GUIDs found."
            [System.Windows.MessageBox]::Show("Could not find GUIDs for selected users. Please refresh the user list.", "Error", "OK", "Error")
            return
        }
        
        # Build the policy object
        $policyObject = @{
            "displayName" = $policyName
            "state" = "disabled"
            "conditions" = @{
                "applications" = @{
                    "includeApplications" = @("All")
                    "excludeApplications" = @()
                }
                "users" = @{
                    "includeUsers" = $userGuids
                    "excludeUsers" = @()
                    "includeGroups" = @()
                    "excludeGroups" = @()
                }
                "locations" = @{
                    "includeLocations" = @("All")
                    "excludeLocations" = @($vacationLocationId, $userLocationId)
                }
            }
            "grantControls" = @{
                "operator" = "OR"
                "builtInControls" = @("block")
            }
        }
        
        # Create the policy using Microsoft Graph
        $policyJson = $policyObject | ConvertTo-Json -Depth 10
        
        Add-StatusMessage "Sending policy creation request to Microsoft Graph..."
        
        $newPolicy = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies" -Body $policyJson -ContentType "application/json"
        
        Add-StatusMessage "SUCCESS: Conditional Access policy created!"
        Add-StatusMessage "Policy ID: $($newPolicy.id)"
        Add-StatusMessage "Policy Name: $($newPolicy.displayName)"
        Add-StatusMessage "State: $($newPolicy.state) (remember to enable after review)"
        
        # Update existing geofencing policy to exclude these users
        if (-not [string]::IsNullOrWhiteSpace($selectedExistingPolicy)) {
            try {
                Add-StatusMessage "Updating main geofencing policy to exclude vacation users..."
                
                # Get existing policy ID from cache
                $existingPolicyId = $null
                if ($script:CAPoliciesCache.ContainsKey($selectedExistingPolicy)) {
                    $existingPolicyId = $script:CAPoliciesCache[$selectedExistingPolicy]
                }
                
                if ($existingPolicyId) {
                    # Fetch current policy
                    $currentPolicy = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies/$existingPolicyId"
                    
                    # Get current excluded users
                    $currentExcludedUsers = @()
                    if ($currentPolicy.conditions.users.excludeUsers) {
                        $currentExcludedUsers = $currentPolicy.conditions.users.excludeUsers
                    }
                    
                    # Add new users to exclusion list (avoid duplicates)
                    $updatedExcludedUsers = @($currentExcludedUsers)
                    foreach ($guid in $userGuids) {
                        if ($guid -notin $updatedExcludedUsers) {
                            $updatedExcludedUsers += $guid
                        }
                    }
                    
                    # Update the policy
                    $updateBody = @{
                        "conditions" = @{
                            "users" = @{
                                "includeUsers" = $currentPolicy.conditions.users.includeUsers
                                "excludeUsers" = $updatedExcludedUsers
                                "includeGroups" = $currentPolicy.conditions.users.includeGroups
                                "excludeGroups" = $currentPolicy.conditions.users.excludeGroups
                            }
                            "applications" = $currentPolicy.conditions.applications
                            "locations" = $currentPolicy.conditions.locations
                            "platforms" = $currentPolicy.conditions.platforms
                            "signInRiskLevels" = $currentPolicy.conditions.signInRiskLevels
                            "userRiskLevels" = $currentPolicy.conditions.userRiskLevels
                            "clientAppTypes" = $currentPolicy.conditions.clientAppTypes
                        }
                        "grantControls" = $currentPolicy.grantControls
                        "sessionControls" = $currentPolicy.sessionControls
                        "state" = $currentPolicy.state
                    }
                    
                    $updateJson = $updateBody | ConvertTo-Json -Depth 10
                    Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies/$existingPolicyId" -Body $updateJson -ContentType "application/json"
                    
                    Add-StatusMessage "SUCCESS: Updated '$selectedExistingPolicy' to exclude vacation users"
                } else {
                    Add-StatusMessage "WARNING: Could not find policy ID for '$selectedExistingPolicy'"
                }
            } catch {
                Add-StatusMessage "ERROR: Failed to update existing policy - $($_.Exception.Message)"
                [System.Windows.MessageBox]::Show("Vacation policy created but failed to update main policy:`n`n$($_.Exception.Message)", "Partial Success", "OK", "Warning")
            }
        }
        
        # Show success message
        $successMsg = "Conditional Access policy created successfully!`n`nPolicy Name: $policyName`nPolicy ID: $($newPolicy.id)`nState: disabled`n`n"
        
        if (-not [string]::IsNullOrWhiteSpace($selectedExistingPolicy)) {
            $successMsg += "Main policy '$selectedExistingPolicy' updated to exclude vacation users.`n`n"
        }
        
        $successMsg += "Please review and enable the policy in the Azure Portal."
        
        [System.Windows.MessageBox]::Show($successMsg, "Success", "OK", "Information")
        
    } catch {
        Add-StatusMessage "ERROR: Policy creation failed - $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Failed to create policy:`n`n$($_.Exception.Message)", "Creation Error", "OK", "Error")
    }
})

$DisconnectBtn.Add_Click({
    try {
        Add-StatusMessage "Disconnecting from Microsoft Graph..."
        
        # Disconnect from Graph
        Disconnect-MgGraph | Out-Null
        
        # Reset connection state
        $script:GraphConnected = $false
        
        # Clear all caches
        $script:UserCache = @{}
        $script:NamedLocationsCache = @{}
        $script:CAPoliciesCache = @{}
        
        # Clear UI elements
        $UsersListBox.Items.Clear()
        $CountryComboBox.Items.Clear()
        $ExistingPolicyComboBox.Items.Clear()
        
        # Reset UI state
        $SignInBtn.Visibility = "Visible"
        $DisconnectBtn.Visibility = "Collapsed"
        $RefreshUsersBtn.IsEnabled = $false
        $CountryComboBox.IsEnabled = $false
        $ExistingPolicyComboBox.IsEnabled = $false
        $CreatePolicyBtn.IsEnabled = $true
        
        # Update connection status
        $ConnectionStatusText.Text = "Not Connected"
        $ConnectionStatusText.Foreground = "Gray"
        $ConnectionStatusBorder.Background = "#F0F0F0"
        $ConnectionStatusBorder.BorderBrush = "Gray"
        
        Add-StatusMessage "SUCCESS: Disconnected from Microsoft Graph"
        Add-StatusMessage "You can now sign in with a different account."
        
    } catch {
        Add-StatusMessage "ERROR: Disconnect failed - $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Failed to disconnect: $($_.Exception.Message)", "Disconnect Error", "OK", "Error")
    }
})

$CloseBtn.Add_Click({
    # Disconnect from Microsoft Graph if connected
    if ($script:GraphConnected) {
        try {
            Disconnect-MgGraph | Out-Null
            Add-StatusMessage "Disconnected from Microsoft Graph"
        } catch {
            # Ignore errors during disconnect
        }
    }
    $window.Close()
})

# Exclude Countries button handler
$ClearStatusBtn.Add_Click({
    $StatusTextBox.Clear()
    Add-StatusMessage "Status cleared."
})

$GitHubBtn.Add_Click({
    Start-Process "https://github.com/Cavanite"
})

$EmailBtn.Add_Click({
    Start-Process "mailto:b.dezeeuw@bizway.nl"
})

$FixGraphModulesBtn.Add_Click({
    try {
        Add-StatusMessage "Starting Graph Modules cleanup..."
        
        # Disconnect from Graph if connected
        if ($script:GraphConnected) {
            Add-StatusMessage "Disconnecting from Microsoft Graph..."
            try {
                Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            } catch { }
            $script:GraphConnected = $false
            $SignInBtn.Visibility = "Visible"
            $DisconnectBtn.Visibility = "Collapsed"
            $ConnectionStatusText.Text = "Not Connected"
            $ConnectionStatusText.Foreground = "Gray"
            $ConnectionStatusBorder.Background = "#F0F0F0"
            $ConnectionStatusBorder.BorderBrush = "Gray"
            Add-StatusMessage "Disconnected."
        }
        
        # Unload all Graph modules from current session
        Add-StatusMessage "Unloading Graph modules from current session..."
        $loadedGraphModules = Get-Module -Name Microsoft.Graph.*
        foreach ($mod in $loadedGraphModules) {
            try {
                Remove-Module -Name $mod.Name -Force -ErrorAction SilentlyContinue
                Add-StatusMessage "  Unloaded: $($mod.Name)"
            } catch { }
        }
        
        Add-StatusMessage "Restarting script with Administrator permissions..."
        Add-StatusMessage "Please wait while modules are reinstalled..."
        
        Start-Sleep -Milliseconds 500
        
        # Get current script path
        $scriptPath = $PSCommandPath
        if (-not $scriptPath) {
            $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "Run-VacationMode.ps1"
        }
        
        # Build PowerShell command to run as admin
        $psCommand = @"
# Clear screen and set up progress display
Clear-Host
`$host.UI.RawUI.WindowTitle = 'Fixing Microsoft Graph Modules'

# Save current process ID to avoid killing ourselves
`$currentPID = `$PID

Write-Host '========================================' -ForegroundColor Cyan
Write-Host ' Microsoft Graph Module Repair Tool' -ForegroundColor Cyan
Write-Host '========================================' -ForegroundColor Cyan
Write-Host ''

# Wait for calling process to close completely
Write-Host '[1/5] Waiting for application to close' -ForegroundColor Yellow
for (`$i = 2; `$i -gt 0; `$i--) {
    Write-Host "      `$i seconds..." -NoNewline
    Start-Sleep -Seconds 1
    Write-Host "`r      `$i seconds... Done" -ForegroundColor Green
}
Write-Host ''

# Close all other PowerShell processes to release module locks
Write-Host '[2/5] Closing other PowerShell processes' -ForegroundColor Yellow
`$psProcesses = Get-Process -Name powershell, pwsh -ErrorAction SilentlyContinue | Where-Object { `$_.Id -ne `$currentPID }
if (`$psProcesses) {
    foreach (`$proc in `$psProcesses) {
        try {
            Write-Host "      Closing process PID `$(`$proc.Id)..." -NoNewline
            `$proc.CloseMainWindow() | Out-Null
            Start-Sleep -Milliseconds 500
            if (-not `$proc.HasExited) {
                `$proc.Kill()
            }
            Write-Host " Done" -ForegroundColor Green
        } catch { }
    }
} else {
    Write-Host "      No other PowerShell processes found" -ForegroundColor Green
}

# Give Windows time to release file locks
Write-Host "      Waiting for file locks to release..." -NoNewline
Start-Sleep -Seconds 2
Write-Host " Done" -ForegroundColor Green
Write-Host ''

# Uninstall only required Graph modules for this application
Write-Host '[3/5] Uninstalling Graph modules' -ForegroundColor Yellow
`$requiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Users', 'Microsoft.Graph.Identity.SignIns')
`$moduleIndex = 1
foreach (`$mod in `$requiredModules) {
    Write-Host "      [`$moduleIndex/`$(`$requiredModules.Count)] Uninstalling `$mod..." -NoNewline
    try {
        Uninstall-Module -Name `$mod -AllVersions -Force -ErrorAction Stop -WarningAction SilentlyContinue
        Write-Host " Done" -ForegroundColor Green
    } catch {
        Write-Host " Warning (will overwrite)" -ForegroundColor DarkYellow
    }
    `$moduleIndex++
}
Write-Host ''

# Install fresh versions
Write-Host '[4/5] Installing fresh Graph modules' -ForegroundColor Yellow
`$installModules = @(
    @{Name='Microsoft.Graph.Authentication'; Display='Authentication'},
    @{Name='Microsoft.Graph.Users'; Display='Users'},
    @{Name='Microsoft.Graph.Identity.SignIns'; Display='Identity.SignIns'}
)
`$installIndex = 1
foreach (`$mod in `$installModules) {
    Write-Host "      [`$installIndex/`$(`$installModules.Count)] Installing `$(`$mod.Display)..." -NoNewline
    Install-Module -Name `$mod.Name -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck -Repository PSGallery -WarningAction SilentlyContinue
    Write-Host " Done" -ForegroundColor Green
    `$installIndex++
}
Write-Host ''

# Restart application
Write-Host '[5/5] Restarting application' -ForegroundColor Yellow
Write-Host ''
Write-Host 'Graph modules reinstalled successfully!' -ForegroundColor Green
Write-Host ''
for (`$i = 3; `$i -gt 0; `$i--) {
    Write-Host "Launching application in `$i..." -NoNewline
    Start-Sleep -Seconds 1
    Write-Host "`r                                  `r" -NoNewline
}

# Restart the application
Start-Process -FilePath 'powershell.exe' -ArgumentList '-ExecutionPolicy Bypass -File "$scriptPath"'

# Close this admin window
Write-Host 'Closing this window...' -ForegroundColor Cyan
Start-Sleep -Seconds 1
exit
"@
        
        # Save command to temp file
        $tempScript = Join-Path -Path $env:TEMP -ChildPath "FixGraphModules_$(Get-Date -Format 'yyyyMMddHHmmss').ps1"
        $psCommand | Out-File -FilePath $tempScript -Encoding UTF8 -Force
        
        # Start elevated process (changed -NoExit to just run and close)
        Start-Process -FilePath "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass -File `"$tempScript`"" -Verb RunAs
        
        # Close current window
        $window.Close()
        
    } catch {
        Add-StatusMessage "ERROR: Failed to fix Graph modules - $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Failed to fix Graph modules: $($_.Exception.Message)", "Error", "OK", "Error")
    }
})

$ExcludeCountriesBtn.Add_Click({
    # Create the country creation window
    [xml]$createCountriesXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Create Countries" Height="600" Width="700"
    WindowStartupLocation="CenterScreen" Topmost="True">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <TextBlock Grid.Row="0" Text="Create New Countries" 
                   FontSize="18" FontWeight="Bold" 
                   HorizontalAlignment="Left" Margin="10"/>
        
        <!-- Input Section -->
        <GroupBox Grid.Row="1" Header="Select Country to Add" 
                  FontSize="12" FontWeight="Bold" Margin="10,5,10,10">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <TextBlock Grid.Row="0" Grid.Column="0" Grid.ColumnSpan="2"
                           Text="Select from the list of available countries:" 
                           Margin="5,5,5,5" FontSize="11"/>
                
                <ComboBox Grid.Row="1" Grid.Column="0" Name="CountrySelectionComboBox" 
                          Height="30" Margin="5,0,5,5"
                          VerticalContentAlignment="Center"
                          IsEditable="True"
                          IsTextSearchEnabled="True"
                          Padding="5"/>
                
                <Button Grid.Row="1" Grid.Column="1" Name="AddCountryBtn" 
                        Content="Add Country" Width="100" Height="30" Margin="0,0,5,5"
                        FontWeight="Bold" Background="#0078D4" Foreground="White"/>
            </Grid>
        </GroupBox>
        
        <!-- List Section -->
        <Grid Grid.Row="2" Margin="10">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            
            <TextBlock Grid.Row="0" Text="Countries to Create:" 
                       FontSize="12" FontWeight="Bold" Margin="0,0,0,5"/>
            
            <ListBox Grid.Row="1" Name="CountriesToCreateListBox" 
                     SelectionMode="Single"
                     Margin="0"
                     VerticalAlignment="Stretch"/>
        </Grid>
        
        <!-- Buttons -->
        <StackPanel Grid.Row="3" Orientation="Horizontal" 
                    HorizontalAlignment="Right" Margin="10">
            <Button Name="RemoveCountryBtn" Content="Remove Selected" 
                    Width="120" Height="35" Margin="5" 
                    Background="#D13438" Foreground="White"/>
            <Button Name="CreateCountriesBtn" Content="Create Countries" 
                    Width="120" Height="35" Margin="5" 
                    FontWeight="Bold" Background="#107C10" Foreground="White"/>
            <Button Name="CancelCreateCountriesBtn" Content="Cancel" 
                    Width="100" Height="35" Margin="5"/>
        </StackPanel>
    </Grid>
</Window>
"@
    
    try {
        $createCountriesReader = New-Object System.Xml.XmlNodeReader $createCountriesXaml
        $createCountriesWindow = [Windows.Markup.XamlReader]::Load($createCountriesReader)
        
        # Get UI elements
        $CountrySelectionComboBox = $createCountriesWindow.FindName("CountrySelectionComboBox")
        $AddCountryBtn = $createCountriesWindow.FindName("AddCountryBtn")
        $CountriesToCreateListBox = $createCountriesWindow.FindName("CountriesToCreateListBox")
        $RemoveCountryBtn = $createCountriesWindow.FindName("RemoveCountryBtn")
        $CreateCountriesBtn = $createCountriesWindow.FindName("CreateCountriesBtn")
        $CancelCreateCountriesBtn = $createCountriesWindow.FindName("CancelCreateCountriesBtn")
        
        # Check if connected to Graph
        if (-not $script:GraphConnected) {
            [System.Windows.MessageBox]::Show(
                "Please sign in to Microsoft Graph first.",
                "Not Connected",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning
            )
            return
        }
        
        # Load countries from Initialize-Named-Location module
        try {
            $allCountries = @()
            $script:CountryCodeLookup = @{}
            
            # Get existing Named Locations from Graph
            $existingNamedLocations = @()
            try {
                $existingNamedLocations = Get-MgIdentityConditionalAccessNamedLocation -All -ErrorAction Stop | Select-Object -ExpandProperty DisplayName
            } catch {
                Write-Verbose "Could not fetch existing named locations: $_"
            }
            
            # Load from the module content
            $modulePath = Join-Path -Path $ModulesPath -ChildPath "Initialize-Named-Location.psm1"
            if (Test-Path $modulePath) {
                # Read the module and extract country list
                $moduleContent = Get-Content -Path $modulePath -Raw
                
                # Extract the country definitions (simplified parsing)
                # The countries are defined as @() array with PSCustomObject items
                $countryMatches = [regex]::Matches($moduleContent, '\[PSCustomObject\]@\{Name = "([^"]+)"; Code = "([^"]+)"\}')
                
                foreach ($match in $countryMatches) {
                    $countryName = $match.Groups[1].Value
                    $countryCode = $match.Groups[2].Value
                    $allCountries += @{
                        Name = $countryName
                        Code = $countryCode
                        Exists = $existingNamedLocations -contains $countryName
                    }
                    # Create lookup table for country name -> code
                    $script:CountryCodeLookup[$countryName] = $countryCode
                }
            }
            
            if ($allCountries.Count -eq 0) {
                Add-StatusMessage "WARNING: Could not load countries from module, using basic country list"
                $allCountries = @(
                    @{Name="United States"; Code="US"; Exists = $existingNamedLocations -contains "United States"},
                    @{Name="United Kingdom"; Code="GB"; Exists = $existingNamedLocations -contains "United Kingdom"},
                    @{Name="Canada"; Code="CA"; Exists = $existingNamedLocations -contains "Canada"},
                    @{Name="Australia"; Code="AU"; Exists = $existingNamedLocations -contains "Australia"},
                    @{Name="Germany"; Code="DE"; Exists = $existingNamedLocations -contains "Germany"},
                    @{Name="France"; Code="FR"; Exists = $existingNamedLocations -contains "France"},
                    @{Name="Spain"; Code="ES"; Exists = $existingNamedLocations -contains "Spain"},
                    @{Name="Netherlands"; Code="NL"; Exists = $existingNamedLocations -contains "Netherlands"}
                )
                foreach ($country in $allCountries) {
                    $script:CountryCodeLookup[$country.Name] = $country.Code
                }
            }
            
            # Populate the combo box with indicator for existing countries
            foreach ($country in ($allCountries | Sort-Object { $_.Name })) {
                if ($country.Exists) {
                    $displayText = "$($country.Name) (Already Exists)"
                } else {
                    $displayText = $country.Name
                }
                $CountrySelectionComboBox.Items.Add($displayText) | Out-Null
            }
        } catch {
            Add-StatusMessage "ERROR: Failed to load countries - $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show(
                "Failed to load countries from module: $($_.Exception.Message)",
                "Error Loading Countries",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
            return
        }
        
        # Button event handlers
        $AddCountryBtn.Add_Click({
            $selectedCountryDisplay = $CountrySelectionComboBox.SelectedItem
            
            if ([string]::IsNullOrWhiteSpace($selectedCountryDisplay)) {
                [System.Windows.MessageBox]::Show(
                    "Please select a country from the dropdown list.",
                    "No Country Selected",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Warning
                )
                return
            }
            
            # Strip the "(Already Exists)" label to get the actual country name
            $selectedCountry = $selectedCountryDisplay -replace '\s*\(Already Exists\)$', ''
            
            # Check if country already exists in list
            if ($CountriesToCreateListBox.Items -contains $selectedCountry) {
                [System.Windows.MessageBox]::Show(
                    "Country '$selectedCountry' is already in the list.",
                    "Duplicate Country",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Warning
                )
                return
            }
            
            $CountriesToCreateListBox.Items.Add($selectedCountry) | Out-Null
            $CountrySelectionComboBox.SelectedIndex = -1
            $CountrySelectionComboBox.Focus()
        })
        
        # Allow Enter key to add country
        $CountrySelectionComboBox.Add_KeyDown({
            if ($_.Key -eq 'Return') {
                $AddCountryBtn.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
            }
        })
        
        $RemoveCountryBtn.Add_Click({
            if ($CountriesToCreateListBox.SelectedItem -ne $null) {
                $CountriesToCreateListBox.Items.Remove($CountriesToCreateListBox.SelectedItem)
            } else {
                [System.Windows.MessageBox]::Show(
                    "Please select a country to remove.",
                    "No Selection",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Warning
                )
            }
        })
        
        $CreateCountriesBtn.Add_Click({
            $countriesToCreate = @($CountriesToCreateListBox.Items)
            
            if ($countriesToCreate.Count -eq 0) {
                [System.Windows.MessageBox]::Show(
                    "Please add at least one country to create.",
                    "No Countries",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Warning
                )
                return
            }
            
            # Check for duplicate countries before creating
            $existingCountries = @()
            try {
                $namedLocations = Get-MgIdentityConditionalAccessNamedLocation -All -ErrorAction Stop
                $existingCountries = $namedLocations.DisplayName
            } catch {
                Add-StatusMessage "WARNING: Could not verify existing named locations - $($_.Exception.Message)"
            }
            
            # Check if any countries to create already exist
            $duplicateCountries = $countriesToCreate | Where-Object { $_ -in $existingCountries }
            
            if ($duplicateCountries.Count -gt 0) {
                $duplicateList = $duplicateCountries -join ", "
                $result = [System.Windows.MessageBox]::Show(
                    "Duplicate Country name found:`n`n$duplicateList`n`nDo you want to continue?",
                    "Duplicate Country",
                    [System.Windows.MessageBoxButton]::YesNo,
                    [System.Windows.MessageBoxImage]::Warning
                )
                
                if ($result -ne "Yes") {
                    return
                }
            }
            
            # Create named locations for each country
            $successCount = 0
            $failureCount = 0
            
            foreach ($country in $countriesToCreate) {
                try {
                    Add-StatusMessage "Creating named location: $country..."
                    
                    # Get the ISO 3166-1 alpha-2 country code from lookup table
                    $countryCode = $script:CountryCodeLookup[$country]
                    
                    if ([string]::IsNullOrWhiteSpace($countryCode)) {
                        Add-StatusMessage "ERROR: Could not find country code for '$country'"
                        $failureCount++
                        continue
                    }
                    
                    # Create a country-based named location using the correct odata type
                    $params = @{
                        DisplayName = $country
                        "@odata.type" = "#microsoft.graph.countryNamedLocation"
                        countriesAndRegions = @($countryCode)
                        includeUnknownCountriesAndRegions = $false
                    }
                    
                    $result = New-MgIdentityConditionalAccessNamedLocation -BodyParameter $params -ErrorAction Stop
                    
                    if ($result) {
                        Add-StatusMessage "SUCCESS: Created named location '$country' (Code: $countryCode, ID: $($result.Id))"
                        $successCount++
                    }
                } catch {
                    Add-StatusMessage "ERROR: Failed to create '$country' - $($_.Exception.Message)"
                    $failureCount++
                }
            }
            
            # Refresh the main country combo box
            try {
                $namedLocations = Get-MgIdentityConditionalAccessNamedLocation -All -ErrorAction Stop
                $script:NamedLocationsCache = @{}
                $CountryComboBox.Items.Clear()
                foreach ($location in $namedLocations) {
                    $script:NamedLocationsCache[$location.DisplayName] = $location.Id
                    $CountryComboBox.Items.Add($location.DisplayName) | Out-Null
                }
            } catch {
                Add-StatusMessage "WARNING: Could not refresh country list - $($_.Exception.Message)"
            }
            
            [System.Windows.MessageBox]::Show(
                "Created $successCount countries successfully." + $(if ($failureCount -gt 0) { "`nFailed to create $failureCount countries." } else { "" }),
                "Countries Created",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )
            
            $createCountriesWindow.Close()
        })
        
        $CancelCreateCountriesBtn.Add_Click({
            $createCountriesWindow.Close()
        })
        
        $createCountriesWindow.ShowDialog() | Out-Null
        
    } catch {
        [System.Windows.MessageBox]::Show(
            "Failed to open country creation window: $($_.Exception.Message)",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
})

# Show the window
Add-StatusMessage "Ready."
$window.ShowDialog() | Out-Null
