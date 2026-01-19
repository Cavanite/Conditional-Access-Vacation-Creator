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
    Title="Conditional Access Vacation Creator" Height="900" Width="1200"
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
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="200"/>
                    </Grid.RowDefinitions>
                    
                    <TextBlock Grid.Row="0" Text="Select Users:" 
                               Margin="5,5,5,2" FontSize="11" FontWeight="Bold"/>
                    
                    <ListBox Grid.Row="1" Name="UsersListBox" 
                             SelectionMode="Multiple"
                             Margin="5"
                             VerticalAlignment="Stretch"/>
                    
                    <Grid Grid.Row="2" Margin="5">
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
                    <GroupBox Grid.Row="4" Header="Status" 
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
                                     Height="180"/>
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
                        
                        <TextBlock Grid.Row="0" Text="Ticket Number:" Margin="5,5,5,2"/>
                        <TextBox Grid.Row="1" Name="TicketNumberTextBox" 
                                 Margin="5,0,5,5" Height="25"/>
                        
                        <TextBlock Grid.Row="2" Text="End Date (dd-mm-yyyy):" Margin="5,5,5,2"/>
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
    $requiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Users', 'Microsoft.Graph.Identity.SignIns')
    
    foreach ($moduleName in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $moduleName)) {
            Add-StatusMessage "Installing $moduleName..."
            try {
                Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber
                Add-StatusMessage "Successfully installed $moduleName"
            } catch {
                Add-StatusMessage "ERROR: Failed to install $moduleName - $($_.Exception.Message)"
                throw
            }
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
        Connect-MgGraph -Scopes "User.Read.All", "Policy.Read.All", "Policy.ReadWrite.ConditionalAccess" -NoWelcome
        
        $context = Get-MgContext
        if ($context) {
            $script:GraphConnected = $true
            Add-StatusMessage "SUCCESS: Connected as $($context.Account)"
            
            # Check for Conditional Access license (Azure AD Premium P1 or higher)
            Add-StatusMessage "Checking for Conditional Access license..."
            $licenseCheckFailed = $false
            try {
                # Try to read CA policies - if tenant doesn't have P1, this will fail
                $testPolicy = Get-MgIdentityConditionalAccessPolicy -Top 1 -ErrorAction Stop
                Add-StatusMessage "SUCCESS: Conditional Access is available in this tenant"
            } catch {
                $errorMessage = $_.Exception.Message
                $innerError = $_.Exception.InnerException.Message
                
                Add-StatusMessage "ERROR: Conditional Access not available - $errorMessage"
                
                # Check if it's a license/premium issue
                if ($errorMessage -like "*does not have a premium license*" -or 
                    $errorMessage -like "*Premium*" -or 
                    $errorMessage -like "*license*" -or 
                    $errorMessage -like "*subscription*" -or
                    $innerError -like "*Premium*" -or
                    $innerError -like "*license*") {
                    
                    $licenseCheckFailed = $true
                } else {
                    # Generic error - might still be a license issue, so warn
                    Add-StatusMessage "WARNING: Could not verify Conditional Access license"
                    $licenseCheckFailed = $true
                }
            }
            
            if ($licenseCheckFailed) {
                $window.Topmost = $true
                [System.Windows.MessageBox]::Show(
                    $window,
                    "Conditional Access License Required`n`n" +
                    "This tenant appears to be using Entra ID Free and does not have the required licenses to use Conditional Access policies.`n`n" +
                    "To use Conditional Access, you need one of:`n" +
                    "  - Azure AD Premium P1`n" +
                    "  - Azure AD Premium P2`n" +
                    "  - Microsoft 365 Business Premium`n`n" +
                    "Please upgrade your tenant's subscription to continue.`n`n" +
                    "For assistance, contact: b.dezeeuw@bizway.nl",
                    "License Required",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Warning
                )
                $window.Topmost = $true
                
                # Disconnect and reset
                try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
                $script:GraphConnected = $false
                $SignInBtn.Visibility = "Visible"
                $DisconnectBtn.Visibility = "Collapsed"
                $RefreshUsersBtn.IsEnabled = $false
                $ConnectionStatusText.Text = "Not Connected"
                $ConnectionStatusText.Foreground = "Gray"
                $ConnectionStatusBorder.Background = "#F0F0F0"
                $ConnectionStatusBorder.BorderBrush = "Gray"
                return
            }
            
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
                Add-StatusMessage "ERROR: Failed to load named locations - Named Locations require Azure AD Premium P1 or higher"
                
                # This is likely a license issue too
                $window.Topmost = $true
                [System.Windows.MessageBox]::Show(
                    $window,
                    "Named Locations require Azure AD Premium`n`n" +
                    "Named Locations are a premium feature that requires Azure AD Premium P1 or higher.`n`n" +
                    "This tenant appears to be using Entra ID Free.`n`n" +
                    "For assistance, contact: b.dezeeuw@bizway.nl",
                    "Premium Feature Required",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Warning
                )
                $window.Topmost = $true
                
                # Disconnect and reset
                try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
                $script:GraphConnected = $false
                $SignInBtn.Visibility = "Visible"
                $DisconnectBtn.Visibility = "Collapsed"
                $RefreshUsersBtn.IsEnabled = $false
                $ConnectionStatusText.Text = "Not Connected"
                $ConnectionStatusText.Foreground = "Gray"
                $ConnectionStatusBorder.Background = "#F0F0F0"
                $ConnectionStatusBorder.BorderBrush = "Gray"
                return
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
                    Add-StatusMessage "ERROR: No geofencing policy found!"
                    [System.Windows.MessageBox]::Show("No Conditional Access policy with 'Geofencing' or 'Countrywhitelist' in the name was found in your tenant.`n`nThis tool requires an existing geofencing policy to function properly.`n`nPlease create a policy with 'Geofencing' or 'Countrywhitelist' in its display name.`n`nFor more information, contact Bert de Zeeuw at b.dezeeuw@bizway.nl", "Configuration Required", "OK", "Error")
                    
                    # Disable policy creation
                    $ExistingPolicyComboBox.IsEnabled = $false
                    $CreatePolicyBtn.IsEnabled = $false
                    return
                }
                
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
                
                # Automatically load users after successful connection
                Add-StatusMessage "Automatically loading users..."
                $RefreshUsersBtn.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
            } catch {
                $errorMsg = $_.Exception.Message
                Add-StatusMessage "ERROR: Failed to load CA policies - Conditional Access requires Azure AD Premium P1 or higher"
                
                # This confirms it's a license issue
                $window.Topmost = $true
                [System.Windows.MessageBox]::Show(
                    $window,
                    "Conditional Access requires Azure AD Premium`n`n" +
                    "Conditional Access Policies require Azure AD Premium P1 or higher.`n`n" +
                    "This tenant appears to be using Entra ID Free.`n`n" +
                    "For assistance, contact: b.dezeeuw@bizway.nl",
                    "Premium Feature Required",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Warning
                )
                $window.Topmost = $true
                
                # Disconnect and reset
                try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
                $script:GraphConnected = $false
                $SignInBtn.Visibility = "Visible"
                $DisconnectBtn.Visibility = "Collapsed"
                $RefreshUsersBtn.IsEnabled = $false
                $ConnectionStatusText.Text = "Not Connected"
                $ConnectionStatusText.Foreground = "Gray"
                $ConnectionStatusBorder.Background = "#F0F0F0"
                $ConnectionStatusBorder.BorderBrush = "Gray"
                return
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
        $users = Get-MgUser -All -Property Id,DisplayName,UserPrincipalName | Select-Object Id,DisplayName,UserPrincipalName
        
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
        
        # Clear and populate the cache as a hashtable
        $script:UserCache = @{}
        $UsersListBox.Items.Clear()
        
        $filteredCount = 0
        foreach ($user in $users) {
            # Check if user is external (contains #EXT# in UPN)
            $isExternal = $user.UserPrincipalName -like '*#EXT#*'
            
            # Check if user matches any exclusion pattern
            $shouldExclude = $false
            foreach ($pattern in $excludePatterns) {
                if (($user.DisplayName -like $pattern) -or ($user.UserPrincipalName -like $pattern)) {
                    $shouldExclude = $true
                    break
                }
            }
            
            # Exclude if external or matches exclusion pattern
            if ($isExternal -or $shouldExclude) {
                $filteredCount++
            }
            # Exclude if external or matches exclusion pattern
            if ($isExternal -or $shouldExclude) {
                $filteredCount++
            } else {
                # Only add internal, non-admin/non-breakglass users
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
