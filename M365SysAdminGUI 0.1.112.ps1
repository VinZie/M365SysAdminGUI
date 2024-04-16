Add-Type -AssemblyName PresentationFramework

# Function to check if a module is installed
function CheckModuleInstalled {
    param (
        [string]$moduleName
    )
    $module = Get-Module -Name $moduleName -ListAvailable
    return [bool]($module -ne $null)
}

# Function to handle Teams tasks
function ManageTeams {
    if (CheckModuleInstalled "MicrosoftTeams") {
        # Attempt to connect to Microsoft Teams
        try {
            Connect-MicrosoftTeams -ErrorAction Stop
            [System.Windows.MessageBox]::Show("Successfully connected to Microsoft Teams. You can proceed with Teams management tasks.", "Connection Successful", "OK", "Information")
        } catch {
            [System.Windows.MessageBox]::Show("Failed to connect to Microsoft Teams. Error: $($_.Exception.Message)", "Connection Error", "OK", "Error")
        }
    } else {
        # Show a message indicating the Teams module is not installed
        [System.Windows.MessageBox]::Show("Microsoft Teams PowerShell module is not installed. Please install the module to proceed with Teams management tasks.", "Module Not Installed", "OK", "Error")
    }
}

# Function to handle Azure tasks
function ManageAzure {
    if (CheckModuleInstalled "AzureAD") {
        # Attempt to connect to Azure
        try {
            Connect-AzAccount -ErrorAction Stop
            [System.Windows.MessageBox]::Show("Successfully connected to Azure. You can proceed with Azure management tasks.", "Connection Successful", "OK", "Information")

            # Create a new window for displaying additional options after authentication
            $AzureAuthWindow = New-Object -TypeName System.Windows.Window
            $AzureAuthWindow.Title = "Azure Management"
            $AzureAuthWindow.Width = 300
            $AzureAuthWindow.Height = 180
            $AzureAuthWindow.WindowStartupLocation = "CenterOwner"
            $AzureAuthWindow.Background = "#2b2b2b"

            # Add a text block to inform the user about the authentication success
            $TextBlock = New-Object -TypeName System.Windows.Controls.TextBlock
            $TextBlock.Text = "Azure authentication successful."
            $TextBlock.Margin = "10"
            $TextBlock.Foreground = "White"

            # Add a button to execute Get-AzADUser
            $ButtonGetUsers = New-Object -TypeName System.Windows.Controls.Button
            $ButtonGetUsers.Content = "All Users in AAD"
            $ButtonGetUsers.Margin = "10"
            $ButtonGetUsers.Background = "#444"
            $ButtonGetUsers.Foreground = "White"
            $ButtonGetUsers.Add_Click({
                Get-AzADUser | Out-GridView -Title "All Users in Azure Active Directory"
            })

            # Add a button to execute New-AzADUser
            $ButtonNewUser = New-Object -TypeName System.Windows.Controls.Button
            $ButtonNewUser.Content = "New User in AAD"
            $ButtonNewUser.Margin = "10"
            $ButtonNewUser.Background = "#444"
            $ButtonNewUser.Foreground = "White"
            $ButtonNewUser.Add_Click({
                New-AzureADUser
            })

            # Add controls to the window
            $AzureAuthWindow.Content = $TextBlock
            $AzureAuthWindow.Content.AddChild($ButtonGetUsers)
            $AzureAuthWindow.Content.AddChild($ButtonNewUser)

            # Show the window
            $AzureAuthWindow.ShowDialog() | Out-Null

        } catch {
            [System.Windows.MessageBox]::Show("Failed to connect to Azure. Error: $($_.Exception.Message)", "Connection Error", "OK", "Error")
        }
    } else {
        # Show a message indicating the Az module is not installed
        [System.Windows.MessageBox]::Show("Azure PowerShell module is not installed. Please install the module to proceed with Azure management tasks.", "Module Not Installed", "OK", "Error")
    }
}

# Function to handle SharePoint tasks
function ManageSharePoint {
    if (CheckModuleInstalled "Microsoft.Online.SharePoint.PowerShell") {
        # Show a message indicating the SharePoint module is installed
        [System.Windows.MessageBox]::Show("Microsoft Online SharePoint PowerShell module is installed. You can proceed with SharePoint management tasks.", "Module Installed", "OK", "Information")
    } else {
        # Show a message indicating the SharePoint module is not installed
        [System.Windows.MessageBox]::Show("Microsoft Online SharePoint PowerShell module is not installed. Please install the module to proceed with SharePoint management tasks.", "Module Not Installed", "OK", "Error")
    }
}

# Function to handle Graph tasks
function ManageGraph {
    if (CheckModuleInstalled "Microsoft.Graph") {
        # Show a message indicating the Graph module is installed
        [System.Windows.MessageBox]::Show("Microsoft Graph PowerShell module is installed. You can proceed with Graph management tasks.", "Module Installed", "OK", "Information")
    } else {
        # Show a message indicating the Graph module is not installed
        [System.Windows.MessageBox]::Show("Microsoft Graph PowerShell module is not installed. Please install the module to proceed with Graph management tasks.", "Module Not Installed", "OK", "Error")
    }
}

# Define XAML for the email input window
$EmailXAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Enter Email Address" Height="200" Width="400">
    <StackPanel Background="#2b2b2b" Margin="20">
        <TextBlock Foreground="White" Text="Please enter your email address:" Margin="0 0 0 10"/>
        <TextBox x:Name="EmailTextBox" Width="250" Background="#444" Foreground="White" Margin="0 0 0 10"/>
        <Button Content="Submit" x:Name="EmailSubmitButton" Background="#444" Foreground="White" HorizontalAlignment="Right" Width="80"/>
    </StackPanel>
</Window>
"@

# Load the XAML content for the email input window
$EmailWindow = [Windows.Markup.XamlReader]::Parse($EmailXAML)

# Define the email address variable
$global:EmailAddress = $null

# Define XAML for the Exchange management window
$ExchangeXAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Exchange Management" Height="300" Width="400" Background="#2b2b2b" ResizeMode="NoResize">
    <Grid>
        <Label Content="Select an Exchange task to proceed:" HorizontalAlignment="Center" VerticalAlignment="Top" Margin="0,20,0,0" Foreground="White"/>
        <StackPanel Margin="20,60,20,0">
            <Button Name="btnCreateMailbox" Content="Create Mailbox" HorizontalAlignment="Stretch" VerticalAlignment="Center" Height="30" Background="#444" Foreground="White" Margin="0,5,0,0"/>
            <Button Name="btnConvertToSharedMailbox" Content="Convert to Shared Mailbox" HorizontalAlignment="Stretch" VerticalAlignment="Center" Height="30" Background="#444" Foreground="White" Margin="0,5,0,0"/>
            <Button Name="btnRemoveMailbox" Content="Remove Mailbox" HorizontalAlignment="Stretch" VerticalAlignment="Center" Height="30" Background="#444" Foreground="White" Margin="0,5,0,0"/>
            <!-- Add more buttons for additional tasks as needed -->
        </StackPanel>
    </Grid>
</Window>
"@

# Load the XAML content for the Exchange management window
$ExchangeWindow = [Windows.Markup.XamlReader]::Parse($ExchangeXAML)

# Load XAML content
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="M365 PowerShell Modules GUI" Height="300" Width="400" Background="#2b2b2b" ResizeMode="NoResize">
    <Grid>
        <Label Content="Select a task to proceed:" HorizontalAlignment="Center" VerticalAlignment="Top" Margin="0,20,0,0" Foreground="White"/>
        <StackPanel Margin="20,60,20,0">
            <Button Name="btnExchange" Content="Manage Exchange" HorizontalAlignment="Stretch" VerticalAlignment="Center" Height="30" Background="#444" Foreground="White" Margin="0,5,0,0"/>
            <Button Name="btnAzure" Content="Manage Azure" HorizontalAlignment="Stretch" VerticalAlignment="Center" Height="30" Background="#444" Foreground="White" Margin="0,5,0,0"/>
            <Button Name="btnSharePoint" Content="Manage SharePoint" HorizontalAlignment="Stretch" VerticalAlignment="Center" Height="30" Background="#444" Foreground="White" Margin="0,5,0,0"/>
            <Button Name="btnTeams" Content="Manage Teams" HorizontalAlignment="Stretch" VerticalAlignment="Center" Height="30" Background="#444" Foreground="White" Margin="0,5,0,0"/>
            <Button Name="btnGraph" Content="Manage Graph" HorizontalAlignment="Stretch" VerticalAlignment="Center" Height="30" Background="#444" Foreground="White" Margin="0,5,0,0"/>
        </StackPanel>
    </Grid>
</Window>
"@

# Load XAML into a XAML reader
$reader=(New-Object System.Xml.XmlNodeReader $xaml)

# Create XAML object
$Window=[Windows.Markup.XamlReader]::Load($reader)

# Find controls by name
$btnExchange = $Window.FindName("btnExchange")
$btnAzure = $Window.FindName("btnAzure")
$btnSharePoint = $Window.FindName("btnSharePoint")
$btnTeams = $Window.FindName("btnTeams")
$btnGraph = $Window.FindName("btnGraph")

# Add event handlers for the task buttons
$btnExchange.Add_Click({
    ManageAzure
})

$btnAzure.Add_Click({
    ManageAzure
})

$btnSharePoint.Add_Click({
    ManageSharePoint
})

$btnTeams.Add_Click({
    ManageTeams
})

$btnGraph.Add_Click({
    ManageGraph
})

# Define the button click event handler for the email input window
$EmailWindow.FindName("EmailSubmitButton").Add_Click({
    # Retrieve the email address entered by the user
    $global:EmailAddress = $EmailWindow.FindName("EmailTextBox").Text
    $EmailWindow.Close()

    # If email is provided, connect to Exchange Online
    if ($global:EmailAddress) {
        try {
            Connect-IPPSSession -UserPrincipalName $global:EmailAddress -ErrorAction Stop
            [System.Windows.MessageBox]::Show("Successfully connected to Exchange Online with the provided email address: $($global:EmailAddress)", "Task Executed", "OK", "Information")
            $ExchangeWindow.ShowDialog() | Out-Null
        } catch {
            [System.Windows.MessageBox]::Show("Failed to connect to Exchange Online. Error: $($_.Exception.Message)", "Error", "OK", "Error")
        }
    } else {
        [System.Windows.MessageBox]::Show("No email address provided.", "Task Canceled", "OK", "Warning")
    }
})

# Display the window
$Window.ShowDialog() | Out-Null
