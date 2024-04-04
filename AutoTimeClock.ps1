function Get-AccessToken {
    param (
        [string]$clientId,
        [string]$tenantId,
        [string]$clientSecret
    )

    $tokenEndpoint = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
    $body = @{
        grant_type    = "client_credentials"
        client_id     = $clientId
        client_secret = $clientSecret
        scope         = "https://graph.microsoft.com/.default"
    }

    $response = Invoke-RestMethod -Uri $tokenEndpoint -Method Post -Body $body
    return $response.access_token
}

function Get-UserIdByEmail {
    param (
        [string]$accessToken,
        [string]$email
    )

    $apiUrl = "https://graph.microsoft.com/v1.0/users?$filter=mail eq '$email'"
    $headers = @{
        Authorization = "Bearer $accessToken"
    }

    $response = Invoke-RestMethod -Uri $apiUrl -Method Get -Headers $headers
    if ($response.value.Count -gt 0) {
        return $response.value[0].id
    }
    else {
        Write-Host "User not found."
        return $null
    }
}

function Get-TeamId {
    param (
        [string]$teamName,
        [string]$accessToken,
        [string]$userId
    )

    $apiUrl = "https://graph.microsoft.com/v1.0/me/joinedTeams"
    $headers = @{
        Authorization    = "Bearer $accessToken"
        "MS-APP-ACTS-AS" = $userId
    }

    $response = Invoke-RestMethod -Uri $apiUrl -Method Get -Headers $headers
    $team = $response.value | Where-Object { $_.displayName -eq $teamName }
    if ($team) {
        return $team.id
    }
    else {
        Write-Host "Team '$teamName' not found."
        return $null
    }
}

function IsTeamsRunning {
    return $null -ne (Get-Process | Where-Object { $_.Name -match "Teams" })
}

function Get-TeamsStatus{
    param(
        [string]$accessToken,
        [string]$userId
    )

    $apiUrl = "https://graph.microsoft.com/v1.0/me/presence"
    $headers = @{
        Authorization   = "Bearer $accessToken"
        "MS-APP-ACTS-AS" = $userId
    }

    $response = Invoke-RestMethod -Uri $apiUrl -Method Get -Headers $headers
    return $response.availability
}
function ClockIn {
    param (
        [string]$teamId,
        [string]$accessToken,
        [string]$userId
    )

    $apiUrl = "https://graph.microsoft.com/beta/teams/$teamId/schedule/timeCards/clockIn"
    $headers = @{
        Authorization    = "Bearer $accessToken"
        "MS-APP-ACTS-AS" = $userId
    }

    if ($null -ne $teamId -and $null -ne $userId) {
        # Prompt the user with a MessageBox
        Add-Type -AssemblyName System.Windows.Forms
        $result = [System.Windows.Forms.MessageBox]::Show("It's " + (Get-Date -Format "HH:mm") + " right now. Do you want to Clock In?", "Clock In", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    }
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $response = Invoke-RestMethod -Uri $apiUrl -Method Post -Headers $headers -ContentType "application/json"
        return $response.id
    }
    else {
        Write-Host "Clock In cancelled."
        return $null
    }
}

function ClockOut {
    param (
        [string]$teamId,
        [string]$timeCardId,
        [string]$accessToken,
        [string]$userId
    )

    $result = $null
    $apiUrl = "https://graph.microsoft.com/beta/teams/$teamId/schedule/timeCards/$timeCardId/clockOut"
    $headers = @{
        Authorization    = "Bearer $accessToken"
        "MS-APP-ACTS-AS" = $userId
    }

    if (![string]::IsNullOrEmpty($timeCardId)) {        
        # Prompt the user with a MessageBox
        Add-Type -AssemblyName System.Windows.Forms
        $result = [System.Windows.Forms.MessageBox]::Show("It's " + (Get-Date -Format "HH:mm") + " right now. Do you want to Clock Out?", "Clock Out", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    }
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        Invoke-RestMethod -Uri $apiUrl -Method Post -Headers $headers
    }
    else {
        Write-Host "Clock Out cancelled."
    }
}

function StartBreak {
    param (
        [string]$teamId,
        [string]$timeCardId,
        [string]$accessToken,
        [string]$userId
    )

    $apiUrl = "https://graph.microsoft.com/beta/teams/$teamId/schedule/timeCards/$timeCardId/startBreak"
    $headers = @{
        Authorization    = "Bearer $accessToken"
        "MS-APP-ACTS-AS" = $userId
    }

    Invoke-RestMethod -Uri $apiUrl -Method Post -Headers $headers -ContentType "application/json"
}

function EndBreak{
    param (
        [string]$teamId,
        [string]$timeCardId,
        [string]$accessToken,
        [string]$userId
    )

    $apiUrl = "https://graph.microsoft.com/beta/teams/$teamId/schedule/timeCards/$timeCardId/endBreak"
    $headers = @{
        Authorization    = "Bearer $accessToken"
        "MS-APP-ACTS-AS" = $userId
    }

    Invoke-RestMethod -Uri $apiUrl -Method Post -Headers $headers -ContentType "application/json"
}

# Load necessary assemblies
Add-Type -AssemblyName PresentationFramework

# Create the form
$form = New-Object System.Windows.Window
$form.Title = "Enter Details"
$form.Width = 400
$form.Height = 250

# Create a stack panel to hold the controls
$stackPanel = New-Object System.Windows.Controls.StackPanel

# Create the email ID label and text box
$emailLabel = New-Object System.Windows.Controls.Label
$emailLabel.Content = "Email ID:"
$emailTextBox = New-Object System.Windows.Controls.TextBox
$emailTextBox.Width = 200

# Create the team name label and text box
$teamNameLabel = New-Object System.Windows.Controls.Label
$teamNameLabel.Content = "Team Name:"
$teamNameTextBox = New-Object System.Windows.Controls.TextBox
$teamNameTextBox.Width = 200

# Create the submit button
$submitButton = New-Object System.Windows.Controls.Button
$submitButton.Content = "Submit"
$submitButton.Width = 100
$submitButton.Height = 30
$submitButton.Margin = "0,10,0,0"
$submitButton.Add_Click({
        $form.DialogResult = $true
        $form.Close()
    })

# Add the controls to the stack panel
$stackPanel.Children.Add($emailLabel)
$stackPanel.Children.Add($emailTextBox)
$stackPanel.Children.Add($teamNameLabel)
$stackPanel.Children.Add($teamNameTextBox)
$stackPanel.Children.Add($submitButton)

# Add the stack panel to the form
$form.Content = $stackPanel

# Set the configuration file path
$ConfigPath = "$env:userprofile\userconfig.xml"

# Check if the configuration file exists
if (Test-Path $ConfigPath) {
    # Load the configuration
    $Config = Import-Clixml $ConfigPath
    $email = $Config.email
    $teamName = $Config.teamName
} else {
    # The file does not exist, create it
    New-Item -Path $ConfigPath -ItemType File

    # Show the form and wait for the user to submit
    $result = $form.ShowDialog()
    if ($result -eq $true) {
        $email = $emailTextBox.Text
        $teamName = $teamNameTextBox.Text
        Write-Host "Email: $email"
        Write-Host "Team Name: $teamName"

        # Save the configuration to the XML file
        $Config = @{
            email = $email
            teamName = $teamName
        }
        $Config | Export-Clixml -Path $ConfigPath
    } else {
        Write-Host "Operation cancelled."
    }
}
# Load configuration from file
$configFile = "config.json"
$config = Get-Content $configFile | ConvertFrom-Json

$clientId = $config.clientId
$clientSecret = $config.clientSecret
$tenantId = $config.tenantId

# Get access token
$accessToken = Get-AccessToken -clientId $clientId -tenantId $tenantId -clientSecret $clientSecret

# $userId = Get-UserIdByEmail -accessToken $accessToken -email $email
$userId = "e7ac535a-2bc3-47ad-a9d5-35ac2c93fb04"
# $teamId = Get-TeamId -teamName $teamName -accessToken $accessToken -userId $userId
$teamId = "c5522a2d-bccf-4d2b-950c-252a06cf632f"

# Variables to track clock-in and clock-out state
$clockedIn = $false
$onBreak = $false
$timeCardId = $null

# Main loop to monitor Teams state
while ($true) {
    $userStatus = Get-TeamsStatus
    if (IsTeamsRunning) {
        if (-not $clockedIn) {
            # Attempt to clock in
            $timeCardId = ClockIn -teamId $teamId -accessToken $accessToken -userId $userId
            $clockedIn = $true
        }
        
        if ($userStatus -eq "Away" -or $userStatus -eq "BeRightBack") {
            # Start break if user is Away or BeRightBack and not already on a break
            if (-not $onBreak) {
                StartBreak -teamId $teamId -accessToken $accessToken -userId $userId
                $onBreak = $true
            }
        } elseif ($userStatus -ne "Offline" -and $onBreak) {
            # End break if user is not Offline and currently on a break
            EndBreak -teamId $teamId -accessToken $accessToken -userId $userId
            $onBreak = $false
        }
    } else {
        if ($clockedIn) {
            # Attempt to clock out
            ClockOut -teamId $teamId -timeCardId $timeCardId -accessToken $accessToken -userId $userId
            $clockedIn = $false
        }
    }
    Start-Sleep -Seconds 60 # Check every minute
}
