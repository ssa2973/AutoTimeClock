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
        [string] $accessToken,
        [string] $email
    )
  
    $apiUrl = "https://graph.microsoft.com/v1.0/users"  
    $headers = @{
        Authorization = "Bearer $accessToken"
    }
  
    $allUsers = @()
  
    do {
        try {
            $response = Invoke-RestMethod -Uri $apiUrl -Method Get -Headers $headers -ErrorAction Stop
  
            # Check if any users are found
            if ($response.value.Count -gt 0) {
                $allUsers += $response.value
  
                # Look for @odata.nextLink for subsequent pages
                $apiUrl = $response."@odata.nextLink"
            }
            else {
                # No users found in this page, exit the loop
                $apiUrl = $null
            }
        }
        catch {
            Write-Error "Error retrieving users: $_.Exception.Message"
            $apiUrl = $null  # Exit the loop on errors
        }
    } while ($apiUrl)
  
    # Search for the user among all retrieved users
    foreach ($user in $allUsers) {
        if ($user.mail -eq $email) {
            return $user.id  # User found, return its ID
        }
    }
  
    # No user found with matching email
    Write-Host "User not found with email: $email"
    return $null
}

function Get-TeamId {
    param (
        [string]$teamName,
        [string]$accessToken,
        [string]$userId
    )
        
    $apiUrl = "https://graph.microsoft.com/v1.0/users/$userId/joinedTeams"
    $headers = @{
        Authorization = "Bearer $accessToken"
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

function Get-Owners {
    param (
        [string]$teamId,
        [string]$accessToken
    )
        
    $apiUrl = "https://graph.microsoft.com/v1.0/groups/$teamId/owners"
    $headers = @{
        Authorization = "Bearer $accessToken"
    }
        
    $response = Invoke-RestMethod -Uri $apiUrl -Method Get -Headers $headers
        
    if ($response.value.Count -ge 1) {
        $mails = $response.value | ForEach-Object { $_.mail }
        return $mails
    }
    else {
        return @()
    }
}
    
function IsTeamsRunning {
    return $null -ne (Get-Process | Where-Object { $_.Name -match "Teams" })
}
    
$mainOtp = Get-Random -Minimum 100000 -Maximum 999999
function Send-Otp {
    param (
        [string]$email,
        [string]$otp
    )
    # Load configuration from file
    $configFile = "mail-config.json"
    $config = Get-Content $configFile | ConvertFrom-Json
            
    $mailId = $config.mail
    $password = $config.password
    $smtpServer = $config.smtpServer
    $smtpPort = $config.smtpPort
            
    $credentials = New-Object -TypeName PSCredential -ArgumentList $mailId, ($password | ConvertTo-SecureString -AsPlainText -Force)
    $subject = "One-Time Password (OTP) for mail verification"
    $message = "Your OTP for verifying your email address is: $otp"
    Send-MailMessage -From $mailId -To $email -Subject $subject -Body $message -SmtpServer $smtpServer -Port $smtpPort -UseSsl -Credential $credentials
}

function Get-TeamsStatus {
    param(
        [string]$accessToken,
        [string]$userId
    )
                
    $apiUrl = "https://graph.microsoft.com/v1.0/users/$userId/presence"
    $headers = @{
        Authorization = "Bearer $accessToken"
    }
                
    $response = Invoke-RestMethod -Uri $apiUrl -Method Get -Headers $headers
    return $response.availability
}
            
function Start-ClockInReminder {
    # Create the reminder popup form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Set Reminder"
    $form.Size = New-Object System.Drawing.Size(400, 150)
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true  # Always on top
    
    # Create the reminder time label
    $reminderTimeLabel = New-Object System.Windows.Forms.Label
    $reminderTimeLabel.Text = "Remind to clock in after:"
    $reminderTimeLabel.Location = New-Object System.Drawing.Point(10, 20)
    $form.Controls.Add($reminderTimeLabel)
        
    # Create a typable dropdown with default values
    $reminderTimeComboBox = New-Object System.Windows.Forms.ComboBox
    $reminderTimeComboBox.Location = New-Object System.Drawing.Point(180, 15)  # Adjusted position
    $reminderTimeComboBox.Width = 80
    $reminderTimeComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
    $reminderTimeComboBox.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
    $reminderTimeComboBox.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::ListItems
    $reminderTimeComboBox.Items.AddRange(@("10", "15", "30"))  # Default values
    $reminderTimeComboBox.Text = ""  # Initial value
    $form.Controls.Add($reminderTimeComboBox)
    
    $minutesLabel = New-Object System.Windows.Forms.Label
    $minutesLabel.Text = "minutes"
    $minutesLabel.Location = New-Object System.Drawing.Point(265, 15)  # Adjusted position
    $form.Controls.Add($minutesLabel)
    
    # Create an Ok button (initially disabled)
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "Ok"
    $okButton.Location = New-Object System.Drawing.Point(120, 50)
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $okButton.Enabled = $false  # Initially disabled
    $form.Controls.Add($okButton)
        
    # Enable Ok button only for valid numbers
    $reminderTimeComboBox.Add_TextChanged({
            $okButton.Enabled = [int]($reminderTimeComboBox.Text) -ge 1  # Enable for numbers 1 or greater
        })
    
    $reminderTimeComboBox.Add_SelectedIndexChanged({
            $okButton.Enabled = $null -ne $reminderTimeComboBox.SelectedItem
        })
    $okButton.Add_Click({
            $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $form.Close()
        })
    
    $result = $form.ShowDialog()  # Capture the dialog result

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return [int]$reminderTimeComboBox.Text
    }
    else {
        return $null  # Return null for other results (Cancel, etc.)
    }
}

function SendMail {
    param(
        [string]$userId,
        [string]$accessToken,
        [string]$subject,
        [string]$message,
        [string[]]$toRecipients,
        [string[]]$ccRecipients
    )
  
    $apiUrl = "https://graph.microsoft.com/v1.0/users/$userId/sendMail"
    $headers = @{
        Authorization  = "Bearer $accessToken"
        "Content-Type" = "application/json"
    }
  
    $emailBody = @{
        message =
        @{
            subject      = $subject
            body         = @{
                contentType = "Text"
                content     = $message
            }
            toRecipients = @(
                foreach ($recipient in $toRecipients) {
                    @{
                        emailAddress = @{
                            address = $recipient
                        }
                    }
                }
            )
            ccRecipients = @(
                foreach ($recipient in $ccRecipients) {
                    @{
                        emailAddress = @{
                            address = $recipient
                        }
                    }
                }
            )
        }
    }

    $bodyJson = $emailBody | ConvertTo-Json -Depth 4
    Invoke-RestMethod -Uri $apiUrl -Method Post -Headers $headers -Body $bodyJson
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
        $form = New-Object System.Windows.Forms.Form
        $form.TopMost = $true  # Set the form to appear in the foreground
        $result = [System.Windows.Forms.MessageBox]::Show($form, "It's " + (Get-Date -Format "HH:mm") + " right now. Do you want to Clock In?", "Clock In", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)    
    }

    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $response = Invoke-RestMethod -Uri $apiUrl -Method Post -Headers $headers -ContentType "application/json"
        return $response.id
    }
    else {
        $selectedReminderTime = Start-ClockInReminder
        if ($null -ne $selectedReminderTime) {
            Start-Sleep -Seconds ($selectedReminderTime * 60)
            $timeCardId = ClockIn -teamId $teamId -accessToken $accessToken -userId $userId  
            if ($timeCardId) {
                return $timeCardId 
            }
        }
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
        $form = New-Object System.Windows.Forms.Form
        $form.TopMost = $true  # Set the form to appear in the foreground
        $result = [System.Windows.Forms.MessageBox]::Show($form, "It's " + (Get-Date -Format "HH:mm") + " right now. Do you want to Clock Out?", "Clock Out", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)    
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

function EndBreak {
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

function VerifyOtp {
    param(
        [string]$mainOtp
    )
    # Prompt for OTP verification
    $otpForm = New-Object System.Windows.Window
    $otpForm.Title = "OTP Verification"
    $otpForm.Width = 300
    $otpForm.Height = 150
    
    # Create a Grid to hold the child elements
    $otpGrid = New-Object System.Windows.Controls.Grid
    
    # Define rows for the grid
    $rowDefinition1 = New-Object System.Windows.Controls.RowDefinition
    $rowDefinition2 = New-Object System.Windows.Controls.RowDefinition
    $rowDefinition1.Height = [System.Windows.GridLength]::Auto
    $rowDefinition2.Height = [System.Windows.GridLength]::Auto
    $otpGrid.RowDefinitions.Add($rowDefinition1)
    $otpGrid.RowDefinitions.Add($rowDefinition2)
    
    # Define columns for the grid
    $columnDefinition1 = New-Object System.Windows.Controls.ColumnDefinition
    $columnDefinition1.Width = [System.Windows.GridLength]::Auto
    $otpGrid.ColumnDefinitions.Add($columnDefinition1)
    
    $columnDefinition2 = New-Object System.Windows.Controls.ColumnDefinition
    $columnDefinition2.Width = [System.Windows.GridLength]::Auto
    $otpGrid.ColumnDefinitions.Add($columnDefinition2)
    
    $columnDefinition3 = New-Object System.Windows.Controls.ColumnDefinition
    $columnDefinition3.Width = [System.Windows.GridLength]::Auto
    $otpGrid.ColumnDefinitions.Add($columnDefinition3)
    
    $otpLabel = New-Object System.Windows.Controls.Label
    $otpLabel.Content = "Enter OTP:"
    $otpLabel.Margin = "10,10,0,0"
    $otpLabel.VerticalAlignment = "Center"
    [System.Windows.Controls.Grid]::SetRow($otpLabel, 0)
    [System.Windows.Controls.Grid]::SetColumn($otpLabel, 0)
    
    $otpTextBox = New-Object System.Windows.Controls.TextBox
    $otpTextBox.Width = 100
    $otpTextBox.Height = 22.5
    $otpTextBox.Margin = "10,10,0,0"
    $otpTextBox.VerticalAlignment = "Center"
    [System.Windows.Controls.Grid]::SetRow($otpTextBox, 0)
    [System.Windows.Controls.Grid]::SetColumn($otpTextBox, 1)
    
    $otpSubmitButton = New-Object System.Windows.Controls.Button
    $otpSubmitButton.Content = "Submit"
    $otpSubmitButton.Width = 100
    $otpSubmitButton.Height = 30
    $otpSubmitButton.Margin = "10,10,0,0"
    $otpSubmitButton.VerticalAlignment = "Center"
    $otpSubmitButton.Add_Click({
            if ($otpTextBox.Text -eq $mainOtp) {
                $otpForm.DialogResult = $true
                $otpForm.Close()
            }
            else {
                [System.Windows.MessageBox]::Show("Invalid OTP. Please try again.")
            }
        })
    [System.Windows.Controls.Grid]::SetRow($otpSubmitButton, 1)
    [System.Windows.Controls.Grid]::SetColumn($otpSubmitButton, 1)
    
    # Add child elements to the Grid
    [void]$otpGrid.Children.Add($otpLabel)
    [void]$otpGrid.Children.Add($otpTextBox)
    [void]$otpGrid.Children.Add($otpSubmitButton)
    
    # Set the Grid as the content of the window
    $otpForm.Content = $otpGrid
    
    $otpForm.WindowStartupLocation = "CenterScreen"
    $otpForm.Topmost = $true
    
    return $otpForm.ShowDialog()
}

# Load necessary assemblies
Add-Type -AssemblyName PresentationFramework

# Create the form
$form = New-Object System.Windows.Window
$form.Title = "Enter Details"
$form.Width = 450
$form.Height = 200
$form.WindowStartupLocation = "CenterScreen"
$form.Topmost = $true

# Create a grid to hold the controls
$grid = New-Object System.Windows.Controls.Grid

# Define rows
$row1 = New-Object System.Windows.Controls.RowDefinition
$row2 = New-Object System.Windows.Controls.RowDefinition
$row3 = New-Object System.Windows.Controls.RowDefinition
$row4 = New-Object System.Windows.Controls.RowDefinition
$row1.Height = "Auto"  # Adjust height as needed
$row2.Height = "Auto"  # Adjust height as needed
$row3.Height = "Auto"  # Adjust height as needed
$row4.Height = "Auto"  # Adjust height as needed
$grid.RowDefinitions.Add($row1)
$grid.RowDefinitions.Add($row2)
$grid.RowDefinitions.Add($row3)
$grid.RowDefinitions.Add($row4)

# Define columns
$col1 = New-Object System.Windows.Controls.ColumnDefinition
$col2 = New-Object System.Windows.Controls.ColumnDefinition
$grid.ColumnDefinitions.Add($col1)
$grid.ColumnDefinitions.Add($col2)

# Create the email ID label and text box
$emailLabel = New-Object System.Windows.Controls.Label
$emailLabel.Content = "Email ID:"
$emailLabel.Margin = "50,0,0,0"
$emailTextBox = New-Object System.Windows.Controls.TextBox
$emailTextBox.Width = 180
$emailTextBox.Height = 22.5
$emailTextBox.Margin = "0,0,25,0"

# Create the team name label and text box
$teamNameLabel = New-Object System.Windows.Controls.Label
$teamNameLabel.Content = "Team Name:"
$teamNameLabel.Margin = "50,0,0,0"
$teamNameTextBox = New-Object System.Windows.Controls.TextBox
$teamNameTextBox.Width = 180
$teamNameTextBox.Height = 22.5
$teamNameTextBox.Margin = "0,0,25,0"

# Create the submit button
$submitButton = New-Object System.Windows.Controls.Button
$submitButton.Content = "Submit"
$submitButton.Width = 100
$submitButton.Height = 30
$submitButton.Margin = "0,35,0,0"
$submitButton.Add_Click({
        $form.DialogResult = $true
        $form.Close()
    })

# Add controls to the grid
[void]$grid.Children.Add($emailLabel)
[void]$grid.Children.Add($emailTextBox)
[void]$grid.Children.Add($teamNameLabel)
[void]$grid.Children.Add($teamNameTextBox)
[void]$grid.Children.Add($submitButton)

# Set Grid's row and column positions for each control
[System.Windows.Controls.Grid]::SetRow($emailLabel, 0)
[System.Windows.Controls.Grid]::SetColumn($emailLabel, 0)
[System.Windows.Controls.Grid]::SetRow($emailTextBox, 0)
[System.Windows.Controls.Grid]::SetColumn($emailTextBox, 1)
[System.Windows.Controls.Grid]::SetRow($teamNameLabel, 1)
[System.Windows.Controls.Grid]::SetColumn($teamNameLabel, 0)
[System.Windows.Controls.Grid]::SetRow($teamNameTextBox, 1)
[System.Windows.Controls.Grid]::SetColumn($teamNameTextBox, 1)
[System.Windows.Controls.Grid]::SetRow($submitButton, 3)
[System.Windows.Controls.Grid]::SetColumn($submitButton, 0)
[System.Windows.Controls.Grid]::SetColumnSpan($submitButton, 2)

# Add the grid to the form
$form.Content = $grid

# Set the configuration file path
$ConfigPath = "$env:userprofile\userconfig.xml"

# Check if the configuration file exists
if (Test-Path $ConfigPath) {
    # Load the configuration
    $Config = Import-Clixml $ConfigPath
    $email = $Config.email
    $teamName = $Config.teamName
}
else {
    # Show the form and wait for the user to submit
    $result = $form.ShowDialog()
    if ($result -eq $true) {
        $email = $emailTextBox.Text
        $teamName = $teamNameTextBox.Text
        Write-Host "Email: $email"
        Write-Host "Team Name: $teamName"
        
        # Send OTP to the entered email
        $otp = Send-Otp -email $email -otp $mainOtp
        $otpResult = VerifyOtp -mainOtp $mainOtp
        
        if ($otpResult -eq $true) {
            # The file does not exist, create it
            New-Item -Path $ConfigPath -ItemType File
            # Save the configuration to the XML file
            $Config = @{
                email    = $email
                teamName = $teamName
            }
            $Config | Export-Clixml -Path $ConfigPath
        }
        else {
            Write-Host "OTP validation failed. Please enter email and team name again."
            # Terminate the program
            exit
        }
    }
    else {
        Write-Host "Operation cancelled."
    }
}

# Load configuration from file
$configFile = "app-config.json"
$config = Get-Content $configFile | ConvertFrom-Json

$clientId = $config.clientId
$clientSecret = $config.clientSecret
$tenantId = $config.tenantId

# Get access token
$accessTokenExpiration = (Get-Date).AddSeconds(3599)
$accessToken = Get-AccessToken -clientId $clientId -tenantId $tenantId -clientSecret $clientSecret

$userId = Get-UserIdByEmail -accessToken $accessToken -email $email
$teamId = Get-TeamId -teamName $teamName -accessToken $accessToken -userId $userId
$ownerMails = Get-Owners -teamId $teamId -accessToken $accessToken

# Variables to track clock-in and clock-out state
$clockedIn = $false
$onBreak = $false
$timeCardId = $null
$clockInTime = $null
$clockOutTime = $null
$breakStartTime = $null
$breakEndTime = $null
$breaksDuration = 0

# Main loop to monitor Teams state
while ($null -ne $userId -and $null -ne $teamId) {
    if ((Get-Date) -ge $accessTokenExpiration) {
        $accessToken = Get-AccessToken -clientId $clientId -tenantId $tenantId -clientSecret $clientSecret
        $accessTokenExpiration = (Get-Date).AddSeconds(3599)
    }
    $userStatus = Get-TeamsStatus -accessToken $accessToken -userId $userId
    if (IsTeamsRunning) {
        if (-not $clockedIn) {
            # Attempt to clock in
            $timeCardId = ClockIn -teamId $teamId -accessToken $accessToken -userId $userId
            if ($null -ne $timeCardId) {
                $clockedIn = $true
                SendMail -userId $userId -accessToken $accessToken -subject "Clock in update" -message "User $email has successfully clocked in at $(Get-Date) in $teamName" -toRecipients $ownerMails -ccRecipients $email
                $clockInTime = Get-Date
            }
        }
        
        if ($userStatus -eq "Away" -or $userStatus -eq "BeRightBack") {
            # Start break if user is Away or BeRightBack and not already on a break
            if (-not $onBreak -and $clockedIn) {
                StartBreak -teamId $teamId -timeCardId $timeCardId -accessToken $accessToken -userId $userId
                $onBreak = $true
                SendMail -userId $userId -accessToken $accessToken -subject "Break update" -message "User $email has started a break at $(Get-Date)" -toRecipients $ownerMails -ccRecipients $email
                $breakStartTime = Get-Date
            }
        }
        elseif ($userStatus -ne "Offline" -and $onBreak -and $clockedIn) {
            # End break if user is not Offline and currently on a break
            EndBreak -teamId $teamId -timeCardId $timeCardId -accessToken $accessToken -userId $userId
            $onBreak = $false
            $breakEndTime = Get-Date
            $duration = $breakEndTime - $breakStartTime
            $breaksDuration += $duration
            SendMail -userId $userId -accessToken $accessToken -subject "Break update" -message "User $email has ended a break at $(Get-Date). Break duration - $duration" -toRecipients $ownerMails -ccRecipients $email
        }
    }
    else {
        if ($clockedIn) {
            # Attempt to clock out
            ClockOut -teamId $teamId -timeCardId $timeCardId -accessToken $accessToken -userId $userId
            $clockedIn = $false
            $clockOutTime = Get-Date
            $duration = $clockOutTime - $clockInTime
            $activeDuration = $duration - $breaksDuration
            SendMail -userId $userId -accessToken $accessToken -subject "Clock out update" -message "User $email has successfully clocked out at $(Get-Date) in $teamName. Total duration - $duration, Active duration - $activeDuration" -toRecipients $ownerMails -ccRecipients $email
        }
    }
    Start-Sleep -Seconds 60 # Check every minute
}
