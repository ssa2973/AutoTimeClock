[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$global:authorization_code = $null
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

function Get-DelegatedAccessToken {
    param(
        [string]$clientId,
        [string]$tenantId,
        [string]$redirectUri
    )
    Get-AuthorizationCode -clientId $clientId -tenantId $tenantId -redirectUri $redirectUri

    $tokenEndpoint = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
    $body = @{
        grant_type   = "authorization_code"
        client_id    = $clientId
        redirect_uri = $redirectUri
        code         = $global:authorization_code
        scope        = "https://graph.microsoft.com/.default"
    }
    $headers = @{
        "Content-Type" = "application/x-www-form-urlencoded"
    }
    $response = Invoke-RestMethod -Uri $tokenEndpoint -Method Post -Body $body -Headers $headers
    return $response.access_token
}

function Get-AuthorizationCode {
    param (
        [string]$clientId,
        [string]$tenantId,
        [string]$redirectUri
    )
    $authCodeEndpoint = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/authorize?response_type=code&client_id=$clientId&state=12345&scope=offline_access%20https%3A%2F%2Fgraph.microsoft.com%2F.default&redirect_uri=$redirectUri&response_mode=query"
        
    Add-Type -AssemblyName System.Windows.Forms

    # Create a form with a web browser control
    $form = New-Object -TypeName System.Windows.Forms.Form
    $form.Size = New-Object -TypeName System.Drawing.Size(800, 600)
    $form.StartPosition = "CenterScreen"
    $form.Text = "Sign in to your Microsoft account"
    $form.TopMost = $true
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("src/logo.ico")
    $webBrowser = New-Object -TypeName System.Windows.Forms.WebBrowser
    $webBrowser.Dock = [System.Windows.Forms.DockStyle]::Fill
    $form.Controls.Add($webBrowser)
    
    # Navigate to the authorization code endpoint URL
    $webBrowser.Navigate($authCodeEndpoint)
    
    # Wait for the user to authorize and for the redirection to happen
    $handler = {
        param(
            [System.Object]$_sender,
            [System.EventArgs]$e
        )
        
        $url = $_sender.Url.AbsoluteUri
        if ($url -match 'code=(.*)&state=12345') {
            $global:authorization_code = $matches[1]
            $form.Close()
        }
    }
          
    # Register event handler for Navigated event
    $webBrowser.add_Navigated($handler)
    
    # Show the form
    $form.ShowDialog() | Out-Null
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
    [System.Windows.MessageBox]::Show("User not found with email: $email")
    Remove-Item -Path $ConfigPath -Force
    exit
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
        [System.Windows.MessageBox]::Show("Team '$teamName' not found. Restart the application and enter a valid team name.")
        Remove-Item -Path $ConfigPath -Force
        exit
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
    return $null -ne (Get-Process | Where-Object { $_.Name -eq "ms-teams" })
}
    
$mainOtp = Get-Random -Minimum 100000 -Maximum 999999
function Send-Otp {
    param (
        [string]$email,
        [string]$otp
    )
    # Load configuration from file
    $configFile = "src/mail-config.json"
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
    param (
        [string]$userId,
        [string]$accessToken
    )

    $apiUrl = "https://graph.microsoft.com/beta/communications/presences/$userId"
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

    # Calculate form size based on percentages
    $formWidth = ($screenWidth * $formWidthPercentage).ToString("0")  # Round to nearest integer
    $formHeight = ($screenHeight * $formHeightPercentage).ToString("0")
    $form.Size = New-Object System.Drawing.Size($formWidth, $formHeight)
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true  # Always on top
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("src/logo.ico")
    $form.AutoSize = $true
    $form.AutoSizeMode = "GrowAndShrink"
    $form.FormBorderStyle = "FixedSingle"
    
    # Create the reminder time label
    $reminderTimeLabel = New-Object System.Windows.Forms.Label
    $reminderTimeLabel.Text = "Remind to clock in after:"
    $reminderTimeLabel.Location = New-Object System.Drawing.Point(10, 20)
    $reminderTimeLabel.Width = $formWidth / 4
    $form.Controls.Add($reminderTimeLabel)

    $reminderTimeComboBoxLocationX = $reminderTimeLabel.Width + 10
    # Create a typable dropdown with default values
    $reminderTimeComboBox = New-Object System.Windows.Forms.ComboBox
    $reminderTimeComboBox.Location = New-Object System.Drawing.Point($reminderTimeComboBoxLocationX, 15)  # Adjusted position
    $reminderTimeComboBox.Width = $formWidth / 10
    $reminderTimeComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
    $reminderTimeComboBox.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
    $reminderTimeComboBox.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::ListItems
    $reminderTimeComboBox.Items.AddRange(@("10", "15", "30"))  # Default values
    $reminderTimeComboBox.Text = ""  # Initial value
    $reminderTimeComboBox.MaxLength = 3  # Limit input to 3 characters
    # Add the KeyPress event handler to restrict input to numbers only
    $reminderTimeComboBox.Add_KeyPress({
            param (
                [object] $_sender,
                [System.Windows.Forms.KeyPressEventArgs] $e
            )
    
            # Check if the key pressed is not a digit or a control key (like backspace)
            if (-not [char]::IsDigit($e.KeyChar) -and -not [char]::IsControl($e.KeyChar)) {
                $e.Handled = $true  # Prevent the character from being entered into the control
            }
        })

    $form.Controls.Add($reminderTimeComboBox)
    
    $minutesLabelLocationX = $reminderTimeComboBoxLocationX + $reminderTimeComboBox.Width + 10
    $minutesLabel = New-Object System.Windows.Forms.Label
    $minutesLabel.Text = "minutes"
    $minutesLabel.Location = New-Object System.Drawing.Point($minutesLabelLocationX, 20)  # Adjusted position
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
        [string]$subject,
        [string]$message,
        [string[]]$toRecipients,
        [string[]]$ccRecipients
    )
  
    # Load configuration from file
    $configFile = "src/mail-config.json"
    $config = Get-Content $configFile | ConvertFrom-Json
                
    $mailId = $config.mail
    $password = $config.password
    $smtpServer = $config.smtpServer
    $smtpPort = $config.smtpPort
                
    $credentials = New-Object -TypeName PSCredential -ArgumentList $mailId, ($password | ConvertTo-SecureString -AsPlainText -Force)
    $mail = @{
        From       = $mailId
        To         = $toRecipients
        Cc         = $ccRecipients
        Subject    = $subject
        Body       = $message
        SmtpServer = $smtpServer
        Port       = $smtpPort
        UseSsl     = $true
        Credential = $credentials
    }
    Send-MailMessage @mail -BodyAsHtml    
}

function Send-MailAsUser {
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

function Get-ClockedInSession {
    param (
        [string]$teamId,
        [string]$accessToken,
        [string]$userId
    )

    $apiUrl = "https://graph.microsoft.com/beta/teams/$teamId/schedule/timeCards"
    $headers = @{
        Authorization    = "Bearer $accessToken"
        "MS-APP-ACTS-AS" = $userId
    }

    $response = Invoke-RestMethod -Uri $apiUrl -Method Get -Headers $headers
    $timeCard = $response.value | Where-Object { $_.state -ne "clockedOut" -and $_.userId -eq $userId }
    return $timeCard
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
        $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("src/logo.ico")
        $form.TopMost = $true  # Set the form to appear in the foreground
        $result = [System.Windows.Forms.MessageBox]::Show($form, "It's " + (Get-Date -Format "HH:mm") + " right now. Do you want to Clock In?", "Clock In", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)    
    }

    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $response = Invoke-RestMethod -Uri $apiUrl -Method Post -Headers $headers -ContentType "application/json"
        return $response
    }
    else {
        $selectedReminderTime = Start-ClockInReminder
        if ($null -ne $selectedReminderTime) {
            Start-Sleep -Seconds ($selectedReminderTime * 60)
            if ((Get-Date) -ge $accessTokenExpiration -or $null -eq $accessToken) {
                $accessToken = Get-AccessToken -clientId $clientId -tenantId $tenantId -clientSecret $clientSecret
                $accessTokenExpiration = (Get-Date).AddSeconds(3599)
            }        
            if (IsTeamsRunning) {
                $timeCard = ClockIn -teamId $teamId -accessToken $accessToken -userId $userId  
            }
            if ($timeCard) {
                return $timeCard
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
    
    if (![string]::IsNullOrEmpty($timeCardId)) {
        # Prompt the user with a MessageBox
        Add-Type -AssemblyName System.Windows.Forms
        $form = New-Object System.Windows.Forms.Form
        $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("src/logo.ico")
        $form.TopMost = $true  # Set the form to appear in the foreground
        # Loop until user clocks out or cancels repeatedly
        if ((Get-Date) -ge $accessTokenExpiration -or $null -eq $accessToken) {
            $accessToken = Get-AccessToken -clientId $clientId -tenantId $tenantId -clientSecret $clientSecret
            $accessTokenExpiration = (Get-Date).AddSeconds(3599)
        }
        $apiUrl = "https://graph.microsoft.com/beta/teams/$teamId/schedule/timeCards/$timeCardId/clockOut"
        $headers = @{
            Authorization    = "Bearer $accessToken"
            "MS-APP-ACTS-AS" = $userId
        }
        $result = [System.Windows.Forms.MessageBox]::Show($form, "It's " + (Get-Date -Format "HH:mm") + " right now. Do you want to Clock Out?", "Clock Out", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)  
      
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            Invoke-RestMethod -Uri $apiUrl -Method Post -Headers $headers
            return $true
            # Exit the loop after clocking out
            break
        }
        else {
            # Display message and wait 10 seconds before prompting again
            Write-Host "Clock Out cancelled."
            Start-Sleep -Seconds 10
            return $false
        }
    }
}

function Get-PresenceSubscription {
    param(
        [string]$accessToken
    )
    $apiUrl = "https://graph.microsoft.com/beta/subscriptions"
    $headers = @{
        Authorization = "Bearer $accessToken"
    }
    $response = Invoke-RestMethod -Uri $apiUrl -Method Get -Headers $headers
    return $response.value[0].id
}

function New-PresenceSubscription {
    param(
        [string]$accessToken,
        [string]$userId,
        [string]$tenantId
    )
    $apiUrl = "https://graph.microsoft.com/beta/subscriptions" 
    $headers = @{
        Authorization  = "Bearer $accessToken"
        "Content-Type" = "application/json"
    }
    $body = @{
        changeType         = "updated"
        notificationUrl    = "$webhookUrl/notifications"
        resource           = "communications/presences/$userId"
        expirationDateTime = [DateTime]::UtcNow.AddHours(1).ToString("yyyy-MM-dd'T'HH:mm:ss.fffffffZ")
        clientState        = "secret"
        tenantId           = $tenantId
    }
    $jsonBody = $body | ConvertTo-Json
    try {
        $response = Invoke-RestMethod -Uri $apiUrl -Method Post -Headers $headers -Body $jsonBody
        return $response.id
    }
    catch {
        Write-Host "Error creating subscription: $_"
    }
}

function Update-PresenceSubscription {
    param(
        [string]$subscriptionId,
        [string]$accessToken
    )
    $apiUrl = "https://graph.microsoft.com/beta/subscriptions/$subscriptionId"
    $headers = @{
        Authorization  = "Bearer $accessToken"
        "Content-Type" = "application/json"
    }
    $body = @{
        expirationDateTime = [DateTime]::UtcNow.AddHours(1).ToString("yyyy-MM-dd'T'HH:mm:ss.fffffffZ")
        notificationUrl    = "$webhookUrl/notifications"
    }
    $jsonBody = $body | ConvertTo-Json
    try {
        Invoke-RestMethod -Uri $apiUrl -Method Patch -Headers $headers -Body $jsonBody | Out-Null
    }
    catch {
        Write-Host "Error updating subscription: $_"
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
    $otpForm.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create("src/logo.ico")
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

# Unique process name
$processName = "AutoTimeClock"
$process = Get-Process -Name $processName -ErrorAction SilentlyContinue
# Check if another instance is running
if ($process.Count -gt 1) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Error"
    $form.Size = New-Object System.Drawing.Size(300, 200)
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("src/logo.ico")
    $form.StartPosition = 'CenterScreen'
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $label.Size = New-Object System.Drawing.Size(280, 20)
    $label.Text = "An instance of the application is already running."
    $form.Controls.Add($label)
    $form.ShowDialog()
    exit
}
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
# Define desired form size as a percentage of screen size
$formWidthPercentage = 0.3  # 30% of screen width
$formHeightPercentage = 0.25  # 25% of screen height

$screenWidth = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Width
$screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height

# Load necessary assemblies
Add-Type -AssemblyName PresentationFramework

# Create the form
$form = New-Object System.Windows.Window
$form.Title = "Enter Details"
$form.Width = ($screenWidth * $formWidthPercentage).ToString("0")  # Round to nearest integer
$form.Height = ($screenHeight * $formHeightPercentage).ToString("0")
$form.WindowStartupLocation = "CenterScreen"
$form.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create("src/logo.ico")
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
$emailTextBox.Width = $form.Width / 3
$emailTextBox.Height = $form.Height / 9
$emailTextBox.Margin = "0,0,25,0"

# Create the team name label and text box
$teamNameLabel = New-Object System.Windows.Controls.Label
$teamNameLabel.Content = "Team Name:"
$teamNameLabel.Margin = "50,0,0,0"
$teamNameTextBox = New-Object System.Windows.Controls.TextBox
$teamNameTextBox.Width = $form.Width / 3
$teamNameTextBox.Height = $form.Height / 9
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
        exit
    }
}

# Load configuration from file
$configFile = "src/app-config.json"
$config = Get-Content $configFile | ConvertFrom-Json

$clientId = $config.clientId
$clientSecret = $config.clientSecret
$tenantId = $config.tenantId
$redirectUri = $config.redirectUri
$webhookUrl = $config.webhookUrl

# Define the WebSocket URL
$webSocketUrl = $webhookUrl.Replace("http", "ws")
# $webSocketUrl = "ws://localhost:3000"

# Create a ClientWebSocket instance
$webSocket = New-Object System.Net.WebSockets.ClientWebSocket
# Function to receive WebSocket messages
function Receive-WebSocketMessage {
    param (
        [Parameter(Mandatory = $true)]
        [System.Net.WebSockets.ClientWebSocket]$socket,
        [Parameter(Mandatory = $true)]
        [string]$userId
    )
    $buffer = New-Object byte[] 4096
    $segment = [System.ArraySegment[byte]]::new($buffer)
    
    $result = $webSocket.ReceiveAsync($segment, [System.Threading.CancellationToken]::None).Result

    if ($result.CloseStatus -ne [System.Nullable[System.Net.WebSockets.WebSocketCloseStatus]]::null) {
        Write-Host "WebSocket closed with status: $($result.CloseStatus)"
    }

    $message = [System.Text.Encoding]::UTF8.GetString($buffer, 0, $result.Count)
    if ($message -ne "" -and $null -ne $message) {
        Write-Host "Received message: $message"
    }
    $jsonObject = $message | ConvertFrom-Json
    $id = $jsonObject.id
    if ($id -eq $userId) {
        $userStatus = $jsonObject[-1].availability
    }
    return $userStatus
}

# Get access token
$accessTokenExpiration = (Get-Date).AddSeconds(3599)
$accessToken = Get-AccessToken -clientId $clientId -tenantId $tenantId -clientSecret $clientSecret
$delegatedToken = Get-DelegatedAccessToken -clientId $clientId -tenantId $tenantId -redirectUri $redirectUri

$userId = Get-UserIdByEmail -accessToken $accessToken -email $email
$teamId = Get-TeamId -teamName $teamName -accessToken $accessToken -userId $userId
$ownerMails = Get-Owners -teamId $teamId -accessToken $accessToken

$timeCard = Get-ClockedInSession -teamId $teamId -accessToken $accessToken -userId $userId

$clockedIn = $false
$clockInTime = $null
$timeCardId = $timeCard.id
# Variables to track clock-in and clock-out state
if ($null -ne $timeCard.id) {
    $clockedIn = $true
    $clockInTime = [dateTime]$timeCard.clockInEvent.dateTime
}
$onBreak = $false
$clockOutTime = $null
$breakStartTime = $null
$breakEndTime = $null
$breaksDuration = 0

$presenceSubscriptionId = Get-PresenceSubscription -accessToken $delegatedToken

if ($null -eq $presenceSubscriptionId) {
    $presenceSubscriptionId = New-PresenceSubscription -accessToken $delegatedToken -userId $userId -tenantId $tenantId
}
else {
    Update-PresenceSubscription -subscriptionId $presenceSubscriptionId -accessToken $delegatedToken
}
$userStatus = Get-TeamsStatus -userId $userId -accessToken $accessToken

# Main loop to monitor Teams state
while ($null -ne $userId -and $null -ne $teamId) {
    if ((Get-Date) -ge $accessTokenExpiration -or $null -eq $accessToken) {
        $accessToken = Get-AccessToken -clientId $clientId -tenantId $tenantId -clientSecret $clientSecret
        $accessTokenExpiration = (Get-Date).AddSeconds(3599)
        $delegatedToken = Get-DelegatedAccessToken -clientId $clientId -tenantId $tenantId -redirectUri $redirectUri
        Update-PresenceSubscription -subscriptionId $presenceSubscriptionId -accessToken $delegatedToken
    }
    # $userStatus = Get-TeamsStatus -userId $userId
    if (IsTeamsRunning) {
        try {
            if ($webSocket.State -ne [System.Net.WebSockets.WebSocketState]::Open) {
                $webSocket.ConnectAsync([System.Uri]$webSocketUrl, [System.Threading.CancellationToken]::None).Wait()
            }
        } catch {
            if ($_.Exception.InnerException.Message -eq "The WebSocket has already been started.") {
                Write-Host "WebSocket already connected."
            } else {
                Write-Host "WebSocket connection error: $_"
            }
        }
        $userStatus = Receive-WebSocketMessage -socket $webSocket -userId $userId
        if ($null -eq $timeCard) {
            $timeCard = Get-ClockedInSession -teamId $teamId -accessToken $accessToken -userId $userId
            $timeCardId = $timeCard.id
            if ($null -ne $timeCard.id) {
                $clockedIn = $true
                $clockInTime = [dateTime]$timeCard.clockInEvent.dateTime
            }        
        }
        if (-not $clockedIn) {
            # Attempt to clock in
            $timeCard = ClockIn -teamId $teamId -accessToken $accessToken -userId $userId
            if ($null -ne $timeCard.id) {
                $clockedIn = $true
                $clockInTime = [dateTime]$timeCard.clockInEvent.dateTime
                $timeCardId = $timeCard.id
                Send-MailAsUser -userId $userId -accessToken $accessToken -subject "Clock in update for $email" -message "User $email has successfully clocked in at $(Get-Date) in $teamName" -toRecipients $ownerMails -ccRecipients $email
            }
        }
        
        if ($null -eq $userStatus) {
            continue
        }
        
        elseif ($userStatus -eq "Away" -or $userStatus -eq "BeRightBack") {
            # Start break if user is Away or BeRightBack and not already on a break
            if (-not $onBreak -and $clockedIn) {
                StartBreak -teamId $teamId -timeCardId $timeCardId -accessToken $accessToken -userId $userId
                $onBreak = $true
                Send-MailAsUser -userId $userId -subject "Break update for $email" -message "User $email has started a break at $(Get-Date)" -toRecipients $ownerMails -ccRecipients $email -accessToken $accessToken
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
            Send-MailAsUser -userId $userId -subject "Break update for $email" -message "User $email has ended a break at $(Get-Date). Break duration - $duration" -toRecipients $ownerMails -ccRecipients $email -accessToken $accessToken
        }
    }
    else {
        if ($clockedIn) {
            if ($onBreak) {
                EndBreak -teamId $teamId -timeCardId $timeCardId -accessToken $accessToken -userId $userId
                $onBreak = $false
                $breakEndTime = Get-Date
                $duration = $breakEndTime - $breakStartTime
                $breaksDuration += $duration
                Send-MailAsUser -userId $userId -subject "Break update for $email" -message "User $email has ended a break at $(Get-Date). Break duration - $duration" -toRecipients $ownerMails -ccRecipients $email -accessToken $accessToken
            }
            # Attempt to clock out
            $result = ClockOut -teamId $teamId -timeCardId $timeCardId -accessToken $accessToken -userId $userId
            if ($result[-1]) {
                $clockedIn = $false
                $clockOutTime = Get-Date
                $duration = $clockOutTime - $clockInTime
                $activeDuration = $duration - $breaksDuration
                Send-MailAsUser -userId $userId -subject "Clock out update for $email" -message "User $email has successfully clocked out at $(Get-Date) in $teamName. Total duration - $duration, Active duration - $activeDuration" -toRecipients $ownerMails -ccRecipients $email -accessToken $accessToken
                $timeCard = $null
                $timeCardId = $null
            }
            $result = $null
            $webSocket.CloseAsync([System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure, "Closing connection", [System.Threading.CancellationToken]::None).Wait()
        }
    }
    # Start-Sleep -Seconds 60 # Check every minute
}
