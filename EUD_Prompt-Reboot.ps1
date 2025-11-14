<#
.SYNOPSIS
    PowerShell script to manage system restart prompts with graphical user interface (GUI) components.
.DESCRIPTION
    This script loads the necessary assemblies for GUI components and manages system restart prompts by displaying a dialog box.
    Users can either restart immediately or defer the restart multiple times for a defined duration. After reaching
    the deferral limit, the script forces a system restart.
.NOTES
    Version : 4.0
    Template version for public distribution
#>

# ============================================================================
# CONFIGURATION SECTION - Customize these values for your organization
# ============================================================================

# Company/Organization Information
$CompanyName = "Your Company Name"
$CompanyITDepartment = "IT Services / Services Informatiques"  # Bilingual department name
$WindowTitle = "$CompanyName - $CompanyITDepartment"

# UI Colors (Hex format: #RRGGBB)
$ColorBackground = "#C0C0C0"      # Main background color
$ColorButtonBackground = "#C0C0C0" # Button background color
$ColorText = "#000000"             # Text color

# Font Configuration
$FontFamily = "Aptos"
$FontSize = 10
$FontSizeLarge = 20
$FontSizeCountdown = 32

# User Messages (Bilingual - English / French)
$MessageWarning = @"
WARNING / ATTENTION

This message is sent to keep the computer park healthy.
Your workstation requires a restart.


Ce message est envoye afin de maintenir la sante du parc informatique.
Votre poste de travail necessite un redemarrage.

"@

$MessageRestartNow = "Restart now`nRedemarrer maintenant"
$MessageRemindMe = "Remind me in 2 hours`nRappeler dans 2 heures"
$MessageRebootImminent = "Reboot imminent...`nRedemarrage en cours..."
$MessageRebootInProgress = "Reboot in progress...`nRedemarrage en cours..."
$MessageAutoRestart = "Auto-restart at: / Redemarrage automatique a : {0}"

# Timing Configuration
$SnoozeLimit = 13          # Maximum number of snoozes (13 × 2 hours = 24 hours max)
$SnoozeIncrement = 2*60*60 # 2 hours in seconds per deferral
$InitialDelay = 24*60*60   # 24 hours in seconds - ABSOLUTE DEADLINE
$RestartTimeout = 15 * 60  # 15 minutes before restart (in seconds)

# Logging Configuration
$LogDir = "C:\Temp"
$LogFileName = "Prompt-Reboot.log"

# ============================================================================
# END OF CONFIGURATION SECTION
# ============================================================================

# Load necessary assemblies for GUI components
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global variables for deferral management
$global:SnoozeCount = 0

# Logging setup
$LogFile = Join-Path $LogDir $LogFileName
if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory -Force | Out-Null }
function Write-Log {
    param([string]$Message)
    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Add-Content -Path $LogFile -Value "[$timestamp] $Message"
    } catch {}
}

# Global variable for planned restart time
$global:PlannedRestart = (Get-Date).AddSeconds($InitialDelay)
Write-Log "Script started. Planned restart at $($global:PlannedRestart.ToString('yyyy-MM-dd HH:mm:ss'))"

# Function to display restart prompt with timer and progress bar
function Show-RestartPrompt {
    # Calculate remaining time based on global planned restart
    $remainingTotal = [int]([TimeSpan]($global:PlannedRestart - (Get-Date))).TotalSeconds
    if ($remainingTotal -lt 0) { $remainingTotal = 0 }
    
    $script:CurrentSessionStart = $remainingTotal
    
    # Create new form for restart prompt
    $form = New-Object System.Windows.Forms.Form
    $form.TopMost = $true
    $form.Text = $WindowTitle
    $form.Size = New-Object System.Drawing.Size(600, 420)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.ControlBox = $false
    $form.BackColor = [System.Drawing.ColorTranslator]::FromHtml($ColorBackground)

    # Define font
    $font = New-Object System.Drawing.Font($FontFamily, $FontSize)

    # Create label for message text
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $label.Size = New-Object System.Drawing.Size(580, 160)
    $label.Text = $MessageWarning
    $label.Font = $font
    $label.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($ColorText)
    $label.TextAlign = [System.Drawing.ContentAlignment]::TopCenter

    # Create countdown label
    $timeLabel = New-Object System.Windows.Forms.Label
    $timeLabel.Font = New-Object System.Drawing.Font($FontFamily, $FontSizeLarge, [System.Drawing.FontStyle]::Bold)
    $timeLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($ColorText)
    $timeLabel.Size = New-Object System.Drawing.Size(580, 40)
    $timeLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $timeLabel.Location = New-Object System.Drawing.Point(10, 170)
    $ts = [TimeSpan]::FromSeconds($remainingTotal)
    $timeLabel.Text = "{0:00}:{1:00}:{2:00}" -f $ts.Hours, $ts.Minutes, $ts.Seconds

    # Create ETA label
    $etaLabel = New-Object System.Windows.Forms.Label
    $etaLabel.Font = $font
    $etaLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($ColorText)
    $etaLabel.Size = New-Object System.Drawing.Size(580, 25)
    $etaLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $etaLabel.Location = New-Object System.Drawing.Point(10, 215)
    $etaLabel.Text = $MessageAutoRestart -f $global:PlannedRestart.ToString("yyyy-MM-dd HH:mm:ss")

    # Create progress bar
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Style = 'Continuous'
    $progressBar.Minimum = 0
    $progressBar.Maximum = $InitialDelay
    $progressBar.Value = $InitialDelay - $remainingTotal
    $progressBar.Size = New-Object System.Drawing.Size(560, 20)
    $progressBar.Location = New-Object System.Drawing.Point(20, 240)

    # Create buttons for user interaction
    $buttonYes = New-Object System.Windows.Forms.Button
    $buttonYes.Location = New-Object System.Drawing.Point(100, 280)
    $buttonYes.Size = New-Object System.Drawing.Size(200, 50)
    $buttonYes.Text = $MessageRestartNow
    $buttonYes.Font = $font
    $buttonYes.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($ColorText)
    $buttonColor = [System.Drawing.ColorTranslator]::FromHtml($ColorButtonBackground)
    $buttonYes.BackColor = [System.Drawing.Color]::FromArgb($buttonColor.R, $buttonColor.G, $buttonColor.B)
    $buttonYes.TabIndex = 1
    $buttonYes.Add_Click({ 
        $timer.Stop()
        $form.Tag = "Yes"
        $form.Close() 
    })

    $buttonNo = New-Object System.Windows.Forms.Button
    $buttonNo.Location = New-Object System.Drawing.Point(300, 280)
    $buttonNo.Size = New-Object System.Drawing.Size(200, 50)
    $buttonNo.Text = $MessageRemindMe
    $buttonNo.Font = $font
    $buttonNo.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($ColorText)
    $buttonNo.BackColor = [System.Drawing.Color]::FromArgb($buttonColor.R, $buttonColor.G, $buttonColor.B)
    $buttonNo.TabIndex = 0
    $buttonNo.Add_Click({ 
        $timer.Stop()
        $form.Tag = "Remind"
        $form.Close() 
    })

    # Disable deferral button if no more deferrals available
    if ($global:SnoozeCount -ge $SnoozeLimit) {
        $buttonNo.Enabled = $false
    }

    # Add controls to form
    $form.Controls.Add($label)
    $form.Controls.Add($timeLabel)
    $form.Controls.Add($etaLabel)
    $form.Controls.Add($progressBar)
    $form.Controls.Add($buttonYes)
    $form.Controls.Add($buttonNo)

    # Set default action on Enter and intercept Enter for $buttonNo
    $form.AcceptButton = $buttonNo
    $form.KeyPreview = $true
    $form.Add_KeyDown({
        param($src, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $buttonNo.PerformClick()
        }
    })

    # Create and configure timer
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 1000 # 1 second

    $timer.Add_Tick({
        $remaining = [int]([TimeSpan]($global:PlannedRestart - (Get-Date))).TotalSeconds
        if ($remaining -lt 0) { $remaining = 0 }

        # Update time display
        $ts = [TimeSpan]::FromSeconds($remaining)
        $timeLabel.Text = "{0:00}:{1:00}:{2:00}" -f $ts.Hours, $ts.Minutes, $ts.Seconds

        # Update progress bar (based on total 24h time)
        $elapsed = $InitialDelay - $remaining
        if ($elapsed -lt 0) { $elapsed = 0 }
        if ($elapsed -gt $progressBar.Maximum) { $elapsed = $progressBar.Maximum }
        $progressBar.Value = $elapsed

        # If time has expired, show final countdown
        if ($remaining -le 0) {
            $timer.Stop()
            $form.Tag = "TimeExpired"
            $form.Close()
        }
    })

    # Handle form closing (X or Alt+F4)
    $form.Add_FormClosing({
        param($src, $e)
        
        # If user closes form without clicking a button
        if ([string]::IsNullOrEmpty($form.Tag)) {
            $timer.Stop()
            
            # Check if enough time remains for another snooze
            $remainingUntilDeadline = [int]([TimeSpan]($global:PlannedRestart - (Get-Date))).TotalSeconds
            
            if ($global:SnoozeCount -ge $SnoozeLimit -or $remainingUntilDeadline -le $SnoozeIncrement) {
                # No snooze possible - force final countdown display
                $e.Cancel = $true  # Cancel closing
                $timer.Stop()
                $form.Hide()  # Hide current form
                Write-Log "User closed the window without action."
                
                $rebootMessage = $MessageRebootImminent

                # Wait until deadline if necessary
                if ($remainingUntilDeadline -gt 0 -and $remainingUntilDeadline -le $SnoozeIncrement) {
                    Write-Log "Not enough time for another snooze. Waiting for deadline"
                    Start-Sleep -Seconds $remainingUntilDeadline
                }
                Write-Log "Starting final $($RestartTimeout/60)-minute countdown"
                Show-FinalCountdown -Message $rebootMessage
                Restart-Computer -Force
            } else {
                # Treat as normal snooze
                Write-Log "User clicked to be reminded"
                $form.Tag = "Remind"
            }
        }
    })

    # Start timer when form is shown and set focus on $buttonNo
    $form.Add_Shown({ 
        $timer.Start()
        $buttonNo.Select()
        $buttonNo.Focus()
    })

    # Display form as dialog
    $form.ShowDialog() | Out-Null

    # Act based on user response
    if ($form.Tag -eq "Yes") {
        Write-Log "User clicked to restart now"
        $immediateRestartTimeout = 30 # Give a 30 seconds timeout when user clicks yes
        $rebootMessage = $MessageRebootInProgress
        Show-FinalCountdown -Message $rebootMessage -Timeout $immediateRestartTimeout
        Restart-Computer -Force
    } elseif ($form.Tag -eq "TimeExpired") {
        # Initial delay expired - show final countdown
        $rebootMessage = $MessageRebootInProgress
        Write-Log "Initial delay expired; starting final $($RestartTimeout/60)-minute countdown"
        Show-FinalCountdown -Message $rebootMessage
        Restart-Computer -Force
    } else {
        Snooze
    }
}

function Snooze {
    $global:SnoozeCount++
    
    # Check how much time remains before absolute 24h deadline
    $remainingUntilDeadline = [int]([TimeSpan]($global:PlannedRestart - (Get-Date))).TotalSeconds
    
    if ($global:SnoozeCount -ge $SnoozeLimit) {
        # Deferral limit reached - show final countdown
        $rebootMessage = $MessageRebootImminent
        Write-Log "Starting final $($RestartTimeout/60)-minute countdown"
        Show-FinalCountdown -Message $rebootMessage
        Restart-Computer -Force
    } elseif ($remainingUntilDeadline -le $SnoozeIncrement) {
        # Not enough time for another complete snooze
        # Wait until deadline, then show final countdown
        if ($remainingUntilDeadline -gt 0) {
            Write-Log "Not enough time for another snooze. Waiting for deadline"
            Start-Sleep -Seconds $remainingUntilDeadline
        }
        $rebootMessage = $MessageRebootImminent
        Write-Log "Starting final $($RestartTimeout/60)-minute countdown"
        Show-FinalCountdown -Message $rebootMessage
        Restart-Computer -Force
    } else {
        # Wait 2 hours before showing prompt again
        Write-Log "User clicked to be reminded"
        Start-Sleep -Seconds $SnoozeIncrement
        Show-RestartPrompt
    }
}

function Show-FinalCountdown {
    param(
        [string]$Message,
        [int]$Timeout = $RestartTimeout
    )

    $form = New-Object System.Windows.Forms.Form
    $form.TopMost = $true
    $form.Text = $WindowTitle
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.ControlBox = $false
    $form.BackColor = [System.Drawing.ColorTranslator]::FromHtml($ColorBackground)
    $form.Font = New-Object System.Drawing.Font($FontFamily, $FontSize)
    $form.Size = New-Object System.Drawing.Size(500, 180)
    $form.StartPosition = "CenterScreen"

    # Main message label
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Message
    $label.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($ColorText)
    $label.Font = New-Object System.Drawing.Font($FontFamily, 11, [System.Drawing.FontStyle]::Bold)
    $label.Size = New-Object System.Drawing.Size(480, 60)
    $label.TextAlign = [System.Drawing.ContentAlignment]::TopCenter
    $label.Location = New-Object System.Drawing.Point(10, 15)
    $form.Controls.Add($label)

    # Countdown label
    $countdownLabel = New-Object System.Windows.Forms.Label
    $countdownLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($ColorText)
    $countdownLabel.Font = New-Object System.Drawing.Font($FontFamily, $FontSizeCountdown, [System.Drawing.FontStyle]::Bold)
    $countdownLabel.Size = New-Object System.Drawing.Size(480, 70)
    $countdownLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $countdownLabel.Location = New-Object System.Drawing.Point(10, 75)
    $form.Controls.Add($countdownLabel)

    # Timer for countdown
    $script:remainingSeconds = $Timeout
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 1000
    $timer.Add_Tick({
        $script:remainingSeconds--
        $minutes = [Math]::Floor($script:remainingSeconds / 60)
        $seconds = $script:remainingSeconds % 60
        $countdownLabel.Text = "{0:00}:{1:00}" -f $minutes, $seconds
        
        if ($script:remainingSeconds -le 0) {
            $timer.Stop()
            $form.Close()
        }
    })

    $form.Add_Shown({ 
        $minutes = [Math]::Floor($script:remainingSeconds / 60)
        $seconds = $script:remainingSeconds % 60
        $countdownLabel.Text = "{0:00}:{1:00}" -f $minutes, $seconds
        Write-Log "Final countdown started"
        $timer.Start()
    })

    $form.ShowDialog() | Out-Null
}

# Start the restart management process
Show-RestartPrompt
