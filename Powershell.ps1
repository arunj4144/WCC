# Windows Update Management Script - JUMPCLOUD INTERACTIVE VERSION
# Version 3.1 - "It Just Works" Edition
# Combines the proven Scheduled Task prompt with the reliable shutdown.exe reboot command.

# --- SCRIPT CONFIGURATION ---
param(
    [string]$GoogleSheetsWebhookUrl = "google sheets webhook url here"
)

# --- MODULE INSTALLATION & GLOBAL VARIABLES ---
try { Import-Module PSWindowsUpdate -ErrorAction Stop } catch { Write-Host "PSWindowsUpdate module not found. Installing..." -ForegroundColor Yellow; Install-Module PSWindowsUpdate -Force -Scope AllUsers -AcceptLicense; Import-Module PSWindowsUpdate -Force }
$ComputerName = $env:COMPUTERNAME; $UserName = $env:USERNAME; $CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name; $ScriptStartTime = Get-Date; $LogFile = "$env:TEMP\WindowsUpdate_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"; $ScriptVersion = "3.1-ItJustWorks"

# --- HELPER FUNCTIONS ---
function Write-Log { param([string]$Message, [string]$Level = "INFO") $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"; $LogEntry = "[$Timestamp] [$Level] $Message"; Add-Content -Path $LogFile -Value $LogEntry; Write-Host $LogEntry -ForegroundColor $(if($Level -eq "ERROR"){"Red"} elseif($Level -eq "WARNING"){"Yellow"} else{"White"}) }
function Get-SystemInfo { $os = Get-CimInstance -ClassName Win32_OperatingSystem; $cs = Get-CimInstance -ClassName Win32_ComputerSystem; $lastBoot = "N/A"; try { $lastBoot = ([System.Management.ManagementDateTimeConverter]::ToDateTime($os.LastBootUpTime)).ToString("u") } catch { try { $lastBoot = ([datetime]::Parse($os.LastBootUpTime)).ToString("u"); Write-Log "Parsed LastBootUpTime using fallback method." } catch { Write-Log "Could not parse LastBootUpTime: '$($os.LastBootUpTime)'" "WARNING" } }; return @{ ComputerName = $ComputerName; UserName = $UserName; CurrentUser = $CurrentUser; OSVersion = $os.Caption; OSBuild = $os.BuildNumber; LastBootTime = $lastBoot; Domain = $cs.Domain; Manufacturer = $cs.Manufacturer; Model = $cs.Model; TotalRAM = [math]::Round($cs.TotalPhysicalMemory/1GB, 2); IPAddress = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object {$_.InterfaceAlias -notlike "*Loopback*" -and $_.IPAddress -notlike "169.254.*"}).IPAddress -join ", " } }
function Get-AvailableUpdates { Write-Log "Scanning for available Windows updates..."; try { return Get-WUList -MicrosoftUpdate } catch { Write-Log "Error scanning for updates: $($_.Exception.Message)" "ERROR"; return $null } }
function Install-WindowsUpdates { Write-Log "Starting Windows Updates installation..."; try { return Get-WUInstall -MicrosoftUpdate -AcceptAll -IgnoreReboot } catch { Write-Log "Error installing updates: $($_.Exception.Message)" "ERROR"; return $null } }
function Test-RebootRequired { $keyPaths = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired", "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending"); if ($keyPaths | ForEach-Object { Test-Path $_ }) { return $true }; if (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name "PendingFileRenameOperations" -ErrorAction SilentlyContinue) { return $true }; return $false }
function Send-ToGoogleSheets { param([hashtable]$Data, [string]$WebhookUrl) Write-Log "Sending data to Google Sheets..."; try { $JsonData = $Data | ConvertTo-Json -Depth 10; Invoke-RestMethod -Uri $WebhookUrl -Method POST -Body $JsonData -ContentType "application/json"; Write-Log "Data sent to Google Sheets successfully"; return $true } catch { Write-Log "Error sending data to Google Sheets: $($_.Exception.Message)" "ERROR"; return $false } }

# --- PROVEN INTERACTIVE PROMPT VIA SCHEDULED TASK ---
function Get-ActiveUser {
    try {
        $activeSessionLine = qwinsta | Select-String -Pattern '^\s*([a-zA-Z0-9]+.*)\s+[0-9]+\s+Active' | Select-Object -First 1
        if ($activeSessionLine) { $userName = ($activeSessionLine.Line -split '\s+' | Where-Object { $_ })[1]; Write-Log "Active user found: $userName"; return $userName }
        Write-Log "No active user session found."; return $null
    } catch { Write-Log "Could not run 'qwinsta' to check for active sessions." "WARNING"; return $null }
}
function Show-InteractiveRebootPromptViaTask {
    param ([string]$UserName)
    $taskName = "RebootPromptTask"; $resultFile = "C:\Windows\Temp\reboot_choice.txt"; $promptScriptPath = "C:\Windows\Temp\prompt_script.ps1"
    $promptScriptBlock = { Add-Type -AssemblyName System.Windows.Forms; $promptResult = [System.Windows.Forms.MessageBox]::Show("Windows updates have been installed and a restart is required to complete the process.`n`nWould you like to restart your computer now?", "Restart Required", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question, [System.Windows.Forms.MessageBoxDefaultButton]::Button2); $promptResult | Out-File -FilePath "C:\Windows\Temp\reboot_choice.txt" -Encoding UTF8 -Force }
    $promptScriptBlock | Out-File -FilePath $promptScriptPath -Encoding UTF8 -Force
    try {
        Write-Log "Creating scheduled task '$taskName' to run as user '$UserName'..."
        $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -File `"$promptScriptPath`""
        $principal = New-ScheduledTaskPrincipal -UserId $UserName -LogonType Interactive
        Register-ScheduledTask -TaskName $taskName -Action $action -Principal $principal -Force | Out-Null
        Start-ScheduledTask -TaskName $taskName
        $timeout = (Get-Date).AddMinutes(5); $userChoice = "Timeout"
        while ((Get-Date) -lt $timeout) { if (Test-Path $resultFile) { $userChoice = Get-Content $resultFile; Write-Log "User responded with: $userChoice"; break }; Start-Sleep -Seconds 5 }
        if ($userChoice -eq "Timeout") { Write-Log "User did not respond to the prompt within the timeout period." "WARNING" }
        return $userChoice
    } catch { Write-Log "An error occurred during the scheduled task prompt process: $($_.Exception.Message)" "ERROR"; return "Error" } 
    finally { Write-Log "Cleaning up task and temporary files..."; Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue; Remove-Item $promptScriptPath, $resultFile -ErrorAction SilentlyContinue }
}

# --- MAIN SCRIPT EXECUTION ---
Write-Log "Starting Windows Update Management Script (v$ScriptVersion)"
$SystemInfo = Get-SystemInfo
$ReportData = @{ Timestamp = Get-Date -Format "o"; ComputerName = $ComputerName; SystemInfo = $SystemInfo; AvailableUpdates = 0; UpdatesInstalled = 0; RebootRequired = $false; RebootAction = "None"; PostponeReason = ""; ComplianceStatus = "In Progress"; ScriptVersion = $ScriptVersion }
$AvailableUpdates = Get-AvailableUpdates; $ReportData.AvailableUpdates = if ($AvailableUpdates) { $AvailableUpdates.Count } else { 0 }
Write-Log "Found $($ReportData.AvailableUpdates) available updates"

if ($ReportData.AvailableUpdates -gt 0) {
    $InstallResult = Install-WindowsUpdates
    if ($InstallResult) { $ReportData.UpdatesInstalled = @($InstallResult).Count; $ReportData.InstallDetails = $InstallResult | ForEach-Object { @{ Title = $_.Title; KB = $_.KB; Size = $_.Size; Status = $_.Result } } }
    if (Test-RebootRequired) {
        $ReportData.RebootRequired = $true
        $activeUserName = Get-ActiveUser
        if ($activeUserName) {
            $userChoice = Show-InteractiveRebootPromptViaTask -UserName $activeUserName
            if ($userChoice -eq "Yes") {
                $ReportData.RebootAction = "Immediate (User Approved)"; $ReportData.ComplianceStatus = "Compliant - Rebooting"
                $ReportData.ExecutionDuration = [math]::Round((New-TimeSpan -Start $ScriptStartTime -End (Get-Date)).TotalMinutes, 2)
                if ($GoogleSheetsWebhookUrl) { Send-ToGoogleSheets -Data $ReportData -WebhookUrl $GoogleSheetsWebhookUrl }
                # --- RELIABLE REBOOT FIX ---
                Write-Log "Initiating system restart in 60 seconds via shutdown.exe..."; 
                shutdown.exe /r /f /t 60 /c "System restart initiated by compliance script after user approval."
            } else {
                $ReportData.RebootAction = "Postponed (User Declined)"
                if ($userChoice -eq "Timeout") { $ReportData.PostponeReason = "User did not respond to prompt within 5-minute timeout." }
                else { $ReportData.PostponeReason = "User declined immediate reboot via prompt." }
                $ReportData.ComplianceStatus = "Non-Compliant - Reboot Pending"
            }
        } else {
            $ReportData.RebootAction = "Skipped (No User Logged In)"; $ReportData.PostponeReason = "A reboot is pending, but prompt was skipped as no active user was found."; $ReportData.ComplianceStatus = "Non-Compliant - Reboot Pending"
        }
    } else { $ReportData.ComplianceStatus = "Compliant - Patched" }
} else { $ReportData.ComplianceStatus = "Compliant - Up to Date" }

# --- FINAL REPORTING ---
if ($ReportData.RebootAction -notlike "*Rebooting*") {
    $ReportData.ExecutionDuration = [math]::Round((New-TimeSpan -Start $ScriptStartTime -End (Get-Date)).TotalMinutes, 2)
    if ($GoogleSheetsWebhookUrl) { Send-ToGoogleSheets -Data $ReportData -WebhookUrl $GoogleSheetsWebhookUrl }
}
$ReportData | ConvertTo-Json -Depth 10 | Out-File -FilePath "$env:TEMP\WindowsUpdateReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
Write-Log "Script execution completed."
if ($ReportData.ComplianceStatus -like "*Compliant*") { exit 0 } else { exit 1 }