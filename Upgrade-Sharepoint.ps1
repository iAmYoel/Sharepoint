<#
.SYNOPSIS
    Automates SharePoint farm upgrade with content DB version logging, version verification, and email notification.

.NOTES
    Run this script as the SharePoint setup account with local admin rights.
    Requires SMTP server access for email notification.
#>

# === FUNCTIONS ===
function Invoke-PsconfigWithRealTimeOutput {
    param (
        [string]$Arguments
    )

    $errorString = "Exception: The upgraded database schema doesn't match the TargetSchema"

    $procInfo = New-Object System.Diagnostics.ProcessStartInfo
    $procInfo.FileName = "psconfig.exe"
    $procInfo.Arguments = $Arguments
    $procInfo.UseShellExecute = $false
    $procInfo.RedirectStandardOutput = $true
    $procInfo.RedirectStandardError = $true
    $procInfo.CreateNoWindow = $true

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $procInfo

    $schemaMismatchDetected = $false

    $process.Start() | Out-Null

    while (-not $process.StandardOutput.EndOfStream) {
        $line = $process.StandardOutput.ReadLine()
        Write-Host $line

        if ($line -match $errorString) {
            $schemaMismatchDetected = $true
        }
    }

    while (-not $process.StandardError.EndOfStream) {
        $line = $process.StandardError.ReadLine()
        Write-Host $line

        if ($line -match $errorString) {
            $schemaMismatchDetected = $true
        }
    }

    $process.WaitForExit()
    return @{ ExitCode = $process.ExitCode; SchemaMismatch = $schemaMismatchDetected }
}


# === CONFIGURATION ===
$logPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$logPath = "C:\SPUpgradeLogs"
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $logPath "SharePointUpgrade_$timestamp.log"

# Email settings (configure these)
$emailFrom = "sharepoint-upgrade@yourdomain.com"
$emailTo = "sp-admins@yourdomain.com"
$smtpServer = "smtp.yourdomain.com"
$emailSubject = "SharePoint Farm Upgrade Log - $timestamp"

$psconfigArgs = "-cmd helpcollections -installall -cmd secureresources -cmd services -install -cmd installfeatures -cmd applicationcontent -install -cmd upgrade -inplace b2b -force -wait"

# === SETUP ===
<# if (!(Test-Path $logPath)) {
    New-Item -Path $logPath -ItemType Directory | Out-Null
} #>

Start-Transcript -Path $logFile -Append
Write-Host "Starting SharePoint Upgrade Process`n"

# Load SharePoint PowerShell snap-in
try {
    Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop
} catch {
    Write-Error "Could not load SharePoint PowerShell snap-in: $_"
    Stop-Transcript
    exit 1
}

# === STEP 0: Log current Content DB versions ===
try {
    Write-Host "Step 0: Logging pre-upgrade content database versions..."
    Get-SPContentDatabase -ErrorAction Stop | Select Name, Id, Server, NeedsUpgrade, Version | Format-List
}
catch {
    Write-Warning "Failed to log pre-upgrade content DB info: $_"
}

# === STEP 1: Run psconfig.exe and upgrade content database if needed ===

$result = Invoke-PsconfigWithRealTimeOutput -Arguments $psconfigArgs

if ($result.ExitCode -ne 0) {
    if ($result.SchemaMismatch) {
        Write-Warning "Detected schema mismatch. Running Upgrade-SPContentDatabase..."

        $dbsToUpgrade = Get-SPContentDatabase | Where-Object { $_.NeedsUpgrade -eq $true }
        foreach ($db in $dbsToUpgrade) {
            try {
                Write-Host "Upgrading content database: $($db.Name)"
                Upgrade-SPContentDatabase -Identity $db -Confirm:$false
            } catch {
                Write-Error "Failed while upgrading content database '$($db.Name)': $_"
                Stop-Transcript
                exit 1
            }
        }

        Write-Host "Retrying psconfig.exe after upgrading databases..."
        $retryResult = Invoke-PsconfigWithRealTimeOutput -Arguments $psconfigArgs

        if ($retryResult.ExitCode -ne 0) {
            Write-Error "psconfig.exe failed again with exit code $($retryResult.ExitCode)"
            Stop-Transcript
            exit 1
        }
        else {
            Write-Host "psconfig.exe completed successfully after retry."
        }
    }
    else {
        Write-Error "psconfig.exe failed with unexpected error. Exit code: $($result.ExitCode)"
        Stop-Transcript
        exit 1
    }
}
else {
    Write-Host "psconfig.exe completed successfully."
}


# === STEP 2: Upgrade content databases if needed ===
<# try {
    Write-Host "Step 2: Checking for content databases that need upgrade..."
    $databasesNeedingUpgrade = Get-SPContentDatabase | Where-Object { $_.NeedsUpgrade -eq $true }

    if ($databasesNeedingUpgrade.Count -gt 0) {
        Write-Host "Found $($databasesNeedingUpgrade.Count) content database(s) requiring upgrade."

        foreach ($db in $databasesNeedingUpgrade) {
            Write-Host "Upgrading content DB: $($db.Name)"
            Upgrade-SPContentDatabase -Identity $db -Confirm:$false
        }

        Write-Host "Content databases upgraded."
    } else {
        Write-Host "All content databases are already up to date."
    }
}
catch {
    Write-Error "Failed while upgrading content databases: $_"
    Stop-Transcript
    exit 1
} #>


# === STEP 3: Log post-upgrade DB state ===
try {
    Write-Host "Step 3: Logging post-upgrade content database versions..."
    Get-SPContentDatabase -ErrorAction Stop | Select Name, Id, Server, NeedsUpgrade, Version | Format-List
}
catch {
    Write-Warning "Failed to log post-upgrade content DB info: $_"
}

<# # === STEP 4: Check for pending reboot ===
try {
    $rebootKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
    if (Test-Path $rebootKey) {
        Write-Warning "A system reboot is pending. Consider restarting this server."
    } else {
        Write-Host "No pending reboot detected."
    }
}
catch {
    Write-Warning "Could not determine reboot status."
} #>

<# # === STEP 5: Email the upgrade log ===
try {
    Write-Host "Sending upgrade log via email..."
    Send-MailMessage -From $emailFrom -To $emailTo -Subject $emailSubject `
        -Body "SharePoint farm upgrade completed. See attached log." `
        -Attachments $logFile, $preUpgradeDbLog, $postUpgradeDbLog `
        -SmtpServer $smtpServer

    Write-Host "Email sent successfully to $emailTo"
}
catch {
    Write-Warning "Failed to send email log: $_"
} #>

# === STEP 6: Final version validation ===
try {
    # Detect farm build version
    $farmBuild = (Get-SPFarm -ErrorAction Stop).BuildVersion
    $farmVersion = $farmBuild.ToString()
    $majorVersion = $farmBuild.Major

    # Determine correct ISAPI path based on version
    switch ($majorVersion) {
        15 { $dllPath = "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.dll" }
        16 { $dllPath = "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.dll" }
        default {
            throw "Unsupported SharePoint version: $majorVersion. Please update the script for this version."
        }
    }

    $dllVersion = (Get-Item $dllPath -ErrorAction Stop).VersionInfo.ProductVersion
    $dbVersion = (Get-SPContentDatabase -ErrorAction Stop | Select-Object -First 1).BuildVersion.ToString()

    Write-Host ""
    Write-Host "Final Version Validation:"
    Write-Host "---------------------------------------------"
    Write-Host "DLL Version (Microsoft.SharePoint.dll): $dllVersion"
    Write-Host "Farm Build Version (Get-SPFarm):        $farmVersion"
    Write-Host "Content DB Build Version:               $dbVersion"
    Write-Host "---------------------------------------------"

    if ($dllVersion -eq $farmVersion -and $farmVersion -eq $dbVersion) {
        Write-Host "SharePoint farm successfully upgraded to version $farmVersion."
    } else {
        Write-Warning "Version mismatch detected. Verify upgrade status manually."
    }
}
catch {
    Write-Warning "Failed to retrieve version information: $_"
}

Stop-Transcript
