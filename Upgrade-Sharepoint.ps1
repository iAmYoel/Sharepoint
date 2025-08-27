<#
.SYNOPSIS
    Automates SharePoint farm upgrade with content DB version logging, version verification, and email notification.

.NOTES
    Run this script as the SharePoint setup account with local admin rights.
    Requires SMTP server access for email notification.
#>

[CmdletBinding(DefaultParameterSetName = "Default")]
param (
    [Parameter(Mandatory = $true, HelpMessage = "SharePoint farm version to upgrade.", ParameterSetName = "Download")]
    [Parameter(Mandatory = $true, HelpMessage = "SharePoint farm version to upgrade.", ParameterSetName = "UpgradeOnly")]
    [Parameter(Mandatory = $true, HelpMessage = "SharePoint farm version to upgrade.", ParameterSetName = "Default")]
    [ValidateSet('SE', '2019', '2016', '2013')]
    [String]$SharePointVersion,

    [Parameter(Mandatory = $false, HelpMessage = "Automatically download patch file and install them.", ParameterSetName = "Download")]
    [Switch]$DownloadPatch,

    [Parameter(Mandatory = $false, HelpMessage = "Only run Sharepoint upgrade.", ParameterSetName = "UpgradeOnly")]
    [Switch]$UpgradeOnly
)

# === FUNCTIONS ===
function Browse-File {
    [CmdletBinding()]
    param (
        [string]$Title = "Select SharePoint patch file(s)",
        [string]$InitialDirectory = [Environment]::GetFolderPath("Desktop"),
        [string]$Filter = "Sharepoint update files (*.exe)|*.exe|All files (*.*)|*.*"
    )

    Add-Type -AssemblyName System.Windows.Forms

    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = $Title
    $dialog.InitialDirectory = $InitialDirectory
    $dialog.Filter = $Filter
    $dialog.Multiselect = $true

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    }
    else {
        return $null
    }
}

function Invoke-SPPsConfigUpgrade {

    $errorString = "Exception: The upgraded database schema doesn't match the TargetSchema"
    $psconfigArgs = "-cmd helpcollections -installall -cmd secureresources -cmd services -install -cmd installfeatures -cmd applicationcontent -install -cmd upgrade -inplace b2b -force -wait"


    $procInfo = New-Object System.Diagnostics.ProcessStartInfo
    $procInfo.FileName = "psconfig.exe"
    $procInfo.Arguments = $psconfigArgs
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

function Get-SPLatestPatch {
    param(
        [Parameter(Mandatory = $true, HelpMessage = "Select the SharePoint product to check for updates.")]
        [ValidateSet('SE','2019','2016','2013')]
        [string]$Product
    )

    # --- Script Setup ---
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    # Ensure modern TLS protocols for web requests
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor 3072
    } catch {}

    # --- Product Metadata ---
    $updatesUrl = 'https://learn.microsoft.com/en-us/officeupdates/sharepoint-updates'
    $products = @{
        'SE'   = @{ Name = 'SharePoint Server Subscription Edition'; Header = 'SharePoint Server Subscription Edition update history'; ExpectedKbCount = 1 }
        '2019' = @{ Name = 'SharePoint Server 2019';                 Header = 'SharePoint Server 2019 update history';        ExpectedKbCount = 2 }
        '2016' = @{ Name = 'SharePoint Server 2016';                 Header = 'SharePoint Server 2016 update history';        ExpectedKbCount = 2 }
        '2013' = @{ Name = 'SharePoint 2013';                        Header = 'SharePoint 2013 update history';               ExpectedKbCount = 2 }
    }

    # --- Parameter Validation ---
    if (-not $Product) {
        throw "Product parameter is required. Valid values: SE, 2019, 2016, 2013."
    }
    if (-not $products.ContainsKey($Product)) {
        throw "Unknown product selection: $Product"
    }
    $selected = $products[$Product]

    # --- Helper: Download page content ---
    function Get-ContentFromUrl {
        param([Parameter(Mandatory)][string] $Url)
        $headers = @{ 'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) PowerShell' }
        (Invoke-WebRequest -Uri $Url -Headers $headers -UseBasicParsing).Content
    }

    # --- Helper: Find section positions for each product ---
    function Get-LearnPageSections {
        param([string] $Content)
        # Dynamically find all headings containing 'SharePoint Server' or 'SharePoint'
        $headerRegex = [regex]'SharePoint (Server )?\d{4,4}.*?update history'
        $matches = $headerRegex.Matches($Content)
        $headers = @()
        $positions = @{}
        foreach ($m in $matches) {
            $header = $m.Value.Trim()
            $idx = $Content.IndexOf($header, [StringComparison]::OrdinalIgnoreCase)
            if ($idx -ge 0) {
                $headers += $header
                $positions[$header] = $idx
            }
        }
        # Always include SE and 2013 if not matched by regex
        foreach ($static in @(
            'SharePoint Server Subscription Edition update history',
            'SharePoint 2013 update history'
        )) {
            $idx = $Content.IndexOf($static, [StringComparison]::OrdinalIgnoreCase)
            if ($idx -ge 0 -and -not $positions.ContainsKey($static)) {
                $headers += $static
                $positions[$static] = $idx
            }
        }
        return @{ Content = $Content; Positions = $positions; Headers = $headers }
    }

    # --- Helper: Extract the update section for the selected product ---
    function Get-SectionWindowForProduct {
        param(
            [Parameter(Mandatory)] $SectionsMeta,
            [Parameter(Mandatory)] [string] $HeaderForProduct
        )
        $content   = $SectionsMeta.Content
        $positions = $SectionsMeta.Positions
        $headers   = $SectionsMeta.Headers

        if (-not $positions.ContainsKey($HeaderForProduct)) {
            throw "Could not find section '$HeaderForProduct' on the page. Structure may have changed."
        }

        $start = [int]$positions[$HeaderForProduct]

        # Find the next section header after the selected product
        $nextStarts = @()
        foreach ($h in $headers) {
            if ($h -ne $HeaderForProduct -and $positions.ContainsKey($h)) {
                $p = [int]$positions[$h]
                if ($p -gt $start) { $nextStarts += $p }
            }
        }
        $end = if ($nextStarts.Count -gt 0) { ($nextStarts | Measure-Object -Minimum).Minimum } else { $content.Length }

        return $content.Substring($start, $end - $start)
    }

    # --- Helper: Find KB IDs in the latest update row ---
    function Get-LatestRowKbIds {
        param(
            [Parameter(Mandatory)][string] $SectionContent,
            [Parameter(Mandatory)][string] $ProductName,
            [int] $ExpectedKbCount = 2
        )
        # Search for KB links near the product name
        $searchStart = $SectionContent.IndexOf($ProductName, [StringComparison]::OrdinalIgnoreCase)
        $window = if ($searchStart -ge 0) { $SectionContent.Substring($searchStart, [Math]::Min(6000, $SectionContent.Length - $searchStart)) } else { $SectionContent }

        $kbRegex = [regex]'https?://support\.microsoft\.com/help/(?<kb>\d{7,8})'
        $matches = $kbRegex.Matches($window)
        $kbIds = @()
        foreach ($m in $matches) {
            $kb = $m.Groups['kb'].Value
            if (-not $kbIds.Contains($kb)) { $kbIds += $kb }
            if ($kbIds.Count -ge [Math]::Max(1, $ExpectedKbCount)) { break }
        }

        # Fallback: search the whole section if not found near product name
        if ($kbIds.Count -eq 0) {
            $matches = $kbRegex.Matches($SectionContent)
            foreach ($m in $matches) {
                $kb = $m.Groups['kb'].Value
                if (-not $kbIds.Contains($kb)) { $kbIds += $kb }
                if ($kbIds.Count -ge [Math]::Max(1, $ExpectedKbCount)) { break }
            }
        }

        if ($kbIds.Count -eq 0) {
            throw 'No KB numbers found in section.'
        }

        return $kbIds
    }

    # --- Helper: Find Microsoft Download Center link, direct download link and file name ---
    function Get-DownloadLinkFromMsSupport {
        param (
            [string]$uri
        )

        try {
            # Get site info
            $response = Invoke-WebRequest -Uri $uri

            # Filter the links to find the one for the Microsoft Download Center
            $MsDcLink = ($response.Links | Where-Object { $_.href -match "microsoft.com/[A-Za-z]{2}-[A-Za-z]{2}/download/.*$" } | Select-Object -First 1).href

            if ($MsDcLink) {
                Write-Verbose "Found Download Center link: $($MsDcLink)"

                Write-Verbose "Fetching page content and links from $MsDcLink..."
                # 2. Use Invoke-WebRequest to get the webpage content and its parsed structure.
                # The 'Links' property automatically contains all <a> tags on the page.
                $response = Invoke-WebRequest -Uri $MsDcLink
                # 3. Filter the links to find the direct download URLs.
                # We are looking for links that point to 'download.microsoft.com' and end in '.exe'.
                $downloadLinks = $response.Links | Where-Object { $_.href -like '*download.microsoft.com/*.exe' }
                if ($downloadLinks) {
                    Write-Verbose "`n[SUCCESS] Found the following direct download links by scraping the page:"
                    # 4. Create a clean output table with the FileName and the full URL.
                    # The FileName is extracted from the URL itself.
                    $results = $downloadLinks | ForEach-Object {
                        [PSCustomObject]@{
                            FileName = ([System.IO.Path]::GetFileName($_.href))
                            Url      = $_.href
                        }
                    }
                } else {
                    Write-Error "Scraping failed. No direct '.exe' download links were found on the page. The page layout may have changed to hide links behind a button click."
                }
            } else {
                Write-Warning "No Microsoft Download Center link was found on the page."
            }
        }
        catch {
            Write-Error "An error occurred while trying to fetch the webpage: $($_.Exception.Message)"
        }
        return $results
    }

    # --- Helper: Extract KB metadata from update table section ---
    function Get-KbMetaInfoFromSection {
        param(
            [Parameter(Mandatory)][string] $SectionContent,
            [Parameter(Mandatory)][string] $KbId
        )

        # Find the first table row containing any of our KBs
        $rowPattern = "(?s)<tr[^>]*?>((?!</tr>).)*?KB\s*$KbId.*?</tr>"
        $rowMatch = [regex]::Match($SectionContent, $rowPattern)
        Write-Verbose "Looking for KB$KbId in section content..."

        $kbIndex = $kbIds.IndexOf($KbId)

        $version = $null
        $releaseDate = $null
        $downloadUrl = $null
        $packageName = $null

        if ($rowMatch.Success) {
            $row = $rowMatch.Value
            Write-Verbose "Found row: $row"

            # Extract all table cells
            $colPattern = "(?s)<td[^>]*>(.*?)</td>"
            $colMatches = [regex]::Matches($row, $colPattern)
            $cols = @()
            foreach ($match in $colMatches) {
                $cellContent = $match.Groups[1].Value.Trim()
                $cols += $cellContent
            }

            if ($cols.Count -ge 4) {
                # Package Name (column 0)
                $packageParts = $cols[0] -split '<br>'
                $packageName = $packageParts[$kbIndex].Trim()

                # Release Date (column 3)
                $releaseDate = $cols[3].Trim()

                # Version column (column 3)
                $versionList = $cols[2] -split '<br>'

                if ($kbIndex -lt $versionList.Count) {
                    $version = $versionList[$kbIndex].Trim()
                } elseif ($versionList.Count -gt 0) {
                    $version = $versionList[0].Trim()
                }

                # Download column (column 4) - split by <br>
                $downloadParts = $cols[1] -split '<br>'
                $downloadCell = $downloadParts[$kbIndex].Trim()
                if ($downloadCell -match "<a[^>]*href=['""]([^'""]+)['""][^>]*>") {
                        $downloadUri = $matches[1]
                    }
                } else {
                    # fallback: first link
                    foreach ($part in $downloadParts[0].Trim()) {
                        if ($part -match "<a[^>]*href=['""]([^'""]+)['""][^>]*>") {
                            $downloadUri = $matches[1]
                            break
                        }
                    }
                }

                $downloadUrl = Get-DownloadLinkFromMsSupport -uri $downloadUri
            }

        return @{
            PackageName = $packageName
            Version = $version
            ReleaseDate = $releaseDate
            DownloadUrl = $downloadUrl.Url
            FileName = $downloadUrl.FileName
        }
    }

    # --- Helper: Get KB metadata (version, release date, file names, download URLs) ---
<#     function Get-KbMetaInfo {
        param(
            [Parameter(Mandatory)][string] $KbId,
            [Parameter(Mandatory)][string] $SectionContent
        )
        $kbUrl = "https://support.microsoft.com/help/$KbId"
        Write-Host "  • Fetching KB page: $kbUrl" -ForegroundColor DarkGray

        $resp = Invoke-WebRequest -Uri $kbUrl -UseBasicParsing -Headers @{ 'User-Agent' = 'Mozilla/5.0 PowerShell' }
        $html = $resp.Content

        # Find download links and extract file names
        $linkRegex = [regex]'href\s*=\s*["''](?<u>https?://(?:download\.microsoft\.com|go\.microsoft\.com/.*?fwlink|aka\.ms/[^"'']+|www\.catalog\.update\.microsoft\.com/[^"'']+)[^"'']*)["'']'
        $matches = $linkRegex.Matches($html)
        $urls = @()
        foreach ($m in $matches) {
            $u = $m.Groups['u'].Value
            if ($u -match 'catalog\.update\.microsoft\.com') { continue }
            if (-not $urls.Contains($u)) { $urls += $u }
        }
        $resolved = @()
        foreach ($u in $urls) {
            $final = $u
            if ($u -match 'go\.microsoft\.com/.*fwlink|aka\.ms/') {
                try {
                    $finalResp = Invoke-WebRequest -Uri $u -MaximumRedirection 10 -UseBasicParsing -Headers @{ 'User-Agent' = 'Mozilla/5.0 PowerShell' }
                    if ($finalResp.BaseResponse -and $finalResp.BaseResponse.ResponseUri) {
                        $final = $finalResp.BaseResponse.ResponseUri.AbsoluteUri
                    }
                } catch {}
            }
            if (-not $resolved.Contains($final)) { $resolved += $final }
        }
        $binaryLinks = @($resolved | Where-Object { $_ -match '\.(exe|msu|cab)(\?|$)' -or $_ -match 'download\.microsoft\.com/.+' })
        $fileName = $null
        if ($binaryLinks -and $binaryLinks.Count -gt 0) {
            $fileName = [System.IO.Path]::GetFileName(($binaryLinks[0] -split '\?')[0])
        }

        # Get version and release date from section table
        $meta = Get-KbMetaInfoFromSection -SectionContent $SectionContent -KbId $KbId

        return [PSCustomObject]@{
            KbNumber    = $KbId
            Version     = $meta.Version
            ReleaseDate = $meta.ReleaseDate
            FileName    = $fileName
            DownloadUrl = if ($binaryLinks.Count -gt 0) { $binaryLinks[0] } else { $null }
        }
    } #>

    # --- Main Logic: Find and return latest KB info for selected product ---
    Write-Host "Selected product: $($selected.Name)" -ForegroundColor Cyan
    Write-Host "Fetching update list from: $updatesUrl" -ForegroundColor Cyan

    $learnContent = Get-ContentFromUrl -Url $updatesUrl
    $sectionsMeta = Get-LearnPageSections -Content $learnContent

    # Try to find the exact header, fallback to closest match
    $headerForProduct = $selected.Header
    if (-not $sectionsMeta.Positions.ContainsKey($headerForProduct)) {
        # Fallback logic for SE and others
        if ($Product -eq 'SE') {
            $fallbackHeader = $sectionsMeta.Headers | Where-Object { $_ -match 'Subscription' }
        } else {
            $year = $Product -replace '[^\d]', ''
            if ($year) {
                $fallbackHeader = $sectionsMeta.Headers | Where-Object { $_ -match $year }
            } else {
                $fallbackHeader = @()
            }
        }
        $fallbackHeader = @($fallbackHeader) # Ensure it's always an array
        if ($fallbackHeader -and $fallbackHeader.Count -gt 0) {
            $headerForProduct = $fallbackHeader[0]
            Write-Verbose "Auto-corrected header for ${Product}: $headerForProduct" #-ForegroundColor Yellow
        } else {
            throw "Could not find section for product $Product. Page structure may have changed."
        }
    }

    $sectionText = Get-SectionWindowForProduct -SectionsMeta $sectionsMeta -HeaderForProduct $headerForProduct
    $kbIds = Get-LatestRowKbIds -SectionContent $sectionText -ProductName $selected.Name -ExpectedKbCount $selected.ExpectedKbCount

    Write-Host "Latest KB row for $($selected.Name): $($kbIds | foreach {"KB$_"})" -ForegroundColor Yellow

    # --- Main Logic: Return info for all KBs found in the latest row ---
    $kbInfoList = @()
    foreach ($kb in $kbIds) {
        $meta = Get-KbMetaInfoFromSection -SectionContent $sectionText -KbId $kb
        $kbInfoList += [PSCustomObject]@{
            KbNumber    = "KB$kb"
            PackageName = $meta.PackageName
            Version     = $meta.Version
            ReleaseDate = $meta.ReleaseDate
            FileName    = $meta.FileName
            DownloadUrl = $meta.DownloadUrl
        }
    }
    return $kbInfoList
}

function Get-SPVersion {
    # Detect farm build version
    $farmBuild = (Get-SPFarm -ErrorAction Stop).BuildVersion
    $farmVersion = $farmBuild.ToString()
    $farmMajorVersion = $farmBuild.Major

    # Determine correct ISAPI path based on version
    switch ($farmMajorVersion) {
        15 { $dllPath = "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.dll" }
        16 { $dllPath = "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.dll" }
        default {
            throw "Unsupported SharePoint version: $farmMajorVersion. Please update the script for this version."
        }
    }

    $dllVersion = (Get-Item $dllPath -ErrorAction Stop).VersionInfo.ProductVersion
    $dbVersion = (Get-SPContentDatabase -ErrorAction Stop | sort BuildVersion).BuildVersion

    return @{
        InstalledVersion = $dllVersion
        FarmVersion = $farmVersion
        DatabaseVersion = $dbVersion
    }
}

function Run-SPUpgrade {
    Write-Host "Upgrading SharePoint farm to latest installed version..." -ForegroundColor Yellow
    $result = Invoke-SPPsConfigUpgrade

    if ($result.ExitCode -ne 0) {
        if ($result.SchemaMismatch) {
            Write-Warning "Detected schema mismatch. Running Upgrade-SPContentDatabase..."

            $dbsToUpgrade = Get-SPContentDatabase | Where-Object { $_.NeedsUpgrade -eq $true }
            foreach ($db in $dbsToUpgrade) {
                try {
                    Write-Host "Upgrading content database: $($db.Name)..."
                    Upgrade-SPContentDatabase -Identity $db -Confirm:$false
                } catch {
                    Write-Error "Failed while upgrading content database '$($db.Name)': $_"
                    End-Script -exitCode 1
                }
            }

            Write-Host "Retrying psconfig.exe after upgrading databases..."
            $retryResult = Invoke-SPPsConfigUpgrade

            if ($retryResult.ExitCode -ne 0) {
                Write-Error "psconfig.exe failed again with exit code $($retryResult.ExitCode)"
                End-Script -exitCode 1
            }
            else {
                Write-Host "psconfig.exe completed successfully after retry."
            }
        }

        else {
            Write-Error "psconfig.exe failed with unexpected error. Exit code: $($result.ExitCode)"
            End-Script -exitCode 1
        }
    }
    else {
        Write-Host "psconfig.exe completed successfully."
    }

    try {
            Write-Host "Step 4: Logging post-upgrade content database versions..."
            Get-SPContentDatabase -ErrorAction Stop | Select Name, Id, Server, NeedsUpgrade, Version | Format-List
        }
        catch {
            Write-Warning "Failed to log post-upgrade content DB info: $_"
        }

        try {
            $currentVersion = Get-SPVersion

            Write-Host ""
            Write-Host "Final version validation:"
            Write-Host "---------------------------------------------"
            Write-Host "DLL Version (Microsoft.SharePoint.dll): $($currentVersion.InstalledVersion)"
            Write-Host "Farm Build Version (Get-SPFarm):        $($currentVersion.FarmVersion)"
            Write-Host "Content DB Build Version:               $(($currentVersion.DatabaseVersion -join ', ').ToString())"
            Write-Host "---------------------------------------------"

            if ($currentVersion.InstalledVersion -eq $currentVersion.FarmVersion -and $currentVersion.FarmVersion -eq ($currentVersion.DatabaseVersion)[0].ToString() -and $currentDatabaseInfo.NeedsUpgrade -notcontains $true) {
                Write-Host "SharePoint farm successfully upgraded to version '$($currentVersion.InstalledVersion)'." -ForegroundColor Green
            } else {
                Write-Warning "Version mismatch detected. "
                if ($currentDatabaseInfo.NeedsUpgrade -contains $true) {
                    Write-Warning "Follwing content databases still need upgrade: $($currentDatabaseInfo | Where-Object { $_.NeedsUpgrade -eq $true } | Select-Object -ExpandProperty Name)"
                }
            }
        }
        catch {
            Write-Warning "Failed to retrieve version information: $_"
        }
}

Function Install-SPPatch {
    <#
    .SYNOPSIS
        Install-SPPatch
    .DESCRIPTION
        Install-SPPatch reduces the amount of time it takes to install SharePoint patches. This cmdlet supports SharePoint 2013 and above. Additional information
        can be found at https://github.com/Nauplius.
    .PARAMETER Path
        The folder where the patch file(s) reside.
    .PARAMETER Pause
        Pauses the Search Service Application(s) prior to stopping the SharePoint Search Services.
    .PARAMETER Stop
        Stop the SharePoint Search Services without pausing the Search Service Application(s).
    .PARAMETER SilentInstall
        Silently installs the patches without user input. Not specifying this parameter will cause each patch to prompt to install.
    .PARAMETER KeepSearchPaused
        Keeps the Search Service Application(s) in a paused state after the installation of the patch has completed. Useful for when applying the patch to multiple
        servers in the farm. Default to false.
    .PARAMETER OnlySTS
        Only apply the STS (non-language dependent) patch. This switch may be used when only an STS patch is available.
    .EXAMPLE
        Install-SPPatch -Path C:\Updates -Pause -SilentInstall

        Install the available patches in C:\Updates, pauses the Search Service Application(s) on the farm, and performs a silent installation.
    .EXAMPLE
        Install-SPPatch -Path C:\Updates -Pause -KeepSearchPaused:$true -SilentInstall

        Install the available patches in C:\Updates, pauses the Search Service Application(s) on the farm,
        does not resume the Search Service Application(s) after the installation is complete, and performs a silent installation.
    .NOTES
        Author: Trevor Seward
        Date: 01/16/2020
    .LINK
        https://thesharepointfarm.com
    .LINK
        https://github.com/Nauplius
    .LINK
        https://sharepointupdates.com
#>
    param
    (
        [string]
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        $Path,
        [switch]
        [Parameter(Mandatory = $true, ParameterSetName = "PauseSearch")]
        $Pause,
        [switch]
        [Parameter(Mandatory = $true, ParameterSetName = "StopSearch")]
        $Stop,
        [switch]
        [Parameter(Mandatory = $false, ParameterSetName = "PauseSearch")]
        $KeepSearchPaused = $false,
        [switch]
        [Parameter(Mandatory = $false)]
        $SilentInstall,
        [switch]
        [Parameter(Mandatory = $false)]
        $OnlySTS
    )

    $version = (Get-SPFarm).BuildVersion
    $majorVersion = $version.Major
    $startTime = Get-Date
    $exitRebootCodes = @(3010, 17022)
    $searchSvcRunning = $false

    #Write-Host -ForegroundColor Green "Current build: $version"

    ###########################
    ##Ensure Patch is Present##
    ###########################

    if ($majorVersion -eq '16') {
        $sts = Get-ChildItem -LiteralPath $Path  -Filter *.exe | ? { $_.Name -match 'sts([A-Za-z0-9\-]+).exe' }
        $wssloc = Get-ChildItem -LiteralPath $Path  -Filter *.exe | ? { $_.Name -match 'wssloc([A-Za-z0-9\-]+).exe' }

        if ($OnlySTS) {
            if ($sts -eq $null) {
                Write-Host 'Missing the sts patch. Please make sure the sts patch present in the specified directory.' -ForegroundColor Red
                return
            }
        }
        else {
            if ($sts -eq $null -and $wssloc -eq $null) {
                Write-Host 'Missing the sts and wssloc patch. Please make sure both patches are present in the specified directory.' -ForegroundColor Red
                return
            }

            if ($sts -eq $null -or $wssloc -eq $null) {
                Write-Host '[Warning] Either the sts and wssloc patch is not available. Please make sure both patches are present in the same directory or safely ignore if only single patch is available.' -ForegroundColor Yellow
                return
            }
        }

        if ($OnlySTS) {
            $patchfiles = $sts
            Write-Host -for Yellow "Installing $sts"
        }
        else {
            $patchfiles = $sts, $wssloc
            Write-Host -for Yellow "Installing $sts and $wssloc"
        }
    }
    elseif ($majorVersion -eq '15') {
        $patchfiles = Get-ChildItem -LiteralPath $Path  -Filter *.exe | ? { $_.Name -match '([A-Za-z0-9\-]+)2013-kb([A-Za-z0-9\-]+)glb.exe' }

        if ($patchfiles -eq $null) {
            Write-Host 'Unable to retrieve the file(s). Exiting Script' -ForegroundColor Red
            return
        }

        Write-Host -ForegroundColor Yellow "Installing $patchfiles"
    }
    elseif ($majorVersion -lt '15') {
        throw 'This cmdlet supports SharePoint 2013 and above.'
    }

    ########################
    ##Stop Search Services##
    ########################
    ##Checking Search services##

    $oSearchSvc = Get-Service "OSearch$majorVersion"
    $sPSearchHCSvc = Get-Service "SPSearchHostController"

    if (($oSearchSvc.status -eq 'Running') -or ($sPSearchHCSvc.status -eq 'Running')) {
        $searchSvcRunning = $true
        if ($Pause) {
            $ssas = Get-SPEnterpriseSearchServiceApplication

            foreach ($ssa in $ssas) {
                Write-Host -ForegroundColor Yellow "Pausing the Search Service Application: $($ssa.DisplayName)"
                Write-Host  -ForegroundColor Yellow  ' This could take a few minutes...'
                Suspend-SPEnterpriseSearchServiceApplication -Identity $ssa | Out-Null
            }
        }
        elseif ($Stop) {
            Write-Host -ForegroundColor Cyan ' Continuing without pausing the Search Service Application'
        }
    }

    #We don't need to stop SharePoint Services for 2016 and above
    if ($majorVersion -lt '16') {
        Write-Host -ForegroundColor Yellow 'Stopping Search Services if they are running'

        if ($oSearchSvc.status -eq 'Running') {
            Set-Service -Name "OSearch$majorVersion" -StartupType Disabled
            Stop-Service "OSearch$majorVersion" -WA 0
        }

        if ($sPSearchHCSvc.status -eq 'Running') {
            Set-Service 'SPSearchHostController' -StartupType Disabled
            Stop-Service 'SPSearchHostController' -WA 0
        }

        Write-Host -ForegroundColor Green 'Search Services are Stopped'
        Write-Host

        #######################
        ##Stop Other Services##
        #######################
        Set-Service -Name 'IISADMIN' -StartupType Disabled
        Set-Service -Name 'SPTimerV4' -StartupType Disabled

        Write-Host -ForegroundColor Green 'Gracefully stopping IIS...'
        Write-Host
        iisreset -stop -noforce
        Write-Host -ForegroundColor Yellow 'Stopping SPTimerV4'
        Write-Host

        $sPTimer = Get-Service 'SPTimerV4'
        if ($sPTimer.Status -eq 'Running') {
            Stop-Service 'SPTimerV4'
        }

        Write-Host -ForegroundColor Green 'Services are Stopped'
        Write-Host
        Write-Host
    }

    ##################
    ##Start patching##
    ##################
    Write-Host -ForegroundColor Yellow 'Working on it... Please keep this PowerShell window open...'
    Write-Host

    $patchStartTime = Get-Date

    foreach ($patchfile in $patchfiles) {
        $filename = $patchfile.Fullname
        #unblock the file, to get rid of the prompts
        Unblock-File -Path $filename -Confirm:$false

        if ($SilentInstall) {
            $process = Start-Process $filename -ArgumentList '/passive /quiet' -PassThru -Wait
        }
        else {
            $process = Start-Process $filename -ArgumentList '/norestart' -PassThru -Wait
        }

        if ($exitRebootCodes.Contains($process.ExitCode)) {
            $reboot = $true
        }

        Write-Host -ForegroundColor Yellow "Patch $patchfile installed with Exit Code $($process.ExitCode)"
    }

    $patchEndTime = Get-Date

    Write-Host
    Write-Host -ForegroundColor Yellow ('Patch installation completed in {0:g}' -f ($patchEndTime - $patchStartTime))
    Write-Host

    if ($majorVersion -lt '16') {
        ##################
        ##Start Services##
        ##################
        Write-Host -ForegroundColor Yellow 'Starting Services'
        Set-Service -Name 'SPTimerV4' -StartupType Automatic
        Set-Service -Name 'IISADMIN' -StartupType Automatic

        Start-Service 'SPTimerV4'
        Start-Service 'IISAdmin'

        ###Ensuring Search Services were stopped by script before Starting"
        if ($searchSvcRunning = $true) {
            Set-Service -Name "OSearch$majorVersion" -StartupType Manual
            Start-Service "OSearch$majorVersion" -WA 0
            Set-Service 'SPSearchHostController' -StartupType Automatic
            Start-Service 'SPSearchHostController' -WA 0
        }
    }

    ###Resuming Search Service Application if paused###
    if ($Pause -and $KeepSearchPaused -eq $false) {
        $ssas = Get-SPEnterpriseSearchServiceApplication

        foreach ($ssa in $ssas) {
            Write-Host -ForegroundColor Yellow "Resuming the Search Service Application: $($ssa.DisplayName)"
            Write-Host -ForegroundColor Yellow ' This could take a few minutes...'
            Resume-SPEnterpriseSearchServiceApplication -Identity $ssa | Out-Null
        }
    }
    elseif ($pause -and $KeepSearchPaused -eq $true) {
        Write-Host -ForegroundColor Yellow 'Not resuming the Search Service Application(s)'
    }

    ###Resuming IIS###
    iisreset -start

    $endTime = Get-Date
    Write-Host -ForegroundColor Green 'Services are Started'

    if ($reboot) {
        Write-Host -ForegroundColor Yellow 'A reboot is required'
    }

    return @(
        @{
            ExitCode = $process.ExitCode
            Duration = $endTime - $startTime
            RebootRequired = $reboot
        }
    )
}

function End-Script {
    param(
        [int]$exitCode = 0
    )
    $endTime = Get-Date
    Write-Host -ForegroundColor Yellow ('Script completed in {0:g}' -f ($endTime - $startTime))
    Write-Host -ForegroundColor Yellow 'Started:'  $startTime
    Write-Host -ForegroundColor Yellow 'Finished:'  $endTime

    Stop-Transcript
    if (-not $DownloadPatch) {Pause}
    Exit $exitCode
}

# === CONFIGURATION ===
$startTime = Get-Date
$scriptPath = $MyInvocation.MyCommand.Definition
$scriptFolder = Split-Path -Parent $scriptPath
$logPath = Join-Path -Path $scriptFolder -ChildPath "SPLogs"
$patchPath = $logPath = Join-Path -Path $scriptFolder -ChildPath "SPUpdateFiles"
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $logPath "SharePointUpgrade_$timestamp.log"

$taskName = "SharePointUpgrade-AfterRebootTask"
$taskDescription = "Automatically created by Upgrade-Sharepoint script to continue upgrade after install and reboot of patch."
$taskTrigger = New-ScheduledTaskTrigger -AtLogOn -User "$env:USERDOMAIN\$env:USERNAME"
$taskAction = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -DownloadPatch -UpgradeOnly"
$taskSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -DisallowDemandStart

<#
# Email settings (configure these)
$emailFrom = "sharepoint-upgrade@yourdomain.com"
$emailTo = "sp-admins@yourdomain.com"
$smtpServer = "smtp.yourdomain.com"
$emailSubject = "SharePoint Farm Upgrade Log - $timestamp"
 #>

# === SETUP ===
if (!(Test-Path $logPath)) {
    New-Item -Path $logPath -ItemType Directory | Out-Null
}

if (!(Test-Path $patchPath)) {
    New-Item -Path $patchPath -ItemType Directory | Out-Null
}

Start-Transcript -Path $logFile -Append
if ($UpgradeOnly.IsPresent) {
    Write-Host "Running SharePoint upgrade only..." -ForegroundColor Yellow
} elseif ($DownloadPatch.IsPresent) {
    Write-Host "Running SharePoint farm upgrade with automatic patch download/install..." -ForegroundColor Yellow
}

# Remove old scheduled task if it exists
if ($oldTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
    Write-Host "Removing old scheduled task: $taskName" -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
}

# Load SharePoint PowerShell snap-in
try {
    Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop
} catch {
    Write-Error "Could not load SharePoint PowerShell snap-in: $_"
    End-Script -exitCode 1
}

# Log current Content DB versions
try {
    Write-Host "Logging pre-upgrade content database versions..." -ForegroundColor Yellow
    $currentDatabaseInfo = Get-SPContentDatabase -ErrorAction Stop | Select Name, Id, Server, NeedsUpgrade, Version
    $currentDatabaseInfo | Format-List
}
catch {
    Write-Warning "Failed to log pre-upgrade content DB info: $_"
}

# === STEP 1: Compare latest version online with installed current version ===
Write-Host "`n`tStep 1: Checking current installed Sharepoint Version..." -ForegroundColor Yellow
$currentVersion = Get-SPVersion

if ($currentVersion.InstalledVersion -eq $currentVersion.FarmVersion -and $currentVersion.FarmVersion -eq ($currentVersion.DatabaseVersion)[0].ToString() -and $currentDatabaseInfo.NeedsUpgrade -notcontains $true) {
    Write-Host "Checking latest Sharepoint version online..."
    $latestVersion = Get-SPLatestPatch -Product '2016'

    Write-Host "Installed version: $($currentVersion.InstalledVersion)" -ForegroundColor Cyan
    Write-Host "Latest release version: $($latestVersion.Version[0])" -ForegroundColor Cyan

    if ($currentVersion.FarmVersion -like $($latestVersion.Version[0])) {
        Write-Host "Latest Sharepoint version is already installed. No upgrade needed." -ForegroundColor Green
        End-Script
    } else {
        Write-Host "New Sharepoint version available: $($latestVersion.Version[0]). Proceeding with upgrade..." -ForegroundColor Yellow
        Write-Host "Latest KBs: $($latestVersion.KbNumber -join ', ')" -ForegroundColor Yellow
    }
} else {
    if (-not $UpgradeOnly.IsPresent) {
        Write-Warning "Version mismatch detected!"
        Write-Host "Current SharePoint version is not up-to-date or requires upgrade." -ForegroundColor Yellow
    }

    if ($currentDatabaseInfo.NeedsUpgrade -contains $true) {
        Write-Host "Following content databases needs upgrade: $($currentDatabaseInfo | Where-Object { $_.NeedsUpgrade -eq $true } | Select-Object -ExpandProperty Name -Join ', ')" -ForegroundColor Yellow

        if (-not $UpgradeOnly.IsPresent) {
            do {
                $Prompt1 = Read-Host "Do you want the script to try upgrading the database(s)? (Y/N)" -ForegroundColor Cyan
            } Until ($Prompt1 -eq "Y" -or $Prompt1 -eq "N")

            if ($Prompt1 -eq "N") {
                Write-Host "Please manually check that Sharepoint has previously been upgraded properly." -ForegroundColor Yellow
                Write-Host "Current Installed Version: $($currentVersion.installedVersion)" -ForegroundColor Yellow
                Write-Host "Current Farm Version: $($currentVersion.FarmVersion)" -ForegroundColor Yellow
                Write-Host "Current Database(s) Version: $($currentDatabaseInfo.Version -join ", ")" -ForegroundColor Yellow
                End-Script
            }
        }
    }
}

# === STEP 2: Download or browse patch file ===
if (-not $UpgradeOnly.IsPresent -and $Prompt1 -ne "Y") {
    Write-Host "`n`tStep 2: Acquiring patch file..." -ForegroundColor Yellow

    if (-not $DownloadPatch.IsPresent) {
        do {
            $Prompt2 = Read-Host "Automatically download and install latest patch file(s)? (Y/N)" -ForegroundColor Cyan
        } Until ($Prompt2 -eq "Y" -or $Prompt2 -eq "N")
    }

    if ($DownloadPatch.IsPresent -or $Prompt2 -eq 'Y') {
        try {
            if ($latestVersion.Count -gt 0) {
                $exeFiles = @()
                foreach ($patch in $latestVersion) {
                    $patchDownloadUrl = $patch.DownloadUrl
                    $patchFileName = $patch.FileName

                    Write-Host "Downloading patch file: $patchFileName..." -ForegroundColor Cyan
                    $localFilePath = Join-Path -Path $patchPath -ChildPath $patchFileName

                    $ProgressPreference = 'SilentlyContinue'
                    Invoke-WebRequest -Uri $patchDownloadUrl -OutFile $localFilePath -ErrorAction Stop
                    $ProgressPreference = 'Continue'
                    Write-Host "Successfully downloaded to: $localFilePath" -ForegroundColor Green

                    $exeFiles += $localFilePath
                }
            } else {
                Write-Warning "No patches found for SharePoint 2016."
                End-Script
            }
        } catch {
            Write-Error "Failed to download patch: $_"
            End-Script -exitCode 1
        }
    } else {
        Write-Host "Skipping automatic download. Please choose patch file(s)..." -ForegroundColor Yellow
        $exeFiles = Browse-File
    }
} else {
    Write-Host "Skipping patch download. Continuing with next step..." -ForegroundColor Yellow
}

# === STEP 3: Install patch file ===
if ($Prompt2 -eq "Y" -or $DownloadPatch.IsPresent -and $Prompt1 -ne "Y") {
    Write-Host "`n`tStep 3: Install patch file..." -ForegroundColor Yellow
    foreach ($file in $exeFiles) {
        if (-not (Test-Path $file)) {
            Write-Error "Patch file not found: $file"
            End-Script -exitCode 1
        }

        try {
            $returnExitCodes = @(0, 3010, 17022) # 0 = success, 3010 = reboot required, 17022 = patch already installed
            Write-Host "Installing patch $(Split-Path -Path $file -Leaf)..." -ForegroundColor Yellow
            $exeInstall = Install-SPPatch -Path (Split-Path -Path $file -Parent) -Pause -SilentInstall -ErrorAction Stop
            Write-Host "Patch installed successfully." -ForegroundColor Green

            if ($exeInstall.RebootRequired) {
                $triggerReboot = $true
                Write-Host "A reboot is required after patch installation." -ForegroundColor Yellow
            }

            if ($exeInstall.ExitCode -in $returnExitCodes -and $localFilePath) {
                Remove-Item -Path $file -Force -Confirm:$false -ErrorAction SilentlyContinue
            }
        }
        catch {
            Write-Error "Failed to install patch: $_"
            End-Script -exitCode 1
        }
    }

    $afterInstallVersion = Get-SPVersion
    Write-Host "Installed SharePoint version: $($afterInstallVersion.InstalledVersion)" -ForegroundColor Cyan
    Write-Host "Farm version: $($afterInstallVersion.FarmVersion)" -ForegroundColor Cyan
    Write-Host "Content DB version: $(($afterInstallVersion.DatabaseVersion -join ", ").ToString())" -ForegroundColor Cyan

} else {
    Write-Host "Skipping patch installation. Continuing with next step..." -ForegroundColor Yellow
}

if ($triggerReboot) {
    Register-ScheduledTask -TaskName $taskName -Description $taskDescription -Trigger $taskTrigger -Action $taskAction -Settings $taskSettings -User $(whoami) -RunLevel Highest
    Write-Host "A reboot is required to complete the installation. Please restart the server and this script will automatically resume the upgrade next time this account '$env:USERDOMAIN\$env:USERNAME' logs in." -ForegroundColor Yellow

    if (-not $DownloadPatch) {
        do {
            $Prompt3 = Read-Host "Do you want to reboot now? (Y/N)" -ForegroundColor Cyan
        } Until ($Prompt3 -eq "Y" -or $Prompt3 -eq "N")

        if ($Prompt3 -eq "Y") {
            Restart-Computer
        } else {
            Write-Host "Please remember to reboot the server later." -ForegroundColor Yellow
        }
    } else {
        Restart-Computer
    }
}


# === STEP 4: Run psconfig.exe and upgrade content database if needed ===
Write-Host "`n`tStep 4: Upgrading Sharepoint Farm..." -ForegroundColor Yellow
Run-SPUpgrade

Write-Host "SharePoint upgrade completed!" -ForegroundColor Yellow
if ($prompt -eq "Y") {
    Write-Host "Please re-run the script to check for new updates again." -ForegroundColor Yellow
} else {
    Write-Host "Please manually verify that all content databases are up to date and that the farm is functioning correctly." -ForegroundColor Yellow
}
# === STEP 4: Upgrade content databases if needed ===
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
<# try {
    Write-Host "Step 4: Logging post-upgrade content database versions..."
    Get-SPContentDatabase -ErrorAction Stop | Select Name, Id, Server, NeedsUpgrade, Version | Format-List
}
catch {
    Write-Warning "Failed to log post-upgrade content DB info: $_"
}
 #>
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

# === STEP 5: Final version validation ===
<# try {
    $currentVersion = Get-SPVersion

    Write-Host ""
    Write-Host "Final Version Validation:"
    Write-Host "---------------------------------------------"
    Write-Host "DLL Version (Microsoft.SharePoint.dll): $($currentVersion.InstalledVersion)"
    Write-Host "Farm Build Version (Get-SPFarm):        $($currentVersion.FarmVersion)"
    Write-Host "Content DB Build Version:               $($currentVersion.DatabaseVersion)"
    Write-Host "---------------------------------------------"

    if ($currentVersion.InstalledVersion -eq $currentVersion.FarmVersion -and $currentVersion.FarmVersion -eq $currentVersion.DatabaseVersion) {
        Write-Host "SharePoint farm successfully upgraded to version $farmVersion."
    } else {
        Write-Warning "Version mismatch detected. "
    }
    Write-Host "Please verify that all content databases are up to date and that the farm is functioning correctly manually." -ForegroundColor Yellow
}
catch {
    Write-Warning "Failed to retrieve version information: $_"
} #>

Stop-Transcript
End-Script