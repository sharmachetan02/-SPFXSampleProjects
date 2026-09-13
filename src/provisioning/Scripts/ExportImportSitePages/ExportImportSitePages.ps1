param(
    [Parameter(Mandatory = $true)]
    [string]$SourceSiteUrl,

    [Parameter(Mandatory = $true)]
    [string]$TargetSiteUrl,

    [Parameter(Mandatory = $true)]
    [string]$LogFilePath,

    [Parameter(Mandatory = $false)]
    [string]$PageNamesCsv,

    [Parameter(Mandatory = $false)]
    [int]$BatchSize = 10,

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [switch]$DryRun
)

# ================================
# Configuration - Site Pages Library Settings
# ================================
$SITE_PAGES_CONFIG = @{
    SOURCE_LIST_TITLE   = "Site Pages"
    TARGET_LIST_TITLE   = "Site Pages"
    SOURCE_FOLDER_NAME  = "SitePages"
    TARGET_FOLDER_NAME  = "SitePages"
    PAGE_EXTENSION      = ".aspx"
}

# ================================
# Configuration - Metadata Fields to Copy
# ================================
$METADATA_FIELDS = @{
    TITLE       = "Title"
    DESCRIPTION = "Description"
}

# ================================
# Configuration - Runtime Behavior
# ================================
$RUNTIME = @{
    RETRY_MAX               = 3
    RETRY_DELAY_SECONDS     = 2
    TEMP_EXPORT_FOLDER_NAME = "Temp_SitePagesExport"
    PUBLISH_AFTER_IMPORT    = $true
}

# Initialize script variables
$script:LogInitialized = $false
$script:CopiedCount = 0
$script:SkippedCount = 0
$script:ErrorCount = 0
$script:PageErrors = @()
$script:PageSkips = @()

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )

    if (-not $script:LogInitialized) {
        $script:LogInitialized = $true

        $logDir = Split-Path $LogFilePath
        if (-not (Test-Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        }

        if (Test-Path $LogFilePath) {
            try {
                Clear-Content -Path $LogFilePath -ErrorAction Stop
            }
            catch {
                Set-Content -Path $LogFilePath -Value $null -Force
            }
        }
        else {
            New-Item -Path $LogFilePath -ItemType File -Force | Out-Null
        }
    }

    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $logEntry = "$timestamp [$Level] $Message"

    Add-Content -Path $LogFilePath -Value $logEntry
    Write-Host $logEntry
}

function Add-PageError {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PageName,

        [Parameter(Mandatory = $true)]
        [string]$Reason
    )

    $script:ErrorCount++
    $errorEntry = [PSCustomObject]@{
        PageName = $PageName
        Reason   = $Reason
    }
    $script:PageErrors += $errorEntry
    Write-Log "Page '$PageName': $Reason" "ERROR"
}

function Add-PageSkip {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PageName,

        [Parameter(Mandatory = $true)]
        [string]$Reason
    )

    $script:SkippedCount++
    $skipEntry = [PSCustomObject]@{
        PageName = $PageName
        Reason   = $Reason
    }
    $script:PageSkips += $skipEntry
    Write-Log "Page '$PageName': $Reason" "WARN"
}

function Get-NormalizedPageName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PageName
    )

    $normalized = $PageName.Trim()
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return $null
    }

    if (-not $normalized.ToLowerInvariant().EndsWith($SITE_PAGES_CONFIG.PAGE_EXTENSION)) {
        return "$normalized$($SITE_PAGES_CONFIG.PAGE_EXTENSION)"
    }

    return $normalized
}

function Get-RequestedPagesSet {
    param(
        [Parameter(Mandatory = $false)]
        [string]$NamesCsv
    )

    if ([string]::IsNullOrWhiteSpace($NamesCsv)) {
        return $null
    }

    $set = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $parts = $NamesCsv -split ","

    foreach ($part in $parts) {
        $normalized = Get-NormalizedPageName -PageName $part
        if (-not [string]::IsNullOrWhiteSpace($normalized)) {
            [void]$set.Add($normalized)
        }
    }

    if ($set.Count -eq 0) {
        return $null
    }

    return $set
}

function Invoke-WithRetry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ActionName,

        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock
    )

    $attempt = 0
    while ($attempt -lt $RUNTIME.RETRY_MAX) {
        try {
            return & $ScriptBlock
        }
        catch {
            $attempt++
            if ($attempt -lt $RUNTIME.RETRY_MAX) {
                Write-Log "$ActionName failed (attempt $attempt). Retrying..." "WARN"
                Start-Sleep -Seconds $RUNTIME.RETRY_DELAY_SECONDS
            }
            else {
                throw "$ActionName failed after $($RUNTIME.RETRY_MAX) attempts: $($_.Exception.Message)"
            }
        }
    }
}

function Connect-SharePointSite {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,

        [Parameter(Mandatory = $true)]
        [string]$Label,

        [Parameter(Mandatory = $false)]
        [string]$EffectiveClientId
    )

    Write-Log "Connecting to $Label site: $SiteUrl"

    $connectCommand = Get-Command Connect-PnPOnline -ErrorAction Stop
    $connectParams = $connectCommand.Parameters.Keys
    $supportsInteractive = $connectParams -contains "Interactive"
    $supportsWebLogin = $connectParams -contains "UseWebLogin"
    $supportsClientId = $connectParams -contains "ClientId"

    return Invoke-WithRetry -ActionName "Connect to $Label site" -ScriptBlock {
        if ($supportsInteractive) {
            $connectArgs = @{ Url = $SiteUrl; ReturnConnection = $true }
            if ($supportsClientId -and -not [string]::IsNullOrWhiteSpace($EffectiveClientId)) {
                $connectArgs.ClientId = $EffectiveClientId
            }

            return Connect-PnPOnline @connectArgs -Interactive -ErrorAction Stop
        }

        if ($supportsWebLogin) {
            return Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ReturnConnection -ErrorAction Stop
        }

        Write-Log "Falling back to credential prompt sign-in for $Label site." "WARN"
        $credentials = Get-Credential -Message "Enter your SharePoint Online credentials for $Label site"
        return Connect-PnPOnline -Url $SiteUrl -Credentials $credentials -ReturnConnection -ErrorAction Stop
    }
}

function Get-TargetPageMap {
    param(
        [Parameter(Mandatory = $true)]
        [object]$TargetConnection
    )

    $map = @{}

    $targetItems = Invoke-WithRetry -ActionName "Load target pages" -ScriptBlock {
        Get-PnPListItem -List $SITE_PAGES_CONFIG.TARGET_LIST_TITLE -Fields "FileLeafRef", "FileRef" -PageSize 2000 -Connection $TargetConnection -ErrorAction Stop
    }

    foreach ($item in $targetItems) {
        $leafRef = [string]$item["FileLeafRef"]
        $fileRef = [string]$item["FileRef"]

        if (-not [string]::IsNullOrWhiteSpace($leafRef)) {
            $map[$leafRef.ToLowerInvariant()] = [PSCustomObject]@{
                FileLeafRef = $leafRef
                FileRef     = $fileRef
                Id          = $item.Id
            }
        }
    }

    return $map
}

try {
    Write-Log "================================"
    Write-Log "Export/Import Site Pages Started"
    Write-Log "================================"
    Write-Log "Source Site URL: $SourceSiteUrl"
    Write-Log "Target Site URL: $TargetSiteUrl"
    Write-Log "Page Names CSV: $PageNamesCsv"
    Write-Log "Dry Run Mode: $DryRun"

    Write-Log "Checking for supported PnP PowerShell module..."
    $moduleToUse = $null
    if (Get-Module -ListAvailable -Name PnP.PowerShell) {
        $moduleToUse = "PnP.PowerShell"
    }
    elseif (Get-Module -ListAvailable -Name SharePointPnPPowerShellOnline) {
        $moduleToUse = "SharePointPnPPowerShellOnline"
    }
    else {
        throw "No supported PnP module found. Install PnP.PowerShell: Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force"
    }

    Import-Module $moduleToUse -ErrorAction Stop
    Write-Log "Using module: $moduleToUse"

    Write-Log "Configuring secure connection settings..."
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    [Net.ServicePointManager]::Expect100Continue = $true
    [Net.ServicePointManager]::DefaultConnectionLimit = 1024
    [Net.ServicePointManager]::MaxServicePointIdleTime = 100000

    $effectiveClientId = if (-not [string]::IsNullOrWhiteSpace($ClientId)) { $ClientId } else { $env:ENTRAID_APP_ID }

    $sourceConnection = Connect-SharePointSite -SiteUrl $SourceSiteUrl -Label "source" -EffectiveClientId $effectiveClientId
    Write-Log "Connected to source site successfully." "SUCCESS"

    $targetConnection = Connect-SharePointSite -SiteUrl $TargetSiteUrl -Label "target" -EffectiveClientId $effectiveClientId
    Write-Log "Connected to target site successfully." "SUCCESS"

    Invoke-WithRetry -ActionName "Validate source Site Pages list" -ScriptBlock {
        Get-PnPList -Identity $SITE_PAGES_CONFIG.SOURCE_LIST_TITLE -Connection $sourceConnection -ErrorAction Stop | Out-Null
    }

    Invoke-WithRetry -ActionName "Validate target Site Pages list" -ScriptBlock {
        Get-PnPList -Identity $SITE_PAGES_CONFIG.TARGET_LIST_TITLE -Connection $targetConnection -ErrorAction Stop | Out-Null
    }

    $requestedPagesSet = Get-RequestedPagesSet -NamesCsv $PageNamesCsv

    if ($requestedPagesSet) {
        Write-Log "Filtered mode enabled. Requested page count: $($requestedPagesSet.Count)"
    }
    else {
        Write-Log "No page filter provided. All pages from source Site Pages library will be processed."
    }

    Write-Log "Loading source Site Pages items..."
    $sourceItems = Invoke-WithRetry -ActionName "Load source pages" -ScriptBlock {
        Get-PnPListItem -List $SITE_PAGES_CONFIG.SOURCE_LIST_TITLE -Fields "FileLeafRef", "FileRef", "Title", "Description" -PageSize 2000 -Connection $sourceConnection -ErrorAction Stop
    }

    $pagesToCopy = @()
    foreach ($item in $sourceItems) {
        $fileLeafRef = [string]$item["FileLeafRef"]
        $fileRef = [string]$item["FileRef"]

        if ([string]::IsNullOrWhiteSpace($fileLeafRef) -or [string]::IsNullOrWhiteSpace($fileRef)) {
            continue
        }

        if (-not $fileLeafRef.ToLowerInvariant().EndsWith($SITE_PAGES_CONFIG.PAGE_EXTENSION)) {
            continue
        }

        if ($requestedPagesSet -and -not $requestedPagesSet.Contains($fileLeafRef)) {
            continue
        }

        $pagesToCopy += [PSCustomObject]@{
            Item        = $item
            FileLeafRef = $fileLeafRef
            FileRef     = $fileRef
        }
    }

    if ($requestedPagesSet) {
        $foundSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($page in $pagesToCopy) {
            [void]$foundSet.Add($page.FileLeafRef)
        }

        foreach ($requestedPage in $requestedPagesSet) {
            if (-not $foundSet.Contains($requestedPage)) {
                Add-PageSkip -PageName $requestedPage -Reason "Requested page not found in source site"
            }
        }
    }

    if (-not $pagesToCopy -or $pagesToCopy.Count -eq 0) {
        throw "No pages found to process based on the provided filter and source Site Pages library."
    }

    Write-Log "Total pages selected for processing: $($pagesToCopy.Count)"

    Write-Log "Loading existing target page map..."
    $targetPageMap = Get-TargetPageMap -TargetConnection $targetConnection
    Write-Log "Existing pages in target library: $($targetPageMap.Count)"

    $tempFolder = Join-Path $PSScriptRoot $RUNTIME.TEMP_EXPORT_FOLDER_NAME
    if (-not (Test-Path $tempFolder)) {
        New-Item -Path $tempFolder -ItemType Directory -Force | Out-Null
    }

    $processed = 0
    for ($index = 0; $index -lt $pagesToCopy.Count; $index += $BatchSize) {
        $endIndex = [Math]::Min($index + $BatchSize - 1, $pagesToCopy.Count - 1)
        $currentBatch = $pagesToCopy[$index..$endIndex]

        Write-Log "Processing batch: pages $($index + 1) to $($endIndex + 1)"

        foreach ($page in $currentBatch) {
            $processed++
            $pageName = $page.FileLeafRef
            $sourceFileRef = $page.FileRef
            $targetKey = $pageName.ToLowerInvariant()

            try {
                Write-Log ("Processing page {0}/{1}: {2}" -f $processed, $pagesToCopy.Count, $pageName)

                if ($targetPageMap.ContainsKey($targetKey)) {
                    Add-PageSkip -PageName $pageName -Reason "Page already exists in target site. Skipping."
                    continue
                }

                if ($DryRun) {
                    Write-Log "DRY RUN: Would copy page '$pageName' from source to target."
                    $script:CopiedCount++
                    continue
                }

                $localExportPath = Join-Path $tempFolder $pageName

                Invoke-WithRetry -ActionName "Export page '$pageName'" -ScriptBlock {
                    Get-PnPFile -Url $sourceFileRef -Path $tempFolder -FileName $pageName -AsFile -Force -Connection $sourceConnection -ErrorAction Stop | Out-Null
                }

                Invoke-WithRetry -ActionName "Import page '$pageName'" -ScriptBlock {
                    Add-PnPFile -Path $localExportPath -Folder $SITE_PAGES_CONFIG.TARGET_FOLDER_NAME -Connection $targetConnection -ErrorAction Stop | Out-Null
                }

                $metadata = @{}
                $sourceTitle = [string]$page.Item[$METADATA_FIELDS.TITLE]
                $sourceDescription = [string]$page.Item[$METADATA_FIELDS.DESCRIPTION]

                if (-not [string]::IsNullOrWhiteSpace($sourceTitle)) {
                    $metadata[$METADATA_FIELDS.TITLE] = $sourceTitle
                }

                if (-not [string]::IsNullOrWhiteSpace($sourceDescription)) {
                    $metadata[$METADATA_FIELDS.DESCRIPTION] = $sourceDescription
                }

                if ($metadata.Count -gt 0) {
                    $itemQuery = "<View><Query><Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='File'>$pageName</Value></Eq></Where></Query><RowLimit>1</RowLimit></View>"
                    $targetItem = Invoke-WithRetry -ActionName "Get target list item for '$pageName'" -ScriptBlock {
                        Get-PnPListItem -List $SITE_PAGES_CONFIG.TARGET_LIST_TITLE -Query $itemQuery -Connection $targetConnection -ErrorAction Stop
                    }

                    $targetItemSingle = @($targetItem) | Select-Object -First 1

                    if ($targetItemSingle) {
                        Invoke-WithRetry -ActionName "Apply metadata for '$pageName'" -ScriptBlock {
                            Set-PnPListItem -List $SITE_PAGES_CONFIG.TARGET_LIST_TITLE -Identity $targetItemSingle.Id -Values $metadata -Connection $targetConnection -ErrorAction Stop | Out-Null
                        }
                    }
                }

                if ($RUNTIME.PUBLISH_AFTER_IMPORT) {
                    try {
                        Invoke-WithRetry -ActionName "Publish page '$pageName'" -ScriptBlock {
                            Publish-PnPFile -Url "$($SITE_PAGES_CONFIG.TARGET_FOLDER_NAME)/$pageName" -Comment "Imported by ExportImportSitePages script" -Connection $targetConnection -ErrorAction Stop
                        }
                    }
                    catch {
                        Write-Log "Publish warning for '$pageName': $($_.Exception.Message)" "WARN"
                    }
                }

                $script:CopiedCount++
                Write-Log "Successfully copied page: $pageName" "SUCCESS"
            }
            catch {
                Add-PageError -PageName $pageName -Reason "Failed to copy page: $($_.Exception.Message)"
            }
            finally {
                try {
                    $localExportPath = Join-Path $tempFolder $pageName
                    if (Test-Path $localExportPath) {
                        Remove-Item -Path $localExportPath -Force -ErrorAction SilentlyContinue
                    }
                }
                catch {
                    # Ignore temp file cleanup errors
                }
            }
        }
    }

    Write-Log "================================"
    Write-Log "Export/Import Site Pages Summary"
    Write-Log "================================"
    Write-Log "Total Pages Selected: $($pagesToCopy.Count)"
    Write-Log "Successfully Copied: $($script:CopiedCount)"
    Write-Log "Skipped: $($script:SkippedCount)"
    Write-Log "Errors: $($script:ErrorCount)"

    if ($script:PageSkips.Count -gt 0) {
        Write-Log "Page Skip Details:" "WARN"
        foreach ($pageSkip in ($script:PageSkips | Sort-Object PageName)) {
            Write-Log "  Page '$($pageSkip.PageName)': $($pageSkip.Reason)" "WARN"
        }
    }

    if ($script:PageErrors.Count -gt 0) {
        Write-Log "Page Error Details:" "ERROR"
        foreach ($pageError in ($script:PageErrors | Sort-Object PageName)) {
            Write-Log "  Page '$($pageError.PageName)': $($pageError.Reason)" "ERROR"
        }
    }

    Write-Log "================================"

    if ($DryRun) {
        Write-Log "Dry Run Mode: No pages were actually copied to target SharePoint site."
    }
    else {
        Write-Log "Export/Import process completed successfully."
    }
}
catch {
    Write-Log "Fatal Error: $($_.Exception.Message)" "ERROR"
    Write-Host "An error occurred. Check log file: $LogFilePath"
    exit 1
}
finally {
    Write-Log "Disconnecting from SharePoint..."

    try {
        if ($sourceConnection) {
            Disconnect-PnPOnline -Connection $sourceConnection -ErrorAction SilentlyContinue
        }
    }
    catch {
        # Ignore disconnect errors
    }

    try {
        if ($targetConnection) {
            Disconnect-PnPOnline -Connection $targetConnection -ErrorAction SilentlyContinue
        }
    }
    catch {
        # Ignore disconnect errors
    }

    try {
        $tempFolder = Join-Path $PSScriptRoot $RUNTIME.TEMP_EXPORT_FOLDER_NAME
        if (Test-Path $tempFolder) {
            Remove-Item -Path $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    catch {
        # Ignore temp folder cleanup errors
    }

    Write-Log "Script execution completed."
}
