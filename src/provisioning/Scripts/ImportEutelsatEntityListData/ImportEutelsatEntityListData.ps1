param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $true)]
    [string]$ListName,

    [Parameter(Mandatory = $true)]
    [string]$CsvFilePath,

    [Parameter(Mandatory = $true)]
    [string]$LogFilePath,

    [Parameter(Mandatory = $false)]
    [int]$BatchSize = 50,

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [switch]$DryRun
)

# ================================
# Configuration - CSV Columns Mapping
# ================================
$CSV_COLUMNS = @{
    TITLE         = "Title"
    COMMENTS      = "Comments"
    COUNTRY       = "Country"
    LEGAL_ADDRESS = "Legal Address"
    STATUS        = "Status"
    SUMMARY       = "Summary"
}

# ================================
# Configuration - SharePoint Field Internal Names
# ================================
$SP_FIELDS = @{
    TITLE          = "Title"
    CONTENTTYPE_ID = "ContentTypeId"
    COMMENTS       = "MA_Comments"
    COUNTRY        = "MA_Country"
    LEGAL_ADDRESS  = "MA_Address"
    STATUS         = "MA_EutelsatEntityStatus"
    SUMMARY        = "MA_Summary"
    DOCUMENTS_SPACE = "MA_DocumentsSpace"
}

# ================================
# Configuration - SharePoint List Names
# ================================
$SP_LISTS = @{
    COUNTRY         = "Country"
    EUTELSAT_ENTITY = "Eutelsat Entity"
}

# ================================
# Configuration - Document Space Settings
# ================================
$DOCUMENT_SPACE = @{
    LIBRARY_NAME   = "documents_space"
    MARKET_FOLDER  = "/Eutelsat Entities"
    HYPERLINK_DESC = "Link"
}

# Content Type ID
$CONTENT_TYPE = "0x0100429E3F54EE591F459DA18E05E550F989"

# Initialize script variables
$script:LogInitialized = $false
$script:ImportedCount = 0
$script:SkippedCount = 0
$script:ErrorCount = 0
$script:RowErrors = @()
$script:RowSkips = @()

$STATUS_MAPPING = @{
    "active"    = "Active"
    "inactive"  = "Inactive"
    "dissolved" = "Dissolved"
}

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

function Add-RowError {
    param(
        [Parameter(Mandatory = $true)]
        [int]$RowNumber,

        [Parameter(Mandatory = $true)]
        [string]$Reason
    )

    $script:ErrorCount++
    $errorEntry = [PSCustomObject]@{
        RowNumber = $RowNumber
        Reason    = $Reason
    }
    $script:RowErrors += $errorEntry
    Write-Log "Row ${RowNumber}: $Reason" "ERROR"
}

function Add-RowSkip {
    param(
        [Parameter(Mandatory = $true)]
        [int]$RowNumber,

        [Parameter(Mandatory = $true)]
        [string]$Reason
    )

    $script:SkippedCount++
    $skipEntry = [PSCustomObject]@{
        RowNumber = $RowNumber
        Reason    = $Reason
    }
    $script:RowSkips += $skipEntry
    Write-Log "Row ${RowNumber}: $Reason" "WARN"
}

function Get-MappedStatus {
    param(
        [Parameter(Mandatory = $false)]
        [string]$StatusValue
    )

    if ([string]::IsNullOrWhiteSpace($StatusValue)) {
        return $null
    }

    $normalizedStatus = $StatusValue.Trim().ToLowerInvariant()
    if ($STATUS_MAPPING.ContainsKey($normalizedStatus)) {
        return $STATUS_MAPPING[$normalizedStatus]
    }

    return $null
}

function Get-SanitizedFolderName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FolderName
    )

    if ([string]::IsNullOrWhiteSpace($FolderName)) {
        return ""
    }

    $sanitized = $FolderName -replace '[~"''#%&*:<>?/\\{|}]', '' -replace '[^\u0020-\u007E]', ''
    $sanitized = $sanitized -replace '^[.\s]+|[.\s]+$', ''
    $sanitized = $sanitized -replace '\s+', ' '
    $sanitized = $sanitized.Trim()

    return $sanitized
}

function New-UniqueFolder {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ParentFolderUrl,

        [Parameter(Mandatory = $true)]
        [string]$BaseName
    )

    try {
        $folderName = $BaseName
        $attempt = 0
        $maxAttempts = 100

        while ($attempt -lt $maxAttempts) {
            $targetFolderUrl = "$ParentFolderUrl/$folderName"

            $folderExists = $false
            try {
                $folder = Get-PnPFolder -Url $targetFolderUrl -ErrorAction SilentlyContinue
                if ($folder) {
                    $folderExists = $true
                }
            }
            catch {
                $folderExists = $false
            }

            if ($folderExists) {
                $attempt++
                $folderName = "$BaseName-$attempt"
            }
            else {
                Add-PnPFolder -Name $folderName -Folder $ParentFolderUrl -ErrorAction Stop | Out-Null
                Write-Log "Created folder: $targetFolderUrl" "INFO"
                return $targetFolderUrl
            }
        }

        throw "Failed to create unique folder after $maxAttempts attempts"
    }
    catch {
        Write-Log "Error creating unique folder: $_" "ERROR"
        throw
    }
}

function Get-LookupMap {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ListName,

        [Parameter(Mandatory = $false)]
        [string]$FieldName = "Title"
    )

    $lookupMap = @{}
    $retryCount = 0
    $maxRetries = 3

    while ($retryCount -lt $maxRetries) {
        try {
            $items = Get-PnPListItem -List $ListName -Fields $FieldName -PageSize 2000 -ErrorAction Stop
            foreach ($item in $items) {
                $key = [string]$item[$FieldName]
                if (-not [string]::IsNullOrWhiteSpace($key)) {
                    $normalizedKey = $key.Trim().ToLowerInvariant()
                    if (-not $lookupMap.ContainsKey($normalizedKey)) {
                        $lookupMap[$normalizedKey] = $item.Id
                    }
                }
            }

            Write-Log "Loaded $($lookupMap.Count) entries from lookup list '$ListName'." "INFO"
            return $lookupMap
        }
        catch {
            $retryCount++
            if ($retryCount -lt $maxRetries) {
                Write-Log "Failed to preload lookup map '$ListName' (attempt $retryCount). Retrying..." "WARN"
                Start-Sleep -Seconds 2
            }
            else {
                throw "Failed to preload lookup map '$ListName' after $maxRetries attempts: $_"
            }
        }
    }

    return $lookupMap
}

try {
    Write-Log "================================"
    Write-Log "Eutelsat Entity List Data Import Started"
    Write-Log "================================"
    Write-Log "Site URL: $SiteUrl"
    Write-Log "List Name: $ListName"
    Write-Log "CSV File: $CsvFilePath"
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
        throw "No supported PnP module found. Install PnP.PowerShell (recommended): Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force"
    }

    Import-Module $moduleToUse -ErrorAction Stop
    Write-Log "✅ Using module: $moduleToUse"

    if (-not (Test-Path $CsvFilePath)) {
        throw "CSV file not found: $CsvFilePath"
    }

    Write-Log "Configuring secure connection settings..."
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    [Net.ServicePointManager]::Expect100Continue = $true
    [Net.ServicePointManager]::DefaultConnectionLimit = 1024
    [Net.ServicePointManager]::MaxServicePointIdleTime = 100000

    Write-Log "Connecting to SharePoint Online site: $SiteUrl"
    $connectCommand = Get-Command Connect-PnPOnline -ErrorAction Stop
    $connectParams = $connectCommand.Parameters.Keys
    $effectiveClientId = if (-not [string]::IsNullOrWhiteSpace($ClientId)) { $ClientId } else { $env:ENTRAID_APP_ID }

    $retryCount = 0
    $maxRetries = 3
    $connected = $false

    while (-not $connected -and $retryCount -lt $maxRetries) {
        try {
            if ($connectParams -contains "Interactive") {
                Write-Log "Using interactive browser sign-in..."
                $connectArgs = @{ Url = $SiteUrl }

                if (($connectParams -contains "ClientId") -and -not [string]::IsNullOrWhiteSpace($effectiveClientId)) {
                    $connectArgs.ClientId = $effectiveClientId
                }

                Connect-PnPOnline @connectArgs -Interactive -ErrorAction Stop
            }
            elseif ($connectParams -contains "UseWebLogin") {
                Write-Log "Using web login sign-in..."
                Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ErrorAction Stop
            }
            else {
                Write-Log "Falling back to credential prompt sign-in..." "WARN"
                $credentials = Get-Credential -Message "Enter your SharePoint Online credentials"
                Connect-PnPOnline -Url $SiteUrl -Credentials $credentials -ErrorAction Stop
            }

            $connected = $true
            Write-Log "✅ Connected successfully."
        }
        catch {
            $retryCount++
            if ($retryCount -lt $maxRetries) {
                Write-Log "Connection attempt $retryCount failed. Retrying..." "WARN"
                Start-Sleep -Seconds 2
            }
            else {
                throw "Failed to connect after $maxRetries attempts: $_"
            }
        }
    }

    Write-Log "Fetching SharePoint list: $ListName"
    $list = $null
    $retryCount = 0

    while ($null -eq $list -and $retryCount -lt $maxRetries) {
        try {
            $list = Get-PnPList -Identity $ListName -ErrorAction Stop
            Write-Log "✅ List found."
        }
        catch {
            $retryCount++
            if ($retryCount -lt $maxRetries) {
                Write-Log "Failed to fetch list (attempt $retryCount). Retrying..." "WARN"
                Start-Sleep -Seconds 2
            }
            else {
                throw "Failed to fetch list '$ListName' after $maxRetries attempts: $_"
            }
        }
    }

    if (-not $list) {
        throw "List '$ListName' not found in site."
    }

    Write-Log "Reading CSV file: $CsvFilePath"
    $csvData = Import-Csv -Path $CsvFilePath -Encoding UTF8

    if (-not $csvData -or $csvData.Count -eq 0) {
        throw "No data found in CSV file."
    }

    Write-Log "Total records in CSV: $($csvData.Count)"

    Write-Log "Preloading lookup maps..."
    $countryLookupMap = Get-LookupMap -ListName $SP_LISTS.COUNTRY -FieldName "Title"

    $batch = @()
    $recordCount = 0

    foreach ($record in $csvData) {
        $recordCount++

        try {
            Write-Log "Processing record $recordCount/$($csvData.Count): Entity=$($record.($CSV_COLUMNS.TITLE))"

            if ([string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.TITLE))) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Title field is empty"
                continue
            }

            if ([string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.COUNTRY))) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Country field is empty"
                continue
            }

            $itemData = @{
                $SP_FIELDS.TITLE          = $record.($CSV_COLUMNS.TITLE).Trim()
                $SP_FIELDS.CONTENTTYPE_ID = $CONTENT_TYPE
            }

            $countryValue = $record.($CSV_COLUMNS.COUNTRY).Trim()
            $countryKey = $countryValue.ToLowerInvariant()
            $countryId = if ($countryLookupMap.ContainsKey($countryKey)) { $countryLookupMap[$countryKey] } else { $null }

            if ($countryId) {
                $itemData[$SP_FIELDS.COUNTRY] = $countryId
            }
            else {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Country lookup not found: '$countryValue'"
                continue
            }

            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.COMMENTS))) {
                $itemData[$SP_FIELDS.COMMENTS] = $record.($CSV_COLUMNS.COMMENTS).Trim()
            }

            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.LEGAL_ADDRESS))) {
                $itemData[$SP_FIELDS.LEGAL_ADDRESS] = $record.($CSV_COLUMNS.LEGAL_ADDRESS).Trim()
            }

            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.STATUS))) {
                $statusValue = $record.($CSV_COLUMNS.STATUS).Trim()
                $mappedStatus = Get-MappedStatus -StatusValue $statusValue
                if ($mappedStatus) {
                    $itemData[$SP_FIELDS.STATUS] = $mappedStatus
                }
                else {
                    Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Invalid Status value: '$statusValue'"
                    continue
                }
            }

            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.SUMMARY))) {
                $itemData[$SP_FIELDS.SUMMARY] = $record.($CSV_COLUMNS.SUMMARY).Trim()
            }

            $batch += [PSCustomObject]@{
                Data  = $itemData
                Index = $recordCount
            }

            if ($batch.Count -ge $BatchSize -or $recordCount -eq $csvData.Count) {
                Write-Log "Processing batch with $($batch.Count) items..."

                foreach ($batchItem in $batch) {
                    try {
                        if ($DryRun) {
                            Write-Log "DRY RUN: Would add item for Eutelsat Entity: $($batchItem.Data.Title)"
                        }
                        else {
                            $newItem = $null
                            $retryCount = 0
                            $maxRetries = 3

                            while ($null -eq $newItem -and $retryCount -lt $maxRetries) {
                                try {
                                    $newItem = Add-PnPListItem -List $ListName -Values $batchItem.Data -ErrorAction Stop
                                }
                                catch {
                                    $retryCount++
                                    if ($retryCount -lt $maxRetries) {
                                        Write-Log "Failed to add item (attempt $retryCount). Retrying..." "WARN"
                                        Start-Sleep -Seconds 2
                                    }
                                    else {
                                        throw "Failed to add item after $maxRetries attempts: $_"
                                    }
                                }
                            }

                            $itemId = if ($newItem -and $newItem.Id) { $newItem.Id } else { "N/A" }

                            try {
                                $sanitizedFolderName = Get-SanitizedFolderName -FolderName $batchItem.Data.Title
                                $parentFolderUrl = "$($DOCUMENT_SPACE.LIBRARY_NAME)$($DOCUMENT_SPACE.MARKET_FOLDER)"

                                $folderUrl = New-UniqueFolder -ParentFolderUrl $parentFolderUrl -BaseName $sanitizedFolderName

                                $web = Get-PnPWeb
                                $siteRelativeFolderUrl = "$($web.ServerRelativeUrl.TrimEnd('/'))/$folderUrl"

                                $hyperlinkValue = "$siteRelativeFolderUrl, $($DOCUMENT_SPACE.HYPERLINK_DESC)"
                                Set-PnPListItem -List $ListName -Identity $newItem.Id -Values @{
                                    $SP_FIELDS.DOCUMENTS_SPACE = $hyperlinkValue
                                } -ErrorAction Stop

                                Write-Log "✅ Created document space folder and updated field for item ID $itemId" "INFO"
                            }
                            catch {
                                Write-Log "⚠ Warning: Failed to create document space folder for item ID $itemId : $_" "WARN"
                            }

                            Write-Log "✅ Successfully added item ID $itemId for Eutelsat Entity: $($batchItem.Data.Title)" "SUCCESS"
                            $script:ImportedCount++
                        }
                    }
                    catch {
                        Add-RowError -RowNumber $batchItem.Index -Reason "Error adding item for Eutelsat Entity '$($batchItem.Data.Title)': $($_.Exception.Message)"
                    }
                }

                $batch = @()
            }
        }
        catch {
            Add-RowError -RowNumber $recordCount -Reason "Error processing record: $($_.Exception.Message)"
        }
    }

    Write-Log "================================"
    Write-Log "Import Process Summary"
    Write-Log "================================"
    Write-Log "Total Records Processed: $recordCount"
    Write-Log "✅ Successfully Imported: $($script:ImportedCount)"
    Write-Log "⊘ Skipped: $($script:SkippedCount)"
    Write-Log "❌ Errors: $($script:ErrorCount)"

    if ($script:RowSkips.Count -gt 0) {
        Write-Log "Row Skip Details:" "WARN"
        foreach ($rowSkip in ($script:RowSkips | Sort-Object RowNumber)) {
            Write-Log "  Row $($rowSkip.RowNumber): $($rowSkip.Reason)" "WARN"
        }
    }

    if ($script:RowErrors.Count -gt 0) {
        Write-Log "Row Error Details:" "ERROR"
        foreach ($rowError in ($script:RowErrors | Sort-Object RowNumber)) {
            Write-Log "  Row $($rowError.RowNumber): $($rowError.Reason)" "ERROR"
        }
    }

    Write-Log "================================"

    if ($DryRun) {
        Write-Log "⚠ Dry Run Mode: No data was actually written to SharePoint"
    }
    else {
        Write-Log "✅ Import completed successfully."
    }
}
catch {
    Write-Log "❌ Fatal Error: $_" "ERROR"
    Write-Host "An error occurred. Check log file: $LogFilePath"
    exit 1
}
finally {
    Write-Log "Disconnecting from SharePoint..."
    try {
        if (Get-PnPConnection -ErrorAction SilentlyContinue) {
            Disconnect-PnPOnline
        }
    }
    catch {
        # Ignore disconnect errors
    }
    Write-Log "Script execution completed."
}
