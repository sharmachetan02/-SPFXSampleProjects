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
    TYPE          = "Type"
    COMMENTS      = "Comments"
    COUNTRY       = "Country"
    LEGAL_ADDRESS = "Legal Address"
    SUMMARY       = "Summary"
    PORTAL        = "Portal"
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
    SUMMARY        = "MA_Summary"
    PORTAL         = "MA_Portal"
    DOCUMENTS_SPACE = "MA_DocumentsSpace"
}

# ================================
# Configuration - SharePoint List Names
# ================================
$SP_LISTS = @{
    COUNTRY   = "Country"
    AUTHORITY = "Authority"
}

# ================================
# Configuration - Document Space Settings
# ================================
$DOCUMENT_SPACE = @{
    LIBRARY_NAME           = "documents_space"
    ADMINISTRATION_FOLDER  = "/Authority/Administration"
    ORGANIZATION_FOLDER    = "/Authority/Organization"
    REGULATOR_FOLDER       = "/Authority/Regulator"
    HYPERLINK_DESC         = "Link"
}

# ================================
# Configuration - Content Type Mappings
# ================================
$CONTENT_TYPES = @{
    "Regulator"      = "0x0100DD016807C9F8854299FFFF08E0D9D2C501"
    "Organization"   = "0x0100DD016807C9F8854299FFFF08E0D9D2C503"
    "Administration" = "0x0100DD016807C9F8854299FFFF08E0D9D2C502"
}

# ================================
# Configuration - Folder URL Map with type
# ================================
$FOLDER_URL_MAP = @{
    "Regulator"      = $DOCUMENT_SPACE.REGULATOR_FOLDER
    "Organization"   = $DOCUMENT_SPACE.ORGANIZATION_FOLDER
    "Administration" = $DOCUMENT_SPACE.ADMINISTRATION_FOLDER
}

# Initialize script variables
$script:LogInitialized = $false
$script:ImportedCount = 0
$script:SkippedCount = 0
$script:ErrorCount = 0
$script:RowErrors = @()
$script:RowSkips = @()

# Logging function
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

function Get-LookupId {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ListName,

        [Parameter(Mandatory = $true)]
        [string]$FieldName,

        [Parameter(Mandatory = $true)]
        [string]$FieldValue
    )

    try {
        if ([string]::IsNullOrWhiteSpace($FieldValue)) {
            return $null
        }

        $item = $null
        $retryCount = 0
        $maxRetries = 3

        while ($null -eq $item -and $retryCount -lt $maxRetries) {
            try {
                $item = Get-PnPListItem -List $ListName -Fields $FieldName -ErrorAction Stop | Where-Object { $_[$FieldName] -eq $FieldValue } | Select-Object -First 1
            }
            catch {
                $retryCount++
                if ($retryCount -lt $maxRetries) {
                    Write-Log "Failed to fetch lookup item (attempt $retryCount). Retrying..." "WARN"
                    Start-Sleep -Seconds 1
                }
                else {
                    throw "Failed to fetch lookup item after $maxRetries attempts: $_"
                }
            }
        }

        if ($item) {
            return $item.Id
        }
        else {
            Write-Log "Lookup value '$FieldValue' not found in list '$ListName'" "WARN"
            return $null
        }
    }
    catch {
        Write-Log "Error retrieving lookup ID for '$FieldValue' in list '$ListName': $_" "ERROR"
        return $null
    }
}

# Main script
try {
    Write-Log "================================"
    Write-Log "Authority List Data Import Started"
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
    $supportsInteractive = $connectParams -contains "Interactive"
    $supportsWebLogin = $connectParams -contains "UseWebLogin"
    $supportsClientId = $connectParams -contains "ClientId"
    $effectiveClientId = if (-not [string]::IsNullOrWhiteSpace($ClientId)) { $ClientId } else { $env:ENTRAID_APP_ID }

    $retryCount = 0
    $maxRetries = 3
    $connected = $false

    while (-not $connected -and $retryCount -lt $maxRetries) {
        try {
            if ($supportsInteractive) {
                Write-Log "Using interactive browser sign-in..."
                $connectArgs = @{ Url = $SiteUrl }

                if ($supportsClientId -and -not [string]::IsNullOrWhiteSpace($effectiveClientId)) {
                    $connectArgs.ClientId = $effectiveClientId
                }

                Connect-PnPOnline @connectArgs -Interactive -ErrorAction Stop
            }
            elseif ($supportsWebLogin) {
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
    $maxRetries = 3

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

    $batch = @()
    $recordCount = 0

    foreach ($record in $csvData) {
        $recordCount++

        try {
            Write-Log "Processing record $recordCount/$($csvData.Count): Authority=$($record.($CSV_COLUMNS.TITLE))"

            if ([string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.TITLE))) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Title field is empty"
                continue
            }

            if ([string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.TYPE))) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Type field is empty"
                continue
            }

            $itemData = @{
                Title = $record.($CSV_COLUMNS.TITLE).Trim()
            }

            $authorityType = $record.($CSV_COLUMNS.TYPE).Trim()
            if ($CONTENT_TYPES.ContainsKey($authorityType)) {
                $itemData[$SP_FIELDS.CONTENTTYPE_ID] = $CONTENT_TYPES[$authorityType]
            }
            else {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Invalid Type value: '$authorityType'"
                continue
            }

            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.COMMENTS))) {
                $itemData[$SP_FIELDS.COMMENTS] = $record.($CSV_COLUMNS.COMMENTS).Trim()
            }

            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.LEGAL_ADDRESS))) {
                $itemData[$SP_FIELDS.LEGAL_ADDRESS] = $record.($CSV_COLUMNS.LEGAL_ADDRESS).Trim()
            }

            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.SUMMARY))) {
                $itemData[$SP_FIELDS.SUMMARY] = $record.($CSV_COLUMNS.SUMMARY).Trim()
            }

            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.COUNTRY))) {
                $countryRawValue = $record.($CSV_COLUMNS.COUNTRY).Trim()
                $countryValues = $countryRawValue -split '[,;|]' |
                    ForEach-Object { $_.Trim() } |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                    Select-Object -Unique

                if (-not $countryValues -or $countryValues.Count -eq 0) {
                    Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Country field is empty"
                    continue
                }

                $countryIds = @()
                $missingCountries = @()

                foreach ($countryValue in $countryValues) {
                    $countryId = Get-LookupId -ListName $SP_LISTS.COUNTRY -FieldName "Title" -FieldValue $countryValue
                    if ($countryId) {
                        $countryIds += [int]$countryId
                    }
                    else {
                        $missingCountries += $countryValue
                    }
                }

                if ($missingCountries.Count -gt 0) {
                    Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Country lookup not found: '$($missingCountries -join ", ")'"
                    continue
                }

                if ($countryIds.Count -gt 0) {
                    $itemData[$SP_FIELDS.COUNTRY] = $countryIds
                }
                else {
                    Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Country lookup resolution returned no IDs"
                    continue
                }
            }
            else {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Country field is empty"
                continue
            }

            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.PORTAL))) {
                $portalUrl = $record.($CSV_COLUMNS.PORTAL).Trim()
                $itemData[$SP_FIELDS.PORTAL] = "$portalUrl, Link"
            }

            $batch += [PSCustomObject]@{
                Data   = $itemData
                Index  = $recordCount
                Record = $record
            }

            if ($batch.Count -ge $BatchSize -or $recordCount -eq $csvData.Count) {
                Write-Log "Processing batch with $($batch.Count) items..."

                foreach ($batchItem in $batch) {
                    try {
                        if ($DryRun) {
                            Write-Log "DRY RUN: Would create item for Authority: $($batchItem.Data.Title)"
                        }
                        else {
                            $retryCount = 0
                            $maxRetries = 3
                            $newItem = $null

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
                                $authorityType = $batchItem.Record.($CSV_COLUMNS.TYPE).Trim()
                                $parentFolderPath = $FOLDER_URL_MAP[$authorityType]

                                if (-not $parentFolderPath) {
                                    Write-Log "Warning: No folder mapping found for type '$authorityType'. Using default REGULATOR folder" "WARN"
                                    $parentFolderPath = $DOCUMENT_SPACE.REGULATOR_FOLDER
                                }

                                $sanitizedFolderName = Get-SanitizedFolderName -FolderName $batchItem.Data.Title
                                $parentFolderUrl = "$($DOCUMENT_SPACE.LIBRARY_NAME)$parentFolderPath"

                                $folderUrl = New-UniqueFolder -ParentFolderUrl $parentFolderUrl -BaseName $sanitizedFolderName

                                $web = Get-PnPWeb
                                $siteRelativeFolderUrl = "$($web.ServerRelativeUrl.TrimEnd('/'))/$folderUrl"

                                $hyperlinkValue = "$siteRelativeFolderUrl, $($DOCUMENT_SPACE.HYPERLINK_DESC)"
                                Set-PnPListItem -List $ListName -Identity $itemId -Values @{
                                    $SP_FIELDS.DOCUMENTS_SPACE = $hyperlinkValue
                                } -ErrorAction Stop

                                Write-Log "✅ Created document space folder and updated field for item ID $itemId" "INFO"
                            }
                            catch {
                                Write-Log "⚠ Warning: Failed to create document space folder for item ID $itemId : $_" "WARN"
                            }

                            Write-Log "✅ Successfully added item ID $itemId for Authority: $($batchItem.Data.Title)" "SUCCESS"
                            $script:ImportedCount++
                        }
                    }
                    catch {
                        Add-RowError -RowNumber $batchItem.Index -Reason "Error adding/updating item for Authority '$($batchItem.Data.Title)': $($_.Exception.Message)"
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
