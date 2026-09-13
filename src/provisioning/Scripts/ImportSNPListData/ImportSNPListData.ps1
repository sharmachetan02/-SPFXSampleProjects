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
    TITLE            = "Title"
    TYPE             = "Type"
    CITY             = "City"
    COMMENTS         = "Comments"
    COUNTRY          = "Country"
    EUTELSAT_ENTITY  = "Eutelsat Entity"
    STATUS           = "Status"
    SUMMARY          = "Summary"
    TELEPORT_PARTNER = "Teleport Partner"
}

# ================================
# Configuration - SharePoint Field Internal Names
# ================================
$SP_FIELDS = @{
    TITLE            = "Title"
    CONTENTTYPE_ID   = "ContentTypeId"
    CITY             = "MA_City"
    COMMENTS         = "MA_Comments"
    COUNTRY          = "MA_Country"
    EUTELSAT_ENTITY  = "MA_EutelsatEntity"
    STATUS           = "MA_SNPStatus"
    SUMMARY          = "MA_Summary"
    TELEPORT_PARTNER = "MA_TeleportPartner"
    DOCUMENTS_SPACE  = "MA_DocumentsSpace"
}

# ================================
# Configuration - SharePoint List Names
# ================================
$SP_LISTS = @{
    COUNTRY          = "Country"
    EUTELSAT_ENTITY  = "Eutelsat Entity"
    TELEPORT_PARTNER = "Teleport Partner"
}

# ================================
# Configuration - Document Space Settings
# ================================
$DOCUMENT_SPACE = @{
    LIBRARY_NAME              = "documents_space"
    SNP_STARGATE_FOLDER       = "/SNPs/Stargates"
    SNP_EUTELSAT_SNP_FOLDER   = "/SNPs/Eutelsat"
    SNP_PARTNER_SNP_FOLDER    = "/SNPs/Partners"
    HYPERLINK_DESC            = "Link"
}

# ================================
# Configuration - Content Type Mappings
# ================================
$CONTENT_TYPES = @{
    "Stargate"      = "0x01007DBD40A3236C4447848111DE54A6D86301"
    "Eutelsat Owned"  = "0x01007DBD40A3236C4447848111DE54A6D86302"
    "Other"   = "0x01007DBD40A3236C4447848111DE54A6D86303"
    #"PARTNER_SNP"   = "0x01007DBD40A3236C4447848111DE54A6D86303"
}

# ================================
# Configuration - Folder URL Map with type
# ================================
$FOLDER_URL_MAP = @{
    "Stargate"     = $DOCUMENT_SPACE.SNP_STARGATE_FOLDER
    "Eutelsat Owned" = $DOCUMENT_SPACE.SNP_EUTELSAT_SNP_FOLDER
    "Other"  = $DOCUMENT_SPACE.SNP_PARTNER_SNP_FOLDER
    #"Eutelsat SNP" = $DOCUMENT_SPACE.SNP_EUTELSAT_SNP_FOLDER
    #"Partner SNP"  = $DOCUMENT_SPACE.SNP_PARTNER_SNP_FOLDER
}

$STATUS_MAPPING = @{
    "active"    = "Active"
    "commissioning"  = "Commissioning"
    "maintenance" = "Maintenance"
    "decommissioned" = "Decommissioned"
}


# Initialize script variables
$script:LogInitialized = $false
$script:ImportedCount = 0
$script:ErrorCount = 0
$script:RowErrors = @()

# Logging function
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )

    # Initialize/clear log file on first use
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

# Function to sanitize folder name for SharePoint
function Get-SanitizedFolderName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FolderName
    )

    if ([string]::IsNullOrWhiteSpace($FolderName)) {
        return ""
    }

    # Remove forbidden SharePoint characters and non-ASCII
    $sanitized = $FolderName -replace '[~"''#%&*:<>?/\\{|}]', '' -replace '[^\u0020-\u007E]', ''
    # Remove leading/trailing dots and spaces
    $sanitized = $sanitized -replace '^[.\s]+|[.\s]+$', ''
    # Collapse multiple spaces
    $sanitized = $sanitized -replace '\s+', ' '
    # Trim
    $sanitized = $sanitized.Trim()

    return $sanitized
}

# Function to create a unique folder in SharePoint
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
            
            # Check if folder exists
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
                # Create the folder
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

# Function to get lookup ID from a list
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

# Function to sanitize title for SharePoint
function Get-SanitizedTitle {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title
    )

    if ([string]::IsNullOrWhiteSpace($Title)) {
        return ""
    }

    # Remove forbidden SharePoint characters
    $sanitized = $Title -replace '[~"''#%&*:<>?/\\{|}]', '' -replace '[^\u0020-\u007E]', ''
    # Collapse multiple spaces
    $sanitized = $sanitized -replace '\s+', ' '
    # Trim
    $sanitized = $sanitized.Trim()

    return $sanitized
}

# Main script
try {
    Write-Log "================================"
    Write-Log "SNP List Data Import Started"
    Write-Log "================================"
    Write-Log "Site URL: $SiteUrl"
    Write-Log "List Name: $ListName"
    Write-Log "CSV File: $CsvFilePath"
    Write-Log "Dry Run Mode: $DryRun"

    # Check if a supported PnP module is installed (prefer modern module)
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

    # Validate CSV file exists
    if (-not (Test-Path $CsvFilePath)) {
        throw "CSV file not found: $CsvFilePath"
    }

    # Force TLS 1.2 and configure service point manager for SharePoint Online
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
    
    # Retry connection logic
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

    # Get the list with retry logic
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

    # Read CSV file
    Write-Log "Reading CSV file: $CsvFilePath"
    $csvData = Import-Csv -Path $CsvFilePath -Encoding UTF8

    if (-not $csvData -or $csvData.Count -eq 0) {
        throw "No data found in CSV file."
    }

    Write-Log "Total records in CSV: $($csvData.Count)"

    # Process records in batches
    $batch = @()
    $recordCount = 0

    foreach ($record in $csvData) {
        $recordCount++

        try {
            Write-Log "Processing record $recordCount/$($csvData.Count): SNP=$($record.($CSV_COLUMNS.TITLE))"

            # Validate required fields
            if ([string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.TITLE))) {
                Add-RowError -RowNumber $recordCount -Reason "Title field is empty"
                continue
            }

            # Validate Type field
            if ([string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.TYPE))) {
                Add-RowError -RowNumber $recordCount -Reason "Type field is empty"
                continue
            }

            # Build the item hashtable
            $itemData = @{
                Title = (Get-SanitizedTitle -Title $record.($CSV_COLUMNS.TITLE).Trim())
            }

            # Map Type to ContentTypeId
            $snpType = $record.($CSV_COLUMNS.TYPE).Trim()
            if ($CONTENT_TYPES.ContainsKey($snpType)) {
                $itemData[$SP_FIELDS.CONTENTTYPE_ID] = $CONTENT_TYPES[$snpType]
            }
            else {
               Write-Log "Warning: Unknown SNP Type '$snpType' for record $recordCount. Using Type value as-is." "WARN"
                # Store the type value for folder mapping
            }

            # Map City (Single line of text)
            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.CITY))) {
                $itemData[$SP_FIELDS.CITY] = $record.($CSV_COLUMNS.CITY).Trim()
            }

            # Map Summary (Multiple lines of text)
            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.SUMMARY))) {
                $itemData[$SP_FIELDS.SUMMARY] = $record.($CSV_COLUMNS.SUMMARY).Trim()
            }

            # Map Comments (Multiple lines of text)
            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.COMMENTS))) {
                $itemData[$SP_FIELDS.COMMENTS] = $record.($CSV_COLUMNS.COMMENTS).Trim()
            }

            # Map Status (Choice)          
            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.STATUS))) {
                $statusValue = $record.($CSV_COLUMNS.STATUS).Trim()
                $mappedStatus = Get-MappedStatus -StatusValue $statusValue
                if ($mappedStatus) {
                    $itemData[$SP_FIELDS.STATUS] = $mappedStatus
                }              
            }

            # Map Country (Lookup)
            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.COUNTRY))) {
                $countryId = Get-LookupId -ListName $SP_LISTS.COUNTRY -FieldName "Title" -FieldValue $record.($CSV_COLUMNS.COUNTRY).Trim()
                if ($countryId) {
                    $itemData[$SP_FIELDS.COUNTRY] = $countryId
                }
            }

            # Map Eutelsat Entity (Lookup)
            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.EUTELSAT_ENTITY))) {
                $entityId = Get-LookupId -ListName $SP_LISTS.EUTELSAT_ENTITY -FieldName "Title" -FieldValue $record.($CSV_COLUMNS.EUTELSAT_ENTITY).Trim()
                if ($entityId) {
                    $itemData[$SP_FIELDS.EUTELSAT_ENTITY] = $entityId
                }
            }

            # Map Teleport Partner (Lookup)
            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.TELEPORT_PARTNER))) {
                $partnerId = Get-LookupId -ListName $SP_LISTS.TELEPORT_PARTNER -FieldName "Title" -FieldValue $record.($CSV_COLUMNS.TELEPORT_PARTNER).Trim()
                if ($partnerId) {
                    $itemData[$SP_FIELDS.TELEPORT_PARTNER] = $partnerId
                }
            }

            # Add to batch with record reference for type mapping
            $batch += [PSCustomObject]@{
                Data   = $itemData
                Index  = $recordCount
                Record = $record
            }

            # Process batch when it reaches BatchSize or at the end
            if ($batch.Count -ge $BatchSize -or $recordCount -eq $csvData.Count) {
                Write-Log "Processing batch with $($batch.Count) items..."

                foreach ($batchItem in $batch) {
                    try {
                        if ($DryRun) {
                            Write-Log "DRY RUN: Would create item for SNP: $($batchItem.Data.Title)"
                        }
                        else {
                            # Add new item with retry logic
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
                            
                            # Create Document Space folder and update the hyperlink field
                            try {
                                # Get the SNP type from the batch item record
                                $snpType = $batchItem.Record.($CSV_COLUMNS.TYPE).Trim()
                                
                                # Determine the parent folder based on type using folder URL map
                                $parentFolderPath = $FOLDER_URL_MAP[$snpType]
                                if (-not $parentFolderPath) {
                                    Write-Log "Warning: No folder mapping found for type '$snpType'. Using default SNP_STARGATE_FOLDER" "WARN"
                                    $parentFolderPath = $DOCUMENT_SPACE.SNP_STARGATE_FOLDER
                                }
                                
                                $sanitizedFolderName = Get-SanitizedFolderName -FolderName $batchItem.Data.Title
                                $parentFolderUrl = "$($DOCUMENT_SPACE.LIBRARY_NAME)$parentFolderPath"
                                
                                # Create unique folder
                                $folderUrl = New-UniqueFolder -ParentFolderUrl $parentFolderUrl -BaseName $sanitizedFolderName
                                
                                # Get the site URL to construct the full folder URL
                                $web = Get-PnPWeb
                                $siteRelativeFolderUrl = "$($web.ServerRelativeUrl.TrimEnd('/'))/$folderUrl"
                                $folderUrl = $siteRelativeFolderUrl
                                
                                # Update the Documents Space field with hyperlink value
                                $hyperlinkValue = "$folderUrl, $($DOCUMENT_SPACE.HYPERLINK_DESC)"
                                Set-PnPListItem -List $ListName -Identity $itemId -Values @{
                                    $SP_FIELDS.DOCUMENTS_SPACE = $hyperlinkValue
                                } -ErrorAction Stop
                                
                                Write-Log "✅ Created document space folder and updated field for item ID $itemId" "INFO"
                            }
                            catch {
                                Write-Log "⚠ Warning: Failed to create document space folder for item ID $itemId : $_" "WARN"
                                # Continue processing despite folder creation failure
                            }
                            
                            Write-Log "✅ Successfully added item ID $itemId for SNP: $($batchItem.Data.Title)" "SUCCESS"
                            $script:ImportedCount++
                        }
                    }
                    catch {
                        Add-RowError -RowNumber $batchItem.Index -Reason "Error adding/updating item for SNP '$($batchItem.Data.Title)': $($_.Exception.Message)"
                    }
                }

                $batch = @()
            }
        }
        catch {
            Add-RowError -RowNumber $recordCount -Reason "Error processing record: $($_.Exception.Message)"
        }
    }

    # Summary
    Write-Log "================================"
    Write-Log "Import Process Summary"
    Write-Log "================================"
    Write-Log "Total Records Processed: $recordCount"
    Write-Log "✅ Successfully Imported: $($script:ImportedCount)"
    Write-Log "❌ Errors: $($script:ErrorCount)"

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
