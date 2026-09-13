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
    FIRST_NAME   = "First Name"
    LAST_NAME    = "Last Name"
    SUMMARY      = "Summary"
    EMAIL        = "Email"
    PHONE_NUMBER = "Phone Number"
    JOB_TITLE    = "Job Title"
    DEPARTMENT   = "Department"
    TYPE         = "Type"
    RELATED_ITEM = "Related Item"
    COMMENTS     = "Comments"
}

# ================================
# Configuration - SharePoint Field Internal Names
# ================================
$SP_FIELDS = @{
    TITLE                  = "Title"
    CONTENTTYPE_ID         = "ContentTypeId"
    FIRST_NAME             = "MA_FirstName"
    LAST_NAME              = "MA_LastName"
    SUMMARY                = "MA_Summary"
    EMAIL                  = "MA_Email"
    PHONE_NUMBER           = "MA_PhoneNumber"
    JOB_TITLE              = "MA_JobTitle"
    DEPARTMENT             = "MA_Department"
    COMMENTS               = "MA_Comments"
    AUTHORITY              = "MA_Authority"
    DISTRIBUTION_PARTNER   = "MA_DistributionPartner"
    LEGAL_SERVICE_PROVIDER = "MA_LegalSvcProvider"
    TELEPORT_PARTNER       = "MA_TeleportPartner"
}

# ================================
# Configuration - SharePoint List Names
# ================================
$SP_LISTS = @{
    CONTACT                = "Contact"
    AUTHORITY              = "Authority"
    DISTRIBUTION_PARTNER   = "Distribution Partner"
    LEGAL_SERVICE_PROVIDER = "Legal Service Provider"
    TELEPORT_PARTNER       = "Teleport Partner"
    EUTELSAT_ENTITY        = "Eutelsat Entity"
}

# ================================
# Configuration - Content Type Mappings
# ================================
$CONTENT_TYPES = @{
    "Regulator"            = "0x01009683AA37F0D20E469880B9832370E8260201"
    "Organization"         = "0x01009683AA37F0D20E469880B9832370E8260203"
    "Administration"       = "0x01009683AA37F0D20E469880B9832370E8260202"
    "Distribution Partner" = "0x01009683AA37F0D20E469880B9832370E82604"
    "Eutelsat"             = "0x01009683AA37F0D20E469880B9832370E82605"
    "Law Firm"             = "0x01009683AA37F0D20E469880B9832370E8260102"
    "Legal Representative" = "0x01009683AA37F0D20E469880B9832370E8260101"
    "Teleport Partner"     = "0x01009683AA37F0D20E469880B9832370E82603"
}

# ================================
# Configuration - Related Item Routing by Type
# ================================
$RELATED_ITEM_ROUTE = @{
    "Regulator" = @{
        ListName = $SP_LISTS.AUTHORITY
        Field    = $SP_FIELDS.AUTHORITY
    }
    "Organization" = @{
        ListName = $SP_LISTS.AUTHORITY
        Field    = $SP_FIELDS.AUTHORITY
    }
    "Administration" = @{
        ListName = $SP_LISTS.AUTHORITY
        Field    = $SP_FIELDS.AUTHORITY
    }
    "Law Firm" = @{
        ListName = $SP_LISTS.LEGAL_SERVICE_PROVIDER
        Field    = $SP_FIELDS.LEGAL_SERVICE_PROVIDER
    }
    "Legal Representative" = @{
        ListName = $SP_LISTS.LEGAL_SERVICE_PROVIDER
        Field    = $SP_FIELDS.LEGAL_SERVICE_PROVIDER
    }
    "Distribution Partner" = @{
        ListName = $SP_LISTS.DISTRIBUTION_PARTNER
        Field    = $SP_FIELDS.DISTRIBUTION_PARTNER
    }
    "Teleport Partner" = @{
        ListName = $SP_LISTS.TELEPORT_PARTNER
        Field    = $SP_FIELDS.TELEPORT_PARTNER
    }
    "Eutelsat" = @{
        ListName = $null
        Field    = $null
    }
}

# Initialize script variables
$script:LogInitialized = $false
$script:ImportedCount = 0
$script:SkippedCount = 0
$script:ErrorCount = 0
$script:RowErrors = @()
$script:RowSkips = @()

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

function Get-NormalizedContentTypeKey {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TypeValue
    )

    $normalizedInput = $TypeValue.Trim().ToLowerInvariant()
    foreach ($key in $CONTENT_TYPES.Keys) {
        if ($key.ToLowerInvariant() -eq $normalizedInput) {
            return $key
        }
    }

    return $null
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
    Write-Log "Contact List Data Import Started"
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
    Write-Log "Using module: $moduleToUse"

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
            Write-Log "Connected successfully."
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
            Write-Log "List found."
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

    Write-Log "Preloading lookup maps..."
    $authorityLookupMap = Get-LookupMap -ListName $SP_LISTS.AUTHORITY
    $distributionPartnerLookupMap = Get-LookupMap -ListName $SP_LISTS.DISTRIBUTION_PARTNER
    $legalServiceProviderLookupMap = Get-LookupMap -ListName $SP_LISTS.LEGAL_SERVICE_PROVIDER
    $teleportPartnerLookupMap = Get-LookupMap -ListName $SP_LISTS.TELEPORT_PARTNER

    $lookupMapsByList = @{
        $SP_LISTS.AUTHORITY = $authorityLookupMap
        $SP_LISTS.DISTRIBUTION_PARTNER = $distributionPartnerLookupMap
        $SP_LISTS.LEGAL_SERVICE_PROVIDER = $legalServiceProviderLookupMap
        $SP_LISTS.TELEPORT_PARTNER = $teleportPartnerLookupMap
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
            $rawFirstName = [string]$record.($CSV_COLUMNS.FIRST_NAME)
            $rawLastName = [string]$record.($CSV_COLUMNS.LAST_NAME)
            $rawType = [string]$record.($CSV_COLUMNS.TYPE)
            $rawRelatedItem = [string]$record.($CSV_COLUMNS.RELATED_ITEM)

            $firstName = if (-not [string]::IsNullOrWhiteSpace($rawFirstName)) { $rawFirstName.Trim() } else { "" }
            $lastName = if (-not [string]::IsNullOrWhiteSpace($rawLastName)) { $rawLastName.Trim() } else { "" }
            $typeValue = if (-not [string]::IsNullOrWhiteSpace($rawType)) { $rawType.Trim() } else { "" }
            $relatedItemValue = if (-not [string]::IsNullOrWhiteSpace($rawRelatedItem)) { $rawRelatedItem.Trim() } else { "" }

            $titleParts = @()
            if (-not [string]::IsNullOrWhiteSpace($firstName)) { $titleParts += $firstName }
            if (-not [string]::IsNullOrWhiteSpace($lastName)) { $titleParts += $lastName }
            $titleValue = ($titleParts -join " ").Trim()

            Write-Log ("Processing record {0}/{1} - Contact={2}" -f $recordCount, $csvData.Count, $titleValue)

            if ([string]::IsNullOrWhiteSpace($titleValue)) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - First Name and Last Name are both empty"
                continue
            }

            if ([string]::IsNullOrWhiteSpace($typeValue)) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Type field is empty"
                continue
            }

            $contentTypeKey = Get-NormalizedContentTypeKey -TypeValue $typeValue
            if ([string]::IsNullOrWhiteSpace($contentTypeKey)) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Invalid Type value: '$typeValue'"
                continue
            }

            $itemData = @{
                $SP_FIELDS.TITLE = $titleValue
                $SP_FIELDS.CONTENTTYPE_ID = $CONTENT_TYPES[$contentTypeKey]
            }

            if (-not [string]::IsNullOrWhiteSpace($firstName)) {
                $itemData[$SP_FIELDS.FIRST_NAME] = $firstName
            }

            if (-not [string]::IsNullOrWhiteSpace($lastName)) {
                $itemData[$SP_FIELDS.LAST_NAME] = $lastName
            }

            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.SUMMARY))) {
                $itemData[$SP_FIELDS.SUMMARY] = $record.($CSV_COLUMNS.SUMMARY).Trim()
            }

            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.EMAIL))) {
                $itemData[$SP_FIELDS.EMAIL] = $record.($CSV_COLUMNS.EMAIL).Trim()
            }

            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.PHONE_NUMBER))) {
                $itemData[$SP_FIELDS.PHONE_NUMBER] = $record.($CSV_COLUMNS.PHONE_NUMBER).Trim()
            }

            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.JOB_TITLE))) {
                $itemData[$SP_FIELDS.JOB_TITLE] = $record.($CSV_COLUMNS.JOB_TITLE).Trim()
            }

            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.DEPARTMENT))) {
                $itemData[$SP_FIELDS.DEPARTMENT] = $record.($CSV_COLUMNS.DEPARTMENT).Trim()
            }

            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.COMMENTS))) {
                $itemData[$SP_FIELDS.COMMENTS] = $record.($CSV_COLUMNS.COMMENTS).Trim()
            }

            $route = $RELATED_ITEM_ROUTE[$contentTypeKey]
            if (-not $route) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - No related item routing configuration found for Type '$contentTypeKey'"
                continue
            }

            if (-not [string]::IsNullOrWhiteSpace($route.ListName)) {
                if ([string]::IsNullOrWhiteSpace($relatedItemValue)) {
                    Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Related Item is required for Type '$contentTypeKey'"
                    continue
                }

                $lookupMap = $lookupMapsByList[$route.ListName]
                if (-not $lookupMap) {
                    Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Lookup map not available for list '$($route.ListName)'"
                    continue
                }

                $lookupKey = $relatedItemValue.Trim().ToLowerInvariant()
                if ($lookupMap.ContainsKey($lookupKey)) {
                    $itemData[$route.Field] = [int]$lookupMap[$lookupKey]
                }
                else {
                    Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Related Item '$relatedItemValue' not found in list '$($route.ListName)'"
                    continue
                }
            }
            else {
                if (-not [string]::IsNullOrWhiteSpace($relatedItemValue)) {
                    Write-Log ("Row {0} - Type '{1}' does not use Related Item. Value will be ignored." -f $recordCount, $contentTypeKey) "WARN"
                }
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
                            Write-Log "DRY RUN: Would create item for Contact: $($batchItem.Data[$SP_FIELDS.TITLE])"
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
                            Write-Log "Successfully added item ID $itemId for Contact: $($batchItem.Data[$SP_FIELDS.TITLE])" "SUCCESS"
                            $script:ImportedCount++
                        }
                    }
                    catch {
                        Add-RowError -RowNumber $batchItem.Index -Reason "Error adding item for Contact '$($batchItem.Data[$SP_FIELDS.TITLE])': $($_.Exception.Message)"
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
    Write-Log "Successfully Imported: $($script:ImportedCount)"
    Write-Log "Skipped: $($script:SkippedCount)"
    Write-Log "Errors: $($script:ErrorCount)"

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
        Write-Log "Dry Run Mode: No data was actually written to SharePoint"
    }
    else {
        Write-Log "Import completed successfully."
    }
}
catch {
    Write-Log "Fatal Error: $_" "ERROR"
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
