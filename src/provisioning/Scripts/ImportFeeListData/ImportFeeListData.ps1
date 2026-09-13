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
    TITLE             = "Title"
    TYPE              = "Type"
    COST_TYPE         = "Cost Type"
    CATEGORY          = "Category"
    SUMMARY           = "Summary"
    ISSUER_TYPE       = "Issuer Type"
    CURRENCY          = "Currency"
    PO_TYPE           = "Po Type"
    RELATED_ITEM      = "Related Item"
    ISSUER            = "Issuer"
    DUE_DATE          = "Due Date"
    REPEATING_FEE     = "Repeating Fee"
    VAT               = "VAT"
    VAT_FREE          = "VAT Free"
    VAT_RATE          = "VAT Rate"
    STATUS            = "Status"
    COMMENTS          = "Comments"
}

# ================================
# Configuration - SharePoint Field Internal Names
# ================================
$SP_FIELDS = @{
    TITLE                    = "Title"
    CONTENTTYPE_ID           = "ContentTypeId"
    CLASSIFICATION           = "MA_CostClassification"
    FEE_TYPE                 = "MA_FeeType"
    SUMMARY                  = "MA_Summary"
    FEE_CHARGED_BY           = "MA_FeeChargedBy"
    CURRENCY                 = "MA_Currency"
    PO_TYPE                  = "MA_PoType"
    AUTHORITY                = "MA_Authority"
    LEGAL_SERVICE_PROVIDER   = "MA_LegalSvcProvider"
    TELEPORT_PARTNER         = "MA_TeleportPartner"
    MA_REQUIREMENT           = "MA_Requirement"
    SATELLITE_NETWORK_PORTAL = "MA_SNP"
    DUE_DATE                 = "MA_DueDate"
    REPEATING_FEE            = "MA_IsRepeatingFee"
    RECURRENCE_PATTERN       = "MA_RecurrencePattern"
    VAT                      = "MA_Vat"
    VAT_FREE                 = "MA_VatFree"
    VAT_RATE                 = "MA_VatRate"
    STATUS                   = "MA_FeeStatus"
    COMMENTS                 = "MA_Comments"
    DOCUMENTS_SPACE          = "MA_DocumentsSpace"
}

# ================================
# Configuration - SharePoint List Names
# ================================
$SP_LISTS = @{
    FEE                      = "Fee"
    CURRENCY                 = "Currency"
    AUTHORITY                = "Authority"
    MA_REQUIREMENT           = "MA Requirement"
    LEGAL_SERVICE_PROVIDER   = "Legal Service Provider"
    TELEPORT_PARTNER         = "Teleport Partner"
    SATELLITE_NETWORK_PORTAL = "Satellite Network Portal"
    RECURRENCE_PATTERN       = "Recurrence Pattern"
}

# ================================
# Configuration - Document Space Settings
# ================================
$DOCUMENT_SPACE = @{
    LIBRARY_NAME          = "documents_space"
    MA_REQUIREMENTS_FOLDER = "/Fees/MA Requirements"
    SNP_FOLDER             = "/Fees/SNPs"
    HYPERLINK_DESC        = "Link"
}

# ================================
# Configuration - Content Type Mappings
# ================================
$CONTENT_TYPES = @{
    "SNP"            = "0x010092D96D11717F014497021E5516FB818201"
    "MA Requirement" = "0x010092D96D11717F014497021E5516FB818202"
}

# ================================
# Configuration - Folder URL Map with type
# ================================
$FOLDER_URL_MAP = @{
    "SNP"            = $DOCUMENT_SPACE.SNP_FOLDER
    "MA Requirement" = $DOCUMENT_SPACE.MA_REQUIREMENTS_FOLDER
}

$STATUS_MAPPING = @{
    "active"   = "Active"
    "inactive" = "Inactive"
    "revoked"  = "Revoked"
}

$CLASSIFICATION_MAPPING = @{
    "capex" = "CAPEX"
    "opex"  = "OPEX"
}

$FEE_CHARGED_BY_MAPPING = @{
    "regulator"            = "Regulator"
    "administration"       = "Administration"
    "organization"         = "Organization"
    "teleport partner"     = "Teleport Partner"
    "law firm"             = "Law Firm"
    "legal representative" = "Legal Representative"
}

$FEE_TYPE_MAPPING = @{
    "commercial"     = "Commercial"
    "license"        = "License"
    "administrative" = "Administrative"
}

$PO_TYPE_MAPPING = @{
    "po"     = "PO"
    "non-po" = "Non-PO"
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

function Get-NormalizedKey {
    param(
        [Parameter(Mandatory = $false)]
        [string]$RawValue,

        [Parameter(Mandatory = $true)]
        [hashtable]$Map
    )

    if ([string]::IsNullOrWhiteSpace($RawValue)) {
        return $null
    }

    $normalizedInput = $RawValue.Trim().ToLowerInvariant()
    foreach ($key in $Map.Keys) {
        if ($key.ToLowerInvariant() -eq $normalizedInput) {
            return $key
        }
    }

    return $null
}

function Get-MappedChoiceValue {
    param(
        [Parameter(Mandatory = $false)]
        [string]$RawValue,

        [Parameter(Mandatory = $true)]
        [hashtable]$Map
    )

    $normalizedKey = Get-NormalizedKey -RawValue $RawValue -Map $Map
    if ([string]::IsNullOrWhiteSpace($normalizedKey)) {
        return $null
    }

    return $Map[$normalizedKey]
}

function Get-DateOnlyValue {
    param(
        [Parameter(Mandatory = $false)]
        [string]$RawValue
    )

    if ([string]::IsNullOrWhiteSpace($RawValue)) {
        return $null
    }

    $normalized = $RawValue.Trim()
    $formats = @("yyyy-MM-dd", "dd-MM-yyyy", "MM-dd-yyyy", "dd/MM/yyyy", "MM/dd/yyyy")

    foreach ($format in $formats) {
        try {
            return [DateTime]::ParseExact($normalized, $format, [System.Globalization.CultureInfo]::InvariantCulture).Date
        }
        catch {
            # Try next format
        }
    }

    try {
        return ([DateTime]::Parse($normalized)).Date
    }
    catch {
        return $null
    }
}

function Get-DecimalValue {
    param(
        [Parameter(Mandatory = $false)]
        [string]$RawValue
    )

    if ([string]::IsNullOrWhiteSpace($RawValue)) {
        return $null
    }

    $value = 0.0
    if ([double]::TryParse($RawValue.Trim(), [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$value)) {
        return [double]$value
    }

    return $null
}

function Get-BoolValue {
    param(
        [Parameter(Mandatory = $false)]
        [string]$RawValue
    )

    if ([string]::IsNullOrWhiteSpace($RawValue)) {
        return $null
    }

    switch -Regex ($RawValue.Trim().ToLowerInvariant()) {
        "^(true|yes|y|1)$" { return $true }
        "^(false|no|n|0)$" { return $false }
        default { return $null }
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

function Find-RecurrencePatternLookupId {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FeeTitle,

        [Parameter(Mandatory = $true)]
        [hashtable]$RecurrenceLookupMap
    )

    $normalizedTitle = $FeeTitle.Trim().ToLowerInvariant()
    $exactPatternKey = "$normalizedTitle`_pattern"

    if ($RecurrenceLookupMap.ContainsKey($exactPatternKey)) {
        return [int]$RecurrenceLookupMap[$exactPatternKey]
    }

    # $containsMatches = @($RecurrenceLookupMap.Keys | Where-Object { $_ -like "*$normalizedTitle*" })
    # if ($containsMatches.Count -gt 0) {
    #     $matchedKey = $containsMatches[0]
    #     Write-Log "Recurrence Pattern exact match not found for '$FeeTitle'. Using contains match '$matchedKey'." "WARN"
    #     return [int]$RecurrenceLookupMap[$matchedKey]
    # }

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

try {
    Write-Log "================================"
    Write-Log "Fee List Data Import Started"
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
    $currencyLookupMap = Get-LookupMap -ListName $SP_LISTS.CURRENCY
    $authorityLookupMap = Get-LookupMap -ListName $SP_LISTS.AUTHORITY
    $legalServiceProviderLookupMap = Get-LookupMap -ListName $SP_LISTS.LEGAL_SERVICE_PROVIDER
    $teleportPartnerLookupMap = Get-LookupMap -ListName $SP_LISTS.TELEPORT_PARTNER
    $maRequirementLookupMap = Get-LookupMap -ListName $SP_LISTS.MA_REQUIREMENT
    $snpLookupMap = Get-LookupMap -ListName $SP_LISTS.SATELLITE_NETWORK_PORTAL
    $recurrencePatternLookupMap = Get-LookupMap -ListName $SP_LISTS.RECURRENCE_PATTERN

    $lookupMapsByList = @{
        $SP_LISTS.CURRENCY = $currencyLookupMap
        $SP_LISTS.AUTHORITY = $authorityLookupMap
        $SP_LISTS.LEGAL_SERVICE_PROVIDER = $legalServiceProviderLookupMap
        $SP_LISTS.TELEPORT_PARTNER = $teleportPartnerLookupMap
        $SP_LISTS.MA_REQUIREMENT = $maRequirementLookupMap
        $SP_LISTS.SATELLITE_NETWORK_PORTAL = $snpLookupMap
        $SP_LISTS.RECURRENCE_PATTERN = $recurrencePatternLookupMap
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
            $titleValue = [string]$record.($CSV_COLUMNS.TITLE)
            $typeValue = [string]$record.($CSV_COLUMNS.TYPE)
            $feeChargedByValue = [string]$record.($CSV_COLUMNS.ISSUER_TYPE)
            $relatedAuthorityValue = [string]$record.($CSV_COLUMNS.ISSUER)
            $relatedItemValue = [string]$record.($CSV_COLUMNS.RELATED_ITEM)

            $titleValue = if (-not [string]::IsNullOrWhiteSpace($titleValue)) { $titleValue.Trim() } else { "" }
            $typeValue = if (-not [string]::IsNullOrWhiteSpace($typeValue)) { $typeValue.Trim() } else { "" }
            $feeChargedByValue = if (-not [string]::IsNullOrWhiteSpace($feeChargedByValue)) { $feeChargedByValue.Trim() } else { "" }
            $relatedAuthorityValue = if (-not [string]::IsNullOrWhiteSpace($relatedAuthorityValue)) { $relatedAuthorityValue.Trim() } else { "" }
            $relatedItemValue = if (-not [string]::IsNullOrWhiteSpace($relatedItemValue)) { $relatedItemValue.Trim() } else { "" }

            Write-Log ("Processing record {0}/{1} - Fee={2}" -f $recordCount, $csvData.Count, $titleValue)

            if ([string]::IsNullOrWhiteSpace($titleValue)) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Title field is empty"
                continue
            }

            $typeKey = Get-NormalizedKey -RawValue $typeValue -Map $CONTENT_TYPES
            if ([string]::IsNullOrWhiteSpace($typeKey)) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Invalid Type value: '$typeValue'"
                continue
            }

            $classificationValue = Get-MappedChoiceValue -RawValue ([string]$record.($CSV_COLUMNS.COST_TYPE)) -Map $CLASSIFICATION_MAPPING
            if ([string]::IsNullOrWhiteSpace($classificationValue)) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Invalid Classification value: '$($record.($CSV_COLUMNS.COST_TYPE))'"
                continue
            }

            $feeTypeValue = Get-MappedChoiceValue -RawValue ([string]$record.($CSV_COLUMNS.CATeGORY)) -Map $FEE_TYPE_MAPPING
            if ([string]::IsNullOrWhiteSpace($feeTypeValue)) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Invalid Fee Type value: '$($record.($CSV_COLUMNS.CATeGORY))'"
                continue
            }

            $feeChargedByMappedValue = Get-MappedChoiceValue -RawValue $feeChargedByValue -Map $FEE_CHARGED_BY_MAPPING
            if ([string]::IsNullOrWhiteSpace($feeChargedByMappedValue)) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Invalid Fee Charged By value: '$feeChargedByValue'"
                continue
            }

            $poTypeValue = Get-MappedChoiceValue -RawValue ([string]$record.($CSV_COLUMNS.PO_TYPE)) -Map $PO_TYPE_MAPPING
            if ([string]::IsNullOrWhiteSpace($poTypeValue)) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Invalid Po Type value: '$($record.($CSV_COLUMNS.PO_TYPE))'"
                continue
            }

            $statusValue = Get-MappedChoiceValue -RawValue ([string]$record.($CSV_COLUMNS.STATUS)) -Map $STATUS_MAPPING
            if ([string]::IsNullOrWhiteSpace($statusValue)) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Invalid Status value: '$($record.($CSV_COLUMNS.STATUS))'"
                continue
            }

            $dueDateValue = Get-DateOnlyValue -RawValue ([string]$record.($CSV_COLUMNS.DUE_DATE))
            if ($null -eq $dueDateValue) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Invalid Due Date value: '$($record.($CSV_COLUMNS.DUE_DATE))'"
                continue
            }

            $isRepeatingFeeValue = Get-BoolValue -RawValue ([string]$record.($CSV_COLUMNS.REPEATING_FEE))
            if ($null -eq $isRepeatingFeeValue) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Invalid Repeating Fee value: '$($record.($CSV_COLUMNS.REPEATING_FEE))'. Use Yes/No or True/False"
                continue
            }

            $vatValue = Get-DecimalValue -RawValue ([string]$record.($CSV_COLUMNS.VAT))
            if ($null -eq $vatValue) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Invalid VAT value: '$($record.($CSV_COLUMNS.VAT))'"
                continue
            }

            $vatFreeValue = Get-DecimalValue -RawValue ([string]$record.($CSV_COLUMNS.VAT_FREE))
            if ($null -eq $vatFreeValue) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Invalid VAT Free value: '$($record.($CSV_COLUMNS.VAT_FREE))'"
                continue
            }

            $vatRateValue = Get-DecimalValue -RawValue ([string]$record.($CSV_COLUMNS.VAT_RATE))
            if ($null -eq $vatRateValue) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Invalid VAT Rate value: '$($record.($CSV_COLUMNS.VAT_RATE))'"
                continue
            }

            $itemData = @{
                $SP_FIELDS.TITLE          = $titleValue
                $SP_FIELDS.CONTENTTYPE_ID = $CONTENT_TYPES[$typeKey]
                $SP_FIELDS.CLASSIFICATION = $classificationValue
                $SP_FIELDS.FEE_TYPE       = $feeTypeValue
                $SP_FIELDS.FEE_CHARGED_BY = $feeChargedByMappedValue
                $SP_FIELDS.PO_TYPE        = $poTypeValue
                $SP_FIELDS.DUE_DATE       = $dueDateValue.ToString("yyyy-MM-dd")
                $SP_FIELDS.REPEATING_FEE  = [bool]$isRepeatingFeeValue
                $SP_FIELDS.VAT            = [double]$vatValue
                $SP_FIELDS.VAT_FREE       = [double]$vatFreeValue
                $SP_FIELDS.VAT_RATE       = [double]$vatRateValue
                $SP_FIELDS.STATUS         = $statusValue
            }

            if (-not [string]::IsNullOrWhiteSpace([string]$record.($CSV_COLUMNS.SUMMARY))) {
                $itemData[$SP_FIELDS.SUMMARY] = $record.($CSV_COLUMNS.SUMMARY).Trim()
            }

            if (-not [string]::IsNullOrWhiteSpace([string]$record.($CSV_COLUMNS.COMMENTS))) {
                $itemData[$SP_FIELDS.COMMENTS] = $record.($CSV_COLUMNS.COMMENTS).Trim()
            }

            $currencyValue = [string]$record.($CSV_COLUMNS.CURRENCY)
            if (-not [string]::IsNullOrWhiteSpace($currencyValue)) {
                $currencyKey = $currencyValue.Trim().ToLowerInvariant()
                if ($currencyLookupMap.ContainsKey($currencyKey)) {
                    $itemData[$SP_FIELDS.CURRENCY] = [int]$currencyLookupMap[$currencyKey]
                }
                else {
                    Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Currency '$currencyValue' not found in list '$($SP_LISTS.CURRENCY)'"
                    continue
                }
            }
            else {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Currency is required"
                continue
            }

            $authorityRoute = $null
            switch ($feeChargedByMappedValue) {
                { $_ -in @("Regulator", "Administration", "Organization") } {
                    $authorityRoute = @{
                        ListName = $SP_LISTS.AUTHORITY
                        Field = $SP_FIELDS.AUTHORITY
                    }
                    break
                }
                { $_ -in @("Law Firm", "Legal Representative") } {
                    $authorityRoute = @{
                        ListName = $SP_LISTS.LEGAL_SERVICE_PROVIDER
                        Field = $SP_FIELDS.LEGAL_SERVICE_PROVIDER
                    }
                    break
                }
                "Teleport Partner" {
                    $authorityRoute = @{
                        ListName = $SP_LISTS.TELEPORT_PARTNER
                        Field = $SP_FIELDS.TELEPORT_PARTNER
                    }
                    break
                }
            }

            if ($authorityRoute) {
                if ([string]::IsNullOrWhiteSpace($relatedAuthorityValue)) {
                    Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Related Authority is required for Fee Charged By '$feeChargedByMappedValue'"
                    continue
                }

                $authorityLookupMap = $lookupMapsByList[$authorityRoute.ListName]
                $authorityKey = $relatedAuthorityValue.Trim().ToLowerInvariant()
                if ($authorityLookupMap.ContainsKey($authorityKey)) {
                    $itemData[$authorityRoute.Field] = [int]$authorityLookupMap[$authorityKey]
                }
                else {
                    Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Related Authority '$relatedAuthorityValue' not found in list '$($authorityRoute.ListName)'"
                    continue
                }
            }

            $typeRoute = $null
            switch ($typeKey) {
                "SNP" {
                    $typeRoute = @{
                        ListName = $SP_LISTS.SATELLITE_NETWORK_PORTAL
                        Field = $SP_FIELDS.SATELLITE_NETWORK_PORTAL
                    }
                    break
                }
                "MA Requirement" {
                    $typeRoute = @{
                        ListName = $SP_LISTS.MA_REQUIREMENT
                        Field = $SP_FIELDS.MA_REQUIREMENT
                    }
                    break
                }
            }

            if ($typeRoute) {
                if ([string]::IsNullOrWhiteSpace($relatedItemValue)) {
                    Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Related Item is required for Type '$typeKey'"
                    continue
                }

                $relatedItemLookupMap = $lookupMapsByList[$typeRoute.ListName]
                $relatedItemKey = $relatedItemValue.Trim().ToLowerInvariant()
                if ($relatedItemLookupMap.ContainsKey($relatedItemKey)) {
                    $itemData[$typeRoute.Field] = [int]$relatedItemLookupMap[$relatedItemKey]
                }
                else {
                    Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Related Item '$relatedItemValue' not found in list '$($typeRoute.ListName)'"
                    continue
                }
            }

            if ($isRepeatingFeeValue) {
                $recurrenceId = Find-RecurrencePatternLookupId -FeeTitle $titleValue -RecurrenceLookupMap $recurrencePatternLookupMap
                if ($null -eq $recurrenceId) {
                    Add-RowSkip -RowNumber $recordCount -Reason "Skipped - No Recurrence Pattern found for repeating fee '$titleValue'"
                    continue
                }

                $itemData[$SP_FIELDS.RECURRENCE_PATTERN] = [int]$recurrenceId
            }
            elseif ($itemData.ContainsKey($SP_FIELDS.RECURRENCE_PATTERN)) {
                $itemData.Remove($SP_FIELDS.RECURRENCE_PATTERN) | Out-Null
            }

            $batch += [PSCustomObject]@{
                Data   = $itemData
                Index  = $recordCount
                Type   = $typeKey
                Title  = $titleValue
            }

            if ($batch.Count -ge $BatchSize -or $recordCount -eq $csvData.Count) {
                Write-Log "Processing batch with $($batch.Count) items..."

                foreach ($batchItem in $batch) {
                    try {
                        if ($DryRun) {
                            Write-Log "DRY RUN: Would create item for Fee: $($batchItem.Title)"
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
                                $folderPath = $FOLDER_URL_MAP[$batchItem.Type]
                                if ([string]::IsNullOrWhiteSpace($folderPath)) {
                                    $folderPath = $DOCUMENT_SPACE.MA_REQUIREMENTS_FOLDER
                                    Write-Log "No folder mapping found for type '$($batchItem.Type)'. Using default '$folderPath'." "WARN"
                                }

                                $sanitizedFolderName = Get-SanitizedFolderName -FolderName $batchItem.Title
                                $parentFolderUrl = "$($DOCUMENT_SPACE.LIBRARY_NAME)$folderPath"
                                $folderUrl = New-UniqueFolder -ParentFolderUrl $parentFolderUrl -BaseName $sanitizedFolderName

                                $web = Get-PnPWeb
                                $siteRelativeFolderUrl = "$($web.ServerRelativeUrl.TrimEnd('/'))/$folderUrl"
                                $hyperlinkValue = "$siteRelativeFolderUrl, $($DOCUMENT_SPACE.HYPERLINK_DESC)"

                                Set-PnPListItem -List $ListName -Identity $itemId -Values @{
                                    $SP_FIELDS.DOCUMENTS_SPACE = $hyperlinkValue
                                } -ErrorAction Stop

                                Write-Log "Created document space folder and updated field for item ID $itemId" "INFO"
                            }
                            catch {
                                Write-Log ("Warning: Failed to create/update Documents Space for item ID {0}: {1}" -f $itemId, $_.Exception.Message) "WARN"
                            }

                            Write-Log "Successfully added item ID $itemId for Fee: $($batchItem.Title)" "SUCCESS"
                            $script:ImportedCount++
                        }
                    }
                    catch {
                        Add-RowError -RowNumber $batchItem.Index -Reason "Error adding item for Fee '$($batchItem.Title)': $($_.Exception.Message)"
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
