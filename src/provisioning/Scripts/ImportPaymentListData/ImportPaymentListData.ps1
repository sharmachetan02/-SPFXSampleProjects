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
    FEES              = "Fee(s)"
    CONTACTS          = "Contact(s)"
    SUMMARY           = "Summary"
    CHARGED_BY        = "Charged By"
    CURRENCY          = "Currency"
    RELATED_AUTHORITY = "Related Authority"
    DUE_DATE          = "Due Date"
    PAYMENT_DATE      = "Payment Date"
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
    TITLE                  = "Title"
    CONTENTTYPE_ID         = "ContentTypeId"
    FEES                   = "MA_Fee"
    CONTACTS               = "MA_Contact"
    SUMMARY                = "MA_Summary"
    CHARGED_BY             = "MA_FeeChargedBy"
    CURRENCY               = "MA_Currency"
    AUTHORITY              = "MA_Authority"
    LEGAL_SERVICE_PROVIDER = "MA_LegalSvcProvider"
    TELEPORT_PARTNER       = "MA_TeleportPartner"
    DUE_DATE               = "MA_DueDate"
    PAYMENT_DATE           = "MA_PaymentDate"
    VAT                    = "MA_Vat"
    VAT_FREE               = "MA_VatFree"
    VAT_RATE               = "MA_VatRate"
    STATUS                 = "MA_PaymentStatus"
    COMMENTS               = "MA_Comments"
    DOCUMENTS_SPACE        = "MA_DocumentsSpace"
}

# ================================
# Configuration - SharePoint List Names
# ================================
$SP_LISTS = @{
    FEE                    = "Fee"
    CURRENCY               = "Currency"
    AUTHORITY              = "Authority"
    CONTACT                = "Contact"
    LEGAL_SERVICE_PROVIDER = "Legal Service Provider"
    TELEPORT_PARTNER       = "Teleport Partner"
}

# ================================
# Configuration - Document Space Settings
# ================================
$DOCUMENT_SPACE = @{
    LIBRARY_NAME   = "documents_space"
    PAYMENT_FOLDER = "/Payment"
    HYPERLINK_DESC = "Link"
}

# ================================
# Configuration - Content Type Mappings
# ================================
$CONTENT_TYPES = @{
    "Payment" = "0x0100617446125FA54D4CAA177629C85FC7CD"
}

# ================================
# Configuration - Folder URL Map
# ================================
$FOLDER_URL_MAP = @{
    "Payment" = $DOCUMENT_SPACE.PAYMENT_FOLDER
}

# ================================
# Configuration - Choice Mappings
# ================================
$STATUS_MAPPING = @{
    "paid"             = "Paid"
    "partially paid"   = "Partially Paid"
    "overpaid"         = "Overpaid"
    "in process"       = "In Process"
    "cancelled"        = "Cancelled"
    "refunded"         = "Refunded"
    "rejected"         = "Rejected"
}

$CHARGED_BY_MAPPING = @{
    "regulator"            = "Regulator"
    "administration"       = "Administration"
    "organization"         = "Organization"
    "teleport partner"     = "Teleport Partner"
    "law firm"             = "Law Firm"
    "legal representative" = "Legal Representative"
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

function Get-MappedChoiceValue {
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
            return $Map[$key]
        }
    }

    return $null
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

function Resolve-MultiLookupIds {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RawValue,

        [Parameter(Mandatory = $true)]
        [hashtable]$LookupMap,

        [Parameter(Mandatory = $true)]
        [string]$FieldLabel,

        [Parameter(Mandatory = $true)]
        [int]$RowNumber
    )

    $ids = @()
    $tokens = $RawValue -split ',' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    foreach ($token in $tokens) {
        $key = $token.ToLowerInvariant()
        if ($LookupMap.ContainsKey($key)) {
            $ids += [int]$LookupMap[$key]
        }
        else {
            Write-Log ("Row {0}: {1} value '{2}' not found in lookup list. It will be skipped." -f $RowNumber, $FieldLabel, $token) "WARN"
        }
    }

    return ,$ids
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
    Write-Log "Payment List Data Import Started"
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
    $feeLookupMap = Get-LookupMap -ListName $SP_LISTS.FEE
    $contactLookupMap = Get-LookupMap -ListName $SP_LISTS.CONTACT
    $currencyLookupMap = Get-LookupMap -ListName $SP_LISTS.CURRENCY
    $authorityLookupMap = Get-LookupMap -ListName $SP_LISTS.AUTHORITY
    $legalServiceProviderLookupMap = Get-LookupMap -ListName $SP_LISTS.LEGAL_SERVICE_PROVIDER
    $teleportPartnerLookupMap = Get-LookupMap -ListName $SP_LISTS.TELEPORT_PARTNER

    $lookupMapsByList = @{
        $SP_LISTS.AUTHORITY              = $authorityLookupMap
        $SP_LISTS.LEGAL_SERVICE_PROVIDER = $legalServiceProviderLookupMap
        $SP_LISTS.TELEPORT_PARTNER       = $teleportPartnerLookupMap
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
            $chargedByValue = [string]$record.($CSV_COLUMNS.CHARGED_BY)
            $relatedAuthorityValue = [string]$record.($CSV_COLUMNS.RELATED_AUTHORITY)

            $titleValue = if (-not [string]::IsNullOrWhiteSpace($titleValue)) { $titleValue.Trim() } else { "" }
            $chargedByValue = if (-not [string]::IsNullOrWhiteSpace($chargedByValue)) { $chargedByValue.Trim() } else { "" }
            $relatedAuthorityValue = if (-not [string]::IsNullOrWhiteSpace($relatedAuthorityValue)) { $relatedAuthorityValue.Trim() } else { "" }

            Write-Log ("Processing record {0}/{1} - Payment={2}" -f $recordCount, $csvData.Count, $titleValue)

            if ([string]::IsNullOrWhiteSpace($titleValue)) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Title field is empty"
                continue
            }

            # ── Charged By (required Choice) ──────────────────────────────
            $chargedByMappedValue = Get-MappedChoiceValue -RawValue $chargedByValue -Map $CHARGED_BY_MAPPING
            if ([string]::IsNullOrWhiteSpace($chargedByMappedValue)) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Invalid Charged By value: '$chargedByValue'"
                continue
            }

            # ── Status (required Choice) ───────────────────────────────────
            $statusValue = Get-MappedChoiceValue -RawValue ([string]$record.($CSV_COLUMNS.STATUS)) -Map $STATUS_MAPPING
            if ([string]::IsNullOrWhiteSpace($statusValue)) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Invalid Status value: '$($record.($CSV_COLUMNS.STATUS))'"
                continue
            }

            # ── Due Date (required) ────────────────────────────────────────
            $dueDateValue = Get-DateOnlyValue -RawValue ([string]$record.($CSV_COLUMNS.DUE_DATE))
            if ($null -eq $dueDateValue) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Invalid Due Date value: '$($record.($CSV_COLUMNS.DUE_DATE))'"
                continue
            }

            # ── VAT figures (required Numbers) ────────────────────────────
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

            # ── Currency (required Lookup) ─────────────────────────────────
            $currencyRaw = [string]$record.($CSV_COLUMNS.CURRENCY)
            if ([string]::IsNullOrWhiteSpace($currencyRaw)) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Currency is required"
                continue
            }

            $currencyKey = $currencyRaw.Trim().ToLowerInvariant()
            if (-not $currencyLookupMap.ContainsKey($currencyKey)) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Currency '$currencyRaw' not found in list '$($SP_LISTS.CURRENCY)'"
                continue
            }

            # ── Charged By -> authority lookup routing ─────────────────────
            $authorityRoute = $null
            switch ($chargedByMappedValue) {
                { $_ -in @("Regulator", "Administration", "Organization") } {
                    $authorityRoute = @{
                        ListName = $SP_LISTS.AUTHORITY
                        Field    = $SP_FIELDS.AUTHORITY
                    }
                    break
                }
                { $_ -in @("Law Firm", "Legal Representative") } {
                    $authorityRoute = @{
                        ListName = $SP_LISTS.LEGAL_SERVICE_PROVIDER
                        Field    = $SP_FIELDS.LEGAL_SERVICE_PROVIDER
                    }
                    break
                }
                "Teleport Partner" {
                    $authorityRoute = @{
                        ListName = $SP_LISTS.TELEPORT_PARTNER
                        Field    = $SP_FIELDS.TELEPORT_PARTNER
                    }
                    break
                }
            }

            $resolvedAuthorityId = $null
            if ($authorityRoute) {
                if ([string]::IsNullOrWhiteSpace($relatedAuthorityValue)) {
                    Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Related Authority is required for Charged By '$chargedByMappedValue'"
                    continue
                }

                $authorityLookup = $lookupMapsByList[$authorityRoute.ListName]
                $authorityKey = $relatedAuthorityValue.Trim().ToLowerInvariant()
                if ($authorityLookup.ContainsKey($authorityKey)) {
                    $resolvedAuthorityId = @{
                        Field = $authorityRoute.Field
                        Id    = [int]$authorityLookup[$authorityKey]
                    }
                }
                else {
                    Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Related Authority '$relatedAuthorityValue' not found in list '$($authorityRoute.ListName)'"
                    continue
                }
            }

            # ── Build item data hashtable ──────────────────────────────────
            $itemData = @{
                $SP_FIELDS.TITLE          = $titleValue
                $SP_FIELDS.CONTENTTYPE_ID = $CONTENT_TYPES["Payment"]
                $SP_FIELDS.CHARGED_BY     = $chargedByMappedValue
                $SP_FIELDS.CURRENCY       = [int]$currencyLookupMap[$currencyKey]
                $SP_FIELDS.DUE_DATE       = $dueDateValue.ToString("yyyy-MM-dd")
                $SP_FIELDS.VAT            = [double]$vatValue
                $SP_FIELDS.VAT_FREE       = [double]$vatFreeValue
                $SP_FIELDS.VAT_RATE       = [double]$vatRateValue
                $SP_FIELDS.STATUS         = $statusValue
            }

            if ($null -ne $resolvedAuthorityId) {
                $itemData[$resolvedAuthorityId.Field] = $resolvedAuthorityId.Id
            }

            if (-not [string]::IsNullOrWhiteSpace([string]$record.($CSV_COLUMNS.SUMMARY))) {
                $itemData[$SP_FIELDS.SUMMARY] = $record.($CSV_COLUMNS.SUMMARY).Trim()
            }

            if (-not [string]::IsNullOrWhiteSpace([string]$record.($CSV_COLUMNS.COMMENTS))) {
                $itemData[$SP_FIELDS.COMMENTS] = $record.($CSV_COLUMNS.COMMENTS).Trim()
            }

            # ── Payment Date (optional Date) ───────────────────────────────
            $paymentDateRaw = [string]$record.($CSV_COLUMNS.PAYMENT_DATE)
            if (-not [string]::IsNullOrWhiteSpace($paymentDateRaw)) {
                $paymentDateValue = Get-DateOnlyValue -RawValue $paymentDateRaw
                if ($null -eq $paymentDateValue) {
                    Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Invalid Payment Date value: '$paymentDateRaw'"
                    continue
                }

                $itemData[$SP_FIELDS.PAYMENT_DATE] = $paymentDateValue.ToString("yyyy-MM-dd")
            }

            # ── Multi-select Fees lookup (Country-style strict resolution) ─
            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.FEES))) {
                $feesRawValue = $record.($CSV_COLUMNS.FEES).Trim()
                $feeValues = $feesRawValue -split '[,;|]' |
                    ForEach-Object { $_.Trim() } |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                    Select-Object -Unique

                if (-not $feeValues -or $feeValues.Count -eq 0) {
                    Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Fee(s) field is empty"
                    continue
                }

                $feeIds = @()
                $missingFees = @()

                foreach ($feeValue in $feeValues) {
                    $feeKey = $feeValue.ToLowerInvariant()
                    if ($feeLookupMap.ContainsKey($feeKey)) {
                        $feeIds += [int]$feeLookupMap[$feeKey]
                    }
                    else {
                        $missingFees += $feeValue
                    }
                }

                if ($missingFees.Count -gt 0) {
                    Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Fee lookup not found: '$($missingFees -join ", ")'"
                    continue
                }

                if ($feeIds.Count -gt 0) {
                    $itemData[$SP_FIELDS.FEES] = $feeIds
                }
                else {
                    Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Fee lookup resolution returned no IDs"
                    continue
                }
            }
            else {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Fee(s) field is empty"
                continue
            }

            # ── Multi-select Contacts lookup (Country-style strict resolution) ─
            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.CONTACTS))) {
                $contactsRawValue = $record.($CSV_COLUMNS.CONTACTS).Trim()
                $contactValues = $contactsRawValue -split '[,;|]' |
                    ForEach-Object { $_.Trim() } |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                    Select-Object -Unique

                if (-not $contactValues -or $contactValues.Count -eq 0) {
                    Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Contact(s) field is empty"
                    continue
                }

                $contactIds = @()
                $missingContacts = @()

                foreach ($contactValue in $contactValues) {
                    $contactKey = $contactValue.ToLowerInvariant()
                    if ($contactLookupMap.ContainsKey($contactKey)) {
                        $contactIds += [int]$contactLookupMap[$contactKey]
                    }
                    else {
                        $missingContacts += $contactValue
                    }
                }

                if ($missingContacts.Count -gt 0) {
                    Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Contact lookup not found: '$($missingContacts -join ", ")'"
                    continue
                }

                if ($contactIds.Count -gt 0) {
                    $itemData[$SP_FIELDS.CONTACTS] = $contactIds
                }
                else {
                    Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Contact lookup resolution returned no IDs"
                    continue
                }
            }
            else {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Contact(s) field is empty"
                continue
            }

            $batch += [PSCustomObject]@{
                Data  = $itemData
                Index = $recordCount
                Title = $titleValue
            }

            if ($batch.Count -ge $BatchSize -or $recordCount -eq $csvData.Count) {
                Write-Log "Processing batch with $($batch.Count) items..."

                foreach ($batchItem in $batch) {
                    try {
                        if ($DryRun) {
                            Write-Log "DRY RUN: Would create item for Payment: $($batchItem.Title)"
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
                                $sanitizedFolderName = Get-SanitizedFolderName -FolderName $batchItem.Title
                                $parentFolderUrl = "$($DOCUMENT_SPACE.LIBRARY_NAME)$($DOCUMENT_SPACE.PAYMENT_FOLDER)"
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
                                Write-Log ("Warning: Failed to create/update Documents Space for Payment item ID {0}: {1}" -f $itemId, $_.Exception.Message) "WARN"
                            }

                            Write-Log "Successfully added item ID $itemId for Payment: $($batchItem.Title)" "SUCCESS"
                            $script:ImportedCount++
                        }
                    }
                    catch {
                        Add-RowError -RowNumber $batchItem.Index -Reason "Error adding item for Payment '$($batchItem.Title)': $($_.Exception.Message)"
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
