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
    TITLE                             = "Title"
    MARKET                            = "Market"
    RESPONSIBLE                       = "Responsible"
    RELATED_ITEM                      = "Related Item"
    EUTELSAT_OWNERS                   = "Eutelsat Owner(s)"
    MARKET_REDINESS_VERTICAL          = "Market Readiness Vertical(s)"

    VERTICAL_RAG_STATUS               = "Vertical RAG Status"
    VERTICAL_ESTIMATED_DATE           = "Vertical Estimated Date"
    VERTICAL_EFFECTIVE_DATE           = "Vertical Effective Date"
    VERTICAL_READINESS_COMMENTS       = "Vertical Readiness Comments"

    AVIATION_RAG_STATUS               = "Aviation RAG Status"
    AVIATION_ESTIMATED_DATE           = "Aviation Estimated Date"
    AVIATION_EFFECTIVE_DATE           = "Aviation Effective Date"
    AVIATION_READINESS_COMMENTS       = "Aviation Readiness Comments"

    LAND_FIXED_RAG_STATUS             = "Land Fixed RAG Status"
    LAND_FIXED_ESTIMATED_DATE         = "Land Fixed Estimated Date"
    LAND_FIXED_EFFECTIVE_DATE         = "Land Fixed Effective Date"
    LAND_FIXED_READINESS_COMMENTS     = "Land Fixed Readiness Comments"

    LAND_MOBILITY_RAG_STATUS          = "Land Mobility RAG Status"
    LAND_MOBILITY_ESTIMATED_DATE      = "Land Mobility Estimated Date"
    LAND_MOBILITY_EFFECTIVE_DATE      = "Land Mobility Effective Date"
    LAND_MOBILITY_READINESS_COMMENTS  = "Land Mobility Readiness Comments"

    MARITINE_RAG_STATUS               = "Maritime RAG Status"
    MARITINE_ESTIMATED_DATE           = "Maritime Estimated Date"
    MARITINE_EFFECTIVE_DATE           = "Maritime Effective Date"
    MARITINE_READINESS_COMMENTS       = "Maritime Readiness Comments"

    SPACE_NW_RAG_STATUS               = "Space & NW RAG Status"
    SPACE_NW_ESTIMATED_DATE           = "Space & NW Estimated Date"
    SPACE_NW_EFFECTIVE_DATE           = "Space & NW Effective Date"
    SPACE_NW_READINESS_COMMENTS       = "Space & NW Readiness Comments"
}

# ================================
# Configuration - SharePoint Field Internal Names
# ================================
$SP_FIELDS = @{
    TITLE                             = "Title"
    CONTENTTYPE_ID                    = "ContentTypeId"

    MARKET                            = "MA_Market"
    RESPONSIBLE                       = "MA_ReadinessResponsible"
    EUTELSAT_OWNERS                   = "MA_Owner"
    MARKET_REDINESS_VERTICAL          = "MA_RequirementVertical"

    VERTICAL_RAG_STATUS               = "MA_ReadinessRAGStatus"
    VERTICAL_ESTIMATED_DATE           = "MA_ReadinessEstimatedDate"
    VERTICAL_EFFECTIVE_DATE           = "MA_ReadinessEffectiveDate"
    VERTICAL_READINESS_COMMENTS       = "MA_ReadinessComments"

    AVIATION_RAG_STATUS               = "MA_AviationRAGStatus"
    AVIATION_ESTIMATED_DATE           = "MA_AviationEstimatedDate"
    AVIATION_EFFECTIVE_DATE           = "MA_AviationEffectiveDate"
    AVIATION_READINESS_COMMENTS       = "MA_AviationComments"

    LAND_FIXED_RAG_STATUS             = "MA_LandFixedRAGStatus"
    LAND_FIXED_ESTIMATED_DATE         = "MA_LandFixedEstimatedDate"
    LAND_FIXED_EFFECTIVE_DATE         = "MA_LandFixedEffectiveDate"
    LAND_FIXED_READINESS_COMMENTS     = "MA_LandFixedComments"

    LAND_MOBILITY_RAG_STATUS          = "MA_LandMobilityRAGStatus"
    LAND_MOBILITY_ESTIMATED_DATE      = "MA_LandMobilityEstimatedDate"
    LAND_MOBILITY_EFFECTIVE_DATE      = "MA_LandMobilityEffectiveDate"
    LAND_MOBILITY_READINESS_COMMENTS  = "MA_LandMobilityComments"

    MARITINE_RAG_STATUS               = "MA_MaritimeRAGStatus"
    MARITINE_ESTIMATED_DATE           = "MA_MaritimeEstimatedDate"
    MARITINE_EFFECTIVE_DATE           = "MA_MaritimeEffectiveDate"
    MARITINE_READINESS_COMMENTS       = "MA_MaritimeComments"

    SPACE_NW_RAG_STATUS               = "MA_SpaceRAGStatus"
    SPACE_NW_ESTIMATED_DATE           = "MA_SpaceEstimatedDate"
    SPACE_NW_EFFECTIVE_DATE           = "MA_SpaceEffectiveDate"
    SPACE_NW_READINESS_COMMENTS       = "MA_SpaceComments"

    DISTRIBUTION_PARTNER              = "MA_DistributionPartner"
    TELEPORT_PARTNER                  = "MA_TeleportPartner"
}

# ================================
# Configuration - SharePoint List Names
# ================================
$SP_LISTS = @{
    MARKET               = "Market"
    DISTRIBUTION_PARTNER = "Distribution Partner"
    CONTACT              = "Contact"
    TELEPORT_PARTNER     = "Teleport Partner"
}

# ================================
# Configuration - Content Type Mappings
# ================================
$CONTENT_TYPES = @{
    "ReadinessCountry" = "0x0100518D373795382B4BAC75508AC0E73D33"
}

# ================================
# Configuration - Choice Mappings
# ================================
$RESPONSIBLE_MAPPING = @{
    "overall"              = "Overall"
    "eutelsat"             = "Eutelsat"
    "teleport partner"     = "Teleport Partner"
    "distribution partner" = "Distribution Partner"
}

$MARKET_REDINESS_VERTICAL_MAPPING = @{
    "all verticals" = "All Verticals"
    "space & nw"    = "Space & NW"
    "land fixed"    = "Land Fixed"
    "land mobility" = "Land Mobility"
    "aviation"      = "Aviation"
    "maritime"      = "Maritime"
    "snp"           = "SNP"
}

$VERTICAL_RAG_STATUS_MAPPING = @{
    "major hurdles" = "Major hurdles"
    "in progress"   = "In Progress"
    "not started"   = "Not Started"
    "done"          = "Done"
    "sanctioned"    = "Sanctioned"
}

$AVIATION_RAG_STATUS_MAPPING = @{
    "major hurdles" = "Major hurdles"
    "in progress"   = "In Progress"
    "not started"   = "Not Started"
    "done"          = "Done"
    "sanctioned"    = "Sanctioned"
}

$LAND_FIXED_RAG_STATUS_MAPPING = @{
    "major hurdles" = "Major hurdles"
    "in progress"   = "In Progress"
    "not started"   = "Not Started"
    "done"          = "Done"
    "sanctioned"    = "Sanctioned"
}

$LAND_MOBILITY_RAG_STATUS_MAPPING = @{
    "major hurdles" = "Major hurdles"
    "in progress"   = "In Progress"
    "not started"   = "Not Started"
    "done"          = "Done"
    "sanctioned"    = "Sanctioned"
}

$MARITINE_RAG_STATUS_MAPPING = @{
    "major hurdles" = "Major hurdles"
    "in progress"   = "In Progress"
    "not started"   = "Not Started"
    "done"          = "Done"
    "sanctioned"    = "Sanctioned"
}

$SPACE_NW_RAG_STATUS_MAPPING = @{
    "major hurdles" = "Major hurdles"
    "in progress"   = "In Progress"
    "not started"   = "Not Started"
    "done"          = "Done"
    "sanctioned"    = "Sanctioned"
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

function Get-MappedMultiChoiceValues {
    param(
        [Parameter(Mandatory = $false)]
        [string]$RawValue,

        [Parameter(Mandatory = $true)]
        [hashtable]$Map,

        [Parameter(Mandatory = $true)]
        [string]$FieldLabel,

        [Parameter(Mandatory = $true)]
        [int]$RowNumber
    )

    if ([string]::IsNullOrWhiteSpace($RawValue)) {
        return @()
    }

    $tokens = $RawValue -split '[,;|]' |
        ForEach-Object { $_.Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique

    $mappedValues = @()
    $missingValues = @()

    foreach ($token in $tokens) {
        $mapped = Get-MappedChoiceValue -RawValue $token -Map $Map
        if ([string]::IsNullOrWhiteSpace($mapped)) {
            $missingValues += $token
        }
        else {
            $mappedValues += $mapped
        }
    }

    if ($missingValues.Count -gt 0) {
        Add-RowSkip -RowNumber $RowNumber -Reason "Skipped - Invalid $FieldLabel value(s): '$($missingValues -join ", ")'"
        return $null
    }

    return $mappedValues
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
            # Try next format.
        }
    }

    try {
        return ([DateTime]::Parse($normalized)).Date
    }
    catch {
        return $null
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

            Write-Log "Loaded $($lookupMap.Count) entries from lookup list '$ListName'."
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

function Resolve-MultiLookupIdsStrict {
    param(
        [Parameter(Mandatory = $false)]
        [string]$RawValue,

        [Parameter(Mandatory = $true)]
        [hashtable]$LookupMap,

        [Parameter(Mandatory = $true)]
        [string]$FieldLabel,

        [Parameter(Mandatory = $true)]
        [int]$RowNumber
    )

    if ([string]::IsNullOrWhiteSpace($RawValue)) {
        return @()
    }

    $tokens = $RawValue -split '[,;|]' |
        ForEach-Object { $_.Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique

    $ids = @()
    $missingValues = @()

    foreach ($token in $tokens) {
        $key = $token.ToLowerInvariant()
        if ($LookupMap.ContainsKey($key)) {
            $ids += [int]$LookupMap[$key]
        }
        else {
            $missingValues += $token
        }
    }

    if ($missingValues.Count -gt 0) {
        Add-RowSkip -RowNumber $RowNumber -Reason "Skipped - $FieldLabel lookup not found: '$($missingValues -join ", ")'"
        return $null
    }

    return $ids
}

function Set-DateFieldIfProvided {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$ItemData,

        [Parameter(Mandatory = $true)]
        [string]$FieldInternalName,

        [Parameter(Mandatory = $false)]
        [string]$RawValue,

        [Parameter(Mandatory = $true)]
        [string]$FieldLabel,

        [Parameter(Mandatory = $true)]
        [int]$RowNumber
    )

    if ([string]::IsNullOrWhiteSpace($RawValue)) {
        return $true
    }

    $dateValue = Get-DateOnlyValue -RawValue $RawValue
    if ($null -eq $dateValue) {
        Add-RowSkip -RowNumber $RowNumber -Reason "Skipped - Invalid $FieldLabel value: '$RawValue'"
        return $false
    }

    $ItemData[$FieldInternalName] = $dateValue.ToString("yyyy-MM-dd")
    return $true
}

function Set-ChoiceFieldIfProvided {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$ItemData,

        [Parameter(Mandatory = $true)]
        [string]$FieldInternalName,

        [Parameter(Mandatory = $false)]
        [string]$RawValue,

        [Parameter(Mandatory = $true)]
        [hashtable]$Map,

        [Parameter(Mandatory = $true)]
        [string]$FieldLabel,

        [Parameter(Mandatory = $true)]
        [int]$RowNumber
    )

    if ([string]::IsNullOrWhiteSpace($RawValue)) {
        return $true
    }

    $mappedValue = Get-MappedChoiceValue -RawValue $RawValue -Map $Map
    if ([string]::IsNullOrWhiteSpace($mappedValue)) {
        Add-RowSkip -RowNumber $RowNumber -Reason "Skipped - Invalid $FieldLabel value: '$RawValue'"
        return $false
    }

    $ItemData[$FieldInternalName] = $mappedValue
    return $true
}

function Set-CommentFieldIfProvided {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$ItemData,

        [Parameter(Mandatory = $true)]
        [string]$FieldInternalName,

        [Parameter(Mandatory = $false)]
        [string]$RawValue
    )

    if (-not [string]::IsNullOrWhiteSpace($RawValue)) {
        $ItemData[$FieldInternalName] = $RawValue.Trim()
    }
}

try {
    Write-Log "================================"
    Write-Log "MA Readiness List Data Import Started"
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
    $marketLookupMap = Get-LookupMap -ListName $SP_LISTS.MARKET
    $contactLookupMap = Get-LookupMap -ListName $SP_LISTS.CONTACT
    $distributionPartnerLookupMap = Get-LookupMap -ListName $SP_LISTS.DISTRIBUTION_PARTNER
    $teleportPartnerLookupMap = Get-LookupMap -ListName $SP_LISTS.TELEPORT_PARTNER

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
            $marketRaw = [string]$record.($CSV_COLUMNS.MARKET)
            $responsibleRaw = [string]$record.($CSV_COLUMNS.RESPONSIBLE)
            $relatedItemRaw = [string]$record.($CSV_COLUMNS.RELATED_ITEM)

            $titleValue = if (-not [string]::IsNullOrWhiteSpace($titleValue)) { $titleValue.Trim() } else { "" }
            $marketRaw = if (-not [string]::IsNullOrWhiteSpace($marketRaw)) { $marketRaw.Trim() } else { "" }
            $responsibleRaw = if (-not [string]::IsNullOrWhiteSpace($responsibleRaw)) { $responsibleRaw.Trim() } else { "" }
            $relatedItemRaw = if (-not [string]::IsNullOrWhiteSpace($relatedItemRaw)) { $relatedItemRaw.Trim() } else { "" }

            Write-Log ("Processing record {0}/{1} - Title={2}" -f $recordCount, $csvData.Count, $titleValue)

            if ([string]::IsNullOrWhiteSpace($titleValue)) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Title field is empty"
                continue
            }

            if ([string]::IsNullOrWhiteSpace($marketRaw)) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Market is required"
                continue
            }

            $marketKey = $marketRaw.ToLowerInvariant()
            if (-not $marketLookupMap.ContainsKey($marketKey)) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Market '$marketRaw' not found in list '$($SP_LISTS.MARKET)'"
                continue
            }

            $responsibleMapped = Get-MappedChoiceValue -RawValue $responsibleRaw -Map $RESPONSIBLE_MAPPING
            if ([string]::IsNullOrWhiteSpace($responsibleMapped)) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Invalid Responsible value: '$responsibleRaw'"
                continue
            }

            $itemData = @{
                $SP_FIELDS.TITLE          = $titleValue
                $SP_FIELDS.CONTENTTYPE_ID = $CONTENT_TYPES["ReadinessCountry"]
                $SP_FIELDS.MARKET         = [int]$marketLookupMap[$marketKey]
                $SP_FIELDS.RESPONSIBLE    = $responsibleMapped
            }

            $ownerRaw = [string]$record.($CSV_COLUMNS.EUTELSAT_OWNERS)
            if (-not [string]::IsNullOrWhiteSpace($ownerRaw)) {
                $ownerIds = Resolve-MultiLookupIdsStrict -RawValue $ownerRaw -LookupMap $contactLookupMap -FieldLabel "Eutelsat Owner(s)" -RowNumber $recordCount
                if ($null -eq $ownerIds) {
                    continue
                }

                if ($ownerIds.Count -gt 0) {
                    $itemData[$SP_FIELDS.EUTELSAT_OWNERS] = $ownerIds
                }
            }

            $verticalsRaw = [string]$record.($CSV_COLUMNS.MARKET_REDINESS_VERTICAL)
            if (-not [string]::IsNullOrWhiteSpace($verticalsRaw)) {
                $verticals = Get-MappedMultiChoiceValues -RawValue $verticalsRaw -Map $MARKET_REDINESS_VERTICAL_MAPPING -FieldLabel "Market Readiness Vertical(s)" -RowNumber $recordCount
                if ($null -eq $verticals) {
                    continue
                }

                if ($verticals.Count -gt 0) {
                    $itemData[$SP_FIELDS.MARKET_REDINESS_VERTICAL] = $verticals
                }
            }

            if (-not (Set-ChoiceFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.VERTICAL_RAG_STATUS -RawValue ([string]$record.($CSV_COLUMNS.VERTICAL_RAG_STATUS)) -Map $VERTICAL_RAG_STATUS_MAPPING -FieldLabel "Vertical RAG Status" -RowNumber $recordCount)) { continue }
            if (-not (Set-DateFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.VERTICAL_ESTIMATED_DATE -RawValue ([string]$record.($CSV_COLUMNS.VERTICAL_ESTIMATED_DATE)) -FieldLabel "Vertical Estimated Date" -RowNumber $recordCount)) { continue }
            if (-not (Set-DateFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.VERTICAL_EFFECTIVE_DATE -RawValue ([string]$record.($CSV_COLUMNS.VERTICAL_EFFECTIVE_DATE)) -FieldLabel "Vertical Effective Date" -RowNumber $recordCount)) { continue }
            Set-CommentFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.VERTICAL_READINESS_COMMENTS -RawValue ([string]$record.($CSV_COLUMNS.VERTICAL_READINESS_COMMENTS))

            if (-not (Set-ChoiceFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.AVIATION_RAG_STATUS -RawValue ([string]$record.($CSV_COLUMNS.AVIATION_RAG_STATUS)) -Map $AVIATION_RAG_STATUS_MAPPING -FieldLabel "Aviation RAG Status" -RowNumber $recordCount)) { continue }
            if (-not (Set-DateFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.AVIATION_ESTIMATED_DATE -RawValue ([string]$record.($CSV_COLUMNS.AVIATION_ESTIMATED_DATE)) -FieldLabel "Aviation Estimated Date" -RowNumber $recordCount)) { continue }
            if (-not (Set-DateFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.AVIATION_EFFECTIVE_DATE -RawValue ([string]$record.($CSV_COLUMNS.AVIATION_EFFECTIVE_DATE)) -FieldLabel "Aviation Effective Date" -RowNumber $recordCount)) { continue }
            Set-CommentFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.AVIATION_READINESS_COMMENTS -RawValue ([string]$record.($CSV_COLUMNS.AVIATION_READINESS_COMMENTS))

            if (-not (Set-ChoiceFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.LAND_FIXED_RAG_STATUS -RawValue ([string]$record.($CSV_COLUMNS.LAND_FIXED_RAG_STATUS)) -Map $LAND_FIXED_RAG_STATUS_MAPPING -FieldLabel "Land Fixed RAG Status" -RowNumber $recordCount)) { continue }
            if (-not (Set-DateFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.LAND_FIXED_ESTIMATED_DATE -RawValue ([string]$record.($CSV_COLUMNS.LAND_FIXED_ESTIMATED_DATE)) -FieldLabel "Land Fixed Estimated Date" -RowNumber $recordCount)) { continue }
            if (-not (Set-DateFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.LAND_FIXED_EFFECTIVE_DATE -RawValue ([string]$record.($CSV_COLUMNS.LAND_FIXED_EFFECTIVE_DATE)) -FieldLabel "Land Fixed Effective Date" -RowNumber $recordCount)) { continue }
            Set-CommentFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.LAND_FIXED_READINESS_COMMENTS -RawValue ([string]$record.($CSV_COLUMNS.LAND_FIXED_READINESS_COMMENTS))

            if (-not (Set-ChoiceFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.LAND_MOBILITY_RAG_STATUS -RawValue ([string]$record.($CSV_COLUMNS.LAND_MOBILITY_RAG_STATUS)) -Map $LAND_MOBILITY_RAG_STATUS_MAPPING -FieldLabel "Land Mobility RAG Status" -RowNumber $recordCount)) { continue }
            if (-not (Set-DateFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.LAND_MOBILITY_ESTIMATED_DATE -RawValue ([string]$record.($CSV_COLUMNS.LAND_MOBILITY_ESTIMATED_DATE)) -FieldLabel "Land Mobility Estimated Date" -RowNumber $recordCount)) { continue }
            if (-not (Set-DateFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.LAND_MOBILITY_EFFECTIVE_DATE -RawValue ([string]$record.($CSV_COLUMNS.LAND_MOBILITY_EFFECTIVE_DATE)) -FieldLabel "Land Mobility Effective Date" -RowNumber $recordCount)) { continue }
            Set-CommentFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.LAND_MOBILITY_READINESS_COMMENTS -RawValue ([string]$record.($CSV_COLUMNS.LAND_MOBILITY_READINESS_COMMENTS))

            if (-not (Set-ChoiceFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.MARITINE_RAG_STATUS -RawValue ([string]$record.($CSV_COLUMNS.MARITINE_RAG_STATUS)) -Map $MARITINE_RAG_STATUS_MAPPING -FieldLabel "Maritime RAG Status" -RowNumber $recordCount)) { continue }
            if (-not (Set-DateFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.MARITINE_ESTIMATED_DATE -RawValue ([string]$record.($CSV_COLUMNS.MARITINE_ESTIMATED_DATE)) -FieldLabel "Maritime Estimated Date" -RowNumber $recordCount)) { continue }
            if (-not (Set-DateFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.MARITINE_EFFECTIVE_DATE -RawValue ([string]$record.($CSV_COLUMNS.MARITINE_EFFECTIVE_DATE)) -FieldLabel "Maritime Effective Date" -RowNumber $recordCount)) { continue }
            Set-CommentFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.MARITINE_READINESS_COMMENTS -RawValue ([string]$record.($CSV_COLUMNS.MARITINE_READINESS_COMMENTS))

            if (-not (Set-ChoiceFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.SPACE_NW_RAG_STATUS -RawValue ([string]$record.($CSV_COLUMNS.SPACE_NW_RAG_STATUS)) -Map $SPACE_NW_RAG_STATUS_MAPPING -FieldLabel "Space & NW RAG Status" -RowNumber $recordCount)) { continue }
            if (-not (Set-DateFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.SPACE_NW_ESTIMATED_DATE -RawValue ([string]$record.($CSV_COLUMNS.SPACE_NW_ESTIMATED_DATE)) -FieldLabel "Space & NW Estimated Date" -RowNumber $recordCount)) { continue }
            if (-not (Set-DateFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.SPACE_NW_EFFECTIVE_DATE -RawValue ([string]$record.($CSV_COLUMNS.SPACE_NW_EFFECTIVE_DATE)) -FieldLabel "Space & NW Effective Date" -RowNumber $recordCount)) { continue }
            Set-CommentFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.SPACE_NW_READINESS_COMMENTS -RawValue ([string]$record.($CSV_COLUMNS.SPACE_NW_READINESS_COMMENTS))

            if ($responsibleMapped -eq "Teleport Partner") {
                if ([string]::IsNullOrWhiteSpace($relatedItemRaw)) {
                    Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Related Item is required when Responsible is 'Teleport Partner'"
                    continue
                }

                $relatedKey = $relatedItemRaw.ToLowerInvariant()
                if (-not $teleportPartnerLookupMap.ContainsKey($relatedKey)) {
                    Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Related Item '$relatedItemRaw' not found in list '$($SP_LISTS.TELEPORT_PARTNER)'"
                    continue
                }

                $itemData[$SP_FIELDS.TELEPORT_PARTNER] = [int]$teleportPartnerLookupMap[$relatedKey]
            }
            elseif ($responsibleMapped -eq "Distribution Partner") {
                if ([string]::IsNullOrWhiteSpace($relatedItemRaw)) {
                    Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Related Item is required when Responsible is 'Distribution Partner'"
                    continue
                }

                $relatedKey = $relatedItemRaw.ToLowerInvariant()
                if (-not $distributionPartnerLookupMap.ContainsKey($relatedKey)) {
                    Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Related Item '$relatedItemRaw' not found in list '$($SP_LISTS.DISTRIBUTION_PARTNER)'"
                    continue
                }

                $itemData[$SP_FIELDS.DISTRIBUTION_PARTNER] = [int]$distributionPartnerLookupMap[$relatedKey]
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
                            Write-Log "DRY RUN: Would create MA Readiness item for '$($batchItem.Title)'"
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
                            Write-Log "Successfully added item ID $itemId for '$($batchItem.Title)'" "SUCCESS"
                            $script:ImportedCount++
                        }
                    }
                    catch {
                        Add-RowError -RowNumber $batchItem.Index -Reason "Error adding item for '$($batchItem.Title)': $($_.Exception.Message)"
                    }
                }

                $batch = @()
            }
        }
        catch {
            Add-RowError -RowNumber $recordCount -Reason "Error processing record: $($_.Exception.Message)"
        }
    }

    if ($batch.Count -gt 0) {
        Write-Log "Processing final batch with $($batch.Count) items..."

        foreach ($batchItem in $batch) {
            try {
                if ($DryRun) {
                    Write-Log "DRY RUN: Would create MA Readiness item for '$($batchItem.Title)'"
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
                    Write-Log "Successfully added item ID $itemId for '$($batchItem.Title)'" "SUCCESS"
                    $script:ImportedCount++
                }
            }
            catch {
                Add-RowError -RowNumber $batchItem.Index -Reason "Error adding item for '$($batchItem.Title)': $($_.Exception.Message)"
            }
        }

        $batch = @()
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
        # Ignore disconnect errors.
    }

    Write-Log "Script execution completed."
}
