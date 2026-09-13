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
    TITLE                    = "Title"
    SUMMARY                  = "Summary"
    MARKET                   = "Market"
    REQUIREMENT_TYPE         = "Requirement Type"
    VERTICAL                 = "Vertical(s)"
    SYNTHESIS_STATUS         = "Synthesis Status"
    RAG_STATUS               = "RAG Status"
    RESPONSIBLE              = "Responsible"
    RESPONSIBLE_CONTACTS     = "Responsible Contact(s)"
    RELATED_ITEM             = "Related Item"
    EUTELSAT_OWNERS          = "Eutelsat Owner(s)"
    COMMETNS                 = "Comments"
    SATELLITE_NETWORK_PORTAL = "Satellite Network Portal(s)"
    SATELLITE_ORBIT          = "Satellite Orbit"
    REQUESTED_PARTY_TYPE     = "Requested Party Type"
    REQUESTED_PARTY          = "Requested Party"
    REQUESTED_PARTY_CONTACTS = "Requested Party Contact(s)"
    LETTER_1_MILESTONE       = "Letter 1 Milestone"
    LETTER_1_SENT_DATE       = "Letter 1 Sent Date"
    LETTER_2_MILESTONE       = "Letter 2 Milestone"
    LETTER_2_SENT_DATE       = "Letter 2 Sent Date"
    RESPONSE_DATE            = "Response Date"
    RESPONSE_SUMMARY         = "Response Summary"
}

# ================================
# Configuration - SharePoint Field Internal Names
# ================================
$SP_FIELDS = @{
    TITLE                    = "Title"
    CONTENTTYPE_ID           = "ContentTypeId"
    SUMMARY                  = "MA_Summary"
    MARKET                   = "MA_Market"
    REQUIREMENT_TYPE         = "MA_RequirementType"
    VERTICAL                 = "MA_RequirementVertical"
    SYNTHESIS_STATUS         = "MA_RequirementStatus"
    RAG_STATUS               = "MA_RAGStatus"
    RESPONSIBLE              = "MA_RequirementResponsible"
    RESPONSIBLE_CONTACTS     = "MA_ResponsiblePartyContact"
    TELEPORT_PARTNER         = "MA_TeleportPartner"
    DISTRIBUTION_PARTNER     = "MA_DistributionPartner"
    EUTELSAT_ENTITY          = "MA_EutelsatEntity"
    EUTELSAT_OWNERS          = "MA_Owner"
    COMMETNS                 = "MA_Comments"
    SATELLITE_ORBIT          = "MA_SatelliteOrbit"
    SATELLITE_NETWORK_PORTAL = "MA_SNP"
    REQUESTED_PARTY_TYPE     = "MA_RequirementRequestedPartyType"
    REQUESTED_PARTY          = "MA_Authority"
    REQUESTED_PARTY_CONTACTS = "MA_RequestedPartyContact"
    LETTER_1_MILESTONE       = "MA_Letter1Milestone"
    LETTER_1_SENT_DATE       = "MA_Letter1SentDate"
    LETTER_2_MILESTONE       = "MA_Letter2Milestone"
    LETTER_2_SENT_DATE       = "MA_Letter2SentDate"
    RESPONSE_DATE            = "MA_ResponseDate"
    RESPONSE_SUMMARY         = "MA_ResponseSummary"
    DOCUMENTS_SPACE          = "MA_DocumentsSpace"
}

# ================================
# Configuration - SharePoint List Names
# ================================
$SP_LISTS = @{
    MARKET                   = "Market"
    REQUIREMENT_TYPE         = "Requirement Type"
    REQUIREMENT_STATUS       = "Requirement Status"   
    CONTACT                  = "Contact"
    AUTHORITY                = "Authority"
    TELEPORT_PARTNER         = "Teleport Partner"
    DISTRIBUTION_PARTNER     = "Distribution Partner"
    EUTELSAT_ENTITY          = "Eutelsat Entity"
    SATELLITE_NETWORK_PORTAL = "Satellite Network Portal"
}

# ================================
# Configuration - Document Space Settings
# ================================
$DOCUMENT_SPACE = @{
    LIBRARY_NAME      = "documents_space"
    NOL_FOLDER        = "/MA Requirements/Nol"
    HYPERLINK_DESC    = "Link"
}

# ================================
# Configuration - Content Type Mappings
# ================================
$CONTENT_TYPES = @{
    "Nol" = "0x01009E58F9850C7D4144AD7FB9331AD6A24803"
}

# ================================
# Configuration - Folder URL Map with type
# ================================
$FOLDER_URL_MAP = @{
    "Nol" = $DOCUMENT_SPACE.NOL_FOLDER
}

$RAG_STATUS_MAPPING = @{
    "major hurdles" = "Major hurdles"
    "in progress"   = "In Progress"
    "not started"   = "Not Started"
    "done"          = "Done"
    "sanctioned"    = "Sanctioned"
}

$RESPONSIBLE_MAPPING = @{
    "eutelsat"             = "Eutelsat"
    "teleport partner"     = "Teleport Partner"
    "distribution partner" = "Distribution Partner"
}

$SATELLITE_ORBIT_MAPPING = @{
    "geo" = "GEO"
    "leo" = "LEO"
}

$REQUESTED_PARTY_TYPE_MAPPING = @{
    "regulator"      = "Regulator"
    "administration" = "Administration"
    "organization"   = "Organization"
}

$VERTICAL_MAPPING = @{
    "space & nw"    = "Space & NW"
    "land fixed"    = "Land Fixed"
    "land mobility" = "Land Mobility"
    "aviation"      = "Aviation"
    "maritime"      = "Maritime"
    "snp"           = "SNP"
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

function Get-LookupIdFromFieldValue {
    param(
        [Parameter(Mandatory = $false)]
        $FieldValue
    )

    if ($null -eq $FieldValue) {
        return @()
    }

    $result = @()

    if ($FieldValue -is [Array]) {
        foreach ($entry in $FieldValue) {
            if ($entry -and $entry.PSObject.Properties.Name -contains "LookupId") {
                $result += [int]$entry.LookupId
            }
            elseif ($entry -is [int]) {
                $result += [int]$entry
            }
            elseif ($entry -is [string] -and $entry -match '^\d+$') {
                $result += [int]$entry
            }
        }

        return $result
    }

    if ($FieldValue.PSObject.Properties.Name -contains "LookupId") {
        $result += [int]$FieldValue.LookupId
        return $result
    }

    if ($FieldValue -is [int]) {
        $result += [int]$FieldValue
        return $result
    }

    if ($FieldValue -is [string] -and $FieldValue -match '^\d+$') {
        $result += [int]$FieldValue
        return $result
    }

    return @()
}

function Get-RequirementStatusLookupMap {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ListName,

        [Parameter(Mandatory = $true)]
        [string]$StatusFieldName,

        [Parameter(Mandatory = $true)]
        [string]$RequirementTypeFieldName
    )

    $map = @{}
    $retryCount = 0
    $maxRetries = 3

    while ($retryCount -lt $maxRetries) {
        try {
            $items = Get-PnPListItem -List $ListName -Fields $StatusFieldName, $RequirementTypeFieldName -PageSize 2000 -ErrorAction Stop
            foreach ($item in $items) {
                $statusTitle = [string]$item[$StatusFieldName]
                if ([string]::IsNullOrWhiteSpace($statusTitle)) {
                    continue
                }

                $normalizedStatus = $statusTitle.Trim().ToLowerInvariant()
                $typeIds = Get-LookupIdFromFieldValue -FieldValue $item[$RequirementTypeFieldName]

                foreach ($typeId in $typeIds) {
                    $key = "$typeId|$normalizedStatus"
                    if (-not $map.ContainsKey($key)) {
                        $map[$key] = [int]$item.Id
                    }
                }
            }

            Write-Log "Loaded $($map.Count) requirement-status mappings from '$ListName'."
            return $map
        }
        catch {
            $retryCount++
            if ($retryCount -lt $maxRetries) {
                Write-Log "Failed to preload requirement-status map '$ListName' (attempt $retryCount). Retrying..." "WARN"
                Start-Sleep -Seconds 2
            }
            else {
                throw "Failed to preload requirement-status map '$ListName' after $maxRetries attempts: $_"
            }
        }
    }

    return $map
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

function Resolve-SingleLookupIdStrict {
    param(
        [Parameter(Mandatory = $false)]
        [string]$RawValue,

        [Parameter(Mandatory = $true)]
        [hashtable]$LookupMap,

        [Parameter(Mandatory = $true)]
        [string]$FieldLabel,

        [Parameter(Mandatory = $true)]
        [int]$RowNumber,

        [Parameter(Mandatory = $false)]
        [bool]$Required = $false
    )

    if ([string]::IsNullOrWhiteSpace($RawValue)) {
        if ($Required) {
            Add-RowSkip -RowNumber $RowNumber -Reason "Skipped - $FieldLabel is required"
            return $null
        }

        return ""
    }

    $normalized = $RawValue.Trim().ToLowerInvariant()
    if (-not $LookupMap.ContainsKey($normalized)) {
        Add-RowSkip -RowNumber $RowNumber -Reason "Skipped - $FieldLabel '$RawValue' not found"
        return $null
    }

    return [int]$LookupMap[$normalized]
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
        [int]$RowNumber,

        [Parameter(Mandatory = $false)]
        [bool]$Required = $false
    )

    if ([string]::IsNullOrWhiteSpace($RawValue)) {
        if ($Required) {
            Add-RowSkip -RowNumber $RowNumber -Reason "Skipped - $FieldLabel is required"
            return $false
        }

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

function Update-DocumentSpaceLink {
    param(
        [Parameter(Mandatory = $true)]
        [int]$ItemId,

        [Parameter(Mandatory = $true)]
        [string]$ItemTitle
    )

    try {
        $parentFolderPath = $FOLDER_URL_MAP["Nol"]
        if ([string]::IsNullOrWhiteSpace($parentFolderPath)) {
            $parentFolderPath = $DOCUMENT_SPACE.NOL_FOLDER
        }

        $sanitizedFolderName = Get-SanitizedFolderName -FolderName $ItemTitle
        if ([string]::IsNullOrWhiteSpace($sanitizedFolderName)) {
            $sanitizedFolderName = "Item-$ItemId"
        }

        $parentFolderUrl = "$($DOCUMENT_SPACE.LIBRARY_NAME)$parentFolderPath"
        $folderUrl = New-UniqueFolder -ParentFolderUrl $parentFolderUrl -BaseName $sanitizedFolderName

        $web = Get-PnPWeb
        $siteRelativeFolderUrl = "$($web.ServerRelativeUrl.TrimEnd('/'))/$folderUrl"
        $hyperlinkValue = "$siteRelativeFolderUrl, $($DOCUMENT_SPACE.HYPERLINK_DESC)"

        Set-PnPListItem -List $ListName -Identity $ItemId -Values @{
            $SP_FIELDS.DOCUMENTS_SPACE = $hyperlinkValue
        } -ErrorAction Stop

        Write-Log "Updated Documents Space for item ID $ItemId" "INFO"
    }
    catch {
        Write-Log "Warning: Failed to create/update Documents Space for item ID $ItemId : $_" "WARN"
    }
}

try {
    Write-Log "================================"
    Write-Log "MA Requirement NOL List Data Import Started"
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
    $requirementTypeLookupMap = Get-LookupMap -ListName $SP_LISTS.REQUIREMENT_TYPE
    $contactLookupMap = Get-LookupMap -ListName $SP_LISTS.CONTACT
    $authorityLookupMap = Get-LookupMap -ListName $SP_LISTS.AUTHORITY
    $teleportPartnerLookupMap = Get-LookupMap -ListName $SP_LISTS.TELEPORT_PARTNER
    $distributionPartnerLookupMap = Get-LookupMap -ListName $SP_LISTS.DISTRIBUTION_PARTNER
    $eutelsatEntityLookupMap = Get-LookupMap -ListName $SP_LISTS.EUTELSAT_ENTITY
    $satelliteNetworkPortalLookupMap = Get-LookupMap -ListName $SP_LISTS.SATELLITE_NETWORK_PORTAL
    $requirementStatusLookupMap = Get-RequirementStatusLookupMap -ListName $SP_LISTS.REQUIREMENT_STATUS -StatusFieldName "Title" -RequirementTypeFieldName $SP_FIELDS.REQUIREMENT_TYPE

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
            $titleRaw = [string]$record.($CSV_COLUMNS.TITLE)
            $marketRaw = [string]$record.($CSV_COLUMNS.MARKET)
            $requirementTypeRaw = [string]$record.($CSV_COLUMNS.REQUIREMENT_TYPE)
            $synthesisStatusRaw = [string]$record.($CSV_COLUMNS.SYNTHESIS_STATUS)
            $responsibleRaw = [string]$record.($CSV_COLUMNS.RESPONSIBLE)
            $relatedItemRaw = [string]$record.($CSV_COLUMNS.RELATED_ITEM)
            $satelliteNetworkPortalRaw = [string]$record.($CSV_COLUMNS.SATELLITE_NETWORK_PORTAL)

            $titleValue = if (-not [string]::IsNullOrWhiteSpace($titleRaw)) { $titleRaw.Trim() } else { "" }
            $marketRaw = if (-not [string]::IsNullOrWhiteSpace($marketRaw)) { $marketRaw.Trim() } else { "" }
            $requirementTypeRaw = if (-not [string]::IsNullOrWhiteSpace($requirementTypeRaw)) { $requirementTypeRaw.Trim() } else { "" }
            $synthesisStatusRaw = if (-not [string]::IsNullOrWhiteSpace($synthesisStatusRaw)) { $synthesisStatusRaw.Trim() } else { "" }
            $responsibleRaw = if (-not [string]::IsNullOrWhiteSpace($responsibleRaw)) { $responsibleRaw.Trim() } else { "" }
            $relatedItemRaw = if (-not [string]::IsNullOrWhiteSpace($relatedItemRaw)) { $relatedItemRaw.Trim() } else { "" }
            $satelliteNetworkPortalRaw = if (-not [string]::IsNullOrWhiteSpace($satelliteNetworkPortalRaw)) { $satelliteNetworkPortalRaw.Trim() } else { "" }

            Write-Log ("Processing record {0}/{1} - Title={2}" -f $recordCount, $csvData.Count, $titleValue)

            if ([string]::IsNullOrWhiteSpace($titleValue)) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Title field is empty"
                continue
            }

            if ([string]::IsNullOrWhiteSpace($requirementTypeRaw)) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Requirement Type is required"
                continue
            }

            $requirementTypeKey = $requirementTypeRaw.ToLowerInvariant()
            if (-not $requirementTypeLookupMap.ContainsKey($requirementTypeKey)) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Requirement Type '$requirementTypeRaw' not found in list '$($SP_LISTS.REQUIREMENT_TYPE)'"
                continue
            }
            $requirementTypeId = [int]$requirementTypeLookupMap[$requirementTypeKey]

            if ([string]::IsNullOrWhiteSpace($synthesisStatusRaw)) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Synthesis Status is required"
                continue
            }

            $synthesisStatusKey = "$requirementTypeId|$($synthesisStatusRaw.ToLowerInvariant())"
            if (-not $requirementStatusLookupMap.ContainsKey($synthesisStatusKey)) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Synthesis Status '$synthesisStatusRaw' not found for Requirement Type '$requirementTypeRaw' in list '$($SP_LISTS.REQUIREMENT_STATUS)'"
                continue
            }
            $synthesisStatusId = [int]$requirementStatusLookupMap[$synthesisStatusKey]

            $responsibleMapped = Get-MappedChoiceValue -RawValue $responsibleRaw -Map $RESPONSIBLE_MAPPING
            if ([string]::IsNullOrWhiteSpace($responsibleMapped)) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Invalid Responsible value: '$responsibleRaw'"
                continue
            }

            $itemData = @{
                $SP_FIELDS.TITLE          = $titleValue
                $SP_FIELDS.CONTENTTYPE_ID = $CONTENT_TYPES["Nol"]
                $SP_FIELDS.REQUIREMENT_TYPE = $requirementTypeId
                $SP_FIELDS.SYNTHESIS_STATUS = $synthesisStatusId
                $SP_FIELDS.RESPONSIBLE    = $responsibleMapped
            }

            if (-not [string]::IsNullOrWhiteSpace($marketRaw)) {
                $marketIds = Resolve-MultiLookupIdsStrict -RawValue $marketRaw -LookupMap $marketLookupMap -FieldLabel "Market" -RowNumber $recordCount
                if ($null -eq $marketIds) {
                    continue
                }

                if ($marketIds.Count -gt 0) {
                    $itemData[$SP_FIELDS.MARKET] = $marketIds
                }
            }

            $verticalsRaw = [string]$record.($CSV_COLUMNS.VERTICAL)
            $verticals = @()
            if (-not [string]::IsNullOrWhiteSpace($verticalsRaw)) {
                $verticals = Get-MappedMultiChoiceValues -RawValue $verticalsRaw -Map $VERTICAL_MAPPING -FieldLabel "Vertical(s)" -RowNumber $recordCount
                if ($null -eq $verticals) {
                    continue
                }

                if ($verticals.Count -gt 0) {
                    $itemData[$SP_FIELDS.VERTICAL] = $verticals
                }
            }

            if (-not [string]::IsNullOrWhiteSpace($satelliteNetworkPortalRaw)) {
                $verticalIncludesSnp = $false
                if (-not [string]::IsNullOrWhiteSpace($verticalsRaw)) {
                    $verticalTokens = $verticalsRaw -split '[,;|]' |
                        ForEach-Object { $_.Trim() } |
                        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

                    foreach ($verticalToken in $verticalTokens) {
                        if ($verticalToken.ToLowerInvariant() -eq "snp") {
                            $verticalIncludesSnp = $true
                            break
                        }
                    }
                }

                if (-not $verticalIncludesSnp) {
                    Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Satellite Network Portal(s) can be set only when Vertical(s) contains 'SNP'"
                    continue
                }

                $satelliteNetworkPortalIds = Resolve-MultiLookupIdsStrict -RawValue $satelliteNetworkPortalRaw -LookupMap $satelliteNetworkPortalLookupMap -FieldLabel "Satellite Network Portal(s)" -RowNumber $recordCount
                if ($null -eq $satelliteNetworkPortalIds) {
                    continue
                }

                if ($satelliteNetworkPortalIds.Count -gt 0) {
                    $itemData[$SP_FIELDS.SATELLITE_NETWORK_PORTAL] = $satelliteNetworkPortalIds
                }
            }

            if (-not (Set-ChoiceFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.RAG_STATUS -RawValue ([string]$record.($CSV_COLUMNS.RAG_STATUS)) -Map $RAG_STATUS_MAPPING -FieldLabel "RAG Status" -RowNumber $recordCount)) { continue }
            if (-not (Set-ChoiceFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.SATELLITE_ORBIT -RawValue ([string]$record.($CSV_COLUMNS.SATELLITE_ORBIT)) -Map $SATELLITE_ORBIT_MAPPING -FieldLabel "Satellite Orbit" -RowNumber $recordCount)) { continue }
            if (-not (Set-ChoiceFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.REQUESTED_PARTY_TYPE -RawValue ([string]$record.($CSV_COLUMNS.REQUESTED_PARTY_TYPE)) -Map $REQUESTED_PARTY_TYPE_MAPPING -FieldLabel "Requested Party Type" -RowNumber $recordCount)) { continue }

            if (-not (Set-DateFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.LETTER_1_MILESTONE -RawValue ([string]$record.($CSV_COLUMNS.LETTER_1_MILESTONE)) -FieldLabel "Letter 1 Milestone" -RowNumber $recordCount)) { continue }
            if (-not (Set-DateFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.LETTER_1_SENT_DATE -RawValue ([string]$record.($CSV_COLUMNS.LETTER_1_SENT_DATE)) -FieldLabel "Letter 1 Sent Date" -RowNumber $recordCount)) { continue }
            if (-not (Set-DateFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.LETTER_2_MILESTONE -RawValue ([string]$record.($CSV_COLUMNS.LETTER_2_MILESTONE)) -FieldLabel "Letter 2 Milestone" -RowNumber $recordCount)) { continue }
            if (-not (Set-DateFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.LETTER_2_SENT_DATE -RawValue ([string]$record.($CSV_COLUMNS.LETTER_2_SENT_DATE)) -FieldLabel "Letter 2 Sent Date" -RowNumber $recordCount)) { continue }
            if (-not (Set-DateFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.RESPONSE_DATE -RawValue ([string]$record.($CSV_COLUMNS.RESPONSE_DATE)) -FieldLabel "Response Date" -RowNumber $recordCount)) { continue }

            Set-CommentFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.SUMMARY -RawValue ([string]$record.($CSV_COLUMNS.SUMMARY))
            Set-CommentFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.COMMETNS -RawValue ([string]$record.($CSV_COLUMNS.COMMETNS))
            Set-CommentFieldIfProvided -ItemData $itemData -FieldInternalName $SP_FIELDS.RESPONSE_SUMMARY -RawValue ([string]$record.($CSV_COLUMNS.RESPONSE_SUMMARY))

            $responsibleContactsRaw = [string]$record.($CSV_COLUMNS.RESPONSIBLE_CONTACTS)
            if (-not [string]::IsNullOrWhiteSpace($responsibleContactsRaw)) {
                $responsibleContactIds = Resolve-MultiLookupIdsStrict -RawValue $responsibleContactsRaw -LookupMap $contactLookupMap -FieldLabel "Responsible Contact(s)" -RowNumber $recordCount
                if ($null -eq $responsibleContactIds) {
                    continue
                }

                if ($responsibleContactIds.Count -gt 0) {
                    $itemData[$SP_FIELDS.RESPONSIBLE_CONTACTS] = $responsibleContactIds
                }
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

            $requestedPartyRaw = [string]$record.($CSV_COLUMNS.REQUESTED_PARTY)
            if (-not [string]::IsNullOrWhiteSpace($requestedPartyRaw)) {
                $requestedPartyId = Resolve-SingleLookupIdStrict -RawValue $requestedPartyRaw -LookupMap $authorityLookupMap -FieldLabel "Requested Party" -RowNumber $recordCount
                if ($null -eq $requestedPartyId) {
                    continue
                }

                if ($requestedPartyId -is [int]) {
                    $itemData[$SP_FIELDS.REQUESTED_PARTY] = $requestedPartyId
                }
            }

            $requestedPartyContactsRaw = [string]$record.($CSV_COLUMNS.REQUESTED_PARTY_CONTACTS)
            if (-not [string]::IsNullOrWhiteSpace($requestedPartyContactsRaw)) {
                $requestedPartyContactIds = Resolve-MultiLookupIdsStrict -RawValue $requestedPartyContactsRaw -LookupMap $contactLookupMap -FieldLabel "Requested Party Contact(s)" -RowNumber $recordCount
                if ($null -eq $requestedPartyContactIds) {
                    continue
                }

                if ($requestedPartyContactIds.Count -gt 0) {
                    $itemData[$SP_FIELDS.REQUESTED_PARTY_CONTACTS] = $requestedPartyContactIds
                }
            }

            if ($responsibleMapped -eq "Teleport Partner") {
                if ([string]::IsNullOrWhiteSpace($relatedItemRaw)) {
                    Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Related Item is required when Responsible is 'Teleport Partner'"
                    continue
                }

                $relatedId = Resolve-SingleLookupIdStrict -RawValue $relatedItemRaw -LookupMap $teleportPartnerLookupMap -FieldLabel "Related Item (Teleport Partner)" -RowNumber $recordCount -Required $true
                if ($null -eq $relatedId) {
                    continue
                }

                $itemData[$SP_FIELDS.TELEPORT_PARTNER] = [int]$relatedId
            }
            elseif ($responsibleMapped -eq "Distribution Partner") {
                if ([string]::IsNullOrWhiteSpace($relatedItemRaw)) {
                    Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Related Item is required when Responsible is 'Distribution Partner'"
                    continue
                }

                $relatedId = Resolve-SingleLookupIdStrict -RawValue $relatedItemRaw -LookupMap $distributionPartnerLookupMap -FieldLabel "Related Item (Distribution Partner)" -RowNumber $recordCount -Required $true
                if ($null -eq $relatedId) {
                    continue
                }

                $itemData[$SP_FIELDS.DISTRIBUTION_PARTNER] = [int]$relatedId
            }
            elseif ($responsibleMapped -eq "Eutelsat") {
                if ([string]::IsNullOrWhiteSpace($relatedItemRaw)) {
                    Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Related Item is required when Responsible is 'Eutelsat'"
                    continue
                }

                $relatedId = Resolve-SingleLookupIdStrict -RawValue $relatedItemRaw -LookupMap $eutelsatEntityLookupMap -FieldLabel "Related Item (Eutelsat Entity)" -RowNumber $recordCount -Required $true
                if ($null -eq $relatedId) {
                    continue
                }

                $itemData[$SP_FIELDS.EUTELSAT_ENTITY] = [int]$relatedId
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
                            Write-Log "DRY RUN: Would create MA Requirement NOL item for '$($batchItem.Title)'"
                            $script:ImportedCount++
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
                            if ($itemId -ne "N/A") {
                                Update-DocumentSpaceLink -ItemId ([int]$itemId) -ItemTitle $batchItem.Title
                            }
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
                    Write-Log "DRY RUN: Would create MA Requirement NOL item for '$($batchItem.Title)'"
                    $script:ImportedCount++
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
                    if ($itemId -ne "N/A") {
                        Update-DocumentSpaceLink -ItemId ([int]$itemId) -ItemTitle $batchItem.Title
                    }
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
