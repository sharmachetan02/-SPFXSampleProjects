param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $true)]
    [string]$ListName,

    [Parameter(Mandatory = $true)]
    [string]$CsvFilePath,

    [Parameter(Mandatory = $true)]
    [string]$TermGroupName,

    [Parameter(Mandatory = $true)]
    [string]$TermSetName,

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
    MARKET           = "Market"
    COUNTRY          = "Country"
    IS_SANCTIONED    = "Is Sanctioned"
    SANCTION_CATEGORY = "Sanction Category"
    PRIORITY         = "Priority"
    CAPACITY         = "Capacity"
    COMMENTS         = "Comments"
    AOP_REGIONS      = "AOP Regions"
    GEO_REGIONS      = "Geo Regions"
    GLOBAL_REGIONS   = "Global Regions"
}

# ================================
# Configuration - SharePoint Field Internal Names
# ================================
$SP_FIELDS = @{
    COUNTRY          = "MA_Country"
    CONTENTTYPE_ID   = "ContentTypeId"
    IS_SANCTIONED    = "MA_IsSanctioned"
    SANCTION_CATEGORY = "MA_SanctionCategory"
    PRIORITY         = "MA_Priority"
    CAPACITY         = "MA_Capacity"
    COMMENTS         = "MA_Comments"
    GROUPING         = "MA_Grouping"
    DOCUMENTS_SPACE  = "MA_DocumentsSpace"
}

# ================================
# Configuration - SharePoint List Names
# ================================
$SP_LISTS = @{
    COUNTRY          = "Country"
    SANCTION_CATEGORY = "Sanction Category"
}

# ================================
# Configuration - Document Space Settings
# ================================
$DOCUMENT_SPACE = @{
    LIBRARY_NAME     = "documents_space"
    MARKET_FOLDER    = "/Markets"
    HYPERLINK_DESC   = "Link"
}

# ================================
# Configuration - Term Store Settings
# ================================
$TERM_PARENTS = @{
    AOP_REGIONS      = "AOP Regions"
    GEO_REGIONS      = "Geo Regions"
    GLOBAL_REGIONS   = "Global Regions"
}

# Content Type ID for "Market" content type
$CONTENT_TYPE = "0x010040CF47236D86974EBD56BBB62FC23C79" 

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

# Function to get Term ID from Term Store within a specific parent collection
function Get-TermId {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TermLabel,

        [Parameter(Mandatory = $true)]
        [object]$TermSet,

        [Parameter(Mandatory = $false)]
        [string]$ParentTermName
    )

    try {
        if ([string]::IsNullOrWhiteSpace($TermLabel)) {
            return $null
        }

        # Search for the term in the term set recursively
        function Find-TermRecursive {
            param($term, $searchLabel)

            if ($term.Name -eq $searchLabel) {
                return $term.Id
            }

            if ($term.Terms -and $term.Terms.Count -gt 0) {
                foreach ($childTerm in $term.Terms) {
                    $result = Find-TermRecursive -term $childTerm -searchLabel $searchLabel
                    if ($result) {
                        return $result
                    }
                }
            }

            return $null
        }

        # If parent term name is specified, search only within that parent
        if (-not [string]::IsNullOrWhiteSpace($ParentTermName)) {
            # Find the parent term first
            $parentTerm = $null
            foreach ($term in $TermSet.Terms) {
                if ($term.Name -eq $ParentTermName) {
                    $parentTerm = $term
                    break
                }
            }

            if (-not $parentTerm) {
                Write-Log "Parent term '$ParentTermName' not found in term set '$($TermSet.Name)'" "WARN"
                return $null
            }

            # Reload the parent term to ensure its children are loaded
            $parentTerm = Get-PnPTerm -Identity $parentTerm.Id -TermGroup 'Market Access' -TermSet $TermSet -Includes Terms

            # Search within the parent term's children
            if ($parentTerm.Terms -and $parentTerm.Terms.Count -gt 0) {
                foreach ($childTerm in $parentTerm.Terms) {
                    $result = Find-TermRecursive -term $childTerm -searchLabel $TermLabel
                    if ($result) {
                        return $result
                    }
                }
            }

            Write-Log "Term '$TermLabel' not found under parent term '$ParentTermName'" "WARN"
            return $null
        }

        # If no parent specified, search through all terms in the term set
        foreach ($term in $TermSet.Terms) {
            $termId = Find-TermRecursive -term $term -searchLabel $TermLabel
            if ($termId) {
                return $termId
            }
        }

        Write-Log "Term '$TermLabel' not found in term set '$($TermSet.Name)'" "WARN"
        return $null
    }
    catch {
        Write-Log "Error retrieving term ID for '$TermLabel': $_" "ERROR"
        return $null
    }
}

# Function to build term collection for managed metadata with parent term context
function New-TermCollection {
    param(
        [Parameter(Mandatory = $true)]
        [object]$TermSet,

        [Parameter(Mandatory = $false)]
        [hashtable]$TermsByParent
    )

    try {
        $termCollection = @()

        if ($TermsByParent -and $TermsByParent.Count -gt 0) {
            foreach ($parentTerm in $TermsByParent.Keys) {
                $termLabel = $TermsByParent[$parentTerm]
                
                if (-not [string]::IsNullOrWhiteSpace($termLabel)) {
                    # Search for term within the specified parent term collection
                    $termId = Get-TermId -TermLabel $termLabel -TermSet $TermSet -ParentTermName $parentTerm
                    if ($termId) {
                        $termCollection += $termId
                    }
                }
            }
        }

        return $termCollection
    }
    catch {
        Write-Log "Error building term collection: $_" "ERROR"
        return @()
    }
}

# Function to convert value to proper type
function Convert-FieldValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value,

        [Parameter(Mandatory = $true)]
        [string]$FieldType
    )

    if ([string]::IsNullOrWhiteSpace($Value) -or $Value -eq "#N/A") {
        return $null
    }

    switch ($FieldType.ToLower()) {
        "boolean" {
            return $Value -eq "Yes" -or $Value -eq "True" -or $Value -eq "1"
        }
        "number" {
            try {
                return [double]::Parse($Value)
            }
            catch {
                return $Value
            }
        }
        default {
            return $Value
        }
    }
}

# Function to set multi-select taxonomy field values
function Set-MultiSelectTaxonomyField {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ListItem,

        [Parameter(Mandatory = $true)]
        [string]$FieldInternalName,

        [Parameter(Mandatory = $true)]
        [array]$TermIds,

        [Parameter(Mandatory = $true)]
        [string]$ListName
    )

    try {
        if ($TermIds -and $TermIds.Count -gt 0) {
            # Build array of taxonomy values in format "TermLabel|TermGuid"
            $taxValues = @()
            
            foreach ($termId in $TermIds) {
                try {
                    # Get term details to retrieve the label
                    $term = Get-PnPTerm -Identity $termId -ErrorAction Stop
                    # Format: "TermLabel|TermGuid"
                    $taxValues += "$($term.Name)|$termId"
                }
                catch {
                    Write-Log "Warning: Could not retrieve term details for ID '$termId': $_" "WARN"
                }
            }
            
            if ($taxValues.Count -gt 0) {
                $retryCount = 0
                $maxRetries = 3
                $success = $false
                
                while (-not $success -and $retryCount -lt $maxRetries) {
                    try {
                        # Set all taxonomy values in one operation
                        Set-PnPListItem -List $ListName -Identity $ListItem.Id -Values @{$FieldInternalName = $taxValues} -ErrorAction Stop
                        $success = $true
                        Write-Log "Set $($taxValues.Count) taxonomy term(s) for field '$FieldInternalName'" "INFO"
                    }
                    catch {
                        $retryCount++
                        if ($retryCount -lt $maxRetries) {
                            Write-Log "Failed to set taxonomy field (attempt $retryCount). Retrying..." "WARN"
                            Start-Sleep -Seconds 1
                        }
                        else {
                            throw "Failed to set taxonomy field after $maxRetries attempts: $_"
                        }
                    }
                }
                return $true
            }
        }
        return $false
    }
    catch {
        Write-Log "Warning: Failed to set taxonomy field '$FieldInternalName': $_" "WARN"
        return $false
    }
}

# Main script
try {
    Write-Log "================================"
    Write-Log "Market List Data Import Started"
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
    $effectiveClientId = if (-not [string]::IsNullOrWhiteSpace($ClientId)) { $ClientId } else { $env:ENTRAID_APP_ID }
    
    # Retry connection logic
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

    # Get Term Group and Term Set with retry logic
    Write-Log "Fetching Term Group: $TermGroupName"
    $termGroup = $null
    $retryCount = 0
    
    while ($null -eq $termGroup -and $retryCount -lt $maxRetries) {
        try {
            $termGroup = Get-PnPTermGroup -Identity $TermGroupName -ErrorAction Stop
        }
        catch {
            $retryCount++
            if ($retryCount -lt $maxRetries) {
                Write-Log "Failed to fetch Term Group (attempt $retryCount). Retrying..." "WARN"
                Start-Sleep -Seconds 2
            }
            else {
                throw "Failed to fetch Term Group '$TermGroupName' after $maxRetries attempts: $_"
            }
        }
    }
    
    if (-not $termGroup) {
        throw "Term Group '$TermGroupName' not found."
    }

    Write-Log "Fetching Term Set: $TermSetName"
    $termSet = $null
    $retryCount = 0
    
    while ($null -eq $termSet -and $retryCount -lt $maxRetries) {
        try {
            $termSet = Get-PnPTermSet -TermGroup $termGroup -Identity $TermSetName -Includes Terms -ErrorAction Stop
        }
        catch {
            $retryCount++
            if ($retryCount -lt $maxRetries) {
                Write-Log "Failed to fetch Term Set (attempt $retryCount). Retrying..." "WARN"
                Start-Sleep -Seconds 2
            }
            else {
                throw "Failed to fetch Term Set '$TermSetName' after $maxRetries attempts: $_"
            }
        }
    }
    
    if (-not $termSet) {
        throw "Term Set '$TermSetName' not found."
    }
    Write-Log "✅ Term Set found."

    # Read CSV file
    Write-Log "Reading CSV file: $CsvFilePath"
    $csvData = Import-Csv -Path $CsvFilePath -Encoding UTF8

    if (-not $csvData -or $csvData.Count -eq 0) {
        throw "No data found in CSV file."
    }

    Write-Log "Total records in CSV: $($csvData.Count)"

    Write-Log "Preloading lookup maps..."
    $countryLookupMap = Get-LookupMap -ListName $SP_LISTS.COUNTRY -FieldName "Title"
    $sanctionLookupMap = Get-LookupMap -ListName $SP_LISTS.SANCTION_CATEGORY -FieldName "Title"

    # Process records in batches
    $batch = @()
    $recordCount = 0

    foreach ($record in $csvData) {
        $recordCount++

        try {
            Write-Log "Processing record $recordCount/$($csvData.Count): Market=$($record.($CSV_COLUMNS.MARKET))"

            # Validate required fields
            if ([string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.MARKET))) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Market field is empty"
                continue
            }

            # Build the item hashtable
            $itemData = @{
                Title = $record.($CSV_COLUMNS.MARKET).Trim()
            }

             
                $itemData[$SP_FIELDS.CONTENTTYPE_ID] = $CONTENT_TYPE
            

            # Map Country (Lookup)
            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.COUNTRY))) {
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
            }

            # Map Is Sanctioned (Boolean)
            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.IS_SANCTIONED))) {
                $itemData[$SP_FIELDS.IS_SANCTIONED] = Convert-FieldValue -Value $record.($CSV_COLUMNS.IS_SANCTIONED) -FieldType "boolean"
            }

            # Map Sanction Category (Lookup)
            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.SANCTION_CATEGORY))) {
                $sanctionCategoryValue = $record.($CSV_COLUMNS.SANCTION_CATEGORY).Trim()
                $sanctionKey = $sanctionCategoryValue.ToLowerInvariant()
                $sanctionId = if ($sanctionLookupMap.ContainsKey($sanctionKey)) { $sanctionLookupMap[$sanctionKey] } else { $null }
                if ($sanctionId) {
                    $itemData[$SP_FIELDS.SANCTION_CATEGORY] = $sanctionId
                }
                else {
                    Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Sanction Category lookup not found: '$sanctionCategoryValue'"
                    continue
                }
            }

            # Map Priority (Choice)
            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.PRIORITY))) {
                $itemData[$SP_FIELDS.PRIORITY] = $record.($CSV_COLUMNS.PRIORITY).Trim()
            }

            # Map Capacity (Single line of text)
            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.CAPACITY))) {
                $itemData[$SP_FIELDS.CAPACITY] = $record.($CSV_COLUMNS.CAPACITY).Trim()
            }

            # Map Comments (Single line of text)
            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.COMMENTS))) {
                $itemData[$SP_FIELDS.COMMENTS] = $record.($CSV_COLUMNS.COMMENTS).Trim()
            }

            # Map Managed Metadata (AOP Regions, Geo Regions, Global Regions)
            # Each region type is a separate parent term collection, so we query by parent term name
            $termsByParent = @{}

            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.AOP_REGIONS))) {
                $termsByParent[$TERM_PARENTS.AOP_REGIONS] = $record.($CSV_COLUMNS.AOP_REGIONS).Trim()
            }

            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.GEO_REGIONS))) {
                $termsByParent[$TERM_PARENTS.GEO_REGIONS] = $record.($CSV_COLUMNS.GEO_REGIONS).Trim()
            }

            if (-not [string]::IsNullOrWhiteSpace($record.($CSV_COLUMNS.GLOBAL_REGIONS))) {
                $termsByParent[$TERM_PARENTS.GLOBAL_REGIONS] = $record.($CSV_COLUMNS.GLOBAL_REGIONS).Trim()
            }

            # Store taxonomy terms separately (cannot be added directly to item)
            $taxonomyTermIds = @()
            if ($termsByParent.Count -gt 0) {
                $termCollection = New-TermCollection -TermSet $termSet -TermsByParent $termsByParent
                if ($termCollection.Count -gt 0) {
                    $taxonomyTermIds = $termCollection
                }
            }

            # Add to batch
            $batch += [PSCustomObject]@{
                Data          = $itemData
                TaxonomyTerms = $taxonomyTermIds
                Index         = $recordCount
            }

            # Process batch when it reaches BatchSize or at the end
            if ($batch.Count -ge $BatchSize -or $recordCount -eq $csvData.Count) {
                Write-Log "Processing batch with $($batch.Count) items..."

                foreach ($batchItem in $batch) {
                    try {
                        if ($DryRun) {
                            Write-Log "DRY RUN: Would add item for market: $($batchItem.Data.Title)"
                        }
                        else {
                            # Add the item without taxonomy field with retry logic
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
                            
                            # Get the item ID for logging
                            $itemId = if ($newItem -and $newItem.Id) { $newItem.Id } else { "N/A" }
                            
                            # Set multi-select taxonomy field if there are terms
                            if ($batchItem.TaxonomyTerms -and $batchItem.TaxonomyTerms.Count -gt 0) {
                                Set-MultiSelectTaxonomyField -ListItem $newItem -FieldInternalName $SP_FIELDS.GROUPING -TermIds $batchItem.TaxonomyTerms -ListName $ListName
                            }
                            
                            # Create Document Space folder and update the hyperlink field
                            try {
                                $sanitizedFolderName = Get-SanitizedFolderName -FolderName $batchItem.Data.Title
                                $parentFolderUrl = "$($DOCUMENT_SPACE.LIBRARY_NAME)$($DOCUMENT_SPACE.MARKET_FOLDER)"
                                
                                # Create unique folder
                                $folderUrl = New-UniqueFolder -ParentFolderUrl $parentFolderUrl -BaseName $sanitizedFolderName
                                
                                # Get the site URL to construct the full folder URL
                                $web = Get-PnPWeb
                                $siteRelativeFolderUrl = "$($web.ServerRelativeUrl.TrimEnd('/'))/$folderUrl"
                                
                                # Update the Documents Space field with hyperlink value
                                $hyperlinkValue = "$siteRelativeFolderUrl, $($DOCUMENT_SPACE.HYPERLINK_DESC)"
                                Set-PnPListItem -List $ListName -Identity $newItem.Id -Values @{
                                    $SP_FIELDS.DOCUMENTS_SPACE = $hyperlinkValue
                                } -ErrorAction Stop
                                
                                Write-Log "✅ Created document space folder and updated field for item ID $itemId" "INFO"
                            }
                            catch {
                                Write-Log "⚠ Warning: Failed to create document space folder for item ID $itemId : $_" "WARN"
                                # Continue processing despite folder creation failure
                            }
                            
                            Write-Log "✅ Successfully added item ID $itemId for market: $($batchItem.Data.Title)" "SUCCESS"
                            $script:ImportedCount++
                        }
                    }
                    catch {
                        Add-RowError -RowNumber $batchItem.Index -Reason "Error adding item for market '$($batchItem.Data.Title)': $($_.Exception.Message)"
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

