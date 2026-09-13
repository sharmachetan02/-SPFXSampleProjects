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
    TITLE     = "Title"
    DAY       = "Day"
    END_DATE  = "End Date"
    FREQUENCY = "Frequency"
    INTERVAL  = "Interval"
    MONTH     = "Month"
    WEEKDAY   = "Weekday"
}

# ================================
# Configuration - SharePoint Field Internal Names
# ================================
$SP_FIELDS = @{
    TITLE          = "Title"
    CONTENTTYPE_ID = "ContentTypeId"
    DAY            = "MA_RecurrenceDay"
    END_DATE       = "MA_RecurrenceEndDate"
    FREQUENCY      = "MA_RecurrenceFrequency"
    INTERVAL       = "MA_RecurrenceInterval"
    MONTH          = "MA_RecurrenceMonth"
    WEEKDAY        = "MA_RecurrenceWeekDay"
}

# ================================
# Configuration - SharePoint List Names
# ================================
$SP_LISTS = @{
    RECURRENCE_PATTERN = "Recurrence Pattern"
}

# ================================
# Configuration - Content Type Mappings
# ================================
$CONTENT_TYPES = @{
    "Recurrence Pattern" = "0x01003C15960223CC59478A7EF2462B55A50D00"
}

# ================================
# Configuration - Choice Mappings
# ================================
$FREQUENCY_MAPPING = @{
    "weekly"       = "Weekly"
    "monthly"      = "Monthly"
    "quarterly"    = "Quarterly"
    "half annual"  = "Half Annual"
    "annual"       = "Annual"
    "multi annual" = "Multi Annual"
}

$MONTH_MAPPING = @{
    "january"   = "January"
    "february"  = "February"
    "march"     = "March"
    "april"     = "April"
    "may"       = "May"
    "june"      = "June"
    "july"      = "July"
    "august"    = "August"
    "september" = "September"
    "october"   = "October"
    "november"  = "November"
    "december"  = "December"
}

$WEEKDAY_MAPPING = @{
    "monday"    = "Monday"
    "tuesday"   = "Tuesday"
    "wednesday" = "Wednesday"
    "thursday"  = "Thursday"
    "friday"    = "Friday"
    "saturday"  = "Saturday"
    "sunday"    = "Sunday"
}

# Initialize script variables
$script:LogInitialized = $false
$script:ImportedCount = 0
$script:SkippedCount = 0
$script:ErrorCount = 0
$script:RowErrors = @()
$script:RowSkips = @()
$script:AvailableFields = @{}

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
        [hashtable]$Mapping
    )

    if ([string]::IsNullOrWhiteSpace($RawValue)) {
        return $null
    }

    $normalized = $RawValue.Trim().ToLowerInvariant()
    if ($Mapping.ContainsKey($normalized)) {
        return $Mapping[$normalized]
    }

    return $null
}

function Get-IntValue {
    param(
        [Parameter(Mandatory = $false)]
        [string]$RawValue
    )

    if ([string]::IsNullOrWhiteSpace($RawValue)) {
        return $null
    }

    $parsed = 0
    if ([int]::TryParse($RawValue.Trim(), [ref]$parsed)) {
        return $parsed
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
            $parsedDate = [DateTime]::ParseExact(
                $normalized,
                $format,
                [System.Globalization.CultureInfo]::InvariantCulture
            )
            return $parsedDate.Date
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

function Invoke-RecurrencePatternBatch {
    param(
        [Parameter(Mandatory = $true)]
        [array]$BatchItems,

        [Parameter(Mandatory = $true)]
        [string]$TargetListName,

        [Parameter(Mandatory = $true)]
        [hashtable]$FieldMap,

        [Parameter(Mandatory = $true)]
        [bool]$IsDryRun
    )

    if ($null -eq $BatchItems -or $BatchItems.Count -eq 0) {
        return
    }

    Write-Log "Processing batch with $($BatchItems.Count) items..."

    foreach ($batchItem in $BatchItems) {
        try {
            if ($IsDryRun) {
                Write-Log "DRY RUN: Would create item for Recurrence Pattern: $($batchItem.Data[$FieldMap.TITLE])"
                $script:ImportedCount++
            }
            else {
                $retryCount = 0
                $maxRetries = 3
                $newItem = $null

                while ($null -eq $newItem -and $retryCount -lt $maxRetries) {
                    try {
                        $newItem = Add-PnPListItem -List $TargetListName -Values $batchItem.Data -ErrorAction Stop
                    }
                    catch {
                        $retryCount++
                        if ($retryCount -lt $maxRetries) {
                            Write-Log "Failed to add item (attempt $retryCount). Retrying..." "WARN"
                            Start-Sleep -Seconds 2
                        }
                        else {
                            $payloadForLog = ($batchItem.Data | ConvertTo-Json -Depth 5 -Compress)
                            Write-Log "Failed payload for row $($batchItem.Index): $payloadForLog" "ERROR"
                            throw "Failed to add item after $maxRetries attempts: $_"
                        }
                    }
                }

                $itemId = if ($newItem -and $newItem.Id) { $newItem.Id } else { "N/A" }
                Write-Log "Successfully added item ID $itemId for Recurrence Pattern: $($batchItem.Data[$FieldMap.TITLE])" "SUCCESS"
                $script:ImportedCount++
            }
        }
        catch {
            Add-RowError -RowNumber $batchItem.Index -Reason "Error adding item for Recurrence Pattern '$($batchItem.Data[$FieldMap.TITLE])': $($_.Exception.Message)"
        }
    }
}

function Initialize-ListFieldMap {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetListName
    )

    $map = @{}
    $fields = Get-PnPField -List $TargetListName -ErrorAction Stop
    foreach ($field in $fields) {
        if (-not [string]::IsNullOrWhiteSpace($field.InternalName)) {
            $map[$field.InternalName] = $true
        }
    }

    return $map
}

function Test-FieldAvailable {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InternalName
    )

    return $script:AvailableFields.ContainsKey($InternalName)
}

function Get-ListContentTypeId {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetListName,

        [Parameter(Mandatory = $true)]
        [string]$ConfiguredContentTypeId
    )

    try {
        $contentTypes = Get-PnPContentType -List $TargetListName -ErrorAction Stop
        foreach ($ct in $contentTypes) {
            $ctId = [string]$ct.StringId
            if ([string]::IsNullOrWhiteSpace($ctId) -and $ct.Id) {
                $ctId = [string]$ct.Id.StringValue
            }

            if (-not [string]::IsNullOrWhiteSpace($ctId) -and $ctId.StartsWith($ConfiguredContentTypeId)) {
                return $ctId
            }
        }
    }
    catch {
        Write-Log "Could not validate content types for list '$TargetListName': $($_.Exception.Message)" "WARN"
    }

    return $null
}

try {
    Write-Log "================================"
    Write-Log "Recurrence Pattern List Data Import Started"
    Write-Log "================================"
    Write-Log "Site URL: $SiteUrl"
    Write-Log "List Name: $ListName"
    Write-Log "CSV File: $CsvFilePath"
    Write-Log "Dry Run Mode: $DryRun"

    if ($ListName -ne $SP_LISTS.RECURRENCE_PATTERN) {
        Write-Log "Provided list name '$ListName' differs from configured '$($SP_LISTS.RECURRENCE_PATTERN)'." "WARN"
    }

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

    Write-Log "Loading list fields for schema validation..."
    $script:AvailableFields = Initialize-ListFieldMap -TargetListName $ListName
    Write-Log "Loaded $($script:AvailableFields.Count) field internal names from list '$ListName'."

    $configuredContentTypeId = $CONTENT_TYPES["Recurrence Pattern"]
    $resolvedContentTypeId = Get-ListContentTypeId -TargetListName $ListName -ConfiguredContentTypeId $configuredContentTypeId
    if ([string]::IsNullOrWhiteSpace($resolvedContentTypeId)) {
        Write-Log "Configured ContentTypeId '$configuredContentTypeId' was not found on list '$ListName'. Items will be created without explicit ContentTypeId." "WARN"
    }
    else {
        Write-Log "Using list content type ID: $resolvedContentTypeId"
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
            $dayValueRaw = [string]$record.($CSV_COLUMNS.DAY)
            $endDateRaw = [string]$record.($CSV_COLUMNS.END_DATE)
            $frequencyRaw = [string]$record.($CSV_COLUMNS.FREQUENCY)
            $intervalRaw = [string]$record.($CSV_COLUMNS.INTERVAL)
            $monthRaw = [string]$record.($CSV_COLUMNS.MONTH)
            $weekdayRaw = [string]$record.($CSV_COLUMNS.WEEKDAY)

            $titleValue = if (-not [string]::IsNullOrWhiteSpace($titleValue)) { $titleValue.Trim() } else { "" }
            Write-Log ("Processing record {0}/{1} - Recurrence Pattern={2}" -f $recordCount, $csvData.Count, $titleValue)

            if ([string]::IsNullOrWhiteSpace($titleValue)) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Title field is empty"
                continue
            }

            if (-not $titleValue.EndsWith("_Pattern")) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Title must end with '_Pattern'"
                continue
            }

            $dayValue = Get-IntValue -RawValue $dayValueRaw
            if ($null -eq $dayValue) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Day must be a valid number"
                continue
            }

            $intervalValue = Get-IntValue -RawValue $intervalRaw
            if ($null -eq $intervalValue) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Interval must be a valid number"
                continue
            }

            $endDateValue = Get-DateOnlyValue -RawValue $endDateRaw
            if ($null -eq $endDateValue) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - End Date is invalid. Use a date-only format such as yyyy-MM-dd"
                continue
            }

            $mappedFrequency = Get-MappedChoiceValue -RawValue $frequencyRaw -Mapping $FREQUENCY_MAPPING
            if ([string]::IsNullOrWhiteSpace($mappedFrequency)) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Invalid Frequency value: '$frequencyRaw'"
                continue
            }

            $mappedMonth = Get-MappedChoiceValue -RawValue $monthRaw -Mapping $MONTH_MAPPING
            if ([string]::IsNullOrWhiteSpace($mappedMonth)) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Invalid Month value: '$monthRaw'"
                continue
            }

            $mappedWeekday = Get-MappedChoiceValue -RawValue $weekdayRaw -Mapping $WEEKDAY_MAPPING
            if ([string]::IsNullOrWhiteSpace($mappedWeekday)) {
                Add-RowSkip -RowNumber $recordCount -Reason "Skipped - Invalid Weekday value: '$weekdayRaw'"
                continue
            }

            $itemData = @{}
            $itemData[$SP_FIELDS.TITLE] = $titleValue

            if (-not [string]::IsNullOrWhiteSpace($resolvedContentTypeId)) {
                $itemData[$SP_FIELDS.CONTENTTYPE_ID] = $resolvedContentTypeId
            }

            if (Test-FieldAvailable -InternalName $SP_FIELDS.DAY) {
                $itemData[$SP_FIELDS.DAY] = [int]$dayValue
            }

            if (Test-FieldAvailable -InternalName $SP_FIELDS.END_DATE) {
                # Date-only fields are most stable when sent as yyyy-MM-dd.
                $itemData[$SP_FIELDS.END_DATE] = $endDateValue.ToString("yyyy-MM-dd")
            }

            if (Test-FieldAvailable -InternalName $SP_FIELDS.FREQUENCY) {
                $itemData[$SP_FIELDS.FREQUENCY] = $mappedFrequency
            }

            if (Test-FieldAvailable -InternalName $SP_FIELDS.INTERVAL) {
                $itemData[$SP_FIELDS.INTERVAL] = [int]$intervalValue
            }

            if (Test-FieldAvailable -InternalName $SP_FIELDS.MONTH) {
                $itemData[$SP_FIELDS.MONTH] = $mappedMonth
            }

            if (Test-FieldAvailable -InternalName $SP_FIELDS.WEEKDAY) {
                $itemData[$SP_FIELDS.WEEKDAY] = $mappedWeekday
            }

            $batch += [PSCustomObject]@{
                Data  = $itemData
                Index = $recordCount
            }

            if ($batch.Count -ge $BatchSize) {
                Invoke-RecurrencePatternBatch -BatchItems $batch -TargetListName $ListName -FieldMap $SP_FIELDS -IsDryRun $DryRun.IsPresent
                $batch = @()
            }
        }
        catch {
            Add-RowError -RowNumber $recordCount -Reason "Error processing record: $($_.Exception.Message)"
        }
    }

    if ($batch.Count -gt 0) {
        Invoke-RecurrencePatternBatch -BatchItems $batch -TargetListName $ListName -FieldMap $SP_FIELDS -IsDryRun $DryRun.IsPresent
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
        # Ignore disconnect errors
    }
    Write-Log "Script execution completed."
}
