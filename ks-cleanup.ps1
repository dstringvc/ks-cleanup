# EnableDatabaseUpdates controls whether this script writes to the database (and, in turn,
# runs the NodeJS downloader). Omit it to do a dry run: the spreadsheet is still populated and
# ks-cleanup.log records what would have been changed, but no UPDATE statements are executed and
# the downloader is not launched. Example: .\ks-cleanup.ps1 -EnableDatabaseUpdates
param(
    [switch]$EnableDatabaseUpdates
)

$ErrorActionPreference = 'Stop'

# KSUpdates.xlsx must be located in the same directory as this script file. Output/log files
# are also written there, regardless of the caller's current working directory.
$currentDir = $PSScriptRoot
$excelPath  = Join-Path $currentDir 'KSUpdates.xlsx'
$outputPath = Join-Path $currentDir 'download-rawfileid-ks.txt'
$logPath    = Join-Path $currentDir 'ks-cleanup.log'

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Add-Content -Path $logPath -Value "[$timestamp] $Message"
}

function Install-RequiredModule {
    param([Parameter(Mandatory)][string]$Name)

    if (Get-Module -ListAvailable -Name $Name) {
        return
    }

    if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null
    }

    if ((Get-PSRepository -Name PSGallery).InstallationPolicy -ne 'Trusted') {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    }

    Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
}

$connection   = $null
$excelPackage = $null

try {
    if ($EnableDatabaseUpdates) {
        Write-Log 'Starting run with database updates ENABLED.'
    }
    else {
        Write-Log 'Starting run with database updates DISABLED (dry run mode).'
    }

    Install-RequiredModule -Name ImportExcel
    Import-Module ImportExcel -ErrorAction Stop

    $connectionString = $env:DBConnectionStringProd
    if ([string]::IsNullOrWhiteSpace($connectionString)) {
        throw "Environment variable 'DBConnectionStringProd' is not set."
    }

    $connection = New-Object System.Data.SqlClient.SqlConnection $connectionString
    $connection.Open()

    $excelPackage = Open-ExcelPackage -Path $excelPath
    $worksheet = $excelPackage.Workbook.Worksheets[1]

    $columns = @{}
    $endCol = $worksheet.Dimension.End.Column
    for ($c = 1; $c -le $endCol; $c++) {
        $header = $worksheet.Cells[1, $c].Text
        if ($header) { $columns[$header] = $c }
    }

    # Validate the spreadsheet contains the expected columns
    foreach ($col in 'State', 'Session', 'BillType', 'BillNumber', 'RawFileID', 'BillTextURL', 'EffectiveDate', 'EffectiveDateUpdated', 'ChapterNumber') {
        if (-not $columns.ContainsKey($col)) {
            throw "Expected column '$col' not found in worksheet."
        }
    }

    $foundRawFileIds = New-Object System.Collections.Generic.List[string]
    $endRow = $worksheet.Dimension.End.Row

    # Determine the last row that actually contains data (Dimension.End.Row can overstate this
    # if trailing rows only carry formatting, e.g. from a Google Sheets export)
    while ($endRow -gt 1 -and
        [string]::IsNullOrWhiteSpace(($worksheet.Cells[$endRow, $columns['State']]).Text) -and
        [string]::IsNullOrWhiteSpace(($worksheet.Cells[$endRow, $columns['BillNumber']]).Text)) {
        $endRow--
    }

    # Iterate over the rows in the spreadsheet
    for ($row = 2; $row -le $endRow; $row++) {
        try {
            $state       = $worksheet.Cells[$row, $columns['State']].Text
            $session     = $worksheet.Cells[$row, $columns['Session']].Text
            $billType    = $worksheet.Cells[$row, $columns['BillType']].Text
            $billTextUrl = $worksheet.Cells[$row, $columns['BillTextURL']].Text
            $billNumber  = $worksheet.Cells[$row, $columns['BillNumber']].Text

            if ([string]::IsNullOrWhiteSpace($state) -and [string]::IsNullOrWhiteSpace($billNumber)) {
                continue
            }

            # Query that finds a row in the RawFileID table matching this spreadsheet row
            $selectCommand = $connection.CreateCommand()
            $selectCommand.CommandText = 'SELECT RawFileID FROM [RawFileID] WHERE [State]=@state AND [Session]=@session AND BillType=@billType AND BillNumber=@billNumber'
            $selectCommand.Parameters.AddWithValue('@state', $state) | Out-Null
            $selectCommand.Parameters.AddWithValue('@session', $session) | Out-Null
            $selectCommand.Parameters.AddWithValue('@billType', $billType) | Out-Null
            $selectCommand.Parameters.AddWithValue('@billNumber', $billNumber) | Out-Null

            $rawFileId = $selectCommand.ExecuteScalar()

            if ($null -eq $rawFileId -or [System.DBNull]::Value.Equals($rawFileId)) {
                $worksheet.Cells[$row, $columns['RawFileID']].Value = 'NotFound'
                continue
            }

            $worksheet.Cells[$row, $columns['RawFileID']].Value = $rawFileId

            # Query that updates the LegislationPath column in the RawFileID table
            if ($EnableDatabaseUpdates) {
                $updateRawFileCommand = $connection.CreateCommand()
                $updateRawFileCommand.CommandText = 'UPDATE [RawFileID] SET LegislationPath=@BillTextURL, LockURL=1 WHERE RawFileID=@rawFileID'
                $updateRawFileCommand.Parameters.AddWithValue('@BillTextURL', [string]$billTextUrl) | Out-Null
                $updateRawFileCommand.Parameters.AddWithValue('@rawFileID', $rawFileId) | Out-Null
                $updateRawFileCommand.ExecuteNonQuery() | Out-Null
            }
            else {
                Write-Log "Row $row [DRY RUN]: would UPDATE [RawFileID] SET LegislationPath='$billTextUrl', LockURL=1 WHERE RawFileID=$rawFileId"
            }

            $effectiveDateValue = $worksheet.Cells[$row, $columns['EffectiveDate']].Value

            # Query that tries to find a row in the Legis table for this RawFileID and, if found,
            # updates the EffectiveDate and PublicActNo columns (no matching row is expected and
            # OK). Skipped entirely when the spreadsheet's EffectiveDate cell is blank.
            if ($null -ne $effectiveDateValue) {
                $chapterNumber = $worksheet.Cells[$row, $columns['ChapterNumber']].Text

                if ($EnableDatabaseUpdates) {
                    $updateLegisCommand = $connection.CreateCommand()
                    $updateLegisCommand.CommandText = 'UPDATE [Legis] SET EffectiveDate=@effectiveDate, PublicActNo=@publicActNo WHERE RawDataID=@rawFileID'
                    $updateLegisCommand.Parameters.AddWithValue('@effectiveDate', $effectiveDateValue) | Out-Null
                    $updateLegisCommand.Parameters.AddWithValue('@publicActNo', $chapterNumber) | Out-Null
                    $updateLegisCommand.Parameters.AddWithValue('@rawFileID', $rawFileId) | Out-Null
                    $legisRowsAffected = $updateLegisCommand.ExecuteNonQuery()

                    if ($legisRowsAffected -gt 0) {
                        $worksheet.Cells[$row, $columns['EffectiveDateUpdated']].Value = $true
                    }
                }
                else {
                    Write-Log "Row $row [DRY RUN]: would UPDATE [Legis] SET EffectiveDate='$effectiveDateValue', PublicActNo='$chapterNumber' WHERE RawDataID=$rawFileId"
                }
            }

            $foundRawFileIds.Add([string]$rawFileId)
        }
        catch {
            Write-Log "Row $row error: $($_.Exception.Message)"
        }
    }

    Close-ExcelPackage -ExcelPackage $excelPackage
    $excelPackage = $null

    $foundRawFileIds | Set-Content -Path $outputPath -Encoding UTF8
}
catch {
    Write-Log "Fatal error: $($_.Exception.Message)"
    if ($excelPackage) {
        Close-ExcelPackage -ExcelPackage $excelPackage -NoSave
    }
    exit 1
}
finally {
    if ($connection -and $connection.State -eq 'Open') {
        $connection.Close()
    }
}

# Run the NodeJS application to download the legislation files found above,
# but only when database updates are actually enabled for this run
if ($EnableDatabaseUpdates) {
    try {
        & node C:\apps\llarts\downloader custom --state=KS "--rawfilelist=$outputPath" -d
    }
    catch {
        Write-Log "Node downloader error: $($_.Exception.Message)"
    }
}
