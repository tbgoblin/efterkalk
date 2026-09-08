$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$folder = Split-Path -Parent $MyInvocation.MyCommand.Path
$outputPath = Join-Path $folder 'mixline-report.json'
$patterns = @('MIX', 'MIXLINE', 'ST000-08.00M-MIX-S0', 'ST000-08.00M-MIX-S5',
    'NITROGEN', 'OXYGEN', 'AZOT', 'N2', 'O2', 'GAS', 'BLEND', 'RATIO', 'PERCENT', 'FLOW')

function Open-ReadOnlyAccessDatabase([string]$path) {
    foreach ($provider in @('Microsoft.ACE.OLEDB.16.0', 'Microsoft.ACE.OLEDB.12.0')) {
        try {
            $connection = [System.Data.OleDb.OleDbConnection]::new(
                "Provider=$provider;Data Source=$path;Mode=Read;Persist Security Info=False;")
            $connection.Open()
            return $connection
        } catch {
            if ($connection) { $connection.Dispose() }
        }
    }
    throw "Nessun provider ACE riesce ad aprire $path"
}

function Convert-Row([System.Data.OleDb.OleDbDataReader]$reader) {
    $row = [ordered]@{}
    for ($index = 0; $index -lt $reader.FieldCount; $index++) {
        $value = if ($reader.IsDBNull($index)) { $null } else { $reader.GetValue($index) }
        if ($value -is [byte[]]) { $value = "<binary:$($value.Length)>" }
        $row[$reader.GetName($index)] = $value
    }
    return $row
}

$report = [ordered]@{ generatedAt = (Get-Date).ToString('o'); databases = @() }

Get-ChildItem -Path $folder -Filter '*.mdb' | Sort-Object Name | ForEach-Object {
    $database = [ordered]@{ file = $_.Name; tables = @(); errors = @() }
    $connection = $null
    try {
        $connection = Open-ReadOnlyAccessDatabase $_.FullName
        $schema = $connection.GetOleDbSchemaTable([System.Data.OleDb.OleDbSchemaGuid]::Tables,
            @($null, $null, $null, 'TABLE'))
        foreach ($schemaRow in $schema.Rows) {
            $tableName = [string]$schemaRow.TABLE_NAME
            if ($tableName -like 'MSys*') { continue }
            $table = [ordered]@{ name = $tableName; columns = @(); matchingRows = @(); error = $null }
            try {
                $columns = $connection.GetOleDbSchemaTable([System.Data.OleDb.OleDbSchemaGuid]::Columns,
                    @($null, $null, $tableName, $null))
                $table.columns = @($columns.Rows | Sort-Object ORDINAL_POSITION | ForEach-Object {
                    $columnSize = if ($_.IsNull('CHARACTER_MAXIMUM_LENGTH')) { $null } else { [int]$_.CHARACTER_MAXIMUM_LENGTH }
                    [ordered]@{ name = [string]$_.COLUMN_NAME; type = [int]$_.DATA_TYPE; size = $columnSize }
                })

                $command = $connection.CreateCommand()
                $command.CommandText = "SELECT * FROM [$($tableName.Replace(']', ']]'))]"
                $command.CommandTimeout = 60
                $reader = $command.ExecuteReader()
                try {
                    while ($reader.Read()) {
                        $row = Convert-Row $reader
                        $searchText = ($row.Values | Where-Object { $_ -ne $null } | ForEach-Object { [string]$_ }) -join ' | '
                        if ($patterns | Where-Object { $searchText.IndexOf($_, [StringComparison]::OrdinalIgnoreCase) -ge 0 }) {
                            $table.matchingRows += $row
                            if ($table.matchingRows.Count -ge 500) { break }
                        }
                    }
                } finally { $reader.Dispose() }
            } catch { $table.error = $_.Exception.Message }
            if ($table.matchingRows.Count -gt 0 -or ($table.columns.name -match 'gas|mix|oxy|nit|flow|ratio|percent|tryk').Count -gt 0) {
                $database.tables += $table
            }
        }
    } catch {
        $database.errors += $_.Exception.Message
    } finally {
        if ($connection) { $connection.Close(); $connection.Dispose() }
    }
    $report.databases += $database
}

$report | ConvertTo-Json -Depth 12 | Set-Content -Path $outputPath -Encoding UTF8
Write-Host "Report creato: $outputPath"