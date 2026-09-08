$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$folder = Split-Path -Parent $MyInvocation.MyCommand.Path
$outputPath = Join-Path $folder 'mixline-fast-report.json'
$tables = @('Generale', 'Generalestern', 'NuoviParametri', 'NuoviParamestern', 'MaterialThickness')

function Open-Database([string]$path) {
    foreach ($provider in @('Microsoft.ACE.OLEDB.16.0', 'Microsoft.ACE.OLEDB.12.0')) {
        try {
            $connection = [System.Data.OleDb.OleDbConnection]::new("Provider=$provider;Data Source=$path;Mode=Read;")
            $connection.Open()
            return $connection
        } catch {
            if ($connection) { $connection.Dispose() }
        }
    }
    throw "ACE non riesce ad aprire $path"
}

function Read-Query($connection, [string]$sql) {
    $command = $connection.CreateCommand()
    $command.CommandText = $sql
    $command.CommandTimeout = 20
    $reader = $command.ExecuteReader()
    $rows = @()
    try {
        while ($reader.Read()) {
            $row = [ordered]@{}
            for ($index = 0; $index -lt $reader.FieldCount; $index++) {
                $row[$reader.GetName($index)] = if ($reader.IsDBNull($index)) { $null } else { $reader.GetValue($index) }
            }
            $rows += $row
        }
    } finally { $reader.Dispose() }
    return $rows
}

$report = [ordered]@{ generatedAt = (Get-Date).ToString('o'); databases = @() }
Get-ChildItem $folder -Filter '*.mdb' | Sort-Object Name | ForEach-Object {
    $database = [ordered]@{ file = $_.Name; tables = @(); errors = @() }
    $connection = $null
    try {
        $connection = Open-Database $_.FullName
        foreach ($table in $tables) {
            try {
                $rows = Read-Query $connection "SELECT * FROM [$table] WHERE [Mat_Id] = 'ST-MIX' AND [Spess] = 8"
                if ($rows.Count) { $database.tables += [ordered]@{ name = $table; rows = $rows } }
            } catch {
                if ($_.Exception.Message -notmatch 'could not find|non è stato trovato|cannot find') {
                    $database.errors += "$table`: $($_.Exception.Message)"
                }
            }
        }
    } catch { $database.errors += $_.Exception.Message }
    finally {
        if ($connection) { $connection.Close(); $connection.Dispose() }
    }
    if ($database.tables.Count -or $database.errors.Count) { $report.databases += $database }
}
$report | ConvertTo-Json -Depth 10 | Set-Content $outputPath -Encoding UTF8
Write-Host "Report rapido creato: $outputPath"
