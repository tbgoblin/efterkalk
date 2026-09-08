const fs = require('fs');
const zlib = require('zlib');
const { execFile } = require('child_process');

function decodeXml(value) {
    return String(value || '')
        .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'").replace(/&amp;/g, '&');
}

function readZipEntry(buffer, wantedName) {
    let eocd = -1;
    for (let offset = buffer.length - 22; offset >= Math.max(0, buffer.length - 65557); offset -= 1) {
        if (buffer.readUInt32LE(offset) === 0x06054b50) { eocd = offset; break; }
    }
    if (eocd < 0) throw new Error('Ugyldig XLSM/ZIP-fil');
    const entryCount = buffer.readUInt16LE(eocd + 10);
    let offset = buffer.readUInt32LE(eocd + 16);
    for (let index = 0; index < entryCount; index += 1) {
        if (buffer.readUInt32LE(offset) !== 0x02014b50) throw new Error('Ugyldigt ZIP-katalog');
        const method = buffer.readUInt16LE(offset + 10);
        const compressedSize = buffer.readUInt32LE(offset + 20);
        const nameLength = buffer.readUInt16LE(offset + 28);
        const extraLength = buffer.readUInt16LE(offset + 30);
        const commentLength = buffer.readUInt16LE(offset + 32);
        const localOffset = buffer.readUInt32LE(offset + 42);
        const name = buffer.toString('utf8', offset + 46, offset + 46 + nameLength).replace(/\\/g, '/');
        if (name === wantedName) {
            const localNameLength = buffer.readUInt16LE(localOffset + 26);
            const localExtraLength = buffer.readUInt16LE(localOffset + 28);
            const dataOffset = localOffset + 30 + localNameLength + localExtraLength;
            const compressed = buffer.subarray(dataOffset, dataOffset + compressedSize);
            if (method === 0) return compressed;
            if (method === 8) return zlib.inflateRawSync(compressed);
            throw new Error('Ikke-understøttet ZIP-komprimering: ' + method);
        }
        offset += 46 + nameLength + extraLength + commentLength;
    }
    return null;
}

function parseSharedStrings(xml) {
    if (!xml) return [];
    return Array.from(xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g), match =>
        Array.from(match[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g), text => decodeXml(text[1])).join(''));
}

function columnIndex(cellRef) {
    const letters = String(cellRef || '').match(/^[A-Z]+/i);
    if (!letters) return -1;
    return letters[0].toUpperCase().split('').reduce((value, char) => value * 26 + char.charCodeAt(0) - 64, 0) - 1;
}

function parseWorksheetXml(xml, sharedStrings) {
    const rows = new Map();
    for (const match of xml.matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
        const attributes = match[1];
        const body = match[2];
        const refMatch = attributes.match(/\br="([A-Z]+)(\d+)"/i);
        if (!refMatch) continue;
        const rowNo = Number(refMatch[2]);
        const colNo = columnIndex(refMatch[1]);
        const typeMatch = attributes.match(/\bt="([^"]+)"/);
        const type = typeMatch ? typeMatch[1] : '';
        const valueMatch = body.match(/<v>([\s\S]*?)<\/v>/);
        let value = null;
        if (type === 'inlineStr') {
            value = Array.from(body.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g), item => decodeXml(item[1])).join('');
        } else if (valueMatch) {
            const raw = decodeXml(valueMatch[1]);
            if (type === 's') value = sharedStrings[Number(raw)] ?? '';
            else if (type === 'str') value = raw;
            else if (type === 'b') value = raw === '1';
            else value = raw === '' || Number.isNaN(Number(raw)) ? raw : Number(raw);
        }
        if (!rows.has(rowNo)) rows.set(rowNo, []);
        rows.get(rowNo)[colNo] = value;
    }
    return rows;
}

function findWorksheetPath(buffer, sheetName) {
    const workbookXml = readZipEntry(buffer, 'xl/workbook.xml');
    const relsXml = readZipEntry(buffer, 'xl/_rels/workbook.xml.rels');
    if (!workbookXml || !relsXml) throw new Error('Workbook metadata mangler');
    const escapedName = String(sheetName).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const sheetMatch = workbookXml.toString('utf8').match(new RegExp('<sheet\\b[^>]*name="' + escapedName + '"[^>]*r:id="([^"]+)"', 'i'));
    if (!sheetMatch) throw new Error('Excel-arket ' + sheetName + ' blev ikke fundet');
    const relMatch = relsXml.toString('utf8').match(new RegExp('<Relationship\\b[^>]*Id="' + sheetMatch[1] + '"[^>]*Target="([^"]+)"', 'i'));
    if (!relMatch) throw new Error('Excel-arkets relation blev ikke fundet');
    return ('xl/' + relMatch[1].replace(/^\//, '').replace(/^\.\.\//, '')).replace(/\/\.\//g, '/');
}

function readLaserParameters(workbookPath) {
    const buffer = fs.readFileSync(workbookPath);
    const sharedXml = readZipEntry(buffer, 'xl/sharedStrings.xml');
    const sharedStrings = parseSharedStrings(sharedXml ? sharedXml.toString('utf8') : '');
    const sheetPath = findWorksheetPath(buffer, 'skæreparametre');
    const sheetXml = readZipEntry(buffer, sheetPath);
    if (!sheetXml) throw new Error('Data for arket skæreparametre mangler');
    const rows = parseWorksheetXml(sheetXml.toString('utf8'), sharedStrings);
    const result = [];
    for (const [rowNo, cells] of rows) {
        if (rowNo === 1) continue;
        const row = {
            prodNo: String(cells[0] || '').trim(), description: String(cells[1] || '').trim(),
            thickness: Number(cells[2] || 0), machine: String(cells[3] || '').trim(),
            cutSpeedMPerMin: Number(cells[5]), piercingMinutes: Number(cells[6] || 0),
            surchargePercent: Number(cells[7] || 0), lens: String(cells[8] || '').trim()
        };
        if (row.prodNo && row.machine && Number.isFinite(row.cutSpeedMPerMin) && row.cutSpeedMPerMin >= 0) result.push(row);
    }
    return Array.from(new Map(result.map(row => [(row.prodNo + '\u0000' + row.machine).toLowerCase(), row])).values());
}

function normalizeLaserTechnicalRows(sourceRows) {
    const numberOrZero = value => {
        const parsed = Number(value);
        return Number.isFinite(parsed) ? parsed : 0;
    };
    const result = [];
    for (const cells of sourceRows) {
        const row = {
            technology: String(cells[0] || '').trim(), material: String(cells[1] || '').trim(),
            thickness: Number(cells[2]), lens: String(cells[3] || '').trim(),
            piercingMilliseconds: numberOrZero(cells[4]), vaporPowerW: numberOrZero(cells[5]),
            reducedPowerW: numberOrZero(cells[6]), feedrateLargeMmMin: numberOrZero(cells[7]),
            feedrateMediumMmMin: numberOrZero(cells[8]), feedrateSmallMmMin: numberOrZero(cells[9]),
            feedrateEngravingMmMin: numberOrZero(cells[10]), gasPressureBar: numberOrZero(cells[12]),
            nozzleSizeMm: numberOrZero(cells[13])
        };
        if (row.technology && row.material && Number.isFinite(row.thickness) && row.thickness > 0) result.push(row);
    }
    return Array.from(new Map(result.map(row => [row.technology.toLowerCase(), row])).values());
}

function readLaserTechnicalParameters(workbookPath) {
    const buffer = fs.readFileSync(workbookPath);
    const sharedXml = readZipEntry(buffer, 'xl/sharedStrings.xml');
    const sharedStrings = parseSharedStrings(sharedXml ? sharedXml.toString('utf8') : '');
    const sheetPath = findWorksheetPath(buffer, 'Laserberegner2');
    const sheetXml = readZipEntry(buffer, sheetPath);
    if (!sheetXml) throw new Error('Data for arket Laserberegner2 mangler');
    const rows = parseWorksheetXml(sheetXml.toString('utf8'), sharedStrings);
    return normalizeLaserTechnicalRows(Array.from(rows.entries())
        .filter(([rowNo]) => rowNo > 2)
        .map(([, cells]) => cells.slice(1, 16)));
}

function readLaserTechnicalParametersViaExcel(workbookPath) {
    const script = `
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$connection = $null
try {
    $providers = @('Microsoft.ACE.OLEDB.16.0', 'Microsoft.ACE.OLEDB.12.0')
    $lastError = $null
    foreach ($provider in $providers) {
        try {
            $connectionString = "Provider=$provider;Data Source=$env:BOM_WORKBOOK_PATH;Extended Properties='Excel 12.0 Macro;HDR=NO;IMEX=1;READONLY=TRUE'"
            $connection = New-Object System.Data.OleDb.OleDbConnection($connectionString)
            $connection.Open()
            break
        } catch {
            $lastError = $_
            if ($connection) { try { $connection.Dispose() } catch {} }
            $connection = $null
        }
    }
    if (-not $connection) { throw $lastError }
    $command = $connection.CreateCommand()
    $command.CommandText = 'SELECT * FROM [Laserberegner2$B3:P305]'
    $reader = $command.ExecuteReader()
    $rows = @()
    while ($reader.Read()) {
        $item = @()
        for ($column = 0; $column -lt 15; $column++) {
            $item += $(if ($reader.IsDBNull($column)) { $null } else { $reader.GetValue($column) })
        }
        $rows += ,$item
    }
    $reader.Close()
    ConvertTo-Json -InputObject $rows -Depth 3 -Compress
} finally {
    if ($connection) { try { $connection.Close(); $connection.Dispose() } catch {} }
}`;
    return new Promise((resolve, reject) => {
        execFile('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', script],
            { windowsHide: true, timeout: 60000, maxBuffer: 4 * 1024 * 1024,
                env: { ...process.env, BOM_WORKBOOK_PATH: workbookPath } }, (error, stdout, stderr) => {
                if (error) return reject(new Error('Excel ACE kunne ikke læse Tabel22: ' + String(stderr || error.message).trim()));
                try {
                    const rows = JSON.parse(String(stdout || '[]').replace(/^\uFEFF/, ''));
                    return resolve(normalizeLaserTechnicalRows(Array.isArray(rows) ? rows : []));
                } catch (parseError) {
                    return reject(new Error('Excel returnerede ugyldige Tabel22-data: ' + parseError.message));
                }
            });
    });
}

async function readCompleteLaserTechnicalParameters(workbookPath) {
    const xmlRows = readLaserTechnicalParameters(workbookPath);
    try {
        const excelRows = await readLaserTechnicalParametersViaExcel(workbookPath);
        if (excelRows.length < 100) throw new Error('Excel returnerede kun ' + excelRows.length + ' laserteknologier fra B3:P305');
        return excelRows.length > xmlRows.length ? excelRows : xmlRows;
    } catch (error) {
        if (xmlRows.length < 100) throw error;
        return xmlRows;
    }
}

module.exports = { readLaserParameters, readLaserTechnicalParameters, readCompleteLaserTechnicalParameters, normalizeLaserTechnicalRows, parseWorksheetXml };