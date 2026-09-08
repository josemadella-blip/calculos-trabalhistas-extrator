# Brasanitas - TODOS os 8217 casos em TODAS as subpastas
# Estrategia: pegar o primeiro PDF ou xlsx de cada pasta de reclamante
# PDF -> pdftotext | xlsx -> ZIP/XML
# Colunas A-K | CSV + conversao unica para xlsx

$ErrorActionPreference = 'Continue'
$log = "C:\Users\jose.madella\Desktop\extrair_log_brasanitas_tudo.txt"
"INICIO: $(Get-Date)" | Out-File $log -Encoding UTF8

$calcBase = "X:\Trabalhista\Arquivos Trabalhistas\C" + [char]0x00E1 + "lculos Elaborados"
$pdftotext = "C:\Users\jose.madella\AppData\Local\Microsoft\WinGet\Packages\oschwartz10612.Poppler_Microsoft.Winget.Source_8wekyb3d8bbwe\poppler-25.07.0\Library\bin\pdftotext.exe"

# Funcoes XLSX (ZIP/XML)
function Ler-XlsxComoXml($caminho) {
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [System.IO.Compression.ZipFile]::OpenRead($caminho)
        $ssEntry = $zip.Entries | Where-Object { $_.FullName -eq "xl/sharedStrings.xml" }
        $sharedStrings = @()
        if ($ssEntry) {
            $reader = New-Object System.IO.StreamReader($ssEntry.Open())
            $ssXml = [xml]$reader.ReadToEnd(); $reader.Close()
            $sharedStrings = @($ssXml.sst.si | ForEach-Object {
                $t = $_.t
                if (-not $t -and $_.r) { $t = ($_.r | ForEach-Object { $_.t } | Out-String) }
                "$t"
            })
        }
        $sheetEntry = $zip.Entries | Where-Object { $_.FullName -eq "xl/worksheets/sheet1.xml" }
        if (-not $sheetEntry) { $sheetEntry = $zip.Entries | Where-Object { $_.FullName -like "xl/worksheets/sheet*.xml" } | Select-Object -First 1 }
        $sheetXml = $null
        if ($sheetEntry) { $reader = New-Object System.IO.StreamReader($sheetEntry.Open()); $sheetXml = [xml]$reader.ReadToEnd(); $reader.Close() }
        $zip.Dispose()
        return @{ SharedStrings = $sharedStrings; Sheet = $sheetXml }
    } catch { return $null }
}

function Col-LetraParaNum($letras) { $num = 0; foreach ($ch in $letras.ToCharArray()) { $num = $num * 26 + ([int]$ch - 64) }; return $num }

function Extrair-Celulas($sheetXml, $sharedStrings) {
    $vals = @{}
    if (-not $sheetXml -or -not $sheetXml.worksheet -or -not $sheetXml.worksheet.sheetData) { return $vals }
    $rows = $sheetXml.worksheet.sheetData.row
    if ($rows -is [xml]) { $rows = @($rows) }
    foreach ($row in $rows) {
        $r = [int]$row.r
        $cells = $row.c
        if ($cells -is [xml]) { $cells = @($cells) }
        foreach ($cell in $cells) {
            $ref = $cell.r; if (-not $ref) { continue }
            $match = [regex]::Match($ref, "^([A-Z]+)(\d+)$")
            if (-not $match.Success) { continue }
            $colNum = Col-LetraParaNum $match.Groups[1].Value
            $rowNum = [int]$match.Groups[2].Value
            $val = ""
            if ($cell.t -eq "s") { $idx = [int]$cell.v; if ($idx -lt $sharedStrings.Count) { $val = $sharedStrings[$idx] } }
            elseif ($cell.t -eq "inlineStr") { $val = $cell.is.t }
            else { $val = "$($cell.v)" }
            if ($val) { $vals["$rowNum,$colNum"] = $val.Trim() }
        }
    }
    return $vals
}

function Processar-Xlsx($arquivo) {
    $data = Ler-XlsxComoXml $arquivo.Arquivo
    if (-not $data) { return $null }
    $vals = Extrair-Celulas $data.Sheet $data.SharedStrings

    $linhaInicio = 0; $linhaFim = 0
    for ($r = 1; $r -le 80; $r++) {
        $comb = "$($vals["$r,1"]) $($vals["$r,2"]) $($vals["$r,3"]) $($vals["$r,4"])"
        if ($comb -match "VALORES\s+APURADOS" -and $linhaInicio -eq 0) { $linhaInicio = $r }
        if ($comb -match "TOTAL\s+BRUTO\s+APURADO") { $linhaFim = $r; break }
    }

    $processo = ""; $reclamante = ""; $reclamada = ""; $dataRef = ""
    for ($r = 1; $r -le 12; $r++) {
        $t = $vals["$r,4"]
        if ($t -match "PROCESSO:\s*(.+)") { $processo = $Matches[1].Trim() }
        if ($t -match "RECLAMANTE:\s*(.+)") { $reclamante = $Matches[1].Trim() }
        if ($t -match "RECLAMADA:\s*(.+)") { $reclamada = $Matches[1].Trim() }
    }
    if ($linhaInicio -gt 0) { $linhaData = $linhaInicio + 1; $dataRef = $vals["$linhaData,8"]; if (-not $dataRef) { $dataRef = $vals["$linhaData,6"] } }

    $totalBruto = ""; $correcaoTotal = ""; $totalApurado = ""
    if ($linhaFim -gt 0) {
        $totalBruto = $vals["$linhaFim,8"]
        $correcaoTotal = $vals["$linhaFim,9"]
        $totalApurado = $vals["$linhaFim,11"]
    }

    return @($arquivo.Tipo, $arquivo.Empresa, $arquivo.Reclamante, (Split-Path $arquivo.Arquivo -Leaf), $processo, $reclamante, $reclamada, $dataRef, $correcaoTotal, $totalApurado, $totalBruto)
}

function Processar-Pdf($arquivo, $pdftotext) {
    $tmpTxt = [System.IO.Path]::GetTempFileName()
    try {
        & $pdftotext -layout $arquivo.Arquivo $tmpTxt 2>$null
        $linhas = Get-Content $tmpTxt -Encoding UTF8
    } catch { return $null }
    finally { if (Test-Path $tmpTxt) { Remove-Item $tmpTxt -Force } }

    $processo = ""; $reclamante = ""; $reclamada = ""; $dataRef = ""
    $totalBruto = ""; $totalApurado = ""

    foreach ($linha in $linhas) {
        if ($linha -match "Processo:\s*(.+?)(\s|$)" -and -not $processo) { $processo = $Matches[1].Trim() }
        if ($linha -match "^Reclamante:\s*(.+)") { $reclamante = $Matches[1].Trim() }
        if ($linha -match "^Reclamado:\s*(.+)") { $reclamada = $Matches[1].Trim() }
        if ($linha -match "Data Liquida" -and $linha -match "(\d{2}/\d{2}/\d{4})") { $dataRef = $Matches[1] }
        if ($linha -match "^\s*Total\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s*$" -and -not $totalApurado) {
            $totalBruto = $Matches[1]; $totalApurado = $Matches[3]
        }
        if ($linha -match "Bruto Devido ao Reclamante\s+([\d.,]+)" -and -not $totalApurado) { $totalApurado = $Matches[1] }
    }

    return @($arquivo.Tipo, $arquivo.Empresa, $arquivo.Reclamante, (Split-Path $arquivo.Arquivo -Leaf), $processo, $reclamante, $reclamada, $dataRef, "", $totalApurado, $totalBruto)
}

try {
    $subpastas = Get-ChildItem -Path $calcBase -Directory -ErrorAction SilentlyContinue
    "Subpastas: $($subpastas.Count)" | Out-File $log -Append

    # Coleta TODOS os arquivos de todas as pastas Brasanitas
    # Prioridade: xlsx RESUMO > xlsx > PDF
    $arquivos = @()
    foreach ($sp in $subpastas) {
        $brasDirs = Get-ChildItem -Path $sp.FullName -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match "Brasa" }
        foreach ($bras in $brasDirs) {
            $reclamantes = Get-ChildItem -Path $bras.FullName -Directory -ErrorAction SilentlyContinue
            $countBras = 0
            foreach ($d in $reclamantes) {
                # Prioridade 1: xlsx RESUMO
                $file = Get-ChildItem -Path $d.FullName -Filter "*RESUMO*.xlsx" -File -ErrorAction SilentlyContinue | Select-Object -First 1
                $tipoArq = "xlsx"
                # Prioridade 2: qualquer xlsx
                if (-not $file) {
                    $file = Get-ChildItem -Path $d.FullName -Filter "*.xlsx" -File -ErrorAction SilentlyContinue | Select-Object -First 1
                }
                # Prioridade 3: PDF
                if (-not $file) {
                    $file = Get-ChildItem -Path $d.FullName -Filter "*.pdf" -File -ErrorAction SilentlyContinue | Select-Object -First 1
                    $tipoArq = "pdf"
                }
                if ($file) {
                    $arquivos += [PSCustomObject]@{
                        Tipo = $sp.Name; Empresa = $bras.Name; Reclamante = $d.Name; Arquivo = $file.FullName; TipoArq = $tipoArq
                    }
                    $countBras++
                }
            }
            "  $($sp.Name) \ $($bras.Name): $countBras arquivos" | Out-File $log -Append
        }
    }
    "Total arquivos coletados: $($arquivos.Count)" | Out-File $log -Append

    # CSV (semicolon-separated)
    $csvPath = "C:\Users\jose.madella\Desktop\Brasanitas_Completo.csv"
    $header = "Tipo (Subpasta);Empresa;Reclamante;Arquivo;Processo;Reclamante (Planilha);Reclamada;Data Referencia;correcao;total apurado;TOTAL BRUTO APURADO"
    $header | Out-File $csvPath -Encoding UTF8

    $i = 0
    foreach ($amostra in $arquivos) {
        $i++
        if ($i % 200 -eq 0) { "Progresso: $i / $($arquivos.Count)" | Out-File $log -Append }
        try {
            $row = $null
            if ($amostra.TipoArq -eq "xlsx") {
                $row = Processar-Xlsx $amostra
            } else {
                $row = Processar-Pdf $amostra $pdftotext
            }
            if ($row) {
                $csvLine = ($row | ForEach-Object { '"' + ($_ -replace '"','""') + '"' }) -join ";"
                $csvLine | Out-File $csvPath -Encoding UTF8 -Append
            }
        } catch {
            "  ERRO [$i]: $($_.Exception.Message)" | Out-File $log -Append
        }
    }
    "CSV gerado: $csvPath | Linhas: $i" | Out-File $log -Append

    # Converte CSV para XLSX
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
    taskkill /F /IM EXCEL.EXE /T 2>$null
    Start-Sleep 2
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($csvPath)
    $ws = $wb.Sheets.Item(1)
    $ws.UsedRange.EntireColumn.AutoFit() | Out-Null
    $ws.Rows.Item(1).Font.Bold = $true
    $ws.Rows.Item(1).Interior.Color = 13434828
    $xlsxPath = "C:\Users\jose.madella\Desktop\Brasanitas_Completo.xlsx"
    $wb.SaveAs($xlsxPath, 51)
    $wb.Close(); $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    "XLSX gerado: $xlsxPath" | Out-File $log -Append
    "CONCLUIDO: $(Get-Date)" | Out-File $log -Append
} catch {
    "ERRO FATAL: $($_.Exception.Message)" | Out-File $log -Append
    "STACK: $($_.ScriptStackTrace)" | Out-File $log -Append
}