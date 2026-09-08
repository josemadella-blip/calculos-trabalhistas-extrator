# Brasanitas - TODOS os casos em TODAS as subpastas de Calculos Elaborados
# Colunas A-K apenas | xlsx via ZIP/XML (Brasanitas tem xlsx, nao PDF)
# CSV + conversao unica para xlsx

$ErrorActionPreference = 'Continue'
$log = "C:\Users\jose.madella\Desktop\extrair_log_brasanitas_completo.txt"
"INICIO: $(Get-Date)" | Out-File $log -Encoding UTF8

$calcBase = "X:\Trabalhista\Arquivos Trabalhistas\C" + [char]0x00E1 + "lculos Elaborados"

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

try {
    $subpastas = Get-ChildItem -Path $calcBase -Directory -ErrorAction SilentlyContinue
    "Subpastas de Calculos Elaborados: $($subpastas.Count)" | Out-File $log -Append

    # Coleta TODOS os xlsx RESUMO de todas as pastas que contenham "Brasanitas" no nome
    $arquivos = @()
    foreach ($sp in $subpastas) {
        $brasDirs = Get-ChildItem -Path $sp.FullName -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match "Brasanitas" }
        foreach ($bras in $brasDirs) {
            $reclamantes = Get-ChildItem -Path $bras.FullName -Directory -ErrorAction SilentlyContinue
            $countBras = 0
            foreach ($d in $reclamantes) {
                $file = Get-ChildItem -Path $d.FullName -Filter "*RESUMO*.xlsx" -File -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($file) {
                    $arquivos += [PSCustomObject]@{
                        Tipo = $sp.Name; Empresa = $bras.Name; Reclamante = $d.Name; Arquivo = $file.FullName
                    }
                    $countBras++
                }
            }
            "  $($sp.Name) \ $($bras.Name): $countBras arquivos" | Out-File $log -Append
        }
    }
    "Total xlsx coletados: $($arquivos.Count)" | Out-File $log -Append

    # CSV (semicolon-separated para Excel PT-BR)
    $csvPath = "C:\Users\jose.madella\Desktop\Brasanitas_Completo.csv"
    $header = "Tipo (Subpasta);Empresa;Reclamante;Arquivo;Processo;Reclamante (Planilha);Reclamada;Data Referencia;correcao;total apurado;TOTAL BRUTO APURADO"
    $header | Out-File $csvPath -Encoding UTF8

    $i = 0
    foreach ($amostra in $arquivos) {
        $i++
        if ($i % 100 -eq 0) { "Progresso: $i / $($arquivos.Count)" | Out-File $log -Append }
        try {
            $data = Ler-XlsxComoXml $amostra.Arquivo
            if (-not $data) { "  ERRO [$i]: nao leu xlsx" | Out-File $log -Append; continue }
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

            $row = @($amostra.Tipo, $amostra.Empresa, $amostra.Reclamante, (Split-Path $amostra.Arquivo -Leaf), $processo, $reclamante, $reclamada, $dataRef, $correcaoTotal, $totalApurado, $totalBruto)
            $csvLine = ($row | ForEach-Object { '"' + ($_ -replace '"','""') + '"' }) -join ";"
            $csvLine | Out-File $csvPath -Encoding UTF8 -Append
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