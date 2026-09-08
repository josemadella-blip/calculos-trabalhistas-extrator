# Script unificado: Brasanitas (xlsx via ZIP/XML) + Solucoes (PDF via pdftotext)
# 15 amostras de cada empresa | Colunas A/B/C = Tipo/Empresa/Reclamante
# Saida: xlsx com escrita direta via COM (célula por célula)

$ErrorActionPreference = 'Continue'
$log = "C:\Users\jose.madella\Desktop\extrair_log_final.txt"
"INICIO: $(Get-Date)" | Out-File $log -Encoding UTF8

# Caminhos com acentos via Unicode
$calcBase = "X:\Trabalhista\Arquivos Trabalhistas\C" + [char]0x00E1 + "lculos Elaborados"
$tipoCalc = "Materializa" + [char]0x00E7 + [char]0x00E3 + "o de Decis" + [char]0x00E3 + "o"
$base = Join-Path $calcBase $tipoCalc
$empBrasanitas = "Brasanitas"
$empSolucoes = [string]::new([char[]](83,111,108,117,0x00E7,0x00F5,101,115,32,83,101,114,118,105,0x00E7,111,115,32,84,101,114,99,101,105,114,105,122,97,100,111,115))
$empresas = @($empBrasanitas, $empSolucoes)
$pdftotext = "C:\Users\jose.madella\AppData\Local\Microsoft\WinGet\Packages\oschwartz10612.Poppler_Microsoft.Winget.Source_8wekyb3d8bbwe\poppler-25.07.0\Library\bin\pdftotext.exe"

# ===== FUNCOES XLSX (ZIP/XML) =====
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

function Processar-Brasanitas($arquivo, $log) {
    $data = Ler-XlsxComoXml $arquivo.Arquivo
    if (-not $data) { "  ERRO: nao leu xlsx" | Out-File $log -Append; return $null }
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
    
    $rubricasArquivo = @{}
    if ($linhaInicio -gt 0 -and $linhaFim -gt 0) {
        for ($r = $linhaInicio + 2; $r -lt $linhaFim; $r++) {
            $nomeRub = $vals["$r,4"]
            if ($nomeRub -and $nomeRub -notmatch "^\s*$") {
                $nomeRub = $nomeRub.Trim()
                $rubricasArquivo[$nomeRub] = [PSCustomObject]@{
                    Principal = $vals["$r,8"]; Correcao = $vals["$r,9"]; Juros = $vals["$r,10"]; Total = $vals["$r,11"]
                }
            }
        }
    }
    
    $totalBruto = ""; $correcaoTotal = ""; $totalApurado = ""
    if ($linhaFim -gt 0) { $totalBruto = $vals["$linhaFim,8"]; $correcaoTotal = $vals["$linhaFim,9"]; $totalApurado = $vals["$linhaFim,11"] }
    
    return [PSCustomObject]@{
        Tipo = $arquivo.Tipo; Empresa = $arquivo.Empresa; ReclamantePasta = $arquivo.Reclamante
        Arquivo = (Split-Path $arquivo.Arquivo -Leaf); Processo = $processo; Reclamante = $reclamante
        Reclamada = $reclamada; DataReferencia = $dataRef; CorrecaoTotal = $correcaoTotal
        TotalApurado = $totalApurado; TotalBrutoApurado = $totalBruto; Rubricas = $rubricasArquivo
    }
}

# ===== FUNCOES PDF (pdftotext) =====
function Processar-Solucoes($arquivo, $pdftotext, $log) {
    $tmpTxt = [System.IO.Path]::GetTempFileName()
    try {
        & $pdftotext -layout $arquivo.Arquivo $tmpTxt 2>$null
        $linhas = Get-Content $tmpTxt -Encoding UTF8
    } catch { "  ERRO pdftotext: $($_.Exception.Message)" | Out-File $log -Append; return $null }
    finally { if (Test-Path $tmpTxt) { Remove-Item $tmpTxt -Force } }
    
    # Extrai dados do cabecalho
    $processo = ""; $reclamante = ""; $reclamada = ""; $dataRef = ""
    foreach ($linha in $linhas) {
        if ($linha -match "Processo:\s*(.+?)(\s|$)" -and -not $processo) { $processo = $Matches[1].Trim() }
        if ($linha -match "^Reclamante:\s*(.+)") { $reclamante = $Matches[1].Trim() }
        if ($linha -match "^Reclamado:\s*(.+)") { $reclamada = $Matches[1].Trim() }
        if ($linha -match "Data Liquida" -and $linha -match "(\d{2}/\d{2}/\d{4})") { $dataRef = $Matches[1] }
    }
    
    # Encontra a secao "Resumo do Calculo" e extrai rubricas
    # Formato: DESCRICAO | Valor Corrigido | Juros | Total
    $rubricasArquivo = @{}
    $totalBruto = ""; $correcaoTotal = ""; $totalApurado = ""
    $inResumo = $false
    $inRubricas = $false
    
    foreach ($linha in $linhas) {
        $trim = $linha.Trim()
        if ($trim -match "Resumo do C" -and $trim -match "lculo") { $inResumo = $true; continue }
        if (-not $inResumo) { continue }
        
        # Cabecalho da tabela de rubricas
        if ($trim -match "Descri.*Bruto Devido") { $inRubricas = $true; continue }
        if (-not $inRubricas) { continue }
        
        # Linha Total
        if ($trim -match "^\s*Total\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s*$") {
            $totalBruto = $Matches[1]      # Valor Corrigido total
            $correcaoTotal = ""             # PDF nao separa correcao
            $totalApurado = $Matches[3]     # Total
            $inRubricas = $false; $inResumo = $false; break
        }
        
        # Rubrica: DESCRICAO + 3 valores numericos
        # Padrao: texto no inicio, depois 3 numeros (pode ter - para zero)
        if ($trim -match "^(.+?)\s+([\d.]+,\d{2}|-)\s+([\d.]+,\d{2}|-)\s+([\d.]+,\d{2}|-)\s*$") {
            $nomeRub = $Matches[1].Trim()
            $valCorrigido = $Matches[2]
            $juros = $Matches[3]
            $total = $Matches[4]
            if ($valCorrigido -eq "-") { $valCorrigido = "" }
            if ($juros -eq "-") { $juros = "" }
            if ($total -eq "-") { $total = "" }
            $rubricasArquivo[$nomeRub] = [PSCustomObject]@{
                Principal = $valCorrigido; Correcao = ""; Juros = $juros; Total = $total
            }
        }
    }
    
    # Se nao achou total na secao Resumo, procura "Bruto Devido ao Reclamante"
    if (-not $totalApurado) {
        foreach ($linha in $linhas) {
            if ($linha -match "Bruto Devido ao Reclamante\s+([\d.,]+)") { $totalApurado = $Matches[1]; break }
        }
    }
    
    return [PSCustomObject]@{
        Tipo = $arquivo.Tipo; Empresa = $arquivo.Empresa; ReclamantePasta = $arquivo.Reclamante
        Arquivo = (Split-Path $arquivo.Arquivo -Leaf); Processo = $processo; Reclamante = $reclamante
        Reclamada = $reclamada; DataReferencia = $dataRef; CorrecaoTotal = $correcaoTotal
        TotalApurado = $totalApurado; TotalBrutoApurado = $totalBruto; Rubricas = $rubricasArquivo
    }
}

# ===== COLETA DE ARQUIVOS =====
try {
    "Base: $base | Existe: $(Test-Path $base)" | Out-File $log -Append
    
    $arquivos = @()
    foreach ($emp in $empresas) {
        $empPath = Join-Path $base $emp
        "Empresa: $emp | Existe: $(Test-Path $empPath)" | Out-File $log -Append
        if (Test-Path $empPath) {
            $countEmp = 0
            $dirs = Get-ChildItem -Path $empPath -Directory -ErrorAction SilentlyContinue
            "  Dirs: $($dirs.Count)" | Out-File $log -Append
            foreach ($d in $dirs) {
                if ($countEmp -ge 15) { break }
                # Brasanitas: procura xlsx RESUMO | Solucoes: procura PDF
                if ($emp -eq "Brasanitas") {
                    $file = Get-ChildItem -Path $d.FullName -Filter "*RESUMO*.xlsx" -File -ErrorAction SilentlyContinue | Select-Object -First 1
                } else {
                    $file = Get-ChildItem -Path $d.FullName -Filter "*.pdf" -File -ErrorAction SilentlyContinue | Select-Object -First 1
                }
                if ($file) {
                    $arquivos += [PSCustomObject]@{ Tipo = $tipoCalc; Empresa = $emp; Reclamante = $d.Name; Arquivo = $file.FullName }
                    $countEmp++
                }
            }
            "  Coletados: $countEmp" | Out-File $log -Append
        }
    }
    "Total: $($arquivos.Count) arquivos" | Out-File $log -Append

    # ===== PROCESSAMENTO =====
    $todasRubricas = New-Object System.Collections.Generic.List[string]
    $linhas = New-Object System.Collections.Generic.List[object]
    
    $i = 0
    foreach ($amostra in $arquivos) {
        $i++
        "[$i] $($amostra.Empresa) \ $($amostra.Reclamante)" | Out-File $log -Append
        try {
            $resultado = $null
            if ($amostra.Empresa -eq "Brasanitas") {
                $resultado = Processar-Brasanitas $amostra $log
            } else {
                $resultado = Processar-Solucoes $amostra $pdftotext $log
            }
            if ($resultado) {
                foreach ($k in $resultado.Rubricas.Keys) { if (-not $todasRubricas.Contains($k)) { $todasRubricas.Add($k) } }
                $linhas.Add($resultado)
                "  OK - Rubricas: $($resultado.Rubricas.Count)" | Out-File $log -Append
            }
        } catch {
            "  ERRO: $($_.Exception.Message)" | Out-File $log -Append
        }
    }
    
    "Montando planilha. Rubricas distintas: $($todasRubricas.Count) | Linhas: $($linhas.Count)" | Out-File $log -Append
    
    # ===== GERACAO XLSX VIA COM =====
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wbOut = $excel.Workbooks.Add()
    $wsOut = $wbOut.Sheets.Item(1)
    $wsOut.Name = "Consolidado"
    
    # Cabecalho
    $col = 1
    $headers = @("Tipo (Subpasta)","Empresa","Reclamante","Arquivo","Processo","Reclamante (Planilha)","Reclamada","Data Referencia","correcao","total apurado","TOTAL BRUTO APURADO")
    foreach ($h in $headers) { $wsOut.Cells.Item(1, $col).Value2 = $h; $col++ }
    
    $colunasRub = @()
    foreach ($rub in $todasRubricas) {
        $wsOut.Cells.Item(1, $col).Value2 = "$rub - Principal"; $colunasRub += ,@($rub, "Principal"); $col++
        $wsOut.Cells.Item(1, $col).Value2 = "$rub - Correcao";  $colunasRub += ,@($rub, "Correcao");  $col++
        $wsOut.Cells.Item(1, $col).Value2 = "$rub - Juros";     $colunasRub += ,@($rub, "Juros");     $col++
        $wsOut.Cells.Item(1, $col).Value2 = "$rub - Total";     $colunasRub += ,@($rub, "Total");     $col++
    }
    
    # Dados
    $rowOut = 2
    foreach ($linha in $linhas) {
        $c = 1
        $wsOut.Cells.Item($rowOut, $c).Value2 = $linha.Tipo; $c++
        $wsOut.Cells.Item($rowOut, $c).Value2 = $linha.Empresa; $c++
        $wsOut.Cells.Item($rowOut, $c).Value2 = $linha.ReclamantePasta; $c++
        $wsOut.Cells.Item($rowOut, $c).Value2 = $linha.Arquivo; $c++
        $wsOut.Cells.Item($rowOut, $c).Value2 = $linha.Processo; $c++
        $wsOut.Cells.Item($rowOut, $c).Value2 = $linha.Reclamante; $c++
        $wsOut.Cells.Item($rowOut, $c).Value2 = $linha.Reclamada; $c++
        $wsOut.Cells.Item($rowOut, $c).Value2 = $linha.DataReferencia; $c++
        $wsOut.Cells.Item($rowOut, $c).Value2 = $linha.CorrecaoTotal; $c++
        $wsOut.Cells.Item($rowOut, $c).Value2 = $linha.TotalApurado; $c++
        $wsOut.Cells.Item($rowOut, $c).Value2 = $linha.TotalBrutoApurado; $c++
        foreach ($cr in $colunasRub) {
            $rub = $cr[0]; $campo = $cr[1]
            $val = ""
            if ($linha.Rubricas.ContainsKey($rub)) { $val = $linha.Rubricas[$rub].$campo }
            $wsOut.Cells.Item($rowOut, $c).Value2 = $val
            $c++
        }
        $rowOut++
    }
    
    $wsOut.UsedRange.EntireColumn.AutoFit() | Out-Null
    $wsOut.Rows.Item(1).Font.Bold = $true
    $wsOut.Rows.Item(1).Interior.Color = 13434828
    
    $outPath = "C:\Users\jose.madella\Desktop\Resumo_Rubricas_30amostras.xlsx"
    $wbOut.SaveAs($outPath, 51)
    $wbOut.Close()
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    
    "XLSX gerado: $outPath" | Out-File $log -Append
    "CONCLUIDO: $(Get-Date)" | Out-File $log -Append
} catch {
    "ERRO FATAL: $($_.Exception.Message)" | Out-File $log -Append
    "STACK: $($_.ScriptStackTrace)" | Out-File $log -Append
    if ($excel) { try { $excel.Quit() } catch {} }
}
