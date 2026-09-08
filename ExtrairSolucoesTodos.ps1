# Script: Todos os casos da Solucoes (PDF via pdftotext)
# Saida: xlsx com todas as rubricas, uma linha por caso

$ErrorActionPreference = 'Continue'
$log = "C:\Users\jose.madella\Desktop\extrair_log_solucoes_todos.txt"
"INICIO: $(Get-Date)" | Out-File $log -Encoding UTF8

$calcBase = "X:\Trabalhista\Arquivos Trabalhistas\C" + [char]0x00E1 + "lculos Elaborados"
$tipoCalc = "Materializa" + [char]0x00E7 + [char]0x00E3 + "o de Decis" + [char]0x00E3 + "o"
$base = Join-Path $calcBase $tipoCalc
$empSolucoes = [string]::new([char[]](83,111,108,117,0x00E7,0x00F5,101,115,32,83,101,114,118,105,0x00E7,111,115,32,84,101,114,99,101,105,114,105,122,97,100,111,115))
$pdftotext = "C:\Users\jose.madella\AppData\Local\Microsoft\WinGet\Packages\oschwartz10612.Poppler_Microsoft.Winget.Source_8wekyb3d8bbwe\poppler-25.07.0\Library\bin\pdftotext.exe"

function Processar-Solucoes($arquivo, $pdftotext, $log) {
    $tmpTxt = [System.IO.Path]::GetTempFileName()
    try {
        & $pdftotext -layout $arquivo.Arquivo $tmpTxt 2>$null
        $linhas = Get-Content $tmpTxt -Encoding UTF8
    } catch { return $null }
    finally { if (Test-Path $tmpTxt) { Remove-Item $tmpTxt -Force } }
    
    $processo = ""; $reclamante = ""; $reclamada = ""; $dataRef = ""
    foreach ($linha in $linhas) {
        if ($linha -match "Processo:\s*(.+?)(\s|$)" -and -not $processo) { $processo = $Matches[1].Trim() }
        if ($linha -match "^Reclamante:\s*(.+)") { $reclamante = $Matches[1].Trim() }
        if ($linha -match "^Reclamado:\s*(.+)") { $reclamada = $Matches[1].Trim() }
        if ($linha -match "Data Liquida" -and $linha -match "(\d{2}/\d{2}/\d{4})") { $dataRef = $Matches[1] }
    }
    
    $rubricasArquivo = @{}
    $totalBruto = ""; $correcaoTotal = ""; $totalApurado = ""
    $inResumo = $false; $inRubricas = $false
    
    foreach ($linha in $linhas) {
        $trim = $linha.Trim()
        if ($trim -match "Resumo do C" -and $trim -match "lculo") { $inResumo = $true; continue }
        if (-not $inResumo) { continue }
        if ($trim -match "Descri.*Bruto Devido") { $inRubricas = $true; continue }
        if (-not $inRubricas) { continue }
        if ($trim -match "^\s*Total\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s*$") {
            $totalBruto = $Matches[1]; $totalApurado = $Matches[3]
            $inRubricas = $false; $inResumo = $false; break
        }
        if ($trim -match "^(.+?)\s+([\d.]+,\d{2}|-)\s+([\d.]+,\d{2}|-)\s+([\d.]+,\d{2}|-)\s*$") {
            $nomeRub = $Matches[1].Trim()
            $valCorrigido = $Matches[2]; $juros = $Matches[3]; $total = $Matches[4]
            if ($valCorrigido -eq "-") { $valCorrigido = "" }
            if ($juros -eq "-") { $juros = "" }
            if ($total -eq "-") { $total = "" }
            $rubricasArquivo[$nomeRub] = [PSCustomObject]@{
                Principal = $valCorrigido; Correcao = ""; Juros = $juros; Total = $total
            }
        }
    }
    
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

try {
    $empPath = Join-Path $base $empSolucoes
    "Empresa: $empSolucoes | Existe: $(Test-Path $empPath)" | Out-File $log -Append
    
    # Coleta TODOS os arquivos PDF (sem limite)
    $arquivos = @()
    $dirs = Get-ChildItem -Path $empPath -Directory -ErrorAction SilentlyContinue
    "Dirs: $($dirs.Count)" | Out-File $log -Append
    foreach ($d in $dirs) {
        $file = Get-ChildItem -Path $d.FullName -Filter "*.pdf" -File -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($file) {
            $arquivos += [PSCustomObject]@{ Tipo = $tipoCalc; Empresa = "Solucoes"; Reclamante = $d.Name; Arquivo = $file.FullName }
        }
    }
    "Total arquivos PDF: $($arquivos.Count)" | Out-File $log -Append

    $todasRubricas = New-Object System.Collections.Generic.List[string]
    $linhas = New-Object System.Collections.Generic.List[object]
    
    $i = 0
    foreach ($amostra in $arquivos) {
        $i++
        if ($i % 50 -eq 0) { "Progresso: $i / $($arquivos.Count)" | Out-File $log -Append }
        try {
            $resultado = Processar-Solucoes $amostra $pdftotext $log
            if ($resultado) {
                foreach ($k in $resultado.Rubricas.Keys) { if (-not $todasRubricas.Contains($k)) { $todasRubricas.Add($k) } }
                $linhas.Add($resultado)
            }
        } catch {
            "  ERRO [$i]: $($_.Exception.Message)" | Out-File $log -Append
        }
    }
    
    "Processamento concluido. Linhas: $($linhas.Count) | Rubricas: $($todasRubricas.Count)" | Out-File $log -Append
    
    # Geracao XLSX via COM
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wbOut = $excel.Workbooks.Add()
    $wsOut = $wbOut.Sheets.Item(1)
    $wsOut.Name = "Solucoes"
    
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
    
    $outPath = "C:\Users\jose.madella\Desktop\Solucoes_Todos_Casos.xlsx"
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
