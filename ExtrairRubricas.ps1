# Extrai rubricas entre "VALORES APURADOS" e "TOTAL BRUTO APURADO" dos arquivos RESUMO ANALÍTICO.xlsx
# Gera um xlsx consolidado com 30 exemplos (uma linha por arquivo, colunas = rubricas)

$ErrorActionPreference = 'Stop'
$base = "X:\Trabalhista\Arquivos Trabalhistas\Cálculos Elaborados\Materialização de Decisão\Brasanitas"
$outPath = "C:\Users\jose.madella\Desktop\Resumo_Rubricas_30amostras.xlsx"

# Coleta arquivos RESUMO ANALÍTICO.xlsx
$dirs = Get-ChildItem -Path $base -Directory
$arquivos = @()
foreach ($d in $dirs) {
    $xlsx = Get-ChildItem -Path $d.FullName -Filter "*RESUMO*.xlsx" -File -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($xlsx) { $arquivos += [PSCustomObject]@{ Reclamante = $d.Name; Arquivo = $xlsx.FullName } }
}
$amostras = $arquivos | Select-Object -First 30
Write-Host ("Total de arquivos encontrados: " + $arquivos.Count)
Write-Host ("Amostras selecionadas: " + $amostras.Count)

# Abre Excel via COM
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# Workbook de saída
$wbOut = $excel.Workbooks.Add()
$wsOut = $wbOut.Sheets.Item(1)
$wsOut.Name = "Rubricas"

# Conjunto ordenado de todas as rubricas encontradas (para criar colunas dinamicamente)
$todasRubricas = New-Object System.Collections.Generic.List[string]
$linhas = New-Object System.Collections.Generic.List[object]

foreach ($amostra in $amostras) {
    Write-Host ("Processando: " + $amostra.Reclamante)
    try {
        $wb = $excel.Workbooks.Open($amostra.Arquivo, 0, $true)
        $ws = $wb.Sheets.Item(1)
        $used = $ws.UsedRange
        $rows = $used.Rows.Count
        $cols = $used.Columns.Count

        # Lê tudo para uma matriz
        $vals = @{}
        for ($r = 1; $r -le $rows; $r++) {
            for ($c = 1; $c -le $cols; $c++) {
                $v = $used.Cells.Item($r, $c).Text
                if ($v) { $vals["$r,$c"] = $v.Trim() }
            }
        }

        # Localiza linha de "VALORES APURADOS" e "TOTAL BRUTO APURADO"
        $linhaInicio = 0
        $linhaFim = 0
        for ($r = 1; $r -le $rows; $r++) {
            $txtCol3 = $vals["$r,3"]
            $txtCol4 = $vals["$r,4"]
            $txtCol1 = $vals["$r,1"]
            $txtCol2 = $vals["$r,2"]
            $combinado = ($txtCol1 + " " + $txtCol2 + " " + $txtCol3 + " " + $txtCol4)
            if ($combinado -match "VALORES\s+APURADOS" -and $linhaInicio -eq 0) {
                $linhaInicio = $r
            }
            if ($combinado -match "TOTAL\s+BRUTO\s+APURADO") {
                $linhaFim = $r
                break
            }
        }

        # Dados do processo (cabeçalho)
        $processo = ""; $reclamante = ""; $reclamada = ""; $dataRef = ""
        for ($r = 1; $r -le [Math]::Min($rows, 12); $r++) {
            $t = $vals["$r,4"]
            if ($t -match "PROCESSO:\s*(.+)") { $processo = $Matches[1].Trim() }
            if ($t -match "RECLAMANTE:\s*(.+)") { $reclamante = $Matches[1].Trim() }
            if ($t -match "RECLAMADA:\s*(.+)") { $reclamada = $Matches[1].Trim() }
        }

        # Data de referência: linha logo após VALORES APURADOS (cabeçalho de datas)
        if ($linhaInicio -gt 0) {
            $linhaData = $linhaInicio + 1
            $dataRef = $vals["$linhaData,8"]
            if (-not $dataRef) { $dataRef = $vals["$linhaData,6"] }
        }

        # Extrai rubricas: linhas entre linhaInicio+2 e linhaFim-1
        # Coluna 4 = nome da rubrica, coluna 8 = Principal, 9 = Correção, 10 = Juros, 11 = Total Apurado
        $rubricasArquivo = @{}
        if ($linhaInicio -gt 0 -and $linhaFim -gt 0) {
            for ($r = $linhaInicio + 2; $r -lt $linhaFim; $r++) {
                $nomeRub = $vals["$r,4"]
                if ($nomeRub -and $nomeRub -notmatch "^\s*$") {
                    $nomeRub = $nomeRub.Trim()
                    $principal = $vals["$r,8"]
                    $correcao  = $vals["$r,9"]
                    $juros     = $vals["$r,10"]
                    $total     = $vals["$r,11"]
                    $rubricasArquivo[$nomeRub] = [PSCustomObject]@{
                        Principal = $principal
                        Correcao  = $correcao
                        Juros     = $juros
                        Total     = $total
                    }
                    if (-not $todasRubricas.Contains($nomeRub)) {
                        $todasRubricas.Add($nomeRub)
                    }
                }
            }
        }

        # Total Bruto Apurado
        $totalBruto = ""
        if ($linhaFim -gt 0) {
            $totalBruto = $vals["$linhaFim,11"]
            if (-not $totalBruto) { $totalBruto = $vals["$linhaFim,8"] }
        }

        $linha = [PSCustomObject]@{
            ReclamantePasta = $amostra.Reclamante
            Arquivo         = (Split-Path $amostra.Arquivo -Leaf)
            Processo        = $processo
            Reclamante      = $reclamante
            Reclamada       = $reclamada
            DataReferencia  = $dataRef
            Rubricas        = $rubricasArquivo
            TotalBrutoApurado = $totalBruto
        }
        $linhas.Add($linha)

        $wb.Close($false)
    } catch {
        Write-Host ("ERRO em " + $amostra.Reclamante + ": " + $_.Exception.Message)
        $linhas.Add([PSCustomObject]@{
            ReclamantePasta = $amostra.Reclamante
            Arquivo         = (Split-Path $amostra.Arquivo -Leaf)
            Processo        = "ERRO: " + $_.Exception.Message
            Reclamante      = ""
            Reclamada       = ""
            DataReferencia  = ""
            Rubricas        = @{}
            TotalBrutoApurado = ""
        })
    }
}

# Monta cabeçalho da planilha de saída
# Colunas fixas + 4 colunas (Principal/Correção/Juros/Total) por rubrica + Total Bruto
$col = 1
$wsOut.Cells.Item(1, $col).Value2 = "Reclamante (Pasta)"; $col++
$wsOut.Cells.Item(1, $col).Value2 = "Arquivo"; $col++
$wsOut.Cells.Item(1, $col).Value2 = "Processo"; $col++
$wsOut.Cells.Item(1, $col).Value2 = "Reclamante"; $col++
$wsOut.Cells.Item(1, $col).Value2 = "Reclamada"; $col++
$wsOut.Cells.Item(1, $col).Value2 = "Data Referência"; $col++

$colunasRubrica = @()
foreach ($rub in $todasRubricas) {
    $wsOut.Cells.Item(1, $col).Value2 = "$rub - Principal"; $colunasRubrica += ,@($rub, "Principal"); $col++
    $wsOut.Cells.Item(1, $col).Value2 = "$rub - Correção";  $colunasRubrica += ,@($rub, "Correcao");  $col++
    $wsOut.Cells.Item(1, $col).Value2 = "$rub - Juros";     $colunasRubrica += ,@($rub, "Juros");     $col++
    $wsOut.Cells.Item(1, $col).Value2 = "$rub - Total";     $colunasRubrica += ,@($rub, "Total");     $col++
}
$wsOut.Cells.Item(1, $col).Value2 = "TOTAL BRUTO APURADO"; $colTotalBruto = $col

# Preenche linhas
$rowOut = 2
foreach ($linha in $linhas) {
    $c = 1
    $wsOut.Cells.Item($rowOut, $c).Value2 = $linha.ReclamantePasta; $c++
    $wsOut.Cells.Item($rowOut, $c).Value2 = $linha.Arquivo; $c++
    $wsOut.Cells.Item($rowOut, $c).Value2 = $linha.Processo; $c++
    $wsOut.Cells.Item($rowOut, $c).Value2 = $linha.Reclamante; $c++
    $wsOut.Cells.Item($rowOut, $c).Value2 = $linha.Reclamada; $c++
    $wsOut.Cells.Item($rowOut, $c).Value2 = $linha.DataReferencia; $c++

    foreach ($cr in $colunasRubrica) {
        $rub = $cr[0]; $campo = $cr[1]
        $val = ""
        if ($linha.Rubricas.ContainsKey($rub)) {
            $val = $linha.Rubricas[$rub].$campo
        }
        $wsOut.Cells.Item($rowOut, $c).Value2 = $val
        $c++
    }
    $wsOut.Cells.Item($rowOut, $colTotalBruto).Value2 = $linha.TotalBrutoApurado
    $rowOut++
}

# Formatação
$wsOut.UsedRange.EntireColumn.AutoFit() | Out-Null
$wsOut.Rows.Item(1).Font.Bold = $true
$wsOut.Rows.Item(1).Interior.Color = 13434828  # verde claro

$wbOut.SaveAs($outPath, 51)  # xlsx
$wbOut.Close()
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host ""
Write-Host ("CONCLUIDO. Arquivo gerado: " + $outPath)
Write-Host ("Rubricas distintas encontradas: " + $todasRubricas.Count)
Write-Host ("Linhas processadas: " + $linhas.Count)
