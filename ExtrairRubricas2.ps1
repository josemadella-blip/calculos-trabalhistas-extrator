$ErrorActionPreference = 'Continue'
$log = "C:\Users\jose.madella\Desktop\extrair_log.txt"
"INICIO: $(Get-Date)" | Out-File $log -Encoding UTF8

# Caminho com acentos - construido via Unicode para evitar problemas de codepage
$base = "X:\Trabalhista\Arquivos Trabalhistas\C" + [char]0x00E1 + "lculos Elaborados\Materializa" + [char]0x00E7 + [char]0x00E3 + "o de Decis" + [char]0x00E3 + "o\Brasanitas"

try {
    "Base: $base" | Out-File $log -Append
    "Existe base: $(Test-Path $base)" | Out-File $log -Append

    $dirs = Get-ChildItem -Path $base -Directory -ErrorAction Stop
    "Dirs encontrados: $($dirs.Count)" | Out-File $log -Append

    $arquivos = @()
    foreach ($d in $dirs) {
        $xlsx = Get-ChildItem -Path $d.FullName -Filter "*RESUMO*.xlsx" -File -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($xlsx) { $arquivos += [PSCustomObject]@{ Reclamante = $d.Name; Arquivo = $xlsx.FullName } }
    }
    "Arquivos RESUMO encontrados: $($arquivos.Count)" | Out-File $log -Append

    $amostras = $arquivos | Select-Object -First 30

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    "Excel COM aberto" | Out-File $log -Append

    $wbOut = $excel.Workbooks.Add()
    $wsOut = $wbOut.Sheets.Item(1)
    $wsOut.Name = "Rubricas"

    $excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable - desabilita macros
    $excel.AskToUpdateLinks = $false
    $excel.AlertBeforeOverwriting = $false

    $todasRubricas = New-Object System.Collections.Generic.List[string]
    $linhas = New-Object System.Collections.Generic.List[object]

    $i = 0
    foreach ($amostra in $amostras) {
        $i++
        "[$i] Processando: $($amostra.Reclamante)" | Out-File $log -Append
        try {
            # Abre com UpdateLinks=0 e ReadOnly=true para evitar dialogos
            $wb = $excel.Workbooks.Open($amostra.Arquivo, 0, $true, 0, "", "", $false, "", "", $false, $false)
            $ws = $wb.Sheets.Item(1)
            $used = $ws.UsedRange
            $rows = $used.Rows.Count
            $cols = $used.Columns.Count
            "  Linhas: $rows, Colunas: $cols" | Out-File $log -Append

            # Leitura em bloco (array 2D) - muito mais rapido que celula por celula
            $arr = $used.Value2
            $vals = @{}
            for ($r = 1; $r -le $rows; $r++) {
                for ($c = 1; $c -le $cols; $c++) {
                    $v = $arr[$r, $c]
                    if ($v) { $vals["$r,$c"] = "$v".Trim() }
                }
            }

            $linhaInicio = 0
            $linhaFim = 0
            for ($r = 1; $r -le $rows; $r++) {
                $combinado = "$($vals["$r,1"]) $($vals["$r,2"]) $($vals["$r,3"]) $($vals["$r,4"])"
                if ($combinado -match "VALORES\s+APURADOS" -and $linhaInicio -eq 0) { $linhaInicio = $r }
                if ($combinado -match "TOTAL\s+BRUTO\s+APURADO") { $linhaFim = $r; break }
            }
            "  LinhaInicio: $linhaInicio, LinhaFim: $linhaFim" | Out-File $log -Append

            $processo = ""; $reclamante = ""; $reclamada = ""; $dataRef = ""
            for ($r = 1; $r -le [Math]::Min($rows, 12); $r++) {
                $t = $vals["$r,4"]
                if ($t -match "PROCESSO:\s*(.+)") { $processo = $Matches[1].Trim() }
                if ($t -match "RECLAMANTE:\s*(.+)") { $reclamante = $Matches[1].Trim() }
                if ($t -match "RECLAMADA:\s*(.+)") { $reclamada = $Matches[1].Trim() }
            }
            if ($linhaInicio -gt 0) {
                $linhaData = $linhaInicio + 1
                $dataRef = $vals["$linhaData,8"]
                if (-not $dataRef) { $dataRef = $vals["$linhaData,6"] }
            }

            $rubricasArquivo = @{}
            if ($linhaInicio -gt 0 -and $linhaFim -gt 0) {
                for ($r = $linhaInicio + 2; $r -lt $linhaFim; $r++) {
                    $nomeRub = $vals["$r,4"]
                    if ($nomeRub -and $nomeRub -notmatch "^\s*$") {
                        $nomeRub = $nomeRub.Trim()
                        $rubricasArquivo[$nomeRub] = [PSCustomObject]@{
                            Principal = $vals["$r,8"]
                            Correcao  = $vals["$r,9"]
                            Juros     = $vals["$r,10"]
                            Total     = $vals["$r,11"]
                        }
                        if (-not $todasRubricas.Contains($nomeRub)) { $todasRubricas.Add($nomeRub) }
                    }
                }
            }
            "  Rubricas neste arquivo: $($rubricasArquivo.Count)" | Out-File $log -Append

            # Na linha TOTAL BRUTO APURADO: col8=Principal, col9=Correcao, col10=Juros, col11=Total Apurado
            $totalBruto = ""
            $correcaoTotal = ""
            $totalApurado = ""
            if ($linhaFim -gt 0) {
                $totalBruto    = $vals["$linhaFim,8"]   # Principal Histórico
                $correcaoTotal = $vals["$linhaFim,9"]   # Correção Monetária
                $totalApurado  = $vals["$linhaFim,11"]  # Total Apurado
            }

            $linhas.Add([PSCustomObject]@{
                ReclamantePasta = $amostra.Reclamante
                Arquivo         = (Split-Path $amostra.Arquivo -Leaf)
                Processo        = $processo
                Reclamante      = $reclamante
                Reclamada       = $reclamada
                DataReferencia  = $dataRef
                CorrecaoTotal   = $correcaoTotal
                TotalApurado    = $totalApurado
                TotalBrutoApurado = $totalBruto
                Rubricas        = $rubricasArquivo
            })
            $wb.Close($false)
        } catch {
            "  ERRO: $($_.Exception.Message)" | Out-File $log -Append
            $linhas.Add([PSCustomObject]@{
                ReclamantePasta = $amostra.Reclamante
                Arquivo         = (Split-Path $amostra.Arquivo -Leaf)
                Processo        = "ERRO"
                Reclamante      = ""
                Reclamada       = ""
                DataReferencia  = ""
                CorrecaoTotal   = ""
                TotalApurado    = ""
                TotalBrutoApurado = ""
                Rubricas        = @{}
            })
        }
    }

    "Montando planilha de saida. Rubricas distintas: $($todasRubricas.Count)" | Out-File $log -Append

    $col = 1
    $wsOut.Cells.Item(1, $col).Value2 = "Reclamante (Pasta)"; $col++
    $wsOut.Cells.Item(1, $col).Value2 = "Arquivo"; $col++
    $wsOut.Cells.Item(1, $col).Value2 = "Processo"; $col++
    $wsOut.Cells.Item(1, $col).Value2 = "Reclamante"; $col++
    $wsOut.Cells.Item(1, $col).Value2 = "Reclamada"; $col++
    $wsOut.Cells.Item(1, $col).Value2 = "Data Referencia"; $col++
    $wsOut.Cells.Item(1, $col).Value2 = "correcao"; $col++
    $wsOut.Cells.Item(1, $col).Value2 = "total apurado"; $col++
    $wsOut.Cells.Item(1, $col).Value2 = "TOTAL BRUTO APURADO"; $colTotalBruto = $col; $col++

    $colunasRubrica = @()
    foreach ($rub in $todasRubricas) {
        $wsOut.Cells.Item(1, $col).Value2 = "$rub - Principal"; $colunasRubrica += ,@($rub, "Principal"); $col++
        $wsOut.Cells.Item(1, $col).Value2 = "$rub - Correcao";  $colunasRubrica += ,@($rub, "Correcao");  $col++
        $wsOut.Cells.Item(1, $col).Value2 = "$rub - Juros";     $colunasRubrica += ,@($rub, "Juros");     $col++
        $wsOut.Cells.Item(1, $col).Value2 = "$rub - Total";     $colunasRubrica += ,@($rub, "Total");     $col++
    }

    $rowOut = 2
    foreach ($linha in $linhas) {
        $c = 1
        $wsOut.Cells.Item($rowOut, $c).Value2 = $linha.ReclamantePasta; $c++
        $wsOut.Cells.Item($rowOut, $c).Value2 = $linha.Arquivo; $c++
        $wsOut.Cells.Item($rowOut, $c).Value2 = $linha.Processo; $c++
        $wsOut.Cells.Item($rowOut, $c).Value2 = $linha.Reclamante; $c++
        $wsOut.Cells.Item($rowOut, $c).Value2 = $linha.Reclamada; $c++
        $wsOut.Cells.Item($rowOut, $c).Value2 = $linha.DataReferencia; $c++
        $wsOut.Cells.Item($rowOut, $c).Value2 = $linha.CorrecaoTotal; $c++
        $wsOut.Cells.Item($rowOut, $c).Value2 = $linha.TotalApurado; $c++
        $wsOut.Cells.Item($rowOut, $colTotalBruto).Value2 = $linha.TotalBrutoApurado; $c++
        foreach ($cr in $colunasRubrica) {
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

    "CONCLUIDO: $outPath" | Out-File $log -Append
    "FIM: $(Get-Date)" | Out-File $log -Append
} catch {
    "ERRO FATAL: $($_.Exception.Message)" | Out-File $log -Append
    "STACK: $($_.ScriptStackTrace)" | Out-File $log -Append
    if ($excel) { try { $excel.Quit() } catch {} }
}
