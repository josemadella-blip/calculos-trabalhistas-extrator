# Extrai rubricas entre "VALORES APURADOS" e "TOTAL BRUTO APURADO"
# Colunas: A=Tipo (subpasta de Calculos Elaborados), B=Empresa, C=Reclamante
# Filtro: apenas Brasanitas e Solucoes Servicos Terceirizados
# 30 amostras (15 de cada empresa)
# Tabela separada por empresa (3 abas: Consolidado, Brasanitas, Solucoes)

$ErrorActionPreference = 'Continue'
$log = "C:\Users\jose.madella\Desktop\extrair_log.txt"
"INICIO: $(Get-Date)" | Out-File $log -Encoding UTF8

# Caminho base com acentos via Unicode escape
$calcBase = "X:\Trabalhista\Arquivos Trabalhistas\C" + [char]0x00E1 + "lculos Elaborados"
$tipoCalc = "Materializa" + [char]0x00E7 + [char]0x00E3 + "o de Decis" + [char]0x00E3 + "o"
$base = Join-Path $calcBase $tipoCalc

# Empresas alvo - nomes construidos via Unicode para evitar problemas de codepage
$empBrasanitas = "Brasanitas"
$empSolucoes = [string]::new([char[]](83,111,108,117,0x00E7,0x00F5,101,115,32,83,101,114,118,105,0x00E7,111,115,32,84,101,114,99,101,105,114,105,122,97,100,111,115))
$empresas = @($empBrasanitas, $empSolucoes)

try {
    "Base: $base" | Out-File $log -Append
    "Existe base: $(Test-Path $base)" | Out-File $log -Append

    # Coleta arquivos RESUMO ANALITICO.xlsx de cada empresa (max 15 por empresa)
    $arquivos = @()
    foreach ($emp in $empresas) {
        $empPath = Join-Path $base $emp
        "Empresa: $emp | Existe: $(Test-Path $empPath)" | Out-File $log -Append
        if (Test-Path $empPath) {
            $dirs = Get-ChildItem -Path $empPath -Directory -ErrorAction SilentlyContinue
            "  Dirs encontrados: $($dirs.Count)" | Out-File $log -Append
            $countEmp = 0
            foreach ($d in $dirs) {
                if ($countEmp -ge 15) { break }
                $xlsx = Get-ChildItem -Path $d.FullName -Filter "*RESUMO*.xlsx" -File -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($xlsx) {
                    $arquivos += [PSCustomObject]@{
                        Tipo      = $tipoCalc
                        Empresa   = $emp
                        Reclamante = $d.Name
                        Arquivo   = $xlsx.FullName
                    }
                    $countEmp++
                }
            }
            "  Arquivos coletados: $countEmp" | Out-File $log -Append
        }
    }

    "Total arquivos encontrados: $($arquivos.Count)" | Out-File $log -Append

    # 15 amostras de cada empresa (ja coletadas acima)
    $amostras = $arquivos
    "Total amostras: $($amostras.Count)" | Out-File $log -Append

    # Abre Excel via COM
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $excel.AskToUpdateLinks = $false
    $excel.AlertBeforeOverwriting = $false
    "Excel COM aberto" | Out-File $log -Append

    # Workbook de saida
    $wbOut = $excel.Workbooks.Add()
    $wsOut = $wbOut.Sheets.Item(1)
    $wsOut.Name = "Consolidado"

    $todasRubricas = New-Object System.Collections.Generic.List[string]
    $linhas = New-Object System.Collections.Generic.List[object]

    $i = 0
    foreach ($amostra in $amostras) {
        $i++
        "[$i] Processando: $($amostra.Empresa) \ $($amostra.Reclamante)" | Out-File $log -Append
        try {
            $wb = $excel.Workbooks.Open($amostra.Arquivo, 0, $true)
            $ws = $wb.Sheets.Item(1)
            $used = $ws.UsedRange
            $rows = $used.Rows.Count
            $cols = $used.Columns.Count
            "  Linhas: $rows, Colunas: $cols" | Out-File $log -Append

            # Leitura em bloco (array 2D)
            $arr = $used.Value2
            $vals = @{}
            for ($r = 1; $r -le $rows; $r++) {
                for ($c = 1; $c -le $cols; $c++) {
                    $v = $arr[$r, $c]
                    if ($v) { $vals["$r,$c"] = "$v".Trim() }
                }
            }

            # Localiza VALORES APURADOS e TOTAL BRUTO APURADO
            $linhaInicio = 0
            $linhaFim = 0
            for ($r = 1; $r -le $rows; $r++) {
                $combinado = "$($vals["$r,1"]) $($vals["$r,2"]) $($vals["$r,3"]) $($vals["$r,4"])"
                if ($combinado -match "VALORES\s+APURADOS" -and $linhaInicio -eq 0) { $linhaInicio = $r }
                if ($combinado -match "TOTAL\s+BRUTO\s+APURADO") { $linhaFim = $r; break }
            }
            "  LinhaInicio: $linhaInicio, LinhaFim: $linhaFim" | Out-File $log -Append

            # Dados do processo
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

            # Extrai rubricas
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

            # Totais da linha TOTAL BRUTO APURADO
            $totalBruto = ""; $correcaoTotal = ""; $totalApurado = ""
            if ($linhaFim -gt 0) {
                $totalBruto    = $vals["$linhaFim,8"]
                $correcaoTotal = $vals["$linhaFim,9"]
                $totalApurado  = $vals["$linhaFim,11"]
            }

            $linhas.Add([PSCustomObject]@{
                Tipo            = $amostra.Tipo
                Empresa         = $amostra.Empresa
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
                Tipo            = $amostra.Tipo
                Empresa         = $amostra.Empresa
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

    "Montando planilha. Rubricas distintas: $($todasRubricas.Count)" | Out-File $log -Append

    # Funcao para escrever cabeçalho e dados em uma aba
    function Escrever-Aba($ws, $linhasAba, $rubricasLista) {
        $col = 1
        $ws.Cells.Item(1, $col).Value2 = "Tipo (Subpasta)"; $col++
        $ws.Cells.Item(1, $col).Value2 = "Empresa"; $col++
        $ws.Cells.Item(1, $col).Value2 = "Reclamante"; $col++
        $ws.Cells.Item(1, $col).Value2 = "Arquivo"; $col++
        $ws.Cells.Item(1, $col).Value2 = "Processo"; $col++
        $ws.Cells.Item(1, $col).Value2 = "Reclamante (Planilha)"; $col++
        $ws.Cells.Item(1, $col).Value2 = "Reclamada"; $col++
        $ws.Cells.Item(1, $col).Value2 = "Data Referencia"; $col++
        $ws.Cells.Item(1, $col).Value2 = "correcao"; $col++
        $ws.Cells.Item(1, $col).Value2 = "total apurado"; $col++
        $ws.Cells.Item(1, $col).Value2 = "TOTAL BRUTO APURADO"; $colTotalBruto = $col; $col++

        $colunasRub = @()
        foreach ($rub in $rubricasLista) {
            $ws.Cells.Item(1, $col).Value2 = "$rub - Principal"; $colunasRub += ,@($rub, "Principal"); $col++
            $ws.Cells.Item(1, $col).Value2 = "$rub - Correcao";  $colunasRub += ,@($rub, "Correcao");  $col++
            $ws.Cells.Item(1, $col).Value2 = "$rub - Juros";     $colunasRub += ,@($rub, "Juros");     $col++
            $ws.Cells.Item(1, $col).Value2 = "$rub - Total";     $colunasRub += ,@($rub, "Total");     $col++
        }

        $rowOut = 2
        foreach ($linha in $linhasAba) {
            $c = 1
            $ws.Cells.Item($rowOut, $c).Value2 = $linha.Tipo; $c++
            $ws.Cells.Item($rowOut, $c).Value2 = $linha.Empresa; $c++
            $ws.Cells.Item($rowOut, $c).Value2 = $linha.ReclamantePasta; $c++
            $ws.Cells.Item($rowOut, $c).Value2 = $linha.Arquivo; $c++
            $ws.Cells.Item($rowOut, $c).Value2 = $linha.Processo; $c++
            $ws.Cells.Item($rowOut, $c).Value2 = $linha.Reclamante; $c++
            $ws.Cells.Item($rowOut, $c).Value2 = $linha.Reclamada; $c++
            $ws.Cells.Item($rowOut, $c).Value2 = $linha.DataReferencia; $c++
            $ws.Cells.Item($rowOut, $c).Value2 = $linha.CorrecaoTotal; $c++
            $ws.Cells.Item($rowOut, $c).Value2 = $linha.TotalApurado; $c++
            $ws.Cells.Item($rowOut, $colTotalBruto).Value2 = $linha.TotalBrutoApurado; $c++
            foreach ($cr in $colunasRub) {
                $rub = $cr[0]; $campo = $cr[1]
                $val = ""
                if ($linha.Rubricas.ContainsKey($rub)) { $val = $linha.Rubricas[$rub].$campo }
                $ws.Cells.Item($rowOut, $c).Value2 = $val
                $c++
            }
            $rowOut++
        }

        $ws.UsedRange.EntireColumn.AutoFit() | Out-Null
        $ws.Rows.Item(1).Font.Bold = $true
        $ws.Rows.Item(1).Interior.Color = 13434828
    }

    # Aba 1: Consolidado (todas as 30 amostras)
    Escrever-Aba $wsOut $linhas $todasRubricas

    # Aba 2: Brasanitas
    $wsBras = $wbOut.Worksheets.Add([System.Reflection.Missing]::Value, $wsOut)
    $wsBras.Name = "Brasanitas"
    $rubBras = New-Object System.Collections.Generic.List[string]
    foreach ($l in ($linhas | Where-Object { $_.Empresa -eq "Brasanitas" })) {
        foreach ($k in $l.Rubricas.Keys) { if (-not $rubBras.Contains($k)) { $rubBras.Add($k) } }
    }
    Escrever-Aba $wsBras ($linhas | Where-Object { $_.Empresa -eq "Brasanitas" }) $rubBras

    # Aba 3: Solucoes
    $wsSol = $wbOut.Worksheets.Add([System.Reflection.Missing]::Value, $wsBras)
    $wsSol.Name = "Solucoes"
    $rubSol = New-Object System.Collections.Generic.List[string]
    $empSol = $empresas[1]
    foreach ($l in ($linhas | Where-Object { $_.Empresa -eq $empSol })) {
        foreach ($k in $l.Rubricas.Keys) { if (-not $rubSol.Contains($k)) { $rubSol.Add($k) } }
    }
    Escrever-Aba $wsSol ($linhas | Where-Object { $_.Empresa -eq $empSol }) $rubSol

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
