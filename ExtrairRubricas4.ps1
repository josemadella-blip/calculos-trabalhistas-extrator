# Extrai rubricas lendo xlsx como ZIP/XML (sem Excel COM)
# Colunas: A=Tipo, B=Empresa, C=Reclamante
# Filtro: Brasanitas e Solucoes | 30 amostras (15+15) | 3 abas

$ErrorActionPreference = 'Continue'
$log = "C:\Users\jose.madella\Desktop\extrair_log_zip.txt"
"INICIO: $(Get-Date)" | Out-File $log -Encoding UTF8

# Caminhos com acentos via Unicode
$calcBase = "X:\Trabalhista\Arquivos Trabalhistas\C" + [char]0x00E1 + "lculos Elaborados"
$tipoCalc = "Materializa" + [char]0x00E7 + [char]0x00E3 + "o de Decis" + [char]0x00E3 + "o"
$base = Join-Path $calcBase $tipoCalc

$empBrasanitas = "Brasanitas"
$empSolucoes = [string]::new([char[]](83,111,108,117,0x00E7,0x00F5,101,115,32,83,101,114,118,105,0x00E7,111,115,32,84,101,114,99,101,105,114,105,122,97,100,111,115))
$empresas = @($empBrasanitas, $empSolucoes)

# Funcao: ler xlsx como ZIP e extrair sharedStrings + sheet1 XML
function Ler-XlsxComoXml($caminho) {
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [System.IO.Compression.ZipFile]::OpenRead($caminho)
        
        # SharedStrings (strings compartilhadas)
        $ssEntry = $zip.Entries | Where-Object { $_.FullName -eq "xl/sharedStrings.xml" }
        $sharedStrings = @()
        if ($ssEntry) {
            $reader = New-Object System.IO.StreamReader($ssEntry.Open())
            $ssXml = [xml]$reader.ReadToEnd()
            $reader.Close()
            $sharedStrings = @($ssXml.sst.si | ForEach-Object {
                $t = $_.t
                if (-not $t -and $_.r) { $t = ($_.r | ForEach-Object { $_.t } | Out-String) }
                "$t"
            })
        }
        
        # Sheet1
        $sheetEntry = $zip.Entries | Where-Object { $_.FullName -eq "xl/worksheets/sheet1.xml" }
        if (-not $sheetEntry) {
            $sheetEntry = $zip.Entries | Where-Object { $_.FullName -like "xl/worksheets/sheet*.xml" } | Select-Object -First 1
        }
        $sheetXml = $null
        if ($sheetEntry) {
            $reader = New-Object System.IO.StreamReader($sheetEntry.Open())
            $sheetXml = [xml]$reader.ReadToEnd()
            $reader.Close()
        }
        
        $zip.Dispose()
        return @{ SharedStrings = $sharedStrings; Sheet = $sheetXml }
    } catch {
        return $null
    }
}

# Funcao: converter referencia de coluna (A, B, AA) para numero
function Col-LetraParaNum($letras) {
    $num = 0
    foreach ($ch in $letras.ToCharArray()) {
        $num = $num * 26 + ([int]$ch - 64)
    }
    return $num
}

# Funcao: extrair celulas do sheet XML para hashtable "r,c" => valor
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
            $ref = $cell.r  # ex: "B9"
            if (-not $ref) { continue }
            # Separar coluna(letras) de linha(numeros)
            $match = [regex]::Match($ref, "^([A-Z]+)(\d+)$")
            if (-not $match.Success) { continue }
            $colNum = Col-LetraParaNum $match.Groups[1].Value
            $rowNum = [int]$match.Groups[2].Value
            
            $val = ""
            if ($cell.t -eq "s") {
                # String compartilhada
                $idx = [int]$cell.v
                if ($idx -lt $sharedStrings.Count) { $val = $sharedStrings[$idx] }
            } elseif ($cell.t -eq "inlineStr") {
                $val = $cell.is.t
            } else {
                $val = "$($cell.v)"
            }
            if ($val) { $vals["$rowNum,$colNum"] = $val.Trim() }
        }
    }
    return $vals
}

try {
    "Base: $base | Existe: $(Test-Path $base)" | Out-File $log -Append

    # Coleta 15 arquivos por empresa
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
                $xlsx = Get-ChildItem -Path $d.FullName -Filter "*RESUMO*.xlsx" -File -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($xlsx) {
                    $arquivos += [PSCustomObject]@{
                        Tipo = $tipoCalc; Empresa = $emp; Reclamante = $d.Name; Arquivo = $xlsx.FullName
                    }
                    $countEmp++
                }
            }
            "  Coletados: $countEmp" | Out-File $log -Append
        }
    }
    "Total: $($arquivos.Count) arquivos" | Out-File $log -Append

    $todasRubricas = New-Object System.Collections.Generic.List[string]
    $linhas = New-Object System.Collections.Generic.List[object]

    $i = 0
    foreach ($amostra in $arquivos) {
        $i++
        "[$i] $($amostra.Empresa) \ $($amostra.Reclamante)" | Out-File $log -Append
        try {
            $data = Ler-XlsxComoXml $amostra.Arquivo
            if (-not $data) { "  ERRO: nao leu zip" | Out-File $log -Append; continue }
            
            $vals = Extrair-Celulas $data.Sheet $data.SharedStrings
            
            # Localiza VALORES APURADOS e TOTAL BRUTO APURADO
            $linhaInicio = 0; $linhaFim = 0
            for ($r = 1; $r -le 80; $r++) {
                $comb = "$($vals["$r,1"]) $($vals["$r,2"]) $($vals["$r,3"]) $($vals["$r,4"])"
                if ($comb -match "VALORES\s+APURADOS" -and $linhaInicio -eq 0) { $linhaInicio = $r }
                if ($comb -match "TOTAL\s+BRUTO\s+APURADO") { $linhaFim = $r; break }
            }
            "  Inicio: $linhaInicio Fim: $linhaFim" | Out-File $log -Append

            # Dados do processo
            $processo = ""; $reclamante = ""; $reclamada = ""; $dataRef = ""
            for ($r = 1; $r -le 12; $r++) {
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

            # Rubricas
            $rubricasArquivo = @{}
            if ($linhaInicio -gt 0 -and $linhaFim -gt 0) {
                for ($r = $linhaInicio + 2; $r -lt $linhaFim; $r++) {
                    $nomeRub = $vals["$r,4"]
                    if ($nomeRub -and $nomeRub -notmatch "^\s*$") {
                        $nomeRub = $nomeRub.Trim()
                        $rubricasArquivo[$nomeRub] = [PSCustomObject]@{
                            Principal = $vals["$r,8"]; Correcao = $vals["$r,9"]; Juros = $vals["$r,10"]; Total = $vals["$r,11"]
                        }
                        if (-not $todasRubricas.Contains($nomeRub)) { $todasRubricas.Add($nomeRub) }
                    }
                }
            }
            "  Rubricas: $($rubricasArquivo.Count)" | Out-File $log -Append

            # Totais
            $totalBruto = ""; $correcaoTotal = ""; $totalApurado = ""
            if ($linhaFim -gt 0) {
                $totalBruto = $vals["$linhaFim,8"]
                $correcaoTotal = $vals["$linhaFim,9"]
                $totalApurado = $vals["$linhaFim,11"]
            }

            $linhas.Add([PSCustomObject]@{
                Tipo = $amostra.Tipo; Empresa = $amostra.Empresa; ReclamantePasta = $amostra.Reclamante
                Arquivo = (Split-Path $amostra.Arquivo -Leaf); Processo = $processo; Reclamante = $reclamante
                Reclamada = $reclamada; DataReferencia = $dataRef; CorrecaoTotal = $correcaoTotal
                TotalApurado = $totalApurado; TotalBrutoApurado = $totalBruto; Rubricas = $rubricasArquivo
            })
        } catch {
            "  ERRO: $($_.Exception.Message)" | Out-File $log -Append
        }
    }

    "Montando planilha. Rubricas: $($todasRubricas.Count)" | Out-File $log -Append

    # Funcao auxiliar: converter data serial do Excel para texto legivel
    function Converter-DataSerial($val) {
        if ($val -match "^\d{4,5}$") {
            $n = [int]$val
            if ($n -gt 30000 -and $n -lt 60000) {
                try { return ([datetime]::FromOADate($n)).ToString("dd/MM/yyyy") } catch {}
            }
        }
        return "$val"
    }

    # Funcao auxiliar: limpar valor (remover System.Xml.XmlElement etc)
    function Limpar-Valor($val) {
        if (-not $val) { return "" }
        $s = "$val"
        if ($s -match "System\.Xml") { return "" }
        return $s.Trim()
    }

    # Gera XLSX diretamente via Excel COM (celula por celula)
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    $wbOut = $excel.Workbooks.Add()
    $wsOut = $wbOut.Sheets.Item(1)
    $wsOut.Name = "Consolidado"

    # Cabecalho
    $col = 1
    $cabecalhos = @("Tipo (Subpasta)","Empresa","Reclamante","Arquivo","Processo","Reclamante (Planilha)","Reclamada","Data Referencia","correcao","total apurado","TOTAL BRUTO APURADO")
    foreach ($h in $cabecalhos) {
        $wsOut.Cells.Item(1, $col).Value2 = $h
        $col++
    }
    $colTotalBruto = $col - 1
    $colunasRub = @()
    foreach ($rub in $todasRubricas) {
        $rubLimpo = Limpar-Valor $rub
        $wsOut.Cells.Item(1, $col).Value2 = "$rubLimpo - Principal"; $colunasRub += ,@($rub, "Principal"); $col++
        $wsOut.Cells.Item(1, $col).Value2 = "$rubLimpo - Correcao";  $colunasRub += ,@($rub, "Correcao");  $col++
        $wsOut.Cells.Item(1, $col).Value2 = "$rubLimpo - Juros";     $colunasRub += ,@($rub, "Juros");     $col++
        $wsOut.Cells.Item(1, $col).Value2 = "$rubLimpo - Total";     $colunasRub += ,@($rub, "Total");     $col++
    }

    # Dados
    $rowOut = 2
    foreach ($linha in $linhas) {
        $c = 1
        $wsOut.Cells.Item($rowOut, $c).Value2 = Limpar-Valor $linha.Tipo; $c++
        $wsOut.Cells.Item($rowOut, $c).Value2 = Limpar-Valor $linha.Empresa; $c++
        $wsOut.Cells.Item($rowOut, $c).Value2 = Limpar-Valor $linha.ReclamantePasta; $c++
        $wsOut.Cells.Item($rowOut, $c).Value2 = Limpar-Valor $linha.Arquivo; $c++
        $wsOut.Cells.Item($rowOut, $c).Value2 = Limpar-Valor $linha.Processo; $c++
        $wsOut.Cells.Item($rowOut, $c).Value2 = Limpar-Valor $linha.Reclamante; $c++
        $wsOut.Cells.Item($rowOut, $c).Value2 = Limpar-Valor $linha.Reclamada; $c++
        $wsOut.Cells.Item($rowOut, $c).Value2 = Converter-DataSerial (Limpar-Valor $linha.DataReferencia); $c++
        $wsOut.Cells.Item($rowOut, $c).Value2 = Limpar-Valor $linha.CorrecaoTotal; $c++
        $wsOut.Cells.Item($rowOut, $c).Value2 = Limpar-Valor $linha.TotalApurado; $c++
        $wsOut.Cells.Item($rowOut, $colTotalBruto).Value2 = Limpar-Valor $linha.TotalBrutoApurado; $c++
        foreach ($cr in $colunasRub) {
            $rub = $cr[0]; $campo = $cr[1]
            $val = ""
            if ($linha.Rubricas.ContainsKey($rub)) { $val = $linha.Rubricas[$rub].$campo }
            $wsOut.Cells.Item($rowOut, $c).Value2 = Limpar-Valor $val
            $c++
        }
        $rowOut++
    }

    # Formatacao
    $wsOut.UsedRange.EntireColumn.AutoFit() | Out-Null
    $wsOut.Rows.Item(1).Font.Bold = $true
    $wsOut.Rows.Item(1).Interior.Color = 13434828

    $xlsxPath = "C:\Users\jose.madella\Desktop\Resumo_Rubricas_30amostras.xlsx"
    $wbOut.SaveAs($xlsxPath, 51)
    $wbOut.Close()
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    
    "XLSX gerado: $xlsxPath" | Out-File $log -Append
    "CONCLUIDO: $(Get-Date)" | Out-File $log -Append
} catch {
    "ERRO FATAL: $($_.Exception.Message)" | Out-File $log -Append
    "STACK: $($_.ScriptStackTrace)" | Out-File $log -Append
    if ($excel) { try { $excel.Quit() } catch {} }
}
