# Solucoes - TODOS os casos | apenas colunas A-K (sem rubricas detalhadas)
# CSV instantaneo + conversao unica para xlsx (rapido)

$ErrorActionPreference = 'Continue'
$log = "C:\Users\jose.madella\Desktop\extrair_log_solucoes_final.txt"
"INICIO: $(Get-Date)" | Out-File $log -Encoding UTF8

$calcBase = "X:\Trabalhista\Arquivos Trabalhistas\C" + [char]0x00E1 + "lculos Elaborados"
$tipoCalc = "Materializa" + [char]0x00E7 + [char]0x00E3 + "o de Decis" + [char]0x00E3 + "o"
$base = Join-Path $calcBase $tipoCalc
$empSolucoes = [string]::new([char[]](83,111,108,117,0x00E7,0x00F5,101,115,32,83,101,114,118,105,0x00E7,111,115,32,84,101,114,99,101,105,114,105,122,97,100,111,115))
$pdftotext = "C:\Users\jose.madella\AppData\Local\Microsoft\WinGet\Packages\oschwartz10612.Poppler_Microsoft.Winget.Source_8wekyb3d8bbwe\poppler-25.07.0\Library\bin\pdftotext.exe"

try {
    $empPath = Join-Path $base $empSolucoes
    "Empresa: $empSolucoes | Existe: $(Test-Path $empPath)" | Out-File $log -Append
    $dirs = Get-ChildItem -Path $empPath -Directory -ErrorAction SilentlyContinue
    "Dirs: $($dirs.Count)" | Out-File $log -Append

    # Coleta TODOS os PDFs
    $arquivos = @()
    foreach ($d in $dirs) {
        $file = Get-ChildItem -Path $d.FullName -Filter "*.pdf" -File -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($file) { $arquivos += [PSCustomObject]@{ Reclamante = $d.Name; Arquivo = $file.FullName } }
    }
    "Total PDFs: $($arquivos.Count)" | Out-File $log -Append

    # CSV (tab-separated, UTF-8)
    $csvPath = "C:\Users\jose.madella\Desktop\Solucoes_Todos_Casos.csv"
    $header = "Tipo (Subpasta)`tEmpresa`tReclamante`tArquivo`tProcesso`tReclamante (Planilha)`tReclamada`tData Referencia`tcorrecao`ttotal apurado`tTOTAL BRUTO APURADO"
    $header | Out-File $csvPath -Encoding UTF8

    $i = 0
    foreach ($amostra in $arquivos) {
        $i++
        if ($i % 50 -eq 0) { "Progresso: $i / $($arquivos.Count)" | Out-File $log -Append }
        try {
            $tmpTxt = [System.IO.Path]::GetTempFileName()
            & $pdftotext -layout $amostra.Arquivo $tmpTxt 2>$null
            $linhas = Get-Content $tmpTxt -Encoding UTF8
            Remove-Item $tmpTxt -Force -ErrorAction SilentlyContinue

            $processo = ""; $reclamante = ""; $reclamada = ""; $dataRef = ""
            $totalBruto = ""; $totalApurado = ""; $liquidoDevido = ""; $totalDevido = ""

            foreach ($linha in $linhas) {
                if ($linha -match "Processo:\s*(.+?)(\s|$)" -and -not $processo) { $processo = $Matches[1].Trim() }
                if ($linha -match "^Reclamante:\s*(.+)") { $reclamante = $Matches[1].Trim() }
                if ($linha -match "^Reclamado:\s*(.+)") { $reclamada = $Matches[1].Trim() }
                if ($linha -match "Data Liquida" -and $linha -match "(\d{2}/\d{2}/\d{4})") { $dataRef = $Matches[1] }
                # Total da secao Resumo: "Total  25.468,05  3.152,85  28.620,90"
                if ($linha -match "^\s*Total\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s*$" -and -not $totalApurado) {
                    $totalBruto = $Matches[1]; $totalApurado = $Matches[3]
                }
                if ($linha -match "Bruto Devido ao Reclamante\s+([\d.,]+)" -and -not $totalApurado) { $totalApurado = $Matches[1] }
                if ($linha -match "L.quido Devido ao Reclamante\s+([\d.,]+)" -and -not $liquidoDevido) { $liquidoDevido = $Matches[1] }
                if ($linha -match "Total Devido pelo Reclamado\s+([\d.,]+)" -and -not $totalDevido) { $totalDevido = $Matches[1] }
            }

            $row = @($tipoCalc, "Solucoes", $amostra.Reclamante, (Split-Path $amostra.Arquivo -Leaf), $processo, $reclamante, $reclamada, $dataRef, "", $totalApurado, $totalBruto)
            $csvLine = ($row | ForEach-Object { '"' + ($_ -replace '"','""') + '"' }) -join "`t"
            $csvLine | Out-File $csvPath -Encoding UTF8 -Append
        } catch {
            "  ERRO [$i]: $($_.Exception.Message)" | Out-File $log -Append
        }
    }
    "CSV gerado: $csvPath | Linhas: $i" | Out-File $log -Append

    # Converte CSV para XLSX (unica operacao COM - rapido)
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($csvPath)
    $ws = $wb.Sheets.Item(1)
    $ws.UsedRange.EntireColumn.AutoFit() | Out-Null
    $ws.Rows.Item(1).Font.Bold = $true
    $ws.Rows.Item(1).Interior.Color = 13434828
    $xlsxPath = "C:\Users\jose.madella\Desktop\Solucoes_Todos_Casos.xlsx"
    $wb.SaveAs($xlsxPath, 51)
    $wb.Close(); $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    "XLSX gerado: $xlsxPath" | Out-File $log -Append
    "CONCLUIDO: $(Get-Date)" | Out-File $log -Append
} catch {
    "ERRO FATAL: $($_.Exception.Message)" | Out-File $log -Append
    "STACK: $($_.ScriptStackTrace)" | Out-File $log -Append
}
