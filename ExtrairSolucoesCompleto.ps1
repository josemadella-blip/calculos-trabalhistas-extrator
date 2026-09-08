# Solucoes - TODOS os casos em TODAS as subpastas de Calculos Elaborados
# Colunas A-K apenas | CSV + conversao unica para xlsx

$ErrorActionPreference = 'Continue'
$log = "C:\Users\jose.madella\Desktop\extrair_log_solucoes_completo.txt"
"INICIO: $(Get-Date)" | Out-File $log -Encoding UTF8

$calcBase = "X:\Trabalhista\Arquivos Trabalhistas\C" + [char]0x00E1 + "lculos Elaborados"
$pdftotext = "C:\Users\jose.madella\AppData\Local\Microsoft\WinGet\Packages\oschwartz10612.Poppler_Microsoft.Winget.Source_8wekyb3d8bbwe\poppler-25.07.0\Library\bin\pdftotext.exe"

try {
    $subpastas = Get-ChildItem -Path $calcBase -Directory -ErrorAction SilentlyContinue
    "Subpastas de Calculos Elaborados: $($subpastas.Count)" | Out-File $log -Append

    # Coleta TODOS os PDFs de todas as pastas que contenham "Solu" no nome
    $arquivos = @()
    foreach ($sp in $subpastas) {
        $solDirs = Get-ChildItem -Path $sp.FullName -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match "Solu" }
        foreach ($sol in $solDirs) {
            $reclamantes = Get-ChildItem -Path $sol.FullName -Directory -ErrorAction SilentlyContinue
            foreach ($d in $reclamantes) {
                $file = Get-ChildItem -Path $d.FullName -Filter "*.pdf" -File -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($file) {
                    $arquivos += [PSCustomObject]@{
                        Tipo = $sp.Name; Empresa = $sol.Name; Reclamante = $d.Name; Arquivo = $file.FullName
                    }
                }
            }
            "  $($sp.Name) \ $($sol.Name): $($reclamantes.Count) dirs" | Out-File $log -Append
        }
    }
    "Total PDFs coletados: $($arquivos.Count)" | Out-File $log -Append

    # CSV (semicolon-separated para Excel PT-BR)
    $csvPath = "C:\Users\jose.madella\Desktop\Solucoes_Completo.csv"
    $header = "Tipo (Subpasta);Empresa;Reclamante;Arquivo;Processo;Reclamante (Planilha);Reclamada;Data Referencia;correcao;total apurado;TOTAL BRUTO APURADO"
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

            $row = @($amostra.Tipo, $amostra.Empresa, $amostra.Reclamante, (Split-Path $amostra.Arquivo -Leaf), $processo, $reclamante, $reclamada, $dataRef, "", $totalApurado, $totalBruto)
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
    $xlsxPath = "C:\Users\jose.madella\Desktop\Solucoes_Completo.xlsx"
    $wb.SaveAs($xlsxPath, 51)
    $wb.Close(); $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    "XLSX gerado: $xlsxPath" | Out-File $log -Append
    "CONCLUIDO: $(Get-Date)" | Out-File $log -Append
} catch {
    "ERRO FATAL: $($_.Exception.Message)" | Out-File $log -Append
    "STACK: $($_.ScriptStackTrace)" | Out-File $log -Append
}