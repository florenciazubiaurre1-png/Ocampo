Install-Module -Name ImportExcel -Force -AllowClobber -Scope CurrentUser
# 1. Configuración
$rootPath = "C:\OCAMPO\Solucion\OD"
$filaInicio = 3 

$files = Get-ChildItem -Path $rootPath -Include *.xls, *.xlsx -Recurse


Write-Host "Solucionando duplicados en fila $filaInicio..." -ForegroundColor Cyan

foreach ($file in $files) {
    try {
        $sheetNames = Get-ExcelSheetInfo -Path $file.FullName | Select-Object -ExpandProperty Name
        Write-Host "`nArchivo: $($file.Name)" -ForegroundColor White
        
        foreach ($sheet in $sheetNames) {
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
            $csvName = "${baseName}_${sheet}.csv"
            $csvPath = Join-Path -Path $file.DirectoryName -ChildPath $csvName
            
            Write-Host "  -> Procesando [$sheet] (Forzando cabeceras únicas)... " -NoNewline
            
            # USAMOS -NoHeader para saltar el error de duplicados
            # ImportExcel leerá la fila 3 como la primera fila de datos.
            $data = Import-Excel -Path $file.FullName -WorksheetName $sheet -StartRow $filaInicio -NoHeader
            
            if ($null -ne $data) {
                # Exportamos. Las columnas se llamarán P1, P2, P3, etc.
                $data | Export-Csv -Path $csvPath -NoTypeInformation -Delimiter "," -Encoding UTF8
                Write-Host "[OK]" -ForegroundColor Green
            }
        }
    }
    catch {
        Write-Host " [ERROR]: $($_.Exception.Message)" -ForegroundColor Red
    }
}
