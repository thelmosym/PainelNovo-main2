Add-Type -AssemblyName System.IO.Compression.FileSystem

$paths = @("Monitoramento", "Medição - Contratada", "Versoes Anteriores", "PainelAgosto")
foreach ($dir in $paths) {
    if (Test-Path $dir) {
        $files = Get-ChildItem -Path $dir -Recurse -Include *.xlsm, *.xlsx
        foreach ($file in $files) {
            try {
                $zip = [System.IO.Compression.ZipFile]::OpenRead($file.FullName)
                foreach ($entry in $zip.Entries) {
                    if ($entry.FullName -like "xl/sharedStrings.xml" -or $entry.FullName -like "xl/worksheets/sheet*.xml") {
                        $stream = $entry.Open()
                        $reader = New-Object System.IO.StreamReader($stream)
                        $content = $reader.ReadToEnd()
                        $reader.Close()
                        $stream.Close()
                        if ($content -match "CSC0823719" -or $content -match "000023/2026") {
                            Write-Host "FOUND in $($file.FullName) entry $($entry.FullName)"
                        }
                    }
                }
                $zip.Dispose()
            } catch {
                # ignore read error
            }
        }
    }
}
