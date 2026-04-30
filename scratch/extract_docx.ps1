$wordPath = "C:\Users\Yolfzamb\Documents\PROGRAMACION WEB\Web app ValorM2Vzla\Términos Condiciones , Política de Privacidad y aviso legal.docx"
$outputPath = "C:\Users\Yolfzamb\Documents\PROGRAMACION WEB\T-rminos-y-Condiciones-Pol-tica-de-Privacidad---ValorM2Vzla-main\extracted_text.txt"

try {
    Add-Type -AssemblyName "System.IO.Compression.FileSystem"
    $tempDir = Join-Path $env:TEMP ([Guid]::NewGuid().ToString())
    New-Item -ItemType Directory -Path $tempDir | Out-Null
    
    [System.IO.Compression.ZipFile]::ExtractToDirectory($wordPath, $tempDir)
    
    $xmlPath = Join-Path $tempDir "word\document.xml"
    if (Test-Path $xmlPath) {
        [xml]$xml = Get-Content $xmlPath
        $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
        
        $paragraphs = $xml.SelectNodes("//w:p", $ns)
        $text = foreach ($p in $paragraphs) {
            $runs = $p.SelectNodes(".//w:t", $ns)
            $line = ($runs | ForEach-Object { $_."#text" }) -join ""
            $line
        }
        $text | Out-File $outputPath -Encoding utf8
        Write-Host "Success: Text extracted to $outputPath"
    } else {
        Write-Error "Could not find word/document.xml in the docx file."
    }
    
    Remove-Item -Path $tempDir -Recurse -Force
} catch {
    Write-Error "Error extracting text: $($_.Exception.Message)"
}
