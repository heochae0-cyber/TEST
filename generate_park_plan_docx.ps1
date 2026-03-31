$ErrorActionPreference = "Stop"

function Escape-Xml {
    param([string]$Text)
    if ($null -eq $Text) { return "" }
    return [System.Security.SecurityElement]::Escape($Text)
}

function New-ParagraphXml {
    param(
        [string]$Text,
        [int]$FontSize = 21,
        [switch]$Bold,
        [switch]$Center,
        [int]$SpacingBefore = 0,
        [int]$SpacingAfter = 120
    )

    $escaped = Escape-Xml $Text
    $boldXml = ""
    if ($Bold) { $boldXml = "<w:b/>" }

    $jcXml = ""
    if ($Center) { $jcXml = "<w:jc w:val=`"center`"/>" }

    return @"
<w:p>
  <w:pPr>
    $jcXml
    <w:spacing w:before="$SpacingBefore" w:after="$SpacingAfter"/>
  </w:pPr>
  <w:r>
    <w:rPr>
      $boldXml
      <w:sz w:val="$FontSize"/>
      <w:szCs w:val="$FontSize"/>
    </w:rPr>
    <w:t xml:space="preserve">$escaped</w:t>
  </w:r>
</w:p>
"@
}

function Convert-LineToParagraph {
    param([string]$Line)

    if ([string]::IsNullOrWhiteSpace($Line)) {
        return New-ParagraphXml -Text "" -FontSize 18 -SpacingAfter 40
    }

    if ($Line.StartsWith("[")) {
        return New-ParagraphXml -Text $Line -FontSize 28 -Bold -Center -SpacingAfter 220
    }

    if ($Line -match '^\d+\.') {
        return New-ParagraphXml -Text $Line -FontSize 24 -Bold -SpacingBefore 100 -SpacingAfter 140
    }

    if ($Line.StartsWith("NOTE:")) {
        return New-ParagraphXml -Text $Line.Substring(5).Trim() -FontSize 19 -SpacingBefore 160 -SpacingAfter 20
    }

    if ($Line.Contains(":") -or $Line.Contains("`t")) {
        return New-ParagraphXml -Text $Line -FontSize 21 -SpacingAfter 80
    }

    if ($Line.StartsWith("- ")) {
        return New-ParagraphXml -Text $Line -FontSize 21 -SpacingAfter 80
    }

    return New-ParagraphXml -Text $Line -FontSize 21 -SpacingAfter 80
}

$root = $PSScriptRoot
$outputDir = Join-Path $root "output"
$docxPath = Join-Path $outputDir "park_plan.docx"
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

$sourceFile = Get-ChildItem -Path $root -Filter "*.txt" | Sort-Object Name | Select-Object -First 1
if ($null -eq $sourceFile) {
    throw "No source txt file found."
}

$allLines = [System.Text.Encoding]::UTF8.GetString([System.IO.File]::ReadAllBytes($sourceFile.FullName)).Split([Environment]::NewLine)
$selectedLines = $allLines[474..578] | ForEach-Object { $_.TrimEnd() }

$normalizedLines = New-Object System.Collections.Generic.List[string]
$previousBlank = $false
foreach ($line in $selectedLines) {
    $isBlank = [string]::IsNullOrWhiteSpace($line)
    if ($isBlank -and $previousBlank) {
        continue
    }
    $normalizedLines.Add($line)
    $previousBlank = $isBlank
}

$bodyParts = foreach ($line in $normalizedLines) { Convert-LineToParagraph -Line $line }

$documentXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
 xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:w10="urn:schemas-microsoft-com:office:word"
 xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
 xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
 xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
 xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
 xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
 mc:Ignorable="w14 wp14">
  <w:body>
    $($bodyParts -join "`n")
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>
"@

$contentTypesXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"@

$relsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"@

if (Test-Path $docxPath) {
    Remove-Item -LiteralPath $docxPath -Force
}

Add-Type -AssemblyName WindowsBase
$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
$package = [System.IO.Packaging.Package]::Open($docxPath, [System.IO.FileMode]::Create)
try {
    $documentUri = New-Object System.Uri('/word/document.xml', [System.UriKind]::Relative)
    $documentPart = $package.CreatePart($documentUri, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml', [System.IO.Packaging.CompressionOption]::Maximum)
    $writer = New-Object System.IO.StreamWriter($documentPart.GetStream(), $utf8NoBom)
    try {
        $writer.Write($documentXml)
    }
    finally {
        $writer.Dispose()
    }

    $package.CreateRelationship($documentUri, [System.IO.Packaging.TargetMode]::Internal, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument') | Out-Null
}
finally {
    $package.Close()
}

Write-Output $docxPath
