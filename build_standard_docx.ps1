$ErrorActionPreference = "Stop"

function Escape-Xml {
    param([string]$Text)
    if ($null -eq $Text) { return "" }
    return [System.Security.SecurityElement]::Escape($Text)
}

function Para {
    param(
        [string]$Text,
        [int]$Size = 22,
        [switch]$Bold,
        [string]$Justify = "left",
        [int]$Before = 0,
        [int]$After = 60,
        [int]$Left = 0
    )

    $b = if ($Bold) { "<w:b/>" } else { "" }
    $ind = if ($Left -gt 0) { "<w:ind w:left=`"$Left`"/>" } else { "" }
    $t = Escape-Xml $Text
    return @"
<w:p>
  <w:pPr>
    <w:jc w:val="$Justify"/>
    <w:spacing w:before="$Before" w:after="$After" w:line="300" w:lineRule="auto"/>
    $ind
  </w:pPr>
  <w:r>
    <w:rPr>
      $b
      <w:rFonts w:ascii="Malgun Gothic" w:hAnsi="Malgun Gothic" w:eastAsia="Malgun Gothic"/>
      <w:sz w:val="$Size"/>
      <w:szCs w:val="$Size"/>
      <w:lang w:val="ko-KR" w:eastAsia="ko-KR"/>
    </w:rPr>
    <w:t xml:space="preserve">$t</w:t>
  </w:r>
</w:p>
"@
}

$root = $PSScriptRoot
$src = Join-Path $root "output\park_plan_clean_utf8.txt"
$out = Join-Path $root "output\park_plan_standard.docx"

if (-not (Test-Path $src)) { throw "Missing source file: $src" }
if (Test-Path $out) { Remove-Item -LiteralPath $out -Force }

$lines = [System.IO.File]::ReadAllLines($src, [System.Text.Encoding]::UTF8)
$paras = New-Object System.Collections.Generic.List[string]
foreach ($line in $lines) {
    if ([string]::IsNullOrWhiteSpace($line)) {
        $paras.Add((Para -Text "" -Size 20 -After 20))
    } elseif ($line -eq "[행사 계획서]") {
        $paras.Add((Para -Text $line -Size 32 -Bold -Justify center -After 100))
    } elseif ($line -match '^\d+\.\s') {
        $paras.Add((Para -Text $line -Size 26 -Bold -Before 100 -After 60))
    } elseif ($line -match '^\d{2}:\d{2}~\d{2}:\d{2}') {
        $paras.Add((Para -Text $line -Size 22 -Bold -Before 60 -After 30))
    } elseif ($line -like '- *') {
        $paras.Add((Para -Text $line -Size 22 -Left 360 -After 20))
    } else {
        $paras.Add((Para -Text $line -Size 22))
    }
}

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
    $($paras -join "`n")
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>
      <w:cols w:space="708"/>
      <w:docGrid w:linePitch="360"/>
    </w:sectPr>
  </w:body>
</w:document>
"@

$stylesXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Malgun Gothic" w:hAnsi="Malgun Gothic" w:eastAsia="Malgun Gothic"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
        <w:lang w:val="ko-KR" w:eastAsia="ko-KR"/>
      </w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
</w:styles>
"@

$settingsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="720"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
</w:settings>
"@

$webSettingsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:optimizeForBrowser/>
</w:webSettings>
"@

$fontTableXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:font w:name="Malgun Gothic"/>
</w:fonts>
"@

$themeXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
      <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
      <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
      <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
      <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
      <a:accent6><a:srgbClr val="F79646"/></a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri"/><a:ea typeface="Malgun Gothic"/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface="Malgun Gothic"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>
"@

$documentRelsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>
"@

$rootRelsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"@

$contentTypesXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>
  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
  <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>
"@

$timestamp = (Get-Date).ToString("s") + "Z"
$coreXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:dcmitype="http://purl.org/dc/dcmitype/"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Park Plan</dc:title>
  <dc:creator>OpenAI Codex</dc:creator>
  <cp:lastModifiedBy>OpenAI Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">$timestamp</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">$timestamp</dcterms:modified>
</cp:coreProperties>
"@

$appXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Office Word</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0000</AppVersion>
</Properties>
"@

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$utf8 = [System.Text.UTF8Encoding]::new($false)
$archive = [System.IO.Compression.ZipFile]::Open($out, [System.IO.Compression.ZipArchiveMode]::Create)
try {
    $entries = @(
        @{ Name='[Content_Types].xml'; Content=$contentTypesXml },
        @{ Name='_rels/.rels'; Content=$rootRelsXml },
        @{ Name='docProps/core.xml'; Content=$coreXml },
        @{ Name='docProps/app.xml'; Content=$appXml },
        @{ Name='word/document.xml'; Content=$documentXml },
        @{ Name='word/styles.xml'; Content=$stylesXml },
        @{ Name='word/settings.xml'; Content=$settingsXml },
        @{ Name='word/webSettings.xml'; Content=$webSettingsXml },
        @{ Name='word/fontTable.xml'; Content=$fontTableXml },
        @{ Name='word/theme/theme1.xml'; Content=$themeXml },
        @{ Name='word/_rels/document.xml.rels'; Content=$documentRelsXml }
    )
    foreach ($e in $entries) {
        $entry = $archive.CreateEntry($e.Name, [System.IO.Compression.CompressionLevel]::Optimal)
        $writer = New-Object System.IO.StreamWriter($entry.Open(), $utf8)
        try { $writer.Write($e.Content) } finally { $writer.Dispose() }
    }
}
finally {
    $archive.Dispose()
}

Write-Output $out
