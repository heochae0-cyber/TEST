Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = 'Stop'

function Get-OleColor {
    param([int]$R, [int]$G, [int]$B)
    [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb($R, $G, $B))
}

function SX { param([double]$Value) [double]($Value * 0.75) }
function SY { param([double]$Value) [double]($Value * 0.75) }

function Set-TextBoxStyle {
    param(
        $Shape,
        [string]$Text,
        [string]$FontName = 'Malgun Gothic',
        [double]$FontSize = 18,
        [int]$FontColor = 0,
        [bool]$Bold = $false,
        [int]$Align = 1
    )

    $Shape.TextFrame.TextRange.Text = $Text
    $Shape.TextFrame.WordWrap = -1
    $Shape.TextFrame.AutoSize = 0
    $Shape.TextFrame.TextRange.Font.Name = $FontName
    $Shape.TextFrame.TextRange.Font.Size = $FontSize
    $Shape.TextFrame.TextRange.Font.Color.RGB = $FontColor
    $Shape.TextFrame.TextRange.Font.Bold = if ($Bold) { -1 } else { 0 }
    $Shape.TextFrame.TextRange.ParagraphFormat.Alignment = $Align
}

function Add-TextBox {
    param(
        $Slide,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height,
        [string]$Text,
        [string]$FontName = 'Malgun Gothic',
        [double]$FontSize = 18,
        [int]$FontColor = 0,
        [bool]$Bold = $false,
        [int]$Align = 1
    )

    $shape = $Slide.Shapes.AddTextbox(1, (SX $Left), (SY $Top), (SX $Width), (SY $Height))
    $shape.Fill.Visible = 0
    $shape.Line.Visible = 0
    Set-TextBoxStyle -Shape $shape -Text $Text -FontName $FontName -FontSize $FontSize -FontColor $FontColor -Bold $Bold -Align $Align
    return $shape
}

function Add-BoxShape {
    param(
        $Slide,
        [int]$Type,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height,
        [int]$FillColor,
        [bool]$NoLine = $true,
        [int]$LineColor = 0,
        [double]$LineWeight = 1
    )

    $shape = $Slide.Shapes.AddShape($Type, (SX $Left), (SY $Top), (SX $Width), (SY $Height))
    $shape.Fill.ForeColor.RGB = $FillColor
    if ($NoLine) {
        $shape.Line.Visible = 0
    } else {
        $shape.Line.ForeColor.RGB = $LineColor
        $shape.Line.Weight = $LineWeight
    }
    return $shape
}

function Add-Rect {
    param($Slide, [double]$Left, [double]$Top, [double]$Width, [double]$Height, [int]$FillColor)
    Add-BoxShape -Slide $Slide -Type 1 -Left $Left -Top $Top -Width $Width -Height $Height -FillColor $FillColor -NoLine $true
}

function Add-RoundRect {
    param($Slide, [double]$Left, [double]$Top, [double]$Width, [double]$Height, [int]$FillColor, [int]$LineColor = 0, [double]$LineWeight = 1, [bool]$NoLine = $false)
    Add-BoxShape -Slide $Slide -Type 5 -Left $Left -Top $Top -Width $Width -Height $Height -FillColor $FillColor -NoLine $NoLine -LineColor $LineColor -LineWeight $LineWeight
}

function Add-Line {
    param($Slide, [double]$X1, [double]$Y1, [double]$X2, [double]$Y2, [int]$Color, [double]$Weight = 1.5)
    $shape = $Slide.Shapes.AddLine((SX $X1), (SY $Y1), (SX $X2), (SY $Y2))
    $shape.Line.ForeColor.RGB = $Color
    $shape.Line.Weight = $Weight
}

function Add-TopBar {
    param($Slide, [int]$BrandColor)
    $null = Add-Rect -Slide $Slide -Left 0 -Top 0 -Width 1280 -Height 8 -FillColor $BrandColor
}

function Add-IntentTag {
    param($Slide, [string]$Text, [int]$BrandColor, [int]$SoftColor)
    $tag = Add-RoundRect -Slide $Slide -Left 1010 -Top 30 -Width 190 -Height 42 -FillColor $SoftColor -LineColor $BrandColor -LineWeight 1.25
    Set-TextBoxStyle -Shape $tag -Text $Text -FontSize 14 -FontColor $BrandColor -Bold $true -Align 2
    $tag.TextFrame.MarginTop = SY 7
    $tag.TextFrame.MarginLeft = SX 6
}

function Add-Header {
    param($Slide, [string]$Title, [string]$Subtitle, [int]$BrandColor, [int]$TextMuted, [int]$LightGrey)
    $null = Add-TextBox -Slide $Slide -Left 80 -Top 74 -Width 900 -Height 42 -Text $Title -FontSize 38 -FontColor $BrandColor -Bold $true
    $null = Add-TextBox -Slide $Slide -Left 80 -Top 118 -Width 980 -Height 28 -Text $Subtitle -FontSize 20 -FontColor $TextMuted
    Add-Line -Slide $Slide -X1 80 -Y1 168 -X2 1200 -Y2 168 -Color $LightGrey -Weight 1.5
}

function Add-Card {
    param($Slide, [double]$Left, [double]$Top, [double]$Width, [double]$Height, [string]$Title, [string]$Body, [int]$LightGrey, [int]$BrandColor, [int]$BodyColor)
    $null = Add-RoundRect -Slide $Slide -Left $Left -Top $Top -Width $Width -Height $Height -FillColor $LightGrey -NoLine $true
    $null = Add-Rect -Slide $Slide -Left $Left -Top $Top -Width 5 -Height $Height -FillColor $BrandColor
    $null = Add-TextBox -Slide $Slide -Left ($Left + 24) -Top ($Top + 20) -Width ($Width - 48) -Height 32 -Text $Title -FontSize 24 -FontColor (Get-OleColor 45 52 54) -Bold $true
    $null = Add-TextBox -Slide $Slide -Left ($Left + 24) -Top ($Top + 60) -Width ($Width - 48) -Height ($Height - 74) -Text $Body -FontSize 18 -FontColor $BodyColor
}

function Add-FocusBox {
    param($Slide, [double]$Left, [double]$Top, [double]$Width, [double]$Height, [string]$Title, [string]$Body, [int]$SoftColor, [int]$BrandColor, [int]$BodyColor, [int]$Align = 2)
    $shape = Add-RoundRect -Slide $Slide -Left $Left -Top $Top -Width $Width -Height $Height -FillColor $SoftColor -LineColor $BrandColor -LineWeight 1.5
    $shape.Line.DashStyle = 4
    $null = Add-TextBox -Slide $Slide -Left ($Left + 20) -Top ($Top + 20) -Width ($Width - 40) -Height 30 -Text $Title -FontSize 24 -FontColor $BrandColor -Bold $true -Align $Align
    $null = Add-TextBox -Slide $Slide -Left ($Left + 20) -Top ($Top + 62) -Width ($Width - 40) -Height ($Height - 72) -Text $Body -FontSize 20 -FontColor $BodyColor -Align $Align
}

function Add-AgendaItem {
    param($Slide, [double]$Top, [string]$TimeText, [string]$Body, [int]$BrandColor, [bool]$Highlight = $false)
    $border = if ($Highlight) { $BrandColor } else { Get-OleColor 221 221 221 }
    $weight = if ($Highlight) { 2 } else { 1 }
    $null = Add-RoundRect -Slide $Slide -Left 80 -Top $Top -Width 1120 -Height 62 -FillColor (Get-OleColor 255 255 255) -LineColor $border -LineWeight $weight
    $timeShape = Add-RoundRect -Slide $Slide -Left 105 -Top ($Top + 11) -Width 92 -Height 38 -FillColor $BrandColor -NoLine $true
    Set-TextBoxStyle -Shape $timeShape -Text $TimeText -FontSize 16 -FontColor (Get-OleColor 255 255 255) -Bold $true -Align 2
    $timeShape.TextFrame.MarginTop = SY 6
    $null = Add-TextBox -Slide $Slide -Left 220 -Top ($Top + 14) -Width 930 -Height 30 -Text $Body -FontSize 22 -FontColor (Get-OleColor 45 52 54) -Bold $true
}

function Convert-NewLines {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    return $Text -replace "`n", "`r`n"
}

$brandColor = Get-OleColor 239 80 48
$darkGrey = Get-OleColor 45 52 54
$lightGrey = Get-OleColor 241 242 246
$white = Get-OleColor 255 255 255
$accentSoft = Get-OleColor 255 241 240
$textMuted = Get-OleColor 108 117 125
$bodyText = Get-OleColor 75 75 75
$footerText = Get-OleColor 153 153 153

$dataPath = Join-Path $PSScriptRoot 'org_culture_briefing_data.json'
$data = Get-Content -LiteralPath $dataPath -Raw -Encoding UTF8 | ConvertFrom-Json

$outputDir = Join-Path $PSScriptRoot 'output'
if (-not (Test-Path $outputDir)) {
    $null = New-Item -ItemType Directory -Path $outputDir
}

$outputPath = Join-Path $outputDir '1Q_org_culture_briefing_draft.pptx'

$ppt = $null
$presentation = $null

try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = -1
    $presentation = $ppt.Presentations.Add()
    $presentation.PageSetup.SlideWidth = 960
    $presentation.PageSetup.SlideHeight = 540
    $blank = 12

    $slide = $presentation.Slides.Add(1, $blank)
    Add-TopBar -Slide $slide -BrandColor $brandColor
    $null = Add-TextBox -Slide $slide -Left 0 -Top 210 -Width 1280 -Height 24 -Text $data.cover.label -FontSize 20 -FontColor $brandColor -Bold $true -Align 2
    $null = Add-TextBox -Slide $slide -Left 0 -Top 255 -Width 1280 -Height 80 -Text $data.cover.title -FontName 'Montserrat' -FontSize 50 -FontColor $brandColor -Bold $true -Align 2
    $null = Add-TextBox -Slide $slide -Left 0 -Top 336 -Width 1280 -Height 34 -Text $data.cover.subtitle -FontSize 28 -FontColor $textMuted -Align 2
    Add-Line -Slide $slide -X1 470 -Y1 455 -X2 810 -Y2 455 -Color (Get-OleColor 221 221 221) -Weight 1
    $null = Add-TextBox -Slide $slide -Left 380 -Top 470 -Width 520 -Height 28 -Text $data.cover.owner -FontSize 20 -FontColor $darkGrey -Align 2

    $s = $data.slides[0]
    $slide = $presentation.Slides.Add(2, $blank)
    Add-TopBar -Slide $slide -BrandColor $brandColor
    Add-IntentTag -Slide $slide -Text $s.tag -BrandColor $brandColor -SoftColor $accentSoft
    Add-Header -Slide $slide -Title $s.title -Subtitle $s.subtitle -BrandColor $brandColor -TextMuted $textMuted -LightGrey $lightGrey
    Add-Card -Slide $slide -Left 80 -Top 220 -Width 530 -Height 150 -Title $s.cards[0].title -Body $s.cards[0].body -LightGrey $lightGrey -BrandColor $brandColor -BodyColor $bodyText
    Add-Card -Slide $slide -Left 650 -Top 220 -Width 530 -Height 150 -Title $s.cards[1].title -Body $s.cards[1].body -LightGrey $lightGrey -BrandColor $brandColor -BodyColor $bodyText
    Add-Card -Slide $slide -Left 80 -Top 395 -Width 530 -Height 150 -Title $s.cards[2].title -Body $s.cards[2].body -LightGrey $lightGrey -BrandColor $brandColor -BodyColor $bodyText
    Add-Card -Slide $slide -Left 650 -Top 395 -Width 530 -Height 150 -Title $s.cards[3].title -Body $s.cards[3].body -LightGrey $lightGrey -BrandColor $brandColor -BodyColor $bodyText

    $s = $data.slides[1]
    $slide = $presentation.Slides.Add(3, $blank)
    Add-TopBar -Slide $slide -BrandColor $brandColor
    Add-IntentTag -Slide $slide -Text $s.tag -BrandColor $brandColor -SoftColor $accentSoft
    Add-Header -Slide $slide -Title $s.title -Subtitle $s.subtitle -BrandColor $brandColor -TextMuted $textMuted -LightGrey $lightGrey
    Add-AgendaItem -Slide $slide -Top 225 -TimeText $s.agenda[0].time -Body $s.agenda[0].body -BrandColor $brandColor
    Add-AgendaItem -Slide $slide -Top 305 -TimeText $s.agenda[1].time -Body $s.agenda[1].body -BrandColor $brandColor -Highlight $s.agenda[1].highlight
    Add-AgendaItem -Slide $slide -Top 385 -TimeText $s.agenda[2].time -Body $s.agenda[2].body -BrandColor $brandColor
    $null = Add-TextBox -Slide $slide -Left 80 -Top 475 -Width 620 -Height 24 -Text $s.footnote -FontSize 18 -FontColor $brandColor -Bold $true

    $s = $data.slides[2]
    $slide = $presentation.Slides.Add(4, $blank)
    Add-TopBar -Slide $slide -BrandColor $brandColor
    Add-IntentTag -Slide $slide -Text $s.tag -BrandColor $brandColor -SoftColor $accentSoft
    Add-Header -Slide $slide -Title $s.title -Subtitle $s.subtitle -BrandColor $brandColor -TextMuted $textMuted -LightGrey $lightGrey
    $null = Add-TextBox -Slide $slide -Left 100 -Top 240 -Width 200 -Height 30 -Text $s.leftTitle -FontSize 28 -FontColor $darkGrey -Bold $true
    $null = Add-TextBox -Slide $slide -Left 100 -Top 285 -Width 470 -Height 170 -Text (Convert-NewLines $s.leftBody) -FontSize 20 -FontColor $bodyText
    Add-FocusBox -Slide $slide -Left 700 -Top 245 -Width 420 -Height 200 -Title $s.focusTitle -Body $s.focusBody -SoftColor $accentSoft -BrandColor $brandColor -BodyColor $darkGrey -Align 2

    $s = $data.slides[3]
    $slide = $presentation.Slides.Add(5, $blank)
    Add-TopBar -Slide $slide -BrandColor $brandColor
    Add-IntentTag -Slide $slide -Text $s.tag -BrandColor $brandColor -SoftColor $accentSoft
    Add-Header -Slide $slide -Title $s.title -Subtitle $s.subtitle -BrandColor $brandColor -TextMuted $textMuted -LightGrey $lightGrey
    Add-Card -Slide $slide -Left 80 -Top 255 -Width 350 -Height 220 -Title $s.cards[0].title -Body $s.cards[0].body -LightGrey $lightGrey -BrandColor $brandColor -BodyColor $bodyText
    Add-Card -Slide $slide -Left 465 -Top 255 -Width 350 -Height 220 -Title $s.cards[1].title -Body $s.cards[1].body -LightGrey $lightGrey -BrandColor $brandColor -BodyColor $bodyText
    Add-Card -Slide $slide -Left 850 -Top 255 -Width 350 -Height 220 -Title $s.cards[2].title -Body $s.cards[2].body -LightGrey $lightGrey -BrandColor $brandColor -BodyColor $bodyText

    $s = $data.slides[4]
    $slide = $presentation.Slides.Add(6, $blank)
    Add-TopBar -Slide $slide -BrandColor $brandColor
    Add-IntentTag -Slide $slide -Text $s.tag -BrandColor $brandColor -SoftColor $accentSoft
    Add-Header -Slide $slide -Title $s.title -Subtitle $s.subtitle -BrandColor $brandColor -TextMuted $textMuted -LightGrey $lightGrey
    Add-FocusBox -Slide $slide -Left 80 -Top 220 -Width 1120 -Height 130 -Title $s.focusTitle -Body $s.focusBody -SoftColor $accentSoft -BrandColor $brandColor -BodyColor $darkGrey -Align 1
    Add-Card -Slide $slide -Left 80 -Top 385 -Width 530 -Height 155 -Title $s.cards[0].title -Body $s.cards[0].body -LightGrey $lightGrey -BrandColor $brandColor -BodyColor $bodyText
    Add-Card -Slide $slide -Left 650 -Top 385 -Width 530 -Height 155 -Title $s.cards[1].title -Body $s.cards[1].body -LightGrey $lightGrey -BrandColor $brandColor -BodyColor $bodyText

    $s = $data.slides[5]
    $slide = $presentation.Slides.Add(7, $blank)
    Add-TopBar -Slide $slide -BrandColor $brandColor
    Add-IntentTag -Slide $slide -Text $s.tag -BrandColor $brandColor -SoftColor $accentSoft
    Add-Header -Slide $slide -Title $s.title -Subtitle $s.subtitle -BrandColor $brandColor -TextMuted $textMuted -LightGrey $lightGrey
    Add-Card -Slide $slide -Left 80 -Top 240 -Width 530 -Height 190 -Title $s.cards[0].title -Body (Convert-NewLines $s.cards[0].body) -LightGrey $lightGrey -BrandColor $brandColor -BodyColor $bodyText
    Add-Card -Slide $slide -Left 650 -Top 240 -Width 530 -Height 190 -Title $s.cards[1].title -Body (Convert-NewLines $s.cards[1].body) -LightGrey $lightGrey -BrandColor $brandColor -BodyColor $bodyText
    $null = Add-TextBox -Slide $slide -Left 170 -Top 475 -Width 940 -Height 30 -Text $s.closing -FontSize 24 -FontColor $brandColor -Bold $true -Align 2

    $s = $data.slides[6]
    $slide = $presentation.Slides.Add(8, $blank)
    Add-TopBar -Slide $slide -BrandColor $brandColor
    Add-IntentTag -Slide $slide -Text $s.tag -BrandColor $brandColor -SoftColor $accentSoft
    Add-Header -Slide $slide -Title $s.title -Subtitle $s.subtitle -BrandColor $brandColor -TextMuted $textMuted -LightGrey $lightGrey
    $guide = Add-RoundRect -Slide $slide -Left 80 -Top 220 -Width 1120 -Height 255 -FillColor $white -NoLine $false -LineColor (Get-OleColor 221 221 221) -LineWeight 1
    $guide.TextFrame.MarginLeft = SX 24
    $guide.TextFrame.MarginRight = SX 20
    $guide.TextFrame.MarginTop = SY 18
    $guide.TextFrame.MarginBottom = SY 18
    $guide.TextFrame.TextRange.Text = ([string]::Join("`r`n`r`n", ($s.guideItems | ForEach-Object { "• $_" })))
    $guide.TextFrame.TextRange.Font.Name = 'Malgun Gothic'
    $guide.TextFrame.TextRange.Font.Size = 20
    $guide.TextFrame.TextRange.Font.Color.RGB = $darkGrey
    $null = Add-TextBox -Slide $slide -Left 80 -Top 650 -Width 520 -Height 20 -Text $s.footer -FontSize 14 -FontColor $footerText

    if (Test-Path $outputPath) {
        Remove-Item -LiteralPath $outputPath -Force
    }

    $presentation.SaveAs($outputPath)
}
finally {
    if ($presentation) {
        $presentation.Close()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation)
    }
    if ($ppt) {
        $ppt.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Output "Created: $outputPath"
