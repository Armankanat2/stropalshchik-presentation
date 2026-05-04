param(
    [string]$InputPath = "deliverables\издательство\КУРС для СТРОПАЛЬЩИКА_ФИНАЛ_95_слайдов.pptx",
    [string]$OutputPath = "deliverables\черновики\КУРС для СТРОПАЛЬЩИКА_черновик_актуальный.pptx"
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing

function Resolve-ProjectPath {
    param([string]$RelativePath)

    $projectRoot = Split-Path -Parent $PSScriptRoot
    [System.IO.Path]::GetFullPath((Join-Path $projectRoot $RelativePath))
}

function Ensure-Directory {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path | Out-Null
    }
}

function Get-NextVersionedDraftPath {
    param([string]$CurrentOutputPath)

    $directory = Split-Path -Parent $CurrentOutputPath
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($CurrentOutputPath)
    $extension = [System.IO.Path]::GetExtension($CurrentOutputPath)
    $pattern = '^' + [regex]::Escape($baseName) + '_(\d+)-(\d+)' + [regex]::Escape($extension) + '$'
    $existing = Get-ChildItem -LiteralPath $directory -File -ErrorAction SilentlyContinue

    $maxMinor = -1
    foreach ($file in $existing) {
        if ($file.Name -match $pattern) {
            $minor = [int]$matches[2]
            if ($minor -gt $maxMinor) {
                $maxMinor = $minor
            }
        }
    }

    $nextMinor = $maxMinor + 1
    Join-Path $directory ("{0}_0-{1}{2}" -f $baseName, $nextMinor, $extension)
}

function Test-IsVersionedDraftPath {
    param([string]$Path)

    $fileName = [System.IO.Path]::GetFileName($Path)
    $fileName -match '_\d+-\d+\.[^.]+$'
}

function Get-OleColor {
    param(
        [int]$R,
        [int]$G,
        [int]$B
    )

    [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb($R, $G, $B))
}

function Remove-GeneratedShapes {
    param($Slide)

    $names = @()
    foreach ($shape in @($Slide.Shapes)) {
        if ($shape.Name -like "Gen_*") {
            $names += $shape.Name
        }
    }

    foreach ($name in $names) {
        try {
            $Slide.Shapes.Item($name).Delete()
        } catch {
        }
    }
}

function Remove-ShapeIfExists {
    param(
        $Slide,
        [string]$Name
    )

    try {
        $Slide.Shapes.Item($Name).Delete()
    } catch {
    }
}

function Set-TextBoxText {
    param(
        $Shape,
        [string]$Text,
        [double]$Size,
        [int]$Color,
        [bool]$Bold = $false,
        [int]$Alignment = 1
    )

    $Shape.TextFrame.TextRange.Text = $Text
    $Shape.TextFrame.TextRange.Font.Name = "Verdana"
    $Shape.TextFrame.TextRange.Font.Size = $Size
    $Shape.TextFrame.TextRange.Font.Bold = [int]$Bold
    $Shape.TextFrame.TextRange.Font.Color.RGB = $Color
    $Shape.TextFrame.TextRange.ParagraphFormat.Alignment = $Alignment
    $Shape.TextFrame.MarginLeft = 6
    $Shape.TextFrame.MarginRight = 6
    $Shape.TextFrame.MarginTop = 4
    $Shape.TextFrame.MarginBottom = 4
    try {
        $Shape.TextFrame.TextRange.ParagraphFormat.Bullet.Type = 0
    } catch {
    }
}

function Add-GeneratedTextBox {
    param(
        $Slide,
        [string]$Name,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height,
        [string]$Text,
        [double]$Size,
        [int]$Color,
        [bool]$Bold = $false,
        [int]$Alignment = 1
    )

    $shape = $Slide.Shapes.AddTextbox(1, $Left, $Top, $Width, $Height)
    $shape.Name = $Name
    $shape.Fill.Visible = 0
    $shape.Line.Visible = 0
    Set-TextBoxText -Shape $shape -Text $Text -Size $Size -Color $Color -Bold $Bold -Alignment $Alignment
    $shape
}

function Add-GeneratedCard {
    param(
        $Slide,
        [string]$Name,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height,
        [int]$FillColor,
        [string]$Title,
        [string]$Body,
        [int]$TitleColor,
        [int]$BodyColor,
        [double]$TitleSize = 18,
        [double]$BodySize = 11.5,
        [double]$TitleHeight = 32,
        [double]$BodyTopOffset = 48,
        [double]$BodyHeight = -1,
        [double]$BodyLeftOffset = 12
    )

    if ($BodyHeight -lt 0) {
        $BodyHeight = $Height - ($BodyTopOffset + 12)
    }

    $shape = $Slide.Shapes.AddShape(5, $Left, $Top, $Width, $Height)
    $shape.Name = $Name
    $shape.Fill.ForeColor.RGB = $FillColor
    $shape.Line.Visible = 0
    $shape.Adjustments.Item(1) = 0.15

    $titleBox = Add-GeneratedTextBox -Slide $Slide -Name "${Name}_Title" -Left ($Left + 12) -Top ($Top + 12) -Width ($Width - 24) -Height $TitleHeight -Text $Title -Size $TitleSize -Color $TitleColor -Bold $true
    $bodyBox = Add-GeneratedTextBox -Slide $Slide -Name "${Name}_Body" -Left ($Left + $BodyLeftOffset) -Top ($Top + $BodyTopOffset) -Width ($Width - 24) -Height $BodyHeight -Text $Body -Size $BodySize -Color $BodyColor

    @($shape, $titleBox, $bodyBox)
}

function Add-PhotoInfoCard {
    param(
        $Slide,
        [string]$Name,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height,
        [int]$FillColor,
        [string]$Title,
        [string]$Body,
        [int]$TitleColor,
        [int]$BodyColor,
        [string]$ImagePath,
        [double]$TextWidth = 158,
        [double]$TitleTopOffset = 14,
        [double]$TitleHeight = 28,
        [double]$BodyTopOffset = 50,
        [double]$TitleSize = 16,
        [double]$BodySize = 11.5,
        [double]$ImageBoxLeftOffset = 176,
        [double]$ImageBoxTopOffset = 12,
        [double]$ImageBoxWidth = 108,
        [double]$ImageBoxHeight = 88
    )

    $shape = $Slide.Shapes.AddShape(5, $Left, $Top, $Width, $Height)
    $shape.Name = $Name
    $shape.Fill.ForeColor.RGB = $FillColor
    $shape.Line.Visible = 0
    $shape.Adjustments.Item(1) = 0.15

    if ((-not [string]::IsNullOrWhiteSpace($ImagePath)) -and (Test-Path -LiteralPath $ImagePath)) {
        $placement = Get-FitPlacement -ImagePath $ImagePath -Left ($Left + $ImageBoxLeftOffset) -Top ($Top + $ImageBoxTopOffset) -MaxWidth $ImageBoxWidth -MaxHeight $ImageBoxHeight
        $picture = $Slide.Shapes.AddPicture($ImagePath, $false, $true, $placement.Left, $placement.Top, $placement.Width, $placement.Height)
        $picture.Name = "${Name}_Image"
    }

    $titleBox = Add-GeneratedTextBox -Slide $Slide -Name "${Name}_Title" -Left ($Left + 16) -Top ($Top + $TitleTopOffset) -Width $TextWidth -Height $TitleHeight -Text $Title -Size $TitleSize -Color $TitleColor -Bold $true
    $bodyBox = Add-GeneratedTextBox -Slide $Slide -Name "${Name}_Body" -Left ($Left + 16) -Top ($Top + $BodyTopOffset) -Width $TextWidth -Height ($Height - ($BodyTopOffset + 12)) -Text $Body -Size $BodySize -Color $BodyColor

    @($shape, $titleBox, $bodyBox)
}

function Add-GeneratedArrow {
    param(
        $Slide,
        [string]$Name,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height,
        [int]$Color
    )

    $shape = $Slide.Shapes.AddShape(33, $Left, $Top, $Width, $Height)
    $shape.Name = $Name
    $shape.Fill.ForeColor.RGB = $Color
    $shape.Line.Visible = 0
    $shape
}

function Get-FitPlacement {
    param(
        [string]$ImagePath,
        [double]$Left,
        [double]$Top,
        [double]$MaxWidth,
        [double]$MaxHeight
    )

    $image = [System.Drawing.Image]::FromFile($ImagePath)
    try {
        $widthRatio = $MaxWidth / $image.Width
        $heightRatio = $MaxHeight / $image.Height
        $ratio = [Math]::Min($widthRatio, $heightRatio)
        $targetWidth = [Math]::Round($image.Width * $ratio, 1)
        $targetHeight = [Math]::Round($image.Height * $ratio, 1)
        $targetLeft = [Math]::Round($Left + (($MaxWidth - $targetWidth) / 2), 1)
        $targetTop = [Math]::Round($Top + (($MaxHeight - $targetHeight) / 2), 1)

        @{
            Left = $targetLeft
            Top = $targetTop
            Width = $targetWidth
            Height = $targetHeight
        }
    } finally {
        $image.Dispose()
    }
}

function Replace-ContentPicture {
    param(
        $Slide,
        [string]$ImagePath,
        [double]$Left,
        [double]$Top,
        [double]$MaxWidth,
        [double]$MaxHeight
    )

    Remove-ShapeIfExists -Slide $Slide -Name "Picture 2"
    $placement = Get-FitPlacement -ImagePath $ImagePath -Left $Left -Top $Top -MaxWidth $MaxWidth -MaxHeight $MaxHeight
    $picture = $Slide.Shapes.AddPicture($ImagePath, $false, $true, $placement.Left, $placement.Top, $placement.Width, $placement.Height)
    $picture.Name = "Picture 2"
    $picture
}

function Set-SlideNumber {
    param(
        $Slide,
        [int]$Number
    )

    $shape = $Slide.Shapes.Item("TextBox 5")
    Set-TextBoxText -Shape $shape -Text ([string]$Number) -Size 16 -Color (Get-OleColor 255 255 255) -Bold $true -Alignment 2
}

function Set-HeaderAndLead {
    param(
        $Slide,
        [string]$Header,
        [string]$Lead
    )

    Remove-GeneratedShapes -Slide $Slide
    Remove-ShapeIfExists -Slide $Slide -Name "Picture 2"

    $headerShape = $Slide.Shapes.Item("TextBox 10")
    $leadShape = $Slide.Shapes.Item("TextBox 11")
    $bodyShape = $Slide.Shapes.Item("TextBox 9")

    Set-TextBoxText -Shape $headerShape -Text $Header -Size 19 -Color (Get-OleColor 255 255 255) -Bold $true

    $leadShape.Left = 20
    $leadShape.Top = 74
    $leadShape.Width = 680
    $leadShape.Height = 42
    Set-TextBoxText -Shape $leadShape -Text $Lead -Size 20 -Color (Get-OleColor 23 54 93) -Bold $true

    $bodyShape.Left = 0
    $bodyShape.Top = 0
    $bodyShape.Width = 1
    $bodyShape.Height = 1
    Set-TextBoxText -Shape $bodyShape -Text "" -Size 8 -Color (Get-OleColor 255 255 255)
}

function Configure-RightImageSlide {
    param(
        $Slide,
        [string]$Header,
        [string]$Lead,
        [string]$Body,
        [string]$ImagePath,
        [double]$LeadLeft = 20,
        [double]$LeadTop = 76,
        [double]$LeadWidth = 392,
        [double]$LeadHeight = 58,
        [double]$BodyLeft = 20,
        [double]$BodyTop = 145,
        [double]$BodyWidth = 390,
        [double]$BodyHeight = 272,
        [double]$BodySize = 14,
        [double]$ImageLeft = 430,
        [double]$ImageTop = 92,
        [double]$ImageMaxWidth = 255,
        [double]$ImageMaxHeight = 300
    )

    Remove-GeneratedShapes -Slide $Slide
    $headerShape = $Slide.Shapes.Item("TextBox 10")
    $leadShape = $Slide.Shapes.Item("TextBox 11")
    $bodyShape = $Slide.Shapes.Item("TextBox 9")

    Set-TextBoxText -Shape $headerShape -Text $Header -Size 19 -Color (Get-OleColor 255 255 255) -Bold $true

    $leadShape.Left = $LeadLeft
    $leadShape.Top = $LeadTop
    $leadShape.Width = $LeadWidth
    $leadShape.Height = $LeadHeight
    Set-TextBoxText -Shape $leadShape -Text $Lead -Size 18 -Color (Get-OleColor 23 54 93) -Bold $true

    $bodyShape.Left = $BodyLeft
    $bodyShape.Top = $BodyTop
    $bodyShape.Width = $BodyWidth
    $bodyShape.Height = $BodyHeight
    Set-TextBoxText -Shape $bodyShape -Text $Body -Size $BodySize -Color (Get-OleColor 45 45 45)

    Replace-ContentPicture -Slide $Slide -ImagePath $ImagePath -Left $ImageLeft -Top $ImageTop -MaxWidth $ImageMaxWidth -MaxHeight $ImageMaxHeight | Out-Null
}

function Add-DeckCardGrid {
    param(
        $Slide,
        [object[]]$Cards
    )

    $lefts = @(20, 365)
    $tops = @(128, 235, 342)
    $cardIndex = 0

    foreach ($top in $tops) {
        foreach ($left in $lefts) {
            if ($cardIndex -ge $Cards.Count) {
                return
            }

            $card = $Cards[$cardIndex]
            $height = if ($card.ContainsKey('Height')) { $card.Height } else { 88 }
            $titleHeight = if ($card.ContainsKey('TitleHeight')) { $card.TitleHeight } else { 32 }
            $bodyTopOffset = if ($card.ContainsKey('BodyTopOffset')) { $card.BodyTopOffset } else { 48 }
            $bodyHeight = if ($card.ContainsKey('BodyHeight')) { $card.BodyHeight } else { -1 }
            $bodyLeftOffset = if ($card.ContainsKey('BodyLeftOffset')) { $card.BodyLeftOffset } else { 12 }
            Add-GeneratedCard -Slide $Slide -Name ("Gen_Card_{0}" -f $cardIndex) -Left $left -Top $top -Width 315 -Height $height -FillColor $card.Fill -Title $card.Title -Body $card.Body -TitleColor $card.TitleColor -BodyColor $card.BodyColor -TitleHeight $titleHeight -BodyTopOffset $bodyTopOffset -BodyHeight $bodyHeight -BodyLeftOffset $bodyLeftOffset | Out-Null
            $cardIndex++
        }
    }
}

function Configure-AgendaSlide {
    param($Slide)

    Set-HeaderAndLead -Slide $Slide -Header "Введение" -Lead "Маршрут обучения"
    $cards = @(
        @{ Fill = (Get-OleColor 234 243 255); Title = "1. Роль стропальщика"; Body = "Кто отвечает за подготовку и безопасное перемещение груза."; TitleColor = (Get-OleColor 23 54 93); BodyColor = (Get-OleColor 45 45 45) },
        @{ Fill = (Get-OleColor 255 243 232); Title = "2. Допуск к работе"; Body = "Кого допускают, что нужно проверить до начала работ."; TitleColor = (Get-OleColor 163 79 20); BodyColor = (Get-OleColor 60 60 60) },
        @{ Fill = (Get-OleColor 234 243 255); Title = "3. Сигналы и связь"; Body = "Единый язык команд, сигнальщик, жестовая, голосовая связь, радио связь."; TitleColor = (Get-OleColor 23 54 93); BodyColor = (Get-OleColor 45 45 45) },
        @{ Fill = (Get-OleColor 255 243 232); Title = "4. Оснастка"; Body = "Стропы, захваты, траверсы, тара и критерии браковки."; TitleColor = (Get-OleColor 163 79 20); BodyColor = (Get-OleColor 60 60 60) },
        @{ Fill = (Get-OleColor 234 243 255); Title = "5. Строповка и безопасность"; Body = "Алгоритм работ, типовые ошибки, опасные зоны и аварийные ситуации."; TitleColor = (Get-OleColor 23 54 93); BodyColor = (Get-OleColor 45 45 45); Height = 103.8; TitleHeight = 52; BodyTopOffset = 61.2; BodyHeight = 35.9; BodyLeftOffset = 5 },
        @{ Fill = (Get-OleColor 255 243 232); Title = "6. Аттестация"; Body = "Повторение, критерии зачета и итоговый тест на 15 вопросов."; TitleColor = (Get-OleColor 163 79 20); BodyColor = (Get-OleColor 60 60 60); Height = 103.8; BodyHeight = 35.9 }
    )
    Add-DeckCardGrid -Slide $Slide -Cards $cards
}

function Configure-NormativeSlide {
    param($Slide)

    Set-HeaderAndLead -Slide $Slide -Header "Введение" -Lead "Нормативная база курса"

    $items = @(
        "Правила по охране труда при эксплуатации подъемных сооружений"
        "Требования промышленной безопасности для работ с грузами"
        "Инструкции по охране труда"
        "Технологические карты, схемы строповки грузов"
        "Внутренние требования компании Велесстрой"
    )

    $top = 132
    for ($i = 0; $i -lt $items.Count; $i++) {
        $cardTop = $top + ($i * 61)
        $row = $Slide.Shapes.AddShape(5, 34, $cardTop, 652, 48)
        $row.Name = ("Gen_Norm_{0}" -f $i)
        $row.Fill.ForeColor.RGB = (Get-OleColor 235 240 246)
        $row.Line.Visible = 0
        $badge = $Slide.Shapes.AddShape(9, 44, $cardTop + 10, 24, 24)
        $badge.Name = ("Gen_NormBadge_{0}" -f $i)
        $badge.Fill.ForeColor.RGB = (Get-OleColor 20 83 144)
        $badge.Line.Visible = 0
        Add-GeneratedTextBox -Slide $Slide -Name ("Gen_NormNum_{0}" -f $i) -Left 46 -Top ($cardTop + 10) -Width 24 -Height 24 -Text ([string]($i + 1)) -Size 12 -Color (Get-OleColor 255 255 255) -Bold $true -Alignment 2 | Out-Null
        Add-GeneratedTextBox -Slide $Slide -Name ("Gen_NormText_{0}" -f $i) -Left 82 -Top ($cardTop + 6) -Width 585 -Height 34 -Text $items[$i] -Size 13 -Color (Get-OleColor 45 45 45) | Out-Null
    }
}

function Configure-LearningFlowSlide {
    param($Slide)

    Set-HeaderAndLead -Slide $Slide -Header "Введение" -Lead "Как проходит обучение"

    Add-GeneratedCard -Slide $Slide -Name "Gen_Theory" -Left 38 -Top 180 -Width 180 -Height 160 -FillColor (Get-OleColor 234 243 255) -Title "Теория" -Body "Основы профессии, оснастка, сигналы, правила и опасные зоны." -TitleColor (Get-OleColor 23 54 93) -BodyColor (Get-OleColor 45 45 45) | Out-Null
    Add-GeneratedArrow -Slide $Slide -Name "Gen_Arrow1" -Left 238 -Top 232 -Width 60 -Height 48 -Color (Get-OleColor 246 145 32) | Out-Null
    Add-GeneratedCard -Slide $Slide -Name "Gen_Practice" -Left 300 -Top 180 -Width 180 -Height 160 -FillColor (Get-OleColor 255 243 232) -Title "Практика" -Body "Алгоритм работ стропальщика, выбор оснастки и безопасные действия, знаковая сигнолизация." -TitleColor (Get-OleColor 163 79 20) -BodyColor (Get-OleColor 60 60 60) | Out-Null
    Add-GeneratedArrow -Slide $Slide -Name "Gen_Arrow2" -Left 500 -Top 232 -Width 60 -Height 48 -Color (Get-OleColor 246 145 32) | Out-Null
    Add-GeneratedCard -Slide $Slide -Name "Gen_Test" -Left 562 -Top 180 -Width 120 -Height 160 -FillColor (Get-OleColor 234 243 255) -Title "Тест" -Body "Итоговая аттестация и проверка понимания ключевых правил." -TitleColor (Get-OleColor 23 54 93) -BodyColor (Get-OleColor 45 45 45) | Out-Null
    Add-GeneratedTextBox -Slide $Slide -Name "Gen_FlowNote" -Left 80 -Top 368 -Width 570 -Height 54 -Text "Логика курса простая: сначала понимаем правила, затем видим их на практических примерах и только после этого переходим к аттестации." -Size 13 -Color (Get-OleColor 60 60 60) | Out-Null
}

function Configure-EquipmentOverviewSlide {
    param($Slide)

    Set-HeaderAndLead -Slide $Slide -Header "Общие сведения о подъемных сооружениях" -Lead "Какие подъемные сооружения встречаются в работе стропальщика"

    Add-PhotoInfoCard -Slide $Slide -Name "Gen_Equip0" -Left 40 -Top 160 -Width 300 -Height 126 -FillColor (Get-OleColor 234 243 255) -Title "Автокран" -Body "Самый частый вариант на строительной площадке." -TitleColor (Get-OleColor 23 54 93) -BodyColor (Get-OleColor 45 45 45) -ImagePath $autocraneImage -TextWidth 152 -ImageBoxLeftOffset 165 -ImageBoxTopOffset 10 -ImageBoxWidth 128 -ImageBoxHeight 94 | Out-Null
    Add-PhotoInfoCard -Slide $Slide -Name "Gen_Equip1" -Left 380 -Top 160 -Width 300 -Height 126 -FillColor (Get-OleColor 255 243 232) -Title "Мостовой кран" -Body "Работа в цехах, на складах и внутри производственных помещений." -TitleColor (Get-OleColor 163 79 20) -BodyColor (Get-OleColor 60 60 60) -ImagePath $bridgeCraneImage -TextWidth 156 -ImageBoxLeftOffset 168 -ImageBoxTopOffset 14 -ImageBoxWidth 124 -ImageBoxHeight 85 | Out-Null
    Add-PhotoInfoCard -Slide $Slide -Name "Gen_Equip2" -Left 40 -Top 306 -Width 300 -Height 126 -FillColor (Get-OleColor 234 243 255) -Title "Кран-манипулятор" -Body "Подъем и подача грузов на ограниченной площадке." -TitleColor (Get-OleColor 23 54 93) -BodyColor (Get-OleColor 45 45 45) -ImagePath $craneManipulatorImage -TextWidth 152 -TitleHeight 34 -BodyTopOffset 64 -ImageBoxLeftOffset 165 -ImageBoxTopOffset 12 -ImageBoxWidth 128 -ImageBoxHeight 97 | Out-Null
    Add-PhotoInfoCard -Slide $Slide -Name "Gen_Equip3" -Left 380 -Top 306 -Width 300 -Height 126 -FillColor (Get-OleColor 255 243 232) -Title "Трубоукладчик" -Body "Специальная техника для работ с трубами и длинномерными грузами." -TitleColor (Get-OleColor 163 79 20) -BodyColor (Get-OleColor 60 60 60) -ImagePath $pipeLayerImage -TextWidth 156 -ImageBoxLeftOffset 166 -ImageBoxTopOffset 8 -ImageBoxWidth 126 -ImageBoxHeight 108 | Out-Null
}

function Configure-CraneKnowledgeSlide {
    param($Slide)

    Set-HeaderAndLead -Slide $Slide -Header "Общие сведения о подъемных сооружениях" -Lead "Что стропальщик должен понимать о возможностях крана"

    Add-GeneratedCard -Slide $Slide -Name "Gen_CraneCore" -Left 204 -Top 154 -Width 312 -Height 104 -FillColor (Get-OleColor 20 83 144) -Title "Перед подъемом" -Body "Стропальщик оценивает не только технические характеристики крана, но и условия безопасной работы." -TitleColor (Get-OleColor 255 255 255) -BodyColor (Get-OleColor 255 255 255) -TitleSize 18 -BodySize 11.5 -TitleHeight 28 -BodyTopOffset 48 -BodyHeight 44 | Out-Null

    $cards = @(
        @{ Name = "Gen_Crane1"; Left = 20; Fill = (Get-OleColor 234 243 255); Label = "Грузоподъемность"; LabelColor = (Get-OleColor 23 54 93); LabelLeft = 8; LabelWidth = 178; LabelSize = 12.2; Body = "Хватает ли ее для груза с учетом условий подъема."; BodyColor = (Get-OleColor 45 45 45) },
        @{ Name = "Gen_Crane2"; Left = 196; Fill = (Get-OleColor 255 243 232); Label = "Вылет стрелы"; LabelColor = (Get-OleColor 163 79 20); LabelSize = 13.5; Body = "Меняет допустимую нагрузку и запас безопасности."; BodyColor = (Get-OleColor 60 60 60) },
        @{ Name = "Gen_Crane3"; Left = 372; Fill = (Get-OleColor 234 243 255); Label = "Рабочая зона"; LabelColor = (Get-OleColor 23 54 93); LabelSize = 13.5; Body = "Где проходит траектория груза и кто попадает в опасную зону."; BodyColor = (Get-OleColor 45 45 45) },
        @{ Name = "Gen_Crane4"; Left = 548; Fill = (Get-OleColor 255 243 232); Label = "Ограничения"; LabelColor = (Get-OleColor 163 79 20); LabelSize = 13.5; Body = "Площадка, опоры, препятствия, ЛЭП, обзорность и связь."; BodyColor = (Get-OleColor 60 60 60) }
    )

    foreach ($card in $cards) {
        $labelLeft = if ($card.ContainsKey("LabelLeft")) { $card.LabelLeft } else { $card.Left }
        $labelWidth = if ($card.ContainsKey("LabelWidth")) { $card.LabelWidth } else { 152 }
        $labelSize = if ($card.ContainsKey("LabelSize")) { $card.LabelSize } else { 13.5 }
        Add-GeneratedTextBox -Slide $Slide -Name ("{0}_Label" -f $card.Name) -Left $labelLeft -Top 268 -Width $labelWidth -Height 22 -Text $card.Label -Size $labelSize -Color $card.LabelColor -Bold $true | Out-Null

        $panel = $Slide.Shapes.AddShape(5, $card.Left, 300, 152, 112)
        $panel.Name = ("{0}_Panel" -f $card.Name)
        $panel.Fill.ForeColor.RGB = $card.Fill
        $panel.Line.Visible = 0
        $panel.Adjustments.Item(1) = 0.15

        Add-GeneratedTextBox -Slide $Slide -Name ("{0}_Body" -f $card.Name) -Left ($card.Left + 12) -Top 316 -Width 128 -Height 78 -Text $card.Body -Size 11.5 -Color $card.BodyColor | Out-Null
    }
}

function Configure-DangerZonesSlide {
    param(
        $Slide,
        [string]$ImagePath
    )

    Configure-RightImageSlide -Slide $Slide -Header "Общие сведения о подъемных сооружениях" -Lead "Опасные зоны при работе крана" -Body "Наибольший риск возникает:`r`n`r`n- под подвешенным грузом;`r`n- рядом с траекторией его перемещения;`r`n- в зоне поворота стрелы;`r`n- возле препятствий и неогражденных участков." -ImagePath $ImagePath -BodyHeight 150 -ImageLeft 454 -ImageTop 92 -ImageMaxWidth 207 -ImageMaxHeight 300

    $noteBox = $Slide.Shapes.AddShape(5, 20, 332, 390, 92)
    $noteBox.Name = "Gen_DangerNoteBox"
    $noteBox.Fill.ForeColor.RGB = (Get-OleColor 234 243 255)
    $noteBox.Fill.Transparency = 0.08
    $noteBox.Line.Visible = 0
    $noteBox.Adjustments.Item(1) = 0.15

    $accent = $Slide.Shapes.AddShape(5, 20, 332, 10, 92)
    $accent.Name = "Gen_DangerNoteAccent"
    $accent.Fill.ForeColor.RGB = (Get-OleColor 246 145 32)
    $accent.Line.Visible = 0
    $accent.Adjustments.Item(1) = 0.15

    Add-GeneratedTextBox -Slide $Slide -Name "Gen_DangerNoteTitle" -Left 40 -Top 343 -Width 130 -Height 18 -Text "Примечание" -Size 11.5 -Color (Get-OleColor 163 79 20) -Bold $true | Out-Null
    Add-GeneratedTextBox -Slide $Slide -Name "Gen_DangerNoteBody" -Left 40 -Top 362 -Width 354 -Height 48 -Text "- от крана до ближайшего препятствия или конструкции расстояние должно быть не менее 1 м;`r`n- при переносе груза расстояние от груза до препятствия должно быть не менее 50 см." -Size 10.8 -Color (Get-OleColor 60 60 60) | Out-Null
}

function Configure-CoordinationSlide {
    param(
        $Slide,
        [string]$ImagePath
    )

    Configure-RightImageSlide -Slide $Slide -Header "Общие сведения о подъемных сооружениях" -Lead "Безопасный подъем требует согласованных действий" -Body "Даже исправная техника не компенсирует неправильные команды.`r`n`r`nПеред началом работ участники должны:`r`n- договориться о едином языке команд;`r`n- определить, кто подает сигналы;`r`n- подтвердить порядок связи и остановки." -ImagePath $ImagePath -BodyHeight 206 -ImageLeft 430 -ImageTop 104 -ImageMaxWidth 255 -ImageMaxHeight 250

    $noteBox = $Slide.Shapes.AddShape(5, 38, 386, 644, 48)
    $noteBox.Name = "Gen_CoordinationNoteBox"
    $noteBox.Fill.ForeColor.RGB = (Get-OleColor 255 243 232)
    $noteBox.Fill.Transparency = 0.06
    $noteBox.Line.Visible = 0
    $noteBox.Adjustments.Item(1) = 0.15

    Add-GeneratedTextBox -Slide $Slide -Name "Gen_CoordinationNote" -Left 56 -Top 398 -Width 606 -Height 23 -Text "Отсутствие взаимопонимания между машинистом крана и стропальщиком - частая причина аварий на площадке." -Size 11.5 -Color (Get-OleColor 163 79 20) -Bold $true -Alignment 2 | Out-Null
}

function Configure-AdmissionSlide {
    param($Slide)

    Set-HeaderAndLead -Slide $Slide -Header "Допуск и организация работ" -Lead "Кто может быть допущен к работе стропальщиком"

    Add-GeneratedCard -Slide $Slide -Name "Gen_AdmissionCore" -Left 176 -Top 136 -Width 368 -Height 72 -FillColor (Get-OleColor 20 83 144) -Title "К работе допускаются" -Body "Только работники, которые выполнили все условия допуска к подъемным работам." -TitleColor (Get-OleColor 255 255 255) -BodyColor (Get-OleColor 255 255 255) -TitleSize 17 -BodySize 10.8 -TitleHeight 24 -BodyTopOffset 38 -BodyHeight 24 | Out-Null

    Add-GeneratedCard -Slide $Slide -Name "Gen_Admission1" -Left 46 -Top 232 -Width 290 -Height 102 -FillColor (Get-OleColor 234 243 255) -Title "18+ и медицинский допуск" -Body "Работник не моложе 18 лет и не имеет противопоказаний к выполнению работ." -TitleColor (Get-OleColor 23 54 93) -BodyColor (Get-OleColor 45 45 45) -TitleSize 15.5 -BodySize 11.2 -TitleHeight 28 -BodyTopOffset 46 -BodyHeight 40 | Out-Null
    Add-GeneratedCard -Slide $Slide -Name "Gen_Admission2" -Left 384 -Top 232 -Width 290 -Height 102 -FillColor (Get-OleColor 255 243 232) -Title "Обучение и инструктажи" -Body "Прошедшие обучение по учебной программе и инструктажи по безопасности труда." -TitleColor (Get-OleColor 163 79 20) -BodyColor (Get-OleColor 60 60 60) -TitleSize 15.5 -BodySize 11.1 -TitleHeight 28 -BodyTopOffset 46 -BodyHeight 40 | Out-Null
    Add-GeneratedCard -Slide $Slide -Name "Gen_Admission3" -Left 46 -Top 344 -Width 290 -Height 102 -FillColor (Get-OleColor 255 243 232) -Title "Производственная стажировка" -Body "Получившие производственную стажировку под руководством опытных стропальщиков." -TitleColor (Get-OleColor 163 79 20) -BodyColor (Get-OleColor 60 60 60) -TitleSize 15 -BodySize 11 -TitleHeight 30 -BodyTopOffset 48 -BodyHeight 40 | Out-Null
    Add-GeneratedCard -Slide $Slide -Name "Gen_Admission4" -Left 384 -Top 344 -Width 290 -Height 102 -FillColor (Get-OleColor 234 243 255) -Title "Удостоверение при себе" -Body "Имеющие удостоверение на право производства работ, которое он обязан во время работы иметь при себе." -TitleColor (Get-OleColor 23 54 93) -BodyColor (Get-OleColor 45 45 45) -TitleSize 15 -BodySize 10.9 -TitleHeight 28 -BodyTopOffset 46 -BodyHeight 42 | Out-Null

    $noteBox = $Slide.Shapes.AddShape(5, 82, 460, 556, 42)
    $noteBox.Name = "Gen_AdmissionNoteBox"
    $noteBox.Fill.ForeColor.RGB = (Get-OleColor 255 243 232)
    $noteBox.Line.Visible = 0
    $noteBox.Adjustments.Item(1) = 0.15

    Add-GeneratedTextBox -Slide $Slide -Name "Gen_AdmissionNote" -Left 98 -Top 468 -Width 522 -Height 28 -Text "Если хотя бы одно условие не выполнено, работника к подъему груза не допускают.`r`nРаботать без удостоверения запрещается." -Size 10.8 -Color (Get-OleColor 163 79 20) -Bold $true -Alignment 2 | Out-Null
}

function Configure-PrestartChecksSlide {
    param($Slide)

    Set-HeaderAndLead -Slide $Slide -Header "Допуск и организация работ" -Lead "Что необходимо проверить до начала работы"

    $cards = @(
        @{ Left = 28; Top = 152; Fill = (Get-OleColor 234 243 255); Title = "1. Задача понятна"; Body = "Известно, что поднимать, куда перемещать и как устанавливать груз."; TitleColor = (Get-OleColor 23 54 93); BodyColor = (Get-OleColor 45 45 45) },
        @{ Left = 256; Top = 152; Fill = (Get-OleColor 255 243 232); Title = "2. Документы на месте"; Body = "Есть наряд-допуск, разрешение на работу и указания ответственного лица."; TitleColor = (Get-OleColor 163 79 20); BodyColor = (Get-OleColor 60 60 60) },
        @{ Left = 484; Top = 152; Fill = (Get-OleColor 234 243 255); Title = "3. Масса груза известна"; Body = "Груз не поднимают на глаз и без подтвержденной массы."; TitleColor = (Get-OleColor 23 54 93); BodyColor = (Get-OleColor 45 45 45) },
        @{ Left = 28; Top = 294; Fill = (Get-OleColor 255 243 232); Title = "4. Схема строповки понятна"; Body = "Выбрана правильная оснастка и понятны точки захвата груза."; TitleColor = (Get-OleColor 163 79 20); BodyColor = (Get-OleColor 60 60 60) },
        @{ Left = 256; Top = 294; Fill = (Get-OleColor 234 243 255); Title = "5. Команды согласованы"; Body = "Определено, кто подает сигналы, и как поддерживается связь."; TitleColor = (Get-OleColor 23 54 93); BodyColor = (Get-OleColor 45 45 45) },
        @{ Left = 484; Top = 294; Fill = (Get-OleColor 255 243 232); Title = "6. Зона безопасна"; Body = "Путь перемещения свободен, люди выведены из опасной зоны."; TitleColor = (Get-OleColor 163 79 20); BodyColor = (Get-OleColor 60 60 60) }
    )

    foreach ($card in $cards) {
        Add-GeneratedCard -Slide $Slide -Name ("Gen_Prestart_" + ($card.Title -replace '[^0-9A-Za-zА-Яа-я]', '')) -Left $card.Left -Top $card.Top -Width 208 -Height 122 -FillColor $card.Fill -Title $card.Title -Body $card.Body -TitleColor $card.TitleColor -BodyColor $card.BodyColor -TitleSize 14.2 -BodySize 11.1 -TitleHeight 32 -BodyTopOffset 54 -BodyHeight 48 | Out-Null
    }

    $noteBox = $Slide.Shapes.AddShape(5, 66, 438, 588, 40)
    $noteBox.Name = "Gen_PrestartNoteBox"
    $noteBox.Fill.ForeColor.RGB = (Get-OleColor 234 243 255)
    $noteBox.Line.Visible = 0
    $noteBox.Adjustments.Item(1) = 0.15

    Add-GeneratedTextBox -Slide $Slide -Name "Gen_PrestartNote" -Left 88 -Top 448 -Width 544 -Height 20 -Text "Если хотя бы один пункт не подтвержден, к подъему груза не приступают." -Size 11.3 -Color (Get-OleColor 23 54 93) -Bold $true -Alignment 2 | Out-Null
}

function Configure-ParticipantsSlide {
    param($Slide)

    Set-HeaderAndLead -Slide $Slide -Header "Допуск и организация работ" -Lead "Участники процесса подъема и взаимодействие на площадке"
    $leadShape = $Slide.Shapes.Item("TextBox 11")
    $leadShape.Left = 18
    $leadShape.Top = 74
    $leadShape.Width = 704
    $leadShape.Height = 52
    Set-TextBoxText -Shape $leadShape -Text "Участники процесса подъема и взаимодействие на площадке" -Size 18.5 -Color (Get-OleColor 23 54 93) -Bold $true -Alignment 2

    Add-GeneratedCard -Slide $Slide -Name "Gen_ParticipantTop" -Left 145 -Top 138 -Width 430 -Height 86 -FillColor (Get-OleColor 20 83 144) -Title "Ответственный за безопасное производство работ" -Body "Определяет порядок работ, разрешает подъем и контролирует безопасность." -TitleColor (Get-OleColor 255 255 255) -BodyColor (Get-OleColor 255 255 255) -TitleSize 12.8 -BodySize 10.1 -TitleHeight 32 -BodyTopOffset 46 -BodyHeight 26 | Out-Null

    $line1 = $Slide.Shapes.AddShape(1, 358, 212, 4, 26)
    $line1.Name = "Gen_ParticipantLine1"
    $line1.Fill.ForeColor.RGB = (Get-OleColor 200 210 220)
    $line1.Line.Visible = 0
    $line2 = $Slide.Shapes.AddShape(1, 178, 238, 364, 4)
    $line2.Name = "Gen_ParticipantLine2"
    $line2.Fill.ForeColor.RGB = (Get-OleColor 200 210 220)
    $line2.Line.Visible = 0
    $line3 = $Slide.Shapes.AddShape(1, 178, 238, 4, 18)
    $line3.Name = "Gen_ParticipantLine3"
    $line3.Fill.ForeColor.RGB = (Get-OleColor 200 210 220)
    $line3.Line.Visible = 0
    $line4 = $Slide.Shapes.AddShape(1, 538, 238, 4, 18)
    $line4.Name = "Gen_ParticipantLine4"
    $line4.Fill.ForeColor.RGB = (Get-OleColor 200 210 220)
    $line4.Line.Visible = 0
    $line5 = $Slide.Shapes.AddShape(1, 358, 242, 4, 126)
    $line5.Name = "Gen_ParticipantLine5"
    $line5.Fill.ForeColor.RGB = (Get-OleColor 200 210 220)
    $line5.Line.Visible = 0

    Add-GeneratedCard -Slide $Slide -Name "Gen_ParticipantMaster" -Left 34 -Top 256 -Width 288 -Height 92 -FillColor (Get-OleColor 255 243 232) -Title "Мастер / производитель работ" -Body "Ставит задачу, организует место работ и координирует участников." -TitleColor (Get-OleColor 163 79 20) -BodyColor (Get-OleColor 60 60 60) -TitleSize 14.2 -BodySize 10.6 -TitleHeight 30 -BodyTopOffset 48 -BodyHeight 32 | Out-Null
    Add-GeneratedCard -Slide $Slide -Name "Gen_ParticipantOperator" -Left 398 -Top 256 -Width 288 -Height 92 -FillColor (Get-OleColor 234 243 255) -Title "Машинист крана" -Body "Выполняет подъем только по согласованным командам и в пределах возможностей крана." -TitleColor (Get-OleColor 23 54 93) -BodyColor (Get-OleColor 45 45 45) -TitleSize 14.2 -BodySize 10.5 -TitleHeight 30 -BodyTopOffset 48 -BodyHeight 34 | Out-Null
    Add-GeneratedCard -Slide $Slide -Name "Gen_ParticipantSignalman" -Left 34 -Top 378 -Width 200 -Height 84 -FillColor (Get-OleColor 234 243 255) -Title "Сигнальщик" -Body "Передает сигналы, если обзор ограничен." -TitleColor (Get-OleColor 23 54 93) -BodyColor (Get-OleColor 45 45 45) -TitleSize 14.2 -BodySize 10.4 -TitleHeight 26 -BodyTopOffset 44 -BodyHeight 28 | Out-Null
    Add-GeneratedCard -Slide $Slide -Name "Gen_ParticipantSlinger" -Left 260 -Top 366 -Width 200 -Height 108 -FillColor (Get-OleColor 255 243 232) -Title "Стропальщик" -Body "Выбирает оснастку, стропит груз, подает команды и сопровождает подъем." -TitleColor (Get-OleColor 163 79 20) -BodyColor (Get-OleColor 60 60 60) -TitleSize 14.2 -BodySize 10.4 -TitleHeight 26 -BodyTopOffset 42 -BodyHeight 52 | Out-Null
    Add-GeneratedCard -Slide $Slide -Name "Gen_ParticipantContact" -Left 486 -Top 378 -Width 200 -Height 84 -FillColor (Get-OleColor 234 243 255) -Title "К кому обращаться" -Body "По вопросам организации и безопасности - к мастеру или ответственному лицу." -TitleColor (Get-OleColor 23 54 93) -BodyColor (Get-OleColor 45 45 45) -TitleSize 14.2 -BodySize 10.1 -TitleHeight 26 -BodyTopOffset 44 -BodyHeight 34 | Out-Null

    $noteBox = $Slide.Shapes.AddShape(5, 48, 474, 624, 28)
    $noteBox.Name = "Gen_ParticipantNoteBox"
    $noteBox.Fill.ForeColor.RGB = (Get-OleColor 255 243 232)
    $noteBox.Line.Visible = 0
    $noteBox.Adjustments.Item(1) = 0.15

    Add-GeneratedTextBox -Slide $Slide -Name "Gen_ParticipantNote" -Left 70 -Top 480 -Width 580 -Height 18 -Text "Самовольные действия и противоречивые команды при подъеме груза недопустимы." -Size 11.2 -Color (Get-OleColor 163 79 20) -Bold $true -Alignment 2 | Out-Null
}

function Configure-CommandAuthoritySlide {
    param($Slide)

    Set-HeaderAndLead -Slide $Slide -Header "Допуск и организация работ" -Lead "Кто подает команды машинисту крана"

    Add-GeneratedCard -Slide $Slide -Name "Gen_CommandTop" -Left 114 -Top 148 -Width 492 -Height 74 -FillColor (Get-OleColor 20 83 144) -Title "Главное правило" -Body "Во время подъема команды машинисту крана подает один назначенный стропальщик." -TitleColor (Get-OleColor 255 255 255) -BodyColor (Get-OleColor 255 255 255) -TitleSize 16 -BodySize 11.3 -TitleHeight 24 -BodyTopOffset 40 -BodyHeight 24 | Out-Null

    Add-GeneratedCard -Slide $Slide -Name "Gen_CommandRegular" -Left 48 -Top 260 -Width 286 -Height 106 -FillColor (Get-OleColor 234 243 255) -Title "Обычная ситуация" -Body "Команды дает стропальщик, который руководит строповкой и сопровождает подъем груза." -TitleColor (Get-OleColor 23 54 93) -BodyColor (Get-OleColor 45 45 45) -TitleSize 15 -BodySize 11 -TitleHeight 26 -BodyTopOffset 44 -BodyHeight 42 | Out-Null
    Add-GeneratedCard -Slide $Slide -Name "Gen_CommandStop" -Left 386 -Top 260 -Width 286 -Height 106 -FillColor (Get-OleColor 255 243 232) -Title 'Исключение: команда "Стоп"' -Body "Если кто-то заметил опасность, команду немедленной остановки может подать любой работник." -TitleColor (Get-OleColor 163 79 20) -BodyColor (Get-OleColor 60 60 60) -TitleSize 15 -BodySize 10.9 -TitleHeight 28 -BodyTopOffset 46 -BodyHeight 40 | Out-Null

    $stopPlacement = Get-FitPlacement -ImagePath $stopSignalImage -Left 540 -Top 142 -MaxWidth 110 -MaxHeight 104
    $stopPicture = $Slide.Shapes.AddPicture($stopSignalImage, $false, $true, $stopPlacement.Left, $stopPlacement.Top, $stopPlacement.Width, $stopPlacement.Height)
    $stopPicture.Name = "Gen_CommandStopImage"

    $noteBox = $Slide.Shapes.AddShape(5, 72, 388, 576, 72)
    $noteBox.Name = "Gen_CommandNoteBox"
    $noteBox.Fill.ForeColor.RGB = (Get-OleColor 234 243 255)
    $noteBox.Line.Visible = 0
    $noteBox.Adjustments.Item(1) = 0.15

    Add-GeneratedTextBox -Slide $Slide -Name "Gen_CommandNoteTitle" -Left 92 -Top 400 -Width 250 -Height 28 -Text "Если крановщик не видит стропальщика" -Size 10.8 -Color (Get-OleColor 23 54 93) -Bold $true | Out-Null
    Add-GeneratedTextBox -Slide $Slide -Name "Gen_CommandNoteBody" -Left 92 -Top 432 -Width 534 -Height 28 -Text "Назначается сигнальщик из числа стропальщиков. Его назначает ответственный руководитель работ." -Size 10.5 -Color (Get-OleColor 45 45 45) | Out-Null
}

function Configure-CommandLanguageSlide {
    param($Slide)

    Set-HeaderAndLead -Slide $Slide -Header "Допуск и организация работ" -Lead "Единый язык команд до начала работ"

    Add-GeneratedCard -Slide $Slide -Name "Gen_LanguageTop" -Left 116 -Top 146 -Width 488 -Height 74 -FillColor (Get-OleColor 20 83 144) -Title "До начала работ договориться обязательно" -Body "Машинист крана и стропальщик заранее согласуют понятный для обоих язык команд." -TitleColor (Get-OleColor 255 255 255) -BodyColor (Get-OleColor 255 255 255) -TitleSize 15.5 -BodySize 11.1 -TitleHeight 24 -BodyTopOffset 40 -BodyHeight 24 | Out-Null

    Add-GeneratedCard -Slide $Slide -Name "Gen_LanguageAgree" -Left 46 -Top 258 -Width 300 -Height 130 -FillColor (Get-OleColor 234 243 255) -Title "Что нужно согласовать" -Body "- основные команды подъема и остановки;`r`n- кто подает сигналы;`r`n- как подтверждается команда;`r`n- что делать при непонимании." -TitleColor (Get-OleColor 23 54 93) -BodyColor (Get-OleColor 45 45 45) -TitleSize 15 -BodySize 10.8 -TitleHeight 26 -BodyTopOffset 44 -BodyHeight 72 | Out-Null
    Add-GeneratedCard -Slide $Slide -Name "Gen_LanguageWhy" -Left 374 -Top 258 -Width 300 -Height 130 -FillColor (Get-OleColor 255 243 232) -Title "Почему это критично" -Body "На площадке могут работать люди с разным опытом и разным уровнем владения профессией.`r`nНепонятая команда быстро превращается в риск." -TitleColor (Get-OleColor 163 79 20) -BodyColor (Get-OleColor 60 60 60) -TitleSize 15 -BodySize 10.5 -TitleHeight 26 -BodyTopOffset 44 -BodyHeight 72 | Out-Null

    $warnPlacement = Get-FitPlacement -ImagePath $warningSignalImage -Left 520 -Top 140 -MaxWidth 120 -MaxHeight 106
    $warnPicture = $Slide.Shapes.AddPicture($warningSignalImage, $false, $true, $warnPlacement.Left, $warnPlacement.Top, $warnPlacement.Width, $warnPlacement.Height)
    $warnPicture.Name = "Gen_LanguageWarnImage"

    $noteBox = $Slide.Shapes.AddShape(5, 86, 406, 548, 42)
    $noteBox.Name = "Gen_LanguageNoteBox"
    $noteBox.Fill.ForeColor.RGB = (Get-OleColor 255 243 232)
    $noteBox.Line.Visible = 0
    $noteBox.Adjustments.Item(1) = 0.15

    Add-GeneratedTextBox -Slide $Slide -Name "Gen_LanguageNote" -Left 104 -Top 418 -Width 512 -Height 20 -Text "Если команда не понята однозначно, подъем не продолжают до уточнения." -Size 11.2 -Color (Get-OleColor 163 79 20) -Bold $true -Alignment 2 | Out-Null
}

function Configure-CommunicationTypesSlide {
    param($Slide)

    Set-HeaderAndLead -Slide $Slide -Header "Допуск и организация работ" -Lead "Виды связи при выполнении работ"

    Add-GeneratedCard -Slide $Slide -Name "Gen_CommIntro" -Left 126 -Top 138 -Width 468 -Height 82 -FillColor (Get-OleColor 20 83 144) -Title "Способ связи выбирают по обзору, шуму и расстоянию" -Body "Главная цель - чтобы команда была понятна сразу и без искажений." -TitleColor (Get-OleColor 255 255 255) -BodyColor (Get-OleColor 255 255 255) -TitleSize 13.6 -BodySize 10.2 -TitleHeight 36 -BodyTopOffset 48 -BodyHeight 22 | Out-Null

    Add-GeneratedCard -Slide $Slide -Name "Gen_CommSigns" -Left 32 -Top 236 -Width 212 -Height 186 -FillColor (Get-OleColor 234 243 255) -Title "Знаковая сигнализация" -Body "Применяется, когда машинист видит стропальщика.`r`n`r`nОсновной способ связи при подъеме груза.`r`n`r`nИспользуется на дистанции от 10 до 22 м." -TitleColor (Get-OleColor 23 54 93) -BodyColor (Get-OleColor 45 45 45) -TitleSize 14.4 -BodySize 10.5 -TitleHeight 26 -BodyTopOffset 46 -BodyHeight 116 | Out-Null
    Add-GeneratedCard -Slide $Slide -Name "Gen_CommVoice" -Left 254 -Top 236 -Width 212 -Height 186 -FillColor (Get-OleColor 255 243 232) -Title "Голосовая связь" -Body "Подходит на короткой дистанции и при низком уровне шума.`r`n`r`nИспользуется только если команды слышны без сомнений.`r`n`r`nИспользуется на дистанции до 10 м." -TitleColor (Get-OleColor 163 79 20) -BodyColor (Get-OleColor 60 60 60) -TitleSize 14.4 -BodySize 10.2 -TitleHeight 26 -BodyTopOffset 46 -BodyHeight 120 | Out-Null
    Add-GeneratedCard -Slide $Slide -Name "Gen_CommRadio" -Left 476 -Top 236 -Width 212 -Height 186 -FillColor (Get-OleColor 234 243 255) -Title "Радиосвязь" -Body "Нужна, когда обзор ограничен, расстояние большое или площадка шумная.`r`n`r`nЕсли связь пропала - подъем останавливают.`r`n`r`nИспользуется на дистанции свыше 22 м." -TitleColor (Get-OleColor 23 54 93) -BodyColor (Get-OleColor 45 45 45) -TitleSize 14.4 -BodySize 10.2 -TitleHeight 26 -BodyTopOffset 46 -BodyHeight 120 | Out-Null

    $noteBox = $Slide.Shapes.AddShape(5, 94, 414, 532, 34)
    $noteBox.Name = "Gen_CommNoteBox"
    $noteBox.Fill.ForeColor.RGB = (Get-OleColor 234 243 255)
    $noteBox.Line.Visible = 0
    $noteBox.Adjustments.Item(1) = 0.15

    Add-GeneratedTextBox -Slide $Slide -Name "Gen_CommNote" -Left 110 -Top 421 -Width 500 -Height 22 -Text "Неустойчивая связь и потеря визуального контакта - основание немедленно остановить подъем." -Size 10.7 -Color (Get-OleColor 23 54 93) -Bold $true -Alignment 2 | Out-Null
}

function Configure-SignalCheatsheetSlide {
    param($Slide)

    Set-HeaderAndLead -Slide $Slide -Header "Допуск и организация работ" -Lead "Знаковая сигнализация"

    $leadShape = $Slide.Shapes.Item("TextBox 11")
    $leadShape.Left = 40
    $leadShape.Top = 74
    $leadShape.Width = 640
    $leadShape.Height = 42
    Set-TextBoxText -Shape $leadShape -Text "Знаковая сигнализация" -Size 22 -Color (Get-OleColor 23 54 93) -Bold $true -Alignment 2

    $frame = $Slide.Shapes.AddShape(1, 76, 126, 568, 356)
    $frame.Name = "Gen_SignalSheetFrame"
    $frame.Fill.ForeColor.RGB = (Get-OleColor 248 250 253)
    $frame.Line.ForeColor.RGB = (Get-OleColor 205 214 225)
    $frame.Line.Weight = 1.5

    $sheetPlacement = Get-FitPlacement -ImagePath $signalCheatsheetImage -Left 88 -Top 138 -MaxWidth 544 -MaxHeight 332
    $sheetPicture = $Slide.Shapes.AddPicture($signalCheatsheetImage, $false, $true, $sheetPlacement.Left, $sheetPlacement.Top, $sheetPlacement.Width, $sheetPlacement.Height)
    $sheetPicture.Name = "Gen_SignalSheetLarge"
}

function Configure-SafeStartAlgorithmSlide {
    param($Slide)

    Set-HeaderAndLead -Slide $Slide -Header "Подготовка к подъему" -Lead "Общий алгоритм безопасного начала работ"

    $heroPlacement = Get-FitPlacement -ImagePath $algorithmBackgroundImage -Left 500 -Top 116 -MaxWidth 188 -MaxHeight 136
    $hero = $Slide.Shapes.AddPicture($algorithmBackgroundImage, $false, $true, $heroPlacement.Left, $heroPlacement.Top, $heroPlacement.Width, $heroPlacement.Height)
    $hero.Name = "Gen_StartHero"

    Add-GeneratedCard -Slide $Slide -Name "Gen_StartPermit" -Left 156 -Top 124 -Width 412 -Height 58 -FillColor (Get-OleColor 20 83 144) -Title "Работы начинают только после задания и наряда-допуска" -Body "" -TitleColor (Get-OleColor 255 255 255) -BodyColor (Get-OleColor 255 255 255) -TitleSize 13.4 -TitleHeight 30 -BodyTopOffset 36 -BodyHeight 0 | Out-Null

    Add-GeneratedCard -Slide $Slide -Name "Gen_StartStep1" -Left 34 -Top 226 -Width 188 -Height 92 -FillColor (Get-OleColor 234 243 255) -Title "1. Получить задание" -Body "Уточнить груз, место работ и ответственных лиц." -TitleColor (Get-OleColor 23 54 93) -BodyColor (Get-OleColor 45 45 45) -TitleSize 12.4 -BodySize 10.1 -TitleHeight 24 -BodyTopOffset 40 -BodyHeight 36 | Out-Null
    Add-GeneratedCard -Slide $Slide -Name "Gen_StartStep2" -Left 266 -Top 226 -Width 188 -Height 92 -FillColor (Get-OleColor 255 243 232) -Title "2. Проверить данные" -Body "Известны масса, схема строповки и нужная оснастка." -TitleColor (Get-OleColor 163 79 20) -BodyColor (Get-OleColor 60 60 60) -TitleSize 12.2 -BodySize 9.8 -TitleHeight 24 -BodyTopOffset 40 -BodyHeight 38 | Out-Null
    Add-GeneratedCard -Slide $Slide -Name "Gen_StartStep3" -Left 498 -Top 226 -Width 188 -Height 92 -FillColor (Get-OleColor 234 243 255) -Title "3. Оценить условия" -Body "Проверить зону работ, путь груза и наличие помех." -TitleColor (Get-OleColor 23 54 93) -BodyColor (Get-OleColor 45 45 45) -TitleSize 12.2 -BodySize 9.8 -TitleHeight 24 -BodyTopOffset 40 -BodyHeight 38 | Out-Null
    Add-GeneratedCard -Slide $Slide -Name "Gen_StartStep4" -Left 150 -Top 332 -Width 188 -Height 92 -FillColor (Get-OleColor 255 243 232) -Title "4. Согласовать команды" -Body "Определить, кто подает сигналы и как подтверждаются команды." -TitleColor (Get-OleColor 163 79 20) -BodyColor (Get-OleColor 60 60 60) -TitleSize 11.8 -BodySize 9.6 -TitleHeight 24 -BodyTopOffset 40 -BodyHeight 40 | Out-Null
    Add-GeneratedCard -Slide $Slide -Name "Gen_StartStep5" -Left 382 -Top 332 -Width 188 -Height 92 -FillColor (Get-OleColor 234 243 255) -Title "5. Принять решение" -Body "Если хоть один пункт не подтвержден, подъем не начинают." -TitleColor (Get-OleColor 23 54 93) -BodyColor (Get-OleColor 45 45 45) -TitleSize 12 -BodySize 9.7 -TitleHeight 24 -BodyTopOffset 40 -BodyHeight 38 | Out-Null

    $noteBox = $Slide.Shapes.AddShape(5, 74, 436, 610, 40)
    $noteBox.Name = "Gen_StartNoteBox"
    $noteBox.Fill.ForeColor.RGB = (Get-OleColor 234 243 255)
    $noteBox.Line.Visible = 0
    $noteBox.Adjustments.Item(1) = 0.15
    Add-GeneratedTextBox -Slide $Slide -Name "Gen_StartNote" -Left 94 -Top 445 -Width 570 -Height 20 -Text "Этот алгоритм помогает остановить ошибку еще до подъема груза." -Size 11 -Color (Get-OleColor 23 54 93) -Bold $true -Alignment 2 | Out-Null
}

function Configure-StartMistakesSlide {
    param($Slide)

    Set-HeaderAndLead -Slide $Slide -Header "Подготовка к подъему" -Lead "Типовые ошибки в начале работ"

    Add-GeneratedCard -Slide $Slide -Name "Gen_MistakesIntro" -Left 42 -Top 116 -Width 412 -Height 78 -FillColor (Get-OleColor 255 243 232) -Title "Большинство инцидентов начинается до подъема груза" -Body "Ошибки появляются на этапе подготовки, когда исходные данные не проверили и начали работать на доверии." -TitleColor (Get-OleColor 163 79 20) -BodyColor (Get-OleColor 60 60 60) -TitleSize 14.2 -BodySize 10 -TitleHeight 28 -BodyTopOffset 46 -BodyHeight 22 | Out-Null

    $imageFrame = $Slide.Shapes.AddShape(5, 478, 100, 194, 320)
    $imageFrame.Name = "Gen_MistakesFrame"
    $imageFrame.Fill.ForeColor.RGB = (Get-OleColor 250 250 250)
    $imageFrame.Line.ForeColor.RGB = (Get-OleColor 225 225 225)
    $imageFrame.Adjustments.Item(1) = 0.08
    $prohibitedPlacement = Get-FitPlacement -ImagePath $slide19UserLayoutImage -Left 486 -Top 110 -MaxWidth 178 -MaxHeight 302
    $prohibitedPicture = $Slide.Shapes.AddPicture($slide19UserLayoutImage, $false, $true, $prohibitedPlacement.Left, $prohibitedPlacement.Top, $prohibitedPlacement.Width, $prohibitedPlacement.Height)
    $prohibitedPicture.Name = "Gen_MistakesImage"

    Add-GeneratedCard -Slide $Slide -Name "Gen_Mistake1" -Left 42 -Top 208 -Width 412 -Height 44 -FillColor (Get-OleColor 234 243 255) -Title "Команды не согласовали" -Body "Крановщик и стропальщик по-разному понимают начало подъема." -TitleColor (Get-OleColor 23 54 93) -BodyColor (Get-OleColor 45 45 45) -TitleSize 13.5 -BodySize 10 -TitleHeight 18 -BodyTopOffset 24 -BodyHeight 16 | Out-Null
    Add-GeneratedCard -Slide $Slide -Name "Gen_Mistake2" -Left 42 -Top 260 -Width 412 -Height 44 -FillColor (Get-OleColor 255 243 232) -Title "Масса груза не известна" -Body "Оснастку выбирают наугад, не понимая реальной нагрузки." -TitleColor (Get-OleColor 163 79 20) -BodyColor (Get-OleColor 60 60 60) -TitleSize 13.5 -BodySize 10 -TitleHeight 18 -BodyTopOffset 24 -BodyHeight 16 | Out-Null
    Add-GeneratedCard -Slide $Slide -Name "Gen_Mistake3" -Left 42 -Top 312 -Width 412 -Height 44 -FillColor (Get-OleColor 234 243 255) -Title "Стропы не проверили" -Body "Повреждения и отсутствие бирки замечают слишком поздно." -TitleColor (Get-OleColor 23 54 93) -BodyColor (Get-OleColor 45 45 45) -TitleSize 13.5 -BodySize 10 -TitleHeight 18 -BodyTopOffset 24 -BodyHeight 16 | Out-Null
    Add-GeneratedCard -Slide $Slide -Name "Gen_Mistake4" -Left 42 -Top 364 -Width 412 -Height 44 -FillColor (Get-OleColor 255 243 232) -Title "Схема строповки не ясна" -Body "Груз цепляют без понимания центра тяжести и точек захвата." -TitleColor (Get-OleColor 163 79 20) -BodyColor (Get-OleColor 60 60 60) -TitleSize 13.2 -BodySize 9.8 -TitleHeight 18 -BodyTopOffset 24 -BodyHeight 16 | Out-Null
    Add-GeneratedCard -Slide $Slide -Name "Gen_Mistake5" -Left 42 -Top 416 -Width 412 -Height 44 -FillColor (Get-OleColor 234 243 255) -Title "Неподготовленный персонал" -Body "Человек не допущен к работе или не понимает порядок действий." -TitleColor (Get-OleColor 23 54 93) -BodyColor (Get-OleColor 45 45 45) -TitleSize 12.8 -BodySize 9.7 -TitleHeight 18 -BodyTopOffset 24 -BodyHeight 16 | Out-Null

    Add-GeneratedTextBox -Slide $Slide -Name "Gen_MistakesRight1" -Left 572 -Top 136 -Width 82 -Height 42 -Text "Применять для обвязки груза случайные средства." -Size 8.5 -Color (Get-OleColor 60 60 60) | Out-Null
    Add-GeneratedTextBox -Slide $Slide -Name "Gen_MistakesRight2" -Left 572 -Top 192 -Width 82 -Height 40 -Text "Забивать крюки стропов и монтажные петли." -Size 8.5 -Color (Get-OleColor 60 60 60) | Out-Null
    Add-GeneratedTextBox -Slide $Slide -Name "Gen_MistakesRight3" -Left 572 -Top 248 -Width 82 -Height 46 -Text "Освобождать краном защемленные стропы, цепи и канаты." -Size 8.3 -Color (Get-OleColor 60 60 60) | Out-Null
    Add-GeneratedTextBox -Slide $Slide -Name "Gen_MistakesRight4" -Left 572 -Top 308 -Width 82 -Height 40 -Text "Поднимать тару, заполненную выше бортов." -Size 8.5 -Color (Get-OleColor 60 60 60) | Out-Null
    Add-GeneratedTextBox -Slide $Slide -Name "Gen_MistakesRight5" -Left 572 -Top 364 -Width 82 -Height 60 -Text "Поднимать или опускать груз на автомобиль, если в кабине или кузове есть люди." -Size 8.1 -Color (Get-OleColor 60 60 60) | Out-Null
}

function Configure-InteractionProcessSlide {
    param($Slide)

    Set-HeaderAndLead -Slide $Slide -Header "Взаимодействие крановщика и стропальщика" -Lead "Взаимодействие крановщика и стропальщика в процессе работ"

    Add-GeneratedCard -Slide $Slide -Name "Gen_InteractionTop" -Left 118 -Top 130 -Width 484 -Height 62 -FillColor (Get-OleColor 20 83 144) -Title "Безопасный подъем - это непрерывный обмен понятными действиями" -Body "" -TitleColor (Get-OleColor 255 255 255) -BodyColor (Get-OleColor 255 255 255) -TitleSize 13.8 -TitleHeight 28 -BodyTopOffset 36 -BodyHeight 0 | Out-Null

    $rwPlacement = Get-FitPlacement -ImagePath $reportWarnImage -Left 548 -Top 116 -MaxWidth 120 -MaxHeight 102
    $rwPicture = $Slide.Shapes.AddPicture($reportWarnImage, $false, $true, $rwPlacement.Left, $rwPlacement.Top, $rwPlacement.Width, $rwPlacement.Height)
    $rwPicture.Name = "Gen_InteractionImage"

    Add-GeneratedCard -Slide $Slide -Name "Gen_InteractionSlinger" -Left 42 -Top 244 -Width 244 -Height 122 -FillColor (Get-OleColor 234 243 255) -Title "Стропальщик" -Body "Проверяет груз и оснастку.`r`nПодает согласованную команду.`r`nСледит за опасной зоной и движением груза." -TitleColor (Get-OleColor 23 54 93) -BodyColor (Get-OleColor 45 45 45) -TitleSize 15 -BodySize 10.4 -TitleHeight 26 -BodyTopOffset 44 -BodyHeight 66 | Out-Null
    Add-GeneratedCard -Slide $Slide -Name "Gen_InteractionOperator" -Left 438 -Top 244 -Width 244 -Height 122 -FillColor (Get-OleColor 255 243 232) -Title "Крановщик" -Body "Принимает только понятную команду.`r`nВыполняет движение в пределах возможностей крана.`r`nОстанавливает подъем при риске." -TitleColor (Get-OleColor 163 79 20) -BodyColor (Get-OleColor 60 60 60) -TitleSize 15 -BodySize 10.1 -TitleHeight 26 -BodyTopOffset 44 -BodyHeight 68 | Out-Null

    Add-GeneratedArrow -Slide $Slide -Name "Gen_InteractionArrow1" -Left 292 -Top 282 -Width 74 -Height 24 -Color (Get-OleColor 245 135 31) | Out-Null
    Add-GeneratedTextBox -Slide $Slide -Name "Gen_InteractionFeedback" -Left 298 -Top 308 -Width 64 -Height 16 -Text "обратная связь" -Size 8.8 -Color (Get-OleColor 23 54 93) -Bold $true -Alignment 2 | Out-Null

    Add-GeneratedCard -Slide $Slide -Name "Gen_InteractionMid" -Left 214 -Top 382 -Width 296 -Height 72 -FillColor (Get-OleColor 234 243 255) -Title "Рабочий цикл" -Body "Команда -> подтверждение -> движение -> контроль -> остановка при риске." -TitleColor (Get-OleColor 23 54 93) -BodyColor (Get-OleColor 45 45 45) -TitleSize 14 -BodySize 10.4 -TitleHeight 22 -BodyTopOffset 38 -BodyHeight 24 | Out-Null

    $noteBox = $Slide.Shapes.AddShape(5, 78, 462, 596, 30)
    $noteBox.Name = "Gen_InteractionNoteBox"
    $noteBox.Fill.ForeColor.RGB = (Get-OleColor 255 243 232)
    $noteBox.Line.Visible = 0
    $noteBox.Adjustments.Item(1) = 0.15
    Add-GeneratedTextBox -Slide $Slide -Name "Gen_InteractionNote" -Left 96 -Top 468 -Width 560 -Height 18 -Text "Ошибка одного участника сразу становится общей угрозой для всей операции." -Size 10.8 -Color (Get-OleColor 163 79 20) -Bold $true -Alignment 2 | Out-Null
}

function Configure-KeySignalsSlide {
    param($Slide)

    Set-HeaderAndLead -Slide $Slide -Header "Знаковая сигнализация" -Lead "Основные знаковые сигналы стропальщика"

    Add-GeneratedCard -Slide $Slide -Name "Gen_KeySignalsTop" -Left 170 -Top 126 -Width 380 -Height 54 -FillColor (Get-OleColor 20 83 144) -Title "Базовые сигналы без повтора полной таблицы" -Body "" -TitleColor (Get-OleColor 255 255 255) -BodyColor (Get-OleColor 255 255 255) -TitleSize 13 -TitleHeight 24 -BodyTopOffset 32 -BodyHeight 0 | Out-Null

    Add-PhotoInfoCard -Slide $Slide -Name "Gen_KeyStop" -Left 18 -Top 210 -Width 166 -Height 186 -FillColor (Get-OleColor 234 243 255) -Title "Стоп" -Body "Немедленная остановка.`r`nИсполняется сразу." -TitleColor (Get-OleColor 23 54 93) -BodyColor (Get-OleColor 45 45 45) -ImagePath $stopPoseImage -TextWidth 136 -TitleTopOffset 12 -TitleHeight 20 -BodyTopOffset 38 -TitleSize 14.5 -BodySize 10.2 -ImageBoxLeftOffset 20 -ImageBoxTopOffset 94 -ImageBoxWidth 126 -ImageBoxHeight 78 | Out-Null
    Add-PhotoInfoCard -Slide $Slide -Name "Gen_KeyRaise" -Left 192 -Top 210 -Width 166 -Height 186 -FillColor (Get-OleColor 255 243 232) -Title "Поднять груз" -Body "Команда на начало подъема.`r`nПодается четко." -TitleColor (Get-OleColor 163 79 20) -BodyColor (Get-OleColor 60 60 60) -ImagePath $raiseLoadImage -TextWidth 136 -TitleTopOffset 12 -TitleHeight 20 -BodyTopOffset 38 -TitleSize 14.5 -BodySize 10.2 -ImageBoxLeftOffset 20 -ImageBoxTopOffset 94 -ImageBoxWidth 126 -ImageBoxHeight 78 | Out-Null
    Add-PhotoInfoCard -Slide $Slide -Name "Gen_KeyMove" -Left 366 -Top 210 -Width 166 -Height 186 -FillColor (Get-OleColor 234 243 255) -Title "Перемещение стрелы" -Body "Показывает направление требуемого движения." -TitleColor (Get-OleColor 23 54 93) -BodyColor (Get-OleColor 45 45 45) -ImagePath $signalImage -TextWidth 136 -TitleTopOffset 12 -TitleHeight 20 -BodyTopOffset 38 -TitleSize 12.6 -BodySize 9.9 -ImageBoxLeftOffset 20 -ImageBoxTopOffset 94 -ImageBoxWidth 126 -ImageBoxHeight 78 | Out-Null
    Add-PhotoInfoCard -Slide $Slide -Name "Gen_KeyWarn" -Left 540 -Top 210 -Width 166 -Height 186 -FillColor (Get-OleColor 255 243 232) -Title "Осторожно" -Body "Перед точным и медленным перемещением." -TitleColor (Get-OleColor 163 79 20) -BodyColor (Get-OleColor 60 60 60) -ImagePath $warningSignalImage -TextWidth 136 -TitleTopOffset 12 -TitleHeight 20 -BodyTopOffset 38 -TitleSize 14.5 -BodySize 10 -ImageBoxLeftOffset 20 -ImageBoxTopOffset 94 -ImageBoxWidth 126 -ImageBoxHeight 78 | Out-Null

    Add-GeneratedCard -Slide $Slide -Name "Gen_KeyRules" -Left 92 -Top 418 -Width 552 -Height 58 -FillColor (Get-OleColor 234 243 255) -Title "Общие требования" -Body "Сигнал заранее согласуют, подают ясно и без двусмысленности. Команду дает один назначенный стропальщик." -TitleColor (Get-OleColor 23 54 93) -BodyColor (Get-OleColor 45 45 45) -TitleSize 13.5 -BodySize 10.1 -TitleHeight 20 -BodyTopOffset 30 -BodyHeight 20 | Out-Null
}

function Configure-TitleSlide {
    param(
        $Slide,
        [string]$LogoPath
    )

    Remove-GeneratedShapes -Slide $Slide
    $title = $Slide.Shapes.Item("Title 8")
    Set-TextBoxText -Shape $title -Text "УЧЕБНАЯ ПРОГРАММА`r`nПРОФЕССИЯ: СТРОПАЛЬЩИК`r`nБезопасное выполнение работ" -Size 23 -Color (Get-OleColor 23 54 93) -Bold $true -Alignment 2
    $title.Left = 35
    $title.Top = 136
    $title.Width = 650
    $title.Height = 110

    Remove-ShapeIfExists -Slide $Slide -Name "Gen_TitleSubtitle"
    Add-GeneratedTextBox -Slide $Slide -Name "Gen_TitleSubtitle" -Left 120 -Top 322 -Width 480 -Height 34 -Text "Велесстрой | учебный курс для подготовки и проверки знаний" -Size 17 -Color (Get-OleColor 70 70 70) -Bold $true -Alignment 2 | Out-Null
}

function Add-TemplateSlideAt {
    param(
        $Presentation,
        [int]$TemplateSlideId,
        [int]$TargetIndex
    )

    $template = $Presentation.Slides.FindBySlideID($TemplateSlideId)
    $duplicate = $template.Duplicate()
    $slide = $duplicate.Item(1)
    $slide.MoveTo($TargetIndex)
    $slide
}

$resolvedInputPath = Resolve-ProjectPath $InputPath
$resolvedOutputPath = Resolve-ProjectPath $OutputPath
Ensure-Directory -Path (Split-Path -Parent $resolvedOutputPath)
Copy-Item -LiteralPath $resolvedInputPath -Destination $resolvedOutputPath -Force

$dangerZoneImage = Resolve-ProjectPath "assets\working-visuals\60-68\slide-62-danger-zone.png"
$mistakesImage = Resolve-ProjectPath "assets\working-visuals\60-68\slide-66-prohibited-lifts.png"
$autocraneImage = Resolve-ProjectPath "assets\working-visuals\60-68\slide-62-danger-zone-photo.jfif"
$bridgeCraneImage = Resolve-ProjectPath "assets\working-visuals\01-10\slide-06-bridge-crane.png"
$craneManipulatorImage = Resolve-ProjectPath "assets\working-visuals\01-10\slide-06-crane-manipulator.png"
$pipeLayerImage = Resolve-ProjectPath "assets\working-visuals\01-10\slide-06-pipe-layer.png"
$roleImage = Resolve-ProjectPath "assets\исходные-фото\Осторожно. Работник велесстрой кадр 1..png"
$signalImage = Resolve-ProjectPath "assets\исходные-фото\переместить стрелу..png"
$raiseLoadImage = Resolve-ProjectPath "assets\исходные-фото\поднять груз.png"
$stopPoseImage = Resolve-ProjectPath "assets\исходные-фото\стоп.png"
$stopSignalImage = Resolve-ProjectPath "assets\working-visuals\69-76\slide-70-stop-signal.png"
$warningSignalImage = Resolve-ProjectPath "assets\working-visuals\69-76\slide-70-warning-signal.png"
$signalCheatsheetImage = Resolve-ProjectPath "assets\working-visuals\69-76\slide-74-signal-cheatsheet.png"
$reportWarnImage = Resolve-ProjectPath "assets\working-visuals\69-76\slide-71-report-and-warn.png"
$prohibitedActionsImage = Resolve-ProjectPath "assets\working-visuals\69-76\slide-72-prohibited-actions.png"
$prohibitedActionsCleanImage = Resolve-ProjectPath "assets\working-visuals\18-21\slide-19-prohibited-actions-clean.png"
$slide19UserLayoutImage = Resolve-ProjectPath "assets\working-visuals\18-21\slide-19-user-layout.png"
$algorithmBackgroundImage = Resolve-ProjectPath "assets\working-visuals\69-76\slide-75-algorithm-background.png"
$logoPath = Resolve-ProjectPath "assets\бренд-материалы\логотип велесстрой.png"

$powerPoint = $null
$presentation = $null

try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.Visible = -1
    $presentation = $powerPoint.Presentations.Open($resolvedOutputPath, $false, $false, $false)

    $templateSlideId = $presentation.Slides.Item(74).SlideID

    Configure-TitleSlide -Slide $presentation.Slides.Item(1) -LogoPath $logoPath

    $specs = @(
        @{ Number = 2; Config = "Agenda" }
        @{ Number = 3; Config = "Norms" }
        @{ Number = 4; Config = "Flow" }
        @{ Number = 5; Config = "Role" }
        @{ Number = 6; Config = "Equipment" }
        @{ Number = 7; Config = "Crane" }
        @{ Number = 8; Config = "Danger" }
        @{ Number = 9; Config = "Mistakes" }
        @{ Number = 10; Config = "Bridge" }
        @{ Number = 11; Config = "Admission" }
        @{ Number = 12; Config = "Prestart" }
        @{ Number = 13; Config = "Participants" }
        @{ Number = 14; Config = "CommandAuthority" }
        @{ Number = 15; Config = "CommandLanguage" }
        @{ Number = 16; Config = "CommunicationTypes" }
        @{ Number = 17; Config = "SignalCheatsheet" }
        @{ Number = 18; Config = "SafeStartAlgorithm" }
        @{ Number = 19; Config = "StartMistakes" }
    )

    foreach ($spec in $specs) {
        $slide = Add-TemplateSlideAt -Presentation $presentation -TemplateSlideId $templateSlideId -TargetIndex $spec.Number
        Set-SlideNumber -Slide $slide -Number $spec.Number

        switch ($spec.Config) {
            "Agenda" {
                Configure-AgendaSlide -Slide $slide
            }
            "Norms" {
                Configure-NormativeSlide -Slide $slide
            }
            "Flow" {
                Configure-LearningFlowSlide -Slide $slide
            }
            "Role" {
                Configure-RightImageSlide -Slide $slide -Header "Введение" -Lead "Роль стропальщика в производственном процессе" -Body "Стропальщик отвечает за подготовку груза к подъему.`r`n`r`n- выбирает подходящие грузозахватные приспособления;`r`n- проверяет схему строповки;`r`n- проводит строповку;`r`n- подает команды крановщику;`r`n- контролирует опасную зону;`r`n- сопровождает груз до безопасной установки или складирования." -ImagePath $roleImage -ImageLeft 410 -ImageTop 112 -ImageMaxWidth 280 -ImageMaxHeight 300
            }
            "Equipment" {
                Configure-EquipmentOverviewSlide -Slide $slide
            }
            "Crane" {
                Configure-CraneKnowledgeSlide -Slide $slide
            }
            "Danger" {
                Configure-DangerZonesSlide -Slide $slide -ImagePath $dangerZoneImage
            }
            "Mistakes" {
                Configure-RightImageSlide -Slide $slide -Header "Общие сведения о подъемных сооружениях" -Lead "Типовые ошибки в зоне работы крана" -Body "Самые опасные действия:`r`n`r`n- заходить под груз;`r`n- пытаться удержать груз руками;`r`n- стоять на линии перемещения;`r`n- приближаться без команды и без контроля зоны." -ImagePath $mistakesImage
            }
            "Bridge" {
                Configure-CoordinationSlide -Slide $slide -ImagePath $signalImage
            }
            "Admission" {
                Configure-AdmissionSlide -Slide $slide
            }
            "Prestart" {
                Configure-PrestartChecksSlide -Slide $slide
            }
            "Participants" {
                Configure-ParticipantsSlide -Slide $slide
            }
            "CommandAuthority" {
                Configure-CommandAuthoritySlide -Slide $slide
            }
            "CommandLanguage" {
                Configure-CommandLanguageSlide -Slide $slide
            }
            "CommunicationTypes" {
                Configure-CommunicationTypesSlide -Slide $slide
            }
            "SignalCheatsheet" {
                Configure-SignalCheatsheetSlide -Slide $slide
            }
            "SafeStartAlgorithm" {
                Configure-SafeStartAlgorithmSlide -Slide $slide
            }
            "StartMistakes" {
                Configure-StartMistakesSlide -Slide $slide
            }
        }
    }

    $firstOriginalIndex = 2 + $specs.Count
    $lastOriginalIndex = (2 * $specs.Count) + 1
    for ($index = $lastOriginalIndex; $index -ge $firstOriginalIndex; $index--) {
        $presentation.Slides.Item($index).Delete()
    }

    for ($i = 2; $i -le $presentation.Slides.Count; $i++) {
        try {
            Set-SlideNumber -Slide $presentation.Slides.Item($i) -Number $i
        } catch {
        }
    }

    $presentation.Save()

    $previewDir = Resolve-ProjectPath ".codex-temp\rebuilt-slides-1-10"
    Ensure-Directory -Path $previewDir
    Get-ChildItem -LiteralPath $previewDir -File -ErrorAction SilentlyContinue | Remove-Item -Force
    for ($i = 1; $i -le 19; $i++) {
        $presentation.Slides.Item($i).Export((Join-Path $previewDir ("slide-{0:D3}.png" -f $i)), "PNG", 1600, 900)
    }
}
finally {
    if ($presentation) {
        $presentation.Close()
    }
    if ($powerPoint) {
        $powerPoint.Quit()
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if (-not (Test-IsVersionedDraftPath -Path $resolvedOutputPath)) {
    try {
        $versionedDraftPath = Get-NextVersionedDraftPath -CurrentOutputPath $resolvedOutputPath
        Copy-Item -LiteralPath $resolvedOutputPath -Destination $versionedDraftPath -Force
        Write-Output "Версия: $versionedDraftPath"
    } catch {
        Write-Warning ("Не удалось создать версионную копию черновика: {0}" -f $_.Exception.Message)
    }
}

Write-Output "Готово: $resolvedOutputPath"
