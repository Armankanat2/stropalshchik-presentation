param(
    [string]$InputPath = "deliverables\черновики\КУРС для СТРОПАЛЬЩИКА_черновик_актуальный_0-22.pptx",
    [string]$OutputPath = ""
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing

function Resolve-ProjectPath {
    param([string]$RelativePath)

    if ([System.IO.Path]::IsPathRooted($RelativePath)) {
        return [System.IO.Path]::GetFullPath($RelativePath)
    }

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
    param([string]$CurrentPath)

    $directory = Split-Path -Parent $CurrentPath
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($CurrentPath)
    $extension = [System.IO.Path]::GetExtension($CurrentPath)
    $baseName = $fileName -replace '_\d+-\d+$', ''
    $pattern = '^' + [regex]::Escape($baseName) + '_(\d+)-(\d+)' + [regex]::Escape($extension) + '$'

    $maxMajor = 0
    $maxMinor = -1
    foreach ($file in Get-ChildItem -LiteralPath $directory -File -ErrorAction SilentlyContinue) {
        if ($file.Name -match $pattern) {
            $major = [int]$matches[1]
            $minor = [int]$matches[2]
            if (($major -gt $maxMajor) -or (($major -eq $maxMajor) -and ($minor -gt $maxMinor))) {
                $maxMajor = $major
                $maxMinor = $minor
            }
        }
    }

    $nextMinor = $maxMinor + 1
    Join-Path $directory ("{0}_{1}-{2}{3}" -f $baseName, $maxMajor, $nextMinor, $extension)
}

function Get-OleColor {
    param([int]$R, [int]$G, [int]$B)

    [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb($R, $G, $B))
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

function Set-ShapeText {
    param(
        $Shape,
        [string]$Text,
        [double]$FontSize,
        [int]$Color,
        [bool]$Bold = $false,
        [int]$Alignment = 1
    )

    $Shape.TextFrame.TextRange.Text = $Text
    $Shape.TextFrame.TextRange.Font.Name = "Verdana"
    $Shape.TextFrame.TextRange.Font.Size = [single]$FontSize
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
        [double]$FontSize,
        [int]$Color,
        [bool]$Bold = $false,
        [int]$Alignment = 1
    )

    Remove-ShapeIfExists -Slide $Slide -Name $Name
    $shape = $Slide.Shapes.AddTextbox(1, $Left, $Top, $Width, $Height)
    $shape.Name = $Name
    $shape.Fill.Visible = 0
    $shape.Line.Visible = 0
    Set-ShapeText -Shape $shape -Text $Text -FontSize $FontSize -Color $Color -Bold $Bold -Alignment $Alignment
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
        $ratio = [Math]::Min($MaxWidth / $image.Width, $MaxHeight / $image.Height)
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
        [double]$MaxHeight,
        [string]$Name = "Picture 2"
    )

    Remove-ShapeIfExists -Slide $Slide -Name $Name
    $placement = Get-FitPlacement -ImagePath $ImagePath -Left $Left -Top $Top -MaxWidth $MaxWidth -MaxHeight $MaxHeight
    $picture = $Slide.Shapes.AddPicture($ImagePath, $false, $true, $placement.Left, $placement.Top, $placement.Width, $placement.Height)
    $picture.Name = $Name
    $picture
}

function Clear-ToTemplate {
    param($Slide)

    $keep = @('TextBox 5', 'TextBox 10', 'Rectangle 2', 'TextBox 11', 'TextBox 9')
    $remove = @()
    foreach ($shape in @($Slide.Shapes)) {
        if ($keep -notcontains $shape.Name) {
            $remove += $shape.Name
        }
    }

    foreach ($name in $remove) {
        try {
            $Slide.Shapes.Item($name).Delete()
        } catch {
        }
    }
}

function Set-SlideNumber {
    param(
        $Slide,
        [int]$Number
    )

    try {
        $shape = $Slide.Shapes.Item("TextBox 5")
        Set-ShapeText -Shape $shape -Text ([string]$Number) -FontSize 16 -Color (Get-OleColor 255 255 255) -Bold $true -Alignment 2
    } catch {
    }
}

function Initialize-Slide {
    param(
        $Slide,
        [string]$Section,
        [string]$Lead,
        [double]$LeadFontSize = 19,
        [double]$LeadWidth = 640,
        [double]$LeadHeight = 54
    )

    Clear-ToTemplate -Slide $Slide

    $headerShape = $Slide.Shapes.Item("TextBox 10")
    $leadShape = $Slide.Shapes.Item("TextBox 11")
    $bodyShape = $Slide.Shapes.Item("TextBox 9")

    $headerShape.Left = 0
    $headerShape.Top = 20.5
    $headerShape.Width = 513.1
    $headerShape.Height = 26.7
    Set-ShapeText -Shape $headerShape -Text $Section -FontSize 18.5 -Color (Get-OleColor 255 255 255) -Bold $true

    $leadShape.Left = 24
    $leadShape.Top = 72
    $leadShape.Width = $LeadWidth
    $leadShape.Height = $LeadHeight
    Set-ShapeText -Shape $leadShape -Text $Lead -FontSize $LeadFontSize -Color (Get-OleColor 23 54 93) -Bold $true

    $bodyShape.Left = 0
    $bodyShape.Top = 0
    $bodyShape.Width = 1
    $bodyShape.Height = 1
    Set-ShapeText -Shape $bodyShape -Text "" -FontSize 8 -Color (Get-OleColor 255 255 255)
}

function Configure-TextSlide {
    param(
        $Slide,
        [string]$Section,
        [string]$Lead,
        [string]$Body
    )

    Initialize-Slide -Slide $Slide -Section $Section -Lead $Lead
    $bodyShape = $Slide.Shapes.Item("TextBox 9")
    $bodyShape.Left = 24
    $bodyShape.Top = 150
    $bodyShape.Width = 660
    $bodyShape.Height = 290
    Set-ShapeText -Shape $bodyShape -Text $Body -FontSize 17 -Color (Get-OleColor 60 60 60)
}

function Configure-RightImageSlide {
    param(
        $Slide,
        [string]$Section,
        [string]$Lead,
        [string]$Body,
        [string]$ImagePath
    )

    Initialize-Slide -Slide $Slide -Section $Section -Lead $Lead
    $bodyShape = $Slide.Shapes.Item("TextBox 9")
    $bodyShape.Left = 24
    $bodyShape.Top = 150
    $bodyShape.Width = 350
    $bodyShape.Height = 280
    Set-ShapeText -Shape $bodyShape -Text $Body -FontSize 16.5 -Color (Get-OleColor 60 60 60)

    if (-not [string]::IsNullOrWhiteSpace($ImagePath) -and (Test-Path -LiteralPath $ImagePath)) {
        Replace-ContentPicture -Slide $Slide -ImagePath $ImagePath -Left 400 -Top 145 -MaxWidth 290 -MaxHeight 300 | Out-Null
    }
}

function Configure-FullImageSlide {
    param(
        $Slide,
        [string]$Section,
        [string]$Lead,
        [string]$Caption,
        [string]$ImagePath
    )

    Initialize-Slide -Slide $Slide -Section $Section -Lead $Lead -LeadFontSize 18.5 -LeadWidth 660 -LeadHeight 50
    if (-not [string]::IsNullOrWhiteSpace($ImagePath) -and (Test-Path -LiteralPath $ImagePath)) {
        Replace-ContentPicture -Slide $Slide -ImagePath $ImagePath -Left 55 -Top 135 -MaxWidth 610 -MaxHeight 270 | Out-Null
    }

    $bodyShape = $Slide.Shapes.Item("TextBox 9")
    $bodyShape.Left = 36
    $bodyShape.Top = 420
    $bodyShape.Width = 648
    $bodyShape.Height = 34
    Set-ShapeText -Shape $bodyShape -Text $Caption -FontSize 13.8 -Color (Get-OleColor 163 79 20) -Bold $true -Alignment 2
}

function Configure-TransitionSlide {
    param(
        $Slide,
        [string]$Section,
        [string]$Lead,
        [string]$Caption,
        [string]$ImagePath
    )

    Initialize-Slide -Slide $Slide -Section $Section -Lead $Lead -LeadFontSize 21 -LeadWidth 650 -LeadHeight 48
    if (-not [string]::IsNullOrWhiteSpace($ImagePath) -and (Test-Path -LiteralPath $ImagePath)) {
        Replace-ContentPicture -Slide $Slide -ImagePath $ImagePath -Left 22 -Top 132 -MaxWidth 676 -MaxHeight 278 | Out-Null
    }

    $bodyShape = $Slide.Shapes.Item("TextBox 9")
    $bodyShape.Left = 48
    $bodyShape.Top = 422
    $bodyShape.Width = 624
    $bodyShape.Height = 28
    Set-ShapeText -Shape $bodyShape -Text $Caption -FontSize 14 -Color (Get-OleColor 23 54 93) -Bold $true -Alignment 2
}

function Configure-ImageQuestionSlide {
    param(
        $Slide,
        [string]$Section,
        [string]$Lead,
        [string]$Question,
        [string]$ImagePath
    )

    Initialize-Slide -Slide $Slide -Section $Section -Lead $Lead -LeadFontSize 18.5 -LeadWidth 660 -LeadHeight 50

    $bodyShape = $Slide.Shapes.Item("TextBox 9")
    $bodyShape.Left = 26
    $bodyShape.Top = 126
    $bodyShape.Width = 664
    $bodyShape.Height = 40
    Set-ShapeText -Shape $bodyShape -Text $Question -FontSize 16.5 -Color (Get-OleColor 60 60 60) -Bold $true -Alignment 2

    if (-not [string]::IsNullOrWhiteSpace($ImagePath) -and (Test-Path -LiteralPath $ImagePath)) {
        Replace-ContentPicture -Slide $Slide -ImagePath $ImagePath -Left 60 -Top 175 -MaxWidth 600 -MaxHeight 255 | Out-Null
    }
}

function Configure-FinalSlide {
    param(
        $Slide,
        [string]$Section,
        [string]$Lead,
        [string]$ImagePath
    )

    Initialize-Slide -Slide $Slide -Section $Section -Lead $Lead -LeadFontSize 19 -LeadWidth 640 -LeadHeight 44
    if (-not [string]::IsNullOrWhiteSpace($ImagePath) -and (Test-Path -LiteralPath $ImagePath)) {
        Replace-ContentPicture -Slide $Slide -ImagePath $ImagePath -Left 0 -Top 58 -MaxWidth 720 -MaxHeight 432 | Out-Null
        try { $Slide.Shapes.Item("Picture 2").ZOrder(1) | Out-Null } catch {}
        try { $Slide.Shapes.Item("TextBox 10").ZOrder(0) | Out-Null } catch {}
        try { $Slide.Shapes.Item("TextBox 11").ZOrder(0) | Out-Null } catch {}
        try { $Slide.Shapes.Item("TextBox 5").ZOrder(0) | Out-Null } catch {}
    }

    $bodyShape = $Slide.Shapes.Item("TextBox 9")
    $bodyShape.Left = 40
    $bodyShape.Top = 414
    $bodyShape.Width = 420
    $bodyShape.Height = 28
    Set-ShapeText -Shape $bodyShape -Text "Курс завершён. Безопасная работа начинается с понятного алгоритма и дисциплины на площадке." -FontSize 13.5 -Color (Get-OleColor 255 255 255) -Bold $true
}

function Configure-Slide23 {
    param(
        $Slide,
        [string]$ImagePath
    )

    Configure-RightImageSlide -Slide $Slide `
        -Section "Стропы и грузозахватные приспособления" `
        -Lead "Иконки и знаки на бирках стропов" `
        -Body "На бирке ищут схему строповки, ограничения применения и рабочую нагрузку.`r`n`r`nТаблица на бирке показывает, как меняется допустимая нагрузка при разных способах строповки.`r`n`r`nЕсли маркировка непонятна или не читается, строп до уточнения данных не используют." `
        -ImagePath $ImagePath

    Add-GeneratedTextBox -Slide $Slide -Name "Gen23_Footer" -Left 40 -Top 430 -Width 640 -Height 20 -Text "Бирка - это рабочая инструкция производителя, а не формальный ярлык." -FontSize 13.2 -Color (Get-OleColor 163 79 20) -Bold $true -Alignment 2 | Out-Null
}

function Replace-SlideWithTemplate {
    param(
        $Presentation,
        [int]$Index,
        [int]$TemplateIndex
    )

    $template = $Presentation.Slides.Item($TemplateIndex)
    $duplicateRange = $template.Duplicate()
    $duplicate = $duplicateRange.Item(1)
    $duplicate.MoveTo($Index)
    $Presentation.Slides.Item($Index + 1).Delete()
    $Presentation.Slides.Item($Index)
}

$resolvedInputPath = Resolve-ProjectPath $InputPath
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Get-NextVersionedDraftPath -CurrentPath $resolvedInputPath
}
$resolvedOutputPath = Resolve-ProjectPath $OutputPath
Ensure-Directory -Path (Split-Path -Parent $resolvedOutputPath)

Copy-Item -LiteralPath $resolvedInputPath -Destination $resolvedOutputPath -Force

$root = Split-Path -Parent $PSScriptRoot
$specs = @(
    @{ Index = 24; Kind = "RightImage"; Section = "Стропы и грузозахватные приспособления"; Lead = "Браковка стропа и типовые дефекты"; Body = "Перед работой строп осматривают целиком.`r`n- обрывы проволок и прядей`r`n- деформация, надрывы и трещины`r`n- отсутствие бирки или нечитаемая маркировка`r`n`r`nСомнительный строп сразу выводят из работы."; Image = Resolve-ProjectPath "assets\working-visuals\22-36\slide-26-sling-defect.jpg" }
    @{ Index = 25; Kind = "RightImage"; Section = "Стропы и грузозахватные приспособления"; Lead = "Правильное использование стропов"; Body = "Строп выбирают по массе, форме груза и схеме подъема.`r`n- проверяют маркировку и состояние`r`n- защищают от острых кромок и перегибов`r`n- не допускают рывков, перекоса и скручивания`r`n`r`nПри сомнении работу останавливают и уточняют решение."; Image = Resolve-ProjectPath "assets\working-visuals\22-36\slide-27-correct-sling-use.jfif" }
    @{ Index = 26; Kind = "RightImage"; Section = "Стропы и грузозахватные приспособления"; Lead = "Основные грузозахватные приспособления и их маркировка"; Body = "Кроме стропов, в работе применяют крюки, скобы, захваты и другие элементы оснастки.`r`n`r`nПеред работой читают маркировку, грузоподъемность, номер, тип и производителя.`r`n`r`nЛюбое приспособление осматривают так же внимательно, как и сам строп."; Image = Resolve-ProjectPath "assets\working-visuals\22-36\slide-28-lifting-device-basic.jfif" }
    @{ Index = 27; Kind = "Text"; Section = "Стропы и грузозахватные приспособления"; Lead = "Браковка грузозахватных приспособлений и типовые дефекты"; Body = "- деформация корпуса, зева и осей`r`n- трещины, износ, отсутствие замков`r`n- повреждение сварных зон и посадочных мест`r`n- нечитаемая маркировка`r`n`r`nЕсли исправность вызывает сомнение, приспособление в работу не допускают." }
    @{ Index = 28; Kind = "RightImage"; Section = "Стропы и грузозахватные приспособления"; Lead = "Правильное использование грузозахватных приспособлений"; Body = "Приспособление должно соответствовать массе и конфигурации груза.`r`n`r`nПеред подъемом проверяют правильность зацепления, отсутствие перекоса и возможность безопасного освобождения после установки груза.`r`n`r`nНельзя использовать приспособление не по назначению."; Image = Resolve-ProjectPath "assets\working-visuals\22-36\slide-28-lifting-device-in-work.jfif" }
    @{ Index = 29; Kind = "RightImage"; Section = "Траверсы"; Lead = "Траверса: назначение и область применения"; Body = "Траверса помогает распределять нагрузку между точками подвеса и уменьшать усилия на ветвях стропа.`r`n`r`nЕё применяют для длинномерных, крупногабаритных и чувствительных к деформации грузов.`r`n`r`nВыбор траверсы зависит от схемы подъема и точки захвата."; Image = Resolve-ProjectPath "assets\working-visuals\22-36\slide-31-traverse-in-work.jfif" }
    @{ Index = 30; Kind = "Text"; Section = "Траверсы"; Lead = "Браковка траверсы и типовые дефекты"; Body = "- трещины и остаточные деформации`r`n- повреждение сварных соединений`r`n- износ проушин, пальцев и мест подвеса`r`n- отсутствие маркировки или таблички`r`n`r`nТраверсу с дефектом не испытывают в работе и не оставляют «до удобного случая»." }
    @{ Index = 31; Kind = "RightImage"; Section = "Траверсы"; Lead = "Правильное использование траверсы"; Body = "Перед работой проверяют схему строповки, точки подвеса и равномерность распределения нагрузки.`r`n`r`nПодъем с траверсой выполняют без рывков, перекоса и самовольной перестановки элементов подвеса.`r`n`r`nЕсли схема не ясна, работу приостанавливают."; Image = Resolve-ProjectPath "assets\working-visuals\22-36\slide-31-traverse-illustration.jfif" }
    @{ Index = 32; Kind = "RightImage"; Section = "Тара"; Lead = "Тара: назначение и область применения"; Body = "Тару используют для безопасного перемещения штучных, сыпучих и мелкоразмерных грузов.`r`n`r`nСтропальщик должен понимать назначение тары, её массу, допустимую загрузку и способ зацепления.`r`n`r`nИспользуют только исправную и промаркированную тару."; Image = Resolve-ProjectPath "assets\working-visuals\22-36\slide-34-container-photo-1.png" }
    @{ Index = 33; Kind = "RightImage"; Section = "Тара"; Lead = "Браковка тары и типовые дефекты"; Body = "- трещины, пробоины и остаточная деформация`r`n- поврежденные петли, проушины и элементы захвата`r`n- отсутствие маркировки и перегруз по массе`r`n`r`nТару с дефектами или неизвестной грузоподъемностью в работу не принимают."; Image = Resolve-ProjectPath "assets\working-visuals\22-36\slide-35-container-photo-2.png" }
    @{ Index = 34; Kind = "RightImage"; Section = "Тара"; Lead = "Правильное использование тары"; Body = "Перед подъемом убеждаются, что груз размещен устойчиво и не может выпасть при движении.`r`n`r`nТару загружают в пределах допустимой массы, а стропы присоединяют к штатным местам зацепления.`r`n`r`nПерегруженная или неправильно зацепленная тара опасна так же, как и неисправный строп."; Image = Resolve-ProjectPath "assets\working-visuals\22-36\slide-36-container-slinging.jfif" }
    @{ Index = 35; Kind = "Transition"; Section = "Алгоритм выполнения работ стропальщика"; Lead = "Общий алгоритм выполнения работ стропальщика"; Caption = "Задание -> оценка условий -> выбор оснастки -> строповка -> пробный подъем -> перемещение -> установка -> завершение."; Image = Resolve-ProjectPath "assets\working-visuals\37-59\slide-37-algorithm-background.png" }
    @{ Index = 36; Kind = "RightImage"; Section = "Алгоритм выполнения работ стропальщика"; Lead = "Получение задания и уточнение условий работы"; Body = "До начала работ стропальщик должен понимать задачу, массу груза, маршрут перемещения и особые ограничения.`r`n`r`nЕсли не хватает данных по схеме строповки или условиям площадки, работу не начинают до уточнения."; Image = Resolve-ProjectPath "assets\working-visuals\37-59\slide-38-task-briefing.png" }
    @{ Index = 37; Kind = "RightImage"; Section = "Алгоритм выполнения работ стропальщика"; Lead = "Когда стропальщик не приступает к работе"; Body = "К работе не приступают, если:`r`n- задача неясна`r`n- неизвестна масса груза`r`n- нет схемы или маркировки`r`n- оснастка вызывает сомнение`r`n- условия на площадке небезопасны."; Image = Resolve-ProjectPath "assets\working-visuals\37-59\slide-39-do-not-start.jfif" }
    @{ Index = 38; Kind = "RightImage"; Section = "Алгоритм выполнения работ стропальщика"; Lead = "Оценка условий, проверка СИЗ и готовности к работе"; Body = "Перед началом работ проверяют СИЗ, обзорность, состояние площадки, освещенность и наличие опасных факторов.`r`n`r`nБез готовности людей и места подъем даже исправным краном выполнять нельзя."; Image = Resolve-ProjectPath "assets\working-visuals\37-59\slide-40-ppe-and-readiness.png" }
    @{ Index = 39; Kind = "RightImage"; Section = "Алгоритм выполнения работ стропальщика"; Lead = "Выбор способа строповки и подбор оснастки"; Body = "Способ строповки выбирают по массе, форме, центру тяжести и точкам захвата груза.`r`n`r`nОснастка должна соответствовать задаче по грузоподъемности и длине ветвей."; Image = Resolve-ProjectPath "assets\working-visuals\37-59\slide-41-select-slinging-scheme.png" }
    @{ Index = 40; Kind = "RightImage"; Section = "Алгоритм выполнения работ стропальщика"; Lead = "Подготовка груза к подъему"; Body = "Перед подъемом освобождают груз от креплений, проверяют устойчивость, укладывают подкладки и убирают лишние предметы из зоны работы.`r`n`r`nГруз должен быть готов к безопасной строповке и перемещению."; Image = Resolve-ProjectPath "assets\working-visuals\37-59\slide-42-prepare-load-area.jfif" }
    @{ Index = 41; Kind = "RightImage"; Section = "Алгоритм выполнения работ стропальщика"; Lead = "Строповка груза"; Body = "Стропы устанавливают в штатные точки или по утвержденной схеме.`r`n`r`nНельзя допускать перекоса, скручивания, перегиба и случайного смещения центра тяжести."; Image = Resolve-ProjectPath "assets\working-visuals\37-59\slide-43-slinging-process.jfif" }
    @{ Index = 42; Kind = "RightImage"; Section = "Алгоритм выполнения работ стропальщика"; Lead = "Пробный подъем и контроль правильности строповки"; Body = "Пробный подъем выполняют на небольшую высоту.`r`n`r`nНа этом этапе проверяют устойчивость груза, работу тормозов, правильность схемы и отсутствие перекоса."; Image = Resolve-ProjectPath "assets\working-visuals\37-59\slide-44-test-lift.png" }
    @{ Index = 43; Kind = "RightImage"; Section = "Алгоритм выполнения работ стропальщика"; Lead = "Сопровождение и перемещение груза"; Body = "Во время перемещения груз сопровождают с безопасной позиции, не заходя под подвешенный груз и не поправляя его руками в опасной фазе.`r`n`r`nДвижение должно быть плавным и управляемым."; Image = Resolve-ProjectPath "assets\working-visuals\37-59\slide-45-guide-load.jfif" }
    @{ Index = 44; Kind = "RightImage"; Section = "Алгоритм выполнения работ стропальщика"; Lead = "Установка груза на место"; Body = "Перед опусканием проверяют место установки, подкладки и наличие свободного пространства.`r`n`r`nГруз ставят устойчиво, чтобы после освобождения стропов он не сместился и не опрокинулся."; Image = Resolve-ProjectPath "assets\working-visuals\37-59\slide-46-place-load.jfif" }
    @{ Index = 45; Kind = "RightImage"; Section = "Алгоритм выполнения работ стропальщика"; Lead = "Снятие стропов и завершение операции"; Body = "Стропы снимают только после полной устойчивой установки груза.`r`n`r`nПосле операции осматривают оснастку и рабочую зону, убирают оборудование и сообщают о замеченных дефектах."; Image = Resolve-ProjectPath "assets\working-visuals\37-59\slide-47-remove-slings.jfif" }
    @{ Index = 46; Kind = "RightImage"; Section = "Алгоритм выполнения работ стропальщика"; Lead = "Типовые ошибки при выполнении работ стропальщика"; Body = "- поспешный старт без уточнений`r`n- неправильный выбор оснастки`r`n- отсутствие пробного подъема`r`n- попытка поправить груз в опасной фазе`r`n- игнорирование замеченных нарушений"; Image = Resolve-ProjectPath "assets\working-visuals\37-59\slide-48-typical-errors.jfif" }
    @{ Index = 47; Kind = "RightImage"; Section = "Основные принципы строповки"; Lead = "Что такое правильная строповка и её основные принципы"; Body = "Правильная строповка удерживает груз устойчиво, не повреждает его и не перегружает ветви стропа.`r`n`r`nГлавные принципы: устойчивость, сохранение центра тяжести, защита от кромок и понятная схема подъема."; Image = Resolve-ProjectPath "assets\working-visuals\37-59\slide-49-correct-slinging-principles.png" }
    @{ Index = 48; Kind = "RightImage"; Section = "Основные принципы строповки"; Lead = "Что нужно проверить перед строповкой груза"; Body = "Проверяют массу, форму, центр тяжести, точки зацепления, наличие острых кромок и маршрут перемещения.`r`n`r`nБез этих данных подобрать безопасную схему строповки нельзя."; Image = Resolve-ProjectPath "assets\working-visuals\37-59\slide-50-check-before-slinging.png" }
    @{ Index = 49; Kind = "RightImage"; Section = "Основные принципы строповки"; Lead = "Типовые ошибки при строповке"; Body = "- зацепление за случайные элементы`r`n- перекос ветвей`r`n- угол, увеличивающий нагрузку на строп`r`n- отсутствие защиты от острых кромок`r`n- игнорирование центра тяжести"; Image = Resolve-ProjectPath "assets\working-visuals\37-59\slide-51-slinging-errors.jfif" }
    @{ Index = 50; Kind = "FullImage"; Section = "Схемы строповки по типам грузов"; Lead = "Схемы строповки: как читать этот раздел"; Caption = "Схема показывает точки зацепления, положение центра тяжести, число ветвей и ограничения по способу подъема."; Image = Resolve-ProjectPath "assets\working-visuals\37-59\slide-52-how-to-read-schemes.png" }
    @{ Index = 51; Kind = "FullImage"; Section = "Схемы строповки по типам грузов"; Lead = "Строповка труб"; Caption = "Для труб важны устойчивость связки, защита от скатывания и правильное распределение точек подвеса."; Image = Resolve-ProjectPath "assets\working-visuals\37-59\slide-53-pipes.png" }
    @{ Index = 52; Kind = "FullImage"; Section = "Схемы строповки по типам грузов"; Lead = "Строповка металла"; Caption = "Металл поднимают по схеме, которая исключает соскальзывание, повреждение кромками и потерю устойчивости."; Image = Resolve-ProjectPath "assets\working-visuals\37-59\slide-54-metal.png" }
    @{ Index = 53; Kind = "FullImage"; Section = "Схемы строповки по типам грузов"; Lead = "Строповка железобетонных изделий"; Caption = "Железобетон поднимают только за расчетные точки захвата и без нагрузки на поврежденные петли."; Image = Resolve-ProjectPath "assets\working-visuals\37-59\slide-55-concrete-elements.png" }
    @{ Index = 54; Kind = "FullImage"; Section = "Схемы строповки по типам грузов"; Lead = "Строповка контейнеров и тары с грузом"; Caption = "Контейнеры и тару перемещают только за штатные элементы, без перегруза и перекоса тары."; Image = Resolve-ProjectPath "assets\working-visuals\37-59\slide-56-containers-scheme.jfif" }
    @{ Index = 55; Kind = "FullImage"; Section = "Схемы строповки по типам грузов"; Lead = "Строповка оборудования"; Caption = "Оборудование требует учета центра тяжести, выступающих частей и мест, чувствительных к повреждению."; Image = Resolve-ProjectPath "assets\working-visuals\37-59\slide-57-equipment.png" }
    @{ Index = 56; Kind = "FullImage"; Section = "Схемы строповки по типам грузов"; Lead = "Строповка длинномерных грузов"; Caption = "Длинномерные грузы поднимают так, чтобы исключить прогиб, раскачивание и перегруз ветвей."; Image = Resolve-ProjectPath "assets\working-visuals\37-59\slide-58-long-loads.png" }
    @{ Index = 57; Kind = "FullImage"; Section = "Схемы строповки по типам грузов"; Lead = "Общие ошибки в схемах строповки разных грузов"; Caption = "Типовые ошибки: неправильные точки подвеса, неверный угол ветвей, отсутствие защиты кромок и потеря устойчивости."; Image = Resolve-ProjectPath "assets\working-visuals\37-59\slide-59-common-scheme-errors.jfif" }
    @{ Index = 58; Kind = "RightImage"; Section = "Требования безопасности при подъеме и перемещении груза"; Lead = "Подготовка крана и рабочей площадки перед началом работ"; Body = "До начала работ проверяют площадку, устойчивость крана, наличие препятствий и организацию опасной зоны.`r`n`r`nПодъем начинают только на подготовленном и понятном рабочем месте."; Image = Resolve-ProjectPath "assets\working-visuals\60-68\slide-60-crane-setup-rules.png" }
    @{ Index = 59; Kind = "RightImage"; Section = "Требования безопасности при подъеме и перемещении груза"; Lead = "Работа крана на опорах, площадка и расстояние до котлована"; Body = "Кран на опорах работает только на подготовленной площадке.`r`n`r`nОсобенно важно соблюдать расстояние до откосов, траншей и котлованов, чтобы исключить просадку и потерю устойчивости."; Image = Resolve-ProjectPath "assets\working-visuals\60-68\slide-61-trench-slope-distance.png" }
    @{ Index = 60; Kind = "FullImage"; Section = "Требования безопасности при подъеме и перемещении груза"; Lead = "Ограждение опасной зоны"; Caption = "В опасную зону не заходят посторонние, а расстояния до груза и препятствий контролируют заранее."; Image = Resolve-ProjectPath "assets\working-visuals\60-68\slide-62-danger-zone.png" }
    @{ Index = 61; Kind = "RightImage"; Section = "Требования безопасности при подъеме и перемещении груза"; Lead = "Работа вблизи ЛЭП"; Body = "Работы возле линий электропередачи выполняют только при соблюдении установленных расстояний и организационных мер.`r`n`r`nНаряд-допуск и контроль безопасной зоны обязательны."; Image = Resolve-ProjectPath "assets\working-visuals\60-68\slide-63-power-lines.png" }
    @{ Index = 62; Kind = "RightImage"; Section = "Требования безопасности при подъеме и перемещении груза"; Lead = "Складирование грузов"; Body = "Грузы складируют устойчиво, с проходами, подкладками и без создания опасности опрокидывания.`r`n`r`nПосле установки груз не должен мешать движению техники и людей."; Image = Resolve-ProjectPath "assets\working-visuals\60-68\slide-64-storage-1.png" }
    @{ Index = 63; Kind = "RightImage"; Section = "Требования безопасности при подъеме и перемещении груза"; Lead = "Основные требования безопасности при работе стропальщика"; Body = "- не работать без ясной задачи`r`n- применять исправную оснастку`r`n- не находиться под грузом`r`n- соблюдать сигналы и дисциплину на площадке`r`n- останавливать работу при риске"; Image = Resolve-ProjectPath "assets\working-visuals\60-68\slide-65-general-safety.png" }
    @{ Index = 64; Kind = "RightImage"; Section = "Требования безопасности при подъеме и перемещении груза"; Lead = "Что запрещено при подъеме и перемещении груза"; Body = "Запрещено подтаскивать груз краном, исправлять его руками в опасной фазе, перемещать людей на грузе и работать с неизвестной массой.`r`n`r`nЗапреты действуют всегда, даже при спешке."; Image = Resolve-ProjectPath "assets\working-visuals\60-68\slide-66-prohibited-actions.png" }
    @{ Index = 65; Kind = "RightImage"; Section = "Требования безопасности при подъеме и перемещении груза"; Lead = "Безопасное нахождение в рабочей зоне"; Body = "Стропальщик выбирает позицию, где видит груз, не попадает под траекторию его движения и сохраняет путь отхода.`r`n`r`nПравильная позиция снижает риск травмы даже при нештатной ситуации."; Image = Resolve-ProjectPath "assets\working-visuals\60-68\slide-67-safe-positioning.png" }
    @{ Index = 66; Kind = "Text"; Section = "Требования безопасности при подъеме и перемещении груза"; Lead = "Безопасность при работе с разными типами грузов"; Body = "Для труб, металла, железобетона, оборудования и длинномерных грузов схема подъема отличается, но правило одно: учитывают центр тяжести, устойчивость, острые кромки и ограничения оснастки.`r`n`r`nНельзя переносить привычную схему с одного типа груза на другой без проверки." }
    @{ Index = 67; Kind = "RightImage"; Section = "Аварийные ситуации и действия при нарушениях"; Lead = "Что относится к аварийной ситуации"; Body = "К аварийным относят потерю устойчивости груза, повреждение оснастки, отказ техники, опасное приближение к препятствиям и любые события, при которых продолжение работ становится рискованным.`r`n`r`nГлавная задача - вовремя распознать угрозу."; Image = Resolve-ProjectPath "assets\working-visuals\69-76\slide-69-emergency-case.png" }
    @{ Index = 68; Kind = "RightImage"; Section = "Аварийные ситуации и действия при нарушениях"; Lead = "Первые действия стропальщика при аварийной ситуации"; Body = "Первое действие - немедленно остановить опасное движение и предупредить участников работ.`r`n`r`nДалее выводят людей из опасной зоны, докладывают ответственному лицу и действуют по установленному порядку."; Image = Resolve-ProjectPath "assets\working-visuals\69-76\slide-70-stop-signal.png" }
    @{ Index = 69; Kind = "RightImage"; Section = "Аварийные ситуации и действия при нарушениях"; Lead = "Действия при выявлении нарушений до начала или в ходе работ"; Body = "Если нарушение замечено до старта или в процессе работ, стропальщик обязан остановить операцию, сообщить о проблеме и не продолжать подъем до устранения причины.`r`n`r`nМолчаливое согласие с нарушением недопустимо."; Image = Resolve-ProjectPath "assets\working-visuals\69-76\slide-71-report-and-warn.png" }
    @{ Index = 70; Kind = "RightImage"; Section = "Аварийные ситуации и действия при нарушениях"; Lead = "Чего нельзя делать в аварийной ситуации"; Body = "Нельзя исправлять ситуацию рывком, заходить под груз, действовать в одиночку без команды и игнорировать признаки потери устойчивости.`r`n`r`nСамовольные действия в аварийной фазе часто усугубляют последствия."; Image = Resolve-ProjectPath "assets\working-visuals\69-76\slide-72-prohibited-actions.png" }
    @{ Index = 71; Kind = "RightImage"; Section = "Аварийные ситуации и действия при нарушениях"; Lead = "Разбор типовой аварийной ситуации"; Body = "Разбор инцидента помогает увидеть цепочку ошибок: неподготовленная площадка, неверная схема, отсутствие остановки при первых признаках опасности.`r`n`r`nАвария почти всегда развивается из нескольких пропущенных предупреждений."; Image = Resolve-ProjectPath "assets\working-visuals\69-76\slide-73-typical-incident-photo.jfif" }
    @{ Index = 72; Kind = "FullImage"; Section = "Итоговое повторение"; Lead = "Повторение знаковых сигналов"; Caption = "Перед итоговой проверкой важно ещё раз закрепить основные сигналы и их однозначное понимание."; Image = Resolve-ProjectPath "assets\working-visuals\69-76\slide-74-signal-cheatsheet.png" }
    @{ Index = 73; Kind = "FullImage"; Section = "Итоговое повторение"; Lead = "Повторение алгоритма выполнения работ стропальщика"; Caption = "Безопасная работа строится на последовательности действий, а не на импровизации в опасной фазе."; Image = Resolve-ProjectPath "assets\working-visuals\69-76\slide-75-algorithm-background.png" }
    @{ Index = 74; Kind = "FullImage"; Section = "Итоговое повторение"; Lead = "Повторение ключевых запретов и требований безопасности"; Caption = "Под груз не заходят, сомнительную оснастку не используют, а при риске работу немедленно останавливают."; Image = Resolve-ProjectPath "assets\working-visuals\69-76\slide-76-prohibited-actions.png" }
    @{ Index = 75; Kind = "Transition"; Section = "Аттестация и тестовый блок"; Lead = "Переход к итоговой проверке знаний"; Caption = "Теоретическая часть завершена. Дальше - итоговая аттестация по ключевым правилам работы стропальщика."; Image = Resolve-ProjectPath "assets\working-visuals\77-79\slide-77-attestation-transition.png" }
    @{ Index = 76; Kind = "Text"; Section = "Аттестация и тестовый блок"; Lead = "Критерии аттестации"; Body = "Итоговая аттестация проводится в форме тестирования.`r`n`r`n- 15 вопросов`r`n- по каждому вопросу один правильный ответ`r`n- успешный результат: 12 правильных ответов и более`r`n`r`nОценивается понимание ключевых правил безопасности и действий стропальщика." }
    @{ Index = 77; Kind = "RightImage"; Section = "Аттестация и тестовый блок"; Lead = "Инструкция по заполнению бланка ответов"; Body = "Перед началом заполните ФИО, дату и подпись.`r`n`r`nОтветы отмечайте только в строках, соответствующих номеру вопроса. Для каждого вопроса выбирайте один вариант ответа. Отметка должна быть четкой и однозначной."; Image = Resolve-ProjectPath "assets\working-visuals\77-79\slide-79-answer-sheet.png" }
    @{ Index = 78; Kind = "Text"; Section = "Аттестация и тестовый блок"; Lead = "Тестовый вопрос 1"; Body = "Какое наименьшее расстояние допускается при работе крана вблизи ЛЭП напряжением 380 вольт от выступающей части крана, груза до ближайшего провода по воздуху, при наличии наряда-допуска и разрешения на работу в охранной зоне ЛЭП?`r`n`r`nА. 2 метра`r`nБ. 4 метра`r`nВ. 1,5 метра`r`nГ. 8 метров" }
    @{ Index = 79; Kind = "Text"; Section = "Аттестация и тестовый блок"; Lead = "Тестовый вопрос 2"; Body = "Какие требования предъявляются при подъеме и опускании груза, установленного вблизи стены, штабеля, вагона?`r`n`r`nА. Работа производится в присутствии лица, ответственного за безопасное производство работ.`r`nБ. В подобном случае всегда назначается сигнальщик.`r`nВ. Чтобы между стеной, штабелем, вагоном и грузом не находились люди.`r`nГ. Чтобы между стеной, грузом и другими предметами было расстояние не менее 1 м." }
    @{ Index = 80; Kind = "Text"; Section = "Аттестация и тестовый блок"; Lead = "Тестовый вопрос 3"; Body = "На какую высоту предварительно должен быть приподнят предельный груз при подъеме?`r`n`r`nА. Приподнять груз на 0,5 для проверки устойчивости крана.`r`nБ. Поднимать груз только после дополнительного инструктажа.`r`nВ. Обязательно должна быть схема строповки груза.`r`nГ. Приподнять груз на 200-300 мм для проверки работы тормозов." }
    @{ Index = 81; Kind = "ImageQuestion"; Section = "Аттестация и тестовый блок"; Lead = "Тестовый вопрос 4"; Caption = "Укажите, на каком рисунке изображен многоветвевой строп?"; Image = Resolve-ProjectPath "assets\working-visuals\80-94\slide-83-mnogovetvevoy-strop-options.png" }
    @{ Index = 82; Kind = "Text"; Section = "Аттестация и тестовый блок"; Lead = "Тестовый вопрос 5"; Body = "В каком случае стропальщику запрещается подавать сигнал на опускание груза в кузов машины, стоящей под погрузкой?`r`n`r`nА. Если производится погрузка кирпича на поддонах без ограждений.`r`nБ. Если водитель машины находится в кабине машины.`r`nВ. Если стропальщик вышел из кузова, куда должен быть опущен груз.`r`nГ. Если на строповку груза нет схемы строповки, но присутствует лицо, ответственное за безопасное производство работ кранами." }
    @{ Index = 83; Kind = "Text"; Section = "Аттестация и тестовый блок"; Lead = "Тестовый вопрос 6"; Body = "Кем и в каких случаях назначается сигнальщик?`r`n`r`nА. Бригадиром при недостаточном освещении рабочего места.`r`nБ. Инженером по охране труда при снегопаде.`r`nВ. ИТР по надзору за кранами, когда зона, обслуживаемая краном, не обозревается из кабины крановщика.`r`nГ. Лицом, ответственным за безопасное производство работ кранами, когда крановщик не видит стропальщика из-за плохой обзорности." }
    @{ Index = 84; Kind = "ImageQuestion"; Section = "Аттестация и тестовый блок"; Lead = "Тестовый вопрос 7"; Caption = "Укажите, на каком рисунке изображен цепной многоветвевой строп?"; Image = Resolve-ProjectPath "assets\working-visuals\80-94\slide-86-chain-multibranch-options.png" }
    @{ Index = 85; Kind = "Text"; Section = "Аттестация и тестовый блок"; Lead = "Тестовый вопрос 8"; Body = "На каком расстоянии от края детали производится обвязка длинномерных грузов, во избежание их прогиба?`r`n`r`nА. На 1/3 длины груза.`r`nБ. На 1/4 длины груза.`r`nВ. На 0,3 длины груза.`r`nГ. На 0,4 длины груза." }
    @{ Index = 86; Kind = "Text"; Section = "Аттестация и тестовый блок"; Lead = "Тестовый вопрос 9"; Body = "Как застропить железобетонную плиту, если у нее сломана одна петля?`r`n`r`nА. Стропить нельзя.`r`nБ. Двумя облегченными петлевыми стропами в обхват, с применением подкладок под острые углы.`r`nВ. Цеплять за оставшиеся петли.`r`nГ. Цеплять за две петли по диагонали." }
    @{ Index = 87; Kind = "Text"; Section = "Аттестация и тестовый блок"; Lead = "Тестовый вопрос 10"; Body = "Как подать сигнал «ПОВЕРНИ СТРЕЛУ»?`r`n`r`nА. Движение вытянутой рукой по направлению движения.`r`nБ. Движение согнутой в локте рукой по направлению движения стрелы.`r`nВ. Движение вытянутой рукой, ладонь по направлению движения.`r`nГ. Движение согнутой в локте рукой из стороны в сторону, ладонью вниз." }
    @{ Index = 88; Kind = "Text"; Section = "Аттестация и тестовый блок"; Lead = "Тестовый вопрос 11"; Body = "Что должен сделать стропальщик перед опусканием груза?`r`n`r`nА. Осмотреть место, на которое укладывается груз и доложить мастеру.`r`nБ. Убедиться, что в проходе, куда будет опускаться груз, осталось свободное место шириной не менее 15 см.`r`nВ. Уложить прочные подкладки при установке груза на подземные кабеля.`r`nГ. На место установки груза предварительно уложить прочные прокладки для удобства извлечения строп из-под груза." }
    @{ Index = 89; Kind = "Text"; Section = "Аттестация и тестовый блок"; Lead = "Тестовый вопрос 12"; Body = "Как подать сигнал «передвинуть кран»?`r`n`r`nА. Движение согнутой в локте рукой, ладонью по направлению движения.`r`nБ. Движение вытянутой рукой, ладонь обращена в сторону требуемого движения.`r`nВ. Движение согнутой в локте рукой из стороны в сторону ладонью вниз.`r`nГ. Движение вытянутой руки снизу вверх, ладонью вверх." }
    @{ Index = 90; Kind = "ImageQuestion"; Section = "Аттестация и тестовый блок"; Lead = "Тестовый вопрос 13"; Caption = "Укажите на рисунке чалочный крюк."; Image = Resolve-ProjectPath "assets\working-visuals\80-94\slide-92-hook-options.png" }
    @{ Index = 91; Kind = "Text"; Section = "Аттестация и тестовый блок"; Lead = "Тестовый вопрос 14"; Body = "Как подобрать стропы для подъема груза?`r`n`r`nА. Угол между ветвями не должен превышать 90°.`r`nБ. Соответствующие весу, виду и габариту поднимаемого груза с учетом числа ветвей и такой длины, чтобы угол между ветвями не превышал 90°.`r`nВ. Испытанные в соответствии с требованиями правил по кранам, в соответствии со схемой строповки груза.`r`nГ. По грузоподъемности с учетом числа ветвей." }
    @{ Index = 92; Kind = "Text"; Section = "Аттестация и тестовый блок"; Lead = "Тестовый вопрос 15"; Body = "Как показать сигнал «опустить груз или крюк»?`r`n`r`nА. Прерывистое движение вверх, руки перед грудью, рука согнута в локте.`r`nБ. Прерывистое движение вниз, руки перед грудью, ладонь вниз, рука согнута в локте.`r`nВ. Подъем вытянутой руки, предварительно опущенной до вертикального положения, ладонь раскрыта.`r`nГ. Прерывистое движение вниз вытянутой руки, ладонь обращена вниз." }
    @{ Index = 93; Kind = "Final"; Section = "Аттестация и тестовый блок"; Lead = "Спасибо за внимание"; Image = Resolve-ProjectPath "assets\working-visuals\95\slide-95-final-background.jpeg" }
)

$powerPoint = $null
$presentation = $null

try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.Visible = -1
    $presentation = $powerPoint.Presentations.Open($resolvedOutputPath, $false, $false, $false)

    Configure-Slide23 -Slide $presentation.Slides.Item(23) -ImagePath (Resolve-ProjectPath "assets\working-visuals\22-36\slide-23-strength-loss-chart.png")
    $presentation.Slides.Item(24).Delete()

    foreach ($spec in ($specs | Sort-Object Index -Descending)) {
        $slide = Replace-SlideWithTemplate -Presentation $presentation -Index $spec.Index -TemplateIndex 23
        switch ($spec.Kind) {
            "Text" {
                Configure-TextSlide -Slide $slide -Section $spec.Section -Lead $spec.Lead -Body $spec.Body
            }
            "RightImage" {
                Configure-RightImageSlide -Slide $slide -Section $spec.Section -Lead $spec.Lead -Body $spec.Body -ImagePath $spec.Image
            }
            "FullImage" {
                Configure-FullImageSlide -Slide $slide -Section $spec.Section -Lead $spec.Lead -Caption $spec.Caption -ImagePath $spec.Image
            }
            "Transition" {
                Configure-TransitionSlide -Slide $slide -Section $spec.Section -Lead $spec.Lead -Caption $spec.Caption -ImagePath $spec.Image
            }
            "ImageQuestion" {
                Configure-ImageQuestionSlide -Slide $slide -Section $spec.Section -Lead $spec.Lead -Question $spec.Caption -ImagePath $spec.Image
            }
            "Final" {
                Configure-FinalSlide -Slide $slide -Section $spec.Section -Lead $spec.Lead -ImagePath $spec.Image
            }
        }
    }

    for ($index = 1; $index -le $presentation.Slides.Count; $index++) {
        Set-SlideNumber -Slide $presentation.Slides.Item($index) -Number $index
    }

    $presentation.Save()

    $previewDirectory = Resolve-ProjectPath ".codex-temp"
    Ensure-Directory -Path $previewDirectory
    foreach ($index in @(23, 24, 35, 50, 60, 75, 81, 93)) {
        $presentation.Slides.Item($index).Export((Join-Path $previewDirectory ("slide-{0:D3}-complete.png" -f $index)), "PNG", 1600, 900)
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

Write-Output "Готово: $resolvedOutputPath"
