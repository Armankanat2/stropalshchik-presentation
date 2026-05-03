param(
    [string]$SourcePath = "assets\исходные-презентации\КУРС для СТРОПАЛЬЩИКА_III.pptx",
    [string]$OutputPath = "deliverables\издательство\КУРС для СТРОПАЛЬЩИКА_ФИНАЛ_95_слайдов.pptx"
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing

function Resolve-ProjectPath {
    param([string]$RelativePath)

    $projectRoot = Split-Path -Parent $PSScriptRoot
    return [System.IO.Path]::GetFullPath((Join-Path $projectRoot $RelativePath))
}

function Ensure-Directory {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path | Out-Null
    }
}

function Remove-ShapeIfExists {
    param(
        $Slide,
        [string]$Name
    )

    try {
        $shape = $Slide.Shapes.Item($Name)
        $shape.Delete()
    } catch {
    }
}

function Set-ShapeText {
    param(
        $Shape,
        [string]$Text,
        [double]$FontSize
    )

    $Shape.TextFrame.TextRange.Text = $Text
    $Shape.TextFrame.TextRange.Font.Size = $FontSize
    try {
        $Shape.TextFrame.TextRange.ParagraphFormat.Bullet.Type = 0
    } catch {
    }
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

        return @{
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
    return $picture
}

function Set-SlideNumber {
    param(
        $Slide,
        [int]$Number
    )

    $shape = $Slide.Shapes.Item("TextBox 5")
    Set-ShapeText -Shape $shape -Text ([string]$Number) -FontSize 16
}

function Configure-TextSlide {
    param(
        $Slide,
        [string]$Title,
        [string]$BodyTop,
        [string]$BodyBottom
    )

    Remove-ShapeIfExists -Slide $Slide -Name "Picture 2"

    $titleShape = $Slide.Shapes.Item("TextBox 10")
    $bodyTopShape = $Slide.Shapes.Item("TextBox 11")
    $bodyBottomShape = $Slide.Shapes.Item("TextBox 9")

    $titleShape.Left = 0
    $titleShape.Top = 20.5
    $titleShape.Width = 513.1
    $titleShape.Height = 26.7

    $bodyTopShape.Left = 15
    $bodyTopShape.Top = 88
    $bodyTopShape.Width = 690
    $bodyTopShape.Height = 105

    $bodyBottomShape.Left = 15
    $bodyBottomShape.Top = 200
    $bodyBottomShape.Width = 690
    $bodyBottomShape.Height = 255

    Set-ShapeText -Shape $titleShape -Text $Title -FontSize 20
    Set-ShapeText -Shape $bodyTopShape -Text $BodyTop -FontSize 22
    Set-ShapeText -Shape $bodyBottomShape -Text $BodyBottom -FontSize 18
}

function Configure-RightImageSlide {
    param(
        $Slide,
        [string]$Title,
        [string]$BodyTop,
        [string]$BodyBottom,
        [string]$ImagePath
    )

    $titleShape = $Slide.Shapes.Item("TextBox 10")
    $bodyTopShape = $Slide.Shapes.Item("TextBox 11")
    $bodyBottomShape = $Slide.Shapes.Item("TextBox 9")

    $titleShape.Left = 0
    $titleShape.Top = 20.5
    $titleShape.Width = 513.1
    $titleShape.Height = 26.7

    $bodyTopShape.Left = 15
    $bodyTopShape.Top = 88
    $bodyTopShape.Width = 400
    $bodyTopShape.Height = 80

    $bodyBottomShape.Left = 15
    $bodyBottomShape.Top = 175
    $bodyBottomShape.Width = 400
    $bodyBottomShape.Height = 280

    Set-ShapeText -Shape $titleShape -Text $Title -FontSize 20
    Set-ShapeText -Shape $bodyTopShape -Text $BodyTop -FontSize 22
    Set-ShapeText -Shape $bodyBottomShape -Text $BodyBottom -FontSize 18

    Replace-ContentPicture -Slide $Slide -ImagePath $ImagePath -Left 430 -Top 95 -MaxWidth 270 -MaxHeight 320 | Out-Null
}

function Configure-ImageQuestionSlide {
    param(
        $Slide,
        [string]$Title,
        [string]$Question,
        [string]$ImagePath
    )

    $titleShape = $Slide.Shapes.Item("TextBox 10")
    $bodyTopShape = $Slide.Shapes.Item("TextBox 11")
    $bodyBottomShape = $Slide.Shapes.Item("TextBox 9")

    $titleShape.Left = 0
    $titleShape.Top = 20.5
    $titleShape.Width = 513.1
    $titleShape.Height = 26.7

    $bodyTopShape.Left = 15
    $bodyTopShape.Top = 80
    $bodyTopShape.Width = 690
    $bodyTopShape.Height = 50

    $bodyBottomShape.Left = 15
    $bodyBottomShape.Top = 448
    $bodyBottomShape.Width = 690
    $bodyBottomShape.Height = 20

    Set-ShapeText -Shape $titleShape -Text $Title -FontSize 20
    Set-ShapeText -Shape $bodyTopShape -Text $Question -FontSize 20
    Set-ShapeText -Shape $bodyBottomShape -Text "" -FontSize 10

    Replace-ContentPicture -Slide $Slide -ImagePath $ImagePath -Left 65 -Top 135 -MaxWidth 590 -MaxHeight 300 | Out-Null
}

function Configure-FinalSlide {
    param(
        $Slide,
        [string]$Title,
        [string]$ImagePath
    )

    $titleShape = $Slide.Shapes.Item("TextBox 10")
    $bodyTopShape = $Slide.Shapes.Item("TextBox 11")
    $bodyBottomShape = $Slide.Shapes.Item("TextBox 9")

    Set-ShapeText -Shape $titleShape -Text $Title -FontSize 20
    Set-ShapeText -Shape $bodyTopShape -Text "" -FontSize 10
    Set-ShapeText -Shape $bodyBottomShape -Text "" -FontSize 10

    $titleShape.Left = 0
    $titleShape.Top = 20.5
    $titleShape.Width = 513.1
    $titleShape.Height = 26.7

    $bodyTopShape.Left = 0
    $bodyTopShape.Top = 0
    $bodyTopShape.Width = 1
    $bodyTopShape.Height = 1

    $bodyBottomShape.Left = 0
    $bodyBottomShape.Top = 0
    $bodyBottomShape.Width = 1
    $bodyBottomShape.Height = 1

    Replace-ContentPicture -Slide $Slide -ImagePath $ImagePath -Left 0 -Top 60 -MaxWidth 720 -MaxHeight 430 | Out-Null

    $textBox = $Slide.Shapes.AddTextbox(1, 28, 320, 280, 90)
    $textBox.Name = "FinalMessage"
    Set-ShapeText -Shape $textBox -Text "Стропальщик`r`nSafety ferst" -FontSize 26
    $textBox.TextFrame.TextRange.Font.Color.RGB = 16777215
    $textBox.Fill.Visible = 0
    $textBox.Line.Visible = 0

    $captionBox = $Slide.Shapes.AddTextbox(1, 30, 422, 220, 22)
    $captionBox.Name = "FinalCaption"
    Set-ShapeText -Shape $captionBox -Text "Первое безопасность" -FontSize 14
    $captionBox.TextFrame.TextRange.Font.Color.RGB = 16777215
    $captionBox.Fill.Visible = 0
    $captionBox.Line.Visible = 0

    $Slide.Shapes.Item("Picture 4").ZOrder(0) | Out-Null
    $Slide.Shapes.Item("TextBox 5").ZOrder(0) | Out-Null
    $titleShape.ZOrder(0) | Out-Null
    $textBox.ZOrder(0) | Out-Null
    $captionBox.ZOrder(0) | Out-Null
}

function Add-TemplateSlide {
    param(
        $Presentation,
        [int]$TemplateIndex
    )

    $template = $Presentation.Slides.Item($TemplateIndex)
    $duplicateRange = $template.Duplicate()
    $slide = $duplicateRange.Item(1)
    $slide.MoveTo($Presentation.Slides.Count)
    return $slide
}

$resolvedSourcePath = Resolve-ProjectPath $SourcePath
$resolvedOutputPath = Resolve-ProjectPath $OutputPath
$outputDirectory = Split-Path -Parent $resolvedOutputPath
Ensure-Directory -Path $outputDirectory

Copy-Item -LiteralPath $resolvedSourcePath -Destination $resolvedOutputPath -Force

$slides = @(
    @{
        Number = 77
        Kind = "RightImage"
        Title = "Итоговая аттестация"
        BodyTop = "Теоретическая часть курса завершена."
        BodyBottom = "Далее проводится итоговая аттестация в форме тестирования.`r`nТестовый блок проверяет понимание ключевых правил и действий стропальщика.`r`nВо время тестирования внимательно читайте формулировки вопросов и вариантов ответа."
        Image = Resolve-ProjectPath "assets\working-visuals\77-79\slide-77-attestation-transition.png"
    }
    @{
        Number = 78
        Kind = "Text"
        Title = "Критерии аттестации"
        BodyTop = "Итоговая аттестация проводится в форме тестирования."
        BodyBottom = "Тест включает 15 вопросов.`r`nК каждому вопросу предусмотрен один правильный ответ.`r`nУспешная сдача теста: 12 правильных ответов и более.`r`nРезультат определяется по количеству верных ответов."
    }
    @{
        Number = 79
        Kind = "RightImage"
        Title = "Порядок заполнения бланка"
        BodyTop = "Перед началом заполните ФИО, дату и подпись."
        BodyBottom = "Ответы отмечайте только в строках, соответствующих номеру вопроса.`r`nПо каждому вопросу выбирайте один вариант ответа.`r`nОтметка должна быть четкой и однозначной.`r`nДля текущего теста используются вопросы 1-15."
        Image = Resolve-ProjectPath "assets\working-visuals\77-79\slide-79-answer-sheet.png"
    }
    @{
        Number = 80
        Kind = "Text"
        Title = "Тестовый вопрос 1"
        BodyTop = "Какое наименьшее расстояние допускается при работе крана вблизи ЛЭП напряжением 380 вольт от выступающей части крана, груза до ближайшего провода по воздуху, при наличии наряда-допуска и разрешения на работу в охранной зоне ЛЭП?"
        BodyBottom = "А. 2 метра`r`nБ. 4 метра`r`nВ. 1,5 метра`r`nГ. 8 метров"
    }
    @{
        Number = 81
        Kind = "Text"
        Title = "Тестовый вопрос 2"
        BodyTop = "Какие требования предъявляются при подъеме и опускании груза, установленного вблизи стены, штабеля, вагона?"
        BodyBottom = "А. Работа производится в присутствии лица, ответственного за безопасное производство работ.`r`nБ. В подобном случае всегда назначается сигнальщик.`r`nВ. Чтобы между стеной, штабелем, вагоном и грузом не находились люди.`r`nГ. Чтобы между стеной, грузом и другими предметами было расстояние не менее 1 м."
    }
    @{
        Number = 82
        Kind = "Text"
        Title = "Тестовый вопрос 3"
        BodyTop = "На какую высоту предварительно должен быть приподнят предельный груз при подъеме?"
        BodyBottom = "А. Приподнять груз на 0,5 для проверки устойчивости крана.`r`nБ. Поднимать груз только после дополнительного инструктажа.`r`nВ. Обязательно должна быть схема строповки груза.`r`nГ. Приподнять груз на 200-300 мм для проверки работы тормозов."
    }
    @{
        Number = 83
        Kind = "ImageQuestion"
        Title = "Тестовый вопрос 4"
        BodyTop = "Укажите, на каком рисунке изображен многоветвевой строп?"
        Image = Resolve-ProjectPath "assets\working-visuals\80-94\slide-83-mnogovetvevoy-strop-options.png"
    }
    @{
        Number = 84
        Kind = "Text"
        Title = "Тестовый вопрос 5"
        BodyTop = "В каком случае стропальщику запрещается подавать сигнал на опускание груза в кузов машины, стоящей под погрузкой?"
        BodyBottom = "А. Если производится погрузка кирпича на поддонах без ограждений.`r`nБ. Если водитель машины находится в кабине машины.`r`nВ. Если стропальщик вышел из кузова, куда должен быть опущен груз.`r`nГ. Если на строповку груза нет схемы строповки, но присутствует лицо, ответственное за безопасное производство работ кранами."
    }
    @{
        Number = 85
        Kind = "Text"
        Title = "Тестовый вопрос 6"
        BodyTop = "Кем и в каких случаях назначается сигнальщик?"
        BodyBottom = "А. Бригадиром при недостаточном освещении рабочего места.`r`nБ. Инженером по охране труда при снегопаде.`r`nВ. ИТР по надзору за кранами, когда зона, обслуживаемая краном, не обозревается из кабины крановщика.`r`nГ. Лицом, ответственным за безопасное производство работ кранами, когда крановщик не видит стропальщика из-за плохой обзорности."
    }
    @{
        Number = 86
        Kind = "ImageQuestion"
        Title = "Тестовый вопрос 7"
        BodyTop = "Укажите, на каком рисунке изображен цепной многоветвевой строп?"
        Image = Resolve-ProjectPath "assets\working-visuals\80-94\slide-86-chain-multibranch-options.png"
    }
    @{
        Number = 87
        Kind = "Text"
        Title = "Тестовый вопрос 8"
        BodyTop = "На каком расстоянии от края детали производится обвязка длинномерных грузов, во избежание их прогиба?"
        BodyBottom = "А. На 1/3 длины груза.`r`nБ. На 1/4 длины груза.`r`nВ. На 0,3 длины груза.`r`nГ. На 0,4 длины груза."
    }
    @{
        Number = 88
        Kind = "Text"
        Title = "Тестовый вопрос 9"
        BodyTop = "Как застропить железобетонную плиту, если у нее сломана одна петля?"
        BodyBottom = "А. Стропить нельзя.`r`nБ. Двумя облегченными петлевыми стропами в обхват, с применением подкладок под острые углы.`r`nВ. Цеплять за оставшиеся петли.`r`nГ. Цеплять за две петли по диагонали."
    }
    @{
        Number = 89
        Kind = "Text"
        Title = "Тестовый вопрос 10"
        BodyTop = "Как подать сигнал «ПОВЕРНИ СТРЕЛУ»?"
        BodyBottom = "А. Движение вытянутой рукой по направлению движения.`r`nБ. Движение согнутой в локте рукой по направлению движения стрелы.`r`nВ. Движение вытянутой рукой, ладонь по направлению движения.`r`nГ. Движение согнутой в локте рукой из стороны в сторону, ладонью вниз."
    }
    @{
        Number = 90
        Kind = "Text"
        Title = "Тестовый вопрос 11"
        BodyTop = "Что должен сделать стропальщик перед опусканием груза?"
        BodyBottom = "А. Осмотреть место, на которое укладывается груз и доложить мастеру.`r`nБ. Убедиться, что в проходе, куда будет опускаться груз, осталось свободное место шириной не менее 15 см.`r`nВ. Уложить прочные подкладки при установке груза на подземные кабеля.`r`nГ. На место установки груза предварительно уложить прочные прокладки для удобства извлечения строп из-под груза."
    }
    @{
        Number = 91
        Kind = "Text"
        Title = "Тестовый вопрос 12"
        BodyTop = "Как подать сигнал «передвинуть кран»?"
        BodyBottom = "А. Движение согнутой в локте рукой, ладонью по направлению движения.`r`nБ. Движение вытянутой рукой, ладонь обращена в сторону требуемого движения.`r`nВ. Движение согнутой в локте рукой из стороны в сторону ладонью вниз.`r`nГ. Движение вытянутой руки снизу вверх, ладонью вверх."
    }
    @{
        Number = 92
        Kind = "ImageQuestion"
        Title = "Тестовый вопрос 13"
        BodyTop = "Укажите на рисунке чалочный крюк."
        Image = Resolve-ProjectPath "assets\working-visuals\80-94\slide-92-hook-options.png"
    }
    @{
        Number = 93
        Kind = "Text"
        Title = "Тестовый вопрос 14"
        BodyTop = "Как подобрать стропы для подъема груза?"
        BodyBottom = "А. Угол между ветвями не должен превышать 90°.`r`nБ. Соответствующие весу, виду и габариту поднимаемого груза с учетом числа ветвей и такой длины, чтобы угол между ветвями не превышал 90°.`r`nВ. Испытанные в соответствии с требованиями правил по кранам, в соответствии со схемой строповки груза.`r`nГ. По грузоподъемности с учетом числа ветвей."
    }
    @{
        Number = 94
        Kind = "Text"
        Title = "Тестовый вопрос 15"
        BodyTop = "Как показать сигнал «опустить груз или крюк»?"
        BodyBottom = "А. Прерывистое движение вверх, руки перед грудью, рука согнута в локте.`r`nБ. Прерывистое движение вниз, руки перед грудью, ладонь вниз, рука согнута в локте.`r`nВ. Подъем вытянутой руки, предварительно опущенной до вертикального положения, ладонь раскрыта.`r`nГ. Прерывистое движение вниз вытянутой руки, ладонь обращена вниз."
    }
    @{
        Number = 95
        Kind = "Final"
        Title = "Спасибо за внимание"
        Image = Resolve-ProjectPath "assets\working-visuals\95\slide-95-final-background.jpeg"
    }
)

$powerPoint = $null
$presentation = $null

try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.Visible = -1
    $presentation = $powerPoint.Presentations.Open($resolvedOutputPath, $false, $false, $false)

    for ($index = 79; $index -ge 77; $index--) {
        $presentation.Slides.Item($index).Delete()
    }

    foreach ($slideSpec in $slides) {
        $slide = Add-TemplateSlide -Presentation $presentation -TemplateIndex 74
        Set-SlideNumber -Slide $slide -Number $slideSpec.Number

        switch ($slideSpec.Kind) {
            "Text" {
                Configure-TextSlide -Slide $slide -Title $slideSpec.Title -BodyTop $slideSpec.BodyTop -BodyBottom $slideSpec.BodyBottom
            }
            "RightImage" {
                Configure-RightImageSlide -Slide $slide -Title $slideSpec.Title -BodyTop $slideSpec.BodyTop -BodyBottom $slideSpec.BodyBottom -ImagePath $slideSpec.Image
            }
            "ImageQuestion" {
                Configure-ImageQuestionSlide -Slide $slide -Title $slideSpec.Title -Question $slideSpec.BodyTop -ImagePath $slideSpec.Image
            }
            "Final" {
                Configure-FinalSlide -Slide $slide -Title $slideSpec.Title -ImagePath $slideSpec.Image
            }
        }
    }

    $presentation.Save()
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

