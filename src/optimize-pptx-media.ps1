param(
    [string]$InputPath = "deliverables\издательство\КУРС для СТРОПАЛЬЩИКА_ФИНАЛ_95_слайдов.pptx",
    [string]$OutputPath = "deliverables\издательство\КУРС для СТРОПАЛЬЩИКА_ФИНАЛ_95_слайдов_облегченная.pptx",
    [int]$MaxLongEdge = 1400,
    [int]$JpegQuality = 82
)

$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem

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

function Get-JpegEncoder {
    return [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() |
        Where-Object { $_.MimeType -eq "image/jpeg" } |
        Select-Object -First 1
}

function Save-Jpeg {
    param(
        [System.Drawing.Image]$Image,
        [string]$Path,
        [int]$Quality
    )

    $encoder = Get-JpegEncoder
    $encoderParams = New-Object System.Drawing.Imaging.EncoderParameters(1)
    $encoderParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality, [long]$Quality)
    $Image.Save($Path, $encoder, $encoderParams)
    $encoderParams.Dispose()
}

function Resize-Image {
    param(
        [System.Drawing.Image]$Image,
        [int]$MaxLongEdge
    )

    $currentLongEdge = [Math]::Max($Image.Width, $Image.Height)
    if ($currentLongEdge -le $MaxLongEdge) {
        return $null
    }

    $ratio = $MaxLongEdge / $currentLongEdge
    $targetWidth = [Math]::Max(1, [int][Math]::Round($Image.Width * $ratio))
    $targetHeight = [Math]::Max(1, [int][Math]::Round($Image.Height * $ratio))

    $bitmap = New-Object System.Drawing.Bitmap($targetWidth, $targetHeight)
    $bitmap.SetResolution($Image.HorizontalResolution, $Image.VerticalResolution)

    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    try {
        $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
        $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
        $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $graphics.DrawImage($Image, 0, 0, $targetWidth, $targetHeight)
    }
    finally {
        $graphics.Dispose()
    }

    return $bitmap
}

function Replace-TextInFile {
    param(
        [string]$Path,
        [string]$OldValue,
        [string]$NewValue
    )

    $content = [System.IO.File]::ReadAllText($Path)
    if ($content.Contains($OldValue)) {
        $updated = $content.Replace($OldValue, $NewValue)
        [System.IO.File]::WriteAllText($Path, $updated)
        return $true
    }

    return $false
}

$resolvedInputPath = Resolve-ProjectPath $InputPath
$resolvedOutputPath = Resolve-ProjectPath $OutputPath
Ensure-Directory -Path (Split-Path -Parent $resolvedOutputPath)

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("pptx-opt-" + [guid]::NewGuid().ToString("N"))
$extractRoot = Join-Path $tempRoot "unzipped"
Ensure-Directory -Path $extractRoot

$beforeLength = (Get-Item -LiteralPath $resolvedInputPath).Length
Copy-Item -LiteralPath $resolvedInputPath -Destination $resolvedOutputPath -Force

try {
    [System.IO.Compression.ZipFile]::ExtractToDirectory($resolvedOutputPath, $extractRoot)
    Remove-Item -LiteralPath $resolvedOutputPath -Force

    $mediaRoot = Join-Path $extractRoot "ppt\media"
    $allTextFiles = Get-ChildItem -LiteralPath $extractRoot -Recurse -File |
        Where-Object { $_.Extension -in ".xml", ".rels" -or $_.Name -eq "[Content_Types].xml" }

    $report = @()

    foreach ($file in Get-ChildItem -LiteralPath $mediaRoot -File) {
        $extension = $file.Extension.ToLowerInvariant()
        if ($extension -notin ".png", ".jpg", ".jpeg") {
            continue
        }

        $originalLength = $file.Length
        $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
        $memoryStream = New-Object System.IO.MemoryStream(,$bytes)
        $image = [System.Drawing.Image]::FromStream($memoryStream, $false, $false)
        try {
            $resized = Resize-Image -Image $image -MaxLongEdge $MaxLongEdge
            $workingImage = if ($resized) { $resized } else { $image }

            $hasAlpha = [System.Drawing.Image]::IsAlphaPixelFormat($image.PixelFormat)
            $shouldConvertToJpeg = (
                $extension -eq ".png" -and
                -not $hasAlpha -and
                $originalLength -ge 350KB
            )

            if ($shouldConvertToJpeg) {
                $newFileName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name) + ".jpg"
                $newPath = Join-Path $mediaRoot $newFileName
                Save-Jpeg -Image $workingImage -Path $newPath -Quality $JpegQuality

                foreach ($textFile in $allTextFiles) {
                    Replace-TextInFile -Path $textFile.FullName -OldValue $file.Name -NewValue $newFileName | Out-Null
                }

                Remove-Item -LiteralPath $file.FullName -Force
                $newLength = (Get-Item -LiteralPath $newPath).Length
                $report += [PSCustomObject]@{
                    File = $file.Name
                    Action = if ($resized) { "resize+convert" } else { "convert" }
                    Before = $originalLength
                    After = $newLength
                }
            }
            elseif ($resized) {
                if ($extension -in ".jpg", ".jpeg") {
                    Save-Jpeg -Image $workingImage -Path $file.FullName -Quality $JpegQuality
                }
                else {
                    $workingImage.Save($file.FullName, [System.Drawing.Imaging.ImageFormat]::Png)
                }

                $newLength = (Get-Item -LiteralPath $file.FullName).Length
                $report += [PSCustomObject]@{
                    File = $file.Name
                    Action = "resize"
                    Before = $originalLength
                    After = $newLength
                }
            }
        }
        finally {
            if ($resized) {
                $resized.Dispose()
            }
            $image.Dispose()
            $memoryStream.Dispose()
        }
    }

    [System.IO.Compression.ZipFile]::CreateFromDirectory($extractRoot, $resolvedOutputPath, [System.IO.Compression.CompressionLevel]::Optimal, $false)
    $afterLength = (Get-Item -LiteralPath $resolvedOutputPath).Length

    $report | Sort-Object Before -Descending | Select-Object -First 15 | Format-Table -AutoSize | Out-String | Write-Output
    Write-Output ("Before={0}" -f $beforeLength)
    Write-Output ("After={0}" -f $afterLength)
    Write-Output ("Saved={0}" -f ($beforeLength - $afterLength))
}
finally {
    if (Test-Path -LiteralPath $tempRoot) {
        Remove-Item -LiteralPath $tempRoot -Recurse -Force
    }
}


