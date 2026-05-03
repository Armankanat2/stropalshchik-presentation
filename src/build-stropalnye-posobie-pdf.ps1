param(
    [Parameter(Mandatory = $true)]
    [string]$SourceDir,
    [Parameter(Mandatory = $true)]
    [string]$OutputPdf,
    [double]$DeskewMinAngle = -1.5,
    [double]$DeskewMaxAngle = 1.5,
    [double]$DeskewStep = 0.1,
    [switch]$EnhanceScan
)

$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Drawing
Add-Type -ReferencedAssemblies @("System.Drawing.dll") -TypeDefinition @"
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;

public static class DocumentDeskew
{
    public static double EstimateFrameAngle(Bitmap source, double minAngle, double maxAngle, double step)
    {
        using (Bitmap scaled = ScaleForAnalysis(source, 400))
        {
            double bestAngle = 0.0;
            double bestScore = double.MinValue;

            for (double angle = minAngle; angle <= maxAngle + 0.0001; angle += step)
            {
                using (Bitmap rotated = RotateBitmap(scaled, angle, Color.Black))
                {
                    double score = ComputeBrightAlignmentScore(rotated, 218);
                    if (score > bestScore)
                    {
                        bestScore = score;
                        bestAngle = angle;
                    }
                }
            }

            return bestAngle;
        }
    }

    public static Rectangle DetectPageBounds(Bitmap source)
    {
        using (Bitmap scaled = ScaleForAnalysis(source, 400))
        {
            Rectangle detected = DetectPageBoundsInBitmap(scaled, 218, 0.62, 6);
            if (detected.Width <= 0 || detected.Height <= 0)
            {
                return new Rectangle(0, 0, source.Width, source.Height);
            }

            double scaleX = source.Width / (double)scaled.Width;
            double scaleY = source.Height / (double)scaled.Height;

            int paddingX = Math.Max(4, (int)Math.Round(6 * scaleX));
            int paddingY = Math.Max(4, (int)Math.Round(6 * scaleY));

            int left = Math.Max(0, (int)Math.Floor(detected.Left * scaleX) + paddingX);
            int top = Math.Max(0, (int)Math.Floor(detected.Top * scaleY) + paddingY);
            int right = Math.Min(source.Width, (int)Math.Ceiling(detected.Right * scaleX) - paddingX);
            int bottom = Math.Min(source.Height, (int)Math.Ceiling(detected.Bottom * scaleY) - paddingY);

            if (right <= left || bottom <= top)
            {
                return new Rectangle(0, 0, source.Width, source.Height);
            }

            return Rectangle.FromLTRB(left, top, right, bottom);
        }
    }

    private static Bitmap ScaleForAnalysis(Bitmap source, int targetWidth)
    {
        if (source.Width <= targetWidth)
        {
            return (Bitmap)source.Clone();
        }

        int targetHeight = (int)Math.Round(source.Height * (targetWidth / (double)source.Width));
        Bitmap scaled = new Bitmap(targetWidth, targetHeight, PixelFormat.Format24bppRgb);
        scaled.SetResolution(source.HorizontalResolution, source.VerticalResolution);

        using (Graphics graphics = Graphics.FromImage(scaled))
        {
            graphics.Clear(Color.White);
            graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            graphics.SmoothingMode = SmoothingMode.HighQuality;
            graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
            graphics.DrawImage(source, 0, 0, targetWidth, targetHeight);
        }

        return scaled;
    }

    public static Bitmap RotateBitmap(Bitmap source, double angle)
    {
        return RotateBitmap(source, angle, Color.White);
    }

    public static Bitmap RotateBitmap(Bitmap source, double angle, Color backgroundColor)
    {
        Bitmap rotated = new Bitmap(source.Width, source.Height, PixelFormat.Format24bppRgb);
        rotated.SetResolution(source.HorizontalResolution, source.VerticalResolution);

        using (Graphics graphics = Graphics.FromImage(rotated))
        {
            graphics.Clear(backgroundColor);
            graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            graphics.SmoothingMode = SmoothingMode.HighQuality;
            graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
            graphics.TranslateTransform(source.Width / 2f, source.Height / 2f);
            graphics.RotateTransform((float)angle);
            graphics.TranslateTransform(-source.Width / 2f, -source.Height / 2f);
            graphics.DrawImage(source, 0, 0, source.Width, source.Height);
        }

        return rotated;
    }

    public static Bitmap CropBitmap(Bitmap source, Rectangle rect)
    {
        return source.Clone(rect, PixelFormat.Format24bppRgb);
    }

    public static Bitmap EnhanceBitmap(Bitmap source)
    {
        Bitmap enhanced = new Bitmap(source.Width, source.Height, PixelFormat.Format24bppRgb);
        enhanced.SetResolution(source.HorizontalResolution, source.VerticalResolution);

        Rectangle rect = new Rectangle(0, 0, source.Width, source.Height);
        BitmapData srcData = source.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
        BitmapData dstData = enhanced.LockBits(rect, ImageLockMode.WriteOnly, PixelFormat.Format24bppRgb);

        try
        {
            int stride = srcData.Stride;
            int bytes = stride * source.Height;
            byte[] srcBuffer = new byte[bytes];
            byte[] dstBuffer = new byte[bytes];
            int[] histogram = new int[256];

            System.Runtime.InteropServices.Marshal.Copy(srcData.Scan0, srcBuffer, 0, bytes);

            for (int y = 0; y < source.Height; y++)
            {
                int rowOffset = y * stride;
                for (int x = 0; x < source.Width; x++)
                {
                    int offset = rowOffset + (x * 3);
                    int gray = (int)Math.Round(
                        (srcBuffer[offset + 2] * 0.299) +
                        (srcBuffer[offset + 1] * 0.587) +
                        (srcBuffer[offset] * 0.114)
                    );
                    gray = ClampByte(gray);
                    histogram[gray]++;
                }
            }

            int totalPixels = source.Width * source.Height;
            int blackPoint = GetPercentile(histogram, totalPixels, 0.01);
            int whitePoint = GetPercentile(histogram, totalPixels, 0.92);

            if (whitePoint <= blackPoint + 10)
            {
                blackPoint = Math.Max(0, blackPoint - 10);
                whitePoint = Math.Min(255, whitePoint + 10);
            }

            double range = Math.Max(1.0, whitePoint - blackPoint);

            for (int y = 0; y < source.Height; y++)
            {
                int rowOffset = y * stride;
                for (int x = 0; x < source.Width; x++)
                {
                    int offset = rowOffset + (x * 3);
                    int gray = (int)Math.Round(
                        (srcBuffer[offset + 2] * 0.299) +
                        (srcBuffer[offset + 1] * 0.587) +
                        (srcBuffer[offset] * 0.114)
                    );

                    double normalized = (gray - blackPoint) / range;
                    if (normalized < 0.0) normalized = 0.0;
                    if (normalized > 1.0) normalized = 1.0;

                    normalized = Math.Pow(normalized, 1.08);

                    if (normalized > 0.72)
                    {
                        double highlight = (normalized - 0.72) / 0.28;
                        normalized = 0.72 + (1.0 - 0.72) * Math.Pow(highlight, 0.72);
                    }

                    if (normalized < 0.55)
                    {
                        double shadow = normalized / 0.55;
                        normalized = 0.55 * Math.Pow(shadow, 1.22);
                    }

                    int outputGray = ClampByte((int)Math.Round(normalized * 255.0));
                    dstBuffer[offset] = (byte)outputGray;
                    dstBuffer[offset + 1] = (byte)outputGray;
                    dstBuffer[offset + 2] = (byte)outputGray;
                }
            }

            System.Runtime.InteropServices.Marshal.Copy(dstBuffer, 0, dstData.Scan0, bytes);
        }
        finally
        {
            source.UnlockBits(srcData);
            enhanced.UnlockBits(dstData);
        }

        return enhanced;
    }

    private static Rectangle DetectPageBoundsInBitmap(Bitmap bitmap, int brightnessThreshold, double brightRatioThreshold, int runLength)
    {
        Rectangle rect = new Rectangle(0, 0, bitmap.Width, bitmap.Height);
        BitmapData data = bitmap.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);

        try
        {
            int stride = data.Stride;
            int bytes = stride * bitmap.Height;
            byte[] buffer = new byte[bytes];
            System.Runtime.InteropServices.Marshal.Copy(data.Scan0, buffer, 0, bytes);

            int[] brightByColumn = new int[bitmap.Width];
            int[] brightByRow = new int[bitmap.Height];

            for (int y = 0; y < bitmap.Height; y++)
            {
                int rowOffset = y * stride;
                for (int x = 0; x < bitmap.Width; x++)
                {
                    int offset = rowOffset + (x * 3);
                    int gray = (buffer[offset] + buffer[offset + 1] + buffer[offset + 2]) / 3;
                    if (gray >= brightnessThreshold)
                    {
                        brightByColumn[x]++;
                        brightByRow[y]++;
                    }
                }
            }

            int left = FindRunStart(brightByColumn, bitmap.Height, brightRatioThreshold, runLength);
            int right = FindRunEnd(brightByColumn, bitmap.Height, brightRatioThreshold, runLength);

            if (left < 0 || right < 0 || right <= left)
            {
                return new Rectangle(0, 0, bitmap.Width, bitmap.Height);
            }

            int usableWidth = Math.Max(1, right - left + 1);
            int top = FindRunStartWithinColumns(buffer, stride, bitmap.Width, bitmap.Height, left, right, brightnessThreshold, brightRatioThreshold, runLength);
            int bottom = FindRunEndWithinColumns(buffer, stride, bitmap.Width, bitmap.Height, left, right, brightnessThreshold, brightRatioThreshold, runLength);

            if (top < 0 || bottom < 0 || bottom <= top)
            {
                return new Rectangle(left, 0, usableWidth, bitmap.Height);
            }

            return Rectangle.FromLTRB(left, top, right + 1, bottom + 1);
        }
        finally
        {
            bitmap.UnlockBits(data);
        }
    }

    private static int FindRunStart(int[] brightCounts, int total, double threshold, int runLength)
    {
        int run = 0;
        for (int i = 0; i < brightCounts.Length; i++)
        {
            double ratio = brightCounts[i] / (double)total;
            if (ratio >= threshold)
            {
                run++;
                if (run >= runLength)
                {
                    return i - runLength + 1;
                }
            }
            else
            {
                run = 0;
            }
        }

        return -1;
    }

    private static int FindRunEnd(int[] brightCounts, int total, double threshold, int runLength)
    {
        int run = 0;
        for (int i = brightCounts.Length - 1; i >= 0; i--)
        {
            double ratio = brightCounts[i] / (double)total;
            if (ratio >= threshold)
            {
                run++;
                if (run >= runLength)
                {
                    return i + runLength - 1;
                }
            }
            else
            {
                run = 0;
            }
        }

        return -1;
    }

    private static int FindRunStartWithinColumns(byte[] buffer, int stride, int width, int height, int left, int right, int brightnessThreshold, double threshold, int runLength)
    {
        int run = 0;
        int usableWidth = Math.Max(1, right - left + 1);

        for (int y = 0; y < height; y++)
        {
            int bright = 0;
            int rowOffset = y * stride;
            for (int x = left; x <= right; x++)
            {
                int offset = rowOffset + (x * 3);
                int gray = (buffer[offset] + buffer[offset + 1] + buffer[offset + 2]) / 3;
                if (gray >= brightnessThreshold)
                {
                    bright++;
                }
            }

            double ratio = bright / (double)usableWidth;
            if (ratio >= threshold)
            {
                run++;
                if (run >= runLength)
                {
                    return y - runLength + 1;
                }
            }
            else
            {
                run = 0;
            }
        }

        return -1;
    }

    private static int FindRunEndWithinColumns(byte[] buffer, int stride, int width, int height, int left, int right, int brightnessThreshold, double threshold, int runLength)
    {
        int run = 0;
        int usableWidth = Math.Max(1, right - left + 1);

        for (int y = height - 1; y >= 0; y--)
        {
            int bright = 0;
            int rowOffset = y * stride;
            for (int x = left; x <= right; x++)
            {
                int offset = rowOffset + (x * 3);
                int gray = (buffer[offset] + buffer[offset + 1] + buffer[offset + 2]) / 3;
                if (gray >= brightnessThreshold)
                {
                    bright++;
                }
            }

            double ratio = bright / (double)usableWidth;
            if (ratio >= threshold)
            {
                run++;
                if (run >= runLength)
                {
                    return y + runLength - 1;
                }
            }
            else
            {
                run = 0;
            }
        }

        return -1;
    }

    private static double ComputeBrightAlignmentScore(Bitmap bitmap, int brightnessThreshold)
    {
        Rectangle rect = new Rectangle(0, 0, bitmap.Width, bitmap.Height);
        BitmapData data = bitmap.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);

        try
        {
            int stride = data.Stride;
            int bytes = stride * bitmap.Height;
            byte[] buffer = new byte[bytes];
            System.Runtime.InteropServices.Marshal.Copy(data.Scan0, buffer, 0, bytes);

            int[] brightColumns = new int[bitmap.Width];
            double score = 0.0;
            for (int y = 0; y < bitmap.Height; y++)
            {
                int brightPixels = 0;
                int rowOffset = y * stride;

                for (int x = 0; x < bitmap.Width; x++)
                {
                    int offset = rowOffset + (x * 3);
                    int gray = (buffer[offset] + buffer[offset + 1] + buffer[offset + 2]) / 3;

                    if (gray >= brightnessThreshold)
                    {
                        brightPixels++;
                        brightColumns[x]++;
                    }
                }

                score += brightPixels * brightPixels;
            }

            for (int x = 0; x < brightColumns.Length; x++)
            {
                score += brightColumns[x] * brightColumns[x];
            }

            return score;
        }
        finally
        {
            bitmap.UnlockBits(data);
        }
    }

    private static int GetPercentile(int[] histogram, int totalPixels, double percentile)
    {
        int target = (int)Math.Round(totalPixels * percentile);
        int cumulative = 0;

        for (int i = 0; i < histogram.Length; i++)
        {
            cumulative += histogram[i];
            if (cumulative >= target)
            {
                return i;
            }
        }

        return histogram.Length - 1;
    }

    private static int ClampByte(int value)
    {
        if (value < 0) return 0;
        if (value > 255) return 255;
        return value;
    }
}
"@

function Get-JpegEncoder {
    return [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() |
        Where-Object { $_.MimeType -eq "image/jpeg" } |
        Select-Object -First 1
}

function Save-JpegFile {
    param(
        [System.Drawing.Image]$Image,
        [string]$OutputPath,
        [System.Drawing.Imaging.ImageCodecInfo]$Encoder,
        [long]$Quality = 95L
    )

    $qualityEncoder = [System.Drawing.Imaging.Encoder]::Quality
    $encoderParams = New-Object System.Drawing.Imaging.EncoderParameters(1)
    $encoderParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter($qualityEncoder, $Quality)

    try {
        $Image.Save($OutputPath, $Encoder, $encoderParams)
    }
    finally {
        $encoderParams.Dispose()
    }
}

function Rotate-IntoCanvas {
    param(
        [System.Drawing.Bitmap]$SourceBitmap,
        [double]$Angle,
        [System.Drawing.Color]$BackgroundColor = [System.Drawing.Color]::White
    )

    return [DocumentDeskew]::RotateBitmap($SourceBitmap, $Angle, $BackgroundColor)
}

function Get-FrameAngle {
    param(
        [System.Drawing.Bitmap]$PageBitmap,
        [double]$MinAngle,
        [double]$MaxAngle,
        [double]$Step
    )

    return [DocumentDeskew]::EstimateFrameAngle($PageBitmap, $MinAngle, $MaxAngle, $Step)
}

function Get-PageBounds {
    param([System.Drawing.Bitmap]$PageBitmap)

    return [DocumentDeskew]::DetectPageBounds($PageBitmap)
}

function Crop-Bitmap {
    param(
        [System.Drawing.Bitmap]$Bitmap,
        [System.Drawing.Rectangle]$Rectangle
    )

    return [DocumentDeskew]::CropBitmap($Bitmap, $Rectangle)
}

function Enhance-Bitmap {
    param([System.Drawing.Bitmap]$Bitmap)

    return [DocumentDeskew]::EnhanceBitmap($Bitmap)
}

function Get-OrderedSourceFiles {
    param([string]$DirectoryPath)

    $files = Get-ChildItem -LiteralPath $DirectoryPath -File
    if (-not $files) {
        throw "No files found in '$DirectoryPath'."
    }

    $ordered = New-Object System.Collections.Generic.List[object]
    $nonNumericFiles = New-Object System.Collections.Generic.List[System.IO.FileInfo]

    foreach ($file in $files) {
        if ($file.Extension -ieq ".pdf") {
            continue
        }

        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)

        if ($baseName -match '^(\d+)-(\d+)$') {
            $ordered.Add([pscustomobject]@{
                File      = $file
                SortOrder = [int]$matches[1]
            })
            continue
        }

        $nonNumericFiles.Add($file)
    }

    if ($nonNumericFiles.Count -gt 1) {
        $unknownNames = ($nonNumericFiles | Select-Object -ExpandProperty Name) -join ", "
        throw "Unable to determine page order for: $unknownNames"
    }

    if ($nonNumericFiles.Count -eq 1) {
        $ordered.Add([pscustomobject]@{
            File      = $nonNumericFiles[0]
            SortOrder = 3
        })
    }

    return $ordered |
        Sort-Object SortOrder, @{ Expression = { $_.File.Name } } |
        Select-Object -ExpandProperty File
}

function Export-PageImages {
    param(
        [System.IO.FileInfo[]]$Files,
        [string]$PageDir,
        [double]$MinAngle,
        [double]$MaxAngle,
        [double]$Step,
        [bool]$UseEnhancement
    )

    $jpegEncoder = Get-JpegEncoder
    if (-not $jpegEncoder) {
        throw "JPEG encoder not found."
    }

    New-Item -ItemType Directory -Force -Path $PageDir | Out-Null

    $pages = New-Object System.Collections.Generic.List[object]
    $pageNumber = 1

    foreach ($file in $Files) {
        $sourceImage = [System.Drawing.Image]::FromFile($file.FullName)
        try {
            $bitmap = New-Object System.Drawing.Bitmap($sourceImage)
            try {
                $splitHeight = [int][Math]::Floor($bitmap.Height / 2)
                $rectangles = @(
                    (New-Object System.Drawing.Rectangle 0, 0, $bitmap.Width, $splitHeight),
                    (New-Object System.Drawing.Rectangle 0, $splitHeight, $bitmap.Width, ($bitmap.Height - $splitHeight))
                )

                foreach ($rect in $rectangles) {
                    $pageBitmap = $bitmap.Clone($rect, $bitmap.PixelFormat)
                    try {
                        $pageBitmap.RotateFlip([System.Drawing.RotateFlipType]::Rotate270FlipNone)
                        $pageAngle = Get-FrameAngle -PageBitmap $pageBitmap -MinAngle $MinAngle -MaxAngle $MaxAngle -Step $Step
                        $alignedBitmap = if ([math]::Abs($pageAngle) -gt 0.0001) {
                            Rotate-IntoCanvas -SourceBitmap $pageBitmap -Angle $pageAngle -BackgroundColor ([System.Drawing.Color]::Black)
                        }
                        else {
                            $pageBitmap.Clone()
                        }

                        try {
                            $pageBounds = Get-PageBounds -PageBitmap $alignedBitmap
                            $croppedBitmap = Crop-Bitmap -Bitmap $alignedBitmap -Rectangle $pageBounds
                        }
                        finally {
                            $alignedBitmap.Dispose()
                        }

                        try {
                            $finalBitmap = if ($UseEnhancement) {
                                Enhance-Bitmap -Bitmap $croppedBitmap
                            }
                            else {
                                $croppedBitmap.Clone()
                            }
                        }
                        finally {
                            $croppedBitmap.Dispose()
                        }

                        try {
                            $pageFileName = "page-{0:D3}.jpg" -f $pageNumber
                            $pagePath = Join-Path $PageDir $pageFileName
                            Save-JpegFile -Image $finalBitmap -OutputPath $pagePath -Encoder $jpegEncoder

                            $pages.Add([pscustomobject]@{
                                PageNumber = $pageNumber
                                FileName   = $pageFileName
                                Width      = $finalBitmap.Width
                                Height     = $finalBitmap.Height
                                Source     = $file.Name
                                Angle      = [math]::Round($pageAngle, 3)
                                Crop       = "$($pageBounds.X),$($pageBounds.Y),$($pageBounds.Width),$($pageBounds.Height)"
                            })

                            $pageNumber++
                        }
                        finally {
                            $finalBitmap.Dispose()
                        }
                    }
                    finally {
                        $pageBitmap.Dispose()
                    }
                }
            }
            finally {
                $bitmap.Dispose()
            }
        }
        finally {
            $sourceImage.Dispose()
        }
    }

    return $pages
}

function Get-ChromePath {
    $candidates = @(
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "$env:ProgramFiles(x86)\Google\Chrome\Application\chrome.exe",
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles(x86)\Microsoft\Edge\Application\msedge.exe"
    )

    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    throw "Chrome or Edge executable not found."
}

function Write-HtmlDocument {
    param(
        [System.Collections.Generic.List[object]]$Pages,
        [string]$HtmlPath
    )

    $pageMarkup = foreach ($page in $Pages) {
        @"
<section class="page">
  <img src="pages/$($page.FileName)" alt="Page $($page.PageNumber)">
</section>
"@
    }

    $html = @"
<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <title>Stropalnye raboty posobie</title>
  <style>
    @page {
      size: A4 portrait;
      margin: 0;
    }

    html, body {
      margin: 0;
      padding: 0;
      background: #ffffff;
    }

    body {
      font-family: sans-serif;
    }

    .page {
      width: 210mm;
      height: 297mm;
      display: flex;
      align-items: center;
      justify-content: center;
      page-break-after: always;
      break-after: page;
      overflow: hidden;
    }

    .page:last-child {
      page-break-after: auto;
      break-after: auto;
    }

    img {
      display: block;
      width: 100%;
      height: 100%;
      object-fit: contain;
    }
  </style>
</head>
<body>
$($pageMarkup -join "`n")
</body>
</html>
"@

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($HtmlPath, $html, $utf8NoBom)
}

$resolvedSourceDir = (Resolve-Path $SourceDir).Path
$resolvedOutputPdf = Join-Path (Resolve-Path ".").Path $OutputPdf
$outputDirectory = Split-Path -Parent $resolvedOutputPdf
$chromePath = Get-ChromePath

New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null

$buildRoot = Join-Path (Join-Path (Resolve-Path ".").Path ".codex-temp") "stropalnye-posobie-pdf"
$pagesDir = Join-Path $buildRoot "pages"
$profileDir = Join-Path $buildRoot "chrome-profile"
$htmlPath = Join-Path $buildRoot "index.html"
$tempPdfPath = Join-Path $buildRoot "output.pdf"

if (Test-Path $buildRoot) {
    Remove-Item -LiteralPath $buildRoot -Recurse -Force
}

New-Item -ItemType Directory -Force -Path $buildRoot | Out-Null
New-Item -ItemType Directory -Force -Path $profileDir | Out-Null

$orderedFiles = Get-OrderedSourceFiles -DirectoryPath $resolvedSourceDir
$pages = Export-PageImages -Files $orderedFiles -PageDir $pagesDir -MinAngle $DeskewMinAngle -MaxAngle $DeskewMaxAngle -Step $DeskewStep -UseEnhancement $EnhanceScan.IsPresent
Write-HtmlDocument -Pages $pages -HtmlPath $htmlPath

$htmlUri = ([System.Uri]$htmlPath).AbsoluteUri
$chromeArgs = @(
    "--headless=new",
    "--disable-gpu",
    "--allow-file-access-from-files",
    "--run-all-compositor-stages-before-draw",
    "--virtual-time-budget=10000",
    "--user-data-dir=$profileDir",
    "--print-to-pdf=$tempPdfPath",
    "--print-to-pdf-no-header",
    $htmlUri
)

& $chromePath $chromeArgs | Out-Null
if ($LASTEXITCODE -ne 0) {
    throw "Browser PDF generation failed with exit code $LASTEXITCODE."
}

if (-not (Test-Path $tempPdfPath)) {
    throw "Generated PDF was not created."
}

Copy-Item -LiteralPath $tempPdfPath -Destination $resolvedOutputPdf -Force

Write-Host "Created PDF: $resolvedOutputPdf"
Write-Host "Source spreads: $($orderedFiles.Count)"
Write-Host "Output pages: $($pages.Count)"
Write-Host "Deskew range: $DeskewMinAngle..$DeskewMaxAngle"
Write-Host "Deskew step: $DeskewStep"
Write-Host "Enhancement: $($EnhanceScan.IsPresent)"
