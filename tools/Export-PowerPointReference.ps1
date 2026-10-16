param(
    [Parameter(Mandatory = $true)][string]$InputPptx,
    [Parameter(Mandatory = $true)][string]$OutputDirectory
)
$ErrorActionPreference = 'Stop'
$inputPath = (Resolve-Path -LiteralPath $InputPptx).Path
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$outputPath = [IO.Path]::GetFullPath($OutputDirectory)
$privateRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot 'plan/powerpoint/screenshots'))
if (-not $outputPath.StartsWith($privateRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'PowerPoint output must stay inside this checkout''s private plan/powerpoint/screenshots/ directory.'
}
if (Test-Path -LiteralPath $outputPath) { throw 'Use a new output directory; existing evidence is preserved.' }
if (@(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count -ne 0) {
    throw 'PowerPoint is already running. No existing session will be attached to or closed.'
}
New-Item -ItemType Directory -Path $outputPath -Force | Out-Null
$app = $null
$presentation = $null
$startedAt = Get-Date
$report = [ordered]@{ inputSha256 = (Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash; startedUtc = $startedAt.ToUniversalTime().ToString('o'); windowless = $true; completed = $false; slides = @() }
try {
    $app = New-Object -ComObject PowerPoint.Application
    $report.powerPointVersion = $app.Version
    if ($app.Presentations.Count -ne 0) { throw 'A presentation appeared during startup; preserving that session.' }
    $presentation = $app.Presentations.Open($inputPath, -1, 0, 0)
    $report.slideCount = $presentation.Slides.Count
    $report.widthPoints = $presentation.PageSetup.SlideWidth
    $report.heightPoints = $presentation.PageSetup.SlideHeight
    $width = [int][Math]::Round($report.widthPoints)
    $height = [int][Math]::Round($report.heightPoints)
    for ($i = 1; $i -le $presentation.Slides.Count; $i++) {
        $slide = $presentation.Slides.Item($i)
        try {
            $imagePath = Join-Path $outputPath ('slide-{0:D2}.png' -f $i)
            $slide.Export($imagePath, 'PNG', $width, $height)
            $report.slides += [ordered]@{ index = $i; shapes = $slide.Shapes.Count; imageSha256 = (Get-FileHash -LiteralPath $imagePath -Algorithm SHA256).Hash }
        } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($slide) }
    }
    $presentation.SaveCopyAs((Join-Path $outputPath 'powerpoint-normalized.pptx'), 24)
    $report.completed = $true
} catch {
    $report.error = $_.Exception.Message
    throw
} finally {
    if ($null -ne $presentation) {
        try { $presentation.Close() } catch { $report.closeError = $_.Exception.Message }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($presentation)
    }
    if ($null -ne $app) {
        # Never quit if another document was opened by the user during the test.
        try { if ($app.Presentations.Count -eq 0) { $app.Quit() } else { $report.sessionPreserved = $true } }
        catch { $report.quitError = $_.Exception.Message }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
    }
    $report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $outputPath 'reference.json') -Encoding utf8
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
$report | ConvertTo-Json -Depth 8
