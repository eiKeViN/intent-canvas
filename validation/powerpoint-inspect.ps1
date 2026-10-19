#requires -Version 5.1
<# Read-only probe for an already open PowerPoint presentation.
   Run with Windows PowerShell 5.1 (powershell.exe), not PowerShell Core.
   Does not open, edit, save, close, or quit the user's presentation/application.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$PresentationPath,
    [Parameter(Mandatory = $true)][string]$OutputPath,
    [switch]$IncludeContent
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
if ($PSVersionTable.PSEdition -ne 'Desktop') {
    throw 'Run this probe in Windows PowerShell 5.1 (powershell.exe).'
}
$expectedPath = (Resolve-Path -LiteralPath $PresentationPath).Path
$destination = [IO.Path]::GetFullPath($OutputPath)
if ([IO.Path]::GetExtension($destination) -ne '.json') { throw 'Output must be a new .json file.' }
if (Test-Path -LiteralPath $destination) { throw 'Output already exists; choose a new path.' }

function Get-TextHash([string]$Text) {
    $algorithm = [Security.Cryptography.SHA256]::Create()
    try {
        return ([BitConverter]::ToString($algorithm.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text)))).Replace('-', '').ToLowerInvariant()
    } finally { $algorithm.Dispose() }
}

function Read-Shape($Shape, [int]$SlideId, [int[]]$Parents) {
    $record = [ordered]@{ slideId = $SlideId; parents = @($Parents); errors = @() }
    try {
        $record.id = [int]$Shape.Id
        $record.type = [int]$Shape.Type
        $record.name = [string]$Shape.Name
        # Raw COM geometry; nested-group coordinates are not normalized to screen pixels.
        $record.bounds = @{ left = [double]$Shape.Left; top = [double]$Shape.Top;
                           width = [double]$Shape.Width; height = [double]$Shape.Height;
                           rotation = [double]$Shape.Rotation }
        if ([int]$Shape.HasTextFrame -eq -1) {
            $text = [string]$Shape.TextFrame.TextRange.Text
            $record.textLength = $text.Length
            $record.textSha256 = Get-TextHash $text
            if ($IncludeContent) { $record.text = $text }
        }
        if ($record.type -eq 6) { # msoGroup
            $children = @()
            for ($childIndex = 1; $childIndex -le $Shape.GroupItems.Count; $childIndex++) {
                $children += Read-Shape ($Shape.GroupItems.Item($childIndex)) $SlideId (@($Parents) + @($record.id))
            }
            $record.children = $children
        }
        if ([int]$Shape.HasChart -eq -1) {
            try {
                $series = @()
                $collection = $Shape.Chart.SeriesCollection()
                for ($seriesIndex = 1; $seriesIndex -le $collection.Count; $seriesIndex++) {
                    $current = $collection.Item($seriesIndex)
                    $values = @($current.Values)
                    $entry = [ordered]@{ index = $seriesIndex; pointCount = $values.Count;
                        valuesSha256 = Get-TextHash (ConvertTo-Json -InputObject $values -Compress -Depth 5) }
                    if ($IncludeContent) { $entry.values = $values; $entry.name = [string]$current.Name }
                    $series += [pscustomobject]$entry
                }
                $record.chartSeries = $series
            } catch { $record.errors += 'Chart series: ' + $_.Exception.Message }
        }
    } catch { $record.errors += $_.Exception.Message }
    return [pscustomobject]$record
}

$application = $null
try {
    $application = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    if ($application.Presentations.Count -eq 0) { throw 'Open the intended presentation in PowerPoint first.' }
    $presentation = $application.ActivePresentation
    if ([IO.Path]::GetFullPath([string]$presentation.FullName) -ne $expectedPath) {
        throw 'The active presentation differs from PresentationPath. Activate the intended file first.'
    }
    $slides = @()
    for ($slideIndex = 1; $slideIndex -le $presentation.Slides.Count; $slideIndex++) {
        $slide = $presentation.Slides.Item($slideIndex)
        $shapes = @()
        for ($shapeIndex = 1; $shapeIndex -le $slide.Shapes.Count; $shapeIndex++) {
            $shapes += Read-Shape ($slide.Shapes.Item($shapeIndex)) ([int]$slide.SlideID) @()
        }
        $slides += @{ index = $slideIndex; id = [int]$slide.SlideID; shapes = $shapes }
    }
    $selection = [ordered]@{ errors = @() }
    try {
        $selected = $application.ActiveWindow.Selection
        $selection.type = [int]$selected.Type
        $selection.slideId = [int]$application.ActiveWindow.View.Slide.SlideID
        if ($selection.type -in @(2, 3)) { # ppSelectionShapes / ppSelectionText
            $range = $selected.ShapeRange
            $selection.childRange = ([int]$selected.HasChildShapeRange -eq -1)
            if ($selection.childRange) { $range = $selected.ChildShapeRange }
            $selection.shapeIds = @(for ($i = 1; $i -le $range.Count; $i++) { [int]$range.Item($i).Id })
        }
        if ($selection.type -eq 3) {
            $selection.textStart = [int]$selected.TextRange.Start
            $selection.textLength = [int]$selected.TextRange.Length
        }
    } catch { $selection.errors += $_.Exception.Message }
    if ([IO.Path]::GetFullPath([string]$application.ActivePresentation.FullName) -ne $expectedPath) {
        throw 'The active presentation changed during capture; discard and retry.'
    }
    $result = [ordered]@{
        schema = 'intent-canvas-powerpoint-probe/1'; hostVersion = [string]$application.Version;
        capturedAtUtc = [DateTime]::UtcNow.ToString('o'); saved = ([int]$presentation.Saved -eq -1);
        slideSize = @{ width = [double]$presentation.PageSetup.SlideWidth; height = [double]$presentation.PageSetup.SlideHeight };
        selection = $selection; slides = $slides;
        scope = 'Read-only API sample, not complete object coverage. Errors must be reviewed. Snapshot is not atomic; do not edit during capture.'
    }
    $json = ConvertTo-Json -InputObject $result -Depth 100
    [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($destination))
    $stream = [IO.File]::Open($destination, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write)
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes($json)
        $stream.Write($bytes, 0, $bytes.Length)
    } finally { $stream.Dispose() }
    Write-Output ('Wrote read-only snapshot: ' + $destination)
} finally {
    # Do not call Quit or forcibly release proxies owned by the interactive host.
    if ($null -ne $application) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($application) }
}
