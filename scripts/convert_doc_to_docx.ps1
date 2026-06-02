param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputPath,

    [Parameter(Position = 1)]
    [string]$OutputPath,

    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-AbsolutePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return [System.IO.Path]::GetFullPath($Path)
}

$inputItem = Get-Item -LiteralPath $InputPath -ErrorAction Stop
if ($inputItem.PSIsContainer) {
    throw "Input path must point to a .doc file, not a directory: $InputPath"
}

if ($inputItem.Extension.ToLowerInvariant() -ne '.doc') {
    throw "Input file must have a .doc extension: $InputPath"
}

$resolvedInputPath = Get-AbsolutePath -Path $inputItem.FullName

if (-not $OutputPath) {
    $OutputPath = [System.IO.Path]::ChangeExtension($resolvedInputPath, '.docx')
}

$resolvedOutputPath = Get-AbsolutePath -Path $OutputPath
$outputDirectory = Split-Path -Parent $resolvedOutputPath

if ($outputDirectory -and -not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
}

if ((Test-Path -LiteralPath $resolvedOutputPath) -and -not $Force) {
    throw "Output file already exists. Use -Force to overwrite: $resolvedOutputPath"
}

$word = $null
$document = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    $document = $word.Documents.Open($resolvedInputPath, $false, $false, $false)
    $wdFormatXMLDocument = 16
    $document.SaveAs2($resolvedOutputPath, $wdFormatXMLDocument)

    Write-Host "Converted $resolvedInputPath -> $resolvedOutputPath"
}
catch {
    throw "Failed to convert .doc using Word automation. Ensure Microsoft Word is installed and available on this Windows machine. $($_.Exception.Message)"
}
finally {
    if ($document) {
        $document.Close($false) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null
    }

    if ($word) {
        $word.Quit() | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}