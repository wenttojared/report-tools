param(
  [string]$RepoRoot = (Resolve-Path "$PSScriptRoot\..").Path
)

$srcVba      = Join-Path $RepoRoot "src\vba"
$srcRibbon   = Join-Path $RepoRoot "src\ribbon\CustomUI14.xml"
$template    = Join-Path $RepoRoot "build\CoreTemplate.xlam"
$outDir      = Join-Path $RepoRoot "dist"
$outFile     = Join-Path $outDir   "ReportTools_Core.xlam"

function Update-ZipEntryText {
  param(
    [Parameter(Mandatory=$true)][string]$ZipPath,
    [Parameter(Mandatory=$true)][string]$EntryName,
    [Parameter(Mandatory=$true)][string]$Text
  )

  Add-Type -AssemblyName System.IO.Compression
  Add-Type -AssemblyName System.IO.Compression.FileSystem

  $fs = [System.IO.File]::Open($ZipPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
  try {
    $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Update, $true)

    # Remove existing entry if present
    $existing = $zip.GetEntry($EntryName)
    if ($existing) { $existing.Delete() }

    # Create new entry
    $entry = $zip.CreateEntry($EntryName, [System.IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    try {
      $writer = New-Object System.IO.StreamWriter($stream, (New-Object System.Text.UTF8Encoding($false)))
      $writer.Write($Text)
      $writer.Flush()
    } finally {
      $stream.Dispose()
    }

    $zip.Dispose()
  }
  finally {
    $fs.Dispose()
  }
}

New-Item -ItemType Directory -Force -Path $outDir | Out-Null
Copy-Item -Force $template $outFile

# --- Import VBA from src into the copied template ---
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
  $wb = $excel.Workbooks.Open($outFile)

  # Requires Trust Center setting:
  # "Trust access to the VBA project object model"
  $vbproj = $wb.VBProject

  # Remove existing modules/classes/forms (keep document modules)
  for ($i = $vbproj.VBComponents.Count; $i -ge 1; $i--) {
    $comp = $vbproj.VBComponents.Item($i)
    # 1=StdModule, 2=ClassModule, 3=MSForm, 100=Document
    if ($comp.Type -in 1,2,3) {
      $vbproj.VBComponents.Remove($comp)
    }
  }

  # Import repo modules
  Get-ChildItem $srcVba -File | ForEach-Object {
    $ext = $_.Extension.ToLowerInvariant()
    if ($ext -in ".bas",".cls",".frm") {
      $vbproj.VBComponents.Import($_.FullName) | Out-Null
    }
  }

  $wb.Save()
  $wb.Close($true)
}
finally {
  $excel.Quit() | Out-Null
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

# --- Inject RibbonX XML into the package ---
$ribbonXml = Get-Content -Raw -Encoding UTF8 $srcRibbon
Update-ZipEntryText -ZipPath $outFile -EntryName "customUI/customUI14.xml" -Text $ribbonXml

Write-Host "Built: $outFile"