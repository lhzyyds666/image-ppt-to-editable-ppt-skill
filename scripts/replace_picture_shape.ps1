param(
  [Parameter(Mandatory=$true)][string]$Pptx,
  [Parameter(Mandatory=$true)][int]$Slide,
  [Parameter(Mandatory=$true)][int]$ShapeId,
  [Parameter(Mandatory=$true)][string]$Image,
  [double]$LeftIn,
  [double]$TopIn,
  [double]$WidthIn,
  [double]$HeightIn,
  [string]$Name
)

$ErrorActionPreference = 'Stop'

$pptPath = (Resolve-Path -LiteralPath $Pptx).Path
$imagePath = (Resolve-Path -LiteralPath $Image).Path
$pptApp = $null
$pres = $null

try {
  $pptApp = New-Object -ComObject PowerPoint.Application
  $pres = $pptApp.Presentations.Open($pptPath, $false, $false, $false)
  $slideObj = $pres.Slides.Item($Slide)
  $old = $null
  foreach ($shape in @($slideObj.Shapes)) {
    if ($shape.Id -eq $ShapeId) {
      $old = $shape
      break
    }
  }
  if (-not $old) {
    throw "Shape id $ShapeId was not found on slide $Slide"
  }

  $left = if ($PSBoundParameters.ContainsKey('LeftIn')) { $LeftIn * 72 } else { $old.Left }
  $top = if ($PSBoundParameters.ContainsKey('TopIn')) { $TopIn * 72 } else { $old.Top }
  $width = if ($PSBoundParameters.ContainsKey('WidthIn')) { $WidthIn * 72 } else { $old.Width }
  $height = if ($PSBoundParameters.ContainsKey('HeightIn')) { $HeightIn * 72 } else { $old.Height }
  $z = $old.ZOrderPosition
  $oldName = $old.Name

  $old.Delete()
  $new = $slideObj.Shapes.AddPicture($imagePath, $false, $true, $left, $top, $width, $height)
  $new.Name = if ($Name) { $Name } else { "${oldName}_recut" }
  while ($new.ZOrderPosition -gt $z) { $new.ZOrder(3) | Out-Null }
  while ($new.ZOrderPosition -lt $z) { $new.ZOrder(2) | Out-Null }

  $pres.Save()
  "replaced shape $ShapeId -> new id $($new.Id), z=$($new.ZOrderPosition)"
}
finally {
  if ($pres) {
    $pres.Close()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
  }
  if ($pptApp) {
    $pptApp.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pptApp) | Out-Null
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
