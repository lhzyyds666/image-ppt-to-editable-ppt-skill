param(
  [Parameter(Mandatory=$true)][string]$Pptx,
  [Parameter(Mandatory=$true)][string]$OutDir,
  [int]$Width = 3840,
  [int]$Height = 2160,
  [string]$Title = 'PPT HTML Review'
)

$ErrorActionPreference = 'Stop'

function HtmlAttr([string]$Value) {
  return [System.Net.WebUtility]::HtmlEncode(($Value -replace "`r|`n", ' / '))
}

$pptPath = (Resolve-Path -LiteralPath $Pptx).Path
$outPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutDir)
$assetsDir = Join-Path $outPath 'assets'
$htmlPath = Join-Path $outPath 'review.html'
$manifestPath = Join-Path $outPath 'manifest.json'
New-Item -ItemType Directory -Force -Path $assetsDir | Out-Null

$version = Get-Date -Format 'yyyyMMddHHmmss'
$pptApp = $null
$pres = $null

try {
  $pptApp = New-Object -ComObject PowerPoint.Application
  $pres = $pptApp.Presentations.Open($pptPath, $true, $false, $false)
  $slides = @()

  for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $slide = $pres.Slides.Item($i)
    $preview = Join-Path $assetsDir ("slide_preview_{0}.png" -f $i)
    $slide.Export($preview, 'PNG', $Width, $Height)
    $items = @()

    foreach ($shape in @($slide.Shapes)) {
      $kind = 'shape'
      $text = ''
      try {
        if ($shape.Type -eq 13 -or $shape.Type -eq 11) { $kind = 'picture' }
        if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
          $kind = 'text'
          $text = $shape.TextFrame.TextRange.Text
        }
      } catch { }

      $items += [ordered]@{
        shape_id = [int]$shape.Id
        name = [string]$shape.Name
        kind = $kind
        x = [math]::Round([double]$shape.Left * 2, 3)
        y = [math]::Round([double]$shape.Top * 2, 3)
        w = [math]::Round([double]$shape.Width * 2, 3)
        h = [math]::Round([double]$shape.Height * 2, 3)
        z = [int]$shape.ZOrderPosition
        text = [string]$text
      }
    }

    $slides += [ordered]@{
      slide_number = [int]$i
      slide_id = [int]$slide.SlideID
      items = $items
    }
  }

  $manifest = [ordered]@{
    source_pptx = $pptPath
    mode = 'preview-background-with-shape-hotspots'
    generated_at = (Get-Date).ToString('s')
    slides = $slides
  }
  [System.IO.File]::WriteAllText($manifestPath, ($manifest | ConvertTo-Json -Depth 10), [System.Text.UTF8Encoding]::new($false))

  $sections = New-Object System.Collections.Generic.List[string]
  foreach ($s in $slides) {
    $slideNo = $s.slide_number
    $parts = New-Object System.Collections.Generic.List[string]
    $parts.Add(("<section class=`"slide`" id=`"slide-{0}`" data-slide-id=`"{1}`"><img class=`"slide-bg`" src=`"assets/slide_preview_{0}.png?v={2}`" alt=`"slide {0} preview`"><div class=`"slide-label`">{0}</div>" -f $slideNo, $s.slide_id, $version))
    foreach ($it in ($s.items | Sort-Object z)) {
      $label = if ($it.text) { $it.text } else { $it.kind }
      $parts.Add(("<div class=`"el hotspot {0}`" data-slide=`"{1}`" data-shape-id=`"{2}`" data-name=`"{3}`" data-kind=`"{0}`" data-label=`"{4}`" style=`"left:{5}px;top:{6}px;width:{7}px;height:{8}px;z-index:{9};`"></div>" -f (HtmlAttr $it.kind), $slideNo, $it.shape_id, (HtmlAttr $it.name), (HtmlAttr $label), $it.x, $it.y, $it.w, $it.h, $it.z))
    }
    $parts.Add('</section>')
    $sections.Add(($parts -join ''))
  }

  $nav = (1..$pres.Slides.Count | ForEach-Object { "<a href=`"#slide-$_`">$_</a>" }) -join ''
  $html = @"
<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="icon" href="data:,">
<title>$([System.Net.WebUtility]::HtmlEncode($Title))</title>
<style>
  :root { --ink:#061630; --accent:#f0440d; --blue:#1875ff; --bg:#f5f6f8; }
  * { box-sizing:border-box; }
  body { margin:0; background:var(--bg); color:var(--ink); font-family:Arial, "Microsoft YaHei", sans-serif; }
  .topbar { position:sticky; top:0; z-index:9999; display:flex; gap:12px; align-items:center; padding:10px 16px; background:rgba(255,255,255,.94); border-bottom:1px solid #d8dde7; }
  .topbar strong { font-size:15px; }
  .topbar a, .topbar button { border:1px solid #c9d0dc; background:#fff; color:var(--ink); border-radius:6px; padding:6px 10px; font-size:13px; text-decoration:none; cursor:pointer; }
  .deck { width:min(100vw, 1980px); margin:18px auto 48px; display:grid; gap:28px; justify-items:center; }
  .slide { position:relative; width:1920px; height:1080px; background:#fcfcfb; overflow:hidden; box-shadow:0 8px 28px rgba(5,22,48,.16); transform-origin:top center; }
  .slide-bg { position:absolute; inset:0; width:100%; height:100%; object-fit:fill; z-index:0; user-select:none; pointer-events:none; }
  .slide-label { position:absolute; right:12px; top:10px; z-index:9000; font:700 18px/1 Arial; color:#fff; background:rgba(5,22,48,.55); padding:5px 9px; border-radius:6px; }
  .el { position:absolute; background:rgba(24,117,255,0); z-index:10; }
  .el:hover { outline:3px solid rgba(240,68,13,.95); outline-offset:2px; background:rgba(240,68,13,.08); }
  .el:hover::before { content:"slide " attr(data-slide) " / shape " attr(data-shape-id) " / " attr(data-kind) " / " attr(data-label); position:absolute; left:0; bottom:100%; z-index:10000; transform:translateY(-6px); white-space:nowrap; font:700 14px/1 Arial; color:#fff; background:rgba(5,22,48,.94); padding:6px 8px; border-radius:5px; pointer-events:none; max-width:900px; overflow:hidden; text-overflow:ellipsis; }
  body.show-boxes .el { outline:2px solid rgba(24,117,255,.6); outline-offset:1px; background:rgba(24,117,255,.04); }
  body.show-boxes .el::after { content:"#" attr(data-shape-id); position:absolute; left:0; top:-20px; font:700 14px/1 Arial; color:#fff; background:var(--blue); padding:3px 5px; border-radius:4px; pointer-events:none; }
  @media (max-width: 1940px) { .slide { transform:scale(calc((100vw - 32px) / 1920)); margin-bottom:calc(-1080px + ((100vw - 32px) / 1920 * 1080px)); } }
</style>
</head>
<body>
  <div class="topbar"><strong>$([System.Net.WebUtility]::HtmlEncode($Title))</strong><button id="toggleBoxes" type="button">Toggle boxes</button>$nav</div>
  <main class="deck">$($sections -join '')</main>
<script>document.getElementById('toggleBoxes').addEventListener('click', () => document.body.classList.toggle('show-boxes'));</script>
</body>
</html>
"@
  [System.IO.File]::WriteAllText($htmlPath, $html, [System.Text.UTF8Encoding]::new($false))
  "wrote $htmlPath"
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
