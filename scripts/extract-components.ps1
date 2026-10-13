param([Parameter(Mandatory=$true)][string]$Source,[Parameter(Mandatory=$true)][string]$Plan,[Parameter(Mandatory=$true)][string]$OutputDir)
$ErrorActionPreference='Stop'
function Detach($shape){if($shape.Type -eq 6){for($j=1;$j -le $shape.GroupItems.Count;$j++){Detach $shape.GroupItems.Item($j)}}elseif($shape.Connector -eq -1){if($shape.ConnectorFormat.BeginConnected){$shape.ConnectorFormat.BeginDisconnect()};if($shape.ConnectorFormat.EndConnected){$shape.ConnectorFormat.EndDisconnect()}}}
function ReplaceText($shape,$map){if($shape.Type -eq 6){for($j=1;$j -le $shape.GroupItems.Count;$j++){ReplaceText $shape.GroupItems.Item($j) $map}}elseif($shape.HasTextFrame -eq -1 -and $shape.TextFrame2.HasText -eq -1){foreach($p in $map.PSObject.Properties){if($shape.TextFrame2.TextRange.Text -eq $p.Name){$shape.TextFrame2.TextRange.Text=$p.Value}}}}
try{$app=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')}catch{$app=New-Object -ComObject PowerPoint.Application}
$sourcePath=(Resolve-Path -LiteralPath $Source).Path;$planData=Get-Content -LiteralPath $Plan -Raw -Encoding utf8|ConvertFrom-Json
$out=[IO.Path]::GetFullPath($OutputDir);New-Item -ItemType Directory -Force -Path $out|Out-Null
$pres=$app.Presentations.Open($sourcePath,-1,0,0);$report=@()
try{foreach($item in $planData){
 if($item.id -notmatch '^[a-z0-9]+(-[a-z0-9]+)*$'){throw 'Invalid component ID'}
 $dest=Join-Path $out $item.id;if(Test-Path -LiteralPath $dest){throw ('Destination exists: '+$dest)};New-Item -ItemType Directory -Path $dest|Out-Null
 $p=$app.Presentations.Add(0)
 try{
  $p.PageSetup.SlideWidth=$pres.PageSetup.SlideWidth;$p.PageSetup.SlideHeight=$pres.PageSetup.SlideHeight
  $slide=$p.Slides.Add(1,12);$slide.FollowMasterBackground=0;$slide.Background.Fill.Solid();$slide.Background.Fill.ForeColor.RGB=16777215
  if($item.preview_background){$h=$item.preview_background.TrimStart('#');$slide.Background.Fill.ForeColor.RGB=[Convert]::ToInt32($h.Substring(0,2),16)+256*[Convert]::ToInt32($h.Substring(2,2),16)+65536*[Convert]::ToInt32($h.Substring(4,2),16)}
  $pres.Slides.Item($item.slide).Shapes.Range([object[]]@($item.shapes)).Copy();$range=$slide.Shapes.Paste()
  $g=if($range.Count -gt 1){$range.Group()}else{$range.Item(1)};$g.Name='component_'+$item.id;Detach $g
  if($item.replacements){ReplaceText $g $item.replacements}
  $g.LockAspectRatio=-1;$factor=[Math]::Min(($p.PageSetup.SlideWidth*.78)/$g.Width,($p.PageSetup.SlideHeight*.65)/$g.Height)
  $g.Width=$g.Width*$factor;$g.Left=($p.PageSetup.SlideWidth-$g.Width)/2;$g.Top=($p.PageSetup.SlideHeight-$g.Height)/2
  $g.AlternativeText='Reusable native component; source page '+$item.slide
  $p.SaveAs((Join-Path $dest 'component.pptx'),24);$slide.Export((Join-Path $dest 'preview.png'),'PNG',1000,563)
  $report+=@{id=$item.id;slide=$item.slide;shapes=$item.shapes;source=$sourcePath;group=$g.Name}
 }finally{$p.Close()}
}}finally{$pres.Close()}
$report|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $out 'extraction.json') -Encoding utf8
Write-Output ('Extracted '+$report.Count+' components')
