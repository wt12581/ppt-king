param([Parameter(Mandatory=$true)][string]$Component,[Parameter(Mandatory=$true)][string]$Target,[Parameter(Mandatory=$true)][int]$Slide,[Parameter(Mandatory=$true)][double]$X,[Parameter(Mandatory=$true)][double]$Y,[double]$Width=0,[string]$Font='',[string]$ColorMap='',[Parameter(Mandatory=$true)][string]$Output)
$ErrorActionPreference='Stop'
function RGB($h){$h=$h.TrimStart('#');return [Convert]::ToInt32($h.Substring(0,2),16)+256*[Convert]::ToInt32($h.Substring(2,2),16)+65536*[Convert]::ToInt32($h.Substring(4,2),16)}
function Adjust($s){if($s.Type -eq 6){for($j=1;$j -le $s.GroupItems.Count;$j++){Adjust $s.GroupItems.Item($j)}}else{if($s.Connector -eq -1){if($s.ConnectorFormat.BeginConnected){$s.ConnectorFormat.BeginDisconnect()};if($s.ConnectorFormat.EndConnected){$s.ConnectorFormat.EndDisconnect()}};if($Font -and $s.HasTextFrame -eq -1){$s.TextFrame2.TextRange.Font.Name=$Font;$s.TextFrame2.TextRange.Font.NameFarEast=$Font};if($colors){foreach($pair in $colors.PSObject.Properties){$old=RGB $pair.Name;$new=RGB $pair.Value;if($s.Fill.Visible -eq -1){if($s.Fill.ForeColor.RGB -eq $old){$s.Fill.ForeColor.RGB=$new};if($s.Fill.Type -eq 3 -and $s.Fill.BackColor.RGB -eq $old){$s.Fill.BackColor.RGB=$new}};if($s.Line.Visible -eq -1 -and $s.Line.ForeColor.RGB -eq $old){$s.Line.ForeColor.RGB=$new};if($s.HasTextFrame -eq -1 -and $s.TextFrame2.HasText -eq -1 -and $s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB -eq $old){$s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB=$new}}}}}
$out=[IO.Path]::GetFullPath($Output);if(Test-Path -LiteralPath $out){throw 'Refuse to overwrite output'}
$compPath=(Resolve-Path -LiteralPath $Component).Path;$targetPath=(Resolve-Path -LiteralPath $Target).Path
$colors=if($ColorMap){Get-Content -LiteralPath $ColorMap -Raw|ConvertFrom-Json}else{$null}
try{$app=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')}catch{$app=New-Object -ComObject PowerPoint.Application}
$cp=$null;$tp=$null
try{$cp=$app.Presentations.Open($compPath,-1,0,0);$tp=$app.Presentations.Open($targetPath,-1,0,0)
 if($cp.Slides.Count -ne 1 -or $Slide -lt 1 -or $Slide -gt $tp.Slides.Count){throw 'Invalid component or target slide'}
 $cp.Slides.Item(1).Shapes.Range([object[]]@(1..$cp.Slides.Item(1).Shapes.Count)).Copy();$range=$tp.Slides.Item($Slide).Shapes.Paste();$g=if($range.Count -gt 1){$range.Group()}else{$range.Item(1)}
 Adjust $g;$g.LockAspectRatio=-1;if($Width -gt 0){$g.Width=$Width};$g.Left=$X;$g.Top=$Y
 if($X -lt 0 -or $Y -lt 0 -or $X+$g.Width -gt $tp.PageSetup.SlideWidth -or $Y+$g.Height -gt $tp.PageSetup.SlideHeight){throw 'Component exceeds target canvas'}
 New-Item -ItemType Directory -Force -Path (Split-Path $out -Parent)|Out-Null;$tp.SaveAs($out,24);Write-Output $out
}finally{if($tp){$tp.Close()};if($cp){$cp.Close()}}
