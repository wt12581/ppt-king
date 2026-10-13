param([Parameter(Mandatory=$true)][string]$Source,[Parameter(Mandatory=$true)][string]$Slides,[Parameter(Mandatory=$true)][string]$OutputDir)
$ErrorActionPreference='Stop'
$sourcePath=(Resolve-Path -LiteralPath $Source).Path;$out=[IO.Path]::GetFullPath($OutputDir)
if(Test-Path -LiteralPath $out){if((Get-ChildItem -LiteralPath $out).Count){throw 'Output directory must be empty'}}else{New-Item -ItemType Directory -Path $out -Force|Out-Null}
try{$app=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')}catch{$app=New-Object -ComObject PowerPoint.Application}
$pres=$app.Presentations.Open($sourcePath,-1,0,0);$images=@()
try{foreach($token in $Slides.Split(',')){$n=[int]$token;if($n -lt 1 -or $n -gt $pres.Slides.Count){throw 'Slide out of range'};$name='slide-'+$n.ToString('00')+'.png';$pres.Slides.Item($n).Export((Join-Path $out $name),'PNG',1600,900);$images+=@{file=$name;role=('source-slide-'+$n)}}}finally{$pres.Close()}
Copy-Item -LiteralPath $sourcePath -Destination (Join-Path $out 'reference.pptx')
ConvertTo-Json -InputObject @($images) -Depth 4|Set-Content -LiteralPath (Join-Path $out 'images.json') -Encoding utf8
Write-Output ('Rendered '+$images.Count+' reference pages')
