param([Parameter(Mandatory=$true)][string]$Request,
      [ValidateSet('Excel','WPS')][string]$Engine='Excel')
$ErrorActionPreference='Stop'
$OutputEncoding=[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
$cfg=Get-Content -LiteralPath $Request -Raw -Encoding utf8 | ConvertFrom-Json
$dir=[string]$cfg.work_dir
$app=$null;$book=$null;$bootstrap=$null;$owned=$false
$existing=@(Get-Process -Name EXCEL,et -ErrorAction SilentlyContinue | ForEach-Object {$_.Id})
function Log([string]$s){[IO.File]::AppendAllText((Join-Path $dir 'native.log'),([DateTime]::UtcNow.ToString('o')+' '+$s+[Environment]::NewLine))}
if(-not ('ContractPrintNativeWindow' -as [type])){
 Add-Type -TypeDefinition @'
using System; using System.Runtime.InteropServices;
public static class ContractPrintNativeWindow {
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);
}
'@
}
function Open-Book {
 $b=$app.Workbooks.Open([string]$cfg.source,0,$false)
 if($b.Connections.Count -ne 0){$b.Close($false);throw 'Workbook has external data connections.'}
 foreach($s in $b.Worksheets){if($s.QueryTables.Count -ne 0){$b.Close($false);throw 'Workbook has QueryTables.'}}
 return $b
}
function Sheet-Info($s){
 $cells=@();$blocked=@();$excluded=@();$mergeSeen=@{};$minr=1048576;$minc=16384;$maxr=0;$maxc=0;$totalRow=0
 $cached=$null
 if($null -ne $cfg.sheet_meta){$cached=$cfg.sheet_meta.PSObject.Properties[[string]$s.Name].Value}
 if($null -ne $cached){
  foreach($p in $cached.cells){$cells+=,@([int]$p[0],[int]$p[1]);$minr=[Math]::Min($minr,[int]$p[0]);$minc=[Math]::Min($minc,[int]$p[1]);$maxr=[Math]::Max($maxr,[int]$p[0]);$maxc=[Math]::Max($maxc,[int]$p[1])}
  $blocked=@($cached.blocks);$totalRow=[int]$cached.total_row
  foreach($b in $cached.merge_ends){$maxr=[Math]::Max($maxr,[int]$b[0]);$maxc=[Math]::Max($maxc,[int]$b[1])}
 } else {
 foreach($kind in @(2,-4123)){
  $selection=$null
  try{$selection=$s.UsedRange.SpecialCells($kind)}catch{}
  if($null -eq $selection){continue}
  if($selection.CountLarge -gt 60000){throw ('Too many content cells: '+$s.Name)}
  foreach($c in $selection.Cells){
   if($c.EntireRow.Hidden -or $c.EntireColumn.Hidden){continue}
   $v=[string]$c.Value2
   if([string]::IsNullOrWhiteSpace($v)){continue}
   $rr=[int]$c.Row;$cc=[int]$c.Column
   $cells+=,@($rr,$cc)
   $minr=[Math]::Min($minr,$rr);$minc=[Math]::Min($minc,$cc)
   $maxr=[Math]::Max($maxr,$rr);$maxc=[Math]::Max($maxc,$cc)
   if($v -match '^\s*(?:GRAND\s+)?TOTAL\b|^\s*合计'){$totalRow=[Math]::Max($totalRow,$rr)}
   if($c.MergeCells){
    $m=$c.MergeArea;$key=[string]$m.Address()
    $maxr=[Math]::Max($maxr,[int]($m.Row+$m.Rows.Count-1));$maxc=[Math]::Max($maxc,[int]($m.Column+$m.Columns.Count-1))
    if(-not $mergeSeen.ContainsKey($key)){
     $mergeSeen[$key]=$true
     if($m.Rows.Count -gt 1){$blocked+=,@([int]$m.Row,[int]($m.Row+$m.Rows.Count-1))}
    }
   }
  }
 }
 }
 foreach($shape in $s.Shapes){
  if($shape.Visible -eq 0){continue}
  # Lines and decorative table rectangles can span many products; they are not
  # indivisible signatures. Pictures, charts and groups remain indivisible.
  if([int]$shape.Type -notin @(3,6,11,13)){continue}
  try{if(-not $shape.DrawingObject.PrintObject){continue}}catch{}
  $tl=$shape.TopLeftCell;$br=$shape.BottomRightCell
  $minr=[Math]::Min($minr,[int]$tl.Row);$minc=[Math]::Min($minc,[int]$tl.Column)
  $maxr=[Math]::Max($maxr,[int]$br.Row);$maxc=[Math]::Max($maxc,[int]$br.Column)
  $blocked+=,@([int]$tl.Row,[int]$br.Row)
  $cells+=,@([int]$tl.Row,[int]$tl.Column);$cells+=,@([int]$br.Row,[int]$br.Column)
 }
 if($maxr -eq 0){return $null}
 if($maxr -gt 5000 -or $maxc -gt 200){throw ('Oversized content range: '+$s.Name)}
 $area=[string]$s.PageSetup.PrintArea
 if($area){
  $range=$s.Range($area)
  if($range.Areas.Count -ne 1){throw ('Multiple print areas need review: '+$s.Name)}
  $r0=[int]$range.Row;$r1=[int]($range.Row+$range.Rows.Count-1)
  $c0=[int]$range.Column;$c1=[int]($range.Column+$range.Columns.Count-1)
  foreach($point in $cells){if($point[0] -lt $r0 -or $point[0] -gt $r1 -or $point[1] -lt $c0 -or $point[1] -gt $c1){
   $outside=$s.Cells.Item($point[0],$point[1]);$excluded+=@{cell=[string]$outside.Address();text=[string]$outside.Value2}
  }}
 }else{$range=$s.Range($s.Cells.Item($minr,$minc),$s.Cells.Item($maxr,$maxc))}
 return @{name=[string]$s.Name;area=[string]$range.Address();first_row=[int]$range.Row;
  last_row=[int]($range.Row+$range.Rows.Count-1);first_col=[int]$range.Column;
  last_col=[int]($range.Column+$range.Columns.Count-1);width=[double]$range.Width;height=[double]$range.Height;blocked=$blocked;total_row=$totalRow;excluded=$excluded}
}
function Is-Safe([int]$row,$blocks){
 foreach($b in $blocks){if($row -gt $b[0] -and $row -le $b[1]){return $false}}
 return $true
}
try{
 Log ('creating '+$Engine)
 $prog='Excel.Application';if($Engine -eq 'WPS'){$prog='ket.Application'}
 $app=New-Object -ComObject $prog
 [uint32]$nativeId=0
 [void][ContractPrintNativeWindow]::GetWindowThreadProcessId([intptr]$app.Hwnd,[ref]$nativeId)
 $owned=($nativeId -gt 0 -and $nativeId -notin $existing)
 if(-not $owned){throw 'Office reused a user process; no workbook was opened or changed.'}
 if($app.Workbooks.Count -ne 0){throw 'Unexpected user workbook in new Office process.'}
 @{pid=$nativeId;started=(Get-Process -Id $nativeId).StartTime.ToUniversalTime().ToString('o');engine=$Engine;owned=$true} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $dir 'session.json') -Encoding utf8
 $app.Visible=$false;$app.DisplayAlerts=$false;$app.AutomationSecurity=3
 $app.EnableEvents=$false;$app.AskToUpdateLinks=$false
 $bootstrap=$app.Workbooks.Add();$app.Calculation=-4135;$app.CalculateBeforeSave=$false
 $device=(Get-ItemProperty -LiteralPath 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Devices' -ErrorAction SilentlyContinue).'Microsoft Print to PDF'
 if($device -match ',([^,]+)$'){
  $port=$Matches[1]
  try{$connector=' on ';if([string]$app.ActivePrinter -match ' 在 '){$connector=' 在 '};$app.ActivePrinter='Microsoft Print to PDF'+$connector+$port}catch{Log ('printer: '+$_.Exception.Message)}
 }
 $book=Open-Book
 if($cfg.export_only){
  $book.ExportAsFixedFormat(0,[string]$cfg.pdf)
  $book.Close($false);$book=$null
  exit 0
 }
 $sheetInfo=@()
 foreach($s in $book.Worksheets){
  if($s.Visible -ne -1){continue}
  if($s.ProtectContents){throw ('Protected sheet requires review: '+$s.Name)}
  $info=Sheet-Info $s
  if($null -ne $info){$sheetInfo+=,$info}
 }
 if($sheetInfo.Count -eq 0){throw 'No visible printable content.'}
 Log 'baseline export'
 $book.ExportAsFixedFormat(0,(Join-Path $dir 'baseline.pdf'))
 $book.Close($false);$book=$null
 $candidates=@();$seenSettings=@{}
 foreach($layout in @(@{orientation=1;margin=28.35;label='o1'},@{orientation=1;margin=17.01;label='o1m6'},@{orientation=2;margin=28.35;label='o2'})){
  $orientation=[int]$layout.orientation;$margin=[double]$layout.margin
  foreach($target in @(1,2,0)){
   $id=($layout.label+'p'+$target)
   Log ('candidate '+$id)
   $book=Open-Book;$settings=@();$unsafe=@();$verify=@()
   try {
   foreach($info in $sheetInfo){
    $s=$book.Worksheets.Item($info.name);$ps=$s.PageSetup
    $ps.PrintArea=$info.area;$ps.PaperSize=9;$ps.Orientation=$orientation
    $ps.LeftMargin=$margin;$ps.RightMargin=$margin;$ps.TopMargin=$margin;$ps.BottomMargin=$margin
    $ps.CenterHorizontally=$true
    $paperW=595.28;$paperH=841.89;if($orientation -eq 2){$paperW=841.89;$paperH=595.28}
    # Extra room for native header/footer/rounding; actual PDF is checked later.
    $usableW=$paperW-2*$margin;$usableH=$paperH-2*$margin-13.3
    $zoom=[Math]::Min(100,[Math]::Floor(100*$usableW/[Math]::Max(1,$info.width)))
    if($target -gt 0){$zoom=[Math]::Min($zoom,[Math]::Floor(100*$usableH*$target/[Math]::Max(1,$info.height)))}
    if($zoom -lt 10){throw ('Required scale is below Office limit: '+$info.name)}
    $s.ResetAllPageBreaks();$ps.FitToPagesWide=1;$ps.FitToPagesTall=$false;$ps.Zoom=[int]$zoom
    $capacity=$usableH*100/$zoom
    $blocks=@($info.blocked)
    if($info.total_row -gt $info.first_row -and $info.total_row -ge ($info.first_row+($info.last_row-$info.first_row)*0.5)){
     $tail=$s.Range($s.Cells.Item($info.total_row,1),$s.Cells.Item($info.last_row,1))
     if($tail.Height -le $capacity*0.8){$blocks+=,@([int]$info.total_row,[int]$info.last_row)}
    }
    $repeat=0
    if([string]$ps.PrintTitleRows){try{$repeat=[double]$s.Range([string]$ps.PrintTitleRows).Height}catch{}}
    $first=$info.first_row;$breaks=@()
    while($first -le $info.last_row){
     $height=0;$end=$first;$available=$capacity
     if($first -gt $info.first_row){$available-=$repeat}
     while($end -le $info.last_row){
      $rh=[double]$s.Rows.Item($end).Height
      if($height+$rh -gt $available){break}
      $height+=$rh;$end++
     }
     if($end -gt $info.last_row){break}
     $cut=$end
     while($cut -gt $first -and -not (Is-Safe $cut $blocks)){$cut--}
     if($cut -le $first){$unsafe+=($info.name+': unbreakable row block');break}
     [void]$s.HPageBreaks.Add($s.Cells.Item($cut,$info.first_col))
     $breaks+=,$cut;$first=$cut
     if($breaks.Count -gt 100){throw 'Excessive page count.'}
    }
    $settings+=@{sheet=$info.name;area=$info.area;scale=[int]$zoom;orientation=$orientation;margin_mm=[Math]::Round($margin/2.835,1);breaks=$breaks}
    $verify+=@{info=$info;blocks=$blocks}
   }
   $ext=[IO.Path]::GetExtension([string]$cfg.source)
   $signature=($settings | ConvertTo-Json -Depth 10 -Compress)
   if($seenSettings.ContainsKey($signature)){Log ('duplicate settings '+$id);continue}
   $seenSettings[$signature]=$true
   $copy=Join-Path $dir ($id+$ext);$pdf=Join-Path $dir ($id+'.pdf')
   $book.SaveCopyAs($copy);$book.ExportAsFixedFormat(0,$pdf)
   # Export forces repagination. Before export Excel may report breaks cached
   # from the original printer/scale, even when the PDF is a single page.
   foreach($v in $verify){
    $info=$v.info;$blocks=$v.blocks;$s=$book.Worksheets.Item($info.name)
    try{
     $s.DisplayPageBreaks=$true
     foreach($b in $s.HPageBreaks){$r=[int]$b.Location.Row;if($r -gt $info.first_row -and $r -le $info.last_row -and -not (Is-Safe $r $blocks)){$unsafe+=($info.name+': native break '+$r)}}
     foreach($b in $s.VPageBreaks){$c=[int]$b.Location.Column;if($c -gt $info.first_col -and $c -le $info.last_col){$unsafe+=($info.name+': horizontal overflow at column '+$c)}}
    }catch{throw ('Unable to verify native page breaks: '+$_.Exception.Message)}
   }
   $candidates+=@{id=$id;excel=$copy;pdf=$pdf;settings=$settings;unsafe=$unsafe}
   } catch {$candidates+=@{id=$id;error=$_.Exception.Message};Log ('candidate failed: '+$_.Exception.Message)}
   finally {$book.Close($false);$book=$null}
  }
 }
 @{engine=$Engine;sheets=$sheetInfo;candidates=$candidates} | ConvertTo-Json -Depth 16 | Set-Content -LiteralPath (Join-Path $dir 'native.json') -Encoding utf8
 Log 'done'
}finally{
 if($null -ne $book){$book.Close($false)}
 if($null -ne $bootstrap){$bootstrap.Close($false)}
 if($null -ne $app){if($owned -and $app.Workbooks.Count -eq 0){$app.Quit()};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
