param(
    [Parameter(Mandatory=$true)][string]$PlanPath,
    [switch]$ExportPreview,
    [ValidateSet('Excel','WPS')][string]$Engine = 'Excel'
)
$ErrorActionPreference = 'Stop'
$culture = [cultureinfo]::InvariantCulture
function Require($condition, [string]$message) { if (-not $condition) { throw $message } }
function Iso-Date([string]$text) { return [datetime]::ParseExact($text, 'yyyy-MM-dd', $culture) }
function Hash-File([string]$path) { return (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash }
function Text-Date($datePlan, [datetime]$day) {
    Require (-not [string]::IsNullOrWhiteSpace($datePlan.format)) 'Text date requires its original format.'
    $pattern=[string]$datePlan.format
    if ($pattern.Contains('{ordinal}')) {
        $ordinal='th'
        if (($day.Day % 100) -notin 11,12,13) {
            switch ($day.Day % 10) {1 {$ordinal='st'} 2 {$ordinal='nd'} 3 {$ordinal='rd'}}
        }
        Require ([string]$datePlan.ordinal_case -in '', 'lower', 'upper') 'Unknown ordinal_case.'
        if ($datePlan.ordinal_case -eq 'upper') {$ordinal=$ordinal.ToUpperInvariant()}
        $pattern=$pattern.Replace('{ordinal}', "'$ordinal'")
    }
    $value = $day.ToString($pattern, $culture)
    switch ([string]$datePlan.text_case) {
        'upper' {$value=$value.ToUpperInvariant()}
        'lower' {$value=$value.ToLowerInvariant()}
        'none' {}
        '' {}
        default {throw 'Unknown text_case.'}
    }
    return ([string]$datePlan.prefix + $value + [string]$datePlan.suffix)
}
function Target-Cell($book, $spec) {
    Require ($spec.cell -cmatch '^[A-Z]{1,3}[1-9][0-9]*$') 'Cell must be one A1 address.'
    $cell = $book.Worksheets.Item([string]$spec.sheet).Range([string]$spec.cell)
    Require ($cell.Count -eq 1) 'Target must be a single cell.'
    Require (-not $cell.HasFormula) 'Refusing to overwrite a formula; preserve its logic or mark pending.'
    if ($cell.MergeCells) {
        Require ($cell.MergeArea.Cells.Item(1,1).Address() -eq $cell.Address()) 'Use the top-left merged cell.'
    }
    return ,$cell
}
function Same-Value($actual, $expected) {
    if ($null -eq $actual -or $null -eq $expected) {return ($null -eq $actual -and $null -eq $expected)}
    if ($actual -is [string] -or $expected -is [string]) {
        return ($actual -is [string] -and $expected -is [string] -and $actual -ceq $expected)
    }
    return ([double]$actual -eq [double]$expected)
}
function Range-Xml($sheet) {
    $used = $sheet.UsedRange
    $lastRow = $used.Row + $used.Rows.Count - 1
    $lastCol = $used.Column + $used.Columns.Count - 1
    Require (($lastRow * $lastCol) -le 200000) 'Used range too large for the preservation check; inspect excess formatting first.'
    return [string]$sheet.Range($sheet.Cells.Item(1,1), $sheet.Cells.Item($lastRow,$lastCol)).Value(11)
}
function Find-XmlCell($document, [int]$wantRow, [int]$wantCol) {
    $ns = [xml.xmlnamespacemanager]::new($document.NameTable)
    $uri='urn:schemas-microsoft-com:office:spreadsheet'
    $ns.AddNamespace('s',$uri)
    $rowIndex=0
    foreach ($row in $document.SelectNodes('//s:Table/s:Row',$ns)) {
        $explicit=$row.GetAttribute('Index',$uri)
        if ($explicit) {$rowIndex=[int]$explicit} else {$rowIndex++}
        if ($rowIndex -ne $wantRow) {continue}
        $colIndex=0
        foreach ($cell in $row.SelectNodes('s:Cell',$ns)) {
            $explicit=$cell.GetAttribute('Index',$uri)
            if ($explicit) {$colIndex=[int]$explicit} else {$colIndex++}
            if ($colIndex -eq $wantCol) {return ,$cell}
            $merged=$cell.GetAttribute('MergeAcross',$uri)
            if ($merged) {$colIndex += [int]$merged}
        }
    }
    throw 'Planned cell missing from XML snapshot.'
}
function Normalize-StyleIds($doc) {
    # Excel can renumber identical style records on save. Compare their actual
    # definitions and cell links, rather than incidental sNN storage IDs.
    $uri='urn:schemas-microsoft-com:office:spreadsheet'
    $ns=[xml.xmlnamespacemanager]::new($doc.NameTable);$ns.AddNamespace('s',$uri)
    $map=@{};$hasher=[security.cryptography.sha256]::Create()
    $nodes=@($doc.SelectNodes('//s:Styles/s:Style',$ns))
    $known=@{};foreach($style in $nodes){$known[$style.GetAttribute('ID',$uri)]=$true}
    try {
        while($map.Count -lt $nodes.Count) {
            $countBefore=$map.Count
            foreach($style in $nodes) {
                $id=$style.GetAttribute('ID',$uri)
                if($map.ContainsKey($id)){continue}
                $parent=$style.GetAttribute('Parent',$uri)
                if($parent -and $known.ContainsKey($parent) -and -not $map.ContainsKey($parent)){continue}
                $copy=$style.CloneNode($true);$copy.RemoveAttribute('ID',$uri)
                if($parent -and $map.ContainsKey($parent)){$copy.SetAttribute('Parent',$uri,$map[$parent]) | Out-Null}
                $bytes=[text.encoding]::UTF8.GetBytes($copy.OuterXml)
                $map[$id]='style-'+[convert]::ToHexString($hasher.ComputeHash($bytes))
            }
            Require ($map.Count -gt $countBefore) 'Cannot resolve style inheritance.'
        }
    } finally {$hasher.Dispose()}
    foreach($node in $doc.SelectNodes('//*[@s:StyleID]',$ns)) {
        $id=$node.GetAttribute('StyleID',$uri)
        if($map.ContainsKey($id)){$node.SetAttribute('StyleID',$uri,$map[$id]) | Out-Null}
    }
    $styles=$doc.SelectSingleNode('//s:Styles',$ns)
    if($null -ne $styles) {
        foreach($style in $styles.ChildNodes) {
            $id=$style.GetAttribute('ID',$uri)
            if($map.ContainsKey($id)){$style.SetAttribute('ID',$uri,$map[$id]) | Out-Null}
            $parent=$style.GetAttribute('Parent',$uri)
            if($parent -and $map.ContainsKey($parent)){$style.SetAttribute('Parent',$uri,$map[$parent]) | Out-Null}
        }
        $ordered=@($styles.ChildNodes | Sort-Object {$_.GetAttribute('ID',$uri)})
        foreach($style in $ordered){[void]$styles.AppendChild($style)}
    }
}
function Snapshot($book, $changes) {
    $sheets = @()
    foreach ($sheet in $book.Worksheets) {
        Write-Verbose ('Snapshot sheet XML: ' + $sheet.Name)
        $doc=[xml.xmldocument]::new()
        $doc.PreserveWhitespace=$false
        $doc.LoadXml((Range-Xml $sheet))
        foreach ($change in $changes) {
            if ($change.sheet -cne $sheet.Name) {continue}
            if ($Engine -eq 'WPS') {continue}
            $cell=$sheet.Range($change.cell)
            $node=Find-XmlCell $doc $cell.Row $cell.Column
            $data=@($node.ChildNodes | Where-Object {$_.LocalName -eq 'Data'})
            if ($data.Count -eq 1) {
                Require (@($data[0].ChildNodes | Where-Object {$_.NodeType -eq 'Element'}).Count -eq 0) 'Rich-text target requires separate formatting-preserving handling.'
                $data[0].InnerText='__AUTHORIZED_VALUE__'
            } else {Require ($Engine -eq 'WPS') 'Expected one data node in planned cell.'}
        }
        # WPS Value(11) omits cell data and some applied-style links. Compare
        # actual native cell properties as well; never trust that XML alone.
        $cells=@(); $rows=@(); $cols=@()
        if ($Engine -eq 'WPS') {
            $used=$sheet.UsedRange
            $lastRow=$used.Row+$used.Rows.Count-1; $lastCol=$used.Column+$used.Columns.Count-1
            Require (($lastRow*$lastCol) -le 20000) 'WPS range too large for per-cell preservation checks.'
            $mask=@{}
            foreach ($change in $changes) {if($change.sheet -ceq $sheet.Name){$mask[$change.cell]=$true}}
            for($r=1;$r -le $lastRow;$r++) {
                $rows+=@{height=$sheet.Rows.Item($r).RowHeight; hidden=$sheet.Rows.Item($r).Hidden}
                for($c=1;$c -le $lastCol;$c++) {
                    $cell=$sheet.Cells.Item($r,$c); $address=$cell.Address($false,$false)
                    $value=$cell.Value2; $formula=$cell.Formula
                    if($mask.ContainsKey($address)) {$value='__AUTHORIZED_VALUE__';$formula='__AUTHORIZED_VALUE__'}
                    $font=$cell.Font; $fill=$cell.Interior
                    $borders=@()
                    foreach($edge in @(5,6,7,8,9,10)) {
                        $border=$cell.Borders.Item($edge)
                        $borders+=@{line=$border.LineStyle; weight=$border.Weight; color=$border.Color}
                    }
                    $merge=''; if($cell.MergeCells){$merge=$cell.MergeArea.Address()}
                    $cells += [ordered]@{
                        address=$address;value=$value;formula=$formula;format=$cell.NumberFormat;style=$cell.Style
                        font=@{name=$font.Name;size=$font.Size;bold=$font.Bold;italic=$font.Italic;color=$font.Color;underline=$font.Underline;strike=$font.Strikethrough}
                        fill=@{color=$fill.Color;pattern=$fill.Pattern};borders=$borders;merge=$merge
                        horizontal=$cell.HorizontalAlignment;vertical=$cell.VerticalAlignment;wrap=$cell.WrapText;orientation=$cell.Orientation
                    }
                }
            }
            for($c=1;$c -le $lastCol;$c++) {$cols+=@{width=$sheet.Columns.Item($c).ColumnWidth;hidden=$sheet.Columns.Item($c).Hidden}}
        }
        # PageSetup getters can block on an unavailable network printer in both
        # Excel and WPS. Print-area/title names are covered below; page layout
        # requires the separate rendered/native preview check.
        Write-Verbose ('Snapshot shapes: ' + $sheet.Name)
        $shapes=@()
        foreach ($shape in $sheet.Shapes) {
            $text=''
            try {if ($shape.TextFrame2.HasText) {$text=$shape.TextFrame2.TextRange.Text}} catch {}
            $shapes += [ordered]@{name=$shape.Name; type=$shape.Type; left=$shape.Left; top=$shape.Top; width=$shape.Width; height=$shape.Height; rotation=$shape.Rotation; placement=$shape.Placement; text=$text}
        }
        $snapshotXml=$doc.OuterXml
        if ($Engine -eq 'WPS') {$snapshotXml=''}
        else {Normalize-StyleIds $doc; $snapshotXml=$doc.OuterXml}
        $sheets += [ordered]@{name=$sheet.Name; visible=$sheet.Visible; xml=$snapshotXml; cells=$cells; rows=$rows; columns=$cols; shapes=$shapes}
    }
    Write-Verbose 'Snapshot workbook names.'
    $names=@()
    foreach ($name in $book.Names) {$names += [ordered]@{name=$name.Name; refers_to=$name.RefersTo; visible=$name.Visible}}
    return ([ordered]@{date1904=$book.Date1904; sheets=$sheets; names=$names} | ConvertTo-Json -Depth 30 -Compress)
}

$plan=Get-Content -LiteralPath $PlanPath -Raw -Encoding utf8 | ConvertFrom-Json
Require ($plan.association.status -ceq 'confirmed') 'Association is not confirmed.'
foreach ($record in @($plan.association.order_confirmation, $plan.association.purchase_contract)) {
    Require ($null -ne $record) 'Order confirmation and purchase contract evidence are required.'
    Require ([IO.Path]::IsPathFullyQualified([string]$record.path)) 'Input evidence requires an absolute path.'
    Require (Test-Path -LiteralPath $record.path -PathType Leaf) 'Evidence input file not found.'
    Require (-not [string]::IsNullOrWhiteSpace($record.locator)) 'Evidence locator is required.'
    Require (-not [string]::IsNullOrWhiteSpace($record.evidence)) 'Association evidence is required.'
}
Require ([IO.Path]::IsPathFullyQualified([string]$plan.source)) 'Source must be an absolute path.'
Require ([IO.Path]::IsPathFullyQualified([string]$plan.output)) 'Output must be an absolute path.'
$source=[IO.Path]::GetFullPath([string]$plan.source)
$output=[IO.Path]::GetFullPath([string]$plan.output)
Require (Test-Path -LiteralPath $source -PathType Leaf) 'Source PI not found.'
Require ($source -ine $output) 'Cannot overwrite source PI.'
$extension=[IO.Path]::GetExtension($source).ToLowerInvariant()
Require ($extension -in @('.xls','.xlsx','.xlsm')) 'Only native xls/xlsx/xlsm supported by this helper.'
Require ($extension -eq [IO.Path]::GetExtension($output).ToLowerInvariant()) 'Output extension must match source.'
foreach ($path in @($output, ($output+'.audit.json'))) {Require (-not (Test-Path -LiteralPath $path)) 'Output already exists; choose a new name.'}
$sourceHash=Hash-File $source
Require ($sourceHash -ieq $plan.source_sha256) 'Source PI hash differs from inspected input.'
Require ($null -ne $plan.date) 'A verified PI document date is required; otherwise mark pending.'
$piDate=Iso-Date $plan.date.pi_date
$purchaseDate=Iso-Date $plan.association.purchase_contract.signing_date
$shouldChangeDate=$piDate -gt $purchaseDate
$inputHashes=@(
    @{role='PI'; path=$source; sha256=$sourceHash},
    @{role='order_confirmation'; path=$plan.association.order_confirmation.path; sha256=(Hash-File $plan.association.order_confirmation.path)},
    @{role='purchase_contract'; path=$plan.association.purchase_contract.path; sha256=(Hash-File $plan.association.purchase_contract.path)}
)
$outputDir=[IO.Path]::GetDirectoryName($output)
[void][IO.Directory]::CreateDirectory($outputDir)
$temp=Join-Path $outputDir ('.contract-working-' + [guid]::NewGuid().ToString('N') + $extension)
$excel=$null; $book=$null; $bootstrap=$null; $ownsApplication=$false
$tempAudit=$temp+'.audit.json'; $published=$false; $auditPublished=$false
$changes=[System.Collections.Generic.List[object]]::new()
$previews=@(); $previewErrors=@()
try {
    Write-Verbose ('Starting native spreadsheet engine: ' + $Engine)
    $progid='Excel.Application'
    if ($Engine -eq 'WPS') {$progid='ket.Application'}
    $existingOfficeIds=@(Get-Process -Name EXCEL,et -ErrorAction SilentlyContinue | ForEach-Object {$_.Id})
    if (-not ('ContractNativeWindow' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class ContractNativeWindow {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@
    }
    $excel=New-Object -ComObject $progid
    [uint32]$nativeProcessId=0
    [void][ContractNativeWindow]::GetWindowThreadProcessId([intptr]$excel.Hwnd,[ref]$nativeProcessId)
    $ownsApplication=($nativeProcessId -gt 0 -and $nativeProcessId -notin $existingOfficeIds)
    Require $ownsApplication 'Native engine reused an existing user process; refusing to hide, alter or quit that session.'
    Require ($excel.Workbooks.Count -eq 0) 'New native instance contains unexpected workbooks.'
    $excel.Visible=$false; $excel.DisplayAlerts=$false
    $excel.AutomationSecurity=3; $excel.EnableEvents=$false; $excel.AskToUpdateLinks=$false
    # Calculation cannot reliably be set with zero open workbooks. Establish
    # manual mode in a disposable blank book BEFORE opening any business input.
    $bootstrap=$excel.Workbooks.Add()
    $excel.Calculation=-4135
    $excel.CalculateBeforeSave=$false
    # Some native spreadsheet engines can update legacy OLE metadata even on
    # read-only opens. Never hand a business original to the native engine.
    [IO.File]::Copy($source,$temp,$false)
    Write-Verbose 'Opening isolated PI copy read-only.'
    $book=$excel.Workbooks.Open($temp,0,$true)
    $bootstrap.Close($false); $bootstrap=$null
    Write-Verbose 'Validating original target cells.'
    Require ($book.Connections.Count -eq 0) 'Workbook has data connections; inspect refresh behavior before using this helper.'
    foreach ($sheet in $book.Worksheets) {Require ($sheet.QueryTables.Count -eq 0) 'Workbook has query tables; separate refresh-preserving handling required.'}
    $targetKeys=@{}
    foreach ($targetSpec in @($plan.title, $plan.invoice_label, $plan.date)) {
        if ($null -eq $targetSpec) {continue}
        $targetKey=([string]$targetSpec.sheet).ToUpperInvariant()+'!'+[string]$targetSpec.cell
        Require (-not $targetKeys.ContainsKey($targetKey)) 'Title, invoice label and date must use different cells; combined fields require separate handling.'
        $targetKeys[$targetKey]=$true
    }
    $dateCell=Target-Cell $book $plan.date
    Require (Same-Value $dateCell.Value2 $plan.date.expected) 'PI date cell differs from expected original value.'
    $dateFormat=[string]$dateCell.NumberFormat
    if ($dateCell.Value2 -is [string]) {
        Require ((Text-Date $plan.date $piDate) -ceq $dateCell.Value2) 'Original text date cannot be reproduced with the supplied format.'
        $dateTarget=Text-Date $plan.date $purchaseDate
    } else {
        Require ($null -ne $dateCell.Value2) 'PI date is empty.'
        $epochOffset=0
        if ($book.Date1904) {$epochOffset=1462}
        $actualDate=[datetime]::FromOADate([double]$dateCell.Value2+$epochOffset)
        Require ($actualDate.Date -eq $piDate) 'PI numeric value does not match interpreted date.'
        $dateTarget=$purchaseDate.ToOADate()-$epochOffset+$actualDate.TimeOfDay.TotalDays
    }
    if ($null -ne $plan.title) {
        $titleCell=Target-Cell $book $plan.title
        Require (Same-Value $titleCell.Value2 $plan.title.expected) 'Title differs from expected original value.'
        Require ([string]$titleCell.Value2 -match '^\s*(PRO\s*[-–]?\s*FORMA\s*(INVOICE)?|INVOICE|PI|CONTRACT)\s*$') 'Target is not a standalone document title.'
        $titleTarget=[regex]::Replace([string]$titleCell.Value2, '\S(?:[\s\S]*\S)?', 'CONTRACT')
        if ($titleTarget -cne $titleCell.Value2) {
            $changes.Add(@{kind='title'; sheet=[string]$plan.title.sheet; cell=[string]$plan.title.cell; before=$titleCell.Value2; after=$titleTarget})
        }
    }
    if ($shouldChangeDate) {
        $changes.Add(@{kind='date'; sheet=[string]$plan.date.sheet; cell=[string]$plan.date.cell; before=$dateCell.Value2; after=$dateTarget})
    }
    if ($null -ne $plan.invoice_label) {
        $labelCell=Target-Cell $book $plan.invoice_label
        Require (Same-Value $labelCell.Value2 $plan.invoice_label.expected) 'Invoice label differs from expected original value.'
        Require ($labelCell.Value2 -is [string]) 'Invoice number label must be text.'
        $labelMatch=[regex]::Match([string]$labelCell.Value2, '^(\s*)(?:Invoice[ \t]+)?No\.', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
        Require $labelMatch.Success 'Target is not an Invoice No. / No. label.'
        $labelTarget=$labelMatch.Groups[1].Value+'No.'+([string]$labelCell.Value2).Substring($labelMatch.Length)
        if ($labelTarget -cne $labelCell.Value2) {
            $changes.Add(@{kind='invoice_label'; sheet=[string]$plan.invoice_label.sheet; cell=[string]$plan.invoice_label.cell; before=$labelCell.Value2; after=$labelTarget})
        }
    }
    foreach($change in $changes) {
        if($change.before -isnot [string]){continue}
        $target=$book.Worksheets.Item($change.sheet).Range($change.cell)
        $firstFont=$null
        for($i=1;$i -le $change.before.Length;$i++) {
            $charFont=$target.Characters($i,1).Font
            $fontKey=@($charFont.Name,$charFont.Size,$charFont.Bold,$charFont.Italic,$charFont.Color,$charFont.Underline,$charFont.Strikethrough) | ConvertTo-Json -Compress
            if($null -eq $firstFont){$firstFont=$fontKey}
            Require ($fontKey -ceq $firstFont) 'Mixed character formatting requires separate preservation handling.'
        }
    }
    Write-Verbose 'Taking original content and layout snapshot.'
    $before=Snapshot $book $changes
    if ($ExportPreview) {
        try {
            $path=$output+'.before.pdf'
            Require (-not (Test-Path -LiteralPath $path)) 'Preview output already exists.'
            $book.ExportAsFixedFormat(0,$path)
            $previews += @{role='before'; path=$path}
        } catch {$previewErrors += $_.Exception.Message}
    }
    Write-Verbose 'Closing source and creating native copy.'
    $book.Close($false); $book=$null
    $book=$excel.Workbooks.Open($temp,0,$false)
    Write-Verbose 'Applying planned edits.'
    Require (-not $book.ReadOnly) 'Output copy is read-only.'
    foreach ($change in $changes) {
        $cell=Target-Cell $book $change
        Require (Same-Value $cell.Value2 $change.before) 'Source copy changed before applying plan.'
        # Separate COM setter sites: WPS otherwise caches the first string
        # assignment's dispatch signature and rejects a later numeric date.
        if ($change.after -is [string]) {$cell.Value2=[string]$change.after}
        else {$cell.Value2=[double]$change.after}
    }
    if ($changes.Count -gt 0) {$book.Save()}
    $book.Close($false); $book=$null
    Write-Verbose 'Reopening saved copy for verification.'
    $book=$excel.Workbooks.Open($temp,0,$true)
    foreach ($change in $changes) {
        $actual=$book.Worksheets.Item($change.sheet).Range($change.cell).Value2
        Require (Same-Value $actual $change.after) 'Saved target value does not match expected value.'
    }
    $savedDate=$book.Worksheets.Item([string]$plan.date.sheet).Range([string]$plan.date.cell)
    Require ([string]$savedDate.NumberFormat -ceq $dateFormat) 'Date display format changed.'
    $after=Snapshot $book $changes
    if ($before -cne $after) {
        $failureTag='.failed-'+[guid]::NewGuid().ToString('N')
        [IO.File]::WriteAllText($output+$failureTag+'-before.json',$before)
        [IO.File]::WriteAllText($output+$failureTag+'-after.json',$after)
        throw 'Unapproved content/style/layout change detected; no contract published. Inspect failed snapshots.'
    }
    if ($ExportPreview) {
        try {
            $path=$output+'.after.pdf'
            Require (-not (Test-Path -LiteralPath $path)) 'Preview output already exists.'
            $book.ExportAsFixedFormat(0,$path)
            $previews += @{role='after'; path=$path}
        } catch {$previewErrors += $_.Exception.Message}
    }
    $book.Close($false); $book=$null
    foreach ($inputFile in $inputHashes) {Require ((Hash-File $inputFile.path) -eq $inputFile.sha256) 'An original input changed during execution.'}
    $dateReason='PI date is earlier than or equal to purchase signing date; unchanged.'
    if ($shouldChangeDate) {$dateReason='PI date is later than purchase signing date; set to purchase signing date.'}
    $audit=[ordered]@{
        status='verified'; engine=$Engine; verification_scope='Saved target values, range snapshots, defined names and shape geometry/text; visual review is separate.'
        unchecked_features=@('picture_binary_content','chart_content','conditional_formatting','rich_text_character_formatting_outside_edit_targets','comments_and_hyperlinks','printer_dependent_page_setup','digital_signature_validity','other_advanced_objects')
        inputs=$inputHashes; association=$plan.association; output=$output; output_sha256=(Hash-File $temp)
        title_present=($null -ne $plan.title); pi_date=$piDate.ToString('yyyy-MM-dd'); purchase_signing_date=$purchaseDate.ToString('yyyy-MM-dd')
        date_changed=$shouldChangeDate; date_reason=$dateReason; original_date_format=$dateFormat
        changes=@($changes.ToArray()); checked_snapshot_fields_unchanged=$true; original_inputs_unchanged=$true
        preview_files=$previews; preview_errors=$previewErrors; visual_review='not_performed_by_script'
    }
    $audit | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath $tempAudit -Encoding utf8
    $savedAudit=Get-Content -LiteralPath $tempAudit -Raw -Encoding utf8 | ConvertFrom-Json
    Require ($savedAudit.output_sha256 -eq (Hash-File $temp)) 'Temporary audit verification failed.'
    [IO.File]::Move($temp,$output); $published=$true
    try {
        [IO.File]::Move($tempAudit,$output+'.audit.json'); $auditPublished=$true
    } catch {
        if ($published -and -not $auditPublished) {Remove-Item -LiteralPath $output -Force; $published=$false}
        throw
    }
    Write-Output "Verified contract saved: $output"
} finally {
    if ($null -ne $book) {$book.Close($false)}
    if ($null -ne $bootstrap) {$bootstrap.Close($false)}
    if ($null -ne $excel) {
        if($ownsApplication -and $excel.Workbooks.Count -eq 0){$excel.Quit()}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
    if (Test-Path -LiteralPath $temp) {Remove-Item -LiteralPath $temp -Force}
    if (Test-Path -LiteralPath $tempAudit) {Remove-Item -LiteralPath $tempAudit -Force}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
