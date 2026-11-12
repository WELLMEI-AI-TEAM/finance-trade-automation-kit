param([Parameter(Mandatory=$true)][string]$Session)
$ErrorActionPreference='Stop'
$s=Get-Content -LiteralPath $Session -Raw -Encoding utf8 | ConvertFrom-Json
if($s.owned -ne $true){throw 'Process is not owned by this batch.'}
$p=Get-Process -Id ([int]$s.pid) -ErrorAction SilentlyContinue
if($null -eq $p){exit 0}
if($p.ProcessName -notin @('EXCEL','et')){throw 'Unexpected process executable.'}
$start=if($s.started -is [DateTime]){$s.started.ToUniversalTime()}else{[DateTime]::Parse([string]$s.started).ToUniversalTime()}
if($p.StartTime.ToUniversalTime().Ticks -ne $start.Ticks){throw 'Process start time changed; refusing cleanup.'}
Stop-Process -Id $p.Id
