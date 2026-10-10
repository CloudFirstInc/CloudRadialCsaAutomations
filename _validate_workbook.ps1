$wb = "C:\Users\DennisGentles\Downloads\City_of_Rye_City_Hall_ConnectWise_Project_Import_CloudRADIUS_Rebalanced 1.xlsx"
Import-Module ImportExcel -ErrorAction Stop

function Protect-Text([object]$Value){if($null-eq $Value){return ''};return ([string]$Value).Trim()}
function As-Decimal([object]$Value,[string]$Field){
  $d=0D
  if(-not[decimal]::TryParse((Protect-Text $Value),[Globalization.NumberStyles]::Any,[Globalization.CultureInfo]::InvariantCulture,[ref]$d)){
    return $null
  }
  return $d
}
function As-CWDate([object]$Value,[string]$Field){
  $d=[datetime]::MinValue
  if($Value -is [datetime]){ return $Value }
  elseif(-not[datetime]::TryParse((Protect-Text $Value),[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal,[ref]$d)){
    return $null
  }
  return $d
}

$projectRows=@(Import-Excel $wb -WorksheetName Import_Project)
$phases=@(Import-Excel $wb -WorksheetName Import_Phases)
$tickets=@(Import-Excel $wb -WorksheetName Import_Tickets)
$tasks=@(Import-Excel $wb -WorksheetName Import_Tasks)

$errors = New-Object System.Collections.Generic.List[string]

if($projectRows.Count -ne 1){$errors.Add("Import_Project must contain exactly one data row (found $($projectRows.Count)).")}

# --- Phases ---
$phaseKeys=@{}
foreach($p in $phases){
  $phaseValue=Protect-Text $p.PhaseNumber;$phaseInt=0
  if(-not[int]::TryParse($phaseValue,[ref]$phaseInt) -or $phaseInt -lt 1){$errors.Add("Invalid positive PhaseNumber: '$phaseValue' (Phase '$($p.PhaseName)')")}
  else{
    $k=$phaseInt.ToString().PadLeft(2,'0')
    if($phaseKeys.ContainsKey($k)){$errors.Add("Duplicate phase number $k")}
    $phaseKeys[$k]=$true
  }
  if($null -eq (As-Decimal $p.BudgetHours "Phase $($p.PhaseNumber) BudgetHours")){$errors.Add("Phase $($p.PhaseNumber) has bad BudgetHours: '$($p.BudgetHours)'")}
  if($null -eq (As-CWDate $p.StartDate "Phase $($p.PhaseNumber) StartDate")){$errors.Add("Phase $($p.PhaseNumber) has bad StartDate: '$($p.StartDate)'")}
  if($null -eq (As-CWDate $p.DueDate "Phase $($p.PhaseNumber) DueDate")){$errors.Add("Phase $($p.PhaseNumber) has bad DueDate: '$($p.DueDate)'")}
  if([string]::IsNullOrWhiteSpace((Protect-Text $p.PhaseName))){$errors.Add("Phase $($p.PhaseNumber) has empty PhaseName")}
  if([string]::IsNullOrWhiteSpace((Protect-Text $p.Summary))){$errors.Add("Phase $($p.PhaseNumber) has empty Summary")}
  if([string]::IsNullOrWhiteSpace((Protect-Text $p.ExitCriteria))){$errors.Add("Phase $($p.PhaseNumber) has empty ExitCriteria")}
}

# --- Tickets ---
$ticketKeys=@{}
foreach($t in $tickets){
  $k=Protect-Text $t.TicketNumber
  if([string]::IsNullOrWhiteSpace($k)){$errors.Add("A ticket has an empty TicketNumber (Summary: '$($t.Summary)')"); continue}
  if($ticketKeys.ContainsKey($k)){$errors.Add("Duplicate ticket number $k")}
  $ticketKeys[$k]=$true
  $phaseValue=Protect-Text $t.PhaseNumber;$phaseInt=0
  if(-not[int]::TryParse($phaseValue,[ref]$phaseInt) -or $phaseInt -lt 1){$errors.Add("Ticket $k has invalid phase number '$phaseValue'")}
  else{
    $pk=$phaseInt.ToString().PadLeft(2,'0')
    if(-not$phaseKeys.ContainsKey($pk)){$errors.Add("Ticket $k references missing phase $pk")}
  }
  if($null -eq (As-Decimal $t.BudgetHours "Ticket $k BudgetHours")){$errors.Add("Ticket $k has bad BudgetHours: '$($t.BudgetHours)'")}
  if($null -eq (As-CWDate $t.StartDate "Ticket $k StartDate")){$errors.Add("Ticket $k has bad StartDate: '$($t.StartDate)'")}
  if($null -eq (As-CWDate $t.DueDate "Ticket $k DueDate")){$errors.Add("Ticket $k has bad DueDate: '$($t.DueDate)'")}
  $pred = Protect-Text $t.PredecessorTicket
  if($pred -and $pred -ne 'None' -and -not $ticketKeys.ContainsKey($pred) -and ($tickets.TicketNumber -notcontains $pred)){$errors.Add("Ticket $k PredecessorTicket '$pred' not found among ticket numbers")}
}

# --- Tasks ---
$taskKeys=@{}
foreach($x in $tasks){
  $ticket=Protect-Text $x.TicketNumber
  if(-not$ticketKeys.ContainsKey($ticket)){$errors.Add("Task '$($x.TaskName)' references missing ticket $ticket")}
  $seq=0;$sequenceText=Protect-Text $x.TaskSequence
  if(-not[int]::TryParse($sequenceText,[ref]$seq) -or $seq -lt 1){$errors.Add("Invalid positive TaskSequence '$sequenceText' for ticket $ticket task '$($x.TaskName)'")}
  $k="$ticket|$seq"
  if($taskKeys.ContainsKey($k)){$errors.Add("Duplicate task key $k")}
  $taskKeys[$k]=$true
  if($null -eq (As-Decimal $x.EstimatedHours "Task $k EstimatedHours")){$errors.Add("Task $k has bad EstimatedHours: '$($x.EstimatedHours)'")}
}

# --- Hour totals ---
$taskTotal=($tasks|ForEach-Object{ $d=As-Decimal $_.EstimatedHours 'EstimatedHours'; if($d){$d}else{0} }|Measure-Object -Sum).Sum
$ticketTotal=($tickets|ForEach-Object{ $d=As-Decimal $_.BudgetHours 'BudgetHours'; if($d){$d}else{0} }|Measure-Object -Sum).Sum
$phaseTotal=($phases|ForEach-Object{ $d=As-Decimal $_.BudgetHours 'BudgetHours'; if($d){$d}else{0} }|Measure-Object -Sum).Sum
Write-Host ""
Write-Host "Task total hours:   $taskTotal"
Write-Host "Ticket total hours: $ticketTotal"
Write-Host "Phase total hours:  $phaseTotal"
if([math]::Abs($taskTotal-$ticketTotal) -gt .01 -or [math]::Abs($ticketTotal-$phaseTotal) -gt .01){
  $errors.Add("Hour totals disagree: tasks=$taskTotal tickets=$ticketTotal phases=$phaseTotal")
}

# --- BillableFlag / Is-Yes sanity ---
function Is-Yes([object]$Value){return (Protect-Text $Value) -match '^(yes|true|1)$'}
$badTicketFlags = $tickets | Where-Object { (Protect-Text $_.BillableFlag) -and (Protect-Text $_.BillableFlag) -notmatch '^(yes|true|1|no|false|0)$' }
foreach($b in $badTicketFlags){$errors.Add("Ticket $($b.TicketNumber) has unrecognized BillableFlag value '$($b.BillableFlag)' (will be treated as non-billable)")}
$badTaskFlags = $tasks | Where-Object { (Protect-Text $_.BillableFlag) -and (Protect-Text $_.BillableFlag) -notmatch '^(yes|true|1|no|false|0)$' }
foreach($b in $badTaskFlags){$errors.Add("Task $($b.TicketNumber)/$($b.TaskSequence) has unrecognized BillableFlag value '$($b.BillableFlag)' (will be treated as non-billable)")}

Write-Host ""
if($errors.Count -eq 0){
  Write-Host "NO VALIDATION ERRORS FOUND" -ForegroundColor Green
}else{
  Write-Host "$($errors.Count) VALIDATION ISSUE(S) FOUND:" -ForegroundColor Red
  $errors | ForEach-Object { Write-Host " - $_" -ForegroundColor Yellow }
}

Write-Host ""
Write-Host "--- All phase rows ---"
$phases | ForEach-Object { "[{0}] {1} | Budget={2} | Start={3} | Due={4}" -f $_.PhaseNumber,$_.PhaseName,$_.BudgetHours,$_.StartDate,$_.DueDate } | Write-Host

Write-Host ""
Write-Host "--- All ticket rows (phase, ticket, budget) ---"
$tickets | ForEach-Object { "Phase {0} | Ticket {1} | Budget={2}" -f $_.PhaseNumber,$_.TicketNumber,$_.BudgetHours } | Write-Host
