#requires -Version 7.2
<#
.SYNOPSIS
Synchronizes the governed SharePoint list "ConnectWise Activity Types" to ConnectWise PSA.

.DESCRIPTION
SharePoint is authoritative. On each timer run the function:
  1. Reads all SharePoint list rows.
  2. Reads all current ConnectWise Activity Types.
  3. Matches by ConnectWise ID first and exact name second.
  4. Creates missing ConnectWise types.
  5. Updates name, point value, and active/inactive state.
  6. Writes the ConnectWise ID and Sync Status back to SharePoint.
  7. Optionally hides ConnectWise types that are not represented in SharePoint.

SAFETY
Every run creates and updates ConnectWise Activity Types and writes the ConnectWise ID
and Sync Status back to SharePoint. Hiding unmanaged ConnectWise types is enabled by
default; set ActivityTypeSync_HideUnmanaged=false to disable it.

The defaults are:
  resource: sales/activities/types
  name field: name
  points field: points
  inactive field: inactiveFlag
Confirm these against the GET response and your ConnectWise Developer Network contract.
#>
param($Timer, $TriggerMetadata)

Set-StrictMode -Version 1
$ErrorActionPreference = 'Stop'

function Get-Setting {
    param([string]$Name,[string]$Default=$null)
    $v=[Environment]::GetEnvironmentVariable($Name)
    if ([string]::IsNullOrWhiteSpace($v)) { return $Default }
    return $v
}
function Test-True { param([string]$Name,[bool]$Default=$false) return (Get-Setting $Name ([string]$Default)) -match '^(?i:true|1|yes)$' }
function Assert-Setting { param([string]$Name,[string]$Value) if([string]::IsNullOrWhiteSpace($Value)){throw "Required setting '$Name' is missing."} }
function Get-Prop { param($Object,[string]$Name) if($null -ne $Object -and $Object.PSObject.Properties.Name -contains $Name){return $Object.$Name}; return $null }
function Get-ConnectWiseApiBaseUrl {
    param([string]$BaseUrl)
    $value=$BaseUrl.Trim()
    if($value -notmatch '^[a-z][a-z0-9+.-]*://'){
        $value="https://$value"
    }

    $uri=$null
    if(-not [uri]::TryCreate($value,[System.UriKind]::Absolute,[ref]$uri) -or
        $uri.Scheme -ne 'https' -or
        [string]::IsNullOrWhiteSpace($uri.Host) -or
        -not [string]::IsNullOrEmpty($uri.Query) -or
        -not [string]::IsNullOrEmpty($uri.Fragment)){
        throw "ConnectWisePsa_ApiBaseUrl must be a valid HTTPS host, optionally followed by /v4_6_release/apis/3.0."
    }

    $apiPath='/v4_6_release/apis/3.0'
    $path=$uri.AbsolutePath.TrimEnd('/')
    if($path -and $path -ne $apiPath){
        throw "ConnectWisePsa_ApiBaseUrl must not contain an unsupported path. Use the host or the complete ConnectWise API path."
    }

    return '{0}://{1}{2}' -f $uri.Scheme,$uri.Authority,$apiPath
}
function ConvertTo-RequiredBoolean {
    param($Value,[string]$FieldName,[int]$ItemId)
    if($Value -is [bool]){return $Value}
    switch -Regex ([string]$Value) {
        '^(?i:true|yes|1)$' { return $true }
        '^(?i:false|no|0)$' { return $false }
        default { throw "SharePoint item $ItemId has a missing or invalid $FieldName value." }
    }
}

# ConnectWise
$CWHost=Get-Setting 'ConnectWisePsa_ApiBaseUrl'
$CWCompany=Get-Setting 'ConnectWisePsa_ApiCompanyId'
$CWPublic=Get-Setting 'ConnectWisePsa_ApiPublicKey'
$CWPrivate=Get-Setting 'ConnectWisePsa_ApiPrivateKey'
$CWClientId=Get-Setting 'ConnectWisePsa_ApiClientId'
$CWVersion=Get-Setting 'ConnectWisePsa_ApiVersion' '2022.1'
$CWResource=(Get-Setting 'ActivityTypeSync_ConnectWiseResource' 'sales/activities/types').Trim('/')
$CWNameField=Get-Setting 'ActivityTypeSync_CWNameField' 'name'
$CWPointsField=Get-Setting 'ActivityTypeSync_CWPointsField' 'points'
$CWInactiveField=Get-Setting 'ActivityTypeSync_CWInactiveField' 'inactiveFlag'

# SharePoint
$SPSite=Get-Setting 'ActivityTypeSync_SharePointSiteUrl' 'https://palaisparc.sharepoint.com/IT/connectwise'
$SPList=Get-Setting 'ActivityTypeSync_SharePointListName' 'ConnectWise Activity Types'
$TenantId=Get-Setting 'Ms365_TenantId'
$SPClientId=Get-Setting 'Ms365_AuthAppId'
$CertBase64=Get-Setting 'Ms365_CertBase64'
$CertPassword=Get-Setting 'Ms365_CertPassword'

$HideUnmanaged=Test-True 'ActivityTypeSync_HideUnmanaged' $true
$ProtectedNames=@((Get-Setting 'ActivityTypeSync_ProtectedNames' 'Historical Entry') -split ',' | ForEach-Object {$_.Trim()} | Where-Object {$_})
$ProtectedIds=@((Get-Setting 'ActivityTypeSync_ProtectedIds' '') -split ',' | ForEach-Object {$_.Trim()} | Where-Object {$_})

foreach($pair in @{
    ConnectWisePsa_ApiBaseUrl=$CWHost; ConnectWisePsa_ApiCompanyId=$CWCompany;
    ConnectWisePsa_ApiPublicKey=$CWPublic; ConnectWisePsa_ApiPrivateKey=$CWPrivate;
    ConnectWisePsa_ApiClientId=$CWClientId; Ms365_TenantId=$TenantId;
    Ms365_AuthAppId=$SPClientId; Ms365_CertBase64=$CertBase64; Ms365_CertPassword=$CertPassword
}.GetEnumerator()){ Assert-Setting $pair.Key ([string]$pair.Value) }

$CWBase=Get-ConnectWiseApiBaseUrl -BaseUrl $CWHost
$auth="$CWCompany+$CWPublic`:$CWPrivate"
$encoded=[Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($auth))
$script:CWHeaders=@{Authorization="Basic $encoded";ClientId=$CWClientId;Accept="application/vnd.connectwise.com+json; version=$CWVersion"}

function Invoke-CW {
    param([ValidateSet('GET','POST','PATCH')][string]$Method,[string]$Path,$Body=$null,[hashtable]$Query=@{},[int]$MaxAttempts=5)
    $pairs=@(foreach($k in $Query.Keys){'{0}={1}' -f [uri]::EscapeDataString([string]$k),[uri]::EscapeDataString([string]$Query[$k])})
    $uri="$CWBase/$($Path.TrimStart('/'))"; if($pairs.Count){$uri+='?'+($pairs -join '&')}
    for($attempt=1;$attempt -le $MaxAttempts;$attempt++){
        try{
            $requestArgs=@{Method=$Method;Uri=$uri;Headers=$script:CWHeaders;TimeoutSec=180}
            if($null -ne $Body){
                $json=ConvertTo-Json -InputObject $Body -Depth 20 -Compress
                $requestArgs.Body=[Text.Encoding]::UTF8.GetBytes($json)
                $requestArgs.ContentType='application/json'
            }
            return Invoke-RestMethod @requestArgs
        }catch{
            $status=$null; try{$status=[int]$_.Exception.Response.StatusCode}catch{}
            $transient=($null -eq $status)-or($status -eq 408)-or($status -eq 429)-or($status -ge 500)
            if(-not $transient -or $Method -eq 'POST' -or $attempt -eq $MaxAttempts){throw "ConnectWise $Method failed. Path=$Path Status=$status ExceptionType=$($_.Exception.GetType().FullName) Error=$($_.Exception.Message) Details=$($_.ErrorDetails.Message) Inner=$($_.Exception.InnerException.Message)"}
            Start-Sleep -Seconds ([Math]::Min(60,[Math]::Pow(2,$attempt)))
        }
    }
}
function Expand-CWResult {
    param($Result)
    # Invoke-RestMethod can hand back a nested array; flatten so each element is one record.
    $items=[Collections.Generic.List[object]]::new()
    foreach($r in @($Result)){
        if($null -eq $r){continue}
        if($r -is [System.Collections.IEnumerable] -and $r -isnot [string] -and $r -isnot [System.Collections.IDictionary] -and $r -isnot [pscustomobject]){
            foreach($inner in (Expand-CWResult $r)){$items.Add($inner)}
        }else{$items.Add($r)}
    }
    return @($items)
}
function Get-CWAll {
    $all=[Collections.Generic.List[object]]::new();$page=1;$size=1000
    $seen=[Collections.Generic.HashSet[string]]::new()
    # The default list can omit inactive records, so fetch them explicitly and merge by id.
    foreach($condition in @($null,"$CWInactiveField=true")){
        $page=1
        do{
            $query=@{page=$page;pageSize=$size;orderBy="$CWNameField asc"}
            if($condition){$query.conditions=$condition}
            $batch=@(Expand-CWResult (Invoke-CW GET $CWResource -Query $query))
            foreach($x in $batch){if($seen.Add([string]$x.id)){$all.Add($x)}}
            $page++
        }while($batch.Count -eq $size)
    }
    return @($all)
}
function Find-CWTypeById {
    param([string]$Id)
    try{return Invoke-CW GET "$CWResource/$Id"}
    catch{
        if("$_" -match 'Status=404'){return $null}
        throw
    }
}
function Find-CWTypeByName {
    param([string]$Name)
    $escaped=$Name.Replace('"','\"')
    $found=@(Expand-CWResult (Invoke-CW GET $CWResource -Query @{conditions="$CWNameField=`"$escaped`"";pageSize=10}))
    if($found.Count -eq 0){return $null}
    return $found[0]
}
function New-CWType {
    param($Desired)
    $body=@{};$body[$CWNameField]=$Desired.Name;$body[$CWPointsField]=$Desired.Points;$body[$CWInactiveField]=(-not $Desired.Active)
    try{return Invoke-CW POST $CWResource $body}
    catch{
        # A redirect/transport failure can happen after ConnectWise already created the record; never retry blindly.
        $existing=Find-CWTypeByName $Desired.Name
        if($existing){Write-Warning "POST reported an error but '$($Desired.Name)' exists as id=$($existing.id); treating it as created. Error: $($_.Exception.Message)";return $existing}
        throw
    }
}
function Patch-CWType {
    param($Current,$Desired)
    $ops=@()
    if([string](Get-Prop $Current $CWNameField) -cne [string]$Desired.Name){$ops+=@{op='replace';path="/$CWNameField";value=$Desired.Name}}
    if([int](Get-Prop $Current $CWPointsField) -ne [int]$Desired.Points){$ops+=@{op='replace';path="/$CWPointsField";value=[int]$Desired.Points}}
    $wantedInactive=(-not [bool]$Desired.Active)
    if([bool](Get-Prop $Current $CWInactiveField) -ne $wantedInactive){$ops+=@{op='replace';path="/$CWInactiveField";value=$wantedInactive}}
    if($ops.Count){return Invoke-CW PATCH "$CWResource/$($Current.id)" $ops}
    return $Current
}
function Hide-CWType {
    param($Current)
    if([bool](Get-Prop $Current $CWInactiveField)){return $Current}
    $ops=@(@{op='replace';path="/$CWInactiveField";value=$true})
    return Invoke-CW PATCH "$CWResource/$($Current.id)" $ops
}

Import-Module PnP.PowerShell -ErrorAction Stop
$secure=ConvertTo-SecureString $CertPassword -AsPlainText -Force
$spConn=Connect-PnPOnline -Url $SPSite -ClientId $SPClientId -Tenant $TenantId -CertificateBase64Encoded $CertBase64 -CertificatePassword $secure -ReturnConnection

$summary=[ordered]@{SharePointRows=0;CWExisting=0;Created=0;Updated=0;Unchanged=0;Hidden=0;HideSkipped=0;Errors=0}
try{
    $spItems=@(Get-PnPListItem -List $SPList -PageSize 5000 -Connection $spConn)
    $summary.SharePointRows=$spItems.Count
    $desired=@(foreach($item in $spItems){
        $points=0
        $pointsValue=$item.FieldValues.CWPointValue
        if($null -eq $pointsValue -or -not [int]::TryParse([string]$pointsValue,[ref]$points)){
            throw "SharePoint item $($item.Id) has a missing or invalid CWPointValue."
        }
        [pscustomobject]@{
            Item=$item; ItemId=$item.Id; Name=[string]$item.FieldValues.Title;
            Category=[string]$item.FieldValues.CWCategory; Purpose=[string]$item.FieldValues.CWPurpose;
            Points=$points; Active=(ConvertTo-RequiredBoolean $item.FieldValues.CWActive 'CWActive' $item.Id);
            CWId=if($item.FieldValues.ConnectWiseID){[string][int]$item.FieldValues.ConnectWiseID}else{$null}
        }
    })
    if($desired.Count -eq 0){throw "SharePoint list '$SPList' contains no governance rows; refusing to alter ConnectWise."}
    if(@($desired|Where-Object{[string]::IsNullOrWhiteSpace($_.Name)}).Count){throw 'One or more SharePoint rows has a blank Activity Type Name.'}
    $dupes=@($desired|Group-Object{$_.Name.Trim().ToLowerInvariant()}|Where-Object Count -gt 1)
    if($dupes.Count){throw "Duplicate SharePoint activity names: $($dupes.Name -join ', ')"}

    $cw=@(Get-CWAll);$summary.CWExisting=$cw.Count
    if($cw.Count -eq 0){throw "ConnectWise returned zero Activity Types from '$CWResource'; refusing to continue."}
    $byId=@{};$byName=@{};foreach($x in $cw){$byId[[string]$x.id]=$x;$n=[string](Get-Prop $x $CWNameField);if($n){$key=$n.Trim().ToLowerInvariant();if($byName.ContainsKey($key)){throw "ConnectWise returned duplicate activity type names '$n'; refusing to make changes."};$byName[$key]=$x}}
    $managedIds=[Collections.Generic.HashSet[string]]::new()
    foreach($d in $desired){
        if($d.CWId -and $byId.ContainsKey($d.CWId)){[void]$managedIds.Add($d.CWId)}
        elseif($byName.ContainsKey($d.Name.Trim().ToLowerInvariant())){[void]$managedIds.Add([string]$byName[$d.Name.Trim().ToLowerInvariant()].id)}
    }

    foreach($d in $desired){
        try{
            $current=$null
            if($d.CWId -and $byId.ContainsKey($d.CWId)){$current=$byId[$d.CWId]}
            elseif($byName.ContainsKey($d.Name.Trim().ToLowerInvariant())){$current=$byName[$d.Name.Trim().ToLowerInvariant()]}

            # The bulk list may be incomplete; a stored ConnectWise ID is authoritative, so renames update instead of creating.
            if($null -eq $current -and $d.CWId){
                $current=Find-CWTypeById $d.CWId
                if($current){Write-Host "MATCHED by id lookup: id=$($current.id) '$($d.Name)'"}
                else{Write-Warning "SharePoint item $($d.ItemId) has ConnectWiseID=$($d.CWId) but ConnectWise returned 404 for it."}
            }elseif($null -eq $current -and -not $d.CWId){
                Write-Warning "SharePoint item $($d.ItemId) '$($d.Name)' has no ConnectWiseID; matching by name only."
            }
            if($null -eq $current){
                $current=Find-CWTypeByName $d.Name
                if($current){Write-Host "MATCHED by name lookup: id=$($current.id) '$($d.Name)'"}
            }
            if($current -and $d.CWId -ne [string]$current.id -and @($desired|Where-Object{$_.CWId -eq [string]$current.id -and $_.ItemId -ne $d.ItemId}).Count){
                throw "ConnectWise type id=$($current.id) is already claimed by another SharePoint row."
            }

            if($null -eq $current){
                Write-Host "CREATE: $($d.Name) points=$($d.Points) active=$($d.Active)"
                $current=New-CWType $d;$summary.Created++
            }else{
                $nameDiff=([string](Get-Prop $current $CWNameField) -cne $d.Name)
                $pointsDiff=([int](Get-Prop $current $CWPointsField) -ne $d.Points)
                $inactiveDiff=([bool](Get-Prop $current $CWInactiveField) -ne (-not $d.Active))
                if($nameDiff -or $pointsDiff -or $inactiveDiff){
                    Write-Host "UPDATE: id=$($current.id) '$([string](Get-Prop $current $CWNameField))' -> '$($d.Name)' points=$($d.Points) active=$($d.Active)"
                    $current=Patch-CWType $current $d;$summary.Updated++
                }else{$summary.Unchanged++}
            }
            if($current -and $current.id){[void]$managedIds.Add([string]$current.id)}

            if($current -and $current.id){
                Set-PnPListItem -List $SPList -Identity $d.ItemId -Values @{ConnectWiseID=[int]$current.id;CWSyncStatus='Synced'} -Connection $spConn | Out-Null
            }
        }catch{
            $summary.Errors++;Write-Error "Failed syncing SharePoint item $($d.ItemId) '$($d.Name)': $($_.Exception.Message)"
            try{Set-PnPListItem -List $SPList -Identity $d.ItemId -Values @{CWSyncStatus='Error'} -Connection $spConn | Out-Null}catch{Write-Warning "Could not write Error status for SharePoint item $($d.ItemId): $($_.Exception.Message)"}
        }
    }

    if($HideUnmanaged -and $summary.Errors -eq 0){
        foreach($x in $cw){
            $id=[string]$x.id;$name=[string](Get-Prop $x $CWNameField)
            if($id -notmatch '^\d+$'){throw "Refusing to hide: unexpected ConnectWise id '$id'."}
            if($managedIds.Contains($id) -or $ProtectedIds -contains $id -or $ProtectedNames -contains $name){continue}
            if(-not [bool](Get-Prop $x $CWInactiveField)){
                Write-Warning "HIDE UNMANAGED: id=$id name='$name'"
                try{
                    Hide-CWType $x|Out-Null;$summary.Hidden++
                }catch{
                    # Types referenced by workflow rules cannot be deactivated; report and continue.
                    $summary.HideSkipped++
                    Write-Warning "HIDE SKIPPED: id=$id name='$name' - $($_.ErrorDetails.Message)$(if(-not $_.ErrorDetails.Message){$_.Exception.Message})"
                }
            }
        }
    }elseif($HideUnmanaged){
        Write-Warning 'Skipping hide-unmanaged pass because one or more SharePoint rows failed to sync.'
    }

    Write-Host ($summary|ConvertTo-Json -Compress)
    if($summary.Errors -gt 0){throw "Activity type synchronization completed with $($summary.Errors) error(s)."}
}finally{
    $spConn=$null
}
