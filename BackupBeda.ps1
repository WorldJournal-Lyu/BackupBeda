<#

BackupBeda.ps1

    2017-11-20 Initial Creation

#>

if (!($env:PSModulePath -match 'C:\\PowerShell\\_Modules')) {
    $env:PSModulePath = $env:PSModulePath + ';C:\PowerShell\_Modules\'
}

Get-Module -ListAvailable WorldJournal.* | Remove-Module -Force
Get-Module -ListAvailable WorldJournal.* | Import-Module -Force

$scriptPath = $MyInvocation.MyCommand.Path
$scriptName = (($MyInvocation.MyCommand) -Replace ".ps1")
$hasError   = $false

$newlog     = New-Log -Path $scriptPath -LogFormat yyyyMMdd-HHmmss
$log        = $newlog.FullName
$logPath    = $newlog.Directory

$mailFrom   = (Get-WJEmail -Name noreply).MailAddress
$mailPass   = (Get-WJEmail -Name noreply).Password
$mailTo     = (Get-WJEmail -Name lyu).MailAddress
$mailSbj    = $scriptName
$mailMsg    = ""

$localTemp = "C:\temp\" + $scriptName + "\"
if (!(Test-Path($localTemp))) {New-Item $localTemp -Type Directory | Out-Null}

Write-Log -Verb "LOG START" -Noun $log -Path $log -Type Long -Status Normal
Write-Line -Length 50 -Path $log

###################################################################################





# Define date variables

$workDate     = (Get-Date).AddDays(0)
$workDate_30  = (Get-Date).AddDays(-30)
$workDate_30w = (Get-Date).AddDays(-(((Get-Date).DayOfWeek.value__)+(7*30)))
$workDay      = ($workDate).DayOfWeek.value__ # 0, 1, 2, 3, 4, 5, 6

# Define server variables

$beda      = (Get-WJPath -Name beda).Path
$back45    = (Get-WJPath -Name back45).Path
$back45_wd = $back45 + $workDate.ToString("yyyyMMdd") + "\"
$back45_30 = $back45 + $workDate_30.ToString("yyyyMMdd") + "\"
$bedaCount = (Get-ChildItem $beda -Recurse -File -Exclude Thumbs.db).Count

if($workDay -eq 4){
    $weeklyPath = $back45_wd + "weeklyPDF\"
    $splin = (Get-WJEmail -Name lyu).MailAddress
}

$tpe       = (Get-WJPath -Name tpe).Path
$backup    = (Get-WJPath -Name backup).Path
$backup_wd = $backup + $workDate.ToString("yyyyMMdd") + "\"
$backup_30 = $backup + $workDate_30.ToString("yyyyMMdd") + "\"
$tpeCount  = (Get-ChildItem $tpe -Recurse -File -Exclude Thumbs.db).Count

$graphic   = (Get-WJPath -Name graphic).Path
$udngroup  = (Get-WJPath -Name udngroup).Path
$overseas  = (Get-WJPath -Name overseas).Path

$marco_production = (Get-WJPath -Name marco_production).Path
$cmpsAT_30 = $marco_production + "AT\Compose\" + $workDate_30.ToString("yyyyMMdd") + "\"
$cmpsBO_30 = $marco_production + "BO\Compose\" + $workDate_30.ToString("yyyyMMdd") + "\"
$cmpsCH_30 = $marco_production + "CH\Compose\" + $workDate_30.ToString("yyyyMMdd") + "\"
$cmpsDC_30 = $marco_production + "DC\Compose\" + $workDate_30.ToString("yyyyMMdd") + "\"
$cmpsNJ_30 = $marco_production + "NJ\Compose\" + $workDate_30.ToString("yyyyMMdd") + "\"
$cmpsNY_30 = $marco_production + "NY\Compose\" + $workDate_30.ToString("yyyyMMdd") + "\"
$cmpsNW_30 = $marco_production + "NW\Compose\" + $workDate_30w.ToString("yyyyMMdd") + "\"

# Construct regex

$list      = @("weekly",4)
$safeList  = @()
For( $i=0; $i -lt $list.Count; $i+=2 ){ 
    if( !($list[$i+1] -eq $workDay) ){ 
        $safeList += $list[$i] 
    } 
}
$regex     = ("^"+(($beda -replace "\\", "\\")) -replace ":", "\:")
if($safeList.Count -gt 1){
    $regex = $regex + "("
    $safeList | ForEach-Object{ $regex = ($regex + $_ + "|") }
    $regex = $regex.Substring(0, $regex.Length-1)
    $regex = $regex + ")"
}elseif($safeList.Count -eq 1){
    $regex = $regex + "(" + $safeList + ")"
}elseif($safeList.Count -eq 0){
    $regex = $regex + "INCLUDE_ALL_FOLDERS"
}


# Define arrays

if($workDay -eq 4){
    $newList = @( $back45_wd, $weeklyPath, $backup_wd )
}else{
    $newList = @( $back45_wd, $backup_wd )
}

$thumbList   = @( $beda, $tpe, $graphic, $udngroup, $overseas, 
                  $cmpsAT_30, $cmpsBO_30, $cmpsCH_30, $cmpsDC_30, $cmpsNJ_30, $cmpsNY_30, $cmpsNW_30 )

$clearList   = @( $graphic, $udngroup, $overseas )

$deleteList  = @( $back45_30, $backup_30, 
                  $cmpsAT_30, $cmpsBO_30, $cmpsCH_30, $cmpsDC_30, $cmpsNJ_30, $cmpsNY_30, $cmpsNW_30 )

# Log variables

Write-Log -Verb "workDate" -Noun $workDate.ToString("yyyyMMdd") -Path $log -Type Short -Status Normal
Write-Log -Verb "workDate_30" -Noun $workDate_30.ToString("yyyyMMdd") -Path $log -Type Short -Status Normal
Write-Log -Verb "workDate_30w" -Noun $workDate_30w.ToString("yyyyMMdd") -Path $log -Type Short -Status Normal
Write-Log -Verb "workDay" -Noun $workDay -Path $log -Type Short -Status Normal
Write-Line -Length 50 -Path $log

Write-Log -Verb "beda" -Noun $beda -Path $log -Type Short -Status Normal
Write-Log -Verb "back45" -Noun $back45 -Path $log -Type Short -Status Normal
Write-Log -Verb "back45_wd" -Noun $back45_wd -Path $log -Type Short -Status Normal
Write-Log -Verb "back45_30" -Noun $back45_30 -Path $log -Type Short -Status Normal
Write-Log -Verb "bedaCount" -Noun $bedaCount -Path $log -Type Short -Status Normal
Write-Line -Length 50 -Path $log

if($workDay -eq 4){
    Write-Log -Verb "weeklyPath" -Noun $weeklyPath -Path $log -Type Short -Status Normal
    Write-Log -Verb "splin" -Noun $splin -Path $log -Type Short -Status Normal
    Write-Line -Length 50 -Path $log
}

Write-Log -Verb "tpe" -Noun $tpe -Path $log -Type Short -Status Normal
Write-Log -Verb "backup" -Noun $backup -Path $log -Type Short -Status Normal
Write-Log -Verb "backup_wd" -Noun $backup_wd -Path $log -Type Short -Status Normal
Write-Log -Verb "backup_30" -Noun $backup_30 -Path $log -Type Short -Status Normal
Write-Log -Verb "tpeCount" -Noun $tpeCount -Path $log -Type Short -Status Normal
Write-Line -Length 50 -Path $log

Write-Log -Verb "graphic" -Noun $graphic -Path $log -Type Short -Status Normal
Write-Log -Verb "udngroup" -Noun $udngroup -Path $log -Type Short -Status Normal
Write-Log -Verb "overseas" -Noun $overseas -Path $log -Type Short -Status Normal
Write-Line -Length 50 -Path $log

Write-Log -Verb "marco_production" -Noun $marco_production -Path $log -Type Short -Status Normal
Write-Log -Verb "cmpsAT_30" -Noun $cmpsAT_30 -Path $log -Type Short -Status Normal
Write-Log -Verb "cmpsBO_30" -Noun $cmpsBO_30 -Path $log -Type Short -Status Normal
Write-Log -Verb "cmpsCH_30" -Noun $cmpsCH_30 -Path $log -Type Short -Status Normal
Write-Log -Verb "cmpsDC_30" -Noun $cmpsDC_30 -Path $log -Type Short -Status Normal
Write-Log -Verb "cmpsNJ_30" -Noun $cmpsNJ_30 -Path $log -Type Short -Status Normal
Write-Log -Verb "cmpsNY_30" -Noun $cmpsNY_30 -Path $log -Type Short -Status Normal
Write-Log -Verb "cmpsNW_30" -Noun $cmpsNW_30 -Path $log -Type Short -Status Normal
Write-Line -Length 50 -Path $log

Write-Log -Verb "safeList" -Noun ($safeList -join ", ") -Path $log -Type Short -Status Normal
Write-Log -Verb "regex" -Noun $regex -Path $log -Type Short -Status Normal
Write-Line -Length 50 -Path $log

$newList | ForEach-Object{ Write-Log -Verb "newList" -Noun $_ -Path $log -Type Short -Status Normal }
$thumbList | ForEach-Object{ Write-Log -Verb "thumbList" -Noun $_ -Path $log -Type Short -Status Normal }
$clearList | ForEach-Object{ Write-Log -Verb "clearList" -Noun $_ -Path $log -Type Short -Status Normal }
$deleteList | ForEach-Object{ Write-Log -Verb "deleteList" -Noun $_ -Path $log -Type Short -Status Normal }

Write-Line -Length 50 -Path $log





# 1 Create new folders in $newList

Write-Log -Verb "CREATE FOLDERS" -Noun "newList" -Path $log -Type Long -Status System; Pause

$newList | ForEach-Object{
    if(Test-Path $_){
        Write-Log -Verb "IS EXIST" -Noun $_ -Path $log -Type Long -Status Good
    }else{
        try{
            New-Item -ItemType Directory -Path $_ | Out-Null
            Write-Log -Verb "NEW" -Noun $_ -Path $log -Type Long -Status Good
        }catch{
            Write-Log -Verb "NEW" -Noun $_ -Path $log -Type Long -Status Bad
            $mailMsg = $mailMsg + (Write-Log -Verb "NEW" -Noun $_ -Path $log -Type Long -Status Bad -Output String) + "`n"
            $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $_.Exception -Path $log -Type Short -Status Bad -Output String) + "`n"
            $hasError = $true
        }
    }
    Pause
}

Write-Line -Length 50 -Path $log





# 2 Delete thumbnails in folders in $thumbList

Write-Log -Verb "DELETE THUMBNAILS" -Noun "thumbList" -Path $log -Type Long -Status System; Pause

$thumbList | ForEach-Object{
    Delete-Thumbs $_ | ForEach-Object{
        if( ($_.Status -eq "Bad") -or ($_.Status -eq "Warning") ){
            Write-Log -Verb $_.Verb -Noun $_.Noun -Path $log -Type Long -Status $_.Status
            Write-Log -Verb "Exception" -Noun $_.Exception -Path $log -Type Short -Status $_.Status
        }else{
            Write-Log -Verb $_.Verb -Noun $_.Noun -Path $log -Type Long -Status $_.Status
        }
    }
}

Write-Line -Length 50 -Path $log





# (Thursdays only) Backup weekly PDF for splin

if(($workDay -eq 4) -and (Test-Path $weeklyPath)){
    Write-Log -Verb "BACKUP WEEKLY" -Noun $weeklyPath -Path $log -Type Long -Status System; Pause
    Get-ChildItem ($beda+"weekly") -Include 455*.pdf, 43*.pdf -Recurse | ForEach-Object{
        try{
            Copy-Item $_.FullName $weeklyPath -ErrorAction Stop
            Write-Log -Verb "COPY" -Noun $_.FullName -Path $log -Type Long -Status Good
        }catch{
            $mailMsg = $mailMsg + (Write-Log -Verb "COPY" -Noun $_.FullName -Path $log -Type Long -Status Bad -Output String) + "`n"
            $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $_.Exception -Path $log -Type Short -Status Bad -Output String) + "`n"
            $hasError = $true
        }
    }
    $weeklyPdfCount = (Get-ChildItem $weeklyPath).Count
    if($weeklyPdfCount -eq 0){
        Emailv3 -From $mailFrom -Pass $mailPass -To $mailTo -Subject ("ERROR Weekly PDF " + $workDate.ToString("yyyy-MM-dd")) -Body ("Path: "+$weeklyPath+" ("+$weeklyPdfCount+" files)")
    }else{
        Emailv3 -From $mailFrom -Pass $mailPass -To $splin -Subject ("Weekly PDF " + $workDate.ToString("yyyy-MM-dd")) -Body ("Path: Back45\"+$workDate.ToString("yyyyMMdd")+"\weeklyPDF"+" ("+$weeklyPdfCount+" files)")
    }
}

Write-Line -Length 50 -Path $log





# 3 Backup $beda to $back45_wd, exclude folders in $safeList

Write-Log -Verb "BACKUP BEDA" -Noun "back45_wd "-Path $log -Type Long -Status System; Pause

if(Test-Path $back45_wd){
    Get-ChildItem $beda -Recurse | Where-Object{
        !($_.FullName -match $regex)
    } | Sort-Object FullName -Descending | Move-Files -From $beda -To $back45_wd | ForEach-Object{
        Write-Log -Verb "moveFrom" -Noun $_.MoveFrom -Path $log -Type Short -Status Normal
        Write-Log -Verb "moveTo" -Noun $_.MoveTo -Path $log -Type Short -Status Normal
        if($_.Status -eq "Bad"){
            $mailMsg = $mailMsg + (Write-Log -Verb $_.Verb -Noun $_.Noun -Path $log -Type Long -Status $_.Status -Output String) + "`n"
            $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $_.Exception -Path $log -Type Short -Status $_.Status -Output String) + "`n"
            $hasError = $true
        }else{
            Write-Log -Verb $_.Verb -Noun $_.Noun -Path $log -Type Long -Status $_.Status
        }
    }
}else{
    $mailMsg = $mailMsg + (Write-Log -Verb "NOT EXIST" -Noun $back45_wd -Path $log -Type Long -Status Bad -Output String) + "`n"
    $hasError = $true
}

Write-Line -Length 50 -Path $log




# (Monday to Friday) Create 45101 folder

if(($weekDay -ne 6) -or ($weekDay -ne 0)){
    try{    
        New-Item -ItemType Directory -Path ($beda+"45101") | Out-Null
        Write-Log -Verb "NEW" -Noun ($beda+"45101") -Path $log -Type Long -Status Good
    }catch{
        $mailMsg = $mailMsg + (Write-Log -Verb "NEW" -Noun ($beda+"45101") -Path $log -Type Long -Status Bad -Output String) + "`n"
        $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $_.Exception -Path $log -Type Short -Status $_.Status -Output String) + "`n"
        $hasError = $true
    }
}





# 4 Backup $tpe to $backup_wd

Write-Log -Verb "BACKUP TPE" -Noun "backup_wd "-Path $log -Type Long -Status System; Pause

if(Test-Path $backup_wd){
    Get-ChildItem $tpe -Recurse | Sort-Object FullName -Descending | Move-Files -From $tpe -To $backup_wd | ForEach-Object{
        Write-Log -Verb "moveFrom" -Noun $_.MoveFrom -Path $log -Type Short -Status Normal
        Write-Log -Verb "moveTo" -Noun $_.MoveTo -Path $log -Type Short -Status Normal
        if($_.Status -eq "Bad"){
            $mailMsg = $mailMsg + (Write-Log -Verb $_.Verb -Noun $_.Noun -Path $log -Type Long -Status $_.Status -Output String) + "`n"
            $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $_.Exception -Path $log -Type Short -Status $_.Status -Output String) + "`n"
            $hasError = $true
        }else{
            Write-Log -Verb $_.Verb -Noun $_.Noun -Path $log -Type Long -Status $_.Status
        }
    }
}else{
    $mailMsg = $mailMsg + (Write-Log -Verb "NOT EXIST" -Noun $backup_wd -Path $log -Type Long -Status Bad -Output String) + "`n"
    $hasError = $true
}
Write-Line -Length 50 -Path $log





# 5 Clear contents in folders in $clearList

Write-Log -Verb "CLEAR FOLDERS" -Noun "clearList" -Path $log -Type Long -Status System; Pause

$clearList | ForEach-Object{
    if(Test-Path $_){
        if((Get-ChildItem $_).Count -eq 0){
            Write-Log -Verb "IS EMPTY" -Noun $_ -Path $log -Type Long -Status Normal
        }else{
            Get-ChildItem $_ -Recurse | Sort-Object FullName -Descending | ForEach-Object{
                try{
                    $temp = $_.FullName
                    Remove-Item $_.FullName -Force -ErrorAction Stop
                    Write-Log -Verb "REMOVE" -Noun $temp -Path $log -Type Long -Status Good
                }catch{
                    $mailMsg = $mailMsg + (Write-Log -Verb "REMOVE" -Noun $_.FullName -Path $log -Type Long -Status Bad -Output String) + "`n"
                    $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $_.Exception -Path $log -Type Short -Status Bad -Output String) + "`n"
                    $hasError = $true
                }
            }
        }
    }else{
        $mailMsg = $mailMsg + (Write-Log -Verb "NOT EXIST" -Noun $_ -Path $log -Type Long -Status Warning -Output String) + "`n"
        $hasError = $true
    }
    Pause
}





# 6 Clear and delete folders in $deleteList

Write-Log -Verb "DELETE FOLDERS" -Noun "deleteList" -Path $log -Type Long -Status System; Pause

$deleteList | ForEach-Object{
    if(Test-Path $_){
        Get-ChildItem $_ -Recurse | Sort-Object FullName -Descending | ForEach-Object{
            try{
                $temp = $_.FullName
                Remove-Item $_.FullName -Force -ErrorAction Stop
                Write-Log -Verb "REMOVE" -Noun $temp -Path $log -Type Long -Status Good

            }catch{
                $mailMsg = $mailMsg + (Write-Log -Verb "REMOVE" -Noun $_.FullName -Path $log -Type Long -Status Bad -Output String) + "`n"
                $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $_.Exception -Path $log -Type Short -Status Bad -Output String) + "`n"
                $hasError = $true
            }
        }
        try{
            $temp = $_
            Remove-Item $_ -Recurse -Force -ErrorAction Stop
            Write-Log -Verb "REMOVE" -Noun $temp -Path $log -Type Long -Status Good
        }catch{
            $mailMsg = $mailMsg + (Write-Log -Verb "REMOVE" -Noun $temp -Path $log -Type Long -Status Bad -Output String) + "`n"
            $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $_.Exception.Message -Path $log -Type Short -Status Bad -Output String) + "`n"
            $hasError = $true
        }
    }else{
        Write-Log -Verb "NOT EXIST" -Noun $_ -Path $log -Type Long -Status Normal
    }
    Pause
}

Write-Line -Length 50 -Path $log





# 7 Check result and compose mail

if($workDay -eq 4){
    $back45Count = (Get-ChildItem $back45_wd -Recurse -File -Exclude Thumbs.db).Count - (Get-ChildItem $weeklyPath -Recurse -File -Exclude Thumbs.db).Count
}else{
    $back45Count = (Get-ChildItem $back45_wd -Recurse -File -Exclude Thumbs.db).Count
}
$backupCount = (Get-ChildItem $backup_wd -Recurse -File -Exclude Thumbs.db).Count

$mailMsg = $mailMsg + $back45_wd + "`n" + "RESULT " + $back45Count + " (EXPECTED " + $bedaCount + ")`n`n"
$mailMsg = $mailMsg + $backup_wd + "`n" + "RESULT " + $backupCount + " (EXPECTED " + $tpeCount + ")`n`n"


$clearList | ForEach-Object{ 
    $count = (Get-ChildItem $_ -Recurse -Exclude Thumbs.db).Count
    if( $count -eq 0 ){ $result = "CLEARED" }else{ $result = "NOT CLEARED"; $hasError = $true }
    $mailMsg = $mailMsg + $_ + "`n" + $result + "`n`n"
}

$deleteList | ForEach-Object{ 
    $testpath = (Test-Path $_)
    if( $testpath -eq $false ){ $result = "DELETED" }else{ $result = "NOT DELETED"; $hasError = $true }
    $mailMsg = $mailMsg + $_ + "`n" + $result + "`n`n"
}






###################################################################################

Write-Line -Length 50 -Path $log

# Delete temp folder

Write-Log -Verb "REMOVE" -Noun $localTemp -Path $log -Type Long -Status Normal
try{
    $temp = $localTemp
    Remove-Item $localTemp -Recurse -Force -ErrorAction Stop
    Write-Log -Verb "REMOVE" -Noun $temp -Path $log -Type Long -Status Good
}catch{
    $mailMsg = $mailMsg + (Write-Log -Verb "REMOVE" -Noun $temp -Path $log -Type Long -Status Bad -Output String) + "`n"
    $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $_.Exception.Message -Path $log -Type Short -Status Bad -Output String) + "`n"
}

Write-Line -Length 50 -Path $log
Write-Log -Verb "LOG END" -Noun $log -Path $log -Type Long -Status Normal
if($hasError){ $mailSbj = "ERROR " + $scriptName }

$emailParam = @{
    From    = $mailFrom
    Pass    = $mailPass
    To      = $mailTo
    Subject = $mailSbj
    Body    = $mailMsg
    ScriptPath = $scriptPath
    Attachment = $log
}
$mailMsg
Emailv2 @emailParam