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





# Date values

$workDate     = (Get-Date).AddDays(0)
$workDate_30  = (Get-Date).AddDays(-30)
$workDate_30w = (Get-Date).AddDays(-(((Get-Date).DayOfWeek.value__)+(7*30)))
$workDay      = ($workDate).DayOfWeek.value__ # 0, 1, 2, 3, 4, 5, 6

# Server paths

$beda      = (Get-WJPath -Name beda).Path
$back45    = (Get-WJPath -Name back45).Path
$back45_wd = $back45 + $workDate.ToString("yyyyMMdd") + "\"
$back45_30 = $back45 + $workDate_30.ToString("yyyyMMdd") + "\"

if($workDate -eq 4){
    $weeklyPath = $back45_wd + "weeklyPDF\"
}

$tpe       = (Get-WJPath -Name tpe).Path
$backup    = (Get-WJPath -Name backup).Path
$backup_wd = $backup + $workDate.ToString("yyyyMMdd") + "\"
$backup_30 = $backup + $workDate_30.ToString("yyyyMMdd") + "\"

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

$list      = @("weekly",4)
$safeList  = @()
For( $i=0; $i -lt $list.Count; $i+=2 ){ 
    if( !($list[$i+1] -eq $workDay) ){ 
        $safeList += $list[$i] 
    } 
}


Write-Log -Verb "workDate" -Noun $workDate.ToString("yyyyMMdd") -Path $log -Type Short -Status Normal
Write-Log -Verb "workDate_30" -Noun $workDate_30.ToString("yyyyMMdd") -Path $log -Type Short -Status Normal
Write-Log -Verb "workDate_30w" -Noun $workDate_30w.ToString("yyyyMMdd") -Path $log -Type Short -Status Normal
Write-Log -Verb "workDay" -Noun $workDay -Path $log -Type Short -Status Normal
Write-Line -Length 50 -Path $log

Write-Log -Verb "beda" -Noun $beda -Path $log -Type Short -Status Normal
Write-Log -Verb "back45" -Noun $back45 -Path $log -Type Short -Status Normal
Write-Log -Verb "back45_wd" -Noun $back45_wd -Path $log -Type Short -Status Normal
Write-Log -Verb "back45_30" -Noun $back45_30 -Path $log -Type Short -Status Normal
Write-Line -Length 50 -Path $log

if($workDay -eq 4){
    Write-Log -Verb "weeklyPath" -Noun $weeklyPath -Path $log -Type Short -Status Normal
}

Write-Log -Verb "tpe" -Noun $tpe -Path $log -Type Short -Status Normal
Write-Log -Verb "backup" -Noun $backup -Path $log -Type Short -Status Normal
Write-Log -Verb "backup_wd" -Noun $backup_wd -Path $log -Type Short -Status Normal
Write-Log -Verb "backup_30" -Noun $backup_30 -Path $log -Type Short -Status Normal
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

Write-Line -Length 50 -Path $log





# Create new paths in $newList

Write-Log -Verb "NEW-ITEM" -Noun "newList" -Path $log -Type Long -Status System

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
}

Write-Line -Length 50 -Path $log





# Delete thumbnails in paths in $thumbList

Write-Log -Verb "DELETE-THUMBS" -Noun "thumbList" -Path $log -Type Long -Status System

$thumbList | ForEach-Object{
    Delete-Thumbs $_ | ForEach-Object{
        if( ($_.Status -eq "Bad") -or ($_.Status -eq "Warning") ){
            $mailMsg = $mailMsg + (Write-Log -Verb $_.Verb -Noun $_.Noun -Path $log -Type Long -Status $_.Status -Output String) + "`n"
            $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $_.Exception -Path $log -Type Short -Status $_.Status) + "`n"
        }else{
            Write-Log -Verb $_.Verb -Noun $_.Noun -Path $log -Type Long -Status $_.Status
        }
    }
}

Write-Line -Length 50 -Path $log





# Backup weekly PDF on Thursday

if(($workDay -eq 4) -and (Test-Path $weeklyPath)){
    Write-Log -Verb "WEEKLY-PDF" -Noun $weeklyPath -Path $log -Type Long -Status System
    $splin = (Get-WJEmail -Name splin).MailAddress
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





# Move $beda to $back45_wd, exclude folders in $safeList

Write-Log -Verb "MOVE-FILES" -Noun "back45_wd "-Path $log -Type Long -Status System

if((Test-Path $back45_wd)){
    $regex = ("^"+(($beda -replace "\\", "\\")) -replace ":", "\:")
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
    Write-Log -Verb "regex" -Noun $regex -Path $log -Type Short -Status Normal

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





# Create beda\45101 from Monday thru Friday

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





# Move $tpe to $backup_wd

Write-Log -Verb "MOVE-FILES" -Noun "backup_wd "-Path $log -Type Long -Status System

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





# Clear contents in $clearList

Write-Log -Verb "REMOVE-FILES" -Noun "clearList" -Path $log -Type Long -Status System

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
}

Write-Line -Length 50 -Path $log





# Delete paths in $deleteList

Write-Log -Verb "REMOVE-FILES & REMOVE-ITEM" -Noun "deleteList" -Path $log -Type Long -Status System

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
}

Write-Line -Length 50 -Path $log





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
Emailv2 @emailParam