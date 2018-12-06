<#
##################################################################################################################
#
# Microsoft Premier Field Engineering
# 
# Migrate1.ps1
# v1.0 Initial creation 06/12/2017 - Perform AD User Migration Tasks
#
# 
# 
# Microsoft Disclaimer for custom scripts
# ================================================================================================================
# The sample scripts are not supported under any Microsoft standard support program or service. The sample scripts
# are provided AS IS without warranty of any kind. Microsoft further disclaims all implied warranties including, 
# without limitation, any implied warranties of merchantability or of fitness for a particular purpose. The entire
# risk arising out of the use or performance of the sample scripts and documentation remains with you. In no event
# shall Microsoft, its authors, or anyone else involved in the creation, production, or delivery of the scripts be
# liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business
# interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability to 
# use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.
# ================================================================================================================
#
##################################################################################################################
# Script variables - please do not change these unless you know what you are doing
##################################################################################################################
#>

$VBCrLf = "`r`n"
$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$Log = "$scriptDir\Phase1Log.txt"
$users = Import-Csv $scriptDir\users.csv
$Global:ErrorActionPreference = 'Stop'

##################################################################################################################
# Functions - please do not change these unless you know what you are doing
##################################################################################################################

Function startScript()
{
    $msg = "Beginning Migration tasks from $env:COMPUTERNAME" + $VBCrLf + "@ $(get-date) via PowerShell Script.  Logging is enabled."
    Write-Host "######################################################################################" -ForegroundColor Yellow
    Write-Host  "$msg" -ForegroundColor Green
    Write-Host "######################################################################################" -ForegroundColor Yellow
    Write-Host
    Write-This $msg $Log
}

Function Write-This($data, $script)
{
    try
    {
        Add-Content -Path $script -Value $data -ErrorAction Stop
    }
    catch
    {
        write-host $_.Exception.Message
    }
}

Function closeScript($exitCode)
{
    if($exitCode -ne 0)
    {
        Write-Host
        Write-Host "######################################################################################" -ForegroundColor Yellow
        $msg = "Script execution unsuccessful, and terminted at $(get-date)" + $VBCrLf + "Time Elapsed: ($($elapsed.Elapsed.ToString()))" `
        + $VBCrLf + "Examine the script output and previous events logged to resolve errors."
        Write-Host $msg -ForegroundColor Red
        Write-Host "######################################################################################" -ForegroundColor Yellow
        Write-This $msg $log
    }
    else
    {
        Write-Host "######################################################################################" -ForegroundColor Yellow
        $msg = "Successfully completed script at $(get-date)" + $VBCrLf + "Time Elapsed: ($($elapsed.Elapsed.ToString()))" + $VBCrLf `
        + "Review the logs."
        Write-Host $msg -ForegroundColor Green
        Write-Host "######################################################################################" -ForegroundColor Yellow
        Write-This $msg $log
    }
    exit $exitCode
}

##################################################################################################################
# Prerequisites  - please do not change unless you know what you are doing
##################################################################################################################

#cred requires target domain Exchange and SKYPE rights
$TARGETCred = Get-Credential -Message "Enter your target domain General Admin Credential."

#cred requires source domain Exchange rights
$SOURCECred = Get-Credential -Message "Enter your source domain general ADM Credential."

##################################################################################################################
# Begin Script  - please do not change unless you know what you are doing
##################################################################################################################

$elapsed = [System.Diagnostics.Stopwatch]::StartNew()
StartScript

#process all target domain users in Bulk to mail enable
Write-Host "*************************************************************************************" -ForegroundColor White
Write-Host "Phase 1 - Bulk Mail Enabling the users in target domain..." -ForegroundColor White
Write-Host "*************************************************************************************" -ForegroundColor White
Write-Host

#targetdomain.com Exchange Session
$HQPSSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchangeserver.targetdomain.com/powershell -Credential $HQCred
Import-PSSession $HQPSSession -DisableNameChecking -AllowClobber | Out-Null

$i = 1
foreach($user in $users)
{
    try
    {
        Write-Host "Mail Enabling" $i "of" $users.Count "Users..." -ForegroundColor Cyan
        #test if mail enabled first
        if(!(Get-Mailbox -Identity $user.targetSAM -ErrorAction SilentlyContinue))
        {
            Enable-Mailbox -identity $user.targetSAM -database $user.targetExchDB -ErrorAction Stop | Out-Null
            Write-Host "TARGET\$($user.targetSAM) was successfully Mail enabled on DAG005" -ForegroundColor Green
            Write-This "$(Get-Date -DisplayHint Time -uformat %T) - TARGET\$($user.targetSAM) was successfully Mail enabled on target DAG." $Log
        }
        else
        {
            Write-Host "TARGET\$($user.targetSAM) is already mail enabled." -ForegroundColor Green
            Write-This "$(Get-Date -DisplayHint Time -uformat %T) - TARGET\$($user.targetSAM) is already mail enabled." $Log
        }
    }
    catch
    {
        Write-Host $_.Exception.Message
        Write-This "$(Get-Date -DisplayHint Time -uformat %T) - TARGET\$($user.targetSAM) was not successfully mail enabled on target DAG: $($_.Exception.Message)" $log
        Remove-PSSession $TARGETPSSession
        closeScript 1
    }
    $i = $i + 1
    Write-Host
}

Remove-PSSession $TARGETPSSession

#Wait for TARGET FIM GALSync to create the new contact in Source (SOURCE) AD, verify one exists for each user before proceeding
Write-Host "*************************************************************************************" -ForegroundColor White
Write-Host "Phase 2 - Verifying mail contacts exist in SOURCE domain and adding mail forwarding..." -ForegroundColor White
Write-Host "*************************************************************************************" -ForegroundColor White
Write-Host

#sourcedomain.com Exchange Session
$SOURCEPSSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchangeserver.sourcedomain.com/powershell -Credential $SOURCECred
Import-PSSession $SOURCEPSSession -DisableNameChecking -AllowClobber | Out-Null

$i = 1
foreach($user in $users)
{
    try
    {
        #Get new targetdomain.com Mail contact in SOURCE and set SOURCE mailbox to deliver to mailbox AND forward to targetdomain.com contact
        Write-Host "Processing" $i "of" $users.Count "Users..." -ForegroundColor White
        Write-Host
        Write-Host "Waiting for contact $($user.SourceName) to populate in the SOURCE domain..." -ForegroundColor Cyan
        $contactPresent = $false
        while($contactPresent -ne $true)
        {
            Start-Sleep -Seconds 3
            $testContact = Get-MailContact -Identity $user.TargetSAM -ErrorAction SilentlyContinue | Where-Object {$_.primarySMTPAddress -like "$($user.TargetSAM)@targetdomain.com"} 
            if($testContact)#Get-MailContact -Identity $user.TargetSAM -ErrorAction SilentlyContinue | Where-Object {$_.primarySMTPAddress -like "$($user.TargetSAM)@targetdomain.com"}) #| Where-Object {$_.primarySMTPAddress -like "$($user.TargetSAM)@targetdomain.com"}
            {
                $contactPresent = $true
                $mailContactIdentity = Get-MailContact -Identity $user.TargetSAM | Where-Object {$_.primarySMTPAddress -like "$($user.TargetSAM)@targetdomain.com"} -ErrorAction SilentlyContinue | Select Identity
                Write-Host "TARGET\$($user.TargetSAM) now has a contact in the SOURCE source domain: $($mailContactIdentity.Identity)" -ForegroundColor Green
                Write-This "$(Get-Date -DisplayHint Time -uformat %T) - TARGET\$($user.TargetSAM) now has an TARGET contact in source domain: $($mailContactIdentity.Identity)" $Log
                if($i -eq 1)
                {
                    Start-Sleep -Seconds 120 #initial user only, wait for exchange to catch up to AD repl
                }
                Set-Mailbox $user.sourcename -forwardingaddress $mailContactIdentity.Identity -DeliverToMailboxAndForward $true -ErrorAction Stop | Out-Null
                Write-Host "TARGET\$($user.TargetSAM) now has mail forwarding enabled from SOURCE to TARGET: $($mailContactIdentity.Identity)" -ForegroundColor Green
                Write-This "$(Get-Date -DisplayHint Time -uformat %T) - TARGET\$($user.TargetSAM) now has mail forwarding enabled from SOURCE to TARGET: $($mailContactIdentity.Identity)" $Log
            }
        }
    }
    catch
    {
        Write-Host "Failed to configure mail forwarding for TARGET\$($user.TargetSAM): $($_.Exception.Message)" -ForegroundColor Red
        Write-This "$(Get-Date -DisplayHint Time -uformat %T) - Failed to configure mail forwarding for TARGET\$($user.TargetSAM): $($_.Exception.Message)"  $log
        Remove-PSSession $SOURCEPSSession
        closeScript 1
    }
    $i = $i + 1
    Write-Host
}

#Remove-PSSession $SOURCEPSSession - do not remove session as we will use for next phase

#Export SOURCE mailbox to PST file
Write-Host "*************************************************************************************" -ForegroundColor White
Write-Host "Phase 3 - Export SOURCE Exchange Mailboxes to PST..." -ForegroundColor White
Write-Host "*************************************************************************************" -ForegroundColor White
Write-Host

$i = 1
foreach($user in $users)
{
    try
    {
        Write-Host "Processing" $i "of" $users.Count "Users..." -ForegroundColor White
        Write-Host
        Write-Host "Exporting mailbox to pst for user: SOURCE\$($user.sourceName)..." -ForegroundColor Cyan
        New-MailboxExportRequest -mailbox $user.sourcename -filepath "\\netapp-snap.sourcedomain.com\Migration\$($user.sourcename).pst" -BadItemLimit 50 -AcceptLargeDataLoss -ErrorAction Stop | Out-Null
        Write-Host "NSWDG\$($user.sourceName) has been queued for export to: \\netapp-snap.sourcedomain.com\Migration\$($user.sourcename).pst" -ForegroundColor Green
        Write-This "$(Get-Date -DisplayHint Time -uformat %T) - SOURCE\$($user.sourceName) has been queued for export: \\netapp-snap.sourcedomain.com\Migration\$($user.sourcename).pst" -ForegroundColor Green $Log
    }
    catch
    {
        Write-Host "Error exporting SOURCE mailbox to pst for user: SOURCE\$($user.sourcename): $($_.Exception.Message)"
        Write-This "$(Get-Date -DisplayHint Time -uformat %T) - Error exporting SOURCE mailbox to pst for user: SOURCE\$($user.sourcename): $($_.Exception.Message)" $log
        Remove-PSSession $SOURCEPSSession
        closeScript 1
    }
    $i = $i + 1
    Write-Host
}

Remove-PSSession $SOURCEPSSession

#Halfway there - Wait for all queued mailbox export requests to finish before executing migrate2.ps1
closescript 0
