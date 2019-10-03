<#
##################################################################################################################
#
# Microsoft Premier Field Engineering
# jesse.esquivel@microsoft.com
# Migrate1.ps1
# v1.0 Initial creation 05/16/2019 - Perform AD/Exchange User Migration Tasks
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
$Log = "$scriptDir\stage1Log.txt"
$Global:ErrorActionPreference = 'Stop'
$users = Import-Csv $scriptDir\users.csv

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
# Prerequisites  - Place all infinity stones into infinity guantlet to prepare for migration
##################################################################################################################

#target Exchange org rights
$EXCHADMFRSTCred = Get-Credential -Message "Enter your EXCHADMFRST ADM Credential."

#cred requires source domain Exchange rights
$SOURCEDOMCred = Get-Credential -Message "Enter your SOURCEDOM Credential."

#cred requires EXCHFRST DA
$EXCHFRSTCred = Get-Credential -Message "Enter your EXCHFRST DA Credential"

##################################################################################################################
# Begin Script  - please do not change unless you know what you are doing
##################################################################################################################

$elapsed = [System.Diagnostics.Stopwatch]::StartNew()
StartScript

#process all EXCHFRST users in Bulk to mail enable
Write-Host "*************************************************************************************" -ForegroundColor White
Write-Host "Phase 1 - Bulk Mail Enabling Users in EXCHFRST..." -ForegroundColor White
Write-Host "*************************************************************************************" -ForegroundColor White
Write-Host

#exchfrst.com Exchange Session
$EXCHFRSTPSSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://someexchserver.exchfrst.com/powershell -Credential $EXCHADMFRSTCred
Import-PSSession $EXCHFRSTPSSession -DisableNameChecking -AllowClobber | Out-Null

$i = 1
foreach($user in $users)
{
    try
    {
        Write-Host "Mail Enabling" $i "of" $users.Count "Users ..." -ForegroundColor Cyan
        $ADUserResult = Get-ADUser -identity $user.targetSAM -Properties msExchRecipientTypeDetails -Server "somedc.exchfrst.com" -ErrorAction Stop
        if($ADUserResult.msExchRecipientTypeDetails -ne "128")
        {
            Enable-MailUser -identity $user.primarySMTPAddress -DomainController somedc.exchfrst.com -externalEmailAddress $user.primarySMTPAddress -Alias "$($user.targetSAM).$($user.persona)" -ErrorAction Stop | Out-Null
            Write-Host "EXCHFRST$\($user.targetSAM) was mail enabled!" -ForegroundColor Green
            Write-This "$(Get-Date -DisplayHint Time -uformat %T) - EXCHFRST$\($user.targetSAM) was mail successfully mail enabled." $Log
        }
        else
        {
            Write-Host "EXCHFRST$\($user.targetSAM) is already mail enabled from a previous run." -ForegroundColor Yellow
            Write-This "$(Get-Date -DisplayHint Time -uformat %T) - EXCHFRST$\($user.targetSAM) is already mail enabled from a previous run." $Log
        }
    }
    catch
    {
        Write-Host $_.Exception.Message -ForegroundColor Red
        Write-This "$(Get-Date -DisplayHint Time -uformat %T) - EXCHFRST$\($user.targetSAM) was not successfully mail enabled: $($_.Exception.Message)" $log
        Remove-PSSession $EXCHFRSTPSSession
        closeScript 1
    }
    $i = $i + 1
    Write-Host
}

#run exchange prepare move request script against all users
Write-Host "*************************************************************************************" -ForegroundColor White
Write-Host "Phase 2 - Run prepareMoveRequest script to prepare for users for mailbox move..." -ForegroundColor White
Write-Host "*************************************************************************************" -ForegroundColor White
Write-Host

$i = 1
foreach($user in $users)
{
    try
    {
        Write-Host "Preparing" $i "of" $users.Count "users to move..." -ForegroundColor Cyan
        F:\Exchange\Scripts\Prepare-MoveRequest.ps1 -Identity $user.primarySMTPAddress -RemoteForestDomainController somedc.sourcedom.com -RemoteForestCredential $SOURCEDOMCred -LocalForestDomainController somedc.exchfrst.com -LocalForestCredential $EXCHFRSTCred -UseLocalObject -OverwriteLocalObject -Verbose
        Write-This "$(Get-Date -DisplayHint Time -uformat %T) - Prepare-MoveRequest ran successfully against: SOURCEDOM\$($user.sourceName)" $Log
        Write-Host
        Write-Host
    }
    catch
    {
        Write-Host $_.Exception.Message -ForegroundColor Red
        Write-This "$(Get-Date -DisplayHint Time -uformat %T) - Prepare-MoveRequest did not run successfully against: SOURCEDOM\$($user.sourceName): $($_.Exception.Message)" $log
        Remove-PSSession $EXCHFRSTPSSession
        closeScript 1
    }
    $i = $i + 1
    Write-Host
}

#process all EXCHFRST users in Bulk to mail enable
Write-Host "*************************************************************************************" -ForegroundColor White
Write-Host "Phase 3 - Stamping @targetsmtp.com address on target mailusers..." -ForegroundColor White
Write-Host "*************************************************************************************" -ForegroundColor White
Write-Host

$i = 1
foreach($user in $users)
{
    try
    {
        Write-Host "Stamping @targetsmtp.com address on" $i "of" $users.Count "Mail Users ..." -ForegroundColor Cyan
        $mailUser = Get-MailUser -Identity $user.targetSAM -DomainController somedc.exchfrst.com -ErrorAction Stop
        if($mailUser.EmailAddresses -notlike "*smtp:$($user.targetMailAddress)*")
        {
            Do
            {
                $mailUser = Get-MailUser -Identity $user.targetSAM -DomainController somedc.exchfrst.com -ErrorAction Stop
                Set-MailUser -identity $user.targetSAM -EmailAddresses @{add="$($user.targetMailAddress)"} -DomainController somedc.exchfrst.com -ErrorAction Stop -EmailAddressPolicyEnabled $true | Out-Null
                #Set-MailUser -identity $user.targetSAM -EmailAddressPolicyEnabled $true -DomainController somedc.exchfrst.com -ErrorAction Stop | Out-Null
                Start-Sleep -Seconds 2
            }
            Until($mailUser.EmailAddresses -like "*smtp:$($user.targetMailAddress)*")
            Write-Host "EXCHFRST$($user.targetSAM) SOCOM Email address stamped to target mailuser and policy enabled!" -ForegroundColor Green
            Write-This "$($user.targetSAM) mailuser was successfully stamped with @targetsmtp.com address: $($user.targetMailAddress)" $Log
        }
    }
    catch
    {
        Write-Host $_.Exception.Message -ForegroundColor Red
        Write-This "$(Get-Date -DisplayHint Time -uformat %T) - $($user.targetSAM) mailuser was NOT successfully stamped with @targetsmtp.com address $($user.targetMailAddress): $($_.Exception.Message)" $Log
        Remove-PSSession $EXCHFRSTPSSession
        closeScript 1
    }
    $i = $i + 1
    Write-Host
}

#submit mailbox move requests to exchange
Write-Host "*************************************************************************************" -ForegroundColor White
Write-Host "Phase 4 - Submit New Mailbox Move Requests to Exchange..." -ForegroundColor White
Write-Host "*************************************************************************************" -ForegroundColor White
Write-Host

$i = 1
foreach($user in $users)
{
    try
    {
        Write-Host "Processing" $i "of" $users.Count "Mailbox Move Requests..." -ForegroundColor Cyan
        New-MoveRequest -identity $user.primarySMTPAddress -TargetDatabase $user.targetDatabase -SuspendWhenReadyToComplete -Allowlargeitems -Remote -RemoteGlobalCatalog somedc.sourcedom.com -RemoteCredential $SOURCEDOMCred -DomainController somedc.exchfrst.com -BadItemLimit 10 -RemoteHostName mail.sourcedom.com -TargetDeliveryDomain targetsmtp.com -BatchName "SOURCEDOM MIGRATION - DO NOT RESUME" -ErrorAction Stop | Out-Null
        Write-Host "Mailbox move for user SOURCEDOM\$($user.SourceName) initiated Successfully!" -ForegroundColor Green
        Write-This "$(Get-Date -DisplayHint Time -uformat %T) - Mailbox move for user SOURCEDOM\$($user.sourceName) initiated Successfully!" $Log
    }
    catch
    {
        Write-Host $_.Exception.Message -ForegroundColor Red
        Write-This "$(Get-Date -DisplayHint Time -uformat %T) - New mailbox move request failed for: SOURCEDOM\$($user.sourceName): $($_.Exception.Message)" $log
        Remove-PSSession $EXCHFRSTPSSession
        closeScript 1
    }
    $i = $i + 1
    Write-Host
}

Remove-PSSession $EXCHFRSTPSSession
closescript 0