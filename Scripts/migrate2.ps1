<#
##################################################################################################################
#
# Microsoft Premier Field Engineering
# jesse.esquivel@microsoft.com
# Migrate2.ps1
# v1.0 Initial creation 05/29/2019 - Perform AD User Migration Tasks Part two
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
$Log = "$scriptDir\Phase2Log.txt"
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

function wackFolder($folderName)
{
    if(Test-Path -Path $folderName)
    {
        Remove-Item -Path $folderName -Force -Recurse -ErrorAction SilentlyContinue
    }

}

##################################################################################################################
# Prerequisites  - Place all infinity stones into infinity guantlet to prepare for migration
##################################################################################################################

#cred requires target Exchange org rights - admin forest used to administer target exchange
$EXCHADMFRSTCred = Get-Credential -Message "Enter your EXCHADMFRST ADM Credential."

#cred requires source domain Exchange Org rights - source forest exchange org
$SOURCEDOMCred = Get-Credential -Message "Enter your SOURCEDOM ADM Credential."

#cred requires EXCHFRST DA - where target exchange org is
$EXCHFRSTCred = Get-Credential -Message "Enter your EXCHFRST DA Credential"

#cred requires TARGETDOM DA - where users will live
$TARGETDOMCred = Get-Credential -Message "Enter your TARGETDOM DA Credential"

##################################################################################################################
# Begin Script  - please do not change unless you know what you are doing
##################################################################################################################

$elapsed = [System.Diagnostics.Stopwatch]::StartNew()
StartScript

#Resume Exchange Move-Requests manually FIRST and ensure mailboxes are cutover successfully before executing this script and ensure all have finished successfully!!!
Write-Host "*************************************************************************************" -ForegroundColor White
Write-Host "Phase 4 - Disconnect mailboxes from EXCHFRST users and re-connect to SOF users..." -ForegroundColor White
Write-Host "*************************************************************************************" -ForegroundColor White
Write-Host

#socrs.mil Exchange Session
$EXCHFRSTPSSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://someexchserver.exchfrst.com/powershell -Credential $EXCHADMFRSTCred
Import-PSSession $EXCHFRSTPSSession -DisableNameChecking -AllowClobber | Out-Null

$i = 1
foreach($user in $users)
{
    try
    {
        Write-Host "Processing" $i "of" $users.Count "Users..." -ForegroundColor White
        Write-Host
        $ADUserResult = Get-ADUser -identity $user.targetSAM -Properties msExchRecipientTypeDetails -Server "somedc.exchfrst.com" -ErrorAction Stop
        if($ADUserResult.msExchRecipientTypeDetails -ne "2")
        {
            Write-Host "Re-configuring linked mailbox for TARGET user: $($user.targetSAM)..." -ForegroundColor Cyan
            #disconnect mailbox from EXCHFRST user
            Disable-Mailbox -Identity  $user.primarySMTPAddress -Confirm:$false -DomainController somedc.exchfrst.com -ErrorAction Stop | Out-Null
            Write-Host "EXCHFRST\$($user.targetSAM) has been disconnected from mailbox successfully!" -ForegroundColor Green
            Write-This "$(Get-Date -DisplayHint Time -uformat %T) - EXCHFRST\$($user.targetSAM) has been disconnected from mailbox successfully" -ForegroundColor Green $Log
            #disable EXCHFRST account
            Set-ADUser -Identity $user.targetSAM -Server somedc.exchfrst.com -Enabled $false -Credential $EXCHFRSTCred -ErrorAction Stop | Out-Null
            Write-Host "EXCHFRST\$($user.targetSAM) account has been disabled!" -ForegroundColor Green
            #connect mailbox to SOF user - changed -linkedMasterAccount switch to name attribute!!! <-----
            Connect-Mailbox -Identity $user.displayName -User $user.displayName -database $user.targetDatabase -linkedDomainController "somedc.targetdom.com" -linkedMasterAccount $user.displayName -linkedCredential $TARGETDOMCred -Alias "$($user.targetSAM).$($user.persona)" -DomainController somedc.exchfrst.com -ErrorAction Stop | Out-Null
            Write-Host "TARGETDOM\$($user.targetSAM) has been connected to their linked mailbox successfully!" -ForegroundColor Green
            Write-This "$(Get-Date -DisplayHint Time -uformat %T) - TARGETDOM\$($user.targetSAM) has been connected to their linked mailbox successfully!" -ForegroundColor Green $Log
            #add sourcedom smtp address as secondary
       	    Set-Mailbox -Identity $user.targetSAM -EmailAddresses @{add="$($user.primarySMTPAddress)"} -DomainController somedc.exchfrst.com -ErrorAction Stop | Out-Null
            Write-Host "TARGETDOM\$($user.targetSAM) secondary SMTP address $($user.primarySMTPAddress) has been stamped!" -ForegroundColor Green
            Write-This "$(Get-Date -DisplayHint Time -uformat %T) - TARGETDOM\$($user.targetSAM) secondary SMTP address $($user.primarySMTPAddress) has been stamped!" -ForegroundColor Green $Log
        }
        else
        {
            Write-Host "User EXCHFRST\$($user.targetSAM) is already a linked mailbox!" -ForegroundColor Yellow
        }

    }
    catch
    {
        Write-Host "Error re-connecting mailbox to TARGET user: TARGETDOM\$($user.targetSAM):" $_.Exception.Message -ForegroundColor Red
        Write-This "$(Get-Date -DisplayHint Time -uformat %T) - Error re-connecting mailbox to TARGET user: TARGETDOM\$($user.targetSAM): $($_.Exception.Message)" $log
        Remove-PSSession $EXCHFRSTPSSession
        closeScript 1
    }
    $i = $i + 1
    Write-Host
}

Remove-PSSession $EXCHFRSTPSSession

#Re-ACL, rename home share, and set home directory attributes on TARGET account - COPY Home shares first for N64.
Write-Host "*************************************************************************************" -ForegroundColor White
Write-Host "Phase 5 - Re-configuring user home shares for TARGETDOM access..." -ForegroundColor White
Write-Host "*************************************************************************************" -ForegroundColor White
Write-Host

$i = 1
foreach($user in $users)
{
    try
    {
        Write-Host "Processing" $i "of" $users.Count "Users..." -ForegroundColor White
        Write-Host
        Write-Host "Re-configuring home share for TARGETDOM\$($user.targetSAM)..." -ForegroundColor Cyan
        if(Test-Path -Path "\\somesmbserver.sourcedom.com\HomeFolders$\$($user.SourceName)")
        {
            #get the current ACL
            $acl = ((Get-Item "\\somesmbserver.sourcedom.com\HomeFolders$\$($user.SourceName)").GetAccessControl('Access'))
            #build the new ACE granting full control to the TARGET identity
            $ace =  New-Object System.Security.AccessControl.FileSystemAccessRule("TARGETDOM\$($user.targetSAM)",'FullControl','ContainerInherit,ObjectInherit','None','Allow') -ErrorAction Stop
            #add the new ACE to the ACL
            $acl.SetAccessRule($ace)
            #write the changes to the folder
            Set-Acl -Path "\\somesmbserver.sourcedom.com\HomeFolders$\$($user.SourceName)" -AclObject $acl -ErrorAction Stop
            #rename directory to TARGET target SamAccountName
            Rename-Item -Path "\\somesmbserver.sourcedom.com\HomeFolders$\$($user.SourceName)" -NewName "\\somesmbserver.sourcedom.com\HomeFolders$\$($user.targetSAM)" -ErrorAction Stop
            #set home share attribute on target account
            Set-ADUser -Identity $user.targetSAM -Server somedc.targetdom.com -HomeDrive H: -HomeDirectory "\\somesmbserver.sourcedom.com\HomeFolders$\$($user.targetSAM)" -Credential $TARGETDOMCred -ErrorAction Stop
            Write-Host "Success." -ForegroundColor Green
            Write-This "$(Get-Date -DisplayHint Time -uformat %T) - User TARGETDOM\$($user.targetSAM) now has access to \\somesmbserver.sourcedom.com\HomeFolders$\$($user.targetSAM)" $Log
            wackFolder "\\somesmbserver.sourcedom.com\HomeFolders$\$($user.targetSAM)\UEMArchives"
            wackFolder "\\somesmbserver.sourcedom.com\HomeFolders$\$($user.targetSAM)\UEMbackups"
            wackFolder "\\somesmbserver.sourcedom.com\HomeFolders$\$($user.targetSAM)\UEMLogs"
            wackFolder "\\somesmbserver.sourcedom.com\HomeFolders$\$($user.targetSAM)\UEV"
        }
        else
        {
            Write-Host "Source Home Folder does not exist for SOURCEDOM\$($user.SourceName), skipping." -ForegroundColor Yellow
            Write-This "$(Get-Date -DisplayHint Time -uformat %T) - Source Home Folder does not exist for SOURCEDOM\$($user.SourceName), skipping." $Log
        }
    }
    catch
    {
        Write-Host "Failed to re-configure home share access for TARGETDOM\$($user.targetSAM):" $_.Exception.Message -ForegroundColor Red
        Write-This "$(Get-Date -DisplayHint Time -uformat %T) - Failed to re-configure home share access for TARGETDOM\$($user.targetSAM): $($_.Exception.Message)" $log
        closeScript 1
    }
    $i = $i + 1
    Write-Host
}

#disable and move source AD user account objects
Write-Host "*************************************************************************************" -ForegroundColor White
Write-Host "Phase 6 - Disable and move source AD objects..." -ForegroundColor White
Write-Host "*************************************************************************************" -ForegroundColor White
Write-Host

$i = 1
foreach($user in $users)
{
    try
    {
        Write-Host "Processing" $i "of" $users.Count "Users..." -ForegroundColor White
        Write-Host
        Write-Host "Disabling and moving SOURCEDOM\$($user.SourceName)..." -ForegroundColor Cyan
        Disable-ADAccount -Identity $user.SourceName -Server somedc.sourcedom.com -ErrorAction Stop -Credential $SOURCEDOMCred
        Get-ADuser -Identity $user.SourceName -Server somedc.sourcedom.com | Move-ADObject -Server somedc.sourcedom.com -TargetPath "OU=Migrated,OU=Disabled,OU=Users,DC=sourcedcom,DC=com" -Credential $SOURCEDOMCred -ErrorAction Stop
        Write-Host "Success." -ForegroundColor Green
        Write-This "$(Get-Date -DisplayHint Time -uformat %T) - Successfully disabled and moved SOURCEDOM\$($user.SourceName) to OU=Migrated,OU=Disabled,OU=Users,DC=sourcedcom,DC=com" $Log
    }
    catch
    {
        Write-Host "Failed to disable and move AD object SOURCEDOM\$($user.SourceName):" $_.Exception.Message -ForegroundColor Red
        Write-This "$(Get-Date -DisplayHint Time -uformat %T) - Failed to disable and move AD object SOURCEDOM\$($user.SourceName): $($_.Exception.Message)" $log
        closeScript 1
    }
    $i = $i + 1
    Write-Host
}

# :) - Review the logs and verify - All your Migrations are belong to us.
closeScript 0