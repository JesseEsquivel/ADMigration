<#
##################################################################################################################
#
# Microsoft Premier Field Engineering
# 
# Migrate2.ps1
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

##################################################################################################################
# Prerequisites  - please do not change unless you know what you are doing
##################################################################################################################

#cred requires targetdomain.com Exchange AND SKYPE rights
$TARGETCred = Get-Credential -Message "Enter your TARGET General Admin Credential."

#cred requires sourcedomain.com Exchange rights
$SOURCECred = Get-Credential -Message "Enter your SOURCE ADM Credential."

##################################################################################################################
# Begin Script  - please do not change unless you know what you are doing
##################################################################################################################

$elapsed = [System.Diagnostics.Stopwatch]::StartNew()
StartScript

#Import SOURCE PST file to TARGET mailbox
Write-Host "*************************************************************************************" -ForegroundColor White
Write-Host "Phase 4 - Import SOURCE PST file to TARGET mailbox..." -ForegroundColor White
Write-Host "*************************************************************************************" -ForegroundColor White
Write-Host

#targetdomain.com Exchange Session
$TARGETPSSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchangeserver.targetdomain.com/powershell -Credential $TARGETCred
Import-PSSession $TARGETPSSession -DisableNameChecking -AllowClobber | Out-Null

$i = 1
foreach($user in $users)
{
    try
    {
        Write-Host "Processing" $i "of" $users.Count "Users..." -ForegroundColor White
        Write-Host
        Write-Host "Importing mailbox from SOURCE pst for TARGET user: $($user.targetSAM)..." -ForegroundColor Cyan
        New-MailboxImportRequest -mailbox $user.targetSAM -filepath "\\netapp-snap.targetdomain.com\Migration\$($user.sourcename).pst" -BadItemLimit 50 -AcceptLargeDataLoss -ErrorAction Stop | Out-Null
        Write-Host "TARGET\$($user.targetSAM) has been queued for import from: \\netapp-snap.targetdomain.com\Migration\$($user.sourcename).pst" -ForegroundColor Green
        Write-This "$(Get-Date -DisplayHint Time -uformat %T) - TARGET\$($user.targetSAM) mail has been queued for import from: \\netapp-snap.targetdomain.com\Migration\$($user.sourcename).pst" -ForegroundColor Green $Log
    }
    catch
    {
        Write-Host "Error importing mailbox from SOURCE pst for TARGET user: $($user.targetSAM):" $_.Exception.Message -ForegroundColor Red
        Write-This "$(Get-Date -DisplayHint Time -uformat %T) - Error importing mailbox from SOURCE pst for TARGET user: $($user.targetSAM): $($_.Exception.Message)" $log
        Remove-PSSession $TARGETPSSession
        closeScript 1
    }
    $i = $i + 1
    Write-Host
}

Remove-PSSession $TARGETPSSession

#Hide SOURCE user from GAL
Write-Host "*************************************************************************************" -ForegroundColor White
Write-Host "Phase 5 - Hiding SOURCE Users from GAL..." -ForegroundColor White
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
        Write-Host "Processing" $i "of" $users.Count "Users..." -ForegroundColor White
        Write-Host
        Write-Host "Hiding SOURCE\$($user.SourceName) from SOURCE Global Address List..." -ForegroundColor Cyan
        Set-Mailbox $user.sourcename -HiddenFromAddressListsEnabled $true -ErrorAction Stop | Out-Null
        Write-Host "Success." -ForegroundColor Green
        Write-This "$(Get-Date -DisplayHint Time -uformat %T) - SOURCE\$($user.SourceName) is now hidden from SOURCE GAL." $Log
    }
    catch
    {
        Write-Host "Error hiding SOURCE\$($user.SourceName) from SOURCE Global Address List:" $_.Exception.Message -ForegroundColor Red
        Write-This "$(Get-Date -DisplayHint Time -uformat %T) - Error hiding SOURCE\$($user.SourceName) from SOURCE Global Address List: $($_.Exception.Message)" $log
        Remove-PSSession $SOURCEPSSession
        closeScript 1
    }
    $i = $i + 1
    Write-Host
}

Remove-PSSession $SOURCEPSSession

#SIP enable user accounts in TARGET
Write-Host "*************************************************************************************" -ForegroundColor White
Write-Host "Phase 6 - SIP Enabling Users in targetdomain.com..." -ForegroundColor White
Write-Host "*************************************************************************************" -ForegroundColor White
Write-Host

#targetdomain.com Skype session
$TARGETSkypeSession = New-PSSession -ConnectionUri "https://skypeserver.targetdomain.com/ocspowershell" -Credential $TARGETCred
Import-PSSession $TARGETSkypeSession -DisableNameChecking -AllowClobber | Out-Null

$i = 1
foreach($user in $users)
{
    try
    {
        Write-Host "Processing" $i "of" $users.Count "Users..." -ForegroundColor White
        Write-Host
        Write-Host "SIP Enabling TARGET\$($user.targetSAM) on targetdomain.com SKYPE..." -ForegroundColor Cyan
        #test to see if user is already sip enabled and if not sip enable
        if(!(Get-CSUser -Identity $user.targetSAM -DomainController DC01.targetdomain.com -ErrorAction SilentlyContinue))
        {
            Enable-CsUser -Identity $user.targetSAM -RegistrarPool N005SVSKYPE01.jiant.net -sipaddressType SamAccountName -SipDomain soc.smil.mil -DomainController DC01.targetdomain.com -ErrorAction Stop | Out-Null
            Grant-CsClientPolicy -Identity $user.targetSAM -PolicyName "Online Address Book" -DomainController DC01.targetdomain.com -ErrorAction Stop | Out-Null
            Write-Host "Success." -ForegroundColor Green
            Write-This "$(Get-Date -DisplayHint Time -uformat %T) - TARGET\$($user.targetSAM) is now skype enabled." $Log
        }
        else
        {
            Write-Host "TARGET$($user.targetSAM) is already SIP Enabled." -ForegroundColor Green
            Write-This "$(Get-Date -DisplayHint Time -uformat %T) - TARGET\$($user.targetSAM) $($user.targetSAM) is already SIP Enabled." $Log
        }
    }
    catch
    {
        Write-Host "Error SIP enabling user TARGET\$($user.targetSAM):" $_.Exception.Message -ForegroundColor Red
        Write-This "$(Get-Date -DisplayHint Time -uformat %T) - Error SIP enabling user TARGET\$($user.targetSAM): $($_.Exception.Message)" $log
        Remove-PSSession $TARGETSkypeSession
        closeScript 1
    }
    $i = $i + 1
    Write-Host
}

Remove-PSSession $TARGETSkypeSession

#Re-ACL, rename home share, and set home directory attributes on TARGET account
Write-Host "*************************************************************************************" -ForegroundColor White
Write-Host "Phase 7 - Re-configuring user home shares for TARGET access..." -ForegroundColor White
Write-Host "*************************************************************************************" -ForegroundColor White
Write-Host

$i = 1
foreach($user in $users)
{
    try
    {
        Write-Host "Processing" $i "of" $users.Count "Users..." -ForegroundColor White
        Write-Host
        Write-Host "Re-configuring home share for TARGET\$($user.targetSAM)..." -ForegroundColor Cyan
        if(Test-Path -Path "\\netapp-cifs.sourcedomain.com\home_folders\$($user.SourceName)")
        {
            #get the current ACL
            $acl = ((Get-Item "\\netapp-cifs.sourcedomain.com\home_folders\$($user.SourceName)").GetAccessControl('Access'))
            #build the new ACE granting full control to the TARGET identity
            $ace =  New-Object System.Security.AccessControl.FileSystemAccessRule("TARGET\$($user.targetSAM)",'FullControl','ContainerInherit,ObjectInherit','None','Allow')
            #add the new ACE to the ACL
            $acl.SetAccessRule($ace)
            #write the changes to the folder
            Set-Acl -Path "\\netapp-cifs.sourcedomain.com\home_folders\$($user.SourceName)" -AclObject $acl
            #rename directory to TARGET target SamAccountName
            Rename-Item -Path "\\netapp-cifs.sourcedomain.com\home_folders\$($user.SourceName)" -NewName "\\netapp-cifs.sourcedomain.com\home_folders\$($user.targetSAM)"
            #set home share attribute on TARGET target account
            Set-ADUser -Identity $user.targetSAM -Server DC01.targetdomain.com -HomeDrive N: -HomeDirectory "\\sourcedomain.com\source\home folders\$($user.targetSAM)" -Credential $TARGETCred
            Write-Host "Success." -ForegroundColor Green
            Write-This "$(Get-Date -DisplayHint Time -uformat %T) - User TARGET\$($user.targetSAM) now has access to \\netapp-cifs.sourcedomain.com\home folders\$($user.targetSAM)" $Log
        }
        else
        {
            Write-Host "Source Home Folder does not exist for SOURCE\$($user.SourceName), skipping." -ForegroundColor Yellow
            Write-This "$(Get-Date -DisplayHint Time -uformat %T) - Source Home Folder does not exist for SOURCE\$($user.SourceName), skipping." $Log
        }
    }
    catch
    {
        Write-Host "Failed to re-configure home share access for TARGET\$($user.targetSAM):" $_.Exception.Message -ForegroundColor Red
        Write-This "$(Get-Date -DisplayHint Time -uformat %T) - Failed to re-configure home share access for TARGET\$($user.targetSAM): $($_.Exception.Message)" $log
        closeScript 1
    }
    $i = $i + 1
    Write-Host
}

#disable and move source AD user account objects
Write-Host "*************************************************************************************" -ForegroundColor White
Write-Host "Phase 8 - Disable and move source AD objects..." -ForegroundColor White
Write-Host "*************************************************************************************" -ForegroundColor White
Write-Host

$i = 1
foreach($user in $users)
{
    try
    {
        Write-Host "Processing" $i "of" $users.Count "Users..." -ForegroundColor White
        Write-Host
        Write-Host "Disabling and moving SOURCE\$($user.SourceName)..." -ForegroundColor Cyan
        #for certain people only - remove expiration and re-enable disabling of AD account for all others
        #Set-ADAccountExpiration -Identity $user.SourceName -DateTime "07/19/2017" -Server dc01.sourcedomain.com -ErrorAction Stop
        Disable-ADAccount -Identity $user.SourceName -Server dc01.sourcedomain.com -ErrorAction Stop
        Get-ADuser -Identity $user.SourceName | Move-ADObject -Server dc01.sourcedomain.com -TargetPath "OU=Migrated,OU=Disabled,OU=Users,DC=sourcedomain,DC=com" -ErrorAction Stop
        Write-Host "Success." -ForegroundColor Green
        Write-This "$(Get-Date -DisplayHint Time -uformat %T) - Successfully disabled and moved SOURCE\$($user.SourceName) to OU=Migrated,OU=Disabled,OU=Users,OU=Account/Object Management,DC=SOURCEdomain,DC=com" $Log
        #Write-This "$(Get-Date -DisplayHint Time -uformat %T) - Successfully expired and moved SOURCE\$($user.SourceName) to OU=Migrated,OU=test,OU=SOURCE Users,DC=SOURCEdomain,DC=com" $Log
    }
    catch
    {
        Write-Host "Failed to disable and move AD object SOURCE\$($user.SourceName):" $_.Exception.Message -ForegroundColor Red
        Write-This "$(Get-Date -DisplayHint Time -uformat %T) - Failed to disable and move AD object SOURCE\$($user.SourceName): $($_.Exception.Message)" $log
        closeScript 1
    }
    $i = $i + 1
    Write-Host
}

# :) - Review the logs and verify.
closeScript 0