<#
.Synopsis
   Folder Archive, based on AD user availability. (Home drive or folder cleanup)
.DESCRIPTION
   Moves user home folders to destination folder after checking if user exists in AD.
   Additionally, emails listed admin on cleanup of achive folder at designated interval.
.EXAMPLE
   Example of how to use:  Set as scheduled task and run
   .\script.ps1
.COMMENT
   .Net Frmaework 4.5 required for system.io.compression.filesystem Tested and verified in Powershell v4.0 (en-US)
   NOTE: also be sure to update all parameters for your needs
#>

Param (
    [string]$Directory = "\\home$\USERS",
    [string]$MoveFrom = "\\home$\USERS",
    [string]$MoveTo = "\\C$\Test",
    [string]$Archive = "\\C$\Test\*",
    [string[]]$To = @("admin@admin.com"),
    [string]$From = "no-reply@admin.com",
    [string]$SMTPServer = "smtpserver",
    [string]$Logfile = "C:\Logs\ArchiveFolders.log"
    )
#Defines logging function called later on in script
Function LogWrite
{
   Param ([string]$logstring)

   Add-content $Logfile -value $logstring
}

#Imports AD Module for use
Try { Import-Module ActiveDirectory -ErrorAction Stop }
Catch { Write-Host "Unable to load Active Directory module, is RSAT installed?"; Break }
#Imports the System IO compression assembly needed for Zip files
Add-Type -assembly "system.io.compression.filesystem"

$HomeFolders = Get-ChildItem $Directory | ?{ $_.PSIsContainer }

#For each folder in Home Folders, Checks AD to see if the folder name matches a samAccountName attribute of a current user.  If not, it outputs the results.
$(ForEach ($User in $HomeFolders)
{
   #Try getting the user from home folder name
    $ADUser = Get-ADUser -Filter {SamAccountName -eq $User.Name}

    #Test to see if the user exists in AD
    If($ADUser)
    {
        #User Exists
        #Write-host "$($User.Name) Exists"
    }
    Else
    {
        LogWrite "$($User.Name) Does Not Exist, Compressing and Moving to Archive $(get-date -f MM-dd-yyyy_HH_mm_ss)" -Append
        $Source = "$MoveFrom\$($User.Name)"
        $Destination = "$MoveTo\$($User.Name).zip"
        If(Test-path $Destination) {Remove-item $Destination}
        [io.compression.zipfile]::CreateFromDirectory($Source,$Destination) 
        Remove-Item $Source -Recurse -Force

    } 
})

#----------------------------------------------------------------------------------
#Notify $To list of Folders being removed after 45 Days (15 Days' notice of removal)
#----------------------------------------------------------------------------------
$limit = (Get-Date).AddDays(-30)
$RetainLimit = (Get-Date).AddDays(-45)
$ExpiredArchives = Get-ChildItem -Path $Archive -Force | Where-Object { $_.PSIsContainer -and $_.CreationTime -lt $limit } 

If ($ExpiredArchives)
{   $Header = @"
<style>
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
</style>
"@
    $Pre = "Archived user folders older then 30 days"
    $BodyAddon = "*All archives listed below will be automatically removed after 15 days*"
    $Body = $ExpiredArchives | Select Name,LastWriteTime | ConvertTo-HTML -PreContent $Pre -Head $Header, $BodyAddon | Out-String
    $SMTPSettings = @{
        To = $To
        From = $From
        Subject = $Pre
        SMTPServer = $SMTPServer
    }
    Send-MailMessage @SMTPSettings -Body $Body -BodyAsHtml
}

#----------------------------------------------------------------------------------
#Clean Archive folders older than 45 Days (30 Days, 15 Days wait from notification)
#----------------------------------------------------------------------------------
$Now = Get-Date
$Days = "45"
$Extension = "*.zip"
#----- define LastWriteTime parameter based on $Days ---#
$LastWrite = $Now.AddDays(-$Days)
#----- get files based on lastwrite filter and specified folder ---#
$Files = Get-Childitem $Archive -Include $Extension -Recurse | Where {$_.LastWriteTime -le "$LastWrite"}
foreach ($File in $Files) 
    {
    try 
        {
        LogWrite "'nDeleting File: $File $(get-date -f MM-dd-yyyy_HH_mm_ss)"
        #Remove-Item $File.FullName | out-null
        }
    catch
        {
        LogWrite "'nError deleting file: $File"
        }
    }
