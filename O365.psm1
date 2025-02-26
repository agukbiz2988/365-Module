function Connect-O365{
    #Uses Function to installed Exchange Online Management
    try {
        Disconnect-ExchangeOnline -Confirm:$false
        Connect-ExchangeOnline -ShowBanner:$false
        Write-Host "You have successfully connected to Exchange Online" -ForegroundColor Green
    } catch {
        Write-Host "
        Please check that you have installed the correct modules for this script to work. 
        
            Try using the command: Install-O365Modules
        " -ForegroundColor Red
    }
}

function Install-O365Modules{
    Install-Module ExchangeOnlineManagement -Force
}

function Get-ModuleCommands{
    Command -Module ExchangeOnlineManagement
}

function Update-O365Modules{
    Update-Module -name powershellget -allowprerelease -force
    update-Module -name ExchangeOnlineManagement -Force
}

function Get-Mailboxes {
    param(
        [string]$value
    )

    Get-Exomailbox -ResultSize unlimited | where Alias -like "*$($value)*"   | Format-List -Property DisplayName, UserPrincipalName, RecipientType, RecipientTypeDetails

}

function New-FolderPath {

        #Folder Path
        $folder = Test-Path -Path "C:\365Module"
    
        #Statement to check Paths and create a folder if it doesn't exist
        if(!$folder){
            new-Item -Path "C:\" -Name "365Module" -ItemType "directory"
            Write-Host "Path Created" -ForegroundColor Green
        }else{
            Write-Error "Path directory already exists"
        }
    
}

function Get-DelegateFileDownload {

    New-FolderPath

    $delegateFile = Test-Path -Path "C:\365Module\delegatePermissions.csv"
    
    #Download 365 Delegate.csv file
    if(!$delegateFile){
        Write-Host "`nDownloading delegatePermissions.csv"
        (New-Object System.Net.WebClient).DownloadFile("https://raw.githubusercontent.com/agukbiz2988/365-Module/refs/heads/main/delegatePermissions.csv", "C:\365Module\delegatePermissions.csv")
        Write-Host "`ndelegatePermissions.csv successfully downloaded" -ForegroundColor Green
    }else{
        Write-Error "delegatePermissions.csv Already Exists" 
    }

}


#This function can remove and add delegate permissions
#Command Example: Update-DelegatePermissions -choice 1
function Update-DelegatePermissions {
    
    param (
        [int]$choice
    )

    if($choice -eq 1){
        $addOrRemove = "Adding"
    }else{
        $addOrRemove = "Removing"
    }

    $filePath = "C:\365Module\delegatePermissions.csv"
    $list = Import-CSV $filePath

    foreach($user in $list){

        #FullAccess Permissions
        if($list.FullAccess -eq "x") {
           try {
                write-host "`n$($addOrRemove) FullAccess Permissions for $($user.User) to $($user.Mailbox)" -ForegroundColor Yellow
                if($choice -eq 1){
                    Add-MailboxPermission -Identity $user.Mailbox -User $user.User -AccessRights FullAccess -InheritanceType All -Confirm:$false
                }else{
                    Remove-MailboxPermission -Identity $user.Mailbox -User $user.User -AccessRights FullAccess -InheritanceType All -Confirm:$false
                }
                Write-Host "Complete`n" -ForegroundColor Green
            }
            catch {
                Write-Error -Message "Error occurred while $($addOrRemove) FullAccess permissions for $($user.User) to $($user.Mailbox)"
            }
        }

        #SendAs Permissions
        if($list.SendAs -eq "x") {
            try {
                write-host "`n$($addOrRemove) SendAS Permissions for $($user.User) to $($user.Mailbox)" -ForegroundColor Yellow
                if($choice -eq 1){
                    Add-RecipientPermission -Identity $user.Mailbox -Trustee $user.User -AccessRights SendAs -Confirm:$false
                }else{
                    Remove-RecipientPermission -Identity $user.Mailbox -Trustee $user.User -AccessRights SendAs -Confirm:$false
                }
                Write-Host "Complete`n" -ForegroundColor Green
            }
            catch {
                Write-Error -Message "Error occurred while $($addOrRemove) SendAs permissions for $($user.User) to $($user.Mailbox)"
            }
        }
    }
}

function New-Account{
    [CmdletBinding()]
    param(
        [parameter(Position=0,mandatory=$true)]
        [string]$firstname,
        [parameter(Position=1,mandatory=$true)]
        [string]$surname,
        [parameter(Position=2,mandatory=$true)]
        [string]$email
    )

    New-Mailbox -Name $firstname -FirstName $firstname -LastName $surname -DisplayName "$($firstname) $($surname)" -MicrosoftOnlineServicesID $email -Password (Read-Host "Enter password" -AsSecureString) -ResetPasswordOnNextLogon $false
}

Write-Host "
Office 365 Exchange Online Module Connected

To Get all Commands Type: Get-Command -Module O365"