function Connect-O365{
    #Uses Function to installed Exchange Online Management
    try {
        Disconnect-ExchangeOnline -Confirm:$false
        Connect-ExchangeOnline -ShowBanner:$false
        Write-Host "You have successfully connected to Exchange Online"
    catch {
        Write-Host "
        Please check that you have installed the correct modules for this script to work. 
        
            Try using the command: Install-O365Modules
        "
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

function Get-Licences {
    Get-MsolAccountSku
}

function Get-Mailboxes {
    param(
        [string]$value
    )

    Get-Exomailbox -ResultSize unlimited | where Alias -like "*$($value)*"   | Format-List -Property DisplayName, UserPrincipalName, RecipientType, RecipientTypeDetails

}

function createFolderPath {

        #Folder Path
        $folder = Test-Path -Path "C:\365Module"
    
        #Statement to check Paths and create a folder if it doesn't exist
        if(!$folder){
            new-Item -Path "C:\" -Name "365Module" -ItemType "directory"
            Write-Output "Path Created"
        }else{
            Write-Output "Path directory already exists"
        }
    
}

function Get-DelegateFileDownload {

    createFolderPath

    $delegateFile = Test-Path -Path "C:\365Module\delegatePermissions.csv"
    
    #Download 365 Delegate.csv file
    if(!$delegate){
        Write-Output "Downloading delegatePermissions.csv"
        (New-Object System.Net.WebClient).DownloadFile("https://raw.githubusercontent.com/agukbiz2988/365-Module/refs/heads/main/delegatePermissions.csv", "C:\365Module\delegatePermissions.csv")
    }else{
        Write-Output "delegatePermissions.csv Already Exists"
    }

}

# function Set-DelegatePermissions {

#     $filePath = "C:\365Module\delegatePermissions.csv"
#     $list = Import-CSV $filePath

#     foreach($user in $list){

#         #FullAccess Permissions
#         if($list.FullAccess -eq "x") {
#            try {
#                 write-host "`nAdding FullAccess Permissions for $($user.User) to $($user.Mailbox)" -ForegroundColor Yellow
#                 Add-MailboxPermission -Identity $user.Mailbox -User $user.User -AccessRights FullAccess -InheritanceType All -Confirm:$false
#                 Write-Host "Complete" -ForegroundColor Green
#             }
#             catch {
#                 Write-Error -Message "Error occurred while adding FullAccess permissions for $($user.User) to $($user.Mailbox)"
#             }
#         }

#         #SendAs Permissions
#         if($list.SendAs -eq "x") {
#             try {
#                 write-host "`nAdding SendAS Permissions for $($user.User) to $($user.Mailbox)" -ForegroundColor Yellow
#                 Add-RecipientPermission -Identity $user.Mailbox -Trustee $user.User -AccessRights SendAs -Confirm:$false
#                 Write-Host "Complete" -ForegroundColor Green
#             }
#             catch {
#                 Write-Error -Message "Error occurred while adding SendAs permissions for $($user.User) to $($user.Mailbox)"
#             }
#         }
#     }
# }

# function Remove-DelegatePermissions {

#     $filePath = "C:\365Module\delegatePermissions.csv"
#     $list = Import-CSV $filePath

#     foreach($user in $list){

#         #FullAccess Permissions
#         if($list.FullAccess -eq "x") {
#            try {
#                 write-host "`nRemoving FullAccess Permissions for $($user.User) to $($user.Mailbox)" -ForegroundColor Yellow
#                 Remove-MailboxPermission -Identity $user.Mailbox -User $user.User -AccessRights FullAccess -InheritanceType All -Confirm:$false
#                 Write-Host "Complete" -ForegroundColor Green
#             }
#             catch {
#                 Write-Error -Message "Error occurred while Removing FullAccess permissions for $($user.User) to $($user.Mailbox)"
#             }
#         }

#         #SendAs Permissions
#         if($list.SendAs -eq "x") {
#             try {
#                 write-host "`nRemoving SendAS Permissions for $($user.User) to $($user.Mailbox)" -ForegroundColor Yellow
#                 Remove-RecipientPermission -Identity $user.Mailbox -Trustee $user.User -AccessRights SendAs -Confirm:$false
#                 Write-Host "Complete" -ForegroundColor Green
#             }
#             catch {
#                 Write-Error -Message "Error occurred while Removing SendAs permissions for $($user.User) to $($user.Mailbox)"
#             }
#         }
#     }
# }

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

function Get-AllCommands{
    Write-Host "
        Connect-O365
        Update-O365Modules
        Get-Mailboxes
        Get-Licences
        Get-DelegateFileDownload
        Update-DelegatePermissions
        Get-O365Help
    "
}


Write-Host "
Office 365 Exchange Online Module Connected

To Get all Commands Type Get-AllCommands"

$(Get-AllCommands)