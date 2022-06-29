$AdminAccount = Read-Host "Please enter the O365 Admin email address" #This prompts the user to type in their O365 admin account
Write-Host "Please press Enter..."
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown'); #I put this in here because for some reason it was requiring Enter to be pressed twice and I wanted people to understand
Connect-ExchangeOnline -UserPrincipalName $AdminAccount #Connects the user to Exchange Online using the email address provided. They will get a pop-up prompt for password and MFA where applicable.

do {
    $TargetCalendar = Read-Host "Please enter the email address of the calendar that is BEING SHARED"` #Prompts the user for the email address being shared

    Write-Host 'Press any key to continue...'; `
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown'); #The prompts for Press any key are put in just for user confirmation and ability to close script if no longer necessary

    $DelegateAccount = Read-Host "Please enter the email address of the user GETTING ACCESS to the calendar"` #Prompts the user for the email address that is getting access

    Write-Host 'Press any key to continue...'; `
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown'); #Another confirmation prompt, and opportunity to easily close the script

    Add-MailboxFolderPermission -Identity ${TargetCalendar}:\Calendar -User $DelegateAccount -AccessRights Editor -SharingPermissionFlags Delegate #Runs the following command based on the input provided by the user.

    Write-Host 'Press any key to continue...'; `
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown'); #Another confirmation prompt, and opportunity to easily close the script

    Write-Output 'Access to the calendar of '${TargetCalendar}' has been given to '${DelegateAccount}''`} #Visual confirmation for the user that the permissions have been added.

    function Show-Menu { #The following creates a visual menu showing the user options to repeat the commands for more/different delegates
        param (
            [string]$Title = 'Confirmation Menu'
        )
        Clear-Host
        Write-Host "================ $Title ================" 

        Write-Host "Do you need to add additional delegates?"
        Write-Host "1: Press '1' to add more delegates."
        Write-Host "2: Press '2' to quit."
    }
    Show-Menu
    $selection = Read-Host "Please make your selection" #The following waits for user input on whether to repeat or not. If option 1 is selected it will go back to line 7; If option 2 is selected the script will end. 
    switch ($selection) {
        '1' {
            'Adding additional delegates'
        } '2' {
            'Closing script.'
        }
    }
    pause
} until ($selection -eq '2' )
