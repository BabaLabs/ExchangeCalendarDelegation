    $AdminAccount = Read-Host "Please enter the O365 Admin email address"
        Write-Host "Please press Enter..."
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
            Connect-ExchangeOnline -UserPrincipalName $AdminAccount

do {
    $TargetCalendar = Read-Host "Please enter the email address of the calendar that is BEING SHARED"`

    Write-Host 'Press any key to continue...'; `
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');

    $DelegateAccount = Read-Host "Please enter the email address of the user GETTING ACCESS to the calendar"`

    Write-Host 'Press any key to continue...'; `
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');

        Add-MailboxFolderPermission -Identity ${TargetCalendar}:\Calendar -User $DelegateAccount -AccessRights Editor -SharingPermissionFlags Delegate

    Write-Host 'Press any key to continue...'; `
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');

    Write-Output 'Access to the calendar of '${TargetCalendar}' has been given to '${DelegateAccount}''`}

    function Show-Menu {
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
    $selection = Read-Host "Please make your selection"
    switch ($selection) {
        '1' {
            'Adding additional delegates'
        } '2' {
            'Closing script.'
        }
    }
    pause
} until ($selection -eq '2' )