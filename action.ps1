# HelloID-Task-SA-Target-ExchangeOnPremises-MailboxUnhideFromAddresslist
########################################################################
# Form mapping
$formObject = @{
    DisplayName     = $form.MailBoxDisplayName
    MailboxIdentity = $form.MailboxIdentity
}


[bool]$IsConnected = $false
try {
    Write-Information "Executing ExchangeOnPremises action: [MailboxUnhideFromAddresslist] for: [$($formObject.DisplayName)]"
    $adminSecurePassword = ConvertTo-SecureString -String $ExchangeAdminPassword -AsPlainText -Force
    $adminCredential = [System.Management.Automation.PSCredential]::new($ExchangeAdminUsername, $adminSecurePassword)
    $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck
    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangeConnectionUri -Credential $adminCredential -SessionOption $sessionOption -Authentication Kerberos  -ErrorAction Stop
    $null = Import-PSSession $exchangeSession -DisableNameChecking -AllowClobber -CommandName 'Set-Mailbox'
    $IsConnected = $true

    $null = Set-Mailbox -Identity $formObject.MailboxIdentity -HiddenFromAddressListsEnabled $false -ErrorAction Stop

    $auditLog = @{
        Action            = 'UpdateResource'
        System            = 'ExchangeOnPremises'
        TargetIdentifier  = $formObject.MailboxIdentity
        TargetDisplayName = $formObject.DisplayName
        Message           = "ExchangeOnPremises action: [MailboxUnhideFromAddresslist] for: [$($formObject.DisplayName)] executed successfully"
        IsError           = $false
    }
    Write-Information -Tags 'Audit' -MessageData $auditLog
    Write-Information "ExchangeOnPremises action: [MailboxUnhideFromAddresslist] for: [$($formObject.DisplayName)] executed successfully"
} catch {
    $ex = $_
    $auditLog = @{
        Action            = 'UpdateResource'
        System            = 'ExchangeOnPremises'
        TargetIdentifier  = $formObject.MailboxIdentity
        TargetDisplayName = $formObject.DisplayName
        Message           = "Could not execute ExchangeOnPremises action: [MailboxUnhideFromAddresslist] for: [$($formObject.DisplayName)], error: $($ex.Exception.Message)"
        IsError           = $true
    }
    Write-Information -Tags 'Audit' -MessageData $auditLog
    Write-Error "Could not execute ExchangeOnPremises action: [MailboxUnhideFromAddresslist] for: [$($formObject.DisplayName)], error: $($ex.Exception.Message)"
} finally {
    if ($IsConnected) {
        Remove-PSSession -Session $exchangeSession -Confirm:$false  -ErrorAction Stop
    }
}
########################################################################
