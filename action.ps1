# HelloID-Task-SA-Target-ExchangeOnPremises-DistributionGroupRevokeMembership
#############################################################################
# Form mapping
$formObject = @{
    GroupIdentity = $form.GroupIdentity
    UsersToRemove = [array]$form.Users
}

[bool]$IsConnected = $false
try {
    $adminSecurePassword = ConvertTo-SecureString -String $ExchangeAdminPassword -AsPlainText -Force
    $adminCredential = [System.Management.Automation.PSCredential]::new($ExchangeAdminUsername, $adminSecurePassword)
    $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck
    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangeConnectionUri -Credential $adminCredential -SessionOption $sessionOption -Authentication Kerberos  -ErrorAction Stop
    $null = Import-PSSession $exchangeSession -DisableNameChecking -AllowClobber -CommandName 'Remove-DistributionGroupMember'
    $IsConnected = $true

    foreach ($user in $formObject.UsersToRemove) {
        try {
            Write-Information "Executing ExchangeOnPremises action: [DistributionGroupRevokeMembership] user [$($user.UserPrincipalName)] from: [$($formObject.GroupIdentity)]"

            $null = Remove-DistributionGroupMember -Identity $formObject.GroupIdentity -Member $user.UserPrincipalName -Confirm:$false -ErrorAction Stop

            $auditLog = @{
                Action            = 'RevokeMembership'
                System            = 'ExchangeOnPremises'
                TargetIdentifier  = $formObject.GroupIdentity
                TargetDisplayName = $formObject.GroupIdentity
                Message           = "ExchangeOnPremises action: [DistributionGroupRevokeMembership] user [$($user.UserPrincipalName)] from: [$($formObject.GroupIdentity)] executed successfully"
                IsError           = $false
            }
            Write-Information -Tags 'Audit' -MessageData $auditLog
            Write-Information "ExchangeOnPremises action: [DistributionGroupRevokeMembership] user [$($user.UserPrincipalName)] from: [$($formObject.GroupIdentity)] executed successfully"

        } catch {
            $ex = $_
            if ($ex.CategoryInfo.Reason -eq 'MemberNotFoundException') {
                $auditLog = @{
                    Action            = 'RevokeMembership'
                    System            = 'ExchangeOnPremises'
                    TargetIdentifier  = $formObject.GroupIdentity
                    TargetDisplayName = $formObject.GroupIdentity
                    Message           = "ExchangeOnPremises action: [DistributionGroupRevokeMembership] user [$($user.UserPrincipalName)] from: [$($formObject.GroupIdentity)] executed successfully"
                    IsError           = $false
                }
                Write-Information -Tags 'Audit' -MessageData $auditLog
                Write-Information "ExchangeOnPremises action: [DistributionGroupRevokeMembership] user [$($user.UserPrincipalName)] from: [$($formObject.GroupIdentity)] executed successfully"
            } else {
                $auditLog = @{
                    Action            = 'RevokeMembership'
                    System            = 'ExchangeOnPremises'
                    TargetIdentifier  = $formObject.GroupIdentity
                    TargetDisplayName = $formObject.GroupIdentity
                    Message           = "Could not execute ExchangeOnPremises action: [DistributionGroupRevokeMembership] user [$($user.UserPrincipalName)] from: [$($formObject.GroupIdentity)], error: $($ex.Exception.Message)"
                    IsError           = $true
                }
                Write-Information -Tags 'Audit' -MessageData $auditLog
                Write-Error "Could not execute ExchangeOnPremises action: [DistributionGroupRevokeMembership] user [$($user.UserPrincipalName)] from: [$($formObject.GroupIdentity)], error: $($ex.Exception.Message)"
            }

        }
    }
} catch {
    $ex = $_
    $auditLog = @{
        Action            = 'RevokeMembership'
        System            = 'ExchangeOnPremises'
        TargetIdentifier  = $formObject.GroupIdentity
        TargetDisplayName = $formObject.GroupIdentity
        Message           = "Could not execute ExchangeOnPremises action: [DistributionGroupRevokeMembership] from: [$($formObject.GroupIdentity)], error: $($ex.Exception.Message)"
        IsError           = $true
    }
    Write-Information -Tags 'Audit' -MessageData $auditLog
    Write-Error "Could not execute ExchangeOnPremises action: [DistributionGroupRevokeMembership] from: [$($formObject.GroupIdentity)], error: $($ex.Exception.Message)"
} finally {
    if ($IsConnected) {
        Remove-PSSession -Session $exchangeSession -Confirm:$false  -ErrorAction Stop
    }
}
#############################################################################
