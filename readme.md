
# HelloID-Task-SA-Target-ExchangeOnPremises-DistributionGroupRevokeMembership

## Prerequisites
Before using this snippet, verify you've met with the following requirements:
- [ ] User defined variables: `$ExchangeAdminUsername`, `$ExchangeAdminPassword` and `$ExchangeConnectionUri` created in your HelloID portal. [See also Custom Variables](https://docs.helloid.com/en/variables/custom-variables.html)


## Description

This code snippet executes the following tasks:

1. Define a hash table `$formObject`. The keys of the hash table represent the properties of the `Remove-DistributionGroupMember` cmdlet, while the values represent the values entered in the form.

> To view an example of the form output, please refer to the JSON code pasted below.

```json
{
    "GroupIdentity": "TestDistributionGroup",
    "Users": [
        {
            "UserPrincipalName": "jan@connectors.com"
        },
        {
            "UserPrincipalName": "koppelaar@connectors.com"
        }
    ]
}
```

> :exclamation: It is important to note that the names of your form fields might differ. Ensure that the `$formObject` hashtable is appropriately adjusted to match your form fields.
> The **GroupIdentity** can hold different values [See the Microsoft Docs page](https://learn.microsoft.com/en-us/powershell/module/exchange/remove-distributiongroupmember?view=exchange-ps)

2. Constructs a PowerShell credential object from the supplied administrative username and password

3. Connects with the credentials to the Exchange on premises environment by means of the `New-PSSession` cmdlet

4. Calls the `Remove-DistributionGroupMember` cmdlet to Remove the users from the DistributionGroup

5. Disconnects from the Exchange session by means of the `Remove-PsSession` cmdlet
