# Intune Group Policy Preferences Utility

## About
Export an existing Group Policy that contains Group Policy Preferences to a JSON file. Then use the JSON file as a rules list for PowerShell to process entries just like GPP does.

## Syntax

``` 
#Generate JSON Rules File(s)
.\Admin\Convert-GPPtoJSON -GPOName "My GPO" -ExportPath ".\Rules"

#Process the rules file on a client
Process-RegistryPreferenceRules -GPPLogKeyPath "HKLM:\SOFTWARE\ASD\GPP" -Path ".\Rules" -RulesFileName "My GPO.JSON"

```