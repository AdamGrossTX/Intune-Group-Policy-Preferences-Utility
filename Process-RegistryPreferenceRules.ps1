[cmdletbinding()]
param(
    $GPPLogKeyPath = "HKLM:\SOFTWARE\ASD\GPP",
    [System.IO.DirectoryInfo]$Path = ".\Rules",
    [string[]]$RulesFileName
)

#region Main
$Main = {
        #Create Logging RegKey to track changes
        $LogKey = Invoke-ProcessRegistryItem -Path $GPPLogKeyPath -Action 'C' -PropertyName '(Default)'

        #Get rules from JSON file(s)
        $RuleObj = Get-RegistryRuleList -Path $Path -RulesFileName $RulesFileName
        
        #Process rules
        if($RuleObj) {
        [string[]]$GUIDList = @()
        foreach($Key in $RuleObj.Keys) {
            foreach($Rule in $RuleObj[$key]) {
                $GUIDList += $Rule.GUID
                $RegKeySplat = @{
                    Group = $Key
                    GUID = [string]$Rule.GUID
                    Path = Join-Path -Path "Registry::$($Rule.hive)" -ChildPath $Rule.key
                    EntryName = $Rule.EntryName
                    PropertyName = if($Rule.default -eq 1) {"(Default)"} else {[string]$Rule.PropertyName} 
                    PropertyValue = [object]$Rule.PropertyValue
                    PropertyType = $RegistryValueKind["$($Rule.PropertyType)"]
                    Action = [string]$Rule.Action
                    RemovePolicy = if($Rule.RemovePolicy -eq 1) {$true} else {$false}
                    GPOSettingOrder = [int]$Rule.GPOSettingOrder
                    Desc = [string]$Rule.desc
                    BypassErrors = if($Rule.bypassErrors -eq 1) {$true} else {$false}
                    DisplayDecimal = if($Rule.displayDecimal -eq 1) {$true} else {$false}
                    Name = [string]$Rule.name
                    RunOnce = if($Rule.runOnce -eq 1) {$true} else {$false}
                    Filters = $Rule.Filters
                    LogKey = $LogKey
                }
                #If the RemovePolicy flag is no longer set, remove the blob from the registry to prevent issues later on.
                Invoke-ProcessRegistryItem @RegKeySplat | Out-Null
            }
        }
        #Clean up old log entries (any items that are no longer used).
        Remove-LogEntry -GUIDList $GUIDList -LogKey $LogKey
    }
}

#endregion

#region Variables
[hashtable] $RegistryValueKind = @{
    "REG_MULTI_SZ" = [Microsoft.Win32.RegistryValueKind]::MultiString
    "REG_DWORD" = [Microsoft.Win32.RegistryValueKind]::DWord
    "REG_SZ" = [Microsoft.Win32.RegistryValueKind]::String
    "REG_QWORD" = [Microsoft.Win32.RegistryValueKind]::QWord
    "REG_BINARY" = [Microsoft.Win32.RegistryValueKind]::Binary
    "REG_EXPAND_SZ" = [Microsoft.Win32.RegistryValueKind]::ExpandString
}
#endregion

#region Functions
function Get-RegistryRuleList {
    [cmdletbinding()]
    param (
        [OutputType([hashtable])]
        [CmdletBinding()]

        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelinebyPropertyName = $true)]
        [System.IO.DirectoryInfo]$Path,

        [Parameter(Mandatory = $false, ValueFromPipeline = $true, ValueFromPipelinebyPropertyName = $true)]
        [string[]]$RulesFileName

    )
    try {

        #Get JSON files if none were passed as params
        if(-not $RulesFileName) {
            $RulesFileName = Get-ChildItem -Path $Path -File '*.JSON'
        }

        #Get JSON Content
        $Rules = Foreach($file in $RulesFileName) {
            Get-Content -Path (Join-Path -Path $Path.FullName -ChildPath $File) -Raw | ConvertFrom-Json -ErrorAction SilentlyContinue
        }

        #Create rule groups based on parent structure
        if($Rules) {
            $properties = $Rules | Get-Member -MemberType NoteProperty -ErrorAction SilentlyContinue | select-Object -ExpandProperty Name
            [hashtable]$RuleObj = @{}
            foreach($Property in $properties) {
                $RuleObj[$Property] = @()
                $RuleObj[$Property] += $Rules.$Property | ForEach-Object {$_}
            }
        }
        else {
            $RuleObj = $null
        }

        return $RuleObj
    }
    catch {
        throw $_
    }
}

function Invoke-ProcessRegistryItem {
    [cmdletbinding()]
    param(

        [OutputType([Microsoft.Win32.RegistryKey])]
        [CmdletBinding()]

        [Parameter(Mandatory = $False, ValueFromPipeline = $true, ValueFromPipelinebyPropertyName = $true)]
        [string]$Group,

        [Parameter(Mandatory = $False, ValueFromPipeline = $true, ValueFromPipelinebyPropertyName = $true)]
        [string]$GUID,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelinebyPropertyName = $true)]
        [string]$Path,

        [Parameter(Mandatory = $false, ValueFromPipeline = $true, ValueFromPipelinebyPropertyName = $true)]
        [string]$EntryName,

        [Parameter(Mandatory = $false, ValueFromPipeline = $true, ValueFromPipelinebyPropertyName = $true)]
        [string]$PropertyName,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false, ValueFromPipelinebyPropertyName = $true)]
        [object]$PropertyValue,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false, ValueFromPipelinebyPropertyName = $true)]
        [Microsoft.Win32.RegistryValueKind]$PropertyType,
        
        [Parameter(Mandatory = $false, ValueFromPipeline = $false, ValueFromPipelinebyPropertyName = $true)]
        [ValidateSet('C','R','U','D')]
        [string]$Action,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false, ValueFromPipelinebyPropertyName = $true)]
        [switch]$RemovePolicy,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false, ValueFromPipelinebyPropertyName = $true)]
        [int]$GPOSettingOrder,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false, ValueFromPipelinebyPropertyName = $true)]
        [string]$Desc,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false, ValueFromPipelinebyPropertyName = $true)]
        [switch]$BypassErrors,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false, ValueFromPipelinebyPropertyName = $true)]
        [switch]$DisplayDecimal,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false, ValueFromPipelinebyPropertyName = $true)]
        [string]$Name,
        
        [Parameter(Mandatory = $false, ValueFromPipeline = $false, ValueFromPipelinebyPropertyName = $true)]
        [switch]$RunOnce,
        
        [Parameter(Mandatory = $false, ValueFromPipeline = $false, ValueFromPipelinebyPropertyName = $true)]
        $Filters,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false, ValueFromPipelinebyPropertyName = $true)]
        $LogKey

    )

    try {

        $Updated = $False

        #If a GUID is present, a log entry will be created to track the entry.
        if($GUID) {
            $CurrentPropertyLog = $LogKey | Get-ItemProperty -Name $GUID -ErrorAction SilentlyContinue
            $LastUpdated = $CurrentPropertyLog | Get-ItemPropertyValue -Name $GUID -ErrorAction SilentlyContinue

            #If the rule is set to remove the policy when no longer applied, export the settings to a registry entry for later use
            if($RemovePolicy.IsPresent) {
                $EntryBlob = [PSCustomObject]@{
                    GUID = $GUID
                    Path = $Path
                    PropertyName = $PropertyName
                    PropertyValue = $PropertyValue
                } | ConvertTo-Json -ErrorAction SilentlyContinue
                $Type = [Microsoft.Win32.RegistryValueKind]::String
                $LogKey | New-ItemProperty -Name "$($GUID)_Blob" -Value "$($EntryBlob)" -PropertyType $Type -Force
            }
            else {
                #delete any rule blog registry entries when RemovePolicy is not set on the rule
                $LogKey | Remove-ItemProperty -Name "$($GUID)_Blob" -Force -ErrorAction SilentlyContinue
            }

        }

        #If the rule has already been run, don't run again if RunOnce is set.
        if($RunOnce.IsPresent -and $null -ne $LastUpdated) {
            $CurrentProperty = Get-ItemProperty -Path $Path -Name $PropertyName -ErrorAction SilentlyContinue
        }
        else {
            $CurrentItemPath = Get-Item -Path $Path -ErrorAction SilentlyContinue

            #Delete item and children if Action is Replace or Delete and PropertyName/PropertyValue are not set
            if($CurrentItemPath -and ($action -in ('R','D')) -and (-not $PropertyName) -and (-not $PropertyValue)) {
                $CurrentItemPath | Remove-Item -Force
                $CurrentItemPath = $Null
            }

            #Create new key if Action is Create or Replace
            if(-not $CurrentItemPath -and ($action -in ('C','R','U'))) {
                $CurrentItemPath = New-Item -Path $Path -Force
                $Updated = $True
            }

            if($PropertyName) {
                $CurrentProperty = $CurrentItemPath | Get-ItemProperty -Name $PropertyName -ErrorAction SilentlyContinue

                #Delete item and children if Action is Replace and PropertyValue is not set or if Action is Delete (no PropertyValue is passed for a Delete Action)
                if($CurrentProperty -and (($action -in ('R')) -and (-not $PropertyValue)) -or ($action -in ('D'))) {
                    $CurrentProperty | Remove-ItemProperty -Name $PropertyName -Force -ErrorAction SilentlyContinue
                    $CurrentProperty = $Null
                }
                
                #Create new property if action is Create or Replace
                if(-not $CurrentProperty -and ($action -in ('C','R','U'))) {
                    $CurrentProperty = $CurrentItemPath | New-ItemProperty -Name $PropertyName -PropertyType $PropertyType -Force
                    $Updated = $True
                }
            }

            if($PropertyValue -and $PropertyName -and ($action -in ('C','R','U'))) {
                $CurrentPropertyValue = $CurrentProperty | Get-ItemPropertyValue -Name $PropertyName -ErrorAction SilentlyContinue

                #Delete item and children if Action is Replace
                if($CurrentPropertyValue -and ($action -in ('R'))) {
                    $CurrentProperty | Set-ItemProperty -Name $PropertyName -Value $Null
                    $CurrentPropertyValue = $Null
                }
                
                #Set value if action is Create Update or Replace 
                if($action -in ('C','R','U')) {
                    $CurrentPropertyValue = $CurrentProperty | Set-ItemProperty -Name $PropertyName -Value $PropertyValue
                    $Updated = $True
                }
            }

            #Update LogEntry with GUID and TimeStamp when update has occurred
            if($GUID -and $Updated -and (-not $LastUpdated)) {
                $timestamp = Get-Date -Format o | ForEach-Object { $_ -replace ":", "." }
                $LogKey | New-ItemProperty -Name $GUID -Value $timestamp -Force
            }
            
        }
        return $CurrentProperty
    }
    catch {
        throw $_
   }
}

function Remove-LogEntry {
    [CmdletBinding()]
    param(

        [OutputType([Microsoft.Win32.RegistryKey])]
        [CmdletBinding()]

        [Parameter(Mandatory = $False, ValueFromPipeline = $true, ValueFromPipelinebyPropertyName = $true)]
        [string[]]$GUIDList,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false, ValueFromPipelinebyPropertyName = $true)]
        $LogKey
    )


    #Get all rule blob log entries so they can be removed when not applied any longer
    $BlobEntries = ((Get-ItemProperty -Path $LogKey.PSPath -ErrorAction SilentlyContinue).psobject.Properties | Where-Object {$_.Name -like '*_blob'}).value | ConvertFrom-Json -ErrorAction SilentlyContinue
    ForEach($Entry in $BlobEntries) {
        $RegKeySplat = @{
            GUID = $Entry.GUID
            Path = $Entry.Path
            PropertyName = $Entry.PropertyName
            PropertyValue = $Entry.PropertyValue
            Action = 'D'
            LogKey = $LogKey
        }

        #Process the Delete action for all blobs that aren't in the current GUID rules list
        if($GUIDList -notcontains $Entry.GUID) {
            Invoke-ProcessRegistryItem @RegKeySplat | Out-Null
            Invoke-ProcessRegistryItem -Path $LogKey.PSPath -Action 'D' -PropertyName "$($Entry.GUID)_blob" -BypassErrors | Out-Null
        }
    }
    
    #Delete all log entries for GUIDs that are no longer applied
    $LogEntries = Get-Item -Path $LogKey.PSPath -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Property | Where-Object {$_ -notlike '*_Blob' -and $_ -ne '(Default)'}
    foreach($LogEntry in $LogEntries) {
        if($GUIDList -notcontains $LogEntry) {
            Invoke-ProcessRegistryItem -Path $LogKey.PSPath -Action 'D' -PropertyName $LogEntry -BypassErrors | Out-Null
        }
    }
}
#endregion

#region Launch Main
& $Main
#endregion