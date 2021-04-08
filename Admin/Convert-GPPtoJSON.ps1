[cmdletbinding()]
param(
    [Parameter(Mandatory = $false, ValueFromPipeline = $true, ValueFromPipelinebyPropertyName = $true)]
    [string[]]$GPOName = "Test GPP",

    [Parameter(Mandatory = $false, ValueFromPipeline = $true, ValueFromPipelinebyPropertyName = $true)]
    [System.IO.DirectoryInfo]$ExportPath = ".\Rules"
)

#region Functions
function Get-GPOSettings {
    param(
        [cmdletbinding()]

        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelinebyPropertyName = $true)]
        [string[]]$GPOName,

        [Parameter(Mandatory = $false, ValueFromPipeline = $true, ValueFromPipelinebyPropertyName = $true)]
        [System.IO.DirectoryInfo]$ExportPath
    )

    try {
        foreach ($Name in $GPOName) {
            $GPO = Get-GPO -Name $Name -ErrorAction SilentlyContinue
            if ($GPO) {
                [xml]$GPOReport = Get-GPOReport -Guid $GPO.Id -Domain $env:USERDNSDOMAIN -ReportType xml
                [string[]]$TypeList = $GPOReport.GPO.Computer.ExtensionData.Extension.FirstChild | Select-Object -ExpandProperty LocalName
                $ValueList = @{}
                foreach ($Type in $TypeList) {
                    $Settings = $GPOReport.GPO.Computer.ExtensionData.Extension.$Type
                    $AllSettingNames = $Settings | Get-Member -MemberType Property | Where-Object { $_.Name -ne 'clsid' } | Select-Object -ExpandProperty Name
                    $SettingNames = $AllSettingNames | Where-Object { $_ -ne 'Collection' }
                    if ($AllSettingNames -contains 'Collection') {
                        foreach ($SettingName in $SettingNames) {
                            $ValueList["$($SettingName)_$($Setting.ParentNode.Name)"] = 
                            foreach ($Setting in $Settings.Collection.$SettingName) {
                                New-RuleObj -Object $Setting -PreferenceType $SettingName
                            }
                        }
                    }
                    foreach ($SettingName in $SettingNames) {
                        $ValueList["$($SettingName)_Root"] = 
                        foreach ($Setting in $Settings.$SettingName) {
                            if ($Setting) {
                                New-RuleObj -Object $Setting -PreferenceType $SettingName
                            }
                        }
                    }
                }
                if ($ValueList) {
                    if (-Not (Test-Path -Path $ExportPath)) {
                        $ExportPath = $ExportPath | New-Item -Path $ExportPath.FullName -ItemType Directory -ErrorAction SilentlyContinue -Force
                    }
                    $ValueList | ConvertTo-Json -Depth 10 | Out-File -FilePath "$(Join-Path -Path $ExportPath.FullName -ChildPath $Name).JSON" -Force
                }
            }
        }
    }
    catch {
        throw $_
    }
}

function New-RuleObj {
    [CmdletBinding()]
    param (
        [OutputType([PSCustomObject])]

        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelinebyPropertyName = $true)]
        [PSCustomObject]$Object,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelinebyPropertyName = $true)]
        [PSCustomObject]$PreferenceType
    )
    try {
        foreach ($Child in $Object) {
            $Filters = @{}
            [int]$runOnce = 0
            foreach ($Node in $Child.Filters.ChildNodes) {
                if ($node.LocalName -contains 'FilterRunOnce') {
                    $runOnce = 1
                } 
                else {
                    $Filters[$Node.LocalName] = 
                    $Item = @{}
                    foreach ($Attribute in $Node.Attributes) {
                        $Item[$Attribute.LocalName] = $Attribute.Value
                    }
                }
            }

            #ItemProperties
            $PreferenceType = $PreferenceType
            $GUID = (New-GUID)
            $GPOSettingOrder = [int]$child.GPOSettingOrder
            $entryName = [string]$child.name
            $desc = [string]$child.desc
            $bypassErrors = [int]$child.bypassErrors
            $disabled = [string]$child.disabled
            $removePolicy = [int]$child.removePolicy
            $userContext = [int]$child.userContext
            $runOnce = [int]$runOnce
            $action = [string]$child.Properties.action

            $filters = if ($Filters.keys) { $filters } else { $null }
            
            $obj = switch ($preferenceType) {
                "Registry" {
                    [PSCustomObject]@{
                        preferenceType  = $preferenceType
                        GUID            = $GUID
                        GPOSettingOrder = $GPOSettingOrder
                        entryName       = $entryName
                        desc            = $desc
                        bypassErrors    = $bypassErrors
                        disabled        = $disabled
                        removePolicy    = $removePolicy
                        runOnce         = $runOnce
                        action          = $action
                        displayDecimal  = [int]$child.Properties.displayDecimal
                        default         = [int]$child.Properties.default
                        hive            = [string]$child.Properties.hive
                        key             = [string]$child.Properties.key
                        name            = [string]$child.Properties.name
                        value           = [string]$child.Properties.value
                        type            = [string]$child.Properties.type
                           
                    }
                }
                "Folder" {
                    [PSCustomObject]@{
                        preferenceType     = $preferenceType
                        GUID               = $GUID
                        GPOSettingOrder    = $GPOSettingOrder
                        entryName          = $entryName
                        desc               = $desc
                        bypassErrors       = $bypassErrors
                        disabled           = $disabled
                        removePolicy       = $removePolicy
                        runOnce            = $runOnce
                        action             = $action
                        hidden             = [int]$child.Properties.hidden
                        archive            = [int]$child.Properties.archive
                        readOnly           = [int]$child.Properties.readOnly
                        deleteIgnoreErrors = [int]$child.Properties.deleteIgnoreErrors
                        deleteReadOnly     = [int]$child.Properties.deleteReadOnly
                        deleteFiles        = [int]$child.Properties.deleteFiles
                        deleteSubFolders   = [int]$child.Properties.deleteSubFolders
                        deleteFolder       = [int]$child.Properties.deleteFolder
                        path               = [string]$child.Properties.path                        
                    }
                }
                "File" {
                    [PSCustomObject]@{
                        preferenceType  = $preferenceType
                        GUID            = $GUID
                        GPOSettingOrder = $GPOSettingOrder
                        entryName       = $entryName
                        desc            = $desc
                        bypassErrors    = $bypassErrors
                        disabled        = $disabled
                        removePolicy    = $removePolicy
                        runOnce         = $runOnce
                        action          = $action
                        fromPath        = [string]$child.Properties.fromPath
                        targetPath      = [string]$child.Properties.targetPath
                        readOnly        = [int]$child.Properties.readOnly
                        archive         = [int]$child.Properties.archive
                        hidden          = [int]$child.Properties.hidden
                        suppress        = [int]$child.Properties.suppress                        
                    }
                }
                "Group" {
                    [PSCustomObject]@{
                        preferenceType  = $preferenceType
                        GUID            = $GUID
                        GPOSettingOrder = $GPOSettingOrder
                        entryName       = $entryName
                        desc            = $desc
                        bypassErrors    = $bypassErrors
                        disabled        = $disabled
                        removePolicy    = $removePolicy
                        runOnce         = $runOnce
                        action          = $action
                        userContext     = $userContext

                        newName         = [string]$child.Properties.newName
                        description     = [string]$child.Properties.description
                        deleteAllUsers  = [int]$child.Properties.deleteAllUsers
                        deleteAllGroups = [int]$child.Properties.deleteAllGroups
                        groupSid        = [string]$child.Properties.groupSid
                        groupName       = [string]$child.Properties.groupName

                        name            = [string]$child.Properties.Members.Member.name
                        memberAction    = [string]$child.Properties.Members.Member.action
                        sid             = [string]$child.Properties.Members.Member.sid                        
                    }
                }
                "EnvironmentVariable" {
                    [PSCustomObject]@{
                        preferenceType  = $preferenceType
                        GUID            = $GUID
                        GPOSettingOrder = $GPOSettingOrder
                        entryName       = $entryName
                        desc            = $desc
                        bypassErrors    = $bypassErrors
                        disabled        = $disabled
                        removePolicy    = $removePolicy
                        runOnce         = $runOnce
                        action          = $action

                        name            = [string]$child.Properties.name
                        value           = [string]$child.Properties.value
                        user            = [int]$child.Properties.user
                        partial         = [int]$child.Properties.partial                        
                    }
                }
                "NTService" {
                    [PSCustomObject]@{
                        preferenceType  = $preferenceType
                        GUID            = $GUID
                        GPOSettingOrder = $GPOSettingOrder
                        entryName       = $entryName
                        desc            = $desc
                        bypassErrors    = $bypassErrors
                        disabled        = $disabled
                        removePolicy    = $removePolicy
                        userContext     = $userContext
                        runOnce         = $runOnce
                        action          = $action
                        startupType     = [string]$child.Properties.startupType
                        serviceName     = [string]$child.Properties.serviceName
                        serviceAction   = [string]$child.Properties.serviceAction
                        timeout         = [int]$child.Properties.timeout                        
                    }
                }
                "Shortcut" {
                    [PSCustomObject]@{
                        preferenceType  = $preferenceType
                        GUID            = $GUID
                        GPOSettingOrder = $GPOSettingOrder
                        entryName       = $entryName
                        desc            = $desc
                        bypassErrors    = $bypassErrors
                        disabled        = $disabled
                        removePolicy    = $removePolicy
                        userContext     = $userContext
                        runOnce         = $runOnce
                        action          = $action
                        pidl            = [string]$child.Properties.pidl
                        targetType      = [string]$child.Properties.targetType
                        comment         = [string]$child.Properties.comment
                        shortcutKey     = [string]$child.Properties.shortcutKey
                        startIn         = [string]$child.Properties.startIn
                        arguments       = [string]$child.Properties.arguments
                        iconIndex       = [int]$child.Properties.iconIndex
                        targetPath      = [string]$child.Properties.targetPath
                        iconPath        = [string]$child.Properties.iconPath
                        window          = [string]$child.Properties.window
                        shortcutPath    = [string]$child.Properties.shortcutPath                        
                    }
                }
                "Task" {
                    [PSCustomObject]@{
                        preferenceType          = $preferenceType
                        GUID                    = $GUID
                        GPOSettingOrder         = $GPOSettingOrder
                        entryName               = $entryName
                        desc                    = $desc
                        bypassErrors            = $bypassErrors
                        disabled                = $disabled
                        removePolicy            = $removePolicy
                        userContext             = $userContext
                        runOnce                 = $runOnce
                        action                  = $action
                        name                   = [string]$child.Properties.name
                        systemRequired         = [int]$child.Properties.systemRequired
                        stopIfGoingOnBatteries = [int]$child.Properties.stopIfGoingOnBatteries
                        noStartIfOnBatteries   = [int]$child.Properties.noStartIfOnBatteries
                        stopOnIdleEnd          = [int]$child.Properties.stopOnIdleEnd
                        deadlineMinutes        = [int]$child.Properties.deadlineMinutes
                        idleMinutes            = [int]$child.Properties.idleMinutes
                        startOnlyIfIdle        = [int]$child.Properties.startOnlyIfIdle
                        maxRunTime             = [int]$child.Properties.maxRunTime
                        deleteWhenDone         = [int]$child.Properties.deleteWhenDone
                        enabled                = [int]$child.Properties.enabled
                        comment                = [string]$child.Properties.comment
                        startIn                = [string]$child.Properties.startIn
                        args                   = [int]$child.Properties.args
                        appName                = [string]$child.Properties.appName
                    }
                }
                    <#
                    $Triggers = 
                                            #trigger
                                            $interval               = [int]$child.Properties.interval
                                            $repeatTask             = [int]$child.Properties.repeatTask
                                            $hasEndDate             = [int]$child.Properties.hasEndDate
                                            $beginDay               = [int]$child.Properties.beginDay
                                            $beginMonth             = [int]$child.Properties.beginMonth
                                            $beginYear              = [int]$child.Properties.beginYear
                                            $startMinutes           = [int]$child.Properties.startMinutes
                                            $startHour              = [int]$child.Properties.startHour
                                            $type                   = [string]$child.Properties.type
                                            $killAtDurationEnd      = [int]$child.Properties.killAtDurationEnd
                                            $minutesDuration        = [int]$child.Properties.minutesDuration
                                            $repeatUnit             = [string]$child.Properties.repeatUnit
                                            $minutesInterval        = [int]$child.Properties.minutesInterval
                                            $endDay                 = [int]$child.Properties.endDay
                                            $endMonth               = [int]$child.Properties.endMonth
                                            $endYear                = [int]$child.Properties.endYear
                    #>
            }
            return $obj
        }
    }
    catch {
        throw $_
    }
}
    #endregion

    #region Main

    Get-GPOSettings -GPOName $GPOName -ExportPath $ExportPath

    #endregion