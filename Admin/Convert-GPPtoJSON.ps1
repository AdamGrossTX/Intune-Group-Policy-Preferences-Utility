[cmdletbinding()]
param(
    [Parameter(Mandatory = $false, ValueFromPipeline = $true, ValueFromPipelinebyPropertyName = $true)]
    [string[]]$GPOName = "Test GPP",

    [Parameter(Mandatory = $false, ValueFromPipeline = $true, ValueFromPipelinebyPropertyName = $true)]
    [System.IO.DirectoryInfo]$ExportPath = ".\Rules"
)

#region Main
$Main = {
    Get-GPOSettings -GPOName $GPOName -ExportPath $ExportPath
}
#endregion

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
        foreach($Name in $GPOName) {
            $GPO = Get-GPO -Name $Name -ErrorAction SilentlyContinue
            if($GPO) {
                [xml]$GPOReport = Get-GPOReport -Guid $GPO.Id -Domain $env:USERDNSDOMAIN -ReportType xml
                $RegistrySettings = $GPOReport.GPO.Computer.ExtensionData.Extension.RegistrySettings
                $ValueList = @{}
                foreach($Setting in $RegistrySettings) {
                    if($Setting.Collection){
                        $ValueList[$Setting.Collection.Name] = New-RuleObj -RegistryObject $Setting.Collection.Registry
                    }
                    if($Setting.Registry){
                        $ValueList["Root"] = New-RuleObj -RegistryObject $Setting.Registry
                    }
                }
                if($ValueList) {
                    if(-Not (Test-Path -Path $ExportPath)) {
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
        [PSCustomObject]$RegistryObject
    )
    try {
        $obj = foreach($Child in $RegistryObject) {
            $Filters = @{}
            [int]$runOnce = 0
            foreach($Node in $Child.Filters.ChildNodes) {
                if($node.LocalName -contains 'FilterRunOnce') {
                    $runOnce = 1
                } 
                else {
                    $Filters[$Node.LocalName] = 
                    $Item = @{}
                    foreach($Attribute in $Node.Attributes) {
                        $Item[$Attribute.LocalName] = $Attribute.Value
                    }
                }
            }

            [PSCustomObject]@{
                GUID = (New-GUID)
                GPOSettingOrder = [int]$child.GPOSettingOrder
                desc = [string]$child.desc
                bypassErrors = [int]$child.bypassErrors
                action = [string]$child.Properties.action
                displayDecimal = [int]$child.Properties.displayDecimal
                default = [int]$child.Properties.default
                hive = [string]$child.Properties.hive
                key = [string]$child.Properties.key
                entryName = [string]$child.name
                propertyName = [string]$child.Properties.name
                propertyvalue = [string]$child.Properties.value
                propertytype = [string]$child.Properties.type
                removePolicy = [int]$child.removePolicy
                runOnce = [int]$runOnce
                filters = if($Filters.keys) {$filters} else {$null}
            }
        }

        return $obj
    }
    catch {
        throw $_
    }
}
#endregion

#Region Launch Main
& $Main
#endregion