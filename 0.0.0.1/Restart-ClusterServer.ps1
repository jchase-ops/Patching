# .ExternalHelp $PSScriptRoot\Restart-ClusterServer-help.xml
function Restart-ClusterServer {

    [CmdletBinding(DefaultParameterSetName = 'Collection')]

    Param (

        [Parameter(Position = 0, ParameterSetName = 'Collection')]
        [ValidateNotNullOrEmpty()]
        [System.String[]]
        $CollectionName,

        [Parameter(Mandatory, Position = 0, ParameterSetName = 'Device')]
        [ValidateNotNullOrEmpty()]
        [System.String[]]
        $ComputerName,

        [Parameter(Position = 1, ParameterSetName = 'Collection')]
        [Parameter(Position = 1, ParameterSetName = 'Device')]
        [ValidateSet('Prepare', 'Restart', 'Rebalance')]
        [System.String]
        $PatchingStep,

        [Parameter(Position = 2, ParameterSetName = 'Collection')]
        [Parameter(Position = 2, ParameterSetName = 'Device')]
        [ValidateScript({ $_ -in $script:Config.OutputType })]
        [System.String]
        $OutputType = 'Json',

        [Parameter(Position = 3, ParameterSetName = 'Collection')]
        [Parameter(Position = 3, ParameterSetName = 'Device')]
        [ValidateScript({ $_ -in $script:Config.ResultFormat })]
        [System.String[]]
        $ResultFormat = 'Teams',

        [Parameter(Position = 4, ParameterSetName = 'Collection')]
        [Parameter(Position = 4, ParameterSetName = 'Device')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $DrivePath = $env:TEMP,

        [Parameter(Position = 5, ParameterSetName = 'Collection')]
        [Parameter(Position = 5, ParameterSetName = 'Device')]
        [ValidateNotNullOrEmpty()]
        [PSCredential]
        $Credential,

        [Parameter(ParameterSetName = 'Collection')]
        [Parameter(ParameterSetName = 'Device')]
        [Switch]
        $SaveCredential,

        [Parameter(ParameterSetName = 'Collection')]
        [Parameter(ParameterSetName = 'Device')]
        [Switch]
        $Quiet
    )

    $suppress = if (!(Test-WindowVisible) -or $Quiet) { $true } else { $false }

    if ((Get-PSDrive -PSProvider CMSite -ErrorAction SilentlyContinue).Count -eq 0) {
        if ($suppress) { Connect-SCCM -Quiet }
        else { Connect-SCCM }
    }
    else {
        $siteDrive = Get-PSDrive -PSProvider CMSite
        if ($PWD.Path -ne "$($siteDrive.Name)") {
            Push-Location -Path "$($siteDrive.Name):\" -StackName $MyInvocation.MyCommand.ModuleName
        }
    }

    if (!($DrivePath)) {
        if ($null -eq $script:Config.PatchingDrive.Root) {
            $DrivePath = $env:TEMP
        }
    }
    else {
        $script:Config.PatchingDrive.Root = $DrivePath
    }
    $script:Config | Export-Clixml -Path "$PSScriptRoot\config.xml" -Depth 100

    if (!($Credential)) {
        if ($DrivePath -ne $env:TEMP) {
            if ($null -eq $script:Config.PatchingDrive.Credential) {
                if (!($suppress)) {
                    $script:Config.PatchingDrive.Credential = $Host.UI.PromptForCredential('Patching PS Drive Credentials', 'Enter credentials for Patching PS Drive', "$($env:USERDOMAIN)\$($env:USERNAME)", '')
                }
                else {
                    return 1
                }
            }
        }
        else {
            $script:Config.PatchingDrive.Credential = [PSCredential]::New("${env:USERDOMAIN}\${env:USERNAME}", $(ConvertTo-SecureString -String 'pw' -AsPlainText -Force))
        }
    }
    else {
        $script:Config.PatchingDrive.Credential = $Credential
    }
    if ($SaveCredential) {
        $script:Config | Export-Clixml -Path "$PSScriptRoot\config.xml" -Depth 100
    }

    if ($null -eq $(Get-PSDrive -Name $script:Config.PatchingDrive.Name -ErrorAction SilentlyContinue)) {
        try {
            if ($DrivePath -ne $env:TEMP) {
                $null = New-PSDrive -Name $script:Config.PatchingDrive.Name -PSProvider $script:Config.PatchingDrive.PSProvider -Root $script:Config.PatchingDrive.Root -Scope $script:Config.PatchingDrive.Scope -Credential $script:Config.PatchingDrive.Credential
            }
            else {
                $null = New-PSDrive -Name $script:Config.PatchingDrive.Name -PSProvider $script:Config.PatchingDrive.PSProvider -Root $env:TEMP -Scope $script:Config.PatchingDrive.Scope
            }
        }
        catch {
            if (!($suppress)) {
                Write-Host 'Failed to mount Patching PS Drive'
                Write-Host 'Mounting Temp PS Drive'
            }
            $null = New-PSDrive -Name $script:Config.PatchingDrive.Name -PSProvider $script:Config.PatchingDrive.PSProvider -Root "$PSScriptRoot\Temp" -Scope $script:Config.PatchingDrive.Scope
        }
    }

    if (!($ResultFormat)) {
        if (!($suppress)) {
            $ResultFormat = $script:Config.ResultFormat | Out-GridView -Title 'ResultFormat' -OutputMode Multiple
        }
        else {
            return 1
        }
    }

    if (!($PatchingStep)) {
        if (!($suppress)) {
            $PatchingStep = 'Prepare', 'Restart', 'Rebalance' | Out-GridView -Title 'PatchingStep' -OutputMode Single
        }
        else {
            return 1
        }
    }

    $scriptParameters = [System.Collections.Hashtable]::New()
    Switch ($PatchingStep) {
        'Prepare' { $scriptParameters.CommandType = 'PrepareCluster' }
        'Restart' { $scriptParameters.CommandType = 'RestartClusterServer' }
        'Rebalance' { $scriptParameters.CommandType = 'RebalanceCluster' }
    }
    if ($OutputType) { $scriptParameters.OutputType = $OutputType }

    $invokeParams = @{}
    if ($PSCmdlet.ParameterSetName -eq 'Collection') { $invokeParams.CollectionName = $CollectionName }
    else { $invokeParams.ComputerName = $ComputerName }
    $invokeParams.ScriptName = 'Patching.RestartServer'
    $invokeParams.ScriptParameters = $scriptParameters
    if ($suppress) { $invokeParams.Quiet = $true }
    Start-CMScriptInvocation @invokeParams
    if (!($suppress)) {
        Write-Host 'Waiting 180 seconds for CM Scripts to complete...'
    }
    Start-Sleep -Seconds 180
    if ($suppress) {
        $invokeResults = Receive-CMScriptInvocation -Quiet
    }
    else {
        $invokeResults = Receive-CMScriptInvocation
    }

    $invokeResults | Group-Object -Property Name | ForEach-Object {
        $group = $_
        $group.Group.Members | ForEach-Object {
            if ($PSCmdlet.ParameterSetName -eq 'Collection') {
                Add-Member -InputObject $_ -NotePropertyName CollectionName -NotePropertyValue $group.Name
            }
            else {
                Add-Member -InputObject $_ -NotePropertyName CollectionName -NotePropertyValue 'Assorted'
            }
            Add-Member -InputObject $_ -NotePropertyName Fact -NotePropertyValue $([ordered]@{ name = $_.ComputerName; value = $null })
            Add-Member -InputObject $_ -NotePropertyName ClientOperationID -NotePropertyValue $group.Group.ClientOperationID
            Add-Member -InputObject $_ -NotePropertyName CompletedClients -NotePropertyValue $group.Group.CompletedClients
            Add-Member -InputObject $_ -NotePropertyName FailedClients -NotePropertyValue $group.Group.FailedClients
            Add-Member -InputObject $_ -NotePropertyName NotApplicableClients -NotePropertyValue $group.Group.NotApplicableClients
            Add-Member -InputObject $_ -NotePropertyName OfflineClients -NotePropertyValue $group.Group.OfflineClients
            Add-Member -InputObject $_ -NotePropertyName TotalClients -NotePropertyValue $group.Group.TotalClients
            Add-Member -InputObject $_ -NotePropertyName ScriptStartTime -NotePropertyValue $group.Group.ScriptStartTime
            Add-Member -InputObject $_ -NotePropertyName ScriptName -NotePropertyValue $group.Group.ScriptName
        }
    }

    $Members = $invokeResults.Members | Sort-Object -Property ComputerName
    ForEach ($m in $Members) {
        if ([version]$m.DeviceOSBuild -lt 6.2) {
            if ([version]$m.DeviceOSBuild -lt 6.1) {
                $m.ScriptResults = '2008'
            }
            else {
                $m.ScriptResults = '2008_R2'
            }
        }
        elseif ($invokeResults.Details.ComputerName.IndexOf($m.ComputerName) -lt 0) {
            if (Test-Connection -ComputerName $m.ComputerName -Count 3 -Quiet) {
                $m.ScriptResults = 'UNKNOWN'
            }
            else {
                $m.ScriptResults = 'OFFLINE'
            }
        }
        else {
            $m.ScriptResults = ($invokeResults.Details[$($invokeResults.Details.ComputerName.IndexOf($m.ComputerName))]).Details
        }
    }

    ForEach ($m in $Members) {
        if ($m.ScriptResults -notin @('2008', '2008_R2', 'UNKNOWN', 'OFFLINE', 'EXCEEDS', 'EXCLUDED', 'REBOOTING', 'PREPARED', 'CLUSTER_GROUP_OFFLINE', 'CLUSTER_GROUP_MOVE_FAILED', 'REBALANCED', 'SKIPPED')) {
            $m.ScriptResults = $m.ScriptResults | ForEach-Object {
                if ('ArticleID' -notin $($_ | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name -Unique)) {
                    if ($_.KB -ne 'None') {
                        [PSCustomObject]@{
                            ArticleID       = $_.KB
                            EvaluationState = $script:Config.EvaluationState.$($_.ES)
                            ErrorCode       = $_.EC
                            PercentComplete = $_.PC
                        }
                    }
                    else {
                        [PSCustomObject]@{
                            ArticleID       = $_.KB
                            EvaluationState = $_.ES
                            ErrorCode       = $_.EC
                            PercentComplete = $_.PC
                        }
                    }
                }
                else {
                    if ($_.ArticleID -ne 'None') {
                        $_.EvaluationState = $script:Config.EvaluationState.$($_.EvaluationState)
                    }
                    $_
                }
            } | Sort-Object -Property EvaluationState, PercentComplete
        }
    }

    if (Get-Location -StackName $MyInvocation.MyCommand.ModuleName -ErrorAction SilentlyContinue) {
        Pop-Location -StackName $MyInvocation.MyCommand.ModuleName
    }

    ForEach ($rf in $ResultFormat) {
        Switch ($rf) {
            'Console' {
                $Members | Sort-Object -Property ComputerName
            }
            'Email' {
                Write-Host 'Under Development' -ForegroundColor Yellow
            }
            'File' {
                Write-Host 'Under Development' -ForegroundColor Yellow
            }
            'ServiceDesk' {
                Write-Host 'Under Development' -ForegroundColor Yellow
            }
            'SharePoint' {
                Write-Host 'Under Development' -ForegroundColor Yellow
            }
            'Teams' {
                $params = New-Object -TypeName System.Collections.Hashtable
                $params.Title = "$($MyInvocation.MyCommand.ModuleName): $($MyInvocation.MyCommand.Name) - Step: ${PatchingStep}"
                ForEach ($collection in $($Members.CollectionName | Sort-Object -Unique)) {
                    $collectionMembers = $Members | Where-Object { $_.CollectionName -eq $collection }
                    $params.ActivityTitle = "Members: $(@($collectionMembers | Where-Object { $_.ScriptResults -notin @('2008', '2008_R2', 'OFFLINE', 'UNKNOWN', 'EXCEEDS') }).Count) out of $(@($collectionMembers).Count)"
                    $params.ActivitySubtitle = "Start: $(($collectionMembers.ScriptStartTime | Select-Object -First 1).ToString('ddd MMM %d HH:mm:ss yyyy'))"
                    $params.ActivityText = "Complete: $(Get-Date -UFormat '%c')"
                    $params.FactSectionList = [System.Collections.Generic.List[System.Collections.Hashtable]]::New()
                    $section = @{ Title = $collection; Facts = [System.Collections.Generic.List[System.Collections.Hashtable]]::New(); FactFormatHash = [System.Collections.Generic.List[System.Collections.Hashtable]]::New() }
                    $collectionMembers | ForEach-Object {
                        $factString = New-Object -TypeName System.Text.StringBuilder
                        if ($_.ScriptResults -notin @('2008', '2008_R2', 'OFFLINE', 'UNKNOWN', 'EXCEEDS', 'EXCLUDED', 'REBOOTING', 'PREPARED', 'CLUSTER_GROUP_OFFLINE', 'CLUSTER_GROUP_MOVE_FAILED', 'REBALANCED', 'SKIPPED')) {
                            if ('Error' -in $($_.ScriptResults.EvaluationState | Sort-Object -Unique)) {
                                [void]$factString.Append("ERROR: ")
                                ForEach ($ec in $(($_.ScriptResults | Where-Object { $_.EvaluationState -eq 'Error' }).ErrorCode | Sort-Object -Unique)) {
                                    [void]$factString.Append("${ec}, ")
                                }
                                $section.Facts.Add(@{ name = $_.ComputerName; value = $factString.ToString().TrimEnd(', ') })
                                if ('ERROR' -notin $section.FactFormatHash.Values) {
                                    $section.FactFormatHash.Add(@{ Status = 'ERROR'; Color = $script:Config.Color.ERROR })
                                }
                            }
                            else {
                                $notReady = [System.Collections.Generic.List[System.String]]::New()
                                1, 2, 3, 4, 5, 6, 7, 11, 21, 22 | ForEach-Object { $notReady.Add($script:Config.EvaluationState.$($_)) }
                                [void]$factString.Append("BUSY: ")
                                ForEach ($state in $($_.ScriptResults | Where-Object { $_.EvaluationState -in $notReady } | Select-Object -ExpandProperty EvaluationState -Unique)) {
                                    [void]$factString.Append("${state} - $(@($_.ScriptResults | Where-Object { $_.EvaluationState -eq $state }).Count), ")
                                }
                                $section.Facts.Add(@{ name = $_.ComputerName; value = $factString.ToString().TrimEnd(', ') })
                                if ('BUSY' -notin $section.FactFormatHash.Values) {
                                    $section.FactFormatHash.Add(@{ Status = 'BUSY'; Color = $script:Config.Color.BUSY })
                                }
                            }
                        }
                        else {
                            [void]$factString.Append($_.ScriptResults)
                            $section.Facts.Add(@{ name = $_.ComputerName; value = $factString.ToString() })
                            if ($_.ScriptResults -notin $section.FactFormatHash.Values) {
                                $section.FactFormatHash.Add(@{ Status = $_.ScriptResults; Color = $script:Config.Color.$($_.ScriptResults) })
                            }
                        }
                    }
                    $params.FactSectionList.Add($section)
                }
                New-TeamsNotification @params
            }
        }
    }
}
