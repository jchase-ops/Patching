# .ExternalHelp $PSScriptRoot\Set-ServerInfo-help.xml
function Set-ServerInfo {

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
        [ValidateSet('PatchWindow', 'Service', 'Task')]
        [System.String]
        $InfoType = 'PatchWindow',

        [Parameter(Position = 2, ParameterSetName = 'Collection')]
        [Parameter(Position = 2, ParameterSetName = 'Device')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $NewInfo,

        [Parameter(Position = 3, ParameterSetName = 'Collection')]
        [Parameter(Position = 3, ParameterSetName = 'Device')]
        [ValidateScript({ $_ -in $script:Default.OutputType })]
        [System.String]
        $OutputType = 'Json',

        [Parameter(Position = 4, ParameterSetName = 'Collection')]
        [Parameter(Position = 4, ParameterSetName = 'Device')]
        [ValidateScript({ $_ -in $script:Default.ResultFormat })]
        [System.String[]]
        $ResultFormat = 'Teams',

        [Parameter(Position = 5, ParameterSetName = 'Collection')]
        [Parameter(Position = 5, ParameterSetName = 'Device')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $DrivePath = $env:TEMP,

        [Parameter(Position = 6, ParameterSetName = 'Collection')]
        [Parameter(Position = 6, ParameterSetName = 'Device')]
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

    if (!($NewInfo)) {
        $NewInfo = Read-Host -Prompt "New ${InfoType}"
    }

    $scriptParameters = @{
        InfoType = $InfoType
        NewInfo  = $NewInfo
    }
    if ($OutputType) { $scriptParameters.OutputType = $OutputType }

    $invokeParams = @{}
    if ($PSCmdlet.ParameterSetName -eq 'Collection') { $invokeParams.CollectionName = $CollectionName }
    else { $invokeParams.ComputerName = $ComputerName }
    $invokeParams.ScriptName = 'Patching.SetServerInfo'
    $invokeParams.ScriptParameters = $scriptParameters
    if ($suppress) { $invokeParams.Quiet = $true }
    Start-CMScriptInvocation @invokeParams
    if (!($suppress)) {
        Write-Host 'Waiting 60 seconds for CM Scripts to complete...'
    }
    Start-Sleep -Seconds 60
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
                $params.Title = "$($MyInvocation.MyCommand.ModuleName): $($MyInvocation.MyCommand.Name) - InfoType: ${InfoType}"
                ForEach ($collection in $($Members.CollectionName | Sort-Object -Unique)) {
                    $collectionMembers = $Members | Where-Object { $_.CollectionName -eq $collection }
                    $params.ActivityTitle = "Members: $(@($collectionMembers | Where-Object { $_.ScriptResults -notin @('2008', '2008_R2', 'OFFLINE', 'UNKNOWN', 'EXCEEDS') }).Count) out of $(@($collectionMembers).Count)"
                    $params.ActivitySubtitle = "Start: $(($collectionMembers.ScriptStartTime | Select-Object -First 1).ToString('ddd MMM %d HH:mm:ss yyyy'))"
                    $params.ActivityText = "Complete: $(Get-Date -UFormat '%c')"
                    $params.FactSectionList = [System.Collections.Generic.List[System.Collections.Hashtable]]::New()
                    Switch ($InfoType) {
                        'PatchWindow' {
                            $section = @{ Title = $collection; Facts = [System.Collections.Generic.List[System.Collections.Hashtable]]::New() }
                            $collectionMembers | ForEach-Object {
                                $section.Facts.Add(@{ name = $_.ComputerName; value = "<font style=`"color:$(Get-Random -InputObject $script:Config.HTMLColors -Count 1)`"><b>$($_.ScriptResults)</b></font>" })
                            }
                            $params.FactSectionList.Add($section)
                        }
                    }
                    New-TeamsNotification @params
                }
            }
        }
    }
}
