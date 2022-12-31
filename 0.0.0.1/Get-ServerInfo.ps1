# .ExternalHelp $PSScriptRoot\Get-ServerInfo-help.xml
function Get-ServerInfo {

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
        [ValidateScript({ $_ -in $script:Config.InfoType })]
        [System.String]
        $InfoType = 'DriveSpace',

        [Parameter(Position = 3, ParameterSetName = 'Collection')]
        [Parameter(Position = 3, ParameterSetName = 'Device')]
        [ValidateScript({ $_ -in $script:Config.Win32Class })]
        [System.String]
        $Win32Class = 'BaseBoard',

        [Parameter(Position = 4, ParameterSetName = 'Collection')]
        [Parameter(Position = 4, ParameterSetName = 'Device')]
        [ValidateScript({ $_ -in $script:Config.OutputType })]
        [System.String]
        $OutputType = 'Json',

        [Parameter(Position = 5, ParameterSetName = 'Collection')]
        [Parameter(Position = 5, ParameterSetName = 'Device')]
        [ValidateScript({ $_ -in $script:Config.ResultFormat })]
        [System.String[]]
        $ResultFormat = 'Teams',

        [Parameter(Position = 6, ParameterSetName = 'Collection')]
        [Parameter(Position = 6, ParameterSetName = 'Device')]
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

        [Parameter(ParameterSetName = 'Device')]
        [Switch]
        $Detailed,

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

    $scriptParameters = @{
        InfoType = $InfoType
    }
    if ($InfoType -eq 'Win32Class') {
        $scriptParameters.Win32Class = $Win32Class
    }
    if ($OutputType) { $scriptParameters.OutputType = $OutputType }

    $invokeParams = @{}
    if ($PSCmdlet.ParameterSetName -eq 'Collection') { $invokeParams.CollectionName = $CollectionName }
    else { $invokeParams.ComputerName = $ComputerName }
    $invokeParams.ScriptName = 'Patching.GetServerInfo'
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

    Switch ($InfoType) {
        'Cluster' {
            ForEach ($m in $Members) {
                if ($m.ScriptResults -notin @('2008', '2008_R2', 'UNKNOWN', 'OFFLINE', 'EXCEEDS')) {
                    if ('Cluster' -notin $($m.ScriptResults | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name -Unique)) {
                        $m.ScriptResults = [PSCustomObject]@{
                            Cluster       = $m.ScriptResults.C
                            ClusterNodes  = $m.ScriptResults.CN | ForEach-Object {
                                $_ -Replace '[@{}=INS]' | ConvertFrom-String -Delimiter '; ' -PropertyNames Id, Name, State
                            } | Sort-Object -Property Id
                            ClusterGroups = $m.ScriptResults.CG | ForEach-Object {
                                $_ -Replace '[@{}=NSO]' | ConvertFrom-String -Delimiter '; ' -PropertyNames Name, State, OwnerNode
                            } | Sort-Object -Property Name
                        }
                    }
                    else {
                        $m.ScriptResults = [PSCustomObject]@{
                            Cluster       = $m.ScriptResults.Cluster
                            ClusterNodes  = $m.ScriptResults.ClusterNodes | ForEach-Object {
                                $_ -Replace '[@{}=]|(Id)|(Name)|(State)' | ConvertFrom-String -Delimiter '; ' -PropertyNames Id, Name, State
                            } | Sort-Object -Property Id
                            ClusterGroups = $m.ScriptResults.ClusterGroups | ForEach-Object {
                                $_ -Replace '[@{}=]|(Name)|(State)|(OwnerNode)' | ConvertFrom-String -Delimiter '; ' -PropertyNames Name, State, OwnerNode
                            } | Sort-Object -Property Name
                        }
                    }
                }
            }
        }
        'DefaultService' {
            ForEach ($m in $Members) {
                if ($m.ScriptResults -notin @('2008', '2008_R2', 'UNKNOWN', 'OFFLINE', 'EXCEEDS')) {
                    $m.ScriptResults = $m.ScriptResults | ForEach-Object {
                        if ('Name' -notin $($_ | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name -Unique)) {
                            [PSCustomObject]@{
                                Name   = $script:Config.DefaultService.$($_.N)
                                Status = $script:Config.ServiceStatus.$($_.S)
                            }
                        }
                        else {
                            $_.Name = $script:Config.DefaultService.$($_.Name)
                            $_.Status = $script:Config.ServiceStatus.$($_.Status)
                            $_
                        }   
                    } | Sort-Object -Property Name
                }
            }
        }
        'InstalledPatches' {
            ForEach ($m in $Members) {
                if ($m.ScriptResults -notin @('2008', '2008_R2', 'UNKNOWN', 'OFFLINE', 'EXCEEDS')) {
                    $m.ScriptResults = $m.ScriptResults | ForEach-Object {
                        if ('HotFixID' -notin $($_ | Get-Member -MemberType NoteProperty -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name -Unique)) {
                            if ('HF' -notin $($_ | Get-Member -MemberType NoteProperty -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name -Unique)) {
                                [PSCustomObject]@{
                                    HotFixID = "KB$($_)"
                                    Type     = 'Unknown'
                                }
                            }
                            else {
                                [PSCustomObject]@{
                                    HotFixID = "KB$($_.HF)"
                                    Type     = $script:Config.PatchDescription.$($_.T)
                                }
                            }
                        }
                        else {
                            $_.Type = $script:Config.PatchDescription.$($_.Type)
                            $_
                        }
                    }
                }
            }
        }
        'PatchingStatus' {
            ForEach ($m in $Members) {
                if ($m.ScriptResults -notin @('2008', '2008_R2', 'UNKNOWN', 'OFFLINE', 'EXCEEDS')) {
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
        }
        'Service' {
            ForEach ($m in $Members) {
                if ($m.ScriptResults -notin @('2008', '2008_R2', 'UNKNOWN', 'OFFLINE', 'EXCEEDS')) {
                    $m.ScriptResults = $m.ScriptResults | ForEach-Object {
                        if ('Name' -notin $($_ | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name -Unique)) {
                            [PSCustomObject]@{
                                Name   = $_.N
                                Status = $script:Config.ServiceStatus.$($_.S)
                            }
                        }
                        else {
                            $_.Status = $script:Config.ServiceStatus.$($_.Status)
                            $_
                        }
                    } | Sort-Object -Property Name
                }
            }
        }
        'Task' {
            ForEach ($m in $Members) {
                if ($m.ScriptResults -notin @('2008', '2008_R2', 'UNKNOWN', 'OFFLINE', 'EXCEEDS')) {
                    $m.ScriptResults = $m.ScriptResults | ForEach-Object {
                        if ('TaskName' -notin $($_ | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name -Unique)) {
                            [PSCustomObject]@{
                                TaskName = $_.TN
                                State    = $script:Config.TaskState.$($_.S)
                            }
                        }
                        else {
                            $_.State = $script:Config.TaskState.$($_.State)
                            $_
                        }
                    } | Sort-Object -Property TaskName
                }
            }
        }
        'Win32Class' {
            ForEach ($m in $Members) {
                if ($m.ScriptResults -notin @('2008', '2008_R2', 'UNKNOWN', 'OFFLINE', 'EXCEEDS')) {
                    Switch ($Win32Class) {
                        'DiskDrive' {
                            $m.ScriptResults = $m.ScriptResults | ForEach-Object {
                                [PSCustomObject]@{
                                    Index            = $_.I
                                    Caption          = $_.C
                                    Model            = $_.M
                                    SerialNumber     = $_.SN
                                    FirmwareRevision = $_.FWR
                                    Status           = $_.S
                                }
                            } | Sort-Object -Property Index
                        }
                        'DiskPartition' {
                            $m.ScriptResults = $m.ScriptResults | ForEach-Object {
                                [PSCustomObject]@{
                                    DiskIndex   = $_.DI
                                    Index       = $_.I
                                    Description = $_.D
                                }
                            } | Sort-Object -Property DiskIndex, Index
                        }
                        'LogicalDisk' {
                            $m.ScriptResults = $m.ScriptResults | ForEach-Object {
                                [PSCustomObject]@{
                                    Name               = $_.N
                                    Size               = $_.S
                                    FreeSpace          = $_.F
                                    DriveType          = $_.DT
                                    FileSystem         = $_.FS
                                    Description        = $_.D
                                    VolumeSerialNumber = $_.VSN
                                    Compressed         = $_.C
                                    VolumeDirty        = $_.VD
                                }
                            } | Sort-Object -Property Name
                        }
                        'NetworkAdapter' {
                            $m.ScriptResults = $m.ScriptResults | ForEach-Object {
                                [PSCustomObject]@{
                                    Name            = $_.N
                                    MACAddress      = $_.MAC
                                    Manufacturer    = $_.M
                                    Speed           = $_.SPD
                                    NetConnectionID = $_.NCID
                                    NetEnabled      = $_.NE
                                    PhysicalAdapter = $_.PA
                                }
                            } | Sort-Object -Property Name
                        }
                        'OptionalFeature' {
                            $m.ScriptResults = $m.ScriptResults | ForEach-Object {
                                [PSCustomObject]@{
                                    Name         = $script:Config.OptionalFeature.$($_.N)
                                    InstallState = $_.IS
                                }
                            } | Sort-Object -Property Name
                        }
                        'PhysicalMemory' {
                            $m.ScriptResults = $m.ScriptResults | ForEach-Object {
                                [PSCustomObject]@{
                                    Name          = $_.N
                                    DeviceLocator = $_.DL
                                    FormFactor    = $_.FF
                                    MemoryType    = $_.MT
                                    Speed         = $_.SPD
                                    Capacity      = $_.C
                                    DataWidth     = $_.DW
                                    TotalWidth    = $_.TW
                                    TypeDetail    = $_.TD
                                    SerialNumber  = $_.SN
                                    Manufacturer  = $_.MF
                                    PartNumber    = $_.PN
                                }
                            } | Sort-Object -Property Name
                        }
                        'Process' {
                            $m.ScriptResults = $m.ScriptResults | ForEach-Object {
                                [PSCustomObject]@{
                                    ProcessName = $_.PN
                                    Handles     = $_.H
                                    VM          = $_.VM
                                    WS          = $_.WS
                                }
                            } | Sort-Object -Property ProcessName
                        }
                        Default { }
                    }
                }
            }
        }
        Default { }
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
                #if ($PSCmdlet.ParameterSetName -eq 'Collection') {
                    $params = New-Object -TypeName System.Collections.Hashtable
                    $params.Title = "$($MyInvocation.MyCommand.ModuleName): $($MyInvocation.MyCommand.Name) - InfoType: ${InfoType}"
                    ForEach ($collection in $($Members.CollectionName | Sort-Object -Unique)) {
                        $collectionMembers = $Members | Where-Object { $_.CollectionName -eq $collection }
                        $params.ActivityTitle = "Members: $(@($collectionMembers | Where-Object { $_.ScriptResults -notin @('2008', '2008_R2', 'OFFLINE', 'UNKNOWN', 'EXCEEDS') }).Count) out of $(@($collectionMembers).Count)"
                        $params.ActivitySubtitle = "Start: $(($collectionMembers.ScriptStartTime | Select-Object -First 1).ToString('ddd MMM %d HH:mm:ss yyyy'))"
                        $params.ActivityText = "Complete: $(Get-Date -UFormat '%c')"
                        $params.FactSectionList = [System.Collections.Generic.List[System.Collections.Hashtable]]::New()
                        Switch ($InfoType) {
                            'Cluster' {
                                $collectionMembers.ScriptResults.Cluster | Where-Object { $_ -ne 'None' } | Sort-Object -Unique | ForEach-Object {
                                    $clusterName = $_
                                    $clusterMembers = $collectionMembers | Where-Object { $_.ScriptResults.Cluster -eq $clusterName }
                                    $section = @{ Title = $clusterName; Facts = [System.Collections.Generic.List[System.Collections.Hashtable]]::New() }
                                    $clusterMembers.ScriptResults.ClusterNodes.Id | Sort-Object -Unique | ForEach-Object {
                                        $nodeID = $_
                                        $node = $clusterMembers.ScriptResults.ClusterNodes | Where-Object { $_.Id -eq $nodeID } | Select-Object -First 1
                                        $section.Facts.Add(@{ name = "ClusterNode: $($node.Id)"; value = "$($node.State): $($node.Name)" })
                                    }
                                    $clusterMembers.ScriptResults.ClusterGroups.Name | Sort-Object -Unique | ForEach-Object {
                                        $groupName = $_
                                        $group = $clusterMembers.ScriptResults.ClusterGroups | Where-Object { $_.Name -eq $groupName } | Select-Object -First 1
                                        $section.Facts.Add(@{ name = "ClusterGroup: $($group.Name)"; value = "$($group.State): $(($clusterMembers.ScriptResults.ClusterNodes | Where-Object { $_.Id -eq $group.OwnerNode }).Name | Select-Object -First 1)" })
                                    }
                                    $section.FactFormatHash = [System.Collections.Generic.List[System.Collections.Hashtable]]::New()
                                    ForEach ($status in $($section.Facts.value | ForEach-Object { $_.Split(':') | Select-Object -First 1 } | Sort-Object -Unique)) {
                                        $section.FactFormatHash.Add(@{ Status = $status; Color = $script:Config.Color.$status })
                                    }
                                    $params.FactSectionList.Add($section)
                                }
                            }
                            'DefaultService' {
                                $section = @{ Title = $collection; Facts = [System.Collections.Generic.List[System.Collections.Hashtable]]::New() }
                                $collectionMembers | ForEach-Object {
                                    $factString = New-Object -TypeName System.Text.StringBuilder
                                    $_.ScriptResults | Group-Object -Property Status -NoElement | Sort-Object -Property Name | ForEach-Object {
                                        $colorKey = Get-Random -InputObject $script:Config.HTMLColors -Count 1
                                        [void]$factString.Append("<font style=`"color:${colorKey}`"><b>$($_.Name)</b></font>: $($_.Count), ")
                                    }
                                    $section.Facts.Add(@{ name = $_.ComputerName; value = $factString.ToString().TrimEnd(', ') })
                                }
                                $params.FactSectionList.Add($section)
                            }
                            'DesktopExperience' {
                                $section = @{ Title = $collection; Facts = [System.Collections.Generic.List[System.Collections.Hashtable]]::New() }
                                $collectionMembers | ForEach-Object {
                                    $section.Facts.Add(@{ name = $_.ComputerName; value = "<font style=`"color:$(Get-Random -InputObject $script:Config.HTMLColors -Count 1)`"><b>$($_.ScriptResults.ProductName)</b></font> \ <font style=`"color:$(Get-Random -InputObject $script:Config.HTMLColors -Count 1)`"><b>$($_.ScriptResults.InstallationType)</b></font>" })
                                }
                                $params.FactSectionList.Add($section)
                            }
                            'DriveSpace' {
                                $section = @{ Title = $collection; Facts = [System.Collections.Generic.List[System.Collections.Hashtable]]::New() }
                                $collectionMembers | ForEach-Object {
                                    $factString = New-Object -TypeName System.Text.StringBuilder
                                    $_.ScriptResults | Where-Object { $_.Name -in @('C', 'E') } | ForEach-Object {
                                        $colorKey = Get-Random -InputObject $script:Config.HTMLColors -Count 1
                                        [void]$factString.Append("<font style=`"color:${colorKey}`"><b>$($_.Name)</b></font> $($_.Free)/$($_.Total) - ")
                                    }
                                    $section.Facts.Add(@{ name = $_.ComputerName; value = $factString.ToString().TrimEnd(' - ') })
                                }
                                $params.FactSectionList.Add($section)
                            }
                            'IIS' {
                                $section = @{ Title = $collection; Facts = [System.Collections.Generic.List[System.Collections.Hashtable]]::New() }
                                $collectionMembers | ForEach-Object {
                                    $factString = New-Object -TypeName System.Text.StringBuilder
                                    $_.ScriptResults | Group-Object -Property State -NoElement | Sort-Object -Property Name | ForEach-Object {
                                        $colorKey = Get-Random -InputObject $script:Config.HTMLColors -Count 1
                                        [void]$factString.Append("<font style='color:${colorKey}'><b>$($_.Name)</b></font>: $($_.Count), ")
                                    }
                                    $section.Facts.Add(@{ name = $_.ComputerName; value = $factString.ToString().TrimEnd(', ') })
                                }
                            }
                            'InstalledPatches' {
                                $section = @{ Title = $collection; Facts = [System.Collections.Generic.List[System.Collections.Hashtable]]::New() }
                                $collectionMembers | ForEach-Object {
                                    $factString = New-Object -TypeName System.Text.StringBuilder
                                    $_.ScriptResults | Group-Object -Property Type -NoElement | Sort-Object -Property Name | ForEach-Object {
                                        $colorKey = Get-Random -InputObject $script:Config.HTMLColors -Count 1
                                        [void]$factString.Append("<font style=`"color:${colorKey}`"><b>$($_.Name)</b></font>: $($_.Count), ")
                                    }
                                    $section.Facts.Add(@{ name = $_.ComputerName; value = $factString.ToString().TrimEnd(', ') })
                                }
                                $params.FactSectionList.Add($section)
                            }
                            'PatchingStatus' {
                                $section = @{ Title = $collection; Facts = [System.Collections.Generic.List[System.Collections.Hashtable]]::New(); FactFormatHash = [System.Collections.Generic.List[System.Collections.Hashtable]]::New() }
                                $collectionMembers | ForEach-Object {
                                    $factString = New-Object -TypeName System.Text.StringBuilder
                                    if ($_.ScriptResults -notin @('2008', '2008_R2', 'OFFLINE', 'UNKNOWN', 'EXCEEDS')) {
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
                                        elseif ('None' -eq $($_.ScriptResults.ArticleID | Sort-Object -Unique)) {
                                            [void]$factString.Append("COMPLETE: No Patches Present")
                                            $section.Facts.Add(@{ name = $_.ComputerName; value = $factString.ToString() })
                                            if ('COMPLETE' -notin $section.FactFormatHash.Values) {
                                                $section.FactFormatHash.Add(@{ Status = 'COMPLETE'; Color = $script:Config.Color.COMPLETE })
                                            }
                                        }
                                        else {
                                            $notReady = [System.Collections.Generic.List[System.String]]::New()
                                            1, 2, 3, 4, 5, 6, 7, 11, 21, 22 | ForEach-Object { $notReady.Add($script:Config.EvaluationState.$($_)) }
                                            if (@($_.ScriptResults | Where-Object { $_.EvaluationState -in $notReady }).Count -ge 1) {
                                                [void]$factString.Append("BUSY: ")
                                                ForEach ($state in $($_.ScriptResults | Where-Object { $_.EvaluationState -in $notReady } | Select-Object -ExpandProperty EvaluationState -Unique)) {
                                                    [void]$factString.Append("${state} - $(@($_.ScriptResults | Where-Object { $_.EvaluationState -eq $state }).Count), ")
                                                }
                                                $section.Facts.Add(@{ name = $_.ComputerName; value = $factString.ToString().TrimEnd(', ') })
                                                if ('BUSY' -notin $section.FactFormatHash.Values) {
                                                    $section.FactFormatHash.Add(@{ Status = 'BUSY'; Color = $script:Config.Color.BUSY })
                                                }
                                            }
                                            else {
                                                [void]$factString.Append("READY: ")
                                                ForEach ($state in $($_.ScriptResults | Where-Object { $_.EvaluationState -notin $notReady } | Select-Object -ExpandProperty EvaluationState -Unique)) {
                                                    [void]$factString.Append("${state} - $(@($_.ScriptResults | Where-Object { $_.EvaluationState -eq $state }).Count), ")
                                                }
                                                $section.Facts.Add(@{ name = $_.ComputerName; value = $factString.ToString().TrimEnd(', ') })
                                                if ('READY' -notin $section.FactFormatHash.Values) {
                                                    $section.FactFormatHash.Add(@{ Status = 'READY'; Color = $script:Config.Color.READY })
                                                }
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
                            'PatchWindow' {
                                $section = @{ Title = $collection; Facts = [System.Collections.Generic.List[System.Collections.Hashtable]]::New() }
                                $collectionMembers | ForEach-Object {
                                    $section.Facts.Add(@{ name = $_.ComputerName; value = "<font style=`"color:$(Get-Random -InputObject $script:Config.HTMLColors -Count 1)`"><b>$($_.ScriptResults)</b></font>" })
                                }
                                $params.FactSectionList.Add($section)
                            }
                            'Service' {
                                $section = @{ Title = $collection; Facts = [System.Collections.Generic.List[System.Collections.Hashtable]]::New() }
                                $collectionMembers | ForEach-Object {
                                    $factString = New-Object -TypeName System.Text.StringBuilder
                                    if ($_.ScriptResults -notin @('2008', '2008_R2', 'OFFLINE', 'UNKNOWN', 'EXCEEDS')) {
                                        $_.ScriptResults | Group-Object -Property Status -NoElement | Sort-Object -Property Name | ForEach-Object {
                                            $colorKey = Get-Random -InputObject $script:Config.HTMLColors -Count 1
                                            [void]$factString.Append("<font style=`"color:${colorKey}`"><b>$($_.Name)</b></font>: $($_.Count), ")
                                        }
                                        $section.Facts.Add(@{ name = $_.ComputerName; value = $factString.ToString().TrimEnd(', ') })
                                    }
                                    else {
                                        [void]$factString.Append("<font style=`"color:$($script:Config.Color.$($_.ScriptResults))`"><b>$($_.ScriptResults)</b></font>")
                                        $section.Facts.Add(@{ name = $_.ComputerName; value = $factString.ToString() })
                                    }
                                }
                                $params.FactSectionList.Add($section)
                            }
                            'Task' {
                                $section = @{ Title = $collection; Facts = [System.Collections.Generic.List[System.Collections.Hashtable]]::New() }
                                $collectionMembers | ForEach-Object {
                                    $factString = New-Object -TypeName System.Text.StringBuilder
                                    if ($_.ScriptResults -notin @('2008', '2008_R2', 'OFFLINE', 'UNKNOWN', 'EXCEEDS')) {
                                        $_.ScriptResults | Group-Object -Property State -NoElement | Sort-Object -Property TaskName | ForEach-Object {
                                            $colorKey = Get-Random -InputObject $script:Config.HTMLColors -Count 1
                                            [void]$factString.Append("<font style=`"color:${colorKey}`"><b>$($_.Name)</b></font>: $($_.Count), ")
                                        }
                                        $section.Facts.Add(@{ name = $_.ComputerName; value = $factString.ToString().TrimEnd(', ') })
                                    }
                                    else {
                                        [void]$factString.Append("<font style=`"color:$($script:Config.Color.$($_.ScriptResults))`"><b>$($_.ScriptResults)</b></font>")
                                        $section.Facts.Add(@{ name = $_.ComputerName; value = $factString.ToString() })
                                    }
                                }
                                $params.FactSectionList.Add($section)
                            }
                            #'Win32Class' {
                            #    $section = @{ Title = $collection; Facts = [System.Collections.Generic.List[System.Collections.Hashtable]]::New() }
                            #    $collectionMembers | ForEach-Object {
                            #        $factString = New-Object -TypeName System.Text.StringBuilder
                                    #if ($_.ScriptResults -notin @('2008', '2008_R2', 'OFFLINE', 'UNKNOWN', 'EXCEEDS')) {
                                        
                                    #}
                                }
                            #}
                        }
                        New-TeamsNotification @params
                    }
                #}
                #else {
                #    $params = New-Object -TypeName System.Collections.Hashtable
                #    $params.Title = "$($MyInvocation.MyCommand.ModuleName): $($MyInvocation.MyCommand.Name) - InfoType: ${InfoType}"
                    # ForEach ($)
                #}
            }
        }
    }
#}
