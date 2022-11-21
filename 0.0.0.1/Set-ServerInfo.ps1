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
# SIG # Begin signature block
# MIIFYQYJKoZIhvcNAQcCoIIFUjCCBU4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUatYJJxehZz7cKMNkNvjGd3/v
# 9emgggMAMIIC/DCCAeSgAwIBAgIQFnL4oVNG56NIRjNfzwNXejANBgkqhkiG9w0B
# AQUFADAWMRQwEgYDVQQDDAtDZXJ0LTAzNDU2MDAeFw0yMTEyMDIwNDU5MTJaFw0y
# MjEyMDIwNTE5MTJaMBYxFDASBgNVBAMMC0NlcnQtMDM0NTYwMIIBIjANBgkqhkiG
# 9w0BAQEFAAOCAQ8AMIIBCgKCAQEA8daSAcUBI0Xx8sMMlSpsCV+24lY46RsxX8iC
# bB7ZM19b/GBjwMo0TCb28ssbZ/P8liNJICrSbyIkQDrIrjqtAdyAPdPAYHONTHad
# 0fuOQQT5MkO5HAxUYLz/6H/xq92lKQFxz5Wgzw+3KOyignY8V8ZZ379z/WqQbNCV
# +29zb9YWOK7eXQ9x8s4+SOizqUE3zkOuijf86I9vZmzMYhsxE7if0R0UlQsLlvTA
# kH/m4IjHem8rl/kC+O71lU7l9475XrUUR3Fxebqh9YoCEZh2eE81TLQcnvK8zgqP
# F+X4INdNPD6zO4T1Nbz0Ccev7mj37+pk/eL5R5aV+NJgqAzhvQIDAQABo0YwRDAO
# BgNVHQ8BAf8EBAMCBaAwEwYDVR0lBAwwCgYIKwYBBQUHAwMwHQYDVR0OBBYEFFNN
# e4x6JSqbcnTR354fVSEgQ0VYMA0GCSqGSIb3DQEBBQUAA4IBAQBXfA8VgaMD2c/v
# Sv8gnS/LWri51BBqcUFE9JYMxEIzlEt2ZfJsG+INaQqzBoyCDx/oMQH7wdFRvDjQ
# QsXpNTo7wH7WytFe9KJrOz2uGG0EnIYHK0dTFIMVOcM9VsWWPG40EAzD//55xX/d
# pBL+L4SSTujbR3ptni8Agu5GiRhTpxwl1L/HLC2QYYMoUKiAxL1p61+cHRj6wMzl
# jxnrMIcBhKioaXnwWdKPCN66Jk8IYdqr8afcRYiwtDi+8Hk2/9nB9HwPox3Dtf8H
# jH0O2/8NiJTeOBFSfrWPM9r4j4NWR8IuLwsqHUfXJEQa9SOxhHvxaNMR/Fhq1GVj
# qUClZiXiMYIByzCCAccCAQEwKjAWMRQwEgYDVQQDDAtDZXJ0LTAzNDU2MAIQFnL4
# oVNG56NIRjNfzwNXejAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAA
# oQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4w
# DAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQURJzPuvMcM6hyUZAXUmmYAY0r
# +bkwDQYJKoZIhvcNAQEBBQAEggEAefzOzOPYOniOzh7o3Xp4uoxVBcwBDCmlttvv
# Y/wQrjQrxvdfQeRTij2i+GPgyz4d1ShWP9Ad3W0IwgNuUc6Ir/XsHoj7crdLQRs3
# 0nCXBlbv4LhaAZIrFlGFvJ4QcPi6NDYMQnIcLkyhi8LBMtUkdTxCAxAykajgrbOS
# lzw8gU419FgzRtlEFXHDxpwzmlke4vN33cNk0TgtxnMOF+SgfWjqLNeesowEYTUI
# IY8+gU/mX6FlOUthIzOObkI8Y2UyOAoVIujREunsjZ4rzvI0GpV8ncWloaLIwch5
# F2aRqMYWNBwpcNxQVRle8SuN1L2Ew/o7USxFHDOJhYCGrqFUAw==
# SIG # End signature block
