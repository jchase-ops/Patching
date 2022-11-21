# .ExternalHelp $PSScriptRoot\New-MonthlyPatchDeployment-help.xml
function New-MonthlyPatchDeployment {

    [CmdletBinding(DefaultParameterSetName = 'Current')]

    Param (

        [Parameter(Mandatory, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $CollectionName,

        [Parameter(Mandatory, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [System.DateTime]
        $AvailableDateTime,

        [Parameter(Mandatory, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [System.DateTime]
        $DeadlineDateTime,

        [Parameter(Position = 3, ParameterSetName = 'Current')]
        [Parameter(Position = 3, ParameterSetName = 'Previous')]
        [ValidateScript({ $_ -in $script:Config.ResultFormat })]
        [System.String[]]
        $ResultFormat = 'Teams',

        [Parameter(Position = 4, ParameterSetName = 'Current')]
        [Parameter(Position = 4, ParameterSetName = 'Previous')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $DrivePath = $env:TEMP,

        [Parameter(Position = 5, ParameterSetName = 'Current')]
        [Parameter(Position = 5, ParameterSetName = 'Previous')]
        [ValidateNotNullOrEmpty()]
        [PSCredential]
        $Credential,

        [Parameter(Position = 6, ParameterSetName = 'Current')]
        [Parameter(Position = 6, ParameterSetName = 'Previous')]
        [ValidateScript({ $_ -in $script:DeploymentConfiguration.Keys })]
        [System.String]
        $Configuration = 'Test',

        [Parameter(ParameterSetName = 'Current')]
        [Parameter(ParameterSetName = 'Previous')]
        [Switch]
        $SaveCredential,

        [Parameter(ParameterSetName = 'Current')]
        [Switch]
        $RenameSoftwareUpdateGroup,

        [Parameter(Mandatory, ParameterSetName = 'Previous')]
        [Switch]
        $UseLastMonthlyPackages,

        [Parameter(ParameterSetName = 'Current')]
        [Parameter(ParameterSetName = 'Previous')]
        [Switch]
        $GetUpdateInformation,

        [Parameter(ParameterSetName = 'Current')]
        [Parameter(ParameterSetName = 'Previous')]
        [Switch]
        $CreateDeployment,

        [Parameter(ParameterSetName = 'Current')]
        [Parameter(ParameterSetName = 'Previous')]
        [Switch]
        $Quiet
    )

    $Date = Get-Date

    $suppress = if (!(Test-WindowVisible) -or $Quiet) { $true } else { $false }

    if ((Get-PSDrive -PSProvider CMSite -ErrorAction SilentlyContinue).Count -eq 0) {
        if ($suppress) { Connect-SCCM -Quiet }
        else { Connect-SCCM }
        $siteDrive = Get-PSDrive -PSProvider CMSite
        if ($PWD.Path -ne "$($siteDrive.Name)") {
            Push-Location -Path "$($siteDrive.Name):\" -StackName $MyInvocation.MyCommand.ModuleName
        }
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

    if (!($Configuration)) {
        if (!($suppress)) {
            $Configuration = $script:DeploymentConfiguration.Keys | Out-GridView -Title 'DeploymentConfiguration' -OutputMode Single
        }
        else {
            return 1
        }
    }

    Switch ($PSCmdlet.ParameterSetName) {
        'Current' {
            $SoftwareUpdateGroup = Get-CMSoftwareUpdateGroup -Name "Monthly_Server_ADR $($Date.ToString('yyyy-MM*'))"
            if ($SoftwareUpdateGroup.Count -eq 0) {
                if (!($suppress)) {
                    Write-Host "'Monthly_Server_ADR $($Date.ToString('yyyy-MM*'))' not found" -ForegroundColor Yellow
                    Write-Host "Checking if '$($Date.ToString('MMMM_yyyy'))_Server_Patches' exists..."
                }
                $SoftwareUpdateGroup = Get-CMSoftwareUpdateGroup -Name "$($Date.ToString('MMMM_yyyy'))_Server_Patches"
                if ($SoftwareUpdateGroup.Count -eq 0) {
                    if (!($suppress)) {
                        Write-Host "'$($Date.ToString('MMMM_yyyy'))_Server_Patches not found" -ForegroundColor Red
                        pause
                        exit
                    }
                    else {
                        return 3
                    }
                }
            }
            if ($RenameSoftwareUpdateGroup) {
                if ($SoftwareUpdateGroup.LocalizedDisplayName -ne "$($Date.ToString('MMMM_yyyy'))_Server_Patches") {
                    $SoftwareUpdateGroup = Set-CMSoftwareUpdateGroup -Name $SoftwareUpdateGroup.LocalizedDisplayName -NewName "$($Date.ToString('MMMM_yyyy'))_Server_Patches" -PassThru
                }
            }
        }
        'Previous' {
            $SoftwareUpdateGroup = Get-CMSoftwareUpdateGroup -Name "$($Date.AddMonths(-1).ToString('MMMM_yyyy'))_Server_Patches"
            if ($SoftwareUpdateGroup.Count -eq 0) {
                if (!($suppress)) {
                    Write-Host "'$($Date.AddMonths(-1).ToString('MMMM_yyyy'))_Server_Patches not found" -ForegroundColor Yellow
                    Write-Host "Checking if 'Monthly_Server_ADR $($Date.AddMonths(-1).ToString('yyyy-MM*')) exists..."
                }
                $SoftwareUpdateGroup = Get-CMSoftwareUpdateGroup -Name "Monthly_Server_ADR $($Date.AddMonths(-1).ToString('yyyy-MM*'))"
                if ($SoftwareUpdateGroup.Count -ne 0) {
                    $SoftwareUpdateGroup = Set-CMSoftwareUpdateGroup -Name $SoftwareUpdateGroup.LocalizedDisplayName -NewName "$($Date.AddMonths(-1).ToString('MMMM_yyyy'))_Server_Patches" -PassThru
                }
                else {
                    if (!($suppress)) {
                        Write-Host "'Monthly_Server_ADR $($Date.AddMonths(-1).ToString('yyyy-MM*')) not found" -ForegroundColor Red
                        pause
                        exit
                    }
                    else {
                        return 3
                    }
                }
            }
        }
    }

    $params = @{
        AvailableDateTime = $AvailableDateTime
        CollectionName = $CollectionName
        DeadlineDateTime = $DeadlineDateTime
        SoftwareUpdateGroupName = $SoftwareUpdateGroup.LocalizedDisplayName
    }

    if ($PSCmdlet.ParameterSetName -eq 'Current') { $params.DeploymentName = "${CollectionName}_$($Date.ToString('MMMM_yyyy'))" }
    else { $params.DeploymentName = "${CollectionName}_$($Date.AddMonths(-1).ToString('MMMM_yyyy'))" }

    ForEach ($key in $script:DeploymentConfiguration.$Configuration.Keys) {
        $params.Add($key, $script:DeploymentConfiguration.$Configuration.$key)
    }

    if (!($CreateDeployment)) {
        New-Object -TypeName PSObject -Property $params
    }
    else {
        if (!($suppress)) {
            Write-Host "Checking if $($params.DeploymentName) exists..."
        }
        $ExistingDeployment = Get-CMSoftwareUpdateDeployment -CollectionName $params.CollectionName | Where-Object { $_.AssignmentName -eq $params.DeploymentName }
        if ($ExistingDeployment.Count -gt 0) {
            if (!($suppress)) {
                Write-Host $params.DeploymentName -NoNewline
                Write-Host " found" -ForegroundColor Red
                pause
                exit
            }
            else {
                return 4
            }
        }
        else {
            $Deployment = New-CMSoftwareUpdateDeployment @params
            if ($?) {
                if (!($suppress)) {
                    Write-Host $params.DeploymentName -NoNewline
                    Write-Host ' created' -ForegroundColor Green
                }
            }
            else {
                if (!($suppress)) {
                    Write-Host $params.DeploymentName -NoNewline
                    Write-Host ' failed' -ForegroundColor Red
                    pause
                }
                else {
                    return 2
                }
            }
        }
    }

    if (Get-Location -StackName $MyInvocation.MyCommand.ModuleName -ErrorAction SilentlyContinue) {
        Pop-Location -StackName $MyInvocation.MyCommand.ModuleName
    }
}
# SIG # Begin signature block
# MIIFYQYJKoZIhvcNAQcCoIIFUjCCBU4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUQOzBAhw1au1ajLTf5nupd5A9
# PCWgggMAMIIC/DCCAeSgAwIBAgIQFnL4oVNG56NIRjNfzwNXejANBgkqhkiG9w0B
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
# DAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUlHSWxnljrow4Jbm4LmnL2gpL
# ZsYwDQYJKoZIhvcNAQEBBQAEggEAreipAGJEf1LNWCRUgHVSmekA9MRTiVYJ1+rp
# sncKejmLtbB/Iu52TRCIwsPxx+qwIuCUUG/OUvXM06ff1dUYrJxv9OO04pQjXAIh
# MZpVAtwH0wi9h5mKN6eptSjPRWoFs0G77yOVSoSrRcmP+JJnyMfNdUrdSW0f7xCH
# fF4hEYg5xm5FjYJl2ZrwEB8W5kyrHM5x5YitKNDuGA4h5YdHFT3uuxSx+aOm52Bk
# pA5LafJJBtxj94JwJ4npeKb5BvZfu6NLT8y6XrriO9T+g8y6lGNT8d3ZjK9kiviw
# s3GQsqIjNdSBASTDQE4sSmYGYfLA/sIYG+SO0pdmBtutMamh9Q==
# SIG # End signature block
