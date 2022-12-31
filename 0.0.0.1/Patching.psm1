# Patching PS Module

#region Classes
################################################################################
#                                                                              #
#                                 CLASSES                                      #
#                                                                              #
################################################################################
# . "$PSScriptRoot\$(Split-Path -Path $(Split-Path -Path $PSScriptRoot -Parent) -Leaf).Classes.ps1"
#endregion

#region Variables
################################################################################
#                                                                              #
#                               VARIABLES                                      #
#                                                                              #
################################################################################
try {
    $script:Config = Import-Clixml -Path "$PSScriptRoot\config.xml" -ErrorAction Stop
}
catch {
    $script:Config = [ordered]@{
        CleanupType              = @(
            'RecycleBin'
            'CCMCache'
            'SoftwareDistribution'
            'TEMP'
            'All'
        )
        Color                    = [ordered]@{
            '2008'                      = 'SaddleBrown'
            '2008_R2'                   = 'SaddleBrown'
            'BUSY'                      = 'GreenYellow'
            'CLUSTER'                   = 'LightSalmon'
            'CLUSTER_GROUP_MOVE_FAILED' = 'MediumVioletRed'
            'CLUSTER_GROUP_OFFLINE'     = 'MediumVioletRed'
            'COMPLETE'                  = 'DarkCyan'
            'ERROR'                     = 'OrangeRed'
            'EXCEEDS'                   = 'DarkOrchid'
            'EXCLUDED'                  = 'DarkGray'
            'FAILED'                    = 'Red'
            'NO_KEY_EXISTS'             = 'OrangeRed'
            'NOT_SET'                   = 'MediumVioletRed'
            'OFFLINE'                   = 'DimGray'
            'ONLINE'                    = 'ForestGreen'
            'PARTIALONLINE'             = 'BurlyWood'
            'PREPARED'                  = 'MediumSpringGreen'
            'READY'                     = 'LimeGreen'
            'REBALANCED'                = 'MediumSpringGreen'
            'REBOOTING'                 = 'DeepSkyBlue'
            'SKIPPED'                   = 'DarkGray'
            'SUCCESS'                   = 'LawnGreen'
            'UNKNOWN'                   = 'DarkOrange'
            'UP'                        = 'Chartreuse'
        }
        DefaultService           = @{
            '0'   = 'AJRouter'
            '1'   = 'ALG'
            '2'   = 'AppIDSvc'
            '3'   = 'Appinfo'
            '4'   = 'AppMgmt'
            '5'   = 'AppReadiness'
            '6'   = 'AppVClient'
            '7'   = 'AppXSvc'
            '8'   = 'AudioEndpointBuilder'
            '9'   = 'Audiosrv'
            '10'  = 'AxInstSV'
            '11'  = 'BFE'
            '12'  = 'BITS'
            '13'  = 'BrokerInfrastructure'
            '14'  = 'Browser'
            '15'  = 'bthserv'
            '16'  = 'CDPSvc'
            '17'  = 'CDPUserSvc'
            '18'  = 'CertPropSvc'
            '19'  = 'ClipSVC'
            '20'  = 'COMSysApp'
            '21'  = 'CoreMessagingRegistrar'
            '22'  = 'CryptSvc'
            '23'  = 'CscService'
            '24'  = 'DcomLaunch'
            '25'  = 'DcpSvc'
            '26'  = 'defragsvc'
            '27'  = 'DeviceAssociationService'
            '28'  = 'DeviceInstall'
            '29'  = 'DevQueryBroker'
            '30'  = 'Dhcp'
            '31'  = 'diagnosticshub.standardcollector.service'
            '32'  = 'DiagTrack'
            '33'  = 'DmEnrollmentSvc'
            '34'  = 'dmwappushservice'
            '35'  = 'Dnscache'
            '36'  = 'dot3svc'
            '37'  = 'DPS'
            '38'  = 'DsmSvc'
            '39'  = 'DsSvc'
            '40'  = 'EapHost'
            '41'  = 'EFS'
            '42'  = 'embeddedmode'
            '43'  = 'EntAppSvc'
            '44'  = 'EventLog'
            '45'  = 'EventSystem'
            '46'  = 'fdPHost'
            '47'  = 'FDResPub'
            '48'  = 'FontCache'
            '49'  = 'FrameServer'
            '50'  = 'gpsvc'
            '51'  = 'hidserv'
            '52'  = 'HvHost'
            '53'  = 'icssvc'
            '54'  = 'IKEEXT'
            '55'  = 'Imhosts'
            '56'  = 'iphlpsvc'
            '57'  = 'KeyIso'
            '58'  = 'KPSSVC'
            '59'  = 'KtmRm'
            '60'  = 'LanmanServer'
            '61'  = 'LanmanWorkstation'
            '62'  = 'lfsvc'
            '63'  = 'LicenseManager'
            '64'  = 'lltdsvc'
            '65'  = 'LSM'
            '66'  = 'MapsBroker'
            '67'  = 'MpsSvc'
            '68'  = 'MSDTC'
            '69'  = 'MSiSCSI'
            '70'  = 'msiserver'
            '71'  = 'NcaSvc'
            '72'  = 'NcbService'
            '73'  = 'Netlogon'
            '74'  = 'Netman'
            '75'  = 'netprofm'
            '76'  = 'NetSetupSvc'
            '77'  = 'NetTcpPortSharing'
            '78'  = 'NgcCtnrSvc'
            '79'  = 'NgcSvc'
            '80'  = 'NlaSvc'
            '81'  = 'nsi'
            '82'  = 'OneSyncSvc'
            '83'  = 'PcaSvc'
            '84'  = 'PerfHost'
            '85'  = 'PhoneSvc'
            '86'  = 'PimIndexMaintenanceSvc'
            '87'  = 'pla'
            '88'  = 'PlugPlay'
            '89'  = 'PolicyAgent'
            '90'  = 'Power'
            '91'  = 'PrintNotify'
            '92'  = 'ProfSvc'
            '93'  = 'QWAVE'
            '94'  = 'RasAuto'
            '95'  = 'RasMan'
            '96'  = 'RemoteAccess'
            '97'  = 'RemoteRegistry'
            '98'  = 'RmSvc'
            '99'  = 'RpcEptMapper'
            '100' = 'RpcLocator'
            '101' = 'RpcSs'
            '102' = 'RSoPProv'
            '103' = 'sacsvr'
            '104' = 'SamSs'
            '105' = 'SCardSvr'
            '106' = 'ScDeviceEnum'
            '107' = 'Schedule'
            '108' = 'SCPolicySvc'
            '109' = 'seclogon'
            '110' = 'SENS'
            '111' = 'SensorDataService'
            '112' = 'SensorService'
            '113' = 'SensrSvc'
            '114' = 'SessionEnv'
            '115' = 'SharedAccess'
            '116' = 'ShellHWDetection'
            '117' = 'smphost'
            '118' = 'SNMPTRAP'
            '119' = 'Spooler'
            '120' = 'sppsvc'
            '121' = 'SSDPSRV'
            '122' = 'SstpSvc'
            '123' = 'StateRepository'
            '124' = 'stisvc'
            '125' = 'StorSvc'
            '126' = 'svsvc'
            '127' = 'swprv'
            '128' = 'SysMain'
            '129' = 'SystemEventsBroker'
            '130' = 'TabletInputService'
            '131' = 'TapiSrv'
            '132' = 'TermService'
            '133' = 'Themes'
            '134' = 'TieringEngineService'
            '135' = 'tiledatamodelsvc'
            '136' = 'TimeBrokerSvc'
            '137' = 'TrkWks'
            '138' = 'TrustedInstaller'
            '139' = 'tzautoupdate'
            '140' = 'UALSVC'
            '141' = 'UevAgentService'
            '142' = 'UIODetect'
            '143' = 'UmRdpService'
            '144' = 'UnistoreSvc'
            '145' = 'upnphost'
            '146' = 'UserDataSvc'
            '147' = 'UserManager'
            '148' = 'UsoSvc'
            '149' = 'VaultSvc'
            '150' = 'vds'
            '151' = 'vmicguestinterface'
            '152' = 'vmicheartbeat'
            '153' = 'vmickvpexchange'
            '154' = 'vmicrdv'
            '155' = 'vmicshutdown'
            '156' = 'vmictimesync'
            '157' = 'vmicvmsession'
            '158' = 'vmicvss'
            '159' = 'VSS'
            '160' = 'W32Time'
            '161' = 'WalletService'
            '162' = 'WbioSrvc'
            '163' = 'Wcmsvc'
            '164' = 'WdiServiceHost'
            '165' = 'WdiSystemHost'
            '166' = 'WdNisSvc'
            '167' = 'Wecsvc'
            '168' = 'WEPHOSTSVC'
            '169' = 'wercplsupport'
            '170' = 'WerSvc'
            '171' = 'WiaRpc'
            '172' = 'WinDefend'
            '173' = 'WinHttpAutoProxySvc'
            '174' = 'Winmgmt'
            '175' = 'WinRM'
            '176' = 'wisvc'
            '177' = 'wlidsvc'
            '178' = 'wmiApSrv'
            '179' = 'WPDBusEnum'
            '180' = 'WpnService'
            '181' = 'WpnUserService'
            '182' = 'WSearch'
            '183' = 'wuauserv'
            '184' = 'wudfsvc'
            '185' = 'XblAuthManager'
            '186' = 'XblGameSave'
        }
        DeploymentProperties     = @(
            'AssignmentName'
            'AssignmentID'
            'StartTime'
            'EnforcementDeadline'
            'TargetCollectionID'
            'ContainsExpiredUpdates'
            'Enabled'
            'CreationTime'
            'LastModifiedBy'
            'AssignedUpdateGroup'
            'AssignedCIs'
        )
        ErrorCode                = @(
            '80070057'
            '80070070'
            '8024001E'
            '800F0986'
        )
        EvaluationState          = [ordered]@{
            0  = 'None'
            8  = 'PendingSoftReboot'
            9  = 'PendingHardReboot'
            10 = 'WaitReboot'
            12 = 'InstallComplete'
            14 = 'WaitServiceWindow'
            15 = 'WaitUserLogon'
            16 = 'WaitUserLogoff'
            17 = 'WaitJobUserLogon'
            18 = 'WaitUserReconnect'
            19 = 'PendingUserLogoff'
            20 = 'PendingUpdate'
            1  = 'Available'
            2  = 'Submitted'
            3  = 'Detecting'
            4  = 'PreDownload'
            5  = 'Downloading'
            6  = 'WaitInstall'
            7  = 'Installing'
            11 = 'Verifying'
            13 = 'Error'
            21 = 'WaitingRetry'
            22 = 'WaitPresModeOff'
        }
        HTMLColors               = @(
            'AliceBlue',
            'AntiqueWhite',
            'Aqua',
            'Aquamarine',
            'Azure',
            'Beige',
            'Bisque',
            'Black',
            'BlanchedAlmond',
            'Blue',
            'BlueViolet',
            'Brown',
            'BurlyWood',
            'CadetBlue',
            'Chartreuse',
            'Chocolate',
            'Coral',
            'CornflowerBlue',
            'Cornsilk',
            'Crimson',
            'Cyan',
            'DarkBlue',
            'DarkCyan',
            'DarkGoldenrod',
            'DarkGray',
            'DarkGreen',
            'DarkKhaki',
            'DarkMagenta',
            'DarkOliveGreen',
            'DarkOrange',
            'DarkOrchid',
            'DarkRed',
            'DarkSalmon',
            'DarkSeaGreen',
            'DarkSlateBlue',
            'DarkSlateGray',
            'DarkTurquoise',
            'DarkViolet',
            'DeepPink',
            'DeepSkyBlue',
            'DimGray',
            'DodgerBlue',
            'FireBrick',
            'FloralWhite',
            'ForestGreen',
            'Fuchsia',
            'Gainsboro',
            'GhostWhite',
            'Gold',
            'Goldenrod',
            'Gray',
            'Green',
            'GreenYellow',
            'HoneyDew',
            'HotPink',
            'IndianRed',
            'Indigo',
            'Ivory',
            'Khaki',
            'Lavender',
            'LavenderBlush',
            'LawnGreen',
            'LemonChiffon',
            'LightBlue',
            'LightCoral',
            'LightCyan',
            'LightGoldenrodYellow',
            'LightGray',
            'LightGreen',
            'LightPink',
            'LightSalmon',
            'LightSeaGreen',
            'LightSkyBlue',
            'LightSlateGray',
            'LightSteelBlue',
            'LightYellow',
            'Lime',
            'LimeGreen',
            'Linen',
            'Magenta',
            'Maroon',
            'MediumAquamarine',
            'MediumBlue',
            'MediumOrchid',
            'MediumPurple',
            'MediumSeaGreen',
            'MediumSlateBlue',
            'MediumSpringGreen',
            'MediumTurquoise',
            'MediumVioletRed',
            'MidnightBlue',
            'MintCream',
            'MistyRose',
            'Mocassin',
            'NavajoWhite',
            'Navy',
            'OldLace',
            'Olive',
            'OliveDrab',
            'Orange',
            'OrangeRed',
            'Orchid',
            'PaleGoldenrod',
            'PaleGreen',
            'PaleTurquoise',
            'PaleVioletRed',
            'PapayaWhip',
            'PeachPuff',
            'Peru',
            'Pink',
            'Plum',
            'PowderBlue',
            'Purple',
            'RebeccaPurple',
            'Red',
            'RosyBrown',
            'RoyalBlue',
            'SaddleBrown',
            'Salmon',
            'SandyBrown',
            'SeaGreen',
            'SeaShell',
            'Sienna',
            'Silver',
            'SkyBlue',
            'SlateBlue',
            'Snow',
            'SpringGreen',
            'SteelBlue',
            'Tan',
            'Teal',
            'Thistle',
            'Tomato',
            'Turquoise',
            'Violet',
            'Wheat',
            'White',
            'WhiteSmoke',
            'Yellow',
            'YellowGreen'
        )
        InfoType                 = @(
            'Cluster'
            'DefaultService'
            'DesktopExperience'
            'DriveSpace'
            'IIS'
            'InstalledPatches'
            'PatchingStatus'
            'PatchWindow'
            'SCCM'
            'Service'
            'Task'
            'Win32Class'
        )
        OptionalFeature          = [ordered]@{
            0   = 'ActiveDirectory-PowerShell'
            1   = 'ADCertificateServicesRole'
            2   = 'AuthManager'
            3   = 'BitLocker'
            4   = 'Bitlocker-Utilities'
            5   = 'BITS'
            6   = 'BITSExtensions-Upload'
            7   = 'CCFFilter'
            8   = 'CertificateEnrollmentPolicyServer'
            9   = 'CertificateEnrollmentServer'
            10  = 'CertificateServices'
            11  = 'ClientForNFS-Infrastructure'
            12  = 'Containers'
            13  = 'CoreFileServer'
            14  = 'DataCenterBridging'
            15  = 'DataCenterBridging-LLDP-Tools'
            16  = 'Dedup-Core'
            17  = 'DeviceHealthAttestationService'
            18  = 'DFSN-Server'
            19  = 'DFSR-Infrastructure-ServerEdition'
            20  = 'DHCPServer'
            21  = 'DHCPServer-Tools'
            22  = 'DirectoryServices-ADAM'
            23  = 'DirectoryServices-ADAM-Tools'
            24  = 'DirectoryServices-AdministrativeCenter'
            25  = 'DirectoryServices-DomainController'
            26  = 'DirectoryServices-DomainController-Tools'
            27  = 'DiskIo-QoS'
            28  = 'DNS-Server-Full-Role'
            29  = 'DNS-Server-Tools'
            30  = 'DSC-Service'
            31  = 'EnhancedStorage'
            32  = 'FabricShieldedTools'
            33  = 'FailoverCluster-AdminPak'
            34  = 'FailoverCluster-AutomationServer'
            35  = 'FailoverCluster-CmdInterface'
            36  = 'FailoverCluster-FullServer'
            37  = 'FailoverCluster-PowerShell'
            38  = 'FileAndStorage-Services'
            39  = 'FileServerVSSAgent'
            40  = 'File-Services'
            41  = 'FRS-Infrastructure'
            42  = 'FSRM-Infrastructure'
            43  = 'FSRM-Infrastructure-Services'
            44  = 'HardenedFabricEncryptionTask'
            45  = 'HostGuardianService-Package'
            46  = 'IdentityServer-SecurityTokenService'
            47  = 'IIS-ApplicationDevelopment'
            48  = 'IIS-ApplicationInit'
            49  = 'IIS-ASP'
            50  = 'IIS-ASPNET'
            51  = 'IIS-ASPNET45'
            52  = 'IIS-BasicAuthentication'
            53  = 'IIS-CertProvider'
            54  = 'IIS-CGI'
            55  = 'IIS-ClientCertificateMappingAuthentication'
            56  = 'IIS-CommonHttpFeatures'
            57  = 'IIS-CustomLogging'
            58  = 'IIS-DefaultDocument'
            59  = 'IIS-DigestAuthentication'
            60  = 'IIS-DirectoryBrowsing'
            61  = 'IIS-FTPExtensibility'
            62  = 'IIS-FTPServer'
            63  = 'IIS-FTPSvc'
            64  = 'IIS-HealthAndDiagnostics'
            65  = 'IIS-HostableWebCore'
            66  = 'IIS-HttpCompressionDynamic'
            67  = 'IIS-HttpCompressionStatic'
            68  = 'IIS-HttpErrors'
            69  = 'IIS-HttpLogging'
            70  = 'IIS-HttpRedirect'
            71  = 'IIS-HttpTracing'
            72  = 'IIS-IIS6ManagementCompatibility'
            73  = 'IIS-IISCertificateMappingAuthentication'
            74  = 'IIS-IPSecurity'
            75  = 'IIS-ISAPIExtensions'
            76  = 'IIS-ISAPIFilter'
            77  = 'IIS-LegacyScripts'
            78  = 'IIS-LegacySnapIn'
            79  = 'IIS-LoggingLibraries'
            80  = 'IIS-ManagementConsole'
            81  = 'IIS-ManagementScriptingTools'
            82  = 'IIS-ManagementService'
            83  = 'IIS-Metabase'
            84  = 'IIS-NetFxExtensibility'
            85  = 'IIS-NetFxExtensibility45'
            86  = 'IIS-ODBCLogging'
            87  = 'IIS-Performance'
            88  = 'IIS-RequestFiltering'
            89  = 'IIS-RequestMonitor'
            90  = 'IIS-Security'
            91  = 'IIS-ServerSideIncludes'
            92  = 'IIS-StaticContent'
            93  = 'IIS-URLAuthorization'
            94  = 'IIS-WebDAV'
            95  = 'IIS-WebServer'
            96  = 'IIS-WebServerManagementTools'
            97  = 'IIS-WebServerRole'
            98  = 'IIS-WebSockets'
            99  = 'IIS-WindowsAuthentication'
            100 = 'IIS-WMICompatibility'
            101 = 'IPAMClientFeature'
            102 = 'IPAMServerFeature'
            103 = 'iSCSITargetServer'
            104 = 'iSCSITargetServer-PowerShell'
            105 = 'iSCSITargetStorageProviders'
            106 = 'iSNS_Service'
            107 = 'KeyDistributionService-PSH-Cmdlets'
            108 = 'Licensing'
            109 = 'LightweightServer'
            110 = 'ManagementOdata'
            111 = 'Microsoft-Hyper-V'
            112 = 'Microsoft-Hyper-V-Management-Clients'
            113 = 'Microsoft-Hyper-V-Management-PowerShell'
            114 = 'Microsoft-Hyper-V-Offline'
            115 = 'Microsoft-Hyper-V-Online'
            116 = 'Microsoft-Windows-FCI-Client-Package'
            117 = 'Microsoft-Windows-GroupPolicy-ServerAdminTools-Update'
            118 = 'MicrosoftWindowsPowerShell'
            119 = 'MicrosoftWindowsPowerShellRoot'
            120 = 'MicrosoftWindowsPowerShellV2'
            121 = 'Microsoft-Windows-Web-Services-for-Management-IIS-Extension'
            122 = 'MSMQ'
            123 = 'MSMQ-ADIntegration'
            124 = 'MSMQ-DCOMProxy'
            125 = 'MSMQ-HTTP'
            126 = 'MSMQ-Multicast'
            127 = 'MSMQ-RoutingServer'
            128 = 'MSMQ-Server'
            129 = 'MSMQ-Services'
            130 = 'MSMQ-Triggers'
            131 = 'MSRDC-Infrastructure'
            132 = 'MultipathIo'
            133 = 'MultiPoint-Connector'
            134 = 'MultiPoint-Connector-Services'
            135 = 'MultiPoint-Role'
            136 = 'MultiPoint-Tools'
            137 = 'NetFx3'
            138 = 'NetFx3ServerFeatures'
            139 = 'NetFx4'
            140 = 'NetFx4Extended-ASPNET45'
            141 = 'NetFx4ServerFeatures'
            142 = 'NetworkDeviceEnrollmentServices'
            143 = 'NetworkLoadBalancingFullServer'
            144 = 'OnlineRevocationServices'
            145 = 'P2P-PnrpOnly'
            146 = 'PeerDist'
            147 = 'PKIClient-PSH-Cmdlets'
            148 = 'Printing-Client'
            149 = 'Printing-Client-Gui'
            150 = 'Printing-LPDPrintService'
            151 = 'Printing-PrintToPDFServices-Features'
            152 = 'Printing-Server-Foundation-Features'
            153 = 'Printing-Server-Role'
            154 = 'Printing-XPSServices-Features'
            155 = 'QWAVE'
            156 = 'RasRoutingProtocols'
            157 = 'RemoteAccess'
            158 = 'RemoteAccessMgmtTools'
            159 = 'RemoteAccessPowerShell'
            160 = 'RemoteAccessServer'
            161 = 'Remote-Desktop-Services'
            162 = 'ResumeKeyFilter'
            163 = 'RightsManagementServices'
            164 = 'RightsManagementServices-AdminTools'
            165 = 'RightsManagementServices-Role'
            166 = 'RMS-Federation'
            167 = 'RPC-HTTP_Proxy'
            168 = 'RSAT-ADDS-Tools-Feature'
            169 = 'RSAT-AD-Tools-Feature'
            170 = 'RSAT-Hyper-V-Tools-Feature'
            171 = 'SBMgr-UI'
            172 = 'ServerCore-Drivers-General'
            173 = 'ServerCore-Drivers-General-WOW64'
            174 = 'ServerCore-EA-IME'
            175 = 'ServerCore-EA-IME-WOW64'
            176 = 'ServerCore-WOW64'
            177 = 'ServerForNFS-Infrastructure'
            178 = 'ServerManager-Core-RSAT'
            179 = 'ServerManager-Core-RSAT-Feature-Tools'
            180 = 'ServerManager-Core-RSAT-Role-Tools'
            181 = 'ServerMediaFoundation'
            182 = 'ServerMigration'
            183 = 'Server-Psh-Cmdlets'
            184 = 'ServicesForNFS-ServerAndClient'
            185 = 'SessionDirectory'
            186 = 'SetupAndBootEventCollection'
            187 = 'ShieldedVMToolsAdminPack'
            188 = 'SimpleTCP'
            189 = 'SMB1Protocol'
            190 = 'SMBBW'
            191 = 'SmbDirect'
            192 = 'SMBHashGeneration'
            193 = 'SmbWitness'
            194 = 'Smtpsvc-Admin-Update-Name'
            195 = 'Smtpsvc-Service-Update-Name'
            196 = 'SNMP'
            197 = 'Storage-Replica-AdminPack'
            198 = 'Storage-Services'
            199 = 'TelnetClient'
            200 = 'TlsSessionTicketKey-PSH-Cmdlets'
            201 = 'Tpm-PSH-Cmdlets'
            202 = 'UpdateServices'
            203 = 'UpdateServices-API'
            204 = 'UpdateServices-Database'
            205 = 'UpdateServices-RSAT'
            206 = 'UpdateServices-Services'
            207 = 'UpdateServices-WidDatabase'
            208 = 'VmHostAgent'
            209 = 'VolumeActivation-Full-Role'
            210 = 'WAS-ConfigurationAPI'
            211 = 'WAS-NetFxEnvironment'
            212 = 'WAS-ProcessModel'
            213 = 'WAS-WindowsActivationService'
            214 = 'WCF-HTTP-Activation'
            215 = 'WCF-HTTP-Activation45'
            216 = 'WCF-MSMQ-Activation45'
            217 = 'WCF-NonHTTP-Activation'
            218 = 'WCF-Pipe-Activation45'
            219 = 'WCF-Services45'
            220 = 'WCF-TCP-Activation45'
            221 = 'WCF-TCP-PortSharing45'
            222 = 'WebAccess'
            223 = 'Web-Application-Proxy'
            224 = 'WebEnrollmentServices'
            225 = 'Windows-Defender'
            226 = 'Windows-Defender-Features'
            227 = 'Windows-Internal-Database'
            228 = 'WindowsPowerShellWebAccess'
            229 = 'WindowsServerBackup'
            230 = 'WindowsStorageManagementService'
            231 = 'WINSRuntime'
            232 = 'WINS-Server-Tools'
            233 = 'WMISnmpProvider'
            234 = 'WorkFolders-Server'
            235 = 'WSS-Product-Package'
        }
        OutputType               = @(
            'String'
            'Json'
        )
        PatchDescription         = [ordered]@{
            '0' = 'Hotfix'
            '1' = 'Security Update'
            '2' = 'Update'
        }
        PatchingDrive            = [ordered]@{
            Credential = $null
            Name       = 'Patching'
            PSProvider = 'FileSystem'
            Root       = $env:TEMP
            Scope      = 'Global'
        }
        ResultFormat             = @(
            'Console'
            'Email'
            'File'
            'ServiceDesk'
            'SharePoint'
            'Teams'
        )
        ServiceStatus            = [ordered]@{
            0 = 'None'
            1 = 'Stopped'
            2 = 'StartPending'
            3 = 'StopPending'
            4 = 'Running'
            5 = 'ContinuePending'
            6 = 'PausePending'
            7 = 'Paused'
        }
        TaskState                = [ordered]@{
            0 = 'Unknown'
            1 = 'Disabled'
            2 = 'Queued'
            3 = 'Ready'
            4 = 'Running'
        }
        Win32Class               = @(
            'BaseBoard'
            'BIOS'
            'ComputerSystem'
            'ComputerSystemProduct'
            'DiskDrive'
            'DiskPartition'
            'LogicalDisk'
            'NetworkAdapter'
            'OperatingSystem'
            'OptionalFeature'
            'PhysicalMemory'
            'Process'
            'Processor'
        )
    }
    $script:Config | Export-Clixml -Path "$PSScriptRoot\config.xml" -Depth 100
}

try {
    $script:DeploymentConfiguration = Import-Clixml -Path "$PSScriptRoot\DeploymentConfiguration.xml" -ErrorAction Stop
}
catch {
    $script:DeploymentConfiguration = [ordered]@{
        Test        = [ordered]@{
            AllowRestart                = $false
            DeploymentType              = 'Required'
            DownloadFromMicrosoftUpdate = $true
            ProtectedType               = 'RemoteDistributionPoint'
            RequirePostRebootFullScan   = $true
            RestartServer               = $false
            RestartWorkstation          = $false
            SoftwareInstallation        = $false
            TimeBasedOn                 = 'LocalTime'
            UnprotectedType             = 'UnprotectedDistributionPoint'
            UseMeteredNetwork           = $false
            VerbosityLevel              = 'OnlySuccessAndErrorMessages'
        }
        Development = [ordered]@{
            AllowRestart                = $false
            DeploymentType              = 'Required'
            DownloadFromMicrosoftUpdate = $true
            ProtectedType               = 'RemoteDistributionPoint'
            RequirePostRebootFullScan   = $true
            RestartServer               = $false
            RestartWorkstation          = $false
            SoftwareInstallation        = $false
            TimeBasedOn                 = 'LocalTime'
            UnprotectedType             = 'UnprotectedDistributionPoint'
            UseMeteredNetwork           = $false
            VerbosityLevel              = 'OnlySuccessAndErrorMessages'
        }
        QA = [ordered]@{
            AllowRestart                = $false
            DeploymentType              = 'Required'
            DownloadFromMicrosoftUpdate = $true
            ProtectedType               = 'RemoteDistributionPoint'
            RequirePostRebootFullScan   = $true
            RestartServer               = $false
            RestartWorkstation          = $false
            SoftwareInstallation        = $false
            TimeBasedOn                 = 'LocalTime'
            UnprotectedType             = 'UnprotectedDistributionPoint'
            UseMeteredNetwork           = $false
            VerbosityLevel              = 'OnlySuccessAndErrorMessages'
        }
        Production = [ordered]@{
            AllowRestart                = $false
            DeploymentType              = 'Required'
            DownloadFromMicrosoftUpdate = $true
            ProtectedType               = 'RemoteDistributionPoint'
            RequirePostRebootFullScan   = $true
            RestartServer               = $true
            RestartWorkstation          = $true
            SoftwareInstallation        = $false
            TimeBasedOn                 = 'LocalTime'
            UnprotectedType             = 'UnprotectedDistributionPoint'
            UseMeteredNetwork           = $false
            VerbosityLevel              = 'OnlySuccessAndErrorMessages'
        }
    }
    $script:DeploymentConfiguration | Export-Clixml -Path "$PSScriptRoot\DeploymentConfiguration.xml" -Depth 100
}
#endregion

#region DotSourcedScripts
################################################################################
#                                                                              #
#                           DOT-SOURCED SCRIPTS                                #
#                                                                              #
################################################################################
. "$PSScriptRoot\Get-PatchingStatus.ps1"
. "$PSScriptRoot\Get-ServerInfo.ps1"
. "$PSScriptRoot\New-MonthlyPatchDeployment.ps1"
. "$PSScriptRoot\Restart-ClusterServer.ps1"
. "$PSScriptRoot\Restart-StandardServer.ps1"
. "$PSScriptRoot\Set-ServerInfo.ps1"
. "$PSScriptRoot\Start-AutomaticRemediation.ps1"
#endregion

#region ModuleMembers
################################################################################
#                                                                              #
#                              MODULE MEMBERS                                  #
#                                                                              #
################################################################################
Export-ModuleMember -Function Get-PatchingStatus
Export-ModuleMember -Function Get-ServerInfo
Export-ModuleMember -Function New-MonthlyPatchDeployment
Export-ModuleMember -Function Restart-ClusterServer
Export-ModuleMember -Function Restart-StandardServer
Export-ModuleMember -Function Set-ServerInfo
Export-ModuleMember -Function Start-AutomaticRemediation
#endregion
