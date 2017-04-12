
Namespace Core

    Public MustInherit Class ShutdownCore

#Region "Windows API Function Decleration and its constants"
        Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Integer, ByVal dwReason As Integer) As Boolean

        'Windows API - Shutdown Type Constants (uFlags)
        Private Const EWX_LOGOFF As Integer = 0
        Private Const EWX_SHUTDOWN As Integer = 1
        Private Const EWX_REBOOT As Integer = 2
        Private Const EWX_FORCE As Integer = 4
        Private Const EWX_POWEROFF As Integer = 8
        Private Const EWX_FORCEIFHUNG As Integer = 16
        Private Const EWX_RESET As Integer = EWX_REBOOT Or EWX_FORCE
        Private Const EWX_POWERDOWN As Integer = EWX_POWEROFF Or EWX_SHUTDOWN
        Private Const ENDSESSION_LOGOFF As Integer = &H80000000
#End Region

#Region "Shutdown Reasons"
        'https://msdn.microsoft.com/en-us/library/aa376885(VS.85).aspx

        ''' <summary>Major reason flags. They indicate the general issue type</summary>
        <Flags>
        Public Enum ShutdownReasonMajor
            'Reference: MSDN - System Shutdown Reason Codes
            Application = &H40000
            Hardware = &H10000
            'LegacyApi = &H70000    'MSDN: This shutdown was initiated by the legacy InitiateSystemShutdown function. Applications should use the InitiateSystemShutdownEx function.
            OperatingSystem = &H20000
            Other = &H0
            Power = &H60000
            Software = &H30000
            System = &H50000
        End Enum

        ''' <summary>Minor reason flags. They modify the specified minor reason flag</summary>
        <Flags>
        Public Enum ShutdownReasonMinor
            'Reference: MSDN - System Shutdown Reason Codes
            BlueScreen = &HF
            CordUnplugged = &HB
            Disk = &H7
            Environment = &HC
            HardwareDriver = &HD
            Hotfix = &H11
            Hung = &H5
            Installation = &H2
            Maintenance = &H1
            MMC = &H19
            NetworkConnectivity = &H14
            NetworkCard = &H9
            Other = &H0
            OtherDriver = &HE
            PowerSupply = &HA
            Processor = &H8
            Reconfig = &H4
            Security = &H13
            SecurityFix = &H12
            SecurityFixUninstall = &H18
            ServicePack = &H10
            ServicePackUninstall = &H16
            TerminalServices = &H20
            Unstable = &H6
            Upgrade = &H3
            WMI = &H15
        End Enum

        ''' <summary>Optional flags provide additional information about the event</summary>
        <Flags>
        Public Enum ShutdownReasonOptional
            'Reference: MSDN - System Shutdown Reason Codes
            ''' <summary>The reason code is defined by the user in Windows Registry</summary>
            UserDefined = &H40000000
            ''' <summary>"The shutdown was planned. The system generates a System State Data (SSD) file. This file contains system state information such as the processes, threads, memory usage, and configuration</summary>
            Planned = &H80000000
            ''' <summary>"The shutdown was unplanned </summary>
            Unplanned = 0
        End Enum


        ''' <summary>Combinations are recognized by the system. The members indicate the string that is displayed in the Shutdown Event Tracker</summary>
        <Flags>
        Public Enum ShutdownReasonPredefined
            'Reference: MSDN - System Shutdown Reason Codes

            ''' <summary>"Application: Unresponsive" An unplanned restart or shutdown to troubleshoot an unresponsive application</summary>
            Application_Unresponsive = ShutdownReasonMajor.Application Or ShutdownReasonMinor.Hung

            ''' <summary>"Application: Installation (Planned)" A planned restart or shutdown to perform application installation </summary>
            Application_Installation_Planned = ShutdownReasonMajor.Application Or ShutdownReasonMinor.Installation Or ShutdownReasonOptional.Planned

            ''' <summary>"Application: Maintenance (Unplanned)" An unplanned restart or shutdown to service an application</summary>
            Application_Maintenance_Unplanned = ShutdownReasonMajor.Application Or ShutdownReasonMinor.Maintenance

            ''' <summary>"Application: Maintenance (Planned)" A planned restart or shutdown to perform planned maintenance on an application</summary>
            Application_Maintenance_Planned = ShutdownReasonMajor.Application Or ShutdownReasonMinor.Maintenance Or ShutdownReasonOptional.Planned

            ''' <summary>"Application: Unstable" An unplanned restart or shutdown to troubleshoot an unstable application</summary>
            Application_Unstable = ShutdownReasonMajor.Application Or ShutdownReasonMinor.Unstable

            ''' <summary>"Hardware: Installation (Unplanned)" An unplanned restart or shutdown to begin or complete hardware installation</summary>
            Hardware_Installation_Unplanned = ShutdownReasonMajor.Hardware Or ShutdownReasonMinor.Installation

            ''' <summary>"Hardware: Installation (Planned)" A planned restart or shutdown to begin or complete hardware installation</summary>
            Hardware_Installation_Planned = ShutdownReasonMajor.Hardware Or ShutdownReasonMinor.Installation Or ShutdownReasonOptional.Planned

            ''' <summary>"Hardware: Maintenance (Unplanned)" An unplanned restart or shutdown to service hardware on the system</summary>
            Hardware_Maintenance_Unplanned = ShutdownReasonMajor.Hardware Or ShutdownReasonMinor.Maintenance

            ''' <summary>"Hardware: Maintenance (Planned)" A planned restart or shutdown to service hardware on the system</summary>
            Hardware_Maintenance_Planned = ShutdownReasonMajor.Hardware Or ShutdownReasonMinor.Maintenance Or ShutdownReasonOptional.Planned

            ''' <summary>"Operating System: Hot fix (Unplanned)" An unplanned restart or shutdown to install a hot fix</summary>
            OperatingSystem_Hotfix_Unplanned = ShutdownReasonMajor.OperatingSystem Or ShutdownReasonMinor.Hotfix

            ''' <summary>"Operating System: Hot fix (Planned)" A planned restart or shutdown to install a hot fix</summary>
            OperatingSystem_Hotfix_Planned = ShutdownReasonMajor.OperatingSystem Or ShutdownReasonMinor.Hotfix

            ''' <summary>"Operating System: Reconfiguration (Unplanned)" An unplanned restart or shutdown to change the operating system configuration</summary>
            OperatingSystem_Reconfig_Unplanned = ShutdownReasonMajor.OperatingSystem Or ShutdownReasonMinor.Reconfig

            ''' <summary>"Operating System: Reconfiguration (Planned)" A planned restart or shutdown to change the operating system configuration</summary>
            OperatingSystem_Reconfig_Planned = ShutdownReasonMajor.OperatingSystem Or ShutdownReasonMinor.Reconfig Or ShutdownReasonOptional.Planned

            ''' <summary>"Operating System: Security fix (Unplanned)" An unplanned restart or shutdown to install a security patch</summary>
            OperatingSystem_SecurityFix_Unplanned = ShutdownReasonMajor.OperatingSystem Or ShutdownReasonMinor.SecurityFix

            ''' <summary>"Operating System: Security fix (Planned)" A planned restart or shutdown to install a security patch</summary>
            OperatingSystem_SecurityFix_Planned = ShutdownReasonMajor.OperatingSystem Or ShutdownReasonMinor.SecurityFix Or ShutdownReasonOptional.Planned

            ''' <summary>"Operating System: Service pack (Planned)" A planned restart or shutdown to install a service pack</summary>
            OperatingSystem_ServicePack_Planned = ShutdownReasonMajor.OperatingSystem Or ShutdownReasonMinor.ServicePack Or ShutdownReasonOptional.Planned

            ''' <summary>"Operating System: Upgrade (Planned)" A planned restart or shutdown to upgrade the operating system configuration</summary>
            OperatingSystem_Upgrade_Planned = ShutdownReasonMajor.OperatingSystem Or ShutdownReasonMinor.Upgrade Or ShutdownReasonOptional.Planned

            ''' <summary>"Other (Unplanned)" An unplanned shutdown or restart</summary>
            Other_Unplanned = ShutdownReasonMajor.Other Or ShutdownReasonMinor.Other

            ''' <summary>"Other (Planned)" A planned shutdown or restart</summary>
            Other_Planned = ShutdownReasonMajor.Other Or ShutdownReasonMinor.Other Or ShutdownReasonOptional.Planned

            ''' <summary>"Other Failure: System Unresponsive" The system became unresponsive</summary>
            Other_Failure_SystemUnresponsive = ShutdownReasonMajor.Other Or ShutdownReasonMinor.Hung

            ''' <summary>"Power Failure: Cord Unplugged" The computer was unplugged</summary>
            Power_Failure_CordUnplugged = ShutdownReasonMajor.Power Or ShutdownReasonMinor.CordUnplugged

            ''' <summary>"Power Failure: Environment" There was a power outage</summary>
            Power_Failure_Environment = ShutdownReasonMajor.Power Or ShutdownReasonMinor.Environment

            ''' <summary>"System Failure: Stop error" The computer displayed a blue screen crash event</summary>
            System_Failure_StopErrorBlueScreen = ShutdownReasonMajor.System Or ShutdownReasonMinor.BlueScreen

            ''' <summary>"Loss of network connectivity (Unplanned)" The computer needs to be shut down due to a network connectivity issue</summary>
            System_LossNetworkConnectivity_Unplanned = ShutdownReasonMajor.System Or ShutdownReasonMinor.NetworkConnectivity

            ''' <summary>"Security issue" The computer needs to be shut down due to a security issue</summary>
            System_SecurityIssue = ShutdownReasonMajor.System Or ShutdownReasonMinor.Security

        End Enum
#End Region
        ''' <summary>Shutdown types</summary>
        <Flags>
        Public Enum ShutdownType
            LogOff = EWX_LOGOFF
            Shutdown = EWX_SHUTDOWN
            Reboot = EWX_REBOOT
            PowerOff = EWX_POWEROFF
        End Enum

        Public Enum ForceType
            ''' <summary>Force option is disabled</summary>
            Never
            ''' <summary>Forces processes to terminate if they do not respond to the WM_QUERYENDSESSION or WM_ENDSESSION message within the timeout interval</summary>
            ForceIfHung
            ''' <summary> System does not send the WM_QUERYENDSESSION message. This can cause applications to lose data! Therefore, you should only use this flag in an emergency!</summary>
            ForceAll
        End Enum

        ''' <summary>Logs off the interactive user, shuts down the system, or shuts down and restarts the system. It sends the WM_QUERYENDSESSION message to all applications to determine if they can be terminated</summary>
        ''' <param name="ShutdownType">Shutdown Type: Log Off, Shutdown, Reboot, Power Off</param>
        ''' <param name="ForceOption">Force type</param>
        ''' <param name="ShutdownReason">Predefined shutdown reasons are recognized by the system. The members indicate the string that is displayed in the Shutdown Event Tracker</param>
        ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
        ''' <exception cref="ShutdownPrivilegeException"></exception>
        Protected Overridable Overloads Function ShutdownSystem(ShutdownType As ShutdownType, ForceOption As ForceType, ShutdownReason As ShutdownReasonPredefined) As Boolean
            Return ShutdownSystem(ShutdownType, ForceOption, Convert.ToInt32(ShutdownReason))
        End Function

        ''' <summary>Logs off the interactive user, shuts down the system, or shuts down and restarts the system. It sends the WM_QUERYENDSESSION message to all applications to determine if they can be terminated</summary>
        ''' <param name="ShutdownType">Shutdown Type: Log Off, Shutdown, Reboot, Power Off</param>
        ''' <param name="ForceOption">Force type</param>
        ''' <param name="ReasonMajor">Major reason flag</param>
        ''' <param name="ReasonMinor">Minor reason flag</param>
        ''' <param name="ReasonOptional">Optional flag that provide additional information about the event</param>
        ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
        ''' <exception cref="ShutdownPrivilegeException"></exception>
        Protected Overridable Overloads Function ShutdownSystem(ShutdownType As ShutdownType, ForceOption As ForceType, ReasonMajor As ShutdownReasonMajor, ReasonMinor As ShutdownReasonMinor, ReasonOptional As ShutdownReasonOptional) As Boolean
            Dim Reason As Integer
            Reason = Convert.ToInt32(ReasonMajor) Or Convert.ToInt32(ReasonMinor) Or Convert.ToInt32(ReasonOptional)
            Return ShutdownSystem(ShutdownType, ForceOption, Reason)
        End Function

        Private Overloads Function ShutdownSystem(ShutdownType As ShutdownType, ForceOption As ForceType, Reason As Integer) As Boolean
            Dim TokenPriv As New TokenPrivilegeCore
            Dim ShutdownFlags As Integer
            Dim RetVal As Boolean = False

            ShutdownFlags = Convert.ToInt32(ShutdownType)

            If ForceOption = ForceType.ForceIfHung Then
                ShutdownFlags = ShutdownFlags Or EWX_FORCEIFHUNG
            ElseIf ForceOption = ForceType.ForceAll Then
                ShutdownFlags = ShutdownFlags Or EWX_FORCEIFHUNG Or EWX_FORCE
            End If

            RetVal = TokenPriv.GetShutdownPrivilege(TokenPrivilegeCore.ShutdownPrivilege.CurrentMachine)

            If RetVal Then
                RetVal = ExitWindowsEx(ShutdownFlags, Reason)
            End If

            Return RetVal
        End Function

    End Class

End Namespace



