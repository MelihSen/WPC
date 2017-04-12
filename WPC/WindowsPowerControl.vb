Imports WPC.Core

Public Class WindowsPowerControl
    Inherits ShutdownCore
    Implements ISuspend

    Private Suspendc As New SuspendCore

    ''' <summary>Determines whether the computer supports hibernation </summary>
    ''' <returns>If the computer supports hibernation (power state S4) and the file Hiberfil.sys is present on the system, the property returns True. Otherwise, the property returns False</returns>
    Public ReadOnly Property HibernateSupported As Boolean Implements ISuspend.HibernateSupported
        Get
            Return Suspendc.HibernateSupported
        End Get
    End Property

    ''' <summary>Determines whether the computer supports the sleep states </summary>
    ''' <returns>If the computer supports the sleep states (S1, S2, and S3), the propery returns True. Otherwise, the property returns False</returns>
    Public ReadOnly Property SleepSupported As Boolean Implements ISuspend.SleepSupported
        Get
            Return Suspendc.SleepSupported
        End Get
    End Property

    Private DefaultForceType As ForceType = ForceType.Never
    ''' <summary>Determines default force type for functions have not force parameter</summary>
    ''' <returns>Returns default force type</returns>
    Public Property DefaultForceOption As ForceType
        Get
            Return DefaultForceType
        End Get
        Set(ByVal value As ForceType)
            DefaultForceType = value
        End Set
    End Property

    Private DefaultShutdownReasonValue As ShutdownReasonPredefined = ShutdownReasonPredefined.Other_Planned
    ''' <summary>Determines default shutdown reason for functions have not shutdown reason parameters</summary>
    ''' <returns>Returns default force type</returns>
    Public Property DefaultShutdownReason As ShutdownReasonPredefined
        Get
            Return DefaultShutdownReasonValue
        End Get
        Set(ByVal value As ShutdownReasonPredefined)
            DefaultShutdownReasonValue = value
        End Set
    End Property


#Region "Shutdown Functions"
    ''' <summary>Shuts down the system to a point at which it is safe to turn off the power. All file buffers have been flushed to disk, and all running processes have stopped</summary>
    ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
    ''' <exception cref="ShutdownPrivilegeException"></exception>
    Public Overloads Function Shutdown() As Boolean
        Return ShutdownSystem(ShutdownType.Shutdown, DefaultForceType, DefaultShutdownReasonValue)
    End Function

    ''' <summary>Shuts down the system to a point at which it is safe to turn off the power. All file buffers have been flushed to disk, and all running processes have stopped</summary>
    ''' <param name="ForceOption">Force type</param>
    ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
    ''' <exception cref="ShutdownPrivilegeException"></exception>
    Public Overloads Function Shutdown(ForceOption As ForceType) As Boolean
        Return ShutdownSystem(ShutdownType.Shutdown, ForceOption, DefaultShutdownReasonValue)
    End Function

    ''' <summary>Shuts down the system to a point at which it is safe to turn off the power. All file buffers have been flushed to disk, and all running processes have stopped</summary>
    ''' <param name="ForceOption">Force type</param>
    ''' <param name="ShutdownReason">Predefined shutdown reasons are recognized by the system. The members indicate the string that is displayed in the Shutdown Event Tracker</param>
    ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
    ''' <exception cref="ShutdownPrivilegeException"></exception>
    Public Overloads Function Shutdown(ForceOption As ForceType, ShutdownReason As ShutdownReasonPredefined) As Boolean
        Return ShutdownSystem(ShutdownType.Shutdown, ForceOption, ShutdownReason)
    End Function

    ''' <summary>Shuts down the system to a point at which it is safe to turn off the power. All file buffers have been flushed to disk, and all running processes have stopped</summary>
    ''' <param name="ForceOption">Force type</param>
    ''' <param name="ReasonMajor">Major reason flag</param>
    ''' <param name="ReasonMinor">Minor reason flag</param>
    ''' <param name="Planned">True: Planned shutdown / False: Unplanned shutdown</param>
    ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
    ''' <exception cref="ShutdownPrivilegeException"></exception>
    Public Overloads Function Shutdown(ForceOption As ForceType, ReasonMajor As ShutdownReasonMajor, ReasonMinor As ShutdownReasonMinor, Planned As Boolean) As Boolean
        Return ShutdownSystem(ShutdownType.Shutdown, ForceOption, ReasonMajor, ReasonMinor, ConvertPlannedToOptionalReason(Planned))
    End Function

    ''' <summary>Shuts down the system to a point at which it is safe to turn off the power. All file buffers have been flushed to disk, and all running processes have stopped</summary>
    ''' <param name="ForceOption">Force type</param>
    ''' <param name="ReasonMajor">Major reason flag</param>
    ''' <param name="ReasonMinor">Minor reason flag</param>
    ''' <param name="ReasonOptional">Optional flag that provide additional information about the event</param>
    ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
    ''' <exception cref="ShutdownPrivilegeException"></exception>
    Public Overloads Function Shutdown(ForceOption As ForceType, ReasonMajor As ShutdownReasonMajor, ReasonMinor As ShutdownReasonMinor, ReasonOptional As ShutdownReasonOptional) As Boolean
        Return ShutdownSystem(ShutdownType.Shutdown, ForceOption, ReasonMajor, ReasonMinor, ReasonOptional)
    End Function
#End Region

#Region "PowerOff Functions"
    ''' <summary>Shuts down the system and turns off the power. The system must support the power-off feature</summary>
    ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
    ''' <exception cref="ShutdownPrivilegeException"></exception>
    Public Overloads Function PowerOff() As Boolean
        Return ShutdownSystem(ShutdownType.PowerOff, DefaultForceType, DefaultShutdownReasonValue)
    End Function

    ''' <summary>Shuts down the system and turns off the power. The system must support the power-off feature</summary>
    ''' <param name="ForceOption">Force type</param>
    ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
    ''' <exception cref="ShutdownPrivilegeException"></exception>
    Public Overloads Function PowerOff(ForceOption As ForceType) As Boolean
        Return ShutdownSystem(ShutdownType.PowerOff, ForceOption, DefaultShutdownReasonValue)
    End Function

    ''' <summary>Shuts down the system and turns off the power. The system must support the power-off feature</summary>
    ''' <param name="ForceOption">Force type</param>
    ''' <param name="ShutdownReason">Predefined shutdown reasons are recognized by the system. The members indicate the string that is displayed in the Shutdown Event Tracker</param>
    ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
    ''' <exception cref="ShutdownPrivilegeException"></exception>
    Public Overloads Function PowerOff(ForceOption As ForceType, ShutdownReason As ShutdownReasonPredefined) As Boolean
        Return ShutdownSystem(ShutdownType.PowerOff, ForceOption, ShutdownReason)
    End Function

    ''' <summary>Shuts down the system and turns off the power. The system must support the power-off feature</summary>
    ''' <param name="ForceOption">Force type</param>
    ''' <param name="ReasonMajor">Major reason flag</param>
    ''' <param name="ReasonMinor">Minor reason flag</param>
    ''' <param name="Planned">True: Planned shutdown / False: Unplanned shutdown</param>
    ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
    ''' <exception cref="ShutdownPrivilegeException"></exception>
    Public Overloads Function PowerOff(ForceOption As ForceType, ReasonMajor As ShutdownReasonMajor, ReasonMinor As ShutdownReasonMinor, Planned As Boolean) As Boolean
        Return ShutdownSystem(ShutdownType.PowerOff, ForceOption, ReasonMajor, ReasonMinor, ConvertPlannedToOptionalReason(Planned))
    End Function

    ''' <summary>Shuts down the system and turns off the power. The system must support the power-off feature</summary>
    ''' <param name="ForceOption">Force type</param>
    ''' <param name="ReasonMajor">Major reason flag</param>
    ''' <param name="ReasonMinor">Minor reason flag</param>
    ''' <param name="ReasonOptional">Optional flag that provide additional information about the event</param>
    ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
    ''' <exception cref="ShutdownPrivilegeException"></exception>
    Public Overloads Function PowerOff(ForceOption As ForceType, ReasonMajor As ShutdownReasonMajor, ReasonMinor As ShutdownReasonMinor, ReasonOptional As ShutdownReasonOptional) As Boolean
        Return ShutdownSystem(ShutdownType.PowerOff, ForceOption, ReasonMajor, ReasonMinor, ReasonOptional)
    End Function
#End Region

#Region "Restart Functions"
    ''' <summary>Shuts down the system and then restarts the system</summary>
    ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
    ''' <exception cref="ShutdownPrivilegeException"></exception>
    Public Overloads Function Restart() As Boolean
        Return ShutdownSystem(ShutdownType.Reboot, DefaultForceType, DefaultShutdownReasonValue)
    End Function

    ''' <summary>Shuts down the system and then restarts the system</summary>
    ''' <param name="ForceOption">Force type</param>
    ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
    ''' <exception cref="ShutdownPrivilegeException"></exception>
    Public Overloads Function Restart(ForceOption As ForceType) As Boolean
        Return ShutdownSystem(ShutdownType.Reboot, ForceOption, DefaultShutdownReasonValue)
    End Function

    ''' <summary>Shuts down the system and then restarts the system</summary>
    ''' <param name="ForceOption">Force type</param>
    ''' <param name="ShutdownReason">Predefined shutdown reasons are recognized by the system. The members indicate the string that is displayed in the Shutdown Event Tracker</param>
    ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
    ''' <exception cref="ShutdownPrivilegeException"></exception>
    Public Overloads Function Restart(ForceOption As ForceType, ShutdownReason As ShutdownReasonPredefined) As Boolean
        Return ShutdownSystem(ShutdownType.Reboot, ForceOption, ShutdownReason)
    End Function

    ''' <summary>Shuts down the system and then restarts the system</summary>
    ''' <param name="ForceOption">Force type</param>
    ''' <param name="ReasonMajor">Major reason flag</param>
    ''' <param name="ReasonMinor">Minor reason flag</param>
    ''' <param name="Planned">True: Planned shutdown / False: Unplanned shutdown</param>
    ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
    ''' <exception cref="ShutdownPrivilegeException"></exception>
    Public Overloads Function Restart(ForceOption As ForceType, ReasonMajor As ShutdownReasonMajor, ReasonMinor As ShutdownReasonMinor, Planned As Boolean) As Boolean
        Return ShutdownSystem(ShutdownType.Reboot, ForceOption, ReasonMajor, ReasonMinor, ConvertPlannedToOptionalReason(Planned))
    End Function

    ''' <summary>Shuts down the system and then restarts the system</summary>
    ''' <param name="ForceOption">Force type</param>
    ''' <param name="ReasonMajor">Major reason flag</param>
    ''' <param name="ReasonMinor">Minor reason flag</param>
    ''' <param name="ReasonOptional">Optional flag that provide additional information about the event</param>
    ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
    ''' <exception cref="ShutdownPrivilegeException"></exception>
    Public Overloads Function Restart(ForceOption As ForceType, ReasonMajor As ShutdownReasonMajor, ReasonMinor As ShutdownReasonMinor, ReasonOptional As ShutdownReasonOptional) As Boolean
        Return ShutdownSystem(ShutdownType.Reboot, ForceOption, ReasonMajor, ReasonMinor, ReasonOptional)
    End Function
#End Region

#Region "LogOff Functions"
    ''' <summary>Shuts down all processes running in the logon session. Then it logs the user off</summary>
    ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
    ''' <exception cref="ShutdownPrivilegeException"></exception>
    Public Overloads Function LogOff() As Boolean
        Return ShutdownSystem(ShutdownType.LogOff, DefaultForceType, DefaultShutdownReasonValue)
    End Function

    ''' <summary>Shuts down all processes running in the logon session. Then it logs the user off</summary>
    ''' <param name="ForceOption">Force type</param>
    ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
    ''' <exception cref="ShutdownPrivilegeException"></exception>
    Public Overloads Function LogOff(ForceOption As ForceType) As Boolean
        Return ShutdownSystem(ShutdownType.LogOff, ForceOption, DefaultShutdownReasonValue)
    End Function

    ''' <summary>Shuts down all processes running in the logon session. Then it logs the user off</summary>
    ''' <param name="ForceOption">Force type</param>
    ''' <param name="ShutdownReason">Predefined shutdown reasons are recognized by the system. The members indicate the string that is displayed in the Shutdown Event Tracker</param>
    ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
    ''' <exception cref="ShutdownPrivilegeException"></exception>
    Public Overloads Function LogOff(ForceOption As ForceType, ShutdownReason As ShutdownReasonPredefined) As Boolean
        Return ShutdownSystem(ShutdownType.LogOff, ForceOption, ShutdownReason)
    End Function

    ''' <summary>Shuts down all processes running in the logon session. Then it logs the user off</summary>
    ''' <param name="ForceOption">Force type</param>
    ''' <param name="ReasonMajor">Major reason flag</param>
    ''' <param name="ReasonMinor">Minor reason flag</param>
    ''' <param name="Planned">True: Planned shutdown / False: Unplanned shutdown</param>
    ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
    ''' <exception cref="ShutdownPrivilegeException"></exception>
    Public Overloads Function LogOff(ForceOption As ForceType, ReasonMajor As ShutdownReasonMajor, ReasonMinor As ShutdownReasonMinor, Planned As Boolean) As Boolean
        Return ShutdownSystem(ShutdownType.LogOff, ForceOption, ReasonMajor, ReasonMinor, ConvertPlannedToOptionalReason(Planned))
    End Function

    ''' <summary>Shuts down all processes running in the logon session. Then it logs the user off</summary>
    ''' <param name="ForceOption">Force type</param>
    ''' <param name="ReasonMajor">Major reason flag</param>
    ''' <param name="ReasonMinor">Minor reason flag</param>
    ''' <param name="ReasonOptional">Optional flag that provide additional information about the event</param>
    ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
    ''' <exception cref="ShutdownPrivilegeException"></exception>
    Public Overloads Function LogOff(ForceOption As ForceType, ReasonMajor As ShutdownReasonMajor, ReasonMinor As ShutdownReasonMinor, ReasonOptional As ShutdownReasonOptional) As Boolean
        Return ShutdownSystem(ShutdownType.LogOff, ForceOption, ReasonMajor, ReasonMinor, ReasonOptional)
    End Function
#End Region

#Region "Sleep Functions"
    ''' <summary>Suspends the system by shutting power down. The system enters a sleep state (S1, S2, or S3)</summary>
    ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
    ''' <exception cref="SuspendNotSupportedException">Sleep is not supported by the system. Check SleepSupported property</exception>
    ''' <exception cref="ShutdownPrivilegeException"></exception>
    Public Function Sleep() As Boolean
        Return Suspendc.SuspendSytem(ISuspend.SuspendType.Sleep, False, False)
    End Function

    ''' <summary>Suspends the system by shutting power down. The system enters a sleep state (S1, S2, or S3)</summary>
    ''' <param name="Force">Forse flag</param>
    ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
    ''' <exception cref="SuspendNotSupportedException">Sleep is not supported by the system. Check SleepSupported property</exception>
    ''' <exception cref="ShutdownPrivilegeException"></exception>
    Public Function Sleep(Force As Boolean) As Boolean
        Return Suspendc.SuspendSytem(ISuspend.SuspendType.Sleep, Force, False)
    End Function

    ''' <summary>Suspends the system by shutting power down. The system enters a sleep state (S1, S2, or S3)</summary>
    ''' <param name="Force">Forse flag</param>
    ''' <param name="DisableWakeUpEvent">True: The system disables all wake events / False: Any system wake events remain enabled</param>
    ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
    ''' <exception cref="SuspendNotSupportedException">Sleep is not supported by the system. Check SleepSupported property</exception>
    ''' <exception cref="ShutdownPrivilegeException"></exception>
    Public Function Sleep(Force As Boolean, DisableWakeUpEvent As Boolean) As Boolean Implements ISuspend.Sleep
        Return Suspendc.SuspendSytem(ISuspend.SuspendType.Sleep, Force, DisableWakeUpEvent)
    End Function

#End Region

#Region "Hibernate Functions"
    ''' <summary>Suspends the system by shutting power down. The system enters a hibernation (S4)</summary>
    ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
    ''' <exception cref="SuspendNotSupportedException">Hibernate is not supported by the system. Check HibernateSupported property</exception>
    ''' <exception cref="ShutdownPrivilegeException"></exception>
    Public Function Hibernate() As Boolean
        Return Suspendc.SuspendSytem(ISuspend.SuspendType.Hibernate, False, False)
    End Function

    ''' <summary>Suspends the system by shutting power down. The system enters a hibernation (S4)</summary>
    ''' <param name="Force">Forse flag</param>
    ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
    ''' <exception cref="SuspendNotSupportedException">Hibernate is not supported by the system. Check HibernateSupported property</exception>
    ''' <exception cref="ShutdownPrivilegeException"></exception>
    Public Function Hibernate(Force As Boolean) As Boolean
        Return Suspendc.SuspendSytem(ISuspend.SuspendType.Hibernate, Force, False)
    End Function

    ''' <summary>Suspends the system by shutting power down. The system enters a hibernation (S4)</summary>
    ''' <param name="Force">Forse flag</param>
    ''' <param name="DisableWakeUpEvent">True: The system disables all wake events / False: Any system wake events remain enabled</param>
    ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
    ''' <exception cref="SuspendNotSupportedException">Hibernate is not supported by the system. Check HibernateSupported property</exception>
    ''' <exception cref="ShutdownPrivilegeException"></exception>
    Public Function Hibernate(Force As Boolean, DisableWakeUpEvent As Boolean) As Boolean Implements ISuspend.Hibernate
        Return Suspendc.SuspendSytem(ISuspend.SuspendType.Hibernate, Force, DisableWakeUpEvent)
    End Function

#End Region

#Region "Lock Function"
    ''' <summary>
    ''' Locks the workstation's display. Locking a workstation protects it from unauthorized use.
    ''' The function executes asynchronously
    ''' </summary>
    ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
    Public Function Lock() As Boolean
        Dim Lockc As LockCore = New LockCore()
        Return Lockc.Lock()
    End Function

#End Region

    Private Function ConvertPlannedToOptionalReason(Planned As Boolean) As ShutdownReasonOptional
        If Planned Then
            Return ShutdownReasonOptional.Planned
        Else
            Return ShutdownReasonOptional.Unplanned
        End If
    End Function




End Class
