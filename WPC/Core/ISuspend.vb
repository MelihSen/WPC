Imports WPC.Core

Public Interface ISuspend

    Enum SuspendType
        Hibernate
        Sleep
    End Enum

    ''' <summary>Determines whether the computer supports hibernation </summary>
    ''' <returns>If the computer supports hibernation (power state S4) and the file Hiberfil.sys is present on the system, the property returns True. Otherwise, the property returns False</returns>
    ReadOnly Property HibernateSupported As Boolean

    ''' <summary>Determines whether the computer supports the sleep states </summary>
    ''' <returns>If the computer supports the sleep states (S1, S2, and S3), the propery returns True. Otherwise, the property returns False</returns>
    ReadOnly Property SleepSupported As Boolean

    ''' <summary>Suspends the system by shutting power down. The system enters a sleep state (S1, S2, or S3)</summary>
    ''' <param name="Force">Forse flag</param>
    ''' <param name="DisableWakeUpEvent">True: The system disables all wake events / False: Any system wake events remain enabled</param>
    ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
    ''' <exception cref="NotSupportedException">Sleep is not supported by the system. Check SleepSupported property</exception>
    ''' <exception cref="ShutdownPrivilegeException"></exception>
    Function Sleep(Force As Boolean, DisableWakeUpEvent As Boolean) As Boolean

    ''' <summary>Suspends the system by shutting power down. The system enters a hibernation (S4)</summary>
    ''' <param name="Force">Forse flag</param>
    ''' <param name="DisableWakeUpEvent">True: The system disables all wake events / False: Any system wake events remain enabled</param>
    ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
    ''' <exception cref="NotSupportedException">Hibernate is not supported by the system. Check HibernateSupported property</exception>
    ''' <exception cref="ShutdownPrivilegeException"></exception>
    Function Hibernate(Force As Boolean, DisableWakeUpEvent As Boolean) As Boolean

End Interface
