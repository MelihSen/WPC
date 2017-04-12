
Namespace Core

    Friend Class SuspendCore
        Implements ISuspend

#Region "Windows API Function Declerations"
        Private Declare Function IsPwrHibernateAllowed Lib "powrprof.dll" () As Boolean
        Private Declare Function IsPwrSuspendAllowed Lib "powrprof.dll" () As Boolean
        Private Declare Function SetSuspendState Lib "powrprof.dll" (ByVal Hibernate As Boolean, ByVal ForceCritical As Boolean, ByVal DisableWakeEvent As Boolean) As Boolean
#End Region

        ''' <summary>Determines whether the computer supports hibernation </summary>
        ''' <returns>If the computer supports hibernation (power state S4) and the file Hiberfil.sys is present on the system, the property returns True. Otherwise, the property returns False</returns>
        Public ReadOnly Property HibernateSupported As Boolean Implements ISuspend.HibernateSupported
            Get
                Return IsPwrHibernateAllowed()
            End Get
        End Property

        ''' <summary>Determines whether the computer supports the sleep states </summary>
        ''' <returns>If the computer supports the sleep states (S1, S2, and S3), the propery returns True. Otherwise, the property returns False</returns>
        Public ReadOnly Property SleepSupported As Boolean Implements ISuspend.SleepSupported
            Get
                Return IsPwrSuspendAllowed()
            End Get
        End Property


        ''' <summary>Suspends the system by shutting power down. Depending on the Hibernate parameter, the system either enters a suspend (sleep) state or hibernation (S4) </summary>
        ''' <param name="SuspendType">Hibernate or Sleep</param>
        ''' <param name="Force">ForceCritical parameter</param>
        ''' <param name="DisableWakeUpEvent">If this parameter is True, the system disables all wake events. If the parameter is False, any system wake events remain enabled</param>
        ''' <returns>If the function succeeds, the return value is True. If the function fails, the return value is False</returns>
        ''' <exception cref="SuspendNotSupportedException">SuspendType is not supported by the system. Check HibernateSupported or SleepSupported property</exception>
        ''' <exception cref="ShutdownPrivilegeException"></exception>
        Public Function SuspendSytem(SuspendType As ISuspend.SuspendType, Force As Boolean, DisableWakeUpEvent As Boolean) As Boolean
            Dim TokenPriv As New TokenPrivilegeCore
            Dim RetVal As Boolean = False

            Select Case SuspendType
                Case ISuspend.SuspendType.Hibernate
                    If HibernateSupported Then
                        RetVal = TokenPriv.GetShutdownPrivilege(TokenPrivilegeCore.ShutdownPrivilege.CurrentMachine)
                        If RetVal Then
                            RetVal = SetSuspendState(True, Force, DisableWakeUpEvent)
                        End If
                    Else
                        Throw New SuspendNotSupportedException()
                    End If
                Case ISuspend.SuspendType.Sleep
                    If SleepSupported Then
                        RetVal = TokenPriv.GetShutdownPrivilege(TokenPrivilegeCore.ShutdownPrivilege.CurrentMachine)
                        If RetVal Then
                            RetVal = SetSuspendState(False, Force, DisableWakeUpEvent)
                        End If
                    Else
                        Throw New SuspendNotSupportedException()
                    End If
            End Select
            Return RetVal
        End Function

        Private Function Sleep(Force As Boolean, DisableWakeUpEvent As Boolean) As Boolean Implements ISuspend.Sleep
            Return SuspendSytem(ISuspend.SuspendType.Sleep, Force, DisableWakeUpEvent)
        End Function

        Private Function Hibernate(Force As Boolean, DisableWakeUpEvent As Boolean) As Boolean Implements ISuspend.Hibernate
            Return SuspendSytem(ISuspend.SuspendType.Hibernate, Force, DisableWakeUpEvent)
        End Function

    End Class

    ''' <summary>The exception that is thrown when system does not support suspend type (Sleep or Hibernate)</summary>
    Public Class SuspendNotSupportedException
        Inherits Exception

        Public Sub New()
        End Sub

        Public Sub New(Message As String)
            MyBase.New(Message)
        End Sub

        Public Sub New(Message As String, InnerException As Exception)
            MyBase.New(Message, InnerException)
        End Sub

        Protected Sub New(Info As Runtime.Serialization.SerializationInfo, Context As Runtime.Serialization.StreamingContext)
            MyBase.New(Info, Context)
        End Sub
    End Class

End Namespace


