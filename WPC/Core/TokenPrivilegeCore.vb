Imports System.Runtime.InteropServices

Namespace Core

    Public Class TokenPrivilegeCore

#Region "Windows API Function Declerations"
        Private Declare Function OpenProcessToken Lib "advapi32.dll" (ProcessHandle As IntPtr, DesiredAccess As Integer, ByRef TokenHandle As IntPtr) As Boolean
        Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (lpSystemName As String, lpName As String, ByRef lpLuid As LUID) As Boolean
        Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As IntPtr, ByVal DisableAllPrivileges As Boolean, ByRef NewState As TokenPrivileges, ByVal BufferLength As Integer, ByRef PreviousState As TokenPrivileges, ByRef ReturnLength As IntPtr) As Boolean
        Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Integer) As Integer
#End Region

#Region "Windows API - Constants"
        Private Const SE_SHUTDOWN_NAME As String = "SeShutdownPrivilege"
        Private Const SE_REMOTE_SHUTDOWN_NAME As String = "SeRemoteShutdownPrivilege"
        Private Const SE_PRIVILEGE_ENABLED As Integer = &H2

        ' Access Rights Constants (DesiredAccess)
        Private Const READ_CONTROL As Integer = &H20000
        Private Const STANDARD_RIGHTS_REQUIRED As Integer = &HF0000
        Private Const STANDARD_RIGHTS_READ As Integer = READ_CONTROL
        Private Const STANDARD_RIGHTS_WRITE As Integer = READ_CONTROL
        Private Const STANDARD_RIGHTS_EXECUTE As Integer = READ_CONTROL
        Private Const STANDARD_RIGHTS_ALL As Integer = &H1F0000
        Private Const SPECIFIC_RIGHTS_ALL As Integer = &HFFFF

        Private Const TOKEN_ASSIGN_PRIMARY As Integer = &H1
        Private Const TOKEN_DUPLICATE As Integer = &H2
        Private Const TOKEN_IMPERSONATE As Integer = &H4
        Private Const TOKEN_QUERY As Integer = &H8
        Private Const TOKEN_QUERY_SOURCE As Integer = &H10
        Private Const TOKEN_ADJUST_PRIVILEGES As Integer = &H20
        Private Const TOKEN_ADJUST_GROUPS As Integer = &H40
        Private Const TOKEN_ADJUST_DEFAULT As Integer = &H80
        Private Const TOKEN_ADJUST_SESSIONID As Integer = &H100
        Private Const TOKEN_ALL_ACCESS As Integer = STANDARD_RIGHTS_REQUIRED Or TOKEN_ASSIGN_PRIMARY Or TOKEN_DUPLICATE Or TOKEN_IMPERSONATE Or TOKEN_QUERY Or TOKEN_QUERY_SOURCE Or TOKEN_ADJUST_PRIVILEGES Or TOKEN_ADJUST_GROUPS Or TOKEN_ADJUST_DEFAULT
        Private Const TOKEN_READ As Integer = STANDARD_RIGHTS_READ Or TOKEN_QUERY
        Private Const TOKEN_WRITE As Integer = STANDARD_RIGHTS_WRITE Or TOKEN_ADJUST_PRIVILEGES Or TOKEN_ADJUST_GROUPS Or TOKEN_ADJUST_DEFAULT
        Private Const TOKEN_EXECUTE As Integer = STANDARD_RIGHTS_EXECUTE
#End Region

        ''' <summary> Locally Unique Identifier Structure (LUID)</summary>
        Private Structure LUID
            Dim LowPart As Integer
            Dim HighPart As Integer
        End Structure

        ''' <summary>Contains information about a set of privileges for an access token </summary>
        Private Structure TokenPrivileges
            Dim Count As Integer
            Dim Privileges As LUID
            Dim Attributes As Integer
        End Structure

        ''' <summary>ShutdownPrivilege type</summary>
        Public Enum ShutdownPrivilege
            CurrentMachine
            RemoteMachine
        End Enum

        ''' <summary>Enables the SE_SHUTDOWN_NAME privilege</summary>
        ''' <param name="ShutdownPrivilegeType">Current Machine or Remote Machine</param>
        ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
        ''' <exception cref="ShutdownPrivilegeException"></exception>
        Public Function GetShutdownPrivilege(ShutdownPrivilegeType As ShutdownPrivilege) As Boolean

            Dim ProcessHandle As IntPtr
            Dim TokenHandle As IntPtr
            Dim ReturnLength As IntPtr

            Dim Luid As LUID
            Dim TokenPriv As TokenPrivileges
            Dim PreviousTokenPriv As TokenPrivileges
            Dim RetVal As Boolean = False
            Dim TokenLength As Integer = Marshal.SizeOf(TokenPriv)  'Len(TokenPriv)

            Try
                ProcessHandle = Process.GetCurrentProcess.Handle
                RetVal = OpenProcessToken(ProcessHandle, TOKEN_ALL_ACCESS, TokenHandle)

                If ShutdownPrivilegeType = ShutdownPrivilege.CurrentMachine Then
                    RetVal = LookupPrivilegeValue(vbNullString, SE_SHUTDOWN_NAME, Luid)
                Else
                    RetVal = LookupPrivilegeValue(vbNullString, SE_REMOTE_SHUTDOWN_NAME, Luid)
                End If

                With TokenPriv
                    .Count = 1
                    .Privileges = Luid
                    .Attributes = SE_PRIVILEGE_ENABLED
                End With

                RetVal = AdjustTokenPrivileges(TokenHandle, False, TokenPriv, TokenLength, PreviousTokenPriv, ReturnLength)

            Catch Ex As Exception
                Throw New ShutdownPrivilegeException("Shutdown privilege error", Ex)
            Finally
                Call CloseHandle(TokenHandle.ToInt32())
                Call CloseHandle(ProcessHandle.ToInt32())
            End Try
            Return RetVal
        End Function

    End Class

    ''' <summary>The exception that is thrown when system does not allow to get Shutdown Privilege (SE_SHUTDOWN_NAME)</summary>
    Public Class ShutdownPrivilegeException
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

