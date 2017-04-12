Namespace Core

    Public Class LockCore
        'Windows API Function Decleration
        Private Declare Auto Function LockWorkStation Lib "user32" () As Boolean

        ''' <summary>
        ''' Locks the workstation's display. Locking a workstation protects it from unauthorized use.
        ''' The function executes asynchronously
        ''' </summary>
        ''' <returns>If the system accepts, the function returns True. Otherwise the function returns False</returns>
        Public Function Lock() As Boolean
            Dim RetVal As Boolean
            RetVal = LockWorkStation()
            Return RetVal
        End Function

    End Class

End Namespace