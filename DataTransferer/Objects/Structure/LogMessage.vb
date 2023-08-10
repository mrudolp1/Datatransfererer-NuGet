Public Class LogMessage
    Public Enum MessageType
        DEBUG
        WARNING
        INFO
        [ERROR]
        PROCESS
    End Enum

    Public Property Type As MessageType
    Public Property Description As String

    Public Sub New()
        'Default
    End Sub
    Public Sub New(type As MessageType, description As String)
        Me.Type = type
        Me.Description = description
    End Sub

    Public Sub New(type As String, desctiption As String)
        Me.Type = DirectCast([Enum].Parse(GetType(MessageType), type), MessageType)
        Me.Description = desctiption
    End Sub

    Public Shared Function CreateDebug(ByVal msg As String) As LogMessage
        Return (New LogMessage(LogMessage.MessageType.DEBUG, msg))
    End Function

    Public Shared Function CreateInfo(ByVal msg As String) As LogMessage
        Return (New LogMessage(LogMessage.MessageType.INFO, msg))
    End Function

    Public Shared Function CreateWarning(ByVal msg As String) As LogMessage
        Return (New LogMessage(LogMessage.MessageType.WARNING, msg))
    End Function

    Public Shared Function CreateError(ByVal msg As String) As LogMessage
        Return (New LogMessage(LogMessage.MessageType.ERROR, msg))
    End Function

    Public Shared Function CreateProcess(ByVal msg As String) As LogMessage
        Return (New LogMessage(LogMessage.MessageType.PROCESS, msg))
    End Function
End Class

Public Class MessageLoggedEventArgs
    Inherits EventArgs

    Public Property Message As LogMessage

    Public Sub New()
        ''Default
    End Sub

    Public Sub New(logMessage As LogMessage)
        Me.Message = logMessage
    End Sub
End Class
