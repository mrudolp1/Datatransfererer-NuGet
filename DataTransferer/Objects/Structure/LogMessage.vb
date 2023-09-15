Imports System.Text.RegularExpressions
Imports DevExpress.Xpo.Helpers.AssociatedCollectionCriteriaHelper

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
    Public Property User As String
    Public Property TimeStamp As String

    Public Sub New()
        'Default
    End Sub
    Public Sub New(type As MessageType, description As String, Optional user As String = Nothing, Optional timeStamp As String = Nothing)
        Me.Type = type
        Me.Description = description
        Me.User = If(user, Environment.UserName)
        Me.TimeStamp = If(timeStamp, DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss tt"))
    End Sub

    Public Sub New(type As String, desctiption As String, Optional user As String = Nothing, Optional timeStamp As String = Nothing)
        Me.Type = DirectCast([Enum].Parse(GetType(MessageType), type), MessageType)
        Me.Description = desctiption
        Me.User = If(user, Environment.UserName)
        Me.TimeStamp = If(timeStamp, DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss tt"))
    End Sub

    Private Function RemoveLineBreaks(ByVal description As String)
        Dim pattern As String = "(\r\n|\r|\n|vbCrLf|vbCr|vbLf)"
        ' Use Regex.Replace to remove line breaks and replace them with a space.
        Dim result As String = Regex.Replace(description, pattern, " ")

        Return result
    End Function

    Public Shared Function CreateDebug(ByVal msg As String, Optional user As String = Nothing, Optional timeStamp As String = Nothing) As LogMessage
        Return (New LogMessage(LogMessage.MessageType.DEBUG, msg, user, timeStamp))
    End Function

    Public Shared Function CreateInfo(ByVal msg As String, Optional user As String = Nothing, Optional timeStamp As String = Nothing) As LogMessage
        Return (New LogMessage(LogMessage.MessageType.INFO, msg, user, timeStamp))
    End Function

    Public Shared Function CreateWarning(ByVal msg As String, Optional user As String = Nothing, Optional timeStamp As String = Nothing) As LogMessage
        Return (New LogMessage(LogMessage.MessageType.WARNING, msg, user, timeStamp))
    End Function

    Public Shared Function CreateError(ByVal msg As String, Optional user As String = Nothing, Optional timeStamp As String = Nothing) As LogMessage
        Return (New LogMessage(LogMessage.MessageType.ERROR, msg, user, timeStamp))
    End Function

    Public Shared Function CreateProcess(ByVal msg As String, Optional user As String = Nothing, Optional timeStamp As String = Nothing) As LogMessage
        Return (New LogMessage(LogMessage.MessageType.PROCESS, msg, user, timeStamp))
    End Function

    Public Overrides Function ToString() As String
        Dim dt As String = DateTime.Now.ToString("MM/dd/yyyy")
        ToString = If(Me.TimeStamp.Length < 15, dt & " " & Me.TimeStamp, Me.TimeStamp)
        ToString += String.Format(" | {0}", Me.User)
        ToString += String.Format(" | {0}", Me.Type.ToString())
        ToString += String.Format(" | {0}", Me.Description)
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
