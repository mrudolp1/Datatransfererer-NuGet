Imports System
Imports DevExpress.XtraEditors
Imports DevExpress.Utils.Svg
Imports DevExpress.Utils.Drawing
Imports System.IO
Imports System.Runtime.CompilerServices
Imports CCI_Engineering_Templates

Partial Public Class LogViewer
    Inherits XtraUserControl

    Public Property viewInfo As Boolean
    Public Property viewEvent As Boolean
    Public Property viewDebug As Boolean
    Public Property viewWarning As Boolean
    Public Property viewError As Boolean
    Public Property LogPath As String
    Public Property AdditionalColumnName As String
    Public Property AdditionalColumnDefault As String

    Public Sub New()
        InitializeComponent()

        booInfo.Checked = My.Settings.booInfo
        booDebug.Checked = My.Settings.booDebug
        booWarning.Checked = My.Settings.booWarning
        booError.Checked = My.Settings.booError
        booEvent.Checked = My.Settings.booEvent
    End Sub

    Private Sub LogActivityFilter(sender As Object, e As EventArgs) Handles booInfo.CheckedChanged, booEvent.CheckedChanged, booDebug.CheckedChanged, booWarning.CheckedChanged, booError.CheckedChanged
        Select Case sender.name.ToString
            Case "booInfo"
                viewInfo = sender.Checked
            Case "booEvent"
                viewEvent = sender.Checked
            Case "booDebug"
                viewDebug = sender.Checked
            Case "booWarning"
                viewWarning = sender.Checked
            Case "booError"
                viewError = sender.Checked
        End Select

        ReloadActivityLog()
    End Sub

    Public Sub Clear()
        booInfo.Text = "INFO"
        booDebug.Text = "DEBUG"
        booEvent.Text = "EVENT"
        booError.Text = "ERROR"
        booWarning.Text = "WARNGING"

        gvTestLog.Columns.Clear()
        gcTestLog.DataSource = Nothing
    End Sub


    Private Sub GridView2_CustomDrawRowIndicator(sender As Object, e As DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs) Handles gvTestLog.CustomDrawRowIndicator
        If isOpening Then Exit Sub

        If gvTestLog.Columns.Count = 0 Then Exit Sub

        Dim svgimg As SvgImage
        Dim svgStr As String
        Dim mycolor As Color = Nothing
        Dim img As Image
        Dim type As String
        Dim rect As Rectangle = New Rectangle(e.Bounds.Location, New Size(10.0!, 10.0!))

        Try
            Dim TypeCellValue As Object = gvTestLog.GetRowCellValue(e.RowHandle, gvTestLog.Columns("Type"))
            type = If(TypeCellValue Is Nothing, "", TypeCellValue.ToString.ToLower)
        Catch ex As Exception
            type = ""
        End Try

        Select Case type.Replace(" ", "")
            Case "info"
                svgimg = My.Resources.about
                svgStr = Application.StartupPath & "\Resources\about.svg"
                mycolor = Color.LightBlue
            Case "debug"
                svgimg = My.Resources.charttype_line
                mycolor = Color.LightCoral
            Case "warning"
                svgimg = My.Resources.warning
                mycolor = Color.LightYellow
            Case "error"
                svgimg = My.Resources.highimportance
                mycolor = Color.DarkRed
            Case "eventbegin", "start", "finish", "eventend", "begin", "end", "process"
                svgimg = My.Resources.rightangleconnector
                mycolor = Color.Gray
        End Select


        Dim pnt As New Point(New Size(0.05, 0.05))
        pnt.X = e.Bounds.X
        pnt.Y = e.Bounds.Y
        Dim pal As DevExpress.Utils.Design.ISvgPaletteProvider = SvgPaletteHelper.GetSvgPalette(Me.LookAndFeel.ActiveLookAndFeel, ObjectState.Normal)

        Try
            'e.Cache.DrawSvgImage(svgimg, pnt, pal)
            e.Appearance.BackColor = mycolor
            e.Appearance.FillRectangle(e.Cache, New Rectangle(e.Bounds.X + 2, e.Bounds.Y + 2, e.Bounds.Width - 4, e.Bounds.Y - 4))
            'Dim btm = SvgBitmap.FromFile(Application.StartupPath & "\Resources\about.svg")
        Catch ex As Exception
        Finally
            e.Handled = True
        End Try
    End Sub

    'Load the log and update the checkbuttons with the quantity of the types of messages in the log
    Public Sub ReloadActivityLog(Optional ByVal logToLoad As String = "")
        Dim logStr As String = ""
        Dim tempLine As String
        Dim lineVars As String()
        Dim countInfo As Integer = 0
        Dim countDebug As Integer = 0
        Dim countWarning As Integer = 0
        Dim countError As Integer = 0
        Dim countEvent As Integer = 0
        Dim logDt As New DataTable
        logDt.Columns.Add("Type")
        logDt.Columns.Add("Date")
        logDt.Columns.Add("Time")
        logDt.Columns.Add("User")
        logDt.Columns.Add(Me.AdditionalColumnName)
        logDt.Columns.Add("Description")

        If logToLoad = "" Then
            If Me.LogPath = "" Then
                Exit Sub
            Else
                logToLoad = Me.LogPath
            End If
        End If

        Using sr As New StreamReader(logToLoad)
            While Not sr.EndOfStream
                tempLine = sr.ReadLine.ToString
                lineVars = tempLine.Split("|")
                If tempLine <> "" Then
                    Select Case lineVars(2).ToString.ToLower.Replace(" ", "")
                        Case "info"
                            countInfo += 1
                            If viewInfo Then AddToLog(lineVars, tempLine, logStr, logDt)
                        Case "debug"
                            countDebug += 1
                            If viewDebug Then AddToLog(lineVars, tempLine, logStr, logDt)
                        Case "warning"
                            countWarning += 1
                            If viewWarning Then AddToLog(lineVars, tempLine, logStr, logDt)
                        Case "error"
                            countError += 1
                            If viewError Then AddToLog(lineVars, tempLine, logStr, logDt)
                        Case "eventbegin", "start", "finish", "eventend", "begin", "end", "process"
                            countEvent += 1
                            If viewEvent Then AddToLog(lineVars, tempLine, logStr, logDt)
                        Case Else
                            logStr += tempLine
                    End Select
                End If
            End While

            sr.Close()
        End Using

        With Me
            .booInfo.UpdateLogCount("Info", countInfo, My.Settings.booInfo)
            .booDebug.UpdateLogCount("Debug", countDebug, My.Settings.booDebug)
            .booWarning.UpdateLogCount("Warning", countWarning, My.Settings.booWarning)
            .booError.UpdateLogCount("Error", countError, My.Settings.booError)
            .booEvent.UpdateLogCount("Event", countEvent, My.Settings.booEvent)
        End With

        'rtfactivityLog.Text = ""
        'rtfactivityLog.AppendText(logStr)
        If logStr = "" Then
            Me.gcTestLog.DataSource = Nothing
        Else
            Me.gcTestLog.DataSource = logDt
            Me.gvTestLog.RefreshData()
            Me.gvTestLog.MoveLastVisible()
            Me.gvTestLog.BestFitColumns()
            Me.gvTestLog.Columns(0).Width = 75
            Me.gvTestLog.Columns(1).Width = 75
            Me.gvTestLog.Columns(2).Width = 75
            Me.gvTestLog.Columns(3).Width = 75
            Me.gvTestLog.Columns(4).Width = 75
        End If
        'rtfactivityLog.ScrollToCaret()
    End Sub

    'Datatable must be formatted with specific columns
    '''Type - Date - Time - Username - Description 
    Public Sub AddToLog(ByVal lineVars As String(), ByVal tempLine As String, ByRef logstr As String, ByRef logdt As DataTable)
        logstr += tempLine & vbCrLf
        Dim dt As String() = lineVars(0).Split(" ")
        Dim mdate As String
        Dim mtime As String

        If dt.Count > 2 Then
            mdate = dt(0)
            mtime = dt(1)
        Else
            mdate = ""
            mtime = dt(0)
        End If

        If lineVars.Count > 4 Then
            logdt.Rows.Add(lineVars(2).Trim, mdate.Replace(" ", ""), mtime.Replace(" ", ""), lineVars(1).Replace(" ", ""), lineVars(4).Replace(" ", ""), lineVars(3))
        Else
            logdt.Rows.Add(lineVars(2).Trim, mdate.Replace(" ", ""), mtime.Replace(" ", ""), lineVars(1).Replace(" ", ""), Me.AdditionalColumnName, lineVars(3))
        End If

    End Sub

    Public Sub LogActivity(ByVal msg As String, Optional ByVal reloadLog As Boolean = False)
        ' Get the current date and time
        Dim dt As String = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss tt")
        'Dim splt() As String = dt.Split(" ")
        'dt = splt(1) '& " " & splt(2)

        ' Print the message to the console
        Console.WriteLine(dt & " | " & Environment.UserName & " | " & msg)

        ' Wrap the file operation in a try-catch block to handle exceptions
        Try
            ' If the log file does not exist, establish intro
            If Not File.Exists(Me.LogPath) Then
                File.Create(Me.LogPath).Dispose()
            End If
            ' Use a StreamWriter to write to the log file
            ' The 'True' argument appends to the file if it already exists
            Using sw As New StreamWriter(Me.LogPath, True)
                ' Write the log message to the file
                For Each line In msg.Split(vbCrLf)
                    sw.WriteLine(dt & " | " & Environment.UserName & " | " & line)
                Next
            End Using
            If reloadLog Then
                Me.ReloadActivityLog()
            End If
        Catch ex As Exception
            ' Handle the exception
            Console.WriteLine("Error writing to log file: " & ex.Message)
        End Try
    End Sub
End Class

Public Module LogViewerExtensions
    'Update the count of the items in the check buttons
    '''The button being updated
    '''the type of message (Info, Error, Debug, etc)
    '''Total lines in the log of that type
    '''Whether or not it is checked
    <Extension()>
    Public Sub UpdateLogCount(ByVal chkbtn As CheckButton, ByVal type As String, ByVal total As Integer, ByVal checked As Boolean)
        If checked Then
            chkbtn.Text = total.ToString & " " & type & "(s)"
        Else
            chkbtn.Text = "0 of " & total.ToString & " " & type & "(s)"
        End If
    End Sub

    'Append the maestro log generated to the 
    <Extension()>
    Public Sub AppendMaestroLog(ByVal strc As EDSStructure, ByVal pathToAppend As String, Optional ByVal AdditionalValue As String = "")

        Dim dateTim As String = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss tt")
        Dim splt() As String = dateTim.Split(" ")

        Dim dt As String = DateTime.Now.ToString("MM/dd/yyyy")

        Dim inputs As Char() = {"|"}
        Dim separator As String = "|"

        Using sw As New StreamWriter(pathToAppend, True)
            Using sr As New StreamReader(strc.LogPath)
                While Not sr.EndOfStream
                    Dim myLine As String = sr.ReadLine
                    If myLine.Length > 0 Then
                        Dim vars As String() = myLine.Split(separator)
                        If vars.Count < 3 Then
                            sw.WriteLine(dt & " " & vars(0) & splt(2) & " " & separator & " " & Environment.UserName & " " & separator & "INFO" & " " & separator & vars(1) & " " & separator & " " & AdditionalValue)
                        ElseIf vars.Count = 1 Then
                            sw.WriteLine(dt & " " & separator & " " & Environment.UserName & " " & separator & "DEBUG" & " " & separator & vars(0) & " " & separator & " " & AdditionalValue)
                        Else
                            sw.WriteLine(dt & " " & vars(0) & splt(2) & " " & separator & " " & Environment.UserName & " " & separator & vars(1) & separator & vars(2) & " " & separator & " " & AdditionalValue)
                        End If
                    End If
                End While
                sr.Close()
            End Using
        End Using
    End Sub
End Module