Imports System
Imports DevExpress.XtraEditors
Imports DevExpress.XtraDialogs.FileExplorerExtensions
Imports DevExpress.Data
Imports CCI_Engineering_Templates
Imports DevExpress.Utils.Svg
Imports DevExpress.Utils.Drawing
Imports System.IO

Namespace UnitTesting
    Partial Public Class LogViewer
        Inherits XtraUserControl

        Public Sub New()
            InitializeComponent()

            booInfo.Checked = My.Settings.booInfo
            booDebug.Checked = My.Settings.booDebug
            booWarning.Checked = My.Settings.booWarning
            booError.Checked = My.Settings.booError
            booEvent.Checked = My.Settings.booEvent
        End Sub

        Private Sub LogActivityFilter(sender As Object, e As EventArgs) Handles booInfo.CheckedChanged, booEvent.CheckedChanged, booDebug.CheckedChanged, booWarning.CheckedChanged, booError.CheckedChanged
            If isopening Then Exit Sub

            ButtonclickToggle(Me.Cursor)
            Select Case sender.name.ToString
                Case "booInfo"
                    My.Settings.booInfo = sender.Checked
                Case "booEvent"
                    My.Settings.booEvent = sender.Checked
                Case "booDebug"
                    My.Settings.booDebug = sender.Checked
                Case "booWarning"
                    My.Settings.booWarning = sender.Checked
                Case "booError"
                    My.Settings.booError = sender.Checked
            End Select

            My.Settings.Save()

            ReloadActivityLog()
            ButtonclickToggle(Me.Cursor)
        End Sub

        Public Sub Clear()
            booInfo.Text = "INFO"
            booDebug.Text = "DEBUG"
            booEvent.Text = "EVENT"
            booError.Text = "ERROR"
            booWarning.Text = "WARNGING"

            GridView2.Columns.Clear()
            gcTestLog.DataSource = Nothing
        End Sub


        Private Sub GridView2_CustomDrawRowIndicator(sender As Object, e As DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs) Handles GridView2.CustomDrawRowIndicator
            If isopening Then Exit Sub

            If GridView2.Columns.Count = 0 Then Exit Sub

            Dim svgimg As SvgImage
            Dim svgStr As String
            Dim mycolor As Color = Nothing
            Dim img As Image
            Dim type As String
            Dim rect As Rectangle = New Rectangle(e.Bounds.Location, New Size(10.0!, 10.0!))

            Try
                Dim TypeCellValue As Object = GridView2.GetRowCellValue(e.RowHandle, GridView2.Columns("Type"))
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
        Public Sub ReloadActivityLog(Optional ByVal logPath As String = "")
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
            logDt.Columns.Add("Iteration")
            logDt.Columns.Add("Description")

            If logPath = "" Then logPath = frmMain.TestLogActivityPath

            Using sr As New StreamReader(logPath)
                While Not sr.EndOfStream
                    tempLine = sr.ReadLine.ToString
                    lineVars = tempLine.Split("|")

                    Select Case lineVars(2).ToString.ToLower.Replace(" ", "")
                        Case "info"
                            countInfo += 1
                            If My.Settings.booInfo Then AddToLog(lineVars, tempLine, logStr, logDt)
                        Case "debug"
                            countDebug += 1
                            If My.Settings.booDebug Then AddToLog(lineVars, tempLine, logStr, logDt)
                        Case "warning"
                            countWarning += 1
                            If My.Settings.booWarning Then AddToLog(lineVars, tempLine, logStr, logDt)
                        Case "error"
                            countError += 1
                            If My.Settings.booError Then AddToLog(lineVars, tempLine, logStr, logDt)
                        Case "eventbegin", "start", "finish", "eventend", "begin", "end", "process"
                            countEvent += 1
                            If My.Settings.booEvent Then AddToLog(lineVars, tempLine, logStr, logDt)
                        Case Else
                            logStr += tempLine
                    End Select
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
                Me.GridView2.RefreshData()
                Me.GridView2.MoveLastVisible()
                Me.GridView2.BestFitColumns()
                Me.GridView2.Columns(0).Width = 75
                Me.GridView2.Columns(1).Width = 75
                Me.GridView2.Columns(2).Width = 75
                Me.GridView2.Columns(3).Width = 75
                Me.GridView2.Columns(4).Width = 75
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
                logdt.Rows.Add(lineVars(2).Trim, mdate.Replace(" ", ""), mtime.Replace(" ", ""), lineVars(1).Replace(" ", ""), "1", lineVars(3))
            End If

        End Sub
    End Class
End Namespace
