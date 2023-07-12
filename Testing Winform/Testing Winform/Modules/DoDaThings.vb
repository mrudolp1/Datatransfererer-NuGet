
Imports System.IO
Imports DevExpress.XtraBars.ToastNotifications
Imports System.Xml
Imports System.Text
Imports CCI_Engineering_Templates
Imports System.Runtime.CompilerServices
Imports System.Runtime.Serialization.Json
Imports System.Text.RegularExpressions
Imports DevExpress.XtraEditors

Namespace UnitTesting

    '------------------------------------------
    '------------------------------------------
    'Module Name: DoDaThings
    'Purpose: Designed for anything that returns a value or performs a specific action
    '------------------------------------------
    '------------------------------------------

    Module DoDaThings
        Public toaster As New ToastNotificationsManager

        <DebuggerStepThrough()>
        Public Function FetchUserData(ByVal data As String) As String

            Dim objAd As Object = CreateObject("ADSystemInfo")
            Dim objuser As Object = GetObject("LDAP://" & objAd.UserName)
            Dim UserData As String = ""

            Select Case data
                Case "FirstName" : UserData = objuser.FirstName
                Case "LastName" : UserData = objuser.LastName
                Case "FullName" : UserData = objuser.FullName
                Case "Description" : UserData = objuser.Description
                Case "physicalDeliveryOfficeName" : UserData = objuser.physicalDeliveryOfficeName
                Case "telephoneNumber" : UserData = objuser.telephoneNumber
                Case "EmailAddress" : UserData = objuser.EmailAddress
                Case "streetAddress" : UserData = objuser.streetAddress
                Case "city" : UserData = objuser.l
                Case "state" : UserData = objuser.st
                Case "zip" : UserData = objuser.postalCode
                Case "UserName" : UserData = objuser.sAMAccountName
                Case "Mobile" : UserData = objuser.Mobile
                Case "ipPhone" : UserData = objuser.ipPhone
                Case "Title" : UserData = objuser.Title
                Case "Department" : UserData = objuser.department
                Case "Company" : UserData = objuser.company
            End Select

            objAd = Nothing
            objuser = Nothing

            Return UserData
        End Function

        Public Sub KillRoboCops()
            Dim proc = Process.GetProcessesByName("RoboCopy")
            For i As Integer = 0 To proc.Count - 1
                Try
                    proc(i).Kill()
                Catch
                End Try
            Next i
        End Sub

#Region "Toaster"
        Public Sub Toaster_Activated(sender As Object, e As DevExpress.XtraBars.ToastNotifications.ToastNotificationEventArgs)
            'Dim values As String() = e.NotificationID.ToString.Split("|")
            'XtraMessageBox.Show(values(1), values(0))
        End Sub

        Public Sub Toaster_UpdateToastContent(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ToastNotifications.UpdateToastContentEventArgs)
            Dim content As XmlDocument = e.ToastContent
            Dim toastElement As XmlElement = content.GetElementsByTagName("toast").OfType(Of XmlElement)().FirstOrDefault()
            Dim actions As XmlElement = content.CreateElement("actions")
            Dim action As XmlElement = content.CreateElement("action")

            toastElement.AppendChild(actions)
            actions.AppendChild(action)
            action.SetAttribute("content", "Show details")
            action.SetAttribute("arguments", "viewdetails")
        End Sub

        <DebuggerStepThrough()>
        Public Sub sendToast(ByVal ToastMessage As String, ToastTitle As String)
            Dim bread As New ToastNotification 'Take out your bread

            'Adjust the settings
            bread.Header = ToastTitle
            bread.Body = ToastMessage
            bread.Body2 = ""
            bread.Sound = ToastNotificationSound.Default
            bread.Template = ToastNotificationTemplate.ImageAndText03
            bread.Duration = ToastNotificationDuration.Long
            bread.ID = bread.Header & "|" & ToastMessage
            Try
                bread.AttributionText = "(" & System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString & ")"
            Catch ex As Exception
                bread.AttributionText = "(" & betaVersion & ")"
            End Try

            bread.Template = ToastNotificationTemplate.ImageAndText03
            'bread.Image = My.Resources.Drone_Reconciliation_Icon
            'bread.AppLogoImage = My.Resources.Drone_Reconciliation_Icon

            toaster.Notifications.Add(bread) 'Put bread in toaster
            toaster.ShowNotification(bread.ID) 'Run the toaster for this bread
            bread.ID = "toast" 'Bread turns into toast
            toaster.Notifications.Remove(bread) 'Remove the toast from the toaster
        End Sub
#End Region

        <DebuggerStepThrough()>
        Sub dtClearer(ByVal sqlsrc As String)
            Try
                If ds.Tables.Contains(sqlsrc) Then
                    ds.Tables(sqlsrc).Clear()
                    ds.Tables(sqlsrc).Columns.Clear()
                End If
            Catch
            End Try
        End Sub

        <DebuggerStepThrough()>
        Public Sub releaseObject(ByVal obj As Object)
            Try
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
                obj = Nothing
            Catch ex As Exception
                obj = Nothing
            Finally
                GC.Collect()
            End Try
        End Sub

        <DebuggerStepThrough()>
        Public Sub KillProcess(ByVal processName As String)

            Try
                'Dim MSExcelControl() As Process
                Dim iID As Integer
                Dim lastOpen As DateTime
                Dim obj1(10) As Process
                obj1 = Process.GetProcessesByName(processName)
                lastOpen = obj1(0).StartTime
                For Each p As Process In obj1
                    If lastOpen = p.StartTime Then
                        iID = p.Id
                        Exit For
                    End If
                Next

                For Each p As Process In obj1
                    If p.Id = iID Then
                        p.Kill()
                        Exit For
                    End If
                Next

            Catch ex As Exception

            End Try
        End Sub

        Function QueryBuilderFromFile(ByVal filetoread As String) As String
            Dim sqlReader As New StreamReader(filetoread)
            Dim temp As String = ""

            Do While sqlReader.Peek <> -1
                temp += sqlReader.ReadLine.Trim + vbNewLine
            Loop

            sqlReader.Dispose()
            sqlReader.Close()

            Return temp
        End Function

        '<DebuggerStepThrough()>
        Public Function token(s As String) As String
            Dim m As String = ""
            For x As Integer = 0 To 1000
                Try
                    m = m & Chr(s.Split(":")(x) / Chr(51).ToString)
                Catch
                    Exit For
                End Try
            Next
            Return m
        End Function

        <DebuggerStepThrough()>
        Function TableContainsRecords(ByVal myTable As DataTable) As Boolean
            If myTable.Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
        End Function

        Public Sub RemoveDataTable(ByVal srcName As String)
            Try
                dtClearer(srcName)
                ds.Tables(srcName).Clear()
                ds.Tables(srcName).Columns.Clear()
                ds.Tables.Remove(srcName)
            Catch ex As Exception
            End Try
        End Sub

        Public Sub ButtonClicksStart(ByRef cur As Cursor)
            isLoading = True
            cur = Cursors.WaitCursor
        End Sub

        Public Sub ButtonClickEnd(ByRef cur As Cursor)
            isLoading = False
            cur = Cursors.Default
        End Sub

        'Public Function ReadSignature() As String
        '    Dim appDataDir As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Signatures"
        '    Dim signature As String = String.Empty
        '    Dim diInfo As DirectoryInfo = New DirectoryInfo(appDataDir)

        '    If diInfo.Exists Then
        '        Dim fiSignature As FileInfo() = diInfo.GetFiles("*.htm")

        '        If fiSignature.Length > 0 Then
        '            For i As Integer = 0 To fiSignature.Count - 1
        '                If fiSignature(i).Name.Replace(fiSignature(0).Extension, String.Empty) = My.Settings.mySignature Then
        '                    Dim sr As StreamReader = New StreamReader(fiSignature(i).FullName, System.Text.Encoding.[Default])
        '                    signature = sr.ReadToEnd()

        '                    If Not String.IsNullOrEmpty(signature) Then
        '                        Dim fileName As String = fiSignature(i).Name.Replace(fiSignature(0).Extension, String.Empty)
        '                        signature = signature.Replace(fileName & "_files/", appDataDir & "/" & fileName & "_files/")
        '                    End If
        '                End If

        '            Next i
        '        End If
        '    End If

        '    Return signature
        'End Function

        Function AddBusinessDays(startDate As Date, numberOfDays As Integer) As Date
            Dim newDate As Date = startDate
            While numberOfDays > 0
                newDate = newDate.AddDays(1)

                If newDate.DayOfWeek() > 0 AndAlso newDate.DayOfWeek() < 6 Then '1-5 is Mon-Fri
                    numberOfDays -= 1
                End If

            End While
            Return newDate
        End Function

        'IEM 11/4/2021 Sometimes you need a little something
        Public Function IsSomething(ByVal sender As Object) As Boolean
            If Not IsNothing(sender) Then Return True
            Return False
        End Function


        Public Function FileBrowse(ByVal title As String, ByVal filter As String) As String

            Dim ofd As New OpenFileDialog
            ofd.Multiselect = False
            ofd.Title = title
            ofd.Filter = filter

            If (ofd.ShowDialog() = DialogResult.OK) Then
                Return ofd.FileName
            End If

            Return Nothing
        End Function

        Public Function FolderBrowse(ByVal title As String, ByVal filter As String) As String

            Dim ofd As New OpenFileDialog
            ofd.Multiselect = False
            ofd.Title = title
            ofd.Filter = filter
            ofd.CheckPathExists = True
            ofd.ShowReadOnly = False
            ofd.CheckFileExists = False
            ofd.ValidateNames = False
            ofd.ReadOnlyChecked = True
            ofd.FileName = "Folder Selection."

            If (ofd.ShowDialog() = DialogResult.OK) Then
                Return ofd.FileName.Replace("Folder Selection", "")
            End If

            Return Nothing
        End Function

        'Returns a date/time 
        Public Function GetDateTimeFilePath() As String
            Dim d As Integer = Now.Day
            Dim mo As Integer = Now.Month
            Dim y As Integer = Now.Year
            Dim h As Integer = Now.Hour
            Dim mi As Integer = Now.Minute
            Dim s As Integer = Now.Second

            Dim dstr As String
            Dim mostr As String
            Dim ystr As String = y
            Dim hstr As String
            Dim mistr As String
            Dim sstr As String

            If h < 10 Then hstr = "0" & h Else hstr = h
            If mi < 10 Then mistr = "0" & mi Else mistr = mi
            If s < 10 Then sstr = "0" & s Else sstr = s
            If d < 10 Then dstr = "0" & d Else dstr = d
            If mo < 10 Then mostr = "0" & mo Else mostr = mo

            Return mostr & dstr & ystr & "_" & hstr & mistr & sstr
        End Function

        'This was taken from logic used in the CCI SQL Manager but has been adjusted to use a datatable instead of a datagrid. 
        'If you are trying to output something that a user is editing. The data will need to be converted to a datatable to utilize this
        'This could probably be updated to work similar to the thing Ken Linck wrote that accepts any type of object. Instead of ouputting HTML calls we could output CSV.
        Public Sub DatatableToCSVOld(ByVal dtDataTable As DataTable, ByVal strFilePath As String)
            Dim sw As StreamWriter = New StreamWriter(strFilePath, False)

            For i As Integer = 0 To dtDataTable.Columns.Count - 1
                sw.Write(dtDataTable.Columns(i))

                If i < dtDataTable.Columns.Count - 1 Then
                    sw.Write(",")
                End If
            Next

            sw.Write(sw.NewLine)

            For Each dr As DataRow In dtDataTable.Rows

                For i As Integer = 0 To dtDataTable.Columns.Count - 1

                    If Not Convert.IsDBNull(dr(i)) Then
                        Dim value As String = dr(i).ToString()

                        If value.Contains(","c) Then
                            value = String.Format("""{0}""", value)
                            sw.Write(value)
                        Else
                            sw.Write(dr(i).ToString())
                        End If
                    End If

                    If i < dtDataTable.Columns.Count - 1 Then
                        sw.Write(",")
                    End If
                Next

                sw.Write(sw.NewLine)
            Next

            sw.Close()
        End Sub

        'This function requires data that is very specific. 
        'The columns in the datatable will create a concatenated string in the order of the columns. 
        'You will need to specify your column order in your query or make sure it is exactly the same as your SQL table. 
        Public Function DataTableToSQL(ByVal dtDataTable As DataTable, Optional ByVal firstVariable As String = "") As String
            Dim tempsql As String

            'Blank data tables won't return anything. Check to make sure there are rows in the datatable
            If dtDataTable.Rows.Count = 0 Then
                Return String.Empty
            End If

            For j As Integer = 0 To dtDataTable.Rows.Count - 1
                Dim dr As DataRow = dtDataTable.Rows(j)
                tempsql += " (" & firstVariable
                For i As Integer = 0 To dtDataTable.Columns.Count - 1

                    If Not Convert.IsDBNull(dr(i)) Then
                        Dim value As String = dr(i).ToString()
                        value = String.Format("'{0}'", value.Replace("'", "''"))
                        tempsql += value
                    Else
                        tempsql += "NULL"
                    End If

                    If i < dtDataTable.Columns.Count - 1 Then
                        tempsql += ","
                    End If
                Next
                tempsql += ") "
                If j < dtDataTable.Rows.Count - 1 Then
                    tempsql += ","
                End If
                tempsql += vbNewLine
            Next

            'I'd like to build this logic into the loop above but couldn't quickly think to add NULL when this scenario happens. 
            'Unsure if it is actually happening. Just haven't had time to actually review. 
            Return tempsql.Replace(",)", ",NULL)")
        End Function
    End Module

    Public Module GeneralHelpers 'salute

        'serialize any object to a json
        '''Object being passed in
        '''location to save the file path
        Public Function ObjectToJson(Of T)(ByVal obj As Object, ByVal jsonPath As String) As Tuple(Of Boolean, String)
            Dim objJson As String

            Try
                objJson = ToJsonString(Of T)(CType(obj, T))
                If objJson.Contains("ERROR SERIALIZING") Then
                    Return New Tuple(Of Boolean, String)(False, objJson)
                    Exit Function
                End If

                Using sw As New StreamWriter(jsonPath)
                    sw.Write(objJson)
                    sw.Close()
                End Using
                Return New Tuple(Of Boolean, String)(True, "")
            Catch ex As Exception
                objJson = Nothing
                Return New Tuple(Of Boolean, String)(False, ex.Message)
            End Try
        End Function

        'Determine if the maestro conductor ran successfully
        Public Function DidConductProperly(ByVal logpath As String) As Boolean
            Dim isFailure As Boolean = True
            Using maeSr As New StreamReader(logpath)
                'if an error exists then it did not conduct properly.
                If maeSr.ReadToEnd.Contains("ERROR") Then
                    isFailure = False
                End If
                maeSr.Close()
            End Using

            Return isFailure
        End Function

        'Archive a directory
        Public Sub DoArchiving(ByVal folder As String)
            Dim arch As DirectoryInfo = Directory.CreateDirectory(folder & "\Archive " & Now.ToString("MM/dd/yyyy HH:mm:ss tt").ToDirectoryString)
            arch.ArchiveFiles(folder)

            frmMain.LogActivity("DEBUG | Folder archived: " & folder)
        End Sub

        'Determine if a file is open
        Public Function FileIsOpen(ByVal file As FileInfo) As Boolean
            Dim stream As FileStream = Nothing
            Try
                stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None)
                stream.Close()
                Return False
            Catch ex As Exception
                Return True
            End Try
        End Function

        'This was taken from logic used in the CCI SQL Manager but has been adjusted to use a datatable instead of a datagrid. 
        'If you are trying to output something that a user is editing. The data will need to be converted to a datatable to utilize this
        'This could probably be updated to work similar to the thing Ken Linck wrote that accepts any type of object. Instead of ouputting HTML calls we could output CSV.
        Public Sub DatatableToCSV(ByVal dtDataTable As DataTable, ByVal strFilePath As String)
            Dim counter As Integer = 1
RetryFileOpenCheck:
            If IO.File.Exists(strFilePath) Then
                If FileIsOpen(New FileInfo(strFilePath)) Then
                    MsgBox(strFilePath & " Is currently open. " & vbCrLf & vbCrLf & "Please close the file To Continue.", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "File Is In use")

                    counter += 1
                    If counter > 2 Then
                        MsgBox("It seems the file Is still open." & vbCrLf & vbCrLf & "Data was Not saved To CSV.", vbInformation)
                        Exit Sub
                    End If
                    GoTo RetryFileOpenCheck
                End If
            End If

            Using sw As StreamWriter = New StreamWriter(strFilePath, False)
                For i As Integer = 0 To dtDataTable.Columns.Count - 1
                    sw.Write(dtDataTable.Columns(i))

                    If i < dtDataTable.Columns.Count - 1 Then
                        sw.Write(",")
                    End If
                Next

                sw.Write(sw.NewLine)

                For Each dr As DataRow In dtDataTable.Rows

                    For i As Integer = 0 To dtDataTable.Columns.Count - 1

                        If Not Convert.IsDBNull(dr(i)) Then
                            Dim value As String = dr(i).ToString()

                            If value.Contains(","c) Then
                                value = String.Format("""{0}""", value)
                                sw.Write(value)
                            Else
                                sw.Write(dr(i).ToString())
                            End If
                        End If

                        If i < dtDataTable.Columns.Count - 1 Then
                            sw.Write(",")
                        End If
                    Next

                    sw.Write(sw.NewLine)
                Next

                sw.Close()
            End Using
        End Sub

        'Convert a CSV file to a databale
        'Uses an OLEDBAdpater to SELECT * FROM csv file
        'None string columns load in with incorrect column headers
        'This is extremely similar to how we use the SQL adapter for the SQL loader and Sender
        'Public Function CSVtoDatatable(ByVal info As FileInfo, Optional ByVal hasHeaders As Boolean = True) As DataTable
        '    Dim dssample As New DataSet
        '    Dim folder = info.FullName.Replace(info.Name, "")
        '    Dim CnStr = "Provider= Microsoft.Jet.OLEDB.4.0;Data Source=" & folder & ";Extended Properties=""text;HDR=No;FMT=Delimited"";"

        '    Using Adp As New OleDbDataAdapter("Select * from [" & info.Name & "]", CnStr)

        '        Try
        '            Adp.Fill(dssample)
        '        Catch
        '        End Try
        '    End Using

        '    If hasHeaders Then
        '        For Each dc As DataColumn In dssample.Tables(0).Columns
        '            'If the data is not saved as a string then it will not recognize the header column as it doesn't not assume headers in the SQL query
        '            'I didn't have time to create custom queries. 
        '            'This just works for any selected csv file
        '            Try
        '                dc.ColumnName = dssample.Tables(0).Rows(0).Item(dc)
        '            Catch
        '            End Try
        '        Next

        '        dssample.Tables(0).Rows.Remove(dssample.Tables(0).Rows(0))
        '    End If

        '    If dssample.Tables.Count > 0 Then
        '        'Only 1 table should have been output but it returns that table
        '        Return dssample.Tables(0)
        '    End If
        'End Function

        Public Function CSVtoDatatable(ByVal info As FileInfo, Optional ByVal hasheaders As Boolean = True) As DataTable

            Dim dt As DataTable = New DataTable()
            Dim row As DataRow
            Dim headersAdded As Boolean = False

            Using SR As StreamReader = New StreamReader(info.FullName)
                If hasheaders Then
                    Dim line As String = SR.ReadLine()
                    Dim strArray As String() = line.Split(","c)
                    For Each s As String In strArray
                        dt.Columns.Add(s)
                        headersAdded = True
                    Next
                End If

                Do
                    Dim line As String
                    line = SR.ReadLine
                    If Not line = String.Empty Then
                        If Not headersAdded Then
                            Dim strArray As String() = line.Split(","c)
                            Dim counter As Integer = 1
                            For Each s As String In strArray
                                dt.Columns.Add("F" & counter)
                                headersAdded = True
                            Next
                        End If

                        row = dt.NewRow()
                        row.ItemArray = line.Split(","c)
                        dt.Rows.Add(row)
                    Else
                        Exit Do
                    End If
                Loop
                SR.Close()
            End Using

            Return dt
        End Function


        'Toggles the cursor between default and waiting
        '''Placed at the beginning and end of form events like button clicks or checkbox changes
        '''Should also be placed anywhere you exit sub 
        Public Sub ButtonclickToggle(ByRef cur As Cursor, Optional ByVal type As Cursor = Nothing)
            If type IsNot Nothing Then
                cur = type
                Exit Sub
            End If


            If cur = Cursors.WaitCursor Then
                cur = Cursors.Default
            Else
                cur = Cursors.WaitCursor
            End If
        End Sub
    End Module

    Public Module UnitTestingExtensions

        <Extension()>
        Public Function WriteAllToFile(ByVal writer As String, ByVal filepath As String, Optional ByVal append As Boolean = False) As Boolean
            Try
                ' Use a StreamWriter to write to the file
                ' The 'True' argument appends to the file if it already exists
                Using sw As New StreamWriter(filepath, append)
                    ' Write the log message to the file
                    sw.Write(writer)
                End Using

                Return True
            Catch ex As Exception
                ' Handle the exception
                Console.WriteLine("Error writing to log file: " & ex.Message)
                Return False
            End Try
        End Function

        <Extension()>
        Public Function WriteLineToFile(ByVal writer As String, ByVal filepath As String, Optional ByVal append As Boolean = False) As Boolean
            Try
                ' Use a StreamWriter to write to the file
                ' The 'True' argument appends to the file if it already exists
                Using sw As New StreamWriter(filepath, append)
                    ' Write the log message to the file
                    sw.WriteLine(writer)
                End Using

                Return True
            Catch ex As Exception
                ' Handle the exception
                Console.WriteLine("Error writing to log file: " & ex.Message)
                Return False
            End Try
        End Function

        'Archive files in a directory to another directory.
        <Extension()>
        Public Sub ArchiveFiles(ByVal dirTo As DirectoryInfo, ByVal dirFrom As String)
            For Each fold As DirectoryInfo In New DirectoryInfo(dirFrom).GetDirectories
                If Not fold.Name.ToLower.Contains("archive") Then
                    fold.MoveTo(dirTo.FullName & "\" & fold.Name)
                    frmMain.LogActivity("DEBUG | " & fold.Name & " moved to " & dirTo.FullName.Replace(dirFrom, ""))
                End If
            Next

            For Each file As FileInfo In New DirectoryInfo(dirFrom).GetFiles
                file.MoveTo(dirTo.FullName & "\" & file.Name)
                frmMain.LogActivity("DEBUG | " & file.Name & " moved to " & dirTo.FullName.Replace(dirFrom, ""))
            Next
        End Sub

        'Custom extension to sort the results datatables by check/failure mode and tool name
        '''Extension specific to datatables
        '''Adds a reference column for results comparison
        '''Names the table based on the optional parameter provided
        '''Sorts the datatable by Type and Tool
        <Extension()>
        Public Sub ResultsSorting(ByRef dt As DataTable, Optional ByVal addColumn As String = Nothing)
            If addColumn IsNot Nothing Then
                Dim newcolumn As New Data.DataColumn("Summary Type", GetType(System.String))
                newcolumn.DefaultValue = addColumn
                dt.Columns.Add(newcolumn)
            End If

            dt.Columns(1).ColumnName = "Rating Old"

            Dim newRatingColumn As New Data.DataColumn("Rating", GetType(System.String))
            dt.Columns.Add(newRatingColumn)

            For Each dr As DataRow In dt.Rows
                dr.Item("Rating") = dr.Item("Rating Old").ToString
            Next

            dt.Columns.Remove("Rating Old")
            If addColumn IsNot Nothing Then dt.TableName = addColumn
            dt.AsDataView.Sort = "Type ASC, Tool ASC"
        End Sub

        'Determine if 2 datatables have the same exact values 
        '''Returns a boolean determining if they are the sam
        '''Returns a databale of the compared values
        <Extension()>
        Public Function IsMatching(ByRef dt As DataTable, ByVal comparer As DataTable) As Tuple(Of Boolean, DataTable)
            Dim dtVal As Double = Double.NaN
            Dim comparerVal As Double = Double.NaN
            Dim delta As Double = Double.NaN
            Dim perDelta As Double = Double.NaN

            Dim matching As Boolean = True 'Item1
            Dim diffDt As New DataTable 'Item2
            diffDt.TableName = dt.TableName & " v. " & comparer.TableName
            diffDt.Columns.Add(dt.TableName & " File", GetType(System.String))
            diffDt.Columns.Add("Check/Failure Mode", GetType(System.String))
            diffDt.Columns.Add(dt.TableName & " Val", GetType(System.String))
            diffDt.Columns.Add(comparer.TableName & " Val", GetType(System.String))
            diffDt.Columns.Add("Delta", GetType(System.String))
            diffDt.Columns.Add("% Difference", GetType(System.String))
            diffDt.Columns.Add("Status", GetType(System.String))
            For i As Integer = 0 To Math.Max(dt.Rows.Count, comparer.Rows.Count) - 1
                Dim dtRow As DataRow
                Dim comparerRow As DataRow

                Try
                    dtRow = dt.Rows(i)
                Catch ex As Exception
                    dtRow = Nothing
                End Try

                If dtRow IsNot Nothing Then
                    For Each dr As DataRow In comparer.Rows
                        If dr.Item("Type").ToString = dtRow.Item("Type").ToString Then
                            If dr.Item("Tool").ToString.Contains(dtRow.Item("Tool").ToString) Or dtRow.Item("Tool").ToString.Contains(dr.Item("Tool").ToString) Then
                                comparerRow = dr
                                Exit For
                            Else
                                comparerRow = Nothing
                            End If
                        Else
                            comparerRow = Nothing
                        End If
                    Next

                    If IsNumeric(dtRow.Item("Rating")) Then
                        dtVal = CType(dtRow.Item("Rating"), Double)
                    End If
                Else
                    comparerRow = comparer.Rows(i)
                End If

                If comparerRow IsNot Nothing Then
                    If IsNumeric(comparerRow.Item("Rating")) Then
                        comparerVal = CType(comparerRow.Item("Rating"), Double)
                    End If
                End If

                If dtVal <> Double.NaN And comparerVal <> Double.NaN Then
                    delta = Math.Round(dtVal - comparerVal, 3)
                    perDelta = Math.Round((comparerVal - dtVal) / (dtVal) * 100, 2)
                End If

                Try
                    diffDt.Rows.Add(
                                dtRow.Item("Tool").ToString,
                                dtRow.Item("Type").ToString,
                                IIf(Double.IsNaN(dtVal), "N/A", dtVal),
                                IIf(Double.IsNaN(comparerVal), "N/A", comparerVal),
                                IIf(Double.IsNaN(delta), "N/A", delta),
                                IIf(Double.IsNaN(perDelta), "N/A", perDelta),
                                IIf(Double.IsNaN(delta) Or Math.Abs(delta) > 0.1, "Fail", "Pass")
                               )
                Catch ex As Exception

                End Try

                If delta = Double.NaN Or delta <> 0 Then
                    matching = False
                End If

                comparerVal = Double.NaN
                dtVal = Double.NaN
                delta = Double.NaN
                perDelta = Double.NaN
            Next

            Return New Tuple(Of Boolean, DataTable)(matching, diffDt)
        End Function

        'Extension for datatables to export to CSV using the datatabletocsv method
        '''Requires a filepath for where to save the csv
        <Extension()>
        Public Sub ToCSV(ByVal dt As DataTable, ByVal FilePath As String)
            DatatableToCSV(dt, FilePath)
        End Sub

        'Replaces all special characters in a string that aren't allowed in file folder or file names
        <Extension()>
        Public Function ToDirectoryString(ByVal str As String) As String
            str = str.Replace("#", "")
            str = str.Replace("%", "")
            str = str.Replace("&", "")
            str = str.Replace("{", "")
            str = str.Replace("}", "")
            str = str.Replace("/", "")
            str = str.Replace("\", "")
            str = str.Replace("<", "")
            str = str.Replace(">", "")
            str = str.Replace("*", "")
            str = str.Replace("?", "")
            str = str.Replace("$", "")
            str = str.Replace("!", "")
            str = str.Replace("'", "")
            str = str.Replace("""", "")
            str = str.Replace(":", "")
            str = str.Replace("@", "")
            str = str.Replace("+", "")
            str = str.Replace("`", "")
            str = str.Replace("|", "")
            str = str.Replace("=", "")

            Return str
        End Function

        <Extension()>
        Public Function TemplateVersion(ByVal file As FileInfo) As String
            Dim ver As String = Nothing
            Dim name As String = file.Name.Replace(file.Extension, "")
            Dim pattern As New Regex("\d+(\.\d+)+")
            Dim sMatch As Match = pattern.Match(name)

            If sMatch.Success Then
                ver = sMatch.Value
            Else
                ver = "-"
            End If

            Return ver
        End Function

        <Extension>
        Public Function ResultsToDataTable(ByVal myTnx As tnxModel) As DataTable
            Dim tnxDt As New DataTable
            tnxDt.Columns.Add("Type", GetType(System.String))
            tnxDt.Columns.Add("Rating", GetType(System.String))
            tnxDt.Columns.Add("Tool", GetType(System.String))
            tnxDt.TableName = myTnx.filePath.Replace(frmMain.dirUse & "\Test ID " & frmMain.testCase, "")

            For Each up In myTnx.geometry.upperStructure
                For Each res In up.Results
                    If res.rating > 0 Then tnxDt.Rows.Add(res.result_lkup, res.rating, "TNX Upper Section " & up.Rec)
                Next
            Next

            For Each down In myTnx.geometry.baseStructure
                For Each res In down.Results
                    If res.rating > 0 Then tnxDt.Rows.Add(res.result_lkup, res.rating, "TNX Base Section " & down.Rec)
                Next
            Next

            Return tnxDt
        End Function

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
        Public Sub AppendLog(ByVal strc As EDSStructure, ByVal pathToAppend As String, ByVal iteration As Integer)
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
                                sw.WriteLine(dt & " " & vars(0) & splt(2) & " " & separator & " " & Environment.UserName & " " & separator & "INFO" & " " & separator & vars(1) & " " & separator & " " & iteration)
                            ElseIf vars.Count = 1 Then
                                sw.WriteLine(dt & " " & separator & " " & Environment.UserName & " " & separator & "DEBUG" & " " & separator & vars(0) & " " & separator & " " & iteration)
                            Else
                                sw.WriteLine(dt & " " & vars(0) & splt(2) & " " & separator & " " & Environment.UserName & " " & separator & vars(1) & separator & vars(2) & " " & separator & " " & iteration)
                            End If
                        End If
                    End While
                    sr.Close()
                End Using
            End Using
        End Sub

        <Extension()>
        Public Sub AddERIs(ByVal myList As List(Of String), ByVal myDir As String, Optional ByVal purgeERI As Boolean = True)
            For Each info As FileInfo In New DirectoryInfo(myDir).GetFiles
                If info.Extension.ToLower() = ".eri" Then
                    'All eris permitted
                    myList.Add(info.FullName)
                    frmMain.LogActivity("DEBUG | ERI: " & info.Name & " found")
                ElseIf info.Name.ToLower.Contains(".eri.") Or info.Extension.ToLower = ".tfnx" Then
                    If purgeERI Then
                        info.Delete()
                        frmMain.LogActivity("DEBUG | File Deleted: " & info.Name & "")
                    End If
                End If
            Next
        End Sub
    End Module

    'Test cases are created when a test case is selected
    'These will correlate to the values in the CSV in the R: drive testing location
    Partial Public Class TestCase
        Public Property ID As Integer
        Public Property BU As Integer
        Public Property SID As String
        Public Property WO As Integer
        Public Property COMB As String
        Public Property SAWorkArea As String

        Public Sub New()

        End Sub

        Public Sub New(ByVal csvValue As String())
            Me.ID = csvValue(0)
            Me.BU = csvValue(1)
            Me.SID = csvValue(2)
            Me.WO = csvValue(3)
            Me.COMB = csvValue(4)
            Me.SAWorkArea = csvValue(5)
        End Sub

    End Class

    'JSON Serializer
    Public Module JsonUtil
        Public Function FromJsonString(Of T)(ByVal jsonString As String) As Tuple(Of T, String)
            Using aMemoryStream As MemoryStream = New MemoryStream(Encoding.UTF8.GetBytes(jsonString))
                Dim ser As DataContractJsonSerializer = New DataContractJsonSerializer(GetType(T))
                Dim myObj As T
                Dim resultTxt As String = "Success"
                Try
                    myObj = CType(ser.ReadObject(aMemoryStream), T)
                Catch ex As Exception
                    resultTxt = "ERROR DESERIALIZING " & ex.Message
                End Try

                Return New Tuple(Of T, String)(myObj, resultTxt)
            End Using
        End Function

        Public Function FromJsonString(Of T)(ByVal jsonString As String, ByVal serializerInstance As DataContractJsonSerializer) As T
            Using aMemoryStream As MemoryStream = New MemoryStream(Encoding.UTF8.GetBytes(jsonString))
                Dim ser = New DataContractJsonSerializer(GetType(T))
                Return CType(ser.ReadObject(aMemoryStream), T)
            End Using
        End Function

        Public Function ToJsonString(ByVal valueObject As Object, ByVal serializerInstance As DataContractJsonSerializer) As String
            Using aMemoryStream As MemoryStream = New MemoryStream()
                serializerInstance.WriteObject(aMemoryStream, valueObject)
                Return Encoding.[Default].GetString(aMemoryStream.ToArray())
            End Using
        End Function

        Public Function ToJsonString(Of T)(ByVal valueObject As T) As String
            Using aMemoryStream As MemoryStream = New MemoryStream()
                Dim serializer As DataContractJsonSerializer = New DataContractJsonSerializer(GetType(T))
                Try
                    serializer.WriteObject(aMemoryStream, valueObject)
                Catch ex As Exception
                    Return "ERROR SERIALIZING: " & ex.Message
                End Try

                Return Encoding.[Default].GetString(aMemoryStream.ToArray())
            End Using
        End Function
    End Module
End Namespace