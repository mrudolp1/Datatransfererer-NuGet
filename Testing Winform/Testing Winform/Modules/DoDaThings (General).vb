Imports System.IO
Imports DevExpress.XtraBars.ToastNotifications
Imports System.Xml
Imports System.Text
Imports CCI_Engineering_Templates
Imports System.Runtime.CompilerServices
Imports System.Runtime.Serialization.Json
Imports System.Text.RegularExpressions

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
    Public Function DoArchiving(ByVal folder As String) As String
        Dim arch As DirectoryInfo = Directory.CreateDirectory(folder & "\Archive " & Now.ToString("MM/dd/yyyy HH:mm:ss tt").ToDirectoryString)
        Dim archiveLog As String = ""
        archiveLog = arch.ArchiveFiles(folder)

        archiveLog.NewLine("DEBUG | Folder archived: " & folder)

        Return archiveLog
    End Function

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
        Next
    End Sub
End Module

