Imports System.IO
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Security.Principal
Imports System.Runtime.Serialization.Json
Imports System.Text
Imports DevExpress.CodeParser
Imports System.Threading
Imports CCI_Engineering_Templates.LogMessage

Public Module WorkflowHelpers
#Region "Import Input"


    'Import inputs for all files in a directory
    Public Async Function ImportingInputsAsync(
                                ByVal ImportingFrom As FileInfo,
                                ByVal saveToDirectory As String,
                                ByVal SAPICompatible As Boolean,
                                Optional ByVal excelVisible As Boolean = True,
                                Optional ByVal edsStructure As EDSStructure = Nothing,
                                Optional cancelToken As CancellationToken = Nothing,
                                Optional progress As IProgress(Of LogMessage) = Nothing) As Task(Of Boolean)


        Dim whichFileVar As (curver As Byte(), SAPIversion As Byte(), newFileName As String, resultsSheet As String, resultsRange As String, orderSheet As String, orderRange As String) = Await WhichFile(ImportingFrom, cancelToken, progress)
        Dim success As Boolean = True
        Dim ImportingTo As FileInfo
        'Item 1 = Excel application
        'Item 2 = Boolean (If true that means excel was previously open

        If whichFileVar.curver Is Nothing Then
            Await WriteLineLogLine("WARNING | Invalid Excel Tool: " & ImportingFrom.FullName, progress, True)
            Return False
        End If

        Dim getnewFileNameVar As String = Await GetNewFileName(saveToDirectory, fileName:=edsStructure.bus_unit.ToString() + " " + whichFileVar.newFileName.Replace(ImportingFrom.Extension.ToLower(), " (II).xlsm"), cancelToken:=cancelToken, progress:=progress)
        Dim fileCreated As Boolean = True
        Dim fileCreatedError As String = ""

        Try
            System.IO.File.WriteAllBytes(getnewFileNameVar, whichFileVar.SAPIversion)
            ImportingTo = New FileInfo(getnewFileNameVar)
        Catch ex As Exception
            fileCreated = False
            fileCreatedError = ex.Message
        End Try

        If Not fileCreated Then
            Await WriteLineLogLine("ERROR | New template file could not be created", progress, True)
            Await WriteLineLogLine("ERROR |  " & fileCreatedError, progress, True)
            Return False
        End If

        Dim myXL As (myApp As Excel.Application, alreadyOpen As Boolean) = Await GetXlApp(cancelToken, progress)
        If ImportingFrom.Extension.ToLower = ".xlsm" Then

            Dim macroname As String = "Import_Previous_Version"
            Dim prefix As String = ""
            Dim toolVer As String = ImportingFrom.TemplateVersion

            If progress IsNot Nothing And toolVer <> "-" Then
                Await WriteLineLogLine("DEBUG | Previous Tool Version: " & toolVer & " found for importing", progress, True)
            End If

            Dim params As (String, String, Boolean) = (ImportingFrom.FullName.ToString, toolVer, True)

            If ImportingTo.Name.ToLower.Contains("pile") Then
                macroname = "Button173_Click"
            ElseIf ImportingTo.Name.ToLower.Contains("drilled pier") Then
                If SAPICompatible Then
                    macroname += "_Performer"
                End If
            ElseIf ImportingTo.Name.ToLower.Contains("leg reinforcement") Then
                If SAPICompatible Then
                    prefix = "m_"
                End If
            End If

            Dim importer As Boolean = Await Import_Previous_Version(myXL.Item1, ImportingTo, macroname, params, excelVisible, prefix, whichFileVar.orderSheet, whichFileVar.orderRange, edsStructure, cancelToken, progress)
            success = importer

        End If

        DisposeXlApp(myXL.myApp, myXL.alreadyOpen)

        Return success
    End Function

    Private Async Sub WriteLineLogLine(ByVal msg As String)
        Return
    End Sub

    Private Async Function WriteLineLogLine(msg As String, progress As IProgress(Of LogMessage), Optional SkipMaeLog As Boolean = False) As Task
        Await Task.Run(Sub() WriteLineLogLine(msg))

        ''Raise a message logged event so these notifications can be passed up to the dashboard.
        If progress IsNot Nothing Then
            Dim msgs As String() = msg.Split(vbCrLf)
            For Each msg In msgs
                Dim msgSplt As String() = msg.Split("|")
                Dim logMsg As LogMessage
                Dim myType As LogMessage.MessageType

                If msgSplt.Length = 2 Then
                    'Sometimes it returns information with no type
                    'This information will default to ERROR since we handle everything else in the tools
                    Try
                        'Attempt to set the first item as a type
                        myType = DirectCast([Enum].Parse(GetType(MessageType), msgSplt(0).Trim), MessageType)
                        logMsg = New LogMessage(msgSplt(0).Trim, msgSplt(1).Trim, user:="Maestro")
                    Catch ex As Exception
                        'if it can't be set then it is most likely the acutal message
                        logMsg = New LogMessage("ERROR", msgSplt(1).Trim, user:="Maestro", timeStamp:=msgSplt(0).Trim)
                    End Try
                ElseIf msgSplt.Length = 3 Then
                    logMsg = New LogMessage(msgSplt(1).Trim, msgSplt(2).Trim, user:="Maestro", timeStamp:=msgSplt(0).Trim)
                ElseIf msgSplt.Length = 1 Then
                    'When the length is equal to 1 then it is only a message and should most likely be displayed as an error. 
                    logMsg = New LogMessage("ERROR", msgSplt(0).Trim, user:="Maestro", timeStamp:=Now.ToString())
                End If
                progress.Report(logMsg)
            Next
        End If

    End Function

    'Create or get the excel application to use.
    Public Async Function GetXlApp(Optional cancelToken As CancellationToken = Nothing, Optional progress As IProgress(Of LogMessage) = Nothing) As Task(Of (Excel.Application, Boolean))
        Try
            Return (GetObject(, "Excel.Appliction"), True)
        Catch ex As Exception
            Return (CreateObject("Excel.Application"), False)
        End Try
    End Function

    'Close the excel application if it was created 
    Public Function DisposeXlApp(ByRef xlapp As Excel.Application, isOpen As Boolean)
        If xlapp IsNot Nothing Then
            If Not isOpen Then
                Try
                    xlapp.Quit()
                    Marshal.ReleaseComObject(xlapp)
                Catch ex As Exception

                End Try

            End If

            xlapp = Nothing
        End If
    End Function

    Private Async Function SiteDataIntoSheet(ByVal xlWorkbook As Excel.Workbook,
                                               ByVal toolFileName As String,
                                               ByVal orderSheet As String,
                                               ByVal orderRange As String,
                                               ByVal edsStructure As EDSStructure,
                                               Optional cancelToken As CancellationToken = Nothing,
                                               Optional progress As IProgress(Of LogMessage) = Nothing
                                            ) As Task(Of Boolean)
        Dim savedOrder As Boolean = True
        Dim savedOrderError As String = ""
        'Some specific examples had to be built in because these tools handle the site data differently on the input tab.
        Try
            If toolFileName.ToLower.Contains("ccipole") Then
                xlWorkbook.Worksheets(orderSheet).Range(orderRange).value = edsStructure.work_order_seq_num?.ToString()
                Await WriteLineLogLine("DEBUG | WO Number " & edsStructure.work_order_seq_num?.ToString() & " added to workbook", progress, True)
            ElseIf toolFileName.ToLower.Contains("cciseismic") Then
                xlWorkbook.Worksheets("Site SDC Data").Range("wo").value = edsStructure.work_order_seq_num.ToString()
                xlWorkbook.Worksheets("Site SDC Data").Range("app").value = edsStructure.order?.ToString()
                xlWorkbook.Worksheets("Site SDC Data").Range("rev").value = edsStructure.orderRev?.ToString()
                Await WriteLineLogLine("DEBUG | WO Number " & edsStructure.work_order_seq_num?.ToString() & " added to workbook", progress, True)
                Await WriteLineLogLine("DEBUG | Order Number " & edsStructure.MyOrder().ToString() & " added to workbook", progress, True)
            Else
                xlWorkbook.Worksheets(orderSheet).Range(orderRange).value = edsStructure.MyOrder().ToString()
                Await WriteLineLogLine("DEBUG | Order Number " & edsStructure.MyOrder().ToString() & " added to workbook", progress, True)
            End If
        Catch ex As Exception
            savedOrder = False
            savedOrderError = ex.Message
            'Throwing this in a try-catch for the time being in case these ranges being editted have other impacts
        End Try

        If Not savedOrder Then
            Await WriteLineLogLine("WARNING | Could not add WO/Order data to workbook", progress, True)
            Await WriteLineLogLine("ERROR | " & savedOrderError, progress, True)
        End If

    End Function

    'Seb's macro runner adjusted specifically for unit testing
    Public Async Function Import_Previous_Version(
                                           ByVal xlapp As Excel.Application,
                                           ByVal workbookFile As FileInfo,
                                           ByVal macroName As String,
                                           ByVal params As (filepath As String, version As String, isMaesting As Boolean), 'Item1 = Filepath, Item2 = Version, Item3 = IsMaesting
                                           Optional ByVal xlVisibility As Boolean = False,
                                           Optional ByVal prefix As String = "",
                                           Optional ByVal orderSheet As String = "",
                                           Optional ByVal orderRange As String = "",
                                           Optional ByVal edsStructure As EDSStructure = Nothing,
                                           Optional cancelToken As CancellationToken = Nothing,
                                           Optional progress As IProgress(Of LogMessage) = Nothing
                                           ) As Task(Of Boolean)

        Dim toolFileName As String = Path.GetFileName(workbookFile.Name)
        Dim xlWorkBook As Excel.Workbook = Nothing
        Dim errorMessage As String = ""
        Dim isSuccess As Boolean = True
        Dim workbIssue As Boolean = False
        Dim workbMessage As String = ""

        If workbookFile Is Nothing Or String.IsNullOrEmpty(macroName) Then
            Await WriteLineLogLine("ERROR | workbookFile or macroName parameter is null or empty", progress, True)
            Return False
        End If

        Try
            If workbookFile.Exists Then

                xlapp.Visible = xlVisibility
                xlWorkBook = xlapp.Workbooks.Open(workbookFile.FullName)

                Await WriteLineLogLine("DEBUG | Tool: " & toolFileName, progress, True)

                'Check that the strings aren't empty and that ismaesting = true
                If params.Item1 IsNot Nothing And params.Item2 IsNot Nothing And params.Item3 Then
                    Await WriteLineLogLine("DEBUG | BEGIN MACRO: " & macroName, progress, True)
                    xlapp.Run(prefix & "Import_Previous_Version." & macroName, params.Item1, params.Item2, params.Item3)
                    Await WriteLineLogLine("DEBUG | END MACRO:  " & macroName, progress, True)


                    Await SiteDataIntoSheet(xlWorkBook, toolFileName, orderSheet, orderRange, edsStructure, cancelToken, progress)

                Else
                    Await WriteLineLogLine("ERROR | Parameters not specific ", progress, True)
                    Await WriteLineLogLine("DEBUG | Tool: " & toolFileName & " failed to import inputs", progress, True)
                    isSuccess = False
                End If

                xlWorkBook.Save()
            Else
                Await WriteLineLogLine("ERROR | " & workbookFile.FullName & " path not found!", progress, True)
            End If
        Catch ex As Exception
            errorMessage = ex.Message
            isSuccess = False
        Finally

            If xlWorkBook IsNot Nothing Then
                xlWorkBook.Close(True)
                Marshal.ReleaseComObject(xlWorkBook)
                xlWorkBook = Nothing
            End If
            Try
            Catch ex As Exception
                workbIssue = True
                workbMessage = ex.Message
            End Try
        End Try

        If Not isSuccess Then
            Await WriteLineLogLine("ERROR | " & errorMessage, progress, True)
        End If

        If workbIssue Then
            Await WriteLineLogLine("WARNING | Could not close Excel Workbook: " & toolFileName, progress, True)
            Await WriteLineLogLine("ERROR | " & workbMessage, progress, True)
        End If

        Return isSuccess
    End Function

#End Region

    Public Function LoadFileForViewing(ByVal info As FileInfo) As Tuple(Of DataTable, String)
        Dim loadDt As DataTable
        Dim myLog As String = ""

        'If the info is something (i.e. not a file folder) then attempt to load results or csv data
        If info IsNot Nothing Then
            If info.Extension.ToLower = ".xlsm" Then

                loadDt = SummarizedResults(info)
            ElseIf info.Extension.ToLower = ".csv" Then
                loadDt = CSVtoDatatable(info)
            ElseIf info.Extension.ToLower = ".ccistr" Then
                Dim myfrm As New Form
                Dim myPg As New PropertyGrid


                Dim tempStr As Tuple(Of EDSStructure, String)
                Using sr As New StreamReader(info.FullName)
                    tempStr = FromJsonString(Of EDSStructure)(sr.ReadToEnd)
                    sr.Close()
                End Using

                If tempStr.Item2.Contains("ERROR DESERIALIZING") Then
                    myLog.NewLine("ERROR | Structure not deserialized.")
                    myLog.NewLine("DEBUG | " & tempStr.Item2.ToString)
                Else
                    myPg.SelectedObject = tempStr.Item1
                    myPg.Dock = DockStyle.Fill

                    With myfrm
                        .FormBorderStyle = FormBorderStyle.SizableToolWindow
                        .Height = 600.0!
                        .Width = 500.0!
                        .Controls.Add(myPg)
                        .Text = info.Name
                        .StartPosition = FormStartPosition.CenterParent
                        .Show()
                    End With
                End If


            ElseIf info.Extension.ToLower = ".txt" Or info.Extension.ToLower = ".eri" Or
                       info.Extension.ToLower = ".log" Or info.Extension.ToLower = ".xml" Or
                       info.Extension.ToLower = ".sql" Or info.Extension.ToLower = ".ccimod" Or
                       info.Extension.ToLower = ".json" Or info.Extension.ToLower = ".svi" Then

                Using sr As New StreamReader(info.FullName)
                    Dim tempDt As New DataTable
                    tempDt.Columns.Add("Text")
                    Dim newRow As String()

                    While Not sr.EndOfStream
                        'newRow = sr.ReadLine.Split("|")
                        'If newRow.Count > 0 Then
                        '    If tempDt.Columns.Count = 0 Then
                        '        tempDt.Columns.Add("Time", GetType(System.String))
                        '        tempDt.Columns.Add("Message", GetType(System.String))
                        '    End If
                        '    Dim combined As String

                        '    Try
                        '        combined = newRow(1) & "|" & newRow(2)
                        '    Catch ex As Exception
                        '        Try
                        '            combined = newRow(1)
                        '        Catch ex1 As Exception
                        '            combined = ""
                        '        End Try
                        '    End Try

                        '    tempDt.Rows.Add(newRow(0), combined)
                        'Else
                        '    If tempDt.Columns.Count = 0 Then
                        '        tempDt.Columns.Add("Text", GetType(System.String))
                        '    End If
                        '    tempDt.Rows.Add(sr.ReadLine)
                        'End If
                        tempDt.Rows.Add(sr.ReadLine)
                    End While

                    loadDt = tempDt
                    sr.Close()
                End Using
            End If

            Return New Tuple(Of DataTable, String)(loadDt, myLog)

        End If

    End Function

    ''''Public Sub LoadMyWOS(ByVal site As SiteData, ByVal myGc As GridControl, ByVal mygv As GridView)
    ''''    Dim myDs As DataSet
    ''''    OracleLoader("SELECT wo_seqnum, eng_app_id, crrnt_rvsn_num, bus_unit, structure_id
    ''''                        FROM work_order_reporting_mv@ISITPRD.CROWNCASTLE.COM
    ''''                        WHERE bus_unit = '" & site.bus_unit.ToString & "' AND structure_id = '" & site.structure_id & "'
    ''''                        AND item_type IN ('SA - Structural Analysis','SA - Structural Analysis w/o App','SDD - Structural Design Drawings') 
    ''''                        ORDER BY wo_seqnum DESC",
    ''''                     "MyWOs", 5000, "ords")
    ''''    mygv.Columns.Clear()
    ''''    myGc.DataSource = Nothing
    ''''    myGc.DataSource = IDoDeclare.ds.Tables("MyWOs")
    ''''    myGc.RefreshDataSource()
    ''''    mygv.BestFitColumns(True)
    ''''End Sub

    'Return a datatable of summarized results from  a selected file
    'Invalid files return blank datatables

    Public Function SummarizedResults(ByVal info As IO.FileInfo) As DataTable

        Dim myTemplate As (Byte(), Byte(), String, String, String, String, String) = WhichFile(info).Result
        Dim range As String = myTemplate.Item5
        Dim tempds As New DataSet
        Dim finalDt As New DataTable
        Dim resultsDt As New DataTable

        'Determine if the selected file is a template
        If myTemplate.Item1 IsNot Nothing And myTemplate.Item2 IsNot Nothing And myTemplate.Item4 IsNot Nothing Then

            'There is potential for a template to have 2 specified ranges to import
            '''Drilled Pier
            '''CCIplate
            '''Tables are added to the temp dataset for each range in the workbook
            For Each str As String In myTemplate.Item5.Split("|")
                Try
                    tempds.Tables.Remove("Selected Results " & myTemplate.Item3 & "_" & str)
                Catch
                End Try

                tempds.Tables.Add(
                                        ExcelDatasourceToDataTable(
                                            GetExcelDataSource(
                                                    info.FullName, 'Path
                                                    myTemplate.Item4, 'Sheet Name
                                                    str), 'Range
                                             "Selected Results " & myTemplate.Item3 & "_" & str)) 'Datatable name
            Next

            'If it is a drilled pier determine which range is the correct range
            '''For monopoles and self supports, the range to select is just the summary from the 'Foundation Input' Tab
            If myTemplate.Item3.Contains("Drilled Pier") Then
                Try
                    If tempds.Tables("Selected Results " & myTemplate.Item3 & "_" & "BD8:CF59").Rows(0).Item("Guyed Tower Reactions").ToString = String.Empty Then
                        resultsDt = tempds.Tables("Selected Results " & myTemplate.Item3 & "_" & "H10:L31")
                    Else
                        resultsDt = tempds.Tables("Selected Results " & myTemplate.Item3 & "_" & "BD8:CF59")
                    End If
                Catch
                    resultsDt = tempds.Tables("Selected Results " & myTemplate.Item3 & "_" & "H10:L31")
                End Try
            ElseIf myTemplate.Item3.ToLower.Contains("cciplate") Then
                resultsDt = tempds.Tables("Selected Results " & myTemplate.Item3 & "_" & range)
            Else
                resultsDt = tempds.Tables("Selected Results " & myTemplate.Item3 & "_" & range)
            End If

            'Add columns to the final DT that shows the summary of the component, type and rating.
            finalDt.Columns.Add("Type", System.Type.GetType("System.String"))
            finalDt.Columns.Add("Rating", System.Type.GetType("System.String"))
            finalDt.Columns.Add("Tool", System.Type.GetType("System.String"))

            With resultsDt
                'Select case based on 'Filename_Range'
                Select Case .TableName
                    Case "Selected Results " & "CCIplate.xlsm" & "_" & "B1:BO64"
                        For i = 0 To .Rows.Count - 1
                            Dim dr As DataRow = .Rows(i)
                            Dim addl As String = ""
                            Dim val As String
                            If i > 31 Then addl = "_Seismic"

                            'Plate stress
                            If Not dr.Item("Plate Summary").ToString = "" And Not dr.Item("Plate Summary").ToString = "Max Stress" Then
                                val = dr.Item("Plate").ToString.Replace("%", "")
                                If val <> "N/A" And val <> "" Then finalDt.Rows.Add("Plate " & dr.Item("Flange ID").ToString & "_" & dr.Item("Column63").ToString & addl, val, info.Name.Replace(".xlsm", ""))

                                val = dr.Item("Column5").ToString.Replace("%", "")
                                If val <> "N/A" And val <> "" Then finalDt.Rows.Add("Plate " & dr.Item("Flange ID").ToString & "_" & "Tension Side Ratio" & addl, val, info.Name.Replace(".xlsm", ""))

                                val = dr.Item("Column6").ToString.Replace("%", "")
                                If val <> "N/A" And val <> "" Then finalDt.Rows.Add("Plate " & dr.Item("Flange ID").ToString & "_" & "Horizontal Weld" & addl, val, info.Name.Replace(".xlsm", ""))

                                val = dr.Item("Column7").ToString.Replace("%", "")
                                If val <> "N/A" And val <> "" Then finalDt.Rows.Add("Plate " & dr.Item("Flange ID").ToString & "_" & "Vertical Weld" & addl, val, info.Name.Replace(".xlsm", ""))

                                val = dr.Item("Column8").ToString.Replace("%", "")
                                If val <> "N/A" And val <> "" Then finalDt.Rows.Add("Plate " & dr.Item("Flange ID").ToString & "_" & "Flexure+Shear" & addl, val, info.Name.Replace(".xlsm", ""))

                                val = dr.Item("Column9").ToString.Replace("%", "")
                                If val <> "N/A" And val <> "" Then finalDt.Rows.Add("Plate " & dr.Item("Flange ID").ToString & "_" & "Tension+Shear" & addl, val, info.Name.Replace(".xlsm", ""))

                                val = dr.Item("Column10").ToString.Replace("%", "")
                                If val <> "N/A" And val <> "" Then finalDt.Rows.Add("Plate " & dr.Item("Flange ID").ToString & "_" & "Compression" & addl, val, info.Name.Replace(".xlsm", ""))

                                val = dr.Item("Column11").ToString.Replace("%", "")
                                If val <> "N/A" And val <> "" Then finalDt.Rows.Add("Plate " & dr.Item("Flange ID").ToString & "_" & "Punching Shear" & addl, val, info.Name.Replace(".xlsm", ""))

                            End If

                            'bolt group 1
                            If Not dr.Item("Bolt GR. 1").ToString = "" And Not dr.Item("Column21").ToString = "%" Then
                                val = dr.Item("Column21").ToString.Replace("%", "")
                                If val <> "N/A" And val <> "" Then finalDt.Rows.Add("Plate " & dr.Item("Flange ID").ToString & "_" & "Bolt Group 1" & addl, val, info.Name.Replace(".xlsm", ""))
                            End If

                            'bolt group 2
                            If Not dr.Item("Bolt GR. 2").ToString = "" And Not dr.Item("Column31").ToString = "%" Then
                                val = dr.Item("Column31").ToString.Replace("%", "")
                                If val <> "N/A" And val <> "" Then finalDt.Rows.Add("Plate " & dr.Item("Flange ID").ToString & "_" & "Bolt Group 2" & addl, val, info.Name.Replace(".xlsm", ""))
                            End If

                            'bolt group 3
                            If Not dr.Item("Bolt GR. 3").ToString = "" And Not dr.Item("Column41").ToString = "%" Then
                                val = dr.Item("Column41").ToString.Replace("%", "")
                                If val <> "N/A" And val <> "" Then finalDt.Rows.Add("Plate " & dr.Item("Flange ID").ToString & "_" & "Bolt Group 3" & addl, val, info.Name.Replace(".xlsm", ""))
                            End If

                            'bolt group 4
                            If Not dr.Item("Bolt GR. 4").ToString = "" And Not dr.Item("Column51").ToString = "%" Then
                                val = dr.Item("Column51").ToString.Replace("%", "")
                                If val <> "N/A" And val <> "" Then finalDt.Rows.Add("Plate " & dr.Item("Flange ID").ToString & "_" & "Bolt Group 4" & addl, val, info.Name.Replace(".xlsm", ""))
                            End If

                            'bolt group 5
                            If Not dr.Item("Bolt GR. 5").ToString = "" And Not dr.Item("Column61").ToString = "%" Then
                                val = dr.Item("Column61").ToString.Replace("%", "")
                                If val <> "N/A" And val <> "" Then finalDt.Rows.Add("Plate " & dr.Item("Flange ID").ToString & "_" & "Bolt Group 5" & addl, val, info.Name.Replace(".xlsm", ""))
                            End If


                        Next
                    Case "Selected Results " & "CCIpole.xlsm" & "_" & "AZ4:BT108"
                        For Each dr As DataRow In .Rows()
                            If Not dr.Item("Elevation (ft)").ToString = String.Empty Then
                                Dim val As Double
                                Try
                                    val = dr.Item("% Capacity") * 100
                                Catch ex As Exception
                                    val = 0
                                End Try
                                finalDt.Rows.Add(dr.Item("Elevation (ft)").ToString & "_" & dr.Item("Critical Element").ToString, val, info.Name.Replace(".xlsm", ""))
                            End If
                        Next
                    Case "Selected Results " & "Drilled Pier Foundation.xlsm" & "_" & "BD8:CF59"
                        For Each dr As DataRow In .Rows()
                            If Not dr.Item("Guyed Tower Reactions").ToString = String.Empty Then
                                Dim soilVal As Double
                                Dim strVal
                                Try
                                    soilVal = dr.Item("Soil Rating")
                                Catch ex As Exception
                                    soilVal = 0
                                End Try
                                Try
                                    strVal = dr.Item("Structural Rating")
                                Catch ex As Exception
                                    strVal = 0
                                End Try
                                finalDt.Rows.Add(dr.Item("Column1").ToString & "_" & dr.Item("Guyed Tower Reactions").ToString & "_Soil", soilVal, info.Name.Replace(".xlsm", ""))
                                finalDt.Rows.Add(dr.Item("Column1").ToString & "_" & dr.Item("Guyed Tower Reactions").ToString & "_Structural", strVal, info.Name.Replace(".xlsm", ""))
                            End If
                        Next
                    Case "Selected Results " & "Drilled Pier Foundation.xlsm" & "_" & "H10:L31"
                        NewFoundationRow(finalDt, .Rows(3), "Soil Lateral Check", "Compression", info)
                        NewFoundationRow(finalDt, .Rows(3), "Soil Lateral Check", "Uplift", info)
                        NewFoundationRow(finalDt, .Rows(10), "Soil Vertical Check", "Compression", info)
                        NewFoundationRow(finalDt, .Rows(10), "Soil Vertical Check", "Uplift", info)
                        NewFoundationRow(finalDt, .Rows(15), "Reinforced Concrete Flexure", "Compression", info)
                        NewFoundationRow(finalDt, .Rows(15), "Reinforced Concrete Flexure", "Uplift", info)
                        NewFoundationRow(finalDt, .Rows(20), "Reinforced Concrete Shear", "Compression", info)
                        NewFoundationRow(finalDt, .Rows(20), "Reinforced Concrete Shear", "Uplift", info)
                    Case "Selected Results " & "Guyed Anchor Block Foundation.xlsm" & "_" & "M20:X70"
                        For Each dr As DataRow In .Rows()
                            If Not dr.Item("Reaction Location").ToString = String.Empty Then
                                Dim soilVal As Double
                                Dim strVal As Double
                                Dim ancVal As Double
                                Try
                                    soilVal = dr.Item("Soil Rating") * 100
                                Catch ex As Exception
                                    soilVal = 0
                                End Try
                                Try
                                    strVal = dr.Item("Structural Rating") * 100
                                Catch ex As Exception
                                    strVal = 0
                                End Try
                                Try
                                    ancVal = dr.Item("Anchor Rating") * 100
                                Catch ex As Exception
                                    ancVal = 0
                                End Try
                                finalDt.Rows.Add(dr.Item("Column1").ToString & "_" & dr.Item("Reaction Location").ToString & "_Soil", soilVal, info.Name.Replace(".xlsm", ""))
                                finalDt.Rows.Add(dr.Item("Column1").ToString & "_" & dr.Item("Reaction Location").ToString & "_Structural", strVal, info.Name.Replace(".xlsm", ""))
                                finalDt.Rows.Add(dr.Item("Column1").ToString & "_" & dr.Item("Reaction Location").ToString & "_Anchor", ancVal, info.Name.Replace(".xlsm", ""))
                            End If
                        Next
                    Case "Selected Results " & "Pier and Pad Foundation.xlsm" & "_" & "F12:K25"
                        For Each dr As DataRow In .Rows()
                            If Not dr.Item("Column1").ToString = String.Empty Then
                                Dim val As Double
                                Try
                                    Try
                                        val = dr.Item("Rating*").ToString.Replace("%", "") * 100
                                    Catch exx As Exception
                                        val = dr.Item("Rating").ToString.Replace("%", "") * 100
                                    End Try
                                Catch ex As Exception
                                    val = 0
                                End Try
                                finalDt.Rows.Add(dr.Item("Column1").ToString, val, info.Name.Replace(".xlsm", ""))
                            End If
                        Next
                    Case "Selected Results " & "Pile Foundation.xlsm" & "_" & "G13:M31"
                        For Each dr As DataRow In .Rows()
                            If Not dr.Item("Column1").ToString = String.Empty Then
                                If dr.Item("Column1").ToString <> "PILE CHECKS" And dr.Item("Column1").ToString <> "BLOCK CHECKS" And
                                                     dr.Item("Column1").ToString <> "PAD CHECKS" And dr.Item("Column1").ToString <> "PIER CHECKS" Then
                                    Dim val As Double
                                    Try
                                        val = dr.Item("Rating*").ToString.Replace("%", "") * 100
                                    Catch ex As Exception
                                        val = 0
                                    End Try
                                    finalDt.Rows.Add(dr.Item("Column1").ToString, val, info.Name.Replace(".xlsm", ""))
                                End If
                            End If
                        Next
                    Case "Selected Results " & "SST Unit Base Foundation.xlsm" & "_" & "F12:K24"
                        For Each dr As DataRow In .Rows()
                            If Not dr.Item("Column1").ToString = String.Empty Then
                                Dim val As Double
                                Try
                                    val = dr.Item("Rating*").ToString.Replace("%", "") * 100
                                Catch ex As Exception
                                    val = 0
                                End Try
                                finalDt.Rows.Add(dr.Item("Column1").ToString, val, info.Name.Replace(".xlsm", ""))
                            End If
                        Next
                End Select
            End With
        End If

        Return finalDt
    End Function

    'Create a foundation row of results
    'Turns out this method is specific to Drilled Pier
    Public Sub NewFoundationRow(ByRef finaldt As DataTable, ByVal dr As DataRow, ByVal checkName As String, ByVal checkType As String, ByVal info As IO.FileInfo)
        With dr
            If Not .Item(checkType).ToString = "-" Then
                finaldt.Rows.Add(checkName & " " & checkType, .Item(checkType).ToString.Replace("%", ""), info.Name.Replace(".xlsm", ""))
            End If
        End With
    End Sub


    'Creates a structure object based on the files in the maestro folder for the current iteration
    Public Function CreateStructure(ByVal filesPath As String, ByRef strctr As EDSStructure, ByVal site As SiteData, ByVal EDSnewID As WindowsIdentity, ByVal EDSdbActive As String, Optional ByVal deleteAdditionalTNXfiles As Boolean = True) As String
        Dim myFiles As String()
        Dim myFilesLst As New List(Of String)
        Dim logString As String = ""

        'default resonse to determine if a question needs asked.
        'Dim response As DialogResult = DialogResult.Cancel

        'Loop through all files in the maestro folder for the current test case and iteration
        For Each info As FileInfo In New DirectoryInfo(filesPath).GetFiles
            If info.Extension = ".eri" Then
                'All eris permitted
                myFilesLst.Add(info.FullName)
                logString.NewLine("DEBUG | File found for structure: " & info.Name)
            ElseIf info.Extension = ".xlsm" Then 'All tools are current xlsm files and this should be a safe assumption
                'Determine if the file is one of the templates
                Dim template As (Byte(), Byte(), String, String, String, String, String) = WhichFile(info).Result

                'If the properties of the tuple are nothing then they aren't templates
                If template.Item1 IsNot Nothing And template.Item2 IsNot Nothing And template.Item3 IsNot Nothing Then
                    myFilesLst.Add(info.FullName)
                    logString.NewLine("DEBUG | File found for structure: " & info.Name)
                End If
            ElseIf info.Name.ToLower.Contains(".eri.") Or info.Extension.ToLower = ".tfnx" Then
                'If response = DialogResult.Cancel Then
                '    response = MsgBox("Would you like to rerun the ERI file as well?", vbYesNo + vbInformation, "Rerun TNX?")
                '    If response = DialogResult.No Then
                '        LogActivity("DEBUG | tnx NOT rerun for Maestro")
                '    Else
                '        LogActivity("DEBUG | tnx will be rerun for Maestro")
                '    End If
                'End If
                'If response = DialogResult.Yes Then
                If deleteAdditionalTNXfiles Then
                    info.Delete()
                    logString.NewLine("DEBUG | File Deleted: " & info.FullName)
                End If
                'End If
            End If
        Next

        'Convert the list of valid file names to an array for creating anew structure
        myFiles = myFilesLst.ToArray
        strctr = New EDSStructure(site.bus_unit, site.structure_id, site.work_order_seq_num, filesPath, filesPath, myFiles, EDSnewID, EDSdbActive)

        Return logString
    End Function

    'Determines the file name for the new templates being saved.
    'Increments file names if they arleady exist in the new directory.
    Public Async Function GetNewFileName(ByVal newFolder As String, ByVal Optional file As FileInfo = Nothing, ByVal Optional fileName As String = Nothing,
                                Optional cancelToken As CancellationToken = Nothing,
                                Optional progress As IProgress(Of LogMessage) = Nothing) As Task(Of String)
        Dim counter As Integer = 0
        Dim filePath As String

        If Not file Is Nothing Then
            filePath = newFolder & "\" & file.Name
        Else
            filePath = newFolder & "\" & fileName
        End If

        While IO.File.Exists(filePath)
            counter += 1
            If file IsNot Nothing Then
                filePath = newFolder & "\" & file.Name.Split(".")(0) & "(" & counter.ToString() & ")." & file.Name.Split(".")(1)
            Else
                filePath = newFolder & "\" & fileName.Split(".")(0) & "(" & counter.ToString() & ")." & fileName.Split(".")(1)
            End If
        End While

        Return filePath
    End Function

    'Used to determine which template is being used
    'This could have been set up as a class but ended up going too far and now we have tuples. Enjoy! :)
    Public Async Function WhichFile(ByVal file As FileInfo,
                                Optional cancelToken As CancellationToken = Nothing,
                                Optional progress As IProgress(Of LogMessage) = Nothing) As Task(Of (Byte(), Byte(), String, String, String, String, String))
        Dim returner As (Byte(), Byte(), String, String, String, String, String)
        'Item 1 = current published versions
        'Item 2 = new versions created for SAPI
        'Item 3 = File name to be used with the bytes
        'Item 4 = Worksheet with results
        'Item 5 = Range for results 
        'Item 6 = Worksheet for order input
        'Item 7 = Range for order

        If file.Name.ToLower.Contains("cciplate") Then
            returner = (
                CCI_Engineering_Templates.My.Resources.CCIplate__4_1_2_,
                CCI_Engineering_Templates.My.Resources.CCIplate,
                "CCIplate.xlsm",
                "Results Database",
                "B1:BO64",
                "Main",
                "C5")
        ElseIf file.Name.ToLower.Contains("ccipole") Then
            returner = (
                CCI_Engineering_Templates.My.Resources.CCIpole__4_5_8_,
                CCI_Engineering_Templates.My.Resources.CCIpole,
                "CCIpole.xlsm",
                "Results",
                "AZ4:BT108",
                "Input",
                "WO")
        ElseIf file.Name.ToLower.Contains("cciseismic") Then
            returner = (
                CCI_Engineering_Templates.My.Resources.CCISeismic__3_3_9_,
                CCI_Engineering_Templates.My.Resources.CCISeismic,
                "CCISeismic.xlsm",
                Nothing,
                Nothing,
                "",
                "")
        ElseIf file.Name.ToLower.Contains("drilled pier") Then
            returner = (
                CCI_Engineering_Templates.My.Resources.Drilled_Pier_Foundation__5_0_5_,
                CCI_Engineering_Templates.My.Resources.Drilled_Pier_Foundation,
                "Drilled Pier Foundation.xlsm",
                "Foundation Input",
                "BD8:CF59|H10:L31",
                "Foundation Input",
                "D5")
        ElseIf file.Name.ToLower.Contains("guyed anchor") Then
            returner = (
                CCI_Engineering_Templates.My.Resources.Guyed_Anchor_Block_Foundation__4_0_0_,
                CCI_Engineering_Templates.My.Resources.Guyed_Anchor_Block_Foundation,
                "Guyed Anchor Block Foundation.xlsm",
                "Input",
                "M20:X70",
                "Input",
                "R5")
        ElseIf file.Name.ToLower.Contains("leg reinforcement") Then
            returner = (
                CCI_Engineering_Templates.My.Resources.Leg_Reinforcement_Tool__10_0_4_,
                CCI_Engineering_Templates.My.Resources.Leg_Reinforcement_Tool,
                "Leg Reinforcement Tool.xlsm",
                Nothing,
                Nothing,
                "IMPORT",
                "Order_Import")
        ElseIf file.Name.ToLower.Contains("pier and pad") Then
            returner = (
                CCI_Engineering_Templates.My.Resources.Pier_and_Pad_Foundation__4_1_1_,
                CCI_Engineering_Templates.My.Resources.Pier_and_Pad_Foundation,
                "Pier and Pad Foundation.xlsm",
                "Input",
                "F12:K25",
                "Input",
                "C5")
        ElseIf file.Name.ToLower.Contains("pile") Then
            returner = (
                CCI_Engineering_Templates.My.Resources.Pile_Foundation__2_2_1_,
                CCI_Engineering_Templates.My.Resources.Pile_Foundation,
                "Pile Foundation.xlsm",
                "Input",
                "G13:M31",
                "Input",
                "C7")
        ElseIf file.Name.ToLower.Contains("unit base") Then
            returner = (
                CCI_Engineering_Templates.My.Resources.SST_Unit_Base_Foundation__4_0_3_,
                CCI_Engineering_Templates.My.Resources.SST_Unit_Base_Foundation,
                "SST Unit Base Foundation.xlsm",
                "Input",
                "F12:K24",
                "Input",
                "C5")
        Else
            returner = (
                Nothing,
                Nothing,
                Nothing,
                Nothing,
                Nothing,
                Nothing,
                Nothing)
        End If

        Return returner
    End Function

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

    'Determine if a file is open
    Private Function FileIsOpen(ByVal file As FileInfo) As Boolean
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

    'Gets a combined datatbale of results for all spreadsheets in directory.
    Public Sub GetAllResults(ByVal folder As String)
        Dim combinedResults As New DataTable
        'Loop through all files in the specified folder
        For Each info As FileInfo In New DirectoryInfo(folder).GetFiles
            If info.Extension.ToLower = ".xlsm" And Not info.Name.Contains("~") Then
                'Merge the datatable to append all data together
                combinedResults.Merge(SummarizedResults(info))
            End If
        Next

        'Save the datatable to a CSV in the specified folder location
        DatatableToCSV(combinedResults, folder & "\Summarized Results.csv")
    End Sub
End Module

