Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop

Partial Public Class EDSStructure

    'include file name & extension for LogPath
    Public LogPath As String = ""
    Public WoLogPath As String = "C:\Users\" & Environment.UserName & "\source\repos\SAPI Maestro\_Logs"
    Public VerNum As String '= Assembly.GetEntryAssembly().GetName().Version.ToString
    'Tool paths
    Public SpliceCheckLocation As String = "C:\Users\" & Environment.UserName & "\Crown Castle USA Inc\Tower Assets Engineering - Engineering Templates\Monopole Splice Check"





    'pass in EDS Structure, run analysis based on tower type
    Public Sub Conduct(Optional isDevMode As Boolean = False)
        'exStruct As EDSStructure, workingAreaPath As String,
        Dim CCIPoleExists As Boolean = False
        Dim comp As Double
        Dim sheer As Double
        Dim mom As Double

        Dim barbCL As Double
        Dim plateComp As Double
        Dim plateSheer As Double
        Dim plateMom As Double
        Dim plateCompSeis As Double
        Dim plateSheerSeis As Double
        Dim plateMomSeis As Double
        'creates TNX file
        Dim poleMacCreateTNX As String = ""
        Dim poleMacRunTNXReactions As String = ""
        Dim poleMacRunTNXReactionsBARB As String = ""

        Dim plateMac As String = ""

        Dim spliceMacImportTNX As String = ""
        Dim spliceMacRun As String = ""
        '//
        Dim spliceCheckFile As String = ""

        'fnd macros
        Dim dpMac As String = ""
        Dim pierPadMac As String = ""
        Dim pileMac As String = ""
        Dim guyAnchorMac As String = ""

        Dim tnxFullPath As String = ""
        Dim tnxFileName As String = ""

        Dim strType As String = Me.SiteInfo.tower_type
        Dim poleWeUsin As Pole = Nothing
        Dim plateWeUsin As CCIplate = Nothing
        Dim basePlateConnection As Connection = Nothing
        Dim basePlateBoltGroup As BoltGroup = Nothing

        Dim workingAreaPath As String = Me.path 'ask Dan where this is

        Dim logFileName As String = Me.bus_unit & "_" & Me.structure_id & "_" & Me.work_order_seq_num & ".txt"


        GetVerNum()

        LogPath = String.Concat(workingAreaPath, logFileName)
        WriteLineLogLine("Beginning Maestro process..")

        Select Case strType
            Case "MONOPOLE"

                If Me.Poles.Count > 0 Then
                    CCIPoleExists = True
                    poleWeUsin = Me.Poles.FirstOrDefault
                    'create TNX file
                    OpenExcelRunMacro(poleWeUsin.workBookPath, poleMacCreateTNX, isDevMode)
                End If

                'not done yet
                ' If exStruct.Seismic Then
                'OpenExcelRunMacro()
                'End If

                'Run TNX
                tnxFullPath = Path.Combine(workingAreaPath, tnxFileName)
                RunTNX(tnxFullPath)

                If CCIPoleExists And Not IsNothing(poleWeUsin) Then
                    OpenExcelRunMacro(poleWeUsin.workBookPath, poleMacRunTNXReactions, isDevMode)
                End If

                spliceCheckFile = SpliceCheck(workingAreaPath)

                If Not spliceCheckFile = "" Then
                    OpenExcelRunMacro(spliceCheckFile, spliceMacImportTNX, isDevMode)
                    OpenExcelRunMacro(spliceCheckFile, spliceMacRun, isDevMode)
                End If

                If Me.CCIplates.Count > 0 Then
                    plateWeUsin = Me.CCIplates.FirstOrDefault
                    OpenExcelRunMacro(plateWeUsin.workBookPath, plateMac)

                    WriteLineLogLine("Checking for BARB..")

                    'barb
                    'If BARB exists, include in report = true and CCI Pole exists, execute barb logic
                    If plateWeUsin.barb_cl_elevation >= 0 And plateWeUsin.include_pole_reactions And CCIPoleExists Then
                        WriteLineLogLine("BARB elevation found..")
                        DoBARB(poleWeUsin, poleMacRunTNXReactionsBARB, isDevMode)

                    End If
                End If

                'get compression, sheer, and moment from TNX
                'GetCompSheerMomFromTNX(comp, sheer, mom)

                'loop through FNDs, open, input reactions & run macros
                'drilled pier
                If Me.DrilledPierTools.Count > 0 Then
                    WriteLineLogLine(Me.DrilledPierTools.Count & " Drilled Pier Fnd(s) found..")

                    For Each dp As DrilledPierFoundation In Me.DrilledPierTools
                        OpenExcelRunMacro(dp.workBookPath, dpMac, isDevMode)
                    Next
                End If
                'pier & pad
                If Me.PierandPads.Count > 0 Then
                    WriteLineLogLine(Me.PierandPads.Count & " Pier & Pad Fnd(s) found..")
                    For Each pierPad As PierAndPad In Me.PierandPads
                        OpenExcelRunMacro(pierPad.workBookPath, pierPadMac, isDevMode)
                    Next
                End If
                'Pile
                If Me.Piles.Count > 0 Then
                    WriteLineLogLine(Me.Piles.Count & " Pile Fnd(s) found..")
                    For Each pile In Me.Piles
                        OpenExcelRunMacro(pile.workBookPath, pileMac, isDevMode)
                    Next
                End If
                'Guy Anchor
                If Me.GuyAnchorBlocks.Count > 0 Then
                    WriteLineLogLine(Me.GuyAnchorBlocks.Count & " Guy Anchor Block Fnd(s) found..")
                    For Each guyAnc As GuyedAnchorBlock In Me.GuyAnchorBlocks
                        OpenExcelRunMacro(guyAnc.workbookpath, guyAnchorMac, isDevMode)
                    Next
                End If
                'run pdf macro
            Case "GUYED", "SELF-SUPPORT"

                'run leg reinforcement
                If exStruct.leg Then
                    OpenExcelRunMacro()
                End If

                ' run seismic macro to create eri with seismic loads if needed
                If exStruct.seismic Then
                    OpenExcelRunMacro()
                End If

                'run tnx
                tnxFullPath = Path.Combine(workingAreaPath, tnxFileName)
                RunTNX(tnxFullPath)

                'run leg reinforcement if exists
                If exStruct.leg Then
                    OpenExcelRunMacro()
                End If

                'loop through FNDs, open, input reactions & run macros



            Case Else
                'manual process
                WriteLineLogLine("Manual process due to Tower Type: " & strType)
                Exit Sub

        End Select


        'determine sufficiency

        'generate report

        'save results

    End Sub

    Public Function DoBARB(ByVal poleWeUsin As Pole, ByVal poleMacRunTNXReactions As String, Optional ByVal isDevMode As Boolean = False) As Boolean

        Dim barbCL As Double
        Dim plateComp As Double
        Dim plateSheer As Double
        Dim plateMom As Double



        Dim plateWeUsin As CCIplate = Nothing
        Dim basePlateConnection As Connection = Nothing
        Dim basePlateBoltGroup As BoltGroup = Nothing

        Dim applyBarb As Boolean = False

        'loop through connections to fine baseplate connection
        For Each connection In plateWeUsin.Connections
            If connection.connection_type = "Base" Then
                basePlateConnection = connection
                Exit For
            End If
        Next

        'loop through bolt groups and see if we apply the barb value
        For Each boltGroup In basePlateConnection.BoltGroups
            If boltGroup.apply_barb_elevation Then
                applyBarb = True
                basePlateBoltGroup = boltGroup
                WriteLineLogLine("Baseplate Bolt Group found for BARB..")

                Exit For
            End If
        Next

        If applyBarb And Not IsNothing(basePlateBoltGroup) Then
            WriteLineLogLine("Getting reactions for BARB..")
            barbCL = plateWeUsin.barb_cl_elevation 'exStruct.plate.barbCL

            If GetReactionsBARB(basePlateConnection, plateMom, plateComp, plateSheer) Then
                'replace values in Pole
                If Not BarbValuesIntoPole(poleWeUsin, poleWeUsin.workBookPath, barbCL, plateComp, plateSheer, plateMom, poleMacRunTNXReactions, isDevMode) Then
                    Return False
                End If
            Else
                Return False
            End If
            'Run TNX Reactions Macro in Pole
        ElseIf applyBarb = False Then
                WriteLineLogLine("Apply Barb = False")
            Else
                WriteLineLogLine("No Plate group found..")

        End If
        Return True
    End Function
#Region "Helpers"
    Public Function GetVerNum() As Boolean
        Try
            VerNum = Assembly.GetEntryAssembly().GetName().Version.ToString
        Catch ex As Exception
            VerNum = "BETA"
            Return False
        End Try
        Return True
    End Function

    Public Function OpenExcelRunMacro(excelPath As String, bigMac As String, Optional ByVal isDevEnv As Boolean = False) As String
        If String.IsNullOrEmpty(excelPath) Or String.IsNullOrEmpty(bigMac) Then
            Return "Error: excelPath or bigMac parameter is null or empty"
        End If

        Dim xlApp As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim xlWorkBook As Excel.Workbook = Nothing

        Dim errorMessage As String = String.Empty

        Dim xlVisibility As Boolean = False
        If isDevEnv Then
            xlVisibility = True
        End If

        Try
            If File.Exists(excelPath) Then

                xlApp = CreateObject("Excel.Application")
                xlApp.Visible = xlVisibility

                xlWorkBook = xlApp.Workbooks.Open(excelPath)

                xlApp.Run(bigMac)

                xlWorkBook.Save()
            Else
                errorMessage = $"Error: {excelPath} path not found!"
                WriteLineLogLine(errorMessage)
                Return errorMessage
            End If
        Catch ex As Exception
            errorMessage = ex.Message
            WriteLineLogLine(errorMessage)
            Return errorMessage
        Finally
            If xlWorkBook IsNot Nothing Then
                xlWorkBook.Close()
                Marshal.ReleaseComObject(xlWorkBook)
                xlWorkBook = Nothing
            End If
            If xlApp IsNot Nothing Then
                xlApp.Quit()
                Marshal.ReleaseComObject(xlApp)
                xlApp = Nothing
            End If
        End Try

        Return "Success"
    End Function

    Public Function BarbValuesIntoPole(pole As Pole, excelPath As String, barbCL As Double, plateComp As Double,
                                      plateShear As Double, plateMom As Double, poleMac As String, Optional isDevEnv As Boolean = False) As Boolean
        Dim xlApp As Microsoft.Office.Interop.Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet = Nothing
        Dim xlVisibility As Boolean = False

        Dim secID As String
        Dim rowNum As String
        Dim cellRangeComp As String
        Dim cellRangeShear As String
        Dim cellRangeMom As String

        'set xl visibility to true if dev environment
        If isDevEnv Then
            xlVisibility = True
        End If

        Try
            If File.Exists(excelPath) Then
                xlApp = CreateObject("Excel.Application")
                xlApp.Visible = xlVisibility

                xlWorkBook = xlApp.Workbooks.Open(excelPath)

                'replace reactions in all elevations at or below BARB CL
                For Each ws As Excel.Worksheet In xlWorkBook.Worksheets
                    If ws.Name.ToUpper = "RESULTS" Then
                        ws.Select()
                        xlWorkSheet = ws
                        Exit For
                    End If
                Next

                If Not xlWorkSheet Is Nothing Then
                    'loop through reinforced sections
                    For Each sec As PoleReinfSection In pole.reinf_sections 'exStruct.pole.reinfsection
                        'if top of section is at or below barbCL, edit values in coumns: IJK, row: local section ID +4
                        If sec.elev_top <= barbCL Then
                            secID = sec.ID
                            rowNum = secID + 4

                            'Comp/Pu = column I
                            cellRangeComp = "I" & rowNum
                            xlWorkSheet.Cells(cellRangeComp).Value = plateComp

                            'Moment/Mux= column J
                            cellRangeMom = "J" & rowNum
                            xlWorkSheet.Cells(cellRangeMom).Value = plateMom

                            'Shear/Vu = column K
                            cellRangeShear = "K" & rowNum
                            xlWorkSheet.Cells(cellRangeShear).Value = plateShear

                        End If
                    Next
                End If

                xlApp.Run(poleMac)

                xlWorkBook.Save()
                xlWorkBook.Close()
                xlApp.Quit()
            Else
                WriteLineLogLine(excelPath & " path not found!")
                Return False
            End If
        Catch ex As Exception
            WriteLineLogLine(ex.Message)
            Return False
        End Try


        Return True
    End Function

    Public Function SpliceCheck(workingAreaPath As String) As String
        Dim repoInfo As New tnxCCIReport(Me)
        Dim twrType As String = Me.tnx.geometry.AntennaType.ToUpper 'exStruct.tnx.geometry.upperStructure.ToString.ToUpper
        Dim spliceFile As String
        Dim workingSpliceFile As String
        'exStruct.tnx.geometry.AntennaType

        'exStruct.tnx.geometry.baseStructure.Count = 0
        'exStruct.tnx.geometry.upperStructure.Contains("pole")

        If repoInfo.sReportTowerManufacturer.ToUpper = "PYROD" And twrType = "TAPERED POLE" Then
            'copy splice tool
            'run macro to import TNX
            'run the "run" macro

            spliceFile = FindSpliceTool()

            If Not spliceFile = "" Then
                'copy tool
                workingSpliceFile = CopyFile(spliceFile, workingAreaPath)

                Return workingSpliceFile
            End If
        End If

        Return ""
    End Function

    Public Function RunTNX(tnxFilePath As String) As Boolean

        Try
            WriteLineLogLine("Running TNX..")
            WriteLineLogLine("---------------------------------------------------------------
---------------------------------------------------------------
---------------------------------------------------------------")

            Dim cmdProcess As New Process
            With cmdProcess
                .StartInfo = New ProcessStartInfo("TNX Tower.exe", tnxFilePath)
                With .StartInfo
                    .CreateNoWindow = True
                    .UseShellExecute = False
                    .RedirectStandardOutput = True
                End With
                .Start()
                .WaitForExit()
            End With ' Read output to a string variable.
            Dim ipconfigOutput As String = cmdProcess.StandardOutput.ReadToEnd


            WriteLineLogLine(ipconfigOutput)
            WriteLineLogLine("---------------------------------------------------------------
---------------------------------------------------------------
---------------------------------------------------------------")
            Return True

        Catch ex As Exception
            WriteLineLogLine("Error Running TNX: " & ex.Message)

            Return False
        End Try
    End Function

    'copy file from original path to new path
    'pass original file path with name/extension and new path without file name
    Public Function CopyFile(origPath As String, newPath As String) As String
        Dim fileName As String = ""
        Dim newPathWithFileName As String = ""

        Try
            fileName = Path.GetFileName(origPath)
            newPathWithFileName = Path.Combine(newPath, fileName)

            If Not File.Exists(origPath) Then
                Throw New FileNotFoundException("Original file not found: " + origPath)
            End If

            If Not Directory.Exists(newPath) Then
                Throw New DirectoryNotFoundException("Destination directory not found: " + newPath)
            End If

            File.Copy(origPath, newPathWithFileName, True)

            Return newPathWithFileName

        Catch ex As Exception
            WriteLineLogLine("Error copying file '" & fileName & "': " & ex.Message)
            Return ""
        End Try
    End Function
    'assuming this has a pole passed in
    Public Function FindSpliceTool(Optional extension As String = "xls") As String
        Dim fileName As String = ""

        ' Check if the directory exists
        If Not Directory.Exists(SpliceCheckLocation) Then
            WriteLineLogLine("Splice Check path not found!")
            Return ""
        End If

        ' Loop through the files in the directory
        For Each file In Directory.GetFiles(SpliceCheckLocation)
            fileName = Path.GetFileName(file).ToLower()

            ' Check if the file is a splice check file
            If fileName.Contains("splice check") And fileName.EndsWith("." & extension) Then
                Return file
            End If
        Next

        ' File not found
        WriteLineLogLine("Splice Check file not found!")
        Return ""
    End Function
#End Region

#Region "Maestro Logging"
    Public Sub WriteLineLogLine(msg As String)
        ' Get the current date and time
        Dim dt As DateTime = DateTime.Now

        ' Print the message to the console
        Console.WriteLine(dt & " | " & msg)

        ' Wrap the file operation in a try-catch block to handle exceptions
        Try
            ' Use a StreamWriter to write to the log file
            ' The 'True' argument appends to the file if it already exists
            Using sw As New StreamWriter(LogPath, True)
                ' If the log file does not exist, create it
                If Not File.Exists(LogPath) Then
                    sw.WriteLine("Maestro Log [START]" & vbCrLf & "Version: " & VerNum)
                End If
                ' Write the log message to the file
                sw.WriteLine(dt & " | " & msg)
            End Using
        Catch ex As Exception
            ' Handle the exception
            Console.WriteLine("Error writing to log file: " & ex.Message)
        End Try
    End Sub

    Public Function GetReactionsBARB(ByVal con As Connection, ByRef mom As Double?, ByRef comp As Double?, ByRef sheer As Double?) As Boolean
        Dim plateComp As Double
        Dim plateSheer As Double
        Dim plateMom As Double
        Dim plateCompSeis As Double
        Dim plateSheerSeis As Double
        Dim plateMomSeis As Double

        'get BARB values from Plate
        Try
            For Each result In con.ConnectionResults ' basePlateConnection.ConnectionResults
                Select Case result.result_lkup
                    Case "CONN_BARB_MOMENT"
                        plateMom = result.rating
                    Case "CONN_BARB_AXIAL"
                        plateComp = result.rating
                    Case "CONN_BARB_SHEAR"
                        plateSheer = result.rating
                    Case "CONN_BARB_MOMENT_SEISMIC"
                        plateMomSeis = result.rating
                    Case "CONN_BARB_AXIAL_SEISMIC"
                        plateCompSeis = result.rating
                    Case "CONN_BARB_SHEAR_SEISMIC"
                        plateSheerSeis = result.rating
                End Select
            Next

            CompareRatingsBARB(plateMom, plateComp, plateSheer, plateMomSeis, plateCompSeis, plateSheerSeis, mom, comp, sheer)

        Catch ex As Exception
            WriteLineLogLine("Exception finding BARB Ratings: " & ex.Message)
            Return False
        End Try

        Return True
    End Function
    'figure out which results to return
    'The following combination of reactions may be provided: wind only, seismic only, wind & seismic.
    'Multiple CCIpole objects will also need to be compared as applicable.
    'Suggested: If seismic and wind exist for the BU, compare both for largest moment and override CCIpole with respective reactions. 

    Public Function CompareRatingsBARB(ByVal plateMom As Double?, ByVal plateComp As Double?, ByVal plateSheer As Double?,
                                       ByVal plateMomSeis As Double?, ByVal plateCompSeis As Double?, ByVal plateSheerSeis As Double?,
                                       ByRef momToUse As Double, ByRef compToUse As Double, ByRef sheerToUse As Double) As Boolean

        Try
            If IsSomething(plateMom) And
                IsSomething(plateComp) And
                IsSomething(plateSheer) And
                IsSomething(plateMomSeis) And
                IsSomething(plateCompSeis) And
                IsSomething(plateSheerSeis) Then
                'compare wind and seis to see which one is larger
                'return larger
                momToUse = GetLarger(plateMom, plateMomSeis)
                compToUse = GetLarger(plateComp, plateCompSeis)
                sheerToUse = GetLarger(plateSheer, plateSheerSeis)

            ElseIf IsSomething(plateMomSeis) And
                IsSomething(plateCompSeis) And
                IsSomething(plateSheerSeis) Then
                'use seis values

                momToUse = plateMomSeis
                compToUse = plateCompSeis
                sheerToUse = plateSheerSeis


            ElseIf IsSomething(plateMom) And
            IsSomething(plateComp) And
            IsSomething(plateSheer) Then
                'use wind values

                momToUse = plateMom
                compToUse = plateComp
                sheerToUse = plateSheer

            Else
                'no sufficient values
                WriteLineLogLine("No BARB values found!")

            End If


        Catch ex As Exception
            WriteLineLogLine("Exception comparing BARB Ratings: " & ex.Message)
            Return False
        End Try
        Return True
    End Function
    'return larger of 2 values
    'return-1 if both vals are not numbers
    Public Function GetLarger(val1 As Double, val2 As Double) As Double
        If Double.IsNaN(val1) And Double.IsNaN(val2) Then
            WriteLineLogLine("Null Value for comparison")
            Return -1
        ElseIf Double.IsNaN(val1) Then
            Return val2
        ElseIf Double.IsNaN(val2) Then
            Return val1
        End If
        Return If(val1 > val2, val1, val2)

    End Function



#End Region
End Class
