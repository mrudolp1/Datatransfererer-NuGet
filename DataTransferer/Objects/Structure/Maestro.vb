Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Threading
Imports CCI_Engineering_Templates.LogMessage
Imports Excel = Microsoft.Office.Interop.Excel

Partial Public Class EDSStructure

    'include file name & extension for LogPath
    Public LogPath As String = ""
    Public WoLogPath As String = "C:\Users\" & Environment.UserName & "\source\repos\SAPI Maestro\_Logs"
    Public VerNum As String '= Assembly.GetEntryAssembly().GetName().Version.ToString
    'Tool paths
    Public SpliceCheckLocation As String = "C:\Users\" & Environment.UserName & "\Crown Castle USA Inc\Tower Assets Engineering - Engineering Templates\Monopole Splice Check"
    Public AnalysisRunning As Boolean = False

    ''' <summary>
    ''' Perform structural analysis on all files in the structure.
    ''' </summary>
    ''' <param name="isDevMode">Determines if the beta version of TNX should be used.</param>
    ''' <param name="xlVisibility">Determines excel tools appear in the foreground (true) or are hidden (false).</param>
    Public Async Function ConductAsync(RunTNX As Boolean, Optional isDevMode As Boolean = False, Optional ByVal xlVisibility As Boolean = False, Optional cancelToken As CancellationToken = Nothing, Optional progress As IProgress(Of LogMessage) = Nothing) As Task(Of Boolean)
        AnalysisRunning = True
        Dim dt As String = DateTime.Now.ToString.Replace("/", "-").Replace(":", ".")

        Dim CCIPoleExists As Boolean = False

        'MACRO VARS
        '//pole macros
        Dim poleMacCreateTNX As String = "MaestMe_Step1" ' step 1
        Dim poleMacImportTNXReactions As String = "MaestMe_Step2" 'step 2
        Dim poleMacRunAnalysis As String = "MaestMe_Step3" ' step 3
        'Dim poleMacRunTNXReactionsBARB As String = ""
        '//plate macros
        Dim plateMac As String = "MaestMe"
        '//splice check
        Dim spliceMacImportTNX As String = ""
        Dim spliceMacRun As String = ""
        '//
        Dim spliceCheckFile As String = ""

        '/fnd macros
        '//monopole
        'Dim drilledPierMac As String = ""
        Dim pierPadMac As String = "MaestMe"
        Dim pileMac As String = "MaestMe"
        Dim guyAnchorMac As String = "MaestMe"
        '//lattice
        Dim legReinforcementMac As String = "MaestMe"
        Dim unitBaseMac As String = "MaestMe"
        Dim drilledPierMac As String = "MaestMe"
        'Dim pierAndPadMac As String = "MaestMe"
        Dim seisMac As String = "MaestMe"

        'TNX vars
        Dim tnxFullPath As String = Me.tnx.filePath '""
        ' Dim tnxFileName As String = ""

        Dim excelResult As OpenExcelRunMacroResult = New OpenExcelRunMacroResult()

        Dim strType As String = Me.SiteInfo.tower_type.ToUpper
        Dim poleWeUsin As Pole = Nothing
        Dim plateWeUsin As CCIplate = Nothing
        Dim seismicWeUsin As CCISeismic = Nothing
        Dim basePlateConnection As Connection = Nothing
        Dim basePlateBoltGroup As BoltGroup = Nothing

        Dim workingAreaPath As String = Me.WorkingDirectory 'ask Dan where this is, ' ask Seb if he found this, 'Seb found this

        Dim logFileName As String = Me.bus_unit & "_" & Me.structure_id & "_" & Me.work_order_seq_num & "_" & dt & ".txt"

        GetVerNum()

        LogPath = Path.Combine(workingAreaPath, logFileName)

        CreateLogFile()

        Await WriteLineLogLine("INFO | Beginning Maestro process..", progress)

        Select Case strType
            Case "MONOPOLE"
                'CCI Pole Step 1 - Create TNX
                If Me.Poles.Count > 0 Then
                    If Me.Poles.Count > 1 Then
                        Await WriteLineLogLine("WARNING | " & Me.Poles.Count & " CCIPole files found! Using first or default..", progress)
                    End If
                    CCIPoleExists = True
                    poleWeUsin = Me.Poles.FirstOrDefault
                    'create TNX file
                    excelResult = Await OpenExcelRunMacroAsync(poleWeUsin, poleMacCreateTNX, xlVisibility, toolName:="CCIPole - Step 1", cancelToken:=cancelToken, progress:=progress)
                    If Not excelResult.Success OrElse (cancelToken <> CancellationToken.None AndAlso cancelToken.IsCancellationRequested) Then GoTo ErrorSkip
                End If

                If Me.CCISeismics.Count > 0 Then
                    If Me.CCISeismics.Count > 1 Then
                        Await WriteLineLogLine("WARNING | " & Me.CCISeismics.Count & " CCISeismic files found! Using first or default..", progress)
                    End If
                    seismicWeUsin = Me.CCISeismics.FirstOrDefault
                    excelResult = Await OpenExcelRunMacroAsync(seismicWeUsin, seisMac, xlVisibility, cancelToken:=cancelToken, progress:=progress)
                    If Not excelResult.Success OrElse (cancelToken <> CancellationToken.None AndAlso cancelToken.IsCancellationRequested) Then GoTo ErrorSkip
                End If

                'Run TNX
                If RunTNX Then
                    If Not Await RunTNXAsync(tnxFullPath, isDevMode, cancelToken, progress) OrElse (cancelToken <> CancellationToken.None AndAlso cancelToken.IsCancellationRequested) Then GoTo ErrorSkip
                End If
                'CCI Pole step 2 - pull in reactions
                If CCIPoleExists And Not IsNothing(poleWeUsin) Then
                    excelResult = Await OpenExcelRunMacroAsync(poleWeUsin, poleMacImportTNXReactions, xlVisibility, toolName:="CCIPole - Step 2", cancelToken:=cancelToken, progress:=progress)
                    If Not excelResult.Success OrElse (cancelToken <> CancellationToken.None AndAlso cancelToken.IsCancellationRequested) Then GoTo ErrorSkip
                End If

                'Splice Check
                'spliceCheckFile = "" 'SpliceCheck(workingAreaPath)

                'If Not spliceCheckFile = "" Then
                '    If CheckForSuccess(OpenExcelRunMacro(spliceCheckFile, spliceMacImportTNX, isDevMode), "Splice Check - Import TNX") = False Then
                '        GoTo ErrorSkip
                '    End If
                '    If CheckForSuccess(OpenExcelRunMacro(spliceCheckFile, spliceMacRun, isDevMode), "Splice Check - Run") = False Then
                '        GoTo ErrorSkip
                '    End If
                'End If

                If Me.CCIplates.Count > 0 Then
                    If Me.CCIplates.Count > 1 Then
                        Await WriteLineLogLine("WARNING | " & Me.CCIplates.Count & " CCIPlate files found! Using first or default..", progress)
                    End If
                    plateWeUsin = Me.CCIplates.FirstOrDefault
                    excelResult = Await OpenExcelRunMacroAsync(plateWeUsin, plateMac, xlVisibility, cancelToken:=cancelToken, progress:=progress)
                    If Not excelResult.Success OrElse (cancelToken <> CancellationToken.None AndAlso cancelToken.IsCancellationRequested) Then GoTo ErrorSkip

                    Await WriteLineLogLine("INFO | Checking for BARB..", progress)

                    'barb
                    'If BARB exists, include in report = true and CCI Pole exists, execute barb logic
                    If plateWeUsin.barb_cl_elevation >= 0 And plateWeUsin.include_pole_reactions And CCIPoleExists Then
                        Await WriteLineLogLine("INFO | BARB elevation found..", progress)
                        Await DoBARB(poleWeUsin, plateWeUsin, xlVisibility, cancelToken, progress)
                    End If
                End If

                'CCI Pole step 3 - Run Analysis
                If CCIPoleExists And Not IsNothing(poleWeUsin) Then
                    excelResult = Await OpenExcelRunMacroAsync(poleWeUsin, poleMacRunAnalysis, xlVisibility, toolName:="CCIPole - Step 3", cancelToken:=cancelToken, progress:=progress)
                    If Not excelResult.Success OrElse (cancelToken <> CancellationToken.None AndAlso cancelToken.IsCancellationRequested) Then GoTo ErrorSkip
                End If



                'get compression, sheer, and moment from TNX
                'GetCompSheerMomFromTNX(comp, sheer, mom)

                'loop through FNDs, open, input reactions & run macros
                '//drilled pier
                If Me.DrilledPierTools.Count > 0 Then
                    Await WriteLineLogLine("INFO | " & Me.DrilledPierTools.Count & " Drilled Pier Fnd(s) found..", progress)

                    'For Each dp As DrilledPierFoundation In Me.DrilledPierTools
                    For i As Integer = 0 To DrilledPierTools.Count - 1
                        excelResult = Await OpenExcelRunMacroAsync(DrilledPierTools(i), drilledPierMac, xlVisibility, cancelToken:=cancelToken, progress:=progress)
                        If Not excelResult.Success OrElse (cancelToken <> CancellationToken.None AndAlso cancelToken.IsCancellationRequested) Then GoTo ErrorSkip
                    Next
                End If
                '//pier & pad
                If Me.PierandPads.Count > 0 Then
                    Await WriteLineLogLine("INFO | " & Me.PierandPads.Count & " Pier & Pad Fnd(s) found..", progress)
                    'For Each pierPad As PierAndPad In Me.PierandPads
                    For i As Integer = 0 To PierandPads.Count - 1
                        'Dim tempPath As String = Path.Combine("C:\Users\stanley\Crown Castle USA Inc\ECS - Tools\SAPI Test Cases\808466\2199162", "808466 Pier and Pad Foundation.xlsm")
                        'OpenExcelRunMacro(tempPath, pierPadMac, isDevMode)
                        excelResult = Await OpenExcelRunMacroAsync(PierandPads(i), pierPadMac, xlVisibility, cancelToken:=cancelToken, progress:=progress)
                        If Not excelResult.Success OrElse (cancelToken <> CancellationToken.None AndAlso cancelToken.IsCancellationRequested) Then GoTo ErrorSkip
                    Next
                End If
                '//Pile
                If Me.Piles.Count > 0 Then
                    Await WriteLineLogLine("INFO | " & Me.Piles.Count & " Pile Fnd(s) found..", progress)
                    'For Each pile As Pile In Me.Piles
                    For i As Integer = 0 To Me.Piles.Count - 1
                        excelResult = Await OpenExcelRunMacroAsync(Piles(i), pierPadMac, xlVisibility, cancelToken:=cancelToken, progress:=progress)
                        If Not excelResult.Success OrElse (cancelToken <> CancellationToken.None AndAlso cancelToken.IsCancellationRequested) Then GoTo ErrorSkip
                    Next
                End If
                '//Guy Anchor

                If Me.GuyAnchorBlockTools.Count > 0 Then
                    Await WriteLineLogLine("INFO | " & Me.GuyAnchorBlockTools.Count & " Guy Anchor Block Fnd(s) found..", progress)
                    'For Each guyAnc As AnchorBlockFoundation In Me.GuyAnchorBlockTools
                    For i As Integer = 0 To GuyAnchorBlockTools.Count - 1
                        excelResult = Await OpenExcelRunMacroAsync(GuyAnchorBlockTools(i), guyAnchorMac, xlVisibility, cancelToken:=cancelToken, progress:=progress)
                        If Not excelResult.Success OrElse (cancelToken <> CancellationToken.None AndAlso cancelToken.IsCancellationRequested) Then GoTo ErrorSkip
                    Next
                End If

            Case "GUYED", "SELF SUPPORT"

                'Check if TNX has been ran. if not, run it
                'Dan will provide a file path to check for in the working directory

                If RunTNX Then
                    '/run tnx
                    If Not Await RunTNXAsync(tnxFullPath, isDevMode, cancelToken, progress) OrElse (cancelToken <> CancellationToken.None AndAlso cancelToken.IsCancellationRequested) Then GoTo ErrorSkip
                End If

                '/run seismic macro to create eri with seismic loads if needed
                If Me.CCISeismics.Count > 0 Then
                    If Me.CCISeismics.Count > 1 Then
                        Await WriteLineLogLine("WARNING | " & Me.CCISeismics.Count & " CCISeismic files found! Using first or default..", progress)
                    End If
                    seismicWeUsin = Me.CCISeismics.FirstOrDefault
                    ' plateWeUsin = Me.CCIplates.FirstOrDefault

                    'run seismic.
                    excelResult = Await OpenExcelRunMacroAsync(seismicWeUsin, seisMac, xlVisibility, cancelToken:=cancelToken, progress:=progress)
                    If Not excelResult.Success OrElse (cancelToken <> CancellationToken.None AndAlso cancelToken.IsCancellationRequested) Then GoTo ErrorSkip

                    'Check if seismic analysis is required
                    If excelResult.SeismicRequired Then
                        Await WriteLineLogLine("INFO | Seismic Analysis required. Rerunning TNX.", progress)
                        Await WriteLineLogLine("WARNING | Seismic loading included in TNX analysis. Further evaluation required to determine if 1.5 overstrength factor controls.", progress)

                        '/run tnx
                        If Not Await RunTNXAsync(tnxFullPath, isDevMode, cancelToken, progress) OrElse (cancelToken <> CancellationToken.None AndAlso cancelToken.IsCancellationRequested) Then GoTo ErrorSkip
                    End If

                End If


                '/run leg reinforcement
                '//compare previous geometry to current geometry
                '/run leg reinforcement if exists
                If Me.LegReinforcements.Count > 0 Then
                    If IsSomething(Me.EDSMe) Then
                        'if we get here, there's a previous EDS structure
                        'check if EDS geometry exists
                        'check if EDS leg reinforcement exists
                        'compare geometry
                        'compare leg reinforcement
                        If IsNothing(EDSMe.tnx.geometry) Or
                          EDSMe.LegReinforcements.Count = 0 Or
                          Not Me.tnx.geometry.Equals(EDSMe.tnx.geometry) Or
                          Not Me.LegReinforcements.Equals(EDSMe.LegReinforcements) Then
                            Await WriteLineLogLine("WARNING | Leg Reinforcement not found or could not verify TNX Leg Reinforcement. Make sure it is generated before running maestro.", progress)
                        End If
                    End If

                    'For Each legReinforcement As LegReinforcement In LegReinforcements
                    For i As Integer = 0 To LegReinforcements.Count - 1
                        excelResult = Await OpenExcelRunMacroAsync(LegReinforcements(i), legReinforcementMac, xlVisibility, cancelToken:=cancelToken, progress:=progress)
                        If Not excelResult.Success OrElse (cancelToken <> CancellationToken.None AndAlso cancelToken.IsCancellationRequested) Then GoTo ErrorSkip
                    Next
                Else
                    'await writelinelogline("WARNING | No Leg Reinforcement found! Could not verify Leg Reinforcement.")
                End If

                '/if plate exists, run - if Chris can update plate easily
                If Me.CCIplates.Count > 0 Then
                    If Me.CCIplates.Count > 1 Then
                        Await WriteLineLogLine("WARNING | " & Me.CCIplates.Count & " CCIPlate files found! Using first or default..", progress)
                    End If
                    plateWeUsin = Me.CCIplates.FirstOrDefault
                    excelResult = Await OpenExcelRunMacroAsync(plateWeUsin, plateMac, xlVisibility, cancelToken:=cancelToken, progress:=progress)
                    If Not excelResult.Success OrElse (cancelToken <> CancellationToken.None AndAlso cancelToken.IsCancellationRequested) Then GoTo ErrorSkip
                End If

                '/loop through FNDs, open, input reactions & run macros
                '//Run Unit Base
                If Me.UnitBases.Count > 0 Then
                    Await WriteLineLogLine("INFO | " & Me.UnitBases.Count & " Unit Bases found..", progress)
                    'For Each unitbase In Me.UnitBases
                    For i As Integer = 0 To UnitBases.Count - 1
                        excelResult = Await OpenExcelRunMacroAsync(UnitBases(i), unitBaseMac, xlVisibility, cancelToken:=cancelToken, progress:=progress)
                        If Not excelResult.Success OrElse (cancelToken <> CancellationToken.None AndAlso cancelToken.IsCancellationRequested) Then GoTo ErrorSkip
                    Next
                End If
                '//Run Drilled Pier
                If Me.DrilledPierTools.Count > 0 Then
                    Await WriteLineLogLine("INFO | " & Me.DrilledPierTools.Count & " Drilled Piers found..", progress)
                    'For Each drilledPier In Me.DrilledPierTools
                    For i As Integer = 0 To DrilledPierTools.Count - 1
                        excelResult = Await OpenExcelRunMacroAsync(DrilledPierTools(i), drilledPierMac, xlVisibility, cancelToken:=cancelToken, progress:=progress)
                        If Not excelResult.Success OrElse (cancelToken <> CancellationToken.None AndAlso cancelToken.IsCancellationRequested) Then GoTo ErrorSkip
                    Next
                End If
                '//Run Pad/Pier
                If Me.PierandPads.Count > 0 Then
                    Await WriteLineLogLine("INFO | " & Me.PierandPads.Count & " Pier and Pads found..", progress)
                    'For Each pierAndPad In Me.PierandPads
                    For i As Integer = 0 To PierandPads.Count - 1
                        excelResult = Await OpenExcelRunMacroAsync(PierandPads(i), pierPadMac, xlVisibility, cancelToken:=cancelToken, progress:=progress)
                        If Not excelResult.Success OrElse (cancelToken <> CancellationToken.None AndAlso cancelToken.IsCancellationRequested) Then GoTo ErrorSkip
                    Next
                End If
                '//Run Pile
                If Me.Piles.Count > 0 Then
                    Await WriteLineLogLine("INFO | " & Me.Piles.Count & " Piles found..", progress)
                    'For Each pile In Me.Piles
                    For i As Integer = 0 To Piles.Count - 1
                        excelResult = Await OpenExcelRunMacroAsync(Piles(i), pileMac, xlVisibility, cancelToken:=cancelToken, progress:=progress)
                        If Not excelResult.Success OrElse (cancelToken <> CancellationToken.None AndAlso cancelToken.IsCancellationRequested) Then GoTo ErrorSkip
                    Next
                End If
                '//Run Guy Anchor
                If Me.GuyAnchorBlockTools.Count > 0 Then
                    Await WriteLineLogLine("INFO | " & Me.GuyAnchorBlockTools.Count & " Guy Anchors found..", progress)
                    'For Each guyAnchor In Me.GuyAnchorBlockTools
                    For i As Integer = 0 To GuyAnchorBlockTools.Count - 1
                        excelResult = Await OpenExcelRunMacroAsync(GuyAnchorBlockTools(i), guyAnchorMac, xlVisibility, cancelToken:=cancelToken, progress:=progress)
                        If Not excelResult.Success OrElse (cancelToken <> CancellationToken.None AndAlso cancelToken.IsCancellationRequested) Then GoTo ErrorSkip
                    Next
                End If

            Case Else
                'manual process
                Await WriteLineLogLine("WARNING | Manual process due to Tower Type: " & strType, progress)
                GoTo ErrorSkip
        End Select

        Await WriteLineLogLine("INFO | Maestro Complete", progress)
        AnalysisRunning = False
        Return True

ErrorSkip:
        If cancelToken <> CancellationToken.None AndAlso cancelToken.IsCancellationRequested Then
            Await WriteLineLogLine("WARNING | Maestro Cancelled", progress)
        ElseIf Not excelResult.Success AndAlso String.IsNullOrEmpty(excelResult.MacroName) Then
            Await WriteLineLogLine(excelResult.ErrorMsg, progress)
        End If
        AnalysisRunning = False
        Return False

    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="isDevMode">Determines if the beta version of TNX should be used.</param>
    ''' <param name="xlVisibility">Determines excel tools appear in the foreground (true) or are hidden (false).</param>
    ''' <param name="cancelToken">Provide an optional cancellation token source for the conduct task. If blank, the structures own cancellation token source will be used.</param>
    ''' <returns></returns>
    'Public Async Function ConductAsync(Optional isDevMode As Boolean = False, Optional ByVal xlVisibility As Boolean = False, Optional cancelTokenSource As CancellationTokenSource = Nothing) As Task(Of Boolean)
    '    _Conducting = True
    '    Me.ConductCancelTokenSource = If(cancelTokenSource, New CancellationTokenSource())
    '    Return Await Task.Run(Function() Conduct(isDevMode, xlVisibility), Me.ConductCancelTokenSource.Token)
    '    _Conducting = False
    'End Function

    'Async Cancellation
    'Private _Conducting
    'Public ReadOnly Conducting As Boolean
    'Private ConductCancelTokenSource As CancellationTokenSource


    'Public Sub CancelConduct()
    '    If Conducting AndAlso ConductCancelTokenSource IsNot Nothing Then
    '        ConductCancelTokenSource.Cancel()
    '        _Conducting = False
    '    End If
    'End Sub

    'Public Async Function ConductAsync(cancelToken As CancellationToken, Optional isDevMode As Boolean = False, Optional ByVal xlVisibility As Boolean = False) As Task(Of Boolean)

    '    Return Await Task.Run(Function() Conduct(isDevMode, xlVisibility), cancelToken)
    'End Function

    ''' <summary>
    ''' checks to see if the Excel macro returned success or not
    ''' Returns false if failed
    ''' </summary>
    ''' <param name="result"></param>
    ''' <param name="toolName"></param>
    ''' <returns></returns>

    Public Async Function DoBARB(ByVal poleWeUsin As Pole, ByVal plateWeUsin As CCIplate, Optional ByVal isDevMode As Boolean = False,
                                      Optional cancelToken As CancellationToken = Nothing, Optional progress As IProgress(Of LogMessage) = Nothing) As Task(Of Boolean)

        Dim barbCL As Double
        Dim plateComp As Double
        Dim plateSheer As Double
        Dim plateMom As Double

        'Dim plateWeUsin As CCIplate = Nothing
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
                Await WriteLineLogLine("INFO | Baseplate Bolt Group found for BARB..", progress)

                Exit For
            End If
        Next

        If applyBarb And Not IsNothing(basePlateBoltGroup) Then
            Await WriteLineLogLine("INFO | Getting reactions for BARB..", progress)
            barbCL = plateWeUsin.barb_cl_elevation 'exStruct.plate.barbCL

            Dim getReations As Tuple(Of Boolean, Double?, Double?, Double?) = Await GetReactionsBARB(basePlateConnection, plateMom, plateComp, plateSheer, cancelToken, progress)
            '''Item 1 = Bool
            '''Item 2 = mom
            '''Item 3 = comp
            '''item 4 = sheer
            If getReations.Item1 Then
                plateMom = getReations.Item2
                plateComp = getReations.Item3
                plateSheer = getReations.Item4
                'replace values in Pole
                If Not Await BarbValuesIntoPole(poleWeUsin, poleWeUsin.WorkBookPath, barbCL, plateComp, plateSheer, plateMom, isDevMode, cancelToken, progress) Then
                    Return False
                End If
            Else
                Return False
            End If
            'Run TNX Reactions Macro in Pole
        ElseIf applyBarb = False Then
            Await WriteLineLogLine("INFO | Apply Barb = False", progress)
            Return False
        Else
            Await WriteLineLogLine("ERROR | No Plate group found..", progress)
            Return False
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

    Public Async Function OpenExcelRunMacroAsync(ByVal objectTorun As EDSExcelObject, ByVal bigMac As String,
                                      Optional ByVal xlVisibility As Boolean = False, Optional toolName As String = Nothing,
                                      Optional cancelToken As CancellationToken = Nothing, Optional progress As IProgress(Of LogMessage) = Nothing) As Task(Of OpenExcelRunMacroResult)


        Dim tnxFilePath As String = Me.tnx.filePath
        Dim excelPath As String = objectTorun.WorkBookPath
        Dim toolFileName As String = Path.GetFileName(excelPath)

        Dim logString As String = ""
        Dim Result As OpenExcelRunMacroResult = New OpenExcelRunMacroResult(False, False, If(toolName, objectTorun.EDSObjectName))

        If String.IsNullOrEmpty(excelPath) Or String.IsNullOrEmpty(bigMac) Then
            Await WriteLineLogLine("ERROR | excelPath or bigMac parameter is null or empty", progress)
            Return Result
        End If

        Dim xlApp As Excel.Application = Nothing
        Dim xlWorkBook As Excel.Workbook = Nothing

        Dim errorMessage As String = ""
        Dim excelClose As Boolean = True
        Dim macroRan As Boolean = True
        Dim macroErrorMessage As String = ""

        Try
            If Not File.Exists(excelPath) Then
                errorMessage = $"ERROR | {excelPath} path not found!"
                Await WriteLineLogLine(errorMessage, progress)
                Return Result
            End If

            xlApp = CreateObject("Excel.Application")
            xlApp.Visible = xlVisibility

            xlWorkBook = xlApp.Workbooks.Open(excelPath)

            'check for pole and plate
            If Me.ParentStructure.CCIplates.Count > 0 Then
                If objectTorun.EDSObjectName.ToUpper = "CCIPOLE" And IsSomething(Me.ParentStructure.CCIplates(0).Connections) Then
                    If toolName.ToLower().Contains("step 1") Then
                        'insert baseplate grade into pole from plate
                        Await WriteLineLogLine("INFO | Inserting Baseplate grade into CCIPole from CCIPlate", progress)

                        Await AssignBasePlateGrade(xlWorkBook, cancelToken, progress)
                    End If

                End If
            End If
            Await WriteLineLogLine("INFO | Tool: " & toolFileName, progress)
            Await WriteLineLogLine("INFO | Running macro: " & bigMac, progress)

            If Not IsNothing(tnxFilePath) Then
                logString = xlApp.Run(bigMac, tnxFilePath)
                Await WriteLineLogLine("INFO | Macro result: " & vbCrLf & logString.Trim, progress)
            Else
                Await WriteLineLogLine("WARNING | No TNX file path in structure..", progress)
                logString = xlApp.Run(bigMac)
                Await WriteLineLogLine("INFO | Macro result: " & vbCrLf & logString.Trim, progress)
            End If

            xlWorkBook.Save()

        Catch ex As Exception
            macroRan = False
            macroErrorMessage = ex.Message
        Finally
            Try
                If xlWorkBook IsNot Nothing Then
                    xlWorkBook.Close(True)
                    Marshal.ReleaseComObject(xlWorkBook)
                    xlWorkBook = Nothing
                End If
                If xlApp IsNot Nothing Then
                    xlApp.Quit()
                    Marshal.ReleaseComObject(xlApp)
                    xlApp = Nothing
                End If
            Catch ex As Exception
                excelClose = False
            End Try
        End Try

        If Not macroRan Then
            Await WriteLineLogLine(macroErrorMessage, progress)
            If Not excelClose Then
                Await WriteLineLogLine("WARNING | Could not close Excel file or App", progress)
            End If

            Return Result
        End If

        If Not excelClose Then
            Await WriteLineLogLine("WARNING | Could not close Excel file or App", progress)
        End If

        'check for errors returned from Excel
        If logString.Contains("| ERROR |") Then
            Await WriteLineLogLine(Result.ErrorMsg, progress)
            Return Result
        End If

        'reload structure object
        Dim structureReload As Boolean = True
        Dim strucutreReloadError As String = ""
        Try
            objectTorun.Clear()
            objectTorun.LoadFromExcel()
        Catch ex As Exception
            structureReload = False
            strucutreReloadError = ex.Message
        End Try

        If Not structureReload Then
            Await WriteLineLogLine("ERROR | Could not rebuild structure object! " & strucutreReloadError, progress)

            Return Result
        End If

        'check for seismic
        If objectTorun.EDSObjectName.ToUpper = "CCISEISMIC" And logString.ToUpper.Contains("SEISMIC ANALYSIS REQUIRED") Then
            Result.SeismicRequired = True
        End If

        Result.Success = True
        Return Result
    End Function

    Public Async Function AssignBasePlateGrade(ByVal xlWorkBook As Excel.Workbook,
                                      Optional cancelToken As CancellationToken = Nothing, Optional progress As IProgress(Of LogMessage) = Nothing) As Task(Of Boolean)

        Dim plateGradeAssigned As Boolean = True
        Dim plateGradeAssignedError As String = ""
        Dim pole_flange_fy As Double = 0
        Try
            With xlWorkBook
                Dim row As Integer = 74
                .Worksheets("Macro References").Range("I74:J83").ClearContents
                For Each conn As Connection In Me.ParentStructure.CCIplates(0).Connections 'Does this need to loop through all potential CCIplate files? - MRR
                    If Not IsNothing(conn.connection_elevation) Then
                        For Each plate As PlateDetail In conn.PlateDetails
                            If plate.plate_type <> "Interior" Then
                                For Each matl As CCIplateMaterial In plate.CCIplateMaterials
                                    If matl.ID = plate.plate_material And matl.fy_0 > pole_flange_fy Then
                                        pole_flange_fy = CType(matl.fy_0, Double)
                                    End If
                                Next
                            End If
                        Next
                        .Worksheets("Macro References").Range("I" & row).Value = CType(conn.connection_elevation, Double)
                        .Worksheets("Macro References").Range("J" & row).Value = pole_flange_fy
                        row += 1
                        pole_flange_fy = 0
                    End If
                Next
            End With
        Catch ex As Exception
            plateGradeAssigned = False
            plateGradeAssignedError = ex.Message
        End Try

        If Not plateGradeAssigned Then
            Await WriteLineLogLine("WARNING | Could not determine baseplate grade from CCIPlate. Assumed grades may be used.", progress)
            Await WriteLineLogLine("WARNING | " & plateGradeAssignedError, progress)
            Return False
        End If

        If pole_flange_fy > 0 Then Await WriteLineLogLine("DEBUG | Baseplate grade found: " & pole_flange_fy.ToString(), progress)
        Return True
    End Function


    Public Async Function BarbValuesIntoPole(pole As Pole, excelPath As String, barbCL As Double, plateComp As Double,
                                      plateShear As Double, plateMom As Double, Optional isDevEnv As Boolean = False,
                                      Optional cancelToken As CancellationToken = Nothing, Optional progress As IProgress(Of LogMessage) = Nothing) As Task(Of Boolean)
        Dim xlApp As Excel.Application
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

        Dim BarbValueSet As Boolean = True
        Dim barbValueSetError As String = ""
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
                            'secID = sec.ID

                            If Not IsNothing(sec.local_section_id) Then
                                secID = sec.local_section_id
                            Else
                                Continue For
                            End If
                            rowNum = secID + 4

                            'Comp/Pu = column I
                            cellRangeComp = "I" & rowNum
                            xlWorkSheet.Range(cellRangeComp).Value = plateComp

                            'Moment/Mux= column J
                            cellRangeMom = "J" & rowNum
                            xlWorkSheet.Range(cellRangeMom).Value = plateMom

                            'Shear/Vu = column K
                            cellRangeShear = "K" & rowNum
                            xlWorkSheet.Range(cellRangeShear).Value = plateShear

                        End If
                    Next
                End If

                'Dim poleOutputFromBarb As String

                'poleOutputFromBarb = xlApp.Run(poleMac)

                'WriteLineLogLine("INFO | " & poleOutputFromBarb)

                xlWorkBook.Save()
                xlWorkBook.Close()
                xlApp.Quit()
            Else
                Await WriteLineLogLine("WARNING | " & excelPath & " CCIPole path for BARB not found!", progress)
                Return False
            End If
        Catch ex As Exception
            BarbValueSet = False
            barbValueSetError = ex.Message
        End Try

        If Not BarbValueSet Then
            Await WriteLineLogLine("ERROR | Error putting BARB values into CCIPole" & barbValueSetError, progress)
            Return False
        End If

        Return True
    End Function

    Public Async Function SpliceCheck(workingAreaPath As String,
                                      Optional cancelToken As CancellationToken = Nothing, Optional progress As IProgress(Of LogMessage) = Nothing) As Task(Of String)
        Dim repoInfo As New tnxCCIReport(Me)
        Dim twrType As String = Me.tnx.geometry.AntennaType.ToUpper 'exStruct.tnx.geometry.upperStructure.ToString.ToUpper
        Dim spliceFile As String
        Dim workingSpliceFile As String
        'exStruct.tnx.geometry.AntennaType

        'exStruct.tnx.geometry.baseStructure.Count = 0
        'exStruct.tnx.geometry.upperStructure.Contains("pole")

        If repoInfo.sReportTowerManufacturer.ToUpper = "PIROD" And twrType = "TAPERED POLE" Then
            'copy splice tool
            'run macro to import TNX
            'run the "run" macro

            spliceFile = Await FindSpliceTool(cancelToken:=cancelToken, progress:=progress)

            If Not spliceFile = "" Then
                'copy tool
                workingSpliceFile = Await CopyFile(spliceFile, workingAreaPath, cancelToken, progress)

                Return workingSpliceFile
            End If
        End If

        Return ""
    End Function

    Public Async Function RunTNXAsync(tnxFilePath As String, Optional isDevMode As Boolean = False, Optional cancelToken As CancellationToken = Nothing, Optional progress As IProgress(Of LogMessage) = Nothing) As Task(Of Boolean)
        'tnx vars
        Dim tnxAppLocation As String = "C:\Program Files (x86)\TNX\tnxTower 8.1.5.0 BETA\tnxtower.exe"
        Dim tnxLogFilePath As String = tnxFilePath & ".APIRun.log"
        Dim tnxProdVersion As String = ""
        Dim tnxSuccess As Boolean = False

        Dim twrType As String = Me.SiteInfo.tower_type.ToUpper
        Dim timeOutCounter As Integer 'set in milliseconds

        If Not File.Exists(tnxFilePath) Then
            Await WriteLineLogLine("ERROR | .eri file does not exist: " & tnxFilePath, progress)
            Return False
        End If

        Dim tnxRan As Boolean = True
        Dim tnxRanError As String = ""
        Try
            Dim cmdProcess As New Process


            Await WriteLineLogLine("INFO | Running TNX..", progress)

            'determine TNX File path - newest version
            Dim whereBeTNX As Tuple(Of String, String) = Await WhereInTheWorldIsTNXTower(tnxProdVersion, isDevMode, cancelToken, progress)
            tnxAppLocation = whereBeTNX.Item1
            tnxProdVersion = whereBeTNX.Item2

            If tnxAppLocation = "" Then
                'TNX app not found
                Await WriteLineLogLine("ERROR | TNX Not installed! Cannot proceed.", progress)
                Return False
            End If

            'see whether arch notation is used or not
            If Await ArchNotReg(cancelToken, progress) Then
                Await WriteLineLogLine("WARNING | Architectural notation is set to 'Yes'", progress)
            Else
                Await WriteLineLogLine("INFO | Architectural notation is set to 'No'", progress)
            End If

            Dim tnxVersionGot As Boolean = True
            Try
                Me.tnx.settings.projectInfo.VersionUsed = tnxProdVersion
            Catch ex As Exception
                tnxVersionGot = False
            End Try

            If Not tnxVersionGot Then
                Await WriteLineLogLine("WARNING | Could not set TNX object version. EDS may show older TNX version number.", progress)
            End If

            Dim tnxLogDeleted As Boolean = True
            Try
                'delete tnx log file if it exist
                If File.Exists(tnxLogFilePath) Then
                    File.Delete(tnxLogFilePath)
                End If
            Catch ex As Exception
                tnxLogDeleted = False
            End Try

            If Not tnxLogDeleted Then
                Await WriteLineLogLine("ERROR | Could not delete TNX API log file. Please delete before continuing: " & tnxLogFilePath, progress)
                Return False
            End If
            'Need to determine if word is open prior to running TNX
            'If it is open then it shouldn't be killed when closing the RTF
            'If it isn't open before tnx then it should be killed
            'Dim isWordOpen As Boolean
            'Try
            '    Dim word As Object = GetObject(, "Word.Application")
            '    isWordOpen = True
            'Catch ex As Exception
            '    isWordOpen = False
            'End Try

            'Make sure ReportPrintReactions=Yes in eri file
            If Not Await SetEriOutputVariables(tnxFilePath, cancelToken, progress) Then
                Await WriteLineLogLine("WARNING | Could not verify ReportPrintReactions=Yes in ERI output variables", progress)
            End If

            'seting timout duration - longer for guyeds
            If twrType = "GUYED" Then
                timeOutCounter = 600000
            Else
                timeOutCounter = 300000
            End If

            With cmdProcess
                If Version821OrLater(tnxProdVersion) Then
                    .StartInfo = New ProcessStartInfo(tnxAppLocation, Chr(34) & tnxFilePath & Chr(34) & " RunAnalysis SilentAnalysisRun GenerateDesignReport GenerateE1PDF SaveInputFileOnCompletion") 'GenerateCCIReport 'RunAnalysis 'SilentAnalysisRun 'GenerateDesignReport
                Else
                    .StartInfo = New ProcessStartInfo(tnxAppLocation, Chr(34) & tnxFilePath & Chr(34) & " RunAnalysis SilentAnalysisRun GenerateDesignReport") 'GenerateCCIReport 'RunAnalysis 'SilentAnalysisRun 'GenerateDesignReport
                End If
                With .StartInfo
                    .CreateNoWindow = True
                    .UseShellExecute = False
                    .RedirectStandardOutput = True
                End With
                .Start()


                tnxSuccess = Await CheckLogFileForFinishedAsync(tnxLogFilePath, timeOutCounter, True, cancelToken, progress)
                Dim tnxDead As Boolean = True
                Dim tnxDeadError As String = ""
                Try
                    Await WriteLineLogLine("INFO | TNX finished, attempting to terminate..", progress)
                    .Kill()
                    Await WriteLineLogLine("INFO | TNX termination complete..", progress)
                Catch ex As Exception
                    tnxDead = False
                    tnxDeadError = ex.Message
                Finally
                    'Try
                    '    'For the time being the RTF file still opens 
                    '    'This needs to be closed before returning TRUE
                    '    CloseRTF(tnxFilePath, isWordOpen)
                    'Catch ex As Exception
                    '    WriteLineLogLine("WARNING | Could not close RFT file: " & ex.Message)
                    'End Try
                End Try

                If Not tnxDead Then
                    Await WriteLineLogLine("WARNING | Exception closing TNX - check and close via task manager: " & tnxDeadError, progress)
                End If

                '.WaitForInputIdle()
                '.WaitForExit()
            End With ' Read output to a string variable.

            Dim ipconfigOutput As String = cmdProcess.StandardOutput.ReadToEnd


            'WriteLineLogLine("INFO | " & ipconfigOutput)
            'For the time being the RTF file still opens 
            'This needs to be closed before returning TRUE
            'I put this around a try catch in cases where an ERI is run and files already exist. 

            If Not tnxSuccess Then Return False Else Return True

        Catch ex As Exception
            tnxRan = False
            tnxRanError = ex.Message
        End Try




        If Not tnxRan Then
            Await WriteLineLogLine("ERROR | Exception Running TNX: " & tnxRanError, progress)
            Return False
        End If
    End Function
    ''' <summary>
    ''' Checks registry settings to see if Architectural Notation is set
    ''' Returns True if Yes
    ''' </summary>
    ''' <returns></returns>
    Public Async Function ArchNotReg(Optional cancelToken As CancellationToken = Nothing, Optional progress As IProgress(Of LogMessage) = Nothing) As Task(Of Boolean)
        Dim tnxRegPath As String = "HKEY_CURRENT_USER\SOFTWARE\TNX\tnxTower\US Units"

        Dim archNot As Object
        Dim archBool As Boolean = True
        Dim archBoolError As String = ""
        Try
            archNot = My.Computer.Registry.GetValue(tnxRegPath, "Architectural", Nothing)
        Catch ex As Exception
            archBool = False
            archBoolError = ex.Message
        End Try

        If Not archBool Then
            Await WriteLineLogLine("ERROR | Exception getting Archetectural Notation setting: " & archBoolError, progress)

            Return False
        End If

        If archNot.toupper.contains("Y") Then
            Return True
        ElseIf archNot.toupper.contains("N") Then
            Return False
        Else
            Await WriteLineLogLine("WARNING | Could not determine Architectural Notation value: " & archNot.ToString, progress)
            Return False
        End If

    End Function

    ''' <summary>
    ''' Sets ReportPrintReactions=Yes in eri file prior to running
    ''' </summary>
    ''' <param name="tnxFilePath"></param>
    ''' <returns></returns>
    Public Async Function SetEriOutputVariables(tnxFilePath As String, Optional cancelToken As CancellationToken = Nothing, Optional progress As IProgress(Of LogMessage) = Nothing) As Task(Of Boolean)
        Dim eriAllText As String

        If Not File.Exists(tnxFilePath) Then Return False
        Dim varsSet As Boolean = True
        Dim varsSetError As String = ""
        Try

            Using fs As FileStream = New FileStream(tnxFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Dim replaceArray() As String = {"ReportPrintReactions=No|ReportPrintReactions=Yes",
                "ReportMaxForces=No|ReportMaxForces=Yes",
                "ReportPrintReactions=No|ReportPrintReactions=Yes",
                "ReportPrintDeflection=No|ReportPrintDeflection=Yes",
                "ReportPrintBoltChecks=No|ReportPrintBoltChecks=Yes",
                "ReportPrintStressChecks=No|ReportPrintStressChecks=Yes",
                "ReportInputGeometry=No|ReportInputGeometry=Yes",
                "ReportInputOptions=No|ReportInputOptions=Yes",
                "CapacityReportOutputType=No Capacity Output|CapacityReportOutputType=Capacity Summary",
                "CapacityReportOutputType=Capacity Details|CapacityReportOutputType=Capacity Summary",
                "PrintInputGeometry=No|PrintInputGeometry=Yes"}

                Using r As StreamReader = New StreamReader(fs)
                    'make sure we're at the beginning
                    r.DiscardBufferedData()
                    r.BaseStream.Seek(0, SeekOrigin.Begin)
                    eriAllText = r.ReadToEnd

                    For Each setting As String In replaceArray
                        Dim settingSplt() As String = setting.Split("|")
                        eriAllText = eriAllText.Replace(settingSplt(0), settingSplt(1))
                    Next


                    Using w As StreamWriter = New StreamWriter(tnxFilePath, False)
                        w.Write(eriAllText)
                    End Using
                End Using
            End Using
        Catch ex As Exception
            varsSet = False
            varsSetError = ex.Message
        End Try

        If Not varsSet Then
            Await WriteLineLogLine("ERROR | Exception setting ERI output variables: " & varsSetError, progress)
            Return False
        End If

        Return True
    End Function

    'Close an RTF file based on a tnx file
    'If isWordOpen is true then it will not close word
    Public Sub CloseRTF(ByVal tnxfilepath As String, ByVal iswordOpen As Boolean)
        Dim word As Object
        Dim doc As Object
        Dim rtfPath As String = tnxfilepath & ".rtf"

        Dim file As New FileInfo(rtfPath)
        Dim wordCheck As Boolean = False
        Dim filecheck As Boolean = False
        If file.Exists Then
            While wordCheck = False
                Try
                    word = GetObject(, "Word.Application")
                    wordCheck = True

                    While filecheck = False
                        filecheck = FileIsOpen(file)
                    End While

                    doc = word.Documents(rtfPath)
                    doc.close
                Catch ex As Exception

                End Try
            End While
        End If
        'Thread.Sleep(4000)
        'word = GetObject(, "Word.Application")

        If Not iswordOpen Then word.quit
        word = Nothing
        doc = Nothing

    End Sub

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
    ''' <summary>
    '''check the TNX API silent log to determine when it's finished
    '''maxTimeout is in milliseconds
    ''' </summary>
    ''' <param name="logFilePath"></param>
    ''' <param name="maxTimeout"></param>
    Private Async Function CheckLogFileForFinishedAsync(logFilePath As String, maxTimeout As Integer, generateReport As Boolean, Optional cancelToken As CancellationToken = Nothing, Optional progress As IProgress(Of LogMessage) = Nothing) As Task(Of Boolean)

        ' Set the time interval to check the log file
        Dim checkInterval As Integer = 2000 ' 2 seconds
        ' Set the phrase to look for in the log file
        Dim finishedPhrase As String = "DESIGN END"
        'Set phrase to look for in the log file in case of error
        Dim errorPhrase As String = "ANALYSIS INCOMPLETE"
        ' Set the initial time
        Dim startTime As DateTime = DateTime.Now

        ' Create a FileStream and StreamReader to read the log file
        'we're using a filestream to (hopefully) be able to read the file without preventing TNX from writing to it
        Dim fs As FileStream
        Dim logReader As StreamReader

        If generateReport Then
            finishedPhrase = "CROWN CASTLE REPORT END" '"ANALYSIS AND DESIGN REPORT END"
        Else
            finishedPhrase = "DESIGN END"
        End If
        While True
            If File.Exists(logFilePath) Then
                fs = New FileStream(logFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                logReader = New StreamReader(fs)
                Exit While
            Else
                ' Check if the maximum timeout has been reached
                If (DateTime.Now - startTime).TotalMilliseconds > maxTimeout Then
                    Await WriteLineLogLine("ERROR | Checking TNX log exceeded timeout - could not find log file: " & logFilePath, progress)
                    Return False
                End If
                ' Wait for the check interval before checking again
                'Thread.Sleep(checkInterval)
                Await Task.Delay(checkInterval)
            End If
        End While

        Dim tnxTimeoutBool As Boolean = False
        Dim tnxErrorBool As Boolean = False
        Dim tnxError As String = ""
        Try
            ' Loop until the "Finished" line is found or the maximum timeout is reached
            While True
                ' Read last line from the log file
                Dim line As String = logReader.ReadToEnd()
                ' If the line is null, wait for the check interval and continue
                If line Is Nothing Then
                    Await Task.Delay(checkInterval)
                Else
                    ' If the line contains "Finished", exit the loop
                    If line.ToUpper.Contains(finishedPhrase) Then
                        Exit While
                    ElseIf line.ToUpper.Contains(errorPhrase) Then
                        'found an error in TNX - abort
                        tnxErrorBool = True
                        Exit While
                    End If
                End If
                ' Check if the maximum timeout has been reached
                If (DateTime.Now - startTime).TotalMilliseconds > maxTimeout Then
                    tnxTimeoutBool = True
                    Exit While
                End If
                ' Wait for the check interval before checking again
                Await Task.Delay(checkInterval)
            End While

            ' Close the StreamReader
            fs.Close()
            logReader.Close()

            ' Check if the "Finished" line was found or if the maximum timeout was reached
            If tnxTimeoutBool Then '(DateTime.Now - startTime).TotalMilliseconds > maxTimeout Then
                Await WriteLineLogLine("ERROR | Checking TNX log exceeded timeout", progress)
                Return False
            ElseIf tnxErrorBool Then
                Await WriteLineLogLine("ERROR | Error found in TNX log - Aborting Maestro", progress)
                Return False
            Else
                Await WriteLineLogLine("INFO | TNX API Finished..", progress)
                Return True
            End If
        Catch ex As Exception
            'tnxTimeoutBool = False
            tnxError = ex.Message
        End Try

        'if we get here, there was an exception
        'If Not tnxTimeoutBool Then
        Await WriteLineLogLine("ERROR | Exception checking TNX Log: " & tnxError, progress)
        Return False
        ' End If
    End Function

    Public Async Function WhereInTheWorldIsTNXTower(ByVal tnxVersion As String, Optional isDevMode As Boolean = False, Optional cancelToken As CancellationToken = Nothing, Optional progress As IProgress(Of LogMessage) = Nothing) As Task(Of Tuple(Of String, String))
        Dim defaultAppLocationBase As String = "C:\Program Files (x86)\TNX"
        Dim appName As String = "tnxTower"
        Dim newestAppFolderName As String

        Dim newestVersion As Version = Nothing
        Dim newestFolder As String = Nothing

        Dim appLoc As String

        Dim gottnx As Boolean = True
        Dim gottnxError As String = ""
        Try
            If Not Directory.Exists(defaultAppLocationBase) Then
                'TNX not installed
                Return New Tuple(Of String, String)("", "")
            End If

            For Each folder In Directory.GetDirectories(defaultAppLocationBase)

                Dim folderName As String = New DirectoryInfo(folder).Name.Replace(" BETA", "")
                Dim folderVersion As Version = Nothing

                'skip if beta version is found and we're not in devmode
                If (Not isDevMode And folder.ToUpper.Contains("BETA")) Or folder.ToUpper.Contains("ARCHIVE") Then
                    Continue For
                End If

                'if we get here and folder has beta, strip "BETA" off and readd if necessary
                'If folderName.ToUpper.Contains("BETA") Then
                '    folderName = folderName.ToUpper.Replace(" BETA", "")
                'End If

                If Version.TryParse(folderName.Substring(appName.Length + 1), folderVersion) Then
                    If newestVersion Is Nothing OrElse folderVersion > newestVersion Then
                        newestVersion = folderVersion
                        newestFolder = folder
                    End If
                End If
            Next

            newestAppFolderName = New DirectoryInfo(newestFolder).Name

            Await WriteLineLogLine("INFO | Newest TNX Version found: " & newestAppFolderName, progress)
            appLoc = Path.Combine(newestFolder, appName & ".exe")

            tnxVersion = FileVersionInfo.GetVersionInfo(appLoc).FileVersion

            Return New Tuple(Of String, String)(appLoc, tnxVersion)
        Catch ex As Exception
            gottnx = False
            gottnxError = ex.Message
        End Try

        If Not gottnx Then
            Await WriteLineLogLine("ERROR | Exception finding TNX App: " & gottnxError, progress)
            Return New Tuple(Of String, String)("", "")
        End If
    End Function

    'Written by Bard
    Public Function Version821OrLater(ByVal version As String) As Boolean
        ' Check if the version string has the correct format.
        If Not version.Contains(".") Then
            Exit Function
        End If

        version = version.Replace(" ", "").Replace("BETA", "")

        ' Split the version string into its components.
        Dim versionParts As String() = version.Split(".")

        ' If the version string has 4 components, then check the fourth component as well.
        If versionParts.Length = 4 Then
            ' Return true if all four components are equal to or greater than 8, 2, 1, and 0, respectively.
            Return (Integer.Parse(versionParts(0)) >= 8 And
                Integer.Parse(versionParts(1)) >= 2 And
                Integer.Parse(versionParts(2)) >= 1 And
                Integer.Parse(versionParts(3)) >= 0)
        Else
            ' The version string only has 3 components, so only check the first three.
            ' Return true if all three components are equal to or greater than 8, 2, and 1, respectively.
            Return (Integer.Parse(versionParts(0)) >= 8 And
                Integer.Parse(versionParts(1)) >= 2 And
                Integer.Parse(versionParts(2)) >= 1)
        End If
    End Function

    'copy file from original path to new path
    'pass original file path with name/extension and new path without file name
    Public Async Function CopyFile(origPath As String, newPath As String, Optional cancelToken As CancellationToken = Nothing, Optional progress As IProgress(Of LogMessage) = Nothing) As Task(Of String)
        Dim fileName As String = ""
        Dim newPathWithFileName As String = ""

        Dim gotFileName As Boolean = True
        Dim gotFileNameError As String = ""

        Try
            fileName = Path.GetFileName(origPath)
            newPathWithFileName = Path.Combine(newPath, fileName)

            If Not File.Exists(origPath) Then
                Throw New FileNotFoundException("ERROR | Original file not found: " + origPath)
            End If

            If Not Directory.Exists(newPath) Then
                Throw New DirectoryNotFoundException("ERROR | Destination directory not found: " + newPath)
            End If

            File.Copy(origPath, newPathWithFileName, True)

            Return newPathWithFileName

        Catch ex As Exception
            gotFileName = False
            gotFileNameError = ex.Message
        End Try

        If Not gotFileName Then
            Await WriteLineLogLine("ERROR | Error copying file '" & fileName & "': " & gotFileNameError, progress)
            Return ""
        End If
    End Function
    'assuming this has a pole passed in
    Public Async Function FindSpliceTool(Optional extension As String = "xls", Optional cancelToken As CancellationToken = Nothing, Optional progress As IProgress(Of LogMessage) = Nothing) As Task(Of String)
        Dim fileName As String = ""

        ' Check if the directory exists
        If Not Directory.Exists(SpliceCheckLocation) Then
            Await WriteLineLogLine("ERROR | Splice Check path not found!", progress)
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
        Await WriteLineLogLine("ERROR | Splice Check file not found!", progress)
        Return ""
    End Function
#End Region

#Region "Maestro Logging"
    Public Sub CreateLogFile()
        ' Get the current date and time
        Dim dt As String = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss tt")
        Dim splt() As String = dt.Split(" ")
        dt = splt(1) '& " " & splt(2)

        Dim msg As String = "INFO | Maestro Log [START] " & "Version: " & VerNum

        ' Print the message to the console
        Console.WriteLine(dt & " | " & msg)

        Try
            Using sw As New StreamWriter(LogPath, True)

                ' Write the log message to the file
                sw.WriteLine(dt & " | " & msg)
            End Using
        Catch ex As Exception
            Console.WriteLine("Error creating log file: " & ex.Message)

        End Try
    End Sub

    <DebuggerStepThrough()>
    Public Async Sub WriteLineLogLine(ByVal msg As String)
        ' Get the current date and time
        Dim dt As String = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss tt")
        Dim splt() As String = dt.Split(" ")
        dt = splt(1) '& " " & splt(2)
        ' Print the message to the console
        Console.WriteLine(dt & " | " & msg)

        ' Wrap the file operation in a try-catch block to handle exceptions
        Try
            ' If the log file does not exist, establish intro
            If Not File.Exists(LogPath) Then
                msg = "Maestro Log [START] " & "Version: " & VerNum & vbCrLf &
                      dt & " | " & msg
            End If
            ' Use a StreamWriter to write to the log file
            ' The 'True' argument appends to the file if it already exists
            Using sw As New StreamWriter(LogPath, True)
                ' Write the log message to the file
                sw.WriteLine(dt & " | " & msg)
            End Using
        Catch ex As Exception
            ' Handle the exception
            Console.WriteLine("Error writing to log file: " & ex.Message)
        End Try

    End Sub
    '<DebuggerStepThrough()>
    Public Async Function WriteLineLogLine(msg As String, progress As IProgress(Of LogMessage)) As Task

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

    Public Async Function GetReactionsBARB(ByVal con As Connection, ByVal mom As Double?, ByVal comp As Double?, ByVal sheer As Double?,
                                     Optional cancelToken As CancellationToken = Nothing, Optional progress As IProgress(Of LogMessage) = Nothing) As Task(Of Tuple(Of Boolean, Double?, Double?, Double?))
        '''Item 1 = Bool
        '''Item 2 = mom
        '''Item 3 = comp
        '''item 4 = sheer
        Dim plateComp As Double
        Dim plateSheer As Double
        Dim plateMom As Double
        Dim plateCompSeis As Double
        Dim plateSheerSeis As Double
        Dim plateMomSeis As Double

        'get BARB values from Plate
        Dim gotReactions As Boolean = True
        Dim gotReactionsError As String = ""
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

            Dim compareRatings As Tuple(Of Boolean, Double?, Double?, Double?) = Await CompareRatingsBARB(plateMom, plateComp, plateSheer, plateMomSeis, plateCompSeis, plateSheerSeis, mom, comp, sheer, cancelToken, progress)
            If Not compareRatings.Item1 Then
                Return New Tuple(Of Boolean, Double?, Double?, Double?)(False, 0, 0, 0)
            Else
                mom = compareRatings.Item2
                comp = compareRatings.Item3
                sheer = compareRatings.Item4
            End If

        Catch ex As Exception
            gotReactions = False
            gotReactionsError = ex.Message
        End Try

        If Not gotReactions Then
            Await WriteLineLogLine("ERROR | Exception finding BARB Ratings: " & gotReactionsError, progress)
            Return New Tuple(Of Boolean, Double?, Double?, Double?)(False, 0, 0, 0)
        End If

        Return New Tuple(Of Boolean, Double?, Double?, Double?)(True, mom, comp, sheer)
    End Function
    'figure out which results to return
    'The following combination of reactions may be provided: wind only, seismic only, wind & seismic.
    'Multiple CCIpole objects will also need to be compared as applicable.
    'Suggested: If seismic and wind exist for the BU, compare both for largest moment and override CCIpole with respective reactions. 

    Public Async Function CompareRatingsBARB(ByVal plateMom As Double?, ByVal plateComp As Double?, ByVal plateSheer As Double?,
                                       ByVal plateMomSeis As Double?, ByVal plateCompSeis As Double?, ByVal plateSheerSeis As Double?,
                                       ByVal momToUse As Double, ByVal compToUse As Double, ByVal sheerToUse As Double,
                                       Optional cancelToken As CancellationToken = Nothing, Optional progress As IProgress(Of LogMessage) = Nothing) As Task(Of Tuple(Of Boolean, Double?, Double?, Double?))
        '''Item 1 = Bool
        '''Item 2 = mom
        '''Item 3 = comp
        '''item 4 = sheer
        Dim compareWorked As Boolean = True
        Dim compareWorkedError As String = ""
        Try
            If IsSomething(plateMom) And
                IsSomething(plateComp) And
                IsSomething(plateSheer) And
                IsSomething(plateMomSeis) And
                IsSomething(plateCompSeis) And
                IsSomething(plateSheerSeis) Then
                'compare wind and seis to see which one is larger
                'only compare moment and use the larger set of values
                'return larger
                momToUse = Await GetLarger(plateMom, plateMomSeis, cancelToken, progress)

                If momToUse = plateMom Then
                    'use wind
                    compToUse = plateComp
                    sheerToUse = plateSheer
                    Await WriteLineLogLine("INFO | Using Wind reactions", progress)

                ElseIf momToUse = plateMomSeis Then
                    'use seismic
                    compToUse = plateCompSeis
                    sheerToUse = plateSheerSeis
                    Await WriteLineLogLine("INFO | Using Seismic reactions", progress)

                Else GoTo SkipCompare
                End If

                'compToUse = GetLarger(plateComp, plateCompSeis)
                'sheerToUse = GetLarger(plateSheer, plateSheerSeis)

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
                Await WriteLineLogLine("WARNING | No BARB values found!", progress)

            End If

            If momToUse = -1 Or compToUse = -1 Or sheerToUse = -1 Then
SkipCompare:
                Await WriteLineLogLine("ERROR | Could not compare  reaction values! Moment: " & momToUse & " Compression:  " & compToUse & " Sheer: " & sheerToUse & "", progress)
                Return New Tuple(Of Boolean, Double?, Double?, Double?)(False, 0, 0, 0)
            End If

        Catch ex As Exception
            compareWorked = False
            compareWorkedError = ex.Message
        End Try

        If Not compareWorked Then
            Await WriteLineLogLine("ERROR | Exception comparing BARB Ratings: " & compareWorkedError, progress)
            Return New Tuple(Of Boolean, Double?, Double?, Double?)(False, 0, 0, 0)
        End If

        Return New Tuple(Of Boolean, Double?, Double?, Double?)(True, momToUse, compToUse, sheerToUse)
    End Function
    'return larger of 2 values
    'return-1 if both vals are not numbers
    Public Async Function GetLarger(val1 As Double, val2 As Double, Optional cancelToken As CancellationToken = Nothing, Optional progress As IProgress(Of LogMessage) = Nothing) As Task(Of Double)
        If Double.IsNaN(val1) And Double.IsNaN(val2) Then
            Await WriteLineLogLine("ERROR | Null Value for comparison", progress)
            Return -1
        ElseIf Double.IsNaN(val1) Then
            Return val2
        ElseIf Double.IsNaN(val2) Then
            Return val1
        End If
        Return If(val1 > val2, val1, val2)

    End Function

    ''' <summary>
    ''' Loops through a set of eri files and runs TNX logic on them
    ''' pass in the parent directory. this folder should have a group of folders with an eri in each and no other files - delete generated files if rerun needed
    ''' </summary>
    ''' <param name="parentDirectory"></param>
    Public Sub LoopThroughERIFiles(parentDirectory As String)
        Dim str As New EDSStructure
        str.LogPath = "C:\Users\stanley\Crown Castle USA Inc\ECS - Tools\SAPI Test Cases\ERI Testing\ERI Log.txt"
        For Each fold In Directory.GetDirectories(parentDirectory)
            For Each f In Directory.GetFiles(fold)
                str.RunTNXAsync(f, True)
                Exit For
            Next
        Next

    End Sub

#End Region
End Class

Public Class OpenExcelRunMacroResult
    Public Property MacroName As String
    Public Property Success As Boolean = True
    Public Property SeismicRequired As Boolean = False

    Public ReadOnly Property ErrorMsg As String
        Get
            Return "ERROR | Exception running macro for " & MacroName & vbCrLf
        End Get
    End Property

    Public Sub New()
        'Default
    End Sub
    Public Sub New(success As Boolean, seismicRequired As Boolean, Optional name As String = "")
        Me.Success = success
        Me.SeismicRequired = seismicRequired
        Me.MacroName = name
    End Sub

End Class
