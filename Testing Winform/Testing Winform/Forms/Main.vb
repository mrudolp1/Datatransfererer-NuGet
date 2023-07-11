Imports CCI_Engineering_Templates
Imports System.IO
Imports RoboSharp
Imports System.Threading
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports DevExpress.XtraEditors
Imports SAPIReportGenerator
Imports SAPI_Report_Generator_Editor

Namespace UnitTesting

    Partial Public Class frmMain
        Public strcLocal As EDSStructure
        Public strcEDS As EDSStructure

        Public BUNumber As String = ""
        Public StrcID As String = ""
        Public WorkOrder As String = ""
        Public forceAcrchiving As Boolean = False

        'Import to Excel
        Public ListOfFilesCopied As New List(Of String) From {"C:\Users\" & Environment.UserName & "\Desktop\C Drive Testing\Drilled Pier\EDS\Test Sites\809534 - MP\Drilled Pier Foundation (5.1.0.3)_EDS_3.xlsm"}

        'Import to EDS
        Public ListOfExcelFiles As New List(Of String) From {"C:\Users\" & Environment.UserName & "\Desktop\C Drive Testing\Drilled Pier\EDS\Test Sites\809534 - MP\Drilled Pier Foundation (5.1.0.3)_2.xlsm"}

#Region "Form Handlers"
        Public Sub New()
            InitializeComponent()
        End Sub
        Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
            isopening = True

            StartEverything()

            If Environment.UserName.ToLower = "imiller" Or
               Environment.UserName.ToLower = "stanley" Or
               Environment.UserName.ToLower = "dsmilowitz" Or
               Environment.UserName.ToLower = "chall" Or
               Environment.UserName.ToLower = "mrudolph" Then
                If My.Settings.dbSelection = "DEV" Then
                    toggleDevUat.IsOn = False
                Else
                    toggleDevUat.IsOn = True
                End If
            Else
                toggleDevUat.IsOn = True
            End If

            'Set UI inputs to the values saved in the user settings
            testLocalWorkarea.Text = My.Settings.localWorkArea
            lFolder = My.Settings.localWorkArea
            CheckEditDevMode.Checked = My.Settings.booConductDevMode
            CheckEditExcelVisible.Checked = My.Settings.booConductExcelVis
            CheckEditExcelVisibleII.Checked = My.Settings.booImportInputsExcelVisible
            CheckEditAutoReport.Checked = My.Settings.booReportOption
            txtFndBU.Text = My.Settings.myBU
            txtFndStrc.Text = My.Settings.myStrID
            txtFndWO.Text = My.Settings.myWO
            txtDirectory.Text = My.Settings.myWorkArea

            If My.Settings.localWorkArea = String.Empty Then
                My.Settings.localWorkArea = "C:\Users\" & Environment.UserName & "\source"
                My.Settings.Save()
            End If

            'force local work area it to be true
            chkWorkLocal.Checked = True
            My.Settings.workLocal = True
            My.Settings.Save()
            If unitTestCases.Count = 0 Then
                LoadTestCases(unitTestCases)
            End If

            'Kill all the robocopies active (This can't be used along side the dashboard)
            KillRoboCops()
            isopening = False

            'Isopening is set to false here because All controls should do all actions based on the next if statement
            SetTestIDLabels()

            If My.Settings.MyTestCase > 0 Then
                If Directory.Exists(lFolder & "\Test ID " & My.Settings.MyTestCase) Then
                    testIteration.Text = currentTestingIteration
                    testID.SelectedIndex = My.Settings.MyTestCase - 1
                    SetUpWorkArea(My.Settings.MyTestCase - 1)

                    If testID.Text.Contains("Checked Out") Then
                        btnClose.Enabled = True
                        btnCheckout.Enabled = False
                        testPush.Enabled = True
                    Else
                        btnClose.Enabled = False
                        btnCheckout.Enabled = True
                        testPush.Enabled = False
                    End If

                End If
            End If
        End Sub

        Private Sub frmMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
            CloseHandle(EDStokenHandle)
            Try
                KillRoboCops()
                'DirectorySync.Stop()
            Catch ex As Exception
            End Try
        End Sub


#End Region

#Region "Structure"

        Private Sub btnLoopThroughERI_Click(sender As Object, e As EventArgs) Handles btnLoopThroughERI.Click
            Dim ed As New EDSStructure
            Dim pd As String = txtDirectory.Text
            ed.LoopThroughERIFiles(pd)
        End Sub


        Private Sub btnImportStrcFiles_Click(sender As Object, e As EventArgs) Handles btnImportStrcFiles.Click
            If txtFndBU.Text = "" Or txtFndStrc.Text = "" Then Exit Sub
            BUNumber = txtFndBU.Text
            StrcID = txtFndStrc.Text
            WorkOrder = txtFndWO.Text

            Dim xlFd As New OpenFileDialog
            ''xlFd.InitialDirectory = txtDirectory.Text
            xlFd.Multiselect = True
            'xlFd.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"

            If xlFd.ShowDialog = DialogResult.OK Then
                Dim workingDirectory As String = Path.GetDirectoryName(xlFd.FileNames(0))
                txtDirectory.Text = workingDirectory
                strcLocal = New EDSStructure(txtFndBU.Text, txtFndStrc.Text, txtFndWO.Text, workingDirectory, workingDirectory, xlFd.FileNames, EDSnewId, EDSdbActive)

            End If

            'Test Parents
            'For Each pp In strcLocal.PierandPads
            '    MessageBox.Show("My parent structure is " & pp.ParentStructure.ToString)
            'Next

            propgridFndXL.SelectedObject = strcLocal

        End Sub

        Private Sub btnExportStrcFiles_Click(sender As Object, e As EventArgs) Handles btnExportStrcFiles.Click
            If strcEDS Is Nothing Then Exit Sub


            strcEDS.SaveTools(txtDirectory.Text)

        End Sub
        Private Sub btnLoadStrcFromEDS_Click(sender As Object, e As EventArgs) Handles btnLoadFndFromEDS.Click
            If txtFndBU.Text = "" Or txtFndStrc.Text = "" Then Exit Sub
            'Go to the EDSFoundationGroup.LoadAllFoundationsFromEDS() and uncomment your foundation type when it's ready for testing.
            Dim workingDirectory As String = txtDirectory.Text
            If Not Directory.Exists(workingDirectory) Then
                MessageBox.Show("Working Directory Not Found.")
                Exit Sub
            End If

            strcEDS = New EDSStructure(txtFndBU.Text, txtFndStrc.Text, txtFndWO.Text, workingDirectory, workingDirectory, EDSnewId, EDSdbActive)

            propgridFndEDS.SelectedObject = strcEDS

        End Sub
        Private Sub btnSaveStrcToEDS_Click(sender As Object, e As EventArgs) Handles btnSaveFndToEDS.Click
            If strcLocal Is Nothing Or txtFndBU.Text = "" Or txtFndStrc.Text = "" Then Exit Sub
            'Go to the EDSFoundationGroup.SaveAllFoundationsFromEDS() and uncomment your foundation type when it's ready for testing.

            strcLocal.EDSMe = New EDSStructure(strcLocal.bus_unit, strcLocal.structure_id, strcLocal.work_order_seq_num, strcLocal.databaseIdentity, strcLocal.activeDatabase)

            Try
                My.Computer.Clipboard.SetText(strcLocal.SavetoEDSQuery)
            Catch ex As Exception
                Debug.WriteLine("Failed to copy query to clipboard.")
            End Try

            If MessageBox.Show("Structure query copied to clipboard. Would you like to send the structure to EDS?", "Save Structure to EDS?", MessageBoxButtons.YesNo) = vbYes Then
                Try
                    strcLocal.SavetoEDS()
                Catch ex As Exception
                    Debug.WriteLine("Failed to send sql query.")
                End Try
            End If

        End Sub
        Private Sub btnCompareFnd_Click(sender As Object, e As EventArgs) Handles btnCompareStrc.Click
            If strcLocal Is Nothing Or strcEDS Is Nothing Then Exit Sub
            'strcLocal.CompareMe(strcEDS)
            strcLocal.Equals(strcEDS)
        End Sub
        Private Sub btnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
            Dim strcFBD As New FolderBrowserDialog

            If strcFBD.ShowDialog = DialogResult.OK Then
                txtDirectory.Text = strcFBD.SelectedPath
            End If
        End Sub

        Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
            Dim str As New EDSStructure

            For Each fold As DirectoryInfo In New DirectoryInfo(txtDirectory.Text).GetDirectories
                For Each file As FileInfo In New DirectoryInfo(fold.FullName).GetFiles
                    If Not file.Extension.ToLower = ".eri" Then
                        file.Delete()
                    End If
                Next
                For Each file As FileInfo In New DirectoryInfo(fold.FullName).GetFiles
                    If file.Extension.ToLower = ".eri" Then
                        str.LogPath = fold.FullName
                        str.RunTNX(file.FullName, True)
                        Exit For
                    End If
                Next
            Next
        End Sub


#Region "Structure Tab Textbox Changes"

        Private Sub txtFndBU_TextChanged(sender As Object, e As EventArgs) Handles txtFndBU.TextChanged
            If isopening Then Exit Sub
            My.Settings.myBU = sender.text
            My.Settings.Save()
        End Sub

        Private Sub txtFndStrc_TextChanged(sender As Object, e As EventArgs) Handles txtFndStrc.TextChanged
            If isopening Then Exit Sub
            My.Settings.myStrID = sender.text
            My.Settings.Save()
        End Sub

        Private Sub txtFndWO_TextChanged(sender As Object, e As EventArgs) Handles txtFndWO.TextChanged
            If isopening Then Exit Sub
            My.Settings.myWO = sender.text
            My.Settings.Save()
        End Sub

        Private Sub txtDirectory_TextChanged(sender As Object, e As EventArgs) Handles txtDirectory.TextChanged
            If isopening Then Exit Sub
            My.Settings.myWorkArea = sender.text
            My.Settings.Save()
        End Sub

        Private Sub localWorkArea_TextChanged(sender As Object, e As EventArgs) Handles testLocalWorkarea.TextChanged
            If isopening Then Exit Sub


            If Microsoft.VisualBasic.Right(sender.text, 1) = "\" Then
                lFolder = Microsoft.VisualBasic.Left(sender.text, Len(sender.text) - 1)
            Else
                lFolder = sender.text
            End If

            My.Settings.localWorkArea = lFolder
            My.Settings.Save()
        End Sub

        Private Sub btnConduct_Click(sender As Object, e As EventArgs) Handles btnConduct.Click

            strcLocal.Conduct(CheckEditDevMode.Checked, CheckEditExcelVisible.Checked)

        End Sub
#End Region
#End Region

#Region "Shame"

        Dim tappy As Integer
        Private Sub PictureBox1_Click(sender As Object, e As EventArgs)
            Dim pwd As String
            If tappy = 0 Then
                MessageBox.Show("Stop touching me. GAR it makes me so mad!", "DO NOT TAP ON GLASS")
                tappy += 1
            ElseIf tappy = 1 Then
                MessageBox.Show("What, are you just doing this for the HALIBUT? Please stop.", "DO NOT TAP ON GLASS")
                tappy = 2
            ElseIf tappy = 2 Then
                If MessageBox.Show("Please, I have asked you nicely to stop. Let MINNOW, are you are going to stop?", "DO NOT TAP ON GLASS", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                    MessageBox.Show("WALLEYE just can't thank you enough for being reasonable. Now go away.", "DO NOT TAP ON GLASS")
                    tappy = 0
                Else
                    pwd = InputBox("COD dang it! Fine, I'll tell you what. If by some miracle you can guess my super secret password, I will let you tap as much as you want and I won't say another word.", "ENTER PASSWORD")
                    If pwd = "DanSmellowitz" Then
                        'If pwd = "Password" Or pwd = "password" Or pwd = "PASSWORD" Then
                        MessageBox.Show("What?! HOW?!! Okay fine, I am a fish of my word. You BETTA believe that I won't stop you from tapping on the glass as much as you want now", "GO AHEAD AND TAP ON GLASS, JERK")
                        tappy = 3
                    Else
                        MessageBox.Show("You clearly didn't want it bad enough. Better TUNA round and never try again.", "DO NOT TAP ON GLASS")
                        tappy = 0
                    End If
                End If
            End If
        End Sub

#End Region

#Region "Unit Testing - Control handlers only"

#Region "Simple Explorers"
        'Simple explorer change events based on textbox change events
        Private Sub testSaFolder_EditValueChanged(sender As Object, e As EventArgs) Handles testSaFolder.EditValueChanged
            Try
                seSA.SetCurrentDirectory(testSaFolder.Text)
            Catch ex As Exception
                seSA.SetCurrentDirectory(testSaFolder.Text.Replace(" - SA", ""))
            End Try

        End Sub
        Private Sub testFolder_EditValueChanged(sender As Object, e As EventArgs) Handles testFolder.EditValueChanged
            seNetwork.SetCurrentDirectory(testFolder.Text)
        End Sub
        Private Sub testLocalWorkarea_EditValueChanged(sender As Object, e As EventArgs) Handles testLocalWorkarea.EditValueChanged
            Try
                seLocal.SetCurrentDirectory(testLocalWorkarea.Text)
                My.Settings.localWorkArea = testLocalWorkarea.Text
                My.Settings.Save()
            Catch
            End Try
        End Sub
#End Region

#Region "Button Process Clicks"

        Private Sub TestSteps(sender As Object, e As EventArgs) Handles _
                    btnProcess1.Click, btnProcess2.Click, btnProcess3.Click, btnProcess4.Click,
                    btnProcess5.Click, btnProcess6.Click, btnProcess7.Click, btnProcess8.Click,
                    btnProcess9.Click, btnProcess10.Click, btnProcess11.Click, btnProcess12.Click,
                    btnProcess13.Click, btnProcess14.Click, btnProcess15.Click, btnProcess16.Click,
                    btnProcess17.Click, btnProcess18.Click, btnProcess19.Click, btnProcess20.Click,
                    btnProcess21.Click, btnProcess22.Click, btnProcess23.Click, btnProcess24.Click
            If isopening Then Exit Sub

            ButtonclickToggle(Me.Cursor, Cursors.WaitCursor)
            LogActivity("PROCESS | Start " & sender.tooltip.ToString)
            Dim tags As String() = sender.tag.ToString.Split("|")
            LogActivity("INFO | " & tags(1))

            Select Case tags(0).ToLower
                Case "step1"
                    Dim eriCount As Integer = New DirectoryInfo(Me.RefFolder).GetFiles.Count
                    Dim answer As DialogResult = vbYes

                    If eriCount > 0 Then
                        Dim msg As String = "Are you sure you would Like to get SA Reference files?" &
                                        vbCrLf & vbCrLf &
                                        "This process will archive files in the following folders: " &
                                        vbCrLf & vbCrLf &
                                        "Reference SA Files" &
                                        vbCrLf & vbCrLf &
                                        "Doing so may require creating new published and SAPI files for this test case."
                        answer = MsgBox(msg, vbCritical + vbYesNo, "Archive Files?")
                    End If

                    If answer = vbNo Then
                        Exit Select
                        LogActivity("INFO | Opted to not get SA Reference files.")
                    Else
                        If eriCount > 0 Then DoArchiving(Me.RefFolder)
                        Dim myfilesLst As New List(Of FileInfo)
                        'Loop through all files in the maestro folder for the current test case and iteration
                        For Each info As FileInfo In New DirectoryInfo(testSaFolder.Text).GetFiles
                            If Not info.FullName.Contains("~") Then
                                If info.Extension.ToLower = ".eri" Then
                                    'All eris permitted
                                    myfilesLst.Add(info)
                                    LogActivity("DEBUG | ERI Found: " & info.FullName)
                                ElseIf info.Extension.ToLower = ".xlsm" Then 'All tools are current xlsm files and this should be a safe assumption
                                    'Determine if the file is one of the templates
                                    Dim template As Tuple(Of Byte(), Byte(), String, String, String) = WhichFile(info)

                                    'If the properties of the tuple are nothing then they aren't templates
                                    If template.Item1 IsNot Nothing And template.Item2 IsNot Nothing And template.Item3 IsNot Nothing Then
                                        myfilesLst.Add(info)
                                        LogActivity("DEBUG | Template Found: " & info.FullName)
                                    End If
                                End If
                            End If
                        Next

                        Dim newFileCsv As New DataTable
                        newFileCsv.Columns.Add("FilePath", GetType(System.String))
                        newFileCsv.Columns.Add("Version", GetType(System.String))
                        newFileCsv.Columns.Add("RPAth", GetType(System.String))
                        For Each file As FileInfo In myfilesLst
                            Dim newFile As FileInfo = file.CopyTo(Me.RefFolder & "\" & file.Name)
                            LogActivity("DEBUG | " & newFile.Name & " has been copied to the SA Reference Files Folder")
                            newFileCsv.Rows.Add(newFile.FullName.Replace(dirUse, "").Replace("Iteration " & iteration, "[ITERATION]"), file.TemplateVersion, file.FullName)
                        Next

                        DatatableToCSV(newFileCsv, Me.RefFolder & "\File List.csv")
                        LogActivity("INFO | SA Reference Files have been copied into the test directory.")
                    End If

                Case "step2a", "step2b"
                    '''Create new iteration

                    Dim curItty As Integer = testIteration.Text
                    Dim nextItty As Integer = testNextIteration.Text
                    Dim ittyToCreate As Integer

                    If tags(0).ToLower = "step2a" Then
                        ittyToCreate = curItty
                    Else
                        ittyToCreate = nextItty
                    End If

                    CreateIteration(ittyToCreate)



                Case "step3", "step3a", "step3b", "step3c", "step3d"
                    'Make sure all necessary files exist in the required folders
                    If Me.iteration = 0 Then
                        MsgBox("Please create an iteration to continue.", vbInformation)
                        LogActivity("ERROR | Iteration not created.")
                        Exit Select
                    End If

                    Dim eriCount As Integer = New DirectoryInfo(Me.EriFolder).GetFiles.Count
                    Dim pubCount As Integer = New DirectoryInfo(Me.PubFolder).GetFiles.Count
                    Dim maeCount As Integer = New DirectoryInfo(Me.MaeFolder).GetFiles.Count
                    Dim manCount As Integer = New DirectoryInfo(Me.ManFolder).GetFiles.Count
                    Dim eriMsg As String = vbCrLf & vbCrLf & "    \Manual ERI"
                    Dim pubMsg As String = vbCrLf & vbCrLf & "    \Manual (Current)"
                    Dim maeMsg As String = vbCrLf & vbCrLf & "    \Iteration " & Me.iteration & "\Maestro"
                    Dim manMsg As String = vbCrLf & vbCrLf & "    \Iteration " & Me.iteration & "\Manual (SAPI)"
                    Dim msg As String = "Are you sure you would like to create new template files?" &
                                            vbCrLf & vbCrLf &
                                            "This process will archive files in the following folders:"
                    Dim answer As DialogResult = vbYes
                    Dim booEri As Boolean = IIf(eriCount = 0, True, False)
                    Dim booPub As Boolean = IIf(pubCount = 0, True, False)
                    Dim booMae As Boolean = IIf(maeCount = 0, True, False)
                    Dim booMan As Boolean = IIf(manCount = 0, True, False)

                    If tags(0).ToLower.Contains("a") And pubCount > 0 Then msg += pubMsg
                    If tags(0).ToLower.Contains("b") And eriCount > 0 Then msg += eriMsg
                    If tags(0).ToLower.Contains("c") And maeCount > 0 Then msg += maeMsg
                    If tags(0).ToLower.Contains("d") And manCount > 0 Then msg += manMsg

                    If ((eriCount > 0 Or pubCount > 0 Or maeCount > 0) And tags(0).ToLower = "step3") Or
                        (pubCount > 0 And tags(0).ToLower = "step3a") Or
                        (eriCount > 0 And tags(0).ToLower = "step3b") Or
                        (maeCount > 0 And tags(0).ToLower = "step3c") Or
                        (manCount > 0 And tags(0).ToLower = "step3d") Then
                        If forceAcrchiving Then
                            answer = vbYes
                        Else
                            answer = MsgBox(msg, vbCritical + vbYesNo, "Archive Files?")
                        End If
                    End If

                    If answer = vbNo Then
                        Exit Select
                        LogActivity("INFO | Opted to not create new templates.")
                    Else
                        Dim refDT As New DataTable
                        refDT = RefernceSADT()

                        If tags(0).ToLower = "step3" Or tags(0).ToLower = "step3a" Then
                            CreateCurrentTemplates(refDT, booPub, Not booPub)
                            LogActivity("INFO | All current template files have been created in the directory '\Manual (Current)'.")
                        End If

                        If tags(0).ToLower = "step3" Or tags(0).ToLower = "step3b" Then
                            CreateManualERI(refDT, Me.EriFolder, booEri, Not booEri)
                            LogActivity("INFO | All reference ERIs files have been created in the directory '\Manual ERI'.")
                        End If

                        If tags(0).ToLower = "step3" Or tags(0).ToLower = "step3c" Then
                            CreateSAPITemplates(refDT, Me.MaeFolder, "MaestroPath", booMae, Not booMae)
                            LogActivity("INFO | All files required for Maestro have been created in the directory '\Iteration" & Me.iteration & "\Maestro'.")

                            CreateManualERI(RefernceSADT, Me.MaeFolder, booMae, False)
                            LogActivity("INFO | All reference ERIs files have been created in the directory '\Iteration" & Me.iteration & "\Maestro'.")
                        End If

                        If tags(0).ToLower = "step3" Or tags(0).ToLower = "step3d" Then
                            CreateSAPITemplates(refDT, Me.ManFolder, "ManualPath", booMan, Not booMan)
                            LogActivity("INFO | All files SAPI files have been created in the directory '\Iteration" & Me.iteration & "\Manual (SAPI)'.")

                            'CreateManualERI(RefernceSADT, Me.ManFolder, booMan, False)
                            'LogActivity("INFO | All reference ERIs files have been created in the directory '\Iteration" & Me.iteration & "\Maestro'.")
                        End If

                        DatatableToCSV(refDT, RefFolder & "\File List.csv")
                    End If
                Case "step4", "step4a", "step4b", "step4c"
                    'Create Published versions of the files
                    If Me.iteration = 0 Then
                        MsgBox("Please create an iteration to continue.", vbInformation)
                        LogActivity("ERROR | Iteration not created.")
                        Exit Select
                    End If

                    If tags(0).ToLower = "step4" Or tags(0).ToLower = "step4b" Then
                        Dim pubCount As Integer = New DirectoryInfo(Me.PubFolder).GetFiles.Count
                        If Not pubCount > 0 Then
                            MsgBox("Please create current template files to continue.", vbInformation)
                            LogActivity("ERROR | Current template files not created.")
                            Exit Select
                        End If

                        ImportInputs("PublishedPath", My.Settings.booImportInputsExcelVisible)
                        LogActivity("INFO | Import inputs complete for current published versions.")
                    End If

                    If tags(0).ToLower = "step4" Or tags(0).ToLower = "step4a" Then
                        Dim maeCount As Integer = New DirectoryInfo(Me.MaeFolder).GetFiles.Count
                        If Not maeCount > 0 Then
                            MsgBox("Please create template SAPI files to continue.", vbInformation)
                            LogActivity("ERROR | SAPI template files not created.")
                            Exit Select
                        End If

                        ImportInputs("MaestroPath", My.Settings.booImportInputsExcelVisible)
                        LogActivity("INFO | Import inputs complete for Maestro SAPI versions.")

                        If tags(0).ToLower = "step4a" Then
                            Dim answer As DialogResult
                            answer = MsgBox("Would you like to copy the Maestro files into the '\Manual (SAPI)' folder?", vbYesNo + vbQuestion, "Copy Files?")
                            If answer = vbYes Then
                                DoArchiving(ManFolder)
                                For Each file As FileInfo In New DirectoryInfo(MaeFolder).GetFiles
                                    file.CopyTo(ManFolder & "\" & file.Name & "." & file.Extension)
                                    LogActivity("DEBUG | File Copied: '" & file.Name & "' From MAE to MAN")
                                Next
                            End If
                        End If
                    End If

                    If tags(0).ToLower = "step4" Or tags(0).ToLower = "step4c" Then
                        Dim mancount As Integer = New DirectoryInfo(Me.ManFolder).GetFiles.Count
                        If Not mancount > 0 Then
                            MsgBox("Please create template SAPI files to continue.", vbInformation)
                            LogActivity("ERROR | SAPI template files not created.")
                            Exit Select
                        End If

                        ImportInputs("ManualPath", My.Settings.booImportInputsExcelVisible)
                        LogActivity("INFO | Import inputs complete for Manual SAPI versions.")
                    End If
                Case "step5", "step5a", "step5b", "step5c"
                    'Run the ERI file in the Manual Reference Folder
                    Dim tempStrc As New EDSStructure
                    Dim myERIs As New List(Of String)
                    Dim mydir As String

                    If tags(0).ToLower = "step5" Or tags(0).ToLower = "step5a" Then
                        mydir = Me.EriFolder
                        myERIs.AddERIs(mydir)
                    End If

                    If tags(0).ToLower = "step5" Or tags(0).ToLower = "step5b" Then
                        mydir = Me.PubFolder
                        myERIs.AddERIs(mydir)
                    End If

                    If tags(0).ToLower = "step5" Or tags(0).ToLower = "step5c" Then
                        mydir = Me.ManFolder
                        myERIs.AddERIs(mydir)
                    End If

                    If myERIs.Count = 0 Then LogActivity("WARNING | No ERI files found to analyze.")

                    For Each eri As String In myERIs
                        If Not tempStrc.RunTNX(eri, True) Then
                            LogActivity("ERROR | Failed to run ERI: " & eri)
                            'tempStrc.AppendLog(Me.TestLogActivityPath)
                            'GoTo finishMe
                        Else
                            LogActivity("DEBUG | ERI: " & eri & " successuflly analyzed")
                            'tempStrc.AppendLog(Me.TestLogActivityPath)
                        End If
                    Next

                Case "step6", "step11"
                    Dim conductPath As String = ""
                    Select Case tags(0).ToLower
                        Case "step6"
                            conductPath = MaeFolder
                        Case "step11"
                            conductPath = EDSFolder
                    End Select

                    LogActivity("INFO | Conduct path: " & conductPath.Replace(dirUse, ""))

                    'Conduct the Maestro files
                    CreateStructure(conductPath)

                    'Conduct it!!!
                    '''This is commented out since Seb is actively working on the conduct function
                    '''Uncommented 4-27-2023
                    strcLocal.Conduct(CheckEditDevMode.Checked, CheckEditExcelVisible.Checked)
                    If DidConductProperly(strcLocal.LogPath) Then
                        ObjectToJson(Of EDSStructure)(strcLocal, conductPath & "\" & "EDSStructure_" & Now.ToString("MM/dd/yyyy HH:mm:ss tt").ToDirectoryString & ".ccistr")
                        LogActivity("INFO | Structure conducted successfully.")
                    Else
                        LogActivity("ERROR | Structure NOT conducted successfully.")
                    End If
                    strcLocal.AppendLog(Me.TestLogActivityPath, iteration)
                    SetStructureToPropertyGrid(strcLocal, pgcUnitTesting)
                Case "step7", "step7b", "step7a"
                    Dim checks As Tuple(Of Tuple(Of Boolean, DataTable), Tuple(Of Boolean, DataTable), Tuple(Of Boolean, DataTable), DataSet, Tuple(Of Boolean, DataTable), Tuple(Of Boolean, DataTable)) = CompareResults()
                    ButtonclickToggle(Me.Cursor, Cursors.Default)

                    'Item 1 = Manual Compared to Maestro   
                    '''Item 1 = Boolean specifying if they match
                    '''Item 2 = Data table of the comparisons
                    'Item 2 = Current Tools Compared to Manual
                    '''Item 1 = Boolean specifying if they match
                    '''Item 2 = Data table of the comparisons
                    'Item 3 = Current Tools Compared to Maestro
                    '''Item 1 = Boolean specifying if they match
                    '''Item 2 = Data table of the comparisons
                    'Item 4 = Dataset will all tables
                    If tags(0).ToLower = "step7" Or tags(0).ToLower = "step7a" Then
                        Dim newSum As New frmSummary
                        newSum.myDs = checks.Item4
                        newSum.Show()
                        LogActivity("DEBUG | Manual = Maestro --> " & checks.Item1.Item1.ToString)
                        LogActivity("DEBUG | Prod = Manual --> " & checks.Item2.Item1.ToString)
                        LogActivity("DEBUG | Prod = Maestro --> " & checks.Item3.Item1.ToString)

                        DatatableToCSV(checks.Item4.Tables("Combined Results"), Me.itFolder & "\All Summarized Results.csv")
                        LogActivity("DEBUG | Results output for reference SA files created: " & Me.itFolder & "\All Summarized Results.csv")
                        newSum.Refresh()
                        newSum.Export()
                    End If

                    If tags(0).ToLower = "step7" Or tags(0).ToLower = "step7b" Then
                        CreatetnxResults()
                    End If

                    LogActivity("INFO | Results for all files compared and created in testing directory.")
                Case "step8"
                    'report generator
                    Dim reportTemplate As String = "\\netapp4\common\Installers (Engineering Development)\SA Report Generator\Reference\Template.docx"
                    Dim reportMapping As String = "\\netapp4\common\Installers (Engineering Development)\SA Report Generator\Reference\mapping.xml"

                    Dim mylocation As String = DetermineFolder("Stop Report Generation")
                    If mylocation = "STOP" Then
                        LogActivity("INFO | Report generation cancelled")
                        Exit Select
                    End If

                    CreateStructure(mylocation, False)
                    SetStructureToPropertyGrid(strcLocal, pgcUnitTesting)

                    Dim testReport As New Report(reportTemplate, reportMapping, strcLocal, My.Settings.booReportOption)

                    Try
                        Using reportEditor As New ReportEditorForm(testReport)
                            reportEditor.ShowDialog()
                        End Using
                    Catch ex As Exception
                        MsgBox(ex.Message, vbExclamation, "Failed to load report generator")
                    End Try

                Case "step9"
                    Dim mylocation As String = DetermineFolder("Stop EDS Saving")
                    If mylocation = "STOP" Then
                        LogActivity("INFO | EDS Saving cancelled")
                        Exit Select
                    End If

                    LogActivity("INFO | Saving to selected database: " & My.Settings.dbSelection)
                    Dim dbToSend As String

                    If My.Settings.serverActive = "dbDevelopment" Then
                        If My.Settings.dbSelection = "UAT" Then
                            dbToSend = EDSdbUserAcceptance
                        Else
                            dbToSend = EDSdbDevelopment
                        End If
                    Else
                        dbToSend = EDSdbProduction
                    End If

                    CreateStructure(mylocation, False)

                    Dim resDT As New DataTable
                    Dim nowString As String = Now.ToString("MM/dd/yyyy HH:mm:ss tt").ToDirectoryString
                    resDT = strcLocal.SavetoEDS(EDSnewId, dbToSend, True)
                    resDT.ToCSV(mylocation & "\EDS Result" & nowString & ".csv")
                    LogActivity("INFO | EDS Response Saved: " & mylocation & "\EDS Result" & nowString & ".csv")

                    If resDT.Rows(0).Item("Result").ToString = "Error" Then
                        Dim errNum As String
                        Dim errLine As String
                        Dim errMessage As String
                        Dim errSev As String
                        Dim errState As String
                        Dim dr As DataRow = resDT.Rows(0)

                        errNum = dr.Item("ErrorNumber").ToString
                        errLine = dr.Item("ErrorLine").ToString
                        errMessage = dr.Item("ErrorMessage").ToString
                        errSev = dr.Item("ErrorSeverity").ToString
                        errState = dr.Item("ErrorState").ToString

                        LogActivity("ERROR | Saving EDS Query failed")
                        LogActivity("DEBUG | Error Number: " & errNum)
                        LogActivity("DEBUG | Error Line: " & errLine)
                        LogActivity("DEBUG | Error Message: " & errMessage)
                        LogActivity("DEBUG | Error Severity: " & errSev)
                        LogActivity("DEBUG | Error State: " & errState)
                    Else
                        LogActivity("INFO | EDS data saved successfully")
                    End If

                    Dim myQuery As String = My.Computer.Clipboard.GetText()
                    Dim queryName As String = "\EDS Query " & nowString
                    If myQuery.WriteAllToFile(myQuery) Then
                        LogActivity("INFO | Query saved to " & mylocation & queryName)
                    Else
                        LogActivity("ERROR | Query could not be saved to folder.")
                    End If

                    SetStructureToPropertyGrid(strcLocal, pgcUnitTesting)
                Case "step10"
                    LogActivity("INFO | Loading from selected database: " & My.Settings.dbSelection)
                    Dim workingdirectory As String
                    Dim dbToSend As String

                    If My.Settings.serverActive = "dbDevelopment" Then
                        If My.Settings.dbSelection = "UAT" Then
                            dbToSend = EDSdbUserAcceptance
                        Else
                            dbToSend = EDSdbDevelopment
                        End If
                    Else
                        dbToSend = EDSdbProduction
                    End If


                    'Check if eds folder exists
                    Dim edsExists As Boolean = Directory.Exists(EDSFolder)

                    'create if it doesn't
                    If Not edsExists Then
                        Directory.CreateDirectory(EDSFolder)
                    End If

                    'archive if files exist
                    Dim edsFiles As Integer = Directory.GetFiles(EDSFolder).Count
                    If edsFiles > 0 Then
                        DoArchiving(EDSFolder)
                    End If

                    'strcLocal.Clear()
                    strcLocal = New EDSStructure(bus_unit, structure_id, work_order_seq_num, EDSFolder, EDSFolder, EDSnewId, EDSdbActive)
                    strcLocal.SaveTools(EDSFolder)
                    LogActivity("INFO | All files have been created in the directory '\Iteration " & iteration & "\EDS'.")
            End Select

finishMe:
            LogActivity("PROCESS | End " & sender.tooltip.ToString, True)
            ButtonclickToggle(Me.Cursor, Cursors.Default)
        End Sub

        Private Function DetermineFolder(ByVal stopping As String) As String
            Dim edsExists As Boolean = Directory.Exists(EDSFolder)
            Dim whichFolder As New DialogResult
            Dim maeOption As String = "YES = '\Maestro' Folder" & vbCrLf & vbCrLf
            Dim manOption As String = "NO = '\Manual (SAPI)' Folder" & vbCrLf & vbCrLf
            Dim cancelOption As String = "CANCEL = " & stopping & vbCrLf
            Dim edsOption As String = "NO = '\EDS' Folder" & vbCrLf & vbCrLf
            Dim filesPath As String

            whichFolder = MsgBox("Which folder would you like use to create a report?" & vbCrLf & vbCrLf & IIf(edsExists, maeOption + edsOption + cancelOption, maeOption + manOption + cancelOption), vbYesNoCancel + vbInformation, "Which Folder?")

            If whichFolder = vbCancel Then Return "STOP"

            Select Case whichFolder
                Case vbYes
                    filesPath = MaeFolder
                Case vbNo
                    If edsExists Then
                        filesPath = EDSFolder
                        LogActivity("INFO | " & stopping.Replace("Stop ", "") & ": " & filesPath)
                    Else
                        filesPath = ManFolder
                        LogActivity("INFO | " & stopping & ": " & filesPath)
                    End If
            End Select



            Return filesPath
        End Function

        Private Sub testBugFile_Click(sender As Object, e As EventArgs) Handles testBugFile.Click
            If Not Directory.Exists(Me.itFolder & "\Bug Reference Files") Then
                Directory.CreateDirectory(Me.itFolder & "\Bug Reference Files")
            End If

            Dim ofd As New XtraOpenFileDialog
            ofd.InitialDirectory = Environment.SpecialFolder.UserProfile.ToString
            ofd.Multiselect = True

            If ofd.ShowDialog = DialogResult.OK Then
                For Each file As String In ofd.FileNames
                    Dim info As New FileInfo(file)
                    info.CopyTo(Me.itFolder & "\Bug Reference Files\" & info.Name)
                    LogActivity("DEBUG | Reference file has been copied: " & info.Name)
                Next
            End If

            ofd.Dispose()
        End Sub

        'Close test case and unload eryting
        '''Basically just the opposite of the test case dropdown
        Private Sub testClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
            isopening = True 'Only used because isloading wasn't there and I didn't feel like adding it
            ButtonclickToggle(Me.Cursor)

            LogActivity("FINISH | Test Case " & Me.testCase)

            PushIt(True)
            TearDownWorkArea()
            sender.enabled = False
            btnCheckout.Enabled = True
            My.Settings.MyTestCase = 0
            My.Settings.Save()
            testPush.Enabled = False
            SetTestIDLabels()
            ButtonclickToggle(Me.Cursor)
            isopening = False
        End Sub

        Private Sub testCheckout_click(sender As Object, e As EventArgs) Handles btnCheckout.Click
            If isopening Then Exit Sub

            If IsNumeric(testID.Text) Then
                ButtonclickToggle(Me.Cursor)
                If CheckOut() = vbNo Then GoTo StopLookingAtMeSwan
                sender.enabled = False
                btnClose.Enabled = True
                My.Settings.MyTestCase = testCase
                My.Settings.Save()

                LogActivity("START | Test Case " & Me.testCase, True)
                LogActivity("INFO | Using CCI Engineering Datatransferer " & CCI_Engineering_Templates.myVersion)
                LogActivity("INFO | Using Testing Winform " & testingVersion)

                If Not Directory.Exists(Me.dirUse & "\Test ID " & Me.testCase & "\Iteration " & Me.iteration) Then
                    CreateIteration(Me.iteration)
                End If

                If File.Exists(Me.dirUse & "\Test ID " & Me.testCase & "\Checked Out.txt") Then
                    testPush.Enabled = True
                End If
                SetTestIDLabels()
                isopening = True
                testID.SelectedIndex = testCase - 1
                isopening = False
StopLookingAtMeSwan:
                ButtonclickToggle(Me.Cursor)
            End If

        End Sub

        Private Sub testPush_Click(sender As Object, e As EventArgs) Handles testPush.Click
            ButtonclickToggle(Me.Cursor)
            PushIt(False)
            ButtonclickToggle(Me.Cursor)
        End Sub

        Private Sub testPull_Click(sender As Object, e As EventArgs) Handles testPull.Click
            ButtonclickToggle(Me.Cursor)
            PullIt(False)
            ButtonclickToggle(Me.Cursor)
        End Sub

#End Region

        'Main tab indexchanged event
        Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
            If TabControl1.SelectedTab.Name = pgUnitTesting.Name Then
                If unitTestCases.Count > 0 Then Exit Sub

                LoadTestCases(unitTestCases)
            End If
        End Sub

        'Test ID drop down changed event
        Private Sub testID_SelectedIndexChanged(sender As Object, e As EventArgs) Handles testID.SelectedIndexChanged
            If isopening Then Exit Sub
            Dim setitUp As Boolean = False
            Dim resetit As Boolean = True
            Dim closePush As Boolean = False
            Dim checkOut As Boolean = True
            Dim tempTestcae As Integer = testCase

            ButtonclickToggle(Me.Cursor)
            isopening = True

            If testID.Text.Contains("Checked Out") Then
                setitUp = True
                closePush = True
                checkOut = False
            ElseIf Directory.Exists(lFolder & "\Test ID " & testCase) Then
                setitUp = True
            Else
                TearDownWorkArea()
                My.Settings.MyTestCase = 0
                My.Settings.Save()
                testID.SelectedIndex = tempTestcae - 1
            End If

            If setitUp Then
                If My.Settings.MyTestCase > 0 Then
                    resetit = False
                End If

                SetUpWorkArea(testCase - 1, resetit)
                My.Settings.MyTestCase = testCase
            End If

            btnClose.Enabled = closePush
            testPush.Enabled = closePush
            btnCheckout.Enabled = checkOut

            isopening = False
            ButtonclickToggle(Me.Cursor)
        End Sub

        'Log that a test case is ending
        Private Sub testID_EditValueChanging(sender As Object, e As DevExpress.XtraEditors.Controls.ChangingEventArgs) Handles testID.EditValueChanging
            If isopening Then Exit Sub

            'Dim prevtestcase As String
            'Try
            '    prevtestcase = e.OldValue.ToString
            '    If IsNumeric(prevtestcase) Then
            '        LogActivity("FINISH | Test Case" & prevtestcase)
            '        ResetControls(False)
            '    Else
            '        'Enable all of the buttons for use in the iteration
            '        ResetControls()
            '    End If
            'Catch ex As Exception
            '    ResetControls()
            'End Try
        End Sub

        'Rich textbox changed event for test notes
        Private Sub rtbNotes_TextChanged(sender As Object, e As EventArgs) Handles rtbNotes.TextChanged
            If isopening Then Exit Sub

            Try
                System.IO.File.WriteAllText(Me.dirUse & "\Test ID " & Me.testCase & "\Test Notes.txt", rtbNotes.Text)
            Catch
            End Try
        End Sub

        Private Sub CheckEditDevMode_CheckedChanged(sender As Object, e As EventArgs) Handles CheckEditDevMode.CheckedChanged
            My.Settings.booConductDevMode = CheckEditDevMode.Checked
            My.Settings.Save()
        End Sub
        Private Sub CheckEditAutoReport_CheckedChanged(sender As Object, e As EventArgs) Handles CheckEditAutoReport.CheckedChanged
            My.Settings.booReportOption = CheckEditAutoReport.Checked
            My.Settings.Save()
        End Sub

        Private Sub CheckEditExcelVisible_CheckedChanged(sender As Object, e As EventArgs) Handles CheckEditExcelVisible.CheckedChanged
            My.Settings.booConductExcelVis = CheckEditExcelVisible.Checked
            My.Settings.Save()
        End Sub

        Private Sub CheckEditExcelVisibleII_CheckedChanged(sender As Object, e As EventArgs) Handles CheckEditExcelVisibleII.CheckedChanged
            My.Settings.booImportInputsExcelVisible = CheckEditExcelVisibleII.Checked
            My.Settings.Save()
        End Sub


        Private Sub btnAuto_Click(sender As Object, e As EventArgs) Handles btnAuto.Click
            ButtonclickToggle(Me.Cursor)
            forceAcrchiving = True
            For i As Integer = 0 To 99
                'Test cases we removed and guyed towers with LRT
                If Not (i = 87 Or i = 76 Or i = 52 Or i = 51 Or i = 49 Or i = 91 Or i = 92) Then
                    'Set up iteration folder
                    testID.SelectedIndex = i
                    'Create Iteration folder
                    If Not Directory.Exists(Me.dirUse & "\Test ID " & Me.testCase & "\Iteration " & Me.iteration) Then
                        CreateIteration(Me.iteration)
                    End If
                    'Create MAE files
                    TestSteps(btnProcess11, EventArgs.Empty)
                    'Create Man files just in case compare results needs it
                    TestSteps(btnProcess13, EventArgs.Empty)
                    'Import inputs for MAE files
                    TestSteps(btnProcess5, EventArgs.Empty)
                    'Conduct MAE files
                    TestSteps(btnProcess7, EventArgs.Empty)
                    'Create excel results (MAE v. CUR only, but everything will be created)
                    TestSteps(btnProcess21, EventArgs.Empty)
                End If
            Next
            forceAcrchiving = False
            ButtonclickToggle(Me.Cursor)
        End Sub

        Private Sub btnChecking_Click(sender As Object, e As EventArgs) Handles btnChecking.Click
            ButtonclickToggle(Me.Cursor)
            For i As Integer = 0 To 99
                'Test cases we removed and guyed towers with LRT
                If Not (i = 87 Or i = 76 Or i = 52 Or i = 51 Or i = 49 Or i = 91 Or i = 92) Then
                    'Pull all files and check it out if possible.
                    PullIt(True, True)
                End If
            Next
            SetTestIDLabels()
            ButtonclickToggle(Me.Cursor)
        End Sub

        Private Sub testGetWOs_click(sender As Object, e As EventArgs) Handles testGetWOs.Click
            OracleLoader("SELECT wo_seqnum, eng_app_id, crrnt_rvsn_num, bus_unit, structure_id
                            FROM work_order_reporting_mv@ISITPRD.CROWNCASTLE.COM
                            WHERE bus_unit = '" & bus_unit.ToString & "' AND structure_id = '" & structure_id & "'
                            AND item_type IN ('SA - Structural Analysis','SA - Structural Analysis w/o App','SDD - Structural Design Drawings') 
                            ORDER BY wo_seqnum DESC",
                         "MyWOs", 5000, "ords")
            GridView1.Columns.Clear()
            gcViewer.DataSource = Nothing
            gcViewer.DataSource = ds.Tables("MyWOs")
            gcViewer.RefreshDataSource()
            GridView1.BestFitColumns(True)
        End Sub

        Private Sub toggleDevUat_Toggled(sender As Object, e As EventArgs) Handles toggleDevUat.Toggled
            If isopening Then Exit Sub

            If My.Settings.serverActive = "dbDevelopment" Then
                If toggleDevUat.IsOn Then
                    My.Settings.dbSelection = "UAT"
                    EDSdbActive = EDSdbUserAcceptance
                Else
                    My.Settings.dbSelection = "DEV"
                    EDSdbActive = EDSdbDevelopment
                End If
            Else
            End If

            My.Settings.Save()
        End Sub
#End Region

#Region "My Largely Little Helpers"

#Region "Properties"
        Public Enum SyncDirection
            RtoLocal
            LocaltoR
        End Enum

        'Determine which directory to use. 
        Public ReadOnly Property dirUse As String
            Get
                If chkWorkLocal.Checked Then
                    Return lFolder
                Else
                    Return rFolder
                End If
            End Get
        End Property

        Public ReadOnly Property bus_unit As Integer
            Get
                Return If(IsNumeric(testBu.Text), testBu.Text, Nothing)
            End Get
        End Property
        Public ReadOnly Property work_order_seq_num As Integer
            Get
                Return If(IsNumeric(testWo.Text), testWo.Text, Nothing)
            End Get
        End Property
        Public ReadOnly Property structure_id As String
            Get
                Return testSid.Text
            End Get
        End Property

        Public ReadOnly Property TestLogActivityPath As String
            Get
                Return dirUse & "\Test ID " & testCase & "\Test Activity.txt"
            End Get
        End Property

        'Set the directories to reference based on working local or on the network.
        Public ReadOnly Property itFolder As String
            Get
                Return dirUse & "\Test ID " & testCase & "\Iteration " & iteration
            End Get
        End Property
        Public ReadOnly Property EDSFolder As String
            Get
                Return itFolder & "\EDS"
            End Get
        End Property
        Public ReadOnly Property MaeFolder As String
            Get
                Return itFolder & "\Maestro"
            End Get
        End Property
        Public ReadOnly Property ManFolder As String
            Get
                Return itFolder & "\Manual (SAPI)"
            End Get
        End Property
        Public ReadOnly Property PubFolder As String
            Get
                Return dirUse & "\Test ID " & testCase & "\Manual (Current)"
            End Get
        End Property
        Public ReadOnly Property RefFolder As String
            Get
                Return dirUse & "\Test ID " & testCase & "\Reference SA Files"
            End Get
        End Property
        Public ReadOnly Property EriFolder As String
            Get
                Return dirUse & "\Test ID " & testCase & "\Manual ERI"
            End Get
        End Property
        Public ReadOnly Property BugFolder As String
            Get
                Return dirUse & "\Test ID " & testCase & "\Bugs"
            End Get
        End Property

        'Set the test case to use throughout.
        Public ReadOnly Property testCase As Integer?
            Get
                If testID.Text.Contains("|") Then
                    Try
                        Return If(IsNumeric(testID.Text.Split("|")(0)), testID.Text.Split("|")(0), Nothing)
                    Catch ex As Exception

                    End Try
                Else
                    Return If(IsNumeric(testID.Text), testID.Text, Nothing)
                End If

            End Get
        End Property
        'Set the iteration to use throughout.
        Public ReadOnly Property iteration As Integer?
            Get
                Return If(IsNumeric(testIteration.Text), testIteration.Text, Nothing)
            End Get
        End Property

        Public unitTestCases As New List(Of TestCase)
        Public rFolder As String = "\\netapp4\cad\Development\SAPI Testing\Unit Testing"
        Public lFolder As String
        Public thr1 As Thread
        Public DirectorySync As RoboCommand
#End Region

#Region "Checkin/Checkout"
        Public Sub PullIt(ByVal checkout As Boolean, Optional ByVal checkingOutEverything As Boolean = False)
            ForceSyncing(SyncDirection.RtoLocal, True)

            If Not testPush.Enabled Then
                If Not checkingOutEverything Then
                    File.Delete(Me.dirUse & "\Test ID " & Me.testCase & "\Checked Out.txt")
                End If
            End If

            If checkout Then
                CreateCheckOutFiles()
            End If
        End Sub

        Public Sub PushIt(ByVal checkin As Boolean)
            ForceSyncing(SyncDirection.LocaltoR, True)

            If checkin Then
                DeleteCheckOutFiles()
            End If
        End Sub

        Public Sub CreateCheckOutFiles()
            File.Create(Me.dirUse & "\Test ID " & Me.testCase & "\Checked Out.txt").Dispose()
            File.Create(rFolder & "\Test ID " & Me.testCase & "\Checked Out.txt").Dispose()
            testPush.Enabled = True
        End Sub

        Public Sub DeleteCheckOutFiles()
            File.Delete(Me.dirUse & "\Test ID " & Me.testCase & "\Checked Out.txt")
            File.Delete(rFolder & "\Test ID " & Me.testCase & "\Checked Out.txt")
        End Sub

        Public Function CheckOut() As DialogResult
            'Get the id from the load case lists
            Dim id As Integer = Me.testCase - 1

            'if the directory exists on the R drive but not locally
            '''Determine if data needs to be pulled
            If Directory.Exists(rFolder & "\Test ID " & testCase) And Directory.Exists(lFolder & "\Test ID " & testCase) Then
                Dim answer As DialogResult = vbNo
                'If not checked out locally
                If Not New FileInfo(Me.dirUse & "\Test ID " & Me.testCase & "\Checked Out.txt").Exists Then
                    'if checked out on the network.
                    If New FileInfo(rFolder & "\Test ID " & Me.testCase & "\Checked Out.txt").Exists Then
                        answer = MsgBox("Test Case " & Me.testCase & " has been checked out by someone else and a version of this test case exists locally." & vbCrLf & vbCrLf & "Pulling data from the R: drive will delete all local files." & vbCrLf & vbCrLf & "Would you like to PULL the R: drive files locally for reference?", vbYesNo + vbInformation, "Checkout Not Allowed")
                        If answer = vbYes Then
                            PullIt(False)
                        Else
                            Return answer
                        End If
                    Else
                        answer = MsgBox("Test Case " & Me.testCase & " exists on the R: drive and locally." & vbCrLf & vbCrLf & "Pulling data from the R: drive will delete all local files." & vbCrLf & vbCrLf & "Would you like to PULL the R: drive files locally and CHECKOUT the test case?", vbYesNo + vbInformation, "Pull and Checkout")
                        If answer = vbYes Then
                            PullIt(True)
                        Else
                            Return answer
                            'CreateCheckOutFiles()
                        End If
                    End If
                End If
            ElseIf Not Directory.Exists(rFolder & "\Test ID " & testCase) And Not Directory.Exists(lFolder & "\Test ID " & testCase) Then
                CreateInitialTestWorkArea(id)
                CreateCheckOutFiles()

            ElseIf Directory.Exists(rFolder & "\Test ID " & testCase) And Not Directory.Exists(lFolder & "\Test ID " & testCase) Then
                My.Computer.FileSystem.CopyDirectory(rFolder & "\Test ID " & testCase, lFolder & "\Test ID " & testCase)
                If New FileInfo(rFolder & "\Test ID " & Me.testCase & "\Checked Out.txt").Exists Then
                    File.Delete(Me.dirUse & "\Test ID " & Me.testCase & "\Checked Out.txt")
                Else
                    CreateCheckOutFiles()
                End If
            End If

            SetUpWorkArea(id, Not btnProcess1.Enabled)
        End Function

        Public Sub CreateInitialTestWorkArea(ByVal listID As Integer)
            Directory.CreateDirectory(rFolder & "\Test ID " & Me.testCase)
            Directory.CreateDirectory(Me.dirUse & "\Test ID " & Me.testCase)
            Directory.CreateDirectory(Me.dirUse & "\Test ID " & Me.testCase & "\Manual (Current)")
            Directory.CreateDirectory(Me.dirUse & "\Test ID " & Me.testCase & "\Reference SA Files")
            Directory.CreateDirectory(Me.dirUse & "\Test ID " & Me.testCase & "\Manual ERI")
            'DirectoryCreator("\Test ID " & Me.testCase & "\Bug Tracking")
            File.Create(Me.dirUse & "\Test ID " & Me.testCase & "\Test Notes.txt").Dispose()
            File.Create(Me.dirUse & "\Test ID " & Me.testCase & "\Test Activity.txt").Dispose()

            'When first creating the test case folder general notes (Salute) will be created to get started. 
            Using sw As New StreamWriter(Me.dirUse & "\Test ID " & Me.testCase & "\Test Notes.txt")
                sw.WriteLine("Testing notes for Test ID " & Me.testCase)
                sw.WriteLine("BU = " & unitTestCases(listID).BU)
                sw.WriteLine("Structure ID = " & unitTestCases(listID).SID)
                sw.WriteLine("Wo = " & unitTestCases(listID).WO)
                sw.WriteLine("SA Work Area = " & unitTestCases(listID).SAWorkArea)
                sw.WriteLine("Load Combination = " & unitTestCases(listID).COMB)
                sw.Close()
            End Using
        End Sub

        Public Sub SetUpWorkArea(ByVal listID As Integer, Optional ByVal resetit As Boolean = True)
            'If the directory exists, it just loads in the text file for reference
            rtbNotes.Text = System.IO.File.ReadAllText(Me.dirUse & "\Test ID " & Me.testCase & "\Test Notes.txt")

            'Set the site data loaded from the test case CSV.
            testBu.Text = unitTestCases(listID).BU
            testSid.Text = unitTestCases(listID).SID
            testWo.Text = unitTestCases(listID).WO
            testSaFolder.Text = unitTestCases(listID).SAWorkArea 'This will update the directory for the SA Reference folder
            testFolder.Text = "R:\Development\SAPI Testing\Unit Testing\Test ID " & Me.testCase 'This will update the directory for the network test case
            testComb.Text = unitTestCases(listID).COMB

            'Iteration count is determined
            '''A count of folders containing the word 'iteration' are counted
            Dim itCount As Integer = 0
            Dim newIt As Integer
            Dim maxIt As Integer = 0
            For Each subDir In New DirectoryInfo(Me.dirUse & "\Test ID " & Me.testCase).GetDirectories
                If subDir.Name.Contains("Iteration ") Then
                    itCount += 1
                    newIt = CType(subDir.Name.Replace("Iteration ", ""), Integer)
                    If newIt > maxIt Then maxIt = newIt
                End If
            Next

            itCount = Math.Max(Math.Max(itCount, currentTestingIteration), maxIt)

            testIteration.Text = itCount
            testNextIteration.Text = itCount + 1

            'Update the local directory to the local test case. 
            Try
                seLocal.SetCurrentDirectory(Me.dirUse & "\Test ID " & Me.testCase)
            Catch ex As Exception
            End Try

            If resetit Then ResetControls()
            mainLogViewer.ReloadActivityLog()
        End Sub

        Public Sub TearDownWorkArea()
            If btnProcess1.Enabled Then ResetControls()

            'Set the values of all inputs
            'This can't be in the resetcontrols because it is secific to this process. 
            testBu.Text = ""
            testSid.Text = ""
            testWo.Text = ""
            testSaFolder.Text = ""
            testFolder.Text = ""
            testComb.Text = ""
            testID.SelectedIndex = -1
            'testIteration.Text = ""
            'testNextIteration.Text = ""
            rtbNotes.Text = ""
            GridView1.Columns.Clear()
            gcViewer.DataSource = Nothing
            pgcUnitTesting.SelectedObject = Nothing

            Try
                seNetwork.SetCurrentDirectory(Environment.SpecialFolder.MyDocuments.ToString)
                'seNetwork.Dispose()
            Catch ex As Exception
            End Try
            Try
                seSA.SetCurrentDirectory(Environment.SpecialFolder.MyDocuments.ToString)
                'seSA.Dispose()
            Catch ex As Exception
            End Try
            Try
                seLocal.SetCurrentDirectory(Me.dirUse)
            Catch ex As Exception
            End Try

            mainLogViewer.Clear()
        End Sub

        'Reset form to disable or enable controls required for testing. 
        Public Sub ResetControls(Optional ByVal reset As Boolean = True)
            btnProcess1.Enabled = Not btnProcess1.Enabled
            btnProcess2.Enabled = Not btnProcess2.Enabled
            btnProcess3.Enabled = Not btnProcess3.Enabled
            btnProcess4.Enabled = Not btnProcess4.Enabled
            btnProcess5.Enabled = Not btnProcess5.Enabled
            btnProcess6.Enabled = Not btnProcess6.Enabled
            btnProcess7.Enabled = Not btnProcess7.Enabled
            btnProcess8.Enabled = Not btnProcess8.Enabled
            btnProcess9.Enabled = Not btnProcess9.Enabled
            btnProcess10.Enabled = Not btnProcess10.Enabled
            btnProcess11.Enabled = Not btnProcess11.Enabled
            btnProcess12.Enabled = Not btnProcess12.Enabled
            btnProcess13.Enabled = Not btnProcess13.Enabled
            btnProcess14.Enabled = Not btnProcess14.Enabled
            btnProcess15.Enabled = Not btnProcess15.Enabled
            btnProcess16.Enabled = Not btnProcess16.Enabled
            btnProcess17.Enabled = Not btnProcess17.Enabled
            btnProcess18.Enabled = Not btnProcess18.Enabled
            btnProcess19.Enabled = Not btnProcess19.Enabled
            btnProcess20.Enabled = Not btnProcess20.Enabled
            btnProcess21.Enabled = Not btnProcess21.Enabled
            btnProcess22.Enabled = Not btnProcess22.Enabled
            btnProcess23.Enabled = Not btnProcess23.Enabled
            btnProcess24.Enabled = Not btnProcess24.Enabled
            testGetWOs.Enabled = Not testGetWOs.Enabled

            testWo.ReadOnly = Not testWo.ReadOnly

            XtraTabControl1.Enabled = Not XtraTabControl1.Enabled
            testBugFile.Enabled = Not testBugFile.Enabled
            If New FileInfo(Me.dirUse & "\Test ID " & Me.testCase & "\Checked Out.txt").Exists Then testPush.Enabled = Not testPush.Enabled
            testPull.Enabled = Not testPull.Enabled
            mainLogViewer.Enabled = Not mainLogViewer.Enabled
            rtfactivityLog.Visible = Not rtfactivityLog.Visible
            CheckEditDevMode.Enabled = Not CheckEditDevMode.Enabled
            CheckEditExcelVisible.Enabled = Not CheckEditExcelVisible.Enabled
            CheckEditExcelVisibleII.Enabled = Not CheckEditExcelVisibleII.Enabled
            CheckEditAutoReport.Enabled = Not CheckEditAutoReport.Enabled
            If Environment.UserName.ToLower = "imiller" Or
               Environment.UserName.ToLower = "stanley" Or
               Environment.UserName.ToLower = "dsmilowitz" Or
               Environment.UserName.ToLower = "chall" Or
               Environment.UserName.ToLower = "mrudolph" Then
                toggleDevUat.Enabled = Not toggleDevUat.Enabled
            End If
        End Sub
#End Region

#Region "Logging"
        'Logs any activity happening during the unit testing process
        Public Sub LogActivity(msg As String, Optional ByVal loadLog As Boolean = False)
            ' Get the current date and time
            Dim dt As String = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss tt")
            'Dim splt() As String = dt.Split(" ")
            'dt = splt(1) '& " " & splt(2)

            ' Print the message to the console
            Console.WriteLine(dt & " | " & Environment.UserName & " | " & msg)

            ' Wrap the file operation in a try-catch block to handle exceptions
            Try
                ' If the log file does not exist, establish intro
                If Not File.Exists(TestLogActivityPath) Then
                    File.Create(TestLogActivityPath).Dispose()
                End If
                ' Use a StreamWriter to write to the log file
                ' The 'True' argument appends to the file if it already exists
                Using sw As New StreamWriter(TestLogActivityPath, True)
                    ' Write the log message to the file
                    sw.WriteLine(dt & " | " & Environment.UserName & " | " & msg & " | " & Me.iteration.ToString)
                End Using
                If loadLog Then
                    mainLogViewer.ReloadActivityLog()
                End If
            Catch ex As Exception
                ' Handle the exception
                Console.WriteLine("Error writing to log file: " & ex.Message)
            End Try
        End Sub

#End Region

#Region "Syncing"
        'Being robocommmand to copy files to R: drive on a regular basis.
        Public Sub InitializeLocaltoCentralSync(ByRef syncer As RoboCommand, ByVal source As String, ByVal destination As String, ByVal Purge As Boolean)
            With syncer
                If .IsRunning Then
                    Return
                End If

                .CopyOptions.Source = source
                .CopyOptions.Destination = destination
                .CopyOptions.CopySubdirectories = True
                .CopyOptions.UseUnbufferedIo = True
                .CopyOptions.CopySubdirectoriesIncludingEmpty = True
                .CopyOptions.Purge = Purge
                'If Not waitforsync Then
                '    .CopyOptions.MultiThreadedCopiesCount = 4
                '    .CopyOptions.MonitorSourceChangesLimit = 5
                '    .CopyOptions.MonitorSourceTimeLimit = 5
                'End If
                .Name = "RoboCopy_UnitTesting"
                .RetryOptions.RetryCount = 1
                .RetryOptions.RetryWaitTime = 2
            End With
        End Sub

        Public Sub ForceSyncing(ByVal direction As SyncDirection,
                                Optional ByVal Purge As Boolean = False)

            Dim myResults As RoboSharp.Results.RoboCopyResults
            Dim filestat As RoboSharp.Results.Statistic
            Dim dirStat As RoboSharp.Results.Statistic
            Dim newSync As New RoboCommand

            KillRoboCops()

            If direction = SyncDirection.LocaltoR Then
                InitializeLocaltoCentralSync(newSync, lFolder & "\Test ID " & Me.testCase, rFolder & "\Test ID " & Me.testCase, Purge)
                If Purge Then LogActivity("DEBUG | Purged Folder: " & newSync.CopyOptions.Destination)
                LogActivity("DEBUG | Forced Sync local drive to R: Drive")
            Else
                InitializeLocaltoCentralSync(newSync, rFolder & "\Test ID " & Me.testCase, lFolder & "\Test ID " & Me.testCase, Purge)
                If Purge Then LogActivity("DEBUG | Purged Folder: " & newSync.CopyOptions.Destination)
                LogActivity("DEBUG | Forced Sync R: drive to local drive")
            End If

            newSync.Start.Wait()
            myResults = newSync.GetResults
            filestat = myResults.FilesStatistic
            dirStat = myResults.DirectoriesStatistic
            newSync.Stop()
            newSync.Dispose()
            KillRoboCops()
        End Sub

        'Function to check if 2 directories ahve the same file count
        'No longer used
        Public Function FileCountMatches(ByVal dir1 As String, ByVal dir2 As String) As Boolean
            Dim rCount As Integer = Directory.GetFiles(dir1, "*", SearchOption.AllDirectories).Count
            Dim lCount As Integer = Directory.GetFiles(dir2, "*.*", SearchOption.AllDirectories).Count
            If rCount = lCount Then
                Return True
            Else
                Return False
            End If
        End Function

        'Funciton used to wait to check if 2 test logs match
        'No longer used
        Public Function LogMatches(ByVal dir1 As String, ByVal dir2 As String) As Boolean
            Dim info1 As New FileInfo(dir1 & "\Test Activity.txt")
            Dim info2 As New FileInfo(dir2 & "\Test Activity.txt")
            Dim length1 As String
            Dim length2 As String

            Using sr As New StreamReader(info1.FullName)
                length1 = sr.ReadToEnd.Empty
                sr.Close()
            End Using

            Using sr As New StreamReader(info2.FullName)
                length2 = sr.ReadToEnd.Empty
                sr.Close()
            End Using

            If length1.Length = length2.Length Then
                If length1 = length2 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        End Function

#End Region

#Region "Import Input"
        'Import inputs for all files in a directory
        Public Function ImportInputs(ByVal FileType As String, Optional ByVal excelVisible As Boolean = True) As Boolean
            Dim SAFiles As New DataTable
            SAFiles = CSVtoDatatable(New FileInfo(Me.RefFolder & "\File List.csv"))
            'If SAFiles.Columns.Count > 2 Then
            '    CreateTemplateFiles(frmMain.testIteration.Text)
            'End If

            Dim myXL As Tuple(Of Excel.Application, Boolean) = GetXlApp()
            'Item 1 = Excel application
            'Item 2 = Boolean (If true that means excel was previously open

            For Each dr As DataRow In SAFiles.Rows()
                Dim importingFrom As New FileInfo(dirUse & dr.Item("FilePath").ToString)
                If importingFrom.Extension.ToLower = ".xlsm" Then
                    Dim importingTo As New FileInfo(dirUse & dr.Item(FileType).ToString.Replace("[ITERATION]", "Iteration " & iteration))
                    Dim macroname As String = "Import_Previous_Version"
                    Dim prefix As String = ""
                    Dim params As Tuple(Of String, String, Boolean) = New Tuple(Of String, String, Boolean)(importingFrom.FullName.ToString, importingFrom.TemplateVersion, True)

                    If importingTo.Name.ToLower.Contains("pile") Then
                        macroname = "Button173_Click"
                    ElseIf importingTo.Name.ToLower.Contains("drilled pier") Then
                        If FileType = "MaestroPath" Or FileType = "ManualPath" Then
                            macroname += "_Performer"
                        End If
                    ElseIf importingTo.Name.ToLower.Contains("leg reinforcement") Then
                        If FileType = "MaestroPath" Or FileType = "ManualPath" Then
                            prefix = "m_"
                        End If
                    End If

                    If Import_Previous_Version(myXL.Item1, importingTo, macroname, params, excelVisible, prefix) Then
                        LogActivity("INFO | Import Inputs Completed for: " & importingTo.FullName)
                        'If FileType = "MaestroPath" Then
                        '    Try
                        '        Dim manFile As New FileInfo(Me.ManFolder & "\" & importingTo.Name)
                        '        manFile.Delete()
                        '        importingTo.CopyTo(manFile.FullName)
                        '    Catch ex As Exception
                        '    End Try
                        'End If
                    Else
                        LogActivity("WARNING | Import Inputs NOT Completed for: " & importingTo.FullName)
                    End If
                End If
            Next

            DisposeXlApp(myXL.Item1, myXL.Item2)
        End Function

        'Create or get the excel application to use.
        Public Function GetXlApp() As Tuple(Of Excel.Application, Boolean)
            Try
                Return New Tuple(Of Excel.Application, Boolean)(GetObject(, "Excel.Appliction"), True)
            Catch ex As Exception
                Return New Tuple(Of Excel.Application, Boolean)(CreateObject("Excel.Application"), False)
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

        'Seb's macro runner adjusted specifically for unit testing
        Public Function Import_Previous_Version(ByVal xlapp As Excel.Application,
                                                ByVal workbookFile As FileInfo,
                                                ByVal macroName As String,
                                                ByVal params As Tuple(Of String, String, Boolean), 'Item1 = Filepath, Item2 = Version, Item3 = IsMaesting
                                                Optional ByVal xlVisibility As Boolean = False,
                                                Optional ByVal prefix As String = ""
                                                ) As Boolean

            Dim toolFileName As String = Path.GetFileName(workbookFile.Name)
            Dim xlWorkBook As Excel.Workbook = Nothing
            Dim errorMessage As String = ""
            Dim isSuccess As Boolean = True

            If workbookFile Is Nothing Or String.IsNullOrEmpty(macroName) Then
                LogActivity("ERROR | workbookFile or macroName parameter is null or empty")
                Return False
            End If

            Try
                If workbookFile.Exists Then

                    xlapp.Visible = xlVisibility
                    xlWorkBook = xlapp.Workbooks.Open(workbookFile.FullName)

                    LogActivity("DEBUG | Tool: " & toolFileName)
                    LogActivity("DEBUG | BEGIN MACRO: " & macroName)

                    'Check that the strings aren't empty and that ismaesting = true
                    If params.Item1 IsNot Nothing And params.Item2 IsNot Nothing And params.Item3 Then
                        xlapp.Run(prefix & "Import_Previous_Version." & macroName, params.Item1, params.Item2, params.Item3)
                        LogActivity("DEBUG | END MACRO:  " & macroName)
                    Else
                        LogActivity("ERROR | Parameters not specific ")
                        LogActivity("DEBUG | Tool: " & toolFileName & " failed to import inputs")
                        isSuccess = False
                    End If

                    xlWorkBook.Save()
                Else
                    LogActivity("ERROR | " & workbookFile.FullName & " path not found!")
                End If
            Catch ex As Exception
                errorMessage = ex.Message
                LogActivity("ERROR | " & ex.Message)
                isSuccess = False
            Finally
                Try
                    If xlWorkBook IsNot Nothing Then
                        xlWorkBook.Close(True)
                        Marshal.ReleaseComObject(xlWorkBook)
                        xlWorkBook = Nothing
                    End If
                Catch ex As Exception
                    LogActivity("WARNING | Could not close Excel Workbook: " & toolFileName)
                End Try
            End Try

            Return isSuccess
        End Function

#End Region

#Region "Creating Iteration & Files"
        'Create folders required for unit testing to be conducted
        '''Maestro folder 
        '''Manual folder 
        '''Users will have the option to replace the files in the folder. 
        Public Sub CreateIteration(ByVal nextIteration As Integer, ByVal Optional isFirstTime As Boolean = False)

            Dim dirtocreate As String = dirUse & "\Test ID " & testCase & "\Iteration " & nextIteration

            If Not Directory.Exists(dirtocreate) Then
                Directory.CreateDirectory(Me.itFolder)
                LogActivity("DEBUG | Directory created: " & Me.itFolder)
                Directory.CreateDirectory(Me.MaeFolder)
                LogActivity("DEBUG | Directory created: " & Me.MaeFolder)
                Directory.CreateDirectory(Me.ManFolder)
                LogActivity("DEBUG | Directory created: " & Me.ManFolder)
                LogActivity("INFO | Iteration " & Me.iteration & " has been created.")
                testIteration.Text = nextIteration
                testNextIteration.Text = nextIteration + 1
            Else
                LogActivity("WARNING | Iteration " & Me.iteration.ToString & " folders already exist")
                If nextIteration = testNextIteration.Text Then
                    testIteration.Text = nextIteration
                    testNextIteration.Text = nextIteration + 1
                End If
            End If

            'testFolder.Text = "R:\Development\SAPI Testing\Unit Testing\Test ID " & testCase
            'If GetReferenceFileCount() = 0 Then
            '    FirstTimeWarning(isFirstTime, nextIteration)
            'Else
            '''Create the directories
            '''Increase the iteration (Should be at 0 if this is the first time)
            '''get all required files for testing
            'CreateTemplateFiles(isFirstTime, False)
            'End If

        End Sub

        'Folder option for this one since it could send ERIs to multiple locations. 
        Public Sub CreateManualERI(ByRef refFiles As DataTable, ByVal folder As String, Optional ByVal isFirstTime As Boolean = False, Optional ByVal archive As Boolean = True)
            If archive And Not isFirstTime Then DoArchiving(folder)

            For Each dr As DataRow In refFiles.Rows
                Dim file As New FileInfo(dirUse & dr.Item("FilePath").ToString)
                'All ERIs welcome
                'ERI is copied to the:
                '''Manual ERI Folder
                If file.Extension.Contains("eri") Then
                    file.CopyTo(folder & "\" & file.Name)
                    LogActivity("DEBUG | ERI created: " & folder & "\" & file.Name)
                End If
            Next
        End Sub

        'No folder option for these since they are always going to go to the current folder path
        Public Sub CreateCurrentTemplates(ByRef refFiles As DataTable, Optional ByVal isFirstTime As Boolean = False, Optional ByVal archive As Boolean = True)
            If archive And Not isFirstTime Then DoArchiving(Me.PubFolder)

            For Each dr As DataRow In refFiles.Rows
                Dim file As New FileInfo(dirUse & dr.Item("FilePath").ToString)
                If file.Extension.ToLower = ".xlsm" Then
                    With WhichFile(file)
                        If .Item1 Is Nothing Or .Item2 Is Nothing Or .Item3 Is Nothing Then
                            TemplateNotFoundWarning(file)
                        Else
                            'Copy published versions of the tools into the manual folder 
                            Dim pubPath As String = GetNewFileName(Me.PubFolder, fileName:= .Item3)
                            IO.File.WriteAllBytes(pubPath, .Item1)
                            dr.Item("PublishedPath") = pubPath.Replace(dirUse, "").Replace("Iteration " & iteration, "[ITERATION]")
                            LogActivity("DEBUG | Production version created: " & pubPath)
                        End If
                    End With
                End If
            Next
        End Sub

        'No folder option for these since they are always going to go to the current folder path
        Public Sub CreateSAPITemplates(ByRef refFiles As DataTable, ByVal dirtouse As String, ByVal dtHeader As String, Optional ByVal isFirstTime As Boolean = False, Optional ByVal archive As Boolean = True)
            If archive And Not isFirstTime Then
                DoArchiving(dirtouse)
            End If

            For Each dr As DataRow In refFiles.Rows
                Dim file As New FileInfo(dirUse & dr.Item("FilePath").ToString)
                If file.Extension.ToLower = ".xlsm" Then
                    With WhichFile(file)
                        If .Item1 Is Nothing Or .Item2 Is Nothing Or .Item3 Is Nothing Then
                            TemplateNotFoundWarning(file)
                        Else
                            'Templates are saved as Bytes() and need to be converted appropriately. 
                            Dim mypath As String = GetNewFileName(dirtouse, fileName:= .Item3)
                            IO.File.WriteAllBytes(mypath, .Item2)
                            dr.Item(dtHeader) = mypath.Replace(dirUse, "").Replace("Iteration " & iteration, "[ITERATION]")
                            LogActivity("DEBUG | SAPI version created: " & mypath)

                            '''File will be copied to the manual folder once the files are populated with data via 
                            '''Manual files will be replaces when user imports data into the maestro files.
                            '''Alternative will be to load maestro and manual files manually.
                            '''''Import Inputs
                            '''''Structure import
                            ''Dim manPath As String = GetNewFileName(Me.ManFolder, fileName:= .Item3)
                            ''IO.File.WriteAllBytes(manPath, .Item2)
                            ''dr.Item("ManualPath") = manPath
                            ''LogActivity("DEBUG | SAPI version created: " & manPath)
                        End If
                    End With
                End If
            Next
        End Sub

        'General Warning *salute* that the xlsm file in the reference folder could not be determined as a general template *salute* and the user should do it manually
        Public Sub TemplateNotFoundWarning(ByVal file As FileInfo)
            MsgBox("Could not determine template file type for file: " & vbCrLf & file.Name & vbCrLf & vbCrLf & "Please copy template manually.", vbCritical, "Template Not Found")
            LogActivity("WARNING | Could not determine template file type for file: " & file.Name)
            LogActivity("WARNING | Please copy template manually.")
        End Sub

        'Get count of files in the refernce file folder
        Public Function GetReferenceFileCount() As Integer
            Dim fileCount As Integer = Directory.GetFiles(Me.RefFolder).Count

            Return fileCount
        End Function

        'If the file count of the files in the reference SA folder = 0 no and it is not the first time: 
        '''Users may not continue because they have not copied over SA reference files yet.
        Public Sub FirstTimeWarning(ByVal isFirstTime As Boolean, ByVal nextIteration As Integer)
            If Not isFirstTime Then MsgBox("Files Do Not exist In the 'Reference SA Files' folder yet. Please copy reference files to continue.", vbCritical, "No Reference Files")
            LogActivity("ERROR | Files Do Not exist In the 'Reference SA Files' folder yet. Please copy reference files to continue.")
            'testIteration.Text = nextIteration - 1
            'testNextIteration.Text = nextIteration
        End Sub

#End Region

#Region "Results"
        'Return a datatable of summarized results from  a selected file
        'Invalid files return blank datatables
        Public Function SummarizedResults(ByVal info As IO.FileInfo) As DataTable
            Dim myTemplate As Tuple(Of Byte(), Byte(), String, String, String) = WhichFile(info)
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
                                        Common.ExcelDatasourceToDataTable(
                                            Common.GetExcelDataSource(
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
                finalDt.Columns.Add("Type", Type.GetType("System.String"))
                finalDt.Columns.Add("Rating", Type.GetType("System.String"))
                finalDt.Columns.Add("Tool", Type.GetType("System.String"))

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

        'Compares the results of all results available
        Public Function CompareResults() As Tuple(Of Tuple(Of Boolean, DataTable), Tuple(Of Boolean, DataTable), Tuple(Of Boolean, DataTable), DataSet, Tuple(Of Boolean, DataTable), Tuple(Of Boolean, DataTable))
            Dim manToMae As Tuple(Of Boolean, DataTable) 'Item1
            Dim curToMan As Tuple(Of Boolean, DataTable) 'Item2
            Dim curToMae As Tuple(Of Boolean, DataTable) 'Item3
            Dim edstoCur As Tuple(Of Boolean, DataTable) 'Item5
            Dim edsToMae As Tuple(Of Boolean, DataTable) 'Item6
            Dim resDs As New DataSet 'Item4

            Dim dir As String = IIf(CType(chkWorkLocal.Checked, Boolean) = True, lFolder, rFolder)
            Dim testid As Integer = CType(testCase, Integer)
            Dim testiteration As Integer = CType(Me.testIteration.Text, Integer)

            GetAllResults(Me.RefFolder)
            GetAllResults(Me.PubFolder)
            GetAllResults(Me.MaeFolder)
            GetAllResults(Me.ManFolder)

            Dim refDt As DataTable = CSVtoDatatable(New FileInfo(Me.RefFolder & "\Summarized Results.csv"))
            Dim curDt As DataTable = CSVtoDatatable(New FileInfo(Me.PubFolder & "\Summarized Results.csv"))
            Dim manDt As DataTable = CSVtoDatatable(New FileInfo(Me.ManFolder & "\Summarized Results.csv"))
            Dim maeDt As DataTable = CSVtoDatatable(New FileInfo(Me.MaeFolder & "\Summarized Results.csv"))
            Dim comDt As DataTable = New DataTable("Combined Results")

            comDt.Columns.Add("Type", Type.GetType("System.String"))
            comDt.Columns.Add("Rating", Type.GetType("System.String"))
            comDt.Columns.Add("Tool", Type.GetType("System.String"))
            comDt.Columns.Add("Summary Type", Type.GetType("System.String"))

            refDt.ResultsSorting("Reference SA")
            curDt.ResultsSorting("Published Versions")
            manDt.ResultsSorting("Manual")
            maeDt.ResultsSorting("Maestro")

            manToMae = manDt.IsMatching(maeDt)
            curToMan = curDt.IsMatching(manDt)
            curToMae = curDt.IsMatching(maeDt)

            resDs.Tables.Add(refDt.Copy)
            resDs.Tables.Add(curDt.Copy)
            resDs.Tables.Add(manDt.Copy)
            resDs.Tables.Add(maeDt.Copy)

            If Directory.Exists(EDSFolder) Then
                GetAllResults(EDSFolder)
                Dim edsDt As DataTable = CSVtoDatatable(New FileInfo(EDSFolder & "\Summarized Results.csv"))
                edsDt.ResultsSorting("EDS")
                edstoCur = edsDt.IsMatching(curDt)
                edsToMae = edsDt.IsMatching(maeDt)
                resDs.Tables.Add(edsDt.Copy)
                resDs.Tables.Add(edstoCur.Item2.Copy)
                resDs.Tables.Add(edsToMae.Item2.Copy)
            Else
                edstoCur = Nothing
                edsToMae = Nothing
            End If

            For Each dt As DataTable In resDs.Tables
                comDt.Merge(dt)
            Next

            resDs.Tables.Add(comDt.Copy)
            resDs.Tables.Add(manToMae.Item2.Copy)
            resDs.Tables.Add(curToMan.Item2.Copy)
            resDs.Tables.Add(curToMae.Item2.Copy)

            Return New Tuple(Of
                        Tuple(Of Boolean, DataTable),
                        Tuple(Of Boolean, DataTable),
                        Tuple(Of Boolean, DataTable),
                        DataSet,
                        Tuple(Of Boolean, DataTable),
                        Tuple(Of Boolean, DataTable)
                       )(
                        curToMan,
                        curToMae,
                        manToMae,
                        resDs,
                        edstoCur,
                        edsToMae
                        )
        End Function

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
            LogActivity("DEBUG | Results output for reference SA files created: " & folder & "\Summarized Results.csv")
        End Sub


        Private Sub CreatetnxResults()
            'Add ability to loop through the following folders
            '''ERI
            '''MAN
            '''MAE
            '''CUR
            'Create a datatable for each of them. (Just compare them all. 
            Dim myERIs As New List(Of String)
            Dim myTnxs As New List(Of tnxModel)
            Dim myDTs As New List(Of DataTable)
            Dim myChecks As New List(Of Tuple(Of Boolean, DataTable))
            Dim tnxDS As New DataSet

            myERIs.AddERIs(Me.EriFolder, False)
            myERIs.AddERIs(Me.PubFolder, False)
            myERIs.AddERIs(Me.ManFolder, False)
            myERIs.AddERIs(Me.MaeFolder, False)
            If Directory.Exists(EDSFolder) Then myERIs.AddERIs(EDSFolder, False)

            For Each eri In myERIs
                If File.Exists(eri & ".XMLOUT.xml") Then
                    myTnxs.Add(New tnxModel(eri, Nothing))
                End If
            Next

            For Each tnx In myTnxs
                myDTs.Add(tnx.ResultsToDataTable)
            Next

            For Each dt In myDTs
                dt.ResultsSorting
            Next

            If myDTs.Count > 1 Then
                For i As Integer = 0 To myDTs.Count - 1
                    For j As Integer = i + 1 To myDTs.Count - 1
                        myChecks.Add(myDTs(i).IsMatching(myDTs(j)))
                    Next
                    myDTs(i).ToCSV(dirUse & "\Test ID " & testCase & myDTs(i).TableName & ".csv")
                    tnxDS.Tables.Add(myDTs(i))
                Next

                'Create a tnx Comparison for the iteration
                'Save tnx Comparison dts to this folder
                'Load up the userform that shows comparisons.
                If myChecks.Count > 0 Then


                    Dim dir As String
                    Dim filepath As String
                    dir = dirUse & "\Test ID " & testCase & "\Iteration " & iteration & "\TNX Results Summary " & Now.ToString("MM/dd/yyyy HH:mm:ss tt").ToDirectoryString()
                    If Not IO.Directory.Exists(dir) Then IO.Directory.CreateDirectory(dir)

                    For Each chk In myChecks
                        filepath = dir & "\" & chk.Item2.TableName.ToDirectoryString & ".csv"
                        Try
                            chk.Item2.ToCSV(filepath)
                        Catch ex As Exception
                        End Try
                        tnxDS.Tables.Add(chk.Item2)
                    Next
                End If

                Dim newSum As New frmSummary
                newSum.myDs = tnxDS
                newSum.ToolStripSplitButton1.Enabled = False
                newSum.Show()
            End If

        End Sub

#End Region

#Region "Other Methods"
        Private Sub SetStructureToPropertyGrid(ByVal str As EDSStructure, ByVal pgrid As PropertyGrid)
            'Allow the user to view the opbjects created in the strlocal object
            pgrid.SelectedObject = str
            LogActivity("DEBUG | " & str.EDSObjectFullName & " Set to " & pgrid.Name)
        End Sub

        'A datatable of the reference files in the Reference SA Files folder. 
        Public Function RefernceSADT() As DataTable
            Dim SAFiles As New DataTable
            SAFiles = CSVtoDatatable(New FileInfo(Me.RefFolder & "\File List.csv"))
            If SAFiles.Columns.Count < 4 Then
                SAFiles.Columns.Add("MaestroPath", GetType(System.String))
                SAFiles.Columns.Add("ManualPath", GetType(System.String))
                SAFiles.Columns.Add("PublishedPath", GetType(System.String))
            End If

            'If a file was added manually
            '''it needs to add that to the filelist csv
            If Not GetReferenceFileCount() - 1 = SAFiles.Rows.Count Then
                For Each file As FileInfo In New DirectoryInfo(Me.RefFolder).GetFiles
                    Dim isFound As Boolean = False
                    For Each dr As DataRow In SAFiles.Rows
                        If dirUse & dr.Item("FilePath").ToString = file.FullName Then
                            isFound = True
                            Exit For
                        End If
                    Next

                    If Not isFound Then
                        If Not file.FullName.Contains("~") Then
                            With WhichFile(file)
                                If file.Extension.ToLower = ".eri" Or .Item1 IsNot Nothing Or .Item2 IsNot Nothing Or .Item3 IsNot Nothing Then
                                    LogActivity("DEBUG | File added manually: " & file.Name)
                                    If SAFiles.Columns.Count < 4 Then
                                        SAFiles.Rows.Add(file.FullName.Replace(dirUse, ""), file.TemplateVersion, "")
                                    Else
                                        SAFiles.Rows.Add(file.FullName.Replace(dirUse, ""), file.TemplateVersion, "", "", "", "")
                                    End If
                                End If
                            End With
                        End If
                    End If
                Next
            End If
            Return SAFiles
        End Function

        'Create a directory for unit testing. 
        '''Creates a directory locally and in the network location.
        Public Sub DirectoryCreator(ByVal subFolder As String)
            'Create R drive directory for folder
            If Not Directory.Exists(rFolder & subFolder) Then
                Directory.CreateDirectory(rFolder & subFolder)
            End If

            'Create local directory for folder
            If chkWorkLocal.Checked Then
                If Not Directory.Exists(lFolder & subFolder) Then
                    Directory.CreateDirectory(lFolder & subFolder)
                End If
            End If
        End Sub

        'Creates a structure object based on the files in the maestro folder for the current iteration
        Public Sub CreateStructure(ByVal filesPath As String, Optional ByVal deleteAdditionalTNXfiles As Boolean = True)
            Dim myFiles As String()
            Dim myFilesLst As New List(Of String)

            'default resonse to determine if a question needs asked.
            'Dim response As DialogResult = DialogResult.Cancel

            'Loop through all files in the maestro folder for the current test case and iteration
            For Each info As FileInfo In New DirectoryInfo(filesPath).GetFiles
                If info.Extension = ".eri" Then
                    'All eris permitted
                    myFilesLst.Add(info.FullName)
                    LogActivity("DEBUG | File found for structure: " & info.Name)
                ElseIf info.Extension = ".xlsm" Then 'All tools are current xlsm files and this should be a safe assumption
                    'Determine if the file is one of the templates
                    Dim template As Tuple(Of Byte(), Byte(), String, String, String) = WhichFile(info)

                    'If the properties of the tuple are nothing then they aren't templates
                    If template.Item1 IsNot Nothing And template.Item2 IsNot Nothing And template.Item3 IsNot Nothing Then
                        myFilesLst.Add(info.FullName)
                        LogActivity("DEBUG | File found for structure: " & info.Name)
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
                        LogActivity("DEBUG | File Deleted: " & info.FullName)
                    End If
                    'End If
                End If
            Next

            'Convert the list of valid file names to an array for creating anew structure
            myFiles = myFilesLst.ToArray
            strcLocal = New EDSStructure(bus_unit, structure_id, work_order_seq_num, filesPath, filesPath, myFiles, EDSnewId, EDSdbActive)
        End Sub

        'Loads the CSV with the test cases 
        'CSV is saved here: R:\Development\SAPI Testing
        Public Sub LoadTestCases(ByRef lst As List(Of TestCase))
            'Read csv saved in R drive location
            Using csvReader As New Microsoft.VisualBasic.FileIO.TextFieldParser("\\netapp4\cad\Development\SAPI Testing\Unit Test Cases.csv")
                csvReader.TextFieldType = FileIO.FieldType.Delimited
                csvReader.SetDelimiters(",")
                Dim csvValue As String()

                While Not csvReader.EndOfData
                    'Create a list of test cases based on the CSV data
                    csvValue = csvReader.ReadFields
                    lst.Add(New TestCase(csvValue))
                End While
            End Using
        End Sub

        Public Sub SetTestIDLabels()
            If Directory.Exists(lFolder) Then
                For Each fold As DirectoryInfo In New DirectoryInfo(lFolder).GetDirectories
                    If fold.Name.Contains("Test ID") Then
                        Dim i As Integer
                        i = CType(Split(fold.Name, " ")(2), Integer)
                        If File.Exists(fold.FullName & "\Checked Out.txt") Then
                            testID.Properties.Items(i - 1) = i & "|Checked Out"
                        Else
                            testID.Properties.Items(i - 1) = i
                        End If
                    End If
                Next
            End If
        End Sub

        'Determines the file name for the new templates being saved.
        'Increments file names if they arleady exist in the new directory.
        Public Function GetNewFileName(ByVal newFolder As String, ByVal Optional file As FileInfo = Nothing, ByVal Optional fileName As String = Nothing) As String
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
        Public Function WhichFile(ByVal file As FileInfo) As Tuple(Of Byte(), Byte(), String, String, String)
            Dim returner As Tuple(Of Byte(), Byte(), String, String, String)
            'Item 1 = current published versions
            'Item 2 = new versions created for SAPI
            'Item 3 = File name to be used with the bytes
            'Item 4 = Worksheet with results
            'Item 5 = Range for results 

            If file.Name.ToLower.Contains("cciplate") Then
                returner = New Tuple(Of Byte(), Byte(), String, String, String)(
                Testing_Winform.My.Resources.CCIplate__4_1_2_,
                CCI_Engineering_Templates.My.Resources.CCIplate,
                "CCIplate.xlsm",
                "Results Database",
                "B1:BO64")
            ElseIf file.Name.ToLower.Contains("ccipole") Then
                returner = New Tuple(Of Byte(), Byte(), String, String, String)(
                Testing_Winform.My.Resources.CCIpole__4_5_8_,
                CCI_Engineering_Templates.My.Resources.CCIpole,
                "CCIpole.xlsm",
                "Results",
                "AZ4:BT108")
            ElseIf file.Name.ToLower.Contains("cciseismic") Then
                returner = New Tuple(Of Byte(), Byte(), String, String, String)(
                Testing_Winform.My.Resources.CCISeismic__3_3_9_,
                CCI_Engineering_Templates.My.Resources.CCISeismic,
                "CCISeismic.xlsm",
                Nothing,
                Nothing)
            ElseIf file.Name.ToLower.Contains("drilled pier") Then
                returner = New Tuple(Of Byte(), Byte(), String, String, String)(
                Testing_Winform.My.Resources.Drilled_Pier_Foundation__5_0_5_,
                CCI_Engineering_Templates.My.Resources.Drilled_Pier_Foundation,
                "Drilled Pier Foundation.xlsm",
                "Foundation Input",
                "BD8:CF59|H10:L31")
            ElseIf file.Name.ToLower.Contains("guyed anchor") Then
                returner = New Tuple(Of Byte(), Byte(), String, String, String)(
                Testing_Winform.My.Resources.Guyed_Anchor_Block_Foundation__4_0_0_,
                CCI_Engineering_Templates.My.Resources.Guyed_Anchor_Block_Foundation,
                "Guyed Anchor Block Foundation.xlsm",
                "Input",
                "M20:X70")
            ElseIf file.Name.ToLower.Contains("leg reinforcement") Then
                returner = New Tuple(Of Byte(), Byte(), String, String, String)(
                Testing_Winform.My.Resources.Leg_Reinforcement_Tool__10_0_4_,
                CCI_Engineering_Templates.My.Resources.Leg_Reinforcement_Tool,
                "Leg Reinforcement Tool.xlsm",
                Nothing,
                Nothing)
            ElseIf file.Name.ToLower.Contains("pier and pad") Then
                returner = New Tuple(Of Byte(), Byte(), String, String, String)(
                Testing_Winform.My.Resources.Pier_and_Pad_Foundation__4_1_1_,
                CCI_Engineering_Templates.My.Resources.Pier_and_Pad_Foundation,
                "Pier and Pad Foundation.xlsm",
                "Input",
                "F12:K25")
            ElseIf file.Name.ToLower.Contains("pile") Then
                returner = New Tuple(Of Byte(), Byte(), String, String, String)(
                Testing_Winform.My.Resources.Pile_Foundation__2_2_1_,
                CCI_Engineering_Templates.My.Resources.Pile_Foundation,
                "Pile Foundation.xlsm",
                "Input",
                "G13:M31")
            ElseIf file.Name.ToLower.Contains("unit base") Then
                returner = New Tuple(Of Byte(), Byte(), String, String, String)(
                Testing_Winform.My.Resources.SST_Unit_Base_Foundation__4_0_3_,
                CCI_Engineering_Templates.My.Resources.SST_Unit_Base_Foundation,
                "SST Unit Base Foundation.xlsm",
                "Input",
                "F12:K24")
            Else
                returner = New Tuple(Of Byte(), Byte(), String, String, String)(
                Nothing,
                Nothing,
                Nothing,
                Nothing,
                Nothing)
            End If

            Return returner
        End Function

#End Region

#End Region
    End Class

End Namespace