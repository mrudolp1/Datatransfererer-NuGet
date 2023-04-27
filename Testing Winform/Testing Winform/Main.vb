Imports System.ComponentModel
Imports System.Text
Imports CCI_Engineering_Templates
Imports System.Data.SqlClient
Imports System.Security.Principal
Imports System.IO
Imports Oracle.ManagedDataAccess.Client
Imports RoboSharp
Imports System.Threading
Imports System.Data.OleDb
Imports System.Runtime.CompilerServices

Namespace UnitTesting

    Partial Public Class frmMain
        Public strcLocal As EDSStructure
        Public strcEDS As EDSStructure

#Region "Object Declarations"
        'Public myUnitBases As New DataTransfererUnitBase
        'Public myPierandPads As New DataTransfererPierandPad
        Public myDrilledPiers As New DrilledPierFoundation
        Public myGuyedAnchorBlocks As New AnchorBlockFoundation
        'Public myPiles As New DataTransfererPile
        'Public MyCCIpoles As New DataTransfererCCIpole
        'Public MyCCIplates As New DataTransfererCCIplate

        Public BUNumber As String = ""
        Public StrcID As String = ""
        Public WorkOrder As String = ""

        'Import to Excel
        Public ListOfFilesCopied As New List(Of String) From {"C:\Users\" & Environment.UserName & "\Desktop\C Drive Testing\Drilled Pier\EDS\Test Sites\809534 - MP\Drilled Pier Foundation (5.1.0.3)_EDS_3.xlsm"}
        'Public ListOfFilesCopied As New List(Of String) From {"C:\Users\" & Environment.UserName & "\Desktop\Save to Excel\EDS - 806889 - Pier and Pad Foundation (4.1.2).xlsm"}
        'Public ListOfFilesCopied As New List(Of String) From {"C:\Users\" & Environment.UserName & "\Desktop\Pile Foundation\VB.Net Test Cases\Test Cases\EDS - 800011 - Pile Foundation (2.2.1.6).xlsm"}
        'Public ListOfFilesCopied As New List(Of String) From {"C:\Users\" & Environment.UserName & "\Desktop\Pile Foundation\VB.Net Test Cases\EDS - 800009 - Drilled Pier Foundation (5.1.0).xlsm"}
        'Public ListOfFilesCopied As New List(Of String) From {"C:\Users\" & Environment.UserName & "\Documents\.NET Testing\EDS - 806956 -SST Unit Base Foundation (4.0.4) - from EDS.xlsm"}
        'Public ListOfFilesCopied As New List(Of String) From {"C:\Users\" & Environment.UserName & "\Documents\.NET Testing\EDS - 811236 - Pile Foundation (2.2.1.6).xlsm"}
        'Public ListOfFilesCopied As New List(Of String) From {"C:\Users\" & Environment.UserName & "\Documents\.NET Testing\EDS - 846182 - Guyed Anchor Block Foundation (4.1.0).xlsm"}
        'Public ListOfFilesCopied As New List(Of String) From {"C:\Users\" & Environment.UserName & "\Desktop\C Drive Testing\CCIpole\EDS Testing\Test Sites\800476\CCIpole (4.6.0) - 1 - EDS.xlsm"}
        'Public ListOfFilesCopied As New List(Of String) From {"C:\Users\" & Environment.UserName & "\Documents\.NET Testing\EDS - 812637 - CCIplate (4.1.2.1).xlsm"}

        'Import to EDS
        Public ListOfExcelFiles As New List(Of String) From {"C:\Users\" & Environment.UserName & "\Desktop\C Drive Testing\Drilled Pier\EDS\Test Sites\809534 - MP\Drilled Pier Foundation (5.1.0.3)_2.xlsm"}
        'Public ListOfExcelFiles As New List(Of String) From {"C:\Users\" & Environment.UserName & "\Desktop\Save to EDS\806889 - Pier and Pad Foundation (4.1.2).xlsm"}
        'Public ListOfExcelFiles As New List(Of String) From {"C:\Users\" & Environment.UserName & "\Desktop\Pile Foundation\VB.Net Test Cases\Test Cases\800011 - Pile Foundation (2.2.1.6).xlsm"}
        'Public ListOfExcelFiles As New List(Of String) From {"C:\Users\" & Environment.UserName & "\Desktop\Pile Foundation\VB.Net Test Cases\Test Cases\800009 - Drilled Pier Foundation (5.1.0) - TEMPLATE - 8-27-2021.xlsm"}
        'Public ListOfExcelFiles As New List(Of String) From {"C:\Users\" & Environment.UserName & "\Documents\.NET Testing\Foundations\Unit Base\806956\EDS - 806956 -SST Unit Base Foundation (4.0.4) - from EDS - Change1.xlsm"}
        'Public ListOfExcelFiles As New List(Of String) From {"C:\Users\" & Environment.UserName & "\Documents\.NET Testing\Foundations\Unit Base\806956\806956 SST Unit Base Foundation (4.0.4) - TEMPLATE.xlsm"}
        'Public ListOfExcelFiles As New List(Of String) From {"C:\Users\" & Environment.UserName & "\Desktop\Pile Foundation\VB.Net Test Cases\Test Cases\800009 - Guyed Anchor Block Foundation (4.1.0) - TEMPLATE - 9-9-2021.xlsm"}
        'Public ListOfExcelFiles As New List(Of String) From {"C:\Users\" & Environment.UserName & "\Documents\.NET Testing\Foundations\Guyed Anchor Block\846182\846182 Guyed Anchor Block Foundation (4.1.0) - TEMPLATE - 11-2-2021.xlsm"}
        'Public ListOfExcelFiles As New List(Of String) From {"C:\Users\" & Environment.UserName & "\Documents\.NET Testing\Foundations\Pile\811236\811236 Pile Foundation (2.2.1.6).xlsm"}
        'Public ListOfExcelFiles As New List(Of String) From {"C:\Users\" & Environment.UserName & "\Documents\.NET Testing\Foundations\Pile\811236\EDS - 811236 - Pile Foundation (2.2.1.6) - Change 1.xlsm"}
        'Public ListOfExcelFiles As New List(Of String) From {"C:\Users\" & Environment.UserName & "\Documents\.NET Testing\Foundations\Guyed Anchor Block\846182\EDS - 846182 - Guyed Anchor Block Foundation (4.1.0) - Change 1.xlsm"}
        'Public ListOfExcelFiles As New List(Of String) From {"C:\Users\" & Environment.UserName & "\Desktop\C Drive Testing\CCIpole\EDS Testing\Test Sites\800476\CCIpole (4.6.0) - 0.xlsm"}
        'Public ListOfExcelFiles As New List(Of String) From {"C:\Users\" & Environment.UserName & "\Documents\.NET Testing\Connection\812637\CCIplate (4.1.2.1) - Copy - Copy.xlsm"}

        'Public EDSdbDevelopment As String = "Server=DEVCCICSQL2.US.CROWNCASTLE.COM,60113;Database=EDSDev;Integrated Security=SSPI"
        'Public EDSuserDevelopment As String = "366:204:303:354:207:330:309:207:204:249"
        'Public EDSuserPwDevelopment As String = "210:264:258:99:297:303:213:258:246:318:354:111:345:168:300:318:261:219:303:267:246:300:108:165:144:192:324:153:246:300"
        'Changed to Uat for testing database changes
        'Public EDSdbDevelopment As String = "Server=DEVCCICSQL2.US.CROWNCASTLE.COM,60113;Database=EDSUat;Integrated Security=SSPI"
        'Public EDSuserDevelopment As String = "366:204:303:354:207:330:309:207:204:249"
        'Public EDSuserPwDevelopment As String = "210:264:258:99:297:303:213:258:246:318:354:111:345:168:300:318:261:219:303:267:246:300:108:165:144:192:324:153:246:300"
        Public EDSdbDevelopment As String = "Server=DEVCCICSQL3.US.CROWNCASTLE.COM,58061;Database=EngEDSDev;Integrated Security=SSPI"
        Public EDSuserDevelopment As String = "366:204:303:354:207:330:309:207:204:249"
        Public EDSuserPwDevelopment As String = "210:264:258:99:297:303:213:258:246:318:354:111:345:168:300:318:261:219:303:267:246:300:108:165:144:192:324:153:246:300"

        Public EDSdbProduction As String = "Server=CCICSQLCLST2.US.CROWNCASTLE.COM,64540;Database=EDSProd;Integrated Security=SSPI"
        Public EDSuserProduction As String = "366:207:330:309:207:204:249"
        Public EDSuserPwProduction As String = "147:267:297:216:168:297:270:357:234:282:225:156:114:216:147:321:111:144:168:156:168:333:222:258:366:171:126:342:252:147"

        Public EDSdbActive As String
        Public EDSuserActive As String
        Public EDSuserPwActive As String
        Public EDStokenHandle As New IntPtr(0)
        Public EDSimpersonatedUser As WindowsImpersonationContext
        Public EDSnewId As WindowsIdentity

        Public sqlCon As New SqlConnection
        Public ds As New DataSet
        Public da As SqlDataAdapter
        Public dt As New DataTable
        Public sql As String

        Public Declare Auto Function LogonUser Lib "advapi32.dll" (ByVal nToken As String, ByVal domain As String, ByVal wToken As String, ByVal lType As Integer, ByVal lProvider As Integer, ByRef Token As IntPtr) As Boolean
        Public Declare Auto Function CloseHandle Lib "kernel32.dll" (ByVal handle As IntPtr) As Boolean
        Public tokenHandle As New IntPtr(0)
#End Region

#Region "Form Handlers"
        Public Sub New()
            InitializeComponent()
        End Sub
        Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
            isopening = True
            If My.Settings.serverActive = "dbDevelopment" Then
                EDSdbActive = EDSdbDevelopment
                EDSuserActive = EDSuserDevelopment
                EDSuserPwActive = EDSuserPwDevelopment
            Else
                EDSdbActive = EDSdbProduction
                EDSuserActive = EDSuserProduction
                EDSuserPwActive = EDSuserPwProduction
            End If

            If Environment.UserName.ToLower = "imiller" Or Environment.UserName.ToLower = "stanley" Or Environment.UserName.ToLower = "dsmilowitz" Then
                txtFndBU.Text = My.Settings.myBU
                txtFndStrc.Text = My.Settings.myStrID
                txtFndWO.Text = My.Settings.myWO
                txtDirectory.Text = My.Settings.myWorkArea
                If My.Settings.localWorkArea = String.Empty Then
                    My.Settings.localWorkArea = "C:\Users\" & Environment.UserName & "\source"
                    My.Settings.Save()
                End If
                testLocalWorkarea.Text = My.Settings.localWorkArea
                lFolder = My.Settings.localWorkArea
                chkWorkLocal.Checked = My.Settings.workLocal
            Else
                txtFndBU.Text = "800000"
                txtFndStrc.Text = "A"
                txtFndWO.Text = "1234567"
                txtDirectory.Text = "C:\SAPI Work Area\Test"
            End If

            LogonUser(token(EDSuserActive), "CCIC", token(EDSuserPwActive), 2, 0, EDStokenHandle)
            EDSnewId = New WindowsIdentity(EDStokenHandle)
            isopening = False
        End Sub

        Private Sub frmMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
            CloseHandle(EDStokenHandle)
            DirectorySync.Stop()
        End Sub
#End Region

#Region "Structure"
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
            strcLocal.SavetoEDS(EDSnewId, EDSdbActive)
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
            Dim tempstr As New EDSStructure

            For Each dir As FileInfo In New DirectoryInfo(txtDirectory.Text).GetFiles
                If dir.Extension = "" Then
                    For Each file As FileInfo In New DirectoryInfo(dir.FullName).GetFiles
                        If file.Extension.ToLower = ".eri" Then
                            tempstr.RunTNX(file.FullName)
                        End If
                    Next
                End If
            Next
        End Sub
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

            strcLocal.Conduct(True)

        End Sub
#End Region

#Region "Unit Testing"
        Public unitTestCases As New List(Of TestCase)
        Public rFolder As String = "R:\Development\SAPI Testing\Unit Testing"
        Public lFolder As String
        Public thr1 As Thread
        Public DirectorySync As RoboCommand = New RoboCommand()

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
            Catch
            End Try
        End Sub

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
            If testID.Text = String.Empty Then Exit Sub

            'Check local isn't really being used anymore since I forced it to work local. 
            'The local directory MUST be specified. It is currently defaulting to your source folder outside your repos folder
            If chkWorkLocal.Checked And lFolder = String.Empty Then
                MsgBox("Please specify a local directory to continue.", vbCritical, "No Local Directory")
                isopening = True
                testID.SelectedIndex = -1
                isopening = False
                Exit Sub
            End If

            If unitTestCases.Count = 0 Then
                LoadTestCases(unitTestCases)
            End If

            ButtonclickToggle(Me.Cursor)

            Dim id As Integer = testID.Text - 1
            Dim testCase As Integer = testID.Text
            Dim dirUse As String

            'Again, the check local isn't important as it is forced into local work at this time.
            If chkWorkLocal.Checked Then
                dirUse = lFolder
            Else
                dirUse = rFolder
            End If

            'if the directory exists on the R drive but not locally
            '''Copy the directory locally since you're probably continuing work done by someone else. 
            If Directory.Exists(rFolder & "\Test ID " & testCase) And chkWorkLocal.Checked And Not Directory.Exists(lFolder & "\Test ID " & testCase) Then
                My.Computer.FileSystem.CopyDirectory(rFolder & "\Test ID " & testCase, lFolder & "\Test ID " & testCase)
            End If

            'Create the initial directory
            ''' the direcotrycreator method creates it locally and on the network
            If Not Directory.Exists(dirUse & "\Test ID " & testCase) Then
                DirectoryCreator("\Test ID " & testCase)
                DirectoryCreator("\Test ID " & testCase & "\Manual (Current)")
                DirectoryCreator("\Test ID " & testCase & "\Reference SA Files")
                DirectoryCreator("\Test ID " & testCase & "\Manual ERI")
                File.Create(dirUse & "\Test ID " & testCase & "\Test Notes.txt").Dispose()

                'When first creating the test case folder general notes (Salute) will be created to get started. 
                If rtbNotes.Text.Length = 0 Then
                    rtbNotes.Text = "Testing notes for Test ID " & testCase
                    rtbNotes.Text += vbCrLf & "BU = " & unitTestCases(id).BU
                    rtbNotes.Text += vbCrLf & "Structure ID = " & unitTestCases(id).SID
                    rtbNotes.Text += vbCrLf & "Wo = " & unitTestCases(id).WO
                    rtbNotes.Text += vbCrLf & "SA Work Area = " & unitTestCases(id).SAWorkArea
                    rtbNotes.Text += vbCrLf & "Load Combination = " & unitTestCases(id).COMB
                End If

                'Create the 1st iteration folder. 
                '''Specifying that this is the first time (The boolean in the called method) 
                CreateIteration(1, True)
            Else
                'If the directory exists, it just loads in the text file for reference
                rtbNotes.Text = System.IO.File.ReadAllText(dirUse & "\Test ID " & testCase & "\Test Notes.txt")
            End If

            'Start file sinking....drip drip drip into the R drive
            InitializeLocaltoCentralSync()
            'Attempted to thread to save time but turns out it is just because of the network connection issues at home
            thr1 = New Thread(AddressOf DirectorySync.StartAsync)
            thr1.Start()

            'Set the site data loaded from the test case CSV.
            testBu.Text = unitTestCases(id).BU
            testSid.Text = unitTestCases(id).SID
            testWo.Text = unitTestCases(id).WO
            testSaFolder.Text = unitTestCases(id).SAWorkArea 'This will update the directory for the SA Reference folder
            testFolder.Text = "R:\Development\SAPI Testing\Unit Testing\Test ID " & testCase 'This will update the directory for the network test case
            testComb.Text = unitTestCases(id).COMB

            'Iteration count is determined
            '''A count of folders containing the word 'iteration' are counted
            Dim itCount As Integer = 0
            For Each subDir In New DirectoryInfo("R:\Development\SAPI Testing\Unit Testing\Test ID " & testCase).GetDirectories
                If subDir.Name.Contains("Iteration ") Then itCount += 1
            Next

            testIteration.Text = itCount
            testNextIteration.Text = itCount + 1

            'Enable all of the buttons for use in the iteration
            btnNextIteration.Enabled = True
            testIterationResults.Enabled = True
            testPrevResults.Enabled = True
            testPublishedResults.Enabled = True
            testConduct.Enabled = True
            testCompareAll.Enabled = True

            'Update the local directory to the local test case. 
            Try
                seLocal.SetCurrentDirectory(dirUse & "\Test ID " & testCase)
            Catch
            End Try
            ButtonclickToggle(Me.Cursor)
        End Sub

        'Create a new iteration button click
        Private Sub btnNextIteration_Click(sender As Object, e As EventArgs) Handles btnNextIteration.Click
            ButtonclickToggle(Me.Cursor)
            CreateIteration(testNextIteration.Text)
            ButtonclickToggle(Me.Cursor)
        End Sub

        'Rich textbox changed event for test notes
        Private Sub rtbNotes_TextChanged(sender As Object, e As EventArgs)
            Dim testCase As Integer = testID.Text
            Dim dirUse As String
            If chkWorkLocal.Checked Then
                dirUse = lFolder
            Else
                dirUse = rFolder
            End If

            Try
                System.IO.File.WriteAllText(dirUse & "\Test ID " & testCase & "\Test Notes.txt", rtbNotes.Text)
            Catch
            End Try
        End Sub

        'work local or remote option 
        Private Sub CheckEdit1_CheckedChanged(sender As Object, e As EventArgs) Handles chkWorkLocal.CheckedChanged
            If isopening Then Exit Sub
            My.Settings.workLocal = sender.checked
            My.Settings.Save()
        End Sub

        'Conduct button click for current iteration
        Private Sub testConduct_Click(sender As Object, e As EventArgs) Handles testConduct.Click
            ButtonclickToggle(Me.Cursor)
            Dim iteration As Integer = testIteration.Text
            Dim testcase As Integer = testID.Text
            Dim maeWorkArea As String = lFolder & "\Test ID " & testcase & "\Iteration " & iteration & "\Maestro"
            Dim myFiles As String()
            Dim myFilesLst As New List(Of String)

            'Loop through all files in the maestro folder for the current test case and iteration
            For Each info As FileInfo In New DirectoryInfo(maeWorkArea).GetFiles
                If info.Extension = ".eri" Then
                    'All eris permitted
                    myFilesLst.Add(info.FullName)
                ElseIf info.Extension = ".xlsm" Then 'All tools are current xlsm files and this should be a safe assumption
                    'Determine if the file is one of the templates
                    Dim template As Tuple(Of FileInfo, Byte(), String, String, String) = WhichFile(info)

                    'If the properties of the tuple are nothing then they aren't templates
                    If template.Item1 IsNot Nothing And template.Item2 IsNot Nothing And template.Item3 IsNot Nothing Then
                        myFilesLst.Add(info.FullName)
                    End If
                End If
            Next

            'Convert the list of valid file names to an array for creating anew structure
            myFiles = myFilesLst.ToArray
            strcLocal = New EDSStructure(testBu.Text, testSid.Text, testWo.Text, maeWorkArea, maeWorkArea, myFiles, EDSnewId, EDSdbActive)

            'Conduct it!!!
            '''This is commented out since Seb is actively working on the conduct function
            'strcLocal.Conduct(True)

            'Allow the user to view the opbjects created in the strlocal object
            pgcUnitTesting.SelectedObject = strcLocal
            ButtonclickToggle(Me.Cursor)
        End Sub

        'Create and compare CSV Results files
        Private Sub testPrevResults_Click(sender As Object, e As EventArgs) Handles testPrevResults.Click
            ButtonclickToggle(Me.Cursor)
            GetAllResults(lFolder & "\Test ID " & testID.Text & "\Reference SA Files")
            ButtonclickToggle(Me.Cursor)
        End Sub
        Private Sub testPublishedResults_Click(sender As Object, e As EventArgs) Handles testPublishedResults.Click
            ButtonclickToggle(Me.Cursor)
            GetAllResults(lFolder & "\Test ID " & testID.Text & "\Manual (Current)")
            ButtonclickToggle(Me.Cursor)
        End Sub
        Private Sub testIterationResults_Click(sender As Object, e As EventArgs) Handles testIterationResults.Click
            ButtonclickToggle(Me.Cursor)
            GetAllResults(lFolder & "\Test ID " & testID.Text & "\Iteration " & testIteration.Text & "\Maestro")
            GetAllResults(lFolder & "\Test ID " & testID.Text & "\Iteration " & testIteration.Text & "\Manual (SAPI)")
            ButtonclickToggle(Me.Cursor)
        End Sub
        Private Sub testCompareAll_Click(sender As Object, e As EventArgs) Handles testCompareAll.Click
            ButtonclickToggle(Me.Cursor)
            Dim checks As Tuple(Of Tuple(Of Boolean, DataTable), Tuple(Of Boolean, DataTable), Tuple(Of Boolean, DataTable), DataSet) = CompareResults()
            ButtonclickToggle(Me.Cursor)

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

            Dim newSum As New frmSummary
            newSum.myDs = checks.Item4
            newSum.Show()
        End Sub


        'Custom Methods
        Public Sub CreateIteration(ByVal Iteration As Integer, ByVal Optional isFirstTime As Boolean = False)
            'Determine which directory to use. 
            Dim dirUse As String
            If chkWorkLocal.Checked Then
                dirUse = lFolder
            Else
                dirUse = rFolder
            End If

            'Set the directories to reference based on working local or on the network.
            Dim testCase As Integer = testID.Text
            Dim MaeFolder As String = dirUse & "\Test ID " & testCase & "\Iteration " & Iteration & "\Maestro"
            Dim ManFolder As String = dirUse & "\Test ID " & testCase & "\Iteration " & Iteration & "\Manual (SAPI)"
            Dim PubFolder As String = dirUse & "\Test ID " & testCase & "\Manual (Current)"
            Dim RefFolder As String = dirUse & "\Test ID " & testCase & "\Reference SA Files"
            Dim EriFolder As String = dirUse & "\Test ID " & testCase & "\Manual ERI"

            'Get a file count of all files in the:
            '''SA reference folder
            '''Published tool folder
            '''ERI Reference folder
            Dim fileCount As Integer = Directory.GetFiles(RefFolder).Count
            Dim publishedFileCount As Integer = Directory.GetFiles(PubFolder).Count
            Dim eriFileCount As Integer = Directory.GetFiles(EriFolder).Count

            If fileCount = 0 Then
                'If the file count is no and it is not the first time: 
                '''Users may not continue because they have not copied over SA reference files yet. 
                If Not isFirstTime Then MsgBox("Files do not exist in the 'Reference SA Files' folder yet. Please copy reference files to continue.", vbCritical, "No Reference Files")
                testIteration.Text = Iteration - 1
                testNextIteration.Text = Iteration
            Else
                'If the SA Reference file count is > 0 then:
                '''Create the directories
                '''Increase the iteration (Should be at 0 if this is the first time)
                '''get all required files for testing
                Directory.CreateDirectory("R:\Development\SAPI Testing\Unit Testing\Test ID " & testCase & "\Iteration " & Iteration)
                Directory.CreateDirectory(MaeFolder)
                Directory.CreateDirectory(ManFolder)

                testIteration.Text = Iteration
                testNextIteration.Text = Iteration + 1
                testFolder.Text = "R:\Development\SAPI Testing\Unit Testing\Test ID " & testCase

                'Loop through all files in the SA Reference Files Folder
                For Each file In New DirectoryInfo(RefFolder).GetFiles
                    'All ERIs welcome
                    'ERI is copied to the:
                    '''Manual ERI Folder
                    '''Mae Folder
                    If file.Extension.Contains("eri") Then
                        file.CopyTo(MaeFolder & "\" & file.Name)
                        If eriFileCount = 0 Then file.CopyTo(EriFolder & "\" & file.Name)
                    Else
                        If file.Extension.ToLower = ".eri" Or file.Extension.ToLower = ".xlsm" Then
                            'Determine if the file is a template
                            Dim myTemplate As Tuple(Of FileInfo, Byte(), String, String, String) = WhichFile(file)

                            'If it is determined to be a template:
                            '''The published version will be copied into the published tools folder
                            '''The new SAPI templates will be copied into the mae folder and man folder for the iteration
                            With myTemplate
                                If .Item1 Is Nothing Or .Item2 Is Nothing Or .Item3 Is Nothing Then
                                    MsgBox("Could not determine template file type for file: " & vbCrLf & file.Name & vbCrLf & vbCrLf & "Please copy template manually.", vbCritical, "Template Not Found")
                                Else
                                    If publishedFileCount = 0 Then
                                        'Copy published versions of the tools into the manual folder 
                                        .Item1.CopyTo(GetNewFileName(PubFolder, fileName:= .Item3))
                                    End If

                                    'Templates are saved as Bytes() and need to be converted appropriately. 
                                    IO.File.WriteAllBytes(GetNewFileName(MaeFolder, fileName:= .Item3), .Item2)
                                    IO.File.WriteAllBytes(GetNewFileName(ManFolder, fileName:= .Item3), .Item2)

                                End If
                            End With
                        End If
                    End If
                Next
            End If
        End Sub
        Public Sub InitializeLocaltoCentralSync()
            If DirectorySync.IsRunning Then
                Return
            End If

            Dim testCase As Integer = testID.Text

            DirectorySync.CopyOptions.Source = lFolder & "\Test ID " & testCase
            DirectorySync.CopyOptions.Destination = rFolder & "\Test ID " & testCase
            DirectorySync.CopyOptions.CopySubdirectories = True
            DirectorySync.CopyOptions.UseUnbufferedIo = True
            DirectorySync.CopyOptions.MultiThreadedCopiesCount = 4
            DirectorySync.CopyOptions.CopySubdirectoriesIncludingEmpty = True
            DirectorySync.CopyOptions.Purge = True
            DirectorySync.CopyOptions.MonitorSourceChangesLimit = 3
            DirectorySync.CopyOptions.MonitorSourceTimeLimit = 5
            DirectorySync.RetryOptions.RetryCount = 1
            DirectorySync.RetryOptions.RetryWaitTime = 2
        End Sub
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
#End Region

    End Class

    Public Module MyLargelyLittleHelpers
        Public isopening As Boolean

        'Loads the CSV with the test cases 
        'CSV is saved here: R:\Development\SAPI Testing
        Public Sub LoadTestCases(ByRef lst As List(Of TestCase))
            'Read csv saved in R drive location
            Using csvReader As New Microsoft.VisualBasic.FileIO.TextFieldParser("R:\Development\SAPI Testing\Unit Test Cases.csv")
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

        'Gets a combined datatbale of results for all spreadsheets in directory.
        Public Sub GetAllResults(ByVal folder As String)
            Dim combinedResults As New DataTable
            'Loop through all files in the specified folder
            For Each info As FileInfo In New DirectoryInfo(folder).GetFiles
                If info.Extension.ToLower = ".xlsm" Then
                    'Merge the datatable to append all data together
                    combinedResults.Merge(SummarizedResults(info))
                End If
            Next

            'Save the datatable to a CSV in the specified folder location
            DatatableToCSV(combinedResults, folder & "\Summarized Results.csv")
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
                If file Is Nothing Then
                    filePath = newFolder & "\" & file.Name.Split(".")(0) & "(" & counter.ToString() & ")" & file.Name.Split(".")(1)
                Else
                    filePath = newFolder & "\" & fileName.Split(".")(0) & "(" & counter.ToString() & ")" & fileName.Split(".")(1)
                End If
            End While

            Return filePath
        End Function

        'Used to determine which template is being used
        'This could have been set up as a class but ended up going too far and now we have tuples. Enjoy! :)
        Public Function WhichFile(ByVal file As FileInfo) As Tuple(Of FileInfo, Byte(), String, String, String)
            Dim returner As Tuple(Of FileInfo, Byte(), String, String, String)

            'This templatesfolder needs to be customized if your username doesn't match your user folder or if your engineering templates are synced to a different location
            Dim templatesFolder As String = "C:\Users\" & Environment.UserName & "\Crown Castle USA Inc\Tower Assets Engineering - Engineering Templates\"

            'Item 1 = current published versions
            'Item 2 = new versions created for SAPI
            'Item 3 = File name to be used with the bytes
            'Item 4 = Worksheet with results
            'Item 5 = Range for results 

            If file.Name.ToLower.Contains("cciplate") Then
                returner = New Tuple(Of FileInfo, Byte(), String, String, String)(
                New FileInfo(templatesFolder & "CCIPlate\CCIplate (4.1.2).xlsm"),
                CCI_Engineering_Templates.My.Resources.CCIplate,
                "CCIplate.xlsm",
                "Results Database",
                "B1:BO64")
            ElseIf file.Name.ToLower.Contains("ccipole") Then
                returner = New Tuple(Of FileInfo, Byte(), String, String, String)(
                New FileInfo(templatesFolder & "CCIPole\CCIpole (4.5.8).xlsm"),
                CCI_Engineering_Templates.My.Resources.CCIpole,
                "CCIpole.xlsm",
                "Results",
                "AZ4:BT108")
            ElseIf file.Name.ToLower.Contains("cciseismic") Then
                returner = New Tuple(Of FileInfo, Byte(), String, String, String)(
                New FileInfo(templatesFolder & "CCISeismic\CCISeismic (3.3.9).xlsm"),
                CCI_Engineering_Templates.My.Resources.CCISeismic,
                "CCISeismic.xlsm",
                Nothing,
                Nothing)
            ElseIf file.Name.ToLower.Contains("drilled pier") Then
                returner = New Tuple(Of FileInfo, Byte(), String, String, String)(
                New FileInfo(templatesFolder & "Foundation Templates\Unified Foundation Tools\Drilled Pier Foundation (5.0.5).xlsm"),
                CCI_Engineering_Templates.My.Resources.Drilled_Pier_Foundation,
                "Drilled Pier Foundation.xlsm",
                "Foundation Input",
                "BD8:CF59|H10:L31")
            ElseIf file.Name.ToLower.Contains("guyed anchor") Then
                returner = New Tuple(Of FileInfo, Byte(), String, String, String)(
                New FileInfo(templatesFolder & "Foundation Templates\Unified Foundation Tools\Guyed Anchor Block Foundation (4.0.0).xlsm"),
                CCI_Engineering_Templates.My.Resources.Guyed_Anchor_Block_Foundation,
                "Guyed Anchor Block Foundation.xlsm",
                "Input",
                "M20:X70")
            ElseIf file.Name.ToLower.Contains("leg reinforcement.xlsm") Then
                returner = New Tuple(Of FileInfo, Byte(), String, String, String)(
                New FileInfo(templatesFolder & "Foundation Templates\Unified Foundation Tools\Guyed Anchor Block Foundation (4.0.0).xlsm"),
                CCI_Engineering_Templates.My.Resources.Leg_Reinforcement_Tool,
                "Leg Reinforcement Tool.xlsm",
                Nothing,
                Nothing)
            ElseIf file.Name.ToLower.Contains("pier and pad") Then
                returner = New Tuple(Of FileInfo, Byte(), String, String, String)(
                New FileInfo(templatesFolder & "Half Pipe Leg Modification\Leg Reinforcement Tool (10.0.4).xlsm"),
                CCI_Engineering_Templates.My.Resources.Pier_and_Pad_Foundation,
                "Pier and Pad Foundation.xlsm",
                "Input",
                "F12:K25")
            ElseIf file.Name.ToLower.Contains("pile") Then
                returner = New Tuple(Of FileInfo, Byte(), String, String, String)(
                New FileInfo(templatesFolder & "Foundation Templates\Unified Foundation Tools\Pile Foundation (2.2.1).xlsm"),
                CCI_Engineering_Templates.My.Resources.Pile_Foundation,
                "Pile Foundation.xlsm",
                "Input",
                "G13:M31")
            ElseIf file.Name.ToLower.Contains("unit base") Then
                returner = New Tuple(Of FileInfo, Byte(), String, String, String)(
                New FileInfo(templatesFolder & "Foundation Templates\Unified Foundation Tools\SST Unit Base Foundation (4.0.3).xlsm"),
                CCI_Engineering_Templates.My.Resources.SST_Unit_Base_Foundation,
                "SST Unit Base Foundation.xlsm",
                "Input",
                "F12:K24")
            Else
                returner = New Tuple(Of FileInfo, Byte(), String, String, String)(
                Nothing,
                Nothing,
                Nothing,
                Nothing,
                Nothing)
            End If

            Return returner
        End Function

        'Return a datatable of summarized results from  a selected file
        'Invalid files return blank datatables
        Public Function SummarizedResults(ByVal info As IO.FileInfo) As DataTable
            Dim myTemplate As Tuple(Of IO.FileInfo, Byte(), String, String, String) = WhichFile(info)
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
                                If i > 31 Then addl = "_Seismic"

                                If Not dr.Item("Plate Summary").ToString = String.Empty And Not dr.Item("Plate Summary").ToString = "Max Stress" Then
                                    Dim val As String

                                    'bolt group 1
                                    If Not dr.Item("Bolt GR. 1").ToString = String.Empty Then
                                        val = dr.Item("Column21").ToString.Replace("%", "")
                                        finalDt.Rows.Add("Plate " & dr.Item("Flange ID").ToString & "_" & "Bolt Group 1" & addl, val, info.Name.Replace(".xlsm", ""))
                                    End If

                                    'bolt group 2
                                    If Not dr.Item("Bolt GR. 2").ToString = String.Empty Then
                                        val = dr.Item("Column31").ToString.Replace("%", "")
                                        finalDt.Rows.Add("Plate " & dr.Item("Flange ID").ToString & "_" & "Bolt Group 2" & addl, val, info.Name.Replace(".xlsm", ""))
                                    End If

                                    'bolt group 3
                                    If Not dr.Item("Bolt GR. 3").ToString = String.Empty Then
                                        val = dr.Item("Column41").ToString.Replace("%", "")
                                        finalDt.Rows.Add("Plate " & dr.Item("Flange ID").ToString & "_" & "Bolt Group 3" & addl, val, info.Name.Replace(".xlsm", ""))
                                    End If

                                    'bolt group 4
                                    If Not dr.Item("Bolt GR. 4").ToString = String.Empty Then
                                        val = dr.Item("Column51").ToString.Replace("%", "")
                                        finalDt.Rows.Add("Plate " & dr.Item("Flange ID").ToString & "_" & "Bolt Group 4" & addl, val, info.Name.Replace(".xlsm", ""))
                                    End If

                                    'bolt group 5
                                    If Not dr.Item("Bolt GR. 5").ToString = String.Empty Then
                                        val = dr.Item("Column61").ToString.Replace("%", "")
                                        finalDt.Rows.Add("Plate " & dr.Item("Flange ID").ToString & "_" & "Bolt Group 5" & addl, val, info.Name.Replace(".xlsm", ""))
                                    End If

                                    'Plate stress
                                    val = dr.Item("Plate").ToString.Replace("%", "")
                                    finalDt.Rows.Add("Plate " & dr.Item("Flange ID").ToString & "_" & dr.Item("Column63").ToString & addl, val, info.Name.Replace(".xlsm", ""))
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
                                        ancVal = 0
                                    End Try
                                    Try
                                        strVal = dr.Item("Anchor Rating") * 100
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
                                        val = dr.Item("Rating*").ToString.Replace("%", "")
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
                    MsgBox(strFilePath & " is currently open. " & vbCrLf & vbCrLf & "Please close the file to continue.", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "File is in use")

                    counter += 1
                    If counter > 2 Then
                        MsgBox("It seems the file is still open." & vbCrLf & vbCrLf & "Data was not saved to CSV.", vbInformation)
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
        Public Function CSVtoDatatable(ByVal info As FileInfo, Optional ByVal hasHeaders As Boolean = True) As DataTable
            Dim dssample As New DataSet
            Dim folder = info.FullName.Replace(info.Name, "")
            Dim CnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & folder & ";Extended Properties=""text;HDR=No;FMT=Delimited"";"

            Using Adp As New OleDbDataAdapter("select * from [" & info.Name & "]", CnStr)
                Adp.Fill(dssample)
            End Using

            If hasHeaders Then
                For Each dc As DataColumn In dssample.Tables(0).Columns
                    'If the data is not saved as a string then it will not recognize the header column as it doesn't not assume headers in the SQL query
                    'I didn't have time to create custom queries. 
                    'This just works for any selected csv file
                    Try
                        dc.ColumnName = dssample.Tables(0).Rows(0).Item(dc)
                    Catch
                    End Try
                Next

                dssample.Tables(0).Rows.Remove(dssample.Tables(0).Rows(0))
            End If

            If dssample.Tables.Count > 0 Then
                'Only 1 table should have been output but it returns that table
                Return dssample.Tables(0)
            End If
        End Function

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

        'Compares the results of all results available
        Public Function CompareResults() As Tuple(Of Tuple(Of Boolean, DataTable), Tuple(Of Boolean, DataTable), Tuple(Of Boolean, DataTable), DataSet)
            Dim manToMae As Tuple(Of Boolean, DataTable) 'Item1
            Dim curToMan As Tuple(Of Boolean, DataTable) 'Item2
            Dim curToMae As Tuple(Of Boolean, DataTable) 'Item2
            Dim resDs As New DataSet 'Item4

            Dim dir As String = IIf(CType(frmMain.chkWorkLocal.Checked, Boolean) = True, frmMain.lFolder, frmMain.rFolder)
            Dim testid As Integer = CType(frmMain.testID.Text, Integer)
            Dim testiteration As Integer = CType(frmMain.testIteration.Text, Integer)

            GetAllResults(dir & "\Test ID " & testid & "\Reference SA Files")
            GetAllResults(dir & "\Test ID " & testid & "\Manual (Current)")
            GetAllResults(dir & "\Test ID " & testid & "\Iteration " & testiteration & "\Maestro")
            GetAllResults(dir & "\Test ID " & testid & "\Iteration " & testiteration & "\Manual (SAPI)")

            Dim refDt As DataTable = CSVtoDatatable(New FileInfo(dir & "\Test ID " & testid & "\Reference SA Files\Summarized Results.csv"))
            Dim curDt As DataTable = CSVtoDatatable(New FileInfo(dir & "\Test ID " & testid & "\Manual (Current)\Summarized Results.csv"))
            Dim manDt As DataTable = CSVtoDatatable(New FileInfo(dir & "\Test ID " & testid & "\Iteration " & testiteration & "\Manual (SAPI)\Summarized Results.csv"))
            Dim maeDt As DataTable = CSVtoDatatable(New FileInfo(dir & "\Test ID " & testid & "\Iteration " & testiteration & "\Maestro\Summarized Results.csv"))
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
                        DataSet
                       )(
                        curToMan,
                        curToMae,
                        manToMae,
                        resDs
                        )
        End Function

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
            dt.TableName = addColumn
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
            diffDt.Columns.Add(dt.TableName & " Val", GetType(System.Double))
            diffDt.Columns.Add(comparer.TableName & " Val", GetType(System.Double))
            diffDt.Columns.Add("Delta", GetType(System.Double))
            diffDt.Columns.Add("% Difference", GetType(System.Double))
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
                        If dr.Item("Type").ToString = dtRow.Item("Type").ToString And dr.Item("Tool").ToString = dtRow.Item("Tool").ToString Then
                            comparerRow = dr
                            Exit For
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

                diffDt.Rows.Add(
                                dtRow.Item("Tool").ToString,
                                dtRow.Item("Type").ToString,
                                IIf(Double.IsNaN(dtVal), Nothing, dtVal),
                                IIf(Double.IsNaN(comparerVal), Nothing, comparerVal),
                                IIf(Double.IsNaN(delta), Nothing, delta),
                                IIf(Double.IsNaN(perDelta), Nothing, perDelta),
                                IIf(Double.IsNaN(delta) Or delta > 0.1, "Fail", "Pass")
                               )

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

        'Toggles the cursor between default and waiting
        '''Placed at the beginning and end of form events like button clicks or checkbox changes
        '''Should also be placed anywhere you exit sub 
        Public Sub ButtonclickToggle(ByRef cur As Cursor)
            If cur = Cursors.WaitCursor Then
                cur = Cursors.Default
            Else
                cur = Cursors.WaitCursor
            End If
        End Sub

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
End Namespace


'UNSUED
#Region "Original Excel"


'Public Sub CreateExcelTemplates() Handles sqltoexcel.Click
'    'ClearAllTools()

'    BUNumber = txtSQLBU.Text
'    StrcID = txtSQLStrc.Text

'    Dim xlFndGroup As New EDSFoundationGroup()

'    For Each item As String In ListOfFilesCopied
'        If item.Contains("SST Unit Base Foundation") Then
'            myUnitBases = New DataTransfererUnitBase(ds, EDSnewId, EDSdbActive, BUNumber, StrcID)
'            myUnitBases.ExcelFilePath = item
'            If myUnitBases.LoadFromEDS() Then myUnitBases.SaveToExcel()
'        ElseIf item.Contains("Pier and Pad Foundation") Then
'            myPierandPads = New DataTransfererPierandPad(ds, EDSnewId, EDSdbActive, BUNumber, StrcID)
'            myPierandPads.ExcelFilePath = item
'            If myPierandPads.LoadFromEDS() Then myPierandPads.SaveToExcel()
'        ElseIf item.Contains("Drilled Pier Foundation") Then
'            myDrilledPiers = New DataTransfererDrilledPier(ds, EDSnewId, EDSdbActive, BUNumber, StrcID)
'            myDrilledPiers.ExcelFilePath = item
'            If myDrilledPiers.LoadFromEDS() Then myDrilledPiers.SaveToExcel()
'        ElseIf item.Contains("Guyed Anchor Block Foundation") Then
'            myGuyedAnchorBlocks = New DataTransfererGuyedAnchorBlock(ds, EDSnewId, EDSdbActive, BUNumber, StrcID)
'            myGuyedAnchorBlocks.ExcelFilePath = item
'            If myGuyedAnchorBlocks.LoadFromEDS() Then myGuyedAnchorBlocks.SaveToExcel()
'        ElseIf item.Contains("Pile Foundation") Then
'            myPiles = New DataTransfererPile(ds, EDSnewId, EDSdbActive, BUNumber, StrcID)
'            myPiles.ExcelFilePath = item
'            If myPiles.LoadFromEDS() Then myPiles.SaveToExcel()
'        ElseIf item.Contains("CCIpole") Then
'            MyCCIpoles = New DataTransfererCCIpole(ds, EDSnewId, EDSdbActive, BUNumber, StrcID)
'            MyCCIpoles.ExcelFilePath = item
'            If MyCCIpoles.LoadFromEDS() Then MyCCIpoles.SaveToExcel()
'        ElseIf item.Contains("CCIplate") Then
'            MyCCIplates = New DataTransfererCCIplate(ds, EDSnewId, EDSdbActive, BUNumber, StrcID)
'            MyCCIplates.ExcelFilePath = item
'            'If MyCCIplates.LoadFromEDS() Then MyCCIplates.SaveToExcel()
'        End If
'    Next

'End Sub

'Public Sub UploadExcelFilesToEDS() Handles exceltosql.Click
'    'ClearAllTools()

'    BUNumber = txtSQLBU.Text
'    StrcID = txtSQLStrc.Text

'    Dim xlFd As New OpenFileDialog
'    xlFd.Multiselect = True
'    xlFd.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"

'    If xlFd.ShowDialog = DialogResult.OK Then
'        Dim xlFndGroup As New EDSFoundationGroup()

'        For Each item As String In xlFd.FileNames
'            If item.Contains("SST Unit Base Foundation") Then
'                myUnitBases = New DataTransfererUnitBase(ds, EDSnewId, EDSdbActive, BUNumber, StrcID)
'                myUnitBases.ExcelFilePath = item
'                myUnitBases.LoadFromExcel()
'                myUnitBases.SaveToEDS()
'            ElseIf item.Contains("Pier and Pad Foundation") Then
'                'myPierandPads = New DataTransfererPierandPad(ds, EDSnewId, EDSdbActive, BUNumber, StrcID)
'                'myPierandPads.ExcelFilePath = item
'                'myPierandPads.LoadFromExcel()
'                'myPierandPads.SaveToEDS()

'                xlFndGroup.PierandPads.Add(New PierAndPad(item))

'            ElseIf item.Contains("Drilled Pier Foundation") Then
'                myDrilledPiers = New DataTransfererDrilledPier(ds, EDSnewId, EDSdbActive, BUNumber, StrcID)
'                myDrilledPiers.ExcelFilePath = item
'                myDrilledPiers.LoadFromExcel()
'                myDrilledPiers.SaveToEDS()
'            ElseIf item.Contains("Guyed Anchor Block Foundation") Then
'                myGuyedAnchorBlocks = New DataTransfererGuyedAnchorBlock(ds, EDSnewId, EDSdbActive, BUNumber, StrcID)
'                myGuyedAnchorBlocks.ExcelFilePath = item
'                myGuyedAnchorBlocks.LoadFromExcel()
'                myGuyedAnchorBlocks.SaveToEDS()
'            ElseIf item.Contains("Pile Foundation") Then
'                myPiles = New DataTransfererPile(ds, EDSnewId, EDSdbActive, BUNumber, StrcID)
'                myPiles.ExcelFilePath = item
'                myPiles.LoadFromExcel()
'                myPiles.SaveToEDS()
'            ElseIf item.Contains("CCIpole") Then
'                MyCCIpoles = New DataTransfererCCIpole(ds, EDSnewId, EDSdbActive, BUNumber, StrcID)
'                MyCCIpoles.ExcelFilePath = item
'                MyCCIpoles.LoadFromExcel()
'                MyCCIpoles.SaveToEDS()
'            ElseIf item.Contains("CCIplate") Then
'                MyCCIplates = New DataTransfererCCIplate(ds, EDSnewId, EDSdbActive, BUNumber, StrcID)
'                MyCCIplates.ExcelFilePath = item
'                MyCCIplates.LoadFromExcel()
'                MyCCIplates.SaveToEDS()
'            End If
'        Next

'        'Compare excel foundations to the foundations in the active model for this BU and Str
'        'This will copy IDs on matching foundations and the whole foundation group if all the foundations match
'        xlFndGroup.CompareMe(New EDSFoundationGroup(BUNumber, StrcID, EDSnewId, EDSdbActive), True)

'        'Save all foundations, anything with an ID won't be uploaded, if the whole foundation group has an ID, nothing has changed
'        xlFndGroup.SaveAllFoundationsEDS(EDSnewId, EDSdbActive)

'    End If

'End Sub

'Sub ClearAllTools()
'    myUnitBases.Clear()
'    'myPierandPads.Clear()
'    myDrilledPiers.Clear()
'    myGuyedAnchorBlocks.Clear()
'    myPiles.Clear()
'    MyCCIpoles.Clear()
'    MyCCIplates.Clear()
'End Sub

''Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
''    MsgBox("Stop touching me")
''End Sub

'Private Sub CreateExcelTemplates(sender As Object, e As EventArgs) Handles sqltoexcel.Click

'End Sub

'Private Sub UploadExcelFilesToEDS(sender As Object, e As EventArgs) Handles exceltosql.Click

'End Sub

#End Region
#Region "tnx"

'Public tnxFromERI As tnxModel
'Public tnxFromDB As tnxModel
'Private Sub btnImportERI_Click(sender As Object, e As EventArgs) Handles btnImportERI.Click
'    Dim eriFd As New OpenFileDialog
'    eriFd.Multiselect = False
'    eriFd.Filter = "TNX File|*.eri"

'    If eriFd.ShowDialog = DialogResult.OK Then
'        tnxFromERI = New tnxModel(eriFd.FileName)

'        propgridTNXERI.SelectedObject = tnxFromERI
'    End If
'End Sub

'Private Sub btnExportERI_Click(sender As Object, e As EventArgs) Handles btnExportERI.Click
'    If tnxFromDB Is Nothing Then
'        MessageBox.Show("Import a file first.")
'        Exit Sub
'    End If

'    Dim eriFd As New SaveFileDialog
'    eriFd.Filter = "TNX File|*.eri"

'    If eriFd.ShowDialog = DialogResult.OK Then
'        tnxFromDB.GenerateERI(eriFd.FileName)
'    End If
'End Sub

'Private Sub btnSavetoEDS_Click(sender As Object, e As EventArgs) Handles btnSavetoEDS.Click
'    If txtBU.Text = "" Or txtStrc.Text = "" Or tnxFromERI Is Nothing Then Exit Sub

'    'tnxFromERI.SaveBaseToEDSInd(txtBU.Text, txtStrc.Text, EDSnewId, EDSdbActive)
'    tnxFromERI.SaveToEDS(txtBU.Text, txtStrc.Text, EDSnewId, EDSdbActive)

'End Sub

'Private Sub btnTest_Click(sender As Object, e As EventArgs) Handles btnTest.Click
'    If txtBU.Text = "" Or txtStrc.Text = "" Or tnxFromERI Is Nothing Then Exit Sub

'    'tnxFromERI.SaveBaseToEDSSub(txtBU.Text, txtStrc.Text, EDSnewId, EDSdbActive)
'    'benchmarked at 0.5 sec

'    'tnxFromERI.SaveBaseToEDSFull(txtBU.Text, txtStrc.Text, EDSnewId, EDSdbActive)
'    'benchmarked between 1.5-2.25 sec

'End Sub

'Private Sub btnLoadfromEDS_Click(sender As Object, e As EventArgs) Handles btnLoadfromEDS.Click
'    If txtBU.Text = "" Or txtStrc.Text = "" Then Exit Sub

'    tnxFromDB = New tnxModel(txtBU.Text, txtStrc.Text, EDSnewId, EDSdbActive)

'    propgridTNXEDS.SelectedObject = tnxFromDB
'End Sub

'Private Sub btnCompare_Click(sender As Object, e As EventArgs) Handles btnCompare.Click
'    If tnxFromDB Is Nothing Or tnxFromERI Is Nothing Then
'        MessageBox.Show("Import both models to compare.")
'    End If

'    Dim differences As String = ""
'    'Dim result As Boolean = tnxFromDB.CompareMe(Of tnxModel)(tnxFromERI, , differences)

'    My.Computer.Clipboard.SetText(differences)

'    'MessageBox.Show(result.ToString)

'End Sub


#End Region