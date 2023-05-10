﻿Imports System.ComponentModel
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
Imports Newtonsoft.Json
Imports CciSites.Utils.JsonUtil
Imports System.Runtime.Serialization.Json
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports DevExpress.XtraEditors
Imports DevExpress.Utils.Svg
Imports DevExpress.Utils.Drawing

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

            If unitTestCases.Count = 0 Then
                LoadTestCases(unitTestCases)
            End If

            LogonUser(token(EDSuserActive), "CCIC", token(EDSuserPwActive), 2, 0, EDStokenHandle)
            EDSnewId = New WindowsIdentity(EDStokenHandle)
            KillRoboCops()
            isopening = False

            If My.Settings.MyTestCase > 0 Then
                testID.Text = My.Settings.MyTestCase
                SetUpWorkArea(My.Settings.MyTestCase - 1)
                btnClose.Enabled = True
                btnCheckout.Enabled = False
            End If
        End Sub

        Private Sub frmMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
            CloseHandle(EDStokenHandle)
            Try
                KillRoboCops()
                DirectorySync.Stop()
            Catch ex As Exception
            End Try
        End Sub

        Private Sub KillRoboCops()
            Dim proc = Process.GetProcessesByName("RoboCopy")
            For i As Integer = 0 To proc.Count - 1
                Try
                    proc(i).Kill()
                Catch
                End Try
            Next i
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

#Region "Unit Testing - Control handlers only"

#Region "Unit Testing - Old Buttons"
        'Create bug folder
        Private Sub SimpleButton2_Click(sender As Object, e As EventArgs) Handles SimpleButton2.Click
            Dim answer = InputBox("Enter the ID of the bug from the spreadsheet.", "Bug ID", Nothing)
            If answer IsNot Nothing Then
                Directory.CreateDirectory(Me.itFolder & "\Bug Reference Files")
            End If
        End Sub

        'work local or remote option 
        Private Sub CheckEdit1_CheckedChanged(sender As Object, e As EventArgs) Handles chkWorkLocal.CheckedChanged
            If isopening Then Exit Sub
            My.Settings.workLocal = sender.checked
            My.Settings.Save()
        End Sub


        'Create a new iteration button click
        Private Sub btnNextIteration_Click(sender As Object, e As EventArgs) Handles btnNextIteration.Click
            ButtonclickToggle(Me.Cursor)
            CreateIteration(testNextIteration.Text)
            ButtonclickToggle(Me.Cursor)
        End Sub

        'Conduct button click for current iteration
        Private Sub testConduct_Click(sender As Object, e As EventArgs) Handles testConduct.Click
            ButtonclickToggle(Me.Cursor)

            CreateStructure()

            'Conduct it!!!
            '''This is commented out since Seb is actively working on the conduct function
            '''Uncommented 4-27-2023
            strcLocal.Conduct(True)
            SetStructureToPropertyGrid(strcLocal, pgcUnitTesting)

            ButtonclickToggle(Me.Cursor)
        End Sub
        'Create a json file of the lodaed structure
        Private Sub testJason_Click(sender As Object, e As EventArgs) Handles testJason.Click
            Dim strJson As String

            Try
                strJson = ToJsonString(Of EDSStructure)(strcLocal)
            Catch ex As Exception
            End Try

            Using sw As New StreamWriter(lFolder & "\Test ID " & testID.Text.ToString & "\Iteration " & testIteration.Text.ToString & "\Maestro\" & "EDSStructure_" & Now.ToString("MM/dd/yyyy HH:mm:ss tt").ToDirectoryString & ".ccistr")
                sw.Write(strJson)
                sw.Close()
            End Using
        End Sub
        Private Sub testJasonLoad_click(sender As Object, e As EventArgs) Handles testJasonLoad.Click
            Dim dateCheck As DateTime = "1/1/1900 12:00 AM"
            Dim myFile As FileInfo = Nothing

            For Each file As FileInfo In New DirectoryInfo(lFolder & "\Test ID " & testID.Text.ToString & "\Iteration " & testIteration.Text.ToString & "\Maestro\").GetFiles
                If file.Extension.ToLower = ".ccistr" Then
                    If file.CreationTime > dateCheck Then
                        dateCheck = file.CreationTime
                        myFile = file
                    End If
                End If
            Next

            If myFile IsNot Nothing Then
                Dim tempStr As New EDSStructure
                Using sr As New StreamReader(myFile.FullName)
                    tempStr = FromJsonString(Of EDSStructure)(sr.ReadToEnd)
                    sr.Close()
                End Using

                Console.WriteLine(tempStr.EDSObjectName)

                pgcUnitTesting.SelectedObject = tempStr
            End If
        End Sub

        Private Sub SimpleButton1_Click(sender As Object, e As EventArgs)
            Dim file As New FileInfo("C:\Users\Imiller\Work Area\SAPI Testing\Unit Testing\Test ID 75\Reference SA Files\Drilled Pier Foundation (5.0.3).xlsm")
            MsgBox(file.TemplateVersion)
        End Sub


        Private Sub testStructureOnly_Click(sender As Object, e As EventArgs) Handles testStructureOnly.Click
            ButtonclickToggle(Me.Cursor)

            CreateStructure()
            SetStructureToPropertyGrid(strcLocal, pgcUnitTesting)

            ButtonclickToggle(Me.Cursor)
        End Sub
        Private Sub SetStructureToPropertyGrid(ByVal str As EDSStructure, ByVal pgrid As PropertyGrid)
            'Allow the user to view the opbjects created in the strlocal object
            pgrid.SelectedObject = str
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


#End Region

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
            Catch
            End Try
        End Sub
#End Region

#Region "Button Process Clicks"

        Private Sub TestSteps(sender As Object, e As EventArgs) Handles _
                    btnProcess1.Click, btnProcess2.Click, btnProcess3.Click, btnProcess4.Click,
                    btnProcess5.Click, btnProcess6.Click, btnProcess7.Click, btnProcess8.Click,
                    btnProcess9.Click, btnProcess10.Click, btnProcess11.Click, btnProcess12.Click
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
                        Next

                        Dim newFileCsv As New DataTable
                        newFileCsv.Columns.Add("FilePath", GetType(System.String))
                        newFileCsv.Columns.Add("Version", GetType(System.String))
                        newFileCsv.Columns.Add("RPAth", GetType(System.String))
                        For Each file As FileInfo In myfilesLst
                            Dim newFile As FileInfo = file.CopyTo(Me.RefFolder & "\" & file.Name)
                            LogActivity("DEBUG | " & newFile.Name & " has been copied to the SA Reference Files Folder")
                            newFileCsv.Rows.Add(newFile.FullName, file.TemplateVersion, file.FullName)
                        Next

                        DatatableToCSV(newFileCsv, Me.RefFolder & "\File List.csv")
                        LogActivity("INFO | SA Reference Files have been copied into the test directory.")
                    End If

                Case "step2"
                    '''Create new iteration
                    CreateIteration(testNextIteration.Text)
                    LogActivity("INFO | Iteration " & Me.iteration & " has been created.")


                Case "step3", "step3a", "step3b", "step3c"
                    'Make sure all necessary files exist in the required folders
                    If Me.iteration = 0 Then
                        MsgBox("Please create an iteration to continue.", vbInformation)
                        LogActivity("ERROR | Iteration not created.")
                        Exit Select
                    End If

                    Dim eriCount As Integer = New DirectoryInfo(Me.EriFolder).GetFiles.Count
                    Dim pubCount As Integer = New DirectoryInfo(Me.PubFolder).GetFiles.Count
                    Dim maeCount As Integer = New DirectoryInfo(Me.MaeFolder).GetFiles.Count
                    Dim eriMsg As String = vbCrLf & "    Iteration " & Me.iteration & "\Manual ERI"
                    Dim pubMsg As String = vbCrLf & "    Iteration " & "\Manual (Current)"
                    Dim maeMsg As String = vbCrLf & vbCrLf & "    Iteration " & Me.iteration & "\Maestro" & vbCrLf & "    Iteration " & Me.iteration & "\Manual (SAPI)"
                    Dim msg As String = "Are you sure you would like to create new template files?" &
                                            vbCrLf & vbCrLf &
                                            "This process will archive files in the following folders:"
                    Dim answer As DialogResult = vbYes
                    Dim booEri As Boolean = IIf(eriCount = 0, True, False)
                    Dim booPub As Boolean = IIf(pubCount = 0, True, False)
                    Dim booMae As Boolean = IIf(maeCount = 0, True, False)

                    If tags(0).ToLower.Contains("a") And pubCount > 0 Then msg += pubMsg
                    If tags(0).ToLower.Contains("b") And eriCount > 0 Then msg += eriMsg
                    If tags(0).ToLower.Contains("c") And maeCount > 0 Then msg += maeMsg

                    If ((eriCount > 0 Or pubCount > 0 Or maeCount > 0) And tags(0) = "step3") Or
                        (pubCount > 0 And tags(0).ToLower = "step3a") Or
                        (eriCount > 0 And tags(0).ToLower = "step3b") Or
                        (maeCount > 0 And tags(0).ToLower = "step3c") Then
                        answer = MsgBox(msg, vbCritical + vbYesNo, "Archive Files?")
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
                            CreateSAPITemplates(refDT, booMae, Not booMae)
                            LogActivity("INFO | All files required for Maestro have been created in the directory '\Iteration" & Me.iteration & "\Maestro'.")
                            LogActivity("INFO | All files SAPI files have been created in the directory '\Iteration" & Me.iteration & "\Manual (SAPI)'.")
                            CreateManualERI(RefernceSADT, Me.MaeFolder, booMae, False)
                            LogActivity("INFO | All reference ERIs files have been created in the directory '\Iteration" & Me.iteration & "\Maestro'.")
                        End If

                        DatatableToCSV(refDT, RefFolder & "\File List.csv")
                    End If
                Case "step4", "step4a", "step4b"
                    'Create Published versions of the files
                    If Me.iteration = 0 Then
                        MsgBox("Please create an iteration to continue.", vbInformation)
                        LogActivity("ERROR | Iteration not created.")
                        Exit Select
                    End If

                    Dim pubCount As Integer = New DirectoryInfo(Me.PubFolder).GetFiles.Count
                    Dim maeCount As Integer = New DirectoryInfo(Me.MaeFolder).GetFiles.Count


                    If tags(0).ToLower = "step4" Or tags(0).ToLower = "step4b" Then
                        If Not pubCount > 0 Then
                            MsgBox("Please create current template files to continue.", vbInformation)
                            LogActivity("ERROR | Current template files not created.")
                            Exit Select
                        End If

                        ImportInputs("PublishedPath")
                        LogActivity("INFO | Import inputs complete for current published versions.")
                    End If

                    If tags(0).ToLower = "step4" Or tags(0).ToLower = "step4a" Then
                        If Not maeCount > 0 Then
                            MsgBox("Please create template SAPI files to continue.", vbInformation)
                            LogActivity("ERROR | SAPI template files not created.")
                            Exit Select
                        End If

                        ImportInputs("MaestroPath")
                        LogActivity("INFO | Import inputs complete for SAPI versions.")
                        LogActivity("INFO | Files were created in the '\Maestro' and '\Manual (SAPI)' folders.")
                    End If

                Case "step5"
                    'Run the ERI file in the Manual Reference Folder
                    Dim tempStrc As New EDSStructure
                    Dim myERIs As New List(Of String)
                    For Each info As FileInfo In New DirectoryInfo(Me.EriFolder).GetFiles
                        If info.Extension.ToLower() = ".eri" Then
                            'All eris permitted
                            myERIs.Add(info.FullName)
                            LogActivity("DEBUG | ERI: " & info.Name & " found")
                        ElseIf info.Name.ToLower.Contains(".eri.") Or info.Extension.ToLower = ".tfnx" Then
                            info.Delete()
                            LogActivity("DEBUG | File Deleted: " & info.Name & "")
                        End If
                    Next

                    For Each eri As String In myERIs
                        If Not tempStrc.RunTNX(eri, True) Then
                            LogActivity("ERROR | Failed to run ERI: " & eri)
                            GoTo finishMe
                        Else
                            LogActivity("DEBUG | ERI: " & eri & " successuflly analyzed")
                        End If
                    Next
                Case "step6"
                    'Conduct the Maestro files
                    CreateStructure()

                    'Conduct it!!!
                    '''This is commented out since Seb is actively working on the conduct function
                    '''Uncommented 4-27-2023
                    strcLocal.Conduct(True)
                    If DidConductProperly(strcLocal.LogPath) Then
                        ObjectToJson(Of EDSStructure)(strcLocal, Me.MaeFolder & "\" & "EDSStructure_" & Now.ToString("MM/dd/yyyy HH:mm:ss tt").ToDirectoryString & ".ccistr")
                        LogActivity("INFO | Structure conducted successfully.")
                    Else
                        LogActivity("ERROR | Structure NOT conducted successfully.")
                    End If
                    strcLocal.AppendLog(Me.TestLogActivityPath)
                    SetStructureToPropertyGrid(strcLocal, pgcUnitTesting)
                Case "step7"
                    Dim checks As Tuple(Of Tuple(Of Boolean, DataTable), Tuple(Of Boolean, DataTable), Tuple(Of Boolean, DataTable), DataSet) = CompareResults()
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
                    LogActivity("INFO | Results for all files compared and created in testing directory.")
            End Select

finishMe:
            LogActivity("PROCESS | End " & sender.tooltip.ToString, True)
            ButtonclickToggle(Me.Cursor, Cursors.Default)
        End Sub

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

            LogActivity("FINISH | Test Case" & Me.testCase)

            PushIt(True)
            TearDownWorkArea()
            sender.enabled = False
            btnCheckout.Enabled = True
            My.Settings.MyTestCase = 0
            My.Settings.Save()

            ButtonclickToggle(Me.Cursor)
            isopening = False
        End Sub

        Private Sub testCheckout_click(sender As Object, e As EventArgs) Handles btnCheckout.Click
            If isopening Then Exit Sub

            If IsNumeric(testID.Text) Then
                ButtonclickToggle(Me.Cursor)
                CheckOut()
                sender.enabled = False
                btnClose.Enabled = True
                My.Settings.MyTestCase = testID.Text
                My.Settings.Save()

                LogActivity("START | Test Case" & Me.testCase, True)
                LogActivity("INFO | Using CCI Engineering Datatransferer " & CCI_Engineering_Templates.myVersion)
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
                Return If(IsNumeric(testID.Text), testID.Text, Nothing)
            End Get
        End Property
        'Set the iteration to use throughout.
        Public ReadOnly Property iteration As Integer?
            Get
                Return If(IsNumeric(testIteration.Text), testIteration.Text, Nothing)
            End Get
        End Property

        Public unitTestCases As New List(Of TestCase)
        Public rFolder As String = "R:\Development\SAPI Testing\Unit Testing"
        Public lFolder As String
        Public thr1 As Thread
        Public DirectorySync As RoboCommand
#End Region

#Region "Checkin/Checkout"
        Public Sub PullIt(ByVal checkout As Boolean)
            ForceSyncing(SyncDirection.RtoLocal, True)

            If Not testPush.Enabled Then
                File.Delete(Me.dirUse & "\Test ID " & Me.testCase & "\Checked Out.txt")
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

        Public Sub CheckOut()
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
                            Exit Sub
                        End If
                    Else
                        answer = MsgBox("Test Case " & Me.testCase & " exists on the R: drive and locally." & vbCrLf & vbCrLf & "Pulling data from the R: drive will delete all local files." & vbCrLf & vbCrLf & "Would you like to PULL the R: drive files locally?", vbYesNo + vbInformation, "Pull and Checkout")
                        If answer = vbYes Then
                            PullIt(True)
                        Else
                            CreateCheckOutFiles()
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

            SetUpWorkArea(id)
        End Sub

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

        Public Sub SetUpWorkArea(ByVal listID As Integer)
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
            For Each subDir In New DirectoryInfo(Me.dirUse & "\Test ID " & Me.testCase).GetDirectories
                If subDir.Name.Contains("Iteration ") Then itCount += 1
            Next

            testIteration.Text = itCount
            testNextIteration.Text = itCount + 1

            'Update the local directory to the local test case. 
            Try
                seLocal.SetCurrentDirectory(Me.dirUse & "\Test ID " & Me.testCase)
            Catch ex As Exception
            End Try

            ResetControls()
            mainLogViewer.ReloadActivityLog()
        End Sub

        Public Sub TearDownWorkArea()
            ResetControls()

            'Set the values of all inputs
            'This can't be in the resetcontrols because it is secific to this process. 
            testBu.Text = ""
            testSid.Text = ""
            testWo.Text = ""
            testSaFolder.Text = ""
            testFolder.Text = ""
            testComb.Text = ""
            testID.SelectedIndex = -1
            testIteration.Text = ""
            testNextIteration.Text = ""

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
            btnNextIteration.Enabled = Not btnNextIteration.Enabled
            testIterationResults.Enabled = Not testIterationResults.Enabled
            testPrevResults.Enabled = Not testPrevResults.Enabled
            testPublishedResults.Enabled = Not testPublishedResults.Enabled
            testConduct.Enabled = Not testConduct.Enabled
            testCompareAll.Enabled = Not testCompareAll.Enabled
            testStructureOnly.Enabled = Not testStructureOnly.Enabled
            testJason.Enabled = Not testJason.Enabled
            mainLogViewer.Enabled = Not mainLogViewer.Enabled

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
            XtraTabControl1.Enabled = Not XtraTabControl1.Enabled
            testBugFile.Enabled = Not testBugFile.Enabled
            testSync.Enabled = Not testSync.Enabled
            If New FileInfo(Me.dirUse & "\Test ID " & Me.testCase & "\Checked Out.txt").Exists Then testPush.Enabled = Not testPush.Enabled
            testPull.Enabled = Not testPull.Enabled

            rtfactivityLog.Visible = Not rtfactivityLog.Visible
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
                    sw.WriteLine(dt & " | " & Environment.UserName & " | " & msg)
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
                LogActivity("DEBUG | Purged Folder: " & newSync.CopyOptions.Destination)
                LogActivity("DEBUG | Forced Sync local drive to R: Drive")
            Else
                LogActivity("DEBUG | Purged Folder: " & newSync.CopyOptions.Destination)
                InitializeLocaltoCentralSync(newSync, rFolder & "\Test ID " & Me.testCase, lFolder & "\Test ID " & Me.testCase, Purge)
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
        Public Function ImportInputs(ByVal FileType As String) As Boolean
            Dim SAFiles As New DataTable
            SAFiles = CSVtoDatatable(New FileInfo(Me.RefFolder & "\File List.csv"))
            'If SAFiles.Columns.Count > 2 Then
            '    CreateTemplateFiles(frmMain.testIteration.Text)
            'End If

            Dim myXL As Tuple(Of Excel.Application, Boolean) = GetXlApp()
            'Item 1 = Excel application
            'Item 2 = Boolean (If true that means excel was previously open

            For Each dr As DataRow In SAFiles.Rows()
                Dim importingFrom As New FileInfo(dr.Item("FilePath").ToString)
                If importingFrom.Extension.ToLower = ".xlsm" Then
                    Dim importingTo As New FileInfo(dr.Item(FileType).ToString)
                    Dim macroname As String = "Import_Previous_Version"
                    Dim params As Tuple(Of String, String, Boolean) = New Tuple(Of String, String, Boolean)(importingFrom.FullName.ToString, importingFrom.TemplateVersion, True)

                    If importingTo.Name.ToLower.Contains("pile") Then
                        macroname = "Button173_Click"
                    ElseIf importingTo.Name.ToLower.Contains("drilled pier") Then
                        If FileType = "MaestroPath" Then
                            macroname += "_Performer"
                        End If
                    End If

                    If Import_Previous_Version(myXL.Item1, importingTo, macroname, params, True) Then
                        LogActivity("INFO | Import Inputs Completed for: " & importingTo.FullName)
                        If FileType = "MaestroPath" Then
                            Try
                                Dim manFile As New FileInfo(Me.ManFolder & "\" & importingTo.Name)
                                manFile.Delete()
                                importingTo.CopyTo(manFile.FullName)
                            Catch ex As Exception
                            End Try
                        End If
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
                                                Optional ByVal xlVisibility As Boolean = False
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
                        xlapp.Run("Import_Previous_Version." & macroName, params.Item1, params.Item2, params.Item3)
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
            testIteration.Text = nextIteration
            testNextIteration.Text = nextIteration + 1
            testFolder.Text = "R:\Development\SAPI Testing\Unit Testing\Test ID " & testCase

            If GetReferenceFileCount() = 0 Then
                FirstTimeWarning(isFirstTime, nextIteration)
            Else
                '''Create the directories
                '''Increase the iteration (Should be at 0 if this is the first time)
                '''get all required files for testing
                Directory.CreateDirectory(Me.itFolder)
                Directory.CreateDirectory(Me.MaeFolder)
                Directory.CreateDirectory(Me.ManFolder)
                'CreateTemplateFiles(isFirstTime, False)
            End If

        End Sub

        'Folder option for this one since it could send ERIs to multiple locations. 
        Public Sub CreateManualERI(ByRef refFiles As DataTable, ByVal folder As String, Optional ByVal isFirstTime As Boolean = False, Optional ByVal archive As Boolean = True)
            If archive And Not isFirstTime Then DoArchiving(folder)

            For Each dr As DataRow In refFiles.Rows
                Dim file As New FileInfo(dr.Item("FilePath").ToString)
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
                Dim file As New FileInfo(dr.Item("FilePath").ToString)
                If file.Extension.ToLower = ".xlsm" Then
                    With WhichFile(file)
                        If .Item1 Is Nothing Or .Item2 Is Nothing Or .Item3 Is Nothing Then
                            TemplateNotFoundWarning(file)
                        Else
                            'Copy published versions of the tools into the manual folder 
                            Dim pubPath As String = GetNewFileName(Me.PubFolder, fileName:= .Item3)
                            IO.File.WriteAllBytes(pubPath, .Item1)
                            dr.Item("PublishedPath") = pubPath
                            LogActivity("DEBUG | Production version created: " & pubPath)
                        End If
                    End With
                End If
            Next
        End Sub

        'No folder option for these since they are always going to go to the current folder path
        Public Sub CreateSAPITemplates(ByRef refFiles As DataTable, Optional ByVal isFirstTime As Boolean = False, Optional ByVal archive As Boolean = True)
            If archive And Not isFirstTime Then
                DoArchiving(Me.MaeFolder)
                DoArchiving(Me.ManFolder)
            End If

            For Each dr As DataRow In refFiles.Rows
                Dim file As New FileInfo(dr.Item("FilePath").ToString)
                If file.Extension.ToLower = ".xlsm" Then
                    With WhichFile(file)
                        If .Item1 Is Nothing Or .Item2 Is Nothing Or .Item3 Is Nothing Then
                            TemplateNotFoundWarning(file)
                        Else
                            'Templates are saved as Bytes() and need to be converted appropriately. 
                            Dim maePath As String = GetNewFileName(Me.MaeFolder, fileName:= .Item3)
                            IO.File.WriteAllBytes(maePath, .Item2)
                            dr.Item("MaestroPath") = maePath
                            LogActivity("DEBUG | SAPI version created: " & maePath)

                            'File will be copied to the manual folder once the files are populated with data via 
                            'Manual files will be replaces when user imports data into the maestro files.
                            'Alternative will be to load maestro and manual files manually.
                            '''Import Inputs
                            '''Structure import
                            Dim manPath As String = GetNewFileName(Me.ManFolder, fileName:= .Item3)
                            IO.File.WriteAllBytes(manPath, .Item2)
                            dr.Item("ManualPath") = manPath
                            LogActivity("DEBUG | SAPI version created: " & manPath)
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
            testIteration.Text = nextIteration - 1
            testNextIteration.Text = nextIteration
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
                ElseIf myTemplate.Item3.ToLower.Contains("CCIplate") Then
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
                                        val = dr.Item("Rating*").ToString.Replace("%", "") * 100
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
        Public Function CompareResults() As Tuple(Of Tuple(Of Boolean, DataTable), Tuple(Of Boolean, DataTable), Tuple(Of Boolean, DataTable), DataSet)
            Dim manToMae As Tuple(Of Boolean, DataTable) 'Item1
            Dim curToMan As Tuple(Of Boolean, DataTable) 'Item2
            Dim curToMae As Tuple(Of Boolean, DataTable) 'Item2
            Dim resDs As New DataSet 'Item4

            Dim dir As String = IIf(CType(chkWorkLocal.Checked, Boolean) = True, lFolder, rFolder)
            Dim testid As Integer = CType(Me.testID.Text, Integer)
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
            LogActivity("DEBUG | Results output for reference SA files created: " & folder & "\Summarized Results.csv")
        End Sub

#End Region

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
                        If dr.Item("FilePath").ToString = file.FullName Then
                            isFound = True
                            Exit For
                        End If
                    Next

                    If Not isFound Then
                        With WhichFile(file)
                            If file.Extension.ToLower = ".eri" Or .Item1 IsNot Nothing Or .Item2 IsNot Nothing Or .Item3 IsNot Nothing Then
                                LogActivity("DEBUG | File added manually: " & file.Name)
                                If SAFiles.Columns.Count < 4 Then
                                    SAFiles.Rows.Add(file.FullName, file.TemplateVersion, "")
                                Else
                                    SAFiles.Rows.Add(file.FullName, file.TemplateVersion, "", "", "", "")
                                End If
                            End If
                        End With

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
        Public Sub CreateStructure()
            Dim myFiles As String()
            Dim myFilesLst As New List(Of String)

            'Loop through all files in the maestro folder for the current test case and iteration
            For Each info As FileInfo In New DirectoryInfo(Me.MaeFolder).GetFiles
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
                    info.Delete()
                    LogActivity("DEBUG | File Deleted: " & info.FullName)
                End If
            Next

            'Convert the list of valid file names to an array for creating anew structure
            myFiles = myFilesLst.ToArray
            strcLocal = New EDSStructure(testBu.Text, testSid.Text, testWo.Text, Me.MaeFolder, Me.MaeFolder, myFiles, EDSnewId, EDSdbActive)
        End Sub

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
                    filePath = newFolder & "\" & file.Name.Split(".")(0) & "(" & counter.ToString() & ")" & file.Name.Split(".")(1)
                Else
                    filePath = newFolder & "\" & fileName.Split(".")(0) & "(" & counter.ToString() & ")" & fileName.Split(".")(1)
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
    End Class

    Public Module GeneralHelpers 'salute

        Public isopening As Boolean = True

        'serialize any object to a json
        '''Object being passed in
        '''location to save the file path
        Public Function ObjectToJson(Of T)(ByVal obj As Object, ByVal jsonPath As String) As Boolean
            Dim objJson As String

            Try
                objJson = ToJsonString(Of T)(CType(obj, T))
                Using sw As New StreamWriter(jsonPath)
                    sw.Write(objJson)
                    sw.Close()
                End Using
                Return True
            Catch ex As Exception
                objJson = Nothing
                Return False
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
        'Archive files in a directory to another directory.
        <Extension()>
        Public Sub ArchiveFiles(ByVal dirTo As DirectoryInfo, ByVal dirFrom As String)
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
                                IIf(Double.IsNaN(dtVal), "N/A", dtVal),
                                IIf(Double.IsNaN(comparerVal), "N/A", comparerVal),
                                IIf(Double.IsNaN(delta), "N/A", delta),
                                IIf(Double.IsNaN(perDelta), "N/A", perDelta),
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
        Public Sub AppendLog(ByVal strc As EDSStructure, ByVal pathToAppend As String)
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
                                sw.WriteLine(dt & " " & vars(0) & splt(2) & " " & separator & " " & Environment.UserName & " " & separator & "INFO" & " " & separator & vars(1))
                            Else
                                sw.WriteLine(dt & " " & vars(0) & splt(2) & " " & separator & " " & Environment.UserName & " " & separator & vars(1) & separator & vars(2))
                            End If
                        End If
                    End While
                    sr.Close()
                End Using
            End Using
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
        Public Function FromJsonString(Of T)(ByVal jsonString As String) As T
            Using aMemoryStream As MemoryStream = New MemoryStream(Encoding.UTF8.GetBytes(jsonString))
                Dim ser = New DataContractJsonSerializer(GetType(T))
                Return CType(ser.ReadObject(aMemoryStream), T)
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
                serializer.WriteObject(aMemoryStream, valueObject)
                Return Encoding.[Default].GetString(aMemoryStream.ToArray())
            End Using
        End Function
    End Module
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