Imports System.ComponentModel
Imports System.Text
Imports CCI_Engineering_Templates
Imports System.Data.SqlClient
Imports System.Security.Principal
Imports System.IO
Imports Oracle.ManagedDataAccess.Client

Partial Public Class frmMain
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
#End Region

#Region "Other Required Declarations"
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

    Public isopening As Boolean

    Private Function token(s As String) As String
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

        If Environment.UserName.ToLower = "imiller" Or Environment.UserName.ToLower = "stanley" Then
            txtFndBU.Text = My.Settings.myBU
            txtFndStrc.Text = My.Settings.myStrID
            txtFndWO.Text = My.Settings.myWO
            txtDirectory.Text = My.Settings.myWorkArea
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
    End Sub
#End Region


#Region "Structure"
    Public strcLocal As EDSStructure
    Public strcEDS As EDSStructure

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
            strcLocal = New EDSStructure(txtFndBU.Text, txtFndStrc.Text, txtFndWO.Text, workingDirectory, workingDirectory, xlFd.FileNames, EDSnewId, EDSdbActive, True)

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

#End Region


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

#Region "Textbox Changes"

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

    Private Sub btnConduct_Click(sender As Object, e As EventArgs) Handles btnConduct.Click

        strcLocal.Conduct(True)

    End Sub

    Private Sub testSaFolder_EditValueChanged(sender As Object, e As EventArgs) Handles testSaFolder.EditValueChanged
        Try
            SimpleExplorer2.SetCurrentDirectory(testSaFolder.Text)
        Catch ex As Exception
            SimpleExplorer2.SetCurrentDirectory(testSaFolder.Text.Replace(" - SA", ""))
        End Try

    End Sub

    Private Sub testFolder_EditValueChanged(sender As Object, e As EventArgs) Handles testFolder.EditValueChanged
        SimpleExplorer1.SetCurrentDirectory(testFolder.Text)
    End Sub

    Public unitTestCases As New List(Of TestCase)

    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        If TabControl1.SelectedTab.Name = TabPage1.Name Then
            If unitTestCases.Count > 0 Then Exit Sub

            Using csvReader As New Microsoft.VisualBasic.FileIO.TextFieldParser("R:\Development\SAPI Testing\Unit Test Cases.csv")
                csvReader.TextFieldType = FileIO.FieldType.Delimited
                csvReader.SetDelimiters(",")
                Dim csvValue As String()

                While Not csvReader.EndOfData
                    csvValue = csvReader.ReadFields
                    unitTestCases.Add(New TestCase(csvValue))
                End While
            End Using
        End If
    End Sub

    Private Sub testID_SelectedIndexChanged(sender As Object, e As EventArgs) Handles testID.SelectedIndexChanged
        Dim id As Integer = testID.Text - 1
        Dim testCase As Integer = testID.Text

        If Not Directory.Exists("R:\Development\SAPI Testing\Unit Testing\Test ID " & testCase) Then
            Directory.CreateDirectory("R:\Development\SAPI Testing\Unit Testing\Test ID " & testCase)
            Directory.CreateDirectory("R:\Development\SAPI Testing\Unit Testing\Test ID " & testCase & "\Manual (Current)")
            Directory.CreateDirectory("R:\Development\SAPI Testing\Unit Testing\Test ID " & testCase & "\Reference SA Files")
            Directory.CreateDirectory("R:\Development\SAPI Testing\Unit Testing\Test ID " & testCase & "\Manual ERI")
            CreateIteration(1, True)
        End If

        testBu.Text = unitTestCases(id).BU
        testSid.Text = unitTestCases(id).SID
        testWo.Text = unitTestCases(id).WO
        testSaFolder.Text = unitTestCases(id).SAWorkArea
        testFolder.Text = "R:\Development\SAPI Testing\Unit Testing\Test ID " & testCase
        testComb.Text = unitTestCases(id).COMB

        Dim itCount As Integer = 0
        For Each subDir In New DirectoryInfo("R:\Development\SAPI Testing\Unit Testing\Test ID " & testCase).GetDirectories
            If subDir.Name.Contains("Iteration ") Then itCount += 1
        Next

        testIteration.Text = itCount
        testNextIteration.Text = itCount + 1

        btnNextIteration.Enabled = True
    End Sub

    Private Sub btnNextIteration_Click(sender As Object, e As EventArgs) Handles btnNextIteration.Click
        CreateIteration(testNextIteration.Text)
    End Sub


    Public Sub CreateIteration(ByVal Iteration As Integer, ByVal Optional isFirstTime As Boolean = False)
        Dim answer = vbYes
        Dim testCase As Integer = testID.Text
        Dim MaeFolder As String = "R:\Development\SAPI Testing\Unit Testing\Test ID " & testCase & "\Iteration " & Iteration & "\Maestro"
        Dim ManFolder As String = "R:\Development\SAPI Testing\Unit Testing\Test ID " & testCase & "\Iteration " & Iteration & "\Manual (SAPI)"
        Dim PubFolder As String = "R:\Development\SAPI Testing\Unit Testing\Test ID " & testCase & "\Manual (Current)"
        Dim RefFolder As String = "R:\Development\SAPI Testing\Unit Testing\Test ID " & testCase & "\Reference SA Files"
        Dim EriFolder As String = "R:\Development\SAPI Testing\Unit Testing\Test ID " & testCase & "\Manual ERI"

        Dim fileCount As Integer = Directory.GetFiles(RefFolder).Count
        Dim publishedFileCount As Integer = Directory.GetFiles(PubFolder).Count
        Dim eriFileCount As Integer = Directory.GetFiles(EriFolder).Count

        If fileCount = 0 Then
            If Not isFirstTime Then MsgBox("Files do not exist in the 'Reference SA Files' folder yet. Please copy reference files to continue.", vbCritical, "No Reference Files")
            testIteration.Text = Iteration - 1
            testNextIteration.Text = Iteration
        Else
            Directory.CreateDirectory("R:\Development\SAPI Testing\Unit Testing\Test ID " & testCase & "\Iteration " & Iteration)
            Directory.CreateDirectory(MaeFolder)
            Directory.CreateDirectory(ManFolder)

            testIteration.Text = Iteration
            testNextIteration.Text = Iteration + 1
            testFolder.Text = "R:\Development\SAPI Testing\Unit Testing\Test ID " & testCase

            For Each file In New DirectoryInfo(RefFolder).GetFiles
                If file.Extension.Contains("eri") Then
                    file.CopyTo(MaeFolder)
                    If eriFileCount = 0 Then file.CopyTo(EriFolder)
                Else
                    Dim myTemplate As Tuple(Of FileInfo, FileInfo) = WhichFile(file)

                    If myTemplate.Item1 Is Nothing Or myTemplate.Item2 Is Nothing Then
                        MsgBox("Could not determine template file type for file: " & vbCrLf & file.Name & vbCrLf & vbCrLf & "Please copy template manually.", vbCritical, "Template Not Found")
                    Else
                        With myTemplate
                            If publishedFileCount = 0 Then
                                'Copy published versions of the tools into the manual folder 
                                .Item1.CopyTo(GetNewFileName(PubFolder, .Item1))
                            End If

                            .Item2.CopyTo(GetNewFileName(MaeFolder, .Item2))
                            .Item2.CopyTo(GetNewFileName(ManFolder, .Item2))
                        End With
                    End If
                End If
            Next


        End If
    End Sub

    Public Function GetNewFileName(ByVal newFolder As String, ByVal file As FileInfo) As String
        Dim counter As Integer = 0
        Dim filePath As String = String.Format("{0}({1}", newFolder & "\" & file.Name, counter.ToString())
        While IO.File.Exists(filePath)
            counter += 1
            filePath = String.Format("{0}({1}", newFolder & "\" & file.Name, counter.ToString())
        End While

        Return filePath
    End Function

    Public Function WhichFile(ByVal file As FileInfo) As Tuple(Of FileInfo, FileInfo)
        Dim returner As Tuple(Of FileInfo, FileInfo)

        If file.Name.ToLower.Contains("cciplate") Then
            returner = New Tuple(Of FileInfo, FileInfo)(New FileInfo(""), New FileInfo(""))
        ElseIf file.Name.ToLower.Contains("ccipole") Then
            returner = New Tuple(Of FileInfo, FileInfo)(New FileInfo(""), New FileInfo(""))
        ElseIf file.Name.ToLower.Contains("cciseismic") Then
            returner = New Tuple(Of FileInfo, FileInfo)(New FileInfo(""), New FileInfo(""))
        ElseIf file.Name.ToLower.Contains("drilled pier") Then
            returner = New Tuple(Of FileInfo, FileInfo)(New FileInfo(""), New FileInfo(""))
        ElseIf file.Name.ToLower.Contains("guyed anchor") Then
            returner = New Tuple(Of FileInfo, FileInfo)(New FileInfo(""), New FileInfo(""))
        ElseIf file.Name.ToLower.Contains("leg reinforcement") Then
            returner = New Tuple(Of FileInfo, FileInfo)(New FileInfo(""), New FileInfo(""))
        ElseIf file.Name.ToLower.Contains("pier and pad") Then
            returner = New Tuple(Of FileInfo, FileInfo)(New FileInfo(""), New FileInfo(""))
        ElseIf file.Name.ToLower.Contains("pile") Then
            returner = New Tuple(Of FileInfo, FileInfo)(New FileInfo(""), New FileInfo(""))
        ElseIf file.Name.ToLower.Contains("unit base") Then
            returner = New Tuple(Of FileInfo, FileInfo)(New FileInfo(""), New FileInfo(""))
        Else
            returner = New Tuple(Of FileInfo, FileInfo)(Nothing, Nothing)
        End If

        Return returner
    End Function

#End Region

End Class


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

