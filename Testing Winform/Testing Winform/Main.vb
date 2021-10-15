Imports System.ComponentModel
Imports System.Text
Imports CCI_Engineering_Templates
Imports System.Data.SqlClient
Imports System.Security.Principal
Imports System.IO

Partial Public Class frmMain
#Region "Object Declarations"
    'Public myUnitBases As New DataTransfererUnitBase
    Public myPierandPads As New DataTransfererPierandPad
    Public myDrilledPiers As New DataTransfererDrilledPier
    Public myGuyedAnchorBlocks As New DataTransfererGuyedAnchorBlock
    Public myPiles As New DataTransfererPile

    Public BUNumber As String = ""
    Public StrcID As String = ""

    'Import to Excel
    Public ListOfFilesCopied As New List(Of String) From {"C:\Users\" & Environment.UserName & "\Desktop\Save to Excel\EDS - 870800 - Drilled Pier Foundation (5.1.0) - 10-14-21.xlsm"}
    'Public ListOfFilesCopied As New List(Of String) From {"C:\Users\" & Environment.UserName & "\Desktop\Save to Excel\EDS - 871191 - Guyed Anchor Block Foundation (4.1.0) - 10-14-21.xlsm"}
    'Import to EDS
    Public ListOfExcelFiles As New List(Of String) From {"C:\Users\" & Environment.UserName & "\Desktop\Save to EDS\870800 - Drilled Pier Foundation (5.1.0) - 10-14-21.xlsm"}
    'Public ListOfExcelFiles As New List(Of String) From {"C:\Users\" & Environment.UserName & "\Desktop\Save to EDS\871191 - Guyed Anchor Block Foundation (4.1.0) - 10-13-2021.xlsm"}
#End Region

#Region "Other Required Declarations"
    Public EDSdbDevelopment As String = "Server=DEVCCICSQL2.US.CROWNCASTLE.COM,60113;Database=EDSDev;Integrated Security=SSPI"
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
        If My.Settings.serverActive = "dbDevelopment" Then
            EDSdbActive = EDSdbDevelopment
            EDSuserActive = EDSuserDevelopment
            EDSuserPwActive = EDSuserPwDevelopment
        Else
            EDSdbActive = EDSdbProduction
            EDSuserActive = EDSuserProduction
            EDSuserPwActive = EDSuserPwProduction
        End If

        LogonUser(token(EDSuserActive), "CCIC", token(EDSuserPwActive), 2, 0, EDStokenHandle)
        EDSnewId = New WindowsIdentity(EDStokenHandle)
    End Sub

    Private Sub frmMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        CloseHandle(EDStokenHandle)
    End Sub
#End Region

    Public Sub CreateExcelTemplates() Handles sqltoexcel.Click
        ClearAllTools()

        For Each item As String In ListOfFilesCopied
            If item.Contains("SST Unit Base Foundation") Then
                'myUnitBases = New DataTransfererUnitBase(ds, EDSnewId, EDSdbActive, BUNumber, StrcID)
                'myUnitBases.ExcelFilePath = item
                'If myUnitBases.LoadFromEDS() Then myUnitBases.SaveToExcel()
            ElseIf item.Contains("Pier and Pad Foundation") Then
                myPierandPads = New DataTransfererPierandPad(ds, EDSnewId, EDSdbActive, BUNumber, StrcID)
                myPierandPads.ExcelFilePath = item
                If myPierandPads.LoadFromEDS() Then myPierandPads.SaveToExcel()
            ElseIf item.Contains("Drilled Pier Foundation") Then
                myDrilledPiers = New DataTransfererDrilledPier(ds, EDSnewId, EDSdbActive, BUNumber, StrcID)
                myDrilledPiers.ExcelFilePath = item
                If myDrilledPiers.LoadFromEDS() Then myDrilledPiers.SaveToExcel()
            ElseIf item.Contains("Guyed Anchor Block Foundation") Then
                myGuyedAnchorBlocks = New DataTransfererGuyedAnchorBlock(ds, EDSnewId, EDSdbActive, BUNumber, StrcID)
                myGuyedAnchorBlocks.ExcelFilePath = item
                If myGuyedAnchorBlocks.LoadFromEDS() Then myGuyedAnchorBlocks.SaveToExcel()
            ElseIf item.Contains("Pile Foundation") Then
                myPiles = New DataTransfererPile(ds, EDSnewId, EDSdbActive, BUNumber, StrcID)
                myPiles.ExcelFilePath = item
                If myPiles.LoadFromEDS() Then myPiles.SaveToExcel()
            End If
        Next

    End Sub

    Public Sub UploadExcelFilesToEDS() Handles exceltosql.Click
        ClearAllTools()

        For Each item As String In ListOfExcelFiles
            If item.Contains("SST Unit Base Foundation") Then
                'myUnitBases = New DataTransfererUnitBase(ds, EDSnewId, EDSdbActive, BUNumber, StrcID)
                'myUnitBases.ExcelFilePath = item
                'myUnitBases.LoadFromExcel()
                'myUnitBases.SaveToEDS()
            ElseIf item.Contains("Pier and Pad Foundation") Then
                myPierandPads = New DataTransfererPierandPad(ds, EDSnewId, EDSdbActive, BUNumber, StrcID)
                myPierandPads.ExcelFilePath = item
                myPierandPads.LoadFromExcel()
                myPierandPads.SaveToEDS()
            ElseIf item.Contains("Drilled Pier Foundation") Then
                myDrilledPiers = New DataTransfererDrilledPier(ds, EDSnewId, EDSdbActive, BUNumber, StrcID)
                myDrilledPiers.ExcelFilePath = item
                myDrilledPiers.LoadFromExcel()
                myDrilledPiers.SaveToEDS()
            ElseIf item.Contains("Guyed Anchor Block Foundation") Then
                myGuyedAnchorBlocks = New DataTransfererGuyedAnchorBlock(ds, EDSnewId, EDSdbActive, BUNumber, StrcID)
                myGuyedAnchorBlocks.ExcelFilePath = item
                myGuyedAnchorBlocks.LoadFromExcel()
                myGuyedAnchorBlocks.SaveToEDS()
            ElseIf item.Contains("Pile Foundation") Then
                myPiles = New DataTransfererPile(ds, EDSnewId, EDSdbActive, BUNumber, StrcID)
                myPiles.ExcelFilePath = item
                myPiles.LoadFromExcel()
                myPiles.SaveToEDS()
            End If
        Next

    End Sub

    Sub ClearAllTools()
        'myUnitBases.Clear()
        myPierandPads.Clear()
        myDrilledPiers.Clear()
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        MsgBox("Stop touching me")
    End Sub

    Private Sub CreateExcelTemplates(sender As Object, e As EventArgs) Handles sqltoexcel.Click

    End Sub

    Private Sub UploadExcelFilesToEDS(sender As Object, e As EventArgs) Handles exceltosql.Click

    End Sub

    Public tnxObject As tnxModel
    Private Sub btnImportTNX_Click(sender As Object, e As EventArgs) Handles btnImportTNX.Click
        Dim eriFd As New OpenFileDialog
        eriFd.Multiselect = False
        eriFd.Filter = "TNX File|*.eri"

        If eriFd.ShowDialog = DialogResult.OK Then
            tnxObject = New tnxModel(eriFd.FileName)
            propgridTNXObject.SelectedObject = tnxObject
        End If
    End Sub

    Private Sub btnExportTNX_Click(sender As Object, e As EventArgs) Handles btnExportTNX.Click
        If tnxObject Is Nothing Then
            MessageBox.Show("Import a file first.")
            Exit Sub
        End If

        Dim eriFd As New SaveFileDialog
        eriFd.Filter = "TNX File|*.eri"

        If eriFd.ShowDialog = DialogResult.OK Then
            tnxObject.GenerateERI(eriFd.FileName)
        End If
    End Sub
End Class
