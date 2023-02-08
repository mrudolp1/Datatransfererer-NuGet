'Option Strict Off

'Imports DevExpress.Spreadsheet
'Imports System.Security.Principal
''Imports Microsoft.Office.Interop.Excel
''Imports Microsoft.Office.Interop

'Public Class DataTransfererLegReinforcement

'#Region "Define"
'    Private NewLegReinforcementWb As New DevExpress.Spreadsheet.Workbook 'Excel.Workbook
'    Private prop_ExcelFilePath As String

'    Public Property LegReinforcements As New List(Of LegReinforcement)
'    Private Property LegReinforcementTemplatePath As String = "C:\Users\" & Environment.UserName & "\Desktop\Leg Reinforcement Tool (10.0.4) - TEMPLATE - 9-14-2021.xlsm"
'    Private Property LegReinforcementFileType As DocumentFormat = DocumentFormat.Xlsm

'    Public Property lrDB As String
'    Public Property lrID As WindowsIdentity

'    Public Property ExcelFilePath() As String
'        Get
'            Return Me.prop_ExcelFilePath
'        End Get
'        Set
'            Me.prop_ExcelFilePath = Value
'        End Set
'    End Property

'    Public Property xlApp As Object
'#End Region

'#Region "Constructors"
'    Public Sub New()
'        'Leave method empty
'    End Sub

'    Public Sub New(ByVal MyDataSet As DataSet, ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String, ByVal BU As String, ByVal Strucutre_ID As String)
'        ds = MyDataSet
'        lrID = LogOnUser
'        lrDB = ActiveDatabase
'        'BUNumber = BU 'Need to turn back on when connecting to dashboard. Turned off for testing.
'        'STR_ID = Strucutre_ID 'Need to turn back on when connecting to dashboard. Turned off for testing.
'    End Sub
'#End Region

'#Region "Load Data"
'    Public Function LoadFromEDS() As Boolean
'        Dim refid As Integer

'        Dim LegReinforcementLoader As String

'        'Load data to get pier and pad details data for the existing structure model
'        For Each item As SQLParameter In LegReinforcementSQLDataTables()
'            LegReinforcementLoader = QueryBuilderFromFile(queryPath & "Leg Reinforcement\" & item.sqlQuery).Replace("[EXISTING MODEL]", GetExistingModelQuery())
'            DoDaSQL.sqlLoader(LegReinforcementLoader, item.sqlDatatable, ds, lrDB, lrID, "0")
'            'If ds.Tables(item.sqlDatatable).Rows.Count = 0 Then Return False 'This may need adjusted since some tables can be empty
'        Next

'        'Custom Section to transfer data for the tool. Needs to be adjusted for each tool.
'        For Each LegReinforcementDataRow As DataRow In ds.Tables("Leg Reinforcement General Details SQL").Rows
'            refid = CType(LegReinforcementDataRow.Item("leg_rein_id"), Integer)

'            LegReinforcements.Add(New LegReinforcement(LegReinforcementDataRow, refid))
'        Next

'        Return True
'    End Function 'Create Leg Reinforcement objects based on what is saved in EDS

'    Public Sub LoadFromExcel()
'        Dim refID As Integer
'        Dim refCol As String

'        For Each item As EXCELDTParameter In LegReinforcementExcelDTParameters()
'            'Get tables from excel file 
'            ds.Tables.Add(ExcelDatasourceToDataTable(GetExcelDataSource(ExcelFilePath, item.xlsSheet, item.xlsRange), item.xlsDatatable))
'        Next

'        'Custom Section to transfer data for the tool. Needs to be adjusted for each tool.
'        For Each LegReinforcementDataRow As DataRow In ds.Tables("Leg Reinforcement General Details EXCEL").Rows

'            refCol = "local_tool_id"
'            refID = CType(LegReinforcementDataRow.Item(refCol), Integer)

'            LegReinforcements.Add(New LegReinforcement(LegReinforcementDataRow, refID, refCol))
'        Next
'    End Sub 'Create Leg Reinforcement objects based on what is coming from the excel file
'#End Region

'#Region "Save Data"
'    Public Sub SaveToEDS()

'        For Each lr As LegReinforcement In LegReinforcements

'            Dim LegReinforcementSaver As String = QueryBuilderFromFile(queryPath & "Leg Reinforcement\Leg Reinforcement (IN_UP).sql")

'            'COMPLETE ONCE SQL CODE IS UPDATED (CHRIS/IAN)

'            LegReinforcementSaver = LegReinforcementSaver.Replace("[INSERT ALL LEG REINFORCEMENT DETAILS]", InsertLegReinforcementDetail(lr))

'            sqlSender(LegReinforcementSaver, lrDB, lrID, "0")
'        Next
'    End Sub

'    Public Sub SaveToExcel()

'        Dim lrRow As Integer = 3

'        LoadNewLegReinforcement()

'        With NewLegReinforcementWb

'            Dim colCounter As Integer
'            Dim myCol As String
'            Dim rowStart As Integer = 4

'            For Each lr As LegReinforcement In LegReinforcements

'                'Leg Reinforcement ID
'                colCounter = 22
'                myCol = GetExcelColumnName(colCounter)
'                If Not IsNothing(lr.leg_rein_id) Then 'FIX NEEDED. RESULTS IN 0. May need to specify "lr.ID"
'                    .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(lr.leg_rein_id, Integer)
'                Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                End If

'                'Leg Reinforcement Data Stored
'                colCounter = colCounter + 3
'                myCol = GetExcelColumnName(colCounter)
'                If Not IsNothing(lr.data_stored) Then
'                    .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(lr.data_stored, Boolean)
'                Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = False
'                End If

'                'Leg Reinforcement Type
'                colCounter = colCounter + 3
'                myCol = GetExcelColumnName(colCounter)
'                If Not IsNothing(lr.rein_type) Then
'                    .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(lr.rein_type, String)
'                Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                End If

'                'Leg Load at Time of Modification
'                colCounter = colCounter + 3
'                myCol = GetExcelColumnName(colCounter)
'                If Not IsNothing(lr.leg_load_time_of_mod) Then
'                    .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(lr.leg_load_time_of_mod, Boolean)
'                Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                End If

'                'Leg Reinforcement End Connection Type
'                colCounter = colCounter + 1
'                myCol = GetExcelColumnName(colCounter)
'                If Not IsNothing(lr.end_connections) Then
'                    .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(lr.end_connections, String)
'                Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                End If

'                'Leg Crushing Applied
'                colCounter = colCounter + 1
'                myCol = GetExcelColumnName(colCounter)
'                If Not IsNothing(lr.leg_crushing) Then
'                    If CType(lr.leg_crushing, Boolean) = False Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = "No"
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = "Yes"
'                    End If
'                End If

'                'Applied Load Methodology
'                colCounter = colCounter + 1
'                myCol = GetExcelColumnName(colCounter)
'                If Not IsNothing(lr.applied_load_methodology) Then
'                    .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(lr.applied_load_methodology, String)
'                Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                End If

'                'Slenderness Ratio Type
'                colCounter = colCounter + 1
'                myCol = GetExcelColumnName(colCounter)
'                If Not IsNothing(lr.slenderness_ratio_type) Then
'                    .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(lr.slenderness_ratio_type, String)
'                Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                End If

'                'Intermediate Connection Type
'                colCounter = colCounter + 1
'                myCol = GetExcelColumnName(colCounter)
'                If Not IsNothing(lr.intermediate_conn_type) Then
'                    .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(lr.intermediate_conn_type, String)
'                Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                End If

'                'Intermediate Connection Spacing
'                colCounter = colCounter + 1
'                myCol = GetExcelColumnName(colCounter)
'                If Not IsNothing(lr.intermediate_conn_spacing) Then
'                    .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(lr.intermediate_conn_spacing, Double)
'                Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                End If

'                'Ki Override
'                colCounter = colCounter + 1
'                myCol = GetExcelColumnName(colCounter)
'                If Not IsNothing(lr.ki_override) Then
'                    .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(lr.ki_override, Double)
'                Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                End If

'                'Leg Diameter
'                colCounter = colCounter + 1
'                myCol = GetExcelColumnName(colCounter)
'                If Not IsNothing(lr.leg_dia) Then
'                    .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(lr.leg_dia, Double)
'                Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                End If

'                'Leg Thickness
'                colCounter = colCounter + 1
'                myCol = GetExcelColumnName(colCounter)
'                If Not IsNothing(lr.leg_thickness) Then
'                    .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(lr.leg_thickness, Double)
'                Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                End If

'                'Leg Yield Strength
'                colCounter = colCounter + 1
'                myCol = GetExcelColumnName(colCounter)
'                If Not IsNothing(lr.leg_yield_strength) Then
'                    .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(lr.leg_yield_strength, Double)
'                Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                End If

'                'Leg Unbraced Length
'                colCounter = colCounter + 1
'                myCol = GetExcelColumnName(colCounter)
'                If Not IsNothing(lr.leg_unbraced_length) Then
'                    .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(lr.leg_unbraced_length, Double)
'                Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                End If

'                'Leg Reinforcement Diameter
'                colCounter = colCounter + 1
'                myCol = GetExcelColumnName(colCounter)
'                If Not IsNothing(lr.rein_dia) Then
'                    .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(lr.rein_dia, Double)
'                Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                End If

'                'Leg Reinforcement Thickness
'                colCounter = colCounter + 1
'                myCol = GetExcelColumnName(colCounter)
'                If Not IsNothing(lr.rein_thickness) Then
'                    .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(lr.rein_thickness, Double)
'                Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                End If

'                'Leg Reinforcement Yield Strength
'                colCounter = colCounter + 1
'                myCol = GetExcelColumnName(colCounter)
'                If Not IsNothing(lr.rein_yield_strength) Then
'                    .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(lr.rein_yield_strength, Double)
'                Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                End If

'                'Print Bolt-On Connection Info
'                colCounter = colCounter + 3
'                myCol = GetExcelColumnName(colCounter)
'                If Not IsNothing(lr.print_bolton_conn_info) Then
'                    .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(lr.print_bolton_conn_info, Boolean)
'                Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = False
'                End If


'                For Each boc As BoltOnConnection In lr.BoltOnConnections

'                    'Bolt-On Connections ID
'                    colCounter = 49
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(boc.bolton_id) Then 'FIX NEEDED. RESULTS IN 0. May need to specify "lr.ID"
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(boc.bolton_id, Integer)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Leg Length of Tower Section
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(boc.leg_length_of_tower_section) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(boc.leg_length_of_tower_section, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Physical Split Pipe Length
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(boc.split_pipe_length) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(boc.split_pipe_length, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Set Top Design Info Equal to Bottom Design Info
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(boc.set_top_to_bottom) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(boc.set_top_to_bottom, Boolean)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Quantity of Bottom Flange Bolts
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(boc.qty_flange_bolt_bot) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(boc.qty_flange_bolt_bot, Integer)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Bolt Circle of Bottom Flange Bolts
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(boc.bolt_circle_bot) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(boc.bolt_circle_bot, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Bolt Orientation of Bottom Flange Bolts
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(boc.bolt_orientation_bot) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(boc.bolt_orientation_bot, Integer)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Quantity of Top Flange Bolts
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(boc.qty_flange_bolt_top) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(boc.qty_flange_bolt_top, Integer)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Bolt Circle of Top Flange Bolts
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(boc.bolt_circle_top) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(boc.bolt_circle_top, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Bolt Orientation of Top Flange Bolts
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(boc.bolt_orientation_top) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(boc.bolt_orientation_top, Integer)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Threaded Rod Diameter, Bottom Flange Connection
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(boc.threaded_rod_dia_bot) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(boc.threaded_rod_dia_bot, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Threaded Rod Material, Bottom Flange Connection
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(boc.threaded_rod_mat_bot) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(boc.threaded_rod_mat_bot, String)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Threaded Rod Quantity, Bottom Flange Connection
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(boc.threaded_rod_qty_bot) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(boc.threaded_rod_qty_bot, Integer)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Threaded Rod Unbraced Length, Bottom Flange Connection
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(boc.threaded_rod_unbraced_length_bot) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(boc.threaded_rod_unbraced_length_bot, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Threaded Rod Diameter, Bottom Flange Connection
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(boc.threaded_rod_dia_top) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(boc.threaded_rod_dia_top, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Threaded Rod Material, Top Flange Connection
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(boc.threaded_rod_mat_top) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(boc.threaded_rod_mat_top, String)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Threaded Rod Quantity, Top Flange Connection
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(boc.threaded_rod_qty_top) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(boc.threaded_rod_qty_top, Integer)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Threaded Rod Unbraced Length, Top Flange Connection
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(boc.threaded_rod_unbraced_length_top) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(boc.threaded_rod_unbraced_length_top, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Stiffener Height, Bottom Flange Connection
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(boc.stiffener_height_bot) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(boc.stiffener_height_bot, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Stiffener Length, Bottom Flange Connection
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(boc.stiffener_length_bot) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(boc.stiffener_length_bot, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Stiffener Fillet Weld Size, Bottom Flange Connection
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(boc.fillet_weld_size_bot) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(boc.fillet_weld_size_bot, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Stiffener Fillet Weld Strength, Bottom Flange Connection
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(boc.exx_bot) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(boc.exx_bot, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Flange Thickness, Bottom Flange Connection
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(boc.flange_thickness_bot) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(boc.flange_thickness_bot, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Stiffener Height, Top Flange Connection
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(boc.stiffener_height_top) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(boc.stiffener_height_top, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Stiffener Length, Top Flange Connection
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(boc.stiffener_length_top) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(boc.stiffener_length_top, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Stiffener Fillet Weld Size, Top Flange Connection
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(boc.fillet_weld_size_top) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(boc.fillet_weld_size_top, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Stiffener Fillet Weld Strength, Top Flange Connection
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(boc.exx_top) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(boc.exx_bot, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Flange Thickness, Top Flange Connection
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(boc.flange_thickness_top) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(boc.flange_thickness_top, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                Next


'                For Each arb As ArbitraryShape In lr.ArbitraryShapes

'                    'Arbitrary Shape ID
'                    colCounter = 78
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(arb.arb_shape_id) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(arb.arb_shape_id, Integer)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'US Name
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(arb.arb_shape_id) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(arb.us_name, String)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'SI Name
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(arb.si_name) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(arb.si_name, String)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Height
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(arb.height) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(arb.height, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Width
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(arb.width) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(arb.width, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Wind Projected Width
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(arb.wind_projected_width) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(arb.wind_projected_width, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Perimeter
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(arb.perimeter) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(arb.perimeter, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Modulus of Elasticity
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(arb.modulus_of_elasticity) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(arb.modulus_of_elasticity, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Density
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(arb.density) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(arb.density, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Cross-Sectional Area
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(arb.area) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(arb.area, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Local Buckling Interaction Stress Reduction Factors
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(arb.stress_reduction_factor) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(arb.stress_reduction_factor, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Warping Constant
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(arb.warp_constant) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(arb.warp_constant, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Moment of Inertia about X-X Axis
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(arb.moment_of_inertia_x) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(arb.moment_of_inertia_x, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Moment of Inertia about Y-Y Axis
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(arb.moment_of_inertia_y) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(arb.moment_of_inertia_y, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Torsional Constant
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(arb.tors_constant) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(arb.tors_constant, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Sx_top, Elastic Section Modulus about X-X Axis, Top Direction
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(arb.elastic_sec_mod_x_top) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(arb.elastic_sec_mod_x_top, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Sy_left, Elastic Section Modulus about Y-Y Axis, Left Direction? (TNX specifies this as top)
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(arb.elastic_sec_mod_y_left) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(arb.elastic_sec_mod_y_left, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Sx_bot, Elastic Section Modulus about X-X Axis, Bottom Direction
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(arb.elastic_sec_mod_x_bot) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(arb.elastic_sec_mod_x_bot, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'Sy_right, Elastic Section Modulus about Y-Y Axis, Right Direction? (TNX specifies this as top)
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(arb.elastic_sec_mod_y_right) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(arb.elastic_sec_mod_y_right, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'rx, Radius of Gyration about X-X Axis
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(arb.radius_of_gyration_x) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(arb.radius_of_gyration_x, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'ry, Radius of Gyration about Y-Y Axis
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(arb.radius_of_gyration_y) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(arb.radius_of_gyration_y, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'SFx, Shear Deflection Form Factor about X-X Axis
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(arb.shear_deflection_form_factor_x) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(arb.shear_deflection_form_factor_x, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'SFy, Shear Deflection Form Factor about Y-Y Axis
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(arb.shear_deflection_form_factor_y) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(arb.shear_deflection_form_factor_y, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'K Factor Adjustment. Allows TNX to match tool result.
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(arb.K_factor_adj) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(arb.K_factor_adj, Double)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'TNX Section File= (code to insert into .eri file)
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(arb.sec_file) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(arb.sec_file, String)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'TNX Section USName= (code to insert into .eri file)
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(arb.tnx_us_name) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(arb.tnx_us_name, String)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'TNX Section SIName= (code to insert into .eri file)
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(arb.tnx_si_name) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(arb.tnx_si_name, String)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'TNX Section Values= (code to insert into .eri file)
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(arb.sec_values) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(arb.sec_values, String)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'TNX Material Database File= (code to insert into .eri file)
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(arb.member_mat_file) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(arb.member_mat_file, String)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'TNX Material Database Name= (code to insert into .eri file)
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(arb.mat_name) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(arb.mat_name, String)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                    'TNX Material Database Values= (code to insert into .eri file)
'                    colCounter = colCounter + 1
'                    myCol = GetExcelColumnName(colCounter)
'                    If Not IsNothing(arb.mat_values) Then
'                        .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).Value = CType(arb.mat_values, String)
'                    Else .Worksheets("Database").Range(myCol & rowStart + lr.local_tool_id).ClearContents
'                    End If

'                Next


'                'specify which sections are "Yes" on IMPORT tab of workbook
'                .Worksheets("IMPORT").Range("H" & 16 + lr.local_tool_id).Value = "Yes"


'            Next


'            '~~~~~~~~POPULATE TOOL INPUTS WITH THE FIRST INSTANCE IN TOOL'S LOCAL DATABASE
'            'may not be as critical for this tool, especially if reimporting values from TNX (in case sections, therefore local tool ids change)


'        End With

'    End Sub

'    Private Function GetExcelColumnName(columnNumber As Integer) As String
'        Dim dividend As Integer = columnNumber
'        Dim columnName As String = String.Empty
'        Dim modulo As Integer

'        While dividend > 0
'            modulo = (dividend - 1) Mod 26
'            columnName = Convert.ToChar(65 + modulo).ToString() & columnName
'            dividend = CInt((dividend - modulo) / 26)
'        End While

'        Return columnName
'    End Function

'    Private Sub LoadNewLegReinforcement()
'        NewLegReinforcementWb.LoadDocument(LegReinforcementTemplatePath, LegReinforcementFileType)
'        NewLegReinforcementWb.BeginUpdate()
'    End Sub

'    Private Sub SaveAndCloseLegReinforcement()
'        NewLegReinforcementWb.EndUpdate()
'        NewLegReinforcementWb.SaveDocument(ExcelFilePath, LegReinforcementFileType)
'    End Sub

'#End Region

'#Region "SQL Insert Statements Data"
'    Private Function InsertLegReinforcementDetail(ByVal lr As LegReinforcement) As String
'        Dim insertString As String = ""

'        'insertString += "@FndID"
'        'COMPLETE ONCE SQL CODE IS UPDATED (CHRIS/IAN)

'    End Function
'#End Region

'#Region "SQL Update Statements Data"
'    'NO LONGER REQUIRED (IF WE ONLY INSERT NEW DATA)
'#End Region

'#Region "General"
'    Public Sub Clear()
'        ExcelFilePath = ""
'        LegReinforcements.Clear()
'    End Sub

'    Private Function LegReinforcementSQLDataTables() As List(Of SQLParameter)
'        Dim MyParameters As New List(Of SQLParameter)

'        MyParameters.Add(New SQLParameter("Leg Reinforcement General Details SQL", "Leg Reinforcement (SELECT Details).sql"))
'        MyParameters.Add(New SQLParameter("Leg Reinforcement Bolt-On Connection Details SQL", "Leg Reinforcement (SELECT Bolt-On Connections).sql"))
'        MyParameters.Add(New SQLParameter("Leg Reinforcement Arbitrary Shape Details SQL", "Leg Reinforcement (SELECT Arbitrary Shape).sql"))

'        Return MyParameters
'    End Function

'    Private Function LegReinforcementExcelDTParameters() As List(Of EXCELDTParameter)
'        Dim MyParameters As New List(Of EXCELDTParameter)

'        MyParameters.Add(New EXCELDTParameter("Leg Reinforcement General Details EXCEL", "V5:AU204", "Database"))
'        MyParameters.Add(New EXCELDTParameter("Leg Reinforcement Bolt-On Connection Details EXCEL", "AV5:BW204", "Database"))
'        MyParameters.Add(New EXCELDTParameter("Leg Reinforcement Arbitrary Shape Details EXCEL", "BV5:DC204", "Database"))

'        Return MyParameters
'    End Function
'#End Region

'End Class
