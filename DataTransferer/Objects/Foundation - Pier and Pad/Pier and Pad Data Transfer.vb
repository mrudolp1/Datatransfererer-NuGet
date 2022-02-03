'Option Strict Off

'Imports DevExpress.Spreadsheet
'Imports System.Security.Principal

'Partial Public Class DataTransfererPierandPad

'#Region "Define"
'    Private NewPierAndPadWb As New Workbook
'    Private prop_ExcelFilePath As String

'    Public Property PierAndPads As New List(Of PierAndPad)
'    Public Property sqlPierAndPads As New List(Of PierAndPad)
'    'Private Property PierAndPadTemplatePath As String = "C:\Users\" & Environment.UserName & "\Desktop\Pier and Pad Foundation (4.1.2) - TEMPLATE - 10-6-2021.xlsm"
'    Private Property PierAndPadTemplatePath As String = "C:\Users\" & Environment.UserName & "\Documents\.NET Testing\Foundations\Pier and Pad\Template\Pier and Pad Foundation (4.1.2) - TEMPLATE - 10-6-2021.xlsm"
'    Private Property PierAndPadFileType As DocumentFormat = DocumentFormat.Xlsm

'    'Public Property ppDS As New DataSet
'    Public Property ppDB As String
'    Public Property ppID As WindowsIdentity

'    Public Property ExcelFilePath() As String
'        Get
'            Return Me.prop_ExcelFilePath
'        End Get
'        Set
'            Me.prop_ExcelFilePath = Value
'        End Set
'    End Property
'#End Region

'#Region "Constructors"
'    Public Sub New()
'        'Leave method empty
'    End Sub

'    Public Sub New(ByVal MyDataSet As DataSet, ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String, ByVal BU As String, ByVal Strucutre_ID As String)
'        'ppDS = MyDataSet
'        ds = MyDataSet
'        ppID = LogOnUser
'        ppDB = ActiveDatabase
'        'BUNumber = BU
'        'STR_ID = Strucutre_ID
'    End Sub
'#End Region

'#Region "Load Data"

'    Sub CreateSQLPierAndPad(ByRef ppList As List(Of PierAndPad))
'        Dim refid As Integer
'        Dim PierAndPadLoader As String

'        'Load data to get pier and pad details data for the existing structure model
'        For Each item As SQLParameter In PierAndPadSQLDataTables()
'            PierAndPadLoader = QueryBuilderFromFile(queryPath & "Pier and Pad\" & item.sqlQuery).Replace("[EXISTING MODEL]", GetExistingModelQuery())
'            DoDaSQL.sqlLoader(PierAndPadLoader, item.sqlDatatable, ds, ppDB, ppID, "0")
'        Next

'        'Custom Section to transfer data for the pier and pad tool. Needs to be adjusted for each tool.
'        For Each PierAndPadDataRow As DataRow In ds.Tables("Pier and Pad General Details SQL").Rows
'            refid = CType(PierAndPadDataRow.Item("pp_id"), Integer)
'            'sqlPierAndPads.Add(New PierAndPad(PierAndPadDataRow, refid))
'            ppList.Add(New PierAndPad(PierAndPadDataRow, refid))
'        Next
'    End Sub
'    Public Function LoadFromEDS() As Boolean
'        CreateSQLPierAndPad(PierAndPads)
'        Return True
'    End Function 'Create Pier and Pad objects based on what is saved in EDS

'    Public Sub LoadFromExcel()

'        Dim refID As Integer
'        Dim refCol As String

'        For Each item As EXCELDTParameter In PierAndPadExcelDTParameters()
'            'Get additional tables from excel file 
'            ds.Tables.Add(ExcelDatasourceToDataTable(GetExcelDataSource(ExcelFilePath, item.xlsSheet, item.xlsRange), item.xlsDatatable))
'        Next


'        For Each PierandPadDataRow As DataRow In ds.Tables("Pier and Pad General Details EXCEL").Rows

'            refCol = "pp_id"
'            refID = CType(PierandPadDataRow.Item(refCol), Integer)

'            PierAndPads.Add(New PierAndPad(PierandPadDataRow, refID, refCol))

'        Next


'        'Pull SQL data, if applicable, to compare with excel data
'        CreateSQLPierAndPad(sqlPierAndPads)

'        'If sqlPiles.Count > 0 Then 'same as if checking for id in tool, if ID greater than 0.
'        For Each fnd As PierAndPad In PierAndPads
'            If fnd.pp_id > 0 Then 'can skip loading SQL data if id = 0 (first time adding to EDS)
'                For Each sqlfnd As PierAndPad In sqlPierAndPads
'                    If fnd.pp_id = sqlfnd.pp_id Then
'                        If CheckChanges(fnd, sqlfnd) Then
'                            isModelNeeded = True
'                            isfndGroupNeeded = True
'                            isPierAndPadNeeded = True
'                        End If
'                        Exit For
'                    End If
'                Next
'            Else
'                'Save the data because nothing exists in sql
'                isModelNeeded = True
'                isfndGroupNeeded = True
'                isPierAndPadNeeded = True
'            End If
'        Next

'        'Else
'        '    'Save the data because nothing exists in sql
'        '    isModelNeeded = True
'        '    isfndGroupNeeded = True
'        '    isPileNeeded = True
'        '    End If

'        'End If

'    End Sub 'Create Pier and Pad objects based on what is coming from the excel file
'#End Region

'#Region "Save Data"
'    Sub Save1PierAndPad(ByVal pp As PierAndPad)

'        Dim firstOne As Boolean = True

'        Dim PierAndPadSaver As String = QueryBuilderFromFile(queryPath & "Pier and Pad\Pier and Pad (IN_UP).sql")
'        PierAndPadSaver = PierAndPadSaver.Replace("[BU NUMBER]", BUNumber)
'        PierAndPadSaver = PierAndPadSaver.Replace("[STRUCTURE ID]", STR_ID)
'        PierAndPadSaver = PierAndPadSaver.Replace("[FOUNDATION TYPE]", "Pier and Pad")
'        If pp.pp_id = 0 Or IsDBNull(pp.pp_id) Then
'            PierAndPadSaver = PierAndPadSaver.Replace("'[PIER AND PAD ID]'", "NULL")
'        Else
'            PierAndPadSaver = PierAndPadSaver.Replace("[PIER AND PAD ID]", pp.pp_id.ToString)
'        End If

'        'Determine if new model ID needs created. Shouldn't be added to all individual tools (only needs to be referenced once)
'        If isModelNeeded Then
'            PierAndPadSaver = PierAndPadSaver.Replace("'[Model ID Needed]'", 1)
'        Else
'            PierAndPadSaver = PierAndPadSaver.Replace("'[Model ID Needed]'", 0)
'        End If

'        'Determine if new foundation group ID needs created. 
'        If isfndGroupNeeded Then
'            PierAndPadSaver = PierAndPadSaver.Replace("'[Fnd GRP ID Needed]'", 1)
'        Else
'            PierAndPadSaver = PierAndPadSaver.Replace("'[Fnd GRP ID Needed]'", 0)
'        End If

'        'Determine if new foundation ID needs created
'        If isPierAndPadNeeded Then
'            PierAndPadSaver = PierAndPadSaver.Replace("'[PIER AND PAD ID Needed]'", 1)
'        Else
'            PierAndPadSaver = PierAndPadSaver.Replace("'[PIER AND PAD ID Needed]'", 0)
'        End If

'        PierAndPadSaver = PierAndPadSaver.Replace("'[INSERT ALL PIER AND PAD DETAILS]'", InsertPierAndPadDetail(pp))

'        sqlSender(PierAndPadSaver, ppDB, ppID, "0")

'    End Sub
'    Public Sub SaveToEDS()
'        Dim firstOne As Boolean = True

'        For Each pp As PierAndPad In PierAndPads
'            Save1PierAndPad(pp)
'        Next

'    End Sub

'    Public Sub SaveToExcel()
'        For Each pp As PierAndPad In PierAndPads

'            LoadNewPierAndPad()

'            With NewPierAndPadWb
'                .Worksheets("Input").Range("ID").Value = CType(pp.pp_id, Integer)
'                If Not IsNothing(pp.pier_shape) Then
'                    .Worksheets("Input").Range("shape").Value = CType(pp.pier_shape, String)
'                End If

'                If Not IsNothing(pp.pier_diameter) Then
'                    .Worksheets("Input").Range("dpier").Value = CType(pp.pier_diameter, Double)
'                Else
'                    .Worksheets("Input").Range("dpier").ClearContents
'                End If

'                If Not IsNothing(pp.extension_above_grade) Then
'                    .Worksheets("Input").Range("E").Value = CType(pp.extension_above_grade, Double)
'                Else
'                    .Worksheets("Input").Range("E").ClearContents
'                End If

'                If Not IsNothing(pp.pier_rebar_size) Then
'                    .Worksheets("Input").Range("Sc").Value = CType(pp.pier_rebar_size, Integer)
'                Else
'                    .Worksheets("Input").Range("Sc").ClearContents
'                End If

'                If Not IsNothing(pp.pier_rebar_quantity) Then
'                    .Worksheets("Input").Range("mc").Value = CType(pp.pier_rebar_quantity, Double)
'                Else
'                    .Worksheets("Input").Range("mc").ClearContents
'                End If

'                If Not IsNothing(pp.pier_tie_size) Then
'                    .Worksheets("Input").Range("St").Value = CType(pp.pier_tie_size, Integer)
'                Else
'                    .Worksheets("Input").Range("St").ClearContents
'                End If

'                If Not IsNothing(pp.pier_tie_quantity) Then
'                    .Worksheets("Input").Range("mt").Value = CType(pp.pier_tie_quantity, Double)
'                Else
'                    .Worksheets("Input").Range("mt").ClearContents
'                End If

'                If Not IsNothing(pp.pier_reinforcement_type) Then
'                    .Worksheets("Input").Range("PierReinfType").Value = CType(pp.pier_reinforcement_type, String)
'                End If

'                If Not IsNothing(pp.pier_clear_cover) Then
'                    .Worksheets("Input").Range("ccpier").Value = CType(pp.pier_clear_cover, Double)
'                Else
'                    .Worksheets("Input").Range("ccpier").ClearContents
'                End If

'                If Not IsNothing(pp.foundation_depth) Then
'                    .Worksheets("Input").Range("D").Value = CType(pp.foundation_depth, Double)
'                Else
'                    .Worksheets("Input").Range("D").ClearContents
'                End If

'                If Not IsNothing(pp.pad_width_1) Then
'                    .Worksheets("Input").Range("W").Value = CType(pp.pad_width_1, Double)
'                Else
'                    .Worksheets("Input").Range("W").ClearContents
'                End If

'                If Not IsNothing(pp.pad_width_2) Then
'                    .Worksheets("Input").Range("W.dir2").Value = CType(pp.pad_width_2, Double)
'                Else
'                    .Worksheets("Input").Range("W.dir2").ClearContents
'                End If

'                If Not IsNothing(pp.pad_thickness) Then
'                    .Worksheets("Input").Range("T").Value = CType(pp.pad_thickness, Double)
'                Else
'                    .Worksheets("Input").Range("T").ClearContents
'                End If

'                If Not IsNothing(pp.pad_rebar_size_top_dir1) Then
'                    .Worksheets("Input").Range("sptop").Value = CType(pp.pad_rebar_size_top_dir1, Integer)
'                Else
'                    .Worksheets("Input").Range("sptop").ClearContents
'                End If

'                If Not IsNothing(pp.pad_rebar_size_bottom_dir1) Then
'                    .Worksheets("Input").Range("Sp").Value = CType(pp.pad_rebar_size_bottom_dir1, Integer)
'                Else
'                    .Worksheets("Input").Range("Sp").ClearContents
'                End If

'                If Not IsNothing(pp.pad_rebar_size_top_dir2) Then
'                    .Worksheets("Input").Range("sptop2").Value = CType(pp.pad_rebar_size_top_dir2, Integer)
'                Else
'                    .Worksheets("Input").Range("sptop2").ClearContents
'                End If

'                If Not IsNothing(pp.pad_rebar_size_bottom_dir2) Then
'                    .Worksheets("Input").Range("sp_2").Value = CType(pp.pad_rebar_size_bottom_dir2, Integer)
'                Else
'                    .Worksheets("Input").Range("sp_2").ClearContents
'                End If

'                If Not IsNothing(pp.pad_rebar_quantity_top_dir1) Then
'                    .Worksheets("Input").Range("mptop").Value = CType(pp.pad_rebar_quantity_top_dir1, Double)
'                Else
'                    .Worksheets("Input").Range("mptop").ClearContents
'                End If

'                If Not IsNothing(pp.pad_rebar_quantity_bottom_dir1) Then
'                    .Worksheets("Input").Range("mp").Value = CType(pp.pad_rebar_quantity_bottom_dir1, Double)
'                Else
'                    .Worksheets("Input").Range("mp").ClearContents
'                End If

'                If Not IsNothing(pp.pad_rebar_quantity_top_dir2) Then
'                    .Worksheets("Input").Range("mptop2").Value = CType(pp.pad_rebar_quantity_top_dir2, Double)
'                Else
'                    .Worksheets("Input").Range("mptop2").ClearContents
'                End If

'                If Not IsNothing(pp.pad_rebar_quantity_bottom_dir2) Then
'                    .Worksheets("Input").Range("mp_2").Value = CType(pp.pad_rebar_quantity_bottom_dir2, Double)
'                Else
'                    .Worksheets("Input").Range("mp_2").ClearContents
'                End If

'                If Not IsNothing(pp.pad_clear_cover) Then
'                    .Worksheets("Input").Range("ccpad").Value = CType(pp.pad_clear_cover, Double)
'                Else
'                    .Worksheets("Input").Range("ccpad").ClearContents
'                End If

'                If Not IsNothing(pp.rebar_grade) Then
'                    .Worksheets("Input").Range("Fy").Value = CType(pp.rebar_grade, Double)
'                Else
'                    .Worksheets("Input").Range("Fy").ClearContents
'                End If

'                If Not IsNothing(pp.concrete_compressive_strength) Then
'                    .Worksheets("Input").Range("F\c").Value = CType(pp.concrete_compressive_strength, Double)
'                Else
'                    .Worksheets("Input").Range("F\c").ClearContents
'                End If

'                If Not IsNothing(pp.dry_concrete_density) Then
'                    .Worksheets("Input").Range("ConcreteDensity").Value = CType(pp.dry_concrete_density, Double)
'                Else
'                    .Worksheets("Input").Range("ConcreteDensity").ClearContents
'                End If

'                If Not IsNothing(pp.total_soil_unit_weight) Then
'                    .Worksheets("Input").Range("γ").Value = CType(pp.total_soil_unit_weight, Double)
'                Else
'                    .Worksheets("Input").Range("γ").ClearContents
'                End If

'                If Not IsNothing(pp.bearing_type) Then
'                    .Worksheets("Input").Range("BearingType").Value = CType(pp.bearing_type, String)
'                Else
'                    .Worksheets("Input").Range("BearingType").ClearContents
'                End If

'                If Not IsNothing(pp.nominal_bearing_capacity) Then
'                    .Worksheets("Input").Range("Qinput").Value = CType(pp.nominal_bearing_capacity, Double)
'                Else
'                    .Worksheets("Input").Range("Qinput").ClearContents
'                End If

'                If Not IsNothing(pp.cohesion) Then
'                    .Worksheets("Input").Range("Cu").Value = CType(pp.cohesion, Double)
'                Else
'                    .Worksheets("Input").Range("Cu").ClearContents
'                End If

'                If Not IsNothing(pp.friction_angle) Then
'                    .Worksheets("Input").Range("ϕ").Value = CType(pp.friction_angle, Double)
'                Else
'                    .Worksheets("Input").Range("ϕ").ClearContents
'                End If

'                If Not IsNothing(pp.spt_blow_count) Then
'                    .Worksheets("Input").Range("N_blows").Value = CType(pp.spt_blow_count, Double)
'                Else
'                    .Worksheets("Input").Range("N_blows").ClearContents
'                End If

'                If Not IsNothing(pp.base_friction_factor) Then
'                    .Worksheets("Input").Range("μ").Value = CType(pp.base_friction_factor, Double)
'                Else
'                    .Worksheets("Input").Range("μ").ClearContents
'                End If

'                If Not IsNothing(pp.neglect_depth) Then
'                    .Worksheets("Input").Range("N").Value = CType(pp.neglect_depth, Double)
'                End If

'                If pp.bearing_distribution_type = False Then
'                    .Worksheets("Input").Range("Rock").Value = "No"
'                Else
'                    .Worksheets("Input").Range("Rock").Value = "Yes"
'                End If

'                If pp.groundwater_depth = -1 Then
'                    .Worksheets("Input").Range("gw").Value = "N/A"
'                Else
'                    .Worksheets("Input").Range("gw").Value = CType(pp.groundwater_depth, Double)
'                End If

'                If Not IsNothing(pp.top_and_bottom_rebar_different) Then
'                    .Worksheets("Input").Range("DifferentReinforcementBoolean").Value = CType(pp.top_and_bottom_rebar_different, Boolean)
'                End If

'                If Not IsNothing(pp.block_foundation) Then
'                    .Worksheets("Input").Range("BlockFoundationBoolean").Value = CType(pp.block_foundation, Boolean)
'                End If

'                If Not IsNothing(pp.rectangular_foundation) Then
'                    .Worksheets("Input").Range("RectangularPadBoolean").Value = CType(pp.rectangular_foundation, Boolean)
'                End If

'                If Not IsNothing(pp.base_plate_distance_above_foundation) Then
'                    .Worksheets("Input").Range("bpdist").Value = CType(pp.base_plate_distance_above_foundation, Double)
'                Else
'                    .Worksheets("Input").Range("bpdist").ClearContents
'                End If

'                If Not IsNothing(pp.bolt_circle_bearing_plate_width) Then
'                    .Worksheets("Input").Range("BC").Value = CType(pp.bolt_circle_bearing_plate_width, Double)
'                Else
'                    .Worksheets("Input").Range("BC").ClearContents
'                End If

'                If Not IsNothing(pp.basic_soil_check) Then
'                    .Worksheets("Input").Range("SoilInteractionBoolean").Value = CType(pp.basic_soil_check, Boolean)
'                End If

'                If Not IsNothing(pp.structural_check) Then
'                    .Worksheets("Input").Range("StructuralCheckBoolean").Value = CType(pp.structural_check, Boolean)
'                End If

'            End With

'            SaveAndClosePierAndPad()
'        Next

'    End Sub

'    Private Sub LoadNewPierAndPad()
'        NewPierAndPadWb.LoadDocument(PierAndPadTemplatePath, PierAndPadFileType)
'        NewPierAndPadWb.BeginUpdate()
'    End Sub

'    Private Sub SaveAndClosePierAndPad()
'        NewPierAndPadWb.Calculate()
'        NewPierAndPadWb.EndUpdate()
'        NewPierAndPadWb.SaveDocument(ExcelFilePath, PierAndPadFileType)
'    End Sub
'#End Region

'#Region "SQL Insert Statements"
'    Private Function InsertPierAndPadDetail(ByVal pp As PierAndPad) As String
'        Dim insertString As String = ""

'        insertString += IIf(IsNothing(pp.pier_shape), "Null", "'" & pp.pier_shape.ToString & "'")
'        insertString += "," & IIf(IsNothing(pp.pier_diameter), "Null", pp.pier_diameter.ToString)
'        insertString += "," & IIf(IsNothing(pp.extension_above_grade), "Null", pp.extension_above_grade.ToString)
'        insertString += "," & IIf(IsNothing(pp.pier_rebar_size), "Null", pp.pier_rebar_size.ToString)
'        insertString += "," & IIf(IsNothing(pp.pier_tie_size), "Null", pp.pier_tie_size.ToString)
'        insertString += "," & IIf(IsNothing(pp.pier_tie_quantity), "Null", pp.pier_tie_quantity.ToString)
'        insertString += "," & IIf(IsNothing(pp.pier_reinforcement_type), "Null", "'" & pp.pier_reinforcement_type.ToString & "'")
'        insertString += "," & IIf(IsNothing(pp.pier_clear_cover), "Null", pp.pier_clear_cover.ToString)
'        insertString += "," & IIf(IsNothing(pp.foundation_depth), "Null", pp.foundation_depth.ToString)
'        insertString += "," & IIf(IsNothing(pp.pad_width_1), "Null", pp.pad_width_1.ToString)
'        insertString += "," & IIf(IsNothing(pp.pad_width_2), "Null", pp.pad_width_2.ToString)
'        insertString += "," & IIf(IsNothing(pp.pad_thickness), "Null", pp.pad_thickness.ToString)
'        insertString += "," & IIf(IsNothing(pp.pad_rebar_size_top_dir1), "Null", pp.pad_rebar_size_top_dir1.ToString)
'        insertString += "," & IIf(IsNothing(pp.pad_rebar_size_bottom_dir1), "Null", pp.pad_rebar_size_bottom_dir1.ToString)
'        insertString += "," & IIf(IsNothing(pp.pad_rebar_size_top_dir2), "Null", pp.pad_rebar_size_top_dir2.ToString)
'        insertString += "," & IIf(IsNothing(pp.pad_rebar_size_bottom_dir2), "Null", pp.pad_rebar_size_bottom_dir2.ToString)
'        insertString += "," & IIf(IsNothing(pp.pad_rebar_quantity_top_dir1), "Null", pp.pad_rebar_quantity_top_dir1.ToString)
'        insertString += "," & IIf(IsNothing(pp.pad_rebar_quantity_bottom_dir1), "Null", pp.pad_rebar_quantity_bottom_dir1.ToString)
'        insertString += "," & IIf(IsNothing(pp.pad_rebar_quantity_top_dir2), "Null", pp.pad_rebar_quantity_top_dir2.ToString)
'        insertString += "," & IIf(IsNothing(pp.pad_rebar_quantity_bottom_dir2), "Null", pp.pad_rebar_quantity_bottom_dir2.ToString)
'        insertString += "," & IIf(IsNothing(pp.pad_clear_cover), "Null", pp.pad_clear_cover.ToString)
'        insertString += "," & IIf(IsNothing(pp.rebar_grade), "Null", pp.rebar_grade.ToString)
'        insertString += "," & IIf(IsNothing(pp.concrete_compressive_strength), "Null", pp.concrete_compressive_strength.ToString)
'        insertString += "," & IIf(IsNothing(pp.dry_concrete_density), "Null", pp.dry_concrete_density.ToString)
'        insertString += "," & IIf(IsNothing(pp.total_soil_unit_weight), "Null", pp.total_soil_unit_weight.ToString)
'        insertString += "," & IIf(IsNothing(pp.bearing_type), "Null", "'" & pp.bearing_type.ToString & "'")
'        insertString += "," & IIf(IsNothing(pp.nominal_bearing_capacity), "Null", pp.nominal_bearing_capacity.ToString)
'        insertString += "," & IIf(IsNothing(pp.cohesion), "Null", pp.cohesion.ToString)
'        insertString += "," & IIf(IsNothing(pp.friction_angle), "Null", pp.friction_angle.ToString)
'        insertString += "," & IIf(IsNothing(pp.spt_blow_count), "Null", pp.spt_blow_count.ToString)
'        insertString += "," & IIf(IsNothing(pp.base_friction_factor), "Null", pp.base_friction_factor.ToString)
'        insertString += "," & IIf(IsNothing(pp.neglect_depth), "Null", pp.neglect_depth.ToString)
'        insertString += "," & IIf(IsNothing(pp.bearing_distribution_type), "Null", "'" & pp.bearing_distribution_type.ToString & "'")
'        insertString += "," & IIf(IsNothing(pp.groundwater_depth), "Null", pp.groundwater_depth.ToString)
'        insertString += "," & IIf(IsNothing(pp.top_and_bottom_rebar_different), "Null", "'" & pp.top_and_bottom_rebar_different.ToString & "'")
'        insertString += "," & IIf(IsNothing(pp.block_foundation), "Null", "'" & pp.block_foundation.ToString & "'")
'        insertString += "," & IIf(IsNothing(pp.rectangular_foundation), "Null", "'" & pp.rectangular_foundation.ToString & "'")
'        insertString += "," & IIf(IsNothing(pp.base_plate_distance_above_foundation), "Null", pp.base_plate_distance_above_foundation.ToString)
'        insertString += "," & IIf(IsNothing(pp.bolt_circle_bearing_plate_width), "Null", pp.bolt_circle_bearing_plate_width.ToString)
'        insertString += "," & IIf(IsNothing(pp.pier_rebar_quantity), "Null", pp.pier_rebar_quantity.ToString)
'        insertString += "," & IIf(IsNothing(pp.basic_soil_check), "Null", "'" & pp.basic_soil_check.ToString & "'")
'        insertString += "," & IIf(IsNothing(pp.structural_check), "Null", "'" & pp.structural_check.ToString & "'")
'        insertString += "," & IIf(IsNothing(pp.tool_version), "Null", "'" & pp.tool_version.ToString & "'")

'        Return insertString
'    End Function
'#End Region


'#Region "General"
'Public Sub Clear()
'    ExcelFilePath = ""
'    PierAndPads.Clear()
'End Sub

'    Private Function PierAndPadSQLDataTables() As List(Of SQLParameter)
'        Dim MyParameters As New List(Of SQLParameter)

'        MyParameters.Add(New SQLParameter("Pier and Pad General Details SQL", "Pier and Pad (SELECT Details).sql"))
'        'MyParameters.Add(New SQLParameter("Pier and Pad Modified Ranges SQL", "Pier and Pad (SELECT Modified Ranges).sql"))

'        Return MyParameters
'    End Function

'    Private Function PierAndPadExcelDTParameters() As List(Of EXCELDTParameter)
'        Dim MyParameters As New List(Of EXCELDTParameter)

'        MyParameters.Add(New EXCELDTParameter("Pier and Pad General Details EXCEL", "A2:AR3", "Details (SAPI)"))
'        'MyParameters.Add(New EXCELDTParameter("Pier and Pad Modified Ranges EXCEL", "A1:E1000", "Modified Ranges"))

'        Return MyParameters
'    End Function

'#End Region

'#Region "Check Changes"
'    'Private changeDt As New DataTable
'    'Private changeList As New List(Of AnalysisChanges)
'    Function CheckChanges(ByVal xlPierAndPad As PierAndPad, ByVal sqlPierAndPad As PierAndPad) As Boolean
'        Dim changesMade As Boolean = False

'        'changeDt.Columns.Add("Variable", Type.GetType("System.String"))
'        'changeDt.Columns.Add("New Value", Type.GetType("System.String"))
'        'changeDt.Columns.Add("Previous Value", Type.GetType("System.String"))
'        'changeDt.Columns.Add("WO", Type.GetType("System.String"))

'        'Check Details
'        If Check1Change(xlPierAndPad.pier_shape, sqlPierAndPad.pier_shape, "Pier and Pad", "Pier_Shape") Then changesMade = True
'        If Check1Change(xlPierAndPad.pier_diameter, sqlPierAndPad.pier_diameter, "Pier and Pad", "Pier_Diameter") Then changesMade = True
'        If Check1Change(xlPierAndPad.extension_above_grade, sqlPierAndPad.extension_above_grade, "Pier and Pad", "Extension_Above_Grade") Then changesMade = True
'        If Check1Change(xlPierAndPad.pier_rebar_size, sqlPierAndPad.pier_rebar_size, "Pier and Pad", "Pier_Rebar_Size") Then changesMade = True
'        If Check1Change(xlPierAndPad.pier_tie_size, sqlPierAndPad.pier_tie_size, "Pier and Pad", "Pier_Tie_Size") Then changesMade = True
'        If Check1Change(xlPierAndPad.pier_tie_quantity, sqlPierAndPad.pier_tie_quantity, "Pier and Pad", "Pier_Tie_Quantity") Then changesMade = True
'        If Check1Change(xlPierAndPad.pier_reinforcement_type, sqlPierAndPad.pier_reinforcement_type, "Pier and Pad", "Pier_Reinforcement_Type") Then changesMade = True
'        If Check1Change(xlPierAndPad.pier_clear_cover, sqlPierAndPad.pier_clear_cover, "Pier and Pad", "Pier_Clear_Cover") Then changesMade = True
'        If Check1Change(xlPierAndPad.foundation_depth, sqlPierAndPad.foundation_depth, "Pier and Pad", "Foundation_Depth") Then changesMade = True
'        If Check1Change(xlPierAndPad.pad_width_1, sqlPierAndPad.pad_width_1, "Pier and Pad", "Pad_Width_1") Then changesMade = True
'        If Check1Change(xlPierAndPad.pad_width_2, sqlPierAndPad.pad_width_2, "Pier and Pad", "Pad_Width_2") Then changesMade = True
'        If Check1Change(xlPierAndPad.pad_thickness, sqlPierAndPad.pad_thickness, "Pier and Pad", "Pad_Thickness") Then changesMade = True
'        If Check1Change(xlPierAndPad.pad_rebar_size_top_dir1, sqlPierAndPad.pad_rebar_size_top_dir1, "Pier and Pad", "Pad_Rebar_Size_Top_Dir1") Then changesMade = True
'        If Check1Change(xlPierAndPad.pad_rebar_size_bottom_dir1, sqlPierAndPad.pad_rebar_size_bottom_dir1, "Pier and Pad", "Pad_Rebar_Size_Bottom_Dir1") Then changesMade = True
'        If Check1Change(xlPierAndPad.pad_rebar_size_top_dir2, sqlPierAndPad.pad_rebar_size_top_dir2, "Pier and Pad", "Pad_Rebar_Size_Top_Dir2") Then changesMade = True
'        If Check1Change(xlPierAndPad.pad_rebar_size_bottom_dir2, sqlPierAndPad.pad_rebar_size_bottom_dir2, "Pier and Pad", "Pad_Rebar_Size_Bottom_Dir2") Then changesMade = True
'        If Check1Change(xlPierAndPad.pad_rebar_quantity_top_dir1, sqlPierAndPad.pad_rebar_quantity_top_dir1, "Pier and Pad", "Pad_Rebar_Quantity_Top_Dir1") Then changesMade = True
'        If Check1Change(xlPierAndPad.pad_rebar_quantity_bottom_dir1, sqlPierAndPad.pad_rebar_quantity_bottom_dir1, "Pier and Pad", "Pad_Rebar_Quantity_Bottom_Dir1") Then changesMade = True
'        If Check1Change(xlPierAndPad.pad_rebar_quantity_top_dir2, sqlPierAndPad.pad_rebar_quantity_top_dir2, "Pier and Pad", "Pad_Rebar_Quantity_Top_Dir2") Then changesMade = True
'        If Check1Change(xlPierAndPad.pad_rebar_quantity_bottom_dir2, sqlPierAndPad.pad_rebar_quantity_bottom_dir2, "Pier and Pad", "Pad_Rebar_Quantity_Bottom_Dir2") Then changesMade = True
'        If Check1Change(xlPierAndPad.pad_clear_cover, sqlPierAndPad.pad_clear_cover, "Pier and Pad", "Pad_Clear_Cover") Then changesMade = True
'        If Check1Change(xlPierAndPad.rebar_grade, sqlPierAndPad.rebar_grade, "Pier and Pad", "Rebar_Grade") Then changesMade = True
'        If Check1Change(xlPierAndPad.concrete_compressive_strength, sqlPierAndPad.concrete_compressive_strength, "Pier and Pad", "Concrete_Compressive_Strength") Then changesMade = True
'        If Check1Change(xlPierAndPad.dry_concrete_density, sqlPierAndPad.dry_concrete_density, "Pier and Pad", "Dry_Concrete_Density") Then changesMade = True
'        If Check1Change(xlPierAndPad.total_soil_unit_weight, sqlPierAndPad.total_soil_unit_weight, "Pier and Pad", "Total_Soil_Unit_Weight") Then changesMade = True
'        If Check1Change(xlPierAndPad.bearing_type, sqlPierAndPad.bearing_type, "Pier and Pad", "Bearing_Type") Then changesMade = True
'        If Check1Change(xlPierAndPad.nominal_bearing_capacity, sqlPierAndPad.nominal_bearing_capacity, "Pier and Pad", "Nominal_Bearing_Capacity") Then changesMade = True
'        If Check1Change(xlPierAndPad.cohesion, sqlPierAndPad.cohesion, "Pier and Pad", "Cohesion") Then changesMade = True
'        If Check1Change(xlPierAndPad.friction_angle, sqlPierAndPad.friction_angle, "Pier and Pad", "Friction_Angle") Then changesMade = True
'        If Check1Change(xlPierAndPad.spt_blow_count, sqlPierAndPad.spt_blow_count, "Pier and Pad", "Spt_Blow_Count") Then changesMade = True
'        If Check1Change(xlPierAndPad.base_friction_factor, sqlPierAndPad.base_friction_factor, "Pier and Pad", "Base_Friction_Factor") Then changesMade = True
'        If Check1Change(xlPierAndPad.neglect_depth, sqlPierAndPad.neglect_depth, "Pier and Pad", "Neglect_Depth") Then changesMade = True
'        If Check1Change(xlPierAndPad.bearing_distribution_type, sqlPierAndPad.bearing_distribution_type, "Pier and Pad", "Bearing_Distribution_Type") Then changesMade = True
'        If Check1Change(xlPierAndPad.groundwater_depth, sqlPierAndPad.groundwater_depth, "Pier and Pad", "Groundwater_Depth") Then changesMade = True
'        If Check1Change(xlPierAndPad.top_and_bottom_rebar_different, sqlPierAndPad.top_and_bottom_rebar_different, "Pier and Pad", "Top_And_Bottom_Rebar_Different") Then changesMade = True
'        If Check1Change(xlPierAndPad.block_foundation, sqlPierAndPad.block_foundation, "Pier and Pad", "Block_Foundation") Then changesMade = True
'        If Check1Change(xlPierAndPad.rectangular_foundation, sqlPierAndPad.rectangular_foundation, "Pier and Pad", "Rectangular_Foundation") Then changesMade = True
'        If Check1Change(xlPierAndPad.base_plate_distance_above_foundation, sqlPierAndPad.base_plate_distance_above_foundation, "Pier and Pad", "Base_Plate_Distance_Above_Foundation") Then changesMade = True
'        If Check1Change(xlPierAndPad.bolt_circle_bearing_plate_width, sqlPierAndPad.bolt_circle_bearing_plate_width, "Pier and Pad", "Bolt_Circle_Bearing_Plate_Width") Then changesMade = True
'        If Check1Change(xlPierAndPad.pier_rebar_quantity, sqlPierAndPad.pier_rebar_quantity, "Pier and Pad", "Pier_Rebar_Quantity") Then changesMade = True
'        If Check1Change(xlPierAndPad.basic_soil_check, sqlPierAndPad.basic_soil_check, "Pier and Pad", "Basic_Soil_Check") Then changesMade = True
'        If Check1Change(xlPierAndPad.structural_check, sqlPierAndPad.structural_check, "Pier and Pad", "Structural_Check") Then changesMade = True
'        If Check1Change(xlPierAndPad.tool_version, sqlPierAndPad.tool_version, "Pier and Pad", "Tool_Version") Then changesMade = True

'        CreateChangeSummary(changeDt) 'possible alternative to listing change summary
'        Return changesMade

'    End Function

'    'Function CreateChangeSummary(ByVal changeDt As DataTable) As String
'    '    'Sub CreateChangeSummary(ByVal changeDt As DataTable)
'    '    'Create your string based on data in the datatable
'    '    Dim summary As String
'    '    Dim counter As Integer = 0

'    '    For Each chng As AnalysisChanges In changeList
'    '        If counter = 0 Then
'    '            summary += chng.Name & " = " & chng.NewValue & " | Previously: " & chng.PreviousValue
'    '        Else
'    '            summary += vbNewLine & chng.Name & " = " & chng.NewValue & " | Previously: " & chng.PreviousValue
'    '        End If

'    '        counter += 1
'    '    Next

'    '    'write to text file
'    '    'End Sub
'    'End Function

'    'Function Check1Change(ByVal newValue As Object, ByVal oldvalue As Object, ByVal tolerance As Double, ByVal variable As String) As Boolean
'    '    If newValue <> oldvalue Then
'    '        changeDt.Rows.Add(variable, newValue, oldvalue, CurWO) 'Need to determine what we want to store in this datatable or list (Foundation Type, Foundation ID)?
'    '        changeList.Add(New AnalysisChanges(oldvalue, newValue, variable, "Pier and Pad Foundations"))
'    '        Return True
'    '    ElseIf Not IsNothing(newValue) And IsNothing(oldvalue) Then 'accounts for when new rows are added. New rows from excel=0 where sql=nothing
'    '        changeDt.Rows.Add(variable, newValue, oldvalue, CurWO) 'Need to determine what we want to store in this datatable or list (Foundation Type, Foundation ID)?
'    '        changeList.Add(New AnalysisChanges(oldvalue, newValue, variable, "Pier and Pad Foundations"))
'    '        Return True
'    '    ElseIf IsNothing(newValue) And Not IsNothing(oldvalue) Then 'accounts for when rows are removed. Rows from excel=nothing where sql=value
'    '        changeDt.Rows.Add(variable, newValue, oldvalue, CurWO) 'Need to determine what we want to store in this datatable or list (Foundation Type, Foundation ID)?
'    '        changeList.Add(New AnalysisChanges(oldvalue, newValue, variable, "Pier and Pad Foundations"))
'    '        Return True
'    '    End If
'    'End Function
'#End Region

'End Class


''Class PierAndPadAnalysisChanges
''    Property PreviousValue As String
''    Property NewValue As String
''    Property Name As String
''    Property PartofDatabase As String

''    Public Sub New(prev As String, Newval As String, name As String, db As String)
''        Me.PreviousValue = prev
''        Me.NewValue = Newval
''        Me.Name = name
''        Me.PartofDatabase = db
''    End Sub
''End Class