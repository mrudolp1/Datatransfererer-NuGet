'Option Strict Off

'Imports DevExpress.Spreadsheet
'Imports System.Security.Principal

'Partial Public Class DataTransfererUnitBase

'#Region "Define"
'    Private NewUnitBaseWb As New Workbook
'    Private prop_ExcelFilePath As String

'    Public Property UnitBases As New List(Of SST_Unit_Base)
'    Public Property sqlUnitBases As New List(Of SST_Unit_Base)
'    Private Property UnitBaseTemplatePath As String = "C:\Users\" & Environment.UserName & "\source\repos\Datatransferer NuGet\Reference\SST Unit Base Foundation (4.0.4) - TEMPLATE.xlsm"
'    Private Property UnitBaseFileType As DocumentFormat = DocumentFormat.Xlsm

'    'Public Property ubDS As New DataSet
'    Public Property ubDB As String
'    Public Property ubID As WindowsIdentity

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
'        'ubDS = MyDataSet
'        ds = MyDataSet
'        ubID = LogOnUser
'        ubDB = ActiveDatabase
'        'BUNumber = BU
'        'STR_ID = Strucutre_ID
'    End Sub
'#End Region

'#Region "Load Data"
'    Sub CreateSQLUnitBase(ByRef UnitBaseList As List(Of SST_Unit_Base))
'        Dim refid As Integer
'        Dim UnitBaseLoader As String

'        'Load data to get Unit Base details for the existing structure model
'        For Each item As SQLParameter In UnitBaseSQLDataTables()
'            UnitBaseLoader = QueryBuilderFromFile(queryPath & "Unit Base\" & item.sqlQuery).Replace("[EXISTING MODEL]", GetExistingModelQuery())
'            DoDaSQL.sqlLoader(UnitBaseLoader, item.sqlDatatable, ds, ubDB, ubID, "0")
'        Next

'        'Custom Section to transfer data for the Unit Base tool. Needs to be adjusted for each tool.
'        For Each UnitBaseDataRow As DataRow In ds.Tables("Unit Base General Details SQL").Rows
'            refid = CType(UnitBaseDataRow.Item("unit_base_id"), Integer)
'            UnitBases.Add(New SST_Unit_Base(UnitBaseDataRow, refid)) 'UnitBaseList?
'        Next

'    End Sub

'    Public Function LoadFromEDS() As Boolean
'        'Dim refid As Integer
'        'Dim UnitBaseLoader As String

'        ''Load data to get Unit Base details for the existing structure model
'        'For Each item As SQLParameter In UnitBaseSQLDataTables()
'        '    UnitBaseLoader = QueryBuilderFromFile(queryPath & "Unit Base\" & item.sqlQuery).Replace("[EXISTING MODEL]", GetExistingModelQuery())
'        '    'DoDaSQL.sqlLoader(UnitBaseLoader, item.sqlDatatable, ubDS, ubDB, ubID, "0")
'        '    DoDaSQL.sqlLoader(UnitBaseLoader, item.sqlDatatable, ds, ubDB, ubID, "0")
'        '    'If ubDS.Tables(item.sqlDatatable).Rows.Count = 0 Then Return False
'        'Next

'        ''Custom Section to transfer data for the drilled pier tool. Needs to be adjusted for each tool.
'        ''For Each UnitBaseDataRow As DataRow In ubDS.Tables("Unit Base General Details SQL").Rows
'        'For Each UnitBaseDataRow As DataRow In ds.Tables("Unit Base General Details SQL").Rows
'        '    refid = CType(UnitBaseDataRow.Item("unit_base_id"), Integer)

'        '    UnitBases.Add(New SST_Unit_Base(UnitBaseDataRow, refid))
'        'Next
'        CreateSQLUnitBase(UnitBases)
'        Return True
'    End Function 'Create Unit Base objects based on what is saved in EDS

'    Public Sub LoadFromExcel()

'        For Each item As EXCELDTParameter In UnitBaseExcelDTParameters()
'            'Get additional tables from excel file 
'            ds.Tables.Add(ExcelDatasourceToDataTable(GetExcelDataSource(ExcelFilePath, item.xlsSheet, item.xlsRange), item.xlsDatatable))
'        Next

'        UnitBases.Add(New SST_Unit_Base(ExcelFilePath)) 'Option 1: Connect to excel and pull values from cells

'        'Dim refID As Integer
'        'Dim refCol As String
'        'For Each UnitBaseDataRow As DataRow In ds.Tables("Unit Base General Details EXCEL").Rows 'Option 2: Connect to excel and pull datarow from SAPI tab
'        '    refCol = "unit_base_id"
'        '    refID = CType(UnitBaseDataRow.Item(refCol), Integer)
'        '    UnitBases.Add(New SST_Unit_Base(UnitBaseDataRow, refID, refCol))
'        'Next

'        'Pull SQL data, if applicable, to compare with excel data
'        CreateSQLUnitBase(sqlUnitBases)

'        'If sqlUnitBases.Count > 0 Then 'same as if checking for id in tool, if ID greater than 0.
'        For Each fnd As SST_Unit_Base In UnitBases
'            If fnd.unit_base_id > 0 Then 'can skip loading SQL data if id = 0 (first time adding to EDS)
'                For Each sqlfnd As SST_Unit_Base In UnitBases
'                    If fnd.unit_base_id = sqlfnd.unit_base_id Then
'                        If CheckChanges(fnd, sqlfnd) Then
'                            isModelNeeded = True
'                            isfndGroupNeeded = True
'                            isUnitBaseNeeded = True
'                        End If
'                        Exit For
'                    End If
'                Next
'            Else
'                'Save the data because nothing exists in sql
'                isModelNeeded = True
'                isfndGroupNeeded = True
'                isUnitBaseNeeded = True
'            End If
'        Next

'    End Sub 'Create Unit Base objects based on what is coming from the excel file
'#End Region

'#Region "Save Data"

'    Sub Save1UnitBase(ByVal ub As SST_Unit_Base)

'        Dim firstOne As Boolean = True
'        Dim mySoils As String = ""
'        Dim myLocations As String = ""

'        Dim UnitBaseSaver As String = QueryBuilderFromFile(queryPath & "Unit Base\Unit Base (IN_UP).sql")
'        UnitBaseSaver = UnitBaseSaver.Replace("[BU NUMBER]", BUNumber)
'        UnitBaseSaver = UnitBaseSaver.Replace("[STRUCTURE ID]", STR_ID)
'        UnitBaseSaver = UnitBaseSaver.Replace("[FOUNDATION TYPE]", "Unit Base")
'        If ub.unit_base_id = 0 Or IsDBNull(ub.unit_base_id) Then
'            UnitBaseSaver = UnitBaseSaver.Replace("'[UNIT BASE ID]'", "NULL")
'        Else
'            UnitBaseSaver = UnitBaseSaver.Replace("[UNIT BASE ID]", ub.unit_base_id.ToString)
'        End If

'        'Determine if new model ID needs created. Shouldn't be added to all individual tools (only needs to be referenced once)
'        If isModelNeeded Then
'            UnitBaseSaver = UnitBaseSaver.Replace("'[Model ID Needed]'", 1)
'        Else
'            UnitBaseSaver = UnitBaseSaver.Replace("'[Model ID Needed]'", 0)
'        End If

'        'Determine if new foundation group ID needs created. 
'        If isfndGroupNeeded Then
'            UnitBaseSaver = UnitBaseSaver.Replace("'[Fnd GRP ID Needed]'", 1)
'        Else
'            UnitBaseSaver = UnitBaseSaver.Replace("'[Fnd GRP ID Needed]'", 0)
'        End If

'        'Determine if new ID needs created
'        If isUnitBaseNeeded Then
'            UnitBaseSaver = UnitBaseSaver.Replace("'[UNIT BASE ID Needed]'", 1)
'        Else
'            UnitBaseSaver = UnitBaseSaver.Replace("'[UNIT BASE ID Needed]'", 0)
'        End If

'        UnitBaseSaver = UnitBaseSaver.Replace("'[INSERT ALL UNIT BASE DETAILS]'", InsertUnitBaseDetail(ub))

'        sqlSender(UnitBaseSaver, ubDB, ubID, "0")
'    End Sub

'    Public Sub SaveToEDS()
'        For Each ub As SST_Unit_Base In UnitBases
'            Save1UnitBase(ub)
'            'Dim UnitBaseSaver As String = Common.QueryBuilderFromFile(queryPath & "Unit Base\Unit Base (IN_UP).sql")

'            'UnitBaseSaver = UnitBaseSaver.Replace("[BU NUMBER]", BUNumber)
'            'UnitBaseSaver = UnitBaseSaver.Replace("[STRUCTURE ID]", STR_ID)
'            'UnitBaseSaver = UnitBaseSaver.Replace("[FOUNDATION TYPE]", "Unit Base")
'            'If ub.unit_base_id = 0 Or IsDBNull(ub.unit_base_id) Then
'            '    UnitBaseSaver = UnitBaseSaver.Replace("'[UNIT BASE ID]'", "NULL")
'            'Else
'            '    UnitBaseSaver = UnitBaseSaver.Replace("[UNIT BASE ID]", ub.unit_base_id.ToString)
'            '    UnitBaseSaver = UnitBaseSaver.Replace("(SELECT * FROM TEMPORARY)", UpdateUnitBaseDetail(ub))
'            'End If
'            'UnitBaseSaver = UnitBaseSaver.Replace("[INSERT ALL UNIT BASE DETAILS]", InsertUnitBaseDetail(ub))

'            'sqlSender(UnitBaseSaver, ubDB, ubID, "0")
'        Next
'    End Sub

'    Public Sub SaveToExcel()
'        For Each ub As SST_Unit_Base In UnitBases
'            LoadNewUnitBase()
'            With NewUnitBaseWb
'                .Worksheets("Input").Range("ID").Value = CType(ub.unit_base_id, Integer)
'                If Not IsNothing(ub.extension_above_grade) Then .Worksheets("Input").Range("E").Value = CType(ub.extension_above_grade, Double) Else .Worksheets("Input").Range("E").ClearContents
'                If Not IsNothing(ub.foundation_depth) Then .Worksheets("Input").Range("D").Value = CType(ub.foundation_depth, Double) Else .Worksheets("Input").Range("D").ClearContents
'                If Not IsNothing(ub.concrete_compressive_strength) Then .Worksheets("Input").Range("F\c").Value = CType(ub.concrete_compressive_strength, Double) Else .Worksheets("Input").Range("F\c").ClearContents
'                If Not IsNothing(ub.dry_concrete_density) Then .Worksheets("Input").Range("ConcreteDensity").Value = CType(ub.dry_concrete_density, Double) Else .Worksheets("Input").Range("ConcreteDensity").ClearContents
'                If Not IsNothing(ub.rebar_grade) Then .Worksheets("Input").Range("Fy").Value = CType(ub.rebar_grade, Double) Else .Worksheets("Input").Range("Fy").ClearContents
'                If Not IsNothing(ub.top_and_bottom_rebar_different) Then .Worksheets("Input").Range("DifferentReinforcementBoolean").Value = ub.top_and_bottom_rebar_different
'                If Not IsNothing(ub.block_foundation) Then .Worksheets("Input").Range("BlockFoundationBoolean").Value = ub.block_foundation
'                If Not IsNothing(ub.rectangular_foundation) Then .Worksheets("Input").Range("RectangularPadBoolean").Value = ub.rectangular_foundation
'                If Not IsNothing(ub.base_plate_distance_above_foundation) Then .Worksheets("Input").Range("bpdist").Value = CType(ub.base_plate_distance_above_foundation, Double) Else .Worksheets("Input").Range("bpdist").ClearContents
'                If Not IsNothing(ub.bolt_circle_bearing_plate_width) Then .Worksheets("Input").Range("BC").Value = CType(ub.bolt_circle_bearing_plate_width, Double) Else .Worksheets("Input").Range("BC").ClearContents
'                If Not IsNothing(ub.tower_centroid_offset) Then .Worksheets("Input").Range("TowerCentroidOffsetBoolean").Value = ub.tower_centroid_offset
'                If Not IsNothing(ub.pier_shape) Then .Worksheets("Input").Range("shape").Value = ub.pier_shape
'                If Not IsNothing(ub.pier_diameter) Then .Worksheets("Input").Range("dpier").Value = CType(ub.pier_diameter, Double) Else .Worksheets("Input").Range("dpier").ClearContents
'                If Not IsNothing(ub.pier_rebar_quantity) Then .Worksheets("Input").Range("mc").Value = CType(ub.pier_rebar_quantity, Integer) Else .Worksheets("Input").Range("mc").ClearContents
'                If Not IsNothing(ub.pier_rebar_size) Then .Worksheets("Input").Range("Sc").Value = CType(ub.pier_rebar_size, Integer) Else .Worksheets("Input").Range("Sc").ClearContents
'                If Not IsNothing(ub.pier_tie_quantity) Then .Worksheets("Input").Range("mt").Value = CType(ub.pier_tie_quantity, Integer) Else .Worksheets("Input").Range("mt").ClearContents
'                If Not IsNothing(ub.pier_tie_size) Then .Worksheets("Input").Range("St").Value = CType(ub.pier_tie_size, Integer) Else .Worksheets("Input").Range("St").ClearContents
'                If Not IsNothing(ub.pier_reinforcement_type) Then .Worksheets("Input").Range("PierReinfType").Value = ub.pier_reinforcement_type
'                If Not IsNothing(ub.pier_clear_cover) Then .Worksheets("Input").Range("ccpier").Value = CType(ub.pier_clear_cover, Double) Else .Worksheets("Input").Range("ccpier").ClearContents
'                If Not IsNothing(ub.pad_width_1) Then .Worksheets("Input").Range("W").Value = CType(ub.pad_width_1, Double) Else .Worksheets("Input").Range("W").ClearContents
'                If Not IsNothing(ub.pad_width_2) Then .Worksheets("Input").Range("W.dir2").Value = CType(ub.pad_width_2, Double) Else .Worksheets("Input").Range("W.dir2").ClearContents
'                If Not IsNothing(ub.pad_thickness) Then .Worksheets("Input").Range("T").Value = CType(ub.pad_thickness, Double) Else .Worksheets("Input").Range("T").ClearContents
'                If Not IsNothing(ub.pad_rebar_size_top_dir1) Then .Worksheets("Input").Range("sptop").Value = CType(ub.pad_rebar_size_top_dir1, Integer) Else .Worksheets("Input").Range("sptop").ClearContents
'                If Not IsNothing(ub.pad_rebar_size_bottom_dir1) Then .Worksheets("Input").Range("Sp").Value = CType(ub.pad_rebar_size_bottom_dir1, Integer) Else .Worksheets("Input").Range("Sp").ClearContents
'                If Not IsNothing(ub.pad_rebar_size_top_dir2) Then .Worksheets("Input").Range("sptop2").Value = CType(ub.pad_rebar_size_top_dir2, Integer) Else .Worksheets("Input").Range("sptop2").ClearContents
'                If Not IsNothing(ub.pad_rebar_size_bottom_dir2) Then .Worksheets("Input").Range("sp_2").Value = CType(ub.pad_rebar_size_bottom_dir2, Integer) Else .Worksheets("Input").Range("sp_2").ClearContents
'                If Not IsNothing(ub.pad_rebar_quantity_top_dir1) Then .Worksheets("Input").Range("mptop").Value = CType(ub.pad_rebar_quantity_top_dir1, Integer) Else .Worksheets("Input").Range("mptop").ClearContents
'                If Not IsNothing(ub.pad_rebar_quantity_bottom_dir1) Then .Worksheets("Input").Range("mp").Value = CType(ub.pad_rebar_quantity_bottom_dir1, Integer) Else .Worksheets("Input").Range("mp").ClearContents
'                If Not IsNothing(ub.pad_rebar_quantity_top_dir2) Then .Worksheets("Input").Range("mptop2").Value = CType(ub.pad_rebar_quantity_top_dir2, Integer) Else .Worksheets("Input").Range("mptop2").ClearContents
'                If Not IsNothing(ub.pad_rebar_quantity_bottom_dir2) Then .Worksheets("Input").Range("mp_2").Value = CType(ub.pad_rebar_quantity_bottom_dir2, Integer) Else .Worksheets("Input").Range("mp_2").ClearContents
'                If Not IsNothing(ub.pad_clear_cover) Then .Worksheets("Input").Range("ccpad").Value = CType(ub.pad_clear_cover, Double) Else .Worksheets("Input").Range("ccpad").ClearContents
'                If Not IsNothing(ub.total_soil_unit_weight) Then .Worksheets("Input").Range("γ").Value = CType(ub.total_soil_unit_weight, Double) Else .Worksheets("Input").Range("γ").ClearContents
'                If Not IsNothing(ub.bearing_type) Then .Worksheets("Input").Range("BearingType").Value = ub.bearing_type
'                If Not IsNothing(ub.nominal_bearing_capacity) Then .Worksheets("Input").Range("Qinput").Value = CType(ub.nominal_bearing_capacity, Double) Else .Worksheets("Input").Range("Qinput").ClearContents
'                If Not IsNothing(ub.cohesion) Then .Worksheets("Input").Range("Cu").Value = CType(ub.cohesion, Double) Else .Worksheets("Input").Range("Cu").ClearContents
'                If Not IsNothing(ub.friction_angle) Then .Worksheets("Input").Range("ϕ").Value = CType(ub.friction_angle, Double) Else .Worksheets("Input").Range("ϕ").ClearContents
'                If Not IsNothing(ub.spt_blow_count) Then .Worksheets("Input").Range("N_blows").Value = CType(ub.spt_blow_count, Integer) Else .Worksheets("Input").Range("N_blows").ClearContents
'                If Not IsNothing(ub.base_friction_factor) Then .Worksheets("Input").Range("μ").Value = CType(ub.base_friction_factor, Double) Else .Worksheets("Input").Range("μ").ClearContents
'                If Not IsNothing(ub.neglect_depth) Then .Worksheets("Input").Range("N").Value = CType(ub.neglect_depth, Double)
'                If ub.bearing_distribution_type = False Then .Worksheets("Input").Range("Rock").Value = "Yes" Else .Worksheets("Input").Range("Rock").Value = "No"
'                If ub.groundwater_depth = -1 Then .Worksheets("Input").Range("gw").Value = "N/A" Else .Worksheets("Input").Range("gw").Value = CType(ub.groundwater_depth, Double) 'If -1 then set to N/A
'                If Not IsNothing(ub.basic_soil_check) Then .Worksheets("Input").Range("SoilInteractionBoolean").Value = ub.basic_soil_check
'                If Not IsNothing(ub.structural_check) Then .Worksheets("Input").Range("StructuralCheckBoolean").Value = ub.structural_check
'                'If Not IsNothing(ub.tool_version) Then .Worksheets("Revision").Range("vnum").Value = ub.tool_version


'                'Seismic design category
'                'TIA
'                'BU
'                'App
'                'Site name
'                'tower height
'                'base face width
'                'reactions (From tnx)
'#Region "Alterate method of saving to excel"
'                ''''.Worksheets("Details (SAPI)").Range("A" & ubRow).Value = ub.unit_base_id
'                ''''.Worksheets("Details (SAPI)").Range("B" & ubRow).Value = ub.extension_above_grade
'                ''''.Worksheets("Details (SAPI)").Range("C" & ubRow).Value = ub.foundation_depth
'                ''''.Worksheets("Details (SAPI)").Range("D" & ubRow).Value = ub.concrete_compressive_strength
'                ''''.Worksheets("Details (SAPI)").Range("E" & ubRow).Value = ub.dry_concrete_density
'                ''''.Worksheets("Details (SAPI)").Range("F" & ubRow).Value = ub.rebar_grade
'                ''''.Worksheets("Details (SAPI)").Range("G" & ubRow).Value = ub.top_and_bottom_rebar_different
'                ''''.Worksheets("Details (SAPI)").Range("H" & ubRow).Value = ub.block_foundation
'                ''''.Worksheets("Details (SAPI)").Range("I" & ubRow).Value = ub.rectangular_foundation
'                ''''.Worksheets("Details (SAPI)").Range("J" & ubRow).Value = ub.base_plate_distance_above_foundation
'                ''''.Worksheets("Details (SAPI)").Range("K" & ubRow).Value = ub.bolt_circle_bearing_plate_width
'                ''''.Worksheets("Details (SAPI)").Range("L" & ubRow).Value = ub.tower_centroid_offset
'                ''''.Worksheets("Details (SAPI)").Range("M" & ubRow).Value = ub.pier_shape
'                ''''.Worksheets("Details (SAPI)").Range("N" & ubRow).Value = ub.pier_diameter
'                ''''.Worksheets("Details (SAPI)").Range("O" & ubRow).Value = ub.pier_rebar_quantity
'                ''''.Worksheets("Details (SAPI)").Range("P" & ubRow).Value = ub.pier_rebar_size
'                ''''.Worksheets("Details (SAPI)").Range("Q" & ubRow).Value = ub.pier_tie_quantity
'                ''''.Worksheets("Details (SAPI)").Range("R" & ubRow).Value = ub.pier_tie_size
'                ''''.Worksheets("Details (SAPI)").Range("S" & ubRow).Value = ub.pier_reinforcement_type
'                ''''.Worksheets("Details (SAPI)").Range("T" & ubRow).Value = ub.pier_clear_cover
'                ''''.Worksheets("Details (SAPI)").Range("U" & ubRow).Value = ub.pad_width_1
'                ''''.Worksheets("Details (SAPI)").Range("V" & ubRow).Value = ub.pad_width_2
'                ''''.Worksheets("Details (SAPI)").Range("W" & ubRow).Value = ub.pad_thickness
'                ''''.Worksheets("Details (SAPI)").Range("X" & ubRow).Value = ub.pad_rebar_size_top_dir1
'                ''''.Worksheets("Details (SAPI)").Range("Y" & ubRow).Value = ub.pad_rebar_size_bottom_dir1
'                ''''.Worksheets("Details (SAPI)").Range("Z" & ubRow).Value = ub.pad_rebar_size_top_dir2
'                ''''.Worksheets("Details (SAPI)").Range("AA" & ubRow).Value = ub.pad_rebar_size_bottom_dir2
'                ''''.Worksheets("Details (SAPI)").Range("AB" & ubRow).Value = ub.pad_rebar_quantity_top_dir1
'                ''''.Worksheets("Details (SAPI)").Range("AC" & ubRow).Value = ub.pad_rebar_quantity_bottom_dir1
'                ''''.Worksheets("Details (SAPI)").Range("AD" & ubRow).Value = ub.pad_rebar_quantity_top_dir2
'                ''''.Worksheets("Details (SAPI)").Range("AE" & ubRow).Value = ub.pad_rebar_quantity_bottom_dir2
'                ''''.Worksheets("Details (SAPI)").Range("AF" & ubRow).Value = ub.pad_clear_cover
'                ''''.Worksheets("Details (SAPI)").Range("AG" & ubRow).Value = ub.total_soil_unit_weight
'                ''''.Worksheets("Details (SAPI)").Range("AH" & ubRow).Value = ub.bearing_type
'                ''''.Worksheets("Details (SAPI)").Range("AI" & ubRow).Value = ub.nominal_bearing_capacity
'                ''''.Worksheets("Details (SAPI)").Range("AJ" & ubRow).Value = ub.cohesion
'                ''''.Worksheets("Details (SAPI)").Range("AK" & ubRow).Value = ub.friction_angle
'                ''''.Worksheets("Details (SAPI)").Range("AL" & ubRow).Value = ub.spt_blow_count
'                ''''.Worksheets("Details (SAPI)").Range("AM" & ubRow).Value = ub.base_friction_factor
'                ''''.Worksheets("Details (SAPI)").Range("AN" & ubRow).Value = ub.neglect_depth
'                ''''.Worksheets("Details (SAPI)").Range("AO" & ubRow).Value = ub.bearing_distribution_type
'                ''''.Worksheets("Details (SAPI)").Range("AP" & ubRow).Value = ub.groundwater_depth
'                ''''ubRow += 1
'#End Region

'            End With
'            SaveAndCloseUnitBase()
'        Next
'    End Sub

'    Private Sub LoadNewUnitBase()
'        NewUnitBaseWb.LoadDocument(UnitBaseTemplatePath, UnitBaseFileType)
'        NewUnitBaseWb.BeginUpdate()
'    End Sub

'    Private Sub SaveAndCloseUnitBase()
'        NewUnitBaseWb.EndUpdate()
'        NewUnitBaseWb.SaveDocument(ExcelFilePath, UnitBaseFileType)
'    End Sub
'#End Region

'#Region "SQL Insert Statements"
'    Private Function InsertUnitBaseDetail(ByVal ub As SST_Unit_Base) As String
'        Dim insertString As String = ""

'        'insertString += "@FndID"
'        insertString += "" & IIf(IsNothing(ub.pier_shape), "Null", "'" & ub.pier_shape.ToString & "'")
'        insertString += "," & IIf(IsNothing(ub.pier_diameter), "Null", ub.pier_diameter.ToString)
'        insertString += "," & IIf(IsNothing(ub.extension_above_grade), "Null", ub.extension_above_grade.ToString)
'        insertString += "," & IIf(IsNothing(ub.pier_rebar_size), "Null", ub.pier_rebar_size.ToString)
'        insertString += "," & IIf(IsNothing(ub.pier_tie_size), "Null", ub.pier_tie_size.ToString)
'        insertString += "," & IIf(IsNothing(ub.pier_tie_quantity), "Null", ub.pier_tie_quantity.ToString)
'        insertString += "," & IIf(IsNothing(ub.pier_reinforcement_type), "Null", "'" & ub.pier_reinforcement_type.ToString & "'")
'        insertString += "," & IIf(IsNothing(ub.pier_clear_cover), "Null", ub.pier_clear_cover.ToString)
'        insertString += "," & IIf(IsNothing(ub.foundation_depth), "Null", ub.foundation_depth.ToString)
'        insertString += "," & IIf(IsNothing(ub.pad_width_1), "Null", ub.pad_width_1.ToString)
'        insertString += "," & IIf(IsNothing(ub.pad_width_2), "Null", ub.pad_width_2.ToString)
'        insertString += "," & IIf(IsNothing(ub.pad_thickness), "Null", ub.pad_thickness.ToString)
'        insertString += "," & IIf(IsNothing(ub.pad_rebar_size_top_dir1), "Null", ub.pad_rebar_size_top_dir1.ToString)
'        insertString += "," & IIf(IsNothing(ub.pad_rebar_size_bottom_dir1), "Null", ub.pad_rebar_size_bottom_dir1.ToString)
'        insertString += "," & IIf(IsNothing(ub.pad_rebar_size_top_dir2), "Null", ub.pad_rebar_size_top_dir2.ToString)
'        insertString += "," & IIf(IsNothing(ub.pad_rebar_size_bottom_dir2), "Null", ub.pad_rebar_size_bottom_dir2.ToString)
'        insertString += "," & IIf(IsNothing(ub.pad_rebar_quantity_top_dir1), "Null", ub.pad_rebar_quantity_top_dir1.ToString)
'        insertString += "," & IIf(IsNothing(ub.pad_rebar_quantity_bottom_dir1), "Null", ub.pad_rebar_quantity_bottom_dir1.ToString)
'        insertString += "," & IIf(IsNothing(ub.pad_rebar_quantity_top_dir2), "Null", ub.pad_rebar_quantity_top_dir2.ToString)
'        insertString += "," & IIf(IsNothing(ub.pad_rebar_quantity_bottom_dir2), "Null", ub.pad_rebar_quantity_bottom_dir2.ToString)
'        insertString += "," & IIf(IsNothing(ub.pad_clear_cover), "Null", ub.pad_clear_cover.ToString)
'        insertString += "," & IIf(IsNothing(ub.rebar_grade), "Null", ub.rebar_grade.ToString)
'        insertString += "," & IIf(IsNothing(ub.concrete_compressive_strength), "Null", ub.concrete_compressive_strength.ToString)
'        insertString += "," & IIf(IsNothing(ub.dry_concrete_density), "Null", ub.dry_concrete_density.ToString)
'        insertString += "," & IIf(IsNothing(ub.total_soil_unit_weight), "Null", ub.total_soil_unit_weight.ToString)
'        insertString += "," & IIf(IsNothing(ub.bearing_type), "Null", "'" & ub.bearing_type.ToString & "'")
'        insertString += "," & IIf(IsNothing(ub.nominal_bearing_capacity), "Null", ub.nominal_bearing_capacity.ToString)
'        insertString += "," & IIf(IsNothing(ub.cohesion), "Null", ub.cohesion.ToString)
'        insertString += "," & IIf(IsNothing(ub.friction_angle), "Null", ub.friction_angle.ToString)
'        insertString += "," & IIf(IsNothing(ub.spt_blow_count), "Null", ub.spt_blow_count.ToString)
'        insertString += "," & IIf(IsNothing(ub.base_friction_factor), "Null", ub.base_friction_factor.ToString)
'        insertString += "," & IIf(IsNothing(ub.neglect_depth), "Null", ub.neglect_depth.ToString)
'        insertString += "," & IIf(IsNothing(ub.bearing_distribution_type), "Null", "'" & ub.bearing_distribution_type.ToString & "'")
'        insertString += "," & IIf(IsNothing(ub.groundwater_depth), "Null", ub.groundwater_depth.ToString) '*******
'        insertString += "," & IIf(IsNothing(ub.top_and_bottom_rebar_different), "Null", "'" & ub.top_and_bottom_rebar_different.ToString & "'")
'        insertString += "," & IIf(IsNothing(ub.block_foundation), "Null", "'" & ub.block_foundation.ToString & "'")
'        insertString += "," & IIf(IsNothing(ub.rectangular_foundation), "Null", "'" & ub.rectangular_foundation.ToString & "'")
'        insertString += "," & IIf(IsNothing(ub.base_plate_distance_above_foundation), "Null", ub.base_plate_distance_above_foundation.ToString)
'        insertString += "," & IIf(IsNothing(ub.bolt_circle_bearing_plate_width), "Null", ub.bolt_circle_bearing_plate_width.ToString)
'        insertString += "," & IIf(IsNothing(ub.tower_centroid_offset), "Null", "'" & ub.tower_centroid_offset.ToString & "'")
'        insertString += "," & IIf(IsNothing(ub.pier_rebar_quantity), "Null", ub.pier_rebar_quantity.ToString)
'        insertString += "," & IIf(IsNothing(ub.basic_soil_check), "Null", "'" & ub.basic_soil_check.ToString & "'")
'        insertString += "," & IIf(IsNothing(ub.structural_check), "Null", "'" & ub.structural_check.ToString & "'")
'        insertString += "," & IIf(IsNothing(ub.tool_version), "Null", "'" & ub.tool_version.ToString & "'")

'        Return insertString
'    End Function
'#End Region

'#Region "SQL Update Statements"
'    Private Function UpdateUnitBaseDetail(ByVal ub As SST_Unit_Base) As String
'        Dim updateString As String = ""

'        updateString += "UPDATE unit_base_details SET "
'        updateString += "extension_above_grade=" & IIf(IsNothing(ub.extension_above_grade), "Null", ub.extension_above_grade.ToString)
'        updateString += ", foundation_depth=" & IIf(IsNothing(ub.foundation_depth), "Null", ub.foundation_depth.ToString)
'        updateString += ", concrete_compressive_strength=" & IIf(IsNothing(ub.concrete_compressive_strength), "Null", ub.concrete_compressive_strength.ToString)
'        updateString += ", dry_concrete_density=" & IIf(IsNothing(ub.dry_concrete_density), "Null", ub.dry_concrete_density.ToString)
'        updateString += ", rebar_grade=" & IIf(IsNothing(ub.rebar_grade), "Null", ub.rebar_grade.ToString)
'        updateString += ", top_and_bottom_rebar_different=" & IIf(IsNothing(ub.top_and_bottom_rebar_different), "Null", "'" & ub.top_and_bottom_rebar_different.ToString & "'")
'        updateString += ", block_foundation=" & IIf(IsNothing(ub.block_foundation), "Null", "'" & ub.block_foundation.ToString & "'")
'        updateString += ", rectangular_foundation=" & IIf(IsNothing(ub.rectangular_foundation), "Null", "'" & ub.rectangular_foundation.ToString & "'")
'        updateString += ", base_plate_distance_above_foundation=" & IIf(IsNothing(ub.base_plate_distance_above_foundation), "Null", ub.base_plate_distance_above_foundation.ToString)
'        updateString += ", bolt_circle_bearing_plate_width=" & IIf(IsNothing(ub.bolt_circle_bearing_plate_width), "Null", ub.bolt_circle_bearing_plate_width.ToString)
'        updateString += ", tower_centroid_offset=" & IIf(IsNothing(ub.tower_centroid_offset), "Null", "'" & ub.tower_centroid_offset.ToString & "'")
'        updateString += ", pier_shape=" & IIf(IsNothing(ub.pier_shape), "Null", "'" & ub.pier_shape.ToString & "'")
'        updateString += ", pier_diameter=" & IIf(IsNothing(ub.pier_diameter), "Null", ub.pier_diameter.ToString)
'        updateString += ", pier_rebar_quantity=" & IIf(IsNothing(ub.pier_rebar_quantity), "Null", ub.pier_rebar_quantity.ToString)
'        updateString += ", pier_rebar_size=" & IIf(IsNothing(ub.pier_rebar_size), "Null", ub.pier_rebar_size.ToString)
'        updateString += ", pier_tie_quantity=" & IIf(IsNothing(ub.pier_tie_quantity), "Null", ub.pier_tie_quantity.ToString)
'        updateString += ", pier_tie_size=" & IIf(IsNothing(ub.pier_tie_size), "Null", ub.pier_tie_size.ToString)
'        updateString += ", pier_reinforcement_type=" & IIf(IsNothing(ub.pier_reinforcement_type), "Null", "'" & ub.pier_reinforcement_type.ToString & "'")
'        updateString += ", pier_clear_cover=" & IIf(IsNothing(ub.pier_clear_cover), "Null", ub.pier_clear_cover.ToString)
'        updateString += ", pad_width_2=" & IIf(IsNothing(ub.pad_width_2), "Null", ub.pad_width_2.ToString)
'        updateString += ", pad_thickness=" & IIf(IsNothing(ub.pad_thickness), "Null", ub.pad_thickness.ToString)
'        updateString += ", pad_rebar_size_top_dir1=" & IIf(IsNothing(ub.pad_rebar_size_top_dir1), "Null", ub.pad_rebar_size_top_dir1.ToString)
'        updateString += ", pad_rebar_size_bottom_dir1=" & IIf(IsNothing(ub.pad_rebar_size_bottom_dir1), "Null", ub.pad_rebar_size_bottom_dir1.ToString)
'        updateString += ", pad_rebar_size_top_dir2=" & IIf(IsNothing(ub.pad_rebar_size_top_dir2), "Null", ub.pad_rebar_size_top_dir2.ToString)
'        updateString += ", pad_rebar_size_bottom_dir2=" & IIf(IsNothing(ub.pad_rebar_size_bottom_dir2), "Null", ub.pad_rebar_size_bottom_dir2.ToString)
'        updateString += ", pad_rebar_quantity_top_dir1=" & IIf(IsNothing(ub.pad_rebar_quantity_top_dir1), "Null", ub.pad_rebar_quantity_top_dir1.ToString)
'        updateString += ", pad_rebar_quantity_bottom_dir1=" & IIf(IsNothing(ub.pad_rebar_quantity_bottom_dir1), "Null", ub.pad_rebar_quantity_bottom_dir1.ToString)
'        updateString += ", pad_rebar_quantity_top_dir2=" & IIf(IsNothing(ub.pad_rebar_quantity_top_dir2), "Null", ub.pad_rebar_quantity_top_dir2.ToString)
'        updateString += ", pad_rebar_quantity_bottom_dir2=" & IIf(IsNothing(ub.pad_rebar_quantity_bottom_dir2), "Null", ub.pad_rebar_quantity_bottom_dir2.ToString)
'        updateString += ", pad_clear_cover=" & IIf(IsNothing(ub.pad_clear_cover), "Null", ub.pad_clear_cover.ToString)
'        updateString += ", total_soil_unit_weight=" & IIf(IsNothing(ub.total_soil_unit_weight), "Null", ub.total_soil_unit_weight.ToString)
'        updateString += ", pad_width_1=" & IIf(IsNothing(ub.pad_width_1), "Null", ub.pad_width_1.ToString)
'        updateString += ", bearing_type=" & IIf(IsNothing(ub.bearing_type), "Null", "'" & ub.bearing_type.ToString & "'")
'        updateString += ", nominal_bearing_capacity=" & IIf(IsNothing(ub.nominal_bearing_capacity), "Null", ub.nominal_bearing_capacity.ToString)
'        updateString += ", cohesion=" & IIf(IsNothing(ub.cohesion), "Null", ub.cohesion.ToString)
'        updateString += ", friction_angle=" & IIf(IsNothing(ub.friction_angle), "Null", ub.friction_angle.ToString)
'        updateString += ", spt_blow_count=" & IIf(IsNothing(ub.spt_blow_count), "Null", ub.spt_blow_count.ToString)
'        updateString += ", base_friction_factor=" & IIf(IsNothing(ub.base_friction_factor), "Null", ub.base_friction_factor.ToString)
'        updateString += ", neglect_depth=" & IIf(IsNothing(ub.neglect_depth), "Null", ub.neglect_depth.ToString)
'        updateString += ", bearing_distribution_type=" & IIf(IsNothing(ub.bearing_distribution_type), "Null", "'" & ub.bearing_distribution_type.ToString & "'")
'        updateString += ", groundwater_depth=" & IIf(IsNothing(ub.groundwater_depth), "Null", ub.groundwater_depth.ToString)
'        updateString += ", basic_soil_check=" & IIf(IsNothing(ub.basic_soil_check), "Null", "'" & ub.basic_soil_check.ToString & "'")
'        updateString += ", structural_check=" & IIf(IsNothing(ub.structural_check), "Null", "'" & ub.structural_check.ToString & "'")
'        updateString += ", tool_version=" & IIf(IsNothing(ub.tool_version), "Null", "'" & ub.tool_version.ToString & "'")
'        updateString += " WHERE ID=" & ub.unit_base_id & vbNewLine

'        Return updateString
'    End Function
'#End Region

'#Region "General"
'    Public Sub Clear()
'        ExcelFilePath = ""
'        UnitBases.Clear()
'    End Sub

'    Private Function UnitBaseSQLDataTables() As List(Of SQLParameter)
'        Dim MyParameters As New List(Of SQLParameter)

'        MyParameters.Add(New SQLParameter("Unit Base General Details SQL", "Unit Base (SELECT Details).sql"))

'        Return MyParameters
'    End Function

'    Private Function UnitBaseExcelDTParameters() As List(Of EXCELDTParameter)
'        Dim MyParameters As New List(Of EXCELDTParameter)

'        MyParameters.Add(New EXCELDTParameter("Unit Base General Details EXCEL", "A2:AP3", "Details (SAPI)"))
'        'MyParameters.Add(New EXCELDTParameter("Pile Location EXCEL", "S3:U103", "SAPI"))

'        Return MyParameters
'    End Function

'    'Alternate Excel DataLink Option:
'    'Private Function UnitBaseExcelRngParameters() As List(Of EXCELRngParameter)
'    '    Dim MyParameters As New List(Of EXCELRngParameter)

'    '    MyParameters.Add(New EXCELRngParameter("ID", "unit_base_id"))
'    '    MyParameters.Add(New EXCELRngParameter("E", "extension_above_grade"))
'    '    MyParameters.Add(New EXCELRngParameter("D", "foundation_depth"))
'    '    MyParameters.Add(New EXCELRngParameter("F\c", "concrete_compressive_strength"))
'    '    MyParameters.Add(New EXCELRngParameter("ConcreteDensity", "dry_concrete_density"))
'    '    MyParameters.Add(New EXCELRngParameter("Fy", "rebar_grade"))
'    '    MyParameters.Add(New EXCELRngParameter("DifferentReinforcementBoolean", "top_and_bottom_rebar_different"))
'    '    MyParameters.Add(New EXCELRngParameter("BlockFoundationBoolean", "block_foundation"))
'    '    MyParameters.Add(New EXCELRngParameter("RectangularPadBoolean", "rectangular_foundation"))
'    '    MyParameters.Add(New EXCELRngParameter("bpdist", "base_plate_distance_above_foundation"))
'    '    MyParameters.Add(New EXCELRngParameter("BC", "bolt_circle_bearing_plate_width"))
'    '    MyParameters.Add(New EXCELRngParameter("TowerCentroidOffsetBoolean", "tower_centroid_offset"))
'    '    MyParameters.Add(New EXCELRngParameter("shape", "pier_shape"))
'    '    MyParameters.Add(New EXCELRngParameter("dpier", "pier_diameter"))
'    '    MyParameters.Add(New EXCELRngParameter("mc", "pier_rebar_quantity"))
'    '    MyParameters.Add(New EXCELRngParameter("Sc", "pier_rebar_size"))
'    '    MyParameters.Add(New EXCELRngParameter("mt", "pier_tie_quantity"))
'    '    MyParameters.Add(New EXCELRngParameter("St", "pier_tie_size"))
'    '    MyParameters.Add(New EXCELRngParameter("PierReinfType", "pier_reinforcement_type"))
'    '    MyParameters.Add(New EXCELRngParameter("ccpier", "pier_clear_cover"))
'    '    MyParameters.Add(New EXCELRngParameter("W", "pad_width_1"))
'    '    MyParameters.Add(New EXCELRngParameter("W.dir2", "pad_width_2"))
'    '    MyParameters.Add(New EXCELRngParameter("T", "pad_thickness"))
'    '    MyParameters.Add(New EXCELRngParameter("sptop", "pad_rebar_size_top_dir1"))
'    '    MyParameters.Add(New EXCELRngParameter("Sp", "pad_rebar_size_bottom_dir1"))
'    '    MyParameters.Add(New EXCELRngParameter("sptop2", "pad_rebar_size_top_dir2"))
'    '    MyParameters.Add(New EXCELRngParameter("sp_2", "pad_rebar_size_bottom_dir2"))
'    '    MyParameters.Add(New EXCELRngParameter("mptop", "pad_rebar_quantity_top_dir1"))
'    '    MyParameters.Add(New EXCELRngParameter("mp", "pad_rebar_quantity_bottom_dir1"))
'    '    MyParameters.Add(New EXCELRngParameter("mptop2", "pad_rebar_quantity_top_dir2"))
'    '    MyParameters.Add(New EXCELRngParameter("mp_2", "pad_rebar_quantity_bottom_dir2"))
'    '    MyParameters.Add(New EXCELRngParameter("ccpad", "pad_clear_cover"))
'    '    MyParameters.Add(New EXCELRngParameter("γ", "total_soil_unit_weight"))
'    '    MyParameters.Add(New EXCELRngParameter("BearingType", "bearing_type"))
'    '    MyParameters.Add(New EXCELRngParameter("Qinput", "nominal_bearing_capacity"))
'    '    MyParameters.Add(New EXCELRngParameter("Cu", "cohesion"))
'    '    MyParameters.Add(New EXCELRngParameter("ϕ", "friction_angle"))
'    '    MyParameters.Add(New EXCELRngParameter("N_blows", "spt_blow_count"))
'    '    MyParameters.Add(New EXCELRngParameter("μ", "base_friction_factor"))
'    '    MyParameters.Add(New EXCELRngParameter("N", "neglect_depth"))
'    '    MyParameters.Add(New EXCELRngParameter("Rock", "bearing_distribution_type"))
'    '    MyParameters.Add(New EXCELRngParameter("gw", "groundwater_depth"))

'    '    Return MyParameters
'    'End Function

'#End Region


'#Region "Check Changes"
'    Private changeDt As New DataTable
'    Private changeList As New List(Of AnalysisChanges)
'    Function CheckChanges(ByVal xlUnitBase As SST_Unit_Base, ByVal sqlUnitBase As SST_Unit_Base) As Boolean
'        Dim changesMade As Boolean = False

'        'changeDt.Columns.Add("Variable", Type.GetType("System.String"))
'        'changeDt.Columns.Add("New Value", Type.GetType("System.String"))
'        'changeDt.Columns.Add("Previuos Value", Type.GetType("System.String"))
'        'changeDt.Columns.Add("WO", Type.GetType("System.String"))

'        'Check Details
'        If Check1Change(xlUnitBase.pier_shape, sqlUnitBase.pier_shape, 1, "Pier_Shape") Then changesMade = True
'        If Check1Change(xlUnitBase.pier_diameter, sqlUnitBase.pier_diameter, 1, "Pier_Diameter") Then changesMade = True
'        If Check1Change(xlUnitBase.extension_above_grade, sqlUnitBase.extension_above_grade, 1, "Extension_Above_Grade") Then changesMade = True
'        If Check1Change(xlUnitBase.pier_rebar_size, sqlUnitBase.pier_rebar_size, 1, "Pier_Rebar_Size") Then changesMade = True
'        If Check1Change(xlUnitBase.pier_tie_size, sqlUnitBase.pier_tie_size, 1, "Pier_Tie_Size") Then changesMade = True
'        If Check1Change(xlUnitBase.pier_tie_quantity, sqlUnitBase.pier_tie_quantity, 1, "Pier_Tie_Quantity") Then changesMade = True
'        If Check1Change(xlUnitBase.pier_reinforcement_type, sqlUnitBase.pier_reinforcement_type, 1, "Pier_Reinforcement_Type") Then changesMade = True
'        If Check1Change(xlUnitBase.pier_clear_cover, sqlUnitBase.pier_clear_cover, 1, "Pier_Clear_Cover") Then changesMade = True
'        If Check1Change(xlUnitBase.foundation_depth, sqlUnitBase.foundation_depth, 1, "Foundation_Depth") Then changesMade = True
'        If Check1Change(xlUnitBase.pad_width_1, sqlUnitBase.pad_width_1, 1, "Pad_Width_1") Then changesMade = True
'        If Check1Change(xlUnitBase.pad_width_2, sqlUnitBase.pad_width_2, 1, "Pad_Width_2") Then changesMade = True
'        If Check1Change(xlUnitBase.pad_thickness, sqlUnitBase.pad_thickness, 1, "Pad_Thickness") Then changesMade = True
'        If Check1Change(xlUnitBase.pad_rebar_size_top_dir1, sqlUnitBase.pad_rebar_size_top_dir1, 1, "Pad_Rebar_Size_Top_Dir1") Then changesMade = True
'        If Check1Change(xlUnitBase.pad_rebar_size_bottom_dir1, sqlUnitBase.pad_rebar_size_bottom_dir1, 1, "Pad_Rebar_Size_Bottom_Dir1") Then changesMade = True
'        If Check1Change(xlUnitBase.pad_rebar_size_top_dir2, sqlUnitBase.pad_rebar_size_top_dir2, 1, "Pad_Rebar_Size_Top_Dir2") Then changesMade = True
'        If Check1Change(xlUnitBase.pad_rebar_size_bottom_dir2, sqlUnitBase.pad_rebar_size_bottom_dir2, 1, "Pad_Rebar_Size_Bottom_Dir2") Then changesMade = True
'        If Check1Change(xlUnitBase.pad_rebar_quantity_top_dir1, sqlUnitBase.pad_rebar_quantity_top_dir1, 1, "Pad_Rebar_Quantity_Top_Dir1") Then changesMade = True
'        If Check1Change(xlUnitBase.pad_rebar_quantity_bottom_dir1, sqlUnitBase.pad_rebar_quantity_bottom_dir1, 1, "Pad_Rebar_Quantity_Bottom_Dir1") Then changesMade = True
'        If Check1Change(xlUnitBase.pad_rebar_quantity_top_dir2, sqlUnitBase.pad_rebar_quantity_top_dir2, 1, "Pad_Rebar_Quantity_Top_Dir2") Then changesMade = True
'        If Check1Change(xlUnitBase.pad_rebar_quantity_bottom_dir2, sqlUnitBase.pad_rebar_quantity_bottom_dir2, 1, "Pad_Rebar_Quantity_Bottom_Dir2") Then changesMade = True
'        If Check1Change(xlUnitBase.pad_clear_cover, sqlUnitBase.pad_clear_cover, 1, "Pad_Clear_Cover") Then changesMade = True
'        If Check1Change(xlUnitBase.rebar_grade, sqlUnitBase.rebar_grade, 1, "Rebar_Grade") Then changesMade = True
'        If Check1Change(xlUnitBase.concrete_compressive_strength, sqlUnitBase.concrete_compressive_strength, 1, "Concrete_Compressive_Strength") Then changesMade = True
'        If Check1Change(xlUnitBase.dry_concrete_density, sqlUnitBase.dry_concrete_density, 1, "Dry_Concrete_Density") Then changesMade = True
'        If Check1Change(xlUnitBase.total_soil_unit_weight, sqlUnitBase.total_soil_unit_weight, 1, "Total_Soil_Unit_Weight") Then changesMade = True
'        If Check1Change(xlUnitBase.bearing_type, sqlUnitBase.bearing_type, 1, "Bearing_Type") Then changesMade = True
'        If Check1Change(xlUnitBase.nominal_bearing_capacity, sqlUnitBase.nominal_bearing_capacity, 1, "Nominal_Bearing_Capacity") Then changesMade = True
'        If Check1Change(xlUnitBase.cohesion, sqlUnitBase.cohesion, 1, "Cohesion") Then changesMade = True
'        If Check1Change(xlUnitBase.friction_angle, sqlUnitBase.friction_angle, 1, "Friction_Angle") Then changesMade = True
'        If Check1Change(xlUnitBase.spt_blow_count, sqlUnitBase.spt_blow_count, 1, "Spt_Blow_Count") Then changesMade = True
'        If Check1Change(xlUnitBase.base_friction_factor, sqlUnitBase.base_friction_factor, 1, "Base_Friction_Factor") Then changesMade = True
'        If Check1Change(xlUnitBase.neglect_depth, sqlUnitBase.neglect_depth, 1, "Neglect_Depth") Then changesMade = True
'        If Check1Change(xlUnitBase.bearing_distribution_type, sqlUnitBase.bearing_distribution_type, 1, "Bearing_Distribution_Type") Then changesMade = True
'        If Check1Change(xlUnitBase.groundwater_depth, sqlUnitBase.groundwater_depth, 1, "Groundwater_Depth") Then changesMade = True
'        If Check1Change(xlUnitBase.top_and_bottom_rebar_different, sqlUnitBase.top_and_bottom_rebar_different, 1, "Top_And_Bottom_Rebar_Different") Then changesMade = True
'        If Check1Change(xlUnitBase.block_foundation, sqlUnitBase.block_foundation, 1, "Block_Foundation") Then changesMade = True
'        If Check1Change(xlUnitBase.rectangular_foundation, sqlUnitBase.rectangular_foundation, 1, "Rectangular_Foundation") Then changesMade = True
'        If Check1Change(xlUnitBase.base_plate_distance_above_foundation, sqlUnitBase.base_plate_distance_above_foundation, 1, "Base_Plate_Distance_Above_Foundation") Then changesMade = True
'        If Check1Change(xlUnitBase.bolt_circle_bearing_plate_width, sqlUnitBase.bolt_circle_bearing_plate_width, 1, "Bolt_Circle_Bearing_Plate_Width") Then changesMade = True
'        If Check1Change(xlUnitBase.tower_centroid_offset, sqlUnitBase.tower_centroid_offset, 1, "Tower_Centroid_Offset") Then changesMade = True
'        If Check1Change(xlUnitBase.pier_rebar_quantity, sqlUnitBase.pier_rebar_quantity, 1, "Pier_Rebar_Quantity") Then changesMade = True
'        If Check1Change(xlUnitBase.basic_soil_check, sqlUnitBase.basic_soil_check, 1, "Basic_Soil_Check") Then changesMade = True
'        If Check1Change(xlUnitBase.structural_check, sqlUnitBase.structural_check, 1, "Structural_Check") Then changesMade = True
'        If Check1Change(xlUnitBase.tool_version, sqlUnitBase.tool_version, 1, "Tool_Version") Then changesMade = True
'        'If Check1Change(xlUnitBase.tool_version, sqlUnitBase.tool_version, 1, "Tool_Version") Then changesMade = True

'        CreateChangeSummary(changeDt) 'possible alternative to listing change summary
'        Return changesMade
'    End Function

'    Function CreateChangeSummary(ByVal changeDt As DataTable) As String
'        'Sub CreateChangeSummary(ByVal changeDt As DataTable)
'        'Create your string based on data in the datatable
'        Dim summary As String
'        Dim counter As Integer = 0

'        For Each chng As AnalysisChanges In changeList
'            If counter = 0 Then
'                summary += chng.Name & " = " & chng.NewValue & " | Previously: " & chng.PreviousValue
'            Else
'                summary += vbNewLine & chng.Name & " = " & chng.NewValue & " | Previously: " & chng.PreviousValue
'            End If

'            counter += 1
'        Next

'        'write to text file
'        'End Sub
'    End Function

'    Function Check1Change(ByVal newValue As Object, ByVal oldvalue As Object, ByVal tolerance As Double, ByVal variable As String) As Boolean
'        If newValue <> oldvalue Then
'            changeDt.Rows.Add(variable, newValue, oldvalue, CurWO) 'Need to determine what we want to store in this datatable or list (Foundation Type, Foundation ID)?
'            changeList.Add(New AnalysisChanges(oldvalue, newValue, variable, "Unit Base Foundations"))
'            Return True
'        ElseIf Not IsNothing(newValue) And IsNothing(oldvalue) Then 'accounts for when new rows are added. New rows from excel=0 where sql=nothing
'            changeDt.Rows.Add(variable, newValue, oldvalue, CurWO) 'Need to determine what we want to store in this datatable or list (Foundation Type, Foundation ID)?
'            changeList.Add(New AnalysisChanges(oldvalue, newValue, variable, "Unit Base Foundations"))
'            Return True
'        ElseIf IsNothing(newValue) And Not IsNothing(oldvalue) Then 'accounts for when rows are removed. Rows from excel=nothing where sql=value
'            changeDt.Rows.Add(variable, newValue, oldvalue, CurWO) 'Need to determine what we want to store in this datatable or list (Foundation Type, Foundation ID)?
'            changeList.Add(New AnalysisChanges(oldvalue, newValue, variable, "Unit Base Foundations"))
'            Return True
'        End If
'    End Function

'#End Region
'End Class