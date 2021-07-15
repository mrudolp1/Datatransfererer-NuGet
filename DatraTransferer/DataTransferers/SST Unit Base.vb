Option Strict Off

Imports DevExpress.Spreadsheet
Imports CCI_Engineering_Templates
Imports System.Security.Principal

Public Class DataTransfererUnitBase

#Region "Define"
    Private NewUnitBaseWb As New Workbook
    Private prop_ExcelFilePath As String


    Public Property UnitBases As New List(Of SST_Unit_Base)
    Private Property UnitBaseTemplatePath As String = "C:\Users\" & Environment.UserName & "\source\repos\DevExpress Objects\Reference\SST Unit Base Foundation (4.0.4) - MRR.xlsm"
    Private Property UnitBaseFileType As DocumentFormat = DocumentFormat.Xlsm

    Public Property ubDS As DataSet
    Public Property ubDB As String
    Public Property ubID As WindowsIdentity

    Public Property ExcelFilePath() As String
        Get
            Return Me.prop_ExcelFilePath
        End Get
        Set
            Me.prop_ExcelFilePath = Value
        End Set
    End Property


#End Region

#Region "Constructors"
    Public Sub New()
        'Leave method empty
    End Sub

    Public Sub New(ByVal MyDataSet As DataSet, ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String, ByVal BU As String, ByVal Strucutre_ID As String)
        ubDS = MyDataSet
        ubID = LogOnUser
        ubDB = ActiveDatabase
        BUNumber = BU
        STR_ID = Strucutre_ID
    End Sub
#End Region

#Region "Load Data"
    Public Sub LoadFromSQL()
        Dim refid As Integer
        Dim UnitBaseLoader As String

        'Load data to get Unit Base details for the existing structure model
        For Each item As SQLParameter In UnitBaseSQLDataTables()
            UnitBaseLoader = QueryBuilderFromFile(queryPath & "Unit Base\" & item.sqlQuery).Replace("[EXISTING MODEL]", GetExistingModelQuery())
            DoDaSQL.sqlLoader(UnitBaseLoader, item.sqlDatatable, ubDS, ubDB, ubID, "0")
        Next

        'Custom Section to transfer data for the drilled pier tool. Needs to be adjusted for each tool.
        For Each UnitBaseDataRow As DataRow In ubDS.Tables("Unit Base General Details SQL").Rows
            refid = CType(UnitBaseDataRow.Item("unit_base_id"), Integer)

            UnitBases.Add(New SST_Unit_Base(UnitBaseDataRow, refid))
        Next

    End Sub 'Create Unit Base objects based on what is saved in EDS

    Public Sub LoadFromExcel()
        UnitBases.Add(New SST_Unit_Base(ExcelFilePath))
    End Sub 'Create Unit Base objects based on what is coming from the excel file
#End Region

#Region "Save Data"
    Public Sub SaveToEDS()
        For Each ub As SST_Unit_Base In UnitBases
            Dim UnitBaseSaver As String = Common.QueryBuilderFromFile(queryPath & "Unit Base\Unit Base (IN_UP).sql")

            UnitBaseSaver = UnitBaseSaver.Replace("[BU NUMBER]", BUNumber)
            UnitBaseSaver = UnitBaseSaver.Replace("[STRUCTURE ID]", STR_ID)
            UnitBaseSaver = UnitBaseSaver.Replace("[FOUNDATION TYPE]", "Unit Base")
            If ub.unit_base_id = 0 Or IsDBNull(ub.unit_base_id) Then
                UnitBaseSaver = UnitBaseSaver.Replace("'[UNIT BASE ID]'", "NULL")
            Else
                UnitBaseSaver = UnitBaseSaver.Replace("[UNIT BASE ID]", ub.unit_base_id.ToString)
                UnitBaseSaver = UnitBaseSaver.Replace("(SELECT * FROM TEMPORARY)", UpdateUnitBaseDetail(ub))
            End If
            UnitBaseSaver = UnitBaseSaver.Replace("[INSERT ALL UNIT BASE DETAILS]", InsertUnitBaseDetail(ub))

            sqlSender(UnitBaseSaver, ubDB, ubID, "0")
        Next
    End Sub

    Public Sub SaveToExcel()
        Dim ubRow As Integer = 2

        For Each ub As SST_Unit_Base In UnitBases
            LoadNewUnitBase()
            With NewUnitBaseWb
                .Worksheets("Input").Range("ID").Value = ub.unit_base_id
                If Not IsNothing(ub.extension_above_grade) Then .Worksheets("Input").Range("E").Value = ub.extension_above_grade
                If Not IsNothing(ub.foundation_depth) Then .Worksheets("Input").Range("D").Value = ub.foundation_depth
                If Not IsNothing(ub.concrete_compressive_strength) Then .Worksheets("Input").Range("F\c").Value = ub.concrete_compressive_strength
                If Not IsNothing(ub.dry_concrete_density) Then .Worksheets("Input").Range("ConcreteDensity").Value = ub.dry_concrete_density
                If Not IsNothing(ub.rebar_grade) Then .Worksheets("Input").Range("Fy").Value = ub.rebar_grade
                If Not IsNothing(ub.top_and_bottom_rebar_different) Then .Worksheets("Input").Range("DifferentReinforcementBoolean").Value = ub.top_and_bottom_rebar_different
                If Not IsNothing(ub.block_foundation) Then .Worksheets("Input").Range("BlockFoundationBoolean").Value = ub.block_foundation
                If Not IsNothing(ub.rectangular_foundation) Then .Worksheets("Input").Range("RectangularPadBoolean").Value = ub.rectangular_foundation
                If Not IsNothing(ub.base_plate_distance_above_foundation) Then .Worksheets("Input").Range("bpdist").Value = ub.base_plate_distance_above_foundation
                If Not IsNothing(ub.bolt_circle_bearing_plate_width) Then .Worksheets("Input").Range("BC").Value = ub.bolt_circle_bearing_plate_width
                If Not IsNothing(ub.tower_centroid_offset) Then .Worksheets("Input").Range("TowerCentroidOffsetBoolean").Value = ub.tower_centroid_offset
                If Not IsNothing(ub.pier_shape) Then .Worksheets("Input").Range("shape").Value = ub.pier_shape
                If Not IsNothing(ub.pier_diameter) Then .Worksheets("Input").Range("dpier").Value = ub.pier_diameter
                If Not IsNothing(ub.pier_rebar_quantity) Then .Worksheets("Input").Range("mc").Value = ub.pier_rebar_quantity
                If Not IsNothing(ub.pier_rebar_size) Then .Worksheets("Input").Range("Sc").Value = ub.pier_rebar_size
                If Not IsNothing(ub.pier_tie_quantity) Then .Worksheets("Input").Range("mt").Value = ub.pier_tie_quantity
                If Not IsNothing(ub.pier_tie_size) Then .Worksheets("Input").Range("St").Value = ub.pier_tie_size
                If Not IsNothing(ub.pier_reinforcement_type) Then .Worksheets("Input").Range("PierReinfType").Value = ub.pier_reinforcement_type
                If Not IsNothing(ub.pier_clear_cover) Then .Worksheets("Input").Range("ccpier").Value = ub.pier_clear_cover
                If Not IsNothing(ub.pad_width_1) Then .Worksheets("Input").Range("W").Value = ub.pad_width_1
                If Not IsNothing(ub.pad_width_2) Then .Worksheets("Input").Range("W.dir2").Value = ub.pad_width_2
                If Not IsNothing(ub.pad_thickness) Then .Worksheets("Input").Range("T").Value = ub.pad_thickness
                If Not IsNothing(ub.pad_rebar_size_top_dir1) Then .Worksheets("Input").Range("sptop").Value = ub.pad_rebar_size_top_dir1
                If Not IsNothing(ub.pad_rebar_size_bottom_dir1) Then .Worksheets("Input").Range("Sp").Value = ub.pad_rebar_size_bottom_dir1
                If Not IsNothing(ub.pad_rebar_size_top_dir2) Then .Worksheets("Input").Range("sptop2").Value = ub.pad_rebar_size_top_dir2
                If Not IsNothing(ub.pad_rebar_size_bottom_dir2) Then .Worksheets("Input").Range("sp_2").Value = ub.pad_rebar_size_bottom_dir2
                If Not IsNothing(ub.pad_rebar_quantity_top_dir1) Then .Worksheets("Input").Range("mptop").Value = ub.pad_rebar_quantity_top_dir1
                If Not IsNothing(ub.pad_rebar_quantity_bottom_dir1) Then .Worksheets("Input").Range("mp").Value = ub.pad_rebar_quantity_bottom_dir1
                If Not IsNothing(ub.pad_rebar_quantity_top_dir2) Then .Worksheets("Input").Range("mptop2").Value = ub.pad_rebar_quantity_top_dir2
                If Not IsNothing(ub.pad_rebar_quantity_bottom_dir2) Then .Worksheets("Input").Range("mp_2").Value = ub.pad_rebar_quantity_bottom_dir2
                If Not IsNothing(ub.pad_clear_cover) Then .Worksheets("Input").Range("ccpad").Value = ub.pad_clear_cover
                If Not IsNothing(ub.total_soil_unit_weight) Then .Worksheets("Input").Range("γ").Value = ub.total_soil_unit_weight
                If Not IsNothing(ub.bearing_type) Then .Worksheets("Input").Range("BearingType").Value = ub.bearing_type
                If Not IsNothing(ub.nominal_bearing_capacity) Then .Worksheets("Input").Range("Qinput").Value = ub.nominal_bearing_capacity
                If Not IsNothing(ub.cohesion) Then .Worksheets("Input").Range("Cu").Value = ub.cohesion
                If Not IsNothing(ub.friction_angle) Then .Worksheets("Input").Range("ϕ").Value = ub.friction_angle
                If Not IsNothing(ub.spt_blow_count) Then .Worksheets("Input").Range("N_blows").Value = ub.spt_blow_count
                If Not IsNothing(ub.base_friction_factor) Then .Worksheets("Input").Range("μ").Value = ub.base_friction_factor
                If Not IsNothing(ub.neglect_depth) Then .Worksheets("Input").Range("N").Value = ub.neglect_depth
                If ub.bearing_distribution_type = False Then .Worksheets("Input").Range("Rock").Value = "Yes" Else .Worksheets("Input").Range("Rock").Value = "No"
                If ub.groundwater_depth = -1 Then .Worksheets("Input").Range("gw").Value = "N/A" Else .Worksheets("Input").Range("gw").Value = ub.groundwater_depth 'If -1 then set to N/A
                'Seismic design category
                'TIA
                'BU
                'App
                'Site name
                'tower height
                'base face width
                'reactions (From tnx)
#Region "Alterate method of saving to excel"
                ''''.Worksheets("Details (SAPI)").Range("A" & ubRow).Value = ub.unit_base_id
                ''''.Worksheets("Details (SAPI)").Range("B" & ubRow).Value = ub.extension_above_grade
                ''''.Worksheets("Details (SAPI)").Range("C" & ubRow).Value = ub.foundation_depth
                ''''.Worksheets("Details (SAPI)").Range("D" & ubRow).Value = ub.concrete_compressive_strength
                ''''.Worksheets("Details (SAPI)").Range("E" & ubRow).Value = ub.dry_concrete_density
                ''''.Worksheets("Details (SAPI)").Range("F" & ubRow).Value = ub.rebar_grade
                ''''.Worksheets("Details (SAPI)").Range("G" & ubRow).Value = ub.top_and_bottom_rebar_different
                ''''.Worksheets("Details (SAPI)").Range("H" & ubRow).Value = ub.block_foundation
                ''''.Worksheets("Details (SAPI)").Range("I" & ubRow).Value = ub.rectangular_foundation
                ''''.Worksheets("Details (SAPI)").Range("J" & ubRow).Value = ub.base_plate_distance_above_foundation
                ''''.Worksheets("Details (SAPI)").Range("K" & ubRow).Value = ub.bolt_circle_bearing_plate_width
                ''''.Worksheets("Details (SAPI)").Range("L" & ubRow).Value = ub.tower_centroid_offset
                ''''.Worksheets("Details (SAPI)").Range("M" & ubRow).Value = ub.pier_shape
                ''''.Worksheets("Details (SAPI)").Range("N" & ubRow).Value = ub.pier_diameter
                ''''.Worksheets("Details (SAPI)").Range("O" & ubRow).Value = ub.pier_rebar_quantity
                ''''.Worksheets("Details (SAPI)").Range("P" & ubRow).Value = ub.pier_rebar_size
                ''''.Worksheets("Details (SAPI)").Range("Q" & ubRow).Value = ub.pier_tie_quantity
                ''''.Worksheets("Details (SAPI)").Range("R" & ubRow).Value = ub.pier_tie_size
                ''''.Worksheets("Details (SAPI)").Range("S" & ubRow).Value = ub.pier_reinforcement_type
                ''''.Worksheets("Details (SAPI)").Range("T" & ubRow).Value = ub.pier_clear_cover
                ''''.Worksheets("Details (SAPI)").Range("U" & ubRow).Value = ub.pad_width_1
                ''''.Worksheets("Details (SAPI)").Range("V" & ubRow).Value = ub.pad_width_2
                ''''.Worksheets("Details (SAPI)").Range("W" & ubRow).Value = ub.pad_thickness
                ''''.Worksheets("Details (SAPI)").Range("X" & ubRow).Value = ub.pad_rebar_size_top_dir1
                ''''.Worksheets("Details (SAPI)").Range("Y" & ubRow).Value = ub.pad_rebar_size_bottom_dir1
                ''''.Worksheets("Details (SAPI)").Range("Z" & ubRow).Value = ub.pad_rebar_size_top_dir2
                ''''.Worksheets("Details (SAPI)").Range("AA" & ubRow).Value = ub.pad_rebar_size_bottom_dir2
                ''''.Worksheets("Details (SAPI)").Range("AB" & ubRow).Value = ub.pad_rebar_quantity_top_dir1
                ''''.Worksheets("Details (SAPI)").Range("AC" & ubRow).Value = ub.pad_rebar_quantity_bottom_dir1
                ''''.Worksheets("Details (SAPI)").Range("AD" & ubRow).Value = ub.pad_rebar_quantity_top_dir2
                ''''.Worksheets("Details (SAPI)").Range("AE" & ubRow).Value = ub.pad_rebar_quantity_bottom_dir2
                ''''.Worksheets("Details (SAPI)").Range("AF" & ubRow).Value = ub.pad_clear_cover
                ''''.Worksheets("Details (SAPI)").Range("AG" & ubRow).Value = ub.total_soil_unit_weight
                ''''.Worksheets("Details (SAPI)").Range("AH" & ubRow).Value = ub.bearing_type
                ''''.Worksheets("Details (SAPI)").Range("AI" & ubRow).Value = ub.nominal_bearing_capacity
                ''''.Worksheets("Details (SAPI)").Range("AJ" & ubRow).Value = ub.cohesion
                ''''.Worksheets("Details (SAPI)").Range("AK" & ubRow).Value = ub.friction_angle
                ''''.Worksheets("Details (SAPI)").Range("AL" & ubRow).Value = ub.spt_blow_count
                ''''.Worksheets("Details (SAPI)").Range("AM" & ubRow).Value = ub.base_friction_factor
                ''''.Worksheets("Details (SAPI)").Range("AN" & ubRow).Value = ub.neglect_depth
                ''''.Worksheets("Details (SAPI)").Range("AO" & ubRow).Value = ub.bearing_distribution_type
                ''''.Worksheets("Details (SAPI)").Range("AP" & ubRow).Value = ub.groundwater_depth
                ''''ubRow += 1
#End Region

            End With
            SaveAndCloseUnitBase()
        Next
    End Sub

    Private Sub LoadNewUnitBase()
        NewUnitBaseWb.LoadDocument(UnitBaseTemplatePath, UnitBaseFileType)
        NewUnitBaseWb.BeginUpdate()
    End Sub

    Private Sub SaveAndCloseUnitBase()
        NewUnitBaseWb.EndUpdate()
        NewUnitBaseWb.SaveDocument(ExcelFilePath, UnitBaseFileType)
    End Sub
#End Region

#Region "SQL Insert Statements"
    Private Function InsertUnitBaseDetail(ByVal ub As SST_Unit_Base) As String
        Dim insertString As String = ""

        insertString += "@FndID"
        insertString += "," & IIf(IsNothing(ub.pier_shape), "Null", "'" & ub.pier_shape.ToString & "'")
        insertString += "," & IIf(IsNothing(ub.pier_diameter), "Null", ub.pier_diameter.ToString)
        insertString += "," & IIf(IsNothing(ub.extension_above_grade), "Null", ub.extension_above_grade.ToString)
        insertString += "," & IIf(IsNothing(ub.pier_rebar_size), "Null", ub.pier_rebar_size.ToString)
        insertString += "," & IIf(IsNothing(ub.pier_tie_size), "Null", ub.pier_tie_size.ToString)
        insertString += "," & IIf(IsNothing(ub.pier_tie_quantity), "Null", ub.pier_tie_quantity.ToString)
        insertString += "," & IIf(IsNothing(ub.pier_reinforcement_type), "Null", "'" & ub.pier_reinforcement_type.ToString & "'")
        insertString += "," & IIf(IsNothing(ub.pier_clear_cover), "Null", ub.pier_clear_cover.ToString)
        insertString += "," & IIf(IsNothing(ub.foundation_depth), "Null", ub.foundation_depth.ToString)
        insertString += "," & IIf(IsNothing(ub.pad_width_1), "Null", ub.pad_width_1.ToString)
        insertString += "," & IIf(IsNothing(ub.pad_width_2), "Null", ub.pad_width_2.ToString)
        insertString += "," & IIf(IsNothing(ub.pad_thickness), "Null", ub.pad_thickness.ToString)
        insertString += "," & IIf(IsNothing(ub.pad_rebar_size_top_dir1), "Null", ub.pad_rebar_size_top_dir1.ToString)
        insertString += "," & IIf(IsNothing(ub.pad_rebar_size_bottom_dir1), "Null", ub.pad_rebar_size_bottom_dir1.ToString)
        insertString += "," & IIf(IsNothing(ub.pad_rebar_size_top_dir2), "Null", ub.pad_rebar_size_top_dir2.ToString)
        insertString += "," & IIf(IsNothing(ub.pad_rebar_size_bottom_dir2), "Null", ub.pad_rebar_size_bottom_dir2.ToString)
        insertString += "," & IIf(IsNothing(ub.pad_rebar_quantity_top_dir1), "Null", ub.pad_rebar_quantity_top_dir1.ToString)
        insertString += "," & IIf(IsNothing(ub.pad_rebar_quantity_bottom_dir1), "Null", ub.pad_rebar_quantity_bottom_dir1.ToString)
        insertString += "," & IIf(IsNothing(ub.pad_rebar_quantity_top_dir2), "Null", ub.pad_rebar_quantity_top_dir2.ToString)
        insertString += "," & IIf(IsNothing(ub.pad_rebar_quantity_bottom_dir2), "Null", ub.pad_rebar_quantity_bottom_dir2.ToString)
        insertString += "," & IIf(IsNothing(ub.pad_clear_cover), "Null", ub.pad_clear_cover.ToString)
        insertString += "," & IIf(IsNothing(ub.rebar_grade), "Null", ub.rebar_grade.ToString)
        insertString += "," & IIf(IsNothing(ub.concrete_compressive_strength), "Null", ub.concrete_compressive_strength.ToString)
        insertString += "," & IIf(IsNothing(ub.dry_concrete_density), "Null", ub.dry_concrete_density.ToString)
        insertString += "," & IIf(IsNothing(ub.total_soil_unit_weight), "Null", ub.total_soil_unit_weight.ToString)
        insertString += "," & IIf(IsNothing(ub.bearing_type), "Null", "'" & ub.bearing_type.ToString & "'")
        insertString += "," & IIf(IsNothing(ub.nominal_bearing_capacity), "Null", ub.nominal_bearing_capacity.ToString)
        insertString += "," & IIf(IsNothing(ub.cohesion), "Null", ub.cohesion.ToString)
        insertString += "," & IIf(IsNothing(ub.friction_angle), "Null", ub.friction_angle.ToString)
        insertString += "," & IIf(IsNothing(ub.spt_blow_count), "Null", ub.spt_blow_count.ToString)
        insertString += "," & IIf(IsNothing(ub.base_friction_factor), "Null", ub.base_friction_factor.ToString)
        insertString += "," & IIf(IsNothing(ub.neglect_depth), "Null", ub.neglect_depth.ToString)
        insertString += "," & IIf(IsNothing(ub.bearing_distribution_type), "Null", "'" & ub.bearing_distribution_type.ToString & "'")
        insertString += "," & IIf(IsNothing(ub.groundwater_depth), "Null", ub.groundwater_depth.ToString) '*******
        insertString += "," & IIf(IsNothing(ub.top_and_bottom_rebar_different), "Null", "'" & ub.top_and_bottom_rebar_different.ToString & "'")
        insertString += "," & IIf(IsNothing(ub.block_foundation), "Null", "'" & ub.block_foundation.ToString & "'")
        insertString += "," & IIf(IsNothing(ub.rectangular_foundation), "Null", "'" & ub.rectangular_foundation.ToString & "'")
        insertString += "," & IIf(IsNothing(ub.base_plate_distance_above_foundation), "Null", ub.base_plate_distance_above_foundation.ToString)
        insertString += "," & IIf(IsNothing(ub.bolt_circle_bearing_plate_width), "Null", ub.bolt_circle_bearing_plate_width.ToString)
        insertString += "," & IIf(IsNothing(ub.tower_centroid_offset), "Null", "'" & ub.tower_centroid_offset.ToString & "'")
        insertString += "," & IIf(IsNothing(ub.pier_rebar_quantity), "Null", ub.pier_rebar_quantity.ToString)

        Return insertString
    End Function
#End Region

#Region "SQL Update Statements"
    Private Function UpdateUnitBaseDetail(ByVal ub As SST_Unit_Base) As String
        Dim updateString As String = ""

        updateString += "UPDATE unit_base_details SET "
        updateString += "extension_above_grade=" & IIf(IsNothing(ub.extension_above_grade), "Null", ub.extension_above_grade.ToString)
        updateString += ", foundation_depth=" & IIf(IsNothing(ub.foundation_depth), "Null", ub.foundation_depth.ToString)
        updateString += ", concrete_compressive_strength=" & IIf(IsNothing(ub.concrete_compressive_strength), "Null", ub.concrete_compressive_strength.ToString)
        updateString += ", dry_concrete_density=" & IIf(IsNothing(ub.dry_concrete_density), "Null", ub.dry_concrete_density.ToString)
        updateString += ", rebar_grade=" & IIf(IsNothing(ub.rebar_grade), "Null", ub.rebar_grade.ToString)
        updateString += ", top_and_bottom_rebar_different=" & IIf(IsNothing(ub.top_and_bottom_rebar_different), "Null", "'" & ub.top_and_bottom_rebar_different.ToString & "'")
        updateString += ", block_foundation=" & IIf(IsNothing(ub.block_foundation), "Null", "'" & ub.block_foundation.ToString & "'")
        updateString += ", rectangular_foundation=" & IIf(IsNothing(ub.rectangular_foundation), "Null", "'" & ub.rectangular_foundation.ToString & "'")
        updateString += ", base_plate_distance_above_foundation=" & IIf(IsNothing(ub.base_plate_distance_above_foundation), "Null", ub.base_plate_distance_above_foundation.ToString)
        updateString += ", bolt_circle_bearing_plate_width=" & IIf(IsNothing(ub.bolt_circle_bearing_plate_width), "Null", ub.bolt_circle_bearing_plate_width.ToString)
        updateString += ", tower_centroid_offset=" & IIf(IsNothing(ub.tower_centroid_offset), "Null", "'" & ub.tower_centroid_offset.ToString & "'")
        updateString += ", pier_shape=" & IIf(IsNothing(ub.pier_shape), "Null", "'" & ub.pier_shape.ToString & "'")
        updateString += ", pier_diameter=" & IIf(IsNothing(ub.pier_diameter), "Null", ub.pier_diameter.ToString)
        updateString += ", pier_rebar_quantity=" & IIf(IsNothing(ub.pier_rebar_quantity), "Null", ub.pier_rebar_quantity.ToString)
        updateString += ", pier_rebar_size=" & IIf(IsNothing(ub.pier_rebar_size), "Null", ub.pier_rebar_size.ToString)
        updateString += ", pier_tie_quantity=" & IIf(IsNothing(ub.pier_tie_quantity), "Null", ub.pier_tie_quantity.ToString)
        updateString += ", pier_tie_size=" & IIf(IsNothing(ub.pier_tie_size), "Null", ub.pier_tie_size.ToString)
        updateString += ", pier_reinforcement_type=" & IIf(IsNothing(ub.pier_reinforcement_type), "Null", "'" & ub.pier_reinforcement_type.ToString & "'")
        updateString += ", pier_clear_cover=" & IIf(IsNothing(ub.pier_clear_cover), "Null", ub.pier_clear_cover.ToString)
        updateString += ", pad_width_2=" & IIf(IsNothing(ub.pad_width_2), "Null", ub.pad_width_2.ToString)
        updateString += ", pad_thickness=" & IIf(IsNothing(ub.pad_thickness), "Null", ub.pad_thickness.ToString)
        updateString += ", pad_rebar_size_top_dir1=" & IIf(IsNothing(ub.pad_rebar_size_top_dir1), "Null", ub.pad_rebar_size_top_dir1.ToString)
        updateString += ", pad_rebar_size_bottom_dir1=" & IIf(IsNothing(ub.pad_rebar_size_bottom_dir1), "Null", ub.pad_rebar_size_bottom_dir1.ToString)
        updateString += ", pad_rebar_size_top_dir2=" & IIf(IsNothing(ub.pad_rebar_size_top_dir2), "Null", ub.pad_rebar_size_top_dir2.ToString)
        updateString += ", pad_rebar_size_bottom_dir2=" & IIf(IsNothing(ub.pad_rebar_size_bottom_dir2), "Null", ub.pad_rebar_size_bottom_dir2.ToString)
        updateString += ", pad_rebar_quantity_top_dir1=" & IIf(IsNothing(ub.pad_rebar_quantity_top_dir1), "Null", ub.pad_rebar_quantity_top_dir1.ToString)
        updateString += ", pad_rebar_quantity_bottom_dir1=" & IIf(IsNothing(ub.pad_rebar_quantity_bottom_dir1), "Null", ub.pad_rebar_quantity_bottom_dir1.ToString)
        updateString += ", pad_rebar_quantity_top_dir2=" & IIf(IsNothing(ub.pad_rebar_quantity_top_dir2), "Null", ub.pad_rebar_quantity_top_dir2.ToString)
        updateString += ", pad_rebar_quantity_bottom_dir2=" & IIf(IsNothing(ub.pad_rebar_quantity_bottom_dir2), "Null", ub.pad_rebar_quantity_bottom_dir2.ToString)
        updateString += ", pad_clear_cover=" & IIf(IsNothing(ub.pad_clear_cover), "Null", ub.pad_clear_cover.ToString)
        updateString += ", total_soil_unit_weight=" & IIf(IsNothing(ub.total_soil_unit_weight), "Null", ub.total_soil_unit_weight.ToString)
        updateString += ", pad_width_1=" & IIf(IsNothing(ub.pad_width_1), "Null", ub.pad_width_1.ToString)
        updateString += ", bearing_type=" & IIf(IsNothing(ub.bearing_type), "Null", "'" & ub.bearing_type.ToString & "'")
        updateString += ", nominal_bearing_capacity=" & IIf(IsNothing(ub.nominal_bearing_capacity), "Null", ub.nominal_bearing_capacity.ToString)
        updateString += ", cohesion=" & IIf(IsNothing(ub.cohesion), "Null", ub.cohesion.ToString)
        updateString += ", friction_angle=" & IIf(IsNothing(ub.friction_angle), "Null", ub.friction_angle.ToString)
        updateString += ", spt_blow_count=" & IIf(IsNothing(ub.spt_blow_count), "Null", ub.spt_blow_count.ToString)
        updateString += ", base_friction_factor=" & IIf(IsNothing(ub.base_friction_factor), "Null", ub.base_friction_factor.ToString)
        updateString += ", neglect_depth=" & IIf(IsNothing(ub.neglect_depth), "Null", ub.neglect_depth.ToString)
        updateString += ", bearing_distribution_type=" & IIf(IsNothing(ub.bearing_distribution_type), "Null", "'" & ub.bearing_distribution_type.ToString & "'")
        updateString += ", groundwater_depth=" & IIf(IsNothing(ub.groundwater_depth), "Null", ub.groundwater_depth.ToString)
        updateString += " WHERE ID=" & ub.unit_base_id & vbNewLine

        Return updateString
    End Function
#End Region

#Region "General"
    Public Sub Clear()
        ExcelFilePath = ""
        UnitBases.Clear()
    End Sub

    Private Function UnitBaseSQLDataTables() As List(Of SQLParameter)
        Dim MyParameters As New List(Of SQLParameter)
        MyParameters.Add(New SQLParameter("Unit Base General Details SQL", "Unit Base (SELECT Details).sql"))

        Return MyParameters
    End Function

    Private Function UnitBaseExcelDTParameters() As List(Of EXCELDTParameter)
        Dim MyParameters As New List(Of EXCELDTParameter)
        MyParameters.Add(New EXCELDTParameter("Unit Base General Details EXCEL", "A2:AP1000", "Details (SAPI)"))

        Return MyParameters
    End Function

    'Alternate Excel DataLink Option:
    'Private Function UnitBaseExcelRngParameters() As List(Of EXCELRngParameter)
    '    Dim MyParameters As New List(Of EXCELRngParameter)

    '    MyParameters.Add(New EXCELRngParameter("ID", "unit_base_id"))
    '    MyParameters.Add(New EXCELRngParameter("E", "extension_above_grade"))
    '    MyParameters.Add(New EXCELRngParameter("D", "foundation_depth"))
    '    MyParameters.Add(New EXCELRngParameter("F\c", "concrete_compressive_strength"))
    '    MyParameters.Add(New EXCELRngParameter("ConcreteDensity", "dry_concrete_density"))
    '    MyParameters.Add(New EXCELRngParameter("Fy", "rebar_grade"))
    '    MyParameters.Add(New EXCELRngParameter("DifferentReinforcementBoolean", "top_and_bottom_rebar_different"))
    '    MyParameters.Add(New EXCELRngParameter("BlockFoundationBoolean", "block_foundation"))
    '    MyParameters.Add(New EXCELRngParameter("RectangularPadBoolean", "rectangular_foundation"))
    '    MyParameters.Add(New EXCELRngParameter("bpdist", "base_plate_distance_above_foundation"))
    '    MyParameters.Add(New EXCELRngParameter("BC", "bolt_circle_bearing_plate_width"))
    '    MyParameters.Add(New EXCELRngParameter("TowerCentroidOffsetBoolean", "tower_centroid_offset"))
    '    MyParameters.Add(New EXCELRngParameter("shape", "pier_shape"))
    '    MyParameters.Add(New EXCELRngParameter("dpier", "pier_diameter"))
    '    MyParameters.Add(New EXCELRngParameter("mc", "pier_rebar_quantity"))
    '    MyParameters.Add(New EXCELRngParameter("Sc", "pier_rebar_size"))
    '    MyParameters.Add(New EXCELRngParameter("mt", "pier_tie_quantity"))
    '    MyParameters.Add(New EXCELRngParameter("St", "pier_tie_size"))
    '    MyParameters.Add(New EXCELRngParameter("PierReinfType", "pier_reinforcement_type"))
    '    MyParameters.Add(New EXCELRngParameter("ccpier", "pier_clear_cover"))
    '    MyParameters.Add(New EXCELRngParameter("W", "pad_width_1"))
    '    MyParameters.Add(New EXCELRngParameter("W.dir2", "pad_width_2"))
    '    MyParameters.Add(New EXCELRngParameter("T", "pad_thickness"))
    '    MyParameters.Add(New EXCELRngParameter("sptop", "pad_rebar_size_top_dir1"))
    '    MyParameters.Add(New EXCELRngParameter("Sp", "pad_rebar_size_bottom_dir1"))
    '    MyParameters.Add(New EXCELRngParameter("sptop2", "pad_rebar_size_top_dir2"))
    '    MyParameters.Add(New EXCELRngParameter("sp_2", "pad_rebar_size_bottom_dir2"))
    '    MyParameters.Add(New EXCELRngParameter("mptop", "pad_rebar_quantity_top_dir1"))
    '    MyParameters.Add(New EXCELRngParameter("mp", "pad_rebar_quantity_bottom_dir1"))
    '    MyParameters.Add(New EXCELRngParameter("mptop2", "pad_rebar_quantity_top_dir2"))
    '    MyParameters.Add(New EXCELRngParameter("mp_2", "pad_rebar_quantity_bottom_dir2"))
    '    MyParameters.Add(New EXCELRngParameter("ccpad", "pad_clear_cover"))
    '    MyParameters.Add(New EXCELRngParameter("γ", "total_soil_unit_weight"))
    '    MyParameters.Add(New EXCELRngParameter("BearingType", "bearing_type"))
    '    MyParameters.Add(New EXCELRngParameter("Qinput", "nominal_bearing_capacity"))
    '    MyParameters.Add(New EXCELRngParameter("Cu", "cohesion"))
    '    MyParameters.Add(New EXCELRngParameter("ϕ", "friction_angle"))
    '    MyParameters.Add(New EXCELRngParameter("N_blows", "spt_blow_count"))
    '    MyParameters.Add(New EXCELRngParameter("μ", "base_friction_factor"))
    '    MyParameters.Add(New EXCELRngParameter("N", "neglect_depth"))
    '    MyParameters.Add(New EXCELRngParameter("Rock", "bearing_distribution_type"))
    '    MyParameters.Add(New EXCELRngParameter("gw", "groundwater_depth"))

    '    Return MyParameters
    'End Function
#End Region

End Class
