Option Strict Off

Imports DevExpress.Spreadsheet
Imports System.Security.Principal

Partial Public Class DataTransfererPierandPad

#Region "Define"
    Private NewPierAndPadWb As New Workbook
    Private prop_ExcelFilePath As String

    Public Property PierAndPads As New List(Of Pier_and_Pad)
    Private Property PierAndPadTemplatePath As String = "C:\Users\" & Environment.UserName & "\source\repos\DevExpress Objects\Pier and Pad Foundation (4.1.0) - EDS.xlsm"
    Private Property PierAndPadFileType As DocumentFormat = DocumentFormat.Xlsm

    Public Property ppDS As New DataSet
    Public Property ppDB As String
    Public Property ppID As WindowsIdentity

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
        ppDS = MyDataSet
        ppID = LogOnUser
        ppDB = ActiveDatabase
        BUNumber = BU
        STR_ID = Strucutre_ID
    End Sub
#End Region

#Region "Load Data"
    Public Function LoadFromEDS() As Boolean
        Dim refid As Integer
        Dim PierAndPadLoader As String

        'Load data to get pier and pad details data for the existing structure model
        For Each item As SQLParameter In PierAndPadSQLDataTables()
            PierAndPadLoader = QueryBuilderFromFile(queryPath & "Pier and Pad\" & item.sqlQuery).Replace("[EXISTING MODEL]", GetExistingModelQuery())
            DoDaSQL.sqlLoader(PierAndPadLoader, item.sqlDatatable, ppDS, ppDB, ppID, "0")
            If ppDS.Tables(item.sqlDatatable).Rows.Count = 0 Then Return False
        Next

        'Custom Section to transfer data for the pier and pad tool. Needs to be adjusted for each tool.
        For Each Pier_And_PadDataRow As DataRow In ppDS.Tables("Pier and Pad General Details SQL").Rows
            refid = CType(Pier_And_PadDataRow.Item("pp_id"), Integer)

            PierAndPads.Add(New Pier_and_Pad(Pier_And_PadDataRow, refid))
        Next

        Return True
    End Function 'Create Pier and Pad objects based on what is saved in EDS

    Public Sub LoadFromExcel()
        PierAndPads.Add(New Pier_and_Pad(ExcelFilePath))
    End Sub 'Create Pier and Pad objects based on what is coming from the excel file
#End Region

#Region "Save Data"
    Public Sub SaveToEDS()
        Dim firstOne As Boolean = True

        For Each pp As Pier_and_Pad In PierAndPads
            Dim PierAndPadSaver As String = QueryBuilderFromFile(queryPath & "Pier and Pad\Pier and Pad (IN_UP).sql")

            PierAndPadSaver = PierAndPadSaver.Replace("[BU NUMBER]", BUNumber)
            PierAndPadSaver = PierAndPadSaver.Replace("[STRUCTURE ID]", STR_ID)
            PierAndPadSaver = PierAndPadSaver.Replace("[FOUNDATION TYPE]", "Pier and Pad")
            If pp.pp_id = 0 Or IsDBNull(pp.pp_id) Then
                PierAndPadSaver = PierAndPadSaver.Replace("'[PIER AND PAD ID]'", "NULL")
            Else
                PierAndPadSaver = PierAndPadSaver.Replace("[PIER AND PAD ID]", pp.pp_id.ToString)
                PierAndPadSaver = PierAndPadSaver.Replace("(SELECT * FROM TEMPORARY)", UpdatePierAndPadDetail(pp))
            End If
            PierAndPadSaver = PierAndPadSaver.Replace("[INSERT ALL PIER AND PAD DETAILS]", InsertPierAndPadDetail(pp))

            sqlSender(PierAndPadSaver, ppDB, ppID, "0")
        Next
    End Sub

    Public Sub SaveToExcel()
        For Each pp As Pier_and_Pad In PierAndPads
            'MRP 7/20/21 - clear contents in Excel if object value is nothing (NULL from SQL). If multiple objects exist within the list and values are not cleared, values from another instance of the object may remain populated in the latest tool
            LoadNewPierAndPad()

            With NewPierAndPadWb
                .Worksheets("Input").Range("ID").Value = CType(pp.pp_id, Integer)
                If Not IsNothing(pp.pier_shape) Then
                    .Worksheets("Input").Range("shape").Value = pp.pier_shape
                End If

                If Not IsNothing(pp.pier_diameter) Then
                    .Worksheets("Input").Range("dpier").Value = CType(pp.pier_diameter, Double)
                Else
                    .Worksheets("Input").Range("dpier").ClearContents
                End If

                If Not IsNothing(pp.extension_above_grade) Then
                    .Worksheets("Input").Range("E").Value = CType(pp.extension_above_grade, Double)
                Else
                    .Worksheets("Input").Range("E").ClearContents
                End If

                If Not IsNothing(pp.pier_rebar_size) Then
                    .Worksheets("Input").Range("Sc").Value = CType(pp.pier_rebar_size, Integer)
                Else
                    .Worksheets("Input").Range("Sc").ClearContents
                End If

                If Not IsNothing(pp.pier_rebar_quantity) Then
                    .Worksheets("Input").Range("mc").Value = CType(pp.pier_rebar_quantity, Integer)
                Else
                    .Worksheets("Input").Range("mc").ClearContents
                End If

                If Not IsNothing(pp.pier_tie_size) Then
                    .Worksheets("Input").Range("St").Value = CType(pp.pier_tie_size, Integer)
                Else
                    .Worksheets("Input").Range("St").ClearContents
                End If

                If Not IsNothing(pp.pier_tie_quantity) Then
                    .Worksheets("Input").Range("mt").Value = CType(pp.pier_tie_quantity, Integer)
                Else
                    .Worksheets("Input").Range("mt").ClearContents
                End If

                If Not IsNothing(pp.pier_reinforcement_type) Then
                    .Worksheets("Input").Range("PierReinfType").Value = pp.pier_reinforcement_type
                End If

                If Not IsNothing(pp.pier_clear_cover) Then
                    .Worksheets("Input").Range("ccpier").Value = CType(pp.pier_clear_cover, Double)
                Else
                    .Worksheets("Input").Range("ccpier").ClearContents
                End If

                If Not IsNothing(pp.foundation_depth) Then
                    .Worksheets("Input").Range("D").Value = CType(pp.foundation_depth, Double)
                Else
                    .Worksheets("Input").Range("D").ClearContents
                End If

                If Not IsNothing(pp.pad_width_1) Then
                    .Worksheets("Input").Range("W").Value = CType(pp.pad_width_1, Double)
                Else
                    .Worksheets("Input").Range("W").ClearContents
                End If

                If Not IsNothing(pp.pad_width_2) Then
                    .Worksheets("Input").Range("W.dir2").Value = CType(pp.pad_width_2, Double)
                Else
                    .Worksheets("Input").Range("W.dir2").ClearContents
                End If

                If Not IsNothing(pp.pad_thickness) Then
                    .Worksheets("Input").Range("T").Value = CType(pp.pad_thickness, Double)
                Else
                    .Worksheets("Input").Range("T").ClearContents
                End If

                If Not IsNothing(pp.pad_rebar_size_top_dir1) Then
                    .Worksheets("Input").Range("sptop").Value = CType(pp.pad_rebar_size_top_dir1, Integer)
                Else
                    .Worksheets("Input").Range("sptop").ClearContents
                End If

                If Not IsNothing(pp.pad_rebar_size_bottom_dir1) Then
                    .Worksheets("Input").Range("Sp").Value = CType(pp.pad_rebar_size_bottom_dir1, Integer)
                Else
                    .Worksheets("Input").Range("Sp").ClearContents
                End If

                If Not IsNothing(pp.pad_rebar_size_top_dir2) Then
                    .Worksheets("Input").Range("sptop2").Value = CType(pp.pad_rebar_size_top_dir2, Integer)
                Else
                    .Worksheets("Input").Range("sptop2").ClearContents
                End If

                If Not IsNothing(pp.pad_rebar_size_bottom_dir2) Then
                    .Worksheets("Input").Range("sp_2").Value = CType(pp.pad_rebar_size_bottom_dir2, Integer)
                Else
                    .Worksheets("Input").Range("sp_2").ClearContents
                End If

                If Not IsNothing(pp.pad_rebar_quantity_top_dir1) Then
                    .Worksheets("Input").Range("mptop").Value = CType(pp.pad_rebar_quantity_top_dir1, Integer)
                Else
                    .Worksheets("Input").Range("mptop").ClearContents
                End If

                If Not IsNothing(pp.pad_rebar_quantity_bottom_dir1) Then
                    .Worksheets("Input").Range("mp").Value = CType(pp.pad_rebar_quantity_bottom_dir1, Integer)
                Else
                    .Worksheets("Input").Range("mp").ClearContents
                End If

                If Not IsNothing(pp.pad_rebar_quantity_top_dir2) Then
                    .Worksheets("Input").Range("mptop2").Value = CType(pp.pad_rebar_quantity_top_dir2, Integer)
                Else
                    .Worksheets("Input").Range("mptop2").ClearContents
                End If

                If Not IsNothing(pp.pad_rebar_quantity_bottom_dir2) Then
                    .Worksheets("Input").Range("mp_2").Value = CType(pp.pad_rebar_quantity_bottom_dir2, Integer)
                Else
                    .Worksheets("Input").Range("mp_2").ClearContents
                End If

                If Not IsNothing(pp.pad_clear_cover) Then
                    .Worksheets("Input").Range("ccpad").Value = CType(pp.pad_clear_cover, Double)
                Else
                    .Worksheets("Input").Range("ccpad").ClearContents
                End If

                If Not IsNothing(pp.rebar_grade) Then
                    .Worksheets("Input").Range("Fy").Value = CType(pp.rebar_grade, Double)
                Else
                    .Worksheets("Input").Range("Fy").ClearContents
                End If

                If Not IsNothing(pp.concrete_compressive_strength) Then
                    .Worksheets("Input").Range("F\c").Value = CType(pp.concrete_compressive_strength, Double)
                Else
                    .Worksheets("Input").Range("F\c").ClearContents
                End If

                If Not IsNothing(pp.dry_concrete_density) Then
                    .Worksheets("Input").Range("ConcreteDensity").Value = CType(pp.dry_concrete_density, Double)
                Else
                    .Worksheets("Input").Range("ConcreteDensity").ClearContents
                End If

                If Not IsNothing(pp.total_soil_unit_weight) Then
                    .Worksheets("Input").Range("γ").Value = CType(pp.total_soil_unit_weight, Double)
                Else
                    .Worksheets("Input").Range("γ").ClearContents
                End If

                If Not IsNothing(pp.bearing_type) Then
                    .Worksheets("Input").Range("BearingType").Value = pp.bearing_type
                Else
                    .Worksheets("Input").Range("BearingType").ClearContents
                End If

                If Not IsNothing(pp.nominal_bearing_capacity) Then
                    .Worksheets("Input").Range("Qinput").Value = CType(pp.nominal_bearing_capacity, Double)
                Else
                    .Worksheets("Input").Range("Qinput").ClearContents
                End If

                If Not IsNothing(pp.cohesion) Then
                    .Worksheets("Input").Range("Cu").Value = CType(pp.cohesion, Double)
                Else
                    .Worksheets("Input").Range("Cu").ClearContents
                End If

                If Not IsNothing(pp.friction_angle) Then
                    .Worksheets("Input").Range("ϕ").Value = CType(pp.friction_angle, Double)
                Else
                    .Worksheets("Input").Range("ϕ").ClearContents
                End If

                If Not IsNothing(pp.spt_blow_count) Then
                    .Worksheets("Input").Range("N_blows").Value = CType(pp.spt_blow_count, Integer)
                Else
                    .Worksheets("Input").Range("N_blows").ClearContents
                End If

                If Not IsNothing(pp.base_friction_factor) Then
                    .Worksheets("Input").Range("μ").Value = CType(pp.base_friction_factor, Double)
                Else
                    .Worksheets("Input").Range("μ").ClearContents
                End If

                If Not IsNothing(pp.neglect_depth) Then
                    .Worksheets("Input").Range("N").Value = CType(pp.neglect_depth, Double)
                End If

                If pp.bearing_distribution_type = True Then
                    .Worksheets("Input").Range("Rock").Value = "No"
                Else
                    .Worksheets("Input").Range("Rock").Value = "Yes"
                End If

                If pp.groundwater_depth = -1 Then
                    .Worksheets("Input").Range("gw").Value = "N/A"
                Else
                    .Worksheets("Input").Range("gw").Value = CType(pp.groundwater_depth, Double)
                End If

                If Not IsNothing(pp.top_and_bottom_rebar_different) Then
                    .Worksheets("Input").Range("DifferentReinforcementBoolean").Value = pp.top_and_bottom_rebar_different
                End If

                If Not IsNothing(pp.block_foundation) Then
                    .Worksheets("Input").Range("BlockFoundationBoolean").Value = pp.block_foundation
                End If

                If Not IsNothing(pp.rectangular_foundation) Then
                    .Worksheets("Input").Range("RectangularPadBoolean").Value = pp.rectangular_foundation
                End If

                If Not IsNothing(pp.base_plate_distance_above_foundation) Then
                    .Worksheets("Input").Range("bpdist").Value = CType(pp.base_plate_distance_above_foundation, Double)
                Else
                    .Worksheets("Input").Range("bpdist").ClearContents
                End If

                If Not IsNothing(pp.bolt_circle_bearing_plate_width) Then
                    .Worksheets("Input").Range("BC").Value = CType(pp.bolt_circle_bearing_plate_width, Double)
                Else
                    .Worksheets("Input").Range("BC").ClearContents
                End If
            End With

            SaveAndClosePierAndPad()
        Next

    End Sub

    Private Sub LoadNewPierAndPad()
        NewPierAndPadWb.LoadDocument(PierAndPadTemplatePath, PierAndPadFileType)
        NewPierAndPadWb.BeginUpdate()
    End Sub

    Private Sub SaveAndClosePierAndPad()
        NewPierAndPadWb.EndUpdate()
        NewPierAndPadWb.SaveDocument(ExcelFilePath, PierAndPadFileType)
    End Sub
#End Region

#Region "SQL Insert Statements"
    Private Function InsertPierAndPadDetail(ByVal pp As Pier_and_Pad) As String
        Dim insertString As String = ""

        insertString += "@FndID"
        insertString += "," & IIf(IsNothing(pp.pier_shape), "Null", "'" & pp.pier_shape.ToString & "'")
        insertString += "," & IIf(IsNothing(pp.pier_diameter), "Null", pp.pier_diameter.ToString)
        insertString += "," & IIf(IsNothing(pp.extension_above_grade), "Null", pp.extension_above_grade.ToString)
        insertString += "," & IIf(IsNothing(pp.pier_rebar_size), "Null", pp.pier_rebar_size.ToString)
        insertString += "," & IIf(IsNothing(pp.pier_tie_size), "Null", pp.pier_tie_size.ToString)
        insertString += "," & IIf(IsNothing(pp.pier_tie_quantity), "Null", pp.pier_tie_quantity.ToString)
        insertString += "," & IIf(IsNothing(pp.pier_reinforcement_type), "Null", "'" & pp.pier_reinforcement_type.ToString & "'")
        insertString += "," & IIf(IsNothing(pp.pier_clear_cover), "Null", pp.pier_clear_cover.ToString)
        insertString += "," & IIf(IsNothing(pp.foundation_depth), "Null", pp.foundation_depth.ToString)
        insertString += "," & IIf(IsNothing(pp.pad_width_1), "Null", pp.pad_width_1.ToString)
        insertString += "," & IIf(IsNothing(pp.pad_width_2), "Null", pp.pad_width_2.ToString)
        insertString += "," & IIf(IsNothing(pp.pad_thickness), "Null", pp.pad_thickness.ToString)
        insertString += "," & IIf(IsNothing(pp.pad_rebar_size_top_dir1), "Null", pp.pad_rebar_size_top_dir1.ToString)
        insertString += "," & IIf(IsNothing(pp.pad_rebar_size_bottom_dir1), "Null", pp.pad_rebar_size_bottom_dir1.ToString)
        insertString += "," & IIf(IsNothing(pp.pad_rebar_size_top_dir2), "Null", pp.pad_rebar_size_top_dir2.ToString)
        insertString += "," & IIf(IsNothing(pp.pad_rebar_size_bottom_dir2), "Null", pp.pad_rebar_size_bottom_dir2.ToString)
        insertString += "," & IIf(IsNothing(pp.pad_rebar_quantity_top_dir1), "Null", pp.pad_rebar_quantity_top_dir1.ToString)
        insertString += "," & IIf(IsNothing(pp.pad_rebar_quantity_bottom_dir1), "Null", pp.pad_rebar_quantity_bottom_dir1.ToString)
        insertString += "," & IIf(IsNothing(pp.pad_rebar_quantity_top_dir2), "Null", pp.pad_rebar_quantity_top_dir2.ToString)
        insertString += "," & IIf(IsNothing(pp.pad_rebar_quantity_bottom_dir2), "Null", pp.pad_rebar_quantity_bottom_dir2.ToString)
        insertString += "," & IIf(IsNothing(pp.pad_clear_cover), "Null", pp.pad_clear_cover.ToString)
        insertString += "," & IIf(IsNothing(pp.rebar_grade), "Null", pp.rebar_grade.ToString)
        insertString += "," & IIf(IsNothing(pp.concrete_compressive_strength), "Null", pp.concrete_compressive_strength.ToString)
        insertString += "," & IIf(IsNothing(pp.dry_concrete_density), "Null", pp.dry_concrete_density.ToString)
        insertString += "," & IIf(IsNothing(pp.total_soil_unit_weight), "Null", pp.total_soil_unit_weight.ToString)
        insertString += "," & IIf(IsNothing(pp.bearing_type), "Null", "'" & pp.bearing_type.ToString & "'")
        insertString += "," & IIf(IsNothing(pp.nominal_bearing_capacity), "Null", pp.nominal_bearing_capacity.ToString)
        insertString += "," & IIf(IsNothing(pp.cohesion), "Null", pp.cohesion.ToString)
        insertString += "," & IIf(IsNothing(pp.friction_angle), "Null", pp.friction_angle.ToString)
        insertString += "," & IIf(IsNothing(pp.spt_blow_count), "Null", pp.spt_blow_count.ToString)
        insertString += "," & IIf(IsNothing(pp.base_friction_factor), "Null", pp.base_friction_factor.ToString)
        insertString += "," & IIf(IsNothing(pp.neglect_depth), "Null", pp.neglect_depth.ToString)
        insertString += "," & IIf(IsNothing(pp.bearing_distribution_type), "Null", "'" & pp.bearing_distribution_type.ToString & "'")
        insertString += "," & IIf(IsNothing(pp.groundwater_depth), "Null", pp.groundwater_depth.ToString)
        insertString += "," & IIf(IsNothing(pp.top_and_bottom_rebar_different), "Null", "'" & pp.top_and_bottom_rebar_different.ToString & "'")
        insertString += "," & IIf(IsNothing(pp.block_foundation), "Null", "'" & pp.block_foundation.ToString & "'")
        insertString += "," & IIf(IsNothing(pp.rectangular_foundation), "Null", "'" & pp.rectangular_foundation.ToString & "'")
        insertString += "," & IIf(IsNothing(pp.base_plate_distance_above_foundation), "Null", pp.base_plate_distance_above_foundation.ToString)
        insertString += "," & IIf(IsNothing(pp.bolt_circle_bearing_plate_width), "Null", pp.bolt_circle_bearing_plate_width.ToString)
        insertString += "," & IIf(IsNothing(pp.pier_rebar_quantity), "Null", pp.pier_rebar_quantity.ToString)

        Return insertString
    End Function
#End Region

#Region "SQL Update Statements"
    Private Function UpdatePierAndPadDetail(ByVal pp As Pier_and_Pad) As String
        Dim updateString As String = ""

        updateString += "UPDATE pier_pad_details SET "
        updateString += " pier_shape=" & IIf(IsNothing(pp.pier_shape), "Null", "'" & pp.pier_shape.ToString & "'")
        updateString += ", pier_diameter=" & IIf(IsNothing(pp.pier_diameter), "Null", pp.pier_diameter.ToString)
        updateString += ", extension_above_grade=" & IIf(IsNothing(pp.extension_above_grade), "Null", pp.extension_above_grade.ToString)
        updateString += ", pier_rebar_size=" & IIf(IsNothing(pp.pier_rebar_size), "Null", pp.pier_rebar_size.ToString)
        updateString += ", pier_tie_size=" & IIf(IsNothing(pp.pier_tie_size), "Null", pp.pier_tie_size.ToString)
        updateString += ", pier_tie_quantity=" & IIf(IsNothing(pp.pier_tie_quantity), "Null", pp.pier_tie_quantity.ToString)
        updateString += ", pier_reinforcement_type=" & IIf(IsNothing(pp.pier_reinforcement_type), "Null", "'" & pp.pier_reinforcement_type.ToString & "'")
        updateString += ", pier_clear_cover=" & IIf(IsNothing(pp.pier_clear_cover), "Null", pp.pier_clear_cover.ToString)
        updateString += ", foundation_depth=" & IIf(IsNothing(pp.foundation_depth), "Null", pp.foundation_depth.ToString)
        updateString += ", pad_width_1=" & IIf(IsNothing(pp.pad_width_1), "Null", pp.pad_width_1.ToString)
        updateString += ", pad_width_2=" & IIf(IsNothing(pp.pad_width_2), "Null", pp.pad_width_2.ToString)
        updateString += ", pad_thickness=" & IIf(IsNothing(pp.pad_thickness), "Null", pp.pad_thickness.ToString)
        updateString += ", pad_rebar_size_top_dir1=" & IIf(IsNothing(pp.pad_rebar_size_top_dir1), "Null", pp.pad_rebar_size_top_dir1.ToString)
        updateString += ", pad_rebar_size_bottom_dir1=" & IIf(IsNothing(pp.pad_rebar_size_bottom_dir1), "Null", pp.pad_rebar_size_bottom_dir1.ToString)
        updateString += ", pad_rebar_size_top_dir2=" & IIf(IsNothing(pp.pad_rebar_size_top_dir2), "Null", pp.pad_rebar_size_top_dir2.ToString)
        updateString += ", pad_rebar_size_bottom_dir2=" & IIf(IsNothing(pp.pad_rebar_size_bottom_dir2), "Null", pp.pad_rebar_size_bottom_dir2.ToString)
        updateString += ", pad_rebar_quantity_top_dir1=" & IIf(IsNothing(pp.pad_rebar_quantity_top_dir1), "Null", pp.pad_rebar_quantity_top_dir1.ToString)
        updateString += ", pad_rebar_quantity_bottom_dir1=" & IIf(IsNothing(pp.pad_rebar_quantity_bottom_dir1), "Null", pp.pad_rebar_quantity_bottom_dir1.ToString)
        updateString += ", pad_rebar_quantity_top_dir2=" & IIf(IsNothing(pp.pad_rebar_quantity_top_dir2), "Null", pp.pad_rebar_quantity_top_dir2.ToString)
        updateString += ", pad_rebar_quantity_bottom_dir2=" & IIf(IsNothing(pp.pad_rebar_quantity_bottom_dir2), "Null", pp.pad_rebar_quantity_bottom_dir2.ToString)
        updateString += ", pad_clear_cover=" & IIf(IsNothing(pp.pad_clear_cover), "Null", pp.pad_clear_cover.ToString)
        updateString += ", rebar_grade=" & IIf(IsNothing(pp.rebar_grade), "Null", pp.rebar_grade.ToString)
        updateString += ", concrete_compressive_strength=" & IIf(IsNothing(pp.concrete_compressive_strength), "Null", pp.concrete_compressive_strength.ToString)
        updateString += ", dry_concrete_density=" & IIf(IsNothing(pp.dry_concrete_density), "Null", pp.dry_concrete_density.ToString)
        updateString += ", total_soil_unit_weight=" & IIf(IsNothing(pp.total_soil_unit_weight), "Null", pp.total_soil_unit_weight.ToString)
        updateString += ", bearing_type=" & IIf(IsNothing(pp.bearing_type), "Null", "'" & pp.bearing_type.ToString & "'")
        updateString += ", nominal_bearing_capacity=" & IIf(IsNothing(pp.nominal_bearing_capacity), "Null", pp.nominal_bearing_capacity.ToString)
        updateString += ", cohesion=" & IIf(IsNothing(pp.cohesion), "Null", pp.cohesion.ToString)
        updateString += ", friction_angle=" & IIf(IsNothing(pp.friction_angle), "Null", pp.friction_angle.ToString)
        updateString += ", spt_blow_count=" & IIf(IsNothing(pp.spt_blow_count), "Null", pp.spt_blow_count.ToString)
        updateString += ", base_friction_factor=" & IIf(IsNothing(pp.base_friction_factor), "Null", pp.base_friction_factor.ToString)
        updateString += ", neglect_depth=" & IIf(IsNothing(pp.neglect_depth), "Null", pp.neglect_depth.ToString)
        updateString += ", bearing_distribution_type=" & IIf(IsNothing(pp.bearing_distribution_type), "Null", "'" & pp.bearing_distribution_type.ToString & "'")
        updateString += ", groundwater_depth=" & IIf(IsNothing(pp.groundwater_depth), "Null", pp.groundwater_depth.ToString)
        updateString += ", top_and_bottom_rebar_different=" & IIf(IsNothing(pp.top_and_bottom_rebar_different), "Null", "'" & pp.top_and_bottom_rebar_different.ToString & "'")
        updateString += ", block_foundation=" & IIf(IsNothing(pp.block_foundation), "Null", "'" & pp.block_foundation.ToString & "'")
        updateString += ", rectangular_foundation=" & IIf(IsNothing(pp.rectangular_foundation), "Null", "'" & pp.rectangular_foundation.ToString & "'")
        updateString += ", base_plate_distance_above_foundation=" & IIf(IsNothing(pp.base_plate_distance_above_foundation), "Null", pp.base_plate_distance_above_foundation.ToString)
        updateString += ", bolt_circle_bearing_plate_width=" & IIf(IsNothing(pp.bolt_circle_bearing_plate_width), "Null", pp.bolt_circle_bearing_plate_width.ToString)
        updateString += ", pier_rebar_quantity=" & IIf(IsNothing(pp.pier_rebar_quantity), "Null", pp.pier_rebar_quantity.ToString)
        updateString += " WHERE ID = " & pp.pp_id.ToString

        Return updateString
    End Function
#End Region

#Region "General"
    Public Sub Clear()
        ExcelFilePath = ""
        PierAndPads.Clear()
    End Sub

    Private Function PierAndPadSQLDataTables() As List(Of SQLParameter)
        Dim MyParameters As New List(Of SQLParameter)

        MyParameters.Add(New SQLParameter("Pier and Pad General Details SQL", "Pier and Pad (SELECT Details).sql"))

        Return MyParameters
    End Function

    'Private Function Pier_And_PadExcelDTParameters() As List(Of EXCELDTParameter) 'MRP - Edit this code to reference Input worksheet rather than Details table
    '    Dim MyParameters As New List(Of EXCELDTParameter)

    '    MyParameters.Add(New EXCELDTParameter("Pier and Pad General Details EXCEL", "A2:AP1000", "Details (SAPI)"))

    '    Return MyParameters
    'End Function

    'Private Function Pier_And_PadExcelRngParameters() As List(Of EXCELRngParameter)
    'Dim myLst As New List(Of EXCELRngParameter)

    '''''''myLst.Add(New EXCELRngParameter("test_value", "soil_layer_Count"))
    '''''''myLst.Add(New EXCELRngParameter("ID", "pp_id"))
    '''''''myLst.Add(New EXCELRngParameter("E", "extension_above_grade"))
    '''''''myLst.Add(New EXCELRngParameter("D", "foundation_depth"))
    '''''''myLst.Add(New EXCELRngParameter("F\c", "concrete_compressive_strength"))
    '''''''myLst.Add(New EXCELRngParameter("ConcreteDensity", "dry_concrete_density"))
    '''''''myLst.Add(New EXCELRngParameter("Fy", "rebar_grade"))
    '''''''myLst.Add(New EXCELRngParameter("DifferentReinforcementBoolean", "top_and_bottom_rebar_different"))
    '''''''myLst.Add(New EXCELRngParameter("BlockFoundationBoolean", "block_foundation"))
    '''''''myLst.Add(New EXCELRngParameter("RectangularPadBoolean", "rectangular_foundation"))
    '''''''myLst.Add(New EXCELRngParameter("bpdist", "base_plate_distance_above_foundation"))
    '''''''myLst.Add(New EXCELRngParameter("BC", "bolt_circle_bearing_plate_width"))
    '''''''myLst.Add(New EXCELRngParameter("shape", "pier_shape"))
    '''''''myLst.Add(New EXCELRngParameter("dpier", "pier_diameter"))
    '''''''myLst.Add(New EXCELRngParameter("mc", "pier_rebar_quantity"))
    '''''''myLst.Add(New EXCELRngParameter("Sc", "pier_rebar_size"))
    '''''''myLst.Add(New EXCELRngParameter("mt", "pier_tie_quantity"))
    '''''''myLst.Add(New EXCELRngParameter("St", "pier_tie_size"))
    '''''''myLst.Add(New EXCELRngParameter("PierReinfType", "pier_reinforcement_type"))
    '''''''myLst.Add(New EXCELRngParameter("ccpier", "pier_clear_cover"))
    '''''''myLst.Add(New EXCELRngParameter("W", "pad_width_1"))
    '''''''myLst.Add(New EXCELRngParameter("W.dir2", "pad_width_2"))
    '''''''myLst.Add(New EXCELRngParameter("T", "pad_thickness"))
    '''''''myLst.Add(New EXCELRngParameter("sptop", "pad_rebar_size_top_dir1"))
    '''''''myLst.Add(New EXCELRngParameter("Sp", "pad_rebar_size_bottom_dir1"))
    '''''''myLst.Add(New EXCELRngParameter("sptop2", "pad_rebar_size_top_dir2"))
    '''''''myLst.Add(New EXCELRngParameter("sp_2", "pad_rebar_size_bottom_dir2"))
    '''''''myLst.Add(New EXCELRngParameter("mptop", "pad_rebar_quantity_top_dir1"))
    '''''''myLst.Add(New EXCELRngParameter("mp", "pad_rebar_quantity_bottom_dir1"))
    '''''''myLst.Add(New EXCELRngParameter("mptop2", "pad_rebar_quantity_top_dir2"))
    '''''''myLst.Add(New EXCELRngParameter("mp_2", "pad_rebar_quantity_bottom_dir2"))
    '''''''myLst.Add(New EXCELRngParameter("ccpad", "pad_clear_cover"))
    '''''''myLst.Add(New EXCELRngParameter("γ", "total_soil_unit_weight"))
    '''''''myLst.Add(New EXCELRngParameter("BearingType", "bearing_type"))
    '''''''myLst.Add(New EXCELRngParameter("Qinput", "nominal_bearing_capacity"))
    '''''''myLst.Add(New EXCELRngParameter("Cu", "cohesion"))
    '''''''myLst.Add(New EXCELRngParameter("ϕ", "friction_angle"))
    '''''''myLst.Add(New EXCELRngParameter("N_blows", "spt_blow_count"))
    '''''''myLst.Add(New EXCELRngParameter("μ", "base_friction_factor"))
    '''''''myLst.Add(New EXCELRngParameter("N", "neglect_depth"))
    '''''''myLst.Add(New EXCELRngParameter("Rock", "bearing_distribution_type"))
    '''''''myLst.Add(New EXCELRngParameter("gw", "groundwater_depth"))

    '''''''Return myLst

    'Dim myvar As String

    'For Each itm As EXCELRngParameter In myLst
    '    If itm.variableName = "pp_id" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "extension_above_grade" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "foundation_depth" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "concrete_compressive_strength" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "dry_concrete_density" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "rebar_grade" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "top_and_bottom_rebar_different" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "block_foundation" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "rectangular_foundation" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "base_plate_distance_above_foundation" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "bolt_circle_bearing_plate_width" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "pier_shape" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "pier_diameter" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "pier_rebar_quantity" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "pier_rebar_size" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "pier_tie_quantity" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "pier_tie_size" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "pier_reinforcement_type" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "pier_clear_cover" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "pad_width_1" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "pad_width_2" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "pad_thickness" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "pad_rebar_size_top_dir1" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "pad_rebar_size_bottom_dir1" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "pad_rebar_size_top_dir2" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "pad_rebar_size_bottom_dir2" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "pad_rebar_quantity_top_dir1" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "pad_rebar_quantity_bottom_dir1" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "pad_rebar_quantity_top_dir2" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "pad_rebar_quantity_bottom_dir2" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "pad_clear_cover" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "total_soil_unit_weight" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "bearing_type" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "nominal_bearing_capacity" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "cohesion" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "friction_angle" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "spt_blow_count" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "base_friction_factor" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "neglect_depth" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "bearing_distribution_type" Then
    '        myvar = itm.rangeValue
    '    ElseIf itm.rangeName = "groundwater_depth" Then
    '        myvar = itm.rangeValue
    '    End If
    'Next

    'End Function
#End Region

End Class