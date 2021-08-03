Option Strict Off

Imports DevExpress.Spreadsheet
Imports System.Security.Principal

Partial Public Class DataTransfererPile

#Region "Define"
    Private NewPileWb As New Workbook
    Private prop_ExcelFilePath As String

    Public Property Piles As New List(Of Pile)
    Private Property PileTemplatePath As String = "C:\Users\" & Environment.UserName & "\Documents\.NET Testing\Foundations\Pile\Template\Pile Foundation (2.2.1).xlsm"
    Private Property PileFileType As DocumentFormat = DocumentFormat.Xlsm

    Public Property pileDS As New DataSet
    Public Property pileDB As String
    Public Property pileID As WindowsIdentity
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
    Sub New()
        'Leave method empty
    End Sub

    Public Sub New(ByVal MyDataSet As DataSet, ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String, ByVal BU As String, ByVal Strucutre_ID As String)
        pileDS = MyDataSet
        pileID = LogOnUser
        pileDB = ActiveDatabase
        'BUNumber = BU 'Need to turn back on when connecting to dashboard. Turned off for testing. 
        'STR_ID = Strucutre_ID 'Need to turn back on when connecting to dashboard. Turned off for testing. 
    End Sub
#End Region

#Region "Load Data"
    Public Function LoadFromEDS() As Boolean
        Dim refid As Integer
        Dim PileLoader As String

        'Load data to get Unit Base details for the existing structure model
        For Each item As SQLParameter In PileSQLDataTables()
            PileLoader = QueryBuilderFromFile(queryPath & "Pile\" & item.sqlQuery).Replace("[EXISTING MODEL]", GetExistingModelQuery())
            DoDaSQL.sqlLoader(PileLoader, item.sqlDatatable, pileDS, pileDB, pileID, "0")
            If pileDS.Tables(item.sqlDatatable).Rows.Count = 0 Then Return False 'This may need adjusted since some tables can be empty
        Next

        'Custom Section to transfer data for the pier and pad tool. Needs to be adjusted for each tool.
        ''MRP 7/20/21 - Defined values may default to Nothing from CCI Engineering Templates. This section sets values with database entries that are not NULL for each object in the list
        'Dim n As Integer
        'n = 0 'initial object in list

        For Each PileDataRow As DataRow In pileDS.Tables("Pile General Details SQL").Rows
            refid = CType(PileDataRow.Item("pile_id"), Integer)

            Piles.Add(New Pile(PileDataRow, refid))
        Next

        Return True
    End Function 'Create Pile objects based on what is saved in EDS

    Public Sub LoadFromExcel()
        Piles.Add(New Pile(ExcelFilePath))
    End Sub 'Create Pile objects based on what is coming from the excel file
#End Region

#Region "Save Data"
    Public Sub SaveToEDS()
        Dim firstOne As Boolean = True


        For Each pf As Pile In Piles
            Dim PileSaver As String = QueryBuilderFromFile(queryPath & "Pile\Pile (IN_UP).sql")

            PileSaver = PileSaver.Replace("[BU NUMBER]", BUNumber)
            PileSaver = PileSaver.Replace("[STRUCTURE ID]", STR_ID)
            PileSaver = PileSaver.Replace("[FOUNDATION TYPE]", "Pile")
            If pf.pile_id = 0 Or IsDBNull(pf.pile_id) Then
                PileSaver = PileSaver.Replace("'[Pile ID]'", "NULL")
            Else
                PileSaver = PileSaver.Replace("[Pile ID]", pf.pile_id.ToString)
                PileSaver = PileSaver.Replace("(SELECT * FROM TEMPORARY)", UpdatePileDetail(pf))
            End If
            PileSaver = PileSaver.Replace("[INSERT ALL PILE DETAILS]", InsertPileDetail(pf))

            sqlSender(PileSaver, pileDB, pileID, "0")
        Next


    End Sub

    Public Sub SaveToExcel()
        Dim pfRow As Integer = 3
        LoadNewPile()

        With NewPileWb
            For Each pf As Pile In Piles
                If Not IsNothing(pf.pile_id) Then
                    .Worksheets("Input").Range("ID").Value = CType(pf.pile_id, Integer)
                Else .Worksheets("Input").Range("ID").ClearContents
                End If
                If Not IsNothing(pf.load_eccentricity) Then
                    .Worksheets("Input").Range("Ecc").Value = CType(pf.load_eccentricity, Double)
                Else .Worksheets("Input").Range("Ecc").ClearContents
                End If
                If Not IsNothing(pf.bolt_circle_bearing_plate_width) Then
                    .Worksheets("Input").Range("BC").Value = CType(pf.bolt_circle_bearing_plate_width, Double)
                Else .Worksheets("Input").Range("BC").ClearContents
                End If
                If Not IsNothing(pf.pile_shape) Then .Worksheets("Input").Range("D23").Value = pf.pile_shape
                If Not IsNothing(pf.pile_material) Then .Worksheets("Input").Range("D24").Value = pf.pile_material
                If Not IsNothing(pf.pile_length) Then
                    .Worksheets("Input").Range("Lpile").Value = CType(pf.pile_length, Double)
                Else .Worksheets("Input").Range("Lpile").ClearContents
                End If
                If Not IsNothing(pf.pile_diameter_width) Then
                    .Worksheets("Input").Range("D26").Value = CType(pf.pile_diameter_width, Double)
                Else .Worksheets("Input").Range("D26").ClearContents
                End If
                If Not IsNothing(pf.pile_pipe_thickness) Then
                    .Worksheets("Input").Range("D27").Value = CType(pf.pile_pipe_thickness, Double)
                Else .Worksheets("Input").Range("D27").ClearContents
                End If

                If pf.pile_soil_capacity_given = True Then
                    .Worksheets("Input").Range("D29").Value = "Yes"
                Else
                    .Worksheets("Input").Range("D29").Value = "No"
                End If

                If Not IsNothing(pf.steel_yield_strength) Then
                    .Worksheets("Input").Range("D30").Value = CType(pf.steel_yield_strength, Double)
                Else .Worksheets("Input").Range("D30").ClearContents
                End If
                If Not IsNothing(pf.pile_type_option) Then .Worksheets("Input").Range("Psize").Value = pf.pile_type_option
                If Not IsNothing(pf.rebar_quantity) Then
                    .Worksheets("Input").Range("Pquan").Value = CType(pf.rebar_quantity, Integer)
                Else .Worksheets("Input").Range("Pquan").ClearContents
                End If
                If Not IsNothing(pf.pile_group_config) Then .Worksheets("Input").Range("Config").Value = pf.pile_group_config
                If Not IsNothing(pf.foundation_depth) Then
                    .Worksheets("Input").Range("D").Value = CType(pf.foundation_depth, Double)
                Else .Worksheets("Input").Range("D").ClearContents
                End If
                If Not IsNothing(pf.pad_thickness) Then
                    .Worksheets("Input").Range("T").Value = CType(pf.pad_thickness, Double)
                Else .Worksheets("Input").Range("T").ClearContents
                End If
                If Not IsNothing(pf.pad_width_dir1) Then
                    .Worksheets("Input").Range("Wx").Value = CType(pf.pad_width_dir1, Double)
                Else .Worksheets("Input").Range("Wx").ClearContents
                End If
                If Not IsNothing(pf.pad_width_dir2) Then
                    .Worksheets("Input").Range("Wy").Value = CType(pf.pad_width_dir2, Double)
                Else .Worksheets("Input").Range("Wy").ClearContents
                End If
                If Not IsNothing(pf.pad_rebar_size_bottom) Then
                    .Worksheets("Input").Range("Spad").Value = CType(pf.pad_rebar_size_bottom, Integer)
                Else .Worksheets("Input").Range("Spad").ClearContents
                End If
                If Not IsNothing(pf.pad_rebar_size_top) Then
                    .Worksheets("Input").Range("Spad_top").Value = CType(pf.pad_rebar_size_top, Integer)
                Else .Worksheets("Input").Range("Spad_top").ClearContents
                End If
                If Not IsNothing(pf.pad_rebar_quantity_bottom_dir1) Then
                    .Worksheets("Input").Range("Mpad").Value = CType(pf.pad_rebar_quantity_bottom_dir1, Integer)
                Else .Worksheets("Input").Range("Mpad").ClearContents
                End If
                If Not IsNothing(pf.pad_rebar_quantity_top_dir1) Then
                    .Worksheets("Input").Range("Mpad_top").Value = CType(pf.pad_rebar_quantity_top_dir1, Integer)
                Else .Worksheets("Input").Range("Mpad_top").ClearContents
                End If
                If Not IsNothing(pf.pad_rebar_quantity_bottom_dir2) Then
                    .Worksheets("Input").Range("Mpad_y").Value = CType(pf.pad_rebar_quantity_bottom_dir2, Integer)
                Else .Worksheets("Input").Range("Mpad_y").ClearContents
                End If
                If Not IsNothing(pf.pad_rebar_quantity_top_dir2) Then
                    .Worksheets("Input").Range("Mpad_y_top").Value = CType(pf.pad_rebar_quantity_top_dir2, Integer)
                Else .Worksheets("Input").Range("Mpad_y_top").ClearContents
                End If
                If Not IsNothing(pf.pier_shape) Then .Worksheets("Input").Range("D57").Value = pf.pier_shape
                If Not IsNothing(pf.pier_diameter) Then
                    .Worksheets("Input").Range("di").Value = CType(pf.pier_diameter, Integer)
                Else .Worksheets("Input").Range("di").ClearContents
                End If
                If Not IsNothing(pf.extension_above_grade) Then
                    .Worksheets("Input").Range("E").Value = CType(pf.extension_above_grade, Double)
                Else .Worksheets("Input").Range("E").ClearContents
                End If
                If Not IsNothing(pf.pier_rebar_size) Then
                    .Worksheets("Input").Range("Rs").Value = CType(pf.pier_rebar_size, Integer)
                Else .Worksheets("Input").Range("Rs").ClearContents
                End If
                If Not IsNothing(pf.pier_rebar_quantity) Then
                    .Worksheets("Input").Range("mc").Value = CType(pf.pier_rebar_quantity, Integer)
                Else .Worksheets("Input").Range("mc").ClearContents
                End If
                If Not IsNothing(pf.pier_tie_size) Then
                    .Worksheets("Input").Range("St").Value = CType(pf.pier_tie_size, Integer)
                Else .Worksheets("Input").Range("St").ClearContents
                End If
                'If Not IsNothing(pf.pier_tie_quantity) Then
                '    .Worksheets("").Range("").Value = CType(pf.pier_tie_quantity, Integer)
                'Else .Worksheets("").Range("").ClearContents
                'End If
                If Not IsNothing(pf.rebar_grade) Then
                    .Worksheets("Input").Range("Fy").Value = CType(pf.rebar_grade, Double)
                Else .Worksheets("Input").Range("Fy").ClearContents
                End If
                If Not IsNothing(pf.concrete_compressive_strength) Then
                    .Worksheets("Input").Range("Fc").Value = CType(pf.concrete_compressive_strength, Double)
                Else .Worksheets("Input").Range("Fc").ClearContents
                End If
                If Not IsNothing(pf.groundwater_depth) Then
                    .Worksheets("Input").Range("D69").Value = CType(pf.groundwater_depth, Double)
                Else .Worksheets("Input").Range("D69").ClearContents
                End If
                If Not IsNothing(pf.total_soil_unit_weight) Then
                    .Worksheets("Input").Range("γsoil_dry").Value = CType(pf.total_soil_unit_weight, Double)
                Else .Worksheets("Input").Range("γsoil_dry").ClearContents
                End If
                If Not IsNothing(pf.cohesion) Then
                    .Worksheets("Input").Range("Co").Value = CType(pf.cohesion, Double)
                Else .Worksheets("Input").Range("Co").ClearContents
                End If
                If Not IsNothing(pf.friction_angle) Then
                    .Worksheets("Input").Range("ɸ").Value = CType(pf.friction_angle, Double)
                Else .Worksheets("Input").Range("ɸ").ClearContents
                End If
                If Not IsNothing(pf.neglect_depth) Then
                    .Worksheets("Input").Range("ND").Value = CType(pf.neglect_depth, Double)
                Else .Worksheets("Input").Range("ND").ClearContents
                End If
                If Not IsNothing(pf.spt_blow_count) Then
                    .Worksheets("Input").Range("N_blows").Value = CType(pf.spt_blow_count, Integer)
                Else .Worksheets("Input").Range("N_blows").ClearContents
                End If
                If Not IsNothing(pf.pile_negative_friction_force) Then
                    .Worksheets("Input").Range("Sw").Value = CType(pf.pile_negative_friction_force, Double)
                Else .Worksheets("Input").Range("Sw").ClearContents
                End If
                If Not IsNothing(pf.pile_ultimate_compression) Then
                    .Worksheets("Input").Range("K45").Value = CType(pf.pile_ultimate_compression, Double)
                Else .Worksheets("Input").Range("K45").ClearContents
                End If
                If Not IsNothing(pf.pile_ultimate_tension) Then
                    .Worksheets("Input").Range("K46").Value = CType(pf.pile_ultimate_tension, Double)
                Else .Worksheets("Input").Range("K46").ClearContents
                End If
                If Not IsNothing(pf.top_and_bottom_rebar_different) Then .Worksheets("Input").Range("Z10").Value = pf.top_and_bottom_rebar_different
                If Not IsNothing(pf.ultimate_gross_end_bearing) Then
                    .Worksheets("Input").Range("M71").Value = CType(pf.ultimate_gross_end_bearing, Double)
                Else .Worksheets("Input").Range("M71").ClearContents
                End If

                If pf.skin_friction_given = True Then
                    .Worksheets("Input").Range("N54").Value = "Yes"
                Else
                    .Worksheets("Input").Range("N54").Value = "No"
                End If

                If pf.pile_group_config = "Circular" Then
                    If Not IsNothing(pf.pile_quantity_circular) Then
                        .Worksheets("Input").Range("D36").Value = CType(pf.pile_quantity_circular, Integer)
                    Else .Worksheets("Input").Range("D36").ClearContents
                    End If
                    If Not IsNothing(pf.group_diameter_circular) Then
                        .Worksheets("Input").Range("D37").Value = CType(pf.group_diameter_circular, Double)
                    Else .Worksheets("Input").Range("D37").ClearContents
                    End If
                End If

                If pf.pile_group_config = "Rectangular" Then
                    If Not IsNothing(pf.pile_column_quantity) Then
                        .Worksheets("Input").Range("D36").Value = CType(pf.pile_column_quantity, Integer)
                    Else .Worksheets("Input").Range("D36").ClearContents
                    End If
                    If Not IsNothing(pf.pile_row_quantity) Then
                        .Worksheets("Input").Range("D37").Value = CType(pf.pile_row_quantity, Integer)
                    Else .Worksheets("Input").Range("D37").ClearContents
                    End If
                End If

                If Not IsNothing(pf.pile_columns_spacing) Then
                    .Worksheets("Input").Range("D38").Value = CType(pf.pile_columns_spacing, Double)
                Else .Worksheets("Input").Range("D38").ClearContents
                End If
                If Not IsNothing(pf.pile_row_spacing) Then
                    .Worksheets("Input").Range("D39").Value = CType(pf.pile_row_spacing, Double)
                Else .Worksheets("Input").Range("D39").ClearContents
                End If

                If pf.group_efficiency_factor_given = True Then
                    .Worksheets("Input").Range("D41").Value = "Yes"
                Else
                    .Worksheets("Input").Range("D41").Value = "No"
                End If

                If Not IsNothing(pf.group_efficiency_factor) Then
                    .Worksheets("Input").Range("D42").Value = CType(pf.group_efficiency_factor, Double)
                Else .Worksheets("Input").Range("D42").ClearContents
                End If
                If Not IsNothing(pf.cap_type) Then .Worksheets("Input").Range("D45").Value = pf.cap_type
                If Not IsNothing(pf.pile_quantity_asymmetric) Then
                    .Worksheets("Moment of Inertia").Range("D10").Value = CType(pf.pile_quantity_asymmetric, Integer)
                Else .Worksheets("Moment of Inertia").Range("D10").ClearContents
                End If
                If Not IsNothing(pf.pile_spacing_min_asymmetric) Then
                    .Worksheets("Moment of Inertia").Range("D11").Value = CType(pf.pile_spacing_min_asymmetric, Double)
                Else .Worksheets("Moment of Inertia").Range("D11").ClearContents
                End If
                If Not IsNothing(pf.quantity_piles_surrounding) Then
                    .Worksheets("Moment of Inertia").Range("D12").Value = CType(pf.quantity_piles_surrounding, Integer)
                Else .Worksheets("Moment of Inertia").Range("D12").ClearContents
                End If
            Next

        End With

        SaveAndClosePile()
    End Sub

    Private Sub LoadNewPile()
        NewPileWb.LoadDocument(PileTemplatePath, PileFileType)
        NewPileWb.BeginUpdate()
    End Sub

    Private Sub SaveAndClosePile()
        NewPileWb.EndUpdate()
        NewPileWb.SaveDocument(ExcelFilePath, PileFileType)
    End Sub
#End Region

#Region "SQL Insert Statements"
    Private Function InsertPileDetail(ByVal pf As Pile) As String
        Dim insertString As String = ""

        insertString += "@FndID"
        'insertString += "," & IIf(IsNothing(pf.pile_id), "Null", pf.pile_id.ToString)
        insertString += "," & IIf(IsNothing(pf.load_eccentricity), "Null", pf.load_eccentricity.ToString)
        insertString += "," & IIf(IsNothing(pf.bolt_circle_bearing_plate_width), "Null", pf.bolt_circle_bearing_plate_width.ToString)
        insertString += "," & IIf(IsNothing(pf.pile_shape), "Null", "'" & pf.pile_shape.ToString & "'")
        insertString += "," & IIf(IsNothing(pf.pile_material), "Null", "'" & pf.pile_material.ToString & "'")
        insertString += "," & IIf(IsNothing(pf.pile_length), "Null", pf.pile_length.ToString)
        insertString += "," & IIf(IsNothing(pf.pile_diameter_width), "Null", pf.pile_diameter_width.ToString)
        insertString += "," & IIf(IsNothing(pf.pile_pipe_thickness), "Null", pf.pile_pipe_thickness.ToString)
        insertString += "," & IIf(IsNothing(pf.pile_soil_capacity_given), "Null", "'" & pf.pile_soil_capacity_given.ToString & "'")
        insertString += "," & IIf(IsNothing(pf.steel_yield_strength), "Null", pf.steel_yield_strength.ToString)
        insertString += "," & IIf(IsNothing(pf.pile_type_option), "Null", "'" & pf.pile_type_option.ToString & "'")
        insertString += "," & IIf(IsNothing(pf.rebar_quantity), "Null", pf.rebar_quantity.ToString)
        insertString += "," & IIf(IsNothing(pf.pile_group_config), "Null", "'" & pf.pile_group_config.ToString & "'")
        insertString += "," & IIf(IsNothing(pf.foundation_depth), "Null", pf.foundation_depth.ToString)
        insertString += "," & IIf(IsNothing(pf.pad_thickness), "Null", pf.pad_thickness.ToString)
        insertString += "," & IIf(IsNothing(pf.pad_width_dir1), "Null", pf.pad_width_dir1.ToString)
        insertString += "," & IIf(IsNothing(pf.pad_width_dir2), "Null", pf.pad_width_dir2.ToString)
        insertString += "," & IIf(IsNothing(pf.pad_rebar_size_bottom), "Null", pf.pad_rebar_size_bottom.ToString)
        insertString += "," & IIf(IsNothing(pf.pad_rebar_size_top), "Null", pf.pad_rebar_size_top.ToString)
        insertString += "," & IIf(IsNothing(pf.pad_rebar_quantity_bottom_dir1), "Null", pf.pad_rebar_quantity_bottom_dir1.ToString)
        insertString += "," & IIf(IsNothing(pf.pad_rebar_quantity_top_dir1), "Null", pf.pad_rebar_quantity_top_dir1.ToString)
        insertString += "," & IIf(IsNothing(pf.pad_rebar_quantity_bottom_dir2), "Null", pf.pad_rebar_quantity_bottom_dir2.ToString)
        insertString += "," & IIf(IsNothing(pf.pad_rebar_quantity_top_dir2), "Null", pf.pad_rebar_quantity_top_dir2.ToString)
        insertString += "," & IIf(IsNothing(pf.pier_shape), "Null", "'" & pf.pier_shape.ToString & "'")
        insertString += "," & IIf(IsNothing(pf.pier_diameter), "Null", pf.pier_diameter.ToString)
        insertString += "," & IIf(IsNothing(pf.extension_above_grade), "Null", pf.extension_above_grade.ToString)
        insertString += "," & IIf(IsNothing(pf.pier_rebar_size), "Null", pf.pier_rebar_size.ToString)
        insertString += "," & IIf(IsNothing(pf.pier_rebar_quantity), "Null", pf.pier_rebar_quantity.ToString)
        insertString += "," & IIf(IsNothing(pf.pier_tie_size), "Null", pf.pier_tie_size.ToString)
        'insertString += "," & IIf(IsNothing(pf.pier_tie_quantity), "Null", pf.pier_tie_quantity.ToString)
        insertString += "," & IIf(IsNothing(pf.rebar_grade), "Null", pf.rebar_grade.ToString)
        insertString += "," & IIf(IsNothing(pf.concrete_compressive_strength), "Null", pf.concrete_compressive_strength.ToString)
        insertString += "," & IIf(IsNothing(pf.groundwater_depth), "Null", pf.groundwater_depth.ToString)
        insertString += "," & IIf(IsNothing(pf.total_soil_unit_weight), "Null", pf.total_soil_unit_weight.ToString)
        insertString += "," & IIf(IsNothing(pf.cohesion), "Null", pf.cohesion.ToString)
        insertString += "," & IIf(IsNothing(pf.friction_angle), "Null", pf.friction_angle.ToString)
        insertString += "," & IIf(IsNothing(pf.neglect_depth), "Null", pf.neglect_depth.ToString)
        insertString += "," & IIf(IsNothing(pf.spt_blow_count), "Null", pf.spt_blow_count.ToString)
        insertString += "," & IIf(IsNothing(pf.pile_negative_friction_force), "Null", pf.pile_negative_friction_force.ToString)
        insertString += "," & IIf(IsNothing(pf.pile_ultimate_compression), "Null", pf.pile_ultimate_compression.ToString)
        insertString += "," & IIf(IsNothing(pf.pile_ultimate_tension), "Null", pf.pile_ultimate_tension.ToString)
        insertString += "," & IIf(IsNothing(pf.top_and_bottom_rebar_different), "Null", "'" & pf.top_and_bottom_rebar_different.ToString & "'")
        insertString += "," & IIf(IsNothing(pf.ultimate_gross_end_bearing), "Null", pf.ultimate_gross_end_bearing.ToString)
        insertString += "," & IIf(IsNothing(pf.skin_friction_given), "Null", "'" & pf.skin_friction_given.ToString & "'")
        insertString += "," & IIf(IsNothing(pf.pile_quantity_circular), "Null", pf.pile_quantity_circular.ToString)
        insertString += "," & IIf(IsNothing(pf.group_diameter_circular), "Null", pf.group_diameter_circular.ToString)
        insertString += "," & IIf(IsNothing(pf.pile_column_quantity), "Null", pf.pile_column_quantity.ToString)
        insertString += "," & IIf(IsNothing(pf.pile_row_quantity), "Null", pf.pile_row_quantity.ToString)
        insertString += "," & IIf(IsNothing(pf.pile_columns_spacing), "Null", pf.pile_columns_spacing.ToString)
        insertString += "," & IIf(IsNothing(pf.pile_row_spacing), "Null", pf.pile_row_spacing.ToString)
        insertString += "," & IIf(IsNothing(pf.group_efficiency_factor_given), "Null", "'" & pf.group_efficiency_factor_given.ToString & "'")
        insertString += "," & IIf(IsNothing(pf.group_efficiency_factor), "Null", pf.group_efficiency_factor.ToString)
        insertString += "," & IIf(IsNothing(pf.cap_type), "Null", "'" & pf.cap_type.ToString & "'")
        insertString += "," & IIf(IsNothing(pf.pile_quantity_asymmetric), "Null", pf.pile_quantity_asymmetric.ToString)
        insertString += "," & IIf(IsNothing(pf.pile_spacing_min_asymmetric), "Null", pf.pile_spacing_min_asymmetric.ToString)
        insertString += "," & IIf(IsNothing(pf.quantity_piles_surrounding), "Null", pf.quantity_piles_surrounding.ToString)

        Return insertString
    End Function
#End Region

#Region "SQL Update Statements"
    Private Function UpdatePileDetail(ByVal pf As Pile) As String
        Dim updateString As String = ""

        updateString += "UPDATE Pile_details SET "
        'updateString += ", pile_id=" & IIf(IsNothing(pf.pile_id), "Null", pf.pile_id.ToString)
        updateString += " load_eccentricity=" & IIf(IsNothing(pf.load_eccentricity), "Null", pf.load_eccentricity.ToString)
        updateString += ", bolt_circle_bearing_plate_width=" & IIf(IsNothing(pf.bolt_circle_bearing_plate_width), "Null", pf.bolt_circle_bearing_plate_width.ToString)
        updateString += ", pile_shape=" & IIf(IsNothing(pf.pile_shape), "Null", "'" & pf.pile_shape.ToString & "'")
        updateString += ", pile_material=" & IIf(IsNothing(pf.pile_material), "Null", "'" & pf.pile_material.ToString & "'")
        updateString += ", pile_length=" & IIf(IsNothing(pf.pile_length), "Null", pf.pile_length.ToString)
        updateString += ", pile_diameter_width=" & IIf(IsNothing(pf.pile_diameter_width), "Null", pf.pile_diameter_width.ToString)
        updateString += ", pile_pipe_thickness=" & IIf(IsNothing(pf.pile_pipe_thickness), "Null", pf.pile_pipe_thickness.ToString)
        updateString += ", pile_soil_capacity_given=" & IIf(IsNothing(pf.pile_soil_capacity_given), "Null", "'" & pf.pile_soil_capacity_given.ToString & "'")
        updateString += ", steel_yield_strength=" & IIf(IsNothing(pf.steel_yield_strength), "Null", pf.steel_yield_strength.ToString)
        updateString += ", pile_type_option=" & IIf(IsNothing(pf.pile_type_option), "Null", "'" & pf.pile_type_option.ToString & "'")
        updateString += ", rebar_quantity=" & IIf(IsNothing(pf.rebar_quantity), "Null", pf.rebar_quantity.ToString)
        updateString += ", pile_group_config=" & IIf(IsNothing(pf.pile_group_config), "Null", "'" & pf.pile_group_config.ToString & "'")
        updateString += ", foundation_depth=" & IIf(IsNothing(pf.foundation_depth), "Null", pf.foundation_depth.ToString)
        updateString += ", pad_thickness=" & IIf(IsNothing(pf.pad_thickness), "Null", pf.pad_thickness.ToString)
        updateString += ", pad_width_dir1=" & IIf(IsNothing(pf.pad_width_dir1), "Null", pf.pad_width_dir1.ToString)
        updateString += ", pad_width_dir2=" & IIf(IsNothing(pf.pad_width_dir2), "Null", pf.pad_width_dir2.ToString)
        updateString += ", pad_rebar_size_bottom=" & IIf(IsNothing(pf.pad_rebar_size_bottom), "Null", pf.pad_rebar_size_bottom.ToString)
        updateString += ", pad_rebar_size_top=" & IIf(IsNothing(pf.pad_rebar_size_top), "Null", pf.pad_rebar_size_top.ToString)
        updateString += ", pad_rebar_quantity_bottom_dir1=" & IIf(IsNothing(pf.pad_rebar_quantity_bottom_dir1), "Null", pf.pad_rebar_quantity_bottom_dir1.ToString)
        updateString += ", pad_rebar_quantity_top_dir1=" & IIf(IsNothing(pf.pad_rebar_quantity_top_dir1), "Null", pf.pad_rebar_quantity_top_dir1.ToString)
        updateString += ", pad_rebar_quantity_bottom_dir2=" & IIf(IsNothing(pf.pad_rebar_quantity_bottom_dir2), "Null", pf.pad_rebar_quantity_bottom_dir2.ToString)
        updateString += ", pad_rebar_quantity_top_dir2=" & IIf(IsNothing(pf.pad_rebar_quantity_top_dir2), "Null", pf.pad_rebar_quantity_top_dir2.ToString)
        updateString += ", pier_shape=" & IIf(IsNothing(pf.pier_shape), "Null", "'" & pf.pier_shape.ToString & "'")
        updateString += ", pier_diameter=" & IIf(IsNothing(pf.pier_diameter), "Null", pf.pier_diameter.ToString)
        updateString += ", extension_above_grade=" & IIf(IsNothing(pf.extension_above_grade), "Null", pf.extension_above_grade.ToString)
        updateString += ", pier_rebar_size=" & IIf(IsNothing(pf.pier_rebar_size), "Null", pf.pier_rebar_size.ToString)
        updateString += ", pier_rebar_quantity=" & IIf(IsNothing(pf.pier_rebar_quantity), "Null", pf.pier_rebar_quantity.ToString)
        updateString += ", pier_tie_size=" & IIf(IsNothing(pf.pier_tie_size), "Null", pf.pier_tie_size.ToString)
        'updateString += ", pier_tie_quantity=" & IIf(IsNothing(pf.pier_tie_quantity), "Null", pf.pier_tie_quantity.ToString)
        updateString += ", rebar_grade=" & IIf(IsNothing(pf.rebar_grade), "Null", pf.rebar_grade.ToString)
        updateString += ", concrete_compressive_strength=" & IIf(IsNothing(pf.concrete_compressive_strength), "Null", pf.concrete_compressive_strength.ToString)
        updateString += ", groundwater_depth=" & IIf(IsNothing(pf.groundwater_depth), "Null", pf.groundwater_depth.ToString)
        updateString += ", total_soil_unit_weight=" & IIf(IsNothing(pf.total_soil_unit_weight), "Null", pf.total_soil_unit_weight.ToString)
        updateString += ", cohesion=" & IIf(IsNothing(pf.cohesion), "Null", pf.cohesion.ToString)
        updateString += ", friction_angle=" & IIf(IsNothing(pf.friction_angle), "Null", pf.friction_angle.ToString)
        updateString += ", neglect_depth=" & IIf(IsNothing(pf.neglect_depth), "Null", pf.neglect_depth.ToString)
        updateString += ", spt_blow_count=" & IIf(IsNothing(pf.spt_blow_count), "Null", pf.spt_blow_count.ToString)
        updateString += ", pile_negative_friction_force=" & IIf(IsNothing(pf.pile_negative_friction_force), "Null", pf.pile_negative_friction_force.ToString)
        updateString += ", pile_ultimate_compression=" & IIf(IsNothing(pf.pile_ultimate_compression), "Null", pf.pile_ultimate_compression.ToString)
        updateString += ", pile_ultimate_tension=" & IIf(IsNothing(pf.pile_ultimate_tension), "Null", pf.pile_ultimate_tension.ToString)
        updateString += ", top_and_bottom_rebar_different=" & IIf(IsNothing(pf.top_and_bottom_rebar_different), "Null", "'" & pf.top_and_bottom_rebar_different.ToString & "'")
        updateString += ", ultimate_gross_end_bearing=" & IIf(IsNothing(pf.ultimate_gross_end_bearing), "Null", pf.ultimate_gross_end_bearing.ToString)
        updateString += ", skin_friction_given=" & IIf(IsNothing(pf.skin_friction_given), "Null", "'" & pf.skin_friction_given.ToString & "'")
        updateString += ", pile_quantity_circular=" & IIf(IsNothing(pf.pile_quantity_circular), "Null", pf.pile_quantity_circular.ToString)
        updateString += ", group_diameter_circular=" & IIf(IsNothing(pf.group_diameter_circular), "Null", pf.group_diameter_circular.ToString)
        updateString += ", pile_column_quantity=" & IIf(IsNothing(pf.pile_column_quantity), "Null", pf.pile_column_quantity.ToString)
        updateString += ", pile_row_quantity=" & IIf(IsNothing(pf.pile_row_quantity), "Null", pf.pile_row_quantity.ToString)
        updateString += ", pile_columns_spacing=" & IIf(IsNothing(pf.pile_columns_spacing), "Null", pf.pile_columns_spacing.ToString)
        updateString += ", pile_row_spacing=" & IIf(IsNothing(pf.pile_row_spacing), "Null", pf.pile_row_spacing.ToString)
        updateString += ", group_efficiency_factor_given=" & IIf(IsNothing(pf.group_efficiency_factor_given), "Null", "'" & pf.group_efficiency_factor_given.ToString & "'")
        updateString += ", group_efficiency_factor=" & IIf(IsNothing(pf.group_efficiency_factor), "Null", pf.group_efficiency_factor.ToString)
        updateString += ", cap_type=" & IIf(IsNothing(pf.cap_type), "Null", "'" & pf.cap_type.ToString & "'")
        updateString += ", pile_quantity_asymmetric=" & IIf(IsNothing(pf.pile_quantity_asymmetric), "Null", pf.pile_quantity_asymmetric.ToString)
        updateString += ", pile_spacing_min_asymmetric=" & IIf(IsNothing(pf.pile_spacing_min_asymmetric), "Null", pf.pile_spacing_min_asymmetric.ToString)
        updateString += ", quantity_piles_surrounding=" & IIf(IsNothing(pf.quantity_piles_surrounding), "Null", pf.quantity_piles_surrounding.ToString)
        updateString += " WHERE ID = " & pf.pile_id.ToString

        Return updateString
    End Function
#End Region

#Region "General"
    Public Sub Clear()
        ExcelFilePath = ""
        Piles.Clear()
    End Sub

    Private Function PileSQLDataTables() As List(Of SQLParameter)
        Dim MyParameters As New List(Of SQLParameter)

        MyParameters.Add(New SQLParameter("Pile General Details SQL", "Pile (SELECT Details).sql"))

        Return MyParameters
    End Function

#End Region


End Class
