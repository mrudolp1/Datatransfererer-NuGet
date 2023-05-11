Option Strict Off
Option Compare Binary

Imports System.ComponentModel
Imports DevExpress.Spreadsheet
Imports System.Runtime.Serialization

<DataContractAttribute()>
Partial Public Class AnchorBlockFoundation
    Inherits EDSExcelObject

#Region "Define"
    Private _file_ver As String
    Private _modified As Boolean?

     <DataMember()> Public Property AnchorBlocks As New List(Of AnchorBlock)

    <Category("Guy Anchor Block Tool"), Description(""), DisplayName("File Ver")>
     <DataMember()> Public Property file_ver() As String
        Get
            Return Me._file_ver
        End Get
        Set
            Me._file_ver = Value
        End Set
    End Property
    <Category("Guy Anchor Block Tool"), Description(""), DisplayName("Modified")>
     <DataMember()> Public Property modified() As Boolean?
        Get
            Return Me._modified
        End Get
        Set
            Me._modified = Value
        End Set
    End Property
#End Region

#Region "Constructors"
    Public Sub New()
    End Sub

    Private Sub ConstructMe(ByVal dr As DataRow)
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        Me.file_ver = DBtoStr(dr.Item("file_ver"))
        Me.modified = DBtoStr(dr.Item("modified"))
    End Sub


    'EDS CONSTRUCTOR
    Public Sub New(ByVal strDS As DataSet, Optional ByVal Parent As EDSObject = Nothing, Optional ByVal dr As DataRow = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Dim ab As New AnchorBlock

        ConstructMe(dr)

        For Each abDr As DataRow In strDS.Tables(ab.EDSObjectName).Rows
            ab = New AnchorBlock(abDr)
            If Me.ID = ab.anchor_block_tool_id Then
                Me.AnchorBlocks.Add(New AnchorBlock(abDr, strDS, False, Me))
            End If
        Next
    End Sub

    'EXCEL CONSTRUCTOR
    Public Sub New(ByVal ExcelFilePath As String, Optional ByVal Parent As EDSObject = Nothing)
        Me.WorkBookPath = ExcelFilePath
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        LoadFromExcel()

        'ATTEMPT AT MAKING GENERAL LISTS *salute* WORK
        'Add a list of anchor profiles
        'Add a list of soil profiles
        'Add a list of soil layers

        'Keep relationships in addition to these lists
        'We still need to access properties appropriately (i.e. anchorblock.anchorprofile.datapoint)

        'After all objects are created repurpose the constructor for the anchor block to assign relationships
        'If local_soil_id = local_soil_id Then swipe_right Else swipe_left

        'SQL FUNCTIONS
        'Only have logic in tool objcet functionality?
        'Basic insert update delete can be kept for the anchor profile and soil profile / layer
        'Change 

        ''''''''SCENARIOS''''''''
        'Scenario 1 - New anchor profile created in tool. Needs to be referenced by existing anchor block
        ''Profile needs inserted first and that foreign key needs to be used to UPDATE the existing anchor block
        ''Anchor 6 exists
        ''Profile 3 is being added
        ''Anchor 6 no longer references profile 1 but needs to reference profile 3
        ''INSERT profile; UPDATE anchor SET profile_id = @my_new_profile_id WHERE ID = 6;
        ''Main question is how do we get @my_new_profile_id?

        'Scenario 2 - a new anchor block was added to the tool that references existing profiles
        ''Insert block with appropriate foreign key references. These can be set by just using the objects associated to it. 
    End Sub
#End Region

#Region "Inherited"
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Guy Anchor Block Tool"
        End Get
    End Property

    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "fnd.anchor_block_tool"
        End Get
    End Property

    Public Overrides ReadOnly Property TemplatePath As String
        Get
            Return IO.Path.Combine(My.Application.Info.DirectoryPath, "Templates", "Guyed Anchor Block Foundation.xlsm")
        End Get
    End Property

    Public Overrides ReadOnly Property Template As Byte()
        Get
            Return CCI_Engineering_Templates.My.Resources.Guyed_Anchor_Block_Foundation
        End Get
    End Property

    Public Overrides ReadOnly Property ExcelDTParams As List(Of EXCELDTParameter)
        Get
            Dim ab As New AnchorBlock
            Dim abProf As New AnchorBlockProfile
            Dim abSProf As New AnchorBlockSoilProfile
            Dim abSlay As New AnchorBlockSoilLayer
            Dim abRes As New AnchorBlockResult
            Dim abTool As New AnchorBlockFoundation

            Return New List(Of EXCELDTParameter) From {
                                                        New EXCELDTParameter(abTool.EDSObjectName, "A2:I3", "Tool (SAPI)"),  'It is slightly confusing but to keep naming issues consistent in the tool a drilled pier = profile and a drilled pier profile = drilled pier details
                                                        New EXCELDTParameter(ab.EDSObjectName, "A2:L52", "Anchors (SAPI)"),
                                                        New EXCELDTParameter(abProf.EDSObjectName, "A2:AB52", "Anchor Profiles (SAPI)"),
                                                        New EXCELDTParameter(abSProf.EDSObjectName, "A2:F52", "Soil Profiles (SAPI)"),
                                                        New EXCELDTParameter(abSlay.EDSObjectName, "A2:N452", "Soil Layers (SAPI)"),
                                                        New EXCELDTParameter(abRes.EDSObjectName, "A2:I152", "Anchor Results (SAPI)")
                                                                                        }
            '***Add additional table references here****
        End Get
    End Property

    Private _Insert As String
    Private _Update As String
    Private _Delete As String

    Public Overrides Function SQLInsert() As String
        Dim _abInsert As String

        SQLInsert = CCI_Engineering_Templates.My.Resources.Anchor_Block_Tool__INSERT

        SQLInsert = SQLInsert.Replace("[TABLE NAME]", Me.EDSTableName)
        SQLInsert = SQLInsert.Replace("[INSERT VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[INSERT FIELDS]", Me.SQLInsertFields)

        'Guy Anchors
        _abInsert = ""
        For Each ab In Me.AnchorBlocks
            _abInsert += ab.SQLInsert + vbCrLf
        Next

        SQLInsert = SQLInsert.Replace("--[ANCHOR BLOCKS]", _abInsert)
        SQLInsert = SQLInsert.TrimEnd

        Return SQLInsert + vbCrLf

    End Function

    Public Overrides Function SQLUpdate() As String
        SQLUpdate = CCI_Engineering_Templates.My.Resources.General__UPDATE
        SQLUpdate = SQLUpdate.Replace("[TABLE NAME]", Me.EDSTableName)
        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)

        Dim gbUp As String = ""
        For Each gab In Me.AnchorBlocks
            If gab.ID IsNot Nothing And gab?.ID > 0 Then
                If gab.local_anchor_id IsNot Nothing Then
                    gbUp += gab.SQLUpdate
                Else
                    gbUp += gab.SQLDelete
                End If
            Else
                gbUp += gab.SQLInsert
            End If
        Next

        SQLUpdate = SQLUpdate.Replace("--[OPTIONAL]", gbUp)

        Return SQLUpdate + vbCrLf

    End Function

    Public Overrides Function SQLDelete() As String

        'Not needed if SQL tables are updated to allow Cascading Delete - MRR
        Dim gbDel As String = ""
        For Each gab In Me.AnchorBlocks
            gbDel += gab.SQLDelete
        Next
        SQLDelete = (gbDel + vbCrLf)
        SQLDelete += CCI_Engineering_Templates.My.Resources.General__DELETE

        'SQLDelete = CCI_Engineering_Templates.My.Resources.General__DELETE
        SQLDelete = SQLDelete.Replace("[TABLE]", Me.EDSTableName)
        SQLDelete = SQLDelete.Replace("[ID]", Me.ID)

        Return SQLDelete + vbCrLf
    End Function

#End Region

#Region "Load From Excel"
    Public Overrides Sub LoadFromExcel()
        ''''''Customize for each foundation type'''''
        Dim excelDS As New DataSet

        For Each item As EXCELDTParameter In ExcelDTParams
            Try
                excelDS.Tables.Add(ExcelDatasourceToDataTable(GetExcelDataSource(Me.WorkBookPath, item.xlsSheet, item.xlsRange), item.xlsDatatable))
            Catch ex As Exception
                Debug.Print(String.Format("Failed to create datatable for: {0}, {1}, {2}", IO.Path.GetFileName(Me.WorkBookPath), item.xlsSheet, item.xlsRange))
            End Try
        Next

        'If tables exist then construct otherwise just return blank GAB Tool 
        '''''''''''''''''''''
        '''''''''''''''''''''
        If excelDS.Tables.Contains("Guy Anchor Block Tool") Then
            Dim dr As DataRow
            Try
                dr = excelDS.Tables("Guy Anchor Block Tool").Rows(0)
            Catch ex As Exception
            End Try

            ConstructMe(dr)

            Dim myAB As New AnchorBlock
            For Each abrow As DataRow In excelDS.Tables(myAB.EDSObjectName).Rows
                If IsSomething(abrow.Item("local_anchor_id")) Or (IsSomething(abrow.Item("ID")) And IsNothing(abrow.Item("local_anchor_id"))) Then
                    Me.AnchorBlocks.Add(New AnchorBlock(abrow, excelDS, True, Me))
                End If
            Next
        End If
        '''''''''''''''''''''
        '''''''''''''''''''''
    End Sub
#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""

        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bus_unit.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.structure_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.file_ver.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        'SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bus_unit")
        SQLInsertFields = SQLInsertFields.AddtoDBString("structure_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("file_ver")
        SQLInsertFields = SQLInsertFields.AddtoDBString("modified")
        SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""

        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bus_unit = " & Me.bus_unit.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("structure_id = " & Me.structure_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("file_ver = " & Me.file_ver.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified = " & Me.modified.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_person_id = " & Me.modified_person_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("process_stage = " & Me.process_stage.ToString.FormatDBValue)

        Return SQLUpdateFieldsandValues
    End Function
#End Region

    '#Region "Save to Excel"
    '    Public Sub workBookFiller_original(ByRef wb As Workbook)
    '        '''''Customize for each excel tool'''''
    '        'Site Code Criteria
    '        Dim tia_current, site_name, structure_type As String
    '        Dim rev_h_section_15_5 As Boolean?
    '        'Dim returnRow = 2
    '        'Dim sumRow = 10 'Actually starts on row 11 but excel uses a (0, 0) origin


    '        With wb

    '            'GAB Tool
    '            Dim gab_tia_current As String
    '            Dim gab_rev_h_section_15_5 As Boolean?

    '            If Not IsNothing(Me.ID) Then
    '                .Worksheets("Tool (SAPI)").Range("A3").Value = CType(Me.ID, Integer)
    '            Else
    '                .Worksheets("Tool (SAPI)").Range("A3").ClearContents
    '            End If
    '            If Not IsNothing(Me.bus_unit) Then
    '                .Worksheets("Tool (SAPI)").Range("B3").Value = CType(Me.bus_unit, Integer)
    '            Else
    '                .Worksheets("Tool (SAPI)").Range("B3").ClearContents
    '            End If
    '            If Not IsNothing(Me.structure_id) Then
    '                .Worksheets("Tool (SAPI)").Range("C3").Value = CType(Me.structure_id, String)
    '            End If
    '            'If Not IsNothing(Me.file_ver) Then
    '            '    .Worksheets("Tool (SAPI)").Range("D3").Value = CType(Me.file_ver, String)
    '            'End If
    '            If Not IsNothing(Me.modified) Then
    '                .Worksheets("Tool (SAPI)").Range("E3").Value = CType(Me.modified, Boolean)
    '            End If
    '            'If Not IsNothing(Me.modified_person_id) Then
    '            '    .Worksheets("Tool (SAPI)").Range("F3").Value = CType(Me.modified_person_id, Integer)
    '            'Else
    '            '    .Worksheets("Tool (SAPI)").Range("F3").ClearContents
    '            'End If
    '            'If Not IsNothing(Me.process_stage) Then
    '            '    .Worksheets("Tool (SAPI)").Range("G3").Value = CType(Me.process_stage, String)
    '            'End If

    '            'TIA Revision- Defaulting to Rev. H if not available. 
    '            If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.tia_current) Then
    '                If Me.ParentStructure?.structureCodeCriteria?.tia_current = "TIA-222-F" Then
    '                    gab_tia_current = "F"
    '                ElseIf Me.ParentStructure?.structureCodeCriteria?.tia_current = "TIA-222-G" Then
    '                    gab_tia_current = "G"
    '                Else
    '                    gab_tia_current = "H"
    '                End If
    '            Else
    '                gab_tia_current = "H"
    '            End If
    '            .Worksheets("Tool (SAPI)").Range("L3").Value = CType(gab_tia_current, String)
    '            'Load Z Normalization
    '            'If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.load_z_norm) Then
    '            '    rev_h_section_15_5 = Me.ParentStructure?.structureCodeCriteria?.load_z_norm
    '            '    .Worksheets("General (SAPI)").Range("Q3").Value = CType(load_z_norm, Boolean)
    '            'End If
    '            'H Section 15.5
    '            If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.rev_h_section_15_5) Then
    '                gab_rev_h_section_15_5 = Me.ParentStructure?.structureCodeCriteria?.rev_h_section_15_5
    '                .Worksheets("Tool (SAPI)").Range("M3").Value = CType(gab_rev_h_section_15_5, Boolean)
    '            End If
    '            'Work Order
    '            If Not IsNothing(Me.ParentStructure?.work_order_seq_num) Then
    '                work_order_seq_num = Me.ParentStructure?.work_order_seq_num
    '                .Worksheets("Tool (SAPI)").Range("K3").Value = CType(work_order_seq_num, Integer)
    '            End If
    '            'Site Name
    '            If Not IsNothing(Me.ParentStructure?.SiteInfo.site_name) Then
    '                site_name = Me.ParentStructure?.SiteInfo.site_name
    '                .Worksheets("Tool (SAPI)").Range("J3").Value = CType(site_name, String)
    '            End If

    '            'Anchors
    '            Dim i As Integer = 3
    '            For Each gab As AnchorBlock In AnchorBlocks
    '                If Not IsNothing(gab.ID) Then .Worksheets("Anchors (SAPI").Range("A" & i).Value = CType(gab.ID, Integer)
    '                If Not IsNothing(gab.anchor_block_tool_id) Then .Worksheets("Anchors (SAPI").Range("B" & i).Value = CType(gab.anchor_block_tool_id, Integer)
    '                If Not IsNothing(gab.anchor_profile_id) Then .Worksheets("Anchors (SAPI").Range("C" & i).Value = CType(gab.anchor_profile_id, Integer)
    '                If Not IsNothing(gab.soil_profile_id) Then .Worksheets("Anchors (SAPI").Range("D" & i).Value = CType(gab.soil_profile_id, Integer)
    '                If Not IsNothing(gab.local_anchor_id) Then .Worksheets("Anchors (SAPI").Range("E" & i).Value = CType(gab.local_anchor_id, Integer)
    '                If Not IsNothing(gab.reaction_location) Then .Worksheets("Anchors (SAPI").Range("F" & i).Value = CType(gab.reaction_location, String)
    '                If Not IsNothing(gab.local_anchor_profile_id) Then .Worksheets("Anchors (SAPI").Range("G" & i).Value = CType(gab.local_anchor_profile_id, Integer)
    '                If Not IsNothing(gab.local_soil_profile_id) Then .Worksheets("Anchors (SAPI").Range("H" & i).Value = CType(gab.local_soil_profile_id, Integer)
    '                i += 1
    '            Next

    '            'Profile
    '            i = 3
    '            For Each ap As AnchorBlockProfile In AnchorProfiles
    '                With .Worksheets("Anchor Profiles (SAPI)")
    '                    If Not IsNothing(ap.ID) Then .Range("A" & i).Value = CType(ap.ID, Integer)
    '                    If Not IsNothing(ap.local_anchor_profile_id) Then .Range("B" & i).Value = CType(ap.local_anchor_profile_id, Integer)
    '                    If Not IsNothing(ap.anchor_depth) Then .Range("C" & i).Value = CType(ap.anchor_depth, Double)
    '                    If Not IsNothing(ap.anchor_width) Then .Range("D" & i).Value = CType(ap.anchor_width, Double)
    '                    If Not IsNothing(ap.anchor_thickness) Then .Range("E" & i).Value = CType(ap.anchor_thickness, Double)
    '                    If Not IsNothing(ap.anchor_length) Then .Range("F" & i).Value = CType(ap.anchor_length, Double)
    '                    If Not IsNothing(ap.anchor_toe_width) Then .Range("G" & i).Value = CType(ap.anchor_toe_width, Double)
    '                    If Not IsNothing(ap.anchor_top_rebar_size) Then .Range("H" & i).Value = CType(ap.anchor_top_rebar_size, Integer)
    '                    If Not IsNothing(ap.anchor_top_rebar_quantity) Then .Range("I" & i).Value = CType(ap.anchor_top_rebar_quantity, Integer)
    '                    If Not IsNothing(ap.anchor_front_rebar_size) Then .Range("J" & i).Value = CType(ap.anchor_front_rebar_size, Integer)
    '                    If Not IsNothing(ap.anchor_front_rebar_quantity) Then .Range("K" & i).Value = CType(ap.anchor_front_rebar_quantity, Integer)
    '                    If Not IsNothing(ap.anchor_stirrup_size) Then .Range("L" & i).Value = CType(ap.anchor_stirrup_size, Integer)
    '                    If Not IsNothing(ap.anchor_shaft_diameter) Then .Range("M" & i).Value = CType(ap.anchor_shaft_diameter, Double)
    '                    If Not IsNothing(ap.anchor_shaft_quantity) Then .Range("N" & i).Value = CType(ap.anchor_shaft_quantity, Integer)
    '                    If Not IsNothing(ap.anchor_shaft_area_override) Then .Range("O" & i).Value = CType(ap.anchor_shaft_area_override, Double)
    '                    If Not IsNothing(ap.anchor_shaft_shear_leg_factor) Then .Range("P" & i).Value = CType(ap.anchor_shaft_shear_leg_factor, Double)
    '                    If Not IsNothing(ap.anchor_shaft_section) Then .Range("Q" & i).Value = CType(ap.anchor_shaft_section, String)
    '                    If Not IsNothing(ap.anchor_rebar_grade) Then .Range("R" & i).Value = CType(ap.anchor_rebar_grade, Double)
    '                    If Not IsNothing(ap.concrete_compressive_strength) Then .Range("S" & i).Value = CType(ap.concrete_compressive_strength, Double)
    '                    If Not IsNothing(ap.clear_cover) Then .Range("T" & i).Value = CType(ap.clear_cover, Double)
    '                    If Not IsNothing(ap.anchor_shaft_yield_strength) Then .Range("U" & i).Value = CType(ap.anchor_shaft_yield_strength, Double)
    '                    If Not IsNothing(ap.anchor_shaft_ultimate_strength) Then .Range("V" & i).Value = CType(ap.anchor_shaft_ultimate_strength, Double)
    '                    If Not IsNothing(ap.rebar_known) Then .Range("W" & i).Value = CType(ap.rebar_known, Boolean)
    '                    If Not IsNothing(ap.anchor_shaft_known) Then .Range("X" & i).Value = CType(ap.anchor_shaft_known, Boolean)
    '                    If Not IsNothing(ap.basic_soil_check) Then .Range("Y" & i).Value = CType(ap.basic_soil_check, Boolean)
    '                    If Not IsNothing(ap.structural_check) Then .Range("Z" & i).Value = CType(ap.structural_check, Boolean)
    '                End With
    '                i += 1
    '            Next

    '            'Soil Profile
    '            i = 3
    '            Dim j As Integer = 3
    '            For Each sp As AnchorBlockSoilProfile In SoilProfiles
    '                If Not IsNothing(sp.ID) Then .Worksheets("Soil Profiles (SAPI)").Range("A" & i).Value = CType(sp.ID, Integer)
    '                If Not IsNothing(sp.local_soil_profile_id) Then .Worksheets("Soil Profiles (SAPI)").Range("B" & i).Value = CType(sp.local_soil_profile_id, Integer)
    '                If Not IsNothing(sp.groundwater_depth) Then .Worksheets("Soil Profiles (SAPI)").Range("C" & i).Value = CType(sp.groundwater_depth, Double)
    '                If Not IsNothing(sp.neglect_depth) Then .Worksheets("Soil Profiles (SAPI)").Range("D" & i).Value = CType(sp.neglect_depth, Double)

    '                'Soil Layers
    '                For Each layer In sp.SoilLayers
    '                    If Not IsNothing(layer.ID) Then .Worksheets("Soil Layers (SAPI)").Range("A" & j).Value = CType(layer.ID, Integer)
    '                    If Not IsNothing(layer.Soil_Profile_id) Then .Worksheets("Soil Layers (SAPI)").Range("B" & j).Value = CType(layer.Soil_Profile_id, Integer)
    '                    If Not IsNothing(layer.local_soil_profile_id) Then .Worksheets("Soil Layers (SAPI)").Range("C" & j).Value = CType(layer.local_soil_profile_id, Integer)
    '                    If Not IsNothing(layer.local_soil_layer_id) Then .Worksheets("Soil Layers (SAPI)").Range("D" & j).Value = CType(layer.local_soil_layer_id, Integer)
    '                    If Not IsNothing(layer.bottom_depth) Then .Worksheets("Soil Layers (SAPI)").Range("E" & j).Value = CType(layer.bottom_depth, Double)
    '                    If Not IsNothing(layer.effective_soil_density) Then .Worksheets("Soil Layers (SAPI)").Range("F" & j).Value = CType(layer.effective_soil_density, Double)
    '                    If Not IsNothing(layer.cohesion) Then .Worksheets("Soil Layers (SAPI)").Range("G" & j).Value = CType(layer.cohesion, Double)
    '                    If Not IsNothing(layer.friction_angle) Then .Worksheets("Soil Layers (SAPI)").Range("H" & j).Value = CType(layer.friction_angle, Double)
    '                    If Not IsNothing(layer.skin_friction_override_comp) Then .Worksheets("Soil Layers (SAPI)").Range("I" & j).Value = CType(layer.skin_friction_override_comp, Double)
    '                    If Not IsNothing(layer.skin_friction_override_uplift) Then .Worksheets("Soil Layers (SAPI)").Range("J" & j).Value = CType(layer.skin_friction_override_uplift, Double)
    '                    If Not IsNothing(layer.nominal_bearing_capacity) Then .Worksheets("Soil Layers (SAPI)").Range("K" & j).Value = CType(layer.nominal_bearing_capacity, Double)
    '                    If Not IsNothing(layer.spt_blow_count) Then .Worksheets("Soil Layers (SAPI)").Range("L" & j).Value = CType(layer.spt_blow_count, Double)

    '                    j += 1
    '                Next

    '                i += 1
    '            Next

    '        End With

    '    End Sub
    '#End Region

#Region "Save to Excel - IEM"
    Public Overrides Sub workBookFiller(ByRef wb As Workbook)
        '''''Customize for each excel tool'''''
        'Site Code Criteria
        Dim tia_current, site_name, structure_type As String
        Dim rev_h_section_15_5 As Boolean?
        'Dim returnRow = 2
        'Dim sumRow = 10 'Actually starts on row 11 but excel uses a (0, 0) origin


        With wb

            'GAB Tool
            Dim gab_tia_current As String
            Dim gab_rev_h_section_15_5 As Boolean?

            With .Worksheets("SUMMARY")
                .Range("EDSReactions").Value = True
            End With

            With .Worksheets("Tool (SAPI)")
                If Not IsNothing(Me.ID) Then
                    .Range("A3").Value = CType(Me.ID, Integer)
                Else
                    .Range("A3").ClearContents
                End If
                If Not IsNothing(Me.bus_unit) Then
                    .Range("B3").Value = CType(Me.bus_unit, Integer)
                Else
                    .Range("B3").ClearContents
                End If
                If Not IsNothing(Me.structure_id) Then
                    .Range("C3").Value = CType(Me.structure_id, String)
                End If
                'If Not IsNothing(Me.file_ver) Then
                '    .Range("D3").Value = CType(Me.file_ver, String)
                'End If
                If Not IsNothing(Me.modified) Then
                    .Range("E3").Value = CType(Me.modified, Boolean)
                End If
                'If Not IsNothing(Me.modified_person_id) Then
                '    .Range("F3").Value = CType(Me.modified_person_id, Integer)
                'Else
                '    .Range("F3").ClearContents
                'End If
                'If Not IsNothing(Me.process_stage) Then
                '    .Range("G3").Value = CType(Me.process_stage, String)
                'End If

                'TIA Revision- Defaulting to Rev. H if not available. 
                If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.tia_current) Then
                    If Me.ParentStructure?.structureCodeCriteria?.tia_current = "TIA-222-F" Then
                        gab_tia_current = "F"
                    ElseIf Me.ParentStructure?.structureCodeCriteria?.tia_current = "TIA-222-G" Then
                        gab_tia_current = "G"
                    Else
                        gab_tia_current = "H"
                    End If
                Else
                    gab_tia_current = "H"
                End If
                .Range("L3").Value = CType(gab_tia_current, String)
                'Load Z Normalization
                'If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.load_z_norm) Then
                '    rev_h_section_15_5 = Me.ParentStructure?.structureCodeCriteria?.load_z_norm
                '    .Worksheets("General (SAPI)").Range("Q3").Value = CType(load_z_norm, Boolean)
                'End If
                'H Section 15.5
                If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.rev_h_section_15_5) Then
                    gab_rev_h_section_15_5 = Me.ParentStructure?.structureCodeCriteria?.rev_h_section_15_5
                    .Range("M3").Value = CType(gab_rev_h_section_15_5, Boolean)
                End If
                'Work Order
                If Not IsNothing(Me.ParentStructure?.work_order_seq_num) Then
                    work_order_seq_num = Me.ParentStructure?.work_order_seq_num
                    .Range("K3").Value = CType(work_order_seq_num, Integer)
                End If
                'Site Name
                'If Not IsNothing(Me.ParentStructure?.SiteInfo.site_name) Then
                '    site_name = Me.ParentStructure?.SiteInfo.site_name
                '    .Range("J3").Value = CType(site_name, String)
                'End If
            End With

            'Anchors
            Dim anchorRow As Integer = 3
            Dim profileRow As Integer = 3
            Dim soilProfRow As Integer = 3
            Dim layerRow As Integer = 3
            Dim profExists As New List(Of Integer)
            Dim soilProfExists As New List(Of Integer)

            'Loop through all anchor blocks
            For Each gab As AnchorBlock In AnchorBlocks
                With .Worksheets("Anchors (SAPI)")
                    If Not IsNothing(gab.ID) Then .Range("A" & anchorRow).Value = CType(gab.ID, Integer)
                    If Not IsNothing(gab.anchor_block_tool_id) Then .Range("B" & anchorRow).Value = CType(gab.anchor_block_tool_id, Integer)
                    If Not IsNothing(gab.anchor_profile_id) Then .Range("C" & anchorRow).Value = CType(gab.anchor_profile_id, Integer)
                    If Not IsNothing(gab.soil_profile_id) Then .Range("D" & anchorRow).Value = CType(gab.soil_profile_id, Integer)
                    If Not IsNothing(gab.local_anchor_id) Then .Range("E" & anchorRow).Value = CType(gab.local_anchor_id, Integer)
                    If Not IsNothing(gab.reaction_location) Then .Range("F" & anchorRow).Value = CType(gab.reaction_location, String)
                    If Not IsNothing(gab.local_anchor_profile_id) Then .Range("G" & anchorRow).Value = CType(gab.local_anchor_profile_id, Integer)
                    If Not IsNothing(gab.local_soil_profile_id) Then .Range("H" & anchorRow).Value = CType(gab.local_soil_profile_id, Integer)
                End With
                'When the anchor details are added it increments the anchor row to ensure it continues down the table
                anchorRow += 1

                'Profile
                'Add the structural properties of the anchor block to the profiles tab
                Dim ap As AnchorBlockProfile = gab.AnchorProfile
                'If Not profExists.Contains(ap.local_anchor_profile_id) Then
                With .Worksheets("Anchor Profiles (SAPI)")
                    If Not IsNothing(ap.ID) Then .Range("A" & profileRow).Value = CType(ap.ID, Integer)
                    If Not IsNothing(ap.local_anchor_profile_id) Then .Range("B" & profileRow).Value = CType(ap.local_anchor_profile_id, Integer)
                    If Not IsNothing(ap.anchor_depth) Then .Range("C" & profileRow).Value = CType(ap.anchor_depth, Double)
                    If Not IsNothing(ap.anchor_width) Then .Range("D" & profileRow).Value = CType(ap.anchor_width, Double)
                    If Not IsNothing(ap.anchor_thickness) Then .Range("E" & profileRow).Value = CType(ap.anchor_thickness, Double)
                    If Not IsNothing(ap.anchor_length) Then .Range("F" & profileRow).Value = CType(ap.anchor_length, Double)
                    If Not IsNothing(ap.anchor_toe_width) Then .Range("G" & profileRow).Value = CType(ap.anchor_toe_width, Double)
                    If Not IsNothing(ap.anchor_top_rebar_size) Then .Range("H" & profileRow).Value = CType(ap.anchor_top_rebar_size, Integer)
                    If Not IsNothing(ap.anchor_top_rebar_quantity) Then .Range("I" & profileRow).Value = CType(ap.anchor_top_rebar_quantity, Integer)
                    If Not IsNothing(ap.anchor_front_rebar_size) Then .Range("J" & profileRow).Value = CType(ap.anchor_front_rebar_size, Integer)
                    If Not IsNothing(ap.anchor_front_rebar_quantity) Then .Range("K" & profileRow).Value = CType(ap.anchor_front_rebar_quantity, Integer)
                    If Not IsNothing(ap.anchor_stirrup_size) Then .Range("L" & profileRow).Value = CType(ap.anchor_stirrup_size, Integer)
                    If Not IsNothing(ap.anchor_shaft_diameter) Then .Range("M" & profileRow).Value = CType(ap.anchor_shaft_diameter, Double)
                    If Not IsNothing(ap.anchor_shaft_quantity) Then .Range("N" & profileRow).Value = CType(ap.anchor_shaft_quantity, Integer)
                    If Not IsNothing(ap.anchor_shaft_area_override) Then .Range("O" & profileRow).Value = CType(ap.anchor_shaft_area_override, Double)
                    If Not IsNothing(ap.anchor_shaft_shear_leg_factor) Then .Range("P" & profileRow).Value = CType(ap.anchor_shaft_shear_leg_factor, Double)
                    If Not IsNothing(ap.anchor_shaft_section) Then .Range("Q" & profileRow).Value = CType(ap.anchor_shaft_section, String)
                    If Not IsNothing(ap.anchor_rebar_grade) Then .Range("R" & profileRow).Value = CType(ap.anchor_rebar_grade, Double)
                    If Not IsNothing(ap.concrete_compressive_strength) Then .Range("S" & profileRow).Value = CType(ap.concrete_compressive_strength, Double)
                    If Not IsNothing(ap.clear_cover) Then .Range("T" & profileRow).Value = CType(ap.clear_cover, Double)
                    If Not IsNothing(ap.anchor_shaft_yield_strength) Then .Range("U" & profileRow).Value = CType(ap.anchor_shaft_yield_strength, Double)
                    If Not IsNothing(ap.anchor_shaft_ultimate_strength) Then .Range("V" & profileRow).Value = CType(ap.anchor_shaft_ultimate_strength, Double)
                    If Not IsNothing(ap.rebar_known) Then .Range("W" & profileRow).Value = CType(ap.rebar_known, Boolean)
                    If Not IsNothing(ap.anchor_shaft_known) Then .Range("X" & profileRow).Value = CType(ap.anchor_shaft_known, Boolean)
                    If Not IsNothing(ap.basic_soil_check) Then .Range("Y" & profileRow).Value = CType(ap.basic_soil_check, Boolean)
                    If Not IsNothing(ap.structural_check) Then .Range("Z" & profileRow).Value = CType(ap.structural_check, Boolean)
                End With
                'If it is added to the sheet then it needs to move on to the next row for the next structural design
                profileRow += 1
                'If it is added to the sheet then it adds it to a list of intergers containing local profile IDs
                'This ensures it won't be added twice.
                profExists.Add(ap.local_anchor_profile_id)
                'End If

                'Soil Profile
                Dim sp As SoilProfile = gab.SoilProfile
                'If Not soilProfExists.Contains(gab.local_soil_profile_id) Then
                With .Worksheets("Soil Profiles (SAPI)")
                    If Not IsNothing(sp.ID) Then .Range("A" & soilProfRow).Value = CType(sp.ID, Integer)
                    If Not IsNothing(gab.local_soil_profile_id) Then .Range("B" & soilProfRow).Value = CType(gab.local_soil_profile_id, Integer)
                    If Not IsNothing(sp.groundwater_depth) Then .Range("C" & soilProfRow).Value = CType(sp.groundwater_depth, Double)
                    If Not IsNothing(sp.neglect_depth) Then .Range("D" & soilProfRow).Value = CType(sp.neglect_depth, Double)
                End With
                'Increment the soil profile row to ensure it goes to the next row of the table
                soilProfRow += 1
                    'If it is added to the sheet then it adds it to a list of intergers containing local soil profile IDs
                    'This ensures it won't be added twice.
                    soilProfExists.Add(gab.local_soil_profile_id)

                    'Soil Layers
                    Dim SoilInc As Integer = 1
                    For Each layer In sp.SoilLayers
                        With .Worksheets("Soil Layers (SAPI)")
                            If Not IsNothing(layer.ID) Then .Range("A" & layerRow).Value = CType(layer.ID, Integer)
                            If Not IsNothing(layer.Soil_Profile_id) Then .Range("B" & layerRow).Value = CType(layer.Soil_Profile_id, Integer)
                            If Not IsNothing(gab.local_soil_profile_id) Then .Range("C" & layerRow).Value = CType(gab.local_soil_profile_id, Integer)
                            .Range("D" & layerRow).Value = CType(SoilInc, Integer)
                            If Not IsNothing(layer.bottom_depth) Then .Range("E" & layerRow).Value = CType(layer.bottom_depth, Double)
                            If Not IsNothing(layer.effective_soil_density) Then .Range("F" & layerRow).Value = CType(layer.effective_soil_density, Double)
                            If Not IsNothing(layer.cohesion) Then .Range("G" & layerRow).Value = CType(layer.cohesion, Double)
                            If Not IsNothing(layer.friction_angle) Then .Range("H" & layerRow).Value = CType(layer.friction_angle, Double)
                            If Not IsNothing(layer.skin_friction_override_comp) Then .Range("I" & layerRow).Value = CType(layer.skin_friction_override_comp, Double)
                            If Not IsNothing(layer.skin_friction_override_uplift) Then .Range("J" & layerRow).Value = CType(layer.skin_friction_override_uplift, Double)
                            If Not IsNothing(layer.nominal_bearing_capacity) Then .Range("K" & layerRow).Value = CType(layer.nominal_bearing_capacity, Double)
                            If Not IsNothing(layer.spt_blow_count) Then .Range("L" & layerRow).Value = CType(layer.spt_blow_count, Double)
                        End With
                        'SoilInc is used to determine the local soil layer id.
                        'This is not a "Global" variable in the sense of the soil layer tab
                        'It is isolated to soil profiles and will reset for each soil proile being added
                        'This ensures soil layers will always be labeled appropriately and doesn't need to be done in tool.
                        SoilInc += 1
                        'Increment the soil layer row for the next layer in this profile as well as the next layer in the next profile
                        layerRow += 1
                    Next
                'End If
            Next

        End With

    End Sub
#End Region

#Region "Equals"
    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        Dim otherToCompare As AnchorBlockFoundation = TryCast(other, AnchorBlockFoundation)
        If otherToCompare Is Nothing Then Return False

        'Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
        'Equals = If(Me.bus_unit.CheckChange(otherToCompare.bus_unit, changes, categoryName, "Bus Unit"), Equals, False)
        'Equals = If(Me.structure_id.CheckChange(otherToCompare.structure_id, changes, categoryName, "Structure Id"), Equals, False)
        Equals = If(Me.file_ver.CheckChange(otherToCompare.file_ver, changes, categoryName, "File Ver"), Equals, False)
        Equals = If(Me.modified.CheckChange(otherToCompare.modified, changes, categoryName, "Modified"), Equals, False)
        'Equals = If(Me.modified_person_id.CheckChange(otherToCompare.modified_person_id, changes, categoryName, "Modified Person Id"), Equals, False)
        'Equals = If(Me.process_stage.CheckChange(otherToCompare.process_stage, changes, categoryName, "Process Stage"), Equals, False)

        'Anchors
        Equals = If(Me.AnchorBlocks.CheckChange(otherToCompare.AnchorBlocks, changes, categoryName, "Anchor Blocks"), Equals, False)
    End Function
#End Region

End Class

<DataContractAttribute()>
Partial Public Class AnchorBlock
    Inherits EDSObjectWithQueries

     <DataMember()> Public Property AnchorProfile As AnchorBlockProfile
     <DataMember()> Public Property SoilProfile As AnchorBlockSoilProfile

#Region "Inheritted"
    '''Must override these inherited properties
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Guy Anchor Blocks"
        End Get
    End Property
    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "fnd.anchor_block"
        End Get
    End Property

    Private _Insert As String
    Private _Update As String
    Private _Delete As String

    Public Overrides Function SQLInsert() As String
        SQLInsert = ""
        SQLInsert = CCI_Engineering_Templates.My.Resources.Anchor_Block__INSERT

        SQLInsert = SQLInsert.Replace("[TABLE NAME]", Me.EDSTableName)
        SQLInsert = SQLInsert.Replace("[INSERT VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[INSERT FIELDS]", Me.SQLInsertFields)

        SQLInsert = SQLInsert.Replace("--[ANCHOR BLOCK PROFILE]", Me.AnchorProfile.SQLInsert)
        SQLInsert = SQLInsert.Replace("--[ANCHOR BLOCK SOIL PROFILE]", Me.SoilProfile.SQLInsert)

        Dim resInsert As String = ""
        For Each res As EDSResult In Me.Results
            resInsert += res.Insert & vbCrLf
        Next

        SQLInsert = SQLInsert.Replace("--[ANCHOR BLOCK RESULTS]", resInsert)
        SQLInsert = SQLInsert.TrimEnd

        Return SQLInsert + vbCrLf
    End Function

    Public Overrides Function SQLUpdate() As String
        SQLUpdate = ""
        SQLUpdate = CCI_Engineering_Templates.My.Resources.General__UPDATE
        SQLUpdate = SQLUpdate.Replace("[TABLE NAME]", Me.EDSTableName)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID)

        Dim _profUpdate As String
        Dim _soilUpdate As String

        If Me.AnchorProfile?.ID IsNot Nothing And Me.AnchorProfile?.ID > 0 Then
            _profUpdate = Me.AnchorProfile.SQLUpdate
        Else
            _profUpdate = Me.AnchorProfile.SQLInsert
        End If

        If Me.SoilProfile?.ID IsNot Nothing And Me.SoilProfile?.ID > 0 Then
            _soilUpdate = Me.SoilProfile.SQLUpdate
        Else
            _soilUpdate = Me.SoilProfile.SQLInsert
        End If

        SQLUpdate = SQLUpdate.Replace("--[OPTIONAL]", _profUpdate + vbCrLf + _soilUpdate + vbCrLf + Me.ResultQuery(True) + vbCrLf)

        Return SQLUpdate + vbCrLf
    End Function

    Public Overrides Function SQLDelete() As String

        SQLDelete = CCI_Engineering_Templates.My.Resources.Anchor_Block__DELETE_
        SQLDelete = SQLDelete.Replace("[SOIL PROFILE ID]", Me.soil_profile_id)
        SQLDelete = SQLDelete.Replace("[ANCHOR PROFILE ID]", Me.anchor_profile_id)
        SQLDelete = SQLDelete.Replace("[ANCHOR ID]", Me.ID)

        Return SQLDelete + vbCrLf
    End Function

#End Region

#Region "Define"
    Private _anchor_profile_id As Integer?
    Private _anchor_block_tool_id As Integer?
    Private _soil_profile_id As Integer?
    Private _local_anchor_id As Integer?
    'Private _reaction_position As Integer?
    Private _reaction_location As String
    Private _local_soil_profile_id As Integer?
    Private _local_anchor_profile_id As Integer?
    <Category("Guy Anchor Blocks"), Description(""), DisplayName("Anchor Profile Id")>
     <DataMember()> Public Property anchor_profile_id() As Integer?
        Get
            Return Me._anchor_profile_id
        End Get
        Set
            Me._anchor_profile_id = Value
        End Set
    End Property
    <Category("Guy Anchor Blocks"), Description(""), DisplayName("Anchor Block Tool Id")>
     <DataMember()> Public Property anchor_block_tool_id() As Integer?
        Get
            Return Me._anchor_block_tool_id
        End Get
        Set
            Me._anchor_block_tool_id = Value
        End Set
    End Property
    <Category("Guy Anchor Blocks"), Description(""), DisplayName("Soil Profile Id")>
     <DataMember()> Public Property soil_profile_id() As Integer?
        Get
            Return Me._soil_profile_id
        End Get
        Set
            Me._soil_profile_id = Value
        End Set
    End Property
    <Category("Guy Anchor Blocks"), Description(""), DisplayName("Local Anchor Id")>
     <DataMember()> Public Property local_anchor_id() As Integer?
        Get
            Return Me._local_anchor_id
        End Get
        Set
            Me._local_anchor_id = Value
        End Set
    End Property
    '<Category("Guy Anchor Blocks"), Description(""), DisplayName("Reaction Position")>
    ' <DataMember()> Public Property reaction_position() As Integer?
    '    Get
    '        Return Me._reaction_position
    '    End Get
    '    Set
    '        Me._reaction_position = Value
    '    End Set
    'End Property
    <Category("Guy Anchor Blocks"), Description(""), DisplayName("Reaction Location")>
     <DataMember()> Public Property reaction_location() As String
        Get
            Return Me._reaction_location
        End Get
        Set
            Me._reaction_location = Value
        End Set
    End Property
    <Category("Guy Anchor Blocks"), Description(""), DisplayName("Local Soil Profile Id")>
     <DataMember()> Public Property local_soil_profile_id() As Integer?
        Get
            Return Me._local_soil_profile_id
        End Get
        Set
            Me._local_soil_profile_id = Value
        End Set
    End Property
    <Category("Guy Anchor Blocks"), Description(""), DisplayName("Local Anchor Profile Id")>
     <DataMember()> Public Property local_anchor_profile_id() As Integer?
        Get
            Return Me._local_anchor_profile_id
        End Get
        Set
            Me._local_anchor_profile_id = Value
        End Set
    End Property
#End Region

#Region "Constructors"
    Public Sub New()

    End Sub

    Public Sub New(ByVal dr As DataRow)
        ConstructMe(dr)
    End Sub

    Public Sub ConstructMe(ByVal dr As DataRow)
        Me.anchor_profile_id = DBtoNullableInt(dr.Item("anchor_profile_id"))
        Me.anchor_block_tool_id = DBtoNullableInt(dr.Item("anchor_block_tool_id"))
        Me.soil_profile_id = DBtoNullableInt(dr.Item("soil_profile_id"))
        Me.local_anchor_id = DBtoNullableInt(dr.Item("local_anchor_id"))
        'Me.reaction_position = DBtoNullableInt(dr.Item("reaction_position"))
        Me.reaction_location = DBtoStr(dr.Item("reaction_location"))
        Me.local_soil_profile_id = DBtoNullableInt(dr.Item("local_soil_profile_id"))
        Me.local_anchor_profile_id = DBtoNullableInt(dr.Item("local_anchor_profile_id"))
        Me.ID = DBtoNullableInt(dr.Item("ID"))

    End Sub

    Public Sub New(ByVal dr As DataRow, ByVal strDS As DataSet, ByVal isExcel As Boolean, Optional ByVal Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Dim abProfile As New AnchorBlockProfile
        Dim abSProfile As New AnchorBlockSoilProfile
        Dim abLayer As New AnchorBlockSoilLayer

        ConstructMe(dr)

        For Each profileRow As DataRow In strDS.Tables(abProfile.EDSObjectName).Rows
            abProfile = New AnchorBlockProfile(profileRow, Me, isExcel)
            If IsSomething(abProfile.local_anchor_profile_id) Or (IsSomething(abProfile.ID) And IsNothing(abProfile.local_anchor_profile_id)) Then
                If If(isExcel, Me.local_anchor_profile_id = abProfile.local_anchor_profile_id, Me.anchor_profile_id = abProfile.ID) Then
                    Me.AnchorProfile = abProfile
                End If
            End If
        Next

        For Each soilRow As DataRow In strDS.Tables(abSProfile.EDSObjectName).Rows
            abSProfile = (New AnchorBlockSoilProfile(soilRow, Me, isExcel))
            If If(isExcel, Me.local_soil_profile_id = abSProfile.local_soil_profile_id, Me.soil_profile_id = abSProfile.ID) Then
                Me.SoilProfile = abSProfile

                For Each layerRow As DataRow In strDS.Tables(abLayer.EDSObjectName).Rows
                    abLayer = (New AnchorBlockSoilLayer(layerRow, abSProfile))
                    If isExcel And IsNothing(Me.soil_profile_id) And IsNothing(abLayer.Soil_Profile_id) And abSProfile.local_soil_profile_id = abLayer.local_soil_profile_id Then 'First time SA with no EDS IDs in Excel tool
                        abSProfile.ABSoilLayers.Add(abLayer)
                        abSProfile.SoilLayers.Add(abLayer)
                    ElseIf Me.soil_profile_id = abLayer.Soil_Profile_id Then 'From EDS, or second SA where tool has EDS IDs populated
                        abSProfile.ABSoilLayers.Add(abLayer)
                        abSProfile.SoilLayers.Add(abLayer)
                    End If
                Next
            End If
        Next

        If isExcel Then
            Dim res As New EDSResult

            For Each resRow As DataRow In strDS.Tables("Guy Anchor Result").Rows
                If DBtoNullableInt(resRow.Item("local_anchor_id")) = Me.local_anchor_id Then
                    res = New EDSResult(resRow, Me.Parent)
                    res.EDSTableName = "fnd.anchor_block_results"
                    res.ForeignKeyName = "anchor_block_id"
                    res.foreign_key = Me.ID
                    res.EDSTableDepth = 1
                    res.modified_person_id = Me.modified_person_id
                    Dim x As AnchorBlockFoundation = Me.Parent
                    x.Results.Add(res)
                    Me.Results.Add(res)
                End If
            Next
        End If
    End Sub

#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@TopLevelID") 'Me.anchor_block_tool_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel3ID") 'Me.anchor_profile_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel1ID") 'Me.soil_profile_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_anchor_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reaction_location.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_anchor_profile_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_soil_profile_id.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)
        Return SQLInsertValues
    End Function
    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""
        'SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
        SQLInsertFields = SQLInsertFields.AddtoDBString("anchor_block_tool_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("anchor_profile_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("soil_profile_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_anchor_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("reaction_location")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_anchor_profile_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_soil_profile_id")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")
        Return SQLInsertFields
    End Function
    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("anchor_block_tool_id = " & Me.anchor_block_tool_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("anchor_profile_id = " & Me.anchor_profile_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("soil_profile_id = " & Me.soil_profile_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_anchor_id = " & Me.local_anchor_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("reaction_location = " & Me.reaction_location.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_anchor_profile_id = " & Me.local_anchor_profile_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_soil_profile_id = " & Me.local_soil_profile_id.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_person_id = " & Me.modified_person_id.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("process_stage = " & Me.process_stage.ToString.FormatDBValue)
        Return SQLUpdateFieldsandValues
    End Function
#End Region

#Region "Equals"
    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName
        Dim otherToCompare As AnchorBlock = TryCast(other, AnchorBlock)
        If otherToCompare Is Nothing Then Return False
        'Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
        'Equals = If(Me.anchor_block_tool_id.CheckChange(otherToCompare.anchor_block_tool_id, changes, categoryName, "Anchor Block Tool Id"), Equals, False)
        'Equals = If(Me.anchor_profile_id.CheckChange(otherToCompare.anchor_profile_id, changes, categoryName, "Anchor Profile Id"), Equals, False)
        'Equals = If(Me.soil_profile_id.CheckChange(otherToCompare.soil_profile_id, changes, categoryName, "Soil Profile Id"), Equals, False)
        Equals = If(Me.local_anchor_id.CheckChange(otherToCompare.local_anchor_id, changes, categoryName, "Local Anchor Id"), Equals, False)
        Equals = If(Me.reaction_location.CheckChange(otherToCompare.reaction_location, changes, categoryName, "Reaction Location"), Equals, False)
        Equals = If(Me.local_anchor_profile_id.CheckChange(otherToCompare.local_anchor_profile_id, changes, categoryName, "Local Anchor Profile Id"), Equals, False)
        Equals = If(Me.local_soil_profile_id.CheckChange(otherToCompare.local_soil_profile_id, changes, categoryName, "Local Soil Profile Id"), Equals, False)
        Equals = If(Me.modified_person_id.CheckChange(otherToCompare.modified_person_id, changes, categoryName, "Modified Person Id"), Equals, False)
        Equals = If(Me.process_stage.CheckChange(otherToCompare.process_stage, changes, categoryName, "Process Stage"), Equals, False)
        If IsSomething(Me.local_anchor_id) Then 'If local ID isblank then there will be no Anchor Profile or Soil Profile associated to the Anchor Block object becuase it is to be deleted - MRR
            'Anchor Profile
            Equals = If(Me.AnchorProfile.CheckChange(otherToCompare.AnchorProfile, changes, categoryName, "Anchor Profile"), Equals, False)
            'Soil Profile
            Equals = If(Me.SoilProfile.CheckChange(otherToCompare.SoilProfile, changes, categoryName, "Soil Profile"), Equals, False)
        End If

        Return Equals
    End Function
#End Region
End Class

<DataContractAttribute()>
Partial Public Class AnchorBlockProfile
    Inherits EDSObjectWithQueries

#Region "Inheritted"
    '''Must override these inherited properties
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Guy Anchor Profiles"
        End Get
    End Property
    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "fnd.anchor_block_profile"
        End Get
    End Property

    Private _Insert As String
    Private _Update As String
    Private _Delete As String

    Public Overrides Function SQLInsert() As String
        SQLInsert = ""
        SQLInsert = CCI_Engineering_Templates.My.Resources.Anchor_Block_Profile__INSERT
        SQLInsert = SQLInsert.Replace("[TABLE NAME]", Me.EDSTableName)
        SQLInsert = SQLInsert.Replace("[INSERT FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.Replace("[INSERT VALUES]", Me.SQLInsertValues)

        Return SQLInsert.TrimEnd
    End Function

    Public Overrides Function SQLUpdate() As String
        SQLUpdate = ""
        SQLUpdate = CCI_Engineering_Templates.My.Resources.General__UPDATE
        SQLUpdate = SQLUpdate.Replace("[TABLE NAME]", Me.EDSTableName)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID)

        Return SQLUpdate
    End Function

    Public Overrides Function SQLDelete() As String
        SQLDelete = CCI_Engineering_Templates.My.Resources.General__DELETE
        SQLDelete = SQLDelete.Replace("[TABLE]", Me.EDSTableName)
        SQLDelete = SQLDelete.Replace("[ID]", Me.ID)

        Return SQLDelete
    End Function

#End Region

#Region "Define"
    Private _anchor_depth As Double?
    Private _anchor_width As Double?
    Private _anchor_thickness As Double?
    Private _anchor_length As Double?
    Private _anchor_toe_width As Double?
    Private _anchor_top_rebar_size As Integer?
    Private _anchor_top_rebar_quantity As Integer?
    Private _anchor_front_rebar_size As Integer?
    Private _anchor_front_rebar_quantity As Integer?
    Private _anchor_stirrup_size As Integer?
    Private _anchor_shaft_diameter As Double?
    Private _anchor_shaft_quantity As Integer?
    Private _anchor_shaft_area_override As Double?
    Private _anchor_shaft_shear_leg_factor As Double?
    Private _anchor_shaft_section As String
    Private _anchor_rebar_grade As Double?
    Private _concrete_compressive_strength As Double?
    Private _clear_cover As Double?
    Private _anchor_shaft_yield_strength As Double?
    Private _anchor_shaft_ultimate_strength As Double?
    Private _rebar_known As Boolean?
    Private _anchor_shaft_known As Boolean?
    Private _basic_soil_check As Boolean?
    Private _structural_check As Boolean?
    Private _local_anchor_profile_id As Integer?

    <Category("Guy Anchor Profiles"), Description(""), DisplayName("Anchor Depth")>
     <DataMember()> Public Property anchor_depth() As Double?
        Get
            Return Me._anchor_depth
        End Get
        Set
            Me._anchor_depth = Value
        End Set
    End Property
    <Category("Guy Anchor Profiles"), Description(""), DisplayName("Anchor Width")>
     <DataMember()> Public Property anchor_width() As Double?
        Get
            Return Me._anchor_width
        End Get
        Set
            Me._anchor_width = Value
        End Set
    End Property
    <Category("Guy Anchor Profiles"), Description(""), DisplayName("Anchor Thickness")>
     <DataMember()> Public Property anchor_thickness() As Double?
        Get
            Return Me._anchor_thickness
        End Get
        Set
            Me._anchor_thickness = Value
        End Set
    End Property
    <Category("Guy Anchor Profiles"), Description(""), DisplayName("Anchor Length")>
     <DataMember()> Public Property anchor_length() As Double?
        Get
            Return Me._anchor_length
        End Get
        Set
            Me._anchor_length = Value
        End Set
    End Property
    <Category("Guy Anchor Profiles"), Description(""), DisplayName("Anchor Toe Width")>
     <DataMember()> Public Property anchor_toe_width() As Double?
        Get
            Return Me._anchor_toe_width
        End Get
        Set
            Me._anchor_toe_width = Value
        End Set
    End Property
    <Category("Guy Anchor Profiles"), Description(""), DisplayName("Anchor Top Rebar Size")>
     <DataMember()> Public Property anchor_top_rebar_size() As Integer?
        Get
            Return Me._anchor_top_rebar_size
        End Get
        Set
            Me._anchor_top_rebar_size = Value
        End Set
    End Property
    <Category("Guy Anchor Profiles"), Description("column - 6+local drilled pier id"), DisplayName("Anchor Top Rebar Quantity")>
     <DataMember()> Public Property anchor_top_rebar_quantity() As Integer?
        Get
            Return Me._anchor_top_rebar_quantity
        End Get
        Set
            Me._anchor_top_rebar_quantity = Value
        End Set
    End Property
    <Category("Guy Anchor Profiles"), Description(""), DisplayName("Anchor Front Rebar Size")>
     <DataMember()> Public Property anchor_front_rebar_size() As Integer?
        Get
            Return Me._anchor_front_rebar_size
        End Get
        Set
            Me._anchor_front_rebar_size = Value
        End Set
    End Property
    <Category("Guy Anchor Profiles"), Description(""), DisplayName("Anchor Front Rebar Quantity")>
     <DataMember()> Public Property anchor_front_rebar_quantity() As Integer?
        Get
            Return Me._anchor_front_rebar_quantity
        End Get
        Set
            Me._anchor_front_rebar_quantity = Value
        End Set
    End Property
    <Category("Guy Anchor Profiles"), Description(""), DisplayName("Anchor Stirrup Size")>
     <DataMember()> Public Property anchor_stirrup_size() As Integer?
        Get
            Return Me._anchor_stirrup_size
        End Get
        Set
            Me._anchor_stirrup_size = Value
        End Set
    End Property
    <Category("Guy Anchor Profiles"), Description(""), DisplayName("Anchor Shaft Diameter")>
     <DataMember()> Public Property anchor_shaft_diameter() As Double?
        Get
            Return Me._anchor_shaft_diameter
        End Get
        Set
            Me._anchor_shaft_diameter = Value
        End Set
    End Property
    <Category("Guy Anchor Profiles"), Description(""), DisplayName("Anchor Shaft Quantity")>
     <DataMember()> Public Property anchor_shaft_quantity() As Integer?
        Get
            Return Me._anchor_shaft_quantity
        End Get
        Set
            Me._anchor_shaft_quantity = Value
        End Set
    End Property
    <Category("Guy Anchor Profiles"), Description(""), DisplayName("Anchor Shaft Area Override")>
     <DataMember()> Public Property anchor_shaft_area_override() As Double?
        Get
            Return Me._anchor_shaft_area_override
        End Get
        Set
            Me._anchor_shaft_area_override = Value
        End Set
    End Property
    <Category("Guy Anchor Profiles"), Description(""), DisplayName("Anchor Shaft Shear Leg Factor")>
     <DataMember()> Public Property anchor_shaft_shear_leg_factor() As Double?
        Get
            Return Me._anchor_shaft_shear_leg_factor
        End Get
        Set
            Me._anchor_shaft_shear_leg_factor = Value
        End Set
    End Property
    <Category("Guy Anchor Profiles"), Description(""), DisplayName("Anchor Shaft Section")>
     <DataMember()> Public Property anchor_shaft_section() As String
        Get
            Return Me._anchor_shaft_section
        End Get
        Set
            Me._anchor_shaft_section = Value
        End Set
    End Property
    <Category("Guy Anchor Profiles"), Description(""), DisplayName("Anchor Rebar Grade")>
     <DataMember()> Public Property anchor_rebar_grade() As Double?
        Get
            Return Me._anchor_rebar_grade
        End Get
        Set
            Me._anchor_rebar_grade = Value
        End Set
    End Property
    <Category("Guy Anchor Profiles"), Description(""), DisplayName("Concrete Compressive Strength")>
     <DataMember()> Public Property concrete_compressive_strength() As Double?
        Get
            Return Me._concrete_compressive_strength
        End Get
        Set
            Me._concrete_compressive_strength = Value
        End Set
    End Property
    <Category("Guy Anchor Profiles"), Description(""), DisplayName("Clear Cover")>
     <DataMember()> Public Property clear_cover() As Double?
        Get
            Return Me._clear_cover
        End Get
        Set
            Me._clear_cover = Value
        End Set
    End Property
    <Category("Guy Anchor Profiles"), Description(""), DisplayName("Anchor Shaft Yield Strength")>
     <DataMember()> Public Property anchor_shaft_yield_strength() As Double?
        Get
            Return Me._anchor_shaft_yield_strength
        End Get
        Set
            Me._anchor_shaft_yield_strength = Value
        End Set
    End Property
    <Category("Guy Anchor Profiles"), Description(""), DisplayName("Anchor Shaft Ultimate Strength")>
     <DataMember()> Public Property anchor_shaft_ultimate_strength() As Double?
        Get
            Return Me._anchor_shaft_ultimate_strength
        End Get
        Set
            Me._anchor_shaft_ultimate_strength = Value
        End Set
    End Property
    <Category("Guy Anchor Profiles"), Description(""), DisplayName("Rebar Known")>
     <DataMember()> Public Property rebar_known() As Boolean?
        Get
            Return Me._rebar_known
        End Get
        Set
            Me._rebar_known = Value
        End Set
    End Property
    <Category("Guy Anchor Profiles"), Description(""), DisplayName("Anchor Shaft Known")>
     <DataMember()> Public Property anchor_shaft_known() As Boolean?
        Get
            Return Me._anchor_shaft_known
        End Get
        Set
            Me._anchor_shaft_known = Value
        End Set
    End Property
    <Category("Guy Anchor Profiles"), Description(""), DisplayName("Basic Soil Check")>
     <DataMember()> Public Property basic_soil_check() As Boolean?
        Get
            Return Me._basic_soil_check
        End Get
        Set
            Me._basic_soil_check = Value
        End Set
    End Property
    <Category("Guy Anchor Profiles"), Description(""), DisplayName("Structural Check")>
     <DataMember()> Public Property structural_check() As Boolean?
        Get
            Return Me._structural_check
        End Get
        Set
            Me._structural_check = Value
        End Set
    End Property
    <Category("AnchorBlockProfile"), Description(""), DisplayName("Local Anchor Profile Id")>
     <DataMember()> Public Property local_anchor_profile_id() As Integer?
        Get
            Return Me._local_anchor_profile_id
        End Get
        Set
            Me._local_anchor_profile_id = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()

    End Sub

    Public Sub New(ByVal dr As DataRow, Optional ByVal Parent As EDSObject = Nothing, Optional ByVal isExcel As Boolean = False)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)
        Me.anchor_depth = DBtoNullableDbl(dr.Item("anchor_depth"))
        Me.anchor_width = DBtoNullableDbl(dr.Item("anchor_width"))
        Me.anchor_thickness = DBtoNullableDbl(dr.Item("anchor_thickness"))
        Me.anchor_length = DBtoNullableDbl(dr.Item("anchor_length"))
        Me.anchor_toe_width = DBtoNullableDbl(dr.Item("anchor_toe_width"))
        Me.anchor_top_rebar_size = DBtoNullableInt(dr.Item("anchor_top_rebar_size"))
        Me.anchor_top_rebar_quantity = DBtoNullableInt(dr.Item("anchor_top_rebar_quantity"))
        Me.anchor_front_rebar_size = DBtoNullableInt(dr.Item("anchor_front_rebar_size"))
        Me.anchor_front_rebar_quantity = DBtoNullableInt(dr.Item("anchor_front_rebar_quantity"))
        Me.anchor_stirrup_size = DBtoNullableInt(dr.Item("anchor_stirrup_size"))
        Me.anchor_shaft_diameter = DBtoNullableDbl(dr.Item("anchor_shaft_diameter"))
        Me.anchor_shaft_quantity = DBtoNullableInt(dr.Item("anchor_shaft_quantity"))
        Me.anchor_shaft_area_override = DBtoNullableDbl(dr.Item("anchor_shaft_area_override"))
        Me.anchor_shaft_shear_leg_factor = DBtoNullableDbl(dr.Item("anchor_shaft_shear_leg_factor"))
        Me.anchor_shaft_section = DBtoStr(dr.Item("anchor_shaft_section"))
        Me.anchor_rebar_grade = DBtoNullableDbl(dr.Item("anchor_rebar_grade"))
        Me.concrete_compressive_strength = DBtoNullableDbl(dr.Item("concrete_compressive_strength"))
        Me.clear_cover = DBtoNullableDbl(dr.Item("clear_cover"))
        Me.anchor_shaft_yield_strength = DBtoNullableDbl(dr.Item("anchor_shaft_yield_strength"))
        Me.anchor_shaft_ultimate_strength = DBtoNullableDbl(dr.Item("anchor_shaft_ultimate_strength"))
        Me.rebar_known = DBtoNullableBool(dr.Item("rebar_known"))
        Me.anchor_shaft_known = DBtoNullableBool(dr.Item("anchor_shaft_known"))
        Me.basic_soil_check = DBtoNullableBool(dr.Item("basic_soil_check"))
        Me.structural_check = DBtoNullableBool(dr.Item("structural_check"))
        Me.local_anchor_profile_id = DBtoNullableInt(dr.Item("local_anchor_profile_id"))


        Dim tempParent As AnchorBlock = TryCast(Me.Parent, AnchorBlock)
        If isExcel Then
            Me.ID = tempParent.anchor_profile_id
        Else
            Me.ID = DBtoNullableInt(dr.Item("ID"))
        End If

    End Sub

#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_anchor_profile_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.anchor_depth.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.anchor_width.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.anchor_thickness.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.anchor_length.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.anchor_toe_width.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.anchor_top_rebar_size.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.anchor_top_rebar_quantity.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.anchor_front_rebar_size.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.anchor_front_rebar_quantity.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.anchor_stirrup_size.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.anchor_shaft_diameter.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.anchor_shaft_quantity.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.anchor_shaft_area_override.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.anchor_shaft_shear_leg_factor.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.anchor_shaft_section.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.anchor_rebar_grade.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.concrete_compressive_strength.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.clear_cover.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.anchor_shaft_yield_strength.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.anchor_shaft_ultimate_strength.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rebar_known.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.anchor_shaft_known.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.basic_soil_check.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.structural_check.ToString.FormatDBValue)
        Return SQLInsertValues
    End Function
    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""
        'SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_anchor_profile_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("anchor_depth")
        SQLInsertFields = SQLInsertFields.AddtoDBString("anchor_width")
        SQLInsertFields = SQLInsertFields.AddtoDBString("anchor_thickness")
        SQLInsertFields = SQLInsertFields.AddtoDBString("anchor_length")
        SQLInsertFields = SQLInsertFields.AddtoDBString("anchor_toe_width")
        SQLInsertFields = SQLInsertFields.AddtoDBString("anchor_top_rebar_size")
        SQLInsertFields = SQLInsertFields.AddtoDBString("anchor_top_rebar_quantity")
        SQLInsertFields = SQLInsertFields.AddtoDBString("anchor_front_rebar_size")
        SQLInsertFields = SQLInsertFields.AddtoDBString("anchor_front_rebar_quantity")
        SQLInsertFields = SQLInsertFields.AddtoDBString("anchor_stirrup_size")
        SQLInsertFields = SQLInsertFields.AddtoDBString("anchor_shaft_diameter")
        SQLInsertFields = SQLInsertFields.AddtoDBString("anchor_shaft_quantity")
        SQLInsertFields = SQLInsertFields.AddtoDBString("anchor_shaft_area_override")
        SQLInsertFields = SQLInsertFields.AddtoDBString("anchor_shaft_shear_leg_factor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("anchor_shaft_section")
        SQLInsertFields = SQLInsertFields.AddtoDBString("anchor_rebar_grade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("concrete_compressive_strength")
        SQLInsertFields = SQLInsertFields.AddtoDBString("clear_cover")
        SQLInsertFields = SQLInsertFields.AddtoDBString("anchor_shaft_yield_strength")
        SQLInsertFields = SQLInsertFields.AddtoDBString("anchor_shaft_ultimate_strength")
        SQLInsertFields = SQLInsertFields.AddtoDBString("rebar_known")
        SQLInsertFields = SQLInsertFields.AddtoDBString("anchor_shaft_known")
        SQLInsertFields = SQLInsertFields.AddtoDBString("basic_soil_check")
        SQLInsertFields = SQLInsertFields.AddtoDBString("structural_check")
        Return SQLInsertFields
    End Function
    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_anchor_profile_id = " & Me.local_anchor_profile_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("anchor_depth = " & Me.anchor_depth.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("anchor_width = " & Me.anchor_width.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("anchor_thickness = " & Me.anchor_thickness.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("anchor_length = " & Me.anchor_length.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("anchor_toe_width = " & Me.anchor_toe_width.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("anchor_top_rebar_size = " & Me.anchor_top_rebar_size.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("anchor_top_rebar_quantity = " & Me.anchor_top_rebar_quantity.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("anchor_front_rebar_size = " & Me.anchor_front_rebar_size.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("anchor_front_rebar_quantity = " & Me.anchor_front_rebar_quantity.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("anchor_stirrup_size = " & Me.anchor_stirrup_size.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("anchor_shaft_diameter = " & Me.anchor_shaft_diameter.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("anchor_shaft_quantity = " & Me.anchor_shaft_quantity.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("anchor_shaft_area_override = " & Me.anchor_shaft_area_override.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("anchor_shaft_shear_leg_factor = " & Me.anchor_shaft_shear_leg_factor.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("anchor_shaft_section = " & Me.anchor_shaft_section.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("anchor_rebar_grade = " & Me.anchor_rebar_grade.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("concrete_compressive_strength = " & Me.concrete_compressive_strength.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("clear_cover = " & Me.clear_cover.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("anchor_shaft_yield_strength = " & Me.anchor_shaft_yield_strength.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("anchor_shaft_ultimate_strength = " & Me.anchor_shaft_ultimate_strength.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("rebar_known = " & Me.rebar_known.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("anchor_shaft_known = " & Me.anchor_shaft_known.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("basic_soil_check = " & Me.basic_soil_check.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("structural_check = " & Me.structural_check.ToString.FormatDBValue)
        Return SQLUpdateFieldsandValues
    End Function
#End Region

#Region "Equals"
    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName
        Dim otherToCompare As AnchorBlockProfile = TryCast(other, AnchorBlockProfile)
        If otherToCompare Is Nothing Then Return False
        'Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
        Equals = If(Me.local_anchor_profile_id.CheckChange(otherToCompare.local_anchor_profile_id, changes, categoryName, "Local Anchor Profile Id"), Equals, False)
        Equals = If(Me.anchor_depth.CheckChange(otherToCompare.anchor_depth, changes, categoryName, "Anchor Depth"), Equals, False)
        Equals = If(Me.anchor_width.CheckChange(otherToCompare.anchor_width, changes, categoryName, "Anchor Width"), Equals, False)
        Equals = If(Me.anchor_thickness.CheckChange(otherToCompare.anchor_thickness, changes, categoryName, "Anchor Thickness"), Equals, False)
        Equals = If(Me.anchor_length.CheckChange(otherToCompare.anchor_length, changes, categoryName, "Anchor Length"), Equals, False)
        Equals = If(Me.anchor_toe_width.CheckChange(otherToCompare.anchor_toe_width, changes, categoryName, "Anchor Toe Width"), Equals, False)
        Equals = If(Me.anchor_top_rebar_size.CheckChange(otherToCompare.anchor_top_rebar_size, changes, categoryName, "Anchor Top Rebar Size"), Equals, False)
        Equals = If(Me.anchor_top_rebar_quantity.CheckChange(otherToCompare.anchor_top_rebar_quantity, changes, categoryName, "Anchor Top Rebar Quantity"), Equals, False)
        Equals = If(Me.anchor_front_rebar_size.CheckChange(otherToCompare.anchor_front_rebar_size, changes, categoryName, "Anchor Front Rebar Size"), Equals, False)
        Equals = If(Me.anchor_front_rebar_quantity.CheckChange(otherToCompare.anchor_front_rebar_quantity, changes, categoryName, "Anchor Front Rebar Quantity"), Equals, False)
        Equals = If(Me.anchor_stirrup_size.CheckChange(otherToCompare.anchor_stirrup_size, changes, categoryName, "Anchor Stirrup Size"), Equals, False)
        Equals = If(Me.anchor_shaft_diameter.CheckChange(otherToCompare.anchor_shaft_diameter, changes, categoryName, "Anchor Shaft Diameter"), Equals, False)
        Equals = If(Me.anchor_shaft_quantity.CheckChange(otherToCompare.anchor_shaft_quantity, changes, categoryName, "Anchor Shaft Quantity"), Equals, False)
        Equals = If(Me.anchor_shaft_area_override.CheckChange(otherToCompare.anchor_shaft_area_override, changes, categoryName, "Anchor Shaft Area Override"), Equals, False)
        Equals = If(Me.anchor_shaft_shear_leg_factor.CheckChange(otherToCompare.anchor_shaft_shear_leg_factor, changes, categoryName, "Anchor Shaft Shear Lag Factor"), Equals, False)
        Equals = If(Me.anchor_shaft_section.CheckChange(otherToCompare.anchor_shaft_section, changes, categoryName, "Anchor Shaft Section"), Equals, False)
        Equals = If(Me.anchor_rebar_grade.CheckChange(otherToCompare.anchor_rebar_grade, changes, categoryName, "Anchor Rebar Grade"), Equals, False)
        Equals = If(Me.concrete_compressive_strength.CheckChange(otherToCompare.concrete_compressive_strength, changes, categoryName, "Concrete Compressive Strength"), Equals, False)
        Equals = If(Me.clear_cover.CheckChange(otherToCompare.clear_cover, changes, categoryName, "Clear Cover"), Equals, False)
        Equals = If(Me.anchor_shaft_yield_strength.CheckChange(otherToCompare.anchor_shaft_yield_strength, changes, categoryName, "Anchor Shaft Yield Strength"), Equals, False)
        Equals = If(Me.anchor_shaft_ultimate_strength.CheckChange(otherToCompare.anchor_shaft_ultimate_strength, changes, categoryName, "Anchor Shaft Ultimate Strength"), Equals, False)
        Equals = If(Me.rebar_known.CheckChange(otherToCompare.rebar_known, changes, categoryName, "Rebar Known"), Equals, False)
        Equals = If(Me.anchor_shaft_known.CheckChange(otherToCompare.anchor_shaft_known, changes, categoryName, "Anchor Shaft Known"), Equals, False)
        Equals = If(Me.basic_soil_check.CheckChange(otherToCompare.basic_soil_check, changes, categoryName, "Basic Soil Check"), Equals, False)
        Equals = If(Me.structural_check.CheckChange(otherToCompare.structural_check, changes, categoryName, "Structural Check"), Equals, False)
        Return Equals
    End Function
#End Region

End Class

'<DataContractAttribute()>
Partial Public Class AnchorBlockSoilProfile
    Inherits SoilProfile

     <DataMember()> Public Property local_soil_profile_id As Integer?
     <DataMember()> Public Property ABSoilLayers As New List(Of AnchorBlockSoilLayer)

    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Soil Profiles"
        End Get
    End Property

    Public Sub New()

    End Sub

    Public Sub New(ByVal dr As DataRow, Optional ByVal Parent As EDSObject = Nothing, Optional ByVal isExcel As Boolean = False)
        ConstructMe(dr, Parent)

        Try
            Me.local_soil_profile_id = DBtoNullableDbl(dr.Item("local_soil_profile_id"))
        Catch
        End Try

        Dim tempParent As AnchorBlock = TryCast(Me.Parent, AnchorBlock)
        If isExcel Then
            Me.ID = tempParent.soil_profile_id
        Else
            Me.ID = DBtoNullableInt(dr.Item("ID"))
        End If
    End Sub


    Public Overrides Function SQLUpdate() As String

        SQLUpdate = CCI_Engineering_Templates.My.Resources.Soil_Profile_UPDATE
        'SQLUpdate = QueryBuilderFromFile(queryPath & "Soil Profile\Soil Profile (UPDATE).sql")
        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
        SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

        Dim slUp As String

        For Each sl In Me.ABSoilLayers
            Dim absl As AnchorBlockSoilLayer = TryCast(sl, AnchorBlockSoilLayer)

            If absl.ID IsNot Nothing And absl?.ID > 0 Then
                If absl.local_soil_layer_id IsNot Nothing Then
                    slUp += absl.SQLUpdate
                Else
                    slUp += absl.SQLDelete
                End If
            Else
                slUp += absl.SQLInsert
            End If

            slUp += vbCrLf
        Next

        SQLUpdate += vbCrLf + slUp

        Return SQLUpdate
    End Function

End Class

<DataContractAttribute()>
Partial Public Class AnchorBlockSoilLayer
    Inherits SoilLayer

     <DataMember()> Public Property local_soil_profile_id As Integer?
     <DataMember()> Public Property local_soil_layer_id As Integer?

    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Soil Layers"
        End Get
    End Property

    Public Sub New()

    End Sub

    Public Sub New(ByVal dr As DataRow, Optional ByVal Parent As EDSObject = Nothing)
        ConstructMe(dr, Parent)

        Try
            Me.Soil_Profile_id = DBtoNullableDbl(dr.Item("soil_profile_id"))
            Me.local_soil_profile_id = DBtoNullableDbl(dr.Item("local_soil_profile_id"))
            Me.local_soil_layer_id = DBtoNullableDbl(dr.Item("local_soil_layer_id"))
        Catch
        End Try
    End Sub

End Class

<DataContractAttribute()>
Partial Public Class AnchorBlockResult
    Inherits EDSResult

    Public ReadOnly Property EDSObjectName As String
        Get
            Return "Guy Anchor Result"
        End Get
    End Property

    Public ReadOnly Property EDSTableName As String
        Get
            Return "fnd.anchor_block_results"
        End Get
    End Property

End Class


