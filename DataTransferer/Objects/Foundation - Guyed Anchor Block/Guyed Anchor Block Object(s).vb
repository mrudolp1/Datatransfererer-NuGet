Option Strict Off
Option Compare Binary

Imports System.ComponentModel
Imports System.Data
Imports DevExpress.Spreadsheet
'Imports Microsoft.Office.Interop

Partial Public Class AnchorBlockFoundation
    Inherits EDSExcelObject

    Public Property AnchorBlocks As New List(Of AnchorBlock)
    'Example for Rudy
    'Origin row in the driled pier database. Basically just where the profile numbers are in the database worksheet.
    'This is actually 58 but due to the 0,0 origin in excel, it is 1 less
    Private pierProfileRow As Integer = 57

    Private _file_ver As String
    Private _modified As Boolean?
    <Category("Guyed Anchor Block Profile"), Description(""), DisplayName("File Ver")>
    Public Property file_ver() As String
        Get
            Return Me._file_ver
        End Get
        Set
            Me._file_ver = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Profile"), Description(""), DisplayName("Modified")>
    Public Property modified() As Boolean?
        Get
            Return Me._modified
        End Get
        Set
            Me._modified = Value
        End Set
    End Property


#Region "Constructors"
    Public Sub New()
    End Sub

    Private Sub ConstructMe(ByVal dr As DataRow)
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        Me._file_ver = DBtoStr(dr.Item("_file_ver"))
        Me.modified = DBtoStr(dr.Item("modified"))
    End Sub

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

    Public Sub New(ByVal filepath As String, Optional ByVal Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)


        ''''''Customize for each foundation type'''''
        Dim excelDS As New DataSet

        For Each item As EXCELDTParameter In excelDTParams
            Try
                excelDS.Tables.Add(ExcelDatasourceToDataTable(GetExcelDataSource(filepath, item.xlsSheet, item.xlsRange), item.xlsDatatable))
            Catch ex As Exception
                Debug.Print(String.Format("Failed to create datatable for: {0}, {1}, {2}", IO.Path.GetFileName(filepath), item.xlsSheet, item.xlsRange))
            End Try
        Next

        Dim dr As DataRow = excelDS.Tables("Anchor Block Foundation").Rows(0)

        ConstructMe(dr)

        Dim myAB As New AnchorBlock
        For Each abrow As DataRow In excelDS.Tables(myAB.EDSObjectName).Rows
            If IsSomething(abrow.Item("local_anchor_id")) Or (IsSomething(abrow.Item("ID")) And IsNothing(abrow.Item("local_anchor_id"))) Then
                Me.AnchorBlocks.Add(New AnchorBlock(abrow, excelDS, True, Me))
            End If
        Next
    End Sub
#End Region

#Region "Inherited"
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Anchor Block Foundation"
        End Get
    End Property

    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "fnd.anchor_block_tool"
        End Get
    End Property

    Public Overrides ReadOnly Property templatePath As String
        Get
            Return IO.Path.Combine(My.Application.Info.DirectoryPath, "Templates", "Guyed Anchor Block Foundation.xlsm")
        End Get
    End Property

    Public Overrides ReadOnly Property excelDTParams As List(Of EXCELDTParameter)
        Get
            Dim ab As New AnchorBlock
            Dim abProf As New AnchorBlockProfile
            Dim abSProf As New AnchorBlockSoilProfile
            Dim abSlay As New AnchorBlockSoilLayer
            Dim abRes As New AnchorBlockResult
            Dim abTool As New AnchorBlockFoundation

            Return New List(Of EXCELDTParameter) From {
                                                        New EXCELDTParameter(ab.EDSObjectName, "A2:H52", "Profiles (ENTER)"),  'It is slightly confusing but to keep naming issues consistent in the tool a drilled pier = profile and a drilled pier profile = drilled pier details
                                                        New EXCELDTParameter(abProf.EDSObjectName, "A2:X52", "Details (ENTER)"),
                                                        New EXCELDTParameter(abSProf.EDSObjectName, "A2:E52", "Soil Profiles (ENTER)"),
                                                        New EXCELDTParameter(abSlay.EDSObjectName, "A2:N1502", "Soil Layers (ENTER)"),
                                                        New EXCELDTParameter(abRes.EDSObjectName, "BD8:BV58", "Foundation Input"),
                                                        New EXCELDTParameter(abTool.EDSObjectName, "A1:E2", "Tool (RETURN_ENTER)")
                                                                                        }
            '***Add additional table references here****
        End Get
    End Property

    Private _Insert As String
    Private _Update As String
    Private _Delete As String

    Public Overrides Function SQLInsert() As String
        Dim _abInsert As String

        SQLInsert = ""
        SQLInsert = CCI_Engineering_Templates.My.Resources.Anchor_Block_Tool__INSERT

        'Guy Anchors
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
                If gab.local_anchor_profile_id IsNot Nothing Then
                    gbUp += gab.SQLUpdate
                Else
                    gbUp += gab.SQLDelete
                End If
            Else
                gbUp += gab.SQLInsert
            End If
        Next
    End Function

    Public Overrides Function SQLDelete() As String
        SQLDelete = CCI_Engineering_Templates.My.Resources.General__DELETE
        SQLDelete = SQLDelete.Replace("[TABLE]", Me.EDSTableName)
        SQLDelete = SQLDelete.Replace("[ID]", Me.ID)

        Return SQLDelete
    End Function

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

#Region "Save to Excel"
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

            If Not IsNothing(Me.ID) Then
                .Worksheets("Tool (SAPI)").Range("A3").Value = CType(Me.ID, Integer)
            Else
                .Worksheets("Tool (SAPI)").Range("A3").ClearContents
            End If
            If Not IsNothing(Me.bus_unit) Then
                .Worksheets("Tool (SAPI)").Range("B3").Value = CType(Me.bus_unit, Integer)
            Else
                .Worksheets("Tool (SAPI)").Range("B3").ClearContents
            End If
            If Not IsNothing(Me.structure_id) Then
                .Worksheets("Tool (SAPI)").Range("C3").Value = CType(Me.structure_id, String)
            End If
            'If Not IsNothing(Me.file_ver) Then
            '    .Worksheets("Tool (SAPI)").Range("D3").Value = CType(Me.file_ver, String)
            'End If
            If Not IsNothing(Me.modified) Then
                .Worksheets("Tool (SAPI)").Range("E3").Value = CType(Me.modified, Boolean)
            End If
            'If Not IsNothing(Me.modified_person_id) Then
            '    .Worksheets("Tool (SAPI)").Range("F3").Value = CType(Me.modified_person_id, Integer)
            'Else
            '    .Worksheets("Tool (SAPI)").Range("F3").ClearContents
            'End If
            'If Not IsNothing(Me.process_stage) Then
            '    .Worksheets("Tool (SAPI)").Range("G3").Value = CType(Me.process_stage, String)
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
            .Worksheets("Tool (SAPI)").Range("L3").Value = CType(gab_tia_current, String)
            'Load Z Normalization
            'If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.load_z_norm) Then
            '    rev_h_section_15_5 = Me.ParentStructure?.structureCodeCriteria?.load_z_norm
            '    .Worksheets("General (SAPI)").Range("Q3").Value = CType(load_z_norm, Boolean)
            'End If
            'H Section 15.5
            If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.rev_h_section_15_5) Then
                gab_rev_h_section_15_5 = Me.ParentStructure?.structureCodeCriteria?.rev_h_section_15_5
                .Worksheets("Tool (SAPI)").Range("M3").Value = CType(gab_rev_h_section_15_5, Boolean)
            End If
            'Work Order
            If Not IsNothing(Me.ParentStructure?.work_order_seq_num) Then
                work_order_seq_num = Me.ParentStructure?.work_order_seq_num
                .Worksheets("Tool (SAPI)").Range("K3").Value = CType(work_order_seq_num, Integer)
            End If
            'Site Name
            If Not IsNothing(Me.ParentStructure?.SiteInfo.site_name) Then
                site_name = Me.ParentStructure?.SiteInfo.site_name
                .Worksheets("Tool (SAPI)").Range("J3").Value = CType(site_name, String)
            End If

            'Anchors
            Dim i As Integer = 3
            For Each gab As AnchorBlock In AnchorBlocks
                If Not IsNothing(gab.ID) Then .Worksheets("Anchors (SAPI)").Range("A" & i).Value = CType(gab.ID, Integer)
                If Not IsNothing(gab.anchor_block_tool_id) Then .Worksheets("Anchors (SAPI").Range("B" & i).Value = CType(gab.anchor_block_tool_id, Integer)
                If Not IsNothing(gab.anchor_profile_id) Then .Worksheets("Anchors (SAPI").Range("C" & i).Value = CType(gab.anchor_profile_id, Integer)
                If Not IsNothing(gab.soil_profile_id) Then .Worksheets("Anchors (SAPI").Range("D" & i).Value = CType(gab.soil_profile_id, Integer)
                If Not IsNothing(gab.local_anchor_id) Then .Worksheets("Anchors (SAPI").Range("E" & i).Value = CType(gab.local_anchor_id, Integer)
                If Not IsNothing(gab.reaction_location) Then .Worksheets("Anchors (SAPI").Range("F" & i).Value = CType(gab.reaction_location, String)
                If Not IsNothing(gab.local_anchor_profile_id) Then .Worksheets("Anchors (SAPI").Range("G" & i).Value = CType(gab.local_anchor_profile_id, Integer)
                If Not IsNothing(gab.local_soil_profile_id) Then .Worksheets("Anchors (SAPI").Range("H" & i).Value = CType(gab.local_soil_profile_id, Integer)
                i += 1
            Next

            'Profile
            i = 3
            For Each ap As AnchorBlockProfile In AnchorProfiles
                If Not IsNothing(ap.ID) Then .Worksheets("Anchor Profiles (SAPI)").Range("A" & i).Value = CType(ap.ID, Integer)
                If Not IsNothing(ap.local_anchor_profile_id) Then .Worksheets("Anchor Profiles (SAPI)").Range("B" & i).Value = CType(ap.local_anchor_profile_id, Integer)
                If Not IsNothing(ap.anchor_depth) Then .Worksheets("Anchor Profiles (SAPI)").Range("C" & i).Value = CType(ap.anchor_depth, Double)
                If Not IsNothing(ap.anchor_width) Then .Worksheets("Anchor Profiles (SAPI)").Range("D" & i).Value = CType(ap.anchor_width, Double)
                If Not IsNothing(ap.anchor_thickness) Then .Worksheets("Anchor Profiles (SAPI)").Range("E" & i).Value = CType(ap.anchor_thickness, Double)
                If Not IsNothing(ap.anchor_length) Then .Worksheets("Anchor Profiles (SAPI)").Range("F" & i).Value = CType(ap.anchor_length, Double)
                If Not IsNothing(ap.anchor_toe_width) Then .Worksheets("Anchor Profiles (SAPI)").Range("G" & i).Value = CType(ap.anchor_toe_width, Double)
                If Not IsNothing(ap.anchor_top_rebar_size) Then .Worksheets("Anchor Profiles (SAPI)").Range("H" & i).Value = CType(ap.anchor_top_rebar_size, Integer)
                If Not IsNothing(ap.anchor_top_rebar_quantity) Then .Worksheets("Anchor Profiles (SAPI)").Range("I" & i).Value = CType(ap.anchor_top_rebar_quantity, Integer)
                If Not IsNothing(ap.anchor_front_rebar_size) Then .Worksheets("Anchor Profiles (SAPI)").Range("J" & i).Value = CType(ap.anchor_front_rebar_size, Integer)
                If Not IsNothing(ap.anchor_front_rebar_quantity) Then .Worksheets("Anchor Profiles (SAPI)").Range("K" & i).Value = CType(ap.anchor_front_rebar_quantity, Integer)
                If Not IsNothing(ap.anchor_stirrup_size) Then .Worksheets("Anchor Profiles (SAPI)").Range("L" & i).Value = CType(ap.anchor_stirrup_size, Integer)
                If Not IsNothing(ap.anchor_shaft_diameter) Then .Worksheets("Anchor Profiles (SAPI)").Range("M" & i).Value = CType(ap.anchor_shaft_diameter, Double)
                If Not IsNothing(ap.anchor_shaft_quantity) Then .Worksheets("Anchor Profiles (SAPI)").Range("N" & i).Value = CType(ap.anchor_shaft_quantity, Integer)
                If Not IsNothing(ap.anchor_shaft_area_override) Then .Worksheets("Anchor Profiles (SAPI)").Range("O" & i).Value = CType(ap.anchor_shaft_area_override, Double)
                If Not IsNothing(ap.anchor_shaft_shear_lag_factor) Then .Worksheets("Anchor Profiles (SAPI)").Range("P" & i).Value = CType(ap.anchor_shaft_shear_lag_factor, Double)
                If Not IsNothing(ap.anchor_shaft_section) Then .Worksheets("Anchor Profiles (SAPI)").Range("Q" & i).Value = CType(ap.anchor_shaft_section, String)
                If Not IsNothing(ap.anchor_rebar_grade) Then .Worksheets("Anchor Profiles (SAPI)").Range("R" & i).Value = CType(ap.anchor_rebar_grade, Double)
                If Not IsNothing(ap.concrete_compressive_strength) Then .Worksheets("Anchor Profiles (SAPI)").Range("S" & i).Value = CType(ap.concrete_compressive_strength, Double)
                If Not IsNothing(ap.clear_cover) Then .Worksheets("Anchor Profiles (SAPI)").Range("T" & i).Value = CType(ap.clear_cover, Double)
                If Not IsNothing(ap.anchor_shaft_yield_strength) Then .Worksheets("Anchor Profiles (SAPI)").Range("U" & i).Value = CType(ap.anchor_shaft_yield_strength, Double)
                If Not IsNothing(ap.anchor_shaft_ultimate_strength) Then .Worksheets("Anchor Profiles (SAPI)").Range("V" & i).Value = CType(ap.anchor_shaft_ultimate_strength, Double)
                If Not IsNothing(ap.rebar_known) Then .Worksheets("Anchor Profiles (SAPI)").Range("W" & i).Value = CType(ap.rebar_known, Boolean)
                If Not IsNothing(ap.anchor_shaft_known) Then .Worksheets("Anchor Profiles (SAPI)").Range("X" & i).Value = CType(ap.anchor_shaft_known, Boolean)
                If Not IsNothing(ap.basic_soil_check) Then .Worksheets("Anchor Profiles (SAPI)").Range("Y" & i).Value = CType(ap.basic_soil_check, Boolean)
                If Not IsNothing(ap.structural_check) Then .Worksheets("Anchor Profiles (SAPI)").Range("Z" & i).Value = CType(ap.structural_check, Boolean)
                i += 1
            Next

            'Soil Profile
            i = 3
            Dim j As Integer = 3
            For Each sp As AnchorBlockSoilProfile In SoilProfiles
                If Not IsNothing(sp.ID) Then .Worksheets("Soil Profiles (SAPI)").Range("A" & i).Value = CType(sp.ID, Integer)
                If Not IsNothing(sp.local_soil_profile_id) Then .Worksheets("Soil Profiles (SAPI)").Range("B" & i).Value = CType(sp.local_soil_profile_id, Integer)
                If Not IsNothing(sp.groundwater_depth) Then .Worksheets("Soil Profiles (SAPI)").Range("C" & i).Value = CType(sp.groundwater_depth, Double)
                If Not IsNothing(sp.neglect_depth) Then .Worksheets("Soil Profiles (SAPI)").Range("D" & i).Value = CType(sp.neglect_depth, Double)

                'Soil Layers
                For Each layer In sp.SoilLayers
                    If Not IsNothing(layer.ID) Then .Worksheets("Soil Layers (SAPI)").Range("A" & j).Value = CType(layer.ID, Integer)
                    If Not IsNothing(layer.Soil_Profile_id) Then .Worksheets("Soil Layers (SAPI)").Range("B" & j).Value = CType(layer.Soil_Profile_id, Integer)
                    If Not IsNothing(layer.local_soil_profile_id) Then .Worksheets("Soil Layers (SAPI)").Range("C" & j).Value = CType(layer.local_soil_profile_id, Integer)
                    If Not IsNothing(layer.local_soil_layer_id) Then .Worksheets("Soil Layers (SAPI)").Range("D" & j).Value = CType(layer.local_soil_layer_id, Integer)
                    If Not IsNothing(layer.bottom_depth) Then .Worksheets("Soil Layers (SAPI)").Range("E" & j).Value = CType(layer.bottom_depth, Double)
                    If Not IsNothing(layer.effective_soil_density) Then .Worksheets("Soil Layers (SAPI)").Range("F" & j).Value = CType(layer.effective_soil_density, Double)
                    If Not IsNothing(layer.cohesion) Then .Worksheets("Soil Layers (SAPI)").Range("G" & j).Value = CType(layer.cohesion, Double)
                    If Not IsNothing(layer.friction_angle) Then .Worksheets("Soil Layers (SAPI)").Range("H" & j).Value = CType(layer.friction_angle, Double)
                    If Not IsNothing(layer.skin_friction_override_comp) Then .Worksheets("Soil Layers (SAPI)").Range("I" & j).Value = CType(layer.skin_friction_override_comp, Double)
                    If Not IsNothing(layer.skin_friction_override_uplift) Then .Worksheets("Soil Layers (SAPI)").Range("J" & j).Value = CType(layer.skin_friction_override_uplift, Double)
                    If Not IsNothing(layer.nominal_bearing_capacity) Then .Worksheets("Soil Layers (SAPI)").Range("K" & j).Value = CType(layer.nominal_bearing_capacity, Double)
                    If Not IsNothing(layer.spt_blow_count) Then .Worksheets("Soil Layers (SAPI)").Range("L" & j).Value = CType(layer.spt_blow_count, Double)

                    j += 1
                Next

                i += 1
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

Partial Public Class AnchorBlock
    Inherits EDSObjectWithQueries

    Public Property AnchorProfile As AnchorBlockProfile
    Public Property SoilProfile As AnchorBlockSoilProfile

#Region "Inheritted"
    '''Must override these inherited properties
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Anchor Block Profile"
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
        SQLInsert = SQLInsert.Replace("--[ANCHOR BLOCK PROFILE]", Me.AnchorProfile.SQLInsert)
        SQLInsert = SQLInsert.Replace("--[ANCHOR BLOCK SOIL PROFILE]", Me.SoilProfile.SQLInsert)

        Dim resInsert As String = ""
        For Each res As EDSResult In Me.Results
            resInsert += res.Insert & vbCrLf
        Next

        SQLInsert = SQLInsert.Replace("", resInsert)
        SQLInsert = SQLInsert.TrimEnd

        Return SQLInsert
    End Function

    Public Overrides Function SQLUpdate() As String
        SQLUpdate = ""
        SQLUpdate = CCI_Engineering_Templates.My.Resources.General__UPDATE
        SQLUpdate = SQLUpdate.Replace("[TABLE NAME]", Me.EDSTableName)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID)

        Dim _pierInsert As String
        Dim _soilInsert As String

        If Me.AnchorProfile?.ID IsNot Nothing And Me.AnchorProfile?.ID > 0 Then
            _pierInsert = Me.AnchorProfile.SQLUpdate
        Else
            _pierInsert = Me.AnchorProfile.SQLInsert
        End If

        If Me.SoilProfile?.ID IsNot Nothing And Me.SoilProfile?.ID > 0 Then
            _soilInsert = Me.SoilProfile.SQLUpdate
        Else
            _soilInsert = Me.SoilProfile.SQLInsert
        End If

        SQLUpdate = SQLUpdate.Replace("--[OPTIONAL]", _pierInsert + vbCrLf + _soilInsert + vbCrLf + Me.ResultQuery(True))

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
    Private _anchor_profile_id As Integer?
    Private _anchor_block_tool_id As Integer?
    Private _soil_profile_id As Integer?
    Private _local_anchor_id As Integer?
    Private _reaction_position As Integer?
    Private _reaction_location As String
    Private _local_soil_profile_id As Integer?
    Private _local_anchor_profile_id As Integer?
    <Category("Guyed Anchor Block Profile"), Description(""), DisplayName("Anchor Profile Id")>
    Public Property anchor_profile_id() As Integer?
        Get
            Return Me._anchor_profile_id
        End Get
        Set
            Me._anchor_profile_id = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Profile"), Description(""), DisplayName("Anchor Block Tool Id")>
    Public Property anchor_block_tool_id() As Integer?
        Get
            Return Me._anchor_block_tool_id
        End Get
        Set
            Me._anchor_block_tool_id = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Profile"), Description(""), DisplayName("Soil Profile Id")>
    Public Property soil_profile_id() As Integer?
        Get
            Return Me._soil_profile_id
        End Get
        Set
            Me._soil_profile_id = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Profile"), Description(""), DisplayName("Local Anchor Id")>
    Public Property local_anchor_id() As Integer?
        Get
            Return Me._local_anchor_id
        End Get
        Set
            Me._local_anchor_id = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Profile"), Description(""), DisplayName("Reaction Position")>
    Public Property reaction_position() As Integer?
        Get
            Return Me._reaction_position
        End Get
        Set
            Me._reaction_position = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Profile"), Description(""), DisplayName("Reaction Location")>
    Public Property reaction_location() As String
        Get
            Return Me._reaction_location
        End Get
        Set
            Me._reaction_location = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Profile"), Description(""), DisplayName("Local Soil Profile Id")>
    Public Property local_soil_profile_id() As Integer?
        Get
            Return Me._local_soil_profile_id
        End Get
        Set
            Me._local_soil_profile_id = Value
        End Set
    End Property
    <Category("AnchorBlock"), Description(""), DisplayName("Local Anchor Profile Id")>
    Public Property local_anchor_profile_id() As Integer?
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
        Me.reaction_position = DBtoNullableInt(dr.Item("reaction_position"))
        Me.reaction_location = DBtoStr(dr.Item("reaction_location"))
        Me.local_soil_profile_id = DBtoNullableInt(dr.Item("local_soil_profile_id"))
        Me.local_anchor_profile_id = DBtoNullableInt(dr.Item("local_anchor_profile_id"))
        Me.ID = DBtoNullableInt(dr.Item("ID"))

    End Sub

    Public Sub New(ByVal dr As DataRow, ByVal strDS As DataSet, ByVal isExcel As Boolean, Optional ByRef Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Dim abProfile As New AnchorBlockProfile
        Dim abSProfile As New AnchorBlockSoilProfile
        Dim abLayer As New AnchorBlockSoilLayer

        ConstructMe(dr)

        For Each profileRow As DataRow In strDS.Tables(abProfile.EDSObjectName).Rows
            abProfile = New AnchorBlockProfile(profileRow, Me)
            If IsSomething(abProfile.local_anchor_profile_id) Or (IsSomething(abProfile.ID) And IsNothing(abProfile.local_anchor_profile_id)) Then
                If If(isExcel, Me.local_anchor_profile_id = abProfile.local_anchor_profile_id, Me.anchor_profile_id = abProfile.ID) Then
                    Me.AnchorProfile = abProfile
                End If
            End If
        Next

        For Each soilRow As DataRow In strDS.Tables(abSProfile.EDSObjectName).Rows
            abSProfile = (New AnchorBlockSoilProfile(soilRow, Me))
            If If(isExcel, Me.local_soil_profile_id = abSProfile.local_soil_profile_id, Me.soil_profile_id = abSProfile.ID) Then
                Me.SoilProfile = abSProfile

                For Each layerRow As DataRow In strDS.Tables(abLayer.EDSObjectName).Rows
                    abLayer = (New AnchorBlockSoilLayer(layerRow, abSProfile))
                    If If(isExcel, abSProfile.local_soil_profile_id = abLayer.local_soil_profile, Me.soil_profile_id = abLayer.Soil_Profile_id) Then
                        abSProfile.ABSoilLayers.Add(abLayer)
                    End If
                Next
            End If
        Next


        If isExcel Then
            Dim res As New EDSResult
            Dim dt As New DataTable
            dt.Columns.Add("result_lkup", GetType(String))
            dt.Columns.Add("rating", GetType(Double))

            For Each resRow In strDS.Tables("Anchor Block Result").Rows
                If DBtoNullableInt(resRow.item("local_anchor_block_id")) = Me.local_anchor_profile_id Then
                    Try
                        dt.Rows.Add("ABSOIL", Math.Round(CType(resRow.item("Soil Rating"), Double), 3))
                        dt.Rows.Add("ABSTRUC", Math.Round(CType(resRow.item("Structural Rating"), Double), 3))
                    Catch
                    End Try
                    'Exit For
                End If
            Next

            For Each resRow As DataRow In dt.Rows
                res = New EDSResult(resRow, Me)
                res.EDSTableName = "fnd.anchor_block_results"
                res.ForeignKeyName = "anchor_block_id"
                res.foreign_key = Me.ID
                res.EDSTableDepth = 1
                res.modified_person_id = Me.modified_person_id
                Dim x As AnchorBlockFoundation = Me.Parent
                x.Results.Add(res)
            Next
        End If
    End Sub

#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.anchor_block_tool_id.ToString.FormatDBValue) '@TopLevelID
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.anchor_profile_id.ToString.FormatDBValue) 'SubLevel1ID
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.soil_profile_id.ToString.FormatDBValue) 'SubLevel2ID
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
        Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
        Equals = If(Me.anchor_block_tool_id.CheckChange(otherToCompare.anchor_block_tool_id, changes, categoryName, "Anchor Block Tool Id"), Equals, False)
        Equals = If(Me.anchor_profile_id.CheckChange(otherToCompare.anchor_profile_id, changes, categoryName, "Anchor Profile Id"), Equals, False)
        Equals = If(Me.soil_profile_id.CheckChange(otherToCompare.soil_profile_id, changes, categoryName, "Soil Profile Id"), Equals, False)
        Equals = If(Me.local_anchor_id.CheckChange(otherToCompare.local_anchor_id, changes, categoryName, "Local Anchor Id"), Equals, False)
        Equals = If(Me.reaction_location.CheckChange(otherToCompare.reaction_location, changes, categoryName, "Reaction Location"), Equals, False)
        Equals = If(Me.local_anchor_profile_id.CheckChange(otherToCompare.local_anchor_profile_id, changes, categoryName, "Local Anchor Profile Id"), Equals, False)
        Equals = If(Me.local_soil_profile_id.CheckChange(otherToCompare.local_soil_profile_id, changes, categoryName, "Local Soil Profile Id"), Equals, False)
        Equals = If(Me.modified_person_id.CheckChange(otherToCompare.modified_person_id, changes, categoryName, "Modified Person Id"), Equals, False)
        Equals = If(Me.process_stage.CheckChange(otherToCompare.process_stage, changes, categoryName, "Process Stage"), Equals, False)
        'Anchor Profile
        Equals = If(Me.AnchorProfile.CheckChange(otherToCompare.AnchorProfile, changes, categoryName, "Pier Profile"), Equals, False)
        'Soil Profile
        Equals = If(Me.SoilProfile.CheckChange(otherToCompare.SoilProfile, changes, categoryName, "Soil Profile"), Equals, False)
        Return Equals
    End Function
#End Region
End Class

Partial Public Class AnchorBlockProfile
    Inherits EDSObjectWithQueries

#Region "Inheritted"
    '''Must override these inherited properties
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Anchor Block"
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
        SQLInsert = CCI_Engineering_Templates.My.Resources.General__INSERT
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

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Depth")>
    Public Property anchor_depth() As Double?
        Get
            Return Me._anchor_depth
        End Get
        Set
            Me._anchor_depth = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Width")>
    Public Property anchor_width() As Double?
        Get
            Return Me._anchor_width
        End Get
        Set
            Me._anchor_width = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Thickness")>
    Public Property anchor_thickness() As Double?
        Get
            Return Me._anchor_thickness
        End Get
        Set
            Me._anchor_thickness = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Length")>
    Public Property anchor_length() As Double?
        Get
            Return Me._anchor_length
        End Get
        Set
            Me._anchor_length = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Toe Width")>
    Public Property anchor_toe_width() As Double?
        Get
            Return Me._anchor_toe_width
        End Get
        Set
            Me._anchor_toe_width = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Top Rebar Size")>
    Public Property anchor_top_rebar_size() As Integer?
        Get
            Return Me._anchor_top_rebar_size
        End Get
        Set
            Me._anchor_top_rebar_size = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description("column - 6+local drilled pier id"), DisplayName("Anchor Top Rebar Quantity")>
    Public Property anchor_top_rebar_quantity() As Integer?
        Get
            Return Me._anchor_top_rebar_quantity
        End Get
        Set
            Me._anchor_top_rebar_quantity = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Front Rebar Size")>
    Public Property anchor_front_rebar_size() As Integer?
        Get
            Return Me._anchor_front_rebar_size
        End Get
        Set
            Me._anchor_front_rebar_size = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Front Rebar Quantity")>
    Public Property anchor_front_rebar_quantity() As Integer?
        Get
            Return Me._anchor_front_rebar_quantity
        End Get
        Set
            Me._anchor_front_rebar_quantity = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Stirrup Size")>
    Public Property anchor_stirrup_size() As Integer?
        Get
            Return Me._anchor_stirrup_size
        End Get
        Set
            Me._anchor_stirrup_size = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Shaft Diameter")>
    Public Property anchor_shaft_diameter() As Double?
        Get
            Return Me._anchor_shaft_diameter
        End Get
        Set
            Me._anchor_shaft_diameter = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Shaft Quantity")>
    Public Property anchor_shaft_quantity() As Integer?
        Get
            Return Me._anchor_shaft_quantity
        End Get
        Set
            Me._anchor_shaft_quantity = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Shaft Area Override")>
    Public Property anchor_shaft_area_override() As Double?
        Get
            Return Me._anchor_shaft_area_override
        End Get
        Set
            Me._anchor_shaft_area_override = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Shaft Shear Leg Factor")>
    Public Property anchor_shaft_shear_leg_factor() As Double?
        Get
            Return Me._anchor_shaft_shear_leg_factor
        End Get
        Set
            Me._anchor_shaft_shear_leg_factor = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Rebar Grade")>
    Public Property anchor_rebar_grade() As Double?
        Get
            Return Me._anchor_rebar_grade
        End Get
        Set
            Me._anchor_rebar_grade = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Concrete Compressive Strength")>
    Public Property concrete_compressive_strength() As Double?
        Get
            Return Me._concrete_compressive_strength
        End Get
        Set
            Me._concrete_compressive_strength = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Clear Cover")>
    Public Property clear_cover() As Double?
        Get
            Return Me._clear_cover
        End Get
        Set
            Me._clear_cover = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Shaft Yield Strength")>
    Public Property anchor_shaft_yield_strength() As Double?
        Get
            Return Me._anchor_shaft_yield_strength
        End Get
        Set
            Me._anchor_shaft_yield_strength = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Shaft Ultimate Strength")>
    Public Property anchor_shaft_ultimate_strength() As Double?
        Get
            Return Me._anchor_shaft_ultimate_strength
        End Get
        Set
            Me._anchor_shaft_ultimate_strength = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Rebar Known")>
    Public Property rebar_known() As Boolean?
        Get
            Return Me._rebar_known
        End Get
        Set
            Me._rebar_known = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Shaft Known")>
    Public Property anchor_shaft_known() As Boolean?
        Get
            Return Me._anchor_shaft_known
        End Get
        Set
            Me._anchor_shaft_known = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Basic Soil Check")>
    Public Property basic_soil_check() As Boolean?
        Get
            Return Me._basic_soil_check
        End Get
        Set
            Me._basic_soil_check = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Structural Check")>
    Public Property structural_check() As Boolean?
        Get
            Return Me._structural_check
        End Get
        Set
            Me._structural_check = Value
        End Set
    End Property
    <Category("AnchorBlockProfile"), Description(""), DisplayName("Local Anchor Profile Id")>
    Public Property local_anchor_profile_id() As Integer?
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

    Public Sub New(ByVal dr As DataRow, Optional ByRef Parent As EDSObject = Nothing)
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
        Me.ID = DBtoNullableInt(dr.Item("ID"))
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
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.anchor_shaft_shear_lag_factor.ToString.FormatDBValue)
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
        SQLInsertFields = SQLInsertFields.AddtoDBString("anchor_shaft_shear_lag_factor")
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
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("anchor_shaft_shear_lag_factor = " & Me.anchor_shaft_shear_lag_factor.ToString.FormatDBValue)
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
        Equals = If(Me.anchor_shaft_shear_lag_factor.CheckChange(otherToCompare.anchor_shaft_shear_lag_factor, changes, categoryName, "Anchor Shaft Shear Lag Factor"), Equals, False)
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

Partial Public Class AnchorBlockSoilProfile
    Inherits SoilProfile

    Public Property local_soil_profile_id As Integer?
    Public Property ABSoilLayers As New List(Of AnchorBlockSoilLayer)

    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Anchor Block Foundation"
        End Get
    End Property

    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "fnd.anchor_block_tool"
        End Get
    End Property

    Public Sub New()

    End Sub

    Public Sub New(ByVal dr As DataRow, Optional ByRef Parent As EDSObject = Nothing)
        ConstructMe(dr, Parent)

        Try
            Me.local_soil_profile_id = DBtoNullableDbl(dr.Item("local_soil_profile_id"))
        Catch
        End Try
    End Sub


    Public Overrides Function SQLUpdate() As String

        SQLUpdate = QueryBuilderFromFile(queryPath & "Soil Profile\Soil Profile (UPDATE).sql")
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

Partial Public Class AnchorBlockSoilLayer
    Inherits SoilLayer

    Public Property local_soil_profile As Integer?
    Public Property local_soil_layer_id As Integer?

    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Anchor Block Soil Layer"
        End Get
    End Property

    Public Sub New()

    End Sub

    Public Sub New(ByVal dr As DataRow, Optional ByRef Parent As EDSObject = Nothing)
        ConstructMe(dr, Parent)

        Try
            Me.local_soil_profile = DBtoNullableDbl(dr.Item("local_soil_profile"))
            Me.local_soil_layer_id = DBtoNullableDbl(dr.Item("local_soil_layer_id"))
        Catch
        End Try
    End Sub

End Class

Partial Public Class AnchorBlockResult
    Inherits EDSResult

    Public ReadOnly Property EDSObjectName As String
        Get
            Return "Anchor Block Foundation"
        End Get
    End Property

    Public ReadOnly Property EDSTableName As String
        Get
            Return "fnd.anchor_block_tool"
        End Get
    End Property

End Class
