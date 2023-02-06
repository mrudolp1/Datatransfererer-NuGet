Option Strict Off
Option Compare Binary

Imports System.ComponentModel
Imports System.Data
Imports DevExpress.Spreadsheet
'Imports Microsoft.Office.Interop

Partial Public Class AnchorBlockFoundation
    Inherits EDSExcelObject

    Public Property GuyedAnchorBlocks As New List(Of AnchorBlock)

    'Origin row in the driled pier database. Basically just where the profile numbers are in the database worksheet.
    'This is actually 58 but due to the 0,0 origin in excel, it is 1 less
    Private pierProfileRow As Integer = 57

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
        If _Insert = "" Then
            '_Insert = CCI_Engineering_Templates.My.Resources.CCIpole_General_INSERT
        End If
        SQLInsert = _Insert

        'Guy Anchors

        'Anchor Profiles

        'Soil Profiles

        'Soil Layers

        'Results

    End Function

    Public Overrides Function SQLUpdate() As String
        If _Update = "" Then
            '_Update = CCI_Engineering_Templates.My.Resources.CCIpole_General_UPDATE
        End If
        SQLUpdate = _Update

        'Guy Anchors

        'Anchor Profiles

        'Soil Profiles

        'Soil Layers

        'Results

    End Function

    Public Overrides Function SQLDelete() As String
        If _Delete = "" Then
            '_Delete = CCI_Engineering_Templates.My.Resources.CCIpole_General_DELETE
        End If
        SQLDelete = _Delete

        'Guy Anchors

        'Anchor Profiles

        'Soil Profiles

        'Soil Layers

        'Results

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
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        'SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bus_unit")
        SQLInsertFields = SQLInsertFields.AddtoDBString("structure_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("file_ver")
        SQLInsertFields = SQLInsertFields.AddtoDBString("modified")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""

        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bus_unit = " & Me.bus_unit.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("structure_id = " & Me.structure_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("file_ver = " & Me.file_ver.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified = " & Me.modified.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_person_id = " & Me.modified_person_id.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("process_stage = " & Me.process_stage.ToString.FormatDBValue)

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
                .Worksheets("Sub Tables (SAPI)").Range("48" & i).Value = CType(Me.ID, Integer)
            Else
                .Worksheets("Sub Tables (SAPI)").Range("48" & i).ClearContents
            End If
            If Not IsNothing(Me.bus_unit) Then
                .Worksheets("").Range("" & i).Value = CType(Me.bus_unit, Integer)
            Else
                .Worksheets("").Range("" & i).ClearContents
            End If
            If Not IsNothing(Me.structure_id) Then
                .Worksheets("").Range("" & i).Value = CType(Me.structure_id, String)
            End If
            'If Not IsNothing(Me.file_ver) Then
            '    .Worksheets("").Range("" & i).Value = CType(Me.file_ver, String)
            'End If
            If Not IsNothing(Me.modified) Then
                .Worksheets("").Range("" & i).Value = CType(Me.modified, Boolean)
            End If
            'If Not IsNothing(Me.modified_person_id) Then
            '    .Worksheets("").Range("" & i).Value = CType(Me.modified_person_id, Integer)
            'Else
            '    .Worksheets("").Range("" & i).ClearContents
            'End If
            'If Not IsNothing(Me.process_stage) Then
            '    .Worksheets("").Range("" & i).Value = CType(Me.process_stage, String)
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
            .Worksheets("General (SAPI)").Range("P3").Value = CType(gab_tia_current, String)
            'Load Z Normalization
            'If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.load_z_norm) Then
            '    rev_h_section_15_5 = Me.ParentStructure?.structureCodeCriteria?.load_z_norm
            '    .Worksheets("General (SAPI)").Range("Q3").Value = CType(load_z_norm, Boolean)
            'End If
            'H Section 15.5
            If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.rev_h_section_15_5) Then
                gab_rev_h_section_15_5 = Me.ParentStructure?.structureCodeCriteria?.rev_h_section_15_5
                .Worksheets("General (SAPI)").Range("R3").Value = CType(gab_rev_h_section_15_5, Boolean)
            End If
            'Work Order
            If Not IsNothing(Me.ParentStructure?.work_order_seq_num) Then
                work_order_seq_num = Me.ParentStructure?.work_order_seq_num
                .Worksheets("General (SAPI)").Range("S3").Value = CType(work_order_seq_num, Integer)
            End If

            'Anchors
            For Each gab As AnchorBlock In GuyedAnchorBlocks
                'TBD - See Drilled Pier Object for similar config
                'Profile
                'Soil Profile
                'Soil Layers
                Dim layAdj As Integer = 0
                For Each layer In AnchorBlock.SoilProfile.GABSoilLayers
                    If Not IsNothing(layer.ID) Then .Cells(pierProfileRow + layAdj - 35, myCol).Value = CType(layer.ID, Double)
                    If Not IsNothing(layer.bottom_depth) Then .Cells(pierProfileRow + layAdj + 127, myCol).Value = CType(layer.bottom_depth, Double)
                    If Not IsNothing(layer.effective_soil_density) Then .Cells(pierProfileRow + layAdj + 158, myCol).Value = CType(layer.effective_soil_density, Double)
                    If Not IsNothing(layer.cohesion) Then .Cells(pierProfileRow + layAdj + 189, myCol).Value = CType(layer.cohesion, Double)
                    If Not IsNothing(layer.friction_angle) Then .Cells(pierProfileRow + layAdj + 220, myCol).Value = CType(layer.friction_angle, Double)
                    If Not IsNothing(layer.skin_friction_override_comp) Then .Cells(pierProfileRow + layAdj + 251, myCol).Value = CType(layer.skin_friction_override_comp, Double)
                    If Not IsNothing(layer.skin_friction_override_uplift) Then .Cells(pierProfileRow + layAdj + 282, myCol).Value = CType(layer.skin_friction_override_uplift, Double)
                    If Not IsNothing(layer.nominal_bearing_capacity) Then .Cells(pierProfileRow + layAdj + 313, myCol).Value = CType(layer.nominal_bearing_capacity, Double)
                    If Not IsNothing(layer.spt_blow_count) Then .Cells(pierProfileRow + layAdj + 344, myCol).Value = CType(layer.spt_blow_count, Double)
                    layAdj += 1
                Next
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
        Equals = If(Me.GuyedAnchorBlocks.CheckChange(otherToCompare.GuyedAnchorBlocks, changes, categoryName, "Anchor Blocks"), Equals, False)

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
        Return SQLInsert
    End Function

    Public Overrides Function SQLUpdate() As String
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
    Private _ID As Integer?
    Private _anchor_block_profile_id As Integer?
    Private _soil_profile_id As Integer?
    Private _reaction_position As Integer?
    Private _reaction_location As String
    Private _local_anchor_profile As Integer?
    Private _local_soil_profile As Integer?
    <Category("Guyed Anchor Block Profile"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer?
        Get
            Return Me._ID
        End Get
        Set
            Me._ID = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Profile"), Description(""), DisplayName("Anchor Block Profile Id")>
    Public Property anchor_block_profile_id() As Integer?
        Get
            Return Me._anchor_block_profile_id
        End Get
        Set
            Me._anchor_block_profile_id = Value
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
    <Category("Guyed Anchor Block Profile"), Description(""), DisplayName("Local Anchor Profile")>
    Public Property local_anchor_profile() As Integer?
        Get
            Return Me._local_anchor_profile
        End Get
        Set
            Me._local_anchor_profile = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Profile"), Description(""), DisplayName("Local Soil Profile")>
    Public Property local_soil_profile() As Integer?
        Get
            Return Me._local_soil_profile
        End Get
        Set
            Me._local_soil_profile = Value
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
    End Sub

    Public Sub New(ByVal dr As DataRow, ByVal strDS As DataSet, ByVal isExcel As Boolean, Optional ByRef Parent As EDSObject = Nothing)
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
        Return SQLInsert
    End Function

    Public Overrides Function SQLUpdate() As String
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
    Private _ID As Integer?
    Private _anchor_location As String
    Private _guy_anchor_radius As Double?
    Private _anchor_depth As Double?
    Private _anchor_width As Double?
    Private _anchor_thickness As Double?
    Private _anchor_length As Double?
    Private _anchor_toe_width As Double?
    Private _anchor_top_rebar_size As Integer?
    Private _anchor_top_rebar_quantity As Integer?
    Private _anchor_bottom_rebar_size As Integer?
    Private _anchor_bottom_rebar_quantity As Integer?
    Private _anchor_stirrup_size As Integer?
    Private _anchor_shaft_diameter As Double?
    Private _anchor_shaft_quantity As Integer?
    Private _anchor_shaft_area_override As Double?
    Private _anchor_shaft_shear_leg_factor As Double?
    Private _rebar_grade As Double?
    Private _concrete_compressive_strength As Double?
    Private _clear_cover As Double?
    Private _anchor_shaft_yield_strength As Double?
    Private _anchor_shaft_ultimate_strength As Double?
    Private _tool_version As String
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer?
        Get
            Return Me._ID
        End Get
        Set
            Me._ID = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Location")>
    Public Property anchor_location() As String
        Get
            Return Me._anchor_location
        End Get
        Set
            Me._anchor_location = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Guy Anchor Radius")>
    Public Property guy_anchor_radius() As Double?
        Get
            Return Me._guy_anchor_radius
        End Get
        Set
            Me._guy_anchor_radius = Value
        End Set
    End Property
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
    <Category("Guyed Anchor Block Details"), Description("column - 6+local drilled pier id"), DisplayName("Anchor Length")>
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
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Top Rebar Quantity")>
    Public Property anchor_top_rebar_quantity() As Integer?
        Get
            Return Me._anchor_top_rebar_quantity
        End Get
        Set
            Me._anchor_top_rebar_quantity = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Bottom Rebar Size")>
    Public Property anchor_bottom_rebar_size() As Integer?
        Get
            Return Me._anchor_bottom_rebar_size
        End Get
        Set
            Me._anchor_bottom_rebar_size = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Bottom Rebar Quantity")>
    Public Property anchor_bottom_rebar_quantity() As Integer?
        Get
            Return Me._anchor_bottom_rebar_quantity
        End Get
        Set
            Me._anchor_bottom_rebar_quantity = Value
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
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Rebar Grade")>
    Public Property rebar_grade() As Double?
        Get
            Return Me._rebar_grade
        End Get
        Set
            Me._rebar_grade = Value
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
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Tool Version")>
    Public Property tool_version() As String
        Get
            Return Me._tool_version
        End Get
        Set
            Me._tool_version = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()

    End Sub

    Public Sub New(ByVal dr As DataRow, Optional ByRef Parent As EDSObject = Nothing)
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

End Class

Partial Public Class AnchorBlockSoilLayer
    Inherits SoilLayer

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
