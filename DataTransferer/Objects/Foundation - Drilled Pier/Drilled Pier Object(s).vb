Option Strict On

Imports System.ComponentModel
Imports System.Data
Imports DevExpress.Spreadsheet
'Imports Microsoft.Office.Interop

Partial Public Class DrilledPier
    Inherits EDSExcelObject

    Public Property PierProfiles As New List(Of DrilledPierProfile)
    Public Property SoilProfiles As New List(Of SoilProfile)


#Region "Inheritted"
    '''Must override these inherited properties
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Drilled Pier"
        End Get
    End Property
    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "fnd.drilled_pier"
        End Get
    End Property
    Public Overrides ReadOnly Property templatePath As String
        Get
            Return IO.Path.Combine(My.Application.Info.DirectoryPath, "Templates", "Drilled Pier Foundation (5.0.4).xlsm")
        End Get
    End Property
    Public Overrides ReadOnly Property excelDTParams As List(Of EXCELDTParameter)
        Get
            Return New List(Of EXCELDTParameter) From {
                                                        New EXCELDTParameter("Drilled Pier EXCEL", "A1:BC2", "Details (SAPI)"),
                                                        New EXCELDTParameter("Drilled Pier Profile EXCEL", "B3:D4", "Sub Tables (SAPI)"),
                                                        New EXCELDTParameter("Drilled Pier Section EXCEL", "G3:O17", "Sub Tables (SAPI)"),
                                                        New EXCELDTParameter("Drilled Pier Rebar EXCEL", "R3:T103", "Sub Tables (SAPI)"),
                                                        New EXCELDTParameter("Drilled Pier Soil Profile EXCEL", "R3:T103", "Sub Tables (SAPI)"),
                                                        New EXCELDTParameter("Drilled Pier Soil Layer EXCEL", "R3:T103", "Sub Tables (SAPI)"),
                                                        New EXCELDTParameter("Embedded Pole EXCEL", "R3:T103", "Sub Tables (SAPI)"),
                                                        New EXCELDTParameter("Belled Pier EXCEL", "A1:C17", "Sub Tables (SAPI)"),
                                                        New EXCELDTParameter("Drilled Pier Results EXCEL", "A1:C17", "Results (SAPI)")
                                                                                        }
            '***Add additional table references here****
        End Get
    End Property

    Private _Insert As String
    Private _Update As String
    Private _Delete As String

    Public Overrides Function SQLInsert() As String
        Return SQLInsert
    End Function

    Public Overrides Function SQLUpdate() As String
        'This section not only needs to call update commands but also needs to call insert and delete commands since subtables may involve adding or deleting records
        Return SQLUpdate
    End Function

    Public Overrides Function SQLDelete() As String
        Return SQLDelete
    End Function

#End Region

#Region "Define"
    Private _ID As Integer?
    Private _drilled_pier_profile_id As Integer?
    Private _soil_profile_id As Integer?
    Private _reaction_position As Integer?
    Private _reaction_location As String
    Private _local_soil_profile As Integer?
    Private _local_drilled_pier_profile As Integer?
    <Category("Drilled Pier"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer?
        Get
            Return Me._ID
        End Get
        Set
            Me._ID = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Drilled Pier Profile Id")>
    Public Property drilled_pier_profile_id() As Integer?
        Get
            Return Me._drilled_pier_profile_id
        End Get
        Set
            Me._drilled_pier_profile_id = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Soil Profile Id")>
    Public Property soil_profile_id() As Integer?
        Get
            Return Me._soil_profile_id
        End Get
        Set
            Me._soil_profile_id = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Reaction Position")>
    Public Property reaction_position() As Integer?
        Get
            Return Me._reaction_position
        End Get
        Set
            Me._reaction_position = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Reaction Location")>
    Public Property reaction_location() As String
        Get
            Return Me._reaction_location
        End Get
        Set
            Me._reaction_location = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Local Soil Profile")>
    Public Property local_soil_profile() As Integer?
        Get
            Return Me._local_soil_profile
        End Get
        Set
            Me._local_soil_profile = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Local Drilled Pier Profile")>
    Public Property local_drilled_pier_profile() As Integer?
        Get
            Return Me._local_drilled_pier_profile
        End Get
        Set
            Me._local_drilled_pier_profile = Value
        End Set
    End Property
#End Region

#Region "Constructors"
    Public Sub New()

    End Sub

    Public Sub New(ByVal dr As DataRow)
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        Me.drilled_pier_profile_id = DBtoNullableInt(dr.Item("drilled_pier_profile_id"))
        Me.soil_profile_id = DBtoNullableInt(dr.Item("soil_profile_id"))
        Me.reaction_position = DBtoNullableInt(dr.Item("reaction_position"))
        Me.reaction_location = DBtoStr(dr.Item("reaction_location"))
        Me.local_soil_profile = DBtoNullableInt(dr.Item("local_soil_profile"))
        Me.local_drilled_pier_profile = DBtoNullableInt(dr.Item("local_drilled_pier_profile"))

        For Each drr As DataRow In ds.Tables("Drilled Pier Profile").Rows

        Next

        For Each drr As DataRow In ds.Tables("Drilled Pier Section").Rows

        Next

        For Each drr As DataRow In ds.Tables("Drilled Pier Rebar").Rows

        Next

        For Each drr As DataRow In ds.Tables("Embedded Pole").Rows

        Next

        For Each drr As DataRow In ds.Tables("Belled Pier").Rows

        Next

        For Each drr As DataRow In ds.Tables("Soil Profile").Rows

        Next

        For Each drr As DataRow In ds.Tables("Soil Layer").Rows

        Next
    End Sub
#End Region

#Region "Save to Excel"
    Public Overrides Sub workBookFiller(ByRef wb As Workbook)
        '''''Customize for each excel tool'''''

        'Site Code Criteria
        Dim tia_current, site_name, structure_type As String
        Dim rev_h_section_15_5 As Boolean?

        With wb

        End With

    End Sub
#End Region

#Region "Save to EDS"

    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.drilled_pier_profile_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.soil_profile_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reaction_position.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reaction_location.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_soil_profile.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_drilled_pier_profile.ToString.FormatDBValue)
        Return SQLInsertValues

    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
        SQLInsertFields = SQLInsertFields.AddtoDBString("drilled_pier_profile_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("soil_profile_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("reaction_position")
        SQLInsertFields = SQLInsertFields.AddtoDBString("reaction_location")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_soil_profile")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_drilled_pier_profile")
        Return SQLInsertFields

    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("drilled_pier_profile_id = " & Me.drilled_pier_profile_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("soil_profile_id = " & Me.soil_profile_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("reaction_position = " & Me.reaction_position.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("reaction_location = " & Me.reaction_location.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_soil_profile = " & Me.local_soil_profile.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_drilled_pier_profile = " & Me.local_drilled_pier_profile.ToString.FormatDBValue)
        Return SQLUpdateFieldsandValues

    End Function

#End Region

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As DrilledPier = TryCast(other, DrilledPier)
        If otherToCompare Is Nothing Then Return False
        Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
        Equals = If(Me.drilled_pier_profile_id.CheckChange(otherToCompare.drilled_pier_profile_id, changes, categoryName, "Drilled Pier Profile Id"), Equals, False)
        Equals = If(Me.soil_profile_id.CheckChange(otherToCompare.soil_profile_id, changes, categoryName, "Soil Profile Id"), Equals, False)
        Equals = If(Me.reaction_position.CheckChange(otherToCompare.reaction_position, changes, categoryName, "Reaction Position"), Equals, False)
        Equals = If(Me.reaction_location.CheckChange(otherToCompare.reaction_location, changes, categoryName, "Reaction Location"), Equals, False)
        Equals = If(Me.local_soil_profile.CheckChange(otherToCompare.local_soil_profile, changes, categoryName, "Local Soil Profile"), Equals, False)
        Equals = If(Me.local_drilled_pier_profile.CheckChange(otherToCompare.local_drilled_pier_profile, changes, categoryName, "Local Drilled Pier Profile"), Equals, False)
        Return Equals

    End Function

End Class

Partial Public Class DrilledPierProfile
    Inherits EDSObjectWithQueries

    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Drilled Pier Profile"
        End Get
    End Property

    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "fnd.drilled_pier_profile"
        End Get
    End Property

#Region "Define"
    Public Property Sections As New List(Of DrilledPierSection)
    Public Property EmbeddedPoles As New List(Of EmbeddedPole)
    Public Property BelledPiers As New List(Of BelledPier)

    Private _ID As Integer?
    Private _foundation_depth As Double?
    Private _extension_above_grade As Double?
    Private _assume_min_steel As Boolean?
    Private _check_shear_along_depth As Boolean?
    Private _utilize_skin_friction_methodology As Boolean?
    Private _embedded_pole As Boolean?
    Private _belled_pier As Boolean?
    Private _concrete_compressive_strength As Double?
    Private _longitudinal_rebar_yield_strength As Double?
    Private _rebar_cage_2_fy_override As Double?
    Private _rebar_cage_3_fy_override As Double?
    Private _rebar_effective_depths As Boolean?
    Private _shear_crit_depth_override_comp As Double?
    Private _shear_crit_depth_override_uplift As Double?
    Private _shear_override_crit_depth As Boolean?
    Private _tie_yield_strength As Double?
    Private _tool_version As String
    <Category("Drilled Pier"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer?
        Get
            Return Me._ID
        End Get
        Set
            Me._ID = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Foundation Depth")>
    Public Property foundation_depth() As Double?
        Get
            Return Me._foundation_depth
        End Get
        Set
            Me._foundation_depth = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Extension Above Grade")>
    Public Property extension_above_grade() As Double?
        Get
            Return Me._extension_above_grade
        End Get
        Set
            Me._extension_above_grade = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Assume Min Steel")>
    Public Property assume_min_steel() As Boolean?
        Get
            Return Me._assume_min_steel
        End Get
        Set
            Me._assume_min_steel = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Check Shear Along Depth")>
    Public Property check_shear_along_depth() As Boolean?
        Get
            Return Me._check_shear_along_depth
        End Get
        Set
            Me._check_shear_along_depth = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Utilize Skin Friction Methodology")>
    Public Property utilize_skin_friction_methodology() As Boolean?
        Get
            Return Me._utilize_skin_friction_methodology
        End Get
        Set
            Me._utilize_skin_friction_methodology = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Embedded Pole")>
    Public Property embedded_pole() As Boolean?
        Get
            Return Me._embedded_pole
        End Get
        Set
            Me._embedded_pole = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Belled Pier")>
    Public Property belled_pier() As Boolean?
        Get
            Return Me._belled_pier
        End Get
        Set
            Me._belled_pier = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Concrete Compressive Strength")>
    Public Property concrete_compressive_strength() As Double?
        Get
            Return Me._concrete_compressive_strength
        End Get
        Set
            Me._concrete_compressive_strength = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Longitudinal Rebar Yield Strength")>
    Public Property longitudinal_rebar_yield_strength() As Double?
        Get
            Return Me._longitudinal_rebar_yield_strength
        End Get
        Set
            Me._longitudinal_rebar_yield_strength = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Rebar Cage 2 Fy Override")>
    Public Property rebar_cage_2_fy_override() As Double?
        Get
            Return Me._rebar_cage_2_fy_override
        End Get
        Set
            Me._rebar_cage_2_fy_override = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Rebar Cage 3 Fy Override")>
    Public Property rebar_cage_3_fy_override() As Double?
        Get
            Return Me._rebar_cage_3_fy_override
        End Get
        Set
            Me._rebar_cage_3_fy_override = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Rebar Effective Depths")>
    Public Property rebar_effective_depths() As Boolean?
        Get
            Return Me._rebar_effective_depths
        End Get
        Set
            Me._rebar_effective_depths = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Shear Crit Depth Override Comp")>
    Public Property shear_crit_depth_override_comp() As Double?
        Get
            Return Me._shear_crit_depth_override_comp
        End Get
        Set
            Me._shear_crit_depth_override_comp = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Shear Crit Depth Override Uplift")>
    Public Property shear_crit_depth_override_uplift() As Double?
        Get
            Return Me._shear_crit_depth_override_uplift
        End Get
        Set
            Me._shear_crit_depth_override_uplift = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Shear Override Crit Depth")>
    Public Property shear_override_crit_depth() As Boolean?
        Get
            Return Me._shear_override_crit_depth
        End Get
        Set
            Me._shear_override_crit_depth = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Tie Yield Strength")>
    Public Property tie_yield_strength() As Double?
        Get
            Return Me._tie_yield_strength
        End Get
        Set
            Me._tie_yield_strength = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Tool Version")>
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

    Public Sub New(ByVal dr As DataRow)
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        Me.foundation_depth = DBtoNullableDbl(dr.Item("foundation_depth"))
        Me.extension_above_grade = DBtoNullableDbl(dr.Item("extension_above_grade"))
        Me.assume_min_steel = DBtoNullableBool(dr.Item("assume_min_steel"))
        Me.check_shear_along_depth = DBtoNullableBool(dr.Item("check_shear_along_depth"))
        Me.utilize_skin_friction_methodology = DBtoNullableBool(dr.Item("utilize_skin_friction_methodology"))
        Me.embedded_pole = DBtoNullableBool(dr.Item("embedded_pole"))
        Me.belled_pier = DBtoNullableBool(dr.Item("belled_pier"))
        Me.concrete_compressive_strength = DBtoNullableDbl(dr.Item("concrete_compressive_strength"))
        Me.longitudinal_rebar_yield_strength = DBtoNullableDbl(dr.Item("longitudinal_rebar_yield_strength"))
        Me.rebar_cage_2_fy_override = DBtoNullableDbl(dr.Item("rebar_cage_2_fy_override"))
        Me.rebar_cage_3_fy_override = DBtoNullableDbl(dr.Item("rebar_cage_3_fy_override"))
        Me.rebar_effective_depths = DBtoNullableBool(dr.Item("rebar_effective_depths"))
        Me.shear_crit_depth_override_comp = DBtoNullableDbl(dr.Item("shear_crit_depth_override_comp"))
        Me.shear_crit_depth_override_uplift = DBtoNullableDbl(dr.Item("shear_crit_depth_override_uplift"))
        Me.shear_override_crit_depth = DBtoNullableBool(dr.Item("shear_override_crit_depth"))
        Me.tie_yield_strength = DBtoNullableDbl(dr.Item("tie_yield_strength"))
        Me.tool_version = DBtoStr(dr.Item("tool_version"))

    End Sub
#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.foundation_depth.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.extension_above_grade.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.assume_min_steel.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.check_shear_along_depth.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.utilize_skin_friction_methodology.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.embedded_pole.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.belled_pier.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.concrete_compressive_strength.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.longitudinal_rebar_yield_strength.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rebar_cage_2_fy_override.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rebar_cage_3_fy_override.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rebar_effective_depths.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.shear_crit_depth_override_comp.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.shear_crit_depth_override_uplift.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.shear_override_crit_depth.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.tie_yield_strength.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.tool_version.ToString.FormatDBValue)
        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
        SQLInsertFields = SQLInsertFields.AddtoDBString("foundation_depth")
        SQLInsertFields = SQLInsertFields.AddtoDBString("extension_above_grade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("assume_min_steel")
        SQLInsertFields = SQLInsertFields.AddtoDBString("check_shear_along_depth")
        SQLInsertFields = SQLInsertFields.AddtoDBString("utilize_skin_friction_methodology")
        SQLInsertFields = SQLInsertFields.AddtoDBString("embedded_pole")
        SQLInsertFields = SQLInsertFields.AddtoDBString("belled_pier")
        SQLInsertFields = SQLInsertFields.AddtoDBString("concrete_compressive_strength")
        SQLInsertFields = SQLInsertFields.AddtoDBString("longitudinal_rebar_yield_strength")
        SQLInsertFields = SQLInsertFields.AddtoDBString("rebar_cage_2_fy_override")
        SQLInsertFields = SQLInsertFields.AddtoDBString("rebar_cage_3_fy_override")
        SQLInsertFields = SQLInsertFields.AddtoDBString("rebar_effective_depths")
        SQLInsertFields = SQLInsertFields.AddtoDBString("shear_crit_depth_override_comp")
        SQLInsertFields = SQLInsertFields.AddtoDBString("shear_crit_depth_override_uplift")
        SQLInsertFields = SQLInsertFields.AddtoDBString("shear_override_crit_depth")
        SQLInsertFields = SQLInsertFields.AddtoDBString("tie_yield_strength")
        SQLInsertFields = SQLInsertFields.AddtoDBString("tool_version")
        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("foundation_depth = " & Me.foundation_depth.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("extension_above_grade = " & Me.extension_above_grade.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("assume_min_steel = " & Me.assume_min_steel.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("check_shear_along_depth = " & Me.check_shear_along_depth.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("utilize_skin_friction_methodology = " & Me.utilize_skin_friction_methodology.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("embedded_pole = " & Me.embedded_pole.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("belled_pier = " & Me.belled_pier.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("concrete_compressive_strength = " & Me.concrete_compressive_strength.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("longitudinal_rebar_yield_strength = " & Me.longitudinal_rebar_yield_strength.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("rebar_cage_2_fy_override = " & Me.rebar_cage_2_fy_override.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("rebar_cage_3_fy_override = " & Me.rebar_cage_3_fy_override.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("rebar_effective_depths = " & Me.rebar_effective_depths.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("shear_crit_depth_override_comp = " & Me.shear_crit_depth_override_comp.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("shear_crit_depth_override_uplift = " & Me.shear_crit_depth_override_uplift.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("shear_override_crit_depth = " & Me.shear_override_crit_depth.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("tie_yield_strength = " & Me.tie_yield_strength.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("tool_version = " & Me.tool_version.ToString.FormatDBValue)
        Return SQLUpdateFieldsandValues
    End Function
#End Region

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As DrilledPierProfile = TryCast(other, DrilledPierProfile)
        If otherToCompare Is Nothing Then Return False
        Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
        Equals = If(Me.foundation_depth.CheckChange(otherToCompare.foundation_depth, changes, categoryName, "Foundation Depth"), Equals, False)
        Equals = If(Me.extension_above_grade.CheckChange(otherToCompare.extension_above_grade, changes, categoryName, "Extension Above Grade"), Equals, False)
        Equals = If(Me.assume_min_steel.CheckChange(otherToCompare.assume_min_steel, changes, categoryName, "Assume Min Steel"), Equals, False)
        Equals = If(Me.check_shear_along_depth.CheckChange(otherToCompare.check_shear_along_depth, changes, categoryName, "Check Shear Along Depth"), Equals, False)
        Equals = If(Me.utilize_skin_friction_methodology.CheckChange(otherToCompare.utilize_skin_friction_methodology, changes, categoryName, "Utilize Skin Friction Methodology"), Equals, False)
        Equals = If(Me.embedded_pole.CheckChange(otherToCompare.embedded_pole, changes, categoryName, "Embedded Pole"), Equals, False)
        Equals = If(Me.belled_pier.CheckChange(otherToCompare.belled_pier, changes, categoryName, "Belled Pier"), Equals, False)
        Equals = If(Me.concrete_compressive_strength.CheckChange(otherToCompare.concrete_compressive_strength, changes, categoryName, "Concrete Compressive Strength"), Equals, False)
        Equals = If(Me.longitudinal_rebar_yield_strength.CheckChange(otherToCompare.longitudinal_rebar_yield_strength, changes, categoryName, "Longitudinal Rebar Yield Strength"), Equals, False)
        Equals = If(Me.rebar_cage_2_fy_override.CheckChange(otherToCompare.rebar_cage_2_fy_override, changes, categoryName, "Rebar Cage 2 Fy Override"), Equals, False)
        Equals = If(Me.rebar_cage_3_fy_override.CheckChange(otherToCompare.rebar_cage_3_fy_override, changes, categoryName, "Rebar Cage 3 Fy Override"), Equals, False)
        Equals = If(Me.rebar_effective_depths.CheckChange(otherToCompare.rebar_effective_depths, changes, categoryName, "Rebar Effective Depths"), Equals, False)
        Equals = If(Me.shear_crit_depth_override_comp.CheckChange(otherToCompare.shear_crit_depth_override_comp, changes, categoryName, "Shear Crit Depth Override Comp"), Equals, False)
        Equals = If(Me.shear_crit_depth_override_uplift.CheckChange(otherToCompare.shear_crit_depth_override_uplift, changes, categoryName, "Shear Crit Depth Override Uplift"), Equals, False)
        Equals = If(Me.shear_override_crit_depth.CheckChange(otherToCompare.shear_override_crit_depth, changes, categoryName, "Shear Override Crit Depth"), Equals, False)
        Equals = If(Me.tie_yield_strength.CheckChange(otherToCompare.tie_yield_strength, changes, categoryName, "Tie Yield Strength"), Equals, False)
        Equals = If(Me.tool_version.CheckChange(otherToCompare.tool_version, changes, categoryName, "Tool Version"), Equals, False)
        Return Equals

    End Function
End Class

Partial Public Class DrilledPierSection
    Inherits EDSObjectWithQueries

    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Drilled Pier Section"
        End Get
    End Property

    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "fnd.drilled_pier_section"
        End Get
    End Property

#Region "Define"
    Public Property Rebar As New List(Of DrilledPierRebar)

    Private _ID As Integer?
    Private _drilled_pier_profile_id As Integer?
    Private _pier_diameter As Double?
    Private _clear_cover As Double?
    Private _clear_cover_rebar_cage_option As Boolean?
    Private _tie_size As Integer?
    Private _tie_spacing As Double?
    Private _bottom_elevation As Double?
    Private _local_section_id As Integer?
    Private _rho_override As Double?
    <Category("Drilled Pier"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer?
        Get
            Return Me._ID
        End Get
        Set
            Me._ID = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Drilled Pier Profile Id")>
    Public Property drilled_pier_profile_id() As Integer?
        Get
            Return Me._drilled_pier_profile_id
        End Get
        Set
            Me._drilled_pier_profile_id = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Pier Diameter")>
    Public Property pier_diameter() As Double?
        Get
            Return Me._pier_diameter
        End Get
        Set
            Me._pier_diameter = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Clear Cover")>
    Public Property clear_cover() As Double?
        Get
            Return Me._clear_cover
        End Get
        Set
            Me._clear_cover = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Clear Cover Rebar Cage Option")>
    Public Property clear_cover_rebar_cage_option() As Boolean?
        Get
            Return Me._clear_cover_rebar_cage_option
        End Get
        Set
            Me._clear_cover_rebar_cage_option = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Tie Size")>
    Public Property tie_size() As Integer?
        Get
            Return Me._tie_size
        End Get
        Set
            Me._tie_size = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Tie Spacing")>
    Public Property tie_spacing() As Double?
        Get
            Return Me._tie_spacing
        End Get
        Set
            Me._tie_spacing = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Bottom Elevation")>
    Public Property bottom_elevation() As Double?
        Get
            Return Me._bottom_elevation
        End Get
        Set
            Me._bottom_elevation = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Local Section Id")>
    Public Property local_section_id() As Integer?
        Get
            Return Me._local_section_id
        End Get
        Set
            Me._local_section_id = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Rho Override")>
    Public Property rho_override() As Double?
        Get
            Return Me._rho_override
        End Get
        Set
            Me._rho_override = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()

    End Sub

    Public Sub New(ByVal dr As DataRow)
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        Me.drilled_pier_profile_id = DBtoNullableInt(dr.Item("drilled_pier_profile_id"))
        Me.pier_diameter = DBtoNullableDbl(dr.Item("pier_diameter"))
        Me.clear_cover = DBtoNullableDbl(dr.Item("clear_cover"))
        Me.clear_cover_rebar_cage_option = DBtoNullableBool(dr.Item("clear_cover_rebar_cage_option"))
        Me.tie_size = DBtoNullableInt(dr.Item("tie_size"))
        Me.tie_spacing = DBtoNullableDbl(dr.Item("tie_spacing"))
        Me.bottom_elevation = DBtoNullableDbl(dr.Item("bottom_elevation"))
        Me.local_section_id = DBtoNullableInt(dr.Item("local_section_id"))
        Me.rho_override = DBtoNullableDbl(dr.Item("rho_override"))
    End Sub
#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.drilled_pier_profile_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pier_diameter.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.clear_cover.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.clear_cover_rebar_cage_option.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.tie_size.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.tie_spacing.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bottom_elevation.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_section_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rho_override.ToString.FormatDBValue)
        Return SQLInsertValues

    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
        SQLInsertFields = SQLInsertFields.AddtoDBString("drilled_pier_profile_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pier_diameter")
        SQLInsertFields = SQLInsertFields.AddtoDBString("clear_cover")
        SQLInsertFields = SQLInsertFields.AddtoDBString("clear_cover_rebar_cage_option")
        SQLInsertFields = SQLInsertFields.AddtoDBString("tie_size")
        SQLInsertFields = SQLInsertFields.AddtoDBString("tie_spacing")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bottom_elevation")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_section_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("rho_override")
        Return SQLInsertFields

    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("drilled_pier_profile_id = " & Me.drilled_pier_profile_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pier_diameter = " & Me.pier_diameter.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("clear_cover = " & Me.clear_cover.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("clear_cover_rebar_cage_option = " & Me.clear_cover_rebar_cage_option.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("tie_size = " & Me.tie_size.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("tie_spacing = " & Me.tie_spacing.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bottom_elevation = " & Me.bottom_elevation.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_section_id = " & Me.local_section_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("rho_override = " & Me.rho_override.ToString.FormatDBValue)
        Return SQLUpdateFieldsandValues

    End Function
#End Region

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As DrilledPierSection = TryCast(other, DrilledPierSection)
        If otherToCompare Is Nothing Then Return False
        Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
        Equals = If(Me.drilled_pier_profile_id.CheckChange(otherToCompare.drilled_pier_profile_id, changes, categoryName, "Drilled Pier Profile Id"), Equals, False)
        Equals = If(Me.pier_diameter.CheckChange(otherToCompare.pier_diameter, changes, categoryName, "Pier Diameter"), Equals, False)
        Equals = If(Me.clear_cover.CheckChange(otherToCompare.clear_cover, changes, categoryName, "Clear Cover"), Equals, False)
        Equals = If(Me.clear_cover_rebar_cage_option.CheckChange(otherToCompare.clear_cover_rebar_cage_option, changes, categoryName, "Clear Cover Rebar Cage Option"), Equals, False)
        Equals = If(Me.tie_size.CheckChange(otherToCompare.tie_size, changes, categoryName, "Tie Size"), Equals, False)
        Equals = If(Me.tie_spacing.CheckChange(otherToCompare.tie_spacing, changes, categoryName, "Tie Spacing"), Equals, False)
        Equals = If(Me.bottom_elevation.CheckChange(otherToCompare.bottom_elevation, changes, categoryName, "Bottom Elevation"), Equals, False)
        Equals = If(Me.local_section_id.CheckChange(otherToCompare.local_section_id, changes, categoryName, "Local Section Id"), Equals, False)
        Equals = If(Me.rho_override.CheckChange(otherToCompare.rho_override, changes, categoryName, "Rho Override"), Equals, False)
        Return Equals

    End Function
End Class

Partial Public Class DrilledPierRebar
    Inherits EDSObjectWithQueries

    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Drilled Pier Rebar"
        End Get
    End Property

    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "fnd.drilled_pier_rebar"
        End Get
    End Property

#Region "Define"
    Private _ID As Integer?
    Private _section_id As Integer?
    Private _longitudinal_rebar_quantity As Integer?
    Private _longitudinal_rebar_size As Integer?
    Private _longitudinal_rebar_cage_diameter As Double?
    Private _local_rebar_id As Integer?
    <Category("Drilled Pier"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer?
        Get
            Return Me._ID
        End Get
        Set
            Me._ID = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Section Id")>
    Public Property section_id() As Integer?
        Get
            Return Me._section_id
        End Get
        Set
            Me._section_id = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Longitudinal Rebar Quantity")>
    Public Property longitudinal_rebar_quantity() As Integer?
        Get
            Return Me._longitudinal_rebar_quantity
        End Get
        Set
            Me._longitudinal_rebar_quantity = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Longitudinal Rebar Size")>
    Public Property longitudinal_rebar_size() As Integer?
        Get
            Return Me._longitudinal_rebar_size
        End Get
        Set
            Me._longitudinal_rebar_size = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Longitudinal Rebar Cage Diameter")>
    Public Property longitudinal_rebar_cage_diameter() As Double?
        Get
            Return Me._longitudinal_rebar_cage_diameter
        End Get
        Set
            Me._longitudinal_rebar_cage_diameter = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Local Rebar Id")>
    Public Property local_rebar_id() As Integer?
        Get
            Return Me._local_rebar_id
        End Get
        Set
            Me._local_rebar_id = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()

    End Sub

    Public Sub New(ByVal dr As DataRow)
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        Me.section_id = DBtoNullableInt(dr.Item("section_id"))
        Me.longitudinal_rebar_quantity = DBtoNullableInt(dr.Item("longitudinal_rebar_quantity"))
        Me.longitudinal_rebar_size = DBtoNullableInt(dr.Item("longitudinal_rebar_size"))
        Me.longitudinal_rebar_cage_diameter = DBtoNullableDbl(dr.Item("longitudinal_rebar_cage_diameter"))
        Me.local_rebar_id = DBtoNullableInt(dr.Item("local_rebar_id"))

    End Sub
#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.section_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.longitudinal_rebar_quantity.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.longitudinal_rebar_size.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.longitudinal_rebar_cage_diameter.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_rebar_id.ToString.FormatDBValue)
        Return SQLInsertValues

    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
        SQLInsertFields = SQLInsertFields.AddtoDBString("section_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("longitudinal_rebar_quantity")
        SQLInsertFields = SQLInsertFields.AddtoDBString("longitudinal_rebar_size")
        SQLInsertFields = SQLInsertFields.AddtoDBString("longitudinal_rebar_cage_diameter")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_rebar_id")
        Return SQLInsertFields

    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("section_id = " & Me.section_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("longitudinal_rebar_quantity = " & Me.longitudinal_rebar_quantity.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("longitudinal_rebar_size = " & Me.longitudinal_rebar_size.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("longitudinal_rebar_cage_diameter = " & Me.longitudinal_rebar_cage_diameter.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_rebar_id = " & Me.local_rebar_id.ToString.FormatDBValue)
        Return SQLUpdateFieldsandValues

    End Function
#End Region

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As DrilledPierRebar = TryCast(other, DrilledPierRebar)
        If otherToCompare Is Nothing Then Return False
        Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
        Equals = If(Me.section_id.CheckChange(otherToCompare.section_id, changes, categoryName, "Section Id"), Equals, False)
        Equals = If(Me.longitudinal_rebar_quantity.CheckChange(otherToCompare.longitudinal_rebar_quantity, changes, categoryName, "Longitudinal Rebar Quantity"), Equals, False)
        Equals = If(Me.longitudinal_rebar_size.CheckChange(otherToCompare.longitudinal_rebar_size, changes, categoryName, "Longitudinal Rebar Size"), Equals, False)
        Equals = If(Me.longitudinal_rebar_cage_diameter.CheckChange(otherToCompare.longitudinal_rebar_cage_diameter, changes, categoryName, "Longitudinal Rebar Cage Diameter"), Equals, False)
        Equals = If(Me.local_rebar_id.CheckChange(otherToCompare.local_rebar_id, changes, categoryName, "Local Rebar Id"), Equals, False)
        Return Equals

    End Function
End Class

Partial Public Class EmbeddedPole
    Inherits EDSObjectWithQueries

    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Embedded Pole"
        End Get
    End Property

    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "fnd.embedded_pole"
        End Get
    End Property

#Region "Define"
    Private _ID As Integer?
    Private _drilled_pier_profile_id As Integer?
    Private _embedded_pole_option As Boolean?
    Private _encased_in_concrete As Boolean?
    Private _pole_side_quantity As Integer?
    Private _pole_yield_strength As Double?
    Private _pole_thickness As Double?
    Private _embedded_pole_input_type As String
    Private _pole_diameter_toc As Double?
    Private _pole_top_diameter As Double?
    Private _pole_bottom_diameter As Double?
    Private _pole_section_length As Double?
    Private _pole_taper_factor As Double?
    Private _pole_bend_radius_override As Double?
    <Category("Drilled Pier"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer?
        Get
            Return Me._ID
        End Get
        Set
            Me._ID = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Drilled Pier Profile Id")>
    Public Property drilled_pier_profile_id() As Integer?
        Get
            Return Me._drilled_pier_profile_id
        End Get
        Set
            Me._drilled_pier_profile_id = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Embedded Pole Option")>
    Public Property embedded_pole_option() As Boolean?
        Get
            Return Me._embedded_pole_option
        End Get
        Set
            Me._embedded_pole_option = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Encased In Concrete")>
    Public Property encased_in_concrete() As Boolean?
        Get
            Return Me._encased_in_concrete
        End Get
        Set
            Me._encased_in_concrete = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Pole Side Quantity")>
    Public Property pole_side_quantity() As Integer?
        Get
            Return Me._pole_side_quantity
        End Get
        Set
            Me._pole_side_quantity = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Pole Yield Strength")>
    Public Property pole_yield_strength() As Double?
        Get
            Return Me._pole_yield_strength
        End Get
        Set
            Me._pole_yield_strength = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Pole Thickness")>
    Public Property pole_thickness() As Double?
        Get
            Return Me._pole_thickness
        End Get
        Set
            Me._pole_thickness = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Embedded Pole Input Type")>
    Public Property embedded_pole_input_type() As String
        Get
            Return Me._embedded_pole_input_type
        End Get
        Set
            Me._embedded_pole_input_type = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Pole Diameter Toc")>
    Public Property pole_diameter_toc() As Double?
        Get
            Return Me._pole_diameter_toc
        End Get
        Set
            Me._pole_diameter_toc = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Pole Top Diameter")>
    Public Property pole_top_diameter() As Double?
        Get
            Return Me._pole_top_diameter
        End Get
        Set
            Me._pole_top_diameter = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Pole Bottom Diameter")>
    Public Property pole_bottom_diameter() As Double?
        Get
            Return Me._pole_bottom_diameter
        End Get
        Set
            Me._pole_bottom_diameter = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Pole Section Length")>
    Public Property pole_section_length() As Double?
        Get
            Return Me._pole_section_length
        End Get
        Set
            Me._pole_section_length = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Pole Taper Factor")>
    Public Property pole_taper_factor() As Double?
        Get
            Return Me._pole_taper_factor
        End Get
        Set
            Me._pole_taper_factor = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Pole Bend Radius Override")>
    Public Property pole_bend_radius_override() As Double?
        Get
            Return Me._pole_bend_radius_override
        End Get
        Set
            Me._pole_bend_radius_override = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()

    End Sub

    Public Sub New(ByVal dr As DataRow)
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        Me.drilled_pier_profile_id = DBtoNullableInt(dr.Item("drilled_pier_profile_id"))
        Me.embedded_pole_option = DBtoNullableBool(dr.Item("embedded_pole_option"))
        Me.encased_in_concrete = DBtoNullableBool(dr.Item("encased_in_concrete"))
        Me.pole_side_quantity = DBtoNullableInt(dr.Item("pole_side_quantity"))
        Me.pole_yield_strength = DBtoNullableDbl(dr.Item("pole_yield_strength"))
        Me.pole_thickness = DBtoNullableDbl(dr.Item("pole_thickness"))
        Me.embedded_pole_input_type = DBtoStr(dr.Item("embedded_pole_input_type"))
        Me.pole_diameter_toc = DBtoNullableDbl(dr.Item("pole_diameter_toc"))
        Me.pole_top_diameter = DBtoNullableDbl(dr.Item("pole_top_diameter"))
        Me.pole_bottom_diameter = DBtoNullableDbl(dr.Item("pole_bottom_diameter"))
        Me.pole_section_length = DBtoNullableDbl(dr.Item("pole_section_length"))
        Me.pole_taper_factor = DBtoNullableDbl(dr.Item("pole_taper_factor"))
        Me.pole_bend_radius_override = DBtoNullableDbl(dr.Item("pole_bend_radius_override"))
    End Sub
#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.drilled_pier_profile_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.embedded_pole_option.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.encased_in_concrete.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pole_side_quantity.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pole_yield_strength.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pole_thickness.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.embedded_pole_input_type.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pole_diameter_toc.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pole_top_diameter.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pole_bottom_diameter.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pole_section_length.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pole_taper_factor.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pole_bend_radius_override.ToString.FormatDBValue)
        Return SQLInsertValues

    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
        SQLInsertFields = SQLInsertFields.AddtoDBString("drilled_pier_profile_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("embedded_pole_option")
        SQLInsertFields = SQLInsertFields.AddtoDBString("encased_in_concrete")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pole_side_quantity")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pole_yield_strength")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pole_thickness")
        SQLInsertFields = SQLInsertFields.AddtoDBString("embedded_pole_input_type")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pole_diameter_toc")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pole_top_diameter")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pole_bottom_diameter")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pole_section_length")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pole_taper_factor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pole_bend_radius_override")
        Return SQLInsertFields

    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("drilled_pier_profile_id = " & Me.drilled_pier_profile_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("embedded_pole_option = " & Me.embedded_pole_option.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("encased_in_concrete = " & Me.encased_in_concrete.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pole_side_quantity = " & Me.pole_side_quantity.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pole_yield_strength = " & Me.pole_yield_strength.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pole_thickness = " & Me.pole_thickness.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("embedded_pole_input_type = " & Me.embedded_pole_input_type.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pole_diameter_toc = " & Me.pole_diameter_toc.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pole_top_diameter = " & Me.pole_top_diameter.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pole_bottom_diameter = " & Me.pole_bottom_diameter.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pole_section_length = " & Me.pole_section_length.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pole_taper_factor = " & Me.pole_taper_factor.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pole_bend_radius_override = " & Me.pole_bend_radius_override.ToString.FormatDBValue)
        Return SQLUpdateFieldsandValues

    End Function
#End Region

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As EmbeddedPole = TryCast(other, EmbeddedPole)
        If otherToCompare Is Nothing Then Return False
        Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
        Equals = If(Me.drilled_pier_profile_id.CheckChange(otherToCompare.drilled_pier_profile_id, changes, categoryName, "Drilled Pier Profile Id"), Equals, False)
        Equals = If(Me.embedded_pole_option.CheckChange(otherToCompare.embedded_pole_option, changes, categoryName, "Embedded Pole Option"), Equals, False)
        Equals = If(Me.encased_in_concrete.CheckChange(otherToCompare.encased_in_concrete, changes, categoryName, "Encased In Concrete"), Equals, False)
        Equals = If(Me.pole_side_quantity.CheckChange(otherToCompare.pole_side_quantity, changes, categoryName, "Pole Side Quantity"), Equals, False)
        Equals = If(Me.pole_yield_strength.CheckChange(otherToCompare.pole_yield_strength, changes, categoryName, "Pole Yield Strength"), Equals, False)
        Equals = If(Me.pole_thickness.CheckChange(otherToCompare.pole_thickness, changes, categoryName, "Pole Thickness"), Equals, False)
        Equals = If(Me.embedded_pole_input_type.CheckChange(otherToCompare.embedded_pole_input_type, changes, categoryName, "Embedded Pole Input Type"), Equals, False)
        Equals = If(Me.pole_diameter_toc.CheckChange(otherToCompare.pole_diameter_toc, changes, categoryName, "Pole Diameter Toc"), Equals, False)
        Equals = If(Me.pole_top_diameter.CheckChange(otherToCompare.pole_top_diameter, changes, categoryName, "Pole Top Diameter"), Equals, False)
        Equals = If(Me.pole_bottom_diameter.CheckChange(otherToCompare.pole_bottom_diameter, changes, categoryName, "Pole Bottom Diameter"), Equals, False)
        Equals = If(Me.pole_section_length.CheckChange(otherToCompare.pole_section_length, changes, categoryName, "Pole Section Length"), Equals, False)
        Equals = If(Me.pole_taper_factor.CheckChange(otherToCompare.pole_taper_factor, changes, categoryName, "Pole Taper Factor"), Equals, False)
        Equals = If(Me.pole_bend_radius_override.CheckChange(otherToCompare.pole_bend_radius_override, changes, categoryName, "Pole Bend Radius Override"), Equals, False)
        Return Equals

    End Function
End Class

Partial Public Class BelledPier
    Inherits EDSObjectWithQueries

    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Belled Pier"
        End Get
    End Property

    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "fnd.belled_pier"
        End Get
    End Property

#Region "Define"
    Private _ID As Integer?
    Private _drilled_pier_profile_id As Integer?
    Private _belled_pier_option As Boolean?
    Private _bottom_diameter_of_bell As Double?
    Private _bell_input_type As String
    Private _bell_angle As Double?
    Private _bell_height As Double?
    Private _bell_toe_height As Double?
    Private _neglect_top_soil_layer As Boolean?
    Private _swelling_expansive_soil As Boolean?
    Private _depth_of_expansive_soil As Double?
    Private _expansive_soil_force As Double?
    <Category("Drilled Pier"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer?
        Get
            Return Me._ID
        End Get
        Set
            Me._ID = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Drilled Pier Profile Id")>
    Public Property drilled_pier_profile_id() As Integer?
        Get
            Return Me._drilled_pier_profile_id
        End Get
        Set
            Me._drilled_pier_profile_id = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Belled Pier Option")>
    Public Property belled_pier_option() As Boolean?
        Get
            Return Me._belled_pier_option
        End Get
        Set
            Me._belled_pier_option = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Bottom Diameter Of Bell")>
    Public Property bottom_diameter_of_bell() As Double?
        Get
            Return Me._bottom_diameter_of_bell
        End Get
        Set
            Me._bottom_diameter_of_bell = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Bell Input Type")>
    Public Property bell_input_type() As String
        Get
            Return Me._bell_input_type
        End Get
        Set
            Me._bell_input_type = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Bell Angle")>
    Public Property bell_angle() As Double?
        Get
            Return Me._bell_angle
        End Get
        Set
            Me._bell_angle = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Bell Height")>
    Public Property bell_height() As Double?
        Get
            Return Me._bell_height
        End Get
        Set
            Me._bell_height = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Bell Toe Height")>
    Public Property bell_toe_height() As Double?
        Get
            Return Me._bell_toe_height
        End Get
        Set
            Me._bell_toe_height = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Neglect Top Soil Layer")>
    Public Property neglect_top_soil_layer() As Boolean?
        Get
            Return Me._neglect_top_soil_layer
        End Get
        Set
            Me._neglect_top_soil_layer = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Swelling Expansive Soil")>
    Public Property swelling_expansive_soil() As Boolean?
        Get
            Return Me._swelling_expansive_soil
        End Get
        Set
            Me._swelling_expansive_soil = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Depth Of Expansive Soil")>
    Public Property depth_of_expansive_soil() As Double?
        Get
            Return Me._depth_of_expansive_soil
        End Get
        Set
            Me._depth_of_expansive_soil = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Expansive Soil Force")>
    Public Property expansive_soil_force() As Double?
        Get
            Return Me._expansive_soil_force
        End Get
        Set
            Me._expansive_soil_force = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()

    End Sub

    Public Sub New(ByVal dr As DataRow)
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        Me.drilled_pier_profile_id = DBtoNullableInt(dr.Item("drilled_pier_profile_id"))
        Me.belled_pier_option = DBtoNullableBool(dr.Item("belled_pier_option"))
        Me.bottom_diameter_of_bell = DBtoNullableDbl(dr.Item("bottom_diameter_of_bell"))
        Me.bell_input_type = DBtoStr(dr.Item("bell_input_type"))
        Me.bell_angle = DBtoNullableDbl(dr.Item("bell_angle"))
        Me.bell_height = DBtoNullableDbl(dr.Item("bell_height"))
        Me.bell_toe_height = DBtoNullableDbl(dr.Item("bell_toe_height"))
        Me.neglect_top_soil_layer = DBtoNullableBool(dr.Item("neglect_top_soil_layer"))
        Me.swelling_expansive_soil = DBtoNullableBool(dr.Item("swelling_expansive_soil"))
        Me.depth_of_expansive_soil = DBtoNullableDbl(dr.Item("depth_of_expansive_soil"))
        Me.expansive_soil_force = DBtoNullableDbl(dr.Item("expansive_soil_force"))
    End Sub
#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.drilled_pier_profile_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.belled_pier_option.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bottom_diameter_of_bell.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bell_input_type.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bell_angle.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bell_height.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bell_toe_height.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.neglect_top_soil_layer.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.swelling_expansive_soil.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.depth_of_expansive_soil.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.expansive_soil_force.ToString.FormatDBValue)
        Return SQLInsertValues

    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
        SQLInsertFields = SQLInsertFields.AddtoDBString("drilled_pier_profile_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("belled_pier_option")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bottom_diameter_of_bell")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bell_input_type")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bell_angle")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bell_height")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bell_toe_height")
        SQLInsertFields = SQLInsertFields.AddtoDBString("neglect_top_soil_layer")
        SQLInsertFields = SQLInsertFields.AddtoDBString("swelling_expansive_soil")
        SQLInsertFields = SQLInsertFields.AddtoDBString("depth_of_expansive_soil")
        SQLInsertFields = SQLInsertFields.AddtoDBString("expansive_soil_force")
        Return SQLInsertFields

    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("drilled_pier_profile_id = " & Me.drilled_pier_profile_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("belled_pier_option = " & Me.belled_pier_option.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bottom_diameter_of_bell = " & Me.bottom_diameter_of_bell.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bell_input_type = " & Me.bell_input_type.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bell_angle = " & Me.bell_angle.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bell_height = " & Me.bell_height.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bell_toe_height = " & Me.bell_toe_height.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("neglect_top_soil_layer = " & Me.neglect_top_soil_layer.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("swelling_expansive_soil = " & Me.swelling_expansive_soil.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("depth_of_expansive_soil = " & Me.depth_of_expansive_soil.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("expansive_soil_force = " & Me.expansive_soil_force.ToString.FormatDBValue)
        Return SQLUpdateFieldsandValues

    End Function
#End Region

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As BelledPier = TryCast(other, BelledPier)
        If otherToCompare Is Nothing Then Return False
        Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
        Equals = If(Me.drilled_pier_profile_id.CheckChange(otherToCompare.drilled_pier_profile_id, changes, categoryName, "Drilled Pier Profile Id"), Equals, False)
        Equals = If(Me.belled_pier_option.CheckChange(otherToCompare.belled_pier_option, changes, categoryName, "Belled Pier Option"), Equals, False)
        Equals = If(Me.bottom_diameter_of_bell.CheckChange(otherToCompare.bottom_diameter_of_bell, changes, categoryName, "Bottom Diameter Of Bell"), Equals, False)
        Equals = If(Me.bell_input_type.CheckChange(otherToCompare.bell_input_type, changes, categoryName, "Bell Input Type"), Equals, False)
        Equals = If(Me.bell_angle.CheckChange(otherToCompare.bell_angle, changes, categoryName, "Bell Angle"), Equals, False)
        Equals = If(Me.bell_height.CheckChange(otherToCompare.bell_height, changes, categoryName, "Bell Height"), Equals, False)
        Equals = If(Me.bell_toe_height.CheckChange(otherToCompare.bell_toe_height, changes, categoryName, "Bell Toe Height"), Equals, False)
        Equals = If(Me.neglect_top_soil_layer.CheckChange(otherToCompare.neglect_top_soil_layer, changes, categoryName, "Neglect Top Soil Layer"), Equals, False)
        Equals = If(Me.swelling_expansive_soil.CheckChange(otherToCompare.swelling_expansive_soil, changes, categoryName, "Swelling Expansive Soil"), Equals, False)
        Equals = If(Me.depth_of_expansive_soil.CheckChange(otherToCompare.depth_of_expansive_soil, changes, categoryName, "Depth Of Expansive Soil"), Equals, False)
        Equals = If(Me.expansive_soil_force.CheckChange(otherToCompare.expansive_soil_force, changes, categoryName, "Expansive Soil Force"), Equals, False)
        Return Equals

    End Function
End Class

