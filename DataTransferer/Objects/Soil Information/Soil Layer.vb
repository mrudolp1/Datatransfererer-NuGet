Option Strict On

Imports System.ComponentModel
Imports System.Data
Imports DevExpress.Spreadsheet
Imports Microsoft.Office.Interop


Partial Public Class SoilLayer
    Inherits EDSObjectWithQueries
    'Inherits EDSObject

#Region "Inheritted"
    Public Overrides ReadOnly Property EDSObjectName As String = "Soil Layer"
    Public Overrides ReadOnly Property EDSTableName As String = "fnd.soil_layer"

    Public Overrides Function SQLInsert() As String

        'SQLInsert = QueryBuilderFromFile(queryPath & "Soil Layer\Soil Layer (INSERT).sql")
        SQLInsert = CCI_Engineering_Templates.My.Resources.Soil_Layer_INSERT
        SQLInsert = SQLInsert.Replace("[SOIL LAYER VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[SOIL LAYER FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLInsert

    End Function

    Public Overrides Function SQLUpdate() As String

        'SQLUpdate = QueryBuilderFromFile(queryPath & "Soil Layer\Soil Layer (UPDATE).sql")
        SQLUpdate = CCI_Engineering_Templates.My.Resources.Soil_Layer_UPDATE
        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
        SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLUpdate

    End Function

    Public Overrides Function SQLDelete() As String

        'SQLDelete = QueryBuilderFromFile(queryPath & "Soil Layer\Soil Layer (DELETE).sql")
        SQLDelete = CCI_Engineering_Templates.My.Resources.Soil_Layer_DELETE
        SQLDelete = SQLDelete.Replace("[ID]", Me.ID.ToString.FormatDBValue)
        SQLDelete = SQLDelete.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLDelete

    End Function

#End Region
#Region "Define"

    Private _ID As Integer?
    Private _soil_profile_id As Integer?
    Private _bottom_depth As Double?
    Private _effective_soil_density As Double?
    Private _cohesion As Double?
    Private _friction_angle As Double?
    Private _skin_friction_override_comp As Double?
    Private _skin_friction_override_uplift As Double?
    Private _nominal_bearing_capacity As Double? 'Does not apply to Piles
    Private _spt_blow_count As Integer?
    <Category("Soil Layer"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer?
        Get
            Return Me._ID
        End Get
        Set
            Me._ID = Value
        End Set
    End Property
    <Category("Soil Layer"), Description(""), DisplayName("Soil Profile ID")>
    Public Property Soil_Profile_id() As Integer?
        Get
            Return Me._soil_profile_id
        End Get
        Set
            Me._soil_profile_id = Value
        End Set
    End Property
    <Category("Soil Layer"), Description(""), DisplayName("Bottom Depth")>
    Public Property bottom_depth() As Double?
        Get
            Return Me._bottom_depth
        End Get
        Set
            Me._bottom_depth = Value
        End Set
    End Property
    <Category("Soil Layer"), Description(""), DisplayName("Effective Soil Density")>
    Public Property effective_soil_density() As Double?
        Get
            Return Me._effective_soil_density
        End Get
        Set
            Me._effective_soil_density = Value
        End Set
    End Property
    <Category("Soil Layer"), Description(""), DisplayName("Cohesion")>
    Public Property cohesion() As Double?
        Get
            Return Me._cohesion
        End Get
        Set
            Me._cohesion = Value
        End Set
    End Property
    <Category("Soil Layer"), Description(""), DisplayName("Friction Angle")>
    Public Property friction_angle() As Double?
        Get
            Return Me._friction_angle
        End Get
        Set
            Me._friction_angle = Value
        End Set
    End Property
    <Category("Soil Layer"), Description(""), DisplayName("Skin Friction Override Comp")>
    Public Property skin_friction_override_comp() As Double?
        Get
            Return Me._skin_friction_override_comp
        End Get
        Set
            Me._skin_friction_override_comp = Value
        End Set
    End Property
    <Category("Soil Layer"), Description(""), DisplayName("Skin Friction Override Uplift")>
    Public Property skin_friction_override_uplift() As Double?
        Get
            Return Me._skin_friction_override_uplift
        End Get
        Set
            Me._skin_friction_override_uplift = Value
        End Set
    End Property
    <Category("Soil Layer"), Description(""), DisplayName("Nominal Bearing Capacity")>
    Public Property nominal_bearing_capacity() As Double?
        Get
            Return Me._nominal_bearing_capacity
        End Get
        Set
            Me._nominal_bearing_capacity = Value
        End Set
    End Property
    <Category("Soil Layer"), Description(""), DisplayName("Spt Blow Count")>
    Public Property spt_blow_count() As Integer?
        Get
            Return Me._spt_blow_count
        End Get
        Set
            Me._spt_blow_count = Value
        End Set
    End Property


#End Region

#Region "Constructors"
    Public Sub New()
        'Leave Method Empty
    End Sub

    Public Sub New(ByVal Row As DataRow)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        'If Parent IsNot Nothing Then Me.Absorb(Parent)
        ''''''Customize for each foundation type'''''
        Dim excelDS As New DataSet
        Dim dr = Row
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        Me.bottom_depth = DBtoNullableDbl(dr.Item("bottom_depth"))
        Me.effective_soil_density = DBtoNullableDbl(dr.Item("effective_soil_density"))
        Me.cohesion = DBtoNullableDbl(dr.Item("cohesion"))
        Me.friction_angle = DBtoNullableDbl(dr.Item("friction_angle"))
        Me.skin_friction_override_comp = DBtoNullableDbl(dr.Item("skin_friction_override_comp"))
        Me.skin_friction_override_uplift = DBtoNullableDbl(dr.Item("skin_friction_override_uplift"))
        Me.nominal_bearing_capacity = DBtoNullableDbl(dr.Item("nominal_bearing_capacity")) 'Does not apply to Piles
        Me.spt_blow_count = DBtoNullableInt(dr.Item("spt_blow_count"))

    End Sub


#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel1ID")
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bottom_depth.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.effective_soil_density.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cohesion.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.friction_angle.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.skin_friction_override_comp.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.skin_friction_override_uplift.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.nominal_bearing_capacity.ToString.FormatDBValue) 'Does not apply to Piles
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.spt_blow_count.ToString.FormatDBValue)


        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""
        'SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
        SQLInsertFields = SQLInsertFields.AddtoDBString("soil_profile_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bottom_depth")
        SQLInsertFields = SQLInsertFields.AddtoDBString("effective_soil_density")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cohesion")
        SQLInsertFields = SQLInsertFields.AddtoDBString("friction_angle")
        SQLInsertFields = SQLInsertFields.AddtoDBString("skin_friction_override_comp")
        SQLInsertFields = SQLInsertFields.AddtoDBString("skin_friction_override_uplift")
        SQLInsertFields = SQLInsertFields.AddtoDBString("nominal_bearing_capacity") 'Does not apply to Piles
        SQLInsertFields = SQLInsertFields.AddtoDBString("spt_blow_count")


        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bottom_depth = " & Me.bottom_depth.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("effective_soil_density = " & Me.effective_soil_density.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cohesion = " & Me.cohesion.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("friction_angle = " & Me.friction_angle.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("skin_friction_override_comp = " & Me.skin_friction_override_comp.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("skin_friction_override_uplift = " & Me.skin_friction_override_uplift.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("nominal_bearing_capacity = " & Me.nominal_bearing_capacity.ToString.FormatDBValue) 'Does not apply to Piles
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("spt_blow_count = " & Me.spt_blow_count.ToString.FormatDBValue)


        Return SQLUpdateFieldsandValues
    End Function
#End Region


    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As SoilLayer = TryCast(other, SoilLayer)
        If otherToCompare Is Nothing Then Return False

        'Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
        Equals = If(Me.bottom_depth.CheckChange(otherToCompare.bottom_depth, changes, categoryName, "Bottom Depth"), Equals, False)
        Equals = If(Me.effective_soil_density.CheckChange(otherToCompare.effective_soil_density, changes, categoryName, "Effective Soil Density"), Equals, False)
        Equals = If(Me.cohesion.CheckChange(otherToCompare.cohesion, changes, categoryName, "Cohesion"), Equals, False)
        Equals = If(Me.friction_angle.CheckChange(otherToCompare.friction_angle, changes, categoryName, "Friction Angle"), Equals, False)
        Equals = If(Me.skin_friction_override_comp.CheckChange(otherToCompare.skin_friction_override_comp, changes, categoryName, "Skin Friction Override Comp"), Equals, False)
        Equals = If(Me.skin_friction_override_uplift.CheckChange(otherToCompare.skin_friction_override_uplift, changes, categoryName, "Skin Friction Override Uplift"), Equals, False)
        Equals = If(Me.nominal_bearing_capacity.CheckChange(otherToCompare.nominal_bearing_capacity, changes, categoryName, "Nominal Bearing Capacity"), Equals, False)
        Equals = If(Me.spt_blow_count.CheckChange(otherToCompare.spt_blow_count, changes, categoryName, "Spt Blow Count"), Equals, False)


    End Function

End Class