Option Strict On

Imports System.ComponentModel
Imports System.Data
Imports DevExpress.Spreadsheet
Imports Microsoft.Office.Interop


Partial Public Class SoilLayers
    Inherits EDSObjectWithQueries
    'Inherits EDSObject

#Region "Inheritted"
    Public Overrides ReadOnly Property EDSObjectName As String = "Soil Layer"
    Public Overrides ReadOnly Property EDSTableName As String = "fnd.soil_layer"
    'Public Overrides ReadOnly Property templatePath As String = IO.Path.Combine(My.Application.Info.DirectoryPath, "Templates", "Pile Foundation.xlsm")
    'Public Overrides ReadOnly Property excelDTParams As List(Of EXCELDTParameter)
    '    Get
    '        Return New List(Of EXCELDTParameter) From {New EXCELDTParameter("Pile General Details EXCEL", "A1:BG2", "Details (SAPI)"),
    '                                                    New EXCELDTParameter("Pile Soil Profile EXCEL", "BI1:BJ2", "Details (SAPI)")}
    '        '***Add additional table references here****
    '    End Get
    'End Property
    Private _Insert As String
    Private _Update As String
    Private _Delete As String

    Public Overrides ReadOnly Property Insert() As String
        Get
            Insert = QueryBuilderFromFile(queryPath & "Soil Layer\Soil Layer (INSERT).sql")
            Insert = Insert.Replace("[SOIL LAYER VALUES]", Me.SQLInsertValues)
            Insert = Insert.Replace("[SOIL LAYER FIELDS]", Me.SQLInsertFields)
            Insert = Insert.TrimEnd() 'Removes empty rows that generate within query for each record

            Return Insert
        End Get
    End Property

    Public Overrides ReadOnly Property Update() As String
        Get
            Update = QueryBuilderFromFile(queryPath & "Soil Layer\Soil Layer (UPDATE).sql")
            Update = Update.Replace("[ID]", Me.ID.ToString.FormatDBValue)
            Update = Update.Replace("[UPDATE]", Me.SQLUpdate)
            'UpdateString = UpdateString.Replace("[RESULTS]", Me.Results.EDSResultQuery)
            Return Update
        End Get
    End Property

    Public Overrides ReadOnly Property Delete() As String
        Get
            Delete = QueryBuilderFromFile(queryPath & "Soil Layer\Soil Layer (DELETE).sql")
            Delete = Delete.Replace("[ID]", Me.Soil_Profile_id.ToString.FormatDBValue)
            Delete = Delete.TrimEnd() 'Removes empty rows that generate within query for each record
            Return Delete
        End Get
    End Property

#End Region
#Region "Define"

    'Private _ID As Integer?
    Private _soil_profile_id As Integer?
    Private _bottom_depth As Double?
    Private _effective_soil_density As Double?
    Private _cohesion As Double?
    Private _friction_angle As Double?
    Private _skin_friction_override_comp As Double?
    Private _skin_friction_override_uplift As Double?
    Private _nominal_bearing_capacity As Double? 'Does not apply to Piles
    Private _spt_blow_count As Integer?
    '<Category("Soil Layer"), Description(""), DisplayName("Id")>
    'Public Property ID() As Integer?
    '    Get
    '        Return Me._ID
    '    End Get
    '    Set
    '        Me._ID = Value
    '    End Set
    'End Property
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
        'Me.ID = DBtoNullableInt(dr.Item("ID"))
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
        SQLInsertValues = SQLInsertValues.AddtoDBString("@Sub1ID")
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

    Public Overrides Function SQLUpdate() As String
        SQLUpdate = ""
        'SQLUpdate = SQLUpdate.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("bottom_depth = " & Me.bottom_depth.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("effective_soil_density = " & Me.effective_soil_density.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("cohesion = " & Me.cohesion.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("friction_angle = " & Me.friction_angle.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("skin_friction_override_comp = " & Me.skin_friction_override_comp.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("skin_friction_override_uplift = " & Me.skin_friction_override_uplift.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("nominal_bearing_capacity = " & Me.nominal_bearing_capacity.ToString.FormatDBValue) 'Does not apply to Piles
        SQLUpdate = SQLUpdate.AddtoDBString("spt_blow_count = " & Me.spt_blow_count.ToString.FormatDBValue)


        Return SQLUpdate
    End Function
#End Region


    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Throw New NotImplementedException()
    End Function


End Class