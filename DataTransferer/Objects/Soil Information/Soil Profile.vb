Option Strict Off
Option Compare Binary

Imports System.ComponentModel
Imports System.Data
Imports DevExpress.Spreadsheet
Imports Microsoft.Office.Interop


Partial Public Class SoilProfile
    Inherits EDSObjectWithQueries

#Region "Inheritted"
    Public Overrides ReadOnly Property EDSObjectName As String = "Soil Profile"
    Public Overrides ReadOnly Property EDSTableName As String = "fnd.soil_profile"

    Public Overrides Function SQLInsert() As String

        SQLInsert = QueryBuilderFromFile(queryPath & "Soil Profile\Soil Profile (INSERT).sql")
        SQLInsert = SQLInsert.Replace("[SOIL PROFILE VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[SOIL PROFILE FIELDS]", Me.SQLInsertFields)

        Dim _layerInsert As String
        For Each layer In Me.SoilLayers
            _layerInsert += layer.SQLInsert + vbCrLf
        Next

        SQLInsert = SQLInsert.Replace("--[SOIL LAYER INSERT]", _layerInsert + vbCrLf)
        SQLInsert = SQLInsert.TrimEnd

        Return SQLInsert
    End Function

    Public Overrides Function SQLUpdate() As String

        SQLUpdate = QueryBuilderFromFile(queryPath & "Soil Profile\Soil Profile (UPDATE).sql")
        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
        SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLUpdate
    End Function

    Public Overrides Function SQLDelete() As String

        SQLDelete = QueryBuilderFromFile(queryPath & "Soil Profile\Soil Profile (DELETE).sql")
        SQLDelete = SQLDelete.Replace("[ID]", Me.ID.ToString.FormatDBValue)
        SQLDelete = SQLDelete.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLDelete
    End Function

#End Region
#Region "Define"

    Public Property SoilLayers As New List(Of SoilLayer)

    Private _groundwater_depth As Double?
    Private _neglect_depth As Double?

    <Category("Soil Profile"), Description(""), DisplayName("Groundwater Depth")>
    Public Property groundwater_depth() As Double?
        Get
            Return Me._groundwater_depth
        End Get
        Set
            Me._groundwater_depth = Value
        End Set
    End Property
    <Category("Soil Profile"), Description(""), DisplayName("Neglect Depth")>
    Public Property neglect_depth() As Double?
        Get
            Return Me._neglect_depth
        End Get
        Set
            Me._neglect_depth = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()
        'Leave Method Empty
    End Sub

    Public Sub New(ByVal Row As DataRow, Optional ByRef Parent As EDSObject = Nothing)
        ConstructMe(Row, Parent)
    End Sub

    Public Sub ConstructMe(ByVal Row As DataRow, Optional ByRef Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)
        ''''''Customize for each foundation type'''''
        Dim excelDS As New DataSet
        'Dim dr = excelDS.Tables("PILE SOIL PROFILE EXCEL").Rows(0)
        Dim dr = Row
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        'Me.groundwater_depth = DBtoNullableDbl(dr.Item("groundwater_depth"))
        Me.groundwater_depth = If(DBtoStr(dr.Item("groundwater_depth")) = "N/A", -1, DBtoNullableDbl(dr.Item("groundwater_depth")))
        Me.neglect_depth = DBtoNullableDbl(dr.Item("neglect_depth"))
    End Sub

#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.groundwater_depth.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.neglect_depth.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""
        'SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
        SQLInsertFields = SQLInsertFields.AddtoDBString("groundwater_depth")
        SQLInsertFields = SQLInsertFields.AddtoDBString("neglect_depth")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("groundwater_depth = " & Me.groundwater_depth.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("neglect_depth = " & Me.neglect_depth.ToString.FormatDBValue)

        Return SQLUpdateFieldsandValues
    End Function
#End Region


    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As SoilProfile = TryCast(other, SoilProfile)
        If otherToCompare Is Nothing Then Return False

        'Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
        Equals = If(Me.groundwater_depth.CheckChange(otherToCompare.groundwater_depth, changes, categoryName, "Groundwater Depth"), Equals, False)
        Equals = If(Me.neglect_depth.CheckChange(otherToCompare.neglect_depth, changes, categoryName, "Neglect Depth"), Equals, False)

    End Function

End Class