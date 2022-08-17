Option Strict On

Imports System.ComponentModel
Imports System.Data
Imports DevExpress.Spreadsheet
Imports Microsoft.Office.Interop


Partial Public Class SoilProfile
    Inherits EDSObjectWithQueries

#Region "Inheritted"
    Public Overrides ReadOnly Property EDSObjectName As String = "Soil Profile"
    Public Overrides ReadOnly Property EDSTableName As String = "fnd.soil_profile"
    'Public Overrides ReadOnly Property templatePath As String = IO.Path.Combine(My.Application.Info.DirectoryPath, "Templates", "Pile Foundation.xlsm")
    'Public Overrides ReadOnly Property excelDTParams As List(Of EXCELDTParameter)
    '    Get
    '        Return New List(Of EXCELDTParameter) From {New EXCELDTParameter("Pile General Details EXCEL", "A1:BG2", "Details (SAPI)"),
    '                                                    New EXCELDTParameter("Pile Soil Profile EXCEL", "BI1:BJ2", "Details (SAPI)")}
    '        '***Add additional table references here****
    '    End Get
    'End Property

    Public Overrides Function SQLInsert() As String

        SQLInsert = QueryBuilderFromFile(queryPath & "Soil Profile\Soil Profile (INSERT).sql")
        SQLInsert = SQLInsert.Replace("[SOIL PROFILE VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[SOIL PROFILE FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

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


    'Private _Insert As String
    'Private _Update As String
    'Private _Delete As String

    'Public Overrides ReadOnly Property Insert() As String
    '    Get
    '        Insert = QueryBuilderFromFile(queryPath & "Soil Profile\Soil Profile (INSERT).sql")
    '        Insert = Insert.Replace("[SOIL PROFILE VALUES]", Me.SQLInsertValues)
    '        Insert = Insert.Replace("[SOIL PROFILE FIELDS]", Me.SQLInsertFields)

    '        Return Insert
    '    End Get
    'End Property

    'Public Overrides ReadOnly Property Update() As String
    '    Get
    '        Update = QueryBuilderFromFile(queryPath & "Soil Profile\Soil Profile (UPDATE).sql")
    '        Update = Update.Replace("[ID]", Me.ID.ToString.FormatDBValue)
    '        Update = Update.Replace("[UPDATE]", Me.SQLUpdate)
    '        'UpdateString = UpdateString.Replace("[RESULTS]", Me.Results.EDSResultQuery)
    '        Return Update
    '    End Get
    'End Property

    'Public Overrides ReadOnly Property Delete() As String
    '    Get
    '        Delete = QueryBuilderFromFile(queryPath & "Soil Profile\Soil Profile (DELETE).sql")
    '        Delete = Delete.Replace("[ID]", Me.ID.ToString.FormatDBValue)
    '        Return Delete
    '    End Get
    'End Property

#End Region
#Region "Define"

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

    'Public Sub New(ByVal Row As DataRow)
    '    Try
    '        If Not IsDBNull(CType(Row.Item("groundwater_depth"), Double)) Then
    '            Me.groundwater_depth = CType(Row.Item("groundwater_depth"), Double)
    '        Else
    '            Me.groundwater_depth = Nothing
    '        End If
    '    Catch
    '        Me.groundwater_depth = Nothing
    '    End Try 'Pile_X_Coordinate
    '    Try
    '        If Not IsDBNull(CType(Row.Item("neglect_depth"), Double)) Then
    '            Me.neglect_depth = CType(Row.Item("neglect_depth"), Double)
    '        Else
    '            Me.neglect_depth = Nothing
    '        End If
    '    Catch
    '        Me.neglect_depth = Nothing
    '    End Try 'Pile_Y_Coordinate
    'End Sub 'Add a pile location to a pile

    Public Sub New(ByVal Row As DataRow)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        'If Parent IsNot Nothing Then Me.Absorb(Parent)
        ''''''Customize for each foundation type'''''
        Dim excelDS As New DataSet
        'Dim dr = excelDS.Tables("PILE SOIL PROFILE EXCEL").Rows(0)
        Dim dr = Row
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        Me.groundwater_depth = DBtoNullableDbl(dr.Item("groundwater_depth"))
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