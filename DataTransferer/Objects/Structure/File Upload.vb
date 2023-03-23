Option Strict On
Option Compare Binary

Imports System.ComponentModel
Imports System.Data
Imports DevExpress.Spreadsheet
Imports Microsoft.Office.Interop


Partial Public Class FileUpload
    Inherits EDSObjectWithQueries
    'Inherits EDSObject

#Region "Inheritted"
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "File Upload"
        End Get
    End Property
    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "gen.file_upload"
        End Get
    End Property


    Public Overrides Function SQLInsert() As String

        SQLInsert = QueryBuilderFromFile(queryPath & "Structure\Generic Sub Table\Generic (INSERT).sql")
        SQLInsert = SQLInsert.Replace("[TABLE]", Me.EDSTableName)
        SQLInsert = SQLInsert.Replace("[VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLInsert

    End Function

    Public Overrides Function SQLUpdate() As String

        'SQLUpdate = QueryBuilderFromFile(queryPath & "Soil Layer\Soil Layer (UPDATE).sql")
        'SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID.ToString.FormatDBValue)
        'SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
        'SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

        Return ""

    End Function

    Public Overrides Function SQLDelete() As String

        'SQLDelete = QueryBuilderFromFile(queryPath & "Soil Layer\Soil Layer (DELETE).sql")
        'SQLDelete = SQLDelete.Replace("[ID]", Me.ID.ToString.FormatDBValue)
        'SQLDelete = SQLDelete.TrimEnd() 'Removes empty rows that generate within query for each record

        Return ""

    End Function

#End Region
#Region "Define"
    Public Property file_path As String

    Private _file_lkup_code As String
    Private _file_file_name As String
    Private _file_file_ext As String
    Private _file_file_ver As String
    <Category("File Upload"), Description(""), DisplayName("File Lkup Code")>
    Public Property file_lkup_code() As String
        Get
            Return Me._file_lkup_code
        End Get
        Set
            Me._file_lkup_code = Value
        End Set
    End Property
    <Category("File Upload"), Description(""), DisplayName("File File Name")>
    Public Property file_file_name() As String
        Get
            Return Me._file_file_name
        End Get
        Set
            Me._file_file_name = Value
        End Set
    End Property
    <Category("File Upload"), Description(""), DisplayName("File File Ext")>
    Public Property file_file_ext() As String
        Get
            Return Me._file_file_ext
        End Get
        Set
            Me._file_file_ext = Value
        End Set
    End Property
    <Category("File Upload"), Description(""), DisplayName("File File Ver")>
    Public Property file_file_ver() As String
        Get
            Return Me._file_file_ver
        End Get
        Set
            Me._file_file_ver = Value
        End Set
    End Property
#End Region

#Region "Constructors"
    Public Sub New()
        'Leave Method Empty
    End Sub

    Public Sub New(ByVal dr As DataRow, Optional ByVal Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        Me.file_lkup_code = DBtoStr(dr.Item("file_lkup_code"))
        Me.file_file_name = DBtoStr(dr.Item("file_file_name"))
        Me.file_file_ext = DBtoStr(dr.Item("file_file_ext"))
        Me.file_file_ver = DBtoStr(dr.Item("file_file_ver"))

    End Sub

    Public Sub New(ByVal filePath As String, Optional ByVal Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        SetFileParameters(filePath)
    End Sub

    Public Sub SetFileParameters(ByVal filePath As String)
        Dim myFile As New IO.FileInfo(filePath)

        Me.file_lkup_code = "NOTSET"
        Me.file_file_name = myFile.Name
        Me._file_file_ext = myFile.Extension
        Me.file_file_ver = "NOTSET" 'Maybe set it from the child? 
        Me.file_path = filePath
    End Sub

#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.work_order_seq_num.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.file_lkup_code.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.file_file_name.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.file_file_ext.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.file_file_ver.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Now.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)
        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""
        SQLInsertFields = SQLInsertFields.AddtoDBString("work_order_seq_num")
        SQLInsertFields = SQLInsertFields.AddtoDBString("file_lkup_code")
        SQLInsertFields = SQLInsertFields.AddtoDBString("file_file_name")
        SQLInsertFields = SQLInsertFields.AddtoDBString("file_file_ext")
        SQLInsertFields = SQLInsertFields.AddtoDBString("file_file_ver")
        SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")
        SQLInsertFields = SQLInsertFields.AddtoDBString("upload_by_person_id")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("upload_date")
        'SQLInsertFields = SQLInsertFields.AddtoDBString(Me.process_stage?.ToString)
        Return SQLInsertFields

    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        Return Nothing
    End Function
#End Region


    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As FileUpload = TryCast(other, FileUpload)
        If otherToCompare Is Nothing Then Return False
        Equals = If(Me.file_lkup_code.CheckChange(otherToCompare.file_lkup_code, changes, categoryName, "File Lkup Code"), Equals, False)
        Equals = If(Me.file_file_name.CheckChange(otherToCompare.file_file_name, changes, categoryName, "File File Name"), Equals, False)
        Equals = If(Me.file_file_ext.CheckChange(otherToCompare.file_file_ext, changes, categoryName, "File File Ext"), Equals, False)
        Equals = If(Me.file_file_ver.CheckChange(otherToCompare.file_file_ver, changes, categoryName, "File File Ver"), Equals, False)
        Return Equals
    End Function

End Class