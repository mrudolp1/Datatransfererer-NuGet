Imports System.ComponentModel
Imports System.Security.Principal
Imports DevExpress.Spreadsheet
Imports System.IO
Imports DevExpress.DataAccess.Excel
Imports System.Runtime.CompilerServices
Imports System.Data.SqlClient

<TypeConverterAttribute(GetType(ExpandableObjectConverter))>
Partial Public MustInherit Class EDSObject
    Implements IComparable(Of EDSObject), IEquatable(Of EDSObject)
    <Category("EDS"), Description(""), DisplayName("ID")>
    Public Property ID As Integer?
    <Category("EDS"), Description(""), DisplayName("Name")>
    Public MustOverride ReadOnly Property EDSObjectName As String
    <Category("EDS"), Description(""), DisplayName("Full Name")>
    Public Overridable ReadOnly Property EDSObjectFullName As String
        Get
            Return If(Me.Parent Is Nothing, Me.EDSObjectName, Me.Parent.EDSObjectFullName & " - " & Me.EDSObjectName)
        End Get
    End Property
    <Category("EDS"), Description(""), Browsable(False)>
    Public Overridable Property Parent As EDSObject
    <Category("EDS"), Description(""), Browsable(False)>
    Public Overridable ReadOnly Property ParentStructure As EDSStructure
        Get
            Return Me.Parent?.ParentStructure
        End Get
        'Set(value As EDSStructure)

        'End Set
    End Property
    <Category("EDS"), Description(""), DisplayName("BU")>
    Public Property bus_unit As String
    <Category("EDS"), Description(""), DisplayName("Structure ID")>
    Public Property structure_id As String
    <Category("EDS"), Description(""), DisplayName("Work Order")>
    Public Property work_order_seq_num As String
    <Category("EDS"), Description(""), Browsable(False)>
    Public Property activeDatabase As String
    <Category("EDS"), Description(""), Browsable(False)>
    Public Property databaseIdentity As WindowsIdentity
    <Category("EDS"), Description(""), Browsable(False)>
    Public Property modified_person_id As Integer?
    <Category("EDS"), Description(""), Browsable(False)>
    Public Property process_stage As String

    'Public Property differences As List(Of ObjectsComparer.Difference)

    Public Overridable Sub Absorb(ByRef Host As EDSObject)
        Me.Parent = Host
        'Me.ParentStructure = Host.ParentStructure 'If(Host.ParentStructure, Nothing) 'The parent of an EDSObject should be the top level structure.
        Me.bus_unit = Host.bus_unit
        Me.structure_id = Host.structure_id
        Me.work_order_seq_num = Host.work_order_seq_num
        Me.activeDatabase = Host.activeDatabase
        Me.databaseIdentity = Host.databaseIdentity
        Me.modified_person_id = Host.modified_person_id
        Me.process_stage = Host.process_stage
    End Sub


    Public Function CompareTo(other As EDSObject) As Integer Implements IComparable(Of EDSObject).CompareTo
        'This is used to sort EDSObjects
        'They will be sorted by ID by default.
        If other Is Nothing Then
            Return 1
        Else
            Return Nullable.Compare(Me.ID, other.ID)
        End If
    End Function

    'Reference for implementing IEquatable: https://www.codeproject.com/Articles/20592/Implementing-IEquatable-Properly
    Public Overloads Function Equals(other As EDSObject) As Boolean Implements IEquatable(Of EDSObject).Equals

        If other Is Nothing Then
            Return False
        Else
            'Call Equals(other As EDSObject, ByRef changes As List(Of AnalysisChanges))
            Return Me.Equals(other, New List(Of AnalysisChange))
        End If

    End Function
    Public Overloads Overrides Function Equals(other As Object) As Boolean
        'This will be called if an object other than an EDS object is passed in
        Dim EDSOther As EDSObject = TryCast(other, EDSObject)

        If EDSOther Is Nothing Then
            Return False
        Else
            'Call Equals(other As EDSObject) 
            Return Me.Equals(other)
        End If

    End Function
    Public Overrides Function GetHashCode() As Integer
        'Fun Story about hash codes: https://stackoverflow.com/questions/7425142/what-is-hashcode-used-for-is-it-unique
        'Creating hash codes: https://thomaslevesque.com/2020/05/15/things-every-csharp-developer-should-know-1-hash-codes/
        Dim HashTuple As Tuple(Of String, Integer?, String, String) = New Tuple(Of String, Integer?, String, String)(Me.EDSObjectName, Me.ID, Me.bus_unit, Me.structure_id)
        Return HashTuple.GetHashCode
    End Function

    Public MustOverride Overloads Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean

    Public Overrides Function ToString() As String
        Return Me.EDSObjectName
    End Function

End Class

Partial Public MustInherit Class EDSObjectWithQueries
    Inherits EDSObject
    <Category("EDS Queries"), Description("EDS Table Name with schema."), DisplayName("Table Name")>
    Public MustOverride ReadOnly Property EDSTableName As String
    <Category("EDS Queries"), Description("Local path to query templates."), DisplayName("Query Path")>
    Public Overridable ReadOnly Property EDSQueryPath As String = IO.Path.Combine(My.Application.Info.DirectoryPath, "Templates")
    <Category("EDS Queries"), Description("Depth of table in EDS query. This determines where the ID is stored in the query and which parent ID is referenced if needed. 0 = Top Level"), Browsable(False)>
    Public Overridable ReadOnly Property EDSTableDepth As Integer = 0
    Public Overridable Property Results As New List(Of EDSResult)


    <Category("EDS Queries"), Description("Insert this object and results into EDS. For use in whole structure query. Requires two variable in main query [@Prev Table (ID INT)] and [@Prev ID INT]"), DisplayName("SQL Insert Query")>
    Public Overridable Function SQLInsert() As String
        SQLInsert = "BEGIN" & vbCrLf &
                     "  INSERT INTO [TABLE] ([FIELDS])" & vbCrLf &
                     "  OUTPUT INSERTED.ID INTO " & EDSStructure.SQLQueryTableVar(Me.EDSTableDepth) & vbCrLf &
                     "  VALUES([VALUES])" & vbCrLf &
                     "  Select " & EDSStructure.SQLQueryIDVar(Me.EDSTableDepth) & "=ID FROM " & EDSStructure.SQLQueryTableVar(Me.EDSTableDepth) & vbCrLf &
                     "   [RESULTS]" & vbCrLf &
                     "  Delete FROM " & EDSStructure.SQLQueryTableVar(Me.EDSTableDepth) & vbCrLf &
                     "END" & vbCrLf
        SQLInsert = SQLInsert.Replace("[TABLE]", Me.EDSTableName)
        SQLInsert = SQLInsert.Replace("[FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.Replace("[VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[RESULTS]", Me.Results.EDSResultQuery)
        Return SQLInsert
    End Function
    Public Overridable Function SQLSetID(Optional ID As Integer? = Nothing) As String
        Return "SET " & EDSStructure.SQLQueryIDVar(Me.EDSTableDepth) & " = " & If(ID Is Nothing, Me.ID.ToString.FormatDBValue, ID.ToString.FormatDBValue) & vbCrLf
    End Function
    <Category("EDS Queries"), Description("Update existing EDS object and insert results. For use in whole structure query."), DisplayName("SQL Update Query")>
    Public Overridable Function SQLUpdate() As String
        SQLUpdate = "BEGIN" & vbCrLf &
                  "  Update [Table]" &
                  "  SET [UPDATE]" & vbCrLf &
                  "  WHERE ID = [ID]" & vbCrLf &
                  "  [RESULTS]" & vbCrLf &
                  "END" & vbCrLf
        SQLUpdate = SQLUpdate.Replace("[TABLE]", Me.EDSTableName)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID)
        SQLUpdate = SQLUpdate.Replace("[RESULTS]", Me.Results.EDSResultQuery)
        Return SQLUpdate
    End Function
    <Category("EDS Queries"), Description("Delete this object and results from EDS. For use in whole structure query."), DisplayName("SQL Delete Query")>
    Public Overridable Function SQLDelete() As String
        SQLDelete = "BEGIN" & vbCrLf &
                 "  IF EXISTS (SELECT ID FROM [TABLE] WHERE ID = [ID])" & vbCrLf &
                 "      Delete FROM [TABLE] WHERE ID = [ID]" & vbCrLf &
                 "END" & vbCrLf
        SQLDelete = SQLDelete.Replace("[TABLE]", Me.EDSTableName)
        SQLDelete = SQLDelete.Replace("[ID]", Me.ID)
        Return SQLDelete
    End Function

    Public Function ResultQuery(Optional ByVal ResultsParentIDKnown As Boolean = True) As String

        ResultQuery = ""

        For Each result In Me.Results
            ResultQuery += result.Insert(ResultsParentIDKnown) & vbCrLf
        Next

        Return ResultQuery

    End Function

    Public MustOverride Function SQLInsertValues() As String

    Public MustOverride Function SQLInsertFields() As String

    Public MustOverride Function SQLUpdateFieldsandValues() As String

    Public Overridable Function EDSQueryBuilder(ItemToCompare As EDSObjectWithQueries, Optional ByRef AllowUpdate As Boolean = True) As String
        'Compare the ID of the current EDS item to the existing item and determine if the Insert, Update, or Delete query should be used
        'If a parent item is being inserted, existing children should be deleted and new should be inserted. This is handled by the AllowUpdate boolean.
        EDSQueryBuilder = ""

        If ItemToCompare Is Nothing Then
            EDSQueryBuilder += Me.SQLInsert
            AllowUpdate = False
        Else
            If AllowUpdate Then
                EDSQueryBuilder += Me.SQLSetID
                If ItemToCompare.ID = Me.ID And AllowUpdate Then
                    If Not Me.Equals(ItemToCompare) Then
                        EDSQueryBuilder += Me.SQLUpdate
                    Else
                        'Save Results Only
                        EDSQueryBuilder += Me.Results.EDSResultQuery
                    End If

                End If
            Else
                'If this is for a sub table, the delete might need to be handled at the top level table.
                'Top level delete will fail if there are sub tables with a foreign key to the top level table.
                'I tried to avoid an issue by adding the If EXIST line to the standard SQLDelete.
                EDSQueryBuilder += ItemToCompare.SQLDelete
                EDSQueryBuilder += Me.SQLInsert
                AllowUpdate = False
            End If
        End If

        Return EDSQueryBuilder

    End Function

End Class

Partial Public MustInherit Class EDSExcelObject
    'This should be inherited by the main tool class. Subclasses such as soil layers can probably inherit the EDSObjectWithQueries
    Inherits EDSObjectWithQueries
    <Category("Tool"), Description("Local path to query templates."), DisplayName("Tool Path")>
    Public Property workBookPath As String
    <Category("Tool"), Description("Local path to query templates."), Browsable(False)>
    Public MustOverride ReadOnly Property templatePath As String
    <Category("Tool"), Description("Local path to query templates."), DisplayName("File Type")>
    Public Property fileType As DocumentFormat = DocumentFormat.Xlsm
    <Category("Tool"), Description("Data transfer parameters, a list of ranges to import from excel."), DisplayName("Import Ranges")>
    Public MustOverride ReadOnly Property excelDTParams As List(Of EXCELDTParameter)
    <Category("Tool"), Description("Version number of tool."), DisplayName("Tool Version")>
    Public Property tool_version As String
    <Category("Tool"), Description("Have the calculation been modified?"), DisplayName("Modified")>
    Public Property modified As Boolean?

#Region "Save to Excel"
    Public MustOverride Sub workBookFiller(ByRef wb As Workbook)

    Public Sub SavetoExcel()
        Dim wb As New Workbook

        If workBookPath = "" Then
            Debug.Print("No workbook path specified.")
            Exit Sub
        End If

        Try
            wb.LoadDocument(templatePath, fileType)
            wb.BeginUpdate()

            'Put the jelly in the donut
            workBookFiller(wb)

            wb.Calculate()
            wb.EndUpdate()
            wb.SaveDocument(workBookPath, fileType)

        Catch ex As Exception
            Debug.Print("Error Saving Workbook: " & ex.Message)
        End Try

    End Sub
#End Region

End Class

Partial Public Class EDSResult
    Inherits EDSObject

    Public Overrides ReadOnly Property EDSObjectName As String = "Result"

#Region "Define"
    Private _foreign_key As Integer?
    Private _result_lkup As String
    Private _rating As Double?
    Private _EDSTableName As String
    Private _ForeignKeyName As String
    'modified_person_id
    'process_stag

    'Public Shadows Property Parent As EDSObjectWithQueries

    <Category("Results"), Description("The ID of the parent object that this result is associated with. (i.e. Drilled Pier, Tower Leg, Plate)"), DisplayName("Result ID")>
    Public Property foreign_key() As Integer?
        Get
            Return Me._foreign_key
        End Get
        Set
            Me._foreign_key = Value
        End Set
    End Property
    <Category("Results"), Description(""), DisplayName("Result Type")>
    Public Property result_lkup() As String
        Get
            Return Me._result_lkup
        End Get
        Set
            Me._result_lkup = Value
        End Set
    End Property
    <Category("Results"), Description(""), DisplayName("Rating (%)")>
    Public Overridable Property rating() As Double?
        Get
            Return Me._rating
        End Get
        Set
            Me._rating = Value
        End Set
    End Property

    <Category("Results"), Description(""), DisplayName("Result Table Name")>
    Public Property EDSTableName() As String
        Get
            Return Me._EDSTableName
        End Get
        Set
            Me._EDSTableName = Value
        End Set
    End Property

    <Category("Results"), Description(""), DisplayName("Result ID Name")>
    Public Property ForeignKeyName() As String
        Get
            Return Me._ForeignKeyName
        End Get
        Set
            Me._ForeignKeyName = Value
        End Set
    End Property
    <Category("EDS Queries"), Description("Depth of table in EDS query. This determines where the ID is stored in the query and which parent ID is referenced if needed. 0 = Top Level"), Browsable(False)>
    Public Overridable Property EDSTableDepth As Integer = 1

#End Region

    Public Function Insert(Optional ByVal ParentID As Integer? = Nothing) As String
        Insert = "BEGIN" & vbCrLf &
                "  INSERT INTO " & Me.EDSTableName & "(" & Me.SQLInsertFields & ")" & vbCrLf &
                "  VALUES([VALUES])" & vbCrLf &
                "END" & vbCrLf
        Insert = Insert.Replace("[TABLE]", Me.EDSTableName)
        Insert = Insert.Replace("[VALUES]", Me.SQLInsertValues(ParentID))
        Insert = Insert.Replace("[FIELDS]", Me.SQLInsertFields)
        Insert = Insert.TrimEnd()
        Return Insert
    End Function

    Public Function SQLInsertValues(Optional ByVal ParentID As Integer? = Nothing) As String
        SQLInsertValues = ""

        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.work_order_seq_num.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(If(ParentID Is Nothing, EDSStructure.SQLQueryIDVar(Me.EDSTableDepth - 1), Me.foreign_key.ToString.FormatDBValue))
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.result_lkup.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rating.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Function SQLInsertFields() As String
        SQLInsertFields = ""

        SQLInsertFields = SQLInsertFields.AddtoDBString("work_order_seq_num")
        SQLInsertFields = SQLInsertFields.AddtoDBString(Me.ForeignKeyName)
        SQLInsertFields = SQLInsertFields.AddtoDBString("result_lkup")
        SQLInsertFields = SQLInsertFields.AddtoDBString("rating")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")

        Return SQLInsertFields
    End Function

    Public Overloads Sub Absorb(ByRef Host As EDSObjectWithQueries)
        MyBase.Absorb(Host)
        'Results don't have a set table depth, it depends on their parent depth
        Me.EDSTableDepth = Host.EDSTableDepth + 1
        'Results table should be the Parent Table Name + _results (fnd.pier_pad -> fnd.pier_pad_results, tnx.upper_structure_sections -> tnx.upper_structure_section_results)
        Me.EDSTableName = If(Host.EDSTableName(Host.EDSTableName.Length - 1) = "s",
                             Host.EDSTableName.Substring(0, Host.EDSTableName.Length - 1),
                             Host.EDSTableName) & "_results"
        'Result ID name should be Parent Table Name + _id (fnd.pier_pad -> pier_pad_id)
        'Seperate the table name from the schema then add _id
        Me.ForeignKeyName = If(Host.EDSTableName.Contains("."),
                                Host.EDSTableName.Substring(Host.EDSTableName.IndexOf(".") + 1, Host.EDSTableName.Length - Host.EDSTableName.IndexOf(".") - 1) & "_id",
                                Host.EDSTableName & "_id")
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Parent"></param>
    Public Sub New(Optional ByRef Parent As EDSObjectWithQueries = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then
            Me.Absorb(Parent)
        End If
    End Sub

    ''' <summary>
    ''' Create result object with result_lkup and rating
    ''' </summary>
    ''' <param name="result_lkup"></param>
    ''' <param name="rating"></param>
    ''' <param name="Parent"></param>
    Public Sub New(ByVal result_lkup As String, ByVal rating As Double?, Optional ByRef Parent As EDSObjectWithQueries = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then
            Me.Absorb(Parent)
        End If
        Me.result_lkup = result_lkup
        Me.rating = rating
    End Sub

    Public Sub New(ByVal resultDr As DataRow, ByRef Parent As EDSObjectWithQueries)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then
            Me.Absorb(Parent)
        End If

        Me.result_lkup = DBtoStr(resultDr.Item("result_lkup"))
        Me.rating = DBtoNullableDbl(resultDr.Item("rating"))
    End Sub

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Throw New NotImplementedException()
    End Function

End Class








