Imports System.ComponentModel
Imports System.Security.Principal
Imports DevExpress.Spreadsheet
Imports System.IO
Imports System.Runtime.InteropServices
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.Serialization
'Imports Microsoft.Office.Interop 'added for testing running macros

Public Delegate Function OverwriteFile(ByVal FileName As String) As Boolean

'To expand collections in the main property grid, this process may help but doesn't need to be implemented right now
'https://www.codeproject.com/Articles/4448/Customized-display-of-collection-data-in-a-Propert?fid=16073&df=90&mpp=25&sort=Position&view=Normal&spc=Relaxed&prof=True&fr=176

<Serializable()>
<TypeConverterAttribute(GetType(ExpandableObjectConverter))>
<DataContract()>
<RefreshProperties(RefreshProperties.Repaint)>
Partial Public MustInherit Class EDSObject
    Implements IComparable(Of EDSObject), IEquatable(Of EDSObject)

#Region "ReadOnly Properties"
    <Category("EDS"), Description(""), DisplayName("Name")>
    Public MustOverride ReadOnly Property EDSObjectName As String
    <Category("EDS"), Description(""), DisplayName("Full Name")>
    Public Overridable ReadOnly Property EDSObjectFullName As String
        Get
            Return If(Me.Parent Is Nothing, Me.EDSObjectName, Me.Parent.EDSObjectFullName & " - " & Me.EDSObjectName)
        End Get
        'Set(value As String)
        '    Throw New NotSupportedException("Setting the EDSObjectFullName is not supported")
        'End Set
    End Property
    <Category("EDS"), Description(""), Browsable(False)>
    Public Overridable ReadOnly Property ParentStructure As EDSStructure
        Get
            Return Me.Parent?.ParentStructure
        End Get
    End Property
#End Region
    <Category("EDS"), Description(""), Browsable(False)>
    Public Overridable Property Parent As EDSObject

    <Category("EDS"), Description(""), Browsable(False)>
    Public Property databaseIdentity As WindowsIdentity


    <Category("EDS"), Description(""), DisplayName("ID")>
    <DataMember()>
    Public Property ID As Integer?

    <Category("EDS"), Description(""), DisplayName("BU")>
    <DataMember()>
    Public Property bus_unit As String
    <Category("EDS"), Description(""), DisplayName("Structure ID")>
    <DataMember()>
    Public Property structure_id As String

    <Category("EDS"), Description(""), DisplayName("Work Order")>
    <DataMember()>
    Public Property work_order_seq_num As String

    <Category("EDS"), Description(""), DisplayName("Order")>
    <DataMember()>
    Public Property order As String

    <Category("EDS"), Description(""), DisplayName("Order Revision")>
    <DataMember()>
    Public Property orderRev As String

    <Category("EDS"), Description(""), Browsable(False)>
    <DataMember()>
    Public Property activeDatabase As String

    <Category("EDS"), Description("Modified By"), Browsable(True)>
    <DataMember()>
    Public Property modified_person_id As Integer?

    <Category("EDS"), Description(""), Browsable(True)>
    <DataMember()>
    Public Property process_stage As String = "test" 'added "test" since error occured during testing

    ' <DataMember()> Public Property differences As List(Of ObjectsComparer.Difference)

    Public Overridable Sub Absorb(ByRef Host As EDSObject)
        Me.Parent = Host
        'Me.ParentStructure = Host.ParentStructure 'If(Host.ParentStructure, Nothing) 'The parent of an EDSObject should be the top level structure.
        Me.bus_unit = Host.bus_unit
        Me.structure_id = Host.structure_id
        Me.work_order_seq_num = Host.work_order_seq_num
        Me.order = Host.order
        Me.orderRev = Host.orderRev
        Me.activeDatabase = Host.activeDatabase
        Me.databaseIdentity = Host.databaseIdentity
        Me.modified_person_id = Host.modified_person_id
        Me.process_stage = Host.process_stage
    End Sub

    Public Overrides Function ToString() As String
        Return Me.EDSObjectName
    End Function

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


    'Used to just clear properties specific to each individual object and lists associated with the object
    Public Overridable Sub Clear()

    End Sub

End Class

<DataContract()>
Partial Public MustInherit Class EDSObjectWithQueries
    Inherits EDSObject
    <Category("EDS Queries"), Description("EDS Table Name with schema."), DisplayName("Table Name")>
    Public MustOverride ReadOnly Property EDSTableName As String

    <Category("EDS Queries"), Description("Local path to query templates."), DisplayName("Query Path")>
    Public Overridable ReadOnly Property EDSQueryPath As String
        Get
            Return IO.Path.Combine(My.Application.Info.DirectoryPath, "Templates")
        End Get
    End Property

    <Category("EDS Queries"), Description("Depth of table in EDS query. This determines where the ID is stored in the query and which parent ID is referenced if needed. 0 = Top Level"), Browsable(False)>
    Public Overridable ReadOnly Property EDSTableDepth As Integer = 0

    <DataMember()>
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
                  "  Update [TABLE]" &
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

<DataContract()>
Partial Public MustInherit Class EDSExcelObject
    'This should be inherited by the main tool class. Subclasses such as soil layers can probably inherit the EDSObjectWithQueries
    Inherits EDSObjectWithQueries

#Region "ReadOnly Properties"
    <Category("Tool"), Description("Local path to query templates."), Browsable(False)>
    Public MustOverride ReadOnly Property TemplatePath As String
    <Category("Tool"), Description("Template resource."), Browsable(False)>
    Public MustOverride ReadOnly Property Template As Byte()
    <Category("Tool"), Description("Data transfer parameters, a list of ranges to import from excel."), DisplayName("Import Ranges")>
    Public MustOverride ReadOnly Property ExcelDTParams As List(Of EXCELDTParameter)
#End Region

    <Category("Tool"), Description("Workbook Path."), DisplayName("Tool Path")>
    <DataMember()>
    Public Property WorkBookPath As String

    Private _FileType As DocumentFormat
    <Category("Tool"), Description("Workbook Type."), DisplayName("File Type")>
    <DataMember()>
    Public Property FileType As DocumentFormat
        Get
            Return DocumentFormat.Xlsm
        End Get
        Set(value As DocumentFormat)
            Me._FileType = DocumentFormat.Xlsm
        End Set
    End Property

    <Category("Tool"), Description("Version number of tool."), DisplayName("Version")>
    <DataMember()>
    Public Property Version As String

    '<Category("Tool"), Description("Title of tool."), DisplayName("Title")>
    'Public Overridable ReadOnly Property Title As String
    <Category("Tool"), Description("Have the calculation been modified?"), DisplayName("Modified")>
    <DataMember()>
    Public Property Modified As Boolean?

    Public Function MyTIA(Optional ByVal fullTIA As Boolean = False) As String
        Dim tia_current As String = ""
        If IsSomething(Me.ParentStructure?.tnx?.code?.design?.DesignCode) Then
            If Me.ParentStructure.tnx.code.design.DesignCode = "TIA/EIA-222-F" Then
                tia_current = "F"
            ElseIf Me.ParentStructure.tnx.code.design.DesignCode = "TIA-222-G" Then
                tia_current = "G"
            ElseIf Me.ParentStructure.tnx.code.design.DesignCode = "TIA-222-H" Then
                tia_current = "H"
            Else
                tia_current = "H"
            End If
        Else
            tia_current = "H"
        End If
        Dim addlTIA As String = "TIA-222-"
        If fullTIA Then tia_current = addlTIA + tia_current
        Return tia_current
    End Function

    Public Function MyOrder() As String
        Dim site_app As String = ""
        Dim site_rev As String = ""

        If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.eng_app_id) Then
            site_app = Me.ParentStructure?.structureCodeCriteria?.eng_app_id.ToString
            If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.eng_app_id_revision) Then
                site_rev = Me.ParentStructure?.structureCodeCriteria?.eng_app_id_revision.ToString
                site_app += " REV. " & site_rev
            End If
        End If
        Return site_app
    End Function

#Region "Run Excel Macro"
    ''' <summary>
    ''' Adding this in case we want to use it later on.
    ''' Allows individual objects to have an excel macro ran
    ''' </summary>
    ''' <param name="bigMac"></param> --> Macro name from excel VBA
    ''' <param name="isDevEnv"></param> --> Boolean to specify if it is in development mod
    ''' <param name="isSeismic"></param> --> Boolean to specify if seismic is required
    ''' <returns></returns>
    Public Function OpenExcelRunMacro(bigMac As String, Optional ByVal isDevEnv As Boolean = False, Optional ByVal isSeismic As Boolean = False) As String
        'Built this using a with to just add a '.' where something from the structure is needed. 
        '''WriteLineLogLine
        '''tnxFilePath

        With Me.ParentStructure
            Dim tnxFilePath As String = Me.ParentStructure.tnx.filePath
            Dim toolFileName As String = Path.GetFileName(Me.WorkBookPath)
            Dim excelPath As String = Me.WorkBookPath
            Dim logString As String = ""
            If String.IsNullOrEmpty(excelPath) Or String.IsNullOrEmpty(bigMac) Then
                Return "ERROR | excelPath or bigMac parameter is null or empty"
            End If

            Dim xlApp As Excel.Application = Nothing
            Dim xlWorkBook As Excel.Workbook = Nothing

            Dim errorMessage As String = String.Empty

            Dim xlVisibility As Boolean = False
            If isDevEnv Then
                xlVisibility = True
            End If

            Try
                If File.Exists(excelPath) Then

                    xlApp = CreateObject("Excel.Application")
                    xlApp.Visible = xlVisibility

                    xlWorkBook = xlApp.Workbooks.Open(excelPath)

                    .WriteLineLogLine("INFO | Tool: " & toolFileName)
                    .WriteLineLogLine("INFO | Running macro: " & bigMac)

                    If Not IsNothing(tnxFilePath) Then
                        logString = xlApp.Run(bigMac, tnxFilePath)
                        .WriteLineLogLine("INFO | Macro result: " & vbCrLf & logString)
                    Else
                        .WriteLineLogLine("WARNING | No TNX file path in structure..")
                        logString = xlApp.Run(bigMac)
                        .WriteLineLogLine("INFO | Macro result: " & vbCrLf & logString)
                    End If

                    xlWorkBook.Save()
                Else
                    errorMessage = $"ERROR | {excelPath} path not found!"
                    .WriteLineLogLine(errorMessage)
                    Return errorMessage
                End If
            Catch ex As Exception
                errorMessage = ex.Message
                .WriteLineLogLine(errorMessage)
                Return errorMessage
            Finally
                If xlWorkBook IsNot Nothing Then
                    xlWorkBook.Close()
                    Marshal.ReleaseComObject(xlWorkBook)
                    xlWorkBook = Nothing
                End If
                If xlApp IsNot Nothing Then
                    xlApp.Quit()
                    Marshal.ReleaseComObject(xlApp)
                    xlApp = Nothing
                End If
            End Try

            'New method to just laod data and set properties
            Me.LoadFromExcel()

            'check for seismic
            If isSeismic And logString.ToUpper.Contains("SEISMIC ANALYSIS REQUIRED") Then
                Return "SEISMIC ANALYSIS REQUIRED"
            End If
        End With

        Return "Success"
    End Function
#End Region

#Region "Load From Excel"
    Public Overridable Sub LoadFromExcel()

    End Sub
#End Region

#Region "Save to Excel"
    Public MustOverride Sub workBookFiller(ByRef wb As Workbook)

    Public Sub SavetoExcel(Optional workBookPath As String = Nothing, Optional index As Integer = 0, Optional replaceFiles As Boolean = True)
        Dim wb As New Workbook


        If String.IsNullOrEmpty(workBookPath) Then
            If Me.ParentStructure?.WorkingDirectory Is Nothing Then
                Debug.Print("No workbook path specified.")
                Exit Sub
            End If

            'Build Path
            workBookPath = Path.Combine(Me.ParentStructure.WorkingDirectory, Me.bus_unit & " " & Me.EDSObjectName & " EDS" & If(index = 0, "", " " & (index + 1).ToString()) & Me.FileType.GetExtension())

        End If

        Me.WorkBookPath = workBookPath

        If File.Exists(workBookPath) AndAlso Not replaceFiles Then Exit Sub

        wb.LoadDocument(Template, FileType)
        wb.BeginUpdate()

        'Put the jelly in the donut
        workBookFiller(wb)

        wb.Calculate()
        wb.EndUpdate()
        wb.SaveDocument(workBookPath, FileType)

    End Sub

    Public Sub SavetoExcel(overwriteFile As OverwriteFile, Optional workBookPath As String = Nothing, Optional index As Integer = 0)
        Dim wb As New Workbook


        If String.IsNullOrEmpty(workBookPath) Then
            If Me.ParentStructure?.WorkingDirectory Is Nothing Then
                Debug.Print("No workbook path specified.")
                Exit Sub
            End If

            'Build Path
            workBookPath = Path.Combine(Me.ParentStructure.WorkingDirectory, Me.bus_unit & " " & Me.EDSObjectName & " EDS" & If(index = 0, "", " " & (index + 1).ToString()) & Me.FileType.GetExtension())

        End If

        Me.WorkBookPath = workBookPath

        If File.Exists(workBookPath) AndAlso
            Not overwriteFile(Path.GetFileName(workBookPath)) Then Exit Sub

        wb.LoadDocument(Template, FileType)
        wb.BeginUpdate()

        'Put the jelly in the donut
        workBookFiller(wb)

        wb.Calculate()
        wb.EndUpdate()
        wb.SaveDocument(workBookPath, FileType)

    End Sub

    Public Sub SavetoExcel()
        Dim wb As New Workbook

        If WorkBookPath = "" Then
            Debug.Print("No workbook path specified.")
            Exit Sub
        End If

        'Try
        wb.LoadDocument(TemplatePath, FileType)
        wb.BeginUpdate()

        'Put the jelly in the donut
        workBookFiller(wb)

        wb.Calculate()
        wb.EndUpdate()
        wb.SaveDocument(WorkBookPath, FileType)

    End Sub
#End Region

End Class

<DataContract(), KnownTypeAttribute(GetType(EDSResult))>
Partial Public Class EDSResult
    Inherits EDSObject

    Public Overrides ReadOnly Property EDSObjectName As String = "Result"

#Region "Define"
    Private _foreign_key As Integer?
    Private _result_lkup As String
    Private _rating As Decimal?
    Private _EDSTableName As String
    Private _ForeignKeyName As String

    <Category("Results"), Description("The ID of the parent object that this result is associated with. (i.e. Drilled Pier, Tower Leg, Plate)"), DisplayName("Result ID")>
    <DataMember()>
    Public Property foreign_key() As Integer?
        Get
            Return Me._foreign_key
        End Get
        Set
            Me._foreign_key = Value
        End Set
    End Property
    <Category("Results"), Description(""), DisplayName("Result Type")>
    <DataMember()>
    Public Property result_lkup() As String
        Get
            Return Me._result_lkup
        End Get
        Set
            Me._result_lkup = Value
        End Set
    End Property
    <Category("Results"), Description(""), DisplayName("Rating (%)")>
    <DataMember()>
    Public Overridable Property rating() As Decimal?
        Get
            Return Me._rating
        End Get
        Set
            Me._rating = Value
        End Set
    End Property

    <Category("Results"), Description(""), DisplayName("Result Table Name")>
    <DataMember()>
    Public Property EDSTableName() As String
        Get
            Return Me._EDSTableName
        End Get
        Set
            Me._EDSTableName = Value
        End Set
    End Property

    <Category("Results"), Description(""), DisplayName("Result ID Name")>
    <DataMember()>
    Public Property ForeignKeyName() As String
        Get
            Return Me._ForeignKeyName
        End Get
        Set
            Me._ForeignKeyName = Value
        End Set
    End Property

    <Category("EDS Queries"), Description("Depth of table in EDS query. This determines where the ID is stored in the query and which parent ID is referenced if needed. 0 = Top Level"), Browsable(False)>
    <DataMember()>
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

        If Me.EDSTableName = "fnd.anchor_block_results" Then
            EDSTableDepth = 3
        End If

        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.work_order_seq_num.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(If(ParentID Is Nothing, EDSStructure.SQLQueryIDVar(Me.EDSTableDepth - 1), Me.foreign_key.ToString.FormatDBValue))
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.result_lkup.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rating.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Function SQLInsertFields() As String
        SQLInsertFields = ""

        SQLInsertFields = SQLInsertFields.AddtoDBString("work_order_seq_num")
        SQLInsertFields = SQLInsertFields.AddtoDBString(Me.ForeignKeyName)
        SQLInsertFields = SQLInsertFields.AddtoDBString("result_lkup")
        SQLInsertFields = SQLInsertFields.AddtoDBString("rating")
        SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")

        Return SQLInsertFields
    End Function

    Public Overloads Sub Absorb(ByRef Host As EDSObjectWithQueries)
        Me.Parent = Host
        'Me.ParentStructure = Host.ParentStructure 'If(Host.ParentStructure, Nothing) 'The parent of an EDSObject should be the top level structure.
        Me.bus_unit = Host.bus_unit
        Me.structure_id = Host.structure_id
        Me.work_order_seq_num = Host.work_order_seq_num
        Me.order = Host.order
        Me.orderRev = Host.orderRev
        Me.activeDatabase = Host.activeDatabase
        Me.databaseIdentity = Host.databaseIdentity
        Me.modified_person_id = Host.modified_person_id
        Me.process_stage = Host.process_stage
        'Results don't have a set table depth, it depends on their parent depth
        Me.EDSTableDepth = Host.EDSTableDepth + 1
        'Results table should be the Parent Table Name + _results (fnd.pier_pad -> fnd.pier_pad_results, tnx.upper_structure_sections -> tnx.upper_structure_section_results)
        'Me.EDSTableName = If(Host.EDSTableName(Host.EDSTableName.Length - 1) = "s",
        '                     Host.EDSTableName.Substring(0, Host.EDSTableName.Length - 1),
        '                     Host.EDSTableName) & "_results"
        Me.EDSTableName = RemovePlural(Host.EDSTableName) & "_results"
        'Result ID name should be Parent Table Name + _id (fnd.pier_pad -> pier_pad_id)
        'Seperate the table name from the schema then add _id
        'Me.ForeignKeyName = If(Host.EDSTableName.Contains("."),
        '                        Host.EDSTableName.Substring(Host.EDSTableName.IndexOf(".") + 1, Host.EDSTableName.Length - Host.EDSTableName.IndexOf(".") - 1) & "_id",
        '                        Host.EDSTableName & "_id")
        Me.ForeignKeyName = RemovePlural(Host.EDSTableName.Split(".").Last) & "_id"
    End Sub

    Private Function RemovePlural(possiblyPlural As String) As String
        Return If(possiblyPlural.ToLower.Last = "s", possiblyPlural.Remove(possiblyPlural.Length - 1), possiblyPlural)
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Parent"></param>
    Public Sub New(Optional ByVal Parent As EDSObjectWithQueries = Nothing)
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
    Public Sub New(ByVal result_lkup As String, ByVal rating As Decimal?, Optional ByVal Parent As EDSObjectWithQueries = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then
            Me.Absorb(Parent)
        End If
        Me.result_lkup = result_lkup
        Me.rating = rating
    End Sub

    Public Sub New(ByVal resultDr As DataRow, ByVal Parent As EDSObjectWithQueries)
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








