﻿Imports System.ComponentModel
Imports System.Security.Principal
Imports DevExpress.Spreadsheet
Imports System.IO
Imports DevExpress.DataAccess.Excel
Imports System.Runtime.CompilerServices
Imports System.Data.SqlClient

Public Module Extensions

    <Extension()>
    Public Function FormatDBValue(input As String) As String
        'Handles nullable values and quoatations needed for DB values

        If String.IsNullOrEmpty(input) Then
            FormatDBValue = "NULL"
        Else
            FormatDBValue = "'" & input & "'"
        End If

        Return FormatDBValue
    End Function

    <Extension()>
    Public Function AddtoDBString(startingString As String, newString As String, Optional isDBValue As Boolean = False) As String
        'isValue should be false if you're creating a string of column names. They should not be in single quotes like the values.

        If isDBValue Then newString = newString.FormatDBValue

        If String.IsNullOrEmpty(startingString) Then
            startingString = newString
        Else
            startingString += ", " & newString
        End If

        Return startingString
    End Function

    '<Extension()>
    'Public Function ToDBString(aString As String, Optional isValue As Boolean = True) As String
    '    If aString = String.Empty Or aString Is Nothing Then
    '        Return "NULL"
    '    Else
    '        If isValue Then aString = "'" & aString & "'"
    '        Return aString
    '    End If
    'End Function

    '<Extension()>
    'Public Function AddtoDBString(astring As String, ByRef newString As String, Optional isValue As Boolean = True) As String
    '    'isValue should be false if you're creating a string of column names. They should not be in single quotes like the values.
    '    If astring = String.Empty Or astring Is Nothing Then
    '        astring = newString.ToDBString(isValue)
    '    Else
    '        astring += ", " & newString.ToDBString(isValue)
    '    End If
    '    Return astring
    'End Function

    '<Extension()>
    'Public Function GetDistinct(Of T As EDSObject)(alist As List(Of T)) As List(Of T)
    '    'Notes: Removes duplicates from list of tnxDatabaseEntry by using their CompareMe function
    '    'Making this generic (Of T As tnxDatabaseEntry) allows it to work for all subclasses of tnxDatabaseEntry

    '    Dim distinctList As New List(Of T)

    '    For Each item In alist
    '        Dim addToList As Boolean = True
    '        For Each distinctItem In distinctList
    '            If item.CompareMe(distinctItem) Then
    '                'Not distinct
    '                addToList = False
    '                Exit For
    '            End If
    '        Next
    '        If addToList Then distinctList.Add(item)
    '    Next

    '    Return distinctList
    'End Function

    <Extension()>
    Public Function EDSListQuery(Of T As EDSObjectWithQueries)(alist As List(Of T), prevList As List(Of T)) As String

        EDSListQuery = ""

        'Create a shallow copy of the lists for sorting and comparison
        'Sort lists by ID descending with Null IDs at the bottom
        Dim currentSortedList As List(Of T) = alist.ToList
        currentSortedList.Sort()
        currentSortedList.Reverse()

        Dim prevSortedList As List(Of T) = prevList.ToList
        prevSortedList.Sort()
        prevSortedList.Reverse()


        Dim i As Integer = 0
        Do While i <= Math.Max(currentSortedList.Count, prevSortedList.Count) - 1

            If i > currentSortedList.Count - 1 Then
                'Delete items in previous list if there is nothing left in current list
                EDSListQuery += prevSortedList(i).Delete
            ElseIf i > prevSortedList.Count - 1 Then
                'Insert items in current list if there is nothing left in previous list
                EDSListQuery += currentSortedList(i).Insert
            Else
                'Compare IDs
                If currentSortedList(i).ID = prevSortedList(i).ID Then
                    If Not currentSortedList(i).Equals(prevSortedList(i)) Then
                        'Update existing
                        EDSListQuery += currentSortedList(i).Update
                    Else
                        'Save Results Only
                        EDSListQuery += currentSortedList(i).Results.EDSResultQuery
                    End If
                ElseIf currentSortedList(i).ID < prevSortedList(i).ID Then
                    EDSListQuery += prevSortedList(i).Delete
                    currentSortedList.Insert(i, Nothing)
                Else
                    'currentSortedList(i).ID > prevSortedList(i).ID
                    EDSListQuery += currentSortedList(i).Insert
                    prevSortedList.Insert(i, Nothing)
                End If
            End If

            i += 1

        Loop

        Return EDSListQuery

    End Function

    <Extension()>
    Public Function EDSResultQuery(alist As List(Of EDSResult), Optional ByVal ResultsParentIDKnown As Boolean = True) As String

        EDSResultQuery = ""

        For Each result In alist
            EDSResultQuery += result.Insert(ResultsParentIDKnown) & vbCrLf
        Next

        Return EDSResultQuery

    End Function

    <Extension()>
    Public Function CheckChange(Of T)(value1 As T, value2 As T, ByRef changes As List(Of AnalysisChange), Optional categoryName As String = Nothing, Optional fieldName As String = Nothing) As Boolean

        'Check if this is an EDSObject
        Dim EDSValue1 As EDSObject = TryCast(value1, EDSObject)
        Dim EDSValue2 As EDSObject = TryCast(value2, EDSObject)
        If EDSValue1 IsNot Nothing AndAlso EDSValue2 IsNot Nothing Then
            Return EDSValue1.Equals(EDSValue2, changes)
        End If

        'Check if this is a collection (list), iterate through if needed
        Dim CollectionValue1 As IEnumerable(Of Object) = TryCast(value1, IEnumerable(Of Object))
        Dim CollectionValue2 As IEnumerable(Of Object) = TryCast(value2, IEnumerable(Of Object))
        If CollectionValue1 IsNot Nothing AndAlso CollectionValue2 IsNot Nothing Then
            If CollectionValue1.Count <> CollectionValue2.Count Then
                changes.Add(New AnalysisChange(categoryName, fieldName & "Quantity", CollectionValue1.Count.ToString, CollectionValue2.Count.ToString))
                Return False
            Else
                For i As Integer = 0 To CollectionValue1.Count - 1
                    CollectionValue1(i).CheckChange(CollectionValue2(i), changes, categoryName, If(fieldName Is Nothing, Nothing, fieldName & " (" & i & ")"))
                Next
            End If
        End If

        'Try to compare values directly
        Try
            If Not value1.Equals(value2) Then
                changes.Add(New AnalysisChange(categoryName, fieldName, value1.ToString, value2.ToString))
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            changes.Add(New AnalysisChange(categoryName, fieldName, "Comparison Failed", ""))
            Return False
        End Try

    End Function
    '<Extension()>
    'Public Function CheckChange(Of T)(value1 As T, value2 As T, ByRef changes As List(Of AnalysisChange), Optional ParentChange As AnalysisChange = Nothing) As Boolean
    '    Dim Change As AnalysisChange
    '    If ParentChange Is Nothing Then
    '        ParentChange = New AnalysisChange()
    '    Else
    '        Change = ParentChange.
    '    End If

    '    'Check if this is an EDSObject
    '    Dim EDSValue1 As EDSObject = TryCast(value1, EDSObject)
    '    Dim EDSValue2 As EDSObject = TryCast(value2, EDSObject)
    '    If EDSValue1 IsNot Nothing AndAlso EDSValue2 IsNot Nothing Then
    '        Return EDSValue1.Equals(EDSValue2, changes)
    '    End If

    '    'Check if this is a collection (list), iterate through if needed
    '    Dim CollectionValue1 As IEnumerable(Of Object) = TryCast(value1, IEnumerable(Of Object))
    '    Dim CollectionValue2 As IEnumerable(Of Object) = TryCast(value2, IEnumerable(Of Object))
    '    If CollectionValue1 IsNot Nothing AndAlso CollectionValue2 IsNot Nothing Then
    '        If CollectionValue1.Count <> CollectionValue2.Count Then
    '            ParentChange.NewValue = CollectionValue1.Count.ToString
    '            ParentChange.PreviousValue = CollectionValue2.Count.ToString
    '            changes.Add(New AnalysisChange(categoryName, fieldName & "Quantity", CollectionValue1.Count.ToString, CollectionValue2.Count.ToString))
    '            Return False
    '        Else
    '            For i As Integer = 0 To CollectionValue1.Count - 1
    '                CollectionValue1(i).CheckChange(CollectionValue2(i), changes, categoryName, fieldName & "(" & i & ")")
    '            Next
    '        End If
    '    End If

    '    'Try to compare values directly
    '    Try
    '        If Not value1.Equals(value2) Then
    '            changes.Add(New AnalysisChange(categoryName, fieldName, value1.ToString, value2.ToString))
    '            Return False
    '        Else
    '            Return True
    '        End If
    '    Catch ex As Exception
    '        changes.Add(New AnalysisChange(categoryName, fieldName, "Comparison Failed", ""))
    '        Return False
    '    End Try

    'End Function

    '<Extension()>
    'Public Iterator Function Add(Of T As ObjectsComparer.Difference)(ByVal e As IEnumerable(Of T), ByVal value As T, Optional ByVal Path As String = Nothing) As IEnumerable(Of T)
    '    'Allow you to add to an IEnumerable like it is a list.
    '    'Useful for working with the ObjectComparer class which stores the differences as IEnumerable(of Difference)
    '    'Refernce: https://stackoverflow.com/a/1210311
    '    For Each cur In e
    '        Yield cur
    '    Next

    '    If Path IsNot Nothing Then
    '        Yield value.InsertPath(Path)
    '    Else
    '        Yield value
    '    End If
    'End Function

    '<Extension()>
    'Public Iterator Function Add(Of T As ObjectsComparer.Difference)(ByVal e1 As IEnumerable(Of T), ByVal e2 As IEnumerable(Of T), Optional ByVal Path As String = Nothing) As IEnumerable(Of T)
    '    'Allow you to add to an IEnumerable to another IEnumerable.

    '    For Each cur In e1
    '        Yield cur
    '    Next

    '    For Each cur In e2
    '        If Path IsNot Nothing Then
    '            Yield cur.InsertPath(Path)
    '        Else
    '            Yield cur
    '        End If
    '    Next

    'End Function

    '<Extension()>
    'Public Function Compare(Comparer As ObjectsComparer.Comparer, obj1 As Object, obj2 As Object, path As String, ByRef differences As IEnumerable(Of ObjectsComparer.Difference)) As Boolean
    '    'Compare and add path to all new differences 

    '    Dim newDifferences As IEnumerable(Of ObjectsComparer.Difference) = Nothing  '= Comparer.CalculateDifferences(obj1, obj2)

    '    Comparer.Compare(obj1, obj2, newDifferences)

    '    differences.Add(newDifferences, path)

    '    MessageBox.Show(newDifferences.Any())

    '    Return Not newDifferences.Any()
    'End Function

End Module

Public Module myLittleHelpers
    Public Function trueFalseYesNo(input As String) As Boolean?
        If input.ToLower = "yes" Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function trueFalseYesNo(input As Boolean?) As String
        If input Then
            Return "Yes"
        Else
            Return "No"
        End If
    End Function
    Public Function BooltoBitString(input As Boolean?) As String
        If input Then
            Return "1"
        Else
            Return "0"
        End If
    End Function

    Public Function DBtoNullableInt(ByVal item As Object) As Integer?
        If IsDBNull(item) Then
            Return Nothing
        Else
            Try
                Return CInt(item)
            Catch ex As Exception
                Return Nothing
            End Try
        End If
    End Function

    Public Function DBtoNullableDbl(ByVal item As Object) As Double?
        If IsDBNull(item) Then
            Return Nothing
        Else
            Try
                Return Math.Round(CDbl(item), 6)
            Catch ex As Exception
                Return Nothing
            End Try
        End If
    End Function

    Public Function DBtoNullableBool(ByVal item As Object) As Boolean?
        If IsDBNull(item) Then
            Return Nothing
        Else
            Try
                Return CBool(item)
            Catch ex As Exception
                Return Nothing
            End Try
        End If
    End Function

    Public Function DBtoStr(ByVal item As Object) As String
        'Strings are nullable, but the default value is "" so that's what should be used
        'CStr(Nothing) = "" which is the default value of a string, so that's what we should use with DBNull too. This works better for comparing EDS to ERI opbjects
        If IsDBNull(item) Then
            Return ""
        Else
            Try
                Return CStr(item)
            Catch ex As Exception
                Return ""
            End Try
        End If
    End Function

    Public Function GetTypeNullable(Of T)(ByVal obj As T) As Type
        'This should return the type of a Nullable even if it's null
        If Nullable.GetUnderlyingType(GetType(T)) Is Nothing Then
            Return GetType(T)
        Else
            Return Nullable.GetUnderlyingType(GetType(T))
        End If
    End Function
End Module

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
    Public Overridable Property ParentStructure As EDSStructure
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
        Me.ParentStructure = If(Host.ParentStructure, Nothing) 'The parent of an EDSObject should be the top level structure.
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
        'They will be sorted by ID
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
            Return Me.Equals(other, Nothing)
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
        Dim HashTuple As Tuple(Of String, String) = New Tuple(Of String, String)(Me.bus_unit, Me.structure_id)
        Return HashTuple.GetHashCode
    End Function

    Public MustOverride Overloads Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean

End Class

Partial Public MustInherit Class EDSObjectWithQueries
    Inherits EDSObject
    <Category("EDS Queries"), Description("EDS Table Name with schema."), DisplayName("Table Name")>
    Public MustOverride ReadOnly Property EDSTableName As String
    <Category("EDS Queries"), Description("Local path to query templates."), DisplayName("Query Path")>
    Public Overridable ReadOnly Property EDSQueryPath As String = IO.Path.Combine(My.Application.Info.DirectoryPath, "Templates")
    <Category("Results"), Description("List of results."), DisplayName("Results")>
    Public Property Results As New List(Of EDSResult)
    <Category("EDS Queries"), Description("Insert this object and results into EDS. For use in whole structure query. Requires two variable in main query [@Prev Table (ID INT)] and [@Prev ID INT]"), DisplayName("SQL Insert Query")>
    Public Overridable ReadOnly Property Insert() As String
        Get
            Insert = "BEGIN" & vbCrLf &
                     "  INSERT INTO [TABLE] ([FIELDS])" & vbCrLf &
                     "  OUTPUT INSERTED.ID INTO @Prev" & vbCrLf &
                     "  VALUES([VALUES])" & vbCrLf &
                     "  Select @PrevID=ID FROM @Prev" & vbCrLf &
                     "   [RESULTS]" & vbCrLf &
                     "  Delete FROM @Prev" & vbCrLf &
                     "END" & vbCrLf
            Insert = Insert.Replace("[TABLE]", Me.EDSTableName.FormatDBValue)
            Insert = Insert.Replace("[FIELDS]", Me.SQLInsertFields)
            Insert = Insert.Replace("[VALUES]", Me.SQLInsertValues)
            Insert = Insert.Replace("[RESULTS]", Me.Results.EDSResultQuery(False))
            Return Insert
        End Get
    End Property
    <Category("EDS Queries"), Description("Update existing EDS object and insert results. For use in whole structure query."), DisplayName("SQL Update Query")>
    Public Overridable ReadOnly Property Update() As String
        Get
            Update = "BEGIN" & vbCrLf &
                      "  Update [Table]" &
                      "  SET [UPDATE]" & vbCrLf &
                      "  WHERE ID = [ID]" & vbCrLf &
                      "  [RESULTS]" & vbCrLf &
                      "END" & vbCrLf
            Update = Update.Replace("[TABLE]", Me.EDSTableName.FormatDBValue)
            Update = Update.Replace("[UPDATE]", Me.SQLUpdate)
            Update = Update.Replace("[ID]", Me.ID)
            Update = Update.Replace("[RESULTS]", Me.Results.EDSResultQuery)
            Return Update
        End Get
    End Property
    <Category("EDS Queries"), Description("Delete this object and results from EDS. For use in whole structure query."), DisplayName("SQL Delete Query")>
    Public Overridable ReadOnly Property Delete() As String
        Get
            Delete = "BEGIN" & vbCrLf &
                     "  Delete FROM [TABLE] WHERE ID = [ID]" & vbCrLf &
                     "END"
            Delete = Delete.Replace("[TABLE]", Me.EDSTableName.FormatDBValue)
            Delete = Delete.Replace("[ID]", Me.ID)
            Return Delete
        End Get
    End Property

    'Public MustOverride Function SQLInsertUpdateDelete() As String

    Public MustOverride Function SQLInsertValues() As String

    Public MustOverride Function SQLInsertFields() As String

    Public MustOverride Function SQLUpdate() As String

    'Public Overridable Function EDSQuery(Of T As EDSObjectWithQueries)(item As T, prevItem As T) As String
    '    'Compare the ID of the current EDS item to the existing item and determine if the Insert, Update, or Delete query should be used

    '    EDSQuery = ""

    '    'If prevItem.ID = item.ID And Not item.CompareMe(prevItem) Then
    '    If prevItem.ID = item.ID And Not item.Equals(prevItem) Then
    '        EDSQuery += item.Update
    '    Else
    '        'Need to add inserted items to comparison list.
    '        EDSQuery += item.Insert
    '        If prevItem IsNot Nothing Then
    '            EDSQuery += prevItem.Delete
    '        End If
    '    End If

    '    Return EDSQuery

    'End Function

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

'I don't think this is needed at the moment. Added the EDSObjectname property to the EDSObject class which can replace the foundationType
'Partial Public MustInherit Class EDSFoundation
'    Inherits EDSExcelObject

'    Public MustOverride ReadOnly Property foundationType As String

'End Class

Partial Public Class EDSStructure
    Inherits EDSObject

    Public Overrides ReadOnly Property EDSObjectName As String = "Structure Model"

    Public Property tnx As tnxModel
    Public Property connections As DataTransfererCCIplate
    Public Property pole As DataTransfererCCIpole
    Public Property structureCodeCriteria As SiteCodeCriteria
    Public Property PierandPads As New List(Of PierAndPad)
    Public Property Piles As New List(Of Pile)
    Public Property UnitBases As New List(Of UnitBase)
    'Public Property UnitBases As New List(Of SST_Unit_Base) 'Challs version - DNU
    Public Property DrilledPiers As New List(Of DrilledPier)
    Public Property GuyAnchorBlocks As New List(Of GuyedAnchorBlock)

    Public Property reportOptions As ReportOptions

    'The structure class should return itself if the parent is requested
    Private _ParentStructure As EDSStructure
    Public Overrides Property ParentStructure As EDSStructure
        Get
            Return Me
        End Get
        Set(value As EDSStructure)
            _ParentStructure = value
        End Set
    End Property


#Region "Constructors"
    Public Sub New()
        'Leave method empty
    End Sub

    Public Sub New(ByVal BU As String, ByVal structureID As String, filePaths As String())
        Me.bus_unit = BU
        Me.structure_id = structureID
        'Uncomment your foundation type for testing when it's ready. 
        LoadFromFiles(filePaths)
    End Sub
    Public Sub New(ByVal BU As String, ByVal structureID As String, ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String)
        Me.bus_unit = BU
        Me.structure_id = structureID
        Me.databaseIdentity = LogOnUser
        Me.activeDatabase = ActiveDatabase

        LoadFromEDS(BU, structureID, LogOnUser, ActiveDatabase)
    End Sub

    Public Overrides Function ToString() As String
        Return Me.bus_unit & " - " & Me.structure_id
    End Function
#End Region

#Region "EDS"
    Public Sub LoadFromEDS(ByVal BU As String, ByVal structureID As String, ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String)

        Dim query As String = QueryBuilderFromFile(queryPath & "Structure\Structure (SELECT).sql").Replace("[BU]", BU.FormatDBValue()).Replace("[STRID]", structureID.FormatDBValue())
        Dim tableNames() As String = {"TNX", "Base Structure", "Upper Structure", "Guys", "Members", "Materials", "Pier and Pad", "Unit Base", "Pile", "Drilled Pier", "Anchor Block", "Soil Profiles", "Soil Layers", "Connections", "Pole", "Site Code Criteria"}

        Using strDS As New DataSet

            sqlLoader(query, strDS, ActiveDatabase, LogOnUser, 500)

            'name tables from tableNames list
            For i = 0 To strDS.Tables.Count - 1
                strDS.Tables(i).TableName = tableNames(i)
            Next

            'If no site code criteria exists, fetch data from ORACLE to use for the first analysis. 
            'Still need to find all Topo inputs
            'Just set other parameters as default values 
            If Not strDS.Tables("Site Code Criteria").Rows.Count > 0 Then
                OracleLoader("
                    SELECT
                            str.bus_unit
                            ,str.structure_id
                            ,tr.standard_code tia_current
                            ,tr.bldg_code ibc_current
                            ,str.ground_elev elev_agl
                            ,str.hgt_no_appurt
                            ,str.crest_height
                            ,str.distance_from_crest
                            ,sit.site_name
                            ,'False' rev_h_section_15_5
                            ,0 tower_point_elev
                            --,pi.eng_app_id
                            --,pi.crrnt_rvsn_num
                        FROM
                            isit_aim.structure                      str
                            ,isit_aim.site                          sit
                            ,rpt_appl.eng_tower_rating_vw           tr
                            --,isit_aim.work_orders                 wo
                            --,isit_isite.project_info              pi
                        WHERE
                            --wo.work_order_seqnum = 'XXXXXXX'
                            str.bus_unit = '" & bus_unit & "' --Comment out when switching to WO
                            AND str.structure_id = '" & structure_id & "' --Comment out when switching to WO
                            AND str.bus_unit = sit.bus_unit
                            AND str.bus_unit = tr.bus_unit
                            --AND wo.bus_unit = str.bus_unit
                            --AND wo.structure_id = str.structure_id
                            --AND pi.eng_app_id = wo.eng_app_id(+)

                    ", "Site Code Criteria", strDS, 3000, "ords")
            End If
            Me.structureCodeCriteria = New SiteCodeCriteria(strDS.Tables("Site Code Criteria").Rows(0))

            'Load TNX Model
            'Me.tnx = New tnxModel(strDS, Me)

            'Pier and Pad
            For Each dr As DataRow In strDS.Tables("Pier and Pad").Rows
                Me.PierandPads.Add(New PierAndPad(dr, Me))
            Next

            'Unit Base (CHall - DNU)
            'For Each dr As DataRow In strDS.Tables("Unit Base").Rows
            '    Me.UnitBases.Add(New SST_Unit_Base(dr, Me))
            'Next

            'Unit Base
            For Each dr As DataRow In strDS.Tables("Unit Base").Rows
                Me.UnitBases.Add(New UnitBase(dr, Me))
            Next

            'For additional tools we'll need to update the constructor to use a datarow and pass through the dataset byref for sub tables (i.e. soil profiles)
            'That constructor will grab datarows from the sub data tables based on the foreign key in datarow
            'For Each dr As DataRow In strDS.Tables("Drilled Pier").Rows
            '    Me.DrilledPiers.Add(New DrilledPier(dr, strDS))
            'Next

        End Using

    End Sub


    Public Sub SavetoEDS(ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String)

        Me.databaseIdentity = LogOnUser
        Me.activeDatabase = ActiveDatabase

        Dim existingStructure As New EDSStructure(Me.bus_unit, Me.structure_id, Me.databaseIdentity, Me.activeDatabase)

        Dim structureQuery As String =
            "DECLARE @Prev TABLE(ID INT)" & vbCrLf &
            "DECLARE @PrevID INT" & vbCrLf &
            "BEGIN TRANSACTION" & vbCrLf

        'structureQuery += Me.tnx.EDSQuery(existingStructure.tnx)
        structureQuery += Me.PierandPads.EDSListQuery(existingStructure.PierandPads)
        structureQuery += Me.UnitBases.EDSListQuery(existingStructure.UnitBases)
        'structureQuery += Me.Piles.EDSListQuery(existingStructure.PierandPads)
        'structureQuery += Me.DrilledPiers.EDSListQuery(existingStructure.PierandPads)
        'structureQuery += Me.GuyAnchorBlocks.EDSListQuery(existingStructure.PierandPads)
        'structureQuery += Me.connections.EDSQuery(existingStructure.PierandPads)
        'structureQuery += Me.pole.EDSQuery(existingStructure.PierandPads)

        structureQuery += "COMMIT"

        'MessageBox.Show(structureQuery)

        sqlSender(structureQuery, ActiveDatabase, LogOnUser, 0.ToString)

    End Sub
#End Region

#Region "Excel"
    Public Sub LoadFromFiles(filePaths As String())

        For Each item As String In filePaths
            If item.EndsWith(".eri") Then
                Me.tnx = New tnxModel(item)
            ElseIf item.Contains("Pier and Pad Foundation") Then
                Me.PierandPads.Add(New PierAndPad(item, Me))
            ElseIf item.Contains("Pile Foundation") Then
                'Me.Piles.Add(New Pile(item))
            ElseIf item.Contains("SST Unit Base Foundation") Then
                'Me.UnitBases.Add(New SST_Unit_Base(item, Me)) 'Chall version - DNU
                Me.UnitBases.Add(New UnitBase(item, Me))
            ElseIf item.Contains("Drilled Pier Foundation") Then
                'Me.DrilledPiers.Add(New DrilledPier(item))
            ElseIf item.Contains("Guyed Anchor Block Foundation") Then
                'Me.GuyAnchorBlocks.Add(New GuyedAnchorBlock(item))
            End If
        Next
    End Sub

    Public Sub SaveTools(folderPath As String)
        'Uncomment your foundation type for testing when it's ready.
        Dim i As Integer
        Dim fileNum As String = ""

        If Me.tnx IsNot Nothing Then Me.tnx.GenerateERI(Path.Combine(folderPath, Me.bus_unit & ".eri"))

        For i = 0 To Me.PierandPads.Count - 1
            'I think we need a better way to get filename and maintain meaningful file names after they've gone through the database.
            'This works for now, just basing the name off the template name.
            fileNum = String.Format(" ({0})", i.ToString)
            PierandPads(i).workBookPath = Path.Combine(folderPath, Me.bus_unit & "_" & Path.GetFileNameWithoutExtension(PierandPads(i).templatePath) & "_EDS_" & fileNum & Path.GetExtension(PierandPads(i).templatePath))
            PierandPads(i).SavetoExcel()
        Next
        'For i = 0 To Me.Piles.Count - 1
        '    fileNum = If(i = 0, "", Format(" ({0})", i.ToString))
        '    Piles(i).workBookPath = Path.Combine(folderPath, Path.GetFileName(Piles(i).templatePath) & fileNum)
        '    Piles(i).SavetoExcel()
        'Next
        For i = 0 To Me.UnitBases.Count - 1
            fileNum = String.Format(" ({0})", i.ToString)
            UnitBases(i).workBookPath = Path.Combine(folderPath, Me.bus_unit & "_" & Path.GetFileNameWithoutExtension(UnitBases(i).templatePath) & "_EDS_" & fileNum & Path.GetExtension(UnitBases(i).templatePath))
            UnitBases(i).SavetoExcel()
        Next
        'For i = 0 To Me.DrilledPiers.Count - 1
        '    fileNum = If(i = 0, "", Format(" ({0})", i.ToString))
        '    DrilledPiers(i).workBookPath = Path.Combine(folderPath, Path.GetFileName(DrilledPiers(i).templatePath) & fileNum)
        '    DrilledPiers(i).SavetoExcel()
        'Next
        'For i = 0 To Me.GuyAnchorBlocks.Count - 1
        '    fileNum = If(i = 0, "", Format(" ({0})", i.ToString))
        '    GuyAnchorBlocks(i).workBookPath = Path.Combine(folderPath, Path.GetFileName(GuyAnchorBlocks(i).templatePath) & fileNum)
        '    GuyAnchorBlocks(i).SavetoExcel()
        'Next
    End Sub
#End Region

#Region "Check Changes"
    Public Function CompareEDS(other As EDSObject, Optional ByRef changes As List(Of AnalysisChange) = Nothing) As Boolean

    End Function

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Throw New NotImplementedException()
    End Function

#End Region
End Class

Partial Public Class EDSResult
    Inherits EDSObject

    Public Overrides ReadOnly Property EDSObjectName As String = "Result"

    Private _foreign_key As Integer?
    Private _result_lkup As String
    Private _rating As Double?
    Private _Insert As String
    Private _EDSTableName As String
    Private _ForeignKeyName As String
    'modified_person_id
    'process_stag
    'modified_date

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
    Public Property rating() As Double?
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


    Public ReadOnly Property Insert(Optional ByVal ResultsParentIDKnown As Boolean = True) As String
        Get
            Insert =
                "BEGIN" & vbCrLf &
                "  INSERT INTO [TABLE] ([FIELDS])" & vbCrLf &
                "  VALUES([VALUES])" & vbCrLf &
                "END" & vbCrLf
            Insert = Insert.Replace("[TABLE]", Me.EDSTableName)
            Insert = Insert.Replace("[VALUES]", Me.SQLInsertValues(ResultsParentIDKnown))
            Insert = Insert.Replace("[FIELDS]", Me.SQLInsertFields)
            Return Insert
        End Get
    End Property

    'Public ReadOnly Property InsertQuery() As String
    '    Get
    '        InsertQuery =
    '            "BEGIN" & vbCrLf &
    '            "  INSERT INTO [TABLE] ([FIELDS])" & vbCrLf &
    '            "  VALUES([VALUES])" & vbCrLf &
    '            "END" & vbCrLf
    '        InsertQuery = InsertQuery.Replace("[TABLE]", Me.EDSTableName)
    '        InsertQuery = InsertQuery.Replace("[VALUES]", Me.SQLInsertValues(True))
    '        InsertQuery = InsertQuery.Replace("[FIELDS]", Me.SQLInsertFields)
    '        Return InsertQuery
    '    End Get
    'End Property


    Public Function SQLInsertValues(Optional ByVal ResultsParentIDKnown As Boolean = True) As String
        SQLInsertValues = ""

        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.work_order_seq_num.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(If(ResultsParentIDKnown, Me.foreign_key.ToString.FormatDBValue, "@PrevID"))
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

    Public Sub New(ByVal resultDr As DataRow, ByRef Parent As EDSObjectWithQueries)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then
            Me.Absorb(Parent)
            Me._foreign_key = Parent.ID
            'Results table should be the Parent Table Name + _results (fnd.pier_pad -> fnd.pier_pad_results)
            Me.EDSTableName = Parent.EDSTableName & "_results"
            'Result ID name should be Parent Table Name + _id (fnd.pier_pad -> pier_pad_id)
            'Seperate the table name from the schema then add _id
            'MessageBox.Show("Start:" & (Parent.EDSTableName.IndexOf(".") + 1).ToString)
            'MessageBox.Show("Length:" & (Parent.EDSTableName.Length - Parent.EDSTableName.IndexOf(".") - 1).ToString)
            Me.ForeignKeyName = If(Parent.EDSTableName.Contains("."),
                                    Parent.EDSTableName.Substring(Parent.EDSTableName.IndexOf(".") + 1, Parent.EDSTableName.Length - Parent.EDSTableName.IndexOf(".") - 1) & "_id",
                                    Parent.EDSTableName & "_id")
        End If

        Me.result_lkup = DBtoStr(resultDr.Item("result_lkup"))
        Me.rating = DBtoNullableDbl(resultDr.Item("rating"))
    End Sub

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Throw New NotImplementedException()
    End Function

End Class

Partial Public Class SiteCodeCriteria

    Private _ID As Integer?
    Private _bus_unit As String
    Private _ibc_current As String
    Private _asce_current As String
    Private _tia_current As String
    Private _rev_h_accepted As Boolean?
    Private _rev_h_section_15_5 As Boolean?
    Private _seismic_design_category As String
    Private _frost_depth_tia_g As Double?
    Private _elev_agl As Double?
    Private _topo_category As Integer?
    Private _expo_category As String
    Private _crest_height As Double?
    Private _slope_distance As Double?
    Private _distance_from_crest As Double?
    Private _downwind As Boolean?
    Private _topo_feature As String
    Private _crest_point_elev As Double?
    Private _base_point_elev As Double?
    Private _mid_height_elev As Double?
    Private _crest_to_mid_height_distance As Double?
    Private _tower_point_elev As Double?
    Private _base_kzt As Double?

    <Category(""), Description(""), DisplayName("ID")>
    Public Property ID() As Integer?
        Get
            Return Me._ID
        End Get
        Set
            Me._ID = Value
        End Set
    End Property
    <Category(""), Description("Member Type"), DisplayName("bus_unit")>
    Public Property bus_unit() As String
        Get
            Return Me._bus_unit
        End Get
        Set
            Me._bus_unit = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("ibc_current")>
    Public Property ibc_current() As String
        Get
            Return Me._ibc_current
        End Get
        Set
            Me._ibc_current = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("asce_current")>
    Public Property asce_current() As String
        Get
            Return Me._asce_current
        End Get
        Set
            Me._asce_current = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("tia_current")>
    Public Property tia_current() As String
        Get
            Return Me._tia_current
        End Get
        Set
            Me._tia_current = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("rev_h_accepted")>
    Public Property rev_h_accepted() As Boolean?
        Get
            Return Me._rev_h_accepted
        End Get
        Set
            Me._rev_h_accepted = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("rev_h_section_15_5")>
    Public Property rev_h_section_15_5() As Boolean?
        Get
            Return Me._rev_h_section_15_5
        End Get
        Set
            Me._rev_h_section_15_5 = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("seismic_design_category")>
    Public Property seismic_design_category() As String
        Get
            Return Me._seismic_design_category
        End Get
        Set
            Me._seismic_design_category = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("frost_depth_tia_g")>
    Public Property frost_depth_tia_g() As Double?
        Get
            Return Me._frost_depth_tia_g
        End Get
        Set
            Me._frost_depth_tia_g = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("elev_agl")>
    Public Property elev_agl() As Double?
        Get
            Return Me._elev_agl
        End Get
        Set
            Me._elev_agl = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("topo_category")>
    Public Property topo_category() As Integer?
        Get
            Return Me._topo_category
        End Get
        Set
            Me._topo_category = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("expo_category")>
    Public Property expo_category() As String
        Get
            Return Me._expo_category
        End Get
        Set
            Me._expo_category = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("crest_height")>
    Public Property crest_height() As Double?
        Get
            Return Me._crest_height
        End Get
        Set
            Me._crest_height = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("slope_distance")>
    Public Property slope_distance() As Double?
        Get
            Return Me._slope_distance
        End Get
        Set
            Me._slope_distance = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("distance_from_crest")>
    Public Property distance_from_crest() As Double?
        Get
            Return Me._distance_from_crest
        End Get
        Set
            Me._distance_from_crest = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("downwind")>
    Public Property downwind() As Boolean?
        Get
            Return Me._downwind
        End Get
        Set
            Me._downwind = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("topo_feature")>
    Public Property topo_feature() As String
        Get
            Return Me._topo_feature
        End Get
        Set
            Me._topo_feature = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("crest_point_elev")>
    Public Property crest_point_elev() As Double?
        Get
            Return Me._crest_point_elev
        End Get
        Set
            Me._crest_point_elev = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("base_point_elev")>
    Public Property base_point_elev() As Double?
        Get
            Return Me._base_point_elev
        End Get
        Set
            Me._base_point_elev = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("mid_height_elev")>
    Public Property mid_height_elev() As Double?
        Get
            Return Me._mid_height_elev
        End Get
        Set
            Me._mid_height_elev = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("crest_to_mid_height_distance")>
    Public Property crest_to_mid_height_distance() As Double?
        Get
            Return Me._crest_to_mid_height_distance
        End Get
        Set
            Me._crest_to_mid_height_distance = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("tower_point_elev")>
    Public Property tower_point_elev() As Double?
        Get
            Return Me._tower_point_elev
        End Get
        Set
            Me._tower_point_elev = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("base_kzt")>
    Public Property base_kzt() As Double?
        Get
            Return Me._base_kzt
        End Get
        Set
            Me._base_kzt = Value
        End Set
    End Property

#Region "Constructors"
    Public Sub New()
        'Variables need to be passed into another constructor
        'Using this just as an example and assuming BU & structure ID exist
    End Sub

    Public Sub New(ByVal SiteCodeDataRow As DataRow)
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("bus_unit"), String)) Then
                Me.bus_unit = CType(SiteCodeDataRow.Item("bus_unit"), String)
            Else
                Me.bus_unit = Nothing
            End If
        Catch ex As Exception
            Me.bus_unit = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("ibc_current"), String)) Then
                Me.ibc_current = CType(SiteCodeDataRow.Item("ibc_current"), String)
            Else
                Me.ibc_current = Nothing
            End If
        Catch ex As Exception
            Me.ibc_current = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("asce_current"), String)) Then
                Me.asce_current = CType(SiteCodeDataRow.Item("asce_current"), String)
            Else
                Me.asce_current = Nothing
            End If
        Catch ex As Exception
            Me.asce_current = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("tia_current"), String)) Then
                Me.tia_current = CType(SiteCodeDataRow.Item("tia_current"), String)
            Else
                Me.tia_current = Nothing
            End If
        Catch ex As Exception
            Me.tia_current = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("rev_h_accepted"), Boolean)) Then
                Me.rev_h_accepted = CType(SiteCodeDataRow.Item("rev_h_accepted"), Boolean)
            Else
                Me.rev_h_accepted = Nothing
            End If
        Catch ex As Exception
            Me.rev_h_accepted = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("rev_h_section_15_5"), Boolean)) Then
                Me.rev_h_section_15_5 = CType(SiteCodeDataRow.Item("rev_h_section_15_5"), Boolean)
            Else
                Me.rev_h_section_15_5 = Nothing
            End If
        Catch ex As Exception
            Me.rev_h_section_15_5 = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("seismic_design_category"), Boolean)) Then
                Me.seismic_design_category = CType(SiteCodeDataRow.Item("seismic_design_category"), Boolean)
            Else
                Me.seismic_design_category = Nothing
            End If
        Catch ex As Exception
            Me.seismic_design_category = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("frost_depth_tia_g"), Double)) Then
                Me.frost_depth_tia_g = CType(SiteCodeDataRow.Item("frost_depth_tia_g"), Double)
            Else
                Me.frost_depth_tia_g = Nothing
            End If
        Catch ex As Exception
            Me.frost_depth_tia_g = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("elev_agl"), Double)) Then
                Me.elev_agl = CType(SiteCodeDataRow.Item("elev_agl"), Double)
            Else
                Me.elev_agl = Nothing
            End If
        Catch ex As Exception
            Me.elev_agl = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("topo_category"), Integer)) Then
                Me.topo_category = CType(SiteCodeDataRow.Item("topo_category"), Integer)
            Else
                Me.topo_category = Nothing
            End If
        Catch ex As Exception
            Me.topo_category = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("expo_category"), String)) Then
                Me.expo_category = CType(SiteCodeDataRow.Item("expo_category"), String)
            Else
                Me.expo_category = Nothing
            End If
        Catch ex As Exception
            Me.expo_category = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("crest_height"), Double)) Then
                Me.crest_height = CType(SiteCodeDataRow.Item("crest_height"), Double)
            Else
                Me.crest_height = Nothing
            End If
        Catch ex As Exception
            Me.crest_height = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("slope_distance"), Double)) Then
                Me.slope_distance = CType(SiteCodeDataRow.Item("slope_distance"), Double)
            Else
                Me.slope_distance = Nothing
            End If
        Catch ex As Exception
            Me.slope_distance = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("distance_from_crest"), Double)) Then
                Me.distance_from_crest = CType(SiteCodeDataRow.Item("distance_from_crest"), Double)
            Else
                Me.distance_from_crest = Nothing
            End If
        Catch ex As Exception
            Me.distance_from_crest = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("downwind"), Boolean)) Then
                Me.downwind = CType(SiteCodeDataRow.Item("downwind"), Boolean)
            Else
                Me.downwind = Nothing
            End If
        Catch ex As Exception
            Me.downwind = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("topo_feature"), String)) Then
                Me.topo_feature = CType(SiteCodeDataRow.Item("topo_feature"), String)
            Else
                Me.topo_feature = Nothing
            End If
        Catch ex As Exception
            Me.topo_feature = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("crest_point_elev"), Double)) Then
                Me.crest_point_elev = CType(SiteCodeDataRow.Item("crest_point_elev"), Double)
            Else
                Me.crest_point_elev = Nothing
            End If
        Catch ex As Exception
            Me.crest_point_elev = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("base_point_elev"), Double)) Then
                Me.base_point_elev = CType(SiteCodeDataRow.Item("base_point_elev"), Double)
            Else
                Me.base_point_elev = Nothing
            End If
        Catch ex As Exception
            Me.base_point_elev = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("mid_height_elev"), Double)) Then
                Me.mid_height_elev = CType(SiteCodeDataRow.Item("mid_height_elev"), Double)
            Else
                Me.mid_height_elev = Nothing
            End If
        Catch ex As Exception
            Me.mid_height_elev = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("crest_to_mid_height_distance"), Double)) Then
                Me.crest_to_mid_height_distance = CType(SiteCodeDataRow.Item("crest_to_mid_height_distance"), Double)
            Else
                Me.crest_to_mid_height_distance = Nothing
            End If
        Catch ex As Exception
            Me.crest_to_mid_height_distance = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("tower_point_elev"), Double)) Then
                Me.tower_point_elev = CType(SiteCodeDataRow.Item("tower_point_elev"), Double)
            Else
                Me.tower_point_elev = Nothing
            End If
        Catch ex As Exception
            Me.tower_point_elev = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("base_kzt"), Double)) Then
                Me.base_kzt = CType(SiteCodeDataRow.Item("base_kzt"), Double)
            Else
                Me.base_kzt = Nothing
            End If
        Catch ex As Exception
            Me.base_kzt = Nothing
        End Try

    End Sub
#End Region

End Class








