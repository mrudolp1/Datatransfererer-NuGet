Imports System.ComponentModel
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

    <Extension()>
    Public Function GetDistinct(Of T As EDSObject)(alist As List(Of T)) As List(Of T)
        'Notes: Removes duplicates from list of tnxDatabaseEntry by using their CompareMe function
        'Making this generic (Of T As tnxDatabaseEntry) allows it to work for all subclasses of tnxDatabaseEntry

        Dim distinctList As New List(Of T)

        For Each item In alist
            Dim addToList As Boolean = True
            For Each distinctItem In distinctList
                If item.CompareMe(distinctItem) Then
                    'Not distinct
                    addToList = False
                    Exit For
                End If
            Next
            If addToList Then distinctList.Add(item)
        Next

        Return distinctList
    End Function

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
                    If Not currentSortedList(i).CompareMe(prevSortedList(i)) Then
                        EDSListQuery += currentSortedList(i).Update
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
    Public Iterator Function Add(Of T As ObjectsComparer.Difference)(ByVal e As IEnumerable(Of T), ByVal value As T, Optional ByVal Path As String = Nothing) As IEnumerable(Of T)
        'Allow you to add to an IEnumerable like it is a list.
        'Useful for working with the ObjectComparer class which stores the differences as IEnumerable(of Difference)
        'Refernce: https://stackoverflow.com/a/1210311
        For Each cur In e
            Yield cur
        Next

        If Path IsNot Nothing Then
            Yield value.InsertPath(Path)
        Else
            Yield value
        End If
    End Function

    <Extension()>
    Public Iterator Function Add(Of T As ObjectsComparer.Difference)(ByVal e1 As IEnumerable(Of T), ByVal e2 As IEnumerable(Of T), Optional ByVal Path As String = Nothing) As IEnumerable(Of T)
        'Allow you to add to an IEnumerable to another IEnumerable.

        For Each cur In e1
            Yield cur
        Next

        For Each cur In e2
            If Path IsNot Nothing Then
                Yield cur.InsertPath(Path)
            Else
                Yield cur
            End If
        Next

    End Function

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
    Implements IEquatable(Of EDSObject), IComparable(Of EDSObject)

    Public Property ID As Integer?
    Public Overridable Property Parent As EDSObject
    Public Overridable Property ParentStructure As EDSStructure
    Public Property bus_unit As String
    Public Property structure_id As String
    Public Property work_order_seq_num As String
    Public Property activeDatabase As String
    Public Property databaseIdentity As WindowsIdentity
    Public Property differences As List(Of ObjectsComparer.Difference)

    Public Overridable Function CreateChangeSummary() As String
        Dim summary As String = ""

        For Each chng As AnalysisChanges In changeList
            summary += chng.CategoryName & " " & chng.FieldName & " = " & chng.NewValue & " | Previously: " & chng.PreviousValue & vbNewLine
        Next

        Return summary

    End Function



    Public Overridable Function CompareMe(Of T As EDSObject)(toCompare As T) As Boolean
        'Compare another EDSObject object to itself using the objects comparer.
        'Making this generic (Of T As EDSObject) allows it to work for all subclasses of EDSObject

        If toCompare Is Nothing Then Return False

        Dim comparer As New ObjectsComparer.Comparer(Of T)()

        Dim differences As IEnumerable(Of ObjectsComparer.Difference) = Nothing

        CompareMe = comparer.Compare(CType(Me, T), toCompare, differences)

        Me.differences = differences.ToList

        Return CompareMe

    End Function

    Public Overridable Sub Absorb(ByRef Host As EDSObject)
        Me.Parent = Host
        Me.ParentStructure = If(Host.ParentStructure, Nothing) 'The parent of an EDSObject should be the top level structure.
        Me.bus_unit = Host.bus_unit
        Me.structure_id = Host.structure_id
        Me.work_order_seq_num = Host.work_order_seq_num
        Me.activeDatabase = Host.activeDatabase
        Me.databaseIdentity = Host.databaseIdentity
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

    Public Function Equals(other As EDSObject) As Boolean Implements IEquatable(Of EDSObject).Equals
        'Not currently using this but we could implement the whole compare function here

        If other Is Nothing Then Return False

        Return Me.CompareMe(other)

    End Function
End Class

Partial Public MustInherit Class EDSObjectWithQueries
    Inherits EDSObject

    Public MustOverride ReadOnly Property EDSTableName As String
    Public Overridable ReadOnly Property EDSQueryPath As String = IO.Path.Combine(My.Application.Info.DirectoryPath, "Templates")
    Public Overridable ReadOnly Property Insert() As String
        Get
            Insert = "BEGIN" & vbCrLf &
                     "  INSERT INTO [TABLE] ([FIELDS])" & vbCrLf &
                     "  VALUES([VALUES])" & vbCrLf &
                     "END"
            Insert = Insert.Replace("[TABLE]", Me.EDSTableName.FormatDBValue)
            Insert = Insert.Replace("[FIELDS]", Me.SQLInsertFields)
            Insert = Insert.Replace("[VALUES]", Me.SQLInsertValues)
            Return Insert
        End Get
    End Property

    Public Overridable ReadOnly Property Update() As String
        Get
            Update = "BEGIN" & vbCrLf &
                      "  Update [Table]" &
                      "  SET [UPDATE]" & vbCrLf &
                      "  WHERE ID = [ID]" & vbCrLf &
                      "END"
            Update = Update.Replace("[TABLE]", Me.EDSTableName.FormatDBValue)
            Update = Update.Replace("[UPDATE]", Me.SQLUpdate)
            Update = Update.Replace("[ID]", Me.ID)
            Return Update
        End Get
    End Property
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

    Public Overridable Function EDSQuery(Of T As EDSObjectWithQueries)(item As T, prevItem As T) As String
        'Compare the ID of the current EDS item to the existing item and determine if the Insert, Update, or Delete query should be used

        EDSQuery = ""

        If prevItem.ID = item.ID And Not item.CompareMe(prevItem) Then
            EDSQuery += item.Update
        Else
            'Need to add inserted items to comparison list.
            EDSQuery += item.Insert
            If prevItem IsNot Nothing Then
                EDSQuery += prevItem.Delete
            End If
        End If

        Return EDSQuery

    End Function

End Class

Partial Public MustInherit Class EDSExcelObject
    'This should be inherited by the main tool class. Subclasses such as soil layers can probably inherit the EDSObjectWithQueries
    Inherits EDSObjectWithQueries

    Public Property workBookPath As String
    Public MustOverride ReadOnly Property templatePath As String
    Public Property fileType As DocumentFormat = DocumentFormat.Xlsm
    Public MustOverride ReadOnly Property excelDTParams As List(Of EXCELDTParameter)
    Public Property Results As New List(Of EDSResult)

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

Partial Public MustInherit Class EDSFoundation
    Inherits EDSExcelObject

    Public MustOverride ReadOnly Property foundationType As String

End Class

Partial Public Class EDSStructure
    Inherits EDSObject

    Public Property tnx As tnxModel
    Public Property structureCodeCriteria As SiteCodeCriteria
    Public Property PierandPads As New List(Of PierAndPad)
    Public Property Piles As New List(Of Pile)
    Public Property UnitBases As New List(Of UnitBase)
    Public Property DrilledPiers As New List(Of DrilledPier)
    Public Property GuyAnchorBlocks As New List(Of GuyedAnchorBlock)
    Public Property connections As DataTransfererCCIplate
    Public Property pole As DataTransfererCCIpole

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
        Dim tableNames() As String = {"TNX", "Base Structure", "Upper Structure", "Guys", "Members", "Materials", "Pier and Pad", "Unit Base", "Pile", "Drilled Pier", "Anchor Block", "Soil Profiles", "Soil Layers", "Connections", "Pole"}


        Using strDS As New DataSet

            sqlLoader(query, strDS, ActiveDatabase, LogOnUser, 500)

            'name tables from tableNames list
            For i = 0 To strDS.Tables.Count - 1
                strDS.Tables(i).TableName = tableNames(i)
            Next

            'Load TNX Model
            'Me.tnx = New tnxModel(strDS, Me)

            'Pier and Pad
            For Each dr As DataRow In strDS.Tables("Pier and Pad").Rows
                Me.PierandPads.Add(New PierAndPad(dr, Me))
            Next

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

        Dim structureQuery As String = ""
        'structureQuery += Me.tnx.EDSQuery(existingStructure.tnx)
        structureQuery += Me.PierandPads.EDSListQuery(existingStructure.PierandPads)
        structureQuery += Me.UnitBases.EDSListQuery(existingStructure.UnitBases)
        'structureQuery += Me.Piles.EDSListQuery(existingStructure.PierandPads)
        'structureQuery += Me.DrilledPiers.EDSListQuery(existingStructure.PierandPads)
        'structureQuery += Me.GuyAnchorBlocks.EDSListQuery(existingStructure.PierandPads)
        'structureQuery += Me.connections.EDSQuery(existingStructure.PierandPads)
        'structureQuery += Me.pole.EDSQuery(existingStructure.PierandPads)

        MessageBox.Show(structureQuery)

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
            fileNum = If(i = 0, "", String.Format(" ({0})", i.ToString))
            PierandPads(i).workBookPath = Path.Combine(folderPath, Me.bus_unit & "_" & Path.GetFileNameWithoutExtension(PierandPads(i).templatePath) & fileNum & Path.GetExtension(PierandPads(i).templatePath))
            PierandPads(i).SavetoExcel()
        Next
        'For i = 0 To Me.Piles.Count - 1
        '    fileNum = If(i = 0, "", Format(" ({0})", i.ToString))
        '    Piles(i).workBookPath = Path.Combine(folderPath, Path.GetFileName(Piles(i).templatePath) & fileNum)
        '    Piles(i).SavetoExcel()
        'Next
        For i = 0 To Me.UnitBases.Count - 1
            fileNum = If(i = 0, "", String.Format(" ({0})", i.ToString))
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

#End Region
End Class

Partial Public Class EDSResult
    Inherits EDSObject

    Private _foreign_key As Integer?
    Private _result_lkup As String
    Private _rating As Double?
    Private _Insert As String
    'modified_person_id
    'process_stage
    'modified_date

    <Category("Results"), Description("The ID of the parent object that this result is associated with. (i.e. Drilled Pier, Tower Leg, Plate)"), DisplayName("Foreign Key Reference")>
    Public Property foreign_key() As Integer?
        Get
            Return If(Me._foreign_key, Me.Parent.ID)
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

    Public Property Result_Table_Name() As String 'Need to set this from parent object (i.e. pier_pad_results)
    Public Property Result_ID_Name() As String 'Need to set this from parent object (i.e. pier_pad_id)

    Public ReadOnly Property Insert() As String
        Get
            Insert =
                "BEGIN" & vbCrLf &
                     "  INSERT INTO [TABLE] ([FIELDS])" & vbCrLf &
                     "  VALUES([VALUES])" & vbCrLf &
                     "END"
            Insert = Insert.Replace("[TABLE]", Me.Result_Table_Name)
            Insert = Insert.Replace("[VALUES]", Me.SQLInsertValues)
            Insert = Insert.Replace("[FIELDS]", Me.SQLInsertFields)
            Return Insert
        End Get
    End Property


    Public Function SQLInsertValues() As String
        SQLInsertValues = ""

        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.work_order_seq_num.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.foreign_key.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.result_lkup.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rating.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Function SQLInsertFields() As String
        SQLInsertFields = ""

        SQLInsertFields = SQLInsertFields.AddtoDBString("work_order_seq_num")
        SQLInsertFields = SQLInsertFields.AddtoDBString(Result_ID_Name)
        SQLInsertFields = SQLInsertFields.AddtoDBString("result_lkup")
        SQLInsertFields = SQLInsertFields.AddtoDBString("rating")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")

        Return SQLInsertFields
    End Function

    Public Sub New(ByVal resultDr As DataRow, ByRef Parent As EDSObject)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Me.result_lkup = DBtoStr(resultDr("result_lkup"))
        Me.rating = DBtoNullableDbl(resultDr("rating"))

    End Sub

End Class

Partial Public Class SiteCodeCriteria

    Private _ID As Integer?
    Private _bus_unit As Integer?
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
    Public Property bus_unit() As Integer?
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

End Class






