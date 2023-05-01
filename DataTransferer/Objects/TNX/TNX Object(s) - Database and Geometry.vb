Option Strict On
Option Compare Binary 'Trying to speed up parsing the TNX file by using Binary Text comparison instead of Text Comparison

Imports System.ComponentModel
Imports System.Data
Imports System.IO
Imports System.Security.Principal
Imports System.Runtime.CompilerServices
Imports System.Data.SqlClient
Imports System.Runtime.Serialization

Public Module TNXExtensions
    <Extension()>
    Public Function TNXMemberListQueryBuilder(Of T As tnxDatabaseObject)(alist As List(Of T), Optional prevList As List(Of T) = Nothing) As String

        TNXMemberListQueryBuilder = ""

        'Create a shallow copy of the lists for sorting and comparison
        'Sort lists by ID descending with Null IDs at the bottom
        Dim currentSortedList As List(Of T) = alist.ToList
        currentSortedList.Sort()
        'currentSortedList.Reverse()

        Dim prevSortedList As List(Of T)
        If prevList Is Nothing Then
            prevSortedList = New List(Of T)
        Else
            prevSortedList = prevList.ToList
            prevSortedList.Sort()
            'prevSortedList.Reverse()
        End If

        Dim i As Integer = 0
        'Remove items that exist in both lists so you're left with a list of items to insert and a list to delete
        Do While i < currentSortedList.Count
            Dim j As Integer = 0
            Do While j < prevSortedList.Count
                If currentSortedList(i).Equals(prevSortedList(j)) Then
                    currentSortedList.RemoveAt(i)
                    prevSortedList.RemoveAt(j)
                    'j = prevSortedList.Count - 1
                    i -= 1
                    Exit Do
                End If
                j += 1
            Loop
            i += 1
        Loop

        For Each current In currentSortedList
            TNXMemberListQueryBuilder += current.SQLInsert
        Next
        For Each prev In prevSortedList
            TNXMemberListQueryBuilder += prev.SQLDelete
        Next

        Return TNXMemberListQueryBuilder

    End Function

    <Extension()>
    Public Function TNXGeometryRecListQueryBuilder(Of T As tnxGeometryRec)(alist As List(Of T), Optional prevList As List(Of T) = Nothing, Optional ByVal AllowUpdate As Boolean = True) As String

        TNXGeometryRecListQueryBuilder = ""

        'Create a shallow copy of the lists for sorting and comparison
        'Sort lists by ID descending with Null IDs at the bottom
        Dim currentSortedList As List(Of T) = alist.ToList
        currentSortedList.Sort()
        'currentSortedList.Reverse()

        Dim prevSortedList As List(Of T)
        If prevList Is Nothing Then
            prevSortedList = New List(Of T)
        Else
            prevSortedList = prevList.ToList
            prevSortedList.Sort()
            'prevSortedList.Reverse()
        End If

        Dim i As Integer = 0
        Do While i <= Math.Max(currentSortedList.Count, prevSortedList.Count) - 1

            If i > currentSortedList.Count - 1 Then
                'Delete items in previous list if there is nothing left in current list
                TNXGeometryRecListQueryBuilder += prevSortedList(i).SQLDelete
            ElseIf i > prevSortedList.Count - 1 Then
                'Insert items in current list if there is nothing left in previous list
                TNXGeometryRecListQueryBuilder += currentSortedList(i).SQLInsert
            Else
                'Compare IDs
                If currentSortedList(i).Rec = prevSortedList(i).Rec And AllowUpdate Then
                    If Not currentSortedList(i).Equals(prevSortedList(i)) Then
                        'Update existing
                        TNXGeometryRecListQueryBuilder += currentSortedList(i).SQLUpdate
                    Else
                        'Save Results Only
                        TNXGeometryRecListQueryBuilder += currentSortedList(i).TNXResults.EDSResultQuery
                    End If
                ElseIf currentSortedList(i).Rec < prevSortedList(i).Rec Then
                    TNXGeometryRecListQueryBuilder += prevSortedList(i).SQLDelete
                    currentSortedList.Insert(i, Nothing)
                Else
                    'currentSortedList(i).ID > prevSortedList(i).ID
                    TNXGeometryRecListQueryBuilder += currentSortedList(i).SQLInsert
                    prevSortedList.Insert(i, Nothing)
                End If
            End If

            i += 1

        Loop

        Return TNXGeometryRecListQueryBuilder

    End Function
End Module

#Region "Database"
<DataContract()>
Partial Public Class tnxDatabase
    Inherits EDSObject

#Region "Inherits"
    Public Overrides ReadOnly Property EDSObjectName As String = "TNX Database"
#End Region

#Region "Define"
    Private _members As New List(Of tnxMember)
    Private _materials As New List(Of tnxMaterial)
    Private _bolts As New List(Of tnxMaterial)

    <Category("TNX Geometry"), Description("Upper Structure Type"), DisplayName("AntennaType")>
    <DataMember()> Public Property members() As List(Of tnxMember)
        Get
            Return Me._members
        End Get
        Set
            Me._members = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("Base Tower Type"), DisplayName("TowerType")>
    <DataMember()> Public Property materials() As List(Of tnxMaterial)
        Get
            Return Me._materials
        End Get
        Set
            Me._materials = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("Base Tower Type"), DisplayName("TowerType")>
    <DataMember()> Public Property bolts() As List(Of tnxMaterial)
        Get
            Return Me._bolts
        End Get
        Set
            Me._bolts = Value
        End Set
    End Property
#End Region

#Region "Constructors"
    Public Sub New(Optional ByVal Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)
    End Sub

#End Region

#Region "Equals"
    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As tnxDatabase = TryCast(other, tnxDatabase)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.members.CheckChange(otherToCompare.members, changes, categoryName, "Members"), Equals, False)
        Equals = If(Me.materials.CheckChange(otherToCompare.materials, changes, categoryName, "Materials"), Equals, False)
        Equals = If(Me.bolts.CheckChange(otherToCompare.bolts, changes, categoryName, "Bolts"), Equals, False)

        Return Equals
    End Function
#End Region

End Class
<DataContract()>
Partial Public MustInherit Class tnxDatabaseObject
    Inherits EDSObject

    <Category("General"), Description("EDS Table Name with schema."), DisplayName("Table Name")>
    Public MustOverride ReadOnly Property EDSTableName As String
    <Category("General"), Description("EDS Cross Reference Table."), DisplayName("Xref Table Name")>
    Public MustOverride ReadOnly Property EDSXrefTableName As String
    Public Overridable ReadOnly Property EDSQueryPath As String = IO.Path.Combine(My.Application.Info.DirectoryPath, "Templates")
    <Category("EDS Queries"), Description("Depth of table in EDS query. This determines where the ID is stored in the query and which parent ID is referenced if needed. 0 = Top Level"), Browsable(False)>
    Public Overridable ReadOnly Property EDSTableDepth As Integer = 1
    <Category("EDS Queries"), Description("Insert this object. For use in whole structure query."), DisplayName("SQL Insert Query")>
    Public MustOverride Function SQLInsert() As String

    <Category("EDS Queries"), Description("Delete this object and results from EDS. For use in whole structure query."), DisplayName("SQL Delete Query")>
    Public MustOverride Function SQLDelete() As String
    '<Category("EDS Queries"), Description("Insert this object. For use in whole structure query."), DisplayName("SQL Insert Query")>

    Public MustOverride Function SQLInsertValues() As String

    Public MustOverride Function SQLInsertFields() As String


End Class
<DataContract()>
Partial Public Class tnxMember
    Inherits tnxDatabaseObject

#Region "Inherits"
    Public Overrides ReadOnly Property EDSObjectName As String = "TNX Database Member"
    Public Overrides ReadOnly Property EDSTableName As String = "tnx.members"
    Public Overrides ReadOnly Property EDSXrefTableName As String = "tnx.members_xref"

    Private _Insert As String
    Private _Delete As String
    Public Overrides Function SQLInsert() As String
        Return Me.SQLInsert(Nothing)
    End Function
    Public Overloads Function SQLInsert(Optional ByVal ParentID As Integer? = Nothing) As String
        SQLInsert =
            "BEGIN" & vbCrLf &
            "   SELECT TOP 1 " & EDSStructure.SQLQueryIDVar(Me.EDSTableDepth) & " = tbl.ID" & vbCrLf &
            "   From tnx.members tbl" & vbCrLf &
            "   WHERE tbl.[File] = " & Me.File.ToString.FormatDBValue & vbCrLf &
            "   AND tbl.USName = " & Me.USName.ToString.FormatDBValue & vbCrLf &
            "   AND tbl.SIName = " & Me.SIName.ToString.FormatDBValue & vbCrLf &
            "   AND tbl.MatValues = " & Me.Values.ToString.FormatDBValue & vbCrLf &
            "   IF " & EDSStructure.SQLQueryIDVar(Me.EDSTableDepth) & " Is NULL" & vbCrLf &
            "       BEGIN" & vbCrLf &
            "           INSERT INTO tnx.members(" & Me.SQLInsertFields & ") OUTPUT INSERTED.ID INTO " & EDSStructure.SQLQueryTableVar(Me.EDSTableDepth) & " VALUES(" & Me.SQLInsertValues & ")" & vbCrLf &
            "           SELECT " & EDSStructure.SQLQueryIDVar(Me.EDSTableDepth) & "=ID FROM " & EDSStructure.SQLQueryTableVar(Me.EDSTableDepth) & vbCrLf &
            "       End" & vbCrLf &
            "   INSERT INTO tnx.members_xref(member_id, tnx_id) VALUES(" & EDSStructure.SQLQueryIDVar(Me.EDSTableDepth) & "," & If(ParentID Is Nothing, EDSStructure.SQLQueryIDVar(Me.EDSTableDepth - 1), ParentID.ToString.FormatDBValue) & ")" & vbCrLf &
            "   DELETE FROM " & EDSStructure.SQLQueryTableVar(Me.EDSTableDepth) & "" & vbCrLf &
            "END" & vbCrLf
        Return SQLInsert
    End Function

    Public Overrides Function SQLDelete() As String
        SQLDelete = "BEGIN" & vbCrLf &
                     "  DELETE FROM tnx.members_xref" & vbCrLf &
                     "  WHERE ID = " & Me.ID.ToString.FormatDBValue & vbCrLf &
                     "  AND tnx_id = " & Me.Parent.ID.ToString.FormatDBValue & vbCrLf &
                     "End" & vbCrLf
        Return SQLDelete
    End Function

    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.File.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.USName.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.SIName.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.Values.NullableToString.FormatDBValue)
        Return SQLInsertValues
    End Function
    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""
        SQLInsertFields = SQLInsertFields.AddtoDBString("[File]")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USName")
        SQLInsertFields = SQLInsertFields.AddtoDBString("SIName")
        SQLInsertFields = SQLInsertFields.AddtoDBString("MatValues")
        Return SQLInsertFields
    End Function

#End Region

#Region "Define"
    Private _File As String
    Private _USName As String
    Private _SIName As String
    Private _Values As String

    <Category("TNX Member Properties"), Description(""), DisplayName("File")>
    <DataMember()> Public Property File() As String
        Get
            Return Me._File
        End Get
        Set
            Me._File = Value
        End Set
    End Property
    <Category("TNX Material Properties"), Description(""), DisplayName("USName")>
    <DataMember()> Public Property USName() As String
        Get
            Return Me._USName
        End Get
        Set
            Me._USName = Value
        End Set
    End Property
    <Category("TNX Material Properties"), Description(""), DisplayName("SIName")>
    <DataMember()> Public Property SIName() As String
        Get
            Return Me._SIName
        End Get
        Set
            Me._SIName = Value
        End Set
    End Property
    <Category("TNX Material Properties"), Description(""), DisplayName("Values")>
    <DataMember()> Public Property Values() As String
        Get
            Return Me._Values
        End Get
        Set
            Me._Values = Value
        End Set
    End Property
#End Region

#Region "Constructors"
    Public Sub New(Optional ByVal Parent As EDSObject = Nothing)
        If Parent IsNot Nothing Then Me.Absorb(Parent)
    End Sub

    Public Sub New(data As DataRow, Optional ByVal Parent As EDSObject = Nothing)
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Me.ID = DBtoNullableInt(data.Item("ID"))
        Me.File = DBtoStr(data.Item("File"))
        Me.USName = DBtoStr(data.Item("USName"))
        Me.SIName = DBtoStr(data.Item("SIName"))
        Me.Values = DBtoStr(data.Item("MatValues"))

    End Sub
#End Region

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As tnxMember = TryCast(other, tnxMember)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.File.CheckChange(otherToCompare.File, changes, categoryName, "Member type"), Equals, False)
        Equals = If(Me.USName.CheckChange(otherToCompare.USName, changes, categoryName, "US Name"), Equals, False)
        Equals = If(Me.SIName.CheckChange(otherToCompare.SIName, changes, categoryName, "SI Name"), Equals, False)
        Equals = If(Me.Values.CheckChange(otherToCompare.Values, changes, categoryName, "Member values"), Equals, False)

        Return Equals
    End Function
End Class
<DataContract()>
Partial Public Class tnxMaterial
    Inherits tnxDatabaseObject

#Region "Inherits"
    Public Overrides ReadOnly Property EDSObjectName As String = "TNX Database Material"
    Public Overrides ReadOnly Property EDSTableName As String = "tnx.materials"
    Public Overrides ReadOnly Property EDSXrefTableName As String = "tnx.materials_xref"

    Public Overrides Function SQLInsert() As String
        Return Me.SQLInsert(Nothing)
    End Function
    Public Overloads Function SQLInsert(Optional ByVal ParentID As Integer? = Nothing) As String
        SQLInsert =
        "BEGIN" & vbCrLf &
        "   SELECT TOP 1 " & EDSStructure.SQLQueryIDVar(Me.EDSTableDepth) & " = tbl.ID" & vbCrLf &
        "   FROM tnx.materials tbl" & vbCrLf &
        "   WHERE tbl.IsBolt = " & Me.IsBolt.ToString.FormatDBValue & vbCrLf &
        "   AND tbl.MemberMatFile = " & Me.MemberMatFile.ToString.FormatDBValue & vbCrLf &
        "   AND tbl.MatName = " & Me.MatName.ToString.FormatDBValue & vbCrLf &
        "   AND tbl.MatValues = " & Me.MatValues.ToString.FormatDBValue & vbCrLf &
        "   IF " & EDSStructure.SQLQueryIDVar(Me.EDSTableDepth) & " Is NULL" & vbCrLf &
        "       BEGIN" & vbCrLf &
        "           INSERT INTO tnx.materials(" & Me.SQLInsertFields & ") OUTPUT INSERTED.ID INTO " & EDSStructure.SQLQueryTableVar(Me.EDSTableDepth) & " VALUES(" & Me.SQLInsertValues & ")" & vbCrLf &
        "           SELECT " & EDSStructure.SQLQueryIDVar(Me.EDSTableDepth) & "=ID FROM " & EDSStructure.SQLQueryTableVar(Me.EDSTableDepth) & vbCrLf &
        "       END" & vbCrLf &
        "   INSERT INTO tnx.materials_xref(material_id, tnx_id) VALUES(" & EDSStructure.SQLQueryIDVar(Me.EDSTableDepth) & "," & If(ParentID Is Nothing, EDSStructure.SQLQueryIDVar(Me.EDSTableDepth - 1), ParentID.ToString.FormatDBValue) & ")" & vbCrLf &
        "   DELETE FROM " & EDSStructure.SQLQueryTableVar(Me.EDSTableDepth) & vbCrLf &
        "   SET " & EDSStructure.SQLQueryIDVar(Me.EDSTableDepth) & " = NULL" & vbCrLf &
        "END" & vbCrLf
        Return SQLInsert
    End Function

    Public Overrides Function SQLDelete() As String
        SQLDelete = "BEGIN" & vbCrLf &
                 "  Delete FROM tnx.materials_xref WHERE ID = " & Me.ID.ToString.FormatDBValue & vbCrLf &
                 "End" & vbCrLf
        Return SQLDelete
    End Function

    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.IsBolt.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.MemberMatFile.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.MatName.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.MatValues.NullableToString.FormatDBValue)
        Return SQLInsertValues
    End Function
    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""
        SQLInsertFields = SQLInsertFields.AddtoDBString("IsBolt")
        SQLInsertFields = SQLInsertFields.AddtoDBString("MemberMatFile")
        SQLInsertFields = SQLInsertFields.AddtoDBString("MatName")
        SQLInsertFields = SQLInsertFields.AddtoDBString("MatValues")
        Return SQLInsertFields
    End Function
    'Public Overrides Function SQLUpdateFieldsandValues() As String
    '    SQLUpdateFieldsandValues = ""
    '    SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("IsBolt = False")
    '    SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("MemberMatFile = " & Me.MemberMatFile.ToString.FormatDBValue)
    '    SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("MatName = " & Me.MatName.ToString.FormatDBValue)
    '    SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("MatValues = " & Me.MatValues.ToString.FormatDBValue)
    '    Return SQLUpdateFieldsandValues
    'End Function
#End Region

#Region "Define"
    Private _MemberMatFile As String
    Private _MatName As String
    Private _MatValues As String
    Private _IsBolt As Boolean?

    <Category("TNX Material Properties"), Description(""), DisplayName("Material Type")>
    <DataMember()> Public Property MemberMatFile() As String
        Get
            Return Me._MemberMatFile
        End Get
        Set
            Me._MemberMatFile = Value
        End Set
    End Property
    <Category("TNX Material Properties"), Description(""), DisplayName("Material Name")>
    <DataMember()> Public Property MatName() As String
        Get
            Return Me._MatName
        End Get
        Set
            Me._MatName = Value
        End Set
    End Property
    <Category("TNX Material Properties"), Description(""), DisplayName("Material Values")>
    <DataMember()> Public Property MatValues() As String
        Get
            Return Me._MatValues
        End Get
        Set
            Me._MatValues = Value
        End Set
    End Property
    <Category("TNX Material Properties"), Description(""), DisplayName("Is Bolt Material?")>
    <DataMember()> Public Property IsBolt() As Boolean?
        Get
            Return Me._IsBolt
        End Get
        Set
            Me._IsBolt = Value
        End Set
    End Property
#End Region

#Region "Constructors"
    Public Sub New(Optional ByVal Parent As EDSObject = Nothing)
        If Parent IsNot Nothing Then Me.Absorb(Parent)
    End Sub

    Public Sub New(data As DataRow, Optional ByVal Parent As EDSObject = Nothing)
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Me.ID = DBtoNullableInt(data.Item("ID"))
        Me.MemberMatFile = DBtoStr(data.Item("MemberMatFile"))
        Me.MatName = DBtoStr(data.Item("MatName"))
        Me.MatValues = DBtoStr(data.Item("MatValues"))
        Me.IsBolt = DBtoNullableBool(data.Item("IsBolt"))

    End Sub

#End Region

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As tnxMaterial = TryCast(other, tnxMaterial)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.MemberMatFile.CheckChange(otherToCompare.MemberMatFile, changes, categoryName, "Material type"), Equals, False)
        Equals = If(Me.MatName.CheckChange(otherToCompare.MatName, changes, categoryName, "Material Name"), Equals, False)
        Equals = If(Me.MatValues.CheckChange(otherToCompare.MatValues, changes, categoryName, "Material properties"), Equals, False)
        Equals = If(Me.IsBolt.CheckChange(otherToCompare.IsBolt, changes, categoryName, "Is Bolt Material"), Equals, False)

        Return Equals
    End Function
End Class

#End Region

#Region "Geometry"
<DataContract()>
Partial Public Class tnxGeometry
    Inherits EDSObject

    Public Overrides ReadOnly Property EDSObjectName As String = "Geometry"

#Region "Define"
    Private _TowerType As String
    Private _AntennaType As String
    Private _OverallHeight As Double?
    Private _BaseElevation As Double?
    Private _Lambda As Double?
    Private _TowerTopFaceWidth As Double?
    Private _TowerBaseFaceWidth As Double?
    Private _TowerTaper As String
    Private _GuyedMonopoleBaseType As String
    Private _TaperHeight As Double?
    Private _PivotHeight As Double?
    Private _AutoCalcGH As Boolean?
    Private _UserGHElev As Double?
    Private _UseIndexPlate As Boolean?
    Private _EnterUserDefinedGhValues As Boolean?
    Private _BaseTowerGhInput As Double?
    Private _UpperStructureGhInput As Double?
    Private _EnterUserDefinedCgValues As Boolean?
    Private _BaseTowerCgInput As Double?
    Private _UpperStructureCgInput As Double?
    Private _AntennaFaceWidth As Double?
    Private _UseTopTakeup As Boolean?
    Private _ConstantSlope As Boolean?
    Private _upperStructure As New List(Of tnxAntennaRecord)
    Private _baseStructure As New List(Of tnxTowerRecord)
    Private _guyWires As New List(Of tnxGuyRecord)


    <Category("TNX Geometry"), Description("Base Tower Type"), DisplayName("TowerType")>
    <DataMember()> Public Property TowerType() As String
        Get
            Return Me._TowerType
        End Get
        Set
            Me._TowerType = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("Upper Structure Type"), DisplayName("AntennaType")>
    <DataMember()> Public Property AntennaType() As String
        Get
            Return Me._AntennaType
        End Get
        Set
            Me._AntennaType = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description(""), DisplayName("OverallHeight")>
    <DataMember()> Public Property OverallHeight() As Double?
        Get
            Return Me._OverallHeight
        End Get
        Set
            Me._OverallHeight = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description(""), DisplayName("BaseElevation")>
    <DataMember()> Public Property BaseElevation() As Double?
        Get
            Return Me._BaseElevation
        End Get
        Set
            Me._BaseElevation = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description(""), DisplayName("Lambda")>
    <DataMember()> Public Property Lambda() As Double?
        Get
            Return Me._Lambda
        End Get
        Set
            Me._Lambda = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description(""), DisplayName("TowerTopFaceWidth")>
    <DataMember()> Public Property TowerTopFaceWidth() As Double?
        Get
            Return Me._TowerTopFaceWidth
        End Get
        Set
            Me._TowerTopFaceWidth = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description(""), DisplayName("TowerBaseFaceWidth")>
    <DataMember()> Public Property TowerBaseFaceWidth() As Double?
        Get
            Return Me._TowerBaseFaceWidth
        End Get
        Set
            Me._TowerBaseFaceWidth = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("Base Type - None, I - Beam, I - Beam Free, Taper, Taper - Free"), DisplayName("TowerTaper")>
    <DataMember()> Public Property TowerTaper() As String
        Get
            Return Me._TowerTaper
        End Get
        Set
            Me._TowerTaper = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("Base Type - Fixed Base, Pinned Base (Only active When base tower type Is guyed And there are no base tower section In the model)"), DisplayName("GuyedMonopoleBaseType")>
    <DataMember()> Public Property GuyedMonopoleBaseType() As String
        Get
            Return Me._GuyedMonopoleBaseType
        End Get
        Set
            Me._GuyedMonopoleBaseType = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("Base Taper Height"), DisplayName("TaperHeight")>
    <DataMember()> Public Property TaperHeight() As Double?
        Get
            Return Me._TaperHeight
        End Get
        Set
            Me._TaperHeight = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("I-Beam Pivot Dist (replaces base taper height When base type Is I-Beam Or I-Beam Free)"), DisplayName("PivotHeight")>
    <DataMember()> Public Property PivotHeight() As Double?
        Get
            Return Me._PivotHeight
        End Get
        Set
            Me._PivotHeight = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description(""), DisplayName("AutoCalcGH")>
    <DataMember()> Public Property AutoCalcGH() As Boolean?
        Get
            Return Me._AutoCalcGH
        End Get
        Set
            Me._AutoCalcGH = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description(""), DisplayName("UserGHElev")>
    <DataMember()> Public Property UserGHElev() As Double?
        Get
            Return Me._UserGHElev
        End Get
        Set
            Me._UserGHElev = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("Has Index Plate"), DisplayName("UseIndexPlate")>
    <DataMember()> Public Property UseIndexPlate() As Boolean?
        Get
            Return Me._UseIndexPlate
        End Get
        Set
            Me._UseIndexPlate = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("Enter Pre-defined Gh values"), DisplayName("EnterUserDefinedGhValues")>
    <DataMember()> Public Property EnterUserDefinedGhValues() As Boolean?
        Get
            Return Me._EnterUserDefinedGhValues
        End Get
        Set
            Me._EnterUserDefinedGhValues = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("Base Tower - Active When EnterUserDefinedGhValues = Yes"), DisplayName("BaseTowerGhInput")>
    <DataMember()> Public Property BaseTowerGhInput() As Double?
        Get
            Return Me._BaseTowerGhInput
        End Get
        Set
            Me._BaseTowerGhInput = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("Upper Structure - Active When EnterUserDefinedGhValues = Yes"), DisplayName("UpperStructureGhInput")>
    <DataMember()> Public Property UpperStructureGhInput() As Double?
        Get
            Return Me._UpperStructureGhInput
        End Get
        Set
            Me._UpperStructureGhInput = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("CSA code only (This controls two inputs In the UI 'Use Default Cg Values' and 'Enter pre-defined Cg Values'. The checked status of the two inputs are always opposite and 'Use Default Cg Values' is opposite of the ERI value."), DisplayName("EnterUserDefinedCgValues")>
    <DataMember()> Public Property EnterUserDefinedCgValues() As Boolean?
        Get
            Return Me._EnterUserDefinedCgValues
        End Get
        Set
            Me._EnterUserDefinedCgValues = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("CSA code only"), DisplayName("BaseTowerCgInput")>
    <DataMember()> Public Property BaseTowerCgInput() As Double?
        Get
            Return Me._BaseTowerCgInput
        End Get
        Set
            Me._BaseTowerCgInput = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("CSA code only"), DisplayName("UpperStructureCgInput")>
    <DataMember()> Public Property UpperStructureCgInput() As Double?
        Get
            Return Me._UpperStructureCgInput
        End Get
        Set
            Me._UpperStructureCgInput = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("Lattice Pole Width - Only applies to lattice upper structures"), DisplayName("AntennaFaceWidth")>
    <DataMember()> Public Property AntennaFaceWidth() As Double?
        Get
            Return Me._AntennaFaceWidth
        End Get
        Set
            Me._AntennaFaceWidth = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("Top takeup on lambda"), DisplayName("UseTopTakeup")>
    <DataMember()> Public Property UseTopTakeup() As Boolean?
        Get
            Return Me._UseTopTakeup
        End Get
        Set
            Me._UseTopTakeup = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("Constant Slope"), DisplayName("ConstantSlope")>
    <DataMember()> Public Property ConstantSlope() As Boolean?
        Get
            Return Me._ConstantSlope
        End Get
        Set
            Me._ConstantSlope = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description(""), DisplayName("Upper Structure")>
    <DataMember()> Public Property upperStructure() As List(Of tnxAntennaRecord)
        Get
            Return Me._upperStructure
        End Get
        Set
            Me._upperStructure = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description(""), DisplayName("Base Structure")>
    <DataMember()> Public Property baseStructure() As List(Of tnxTowerRecord)
        Get
            Return Me._baseStructure
        End Get
        Set
            Me._baseStructure = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description(""), DisplayName("Guy Wires")>
    <DataMember()> Public Property guyWires() As List(Of tnxGuyRecord)
        Get
            Return Me._guyWires
        End Get
        Set
            Me._guyWires = Value
        End Set
    End Property

    <Category("Settings"), Description("Consider tower sections in the equality comparison."), DisplayName("Consider Geometry Equality")>
    <DataMember()> Public Property ConsiderSectionEquality() As Boolean = True

#End Region

#Region "Constructors"
    Public Sub New(Optional ByVal Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)
    End Sub

#End Region

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As tnxGeometry = TryCast(other, tnxGeometry)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.TowerType.CheckChange(otherToCompare.TowerType, changes, categoryName, "Tower Type"), Equals, False)
        Equals = If(Me.AntennaType.CheckChange(otherToCompare.AntennaType, changes, categoryName, "Antenna Type"), Equals, False)
        Equals = If(Me.OverallHeight.CheckChange(otherToCompare.OverallHeight, changes, categoryName, "Overall Height"), Equals, False)
        Equals = If(Me.BaseElevation.CheckChange(otherToCompare.BaseElevation, changes, categoryName, "Base Elevation"), Equals, False)
        Equals = If(Me.Lambda.CheckChange(otherToCompare.Lambda, changes, categoryName, "Lambda"), Equals, False)
        Equals = If(Me.TowerTopFaceWidth.CheckChange(otherToCompare.TowerTopFaceWidth, changes, categoryName, "Tower Top Face Width"), Equals, False)
        Equals = If(Me.TowerBaseFaceWidth.CheckChange(otherToCompare.TowerBaseFaceWidth, changes, categoryName, "Tower Base Face Width"), Equals, False)
        Equals = If(Me.TowerTaper.CheckChange(otherToCompare.TowerTaper, changes, categoryName, "Tower Taper"), Equals, False)
        Equals = If(Me.GuyedMonopoleBaseType.CheckChange(otherToCompare.GuyedMonopoleBaseType, changes, categoryName, "Guyed Monopole Base Type"), Equals, False)
        Equals = If(Me.TaperHeight.CheckChange(otherToCompare.TaperHeight, changes, categoryName, "Taper Height"), Equals, False)
        Equals = If(Me.PivotHeight.CheckChange(otherToCompare.PivotHeight, changes, categoryName, "Pivot Height"), Equals, False)
        Equals = If(Me.AutoCalcGH.CheckChange(otherToCompare.AutoCalcGH, changes, categoryName, "Auto Calc GH"), Equals, False)
        Equals = If(Me.UserGHElev.CheckChange(otherToCompare.UserGHElev, changes, categoryName, "User GH Elev"), Equals, False)
        Equals = If(Me.UseIndexPlate.CheckChange(otherToCompare.UseIndexPlate, changes, categoryName, "Use Index Plate"), Equals, False)
        Equals = If(Me.EnterUserDefinedGhValues.CheckChange(otherToCompare.EnterUserDefinedGhValues, changes, categoryName, "Enter User Defined Gh Values"), Equals, False)
        Equals = If(Me.BaseTowerGhInput.CheckChange(otherToCompare.BaseTowerGhInput, changes, categoryName, "Base Tower Gh Input"), Equals, False)
        Equals = If(Me.UpperStructureGhInput.CheckChange(otherToCompare.UpperStructureGhInput, changes, categoryName, "Upper Structure Gh Input"), Equals, False)
        Equals = If(Me.EnterUserDefinedCgValues.CheckChange(otherToCompare.EnterUserDefinedCgValues, changes, categoryName, "Enter User Defined Cg Values"), Equals, False)
        Equals = If(Me.BaseTowerCgInput.CheckChange(otherToCompare.BaseTowerCgInput, changes, categoryName, "Base Tower Cg Input"), Equals, False)
        Equals = If(Me.UpperStructureCgInput.CheckChange(otherToCompare.UpperStructureCgInput, changes, categoryName, "Upper Structure Cg Input"), Equals, False)
        Equals = If(Me.AntennaFaceWidth.CheckChange(otherToCompare.AntennaFaceWidth, changes, categoryName, "Antenna Face Width"), Equals, False)
        Equals = If(Me.UseTopTakeup.CheckChange(otherToCompare.UseTopTakeup, changes, categoryName, "Use Top Takeup"), Equals, False)
        Equals = If(Me.ConstantSlope.CheckChange(otherToCompare.ConstantSlope, changes, categoryName, "Constant Slope"), Equals, False)
        If Me.ConsiderSectionEquality Then
            Equals = If(Me.upperStructure.CheckChange(otherToCompare.upperStructure, changes, categoryName, "Upper Structure"), Equals, False)
            Equals = If(Me.baseStructure.CheckChange(otherToCompare.baseStructure, changes, categoryName, "Base Structure"), Equals, False)
            Equals = If(Me.guyWires.CheckChange(otherToCompare.guyWires, changes, categoryName, "Guy Wires"), Equals, False)
        End If

        Return Equals
    End Function

    ''' <summary>
    ''' Converts from XML report section numbering to upper/base geometry record numbers
    ''' </summary>
    ''' <param name="SectionNumber">1 as top section.</param>
    ''' <returns></returns>
    Public Function tnxSectionSelector(SectionNumber As Integer) As tnxGeometryRec
        If SectionNumber <= Me.upperStructure.Count Then
            Return Me.upperStructure(SectionNumber - 1)
        Else
            Return Me.baseStructure(SectionNumber - Me.upperStructure.Count - 1)
        End If
    End Function

End Class


Public Class tnxResult
    Inherits EDSResult

    <Category("Loads"), Description(""), DisplayName("Design Load")>
    <DataMember()> Public Property DesignLoad As Double?
    <Category("Loads"), Description(""), DisplayName("Applied Load")>
    <DataMember()> Public Property AppliedLoad As Double?
    <Category("Ratio"), Description(""), DisplayName("Load Ratio Limit")>
    <DataMember()> Public Property LoadRatioLimit As Double?
    '<Category("Ratio"), Description(""), DisplayName("Required Safety Factor")>
    ' <DataMember()> Public Property RequiredSafteyFactor As Double?
    '<Category("Ratio"), Description(""), DisplayName("Use Safety Factor Instead of Ratio")>
    ' <DataMember()> Public Property UseSFInsteadofRatio As Boolean = False

    <Category("Ratio"), Description("This rating takes into account TIA-222-H Annex S Section 15.5 when applicable."), DisplayName("Rating")>
    Public Overrides Property Rating As Double?
        Get
            Dim designCode As String
            Dim useAnnexS As Boolean
            Try
                designCode = Me.ParentStructure.tnx.code.design.DesignCode
                useAnnexS = Me.ParentStructure.tnx.code.design.UseTIA222H_AnnexS.Value
            Catch ex As Exception
                designCode = ""
                useAnnexS = False
                Debug.Print("Design code unknown. Using nonnormailzed TNX results.")
            End Try

            If designCode = "TIA-222-H" And useAnnexS Then
                Return Me.NormalizedRatio
            Else
                Return Me.Ratio
            End If
        End Get
        Set(value As Double?)
            'Do Nothing
        End Set
    End Property

    Public Sub New()
        'Leave Blank
    End Sub

    ''' <summary>
    ''' Create result object with result_lkup and rating
    ''' </summary>
    ''' <param name="result_lkup"></param>
    ''' <param name="rating"></param>
    ''' <param name="designLoad"></param>
    ''' <param name="appliedLoad"></param>
    ''' <param name="Parent"></param>
    Public Sub New(ByVal result_lkup As String, ByVal rating As Double?, ByVal designLoad As Double?, ByVal appliedLoad As Double?, ByVal RatioLimit As Double?, Optional ByVal Parent As EDSObjectWithQueries = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then
            Me.Absorb(Parent)
        End If

        Me.result_lkup = result_lkup
        Me.Rating = rating
        Me.DesignLoad = designLoad
        Me.AppliedLoad = appliedLoad
        Me.LoadRatioLimit = RatioLimit

    End Sub

    ''' <summary>
    ''' Ratio of the applied load to the design load.
    ''' </summary>
    ''' <returns></returns>
    Public Function Ratio() As Double
        If DesignLoad.HasValue And AppliedLoad.HasValue Then
            Return Math.Abs(AppliedLoad.Value / DesignLoad.Value)
        End If
        Return 0
    End Function

    ''' <summary>
    ''' Ratio of the applied load to the design load and normalized with the load ratio limit (i.e. 105%).
    ''' </summary>
    ''' <returns></returns>
    Public Function NormalizedRatio() As Double
        If DesignLoad.HasValue And AppliedLoad.HasValue And LoadRatioLimit.HasValue Then
            Return Math.Abs(AppliedLoad.Value / DesignLoad.Value) / LoadRatioLimit.Value
        End If
        Return 0
    End Function

End Class

<DataContractAttribute()>
Partial Public MustInherit Class tnxGeometryRec
    Inherits EDSObjectWithQueries
    Implements IComparable(Of tnxGeometryRec)
    'Use this type to allow sorting by Record instead of ID and provide a custom version of EDSListQueryBuilder for tnx geometry records
    <DataMember()> Public Property Rec As Integer?
    Public Overrides ReadOnly Property EDSTableDepth As Integer = 1
    'This cannot override the Results list from the EDSObjectWithQueries because it is a different type, instead we can shadow
    <DataMember()> Public Property TNXResults As New List(Of tnxResult)

    Private _Results As List(Of EDSResult)
    Public Overrides Property Results As List(Of EDSResult)
        Get
            Return Me.TNXResults.ConvertAll(Function(x) CType(x, EDSResult))
        End Get
        Set(value As List(Of EDSResult))
            Me._Results = value
        End Set
    End Property

    Public Overloads Function CompareTo(other As tnxGeometryRec) As Integer Implements IComparable(Of tnxGeometryRec).CompareTo
        'Sorted by Rec
        If other Is Nothing Then
            Return 1
        Else
            Return Nullable.Compare(Me.Rec, other.Rec)
        End If
    End Function

End Class

<DataContractAttribute()>
<KnownType(GetType(tnxAntennaRecord))>
Partial Public Class tnxAntennaRecord
    Inherits tnxGeometryRec
    'upper structure
#Region "Inheritted"

    Public Overrides ReadOnly Property EDSObjectName As String = "Upper Structure Section " & Me.Rec.ToString
    Public Overrides ReadOnly Property EDSTableName As String = "tnx.upper_structure_sections"

    Public Overrides Function SQLInsertValues() As String
        Return SQLInsertValues(Nothing)
    End Function
    'Public Overloads Function SQLInsertValues(Optional ByVal ParentIDKnown As Boolean = True) As String
    Public Overloads Function SQLInsertValues(Optional ByVal ParentID As Integer? = Nothing) As String
        'For any EDSObject that has parent object we will need to overload the update property with a version that excepts the current version being updated.
        SQLInsertValues = ""

        SQLInsertValues = SQLInsertValues.AddtoDBString(If(ParentID Is Nothing, EDSStructure.SQLQueryIDVar(Me.EDSTableDepth - 1), ParentID.NullableToString.FormatDBValue))
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.Rec.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaBraceType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaHeight.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaDiagonalSpacing.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaDiagonalSpacingEx.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaNumSections.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaNumSesctions.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaSectionLength.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaLegType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaLegSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaLegGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaLegMatlGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaDiagonalGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaDiagonalMatlGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaInnerBracingGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaInnerBracingMatlGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaTopGirtGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaTopGirtMatlGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaBotGirtGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaBotGirtMatlGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaInnerGirtGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaInnerGirtMatlGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaLongHorizontalGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaLongHorizontalMatlGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaShortHorizontalGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaShortHorizontalMatlGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaDiagonalType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaDiagonalSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaInnerBracingType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaInnerBracingSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaTopGirtType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaTopGirtSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaBotGirtType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaBotGirtSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaTopGirtOffset.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaBotGirtOffset.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaHasKBraceEndPanels.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaHasHorizontals.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaLongHorizontalType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaLongHorizontalSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaShortHorizontalType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaShortHorizontalSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantMatlGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantDiagType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantSubDiagonalType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantSubHorizontalType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantVerticalType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHipType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHipDiagonalType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHorizontalSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHorizontalSize2.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHorizontalSize3.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHorizontalSize4.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantDiagonalSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantDiagonalSize2.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantDiagonalSize3.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantDiagonalSize4.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantSubHorizontalSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantSubDiagonalSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaSubDiagLocation.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantVerticalSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHipDiagonalSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHipDiagonalSize2.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHipDiagonalSize3.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHipDiagonalSize4.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHipSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHipSize2.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHipSize3.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHipSize4.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaNumInnerGirts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaInnerGirtType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaInnerGirtSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaPoleShapeType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaPoleSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaPoleGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaPoleMatlGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaPoleSpliceLength.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaTaperPoleNumSides.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaTaperPoleTopDiameter.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaTaperPoleBotDiameter.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaTaperPoleWallThickness.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaTaperPoleBendRadius.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaTaperPoleGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaTaperPoleMatlGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaSWMult.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaWPMult.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaAutoCalcKSingleAngle.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaAutoCalcKSolidRound.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaAfGusset.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaTfGusset.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaGussetBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaGussetGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaGussetMatlGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaAfMult.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaArMult.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaFlatIPAPole.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRoundIPAPole.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaFlatIPALeg.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRoundIPALeg.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaFlatIPAHorizontal.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRoundIPAHorizontal.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaFlatIPADiagonal.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRoundIPADiagonal.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaCSA_S37_SpeedUpFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaKLegs.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaKXBracedDiags.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaKKBracedDiags.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaKZBracedDiags.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaKHorzs.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaKSecHorzs.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaKGirts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaKInners.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaKXBracedDiagsY.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaKKBracedDiagsY.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaKZBracedDiagsY.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaKHorzsY.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaKSecHorzsY.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaKGirtsY.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaKInnersY.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaKRedHorz.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaKRedDiag.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaKRedSubDiag.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaKRedSubHorz.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaKRedVert.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaKRedHip.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaKRedHipDiag.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaKTLX.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaKTLZ.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaKTLLeg.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaInnerKTLX.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaInnerKTLZ.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaInnerKTLLeg.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaStitchBoltLocationHoriz.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaStitchBoltLocationDiag.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaStitchSpacing.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaStitchSpacingHorz.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaStitchSpacingDiag.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaStitchSpacingRed.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaLegNetWidthDeduct.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaLegUFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaDiagonalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaTopGirtNetWidthDeduct.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaBotGirtNetWidthDeduct.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaInnerGirtNetWidthDeduct.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaHorizontalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaShortHorizontalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaDiagonalUFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaTopGirtUFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaBotGirtUFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaInnerGirtUFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaHorizontalUFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaShortHorizontalUFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaLegConnType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaLegNumBolts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaDiagonalNumBolts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaTopGirtNumBolts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaBotGirtNumBolts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaInnerGirtNumBolts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaHorizontalNumBolts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaShortHorizontalNumBolts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaLegBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaLegBoltSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaDiagonalBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaDiagonalBoltSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaTopGirtBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaTopGirtBoltSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaBotGirtBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaBotGirtBoltSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaInnerGirtBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaInnerGirtBoltSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaHorizontalBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaHorizontalBoltSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaShortHorizontalBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaShortHorizontalBoltSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaLegBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaDiagonalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaTopGirtBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaBotGirtBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaInnerGirtBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaHorizontalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaShortHorizontalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaDiagonalGageG1Distance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaTopGirtGageG1Distance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaBotGirtGageG1Distance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaInnerGirtGageG1Distance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaHorizontalGageG1Distance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaShortHorizontalGageG1Distance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHorizontalBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHorizontalBoltSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHorizontalNumBolts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHorizontalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHorizontalGageG1Distance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHorizontalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHorizontalUFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantDiagonalBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantDiagonalBoltSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantDiagonalNumBolts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantDiagonalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantDiagonalGageG1Distance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantDiagonalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantDiagonalUFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantSubDiagonalBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantSubDiagonalBoltSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantSubDiagonalNumBolts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantSubDiagonalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantSubDiagonalGageG1Distance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantSubDiagonalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantSubDiagonalUFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantSubHorizontalBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantSubHorizontalBoltSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantSubHorizontalNumBolts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantSubHorizontalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantSubHorizontalGageG1Distance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantSubHorizontalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantSubHorizontalUFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantVerticalBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantVerticalBoltSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantVerticalNumBolts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantVerticalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantVerticalGageG1Distance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantVerticalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantVerticalUFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHipBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHipBoltSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHipNumBolts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHipBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHipGageG1Distance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHipNetWidthDeduct.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHipUFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHipDiagonalBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHipDiagonalBoltSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHipDiagonalNumBolts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHipDiagonalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHipDiagonalGageG1Distance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHipDiagonalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaRedundantHipDiagonalUFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaDiagonalOutOfPlaneRestraint.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaTopGirtOutOfPlaneRestraint.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaBottomGirtOutOfPlaneRestraint.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaMidGirtOutOfPlaneRestraint.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaHorizontalOutOfPlaneRestraint.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaSecondaryHorizontalOutOfPlaneRestraint.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaDiagOffsetNEY.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaDiagOffsetNEX.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaDiagOffsetPEY.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaDiagOffsetPEX.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaKbraceOffsetNEY.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaKbraceOffsetNEX.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaKbraceOffsetPEY.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AntennaKbraceOffsetPEX.NullableToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        SQLInsertFields = SQLInsertFields.AddtoDBString("tnx_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRec")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaBraceType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaHeight")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaDiagonalSpacing")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaDiagonalSpacingEx")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaNumSections")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaNumSesctions")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaSectionLength")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaLegType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaLegSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaLegGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaLegMatlGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaDiagonalGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaDiagonalMatlGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaInnerBracingGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaInnerBracingMatlGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaTopGirtGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaTopGirtMatlGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaBotGirtGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaBotGirtMatlGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaInnerGirtGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaInnerGirtMatlGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaLongHorizontalGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaLongHorizontalMatlGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaShortHorizontalGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaShortHorizontalMatlGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaDiagonalType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaDiagonalSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaInnerBracingType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaInnerBracingSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaTopGirtType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaTopGirtSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaBotGirtType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaBotGirtSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaTopGirtOffset")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaBotGirtOffset")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaHasKBraceEndPanels")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaHasHorizontals")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaLongHorizontalType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaLongHorizontalSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaShortHorizontalType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaShortHorizontalSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantMatlGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantDiagType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantSubDiagonalType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantSubHorizontalType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantVerticalType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHipType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHipDiagonalType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHorizontalSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHorizontalSize2")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHorizontalSize3")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHorizontalSize4")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantDiagonalSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantDiagonalSize2")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantDiagonalSize3")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantDiagonalSize4")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantSubHorizontalSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantSubDiagonalSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaSubDiagLocation")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantVerticalSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHipDiagonalSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHipDiagonalSize2")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHipDiagonalSize3")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHipDiagonalSize4")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHipSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHipSize2")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHipSize3")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHipSize4")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaNumInnerGirts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaInnerGirtType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaInnerGirtSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaPoleShapeType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaPoleSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaPoleGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaPoleMatlGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaPoleSpliceLength")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaTaperPoleNumSides")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaTaperPoleTopDiameter")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaTaperPoleBotDiameter")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaTaperPoleWallThickness")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaTaperPoleBendRadius")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaTaperPoleGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaTaperPoleMatlGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaSWMult")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaWPMult")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaAutoCalcKSingleAngle")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaAutoCalcKSolidRound")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaAfGusset")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaTfGusset")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaGussetBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaGussetGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaGussetMatlGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaAfMult")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaArMult")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaFlatIPAPole")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRoundIPAPole")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaFlatIPALeg")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRoundIPALeg")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaFlatIPAHorizontal")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRoundIPAHorizontal")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaFlatIPADiagonal")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRoundIPADiagonal")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaCSA_S37_SpeedUpFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaKLegs")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaKXBracedDiags")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaKKBracedDiags")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaKZBracedDiags")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaKHorzs")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaKSecHorzs")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaKGirts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaKInners")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaKXBracedDiagsY")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaKKBracedDiagsY")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaKZBracedDiagsY")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaKHorzsY")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaKSecHorzsY")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaKGirtsY")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaKInnersY")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaKRedHorz")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaKRedDiag")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaKRedSubDiag")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaKRedSubHorz")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaKRedVert")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaKRedHip")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaKRedHipDiag")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaKTLX")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaKTLZ")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaKTLLeg")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaInnerKTLX")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaInnerKTLZ")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaInnerKTLLeg")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaStitchBoltLocationHoriz")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaStitchBoltLocationDiag")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaStitchSpacing")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaStitchSpacingHorz")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaStitchSpacingDiag")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaStitchSpacingRed")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaLegNetWidthDeduct")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaLegUFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaDiagonalNetWidthDeduct")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaTopGirtNetWidthDeduct")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaBotGirtNetWidthDeduct")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaInnerGirtNetWidthDeduct")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaHorizontalNetWidthDeduct")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaShortHorizontalNetWidthDeduct")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaDiagonalUFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaTopGirtUFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaBotGirtUFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaInnerGirtUFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaHorizontalUFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaShortHorizontalUFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaLegConnType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaLegNumBolts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaDiagonalNumBolts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaTopGirtNumBolts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaBotGirtNumBolts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaInnerGirtNumBolts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaHorizontalNumBolts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaShortHorizontalNumBolts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaLegBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaLegBoltSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaDiagonalBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaDiagonalBoltSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaTopGirtBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaTopGirtBoltSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaBotGirtBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaBotGirtBoltSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaInnerGirtBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaInnerGirtBoltSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaHorizontalBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaHorizontalBoltSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaShortHorizontalBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaShortHorizontalBoltSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaLegBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaDiagonalBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaTopGirtBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaBotGirtBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaInnerGirtBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaHorizontalBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaShortHorizontalBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaDiagonalGageG1Distance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaTopGirtGageG1Distance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaBotGirtGageG1Distance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaInnerGirtGageG1Distance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaHorizontalGageG1Distance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaShortHorizontalGageG1Distance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHorizontalBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHorizontalBoltSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHorizontalNumBolts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHorizontalBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHorizontalGageG1Distance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHorizontalNetWidthDeduct")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHorizontalUFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantDiagonalBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantDiagonalBoltSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantDiagonalNumBolts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantDiagonalBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantDiagonalGageG1Distance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantDiagonalNetWidthDeduct")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantDiagonalUFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantSubDiagonalBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantSubDiagonalBoltSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantSubDiagonalNumBolts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantSubDiagonalBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantSubDiagonalGageG1Distance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantSubDiagonalNetWidthDeduct")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantSubDiagonalUFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantSubHorizontalBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantSubHorizontalBoltSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantSubHorizontalNumBolts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantSubHorizontalBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantSubHorizontalGageG1Distance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantSubHorizontalNetWidthDeduct")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantSubHorizontalUFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantVerticalBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantVerticalBoltSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantVerticalNumBolts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantVerticalBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantVerticalGageG1Distance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantVerticalNetWidthDeduct")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantVerticalUFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHipBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHipBoltSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHipNumBolts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHipBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHipGageG1Distance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHipNetWidthDeduct")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHipUFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHipDiagonalBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHipDiagonalBoltSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHipDiagonalNumBolts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHipDiagonalBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHipDiagonalGageG1Distance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHipDiagonalNetWidthDeduct")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaRedundantHipDiagonalUFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaDiagonalOutOfPlaneRestraint")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaTopGirtOutOfPlaneRestraint")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaBottomGirtOutOfPlaneRestraint")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaMidGirtOutOfPlaneRestraint")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaHorizontalOutOfPlaneRestraint")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaSecondaryHorizontalOutOfPlaneRestraint")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaDiagOffsetNEY")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaDiagOffsetNEX")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaDiagOffsetPEY")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaDiagOffsetPEX")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaKbraceOffsetNEY")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaKbraceOffsetNEX")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaKbraceOffsetPEY")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaKbraceOffsetPEX")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""

        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRec = " & Me.Rec.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaBraceType = " & Me.AntennaBraceType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaHeight = " & Me.AntennaHeight.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaDiagonalSpacing = " & Me.AntennaDiagonalSpacing.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaDiagonalSpacingEx = " & Me.AntennaDiagonalSpacingEx.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaNumSections = " & Me.AntennaNumSections.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaNumSesctions = " & Me.AntennaNumSesctions.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaSectionLength = " & Me.AntennaSectionLength.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaLegType = " & Me.AntennaLegType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaLegSize = " & Me.AntennaLegSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaLegGrade = " & Me.AntennaLegGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaLegMatlGrade = " & Me.AntennaLegMatlGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaDiagonalGrade = " & Me.AntennaDiagonalGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaDiagonalMatlGrade = " & Me.AntennaDiagonalMatlGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaInnerBracingGrade = " & Me.AntennaInnerBracingGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaInnerBracingMatlGrade = " & Me.AntennaInnerBracingMatlGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaTopGirtGrade = " & Me.AntennaTopGirtGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaTopGirtMatlGrade = " & Me.AntennaTopGirtMatlGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaBotGirtGrade = " & Me.AntennaBotGirtGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaBotGirtMatlGrade = " & Me.AntennaBotGirtMatlGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaInnerGirtGrade = " & Me.AntennaInnerGirtGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaInnerGirtMatlGrade = " & Me.AntennaInnerGirtMatlGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaLongHorizontalGrade = " & Me.AntennaLongHorizontalGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaLongHorizontalMatlGrade = " & Me.AntennaLongHorizontalMatlGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaShortHorizontalGrade = " & Me.AntennaShortHorizontalGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaShortHorizontalMatlGrade = " & Me.AntennaShortHorizontalMatlGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaDiagonalType = " & Me.AntennaDiagonalType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaDiagonalSize = " & Me.AntennaDiagonalSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaInnerBracingType = " & Me.AntennaInnerBracingType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaInnerBracingSize = " & Me.AntennaInnerBracingSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaTopGirtType = " & Me.AntennaTopGirtType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaTopGirtSize = " & Me.AntennaTopGirtSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaBotGirtType = " & Me.AntennaBotGirtType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaBotGirtSize = " & Me.AntennaBotGirtSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaTopGirtOffset = " & Me.AntennaTopGirtOffset.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaBotGirtOffset = " & Me.AntennaBotGirtOffset.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaHasKBraceEndPanels = " & Me.AntennaHasKBraceEndPanels.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaHasHorizontals = " & Me.AntennaHasHorizontals.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaLongHorizontalType = " & Me.AntennaLongHorizontalType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaLongHorizontalSize = " & Me.AntennaLongHorizontalSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaShortHorizontalType = " & Me.AntennaShortHorizontalType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaShortHorizontalSize = " & Me.AntennaShortHorizontalSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantGrade = " & Me.AntennaRedundantGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantMatlGrade = " & Me.AntennaRedundantMatlGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantType = " & Me.AntennaRedundantType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantDiagType = " & Me.AntennaRedundantDiagType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantSubDiagonalType = " & Me.AntennaRedundantSubDiagonalType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantSubHorizontalType = " & Me.AntennaRedundantSubHorizontalType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantVerticalType = " & Me.AntennaRedundantVerticalType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHipType = " & Me.AntennaRedundantHipType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHipDiagonalType = " & Me.AntennaRedundantHipDiagonalType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHorizontalSize = " & Me.AntennaRedundantHorizontalSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHorizontalSize2 = " & Me.AntennaRedundantHorizontalSize2.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHorizontalSize3 = " & Me.AntennaRedundantHorizontalSize3.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHorizontalSize4 = " & Me.AntennaRedundantHorizontalSize4.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantDiagonalSize = " & Me.AntennaRedundantDiagonalSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantDiagonalSize2 = " & Me.AntennaRedundantDiagonalSize2.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantDiagonalSize3 = " & Me.AntennaRedundantDiagonalSize3.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantDiagonalSize4 = " & Me.AntennaRedundantDiagonalSize4.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantSubHorizontalSize = " & Me.AntennaRedundantSubHorizontalSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantSubDiagonalSize = " & Me.AntennaRedundantSubDiagonalSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaSubDiagLocation = " & Me.AntennaSubDiagLocation.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantVerticalSize = " & Me.AntennaRedundantVerticalSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHipDiagonalSize = " & Me.AntennaRedundantHipDiagonalSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHipDiagonalSize2 = " & Me.AntennaRedundantHipDiagonalSize2.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHipDiagonalSize3 = " & Me.AntennaRedundantHipDiagonalSize3.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHipDiagonalSize4 = " & Me.AntennaRedundantHipDiagonalSize4.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHipSize = " & Me.AntennaRedundantHipSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHipSize2 = " & Me.AntennaRedundantHipSize2.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHipSize3 = " & Me.AntennaRedundantHipSize3.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHipSize4 = " & Me.AntennaRedundantHipSize4.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaNumInnerGirts = " & Me.AntennaNumInnerGirts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaInnerGirtType = " & Me.AntennaInnerGirtType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaInnerGirtSize = " & Me.AntennaInnerGirtSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaPoleShapeType = " & Me.AntennaPoleShapeType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaPoleSize = " & Me.AntennaPoleSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaPoleGrade = " & Me.AntennaPoleGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaPoleMatlGrade = " & Me.AntennaPoleMatlGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaPoleSpliceLength = " & Me.AntennaPoleSpliceLength.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaTaperPoleNumSides = " & Me.AntennaTaperPoleNumSides.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaTaperPoleTopDiameter = " & Me.AntennaTaperPoleTopDiameter.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaTaperPoleBotDiameter = " & Me.AntennaTaperPoleBotDiameter.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaTaperPoleWallThickness = " & Me.AntennaTaperPoleWallThickness.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaTaperPoleBendRadius = " & Me.AntennaTaperPoleBendRadius.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaTaperPoleGrade = " & Me.AntennaTaperPoleGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaTaperPoleMatlGrade = " & Me.AntennaTaperPoleMatlGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaSWMult = " & Me.AntennaSWMult.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaWPMult = " & Me.AntennaWPMult.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaAutoCalcKSingleAngle = " & Me.AntennaAutoCalcKSingleAngle.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaAutoCalcKSolidRound = " & Me.AntennaAutoCalcKSolidRound.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaAfGusset = " & Me.AntennaAfGusset.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaTfGusset = " & Me.AntennaTfGusset.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaGussetBoltEdgeDistance = " & Me.AntennaGussetBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaGussetGrade = " & Me.AntennaGussetGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaGussetMatlGrade = " & Me.AntennaGussetMatlGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaAfMult = " & Me.AntennaAfMult.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaArMult = " & Me.AntennaArMult.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaFlatIPAPole = " & Me.AntennaFlatIPAPole.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRoundIPAPole = " & Me.AntennaRoundIPAPole.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaFlatIPALeg = " & Me.AntennaFlatIPALeg.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRoundIPALeg = " & Me.AntennaRoundIPALeg.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaFlatIPAHorizontal = " & Me.AntennaFlatIPAHorizontal.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRoundIPAHorizontal = " & Me.AntennaRoundIPAHorizontal.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaFlatIPADiagonal = " & Me.AntennaFlatIPADiagonal.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRoundIPADiagonal = " & Me.AntennaRoundIPADiagonal.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaCSA_S37_SpeedUpFactor = " & Me.AntennaCSA_S37_SpeedUpFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaKLegs = " & Me.AntennaKLegs.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaKXBracedDiags = " & Me.AntennaKXBracedDiags.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaKKBracedDiags = " & Me.AntennaKKBracedDiags.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaKZBracedDiags = " & Me.AntennaKZBracedDiags.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaKHorzs = " & Me.AntennaKHorzs.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaKSecHorzs = " & Me.AntennaKSecHorzs.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaKGirts = " & Me.AntennaKGirts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaKInners = " & Me.AntennaKInners.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaKXBracedDiagsY = " & Me.AntennaKXBracedDiagsY.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaKKBracedDiagsY = " & Me.AntennaKKBracedDiagsY.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaKZBracedDiagsY = " & Me.AntennaKZBracedDiagsY.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaKHorzsY = " & Me.AntennaKHorzsY.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaKSecHorzsY = " & Me.AntennaKSecHorzsY.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaKGirtsY = " & Me.AntennaKGirtsY.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaKInnersY = " & Me.AntennaKInnersY.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaKRedHorz = " & Me.AntennaKRedHorz.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaKRedDiag = " & Me.AntennaKRedDiag.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaKRedSubDiag = " & Me.AntennaKRedSubDiag.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaKRedSubHorz = " & Me.AntennaKRedSubHorz.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaKRedVert = " & Me.AntennaKRedVert.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaKRedHip = " & Me.AntennaKRedHip.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaKRedHipDiag = " & Me.AntennaKRedHipDiag.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaKTLX = " & Me.AntennaKTLX.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaKTLZ = " & Me.AntennaKTLZ.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaKTLLeg = " & Me.AntennaKTLLeg.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaInnerKTLX = " & Me.AntennaInnerKTLX.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaInnerKTLZ = " & Me.AntennaInnerKTLZ.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaInnerKTLLeg = " & Me.AntennaInnerKTLLeg.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaStitchBoltLocationHoriz = " & Me.AntennaStitchBoltLocationHoriz.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaStitchBoltLocationDiag = " & Me.AntennaStitchBoltLocationDiag.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaStitchSpacing = " & Me.AntennaStitchSpacing.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaStitchSpacingHorz = " & Me.AntennaStitchSpacingHorz.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaStitchSpacingDiag = " & Me.AntennaStitchSpacingDiag.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaStitchSpacingRed = " & Me.AntennaStitchSpacingRed.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaLegNetWidthDeduct = " & Me.AntennaLegNetWidthDeduct.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaLegUFactor = " & Me.AntennaLegUFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaDiagonalNetWidthDeduct = " & Me.AntennaDiagonalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaTopGirtNetWidthDeduct = " & Me.AntennaTopGirtNetWidthDeduct.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaBotGirtNetWidthDeduct = " & Me.AntennaBotGirtNetWidthDeduct.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaInnerGirtNetWidthDeduct = " & Me.AntennaInnerGirtNetWidthDeduct.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaHorizontalNetWidthDeduct = " & Me.AntennaHorizontalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaShortHorizontalNetWidthDeduct = " & Me.AntennaShortHorizontalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaDiagonalUFactor = " & Me.AntennaDiagonalUFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaTopGirtUFactor = " & Me.AntennaTopGirtUFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaBotGirtUFactor = " & Me.AntennaBotGirtUFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaInnerGirtUFactor = " & Me.AntennaInnerGirtUFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaHorizontalUFactor = " & Me.AntennaHorizontalUFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaShortHorizontalUFactor = " & Me.AntennaShortHorizontalUFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaLegConnType = " & Me.AntennaLegConnType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaLegNumBolts = " & Me.AntennaLegNumBolts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaDiagonalNumBolts = " & Me.AntennaDiagonalNumBolts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaTopGirtNumBolts = " & Me.AntennaTopGirtNumBolts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaBotGirtNumBolts = " & Me.AntennaBotGirtNumBolts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaInnerGirtNumBolts = " & Me.AntennaInnerGirtNumBolts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaHorizontalNumBolts = " & Me.AntennaHorizontalNumBolts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaShortHorizontalNumBolts = " & Me.AntennaShortHorizontalNumBolts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaLegBoltGrade = " & Me.AntennaLegBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaLegBoltSize = " & Me.AntennaLegBoltSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaDiagonalBoltGrade = " & Me.AntennaDiagonalBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaDiagonalBoltSize = " & Me.AntennaDiagonalBoltSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaTopGirtBoltGrade = " & Me.AntennaTopGirtBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaTopGirtBoltSize = " & Me.AntennaTopGirtBoltSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaBotGirtBoltGrade = " & Me.AntennaBotGirtBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaBotGirtBoltSize = " & Me.AntennaBotGirtBoltSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaInnerGirtBoltGrade = " & Me.AntennaInnerGirtBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaInnerGirtBoltSize = " & Me.AntennaInnerGirtBoltSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaHorizontalBoltGrade = " & Me.AntennaHorizontalBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaHorizontalBoltSize = " & Me.AntennaHorizontalBoltSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaShortHorizontalBoltGrade = " & Me.AntennaShortHorizontalBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaShortHorizontalBoltSize = " & Me.AntennaShortHorizontalBoltSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaLegBoltEdgeDistance = " & Me.AntennaLegBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaDiagonalBoltEdgeDistance = " & Me.AntennaDiagonalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaTopGirtBoltEdgeDistance = " & Me.AntennaTopGirtBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaBotGirtBoltEdgeDistance = " & Me.AntennaBotGirtBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaInnerGirtBoltEdgeDistance = " & Me.AntennaInnerGirtBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaHorizontalBoltEdgeDistance = " & Me.AntennaHorizontalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaShortHorizontalBoltEdgeDistance = " & Me.AntennaShortHorizontalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaDiagonalGageG1Distance = " & Me.AntennaDiagonalGageG1Distance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaTopGirtGageG1Distance = " & Me.AntennaTopGirtGageG1Distance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaBotGirtGageG1Distance = " & Me.AntennaBotGirtGageG1Distance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaInnerGirtGageG1Distance = " & Me.AntennaInnerGirtGageG1Distance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaHorizontalGageG1Distance = " & Me.AntennaHorizontalGageG1Distance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaShortHorizontalGageG1Distance = " & Me.AntennaShortHorizontalGageG1Distance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHorizontalBoltGrade = " & Me.AntennaRedundantHorizontalBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHorizontalBoltSize = " & Me.AntennaRedundantHorizontalBoltSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHorizontalNumBolts = " & Me.AntennaRedundantHorizontalNumBolts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHorizontalBoltEdgeDistance = " & Me.AntennaRedundantHorizontalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHorizontalGageG1Distance = " & Me.AntennaRedundantHorizontalGageG1Distance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHorizontalNetWidthDeduct = " & Me.AntennaRedundantHorizontalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHorizontalUFactor = " & Me.AntennaRedundantHorizontalUFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantDiagonalBoltGrade = " & Me.AntennaRedundantDiagonalBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantDiagonalBoltSize = " & Me.AntennaRedundantDiagonalBoltSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantDiagonalNumBolts = " & Me.AntennaRedundantDiagonalNumBolts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantDiagonalBoltEdgeDistance = " & Me.AntennaRedundantDiagonalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantDiagonalGageG1Distance = " & Me.AntennaRedundantDiagonalGageG1Distance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantDiagonalNetWidthDeduct = " & Me.AntennaRedundantDiagonalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantDiagonalUFactor = " & Me.AntennaRedundantDiagonalUFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantSubDiagonalBoltGrade = " & Me.AntennaRedundantSubDiagonalBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantSubDiagonalBoltSize = " & Me.AntennaRedundantSubDiagonalBoltSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantSubDiagonalNumBolts = " & Me.AntennaRedundantSubDiagonalNumBolts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantSubDiagonalBoltEdgeDistance = " & Me.AntennaRedundantSubDiagonalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantSubDiagonalGageG1Distance = " & Me.AntennaRedundantSubDiagonalGageG1Distance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantSubDiagonalNetWidthDeduct = " & Me.AntennaRedundantSubDiagonalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantSubDiagonalUFactor = " & Me.AntennaRedundantSubDiagonalUFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantSubHorizontalBoltGrade = " & Me.AntennaRedundantSubHorizontalBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantSubHorizontalBoltSize = " & Me.AntennaRedundantSubHorizontalBoltSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantSubHorizontalNumBolts = " & Me.AntennaRedundantSubHorizontalNumBolts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantSubHorizontalBoltEdgeDistance = " & Me.AntennaRedundantSubHorizontalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantSubHorizontalGageG1Distance = " & Me.AntennaRedundantSubHorizontalGageG1Distance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantSubHorizontalNetWidthDeduct = " & Me.AntennaRedundantSubHorizontalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantSubHorizontalUFactor = " & Me.AntennaRedundantSubHorizontalUFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantVerticalBoltGrade = " & Me.AntennaRedundantVerticalBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantVerticalBoltSize = " & Me.AntennaRedundantVerticalBoltSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantVerticalNumBolts = " & Me.AntennaRedundantVerticalNumBolts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantVerticalBoltEdgeDistance = " & Me.AntennaRedundantVerticalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantVerticalGageG1Distance = " & Me.AntennaRedundantVerticalGageG1Distance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantVerticalNetWidthDeduct = " & Me.AntennaRedundantVerticalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantVerticalUFactor = " & Me.AntennaRedundantVerticalUFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHipBoltGrade = " & Me.AntennaRedundantHipBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHipBoltSize = " & Me.AntennaRedundantHipBoltSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHipNumBolts = " & Me.AntennaRedundantHipNumBolts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHipBoltEdgeDistance = " & Me.AntennaRedundantHipBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHipGageG1Distance = " & Me.AntennaRedundantHipGageG1Distance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHipNetWidthDeduct = " & Me.AntennaRedundantHipNetWidthDeduct.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHipUFactor = " & Me.AntennaRedundantHipUFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHipDiagonalBoltGrade = " & Me.AntennaRedundantHipDiagonalBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHipDiagonalBoltSize = " & Me.AntennaRedundantHipDiagonalBoltSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHipDiagonalNumBolts = " & Me.AntennaRedundantHipDiagonalNumBolts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHipDiagonalBoltEdgeDistance = " & Me.AntennaRedundantHipDiagonalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHipDiagonalGageG1Distance = " & Me.AntennaRedundantHipDiagonalGageG1Distance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHipDiagonalNetWidthDeduct = " & Me.AntennaRedundantHipDiagonalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaRedundantHipDiagonalUFactor = " & Me.AntennaRedundantHipDiagonalUFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaDiagonalOutOfPlaneRestraint = " & Me.AntennaDiagonalOutOfPlaneRestraint.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaTopGirtOutOfPlaneRestraint = " & Me.AntennaTopGirtOutOfPlaneRestraint.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaBottomGirtOutOfPlaneRestraint = " & Me.AntennaBottomGirtOutOfPlaneRestraint.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaMidGirtOutOfPlaneRestraint = " & Me.AntennaMidGirtOutOfPlaneRestraint.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaHorizontalOutOfPlaneRestraint = " & Me.AntennaHorizontalOutOfPlaneRestraint.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaSecondaryHorizontalOutOfPlaneRestraint = " & Me.AntennaSecondaryHorizontalOutOfPlaneRestraint.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaDiagOffsetNEY = " & Me.AntennaDiagOffsetNEY.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaDiagOffsetNEX = " & Me.AntennaDiagOffsetNEX.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaDiagOffsetPEY = " & Me.AntennaDiagOffsetPEY.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaDiagOffsetPEX = " & Me.AntennaDiagOffsetPEX.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaKbraceOffsetNEY = " & Me.AntennaKbraceOffsetNEY.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaKbraceOffsetNEX = " & Me.AntennaKbraceOffsetNEX.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaKbraceOffsetPEY = " & Me.AntennaKbraceOffsetPEY.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaKbraceOffsetPEX = " & Me.AntennaKbraceOffsetPEX.NullableToString.FormatDBValue)

        Return SQLUpdateFieldsandValues
    End Function
#End Region

#Region "Define"
    'Private _AntennaRec As Integer?
    Private _AntennaBraceType As String
    Private _AntennaHeight As Double?
    Private _AntennaDiagonalSpacing As Double?
    Private _AntennaDiagonalSpacingEx As Double?
    Private _AntennaNumSections As Integer?
    Private _AntennaNumSesctions As Integer?
    Private _AntennaSectionLength As Double?
    Private _AntennaLegType As String
    Private _AntennaLegSize As String
    Private _AntennaLegGrade As Double?
    Private _AntennaLegMatlGrade As String
    Private _AntennaDiagonalGrade As Double?
    Private _AntennaDiagonalMatlGrade As String
    Private _AntennaInnerBracingGrade As Double?
    Private _AntennaInnerBracingMatlGrade As String
    Private _AntennaTopGirtGrade As Double?
    Private _AntennaTopGirtMatlGrade As String
    Private _AntennaBotGirtGrade As Double?
    Private _AntennaBotGirtMatlGrade As String
    Private _AntennaInnerGirtGrade As Double?
    Private _AntennaInnerGirtMatlGrade As String
    Private _AntennaLongHorizontalGrade As Double?
    Private _AntennaLongHorizontalMatlGrade As String
    Private _AntennaShortHorizontalGrade As Double?
    Private _AntennaShortHorizontalMatlGrade As String
    Private _AntennaDiagonalType As String
    Private _AntennaDiagonalSize As String
    Private _AntennaInnerBracingType As String
    Private _AntennaInnerBracingSize As String
    Private _AntennaTopGirtType As String
    Private _AntennaTopGirtSize As String
    Private _AntennaBotGirtType As String
    Private _AntennaBotGirtSize As String
    Private _AntennaTopGirtOffset As Double?
    Private _AntennaBotGirtOffset As Double?
    Private _AntennaHasKBraceEndPanels As Boolean?
    Private _AntennaHasHorizontals As Boolean?
    Private _AntennaLongHorizontalType As String
    Private _AntennaLongHorizontalSize As String
    Private _AntennaShortHorizontalType As String
    Private _AntennaShortHorizontalSize As String
    Private _AntennaRedundantGrade As Double?
    Private _AntennaRedundantMatlGrade As String
    Private _AntennaRedundantType As String
    Private _AntennaRedundantDiagType As String
    Private _AntennaRedundantSubDiagonalType As String
    Private _AntennaRedundantSubHorizontalType As String
    Private _AntennaRedundantVerticalType As String
    Private _AntennaRedundantHipType As String
    Private _AntennaRedundantHipDiagonalType As String
    Private _AntennaRedundantHorizontalSize As String
    Private _AntennaRedundantHorizontalSize2 As String
    Private _AntennaRedundantHorizontalSize3 As String
    Private _AntennaRedundantHorizontalSize4 As String
    Private _AntennaRedundantDiagonalSize As String
    Private _AntennaRedundantDiagonalSize2 As String
    Private _AntennaRedundantDiagonalSize3 As String
    Private _AntennaRedundantDiagonalSize4 As String
    Private _AntennaRedundantSubHorizontalSize As String
    Private _AntennaRedundantSubDiagonalSize As String
    Private _AntennaSubDiagLocation As Double?
    Private _AntennaRedundantVerticalSize As String
    Private _AntennaRedundantHipDiagonalSize As String
    Private _AntennaRedundantHipDiagonalSize2 As String
    Private _AntennaRedundantHipDiagonalSize3 As String
    Private _AntennaRedundantHipDiagonalSize4 As String
    Private _AntennaRedundantHipSize As String
    Private _AntennaRedundantHipSize2 As String
    Private _AntennaRedundantHipSize3 As String
    Private _AntennaRedundantHipSize4 As String
    Private _AntennaNumInnerGirts As Integer?
    Private _AntennaInnerGirtType As String
    Private _AntennaInnerGirtSize As String
    Private _AntennaPoleShapeType As String
    Private _AntennaPoleSize As String
    Private _AntennaPoleGrade As Double?
    Private _AntennaPoleMatlGrade As String
    Private _AntennaPoleSpliceLength As Double?
    Private _AntennaTaperPoleNumSides As Integer?
    Private _AntennaTaperPoleTopDiameter As Double?
    Private _AntennaTaperPoleBotDiameter As Double?
    Private _AntennaTaperPoleWallThickness As Double?
    Private _AntennaTaperPoleBendRadius As Double?
    Private _AntennaTaperPoleGrade As Double?
    Private _AntennaTaperPoleMatlGrade As String
    Private _AntennaSWMult As Double?
    Private _AntennaWPMult As Double?
    Private _AntennaAutoCalcKSingleAngle As Double?
    Private _AntennaAutoCalcKSolidRound As Double?
    Private _AntennaAfGusset As Double?
    Private _AntennaTfGusset As Double?
    Private _AntennaGussetBoltEdgeDistance As Double?
    Private _AntennaGussetGrade As Double?
    Private _AntennaGussetMatlGrade As String
    Private _AntennaAfMult As Double?
    Private _AntennaArMult As Double?
    Private _AntennaFlatIPAPole As Double?
    Private _AntennaRoundIPAPole As Double?
    Private _AntennaFlatIPALeg As Double?
    Private _AntennaRoundIPALeg As Double?
    Private _AntennaFlatIPAHorizontal As Double?
    Private _AntennaRoundIPAHorizontal As Double?
    Private _AntennaFlatIPADiagonal As Double?
    Private _AntennaRoundIPADiagonal As Double?
    Private _AntennaCSA_S37_SpeedUpFactor As Double?
    Private _AntennaKLegs As Double?
    Private _AntennaKXBracedDiags As Double?
    Private _AntennaKKBracedDiags As Double?
    Private _AntennaKZBracedDiags As Double?
    Private _AntennaKHorzs As Double?
    Private _AntennaKSecHorzs As Double?
    Private _AntennaKGirts As Double?
    Private _AntennaKInners As Double?
    Private _AntennaKXBracedDiagsY As Double?
    Private _AntennaKKBracedDiagsY As Double?
    Private _AntennaKZBracedDiagsY As Double?
    Private _AntennaKHorzsY As Double?
    Private _AntennaKSecHorzsY As Double?
    Private _AntennaKGirtsY As Double?
    Private _AntennaKInnersY As Double?
    Private _AntennaKRedHorz As Double?
    Private _AntennaKRedDiag As Double?
    Private _AntennaKRedSubDiag As Double?
    Private _AntennaKRedSubHorz As Double?
    Private _AntennaKRedVert As Double?
    Private _AntennaKRedHip As Double?
    Private _AntennaKRedHipDiag As Double?
    Private _AntennaKTLX As Double?
    Private _AntennaKTLZ As Double?
    Private _AntennaKTLLeg As Double?
    Private _AntennaInnerKTLX As Double?
    Private _AntennaInnerKTLZ As Double?
    Private _AntennaInnerKTLLeg As Double?
    Private _AntennaStitchBoltLocationHoriz As String
    Private _AntennaStitchBoltLocationDiag As String
    Private _AntennaStitchSpacing As Double?
    Private _AntennaStitchSpacingHorz As Double?
    Private _AntennaStitchSpacingDiag As Double?
    Private _AntennaStitchSpacingRed As Double?
    Private _AntennaLegNetWidthDeduct As Double?
    Private _AntennaLegUFactor As Double?
    Private _AntennaDiagonalNetWidthDeduct As Double?
    Private _AntennaTopGirtNetWidthDeduct As Double?
    Private _AntennaBotGirtNetWidthDeduct As Double?
    Private _AntennaInnerGirtNetWidthDeduct As Double?
    Private _AntennaHorizontalNetWidthDeduct As Double?
    Private _AntennaShortHorizontalNetWidthDeduct As Double?
    Private _AntennaDiagonalUFactor As Double?
    Private _AntennaTopGirtUFactor As Double?
    Private _AntennaBotGirtUFactor As Double?
    Private _AntennaInnerGirtUFactor As Double?
    Private _AntennaHorizontalUFactor As Double?
    Private _AntennaShortHorizontalUFactor As Double?
    Private _AntennaLegConnType As String
    Private _AntennaLegNumBolts As Integer?
    Private _AntennaDiagonalNumBolts As Integer?
    Private _AntennaTopGirtNumBolts As Integer?
    Private _AntennaBotGirtNumBolts As Integer?
    Private _AntennaInnerGirtNumBolts As Integer?
    Private _AntennaHorizontalNumBolts As Integer?
    Private _AntennaShortHorizontalNumBolts As Integer?
    Private _AntennaLegBoltGrade As String
    Private _AntennaLegBoltSize As Double?
    Private _AntennaDiagonalBoltGrade As String
    Private _AntennaDiagonalBoltSize As Double?
    Private _AntennaTopGirtBoltGrade As String
    Private _AntennaTopGirtBoltSize As Double?
    Private _AntennaBotGirtBoltGrade As String
    Private _AntennaBotGirtBoltSize As Double?
    Private _AntennaInnerGirtBoltGrade As String
    Private _AntennaInnerGirtBoltSize As Double?
    Private _AntennaHorizontalBoltGrade As String
    Private _AntennaHorizontalBoltSize As Double?
    Private _AntennaShortHorizontalBoltGrade As String
    Private _AntennaShortHorizontalBoltSize As Double?
    Private _AntennaLegBoltEdgeDistance As Double?
    Private _AntennaDiagonalBoltEdgeDistance As Double?
    Private _AntennaTopGirtBoltEdgeDistance As Double?
    Private _AntennaBotGirtBoltEdgeDistance As Double?
    Private _AntennaInnerGirtBoltEdgeDistance As Double?
    Private _AntennaHorizontalBoltEdgeDistance As Double?
    Private _AntennaShortHorizontalBoltEdgeDistance As Double?
    Private _AntennaDiagonalGageG1Distance As Double?
    Private _AntennaTopGirtGageG1Distance As Double?
    Private _AntennaBotGirtGageG1Distance As Double?
    Private _AntennaInnerGirtGageG1Distance As Double?
    Private _AntennaHorizontalGageG1Distance As Double?
    Private _AntennaShortHorizontalGageG1Distance As Double?
    Private _AntennaRedundantHorizontalBoltGrade As String
    Private _AntennaRedundantHorizontalBoltSize As Double?
    Private _AntennaRedundantHorizontalNumBolts As Integer?
    Private _AntennaRedundantHorizontalBoltEdgeDistance As Double?
    Private _AntennaRedundantHorizontalGageG1Distance As Double?
    Private _AntennaRedundantHorizontalNetWidthDeduct As Double?
    Private _AntennaRedundantHorizontalUFactor As Double?
    Private _AntennaRedundantDiagonalBoltGrade As String
    Private _AntennaRedundantDiagonalBoltSize As Double?
    Private _AntennaRedundantDiagonalNumBolts As Integer?
    Private _AntennaRedundantDiagonalBoltEdgeDistance As Double?
    Private _AntennaRedundantDiagonalGageG1Distance As Double?
    Private _AntennaRedundantDiagonalNetWidthDeduct As Double?
    Private _AntennaRedundantDiagonalUFactor As Double?
    Private _AntennaRedundantSubDiagonalBoltGrade As String
    Private _AntennaRedundantSubDiagonalBoltSize As Double?
    Private _AntennaRedundantSubDiagonalNumBolts As Integer?
    Private _AntennaRedundantSubDiagonalBoltEdgeDistance As Double?
    Private _AntennaRedundantSubDiagonalGageG1Distance As Double?
    Private _AntennaRedundantSubDiagonalNetWidthDeduct As Double?
    Private _AntennaRedundantSubDiagonalUFactor As Double?
    Private _AntennaRedundantSubHorizontalBoltGrade As String
    Private _AntennaRedundantSubHorizontalBoltSize As Double?
    Private _AntennaRedundantSubHorizontalNumBolts As Integer?
    Private _AntennaRedundantSubHorizontalBoltEdgeDistance As Double?
    Private _AntennaRedundantSubHorizontalGageG1Distance As Double?
    Private _AntennaRedundantSubHorizontalNetWidthDeduct As Double?
    Private _AntennaRedundantSubHorizontalUFactor As Double?
    Private _AntennaRedundantVerticalBoltGrade As String
    Private _AntennaRedundantVerticalBoltSize As Double?
    Private _AntennaRedundantVerticalNumBolts As Integer?
    Private _AntennaRedundantVerticalBoltEdgeDistance As Double?
    Private _AntennaRedundantVerticalGageG1Distance As Double?
    Private _AntennaRedundantVerticalNetWidthDeduct As Double?
    Private _AntennaRedundantVerticalUFactor As Double?
    Private _AntennaRedundantHipBoltGrade As String
    Private _AntennaRedundantHipBoltSize As Double?
    Private _AntennaRedundantHipNumBolts As Integer?
    Private _AntennaRedundantHipBoltEdgeDistance As Double?
    Private _AntennaRedundantHipGageG1Distance As Double?
    Private _AntennaRedundantHipNetWidthDeduct As Double?
    Private _AntennaRedundantHipUFactor As Double?
    Private _AntennaRedundantHipDiagonalBoltGrade As String
    Private _AntennaRedundantHipDiagonalBoltSize As Double?
    Private _AntennaRedundantHipDiagonalNumBolts As Integer?
    Private _AntennaRedundantHipDiagonalBoltEdgeDistance As Double?
    Private _AntennaRedundantHipDiagonalGageG1Distance As Double?
    Private _AntennaRedundantHipDiagonalNetWidthDeduct As Double?
    Private _AntennaRedundantHipDiagonalUFactor As Double?
    Private _AntennaDiagonalOutOfPlaneRestraint As Boolean?
    Private _AntennaTopGirtOutOfPlaneRestraint As Boolean?
    Private _AntennaBottomGirtOutOfPlaneRestraint As Boolean?
    Private _AntennaMidGirtOutOfPlaneRestraint As Boolean?
    Private _AntennaHorizontalOutOfPlaneRestraint As Boolean?
    Private _AntennaSecondaryHorizontalOutOfPlaneRestraint As Boolean?
    Private _AntennaDiagOffsetNEY As Double?
    Private _AntennaDiagOffsetNEX As Double?
    Private _AntennaDiagOffsetPEY As Double?
    Private _AntennaDiagOffsetPEX As Double?
    Private _AntennaKbraceOffsetNEY As Double?
    Private _AntennaKbraceOffsetNEX As Double?
    Private _AntennaKbraceOffsetPEY As Double?
    Private _AntennaKbraceOffsetPEX As Double?

    '<Category("TNX Antenna Record"), Description(""), DisplayName("Antennarec")>
    ' <DataMember()> Public Property Rec() As Integer?
    '    Get
    '        Return Me._AntennaRec
    '    End Get
    '    Set
    '        Me._AntennaRec = Value
    '    End Set
    'End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennabracetype")>
    <DataMember()> Public Property AntennaBraceType() As String
        Get
            Return Me._AntennaBraceType
        End Get
        Set
            Me._AntennaBraceType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaheight")>
    <DataMember()> Public Property AntennaHeight() As Double?
        Get
            Return Me._AntennaHeight
        End Get
        Set
            Me._AntennaHeight = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagonalspacing")>
    <DataMember()> Public Property AntennaDiagonalSpacing() As Double?
        Get
            Return Me._AntennaDiagonalSpacing
        End Get
        Set
            Me._AntennaDiagonalSpacing = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagonalspacingex")>
    <DataMember()> Public Property AntennaDiagonalSpacingEx() As Double?
        Get
            Return Me._AntennaDiagonalSpacingEx
        End Get
        Set
            Me._AntennaDiagonalSpacingEx = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennanumsections")>
    <DataMember()> Public Property AntennaNumSections() As Integer?
        Get
            Return Me._AntennaNumSections
        End Get
        Set
            Me._AntennaNumSections = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennanumsesctions")>
    <DataMember()> Public Property AntennaNumSesctions() As Integer?
        Get
            Return Me._AntennaNumSesctions
        End Get
        Set
            Me._AntennaNumSesctions = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennasectionlength")>
    <DataMember()> Public Property AntennaSectionLength() As Double?
        Get
            Return Me._AntennaSectionLength
        End Get
        Set
            Me._AntennaSectionLength = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennalegtype")>
    <DataMember()> Public Property AntennaLegType() As String
        Get
            Return Me._AntennaLegType
        End Get
        Set
            Me._AntennaLegType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennalegsize")>
    <DataMember()> Public Property AntennaLegSize() As String
        Get
            Return Me._AntennaLegSize
        End Get
        Set
            Me._AntennaLegSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaleggrade")>
    <DataMember()> Public Property AntennaLegGrade() As Double?
        Get
            Return Me._AntennaLegGrade
        End Get
        Set
            Me._AntennaLegGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennalegmatlgrade")>
    <DataMember()> Public Property AntennaLegMatlGrade() As String
        Get
            Return Me._AntennaLegMatlGrade
        End Get
        Set
            Me._AntennaLegMatlGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagonalgrade")>
    <DataMember()> Public Property AntennaDiagonalGrade() As Double?
        Get
            Return Me._AntennaDiagonalGrade
        End Get
        Set
            Me._AntennaDiagonalGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagonalmatlgrade")>
    <DataMember()> Public Property AntennaDiagonalMatlGrade() As String
        Get
            Return Me._AntennaDiagonalMatlGrade
        End Get
        Set
            Me._AntennaDiagonalMatlGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnerbracinggrade")>
    <DataMember()> Public Property AntennaInnerBracingGrade() As Double?
        Get
            Return Me._AntennaInnerBracingGrade
        End Get
        Set
            Me._AntennaInnerBracingGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnerbracingmatlgrade")>
    <DataMember()> Public Property AntennaInnerBracingMatlGrade() As String
        Get
            Return Me._AntennaInnerBracingMatlGrade
        End Get
        Set
            Me._AntennaInnerBracingMatlGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennatopgirtgrade")>
    <DataMember()> Public Property AntennaTopGirtGrade() As Double?
        Get
            Return Me._AntennaTopGirtGrade
        End Get
        Set
            Me._AntennaTopGirtGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennatopgirtmatlgrade")>
    <DataMember()> Public Property AntennaTopGirtMatlGrade() As String
        Get
            Return Me._AntennaTopGirtMatlGrade
        End Get
        Set
            Me._AntennaTopGirtMatlGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennabotgirtgrade")>
    <DataMember()> Public Property AntennaBotGirtGrade() As Double?
        Get
            Return Me._AntennaBotGirtGrade
        End Get
        Set
            Me._AntennaBotGirtGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennabotgirtmatlgrade")>
    <DataMember()> Public Property AntennaBotGirtMatlGrade() As String
        Get
            Return Me._AntennaBotGirtMatlGrade
        End Get
        Set
            Me._AntennaBotGirtMatlGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnergirtgrade")>
    <DataMember()> Public Property AntennaInnerGirtGrade() As Double?
        Get
            Return Me._AntennaInnerGirtGrade
        End Get
        Set
            Me._AntennaInnerGirtGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnergirtmatlgrade")>
    <DataMember()> Public Property AntennaInnerGirtMatlGrade() As String
        Get
            Return Me._AntennaInnerGirtMatlGrade
        End Get
        Set
            Me._AntennaInnerGirtMatlGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennalonghorizontalgrade")>
    <DataMember()> Public Property AntennaLongHorizontalGrade() As Double?
        Get
            Return Me._AntennaLongHorizontalGrade
        End Get
        Set
            Me._AntennaLongHorizontalGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennalonghorizontalmatlgrade")>
    <DataMember()> Public Property AntennaLongHorizontalMatlGrade() As String
        Get
            Return Me._AntennaLongHorizontalMatlGrade
        End Get
        Set
            Me._AntennaLongHorizontalMatlGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennashorthorizontalgrade")>
    <DataMember()> Public Property AntennaShortHorizontalGrade() As Double?
        Get
            Return Me._AntennaShortHorizontalGrade
        End Get
        Set
            Me._AntennaShortHorizontalGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennashorthorizontalmatlgrade")>
    <DataMember()> Public Property AntennaShortHorizontalMatlGrade() As String
        Get
            Return Me._AntennaShortHorizontalMatlGrade
        End Get
        Set
            Me._AntennaShortHorizontalMatlGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagonaltype")>
    <DataMember()> Public Property AntennaDiagonalType() As String
        Get
            Return Me._AntennaDiagonalType
        End Get
        Set
            Me._AntennaDiagonalType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagonalsize")>
    <DataMember()> Public Property AntennaDiagonalSize() As String
        Get
            Return Me._AntennaDiagonalSize
        End Get
        Set
            Me._AntennaDiagonalSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnerbracingtype")>
    <DataMember()> Public Property AntennaInnerBracingType() As String
        Get
            Return Me._AntennaInnerBracingType
        End Get
        Set
            Me._AntennaInnerBracingType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnerbracingsize")>
    <DataMember()> Public Property AntennaInnerBracingSize() As String
        Get
            Return Me._AntennaInnerBracingSize
        End Get
        Set
            Me._AntennaInnerBracingSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennatopgirttype")>
    <DataMember()> Public Property AntennaTopGirtType() As String
        Get
            Return Me._AntennaTopGirtType
        End Get
        Set
            Me._AntennaTopGirtType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennatopgirtsize")>
    <DataMember()> Public Property AntennaTopGirtSize() As String
        Get
            Return Me._AntennaTopGirtSize
        End Get
        Set
            Me._AntennaTopGirtSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennabotgirttype")>
    <DataMember()> Public Property AntennaBotGirtType() As String
        Get
            Return Me._AntennaBotGirtType
        End Get
        Set
            Me._AntennaBotGirtType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennabotgirtsize")>
    <DataMember()> Public Property AntennaBotGirtSize() As String
        Get
            Return Me._AntennaBotGirtSize
        End Get
        Set
            Me._AntennaBotGirtSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennatopgirtoffset")>
    <DataMember()> Public Property AntennaTopGirtOffset() As Double?
        Get
            Return Me._AntennaTopGirtOffset
        End Get
        Set
            Me._AntennaTopGirtOffset = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennabotgirtoffset")>
    <DataMember()> Public Property AntennaBotGirtOffset() As Double?
        Get
            Return Me._AntennaBotGirtOffset
        End Get
        Set
            Me._AntennaBotGirtOffset = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennahaskbraceendpanels")>
    <DataMember()> Public Property AntennaHasKBraceEndPanels() As Boolean?
        Get
            Return Me._AntennaHasKBraceEndPanels
        End Get
        Set
            Me._AntennaHasKBraceEndPanels = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennahashorizontals")>
    <DataMember()> Public Property AntennaHasHorizontals() As Boolean?
        Get
            Return Me._AntennaHasHorizontals
        End Get
        Set
            Me._AntennaHasHorizontals = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennalonghorizontaltype")>
    <DataMember()> Public Property AntennaLongHorizontalType() As String
        Get
            Return Me._AntennaLongHorizontalType
        End Get
        Set
            Me._AntennaLongHorizontalType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennalonghorizontalsize")>
    <DataMember()> Public Property AntennaLongHorizontalSize() As String
        Get
            Return Me._AntennaLongHorizontalSize
        End Get
        Set
            Me._AntennaLongHorizontalSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennashorthorizontaltype")>
    <DataMember()> Public Property AntennaShortHorizontalType() As String
        Get
            Return Me._AntennaShortHorizontalType
        End Get
        Set
            Me._AntennaShortHorizontalType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennashorthorizontalsize")>
    <DataMember()> Public Property AntennaShortHorizontalSize() As String
        Get
            Return Me._AntennaShortHorizontalSize
        End Get
        Set
            Me._AntennaShortHorizontalSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantgrade")>
    <DataMember()> Public Property AntennaRedundantGrade() As Double?
        Get
            Return Me._AntennaRedundantGrade
        End Get
        Set
            Me._AntennaRedundantGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantmatlgrade")>
    <DataMember()> Public Property AntennaRedundantMatlGrade() As String
        Get
            Return Me._AntennaRedundantMatlGrade
        End Get
        Set
            Me._AntennaRedundantMatlGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanttype")>
    <DataMember()> Public Property AntennaRedundantType() As String
        Get
            Return Me._AntennaRedundantType
        End Get
        Set
            Me._AntennaRedundantType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantdiagtype")>
    <DataMember()> Public Property AntennaRedundantDiagType() As String
        Get
            Return Me._AntennaRedundantDiagType
        End Get
        Set
            Me._AntennaRedundantDiagType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubdiagonaltype")>
    <DataMember()> Public Property AntennaRedundantSubDiagonalType() As String
        Get
            Return Me._AntennaRedundantSubDiagonalType
        End Get
        Set
            Me._AntennaRedundantSubDiagonalType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubhorizontaltype")>
    <DataMember()> Public Property AntennaRedundantSubHorizontalType() As String
        Get
            Return Me._AntennaRedundantSubHorizontalType
        End Get
        Set
            Me._AntennaRedundantSubHorizontalType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantverticaltype")>
    <DataMember()> Public Property AntennaRedundantVerticalType() As String
        Get
            Return Me._AntennaRedundantVerticalType
        End Get
        Set
            Me._AntennaRedundantVerticalType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthiptype")>
    <DataMember()> Public Property AntennaRedundantHipType() As String
        Get
            Return Me._AntennaRedundantHipType
        End Get
        Set
            Me._AntennaRedundantHipType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipdiagonaltype")>
    <DataMember()> Public Property AntennaRedundantHipDiagonalType() As String
        Get
            Return Me._AntennaRedundantHipDiagonalType
        End Get
        Set
            Me._AntennaRedundantHipDiagonalType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthorizontalsize")>
    <DataMember()> Public Property AntennaRedundantHorizontalSize() As String
        Get
            Return Me._AntennaRedundantHorizontalSize
        End Get
        Set
            Me._AntennaRedundantHorizontalSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthorizontalsize2")>
    <DataMember()> Public Property AntennaRedundantHorizontalSize2() As String
        Get
            Return Me._AntennaRedundantHorizontalSize2
        End Get
        Set
            Me._AntennaRedundantHorizontalSize2 = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthorizontalsize3")>
    <DataMember()> Public Property AntennaRedundantHorizontalSize3() As String
        Get
            Return Me._AntennaRedundantHorizontalSize3
        End Get
        Set
            Me._AntennaRedundantHorizontalSize3 = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthorizontalsize4")>
    <DataMember()> Public Property AntennaRedundantHorizontalSize4() As String
        Get
            Return Me._AntennaRedundantHorizontalSize4
        End Get
        Set
            Me._AntennaRedundantHorizontalSize4 = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantdiagonalsize")>
    <DataMember()> Public Property AntennaRedundantDiagonalSize() As String
        Get
            Return Me._AntennaRedundantDiagonalSize
        End Get
        Set
            Me._AntennaRedundantDiagonalSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantdiagonalsize2")>
    <DataMember()> Public Property AntennaRedundantDiagonalSize2() As String
        Get
            Return Me._AntennaRedundantDiagonalSize2
        End Get
        Set
            Me._AntennaRedundantDiagonalSize2 = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantdiagonalsize3")>
    <DataMember()> Public Property AntennaRedundantDiagonalSize3() As String
        Get
            Return Me._AntennaRedundantDiagonalSize3
        End Get
        Set
            Me._AntennaRedundantDiagonalSize3 = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantdiagonalsize4")>
    <DataMember()> Public Property AntennaRedundantDiagonalSize4() As String
        Get
            Return Me._AntennaRedundantDiagonalSize4
        End Get
        Set
            Me._AntennaRedundantDiagonalSize4 = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubhorizontalsize")>
    <DataMember()> Public Property AntennaRedundantSubHorizontalSize() As String
        Get
            Return Me._AntennaRedundantSubHorizontalSize
        End Get
        Set
            Me._AntennaRedundantSubHorizontalSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubdiagonalsize")>
    <DataMember()> Public Property AntennaRedundantSubDiagonalSize() As String
        Get
            Return Me._AntennaRedundantSubDiagonalSize
        End Get
        Set
            Me._AntennaRedundantSubDiagonalSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennasubdiaglocation")>
    <DataMember()> Public Property AntennaSubDiagLocation() As Double?
        Get
            Return Me._AntennaSubDiagLocation
        End Get
        Set
            Me._AntennaSubDiagLocation = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantverticalsize")>
    <DataMember()> Public Property AntennaRedundantVerticalSize() As String
        Get
            Return Me._AntennaRedundantVerticalSize
        End Get
        Set
            Me._AntennaRedundantVerticalSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipdiagonalsize")>
    <DataMember()> Public Property AntennaRedundantHipDiagonalSize() As String
        Get
            Return Me._AntennaRedundantHipDiagonalSize
        End Get
        Set
            Me._AntennaRedundantHipDiagonalSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipdiagonalsize2")>
    <DataMember()> Public Property AntennaRedundantHipDiagonalSize2() As String
        Get
            Return Me._AntennaRedundantHipDiagonalSize2
        End Get
        Set
            Me._AntennaRedundantHipDiagonalSize2 = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipdiagonalsize3")>
    <DataMember()> Public Property AntennaRedundantHipDiagonalSize3() As String
        Get
            Return Me._AntennaRedundantHipDiagonalSize3
        End Get
        Set
            Me._AntennaRedundantHipDiagonalSize3 = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipdiagonalsize4")>
    <DataMember()> Public Property AntennaRedundantHipDiagonalSize4() As String
        Get
            Return Me._AntennaRedundantHipDiagonalSize4
        End Get
        Set
            Me._AntennaRedundantHipDiagonalSize4 = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipsize")>
    <DataMember()> Public Property AntennaRedundantHipSize() As String
        Get
            Return Me._AntennaRedundantHipSize
        End Get
        Set
            Me._AntennaRedundantHipSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipsize2")>
    <DataMember()> Public Property AntennaRedundantHipSize2() As String
        Get
            Return Me._AntennaRedundantHipSize2
        End Get
        Set
            Me._AntennaRedundantHipSize2 = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipsize3")>
    <DataMember()> Public Property AntennaRedundantHipSize3() As String
        Get
            Return Me._AntennaRedundantHipSize3
        End Get
        Set
            Me._AntennaRedundantHipSize3 = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipsize4")>
    <DataMember()> Public Property AntennaRedundantHipSize4() As String
        Get
            Return Me._AntennaRedundantHipSize4
        End Get
        Set
            Me._AntennaRedundantHipSize4 = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennanuminnergirts")>
    <DataMember()> Public Property AntennaNumInnerGirts() As Integer?
        Get
            Return Me._AntennaNumInnerGirts
        End Get
        Set
            Me._AntennaNumInnerGirts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnergirttype")>
    <DataMember()> Public Property AntennaInnerGirtType() As String
        Get
            Return Me._AntennaInnerGirtType
        End Get
        Set
            Me._AntennaInnerGirtType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnergirtsize")>
    <DataMember()> Public Property AntennaInnerGirtSize() As String
        Get
            Return Me._AntennaInnerGirtSize
        End Get
        Set
            Me._AntennaInnerGirtSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennapoleshapetype")>
    <DataMember()> Public Property AntennaPoleShapeType() As String
        Get
            Return Me._AntennaPoleShapeType
        End Get
        Set
            Me._AntennaPoleShapeType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennapolesize")>
    <DataMember()> Public Property AntennaPoleSize() As String
        Get
            Return Me._AntennaPoleSize
        End Get
        Set
            Me._AntennaPoleSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennapolegrade")>
    <DataMember()> Public Property AntennaPoleGrade() As Double?
        Get
            Return Me._AntennaPoleGrade
        End Get
        Set
            Me._AntennaPoleGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennapolematlgrade")>
    <DataMember()> Public Property AntennaPoleMatlGrade() As String
        Get
            Return Me._AntennaPoleMatlGrade
        End Get
        Set
            Me._AntennaPoleMatlGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennapolesplicelength")>
    <DataMember()> Public Property AntennaPoleSpliceLength() As Double?
        Get
            Return Me._AntennaPoleSpliceLength
        End Get
        Set
            Me._AntennaPoleSpliceLength = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennataperpolenumsides")>
    <DataMember()> Public Property AntennaTaperPoleNumSides() As Integer?
        Get
            Return Me._AntennaTaperPoleNumSides
        End Get
        Set
            Me._AntennaTaperPoleNumSides = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennataperpoletopdiameter")>
    <DataMember()> Public Property AntennaTaperPoleTopDiameter() As Double?
        Get
            Return Me._AntennaTaperPoleTopDiameter
        End Get
        Set
            Me._AntennaTaperPoleTopDiameter = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennataperpolebotdiameter")>
    <DataMember()> Public Property AntennaTaperPoleBotDiameter() As Double?
        Get
            Return Me._AntennaTaperPoleBotDiameter
        End Get
        Set
            Me._AntennaTaperPoleBotDiameter = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennataperpolewallthickness")>
    <DataMember()> Public Property AntennaTaperPoleWallThickness() As Double?
        Get
            Return Me._AntennaTaperPoleWallThickness
        End Get
        Set
            Me._AntennaTaperPoleWallThickness = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennataperpolebendradius")>
    <DataMember()> Public Property AntennaTaperPoleBendRadius() As Double?
        Get
            Return Me._AntennaTaperPoleBendRadius
        End Get
        Set
            Me._AntennaTaperPoleBendRadius = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennataperpolegrade")>
    <DataMember()> Public Property AntennaTaperPoleGrade() As Double?
        Get
            Return Me._AntennaTaperPoleGrade
        End Get
        Set
            Me._AntennaTaperPoleGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennataperpolematlgrade")>
    <DataMember()> Public Property AntennaTaperPoleMatlGrade() As String
        Get
            Return Me._AntennaTaperPoleMatlGrade
        End Get
        Set
            Me._AntennaTaperPoleMatlGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaswmult")>
    <DataMember()> Public Property AntennaSWMult() As Double?
        Get
            Return Me._AntennaSWMult
        End Get
        Set
            Me._AntennaSWMult = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennawpmult")>
    <DataMember()> Public Property AntennaWPMult() As Double?
        Get
            Return Me._AntennaWPMult
        End Get
        Set
            Me._AntennaWPMult = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaautocalcksingleangle")>
    <DataMember()> Public Property AntennaAutoCalcKSingleAngle() As Double?
        Get
            Return Me._AntennaAutoCalcKSingleAngle
        End Get
        Set
            Me._AntennaAutoCalcKSingleAngle = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaautocalcksolidround")>
    <DataMember()> Public Property AntennaAutoCalcKSolidRound() As Double?
        Get
            Return Me._AntennaAutoCalcKSolidRound
        End Get
        Set
            Me._AntennaAutoCalcKSolidRound = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaafgusset")>
    <DataMember()> Public Property AntennaAfGusset() As Double?
        Get
            Return Me._AntennaAfGusset
        End Get
        Set
            Me._AntennaAfGusset = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennatfgusset")>
    <DataMember()> Public Property AntennaTfGusset() As Double?
        Get
            Return Me._AntennaTfGusset
        End Get
        Set
            Me._AntennaTfGusset = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennagussetboltedgedistance")>
    <DataMember()> Public Property AntennaGussetBoltEdgeDistance() As Double?
        Get
            Return Me._AntennaGussetBoltEdgeDistance
        End Get
        Set
            Me._AntennaGussetBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennagussetgrade")>
    <DataMember()> Public Property AntennaGussetGrade() As Double?
        Get
            Return Me._AntennaGussetGrade
        End Get
        Set
            Me._AntennaGussetGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennagussetmatlgrade")>
    <DataMember()> Public Property AntennaGussetMatlGrade() As String
        Get
            Return Me._AntennaGussetMatlGrade
        End Get
        Set
            Me._AntennaGussetMatlGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaafmult")>
    <DataMember()> Public Property AntennaAfMult() As Double?
        Get
            Return Me._AntennaAfMult
        End Get
        Set
            Me._AntennaAfMult = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaarmult")>
    <DataMember()> Public Property AntennaArMult() As Double?
        Get
            Return Me._AntennaArMult
        End Get
        Set
            Me._AntennaArMult = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaflatipapole")>
    <DataMember()> Public Property AntennaFlatIPAPole() As Double?
        Get
            Return Me._AntennaFlatIPAPole
        End Get
        Set
            Me._AntennaFlatIPAPole = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaroundipapole")>
    <DataMember()> Public Property AntennaRoundIPAPole() As Double?
        Get
            Return Me._AntennaRoundIPAPole
        End Get
        Set
            Me._AntennaRoundIPAPole = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaflatipaleg")>
    <DataMember()> Public Property AntennaFlatIPALeg() As Double?
        Get
            Return Me._AntennaFlatIPALeg
        End Get
        Set
            Me._AntennaFlatIPALeg = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaroundipaleg")>
    <DataMember()> Public Property AntennaRoundIPALeg() As Double?
        Get
            Return Me._AntennaRoundIPALeg
        End Get
        Set
            Me._AntennaRoundIPALeg = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaflatipahorizontal")>
    <DataMember()> Public Property AntennaFlatIPAHorizontal() As Double?
        Get
            Return Me._AntennaFlatIPAHorizontal
        End Get
        Set
            Me._AntennaFlatIPAHorizontal = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaroundipahorizontal")>
    <DataMember()> Public Property AntennaRoundIPAHorizontal() As Double?
        Get
            Return Me._AntennaRoundIPAHorizontal
        End Get
        Set
            Me._AntennaRoundIPAHorizontal = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaflatipadiagonal")>
    <DataMember()> Public Property AntennaFlatIPADiagonal() As Double?
        Get
            Return Me._AntennaFlatIPADiagonal
        End Get
        Set
            Me._AntennaFlatIPADiagonal = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaroundipadiagonal")>
    <DataMember()> Public Property AntennaRoundIPADiagonal() As Double?
        Get
            Return Me._AntennaRoundIPADiagonal
        End Get
        Set
            Me._AntennaRoundIPADiagonal = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennacsa_S37_Speedupfactor")>
    <DataMember()> Public Property AntennaCSA_S37_SpeedUpFactor() As Double?
        Get
            Return Me._AntennaCSA_S37_SpeedUpFactor
        End Get
        Set
            Me._AntennaCSA_S37_SpeedUpFactor = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaklegs")>
    <DataMember()> Public Property AntennaKLegs() As Double?
        Get
            Return Me._AntennaKLegs
        End Get
        Set
            Me._AntennaKLegs = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakxbraceddiags")>
    <DataMember()> Public Property AntennaKXBracedDiags() As Double?
        Get
            Return Me._AntennaKXBracedDiags
        End Get
        Set
            Me._AntennaKXBracedDiags = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakkbraceddiags")>
    <DataMember()> Public Property AntennaKKBracedDiags() As Double?
        Get
            Return Me._AntennaKKBracedDiags
        End Get
        Set
            Me._AntennaKKBracedDiags = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakzbraceddiags")>
    <DataMember()> Public Property AntennaKZBracedDiags() As Double?
        Get
            Return Me._AntennaKZBracedDiags
        End Get
        Set
            Me._AntennaKZBracedDiags = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakhorzs")>
    <DataMember()> Public Property AntennaKHorzs() As Double?
        Get
            Return Me._AntennaKHorzs
        End Get
        Set
            Me._AntennaKHorzs = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaksechorzs")>
    <DataMember()> Public Property AntennaKSecHorzs() As Double?
        Get
            Return Me._AntennaKSecHorzs
        End Get
        Set
            Me._AntennaKSecHorzs = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakgirts")>
    <DataMember()> Public Property AntennaKGirts() As Double?
        Get
            Return Me._AntennaKGirts
        End Get
        Set
            Me._AntennaKGirts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakinners")>
    <DataMember()> Public Property AntennaKInners() As Double?
        Get
            Return Me._AntennaKInners
        End Get
        Set
            Me._AntennaKInners = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakxbraceddiagsy")>
    <DataMember()> Public Property AntennaKXBracedDiagsY() As Double?
        Get
            Return Me._AntennaKXBracedDiagsY
        End Get
        Set
            Me._AntennaKXBracedDiagsY = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakkbraceddiagsy")>
    <DataMember()> Public Property AntennaKKBracedDiagsY() As Double?
        Get
            Return Me._AntennaKKBracedDiagsY
        End Get
        Set
            Me._AntennaKKBracedDiagsY = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakzbraceddiagsy")>
    <DataMember()> Public Property AntennaKZBracedDiagsY() As Double?
        Get
            Return Me._AntennaKZBracedDiagsY
        End Get
        Set
            Me._AntennaKZBracedDiagsY = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakhorzsy")>
    <DataMember()> Public Property AntennaKHorzsY() As Double?
        Get
            Return Me._AntennaKHorzsY
        End Get
        Set
            Me._AntennaKHorzsY = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaksechorzsy")>
    <DataMember()> Public Property AntennaKSecHorzsY() As Double?
        Get
            Return Me._AntennaKSecHorzsY
        End Get
        Set
            Me._AntennaKSecHorzsY = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakgirtsy")>
    <DataMember()> Public Property AntennaKGirtsY() As Double?
        Get
            Return Me._AntennaKGirtsY
        End Get
        Set
            Me._AntennaKGirtsY = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakinnersy")>
    <DataMember()> Public Property AntennaKInnersY() As Double?
        Get
            Return Me._AntennaKInnersY
        End Get
        Set
            Me._AntennaKInnersY = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakredhorz")>
    <DataMember()> Public Property AntennaKRedHorz() As Double?
        Get
            Return Me._AntennaKRedHorz
        End Get
        Set
            Me._AntennaKRedHorz = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakreddiag")>
    <DataMember()> Public Property AntennaKRedDiag() As Double?
        Get
            Return Me._AntennaKRedDiag
        End Get
        Set
            Me._AntennaKRedDiag = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakredsubdiag")>
    <DataMember()> Public Property AntennaKRedSubDiag() As Double?
        Get
            Return Me._AntennaKRedSubDiag
        End Get
        Set
            Me._AntennaKRedSubDiag = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakredsubhorz")>
    <DataMember()> Public Property AntennaKRedSubHorz() As Double?
        Get
            Return Me._AntennaKRedSubHorz
        End Get
        Set
            Me._AntennaKRedSubHorz = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakredvert")>
    <DataMember()> Public Property AntennaKRedVert() As Double?
        Get
            Return Me._AntennaKRedVert
        End Get
        Set
            Me._AntennaKRedVert = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakredhip")>
    <DataMember()> Public Property AntennaKRedHip() As Double?
        Get
            Return Me._AntennaKRedHip
        End Get
        Set
            Me._AntennaKRedHip = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakredhipdiag")>
    <DataMember()> Public Property AntennaKRedHipDiag() As Double?
        Get
            Return Me._AntennaKRedHipDiag
        End Get
        Set
            Me._AntennaKRedHipDiag = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaktlx")>
    <DataMember()> Public Property AntennaKTLX() As Double?
        Get
            Return Me._AntennaKTLX
        End Get
        Set
            Me._AntennaKTLX = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaktlz")>
    <DataMember()> Public Property AntennaKTLZ() As Double?
        Get
            Return Me._AntennaKTLZ
        End Get
        Set
            Me._AntennaKTLZ = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaktlleg")>
    <DataMember()> Public Property AntennaKTLLeg() As Double?
        Get
            Return Me._AntennaKTLLeg
        End Get
        Set
            Me._AntennaKTLLeg = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnerktlx")>
    <DataMember()> Public Property AntennaInnerKTLX() As Double?
        Get
            Return Me._AntennaInnerKTLX
        End Get
        Set
            Me._AntennaInnerKTLX = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnerktlz")>
    <DataMember()> Public Property AntennaInnerKTLZ() As Double?
        Get
            Return Me._AntennaInnerKTLZ
        End Get
        Set
            Me._AntennaInnerKTLZ = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnerktlleg")>
    <DataMember()> Public Property AntennaInnerKTLLeg() As Double?
        Get
            Return Me._AntennaInnerKTLLeg
        End Get
        Set
            Me._AntennaInnerKTLLeg = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennastitchboltlocationhoriz")>
    <DataMember()> Public Property AntennaStitchBoltLocationHoriz() As String
        Get
            Return Me._AntennaStitchBoltLocationHoriz
        End Get
        Set
            Me._AntennaStitchBoltLocationHoriz = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennastitchboltlocationdiag")>
    <DataMember()> Public Property AntennaStitchBoltLocationDiag() As String
        Get
            Return Me._AntennaStitchBoltLocationDiag
        End Get
        Set
            Me._AntennaStitchBoltLocationDiag = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennastitchspacing")>
    <DataMember()> Public Property AntennaStitchSpacing() As Double?
        Get
            Return Me._AntennaStitchSpacing
        End Get
        Set
            Me._AntennaStitchSpacing = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennastitchspacinghorz")>
    <DataMember()> Public Property AntennaStitchSpacingHorz() As Double?
        Get
            Return Me._AntennaStitchSpacingHorz
        End Get
        Set
            Me._AntennaStitchSpacingHorz = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennastitchspacingdiag")>
    <DataMember()> Public Property AntennaStitchSpacingDiag() As Double?
        Get
            Return Me._AntennaStitchSpacingDiag
        End Get
        Set
            Me._AntennaStitchSpacingDiag = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennastitchspacingred")>
    <DataMember()> Public Property AntennaStitchSpacingRed() As Double?
        Get
            Return Me._AntennaStitchSpacingRed
        End Get
        Set
            Me._AntennaStitchSpacingRed = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennalegnetwidthdeduct")>
    <DataMember()> Public Property AntennaLegNetWidthDeduct() As Double?
        Get
            Return Me._AntennaLegNetWidthDeduct
        End Get
        Set
            Me._AntennaLegNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennalegufactor")>
    <DataMember()> Public Property AntennaLegUFactor() As Double?
        Get
            Return Me._AntennaLegUFactor
        End Get
        Set
            Me._AntennaLegUFactor = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagonalnetwidthdeduct")>
    <DataMember()> Public Property AntennaDiagonalNetWidthDeduct() As Double?
        Get
            Return Me._AntennaDiagonalNetWidthDeduct
        End Get
        Set
            Me._AntennaDiagonalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennatopgirtnetwidthdeduct")>
    <DataMember()> Public Property AntennaTopGirtNetWidthDeduct() As Double?
        Get
            Return Me._AntennaTopGirtNetWidthDeduct
        End Get
        Set
            Me._AntennaTopGirtNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennabotgirtnetwidthdeduct")>
    <DataMember()> Public Property AntennaBotGirtNetWidthDeduct() As Double?
        Get
            Return Me._AntennaBotGirtNetWidthDeduct
        End Get
        Set
            Me._AntennaBotGirtNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnergirtnetwidthdeduct")>
    <DataMember()> Public Property AntennaInnerGirtNetWidthDeduct() As Double?
        Get
            Return Me._AntennaInnerGirtNetWidthDeduct
        End Get
        Set
            Me._AntennaInnerGirtNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennahorizontalnetwidthdeduct")>
    <DataMember()> Public Property AntennaHorizontalNetWidthDeduct() As Double?
        Get
            Return Me._AntennaHorizontalNetWidthDeduct
        End Get
        Set
            Me._AntennaHorizontalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennashorthorizontalnetwidthdeduct")>
    <DataMember()> Public Property AntennaShortHorizontalNetWidthDeduct() As Double?
        Get
            Return Me._AntennaShortHorizontalNetWidthDeduct
        End Get
        Set
            Me._AntennaShortHorizontalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagonalufactor")>
    <DataMember()> Public Property AntennaDiagonalUFactor() As Double?
        Get
            Return Me._AntennaDiagonalUFactor
        End Get
        Set
            Me._AntennaDiagonalUFactor = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennatopgirtufactor")>
    <DataMember()> Public Property AntennaTopGirtUFactor() As Double?
        Get
            Return Me._AntennaTopGirtUFactor
        End Get
        Set
            Me._AntennaTopGirtUFactor = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennabotgirtufactor")>
    <DataMember()> Public Property AntennaBotGirtUFactor() As Double?
        Get
            Return Me._AntennaBotGirtUFactor
        End Get
        Set
            Me._AntennaBotGirtUFactor = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnergirtufactor")>
    <DataMember()> Public Property AntennaInnerGirtUFactor() As Double?
        Get
            Return Me._AntennaInnerGirtUFactor
        End Get
        Set
            Me._AntennaInnerGirtUFactor = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennahorizontalufactor")>
    <DataMember()> Public Property AntennaHorizontalUFactor() As Double?
        Get
            Return Me._AntennaHorizontalUFactor
        End Get
        Set
            Me._AntennaHorizontalUFactor = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennashorthorizontalufactor")>
    <DataMember()> Public Property AntennaShortHorizontalUFactor() As Double?
        Get
            Return Me._AntennaShortHorizontalUFactor
        End Get
        Set
            Me._AntennaShortHorizontalUFactor = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennalegconntype")>
    <DataMember()> Public Property AntennaLegConnType() As String
        Get
            Return Me._AntennaLegConnType
        End Get
        Set
            Me._AntennaLegConnType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennalegnumbolts")>
    <DataMember()> Public Property AntennaLegNumBolts() As Integer?
        Get
            Return Me._AntennaLegNumBolts
        End Get
        Set
            Me._AntennaLegNumBolts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagonalnumbolts")>
    <DataMember()> Public Property AntennaDiagonalNumBolts() As Integer?
        Get
            Return Me._AntennaDiagonalNumBolts
        End Get
        Set
            Me._AntennaDiagonalNumBolts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennatopgirtnumbolts")>
    <DataMember()> Public Property AntennaTopGirtNumBolts() As Integer?
        Get
            Return Me._AntennaTopGirtNumBolts
        End Get
        Set
            Me._AntennaTopGirtNumBolts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennabotgirtnumbolts")>
    <DataMember()> Public Property AntennaBotGirtNumBolts() As Integer?
        Get
            Return Me._AntennaBotGirtNumBolts
        End Get
        Set
            Me._AntennaBotGirtNumBolts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnergirtnumbolts")>
    <DataMember()> Public Property AntennaInnerGirtNumBolts() As Integer?
        Get
            Return Me._AntennaInnerGirtNumBolts
        End Get
        Set
            Me._AntennaInnerGirtNumBolts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennahorizontalnumbolts")>
    <DataMember()> Public Property AntennaHorizontalNumBolts() As Integer?
        Get
            Return Me._AntennaHorizontalNumBolts
        End Get
        Set
            Me._AntennaHorizontalNumBolts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennashorthorizontalnumbolts")>
    <DataMember()> Public Property AntennaShortHorizontalNumBolts() As Integer?
        Get
            Return Me._AntennaShortHorizontalNumBolts
        End Get
        Set
            Me._AntennaShortHorizontalNumBolts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennalegboltgrade")>
    <DataMember()> Public Property AntennaLegBoltGrade() As String
        Get
            Return Me._AntennaLegBoltGrade
        End Get
        Set
            Me._AntennaLegBoltGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennalegboltsize")>
    <DataMember()> Public Property AntennaLegBoltSize() As Double?
        Get
            Return Me._AntennaLegBoltSize
        End Get
        Set
            Me._AntennaLegBoltSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagonalboltgrade")>
    <DataMember()> Public Property AntennaDiagonalBoltGrade() As String
        Get
            Return Me._AntennaDiagonalBoltGrade
        End Get
        Set
            Me._AntennaDiagonalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagonalboltsize")>
    <DataMember()> Public Property AntennaDiagonalBoltSize() As Double?
        Get
            Return Me._AntennaDiagonalBoltSize
        End Get
        Set
            Me._AntennaDiagonalBoltSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennatopgirtboltgrade")>
    <DataMember()> Public Property AntennaTopGirtBoltGrade() As String
        Get
            Return Me._AntennaTopGirtBoltGrade
        End Get
        Set
            Me._AntennaTopGirtBoltGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennatopgirtboltsize")>
    <DataMember()> Public Property AntennaTopGirtBoltSize() As Double?
        Get
            Return Me._AntennaTopGirtBoltSize
        End Get
        Set
            Me._AntennaTopGirtBoltSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennabotgirtboltgrade")>
    <DataMember()> Public Property AntennaBotGirtBoltGrade() As String
        Get
            Return Me._AntennaBotGirtBoltGrade
        End Get
        Set
            Me._AntennaBotGirtBoltGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennabotgirtboltsize")>
    <DataMember()> Public Property AntennaBotGirtBoltSize() As Double?
        Get
            Return Me._AntennaBotGirtBoltSize
        End Get
        Set
            Me._AntennaBotGirtBoltSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnergirtboltgrade")>
    <DataMember()> Public Property AntennaInnerGirtBoltGrade() As String
        Get
            Return Me._AntennaInnerGirtBoltGrade
        End Get
        Set
            Me._AntennaInnerGirtBoltGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnergirtboltsize")>
    <DataMember()> Public Property AntennaInnerGirtBoltSize() As Double?
        Get
            Return Me._AntennaInnerGirtBoltSize
        End Get
        Set
            Me._AntennaInnerGirtBoltSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennahorizontalboltgrade")>
    <DataMember()> Public Property AntennaHorizontalBoltGrade() As String
        Get
            Return Me._AntennaHorizontalBoltGrade
        End Get
        Set
            Me._AntennaHorizontalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennahorizontalboltsize")>
    <DataMember()> Public Property AntennaHorizontalBoltSize() As Double?
        Get
            Return Me._AntennaHorizontalBoltSize
        End Get
        Set
            Me._AntennaHorizontalBoltSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennashorthorizontalboltgrade")>
    <DataMember()> Public Property AntennaShortHorizontalBoltGrade() As String
        Get
            Return Me._AntennaShortHorizontalBoltGrade
        End Get
        Set
            Me._AntennaShortHorizontalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennashorthorizontalboltsize")>
    <DataMember()> Public Property AntennaShortHorizontalBoltSize() As Double?
        Get
            Return Me._AntennaShortHorizontalBoltSize
        End Get
        Set
            Me._AntennaShortHorizontalBoltSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennalegboltedgedistance")>
    <DataMember()> Public Property AntennaLegBoltEdgeDistance() As Double?
        Get
            Return Me._AntennaLegBoltEdgeDistance
        End Get
        Set
            Me._AntennaLegBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagonalboltedgedistance")>
    <DataMember()> Public Property AntennaDiagonalBoltEdgeDistance() As Double?
        Get
            Return Me._AntennaDiagonalBoltEdgeDistance
        End Get
        Set
            Me._AntennaDiagonalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennatopgirtboltedgedistance")>
    <DataMember()> Public Property AntennaTopGirtBoltEdgeDistance() As Double?
        Get
            Return Me._AntennaTopGirtBoltEdgeDistance
        End Get
        Set
            Me._AntennaTopGirtBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennabotgirtboltedgedistance")>
    <DataMember()> Public Property AntennaBotGirtBoltEdgeDistance() As Double?
        Get
            Return Me._AntennaBotGirtBoltEdgeDistance
        End Get
        Set
            Me._AntennaBotGirtBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnergirtboltedgedistance")>
    <DataMember()> Public Property AntennaInnerGirtBoltEdgeDistance() As Double?
        Get
            Return Me._AntennaInnerGirtBoltEdgeDistance
        End Get
        Set
            Me._AntennaInnerGirtBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennahorizontalboltedgedistance")>
    <DataMember()> Public Property AntennaHorizontalBoltEdgeDistance() As Double?
        Get
            Return Me._AntennaHorizontalBoltEdgeDistance
        End Get
        Set
            Me._AntennaHorizontalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennashorthorizontalboltedgedistance")>
    <DataMember()> Public Property AntennaShortHorizontalBoltEdgeDistance() As Double?
        Get
            Return Me._AntennaShortHorizontalBoltEdgeDistance
        End Get
        Set
            Me._AntennaShortHorizontalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagonalgageg1Distance")>
    <DataMember()> Public Property AntennaDiagonalGageG1Distance() As Double?
        Get
            Return Me._AntennaDiagonalGageG1Distance
        End Get
        Set
            Me._AntennaDiagonalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennatopgirtgageg1Distance")>
    <DataMember()> Public Property AntennaTopGirtGageG1Distance() As Double?
        Get
            Return Me._AntennaTopGirtGageG1Distance
        End Get
        Set
            Me._AntennaTopGirtGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennabotgirtgageg1Distance")>
    <DataMember()> Public Property AntennaBotGirtGageG1Distance() As Double?
        Get
            Return Me._AntennaBotGirtGageG1Distance
        End Get
        Set
            Me._AntennaBotGirtGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnergirtgageg1Distance")>
    <DataMember()> Public Property AntennaInnerGirtGageG1Distance() As Double?
        Get
            Return Me._AntennaInnerGirtGageG1Distance
        End Get
        Set
            Me._AntennaInnerGirtGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennahorizontalgageg1Distance")>
    <DataMember()> Public Property AntennaHorizontalGageG1Distance() As Double?
        Get
            Return Me._AntennaHorizontalGageG1Distance
        End Get
        Set
            Me._AntennaHorizontalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennashorthorizontalgageg1Distance")>
    <DataMember()> Public Property AntennaShortHorizontalGageG1Distance() As Double?
        Get
            Return Me._AntennaShortHorizontalGageG1Distance
        End Get
        Set
            Me._AntennaShortHorizontalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthorizontalboltgrade")>
    <DataMember()> Public Property AntennaRedundantHorizontalBoltGrade() As String
        Get
            Return Me._AntennaRedundantHorizontalBoltGrade
        End Get
        Set
            Me._AntennaRedundantHorizontalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthorizontalboltsize")>
    <DataMember()> Public Property AntennaRedundantHorizontalBoltSize() As Double?
        Get
            Return Me._AntennaRedundantHorizontalBoltSize
        End Get
        Set
            Me._AntennaRedundantHorizontalBoltSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthorizontalnumbolts")>
    <DataMember()> Public Property AntennaRedundantHorizontalNumBolts() As Integer?
        Get
            Return Me._AntennaRedundantHorizontalNumBolts
        End Get
        Set
            Me._AntennaRedundantHorizontalNumBolts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthorizontalboltedgedistance")>
    <DataMember()> Public Property AntennaRedundantHorizontalBoltEdgeDistance() As Double?
        Get
            Return Me._AntennaRedundantHorizontalBoltEdgeDistance
        End Get
        Set
            Me._AntennaRedundantHorizontalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthorizontalgageg1Distance")>
    <DataMember()> Public Property AntennaRedundantHorizontalGageG1Distance() As Double?
        Get
            Return Me._AntennaRedundantHorizontalGageG1Distance
        End Get
        Set
            Me._AntennaRedundantHorizontalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthorizontalnetwidthdeduct")>
    <DataMember()> Public Property AntennaRedundantHorizontalNetWidthDeduct() As Double?
        Get
            Return Me._AntennaRedundantHorizontalNetWidthDeduct
        End Get
        Set
            Me._AntennaRedundantHorizontalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthorizontalufactor")>
    <DataMember()> Public Property AntennaRedundantHorizontalUFactor() As Double?
        Get
            Return Me._AntennaRedundantHorizontalUFactor
        End Get
        Set
            Me._AntennaRedundantHorizontalUFactor = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantdiagonalboltgrade")>
    <DataMember()> Public Property AntennaRedundantDiagonalBoltGrade() As String
        Get
            Return Me._AntennaRedundantDiagonalBoltGrade
        End Get
        Set
            Me._AntennaRedundantDiagonalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantdiagonalboltsize")>
    <DataMember()> Public Property AntennaRedundantDiagonalBoltSize() As Double?
        Get
            Return Me._AntennaRedundantDiagonalBoltSize
        End Get
        Set
            Me._AntennaRedundantDiagonalBoltSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantdiagonalnumbolts")>
    <DataMember()> Public Property AntennaRedundantDiagonalNumBolts() As Integer?
        Get
            Return Me._AntennaRedundantDiagonalNumBolts
        End Get
        Set
            Me._AntennaRedundantDiagonalNumBolts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantdiagonalboltedgedistance")>
    <DataMember()> Public Property AntennaRedundantDiagonalBoltEdgeDistance() As Double?
        Get
            Return Me._AntennaRedundantDiagonalBoltEdgeDistance
        End Get
        Set
            Me._AntennaRedundantDiagonalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantdiagonalgageg1Distance")>
    <DataMember()> Public Property AntennaRedundantDiagonalGageG1Distance() As Double?
        Get
            Return Me._AntennaRedundantDiagonalGageG1Distance
        End Get
        Set
            Me._AntennaRedundantDiagonalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantdiagonalnetwidthdeduct")>
    <DataMember()> Public Property AntennaRedundantDiagonalNetWidthDeduct() As Double?
        Get
            Return Me._AntennaRedundantDiagonalNetWidthDeduct
        End Get
        Set
            Me._AntennaRedundantDiagonalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantdiagonalufactor")>
    <DataMember()> Public Property AntennaRedundantDiagonalUFactor() As Double?
        Get
            Return Me._AntennaRedundantDiagonalUFactor
        End Get
        Set
            Me._AntennaRedundantDiagonalUFactor = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubdiagonalboltgrade")>
    <DataMember()> Public Property AntennaRedundantSubDiagonalBoltGrade() As String
        Get
            Return Me._AntennaRedundantSubDiagonalBoltGrade
        End Get
        Set
            Me._AntennaRedundantSubDiagonalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubdiagonalboltsize")>
    <DataMember()> Public Property AntennaRedundantSubDiagonalBoltSize() As Double?
        Get
            Return Me._AntennaRedundantSubDiagonalBoltSize
        End Get
        Set
            Me._AntennaRedundantSubDiagonalBoltSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubdiagonalnumbolts")>
    <DataMember()> Public Property AntennaRedundantSubDiagonalNumBolts() As Integer?
        Get
            Return Me._AntennaRedundantSubDiagonalNumBolts
        End Get
        Set
            Me._AntennaRedundantSubDiagonalNumBolts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubdiagonalboltedgedistance")>
    <DataMember()> Public Property AntennaRedundantSubDiagonalBoltEdgeDistance() As Double?
        Get
            Return Me._AntennaRedundantSubDiagonalBoltEdgeDistance
        End Get
        Set
            Me._AntennaRedundantSubDiagonalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubdiagonalgageg1Distance")>
    <DataMember()> Public Property AntennaRedundantSubDiagonalGageG1Distance() As Double?
        Get
            Return Me._AntennaRedundantSubDiagonalGageG1Distance
        End Get
        Set
            Me._AntennaRedundantSubDiagonalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubdiagonalnetwidthdeduct")>
    <DataMember()> Public Property AntennaRedundantSubDiagonalNetWidthDeduct() As Double?
        Get
            Return Me._AntennaRedundantSubDiagonalNetWidthDeduct
        End Get
        Set
            Me._AntennaRedundantSubDiagonalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubdiagonalufactor")>
    <DataMember()> Public Property AntennaRedundantSubDiagonalUFactor() As Double?
        Get
            Return Me._AntennaRedundantSubDiagonalUFactor
        End Get
        Set
            Me._AntennaRedundantSubDiagonalUFactor = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubhorizontalboltgrade")>
    <DataMember()> Public Property AntennaRedundantSubHorizontalBoltGrade() As String
        Get
            Return Me._AntennaRedundantSubHorizontalBoltGrade
        End Get
        Set
            Me._AntennaRedundantSubHorizontalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubhorizontalboltsize")>
    <DataMember()> Public Property AntennaRedundantSubHorizontalBoltSize() As Double?
        Get
            Return Me._AntennaRedundantSubHorizontalBoltSize
        End Get
        Set
            Me._AntennaRedundantSubHorizontalBoltSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubhorizontalnumbolts")>
    <DataMember()> Public Property AntennaRedundantSubHorizontalNumBolts() As Integer?
        Get
            Return Me._AntennaRedundantSubHorizontalNumBolts
        End Get
        Set
            Me._AntennaRedundantSubHorizontalNumBolts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubhorizontalboltedgedistance")>
    <DataMember()> Public Property AntennaRedundantSubHorizontalBoltEdgeDistance() As Double?
        Get
            Return Me._AntennaRedundantSubHorizontalBoltEdgeDistance
        End Get
        Set
            Me._AntennaRedundantSubHorizontalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubhorizontalgageg1Distance")>
    <DataMember()> Public Property AntennaRedundantSubHorizontalGageG1Distance() As Double?
        Get
            Return Me._AntennaRedundantSubHorizontalGageG1Distance
        End Get
        Set
            Me._AntennaRedundantSubHorizontalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubhorizontalnetwidthdeduct")>
    <DataMember()> Public Property AntennaRedundantSubHorizontalNetWidthDeduct() As Double?
        Get
            Return Me._AntennaRedundantSubHorizontalNetWidthDeduct
        End Get
        Set
            Me._AntennaRedundantSubHorizontalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubhorizontalufactor")>
    <DataMember()> Public Property AntennaRedundantSubHorizontalUFactor() As Double?
        Get
            Return Me._AntennaRedundantSubHorizontalUFactor
        End Get
        Set
            Me._AntennaRedundantSubHorizontalUFactor = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantverticalboltgrade")>
    <DataMember()> Public Property AntennaRedundantVerticalBoltGrade() As String
        Get
            Return Me._AntennaRedundantVerticalBoltGrade
        End Get
        Set
            Me._AntennaRedundantVerticalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantverticalboltsize")>
    <DataMember()> Public Property AntennaRedundantVerticalBoltSize() As Double?
        Get
            Return Me._AntennaRedundantVerticalBoltSize
        End Get
        Set
            Me._AntennaRedundantVerticalBoltSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantverticalnumbolts")>
    <DataMember()> Public Property AntennaRedundantVerticalNumBolts() As Integer?
        Get
            Return Me._AntennaRedundantVerticalNumBolts
        End Get
        Set
            Me._AntennaRedundantVerticalNumBolts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantverticalboltedgedistance")>
    <DataMember()> Public Property AntennaRedundantVerticalBoltEdgeDistance() As Double?
        Get
            Return Me._AntennaRedundantVerticalBoltEdgeDistance
        End Get
        Set
            Me._AntennaRedundantVerticalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantverticalgageg1Distance")>
    <DataMember()> Public Property AntennaRedundantVerticalGageG1Distance() As Double?
        Get
            Return Me._AntennaRedundantVerticalGageG1Distance
        End Get
        Set
            Me._AntennaRedundantVerticalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantverticalnetwidthdeduct")>
    <DataMember()> Public Property AntennaRedundantVerticalNetWidthDeduct() As Double?
        Get
            Return Me._AntennaRedundantVerticalNetWidthDeduct
        End Get
        Set
            Me._AntennaRedundantVerticalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantverticalufactor")>
    <DataMember()> Public Property AntennaRedundantVerticalUFactor() As Double?
        Get
            Return Me._AntennaRedundantVerticalUFactor
        End Get
        Set
            Me._AntennaRedundantVerticalUFactor = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipboltgrade")>
    <DataMember()> Public Property AntennaRedundantHipBoltGrade() As String
        Get
            Return Me._AntennaRedundantHipBoltGrade
        End Get
        Set
            Me._AntennaRedundantHipBoltGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipboltsize")>
    <DataMember()> Public Property AntennaRedundantHipBoltSize() As Double?
        Get
            Return Me._AntennaRedundantHipBoltSize
        End Get
        Set
            Me._AntennaRedundantHipBoltSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipnumbolts")>
    <DataMember()> Public Property AntennaRedundantHipNumBolts() As Integer?
        Get
            Return Me._AntennaRedundantHipNumBolts
        End Get
        Set
            Me._AntennaRedundantHipNumBolts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipboltedgedistance")>
    <DataMember()> Public Property AntennaRedundantHipBoltEdgeDistance() As Double?
        Get
            Return Me._AntennaRedundantHipBoltEdgeDistance
        End Get
        Set
            Me._AntennaRedundantHipBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipgageg1Distance")>
    <DataMember()> Public Property AntennaRedundantHipGageG1Distance() As Double?
        Get
            Return Me._AntennaRedundantHipGageG1Distance
        End Get
        Set
            Me._AntennaRedundantHipGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipnetwidthdeduct")>
    <DataMember()> Public Property AntennaRedundantHipNetWidthDeduct() As Double?
        Get
            Return Me._AntennaRedundantHipNetWidthDeduct
        End Get
        Set
            Me._AntennaRedundantHipNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipufactor")>
    <DataMember()> Public Property AntennaRedundantHipUFactor() As Double?
        Get
            Return Me._AntennaRedundantHipUFactor
        End Get
        Set
            Me._AntennaRedundantHipUFactor = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipdiagonalboltgrade")>
    <DataMember()> Public Property AntennaRedundantHipDiagonalBoltGrade() As String
        Get
            Return Me._AntennaRedundantHipDiagonalBoltGrade
        End Get
        Set
            Me._AntennaRedundantHipDiagonalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipdiagonalboltsize")>
    <DataMember()> Public Property AntennaRedundantHipDiagonalBoltSize() As Double?
        Get
            Return Me._AntennaRedundantHipDiagonalBoltSize
        End Get
        Set
            Me._AntennaRedundantHipDiagonalBoltSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipdiagonalnumbolts")>
    <DataMember()> Public Property AntennaRedundantHipDiagonalNumBolts() As Integer?
        Get
            Return Me._AntennaRedundantHipDiagonalNumBolts
        End Get
        Set
            Me._AntennaRedundantHipDiagonalNumBolts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipdiagonalboltedgedistance")>
    <DataMember()> Public Property AntennaRedundantHipDiagonalBoltEdgeDistance() As Double?
        Get
            Return Me._AntennaRedundantHipDiagonalBoltEdgeDistance
        End Get
        Set
            Me._AntennaRedundantHipDiagonalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipdiagonalgageg1Distance")>
    <DataMember()> Public Property AntennaRedundantHipDiagonalGageG1Distance() As Double?
        Get
            Return Me._AntennaRedundantHipDiagonalGageG1Distance
        End Get
        Set
            Me._AntennaRedundantHipDiagonalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipdiagonalnetwidthdeduct")>
    <DataMember()> Public Property AntennaRedundantHipDiagonalNetWidthDeduct() As Double?
        Get
            Return Me._AntennaRedundantHipDiagonalNetWidthDeduct
        End Get
        Set
            Me._AntennaRedundantHipDiagonalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipdiagonalufactor")>
    <DataMember()> Public Property AntennaRedundantHipDiagonalUFactor() As Double?
        Get
            Return Me._AntennaRedundantHipDiagonalUFactor
        End Get
        Set
            Me._AntennaRedundantHipDiagonalUFactor = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagonaloutofplanerestraint")>
    <DataMember()> Public Property AntennaDiagonalOutOfPlaneRestraint() As Boolean?
        Get
            Return Me._AntennaDiagonalOutOfPlaneRestraint
        End Get
        Set
            Me._AntennaDiagonalOutOfPlaneRestraint = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennatopgirtoutofplanerestraint")>
    <DataMember()> Public Property AntennaTopGirtOutOfPlaneRestraint() As Boolean?
        Get
            Return Me._AntennaTopGirtOutOfPlaneRestraint
        End Get
        Set
            Me._AntennaTopGirtOutOfPlaneRestraint = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennabottomgirtoutofplanerestraint")>
    <DataMember()> Public Property AntennaBottomGirtOutOfPlaneRestraint() As Boolean?
        Get
            Return Me._AntennaBottomGirtOutOfPlaneRestraint
        End Get
        Set
            Me._AntennaBottomGirtOutOfPlaneRestraint = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennamidgirtoutofplanerestraint")>
    <DataMember()> Public Property AntennaMidGirtOutOfPlaneRestraint() As Boolean?
        Get
            Return Me._AntennaMidGirtOutOfPlaneRestraint
        End Get
        Set
            Me._AntennaMidGirtOutOfPlaneRestraint = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennahorizontaloutofplanerestraint")>
    <DataMember()> Public Property AntennaHorizontalOutOfPlaneRestraint() As Boolean?
        Get
            Return Me._AntennaHorizontalOutOfPlaneRestraint
        End Get
        Set
            Me._AntennaHorizontalOutOfPlaneRestraint = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennasecondaryhorizontaloutofplanerestraint")>
    <DataMember()> Public Property AntennaSecondaryHorizontalOutOfPlaneRestraint() As Boolean?
        Get
            Return Me._AntennaSecondaryHorizontalOutOfPlaneRestraint
        End Get
        Set
            Me._AntennaSecondaryHorizontalOutOfPlaneRestraint = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagoffsetney")>
    <DataMember()> Public Property AntennaDiagOffsetNEY() As Double?
        Get
            Return Me._AntennaDiagOffsetNEY
        End Get
        Set
            Me._AntennaDiagOffsetNEY = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagoffsetnex")>
    <DataMember()> Public Property AntennaDiagOffsetNEX() As Double?
        Get
            Return Me._AntennaDiagOffsetNEX
        End Get
        Set
            Me._AntennaDiagOffsetNEX = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagoffsetpey")>
    <DataMember()> Public Property AntennaDiagOffsetPEY() As Double?
        Get
            Return Me._AntennaDiagOffsetPEY
        End Get
        Set
            Me._AntennaDiagOffsetPEY = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagoffsetpex")>
    <DataMember()> Public Property AntennaDiagOffsetPEX() As Double?
        Get
            Return Me._AntennaDiagOffsetPEX
        End Get
        Set
            Me._AntennaDiagOffsetPEX = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakbraceoffsetney")>
    <DataMember()> Public Property AntennaKbraceOffsetNEY() As Double?
        Get
            Return Me._AntennaKbraceOffsetNEY
        End Get
        Set
            Me._AntennaKbraceOffsetNEY = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakbraceoffsetnex")>
    <DataMember()> Public Property AntennaKbraceOffsetNEX() As Double?
        Get
            Return Me._AntennaKbraceOffsetNEX
        End Get
        Set
            Me._AntennaKbraceOffsetNEX = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakbraceoffsetpey")>
    <DataMember()> Public Property AntennaKbraceOffsetPEY() As Double?
        Get
            Return Me._AntennaKbraceOffsetPEY
        End Get
        Set
            Me._AntennaKbraceOffsetPEY = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakbraceoffsetpex")>
    <DataMember()> Public Property AntennaKbraceOffsetPEX() As Double?
        Get
            Return Me._AntennaKbraceOffsetPEX
        End Get
        Set
            Me._AntennaKbraceOffsetPEX = Value
        End Set
    End Property
#End Region

#Region "Constructors"
    Public Sub New(Optional ByVal Parent As EDSObject = Nothing)
        If Parent IsNot Nothing Then Me.Absorb(Parent)
    End Sub

    Public Sub New(data As DataRow, Optional ByVal Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Me.ID = DBtoNullableInt(data.Item("ID"))
        Me.Rec = DBtoNullableInt(data.Item("AntennaRec"))
        Me.AntennaBraceType = DBtoStr(data.Item("AntennaBraceType"))
        Me.AntennaHeight = DBtoNullableDbl(data.Item("AntennaHeight"))
        Me.AntennaDiagonalSpacing = DBtoNullableDbl(data.Item("AntennaDiagonalSpacing"))
        Me.AntennaDiagonalSpacingEx = DBtoNullableDbl(data.Item("AntennaDiagonalSpacingEx"))
        Me.AntennaNumSections = DBtoNullableInt(data.Item("AntennaNumSections"))
        Me.AntennaNumSesctions = DBtoNullableInt(data.Item("AntennaNumSesctions"))
        Me.AntennaSectionLength = DBtoNullableDbl(data.Item("AntennaSectionLength"))
        Me.AntennaLegType = DBtoStr(data.Item("AntennaLegType"))
        Me.AntennaLegSize = DBtoStr(data.Item("AntennaLegSize"))
        Me.AntennaLegGrade = DBtoNullableDbl(data.Item("AntennaLegGrade"))
        Me.AntennaLegMatlGrade = DBtoStr(data.Item("AntennaLegMatlGrade"))
        Me.AntennaDiagonalGrade = DBtoNullableDbl(data.Item("AntennaDiagonalGrade"))
        Me.AntennaDiagonalMatlGrade = DBtoStr(data.Item("AntennaDiagonalMatlGrade"))
        Me.AntennaInnerBracingGrade = DBtoNullableDbl(data.Item("AntennaInnerBracingGrade"))
        Me.AntennaInnerBracingMatlGrade = DBtoStr(data.Item("AntennaInnerBracingMatlGrade"))
        Me.AntennaTopGirtGrade = DBtoNullableDbl(data.Item("AntennaTopGirtGrade"))
        Me.AntennaTopGirtMatlGrade = DBtoStr(data.Item("AntennaTopGirtMatlGrade"))
        Me.AntennaBotGirtGrade = DBtoNullableDbl(data.Item("AntennaBotGirtGrade"))
        Me.AntennaBotGirtMatlGrade = DBtoStr(data.Item("AntennaBotGirtMatlGrade"))
        Me.AntennaInnerGirtGrade = DBtoNullableDbl(data.Item("AntennaInnerGirtGrade"))
        Me.AntennaInnerGirtMatlGrade = DBtoStr(data.Item("AntennaInnerGirtMatlGrade"))
        Me.AntennaLongHorizontalGrade = DBtoNullableDbl(data.Item("AntennaLongHorizontalGrade"))
        Me.AntennaLongHorizontalMatlGrade = DBtoStr(data.Item("AntennaLongHorizontalMatlGrade"))
        Me.AntennaShortHorizontalGrade = DBtoNullableDbl(data.Item("AntennaShortHorizontalGrade"))
        Me.AntennaShortHorizontalMatlGrade = DBtoStr(data.Item("AntennaShortHorizontalMatlGrade"))
        Me.AntennaDiagonalType = DBtoStr(data.Item("AntennaDiagonalType"))
        Me.AntennaDiagonalSize = DBtoStr(data.Item("AntennaDiagonalSize"))
        Me.AntennaInnerBracingType = DBtoStr(data.Item("AntennaInnerBracingType"))
        Me.AntennaInnerBracingSize = DBtoStr(data.Item("AntennaInnerBracingSize"))
        Me.AntennaTopGirtType = DBtoStr(data.Item("AntennaTopGirtType"))
        Me.AntennaTopGirtSize = DBtoStr(data.Item("AntennaTopGirtSize"))
        Me.AntennaBotGirtType = DBtoStr(data.Item("AntennaBotGirtType"))
        Me.AntennaBotGirtSize = DBtoStr(data.Item("AntennaBotGirtSize"))
        Me.AntennaTopGirtOffset = DBtoNullableDbl(data.Item("AntennaTopGirtOffset"))
        Me.AntennaBotGirtOffset = DBtoNullableDbl(data.Item("AntennaBotGirtOffset"))
        Me.AntennaHasKBraceEndPanels = DBtoNullableBool(data.Item("AntennaHasKBraceEndPanels"))
        Me.AntennaHasHorizontals = DBtoNullableBool(data.Item("AntennaHasHorizontals"))
        Me.AntennaLongHorizontalType = DBtoStr(data.Item("AntennaLongHorizontalType"))
        Me.AntennaLongHorizontalSize = DBtoStr(data.Item("AntennaLongHorizontalSize"))
        Me.AntennaShortHorizontalType = DBtoStr(data.Item("AntennaShortHorizontalType"))
        Me.AntennaShortHorizontalSize = DBtoStr(data.Item("AntennaShortHorizontalSize"))
        Me.AntennaRedundantGrade = DBtoNullableDbl(data.Item("AntennaRedundantGrade"))
        Me.AntennaRedundantMatlGrade = DBtoStr(data.Item("AntennaRedundantMatlGrade"))
        Me.AntennaRedundantType = DBtoStr(data.Item("AntennaRedundantType"))
        Me.AntennaRedundantDiagType = DBtoStr(data.Item("AntennaRedundantDiagType"))
        Me.AntennaRedundantSubDiagonalType = DBtoStr(data.Item("AntennaRedundantSubDiagonalType"))
        Me.AntennaRedundantSubHorizontalType = DBtoStr(data.Item("AntennaRedundantSubHorizontalType"))
        Me.AntennaRedundantVerticalType = DBtoStr(data.Item("AntennaRedundantVerticalType"))
        Me.AntennaRedundantHipType = DBtoStr(data.Item("AntennaRedundantHipType"))
        Me.AntennaRedundantHipDiagonalType = DBtoStr(data.Item("AntennaRedundantHipDiagonalType"))
        Me.AntennaRedundantHorizontalSize = DBtoStr(data.Item("AntennaRedundantHorizontalSize"))
        Me.AntennaRedundantHorizontalSize2 = DBtoStr(data.Item("AntennaRedundantHorizontalSize2"))
        Me.AntennaRedundantHorizontalSize3 = DBtoStr(data.Item("AntennaRedundantHorizontalSize3"))
        Me.AntennaRedundantHorizontalSize4 = DBtoStr(data.Item("AntennaRedundantHorizontalSize4"))
        Me.AntennaRedundantDiagonalSize = DBtoStr(data.Item("AntennaRedundantDiagonalSize"))
        Me.AntennaRedundantDiagonalSize2 = DBtoStr(data.Item("AntennaRedundantDiagonalSize2"))
        Me.AntennaRedundantDiagonalSize3 = DBtoStr(data.Item("AntennaRedundantDiagonalSize3"))
        Me.AntennaRedundantDiagonalSize4 = DBtoStr(data.Item("AntennaRedundantDiagonalSize4"))
        Me.AntennaRedundantSubHorizontalSize = DBtoStr(data.Item("AntennaRedundantSubHorizontalSize"))
        Me.AntennaRedundantSubDiagonalSize = DBtoStr(data.Item("AntennaRedundantSubDiagonalSize"))
        Me.AntennaSubDiagLocation = DBtoNullableDbl(data.Item("AntennaSubDiagLocation"))
        Me.AntennaRedundantVerticalSize = DBtoStr(data.Item("AntennaRedundantVerticalSize"))
        Me.AntennaRedundantHipDiagonalSize = DBtoStr(data.Item("AntennaRedundantHipDiagonalSize"))
        Me.AntennaRedundantHipDiagonalSize2 = DBtoStr(data.Item("AntennaRedundantHipDiagonalSize2"))
        Me.AntennaRedundantHipDiagonalSize3 = DBtoStr(data.Item("AntennaRedundantHipDiagonalSize3"))
        Me.AntennaRedundantHipDiagonalSize4 = DBtoStr(data.Item("AntennaRedundantHipDiagonalSize4"))
        Me.AntennaRedundantHipSize = DBtoStr(data.Item("AntennaRedundantHipSize"))
        Me.AntennaRedundantHipSize2 = DBtoStr(data.Item("AntennaRedundantHipSize2"))
        Me.AntennaRedundantHipSize3 = DBtoStr(data.Item("AntennaRedundantHipSize3"))
        Me.AntennaRedundantHipSize4 = DBtoStr(data.Item("AntennaRedundantHipSize4"))
        Me.AntennaNumInnerGirts = DBtoNullableInt(data.Item("AntennaNumInnerGirts"))
        Me.AntennaInnerGirtType = DBtoStr(data.Item("AntennaInnerGirtType"))
        Me.AntennaInnerGirtSize = DBtoStr(data.Item("AntennaInnerGirtSize"))
        Me.AntennaPoleShapeType = DBtoStr(data.Item("AntennaPoleShapeType"))
        Me.AntennaPoleSize = DBtoStr(data.Item("AntennaPoleSize"))
        Me.AntennaPoleGrade = DBtoNullableDbl(data.Item("AntennaPoleGrade"))
        Me.AntennaPoleMatlGrade = DBtoStr(data.Item("AntennaPoleMatlGrade"))
        Me.AntennaPoleSpliceLength = DBtoNullableDbl(data.Item("AntennaPoleSpliceLength"))
        Me.AntennaTaperPoleNumSides = DBtoNullableInt(data.Item("AntennaTaperPoleNumSides"))
        Me.AntennaTaperPoleTopDiameter = DBtoNullableDbl(data.Item("AntennaTaperPoleTopDiameter"))
        Me.AntennaTaperPoleBotDiameter = DBtoNullableDbl(data.Item("AntennaTaperPoleBotDiameter"))
        Me.AntennaTaperPoleWallThickness = DBtoNullableDbl(data.Item("AntennaTaperPoleWallThickness"))
        Me.AntennaTaperPoleBendRadius = DBtoNullableDbl(data.Item("AntennaTaperPoleBendRadius"))
        Me.AntennaTaperPoleGrade = DBtoNullableDbl(data.Item("AntennaTaperPoleGrade"))
        Me.AntennaTaperPoleMatlGrade = DBtoStr(data.Item("AntennaTaperPoleMatlGrade"))
        Me.AntennaSWMult = DBtoNullableDbl(data.Item("AntennaSWMult"))
        Me.AntennaWPMult = DBtoNullableDbl(data.Item("AntennaWPMult"))
        Me.AntennaAutoCalcKSingleAngle = DBtoNullableDbl(data.Item("AntennaAutoCalcKSingleAngle"))
        Me.AntennaAutoCalcKSolidRound = DBtoNullableDbl(data.Item("AntennaAutoCalcKSolidRound"))
        Me.AntennaAfGusset = DBtoNullableDbl(data.Item("AntennaAfGusset"))
        Me.AntennaTfGusset = DBtoNullableDbl(data.Item("AntennaTfGusset"))
        Me.AntennaGussetBoltEdgeDistance = DBtoNullableDbl(data.Item("AntennaGussetBoltEdgeDistance"))
        Me.AntennaGussetGrade = DBtoNullableDbl(data.Item("AntennaGussetGrade"))
        Me.AntennaGussetMatlGrade = DBtoStr(data.Item("AntennaGussetMatlGrade"))
        Me.AntennaAfMult = DBtoNullableDbl(data.Item("AntennaAfMult"))
        Me.AntennaArMult = DBtoNullableDbl(data.Item("AntennaArMult"))
        Me.AntennaFlatIPAPole = DBtoNullableDbl(data.Item("AntennaFlatIPAPole"))
        Me.AntennaRoundIPAPole = DBtoNullableDbl(data.Item("AntennaRoundIPAPole"))
        Me.AntennaFlatIPALeg = DBtoNullableDbl(data.Item("AntennaFlatIPALeg"))
        Me.AntennaRoundIPALeg = DBtoNullableDbl(data.Item("AntennaRoundIPALeg"))
        Me.AntennaFlatIPAHorizontal = DBtoNullableDbl(data.Item("AntennaFlatIPAHorizontal"))
        Me.AntennaRoundIPAHorizontal = DBtoNullableDbl(data.Item("AntennaRoundIPAHorizontal"))
        Me.AntennaFlatIPADiagonal = DBtoNullableDbl(data.Item("AntennaFlatIPADiagonal"))
        Me.AntennaRoundIPADiagonal = DBtoNullableDbl(data.Item("AntennaRoundIPADiagonal"))
        Me.AntennaCSA_S37_SpeedUpFactor = DBtoNullableDbl(data.Item("AntennaCSA_S37_SpeedUpFactor"))
        Me.AntennaKLegs = DBtoNullableDbl(data.Item("AntennaKLegs"))
        Me.AntennaKXBracedDiags = DBtoNullableDbl(data.Item("AntennaKXBracedDiags"))
        Me.AntennaKKBracedDiags = DBtoNullableDbl(data.Item("AntennaKKBracedDiags"))
        Me.AntennaKZBracedDiags = DBtoNullableDbl(data.Item("AntennaKZBracedDiags"))
        Me.AntennaKHorzs = DBtoNullableDbl(data.Item("AntennaKHorzs"))
        Me.AntennaKSecHorzs = DBtoNullableDbl(data.Item("AntennaKSecHorzs"))
        Me.AntennaKGirts = DBtoNullableDbl(data.Item("AntennaKGirts"))
        Me.AntennaKInners = DBtoNullableDbl(data.Item("AntennaKInners"))
        Me.AntennaKXBracedDiagsY = DBtoNullableDbl(data.Item("AntennaKXBracedDiagsY"))
        Me.AntennaKKBracedDiagsY = DBtoNullableDbl(data.Item("AntennaKKBracedDiagsY"))
        Me.AntennaKZBracedDiagsY = DBtoNullableDbl(data.Item("AntennaKZBracedDiagsY"))
        Me.AntennaKHorzsY = DBtoNullableDbl(data.Item("AntennaKHorzsY"))
        Me.AntennaKSecHorzsY = DBtoNullableDbl(data.Item("AntennaKSecHorzsY"))
        Me.AntennaKGirtsY = DBtoNullableDbl(data.Item("AntennaKGirtsY"))
        Me.AntennaKInnersY = DBtoNullableDbl(data.Item("AntennaKInnersY"))
        Me.AntennaKRedHorz = DBtoNullableDbl(data.Item("AntennaKRedHorz"))
        Me.AntennaKRedDiag = DBtoNullableDbl(data.Item("AntennaKRedDiag"))
        Me.AntennaKRedSubDiag = DBtoNullableDbl(data.Item("AntennaKRedSubDiag"))
        Me.AntennaKRedSubHorz = DBtoNullableDbl(data.Item("AntennaKRedSubHorz"))
        Me.AntennaKRedVert = DBtoNullableDbl(data.Item("AntennaKRedVert"))
        Me.AntennaKRedHip = DBtoNullableDbl(data.Item("AntennaKRedHip"))
        Me.AntennaKRedHipDiag = DBtoNullableDbl(data.Item("AntennaKRedHipDiag"))
        Me.AntennaKTLX = DBtoNullableDbl(data.Item("AntennaKTLX"))
        Me.AntennaKTLZ = DBtoNullableDbl(data.Item("AntennaKTLZ"))
        Me.AntennaKTLLeg = DBtoNullableDbl(data.Item("AntennaKTLLeg"))
        Me.AntennaInnerKTLX = DBtoNullableDbl(data.Item("AntennaInnerKTLX"))
        Me.AntennaInnerKTLZ = DBtoNullableDbl(data.Item("AntennaInnerKTLZ"))
        Me.AntennaInnerKTLLeg = DBtoNullableDbl(data.Item("AntennaInnerKTLLeg"))
        Me.AntennaStitchBoltLocationHoriz = DBtoStr(data.Item("AntennaStitchBoltLocationHoriz"))
        Me.AntennaStitchBoltLocationDiag = DBtoStr(data.Item("AntennaStitchBoltLocationDiag"))
        Me.AntennaStitchSpacing = DBtoNullableDbl(data.Item("AntennaStitchSpacing"))
        Me.AntennaStitchSpacingHorz = DBtoNullableDbl(data.Item("AntennaStitchSpacingHorz"))
        Me.AntennaStitchSpacingDiag = DBtoNullableDbl(data.Item("AntennaStitchSpacingDiag"))
        Me.AntennaStitchSpacingRed = DBtoNullableDbl(data.Item("AntennaStitchSpacingRed"))
        Me.AntennaLegNetWidthDeduct = DBtoNullableDbl(data.Item("AntennaLegNetWidthDeduct"))
        Me.AntennaLegUFactor = DBtoNullableDbl(data.Item("AntennaLegUFactor"))
        Me.AntennaDiagonalNetWidthDeduct = DBtoNullableDbl(data.Item("AntennaDiagonalNetWidthDeduct"))
        Me.AntennaTopGirtNetWidthDeduct = DBtoNullableDbl(data.Item("AntennaTopGirtNetWidthDeduct"))
        Me.AntennaBotGirtNetWidthDeduct = DBtoNullableDbl(data.Item("AntennaBotGirtNetWidthDeduct"))
        Me.AntennaInnerGirtNetWidthDeduct = DBtoNullableDbl(data.Item("AntennaInnerGirtNetWidthDeduct"))
        Me.AntennaHorizontalNetWidthDeduct = DBtoNullableDbl(data.Item("AntennaHorizontalNetWidthDeduct"))
        Me.AntennaShortHorizontalNetWidthDeduct = DBtoNullableDbl(data.Item("AntennaShortHorizontalNetWidthDeduct"))
        Me.AntennaDiagonalUFactor = DBtoNullableDbl(data.Item("AntennaDiagonalUFactor"))
        Me.AntennaTopGirtUFactor = DBtoNullableDbl(data.Item("AntennaTopGirtUFactor"))
        Me.AntennaBotGirtUFactor = DBtoNullableDbl(data.Item("AntennaBotGirtUFactor"))
        Me.AntennaInnerGirtUFactor = DBtoNullableDbl(data.Item("AntennaInnerGirtUFactor"))
        Me.AntennaHorizontalUFactor = DBtoNullableDbl(data.Item("AntennaHorizontalUFactor"))
        Me.AntennaShortHorizontalUFactor = DBtoNullableDbl(data.Item("AntennaShortHorizontalUFactor"))
        Me.AntennaLegConnType = DBtoStr(data.Item("AntennaLegConnType"))
        Me.AntennaLegNumBolts = DBtoNullableInt(data.Item("AntennaLegNumBolts"))
        Me.AntennaDiagonalNumBolts = DBtoNullableInt(data.Item("AntennaDiagonalNumBolts"))
        Me.AntennaTopGirtNumBolts = DBtoNullableInt(data.Item("AntennaTopGirtNumBolts"))
        Me.AntennaBotGirtNumBolts = DBtoNullableInt(data.Item("AntennaBotGirtNumBolts"))
        Me.AntennaInnerGirtNumBolts = DBtoNullableInt(data.Item("AntennaInnerGirtNumBolts"))
        Me.AntennaHorizontalNumBolts = DBtoNullableInt(data.Item("AntennaHorizontalNumBolts"))
        Me.AntennaShortHorizontalNumBolts = DBtoNullableInt(data.Item("AntennaShortHorizontalNumBolts"))
        Me.AntennaLegBoltGrade = DBtoStr(data.Item("AntennaLegBoltGrade"))
        Me.AntennaLegBoltSize = DBtoNullableDbl(data.Item("AntennaLegBoltSize"))
        Me.AntennaDiagonalBoltGrade = DBtoStr(data.Item("AntennaDiagonalBoltGrade"))
        Me.AntennaDiagonalBoltSize = DBtoNullableDbl(data.Item("AntennaDiagonalBoltSize"))
        Me.AntennaTopGirtBoltGrade = DBtoStr(data.Item("AntennaTopGirtBoltGrade"))
        Me.AntennaTopGirtBoltSize = DBtoNullableDbl(data.Item("AntennaTopGirtBoltSize"))
        Me.AntennaBotGirtBoltGrade = DBtoStr(data.Item("AntennaBotGirtBoltGrade"))
        Me.AntennaBotGirtBoltSize = DBtoNullableDbl(data.Item("AntennaBotGirtBoltSize"))
        Me.AntennaInnerGirtBoltGrade = DBtoStr(data.Item("AntennaInnerGirtBoltGrade"))
        Me.AntennaInnerGirtBoltSize = DBtoNullableDbl(data.Item("AntennaInnerGirtBoltSize"))
        Me.AntennaHorizontalBoltGrade = DBtoStr(data.Item("AntennaHorizontalBoltGrade"))
        Me.AntennaHorizontalBoltSize = DBtoNullableDbl(data.Item("AntennaHorizontalBoltSize"))
        Me.AntennaShortHorizontalBoltGrade = DBtoStr(data.Item("AntennaShortHorizontalBoltGrade"))
        Me.AntennaShortHorizontalBoltSize = DBtoNullableDbl(data.Item("AntennaShortHorizontalBoltSize"))
        Me.AntennaLegBoltEdgeDistance = DBtoNullableDbl(data.Item("AntennaLegBoltEdgeDistance"))
        Me.AntennaDiagonalBoltEdgeDistance = DBtoNullableDbl(data.Item("AntennaDiagonalBoltEdgeDistance"))
        Me.AntennaTopGirtBoltEdgeDistance = DBtoNullableDbl(data.Item("AntennaTopGirtBoltEdgeDistance"))
        Me.AntennaBotGirtBoltEdgeDistance = DBtoNullableDbl(data.Item("AntennaBotGirtBoltEdgeDistance"))
        Me.AntennaInnerGirtBoltEdgeDistance = DBtoNullableDbl(data.Item("AntennaInnerGirtBoltEdgeDistance"))
        Me.AntennaHorizontalBoltEdgeDistance = DBtoNullableDbl(data.Item("AntennaHorizontalBoltEdgeDistance"))
        Me.AntennaShortHorizontalBoltEdgeDistance = DBtoNullableDbl(data.Item("AntennaShortHorizontalBoltEdgeDistance"))
        Me.AntennaDiagonalGageG1Distance = DBtoNullableDbl(data.Item("AntennaDiagonalGageG1Distance"))
        Me.AntennaTopGirtGageG1Distance = DBtoNullableDbl(data.Item("AntennaTopGirtGageG1Distance"))
        Me.AntennaBotGirtGageG1Distance = DBtoNullableDbl(data.Item("AntennaBotGirtGageG1Distance"))
        Me.AntennaInnerGirtGageG1Distance = DBtoNullableDbl(data.Item("AntennaInnerGirtGageG1Distance"))
        Me.AntennaHorizontalGageG1Distance = DBtoNullableDbl(data.Item("AntennaHorizontalGageG1Distance"))
        Me.AntennaShortHorizontalGageG1Distance = DBtoNullableDbl(data.Item("AntennaShortHorizontalGageG1Distance"))
        Me.AntennaRedundantHorizontalBoltGrade = DBtoStr(data.Item("AntennaRedundantHorizontalBoltGrade"))
        Me.AntennaRedundantHorizontalBoltSize = DBtoNullableDbl(data.Item("AntennaRedundantHorizontalBoltSize"))
        Me.AntennaRedundantHorizontalNumBolts = DBtoNullableInt(data.Item("AntennaRedundantHorizontalNumBolts"))
        Me.AntennaRedundantHorizontalBoltEdgeDistance = DBtoNullableDbl(data.Item("AntennaRedundantHorizontalBoltEdgeDistance"))
        Me.AntennaRedundantHorizontalGageG1Distance = DBtoNullableDbl(data.Item("AntennaRedundantHorizontalGageG1Distance"))
        Me.AntennaRedundantHorizontalNetWidthDeduct = DBtoNullableDbl(data.Item("AntennaRedundantHorizontalNetWidthDeduct"))
        Me.AntennaRedundantHorizontalUFactor = DBtoNullableDbl(data.Item("AntennaRedundantHorizontalUFactor"))
        Me.AntennaRedundantDiagonalBoltGrade = DBtoStr(data.Item("AntennaRedundantDiagonalBoltGrade"))
        Me.AntennaRedundantDiagonalBoltSize = DBtoNullableDbl(data.Item("AntennaRedundantDiagonalBoltSize"))
        Me.AntennaRedundantDiagonalNumBolts = DBtoNullableInt(data.Item("AntennaRedundantDiagonalNumBolts"))
        Me.AntennaRedundantDiagonalBoltEdgeDistance = DBtoNullableDbl(data.Item("AntennaRedundantDiagonalBoltEdgeDistance"))
        Me.AntennaRedundantDiagonalGageG1Distance = DBtoNullableDbl(data.Item("AntennaRedundantDiagonalGageG1Distance"))
        Me.AntennaRedundantDiagonalNetWidthDeduct = DBtoNullableDbl(data.Item("AntennaRedundantDiagonalNetWidthDeduct"))
        Me.AntennaRedundantDiagonalUFactor = DBtoNullableDbl(data.Item("AntennaRedundantDiagonalUFactor"))
        Me.AntennaRedundantSubDiagonalBoltGrade = DBtoStr(data.Item("AntennaRedundantSubDiagonalBoltGrade"))
        Me.AntennaRedundantSubDiagonalBoltSize = DBtoNullableDbl(data.Item("AntennaRedundantSubDiagonalBoltSize"))
        Me.AntennaRedundantSubDiagonalNumBolts = DBtoNullableInt(data.Item("AntennaRedundantSubDiagonalNumBolts"))
        Me.AntennaRedundantSubDiagonalBoltEdgeDistance = DBtoNullableDbl(data.Item("AntennaRedundantSubDiagonalBoltEdgeDistance"))
        Me.AntennaRedundantSubDiagonalGageG1Distance = DBtoNullableDbl(data.Item("AntennaRedundantSubDiagonalGageG1Distance"))
        Me.AntennaRedundantSubDiagonalNetWidthDeduct = DBtoNullableDbl(data.Item("AntennaRedundantSubDiagonalNetWidthDeduct"))
        Me.AntennaRedundantSubDiagonalUFactor = DBtoNullableDbl(data.Item("AntennaRedundantSubDiagonalUFactor"))
        Me.AntennaRedundantSubHorizontalBoltGrade = DBtoStr(data.Item("AntennaRedundantSubHorizontalBoltGrade"))
        Me.AntennaRedundantSubHorizontalBoltSize = DBtoNullableDbl(data.Item("AntennaRedundantSubHorizontalBoltSize"))
        Me.AntennaRedundantSubHorizontalNumBolts = DBtoNullableInt(data.Item("AntennaRedundantSubHorizontalNumBolts"))
        Me.AntennaRedundantSubHorizontalBoltEdgeDistance = DBtoNullableDbl(data.Item("AntennaRedundantSubHorizontalBoltEdgeDistance"))
        Me.AntennaRedundantSubHorizontalGageG1Distance = DBtoNullableDbl(data.Item("AntennaRedundantSubHorizontalGageG1Distance"))
        Me.AntennaRedundantSubHorizontalNetWidthDeduct = DBtoNullableDbl(data.Item("AntennaRedundantSubHorizontalNetWidthDeduct"))
        Me.AntennaRedundantSubHorizontalUFactor = DBtoNullableDbl(data.Item("AntennaRedundantSubHorizontalUFactor"))
        Me.AntennaRedundantVerticalBoltGrade = DBtoStr(data.Item("AntennaRedundantVerticalBoltGrade"))
        Me.AntennaRedundantVerticalBoltSize = DBtoNullableDbl(data.Item("AntennaRedundantVerticalBoltSize"))
        Me.AntennaRedundantVerticalNumBolts = DBtoNullableInt(data.Item("AntennaRedundantVerticalNumBolts"))
        Me.AntennaRedundantVerticalBoltEdgeDistance = DBtoNullableDbl(data.Item("AntennaRedundantVerticalBoltEdgeDistance"))
        Me.AntennaRedundantVerticalGageG1Distance = DBtoNullableDbl(data.Item("AntennaRedundantVerticalGageG1Distance"))
        Me.AntennaRedundantVerticalNetWidthDeduct = DBtoNullableDbl(data.Item("AntennaRedundantVerticalNetWidthDeduct"))
        Me.AntennaRedundantVerticalUFactor = DBtoNullableDbl(data.Item("AntennaRedundantVerticalUFactor"))
        Me.AntennaRedundantHipBoltGrade = DBtoStr(data.Item("AntennaRedundantHipBoltGrade"))
        Me.AntennaRedundantHipBoltSize = DBtoNullableDbl(data.Item("AntennaRedundantHipBoltSize"))
        Me.AntennaRedundantHipNumBolts = DBtoNullableInt(data.Item("AntennaRedundantHipNumBolts"))
        Me.AntennaRedundantHipBoltEdgeDistance = DBtoNullableDbl(data.Item("AntennaRedundantHipBoltEdgeDistance"))
        Me.AntennaRedundantHipGageG1Distance = DBtoNullableDbl(data.Item("AntennaRedundantHipGageG1Distance"))
        Me.AntennaRedundantHipNetWidthDeduct = DBtoNullableDbl(data.Item("AntennaRedundantHipNetWidthDeduct"))
        Me.AntennaRedundantHipUFactor = DBtoNullableDbl(data.Item("AntennaRedundantHipUFactor"))
        Me.AntennaRedundantHipDiagonalBoltGrade = DBtoStr(data.Item("AntennaRedundantHipDiagonalBoltGrade"))
        Me.AntennaRedundantHipDiagonalBoltSize = DBtoNullableDbl(data.Item("AntennaRedundantHipDiagonalBoltSize"))
        Me.AntennaRedundantHipDiagonalNumBolts = DBtoNullableInt(data.Item("AntennaRedundantHipDiagonalNumBolts"))
        Me.AntennaRedundantHipDiagonalBoltEdgeDistance = DBtoNullableDbl(data.Item("AntennaRedundantHipDiagonalBoltEdgeDistance"))
        Me.AntennaRedundantHipDiagonalGageG1Distance = DBtoNullableDbl(data.Item("AntennaRedundantHipDiagonalGageG1Distance"))
        Me.AntennaRedundantHipDiagonalNetWidthDeduct = DBtoNullableDbl(data.Item("AntennaRedundantHipDiagonalNetWidthDeduct"))
        Me.AntennaRedundantHipDiagonalUFactor = DBtoNullableDbl(data.Item("AntennaRedundantHipDiagonalUFactor"))
        Me.AntennaDiagonalOutOfPlaneRestraint = DBtoNullableBool(data.Item("AntennaDiagonalOutOfPlaneRestraint"))
        Me.AntennaTopGirtOutOfPlaneRestraint = DBtoNullableBool(data.Item("AntennaTopGirtOutOfPlaneRestraint"))
        Me.AntennaBottomGirtOutOfPlaneRestraint = DBtoNullableBool(data.Item("AntennaBottomGirtOutOfPlaneRestraint"))
        Me.AntennaMidGirtOutOfPlaneRestraint = DBtoNullableBool(data.Item("AntennaMidGirtOutOfPlaneRestraint"))
        Me.AntennaHorizontalOutOfPlaneRestraint = DBtoNullableBool(data.Item("AntennaHorizontalOutOfPlaneRestraint"))
        Me.AntennaSecondaryHorizontalOutOfPlaneRestraint = DBtoNullableBool(data.Item("AntennaSecondaryHorizontalOutOfPlaneRestraint"))
        Me.AntennaDiagOffsetNEY = DBtoNullableDbl(data.Item("AntennaDiagOffsetNEY"))
        Me.AntennaDiagOffsetNEX = DBtoNullableDbl(data.Item("AntennaDiagOffsetNEX"))
        Me.AntennaDiagOffsetPEY = DBtoNullableDbl(data.Item("AntennaDiagOffsetPEY"))
        Me.AntennaDiagOffsetPEX = DBtoNullableDbl(data.Item("AntennaDiagOffsetPEX"))
        Me.AntennaKbraceOffsetNEY = DBtoNullableDbl(data.Item("AntennaKbraceOffsetNEY"))
        Me.AntennaKbraceOffsetNEX = DBtoNullableDbl(data.Item("AntennaKbraceOffsetNEX"))
        Me.AntennaKbraceOffsetPEY = DBtoNullableDbl(data.Item("AntennaKbraceOffsetPEY"))
        Me.AntennaKbraceOffsetPEX = DBtoNullableDbl(data.Item("AntennaKbraceOffsetPEX"))

    End Sub
#End Region

    Public Function GenerateSQL() As String
        Dim insertString As String = ""

        insertString = insertString.AddtoDBString(Me.Rec.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaBraceType.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaHeight.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaDiagonalSpacing.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaDiagonalSpacingEx.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaNumSections.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaNumSesctions.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaSectionLength.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaLegType.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaLegSize.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaLegGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaLegMatlGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaDiagonalGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaDiagonalMatlGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaInnerBracingGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaInnerBracingMatlGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaTopGirtGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaTopGirtMatlGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaBotGirtGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaBotGirtMatlGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaInnerGirtGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaInnerGirtMatlGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaLongHorizontalGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaLongHorizontalMatlGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaShortHorizontalGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaShortHorizontalMatlGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaDiagonalType.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaDiagonalSize.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaInnerBracingType.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaInnerBracingSize.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaTopGirtType.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaTopGirtSize.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaBotGirtType.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaBotGirtSize.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaTopGirtOffset.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaBotGirtOffset.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaHasKBraceEndPanels.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaHasHorizontals.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaLongHorizontalType.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaLongHorizontalSize.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaShortHorizontalType.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaShortHorizontalSize.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantMatlGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantType.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantDiagType.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantSubDiagonalType.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantSubHorizontalType.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantVerticalType.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHipType.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHipDiagonalType.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHorizontalSize.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHorizontalSize2.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHorizontalSize3.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHorizontalSize4.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantDiagonalSize.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantDiagonalSize2.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantDiagonalSize3.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantDiagonalSize4.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantSubHorizontalSize.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantSubDiagonalSize.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaSubDiagLocation.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantVerticalSize.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHipDiagonalSize.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHipDiagonalSize2.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHipDiagonalSize3.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHipDiagonalSize4.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHipSize.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHipSize2.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHipSize3.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHipSize4.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaNumInnerGirts.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaInnerGirtType.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaInnerGirtSize.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaPoleShapeType.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaPoleSize.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaPoleGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaPoleMatlGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaPoleSpliceLength.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaTaperPoleNumSides.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaTaperPoleTopDiameter.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaTaperPoleBotDiameter.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaTaperPoleWallThickness.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaTaperPoleBendRadius.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaTaperPoleGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaTaperPoleMatlGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaSWMult.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaWPMult.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaAutoCalcKSingleAngle.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaAutoCalcKSolidRound.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaAfGusset.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaTfGusset.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaGussetBoltEdgeDistance.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaGussetGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaGussetMatlGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaAfMult.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaArMult.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaFlatIPAPole.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRoundIPAPole.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaFlatIPALeg.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRoundIPALeg.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaFlatIPAHorizontal.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRoundIPAHorizontal.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaFlatIPADiagonal.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRoundIPADiagonal.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaCSA_S37_SpeedUpFactor.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaKLegs.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaKXBracedDiags.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaKKBracedDiags.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaKZBracedDiags.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaKHorzs.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaKSecHorzs.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaKGirts.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaKInners.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaKXBracedDiagsY.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaKKBracedDiagsY.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaKZBracedDiagsY.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaKHorzsY.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaKSecHorzsY.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaKGirtsY.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaKInnersY.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaKRedHorz.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaKRedDiag.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaKRedSubDiag.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaKRedSubHorz.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaKRedVert.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaKRedHip.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaKRedHipDiag.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaKTLX.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaKTLZ.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaKTLLeg.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaInnerKTLX.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaInnerKTLZ.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaInnerKTLLeg.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaStitchBoltLocationHoriz.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaStitchBoltLocationDiag.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaStitchSpacing.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaStitchSpacingHorz.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaStitchSpacingDiag.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaStitchSpacingRed.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaLegNetWidthDeduct.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaLegUFactor.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaDiagonalNetWidthDeduct.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaTopGirtNetWidthDeduct.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaBotGirtNetWidthDeduct.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaInnerGirtNetWidthDeduct.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaHorizontalNetWidthDeduct.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaShortHorizontalNetWidthDeduct.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaDiagonalUFactor.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaTopGirtUFactor.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaBotGirtUFactor.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaInnerGirtUFactor.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaHorizontalUFactor.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaShortHorizontalUFactor.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaLegConnType.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaLegNumBolts.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaDiagonalNumBolts.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaTopGirtNumBolts.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaBotGirtNumBolts.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaInnerGirtNumBolts.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaHorizontalNumBolts.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaShortHorizontalNumBolts.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaLegBoltGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaLegBoltSize.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaDiagonalBoltGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaDiagonalBoltSize.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaTopGirtBoltGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaTopGirtBoltSize.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaBotGirtBoltGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaBotGirtBoltSize.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaInnerGirtBoltGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaInnerGirtBoltSize.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaHorizontalBoltGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaHorizontalBoltSize.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaShortHorizontalBoltGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaShortHorizontalBoltSize.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaLegBoltEdgeDistance.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaDiagonalBoltEdgeDistance.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaTopGirtBoltEdgeDistance.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaBotGirtBoltEdgeDistance.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaInnerGirtBoltEdgeDistance.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaHorizontalBoltEdgeDistance.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaShortHorizontalBoltEdgeDistance.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaDiagonalGageG1Distance.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaTopGirtGageG1Distance.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaBotGirtGageG1Distance.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaInnerGirtGageG1Distance.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaHorizontalGageG1Distance.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaShortHorizontalGageG1Distance.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHorizontalBoltGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHorizontalBoltSize.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHorizontalNumBolts.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHorizontalBoltEdgeDistance.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHorizontalGageG1Distance.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHorizontalNetWidthDeduct.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHorizontalUFactor.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantDiagonalBoltGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantDiagonalBoltSize.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantDiagonalNumBolts.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantDiagonalBoltEdgeDistance.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantDiagonalGageG1Distance.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantDiagonalNetWidthDeduct.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantDiagonalUFactor.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantSubDiagonalBoltGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantSubDiagonalBoltSize.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantSubDiagonalNumBolts.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantSubDiagonalBoltEdgeDistance.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantSubDiagonalGageG1Distance.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantSubDiagonalNetWidthDeduct.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantSubDiagonalUFactor.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantSubHorizontalBoltGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantSubHorizontalBoltSize.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantSubHorizontalNumBolts.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantSubHorizontalBoltEdgeDistance.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantSubHorizontalGageG1Distance.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantSubHorizontalNetWidthDeduct.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantSubHorizontalUFactor.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantVerticalBoltGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantVerticalBoltSize.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantVerticalNumBolts.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantVerticalBoltEdgeDistance.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantVerticalGageG1Distance.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantVerticalNetWidthDeduct.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantVerticalUFactor.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHipBoltGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHipBoltSize.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHipNumBolts.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHipBoltEdgeDistance.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHipGageG1Distance.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHipNetWidthDeduct.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHipUFactor.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHipDiagonalBoltGrade.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHipDiagonalBoltSize.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHipDiagonalNumBolts.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHipDiagonalBoltEdgeDistance.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHipDiagonalGageG1Distance.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHipDiagonalNetWidthDeduct.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaRedundantHipDiagonalUFactor.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaDiagonalOutOfPlaneRestraint.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaTopGirtOutOfPlaneRestraint.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaBottomGirtOutOfPlaneRestraint.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaMidGirtOutOfPlaneRestraint.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaHorizontalOutOfPlaneRestraint.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaSecondaryHorizontalOutOfPlaneRestraint.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaDiagOffsetNEY.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaDiagOffsetNEX.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaDiagOffsetPEY.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaDiagOffsetPEX.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaKbraceOffsetNEY.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaKbraceOffsetNEX.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaKbraceOffsetPEY.ToString)
        insertString = insertString.AddtoDBString(Me.AntennaKbraceOffsetPEX.ToString)

        Return insertString
    End Function

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As tnxAntennaRecord = TryCast(other, tnxAntennaRecord)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.Rec.CheckChange(otherToCompare.Rec, changes, categoryName, "Antenna Rec"), Equals, False)
        Equals = If(Me.AntennaBraceType.CheckChange(otherToCompare.AntennaBraceType, changes, categoryName, "Antenna Brace Type"), Equals, False)
        Equals = If(Me.AntennaHeight.CheckChange(otherToCompare.AntennaHeight, changes, categoryName, "Antenna Height"), Equals, False)
        Equals = If(Me.AntennaDiagonalSpacing.CheckChange(otherToCompare.AntennaDiagonalSpacing, changes, categoryName, "Antenna Diagonal Spacing"), Equals, False)
        Equals = If(Me.AntennaDiagonalSpacingEx.CheckChange(otherToCompare.AntennaDiagonalSpacingEx, changes, categoryName, "Antenna Diagonal Spacing Ex"), Equals, False)
        Equals = If(Me.AntennaNumSections.CheckChange(otherToCompare.AntennaNumSections, changes, categoryName, "Antenna Num Sections"), Equals, False)
        Equals = If(Me.AntennaNumSesctions.CheckChange(otherToCompare.AntennaNumSesctions, changes, categoryName, "Antenna Num Sesctions"), Equals, False)
        Equals = If(Me.AntennaSectionLength.CheckChange(otherToCompare.AntennaSectionLength, changes, categoryName, "Antenna Section Length"), Equals, False)
        Equals = If(Me.AntennaLegType.CheckChange(otherToCompare.AntennaLegType, changes, categoryName, "Antenna Leg Type"), Equals, False)
        Equals = If(Me.AntennaLegSize.CheckChange(otherToCompare.AntennaLegSize, changes, categoryName, "Antenna Leg Size"), Equals, False)
        Equals = If(Me.AntennaLegGrade.CheckChange(otherToCompare.AntennaLegGrade, changes, categoryName, "Antenna Leg Grade"), Equals, False)
        Equals = If(Me.AntennaLegMatlGrade.CheckChange(otherToCompare.AntennaLegMatlGrade, changes, categoryName, "Antenna Leg Matl Grade"), Equals, False)
        Equals = If(Me.AntennaDiagonalGrade.CheckChange(otherToCompare.AntennaDiagonalGrade, changes, categoryName, "Antenna Diagonal Grade"), Equals, False)
        Equals = If(Me.AntennaDiagonalMatlGrade.CheckChange(otherToCompare.AntennaDiagonalMatlGrade, changes, categoryName, "Antenna Diagonal Matl Grade"), Equals, False)
        Equals = If(Me.AntennaInnerBracingGrade.CheckChange(otherToCompare.AntennaInnerBracingGrade, changes, categoryName, "Antenna Inner Bracing Grade"), Equals, False)
        Equals = If(Me.AntennaInnerBracingMatlGrade.CheckChange(otherToCompare.AntennaInnerBracingMatlGrade, changes, categoryName, "Antenna Inner Bracing Matl Grade"), Equals, False)
        Equals = If(Me.AntennaTopGirtGrade.CheckChange(otherToCompare.AntennaTopGirtGrade, changes, categoryName, "Antenna Top Girt Grade"), Equals, False)
        Equals = If(Me.AntennaTopGirtMatlGrade.CheckChange(otherToCompare.AntennaTopGirtMatlGrade, changes, categoryName, "Antenna Top Girt Matl Grade"), Equals, False)
        Equals = If(Me.AntennaBotGirtGrade.CheckChange(otherToCompare.AntennaBotGirtGrade, changes, categoryName, "Antenna Bot Girt Grade"), Equals, False)
        Equals = If(Me.AntennaBotGirtMatlGrade.CheckChange(otherToCompare.AntennaBotGirtMatlGrade, changes, categoryName, "Antenna Bot Girt Matl Grade"), Equals, False)
        Equals = If(Me.AntennaInnerGirtGrade.CheckChange(otherToCompare.AntennaInnerGirtGrade, changes, categoryName, "Antenna Inner Girt Grade"), Equals, False)
        Equals = If(Me.AntennaInnerGirtMatlGrade.CheckChange(otherToCompare.AntennaInnerGirtMatlGrade, changes, categoryName, "Antenna Inner Girt Matl Grade"), Equals, False)
        Equals = If(Me.AntennaLongHorizontalGrade.CheckChange(otherToCompare.AntennaLongHorizontalGrade, changes, categoryName, "Antenna Long Horizontal Grade"), Equals, False)
        Equals = If(Me.AntennaLongHorizontalMatlGrade.CheckChange(otherToCompare.AntennaLongHorizontalMatlGrade, changes, categoryName, "Antenna Long Horizontal Matl Grade"), Equals, False)
        Equals = If(Me.AntennaShortHorizontalGrade.CheckChange(otherToCompare.AntennaShortHorizontalGrade, changes, categoryName, "Antenna Short Horizontal Grade"), Equals, False)
        Equals = If(Me.AntennaShortHorizontalMatlGrade.CheckChange(otherToCompare.AntennaShortHorizontalMatlGrade, changes, categoryName, "Antenna Short Horizontal Matl Grade"), Equals, False)
        Equals = If(Me.AntennaDiagonalType.CheckChange(otherToCompare.AntennaDiagonalType, changes, categoryName, "Antenna Diagonal Type"), Equals, False)
        Equals = If(Me.AntennaDiagonalSize.CheckChange(otherToCompare.AntennaDiagonalSize, changes, categoryName, "Antenna Diagonal Size"), Equals, False)
        Equals = If(Me.AntennaInnerBracingType.CheckChange(otherToCompare.AntennaInnerBracingType, changes, categoryName, "Antenna Inner Bracing Type"), Equals, False)
        Equals = If(Me.AntennaInnerBracingSize.CheckChange(otherToCompare.AntennaInnerBracingSize, changes, categoryName, "Antenna Inner Bracing Size"), Equals, False)
        Equals = If(Me.AntennaTopGirtType.CheckChange(otherToCompare.AntennaTopGirtType, changes, categoryName, "Antenna Top Girt Type"), Equals, False)
        Equals = If(Me.AntennaTopGirtSize.CheckChange(otherToCompare.AntennaTopGirtSize, changes, categoryName, "Antenna Top Girt Size"), Equals, False)
        Equals = If(Me.AntennaBotGirtType.CheckChange(otherToCompare.AntennaBotGirtType, changes, categoryName, "Antenna Bot Girt Type"), Equals, False)
        Equals = If(Me.AntennaBotGirtSize.CheckChange(otherToCompare.AntennaBotGirtSize, changes, categoryName, "Antenna Bot Girt Size"), Equals, False)
        Equals = If(Me.AntennaTopGirtOffset.CheckChange(otherToCompare.AntennaTopGirtOffset, changes, categoryName, "Antenna Top Girt Offset"), Equals, False)
        Equals = If(Me.AntennaBotGirtOffset.CheckChange(otherToCompare.AntennaBotGirtOffset, changes, categoryName, "Antenna Bot Girt Offset"), Equals, False)
        Equals = If(Me.AntennaHasKBraceEndPanels.CheckChange(otherToCompare.AntennaHasKBraceEndPanels, changes, categoryName, "Antenna Has K-Brace End Panels"), Equals, False)
        Equals = If(Me.AntennaHasHorizontals.CheckChange(otherToCompare.AntennaHasHorizontals, changes, categoryName, "Antenna Has Horizontals"), Equals, False)
        Equals = If(Me.AntennaLongHorizontalType.CheckChange(otherToCompare.AntennaLongHorizontalType, changes, categoryName, "Antenna Long Horizontal Type"), Equals, False)
        Equals = If(Me.AntennaLongHorizontalSize.CheckChange(otherToCompare.AntennaLongHorizontalSize, changes, categoryName, "Antenna Long Horizontal Size"), Equals, False)
        Equals = If(Me.AntennaShortHorizontalType.CheckChange(otherToCompare.AntennaShortHorizontalType, changes, categoryName, "Antenna Short Horizontal Type"), Equals, False)
        Equals = If(Me.AntennaShortHorizontalSize.CheckChange(otherToCompare.AntennaShortHorizontalSize, changes, categoryName, "Antenna Short Horizontal Size"), Equals, False)
        Equals = If(Me.AntennaRedundantGrade.CheckChange(otherToCompare.AntennaRedundantGrade, changes, categoryName, "Antenna Redundant Grade"), Equals, False)
        Equals = If(Me.AntennaRedundantMatlGrade.CheckChange(otherToCompare.AntennaRedundantMatlGrade, changes, categoryName, "Antenna Redundant Matl Grade"), Equals, False)
        Equals = If(Me.AntennaRedundantType.CheckChange(otherToCompare.AntennaRedundantType, changes, categoryName, "Antenna Redundant Type"), Equals, False)
        Equals = If(Me.AntennaRedundantDiagType.CheckChange(otherToCompare.AntennaRedundantDiagType, changes, categoryName, "Antenna Redundant Diag Type"), Equals, False)
        Equals = If(Me.AntennaRedundantSubDiagonalType.CheckChange(otherToCompare.AntennaRedundantSubDiagonalType, changes, categoryName, "Antenna Redundant Sub Diagonal Type"), Equals, False)
        Equals = If(Me.AntennaRedundantSubHorizontalType.CheckChange(otherToCompare.AntennaRedundantSubHorizontalType, changes, categoryName, "Antenna Redundant Sub Horizontal Type"), Equals, False)
        Equals = If(Me.AntennaRedundantVerticalType.CheckChange(otherToCompare.AntennaRedundantVerticalType, changes, categoryName, "Antenna Redundant Vertical Type"), Equals, False)
        Equals = If(Me.AntennaRedundantHipType.CheckChange(otherToCompare.AntennaRedundantHipType, changes, categoryName, "Antenna Redundant Hip Type"), Equals, False)
        Equals = If(Me.AntennaRedundantHipDiagonalType.CheckChange(otherToCompare.AntennaRedundantHipDiagonalType, changes, categoryName, "Antenna Redundant Hip Diagonal Type"), Equals, False)
        Equals = If(Me.AntennaRedundantHorizontalSize.CheckChange(otherToCompare.AntennaRedundantHorizontalSize, changes, categoryName, "Antenna Redundant Horizontal Size"), Equals, False)
        Equals = If(Me.AntennaRedundantHorizontalSize2.CheckChange(otherToCompare.AntennaRedundantHorizontalSize2, changes, categoryName, "Antenna Redundant Horizontal Size 2"), Equals, False)
        Equals = If(Me.AntennaRedundantHorizontalSize3.CheckChange(otherToCompare.AntennaRedundantHorizontalSize3, changes, categoryName, "Antenna Redundant Horizontal Size 3"), Equals, False)
        Equals = If(Me.AntennaRedundantHorizontalSize4.CheckChange(otherToCompare.AntennaRedundantHorizontalSize4, changes, categoryName, "Antenna Redundant Horizontal Size 4"), Equals, False)
        Equals = If(Me.AntennaRedundantDiagonalSize.CheckChange(otherToCompare.AntennaRedundantDiagonalSize, changes, categoryName, "Antenna Redundant Diagonal Size"), Equals, False)
        Equals = If(Me.AntennaRedundantDiagonalSize2.CheckChange(otherToCompare.AntennaRedundantDiagonalSize2, changes, categoryName, "Antenna Redundant Diagonal Size 2"), Equals, False)
        Equals = If(Me.AntennaRedundantDiagonalSize3.CheckChange(otherToCompare.AntennaRedundantDiagonalSize3, changes, categoryName, "Antenna Redundant Diagonal Size 3"), Equals, False)
        Equals = If(Me.AntennaRedundantDiagonalSize4.CheckChange(otherToCompare.AntennaRedundantDiagonalSize4, changes, categoryName, "Antenna Redundant Diagonal Size 4"), Equals, False)
        Equals = If(Me.AntennaRedundantSubHorizontalSize.CheckChange(otherToCompare.AntennaRedundantSubHorizontalSize, changes, categoryName, "Antenna Redundant Sub Horizontal Size"), Equals, False)
        Equals = If(Me.AntennaRedundantSubDiagonalSize.CheckChange(otherToCompare.AntennaRedundantSubDiagonalSize, changes, categoryName, "Antenna Redundant Sub Diagonal Size"), Equals, False)
        Equals = If(Me.AntennaSubDiagLocation.CheckChange(otherToCompare.AntennaSubDiagLocation, changes, categoryName, "Antenna Sub Diag Location"), Equals, False)
        Equals = If(Me.AntennaRedundantVerticalSize.CheckChange(otherToCompare.AntennaRedundantVerticalSize, changes, categoryName, "Antenna Redundant Vertical Size"), Equals, False)
        Equals = If(Me.AntennaRedundantHipDiagonalSize.CheckChange(otherToCompare.AntennaRedundantHipDiagonalSize, changes, categoryName, "Antenna Redundant Hip Diagonal Size"), Equals, False)
        Equals = If(Me.AntennaRedundantHipDiagonalSize2.CheckChange(otherToCompare.AntennaRedundantHipDiagonalSize2, changes, categoryName, "Antenna Redundant Hip Diagonal Size 2"), Equals, False)
        Equals = If(Me.AntennaRedundantHipDiagonalSize3.CheckChange(otherToCompare.AntennaRedundantHipDiagonalSize3, changes, categoryName, "Antenna Redundant Hip Diagonal Size 3"), Equals, False)
        Equals = If(Me.AntennaRedundantHipDiagonalSize4.CheckChange(otherToCompare.AntennaRedundantHipDiagonalSize4, changes, categoryName, "Antenna Redundant Hip Diagonal Size 4"), Equals, False)
        Equals = If(Me.AntennaRedundantHipSize.CheckChange(otherToCompare.AntennaRedundantHipSize, changes, categoryName, "Antenna Redundant Hip Size"), Equals, False)
        Equals = If(Me.AntennaRedundantHipSize2.CheckChange(otherToCompare.AntennaRedundantHipSize2, changes, categoryName, "Antenna Redundant Hip Size 2"), Equals, False)
        Equals = If(Me.AntennaRedundantHipSize3.CheckChange(otherToCompare.AntennaRedundantHipSize3, changes, categoryName, "Antenna Redundant Hip Size 3"), Equals, False)
        Equals = If(Me.AntennaRedundantHipSize4.CheckChange(otherToCompare.AntennaRedundantHipSize4, changes, categoryName, "Antenna Redundant Hip Size 4"), Equals, False)
        Equals = If(Me.AntennaNumInnerGirts.CheckChange(otherToCompare.AntennaNumInnerGirts, changes, categoryName, "Antenna Num Inner Girts"), Equals, False)
        Equals = If(Me.AntennaInnerGirtType.CheckChange(otherToCompare.AntennaInnerGirtType, changes, categoryName, "Antenna Inner Girt Type"), Equals, False)
        Equals = If(Me.AntennaInnerGirtSize.CheckChange(otherToCompare.AntennaInnerGirtSize, changes, categoryName, "Antenna Inner Girt Size"), Equals, False)
        Equals = If(Me.AntennaPoleShapeType.CheckChange(otherToCompare.AntennaPoleShapeType, changes, categoryName, "Antenna Pole Shape Type"), Equals, False)
        Equals = If(Me.AntennaPoleSize.CheckChange(otherToCompare.AntennaPoleSize, changes, categoryName, "Antenna Pole Size"), Equals, False)
        Equals = If(Me.AntennaPoleGrade.CheckChange(otherToCompare.AntennaPoleGrade, changes, categoryName, "Antenna Pole Grade"), Equals, False)
        Equals = If(Me.AntennaPoleMatlGrade.CheckChange(otherToCompare.AntennaPoleMatlGrade, changes, categoryName, "Antenna Pole Matl Grade"), Equals, False)
        Equals = If(Me.AntennaPoleSpliceLength.CheckChange(otherToCompare.AntennaPoleSpliceLength, changes, categoryName, "Antenna Pole Splice Length"), Equals, False)
        Equals = If(Me.AntennaTaperPoleNumSides.CheckChange(otherToCompare.AntennaTaperPoleNumSides, changes, categoryName, "Antenna Taper Pole Num Sides"), Equals, False)
        Equals = If(Me.AntennaTaperPoleTopDiameter.CheckChange(otherToCompare.AntennaTaperPoleTopDiameter, changes, categoryName, "Antenna Taper Pole Top Diameter"), Equals, False)
        Equals = If(Me.AntennaTaperPoleBotDiameter.CheckChange(otherToCompare.AntennaTaperPoleBotDiameter, changes, categoryName, "Antenna Taper Pole Bot Diameter"), Equals, False)
        Equals = If(Me.AntennaTaperPoleWallThickness.CheckChange(otherToCompare.AntennaTaperPoleWallThickness, changes, categoryName, "Antenna Taper Pole Wall Thickness"), Equals, False)
        Equals = If(Me.AntennaTaperPoleBendRadius.CheckChange(otherToCompare.AntennaTaperPoleBendRadius, changes, categoryName, "Antenna Taper Pole Bend Radius"), Equals, False)
        Equals = If(Me.AntennaTaperPoleGrade.CheckChange(otherToCompare.AntennaTaperPoleGrade, changes, categoryName, "Antenna Taper Pole Grade"), Equals, False)
        Equals = If(Me.AntennaTaperPoleMatlGrade.CheckChange(otherToCompare.AntennaTaperPoleMatlGrade, changes, categoryName, "Antenna Taper Pole Matl Grade"), Equals, False)
        Equals = If(Me.AntennaSWMult.CheckChange(otherToCompare.AntennaSWMult, changes, categoryName, "Antenna SW Mult"), Equals, False)
        Equals = If(Me.AntennaWPMult.CheckChange(otherToCompare.AntennaWPMult, changes, categoryName, "Antenna WP Mult"), Equals, False)
        Equals = If(Me.AntennaAutoCalcKSingleAngle.CheckChange(otherToCompare.AntennaAutoCalcKSingleAngle, changes, categoryName, "Antenna Auto Calc K Single Angle"), Equals, False)
        Equals = If(Me.AntennaAutoCalcKSolidRound.CheckChange(otherToCompare.AntennaAutoCalcKSolidRound, changes, categoryName, "Antenna Auto Calc K Solid Round"), Equals, False)
        Equals = If(Me.AntennaAfGusset.CheckChange(otherToCompare.AntennaAfGusset, changes, categoryName, "Antenna Af Gusset"), Equals, False)
        Equals = If(Me.AntennaTfGusset.CheckChange(otherToCompare.AntennaTfGusset, changes, categoryName, "Antenna Tf Gusset"), Equals, False)
        Equals = If(Me.AntennaGussetBoltEdgeDistance.CheckChange(otherToCompare.AntennaGussetBoltEdgeDistance, changes, categoryName, "Antenna Gusset Bolt Edge Distance"), Equals, False)
        Equals = If(Me.AntennaGussetGrade.CheckChange(otherToCompare.AntennaGussetGrade, changes, categoryName, "Antenna Gusset Grade"), Equals, False)
        Equals = If(Me.AntennaGussetMatlGrade.CheckChange(otherToCompare.AntennaGussetMatlGrade, changes, categoryName, "Antenna Gusset Matl Grade"), Equals, False)
        Equals = If(Me.AntennaAfMult.CheckChange(otherToCompare.AntennaAfMult, changes, categoryName, "Antenna Af Mult"), Equals, False)
        Equals = If(Me.AntennaArMult.CheckChange(otherToCompare.AntennaArMult, changes, categoryName, "Antenna Ar Mult"), Equals, False)
        Equals = If(Me.AntennaFlatIPAPole.CheckChange(otherToCompare.AntennaFlatIPAPole, changes, categoryName, "Antenna Flat IPA Pole"), Equals, False)
        Equals = If(Me.AntennaRoundIPAPole.CheckChange(otherToCompare.AntennaRoundIPAPole, changes, categoryName, "Antenna Round IPA Pole"), Equals, False)
        Equals = If(Me.AntennaFlatIPALeg.CheckChange(otherToCompare.AntennaFlatIPALeg, changes, categoryName, "Antenna Flat IP ALeg"), Equals, False)
        Equals = If(Me.AntennaRoundIPALeg.CheckChange(otherToCompare.AntennaRoundIPALeg, changes, categoryName, "Antenna Round IPA Leg"), Equals, False)
        Equals = If(Me.AntennaFlatIPAHorizontal.CheckChange(otherToCompare.AntennaFlatIPAHorizontal, changes, categoryName, "Antenna Flat IPA Horizontal"), Equals, False)
        Equals = If(Me.AntennaRoundIPAHorizontal.CheckChange(otherToCompare.AntennaRoundIPAHorizontal, changes, categoryName, "Antenna Round IPA Horizontal"), Equals, False)
        Equals = If(Me.AntennaFlatIPADiagonal.CheckChange(otherToCompare.AntennaFlatIPADiagonal, changes, categoryName, "Antenna Flat IPA Diagonal"), Equals, False)
        Equals = If(Me.AntennaRoundIPADiagonal.CheckChange(otherToCompare.AntennaRoundIPADiagonal, changes, categoryName, "Antenna Round IPA Diagonal"), Equals, False)
        Equals = If(Me.AntennaCSA_S37_SpeedUpFactor.CheckChange(otherToCompare.AntennaCSA_S37_SpeedUpFactor, changes, categoryName, "Antenna CSA-S37 Speed Up Factor"), Equals, False)
        Equals = If(Me.AntennaKLegs.CheckChange(otherToCompare.AntennaKLegs, changes, categoryName, "Antenna KLegs"), Equals, False)
        Equals = If(Me.AntennaKXBracedDiags.CheckChange(otherToCompare.AntennaKXBracedDiags, changes, categoryName, "Antenna K X-Braced Diags"), Equals, False)
        Equals = If(Me.AntennaKKBracedDiags.CheckChange(otherToCompare.AntennaKKBracedDiags, changes, categoryName, "Antenna K K-Braced Diags"), Equals, False)
        Equals = If(Me.AntennaKZBracedDiags.CheckChange(otherToCompare.AntennaKZBracedDiags, changes, categoryName, "Antenna K Z-Braced Diags"), Equals, False)
        Equals = If(Me.AntennaKHorzs.CheckChange(otherToCompare.AntennaKHorzs, changes, categoryName, "Antenna K Horzs"), Equals, False)
        Equals = If(Me.AntennaKSecHorzs.CheckChange(otherToCompare.AntennaKSecHorzs, changes, categoryName, "Antenna K Sec Horzs"), Equals, False)
        Equals = If(Me.AntennaKGirts.CheckChange(otherToCompare.AntennaKGirts, changes, categoryName, "Antenna K Girts"), Equals, False)
        Equals = If(Me.AntennaKInners.CheckChange(otherToCompare.AntennaKInners, changes, categoryName, "Antenna K Inners"), Equals, False)
        Equals = If(Me.AntennaKXBracedDiagsY.CheckChange(otherToCompare.AntennaKXBracedDiagsY, changes, categoryName, "Antenna K X-Braced Diags Y"), Equals, False)
        Equals = If(Me.AntennaKKBracedDiagsY.CheckChange(otherToCompare.AntennaKKBracedDiagsY, changes, categoryName, "Antenna K K-Braced Diags Y"), Equals, False)
        Equals = If(Me.AntennaKZBracedDiagsY.CheckChange(otherToCompare.AntennaKZBracedDiagsY, changes, categoryName, "Antenna K Z-Braced Diags Y"), Equals, False)
        Equals = If(Me.AntennaKHorzsY.CheckChange(otherToCompare.AntennaKHorzsY, changes, categoryName, "Antenna K Horzs Y"), Equals, False)
        Equals = If(Me.AntennaKSecHorzsY.CheckChange(otherToCompare.AntennaKSecHorzsY, changes, categoryName, "Antenna K Sec Horzs Y"), Equals, False)
        Equals = If(Me.AntennaKGirtsY.CheckChange(otherToCompare.AntennaKGirtsY, changes, categoryName, "Antenna K Girts Y"), Equals, False)
        Equals = If(Me.AntennaKInnersY.CheckChange(otherToCompare.AntennaKInnersY, changes, categoryName, "Antenna K Inners Y"), Equals, False)
        Equals = If(Me.AntennaKRedHorz.CheckChange(otherToCompare.AntennaKRedHorz, changes, categoryName, "Antenna K Red Horz"), Equals, False)
        Equals = If(Me.AntennaKRedDiag.CheckChange(otherToCompare.AntennaKRedDiag, changes, categoryName, "Antenna K Red Diag"), Equals, False)
        Equals = If(Me.AntennaKRedSubDiag.CheckChange(otherToCompare.AntennaKRedSubDiag, changes, categoryName, "Antenna K Red Sub Diag"), Equals, False)
        Equals = If(Me.AntennaKRedSubHorz.CheckChange(otherToCompare.AntennaKRedSubHorz, changes, categoryName, "Antenna K Red Sub Horz"), Equals, False)
        Equals = If(Me.AntennaKRedVert.CheckChange(otherToCompare.AntennaKRedVert, changes, categoryName, "Antenna K Red Vert"), Equals, False)
        Equals = If(Me.AntennaKRedHip.CheckChange(otherToCompare.AntennaKRedHip, changes, categoryName, "Antenna K Red Hip"), Equals, False)
        Equals = If(Me.AntennaKRedHipDiag.CheckChange(otherToCompare.AntennaKRedHipDiag, changes, categoryName, "Antenna K Red Hip Diag"), Equals, False)
        Equals = If(Me.AntennaKTLX.CheckChange(otherToCompare.AntennaKTLX, changes, categoryName, "Antenna KTLX"), Equals, False)
        Equals = If(Me.AntennaKTLZ.CheckChange(otherToCompare.AntennaKTLZ, changes, categoryName, "Antenna KTLZ"), Equals, False)
        Equals = If(Me.AntennaKTLLeg.CheckChange(otherToCompare.AntennaKTLLeg, changes, categoryName, "Antenna KTL Leg"), Equals, False)
        Equals = If(Me.AntennaInnerKTLX.CheckChange(otherToCompare.AntennaInnerKTLX, changes, categoryName, "Antenna Inner KTLX"), Equals, False)
        Equals = If(Me.AntennaInnerKTLZ.CheckChange(otherToCompare.AntennaInnerKTLZ, changes, categoryName, "Antenna Inner KTLZ"), Equals, False)
        Equals = If(Me.AntennaInnerKTLLeg.CheckChange(otherToCompare.AntennaInnerKTLLeg, changes, categoryName, "Antenna Inner KTL Leg"), Equals, False)
        Equals = If(Me.AntennaStitchBoltLocationHoriz.CheckChange(otherToCompare.AntennaStitchBoltLocationHoriz, changes, categoryName, "Antenna Stitch Bolt Location Horiz"), Equals, False)
        Equals = If(Me.AntennaStitchBoltLocationDiag.CheckChange(otherToCompare.AntennaStitchBoltLocationDiag, changes, categoryName, "Antenna Stitch Bolt Location Diag"), Equals, False)
        Equals = If(Me.AntennaStitchSpacing.CheckChange(otherToCompare.AntennaStitchSpacing, changes, categoryName, "Antenna Stitch Spacing"), Equals, False)
        Equals = If(Me.AntennaStitchSpacingHorz.CheckChange(otherToCompare.AntennaStitchSpacingHorz, changes, categoryName, "Antenna Stitch Spacing Horz"), Equals, False)
        Equals = If(Me.AntennaStitchSpacingDiag.CheckChange(otherToCompare.AntennaStitchSpacingDiag, changes, categoryName, "Antenna Stitch Spacing Diag"), Equals, False)
        Equals = If(Me.AntennaStitchSpacingRed.CheckChange(otherToCompare.AntennaStitchSpacingRed, changes, categoryName, "Antenna Stitch Spacing Red"), Equals, False)
        Equals = If(Me.AntennaLegNetWidthDeduct.CheckChange(otherToCompare.AntennaLegNetWidthDeduct, changes, categoryName, "Antenna Leg Net Width Deduct"), Equals, False)
        Equals = If(Me.AntennaLegUFactor.CheckChange(otherToCompare.AntennaLegUFactor, changes, categoryName, "Antenna Leg U Factor"), Equals, False)
        Equals = If(Me.AntennaDiagonalNetWidthDeduct.CheckChange(otherToCompare.AntennaDiagonalNetWidthDeduct, changes, categoryName, "Antenna Diagonal Net Width Deduct"), Equals, False)
        Equals = If(Me.AntennaTopGirtNetWidthDeduct.CheckChange(otherToCompare.AntennaTopGirtNetWidthDeduct, changes, categoryName, "Antenna Top Girt Net Width Deduct"), Equals, False)
        Equals = If(Me.AntennaBotGirtNetWidthDeduct.CheckChange(otherToCompare.AntennaBotGirtNetWidthDeduct, changes, categoryName, "Antenna Bot Girt Net Width Deduct"), Equals, False)
        Equals = If(Me.AntennaInnerGirtNetWidthDeduct.CheckChange(otherToCompare.AntennaInnerGirtNetWidthDeduct, changes, categoryName, "Antenna Inner Girt Net Width Deduct"), Equals, False)
        Equals = If(Me.AntennaHorizontalNetWidthDeduct.CheckChange(otherToCompare.AntennaHorizontalNetWidthDeduct, changes, categoryName, "Antenna Horizontal Net Width Deduct"), Equals, False)
        Equals = If(Me.AntennaShortHorizontalNetWidthDeduct.CheckChange(otherToCompare.AntennaShortHorizontalNetWidthDeduct, changes, categoryName, "Antenna Short Horizontal Net Width Deduct"), Equals, False)
        Equals = If(Me.AntennaDiagonalUFactor.CheckChange(otherToCompare.AntennaDiagonalUFactor, changes, categoryName, "Antenna Diagonal U Factor"), Equals, False)
        Equals = If(Me.AntennaTopGirtUFactor.CheckChange(otherToCompare.AntennaTopGirtUFactor, changes, categoryName, "Antenna Top Girt U Factor"), Equals, False)
        Equals = If(Me.AntennaBotGirtUFactor.CheckChange(otherToCompare.AntennaBotGirtUFactor, changes, categoryName, "Antenna Bot Girt U Factor"), Equals, False)
        Equals = If(Me.AntennaInnerGirtUFactor.CheckChange(otherToCompare.AntennaInnerGirtUFactor, changes, categoryName, "Antenna Inner Girt U Factor"), Equals, False)
        Equals = If(Me.AntennaHorizontalUFactor.CheckChange(otherToCompare.AntennaHorizontalUFactor, changes, categoryName, "Antenna Horizontal U Factor"), Equals, False)
        Equals = If(Me.AntennaShortHorizontalUFactor.CheckChange(otherToCompare.AntennaShortHorizontalUFactor, changes, categoryName, "Antenna Short Horizontal U Factor"), Equals, False)
        Equals = If(Me.AntennaLegConnType.CheckChange(otherToCompare.AntennaLegConnType, changes, categoryName, "Antenna Leg Conn Type"), Equals, False)
        Equals = If(Me.AntennaLegNumBolts.CheckChange(otherToCompare.AntennaLegNumBolts, changes, categoryName, "Antenna Leg Num Bolts"), Equals, False)
        Equals = If(Me.AntennaDiagonalNumBolts.CheckChange(otherToCompare.AntennaDiagonalNumBolts, changes, categoryName, "Antenna Diagonal Num Bolts"), Equals, False)
        Equals = If(Me.AntennaTopGirtNumBolts.CheckChange(otherToCompare.AntennaTopGirtNumBolts, changes, categoryName, "Antenna Top Girt Num Bolts"), Equals, False)
        Equals = If(Me.AntennaBotGirtNumBolts.CheckChange(otherToCompare.AntennaBotGirtNumBolts, changes, categoryName, "Antenna Bot Girt Num Bolts"), Equals, False)
        Equals = If(Me.AntennaInnerGirtNumBolts.CheckChange(otherToCompare.AntennaInnerGirtNumBolts, changes, categoryName, "Antenna Inner Girt Num Bolts"), Equals, False)
        Equals = If(Me.AntennaHorizontalNumBolts.CheckChange(otherToCompare.AntennaHorizontalNumBolts, changes, categoryName, "Antenna Horizontal Num Bolts"), Equals, False)
        Equals = If(Me.AntennaShortHorizontalNumBolts.CheckChange(otherToCompare.AntennaShortHorizontalNumBolts, changes, categoryName, "Antenna Short Horizontal Num Bolts"), Equals, False)
        Equals = If(Me.AntennaLegBoltGrade.CheckChange(otherToCompare.AntennaLegBoltGrade, changes, categoryName, "Antenna Leg Bolt Grade"), Equals, False)
        Equals = If(Me.AntennaLegBoltSize.CheckChange(otherToCompare.AntennaLegBoltSize, changes, categoryName, "Antenna Leg Bolt Size"), Equals, False)
        Equals = If(Me.AntennaDiagonalBoltGrade.CheckChange(otherToCompare.AntennaDiagonalBoltGrade, changes, categoryName, "Antenna Diagonal Bolt Grade"), Equals, False)
        Equals = If(Me.AntennaDiagonalBoltSize.CheckChange(otherToCompare.AntennaDiagonalBoltSize, changes, categoryName, "Antenna Diagonal Bolt Size"), Equals, False)
        Equals = If(Me.AntennaTopGirtBoltGrade.CheckChange(otherToCompare.AntennaTopGirtBoltGrade, changes, categoryName, "Antenna Top Girt Bolt Grade"), Equals, False)
        Equals = If(Me.AntennaTopGirtBoltSize.CheckChange(otherToCompare.AntennaTopGirtBoltSize, changes, categoryName, "Antenna Top Girt Bolt Size"), Equals, False)
        Equals = If(Me.AntennaBotGirtBoltGrade.CheckChange(otherToCompare.AntennaBotGirtBoltGrade, changes, categoryName, "Antenna Bot Girt Bolt Grade"), Equals, False)
        Equals = If(Me.AntennaBotGirtBoltSize.CheckChange(otherToCompare.AntennaBotGirtBoltSize, changes, categoryName, "Antenna Bot Girt Bolt Size"), Equals, False)
        Equals = If(Me.AntennaInnerGirtBoltGrade.CheckChange(otherToCompare.AntennaInnerGirtBoltGrade, changes, categoryName, "Antenna Inner Girt Bolt Grade"), Equals, False)
        Equals = If(Me.AntennaInnerGirtBoltSize.CheckChange(otherToCompare.AntennaInnerGirtBoltSize, changes, categoryName, "Antenna Inner Girt Bolt Size"), Equals, False)
        Equals = If(Me.AntennaHorizontalBoltGrade.CheckChange(otherToCompare.AntennaHorizontalBoltGrade, changes, categoryName, "Antenna Horizontal Bolt Grade"), Equals, False)
        Equals = If(Me.AntennaHorizontalBoltSize.CheckChange(otherToCompare.AntennaHorizontalBoltSize, changes, categoryName, "Antenna Horizontal Bolt Size"), Equals, False)
        Equals = If(Me.AntennaShortHorizontalBoltGrade.CheckChange(otherToCompare.AntennaShortHorizontalBoltGrade, changes, categoryName, "Antenna Short Horizontal Bolt Grade"), Equals, False)
        Equals = If(Me.AntennaShortHorizontalBoltSize.CheckChange(otherToCompare.AntennaShortHorizontalBoltSize, changes, categoryName, "Antenna Short Horizontal Bolt Size"), Equals, False)
        Equals = If(Me.AntennaLegBoltEdgeDistance.CheckChange(otherToCompare.AntennaLegBoltEdgeDistance, changes, categoryName, "Antenna Leg Bolt Edge Distance"), Equals, False)
        Equals = If(Me.AntennaDiagonalBoltEdgeDistance.CheckChange(otherToCompare.AntennaDiagonalBoltEdgeDistance, changes, categoryName, "Antenna Diagonal Bolt Edge Distance"), Equals, False)
        Equals = If(Me.AntennaTopGirtBoltEdgeDistance.CheckChange(otherToCompare.AntennaTopGirtBoltEdgeDistance, changes, categoryName, "Antenna Top Girt Bolt Edge Distance"), Equals, False)
        Equals = If(Me.AntennaBotGirtBoltEdgeDistance.CheckChange(otherToCompare.AntennaBotGirtBoltEdgeDistance, changes, categoryName, "Antenna Bot Girt Bolt Edge Distance"), Equals, False)
        Equals = If(Me.AntennaInnerGirtBoltEdgeDistance.CheckChange(otherToCompare.AntennaInnerGirtBoltEdgeDistance, changes, categoryName, "Antenna Inner Girt Bolt Edge Distance"), Equals, False)
        Equals = If(Me.AntennaHorizontalBoltEdgeDistance.CheckChange(otherToCompare.AntennaHorizontalBoltEdgeDistance, changes, categoryName, "Antenna Horizontal Bolt Edge Distance"), Equals, False)
        Equals = If(Me.AntennaShortHorizontalBoltEdgeDistance.CheckChange(otherToCompare.AntennaShortHorizontalBoltEdgeDistance, changes, categoryName, "Antenna Short Horizontal Bolt Edge Distance"), Equals, False)
        Equals = If(Me.AntennaDiagonalGageG1Distance.CheckChange(otherToCompare.AntennaDiagonalGageG1Distance, changes, categoryName, "Antenna Diagonal Gage G1 Distance"), Equals, False)
        Equals = If(Me.AntennaTopGirtGageG1Distance.CheckChange(otherToCompare.AntennaTopGirtGageG1Distance, changes, categoryName, "Antenna Top Girt Gage G1 Distance"), Equals, False)
        Equals = If(Me.AntennaBotGirtGageG1Distance.CheckChange(otherToCompare.AntennaBotGirtGageG1Distance, changes, categoryName, "Antenna Bot Girt Gage G1 Distance"), Equals, False)
        Equals = If(Me.AntennaInnerGirtGageG1Distance.CheckChange(otherToCompare.AntennaInnerGirtGageG1Distance, changes, categoryName, "Antenna Inner Girt Gage G1 Distance"), Equals, False)
        Equals = If(Me.AntennaHorizontalGageG1Distance.CheckChange(otherToCompare.AntennaHorizontalGageG1Distance, changes, categoryName, "Antenna Horizontal Gage G1 Distance"), Equals, False)
        Equals = If(Me.AntennaShortHorizontalGageG1Distance.CheckChange(otherToCompare.AntennaShortHorizontalGageG1Distance, changes, categoryName, "Antenna Short Horizontal Gage G1 Distance"), Equals, False)
        Equals = If(Me.AntennaRedundantHorizontalBoltGrade.CheckChange(otherToCompare.AntennaRedundantHorizontalBoltGrade, changes, categoryName, "Antenna Redundant Horizontal Bolt Grade"), Equals, False)
        Equals = If(Me.AntennaRedundantHorizontalBoltSize.CheckChange(otherToCompare.AntennaRedundantHorizontalBoltSize, changes, categoryName, "Antenna Redundant Horizontal Bolt Size"), Equals, False)
        Equals = If(Me.AntennaRedundantHorizontalNumBolts.CheckChange(otherToCompare.AntennaRedundantHorizontalNumBolts, changes, categoryName, "Antenna Redundant Horizontal Num Bolts"), Equals, False)
        Equals = If(Me.AntennaRedundantHorizontalBoltEdgeDistance.CheckChange(otherToCompare.AntennaRedundantHorizontalBoltEdgeDistance, changes, categoryName, "Antenna Redundant Horizontal Bolt Edge Distance"), Equals, False)
        Equals = If(Me.AntennaRedundantHorizontalGageG1Distance.CheckChange(otherToCompare.AntennaRedundantHorizontalGageG1Distance, changes, categoryName, "Antenna Redundant Horizontal Gage G1 Distance"), Equals, False)
        Equals = If(Me.AntennaRedundantHorizontalNetWidthDeduct.CheckChange(otherToCompare.AntennaRedundantHorizontalNetWidthDeduct, changes, categoryName, "Antenna Redundant Horizontal Net Width Deduct"), Equals, False)
        Equals = If(Me.AntennaRedundantHorizontalUFactor.CheckChange(otherToCompare.AntennaRedundantHorizontalUFactor, changes, categoryName, "Antenna Redundant Horizontal UFactor"), Equals, False)
        Equals = If(Me.AntennaRedundantDiagonalBoltGrade.CheckChange(otherToCompare.AntennaRedundantDiagonalBoltGrade, changes, categoryName, "Antenna Redundant Diagonal Bolt Grade"), Equals, False)
        Equals = If(Me.AntennaRedundantDiagonalBoltSize.CheckChange(otherToCompare.AntennaRedundantDiagonalBoltSize, changes, categoryName, "Antenna Redundant Diagonal Bolt Size"), Equals, False)
        Equals = If(Me.AntennaRedundantDiagonalNumBolts.CheckChange(otherToCompare.AntennaRedundantDiagonalNumBolts, changes, categoryName, "Antenna Redundant Diagonal Num Bolts"), Equals, False)
        Equals = If(Me.AntennaRedundantDiagonalBoltEdgeDistance.CheckChange(otherToCompare.AntennaRedundantDiagonalBoltEdgeDistance, changes, categoryName, "Antenna Redundant Diagonal Bolt Edge Distance"), Equals, False)
        Equals = If(Me.AntennaRedundantDiagonalGageG1Distance.CheckChange(otherToCompare.AntennaRedundantDiagonalGageG1Distance, changes, categoryName, "Antenna Redundant Diagonal Gage G1 Distance"), Equals, False)
        Equals = If(Me.AntennaRedundantDiagonalNetWidthDeduct.CheckChange(otherToCompare.AntennaRedundantDiagonalNetWidthDeduct, changes, categoryName, "Antenna Redundant Diagonal Net Width Deduct"), Equals, False)
        Equals = If(Me.AntennaRedundantDiagonalUFactor.CheckChange(otherToCompare.AntennaRedundantDiagonalUFactor, changes, categoryName, "Antenna Redundant Diagonal UFactor"), Equals, False)
        Equals = If(Me.AntennaRedundantSubDiagonalBoltGrade.CheckChange(otherToCompare.AntennaRedundantSubDiagonalBoltGrade, changes, categoryName, "Antenna Redundant Sub Diagonal Bolt Grade"), Equals, False)
        Equals = If(Me.AntennaRedundantSubDiagonalBoltSize.CheckChange(otherToCompare.AntennaRedundantSubDiagonalBoltSize, changes, categoryName, "Antenna Redundant Sub Diagonal Bolt Size"), Equals, False)
        Equals = If(Me.AntennaRedundantSubDiagonalNumBolts.CheckChange(otherToCompare.AntennaRedundantSubDiagonalNumBolts, changes, categoryName, "Antenna Redundant Sub Diagonal Num Bolts"), Equals, False)
        Equals = If(Me.AntennaRedundantSubDiagonalBoltEdgeDistance.CheckChange(otherToCompare.AntennaRedundantSubDiagonalBoltEdgeDistance, changes, categoryName, "Antenna Redundant Sub Diagonal Bolt Edge Distance"), Equals, False)
        Equals = If(Me.AntennaRedundantSubDiagonalGageG1Distance.CheckChange(otherToCompare.AntennaRedundantSubDiagonalGageG1Distance, changes, categoryName, "Antenna Redundant Sub Diagonal Gage G1 Distance"), Equals, False)
        Equals = If(Me.AntennaRedundantSubDiagonalNetWidthDeduct.CheckChange(otherToCompare.AntennaRedundantSubDiagonalNetWidthDeduct, changes, categoryName, "Antenna Redundant Sub Diagonal Net Width Deduct"), Equals, False)
        Equals = If(Me.AntennaRedundantSubDiagonalUFactor.CheckChange(otherToCompare.AntennaRedundantSubDiagonalUFactor, changes, categoryName, "Antenna Redundant Sub Diagonal UFactor"), Equals, False)
        Equals = If(Me.AntennaRedundantSubHorizontalBoltGrade.CheckChange(otherToCompare.AntennaRedundantSubHorizontalBoltGrade, changes, categoryName, "Antenna Redundant Sub Horizontal Bolt Grade"), Equals, False)
        Equals = If(Me.AntennaRedundantSubHorizontalBoltSize.CheckChange(otherToCompare.AntennaRedundantSubHorizontalBoltSize, changes, categoryName, "Antenna Redundant Sub Horizontal Bolt Size"), Equals, False)
        Equals = If(Me.AntennaRedundantSubHorizontalNumBolts.CheckChange(otherToCompare.AntennaRedundantSubHorizontalNumBolts, changes, categoryName, "Antenna Redundant Sub Horizontal Num Bolts"), Equals, False)
        Equals = If(Me.AntennaRedundantSubHorizontalBoltEdgeDistance.CheckChange(otherToCompare.AntennaRedundantSubHorizontalBoltEdgeDistance, changes, categoryName, "Antenna Redundant Sub Horizontal Bolt Edge Distance"), Equals, False)
        Equals = If(Me.AntennaRedundantSubHorizontalGageG1Distance.CheckChange(otherToCompare.AntennaRedundantSubHorizontalGageG1Distance, changes, categoryName, "Antenna Redundant Sub Horizontal Gage G1 Distance"), Equals, False)
        Equals = If(Me.AntennaRedundantSubHorizontalNetWidthDeduct.CheckChange(otherToCompare.AntennaRedundantSubHorizontalNetWidthDeduct, changes, categoryName, "Antenna Redundant Sub Horizontal Net Width Deduct"), Equals, False)
        Equals = If(Me.AntennaRedundantSubHorizontalUFactor.CheckChange(otherToCompare.AntennaRedundantSubHorizontalUFactor, changes, categoryName, "Antenna Redundant Sub Horizontal UFactor"), Equals, False)
        Equals = If(Me.AntennaRedundantVerticalBoltGrade.CheckChange(otherToCompare.AntennaRedundantVerticalBoltGrade, changes, categoryName, "Antenna Redundant Vertical Bolt Grade"), Equals, False)
        Equals = If(Me.AntennaRedundantVerticalBoltSize.CheckChange(otherToCompare.AntennaRedundantVerticalBoltSize, changes, categoryName, "Antenna Redundant Vertical Bolt Size"), Equals, False)
        Equals = If(Me.AntennaRedundantVerticalNumBolts.CheckChange(otherToCompare.AntennaRedundantVerticalNumBolts, changes, categoryName, "Antenna Redundant Vertical Num Bolts"), Equals, False)
        Equals = If(Me.AntennaRedundantVerticalBoltEdgeDistance.CheckChange(otherToCompare.AntennaRedundantVerticalBoltEdgeDistance, changes, categoryName, "Antenna Redundant Vertical Bolt Edge Distance"), Equals, False)
        Equals = If(Me.AntennaRedundantVerticalGageG1Distance.CheckChange(otherToCompare.AntennaRedundantVerticalGageG1Distance, changes, categoryName, "Antenna Redundant Vertical Gage G1 Distance"), Equals, False)
        Equals = If(Me.AntennaRedundantVerticalNetWidthDeduct.CheckChange(otherToCompare.AntennaRedundantVerticalNetWidthDeduct, changes, categoryName, "Antenna Redundant Vertical Net Width Deduct"), Equals, False)
        Equals = If(Me.AntennaRedundantVerticalUFactor.CheckChange(otherToCompare.AntennaRedundantVerticalUFactor, changes, categoryName, "Antenna Redundant Vertical UFactor"), Equals, False)
        Equals = If(Me.AntennaRedundantHipBoltGrade.CheckChange(otherToCompare.AntennaRedundantHipBoltGrade, changes, categoryName, "Antenna Redundant Hip Bolt Grade"), Equals, False)
        Equals = If(Me.AntennaRedundantHipBoltSize.CheckChange(otherToCompare.AntennaRedundantHipBoltSize, changes, categoryName, "Antenna Redundant Hip Bolt Size"), Equals, False)
        Equals = If(Me.AntennaRedundantHipNumBolts.CheckChange(otherToCompare.AntennaRedundantHipNumBolts, changes, categoryName, "Antenna Redundant Hip Num Bolts"), Equals, False)
        Equals = If(Me.AntennaRedundantHipBoltEdgeDistance.CheckChange(otherToCompare.AntennaRedundantHipBoltEdgeDistance, changes, categoryName, "Antenna Redundant Hip Bolt Edge Distance"), Equals, False)
        Equals = If(Me.AntennaRedundantHipGageG1Distance.CheckChange(otherToCompare.AntennaRedundantHipGageG1Distance, changes, categoryName, "Antenna Redundant Hip Gage G1 Distance"), Equals, False)
        Equals = If(Me.AntennaRedundantHipNetWidthDeduct.CheckChange(otherToCompare.AntennaRedundantHipNetWidthDeduct, changes, categoryName, "Antenna Redundant Hip Net Width Deduct"), Equals, False)
        Equals = If(Me.AntennaRedundantHipUFactor.CheckChange(otherToCompare.AntennaRedundantHipUFactor, changes, categoryName, "Antenna Redundant Hip UFactor"), Equals, False)
        Equals = If(Me.AntennaRedundantHipDiagonalBoltGrade.CheckChange(otherToCompare.AntennaRedundantHipDiagonalBoltGrade, changes, categoryName, "Antenna Redundant Hip Diagonal Bolt Grade"), Equals, False)
        Equals = If(Me.AntennaRedundantHipDiagonalBoltSize.CheckChange(otherToCompare.AntennaRedundantHipDiagonalBoltSize, changes, categoryName, "Antenna Redundant Hip Diagonal Bolt Size"), Equals, False)
        Equals = If(Me.AntennaRedundantHipDiagonalNumBolts.CheckChange(otherToCompare.AntennaRedundantHipDiagonalNumBolts, changes, categoryName, "Antenna Redundant Hip Diagonal Num Bolts"), Equals, False)
        Equals = If(Me.AntennaRedundantHipDiagonalBoltEdgeDistance.CheckChange(otherToCompare.AntennaRedundantHipDiagonalBoltEdgeDistance, changes, categoryName, "Antenna Redundant Hip Diagonal Bolt Edge Distance"), Equals, False)
        Equals = If(Me.AntennaRedundantHipDiagonalGageG1Distance.CheckChange(otherToCompare.AntennaRedundantHipDiagonalGageG1Distance, changes, categoryName, "Antenna Redundant Hip Diagonal Gage G1 Distance"), Equals, False)
        Equals = If(Me.AntennaRedundantHipDiagonalNetWidthDeduct.CheckChange(otherToCompare.AntennaRedundantHipDiagonalNetWidthDeduct, changes, categoryName, "Antenna Redundant Hip Diagonal Net Width Deduct"), Equals, False)
        Equals = If(Me.AntennaRedundantHipDiagonalUFactor.CheckChange(otherToCompare.AntennaRedundantHipDiagonalUFactor, changes, categoryName, "Antenna Redundant Hip Diagonal UFactor"), Equals, False)
        Equals = If(Me.AntennaDiagonalOutOfPlaneRestraint.CheckChange(otherToCompare.AntennaDiagonalOutOfPlaneRestraint, changes, categoryName, "Antenna Diagonal Out Of Plane Restraint"), Equals, False)
        Equals = If(Me.AntennaTopGirtOutOfPlaneRestraint.CheckChange(otherToCompare.AntennaTopGirtOutOfPlaneRestraint, changes, categoryName, "Antenna Top Girt Out Of Plane Restraint"), Equals, False)
        Equals = If(Me.AntennaBottomGirtOutOfPlaneRestraint.CheckChange(otherToCompare.AntennaBottomGirtOutOfPlaneRestraint, changes, categoryName, "Antenna Bottom Girt Out Of Plane Restraint"), Equals, False)
        Equals = If(Me.AntennaMidGirtOutOfPlaneRestraint.CheckChange(otherToCompare.AntennaMidGirtOutOfPlaneRestraint, changes, categoryName, "Antenna Mid Girt Out Of Plane Restraint"), Equals, False)
        Equals = If(Me.AntennaHorizontalOutOfPlaneRestraint.CheckChange(otherToCompare.AntennaHorizontalOutOfPlaneRestraint, changes, categoryName, "Antenna Horizontal Out Of Plane Restraint"), Equals, False)
        Equals = If(Me.AntennaSecondaryHorizontalOutOfPlaneRestraint.CheckChange(otherToCompare.AntennaSecondaryHorizontalOutOfPlaneRestraint, changes, categoryName, "Antenna Secondary Horizontal Out Of Plane Restraint"), Equals, False)
        Equals = If(Me.AntennaDiagOffsetNEY.CheckChange(otherToCompare.AntennaDiagOffsetNEY, changes, categoryName, "Antenna Diag Offset NEY"), Equals, False)
        Equals = If(Me.AntennaDiagOffsetNEX.CheckChange(otherToCompare.AntennaDiagOffsetNEX, changes, categoryName, "Antenna Diag Offset NEX"), Equals, False)
        Equals = If(Me.AntennaDiagOffsetPEY.CheckChange(otherToCompare.AntennaDiagOffsetPEY, changes, categoryName, "Antenna Diag Offset PEY"), Equals, False)
        Equals = If(Me.AntennaDiagOffsetPEX.CheckChange(otherToCompare.AntennaDiagOffsetPEX, changes, categoryName, "Antenna Diag Offset PEX"), Equals, False)
        Equals = If(Me.AntennaKbraceOffsetNEY.CheckChange(otherToCompare.AntennaKbraceOffsetNEY, changes, categoryName, "Antenna K brace Offset NEY"), Equals, False)
        Equals = If(Me.AntennaKbraceOffsetNEX.CheckChange(otherToCompare.AntennaKbraceOffsetNEX, changes, categoryName, "Antenna K brace Offset NEX"), Equals, False)
        Equals = If(Me.AntennaKbraceOffsetPEY.CheckChange(otherToCompare.AntennaKbraceOffsetPEY, changes, categoryName, "Antenna K brace Offset PEY"), Equals, False)
        Equals = If(Me.AntennaKbraceOffsetPEX.CheckChange(otherToCompare.AntennaKbraceOffsetPEX, changes, categoryName, "Antenna K brace Offset PEX"), Equals, False)

        Return Equals
    End Function

End Class

<DataContractAttribute()>
<KnownType(GetType(tnxTowerRecord))>
Partial Public Class tnxTowerRecord
    Inherits tnxGeometryRec
    'base structure

#Region "Inheritted"

    Public Overrides ReadOnly Property EDSObjectName As String = "Base Structure Section " & Me.Rec.ToString
    Public Overrides ReadOnly Property EDSTableName As String = "tnx.base_structure_sections"

    Public Overrides Function SQLInsertValues() As String
        Return SQLInsertValues(Nothing)
    End Function

    Public Overloads Function SQLInsertValues(Optional ByVal ParentID As Integer? = Nothing) As String
        'For any EDSObject that has parent object we will need to overload the update property with a version that excepts the current version being updated.
        SQLInsertValues = ""

        SQLInsertValues = SQLInsertValues.AddtoDBString(If(ParentID Is Nothing, EDSStructure.SQLQueryIDVar(Me.EDSTableDepth - 1), ParentID.NullableToString.FormatDBValue))
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.Rec.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerDatabase.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerName.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerHeight.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerFaceWidth.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerNumSections.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerSectionLength.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerDiagonalSpacing.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerDiagonalSpacingEx.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerBraceType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerFaceBevel.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerTopGirtOffset.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerBotGirtOffset.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerHasKBraceEndPanels.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerHasHorizontals.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLegType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLegSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLegGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLegMatlGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerDiagonalGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerDiagonalMatlGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerInnerBracingGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerInnerBracingMatlGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerTopGirtGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerTopGirtMatlGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerBotGirtGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerBotGirtMatlGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerInnerGirtGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerInnerGirtMatlGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLongHorizontalGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLongHorizontalMatlGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerShortHorizontalGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerShortHorizontalMatlGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerDiagonalType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerDiagonalSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerInnerBracingType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerInnerBracingSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerTopGirtType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerTopGirtSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerBotGirtType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerBotGirtSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerNumInnerGirts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerInnerGirtType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerInnerGirtSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLongHorizontalType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLongHorizontalSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerShortHorizontalType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerShortHorizontalSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantMatlGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantDiagType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantSubDiagonalType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantSubHorizontalType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantVerticalType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHipType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHipDiagonalType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHorizontalSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHorizontalSize2.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHorizontalSize3.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHorizontalSize4.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantDiagonalSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantDiagonalSize2.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantDiagonalSize3.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantDiagonalSize4.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantSubHorizontalSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantSubDiagonalSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerSubDiagLocation.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantVerticalSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHipSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHipSize2.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHipSize3.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHipSize4.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHipDiagonalSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHipDiagonalSize2.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHipDiagonalSize3.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHipDiagonalSize4.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerSWMult.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerWPMult.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerAutoCalcKSingleAngle.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerAutoCalcKSolidRound.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerAfGusset.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerTfGusset.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerGussetBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerGussetGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerGussetMatlGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerAfMult.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerArMult.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerFlatIPAPole.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRoundIPAPole.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerFlatIPALeg.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRoundIPALeg.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerFlatIPAHorizontal.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRoundIPAHorizontal.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerFlatIPADiagonal.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRoundIPADiagonal.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerCSA_S37_SpeedUpFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerKLegs.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerKXBracedDiags.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerKKBracedDiags.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerKZBracedDiags.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerKHorzs.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerKSecHorzs.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerKGirts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerKInners.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerKXBracedDiagsY.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerKKBracedDiagsY.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerKZBracedDiagsY.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerKHorzsY.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerKSecHorzsY.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerKGirtsY.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerKInnersY.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerKRedHorz.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerKRedDiag.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerKRedSubDiag.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerKRedSubHorz.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerKRedVert.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerKRedHip.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerKRedHipDiag.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerKTLX.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerKTLZ.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerKTLLeg.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerInnerKTLX.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerInnerKTLZ.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerInnerKTLLeg.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerStitchBoltLocationHoriz.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerStitchBoltLocationDiag.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerStitchBoltLocationRed.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerStitchSpacing.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerStitchSpacingDiag.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerStitchSpacingHorz.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerStitchSpacingRed.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLegNetWidthDeduct.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLegUFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerDiagonalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerTopGirtNetWidthDeduct.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerBotGirtNetWidthDeduct.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerInnerGirtNetWidthDeduct.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerHorizontalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerShortHorizontalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerDiagonalUFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerTopGirtUFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerBotGirtUFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerInnerGirtUFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerHorizontalUFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerShortHorizontalUFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLegConnType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLegNumBolts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerDiagonalNumBolts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerTopGirtNumBolts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerBotGirtNumBolts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerInnerGirtNumBolts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerHorizontalNumBolts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerShortHorizontalNumBolts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLegBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLegBoltSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerDiagonalBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerDiagonalBoltSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerTopGirtBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerTopGirtBoltSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerBotGirtBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerBotGirtBoltSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerInnerGirtBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerInnerGirtBoltSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerHorizontalBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerHorizontalBoltSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerShortHorizontalBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerShortHorizontalBoltSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLegBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerDiagonalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerTopGirtBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerBotGirtBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerInnerGirtBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerHorizontalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerShortHorizontalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerDiagonalGageG1Distance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerTopGirtGageG1Distance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerBotGirtGageG1Distance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerInnerGirtGageG1Distance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerHorizontalGageG1Distance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerShortHorizontalGageG1Distance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHorizontalBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHorizontalBoltSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHorizontalNumBolts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHorizontalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHorizontalGageG1Distance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHorizontalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHorizontalUFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantDiagonalBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantDiagonalBoltSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantDiagonalNumBolts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantDiagonalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantDiagonalGageG1Distance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantDiagonalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantDiagonalUFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantSubDiagonalBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantSubDiagonalBoltSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantSubDiagonalNumBolts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantSubDiagonalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantSubDiagonalGageG1Distance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantSubDiagonalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantSubDiagonalUFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantSubHorizontalBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantSubHorizontalBoltSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantSubHorizontalNumBolts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantSubHorizontalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantSubHorizontalGageG1Distance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantSubHorizontalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantSubHorizontalUFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantVerticalBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantVerticalBoltSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantVerticalNumBolts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantVerticalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantVerticalGageG1Distance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantVerticalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantVerticalUFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHipBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHipBoltSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHipNumBolts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHipBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHipGageG1Distance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHipNetWidthDeduct.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHipUFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHipDiagonalBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHipDiagonalBoltSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHipDiagonalNumBolts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHipDiagonalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHipDiagonalGageG1Distance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHipDiagonalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerRedundantHipDiagonalUFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerDiagonalOutOfPlaneRestraint.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerTopGirtOutOfPlaneRestraint.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerBottomGirtOutOfPlaneRestraint.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerMidGirtOutOfPlaneRestraint.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerHorizontalOutOfPlaneRestraint.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerSecondaryHorizontalOutOfPlaneRestraint.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerUniqueFlag.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerDiagOffsetNEY.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerDiagOffsetNEX.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerDiagOffsetPEY.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerDiagOffsetPEX.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerKbraceOffsetNEY.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerKbraceOffsetNEX.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerKbraceOffsetPEY.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerKbraceOffsetPEX.NullableToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        SQLInsertFields = SQLInsertFields.AddtoDBString("tnx_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRec")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerDatabase")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerName")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerHeight")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerFaceWidth")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerNumSections")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerSectionLength")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerDiagonalSpacing")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerDiagonalSpacingEx")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerBraceType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerFaceBevel")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerTopGirtOffset")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerBotGirtOffset")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerHasKBraceEndPanels")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerHasHorizontals")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLegType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLegSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLegGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLegMatlGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerDiagonalGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerDiagonalMatlGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerInnerBracingGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerInnerBracingMatlGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerTopGirtGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerTopGirtMatlGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerBotGirtGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerBotGirtMatlGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerInnerGirtGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerInnerGirtMatlGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLongHorizontalGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLongHorizontalMatlGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerShortHorizontalGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerShortHorizontalMatlGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerDiagonalType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerDiagonalSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerInnerBracingType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerInnerBracingSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerTopGirtType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerTopGirtSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerBotGirtType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerBotGirtSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerNumInnerGirts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerInnerGirtType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerInnerGirtSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLongHorizontalType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLongHorizontalSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerShortHorizontalType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerShortHorizontalSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantMatlGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantDiagType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantSubDiagonalType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantSubHorizontalType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantVerticalType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHipType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHipDiagonalType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHorizontalSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHorizontalSize2")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHorizontalSize3")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHorizontalSize4")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantDiagonalSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantDiagonalSize2")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantDiagonalSize3")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantDiagonalSize4")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantSubHorizontalSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantSubDiagonalSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerSubDiagLocation")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantVerticalSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHipSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHipSize2")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHipSize3")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHipSize4")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHipDiagonalSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHipDiagonalSize2")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHipDiagonalSize3")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHipDiagonalSize4")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerSWMult")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerWPMult")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerAutoCalcKSingleAngle")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerAutoCalcKSolidRound")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerAfGusset")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerTfGusset")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerGussetBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerGussetGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerGussetMatlGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerAfMult")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerArMult")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerFlatIPAPole")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRoundIPAPole")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerFlatIPALeg")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRoundIPALeg")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerFlatIPAHorizontal")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRoundIPAHorizontal")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerFlatIPADiagonal")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRoundIPADiagonal")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerCSA_S37_SpeedUpFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerKLegs")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerKXBracedDiags")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerKKBracedDiags")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerKZBracedDiags")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerKHorzs")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerKSecHorzs")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerKGirts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerKInners")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerKXBracedDiagsY")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerKKBracedDiagsY")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerKZBracedDiagsY")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerKHorzsY")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerKSecHorzsY")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerKGirtsY")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerKInnersY")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerKRedHorz")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerKRedDiag")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerKRedSubDiag")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerKRedSubHorz")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerKRedVert")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerKRedHip")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerKRedHipDiag")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerKTLX")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerKTLZ")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerKTLLeg")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerInnerKTLX")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerInnerKTLZ")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerInnerKTLLeg")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerStitchBoltLocationHoriz")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerStitchBoltLocationDiag")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerStitchBoltLocationRed")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerStitchSpacing")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerStitchSpacingDiag")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerStitchSpacingHorz")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerStitchSpacingRed")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLegNetWidthDeduct")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLegUFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerDiagonalNetWidthDeduct")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerTopGirtNetWidthDeduct")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerBotGirtNetWidthDeduct")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerInnerGirtNetWidthDeduct")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerHorizontalNetWidthDeduct")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerShortHorizontalNetWidthDeduct")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerDiagonalUFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerTopGirtUFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerBotGirtUFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerInnerGirtUFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerHorizontalUFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerShortHorizontalUFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLegConnType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLegNumBolts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerDiagonalNumBolts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerTopGirtNumBolts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerBotGirtNumBolts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerInnerGirtNumBolts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerHorizontalNumBolts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerShortHorizontalNumBolts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLegBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLegBoltSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerDiagonalBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerDiagonalBoltSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerTopGirtBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerTopGirtBoltSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerBotGirtBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerBotGirtBoltSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerInnerGirtBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerInnerGirtBoltSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerHorizontalBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerHorizontalBoltSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerShortHorizontalBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerShortHorizontalBoltSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLegBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerDiagonalBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerTopGirtBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerBotGirtBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerInnerGirtBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerHorizontalBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerShortHorizontalBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerDiagonalGageG1Distance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerTopGirtGageG1Distance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerBotGirtGageG1Distance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerInnerGirtGageG1Distance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerHorizontalGageG1Distance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerShortHorizontalGageG1Distance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHorizontalBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHorizontalBoltSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHorizontalNumBolts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHorizontalBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHorizontalGageG1Distance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHorizontalNetWidthDeduct")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHorizontalUFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantDiagonalBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantDiagonalBoltSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantDiagonalNumBolts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantDiagonalBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantDiagonalGageG1Distance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantDiagonalNetWidthDeduct")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantDiagonalUFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantSubDiagonalBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantSubDiagonalBoltSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantSubDiagonalNumBolts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantSubDiagonalBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantSubDiagonalGageG1Distance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantSubDiagonalNetWidthDeduct")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantSubDiagonalUFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantSubHorizontalBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantSubHorizontalBoltSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantSubHorizontalNumBolts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantSubHorizontalBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantSubHorizontalGageG1Distance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantSubHorizontalNetWidthDeduct")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantSubHorizontalUFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantVerticalBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantVerticalBoltSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantVerticalNumBolts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantVerticalBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantVerticalGageG1Distance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantVerticalNetWidthDeduct")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantVerticalUFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHipBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHipBoltSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHipNumBolts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHipBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHipGageG1Distance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHipNetWidthDeduct")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHipUFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHipDiagonalBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHipDiagonalBoltSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHipDiagonalNumBolts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHipDiagonalBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHipDiagonalGageG1Distance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHipDiagonalNetWidthDeduct")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerRedundantHipDiagonalUFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerDiagonalOutOfPlaneRestraint")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerTopGirtOutOfPlaneRestraint")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerBottomGirtOutOfPlaneRestraint")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerMidGirtOutOfPlaneRestraint")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerHorizontalOutOfPlaneRestraint")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerSecondaryHorizontalOutOfPlaneRestraint")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerUniqueFlag")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerDiagOffsetNEY")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerDiagOffsetNEX")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerDiagOffsetPEY")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerDiagOffsetPEX")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerKbraceOffsetNEY")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerKbraceOffsetNEX")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerKbraceOffsetPEY")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerKbraceOffsetPEX")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""

        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRec = " & Me.Rec.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerDatabase = " & Me.TowerDatabase.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerName = " & Me.TowerName.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerHeight = " & Me.TowerHeight.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerFaceWidth = " & Me.TowerFaceWidth.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerNumSections = " & Me.TowerNumSections.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerSectionLength = " & Me.TowerSectionLength.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerDiagonalSpacing = " & Me.TowerDiagonalSpacing.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerDiagonalSpacingEx = " & Me.TowerDiagonalSpacingEx.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerBraceType = " & Me.TowerBraceType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerFaceBevel = " & Me.TowerFaceBevel.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerTopGirtOffset = " & Me.TowerTopGirtOffset.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerBotGirtOffset = " & Me.TowerBotGirtOffset.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerHasKBraceEndPanels = " & Me.TowerHasKBraceEndPanels.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerHasHorizontals = " & Me.TowerHasHorizontals.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerLegType = " & Me.TowerLegType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerLegSize = " & Me.TowerLegSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerLegGrade = " & Me.TowerLegGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerLegMatlGrade = " & Me.TowerLegMatlGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerDiagonalGrade = " & Me.TowerDiagonalGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerDiagonalMatlGrade = " & Me.TowerDiagonalMatlGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerInnerBracingGrade = " & Me.TowerInnerBracingGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerInnerBracingMatlGrade = " & Me.TowerInnerBracingMatlGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerTopGirtGrade = " & Me.TowerTopGirtGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerTopGirtMatlGrade = " & Me.TowerTopGirtMatlGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerBotGirtGrade = " & Me.TowerBotGirtGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerBotGirtMatlGrade = " & Me.TowerBotGirtMatlGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerInnerGirtGrade = " & Me.TowerInnerGirtGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerInnerGirtMatlGrade = " & Me.TowerInnerGirtMatlGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerLongHorizontalGrade = " & Me.TowerLongHorizontalGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerLongHorizontalMatlGrade = " & Me.TowerLongHorizontalMatlGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerShortHorizontalGrade = " & Me.TowerShortHorizontalGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerShortHorizontalMatlGrade = " & Me.TowerShortHorizontalMatlGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerDiagonalType = " & Me.TowerDiagonalType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerDiagonalSize = " & Me.TowerDiagonalSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerInnerBracingType = " & Me.TowerInnerBracingType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerInnerBracingSize = " & Me.TowerInnerBracingSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerTopGirtType = " & Me.TowerTopGirtType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerTopGirtSize = " & Me.TowerTopGirtSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerBotGirtType = " & Me.TowerBotGirtType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerBotGirtSize = " & Me.TowerBotGirtSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerNumInnerGirts = " & Me.TowerNumInnerGirts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerInnerGirtType = " & Me.TowerInnerGirtType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerInnerGirtSize = " & Me.TowerInnerGirtSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerLongHorizontalType = " & Me.TowerLongHorizontalType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerLongHorizontalSize = " & Me.TowerLongHorizontalSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerShortHorizontalType = " & Me.TowerShortHorizontalType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerShortHorizontalSize = " & Me.TowerShortHorizontalSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantGrade = " & Me.TowerRedundantGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantMatlGrade = " & Me.TowerRedundantMatlGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantType = " & Me.TowerRedundantType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantDiagType = " & Me.TowerRedundantDiagType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantSubDiagonalType = " & Me.TowerRedundantSubDiagonalType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantSubHorizontalType = " & Me.TowerRedundantSubHorizontalType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantVerticalType = " & Me.TowerRedundantVerticalType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHipType = " & Me.TowerRedundantHipType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHipDiagonalType = " & Me.TowerRedundantHipDiagonalType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHorizontalSize = " & Me.TowerRedundantHorizontalSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHorizontalSize2 = " & Me.TowerRedundantHorizontalSize2.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHorizontalSize3 = " & Me.TowerRedundantHorizontalSize3.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHorizontalSize4 = " & Me.TowerRedundantHorizontalSize4.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantDiagonalSize = " & Me.TowerRedundantDiagonalSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantDiagonalSize2 = " & Me.TowerRedundantDiagonalSize2.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantDiagonalSize3 = " & Me.TowerRedundantDiagonalSize3.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantDiagonalSize4 = " & Me.TowerRedundantDiagonalSize4.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantSubHorizontalSize = " & Me.TowerRedundantSubHorizontalSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantSubDiagonalSize = " & Me.TowerRedundantSubDiagonalSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerSubDiagLocation = " & Me.TowerSubDiagLocation.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantVerticalSize = " & Me.TowerRedundantVerticalSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHipSize = " & Me.TowerRedundantHipSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHipSize2 = " & Me.TowerRedundantHipSize2.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHipSize3 = " & Me.TowerRedundantHipSize3.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHipSize4 = " & Me.TowerRedundantHipSize4.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHipDiagonalSize = " & Me.TowerRedundantHipDiagonalSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHipDiagonalSize2 = " & Me.TowerRedundantHipDiagonalSize2.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHipDiagonalSize3 = " & Me.TowerRedundantHipDiagonalSize3.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHipDiagonalSize4 = " & Me.TowerRedundantHipDiagonalSize4.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerSWMult = " & Me.TowerSWMult.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerWPMult = " & Me.TowerWPMult.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerAutoCalcKSingleAngle = " & Me.TowerAutoCalcKSingleAngle.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerAutoCalcKSolidRound = " & Me.TowerAutoCalcKSolidRound.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerAfGusset = " & Me.TowerAfGusset.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerTfGusset = " & Me.TowerTfGusset.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerGussetBoltEdgeDistance = " & Me.TowerGussetBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerGussetGrade = " & Me.TowerGussetGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerGussetMatlGrade = " & Me.TowerGussetMatlGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerAfMult = " & Me.TowerAfMult.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerArMult = " & Me.TowerArMult.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerFlatIPAPole = " & Me.TowerFlatIPAPole.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRoundIPAPole = " & Me.TowerRoundIPAPole.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerFlatIPALeg = " & Me.TowerFlatIPALeg.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRoundIPALeg = " & Me.TowerRoundIPALeg.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerFlatIPAHorizontal = " & Me.TowerFlatIPAHorizontal.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRoundIPAHorizontal = " & Me.TowerRoundIPAHorizontal.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerFlatIPADiagonal = " & Me.TowerFlatIPADiagonal.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRoundIPADiagonal = " & Me.TowerRoundIPADiagonal.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerCSA_S37_SpeedUpFactor = " & Me.TowerCSA_S37_SpeedUpFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerKLegs = " & Me.TowerKLegs.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerKXBracedDiags = " & Me.TowerKXBracedDiags.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerKKBracedDiags = " & Me.TowerKKBracedDiags.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerKZBracedDiags = " & Me.TowerKZBracedDiags.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerKHorzs = " & Me.TowerKHorzs.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerKSecHorzs = " & Me.TowerKSecHorzs.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerKGirts = " & Me.TowerKGirts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerKInners = " & Me.TowerKInners.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerKXBracedDiagsY = " & Me.TowerKXBracedDiagsY.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerKKBracedDiagsY = " & Me.TowerKKBracedDiagsY.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerKZBracedDiagsY = " & Me.TowerKZBracedDiagsY.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerKHorzsY = " & Me.TowerKHorzsY.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerKSecHorzsY = " & Me.TowerKSecHorzsY.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerKGirtsY = " & Me.TowerKGirtsY.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerKInnersY = " & Me.TowerKInnersY.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerKRedHorz = " & Me.TowerKRedHorz.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerKRedDiag = " & Me.TowerKRedDiag.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerKRedSubDiag = " & Me.TowerKRedSubDiag.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerKRedSubHorz = " & Me.TowerKRedSubHorz.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerKRedVert = " & Me.TowerKRedVert.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerKRedHip = " & Me.TowerKRedHip.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerKRedHipDiag = " & Me.TowerKRedHipDiag.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerKTLX = " & Me.TowerKTLX.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerKTLZ = " & Me.TowerKTLZ.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerKTLLeg = " & Me.TowerKTLLeg.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerInnerKTLX = " & Me.TowerInnerKTLX.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerInnerKTLZ = " & Me.TowerInnerKTLZ.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerInnerKTLLeg = " & Me.TowerInnerKTLLeg.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerStitchBoltLocationHoriz = " & Me.TowerStitchBoltLocationHoriz.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerStitchBoltLocationDiag = " & Me.TowerStitchBoltLocationDiag.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerStitchBoltLocationRed = " & Me.TowerStitchBoltLocationRed.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerStitchSpacing = " & Me.TowerStitchSpacing.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerStitchSpacingDiag = " & Me.TowerStitchSpacingDiag.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerStitchSpacingHorz = " & Me.TowerStitchSpacingHorz.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerStitchSpacingRed = " & Me.TowerStitchSpacingRed.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerLegNetWidthDeduct = " & Me.TowerLegNetWidthDeduct.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerLegUFactor = " & Me.TowerLegUFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerDiagonalNetWidthDeduct = " & Me.TowerDiagonalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerTopGirtNetWidthDeduct = " & Me.TowerTopGirtNetWidthDeduct.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerBotGirtNetWidthDeduct = " & Me.TowerBotGirtNetWidthDeduct.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerInnerGirtNetWidthDeduct = " & Me.TowerInnerGirtNetWidthDeduct.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerHorizontalNetWidthDeduct = " & Me.TowerHorizontalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerShortHorizontalNetWidthDeduct = " & Me.TowerShortHorizontalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerDiagonalUFactor = " & Me.TowerDiagonalUFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerTopGirtUFactor = " & Me.TowerTopGirtUFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerBotGirtUFactor = " & Me.TowerBotGirtUFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerInnerGirtUFactor = " & Me.TowerInnerGirtUFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerHorizontalUFactor = " & Me.TowerHorizontalUFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerShortHorizontalUFactor = " & Me.TowerShortHorizontalUFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerLegConnType = " & Me.TowerLegConnType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerLegNumBolts = " & Me.TowerLegNumBolts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerDiagonalNumBolts = " & Me.TowerDiagonalNumBolts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerTopGirtNumBolts = " & Me.TowerTopGirtNumBolts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerBotGirtNumBolts = " & Me.TowerBotGirtNumBolts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerInnerGirtNumBolts = " & Me.TowerInnerGirtNumBolts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerHorizontalNumBolts = " & Me.TowerHorizontalNumBolts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerShortHorizontalNumBolts = " & Me.TowerShortHorizontalNumBolts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerLegBoltGrade = " & Me.TowerLegBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerLegBoltSize = " & Me.TowerLegBoltSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerDiagonalBoltGrade = " & Me.TowerDiagonalBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerDiagonalBoltSize = " & Me.TowerDiagonalBoltSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerTopGirtBoltGrade = " & Me.TowerTopGirtBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerTopGirtBoltSize = " & Me.TowerTopGirtBoltSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerBotGirtBoltGrade = " & Me.TowerBotGirtBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerBotGirtBoltSize = " & Me.TowerBotGirtBoltSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerInnerGirtBoltGrade = " & Me.TowerInnerGirtBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerInnerGirtBoltSize = " & Me.TowerInnerGirtBoltSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerHorizontalBoltGrade = " & Me.TowerHorizontalBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerHorizontalBoltSize = " & Me.TowerHorizontalBoltSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerShortHorizontalBoltGrade = " & Me.TowerShortHorizontalBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerShortHorizontalBoltSize = " & Me.TowerShortHorizontalBoltSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerLegBoltEdgeDistance = " & Me.TowerLegBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerDiagonalBoltEdgeDistance = " & Me.TowerDiagonalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerTopGirtBoltEdgeDistance = " & Me.TowerTopGirtBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerBotGirtBoltEdgeDistance = " & Me.TowerBotGirtBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerInnerGirtBoltEdgeDistance = " & Me.TowerInnerGirtBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerHorizontalBoltEdgeDistance = " & Me.TowerHorizontalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerShortHorizontalBoltEdgeDistance = " & Me.TowerShortHorizontalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerDiagonalGageG1Distance = " & Me.TowerDiagonalGageG1Distance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerTopGirtGageG1Distance = " & Me.TowerTopGirtGageG1Distance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerBotGirtGageG1Distance = " & Me.TowerBotGirtGageG1Distance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerInnerGirtGageG1Distance = " & Me.TowerInnerGirtGageG1Distance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerHorizontalGageG1Distance = " & Me.TowerHorizontalGageG1Distance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerShortHorizontalGageG1Distance = " & Me.TowerShortHorizontalGageG1Distance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHorizontalBoltGrade = " & Me.TowerRedundantHorizontalBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHorizontalBoltSize = " & Me.TowerRedundantHorizontalBoltSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHorizontalNumBolts = " & Me.TowerRedundantHorizontalNumBolts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHorizontalBoltEdgeDistance = " & Me.TowerRedundantHorizontalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHorizontalGageG1Distance = " & Me.TowerRedundantHorizontalGageG1Distance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHorizontalNetWidthDeduct = " & Me.TowerRedundantHorizontalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHorizontalUFactor = " & Me.TowerRedundantHorizontalUFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantDiagonalBoltGrade = " & Me.TowerRedundantDiagonalBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantDiagonalBoltSize = " & Me.TowerRedundantDiagonalBoltSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantDiagonalNumBolts = " & Me.TowerRedundantDiagonalNumBolts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantDiagonalBoltEdgeDistance = " & Me.TowerRedundantDiagonalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantDiagonalGageG1Distance = " & Me.TowerRedundantDiagonalGageG1Distance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantDiagonalNetWidthDeduct = " & Me.TowerRedundantDiagonalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantDiagonalUFactor = " & Me.TowerRedundantDiagonalUFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantSubDiagonalBoltGrade = " & Me.TowerRedundantSubDiagonalBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantSubDiagonalBoltSize = " & Me.TowerRedundantSubDiagonalBoltSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantSubDiagonalNumBolts = " & Me.TowerRedundantSubDiagonalNumBolts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantSubDiagonalBoltEdgeDistance = " & Me.TowerRedundantSubDiagonalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantSubDiagonalGageG1Distance = " & Me.TowerRedundantSubDiagonalGageG1Distance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantSubDiagonalNetWidthDeduct = " & Me.TowerRedundantSubDiagonalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantSubDiagonalUFactor = " & Me.TowerRedundantSubDiagonalUFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantSubHorizontalBoltGrade = " & Me.TowerRedundantSubHorizontalBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantSubHorizontalBoltSize = " & Me.TowerRedundantSubHorizontalBoltSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantSubHorizontalNumBolts = " & Me.TowerRedundantSubHorizontalNumBolts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantSubHorizontalBoltEdgeDistance = " & Me.TowerRedundantSubHorizontalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantSubHorizontalGageG1Distance = " & Me.TowerRedundantSubHorizontalGageG1Distance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantSubHorizontalNetWidthDeduct = " & Me.TowerRedundantSubHorizontalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantSubHorizontalUFactor = " & Me.TowerRedundantSubHorizontalUFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantVerticalBoltGrade = " & Me.TowerRedundantVerticalBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantVerticalBoltSize = " & Me.TowerRedundantVerticalBoltSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantVerticalNumBolts = " & Me.TowerRedundantVerticalNumBolts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantVerticalBoltEdgeDistance = " & Me.TowerRedundantVerticalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantVerticalGageG1Distance = " & Me.TowerRedundantVerticalGageG1Distance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantVerticalNetWidthDeduct = " & Me.TowerRedundantVerticalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantVerticalUFactor = " & Me.TowerRedundantVerticalUFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHipBoltGrade = " & Me.TowerRedundantHipBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHipBoltSize = " & Me.TowerRedundantHipBoltSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHipNumBolts = " & Me.TowerRedundantHipNumBolts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHipBoltEdgeDistance = " & Me.TowerRedundantHipBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHipGageG1Distance = " & Me.TowerRedundantHipGageG1Distance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHipNetWidthDeduct = " & Me.TowerRedundantHipNetWidthDeduct.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHipUFactor = " & Me.TowerRedundantHipUFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHipDiagonalBoltGrade = " & Me.TowerRedundantHipDiagonalBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHipDiagonalBoltSize = " & Me.TowerRedundantHipDiagonalBoltSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHipDiagonalNumBolts = " & Me.TowerRedundantHipDiagonalNumBolts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHipDiagonalBoltEdgeDistance = " & Me.TowerRedundantHipDiagonalBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHipDiagonalGageG1Distance = " & Me.TowerRedundantHipDiagonalGageG1Distance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHipDiagonalNetWidthDeduct = " & Me.TowerRedundantHipDiagonalNetWidthDeduct.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerRedundantHipDiagonalUFactor = " & Me.TowerRedundantHipDiagonalUFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerDiagonalOutOfPlaneRestraint = " & Me.TowerDiagonalOutOfPlaneRestraint.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerTopGirtOutOfPlaneRestraint = " & Me.TowerTopGirtOutOfPlaneRestraint.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerBottomGirtOutOfPlaneRestraint = " & Me.TowerBottomGirtOutOfPlaneRestraint.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerMidGirtOutOfPlaneRestraint = " & Me.TowerMidGirtOutOfPlaneRestraint.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerHorizontalOutOfPlaneRestraint = " & Me.TowerHorizontalOutOfPlaneRestraint.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerSecondaryHorizontalOutOfPlaneRestraint = " & Me.TowerSecondaryHorizontalOutOfPlaneRestraint.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerUniqueFlag = " & Me.TowerUniqueFlag.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerDiagOffsetNEY = " & Me.TowerDiagOffsetNEY.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerDiagOffsetNEX = " & Me.TowerDiagOffsetNEX.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerDiagOffsetPEY = " & Me.TowerDiagOffsetPEY.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerDiagOffsetPEX = " & Me.TowerDiagOffsetPEX.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerKbraceOffsetNEY = " & Me.TowerKbraceOffsetNEY.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerKbraceOffsetNEX = " & Me.TowerKbraceOffsetNEX.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerKbraceOffsetPEY = " & Me.TowerKbraceOffsetPEY.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerKbraceOffsetPEX = " & Me.TowerKbraceOffsetPEX.NullableToString.FormatDBValue)

        Return SQLUpdateFieldsandValues
    End Function

#End Region

#Region "Define"
    'Private _TowerRec As Integer?
    Private _TowerDatabase As String
    Private _TowerName As String
    Private _TowerHeight As Double?
    Private _TowerFaceWidth As Double?
    Private _TowerNumSections As Integer?
    Private _TowerSectionLength As Double?
    Private _TowerDiagonalSpacing As Double?
    Private _TowerDiagonalSpacingEx As Double?
    Private _TowerBraceType As String
    Private _TowerFaceBevel As Double?
    Private _TowerTopGirtOffset As Double?
    Private _TowerBotGirtOffset As Double?
    Private _TowerHasKBraceEndPanels As Boolean?
    Private _TowerHasHorizontals As Boolean?
    Private _TowerLegType As String
    Private _TowerLegSize As String
    Private _TowerLegGrade As Double?
    Private _TowerLegMatlGrade As String
    Private _TowerDiagonalGrade As Double?
    Private _TowerDiagonalMatlGrade As String
    Private _TowerInnerBracingGrade As Double?
    Private _TowerInnerBracingMatlGrade As String
    Private _TowerTopGirtGrade As Double?
    Private _TowerTopGirtMatlGrade As String
    Private _TowerBotGirtGrade As Double?
    Private _TowerBotGirtMatlGrade As String
    Private _TowerInnerGirtGrade As Double?
    Private _TowerInnerGirtMatlGrade As String
    Private _TowerLongHorizontalGrade As Double?
    Private _TowerLongHorizontalMatlGrade As String
    Private _TowerShortHorizontalGrade As Double?
    Private _TowerShortHorizontalMatlGrade As String
    Private _TowerDiagonalType As String
    Private _TowerDiagonalSize As String
    Private _TowerInnerBracingType As String
    Private _TowerInnerBracingSize As String
    Private _TowerTopGirtType As String
    Private _TowerTopGirtSize As String
    Private _TowerBotGirtType As String
    Private _TowerBotGirtSize As String
    Private _TowerNumInnerGirts As Integer?
    Private _TowerInnerGirtType As String
    Private _TowerInnerGirtSize As String
    Private _TowerLongHorizontalType As String
    Private _TowerLongHorizontalSize As String
    Private _TowerShortHorizontalType As String
    Private _TowerShortHorizontalSize As String
    Private _TowerRedundantGrade As Double?
    Private _TowerRedundantMatlGrade As String
    Private _TowerRedundantType As String
    Private _TowerRedundantDiagType As String
    Private _TowerRedundantSubDiagonalType As String
    Private _TowerRedundantSubHorizontalType As String
    Private _TowerRedundantVerticalType As String
    Private _TowerRedundantHipType As String
    Private _TowerRedundantHipDiagonalType As String
    Private _TowerRedundantHorizontalSize As String
    Private _TowerRedundantHorizontalSize2 As String
    Private _TowerRedundantHorizontalSize3 As String
    Private _TowerRedundantHorizontalSize4 As String
    Private _TowerRedundantDiagonalSize As String
    Private _TowerRedundantDiagonalSize2 As String
    Private _TowerRedundantDiagonalSize3 As String
    Private _TowerRedundantDiagonalSize4 As String
    Private _TowerRedundantSubHorizontalSize As String
    Private _TowerRedundantSubDiagonalSize As String
    Private _TowerSubDiagLocation As Double?
    Private _TowerRedundantVerticalSize As String
    Private _TowerRedundantHipSize As String
    Private _TowerRedundantHipSize2 As String
    Private _TowerRedundantHipSize3 As String
    Private _TowerRedundantHipSize4 As String
    Private _TowerRedundantHipDiagonalSize As String
    Private _TowerRedundantHipDiagonalSize2 As String
    Private _TowerRedundantHipDiagonalSize3 As String
    Private _TowerRedundantHipDiagonalSize4 As String
    Private _TowerSWMult As Double?
    Private _TowerWPMult As Double?
    Private _TowerAutoCalcKSingleAngle As Boolean?
    Private _TowerAutoCalcKSolidRound As Boolean?
    Private _TowerAfGusset As Double?
    Private _TowerTfGusset As Double?
    Private _TowerGussetBoltEdgeDistance As Double?
    Private _TowerGussetGrade As Double?
    Private _TowerGussetMatlGrade As String
    Private _TowerAfMult As Double?
    Private _TowerArMult As Double?
    Private _TowerFlatIPAPole As Double?
    Private _TowerRoundIPAPole As Double?
    Private _TowerFlatIPALeg As Double?
    Private _TowerRoundIPALeg As Double?
    Private _TowerFlatIPAHorizontal As Double?
    Private _TowerRoundIPAHorizontal As Double?
    Private _TowerFlatIPADiagonal As Double?
    Private _TowerRoundIPADiagonal As Double?
    Private _TowerCSA_S37_SpeedUpFactor As Double?
    Private _TowerKLegs As Double?
    Private _TowerKXBracedDiags As Double?
    Private _TowerKKBracedDiags As Double?
    Private _TowerKZBracedDiags As Double?
    Private _TowerKHorzs As Double?
    Private _TowerKSecHorzs As Double?
    Private _TowerKGirts As Double?
    Private _TowerKInners As Double?
    Private _TowerKXBracedDiagsY As Double?
    Private _TowerKKBracedDiagsY As Double?
    Private _TowerKZBracedDiagsY As Double?
    Private _TowerKHorzsY As Double?
    Private _TowerKSecHorzsY As Double?
    Private _TowerKGirtsY As Double?
    Private _TowerKInnersY As Double?
    Private _TowerKRedHorz As Double?
    Private _TowerKRedDiag As Double?
    Private _TowerKRedSubDiag As Double?
    Private _TowerKRedSubHorz As Double?
    Private _TowerKRedVert As Double?
    Private _TowerKRedHip As Double?
    Private _TowerKRedHipDiag As Double?
    Private _TowerKTLX As Double?
    Private _TowerKTLZ As Double?
    Private _TowerKTLLeg As Double?
    Private _TowerInnerKTLX As Double?
    Private _TowerInnerKTLZ As Double?
    Private _TowerInnerKTLLeg As Double?
    Private _TowerStitchBoltLocationHoriz As String
    Private _TowerStitchBoltLocationDiag As String
    Private _TowerStitchBoltLocationRed As String
    Private _TowerStitchSpacing As Double?
    Private _TowerStitchSpacingDiag As Double?
    Private _TowerStitchSpacingHorz As Double?
    Private _TowerStitchSpacingRed As Double?
    Private _TowerLegNetWidthDeduct As Double?
    Private _TowerLegUFactor As Double?
    Private _TowerDiagonalNetWidthDeduct As Double?
    Private _TowerTopGirtNetWidthDeduct As Double?
    Private _TowerBotGirtNetWidthDeduct As Double?
    Private _TowerInnerGirtNetWidthDeduct As Double?
    Private _TowerHorizontalNetWidthDeduct As Double?
    Private _TowerShortHorizontalNetWidthDeduct As Double?
    Private _TowerDiagonalUFactor As Double?
    Private _TowerTopGirtUFactor As Double?
    Private _TowerBotGirtUFactor As Double?
    Private _TowerInnerGirtUFactor As Double?
    Private _TowerHorizontalUFactor As Double?
    Private _TowerShortHorizontalUFactor As Double?
    Private _TowerLegConnType As String
    Private _TowerLegNumBolts As Integer?
    Private _TowerDiagonalNumBolts As Integer?
    Private _TowerTopGirtNumBolts As Integer?
    Private _TowerBotGirtNumBolts As Integer?
    Private _TowerInnerGirtNumBolts As Integer?
    Private _TowerHorizontalNumBolts As Integer?
    Private _TowerShortHorizontalNumBolts As Integer?
    Private _TowerLegBoltGrade As String
    Private _TowerLegBoltSize As Double?
    Private _TowerDiagonalBoltGrade As String
    Private _TowerDiagonalBoltSize As Double?
    Private _TowerTopGirtBoltGrade As String
    Private _TowerTopGirtBoltSize As Double?
    Private _TowerBotGirtBoltGrade As String
    Private _TowerBotGirtBoltSize As Double?
    Private _TowerInnerGirtBoltGrade As String
    Private _TowerInnerGirtBoltSize As Double?
    Private _TowerHorizontalBoltGrade As String
    Private _TowerHorizontalBoltSize As Double?
    Private _TowerShortHorizontalBoltGrade As String
    Private _TowerShortHorizontalBoltSize As Double?
    Private _TowerLegBoltEdgeDistance As Double?
    Private _TowerDiagonalBoltEdgeDistance As Double?
    Private _TowerTopGirtBoltEdgeDistance As Double?
    Private _TowerBotGirtBoltEdgeDistance As Double?
    Private _TowerInnerGirtBoltEdgeDistance As Double?
    Private _TowerHorizontalBoltEdgeDistance As Double?
    Private _TowerShortHorizontalBoltEdgeDistance As Double?
    Private _TowerDiagonalGageG1Distance As Double?
    Private _TowerTopGirtGageG1Distance As Double?
    Private _TowerBotGirtGageG1Distance As Double?
    Private _TowerInnerGirtGageG1Distance As Double?
    Private _TowerHorizontalGageG1Distance As Double?
    Private _TowerShortHorizontalGageG1Distance As Double?
    Private _TowerRedundantHorizontalBoltGrade As String
    Private _TowerRedundantHorizontalBoltSize As Double?
    Private _TowerRedundantHorizontalNumBolts As Integer?
    Private _TowerRedundantHorizontalBoltEdgeDistance As Double?
    Private _TowerRedundantHorizontalGageG1Distance As Double?
    Private _TowerRedundantHorizontalNetWidthDeduct As Double?
    Private _TowerRedundantHorizontalUFactor As Double?
    Private _TowerRedundantDiagonalBoltGrade As String
    Private _TowerRedundantDiagonalBoltSize As Double?
    Private _TowerRedundantDiagonalNumBolts As Integer?
    Private _TowerRedundantDiagonalBoltEdgeDistance As Double?
    Private _TowerRedundantDiagonalGageG1Distance As Double?
    Private _TowerRedundantDiagonalNetWidthDeduct As Double?
    Private _TowerRedundantDiagonalUFactor As Double?
    Private _TowerRedundantSubDiagonalBoltGrade As String
    Private _TowerRedundantSubDiagonalBoltSize As Double?
    Private _TowerRedundantSubDiagonalNumBolts As Integer?
    Private _TowerRedundantSubDiagonalBoltEdgeDistance As Double?
    Private _TowerRedundantSubDiagonalGageG1Distance As Double?
    Private _TowerRedundantSubDiagonalNetWidthDeduct As Double?
    Private _TowerRedundantSubDiagonalUFactor As Double?
    Private _TowerRedundantSubHorizontalBoltGrade As String
    Private _TowerRedundantSubHorizontalBoltSize As Double?
    Private _TowerRedundantSubHorizontalNumBolts As Integer?
    Private _TowerRedundantSubHorizontalBoltEdgeDistance As Double?
    Private _TowerRedundantSubHorizontalGageG1Distance As Double?
    Private _TowerRedundantSubHorizontalNetWidthDeduct As Double?
    Private _TowerRedundantSubHorizontalUFactor As Double?
    Private _TowerRedundantVerticalBoltGrade As String
    Private _TowerRedundantVerticalBoltSize As Double?
    Private _TowerRedundantVerticalNumBolts As Integer?
    Private _TowerRedundantVerticalBoltEdgeDistance As Double?
    Private _TowerRedundantVerticalGageG1Distance As Double?
    Private _TowerRedundantVerticalNetWidthDeduct As Double?
    Private _TowerRedundantVerticalUFactor As Double?
    Private _TowerRedundantHipBoltGrade As String
    Private _TowerRedundantHipBoltSize As Double?
    Private _TowerRedundantHipNumBolts As Integer?
    Private _TowerRedundantHipBoltEdgeDistance As Double?
    Private _TowerRedundantHipGageG1Distance As Double?
    Private _TowerRedundantHipNetWidthDeduct As Double?
    Private _TowerRedundantHipUFactor As Double?
    Private _TowerRedundantHipDiagonalBoltGrade As String
    Private _TowerRedundantHipDiagonalBoltSize As Double?
    Private _TowerRedundantHipDiagonalNumBolts As Integer?
    Private _TowerRedundantHipDiagonalBoltEdgeDistance As Double?
    Private _TowerRedundantHipDiagonalGageG1Distance As Double?
    Private _TowerRedundantHipDiagonalNetWidthDeduct As Double?
    Private _TowerRedundantHipDiagonalUFactor As Double?
    Private _TowerDiagonalOutOfPlaneRestraint As Boolean?
    Private _TowerTopGirtOutOfPlaneRestraint As Boolean?
    Private _TowerBottomGirtOutOfPlaneRestraint As Boolean?
    Private _TowerMidGirtOutOfPlaneRestraint As Boolean?
    Private _TowerHorizontalOutOfPlaneRestraint As Boolean?
    Private _TowerSecondaryHorizontalOutOfPlaneRestraint As Boolean?
    Private _TowerUniqueFlag As Integer?
    Private _TowerDiagOffsetNEY As Double?
    Private _TowerDiagOffsetNEX As Double?
    Private _TowerDiagOffsetPEY As Double?
    Private _TowerDiagOffsetPEX As Double?
    Private _TowerKbraceOffsetNEY As Double?
    Private _TowerKbraceOffsetNEX As Double?
    Private _TowerKbraceOffsetPEY As Double?
    Private _TowerKbraceOffsetPEX As Double?

    '<Category("TNX Tower Record"), Description(""), DisplayName("Towerrec")>
    ' <DataMember()> Public Property Rec() As Integer?
    '    Get
    '        Return Me._TowerRec
    '    End Get
    '    Set
    '        Me._TowerRec = Value
    '    End Set
    'End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdatabase")>
    <DataMember()> Public Property TowerDatabase() As String
        Get
            Return Me._TowerDatabase
        End Get
        Set
            Me._TowerDatabase = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towername")>
    <DataMember()> Public Property TowerName() As String
        Get
            Return Me._TowerName
        End Get
        Set
            Me._TowerName = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerheight")>
    <DataMember()> Public Property TowerHeight() As Double?
        Get
            Return Me._TowerHeight
        End Get
        Set
            Me._TowerHeight = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerfacewidth")>
    <DataMember()> Public Property TowerFaceWidth() As Double?
        Get
            Return Me._TowerFaceWidth
        End Get
        Set
            Me._TowerFaceWidth = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towernumsections")>
    <DataMember()> Public Property TowerNumSections() As Integer?
        Get
            Return Me._TowerNumSections
        End Get
        Set
            Me._TowerNumSections = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towersectionlength")>
    <DataMember()> Public Property TowerSectionLength() As Double?
        Get
            Return Me._TowerSectionLength
        End Get
        Set
            Me._TowerSectionLength = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagonalspacing")>
    <DataMember()> Public Property TowerDiagonalSpacing() As Double?
        Get
            Return Me._TowerDiagonalSpacing
        End Get
        Set
            Me._TowerDiagonalSpacing = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagonalspacingex")>
    <DataMember()> Public Property TowerDiagonalSpacingEx() As Double?
        Get
            Return Me._TowerDiagonalSpacingEx
        End Get
        Set
            Me._TowerDiagonalSpacingEx = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerbracetype")>
    <DataMember()> Public Property TowerBraceType() As String
        Get
            Return Me._TowerBraceType
        End Get
        Set
            Me._TowerBraceType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerfacebevel")>
    <DataMember()> Public Property TowerFaceBevel() As Double?
        Get
            Return Me._TowerFaceBevel
        End Get
        Set
            Me._TowerFaceBevel = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towertopgirtoffset")>
    <DataMember()> Public Property TowerTopGirtOffset() As Double?
        Get
            Return Me._TowerTopGirtOffset
        End Get
        Set
            Me._TowerTopGirtOffset = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerbotgirtoffset")>
    <DataMember()> Public Property TowerBotGirtOffset() As Double?
        Get
            Return Me._TowerBotGirtOffset
        End Get
        Set
            Me._TowerBotGirtOffset = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerhaskbraceendpanels")>
    <DataMember()> Public Property TowerHasKBraceEndPanels() As Boolean?
        Get
            Return Me._TowerHasKBraceEndPanels
        End Get
        Set
            Me._TowerHasKBraceEndPanels = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerhashorizontals")>
    <DataMember()> Public Property TowerHasHorizontals() As Boolean?
        Get
            Return Me._TowerHasHorizontals
        End Get
        Set
            Me._TowerHasHorizontals = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerlegtype")>
    <DataMember()> Public Property TowerLegType() As String
        Get
            Return Me._TowerLegType
        End Get
        Set
            Me._TowerLegType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerlegsize")>
    <DataMember()> Public Property TowerLegSize() As String
        Get
            Return Me._TowerLegSize
        End Get
        Set
            Me._TowerLegSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerleggrade")>
    <DataMember()> Public Property TowerLegGrade() As Double?
        Get
            Return Me._TowerLegGrade
        End Get
        Set
            Me._TowerLegGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerlegmatlgrade")>
    <DataMember()> Public Property TowerLegMatlGrade() As String
        Get
            Return Me._TowerLegMatlGrade
        End Get
        Set
            Me._TowerLegMatlGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagonalgrade")>
    <DataMember()> Public Property TowerDiagonalGrade() As Double?
        Get
            Return Me._TowerDiagonalGrade
        End Get
        Set
            Me._TowerDiagonalGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagonalmatlgrade")>
    <DataMember()> Public Property TowerDiagonalMatlGrade() As String
        Get
            Return Me._TowerDiagonalMatlGrade
        End Get
        Set
            Me._TowerDiagonalMatlGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnerbracinggrade")>
    <DataMember()> Public Property TowerInnerBracingGrade() As Double?
        Get
            Return Me._TowerInnerBracingGrade
        End Get
        Set
            Me._TowerInnerBracingGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnerbracingmatlgrade")>
    <DataMember()> Public Property TowerInnerBracingMatlGrade() As String
        Get
            Return Me._TowerInnerBracingMatlGrade
        End Get
        Set
            Me._TowerInnerBracingMatlGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towertopgirtgrade")>
    <DataMember()> Public Property TowerTopGirtGrade() As Double?
        Get
            Return Me._TowerTopGirtGrade
        End Get
        Set
            Me._TowerTopGirtGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towertopgirtmatlgrade")>
    <DataMember()> Public Property TowerTopGirtMatlGrade() As String
        Get
            Return Me._TowerTopGirtMatlGrade
        End Get
        Set
            Me._TowerTopGirtMatlGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerbotgirtgrade")>
    <DataMember()> Public Property TowerBotGirtGrade() As Double?
        Get
            Return Me._TowerBotGirtGrade
        End Get
        Set
            Me._TowerBotGirtGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerbotgirtmatlgrade")>
    <DataMember()> Public Property TowerBotGirtMatlGrade() As String
        Get
            Return Me._TowerBotGirtMatlGrade
        End Get
        Set
            Me._TowerBotGirtMatlGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnergirtgrade")>
    <DataMember()> Public Property TowerInnerGirtGrade() As Double?
        Get
            Return Me._TowerInnerGirtGrade
        End Get
        Set
            Me._TowerInnerGirtGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnergirtmatlgrade")>
    <DataMember()> Public Property TowerInnerGirtMatlGrade() As String
        Get
            Return Me._TowerInnerGirtMatlGrade
        End Get
        Set
            Me._TowerInnerGirtMatlGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerlonghorizontalgrade")>
    <DataMember()> Public Property TowerLongHorizontalGrade() As Double?
        Get
            Return Me._TowerLongHorizontalGrade
        End Get
        Set
            Me._TowerLongHorizontalGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerlonghorizontalmatlgrade")>
    <DataMember()> Public Property TowerLongHorizontalMatlGrade() As String
        Get
            Return Me._TowerLongHorizontalMatlGrade
        End Get
        Set
            Me._TowerLongHorizontalMatlGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towershorthorizontalgrade")>
    <DataMember()> Public Property TowerShortHorizontalGrade() As Double?
        Get
            Return Me._TowerShortHorizontalGrade
        End Get
        Set
            Me._TowerShortHorizontalGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towershorthorizontalmatlgrade")>
    <DataMember()> Public Property TowerShortHorizontalMatlGrade() As String
        Get
            Return Me._TowerShortHorizontalMatlGrade
        End Get
        Set
            Me._TowerShortHorizontalMatlGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagonaltype")>
    <DataMember()> Public Property TowerDiagonalType() As String
        Get
            Return Me._TowerDiagonalType
        End Get
        Set
            Me._TowerDiagonalType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagonalsize")>
    <DataMember()> Public Property TowerDiagonalSize() As String
        Get
            Return Me._TowerDiagonalSize
        End Get
        Set
            Me._TowerDiagonalSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnerbracingtype")>
    <DataMember()> Public Property TowerInnerBracingType() As String
        Get
            Return Me._TowerInnerBracingType
        End Get
        Set
            Me._TowerInnerBracingType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnerbracingsize")>
    <DataMember()> Public Property TowerInnerBracingSize() As String
        Get
            Return Me._TowerInnerBracingSize
        End Get
        Set
            Me._TowerInnerBracingSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towertopgirttype")>
    <DataMember()> Public Property TowerTopGirtType() As String
        Get
            Return Me._TowerTopGirtType
        End Get
        Set
            Me._TowerTopGirtType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towertopgirtsize")>
    <DataMember()> Public Property TowerTopGirtSize() As String
        Get
            Return Me._TowerTopGirtSize
        End Get
        Set
            Me._TowerTopGirtSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerbotgirttype")>
    <DataMember()> Public Property TowerBotGirtType() As String
        Get
            Return Me._TowerBotGirtType
        End Get
        Set
            Me._TowerBotGirtType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerbotgirtsize")>
    <DataMember()> Public Property TowerBotGirtSize() As String
        Get
            Return Me._TowerBotGirtSize
        End Get
        Set
            Me._TowerBotGirtSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towernuminnergirts")>
    <DataMember()> Public Property TowerNumInnerGirts() As Integer?
        Get
            Return Me._TowerNumInnerGirts
        End Get
        Set
            Me._TowerNumInnerGirts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnergirttype")>
    <DataMember()> Public Property TowerInnerGirtType() As String
        Get
            Return Me._TowerInnerGirtType
        End Get
        Set
            Me._TowerInnerGirtType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnergirtsize")>
    <DataMember()> Public Property TowerInnerGirtSize() As String
        Get
            Return Me._TowerInnerGirtSize
        End Get
        Set
            Me._TowerInnerGirtSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerlonghorizontaltype")>
    <DataMember()> Public Property TowerLongHorizontalType() As String
        Get
            Return Me._TowerLongHorizontalType
        End Get
        Set
            Me._TowerLongHorizontalType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerlonghorizontalsize")>
    <DataMember()> Public Property TowerLongHorizontalSize() As String
        Get
            Return Me._TowerLongHorizontalSize
        End Get
        Set
            Me._TowerLongHorizontalSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towershorthorizontaltype")>
    <DataMember()> Public Property TowerShortHorizontalType() As String
        Get
            Return Me._TowerShortHorizontalType
        End Get
        Set
            Me._TowerShortHorizontalType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towershorthorizontalsize")>
    <DataMember()> Public Property TowerShortHorizontalSize() As String
        Get
            Return Me._TowerShortHorizontalSize
        End Get
        Set
            Me._TowerShortHorizontalSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantgrade")>
    <DataMember()> Public Property TowerRedundantGrade() As Double?
        Get
            Return Me._TowerRedundantGrade
        End Get
        Set
            Me._TowerRedundantGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantmatlgrade")>
    <DataMember()> Public Property TowerRedundantMatlGrade() As String
        Get
            Return Me._TowerRedundantMatlGrade
        End Get
        Set
            Me._TowerRedundantMatlGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanttype")>
    <DataMember()> Public Property TowerRedundantType() As String
        Get
            Return Me._TowerRedundantType
        End Get
        Set
            Me._TowerRedundantType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantdiagtype")>
    <DataMember()> Public Property TowerRedundantDiagType() As String
        Get
            Return Me._TowerRedundantDiagType
        End Get
        Set
            Me._TowerRedundantDiagType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubdiagonaltype")>
    <DataMember()> Public Property TowerRedundantSubDiagonalType() As String
        Get
            Return Me._TowerRedundantSubDiagonalType
        End Get
        Set
            Me._TowerRedundantSubDiagonalType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubhorizontaltype")>
    <DataMember()> Public Property TowerRedundantSubHorizontalType() As String
        Get
            Return Me._TowerRedundantSubHorizontalType
        End Get
        Set
            Me._TowerRedundantSubHorizontalType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantverticaltype")>
    <DataMember()> Public Property TowerRedundantVerticalType() As String
        Get
            Return Me._TowerRedundantVerticalType
        End Get
        Set
            Me._TowerRedundantVerticalType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthiptype")>
    <DataMember()> Public Property TowerRedundantHipType() As String
        Get
            Return Me._TowerRedundantHipType
        End Get
        Set
            Me._TowerRedundantHipType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipdiagonaltype")>
    <DataMember()> Public Property TowerRedundantHipDiagonalType() As String
        Get
            Return Me._TowerRedundantHipDiagonalType
        End Get
        Set
            Me._TowerRedundantHipDiagonalType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthorizontalsize")>
    <DataMember()> Public Property TowerRedundantHorizontalSize() As String
        Get
            Return Me._TowerRedundantHorizontalSize
        End Get
        Set
            Me._TowerRedundantHorizontalSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthorizontalsize2")>
    <DataMember()> Public Property TowerRedundantHorizontalSize2() As String
        Get
            Return Me._TowerRedundantHorizontalSize2
        End Get
        Set
            Me._TowerRedundantHorizontalSize2 = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthorizontalsize3")>
    <DataMember()> Public Property TowerRedundantHorizontalSize3() As String
        Get
            Return Me._TowerRedundantHorizontalSize3
        End Get
        Set
            Me._TowerRedundantHorizontalSize3 = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthorizontalsize4")>
    <DataMember()> Public Property TowerRedundantHorizontalSize4() As String
        Get
            Return Me._TowerRedundantHorizontalSize4
        End Get
        Set
            Me._TowerRedundantHorizontalSize4 = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantdiagonalsize")>
    <DataMember()> Public Property TowerRedundantDiagonalSize() As String
        Get
            Return Me._TowerRedundantDiagonalSize
        End Get
        Set
            Me._TowerRedundantDiagonalSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantdiagonalsize2")>
    <DataMember()> Public Property TowerRedundantDiagonalSize2() As String
        Get
            Return Me._TowerRedundantDiagonalSize2
        End Get
        Set
            Me._TowerRedundantDiagonalSize2 = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantdiagonalsize3")>
    <DataMember()> Public Property TowerRedundantDiagonalSize3() As String
        Get
            Return Me._TowerRedundantDiagonalSize3
        End Get
        Set
            Me._TowerRedundantDiagonalSize3 = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantdiagonalsize4")>
    <DataMember()> Public Property TowerRedundantDiagonalSize4() As String
        Get
            Return Me._TowerRedundantDiagonalSize4
        End Get
        Set
            Me._TowerRedundantDiagonalSize4 = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubhorizontalsize")>
    <DataMember()> Public Property TowerRedundantSubHorizontalSize() As String
        Get
            Return Me._TowerRedundantSubHorizontalSize
        End Get
        Set
            Me._TowerRedundantSubHorizontalSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubdiagonalsize")>
    <DataMember()> Public Property TowerRedundantSubDiagonalSize() As String
        Get
            Return Me._TowerRedundantSubDiagonalSize
        End Get
        Set
            Me._TowerRedundantSubDiagonalSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towersubdiaglocation")>
    <DataMember()> Public Property TowerSubDiagLocation() As Double?
        Get
            Return Me._TowerSubDiagLocation
        End Get
        Set
            Me._TowerSubDiagLocation = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantverticalsize")>
    <DataMember()> Public Property TowerRedundantVerticalSize() As String
        Get
            Return Me._TowerRedundantVerticalSize
        End Get
        Set
            Me._TowerRedundantVerticalSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipsize")>
    <DataMember()> Public Property TowerRedundantHipSize() As String
        Get
            Return Me._TowerRedundantHipSize
        End Get
        Set
            Me._TowerRedundantHipSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipsize2")>
    <DataMember()> Public Property TowerRedundantHipSize2() As String
        Get
            Return Me._TowerRedundantHipSize2
        End Get
        Set
            Me._TowerRedundantHipSize2 = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipsize3")>
    <DataMember()> Public Property TowerRedundantHipSize3() As String
        Get
            Return Me._TowerRedundantHipSize3
        End Get
        Set
            Me._TowerRedundantHipSize3 = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipsize4")>
    <DataMember()> Public Property TowerRedundantHipSize4() As String
        Get
            Return Me._TowerRedundantHipSize4
        End Get
        Set
            Me._TowerRedundantHipSize4 = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipdiagonalsize")>
    <DataMember()> Public Property TowerRedundantHipDiagonalSize() As String
        Get
            Return Me._TowerRedundantHipDiagonalSize
        End Get
        Set
            Me._TowerRedundantHipDiagonalSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipdiagonalsize2")>
    <DataMember()> Public Property TowerRedundantHipDiagonalSize2() As String
        Get
            Return Me._TowerRedundantHipDiagonalSize2
        End Get
        Set
            Me._TowerRedundantHipDiagonalSize2 = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipdiagonalsize3")>
    <DataMember()> Public Property TowerRedundantHipDiagonalSize3() As String
        Get
            Return Me._TowerRedundantHipDiagonalSize3
        End Get
        Set
            Me._TowerRedundantHipDiagonalSize3 = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipdiagonalsize4")>
    <DataMember()> Public Property TowerRedundantHipDiagonalSize4() As String
        Get
            Return Me._TowerRedundantHipDiagonalSize4
        End Get
        Set
            Me._TowerRedundantHipDiagonalSize4 = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerswmult")>
    <DataMember()> Public Property TowerSWMult() As Double?
        Get
            Return Me._TowerSWMult
        End Get
        Set
            Me._TowerSWMult = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerwpmult")>
    <DataMember()> Public Property TowerWPMult() As Double?
        Get
            Return Me._TowerWPMult
        End Get
        Set
            Me._TowerWPMult = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerautocalcksingleangle")>
    <DataMember()> Public Property TowerAutoCalcKSingleAngle() As Boolean?
        Get
            Return Me._TowerAutoCalcKSingleAngle
        End Get
        Set
            Me._TowerAutoCalcKSingleAngle = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerautocalcksolidround")>
    <DataMember()> Public Property TowerAutoCalcKSolidRound() As Boolean?
        Get
            Return Me._TowerAutoCalcKSolidRound
        End Get
        Set
            Me._TowerAutoCalcKSolidRound = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerafgusset")>
    <DataMember()> Public Property TowerAfGusset() As Double?
        Get
            Return Me._TowerAfGusset
        End Get
        Set
            Me._TowerAfGusset = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towertfgusset")>
    <DataMember()> Public Property TowerTfGusset() As Double?
        Get
            Return Me._TowerTfGusset
        End Get
        Set
            Me._TowerTfGusset = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towergussetboltedgedistance")>
    <DataMember()> Public Property TowerGussetBoltEdgeDistance() As Double?
        Get
            Return Me._TowerGussetBoltEdgeDistance
        End Get
        Set
            Me._TowerGussetBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towergussetgrade")>
    <DataMember()> Public Property TowerGussetGrade() As Double?
        Get
            Return Me._TowerGussetGrade
        End Get
        Set
            Me._TowerGussetGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towergussetmatlgrade")>
    <DataMember()> Public Property TowerGussetMatlGrade() As String
        Get
            Return Me._TowerGussetMatlGrade
        End Get
        Set
            Me._TowerGussetMatlGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerafmult")>
    <DataMember()> Public Property TowerAfMult() As Double?
        Get
            Return Me._TowerAfMult
        End Get
        Set
            Me._TowerAfMult = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerarmult")>
    <DataMember()> Public Property TowerArMult() As Double?
        Get
            Return Me._TowerArMult
        End Get
        Set
            Me._TowerArMult = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerflatipapole")>
    <DataMember()> Public Property TowerFlatIPAPole() As Double?
        Get
            Return Me._TowerFlatIPAPole
        End Get
        Set
            Me._TowerFlatIPAPole = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerroundipapole")>
    <DataMember()> Public Property TowerRoundIPAPole() As Double?
        Get
            Return Me._TowerRoundIPAPole
        End Get
        Set
            Me._TowerRoundIPAPole = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerflatipaleg")>
    <DataMember()> Public Property TowerFlatIPALeg() As Double?
        Get
            Return Me._TowerFlatIPALeg
        End Get
        Set
            Me._TowerFlatIPALeg = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerroundipaleg")>
    <DataMember()> Public Property TowerRoundIPALeg() As Double?
        Get
            Return Me._TowerRoundIPALeg
        End Get
        Set
            Me._TowerRoundIPALeg = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerflatipahorizontal")>
    <DataMember()> Public Property TowerFlatIPAHorizontal() As Double?
        Get
            Return Me._TowerFlatIPAHorizontal
        End Get
        Set
            Me._TowerFlatIPAHorizontal = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerroundipahorizontal")>
    <DataMember()> Public Property TowerRoundIPAHorizontal() As Double?
        Get
            Return Me._TowerRoundIPAHorizontal
        End Get
        Set
            Me._TowerRoundIPAHorizontal = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerflatipadiagonal")>
    <DataMember()> Public Property TowerFlatIPADiagonal() As Double?
        Get
            Return Me._TowerFlatIPADiagonal
        End Get
        Set
            Me._TowerFlatIPADiagonal = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerroundipadiagonal")>
    <DataMember()> Public Property TowerRoundIPADiagonal() As Double?
        Get
            Return Me._TowerRoundIPADiagonal
        End Get
        Set
            Me._TowerRoundIPADiagonal = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towercsa_S37_Speedupfactor")>
    <DataMember()> Public Property TowerCSA_S37_SpeedUpFactor() As Double?
        Get
            Return Me._TowerCSA_S37_SpeedUpFactor
        End Get
        Set
            Me._TowerCSA_S37_SpeedUpFactor = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerklegs")>
    <DataMember()> Public Property TowerKLegs() As Double?
        Get
            Return Me._TowerKLegs
        End Get
        Set
            Me._TowerKLegs = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkxbraceddiags")>
    <DataMember()> Public Property TowerKXBracedDiags() As Double?
        Get
            Return Me._TowerKXBracedDiags
        End Get
        Set
            Me._TowerKXBracedDiags = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkkbraceddiags")>
    <DataMember()> Public Property TowerKKBracedDiags() As Double?
        Get
            Return Me._TowerKKBracedDiags
        End Get
        Set
            Me._TowerKKBracedDiags = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkzbraceddiags")>
    <DataMember()> Public Property TowerKZBracedDiags() As Double?
        Get
            Return Me._TowerKZBracedDiags
        End Get
        Set
            Me._TowerKZBracedDiags = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkhorzs")>
    <DataMember()> Public Property TowerKHorzs() As Double?
        Get
            Return Me._TowerKHorzs
        End Get
        Set
            Me._TowerKHorzs = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerksechorzs")>
    <DataMember()> Public Property TowerKSecHorzs() As Double?
        Get
            Return Me._TowerKSecHorzs
        End Get
        Set
            Me._TowerKSecHorzs = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkgirts")>
    <DataMember()> Public Property TowerKGirts() As Double?
        Get
            Return Me._TowerKGirts
        End Get
        Set
            Me._TowerKGirts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkinners")>
    <DataMember()> Public Property TowerKInners() As Double?
        Get
            Return Me._TowerKInners
        End Get
        Set
            Me._TowerKInners = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkxbraceddiagsy")>
    <DataMember()> Public Property TowerKXBracedDiagsY() As Double?
        Get
            Return Me._TowerKXBracedDiagsY
        End Get
        Set
            Me._TowerKXBracedDiagsY = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkkbraceddiagsy")>
    <DataMember()> Public Property TowerKKBracedDiagsY() As Double?
        Get
            Return Me._TowerKKBracedDiagsY
        End Get
        Set
            Me._TowerKKBracedDiagsY = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkzbraceddiagsy")>
    <DataMember()> Public Property TowerKZBracedDiagsY() As Double?
        Get
            Return Me._TowerKZBracedDiagsY
        End Get
        Set
            Me._TowerKZBracedDiagsY = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkhorzsy")>
    <DataMember()> Public Property TowerKHorzsY() As Double?
        Get
            Return Me._TowerKHorzsY
        End Get
        Set
            Me._TowerKHorzsY = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerksechorzsy")>
    <DataMember()> Public Property TowerKSecHorzsY() As Double?
        Get
            Return Me._TowerKSecHorzsY
        End Get
        Set
            Me._TowerKSecHorzsY = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkgirtsy")>
    <DataMember()> Public Property TowerKGirtsY() As Double?
        Get
            Return Me._TowerKGirtsY
        End Get
        Set
            Me._TowerKGirtsY = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkinnersy")>
    <DataMember()> Public Property TowerKInnersY() As Double?
        Get
            Return Me._TowerKInnersY
        End Get
        Set
            Me._TowerKInnersY = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkredhorz")>
    <DataMember()> Public Property TowerKRedHorz() As Double?
        Get
            Return Me._TowerKRedHorz
        End Get
        Set
            Me._TowerKRedHorz = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkreddiag")>
    <DataMember()> Public Property TowerKRedDiag() As Double?
        Get
            Return Me._TowerKRedDiag
        End Get
        Set
            Me._TowerKRedDiag = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkredsubdiag")>
    <DataMember()> Public Property TowerKRedSubDiag() As Double?
        Get
            Return Me._TowerKRedSubDiag
        End Get
        Set
            Me._TowerKRedSubDiag = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkredsubhorz")>
    <DataMember()> Public Property TowerKRedSubHorz() As Double?
        Get
            Return Me._TowerKRedSubHorz
        End Get
        Set
            Me._TowerKRedSubHorz = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkredvert")>
    <DataMember()> Public Property TowerKRedVert() As Double?
        Get
            Return Me._TowerKRedVert
        End Get
        Set
            Me._TowerKRedVert = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkredhip")>
    <DataMember()> Public Property TowerKRedHip() As Double?
        Get
            Return Me._TowerKRedHip
        End Get
        Set
            Me._TowerKRedHip = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkredhipdiag")>
    <DataMember()> Public Property TowerKRedHipDiag() As Double?
        Get
            Return Me._TowerKRedHipDiag
        End Get
        Set
            Me._TowerKRedHipDiag = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerktlx")>
    <DataMember()> Public Property TowerKTLX() As Double?
        Get
            Return Me._TowerKTLX
        End Get
        Set
            Me._TowerKTLX = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerktlz")>
    <DataMember()> Public Property TowerKTLZ() As Double?
        Get
            Return Me._TowerKTLZ
        End Get
        Set
            Me._TowerKTLZ = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerktlleg")>
    <DataMember()> Public Property TowerKTLLeg() As Double?
        Get
            Return Me._TowerKTLLeg
        End Get
        Set
            Me._TowerKTLLeg = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnerktlx")>
    <DataMember()> Public Property TowerInnerKTLX() As Double?
        Get
            Return Me._TowerInnerKTLX
        End Get
        Set
            Me._TowerInnerKTLX = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnerktlz")>
    <DataMember()> Public Property TowerInnerKTLZ() As Double?
        Get
            Return Me._TowerInnerKTLZ
        End Get
        Set
            Me._TowerInnerKTLZ = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnerktlleg")>
    <DataMember()> Public Property TowerInnerKTLLeg() As Double?
        Get
            Return Me._TowerInnerKTLLeg
        End Get
        Set
            Me._TowerInnerKTLLeg = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerstitchboltlocationhoriz")>
    <DataMember()> Public Property TowerStitchBoltLocationHoriz() As String
        Get
            Return Me._TowerStitchBoltLocationHoriz
        End Get
        Set
            Me._TowerStitchBoltLocationHoriz = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerstitchboltlocationdiag")>
    <DataMember()> Public Property TowerStitchBoltLocationDiag() As String
        Get
            Return Me._TowerStitchBoltLocationDiag
        End Get
        Set
            Me._TowerStitchBoltLocationDiag = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerstitchboltlocationred")>
    <DataMember()> Public Property TowerStitchBoltLocationRed() As String
        Get
            Return Me._TowerStitchBoltLocationRed
        End Get
        Set
            Me._TowerStitchBoltLocationRed = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerstitchspacing")>
    <DataMember()> Public Property TowerStitchSpacing() As Double?
        Get
            Return Me._TowerStitchSpacing
        End Get
        Set
            Me._TowerStitchSpacing = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerstitchspacingdiag")>
    <DataMember()> Public Property TowerStitchSpacingDiag() As Double?
        Get
            Return Me._TowerStitchSpacingDiag
        End Get
        Set
            Me._TowerStitchSpacingDiag = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerstitchspacinghorz")>
    <DataMember()> Public Property TowerStitchSpacingHorz() As Double?
        Get
            Return Me._TowerStitchSpacingHorz
        End Get
        Set
            Me._TowerStitchSpacingHorz = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerstitchspacingred")>
    <DataMember()> Public Property TowerStitchSpacingRed() As Double?
        Get
            Return Me._TowerStitchSpacingRed
        End Get
        Set
            Me._TowerStitchSpacingRed = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerlegnetwidthdeduct")>
    <DataMember()> Public Property TowerLegNetWidthDeduct() As Double?
        Get
            Return Me._TowerLegNetWidthDeduct
        End Get
        Set
            Me._TowerLegNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerlegufactor")>
    <DataMember()> Public Property TowerLegUFactor() As Double?
        Get
            Return Me._TowerLegUFactor
        End Get
        Set
            Me._TowerLegUFactor = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagonalnetwidthdeduct")>
    <DataMember()> Public Property TowerDiagonalNetWidthDeduct() As Double?
        Get
            Return Me._TowerDiagonalNetWidthDeduct
        End Get
        Set
            Me._TowerDiagonalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towertopgirtnetwidthdeduct")>
    <DataMember()> Public Property TowerTopGirtNetWidthDeduct() As Double?
        Get
            Return Me._TowerTopGirtNetWidthDeduct
        End Get
        Set
            Me._TowerTopGirtNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerbotgirtnetwidthdeduct")>
    <DataMember()> Public Property TowerBotGirtNetWidthDeduct() As Double?
        Get
            Return Me._TowerBotGirtNetWidthDeduct
        End Get
        Set
            Me._TowerBotGirtNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnergirtnetwidthdeduct")>
    <DataMember()> Public Property TowerInnerGirtNetWidthDeduct() As Double?
        Get
            Return Me._TowerInnerGirtNetWidthDeduct
        End Get
        Set
            Me._TowerInnerGirtNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerhorizontalnetwidthdeduct")>
    <DataMember()> Public Property TowerHorizontalNetWidthDeduct() As Double?
        Get
            Return Me._TowerHorizontalNetWidthDeduct
        End Get
        Set
            Me._TowerHorizontalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towershorthorizontalnetwidthdeduct")>
    <DataMember()> Public Property TowerShortHorizontalNetWidthDeduct() As Double?
        Get
            Return Me._TowerShortHorizontalNetWidthDeduct
        End Get
        Set
            Me._TowerShortHorizontalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagonalufactor")>
    <DataMember()> Public Property TowerDiagonalUFactor() As Double?
        Get
            Return Me._TowerDiagonalUFactor
        End Get
        Set
            Me._TowerDiagonalUFactor = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towertopgirtufactor")>
    <DataMember()> Public Property TowerTopGirtUFactor() As Double?
        Get
            Return Me._TowerTopGirtUFactor
        End Get
        Set
            Me._TowerTopGirtUFactor = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerbotgirtufactor")>
    <DataMember()> Public Property TowerBotGirtUFactor() As Double?
        Get
            Return Me._TowerBotGirtUFactor
        End Get
        Set
            Me._TowerBotGirtUFactor = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnergirtufactor")>
    <DataMember()> Public Property TowerInnerGirtUFactor() As Double?
        Get
            Return Me._TowerInnerGirtUFactor
        End Get
        Set
            Me._TowerInnerGirtUFactor = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerhorizontalufactor")>
    <DataMember()> Public Property TowerHorizontalUFactor() As Double?
        Get
            Return Me._TowerHorizontalUFactor
        End Get
        Set
            Me._TowerHorizontalUFactor = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towershorthorizontalufactor")>
    <DataMember()> Public Property TowerShortHorizontalUFactor() As Double?
        Get
            Return Me._TowerShortHorizontalUFactor
        End Get
        Set
            Me._TowerShortHorizontalUFactor = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerlegconntype")>
    <DataMember()> Public Property TowerLegConnType() As String
        Get
            Return Me._TowerLegConnType
        End Get
        Set
            Me._TowerLegConnType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerlegnumbolts")>
    <DataMember()> Public Property TowerLegNumBolts() As Integer?
        Get
            Return Me._TowerLegNumBolts
        End Get
        Set
            Me._TowerLegNumBolts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagonalnumbolts")>
    <DataMember()> Public Property TowerDiagonalNumBolts() As Integer?
        Get
            Return Me._TowerDiagonalNumBolts
        End Get
        Set
            Me._TowerDiagonalNumBolts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towertopgirtnumbolts")>
    <DataMember()> Public Property TowerTopGirtNumBolts() As Integer?
        Get
            Return Me._TowerTopGirtNumBolts
        End Get
        Set
            Me._TowerTopGirtNumBolts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerbotgirtnumbolts")>
    <DataMember()> Public Property TowerBotGirtNumBolts() As Integer?
        Get
            Return Me._TowerBotGirtNumBolts
        End Get
        Set
            Me._TowerBotGirtNumBolts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnergirtnumbolts")>
    <DataMember()> Public Property TowerInnerGirtNumBolts() As Integer?
        Get
            Return Me._TowerInnerGirtNumBolts
        End Get
        Set
            Me._TowerInnerGirtNumBolts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerhorizontalnumbolts")>
    <DataMember()> Public Property TowerHorizontalNumBolts() As Integer?
        Get
            Return Me._TowerHorizontalNumBolts
        End Get
        Set
            Me._TowerHorizontalNumBolts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towershorthorizontalnumbolts")>
    <DataMember()> Public Property TowerShortHorizontalNumBolts() As Integer?
        Get
            Return Me._TowerShortHorizontalNumBolts
        End Get
        Set
            Me._TowerShortHorizontalNumBolts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerlegboltgrade")>
    <DataMember()> Public Property TowerLegBoltGrade() As String
        Get
            Return Me._TowerLegBoltGrade
        End Get
        Set
            Me._TowerLegBoltGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerlegboltsize")>
    <DataMember()> Public Property TowerLegBoltSize() As Double?
        Get
            Return Me._TowerLegBoltSize
        End Get
        Set
            Me._TowerLegBoltSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagonalboltgrade")>
    <DataMember()> Public Property TowerDiagonalBoltGrade() As String
        Get
            Return Me._TowerDiagonalBoltGrade
        End Get
        Set
            Me._TowerDiagonalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagonalboltsize")>
    <DataMember()> Public Property TowerDiagonalBoltSize() As Double?
        Get
            Return Me._TowerDiagonalBoltSize
        End Get
        Set
            Me._TowerDiagonalBoltSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towertopgirtboltgrade")>
    <DataMember()> Public Property TowerTopGirtBoltGrade() As String
        Get
            Return Me._TowerTopGirtBoltGrade
        End Get
        Set
            Me._TowerTopGirtBoltGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towertopgirtboltsize")>
    <DataMember()> Public Property TowerTopGirtBoltSize() As Double?
        Get
            Return Me._TowerTopGirtBoltSize
        End Get
        Set
            Me._TowerTopGirtBoltSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerbotgirtboltgrade")>
    <DataMember()> Public Property TowerBotGirtBoltGrade() As String
        Get
            Return Me._TowerBotGirtBoltGrade
        End Get
        Set
            Me._TowerBotGirtBoltGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerbotgirtboltsize")>
    <DataMember()> Public Property TowerBotGirtBoltSize() As Double?
        Get
            Return Me._TowerBotGirtBoltSize
        End Get
        Set
            Me._TowerBotGirtBoltSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnergirtboltgrade")>
    <DataMember()> Public Property TowerInnerGirtBoltGrade() As String
        Get
            Return Me._TowerInnerGirtBoltGrade
        End Get
        Set
            Me._TowerInnerGirtBoltGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnergirtboltsize")>
    <DataMember()> Public Property TowerInnerGirtBoltSize() As Double?
        Get
            Return Me._TowerInnerGirtBoltSize
        End Get
        Set
            Me._TowerInnerGirtBoltSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerhorizontalboltgrade")>
    <DataMember()> Public Property TowerHorizontalBoltGrade() As String
        Get
            Return Me._TowerHorizontalBoltGrade
        End Get
        Set
            Me._TowerHorizontalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerhorizontalboltsize")>
    <DataMember()> Public Property TowerHorizontalBoltSize() As Double?
        Get
            Return Me._TowerHorizontalBoltSize
        End Get
        Set
            Me._TowerHorizontalBoltSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towershorthorizontalboltgrade")>
    <DataMember()> Public Property TowerShortHorizontalBoltGrade() As String
        Get
            Return Me._TowerShortHorizontalBoltGrade
        End Get
        Set
            Me._TowerShortHorizontalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towershorthorizontalboltsize")>
    <DataMember()> Public Property TowerShortHorizontalBoltSize() As Double?
        Get
            Return Me._TowerShortHorizontalBoltSize
        End Get
        Set
            Me._TowerShortHorizontalBoltSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerlegboltedgedistance")>
    <DataMember()> Public Property TowerLegBoltEdgeDistance() As Double?
        Get
            Return Me._TowerLegBoltEdgeDistance
        End Get
        Set
            Me._TowerLegBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagonalboltedgedistance")>
    <DataMember()> Public Property TowerDiagonalBoltEdgeDistance() As Double?
        Get
            Return Me._TowerDiagonalBoltEdgeDistance
        End Get
        Set
            Me._TowerDiagonalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towertopgirtboltedgedistance")>
    <DataMember()> Public Property TowerTopGirtBoltEdgeDistance() As Double?
        Get
            Return Me._TowerTopGirtBoltEdgeDistance
        End Get
        Set
            Me._TowerTopGirtBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerbotgirtboltedgedistance")>
    <DataMember()> Public Property TowerBotGirtBoltEdgeDistance() As Double?
        Get
            Return Me._TowerBotGirtBoltEdgeDistance
        End Get
        Set
            Me._TowerBotGirtBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnergirtboltedgedistance")>
    <DataMember()> Public Property TowerInnerGirtBoltEdgeDistance() As Double?
        Get
            Return Me._TowerInnerGirtBoltEdgeDistance
        End Get
        Set
            Me._TowerInnerGirtBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerhorizontalboltedgedistance")>
    <DataMember()> Public Property TowerHorizontalBoltEdgeDistance() As Double?
        Get
            Return Me._TowerHorizontalBoltEdgeDistance
        End Get
        Set
            Me._TowerHorizontalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towershorthorizontalboltedgedistance")>
    <DataMember()> Public Property TowerShortHorizontalBoltEdgeDistance() As Double?
        Get
            Return Me._TowerShortHorizontalBoltEdgeDistance
        End Get
        Set
            Me._TowerShortHorizontalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagonalgageg1Distance")>
    <DataMember()> Public Property TowerDiagonalGageG1Distance() As Double?
        Get
            Return Me._TowerDiagonalGageG1Distance
        End Get
        Set
            Me._TowerDiagonalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towertopgirtgageg1Distance")>
    <DataMember()> Public Property TowerTopGirtGageG1Distance() As Double?
        Get
            Return Me._TowerTopGirtGageG1Distance
        End Get
        Set
            Me._TowerTopGirtGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerbotgirtgageg1Distance")>
    <DataMember()> Public Property TowerBotGirtGageG1Distance() As Double?
        Get
            Return Me._TowerBotGirtGageG1Distance
        End Get
        Set
            Me._TowerBotGirtGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnergirtgageg1Distance")>
    <DataMember()> Public Property TowerInnerGirtGageG1Distance() As Double?
        Get
            Return Me._TowerInnerGirtGageG1Distance
        End Get
        Set
            Me._TowerInnerGirtGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerhorizontalgageg1Distance")>
    <DataMember()> Public Property TowerHorizontalGageG1Distance() As Double?
        Get
            Return Me._TowerHorizontalGageG1Distance
        End Get
        Set
            Me._TowerHorizontalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towershorthorizontalgageg1Distance")>
    <DataMember()> Public Property TowerShortHorizontalGageG1Distance() As Double?
        Get
            Return Me._TowerShortHorizontalGageG1Distance
        End Get
        Set
            Me._TowerShortHorizontalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthorizontalboltgrade")>
    <DataMember()> Public Property TowerRedundantHorizontalBoltGrade() As String
        Get
            Return Me._TowerRedundantHorizontalBoltGrade
        End Get
        Set
            Me._TowerRedundantHorizontalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthorizontalboltsize")>
    <DataMember()> Public Property TowerRedundantHorizontalBoltSize() As Double?
        Get
            Return Me._TowerRedundantHorizontalBoltSize
        End Get
        Set
            Me._TowerRedundantHorizontalBoltSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthorizontalnumbolts")>
    <DataMember()> Public Property TowerRedundantHorizontalNumBolts() As Integer?
        Get
            Return Me._TowerRedundantHorizontalNumBolts
        End Get
        Set
            Me._TowerRedundantHorizontalNumBolts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthorizontalboltedgedistance")>
    <DataMember()> Public Property TowerRedundantHorizontalBoltEdgeDistance() As Double?
        Get
            Return Me._TowerRedundantHorizontalBoltEdgeDistance
        End Get
        Set
            Me._TowerRedundantHorizontalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthorizontalgageg1Distance")>
    <DataMember()> Public Property TowerRedundantHorizontalGageG1Distance() As Double?
        Get
            Return Me._TowerRedundantHorizontalGageG1Distance
        End Get
        Set
            Me._TowerRedundantHorizontalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthorizontalnetwidthdeduct")>
    <DataMember()> Public Property TowerRedundantHorizontalNetWidthDeduct() As Double?
        Get
            Return Me._TowerRedundantHorizontalNetWidthDeduct
        End Get
        Set
            Me._TowerRedundantHorizontalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthorizontalufactor")>
    <DataMember()> Public Property TowerRedundantHorizontalUFactor() As Double?
        Get
            Return Me._TowerRedundantHorizontalUFactor
        End Get
        Set
            Me._TowerRedundantHorizontalUFactor = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantdiagonalboltgrade")>
    <DataMember()> Public Property TowerRedundantDiagonalBoltGrade() As String
        Get
            Return Me._TowerRedundantDiagonalBoltGrade
        End Get
        Set
            Me._TowerRedundantDiagonalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantdiagonalboltsize")>
    <DataMember()> Public Property TowerRedundantDiagonalBoltSize() As Double?
        Get
            Return Me._TowerRedundantDiagonalBoltSize
        End Get
        Set
            Me._TowerRedundantDiagonalBoltSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantdiagonalnumbolts")>
    <DataMember()> Public Property TowerRedundantDiagonalNumBolts() As Integer?
        Get
            Return Me._TowerRedundantDiagonalNumBolts
        End Get
        Set
            Me._TowerRedundantDiagonalNumBolts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantdiagonalboltedgedistance")>
    <DataMember()> Public Property TowerRedundantDiagonalBoltEdgeDistance() As Double?
        Get
            Return Me._TowerRedundantDiagonalBoltEdgeDistance
        End Get
        Set
            Me._TowerRedundantDiagonalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantdiagonalgageg1Distance")>
    <DataMember()> Public Property TowerRedundantDiagonalGageG1Distance() As Double?
        Get
            Return Me._TowerRedundantDiagonalGageG1Distance
        End Get
        Set
            Me._TowerRedundantDiagonalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantdiagonalnetwidthdeduct")>
    <DataMember()> Public Property TowerRedundantDiagonalNetWidthDeduct() As Double?
        Get
            Return Me._TowerRedundantDiagonalNetWidthDeduct
        End Get
        Set
            Me._TowerRedundantDiagonalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantdiagonalufactor")>
    <DataMember()> Public Property TowerRedundantDiagonalUFactor() As Double?
        Get
            Return Me._TowerRedundantDiagonalUFactor
        End Get
        Set
            Me._TowerRedundantDiagonalUFactor = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubdiagonalboltgrade")>
    <DataMember()> Public Property TowerRedundantSubDiagonalBoltGrade() As String
        Get
            Return Me._TowerRedundantSubDiagonalBoltGrade
        End Get
        Set
            Me._TowerRedundantSubDiagonalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubdiagonalboltsize")>
    <DataMember()> Public Property TowerRedundantSubDiagonalBoltSize() As Double?
        Get
            Return Me._TowerRedundantSubDiagonalBoltSize
        End Get
        Set
            Me._TowerRedundantSubDiagonalBoltSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubdiagonalnumbolts")>
    <DataMember()> Public Property TowerRedundantSubDiagonalNumBolts() As Integer?
        Get
            Return Me._TowerRedundantSubDiagonalNumBolts
        End Get
        Set
            Me._TowerRedundantSubDiagonalNumBolts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubdiagonalboltedgedistance")>
    <DataMember()> Public Property TowerRedundantSubDiagonalBoltEdgeDistance() As Double?
        Get
            Return Me._TowerRedundantSubDiagonalBoltEdgeDistance
        End Get
        Set
            Me._TowerRedundantSubDiagonalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubdiagonalgageg1Distance")>
    <DataMember()> Public Property TowerRedundantSubDiagonalGageG1Distance() As Double?
        Get
            Return Me._TowerRedundantSubDiagonalGageG1Distance
        End Get
        Set
            Me._TowerRedundantSubDiagonalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubdiagonalnetwidthdeduct")>
    <DataMember()> Public Property TowerRedundantSubDiagonalNetWidthDeduct() As Double?
        Get
            Return Me._TowerRedundantSubDiagonalNetWidthDeduct
        End Get
        Set
            Me._TowerRedundantSubDiagonalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubdiagonalufactor")>
    <DataMember()> Public Property TowerRedundantSubDiagonalUFactor() As Double?
        Get
            Return Me._TowerRedundantSubDiagonalUFactor
        End Get
        Set
            Me._TowerRedundantSubDiagonalUFactor = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubhorizontalboltgrade")>
    <DataMember()> Public Property TowerRedundantSubHorizontalBoltGrade() As String
        Get
            Return Me._TowerRedundantSubHorizontalBoltGrade
        End Get
        Set
            Me._TowerRedundantSubHorizontalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubhorizontalboltsize")>
    <DataMember()> Public Property TowerRedundantSubHorizontalBoltSize() As Double?
        Get
            Return Me._TowerRedundantSubHorizontalBoltSize
        End Get
        Set
            Me._TowerRedundantSubHorizontalBoltSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubhorizontalnumbolts")>
    <DataMember()> Public Property TowerRedundantSubHorizontalNumBolts() As Integer?
        Get
            Return Me._TowerRedundantSubHorizontalNumBolts
        End Get
        Set
            Me._TowerRedundantSubHorizontalNumBolts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubhorizontalboltedgedistance")>
    <DataMember()> Public Property TowerRedundantSubHorizontalBoltEdgeDistance() As Double?
        Get
            Return Me._TowerRedundantSubHorizontalBoltEdgeDistance
        End Get
        Set
            Me._TowerRedundantSubHorizontalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubhorizontalgageg1Distance")>
    <DataMember()> Public Property TowerRedundantSubHorizontalGageG1Distance() As Double?
        Get
            Return Me._TowerRedundantSubHorizontalGageG1Distance
        End Get
        Set
            Me._TowerRedundantSubHorizontalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubhorizontalnetwidthdeduct")>
    <DataMember()> Public Property TowerRedundantSubHorizontalNetWidthDeduct() As Double?
        Get
            Return Me._TowerRedundantSubHorizontalNetWidthDeduct
        End Get
        Set
            Me._TowerRedundantSubHorizontalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubhorizontalufactor")>
    <DataMember()> Public Property TowerRedundantSubHorizontalUFactor() As Double?
        Get
            Return Me._TowerRedundantSubHorizontalUFactor
        End Get
        Set
            Me._TowerRedundantSubHorizontalUFactor = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantverticalboltgrade")>
    <DataMember()> Public Property TowerRedundantVerticalBoltGrade() As String
        Get
            Return Me._TowerRedundantVerticalBoltGrade
        End Get
        Set
            Me._TowerRedundantVerticalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantverticalboltsize")>
    <DataMember()> Public Property TowerRedundantVerticalBoltSize() As Double?
        Get
            Return Me._TowerRedundantVerticalBoltSize
        End Get
        Set
            Me._TowerRedundantVerticalBoltSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantverticalnumbolts")>
    <DataMember()> Public Property TowerRedundantVerticalNumBolts() As Integer?
        Get
            Return Me._TowerRedundantVerticalNumBolts
        End Get
        Set
            Me._TowerRedundantVerticalNumBolts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantverticalboltedgedistance")>
    <DataMember()> Public Property TowerRedundantVerticalBoltEdgeDistance() As Double?
        Get
            Return Me._TowerRedundantVerticalBoltEdgeDistance
        End Get
        Set
            Me._TowerRedundantVerticalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantverticalgageg1Distance")>
    <DataMember()> Public Property TowerRedundantVerticalGageG1Distance() As Double?
        Get
            Return Me._TowerRedundantVerticalGageG1Distance
        End Get
        Set
            Me._TowerRedundantVerticalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantverticalnetwidthdeduct")>
    <DataMember()> Public Property TowerRedundantVerticalNetWidthDeduct() As Double?
        Get
            Return Me._TowerRedundantVerticalNetWidthDeduct
        End Get
        Set
            Me._TowerRedundantVerticalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantverticalufactor")>
    <DataMember()> Public Property TowerRedundantVerticalUFactor() As Double?
        Get
            Return Me._TowerRedundantVerticalUFactor
        End Get
        Set
            Me._TowerRedundantVerticalUFactor = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipboltgrade")>
    <DataMember()> Public Property TowerRedundantHipBoltGrade() As String
        Get
            Return Me._TowerRedundantHipBoltGrade
        End Get
        Set
            Me._TowerRedundantHipBoltGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipboltsize")>
    <DataMember()> Public Property TowerRedundantHipBoltSize() As Double?
        Get
            Return Me._TowerRedundantHipBoltSize
        End Get
        Set
            Me._TowerRedundantHipBoltSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipnumbolts")>
    <DataMember()> Public Property TowerRedundantHipNumBolts() As Integer?
        Get
            Return Me._TowerRedundantHipNumBolts
        End Get
        Set
            Me._TowerRedundantHipNumBolts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipboltedgedistance")>
    <DataMember()> Public Property TowerRedundantHipBoltEdgeDistance() As Double?
        Get
            Return Me._TowerRedundantHipBoltEdgeDistance
        End Get
        Set
            Me._TowerRedundantHipBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipgageg1Distance")>
    <DataMember()> Public Property TowerRedundantHipGageG1Distance() As Double?
        Get
            Return Me._TowerRedundantHipGageG1Distance
        End Get
        Set
            Me._TowerRedundantHipGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipnetwidthdeduct")>
    <DataMember()> Public Property TowerRedundantHipNetWidthDeduct() As Double?
        Get
            Return Me._TowerRedundantHipNetWidthDeduct
        End Get
        Set
            Me._TowerRedundantHipNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipufactor")>
    <DataMember()> Public Property TowerRedundantHipUFactor() As Double?
        Get
            Return Me._TowerRedundantHipUFactor
        End Get
        Set
            Me._TowerRedundantHipUFactor = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipdiagonalboltgrade")>
    <DataMember()> Public Property TowerRedundantHipDiagonalBoltGrade() As String
        Get
            Return Me._TowerRedundantHipDiagonalBoltGrade
        End Get
        Set
            Me._TowerRedundantHipDiagonalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipdiagonalboltsize")>
    <DataMember()> Public Property TowerRedundantHipDiagonalBoltSize() As Double?
        Get
            Return Me._TowerRedundantHipDiagonalBoltSize
        End Get
        Set
            Me._TowerRedundantHipDiagonalBoltSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipdiagonalnumbolts")>
    <DataMember()> Public Property TowerRedundantHipDiagonalNumBolts() As Integer?
        Get
            Return Me._TowerRedundantHipDiagonalNumBolts
        End Get
        Set
            Me._TowerRedundantHipDiagonalNumBolts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipdiagonalboltedgedistance")>
    <DataMember()> Public Property TowerRedundantHipDiagonalBoltEdgeDistance() As Double?
        Get
            Return Me._TowerRedundantHipDiagonalBoltEdgeDistance
        End Get
        Set
            Me._TowerRedundantHipDiagonalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipdiagonalgageg1Distance")>
    <DataMember()> Public Property TowerRedundantHipDiagonalGageG1Distance() As Double?
        Get
            Return Me._TowerRedundantHipDiagonalGageG1Distance
        End Get
        Set
            Me._TowerRedundantHipDiagonalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipdiagonalnetwidthdeduct")>
    <DataMember()> Public Property TowerRedundantHipDiagonalNetWidthDeduct() As Double?
        Get
            Return Me._TowerRedundantHipDiagonalNetWidthDeduct
        End Get
        Set
            Me._TowerRedundantHipDiagonalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipdiagonalufactor")>
    <DataMember()> Public Property TowerRedundantHipDiagonalUFactor() As Double?
        Get
            Return Me._TowerRedundantHipDiagonalUFactor
        End Get
        Set
            Me._TowerRedundantHipDiagonalUFactor = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagonaloutofplanerestraint")>
    <DataMember()> Public Property TowerDiagonalOutOfPlaneRestraint() As Boolean?
        Get
            Return Me._TowerDiagonalOutOfPlaneRestraint
        End Get
        Set
            Me._TowerDiagonalOutOfPlaneRestraint = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towertopgirtoutofplanerestraint")>
    <DataMember()> Public Property TowerTopGirtOutOfPlaneRestraint() As Boolean?
        Get
            Return Me._TowerTopGirtOutOfPlaneRestraint
        End Get
        Set
            Me._TowerTopGirtOutOfPlaneRestraint = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerbottomgirtoutofplanerestraint")>
    <DataMember()> Public Property TowerBottomGirtOutOfPlaneRestraint() As Boolean?
        Get
            Return Me._TowerBottomGirtOutOfPlaneRestraint
        End Get
        Set
            Me._TowerBottomGirtOutOfPlaneRestraint = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towermidgirtoutofplanerestraint")>
    <DataMember()> Public Property TowerMidGirtOutOfPlaneRestraint() As Boolean?
        Get
            Return Me._TowerMidGirtOutOfPlaneRestraint
        End Get
        Set
            Me._TowerMidGirtOutOfPlaneRestraint = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerhorizontaloutofplanerestraint")>
    <DataMember()> Public Property TowerHorizontalOutOfPlaneRestraint() As Boolean?
        Get
            Return Me._TowerHorizontalOutOfPlaneRestraint
        End Get
        Set
            Me._TowerHorizontalOutOfPlaneRestraint = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towersecondaryhorizontaloutofplanerestraint")>
    <DataMember()> Public Property TowerSecondaryHorizontalOutOfPlaneRestraint() As Boolean?
        Get
            Return Me._TowerSecondaryHorizontalOutOfPlaneRestraint
        End Get
        Set
            Me._TowerSecondaryHorizontalOutOfPlaneRestraint = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Toweruniqueflag")>
    <DataMember()> Public Property TowerUniqueFlag() As Integer?
        Get
            Return Me._TowerUniqueFlag
        End Get
        Set
            Me._TowerUniqueFlag = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagoffsetney")>
    <DataMember()> Public Property TowerDiagOffsetNEY() As Double?
        Get
            Return Me._TowerDiagOffsetNEY
        End Get
        Set
            Me._TowerDiagOffsetNEY = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagoffsetnex")>
    <DataMember()> Public Property TowerDiagOffsetNEX() As Double?
        Get
            Return Me._TowerDiagOffsetNEX
        End Get
        Set
            Me._TowerDiagOffsetNEX = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagoffsetpey")>
    <DataMember()> Public Property TowerDiagOffsetPEY() As Double?
        Get
            Return Me._TowerDiagOffsetPEY
        End Get
        Set
            Me._TowerDiagOffsetPEY = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagoffsetpex")>
    <DataMember()> Public Property TowerDiagOffsetPEX() As Double?
        Get
            Return Me._TowerDiagOffsetPEX
        End Get
        Set
            Me._TowerDiagOffsetPEX = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkbraceoffsetney")>
    <DataMember()> Public Property TowerKbraceOffsetNEY() As Double?
        Get
            Return Me._TowerKbraceOffsetNEY
        End Get
        Set
            Me._TowerKbraceOffsetNEY = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkbraceoffsetnex")>
    <DataMember()> Public Property TowerKbraceOffsetNEX() As Double?
        Get
            Return Me._TowerKbraceOffsetNEX
        End Get
        Set
            Me._TowerKbraceOffsetNEX = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkbraceoffsetpey")>
    <DataMember()> Public Property TowerKbraceOffsetPEY() As Double?
        Get
            Return Me._TowerKbraceOffsetPEY
        End Get
        Set
            Me._TowerKbraceOffsetPEY = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkbraceoffsetpex")>
    <DataMember()> Public Property TowerKbraceOffsetPEX() As Double?
        Get
            Return Me._TowerKbraceOffsetPEX
        End Get
        Set
            Me._TowerKbraceOffsetPEX = Value
        End Set
    End Property
#End Region

#Region "Constructors"
    Public Sub New(Optional ByVal Parent As EDSObject = Nothing)
        If Parent IsNot Nothing Then Me.Absorb(Parent)
    End Sub

    Public Sub New(data As DataRow, Optional ByVal Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Me.ID = DBtoNullableInt(data.Item("ID"))
        Me.Rec = DBtoNullableInt(data.Item("TowerRec"))
        Me.TowerDatabase = DBtoStr(data.Item("TowerDatabase"))
        Me.TowerName = DBtoStr(data.Item("TowerName"))
        Me.TowerHeight = DBtoNullableDbl(data.Item("TowerHeight"))
        Me.TowerFaceWidth = DBtoNullableDbl(data.Item("TowerFaceWidth"))
        Me.TowerNumSections = DBtoNullableInt(data.Item("TowerNumSections"))
        Me.TowerSectionLength = DBtoNullableDbl(data.Item("TowerSectionLength"))
        Me.TowerDiagonalSpacing = DBtoNullableDbl(data.Item("TowerDiagonalSpacing"))
        Me.TowerDiagonalSpacingEx = DBtoNullableDbl(data.Item("TowerDiagonalSpacingEx"))
        Me.TowerBraceType = DBtoStr(data.Item("TowerBraceType"))
        Me.TowerFaceBevel = DBtoNullableDbl(data.Item("TowerFaceBevel"))
        Me.TowerTopGirtOffset = DBtoNullableDbl(data.Item("TowerTopGirtOffset"))
        Me.TowerBotGirtOffset = DBtoNullableDbl(data.Item("TowerBotGirtOffset"))
        Me.TowerHasKBraceEndPanels = DBtoNullableBool(data.Item("TowerHasKBraceEndPanels"))
        Me.TowerHasHorizontals = DBtoNullableBool(data.Item("TowerHasHorizontals"))
        Me.TowerLegType = DBtoStr(data.Item("TowerLegType"))
        Me.TowerLegSize = DBtoStr(data.Item("TowerLegSize"))
        Me.TowerLegGrade = DBtoNullableDbl(data.Item("TowerLegGrade"))
        Me.TowerLegMatlGrade = DBtoStr(data.Item("TowerLegMatlGrade"))
        Me.TowerDiagonalGrade = DBtoNullableDbl(data.Item("TowerDiagonalGrade"))
        Me.TowerDiagonalMatlGrade = DBtoStr(data.Item("TowerDiagonalMatlGrade"))
        Me.TowerInnerBracingGrade = DBtoNullableDbl(data.Item("TowerInnerBracingGrade"))
        Me.TowerInnerBracingMatlGrade = DBtoStr(data.Item("TowerInnerBracingMatlGrade"))
        Me.TowerTopGirtGrade = DBtoNullableDbl(data.Item("TowerTopGirtGrade"))
        Me.TowerTopGirtMatlGrade = DBtoStr(data.Item("TowerTopGirtMatlGrade"))
        Me.TowerBotGirtGrade = DBtoNullableDbl(data.Item("TowerBotGirtGrade"))
        Me.TowerBotGirtMatlGrade = DBtoStr(data.Item("TowerBotGirtMatlGrade"))
        Me.TowerInnerGirtGrade = DBtoNullableDbl(data.Item("TowerInnerGirtGrade"))
        Me.TowerInnerGirtMatlGrade = DBtoStr(data.Item("TowerInnerGirtMatlGrade"))
        Me.TowerLongHorizontalGrade = DBtoNullableDbl(data.Item("TowerLongHorizontalGrade"))
        Me.TowerLongHorizontalMatlGrade = DBtoStr(data.Item("TowerLongHorizontalMatlGrade"))
        Me.TowerShortHorizontalGrade = DBtoNullableDbl(data.Item("TowerShortHorizontalGrade"))
        Me.TowerShortHorizontalMatlGrade = DBtoStr(data.Item("TowerShortHorizontalMatlGrade"))
        Me.TowerDiagonalType = DBtoStr(data.Item("TowerDiagonalType"))
        Me.TowerDiagonalSize = DBtoStr(data.Item("TowerDiagonalSize"))
        Me.TowerInnerBracingType = DBtoStr(data.Item("TowerInnerBracingType"))
        Me.TowerInnerBracingSize = DBtoStr(data.Item("TowerInnerBracingSize"))
        Me.TowerTopGirtType = DBtoStr(data.Item("TowerTopGirtType"))
        Me.TowerTopGirtSize = DBtoStr(data.Item("TowerTopGirtSize"))
        Me.TowerBotGirtType = DBtoStr(data.Item("TowerBotGirtType"))
        Me.TowerBotGirtSize = DBtoStr(data.Item("TowerBotGirtSize"))
        Me.TowerNumInnerGirts = DBtoNullableInt(data.Item("TowerNumInnerGirts"))
        Me.TowerInnerGirtType = DBtoStr(data.Item("TowerInnerGirtType"))
        Me.TowerInnerGirtSize = DBtoStr(data.Item("TowerInnerGirtSize"))
        Me.TowerLongHorizontalType = DBtoStr(data.Item("TowerLongHorizontalType"))
        Me.TowerLongHorizontalSize = DBtoStr(data.Item("TowerLongHorizontalSize"))
        Me.TowerShortHorizontalType = DBtoStr(data.Item("TowerShortHorizontalType"))
        Me.TowerShortHorizontalSize = DBtoStr(data.Item("TowerShortHorizontalSize"))
        Me.TowerRedundantGrade = DBtoNullableDbl(data.Item("TowerRedundantGrade"))
        Me.TowerRedundantMatlGrade = DBtoStr(data.Item("TowerRedundantMatlGrade"))
        Me.TowerRedundantType = DBtoStr(data.Item("TowerRedundantType"))
        Me.TowerRedundantDiagType = DBtoStr(data.Item("TowerRedundantDiagType"))
        Me.TowerRedundantSubDiagonalType = DBtoStr(data.Item("TowerRedundantSubDiagonalType"))
        Me.TowerRedundantSubHorizontalType = DBtoStr(data.Item("TowerRedundantSubHorizontalType"))
        Me.TowerRedundantVerticalType = DBtoStr(data.Item("TowerRedundantVerticalType"))
        Me.TowerRedundantHipType = DBtoStr(data.Item("TowerRedundantHipType"))
        Me.TowerRedundantHipDiagonalType = DBtoStr(data.Item("TowerRedundantHipDiagonalType"))
        Me.TowerRedundantHorizontalSize = DBtoStr(data.Item("TowerRedundantHorizontalSize"))
        Me.TowerRedundantHorizontalSize2 = DBtoStr(data.Item("TowerRedundantHorizontalSize2"))
        Me.TowerRedundantHorizontalSize3 = DBtoStr(data.Item("TowerRedundantHorizontalSize3"))
        Me.TowerRedundantHorizontalSize4 = DBtoStr(data.Item("TowerRedundantHorizontalSize4"))
        Me.TowerRedundantDiagonalSize = DBtoStr(data.Item("TowerRedundantDiagonalSize"))
        Me.TowerRedundantDiagonalSize2 = DBtoStr(data.Item("TowerRedundantDiagonalSize2"))
        Me.TowerRedundantDiagonalSize3 = DBtoStr(data.Item("TowerRedundantDiagonalSize3"))
        Me.TowerRedundantDiagonalSize4 = DBtoStr(data.Item("TowerRedundantDiagonalSize4"))
        Me.TowerRedundantSubHorizontalSize = DBtoStr(data.Item("TowerRedundantSubHorizontalSize"))
        Me.TowerRedundantSubDiagonalSize = DBtoStr(data.Item("TowerRedundantSubDiagonalSize"))
        Me.TowerSubDiagLocation = DBtoNullableDbl(data.Item("TowerSubDiagLocation"))
        Me.TowerRedundantVerticalSize = DBtoStr(data.Item("TowerRedundantVerticalSize"))
        Me.TowerRedundantHipSize = DBtoStr(data.Item("TowerRedundantHipSize"))
        Me.TowerRedundantHipSize2 = DBtoStr(data.Item("TowerRedundantHipSize2"))
        Me.TowerRedundantHipSize3 = DBtoStr(data.Item("TowerRedundantHipSize3"))
        Me.TowerRedundantHipSize4 = DBtoStr(data.Item("TowerRedundantHipSize4"))
        Me.TowerRedundantHipDiagonalSize = DBtoStr(data.Item("TowerRedundantHipDiagonalSize"))
        Me.TowerRedundantHipDiagonalSize2 = DBtoStr(data.Item("TowerRedundantHipDiagonalSize2"))
        Me.TowerRedundantHipDiagonalSize3 = DBtoStr(data.Item("TowerRedundantHipDiagonalSize3"))
        Me.TowerRedundantHipDiagonalSize4 = DBtoStr(data.Item("TowerRedundantHipDiagonalSize4"))
        Me.TowerSWMult = DBtoNullableDbl(data.Item("TowerSWMult"))
        Me.TowerWPMult = DBtoNullableDbl(data.Item("TowerWPMult"))
        Me.TowerAutoCalcKSingleAngle = DBtoNullableBool(data.Item("TowerAutoCalcKSingleAngle"))
        Me.TowerAutoCalcKSolidRound = DBtoNullableBool(data.Item("TowerAutoCalcKSolidRound"))
        Me.TowerAfGusset = DBtoNullableDbl(data.Item("TowerAfGusset"))
        Me.TowerTfGusset = DBtoNullableDbl(data.Item("TowerTfGusset"))
        Me.TowerGussetBoltEdgeDistance = DBtoNullableDbl(data.Item("TowerGussetBoltEdgeDistance"))
        Me.TowerGussetGrade = DBtoNullableDbl(data.Item("TowerGussetGrade"))
        Me.TowerGussetMatlGrade = DBtoStr(data.Item("TowerGussetMatlGrade"))
        Me.TowerAfMult = DBtoNullableDbl(data.Item("TowerAfMult"))
        Me.TowerArMult = DBtoNullableDbl(data.Item("TowerArMult"))
        Me.TowerFlatIPAPole = DBtoNullableDbl(data.Item("TowerFlatIPAPole"))
        Me.TowerRoundIPAPole = DBtoNullableDbl(data.Item("TowerRoundIPAPole"))
        Me.TowerFlatIPALeg = DBtoNullableDbl(data.Item("TowerFlatIPALeg"))
        Me.TowerRoundIPALeg = DBtoNullableDbl(data.Item("TowerRoundIPALeg"))
        Me.TowerFlatIPAHorizontal = DBtoNullableDbl(data.Item("TowerFlatIPAHorizontal"))
        Me.TowerRoundIPAHorizontal = DBtoNullableDbl(data.Item("TowerRoundIPAHorizontal"))
        Me.TowerFlatIPADiagonal = DBtoNullableDbl(data.Item("TowerFlatIPADiagonal"))
        Me.TowerRoundIPADiagonal = DBtoNullableDbl(data.Item("TowerRoundIPADiagonal"))
        Me.TowerCSA_S37_SpeedUpFactor = DBtoNullableDbl(data.Item("TowerCSA_S37_SpeedUpFactor"))
        Me.TowerKLegs = DBtoNullableDbl(data.Item("TowerKLegs"))
        Me.TowerKXBracedDiags = DBtoNullableDbl(data.Item("TowerKXBracedDiags"))
        Me.TowerKKBracedDiags = DBtoNullableDbl(data.Item("TowerKKBracedDiags"))
        Me.TowerKZBracedDiags = DBtoNullableDbl(data.Item("TowerKZBracedDiags"))
        Me.TowerKHorzs = DBtoNullableDbl(data.Item("TowerKHorzs"))
        Me.TowerKSecHorzs = DBtoNullableDbl(data.Item("TowerKSecHorzs"))
        Me.TowerKGirts = DBtoNullableDbl(data.Item("TowerKGirts"))
        Me.TowerKInners = DBtoNullableDbl(data.Item("TowerKInners"))
        Me.TowerKXBracedDiagsY = DBtoNullableDbl(data.Item("TowerKXBracedDiagsY"))
        Me.TowerKKBracedDiagsY = DBtoNullableDbl(data.Item("TowerKKBracedDiagsY"))
        Me.TowerKZBracedDiagsY = DBtoNullableDbl(data.Item("TowerKZBracedDiagsY"))
        Me.TowerKHorzsY = DBtoNullableDbl(data.Item("TowerKHorzsY"))
        Me.TowerKSecHorzsY = DBtoNullableDbl(data.Item("TowerKSecHorzsY"))
        Me.TowerKGirtsY = DBtoNullableDbl(data.Item("TowerKGirtsY"))
        Me.TowerKInnersY = DBtoNullableDbl(data.Item("TowerKInnersY"))
        Me.TowerKRedHorz = DBtoNullableDbl(data.Item("TowerKRedHorz"))
        Me.TowerKRedDiag = DBtoNullableDbl(data.Item("TowerKRedDiag"))
        Me.TowerKRedSubDiag = DBtoNullableDbl(data.Item("TowerKRedSubDiag"))
        Me.TowerKRedSubHorz = DBtoNullableDbl(data.Item("TowerKRedSubHorz"))
        Me.TowerKRedVert = DBtoNullableDbl(data.Item("TowerKRedVert"))
        Me.TowerKRedHip = DBtoNullableDbl(data.Item("TowerKRedHip"))
        Me.TowerKRedHipDiag = DBtoNullableDbl(data.Item("TowerKRedHipDiag"))
        Me.TowerKTLX = DBtoNullableDbl(data.Item("TowerKTLX"))
        Me.TowerKTLZ = DBtoNullableDbl(data.Item("TowerKTLZ"))
        Me.TowerKTLLeg = DBtoNullableDbl(data.Item("TowerKTLLeg"))
        Me.TowerInnerKTLX = DBtoNullableDbl(data.Item("TowerInnerKTLX"))
        Me.TowerInnerKTLZ = DBtoNullableDbl(data.Item("TowerInnerKTLZ"))
        Me.TowerInnerKTLLeg = DBtoNullableDbl(data.Item("TowerInnerKTLLeg"))
        Me.TowerStitchBoltLocationHoriz = DBtoStr(data.Item("TowerStitchBoltLocationHoriz"))
        Me.TowerStitchBoltLocationDiag = DBtoStr(data.Item("TowerStitchBoltLocationDiag"))
        Me.TowerStitchBoltLocationRed = DBtoStr(data.Item("TowerStitchBoltLocationRed"))
        Me.TowerStitchSpacing = DBtoNullableDbl(data.Item("TowerStitchSpacing"))
        Me.TowerStitchSpacingDiag = DBtoNullableDbl(data.Item("TowerStitchSpacingDiag"))
        Me.TowerStitchSpacingHorz = DBtoNullableDbl(data.Item("TowerStitchSpacingHorz"))
        Me.TowerStitchSpacingRed = DBtoNullableDbl(data.Item("TowerStitchSpacingRed"))
        Me.TowerLegNetWidthDeduct = DBtoNullableDbl(data.Item("TowerLegNetWidthDeduct"))
        Me.TowerLegUFactor = DBtoNullableDbl(data.Item("TowerLegUFactor"))
        Me.TowerDiagonalNetWidthDeduct = DBtoNullableDbl(data.Item("TowerDiagonalNetWidthDeduct"))
        Me.TowerTopGirtNetWidthDeduct = DBtoNullableDbl(data.Item("TowerTopGirtNetWidthDeduct"))
        Me.TowerBotGirtNetWidthDeduct = DBtoNullableDbl(data.Item("TowerBotGirtNetWidthDeduct"))
        Me.TowerInnerGirtNetWidthDeduct = DBtoNullableDbl(data.Item("TowerInnerGirtNetWidthDeduct"))
        Me.TowerHorizontalNetWidthDeduct = DBtoNullableDbl(data.Item("TowerHorizontalNetWidthDeduct"))
        Me.TowerShortHorizontalNetWidthDeduct = DBtoNullableDbl(data.Item("TowerShortHorizontalNetWidthDeduct"))
        Me.TowerDiagonalUFactor = DBtoNullableDbl(data.Item("TowerDiagonalUFactor"))
        Me.TowerTopGirtUFactor = DBtoNullableDbl(data.Item("TowerTopGirtUFactor"))
        Me.TowerBotGirtUFactor = DBtoNullableDbl(data.Item("TowerBotGirtUFactor"))
        Me.TowerInnerGirtUFactor = DBtoNullableDbl(data.Item("TowerInnerGirtUFactor"))
        Me.TowerHorizontalUFactor = DBtoNullableDbl(data.Item("TowerHorizontalUFactor"))
        Me.TowerShortHorizontalUFactor = DBtoNullableDbl(data.Item("TowerShortHorizontalUFactor"))
        Me.TowerLegConnType = DBtoStr(data.Item("TowerLegConnType"))
        Me.TowerLegNumBolts = DBtoNullableInt(data.Item("TowerLegNumBolts"))
        Me.TowerDiagonalNumBolts = DBtoNullableInt(data.Item("TowerDiagonalNumBolts"))
        Me.TowerTopGirtNumBolts = DBtoNullableInt(data.Item("TowerTopGirtNumBolts"))
        Me.TowerBotGirtNumBolts = DBtoNullableInt(data.Item("TowerBotGirtNumBolts"))
        Me.TowerInnerGirtNumBolts = DBtoNullableInt(data.Item("TowerInnerGirtNumBolts"))
        Me.TowerHorizontalNumBolts = DBtoNullableInt(data.Item("TowerHorizontalNumBolts"))
        Me.TowerShortHorizontalNumBolts = DBtoNullableInt(data.Item("TowerShortHorizontalNumBolts"))
        Me.TowerLegBoltGrade = DBtoStr(data.Item("TowerLegBoltGrade"))
        Me.TowerLegBoltSize = DBtoNullableDbl(data.Item("TowerLegBoltSize"))
        Me.TowerDiagonalBoltGrade = DBtoStr(data.Item("TowerDiagonalBoltGrade"))
        Me.TowerDiagonalBoltSize = DBtoNullableDbl(data.Item("TowerDiagonalBoltSize"))
        Me.TowerTopGirtBoltGrade = DBtoStr(data.Item("TowerTopGirtBoltGrade"))
        Me.TowerTopGirtBoltSize = DBtoNullableDbl(data.Item("TowerTopGirtBoltSize"))
        Me.TowerBotGirtBoltGrade = DBtoStr(data.Item("TowerBotGirtBoltGrade"))
        Me.TowerBotGirtBoltSize = DBtoNullableDbl(data.Item("TowerBotGirtBoltSize"))
        Me.TowerInnerGirtBoltGrade = DBtoStr(data.Item("TowerInnerGirtBoltGrade"))
        Me.TowerInnerGirtBoltSize = DBtoNullableDbl(data.Item("TowerInnerGirtBoltSize"))
        Me.TowerHorizontalBoltGrade = DBtoStr(data.Item("TowerHorizontalBoltGrade"))
        Me.TowerHorizontalBoltSize = DBtoNullableDbl(data.Item("TowerHorizontalBoltSize"))
        Me.TowerShortHorizontalBoltGrade = DBtoStr(data.Item("TowerShortHorizontalBoltGrade"))
        Me.TowerShortHorizontalBoltSize = DBtoNullableDbl(data.Item("TowerShortHorizontalBoltSize"))
        Me.TowerLegBoltEdgeDistance = DBtoNullableDbl(data.Item("TowerLegBoltEdgeDistance"))
        Me.TowerDiagonalBoltEdgeDistance = DBtoNullableDbl(data.Item("TowerDiagonalBoltEdgeDistance"))
        Me.TowerTopGirtBoltEdgeDistance = DBtoNullableDbl(data.Item("TowerTopGirtBoltEdgeDistance"))
        Me.TowerBotGirtBoltEdgeDistance = DBtoNullableDbl(data.Item("TowerBotGirtBoltEdgeDistance"))
        Me.TowerInnerGirtBoltEdgeDistance = DBtoNullableDbl(data.Item("TowerInnerGirtBoltEdgeDistance"))
        Me.TowerHorizontalBoltEdgeDistance = DBtoNullableDbl(data.Item("TowerHorizontalBoltEdgeDistance"))
        Me.TowerShortHorizontalBoltEdgeDistance = DBtoNullableDbl(data.Item("TowerShortHorizontalBoltEdgeDistance"))
        Me.TowerDiagonalGageG1Distance = DBtoNullableDbl(data.Item("TowerDiagonalGageG1Distance"))
        Me.TowerTopGirtGageG1Distance = DBtoNullableDbl(data.Item("TowerTopGirtGageG1Distance"))
        Me.TowerBotGirtGageG1Distance = DBtoNullableDbl(data.Item("TowerBotGirtGageG1Distance"))
        Me.TowerInnerGirtGageG1Distance = DBtoNullableDbl(data.Item("TowerInnerGirtGageG1Distance"))
        Me.TowerHorizontalGageG1Distance = DBtoNullableDbl(data.Item("TowerHorizontalGageG1Distance"))
        Me.TowerShortHorizontalGageG1Distance = DBtoNullableDbl(data.Item("TowerShortHorizontalGageG1Distance"))
        Me.TowerRedundantHorizontalBoltGrade = DBtoStr(data.Item("TowerRedundantHorizontalBoltGrade"))
        Me.TowerRedundantHorizontalBoltSize = DBtoNullableDbl(data.Item("TowerRedundantHorizontalBoltSize"))
        Me.TowerRedundantHorizontalNumBolts = DBtoNullableInt(data.Item("TowerRedundantHorizontalNumBolts"))
        Me.TowerRedundantHorizontalBoltEdgeDistance = DBtoNullableDbl(data.Item("TowerRedundantHorizontalBoltEdgeDistance"))
        Me.TowerRedundantHorizontalGageG1Distance = DBtoNullableDbl(data.Item("TowerRedundantHorizontalGageG1Distance"))
        Me.TowerRedundantHorizontalNetWidthDeduct = DBtoNullableDbl(data.Item("TowerRedundantHorizontalNetWidthDeduct"))
        Me.TowerRedundantHorizontalUFactor = DBtoNullableDbl(data.Item("TowerRedundantHorizontalUFactor"))
        Me.TowerRedundantDiagonalBoltGrade = DBtoStr(data.Item("TowerRedundantDiagonalBoltGrade"))
        Me.TowerRedundantDiagonalBoltSize = DBtoNullableDbl(data.Item("TowerRedundantDiagonalBoltSize"))
        Me.TowerRedundantDiagonalNumBolts = DBtoNullableInt(data.Item("TowerRedundantDiagonalNumBolts"))
        Me.TowerRedundantDiagonalBoltEdgeDistance = DBtoNullableDbl(data.Item("TowerRedundantDiagonalBoltEdgeDistance"))
        Me.TowerRedundantDiagonalGageG1Distance = DBtoNullableDbl(data.Item("TowerRedundantDiagonalGageG1Distance"))
        Me.TowerRedundantDiagonalNetWidthDeduct = DBtoNullableDbl(data.Item("TowerRedundantDiagonalNetWidthDeduct"))
        Me.TowerRedundantDiagonalUFactor = DBtoNullableDbl(data.Item("TowerRedundantDiagonalUFactor"))
        Me.TowerRedundantSubDiagonalBoltGrade = DBtoStr(data.Item("TowerRedundantSubDiagonalBoltGrade"))
        Me.TowerRedundantSubDiagonalBoltSize = DBtoNullableDbl(data.Item("TowerRedundantSubDiagonalBoltSize"))
        Me.TowerRedundantSubDiagonalNumBolts = DBtoNullableInt(data.Item("TowerRedundantSubDiagonalNumBolts"))
        Me.TowerRedundantSubDiagonalBoltEdgeDistance = DBtoNullableDbl(data.Item("TowerRedundantSubDiagonalBoltEdgeDistance"))
        Me.TowerRedundantSubDiagonalGageG1Distance = DBtoNullableDbl(data.Item("TowerRedundantSubDiagonalGageG1Distance"))
        Me.TowerRedundantSubDiagonalNetWidthDeduct = DBtoNullableDbl(data.Item("TowerRedundantSubDiagonalNetWidthDeduct"))
        Me.TowerRedundantSubDiagonalUFactor = DBtoNullableDbl(data.Item("TowerRedundantSubDiagonalUFactor"))
        Me.TowerRedundantSubHorizontalBoltGrade = DBtoStr(data.Item("TowerRedundantSubHorizontalBoltGrade"))
        Me.TowerRedundantSubHorizontalBoltSize = DBtoNullableDbl(data.Item("TowerRedundantSubHorizontalBoltSize"))
        Me.TowerRedundantSubHorizontalNumBolts = DBtoNullableInt(data.Item("TowerRedundantSubHorizontalNumBolts"))
        Me.TowerRedundantSubHorizontalBoltEdgeDistance = DBtoNullableDbl(data.Item("TowerRedundantSubHorizontalBoltEdgeDistance"))
        Me.TowerRedundantSubHorizontalGageG1Distance = DBtoNullableDbl(data.Item("TowerRedundantSubHorizontalGageG1Distance"))
        Me.TowerRedundantSubHorizontalNetWidthDeduct = DBtoNullableDbl(data.Item("TowerRedundantSubHorizontalNetWidthDeduct"))
        Me.TowerRedundantSubHorizontalUFactor = DBtoNullableDbl(data.Item("TowerRedundantSubHorizontalUFactor"))
        Me.TowerRedundantVerticalBoltGrade = DBtoStr(data.Item("TowerRedundantVerticalBoltGrade"))
        Me.TowerRedundantVerticalBoltSize = DBtoNullableDbl(data.Item("TowerRedundantVerticalBoltSize"))
        Me.TowerRedundantVerticalNumBolts = DBtoNullableInt(data.Item("TowerRedundantVerticalNumBolts"))
        Me.TowerRedundantVerticalBoltEdgeDistance = DBtoNullableDbl(data.Item("TowerRedundantVerticalBoltEdgeDistance"))
        Me.TowerRedundantVerticalGageG1Distance = DBtoNullableDbl(data.Item("TowerRedundantVerticalGageG1Distance"))
        Me.TowerRedundantVerticalNetWidthDeduct = DBtoNullableDbl(data.Item("TowerRedundantVerticalNetWidthDeduct"))
        Me.TowerRedundantVerticalUFactor = DBtoNullableDbl(data.Item("TowerRedundantVerticalUFactor"))
        Me.TowerRedundantHipBoltGrade = DBtoStr(data.Item("TowerRedundantHipBoltGrade"))
        Me.TowerRedundantHipBoltSize = DBtoNullableDbl(data.Item("TowerRedundantHipBoltSize"))
        Me.TowerRedundantHipNumBolts = DBtoNullableInt(data.Item("TowerRedundantHipNumBolts"))
        Me.TowerRedundantHipBoltEdgeDistance = DBtoNullableDbl(data.Item("TowerRedundantHipBoltEdgeDistance"))
        Me.TowerRedundantHipGageG1Distance = DBtoNullableDbl(data.Item("TowerRedundantHipGageG1Distance"))
        Me.TowerRedundantHipNetWidthDeduct = DBtoNullableDbl(data.Item("TowerRedundantHipNetWidthDeduct"))
        Me.TowerRedundantHipUFactor = DBtoNullableDbl(data.Item("TowerRedundantHipUFactor"))
        Me.TowerRedundantHipDiagonalBoltGrade = DBtoStr(data.Item("TowerRedundantHipDiagonalBoltGrade"))
        Me.TowerRedundantHipDiagonalBoltSize = DBtoNullableDbl(data.Item("TowerRedundantHipDiagonalBoltSize"))
        Me.TowerRedundantHipDiagonalNumBolts = DBtoNullableInt(data.Item("TowerRedundantHipDiagonalNumBolts"))
        Me.TowerRedundantHipDiagonalBoltEdgeDistance = DBtoNullableDbl(data.Item("TowerRedundantHipDiagonalBoltEdgeDistance"))
        Me.TowerRedundantHipDiagonalGageG1Distance = DBtoNullableDbl(data.Item("TowerRedundantHipDiagonalGageG1Distance"))
        Me.TowerRedundantHipDiagonalNetWidthDeduct = DBtoNullableDbl(data.Item("TowerRedundantHipDiagonalNetWidthDeduct"))
        Me.TowerRedundantHipDiagonalUFactor = DBtoNullableDbl(data.Item("TowerRedundantHipDiagonalUFactor"))
        Me.TowerDiagonalOutOfPlaneRestraint = DBtoNullableBool(data.Item("TowerDiagonalOutOfPlaneRestraint"))
        Me.TowerTopGirtOutOfPlaneRestraint = DBtoNullableBool(data.Item("TowerTopGirtOutOfPlaneRestraint"))
        Me.TowerBottomGirtOutOfPlaneRestraint = DBtoNullableBool(data.Item("TowerBottomGirtOutOfPlaneRestraint"))
        Me.TowerMidGirtOutOfPlaneRestraint = DBtoNullableBool(data.Item("TowerMidGirtOutOfPlaneRestraint"))
        Me.TowerHorizontalOutOfPlaneRestraint = DBtoNullableBool(data.Item("TowerHorizontalOutOfPlaneRestraint"))
        Me.TowerSecondaryHorizontalOutOfPlaneRestraint = DBtoNullableBool(data.Item("TowerSecondaryHorizontalOutOfPlaneRestraint"))
        Me.TowerUniqueFlag = DBtoNullableInt(data.Item("TowerUniqueFlag"))
        Me.TowerDiagOffsetNEY = DBtoNullableDbl(data.Item("TowerDiagOffsetNEY"))
        Me.TowerDiagOffsetNEX = DBtoNullableDbl(data.Item("TowerDiagOffsetNEX"))
        Me.TowerDiagOffsetPEY = DBtoNullableDbl(data.Item("TowerDiagOffsetPEY"))
        Me.TowerDiagOffsetPEX = DBtoNullableDbl(data.Item("TowerDiagOffsetPEX"))
        Me.TowerKbraceOffsetNEY = DBtoNullableDbl(data.Item("TowerKbraceOffsetNEY"))
        Me.TowerKbraceOffsetNEX = DBtoNullableDbl(data.Item("TowerKbraceOffsetNEX"))
        Me.TowerKbraceOffsetPEY = DBtoNullableDbl(data.Item("TowerKbraceOffsetPEY"))
        Me.TowerKbraceOffsetPEX = DBtoNullableDbl(data.Item("TowerKbraceOffsetPEX"))

    End Sub
#End Region

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As tnxTowerRecord = TryCast(other, tnxTowerRecord)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.Rec.CheckChange(otherToCompare.Rec, changes, categoryName, "Tower Rec"), Equals, False)
        Equals = If(Me.TowerDatabase.CheckChange(otherToCompare.TowerDatabase, changes, categoryName, "Tower Database"), Equals, False)
        Equals = If(Me.TowerName.CheckChange(otherToCompare.TowerName, changes, categoryName, "Tower Name"), Equals, False)
        Equals = If(Me.TowerHeight.CheckChange(otherToCompare.TowerHeight, changes, categoryName, "Tower Height"), Equals, False)
        Equals = If(Me.TowerFaceWidth.CheckChange(otherToCompare.TowerFaceWidth, changes, categoryName, "Tower Face Width"), Equals, False)
        Equals = If(Me.TowerNumSections.CheckChange(otherToCompare.TowerNumSections, changes, categoryName, "Tower Num Sections"), Equals, False)
        Equals = If(Me.TowerSectionLength.CheckChange(otherToCompare.TowerSectionLength, changes, categoryName, "Tower Section Length"), Equals, False)
        Equals = If(Me.TowerDiagonalSpacing.CheckChange(otherToCompare.TowerDiagonalSpacing, changes, categoryName, "Tower Diagonal Spacing"), Equals, False)
        Equals = If(Me.TowerDiagonalSpacingEx.CheckChange(otherToCompare.TowerDiagonalSpacingEx, changes, categoryName, "Tower Diagonal Spacing Ex"), Equals, False)
        Equals = If(Me.TowerBraceType.CheckChange(otherToCompare.TowerBraceType, changes, categoryName, "Tower Brace Type"), Equals, False)
        Equals = If(Me.TowerFaceBevel.CheckChange(otherToCompare.TowerFaceBevel, changes, categoryName, "Tower Face Bevel"), Equals, False)
        Equals = If(Me.TowerTopGirtOffset.CheckChange(otherToCompare.TowerTopGirtOffset, changes, categoryName, "Tower Top Girt Offset"), Equals, False)
        Equals = If(Me.TowerBotGirtOffset.CheckChange(otherToCompare.TowerBotGirtOffset, changes, categoryName, "Tower Bot Girt Offset"), Equals, False)
        Equals = If(Me.TowerHasKBraceEndPanels.CheckChange(otherToCompare.TowerHasKBraceEndPanels, changes, categoryName, "Tower Has KBrace End Panels"), Equals, False)
        Equals = If(Me.TowerHasHorizontals.CheckChange(otherToCompare.TowerHasHorizontals, changes, categoryName, "Tower Has Horizontals"), Equals, False)
        Equals = If(Me.TowerLegType.CheckChange(otherToCompare.TowerLegType, changes, categoryName, "Tower Leg Type"), Equals, False)
        Equals = If(Me.TowerLegSize.CheckChange(otherToCompare.TowerLegSize, changes, categoryName, "Tower Leg Size"), Equals, False)
        Equals = If(Me.TowerLegGrade.CheckChange(otherToCompare.TowerLegGrade, changes, categoryName, "Tower Leg Grade"), Equals, False)
        Equals = If(Me.TowerLegMatlGrade.CheckChange(otherToCompare.TowerLegMatlGrade, changes, categoryName, "Tower Leg Matl Grade"), Equals, False)
        Equals = If(Me.TowerDiagonalGrade.CheckChange(otherToCompare.TowerDiagonalGrade, changes, categoryName, "Tower Diagonal Grade"), Equals, False)
        Equals = If(Me.TowerDiagonalMatlGrade.CheckChange(otherToCompare.TowerDiagonalMatlGrade, changes, categoryName, "Tower Diagonal Matl Grade"), Equals, False)
        Equals = If(Me.TowerInnerBracingGrade.CheckChange(otherToCompare.TowerInnerBracingGrade, changes, categoryName, "Tower Inner Bracing Grade"), Equals, False)
        Equals = If(Me.TowerInnerBracingMatlGrade.CheckChange(otherToCompare.TowerInnerBracingMatlGrade, changes, categoryName, "Tower Inner Bracing Matl Grade"), Equals, False)
        Equals = If(Me.TowerTopGirtGrade.CheckChange(otherToCompare.TowerTopGirtGrade, changes, categoryName, "Tower Top Girt Grade"), Equals, False)
        Equals = If(Me.TowerTopGirtMatlGrade.CheckChange(otherToCompare.TowerTopGirtMatlGrade, changes, categoryName, "Tower Top Girt Matl Grade"), Equals, False)
        Equals = If(Me.TowerBotGirtGrade.CheckChange(otherToCompare.TowerBotGirtGrade, changes, categoryName, "Tower Bot Girt Grade"), Equals, False)
        Equals = If(Me.TowerBotGirtMatlGrade.CheckChange(otherToCompare.TowerBotGirtMatlGrade, changes, categoryName, "Tower Bot Girt Matl Grade"), Equals, False)
        Equals = If(Me.TowerInnerGirtGrade.CheckChange(otherToCompare.TowerInnerGirtGrade, changes, categoryName, "Tower Inner Girt Grade"), Equals, False)
        Equals = If(Me.TowerInnerGirtMatlGrade.CheckChange(otherToCompare.TowerInnerGirtMatlGrade, changes, categoryName, "Tower Inner Girt Matl Grade"), Equals, False)
        Equals = If(Me.TowerLongHorizontalGrade.CheckChange(otherToCompare.TowerLongHorizontalGrade, changes, categoryName, "Tower Long Horizontal Grade"), Equals, False)
        Equals = If(Me.TowerLongHorizontalMatlGrade.CheckChange(otherToCompare.TowerLongHorizontalMatlGrade, changes, categoryName, "Tower Long Horizontal Matl Grade"), Equals, False)
        Equals = If(Me.TowerShortHorizontalGrade.CheckChange(otherToCompare.TowerShortHorizontalGrade, changes, categoryName, "Tower Short Horizontal Grade"), Equals, False)
        Equals = If(Me.TowerShortHorizontalMatlGrade.CheckChange(otherToCompare.TowerShortHorizontalMatlGrade, changes, categoryName, "Tower Short Horizontal Matl Grade"), Equals, False)
        Equals = If(Me.TowerDiagonalType.CheckChange(otherToCompare.TowerDiagonalType, changes, categoryName, "Tower Diagonal Type"), Equals, False)
        Equals = If(Me.TowerDiagonalSize.CheckChange(otherToCompare.TowerDiagonalSize, changes, categoryName, "Tower Diagonal Size"), Equals, False)
        Equals = If(Me.TowerInnerBracingType.CheckChange(otherToCompare.TowerInnerBracingType, changes, categoryName, "Tower Inner Bracing Type"), Equals, False)
        Equals = If(Me.TowerInnerBracingSize.CheckChange(otherToCompare.TowerInnerBracingSize, changes, categoryName, "Tower Inner Bracing Size"), Equals, False)
        Equals = If(Me.TowerTopGirtType.CheckChange(otherToCompare.TowerTopGirtType, changes, categoryName, "Tower Top Girt Type"), Equals, False)
        Equals = If(Me.TowerTopGirtSize.CheckChange(otherToCompare.TowerTopGirtSize, changes, categoryName, "Tower Top Girt Size"), Equals, False)
        Equals = If(Me.TowerBotGirtType.CheckChange(otherToCompare.TowerBotGirtType, changes, categoryName, "Tower Bot Girt Type"), Equals, False)
        Equals = If(Me.TowerBotGirtSize.CheckChange(otherToCompare.TowerBotGirtSize, changes, categoryName, "Tower Bot Girt Size"), Equals, False)
        Equals = If(Me.TowerNumInnerGirts.CheckChange(otherToCompare.TowerNumInnerGirts, changes, categoryName, "Tower Num Inner Girts"), Equals, False)
        Equals = If(Me.TowerInnerGirtType.CheckChange(otherToCompare.TowerInnerGirtType, changes, categoryName, "Tower Inner Girt Type"), Equals, False)
        Equals = If(Me.TowerInnerGirtSize.CheckChange(otherToCompare.TowerInnerGirtSize, changes, categoryName, "Tower Inner Girt Size"), Equals, False)
        Equals = If(Me.TowerLongHorizontalType.CheckChange(otherToCompare.TowerLongHorizontalType, changes, categoryName, "Tower Long Horizontal Type"), Equals, False)
        Equals = If(Me.TowerLongHorizontalSize.CheckChange(otherToCompare.TowerLongHorizontalSize, changes, categoryName, "Tower Long Horizontal Size"), Equals, False)
        Equals = If(Me.TowerShortHorizontalType.CheckChange(otherToCompare.TowerShortHorizontalType, changes, categoryName, "Tower Short Horizontal Type"), Equals, False)
        Equals = If(Me.TowerShortHorizontalSize.CheckChange(otherToCompare.TowerShortHorizontalSize, changes, categoryName, "Tower Short Horizontal Size"), Equals, False)
        Equals = If(Me.TowerRedundantGrade.CheckChange(otherToCompare.TowerRedundantGrade, changes, categoryName, "Tower Redundant Grade"), Equals, False)
        Equals = If(Me.TowerRedundantMatlGrade.CheckChange(otherToCompare.TowerRedundantMatlGrade, changes, categoryName, "Tower Redundant Matl Grade"), Equals, False)
        Equals = If(Me.TowerRedundantType.CheckChange(otherToCompare.TowerRedundantType, changes, categoryName, "Tower Redundant Type"), Equals, False)
        Equals = If(Me.TowerRedundantDiagType.CheckChange(otherToCompare.TowerRedundantDiagType, changes, categoryName, "Tower Redundant Diag Type"), Equals, False)
        Equals = If(Me.TowerRedundantSubDiagonalType.CheckChange(otherToCompare.TowerRedundantSubDiagonalType, changes, categoryName, "Tower Redundant Sub Diagonal Type"), Equals, False)
        Equals = If(Me.TowerRedundantSubHorizontalType.CheckChange(otherToCompare.TowerRedundantSubHorizontalType, changes, categoryName, "Tower Redundant Sub Horizontal Type"), Equals, False)
        Equals = If(Me.TowerRedundantVerticalType.CheckChange(otherToCompare.TowerRedundantVerticalType, changes, categoryName, "Tower Redundant Vertical Type"), Equals, False)
        Equals = If(Me.TowerRedundantHipType.CheckChange(otherToCompare.TowerRedundantHipType, changes, categoryName, "Tower Redundant Hip Type"), Equals, False)
        Equals = If(Me.TowerRedundantHipDiagonalType.CheckChange(otherToCompare.TowerRedundantHipDiagonalType, changes, categoryName, "Tower Redundant Hip Diagonal Type"), Equals, False)
        Equals = If(Me.TowerRedundantHorizontalSize.CheckChange(otherToCompare.TowerRedundantHorizontalSize, changes, categoryName, "Tower Redundant Horizontal Size"), Equals, False)
        Equals = If(Me.TowerRedundantHorizontalSize2.CheckChange(otherToCompare.TowerRedundantHorizontalSize2, changes, categoryName, "Tower Redundant Horizontal Size 2"), Equals, False)
        Equals = If(Me.TowerRedundantHorizontalSize3.CheckChange(otherToCompare.TowerRedundantHorizontalSize3, changes, categoryName, "Tower Redundant Horizontal Size 3"), Equals, False)
        Equals = If(Me.TowerRedundantHorizontalSize4.CheckChange(otherToCompare.TowerRedundantHorizontalSize4, changes, categoryName, "Tower Redundant Horizontal Size 4"), Equals, False)
        Equals = If(Me.TowerRedundantDiagonalSize.CheckChange(otherToCompare.TowerRedundantDiagonalSize, changes, categoryName, "Tower Redundant Diagonal Size"), Equals, False)
        Equals = If(Me.TowerRedundantDiagonalSize2.CheckChange(otherToCompare.TowerRedundantDiagonalSize2, changes, categoryName, "Tower Redundant Diagonal Size 2"), Equals, False)
        Equals = If(Me.TowerRedundantDiagonalSize3.CheckChange(otherToCompare.TowerRedundantDiagonalSize3, changes, categoryName, "Tower Redundant Diagonal Size 3"), Equals, False)
        Equals = If(Me.TowerRedundantDiagonalSize4.CheckChange(otherToCompare.TowerRedundantDiagonalSize4, changes, categoryName, "Tower Redundant Diagonal Size 4"), Equals, False)
        Equals = If(Me.TowerRedundantSubHorizontalSize.CheckChange(otherToCompare.TowerRedundantSubHorizontalSize, changes, categoryName, "Tower Redundant Sub Horizontal Size"), Equals, False)
        Equals = If(Me.TowerRedundantSubDiagonalSize.CheckChange(otherToCompare.TowerRedundantSubDiagonalSize, changes, categoryName, "Tower Redundant Sub Diagonal Size"), Equals, False)
        Equals = If(Me.TowerSubDiagLocation.CheckChange(otherToCompare.TowerSubDiagLocation, changes, categoryName, "Tower Sub Diag Location"), Equals, False)
        Equals = If(Me.TowerRedundantVerticalSize.CheckChange(otherToCompare.TowerRedundantVerticalSize, changes, categoryName, "Tower Redundant Vertical Size"), Equals, False)
        Equals = If(Me.TowerRedundantHipSize.CheckChange(otherToCompare.TowerRedundantHipSize, changes, categoryName, "Tower Redundant Hip Size"), Equals, False)
        Equals = If(Me.TowerRedundantHipSize2.CheckChange(otherToCompare.TowerRedundantHipSize2, changes, categoryName, "Tower Redundant Hip Size 2"), Equals, False)
        Equals = If(Me.TowerRedundantHipSize3.CheckChange(otherToCompare.TowerRedundantHipSize3, changes, categoryName, "Tower Redundant Hip Size 3"), Equals, False)
        Equals = If(Me.TowerRedundantHipSize4.CheckChange(otherToCompare.TowerRedundantHipSize4, changes, categoryName, "Tower Redundant Hip Size 4"), Equals, False)
        Equals = If(Me.TowerRedundantHipDiagonalSize.CheckChange(otherToCompare.TowerRedundantHipDiagonalSize, changes, categoryName, "Tower Redundant Hip Diagonal Size"), Equals, False)
        Equals = If(Me.TowerRedundantHipDiagonalSize2.CheckChange(otherToCompare.TowerRedundantHipDiagonalSize2, changes, categoryName, "Tower Redundant Hip Diagonal Size 2"), Equals, False)
        Equals = If(Me.TowerRedundantHipDiagonalSize3.CheckChange(otherToCompare.TowerRedundantHipDiagonalSize3, changes, categoryName, "Tower Redundant Hip Diagonal Size 3"), Equals, False)
        Equals = If(Me.TowerRedundantHipDiagonalSize4.CheckChange(otherToCompare.TowerRedundantHipDiagonalSize4, changes, categoryName, "Tower Redundant Hip Diagonal Size 4"), Equals, False)
        Equals = If(Me.TowerSWMult.CheckChange(otherToCompare.TowerSWMult, changes, categoryName, "Tower SW Mult"), Equals, False)
        Equals = If(Me.TowerWPMult.CheckChange(otherToCompare.TowerWPMult, changes, categoryName, "Tower WP Mult"), Equals, False)
        Equals = If(Me.TowerAutoCalcKSingleAngle.CheckChange(otherToCompare.TowerAutoCalcKSingleAngle, changes, categoryName, "Tower Auto Calc K Single Angle"), Equals, False)
        Equals = If(Me.TowerAutoCalcKSolidRound.CheckChange(otherToCompare.TowerAutoCalcKSolidRound, changes, categoryName, "Tower Auto Calc KSolid Round"), Equals, False)
        Equals = If(Me.TowerAfGusset.CheckChange(otherToCompare.TowerAfGusset, changes, categoryName, "Tower Af Gusset"), Equals, False)
        Equals = If(Me.TowerTfGusset.CheckChange(otherToCompare.TowerTfGusset, changes, categoryName, "Tower Tf Gusset"), Equals, False)
        Equals = If(Me.TowerGussetBoltEdgeDistance.CheckChange(otherToCompare.TowerGussetBoltEdgeDistance, changes, categoryName, "Tower Gusset Bolt Edge Distance"), Equals, False)
        Equals = If(Me.TowerGussetGrade.CheckChange(otherToCompare.TowerGussetGrade, changes, categoryName, "Tower Gusset Grade"), Equals, False)
        Equals = If(Me.TowerGussetMatlGrade.CheckChange(otherToCompare.TowerGussetMatlGrade, changes, categoryName, "Tower Gusset Matl Grade"), Equals, False)
        Equals = If(Me.TowerAfMult.CheckChange(otherToCompare.TowerAfMult, changes, categoryName, "Tower Af Mult"), Equals, False)
        Equals = If(Me.TowerArMult.CheckChange(otherToCompare.TowerArMult, changes, categoryName, "Tower Ar Mult"), Equals, False)
        Equals = If(Me.TowerFlatIPAPole.CheckChange(otherToCompare.TowerFlatIPAPole, changes, categoryName, "Tower Flat IPA Pole"), Equals, False)
        Equals = If(Me.TowerRoundIPAPole.CheckChange(otherToCompare.TowerRoundIPAPole, changes, categoryName, "Tower Round IPA Pole"), Equals, False)
        Equals = If(Me.TowerFlatIPALeg.CheckChange(otherToCompare.TowerFlatIPALeg, changes, categoryName, "Tower Flat IPA Leg"), Equals, False)
        Equals = If(Me.TowerRoundIPALeg.CheckChange(otherToCompare.TowerRoundIPALeg, changes, categoryName, "Tower Round IPA Leg"), Equals, False)
        Equals = If(Me.TowerFlatIPAHorizontal.CheckChange(otherToCompare.TowerFlatIPAHorizontal, changes, categoryName, "Tower Flat IPA Horizontal"), Equals, False)
        Equals = If(Me.TowerRoundIPAHorizontal.CheckChange(otherToCompare.TowerRoundIPAHorizontal, changes, categoryName, "Tower Round IPA Horizontal"), Equals, False)
        Equals = If(Me.TowerFlatIPADiagonal.CheckChange(otherToCompare.TowerFlatIPADiagonal, changes, categoryName, "Tower Flat IPA Diagonal"), Equals, False)
        Equals = If(Me.TowerRoundIPADiagonal.CheckChange(otherToCompare.TowerRoundIPADiagonal, changes, categoryName, "Tower Round IPA Diagonal"), Equals, False)
        Equals = If(Me.TowerCSA_S37_SpeedUpFactor.CheckChange(otherToCompare.TowerCSA_S37_SpeedUpFactor, changes, categoryName, "Tower CSA  S 37  Speed Up Factor"), Equals, False)
        Equals = If(Me.TowerKLegs.CheckChange(otherToCompare.TowerKLegs, changes, categoryName, "Tower K Legs"), Equals, False)
        Equals = If(Me.TowerKXBracedDiags.CheckChange(otherToCompare.TowerKXBracedDiags, changes, categoryName, "Tower K X-Braced Diags"), Equals, False)
        Equals = If(Me.TowerKKBracedDiags.CheckChange(otherToCompare.TowerKKBracedDiags, changes, categoryName, "Tower K K-Braced Diags"), Equals, False)
        Equals = If(Me.TowerKZBracedDiags.CheckChange(otherToCompare.TowerKZBracedDiags, changes, categoryName, "Tower K Z-Braced Diags"), Equals, False)
        Equals = If(Me.TowerKHorzs.CheckChange(otherToCompare.TowerKHorzs, changes, categoryName, "Tower K Horzs"), Equals, False)
        Equals = If(Me.TowerKSecHorzs.CheckChange(otherToCompare.TowerKSecHorzs, changes, categoryName, "Tower K Sec Horzs"), Equals, False)
        Equals = If(Me.TowerKGirts.CheckChange(otherToCompare.TowerKGirts, changes, categoryName, "Tower K Girts"), Equals, False)
        Equals = If(Me.TowerKInners.CheckChange(otherToCompare.TowerKInners, changes, categoryName, "Tower K Inners"), Equals, False)
        Equals = If(Me.TowerKXBracedDiagsY.CheckChange(otherToCompare.TowerKXBracedDiagsY, changes, categoryName, "Tower K X-Braced Diags Y"), Equals, False)
        Equals = If(Me.TowerKKBracedDiagsY.CheckChange(otherToCompare.TowerKKBracedDiagsY, changes, categoryName, "Tower K K-Braced Diags Y"), Equals, False)
        Equals = If(Me.TowerKZBracedDiagsY.CheckChange(otherToCompare.TowerKZBracedDiagsY, changes, categoryName, "Tower K Z-Braced Diags Y"), Equals, False)
        Equals = If(Me.TowerKHorzsY.CheckChange(otherToCompare.TowerKHorzsY, changes, categoryName, "Tower K Horzs Y"), Equals, False)
        Equals = If(Me.TowerKSecHorzsY.CheckChange(otherToCompare.TowerKSecHorzsY, changes, categoryName, "Tower K Sec Horzs Y"), Equals, False)
        Equals = If(Me.TowerKGirtsY.CheckChange(otherToCompare.TowerKGirtsY, changes, categoryName, "Tower K Girts Y"), Equals, False)
        Equals = If(Me.TowerKInnersY.CheckChange(otherToCompare.TowerKInnersY, changes, categoryName, "Tower K Inners Y"), Equals, False)
        Equals = If(Me.TowerKRedHorz.CheckChange(otherToCompare.TowerKRedHorz, changes, categoryName, "Tower K Red Horz"), Equals, False)
        Equals = If(Me.TowerKRedDiag.CheckChange(otherToCompare.TowerKRedDiag, changes, categoryName, "Tower K Red Diag"), Equals, False)
        Equals = If(Me.TowerKRedSubDiag.CheckChange(otherToCompare.TowerKRedSubDiag, changes, categoryName, "Tower K Red Sub Diag"), Equals, False)
        Equals = If(Me.TowerKRedSubHorz.CheckChange(otherToCompare.TowerKRedSubHorz, changes, categoryName, "Tower K Red Sub Horz"), Equals, False)
        Equals = If(Me.TowerKRedVert.CheckChange(otherToCompare.TowerKRedVert, changes, categoryName, "Tower K Red Vert"), Equals, False)
        Equals = If(Me.TowerKRedHip.CheckChange(otherToCompare.TowerKRedHip, changes, categoryName, "Tower K Red Hip"), Equals, False)
        Equals = If(Me.TowerKRedHipDiag.CheckChange(otherToCompare.TowerKRedHipDiag, changes, categoryName, "Tower K Red Hip Diag"), Equals, False)
        Equals = If(Me.TowerKTLX.CheckChange(otherToCompare.TowerKTLX, changes, categoryName, "Tower KTLX"), Equals, False)
        Equals = If(Me.TowerKTLZ.CheckChange(otherToCompare.TowerKTLZ, changes, categoryName, "Tower KTLZ"), Equals, False)
        Equals = If(Me.TowerKTLLeg.CheckChange(otherToCompare.TowerKTLLeg, changes, categoryName, "Tower KTL Leg"), Equals, False)
        Equals = If(Me.TowerInnerKTLX.CheckChange(otherToCompare.TowerInnerKTLX, changes, categoryName, "Tower Inner KTLX"), Equals, False)
        Equals = If(Me.TowerInnerKTLZ.CheckChange(otherToCompare.TowerInnerKTLZ, changes, categoryName, "Tower Inner KTLZ"), Equals, False)
        Equals = If(Me.TowerInnerKTLLeg.CheckChange(otherToCompare.TowerInnerKTLLeg, changes, categoryName, "Tower Inner KTL Leg"), Equals, False)
        Equals = If(Me.TowerStitchBoltLocationHoriz.CheckChange(otherToCompare.TowerStitchBoltLocationHoriz, changes, categoryName, "Tower Stitch Bolt Location Horiz"), Equals, False)
        Equals = If(Me.TowerStitchBoltLocationDiag.CheckChange(otherToCompare.TowerStitchBoltLocationDiag, changes, categoryName, "Tower Stitch Bolt Location Diag"), Equals, False)
        Equals = If(Me.TowerStitchBoltLocationRed.CheckChange(otherToCompare.TowerStitchBoltLocationRed, changes, categoryName, "Tower Stitch Bolt Location Red"), Equals, False)
        Equals = If(Me.TowerStitchSpacing.CheckChange(otherToCompare.TowerStitchSpacing, changes, categoryName, "Tower Stitch Spacing"), Equals, False)
        Equals = If(Me.TowerStitchSpacingDiag.CheckChange(otherToCompare.TowerStitchSpacingDiag, changes, categoryName, "Tower Stitch Spacing Diag"), Equals, False)
        Equals = If(Me.TowerStitchSpacingHorz.CheckChange(otherToCompare.TowerStitchSpacingHorz, changes, categoryName, "Tower Stitch Spacing Horz"), Equals, False)
        Equals = If(Me.TowerStitchSpacingRed.CheckChange(otherToCompare.TowerStitchSpacingRed, changes, categoryName, "Tower Stitch Spacing Red"), Equals, False)
        Equals = If(Me.TowerLegNetWidthDeduct.CheckChange(otherToCompare.TowerLegNetWidthDeduct, changes, categoryName, "Tower Leg Net Width Deduct"), Equals, False)
        Equals = If(Me.TowerLegUFactor.CheckChange(otherToCompare.TowerLegUFactor, changes, categoryName, "Tower Leg U Factor"), Equals, False)
        Equals = If(Me.TowerDiagonalNetWidthDeduct.CheckChange(otherToCompare.TowerDiagonalNetWidthDeduct, changes, categoryName, "Tower Diagonal Net Width Deduct"), Equals, False)
        Equals = If(Me.TowerTopGirtNetWidthDeduct.CheckChange(otherToCompare.TowerTopGirtNetWidthDeduct, changes, categoryName, "Tower Top Girt Net Width Deduct"), Equals, False)
        Equals = If(Me.TowerBotGirtNetWidthDeduct.CheckChange(otherToCompare.TowerBotGirtNetWidthDeduct, changes, categoryName, "Tower Bot Girt Net Width Deduct"), Equals, False)
        Equals = If(Me.TowerInnerGirtNetWidthDeduct.CheckChange(otherToCompare.TowerInnerGirtNetWidthDeduct, changes, categoryName, "Tower Inner Girt Net Width Deduct"), Equals, False)
        Equals = If(Me.TowerHorizontalNetWidthDeduct.CheckChange(otherToCompare.TowerHorizontalNetWidthDeduct, changes, categoryName, "Tower Horizontal Net Width Deduct"), Equals, False)
        Equals = If(Me.TowerShortHorizontalNetWidthDeduct.CheckChange(otherToCompare.TowerShortHorizontalNetWidthDeduct, changes, categoryName, "Tower Short Horizontal Net Width Deduct"), Equals, False)
        Equals = If(Me.TowerDiagonalUFactor.CheckChange(otherToCompare.TowerDiagonalUFactor, changes, categoryName, "Tower Diagonal U Factor"), Equals, False)
        Equals = If(Me.TowerTopGirtUFactor.CheckChange(otherToCompare.TowerTopGirtUFactor, changes, categoryName, "Tower Top Girt U Factor"), Equals, False)
        Equals = If(Me.TowerBotGirtUFactor.CheckChange(otherToCompare.TowerBotGirtUFactor, changes, categoryName, "Tower Bot Girt U Factor"), Equals, False)
        Equals = If(Me.TowerInnerGirtUFactor.CheckChange(otherToCompare.TowerInnerGirtUFactor, changes, categoryName, "Tower Inner Girt U Factor"), Equals, False)
        Equals = If(Me.TowerHorizontalUFactor.CheckChange(otherToCompare.TowerHorizontalUFactor, changes, categoryName, "Tower Horizontal U Factor"), Equals, False)
        Equals = If(Me.TowerShortHorizontalUFactor.CheckChange(otherToCompare.TowerShortHorizontalUFactor, changes, categoryName, "Tower Short Horizontal U Factor"), Equals, False)
        Equals = If(Me.TowerLegConnType.CheckChange(otherToCompare.TowerLegConnType, changes, categoryName, "Tower Leg Conn Type"), Equals, False)
        Equals = If(Me.TowerLegNumBolts.CheckChange(otherToCompare.TowerLegNumBolts, changes, categoryName, "Tower Leg Num Bolts"), Equals, False)
        Equals = If(Me.TowerDiagonalNumBolts.CheckChange(otherToCompare.TowerDiagonalNumBolts, changes, categoryName, "Tower Diagonal Num Bolts"), Equals, False)
        Equals = If(Me.TowerTopGirtNumBolts.CheckChange(otherToCompare.TowerTopGirtNumBolts, changes, categoryName, "Tower Top Girt Num Bolts"), Equals, False)
        Equals = If(Me.TowerBotGirtNumBolts.CheckChange(otherToCompare.TowerBotGirtNumBolts, changes, categoryName, "Tower Bot Girt Num Bolts"), Equals, False)
        Equals = If(Me.TowerInnerGirtNumBolts.CheckChange(otherToCompare.TowerInnerGirtNumBolts, changes, categoryName, "Tower Inner Girt Num Bolts"), Equals, False)
        Equals = If(Me.TowerHorizontalNumBolts.CheckChange(otherToCompare.TowerHorizontalNumBolts, changes, categoryName, "Tower Horizontal Num Bolts"), Equals, False)
        Equals = If(Me.TowerShortHorizontalNumBolts.CheckChange(otherToCompare.TowerShortHorizontalNumBolts, changes, categoryName, "Tower Short Horizontal Num Bolts"), Equals, False)
        Equals = If(Me.TowerLegBoltGrade.CheckChange(otherToCompare.TowerLegBoltGrade, changes, categoryName, "Tower Leg Bolt Grade"), Equals, False)
        Equals = If(Me.TowerLegBoltSize.CheckChange(otherToCompare.TowerLegBoltSize, changes, categoryName, "Tower Leg Bolt Size"), Equals, False)
        Equals = If(Me.TowerDiagonalBoltGrade.CheckChange(otherToCompare.TowerDiagonalBoltGrade, changes, categoryName, "Tower Diagonal Bolt Grade"), Equals, False)
        Equals = If(Me.TowerDiagonalBoltSize.CheckChange(otherToCompare.TowerDiagonalBoltSize, changes, categoryName, "Tower Diagonal Bolt Size"), Equals, False)
        Equals = If(Me.TowerTopGirtBoltGrade.CheckChange(otherToCompare.TowerTopGirtBoltGrade, changes, categoryName, "Tower Top Girt Bolt Grade"), Equals, False)
        Equals = If(Me.TowerTopGirtBoltSize.CheckChange(otherToCompare.TowerTopGirtBoltSize, changes, categoryName, "Tower Top Girt Bolt Size"), Equals, False)
        Equals = If(Me.TowerBotGirtBoltGrade.CheckChange(otherToCompare.TowerBotGirtBoltGrade, changes, categoryName, "Tower Bot Girt Bolt Grade"), Equals, False)
        Equals = If(Me.TowerBotGirtBoltSize.CheckChange(otherToCompare.TowerBotGirtBoltSize, changes, categoryName, "Tower Bot Girt Bolt Size"), Equals, False)
        Equals = If(Me.TowerInnerGirtBoltGrade.CheckChange(otherToCompare.TowerInnerGirtBoltGrade, changes, categoryName, "Tower Inner Girt Bolt Grade"), Equals, False)
        Equals = If(Me.TowerInnerGirtBoltSize.CheckChange(otherToCompare.TowerInnerGirtBoltSize, changes, categoryName, "Tower Inner Girt Bolt Size"), Equals, False)
        Equals = If(Me.TowerHorizontalBoltGrade.CheckChange(otherToCompare.TowerHorizontalBoltGrade, changes, categoryName, "Tower Horizontal Bolt Grade"), Equals, False)
        Equals = If(Me.TowerHorizontalBoltSize.CheckChange(otherToCompare.TowerHorizontalBoltSize, changes, categoryName, "Tower Horizontal Bolt Size"), Equals, False)
        Equals = If(Me.TowerShortHorizontalBoltGrade.CheckChange(otherToCompare.TowerShortHorizontalBoltGrade, changes, categoryName, "Tower Short Horizontal Bolt Grade"), Equals, False)
        Equals = If(Me.TowerShortHorizontalBoltSize.CheckChange(otherToCompare.TowerShortHorizontalBoltSize, changes, categoryName, "Tower Short Horizontal Bolt Size"), Equals, False)
        Equals = If(Me.TowerLegBoltEdgeDistance.CheckChange(otherToCompare.TowerLegBoltEdgeDistance, changes, categoryName, "Tower Leg Bolt Edge Distance"), Equals, False)
        Equals = If(Me.TowerDiagonalBoltEdgeDistance.CheckChange(otherToCompare.TowerDiagonalBoltEdgeDistance, changes, categoryName, "Tower Diagonal Bolt Edge Distance"), Equals, False)
        Equals = If(Me.TowerTopGirtBoltEdgeDistance.CheckChange(otherToCompare.TowerTopGirtBoltEdgeDistance, changes, categoryName, "Tower Top Girt Bolt Edge Distance"), Equals, False)
        Equals = If(Me.TowerBotGirtBoltEdgeDistance.CheckChange(otherToCompare.TowerBotGirtBoltEdgeDistance, changes, categoryName, "Tower Bot Girt Bolt Edge Distance"), Equals, False)
        Equals = If(Me.TowerInnerGirtBoltEdgeDistance.CheckChange(otherToCompare.TowerInnerGirtBoltEdgeDistance, changes, categoryName, "Tower Inner Girt Bolt Edge Distance"), Equals, False)
        Equals = If(Me.TowerHorizontalBoltEdgeDistance.CheckChange(otherToCompare.TowerHorizontalBoltEdgeDistance, changes, categoryName, "Tower Horizontal Bolt Edge Distance"), Equals, False)
        Equals = If(Me.TowerShortHorizontalBoltEdgeDistance.CheckChange(otherToCompare.TowerShortHorizontalBoltEdgeDistance, changes, categoryName, "Tower Short Horizontal Bolt Edge Distance"), Equals, False)
        Equals = If(Me.TowerDiagonalGageG1Distance.CheckChange(otherToCompare.TowerDiagonalGageG1Distance, changes, categoryName, "Tower Diagonal Gage G1 Distance"), Equals, False)
        Equals = If(Me.TowerTopGirtGageG1Distance.CheckChange(otherToCompare.TowerTopGirtGageG1Distance, changes, categoryName, "Tower Top Girt Gage G1 Distance"), Equals, False)
        Equals = If(Me.TowerBotGirtGageG1Distance.CheckChange(otherToCompare.TowerBotGirtGageG1Distance, changes, categoryName, "Tower Bot Girt Gage G1 Distance"), Equals, False)
        Equals = If(Me.TowerInnerGirtGageG1Distance.CheckChange(otherToCompare.TowerInnerGirtGageG1Distance, changes, categoryName, "Tower Inner Girt Gage G1 Distance"), Equals, False)
        Equals = If(Me.TowerHorizontalGageG1Distance.CheckChange(otherToCompare.TowerHorizontalGageG1Distance, changes, categoryName, "Tower Horizontal Gage G1 Distance"), Equals, False)
        Equals = If(Me.TowerShortHorizontalGageG1Distance.CheckChange(otherToCompare.TowerShortHorizontalGageG1Distance, changes, categoryName, "Tower Short Horizontal Gage G1 Distance"), Equals, False)
        Equals = If(Me.TowerRedundantHorizontalBoltGrade.CheckChange(otherToCompare.TowerRedundantHorizontalBoltGrade, changes, categoryName, "Tower Redundant Horizontal Bolt Grade"), Equals, False)
        Equals = If(Me.TowerRedundantHorizontalBoltSize.CheckChange(otherToCompare.TowerRedundantHorizontalBoltSize, changes, categoryName, "Tower Redundant Horizontal Bolt Size"), Equals, False)
        Equals = If(Me.TowerRedundantHorizontalNumBolts.CheckChange(otherToCompare.TowerRedundantHorizontalNumBolts, changes, categoryName, "Tower Redundant Horizontal Num Bolts"), Equals, False)
        Equals = If(Me.TowerRedundantHorizontalBoltEdgeDistance.CheckChange(otherToCompare.TowerRedundantHorizontalBoltEdgeDistance, changes, categoryName, "Tower Redundant Horizontal Bolt Edge Distance"), Equals, False)
        Equals = If(Me.TowerRedundantHorizontalGageG1Distance.CheckChange(otherToCompare.TowerRedundantHorizontalGageG1Distance, changes, categoryName, "Tower Redundant Horizontal Gage G1 Distance"), Equals, False)
        Equals = If(Me.TowerRedundantHorizontalNetWidthDeduct.CheckChange(otherToCompare.TowerRedundantHorizontalNetWidthDeduct, changes, categoryName, "Tower Redundant Horizontal Net Width Deduct"), Equals, False)
        Equals = If(Me.TowerRedundantHorizontalUFactor.CheckChange(otherToCompare.TowerRedundantHorizontalUFactor, changes, categoryName, "Tower Redundant Horizontal U Factor"), Equals, False)
        Equals = If(Me.TowerRedundantDiagonalBoltGrade.CheckChange(otherToCompare.TowerRedundantDiagonalBoltGrade, changes, categoryName, "Tower Redundant Diagonal Bolt Grade"), Equals, False)
        Equals = If(Me.TowerRedundantDiagonalBoltSize.CheckChange(otherToCompare.TowerRedundantDiagonalBoltSize, changes, categoryName, "Tower Redundant Diagonal Bolt Size"), Equals, False)
        Equals = If(Me.TowerRedundantDiagonalNumBolts.CheckChange(otherToCompare.TowerRedundantDiagonalNumBolts, changes, categoryName, "Tower Redundant Diagonal Num Bolts"), Equals, False)
        Equals = If(Me.TowerRedundantDiagonalBoltEdgeDistance.CheckChange(otherToCompare.TowerRedundantDiagonalBoltEdgeDistance, changes, categoryName, "Tower Redundant Diagonal Bolt Edge Distance"), Equals, False)
        Equals = If(Me.TowerRedundantDiagonalGageG1Distance.CheckChange(otherToCompare.TowerRedundantDiagonalGageG1Distance, changes, categoryName, "Tower Redundant Diagonal Gage G1 Distance"), Equals, False)
        Equals = If(Me.TowerRedundantDiagonalNetWidthDeduct.CheckChange(otherToCompare.TowerRedundantDiagonalNetWidthDeduct, changes, categoryName, "Tower Redundant Diagonal Net Width Deduct"), Equals, False)
        Equals = If(Me.TowerRedundantDiagonalUFactor.CheckChange(otherToCompare.TowerRedundantDiagonalUFactor, changes, categoryName, "Tower Redundant Diagonal U Factor"), Equals, False)
        Equals = If(Me.TowerRedundantSubDiagonalBoltGrade.CheckChange(otherToCompare.TowerRedundantSubDiagonalBoltGrade, changes, categoryName, "Tower Redundant Sub Diagonal Bolt Grade"), Equals, False)
        Equals = If(Me.TowerRedundantSubDiagonalBoltSize.CheckChange(otherToCompare.TowerRedundantSubDiagonalBoltSize, changes, categoryName, "Tower Redundant Sub Diagonal Bolt Size"), Equals, False)
        Equals = If(Me.TowerRedundantSubDiagonalNumBolts.CheckChange(otherToCompare.TowerRedundantSubDiagonalNumBolts, changes, categoryName, "Tower Redundant Sub Diagonal Num Bolts"), Equals, False)
        Equals = If(Me.TowerRedundantSubDiagonalBoltEdgeDistance.CheckChange(otherToCompare.TowerRedundantSubDiagonalBoltEdgeDistance, changes, categoryName, "Tower Redundant Sub Diagonal Bolt Edge Distance"), Equals, False)
        Equals = If(Me.TowerRedundantSubDiagonalGageG1Distance.CheckChange(otherToCompare.TowerRedundantSubDiagonalGageG1Distance, changes, categoryName, "Tower Redundant Sub Diagonal Gage G1 Distance"), Equals, False)
        Equals = If(Me.TowerRedundantSubDiagonalNetWidthDeduct.CheckChange(otherToCompare.TowerRedundantSubDiagonalNetWidthDeduct, changes, categoryName, "Tower Redundant Sub Diagonal Net Width Deduct"), Equals, False)
        Equals = If(Me.TowerRedundantSubDiagonalUFactor.CheckChange(otherToCompare.TowerRedundantSubDiagonalUFactor, changes, categoryName, "Tower Redundant Sub Diagonal U Factor"), Equals, False)
        Equals = If(Me.TowerRedundantSubHorizontalBoltGrade.CheckChange(otherToCompare.TowerRedundantSubHorizontalBoltGrade, changes, categoryName, "Tower Redundant Sub Horizontal Bolt Grade"), Equals, False)
        Equals = If(Me.TowerRedundantSubHorizontalBoltSize.CheckChange(otherToCompare.TowerRedundantSubHorizontalBoltSize, changes, categoryName, "Tower Redundant Sub Horizontal Bolt Size"), Equals, False)
        Equals = If(Me.TowerRedundantSubHorizontalNumBolts.CheckChange(otherToCompare.TowerRedundantSubHorizontalNumBolts, changes, categoryName, "Tower Redundant Sub Horizontal Num Bolts"), Equals, False)
        Equals = If(Me.TowerRedundantSubHorizontalBoltEdgeDistance.CheckChange(otherToCompare.TowerRedundantSubHorizontalBoltEdgeDistance, changes, categoryName, "Tower Redundant Sub Horizontal Bolt Edge Distance"), Equals, False)
        Equals = If(Me.TowerRedundantSubHorizontalGageG1Distance.CheckChange(otherToCompare.TowerRedundantSubHorizontalGageG1Distance, changes, categoryName, "Tower Redundant Sub Horizontal Gage G1 Distance"), Equals, False)
        Equals = If(Me.TowerRedundantSubHorizontalNetWidthDeduct.CheckChange(otherToCompare.TowerRedundantSubHorizontalNetWidthDeduct, changes, categoryName, "Tower Redundant Sub Horizontal Net Width Deduct"), Equals, False)
        Equals = If(Me.TowerRedundantSubHorizontalUFactor.CheckChange(otherToCompare.TowerRedundantSubHorizontalUFactor, changes, categoryName, "Tower Redundant Sub Horizontal U Factor"), Equals, False)
        Equals = If(Me.TowerRedundantVerticalBoltGrade.CheckChange(otherToCompare.TowerRedundantVerticalBoltGrade, changes, categoryName, "Tower Redundant Vertical Bolt Grade"), Equals, False)
        Equals = If(Me.TowerRedundantVerticalBoltSize.CheckChange(otherToCompare.TowerRedundantVerticalBoltSize, changes, categoryName, "Tower Redundant Vertical Bolt Size"), Equals, False)
        Equals = If(Me.TowerRedundantVerticalNumBolts.CheckChange(otherToCompare.TowerRedundantVerticalNumBolts, changes, categoryName, "Tower Redundant Vertical Num Bolts"), Equals, False)
        Equals = If(Me.TowerRedundantVerticalBoltEdgeDistance.CheckChange(otherToCompare.TowerRedundantVerticalBoltEdgeDistance, changes, categoryName, "Tower Redundant Vertical Bolt Edge Distance"), Equals, False)
        Equals = If(Me.TowerRedundantVerticalGageG1Distance.CheckChange(otherToCompare.TowerRedundantVerticalGageG1Distance, changes, categoryName, "Tower Redundant Vertical Gage G1 Distance"), Equals, False)
        Equals = If(Me.TowerRedundantVerticalNetWidthDeduct.CheckChange(otherToCompare.TowerRedundantVerticalNetWidthDeduct, changes, categoryName, "Tower Redundant Vertical Net Width Deduct"), Equals, False)
        Equals = If(Me.TowerRedundantVerticalUFactor.CheckChange(otherToCompare.TowerRedundantVerticalUFactor, changes, categoryName, "Tower Redundant Vertical U Factor"), Equals, False)
        Equals = If(Me.TowerRedundantHipBoltGrade.CheckChange(otherToCompare.TowerRedundantHipBoltGrade, changes, categoryName, "Tower Redundant Hip Bolt Grade"), Equals, False)
        Equals = If(Me.TowerRedundantHipBoltSize.CheckChange(otherToCompare.TowerRedundantHipBoltSize, changes, categoryName, "Tower Redundant Hip Bolt Size"), Equals, False)
        Equals = If(Me.TowerRedundantHipNumBolts.CheckChange(otherToCompare.TowerRedundantHipNumBolts, changes, categoryName, "Tower Redundant Hip Num Bolts"), Equals, False)
        Equals = If(Me.TowerRedundantHipBoltEdgeDistance.CheckChange(otherToCompare.TowerRedundantHipBoltEdgeDistance, changes, categoryName, "Tower Redundant Hip Bolt Edge Distance"), Equals, False)
        Equals = If(Me.TowerRedundantHipGageG1Distance.CheckChange(otherToCompare.TowerRedundantHipGageG1Distance, changes, categoryName, "Tower Redundant Hip Gage G1 Distance"), Equals, False)
        Equals = If(Me.TowerRedundantHipNetWidthDeduct.CheckChange(otherToCompare.TowerRedundantHipNetWidthDeduct, changes, categoryName, "Tower Redundant Hip Net Width Deduct"), Equals, False)
        Equals = If(Me.TowerRedundantHipUFactor.CheckChange(otherToCompare.TowerRedundantHipUFactor, changes, categoryName, "Tower Redundant Hip UFactor"), Equals, False)
        Equals = If(Me.TowerRedundantHipDiagonalBoltGrade.CheckChange(otherToCompare.TowerRedundantHipDiagonalBoltGrade, changes, categoryName, "Tower Redundant Hip Diagonal Bolt Grade"), Equals, False)
        Equals = If(Me.TowerRedundantHipDiagonalBoltSize.CheckChange(otherToCompare.TowerRedundantHipDiagonalBoltSize, changes, categoryName, "Tower Redundant Hip Diagonal Bolt Size"), Equals, False)
        Equals = If(Me.TowerRedundantHipDiagonalNumBolts.CheckChange(otherToCompare.TowerRedundantHipDiagonalNumBolts, changes, categoryName, "Tower Redundant Hip Diagonal Num Bolts"), Equals, False)
        Equals = If(Me.TowerRedundantHipDiagonalBoltEdgeDistance.CheckChange(otherToCompare.TowerRedundantHipDiagonalBoltEdgeDistance, changes, categoryName, "Tower Redundant Hip Diagonal Bolt Edge Distance"), Equals, False)
        Equals = If(Me.TowerRedundantHipDiagonalGageG1Distance.CheckChange(otherToCompare.TowerRedundantHipDiagonalGageG1Distance, changes, categoryName, "Tower Redundant Hip Diagonal Gage G1 Distance"), Equals, False)
        Equals = If(Me.TowerRedundantHipDiagonalNetWidthDeduct.CheckChange(otherToCompare.TowerRedundantHipDiagonalNetWidthDeduct, changes, categoryName, "Tower Redundant Hip Diagonal Net Width Deduct"), Equals, False)
        Equals = If(Me.TowerRedundantHipDiagonalUFactor.CheckChange(otherToCompare.TowerRedundantHipDiagonalUFactor, changes, categoryName, "Tower Redundant Hip Diagonal UFactor"), Equals, False)
        Equals = If(Me.TowerDiagonalOutOfPlaneRestraint.CheckChange(otherToCompare.TowerDiagonalOutOfPlaneRestraint, changes, categoryName, "Tower Diagonal Out Of Plane Restraint"), Equals, False)
        Equals = If(Me.TowerTopGirtOutOfPlaneRestraint.CheckChange(otherToCompare.TowerTopGirtOutOfPlaneRestraint, changes, categoryName, "Tower Top Girt Out Of Plane Restraint"), Equals, False)
        Equals = If(Me.TowerBottomGirtOutOfPlaneRestraint.CheckChange(otherToCompare.TowerBottomGirtOutOfPlaneRestraint, changes, categoryName, "Tower Bottom Girt Out Of Plane Restraint"), Equals, False)
        Equals = If(Me.TowerMidGirtOutOfPlaneRestraint.CheckChange(otherToCompare.TowerMidGirtOutOfPlaneRestraint, changes, categoryName, "Tower Mid Girt Out Of Plane Restraint"), Equals, False)
        Equals = If(Me.TowerHorizontalOutOfPlaneRestraint.CheckChange(otherToCompare.TowerHorizontalOutOfPlaneRestraint, changes, categoryName, "Tower Horizontal Out Of Plane Restraint"), Equals, False)
        Equals = If(Me.TowerSecondaryHorizontalOutOfPlaneRestraint.CheckChange(otherToCompare.TowerSecondaryHorizontalOutOfPlaneRestraint, changes, categoryName, "Tower Secondary Horizontal Out Of Plane Restraint"), Equals, False)
        Equals = If(Me.TowerUniqueFlag.CheckChange(otherToCompare.TowerUniqueFlag, changes, categoryName, "Tower Unique Flag"), Equals, False)
        Equals = If(Me.TowerDiagOffsetNEY.CheckChange(otherToCompare.TowerDiagOffsetNEY, changes, categoryName, "Tower Diag Offset NEY"), Equals, False)
        Equals = If(Me.TowerDiagOffsetNEX.CheckChange(otherToCompare.TowerDiagOffsetNEX, changes, categoryName, "Tower Diag Offset NEX"), Equals, False)
        Equals = If(Me.TowerDiagOffsetPEY.CheckChange(otherToCompare.TowerDiagOffsetPEY, changes, categoryName, "Tower Diag Offset PEY"), Equals, False)
        Equals = If(Me.TowerDiagOffsetPEX.CheckChange(otherToCompare.TowerDiagOffsetPEX, changes, categoryName, "Tower Diag Offset PEX"), Equals, False)
        Equals = If(Me.TowerKbraceOffsetNEY.CheckChange(otherToCompare.TowerKbraceOffsetNEY, changes, categoryName, "Tower Kbrace Offset NEY"), Equals, False)
        Equals = If(Me.TowerKbraceOffsetNEX.CheckChange(otherToCompare.TowerKbraceOffsetNEX, changes, categoryName, "Tower Kbrace Offset NEX"), Equals, False)
        Equals = If(Me.TowerKbraceOffsetPEY.CheckChange(otherToCompare.TowerKbraceOffsetPEY, changes, categoryName, "Tower Kbrace Offset PEY"), Equals, False)
        Equals = If(Me.TowerKbraceOffsetPEX.CheckChange(otherToCompare.TowerKbraceOffsetPEX, changes, categoryName, "Tower Kbrace Offset PEX"), Equals, False)

        Return Equals
    End Function

#Region "Testing"

    'Public Function GenerateDataTable() As DataTable
    '    GenerateDataTable = New DataTable
    '    '    GenerateDataTable.Columns.Add("ID", GetTypeNullable(Me.ID))
    '    GenerateDataTable.Columns.Add("TowerRec", GetTypeNullable(Me.TowerRec))
    '    'GenerateDataTable.Columns.Add("TowerDatabase", GetTypeNullable(Me.TowerDatabase))
    '    '    GenerateDataTable.Columns.Add("TowerName", GetTypeNullable(Me.TowerName))
    '    '    GenerateDataTable.Columns.Add("TowerHeight", GetTypeNullable(Me.TowerHeight))
    '    '    GenerateDataTable.Columns.Add("TowerFaceWidth", GetTypeNullable(Me.TowerFaceWidth))
    '    '    GenerateDataTable.Columns.Add("TowerNumSections", GetTypeNullable(Me.TowerNumSections))
    '    '    GenerateDataTable.Columns.Add("TowerSectionLength", GetTypeNullable(Me.TowerSectionLength))
    '    '    GenerateDataTable.Columns.Add("TowerDiagonalSpacing", GetTypeNullable(Me.TowerDiagonalSpacing))
    '    '    GenerateDataTable.Columns.Add("TowerDiagonalSpacingEx", GetTypeNullable(Me.TowerDiagonalSpacingEx))
    '    GenerateDataTable.Columns.Add("TowerBraceType", GetTypeNullable(Me.TowerBraceType))
    '    '    GenerateDataTable.Columns.Add("TowerFaceBevel", GetTypeNullable(Me.TowerFaceBevel))
    '    '    GenerateDataTable.Columns.Add("TowerTopGirtOffset", GetTypeNullable(Me.TowerTopGirtOffset))
    '    '    GenerateDataTable.Columns.Add("TowerBotGirtOffset", GetTypeNullable(Me.TowerBotGirtOffset))
    '    '    GenerateDataTable.Columns.Add("TowerHasKBraceEndPanels", GetTypeNullable(Me.TowerHasKBraceEndPanels))
    '    '    GenerateDataTable.Columns.Add("TowerHasHorizontals", GetTypeNullable(Me.TowerHasHorizontals))
    '    '    GenerateDataTable.Columns.Add("TowerLegType", GetTypeNullable(Me.TowerLegType))
    '    '    GenerateDataTable.Columns.Add("TowerLegSize", GetTypeNullable(Me.TowerLegSize))
    '    '    GenerateDataTable.Columns.Add("TowerLegGrade", GetTypeNullable(Me.TowerLegGrade))
    '    '    GenerateDataTable.Columns.Add("TowerLegMatlGrade", GetTypeNullable(Me.TowerLegMatlGrade))
    '    '    GenerateDataTable.Columns.Add("TowerDiagonalGrade", GetTypeNullable(Me.TowerDiagonalGrade))
    '    '    GenerateDataTable.Columns.Add("TowerDiagonalMatlGrade", GetTypeNullable(Me.TowerDiagonalMatlGrade))
    '    '    GenerateDataTable.Columns.Add("TowerInnerBracingGrade", GetTypeNullable(Me.TowerInnerBracingGrade))
    '    '    GenerateDataTable.Columns.Add("TowerInnerBracingMatlGrade", GetTypeNullable(Me.TowerInnerBracingMatlGrade))
    '    '    GenerateDataTable.Columns.Add("TowerTopGirtGrade", GetTypeNullable(Me.TowerTopGirtGrade))
    '    '    GenerateDataTable.Columns.Add("TowerTopGirtMatlGrade", GetTypeNullable(Me.TowerTopGirtMatlGrade))
    '    '    GenerateDataTable.Columns.Add("TowerBotGirtGrade", GetTypeNullable(Me.TowerBotGirtGrade))
    '    '    GenerateDataTable.Columns.Add("TowerBotGirtMatlGrade", GetTypeNullable(Me.TowerBotGirtMatlGrade))
    '    '    GenerateDataTable.Columns.Add("TowerInnerGirtGrade", GetTypeNullable(Me.TowerInnerGirtGrade))
    '    '    GenerateDataTable.Columns.Add("TowerInnerGirtMatlGrade", GetTypeNullable(Me.TowerInnerGirtMatlGrade))
    '    '    GenerateDataTable.Columns.Add("TowerLongHorizontalGrade", GetTypeNullable(Me.TowerLongHorizontalGrade))
    '    '    GenerateDataTable.Columns.Add("TowerLongHorizontalMatlGrade", GetTypeNullable(Me.TowerLongHorizontalMatlGrade))
    '    '    GenerateDataTable.Columns.Add("TowerShortHorizontalGrade", GetTypeNullable(Me.TowerShortHorizontalGrade))
    '    '    GenerateDataTable.Columns.Add("TowerShortHorizontalMatlGrade", GetTypeNullable(Me.TowerShortHorizontalMatlGrade))
    '    '    GenerateDataTable.Columns.Add("TowerDiagonalType", GetTypeNullable(Me.TowerDiagonalType))
    '    '    GenerateDataTable.Columns.Add("TowerDiagonalSize", GetTypeNullable(Me.TowerDiagonalSize))
    '    '    GenerateDataTable.Columns.Add("TowerInnerBracingType", GetTypeNullable(Me.TowerInnerBracingType))
    '    '    GenerateDataTable.Columns.Add("TowerInnerBracingSize", GetTypeNullable(Me.TowerInnerBracingSize))
    '    '    GenerateDataTable.Columns.Add("TowerTopGirtType", GetTypeNullable(Me.TowerTopGirtType))
    '    '    GenerateDataTable.Columns.Add("TowerTopGirtSize", GetTypeNullable(Me.TowerTopGirtSize))
    '    '    GenerateDataTable.Columns.Add("TowerBotGirtType", GetTypeNullable(Me.TowerBotGirtType))
    '    '    GenerateDataTable.Columns.Add("TowerBotGirtSize", GetTypeNullable(Me.TowerBotGirtSize))
    '    '    GenerateDataTable.Columns.Add("TowerNumInnerGirts", GetTypeNullable(Me.TowerNumInnerGirts))
    '    '    GenerateDataTable.Columns.Add("TowerInnerGirtType", GetTypeNullable(Me.TowerInnerGirtType))
    '    '    GenerateDataTable.Columns.Add("TowerInnerGirtSize", GetTypeNullable(Me.TowerInnerGirtSize))
    '    '    GenerateDataTable.Columns.Add("TowerLongHorizontalType", GetTypeNullable(Me.TowerLongHorizontalType))
    '    '    GenerateDataTable.Columns.Add("TowerLongHorizontalSize", GetTypeNullable(Me.TowerLongHorizontalSize))
    '    '    GenerateDataTable.Columns.Add("TowerShortHorizontalType", GetTypeNullable(Me.TowerShortHorizontalType))
    '    '    GenerateDataTable.Columns.Add("TowerShortHorizontalSize", GetTypeNullable(Me.TowerShortHorizontalSize))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantGrade", GetTypeNullable(Me.TowerRedundantGrade))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantMatlGrade", GetTypeNullable(Me.TowerRedundantMatlGrade))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantType", GetTypeNullable(Me.TowerRedundantType))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantDiagType", GetTypeNullable(Me.TowerRedundantDiagType))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantSubDiagonalType", GetTypeNullable(Me.TowerRedundantSubDiagonalType))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantSubHorizontalType", GetTypeNullable(Me.TowerRedundantSubHorizontalType))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantVerticalType", GetTypeNullable(Me.TowerRedundantVerticalType))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHipType", GetTypeNullable(Me.TowerRedundantHipType))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHipDiagonalType", GetTypeNullable(Me.TowerRedundantHipDiagonalType))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHorizontalSize", GetTypeNullable(Me.TowerRedundantHorizontalSize))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHorizontalSize2", GetTypeNullable(Me.TowerRedundantHorizontalSize2))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHorizontalSize3", GetTypeNullable(Me.TowerRedundantHorizontalSize3))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHorizontalSize4", GetTypeNullable(Me.TowerRedundantHorizontalSize4))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantDiagonalSize", GetTypeNullable(Me.TowerRedundantDiagonalSize))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantDiagonalSize2", GetTypeNullable(Me.TowerRedundantDiagonalSize2))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantDiagonalSize3", GetTypeNullable(Me.TowerRedundantDiagonalSize3))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantDiagonalSize4", GetTypeNullable(Me.TowerRedundantDiagonalSize4))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantSubHorizontalSize", GetTypeNullable(Me.TowerRedundantSubHorizontalSize))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantSubDiagonalSize", GetTypeNullable(Me.TowerRedundantSubDiagonalSize))
    '    '    GenerateDataTable.Columns.Add("TowerSubDiagLocation", GetTypeNullable(Me.TowerSubDiagLocation))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantVerticalSize", GetTypeNullable(Me.TowerRedundantVerticalSize))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHipSize", GetTypeNullable(Me.TowerRedundantHipSize))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHipSize2", GetTypeNullable(Me.TowerRedundantHipSize2))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHipSize3", GetTypeNullable(Me.TowerRedundantHipSize3))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHipSize4", GetTypeNullable(Me.TowerRedundantHipSize4))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHipDiagonalSize", GetTypeNullable(Me.TowerRedundantHipDiagonalSize))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHipDiagonalSize2", GetTypeNullable(Me.TowerRedundantHipDiagonalSize2))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHipDiagonalSize3", GetTypeNullable(Me.TowerRedundantHipDiagonalSize3))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHipDiagonalSize4", GetTypeNullable(Me.TowerRedundantHipDiagonalSize4))
    '    '    GenerateDataTable.Columns.Add("TowerSWMult", GetTypeNullable(Me.TowerSWMult))
    '    '    GenerateDataTable.Columns.Add("TowerWPMult", GetTypeNullable(Me.TowerWPMult))
    '    '    GenerateDataTable.Columns.Add("TowerAutoCalcKSingleAngle", GetTypeNullable(Me.TowerAutoCalcKSingleAngle))
    '    '    GenerateDataTable.Columns.Add("TowerAutoCalcKSolidRound", GetTypeNullable(Me.TowerAutoCalcKSolidRound))
    '    '    GenerateDataTable.Columns.Add("TowerAfGusset", GetTypeNullable(Me.TowerAfGusset))
    '    '    GenerateDataTable.Columns.Add("TowerTfGusset", GetTypeNullable(Me.TowerTfGusset))
    '    '    GenerateDataTable.Columns.Add("TowerGussetBoltEdgeDistance", GetTypeNullable(Me.TowerGussetBoltEdgeDistance))
    '    '    GenerateDataTable.Columns.Add("TowerGussetGrade", GetTypeNullable(Me.TowerGussetGrade))
    '    '    GenerateDataTable.Columns.Add("TowerGussetMatlGrade", GetTypeNullable(Me.TowerGussetMatlGrade))
    '    '    GenerateDataTable.Columns.Add("TowerAfMult", GetTypeNullable(Me.TowerAfMult))
    '    '    GenerateDataTable.Columns.Add("TowerArMult", GetTypeNullable(Me.TowerArMult))
    '    '    GenerateDataTable.Columns.Add("TowerFlatIPAPole", GetTypeNullable(Me.TowerFlatIPAPole))
    '    '    GenerateDataTable.Columns.Add("TowerRoundIPAPole", GetTypeNullable(Me.TowerRoundIPAPole))
    '    '    GenerateDataTable.Columns.Add("TowerFlatIPALeg", GetTypeNullable(Me.TowerFlatIPALeg))
    '    '    GenerateDataTable.Columns.Add("TowerRoundIPALeg", GetTypeNullable(Me.TowerRoundIPALeg))
    '    '    GenerateDataTable.Columns.Add("TowerFlatIPAHorizontal", GetTypeNullable(Me.TowerFlatIPAHorizontal))
    '    '    GenerateDataTable.Columns.Add("TowerRoundIPAHorizontal", GetTypeNullable(Me.TowerRoundIPAHorizontal))
    '    '    GenerateDataTable.Columns.Add("TowerFlatIPADiagonal", GetTypeNullable(Me.TowerFlatIPADiagonal))
    '    '    GenerateDataTable.Columns.Add("TowerRoundIPADiagonal", GetTypeNullable(Me.TowerRoundIPADiagonal))
    '    '    GenerateDataTable.Columns.Add("TowerCSA_S37_SpeedUpFactor", GetTypeNullable(Me.TowerCSA_S37_SpeedUpFactor))
    '    '    GenerateDataTable.Columns.Add("TowerKLegs", GetTypeNullable(Me.TowerKLegs))
    '    '    GenerateDataTable.Columns.Add("TowerKXBracedDiags", GetTypeNullable(Me.TowerKXBracedDiags))
    '    '    GenerateDataTable.Columns.Add("TowerKKBracedDiags", GetTypeNullable(Me.TowerKKBracedDiags))
    '    '    GenerateDataTable.Columns.Add("TowerKZBracedDiags", GetTypeNullable(Me.TowerKZBracedDiags))
    '    '    GenerateDataTable.Columns.Add("TowerKHorzs", GetTypeNullable(Me.TowerKHorzs))
    '    '    GenerateDataTable.Columns.Add("TowerKSecHorzs", GetTypeNullable(Me.TowerKSecHorzs))
    '    '    GenerateDataTable.Columns.Add("TowerKGirts", GetTypeNullable(Me.TowerKGirts))
    '    '    GenerateDataTable.Columns.Add("TowerKInners", GetTypeNullable(Me.TowerKInners))
    '    '    GenerateDataTable.Columns.Add("TowerKXBracedDiagsY", GetTypeNullable(Me.TowerKXBracedDiagsY))
    '    '    GenerateDataTable.Columns.Add("TowerKKBracedDiagsY", GetTypeNullable(Me.TowerKKBracedDiagsY))
    '    '    GenerateDataTable.Columns.Add("TowerKZBracedDiagsY", GetTypeNullable(Me.TowerKZBracedDiagsY))
    '    '    GenerateDataTable.Columns.Add("TowerKHorzsY", GetTypeNullable(Me.TowerKHorzsY))
    '    '    GenerateDataTable.Columns.Add("TowerKSecHorzsY", GetTypeNullable(Me.TowerKSecHorzsY))
    '    '    GenerateDataTable.Columns.Add("TowerKGirtsY", GetTypeNullable(Me.TowerKGirtsY))
    '    '    GenerateDataTable.Columns.Add("TowerKInnersY", GetTypeNullable(Me.TowerKInnersY))
    '    '    GenerateDataTable.Columns.Add("TowerKRedHorz", GetTypeNullable(Me.TowerKRedHorz))
    '    '    GenerateDataTable.Columns.Add("TowerKRedDiag", GetTypeNullable(Me.TowerKRedDiag))
    '    '    GenerateDataTable.Columns.Add("TowerKRedSubDiag", GetTypeNullable(Me.TowerKRedSubDiag))
    '    '    GenerateDataTable.Columns.Add("TowerKRedSubHorz", GetTypeNullable(Me.TowerKRedSubHorz))
    '    '    GenerateDataTable.Columns.Add("TowerKRedVert", GetTypeNullable(Me.TowerKRedVert))
    '    '    GenerateDataTable.Columns.Add("TowerKRedHip", GetTypeNullable(Me.TowerKRedHip))
    '    '    GenerateDataTable.Columns.Add("TowerKRedHipDiag", GetTypeNullable(Me.TowerKRedHipDiag))
    '    '    GenerateDataTable.Columns.Add("TowerKTLX", GetTypeNullable(Me.TowerKTLX))
    '    '    GenerateDataTable.Columns.Add("TowerKTLZ", GetTypeNullable(Me.TowerKTLZ))
    '    '    GenerateDataTable.Columns.Add("TowerKTLLeg", GetTypeNullable(Me.TowerKTLLeg))
    '    '    GenerateDataTable.Columns.Add("TowerInnerKTLX", GetTypeNullable(Me.TowerInnerKTLX))
    '    '    GenerateDataTable.Columns.Add("TowerInnerKTLZ", GetTypeNullable(Me.TowerInnerKTLZ))
    '    '    GenerateDataTable.Columns.Add("TowerInnerKTLLeg", GetTypeNullable(Me.TowerInnerKTLLeg))
    '    '    GenerateDataTable.Columns.Add("TowerStitchBoltLocationHoriz", GetTypeNullable(Me.TowerStitchBoltLocationHoriz))
    '    '    GenerateDataTable.Columns.Add("TowerStitchBoltLocationDiag", GetTypeNullable(Me.TowerStitchBoltLocationDiag))
    '    '    GenerateDataTable.Columns.Add("TowerStitchBoltLocationRed", GetTypeNullable(Me.TowerStitchBoltLocationRed))
    '    '    GenerateDataTable.Columns.Add("TowerStitchSpacing", GetTypeNullable(Me.TowerStitchSpacing))
    '    '    GenerateDataTable.Columns.Add("TowerStitchSpacingDiag", GetTypeNullable(Me.TowerStitchSpacingDiag))
    '    '    GenerateDataTable.Columns.Add("TowerStitchSpacingHorz", GetTypeNullable(Me.TowerStitchSpacingHorz))
    '    '    GenerateDataTable.Columns.Add("TowerStitchSpacingRed", GetTypeNullable(Me.TowerStitchSpacingRed))
    '    '    GenerateDataTable.Columns.Add("TowerLegNetWidthDeduct", GetTypeNullable(Me.TowerLegNetWidthDeduct))
    '    '    GenerateDataTable.Columns.Add("TowerLegUFactor", GetTypeNullable(Me.TowerLegUFactor))
    '    '    GenerateDataTable.Columns.Add("TowerDiagonalNetWidthDeduct", GetTypeNullable(Me.TowerDiagonalNetWidthDeduct))
    '    '    GenerateDataTable.Columns.Add("TowerTopGirtNetWidthDeduct", GetTypeNullable(Me.TowerTopGirtNetWidthDeduct))
    '    '    GenerateDataTable.Columns.Add("TowerBotGirtNetWidthDeduct", GetTypeNullable(Me.TowerBotGirtNetWidthDeduct))
    '    '    GenerateDataTable.Columns.Add("TowerInnerGirtNetWidthDeduct", GetTypeNullable(Me.TowerInnerGirtNetWidthDeduct))
    '    '    GenerateDataTable.Columns.Add("TowerHorizontalNetWidthDeduct", GetTypeNullable(Me.TowerHorizontalNetWidthDeduct))
    '    '    GenerateDataTable.Columns.Add("TowerShortHorizontalNetWidthDeduct", GetTypeNullable(Me.TowerShortHorizontalNetWidthDeduct))
    '    '    GenerateDataTable.Columns.Add("TowerDiagonalUFactor", GetTypeNullable(Me.TowerDiagonalUFactor))
    '    '    GenerateDataTable.Columns.Add("TowerTopGirtUFactor", GetTypeNullable(Me.TowerTopGirtUFactor))
    '    '    GenerateDataTable.Columns.Add("TowerBotGirtUFactor", GetTypeNullable(Me.TowerBotGirtUFactor))
    '    '    GenerateDataTable.Columns.Add("TowerInnerGirtUFactor", GetTypeNullable(Me.TowerInnerGirtUFactor))
    '    '    GenerateDataTable.Columns.Add("TowerHorizontalUFactor", GetTypeNullable(Me.TowerHorizontalUFactor))
    '    '    GenerateDataTable.Columns.Add("TowerShortHorizontalUFactor", GetTypeNullable(Me.TowerShortHorizontalUFactor))
    '    '    GenerateDataTable.Columns.Add("TowerLegConnType", GetTypeNullable(Me.TowerLegConnType))
    '    '    GenerateDataTable.Columns.Add("TowerLegNumBolts", GetTypeNullable(Me.TowerLegNumBolts))
    '    '    GenerateDataTable.Columns.Add("TowerDiagonalNumBolts", GetTypeNullable(Me.TowerDiagonalNumBolts))
    '    '    GenerateDataTable.Columns.Add("TowerTopGirtNumBolts", GetTypeNullable(Me.TowerTopGirtNumBolts))
    '    '    GenerateDataTable.Columns.Add("TowerBotGirtNumBolts", GetTypeNullable(Me.TowerBotGirtNumBolts))
    '    '    GenerateDataTable.Columns.Add("TowerInnerGirtNumBolts", GetTypeNullable(Me.TowerInnerGirtNumBolts))
    '    '    GenerateDataTable.Columns.Add("TowerHorizontalNumBolts", GetTypeNullable(Me.TowerHorizontalNumBolts))
    '    '    GenerateDataTable.Columns.Add("TowerShortHorizontalNumBolts", GetTypeNullable(Me.TowerShortHorizontalNumBolts))
    '    '    GenerateDataTable.Columns.Add("TowerLegBoltGrade", GetTypeNullable(Me.TowerLegBoltGrade))
    '    '    GenerateDataTable.Columns.Add("TowerLegBoltSize", GetTypeNullable(Me.TowerLegBoltSize))
    '    '    GenerateDataTable.Columns.Add("TowerDiagonalBoltGrade", GetTypeNullable(Me.TowerDiagonalBoltGrade))
    '    '    GenerateDataTable.Columns.Add("TowerDiagonalBoltSize", GetTypeNullable(Me.TowerDiagonalBoltSize))
    '    '    GenerateDataTable.Columns.Add("TowerTopGirtBoltGrade", GetTypeNullable(Me.TowerTopGirtBoltGrade))
    '    '    GenerateDataTable.Columns.Add("TowerTopGirtBoltSize", GetTypeNullable(Me.TowerTopGirtBoltSize))
    '    '    GenerateDataTable.Columns.Add("TowerBotGirtBoltGrade", GetTypeNullable(Me.TowerBotGirtBoltGrade))
    '    '    GenerateDataTable.Columns.Add("TowerBotGirtBoltSize", GetTypeNullable(Me.TowerBotGirtBoltSize))
    '    '    GenerateDataTable.Columns.Add("TowerInnerGirtBoltGrade", GetTypeNullable(Me.TowerInnerGirtBoltGrade))
    '    '    GenerateDataTable.Columns.Add("TowerInnerGirtBoltSize", GetTypeNullable(Me.TowerInnerGirtBoltSize))
    '    '    GenerateDataTable.Columns.Add("TowerHorizontalBoltGrade", GetTypeNullable(Me.TowerHorizontalBoltGrade))
    '    '    GenerateDataTable.Columns.Add("TowerHorizontalBoltSize", GetTypeNullable(Me.TowerHorizontalBoltSize))
    '    '    GenerateDataTable.Columns.Add("TowerShortHorizontalBoltGrade", GetTypeNullable(Me.TowerShortHorizontalBoltGrade))
    '    '    GenerateDataTable.Columns.Add("TowerShortHorizontalBoltSize", GetTypeNullable(Me.TowerShortHorizontalBoltSize))
    '    '    GenerateDataTable.Columns.Add("TowerLegBoltEdgeDistance", GetTypeNullable(Me.TowerLegBoltEdgeDistance))
    '    '    GenerateDataTable.Columns.Add("TowerDiagonalBoltEdgeDistance", GetTypeNullable(Me.TowerDiagonalBoltEdgeDistance))
    '    '    GenerateDataTable.Columns.Add("TowerTopGirtBoltEdgeDistance", GetTypeNullable(Me.TowerTopGirtBoltEdgeDistance))
    '    '    GenerateDataTable.Columns.Add("TowerBotGirtBoltEdgeDistance", GetTypeNullable(Me.TowerBotGirtBoltEdgeDistance))
    '    '    GenerateDataTable.Columns.Add("TowerInnerGirtBoltEdgeDistance", GetTypeNullable(Me.TowerInnerGirtBoltEdgeDistance))
    '    '    GenerateDataTable.Columns.Add("TowerHorizontalBoltEdgeDistance", GetTypeNullable(Me.TowerHorizontalBoltEdgeDistance))
    '    '    GenerateDataTable.Columns.Add("TowerShortHorizontalBoltEdgeDistance", GetTypeNullable(Me.TowerShortHorizontalBoltEdgeDistance))
    '    '    GenerateDataTable.Columns.Add("TowerDiagonalGageG1Distance", GetTypeNullable(Me.TowerDiagonalGageG1Distance))
    '    '    GenerateDataTable.Columns.Add("TowerTopGirtGageG1Distance", GetTypeNullable(Me.TowerTopGirtGageG1Distance))
    '    '    GenerateDataTable.Columns.Add("TowerBotGirtGageG1Distance", GetTypeNullable(Me.TowerBotGirtGageG1Distance))
    '    '    GenerateDataTable.Columns.Add("TowerInnerGirtGageG1Distance", GetTypeNullable(Me.TowerInnerGirtGageG1Distance))
    '    '    GenerateDataTable.Columns.Add("TowerHorizontalGageG1Distance", GetTypeNullable(Me.TowerHorizontalGageG1Distance))
    '    '    GenerateDataTable.Columns.Add("TowerShortHorizontalGageG1Distance", GetTypeNullable(Me.TowerShortHorizontalGageG1Distance))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHorizontalBoltGrade", GetTypeNullable(Me.TowerRedundantHorizontalBoltGrade))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHorizontalBoltSize", GetTypeNullable(Me.TowerRedundantHorizontalBoltSize))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHorizontalNumBolts", GetTypeNullable(Me.TowerRedundantHorizontalNumBolts))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHorizontalBoltEdgeDistance", GetTypeNullable(Me.TowerRedundantHorizontalBoltEdgeDistance))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHorizontalGageG1Distance", GetTypeNullable(Me.TowerRedundantHorizontalGageG1Distance))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHorizontalNetWidthDeduct", GetTypeNullable(Me.TowerRedundantHorizontalNetWidthDeduct))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHorizontalUFactor", GetTypeNullable(Me.TowerRedundantHorizontalUFactor))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantDiagonalBoltGrade", GetTypeNullable(Me.TowerRedundantDiagonalBoltGrade))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantDiagonalBoltSize", GetTypeNullable(Me.TowerRedundantDiagonalBoltSize))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantDiagonalNumBolts", GetTypeNullable(Me.TowerRedundantDiagonalNumBolts))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantDiagonalBoltEdgeDistance", GetTypeNullable(Me.TowerRedundantDiagonalBoltEdgeDistance))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantDiagonalGageG1Distance", GetTypeNullable(Me.TowerRedundantDiagonalGageG1Distance))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantDiagonalNetWidthDeduct", GetTypeNullable(Me.TowerRedundantDiagonalNetWidthDeduct))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantDiagonalUFactor", GetTypeNullable(Me.TowerRedundantDiagonalUFactor))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantSubDiagonalBoltGrade", GetTypeNullable(Me.TowerRedundantSubDiagonalBoltGrade))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantSubDiagonalBoltSize", GetTypeNullable(Me.TowerRedundantSubDiagonalBoltSize))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantSubDiagonalNumBolts", GetTypeNullable(Me.TowerRedundantSubDiagonalNumBolts))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantSubDiagonalBoltEdgeDistance", GetTypeNullable(Me.TowerRedundantSubDiagonalBoltEdgeDistance))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantSubDiagonalGageG1Distance", GetTypeNullable(Me.TowerRedundantSubDiagonalGageG1Distance))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantSubDiagonalNetWidthDeduct", GetTypeNullable(Me.TowerRedundantSubDiagonalNetWidthDeduct))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantSubDiagonalUFactor", GetTypeNullable(Me.TowerRedundantSubDiagonalUFactor))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantSubHorizontalBoltGrade", GetTypeNullable(Me.TowerRedundantSubHorizontalBoltGrade))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantSubHorizontalBoltSize", GetTypeNullable(Me.TowerRedundantSubHorizontalBoltSize))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantSubHorizontalNumBolts", GetTypeNullable(Me.TowerRedundantSubHorizontalNumBolts))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantSubHorizontalBoltEdgeDistance", GetTypeNullable(Me.TowerRedundantSubHorizontalBoltEdgeDistance))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantSubHorizontalGageG1Distance", GetTypeNullable(Me.TowerRedundantSubHorizontalGageG1Distance))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantSubHorizontalNetWidthDeduct", GetTypeNullable(Me.TowerRedundantSubHorizontalNetWidthDeduct))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantSubHorizontalUFactor", GetTypeNullable(Me.TowerRedundantSubHorizontalUFactor))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantVerticalBoltGrade", GetTypeNullable(Me.TowerRedundantVerticalBoltGrade))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantVerticalBoltSize", GetTypeNullable(Me.TowerRedundantVerticalBoltSize))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantVerticalNumBolts", GetTypeNullable(Me.TowerRedundantVerticalNumBolts))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantVerticalBoltEdgeDistance", GetTypeNullable(Me.TowerRedundantVerticalBoltEdgeDistance))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantVerticalGageG1Distance", GetTypeNullable(Me.TowerRedundantVerticalGageG1Distance))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantVerticalNetWidthDeduct", GetTypeNullable(Me.TowerRedundantVerticalNetWidthDeduct))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantVerticalUFactor", GetTypeNullable(Me.TowerRedundantVerticalUFactor))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHipBoltGrade", GetTypeNullable(Me.TowerRedundantHipBoltGrade))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHipBoltSize", GetTypeNullable(Me.TowerRedundantHipBoltSize))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHipNumBolts", GetTypeNullable(Me.TowerRedundantHipNumBolts))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHipBoltEdgeDistance", GetTypeNullable(Me.TowerRedundantHipBoltEdgeDistance))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHipGageG1Distance", GetTypeNullable(Me.TowerRedundantHipGageG1Distance))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHipNetWidthDeduct", GetTypeNullable(Me.TowerRedundantHipNetWidthDeduct))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHipUFactor", GetTypeNullable(Me.TowerRedundantHipUFactor))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHipDiagonalBoltGrade", GetTypeNullable(Me.TowerRedundantHipDiagonalBoltGrade))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHipDiagonalBoltSize", GetTypeNullable(Me.TowerRedundantHipDiagonalBoltSize))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHipDiagonalNumBolts", GetTypeNullable(Me.TowerRedundantHipDiagonalNumBolts))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHipDiagonalBoltEdgeDistance", GetTypeNullable(Me.TowerRedundantHipDiagonalBoltEdgeDistance))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHipDiagonalGageG1Distance", GetTypeNullable(Me.TowerRedundantHipDiagonalGageG1Distance))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHipDiagonalNetWidthDeduct", GetTypeNullable(Me.TowerRedundantHipDiagonalNetWidthDeduct))
    '    '    GenerateDataTable.Columns.Add("TowerRedundantHipDiagonalUFactor", GetTypeNullable(Me.TowerRedundantHipDiagonalUFactor))
    '    '    GenerateDataTable.Columns.Add("TowerDiagonalOutOfPlaneRestraint", GetTypeNullable(Me.TowerDiagonalOutOfPlaneRestraint))
    '    '    GenerateDataTable.Columns.Add("TowerTopGirtOutOfPlaneRestraint", GetTypeNullable(Me.TowerTopGirtOutOfPlaneRestraint))
    '    '    GenerateDataTable.Columns.Add("TowerBottomGirtOutOfPlaneRestraint", GetTypeNullable(Me.TowerBottomGirtOutOfPlaneRestraint))
    '    '    GenerateDataTable.Columns.Add("TowerMidGirtOutOfPlaneRestraint", GetTypeNullable(Me.TowerMidGirtOutOfPlaneRestraint))
    '    '    GenerateDataTable.Columns.Add("TowerHorizontalOutOfPlaneRestraint", GetTypeNullable(Me.TowerHorizontalOutOfPlaneRestraint))
    '    '    GenerateDataTable.Columns.Add("TowerSecondaryHorizontalOutOfPlaneRestraint", GetTypeNullable(Me.TowerSecondaryHorizontalOutOfPlaneRestraint))
    '    '    GenerateDataTable.Columns.Add("TowerUniqueFlag", GetTypeNullable(Me.TowerUniqueFlag))
    '    '    GenerateDataTable.Columns.Add("TowerDiagOffsetNEY", GetTypeNullable(Me.TowerDiagOffsetNEY))
    '    '    GenerateDataTable.Columns.Add("TowerDiagOffsetNEX", GetTypeNullable(Me.TowerDiagOffsetNEX))
    '    '    GenerateDataTable.Columns.Add("TowerDiagOffsetPEY", GetTypeNullable(Me.TowerDiagOffsetPEY))
    '    '    GenerateDataTable.Columns.Add("TowerDiagOffsetPEX", GetTypeNullable(Me.TowerDiagOffsetPEX))
    '    '    GenerateDataTable.Columns.Add("TowerKbraceOffsetNEY", GetTypeNullable(Me.TowerKbraceOffsetNEY))
    '    '    GenerateDataTable.Columns.Add("TowerKbraceOffsetNEX", GetTypeNullable(Me.TowerKbraceOffsetNEX))
    '    '    GenerateDataTable.Columns.Add("TowerKbraceOffsetPEY", GetTypeNullable(Me.TowerKbraceOffsetPEY))
    '    '    GenerateDataTable.Columns.Add("TowerKbraceOffsetPEX", GetTypeNullable(Me.TowerKbraceOffsetPEX))

    '    Return GenerateDataTable
    'End Function

    'Public Sub GenerateDataRow(ByVal DT As DataTable)
    '    Dim newRow As DataRow = DT.NewRow()

    '    'If Me.ID.HasValue Then newRow("ID") = Me.ID.Value
    '    If Me.TowerRec.HasValue Then newRow("TowerRec") = Me.TowerRec.Value
    '    '    If Not Me.TowerDatabase = "" Then newRow("TowerDatabase") = Me.TowerDatabase
    '    '    If Not Me.TowerName = "" Then newRow("TowerName") = Me.TowerName
    '    '    If Me.TowerHeight.HasValue Then newRow("TowerHeight") = Me.TowerHeight.Value
    '    '    If Me.TowerFaceWidth.HasValue Then newRow("TowerFaceWidth") = Me.TowerFaceWidth.Value
    '    '    If Me.TowerNumSections.HasValue Then newRow("TowerNumSections") = Me.TowerNumSections.Value
    '    '    If Me.TowerSectionLength.HasValue Then newRow("TowerSectionLength") = Me.TowerSectionLength.Value
    '    '    If Me.TowerDiagonalSpacing.HasValue Then newRow("TowerDiagonalSpacing") = Me.TowerDiagonalSpacing.Value
    '    '    If Me.TowerDiagonalSpacingEx.HasValue Then newRow("TowerDiagonalSpacingEx") = Me.TowerDiagonalSpacingEx.Value
    '    If Not Me.TowerBraceType = "" Then newRow("TowerBraceType") = Me.TowerBraceType
    '    '    If Me.TowerFaceBevel.HasValue Then newRow("TowerFaceBevel") = Me.TowerFaceBevel.Value
    '    '    If Me.TowerTopGirtOffset.HasValue Then newRow("TowerTopGirtOffset") = Me.TowerTopGirtOffset.Value
    '    '    If Me.TowerBotGirtOffset.HasValue Then newRow("TowerBotGirtOffset") = Me.TowerBotGirtOffset.Value
    '    '    If Me.TowerHasKBraceEndPanels.HasValue Then newRow("TowerHasKBraceEndPanels") = Me.TowerHasKBraceEndPanels.Value
    '    '    If Me.TowerHasHorizontals.HasValue Then newRow("TowerHasHorizontals") = Me.TowerHasHorizontals.Value
    '    '    If Not Me.TowerLegType = "" Then newRow("TowerLegType") = Me.TowerLegType
    '    '    If Not Me.TowerLegSize = "" Then newRow("TowerLegSize") = Me.TowerLegSize
    '    '    If Me.TowerLegGrade.HasValue Then newRow("TowerLegGrade") = Me.TowerLegGrade.Value
    '    '    If Not Me.TowerLegMatlGrade = "" Then newRow("TowerLegMatlGrade") = Me.TowerLegMatlGrade
    '    '    If Me.TowerDiagonalGrade.HasValue Then newRow("TowerDiagonalGrade") = Me.TowerDiagonalGrade.Value
    '    '    If Not Me.TowerDiagonalMatlGrade = "" Then newRow("TowerDiagonalMatlGrade") = Me.TowerDiagonalMatlGrade
    '    '    If Me.TowerInnerBracingGrade.HasValue Then newRow("TowerInnerBracingGrade") = Me.TowerInnerBracingGrade.Value
    '    '    If Not Me.TowerInnerBracingMatlGrade = "" Then newRow("TowerInnerBracingMatlGrade") = Me.TowerInnerBracingMatlGrade
    '    '    If Me.TowerTopGirtGrade.HasValue Then newRow("TowerTopGirtGrade") = Me.TowerTopGirtGrade.Value
    '    '    If Not Me.TowerTopGirtMatlGrade = "" Then newRow("TowerTopGirtMatlGrade") = Me.TowerTopGirtMatlGrade
    '    '    If Me.TowerBotGirtGrade.HasValue Then newRow("TowerBotGirtGrade") = Me.TowerBotGirtGrade.Value
    '    '    If Not Me.TowerBotGirtMatlGrade = "" Then newRow("TowerBotGirtMatlGrade") = Me.TowerBotGirtMatlGrade
    '    '    If Me.TowerInnerGirtGrade.HasValue Then newRow("TowerInnerGirtGrade") = Me.TowerInnerGirtGrade.Value
    '    '    If Not Me.TowerInnerGirtMatlGrade = "" Then newRow("TowerInnerGirtMatlGrade") = Me.TowerInnerGirtMatlGrade
    '    '    If Me.TowerLongHorizontalGrade.HasValue Then newRow("TowerLongHorizontalGrade") = Me.TowerLongHorizontalGrade.Value
    '    '    If Not Me.TowerLongHorizontalMatlGrade = "" Then newRow("TowerLongHorizontalMatlGrade") = Me.TowerLongHorizontalMatlGrade
    '    '    If Me.TowerShortHorizontalGrade.HasValue Then newRow("TowerShortHorizontalGrade") = Me.TowerShortHorizontalGrade.Value
    '    '    If Not Me.TowerShortHorizontalMatlGrade = "" Then newRow("TowerShortHorizontalMatlGrade") = Me.TowerShortHorizontalMatlGrade
    '    '    If Not Me.TowerDiagonalType = "" Then newRow("TowerDiagonalType") = Me.TowerDiagonalType
    '    '    If Not Me.TowerDiagonalSize = "" Then newRow("TowerDiagonalSize") = Me.TowerDiagonalSize
    '    '    If Not Me.TowerInnerBracingType = "" Then newRow("TowerInnerBracingType") = Me.TowerInnerBracingType
    '    '    If Not Me.TowerInnerBracingSize = "" Then newRow("TowerInnerBracingSize") = Me.TowerInnerBracingSize
    '    '    If Not Me.TowerTopGirtType = "" Then newRow("TowerTopGirtType") = Me.TowerTopGirtType
    '    '    If Not Me.TowerTopGirtSize = "" Then newRow("TowerTopGirtSize") = Me.TowerTopGirtSize
    '    '    If Not Me.TowerBotGirtType = "" Then newRow("TowerBotGirtType") = Me.TowerBotGirtType
    '    '    If Not Me.TowerBotGirtSize = "" Then newRow("TowerBotGirtSize") = Me.TowerBotGirtSize
    '    '    If Me.TowerNumInnerGirts.HasValue Then newRow("TowerNumInnerGirts") = Me.TowerNumInnerGirts.Value
    '    '    If Not Me.TowerInnerGirtType = "" Then newRow("TowerInnerGirtType") = Me.TowerInnerGirtType
    '    '    If Not Me.TowerInnerGirtSize = "" Then newRow("TowerInnerGirtSize") = Me.TowerInnerGirtSize
    '    '    If Not Me.TowerLongHorizontalType = "" Then newRow("TowerLongHorizontalType") = Me.TowerLongHorizontalType
    '    '    If Not Me.TowerLongHorizontalSize = "" Then newRow("TowerLongHorizontalSize") = Me.TowerLongHorizontalSize
    '    '    If Not Me.TowerShortHorizontalType = "" Then newRow("TowerShortHorizontalType") = Me.TowerShortHorizontalType
    '    '    If Not Me.TowerShortHorizontalSize = "" Then newRow("TowerShortHorizontalSize") = Me.TowerShortHorizontalSize
    '    '    If Me.TowerRedundantGrade.HasValue Then newRow("TowerRedundantGrade") = Me.TowerRedundantGrade.Value
    '    '    If Not Me.TowerRedundantMatlGrade = "" Then newRow("TowerRedundantMatlGrade") = Me.TowerRedundantMatlGrade
    '    '    If Not Me.TowerRedundantType = "" Then newRow("TowerRedundantType") = Me.TowerRedundantType
    '    '    If Not Me.TowerRedundantDiagType = "" Then newRow("TowerRedundantDiagType") = Me.TowerRedundantDiagType
    '    '    If Not Me.TowerRedundantSubDiagonalType = "" Then newRow("TowerRedundantSubDiagonalType") = Me.TowerRedundantSubDiagonalType
    '    '    If Not Me.TowerRedundantSubHorizontalType = "" Then newRow("TowerRedundantSubHorizontalType") = Me.TowerRedundantSubHorizontalType
    '    '    If Not Me.TowerRedundantVerticalType = "" Then newRow("TowerRedundantVerticalType") = Me.TowerRedundantVerticalType
    '    '    If Not Me.TowerRedundantHipType = "" Then newRow("TowerRedundantHipType") = Me.TowerRedundantHipType
    '    '    If Not Me.TowerRedundantHipDiagonalType = "" Then newRow("TowerRedundantHipDiagonalType") = Me.TowerRedundantHipDiagonalType
    '    '    If Not Me.TowerRedundantHorizontalSize = "" Then newRow("TowerRedundantHorizontalSize") = Me.TowerRedundantHorizontalSize
    '    '    If Not Me.TowerRedundantHorizontalSize2 = "" Then newRow("TowerRedundantHorizontalSize2") = Me.TowerRedundantHorizontalSize2
    '    '    If Not Me.TowerRedundantHorizontalSize3 = "" Then newRow("TowerRedundantHorizontalSize3") = Me.TowerRedundantHorizontalSize3
    '    '    If Not Me.TowerRedundantHorizontalSize4 = "" Then newRow("TowerRedundantHorizontalSize4") = Me.TowerRedundantHorizontalSize4
    '    '    If Not Me.TowerRedundantDiagonalSize = "" Then newRow("TowerRedundantDiagonalSize") = Me.TowerRedundantDiagonalSize
    '    '    If Not Me.TowerRedundantDiagonalSize2 = "" Then newRow("TowerRedundantDiagonalSize2") = Me.TowerRedundantDiagonalSize2
    '    '    If Not Me.TowerRedundantDiagonalSize3 = "" Then newRow("TowerRedundantDiagonalSize3") = Me.TowerRedundantDiagonalSize3
    '    '    If Not Me.TowerRedundantDiagonalSize4 = "" Then newRow("TowerRedundantDiagonalSize4") = Me.TowerRedundantDiagonalSize4
    '    '    If Not Me.TowerRedundantSubHorizontalSize = "" Then newRow("TowerRedundantSubHorizontalSize") = Me.TowerRedundantSubHorizontalSize
    '    '    If Not Me.TowerRedundantSubDiagonalSize = "" Then newRow("TowerRedundantSubDiagonalSize") = Me.TowerRedundantSubDiagonalSize
    '    '    If Me.TowerSubDiagLocation.HasValue Then newRow("TowerSubDiagLocation") = Me.TowerSubDiagLocation.Value
    '    '    If Not Me.TowerRedundantVerticalSize = "" Then newRow("TowerRedundantVerticalSize") = Me.TowerRedundantVerticalSize
    '    '    If Not Me.TowerRedundantHipSize = "" Then newRow("TowerRedundantHipSize") = Me.TowerRedundantHipSize
    '    '    If Not Me.TowerRedundantHipSize2 = "" Then newRow("TowerRedundantHipSize2") = Me.TowerRedundantHipSize2
    '    '    If Not Me.TowerRedundantHipSize3 = "" Then newRow("TowerRedundantHipSize3") = Me.TowerRedundantHipSize3
    '    '    If Not Me.TowerRedundantHipSize4 = "" Then newRow("TowerRedundantHipSize4") = Me.TowerRedundantHipSize4
    '    '    If Not Me.TowerRedundantHipDiagonalSize = "" Then newRow("TowerRedundantHipDiagonalSize") = Me.TowerRedundantHipDiagonalSize
    '    '    If Not Me.TowerRedundantHipDiagonalSize2 = "" Then newRow("TowerRedundantHipDiagonalSize2") = Me.TowerRedundantHipDiagonalSize2
    '    '    If Not Me.TowerRedundantHipDiagonalSize3 = "" Then newRow("TowerRedundantHipDiagonalSize3") = Me.TowerRedundantHipDiagonalSize3
    '    '    If Not Me.TowerRedundantHipDiagonalSize4 = "" Then newRow("TowerRedundantHipDiagonalSize4") = Me.TowerRedundantHipDiagonalSize4
    '    '    If Me.TowerSWMult.HasValue Then newRow("TowerSWMult") = Me.TowerSWMult.Value
    '    '    If Me.TowerWPMult.HasValue Then newRow("TowerWPMult") = Me.TowerWPMult.Value
    '    '    If Me.TowerAutoCalcKSingleAngle.HasValue Then newRow("TowerAutoCalcKSingleAngle") = Me.TowerAutoCalcKSingleAngle.Value
    '    '    If Me.TowerAutoCalcKSolidRound.HasValue Then newRow("TowerAutoCalcKSolidRound") = Me.TowerAutoCalcKSolidRound.Value
    '    '    If Me.TowerAfGusset.HasValue Then newRow("TowerAfGusset") = Me.TowerAfGusset.Value
    '    '    If Me.TowerTfGusset.HasValue Then newRow("TowerTfGusset") = Me.TowerTfGusset.Value
    '    '    If Me.TowerGussetBoltEdgeDistance.HasValue Then newRow("TowerGussetBoltEdgeDistance") = Me.TowerGussetBoltEdgeDistance.Value
    '    '    If Me.TowerGussetGrade.HasValue Then newRow("TowerGussetGrade") = Me.TowerGussetGrade.Value
    '    '    If Not Me.TowerGussetMatlGrade = "" Then newRow("TowerGussetMatlGrade") = Me.TowerGussetMatlGrade
    '    '    If Me.TowerAfMult.HasValue Then newRow("TowerAfMult") = Me.TowerAfMult.Value
    '    '    If Me.TowerArMult.HasValue Then newRow("TowerArMult") = Me.TowerArMult.Value
    '    '    If Me.TowerFlatIPAPole.HasValue Then newRow("TowerFlatIPAPole") = Me.TowerFlatIPAPole.Value
    '    '    If Me.TowerRoundIPAPole.HasValue Then newRow("TowerRoundIPAPole") = Me.TowerRoundIPAPole.Value
    '    '    If Me.TowerFlatIPALeg.HasValue Then newRow("TowerFlatIPALeg") = Me.TowerFlatIPALeg.Value
    '    '    If Me.TowerRoundIPALeg.HasValue Then newRow("TowerRoundIPALeg") = Me.TowerRoundIPALeg.Value
    '    '    If Me.TowerFlatIPAHorizontal.HasValue Then newRow("TowerFlatIPAHorizontal") = Me.TowerFlatIPAHorizontal.Value
    '    '    If Me.TowerRoundIPAHorizontal.HasValue Then newRow("TowerRoundIPAHorizontal") = Me.TowerRoundIPAHorizontal.Value
    '    '    If Me.TowerFlatIPADiagonal.HasValue Then newRow("TowerFlatIPADiagonal") = Me.TowerFlatIPADiagonal.Value
    '    '    If Me.TowerRoundIPADiagonal.HasValue Then newRow("TowerRoundIPADiagonal") = Me.TowerRoundIPADiagonal.Value
    '    '    If Me.TowerCSA_S37_SpeedUpFactor.HasValue Then newRow("TowerCSA_S37_SpeedUpFactor") = Me.TowerCSA_S37_SpeedUpFactor.Value
    '    '    If Me.TowerKLegs.HasValue Then newRow("TowerKLegs") = Me.TowerKLegs.Value
    '    '    If Me.TowerKXBracedDiags.HasValue Then newRow("TowerKXBracedDiags") = Me.TowerKXBracedDiags.Value
    '    '    If Me.TowerKKBracedDiags.HasValue Then newRow("TowerKKBracedDiags") = Me.TowerKKBracedDiags.Value
    '    '    If Me.TowerKZBracedDiags.HasValue Then newRow("TowerKZBracedDiags") = Me.TowerKZBracedDiags.Value
    '    '    If Me.TowerKHorzs.HasValue Then newRow("TowerKHorzs") = Me.TowerKHorzs.Value
    '    '    If Me.TowerKSecHorzs.HasValue Then newRow("TowerKSecHorzs") = Me.TowerKSecHorzs.Value
    '    '    If Me.TowerKGirts.HasValue Then newRow("TowerKGirts") = Me.TowerKGirts.Value
    '    '    If Me.TowerKInners.HasValue Then newRow("TowerKInners") = Me.TowerKInners.Value
    '    '    If Me.TowerKXBracedDiagsY.HasValue Then newRow("TowerKXBracedDiagsY") = Me.TowerKXBracedDiagsY.Value
    '    '    If Me.TowerKKBracedDiagsY.HasValue Then newRow("TowerKKBracedDiagsY") = Me.TowerKKBracedDiagsY.Value
    '    '    If Me.TowerKZBracedDiagsY.HasValue Then newRow("TowerKZBracedDiagsY") = Me.TowerKZBracedDiagsY.Value
    '    '    If Me.TowerKHorzsY.HasValue Then newRow("TowerKHorzsY") = Me.TowerKHorzsY.Value
    '    '    If Me.TowerKSecHorzsY.HasValue Then newRow("TowerKSecHorzsY") = Me.TowerKSecHorzsY.Value
    '    '    If Me.TowerKGirtsY.HasValue Then newRow("TowerKGirtsY") = Me.TowerKGirtsY.Value
    '    '    If Me.TowerKInnersY.HasValue Then newRow("TowerKInnersY") = Me.TowerKInnersY.Value
    '    '    If Me.TowerKRedHorz.HasValue Then newRow("TowerKRedHorz") = Me.TowerKRedHorz.Value
    '    '    If Me.TowerKRedDiag.HasValue Then newRow("TowerKRedDiag") = Me.TowerKRedDiag.Value
    '    '    If Me.TowerKRedSubDiag.HasValue Then newRow("TowerKRedSubDiag") = Me.TowerKRedSubDiag.Value
    '    '    If Me.TowerKRedSubHorz.HasValue Then newRow("TowerKRedSubHorz") = Me.TowerKRedSubHorz.Value
    '    '    If Me.TowerKRedVert.HasValue Then newRow("TowerKRedVert") = Me.TowerKRedVert.Value
    '    '    If Me.TowerKRedHip.HasValue Then newRow("TowerKRedHip") = Me.TowerKRedHip.Value
    '    '    If Me.TowerKRedHipDiag.HasValue Then newRow("TowerKRedHipDiag") = Me.TowerKRedHipDiag.Value
    '    '    If Me.TowerKTLX.HasValue Then newRow("TowerKTLX") = Me.TowerKTLX.Value
    '    '    If Me.TowerKTLZ.HasValue Then newRow("TowerKTLZ") = Me.TowerKTLZ.Value
    '    '    If Me.TowerKTLLeg.HasValue Then newRow("TowerKTLLeg") = Me.TowerKTLLeg.Value
    '    '    If Me.TowerInnerKTLX.HasValue Then newRow("TowerInnerKTLX") = Me.TowerInnerKTLX.Value
    '    '    If Me.TowerInnerKTLZ.HasValue Then newRow("TowerInnerKTLZ") = Me.TowerInnerKTLZ.Value
    '    '    If Me.TowerInnerKTLLeg.HasValue Then newRow("TowerInnerKTLLeg") = Me.TowerInnerKTLLeg.Value
    '    '    If Not Me.TowerStitchBoltLocationHoriz = "" Then newRow("TowerStitchBoltLocationHoriz") = Me.TowerStitchBoltLocationHoriz
    '    '    If Not Me.TowerStitchBoltLocationDiag = "" Then newRow("TowerStitchBoltLocationDiag") = Me.TowerStitchBoltLocationDiag
    '    '    If Not Me.TowerStitchBoltLocationRed = "" Then newRow("TowerStitchBoltLocationRed") = Me.TowerStitchBoltLocationRed
    '    '    If Me.TowerStitchSpacing.HasValue Then newRow("TowerStitchSpacing") = Me.TowerStitchSpacing.Value
    '    '    If Me.TowerStitchSpacingDiag.HasValue Then newRow("TowerStitchSpacingDiag") = Me.TowerStitchSpacingDiag.Value
    '    '    If Me.TowerStitchSpacingHorz.HasValue Then newRow("TowerStitchSpacingHorz") = Me.TowerStitchSpacingHorz.Value
    '    '    If Me.TowerStitchSpacingRed.HasValue Then newRow("TowerStitchSpacingRed") = Me.TowerStitchSpacingRed.Value
    '    '    If Me.TowerLegNetWidthDeduct.HasValue Then newRow("TowerLegNetWidthDeduct") = Me.TowerLegNetWidthDeduct.Value
    '    '    If Me.TowerLegUFactor.HasValue Then newRow("TowerLegUFactor") = Me.TowerLegUFactor.Value
    '    '    If Me.TowerDiagonalNetWidthDeduct.HasValue Then newRow("TowerDiagonalNetWidthDeduct") = Me.TowerDiagonalNetWidthDeduct.Value
    '    '    If Me.TowerTopGirtNetWidthDeduct.HasValue Then newRow("TowerTopGirtNetWidthDeduct") = Me.TowerTopGirtNetWidthDeduct.Value
    '    '    If Me.TowerBotGirtNetWidthDeduct.HasValue Then newRow("TowerBotGirtNetWidthDeduct") = Me.TowerBotGirtNetWidthDeduct.Value
    '    '    If Me.TowerInnerGirtNetWidthDeduct.HasValue Then newRow("TowerInnerGirtNetWidthDeduct") = Me.TowerInnerGirtNetWidthDeduct.Value
    '    '    If Me.TowerHorizontalNetWidthDeduct.HasValue Then newRow("TowerHorizontalNetWidthDeduct") = Me.TowerHorizontalNetWidthDeduct.Value
    '    '    If Me.TowerShortHorizontalNetWidthDeduct.HasValue Then newRow("TowerShortHorizontalNetWidthDeduct") = Me.TowerShortHorizontalNetWidthDeduct.Value
    '    '    If Me.TowerDiagonalUFactor.HasValue Then newRow("TowerDiagonalUFactor") = Me.TowerDiagonalUFactor.Value
    '    '    If Me.TowerTopGirtUFactor.HasValue Then newRow("TowerTopGirtUFactor") = Me.TowerTopGirtUFactor.Value
    '    '    If Me.TowerBotGirtUFactor.HasValue Then newRow("TowerBotGirtUFactor") = Me.TowerBotGirtUFactor.Value
    '    '    If Me.TowerInnerGirtUFactor.HasValue Then newRow("TowerInnerGirtUFactor") = Me.TowerInnerGirtUFactor.Value
    '    '    If Me.TowerHorizontalUFactor.HasValue Then newRow("TowerHorizontalUFactor") = Me.TowerHorizontalUFactor.Value
    '    '    If Me.TowerShortHorizontalUFactor.HasValue Then newRow("TowerShortHorizontalUFactor") = Me.TowerShortHorizontalUFactor.Value
    '    '    If Not Me.TowerLegConnType = "" Then newRow("TowerLegConnType") = Me.TowerLegConnType
    '    '    If Me.TowerLegNumBolts.HasValue Then newRow("TowerLegNumBolts") = Me.TowerLegNumBolts.Value
    '    '    If Me.TowerDiagonalNumBolts.HasValue Then newRow("TowerDiagonalNumBolts") = Me.TowerDiagonalNumBolts.Value
    '    '    If Me.TowerTopGirtNumBolts.HasValue Then newRow("TowerTopGirtNumBolts") = Me.TowerTopGirtNumBolts.Value
    '    '    If Me.TowerBotGirtNumBolts.HasValue Then newRow("TowerBotGirtNumBolts") = Me.TowerBotGirtNumBolts.Value
    '    '    If Me.TowerInnerGirtNumBolts.HasValue Then newRow("TowerInnerGirtNumBolts") = Me.TowerInnerGirtNumBolts.Value
    '    '    If Me.TowerHorizontalNumBolts.HasValue Then newRow("TowerHorizontalNumBolts") = Me.TowerHorizontalNumBolts.Value
    '    '    If Me.TowerShortHorizontalNumBolts.HasValue Then newRow("TowerShortHorizontalNumBolts") = Me.TowerShortHorizontalNumBolts.Value
    '    '    If Not Me.TowerLegBoltGrade = "" Then newRow("TowerLegBoltGrade") = Me.TowerLegBoltGrade
    '    '    If Me.TowerLegBoltSize.HasValue Then newRow("TowerLegBoltSize") = Me.TowerLegBoltSize.Value
    '    '    If Not Me.TowerDiagonalBoltGrade = "" Then newRow("TowerDiagonalBoltGrade") = Me.TowerDiagonalBoltGrade
    '    '    If Me.TowerDiagonalBoltSize.HasValue Then newRow("TowerDiagonalBoltSize") = Me.TowerDiagonalBoltSize.Value
    '    '    If Not Me.TowerTopGirtBoltGrade = "" Then newRow("TowerTopGirtBoltGrade") = Me.TowerTopGirtBoltGrade
    '    '    If Me.TowerTopGirtBoltSize.HasValue Then newRow("TowerTopGirtBoltSize") = Me.TowerTopGirtBoltSize.Value
    '    '    If Not Me.TowerBotGirtBoltGrade = "" Then newRow("TowerBotGirtBoltGrade") = Me.TowerBotGirtBoltGrade
    '    '    If Me.TowerBotGirtBoltSize.HasValue Then newRow("TowerBotGirtBoltSize") = Me.TowerBotGirtBoltSize.Value
    '    '    If Not Me.TowerInnerGirtBoltGrade = "" Then newRow("TowerInnerGirtBoltGrade") = Me.TowerInnerGirtBoltGrade
    '    '    If Me.TowerInnerGirtBoltSize.HasValue Then newRow("TowerInnerGirtBoltSize") = Me.TowerInnerGirtBoltSize.Value
    '    '    If Not Me.TowerHorizontalBoltGrade = "" Then newRow("TowerHorizontalBoltGrade") = Me.TowerHorizontalBoltGrade
    '    '    If Me.TowerHorizontalBoltSize.HasValue Then newRow("TowerHorizontalBoltSize") = Me.TowerHorizontalBoltSize.Value
    '    '    If Not Me.TowerShortHorizontalBoltGrade = "" Then newRow("TowerShortHorizontalBoltGrade") = Me.TowerShortHorizontalBoltGrade
    '    '    If Me.TowerShortHorizontalBoltSize.HasValue Then newRow("TowerShortHorizontalBoltSize") = Me.TowerShortHorizontalBoltSize.Value
    '    '    If Me.TowerLegBoltEdgeDistance.HasValue Then newRow("TowerLegBoltEdgeDistance") = Me.TowerLegBoltEdgeDistance.Value
    '    '    If Me.TowerDiagonalBoltEdgeDistance.HasValue Then newRow("TowerDiagonalBoltEdgeDistance") = Me.TowerDiagonalBoltEdgeDistance.Value
    '    '    If Me.TowerTopGirtBoltEdgeDistance.HasValue Then newRow("TowerTopGirtBoltEdgeDistance") = Me.TowerTopGirtBoltEdgeDistance.Value
    '    '    If Me.TowerBotGirtBoltEdgeDistance.HasValue Then newRow("TowerBotGirtBoltEdgeDistance") = Me.TowerBotGirtBoltEdgeDistance.Value
    '    '    If Me.TowerInnerGirtBoltEdgeDistance.HasValue Then newRow("TowerInnerGirtBoltEdgeDistance") = Me.TowerInnerGirtBoltEdgeDistance.Value
    '    '    If Me.TowerHorizontalBoltEdgeDistance.HasValue Then newRow("TowerHorizontalBoltEdgeDistance") = Me.TowerHorizontalBoltEdgeDistance.Value
    '    '    If Me.TowerShortHorizontalBoltEdgeDistance.HasValue Then newRow("TowerShortHorizontalBoltEdgeDistance") = Me.TowerShortHorizontalBoltEdgeDistance.Value
    '    '    If Me.TowerDiagonalGageG1Distance.HasValue Then newRow("TowerDiagonalGageG1Distance") = Me.TowerDiagonalGageG1Distance.Value
    '    '    If Me.TowerTopGirtGageG1Distance.HasValue Then newRow("TowerTopGirtGageG1Distance") = Me.TowerTopGirtGageG1Distance.Value
    '    '    If Me.TowerBotGirtGageG1Distance.HasValue Then newRow("TowerBotGirtGageG1Distance") = Me.TowerBotGirtGageG1Distance.Value
    '    '    If Me.TowerInnerGirtGageG1Distance.HasValue Then newRow("TowerInnerGirtGageG1Distance") = Me.TowerInnerGirtGageG1Distance.Value
    '    '    If Me.TowerHorizontalGageG1Distance.HasValue Then newRow("TowerHorizontalGageG1Distance") = Me.TowerHorizontalGageG1Distance.Value
    '    '    If Me.TowerShortHorizontalGageG1Distance.HasValue Then newRow("TowerShortHorizontalGageG1Distance") = Me.TowerShortHorizontalGageG1Distance.Value
    '    '    If Not Me.TowerRedundantHorizontalBoltGrade = "" Then newRow("TowerRedundantHorizontalBoltGrade") = Me.TowerRedundantHorizontalBoltGrade
    '    '    If Me.TowerRedundantHorizontalBoltSize.HasValue Then newRow("TowerRedundantHorizontalBoltSize") = Me.TowerRedundantHorizontalBoltSize.Value
    '    '    If Me.TowerRedundantHorizontalNumBolts.HasValue Then newRow("TowerRedundantHorizontalNumBolts") = Me.TowerRedundantHorizontalNumBolts.Value
    '    '    If Me.TowerRedundantHorizontalBoltEdgeDistance.HasValue Then newRow("TowerRedundantHorizontalBoltEdgeDistance") = Me.TowerRedundantHorizontalBoltEdgeDistance.Value
    '    '    If Me.TowerRedundantHorizontalGageG1Distance.HasValue Then newRow("TowerRedundantHorizontalGageG1Distance") = Me.TowerRedundantHorizontalGageG1Distance.Value
    '    '    If Me.TowerRedundantHorizontalNetWidthDeduct.HasValue Then newRow("TowerRedundantHorizontalNetWidthDeduct") = Me.TowerRedundantHorizontalNetWidthDeduct.Value
    '    '    If Me.TowerRedundantHorizontalUFactor.HasValue Then newRow("TowerRedundantHorizontalUFactor") = Me.TowerRedundantHorizontalUFactor.Value
    '    '    If Not Me.TowerRedundantDiagonalBoltGrade = "" Then newRow("TowerRedundantDiagonalBoltGrade") = Me.TowerRedundantDiagonalBoltGrade
    '    '    If Me.TowerRedundantDiagonalBoltSize.HasValue Then newRow("TowerRedundantDiagonalBoltSize") = Me.TowerRedundantDiagonalBoltSize.Value
    '    '    If Me.TowerRedundantDiagonalNumBolts.HasValue Then newRow("TowerRedundantDiagonalNumBolts") = Me.TowerRedundantDiagonalNumBolts.Value
    '    '    If Me.TowerRedundantDiagonalBoltEdgeDistance.HasValue Then newRow("TowerRedundantDiagonalBoltEdgeDistance") = Me.TowerRedundantDiagonalBoltEdgeDistance.Value
    '    '    If Me.TowerRedundantDiagonalGageG1Distance.HasValue Then newRow("TowerRedundantDiagonalGageG1Distance") = Me.TowerRedundantDiagonalGageG1Distance.Value
    '    '    If Me.TowerRedundantDiagonalNetWidthDeduct.HasValue Then newRow("TowerRedundantDiagonalNetWidthDeduct") = Me.TowerRedundantDiagonalNetWidthDeduct.Value
    '    '    If Me.TowerRedundantDiagonalUFactor.HasValue Then newRow("TowerRedundantDiagonalUFactor") = Me.TowerRedundantDiagonalUFactor.Value
    '    '    If Not Me.TowerRedundantSubDiagonalBoltGrade = "" Then newRow("TowerRedundantSubDiagonalBoltGrade") = Me.TowerRedundantSubDiagonalBoltGrade
    '    '    If Me.TowerRedundantSubDiagonalBoltSize.HasValue Then newRow("TowerRedundantSubDiagonalBoltSize") = Me.TowerRedundantSubDiagonalBoltSize.Value
    '    '    If Me.TowerRedundantSubDiagonalNumBolts.HasValue Then newRow("TowerRedundantSubDiagonalNumBolts") = Me.TowerRedundantSubDiagonalNumBolts.Value
    '    '    If Me.TowerRedundantSubDiagonalBoltEdgeDistance.HasValue Then newRow("TowerRedundantSubDiagonalBoltEdgeDistance") = Me.TowerRedundantSubDiagonalBoltEdgeDistance.Value
    '    '    If Me.TowerRedundantSubDiagonalGageG1Distance.HasValue Then newRow("TowerRedundantSubDiagonalGageG1Distance") = Me.TowerRedundantSubDiagonalGageG1Distance.Value
    '    '    If Me.TowerRedundantSubDiagonalNetWidthDeduct.HasValue Then newRow("TowerRedundantSubDiagonalNetWidthDeduct") = Me.TowerRedundantSubDiagonalNetWidthDeduct.Value
    '    '    If Me.TowerRedundantSubDiagonalUFactor.HasValue Then newRow("TowerRedundantSubDiagonalUFactor") = Me.TowerRedundantSubDiagonalUFactor.Value
    '    '    If Not Me.TowerRedundantSubHorizontalBoltGrade = "" Then newRow("TowerRedundantSubHorizontalBoltGrade") = Me.TowerRedundantSubHorizontalBoltGrade
    '    '    If Me.TowerRedundantSubHorizontalBoltSize.HasValue Then newRow("TowerRedundantSubHorizontalBoltSize") = Me.TowerRedundantSubHorizontalBoltSize.Value
    '    '    If Me.TowerRedundantSubHorizontalNumBolts.HasValue Then newRow("TowerRedundantSubHorizontalNumBolts") = Me.TowerRedundantSubHorizontalNumBolts.Value
    '    '    If Me.TowerRedundantSubHorizontalBoltEdgeDistance.HasValue Then newRow("TowerRedundantSubHorizontalBoltEdgeDistance") = Me.TowerRedundantSubHorizontalBoltEdgeDistance.Value
    '    '    If Me.TowerRedundantSubHorizontalGageG1Distance.HasValue Then newRow("TowerRedundantSubHorizontalGageG1Distance") = Me.TowerRedundantSubHorizontalGageG1Distance.Value
    '    '    If Me.TowerRedundantSubHorizontalNetWidthDeduct.HasValue Then newRow("TowerRedundantSubHorizontalNetWidthDeduct") = Me.TowerRedundantSubHorizontalNetWidthDeduct.Value
    '    '    If Me.TowerRedundantSubHorizontalUFactor.HasValue Then newRow("TowerRedundantSubHorizontalUFactor") = Me.TowerRedundantSubHorizontalUFactor.Value
    '    '    If Not Me.TowerRedundantVerticalBoltGrade = "" Then newRow("TowerRedundantVerticalBoltGrade") = Me.TowerRedundantVerticalBoltGrade
    '    '    If Me.TowerRedundantVerticalBoltSize.HasValue Then newRow("TowerRedundantVerticalBoltSize") = Me.TowerRedundantVerticalBoltSize.Value
    '    '    If Me.TowerRedundantVerticalNumBolts.HasValue Then newRow("TowerRedundantVerticalNumBolts") = Me.TowerRedundantVerticalNumBolts.Value
    '    '    If Me.TowerRedundantVerticalBoltEdgeDistance.HasValue Then newRow("TowerRedundantVerticalBoltEdgeDistance") = Me.TowerRedundantVerticalBoltEdgeDistance.Value
    '    '    If Me.TowerRedundantVerticalGageG1Distance.HasValue Then newRow("TowerRedundantVerticalGageG1Distance") = Me.TowerRedundantVerticalGageG1Distance.Value
    '    '    If Me.TowerRedundantVerticalNetWidthDeduct.HasValue Then newRow("TowerRedundantVerticalNetWidthDeduct") = Me.TowerRedundantVerticalNetWidthDeduct.Value
    '    '    If Me.TowerRedundantVerticalUFactor.HasValue Then newRow("TowerRedundantVerticalUFactor") = Me.TowerRedundantVerticalUFactor.Value
    '    '    If Not Me.TowerRedundantHipBoltGrade = "" Then newRow("TowerRedundantHipBoltGrade") = Me.TowerRedundantHipBoltGrade
    '    '    If Me.TowerRedundantHipBoltSize.HasValue Then newRow("TowerRedundantHipBoltSize") = Me.TowerRedundantHipBoltSize.Value
    '    '    If Me.TowerRedundantHipNumBolts.HasValue Then newRow("TowerRedundantHipNumBolts") = Me.TowerRedundantHipNumBolts.Value
    '    '    If Me.TowerRedundantHipBoltEdgeDistance.HasValue Then newRow("TowerRedundantHipBoltEdgeDistance") = Me.TowerRedundantHipBoltEdgeDistance.Value
    '    '    If Me.TowerRedundantHipGageG1Distance.HasValue Then newRow("TowerRedundantHipGageG1Distance") = Me.TowerRedundantHipGageG1Distance.Value
    '    '    If Me.TowerRedundantHipNetWidthDeduct.HasValue Then newRow("TowerRedundantHipNetWidthDeduct") = Me.TowerRedundantHipNetWidthDeduct.Value
    '    '    If Me.TowerRedundantHipUFactor.HasValue Then newRow("TowerRedundantHipUFactor") = Me.TowerRedundantHipUFactor.Value
    '    '    If Not Me.TowerRedundantHipDiagonalBoltGrade = "" Then newRow("TowerRedundantHipDiagonalBoltGrade") = Me.TowerRedundantHipDiagonalBoltGrade
    '    '    If Me.TowerRedundantHipDiagonalBoltSize.HasValue Then newRow("TowerRedundantHipDiagonalBoltSize") = Me.TowerRedundantHipDiagonalBoltSize.Value
    '    '    If Me.TowerRedundantHipDiagonalNumBolts.HasValue Then newRow("TowerRedundantHipDiagonalNumBolts") = Me.TowerRedundantHipDiagonalNumBolts.Value
    '    '    If Me.TowerRedundantHipDiagonalBoltEdgeDistance.HasValue Then newRow("TowerRedundantHipDiagonalBoltEdgeDistance") = Me.TowerRedundantHipDiagonalBoltEdgeDistance.Value
    '    '    If Me.TowerRedundantHipDiagonalGageG1Distance.HasValue Then newRow("TowerRedundantHipDiagonalGageG1Distance") = Me.TowerRedundantHipDiagonalGageG1Distance.Value
    '    '    If Me.TowerRedundantHipDiagonalNetWidthDeduct.HasValue Then newRow("TowerRedundantHipDiagonalNetWidthDeduct") = Me.TowerRedundantHipDiagonalNetWidthDeduct.Value
    '    '    If Me.TowerRedundantHipDiagonalUFactor.HasValue Then newRow("TowerRedundantHipDiagonalUFactor") = Me.TowerRedundantHipDiagonalUFactor.Value
    '    '    If Me.TowerDiagonalOutOfPlaneRestraint.HasValue Then newRow("TowerDiagonalOutOfPlaneRestraint") = Me.TowerDiagonalOutOfPlaneRestraint.Value
    '    '    If Me.TowerTopGirtOutOfPlaneRestraint.HasValue Then newRow("TowerTopGirtOutOfPlaneRestraint") = Me.TowerTopGirtOutOfPlaneRestraint.Value
    '    '    If Me.TowerBottomGirtOutOfPlaneRestraint.HasValue Then newRow("TowerBottomGirtOutOfPlaneRestraint") = Me.TowerBottomGirtOutOfPlaneRestraint.Value
    '    '    If Me.TowerMidGirtOutOfPlaneRestraint.HasValue Then newRow("TowerMidGirtOutOfPlaneRestraint") = Me.TowerMidGirtOutOfPlaneRestraint.Value
    '    '    If Me.TowerHorizontalOutOfPlaneRestraint.HasValue Then newRow("TowerHorizontalOutOfPlaneRestraint") = Me.TowerHorizontalOutOfPlaneRestraint.Value
    '    '    If Me.TowerSecondaryHorizontalOutOfPlaneRestraint.HasValue Then newRow("TowerSecondaryHorizontalOutOfPlaneRestraint") = Me.TowerSecondaryHorizontalOutOfPlaneRestraint.Value
    '    '    If Me.TowerUniqueFlag.HasValue Then newRow("TowerUniqueFlag") = Me.TowerUniqueFlag.Value
    '    '    If Me.TowerDiagOffsetNEY.HasValue Then newRow("TowerDiagOffsetNEY") = Me.TowerDiagOffsetNEY.Value
    '    '    If Me.TowerDiagOffsetNEX.HasValue Then newRow("TowerDiagOffsetNEX") = Me.TowerDiagOffsetNEX.Value
    '    '    If Me.TowerDiagOffsetPEY.HasValue Then newRow("TowerDiagOffsetPEY") = Me.TowerDiagOffsetPEY.Value
    '    '    If Me.TowerDiagOffsetPEX.HasValue Then newRow("TowerDiagOffsetPEX") = Me.TowerDiagOffsetPEX.Value
    '    '    If Me.TowerKbraceOffsetNEY.HasValue Then newRow("TowerKbraceOffsetNEY") = Me.TowerKbraceOffsetNEY.Value
    '    '    If Me.TowerKbraceOffsetNEX.HasValue Then newRow("TowerKbraceOffsetNEX") = Me.TowerKbraceOffsetNEX.Value
    '    '    If Me.TowerKbraceOffsetPEY.HasValue Then newRow("TowerKbraceOffsetPEY") = Me.TowerKbraceOffsetPEY.Value
    '    '    If Me.TowerKbraceOffsetPEX.HasValue Then newRow("TowerKbraceOffsetPEX") = Me.TowerKbraceOffsetPEX.Value


    '    DT.Rows.Add(newRow)

    'End Sub

    'Public Sub AddSQLParams(ByRef sqlCmd As SqlCommand)
    '    'sqlCmd.Parameters.AddWithValue("@ID", Me.ID)
    '    sqlCmd.Parameters.AddWithValue("@TowerRec", Me.TowerRec)
    '    sqlCmd.Parameters.AddWithValue("@TowerDatabase", Me.TowerDatabase)
    '    sqlCmd.Parameters.AddWithValue("@TowerName", Me.TowerName)
    '    sqlCmd.Parameters.AddWithValue("@TowerHeight", Me.TowerHeight)
    '    sqlCmd.Parameters.AddWithValue("@TowerFaceWidth", Me.TowerFaceWidth)
    '    sqlCmd.Parameters.AddWithValue("@TowerNumSections", Me.TowerNumSections)
    '    sqlCmd.Parameters.AddWithValue("@TowerSectionLength", Me.TowerSectionLength)
    '    sqlCmd.Parameters.AddWithValue("@TowerDiagonalSpacing", Me.TowerDiagonalSpacing)
    '    sqlCmd.Parameters.AddWithValue("@TowerDiagonalSpacingEx", Me.TowerDiagonalSpacingEx)
    '    sqlCmd.Parameters.AddWithValue("@TowerBraceType", Me.TowerBraceType)
    '    sqlCmd.Parameters.AddWithValue("@TowerFaceBevel", Me.TowerFaceBevel)
    '    sqlCmd.Parameters.AddWithValue("@TowerTopGirtOffset", Me.TowerTopGirtOffset)
    '    sqlCmd.Parameters.AddWithValue("@TowerBotGirtOffset", Me.TowerBotGirtOffset)
    '    sqlCmd.Parameters.AddWithValue("@TowerHasKBraceEndPanels", Me.TowerHasKBraceEndPanels)
    '    sqlCmd.Parameters.AddWithValue("@TowerHasHorizontals", Me.TowerHasHorizontals)
    '    sqlCmd.Parameters.AddWithValue("@TowerLegType", Me.TowerLegType)
    '    sqlCmd.Parameters.AddWithValue("@TowerLegSize", Me.TowerLegSize)
    '    sqlCmd.Parameters.AddWithValue("@TowerLegGrade", Me.TowerLegGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerLegMatlGrade", Me.TowerLegMatlGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerDiagonalGrade", Me.TowerDiagonalGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerDiagonalMatlGrade", Me.TowerDiagonalMatlGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerInnerBracingGrade", Me.TowerInnerBracingGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerInnerBracingMatlGrade", Me.TowerInnerBracingMatlGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerTopGirtGrade", Me.TowerTopGirtGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerTopGirtMatlGrade", Me.TowerTopGirtMatlGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerBotGirtGrade", Me.TowerBotGirtGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerBotGirtMatlGrade", Me.TowerBotGirtMatlGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerInnerGirtGrade", Me.TowerInnerGirtGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerInnerGirtMatlGrade", Me.TowerInnerGirtMatlGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerLongHorizontalGrade", Me.TowerLongHorizontalGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerLongHorizontalMatlGrade", Me.TowerLongHorizontalMatlGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerShortHorizontalGrade", Me.TowerShortHorizontalGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerShortHorizontalMatlGrade", Me.TowerShortHorizontalMatlGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerDiagonalType", Me.TowerDiagonalType)
    '    sqlCmd.Parameters.AddWithValue("@TowerDiagonalSize", Me.TowerDiagonalSize)
    '    sqlCmd.Parameters.AddWithValue("@TowerInnerBracingType", Me.TowerInnerBracingType)
    '    sqlCmd.Parameters.AddWithValue("@TowerInnerBracingSize", Me.TowerInnerBracingSize)
    '    sqlCmd.Parameters.AddWithValue("@TowerTopGirtType", Me.TowerTopGirtType)
    '    sqlCmd.Parameters.AddWithValue("@TowerTopGirtSize", Me.TowerTopGirtSize)
    '    sqlCmd.Parameters.AddWithValue("@TowerBotGirtType", Me.TowerBotGirtType)
    '    sqlCmd.Parameters.AddWithValue("@TowerBotGirtSize", Me.TowerBotGirtSize)
    '    sqlCmd.Parameters.AddWithValue("@TowerNumInnerGirts", Me.TowerNumInnerGirts)
    '    sqlCmd.Parameters.AddWithValue("@TowerInnerGirtType", Me.TowerInnerGirtType)
    '    sqlCmd.Parameters.AddWithValue("@TowerInnerGirtSize", Me.TowerInnerGirtSize)
    '    sqlCmd.Parameters.AddWithValue("@TowerLongHorizontalType", Me.TowerLongHorizontalType)
    '    sqlCmd.Parameters.AddWithValue("@TowerLongHorizontalSize", Me.TowerLongHorizontalSize)
    '    sqlCmd.Parameters.AddWithValue("@TowerShortHorizontalType", Me.TowerShortHorizontalType)
    '    sqlCmd.Parameters.AddWithValue("@TowerShortHorizontalSize", Me.TowerShortHorizontalSize)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantGrade", Me.TowerRedundantGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantMatlGrade", Me.TowerRedundantMatlGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantType", Me.TowerRedundantType)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantDiagType", Me.TowerRedundantDiagType)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantSubDiagonalType", Me.TowerRedundantSubDiagonalType)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantSubHorizontalType", Me.TowerRedundantSubHorizontalType)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantVerticalType", Me.TowerRedundantVerticalType)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHipType", Me.TowerRedundantHipType)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHipDiagonalType", Me.TowerRedundantHipDiagonalType)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHorizontalSize", Me.TowerRedundantHorizontalSize)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHorizontalSize2", Me.TowerRedundantHorizontalSize2)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHorizontalSize3", Me.TowerRedundantHorizontalSize3)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHorizontalSize4", Me.TowerRedundantHorizontalSize4)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantDiagonalSize", Me.TowerRedundantDiagonalSize)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantDiagonalSize2", Me.TowerRedundantDiagonalSize2)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantDiagonalSize3", Me.TowerRedundantDiagonalSize3)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantDiagonalSize4", Me.TowerRedundantDiagonalSize4)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantSubHorizontalSize", Me.TowerRedundantSubHorizontalSize)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantSubDiagonalSize", Me.TowerRedundantSubDiagonalSize)
    '    sqlCmd.Parameters.AddWithValue("@TowerSubDiagLocation", Me.TowerSubDiagLocation)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantVerticalSize", Me.TowerRedundantVerticalSize)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHipSize", Me.TowerRedundantHipSize)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHipSize2", Me.TowerRedundantHipSize2)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHipSize3", Me.TowerRedundantHipSize3)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHipSize4", Me.TowerRedundantHipSize4)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHipDiagonalSize", Me.TowerRedundantHipDiagonalSize)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHipDiagonalSize2", Me.TowerRedundantHipDiagonalSize2)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHipDiagonalSize3", Me.TowerRedundantHipDiagonalSize3)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHipDiagonalSize4", Me.TowerRedundantHipDiagonalSize4)
    '    sqlCmd.Parameters.AddWithValue("@TowerSWMult", Me.TowerSWMult)
    '    sqlCmd.Parameters.AddWithValue("@TowerWPMult", Me.TowerWPMult)
    '    sqlCmd.Parameters.AddWithValue("@TowerAutoCalcKSingleAngle", Me.TowerAutoCalcKSingleAngle)
    '    sqlCmd.Parameters.AddWithValue("@TowerAutoCalcKSolidRound", Me.TowerAutoCalcKSolidRound)
    '    sqlCmd.Parameters.AddWithValue("@TowerAfGusset", Me.TowerAfGusset)
    '    sqlCmd.Parameters.AddWithValue("@TowerTfGusset", Me.TowerTfGusset)
    '    sqlCmd.Parameters.AddWithValue("@TowerGussetBoltEdgeDistance", Me.TowerGussetBoltEdgeDistance)
    '    sqlCmd.Parameters.AddWithValue("@TowerGussetGrade", Me.TowerGussetGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerGussetMatlGrade", Me.TowerGussetMatlGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerAfMult", Me.TowerAfMult)
    '    sqlCmd.Parameters.AddWithValue("@TowerArMult", Me.TowerArMult)
    '    sqlCmd.Parameters.AddWithValue("@TowerFlatIPAPole", Me.TowerFlatIPAPole)
    '    sqlCmd.Parameters.AddWithValue("@TowerRoundIPAPole", Me.TowerRoundIPAPole)
    '    sqlCmd.Parameters.AddWithValue("@TowerFlatIPALeg", Me.TowerFlatIPALeg)
    '    sqlCmd.Parameters.AddWithValue("@TowerRoundIPALeg", Me.TowerRoundIPALeg)
    '    sqlCmd.Parameters.AddWithValue("@TowerFlatIPAHorizontal", Me.TowerFlatIPAHorizontal)
    '    sqlCmd.Parameters.AddWithValue("@TowerRoundIPAHorizontal", Me.TowerRoundIPAHorizontal)
    '    sqlCmd.Parameters.AddWithValue("@TowerFlatIPADiagonal", Me.TowerFlatIPADiagonal)
    '    sqlCmd.Parameters.AddWithValue("@TowerRoundIPADiagonal", Me.TowerRoundIPADiagonal)
    '    sqlCmd.Parameters.AddWithValue("@TowerCSA_S37_SpeedUpFactor", Me.TowerCSA_S37_SpeedUpFactor)
    '    sqlCmd.Parameters.AddWithValue("@TowerKLegs", Me.TowerKLegs)
    '    sqlCmd.Parameters.AddWithValue("@TowerKXBracedDiags", Me.TowerKXBracedDiags)
    '    sqlCmd.Parameters.AddWithValue("@TowerKKBracedDiags", Me.TowerKKBracedDiags)
    '    sqlCmd.Parameters.AddWithValue("@TowerKZBracedDiags", Me.TowerKZBracedDiags)
    '    sqlCmd.Parameters.AddWithValue("@TowerKHorzs", Me.TowerKHorzs)
    '    sqlCmd.Parameters.AddWithValue("@TowerKSecHorzs", Me.TowerKSecHorzs)
    '    sqlCmd.Parameters.AddWithValue("@TowerKGirts", Me.TowerKGirts)
    '    sqlCmd.Parameters.AddWithValue("@TowerKInners", Me.TowerKInners)
    '    sqlCmd.Parameters.AddWithValue("@TowerKXBracedDiagsY", Me.TowerKXBracedDiagsY)
    '    sqlCmd.Parameters.AddWithValue("@TowerKKBracedDiagsY", Me.TowerKKBracedDiagsY)
    '    sqlCmd.Parameters.AddWithValue("@TowerKZBracedDiagsY", Me.TowerKZBracedDiagsY)
    '    sqlCmd.Parameters.AddWithValue("@TowerKHorzsY", Me.TowerKHorzsY)
    '    sqlCmd.Parameters.AddWithValue("@TowerKSecHorzsY", Me.TowerKSecHorzsY)
    '    sqlCmd.Parameters.AddWithValue("@TowerKGirtsY", Me.TowerKGirtsY)
    '    sqlCmd.Parameters.AddWithValue("@TowerKInnersY", Me.TowerKInnersY)
    '    sqlCmd.Parameters.AddWithValue("@TowerKRedHorz", Me.TowerKRedHorz)
    '    sqlCmd.Parameters.AddWithValue("@TowerKRedDiag", Me.TowerKRedDiag)
    '    sqlCmd.Parameters.AddWithValue("@TowerKRedSubDiag", Me.TowerKRedSubDiag)
    '    sqlCmd.Parameters.AddWithValue("@TowerKRedSubHorz", Me.TowerKRedSubHorz)
    '    sqlCmd.Parameters.AddWithValue("@TowerKRedVert", Me.TowerKRedVert)
    '    sqlCmd.Parameters.AddWithValue("@TowerKRedHip", Me.TowerKRedHip)
    '    sqlCmd.Parameters.AddWithValue("@TowerKRedHipDiag", Me.TowerKRedHipDiag)
    '    sqlCmd.Parameters.AddWithValue("@TowerKTLX", Me.TowerKTLX)
    '    sqlCmd.Parameters.AddWithValue("@TowerKTLZ", Me.TowerKTLZ)
    '    sqlCmd.Parameters.AddWithValue("@TowerKTLLeg", Me.TowerKTLLeg)
    '    sqlCmd.Parameters.AddWithValue("@TowerInnerKTLX", Me.TowerInnerKTLX)
    '    sqlCmd.Parameters.AddWithValue("@TowerInnerKTLZ", Me.TowerInnerKTLZ)
    '    sqlCmd.Parameters.AddWithValue("@TowerInnerKTLLeg", Me.TowerInnerKTLLeg)
    '    sqlCmd.Parameters.AddWithValue("@TowerStitchBoltLocationHoriz", Me.TowerStitchBoltLocationHoriz)
    '    sqlCmd.Parameters.AddWithValue("@TowerStitchBoltLocationDiag", Me.TowerStitchBoltLocationDiag)
    '    sqlCmd.Parameters.AddWithValue("@TowerStitchBoltLocationRed", Me.TowerStitchBoltLocationRed)
    '    sqlCmd.Parameters.AddWithValue("@TowerStitchSpacing", Me.TowerStitchSpacing)
    '    sqlCmd.Parameters.AddWithValue("@TowerStitchSpacingDiag", Me.TowerStitchSpacingDiag)
    '    sqlCmd.Parameters.AddWithValue("@TowerStitchSpacingHorz", Me.TowerStitchSpacingHorz)
    '    sqlCmd.Parameters.AddWithValue("@TowerStitchSpacingRed", Me.TowerStitchSpacingRed)
    '    sqlCmd.Parameters.AddWithValue("@TowerLegNetWidthDeduct", Me.TowerLegNetWidthDeduct)
    '    sqlCmd.Parameters.AddWithValue("@TowerLegUFactor", Me.TowerLegUFactor)
    '    sqlCmd.Parameters.AddWithValue("@TowerDiagonalNetWidthDeduct", Me.TowerDiagonalNetWidthDeduct)
    '    sqlCmd.Parameters.AddWithValue("@TowerTopGirtNetWidthDeduct", Me.TowerTopGirtNetWidthDeduct)
    '    sqlCmd.Parameters.AddWithValue("@TowerBotGirtNetWidthDeduct", Me.TowerBotGirtNetWidthDeduct)
    '    sqlCmd.Parameters.AddWithValue("@TowerInnerGirtNetWidthDeduct", Me.TowerInnerGirtNetWidthDeduct)
    '    sqlCmd.Parameters.AddWithValue("@TowerHorizontalNetWidthDeduct", Me.TowerHorizontalNetWidthDeduct)
    '    sqlCmd.Parameters.AddWithValue("@TowerShortHorizontalNetWidthDeduct", Me.TowerShortHorizontalNetWidthDeduct)
    '    sqlCmd.Parameters.AddWithValue("@TowerDiagonalUFactor", Me.TowerDiagonalUFactor)
    '    sqlCmd.Parameters.AddWithValue("@TowerTopGirtUFactor", Me.TowerTopGirtUFactor)
    '    sqlCmd.Parameters.AddWithValue("@TowerBotGirtUFactor", Me.TowerBotGirtUFactor)
    '    sqlCmd.Parameters.AddWithValue("@TowerInnerGirtUFactor", Me.TowerInnerGirtUFactor)
    '    sqlCmd.Parameters.AddWithValue("@TowerHorizontalUFactor", Me.TowerHorizontalUFactor)
    '    sqlCmd.Parameters.AddWithValue("@TowerShortHorizontalUFactor", Me.TowerShortHorizontalUFactor)
    '    sqlCmd.Parameters.AddWithValue("@TowerLegConnType", Me.TowerLegConnType)
    '    sqlCmd.Parameters.AddWithValue("@TowerLegNumBolts", Me.TowerLegNumBolts)
    '    sqlCmd.Parameters.AddWithValue("@TowerDiagonalNumBolts", Me.TowerDiagonalNumBolts)
    '    sqlCmd.Parameters.AddWithValue("@TowerTopGirtNumBolts", Me.TowerTopGirtNumBolts)
    '    sqlCmd.Parameters.AddWithValue("@TowerBotGirtNumBolts", Me.TowerBotGirtNumBolts)
    '    sqlCmd.Parameters.AddWithValue("@TowerInnerGirtNumBolts", Me.TowerInnerGirtNumBolts)
    '    sqlCmd.Parameters.AddWithValue("@TowerHorizontalNumBolts", Me.TowerHorizontalNumBolts)
    '    sqlCmd.Parameters.AddWithValue("@TowerShortHorizontalNumBolts", Me.TowerShortHorizontalNumBolts)
    '    sqlCmd.Parameters.AddWithValue("@TowerLegBoltGrade", Me.TowerLegBoltGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerLegBoltSize", Me.TowerLegBoltSize)
    '    sqlCmd.Parameters.AddWithValue("@TowerDiagonalBoltGrade", Me.TowerDiagonalBoltGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerDiagonalBoltSize", Me.TowerDiagonalBoltSize)
    '    sqlCmd.Parameters.AddWithValue("@TowerTopGirtBoltGrade", Me.TowerTopGirtBoltGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerTopGirtBoltSize", Me.TowerTopGirtBoltSize)
    '    sqlCmd.Parameters.AddWithValue("@TowerBotGirtBoltGrade", Me.TowerBotGirtBoltGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerBotGirtBoltSize", Me.TowerBotGirtBoltSize)
    '    sqlCmd.Parameters.AddWithValue("@TowerInnerGirtBoltGrade", Me.TowerInnerGirtBoltGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerInnerGirtBoltSize", Me.TowerInnerGirtBoltSize)
    '    sqlCmd.Parameters.AddWithValue("@TowerHorizontalBoltGrade", Me.TowerHorizontalBoltGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerHorizontalBoltSize", Me.TowerHorizontalBoltSize)
    '    sqlCmd.Parameters.AddWithValue("@TowerShortHorizontalBoltGrade", Me.TowerShortHorizontalBoltGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerShortHorizontalBoltSize", Me.TowerShortHorizontalBoltSize)
    '    sqlCmd.Parameters.AddWithValue("@TowerLegBoltEdgeDistance", Me.TowerLegBoltEdgeDistance)
    '    sqlCmd.Parameters.AddWithValue("@TowerDiagonalBoltEdgeDistance", Me.TowerDiagonalBoltEdgeDistance)
    '    sqlCmd.Parameters.AddWithValue("@TowerTopGirtBoltEdgeDistance", Me.TowerTopGirtBoltEdgeDistance)
    '    sqlCmd.Parameters.AddWithValue("@TowerBotGirtBoltEdgeDistance", Me.TowerBotGirtBoltEdgeDistance)
    '    sqlCmd.Parameters.AddWithValue("@TowerInnerGirtBoltEdgeDistance", Me.TowerInnerGirtBoltEdgeDistance)
    '    sqlCmd.Parameters.AddWithValue("@TowerHorizontalBoltEdgeDistance", Me.TowerHorizontalBoltEdgeDistance)
    '    sqlCmd.Parameters.AddWithValue("@TowerShortHorizontalBoltEdgeDistance", Me.TowerShortHorizontalBoltEdgeDistance)
    '    sqlCmd.Parameters.AddWithValue("@TowerDiagonalGageG1Distance", Me.TowerDiagonalGageG1Distance)
    '    sqlCmd.Parameters.AddWithValue("@TowerTopGirtGageG1Distance", Me.TowerTopGirtGageG1Distance)
    '    sqlCmd.Parameters.AddWithValue("@TowerBotGirtGageG1Distance", Me.TowerBotGirtGageG1Distance)
    '    sqlCmd.Parameters.AddWithValue("@TowerInnerGirtGageG1Distance", Me.TowerInnerGirtGageG1Distance)
    '    sqlCmd.Parameters.AddWithValue("@TowerHorizontalGageG1Distance", Me.TowerHorizontalGageG1Distance)
    '    sqlCmd.Parameters.AddWithValue("@TowerShortHorizontalGageG1Distance", Me.TowerShortHorizontalGageG1Distance)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHorizontalBoltGrade", Me.TowerRedundantHorizontalBoltGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHorizontalBoltSize", Me.TowerRedundantHorizontalBoltSize)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHorizontalNumBolts", Me.TowerRedundantHorizontalNumBolts)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHorizontalBoltEdgeDistance", Me.TowerRedundantHorizontalBoltEdgeDistance)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHorizontalGageG1Distance", Me.TowerRedundantHorizontalGageG1Distance)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHorizontalNetWidthDeduct", Me.TowerRedundantHorizontalNetWidthDeduct)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHorizontalUFactor", Me.TowerRedundantHorizontalUFactor)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantDiagonalBoltGrade", Me.TowerRedundantDiagonalBoltGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantDiagonalBoltSize", Me.TowerRedundantDiagonalBoltSize)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantDiagonalNumBolts", Me.TowerRedundantDiagonalNumBolts)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantDiagonalBoltEdgeDistance", Me.TowerRedundantDiagonalBoltEdgeDistance)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantDiagonalGageG1Distance", Me.TowerRedundantDiagonalGageG1Distance)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantDiagonalNetWidthDeduct", Me.TowerRedundantDiagonalNetWidthDeduct)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantDiagonalUFactor", Me.TowerRedundantDiagonalUFactor)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantSubDiagonalBoltGrade", Me.TowerRedundantSubDiagonalBoltGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantSubDiagonalBoltSize", Me.TowerRedundantSubDiagonalBoltSize)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantSubDiagonalNumBolts", Me.TowerRedundantSubDiagonalNumBolts)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantSubDiagonalBoltEdgeDistance", Me.TowerRedundantSubDiagonalBoltEdgeDistance)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantSubDiagonalGageG1Distance", Me.TowerRedundantSubDiagonalGageG1Distance)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantSubDiagonalNetWidthDeduct", Me.TowerRedundantSubDiagonalNetWidthDeduct)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantSubDiagonalUFactor", Me.TowerRedundantSubDiagonalUFactor)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantSubHorizontalBoltGrade", Me.TowerRedundantSubHorizontalBoltGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantSubHorizontalBoltSize", Me.TowerRedundantSubHorizontalBoltSize)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantSubHorizontalNumBolts", Me.TowerRedundantSubHorizontalNumBolts)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantSubHorizontalBoltEdgeDistance", Me.TowerRedundantSubHorizontalBoltEdgeDistance)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantSubHorizontalGageG1Distance", Me.TowerRedundantSubHorizontalGageG1Distance)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantSubHorizontalNetWidthDeduct", Me.TowerRedundantSubHorizontalNetWidthDeduct)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantSubHorizontalUFactor", Me.TowerRedundantSubHorizontalUFactor)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantVerticalBoltGrade", Me.TowerRedundantVerticalBoltGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantVerticalBoltSize", Me.TowerRedundantVerticalBoltSize)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantVerticalNumBolts", Me.TowerRedundantVerticalNumBolts)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantVerticalBoltEdgeDistance", Me.TowerRedundantVerticalBoltEdgeDistance)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantVerticalGageG1Distance", Me.TowerRedundantVerticalGageG1Distance)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantVerticalNetWidthDeduct", Me.TowerRedundantVerticalNetWidthDeduct)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantVerticalUFactor", Me.TowerRedundantVerticalUFactor)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHipBoltGrade", Me.TowerRedundantHipBoltGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHipBoltSize", Me.TowerRedundantHipBoltSize)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHipNumBolts", Me.TowerRedundantHipNumBolts)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHipBoltEdgeDistance", Me.TowerRedundantHipBoltEdgeDistance)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHipGageG1Distance", Me.TowerRedundantHipGageG1Distance)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHipNetWidthDeduct", Me.TowerRedundantHipNetWidthDeduct)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHipUFactor", Me.TowerRedundantHipUFactor)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHipDiagonalBoltGrade", Me.TowerRedundantHipDiagonalBoltGrade)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHipDiagonalBoltSize", Me.TowerRedundantHipDiagonalBoltSize)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHipDiagonalNumBolts", Me.TowerRedundantHipDiagonalNumBolts)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHipDiagonalBoltEdgeDistance", Me.TowerRedundantHipDiagonalBoltEdgeDistance)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHipDiagonalGageG1Distance", Me.TowerRedundantHipDiagonalGageG1Distance)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHipDiagonalNetWidthDeduct", Me.TowerRedundantHipDiagonalNetWidthDeduct)
    '    sqlCmd.Parameters.AddWithValue("@TowerRedundantHipDiagonalUFactor", Me.TowerRedundantHipDiagonalUFactor)
    '    sqlCmd.Parameters.AddWithValue("@TowerDiagonalOutOfPlaneRestraint", Me.TowerDiagonalOutOfPlaneRestraint)
    '    sqlCmd.Parameters.AddWithValue("@TowerTopGirtOutOfPlaneRestraint", Me.TowerTopGirtOutOfPlaneRestraint)
    '    sqlCmd.Parameters.AddWithValue("@TowerBottomGirtOutOfPlaneRestraint", Me.TowerBottomGirtOutOfPlaneRestraint)
    '    sqlCmd.Parameters.AddWithValue("@TowerMidGirtOutOfPlaneRestraint", Me.TowerMidGirtOutOfPlaneRestraint)
    '    sqlCmd.Parameters.AddWithValue("@TowerHorizontalOutOfPlaneRestraint", Me.TowerHorizontalOutOfPlaneRestraint)
    '    sqlCmd.Parameters.AddWithValue("@TowerSecondaryHorizontalOutOfPlaneRestraint", Me.TowerSecondaryHorizontalOutOfPlaneRestraint)
    '    sqlCmd.Parameters.AddWithValue("@TowerUniqueFlag", Me.TowerUniqueFlag)
    '    sqlCmd.Parameters.AddWithValue("@TowerDiagOffsetNEY", Me.TowerDiagOffsetNEY)
    '    sqlCmd.Parameters.AddWithValue("@TowerDiagOffsetNEX", Me.TowerDiagOffsetNEX)
    '    sqlCmd.Parameters.AddWithValue("@TowerDiagOffsetPEY", Me.TowerDiagOffsetPEY)
    '    sqlCmd.Parameters.AddWithValue("@TowerDiagOffsetPEX", Me.TowerDiagOffsetPEX)
    '    sqlCmd.Parameters.AddWithValue("@TowerKbraceOffsetNEY", Me.TowerKbraceOffsetNEY)
    '    sqlCmd.Parameters.AddWithValue("@TowerKbraceOffsetNEX", Me.TowerKbraceOffsetNEX)
    '    sqlCmd.Parameters.AddWithValue("@TowerKbraceOffsetPEY", Me.TowerKbraceOffsetPEY)
    '    sqlCmd.Parameters.AddWithValue("@TowerKbraceOffsetPEX", Me.TowerKbraceOffsetPEX)
    'End Sub
    'Public Function GenerateSQL() As String
    '    Dim insertString As String = ""

    '    insertString = insertString.AddtoDBString(Me.TowerRec.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerDatabase.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerName.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerHeight.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerFaceWidth.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerNumSections.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerSectionLength.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerDiagonalSpacing.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerDiagonalSpacingEx.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerBraceType.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerFaceBevel.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerTopGirtOffset.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerBotGirtOffset.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerHasKBraceEndPanels.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerHasHorizontals.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerLegType.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerLegSize.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerLegGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerLegMatlGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerDiagonalGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerDiagonalMatlGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerInnerBracingGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerInnerBracingMatlGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerTopGirtGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerTopGirtMatlGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerBotGirtGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerBotGirtMatlGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerInnerGirtGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerInnerGirtMatlGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerLongHorizontalGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerLongHorizontalMatlGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerShortHorizontalGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerShortHorizontalMatlGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerDiagonalType.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerDiagonalSize.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerInnerBracingType.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerInnerBracingSize.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerTopGirtType.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerTopGirtSize.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerBotGirtType.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerBotGirtSize.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerNumInnerGirts.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerInnerGirtType.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerInnerGirtSize.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerLongHorizontalType.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerLongHorizontalSize.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerShortHorizontalType.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerShortHorizontalSize.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantMatlGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantType.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantDiagType.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantSubDiagonalType.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantSubHorizontalType.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantVerticalType.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHipType.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHipDiagonalType.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHorizontalSize.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHorizontalSize2.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHorizontalSize3.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHorizontalSize4.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantDiagonalSize.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantDiagonalSize2.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantDiagonalSize3.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantDiagonalSize4.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantSubHorizontalSize.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantSubDiagonalSize.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerSubDiagLocation.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantVerticalSize.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHipSize.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHipSize2.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHipSize3.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHipSize4.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHipDiagonalSize.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHipDiagonalSize2.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHipDiagonalSize3.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHipDiagonalSize4.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerSWMult.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerWPMult.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerAutoCalcKSingleAngle.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerAutoCalcKSolidRound.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerAfGusset.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerTfGusset.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerGussetBoltEdgeDistance.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerGussetGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerGussetMatlGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerAfMult.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerArMult.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerFlatIPAPole.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRoundIPAPole.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerFlatIPALeg.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRoundIPALeg.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerFlatIPAHorizontal.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRoundIPAHorizontal.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerFlatIPADiagonal.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRoundIPADiagonal.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerCSA_S37_SpeedUpFactor.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerKLegs.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerKXBracedDiags.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerKKBracedDiags.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerKZBracedDiags.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerKHorzs.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerKSecHorzs.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerKGirts.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerKInners.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerKXBracedDiagsY.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerKKBracedDiagsY.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerKZBracedDiagsY.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerKHorzsY.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerKSecHorzsY.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerKGirtsY.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerKInnersY.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerKRedHorz.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerKRedDiag.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerKRedSubDiag.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerKRedSubHorz.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerKRedVert.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerKRedHip.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerKRedHipDiag.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerKTLX.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerKTLZ.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerKTLLeg.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerInnerKTLX.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerInnerKTLZ.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerInnerKTLLeg.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerStitchBoltLocationHoriz.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerStitchBoltLocationDiag.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerStitchBoltLocationRed.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerStitchSpacing.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerStitchSpacingDiag.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerStitchSpacingHorz.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerStitchSpacingRed.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerLegNetWidthDeduct.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerLegUFactor.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerDiagonalNetWidthDeduct.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerTopGirtNetWidthDeduct.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerBotGirtNetWidthDeduct.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerInnerGirtNetWidthDeduct.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerHorizontalNetWidthDeduct.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerShortHorizontalNetWidthDeduct.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerDiagonalUFactor.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerTopGirtUFactor.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerBotGirtUFactor.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerInnerGirtUFactor.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerHorizontalUFactor.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerShortHorizontalUFactor.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerLegConnType.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerLegNumBolts.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerDiagonalNumBolts.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerTopGirtNumBolts.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerBotGirtNumBolts.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerInnerGirtNumBolts.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerHorizontalNumBolts.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerShortHorizontalNumBolts.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerLegBoltGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerLegBoltSize.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerDiagonalBoltGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerDiagonalBoltSize.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerTopGirtBoltGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerTopGirtBoltSize.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerBotGirtBoltGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerBotGirtBoltSize.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerInnerGirtBoltGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerInnerGirtBoltSize.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerHorizontalBoltGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerHorizontalBoltSize.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerShortHorizontalBoltGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerShortHorizontalBoltSize.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerLegBoltEdgeDistance.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerDiagonalBoltEdgeDistance.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerTopGirtBoltEdgeDistance.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerBotGirtBoltEdgeDistance.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerInnerGirtBoltEdgeDistance.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerHorizontalBoltEdgeDistance.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerShortHorizontalBoltEdgeDistance.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerDiagonalGageG1Distance.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerTopGirtGageG1Distance.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerBotGirtGageG1Distance.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerInnerGirtGageG1Distance.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerHorizontalGageG1Distance.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerShortHorizontalGageG1Distance.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHorizontalBoltGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHorizontalBoltSize.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHorizontalNumBolts.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHorizontalBoltEdgeDistance.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHorizontalGageG1Distance.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHorizontalNetWidthDeduct.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHorizontalUFactor.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantDiagonalBoltGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantDiagonalBoltSize.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantDiagonalNumBolts.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantDiagonalBoltEdgeDistance.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantDiagonalGageG1Distance.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantDiagonalNetWidthDeduct.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantDiagonalUFactor.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantSubDiagonalBoltGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantSubDiagonalBoltSize.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantSubDiagonalNumBolts.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantSubDiagonalBoltEdgeDistance.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantSubDiagonalGageG1Distance.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantSubDiagonalNetWidthDeduct.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantSubDiagonalUFactor.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantSubHorizontalBoltGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantSubHorizontalBoltSize.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantSubHorizontalNumBolts.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantSubHorizontalBoltEdgeDistance.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantSubHorizontalGageG1Distance.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantSubHorizontalNetWidthDeduct.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantSubHorizontalUFactor.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantVerticalBoltGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantVerticalBoltSize.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantVerticalNumBolts.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantVerticalBoltEdgeDistance.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantVerticalGageG1Distance.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantVerticalNetWidthDeduct.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantVerticalUFactor.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHipBoltGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHipBoltSize.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHipNumBolts.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHipBoltEdgeDistance.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHipGageG1Distance.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHipNetWidthDeduct.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHipUFactor.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHipDiagonalBoltGrade.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHipDiagonalBoltSize.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHipDiagonalNumBolts.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHipDiagonalBoltEdgeDistance.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHipDiagonalGageG1Distance.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHipDiagonalNetWidthDeduct.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerRedundantHipDiagonalUFactor.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerDiagonalOutOfPlaneRestraint.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerTopGirtOutOfPlaneRestraint.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerBottomGirtOutOfPlaneRestraint.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerMidGirtOutOfPlaneRestraint.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerHorizontalOutOfPlaneRestraint.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerSecondaryHorizontalOutOfPlaneRestraint.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerUniqueFlag.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerDiagOffsetNEY.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerDiagOffsetNEX.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerDiagOffsetPEY.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerDiagOffsetPEX.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerKbraceOffsetNEY.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerKbraceOffsetNEX.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerKbraceOffsetPEY.ToString)
    '    insertString = insertString.AddtoDBString(Me.TowerKbraceOffsetPEX.ToString)

    '    Return insertString
    'End Function
    'Public Function GenerateSQLColumns() As String
    '    Dim insertString As String = ""

    '    insertString = insertString.AddtoDBString("TowerRec", False)
    '    insertString = insertString.AddtoDBString("TowerDatabase", False)
    '    insertString = insertString.AddtoDBString("TowerName", False)
    '    insertString = insertString.AddtoDBString("TowerHeight", False)
    '    insertString = insertString.AddtoDBString("TowerFaceWidth", False)
    '    insertString = insertString.AddtoDBString("TowerNumSections", False)
    '    insertString = insertString.AddtoDBString("TowerSectionLength", False)
    '    insertString = insertString.AddtoDBString("TowerDiagonalSpacing", False)
    '    insertString = insertString.AddtoDBString("TowerDiagonalSpacingEx", False)
    '    insertString = insertString.AddtoDBString("TowerBraceType", False)
    '    insertString = insertString.AddtoDBString("TowerFaceBevel", False)
    '    insertString = insertString.AddtoDBString("TowerTopGirtOffset", False)
    '    insertString = insertString.AddtoDBString("TowerBotGirtOffset", False)
    '    insertString = insertString.AddtoDBString("TowerHasKBraceEndPanels", False)
    '    insertString = insertString.AddtoDBString("TowerHasHorizontals", False)
    '    insertString = insertString.AddtoDBString("TowerLegType", False)
    '    insertString = insertString.AddtoDBString("TowerLegSize", False)
    '    insertString = insertString.AddtoDBString("TowerLegGrade", False)
    '    insertString = insertString.AddtoDBString("TowerLegMatlGrade", False)
    '    insertString = insertString.AddtoDBString("TowerDiagonalGrade", False)
    '    insertString = insertString.AddtoDBString("TowerDiagonalMatlGrade", False)
    '    insertString = insertString.AddtoDBString("TowerInnerBracingGrade", False)
    '    insertString = insertString.AddtoDBString("TowerInnerBracingMatlGrade", False)
    '    insertString = insertString.AddtoDBString("TowerTopGirtGrade", False)
    '    insertString = insertString.AddtoDBString("TowerTopGirtMatlGrade", False)
    '    insertString = insertString.AddtoDBString("TowerBotGirtGrade", False)
    '    insertString = insertString.AddtoDBString("TowerBotGirtMatlGrade", False)
    '    insertString = insertString.AddtoDBString("TowerInnerGirtGrade", False)
    '    insertString = insertString.AddtoDBString("TowerInnerGirtMatlGrade", False)
    '    insertString = insertString.AddtoDBString("TowerLongHorizontalGrade", False)
    '    insertString = insertString.AddtoDBString("TowerLongHorizontalMatlGrade", False)
    '    insertString = insertString.AddtoDBString("TowerShortHorizontalGrade", False)
    '    insertString = insertString.AddtoDBString("TowerShortHorizontalMatlGrade", False)
    '    insertString = insertString.AddtoDBString("TowerDiagonalType", False)
    '    insertString = insertString.AddtoDBString("TowerDiagonalSize", False)
    '    insertString = insertString.AddtoDBString("TowerInnerBracingType", False)
    '    insertString = insertString.AddtoDBString("TowerInnerBracingSize", False)
    '    insertString = insertString.AddtoDBString("TowerTopGirtType", False)
    '    insertString = insertString.AddtoDBString("TowerTopGirtSize", False)
    '    insertString = insertString.AddtoDBString("TowerBotGirtType", False)
    '    insertString = insertString.AddtoDBString("TowerBotGirtSize", False)
    '    insertString = insertString.AddtoDBString("TowerNumInnerGirts", False)
    '    insertString = insertString.AddtoDBString("TowerInnerGirtType", False)
    '    insertString = insertString.AddtoDBString("TowerInnerGirtSize", False)
    '    insertString = insertString.AddtoDBString("TowerLongHorizontalType", False)
    '    insertString = insertString.AddtoDBString("TowerLongHorizontalSize", False)
    '    insertString = insertString.AddtoDBString("TowerShortHorizontalType", False)
    '    insertString = insertString.AddtoDBString("TowerShortHorizontalSize", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantGrade", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantMatlGrade", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantType", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantDiagType", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantSubDiagonalType", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantSubHorizontalType", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantVerticalType", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHipType", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHipDiagonalType", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHorizontalSize", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHorizontalSize2", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHorizontalSize3", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHorizontalSize4", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantDiagonalSize", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantDiagonalSize2", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantDiagonalSize3", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantDiagonalSize4", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantSubHorizontalSize", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantSubDiagonalSize", False)
    '    insertString = insertString.AddtoDBString("TowerSubDiagLocation", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantVerticalSize", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHipSize", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHipSize2", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHipSize3", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHipSize4", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHipDiagonalSize", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHipDiagonalSize2", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHipDiagonalSize3", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHipDiagonalSize4", False)
    '    insertString = insertString.AddtoDBString("TowerSWMult", False)
    '    insertString = insertString.AddtoDBString("TowerWPMult", False)
    '    insertString = insertString.AddtoDBString("TowerAutoCalcKSingleAngle", False)
    '    insertString = insertString.AddtoDBString("TowerAutoCalcKSolidRound", False)
    '    insertString = insertString.AddtoDBString("TowerAfGusset", False)
    '    insertString = insertString.AddtoDBString("TowerTfGusset", False)
    '    insertString = insertString.AddtoDBString("TowerGussetBoltEdgeDistance", False)
    '    insertString = insertString.AddtoDBString("TowerGussetGrade", False)
    '    insertString = insertString.AddtoDBString("TowerGussetMatlGrade", False)
    '    insertString = insertString.AddtoDBString("TowerAfMult", False)
    '    insertString = insertString.AddtoDBString("TowerArMult", False)
    '    insertString = insertString.AddtoDBString("TowerFlatIPAPole", False)
    '    insertString = insertString.AddtoDBString("TowerRoundIPAPole", False)
    '    insertString = insertString.AddtoDBString("TowerFlatIPALeg", False)
    '    insertString = insertString.AddtoDBString("TowerRoundIPALeg", False)
    '    insertString = insertString.AddtoDBString("TowerFlatIPAHorizontal", False)
    '    insertString = insertString.AddtoDBString("TowerRoundIPAHorizontal", False)
    '    insertString = insertString.AddtoDBString("TowerFlatIPADiagonal", False)
    '    insertString = insertString.AddtoDBString("TowerRoundIPADiagonal", False)
    '    insertString = insertString.AddtoDBString("TowerCSA_S37_SpeedUpFactor", False)
    '    insertString = insertString.AddtoDBString("TowerKLegs", False)
    '    insertString = insertString.AddtoDBString("TowerKXBracedDiags", False)
    '    insertString = insertString.AddtoDBString("TowerKKBracedDiags", False)
    '    insertString = insertString.AddtoDBString("TowerKZBracedDiags", False)
    '    insertString = insertString.AddtoDBString("TowerKHorzs", False)
    '    insertString = insertString.AddtoDBString("TowerKSecHorzs", False)
    '    insertString = insertString.AddtoDBString("TowerKGirts", False)
    '    insertString = insertString.AddtoDBString("TowerKInners", False)
    '    insertString = insertString.AddtoDBString("TowerKXBracedDiagsY", False)
    '    insertString = insertString.AddtoDBString("TowerKKBracedDiagsY", False)
    '    insertString = insertString.AddtoDBString("TowerKZBracedDiagsY", False)
    '    insertString = insertString.AddtoDBString("TowerKHorzsY", False)
    '    insertString = insertString.AddtoDBString("TowerKSecHorzsY", False)
    '    insertString = insertString.AddtoDBString("TowerKGirtsY", False)
    '    insertString = insertString.AddtoDBString("TowerKInnersY", False)
    '    insertString = insertString.AddtoDBString("TowerKRedHorz", False)
    '    insertString = insertString.AddtoDBString("TowerKRedDiag", False)
    '    insertString = insertString.AddtoDBString("TowerKRedSubDiag", False)
    '    insertString = insertString.AddtoDBString("TowerKRedSubHorz", False)
    '    insertString = insertString.AddtoDBString("TowerKRedVert", False)
    '    insertString = insertString.AddtoDBString("TowerKRedHip", False)
    '    insertString = insertString.AddtoDBString("TowerKRedHipDiag", False)
    '    insertString = insertString.AddtoDBString("TowerKTLX", False)
    '    insertString = insertString.AddtoDBString("TowerKTLZ", False)
    '    insertString = insertString.AddtoDBString("TowerKTLLeg", False)
    '    insertString = insertString.AddtoDBString("TowerInnerKTLX", False)
    '    insertString = insertString.AddtoDBString("TowerInnerKTLZ", False)
    '    insertString = insertString.AddtoDBString("TowerInnerKTLLeg", False)
    '    insertString = insertString.AddtoDBString("TowerStitchBoltLocationHoriz", False)
    '    insertString = insertString.AddtoDBString("TowerStitchBoltLocationDiag", False)
    '    insertString = insertString.AddtoDBString("TowerStitchBoltLocationRed", False)
    '    insertString = insertString.AddtoDBString("TowerStitchSpacing", False)
    '    insertString = insertString.AddtoDBString("TowerStitchSpacingDiag", False)
    '    insertString = insertString.AddtoDBString("TowerStitchSpacingHorz", False)
    '    insertString = insertString.AddtoDBString("TowerStitchSpacingRed", False)
    '    insertString = insertString.AddtoDBString("TowerLegNetWidthDeduct", False)
    '    insertString = insertString.AddtoDBString("TowerLegUFactor", False)
    '    insertString = insertString.AddtoDBString("TowerDiagonalNetWidthDeduct", False)
    '    insertString = insertString.AddtoDBString("TowerTopGirtNetWidthDeduct", False)
    '    insertString = insertString.AddtoDBString("TowerBotGirtNetWidthDeduct", False)
    '    insertString = insertString.AddtoDBString("TowerInnerGirtNetWidthDeduct", False)
    '    insertString = insertString.AddtoDBString("TowerHorizontalNetWidthDeduct", False)
    '    insertString = insertString.AddtoDBString("TowerShortHorizontalNetWidthDeduct", False)
    '    insertString = insertString.AddtoDBString("TowerDiagonalUFactor", False)
    '    insertString = insertString.AddtoDBString("TowerTopGirtUFactor", False)
    '    insertString = insertString.AddtoDBString("TowerBotGirtUFactor", False)
    '    insertString = insertString.AddtoDBString("TowerInnerGirtUFactor", False)
    '    insertString = insertString.AddtoDBString("TowerHorizontalUFactor", False)
    '    insertString = insertString.AddtoDBString("TowerShortHorizontalUFactor", False)
    '    insertString = insertString.AddtoDBString("TowerLegConnType", False)
    '    insertString = insertString.AddtoDBString("TowerLegNumBolts", False)
    '    insertString = insertString.AddtoDBString("TowerDiagonalNumBolts", False)
    '    insertString = insertString.AddtoDBString("TowerTopGirtNumBolts", False)
    '    insertString = insertString.AddtoDBString("TowerBotGirtNumBolts", False)
    '    insertString = insertString.AddtoDBString("TowerInnerGirtNumBolts", False)
    '    insertString = insertString.AddtoDBString("TowerHorizontalNumBolts", False)
    '    insertString = insertString.AddtoDBString("TowerShortHorizontalNumBolts", False)
    '    insertString = insertString.AddtoDBString("TowerLegBoltGrade", False)
    '    insertString = insertString.AddtoDBString("TowerLegBoltSize", False)
    '    insertString = insertString.AddtoDBString("TowerDiagonalBoltGrade", False)
    '    insertString = insertString.AddtoDBString("TowerDiagonalBoltSize", False)
    '    insertString = insertString.AddtoDBString("TowerTopGirtBoltGrade", False)
    '    insertString = insertString.AddtoDBString("TowerTopGirtBoltSize", False)
    '    insertString = insertString.AddtoDBString("TowerBotGirtBoltGrade", False)
    '    insertString = insertString.AddtoDBString("TowerBotGirtBoltSize", False)
    '    insertString = insertString.AddtoDBString("TowerInnerGirtBoltGrade", False)
    '    insertString = insertString.AddtoDBString("TowerInnerGirtBoltSize", False)
    '    insertString = insertString.AddtoDBString("TowerHorizontalBoltGrade", False)
    '    insertString = insertString.AddtoDBString("TowerHorizontalBoltSize", False)
    '    insertString = insertString.AddtoDBString("TowerShortHorizontalBoltGrade", False)
    '    insertString = insertString.AddtoDBString("TowerShortHorizontalBoltSize", False)
    '    insertString = insertString.AddtoDBString("TowerLegBoltEdgeDistance", False)
    '    insertString = insertString.AddtoDBString("TowerDiagonalBoltEdgeDistance", False)
    '    insertString = insertString.AddtoDBString("TowerTopGirtBoltEdgeDistance", False)
    '    insertString = insertString.AddtoDBString("TowerBotGirtBoltEdgeDistance", False)
    '    insertString = insertString.AddtoDBString("TowerInnerGirtBoltEdgeDistance", False)
    '    insertString = insertString.AddtoDBString("TowerHorizontalBoltEdgeDistance", False)
    '    insertString = insertString.AddtoDBString("TowerShortHorizontalBoltEdgeDistance", False)
    '    insertString = insertString.AddtoDBString("TowerDiagonalGageG1Distance", False)
    '    insertString = insertString.AddtoDBString("TowerTopGirtGageG1Distance", False)
    '    insertString = insertString.AddtoDBString("TowerBotGirtGageG1Distance", False)
    '    insertString = insertString.AddtoDBString("TowerInnerGirtGageG1Distance", False)
    '    insertString = insertString.AddtoDBString("TowerHorizontalGageG1Distance", False)
    '    insertString = insertString.AddtoDBString("TowerShortHorizontalGageG1Distance", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHorizontalBoltGrade", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHorizontalBoltSize", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHorizontalNumBolts", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHorizontalBoltEdgeDistance", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHorizontalGageG1Distance", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHorizontalNetWidthDeduct", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHorizontalUFactor", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantDiagonalBoltGrade", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantDiagonalBoltSize", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantDiagonalNumBolts", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantDiagonalBoltEdgeDistance", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantDiagonalGageG1Distance", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantDiagonalNetWidthDeduct", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantDiagonalUFactor", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantSubDiagonalBoltGrade", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantSubDiagonalBoltSize", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantSubDiagonalNumBolts", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantSubDiagonalBoltEdgeDistance", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantSubDiagonalGageG1Distance", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantSubDiagonalNetWidthDeduct", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantSubDiagonalUFactor", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantSubHorizontalBoltGrade", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantSubHorizontalBoltSize", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantSubHorizontalNumBolts", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantSubHorizontalBoltEdgeDistance", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantSubHorizontalGageG1Distance", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantSubHorizontalNetWidthDeduct", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantSubHorizontalUFactor", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantVerticalBoltGrade", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantVerticalBoltSize", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantVerticalNumBolts", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantVerticalBoltEdgeDistance", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantVerticalGageG1Distance", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantVerticalNetWidthDeduct", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantVerticalUFactor", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHipBoltGrade", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHipBoltSize", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHipNumBolts", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHipBoltEdgeDistance", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHipGageG1Distance", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHipNetWidthDeduct", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHipUFactor", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHipDiagonalBoltGrade", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHipDiagonalBoltSize", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHipDiagonalNumBolts", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHipDiagonalBoltEdgeDistance", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHipDiagonalGageG1Distance", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHipDiagonalNetWidthDeduct", False)
    '    insertString = insertString.AddtoDBString("TowerRedundantHipDiagonalUFactor", False)
    '    insertString = insertString.AddtoDBString("TowerDiagonalOutOfPlaneRestraint", False)
    '    insertString = insertString.AddtoDBString("TowerTopGirtOutOfPlaneRestraint", False)
    '    insertString = insertString.AddtoDBString("TowerBottomGirtOutOfPlaneRestraint", False)
    '    insertString = insertString.AddtoDBString("TowerMidGirtOutOfPlaneRestraint", False)
    '    insertString = insertString.AddtoDBString("TowerHorizontalOutOfPlaneRestraint", False)
    '    insertString = insertString.AddtoDBString("TowerSecondaryHorizontalOutOfPlaneRestraint", False)
    '    insertString = insertString.AddtoDBString("TowerUniqueFlag", False)
    '    insertString = insertString.AddtoDBString("TowerDiagOffsetNEY", False)
    '    insertString = insertString.AddtoDBString("TowerDiagOffsetNEX", False)
    '    insertString = insertString.AddtoDBString("TowerDiagOffsetPEY", False)
    '    insertString = insertString.AddtoDBString("TowerDiagOffsetPEX", False)
    '    insertString = insertString.AddtoDBString("TowerKbraceOffsetNEY", False)
    '    insertString = insertString.AddtoDBString("TowerKbraceOffsetNEX", False)
    '    insertString = insertString.AddtoDBString("TowerKbraceOffsetPEY", False)
    '    insertString = insertString.AddtoDBString("TowerKbraceOffsetPEX", False)

    '    Return insertString
    'End Function

#End Region

End Class

<DataContractAttribute()>
<KnownType(GetType(tnxGuyRecord))>
Partial Public Class tnxGuyRecord
    Inherits tnxGeometryRec

#Region "Inheritted"

    Public Overrides ReadOnly Property EDSObjectName As String = "Guy Level " & Me.Rec.ToString
    Public Overrides ReadOnly Property EDSTableName As String = "tnx.guys"

    Public Overrides Function SQLInsertValues() As String
        Return SQLInsertValues(Nothing)
    End Function

    Public Overloads Function SQLInsertValues(Optional ByVal ParentID As Integer? = Nothing) As String
        'For any EDSObject that has parent object we will need to overload the update property with a version that excepts the current version being updated.
        SQLInsertValues = ""

        SQLInsertValues = SQLInsertValues.AddtoDBString(If(ParentID Is Nothing, EDSStructure.SQLQueryIDVar(Me.EDSTableDepth - 1), ParentID.NullableToString.FormatDBValue))
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.Rec.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyHeight.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyAutoCalcKSingleAngle.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyAutoCalcKSolidRound.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyMount.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TorqueArmStyle.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyRadius.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyRadius120.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyRadius240.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyRadius360.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TorqueArmRadius.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TorqueArmLegAngle.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.Azimuth0Adjustment.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.Azimuth120Adjustment.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.Azimuth240Adjustment.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.Azimuth360Adjustment.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.Anchor0Elevation.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.Anchor120Elevation.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.Anchor240Elevation.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.Anchor360Elevation.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuySize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.Guy120Size.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.Guy240Size.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.Guy360Size.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TorqueArmSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TorqueArmSizeBot.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TorqueArmType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TorqueArmGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TorqueArmMatlGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TorqueArmKFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TorqueArmKFactorY.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyPullOffKFactorX.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyPullOffKFactorY.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyDiagKFactorX.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyDiagKFactorY.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyAutoCalc.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyAllGuysSame.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyAllGuysAnchorSame.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyIsStrapping.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyPullOffSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyPullOffSizeBot.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyPullOffType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyPullOffGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyPullOffMatlGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyUpperDiagSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyLowerDiagSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyDiagType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyDiagGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyDiagMatlGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyDiagNetWidthDeduct.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyDiagUFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyDiagNumBolts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyDiagonalOutOfPlaneRestraint.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyDiagBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyDiagBoltSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyDiagBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyDiagBoltGageDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyPullOffNetWidthDeduct.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyPullOffUFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyPullOffNumBolts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyPullOffOutOfPlaneRestraint.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyPullOffBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyPullOffBoltSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyPullOffBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyPullOffBoltGageDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyTorqueArmNetWidthDeduct.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyTorqueArmUFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyTorqueArmNumBolts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyTorqueArmOutOfPlaneRestraint.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyTorqueArmBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyTorqueArmBoltSize.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyTorqueArmBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyTorqueArmBoltGageDistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyPerCentTension.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyPerCentTension120.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyPerCentTension240.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyPerCentTension360.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyEffFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyEffFactor120.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyEffFactor240.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyEffFactor360.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyNumInsulators.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyInsulatorLength.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyInsulatorDia.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GuyInsulatorWt.NullableToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        SQLInsertFields = SQLInsertFields.AddtoDBString("tnx_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyRec")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyHeight")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyAutoCalcKSingleAngle")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyAutoCalcKSolidRound")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyMount")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TorqueArmStyle")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyRadius")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyRadius120")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyRadius240")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyRadius360")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TorqueArmRadius")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TorqueArmLegAngle")
        SQLInsertFields = SQLInsertFields.AddtoDBString("Azimuth0Adjustment")
        SQLInsertFields = SQLInsertFields.AddtoDBString("Azimuth120Adjustment")
        SQLInsertFields = SQLInsertFields.AddtoDBString("Azimuth240Adjustment")
        SQLInsertFields = SQLInsertFields.AddtoDBString("Azimuth360Adjustment")
        SQLInsertFields = SQLInsertFields.AddtoDBString("Anchor0Elevation")
        SQLInsertFields = SQLInsertFields.AddtoDBString("Anchor120Elevation")
        SQLInsertFields = SQLInsertFields.AddtoDBString("Anchor240Elevation")
        SQLInsertFields = SQLInsertFields.AddtoDBString("Anchor360Elevation")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuySize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("Guy120Size")
        SQLInsertFields = SQLInsertFields.AddtoDBString("Guy240Size")
        SQLInsertFields = SQLInsertFields.AddtoDBString("Guy360Size")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TorqueArmSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TorqueArmSizeBot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TorqueArmType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TorqueArmGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TorqueArmMatlGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TorqueArmKFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TorqueArmKFactorY")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyPullOffKFactorX")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyPullOffKFactorY")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyDiagKFactorX")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyDiagKFactorY")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyAutoCalc")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyAllGuysSame")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyAllGuysAnchorSame")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyIsStrapping")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyPullOffSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyPullOffSizeBot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyPullOffType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyPullOffGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyPullOffMatlGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyUpperDiagSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyLowerDiagSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyDiagType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyDiagGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyDiagMatlGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyDiagNetWidthDeduct")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyDiagUFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyDiagNumBolts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyDiagonalOutOfPlaneRestraint")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyDiagBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyDiagBoltSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyDiagBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyDiagBoltGageDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyPullOffNetWidthDeduct")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyPullOffUFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyPullOffNumBolts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyPullOffOutOfPlaneRestraint")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyPullOffBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyPullOffBoltSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyPullOffBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyPullOffBoltGageDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyTorqueArmNetWidthDeduct")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyTorqueArmUFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyTorqueArmNumBolts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyTorqueArmOutOfPlaneRestraint")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyTorqueArmBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyTorqueArmBoltSize")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyTorqueArmBoltEdgeDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyTorqueArmBoltGageDistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyPerCentTension")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyPerCentTension120")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyPerCentTension240")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyPerCentTension360")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyEffFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyEffFactor120")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyEffFactor240")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyEffFactor360")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyNumInsulators")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyInsulatorLength")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyInsulatorDia")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyInsulatorWt")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""

        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyRec = " & Me.Rec.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyHeight = " & Me.GuyHeight.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyAutoCalcKSingleAngle = " & Me.GuyAutoCalcKSingleAngle.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyAutoCalcKSolidRound = " & Me.GuyAutoCalcKSolidRound.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyMount = " & Me.GuyMount.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TorqueArmStyle = " & Me.TorqueArmStyle.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyRadius = " & Me.GuyRadius.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyRadius120 = " & Me.GuyRadius120.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyRadius240 = " & Me.GuyRadius240.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyRadius360 = " & Me.GuyRadius360.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TorqueArmRadius = " & Me.TorqueArmRadius.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TorqueArmLegAngle = " & Me.TorqueArmLegAngle.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("Azimuth0Adjustment = " & Me.Azimuth0Adjustment.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("Azimuth120Adjustment = " & Me.Azimuth120Adjustment.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("Azimuth240Adjustment = " & Me.Azimuth240Adjustment.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("Azimuth360Adjustment = " & Me.Azimuth360Adjustment.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("Anchor0Elevation = " & Me.Anchor0Elevation.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("Anchor120Elevation = " & Me.Anchor120Elevation.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("Anchor240Elevation = " & Me.Anchor240Elevation.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("Anchor360Elevation = " & Me.Anchor360Elevation.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuySize = " & Me.GuySize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("Guy120Size = " & Me.Guy120Size.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("Guy240Size = " & Me.Guy240Size.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("Guy360Size = " & Me.Guy360Size.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyGrade = " & Me.GuyGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TorqueArmSize = " & Me.TorqueArmSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TorqueArmSizeBot = " & Me.TorqueArmSizeBot.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TorqueArmType = " & Me.TorqueArmType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TorqueArmGrade = " & Me.TorqueArmGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TorqueArmMatlGrade = " & Me.TorqueArmMatlGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TorqueArmKFactor = " & Me.TorqueArmKFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TorqueArmKFactorY = " & Me.TorqueArmKFactorY.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyPullOffKFactorX = " & Me.GuyPullOffKFactorX.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyPullOffKFactorY = " & Me.GuyPullOffKFactorY.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyDiagKFactorX = " & Me.GuyDiagKFactorX.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyDiagKFactorY = " & Me.GuyDiagKFactorY.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyAutoCalc = " & Me.GuyAutoCalc.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyAllGuysSame = " & Me.GuyAllGuysSame.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyAllGuysAnchorSame = " & Me.GuyAllGuysAnchorSame.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyIsStrapping = " & Me.GuyIsStrapping.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyPullOffSize = " & Me.GuyPullOffSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyPullOffSizeBot = " & Me.GuyPullOffSizeBot.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyPullOffType = " & Me.GuyPullOffType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyPullOffGrade = " & Me.GuyPullOffGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyPullOffMatlGrade = " & Me.GuyPullOffMatlGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyUpperDiagSize = " & Me.GuyUpperDiagSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyLowerDiagSize = " & Me.GuyLowerDiagSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyDiagType = " & Me.GuyDiagType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyDiagGrade = " & Me.GuyDiagGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyDiagMatlGrade = " & Me.GuyDiagMatlGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyDiagNetWidthDeduct = " & Me.GuyDiagNetWidthDeduct.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyDiagUFactor = " & Me.GuyDiagUFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyDiagNumBolts = " & Me.GuyDiagNumBolts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyDiagonalOutOfPlaneRestraint = " & Me.GuyDiagonalOutOfPlaneRestraint.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyDiagBoltGrade = " & Me.GuyDiagBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyDiagBoltSize = " & Me.GuyDiagBoltSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyDiagBoltEdgeDistance = " & Me.GuyDiagBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyDiagBoltGageDistance = " & Me.GuyDiagBoltGageDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyPullOffNetWidthDeduct = " & Me.GuyPullOffNetWidthDeduct.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyPullOffUFactor = " & Me.GuyPullOffUFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyPullOffNumBolts = " & Me.GuyPullOffNumBolts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyPullOffOutOfPlaneRestraint = " & Me.GuyPullOffOutOfPlaneRestraint.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyPullOffBoltGrade = " & Me.GuyPullOffBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyPullOffBoltSize = " & Me.GuyPullOffBoltSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyPullOffBoltEdgeDistance = " & Me.GuyPullOffBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyPullOffBoltGageDistance = " & Me.GuyPullOffBoltGageDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyTorqueArmNetWidthDeduct = " & Me.GuyTorqueArmNetWidthDeduct.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyTorqueArmUFactor = " & Me.GuyTorqueArmUFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyTorqueArmNumBolts = " & Me.GuyTorqueArmNumBolts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyTorqueArmOutOfPlaneRestraint = " & Me.GuyTorqueArmOutOfPlaneRestraint.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyTorqueArmBoltGrade = " & Me.GuyTorqueArmBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyTorqueArmBoltSize = " & Me.GuyTorqueArmBoltSize.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyTorqueArmBoltEdgeDistance = " & Me.GuyTorqueArmBoltEdgeDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyTorqueArmBoltGageDistance = " & Me.GuyTorqueArmBoltGageDistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyPerCentTension = " & Me.GuyPerCentTension.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyPerCentTension120 = " & Me.GuyPerCentTension120.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyPerCentTension240 = " & Me.GuyPerCentTension240.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyPerCentTension360 = " & Me.GuyPerCentTension360.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyEffFactor = " & Me.GuyEffFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyEffFactor120 = " & Me.GuyEffFactor120.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyEffFactor240 = " & Me.GuyEffFactor240.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyEffFactor360 = " & Me.GuyEffFactor360.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyNumInsulators = " & Me.GuyNumInsulators.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyInsulatorLength = " & Me.GuyInsulatorLength.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyInsulatorDia = " & Me.GuyInsulatorDia.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyInsulatorWt = " & Me.GuyInsulatorWt.NullableToString.FormatDBValue)

        Return SQLUpdateFieldsandValues
    End Function
#End Region

#Region "Define"
    'Private _GuyRec As Integer?
    Private _GuyHeight As Double?
    Private _GuyAutoCalcKSingleAngle As Boolean?
    Private _GuyAutoCalcKSolidRound As Boolean?
    Private _GuyMount As String
    Private _TorqueArmStyle As String
    Private _GuyRadius As Double?
    Private _GuyRadius120 As Double?
    Private _GuyRadius240 As Double?
    Private _GuyRadius360 As Double?
    Private _TorqueArmRadius As Double?
    Private _TorqueArmLegAngle As Double?
    Private _Azimuth0Adjustment As Double?
    Private _Azimuth120Adjustment As Double?
    Private _Azimuth240Adjustment As Double?
    Private _Azimuth360Adjustment As Double?
    Private _Anchor0Elevation As Double?
    Private _Anchor120Elevation As Double?
    Private _Anchor240Elevation As Double?
    Private _Anchor360Elevation As Double?
    Private _GuySize As String
    Private _Guy120Size As String
    Private _Guy240Size As String
    Private _Guy360Size As String
    Private _GuyGrade As String
    Private _TorqueArmSize As String
    Private _TorqueArmSizeBot As String
    Private _TorqueArmType As String
    Private _TorqueArmGrade As Double?
    Private _TorqueArmMatlGrade As String
    Private _TorqueArmKFactor As Double?
    Private _TorqueArmKFactorY As Double?
    Private _GuyPullOffKFactorX As Double?
    Private _GuyPullOffKFactorY As Double?
    Private _GuyDiagKFactorX As Double?
    Private _GuyDiagKFactorY As Double?
    Private _GuyAutoCalc As Boolean?
    Private _GuyAllGuysSame As Boolean?
    Private _GuyAllGuysAnchorSame As Boolean?
    Private _GuyIsStrapping As Boolean?
    Private _GuyPullOffSize As String
    Private _GuyPullOffSizeBot As String
    Private _GuyPullOffType As String
    Private _GuyPullOffGrade As Double?
    Private _GuyPullOffMatlGrade As String
    Private _GuyUpperDiagSize As String
    Private _GuyLowerDiagSize As String
    Private _GuyDiagType As String
    Private _GuyDiagGrade As Double?
    Private _GuyDiagMatlGrade As String
    Private _GuyDiagNetWidthDeduct As Double?
    Private _GuyDiagUFactor As Double?
    Private _GuyDiagNumBolts As Integer?
    Private _GuyDiagonalOutOfPlaneRestraint As Boolean?
    Private _GuyDiagBoltGrade As String
    Private _GuyDiagBoltSize As Double?
    Private _GuyDiagBoltEdgeDistance As Double?
    Private _GuyDiagBoltGageDistance As Double?
    Private _GuyPullOffNetWidthDeduct As Double?
    Private _GuyPullOffUFactor As Double?
    Private _GuyPullOffNumBolts As Integer?
    Private _GuyPullOffOutOfPlaneRestraint As Boolean?
    Private _GuyPullOffBoltGrade As String
    Private _GuyPullOffBoltSize As Double?
    Private _GuyPullOffBoltEdgeDistance As Double?
    Private _GuyPullOffBoltGageDistance As Double?
    Private _GuyTorqueArmNetWidthDeduct As Double?
    Private _GuyTorqueArmUFactor As Double?
    Private _GuyTorqueArmNumBolts As Integer?
    Private _GuyTorqueArmOutOfPlaneRestraint As Boolean?
    Private _GuyTorqueArmBoltGrade As String
    Private _GuyTorqueArmBoltSize As Double?
    Private _GuyTorqueArmBoltEdgeDistance As Double?
    Private _GuyTorqueArmBoltGageDistance As Double?
    Private _GuyPerCentTension As Double?
    Private _GuyPerCentTension120 As Double?
    Private _GuyPerCentTension240 As Double?
    Private _GuyPerCentTension360 As Double?
    Private _GuyEffFactor As Double?
    Private _GuyEffFactor120 As Double?
    Private _GuyEffFactor240 As Double?
    Private _GuyEffFactor360 As Double?
    Private _GuyNumInsulators As Integer?
    Private _GuyInsulatorLength As Double?
    Private _GuyInsulatorDia As Double?
    Private _GuyInsulatorWt As Double?

    '<Category("TNX Guy Record"), Description(""), DisplayName("Guyrec")>
    ' <DataMember()> Public Property Rec() As Integer?
    '    Get
    '        Return Me._GuyRec
    '    End Get
    '    Set
    '        Me._GuyRec = Value
    '    End Set
    'End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyheight")>
    <DataMember()> Public Property GuyHeight() As Double?
        Get
            Return Me._GuyHeight
        End Get
        Set
            Me._GuyHeight = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyautocalcksingleangle")>
    <DataMember()> Public Property GuyAutoCalcKSingleAngle() As Boolean?
        Get
            Return Me._GuyAutoCalcKSingleAngle
        End Get
        Set
            Me._GuyAutoCalcKSingleAngle = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyautocalcksolidround")>
    <DataMember()> Public Property GuyAutoCalcKSolidRound() As Boolean?
        Get
            Return Me._GuyAutoCalcKSolidRound
        End Get
        Set
            Me._GuyAutoCalcKSolidRound = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guymount")>
    <DataMember()> Public Property GuyMount() As String
        Get
            Return Me._GuyMount
        End Get
        Set
            Me._GuyMount = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Torquearmstyle")>
    <DataMember()> Public Property TorqueArmStyle() As String
        Get
            Return Me._TorqueArmStyle
        End Get
        Set
            Me._TorqueArmStyle = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyradius")>
    <DataMember()> Public Property GuyRadius() As Double?
        Get
            Return Me._GuyRadius
        End Get
        Set
            Me._GuyRadius = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyradius120")>
    <DataMember()> Public Property GuyRadius120() As Double?
        Get
            Return Me._GuyRadius120
        End Get
        Set
            Me._GuyRadius120 = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyradius240")>
    <DataMember()> Public Property GuyRadius240() As Double?
        Get
            Return Me._GuyRadius240
        End Get
        Set
            Me._GuyRadius240 = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyradius360")>
    <DataMember()> Public Property GuyRadius360() As Double?
        Get
            Return Me._GuyRadius360
        End Get
        Set
            Me._GuyRadius360 = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Torquearmradius")>
    <DataMember()> Public Property TorqueArmRadius() As Double?
        Get
            Return Me._TorqueArmRadius
        End Get
        Set
            Me._TorqueArmRadius = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Torquearmlegangle")>
    <DataMember()> Public Property TorqueArmLegAngle() As Double?
        Get
            Return Me._TorqueArmLegAngle
        End Get
        Set
            Me._TorqueArmLegAngle = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Azimuth0Adjustment")>
    <DataMember()> Public Property Azimuth0Adjustment() As Double?
        Get
            Return Me._Azimuth0Adjustment
        End Get
        Set
            Me._Azimuth0Adjustment = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Azimuth120Adjustment")>
    <DataMember()> Public Property Azimuth120Adjustment() As Double?
        Get
            Return Me._Azimuth120Adjustment
        End Get
        Set
            Me._Azimuth120Adjustment = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Azimuth240Adjustment")>
    <DataMember()> Public Property Azimuth240Adjustment() As Double?
        Get
            Return Me._Azimuth240Adjustment
        End Get
        Set
            Me._Azimuth240Adjustment = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Azimuth360Adjustment")>
    <DataMember()> Public Property Azimuth360Adjustment() As Double?
        Get
            Return Me._Azimuth360Adjustment
        End Get
        Set
            Me._Azimuth360Adjustment = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Anchor0Elevation")>
    <DataMember()> Public Property Anchor0Elevation() As Double?
        Get
            Return Me._Anchor0Elevation
        End Get
        Set
            Me._Anchor0Elevation = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Anchor120Elevation")>
    <DataMember()> Public Property Anchor120Elevation() As Double?
        Get
            Return Me._Anchor120Elevation
        End Get
        Set
            Me._Anchor120Elevation = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Anchor240Elevation")>
    <DataMember()> Public Property Anchor240Elevation() As Double?
        Get
            Return Me._Anchor240Elevation
        End Get
        Set
            Me._Anchor240Elevation = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Anchor360Elevation")>
    <DataMember()> Public Property Anchor360Elevation() As Double?
        Get
            Return Me._Anchor360Elevation
        End Get
        Set
            Me._Anchor360Elevation = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guysize")>
    <DataMember()> Public Property GuySize() As String
        Get
            Return Me._GuySize
        End Get
        Set
            Me._GuySize = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guy120Size")>
    <DataMember()> Public Property Guy120Size() As String
        Get
            Return Me._Guy120Size
        End Get
        Set
            Me._Guy120Size = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guy240Size")>
    <DataMember()> Public Property Guy240Size() As String
        Get
            Return Me._Guy240Size
        End Get
        Set
            Me._Guy240Size = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guy360Size")>
    <DataMember()> Public Property Guy360Size() As String
        Get
            Return Me._Guy360Size
        End Get
        Set
            Me._Guy360Size = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guygrade")>
    <DataMember()> Public Property GuyGrade() As String
        Get
            Return Me._GuyGrade
        End Get
        Set
            Me._GuyGrade = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Torquearmsize")>
    <DataMember()> Public Property TorqueArmSize() As String
        Get
            Return Me._TorqueArmSize
        End Get
        Set
            Me._TorqueArmSize = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Torquearmsizebot")>
    <DataMember()> Public Property TorqueArmSizeBot() As String
        Get
            Return Me._TorqueArmSizeBot
        End Get
        Set
            Me._TorqueArmSizeBot = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Torquearmtype")>
    <DataMember()> Public Property TorqueArmType() As String
        Get
            Return Me._TorqueArmType
        End Get
        Set
            Me._TorqueArmType = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Torquearmgrade")>
    <DataMember()> Public Property TorqueArmGrade() As Double?
        Get
            Return Me._TorqueArmGrade
        End Get
        Set
            Me._TorqueArmGrade = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Torquearmmatlgrade")>
    <DataMember()> Public Property TorqueArmMatlGrade() As String
        Get
            Return Me._TorqueArmMatlGrade
        End Get
        Set
            Me._TorqueArmMatlGrade = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Torquearmkfactor")>
    <DataMember()> Public Property TorqueArmKFactor() As Double?
        Get
            Return Me._TorqueArmKFactor
        End Get
        Set
            Me._TorqueArmKFactor = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Torquearmkfactory")>
    <DataMember()> Public Property TorqueArmKFactorY() As Double?
        Get
            Return Me._TorqueArmKFactorY
        End Get
        Set
            Me._TorqueArmKFactorY = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypulloffkfactorx")>
    <DataMember()> Public Property GuyPullOffKFactorX() As Double?
        Get
            Return Me._GuyPullOffKFactorX
        End Get
        Set
            Me._GuyPullOffKFactorX = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypulloffkfactory")>
    <DataMember()> Public Property GuyPullOffKFactorY() As Double?
        Get
            Return Me._GuyPullOffKFactorY
        End Get
        Set
            Me._GuyPullOffKFactorY = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guydiagkfactorx")>
    <DataMember()> Public Property GuyDiagKFactorX() As Double?
        Get
            Return Me._GuyDiagKFactorX
        End Get
        Set
            Me._GuyDiagKFactorX = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guydiagkfactory")>
    <DataMember()> Public Property GuyDiagKFactorY() As Double?
        Get
            Return Me._GuyDiagKFactorY
        End Get
        Set
            Me._GuyDiagKFactorY = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyautocalc")>
    <DataMember()> Public Property GuyAutoCalc() As Boolean?
        Get
            Return Me._GuyAutoCalc
        End Get
        Set
            Me._GuyAutoCalc = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyallguyssame")>
    <DataMember()> Public Property GuyAllGuysSame() As Boolean?
        Get
            Return Me._GuyAllGuysSame
        End Get
        Set
            Me._GuyAllGuysSame = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyallguysanchorsame")>
    <DataMember()> Public Property GuyAllGuysAnchorSame() As Boolean?
        Get
            Return Me._GuyAllGuysAnchorSame
        End Get
        Set
            Me._GuyAllGuysAnchorSame = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyisstrapping")>
    <DataMember()> Public Property GuyIsStrapping() As Boolean?
        Get
            Return Me._GuyIsStrapping
        End Get
        Set
            Me._GuyIsStrapping = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypulloffsize")>
    <DataMember()> Public Property GuyPullOffSize() As String
        Get
            Return Me._GuyPullOffSize
        End Get
        Set
            Me._GuyPullOffSize = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypulloffsizebot")>
    <DataMember()> Public Property GuyPullOffSizeBot() As String
        Get
            Return Me._GuyPullOffSizeBot
        End Get
        Set
            Me._GuyPullOffSizeBot = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypullofftype")>
    <DataMember()> Public Property GuyPullOffType() As String
        Get
            Return Me._GuyPullOffType
        End Get
        Set
            Me._GuyPullOffType = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypulloffgrade")>
    <DataMember()> Public Property GuyPullOffGrade() As Double?
        Get
            Return Me._GuyPullOffGrade
        End Get
        Set
            Me._GuyPullOffGrade = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypulloffmatlgrade")>
    <DataMember()> Public Property GuyPullOffMatlGrade() As String
        Get
            Return Me._GuyPullOffMatlGrade
        End Get
        Set
            Me._GuyPullOffMatlGrade = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyupperdiagsize")>
    <DataMember()> Public Property GuyUpperDiagSize() As String
        Get
            Return Me._GuyUpperDiagSize
        End Get
        Set
            Me._GuyUpperDiagSize = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guylowerdiagsize")>
    <DataMember()> Public Property GuyLowerDiagSize() As String
        Get
            Return Me._GuyLowerDiagSize
        End Get
        Set
            Me._GuyLowerDiagSize = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guydiagtype")>
    <DataMember()> Public Property GuyDiagType() As String
        Get
            Return Me._GuyDiagType
        End Get
        Set
            Me._GuyDiagType = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guydiaggrade")>
    <DataMember()> Public Property GuyDiagGrade() As Double?
        Get
            Return Me._GuyDiagGrade
        End Get
        Set
            Me._GuyDiagGrade = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guydiagmatlgrade")>
    <DataMember()> Public Property GuyDiagMatlGrade() As String
        Get
            Return Me._GuyDiagMatlGrade
        End Get
        Set
            Me._GuyDiagMatlGrade = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guydiagnetwidthdeduct")>
    <DataMember()> Public Property GuyDiagNetWidthDeduct() As Double?
        Get
            Return Me._GuyDiagNetWidthDeduct
        End Get
        Set
            Me._GuyDiagNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guydiagufactor")>
    <DataMember()> Public Property GuyDiagUFactor() As Double?
        Get
            Return Me._GuyDiagUFactor
        End Get
        Set
            Me._GuyDiagUFactor = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guydiagnumbolts")>
    <DataMember()> Public Property GuyDiagNumBolts() As Integer?
        Get
            Return Me._GuyDiagNumBolts
        End Get
        Set
            Me._GuyDiagNumBolts = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guydiagonaloutofplanerestraint")>
    <DataMember()> Public Property GuyDiagonalOutOfPlaneRestraint() As Boolean?
        Get
            Return Me._GuyDiagonalOutOfPlaneRestraint
        End Get
        Set
            Me._GuyDiagonalOutOfPlaneRestraint = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guydiagboltgrade")>
    <DataMember()> Public Property GuyDiagBoltGrade() As String
        Get
            Return Me._GuyDiagBoltGrade
        End Get
        Set
            Me._GuyDiagBoltGrade = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guydiagboltsize")>
    <DataMember()> Public Property GuyDiagBoltSize() As Double?
        Get
            Return Me._GuyDiagBoltSize
        End Get
        Set
            Me._GuyDiagBoltSize = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guydiagboltedgedistance")>
    <DataMember()> Public Property GuyDiagBoltEdgeDistance() As Double?
        Get
            Return Me._GuyDiagBoltEdgeDistance
        End Get
        Set
            Me._GuyDiagBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guydiagboltgagedistance")>
    <DataMember()> Public Property GuyDiagBoltGageDistance() As Double?
        Get
            Return Me._GuyDiagBoltGageDistance
        End Get
        Set
            Me._GuyDiagBoltGageDistance = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypulloffnetwidthdeduct")>
    <DataMember()> Public Property GuyPullOffNetWidthDeduct() As Double?
        Get
            Return Me._GuyPullOffNetWidthDeduct
        End Get
        Set
            Me._GuyPullOffNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypulloffufactor")>
    <DataMember()> Public Property GuyPullOffUFactor() As Double?
        Get
            Return Me._GuyPullOffUFactor
        End Get
        Set
            Me._GuyPullOffUFactor = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypulloffnumbolts")>
    <DataMember()> Public Property GuyPullOffNumBolts() As Integer?
        Get
            Return Me._GuyPullOffNumBolts
        End Get
        Set
            Me._GuyPullOffNumBolts = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypulloffoutofplanerestraint")>
    <DataMember()> Public Property GuyPullOffOutOfPlaneRestraint() As Boolean?
        Get
            Return Me._GuyPullOffOutOfPlaneRestraint
        End Get
        Set
            Me._GuyPullOffOutOfPlaneRestraint = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypulloffboltgrade")>
    <DataMember()> Public Property GuyPullOffBoltGrade() As String
        Get
            Return Me._GuyPullOffBoltGrade
        End Get
        Set
            Me._GuyPullOffBoltGrade = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypulloffboltsize")>
    <DataMember()> Public Property GuyPullOffBoltSize() As Double?
        Get
            Return Me._GuyPullOffBoltSize
        End Get
        Set
            Me._GuyPullOffBoltSize = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypulloffboltedgedistance")>
    <DataMember()> Public Property GuyPullOffBoltEdgeDistance() As Double?
        Get
            Return Me._GuyPullOffBoltEdgeDistance
        End Get
        Set
            Me._GuyPullOffBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypulloffboltgagedistance")>
    <DataMember()> Public Property GuyPullOffBoltGageDistance() As Double?
        Get
            Return Me._GuyPullOffBoltGageDistance
        End Get
        Set
            Me._GuyPullOffBoltGageDistance = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guytorquearmnetwidthdeduct")>
    <DataMember()> Public Property GuyTorqueArmNetWidthDeduct() As Double?
        Get
            Return Me._GuyTorqueArmNetWidthDeduct
        End Get
        Set
            Me._GuyTorqueArmNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guytorquearmufactor")>
    <DataMember()> Public Property GuyTorqueArmUFactor() As Double?
        Get
            Return Me._GuyTorqueArmUFactor
        End Get
        Set
            Me._GuyTorqueArmUFactor = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guytorquearmnumbolts")>
    <DataMember()> Public Property GuyTorqueArmNumBolts() As Integer?
        Get
            Return Me._GuyTorqueArmNumBolts
        End Get
        Set
            Me._GuyTorqueArmNumBolts = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guytorquearmoutofplanerestraint")>
    <DataMember()> Public Property GuyTorqueArmOutOfPlaneRestraint() As Boolean?
        Get
            Return Me._GuyTorqueArmOutOfPlaneRestraint
        End Get
        Set
            Me._GuyTorqueArmOutOfPlaneRestraint = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guytorquearmboltgrade")>
    <DataMember()> Public Property GuyTorqueArmBoltGrade() As String
        Get
            Return Me._GuyTorqueArmBoltGrade
        End Get
        Set
            Me._GuyTorqueArmBoltGrade = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guytorquearmboltsize")>
    <DataMember()> Public Property GuyTorqueArmBoltSize() As Double?
        Get
            Return Me._GuyTorqueArmBoltSize
        End Get
        Set
            Me._GuyTorqueArmBoltSize = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guytorquearmboltedgedistance")>
    <DataMember()> Public Property GuyTorqueArmBoltEdgeDistance() As Double?
        Get
            Return Me._GuyTorqueArmBoltEdgeDistance
        End Get
        Set
            Me._GuyTorqueArmBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guytorquearmboltgagedistance")>
    <DataMember()> Public Property GuyTorqueArmBoltGageDistance() As Double?
        Get
            Return Me._GuyTorqueArmBoltGageDistance
        End Get
        Set
            Me._GuyTorqueArmBoltGageDistance = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypercenttension")>
    <DataMember()> Public Property GuyPerCentTension() As Double?
        Get
            Return Me._GuyPerCentTension
        End Get
        Set
            Me._GuyPerCentTension = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypercenttension120")>
    <DataMember()> Public Property GuyPerCentTension120() As Double?
        Get
            Return Me._GuyPerCentTension120
        End Get
        Set
            Me._GuyPerCentTension120 = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypercenttension240")>
    <DataMember()> Public Property GuyPerCentTension240() As Double?
        Get
            Return Me._GuyPerCentTension240
        End Get
        Set
            Me._GuyPerCentTension240 = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypercenttension360")>
    <DataMember()> Public Property GuyPerCentTension360() As Double?
        Get
            Return Me._GuyPerCentTension360
        End Get
        Set
            Me._GuyPerCentTension360 = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyefffactor")>
    <DataMember()> Public Property GuyEffFactor() As Double?
        Get
            Return Me._GuyEffFactor
        End Get
        Set
            Me._GuyEffFactor = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyefffactor120")>
    <DataMember()> Public Property GuyEffFactor120() As Double?
        Get
            Return Me._GuyEffFactor120
        End Get
        Set
            Me._GuyEffFactor120 = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyefffactor240")>
    <DataMember()> Public Property GuyEffFactor240() As Double?
        Get
            Return Me._GuyEffFactor240
        End Get
        Set
            Me._GuyEffFactor240 = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyefffactor360")>
    <DataMember()> Public Property GuyEffFactor360() As Double?
        Get
            Return Me._GuyEffFactor360
        End Get
        Set
            Me._GuyEffFactor360 = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guynuminsulators")>
    <DataMember()> Public Property GuyNumInsulators() As Integer?
        Get
            Return Me._GuyNumInsulators
        End Get
        Set
            Me._GuyNumInsulators = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyinsulatorlength")>
    <DataMember()> Public Property GuyInsulatorLength() As Double?
        Get
            Return Me._GuyInsulatorLength
        End Get
        Set
            Me._GuyInsulatorLength = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyinsulatordia")>
    <DataMember()> Public Property GuyInsulatorDia() As Double?
        Get
            Return Me._GuyInsulatorDia
        End Get
        Set
            Me._GuyInsulatorDia = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyinsulatorwt")>
    <DataMember()> Public Property GuyInsulatorWt() As Double?
        Get
            Return Me._GuyInsulatorWt
        End Get
        Set
            Me._GuyInsulatorWt = Value
        End Set
    End Property
#End Region

#Region "Constructors"
    Public Sub New(Optional ByVal Parent As EDSObject = Nothing)
        If Parent IsNot Nothing Then Me.Absorb(Parent)
    End Sub

    Public Sub New(data As DataRow, Optional ByVal Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Me.ID = DBtoNullableInt(data.Item("ID"))
        Me.Rec = DBtoNullableInt(data.Item("GuyRec"))
        Me.GuyHeight = DBtoNullableDbl(data.Item("GuyHeight"))
        Me.GuyAutoCalcKSingleAngle = DBtoNullableBool(data.Item("GuyAutoCalcKSingleAngle"))
        Me.GuyAutoCalcKSolidRound = DBtoNullableBool(data.Item("GuyAutoCalcKSolidRound"))
        Me.GuyMount = DBtoStr(data.Item("GuyMount"))
        Me.TorqueArmStyle = DBtoStr(data.Item("TorqueArmStyle"))
        Me.GuyRadius = DBtoNullableDbl(data.Item("GuyRadius"))
        Me.GuyRadius120 = DBtoNullableDbl(data.Item("GuyRadius120"))
        Me.GuyRadius240 = DBtoNullableDbl(data.Item("GuyRadius240"))
        Me.GuyRadius360 = DBtoNullableDbl(data.Item("GuyRadius360"))
        Me.TorqueArmRadius = DBtoNullableDbl(data.Item("TorqueArmRadius"))
        Me.TorqueArmLegAngle = DBtoNullableDbl(data.Item("TorqueArmLegAngle"))
        Me.Azimuth0Adjustment = DBtoNullableDbl(data.Item("Azimuth0Adjustment"))
        Me.Azimuth120Adjustment = DBtoNullableDbl(data.Item("Azimuth120Adjustment"))
        Me.Azimuth240Adjustment = DBtoNullableDbl(data.Item("Azimuth240Adjustment"))
        Me.Azimuth360Adjustment = DBtoNullableDbl(data.Item("Azimuth360Adjustment"))
        Me.Anchor0Elevation = DBtoNullableDbl(data.Item("Anchor0Elevation"))
        Me.Anchor120Elevation = DBtoNullableDbl(data.Item("Anchor120Elevation"))
        Me.Anchor240Elevation = DBtoNullableDbl(data.Item("Anchor240Elevation"))
        Me.Anchor360Elevation = DBtoNullableDbl(data.Item("Anchor360Elevation"))
        Me.GuySize = DBtoStr(data.Item("GuySize"))
        Me.Guy120Size = DBtoStr(data.Item("Guy120Size"))
        Me.Guy240Size = DBtoStr(data.Item("Guy240Size"))
        Me.Guy360Size = DBtoStr(data.Item("Guy360Size"))
        Me.GuyGrade = DBtoStr(data.Item("GuyGrade"))
        Me.TorqueArmSize = DBtoStr(data.Item("TorqueArmSize"))
        Me.TorqueArmSizeBot = DBtoStr(data.Item("TorqueArmSizeBot"))
        Me.TorqueArmType = DBtoStr(data.Item("TorqueArmType"))
        Me.TorqueArmGrade = DBtoNullableDbl(data.Item("TorqueArmGrade"))
        Me.TorqueArmMatlGrade = DBtoStr(data.Item("TorqueArmMatlGrade"))
        Me.TorqueArmKFactor = DBtoNullableDbl(data.Item("TorqueArmKFactor"))
        Me.TorqueArmKFactorY = DBtoNullableDbl(data.Item("TorqueArmKFactorY"))
        Me.GuyPullOffKFactorX = DBtoNullableDbl(data.Item("GuyPullOffKFactorX"))
        Me.GuyPullOffKFactorY = DBtoNullableDbl(data.Item("GuyPullOffKFactorY"))
        Me.GuyDiagKFactorX = DBtoNullableDbl(data.Item("GuyDiagKFactorX"))
        Me.GuyDiagKFactorY = DBtoNullableDbl(data.Item("GuyDiagKFactorY"))
        Me.GuyAutoCalc = DBtoNullableBool(data.Item("GuyAutoCalc"))
        Me.GuyAllGuysSame = DBtoNullableBool(data.Item("GuyAllGuysSame"))
        Me.GuyAllGuysAnchorSame = DBtoNullableBool(data.Item("GuyAllGuysAnchorSame"))
        Me.GuyIsStrapping = DBtoNullableBool(data.Item("GuyIsStrapping"))
        Me.GuyPullOffSize = DBtoStr(data.Item("GuyPullOffSize"))
        Me.GuyPullOffSizeBot = DBtoStr(data.Item("GuyPullOffSizeBot"))
        Me.GuyPullOffType = DBtoStr(data.Item("GuyPullOffType"))
        Me.GuyPullOffGrade = DBtoNullableDbl(data.Item("GuyPullOffGrade"))
        Me.GuyPullOffMatlGrade = DBtoStr(data.Item("GuyPullOffMatlGrade"))
        Me.GuyUpperDiagSize = DBtoStr(data.Item("GuyUpperDiagSize"))
        Me.GuyLowerDiagSize = DBtoStr(data.Item("GuyLowerDiagSize"))
        Me.GuyDiagType = DBtoStr(data.Item("GuyDiagType"))
        Me.GuyDiagGrade = DBtoNullableDbl(data.Item("GuyDiagGrade"))
        Me.GuyDiagMatlGrade = DBtoStr(data.Item("GuyDiagMatlGrade"))
        Me.GuyDiagNetWidthDeduct = DBtoNullableDbl(data.Item("GuyDiagNetWidthDeduct"))
        Me.GuyDiagUFactor = DBtoNullableDbl(data.Item("GuyDiagUFactor"))
        Me.GuyDiagNumBolts = DBtoNullableInt(data.Item("GuyDiagNumBolts"))
        Me.GuyDiagonalOutOfPlaneRestraint = DBtoNullableBool(data.Item("GuyDiagonalOutOfPlaneRestraint"))
        Me.GuyDiagBoltGrade = DBtoStr(data.Item("GuyDiagBoltGrade"))
        Me.GuyDiagBoltSize = DBtoNullableDbl(data.Item("GuyDiagBoltSize"))
        Me.GuyDiagBoltEdgeDistance = DBtoNullableDbl(data.Item("GuyDiagBoltEdgeDistance"))
        Me.GuyDiagBoltGageDistance = DBtoNullableDbl(data.Item("GuyDiagBoltGageDistance"))
        Me.GuyPullOffNetWidthDeduct = DBtoNullableDbl(data.Item("GuyPullOffNetWidthDeduct"))
        Me.GuyPullOffUFactor = DBtoNullableDbl(data.Item("GuyPullOffUFactor"))
        Me.GuyPullOffNumBolts = DBtoNullableInt(data.Item("GuyPullOffNumBolts"))
        Me.GuyPullOffOutOfPlaneRestraint = DBtoNullableBool(data.Item("GuyPullOffOutOfPlaneRestraint"))
        Me.GuyPullOffBoltGrade = DBtoStr(data.Item("GuyPullOffBoltGrade"))
        Me.GuyPullOffBoltSize = DBtoNullableDbl(data.Item("GuyPullOffBoltSize"))
        Me.GuyPullOffBoltEdgeDistance = DBtoNullableDbl(data.Item("GuyPullOffBoltEdgeDistance"))
        Me.GuyPullOffBoltGageDistance = DBtoNullableDbl(data.Item("GuyPullOffBoltGageDistance"))
        Me.GuyTorqueArmNetWidthDeduct = DBtoNullableDbl(data.Item("GuyTorqueArmNetWidthDeduct"))
        Me.GuyTorqueArmUFactor = DBtoNullableDbl(data.Item("GuyTorqueArmUFactor"))
        Me.GuyTorqueArmNumBolts = DBtoNullableInt(data.Item("GuyTorqueArmNumBolts"))
        Me.GuyTorqueArmOutOfPlaneRestraint = DBtoNullableBool(data.Item("GuyTorqueArmOutOfPlaneRestraint"))
        Me.GuyTorqueArmBoltGrade = DBtoStr(data.Item("GuyTorqueArmBoltGrade"))
        Me.GuyTorqueArmBoltSize = DBtoNullableDbl(data.Item("GuyTorqueArmBoltSize"))
        Me.GuyTorqueArmBoltEdgeDistance = DBtoNullableDbl(data.Item("GuyTorqueArmBoltEdgeDistance"))
        Me.GuyTorqueArmBoltGageDistance = DBtoNullableDbl(data.Item("GuyTorqueArmBoltGageDistance"))
        Me.GuyPerCentTension = DBtoNullableDbl(data.Item("GuyPerCentTension"))
        Me.GuyPerCentTension120 = DBtoNullableDbl(data.Item("GuyPerCentTension120"))
        Me.GuyPerCentTension240 = DBtoNullableDbl(data.Item("GuyPerCentTension240"))
        Me.GuyPerCentTension360 = DBtoNullableDbl(data.Item("GuyPerCentTension360"))
        Me.GuyEffFactor = DBtoNullableDbl(data.Item("GuyEffFactor"))
        Me.GuyEffFactor120 = DBtoNullableDbl(data.Item("GuyEffFactor120"))
        Me.GuyEffFactor240 = DBtoNullableDbl(data.Item("GuyEffFactor240"))
        Me.GuyEffFactor360 = DBtoNullableDbl(data.Item("GuyEffFactor360"))
        Me.GuyNumInsulators = DBtoNullableInt(data.Item("GuyNumInsulators"))
        Me.GuyInsulatorLength = DBtoNullableDbl(data.Item("GuyInsulatorLength"))
        Me.GuyInsulatorDia = DBtoNullableDbl(data.Item("GuyInsulatorDia"))
        Me.GuyInsulatorWt = DBtoNullableDbl(data.Item("GuyInsulatorWt"))

    End Sub
#End Region

#Region "Equals"
    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As tnxGuyRecord = TryCast(other, tnxGuyRecord)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.Rec.CheckChange(otherToCompare.Rec, changes, categoryName, "Guy Rec"), Equals, False)
        Equals = If(Me.GuyHeight.CheckChange(otherToCompare.GuyHeight, changes, categoryName, "Guy Height"), Equals, False)
        Equals = If(Me.GuyAutoCalcKSingleAngle.CheckChange(otherToCompare.GuyAutoCalcKSingleAngle, changes, categoryName, "Guy Auto Calc K Single Angle"), Equals, False)
        Equals = If(Me.GuyAutoCalcKSolidRound.CheckChange(otherToCompare.GuyAutoCalcKSolidRound, changes, categoryName, "Guy Auto Calc K Solid Round"), Equals, False)
        Equals = If(Me.GuyMount.CheckChange(otherToCompare.GuyMount, changes, categoryName, "Guy Mount"), Equals, False)
        Equals = If(Me.TorqueArmStyle.CheckChange(otherToCompare.TorqueArmStyle, changes, categoryName, "Torque Arm Style"), Equals, False)
        Equals = If(Me.GuyRadius.CheckChange(otherToCompare.GuyRadius, changes, categoryName, "Guy Radius"), Equals, False)
        Equals = If(Me.GuyRadius120.CheckChange(otherToCompare.GuyRadius120, changes, categoryName, "Guy Radius 120"), Equals, False)
        Equals = If(Me.GuyRadius240.CheckChange(otherToCompare.GuyRadius240, changes, categoryName, "Guy Radius 240"), Equals, False)
        Equals = If(Me.GuyRadius360.CheckChange(otherToCompare.GuyRadius360, changes, categoryName, "Guy Radius 360"), Equals, False)
        Equals = If(Me.TorqueArmRadius.CheckChange(otherToCompare.TorqueArmRadius, changes, categoryName, "Torque Arm Radius"), Equals, False)
        Equals = If(Me.TorqueArmLegAngle.CheckChange(otherToCompare.TorqueArmLegAngle, changes, categoryName, "Torque Arm Leg Angle"), Equals, False)
        Equals = If(Me.Azimuth0Adjustment.CheckChange(otherToCompare.Azimuth0Adjustment, changes, categoryName, "Azimuth 0 Adjustment"), Equals, False)
        Equals = If(Me.Azimuth120Adjustment.CheckChange(otherToCompare.Azimuth120Adjustment, changes, categoryName, "Azimuth 120 Adjustment"), Equals, False)
        Equals = If(Me.Azimuth240Adjustment.CheckChange(otherToCompare.Azimuth240Adjustment, changes, categoryName, "Azimuth 240 Adjustment"), Equals, False)
        Equals = If(Me.Azimuth360Adjustment.CheckChange(otherToCompare.Azimuth360Adjustment, changes, categoryName, "Azimuth 360 Adjustment"), Equals, False)
        Equals = If(Me.Anchor0Elevation.CheckChange(otherToCompare.Anchor0Elevation, changes, categoryName, "Anchor 0 Elevation"), Equals, False)
        Equals = If(Me.Anchor120Elevation.CheckChange(otherToCompare.Anchor120Elevation, changes, categoryName, "Anchor 120 Elevation"), Equals, False)
        Equals = If(Me.Anchor240Elevation.CheckChange(otherToCompare.Anchor240Elevation, changes, categoryName, "Anchor 240 Elevation"), Equals, False)
        Equals = If(Me.Anchor360Elevation.CheckChange(otherToCompare.Anchor360Elevation, changes, categoryName, "Anchor 360 Elevation"), Equals, False)
        Equals = If(Me.GuySize.CheckChange(otherToCompare.GuySize, changes, categoryName, "Guy Size"), Equals, False)
        Equals = If(Me.Guy120Size.CheckChange(otherToCompare.Guy120Size, changes, categoryName, "Guy 120 Size"), Equals, False)
        Equals = If(Me.Guy240Size.CheckChange(otherToCompare.Guy240Size, changes, categoryName, "Guy 240 Size"), Equals, False)
        Equals = If(Me.Guy360Size.CheckChange(otherToCompare.Guy360Size, changes, categoryName, "Guy 360 Size"), Equals, False)
        Equals = If(Me.GuyGrade.CheckChange(otherToCompare.GuyGrade, changes, categoryName, "Guy Grade"), Equals, False)
        Equals = If(Me.TorqueArmSize.CheckChange(otherToCompare.TorqueArmSize, changes, categoryName, "Torque Arm Size"), Equals, False)
        Equals = If(Me.TorqueArmSizeBot.CheckChange(otherToCompare.TorqueArmSizeBot, changes, categoryName, "Torque Arm Size Bot"), Equals, False)
        Equals = If(Me.TorqueArmType.CheckChange(otherToCompare.TorqueArmType, changes, categoryName, "Torque Arm Type"), Equals, False)
        Equals = If(Me.TorqueArmGrade.CheckChange(otherToCompare.TorqueArmGrade, changes, categoryName, "Torque Arm Grade"), Equals, False)
        Equals = If(Me.TorqueArmMatlGrade.CheckChange(otherToCompare.TorqueArmMatlGrade, changes, categoryName, "Torque Arm Matl Grade"), Equals, False)
        Equals = If(Me.TorqueArmKFactor.CheckChange(otherToCompare.TorqueArmKFactor, changes, categoryName, "Torque Arm K Factor"), Equals, False)
        Equals = If(Me.TorqueArmKFactorY.CheckChange(otherToCompare.TorqueArmKFactorY, changes, categoryName, "Torque Arm K Factor Y"), Equals, False)
        Equals = If(Me.GuyPullOffKFactorX.CheckChange(otherToCompare.GuyPullOffKFactorX, changes, categoryName, "Guy Pull Off K Factor X"), Equals, False)
        Equals = If(Me.GuyPullOffKFactorY.CheckChange(otherToCompare.GuyPullOffKFactorY, changes, categoryName, "Guy Pull Off K Factor Y"), Equals, False)
        Equals = If(Me.GuyDiagKFactorX.CheckChange(otherToCompare.GuyDiagKFactorX, changes, categoryName, "Guy Diag K Factor X"), Equals, False)
        Equals = If(Me.GuyDiagKFactorY.CheckChange(otherToCompare.GuyDiagKFactorY, changes, categoryName, "Guy Diag K Factor Y"), Equals, False)
        Equals = If(Me.GuyAutoCalc.CheckChange(otherToCompare.GuyAutoCalc, changes, categoryName, "Guy Auto Calc"), Equals, False)
        Equals = If(Me.GuyAllGuysSame.CheckChange(otherToCompare.GuyAllGuysSame, changes, categoryName, "Guy All Guys Same"), Equals, False)
        Equals = If(Me.GuyAllGuysAnchorSame.CheckChange(otherToCompare.GuyAllGuysAnchorSame, changes, categoryName, "Guy All Guys Anchor Same"), Equals, False)
        Equals = If(Me.GuyIsStrapping.CheckChange(otherToCompare.GuyIsStrapping, changes, categoryName, "Guy Is Strapping"), Equals, False)
        Equals = If(Me.GuyPullOffSize.CheckChange(otherToCompare.GuyPullOffSize, changes, categoryName, "Guy Pull Off Size"), Equals, False)
        Equals = If(Me.GuyPullOffSizeBot.CheckChange(otherToCompare.GuyPullOffSizeBot, changes, categoryName, "Guy Pull Off Size Bot"), Equals, False)
        Equals = If(Me.GuyPullOffType.CheckChange(otherToCompare.GuyPullOffType, changes, categoryName, "Guy Pull Off Type"), Equals, False)
        Equals = If(Me.GuyPullOffGrade.CheckChange(otherToCompare.GuyPullOffGrade, changes, categoryName, "Guy Pull Off Grade"), Equals, False)
        Equals = If(Me.GuyPullOffMatlGrade.CheckChange(otherToCompare.GuyPullOffMatlGrade, changes, categoryName, "Guy Pull Off Matl Grade"), Equals, False)
        Equals = If(Me.GuyUpperDiagSize.CheckChange(otherToCompare.GuyUpperDiagSize, changes, categoryName, "Guy Upper Diag Size"), Equals, False)
        Equals = If(Me.GuyLowerDiagSize.CheckChange(otherToCompare.GuyLowerDiagSize, changes, categoryName, "Guy Lower Diag Size"), Equals, False)
        Equals = If(Me.GuyDiagType.CheckChange(otherToCompare.GuyDiagType, changes, categoryName, "Guy Diag Type"), Equals, False)
        Equals = If(Me.GuyDiagGrade.CheckChange(otherToCompare.GuyDiagGrade, changes, categoryName, "Guy Diag Grade"), Equals, False)
        Equals = If(Me.GuyDiagMatlGrade.CheckChange(otherToCompare.GuyDiagMatlGrade, changes, categoryName, "Guy Diag Matl Grade"), Equals, False)
        Equals = If(Me.GuyDiagNetWidthDeduct.CheckChange(otherToCompare.GuyDiagNetWidthDeduct, changes, categoryName, "Guy Diag Net Width Deduct"), Equals, False)
        Equals = If(Me.GuyDiagUFactor.CheckChange(otherToCompare.GuyDiagUFactor, changes, categoryName, "Guy Diag U Factor"), Equals, False)
        Equals = If(Me.GuyDiagNumBolts.CheckChange(otherToCompare.GuyDiagNumBolts, changes, categoryName, "Guy Diag Num Bolts"), Equals, False)
        Equals = If(Me.GuyDiagonalOutOfPlaneRestraint.CheckChange(otherToCompare.GuyDiagonalOutOfPlaneRestraint, changes, categoryName, "Guy Diagonal Out Of Plane Restraint"), Equals, False)
        Equals = If(Me.GuyDiagBoltGrade.CheckChange(otherToCompare.GuyDiagBoltGrade, changes, categoryName, "Guy Diag Bolt Grade"), Equals, False)
        Equals = If(Me.GuyDiagBoltSize.CheckChange(otherToCompare.GuyDiagBoltSize, changes, categoryName, "Guy Diag Bolt Size"), Equals, False)
        Equals = If(Me.GuyDiagBoltEdgeDistance.CheckChange(otherToCompare.GuyDiagBoltEdgeDistance, changes, categoryName, "Guy Diag Bolt Edge Distance"), Equals, False)
        Equals = If(Me.GuyDiagBoltGageDistance.CheckChange(otherToCompare.GuyDiagBoltGageDistance, changes, categoryName, "Guy Diag Bolt Gage Distance"), Equals, False)
        Equals = If(Me.GuyPullOffNetWidthDeduct.CheckChange(otherToCompare.GuyPullOffNetWidthDeduct, changes, categoryName, "Guy Pull Off Net Width Deduct"), Equals, False)
        Equals = If(Me.GuyPullOffUFactor.CheckChange(otherToCompare.GuyPullOffUFactor, changes, categoryName, "Guy Pull Off U Factor"), Equals, False)
        Equals = If(Me.GuyPullOffNumBolts.CheckChange(otherToCompare.GuyPullOffNumBolts, changes, categoryName, "Guy Pull Off Num Bolts"), Equals, False)
        Equals = If(Me.GuyPullOffOutOfPlaneRestraint.CheckChange(otherToCompare.GuyPullOffOutOfPlaneRestraint, changes, categoryName, "Guy Pull Off Out Of Plane Restraint"), Equals, False)
        Equals = If(Me.GuyPullOffBoltGrade.CheckChange(otherToCompare.GuyPullOffBoltGrade, changes, categoryName, "Guy Pull Off Bolt Grade"), Equals, False)
        Equals = If(Me.GuyPullOffBoltSize.CheckChange(otherToCompare.GuyPullOffBoltSize, changes, categoryName, "Guy Pull Off Bolt Size"), Equals, False)
        Equals = If(Me.GuyPullOffBoltEdgeDistance.CheckChange(otherToCompare.GuyPullOffBoltEdgeDistance, changes, categoryName, "Guy Pull Off Bolt Edge Distance"), Equals, False)
        Equals = If(Me.GuyPullOffBoltGageDistance.CheckChange(otherToCompare.GuyPullOffBoltGageDistance, changes, categoryName, "Guy Pull Off Bolt Gage Distance"), Equals, False)
        Equals = If(Me.GuyTorqueArmNetWidthDeduct.CheckChange(otherToCompare.GuyTorqueArmNetWidthDeduct, changes, categoryName, "Guy Torque Arm Net Width Deduct"), Equals, False)
        Equals = If(Me.GuyTorqueArmUFactor.CheckChange(otherToCompare.GuyTorqueArmUFactor, changes, categoryName, "Guy Torque Arm U Factor"), Equals, False)
        Equals = If(Me.GuyTorqueArmNumBolts.CheckChange(otherToCompare.GuyTorqueArmNumBolts, changes, categoryName, "Guy Torque Arm Num Bolts"), Equals, False)
        Equals = If(Me.GuyTorqueArmOutOfPlaneRestraint.CheckChange(otherToCompare.GuyTorqueArmOutOfPlaneRestraint, changes, categoryName, "Guy Torque Arm Out Of Plane Restraint"), Equals, False)
        Equals = If(Me.GuyTorqueArmBoltGrade.CheckChange(otherToCompare.GuyTorqueArmBoltGrade, changes, categoryName, "Guy Torque Arm Bolt Grade"), Equals, False)
        Equals = If(Me.GuyTorqueArmBoltSize.CheckChange(otherToCompare.GuyTorqueArmBoltSize, changes, categoryName, "Guy Torque Arm Bolt Size"), Equals, False)
        Equals = If(Me.GuyTorqueArmBoltEdgeDistance.CheckChange(otherToCompare.GuyTorqueArmBoltEdgeDistance, changes, categoryName, "Guy Torque Arm Bolt Edge Distance"), Equals, False)
        Equals = If(Me.GuyTorqueArmBoltGageDistance.CheckChange(otherToCompare.GuyTorqueArmBoltGageDistance, changes, categoryName, "Guy Torque Arm Bolt Gage Distance"), Equals, False)
        Equals = If(Me.GuyPerCentTension.CheckChange(otherToCompare.GuyPerCentTension, changes, categoryName, "Guy Per Cent Tension"), Equals, False)
        Equals = If(Me.GuyPerCentTension120.CheckChange(otherToCompare.GuyPerCentTension120, changes, categoryName, "Guy Per Cent Tension 120"), Equals, False)
        Equals = If(Me.GuyPerCentTension240.CheckChange(otherToCompare.GuyPerCentTension240, changes, categoryName, "Guy Per Cent Tension 240"), Equals, False)
        Equals = If(Me.GuyPerCentTension360.CheckChange(otherToCompare.GuyPerCentTension360, changes, categoryName, "Guy Per Cent Tension 360"), Equals, False)
        Equals = If(Me.GuyEffFactor.CheckChange(otherToCompare.GuyEffFactor, changes, categoryName, "Guy Eff Factor"), Equals, False)
        Equals = If(Me.GuyEffFactor120.CheckChange(otherToCompare.GuyEffFactor120, changes, categoryName, "Guy Eff Factor 120"), Equals, False)
        Equals = If(Me.GuyEffFactor240.CheckChange(otherToCompare.GuyEffFactor240, changes, categoryName, "Guy Eff Factor 240"), Equals, False)
        Equals = If(Me.GuyEffFactor360.CheckChange(otherToCompare.GuyEffFactor360, changes, categoryName, "Guy Eff Factor 360"), Equals, False)
        Equals = If(Me.GuyNumInsulators.CheckChange(otherToCompare.GuyNumInsulators, changes, categoryName, "Guy Num Insulators"), Equals, False)
        Equals = If(Me.GuyInsulatorLength.CheckChange(otherToCompare.GuyInsulatorLength, changes, categoryName, "Guy Insulator Length"), Equals, False)
        Equals = If(Me.GuyInsulatorDia.CheckChange(otherToCompare.GuyInsulatorDia, changes, categoryName, "Guy Insulator Dia"), Equals, False)
        Equals = If(Me.GuyInsulatorWt.CheckChange(otherToCompare.GuyInsulatorWt, changes, categoryName, "Guy Insulator Wt"), Equals, False)

        Return Equals
    End Function
#End Region
End Class
#End Region