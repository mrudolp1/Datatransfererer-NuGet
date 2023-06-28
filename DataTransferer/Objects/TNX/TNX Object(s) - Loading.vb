
Option Compare Binary 'Trying to speed up parsing the TNX file by using Binary Text comparison instead of Text Comparison

Imports System.ComponentModel
Imports System.Runtime.Serialization

<DataContract()>
Partial Public Class tnxFeedLine
    Inherits EDSObjectWithQueries

#Region "Inheritted"

    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Feed Line"
        End Get
    End Property

    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "load.linear_output"
        End Get
    End Property
#End Region

#Region "Define"
    Public Property parent_id As Integer?

    Private _tnx_id As Integer?
    Private _FeedLineRec As Integer?
    Private _FeedLineEnabled As Boolean?
    Private _FeedLineDatabase As String
    Private _FeedLineDescription As String
    Private _FeedLineClassificationCategory As String
    Private _FeedLineNote As String
    Private _FeedLineNum As Integer?
    Private _FeedLineUseShielding As Boolean?
    Private _ExcludeFeedLineFromTorque As Boolean?
    Private _FeedLineNumPerRow As Integer?
    Private _FeedLineFace As Integer?
    Private _FeedLineComponentType As String
    Private _FeedLineGroupTreatmentType As String
    Private _FeedLineRoundClusterDia As Decimal?
    Private _FeedLineWidth As Decimal?
    Private _FeedLinePerimeter As Decimal?
    Private _FlatAttachmentEffectiveWidthRatio As Decimal?
    Private _AutoCalcFlatAttachmentEffectiveWidthRatio As Boolean?
    Private _FeedLineShieldingFactorKaNoIce As Decimal?
    Private _FeedLineShieldingFactorKaIce As Decimal?
    Private _FeedLineAutoCalcKa As Boolean?
    Private _FeedLineCaAaNoIce As Decimal?
    Private _FeedLineCaAaIce As Decimal?
    Private _FeedLineCaAaIce_1 As Decimal?
    Private _FeedLineCaAaIce_2 As Decimal?
    Private _FeedLineCaAaIce_4 As Decimal?
    Private _FeedLineWtNoIce As Decimal?
    Private _FeedLineWtIce As Decimal?
    Private _FeedLineWtIce_1 As Decimal?
    Private _FeedLineWtIce_2 As Decimal?
    Private _FeedLineWtIce_4 As Decimal?
    Private _FeedLineFaceOffset As Decimal?
    Private _FeedLineOffsetFrac As Decimal?
    Private _FeedLinePerimeterOffsetStartFrac As Decimal?
    Private _FeedLinePerimeterOffsetEndFrac As Decimal?
    Private _FeedLineStartHt As Decimal?
    Private _FeedLineEndHt As Decimal?
    Private _FeedLineClearSpacing As Decimal?
    Private _FeedLineRowClearSpacing As Decimal?

    <Category("TNX Feed Lines"), Description(""), DisplayName("tnx ID")>
    <DataMember()> Public Property tnx_id() As Integer?
        Get
            Return Me._tnx_id
        End Get
        Set
            Me._tnx_id = Value
        End Set
    End Property

    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineRec")>
    <DataMember()> Public Property FeedLineRec() As Integer?
        Get
            Return Me._FeedLineRec
        End Get
        Set
            Me._FeedLineRec = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineEnabled")>
    <DataMember()> Public Property FeedLineEnabled() As Boolean?
        Get
            Return Me._FeedLineEnabled
        End Get
        Set
            Me._FeedLineEnabled = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineDatabase")>
    <DataMember()> Public Property FeedLineDatabase() As String
        Get
            Return Me._FeedLineDatabase
        End Get
        Set
            Me._FeedLineDatabase = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineDescription")>
    <DataMember()> Public Property FeedLineDescription() As String
        Get
            Return Me._FeedLineDescription
        End Get
        Set
            Me._FeedLineDescription = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineClassificationCategory")>
    <DataMember()> Public Property FeedLineClassificationCategory() As String
        Get
            Return Me._FeedLineClassificationCategory
        End Get
        Set
            Me._FeedLineClassificationCategory = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineNote")>
    <DataMember()> Public Property FeedLineNote() As String
        Get
            Return Me._FeedLineNote
        End Get
        Set
            Me._FeedLineNote = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineNum")>
    <DataMember()> Public Property FeedLineNum() As Integer?
        Get
            Return Me._FeedLineNum
        End Get
        Set
            Me._FeedLineNum = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineUseShielding")>
    <DataMember()> Public Property FeedLineUseShielding() As Boolean?
        Get
            Return Me._FeedLineUseShielding
        End Get
        Set
            Me._FeedLineUseShielding = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("ExcludeFeedLineFromTorque")>
    <DataMember()> Public Property ExcludeFeedLineFromTorque() As Boolean?
        Get
            Return Me._ExcludeFeedLineFromTorque
        End Get
        Set
            Me._ExcludeFeedLineFromTorque = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineNumPerRow")>
    <DataMember()> Public Property FeedLineNumPerRow() As Integer?
        Get
            Return Me._FeedLineNumPerRow
        End Get
        Set
            Me._FeedLineNumPerRow = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description("{0 = A, 1 = B,  2 = C, 3 = D}"), DisplayName("FeedLineFace")>
    <DataMember()> Public Property FeedLineFace() As Integer?
        Get
            Return Me._FeedLineFace
        End Get
        Set
            Me._FeedLineFace = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineComponentType")>
    <DataMember()> Public Property FeedLineComponentType() As String
        Get
            Return Me._FeedLineComponentType
        End Get
        Set
            Me._FeedLineComponentType = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineGroupTreatmentType")>
    <DataMember()> Public Property FeedLineGroupTreatmentType() As String
        Get
            Return Me._FeedLineGroupTreatmentType
        End Get
        Set
            Me._FeedLineGroupTreatmentType = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineRoundClusterDia")>
    <DataMember()> Public Property FeedLineRoundClusterDia() As Decimal?
        Get
            Return Me._FeedLineRoundClusterDia
        End Get
        Set
            Me._FeedLineRoundClusterDia = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineWidth")>
    <DataMember()> Public Property FeedLineWidth() As Decimal?
        Get
            Return Me._FeedLineWidth
        End Get
        Set
            Me._FeedLineWidth = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLinePerimeter")>
    <DataMember()> Public Property FeedLinePerimeter() As Decimal?
        Get
            Return Me._FeedLinePerimeter
        End Get
        Set
            Me._FeedLinePerimeter = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FlatAttachmentEffectiveWidthRatio")>
    <DataMember()> Public Property FlatAttachmentEffectiveWidthRatio() As Decimal?
        Get
            Return Me._FlatAttachmentEffectiveWidthRatio
        End Get
        Set
            Me._FlatAttachmentEffectiveWidthRatio = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("AutoCalcFlatAttachmentEffectiveWidthRatio")>
    <DataMember()> Public Property AutoCalcFlatAttachmentEffectiveWidthRatio() As Boolean?
        Get
            Return Me._AutoCalcFlatAttachmentEffectiveWidthRatio
        End Get
        Set
            Me._AutoCalcFlatAttachmentEffectiveWidthRatio = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineShieldingFactorKaNoIce")>
    <DataMember()> Public Property FeedLineShieldingFactorKaNoIce() As Decimal?
        Get
            Return Me._FeedLineShieldingFactorKaNoIce
        End Get
        Set
            Me._FeedLineShieldingFactorKaNoIce = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineShieldingFactorKaIce")>
    <DataMember()> Public Property FeedLineShieldingFactorKaIce() As Decimal?
        Get
            Return Me._FeedLineShieldingFactorKaIce
        End Get
        Set
            Me._FeedLineShieldingFactorKaIce = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineAutoCalcKa")>
    <DataMember()> Public Property FeedLineAutoCalcKa() As Boolean?
        Get
            Return Me._FeedLineAutoCalcKa
        End Get
        Set
            Me._FeedLineAutoCalcKa = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineCaAaNoIce")>
    <DataMember()> Public Property FeedLineCaAaNoIce() As Decimal?
        Get
            Return Me._FeedLineCaAaNoIce
        End Get
        Set
            Me._FeedLineCaAaNoIce = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineCaAaIce")>
    <DataMember()> Public Property FeedLineCaAaIce() As Decimal?
        Get
            Return Me._FeedLineCaAaIce
        End Get
        Set
            Me._FeedLineCaAaIce = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineCaAaIce_1")>
    <DataMember()> Public Property FeedLineCaAaIce_1() As Decimal?
        Get
            Return Me._FeedLineCaAaIce_1
        End Get
        Set
            Me._FeedLineCaAaIce_1 = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineCaAaIce_2")>
    <DataMember()> Public Property FeedLineCaAaIce_2() As Decimal?
        Get
            Return Me._FeedLineCaAaIce_2
        End Get
        Set
            Me._FeedLineCaAaIce_2 = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineCaAaIce_4")>
    <DataMember()> Public Property FeedLineCaAaIce_4() As Decimal?
        Get
            Return Me._FeedLineCaAaIce_4
        End Get
        Set
            Me._FeedLineCaAaIce_4 = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineWtNoIce")>
    <DataMember()> Public Property FeedLineWtNoIce() As Decimal?
        Get
            Return Me._FeedLineWtNoIce
        End Get
        Set
            Me._FeedLineWtNoIce = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineWtIce")>
    <DataMember()> Public Property FeedLineWtIce() As Decimal?
        Get
            Return Me._FeedLineWtIce
        End Get
        Set
            Me._FeedLineWtIce = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineWtIce_1")>
    <DataMember()> Public Property FeedLineWtIce_1() As Decimal?
        Get
            Return Me._FeedLineWtIce_1
        End Get
        Set
            Me._FeedLineWtIce_1 = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineWtIce_2")>
    <DataMember()> Public Property FeedLineWtIce_2() As Decimal?
        Get
            Return Me._FeedLineWtIce_2
        End Get
        Set
            Me._FeedLineWtIce_2 = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineWtIce_4")>
    <DataMember()> Public Property FeedLineWtIce_4() As Decimal?
        Get
            Return Me._FeedLineWtIce_4
        End Get
        Set
            Me._FeedLineWtIce_4 = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineFaceOffset")>
    <DataMember()> Public Property FeedLineFaceOffset() As Decimal?
        Get
            Return Me._FeedLineFaceOffset
        End Get
        Set
            Me._FeedLineFaceOffset = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineOffsetFrac")>
    <DataMember()> Public Property FeedLineOffsetFrac() As Decimal?
        Get
            Return Me._FeedLineOffsetFrac
        End Get
        Set
            Me._FeedLineOffsetFrac = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLinePerimeterOffsetStartFrac")>
    <DataMember()> Public Property FeedLinePerimeterOffsetStartFrac() As Decimal?
        Get
            Return Me._FeedLinePerimeterOffsetStartFrac
        End Get
        Set
            Me._FeedLinePerimeterOffsetStartFrac = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLinePerimeterOffsetEndFrac")>
    <DataMember()> Public Property FeedLinePerimeterOffsetEndFrac() As Decimal?
        Get
            Return Me._FeedLinePerimeterOffsetEndFrac
        End Get
        Set
            Me._FeedLinePerimeterOffsetEndFrac = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineStartHt")>
    <DataMember()> Public Property FeedLineStartHt() As Decimal?
        Get
            Return Me._FeedLineStartHt
        End Get
        Set
            Me._FeedLineStartHt = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineEndHt")>
    <DataMember()> Public Property FeedLineEndHt() As Decimal?
        Get
            Return Me._FeedLineEndHt
        End Get
        Set
            Me._FeedLineEndHt = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineClearSpacing")>
    <DataMember()> Public Property FeedLineClearSpacing() As Decimal?
        Get
            Return Me._FeedLineClearSpacing
        End Get
        Set
            Me._FeedLineClearSpacing = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineRowClearSpacing")>
    <DataMember()> Public Property FeedLineRowClearSpacing() As Decimal?
        Get
            Return Me._FeedLineRowClearSpacing
        End Get
        Set
            Me._FeedLineRowClearSpacing = Value
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
        Dim otherToCompare As tnxFeedLine = TryCast(other, tnxFeedLine)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.FeedLineRec.CheckChange(otherToCompare.FeedLineRec, changes, categoryName, "Feedlinerec"), Equals, False)
        Equals = If(Me.FeedLineEnabled.CheckChange(otherToCompare.FeedLineEnabled, changes, categoryName, "Feedlineenabled"), Equals, False)
        Equals = If(Me.FeedLineDatabase.CheckChange(otherToCompare.FeedLineDatabase, changes, categoryName, "Feedlinedatabase"), Equals, False)
        Equals = If(Me.FeedLineDescription.CheckChange(otherToCompare.FeedLineDescription, changes, categoryName, "Feedlinedescription"), Equals, False)
        Equals = If(Me.FeedLineClassificationCategory.CheckChange(otherToCompare.FeedLineClassificationCategory, changes, categoryName, "Feedlineclassificationcategory"), Equals, False)
        Equals = If(Me.FeedLineNote.CheckChange(otherToCompare.FeedLineNote, changes, categoryName, "Feedlinenote"), Equals, False)
        Equals = If(Me.FeedLineNum.CheckChange(otherToCompare.FeedLineNum, changes, categoryName, "Feedlinenum"), Equals, False)
        Equals = If(Me.FeedLineUseShielding.CheckChange(otherToCompare.FeedLineUseShielding, changes, categoryName, "Feedlineuseshielding"), Equals, False)
        Equals = If(Me.ExcludeFeedLineFromTorque.CheckChange(otherToCompare.ExcludeFeedLineFromTorque, changes, categoryName, "Excludefeedlinefromtorque"), Equals, False)
        Equals = If(Me.FeedLineNumPerRow.CheckChange(otherToCompare.FeedLineNumPerRow, changes, categoryName, "Feedlinenumperrow"), Equals, False)
        Equals = If(Me.FeedLineFace.CheckChange(otherToCompare.FeedLineFace, changes, categoryName, "Feedlineface"), Equals, False)
        Equals = If(Me.FeedLineComponentType.CheckChange(otherToCompare.FeedLineComponentType, changes, categoryName, "Feedlinecomponenttype"), Equals, False)
        Equals = If(Me.FeedLineGroupTreatmentType.CheckChange(otherToCompare.FeedLineGroupTreatmentType, changes, categoryName, "Feedlinegrouptreatmenttype"), Equals, False)
        Equals = If(Me.FeedLineRoundClusterDia.CheckChange(otherToCompare.FeedLineRoundClusterDia, changes, categoryName, "Feedlineroundclusterdia"), Equals, False)
        Equals = If(Me.FeedLineWidth.CheckChange(otherToCompare.FeedLineWidth, changes, categoryName, "Feedlinewidth"), Equals, False)
        Equals = If(Me.FeedLinePerimeter.CheckChange(otherToCompare.FeedLinePerimeter, changes, categoryName, "Feedlineperimeter"), Equals, False)
        Equals = If(Me.FlatAttachmentEffectiveWidthRatio.CheckChange(otherToCompare.FlatAttachmentEffectiveWidthRatio, changes, categoryName, "Flatattachmenteffectivewidthratio"), Equals, False)
        Equals = If(Me.AutoCalcFlatAttachmentEffectiveWidthRatio.CheckChange(otherToCompare.AutoCalcFlatAttachmentEffectiveWidthRatio, changes, categoryName, "Autocalcflatattachmenteffectivewidthratio"), Equals, False)
        Equals = If(Me.FeedLineShieldingFactorKaNoIce.CheckChange(otherToCompare.FeedLineShieldingFactorKaNoIce, changes, categoryName, "Feedlineshieldingfactorkanoice"), Equals, False)
        Equals = If(Me.FeedLineShieldingFactorKaIce.CheckChange(otherToCompare.FeedLineShieldingFactorKaIce, changes, categoryName, "Feedlineshieldingfactorkaice"), Equals, False)
        Equals = If(Me.FeedLineAutoCalcKa.CheckChange(otherToCompare.FeedLineAutoCalcKa, changes, categoryName, "Feedlineautocalcka"), Equals, False)
        Equals = If(Me.FeedLineCaAaNoIce.CheckChange(otherToCompare.FeedLineCaAaNoIce, changes, categoryName, "Feedlinecaaanoice"), Equals, False)
        Equals = If(Me.FeedLineCaAaIce.CheckChange(otherToCompare.FeedLineCaAaIce, changes, categoryName, "Feedlinecaaaice"), Equals, False)
        Equals = If(Me.FeedLineCaAaIce_1.CheckChange(otherToCompare.FeedLineCaAaIce_1, changes, categoryName, "Feedlinecaaaice 1"), Equals, False)
        Equals = If(Me.FeedLineCaAaIce_2.CheckChange(otherToCompare.FeedLineCaAaIce_2, changes, categoryName, "Feedlinecaaaice 2"), Equals, False)
        Equals = If(Me.FeedLineCaAaIce_4.CheckChange(otherToCompare.FeedLineCaAaIce_4, changes, categoryName, "Feedlinecaaaice 4"), Equals, False)
        Equals = If(Me.FeedLineWtNoIce.CheckChange(otherToCompare.FeedLineWtNoIce, changes, categoryName, "Feedlinewtnoice"), Equals, False)
        Equals = If(Me.FeedLineWtIce.CheckChange(otherToCompare.FeedLineWtIce, changes, categoryName, "Feedlinewtice"), Equals, False)
        Equals = If(Me.FeedLineWtIce_1.CheckChange(otherToCompare.FeedLineWtIce_1, changes, categoryName, "Feedlinewtice 1"), Equals, False)
        Equals = If(Me.FeedLineWtIce_2.CheckChange(otherToCompare.FeedLineWtIce_2, changes, categoryName, "Feedlinewtice 2"), Equals, False)
        Equals = If(Me.FeedLineWtIce_4.CheckChange(otherToCompare.FeedLineWtIce_4, changes, categoryName, "Feedlinewtice 4"), Equals, False)
        Equals = If(Me.FeedLineFaceOffset.CheckChange(otherToCompare.FeedLineFaceOffset, changes, categoryName, "Feedlinefaceoffset"), Equals, False)
        Equals = If(Me.FeedLineOffsetFrac.CheckChange(otherToCompare.FeedLineOffsetFrac, changes, categoryName, "Feedlineoffsetfrac"), Equals, False)
        Equals = If(Me.FeedLinePerimeterOffsetStartFrac.CheckChange(otherToCompare.FeedLinePerimeterOffsetStartFrac, changes, categoryName, "Feedlineperimeteroffsetstartfrac"), Equals, False)
        Equals = If(Me.FeedLinePerimeterOffsetEndFrac.CheckChange(otherToCompare.FeedLinePerimeterOffsetEndFrac, changes, categoryName, "Feedlineperimeteroffsetendfrac"), Equals, False)
        Equals = If(Me.FeedLineStartHt.CheckChange(otherToCompare.FeedLineStartHt, changes, categoryName, "Feedlinestartht"), Equals, False)
        Equals = If(Me.FeedLineEndHt.CheckChange(otherToCompare.FeedLineEndHt, changes, categoryName, "Feedlineendht"), Equals, False)
        Equals = If(Me.FeedLineClearSpacing.CheckChange(otherToCompare.FeedLineClearSpacing, changes, categoryName, "Feedlineclearspacing"), Equals, False)
        Equals = If(Me.FeedLineRowClearSpacing.CheckChange(otherToCompare.FeedLineRowClearSpacing, changes, categoryName, "Feedlinerowclearspacing"), Equals, False)

        Return Equals
    End Function

#End Region

    Public Sub New(ByVal recRow As DataRow, Optional ByVal Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Me.ID = DBtoNullableInt(recRow.Item("ID"))
        'Me.bus_unit = DBtoNullableInt(recRow.Item("bus_unit"))
        'Me.structure_id = DBtoStr(recRow.Item("structure_id"))
        'Me.modified_person_id = DBtoNullableInt(recRow.Item("modified_person_id"))
        'Me.process_stage = DBtoStr(recRow.Item("process_stage"))
        'Me.work_order_seq_num = DBtoNullableInt(recRow.Item("work_order_seq_num"))
        Me.tnx_id = DBtoNullableInt(recRow.Item("tnx_id"))
        Me.FeedLineRec = DBtoNullableInt(recRow.Item("FeedLineRec"))
        Me.FeedLineEnabled = DBtoNullableBool(recRow.Item("FeedLineEnabled"))
        Me.FeedLineDatabase = DBtoStr(recRow.Item("FeedLineDatabase"))
        Me.FeedLineDescription = DBtoStr(recRow.Item("FeedLineDescription"))
        Me.FeedLineClassificationCategory = DBtoStr(recRow.Item("FeedLineClassificationCategory"))
        Me.FeedLineNote = DBtoStr(recRow.Item("FeedLineNote"))
        Me.FeedLineNum = DBtoNullableInt(recRow.Item("FeedLineNum"))
        Me.FeedLineUseShielding = DBtoNullableBool(recRow.Item("FeedLineUseShielding"))
        Me.ExcludeFeedLineFromTorque = DBtoNullableBool(recRow.Item("ExcludeFeedLineFromTorque"))
        Me.FeedLineNumPerRow = DBtoNullableInt(recRow.Item("FeedLineNumPerRow"))
        Me.FeedLineFace = DBtoNullableInt(recRow.Item("FeedLineFace"))
        Me.FeedLineComponentType = DBtoStr(recRow.Item("FeedLineComponentType"))
        Me.FeedLineGroupTreatmentType = DBtoStr(recRow.Item("FeedLineGroupTreatmentType"))
        Me.FeedLineRoundClusterDia = DBtoNullableDec(recRow.Item("FeedLineRoundClusterDia"), 6)
        Me.FeedLineWidth = DBtoNullableDec(recRow.Item("FeedLineWidth"), 6)
        Me.FeedLinePerimeter = DBtoNullableDec(recRow.Item("FeedLinePerimeter"), 6)
        Me.FlatAttachmentEffectiveWidthRatio = DBtoNullableDec(recRow.Item("FlatAttachmentEffectiveWidthRatio"), 6)
        Me.AutoCalcFlatAttachmentEffectiveWidthRatio = DBtoNullableBool(recRow.Item("AutoCalcFlatAttachmentEffectiveWidthRatio"))
        Me.FeedLineShieldingFactorKaNoIce = DBtoNullableDec(recRow.Item("FeedLineShieldingFactorKaNoIce"), 6)
        Me.FeedLineShieldingFactorKaIce = DBtoNullableDec(recRow.Item("FeedLineShieldingFactorKaIce"), 6)
        Me.FeedLineAutoCalcKa = DBtoNullableBool(recRow.Item("FeedLineAutoCalcKa"))
        Me.FeedLineCaAaNoIce = DBtoNullableDec(recRow.Item("FeedLineCaAaNoIce"), 6)
        Me.FeedLineCaAaIce = DBtoNullableDec(recRow.Item("FeedLineCaAaIce"), 6)
        Me.FeedLineCaAaIce_1 = DBtoNullableDec(recRow.Item("FeedLineCaAaIce_1"), 6)
        Me.FeedLineCaAaIce_2 = DBtoNullableDec(recRow.Item("FeedLineCaAaIce_2"), 6)
        Me.FeedLineCaAaIce_4 = DBtoNullableDec(recRow.Item("FeedLineCaAaIce_4"), 6)
        Me.FeedLineWtNoIce = DBtoNullableDec(recRow.Item("FeedLineWtNoIce"), 6)
        Me.FeedLineWtIce = DBtoNullableDec(recRow.Item("FeedLineWtIce"), 6)
        Me.FeedLineWtIce_1 = DBtoNullableDec(recRow.Item("FeedLineWtIce_1"), 6)
        Me.FeedLineWtIce_2 = DBtoNullableDec(recRow.Item("FeedLineWtIce_2"), 6)
        Me.FeedLineWtIce_4 = DBtoNullableDec(recRow.Item("FeedLineWtIce_4"), 6)
        Me.FeedLineFaceOffset = DBtoNullableDec(recRow.Item("FeedLineFaceOffset"), 6)
        Me.FeedLineOffsetFrac = DBtoNullableDec(recRow.Item("FeedLineOffsetFrac"), 6)
        Me.FeedLinePerimeterOffsetStartFrac = DBtoNullableDec(recRow.Item("FeedLinePerimeterOffsetStartFrac"), 6)
        Me.FeedLinePerimeterOffsetEndFrac = DBtoNullableDec(recRow.Item("FeedLinePerimeterOffsetEndFrac"), 6)
        Me.FeedLineStartHt = DBtoNullableDec(recRow.Item("FeedLineStartHt"), 6)
        Me.FeedLineEndHt = DBtoNullableDec(recRow.Item("FeedLineEndHt"), 6)
        Me.FeedLineClearSpacing = DBtoNullableDec(recRow.Item("FeedLineClearSpacing"), 6)
        Me.FeedLineRowClearSpacing = DBtoNullableDec(recRow.Item("FeedLineRowClearSpacing"), 6)

    End Sub

    Public Overrides Function SQLInsert() As String
        SQLInsert = CCI_Engineering_Templates.My.Resources.General__INSERT
        SQLInsert = SQLInsert.Replace("[TABLE NAME]", Me.EDSTableName)
        SQLInsert = SQLInsert.Replace("[INSERT FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.Replace("[INSERT VALUES]", Me.SQLInsertValues)

        Return SQLInsert
    End Function

    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""

        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bus_unit.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.structure_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.work_order_seq_num.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@TopLevelID") 'foreign key reference 'tnx_id
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineRec.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineEnabled.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineDatabase.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineDescription.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineClassificationCategory.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineNote.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineNum.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineUseShielding.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ExcludeFeedLineFromTorque.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineNumPerRow.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineFace.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineComponentType.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineGroupTreatmentType.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineRoundClusterDia.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineWidth.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLinePerimeter.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FlatAttachmentEffectiveWidthRatio.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.AutoCalcFlatAttachmentEffectiveWidthRatio.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineShieldingFactorKaNoIce.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineShieldingFactorKaIce.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineAutoCalcKa.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineCaAaNoIce.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineCaAaIce.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineCaAaIce_1.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineCaAaIce_2.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineCaAaIce_4.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineWtNoIce.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineWtIce.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineWtIce_1.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineWtIce_2.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineWtIce_4.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineFaceOffset.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineOffsetFrac.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLinePerimeterOffsetStartFrac.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLinePerimeterOffsetEndFrac.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineStartHt.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineEndHt.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineClearSpacing.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FeedLineRowClearSpacing.ToString.FormatDBValue)

        Return SQLInsertValues
        'Throw New NotImplementedException()
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        SQLInsertFields = SQLInsertFields.AddtoDBString("bus_unit")
        SQLInsertFields = SQLInsertFields.AddtoDBString("structure_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")
        SQLInsertFields = SQLInsertFields.AddtoDBString("work_order_seq_num")
        SQLInsertFields = SQLInsertFields.AddtoDBString("tnx_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineRec")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineEnabled")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineDatabase")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineDescription")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineClassificationCategory")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineNote")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineNum")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineUseShielding")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ExcludeFeedLineFromTorque")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineNumPerRow")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineFace")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineComponentType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineGroupTreatmentType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineRoundClusterDia")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineWidth")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLinePerimeter")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FlatAttachmentEffectiveWidthRatio")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AutoCalcFlatAttachmentEffectiveWidthRatio")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineShieldingFactorKaNoIce")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineShieldingFactorKaIce")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineAutoCalcKa")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineCaAaNoIce")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineCaAaIce")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineCaAaIce_1")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineCaAaIce_2")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineCaAaIce_4")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineWtNoIce")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineWtIce")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineWtIce_1")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineWtIce_2")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineWtIce_4")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineFaceOffset")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineOffsetFrac")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLinePerimeterOffsetStartFrac")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLinePerimeterOffsetEndFrac")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineStartHt")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineEndHt")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineClearSpacing")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FeedLineRowClearSpacing")

        Return SQLInsertFields
        'Throw New NotImplementedException()
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        Throw New NotImplementedException()
    End Function
End Class

<DataContract()>
Partial Public Class tnxDiscreteLoad
    Inherits EDSObjectWithQueries

#Region "Inheritted"
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Discrete Load"
        End Get
    End Property
    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "load.discrete_output"
        End Get
    End Property
#End Region

#Region "Define"
    Public Property parent_id As Integer?

    Private _tnx_id As Integer?
    Private _TowerLoadRec As Integer?
    Private _TowerLoadEnabled As Boolean?
    Private _TowerLoadDatabase As String
    Private _TowerLoadDescription As String
    Private _TowerLoadType As String
    Private _TowerLoadClassificationCategory As String
    Private _TowerLoadNote As String
    Private _TowerLoadNum As Integer?
    Private _TowerLoadFace As Integer?
    Private _TowerOffsetType As String
    Private _TowerOffsetDist As Decimal?
    Private _TowerVertOffset As Decimal?
    Private _TowerLateralOffset As Decimal?
    Private _TowerAzimuthAdjustment As Decimal?
    Private _TowerAppurtSymbol As String
    Private _TowerLoadShieldingFactorKaNoIce As Decimal?
    Private _TowerLoadShieldingFactorKaIce As Decimal?
    Private _TowerLoadAutoCalcKa As Boolean?
    Private _TowerLoadCaAaNoIce As Decimal?
    Private _TowerLoadCaAaIce As Decimal?
    Private _TowerLoadCaAaIce_1 As Decimal?
    Private _TowerLoadCaAaIce_2 As Decimal?
    Private _TowerLoadCaAaIce_4 As Decimal?
    Private _TowerLoadCaAaNoIce_Side As Decimal?
    Private _TowerLoadCaAaIce_Side As Decimal?
    Private _TowerLoadCaAaIce_Side_1 As Decimal?
    Private _TowerLoadCaAaIce_Side_2 As Decimal?
    Private _TowerLoadCaAaIce_Side_4 As Decimal?
    Private _TowerLoadWtNoIce As Decimal?
    Private _TowerLoadWtIce As Decimal?
    Private _TowerLoadWtIce_1 As Decimal?
    Private _TowerLoadWtIce_2 As Decimal?
    Private _TowerLoadWtIce_4 As Decimal?
    Private _TowerLoadStartHt As Decimal?
    Private _TowerLoadEndHt As Decimal?

    <Category("TNX Discrete Load"), Description(""), DisplayName("tnx ID")>
    <DataMember()> Public Property tnx_id() As Integer?
        Get
            Return Me._tnx_id
        End Get
        Set
            Me._tnx_id = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadRec")>
    <DataMember()> Public Property TowerLoadRec() As Integer?
        Get
            Return Me._TowerLoadRec
        End Get
        Set
            Me._TowerLoadRec = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadEnabled")>
    <DataMember()> Public Property TowerLoadEnabled() As Boolean?
        Get
            Return Me._TowerLoadEnabled
        End Get
        Set
            Me._TowerLoadEnabled = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadDatabase")>
    <DataMember()> Public Property TowerLoadDatabase() As String
        Get
            Return Me._TowerLoadDatabase
        End Get
        Set
            Me._TowerLoadDatabase = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadDescription")>
    <DataMember()> Public Property TowerLoadDescription() As String
        Get
            Return Me._TowerLoadDescription
        End Get
        Set
            Me._TowerLoadDescription = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadType")>
    <DataMember()> Public Property TowerLoadType() As String
        Get
            Return Me._TowerLoadType
        End Get
        Set
            Me._TowerLoadType = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadClassificationCategory")>
    <DataMember()> Public Property TowerLoadClassificationCategory() As String
        Get
            Return Me._TowerLoadClassificationCategory
        End Get
        Set
            Me._TowerLoadClassificationCategory = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadNote")>
    <DataMember()> Public Property TowerLoadNote() As String
        Get
            Return Me._TowerLoadNote
        End Get
        Set
            Me._TowerLoadNote = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadNum")>
    <DataMember()> Public Property TowerLoadNum() As Integer?
        Get
            Return Me._TowerLoadNum
        End Get
        Set
            Me._TowerLoadNum = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description("{0 = A, 1 = B,  2 = C, 3 = D}"), DisplayName("TowerLoadFace")>
    <DataMember()> Public Property TowerLoadFace() As Integer?
        Get
            Return Me._TowerLoadFace
        End Get
        Set
            Me._TowerLoadFace = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerOffsetType")>
    <DataMember()> Public Property TowerOffsetType() As String
        Get
            Return Me._TowerOffsetType
        End Get
        Set
            Me._TowerOffsetType = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerOffsetDist")>
    <DataMember()> Public Property TowerOffsetDist() As Decimal?
        Get
            Return Me._TowerOffsetDist
        End Get
        Set
            Me._TowerOffsetDist = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerVertOffset")>
    <DataMember()> Public Property TowerVertOffset() As Decimal?
        Get
            Return Me._TowerVertOffset
        End Get
        Set
            Me._TowerVertOffset = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLateralOffset")>
    <DataMember()> Public Property TowerLateralOffset() As Decimal?
        Get
            Return Me._TowerLateralOffset
        End Get
        Set
            Me._TowerLateralOffset = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerAzimuthAdjustment")>
    <DataMember()> Public Property TowerAzimuthAdjustment() As Decimal?
        Get
            Return Me._TowerAzimuthAdjustment
        End Get
        Set
            Me._TowerAzimuthAdjustment = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerAppurtSymbol")>
    <DataMember()> Public Property TowerAppurtSymbol() As String
        Get
            Return Me._TowerAppurtSymbol
        End Get
        Set
            Me._TowerAppurtSymbol = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadShieldingFactorKaNoIce")>
    <DataMember()> Public Property TowerLoadShieldingFactorKaNoIce() As Decimal?
        Get
            Return Me._TowerLoadShieldingFactorKaNoIce
        End Get
        Set
            Me._TowerLoadShieldingFactorKaNoIce = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadShieldingFactorKaIce")>
    <DataMember()> Public Property TowerLoadShieldingFactorKaIce() As Decimal?
        Get
            Return Me._TowerLoadShieldingFactorKaIce
        End Get
        Set
            Me._TowerLoadShieldingFactorKaIce = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadAutoCalcKa")>
    <DataMember()> Public Property TowerLoadAutoCalcKa() As Boolean?
        Get
            Return Me._TowerLoadAutoCalcKa
        End Get
        Set
            Me._TowerLoadAutoCalcKa = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadCaAaNoIce")>
    <DataMember()> Public Property TowerLoadCaAaNoIce() As Decimal?
        Get
            Return Me._TowerLoadCaAaNoIce
        End Get
        Set
            Me._TowerLoadCaAaNoIce = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadCaAaIce")>
    <DataMember()> Public Property TowerLoadCaAaIce() As Decimal?
        Get
            Return Me._TowerLoadCaAaIce
        End Get
        Set
            Me._TowerLoadCaAaIce = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadCaAaIce_1")>
    <DataMember()> Public Property TowerLoadCaAaIce_1() As Decimal?
        Get
            Return Me._TowerLoadCaAaIce_1
        End Get
        Set
            Me._TowerLoadCaAaIce_1 = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadCaAaIce_2")>
    <DataMember()> Public Property TowerLoadCaAaIce_2() As Decimal?
        Get
            Return Me._TowerLoadCaAaIce_2
        End Get
        Set
            Me._TowerLoadCaAaIce_2 = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadCaAaIce_4")>
    <DataMember()> Public Property TowerLoadCaAaIce_4() As Decimal?
        Get
            Return Me._TowerLoadCaAaIce_4
        End Get
        Set
            Me._TowerLoadCaAaIce_4 = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadCaAaNoIce_Side")>
    <DataMember()> Public Property TowerLoadCaAaNoIce_Side() As Decimal?
        Get
            Return Me._TowerLoadCaAaNoIce_Side
        End Get
        Set
            Me._TowerLoadCaAaNoIce_Side = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadCaAaIce_Side")>
    <DataMember()> Public Property TowerLoadCaAaIce_Side() As Decimal?
        Get
            Return Me._TowerLoadCaAaIce_Side
        End Get
        Set
            Me._TowerLoadCaAaIce_Side = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadCaAaIce_Side_1")>
    <DataMember()> Public Property TowerLoadCaAaIce_Side_1() As Decimal?
        Get
            Return Me._TowerLoadCaAaIce_Side_1
        End Get
        Set
            Me._TowerLoadCaAaIce_Side_1 = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadCaAaIce_Side_2")>
    <DataMember()> Public Property TowerLoadCaAaIce_Side_2() As Decimal?
        Get
            Return Me._TowerLoadCaAaIce_Side_2
        End Get
        Set
            Me._TowerLoadCaAaIce_Side_2 = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadCaAaIce_Side_4")>
    <DataMember()> Public Property TowerLoadCaAaIce_Side_4() As Decimal?
        Get
            Return Me._TowerLoadCaAaIce_Side_4
        End Get
        Set
            Me._TowerLoadCaAaIce_Side_4 = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadWtNoIce")>
    <DataMember()> Public Property TowerLoadWtNoIce() As Decimal?
        Get
            Return Me._TowerLoadWtNoIce
        End Get
        Set
            Me._TowerLoadWtNoIce = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadWtIce")>
    <DataMember()> Public Property TowerLoadWtIce() As Decimal?
        Get
            Return Me._TowerLoadWtIce
        End Get
        Set
            Me._TowerLoadWtIce = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadWtIce_1")>
    <DataMember()> Public Property TowerLoadWtIce_1() As Decimal?
        Get
            Return Me._TowerLoadWtIce_1
        End Get
        Set
            Me._TowerLoadWtIce_1 = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadWtIce_2")>
    <DataMember()> Public Property TowerLoadWtIce_2() As Decimal?
        Get
            Return Me._TowerLoadWtIce_2
        End Get
        Set
            Me._TowerLoadWtIce_2 = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadWtIce_4")>
    <DataMember()> Public Property TowerLoadWtIce_4() As Decimal?
        Get
            Return Me._TowerLoadWtIce_4
        End Get
        Set
            Me._TowerLoadWtIce_4 = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadStartHt")>
    <DataMember()> Public Property TowerLoadStartHt() As Decimal?
        Get
            Return Me._TowerLoadStartHt
        End Get
        Set
            Me._TowerLoadStartHt = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadEndHt")>
    <DataMember()> Public Property TowerLoadEndHt() As Decimal?
        Get
            Return Me._TowerLoadEndHt
        End Get
        Set
            Me._TowerLoadEndHt = Value
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
        Dim otherToCompare As tnxDiscreteLoad = TryCast(other, tnxDiscreteLoad)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.TowerLoadRec.CheckChange(otherToCompare.TowerLoadRec, changes, categoryName, "Towerloadrec"), Equals, False)
        Equals = If(Me.TowerLoadEnabled.CheckChange(otherToCompare.TowerLoadEnabled, changes, categoryName, "Towerloadenabled"), Equals, False)
        Equals = If(Me.TowerLoadDatabase.CheckChange(otherToCompare.TowerLoadDatabase, changes, categoryName, "Towerloaddatabase"), Equals, False)
        Equals = If(Me.TowerLoadDescription.CheckChange(otherToCompare.TowerLoadDescription, changes, categoryName, "Towerloaddescription"), Equals, False)
        Equals = If(Me.TowerLoadType.CheckChange(otherToCompare.TowerLoadType, changes, categoryName, "Towerloadtype"), Equals, False)
        Equals = If(Me.TowerLoadClassificationCategory.CheckChange(otherToCompare.TowerLoadClassificationCategory, changes, categoryName, "Towerloadclassificationcategory"), Equals, False)
        Equals = If(Me.TowerLoadNote.CheckChange(otherToCompare.TowerLoadNote, changes, categoryName, "Towerloadnote"), Equals, False)
        Equals = If(Me.TowerLoadNum.CheckChange(otherToCompare.TowerLoadNum, changes, categoryName, "Towerloadnum"), Equals, False)
        Equals = If(Me.TowerLoadFace.CheckChange(otherToCompare.TowerLoadFace, changes, categoryName, "Towerloadface"), Equals, False)
        Equals = If(Me.TowerOffsetType.CheckChange(otherToCompare.TowerOffsetType, changes, categoryName, "Toweroffsettype"), Equals, False)
        Equals = If(Me.TowerOffsetDist.CheckChange(otherToCompare.TowerOffsetDist, changes, categoryName, "Toweroffsetdist"), Equals, False)
        Equals = If(Me.TowerVertOffset.CheckChange(otherToCompare.TowerVertOffset, changes, categoryName, "Towervertoffset"), Equals, False)
        Equals = If(Me.TowerLateralOffset.CheckChange(otherToCompare.TowerLateralOffset, changes, categoryName, "Towerlateraloffset"), Equals, False)
        Equals = If(Me.TowerAzimuthAdjustment.CheckChange(otherToCompare.TowerAzimuthAdjustment, changes, categoryName, "Towerazimuthadjustment"), Equals, False)
        Equals = If(Me.TowerAppurtSymbol.CheckChange(otherToCompare.TowerAppurtSymbol, changes, categoryName, "Towerappurtsymbol"), Equals, False)
        Equals = If(Me.TowerLoadShieldingFactorKaNoIce.CheckChange(otherToCompare.TowerLoadShieldingFactorKaNoIce, changes, categoryName, "Towerloadshieldingfactorkanoice"), Equals, False)
        Equals = If(Me.TowerLoadShieldingFactorKaIce.CheckChange(otherToCompare.TowerLoadShieldingFactorKaIce, changes, categoryName, "Towerloadshieldingfactorkaice"), Equals, False)
        Equals = If(Me.TowerLoadAutoCalcKa.CheckChange(otherToCompare.TowerLoadAutoCalcKa, changes, categoryName, "Towerloadautocalcka"), Equals, False)
        Equals = If(Me.TowerLoadCaAaNoIce.CheckChange(otherToCompare.TowerLoadCaAaNoIce, changes, categoryName, "Towerloadcaaanoice"), Equals, False)
        Equals = If(Me.TowerLoadCaAaIce.CheckChange(otherToCompare.TowerLoadCaAaIce, changes, categoryName, "Towerloadcaaaice"), Equals, False)
        Equals = If(Me.TowerLoadCaAaIce_1.CheckChange(otherToCompare.TowerLoadCaAaIce_1, changes, categoryName, "Towerloadcaaaice 1"), Equals, False)
        Equals = If(Me.TowerLoadCaAaIce_2.CheckChange(otherToCompare.TowerLoadCaAaIce_2, changes, categoryName, "Towerloadcaaaice 2"), Equals, False)
        Equals = If(Me.TowerLoadCaAaIce_4.CheckChange(otherToCompare.TowerLoadCaAaIce_4, changes, categoryName, "Towerloadcaaaice 4"), Equals, False)
        Equals = If(Me.TowerLoadCaAaNoIce_Side.CheckChange(otherToCompare.TowerLoadCaAaNoIce_Side, changes, categoryName, "Towerloadcaaanoice Side"), Equals, False)
        Equals = If(Me.TowerLoadCaAaIce_Side.CheckChange(otherToCompare.TowerLoadCaAaIce_Side, changes, categoryName, "Towerloadcaaaice Side"), Equals, False)
        Equals = If(Me.TowerLoadCaAaIce_Side_1.CheckChange(otherToCompare.TowerLoadCaAaIce_Side_1, changes, categoryName, "Towerloadcaaaice Side 1"), Equals, False)
        Equals = If(Me.TowerLoadCaAaIce_Side_2.CheckChange(otherToCompare.TowerLoadCaAaIce_Side_2, changes, categoryName, "Towerloadcaaaice Side 2"), Equals, False)
        Equals = If(Me.TowerLoadCaAaIce_Side_4.CheckChange(otherToCompare.TowerLoadCaAaIce_Side_4, changes, categoryName, "Towerloadcaaaice Side 4"), Equals, False)
        Equals = If(Me.TowerLoadWtNoIce.CheckChange(otherToCompare.TowerLoadWtNoIce, changes, categoryName, "Towerloadwtnoice"), Equals, False)
        Equals = If(Me.TowerLoadWtIce.CheckChange(otherToCompare.TowerLoadWtIce, changes, categoryName, "Towerloadwtice"), Equals, False)
        Equals = If(Me.TowerLoadWtIce_1.CheckChange(otherToCompare.TowerLoadWtIce_1, changes, categoryName, "Towerloadwtice 1"), Equals, False)
        Equals = If(Me.TowerLoadWtIce_2.CheckChange(otherToCompare.TowerLoadWtIce_2, changes, categoryName, "Towerloadwtice 2"), Equals, False)
        Equals = If(Me.TowerLoadWtIce_4.CheckChange(otherToCompare.TowerLoadWtIce_4, changes, categoryName, "Towerloadwtice 4"), Equals, False)
        Equals = If(Me.TowerLoadStartHt.CheckChange(otherToCompare.TowerLoadStartHt, changes, categoryName, "Towerloadstartht"), Equals, False)
        Equals = If(Me.TowerLoadEndHt.CheckChange(otherToCompare.TowerLoadEndHt, changes, categoryName, "Towerloadendht"), Equals, False)

        Return Equals
    End Function
#End Region

    Public Sub New(ByVal recRow As DataRow, Optional ByVal Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Me.ID = DBtoNullableInt(recRow.Item("ID"))
        'Me.bus_unit = DBtoNullableInt(recRow.Item("bus_unit"))
        'Me.structure_id = DBtoStr(recRow.Item("structure_id"))
        'Me.modified_person_id = DBtoNullableInt(recRow.Item("modified_person_id"))
        'Me.process_stage = DBtoStr(recRow.Item("process_stage"))
        'Me.work_order_seq_num = DBtoNullableInt(recRow.Item("work_order_seq_num"))
        Me.tnx_id = DBtoNullableInt(recRow.Item("tnx_id"))
        Me.TowerLoadRec = DBtoNullableInt(recRow.Item("TowerLoadRec"))
        Me.TowerLoadEnabled = DBtoNullableBool(recRow.Item("TowerLoadEnabled"))
        Me.TowerLoadDatabase = DBtoStr(recRow.Item("TowerLoadDatabase"))
        Me.TowerLoadDescription = DBtoStr(recRow.Item("TowerLoadDescription"))
        Me.TowerLoadType = DBtoStr(recRow.Item("TowerLoadType"))
        Me.TowerLoadClassificationCategory = DBtoStr(recRow.Item("TowerLoadClassificationCategory"))
        Me.TowerLoadNote = DBtoStr(recRow.Item("TowerLoadNote"))
        Me.TowerLoadNum = DBtoNullableInt(recRow.Item("TowerLoadNum"))
        Me.TowerLoadFace = DBtoNullableInt(recRow.Item("TowerLoadFace"))
        Me.TowerOffsetType = DBtoStr(recRow.Item("TowerOffsetType"))
        Me.TowerOffsetDist = DBtoNullableDec(recRow.Item("TowerOffsetDist"), 6)
        Me.TowerVertOffset = DBtoNullableDec(recRow.Item("TowerVertOffset"), 6)
        Me.TowerLateralOffset = DBtoNullableDec(recRow.Item("TowerLateralOffset"), 6)
        Me.TowerAzimuthAdjustment = DBtoNullableDec(recRow.Item("TowerAzimuthAdjustment"), 6)
        Me.TowerAppurtSymbol = DBtoStr(recRow.Item("TowerAppurtSymbol"))
        Me.TowerLoadShieldingFactorKaNoIce = DBtoNullableDec(recRow.Item("TowerLoadShieldingFactorKaNoIce"), 6)
        Me.TowerLoadShieldingFactorKaIce = DBtoNullableDec(recRow.Item("TowerLoadShieldingFactorKaIce"), 6)
        Me.TowerLoadAutoCalcKa = DBtoNullableBool(recRow.Item("TowerLoadAutoCalcKa"))
        Me.TowerLoadCaAaNoIce = DBtoNullableDec(recRow.Item("TowerLoadCaAaNoIce"), 6)
        Me.TowerLoadCaAaIce = DBtoNullableDec(recRow.Item("TowerLoadCaAaIce"), 6)
        Me.TowerLoadCaAaIce_1 = DBtoNullableDec(recRow.Item("TowerLoadCaAaIce_1"), 6)
        Me.TowerLoadCaAaIce_2 = DBtoNullableDec(recRow.Item("TowerLoadCaAaIce_2"), 6)
        Me.TowerLoadCaAaIce_4 = DBtoNullableDec(recRow.Item("TowerLoadCaAaIce_4"), 6)
        Me.TowerLoadCaAaNoIce_Side = DBtoNullableDec(recRow.Item("TowerLoadCaAaNoIce_Side"), 6)
        Me.TowerLoadCaAaIce_Side = DBtoNullableDec(recRow.Item("TowerLoadCaAaIce_Side"), 6)
        Me.TowerLoadCaAaIce_Side_1 = DBtoNullableDec(recRow.Item("TowerLoadCaAaIce_Side_1"), 6)
        Me.TowerLoadCaAaIce_Side_2 = DBtoNullableDec(recRow.Item("TowerLoadCaAaIce_Side_2"), 6)
        Me.TowerLoadCaAaIce_Side_4 = DBtoNullableDec(recRow.Item("TowerLoadCaAaIce_Side_4"), 6)
        Me.TowerLoadWtNoIce = DBtoNullableDec(recRow.Item("TowerLoadWtNoIce"), 6)
        Me.TowerLoadWtIce = DBtoNullableDec(recRow.Item("TowerLoadWtIce"), 6)
        Me.TowerLoadWtIce_1 = DBtoNullableDec(recRow.Item("TowerLoadWtIce_1"), 6)
        Me.TowerLoadWtIce_2 = DBtoNullableDec(recRow.Item("TowerLoadWtIce_2"), 6)
        Me.TowerLoadWtIce_4 = DBtoNullableDec(recRow.Item("TowerLoadWtIce_4"), 6)
        Me.TowerLoadStartHt = DBtoNullableDec(recRow.Item("TowerLoadStartHt"), 6)
        Me.TowerLoadEndHt = DBtoNullableDec(recRow.Item("TowerLoadEndHt"), 6)

    End Sub

    Public Overrides Function SQLInsert() As String
        SQLInsert = CCI_Engineering_Templates.My.Resources.General__INSERT
        SQLInsert = SQLInsert.Replace("[TABLE NAME]", Me.EDSTableName)
        SQLInsert = SQLInsert.Replace("[INSERT FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.Replace("[INSERT VALUES]", Me.SQLInsertValues)

        Return SQLInsert
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        SQLInsertFields = SQLInsertFields.AddtoDBString("bus_unit")
        SQLInsertFields = SQLInsertFields.AddtoDBString("structure_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")
        SQLInsertFields = SQLInsertFields.AddtoDBString("work_order_seq_num")
        SQLInsertFields = SQLInsertFields.AddtoDBString("tnx_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLoadRec")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLoadEnabled")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLoadDatabase")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLoadDescription")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLoadType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLoadClassificationCategory")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLoadNote")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLoadNum")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLoadFace")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerOffsetType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerOffsetDist")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerVertOffset")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLateralOffset")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerAzimuthAdjustment")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerAppurtSymbol")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLoadShieldingFactorKaNoIce")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLoadShieldingFactorKaIce")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLoadAutoCalcKa")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLoadCaAaNoIce")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLoadCaAaIce")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLoadCaAaIce_1")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLoadCaAaIce_2")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLoadCaAaIce_4")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLoadCaAaNoIce_Side")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLoadCaAaIce_Side")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLoadCaAaIce_Side_1")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLoadCaAaIce_Side_2")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLoadCaAaIce_Side_4")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLoadWtNoIce")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLoadWtIce")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLoadWtIce_1")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLoadWtIce_2")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLoadWtIce_4")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLoadStartHt")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerLoadEndHt")

        Return SQLInsertFields
        'Throw New NotImplementedException()
    End Function

    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""

        'Dim idStr As String = ""
        'If Me.Parent.ID IsNot Nothing Then idStr = Me.Parent.ID Else idStr = "@TopLevelID"

        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bus_unit.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.structure_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.work_order_seq_num.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@TopLevelID") 'foreign key reference  'tnx_id
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLoadRec.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLoadEnabled.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLoadDatabase.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLoadDescription.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLoadType.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLoadClassificationCategory.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLoadNote.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLoadNum.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLoadFace.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerOffsetType.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerOffsetDist.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerVertOffset.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLateralOffset.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerAzimuthAdjustment.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerAppurtSymbol.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLoadShieldingFactorKaNoIce.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLoadShieldingFactorKaIce.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLoadAutoCalcKa.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLoadCaAaNoIce.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLoadCaAaIce.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLoadCaAaIce_1.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLoadCaAaIce_2.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLoadCaAaIce_4.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLoadCaAaNoIce_Side.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLoadCaAaIce_Side.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLoadCaAaIce_Side_1.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLoadCaAaIce_Side_2.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLoadCaAaIce_Side_4.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLoadWtNoIce.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLoadWtIce.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLoadWtIce_1.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLoadWtIce_2.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLoadWtIce_4.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLoadStartHt.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TowerLoadEndHt.ToString.FormatDBValue)

        Return SQLInsertValues
        'Throw New NotImplementedException()
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        Throw New NotImplementedException()
    End Function
End Class

<DataContract()>
Partial Public Class tnxDish
    Inherits EDSObjectWithQueries

#Region "Inheritted"
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Dish"
        End Get
    End Property
    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "load.dish_output"
        End Get
    End Property
#End Region

#Region "Define"
    Public Property parent_id As Integer?

    Private _tnx_id As Integer?
    Private _DishRec As Integer?
    Private _DishEnabled As Boolean?
    Private _DishDatabase As String
    Private _DishDescription As String
    Private _DishClassificationCategory As String
    Private _DishNote As String
    Private _DishNum As Integer?
    Private _DishFace As Integer?
    Private _DishType As String
    Private _DishOffsetType As String
    Private _DishVertOffset As Decimal?
    Private _DishLateralOffset As Decimal?
    Private _DishOffsetDist As Decimal?
    Private _DishArea As Decimal?
    Private _DishAreaIce As Decimal?
    Private _DishAreaIce_1 As Decimal?
    Private _DishAreaIce_2 As Decimal?
    Private _DishAreaIce_4 As Decimal?
    Private _DishDiameter As Decimal?
    Private _DishWtNoIce As Decimal?
    Private _DishWtIce As Decimal?
    Private _DishWtIce_1 As Decimal?
    Private _DishWtIce_2 As Decimal?
    Private _DishWtIce_4 As Decimal?
    Private _DishStartHt As Decimal?
    Private _DishAzimuthAdjustment As Decimal?
    Private _DishBeamWidth As Decimal?

    <Category("TNX Dish"), Description(""), DisplayName("tnx ID")>
    <DataMember()> Public Property tnx_id() As Integer?
        Get
            Return Me._tnx_id
        End Get
        Set
            Me._tnx_id = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishRec")>
    <DataMember()> Public Property DishRec() As Integer?
        Get
            Return Me._DishRec
        End Get
        Set
            Me._DishRec = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishEnabled")>
    <DataMember()> Public Property DishEnabled() As Boolean?
        Get
            Return Me._DishEnabled
        End Get
        Set
            Me._DishEnabled = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishDatabase")>
    <DataMember()> Public Property DishDatabase() As String
        Get
            Return Me._DishDatabase
        End Get
        Set
            Me._DishDatabase = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishDescription")>
    <DataMember()> Public Property DishDescription() As String
        Get
            Return Me._DishDescription
        End Get
        Set
            Me._DishDescription = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishClassificationCategory")>
    <DataMember()> Public Property DishClassificationCategory() As String
        Get
            Return Me._DishClassificationCategory
        End Get
        Set
            Me._DishClassificationCategory = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishNote")>
    <DataMember()> Public Property DishNote() As String
        Get
            Return Me._DishNote
        End Get
        Set
            Me._DishNote = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishNum")>
    <DataMember()> Public Property DishNum() As Integer?
        Get
            Return Me._DishNum
        End Get
        Set
            Me._DishNum = Value
        End Set
    End Property
    <Category("TNX Dish"), Description("{0 = A, 1 = B,  2 = C, 3 = D}"), DisplayName("DishFace")>
    <DataMember()> Public Property DishFace() As Integer?
        Get
            Return Me._DishFace
        End Get
        Set
            Me._DishFace = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishType")>
    <DataMember()> Public Property DishType() As String
        Get
            Return Me._DishType
        End Get
        Set
            Me._DishType = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishOffsetType")>
    <DataMember()> Public Property DishOffsetType() As String
        Get
            Return Me._DishOffsetType
        End Get
        Set
            Me._DishOffsetType = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishVertOffset")>
    <DataMember()> Public Property DishVertOffset() As Decimal?
        Get
            Return Me._DishVertOffset
        End Get
        Set
            Me._DishVertOffset = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishLateralOffset")>
    <DataMember()> Public Property DishLateralOffset() As Decimal?
        Get
            Return Me._DishLateralOffset
        End Get
        Set
            Me._DishLateralOffset = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishOffsetDist")>
    <DataMember()> Public Property DishOffsetDist() As Decimal?
        Get
            Return Me._DishOffsetDist
        End Get
        Set
            Me._DishOffsetDist = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishArea")>
    <DataMember()> Public Property DishArea() As Decimal?
        Get
            Return Me._DishArea
        End Get
        Set
            Me._DishArea = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishAreaIce")>
    <DataMember()> Public Property DishAreaIce() As Decimal?
        Get
            Return Me._DishAreaIce
        End Get
        Set
            Me._DishAreaIce = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishAreaIce_1")>
    <DataMember()> Public Property DishAreaIce_1() As Decimal?
        Get
            Return Me._DishAreaIce_1
        End Get
        Set
            Me._DishAreaIce_1 = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishAreaIce_2")>
    <DataMember()> Public Property DishAreaIce_2() As Decimal?
        Get
            Return Me._DishAreaIce_2
        End Get
        Set
            Me._DishAreaIce_2 = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishAreaIce_4")>
    <DataMember()> Public Property DishAreaIce_4() As Decimal?
        Get
            Return Me._DishAreaIce_4
        End Get
        Set
            Me._DishAreaIce_4 = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishDiameter")>
    <DataMember()> Public Property DishDiameter() As Decimal?
        Get
            Return Me._DishDiameter
        End Get
        Set
            Me._DishDiameter = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishWtNoIce")>
    <DataMember()> Public Property DishWtNoIce() As Decimal?
        Get
            Return Me._DishWtNoIce
        End Get
        Set
            Me._DishWtNoIce = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishWtIce")>
    <DataMember()> Public Property DishWtIce() As Decimal?
        Get
            Return Me._DishWtIce
        End Get
        Set
            Me._DishWtIce = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishWtIce_1")>
    <DataMember()> Public Property DishWtIce_1() As Decimal?
        Get
            Return Me._DishWtIce_1
        End Get
        Set
            Me._DishWtIce_1 = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishWtIce_2")>
    <DataMember()> Public Property DishWtIce_2() As Decimal?
        Get
            Return Me._DishWtIce_2
        End Get
        Set
            Me._DishWtIce_2 = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishWtIce_4")>
    <DataMember()> Public Property DishWtIce_4() As Decimal?
        Get
            Return Me._DishWtIce_4
        End Get
        Set
            Me._DishWtIce_4 = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishStartHt")>
    <DataMember()> Public Property DishStartHt() As Decimal?
        Get
            Return Me._DishStartHt
        End Get
        Set
            Me._DishStartHt = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishAzimuthAdjustment")>
    <DataMember()> Public Property DishAzimuthAdjustment() As Decimal?
        Get
            Return Me._DishAzimuthAdjustment
        End Get
        Set
            Me._DishAzimuthAdjustment = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishBeamWidth")>
    <DataMember()> Public Property DishBeamWidth() As Decimal?
        Get
            Return Me._DishBeamWidth
        End Get
        Set
            Me._DishBeamWidth = Value
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
        Dim otherToCompare As tnxDish = TryCast(other, tnxDish)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.DishRec.CheckChange(otherToCompare.DishRec, changes, categoryName, "Dishrec"), Equals, False)
        Equals = If(Me.DishEnabled.CheckChange(otherToCompare.DishEnabled, changes, categoryName, "Dishenabled"), Equals, False)
        Equals = If(Me.DishDatabase.CheckChange(otherToCompare.DishDatabase, changes, categoryName, "Dishdatabase"), Equals, False)
        Equals = If(Me.DishDescription.CheckChange(otherToCompare.DishDescription, changes, categoryName, "Dishdescription"), Equals, False)
        Equals = If(Me.DishClassificationCategory.CheckChange(otherToCompare.DishClassificationCategory, changes, categoryName, "Dishclassificationcategory"), Equals, False)
        Equals = If(Me.DishNote.CheckChange(otherToCompare.DishNote, changes, categoryName, "Dishnote"), Equals, False)
        Equals = If(Me.DishNum.CheckChange(otherToCompare.DishNum, changes, categoryName, "Dishnum"), Equals, False)
        Equals = If(Me.DishFace.CheckChange(otherToCompare.DishFace, changes, categoryName, "Dishface"), Equals, False)
        Equals = If(Me.DishType.CheckChange(otherToCompare.DishType, changes, categoryName, "Dishtype"), Equals, False)
        Equals = If(Me.DishOffsetType.CheckChange(otherToCompare.DishOffsetType, changes, categoryName, "Dishoffsettype"), Equals, False)
        Equals = If(Me.DishVertOffset.CheckChange(otherToCompare.DishVertOffset, changes, categoryName, "Dishvertoffset"), Equals, False)
        Equals = If(Me.DishLateralOffset.CheckChange(otherToCompare.DishLateralOffset, changes, categoryName, "Dishlateraloffset"), Equals, False)
        Equals = If(Me.DishOffsetDist.CheckChange(otherToCompare.DishOffsetDist, changes, categoryName, "Dishoffsetdist"), Equals, False)
        Equals = If(Me.DishArea.CheckChange(otherToCompare.DishArea, changes, categoryName, "Disharea"), Equals, False)
        Equals = If(Me.DishAreaIce.CheckChange(otherToCompare.DishAreaIce, changes, categoryName, "Dishareaice"), Equals, False)
        Equals = If(Me.DishAreaIce_1.CheckChange(otherToCompare.DishAreaIce_1, changes, categoryName, "Dishareaice 1"), Equals, False)
        Equals = If(Me.DishAreaIce_2.CheckChange(otherToCompare.DishAreaIce_2, changes, categoryName, "Dishareaice 2"), Equals, False)
        Equals = If(Me.DishAreaIce_4.CheckChange(otherToCompare.DishAreaIce_4, changes, categoryName, "Dishareaice 4"), Equals, False)
        Equals = If(Me.DishDiameter.CheckChange(otherToCompare.DishDiameter, changes, categoryName, "Dishdiameter"), Equals, False)
        Equals = If(Me.DishWtNoIce.CheckChange(otherToCompare.DishWtNoIce, changes, categoryName, "Dishwtnoice"), Equals, False)
        Equals = If(Me.DishWtIce.CheckChange(otherToCompare.DishWtIce, changes, categoryName, "Dishwtice"), Equals, False)
        Equals = If(Me.DishWtIce_1.CheckChange(otherToCompare.DishWtIce_1, changes, categoryName, "Dishwtice 1"), Equals, False)
        Equals = If(Me.DishWtIce_2.CheckChange(otherToCompare.DishWtIce_2, changes, categoryName, "Dishwtice 2"), Equals, False)
        Equals = If(Me.DishWtIce_4.CheckChange(otherToCompare.DishWtIce_4, changes, categoryName, "Dishwtice 4"), Equals, False)
        Equals = If(Me.DishStartHt.CheckChange(otherToCompare.DishStartHt, changes, categoryName, "Dishstartht"), Equals, False)
        Equals = If(Me.DishAzimuthAdjustment.CheckChange(otherToCompare.DishAzimuthAdjustment, changes, categoryName, "Dishazimuthadjustment"), Equals, False)
        Equals = If(Me.DishBeamWidth.CheckChange(otherToCompare.DishBeamWidth, changes, categoryName, "Dishbeamwidth"), Equals, False)

        Return Equals
    End Function
#End Region

    Public Sub New(ByVal recRow As DataRow, Optional ByVal Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Me.ID = DBtoNullableInt(recRow.Item("ID"))
        'Me.bus_unit = DBtoNullableInt(recRow.Item("bus_unit"))
        'Me.structure_id = DBtoStr(recRow.Item("structure_id"))
        'Me.modified_person_id = DBtoNullableInt(recRow.Item("modified_person_id"))
        'Me.process_stage = DBtoStr(recRow.Item("process_stage"))
        'Me.work_order_seq_num = DBtoNullableInt(recRow.Item("work_order_seq_num"))
        Me.tnx_id = DBtoNullableInt(recRow.Item("tnx_id"))
        Me.DishRec = DBtoNullableInt(recRow.Item("DishRec"))
        Me.DishEnabled = DBtoNullableBool(recRow.Item("DishEnabled"))
        Me.DishDatabase = DBtoStr(recRow.Item("DishDatabase"))
        Me.DishDescription = DBtoStr(recRow.Item("DishDescription"))
        Me.DishClassificationCategory = DBtoStr(recRow.Item("DishClassificationCategory"))
        Me.DishNote = DBtoStr(recRow.Item("DishNote"))
        Me.DishNum = DBtoNullableInt(recRow.Item("DishNum"))
        Me.DishFace = DBtoNullableInt(recRow.Item("DishFace"))
        Me.DishType = DBtoStr(recRow.Item("DishType"))
        Me.DishOffsetType = DBtoStr(recRow.Item("DishOffsetType"))
        Me.DishVertOffset = DBtoNullableDec(recRow.Item("DishVertOffset"), 6)
        Me.DishLateralOffset = DBtoNullableDec(recRow.Item("DishLateralOffset"), 6)
        Me.DishOffsetDist = DBtoNullableDec(recRow.Item("DishOffsetDist"), 6)
        Me.DishArea = DBtoNullableDec(recRow.Item("DishArea"), 6)
        Me.DishAreaIce = DBtoNullableDec(recRow.Item("DishAreaIce"), 6)
        Me.DishAreaIce_1 = DBtoNullableDec(recRow.Item("DishAreaIce_1"), 6)
        Me.DishAreaIce_2 = DBtoNullableDec(recRow.Item("DishAreaIce_2"), 6)
        Me.DishAreaIce_4 = DBtoNullableDec(recRow.Item("DishAreaIce_4"), 6)
        Me.DishDiameter = DBtoNullableDec(recRow.Item("DishDiameter"), 6)
        Me.DishWtNoIce = DBtoNullableDec(recRow.Item("DishWtNoIce"), 6)
        Me.DishWtIce = DBtoNullableDec(recRow.Item("DishWtIce"), 6)
        Me.DishWtIce_1 = DBtoNullableDec(recRow.Item("DishWtIce_1"), 6)
        Me.DishWtIce_2 = DBtoNullableDec(recRow.Item("DishWtIce_2"), 6)
        Me.DishWtIce_4 = DBtoNullableDec(recRow.Item("DishWtIce_4"), 6)
        Me.DishStartHt = DBtoNullableDec(recRow.Item("DishStartHt"), 6)
        Me.DishAzimuthAdjustment = DBtoNullableDec(recRow.Item("DishAzimuthAdjustment"), 6)
        Me.DishBeamWidth = DBtoNullableDec(recRow.Item("DishBeamWidth"), 6)

    End Sub

    Public Overrides Function SQLInsert() As String
        SQLInsert = CCI_Engineering_Templates.My.Resources.General__INSERT
        SQLInsert = SQLInsert.Replace("[TABLE NAME]", Me.EDSTableName)
        SQLInsert = SQLInsert.Replace("[INSERT FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.Replace("[INSERT VALUES]", Me.SQLInsertValues)

        Return SQLInsert
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        SQLInsertFields = SQLInsertFields.AddtoDBString("bus_unit")
        SQLInsertFields = SQLInsertFields.AddtoDBString("structure_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")
        SQLInsertFields = SQLInsertFields.AddtoDBString("work_order_seq_num")
        SQLInsertFields = SQLInsertFields.AddtoDBString("tnx_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DishRec")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DishEnabled")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DishDatabase")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DishDescription")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DishClassificationCategory")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DishNote")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DishNum")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DishFace")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DishType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DishOffsetType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DishVertOffset")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DishLateralOffset")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DishOffsetDist")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DishArea")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DishAreaIce")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DishAreaIce_1")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DishAreaIce_2")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DishAreaIce_4")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DishDiameter")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DishWtNoIce")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DishWtIce")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DishWtIce_1")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DishWtIce_2")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DishWtIce_4")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DishStartHt")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DishAzimuthAdjustment")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DishBeamWidth")

        Return SQLInsertFields
        'Throw New NotImplementedException()
    End Function

    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""

        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bus_unit.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.structure_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.work_order_seq_num.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@TopLevelID") 'foreign key reference 'tnx_id
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.DishRec.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.DishEnabled.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.DishDatabase.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.DishDescription.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.DishClassificationCategory.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.DishNote.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.DishNum.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.DishFace.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.DishType.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.DishOffsetType.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.DishVertOffset.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.DishLateralOffset.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.DishOffsetDist.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.DishArea.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.DishAreaIce.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.DishAreaIce_1.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.DishAreaIce_2.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.DishAreaIce_4.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.DishDiameter.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.DishWtNoIce.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.DishWtIce.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.DishWtIce_1.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.DishWtIce_2.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.DishWtIce_4.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.DishStartHt.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.DishAzimuthAdjustment.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.DishBeamWidth.ToString.FormatDBValue)

        Return SQLInsertValues
        'Throw New NotImplementedException()
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        Throw New NotImplementedException()
    End Function
End Class

<DataContract()>
Partial Public Class tnxUserForce
    Inherits EDSObjectWithQueries

#Region "Inheritted"
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "User Force"
        End Get
    End Property
    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "load.user_force_output"
        End Get
    End Property
#End Region

#Region "Define"
    Public Property parent_id As Integer?

    Private _tnx_id As Integer?
    Private _UserForceRec As Integer?
    Private _UserForceEnabled As Boolean?
    Private _UserForceDescription As String
    Private _UserForceStartHt As Decimal?
    Private _UserForceOffset As Decimal?
    Private _UserForceAzimuth As Decimal?
    Private _UserForceFxNoIce As Decimal?
    Private _UserForceFzNoIce As Decimal?
    Private _UserForceAxialNoIce As Decimal?
    Private _UserForceShearNoIce As Decimal?
    Private _UserForceCaAcNoIce As Decimal?
    Private _UserForceFxIce As Decimal?
    Private _UserForceFzIce As Decimal?
    Private _UserForceAxialIce As Decimal?
    Private _UserForceShearIce As Decimal?
    Private _UserForceCaAcIce As Decimal?
    Private _UserForceFxService As Decimal?
    Private _UserForceFzService As Decimal?
    Private _UserForceAxialService As Decimal?
    Private _UserForceShearService As Decimal?
    Private _UserForceCaAcService As Decimal?
    Private _UserForceEhx As Decimal?
    Private _UserForceEhz As Decimal?
    Private _UserForceEv As Decimal?
    Private _UserForceEh As Decimal?

    <Category("TNX User Force"), Description(""), DisplayName("tnx ID")>
    <DataMember()> Public Property tnx_id() As Integer?
        Get
            Return Me._tnx_id
        End Get
        Set
            Me._tnx_id = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceRec")>
    <DataMember()> Public Property UserForceRec() As Integer?
        Get
            Return Me._UserForceRec
        End Get
        Set
            Me._UserForceRec = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceEnabled")>
    <DataMember()> Public Property UserForceEnabled() As Boolean?
        Get
            Return Me._UserForceEnabled
        End Get
        Set
            Me._UserForceEnabled = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceDescription")>
    <DataMember()> Public Property UserForceDescription() As String
        Get
            Return Me._UserForceDescription
        End Get
        Set
            Me._UserForceDescription = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceStartHt")>
    <DataMember()> Public Property UserForceStartHt() As Decimal?
        Get
            Return Me._UserForceStartHt
        End Get
        Set
            Me._UserForceStartHt = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceOffset")>
    <DataMember()> Public Property UserForceOffset() As Decimal?
        Get
            Return Me._UserForceOffset
        End Get
        Set
            Me._UserForceOffset = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceAzimuth")>
    <DataMember()> Public Property UserForceAzimuth() As Decimal?
        Get
            Return Me._UserForceAzimuth
        End Get
        Set
            Me._UserForceAzimuth = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceFxNoIce")>
    <DataMember()> Public Property UserForceFxNoIce() As Decimal?
        Get
            Return Me._UserForceFxNoIce
        End Get
        Set
            Me._UserForceFxNoIce = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceFzNoIce")>
    <DataMember()> Public Property UserForceFzNoIce() As Decimal?
        Get
            Return Me._UserForceFzNoIce
        End Get
        Set
            Me._UserForceFzNoIce = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceAxialNoIce")>
    <DataMember()> Public Property UserForceAxialNoIce() As Decimal?
        Get
            Return Me._UserForceAxialNoIce
        End Get
        Set
            Me._UserForceAxialNoIce = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceShearNoIce")>
    <DataMember()> Public Property UserForceShearNoIce() As Decimal?
        Get
            Return Me._UserForceShearNoIce
        End Get
        Set
            Me._UserForceShearNoIce = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceCaAcNoIce")>
    <DataMember()> Public Property UserForceCaAcNoIce() As Decimal?
        Get
            Return Me._UserForceCaAcNoIce
        End Get
        Set
            Me._UserForceCaAcNoIce = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceFxIce")>
    <DataMember()> Public Property UserForceFxIce() As Decimal?
        Get
            Return Me._UserForceFxIce
        End Get
        Set
            Me._UserForceFxIce = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceFzIce")>
    <DataMember()> Public Property UserForceFzIce() As Decimal?
        Get
            Return Me._UserForceFzIce
        End Get
        Set
            Me._UserForceFzIce = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceAxialIce")>
    <DataMember()> Public Property UserForceAxialIce() As Decimal?
        Get
            Return Me._UserForceAxialIce
        End Get
        Set
            Me._UserForceAxialIce = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceShearIce")>
    <DataMember()> Public Property UserForceShearIce() As Decimal?
        Get
            Return Me._UserForceShearIce
        End Get
        Set
            Me._UserForceShearIce = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceCaAcIce")>
    <DataMember()> Public Property UserForceCaAcIce() As Decimal?
        Get
            Return Me._UserForceCaAcIce
        End Get
        Set
            Me._UserForceCaAcIce = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceFxService")>
    <DataMember()> Public Property UserForceFxService() As Decimal?
        Get
            Return Me._UserForceFxService
        End Get
        Set
            Me._UserForceFxService = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceFzService")>
    <DataMember()> Public Property UserForceFzService() As Decimal?
        Get
            Return Me._UserForceFzService
        End Get
        Set
            Me._UserForceFzService = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceAxialService")>
    <DataMember()> Public Property UserForceAxialService() As Decimal?
        Get
            Return Me._UserForceAxialService
        End Get
        Set
            Me._UserForceAxialService = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceShearService")>
    <DataMember()> Public Property UserForceShearService() As Decimal?
        Get
            Return Me._UserForceShearService
        End Get
        Set
            Me._UserForceShearService = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceCaAcService")>
    <DataMember()> Public Property UserForceCaAcService() As Decimal?
        Get
            Return Me._UserForceCaAcService
        End Get
        Set
            Me._UserForceCaAcService = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceEhx")>
    <DataMember()> Public Property UserForceEhx() As Decimal?
        Get
            Return Me._UserForceEhx
        End Get
        Set
            Me._UserForceEhx = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceEhz")>
    <DataMember()> Public Property UserForceEhz() As Decimal?
        Get
            Return Me._UserForceEhz
        End Get
        Set
            Me._UserForceEhz = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceEv")>
    <DataMember()> Public Property UserForceEv() As Decimal?
        Get
            Return Me._UserForceEv
        End Get
        Set
            Me._UserForceEv = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceEh")>
    <DataMember()> Public Property UserForceEh() As Decimal?
        Get
            Return Me._UserForceEh
        End Get
        Set
            Me._UserForceEh = Value
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
        Dim otherToCompare As tnxUserForce = TryCast(other, tnxUserForce)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.UserForceRec.CheckChange(otherToCompare.UserForceRec, changes, categoryName, "Userforcerec"), Equals, False)
        Equals = If(Me.UserForceEnabled.CheckChange(otherToCompare.UserForceEnabled, changes, categoryName, "Userforceenabled"), Equals, False)
        Equals = If(Me.UserForceDescription.CheckChange(otherToCompare.UserForceDescription, changes, categoryName, "Userforcedescription"), Equals, False)
        Equals = If(Me.UserForceStartHt.CheckChange(otherToCompare.UserForceStartHt, changes, categoryName, "Userforcestartht"), Equals, False)
        Equals = If(Me.UserForceOffset.CheckChange(otherToCompare.UserForceOffset, changes, categoryName, "Userforceoffset"), Equals, False)
        Equals = If(Me.UserForceAzimuth.CheckChange(otherToCompare.UserForceAzimuth, changes, categoryName, "Userforceazimuth"), Equals, False)
        Equals = If(Me.UserForceFxNoIce.CheckChange(otherToCompare.UserForceFxNoIce, changes, categoryName, "Userforcefxnoice"), Equals, False)
        Equals = If(Me.UserForceFzNoIce.CheckChange(otherToCompare.UserForceFzNoIce, changes, categoryName, "Userforcefznoice"), Equals, False)
        Equals = If(Me.UserForceAxialNoIce.CheckChange(otherToCompare.UserForceAxialNoIce, changes, categoryName, "Userforceaxialnoice"), Equals, False)
        Equals = If(Me.UserForceShearNoIce.CheckChange(otherToCompare.UserForceShearNoIce, changes, categoryName, "Userforceshearnoice"), Equals, False)
        Equals = If(Me.UserForceCaAcNoIce.CheckChange(otherToCompare.UserForceCaAcNoIce, changes, categoryName, "Userforcecaacnoice"), Equals, False)
        Equals = If(Me.UserForceFxIce.CheckChange(otherToCompare.UserForceFxIce, changes, categoryName, "Userforcefxice"), Equals, False)
        Equals = If(Me.UserForceFzIce.CheckChange(otherToCompare.UserForceFzIce, changes, categoryName, "Userforcefzice"), Equals, False)
        Equals = If(Me.UserForceAxialIce.CheckChange(otherToCompare.UserForceAxialIce, changes, categoryName, "Userforceaxialice"), Equals, False)
        Equals = If(Me.UserForceShearIce.CheckChange(otherToCompare.UserForceShearIce, changes, categoryName, "Userforceshearice"), Equals, False)
        Equals = If(Me.UserForceCaAcIce.CheckChange(otherToCompare.UserForceCaAcIce, changes, categoryName, "Userforcecaacice"), Equals, False)
        Equals = If(Me.UserForceFxService.CheckChange(otherToCompare.UserForceFxService, changes, categoryName, "Userforcefxservice"), Equals, False)
        Equals = If(Me.UserForceFzService.CheckChange(otherToCompare.UserForceFzService, changes, categoryName, "Userforcefzservice"), Equals, False)
        Equals = If(Me.UserForceAxialService.CheckChange(otherToCompare.UserForceAxialService, changes, categoryName, "Userforceaxialservice"), Equals, False)
        Equals = If(Me.UserForceShearService.CheckChange(otherToCompare.UserForceShearService, changes, categoryName, "Userforceshearservice"), Equals, False)
        Equals = If(Me.UserForceCaAcService.CheckChange(otherToCompare.UserForceCaAcService, changes, categoryName, "Userforcecaacservice"), Equals, False)
        Equals = If(Me.UserForceEhx.CheckChange(otherToCompare.UserForceEhx, changes, categoryName, "Userforceehx"), Equals, False)
        Equals = If(Me.UserForceEhz.CheckChange(otherToCompare.UserForceEhz, changes, categoryName, "Userforceehz"), Equals, False)
        Equals = If(Me.UserForceEv.CheckChange(otherToCompare.UserForceEv, changes, categoryName, "Userforceev"), Equals, False)
        Equals = If(Me.UserForceEh.CheckChange(otherToCompare.UserForceEh, changes, categoryName, "Userforceeh"), Equals, False)

        Return Equals
    End Function
#End Region

    Public Sub New(ByVal recRow As DataRow, Optional ByVal Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Me.ID = DBtoNullableInt(recRow.Item("ID"))
        'Me.bus_unit = DBtoNullableInt(recRow.Item("bus_unit"))
        'Me.structure_id = DBtoStr(recRow.Item("structure_id"))
        'Me.modified_person_id = DBtoNullableInt(recRow.Item("modified_person_id"))
        'Me.process_stage = DBtoStr(recRow.Item("process_stage"))
        'Me.work_order_seq_num = DBtoNullableInt(recRow.Item("work_order_seq_num"))
        Me.tnx_id = DBtoNullableInt(recRow.Item("tnx_id"))
        Me.UserForceRec = DBtoNullableInt(recRow.Item("UserForceRec"))
        Me.UserForceEnabled = DBtoNullableBool(recRow.Item("UserForceEnabled"))
        Me.UserForceDescription = DBtoStr(recRow.Item("UserForceDescription"))
        Me.UserForceStartHt = DBtoNullableDec(recRow.Item("UserForceStartHt"), 6)
        Me.UserForceOffset = DBtoNullableDec(recRow.Item("UserForceOffset"), 6)
        Me.UserForceAzimuth = DBtoNullableDec(recRow.Item("UserForceAzimuth"), 6)
        Me.UserForceFxNoIce = DBtoNullableDec(recRow.Item("UserForceFxNoIce"), 6)
        Me.UserForceFzNoIce = DBtoNullableDec(recRow.Item("UserForceFzNoIce"), 6)
        Me.UserForceAxialNoIce = DBtoNullableDec(recRow.Item("UserForceAxialNoIce"), 6)
        Me.UserForceShearNoIce = DBtoNullableDec(recRow.Item("UserForceShearNoIce"), 6)
        Me.UserForceCaAcNoIce = DBtoNullableDec(recRow.Item("UserForceCaAcNoIce"), 6)
        Me.UserForceFxIce = DBtoNullableDec(recRow.Item("UserForceFxIce"), 6)
        Me.UserForceFzIce = DBtoNullableDec(recRow.Item("UserForceFzIce"), 6)
        Me.UserForceAxialIce = DBtoNullableDec(recRow.Item("UserForceAxialIce"), 6)
        Me.UserForceShearIce = DBtoNullableDec(recRow.Item("UserForceShearIce"), 6)
        Me.UserForceCaAcIce = DBtoNullableDec(recRow.Item("UserForceCaAcIce"), 6)
        Me.UserForceFxService = DBtoNullableDec(recRow.Item("UserForceFxService"), 6)
        Me.UserForceFzService = DBtoNullableDec(recRow.Item("UserForceFzService"), 6)
        Me.UserForceAxialService = DBtoNullableDec(recRow.Item("UserForceAxialService"), 6)
        Me.UserForceShearService = DBtoNullableDec(recRow.Item("UserForceShearService"), 6)
        Me.UserForceCaAcService = DBtoNullableDec(recRow.Item("UserForceCaAcService"), 6)
        Me.UserForceEhx = DBtoNullableDec(recRow.Item("UserForceEhx"), 6)
        Me.UserForceEhz = DBtoNullableDec(recRow.Item("UserForceEhz"), 6)
        Me.UserForceEv = DBtoNullableDec(recRow.Item("UserForceEv"), 6)
        Me.UserForceEh = DBtoNullableDec(recRow.Item("UserForceEh"), 6)

    End Sub

    Public Overrides Function SQLInsert() As String
        SQLInsert = CCI_Engineering_Templates.My.Resources.General__INSERT
        SQLInsert = SQLInsert.Replace("[TABLE NAME]", Me.EDSTableName)
        SQLInsert = SQLInsert.Replace("[INSERT FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.Replace("[INSERT VALUES]", Me.SQLInsertValues)

        Return SQLInsert
    End Function

    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""



        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bus_unit.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.structure_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.work_order_seq_num.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@TopLevelID") 'foreign key reference 'tnx_id
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.UserForceRec.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.UserForceEnabled.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.UserForceDescription.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.UserForceStartHt.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.UserForceOffset.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.UserForceAzimuth.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.UserForceFxNoIce.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.UserForceFzNoIce.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.UserForceAxialNoIce.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.UserForceShearNoIce.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.UserForceCaAcNoIce.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.UserForceFxIce.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.UserForceFzIce.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.UserForceAxialIce.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.UserForceShearIce.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.UserForceCaAcIce.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.UserForceFxService.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.UserForceFzService.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.UserForceAxialService.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.UserForceShearService.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.UserForceCaAcService.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.UserForceEhx.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.UserForceEhz.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.UserForceEv.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.UserForceEh.ToString.FormatDBValue)

        Return SQLInsertValues
        'Throw New NotImplementedException()
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        SQLInsertFields = SQLInsertFields.AddtoDBString("bus_unit")
        SQLInsertFields = SQLInsertFields.AddtoDBString("structure_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")
        SQLInsertFields = SQLInsertFields.AddtoDBString("work_order_seq_num")
        SQLInsertFields = SQLInsertFields.AddtoDBString("tnx_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UserForceRec")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UserForceEnabled")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UserForceDescription")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UserForceStartHt")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UserForceOffset")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UserForceAzimuth")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UserForceFxNoIce")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UserForceFzNoIce")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UserForceAxialNoIce")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UserForceShearNoIce")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UserForceCaAcNoIce")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UserForceFxIce")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UserForceFzIce")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UserForceAxialIce")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UserForceShearIce")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UserForceCaAcIce")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UserForceFxService")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UserForceFzService")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UserForceAxialService")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UserForceShearService")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UserForceCaAcService")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UserForceEhx")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UserForceEhz")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UserForceEv")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UserForceEh")

        Return SQLInsertFields
        'Throw New NotImplementedException()
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        Throw New NotImplementedException()
    End Function
End Class
