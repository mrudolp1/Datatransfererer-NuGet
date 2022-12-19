﻿Option Strict On
Option Compare Binary 'Trying to speed up parsing the TNX file by using Binary Text comparison instead of Text Comparison

Imports System.ComponentModel
Imports System.Data
Imports System.IO
Imports System.Security.Principal
Imports System.Runtime.CompilerServices
Imports System.Data.SqlClient


Partial Public Class tnxFeedLine
    Inherits EDSObject

#Region "Inheritted"

    Public Overrides ReadOnly Property EDSObjectName As String = "Feed Line"

#End Region

#Region "Define"
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
    Private _FeedLineRoundClusterDia As Double?
    Private _FeedLineWidth As Double?
    Private _FeedLinePerimeter As Double?
    Private _FlatAttachmentEffectiveWidthRatio As Double?
    Private _AutoCalcFlatAttachmentEffectiveWidthRatio As Boolean?
    Private _FeedLineShieldingFactorKaNoIce As Double?
    Private _FeedLineShieldingFactorKaIce As Double?
    Private _FeedLineAutoCalcKa As Boolean?
    Private _FeedLineCaAaNoIce As Double?
    Private _FeedLineCaAaIce As Double?
    Private _FeedLineCaAaIce_1 As Double?
    Private _FeedLineCaAaIce_2 As Double?
    Private _FeedLineCaAaIce_4 As Double?
    Private _FeedLineWtNoIce As Double?
    Private _FeedLineWtIce As Double?
    Private _FeedLineWtIce_1 As Double?
    Private _FeedLineWtIce_2 As Double?
    Private _FeedLineWtIce_4 As Double?
    Private _FeedLineFaceOffset As Double?
    Private _FeedLineOffsetFrac As Double?
    Private _FeedLinePerimeterOffsetStartFrac As Double?
    Private _FeedLinePerimeterOffsetEndFrac As Double?
    Private _FeedLineStartHt As Double?
    Private _FeedLineEndHt As Double?
    Private _FeedLineClearSpacing As Double?
    Private _FeedLineRowClearSpacing As Double?

    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineRec")>
    Public Property FeedLineRec() As Integer?
        Get
            Return Me._FeedLineRec
        End Get
        Set
            Me._FeedLineRec = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineEnabled")>
    Public Property FeedLineEnabled() As Boolean?
        Get
            Return Me._FeedLineEnabled
        End Get
        Set
            Me._FeedLineEnabled = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineDatabase")>
    Public Property FeedLineDatabase() As String
        Get
            Return Me._FeedLineDatabase
        End Get
        Set
            Me._FeedLineDatabase = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineDescription")>
    Public Property FeedLineDescription() As String
        Get
            Return Me._FeedLineDescription
        End Get
        Set
            Me._FeedLineDescription = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineClassificationCategory")>
    Public Property FeedLineClassificationCategory() As String
        Get
            Return Me._FeedLineClassificationCategory
        End Get
        Set
            Me._FeedLineClassificationCategory = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineNote")>
    Public Property FeedLineNote() As String
        Get
            Return Me._FeedLineNote
        End Get
        Set
            Me._FeedLineNote = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineNum")>
    Public Property FeedLineNum() As Integer?
        Get
            Return Me._FeedLineNum
        End Get
        Set
            Me._FeedLineNum = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineUseShielding")>
    Public Property FeedLineUseShielding() As Boolean?
        Get
            Return Me._FeedLineUseShielding
        End Get
        Set
            Me._FeedLineUseShielding = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("ExcludeFeedLineFromTorque")>
    Public Property ExcludeFeedLineFromTorque() As Boolean?
        Get
            Return Me._ExcludeFeedLineFromTorque
        End Get
        Set
            Me._ExcludeFeedLineFromTorque = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineNumPerRow")>
    Public Property FeedLineNumPerRow() As Integer?
        Get
            Return Me._FeedLineNumPerRow
        End Get
        Set
            Me._FeedLineNumPerRow = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description("{0 = A, 1 = B,  2 = C, 3 = D}"), DisplayName("FeedLineFace")>
    Public Property FeedLineFace() As Integer?
        Get
            Return Me._FeedLineFace
        End Get
        Set
            Me._FeedLineFace = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineComponentType")>
    Public Property FeedLineComponentType() As String
        Get
            Return Me._FeedLineComponentType
        End Get
        Set
            Me._FeedLineComponentType = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineGroupTreatmentType")>
    Public Property FeedLineGroupTreatmentType() As String
        Get
            Return Me._FeedLineGroupTreatmentType
        End Get
        Set
            Me._FeedLineGroupTreatmentType = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineRoundClusterDia")>
    Public Property FeedLineRoundClusterDia() As Double?
        Get
            Return Me._FeedLineRoundClusterDia
        End Get
        Set
            Me._FeedLineRoundClusterDia = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineWidth")>
    Public Property FeedLineWidth() As Double?
        Get
            Return Me._FeedLineWidth
        End Get
        Set
            Me._FeedLineWidth = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLinePerimeter")>
    Public Property FeedLinePerimeter() As Double?
        Get
            Return Me._FeedLinePerimeter
        End Get
        Set
            Me._FeedLinePerimeter = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FlatAttachmentEffectiveWidthRatio")>
    Public Property FlatAttachmentEffectiveWidthRatio() As Double?
        Get
            Return Me._FlatAttachmentEffectiveWidthRatio
        End Get
        Set
            Me._FlatAttachmentEffectiveWidthRatio = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("AutoCalcFlatAttachmentEffectiveWidthRatio")>
    Public Property AutoCalcFlatAttachmentEffectiveWidthRatio() As Boolean?
        Get
            Return Me._AutoCalcFlatAttachmentEffectiveWidthRatio
        End Get
        Set
            Me._AutoCalcFlatAttachmentEffectiveWidthRatio = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineShieldingFactorKaNoIce")>
    Public Property FeedLineShieldingFactorKaNoIce() As Double?
        Get
            Return Me._FeedLineShieldingFactorKaNoIce
        End Get
        Set
            Me._FeedLineShieldingFactorKaNoIce = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineShieldingFactorKaIce")>
    Public Property FeedLineShieldingFactorKaIce() As Double?
        Get
            Return Me._FeedLineShieldingFactorKaIce
        End Get
        Set
            Me._FeedLineShieldingFactorKaIce = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineAutoCalcKa")>
    Public Property FeedLineAutoCalcKa() As Boolean?
        Get
            Return Me._FeedLineAutoCalcKa
        End Get
        Set
            Me._FeedLineAutoCalcKa = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineCaAaNoIce")>
    Public Property FeedLineCaAaNoIce() As Double?
        Get
            Return Me._FeedLineCaAaNoIce
        End Get
        Set
            Me._FeedLineCaAaNoIce = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineCaAaIce")>
    Public Property FeedLineCaAaIce() As Double?
        Get
            Return Me._FeedLineCaAaIce
        End Get
        Set
            Me._FeedLineCaAaIce = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineCaAaIce_1")>
    Public Property FeedLineCaAaIce_1() As Double?
        Get
            Return Me._FeedLineCaAaIce_1
        End Get
        Set
            Me._FeedLineCaAaIce_1 = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineCaAaIce_2")>
    Public Property FeedLineCaAaIce_2() As Double?
        Get
            Return Me._FeedLineCaAaIce_2
        End Get
        Set
            Me._FeedLineCaAaIce_2 = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineCaAaIce_4")>
    Public Property FeedLineCaAaIce_4() As Double?
        Get
            Return Me._FeedLineCaAaIce_4
        End Get
        Set
            Me._FeedLineCaAaIce_4 = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineWtNoIce")>
    Public Property FeedLineWtNoIce() As Double?
        Get
            Return Me._FeedLineWtNoIce
        End Get
        Set
            Me._FeedLineWtNoIce = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineWtIce")>
    Public Property FeedLineWtIce() As Double?
        Get
            Return Me._FeedLineWtIce
        End Get
        Set
            Me._FeedLineWtIce = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineWtIce_1")>
    Public Property FeedLineWtIce_1() As Double?
        Get
            Return Me._FeedLineWtIce_1
        End Get
        Set
            Me._FeedLineWtIce_1 = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineWtIce_2")>
    Public Property FeedLineWtIce_2() As Double?
        Get
            Return Me._FeedLineWtIce_2
        End Get
        Set
            Me._FeedLineWtIce_2 = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineWtIce_4")>
    Public Property FeedLineWtIce_4() As Double?
        Get
            Return Me._FeedLineWtIce_4
        End Get
        Set
            Me._FeedLineWtIce_4 = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineFaceOffset")>
    Public Property FeedLineFaceOffset() As Double?
        Get
            Return Me._FeedLineFaceOffset
        End Get
        Set
            Me._FeedLineFaceOffset = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineOffsetFrac")>
    Public Property FeedLineOffsetFrac() As Double?
        Get
            Return Me._FeedLineOffsetFrac
        End Get
        Set
            Me._FeedLineOffsetFrac = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLinePerimeterOffsetStartFrac")>
    Public Property FeedLinePerimeterOffsetStartFrac() As Double?
        Get
            Return Me._FeedLinePerimeterOffsetStartFrac
        End Get
        Set
            Me._FeedLinePerimeterOffsetStartFrac = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLinePerimeterOffsetEndFrac")>
    Public Property FeedLinePerimeterOffsetEndFrac() As Double?
        Get
            Return Me._FeedLinePerimeterOffsetEndFrac
        End Get
        Set
            Me._FeedLinePerimeterOffsetEndFrac = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineStartHt")>
    Public Property FeedLineStartHt() As Double?
        Get
            Return Me._FeedLineStartHt
        End Get
        Set
            Me._FeedLineStartHt = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineEndHt")>
    Public Property FeedLineEndHt() As Double?
        Get
            Return Me._FeedLineEndHt
        End Get
        Set
            Me._FeedLineEndHt = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineClearSpacing")>
    Public Property FeedLineClearSpacing() As Double?
        Get
            Return Me._FeedLineClearSpacing
        End Get
        Set
            Me._FeedLineClearSpacing = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineRowClearSpacing")>
    Public Property FeedLineRowClearSpacing() As Double?
        Get
            Return Me._FeedLineRowClearSpacing
        End Get
        Set
            Me._FeedLineRowClearSpacing = Value
        End Set
    End Property
#End Region

#Region "Constructors"
    Public Sub New(Optional ByRef Parent As EDSObject = Nothing)
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

End Class
Partial Public Class tnxDiscreteLoad
    Inherits EDSObject

#Region "Inheritted"
    Public Overrides ReadOnly Property EDSObjectName As String = "Discrete Load"
#End Region

#Region "Define"
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
    Private _TowerOffsetDist As Double?
    Private _TowerVertOffset As Double?
    Private _TowerLateralOffset As Double?
    Private _TowerAzimuthAdjustment As Double?
    Private _TowerAppurtSymbol As String
    Private _TowerLoadShieldingFactorKaNoIce As Double?
    Private _TowerLoadShieldingFactorKaIce As Double?
    Private _TowerLoadAutoCalcKa As Boolean?
    Private _TowerLoadCaAaNoIce As Double?
    Private _TowerLoadCaAaIce As Double?
    Private _TowerLoadCaAaIce_1 As Double?
    Private _TowerLoadCaAaIce_2 As Double?
    Private _TowerLoadCaAaIce_4 As Double?
    Private _TowerLoadCaAaNoIce_Side As Double?
    Private _TowerLoadCaAaIce_Side As Double?
    Private _TowerLoadCaAaIce_Side_1 As Double?
    Private _TowerLoadCaAaIce_Side_2 As Double?
    Private _TowerLoadCaAaIce_Side_4 As Double?
    Private _TowerLoadWtNoIce As Double?
    Private _TowerLoadWtIce As Double?
    Private _TowerLoadWtIce_1 As Double?
    Private _TowerLoadWtIce_2 As Double?
    Private _TowerLoadWtIce_4 As Double?
    Private _TowerLoadStartHt As Double?
    Private _TowerLoadEndHt As Double?


    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadRec")>
    Public Property TowerLoadRec() As Integer?
        Get
            Return Me._TowerLoadRec
        End Get
        Set
            Me._TowerLoadRec = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadEnabled")>
    Public Property TowerLoadEnabled() As Boolean?
        Get
            Return Me._TowerLoadEnabled
        End Get
        Set
            Me._TowerLoadEnabled = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadDatabase")>
    Public Property TowerLoadDatabase() As String
        Get
            Return Me._TowerLoadDatabase
        End Get
        Set
            Me._TowerLoadDatabase = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadDescription")>
    Public Property TowerLoadDescription() As String
        Get
            Return Me._TowerLoadDescription
        End Get
        Set
            Me._TowerLoadDescription = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadType")>
    Public Property TowerLoadType() As String
        Get
            Return Me._TowerLoadType
        End Get
        Set
            Me._TowerLoadType = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadClassificationCategory")>
    Public Property TowerLoadClassificationCategory() As String
        Get
            Return Me._TowerLoadClassificationCategory
        End Get
        Set
            Me._TowerLoadClassificationCategory = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadNote")>
    Public Property TowerLoadNote() As String
        Get
            Return Me._TowerLoadNote
        End Get
        Set
            Me._TowerLoadNote = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadNum")>
    Public Property TowerLoadNum() As Integer?
        Get
            Return Me._TowerLoadNum
        End Get
        Set
            Me._TowerLoadNum = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description("{0 = A, 1 = B,  2 = C, 3 = D}"), DisplayName("TowerLoadFace")>
    Public Property TowerLoadFace() As Integer?
        Get
            Return Me._TowerLoadFace
        End Get
        Set
            Me._TowerLoadFace = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerOffsetType")>
    Public Property TowerOffsetType() As String
        Get
            Return Me._TowerOffsetType
        End Get
        Set
            Me._TowerOffsetType = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerOffsetDist")>
    Public Property TowerOffsetDist() As Double?
        Get
            Return Me._TowerOffsetDist
        End Get
        Set
            Me._TowerOffsetDist = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerVertOffset")>
    Public Property TowerVertOffset() As Double?
        Get
            Return Me._TowerVertOffset
        End Get
        Set
            Me._TowerVertOffset = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLateralOffset")>
    Public Property TowerLateralOffset() As Double?
        Get
            Return Me._TowerLateralOffset
        End Get
        Set
            Me._TowerLateralOffset = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerAzimuthAdjustment")>
    Public Property TowerAzimuthAdjustment() As Double?
        Get
            Return Me._TowerAzimuthAdjustment
        End Get
        Set
            Me._TowerAzimuthAdjustment = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerAppurtSymbol")>
    Public Property TowerAppurtSymbol() As String
        Get
            Return Me._TowerAppurtSymbol
        End Get
        Set
            Me._TowerAppurtSymbol = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadShieldingFactorKaNoIce")>
    Public Property TowerLoadShieldingFactorKaNoIce() As Double?
        Get
            Return Me._TowerLoadShieldingFactorKaNoIce
        End Get
        Set
            Me._TowerLoadShieldingFactorKaNoIce = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadShieldingFactorKaIce")>
    Public Property TowerLoadShieldingFactorKaIce() As Double?
        Get
            Return Me._TowerLoadShieldingFactorKaIce
        End Get
        Set
            Me._TowerLoadShieldingFactorKaIce = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadAutoCalcKa")>
    Public Property TowerLoadAutoCalcKa() As Boolean?
        Get
            Return Me._TowerLoadAutoCalcKa
        End Get
        Set
            Me._TowerLoadAutoCalcKa = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadCaAaNoIce")>
    Public Property TowerLoadCaAaNoIce() As Double?
        Get
            Return Me._TowerLoadCaAaNoIce
        End Get
        Set
            Me._TowerLoadCaAaNoIce = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadCaAaIce")>
    Public Property TowerLoadCaAaIce() As Double?
        Get
            Return Me._TowerLoadCaAaIce
        End Get
        Set
            Me._TowerLoadCaAaIce = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadCaAaIce_1")>
    Public Property TowerLoadCaAaIce_1() As Double?
        Get
            Return Me._TowerLoadCaAaIce_1
        End Get
        Set
            Me._TowerLoadCaAaIce_1 = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadCaAaIce_2")>
    Public Property TowerLoadCaAaIce_2() As Double?
        Get
            Return Me._TowerLoadCaAaIce_2
        End Get
        Set
            Me._TowerLoadCaAaIce_2 = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadCaAaIce_4")>
    Public Property TowerLoadCaAaIce_4() As Double?
        Get
            Return Me._TowerLoadCaAaIce_4
        End Get
        Set
            Me._TowerLoadCaAaIce_4 = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadCaAaNoIce_Side")>
    Public Property TowerLoadCaAaNoIce_Side() As Double?
        Get
            Return Me._TowerLoadCaAaNoIce_Side
        End Get
        Set
            Me._TowerLoadCaAaNoIce_Side = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadCaAaIce_Side")>
    Public Property TowerLoadCaAaIce_Side() As Double?
        Get
            Return Me._TowerLoadCaAaIce_Side
        End Get
        Set
            Me._TowerLoadCaAaIce_Side = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadCaAaIce_Side_1")>
    Public Property TowerLoadCaAaIce_Side_1() As Double?
        Get
            Return Me._TowerLoadCaAaIce_Side_1
        End Get
        Set
            Me._TowerLoadCaAaIce_Side_1 = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadCaAaIce_Side_2")>
    Public Property TowerLoadCaAaIce_Side_2() As Double?
        Get
            Return Me._TowerLoadCaAaIce_Side_2
        End Get
        Set
            Me._TowerLoadCaAaIce_Side_2 = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadCaAaIce_Side_4")>
    Public Property TowerLoadCaAaIce_Side_4() As Double?
        Get
            Return Me._TowerLoadCaAaIce_Side_4
        End Get
        Set
            Me._TowerLoadCaAaIce_Side_4 = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadWtNoIce")>
    Public Property TowerLoadWtNoIce() As Double?
        Get
            Return Me._TowerLoadWtNoIce
        End Get
        Set
            Me._TowerLoadWtNoIce = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadWtIce")>
    Public Property TowerLoadWtIce() As Double?
        Get
            Return Me._TowerLoadWtIce
        End Get
        Set
            Me._TowerLoadWtIce = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadWtIce_1")>
    Public Property TowerLoadWtIce_1() As Double?
        Get
            Return Me._TowerLoadWtIce_1
        End Get
        Set
            Me._TowerLoadWtIce_1 = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadWtIce_2")>
    Public Property TowerLoadWtIce_2() As Double?
        Get
            Return Me._TowerLoadWtIce_2
        End Get
        Set
            Me._TowerLoadWtIce_2 = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadWtIce_4")>
    Public Property TowerLoadWtIce_4() As Double?
        Get
            Return Me._TowerLoadWtIce_4
        End Get
        Set
            Me._TowerLoadWtIce_4 = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadStartHt")>
    Public Property TowerLoadStartHt() As Double?
        Get
            Return Me._TowerLoadStartHt
        End Get
        Set
            Me._TowerLoadStartHt = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadEndHt")>
    Public Property TowerLoadEndHt() As Double?
        Get
            Return Me._TowerLoadEndHt
        End Get
        Set
            Me._TowerLoadEndHt = Value
        End Set
    End Property
#End Region

#Region "Constructors"
    Public Sub New(Optional ByRef Parent As EDSObject = Nothing)
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

End Class

Partial Public Class tnxDish
    Inherits EDSObject

#Region "Inheritted"
    Public Overrides ReadOnly Property EDSObjectName As String = "Dish"
#End Region

#Region "Define"
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
    Private _DishVertOffset As Double?
    Private _DishLateralOffset As Double?
    Private _DishOffsetDist As Double?
    Private _DishArea As Double?
    Private _DishAreaIce As Double?
    Private _DishAreaIce_1 As Double?
    Private _DishAreaIce_2 As Double?
    Private _DishAreaIce_4 As Double?
    Private _DishDiameter As Double?
    Private _DishWtNoIce As Double?
    Private _DishWtIce As Double?
    Private _DishWtIce_1 As Double?
    Private _DishWtIce_2 As Double?
    Private _DishWtIce_4 As Double?
    Private _DishStartHt As Double?
    Private _DishAzimuthAdjustment As Double?
    Private _DishBeamWidth As Double?

    <Category("TNX Dish"), Description(""), DisplayName("DishRec")>
    Public Property DishRec() As Integer?
        Get
            Return Me._DishRec
        End Get
        Set
            Me._DishRec = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishEnabled")>
    Public Property DishEnabled() As Boolean?
        Get
            Return Me._DishEnabled
        End Get
        Set
            Me._DishEnabled = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishDatabase")>
    Public Property DishDatabase() As String
        Get
            Return Me._DishDatabase
        End Get
        Set
            Me._DishDatabase = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishDescription")>
    Public Property DishDescription() As String
        Get
            Return Me._DishDescription
        End Get
        Set
            Me._DishDescription = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishClassificationCategory")>
    Public Property DishClassificationCategory() As String
        Get
            Return Me._DishClassificationCategory
        End Get
        Set
            Me._DishClassificationCategory = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishNote")>
    Public Property DishNote() As String
        Get
            Return Me._DishNote
        End Get
        Set
            Me._DishNote = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishNum")>
    Public Property DishNum() As Integer?
        Get
            Return Me._DishNum
        End Get
        Set
            Me._DishNum = Value
        End Set
    End Property
    <Category("TNX Dish"), Description("{0 = A, 1 = B,  2 = C, 3 = D}"), DisplayName("DishFace")>
    Public Property DishFace() As Integer?
        Get
            Return Me._DishFace
        End Get
        Set
            Me._DishFace = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishType")>
    Public Property DishType() As String
        Get
            Return Me._DishType
        End Get
        Set
            Me._DishType = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishOffsetType")>
    Public Property DishOffsetType() As String
        Get
            Return Me._DishOffsetType
        End Get
        Set
            Me._DishOffsetType = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishVertOffset")>
    Public Property DishVertOffset() As Double?
        Get
            Return Me._DishVertOffset
        End Get
        Set
            Me._DishVertOffset = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishLateralOffset")>
    Public Property DishLateralOffset() As Double?
        Get
            Return Me._DishLateralOffset
        End Get
        Set
            Me._DishLateralOffset = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishOffsetDist")>
    Public Property DishOffsetDist() As Double?
        Get
            Return Me._DishOffsetDist
        End Get
        Set
            Me._DishOffsetDist = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishArea")>
    Public Property DishArea() As Double?
        Get
            Return Me._DishArea
        End Get
        Set
            Me._DishArea = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishAreaIce")>
    Public Property DishAreaIce() As Double?
        Get
            Return Me._DishAreaIce
        End Get
        Set
            Me._DishAreaIce = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishAreaIce_1")>
    Public Property DishAreaIce_1() As Double?
        Get
            Return Me._DishAreaIce_1
        End Get
        Set
            Me._DishAreaIce_1 = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishAreaIce_2")>
    Public Property DishAreaIce_2() As Double?
        Get
            Return Me._DishAreaIce_2
        End Get
        Set
            Me._DishAreaIce_2 = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishAreaIce_4")>
    Public Property DishAreaIce_4() As Double?
        Get
            Return Me._DishAreaIce_4
        End Get
        Set
            Me._DishAreaIce_4 = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishDiameter")>
    Public Property DishDiameter() As Double?
        Get
            Return Me._DishDiameter
        End Get
        Set
            Me._DishDiameter = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishWtNoIce")>
    Public Property DishWtNoIce() As Double?
        Get
            Return Me._DishWtNoIce
        End Get
        Set
            Me._DishWtNoIce = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishWtIce")>
    Public Property DishWtIce() As Double?
        Get
            Return Me._DishWtIce
        End Get
        Set
            Me._DishWtIce = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishWtIce_1")>
    Public Property DishWtIce_1() As Double?
        Get
            Return Me._DishWtIce_1
        End Get
        Set
            Me._DishWtIce_1 = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishWtIce_2")>
    Public Property DishWtIce_2() As Double?
        Get
            Return Me._DishWtIce_2
        End Get
        Set
            Me._DishWtIce_2 = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishWtIce_4")>
    Public Property DishWtIce_4() As Double?
        Get
            Return Me._DishWtIce_4
        End Get
        Set
            Me._DishWtIce_4 = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishStartHt")>
    Public Property DishStartHt() As Double?
        Get
            Return Me._DishStartHt
        End Get
        Set
            Me._DishStartHt = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishAzimuthAdjustment")>
    Public Property DishAzimuthAdjustment() As Double?
        Get
            Return Me._DishAzimuthAdjustment
        End Get
        Set
            Me._DishAzimuthAdjustment = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishBeamWidth")>
    Public Property DishBeamWidth() As Double?
        Get
            Return Me._DishBeamWidth
        End Get
        Set
            Me._DishBeamWidth = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New(Optional ByRef Parent As EDSObject = Nothing)
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
End Class


Partial Public Class tnxUserForce
    Inherits EDSObject

#Region "Inheritted"
    Public Overrides ReadOnly Property EDSObjectName As String = "User Force"
#End Region

#Region "Define"
    Private _UserForceRec As Integer?
    Private _UserForceEnabled As Boolean?
    Private _UserForceDescription As String
    Private _UserForceStartHt As Double?
    Private _UserForceOffset As Double?
    Private _UserForceAzimuth As Double?
    Private _UserForceFxNoIce As Double?
    Private _UserForceFzNoIce As Double?
    Private _UserForceAxialNoIce As Double?
    Private _UserForceShearNoIce As Double?
    Private _UserForceCaAcNoIce As Double?
    Private _UserForceFxIce As Double?
    Private _UserForceFzIce As Double?
    Private _UserForceAxialIce As Double?
    Private _UserForceShearIce As Double?
    Private _UserForceCaAcIce As Double?
    Private _UserForceFxService As Double?
    Private _UserForceFzService As Double?
    Private _UserForceAxialService As Double?
    Private _UserForceShearService As Double?
    Private _UserForceCaAcService As Double?
    Private _UserForceEhx As Double?
    Private _UserForceEhz As Double?
    Private _UserForceEv As Double?
    Private _UserForceEh As Double?

    <Category("TNX User Force"), Description(""), DisplayName("UserForceRec")>
    Public Property UserForceRec() As Integer?
        Get
            Return Me._UserForceRec
        End Get
        Set
            Me._UserForceRec = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceEnabled")>
    Public Property UserForceEnabled() As Boolean?
        Get
            Return Me._UserForceEnabled
        End Get
        Set
            Me._UserForceEnabled = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceDescription")>
    Public Property UserForceDescription() As String
        Get
            Return Me._UserForceDescription
        End Get
        Set
            Me._UserForceDescription = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceStartHt")>
    Public Property UserForceStartHt() As Double?
        Get
            Return Me._UserForceStartHt
        End Get
        Set
            Me._UserForceStartHt = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceOffset")>
    Public Property UserForceOffset() As Double?
        Get
            Return Me._UserForceOffset
        End Get
        Set
            Me._UserForceOffset = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceAzimuth")>
    Public Property UserForceAzimuth() As Double?
        Get
            Return Me._UserForceAzimuth
        End Get
        Set
            Me._UserForceAzimuth = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceFxNoIce")>
    Public Property UserForceFxNoIce() As Double?
        Get
            Return Me._UserForceFxNoIce
        End Get
        Set
            Me._UserForceFxNoIce = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceFzNoIce")>
    Public Property UserForceFzNoIce() As Double?
        Get
            Return Me._UserForceFzNoIce
        End Get
        Set
            Me._UserForceFzNoIce = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceAxialNoIce")>
    Public Property UserForceAxialNoIce() As Double?
        Get
            Return Me._UserForceAxialNoIce
        End Get
        Set
            Me._UserForceAxialNoIce = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceShearNoIce")>
    Public Property UserForceShearNoIce() As Double?
        Get
            Return Me._UserForceShearNoIce
        End Get
        Set
            Me._UserForceShearNoIce = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceCaAcNoIce")>
    Public Property UserForceCaAcNoIce() As Double?
        Get
            Return Me._UserForceCaAcNoIce
        End Get
        Set
            Me._UserForceCaAcNoIce = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceFxIce")>
    Public Property UserForceFxIce() As Double?
        Get
            Return Me._UserForceFxIce
        End Get
        Set
            Me._UserForceFxIce = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceFzIce")>
    Public Property UserForceFzIce() As Double?
        Get
            Return Me._UserForceFzIce
        End Get
        Set
            Me._UserForceFzIce = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceAxialIce")>
    Public Property UserForceAxialIce() As Double?
        Get
            Return Me._UserForceAxialIce
        End Get
        Set
            Me._UserForceAxialIce = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceShearIce")>
    Public Property UserForceShearIce() As Double?
        Get
            Return Me._UserForceShearIce
        End Get
        Set
            Me._UserForceShearIce = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceCaAcIce")>
    Public Property UserForceCaAcIce() As Double?
        Get
            Return Me._UserForceCaAcIce
        End Get
        Set
            Me._UserForceCaAcIce = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceFxService")>
    Public Property UserForceFxService() As Double?
        Get
            Return Me._UserForceFxService
        End Get
        Set
            Me._UserForceFxService = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceFzService")>
    Public Property UserForceFzService() As Double?
        Get
            Return Me._UserForceFzService
        End Get
        Set
            Me._UserForceFzService = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceAxialService")>
    Public Property UserForceAxialService() As Double?
        Get
            Return Me._UserForceAxialService
        End Get
        Set
            Me._UserForceAxialService = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceShearService")>
    Public Property UserForceShearService() As Double?
        Get
            Return Me._UserForceShearService
        End Get
        Set
            Me._UserForceShearService = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceCaAcService")>
    Public Property UserForceCaAcService() As Double?
        Get
            Return Me._UserForceCaAcService
        End Get
        Set
            Me._UserForceCaAcService = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceEhx")>
    Public Property UserForceEhx() As Double?
        Get
            Return Me._UserForceEhx
        End Get
        Set
            Me._UserForceEhx = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceEhz")>
    Public Property UserForceEhz() As Double?
        Get
            Return Me._UserForceEhz
        End Get
        Set
            Me._UserForceEhz = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceEv")>
    Public Property UserForceEv() As Double?
        Get
            Return Me._UserForceEv
        End Get
        Set
            Me._UserForceEv = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceEh")>
    Public Property UserForceEh() As Double?
        Get
            Return Me._UserForceEh
        End Get
        Set
            Me._UserForceEh = Value
        End Set
    End Property
#End Region

#Region "Constructors"
    Public Sub New(Optional ByRef Parent As EDSObject = Nothing)
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

End Class