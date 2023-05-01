Option Strict On
Option Compare Binary 'Trying to speed up parsing the TNX file by using Binary Text comparison instead of Text Comparison

Imports System.ComponentModel
Imports System.Data
Imports System.IO
Imports System.Security.Principal
Imports System.Runtime.CompilerServices
Imports System.Data.SqlClient
Imports System.Runtime.Serialization

#Region "Code"
<DataContract()>
Partial Public Class tnxCode
    Inherits EDSObject

    Public Overrides ReadOnly Property EDSObjectName As String = "Code"

#Region "Define"
    Private _design As New tnxDesign()
    Private _ice As New tnxIce()
    Private _thermal As New tnxThermal()
    Private _wind As New tnxWind()
    Private _misclCode As New tnxMisclCode()
    Private _seismic As New tnxSeismic()

    <Category("TNX Code"), Description(""), DisplayName("Design")>
     <DataMember()> Public Property design() As tnxDesign
        Get
            Return Me._design
        End Get
        Set
            Me._design = Value
        End Set
    End Property

    <Category("TNX Code"), Description(""), DisplayName("Ice")>
     <DataMember()> Public Property ice() As tnxIce
        Get
            Return Me._ice
        End Get
        Set
            Me._ice = Value
        End Set
    End Property

    <Category("TNX Code"), Description(""), DisplayName("Thermal")>
     <DataMember()> Public Property thermal() As tnxThermal
        Get
            Return Me._thermal
        End Get
        Set
            Me._thermal = Value
        End Set
    End Property

    <Category("TNX Code"), Description(""), DisplayName("Wind")>
     <DataMember()> Public Property wind() As tnxWind
        Get
            Return Me._wind
        End Get
        Set
            Me._wind = Value
        End Set
    End Property

    <Category("TNX Code"), Description(""), DisplayName("Miscellaneous Code")>
     <DataMember()> Public Property misclCode() As tnxMisclCode
        Get
            Return Me._misclCode
        End Get
        Set
            Me._misclCode = Value
        End Set
    End Property

    <Category("TNX Code"), Description(""), DisplayName("Seismic")>
     <DataMember()> Public Property seismic() As tnxSeismic
        Get
            Return Me._seismic
        End Get
        Set
            Me._seismic = Value
        End Set
    End Property
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
        Dim otherToCompare As tnxCode = TryCast(other, tnxCode)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.design.CheckChange(otherToCompare.design, changes, categoryName, "Member type"), Equals, False)
        Equals = If(Me.ice.CheckChange(otherToCompare.ice, changes, categoryName, "US Name"), Equals, False)
        Equals = If(Me.thermal.CheckChange(otherToCompare.thermal, changes, categoryName, "SI Name"), Equals, False)
        Equals = If(Me.wind.CheckChange(otherToCompare.wind, changes, categoryName, "Member values"), Equals, False)
        Equals = If(Me.misclCode.CheckChange(otherToCompare.misclCode, changes, categoryName, "SI Name"), Equals, False)
        Equals = If(Me.seismic.CheckChange(otherToCompare.seismic, changes, categoryName, "Member values"), Equals, False)
        Return Equals
    End Function
End Class
<DataContract()>
Partial Public Class tnxDesign
    Inherits EDSObject

    Public Overrides ReadOnly Property EDSObjectName As String = "Design"

#Region "Define"
    Private _DesignCode As String
    Private _ERIDesignMode As String
    Private _DoInteraction As Boolean?
    Private _DoHorzInteraction As Boolean?
    Private _DoDiagInteraction As Boolean?
    Private _UseMomentMagnification As Boolean?
    Private _UseCodeStressRatio As Boolean?
    Private _AllowStressRatio As Double?
    Private _AllowAntStressRatio As Double?
    Private _UseCodeGuySF As Boolean?
    Private _GuySF As Double?
    Private _UseTIA222H_AnnexS As Boolean?
    Private _TIA_222_H_AnnexS_Ratio As Double?
    Private _PrintBitmaps As Boolean?

    <Category("TNX Code Design"), Description(""), DisplayName("DesignCode")>
     <DataMember()> Public Property DesignCode() As String
        Get
            Return Me._DesignCode
        End Get
        Set
            Me._DesignCode = Value
        End Set
    End Property
    <Category("TNX Code Design"), Description("Analysis Only, Check Sections, Cyclic Design"), DisplayName("ERIDesignMode")>
     <DataMember()> Public Property ERIDesignMode() As String
        Get
            Return Me._ERIDesignMode
        End Get
        Set
            Me._ERIDesignMode = Value
        End Set
    End Property
    <Category("TNX Code Design"), Description("consider moments - legs"), DisplayName("DoInteraction")>
     <DataMember()> Public Property DoInteraction() As Boolean?
        Get
            Return Me._DoInteraction
        End Get
        Set
            Me._DoInteraction = Value
        End Set
    End Property
    <Category("TNX Code Design"), Description("consider moments - horizontals"), DisplayName("DoHorzInteraction")>
     <DataMember()> Public Property DoHorzInteraction() As Boolean?
        Get
            Return Me._DoHorzInteraction
        End Get
        Set
            Me._DoHorzInteraction = Value
        End Set
    End Property
    <Category("TNX Code Design"), Description("consider moments - diagonals"), DisplayName("DoDiagInteraction")>
     <DataMember()> Public Property DoDiagInteraction() As Boolean?
        Get
            Return Me._DoDiagInteraction
        End Get
        Set
            Me._DoDiagInteraction = Value
        End Set
    End Property
    <Category("TNX Code Design"), Description(""), DisplayName("UseMomentMagnification")>
     <DataMember()> Public Property UseMomentMagnification() As Boolean?
        Get
            Return Me._UseMomentMagnification
        End Get
        Set
            Me._UseMomentMagnification = Value
        End Set
    End Property
    <Category("TNX Code Design"), Description(""), DisplayName("UseCodeStressRatio")>
     <DataMember()> Public Property UseCodeStressRatio() As Boolean?
        Get
            Return Me._UseCodeStressRatio
        End Get
        Set
            Me._UseCodeStressRatio = Value
        End Set
    End Property
    <Category("TNX Code Design"), Description("base structure allowable stress ratio"), DisplayName("AllowStressRatio")>
     <DataMember()> Public Property AllowStressRatio() As Double?
        Get
            Return Me._AllowStressRatio
        End Get
        Set
            Me._AllowStressRatio = Value
        End Set
    End Property
    <Category("TNX Code Design"), Description("upper structure allowable stress ratio"), DisplayName("AllowAntStressRatio")>
     <DataMember()> Public Property AllowAntStressRatio() As Double?
        Get
            Return Me._AllowAntStressRatio
        End Get
        Set
            Me._AllowAntStressRatio = Value
        End Set
    End Property
    <Category("TNX Code Design"), Description(""), DisplayName("UseCodeGuySF")>
     <DataMember()> Public Property UseCodeGuySF() As Boolean?
        Get
            Return Me._UseCodeGuySF
        End Get
        Set
            Me._UseCodeGuySF = Value
        End Set
    End Property
    <Category("TNX Code Design"), Description(""), DisplayName("GuySF")>
     <DataMember()> Public Property GuySF() As Double?
        Get
            Return Me._GuySF
        End Get
        Set
            Me._GuySF = Value
        End Set
    End Property
    <Category("TNX Code Design"), Description(""), DisplayName("UseTIA222H_AnnexS")>
     <DataMember()> Public Property UseTIA222H_AnnexS() As Boolean?
        Get
            Return Me._UseTIA222H_AnnexS
        End Get
        Set
            Me._UseTIA222H_AnnexS = Value
        End Set
    End Property
    <Category("TNX Code Design"), Description("TIA-222-H Annex S allowable ratio"), DisplayName("TIA_222_H_AnnexS_Ratio")>
     <DataMember()> Public Property TIA_222_H_AnnexS_Ratio() As Double?
        Get
            Return Me._TIA_222_H_AnnexS_Ratio
        End Get
        Set
            Me._TIA_222_H_AnnexS_Ratio = Value
        End Set
    End Property
    <Category("TNX Code Design"), Description(""), DisplayName("PrintBitmaps")>
     <DataMember()> Public Property PrintBitmaps() As Boolean?
        Get
            Return Me._PrintBitmaps
        End Get
        Set
            Me._PrintBitmaps = Value
        End Set
    End Property
#End Region

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As tnxDesign = TryCast(other, tnxDesign)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.DesignCode.CheckChange(otherToCompare.DesignCode, changes, categoryName, "Designcode"), Equals, False)
        Equals = If(Me.ERIDesignMode.CheckChange(otherToCompare.ERIDesignMode, changes, categoryName, "Eridesignmode"), Equals, False)
        Equals = If(Me.DoInteraction.CheckChange(otherToCompare.DoInteraction, changes, categoryName, "Dointeraction"), Equals, False)
        Equals = If(Me.DoHorzInteraction.CheckChange(otherToCompare.DoHorzInteraction, changes, categoryName, "Dohorzinteraction"), Equals, False)
        Equals = If(Me.DoDiagInteraction.CheckChange(otherToCompare.DoDiagInteraction, changes, categoryName, "Dodiaginteraction"), Equals, False)
        Equals = If(Me.UseMomentMagnification.CheckChange(otherToCompare.UseMomentMagnification, changes, categoryName, "Usemomentmagnification"), Equals, False)
        Equals = If(Me.UseCodeStressRatio.CheckChange(otherToCompare.UseCodeStressRatio, changes, categoryName, "Usecodestressratio"), Equals, False)
        Equals = If(Me.AllowStressRatio.CheckChange(otherToCompare.AllowStressRatio, changes, categoryName, "Allowstressratio"), Equals, False)
        Equals = If(Me.AllowAntStressRatio.CheckChange(otherToCompare.AllowAntStressRatio, changes, categoryName, "Allowantstressratio"), Equals, False)
        Equals = If(Me.UseCodeGuySF.CheckChange(otherToCompare.UseCodeGuySF, changes, categoryName, "Usecodeguysf"), Equals, False)
        Equals = If(Me.GuySF.CheckChange(otherToCompare.GuySF, changes, categoryName, "Guysf"), Equals, False)
        Equals = If(Me.UseTIA222H_AnnexS.CheckChange(otherToCompare.UseTIA222H_AnnexS, changes, categoryName, "Usetia222H Annexs"), Equals, False)
        Equals = If(Me.TIA_222_H_AnnexS_Ratio.CheckChange(otherToCompare.TIA_222_H_AnnexS_Ratio, changes, categoryName, "Tia 222 H Annexs Ratio"), Equals, False)
        Equals = If(Me.PrintBitmaps.CheckChange(otherToCompare.PrintBitmaps, changes, categoryName, "Printbitmaps"), Equals, False)

        Return Equals
    End Function

End Class
<DataContract()>
Partial Public Class tnxIce
    Inherits EDSObject

    Public Overrides ReadOnly Property EDSObjectName As String = "Ice"

#Region "Define"
    Private _IceThickness As Double?
    Private _IceDensity As Double?
    Private _UseModified_TIA_222_IceParameters As Boolean?
    Private _TIA_222_IceThicknessMultiplier As Double?
    Private _DoNotUse_TIA_222_IceEscalation As Boolean?
    Private _UseIceEscalation As Boolean?

    <Category("TNX Code Ice"), Description(""), DisplayName("IceThickness")>
     <DataMember()> Public Property IceThickness() As Double?
        Get
            Return Me._IceThickness
        End Get
        Set
            Me._IceThickness = Value
        End Set
    End Property
    <Category("TNX Code Ice"), Description(""), DisplayName("IceDensity")>
     <DataMember()> Public Property IceDensity() As Double?
        Get
            Return Me._IceDensity
        End Get
        Set
            Me._IceDensity = Value
        End Set
    End Property
    <Category("TNX Code Ice"), Description("TIA-222-G/H Custom Ice Options"), DisplayName("UseModified_TIA_222_IceParameters")>
     <DataMember()> Public Property UseModified_TIA_222_IceParameters() As Boolean?
        Get
            Return Me._UseModified_TIA_222_IceParameters
        End Get
        Set
            Me._UseModified_TIA_222_IceParameters = Value
        End Set
    End Property
    <Category("TNX Code Ice"), Description("TIA-222-G/H Custom Ice Options"), DisplayName("TIA_222_IceThicknessMultiplier")>
     <DataMember()> Public Property TIA_222_IceThicknessMultiplier() As Double?
        Get
            Return Me._TIA_222_IceThicknessMultiplier
        End Get
        Set
            Me._TIA_222_IceThicknessMultiplier = Value
        End Set
    End Property
    <Category("TNX Code Ice"), Description("TIA-222-G/H Custom Ice Options"), DisplayName("DoNotUse_TIA_222_IceEscalation")>
     <DataMember()> Public Property DoNotUse_TIA_222_IceEscalation() As Boolean?
        Get
            Return Me._DoNotUse_TIA_222_IceEscalation
        End Get
        Set
            Me._DoNotUse_TIA_222_IceEscalation = Value
        End Set
    End Property
    <Category("TNX Code Ice"), Description("TIA-222-F and earlier"), DisplayName("UseIceEscalation")>
     <DataMember()> Public Property UseIceEscalation() As Boolean?
        Get
            Return Me._UseIceEscalation
        End Get
        Set
            Me._UseIceEscalation = Value
        End Set
    End Property

#End Region
    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As tnxIce = TryCast(other, tnxIce)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.IceThickness.CheckChange(otherToCompare.IceThickness, changes, categoryName, "Ice Thickness"), Equals, False)
        Equals = If(Me.IceDensity.CheckChange(otherToCompare.IceDensity, changes, categoryName, "Ice Density"), Equals, False)
        Equals = If(Me.UseModified_TIA_222_IceParameters.CheckChange(otherToCompare.UseModified_TIA_222_IceParameters, changes, categoryName, "Use Modified TIA-222 Ice Parameters"), Equals, False)
        Equals = If(Me.TIA_222_IceThicknessMultiplier.CheckChange(otherToCompare.TIA_222_IceThicknessMultiplier, changes, categoryName, "TIA-222 Ice Thickness Multiplier"), Equals, False)
        Equals = If(Me.DoNotUse_TIA_222_IceEscalation.CheckChange(otherToCompare.DoNotUse_TIA_222_IceEscalation, changes, categoryName, "Donotuse Tia 222 Iceescalation"), Equals, False)
        Equals = If(Me.UseIceEscalation.CheckChange(otherToCompare.UseIceEscalation, changes, categoryName, "Use Ice Escalation"), Equals, False)

        Return Equals
    End Function
End Class
<DataContract()>
Partial Public Class tnxThermal
    Inherits EDSObject

    Public Overrides ReadOnly Property EDSObjectName As String = "Code"

#Region "Define"
    Private _TempDrop As Double?

    <Category("TNX Code Thermal"), Description(""), DisplayName("TempDrop")>
     <DataMember()> Public Property TempDrop() As Double?
        Get
            Return Me._TempDrop
        End Get
        Set
            Me._TempDrop = Value
        End Set
    End Property
#End Region

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As tnxThermal = TryCast(other, tnxThermal)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.TempDrop.CheckChange(otherToCompare.TempDrop, changes, categoryName, "Tempdrop"), Equals, False)

        Return Equals
    End Function
End Class


<DataContract()>
Partial Public Class tnxWind
    Inherits EDSObject

    Public Overrides ReadOnly Property EDSObjectName As String = "Wind"

#Region "Define"
    Private _WindSpeed As Double?
    Private _WindSpeedIce As Double?
    Private _WindSpeedService As Double?
    Private _UseStateCountyLookup As Boolean?
    Private _State As String
    Private _County As String
    Private _UseMaxKz As Boolean?
    Private _ASCE_7_10_WindData As Boolean?
    Private _ASCE_7_10_ConvertWindToASD As Boolean?
    Private _UseASCEWind As Boolean?
    Private _AutoCalc_ASCE_GH As Boolean?
    Private _ASCE_ExposureCat As Integer?
    Private _ASCE_Year As Integer?
    Private _ASCEGh As Double?
    Private _ASCEI As Double?
    Private _CalcWindAt As Integer?
    Private _WindCalcPoints As Double?
    Private _WindExposure As Integer?
    Private _StructureCategory As Integer?
    Private _RiskCategory As Integer?
    Private _TopoCategory As Integer?
    Private _RSMTopographicFeature As Integer?
    Private _RSM_L As Double?
    Private _RSM_X As Double?
    Private _CrestHeight As Double?
    Private _TIA_222_H_TopoFeatureDownwind As Boolean?
    Private _BaseElevAboveSeaLevel As Double?
    Private _ConsiderRooftopSpeedUp As Boolean?
    Private _RooftopWS As Double?
    Private _RooftopHS As Double?
    Private _RooftopParapetHt As Double?
    Private _RooftopXB As Double?
    Private _WindZone As Integer?
    Private _EIACWindMult As Double?
    Private _EIACWindMultIce As Double?
    Private _EIACIgnoreCableDrag As Boolean?
    Private _CSA_S37_RefVelPress As Double?
    Private _CSA_S37_ReliabilityClass As Integer?
    Private _CSA_S37_ServiceabilityFactor As Double?

    <Category("TNX Code Wind"), Description(""), DisplayName("WindSpeed")>
     <DataMember()> Public Property WindSpeed() As Double?
        Get
            Return Me._WindSpeed
        End Get
        Set
            Me._WindSpeed = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("WindSpeedIce")>
     <DataMember()> Public Property WindSpeedIce() As Double?
        Get
            Return Me._WindSpeedIce
        End Get
        Set
            Me._WindSpeedIce = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("WindSpeedService")>
     <DataMember()> Public Property WindSpeedService() As Double?
        Get
            Return Me._WindSpeedService
        End Get
        Set
            Me._WindSpeedService = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("UseStateCountyLookup")>
     <DataMember()> Public Property UseStateCountyLookup() As Boolean?
        Get
            Return Me._UseStateCountyLookup
        End Get
        Set
            Me._UseStateCountyLookup = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("State")>
     <DataMember()> Public Property State() As String
        Get
            Return Me._State
        End Get
        Set
            Me._State = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("County")>
     <DataMember()> Public Property County() As String
        Get
            Return Me._County
        End Get
        Set
            Me._County = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("UseMaxKz")>
     <DataMember()> Public Property UseMaxKz() As Boolean?
        Get
            Return Me._UseMaxKz
        End Get
        Set
            Me._UseMaxKz = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("TIA-222-G Only"), DisplayName("ASCE_7_10_WindData")>
     <DataMember()> Public Property ASCE_7_10_WindData() As Boolean?
        Get
            Return Me._ASCE_7_10_WindData
        End Get
        Set
            Me._ASCE_7_10_WindData = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("TIA-222-G Only"), DisplayName("ASCE_7_10_ConvertWindToASD")>
     <DataMember()> Public Property ASCE_7_10_ConvertWindToASD() As Boolean?
        Get
            Return Me._ASCE_7_10_ConvertWindToASD
        End Get
        Set
            Me._ASCE_7_10_ConvertWindToASD = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("Use Special Wind Profile"), DisplayName("UseASCEWind")>
     <DataMember()> Public Property UseASCEWind() As Boolean?
        Get
            Return Me._UseASCEWind
        End Get
        Set
            Me._UseASCEWind = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("Use TIA Gh Value"), DisplayName("AutoCalc_ASCE_GH")>
     <DataMember()> Public Property AutoCalc_ASCE_GH() As Boolean?
        Get
            Return Me._AutoCalc_ASCE_GH
        End Get
        Set
            Me._AutoCalc_ASCE_GH = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("{0 = B, 1 = C,  2 = D}"), DisplayName("ASCE_ExposureCat")>
     <DataMember()> Public Property ASCE_ExposureCat() As Integer?
        Get
            Return Me._ASCE_ExposureCat
        End Get
        Set
            Me._ASCE_ExposureCat = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("{0 = ASCE 7-88, 1 = ASCE 7-93, 2= ASCE 7-95, 3 = ASCE 7-98, 4 = ASCE 7-02, 5 = Cook Co., IL, 6 = WIS 53, 7 = Chicago}"), DisplayName("ASCE_Year")>
     <DataMember()> Public Property ASCE_Year() As Integer?
        Get
            Return Me._ASCE_Year
        End Get
        Set
            Me._ASCE_Year = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("ASCEGh")>
     <DataMember()> Public Property ASCEGh() As Double?
        Get
            Return Me._ASCEGh
        End Get
        Set
            Me._ASCEGh = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("ASCEI")>
     <DataMember()> Public Property ASCEI() As Double?
        Get
            Return Me._ASCEI
        End Get
        Set
            Me._ASCEI = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("{0 = Every Section, 1 = between guys, 2 = user specify (WindCalcPoints)}"), DisplayName("CalcWindAt")>
     <DataMember()> Public Property CalcWindAt() As Integer?
        Get
            Return Me._CalcWindAt
        End Get
        Set
            Me._CalcWindAt = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("WindCalcPoints")>
     <DataMember()> Public Property WindCalcPoints() As Double?
        Get
            Return Me._WindCalcPoints
        End Get
        Set
            Me._WindCalcPoints = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("{0 = B, 1 = C, 2 = D}"), DisplayName("WindExposure")>
     <DataMember()> Public Property WindExposure() As Integer?
        Get
            Return Me._WindExposure
        End Get
        Set
            Me._WindExposure = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("Structure Class - TIA-222-G Only {0 = I, 1 = II, 2 = III}"), DisplayName("StructureCategory")>
     <DataMember()> Public Property StructureCategory() As Integer?
        Get
            Return Me._StructureCategory
        End Get
        Set
            Me._StructureCategory = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("TIA-222-H Only {0 = I, 1 = II,  2 = III,  3 = IV}"), DisplayName("RiskCategory")>
     <DataMember()> Public Property RiskCategory() As Integer?
        Get
            Return Me._RiskCategory
        End Get
        Set
            Me._RiskCategory = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("{0 = 1, 1 = 2, 2 = 3, 3 = 4, 4 = 5/Rigorous Procedure}"), DisplayName("TopoCategory")>
     <DataMember()> Public Property TopoCategory() As Integer?
        Get
            Return Me._TopoCategory
        End Get
        Set
            Me._TopoCategory = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("{0 = Continuous Ridge, 1 = Flat Topped Ridge, 2 = Hill, 3 = Flat Topped Hill, 4 = Continuous Escarpment}"), DisplayName("RSMTopographicFeature")>
     <DataMember()> Public Property RSMTopographicFeature() As Integer?
        Get
            Return Me._RSMTopographicFeature
        End Get
        Set
            Me._RSMTopographicFeature = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("RSM_L")>
     <DataMember()> Public Property RSM_L() As Double?
        Get
            Return Me._RSM_L
        End Get
        Set
            Me._RSM_L = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("RSM_X")>
     <DataMember()> Public Property RSM_X() As Double?
        Get
            Return Me._RSM_X
        End Get
        Set
            Me._RSM_X = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("CrestHeight")>
     <DataMember()> Public Property CrestHeight() As Double?
        Get
            Return Me._CrestHeight
        End Get
        Set
            Me._CrestHeight = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("TIA_222_H_TopoFeatureDownwind")>
     <DataMember()> Public Property TIA_222_H_TopoFeatureDownwind() As Boolean?
        Get
            Return Me._TIA_222_H_TopoFeatureDownwind
        End Get
        Set
            Me._TIA_222_H_TopoFeatureDownwind = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("BaseElevAboveSeaLevel")>
     <DataMember()> Public Property BaseElevAboveSeaLevel() As Double?
        Get
            Return Me._BaseElevAboveSeaLevel
        End Get
        Set
            Me._BaseElevAboveSeaLevel = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("ConsiderRooftopSpeedUp")>
     <DataMember()> Public Property ConsiderRooftopSpeedUp() As Boolean?
        Get
            Return Me._ConsiderRooftopSpeedUp
        End Get
        Set
            Me._ConsiderRooftopSpeedUp = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("RooftopWS")>
     <DataMember()> Public Property RooftopWS() As Double?
        Get
            Return Me._RooftopWS
        End Get
        Set
            Me._RooftopWS = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("RooftopHS")>
     <DataMember()> Public Property RooftopHS() As Double?
        Get
            Return Me._RooftopHS
        End Get
        Set
            Me._RooftopHS = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("RooftopParapetHt")>
     <DataMember()> Public Property RooftopParapetHt() As Double?
        Get
            Return Me._RooftopParapetHt
        End Get
        Set
            Me._RooftopParapetHt = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("RooftopXB")>
     <DataMember()> Public Property RooftopXB() As Double?
        Get
            Return Me._RooftopXB
        End Get
        Set
            Me._RooftopXB = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("EIA-222-C and earlier {0 = A, 1 = B, 2 = C}"), DisplayName("WindZone")>
     <DataMember()> Public Property WindZone() As Integer?
        Get
            Return Me._WindZone
        End Get
        Set
            Me._WindZone = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("EIA-222-C and earlier"), DisplayName("EIACWindMult")>
     <DataMember()> Public Property EIACWindMult() As Double?
        Get
            Return Me._EIACWindMult
        End Get
        Set
            Me._EIACWindMult = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("EIA-222-C and earlier"), DisplayName("EIACWindMultIce")>
     <DataMember()> Public Property EIACWindMultIce() As Double?
        Get
            Return Me._EIACWindMultIce
        End Get
        Set
            Me._EIACWindMultIce = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("EIA-222-C and earlier - Set Cable Drag Factor to 1.0"), DisplayName("EIACIgnoreCableDrag")>
     <DataMember()> Public Property EIACIgnoreCableDrag() As Boolean?
        Get
            Return Me._EIACIgnoreCableDrag
        End Get
        Set
            Me._EIACIgnoreCableDrag = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("CSA S37-01 only"), DisplayName("CSA_S37_RefVelPress")>
     <DataMember()> Public Property CSA_S37_RefVelPress() As Double?
        Get
            Return Me._CSA_S37_RefVelPress
        End Get
        Set
            Me._CSA_S37_RefVelPress = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("CSA S37-01 only {0 = I, 1 = II, 2 = III}"), DisplayName("CSA_S37_ReliabilityClass")>
     <DataMember()> Public Property CSA_S37_ReliabilityClass() As Integer?
        Get
            Return Me._CSA_S37_ReliabilityClass
        End Get
        Set
            Me._CSA_S37_ReliabilityClass = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("CSA S37-01 only"), DisplayName("CSA_S37_ServiceabilityFactor")>
     <DataMember()> Public Property CSA_S37_ServiceabilityFactor() As Double?
        Get
            Return Me._CSA_S37_ServiceabilityFactor
        End Get
        Set
            Me._CSA_S37_ServiceabilityFactor = Value
        End Set
    End Property
#End Region
    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As tnxWind = TryCast(other, tnxWind)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.WindSpeed.CheckChange(otherToCompare.WindSpeed, changes, categoryName, "Windspeed"), Equals, False)
        Equals = If(Me.WindSpeedIce.CheckChange(otherToCompare.WindSpeedIce, changes, categoryName, "Windspeedice"), Equals, False)
        Equals = If(Me.WindSpeedService.CheckChange(otherToCompare.WindSpeedService, changes, categoryName, "Windspeedservice"), Equals, False)
        Equals = If(Me.UseStateCountyLookup.CheckChange(otherToCompare.UseStateCountyLookup, changes, categoryName, "Usestatecountylookup"), Equals, False)
        Equals = If(Me.State.CheckChange(otherToCompare.State, changes, categoryName, "State"), Equals, False)
        Equals = If(Me.County.CheckChange(otherToCompare.County, changes, categoryName, "County"), Equals, False)
        Equals = If(Me.UseMaxKz.CheckChange(otherToCompare.UseMaxKz, changes, categoryName, "Usemaxkz"), Equals, False)
        Equals = If(Me.ASCE_7_10_WindData.CheckChange(otherToCompare.ASCE_7_10_WindData, changes, categoryName, "Asce 7 10 Winddata"), Equals, False)
        Equals = If(Me.ASCE_7_10_ConvertWindToASD.CheckChange(otherToCompare.ASCE_7_10_ConvertWindToASD, changes, categoryName, "Asce 7 10 Convertwindtoasd"), Equals, False)
        Equals = If(Me.UseASCEWind.CheckChange(otherToCompare.UseASCEWind, changes, categoryName, "Useascewind"), Equals, False)
        Equals = If(Me.AutoCalc_ASCE_GH.CheckChange(otherToCompare.AutoCalc_ASCE_GH, changes, categoryName, "Autocalc Asce Gh"), Equals, False)
        Equals = If(Me.ASCE_ExposureCat.CheckChange(otherToCompare.ASCE_ExposureCat, changes, categoryName, "Asce Exposurecat"), Equals, False)
        Equals = If(Me.ASCE_Year.CheckChange(otherToCompare.ASCE_Year, changes, categoryName, "Asce Year"), Equals, False)
        Equals = If(Me.ASCEGh.CheckChange(otherToCompare.ASCEGh, changes, categoryName, "Ascegh"), Equals, False)
        Equals = If(Me.ASCEI.CheckChange(otherToCompare.ASCEI, changes, categoryName, "Ascei"), Equals, False)
        Equals = If(Me.CalcWindAt.CheckChange(otherToCompare.CalcWindAt, changes, categoryName, "Calcwindat"), Equals, False)
        Equals = If(Me.WindCalcPoints.CheckChange(otherToCompare.WindCalcPoints, changes, categoryName, "Windcalcpoints"), Equals, False)
        Equals = If(Me.WindExposure.CheckChange(otherToCompare.WindExposure, changes, categoryName, "Windexposure"), Equals, False)
        Equals = If(Me.StructureCategory.CheckChange(otherToCompare.StructureCategory, changes, categoryName, "Structurecategory"), Equals, False)
        Equals = If(Me.RiskCategory.CheckChange(otherToCompare.RiskCategory, changes, categoryName, "Riskcategory"), Equals, False)
        Equals = If(Me.TopoCategory.CheckChange(otherToCompare.TopoCategory, changes, categoryName, "Topocategory"), Equals, False)
        Equals = If(Me.RSMTopographicFeature.CheckChange(otherToCompare.RSMTopographicFeature, changes, categoryName, "Rsmtopographicfeature"), Equals, False)
        Equals = If(Me.RSM_L.CheckChange(otherToCompare.RSM_L, changes, categoryName, "Rsm L"), Equals, False)
        Equals = If(Me.RSM_X.CheckChange(otherToCompare.RSM_X, changes, categoryName, "Rsm X"), Equals, False)
        Equals = If(Me.CrestHeight.CheckChange(otherToCompare.CrestHeight, changes, categoryName, "Crestheight"), Equals, False)
        Equals = If(Me.TIA_222_H_TopoFeatureDownwind.CheckChange(otherToCompare.TIA_222_H_TopoFeatureDownwind, changes, categoryName, "Tia 222 H Topofeaturedownwind"), Equals, False)
        Equals = If(Me.BaseElevAboveSeaLevel.CheckChange(otherToCompare.BaseElevAboveSeaLevel, changes, categoryName, "Baseelevabovesealevel"), Equals, False)
        Equals = If(Me.ConsiderRooftopSpeedUp.CheckChange(otherToCompare.ConsiderRooftopSpeedUp, changes, categoryName, "Considerrooftopspeedup"), Equals, False)
        Equals = If(Me.RooftopWS.CheckChange(otherToCompare.RooftopWS, changes, categoryName, "Rooftopws"), Equals, False)
        Equals = If(Me.RooftopHS.CheckChange(otherToCompare.RooftopHS, changes, categoryName, "Rooftophs"), Equals, False)
        Equals = If(Me.RooftopParapetHt.CheckChange(otherToCompare.RooftopParapetHt, changes, categoryName, "Rooftopparapetht"), Equals, False)
        Equals = If(Me.RooftopXB.CheckChange(otherToCompare.RooftopXB, changes, categoryName, "Rooftopxb"), Equals, False)
        Equals = If(Me.WindZone.CheckChange(otherToCompare.WindZone, changes, categoryName, "Windzone"), Equals, False)
        Equals = If(Me.EIACWindMult.CheckChange(otherToCompare.EIACWindMult, changes, categoryName, "Eiacwindmult"), Equals, False)
        Equals = If(Me.EIACWindMultIce.CheckChange(otherToCompare.EIACWindMultIce, changes, categoryName, "Eiacwindmultice"), Equals, False)
        Equals = If(Me.EIACIgnoreCableDrag.CheckChange(otherToCompare.EIACIgnoreCableDrag, changes, categoryName, "Eiacignorecabledrag"), Equals, False)
        Equals = If(Me.CSA_S37_RefVelPress.CheckChange(otherToCompare.CSA_S37_RefVelPress, changes, categoryName, "Csa S37 Refvelpress"), Equals, False)
        Equals = If(Me.CSA_S37_ReliabilityClass.CheckChange(otherToCompare.CSA_S37_ReliabilityClass, changes, categoryName, "Csa S37 Reliabilityclass"), Equals, False)
        Equals = If(Me.CSA_S37_ServiceabilityFactor.CheckChange(otherToCompare.CSA_S37_ServiceabilityFactor, changes, categoryName, "Csa S37 Serviceabilityfactor"), Equals, False)

        Return Equals
    End Function
End Class
<DataContract()>
Partial Public Class tnxMisclCode
    Inherits EDSObject

    Public Overrides ReadOnly Property EDSObjectName As String = "Miscl"

#Region "Define"
    Private _GroutFc As Double?
    Private _TowerBoltGrade As String
    Private _TowerBoltMinEdgeDist As Double?

    <Category("TNX Code Miscellaneous"), Description(""), DisplayName("GroutFc")>
     <DataMember()> Public Property GroutFc() As Double?
        Get
            Return Me._GroutFc
        End Get
        Set
            Me._GroutFc = Value
        End Set
    End Property
    <Category("TNX Code Miscellaneous"), Description("Default bolt grade"), DisplayName("TowerBoltGrade")>
     <DataMember()> Public Property TowerBoltGrade() As String
        Get
            Return Me._TowerBoltGrade
        End Get
        Set
            Me._TowerBoltGrade = Value
        End Set
    End Property
    <Category("TNX Code Miscellaneous"), Description("Not in UI"), DisplayName("TowerBoltMinEdgeDist")>
     <DataMember()> Public Property TowerBoltMinEdgeDist() As Double?
        Get
            Return Me._TowerBoltMinEdgeDist
        End Get
        Set
            Me._TowerBoltMinEdgeDist = Value
        End Set
    End Property
#End Region
    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As tnxMisclCode = TryCast(other, tnxMisclCode)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.GroutFc.CheckChange(otherToCompare.GroutFc, changes, categoryName, "Grout fc"), Equals, False)
        Equals = If(Me.TowerBoltGrade.CheckChange(otherToCompare.TowerBoltGrade, changes, categoryName, "Tower Bolt Grade"), Equals, False)
        Equals = If(Me.TowerBoltMinEdgeDist.CheckChange(otherToCompare.TowerBoltMinEdgeDist, changes, categoryName, "Tower Bolt Min Edge Dist"), Equals, False)

        Return Equals
    End Function
End Class
<DataContract()>
Partial Public Class tnxSeismic
    Inherits EDSObject

    Public Overrides ReadOnly Property EDSObjectName As String = "Seismic"

#Region "Define"

    Private _UseASCE7_10_Seismic_Lcomb As Boolean?
    Private _SeismicSiteClass As Integer?
    Private _SeismicSs As Double?
    Private _SeismicS1 As Double?

    <Category("TNX Code seismic"), Description(""), DisplayName("UseASCE7_10_Seismic_Lcomb")>
     <DataMember()> Public Property UseASCE7_10_Seismic_Lcomb() As Boolean?
        Get
            Return Me._UseASCE7_10_Seismic_Lcomb
        End Get
        Set
            Me._UseASCE7_10_Seismic_Lcomb = Value
        End Set
    End Property
    <Category("TNX Code seismic"), Description("not in UI {0 = A, 1 = B, 2 = C, 3 = D, 4 = E} "), DisplayName("SeismicSiteClass")>
     <DataMember()> Public Property SeismicSiteClass() As Integer?
        Get
            Return Me._SeismicSiteClass
        End Get
        Set
            Me._SeismicSiteClass = Value
        End Set
    End Property
    <Category("TNX Code seismic"), Description("not in UI"), DisplayName("SeismicSs")>
     <DataMember()> Public Property SeismicSs() As Double?
        Get
            Return Me._SeismicSs
        End Get
        Set
            Me._SeismicSs = Value
        End Set
    End Property
    <Category("TNX Code seismic"), Description("not in UI"), DisplayName("SeismicS1")>
     <DataMember()> Public Property SeismicS1() As Double?
        Get
            Return Me._SeismicS1
        End Get
        Set
            Me._SeismicS1 = Value
        End Set
    End Property
#End Region

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As tnxSeismic = TryCast(other, tnxSeismic)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.UseASCE7_10_Seismic_Lcomb.CheckChange(otherToCompare.UseASCE7_10_Seismic_Lcomb, changes, categoryName, "Useasce7 10 Seismic Lcomb"), Equals, False)
        Equals = If(Me.SeismicSiteClass.CheckChange(otherToCompare.SeismicSiteClass, changes, categoryName, "Seismicsiteclass"), Equals, False)
        Equals = If(Me.SeismicSs.CheckChange(otherToCompare.SeismicSs, changes, categoryName, "Seismicss"), Equals, False)
        Equals = If(Me.SeismicS1.CheckChange(otherToCompare.SeismicS1, changes, categoryName, "Seismics1"), Equals, False)

        Return Equals
    End Function
End Class
#End Region

#Region "Options"
<DataContract()>
Partial Public Class tnxOptions
    Inherits EDSObject

    Public Overrides ReadOnly Property EDSObjectName As String = "Options"

#Region "Define"
    Private _UseClearSpans As Boolean?
    Private _UseClearSpansKlr As Boolean?
    Private _UseFeedlineAsCylinder As Boolean?
    Private _UseLegLoads As Boolean?
    Private _SRTakeCompression As Boolean?
    Private _AllLegPanelsSame As Boolean?
    Private _UseCombinedBoltCapacity As Boolean?
    Private _SecHorzBracesLeg As Boolean?
    Private _SortByComponent As Boolean?
    Private _SRCutEnds As Boolean?
    Private _SRConcentric As Boolean?
    Private _CalcBlockShear As Boolean?
    Private _Use4SidedDiamondBracing As Boolean?
    Private _TriangulateInnerBracing As Boolean?
    Private _PrintCarrierNotes As Boolean?
    Private _AddIBCWindCase As Boolean?
    Private _LegBoltsAtTop As Boolean?
    Private _UseTIA222Exemptions_MinBracingResistance As Boolean?
    Private _UseTIA222Exemptions_TensionSplice As Boolean?
    Private _IgnoreKLryFor60DegAngleLegs As Boolean?
    Private _UseFeedlineTorque As Boolean?
    Private _UsePinnedElements As Boolean?
    Private _UseRigidIndex As Boolean?
    Private _UseTrueCable As Boolean?
    Private _UseASCELy As Boolean?
    Private _CalcBracingForces As Boolean?
    Private _IgnoreBracingFEA As Boolean?
    Private _BypassStabilityChecks As Boolean?
    Private _UseWindProjection As Boolean?
    Private _UseDishCoeff As Boolean?
    Private _AutoCalcTorqArmArea As Boolean?
    Private _foundationStiffness As New tnxFoundationStiffness()
    Private _defaultGirtOffsets As New tnxDefaultGirtOffsets()
    Private _cantileverPoles As New tnxCantileverPoles()
    Private _windDirections As New tnxWindDirections()
    Private _misclOptions As New tnxMisclOptions()

    <Category("TNX Options"), Description(""), DisplayName("UseClearSpans")>
     <DataMember()> Public Property UseClearSpans() As Boolean?
        Get
            Return Me._UseClearSpans
        End Get
        Set
            Me._UseClearSpans = Value
        End Set
    End Property
    <Category("TNX Options"), Description(""), DisplayName("UseClearSpansKlr")>
     <DataMember()> Public Property UseClearSpansKlr() As Boolean?
        Get
            Return Me._UseClearSpansKlr
        End Get
        Set
            Me._UseClearSpansKlr = Value
        End Set
    End Property
    <Category("TNX Options"), Description("treat feedline bundles as cylindrical"), DisplayName("UseFeedlineAsCylinder")>
     <DataMember()> Public Property UseFeedlineAsCylinder() As Boolean?
        Get
            Return Me._UseFeedlineAsCylinder
        End Get
        Set
            Me._UseFeedlineAsCylinder = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Distribute Leg Loads As Uniform"), DisplayName("UseLegLoads")>
     <DataMember()> Public Property UseLegLoads() As Boolean?
        Get
            Return Me._UseLegLoads
        End Get
        Set
            Me._UseLegLoads = Value
        End Set
    End Property
    <Category("TNX Options"), Description("SR Sleeve Bolts Resist Compression"), DisplayName("SRTakeCompression")>
     <DataMember()> Public Property SRTakeCompression() As Boolean?
        Get
            Return Me._SRTakeCompression
        End Get
        Set
            Me._SRTakeCompression = Value
        End Set
    End Property
    <Category("TNX Options"), Description("All Leg Panels Have Same Allowable"), DisplayName("AllLegPanelsSame")>
     <DataMember()> Public Property AllLegPanelsSame() As Boolean?
        Get
            Return Me._AllLegPanelsSame
        End Get
        Set
            Me._AllLegPanelsSame = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Include Bolts In Member Capacity"), DisplayName("UseCombinedBoltCapacity")>
     <DataMember()> Public Property UseCombinedBoltCapacity() As Boolean?
        Get
            Return Me._UseCombinedBoltCapacity
        End Get
        Set
            Me._UseCombinedBoltCapacity = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Secondary Horizontal Braces Leg"), DisplayName("SecHorzBracesLeg")>
     <DataMember()> Public Property SecHorzBracesLeg() As Boolean?
        Get
            Return Me._SecHorzBracesLeg
        End Get
        Set
            Me._SecHorzBracesLeg = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Sort Capacity Reports By Component"), DisplayName("SortByComponent")>
     <DataMember()> Public Property SortByComponent() As Boolean?
        Get
            Return Me._SortByComponent
        End Get
        Set
            Me._SortByComponent = Value
        End Set
    End Property
    <Category("TNX Options"), Description("SR Members Have Cut Ends"), DisplayName("SRCutEnds")>
     <DataMember()> Public Property SRCutEnds() As Boolean?
        Get
            Return Me._SRCutEnds
        End Get
        Set
            Me._SRCutEnds = Value
        End Set
    End Property
    <Category("TNX Options"), Description("SR Members Are Concentric"), DisplayName("SRConcentric")>
     <DataMember()> Public Property SRConcentric() As Boolean?
        Get
            Return Me._SRConcentric
        End Get
        Set
            Me._SRConcentric = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Include Angle Block Shear Check"), DisplayName("CalcBlockShear")>
     <DataMember()> Public Property CalcBlockShear() As Boolean?
        Get
            Return Me._CalcBlockShear
        End Get
        Set
            Me._CalcBlockShear = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Use Diamond Inner Bracing"), DisplayName("Use4SidedDiamondBracing")>
     <DataMember()> Public Property Use4SidedDiamondBracing() As Boolean?
        Get
            Return Me._Use4SidedDiamondBracing
        End Get
        Set
            Me._Use4SidedDiamondBracing = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Triangulate Diamond Inner Bracing"), DisplayName("TriangulateInnerBracing")>
     <DataMember()> Public Property TriangulateInnerBracing() As Boolean?
        Get
            Return Me._TriangulateInnerBracing
        End Get
        Set
            Me._TriangulateInnerBracing = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Print Carrier/Notes"), DisplayName("PrintCarrierNotes")>
     <DataMember()> Public Property PrintCarrierNotes() As Boolean?
        Get
            Return Me._PrintCarrierNotes
        End Get
        Set
            Me._PrintCarrierNotes = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Add IBC .6D+W Combination"), DisplayName("AddIBCWindCase")>
     <DataMember()> Public Property AddIBCWindCase() As Boolean?
        Get
            Return Me._AddIBCWindCase
        End Get
        Set
            Me._AddIBCWindCase = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Leg Bolts Are At Top Of Section"), DisplayName("LegBoltsAtTop")>
     <DataMember()> Public Property LegBoltsAtTop() As Boolean?
        Get
            Return Me._LegBoltsAtTop
        End Get
        Set
            Me._LegBoltsAtTop = Value
        End Set
    End Property
    <Category("TNX Options"), Description(""), DisplayName("UseTIA222Exemptions_MinBracingResistance")>
     <DataMember()> Public Property UseTIA222Exemptions_MinBracingResistance() As Boolean?
        Get
            Return Me._UseTIA222Exemptions_MinBracingResistance
        End Get
        Set
            Me._UseTIA222Exemptions_MinBracingResistance = Value
        End Set
    End Property
    <Category("TNX Options"), Description(""), DisplayName("UseTIA222Exemptions_TensionSplice")>
     <DataMember()> Public Property UseTIA222Exemptions_TensionSplice() As Boolean?
        Get
            Return Me._UseTIA222Exemptions_TensionSplice
        End Get
        Set
            Me._UseTIA222Exemptions_TensionSplice = Value
        End Set
    End Property
    <Category("TNX Options"), Description(""), DisplayName("IgnoreKLryFor60DegAngleLegs")>
     <DataMember()> Public Property IgnoreKLryFor60DegAngleLegs() As Boolean?
        Get
            Return Me._IgnoreKLryFor60DegAngleLegs
        End Get
        Set
            Me._IgnoreKLryFor60DegAngleLegs = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Consider Feed Line Torque"), DisplayName("UseFeedlineTorque")>
     <DataMember()> Public Property UseFeedlineTorque() As Boolean?
        Get
            Return Me._UseFeedlineTorque
        End Get
        Set
            Me._UseFeedlineTorque = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Assume Legs Pinned"), DisplayName("UsePinnedElements")>
     <DataMember()> Public Property UsePinnedElements() As Boolean?
        Get
            Return Me._UsePinnedElements
        End Get
        Set
            Me._UsePinnedElements = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Assume Rigid Index Plate"), DisplayName("UseRigidIndex")>
     <DataMember()> Public Property UseRigidIndex() As Boolean?
        Get
            Return Me._UseRigidIndex
        End Get
        Set
            Me._UseRigidIndex = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Retension Guys To Initial Tension"), DisplayName("UseTrueCable")>
     <DataMember()> Public Property UseTrueCable() As Boolean?
        Get
            Return Me._UseTrueCable
        End Get
        Set
            Me._UseTrueCable = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Use ASCE 10 X-Brace Ly Rules"), DisplayName("UseASCELy")>
     <DataMember()> Public Property UseASCELy() As Boolean?
        Get
            Return Me._UseASCELy
        End Get
        Set
            Me._UseASCELy = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Calculate Forces in Supporing Bracing Members"), DisplayName("CalcBracingForces")>
     <DataMember()> Public Property CalcBracingForces() As Boolean?
        Get
            Return Me._CalcBracingForces
        End Get
        Set
            Me._CalcBracingForces = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Ignore Redundant Bracing in FEA"), DisplayName("IgnoreBracingFEA")>
     <DataMember()> Public Property IgnoreBracingFEA() As Boolean?
        Get
            Return Me._IgnoreBracingFEA
        End Get
        Set
            Me._IgnoreBracingFEA = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Bypass Mast Stability Checks"), DisplayName("BypassStabilityChecks")>
     <DataMember()> Public Property BypassStabilityChecks() As Boolean?
        Get
            Return Me._BypassStabilityChecks
        End Get
        Set
            Me._BypassStabilityChecks = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Project Wind Area Of Appurtenances"), DisplayName("UseWindProjection")>
     <DataMember()> Public Property UseWindProjection() As Boolean?
        Get
            Return Me._UseWindProjection
        End Get
        Set
            Me._UseWindProjection = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Use Azimuth Dish Coefficients"), DisplayName("UseDishCoeff")>
     <DataMember()> Public Property UseDishCoeff() As Boolean?
        Get
            Return Me._UseDishCoeff
        End Get
        Set
            Me._UseDishCoeff = Value
        End Set
    End Property
    <Category("TNX Options"), Description("AutoCalc Torque Arm Area"), DisplayName("AutoCalcTorqArmArea")>
     <DataMember()> Public Property AutoCalcTorqArmArea() As Boolean?
        Get
            Return Me._AutoCalcTorqArmArea
        End Get
        Set
            Me._AutoCalcTorqArmArea = Value
        End Set
    End Property

    <Category("TNX Options"), Description(""), DisplayName("Foundation Stiffness Options")>
     <DataMember()> Public Property foundationStiffness() As tnxFoundationStiffness
        Get
            Return Me._foundationStiffness
        End Get
        Set
            Me._foundationStiffness = Value
        End Set
    End Property

    <Category("TNX Options"), Description(""), DisplayName("Default Girt Offsets Options")>
     <DataMember()> Public Property defaultGirtOffsets() As tnxDefaultGirtOffsets
        Get
            Return Me._defaultGirtOffsets
        End Get
        Set
            Me._defaultGirtOffsets = Value
        End Set
    End Property

    <Category("TNX Options"), Description(""), DisplayName("Cantilever Pole Options")>
     <DataMember()> Public Property cantileverPoles() As tnxCantileverPoles
        Get
            Return Me._cantileverPoles
        End Get
        Set
            Me._cantileverPoles = Value
        End Set
    End Property

    <Category("TNX Options"), Description(""), DisplayName("Wind Direction Options")>
     <DataMember()> Public Property windDirections() As tnxWindDirections
        Get
            Return Me._windDirections
        End Get
        Set
            Me._windDirections = Value
        End Set
    End Property

    <Category("TNX Options"), Description(""), DisplayName("Miscellaneous Options")>
     <DataMember()> Public Property misclOptions() As tnxMisclOptions
        Get
            Return Me._misclOptions
        End Get
        Set
            Me._misclOptions = Value
        End Set
    End Property
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
        Dim otherToCompare As tnxOptions = TryCast(other, tnxOptions)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.UseClearSpans.CheckChange(otherToCompare.UseClearSpans, changes, categoryName, "Use Clear Spans"), Equals, False)
        Equals = If(Me.UseClearSpansKlr.CheckChange(otherToCompare.UseClearSpansKlr, changes, categoryName, "Use Clear Spans Klr"), Equals, False)
        Equals = If(Me.UseFeedlineAsCylinder.CheckChange(otherToCompare.UseFeedlineAsCylinder, changes, categoryName, "Use Feedline As Cylinder"), Equals, False)
        Equals = If(Me.UseLegLoads.CheckChange(otherToCompare.UseLegLoads, changes, categoryName, "Use Leg Loads"), Equals, False)
        Equals = If(Me.SRTakeCompression.CheckChange(otherToCompare.SRTakeCompression, changes, categoryName, "SR Take Compression"), Equals, False)
        Equals = If(Me.AllLegPanelsSame.CheckChange(otherToCompare.AllLegPanelsSame, changes, categoryName, "All Leg Panels Same"), Equals, False)
        Equals = If(Me.UseCombinedBoltCapacity.CheckChange(otherToCompare.UseCombinedBoltCapacity, changes, categoryName, "Use Combined Bolt Capacity"), Equals, False)
        Equals = If(Me.SecHorzBracesLeg.CheckChange(otherToCompare.SecHorzBracesLeg, changes, categoryName, "Sec Horz Braces Leg"), Equals, False)
        Equals = If(Me.SortByComponent.CheckChange(otherToCompare.SortByComponent, changes, categoryName, "Sort By Component"), Equals, False)
        Equals = If(Me.SRCutEnds.CheckChange(otherToCompare.SRCutEnds, changes, categoryName, "SR Cut Ends"), Equals, False)
        Equals = If(Me.SRConcentric.CheckChange(otherToCompare.SRConcentric, changes, categoryName, "SR Concentric"), Equals, False)
        Equals = If(Me.CalcBlockShear.CheckChange(otherToCompare.CalcBlockShear, changes, categoryName, "Calc Block Shear"), Equals, False)
        Equals = If(Me.Use4SidedDiamondBracing.CheckChange(otherToCompare.Use4SidedDiamondBracing, changes, categoryName, "Use 4 Sided Diamond Bracing"), Equals, False)
        Equals = If(Me.TriangulateInnerBracing.CheckChange(otherToCompare.TriangulateInnerBracing, changes, categoryName, "Triangulate Inner Bracing"), Equals, False)
        Equals = If(Me.PrintCarrierNotes.CheckChange(otherToCompare.PrintCarrierNotes, changes, categoryName, "Print Carrier Notes"), Equals, False)
        Equals = If(Me.AddIBCWindCase.CheckChange(otherToCompare.AddIBCWindCase, changes, categoryName, "Add IBCWind Case"), Equals, False)
        Equals = If(Me.LegBoltsAtTop.CheckChange(otherToCompare.LegBoltsAtTop, changes, categoryName, "Leg Bolts At Top"), Equals, False)
        Equals = If(Me.UseTIA222Exemptions_MinBracingResistance.CheckChange(otherToCompare.UseTIA222Exemptions_MinBracingResistance, changes, categoryName, "Use TIA 222 Exemptions  Min Bracing Resistance"), Equals, False)
        Equals = If(Me.UseTIA222Exemptions_TensionSplice.CheckChange(otherToCompare.UseTIA222Exemptions_TensionSplice, changes, categoryName, "Use TIA 222 Exemptions  Tension Splice"), Equals, False)
        Equals = If(Me.IgnoreKLryFor60DegAngleLegs.CheckChange(otherToCompare.IgnoreKLryFor60DegAngleLegs, changes, categoryName, "Ignore KLry For 60 Deg Angle Legs"), Equals, False)
        Equals = If(Me.UseFeedlineTorque.CheckChange(otherToCompare.UseFeedlineTorque, changes, categoryName, "Use Feedline Torque"), Equals, False)
        Equals = If(Me.UsePinnedElements.CheckChange(otherToCompare.UsePinnedElements, changes, categoryName, "Use Pinned Elements"), Equals, False)
        Equals = If(Me.UseRigidIndex.CheckChange(otherToCompare.UseRigidIndex, changes, categoryName, "Use Rigid Index"), Equals, False)
        Equals = If(Me.UseTrueCable.CheckChange(otherToCompare.UseTrueCable, changes, categoryName, "Use True Cable"), Equals, False)
        Equals = If(Me.UseASCELy.CheckChange(otherToCompare.UseASCELy, changes, categoryName, "Use ASCE Ly"), Equals, False)
        Equals = If(Me.CalcBracingForces.CheckChange(otherToCompare.CalcBracingForces, changes, categoryName, "Calc Bracing Forces"), Equals, False)
        Equals = If(Me.IgnoreBracingFEA.CheckChange(otherToCompare.IgnoreBracingFEA, changes, categoryName, "Ignore Bracing FEA"), Equals, False)
        Equals = If(Me.BypassStabilityChecks.CheckChange(otherToCompare.BypassStabilityChecks, changes, categoryName, "Bypass Stability Checks"), Equals, False)
        Equals = If(Me.UseWindProjection.CheckChange(otherToCompare.UseWindProjection, changes, categoryName, "Use Wind Projection"), Equals, False)
        Equals = If(Me.UseDishCoeff.CheckChange(otherToCompare.UseDishCoeff, changes, categoryName, "Use Dish Coeff"), Equals, False)
        Equals = If(Me.AutoCalcTorqArmArea.CheckChange(otherToCompare.AutoCalcTorqArmArea, changes, categoryName, "Auto Calc Torq Arm Area"), Equals, False)
        Equals = If(Me.foundationStiffness.CheckChange(otherToCompare.foundationStiffness, changes, categoryName, "Foundation Stiffness"), Equals, False)
        Equals = If(Me.defaultGirtOffsets.CheckChange(otherToCompare.defaultGirtOffsets, changes, categoryName, "Default Girt Offsets"), Equals, False)
        Equals = If(Me.cantileverPoles.CheckChange(otherToCompare.cantileverPoles, changes, categoryName, "Cantilever Poles"), Equals, False)
        Equals = If(Me.windDirections.CheckChange(otherToCompare.windDirections, changes, categoryName, "Wind Directions"), Equals, False)
        Equals = If(Me.misclOptions.CheckChange(otherToCompare.misclOptions, changes, categoryName, "Miscl"), Equals, False)

        Return Equals
    End Function
End Class
<DataContract()>
Partial Public Class tnxFoundationStiffness
    Inherits EDSObject

    Public Overrides ReadOnly Property EDSObjectName As String = "Foundation Stiffness"

#Region "Define"

    Private _MastVert As Double?
    Private _MastHorz As Double?
    Private _GuyVert As Double?
    Private _GuyHorz As Double?

    <Category("TNX Foundation Stiffness Options"), Description("foundation stiffness"), DisplayName("MastVert")>
     <DataMember()> Public Property MastVert() As Double?
        Get
            Return Me._MastVert
        End Get
        Set
            Me._MastVert = Value
        End Set
    End Property
    <Category("TNX Foundation Stiffness Options"), Description("foundation stiffness"), DisplayName("MastHorz")>
     <DataMember()> Public Property MastHorz() As Double?
        Get
            Return Me._MastHorz
        End Get
        Set
            Me._MastHorz = Value
        End Set
    End Property
    <Category("TNX Foundation Stiffness Options"), Description("foundation stiffness"), DisplayName("GuyVert")>
     <DataMember()> Public Property GuyVert() As Double?
        Get
            Return Me._GuyVert
        End Get
        Set
            Me._GuyVert = Value
        End Set
    End Property
    <Category("TNX Foundation Stiffness Options"), Description("foundation stiffness"), DisplayName("GuyHorz")>
     <DataMember()> Public Property GuyHorz() As Double?
        Get
            Return Me._GuyHorz
        End Get
        Set
            Me._GuyHorz = Value
        End Set
    End Property

#End Region

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As tnxFoundationStiffness = TryCast(other, tnxFoundationStiffness)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.MastVert.CheckChange(otherToCompare.MastVert, changes, categoryName, "Mast Vert"), Equals, False)
        Equals = If(Me.MastHorz.CheckChange(otherToCompare.MastHorz, changes, categoryName, "Mast Horz"), Equals, False)
        Equals = If(Me.GuyVert.CheckChange(otherToCompare.GuyVert, changes, categoryName, "Guy Vert"), Equals, False)
        Equals = If(Me.GuyHorz.CheckChange(otherToCompare.GuyHorz, changes, categoryName, "Guy Horz"), Equals, False)

        Return Equals
    End Function

End Class
<DataContract()>
Partial Public Class tnxDefaultGirtOffsets
    Inherits EDSObject

    Public Overrides ReadOnly Property EDSObjectName As String = "Default Girt Offsets"

#Region "Define"

    Private _GirtOffset As Double?
    Private _GirtOffsetLatticedPole As Double?
    Private _OffsetBotGirt As Boolean?

    <Category("TNX Default Girt Offset Options"), Description(""), DisplayName("GirtOffset")>
     <DataMember()> Public Property GirtOffset() As Double?
        Get
            Return Me._GirtOffset
        End Get
        Set
            Me._GirtOffset = Value
        End Set
    End Property
    <Category("TNX Default Girt Offset Options"), Description(""), DisplayName("GirtOffsetLatticedPole")>
     <DataMember()> Public Property GirtOffsetLatticedPole() As Double?
        Get
            Return Me._GirtOffsetLatticedPole
        End Get
        Set
            Me._GirtOffsetLatticedPole = Value
        End Set
    End Property
    <Category("TNX Default Girt Offset Options"), Description("offset girt at foundation"), DisplayName("OffsetBotGirt")>
     <DataMember()> Public Property OffsetBotGirt() As Boolean?
        Get
            Return Me._OffsetBotGirt
        End Get
        Set
            Me._OffsetBotGirt = Value
        End Set
    End Property
#End Region

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As tnxDefaultGirtOffsets = TryCast(other, tnxDefaultGirtOffsets)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.GirtOffset.CheckChange(otherToCompare.GirtOffset, changes, categoryName, "Girt Offset"), Equals, False)
        Equals = If(Me.GirtOffsetLatticedPole.CheckChange(otherToCompare.GirtOffsetLatticedPole, changes, categoryName, "Girt Offset Latticed Pole"), Equals, False)
        Equals = If(Me.OffsetBotGirt.CheckChange(otherToCompare.OffsetBotGirt, changes, categoryName, "Offset Bot Girt"), Equals, False)

        Return Equals
    End Function
End Class
<DataContract()>
Partial Public Class tnxCantileverPoles
    Inherits EDSObject

    Public Overrides ReadOnly Property EDSObjectName As String = "Cantilever Poles"

#Region "Define"

    Private _CheckVonMises As Boolean?
    Private _SocketTopMount As Boolean?
    Private _PrintMonopoleAtIncrements As Boolean?
    Private _UseSubCriticalFlow As Boolean?
    Private _AssumePoleWithNoAttachments As Boolean?
    Private _AssumePoleWithShroud As Boolean?
    Private _PoleCornerRadiusKnown As Boolean?
    Private _CantKFactor As Double?

    <Category("TNX Cantilever Pole Options"), Description(")Include Shear-Torsion Interaction"), DisplayName("CheckVonMises")>
     <DataMember()> Public Property CheckVonMises() As Boolean?
        Get
            Return Me._CheckVonMises
        End Get
        Set
            Me._CheckVonMises = Value
        End Set
    End Property
    <Category("TNX Cantilever Pole Options"), Description("Use Top Mounted Socket"), DisplayName("SocketTopMount")>
     <DataMember()> Public Property SocketTopMount() As Boolean?
        Get
            Return Me._SocketTopMount
        End Get
        Set
            Me._SocketTopMount = Value
        End Set
    End Property
    <Category("TNX Cantilever Pole Options"), Description("Print Pole Stresses at Increments"), DisplayName("PrintMonopoleAtIncrements")>
     <DataMember()> Public Property PrintMonopoleAtIncrements() As Boolean?
        Get
            Return Me._PrintMonopoleAtIncrements
        End Get
        Set
            Me._PrintMonopoleAtIncrements = Value
        End Set
    End Property
    <Category("TNX Cantilever Pole Options"), Description("Always Yse Sub-Critical Flow"), DisplayName("UseSubCriticalFlow")>
     <DataMember()> Public Property UseSubCriticalFlow() As Boolean?
        Get
            Return Me._UseSubCriticalFlow
        End Get
        Set
            Me._UseSubCriticalFlow = Value
        End Set
    End Property
    <Category("TNX Cantilever Pole Options"), Description("Pole Without Linear Attachments"), DisplayName("AssumePoleWithNoAttachments")>
     <DataMember()> Public Property AssumePoleWithNoAttachments() As Boolean?
        Get
            Return Me._AssumePoleWithNoAttachments
        End Get
        Set
            Me._AssumePoleWithNoAttachments = Value
        End Set
    End Property
    <Category("TNX Cantilever Pole Options"), Description("Pole With Shroud or No Appurtenances"), DisplayName("AssumePoleWithShroud")>
     <DataMember()> Public Property AssumePoleWithShroud() As Boolean?
        Get
            Return Me._AssumePoleWithShroud
        End Get
        Set
            Me._AssumePoleWithShroud = Value
        End Set
    End Property
    <Category("TNX Cantilever Pole Options"), Description("Outside and Inside Corner Radii Are Known"), DisplayName("PoleCornerRadiusKnown")>
     <DataMember()> Public Property PoleCornerRadiusKnown() As Boolean?
        Get
            Return Me._PoleCornerRadiusKnown
        End Get
        Set
            Me._PoleCornerRadiusKnown = Value
        End Set
    End Property
    <Category("TNX Cantilever Pole Options"), Description("Cantilevered Poles K Factor"), DisplayName("CantKFactor")>
     <DataMember()> Public Property CantKFactor() As Double?
        Get
            Return Me._CantKFactor
        End Get
        Set
            Me._CantKFactor = Value
        End Set
    End Property
#End Region

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As tnxCantileverPoles = TryCast(other, tnxCantileverPoles)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.CheckVonMises.CheckChange(otherToCompare.CheckVonMises, changes, categoryName, "Check Von Mises"), Equals, False)
        Equals = If(Me.SocketTopMount.CheckChange(otherToCompare.SocketTopMount, changes, categoryName, "Socket Top Mount"), Equals, False)
        Equals = If(Me.PrintMonopoleAtIncrements.CheckChange(otherToCompare.PrintMonopoleAtIncrements, changes, categoryName, "Print Monopole At Increments"), Equals, False)
        Equals = If(Me.UseSubCriticalFlow.CheckChange(otherToCompare.UseSubCriticalFlow, changes, categoryName, "Use Sub Critical Flow"), Equals, False)
        Equals = If(Me.AssumePoleWithNoAttachments.CheckChange(otherToCompare.AssumePoleWithNoAttachments, changes, categoryName, "Assume Pole With No Attachments"), Equals, False)
        Equals = If(Me.AssumePoleWithShroud.CheckChange(otherToCompare.AssumePoleWithShroud, changes, categoryName, "Assume Pole With Shroud"), Equals, False)
        Equals = If(Me.PoleCornerRadiusKnown.CheckChange(otherToCompare.PoleCornerRadiusKnown, changes, categoryName, "Pole Corner Radius Known"), Equals, False)
        Equals = If(Me.CantKFactor.CheckChange(otherToCompare.CantKFactor, changes, categoryName, "Cant KFactor"), Equals, False)

        Return Equals
    End Function
End Class
<DataContract()>
Partial Public Class tnxWindDirections
    Inherits EDSObject

    Public Overrides ReadOnly Property EDSObjectName As String = "Wind Directions"

#Region "Define"

    Private _WindDirOption As Integer?
    Private _WindDir0_0 As Boolean?
    Private _WindDir0_1 As Boolean?
    Private _WindDir0_2 As Boolean?
    Private _WindDir0_3 As Boolean?
    Private _WindDir0_4 As Boolean?
    Private _WindDir0_5 As Boolean?
    Private _WindDir0_6 As Boolean?
    Private _WindDir0_7 As Boolean?
    Private _WindDir0_8 As Boolean?
    Private _WindDir0_9 As Boolean?
    Private _WindDir0_10 As Boolean?
    Private _WindDir0_11 As Boolean?
    Private _WindDir0_12 As Boolean?
    Private _WindDir0_13 As Boolean?
    Private _WindDir0_14 As Boolean?
    Private _WindDir0_15 As Boolean?
    Private _WindDir1_0 As Boolean?
    Private _WindDir1_1 As Boolean?
    Private _WindDir1_2 As Boolean?
    Private _WindDir1_3 As Boolean?
    Private _WindDir1_4 As Boolean?
    Private _WindDir1_5 As Boolean?
    Private _WindDir1_6 As Boolean?
    Private _WindDir1_7 As Boolean?
    Private _WindDir1_8 As Boolean?
    Private _WindDir1_9 As Boolean?
    Private _WindDir1_10 As Boolean?
    Private _WindDir1_11 As Boolean?
    Private _WindDir1_12 As Boolean?
    Private _WindDir1_13 As Boolean?
    Private _WindDir1_14 As Boolean?
    Private _WindDir1_15 As Boolean?
    Private _WindDir2_0 As Boolean?
    Private _WindDir2_1 As Boolean?
    Private _WindDir2_2 As Boolean?
    Private _WindDir2_3 As Boolean?
    Private _WindDir2_4 As Boolean?
    Private _WindDir2_5 As Boolean?
    Private _WindDir2_6 As Boolean?
    Private _WindDir2_7 As Boolean?
    Private _WindDir2_8 As Boolean?
    Private _WindDir2_9 As Boolean?
    Private _WindDir2_10 As Boolean?
    Private _WindDir2_11 As Boolean?
    Private _WindDir2_12 As Boolean?
    Private _WindDir2_13 As Boolean?
    Private _WindDir2_14 As Boolean?
    Private _WindDir2_15 As Boolean?
    Private _SuppressWindPatternLoading As Boolean?

    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Direction Options")>
     <DataMember()> Public Property WindDirOption() As Integer?
        Get
            Return Me._WindDirOption
        End Get
        Set
            Me._WindDirOption = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind No Ice 0 deg")>
     <DataMember()> Public Property WindDir0_0() As Boolean?
        Get
            Return Me._WindDir0_0
        End Get
        Set
            Me._WindDir0_0 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind No Ice 30 deg")>
     <DataMember()> Public Property WindDir0_1() As Boolean?
        Get
            Return Me._WindDir0_1
        End Get
        Set
            Me._WindDir0_1 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind No Ice 45 deg")>
     <DataMember()> Public Property WindDir0_2() As Boolean?
        Get
            Return Me._WindDir0_2
        End Get
        Set
            Me._WindDir0_2 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind No Ice 60 deg")>
     <DataMember()> Public Property WindDir0_3() As Boolean?
        Get
            Return Me._WindDir0_3
        End Get
        Set
            Me._WindDir0_3 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind No Ice 90 deg")>
     <DataMember()> Public Property WindDir0_4() As Boolean?
        Get
            Return Me._WindDir0_4
        End Get
        Set
            Me._WindDir0_4 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind No Ice 120 deg")>
     <DataMember()> Public Property WindDir0_5() As Boolean?
        Get
            Return Me._WindDir0_5
        End Get
        Set
            Me._WindDir0_5 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind No Ice 135 deg")>
     <DataMember()> Public Property WindDir0_6() As Boolean?
        Get
            Return Me._WindDir0_6
        End Get
        Set
            Me._WindDir0_6 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind No Ice 150 deg")>
     <DataMember()> Public Property WindDir0_7() As Boolean?
        Get
            Return Me._WindDir0_7
        End Get
        Set
            Me._WindDir0_7 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind No Ice 180 deg")>
     <DataMember()> Public Property WindDir0_8() As Boolean?
        Get
            Return Me._WindDir0_8
        End Get
        Set
            Me._WindDir0_8 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind No Ice 210 deg")>
     <DataMember()> Public Property WindDir0_9() As Boolean?
        Get
            Return Me._WindDir0_9
        End Get
        Set
            Me._WindDir0_9 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind No Ice 225 deg")>
     <DataMember()> Public Property WindDir0_10() As Boolean?
        Get
            Return Me._WindDir0_10
        End Get
        Set
            Me._WindDir0_10 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind No Ice 240 deg")>
     <DataMember()> Public Property WindDir0_11() As Boolean?
        Get
            Return Me._WindDir0_11
        End Get
        Set
            Me._WindDir0_11 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind No Ice 270 deg")>
     <DataMember()> Public Property WindDir0_12() As Boolean?
        Get
            Return Me._WindDir0_12
        End Get
        Set
            Me._WindDir0_12 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind No Ice 300 deg")>
     <DataMember()> Public Property WindDir0_13() As Boolean?
        Get
            Return Me._WindDir0_13
        End Get
        Set
            Me._WindDir0_13 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind No Ice 315 deg")>
     <DataMember()> Public Property WindDir0_14() As Boolean?
        Get
            Return Me._WindDir0_14
        End Get
        Set
            Me._WindDir0_14 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind No Ice 330 deg")>
     <DataMember()> Public Property WindDir0_15() As Boolean?
        Get
            Return Me._WindDir0_15
        End Get
        Set
            Me._WindDir0_15 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Ice 0 deg")>
     <DataMember()> Public Property WindDir1_0() As Boolean?
        Get
            Return Me._WindDir1_0
        End Get
        Set
            Me._WindDir1_0 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Ice 30 deg")>
     <DataMember()> Public Property WindDir1_1() As Boolean?
        Get
            Return Me._WindDir1_1
        End Get
        Set
            Me._WindDir1_1 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Ice 45 deg")>
     <DataMember()> Public Property WindDir1_2() As Boolean?
        Get
            Return Me._WindDir1_2
        End Get
        Set
            Me._WindDir1_2 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Ice 60 deg")>
     <DataMember()> Public Property WindDir1_3() As Boolean?
        Get
            Return Me._WindDir1_3
        End Get
        Set
            Me._WindDir1_3 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Ice 90 deg")>
     <DataMember()> Public Property WindDir1_4() As Boolean?
        Get
            Return Me._WindDir1_4
        End Get
        Set
            Me._WindDir1_4 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Ice 120 deg")>
     <DataMember()> Public Property WindDir1_5() As Boolean?
        Get
            Return Me._WindDir1_5
        End Get
        Set
            Me._WindDir1_5 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Ice 135 deg")>
     <DataMember()> Public Property WindDir1_6() As Boolean?
        Get
            Return Me._WindDir1_6
        End Get
        Set
            Me._WindDir1_6 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Ice 150 deg")>
     <DataMember()> Public Property WindDir1_7() As Boolean?
        Get
            Return Me._WindDir1_7
        End Get
        Set
            Me._WindDir1_7 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Ice 180 deg")>
     <DataMember()> Public Property WindDir1_8() As Boolean?
        Get
            Return Me._WindDir1_8
        End Get
        Set
            Me._WindDir1_8 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Ice 210 deg")>
     <DataMember()> Public Property WindDir1_9() As Boolean?
        Get
            Return Me._WindDir1_9
        End Get
        Set
            Me._WindDir1_9 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Ice 225 deg")>
     <DataMember()> Public Property WindDir1_10() As Boolean?
        Get
            Return Me._WindDir1_10
        End Get
        Set
            Me._WindDir1_10 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Ice 240 deg")>
     <DataMember()> Public Property WindDir1_11() As Boolean?
        Get
            Return Me._WindDir1_11
        End Get
        Set
            Me._WindDir1_11 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Ice 270 deg")>
     <DataMember()> Public Property WindDir1_12() As Boolean?
        Get
            Return Me._WindDir1_12
        End Get
        Set
            Me._WindDir1_12 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Ice 300 deg")>
     <DataMember()> Public Property WindDir1_13() As Boolean?
        Get
            Return Me._WindDir1_13
        End Get
        Set
            Me._WindDir1_13 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Ice 315 deg")>
     <DataMember()> Public Property WindDir1_14() As Boolean?
        Get
            Return Me._WindDir1_14
        End Get
        Set
            Me._WindDir1_14 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Ice 330 deg")>
     <DataMember()> Public Property WindDir1_15() As Boolean?
        Get
            Return Me._WindDir1_15
        End Get
        Set
            Me._WindDir1_15 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Service 0 deg")>
     <DataMember()> Public Property WindDir2_0() As Boolean?
        Get
            Return Me._WindDir2_0
        End Get
        Set
            Me._WindDir2_0 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Service 30 deg")>
     <DataMember()> Public Property WindDir2_1() As Boolean?
        Get
            Return Me._WindDir2_1
        End Get
        Set
            Me._WindDir2_1 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Service 45 deg")>
     <DataMember()> Public Property WindDir2_2() As Boolean?
        Get
            Return Me._WindDir2_2
        End Get
        Set
            Me._WindDir2_2 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Service 60 deg")>
     <DataMember()> Public Property WindDir2_3() As Boolean?
        Get
            Return Me._WindDir2_3
        End Get
        Set
            Me._WindDir2_3 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Service 90 deg")>
     <DataMember()> Public Property WindDir2_4() As Boolean?
        Get
            Return Me._WindDir2_4
        End Get
        Set
            Me._WindDir2_4 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Service 120 deg")>
     <DataMember()> Public Property WindDir2_5() As Boolean?
        Get
            Return Me._WindDir2_5
        End Get
        Set
            Me._WindDir2_5 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Service 135 deg")>
     <DataMember()> Public Property WindDir2_6() As Boolean?
        Get
            Return Me._WindDir2_6
        End Get
        Set
            Me._WindDir2_6 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Service 150 deg")>
     <DataMember()> Public Property WindDir2_7() As Boolean?
        Get
            Return Me._WindDir2_7
        End Get
        Set
            Me._WindDir2_7 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Service 180 deg")>
     <DataMember()> Public Property WindDir2_8() As Boolean?
        Get
            Return Me._WindDir2_8
        End Get
        Set
            Me._WindDir2_8 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Service 210 deg")>
     <DataMember()> Public Property WindDir2_9() As Boolean?
        Get
            Return Me._WindDir2_9
        End Get
        Set
            Me._WindDir2_9 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Service 225 deg")>
     <DataMember()> Public Property WindDir2_10() As Boolean?
        Get
            Return Me._WindDir2_10
        End Get
        Set
            Me._WindDir2_10 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Service 240 deg")>
     <DataMember()> Public Property WindDir2_11() As Boolean?
        Get
            Return Me._WindDir2_11
        End Get
        Set
            Me._WindDir2_11 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Service 270 deg")>
     <DataMember()> Public Property WindDir2_12() As Boolean?
        Get
            Return Me._WindDir2_12
        End Get
        Set
            Me._WindDir2_12 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Service 300 deg")>
     <DataMember()> Public Property WindDir2_13() As Boolean?
        Get
            Return Me._WindDir2_13
        End Get
        Set
            Me._WindDir2_13 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Service 315 deg")>
     <DataMember()> Public Property WindDir2_14() As Boolean?
        Get
            Return Me._WindDir2_14
        End Get
        Set
            Me._WindDir2_14 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Wind Service 330 deg")>
     <DataMember()> Public Property WindDir2_15() As Boolean?
        Get
            Return Me._WindDir2_15
        End Get
        Set
            Me._WindDir2_15 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description(""), DisplayName("Suppress Generation of Pattern Loading")>
     <DataMember()> Public Property SuppressWindPatternLoading() As Boolean?
        Get
            Return Me._SuppressWindPatternLoading
        End Get
        Set
            Me._SuppressWindPatternLoading = Value
        End Set
    End Property

#End Region

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As tnxWindDirections = TryCast(other, tnxWindDirections)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.WindDirOption.CheckChange(otherToCompare.WindDirOption, changes, categoryName, "Wind Direction"), Equals, False)
        Equals = If(Me.WindDir0_0.CheckChange(otherToCompare.WindDir0_0, changes, categoryName, "Wind No Ice 0 deg"), Equals, False)
        Equals = If(Me.WindDir0_1.CheckChange(otherToCompare.WindDir0_1, changes, categoryName, "Wind No Ice 30 deg"), Equals, False)
        Equals = If(Me.WindDir0_2.CheckChange(otherToCompare.WindDir0_2, changes, categoryName, "Wind No Ice 45 deg"), Equals, False)
        Equals = If(Me.WindDir0_3.CheckChange(otherToCompare.WindDir0_3, changes, categoryName, "Wind No Ice 60 deg"), Equals, False)
        Equals = If(Me.WindDir0_4.CheckChange(otherToCompare.WindDir0_4, changes, categoryName, "Wind No Ice 90 deg"), Equals, False)
        Equals = If(Me.WindDir0_5.CheckChange(otherToCompare.WindDir0_5, changes, categoryName, "Wind No Ice 120 deg"), Equals, False)
        Equals = If(Me.WindDir0_6.CheckChange(otherToCompare.WindDir0_6, changes, categoryName, "Wind No Ice 135 deg"), Equals, False)
        Equals = If(Me.WindDir0_7.CheckChange(otherToCompare.WindDir0_7, changes, categoryName, "Wind No Ice 150 deg"), Equals, False)
        Equals = If(Me.WindDir0_8.CheckChange(otherToCompare.WindDir0_8, changes, categoryName, "Wind No Ice 180 deg"), Equals, False)
        Equals = If(Me.WindDir0_9.CheckChange(otherToCompare.WindDir0_9, changes, categoryName, "Wind No Ice 210 deg"), Equals, False)
        Equals = If(Me.WindDir0_10.CheckChange(otherToCompare.WindDir0_10, changes, categoryName, "Wind No Ice 225 deg"), Equals, False)
        Equals = If(Me.WindDir0_11.CheckChange(otherToCompare.WindDir0_11, changes, categoryName, "Wind No Ice 240 deg"), Equals, False)
        Equals = If(Me.WindDir0_12.CheckChange(otherToCompare.WindDir0_12, changes, categoryName, "Wind No Ice 270 deg"), Equals, False)
        Equals = If(Me.WindDir0_13.CheckChange(otherToCompare.WindDir0_13, changes, categoryName, "Wind No Ice 300 deg"), Equals, False)
        Equals = If(Me.WindDir0_14.CheckChange(otherToCompare.WindDir0_14, changes, categoryName, "Wind No Ice 315 deg"), Equals, False)
        Equals = If(Me.WindDir0_15.CheckChange(otherToCompare.WindDir0_15, changes, categoryName, "Wind No Ice 330 deg"), Equals, False)
        Equals = If(Me.WindDir1_0.CheckChange(otherToCompare.WindDir1_0, changes, categoryName, "Wind Ice 0 deg"), Equals, False)
        Equals = If(Me.WindDir1_1.CheckChange(otherToCompare.WindDir1_1, changes, categoryName, "Wind Ice 30 deg"), Equals, False)
        Equals = If(Me.WindDir1_2.CheckChange(otherToCompare.WindDir1_2, changes, categoryName, "Wind Ice 45 deg"), Equals, False)
        Equals = If(Me.WindDir1_3.CheckChange(otherToCompare.WindDir1_3, changes, categoryName, "Wind Ice 60 deg"), Equals, False)
        Equals = If(Me.WindDir1_4.CheckChange(otherToCompare.WindDir1_4, changes, categoryName, "Wind Ice 90 deg"), Equals, False)
        Equals = If(Me.WindDir1_5.CheckChange(otherToCompare.WindDir1_5, changes, categoryName, "Wind Ice 120 deg"), Equals, False)
        Equals = If(Me.WindDir1_6.CheckChange(otherToCompare.WindDir1_6, changes, categoryName, "Wind Ice 135 deg"), Equals, False)
        Equals = If(Me.WindDir1_7.CheckChange(otherToCompare.WindDir1_7, changes, categoryName, "Wind Ice 150 deg"), Equals, False)
        Equals = If(Me.WindDir1_8.CheckChange(otherToCompare.WindDir1_8, changes, categoryName, "Wind Ice 180 deg"), Equals, False)
        Equals = If(Me.WindDir1_9.CheckChange(otherToCompare.WindDir1_9, changes, categoryName, "Wind Ice 210 deg"), Equals, False)
        Equals = If(Me.WindDir1_10.CheckChange(otherToCompare.WindDir1_10, changes, categoryName, "Wind Ice 225 deg"), Equals, False)
        Equals = If(Me.WindDir1_11.CheckChange(otherToCompare.WindDir1_11, changes, categoryName, "Wind Ice 240 deg"), Equals, False)
        Equals = If(Me.WindDir1_12.CheckChange(otherToCompare.WindDir1_12, changes, categoryName, "Wind Ice 270 deg"), Equals, False)
        Equals = If(Me.WindDir1_13.CheckChange(otherToCompare.WindDir1_13, changes, categoryName, "Wind Ice 300 deg"), Equals, False)
        Equals = If(Me.WindDir1_14.CheckChange(otherToCompare.WindDir1_14, changes, categoryName, "Wind Ice 315 deg"), Equals, False)
        Equals = If(Me.WindDir1_15.CheckChange(otherToCompare.WindDir1_15, changes, categoryName, "Wind Ice 330 deg"), Equals, False)
        Equals = If(Me.WindDir2_0.CheckChange(otherToCompare.WindDir2_0, changes, categoryName, "Wind Service 0 deg"), Equals, False)
        Equals = If(Me.WindDir2_1.CheckChange(otherToCompare.WindDir2_1, changes, categoryName, "Wind Service 30 deg"), Equals, False)
        Equals = If(Me.WindDir2_2.CheckChange(otherToCompare.WindDir2_2, changes, categoryName, "Wind Service 45 deg"), Equals, False)
        Equals = If(Me.WindDir2_3.CheckChange(otherToCompare.WindDir2_3, changes, categoryName, "Wind Service 60 deg"), Equals, False)
        Equals = If(Me.WindDir2_4.CheckChange(otherToCompare.WindDir2_4, changes, categoryName, "Wind Service 90 deg"), Equals, False)
        Equals = If(Me.WindDir2_5.CheckChange(otherToCompare.WindDir2_5, changes, categoryName, "Wind Service 120 deg"), Equals, False)
        Equals = If(Me.WindDir2_6.CheckChange(otherToCompare.WindDir2_6, changes, categoryName, "Wind Service 135 deg"), Equals, False)
        Equals = If(Me.WindDir2_7.CheckChange(otherToCompare.WindDir2_7, changes, categoryName, "Wind Service 150 deg"), Equals, False)
        Equals = If(Me.WindDir2_8.CheckChange(otherToCompare.WindDir2_8, changes, categoryName, "Wind Service 180 deg"), Equals, False)
        Equals = If(Me.WindDir2_9.CheckChange(otherToCompare.WindDir2_9, changes, categoryName, "Wind Service 210 deg"), Equals, False)
        Equals = If(Me.WindDir2_10.CheckChange(otherToCompare.WindDir2_10, changes, categoryName, "Wind Service 225 deg"), Equals, False)
        Equals = If(Me.WindDir2_11.CheckChange(otherToCompare.WindDir2_11, changes, categoryName, "Wind Service 240 deg"), Equals, False)
        Equals = If(Me.WindDir2_12.CheckChange(otherToCompare.WindDir2_12, changes, categoryName, "Wind Service 270 deg"), Equals, False)
        Equals = If(Me.WindDir2_13.CheckChange(otherToCompare.WindDir2_13, changes, categoryName, "Wind Service 300 deg"), Equals, False)
        Equals = If(Me.WindDir2_14.CheckChange(otherToCompare.WindDir2_14, changes, categoryName, "Wind Service 315 deg"), Equals, False)
        Equals = If(Me.WindDir2_15.CheckChange(otherToCompare.WindDir2_15, changes, categoryName, "Wind Service 330 deg"), Equals, False)
        Equals = If(Me.SuppressWindPatternLoading.CheckChange(otherToCompare.SuppressWindPatternLoading, changes, categoryName, "Suppress Wind Pattern Loading"), Equals, False)


        Return Equals
    End Function
End Class
<DataContract()>
Partial Public Class tnxMisclOptions
    Inherits EDSObject

    Public Overrides ReadOnly Property EDSObjectName As String = "Miscl"

#Region "Define"

    Private _HogRodTakeup As Double?
    Private _RadiusSampleDist As Double?

    <Category("TNX Miscl Options"), Description("Tension Only Take-Up"), DisplayName("Hog Rod Takeup")>
     <DataMember()> Public Property HogRodTakeup() As Double?
        Get
            Return Me._HogRodTakeup
        End Get
        Set
            Me._HogRodTakeup = Value
        End Set
    End Property
    <Category("TNX Miscl Options"), Description("Sampling Distance"), DisplayName("Radius Sample Dist")>
     <DataMember()> Public Property RadiusSampleDist() As Double?
        Get
            Return Me._RadiusSampleDist
        End Get
        Set
            Me._RadiusSampleDist = Value
        End Set
    End Property
#End Region
    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As tnxMisclOptions = TryCast(other, tnxMisclOptions)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.HogRodTakeup.CheckChange(otherToCompare.HogRodTakeup, changes, categoryName, "Tension Only Take-Up"), Equals, False)
        Equals = If(Me.RadiusSampleDist.CheckChange(otherToCompare.RadiusSampleDist, changes, categoryName, "Sampling Distance"), Equals, False)

        Return Equals
    End Function
End Class



#End Region

#Region "Settings"
<DataContract()>
Partial Public Class tnxSettings
    Inherits EDSObject

    Public Overrides ReadOnly Property EDSObjectName As String = "Settings"

#Region "Define"

    'Other settings are not saved in ERI file
    Private _USUnits As New tnxUnits()
    'Private _SIunits As tnxSIUnits 
    Private _projectInfo As New tnxProjectInfo()
    Private _userInfo As New tnxUserInfo()

    <Category("TNX Setings"), Description(""), DisplayName("US Units")>
     <DataMember()> Public Property USUnits() As tnxUnits
        Get
            Return Me._USUnits
        End Get
        Set
            Me._USUnits = Value
        End Set
    End Property
    <Category("TNX Setings"), Description(""), DisplayName("Project Info")>
     <DataMember()> Public Property projectInfo() As tnxProjectInfo
        Get
            Return Me._projectInfo
        End Get
        Set
            Me._projectInfo = Value
        End Set
    End Property
    <Category("TNX Setings"), Description(""), DisplayName("User Info")>
     <DataMember()> Public Property userInfo() As tnxUserInfo
        Get
            Return Me._userInfo
        End Get
        Set
            Me._userInfo = Value
        End Set
    End Property
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
        Dim otherToCompare As tnxSettings = TryCast(other, tnxSettings)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.USUnits.CheckChange(otherToCompare.USUnits, changes, categoryName, "US Units"), Equals, False)
        Equals = If(Me.projectInfo.CheckChange(otherToCompare.projectInfo, changes, categoryName, "Sampling Distance"), Equals, False)
        Equals = If(Me.userInfo.CheckChange(otherToCompare.userInfo, changes, categoryName, "Sampling Distance"), Equals, False)

        Return Equals
    End Function
End Class
<DataContract()>
Partial Public Class tnxSolutionSettings
    Inherits EDSObject

    Public Overrides ReadOnly Property EDSObjectName As String = "Solution"

#Region "Define"

    Private _SolutionUsePDelta As Boolean?
    Private _SolutionMinStiffness As Double?
    Private _SolutionMaxStiffness As Double?
    Private _SolutionMaxCycles As Integer?
    Private _SolutionPower As Double?
    Private _SolutionTolerance As Double?

    <Category("TNX Solution Options"), Description(""), DisplayName("SolutionUsePDelta")>
     <DataMember()> Public Property SolutionUsePDelta() As Boolean?
        Get
            Return Me._SolutionUsePDelta
        End Get
        Set
            Me._SolutionUsePDelta = Value
        End Set
    End Property
    <Category("TNX Solution Options"), Description(""), DisplayName("SolutionMinStiffness")>
     <DataMember()> Public Property SolutionMinStiffness() As Double?
        Get
            Return Me._SolutionMinStiffness
        End Get
        Set
            Me._SolutionMinStiffness = Value
        End Set
    End Property
    <Category("TNX Solution Options"), Description(""), DisplayName("SolutionMaxStiffness")>
     <DataMember()> Public Property SolutionMaxStiffness() As Double?
        Get
            Return Me._SolutionMaxStiffness
        End Get
        Set
            Me._SolutionMaxStiffness = Value
        End Set
    End Property
    <Category("TNX Solution Options"), Description(""), DisplayName("SolutionMaxCycles")>
     <DataMember()> Public Property SolutionMaxCycles() As Integer?
        Get
            Return Me._SolutionMaxCycles
        End Get
        Set
            Me._SolutionMaxCycles = Value
        End Set
    End Property
    <Category("TNX Solution Options"), Description(""), DisplayName("SolutionPower")>
     <DataMember()> Public Property SolutionPower() As Double?
        Get
            Return Me._SolutionPower
        End Get
        Set
            Me._SolutionPower = Value
        End Set
    End Property
    <Category("TNX Solution Options"), Description(""), DisplayName("SolutionTolerance")>
     <DataMember()> Public Property SolutionTolerance() As Double?
        Get
            Return Me._SolutionTolerance
        End Get
        Set
            Me._SolutionTolerance = Value
        End Set
    End Property
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
        Dim otherToCompare As tnxSolutionSettings = TryCast(other, tnxSolutionSettings)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.SolutionUsePDelta.CheckChange(otherToCompare.SolutionUsePDelta, changes, categoryName, "Use PDelta"), Equals, False)
        Equals = If(Me.SolutionMinStiffness.CheckChange(otherToCompare.SolutionMinStiffness, changes, categoryName, "Min Stiffness"), Equals, False)
        Equals = If(Me.SolutionMaxStiffness.CheckChange(otherToCompare.SolutionMaxStiffness, changes, categoryName, "Max Stiffness"), Equals, False)
        Equals = If(Me.SolutionMaxCycles.CheckChange(otherToCompare.SolutionMaxCycles, changes, categoryName, "Max Cycles"), Equals, False)
        Equals = If(Me.SolutionPower.CheckChange(otherToCompare.SolutionPower, changes, categoryName, "Power"), Equals, False)
        Equals = If(Me.SolutionTolerance.CheckChange(otherToCompare.SolutionTolerance, changes, categoryName, "Tolerance"), Equals, False)

        Return Equals
    End Function
End Class
<DataContract()>
Partial Public Class tnxReportSettings
    Inherits EDSObject

    Public Overrides ReadOnly Property EDSObjectName As String = "Report"

#Region "Define"

    Private _ReportInputCosts As Boolean?
    Private _ReportInputGeometry As Boolean?
    Private _ReportInputOptions As Boolean?
    Private _ReportMaxForces As Boolean?
    Private _ReportInputMap As Boolean?
    Private _CostReportOutputType As String
    Private _CapacityReportOutputType As String
    Private _ReportPrintForceTotals As Boolean?
    Private _ReportPrintForceDetails As Boolean?
    Private _ReportPrintMastVectors As Boolean?
    Private _ReportPrintAntPoleVectors As Boolean?
    Private _ReportPrintDiscreteVectors As Boolean?
    Private _ReportPrintDishVectors As Boolean?
    Private _ReportPrintFeedTowerVectors As Boolean?
    Private _ReportPrintUserLoadVectors As Boolean?
    Private _ReportPrintPressures As Boolean?
    Private _ReportPrintAppurtForces As Boolean?
    Private _ReportPrintGuyForces As Boolean?
    Private _ReportPrintGuyStressing As Boolean?
    Private _ReportPrintDeflections As Boolean?
    Private _ReportPrintReactions As Boolean?
    Private _ReportPrintStressChecks As Boolean?
    Private _ReportPrintBoltChecks As Boolean?
    Private _ReportPrintInputGVerificationTables As Boolean?
    Private _ReportPrintOutputGVerificationTables As Boolean?

    <Category("TNX Report Settings"), Description(""), DisplayName("ReportInputCosts")>
     <DataMember()> Public Property ReportInputCosts() As Boolean?
        Get
            Return Me._ReportInputCosts
        End Get
        Set
            Me._ReportInputCosts = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportInputGeometry")>
     <DataMember()> Public Property ReportInputGeometry() As Boolean?
        Get
            Return Me._ReportInputGeometry
        End Get
        Set
            Me._ReportInputGeometry = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportInputOptions")>
     <DataMember()> Public Property ReportInputOptions() As Boolean?
        Get
            Return Me._ReportInputOptions
        End Get
        Set
            Me._ReportInputOptions = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportMaxForces")>
     <DataMember()> Public Property ReportMaxForces() As Boolean?
        Get
            Return Me._ReportMaxForces
        End Get
        Set
            Me._ReportMaxForces = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportInputMap")>
     <DataMember()> Public Property ReportInputMap() As Boolean?
        Get
            Return Me._ReportInputMap
        End Get
        Set
            Me._ReportInputMap = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description("{No Capacity Output, Capacity Summary, Capacity Details}"), DisplayName("CostReportOutputType")>
     <DataMember()> Public Property CostReportOutputType() As String
        Get
            Return Me._CostReportOutputType
        End Get
        Set
            Me._CostReportOutputType = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description("{No Cost Output, Cost Summary, Cost Details}"), DisplayName("CapacityReportOutputType")>
     <DataMember()> Public Property CapacityReportOutputType() As String
        Get
            Return Me._CapacityReportOutputType
        End Get
        Set
            Me._CapacityReportOutputType = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintForceTotals")>
     <DataMember()> Public Property ReportPrintForceTotals() As Boolean?
        Get
            Return Me._ReportPrintForceTotals
        End Get
        Set
            Me._ReportPrintForceTotals = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintForceDetails")>
     <DataMember()> Public Property ReportPrintForceDetails() As Boolean?
        Get
            Return Me._ReportPrintForceDetails
        End Get
        Set
            Me._ReportPrintForceDetails = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintMastVectors")>
     <DataMember()> Public Property ReportPrintMastVectors() As Boolean?
        Get
            Return Me._ReportPrintMastVectors
        End Get
        Set
            Me._ReportPrintMastVectors = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintAntPoleVectors")>
     <DataMember()> Public Property ReportPrintAntPoleVectors() As Boolean?
        Get
            Return Me._ReportPrintAntPoleVectors
        End Get
        Set
            Me._ReportPrintAntPoleVectors = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintDiscreteVectors")>
     <DataMember()> Public Property ReportPrintDiscreteVectors() As Boolean?
        Get
            Return Me._ReportPrintDiscreteVectors
        End Get
        Set
            Me._ReportPrintDiscreteVectors = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintDishVectors")>
     <DataMember()> Public Property ReportPrintDishVectors() As Boolean?
        Get
            Return Me._ReportPrintDishVectors
        End Get
        Set
            Me._ReportPrintDishVectors = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintFeedTowerVectors")>
     <DataMember()> Public Property ReportPrintFeedTowerVectors() As Boolean?
        Get
            Return Me._ReportPrintFeedTowerVectors
        End Get
        Set
            Me._ReportPrintFeedTowerVectors = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintUserLoadVectors")>
     <DataMember()> Public Property ReportPrintUserLoadVectors() As Boolean?
        Get
            Return Me._ReportPrintUserLoadVectors
        End Get
        Set
            Me._ReportPrintUserLoadVectors = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintPressures")>
     <DataMember()> Public Property ReportPrintPressures() As Boolean?
        Get
            Return Me._ReportPrintPressures
        End Get
        Set
            Me._ReportPrintPressures = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintAppurtForces")>
     <DataMember()> Public Property ReportPrintAppurtForces() As Boolean?
        Get
            Return Me._ReportPrintAppurtForces
        End Get
        Set
            Me._ReportPrintAppurtForces = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintGuyForces")>
     <DataMember()> Public Property ReportPrintGuyForces() As Boolean?
        Get
            Return Me._ReportPrintGuyForces
        End Get
        Set
            Me._ReportPrintGuyForces = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintGuyStressing")>
     <DataMember()> Public Property ReportPrintGuyStressing() As Boolean?
        Get
            Return Me._ReportPrintGuyStressing
        End Get
        Set
            Me._ReportPrintGuyStressing = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintDeflections")>
     <DataMember()> Public Property ReportPrintDeflections() As Boolean?
        Get
            Return Me._ReportPrintDeflections
        End Get
        Set
            Me._ReportPrintDeflections = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintReactions")>
     <DataMember()> Public Property ReportPrintReactions() As Boolean?
        Get
            Return Me._ReportPrintReactions
        End Get
        Set
            Me._ReportPrintReactions = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintStressChecks")>
     <DataMember()> Public Property ReportPrintStressChecks() As Boolean?
        Get
            Return Me._ReportPrintStressChecks
        End Get
        Set
            Me._ReportPrintStressChecks = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintBoltChecks")>
     <DataMember()> Public Property ReportPrintBoltChecks() As Boolean?
        Get
            Return Me._ReportPrintBoltChecks
        End Get
        Set
            Me._ReportPrintBoltChecks = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintInputGVerificationTables")>
     <DataMember()> Public Property ReportPrintInputGVerificationTables() As Boolean?
        Get
            Return Me._ReportPrintInputGVerificationTables
        End Get
        Set
            Me._ReportPrintInputGVerificationTables = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintOutputGVerificationTables")>
     <DataMember()> Public Property ReportPrintOutputGVerificationTables() As Boolean?
        Get
            Return Me._ReportPrintOutputGVerificationTables
        End Get
        Set
            Me._ReportPrintOutputGVerificationTables = Value
        End Set
    End Property
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
        Dim otherToCompare As tnxReportSettings = TryCast(other, tnxReportSettings)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.ReportInputCosts.CheckChange(otherToCompare.ReportInputCosts, changes, categoryName, "Report Input Costs"), Equals, False)
        Equals = If(Me.ReportInputGeometry.CheckChange(otherToCompare.ReportInputGeometry, changes, categoryName, "Report Input Geometry"), Equals, False)
        Equals = If(Me.ReportInputOptions.CheckChange(otherToCompare.ReportInputOptions, changes, categoryName, "Report Input Options"), Equals, False)
        Equals = If(Me.ReportMaxForces.CheckChange(otherToCompare.ReportMaxForces, changes, categoryName, "Report Max Forces"), Equals, False)
        Equals = If(Me.ReportInputMap.CheckChange(otherToCompare.ReportInputMap, changes, categoryName, "Report Input Map"), Equals, False)
        Equals = If(Me.CostReportOutputType.CheckChange(otherToCompare.CostReportOutputType, changes, categoryName, "Cost Report Output Type"), Equals, False)
        Equals = If(Me.CapacityReportOutputType.CheckChange(otherToCompare.CapacityReportOutputType, changes, categoryName, "Capacity Report Output Type"), Equals, False)
        Equals = If(Me.ReportPrintForceTotals.CheckChange(otherToCompare.ReportPrintForceTotals, changes, categoryName, "Report Print Force Totals"), Equals, False)
        Equals = If(Me.ReportPrintForceDetails.CheckChange(otherToCompare.ReportPrintForceDetails, changes, categoryName, "Report Print Force Details"), Equals, False)
        Equals = If(Me.ReportPrintMastVectors.CheckChange(otherToCompare.ReportPrintMastVectors, changes, categoryName, "Report Print Mast Vectors"), Equals, False)
        Equals = If(Me.ReportPrintAntPoleVectors.CheckChange(otherToCompare.ReportPrintAntPoleVectors, changes, categoryName, "Report Print Ant Pole Vectors"), Equals, False)
        Equals = If(Me.ReportPrintDiscreteVectors.CheckChange(otherToCompare.ReportPrintDiscreteVectors, changes, categoryName, "Report Print Discrete Vectors"), Equals, False)
        Equals = If(Me.ReportPrintDishVectors.CheckChange(otherToCompare.ReportPrintDishVectors, changes, categoryName, "Report Print Dish Vectors"), Equals, False)
        Equals = If(Me.ReportPrintFeedTowerVectors.CheckChange(otherToCompare.ReportPrintFeedTowerVectors, changes, categoryName, "Report Print Feed Tower Vectors"), Equals, False)
        Equals = If(Me.ReportPrintUserLoadVectors.CheckChange(otherToCompare.ReportPrintUserLoadVectors, changes, categoryName, "Report Print User Load Vectors"), Equals, False)
        Equals = If(Me.ReportPrintPressures.CheckChange(otherToCompare.ReportPrintPressures, changes, categoryName, "Report Print Pressures"), Equals, False)
        Equals = If(Me.ReportPrintAppurtForces.CheckChange(otherToCompare.ReportPrintAppurtForces, changes, categoryName, "Report Print Appurt Forces"), Equals, False)
        Equals = If(Me.ReportPrintGuyForces.CheckChange(otherToCompare.ReportPrintGuyForces, changes, categoryName, "Report Print Guy Forces"), Equals, False)
        Equals = If(Me.ReportPrintGuyStressing.CheckChange(otherToCompare.ReportPrintGuyStressing, changes, categoryName, "Report Print Guy Stressing"), Equals, False)
        Equals = If(Me.ReportPrintDeflections.CheckChange(otherToCompare.ReportPrintDeflections, changes, categoryName, "Report Print Deflections"), Equals, False)
        Equals = If(Me.ReportPrintReactions.CheckChange(otherToCompare.ReportPrintReactions, changes, categoryName, "Report Print Reactions"), Equals, False)
        Equals = If(Me.ReportPrintStressChecks.CheckChange(otherToCompare.ReportPrintStressChecks, changes, categoryName, "Report Print Stress Checks"), Equals, False)
        Equals = If(Me.ReportPrintBoltChecks.CheckChange(otherToCompare.ReportPrintBoltChecks, changes, categoryName, "Report Print Bolt Checks"), Equals, False)
        Equals = If(Me.ReportPrintInputGVerificationTables.CheckChange(otherToCompare.ReportPrintInputGVerificationTables, changes, categoryName, "Report Print Input GVerification Tables"), Equals, False)
        Equals = If(Me.ReportPrintOutputGVerificationTables.CheckChange(otherToCompare.ReportPrintOutputGVerificationTables, changes, categoryName, "Report Print Output GVerification Tables"), Equals, False)

        Return Equals
    End Function
End Class
<DataContract()>
Partial Public Class tnxMTOSettings
    Inherits EDSObject

    Public Overrides ReadOnly Property EDSObjectName As String = "MTO"

#Region "Define"

    Private _IncludeCapacityNote As Boolean?
    Private _IncludeAppurtGraphics As Boolean?
    Private _DisplayNotes As Boolean?
    Private _DisplayReactions As Boolean?
    Private _DisplaySchedule As Boolean?
    Private _DisplayAppurtenanceTable As Boolean?
    Private _DisplayMaterialStrengthTable As Boolean?
    Private _Notes As String
    'Private _Notes As New List(Of tnxNote)

    <Category("TNX MTO Settings"), Description(""), DisplayName("IncludeCapacityNote")>
     <DataMember()> Public Property IncludeCapacityNote() As Boolean?
        Get
            Return Me._IncludeCapacityNote
        End Get
        Set
            Me._IncludeCapacityNote = Value
        End Set
    End Property
    <Category("TNX MTO Settings"), Description(""), DisplayName("IncludeAppurtGraphics")>
     <DataMember()> Public Property IncludeAppurtGraphics() As Boolean?
        Get
            Return Me._IncludeAppurtGraphics
        End Get
        Set
            Me._IncludeAppurtGraphics = Value
        End Set
    End Property
    <Category("TNX MTO Settings"), Description(""), DisplayName("DisplayNotes")>
     <DataMember()> Public Property DisplayNotes() As Boolean?
        Get
            Return Me._DisplayNotes
        End Get
        Set
            Me._DisplayNotes = Value
        End Set
    End Property
    <Category("TNX MTO Settings"), Description(""), DisplayName("DisplayReactions")>
     <DataMember()> Public Property DisplayReactions() As Boolean?
        Get
            Return Me._DisplayReactions
        End Get
        Set
            Me._DisplayReactions = Value
        End Set
    End Property
    <Category("TNX MTO Settings"), Description(""), DisplayName("DisplaySchedule")>
     <DataMember()> Public Property DisplaySchedule() As Boolean?
        Get
            Return Me._DisplaySchedule
        End Get
        Set
            Me._DisplaySchedule = Value
        End Set
    End Property
    <Category("TNX MTO Settings"), Description(""), DisplayName("DisplayAppurtenanceTable")>
     <DataMember()> Public Property DisplayAppurtenanceTable() As Boolean?
        Get
            Return Me._DisplayAppurtenanceTable
        End Get
        Set
            Me._DisplayAppurtenanceTable = Value
        End Set
    End Property
    <Category("TNX MTO Settings"), Description(""), DisplayName("DisplayMaterialStrengthTable")>
     <DataMember()> Public Property DisplayMaterialStrengthTable() As Boolean?
        Get
            Return Me._DisplayMaterialStrengthTable
        End Get
        Set
            Me._DisplayMaterialStrengthTable = Value
        End Set
    End Property
    <Category("TNX MTO Settings"), Description(""), DisplayName("Notes")>
     <DataMember()> Public Property Notes() As String
        Get
            Return Me._Notes
        End Get
        Set
            Me._Notes = Value
        End Set
    End Property
    '<Category("TNX MTO Settings"), Description(""), DisplayName("Notes")>
    ' <DataMember()> Public Property Notes() As List(Of tnxNote)
    '    Get
    '        Return Me._Notes
    '    End Get
    '    Set
    '        Me._Notes = Value
    '    End Set
    'End Property
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
        Dim otherToCompare As tnxMTOSettings = TryCast(other, tnxMTOSettings)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.IncludeCapacityNote.CheckChange(otherToCompare.IncludeCapacityNote, changes, categoryName, "Include Capacity Note"), Equals, False)
        Equals = If(Me.IncludeAppurtGraphics.CheckChange(otherToCompare.IncludeAppurtGraphics, changes, categoryName, "Include Appurt Graphics"), Equals, False)
        Equals = If(Me.DisplayNotes.CheckChange(otherToCompare.DisplayNotes, changes, categoryName, "Display Notes"), Equals, False)
        Equals = If(Me.DisplayReactions.CheckChange(otherToCompare.DisplayReactions, changes, categoryName, "Display Reactions"), Equals, False)
        Equals = If(Me.DisplaySchedule.CheckChange(otherToCompare.DisplaySchedule, changes, categoryName, "Display Schedule"), Equals, False)
        Equals = If(Me.DisplayAppurtenanceTable.CheckChange(otherToCompare.DisplayAppurtenanceTable, changes, categoryName, "Display Appurtenance Table"), Equals, False)
        Equals = If(Me.DisplayMaterialStrengthTable.CheckChange(otherToCompare.DisplayMaterialStrengthTable, changes, categoryName, "Display Material Strength Table"), Equals, False)
        Equals = If(Me.Notes.CheckChange(otherToCompare.Notes, changes, categoryName, "Notes"), Equals, False)

        Return Equals
    End Function
End Class

'Partial Public Class tnxNote
'    Inherits EDSObject

'    Public Overrides ReadOnly Property EDSObjectName As String = "Notes"

'#Region "Define"
'#End Region

'    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
'        Equals = True
'        If changes Is Nothing Then changes = New List(Of AnalysisChange)
'        Dim categoryName As String = Me.EDSObjectFullName

'        'Makes sure you are comparing to the same object type
'        'Customize this to the object type
'        Dim otherToCompare As tnxSeismic = TryCast(other, tnxSeismic)
'        If otherToCompare Is Nothing Then Return False

'        'Equals here

'        Return Equals
'    End Function

'    Private _Note As String

'    <Category("TNX Note"), Description(""), DisplayName("Note")>
'     <DataMember()> Public Property Note As String
'        Get
'            Return Me._Note
'        End Get
'        Set
'            Me._Note = Value
'        End Set
'    End Property

'End Class
<DataContract()>
Partial Public Class tnxProjectInfo
    Inherits EDSObject

    Public Overrides ReadOnly Property EDSObjectName As String = "Project Info"

#Region "Define"

    Private _DesignStandardSeries As String
    Private _UnitsSystem As String
    Private _ClientName As String
    Private _ProjectName As String
    Private _ProjectNumber As String
    Private _CreatedBy As String
    Private _CreatedOn As String
    Private _LastUsedBy As String
    Private _LastUsedOn As String
    Private _VersionUsed As String

    <Category("TNX Project Info"), Description("TIA/EIA or CSA-S37"), DisplayName("DesignStandardSeries")>
     <DataMember()> Public Property DesignStandardSeries() As String
        Get
            Return Me._DesignStandardSeries
        End Get
        Set
            Me._DesignStandardSeries = Value
        End Set
    End Property
    <Category("TNX Project Info"), Description("US or SI"), DisplayName("UnitsSystem")>
     <DataMember()> Public Property UnitsSystem() As String
        Get
            Return Me._UnitsSystem
        End Get
        Set
            Me._UnitsSystem = Value
        End Set
    End Property
    <Category("TNX Project Info"), Description(""), DisplayName("ClientName")>
     <DataMember()> Public Property ClientName() As String
        Get
            Return Me._ClientName
        End Get
        Set
            Me._ClientName = Value
        End Set
    End Property
    <Category("TNX Project Info"), Description(""), DisplayName("ProjectName")>
     <DataMember()> Public Property ProjectName() As String
        Get
            Return Me._ProjectName
        End Get
        Set
            Me._ProjectName = Value
        End Set
    End Property
    <Category("TNX Project Info"), Description(""), DisplayName("ProjectNumber")>
     <DataMember()> Public Property ProjectNumber() As String
        Get
            Return Me._ProjectNumber
        End Get
        Set
            Me._ProjectNumber = Value
        End Set
    End Property
    <Category("TNX Project Info"), Description(""), DisplayName("CreatedBy")>
     <DataMember()> Public Property CreatedBy() As String
        Get
            Return Me._CreatedBy
        End Get
        Set
            Me._CreatedBy = Value
        End Set
    End Property
    <Category("TNX Project Info"), Description(""), DisplayName("CreatedOn")>
     <DataMember()> Public Property CreatedOn() As String
        Get
            Return Me._CreatedOn
        End Get
        Set
            Me._CreatedOn = Value
        End Set
    End Property
    <Category("TNX Project Info"), Description(""), DisplayName("LastUsedBy")>
     <DataMember()> Public Property LastUsedBy() As String
        Get
            Return Me._LastUsedBy
        End Get
        Set
            Me._LastUsedBy = Value
        End Set
    End Property
    <Category("TNX Project Info"), Description(""), DisplayName("LastUsedOn")>
     <DataMember()> Public Property LastUsedOn() As String
        Get
            Return Me._LastUsedOn
        End Get
        Set
            Me._LastUsedOn = Value
        End Set
    End Property
    <Category("TNX Project Info"), Description(""), DisplayName("VersionUsed")>
     <DataMember()> Public Property VersionUsed() As String
        Get
            Return Me._VersionUsed
        End Get
        Set
            Me._VersionUsed = Value
        End Set
    End Property
#End Region

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As tnxProjectInfo = TryCast(other, tnxProjectInfo)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.DesignStandardSeries.CheckChange(otherToCompare.DesignStandardSeries, changes, categoryName, "Design Standard Series"), Equals, False)
        Equals = If(Me.UnitsSystem.CheckChange(otherToCompare.UnitsSystem, changes, categoryName, "Units System"), Equals, False)
        Equals = If(Me.ClientName.CheckChange(otherToCompare.ClientName, changes, categoryName, "Client Name"), Equals, False)
        Equals = If(Me.ProjectName.CheckChange(otherToCompare.ProjectName, changes, categoryName, "Project Name"), Equals, False)
        Equals = If(Me.ProjectNumber.CheckChange(otherToCompare.ProjectNumber, changes, categoryName, "Project Number"), Equals, False)
        Equals = If(Me.CreatedBy.CheckChange(otherToCompare.CreatedBy, changes, categoryName, "Created By"), Equals, False)
        Equals = If(Me.CreatedOn.CheckChange(otherToCompare.CreatedOn, changes, categoryName, "Created On"), Equals, False)
        Equals = If(Me.LastUsedBy.CheckChange(otherToCompare.LastUsedBy, changes, categoryName, "Last Used By"), Equals, False)
        Equals = If(Me.LastUsedOn.CheckChange(otherToCompare.LastUsedOn, changes, categoryName, "Last Used On"), Equals, False)
        Equals = If(Me.VersionUsed.CheckChange(otherToCompare.VersionUsed, changes, categoryName, "Version Used"), Equals, False)

        Return Equals
    End Function

End Class
<DataContract()>
Partial Public Class tnxUserInfo
    Inherits EDSObject

    Public Overrides ReadOnly Property EDSObjectName As String = "User Info"

#Region "Define"

    Private _ViewerUserName As String
    Private _ViewerCompanyName As String
    Private _ViewerStreetAddress As String
    Private _ViewerCityState As String
    Private _ViewerPhone As String
    Private _ViewerFAX As String
    Private _ViewerLogo As String
    Private _ViewerCompanyBitmap As String

    <Category("TNX User Info"), Description(""), DisplayName("ViewerUserName")>
     <DataMember()> Public Property ViewerUserName() As String
        Get
            Return Me._ViewerUserName
        End Get
        Set
            Me._ViewerUserName = Value
        End Set
    End Property
    <Category("TNX User Info"), Description(""), DisplayName("ViewerCompanyName")>
     <DataMember()> Public Property ViewerCompanyName() As String
        Get
            Return Me._ViewerCompanyName
        End Get
        Set
            Me._ViewerCompanyName = Value
        End Set
    End Property
    <Category("TNX User Info"), Description(""), DisplayName("ViewerStreetAddress")>
     <DataMember()> Public Property ViewerStreetAddress() As String
        Get
            Return Me._ViewerStreetAddress
        End Get
        Set
            Me._ViewerStreetAddress = Value
        End Set
    End Property
    <Category("TNX User Info"), Description(""), DisplayName("ViewerCityState")>
     <DataMember()> Public Property ViewerCityState() As String
        Get
            Return Me._ViewerCityState
        End Get
        Set
            Me._ViewerCityState = Value
        End Set
    End Property
    <Category("TNX User Info"), Description(""), DisplayName("ViewerPhone")>
     <DataMember()> Public Property ViewerPhone() As String
        Get
            Return Me._ViewerPhone
        End Get
        Set
            Me._ViewerPhone = Value
        End Set
    End Property
    <Category("TNX User Info"), Description(""), DisplayName("ViewerFAX")>
     <DataMember()> Public Property ViewerFAX() As String
        Get
            Return Me._ViewerFAX
        End Get
        Set
            Me._ViewerFAX = Value
        End Set
    End Property
    <Category("TNX User Info"), Description(""), DisplayName("ViewerLogo")>
     <DataMember()> Public Property ViewerLogo() As String
        Get
            Return Me._ViewerLogo
        End Get
        Set
            Me._ViewerLogo = Value
        End Set
    End Property
    <Category("TNX User Info"), Description(""), DisplayName("ViewerCompanyBitmap")>
     <DataMember()> Public Property ViewerCompanyBitmap() As String
        Get
            Return Me._ViewerCompanyBitmap
        End Get
        Set
            Me._ViewerCompanyBitmap = Value
        End Set
    End Property
#End Region
    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As tnxUserInfo = TryCast(other, tnxUserInfo)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.ViewerUserName.CheckChange(otherToCompare.ViewerUserName, changes, categoryName, "Viewer User Name"), Equals, False)
        Equals = If(Me.ViewerCompanyName.CheckChange(otherToCompare.ViewerCompanyName, changes, categoryName, "Viewer Company Name"), Equals, False)
        Equals = If(Me.ViewerStreetAddress.CheckChange(otherToCompare.ViewerStreetAddress, changes, categoryName, "Viewer Street Address"), Equals, False)
        Equals = If(Me.ViewerCityState.CheckChange(otherToCompare.ViewerCityState, changes, categoryName, "Viewer City State"), Equals, False)
        Equals = If(Me.ViewerPhone.CheckChange(otherToCompare.ViewerPhone, changes, categoryName, "Viewer Phone"), Equals, False)
        Equals = If(Me.ViewerFAX.CheckChange(otherToCompare.ViewerFAX, changes, categoryName, "Viewer FAX"), Equals, False)
        Equals = If(Me.ViewerLogo.CheckChange(otherToCompare.ViewerLogo, changes, categoryName, "Viewer Logo"), Equals, False)
        Equals = If(Me.ViewerCompanyBitmap.CheckChange(otherToCompare.ViewerCompanyBitmap, changes, categoryName, "Viewer Company Bitmap"), Equals, False)

        Return Equals
    End Function
End Class


<DataContract()>
Partial Public Class tnxUnits

    Private _Length As New tnxLengthUnit()
    Private _Coordinate As New tnxCoordinateUnit()
    Private _Force As New tnxForceUnit()
    Private _Load As New tnxLoadUnit()
    Private _Moment As New tnxMomentUnit()
    Private _Properties As New tnxPropertiesUnit()
    Private _Pressure As New tnxPressureUnit()
    Private _Velocity As New tnxVelocityUnit()
    Private _Displacement As New tnxDisplacementUnit()
    Private _Mass As New tnxMassUnit()
    Private _Acceleration As New tnxAccelerationUnit()
    Private _Stress As New tnxStressUnit()
    Private _Density As New tnxDensityUnit()
    Private _UnitWt As New tnxUnitWTUnit()
    Private _Strength As New tnxStrengthUnit()
    Private _Modulus As New tnxModulusUnit()
    Private _Temperature As New tnxTempUnit()
    Private _Printer As New tnxPrinterUnit()
    Private _Rotation As New tnxRotationUnit()
    Private _Spacing As New tnxSpacingUnit()

    <Category("TNX Units"), Description(""), DisplayName("Length")>
     <DataMember()> Public Property Length() As tnxLengthUnit
        Get
            Return Me._Length
        End Get
        Set
            Me._Length = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Coordinate")>
     <DataMember()> Public Property Coordinate() As tnxCoordinateUnit
        Get
            Return Me._Coordinate
        End Get
        Set
            Me._Coordinate = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Force")>
     <DataMember()> Public Property Force() As tnxForceUnit
        Get
            Return Me._Force
        End Get
        Set
            Me._Force = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Load")>
     <DataMember()> Public Property Load() As tnxLoadUnit
        Get
            Return Me._Load
        End Get
        Set
            Me._Load = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Moment")>
     <DataMember()> Public Property Moment() As tnxMomentUnit
        Get
            Return Me._Moment
        End Get
        Set
            Me._Moment = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Properties")>
     <DataMember()> Public Property Properties() As tnxPropertiesUnit
        Get
            Return Me._Properties
        End Get
        Set
            Me._Properties = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Pressure")>
     <DataMember()> Public Property Pressure() As tnxPressureUnit
        Get
            Return Me._Pressure
        End Get
        Set
            Me._Pressure = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Velocity")>
     <DataMember()> Public Property Velocity() As tnxVelocityUnit
        Get
            Return Me._Velocity
        End Get
        Set
            Me._Velocity = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Displacement")>
     <DataMember()> Public Property Displacement() As tnxDisplacementUnit
        Get
            Return Me._Displacement
        End Get
        Set
            Me._Displacement = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Mass")>
     <DataMember()> Public Property Mass() As tnxMassUnit
        Get
            Return Me._Mass
        End Get
        Set
            Me._Mass = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Acceleration")>
     <DataMember()> Public Property Acceleration() As tnxAccelerationUnit
        Get
            Return Me._Acceleration
        End Get
        Set
            Me._Acceleration = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Stress")>
     <DataMember()> Public Property Stress() As tnxStressUnit
        Get
            Return Me._Stress
        End Get
        Set
            Me._Stress = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Density")>
     <DataMember()> Public Property Density() As tnxDensityUnit
        Get
            Return Me._Density
        End Get
        Set
            Me._Density = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Unitwt")>
     <DataMember()> Public Property UnitWt() As tnxUnitWTUnit
        Get
            Return Me._UnitWt
        End Get
        Set
            Me._UnitWt = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Strength")>
     <DataMember()> Public Property Strength() As tnxStrengthUnit
        Get
            Return Me._Strength
        End Get
        Set
            Me._Strength = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Modulus")>
     <DataMember()> Public Property Modulus() As tnxModulusUnit
        Get
            Return Me._Modulus
        End Get
        Set
            Me._Modulus = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Temperature")>
     <DataMember()> Public Property Temperature() As tnxTempUnit
        Get
            Return Me._Temperature
        End Get
        Set
            Me._Temperature = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Printer")>
     <DataMember()> Public Property Printer() As tnxPrinterUnit
        Get
            Return Me._Printer
        End Get
        Set
            Me._Printer = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Rotation")>
     <DataMember()> Public Property Rotation() As tnxRotationUnit
        Get
            Return Me._Rotation
        End Get
        Set
            Me._Rotation = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Spacing")>
     <DataMember()> Public Property Spacing() As tnxSpacingUnit
        Get
            Return Me._Spacing
        End Get
        Set
            Me._Spacing = Value
        End Set
    End Property

    Public Function convertForcePerUnitLengthtoDefault(InputValue As Double?) As Double?
        If Not InputValue.HasValue Then
            Return Nothing
        End If

        Try
            Return InputValue.Value * (Me.Properties.multiplier / Me.Force.multiplier)
        Catch ex As Exception
            Debug.Print("Error Converting Force Per Unit Length to Default")
            Return Nothing
        End Try
    End Function

    Public Function convertForcePerUnitLengthtoDefault(InputValue As String) As Double?
        Dim dblInputValue As Double
        If Not Double.TryParse(InputValue, dblInputValue) Then
            Return Nothing
        End If

        Try
            Return dblInputValue * (Me.Properties.multiplier / Me.Force.multiplier)
        Catch ex As Exception
            Debug.Print("Error Converting Force Per Unit Length to Default")
            Return Nothing
        End Try
    End Function
    Public Function convertForcePerUnitLengthtoERISpecified(InputValue As Double?) As Double?
        If Not InputValue.HasValue Then
            Return Nothing
        End If

        Try
            Return InputValue.Value / (Me.Properties.multiplier / Me.Force.multiplier)
        Catch ex As Exception
            Debug.Print("Error Converting Force Per Unit Length to ERI Specified")
            Return Nothing
        End Try
    End Function

End Class

Partial Public Class tnxUnitProperty
    'Variables need to be public for inheritance
    Public _value As String
    Public _precision As Integer?
    Public _multiplier As Double?

    <Category("TNX Unit Property"), Description(""), DisplayName("Value")>
    Public Overridable Property value() As String
        Get
            Return Me._value
        End Get
        Set
            Me._value = Value
        End Set
    End Property
    <Category("TNX Unit Property"), Description(""), DisplayName("Precision")>
    Public Overridable Property precision() As Integer?
        Get
            Return Me._precision
        End Get
        Set
            If Value < 0 Then
                Me._precision = 0
            ElseIf Value > 4 Then
                Me._precision = 4
            Else
                Me._precision = Value
            End If
        End Set
    End Property
    <Category("TNX Unit Property"), Description("Used to convert TNX file units to default EDS units during import."), DisplayName("Multiplier")>
    Public Overridable Property multiplier() As Double?
        Get
            Return Me._multiplier
        End Get
        Set
            Me._multiplier = Value
        End Set
    End Property

    Public Sub New()
    End Sub

    Public Overridable Function convertToEDSDefaultUnits(InputValue As Double?) As Double?

        If InputValue Is Nothing Then
            Return Nothing
        End If

        If Me._value = "" Or Me._value Is Nothing Then
            Throw New System.Exception("Property value not set")
        ElseIf Me._multiplier = 0 Or Me._multiplier Is Nothing Then
            Throw New System.Exception("Property multiplier not set")
        End If

        Return Math.Round(InputValue.Value / Me.multiplier.Value, 6)

    End Function

    Public Overridable Function convertToEDSDefaultUnits(strInputValue As String) As Double?
        Dim InputValue As Double
        If Not Double.TryParse(strInputValue, InputValue) Then
            Return Nothing
        End If

        If Me._value = "" Or Me._value Is Nothing Then
            Throw New System.Exception("Property value not set")
        ElseIf Me._multiplier = 0 Or Me._multiplier Is Nothing Then
            Throw New System.Exception("Property multiplier not set")
        End If

        Return Math.Round(InputValue / Me.multiplier.Value, 6)

    End Function

    Public Overridable Function convertToERIUnits(InputValue As Double?) As Double?

        If InputValue Is Nothing Then
            Return Nothing
        End If

        If Me._value = "" Or Me._value Is Nothing Then
            Throw New System.Exception("Property value not set")
        ElseIf Me._multiplier = 0 Or Me._multiplier Is Nothing Then
            Throw New System.Exception("Property multiplier not set")
        End If

        Return Math.Round(InputValue.Value * Me.multiplier.Value, 4)

    End Function

End Class

Partial Public Class tnxLengthUnit
    Inherits tnxUnitProperty

    Public Overrides Property value() As String
        Get
            Return Me._value
        End Get
        Set
            Me._value = Value

            If Me._value = "ft" Then
                Me._multiplier = 1
            ElseIf Me._value = "in" Then
                Me._multiplier = 12
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me._value)
            End If
        End Set
    End Property

    Public Overridable Function convertAreaToEDSDefaultUnits(InputValue As Double?) As Double?

        If Me._value = "" Then
            Throw New System.Exception("Property value not set")
        ElseIf Me._multiplier = 0 Then
            Throw New System.Exception("Property multiplier not set")
        End If

        Return InputValue / (Me.multiplier * Me.multiplier)

    End Function

    Public Overridable Function convertAreaToERIUnits(InputValue As Double?) As Double?

        If Me._value = "" Then
            Throw New System.Exception("Property value not set")
        ElseIf Me._multiplier = 0 Then
            Throw New System.Exception("Property multiplier not set")
        End If

        Return InputValue * Me.multiplier * Me.multiplier

    End Function

    Public Sub New()
    End Sub
    Public Sub New(new_value As String)

        Me.value = new_value

    End Sub

End Class

Partial Public Class tnxCoordinateUnit
    Inherits tnxLengthUnit
    Public Sub New()
    End Sub
    Public Sub New(new_value As String)

        Me.value = new_value

    End Sub
End Class

Partial Public Class tnxForceUnit
    Inherits tnxUnitProperty

    Public Overrides Property value() As String
        Get
            Return Me._value
        End Get
        Set
            Me._value = Value

            If Me._value = "K" Then
                Me._multiplier = 1
            ElseIf Me._value = "lb" Then
                Me._multiplier = 1000
            ElseIf Me._value = "T" Then
                Me._multiplier = 0.5
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me._value)
            End If
        End Set
    End Property
    Public Sub New()
    End Sub
    Public Sub New(new_value As String)

        Me.value = new_value

    End Sub
End Class
Partial Public Class tnxLoadUnit
    Inherits tnxUnitProperty

    Public Overrides Property value() As String
        Get
            Return Me._value
        End Get
        Set
            Me._value = Value

            If Me._value = "klf" Then
                Me._multiplier = 1
            ElseIf Me._value = "plf" Then
                Me._multiplier = 1000
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me._value)
            End If
        End Set
    End Property
    Public Sub New()
    End Sub
    Public Sub New(new_value As String)

        Me.value = new_value

    End Sub
End Class
Partial Public Class tnxMomentUnit
    Inherits tnxUnitProperty

    Public Overrides Property value() As String
        Get
            Return Me._value
        End Get
        Set
            Me._value = Value

            If Me._value = "kip-ft" Then
                Me._multiplier = 1
            ElseIf Me._value = "lb-ft" Then
                Me._multiplier = 1000
            ElseIf Me._value = "lb-in" Then
                Me._multiplier = 12000
            ElseIf Me._value = "kip-in" Then
                Me._multiplier = 12
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me._value)
            End If
        End Set
    End Property
    Public Sub New()
    End Sub
    Public Sub New(new_value As String)

        Me.value = new_value

    End Sub
End Class
Partial Public Class tnxPropertiesUnit
    Inherits tnxLengthUnit
    Public Sub New()
    End Sub
    Public Sub New(new_value As String)

        Me.value = new_value

    End Sub
End Class
Partial Public Class tnxPressureUnit
    Inherits tnxUnitProperty

    Public Overrides Property value() As String
        Get
            Return Me._value
        End Get
        Set
            Me._value = Value

            If Me._value = "ksf" Then
                Me._multiplier = 1
            ElseIf Me._value = "psf" Then
                Me._multiplier = 1000
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me._value)
            End If
        End Set
    End Property
    Public Sub New()
    End Sub
    Public Sub New(new_value As String)

        Me.value = new_value

    End Sub
End Class
Partial Public Class tnxVelocityUnit
    Inherits tnxUnitProperty

    Public Overrides Property value() As String
        Get
            Return Me._value
        End Get
        Set
            Me._value = Value

            If Me._value = "mph" Then
                Me._multiplier = 1
            ElseIf Me._value = "fps" Then
                Me._multiplier = 5280 / 3600
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me._value)
            End If
        End Set
    End Property
    Public Sub New()
    End Sub
    Public Sub New(new_value As String)

        Me.value = new_value

    End Sub
End Class
Partial Public Class tnxDisplacementUnit
    'Note: This is called deflection in the TNX UI
    Inherits tnxLengthUnit
    Public Overrides Property precision() As Integer?
        Get
            Return Me._precision
        End Get
        Set
            If Value < 0 Then
                Me._precision = 0
            ElseIf Value > 6 Then
                Me._precision = 6
            Else
                Me._precision = Value
            End If
        End Set
    End Property
    Public Sub New()
    End Sub
    Public Sub New(new_value As String)

        Me.value = new_value

    End Sub
End Class
Partial Public Class tnxMassUnit
    'This property isn't accessible in the TNX UI
    Inherits tnxUnitProperty

    Public Overrides Property value() As String
        Get
            Return Me._value
        End Get
        Set
            Me._value = Value

            If Me._value = "lb" Then
                Me._multiplier = 1
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me._value)
            End If
        End Set
    End Property
    Public Sub New()
    End Sub
    Public Sub New(new_value As String)

        Me.value = new_value

    End Sub
End Class
Partial Public Class tnxAccelerationUnit
    Inherits tnxUnitProperty

    Public Overrides Property value() As String
        Get
            Return Me._value
        End Get
        Set
            Me._value = Value

            If Me._value = "G" Then
                Me._multiplier = 1
            ElseIf Me._value = "fpss" Then
                Me._multiplier = 32.17405
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me._value)
            End If
        End Set
    End Property
    Public Sub New()
    End Sub
    Public Sub New(new_value As String)

        Me.value = new_value

    End Sub
End Class
Partial Public Class tnxStressUnit
    Inherits tnxUnitProperty

    Public Overrides Property value() As String
        Get
            Return Me._value
        End Get
        Set
            Me._value = Value

            If Me._value = "ksi" Then
                Me._multiplier = 1
            ElseIf Me._value = "psi" Then
                Me._multiplier = 1000
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me._value)
            End If
        End Set
    End Property
    Public Sub New()
    End Sub
    Public Sub New(new_value As String)

        Me.value = new_value

    End Sub
End Class
Partial Public Class tnxDensityUnit
    Inherits tnxUnitProperty

    Public Overrides Property value() As String
        Get
            Return Me._value
        End Get
        Set
            Me._value = Value

            If Me._value = "pcf" Then
                Me._multiplier = 1
            ElseIf Me._value = "pci" Then
                Me._multiplier = 1728
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me._value)
            End If
        End Set
    End Property
    Public Sub New()
    End Sub
    Public Sub New(new_value As String)

        Me.value = new_value

    End Sub
End Class
Partial Public Class tnxUnitWTUnit
    Inherits tnxUnitProperty
    'As of version 8.1.1.0 of TNX there is a bug in TNX, the unit wt is always tied to the density units.
    Public Overrides Property value() As String
        Get
            Return Me._value
        End Get
        Set
            Me._value = Value

            If Me._value = "plf" Then
                Me._multiplier = 1
            ElseIf Me._value = "klf" Then
                Me._multiplier = 0.001
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me._value)
            End If
        End Set
    End Property
    Public Sub New()
    End Sub
    Public Sub New(new_value As String)

        Me.value = new_value

    End Sub
End Class
Partial Public Class tnxStrengthUnit
    Inherits tnxUnitProperty

    Public Overrides Property value() As String
        Get
            Return Me._value
        End Get
        Set
            Me._value = Value

            If Me._value = "ksi" Then
                Me._multiplier = 1
            ElseIf Me._value = "psi" Then
                Me._multiplier = 1000
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me._value)
            End If
        End Set
    End Property
    Public Sub New()
    End Sub
    Public Sub New(new_value As String)

        Me.value = new_value

    End Sub
End Class
Partial Public Class tnxModulusUnit
    Inherits tnxUnitProperty

    Public Overrides Property value() As String
        Get
            Return Me._value
        End Get
        Set
            Me._value = Value

            If Me._value = "ksi" Then
                Me._multiplier = 1
            ElseIf Me._value = "psi" Then
                Me._multiplier = 1000
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me._value)
            End If
        End Set
    End Property
    Public Sub New()
    End Sub
    Public Sub New(new_value As String)

        Me.value = new_value

    End Sub
End Class
Partial Public Class tnxTempUnit
    Inherits tnxUnitProperty

    Public Overrides Property value() As String
        Get
            Return Me._value
        End Get
        Set
            Me._value = Value

            If Me._value = "F" Then
                Me._multiplier = 1
            ElseIf Me._value = "C" Then
                'This conversion doesn't use a simple multiplier.
                'Override coversion function to get correct results
                Me._multiplier = 1
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me._value)
            End If
        End Set
    End Property

    Public Overrides Function convertToEDSDefaultUnits(InputValue As Double?) As Double?

        If Not InputValue.HasValue Then
            Return Nothing
        End If

        If Me._value = "" Then
            Throw New System.Exception("Property value not set")
        End If

        If Me._value = "C" Then
            Return Math.Round(InputValue.Value * (9 / 5) + 32, 6)
        Else
            Return InputValue
        End If

    End Function

    Public Overrides Function convertToERIUnits(InputValue As Double?) As Double?

        If Not InputValue.HasValue Then
            Return Nothing
        End If

        If Me._value = "" Then
            Throw New System.Exception("Property value not set")
        End If

        If Me._value = "C" Then
            Return Math.Round((InputValue.Value - 32) * (5 / 9), 6)
        Else
            Return InputValue
        End If

    End Function

    Public Sub New()
    End Sub
    Public Sub New(new_value As String)

        Me.value = new_value

    End Sub
End Class
Partial Public Class tnxPrinterUnit
    'This property isn't accessible in the TNX UI
    Inherits tnxUnitProperty

    Public Overrides Property value() As String
        Get
            Return Me._value
        End Get
        Set
            Me._value = Value

            If Me._value = "in" Then
                Me._multiplier = 1
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me._value)
            End If
        End Set
    End Property

    Public Sub New()
    End Sub
    Public Sub New(new_value As String)

        Me.value = new_value

    End Sub
End Class
Partial Public Class tnxRotationUnit
    Inherits tnxUnitProperty

    Public Overrides Property precision() As Integer?
        Get
            Return Me._precision
        End Get
        Set
            If Value < 0 Then
                Me._precision = 0
            ElseIf Value > 6 Then
                Me._precision = 6
            Else
                Me._precision = Value
            End If
        End Set
    End Property

    Public Overrides Property value() As String
        Get
            Return Me._value
        End Get
        Set
            Me._value = Value

            If Me._value = "deg" Then
                Me._multiplier = 1
            ElseIf Me._value = "rad" Then
                Me._multiplier = 3.14159 / 180
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me._value)
            End If
        End Set
    End Property

    Public Sub New()
    End Sub
    Public Sub New(new_value As String)

        Me.value = new_value

    End Sub
End Class
Partial Public Class tnxSpacingUnit
    Inherits tnxLengthUnit

    Public Sub New()
    End Sub
    Public Sub New(new_value As String)

        Me.value = new_value

    End Sub
End Class

#End Region

<DataContract()>
Partial Public Class tnxCCIReport
    Inherits EDSObject

    Public Overrides ReadOnly Property EDSObjectName As String = "CCI Report"
#Region "Define"
    Private _sReportProjectNumber As String
    Private _sReportJobType As String
    Private _sReportCarrierName As String
    Private _sReportCarrierSiteNumber As String
    Private _sReportCarrierSiteName As String
    Private _sReportSiteAddress As String
    Private _sReportLatitudeDegree As Double?
    Private _sReportLatitudeMinute As Double?
    Private _sReportLatitudeSecond As Double?
    Private _sReportLongitudeDegree As Double?
    Private _sReportLongitudeMinute As Double?
    Private _sReportLongitudeSecond As Double?
    Private _sReportLocalCodeRequirement As String
    Private _sReportSiteHistory As String
    Private _sReportTowerManufacturer As String
    Private _sReportMonthManufactured As String
    Private _sReportYearManufactured As Integer?
    Private _sReportOriginalSpeed As Double?
    Private _sReportOriginalCode As String
    Private _sReportTowerType As String
    Private _sReportEngrName As String
    Private _sReportEngrTitle As String
    Private _sReportHQPhoneNumber As String
    Private _sReportEmailAddress As String
    Private _sReportLogoPath As String
    Private _sReportCCiContactName As String
    Private _sReportCCiAddress1 As String
    Private _sReportCCiAddress2 As String
    Private _sReportCCiBUNumber As String
    Private _sReportCCiSiteName As String
    Private _sReportCCiJDENumber As String
    Private _sReportCCiWONumber As String
    Private _sReportCCiPONumber As String
    Private _sReportCCiAppNumber As String
    Private _sReportCCiRevNumber As String
    Private _sReportDocsProvided As New List(Of String)
    Private _sReportRecommendations As String
    Private _sReportAppurt1 As New List(Of String)
    Private _sReportAppurt2 As New List(Of String)
    Private _sReportAppurt3 As New List(Of String)
    Private _sReportAddlCapacity As New List(Of String)
    Private _sReportAssumption As New List(Of String)
    Private _sReportAppurt1Note1 As String
    Private _sReportAppurt1Note2 As String
    Private _sReportAppurt1Note3 As String
    Private _sReportAppurt1Note4 As String
    Private _sReportAppurt1Note5 As String
    Private _sReportAppurt1Note6 As String
    Private _sReportAppurt1Note7 As String
    Private _sReportAppurt2Note1 As String
    Private _sReportAppurt2Note2 As String
    Private _sReportAppurt2Note3 As String
    Private _sReportAppurt2Note4 As String
    Private _sReportAppurt2Note5 As String
    Private _sReportAppurt2Note6 As String
    Private _sReportAppurt2Note7 As String
    Private _sReportAddlCapacityNote1 As String
    Private _sReportAddlCapacityNote2 As String
    Private _sReportAddlCapacityNote3 As String
    Private _sReportAddlCapacityNote4 As String

    <Category("TNX CCI Report"), Description(""), DisplayName("sReportProjectNumber")>
     <DataMember()> Public Property sReportProjectNumber() As String
        Get
            Return Me._sReportProjectNumber
        End Get
        Set
            Me._sReportProjectNumber = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportJobType")>
     <DataMember()> Public Property sReportJobType() As String
        Get
            Return Me._sReportJobType
        End Get
        Set
            Me._sReportJobType = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportCarrierName")>
     <DataMember()> Public Property sReportCarrierName() As String
        Get
            Return Me._sReportCarrierName
        End Get
        Set
            Me._sReportCarrierName = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportCarrierSiteNumber")>
     <DataMember()> Public Property sReportCarrierSiteNumber() As String
        Get
            Return Me._sReportCarrierSiteNumber
        End Get
        Set
            Me._sReportCarrierSiteNumber = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportCarrierSiteName")>
     <DataMember()> Public Property sReportCarrierSiteName() As String
        Get
            Return Me._sReportCarrierSiteName
        End Get
        Set
            Me._sReportCarrierSiteName = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportSiteAddress")>
     <DataMember()> Public Property sReportSiteAddress() As String
        Get
            Return Me._sReportSiteAddress
        End Get
        Set
            Me._sReportSiteAddress = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportLatitudeDegree")>
     <DataMember()> Public Property sReportLatitudeDegree() As Double?
        Get
            Return Me._sReportLatitudeDegree
        End Get
        Set
            Me._sReportLatitudeDegree = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportLatitudeMinute")>
     <DataMember()> Public Property sReportLatitudeMinute() As Double?
        Get
            Return Me._sReportLatitudeMinute
        End Get
        Set
            Me._sReportLatitudeMinute = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportLatitudeSecond")>
     <DataMember()> Public Property sReportLatitudeSecond() As Double?
        Get
            Return Me._sReportLatitudeSecond
        End Get
        Set
            Me._sReportLatitudeSecond = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportLongitudeDegree")>
     <DataMember()> Public Property sReportLongitudeDegree() As Double?
        Get
            Return Me._sReportLongitudeDegree
        End Get
        Set
            Me._sReportLongitudeDegree = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportLongitudeMinute")>
     <DataMember()> Public Property sReportLongitudeMinute() As Double?
        Get
            Return Me._sReportLongitudeMinute
        End Get
        Set
            Me._sReportLongitudeMinute = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportLongitudeSecond")>
     <DataMember()> Public Property sReportLongitudeSecond() As Double?
        Get
            Return Me._sReportLongitudeSecond
        End Get
        Set
            Me._sReportLongitudeSecond = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportLocalCodeRequirement")>
     <DataMember()> Public Property sReportLocalCodeRequirement() As String
        Get
            Return Me._sReportLocalCodeRequirement
        End Get
        Set
            Me._sReportLocalCodeRequirement = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportSiteHistory")>
     <DataMember()> Public Property sReportSiteHistory() As String
        Get
            Return Me._sReportSiteHistory
        End Get
        Set
            Me._sReportSiteHistory = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportTowerManufacturer")>
     <DataMember()> Public Property sReportTowerManufacturer() As String
        Get
            Return Me._sReportTowerManufacturer
        End Get
        Set
            Me._sReportTowerManufacturer = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportMonthManufactured")>
     <DataMember()> Public Property sReportMonthManufactured() As String
        Get
            Return Me._sReportMonthManufactured
        End Get
        Set
            Me._sReportMonthManufactured = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportYearManufactured")>
     <DataMember()> Public Property sReportYearManufactured() As Integer?
        Get
            Return Me._sReportYearManufactured
        End Get
        Set
            Me._sReportYearManufactured = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportOriginalSpeed")>
     <DataMember()> Public Property sReportOriginalSpeed() As Double?
        Get
            Return Me._sReportOriginalSpeed
        End Get
        Set
            Me._sReportOriginalSpeed = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportOriginalCode")>
     <DataMember()> Public Property sReportOriginalCode() As String
        Get
            Return Me._sReportOriginalCode
        End Get
        Set
            Me._sReportOriginalCode = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportTowerType")>
     <DataMember()> Public Property sReportTowerType() As String
        Get
            Return Me._sReportTowerType
        End Get
        Set
            Me._sReportTowerType = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportEngrName")>
     <DataMember()> Public Property sReportEngrName() As String
        Get
            Return Me._sReportEngrName
        End Get
        Set
            Me._sReportEngrName = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportEngrTitle")>
     <DataMember()> Public Property sReportEngrTitle() As String
        Get
            Return Me._sReportEngrTitle
        End Get
        Set
            Me._sReportEngrTitle = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportHQPhoneNumber")>
     <DataMember()> Public Property sReportHQPhoneNumber() As String
        Get
            Return Me._sReportHQPhoneNumber
        End Get
        Set
            Me._sReportHQPhoneNumber = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportEmailAddress")>
     <DataMember()> Public Property sReportEmailAddress() As String
        Get
            Return Me._sReportEmailAddress
        End Get
        Set
            Me._sReportEmailAddress = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportLogoPath")>
     <DataMember()> Public Property sReportLogoPath() As String
        Get
            Return Me._sReportLogoPath
        End Get
        Set
            Me._sReportLogoPath = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportCCiContactName")>
     <DataMember()> Public Property sReportCCiContactName() As String
        Get
            Return Me._sReportCCiContactName
        End Get
        Set
            Me._sReportCCiContactName = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportCCiAddress1")>
     <DataMember()> Public Property sReportCCiAddress1() As String
        Get
            Return Me._sReportCCiAddress1
        End Get
        Set
            Me._sReportCCiAddress1 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportCCiAddress2")>
     <DataMember()> Public Property sReportCCiAddress2() As String
        Get
            Return Me._sReportCCiAddress2
        End Get
        Set
            Me._sReportCCiAddress2 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportCCiBUNumber")>
     <DataMember()> Public Property sReportCCiBUNumber() As String
        Get
            Return Me._sReportCCiBUNumber
        End Get
        Set
            Me._sReportCCiBUNumber = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportCCiSiteName")>
     <DataMember()> Public Property sReportCCiSiteName() As String
        Get
            Return Me._sReportCCiSiteName
        End Get
        Set
            Me._sReportCCiSiteName = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportCCiJDENumber")>
     <DataMember()> Public Property sReportCCiJDENumber() As String
        Get
            Return Me._sReportCCiJDENumber
        End Get
        Set
            Me._sReportCCiJDENumber = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportCCiWONumber")>
     <DataMember()> Public Property sReportCCiWONumber() As String
        Get
            Return Me._sReportCCiWONumber
        End Get
        Set
            Me._sReportCCiWONumber = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportCCiPONumber")>
     <DataMember()> Public Property sReportCCiPONumber() As String
        Get
            Return Me._sReportCCiPONumber
        End Get
        Set
            Me._sReportCCiPONumber = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportCCiAppNumber")>
     <DataMember()> Public Property sReportCCiAppNumber() As String
        Get
            Return Me._sReportCCiAppNumber
        End Get
        Set
            Me._sReportCCiAppNumber = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportCCiRevNumber")>
     <DataMember()> Public Property sReportCCiRevNumber() As String
        Get
            Return Me._sReportCCiRevNumber
        End Get
        Set
            Me._sReportCCiRevNumber = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description("Reference Document Row. String format: Doc Type<~~>Remarks<~~>Ref No<~~>Source"), DisplayName("sReportDocsProvided")>
     <DataMember()> Public Property sReportDocsProvided() As List(Of String)
        Get
            Return Me._sReportDocsProvided
        End Get
        Set
            Me._sReportDocsProvided = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportRecommendations")>
     <DataMember()> Public Property sReportRecommendations() As String
        Get
            Return Me._sReportRecommendations
        End Get
        Set
            Me._sReportRecommendations = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description("Proposed Equipment Row. String format: MCL<~~>ECL<~~>qty<~~>manufacturer<~~>model<~~>FL qty<~~>FL Size<~~>Note #<~~>?<~~>Proposed"), DisplayName("sReportAppurt1")>
     <DataMember()> Public Property sReportAppurt1() As List(Of String)
        Get
            Return Me._sReportAppurt1
        End Get
        Set
            Me._sReportAppurt1 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description("Existing Equipment Row. String format:MCL<~~>ECL<~~>qty<~~>manufacturer<~~>model<~~>FL qty<~~>FL Size<~~>Note #<~~>?<~~>Existing"), DisplayName("sReportAppurt2")>
     <DataMember()> Public Property sReportAppurt2() As List(Of String)
        Get
            Return Me._sReportAppurt2
        End Get
        Set
            Me._sReportAppurt2 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description("Design Equipment Row. String format: MCL<~~>ECL<~~>qty<~~>manufacturer<~~>model<~~>FL qty<~~>FL Size<~~>"), DisplayName("sReportAppurt2")>
     <DataMember()> Public Property sReportAppurt3() As List(Of String)
        Get
            Return Me._sReportAppurt3
        End Get
        Set
            Me._sReportAppurt3 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description("Additional Capacity Row. String format: Component<~~>Note #<~~>Elevation<~~>Cap%<~~>Pass/Fail<~~>Include in Report {Yes/No}"), DisplayName("sReportAddlCapacity")>
     <DataMember()> Public Property sReportAddlCapacity() As List(Of String)
        Get
            Return Me._sReportAddlCapacity
        End Get
        Set
            Me._sReportAddlCapacity = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAssumption")>
     <DataMember()> Public Property sReportAssumption() As List(Of String)
        Get
            Return Me._sReportAssumption
        End Get
        Set
            Me._sReportAssumption = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAppurt1Note1")>
     <DataMember()> Public Property sReportAppurt1Note1() As String
        Get
            Return Me._sReportAppurt1Note1
        End Get
        Set
            Me._sReportAppurt1Note1 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAppurt1Note2")>
     <DataMember()> Public Property sReportAppurt1Note2() As String
        Get
            Return Me._sReportAppurt1Note2
        End Get
        Set
            Me._sReportAppurt1Note2 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAppurt1Note3")>
     <DataMember()> Public Property sReportAppurt1Note3() As String
        Get
            Return Me._sReportAppurt1Note3
        End Get
        Set
            Me._sReportAppurt1Note3 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAppurt1Note4")>
     <DataMember()> Public Property sReportAppurt1Note4() As String
        Get
            Return Me._sReportAppurt1Note4
        End Get
        Set
            Me._sReportAppurt1Note4 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAppurt1Note5")>
     <DataMember()> Public Property sReportAppurt1Note5() As String
        Get
            Return Me._sReportAppurt1Note5
        End Get
        Set
            Me._sReportAppurt1Note5 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAppurt1Note6")>
     <DataMember()> Public Property sReportAppurt1Note6() As String
        Get
            Return Me._sReportAppurt1Note6
        End Get
        Set
            Me._sReportAppurt1Note6 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAppurt1Note7")>
     <DataMember()> Public Property sReportAppurt1Note7() As String
        Get
            Return Me._sReportAppurt1Note7
        End Get
        Set
            Me._sReportAppurt1Note7 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAppurt2Note1")>
     <DataMember()> Public Property sReportAppurt2Note1() As String
        Get
            Return Me._sReportAppurt2Note1
        End Get
        Set
            Me._sReportAppurt2Note1 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAppurt2Note2")>
     <DataMember()> Public Property sReportAppurt2Note2() As String
        Get
            Return Me._sReportAppurt2Note2
        End Get
        Set
            Me._sReportAppurt2Note2 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAppurt2Note3")>
     <DataMember()> Public Property sReportAppurt2Note3() As String
        Get
            Return Me._sReportAppurt2Note3
        End Get
        Set
            Me._sReportAppurt2Note3 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAppurt2Note4")>
     <DataMember()> Public Property sReportAppurt2Note4() As String
        Get
            Return Me._sReportAppurt2Note4
        End Get
        Set
            Me._sReportAppurt2Note4 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAppurt2Note5")>
     <DataMember()> Public Property sReportAppurt2Note5() As String
        Get
            Return Me._sReportAppurt2Note5
        End Get
        Set
            Me._sReportAppurt2Note5 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAppurt2Note6")>
     <DataMember()> Public Property sReportAppurt2Note6() As String
        Get
            Return Me._sReportAppurt2Note6
        End Get
        Set
            Me._sReportAppurt2Note6 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAppurt2Note7")>
     <DataMember()> Public Property sReportAppurt2Note7() As String
        Get
            Return Me._sReportAppurt2Note7
        End Get
        Set
            Me._sReportAppurt2Note7 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAddlCapacityNote1")>
     <DataMember()> Public Property sReportAddlCapacityNote1() As String
        Get
            Return Me._sReportAddlCapacityNote1
        End Get
        Set
            Me._sReportAddlCapacityNote1 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAddlCapacityNote2")>
     <DataMember()> Public Property sReportAddlCapacityNote2() As String
        Get
            Return Me._sReportAddlCapacityNote2
        End Get
        Set
            Me._sReportAddlCapacityNote2 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAddlCapacityNote3")>
     <DataMember()> Public Property sReportAddlCapacityNote3() As String
        Get
            Return Me._sReportAddlCapacityNote3
        End Get
        Set
            Me._sReportAddlCapacityNote3 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAddlCapacityNote4")>
     <DataMember()> Public Property sReportAddlCapacityNote4() As String
        Get
            Return Me._sReportAddlCapacityNote4
        End Get
        Set
            Me._sReportAddlCapacityNote4 = Value
        End Set
    End Property

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
        Dim otherToCompare As tnxCCIReport = TryCast(other, tnxCCIReport)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.sReportProjectNumber.CheckChange(otherToCompare.sReportProjectNumber, changes, categoryName, "Sreportprojectnumber"), Equals, False)
        Equals = If(Me.sReportJobType.CheckChange(otherToCompare.sReportJobType, changes, categoryName, "Sreportjobtype"), Equals, False)
        Equals = If(Me.sReportCarrierName.CheckChange(otherToCompare.sReportCarrierName, changes, categoryName, "Sreportcarriername"), Equals, False)
        Equals = If(Me.sReportCarrierSiteNumber.CheckChange(otherToCompare.sReportCarrierSiteNumber, changes, categoryName, "Sreportcarriersitenumber"), Equals, False)
        Equals = If(Me.sReportCarrierSiteName.CheckChange(otherToCompare.sReportCarrierSiteName, changes, categoryName, "Sreportcarriersitename"), Equals, False)
        Equals = If(Me.sReportSiteAddress.CheckChange(otherToCompare.sReportSiteAddress, changes, categoryName, "Sreportsiteaddress"), Equals, False)
        Equals = If(Me.sReportLatitudeDegree.CheckChange(otherToCompare.sReportLatitudeDegree, changes, categoryName, "Sreportlatitudedegree"), Equals, False)
        Equals = If(Me.sReportLatitudeMinute.CheckChange(otherToCompare.sReportLatitudeMinute, changes, categoryName, "Sreportlatitudeminute"), Equals, False)
        Equals = If(Me.sReportLatitudeSecond.CheckChange(otherToCompare.sReportLatitudeSecond, changes, categoryName, "Sreportlatitudesecond"), Equals, False)
        Equals = If(Me.sReportLongitudeDegree.CheckChange(otherToCompare.sReportLongitudeDegree, changes, categoryName, "Sreportlongitudedegree"), Equals, False)
        Equals = If(Me.sReportLongitudeMinute.CheckChange(otherToCompare.sReportLongitudeMinute, changes, categoryName, "Sreportlongitudeminute"), Equals, False)
        Equals = If(Me.sReportLongitudeSecond.CheckChange(otherToCompare.sReportLongitudeSecond, changes, categoryName, "Sreportlongitudesecond"), Equals, False)
        Equals = If(Me.sReportLocalCodeRequirement.CheckChange(otherToCompare.sReportLocalCodeRequirement, changes, categoryName, "Sreportlocalcoderequirement"), Equals, False)
        Equals = If(Me.sReportSiteHistory.CheckChange(otherToCompare.sReportSiteHistory, changes, categoryName, "Sreportsitehistory"), Equals, False)
        Equals = If(Me.sReportTowerManufacturer.CheckChange(otherToCompare.sReportTowerManufacturer, changes, categoryName, "Sreporttowermanufacturer"), Equals, False)
        Equals = If(Me.sReportMonthManufactured.CheckChange(otherToCompare.sReportMonthManufactured, changes, categoryName, "Sreportmonthmanufactured"), Equals, False)
        Equals = If(Me.sReportYearManufactured.CheckChange(otherToCompare.sReportYearManufactured, changes, categoryName, "Sreportyearmanufactured"), Equals, False)
        Equals = If(Me.sReportOriginalSpeed.CheckChange(otherToCompare.sReportOriginalSpeed, changes, categoryName, "Sreportoriginalspeed"), Equals, False)
        Equals = If(Me.sReportOriginalCode.CheckChange(otherToCompare.sReportOriginalCode, changes, categoryName, "Sreportoriginalcode"), Equals, False)
        Equals = If(Me.sReportTowerType.CheckChange(otherToCompare.sReportTowerType, changes, categoryName, "Sreporttowertype"), Equals, False)
        Equals = If(Me.sReportEngrName.CheckChange(otherToCompare.sReportEngrName, changes, categoryName, "Sreportengrname"), Equals, False)
        Equals = If(Me.sReportEngrTitle.CheckChange(otherToCompare.sReportEngrTitle, changes, categoryName, "Sreportengrtitle"), Equals, False)
        Equals = If(Me.sReportHQPhoneNumber.CheckChange(otherToCompare.sReportHQPhoneNumber, changes, categoryName, "Sreporthqphonenumber"), Equals, False)
        Equals = If(Me.sReportEmailAddress.CheckChange(otherToCompare.sReportEmailAddress, changes, categoryName, "Sreportemailaddress"), Equals, False)
        Equals = If(Me.sReportLogoPath.CheckChange(otherToCompare.sReportLogoPath, changes, categoryName, "Sreportlogopath"), Equals, False)
        Equals = If(Me.sReportCCiContactName.CheckChange(otherToCompare.sReportCCiContactName, changes, categoryName, "Sreportccicontactname"), Equals, False)
        Equals = If(Me.sReportCCiAddress1.CheckChange(otherToCompare.sReportCCiAddress1, changes, categoryName, "Sreportcciaddress1"), Equals, False)
        Equals = If(Me.sReportCCiAddress2.CheckChange(otherToCompare.sReportCCiAddress2, changes, categoryName, "Sreportcciaddress2"), Equals, False)
        Equals = If(Me.sReportCCiBUNumber.CheckChange(otherToCompare.sReportCCiBUNumber, changes, categoryName, "Sreportccibunumber"), Equals, False)
        Equals = If(Me.sReportCCiSiteName.CheckChange(otherToCompare.sReportCCiSiteName, changes, categoryName, "Sreportccisitename"), Equals, False)
        Equals = If(Me.sReportCCiJDENumber.CheckChange(otherToCompare.sReportCCiJDENumber, changes, categoryName, "Sreportccijdenumber"), Equals, False)
        Equals = If(Me.sReportCCiWONumber.CheckChange(otherToCompare.sReportCCiWONumber, changes, categoryName, "Sreportcciwonumber"), Equals, False)
        Equals = If(Me.sReportCCiPONumber.CheckChange(otherToCompare.sReportCCiPONumber, changes, categoryName, "Sreportcciponumber"), Equals, False)
        Equals = If(Me.sReportCCiAppNumber.CheckChange(otherToCompare.sReportCCiAppNumber, changes, categoryName, "Sreportcciappnumber"), Equals, False)
        Equals = If(Me.sReportCCiRevNumber.CheckChange(otherToCompare.sReportCCiRevNumber, changes, categoryName, "Sreportccirevnumber"), Equals, False)
        Equals = If(Me.sReportDocsProvided.CheckChange(otherToCompare.sReportDocsProvided, changes, categoryName, "Sreportdocsprovided"), Equals, False)
        Equals = If(Me.sReportRecommendations.CheckChange(otherToCompare.sReportRecommendations, changes, categoryName, "Sreportrecommendations"), Equals, False)
        Equals = If(Me.sReportAppurt1.CheckChange(otherToCompare.sReportAppurt1, changes, categoryName, "Sreportappurt1"), Equals, False)
        Equals = If(Me.sReportAppurt2.CheckChange(otherToCompare.sReportAppurt2, changes, categoryName, "Sreportappurt2"), Equals, False)
        Equals = If(Me.sReportAppurt3.CheckChange(otherToCompare.sReportAppurt3, changes, categoryName, "Sreportappurt3"), Equals, False)
        Equals = If(Me.sReportAddlCapacity.CheckChange(otherToCompare.sReportAddlCapacity, changes, categoryName, "Sreportaddlcapacity"), Equals, False)
        Equals = If(Me.sReportAssumption.CheckChange(otherToCompare.sReportAssumption, changes, categoryName, "Sreportassumption"), Equals, False)
        Equals = If(Me.sReportAppurt1Note1.CheckChange(otherToCompare.sReportAppurt1Note1, changes, categoryName, "Sreportappurt1Note1"), Equals, False)
        Equals = If(Me.sReportAppurt1Note2.CheckChange(otherToCompare.sReportAppurt1Note2, changes, categoryName, "Sreportappurt1Note2"), Equals, False)
        Equals = If(Me.sReportAppurt1Note3.CheckChange(otherToCompare.sReportAppurt1Note3, changes, categoryName, "Sreportappurt1Note3"), Equals, False)
        Equals = If(Me.sReportAppurt1Note4.CheckChange(otherToCompare.sReportAppurt1Note4, changes, categoryName, "Sreportappurt1Note4"), Equals, False)
        Equals = If(Me.sReportAppurt1Note5.CheckChange(otherToCompare.sReportAppurt1Note5, changes, categoryName, "Sreportappurt1Note5"), Equals, False)
        Equals = If(Me.sReportAppurt1Note6.CheckChange(otherToCompare.sReportAppurt1Note6, changes, categoryName, "Sreportappurt1Note6"), Equals, False)
        Equals = If(Me.sReportAppurt1Note7.CheckChange(otherToCompare.sReportAppurt1Note7, changes, categoryName, "Sreportappurt1Note7"), Equals, False)
        Equals = If(Me.sReportAppurt2Note1.CheckChange(otherToCompare.sReportAppurt2Note1, changes, categoryName, "Sreportappurt2Note1"), Equals, False)
        Equals = If(Me.sReportAppurt2Note2.CheckChange(otherToCompare.sReportAppurt2Note2, changes, categoryName, "Sreportappurt2Note2"), Equals, False)
        Equals = If(Me.sReportAppurt2Note3.CheckChange(otherToCompare.sReportAppurt2Note3, changes, categoryName, "Sreportappurt2Note3"), Equals, False)
        Equals = If(Me.sReportAppurt2Note4.CheckChange(otherToCompare.sReportAppurt2Note4, changes, categoryName, "Sreportappurt2Note4"), Equals, False)
        Equals = If(Me.sReportAppurt2Note5.CheckChange(otherToCompare.sReportAppurt2Note5, changes, categoryName, "Sreportappurt2Note5"), Equals, False)
        Equals = If(Me.sReportAppurt2Note6.CheckChange(otherToCompare.sReportAppurt2Note6, changes, categoryName, "Sreportappurt2Note6"), Equals, False)
        Equals = If(Me.sReportAppurt2Note7.CheckChange(otherToCompare.sReportAppurt2Note7, changes, categoryName, "Sreportappurt2Note7"), Equals, False)
        Equals = If(Me.sReportAddlCapacityNote1.CheckChange(otherToCompare.sReportAddlCapacityNote1, changes, categoryName, "Sreportaddlcapacitynote1"), Equals, False)
        Equals = If(Me.sReportAddlCapacityNote2.CheckChange(otherToCompare.sReportAddlCapacityNote2, changes, categoryName, "Sreportaddlcapacitynote2"), Equals, False)
        Equals = If(Me.sReportAddlCapacityNote3.CheckChange(otherToCompare.sReportAddlCapacityNote3, changes, categoryName, "Sreportaddlcapacitynote3"), Equals, False)
        Equals = If(Me.sReportAddlCapacityNote4.CheckChange(otherToCompare.sReportAddlCapacityNote4, changes, categoryName, "Sreportaddlcapacitynote4"), Equals, False)

        Return Equals

    End Function

End Class
