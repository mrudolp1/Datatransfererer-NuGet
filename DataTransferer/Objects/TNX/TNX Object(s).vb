'Option Strict On
'Option Compare Binary 'Trying to speed up parsing the TNX file by using Binary Text comparison instead of Text Comparison

Imports System.ComponentModel
Imports System.Data
Imports System.IO
Imports System.Xml
Imports System.Xml.Serialization
Imports System.Security.Principal
Imports System.Runtime.CompilerServices
Imports System.Data.SqlClient
Imports MoreLinq
Imports System.Runtime.Serialization

<DataContract()>
Partial Public Class tnxModel
    Inherits EDSObjectWithQueries

#Region "Inheritted"

    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "TNX Model"
        End Get
    End Property
    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "tnx.tnx"
        End Get
    End Property

    Public Overrides Property Results As List(Of EDSResult)
        Get
            Dim AllResults As List(Of EDSResult) = New List(Of EDSResult)()
            '' baseStructure.Select(Function(x) x.Results) returns a List(Of List(Of EDSResult)
            '' baseStructure.Select(Function(x) x.Results).SelectMany(Function (x) x) returns all the EDSResults in one List(Of EDSResult)
            AllResults.AddRange(Me.geometry.baseStructure.Select(Function(x) x.Results).SelectMany(Function(x) x))
            AllResults.AddRange(Me.geometry.upperStructure.Select(Function(x) x.Results).SelectMany(Function(x) x))
            AllResults.AddRange(Me.geometry.guyWires.Select(Function(x) x.Results).SelectMany(Function(x) x))
            Return AllResults
        End Get
        Set(value As List(Of EDSResult))
            MyBase.Results = value
        End Set
    End Property

#End Region

#Region "Define"
    Private _filePath As String
    Private _database As New tnxDatabase(Me)
    Private _settings As New tnxSettings(Me)
    Private _solutionSettings As New tnxSolutionSettings(Me)
    Private _MTOSettings As New tnxMTOSettings(Me)
    Private _reportSettings As New tnxReportSettings(Me)
    Private _CCIReport As New tnxCCIReport(Me)
    Private _code As New tnxCode(Me)
    Private _options As New tnxOptions(Me)
    Private _geometry As New tnxGeometry(Me)
    Private _feedLines As New List(Of tnxFeedLine)
    Private _discreteLoads As New List(Of tnxDiscreteLoad)
    Private _dishes As New List(Of tnxDish)
    Private _userForces As New List(Of tnxUserForce)
    Private _otherLines As New List(Of String())
    Private _ConsiderLoadingEquality As Boolean = True
    Private _ConsiderGeometryEquality As Boolean = True

    <Category("TNX"), Description(""), DisplayName("filePath")>
    <DataMember()> Public Property filePath() As String
        Get
            Return Me._filePath
        End Get
        Set
            Me._filePath = Value
        End Set
    End Property
    <Category("TNX"), Description(""), DisplayName("tnxDatabase")>
    <DataMember()> Public Property database() As tnxDatabase
        Get
            Return Me._database
        End Get
        Set
            Me._database = Value
        End Set
    End Property
    <Category("TNX"), Description(""), DisplayName("Settings")>
    <DataMember()> Public Property settings() As tnxSettings
        Get
            Return Me._settings
        End Get
        Set
            Me._settings = Value
        End Set
    End Property
    <Category("TNX"), Description(""), DisplayName("Solution Settings")>
    <DataMember()> Public Property solutionSettings() As tnxSolutionSettings
        Get
            Return Me._solutionSettings
        End Get
        Set
            Me._solutionSettings = Value
        End Set
    End Property
    <Category("TNX"), Description(""), DisplayName("MTO Settings")>
    <DataMember()> Public Property MTOSettings() As tnxMTOSettings
        Get
            Return Me._MTOSettings
        End Get
        Set
            Me._MTOSettings = Value
        End Set
    End Property
    <Category("TNX"), Description(""), DisplayName("Report Settings")>
    <DataMember()> Public Property reportSettings() As tnxReportSettings
        Get
            Return Me._reportSettings
        End Get
        Set
            Me._reportSettings = Value
        End Set
    End Property
    <Category("TNX"), Description(""), DisplayName("CCI Report")>
    <DataMember()> Public Property CCIReport() As tnxCCIReport
        Get
            Return Me._CCIReport
        End Get
        Set
            Me._CCIReport = Value
        End Set
    End Property
    <Category("TNX"), Description(""), DisplayName("Code")>
    <DataMember()> Public Property code() As tnxCode
        Get
            Return Me._code
        End Get
        Set
            Me._code = Value
        End Set
    End Property
    <Category("TNX"), Description(""), DisplayName("Options")>
    <DataMember()> Public Property options() As tnxOptions
        Get
            Return Me._options
        End Get
        Set
            Me._options = Value
        End Set
    End Property
    <Category("TNX"), Description(""), DisplayName("Geometry")>
    <DataMember()> Public Property geometry() As tnxGeometry
        Get
            Return Me._geometry
        End Get
        Set
            Me._geometry = Value
        End Set
    End Property
    <Category("TNX"), Description(""), DisplayName("Feed Lines")>
    <DataMember()> Public Property feedLines() As List(Of tnxFeedLine)
        Get
            Return Me._feedLines
        End Get
        Set
            Me._feedLines = Value
        End Set
    End Property
    <Category("TNX"), Description(""), DisplayName("Discrete Loads")>
    <DataMember()> Public Property discreteLoads() As List(Of tnxDiscreteLoad)
        Get
            Return Me._discreteLoads
        End Get
        Set
            Me._discreteLoads = Value
        End Set
    End Property
    <Category("TNX"), Description(""), DisplayName("Dishes")>
    <DataMember()> Public Property dishes() As List(Of tnxDish)
        Get
            Return Me._dishes
        End Get
        Set
            Me._dishes = Value
        End Set
    End Property
    <Category("TNX"), Description(""), DisplayName("User Forces")>
    <DataMember()> Public Property userForces() As List(Of tnxUserForce)
        Get
            Return Me._userForces
        End Get
        Set
            Me._userForces = Value
        End Set
    End Property
    <Category("TNX"), Description(""), DisplayName("All the other stuff")>
    <DataMember()> Public Property otherLines() As List(Of String())
        Get
            Return Me._otherLines
        End Get
        Set
            Me._otherLines = Value
        End Set
    End Property

    <Category("Settings"), Description("Consider loading in the equality comparison."), DisplayName("Consider Loading Equality")>
    <DataMember()> Public Property ConsiderLoadingEquality() As Boolean
        Get
            Return _ConsiderLoadingEquality
        End Get
        Set(value As Boolean)
            _ConsiderLoadingEquality = value
        End Set
    End Property

    <Category("Settings"), Description("Consider database item and tower sections in the equality comparison. Disable this when determining if the main TNX table needs to be updated."), DisplayName("Consider Geometry Equality")>
    <DataMember()> Public Property ConsiderGeometryEquality() As Boolean
        Get
            Return _ConsiderGeometryEquality
        End Get
        Set(value As Boolean)
            _ConsiderGeometryEquality = value
            Try
                Me.geometry.ConsiderSectionEquality = value
            Catch
            End Try
        End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
        'Leave method empty
    End Sub

    <Category("Constructor"), Description("Create TNX object from DataSet")>
    Public Sub New(ByRef StructureDS As DataSet, Optional ByVal Parent As EDSObject = Nothing)
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        If StructureDS.Tables.Contains("TNX") And StructureDS.Tables("TNX").Rows.Count > 0 Then
            setIndInputs(StructureDS.Tables("TNX").Rows(0))

            Me.geometry.Absorb(Me)

            If StructureDS.Tables.Contains("Base Structure") Then
                For Each baseSection As DataRow In StructureDS.Tables("Base Structure").Rows
                    Me.geometry.baseStructure.Add(New tnxTowerRecord(baseSection, Me.geometry))
                Next
                Me.geometry.baseStructure.Sort()
            End If

            If StructureDS.Tables.Contains("Upper Structure") Then
                For Each upperSection As DataRow In StructureDS.Tables("Upper Structure").Rows
                    Me.geometry.upperStructure.Add(New tnxAntennaRecord(upperSection, Me.geometry))
                Next
                Me.geometry.upperStructure.Sort()
            End If

            If StructureDS.Tables.Contains("Guys") Then
                For Each guyLevel As DataRow In StructureDS.Tables("Guys").Rows
                    Me.geometry.guyWires.Add(New tnxGuyRecord(guyLevel, Me))
                Next
                Me.geometry.guyWires.Sort()
            End If

            'Discrete
            If StructureDS.Tables.Contains("Discretes") Then
                For Each disc As DataRow In StructureDS.Tables("Discretes").Rows
                    Me.discreteLoads.Add(New tnxDiscreteLoad(disc, Me))
                Next
                Me.discreteLoads.Sort()
            End If

            'Dish
            If StructureDS.Tables.Contains("Dishes") Then
                For Each dish As DataRow In StructureDS.Tables("Dishes").Rows
                    Me.dishes.Add(New tnxDish(dish, Me))
                Next
                Me.dishes.Sort()
            End If

            'User Forces
            If StructureDS.Tables.Contains("User Forces") Then
                For Each user As DataRow In StructureDS.Tables("User Forces").Rows
                    Me.userForces.Add(New tnxUserForce(user, Me))
                Next
                Me.userForces.Sort()
            End If

            'Linear
            If StructureDS.Tables.Contains("Lines") Then
                For Each line As DataRow In StructureDS.Tables("Lines").Rows
                    Me.feedLines.Add(New tnxFeedLine(line, Me))
                Next
                Me.feedLines.Sort()
            End If

            If StructureDS.Tables.Contains("Materials") Then
                For Each material As DataRow In StructureDS.Tables("Materials").Rows
                    If Not IsDBNull(material.Item("IsBolt")) Then
                        If CBool(material.Item("IsBolt")) Then
                            Me.database.bolts.Add(New tnxMaterial(material, Me.database))
                        Else
                            Me.database.materials.Add(New tnxMaterial(material, Me.database))
                        End If
                    End If
                Next
            End If

            If StructureDS.Tables.Contains("Members") Then
                For Each member As DataRow In StructureDS.Tables("Members").Rows
                    Me.database.members.Add(New tnxMember(member, Me.database))
                Next
            End If

        End If
    End Sub

    Public Sub setIndInputs(Data As DataRow)

        Me.ID = DBtoNullableInt(Data.Item("ID"))
        Me.settings.projectInfo.DesignStandardSeries = DBtoStr(Data.Item("DesignStandardSeries"))
        Me.settings.projectInfo.UnitsSystem = DBtoStr(Data.Item("UnitsSystem"))
        Me.settings.projectInfo.ClientName = DBtoStr(Data.Item("ClientName"))
        Me.settings.projectInfo.ProjectName = DBtoStr(Data.Item("ProjectName"))
        Me.settings.projectInfo.ProjectNumber = DBtoStr(Data.Item("ProjectNumber"))
        Me.settings.projectInfo.CreatedBy = DBtoStr(Data.Item("CreatedBy"))
        Me.settings.projectInfo.CreatedOn = DBtoStr(Data.Item("CreatedOn"))
        Me.settings.projectInfo.LastUsedBy = DBtoStr(Data.Item("LastUsedBy"))
        Me.settings.projectInfo.LastUsedOn = DBtoStr(Data.Item("LastUsedOn"))
        Me.settings.projectInfo.VersionUsed = DBtoStr(Data.Item("VersionUsed"))
        Me.settings.USUnits.Length.value = DBtoStr(Data.Item("USLength"))
        Me.settings.USUnits.Length.precision = DBtoNullableInt(Data.Item("USLengthPrec"))
        Me.settings.USUnits.Coordinate.value = DBtoStr(Data.Item("USCoordinate"))
        Me.settings.USUnits.Coordinate.precision = DBtoNullableInt(Data.Item("USCoordinatePrec"))
        Me.settings.USUnits.Force.value = DBtoStr(Data.Item("USForce"))
        Me.settings.USUnits.Force.precision = DBtoNullableInt(Data.Item("USForcePrec"))
        Me.settings.USUnits.Load.value = DBtoStr(Data.Item("USLoad"))
        Me.settings.USUnits.Load.precision = DBtoNullableInt(Data.Item("USLoadPrec"))
        Me.settings.USUnits.Moment.value = DBtoStr(Data.Item("USMoment"))
        Me.settings.USUnits.Moment.precision = DBtoNullableInt(Data.Item("USMomentPrec"))
        Me.settings.USUnits.Properties.value = DBtoStr(Data.Item("USProperties"))
        Me.settings.USUnits.Properties.precision = DBtoNullableInt(Data.Item("USPropertiesPrec"))
        Me.settings.USUnits.Pressure.value = DBtoStr(Data.Item("USPressure"))
        Me.settings.USUnits.Pressure.precision = DBtoNullableInt(Data.Item("USPressurePrec"))
        Me.settings.USUnits.Velocity.value = DBtoStr(Data.Item("USVelocity"))
        Me.settings.USUnits.Velocity.precision = DBtoNullableInt(Data.Item("USVelocityPrec"))
        Me.settings.USUnits.Displacement.value = DBtoStr(Data.Item("USDisplacement"))
        Me.settings.USUnits.Displacement.precision = DBtoNullableInt(Data.Item("USDisplacementPrec"))
        Me.settings.USUnits.Mass.value = DBtoStr(Data.Item("USMass"))
        Me.settings.USUnits.Mass.precision = DBtoNullableInt(Data.Item("USMassPrec"))
        Me.settings.USUnits.Acceleration.value = DBtoStr(Data.Item("USAcceleration"))
        Me.settings.USUnits.Acceleration.precision = DBtoNullableInt(Data.Item("USAccelerationPrec"))
        Me.settings.USUnits.Stress.value = DBtoStr(Data.Item("USStress"))
        Me.settings.USUnits.Stress.precision = DBtoNullableInt(Data.Item("USStressPrec"))
        Me.settings.USUnits.Density.value = DBtoStr(Data.Item("USDensity"))
        Me.settings.USUnits.Density.precision = DBtoNullableInt(Data.Item("USDensityPrec"))
        Me.settings.USUnits.UnitWt.value = DBtoStr(Data.Item("USUnitWt"))
        Me.settings.USUnits.UnitWt.precision = DBtoNullableInt(Data.Item("USUnitWtPrec"))
        Me.settings.USUnits.Strength.value = DBtoStr(Data.Item("USStrength"))
        Me.settings.USUnits.Strength.precision = DBtoNullableInt(Data.Item("USStrengthPrec"))
        Me.settings.USUnits.Modulus.value = DBtoStr(Data.Item("USModulus"))
        Me.settings.USUnits.Modulus.precision = DBtoNullableInt(Data.Item("USModulusPrec"))
        Me.settings.USUnits.Temperature.value = DBtoStr(Data.Item("USTemperature"))
        Me.settings.USUnits.Temperature.precision = DBtoNullableInt(Data.Item("USTemperaturePrec"))
        Me.settings.USUnits.Printer.value = DBtoStr(Data.Item("USPrinter"))
        Me.settings.USUnits.Printer.precision = DBtoNullableInt(Data.Item("USPrinterPrec"))
        Me.settings.USUnits.Rotation.value = DBtoStr(Data.Item("USRotation"))
        Me.settings.USUnits.Rotation.precision = DBtoNullableInt(Data.Item("USRotationPrec"))
        Me.settings.USUnits.Spacing.value = DBtoStr(Data.Item("USSpacing"))
        Me.settings.USUnits.Spacing.precision = DBtoNullableInt(Data.Item("USSpacingPrec"))
        Me.settings.userInfo.ViewerUserName = DBtoStr(Data.Item("ViewerUserName"))
        Me.settings.userInfo.ViewerCompanyName = DBtoStr(Data.Item("ViewerCompanyName"))
        Me.settings.userInfo.ViewerStreetAddress = DBtoStr(Data.Item("ViewerStreetAddress"))
        Me.settings.userInfo.ViewerCityState = DBtoStr(Data.Item("ViewerCityState"))
        Me.settings.userInfo.ViewerPhone = DBtoStr(Data.Item("ViewerPhone"))
        Me.settings.userInfo.ViewerFAX = DBtoStr(Data.Item("ViewerFAX"))
        Me.settings.userInfo.ViewerLogo = DBtoStr(Data.Item("ViewerLogo"))
        Me.settings.userInfo.ViewerCompanyBitmap = DBtoStr(Data.Item("ViewerCompanyBitmap"))
        Me.CCIReport.sReportProjectNumber = DBtoStr(Data.Item("sReportProjectNumber"))
        Me.CCIReport.sReportJobType = DBtoStr(Data.Item("sReportJobType"))
        Me.CCIReport.sReportCarrierName = DBtoStr(Data.Item("sReportCarrierName"))
        Me.CCIReport.sReportCarrierSiteNumber = DBtoStr(Data.Item("sReportCarrierSiteNumber"))
        Me.CCIReport.sReportCarrierSiteName = DBtoStr(Data.Item("sReportCarrierSiteName"))
        Me.CCIReport.sReportSiteAddress = DBtoStr(Data.Item("sReportSiteAddress"))
        Me.CCIReport.sReportLatitudeDegree = DBtoNullableDbl(Data.Item("sReportLatitudeDegree"), 6)
        Me.CCIReport.sReportLatitudeMinute = DBtoNullableDbl(Data.Item("sReportLatitudeMinute"), 6)
        Me.CCIReport.sReportLatitudeSecond = DBtoNullableDbl(Data.Item("sReportLatitudeSecond"), 6)
        Me.CCIReport.sReportLongitudeDegree = DBtoNullableDbl(Data.Item("sReportLongitudeDegree"), 6)
        Me.CCIReport.sReportLongitudeMinute = DBtoNullableDbl(Data.Item("sReportLongitudeMinute"), 6)
        Me.CCIReport.sReportLongitudeSecond = DBtoNullableDbl(Data.Item("sReportLongitudeSecond"), 6)
        Me.CCIReport.sReportLocalCodeRequirement = DBtoStr(Data.Item("sReportLocalCodeRequirement"))
        Me.CCIReport.sReportSiteHistory = DBtoStr(Data.Item("sReportSiteHistory"))
        Me.CCIReport.sReportTowerManufacturer = DBtoStr(Data.Item("sReportTowerManufacturer"))
        Me.CCIReport.sReportMonthManufactured = DBtoStr(Data.Item("sReportMonthManufactured"))
        Me.CCIReport.sReportYearManufactured = DBtoNullableInt(Data.Item("sReportYearManufactured"))
        Me.CCIReport.sReportOriginalSpeed = DBtoNullableDbl(Data.Item("sReportOriginalSpeed"), 6)
        Me.CCIReport.sReportOriginalCode = DBtoStr(Data.Item("sReportOriginalCode"))
        Me.CCIReport.sReportTowerType = DBtoStr(Data.Item("sReportTowerType"))
        Me.CCIReport.sReportEngrName = DBtoStr(Data.Item("sReportEngrName"))
        Me.CCIReport.sReportEngrTitle = DBtoStr(Data.Item("sReportEngrTitle"))
        Me.CCIReport.sReportHQPhoneNumber = DBtoStr(Data.Item("sReportHQPhoneNumber"))
        Me.CCIReport.sReportEmailAddress = DBtoStr(Data.Item("sReportEmailAddress"))
        Me.CCIReport.sReportLogoPath = DBtoStr(Data.Item("sReportLogoPath"))
        Me.CCIReport.sReportCCiContactName = DBtoStr(Data.Item("sReportCCiContactName"))
        Me.CCIReport.sReportCCiAddress1 = DBtoStr(Data.Item("sReportCCiAddress1"))
        Me.CCIReport.sReportCCiAddress2 = DBtoStr(Data.Item("sReportCCiAddress2"))
        Me.CCIReport.sReportCCiBUNumber = DBtoStr(Data.Item("sReportCCiBUNumber"))
        Me.CCIReport.sReportCCiSiteName = DBtoStr(Data.Item("sReportCCiSiteName"))
        Me.CCIReport.sReportCCiJDENumber = DBtoStr(Data.Item("sReportCCiJDENumber"))
        Me.CCIReport.sReportCCiWONumber = DBtoStr(Data.Item("sReportCCiWONumber"))
        Me.CCIReport.sReportCCiPONumber = DBtoStr(Data.Item("sReportCCiPONumber"))
        Me.CCIReport.sReportCCiAppNumber = DBtoStr(Data.Item("sReportCCiAppNumber"))
        Me.CCIReport.sReportCCiRevNumber = DBtoStr(Data.Item("sReportCCiRevNumber"))
        Me.CCIReport.sReportRecommendations = DBtoStr(Data.Item("sReportRecommendations"))
        Me.CCIReport.sReportAppurt1Note1 = DBtoStr(Data.Item("sReportAppurt1Note1"))
        Me.CCIReport.sReportAppurt1Note2 = DBtoStr(Data.Item("sReportAppurt1Note2"))
        Me.CCIReport.sReportAppurt1Note3 = DBtoStr(Data.Item("sReportAppurt1Note3"))
        Me.CCIReport.sReportAppurt1Note4 = DBtoStr(Data.Item("sReportAppurt1Note4"))
        Me.CCIReport.sReportAppurt1Note5 = DBtoStr(Data.Item("sReportAppurt1Note5"))
        Me.CCIReport.sReportAppurt1Note6 = DBtoStr(Data.Item("sReportAppurt1Note6"))
        Me.CCIReport.sReportAppurt1Note7 = DBtoStr(Data.Item("sReportAppurt1Note7"))
        Me.CCIReport.sReportAppurt2Note1 = DBtoStr(Data.Item("sReportAppurt2Note1"))
        Me.CCIReport.sReportAppurt2Note2 = DBtoStr(Data.Item("sReportAppurt2Note2"))
        Me.CCIReport.sReportAppurt2Note3 = DBtoStr(Data.Item("sReportAppurt2Note3"))
        Me.CCIReport.sReportAppurt2Note4 = DBtoStr(Data.Item("sReportAppurt2Note4"))
        Me.CCIReport.sReportAppurt2Note5 = DBtoStr(Data.Item("sReportAppurt2Note5"))
        Me.CCIReport.sReportAppurt2Note6 = DBtoStr(Data.Item("sReportAppurt2Note6"))
        Me.CCIReport.sReportAppurt2Note7 = DBtoStr(Data.Item("sReportAppurt2Note7"))
        Me.CCIReport.sReportAddlCapacityNote1 = DBtoStr(Data.Item("sReportAddlCapacityNote1"))
        Me.CCIReport.sReportAddlCapacityNote2 = DBtoStr(Data.Item("sReportAddlCapacityNote2"))
        Me.CCIReport.sReportAddlCapacityNote3 = DBtoStr(Data.Item("sReportAddlCapacityNote3"))
        Me.CCIReport.sReportAddlCapacityNote4 = DBtoStr(Data.Item("sReportAddlCapacityNote4"))
        Me.code.design.DesignCode = DBtoStr(Data.Item("DesignCode"))
        Me.geometry.TowerType = DBtoStr(Data.Item("TowerType"))
        Me.geometry.AntennaType = DBtoStr(Data.Item("AntennaType"))
        Me.geometry.OverallHeight = DBtoNullableDbl(Data.Item("OverallHeight"), 6)
        Me.geometry.BaseElevation = DBtoNullableDbl(Data.Item("BaseElevation"), 6)
        Me.geometry.Lambda = DBtoNullableDbl(Data.Item("Lambda"), 6)
        Me.geometry.TowerTopFaceWidth = DBtoNullableDbl(Data.Item("TowerTopFaceWidth"), 6)
        Me.geometry.TowerBaseFaceWidth = DBtoNullableDbl(Data.Item("TowerBaseFaceWidth"), 6)
        Me.code.wind.WindSpeed = DBtoNullableDbl(Data.Item("WindSpeed"), 6)
        Me.code.wind.WindSpeedIce = DBtoNullableDbl(Data.Item("WindSpeedIce"), 6)
        Me.code.wind.WindSpeedService = DBtoNullableDbl(Data.Item("WindSpeedService"), 6)
        Me.code.ice.IceThickness = DBtoNullableDbl(Data.Item("IceThickness"), 6)
        Me.code.wind.CSA_S37_RefVelPress = DBtoNullableDbl(Data.Item("CSA_S37_RefVelPress"), 6)
        Me.code.wind.CSA_S37_ReliabilityClass = DBtoNullableInt(Data.Item("CSA_S37_ReliabilityClass"))
        Me.code.wind.CSA_S37_ServiceabilityFactor = DBtoNullableDbl(Data.Item("CSA_S37_ServiceabilityFactor"), 6)
        Me.code.ice.UseModified_TIA_222_IceParameters = DBtoNullableBool(Data.Item("UseModified_TIA_222_IceParameters"))
        Me.code.ice.TIA_222_IceThicknessMultiplier = DBtoNullableDbl(Data.Item("TIA_222_IceThicknessMultiplier"), 6)
        Me.code.ice.DoNotUse_TIA_222_IceEscalation = DBtoNullableBool(Data.Item("DoNotUse_TIA_222_IceEscalation"))
        Me.code.ice.IceDensity = DBtoNullableDbl(Data.Item("IceDensity"), 6)
        Me.code.seismic.SeismicSiteClass = DBtoNullableInt(Data.Item("SeismicSiteClass"))
        Me.code.seismic.SeismicSs = DBtoNullableDbl(Data.Item("SeismicSs"), 6)
        Me.code.seismic.SeismicS1 = DBtoNullableDbl(Data.Item("SeismicS1"), 6)
        Me.code.thermal.TempDrop = DBtoNullableDbl(Data.Item("TempDrop"), 6)
        Me.code.misclCode.GroutFc = DBtoNullableDbl(Data.Item("GroutFc"), 6)
        Me.options.defaultGirtOffsets.GirtOffset = DBtoNullableDbl(Data.Item("GirtOffset"), 6)
        Me.options.defaultGirtOffsets.GirtOffsetLatticedPole = DBtoNullableDbl(Data.Item("GirtOffsetLatticedPole"), 6)
        Me.options.foundationStiffness.MastVert = DBtoNullableDbl(Data.Item("MastVert"), 6)
        Me.options.foundationStiffness.MastHorz = DBtoNullableDbl(Data.Item("MastHorz"), 6)
        Me.options.foundationStiffness.GuyVert = DBtoNullableDbl(Data.Item("GuyVert"), 6)
        Me.options.foundationStiffness.GuyHorz = DBtoNullableDbl(Data.Item("GuyHorz"), 6)
        Me.options.misclOptions.HogRodTakeup = DBtoNullableDbl(Data.Item("HogRodTakeup"), 6)
        Me.geometry.TowerTaper = DBtoStr(Data.Item("TowerTaper"))
        Me.geometry.GuyedMonopoleBaseType = DBtoStr(Data.Item("GuyedMonopoleBaseType"))
        Me.geometry.TaperHeight = DBtoNullableDbl(Data.Item("TaperHeight"), 6)
        Me.geometry.PivotHeight = DBtoNullableDbl(Data.Item("PivotHeight"), 6)
        Me.geometry.AutoCalcGH = DBtoNullableBool(Data.Item("AutoCalcGH"))
        Me.MTOSettings.IncludeCapacityNote = DBtoNullableBool(Data.Item("IncludeCapacityNote"))
        Me.MTOSettings.IncludeAppurtGraphics = DBtoNullableBool(Data.Item("IncludeAppurtGraphics"))
        Me.MTOSettings.DisplayNotes = DBtoNullableBool(Data.Item("DisplayNotes"))
        Me.MTOSettings.DisplayReactions = DBtoNullableBool(Data.Item("DisplayReactions"))
        Me.MTOSettings.DisplaySchedule = DBtoNullableBool(Data.Item("DisplaySchedule"))
        Me.MTOSettings.DisplayAppurtenanceTable = DBtoNullableBool(Data.Item("DisplayAppurtenanceTable"))
        Me.MTOSettings.DisplayMaterialStrengthTable = DBtoNullableBool(Data.Item("DisplayMaterialStrengthTable"))
        Me.code.wind.AutoCalc_ASCE_GH = DBtoNullableBool(Data.Item("AutoCalc_ASCE_GH"))
        Me.code.wind.ASCE_ExposureCat = DBtoNullableInt(Data.Item("ASCE_ExposureCat"))
        Me.code.wind.ASCE_Year = DBtoNullableInt(Data.Item("ASCE_Year"))
        Me.code.wind.ASCEGh = DBtoNullableDbl(Data.Item("ASCEGh"), 6)
        Me.code.wind.ASCEI = DBtoNullableDbl(Data.Item("ASCEI"), 6)
        Me.code.wind.UseASCEWind = DBtoNullableBool(Data.Item("UseASCEWind"))
        Me.geometry.UserGHElev = DBtoNullableDbl(Data.Item("UserGHElev"), 6)
        Me.code.design.UseCodeGuySF = DBtoNullableBool(Data.Item("UseCodeGuySF"))
        Me.code.design.GuySF = DBtoNullableDbl(Data.Item("GuySF"), 6)
        Me.code.wind.CalcWindAt = DBtoNullableInt(Data.Item("CalcWindAt"))
        Me.code.misclCode.TowerBoltGrade = DBtoStr(Data.Item("TowerBoltGrade"))
        Me.code.misclCode.TowerBoltMinEdgeDist = DBtoNullableDbl(Data.Item("TowerBoltMinEdgeDist"), 6)
        Me.code.design.AllowStressRatio = DBtoNullableDbl(Data.Item("AllowStressRatio"), 6)
        Me.code.design.AllowAntStressRatio = DBtoNullableDbl(Data.Item("AllowAntStressRatio"), 6)
        Me.code.wind.WindCalcPoints = DBtoNullableDbl(Data.Item("WindCalcPoints"), 6)
        Me.geometry.UseIndexPlate = DBtoNullableBool(Data.Item("UseIndexPlate"))
        Me.geometry.EnterUserDefinedGhValues = DBtoNullableBool(Data.Item("EnterUserDefinedGhValues"))
        Me.geometry.BaseTowerGhInput = DBtoNullableDbl(Data.Item("BaseTowerGhInput"), 6)
        Me.geometry.UpperStructureGhInput = DBtoNullableDbl(Data.Item("UpperStructureGhInput"), 6)
        Me.geometry.EnterUserDefinedCgValues = DBtoNullableBool(Data.Item("EnterUserDefinedCgValues"))
        Me.geometry.BaseTowerCgInput = DBtoNullableDbl(Data.Item("BaseTowerCgInput"), 6)
        Me.geometry.UpperStructureCgInput = DBtoNullableDbl(Data.Item("UpperStructureCgInput"), 6)
        Me.options.cantileverPoles.CheckVonMises = DBtoNullableBool(Data.Item("CheckVonMises"))
        Me.options.UseClearSpans = DBtoNullableBool(Data.Item("UseClearSpans"))
        Me.options.UseClearSpansKlr = DBtoNullableBool(Data.Item("UseClearSpansKlr"))
        Me.geometry.AntennaFaceWidth = DBtoNullableDbl(Data.Item("AntennaFaceWidth"), 6)
        Me.code.design.DoInteraction = DBtoNullableBool(Data.Item("DoInteraction"))
        Me.code.design.DoHorzInteraction = DBtoNullableBool(Data.Item("DoHorzInteraction"))
        Me.code.design.DoDiagInteraction = DBtoNullableBool(Data.Item("DoDiagInteraction"))
        Me.code.design.UseMomentMagnification = DBtoNullableBool(Data.Item("UseMomentMagnification"))
        Me.options.UseFeedlineAsCylinder = DBtoNullableBool(Data.Item("UseFeedlineAsCylinder"))
        Me.options.defaultGirtOffsets.OffsetBotGirt = DBtoNullableBool(Data.Item("OffsetBotGirt"))
        Me.code.design.PrintBitmaps = DBtoNullableBool(Data.Item("PrintBitmaps"))
        Me.geometry.UseTopTakeup = DBtoNullableBool(Data.Item("UseTopTakeup"))
        Me.geometry.ConstantSlope = DBtoNullableBool(Data.Item("ConstantSlope"))
        Me.code.design.UseCodeStressRatio = DBtoNullableBool(Data.Item("UseCodeStressRatio"))
        Me.options.UseLegLoads = DBtoNullableBool(Data.Item("UseLegLoads"))
        Me.code.design.ERIDesignMode = DBtoStr(Data.Item("ERIDesignMode"))
        Me.code.wind.WindExposure = DBtoNullableInt(Data.Item("WindExposure"))
        Me.code.wind.WindZone = DBtoNullableInt(Data.Item("WindZone"))
        Me.code.wind.StructureCategory = DBtoNullableInt(Data.Item("StructureCategory"))
        Me.code.wind.RiskCategory = DBtoNullableInt(Data.Item("RiskCategory"))
        Me.code.wind.TopoCategory = DBtoNullableInt(Data.Item("TopoCategory"))
        Me.code.wind.RSMTopographicFeature = DBtoNullableInt(Data.Item("RSMTopographicFeature"))
        Me.code.wind.RSM_L = DBtoNullableDbl(Data.Item("RSM_L"), 6)
        Me.code.wind.RSM_X = DBtoNullableDbl(Data.Item("RSM_X"), 6)
        Me.code.wind.CrestHeight = DBtoNullableDbl(Data.Item("CrestHeight"), 6)
        Me.code.wind.TIA_222_H_TopoFeatureDownwind = DBtoNullableBool(Data.Item("TIA_222_H_TopoFeatureDownwind"))
        Me.code.wind.BaseElevAboveSeaLevel = DBtoNullableDbl(Data.Item("BaseElevAboveSeaLevel"), 6)
        Me.code.wind.ConsiderRooftopSpeedUp = DBtoNullableBool(Data.Item("ConsiderRooftopSpeedUp"))
        Me.code.wind.RooftopWS = DBtoNullableDbl(Data.Item("RooftopWS"), 6)
        Me.code.wind.RooftopHS = DBtoNullableDbl(Data.Item("RooftopHS"), 6)
        Me.code.wind.RooftopParapetHt = DBtoNullableDbl(Data.Item("RooftopParapetHt"), 6)
        Me.code.wind.RooftopXB = DBtoNullableDbl(Data.Item("RooftopXB"), 6)
        Me.code.design.UseTIA222H_AnnexS = DBtoNullableBool(Data.Item("UseTIA222H_AnnexS"))
        Me.code.design.TIA_222_H_AnnexS_Ratio = DBtoNullableDbl(Data.Item("TIA_222_H_AnnexS_Ratio"), 6)
        Me.code.wind.EIACWindMult = DBtoNullableDbl(Data.Item("EIACWindMult"), 6)
        Me.code.wind.EIACWindMultIce = DBtoNullableDbl(Data.Item("EIACWindMultIce"), 6)
        Me.code.wind.EIACIgnoreCableDrag = DBtoNullableBool(Data.Item("EIACIgnoreCableDrag"))
        Me.MTOSettings.Notes = DBtoStr(Data.Item("Notes"))
        Me.reportSettings.ReportInputCosts = DBtoNullableBool(Data.Item("ReportInputCosts"))
        Me.reportSettings.ReportInputGeometry = DBtoNullableBool(Data.Item("ReportInputGeometry"))
        Me.reportSettings.ReportInputOptions = DBtoNullableBool(Data.Item("ReportInputOptions"))
        Me.reportSettings.ReportMaxForces = DBtoNullableBool(Data.Item("ReportMaxForces"))
        Me.reportSettings.ReportInputMap = DBtoNullableBool(Data.Item("ReportInputMap"))
        Me.reportSettings.CostReportOutputType = DBtoStr(Data.Item("CostReportOutputType"))
        Me.reportSettings.CapacityReportOutputType = DBtoStr(Data.Item("CapacityReportOutputType"))
        Me.reportSettings.ReportPrintForceTotals = DBtoNullableBool(Data.Item("ReportPrintForceTotals"))
        Me.reportSettings.ReportPrintForceDetails = DBtoNullableBool(Data.Item("ReportPrintForceDetails"))
        Me.reportSettings.ReportPrintMastVectors = DBtoNullableBool(Data.Item("ReportPrintMastVectors"))
        Me.reportSettings.ReportPrintAntPoleVectors = DBtoNullableBool(Data.Item("ReportPrintAntPoleVectors"))
        Me.reportSettings.ReportPrintDiscreteVectors = DBtoNullableBool(Data.Item("ReportPrintDiscreteVectors"))
        Me.reportSettings.ReportPrintDishVectors = DBtoNullableBool(Data.Item("ReportPrintDishVectors"))
        Me.reportSettings.ReportPrintFeedTowerVectors = DBtoNullableBool(Data.Item("ReportPrintFeedTowerVectors"))
        Me.reportSettings.ReportPrintUserLoadVectors = DBtoNullableBool(Data.Item("ReportPrintUserLoadVectors"))
        Me.reportSettings.ReportPrintPressures = DBtoNullableBool(Data.Item("ReportPrintPressures"))
        Me.reportSettings.ReportPrintAppurtForces = DBtoNullableBool(Data.Item("ReportPrintAppurtForces"))
        Me.reportSettings.ReportPrintGuyForces = DBtoNullableBool(Data.Item("ReportPrintGuyForces"))
        Me.reportSettings.ReportPrintGuyStressing = DBtoNullableBool(Data.Item("ReportPrintGuyStressing"))
        Me.reportSettings.ReportPrintDeflections = DBtoNullableBool(Data.Item("ReportPrintDeflections"))
        Me.reportSettings.ReportPrintReactions = DBtoNullableBool(Data.Item("ReportPrintReactions"))
        Me.reportSettings.ReportPrintStressChecks = DBtoNullableBool(Data.Item("ReportPrintStressChecks"))
        Me.reportSettings.ReportPrintBoltChecks = DBtoNullableBool(Data.Item("ReportPrintBoltChecks"))
        Me.reportSettings.ReportPrintInputGVerificationTables = DBtoNullableBool(Data.Item("ReportPrintInputGVerificationTables"))
        Me.reportSettings.ReportPrintOutputGVerificationTables = DBtoNullableBool(Data.Item("ReportPrintOutputGVerificationTables"))
        Me.options.cantileverPoles.SocketTopMount = DBtoNullableBool(Data.Item("SocketTopMount"))
        Me.options.SRTakeCompression = DBtoNullableBool(Data.Item("SRTakeCompression"))
        Me.options.AllLegPanelsSame = DBtoNullableBool(Data.Item("AllLegPanelsSame"))
        Me.options.UseCombinedBoltCapacity = DBtoNullableBool(Data.Item("UseCombinedBoltCapacity"))
        Me.options.SecHorzBracesLeg = DBtoNullableBool(Data.Item("SecHorzBracesLeg"))
        Me.options.SortByComponent = DBtoNullableBool(Data.Item("SortByComponent"))
        Me.options.SRCutEnds = DBtoNullableBool(Data.Item("SRCutEnds"))
        Me.options.SRConcentric = DBtoNullableBool(Data.Item("SRConcentric"))
        Me.options.CalcBlockShear = DBtoNullableBool(Data.Item("CalcBlockShear"))
        Me.options.Use4SidedDiamondBracing = DBtoNullableBool(Data.Item("Use4SidedDiamondBracing"))
        Me.options.TriangulateInnerBracing = DBtoNullableBool(Data.Item("TriangulateInnerBracing"))
        Me.options.PrintCarrierNotes = DBtoNullableBool(Data.Item("PrintCarrierNotes"))
        Me.options.AddIBCWindCase = DBtoNullableBool(Data.Item("AddIBCWindCase"))
        Me.code.wind.UseStateCountyLookup = DBtoNullableBool(Data.Item("UseStateCountyLookup"))
        Me.code.wind.State = DBtoStr(Data.Item("State"))
        Me.code.wind.County = DBtoStr(Data.Item("County"))
        Me.options.LegBoltsAtTop = DBtoNullableBool(Data.Item("LegBoltsAtTop"))
        Me.options.cantileverPoles.PrintMonopoleAtIncrements = DBtoNullableBool(Data.Item("PrintMonopoleAtIncrements"))
        Me.options.UseTIA222Exemptions_MinBracingResistance = DBtoNullableBool(Data.Item("UseTIA222Exemptions_MinBracingResistance"))
        Me.options.UseTIA222Exemptions_TensionSplice = DBtoNullableBool(Data.Item("UseTIA222Exemptions_TensionSplice"))
        Me.options.IgnoreKLryFor60DegAngleLegs = DBtoNullableBool(Data.Item("IgnoreKLryFor60DegAngleLegs"))
        Me.code.wind.ASCE_7_10_WindData = DBtoNullableBool(Data.Item("ASCE_7_10_WindData"))
        Me.code.wind.ASCE_7_10_ConvertWindToASD = DBtoNullableBool(Data.Item("ASCE_7_10_ConvertWindToASD"))
        Me.solutionSettings.SolutionUsePDelta = DBtoNullableBool(Data.Item("SolutionUsePDelta"))
        Me.options.UseFeedlineTorque = DBtoNullableBool(Data.Item("UseFeedlineTorque"))
        Me.options.UsePinnedElements = DBtoNullableBool(Data.Item("UsePinnedElements"))
        Me.code.wind.UseMaxKz = DBtoNullableBool(Data.Item("UseMaxKz"))
        Me.options.UseRigidIndex = DBtoNullableBool(Data.Item("UseRigidIndex"))
        Me.options.UseTrueCable = DBtoNullableBool(Data.Item("UseTrueCable"))
        Me.options.UseASCELy = DBtoNullableBool(Data.Item("UseASCELy"))
        Me.options.CalcBracingForces = DBtoNullableBool(Data.Item("CalcBracingForces"))
        Me.options.IgnoreBracingFEA = DBtoNullableBool(Data.Item("IgnoreBracingFEA"))
        Me.options.cantileverPoles.UseSubCriticalFlow = DBtoNullableBool(Data.Item("UseSubCriticalFlow"))
        Me.options.cantileverPoles.AssumePoleWithNoAttachments = DBtoNullableBool(Data.Item("AssumePoleWithNoAttachments"))
        Me.options.cantileverPoles.AssumePoleWithShroud = DBtoNullableBool(Data.Item("AssumePoleWithShroud"))
        Me.options.cantileverPoles.PoleCornerRadiusKnown = DBtoNullableBool(Data.Item("PoleCornerRadiusKnown"))
        Me.solutionSettings.SolutionMinStiffness = DBtoNullableDbl(Data.Item("SolutionMinStiffness"), 6)
        Me.solutionSettings.SolutionMaxStiffness = DBtoNullableDbl(Data.Item("SolutionMaxStiffness"), 6)
        Me.solutionSettings.SolutionMaxCycles = DBtoNullableInt(Data.Item("SolutionMaxCycles"))
        Me.solutionSettings.SolutionPower = DBtoNullableDbl(Data.Item("SolutionPower"), 6)
        Me.solutionSettings.SolutionTolerance = DBtoNullableDbl(Data.Item("SolutionTolerance"), 6)
        Me.options.cantileverPoles.CantKFactor = DBtoNullableDbl(Data.Item("CantKFactor"), 6)
        Me.options.misclOptions.RadiusSampleDist = DBtoNullableDbl(Data.Item("RadiusSampleDist"), 6)
        Me.options.BypassStabilityChecks = DBtoNullableBool(Data.Item("BypassStabilityChecks"))
        Me.options.UseWindProjection = DBtoNullableBool(Data.Item("UseWindProjection"))
        Me.code.ice.UseIceEscalation = DBtoNullableBool(Data.Item("UseIceEscalation"))
        Me.options.UseDishCoeff = DBtoNullableBool(Data.Item("UseDishCoeff"))
        Me.options.AutoCalcTorqArmArea = DBtoNullableBool(Data.Item("AutoCalcTorqArmArea"))
        Me.options.windDirections.WindDirOption = DBtoNullableInt(Data.Item("WindDirOption"))
        Me.options.windDirections.WindDir0_0 = DBtoNullableBool(Data.Item("WindDir0_0"))
        Me.options.windDirections.WindDir0_1 = DBtoNullableBool(Data.Item("WindDir0_1"))
        Me.options.windDirections.WindDir0_2 = DBtoNullableBool(Data.Item("WindDir0_2"))
        Me.options.windDirections.WindDir0_3 = DBtoNullableBool(Data.Item("WindDir0_3"))
        Me.options.windDirections.WindDir0_4 = DBtoNullableBool(Data.Item("WindDir0_4"))
        Me.options.windDirections.WindDir0_5 = DBtoNullableBool(Data.Item("WindDir0_5"))
        Me.options.windDirections.WindDir0_6 = DBtoNullableBool(Data.Item("WindDir0_6"))
        Me.options.windDirections.WindDir0_7 = DBtoNullableBool(Data.Item("WindDir0_7"))
        Me.options.windDirections.WindDir0_8 = DBtoNullableBool(Data.Item("WindDir0_8"))
        Me.options.windDirections.WindDir0_9 = DBtoNullableBool(Data.Item("WindDir0_9"))
        Me.options.windDirections.WindDir0_10 = DBtoNullableBool(Data.Item("WindDir0_10"))
        Me.options.windDirections.WindDir0_11 = DBtoNullableBool(Data.Item("WindDir0_11"))
        Me.options.windDirections.WindDir0_12 = DBtoNullableBool(Data.Item("WindDir0_12"))
        Me.options.windDirections.WindDir0_13 = DBtoNullableBool(Data.Item("WindDir0_13"))
        Me.options.windDirections.WindDir0_14 = DBtoNullableBool(Data.Item("WindDir0_14"))
        Me.options.windDirections.WindDir0_15 = DBtoNullableBool(Data.Item("WindDir0_15"))
        Me.options.windDirections.WindDir1_0 = DBtoNullableBool(Data.Item("WindDir1_0"))
        Me.options.windDirections.WindDir1_1 = DBtoNullableBool(Data.Item("WindDir1_1"))
        Me.options.windDirections.WindDir1_2 = DBtoNullableBool(Data.Item("WindDir1_2"))
        Me.options.windDirections.WindDir1_3 = DBtoNullableBool(Data.Item("WindDir1_3"))
        Me.options.windDirections.WindDir1_4 = DBtoNullableBool(Data.Item("WindDir1_4"))
        Me.options.windDirections.WindDir1_5 = DBtoNullableBool(Data.Item("WindDir1_5"))
        Me.options.windDirections.WindDir1_6 = DBtoNullableBool(Data.Item("WindDir1_6"))
        Me.options.windDirections.WindDir1_7 = DBtoNullableBool(Data.Item("WindDir1_7"))
        Me.options.windDirections.WindDir1_8 = DBtoNullableBool(Data.Item("WindDir1_8"))
        Me.options.windDirections.WindDir1_9 = DBtoNullableBool(Data.Item("WindDir1_9"))
        Me.options.windDirections.WindDir1_10 = DBtoNullableBool(Data.Item("WindDir1_10"))
        Me.options.windDirections.WindDir1_11 = DBtoNullableBool(Data.Item("WindDir1_11"))
        Me.options.windDirections.WindDir1_12 = DBtoNullableBool(Data.Item("WindDir1_12"))
        Me.options.windDirections.WindDir1_13 = DBtoNullableBool(Data.Item("WindDir1_13"))
        Me.options.windDirections.WindDir1_14 = DBtoNullableBool(Data.Item("WindDir1_14"))
        Me.options.windDirections.WindDir1_15 = DBtoNullableBool(Data.Item("WindDir1_15"))
        Me.options.windDirections.WindDir2_0 = DBtoNullableBool(Data.Item("WindDir2_0"))
        Me.options.windDirections.WindDir2_1 = DBtoNullableBool(Data.Item("WindDir2_1"))
        Me.options.windDirections.WindDir2_2 = DBtoNullableBool(Data.Item("WindDir2_2"))
        Me.options.windDirections.WindDir2_3 = DBtoNullableBool(Data.Item("WindDir2_3"))
        Me.options.windDirections.WindDir2_4 = DBtoNullableBool(Data.Item("WindDir2_4"))
        Me.options.windDirections.WindDir2_5 = DBtoNullableBool(Data.Item("WindDir2_5"))
        Me.options.windDirections.WindDir2_6 = DBtoNullableBool(Data.Item("WindDir2_6"))
        Me.options.windDirections.WindDir2_7 = DBtoNullableBool(Data.Item("WindDir2_7"))
        Me.options.windDirections.WindDir2_8 = DBtoNullableBool(Data.Item("WindDir2_8"))
        Me.options.windDirections.WindDir2_9 = DBtoNullableBool(Data.Item("WindDir2_9"))
        Me.options.windDirections.WindDir2_10 = DBtoNullableBool(Data.Item("WindDir2_10"))
        Me.options.windDirections.WindDir2_11 = DBtoNullableBool(Data.Item("WindDir2_11"))
        Me.options.windDirections.WindDir2_12 = DBtoNullableBool(Data.Item("WindDir2_12"))
        Me.options.windDirections.WindDir2_13 = DBtoNullableBool(Data.Item("WindDir2_13"))
        Me.options.windDirections.WindDir2_14 = DBtoNullableBool(Data.Item("WindDir2_14"))
        Me.options.windDirections.WindDir2_15 = DBtoNullableBool(Data.Item("WindDir2_15"))
        Me.options.windDirections.SuppressWindPatternLoading = DBtoNullableBool(Data.Item("SuppressWindPatternLoading"))
    End Sub


    <Category("Constructor"), Description("Create TNX object from TNX file. Accepts EDSObject as parent.")>
    Public Sub New(ByVal tnxPath As String, Optional ByVal Parent As EDSObject = Nothing)
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        BuildFromFile(tnxPath)

    End Sub

    <Category("Constructor"), Description("Create TNX object from TNX file. Accepts EDSStructure as parent.")>
    Public Sub New(ByVal tnxPath As String, Optional ByVal Parent As EDSStructure = Nothing)
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        BuildFromFile(tnxPath)

    End Sub


    <Description("Build TNX object from TNX file.")>
    Public Sub BuildFromFile(ByVal tnxPath As String)
        Me.geometry.Absorb(Me)
        Me.filePath = tnxPath

        Dim tnxVar As String
        Dim tnxValue As String
        Dim recIndex As Integer
        Dim recordUSUnits As Boolean? = False
        Dim sectionFilter As String = ""
        Dim caseFilter As String = ""
        Dim dbFileFilter As String = ""

        For Each line In File.ReadLines(tnxPath)

            If Not line.Contains("=") Then
                tnxVar = line
                tnxValue = ""
            Else
                tnxVar = Left(line, line.IndexOf("="))
                tnxValue = Right(line, Len(line) - line.IndexOf("=") - 1)
            End If

            'Set caseFilter
            Select Case sectionFilter
                Case "db"
                    caseFilter = sectionFilter
                    'This if statement isn't needed because the [Structure] line can be used to deactivate the database filter
                    'If tnxVar = "File" Or tnxVar = "MemberMatFile" Or tnxVar = "USName" Or tnxVar = "SIName" Or tnxVar = "Values" Then
                    '    caseFilter = sectionFilter
                    'Else
                    '    caseFilter = ""
                    'End If
                Case "Antenna"
                    If Left(tnxVar, 7) = "Antenna" Then
                        caseFilter = sectionFilter
                    Else
                        caseFilter = ""
                    End If
                Case "Tower"
                    If Left(tnxVar, 5) = "Tower" Then
                        caseFilter = sectionFilter
                    Else
                        caseFilter = ""
                    End If
                Case "Guy"
                    If Left(tnxVar, 3) = "Guy" Or Left(tnxVar, 6) = "Anchor" Or Left(tnxVar, 7) = "Azimuth" Or Left(tnxVar, 6) = "Torque" Then
                        caseFilter = sectionFilter
                    Else
                        caseFilter = ""
                    End If
                Case "FeedLine"
                    If Left(tnxVar, 8) = "FeedLine" Or Left(tnxVar, 8) = "AutoCalc" Or Left(tnxVar, 7) = "Exclude" Or Left(tnxVar, 4) = "Flat" Then
                        caseFilter = sectionFilter
                    Else
                        caseFilter = ""
                    End If
                Case "Discrete"
                    'If Left(tnxVar, 9) = "TowerLoad" Or Left(tnxVar, 9) = "TowerVert" Or Left(tnxVar, 8) = "TowerOff" Or Left(tnxVar, 8) = "TowerLat" Or Left(tnxVar, 8) = "TowerAzi" Or Left(tnxVar, 8) = "TowerApp" Then
                    If Left(tnxVar, 5) = "Tower" Then
                        caseFilter = sectionFilter
                    Else
                        caseFilter = ""
                    End If
                Case "Dish"
                    If Left(tnxVar, 4) = "Dish" Then
                        caseFilter = sectionFilter
                    Else
                        caseFilter = ""
                    End If
                Case "UserForce"
                    If Left(tnxVar, 9) = "UserForce" Then
                        caseFilter = sectionFilter
                    Else
                        caseFilter = ""
                    End If
                Case Else
                    caseFilter = ""
            End Select

            Select Case True
                ''''Check for main file sections and activate corresponding filter''''
                Case tnxVar.Equals("[Databases]")
                    Try
                        sectionFilter = "db"
                        Me.otherLines.Add(New String() {tnxVar, tnxValue})
                    Catch ex As Exception
                        Debug.Print("Error parsing TNX variable: " & tnxVar)
                    End Try
                Case tnxVar.Equals("[Structure]")
                    Try
                        sectionFilter = "" 'Deactivate main section filter from database
                        Me.otherLines.Add(New String() {tnxVar, tnxValue})
                    Catch ex As Exception
                        Debug.Print("Error parsing TNX variable: " & tnxVar)
                    End Try
                Case tnxVar.Equals("NumAntennaRecs")
                    Try
                        sectionFilter = "Antenna"
                        Me.otherLines.Add(New String() {tnxVar, tnxValue})
                    Catch ex As Exception
                        Debug.Print("Error parsing TNX variable: " & tnxVar)
                    End Try
                Case tnxVar.Equals("NumTowerRecs")
                    Try
                        sectionFilter = "Tower"
                        Me.otherLines.Add(New String() {tnxVar, tnxValue})
                    Catch ex As Exception
                        Debug.Print("Error parsing TNX variable: " & tnxVar)
                    End Try
                Case tnxVar.Equals("NumGuyRecs")
                    Try
                        sectionFilter = "Guy"
                        Me.otherLines.Add(New String() {tnxVar, tnxValue})
                    Catch ex As Exception
                        Debug.Print("Error parsing TNX variable: " & tnxVar)
                    End Try
                Case tnxVar.Equals("NumFeedLineRecs")
                    Try
                        sectionFilter = "FeedLine"
                        Me.otherLines.Add(New String() {tnxVar, tnxValue})
                    Catch ex As Exception
                        Debug.Print("Error parsing TNX variable: " & tnxVar)
                    End Try
                Case tnxVar.Equals("NumTowerLoadRecs")
                    Try
                        sectionFilter = "Discrete"
                        Me.otherLines.Add(New String() {tnxVar, tnxValue})
                    Catch ex As Exception
                        Debug.Print("Error parsing TNX variable: " & tnxVar)
                    End Try
                Case tnxVar.Equals("NumDishRecs")
                    Try
                        sectionFilter = "Dish"
                        Me.otherLines.Add(New String() {tnxVar, tnxValue})
                    Catch ex As Exception
                        Debug.Print("Error parsing TNX variable: " & tnxVar)
                    End Try
                Case tnxVar.Equals("NumUserForceRecs")
                    Try
                        sectionFilter = "UserForce"
                        Me.otherLines.Add(New String() {tnxVar, tnxValue})
                    Catch ex As Exception
                        Debug.Print("Error parsing TNX variable: " & tnxVar)
                    End Try
                '''''If not a main section divider, use the caseFilter to minimize the number of cases checked per line.
                Case caseFilter = ""
                    ''''These are all the individual options for the eri file. They are not part of a record which there may be multiple of.'''
                    Select Case True
                    ''''Units''''
                        Case tnxVar.Equals("UnitsSystem")
                            If tnxValue <> "US" Then
                                Throw New System.Exception("TNX file is not in US units.")
                            End If
                            Me.settings.projectInfo.UnitsSystem = tnxValue
                        Case tnxVar.Equals("[US Units]")
                            recordUSUnits = True
                            Me.otherLines.Add(New String() {tnxVar})
                        Case tnxVar.Equals("[SI Units]")
                            recordUSUnits = False
                            Me.otherLines.Add(New String() {tnxVar})
                        Case tnxVar.Equals("Length")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Length = New tnxLengthUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("LengthPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Length.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Coordinate")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Coordinate = New tnxCoordinateUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("CoordinatePrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Coordinate.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Force")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Force = New tnxForceUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ForcePrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Force.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Load")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Load = New tnxLoadUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("LoadPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Load.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Moment")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Moment = New tnxMomentUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("MomentPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Moment.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Properties")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Properties = New tnxPropertiesUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("PropertiesPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Properties.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Pressure")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Pressure = New tnxPressureUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("PressurePrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Pressure.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Velocity")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Velocity = New tnxVelocityUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("VelocityPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Velocity.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Displacement")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Displacement = New tnxDisplacementUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DisplacementPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Displacement.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Mass")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Mass = New tnxMassUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("MassPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Mass.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Acceleration")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Acceleration = New tnxAccelerationUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AccelerationPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Acceleration.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Stress")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Stress = New tnxStressUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("StressPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Stress.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Density")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Density = New tnxDensityUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DensityPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Density.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UnitWt")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.UnitWt = New tnxUnitWTUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UnitWtPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.UnitWt.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Strength")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Strength = New tnxStrengthUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("StrengthPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Strength.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Modulus")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Modulus = New tnxModulusUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ModulusPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Modulus.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Temperature")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Temperature = New tnxTempUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TemperaturePrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Temperature.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Printer")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Printer = New tnxPrinterUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("PrinterPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Printer.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Rotation")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Rotation = New tnxRotationUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("RotationPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Rotation.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Spacing")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Spacing = New tnxSpacingUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("SpacingPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Spacing.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                ''''Project Info Settings
                        Case tnxVar.Equals("DesignStandardSeries")
                            Try
                                Me.settings.projectInfo.DesignStandardSeries = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UnitsSystem")
                            Try
                                Me.settings.projectInfo.UnitsSystem = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ClientName")
                            Try
                                Me.settings.projectInfo.ClientName = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ProjectName")
                            Try
                                Me.settings.projectInfo.ProjectName = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ProjectNumber")
                            Try
                                Me.settings.projectInfo.ProjectNumber = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("CreatedBy")
                            Try
                                Me.settings.projectInfo.CreatedBy = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("CreatedOn")
                            Try
                                Me.settings.projectInfo.CreatedOn = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("LastUsedBy")
                            Try
                                Me.settings.projectInfo.LastUsedBy = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("LastUsedOn")
                            Try
                                Me.settings.projectInfo.LastUsedOn = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("VersionUsed")
                            Try
                                Me.settings.projectInfo.VersionUsed = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                            '''User Info Settings
                        Case tnxVar.Equals("ViewerUserName")
                            Try
                                Me.settings.userInfo.ViewerUserName = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ViewerCompanyName")
                            Try
                                Me.settings.userInfo.ViewerCompanyName = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ViewerStreetAddress")
                            Try
                                Me.settings.userInfo.ViewerStreetAddress = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ViewerCityState")
                            Try
                                Me.settings.userInfo.ViewerCityState = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ViewerPhone")
                            Try
                                Me.settings.userInfo.ViewerPhone = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ViewerFAX")
                            Try
                                Me.settings.userInfo.ViewerFAX = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ViewerLogo")
                            Try
                                Me.settings.userInfo.ViewerLogo = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ViewerCompanyBitmap")
                            Try
                                Me.settings.userInfo.ViewerCompanyBitmap = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                ''''Code''''
                        Case tnxVar.Equals("DesignCode")
                            Try
                                Me.code.design.DesignCode = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ERIDesignMode")
                            Try
                                Me.code.design.ERIDesignMode = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DoInteraction")
                            Try
                                Me.code.design.DoInteraction = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DoHorzInteraction")
                            Try
                                Me.code.design.DoHorzInteraction = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DoDiagInteraction")
                            Try
                                Me.code.design.DoDiagInteraction = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UseMomentMagnification")
                            Try
                                Me.code.design.UseMomentMagnification = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UseCodeStressRatio")
                            Try
                                Me.code.design.UseCodeStressRatio = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AllowStressRatio")
                            Try
                                Me.code.design.AllowStressRatio = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AllowAntStressRatio")
                            Try
                                Me.code.design.AllowAntStressRatio = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UseCodeGuySF")
                            Try
                                Me.code.design.UseCodeGuySF = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuySF")
                            Try
                                Me.code.design.GuySF = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UseTIA222H_AnnexS")
                            Try
                                Me.code.design.UseTIA222H_AnnexS = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TIA_222_H_AnnexS_Ratio")
                            Try
                                Me.code.design.TIA_222_H_AnnexS_Ratio = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("PrintBitmaps")
                            Try
                                Me.code.design.PrintBitmaps = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("IceThickness")
                            Try
                                Me.code.ice.IceThickness = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("IceDensity")
                            Try
                                Me.code.ice.IceDensity = Me.settings.USUnits.Density.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UseModified_TIA_222_IceParameters")
                            Try
                                Me.code.ice.UseModified_TIA_222_IceParameters = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TIA_222_IceThicknessMultiplier")
                            Try
                                Me.code.ice.TIA_222_IceThicknessMultiplier = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DoNotUse_TIA_222_IceEscalation")
                            Try
                                Me.code.ice.DoNotUse_TIA_222_IceEscalation = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UseIceEscalation")
                            Try
                                Me.code.ice.UseIceEscalation = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TempDrop")
                            Try
                                Me.code.thermal.TempDrop = Me.settings.USUnits.Temperature.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GroutFc")
                            Try
                                Me.code.misclCode.GroutFc = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerBoltGrade")
                            Try
                                Me.code.misclCode.TowerBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerBoltMinEdgeDist")
                            Try
                                Me.code.misclCode.TowerBoltMinEdgeDist = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindSpeed")
                            Try
                                Me.code.wind.WindSpeed = Me.settings.USUnits.Velocity.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindSpeedIce")
                            Try
                                Me.code.wind.WindSpeedIce = Me.settings.USUnits.Velocity.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindSpeedService")
                            Try
                                Me.code.wind.WindSpeedService = Me.settings.USUnits.Velocity.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UseStateCountyLookup")
                            Try
                                Me.code.wind.UseStateCountyLookup = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("State")
                            Try
                                Me.code.wind.State = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("County")
                            Try
                                Me.code.wind.County = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UseMaxKz")
                            Try
                                Me.code.wind.UseMaxKz = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ASCE_7_10_WindData")
                            Try
                                Me.code.wind.ASCE_7_10_WindData = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ASCE_7_10_ConvertWindToASD")
                            Try
                                Me.code.wind.ASCE_7_10_ConvertWindToASD = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UseASCEWind")
                            Try
                                Me.code.wind.UseASCEWind = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AutoCalc_ASCE_GH")
                            Try
                                Me.code.wind.AutoCalc_ASCE_GH = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ASCE_ExposureCat")
                            Try
                                Me.code.wind.ASCE_ExposureCat = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ASCE_Year")
                            Try
                                Me.code.wind.ASCE_Year = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ASCEGh")
                            Try
                                Me.code.wind.ASCEGh = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ASCEI")
                            Try
                                Me.code.wind.ASCEI = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("CalcWindAt")
                            Try
                                Me.code.wind.CalcWindAt = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindCalcPoints")
                            Try
                                Me.code.wind.WindCalcPoints = Me.settings.USUnits.Length.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindExposure")
                            Try
                                Me.code.wind.WindExposure = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("StructureCategory")
                            Try
                                Me.code.wind.StructureCategory = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("RiskCategory")
                            Try
                                Me.code.wind.RiskCategory = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TopoCategory")
                            Try
                                Me.code.wind.TopoCategory = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("RSMTopographicFeature")
                            Try
                                Me.code.wind.RSMTopographicFeature = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("RSM_L")
                            Try
                                Me.code.wind.RSM_L = Me.settings.USUnits.Length.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("RSM_X")
                            Try
                                Me.code.wind.RSM_X = Me.settings.USUnits.Length.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("CrestHeight")
                            Try
                                Me.code.wind.CrestHeight = Me.settings.USUnits.Length.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TIA_222_H_TopoFeatureDownwind")
                            Try
                                Me.code.wind.TIA_222_H_TopoFeatureDownwind = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("BaseElevAboveSeaLevel")
                            Try
                                Me.code.wind.BaseElevAboveSeaLevel = Me.settings.USUnits.Length.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ConsiderRooftopSpeedUp")
                            Try
                                Me.code.wind.ConsiderRooftopSpeedUp = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("RooftopWS")
                            Try
                                Me.code.wind.RooftopWS = Me.settings.USUnits.Length.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("RooftopHS")
                            Try
                                Me.code.wind.RooftopHS = Me.settings.USUnits.Length.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("RooftopParapetHt")
                            Try
                                Me.code.wind.RooftopParapetHt = Me.settings.USUnits.Length.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("RooftopXB")
                            Try
                                Me.code.wind.RooftopXB = Me.settings.USUnits.Length.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindZone")
                            Try
                                Me.code.wind.WindZone = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("EIACWindMult")
                            Try
                                Me.code.wind.EIACWindMult = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("EIACWindMultIce")
                            Try
                                Me.code.wind.EIACWindMultIce = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("EIACIgnoreCableDrag")
                            Try
                                Me.code.wind.EIACIgnoreCableDrag = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("CSA_S37_RefVelPress")
                            Try
                                Me.code.wind.CSA_S37_RefVelPress = Me.settings.USUnits.Pressure.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("CSA_S37_ReliabilityClass")
                            Try
                                Me.code.wind.CSA_S37_ReliabilityClass = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("CSA_S37_ServiceabilityFactor")
                            Try
                                Me.code.wind.CSA_S37_ServiceabilityFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UseASCE7_10_Seismic_Lcomb")
                            Try
                                Me.code.seismic.UseASCE7_10_Seismic_Lcomb = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("SeismicSiteClass")
                            Try
                                Me.code.seismic.SeismicSiteClass = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("SeismicSs")
                            Try
                                Me.code.seismic.SeismicSs = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("SeismicS1")
                            Try
                                Me.code.seismic.SeismicS1 = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                ''''Options''''
                        Case tnxVar.Equals("UseClearSpans")
                            Try
                                Me.options.UseClearSpans = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UseClearSpansKlr")
                            Try
                                Me.options.UseClearSpansKlr = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UseFeedlineAsCylinder")
                            Try
                                Me.options.UseFeedlineAsCylinder = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UseLegLoads")
                            Try
                                Me.options.UseLegLoads = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("SRTakeCompression")
                            Try
                                Me.options.SRTakeCompression = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AllLegPanelsSame")
                            Try
                                Me.options.AllLegPanelsSame = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UseCombinedBoltCapacity")
                            Try
                                Me.options.UseCombinedBoltCapacity = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("SecHorzBracesLeg")
                            Try
                                Me.options.SecHorzBracesLeg = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("SortByComponent")
                            Try
                                Me.options.SortByComponent = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("SRCutEnds")
                            Try
                                Me.options.SRCutEnds = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("SRConcentric")
                            Try
                                Me.options.SRConcentric = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("CalcBlockShear")
                            Try
                                Me.options.CalcBlockShear = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Use4SidedDiamondBracing")
                            Try
                                Me.options.Use4SidedDiamondBracing = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TriangulateInnerBracing")
                            Try
                                Me.options.TriangulateInnerBracing = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("PrintCarrierNotes")
                            Try
                                Me.options.PrintCarrierNotes = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AddIBCWindCase")
                            Try
                                Me.options.AddIBCWindCase = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("LegBoltsAtTop")
                            Try
                                Me.options.LegBoltsAtTop = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UseTIA222Exemptions_MinBracingResistance")
                            Try
                                Me.options.UseTIA222Exemptions_MinBracingResistance = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UseTIA222Exemptions_TensionSplice")
                            Try
                                Me.options.UseTIA222Exemptions_TensionSplice = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("IgnoreKLryFor60DegAngleLegs")
                            Try
                                Me.options.IgnoreKLryFor60DegAngleLegs = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UseFeedlineTorque")
                            Try
                                Me.options.UseFeedlineTorque = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UsePinnedElements")
                            Try
                                Me.options.UsePinnedElements = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UseRigidIndex")
                            Try
                                Me.options.UseRigidIndex = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UseTrueCable")
                            Try
                                Me.options.UseTrueCable = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UseASCELy")
                            Try
                                Me.options.UseASCELy = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("CalcBracingForces")
                            Try
                                Me.options.CalcBracingForces = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("IgnoreBracingFEA")
                            Try
                                Me.options.IgnoreBracingFEA = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("BypassStabilityChecks")
                            Try
                                Me.options.BypassStabilityChecks = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UseWindProjection")
                            Try
                                Me.options.UseWindProjection = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UseDishCoeff")
                            Try
                                Me.options.UseDishCoeff = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AutoCalcTorqArmArea")
                            Try
                                Me.options.AutoCalcTorqArmArea = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("MastVert")
                            'The units of this input are dependant on the TNX force unit setting and the TNX property unit setting. (ie K/ft or Force/Properties) 
                            Try
                                'Me.options.foundationStiffness.MastVert = (CDbl(tnxValue)) * (Me.settings.USUnits.Properties.multiplier / Me.settings.USUnits.Force.multiplier)
                                Me.options.foundationStiffness.MastVert = Me.settings.USUnits.convertForcePerUnitLengthtoDefault(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("MastHorz")
                            'The units of this input are dependant on the TNX force unit setting and the TNX property unit setting. (ie K/ft or Force/Properties) 
                            Try
                                'Me.options.foundationStiffness.MastHorz = (CDbl(tnxValue)) * (Me.settings.USUnits.Properties.multiplier / Me.settings.USUnits.Force.multiplier)
                                Me.options.foundationStiffness.MastHorz = Me.settings.USUnits.convertForcePerUnitLengthtoDefault(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyVert")
                            'The units of this input are dependant on the TNX force unit setting and the TNX property unit setting. (ie K/ft or Force/Properties) 
                            Try
                                'Me.options.foundationStiffness.GuyVert = (CDbl(tnxValue)) * (Me.settings.USUnits.Properties.multiplier / Me.settings.USUnits.Force.multiplier)
                                Me.options.foundationStiffness.GuyVert = Me.settings.USUnits.convertForcePerUnitLengthtoDefault(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyHorz")
                            'The units of this input are dependant on the TNX force unit setting and the TNX property unit setting. (ie K/ft or Force/Properties) 
                            Try
                                'Me.options.foundationStiffness.GuyHorz = (CDbl(tnxValue)) * (Me.settings.USUnits.Properties.multiplier / Me.settings.USUnits.Force.multiplier)
                                Me.options.foundationStiffness.GuyHorz = Me.settings.USUnits.convertForcePerUnitLengthtoDefault(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GirtOffset")
                            Try
                                Me.options.defaultGirtOffsets.GirtOffset = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GirtOffsetLatticedPole")
                            Try
                                Me.options.defaultGirtOffsets.GirtOffsetLatticedPole = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("OffsetBotGirt")
                            Try
                                Me.options.defaultGirtOffsets.OffsetBotGirt = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("CheckVonMises")
                            Try
                                Me.options.cantileverPoles.CheckVonMises = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("SocketTopMount")
                            Try
                                Me.options.cantileverPoles.SocketTopMount = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("PrintMonopoleAtIncrements")
                            Try
                                Me.options.cantileverPoles.PrintMonopoleAtIncrements = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UseSubCriticalFlow")
                            Try
                                Me.options.cantileverPoles.UseSubCriticalFlow = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AssumePoleWithNoAttachments")
                            Try
                                Me.options.cantileverPoles.AssumePoleWithNoAttachments = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AssumePoleWithShroud")
                            Try
                                Me.options.cantileverPoles.AssumePoleWithShroud = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("PoleCornerRadiusKnown")
                            Try
                                Me.options.cantileverPoles.PoleCornerRadiusKnown = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("CantKFactor")
                            Try
                                Me.options.cantileverPoles.CantKFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("HogRodTakeup")
                            Try
                                Me.options.misclOptions.HogRodTakeup = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("RadiusSampleDist")
                            Try
                                Me.options.misclOptions.RadiusSampleDist = Me.settings.USUnits.Length.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDirOption")
                            Try
                                Me.options.windDirections.WindDirOption = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir0_0")
                            Try
                                Me.options.windDirections.WindDir0_0 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir0_1")
                            Try
                                Me.options.windDirections.WindDir0_1 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir0_2")
                            Try
                                Me.options.windDirections.WindDir0_2 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir0_3")
                            Try
                                Me.options.windDirections.WindDir0_3 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir0_4")
                            Try
                                Me.options.windDirections.WindDir0_4 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir0_5")
                            Try
                                Me.options.windDirections.WindDir0_5 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir0_6")
                            Try
                                Me.options.windDirections.WindDir0_6 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir0_7")
                            Try
                                Me.options.windDirections.WindDir0_7 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir0_8")
                            Try
                                Me.options.windDirections.WindDir0_8 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir0_9")
                            Try
                                Me.options.windDirections.WindDir0_9 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir0_10")
                            Try
                                Me.options.windDirections.WindDir0_10 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir0_11")
                            Try
                                Me.options.windDirections.WindDir0_11 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir0_12")
                            Try
                                Me.options.windDirections.WindDir0_12 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir0_13")
                            Try
                                Me.options.windDirections.WindDir0_13 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir0_14")
                            Try
                                Me.options.windDirections.WindDir0_14 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir0_15")
                            Try
                                Me.options.windDirections.WindDir0_15 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir1_0")
                            Try
                                Me.options.windDirections.WindDir1_0 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir1_1")
                            Try
                                Me.options.windDirections.WindDir1_1 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir1_2")
                            Try
                                Me.options.windDirections.WindDir1_2 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir1_3")
                            Try
                                Me.options.windDirections.WindDir1_3 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir1_4")
                            Try
                                Me.options.windDirections.WindDir1_4 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir1_5")
                            Try
                                Me.options.windDirections.WindDir1_5 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir1_6")
                            Try
                                Me.options.windDirections.WindDir1_6 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir1_7")
                            Try
                                Me.options.windDirections.WindDir1_7 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir1_8")
                            Try
                                Me.options.windDirections.WindDir1_8 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir1_9")
                            Try
                                Me.options.windDirections.WindDir1_9 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir1_10")
                            Try
                                Me.options.windDirections.WindDir1_10 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir1_11")
                            Try
                                Me.options.windDirections.WindDir1_11 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir1_12")
                            Try
                                Me.options.windDirections.WindDir1_12 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir1_13")
                            Try
                                Me.options.windDirections.WindDir1_13 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir1_14")
                            Try
                                Me.options.windDirections.WindDir1_14 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir1_15")
                            Try
                                Me.options.windDirections.WindDir1_15 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir2_0")
                            Try
                                Me.options.windDirections.WindDir2_0 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir2_1")
                            Try
                                Me.options.windDirections.WindDir2_1 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir2_2")
                            Try
                                Me.options.windDirections.WindDir2_2 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir2_3")
                            Try
                                Me.options.windDirections.WindDir2_3 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir2_4")
                            Try
                                Me.options.windDirections.WindDir2_4 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir2_5")
                            Try
                                Me.options.windDirections.WindDir2_5 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir2_6")
                            Try
                                Me.options.windDirections.WindDir2_6 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir2_7")
                            Try
                                Me.options.windDirections.WindDir2_7 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir2_8")
                            Try
                                Me.options.windDirections.WindDir2_8 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir2_9")
                            Try
                                Me.options.windDirections.WindDir2_9 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir2_10")
                            Try
                                Me.options.windDirections.WindDir2_10 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir2_11")
                            Try
                                Me.options.windDirections.WindDir2_11 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir2_12")
                            Try
                                Me.options.windDirections.WindDir2_12 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir2_13")
                            Try
                                Me.options.windDirections.WindDir2_13 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir2_14")
                            Try
                                Me.options.windDirections.WindDir2_14 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("WindDir2_15")
                            Try
                                Me.options.windDirections.WindDir2_15 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("SuppressWindPatternLoading")
                            Try
                                Me.options.windDirections.SuppressWindPatternLoading = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try

                ''''General Geometry''''
                        Case tnxVar.Equals("TowerType")
                            Try
                                Me.geometry.TowerType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaType")
                            Try
                                Me.geometry.AntennaType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("OverallHeight")
                            Try
                                Me.geometry.OverallHeight = Me.settings.USUnits.Length.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("BaseElevation")
                            Try
                                Me.geometry.BaseElevation = Me.settings.USUnits.Length.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Lambda")
                            Try
                                Me.geometry.Lambda = Me.settings.USUnits.Spacing.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerTopFaceWidth")
                            Try
                                Me.geometry.TowerTopFaceWidth = Me.settings.USUnits.Spacing.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerBaseFaceWidth")
                            Try
                                Me.geometry.TowerBaseFaceWidth = Me.settings.USUnits.Spacing.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerTaper")
                            Try
                                Me.geometry.TowerTaper = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyedMonopoleBaseType")
                            Try
                                Me.geometry.GuyedMonopoleBaseType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TaperHeight")
                            Try
                                Me.geometry.TaperHeight = Me.settings.USUnits.Length.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("PivotHeight")
                            Try
                                Me.geometry.PivotHeight = Me.settings.USUnits.Length.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AutoCalcGH")
                            Try
                                Me.geometry.AutoCalcGH = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UserGHElev")
                            Try
                                Me.geometry.UserGHElev = Me.settings.USUnits.Length.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UseIndexPlate")
                            Try
                                Me.geometry.UseIndexPlate = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("EnterUserDefinedGhValues")
                            Try
                                Me.geometry.EnterUserDefinedGhValues = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("BaseTowerGhInput")
                            Try
                                Me.geometry.BaseTowerGhInput = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UpperStructureGhInput")
                            Try
                                Me.geometry.UpperStructureGhInput = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("EnterUserDefinedCgValues")
                            Try
                                Me.geometry.EnterUserDefinedCgValues = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("BaseTowerCgInput")
                            Try
                                Me.geometry.BaseTowerCgInput = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UpperStructureCgInput")
                            Try
                                Me.geometry.UpperStructureCgInput = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaFaceWidth")
                            Try
                                Me.geometry.AntennaFaceWidth = Me.settings.USUnits.Spacing.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UseTopTakeup")
                            Try
                                Me.geometry.UseTopTakeup = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ConstantSlope")
                            Try
                                Me.geometry.ConstantSlope = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("[End Application]")
                            Me.otherLines.Add(New String() {tnxVar})
                            Exit For
                    ''''Solution Options''''
                        Case tnxVar.Equals("SolutionUsePDelta")
                            Try
                                Me.solutionSettings.SolutionUsePDelta = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("SolutionMinStiffness")
                            Try
                                Me.solutionSettings.SolutionMinStiffness = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("SolutionMaxStiffness")
                            Try
                                Me.solutionSettings.SolutionMaxStiffness = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("SolutionMaxCycles")
                            Try
                                Me.solutionSettings.SolutionMaxCycles = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("SolutionPower")
                            Try
                                Me.solutionSettings.SolutionPower = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("SolutionTolerance")
                            Try
                                Me.solutionSettings.SolutionTolerance = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        '''''MTO Settings''''
                        Case tnxVar.Equals("IncludeCapacityNote")
                            Try
                                Me.MTOSettings.IncludeCapacityNote = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("IncludeAppurtGraphics")
                            Try
                                Me.MTOSettings.IncludeAppurtGraphics = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DisplayNotes")
                            Try
                                Me.MTOSettings.DisplayNotes = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DisplayReactions")
                            Try
                                Me.MTOSettings.DisplayReactions = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DisplaySchedule")
                            Try
                                Me.MTOSettings.DisplaySchedule = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DisplayAppurtenanceTable")
                            Try
                                Me.MTOSettings.DisplayAppurtenanceTable = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DisplayMaterialStrengthTable")
                            Try
                                Me.MTOSettings.DisplayMaterialStrengthTable = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Notes")
                            Try
                                'Me.MTOSettings.Notes.Add(New tnxNote With {.Note = tnxValue})
                                If Not Me.MTOSettings.Notes = "" Then Me.MTOSettings.Notes += "||"
                                Me.MTOSettings.Notes += tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                    ''''Report Settings''''
                        Case tnxVar.Equals("ReportInputCosts")
                            Try
                                Me.reportSettings.ReportInputCosts = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ReportInputGeometry")
                            Try
                                Me.reportSettings.ReportInputGeometry = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ReportInputOptions")
                            Try
                                Me.reportSettings.ReportInputOptions = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ReportMaxForces")
                            Try
                                Me.reportSettings.ReportMaxForces = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ReportInputMap")
                            Try
                                Me.reportSettings.ReportInputMap = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("CostReportOutputType")
                            Try
                                Me.reportSettings.CostReportOutputType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("CapacityReportOutputType")
                            Try
                                Me.reportSettings.CapacityReportOutputType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ReportPrintForceTotals")
                            Try
                                Me.reportSettings.ReportPrintForceTotals = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ReportPrintForceDetails")
                            Try
                                Me.reportSettings.ReportPrintForceDetails = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ReportPrintMastVectors")
                            Try
                                Me.reportSettings.ReportPrintMastVectors = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ReportPrintAntPoleVectors")
                            Try
                                Me.reportSettings.ReportPrintAntPoleVectors = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ReportPrintDiscreteVectors")
                            Try
                                Me.reportSettings.ReportPrintDiscreteVectors = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ReportPrintDishVectors")
                            Try
                                Me.reportSettings.ReportPrintDishVectors = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ReportPrintFeedTowerVectors")
                            Try
                                Me.reportSettings.ReportPrintFeedTowerVectors = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ReportPrintUserLoadVectors")
                            Try
                                Me.reportSettings.ReportPrintUserLoadVectors = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ReportPrintPressures")
                            Try
                                Me.reportSettings.ReportPrintPressures = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ReportPrintAppurtForces")
                            Try
                                Me.reportSettings.ReportPrintAppurtForces = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ReportPrintGuyForces")
                            Try
                                Me.reportSettings.ReportPrintGuyForces = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ReportPrintGuyStressing")
                            Try
                                Me.reportSettings.ReportPrintGuyStressing = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ReportPrintDeflections")
                            Try
                                Me.reportSettings.ReportPrintDeflections = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ReportPrintReactions")
                            Try
                                Me.reportSettings.ReportPrintReactions = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ReportPrintStressChecks")
                            Try
                                Me.reportSettings.ReportPrintStressChecks = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ReportPrintBoltChecks")
                            Try
                                Me.reportSettings.ReportPrintBoltChecks = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ReportPrintInputGVerificationTables")
                            Try
                                Me.reportSettings.ReportPrintInputGVerificationTables = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ReportPrintOutputGVerificationTables")
                            Try
                                Me.reportSettings.ReportPrintOutputGVerificationTables = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                    ''''CCI Report''''
                        Case tnxVar.Equals("sReportProjectNumber")
                            Try
                                Me.CCIReport.sReportProjectNumber = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportJobType")
                            Try
                                Me.CCIReport.sReportJobType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportCarrierName")
                            Try
                                Me.CCIReport.sReportCarrierName = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportCarrierSiteNumber")
                            Try
                                Me.CCIReport.sReportCarrierSiteNumber = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportCarrierSiteName")
                            Try
                                Me.CCIReport.sReportCarrierSiteName = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportSiteAddress")
                            Try
                                Me.CCIReport.sReportSiteAddress = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportLatitudeDegree")
                            Try
                                Me.CCIReport.sReportLatitudeDegree = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportLatitudeMinute")
                            Try
                                Me.CCIReport.sReportLatitudeMinute = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportLatitudeSecond")
                            Try
                                Me.CCIReport.sReportLatitudeSecond = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportLongitudeDegree")
                            Try
                                Me.CCIReport.sReportLongitudeDegree = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportLongitudeMinute")
                            Try
                                Me.CCIReport.sReportLongitudeMinute = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportLongitudeSecond")
                            Try
                                Me.CCIReport.sReportLongitudeSecond = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportLocalCodeRequirement")
                            Try
                                Me.CCIReport.sReportLocalCodeRequirement = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportSiteHistory")
                            Try
                                Me.CCIReport.sReportSiteHistory = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportTowerManufacturer")
                            Try
                                Me.CCIReport.sReportTowerManufacturer = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportMonthManufactured")
                            Try
                                Me.CCIReport.sReportMonthManufactured = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportYearManufactured")
                            Try
                                Me.CCIReport.sReportYearManufactured = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportOriginalSpeed")
                            Try
                                Me.CCIReport.sReportOriginalSpeed = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportOriginalCode")
                            Try
                                Me.CCIReport.sReportOriginalCode = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportTowerType")
                            Try
                                Me.CCIReport.sReportTowerType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportEngrName")
                            Try
                                Me.CCIReport.sReportEngrName = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportEngrTitle")
                            Try
                                Me.CCIReport.sReportEngrTitle = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportHQPhoneNumber")
                            Try
                                Me.CCIReport.sReportHQPhoneNumber = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportEmailAddress")
                            Try
                                Me.CCIReport.sReportEmailAddress = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportLogoPath")
                            Try
                                Me.CCIReport.sReportLogoPath = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportCCiContactName")
                            Try
                                Me.CCIReport.sReportCCiContactName = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportCCiAddress1")
                            Try
                                Me.CCIReport.sReportCCiAddress1 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportCCiAddress2")
                            Try
                                Me.CCIReport.sReportCCiAddress2 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportCCiBUNumber")
                            Try
                                Me.CCIReport.sReportCCiBUNumber = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportCCiSiteName")
                            Try
                                Me.CCIReport.sReportCCiSiteName = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportCCiJDENumber")
                            Try
                                Me.CCIReport.sReportCCiJDENumber = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportCCiWONumber")
                            Try
                                Me.CCIReport.sReportCCiWONumber = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportCCiPONumber")
                            Try
                                Me.CCIReport.sReportCCiPONumber = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportCCiAppNumber")
                            Try
                                Me.CCIReport.sReportCCiAppNumber = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportCCiRevNumber")
                            Try
                                Me.CCIReport.sReportCCiRevNumber = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportDocsProvided")
                            Try
                                Me.CCIReport.sReportDocsProvided.Add(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportRecommendations")
                            Try
                                Me.CCIReport.sReportRecommendations = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportAppurt1")
                            Try
                                Me.CCIReport.sReportAppurt1.Add(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportAppurt2")
                            Try
                                Me.CCIReport.sReportAppurt2.Add(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportAppurt3")
                            Try
                                Me.CCIReport.sReportAppurt3.Add(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportAddlCapacity")
                            Try
                                Me.CCIReport.sReportAddlCapacity.Add(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportAssumption")
                            Try
                                Me.CCIReport.sReportAssumption.Add(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportAppurt1Note1")
                            Try
                                Me.CCIReport.sReportAppurt1Note1 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportAppurt1Note2")
                            Try
                                Me.CCIReport.sReportAppurt1Note2 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportAppurt1Note3")
                            Try
                                Me.CCIReport.sReportAppurt1Note3 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportAppurt1Note4")
                            Try
                                Me.CCIReport.sReportAppurt1Note4 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportAppurt1Note5")
                            Try
                                Me.CCIReport.sReportAppurt1Note5 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportAppurt1Note6")
                            Try
                                Me.CCIReport.sReportAppurt1Note6 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportAppurt1Note7")
                            Try
                                Me.CCIReport.sReportAppurt1Note7 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportAppurt2Note1")
                            Try
                                Me.CCIReport.sReportAppurt2Note1 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportAppurt2Note2")
                            Try
                                Me.CCIReport.sReportAppurt2Note2 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportAppurt2Note3")
                            Try
                                Me.CCIReport.sReportAppurt2Note3 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportAppurt2Note4")
                            Try
                                Me.CCIReport.sReportAppurt2Note4 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportAppurt2Note5")
                            Try
                                Me.CCIReport.sReportAppurt2Note5 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportAppurt2Note6")
                            Try
                                Me.CCIReport.sReportAppurt2Note6 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportAppurt2Note7")
                            Try
                                Me.CCIReport.sReportAppurt2Note7 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportAddlCapacityNote1")
                            Try
                                Me.CCIReport.sReportAddlCapacityNote1 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportAddlCapacityNote2")
                            Try
                                Me.CCIReport.sReportAddlCapacityNote2 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportAddlCapacityNote3")
                            Try
                                Me.CCIReport.sReportAddlCapacityNote3 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("sReportAddlCapacityNote4")
                            Try
                                Me.CCIReport.sReportAddlCapacityNote4 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case Else ''''All other lines
                            If line.Contains("=") Then
                                Me.otherLines.Add(New String() {tnxVar, tnxValue})
                            Else
                                Me.otherLines.Add(New String() {tnxVar})
                            End If
                    End Select
                Case caseFilter = "db"
                    Select Case True
                        Case tnxVar.Equals("File")
                            dbFileFilter = "member"
                            Me.database.members.Add(New tnxMember(Me.database) With {.File = tnxValue})
                        Case tnxVar.Equals("MemberMatFile")
                            dbFileFilter = "material"
                            Me.database.materials.Add(New tnxMaterial(Me.database) With {.MemberMatFile = tnxValue, .IsBolt = False})
                        Case tnxVar.Equals("BoltMatFile")
                            dbFileFilter = "bolt"
                            Me.database.bolts.Add(New tnxMaterial(Me.database) With {.MemberMatFile = tnxValue, .IsBolt = True})
                        Case tnxVar.Equals("USName")
                            If dbFileFilter = "member" Then Me.database.members.Last.USName = tnxValue
                        Case tnxVar.Equals("SIName")
                            If dbFileFilter = "member" Then Me.database.members.Last.SIName = tnxValue
                        Case tnxVar.Equals("Values")
                            If dbFileFilter = "member" Then Me.database.members.Last.Values = tnxValue
                        Case tnxVar.Equals("MatName")
                            Select Case dbFileFilter
                                Case "material"
                                    Me.database.materials.Last.MatName = tnxValue
                                Case "bolt"
                                    Me.database.bolts.Last.MatName = tnxValue
                            End Select
                        Case tnxVar.Equals("MatValues")
                            Select Case dbFileFilter
                                Case "material"
                                    Me.database.materials.Last.MatValues = tnxValue
                                Case "bolt"
                                    Me.database.bolts.Last.MatValues = tnxValue
                            End Select
                    End Select
                Case caseFilter = "Antenna"
                    ''''Antenna Rec (Upper Structure)''''
                    Select Case True
                        Case tnxVar.Equals("AntennaRec")
                            Try
                                recIndex = CInt(tnxValue) - 1
                                Me.geometry.upperStructure.Add(New tnxAntennaRecord(Me.geometry))
                                Me.geometry.upperStructure(recIndex).Rec = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaBraceType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaBraceType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaHeight")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaHeight = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaDiagonalSpacing")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagonalSpacing = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaDiagonalSpacingEx")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagonalSpacingEx = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaNumSections")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaNumSections = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaNumSesctions")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaNumSesctions = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaSectionLength")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaSectionLength = Me.settings.USUnits.Length.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaLegType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaLegType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaLegSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaLegSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaLegGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaLegGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaLegMatlGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaLegMatlGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaDiagonalGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagonalGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaDiagonalMatlGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagonalMatlGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaInnerBracingGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerBracingGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaInnerBracingMatlGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerBracingMatlGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaTopGirtGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTopGirtGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaTopGirtMatlGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTopGirtMatlGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaBotGirtGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaBotGirtGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaBotGirtMatlGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaBotGirtMatlGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaInnerGirtGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerGirtGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaInnerGirtMatlGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerGirtMatlGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaLongHorizontalGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaLongHorizontalGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaLongHorizontalMatlGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaLongHorizontalMatlGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaShortHorizontalGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaShortHorizontalGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaShortHorizontalMatlGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaShortHorizontalMatlGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaDiagonalType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagonalType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaDiagonalSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagonalSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaInnerBracingType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerBracingType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaInnerBracingSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerBracingSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaTopGirtType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTopGirtType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaTopGirtSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTopGirtSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaBotGirtType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaBotGirtType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaBotGirtSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaBotGirtSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaTopGirtOffset")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTopGirtOffset = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaBotGirtOffset")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaBotGirtOffset = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaHasKBraceEndPanels")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaHasKBraceEndPanels = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaHasHorizontals")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaHasHorizontals = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaLongHorizontalType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaLongHorizontalType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaLongHorizontalSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaLongHorizontalSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaShortHorizontalType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaShortHorizontalType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaShortHorizontalSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaShortHorizontalSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantMatlGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantMatlGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantDiagType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantDiagType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubDiagonalType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubDiagonalType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubHorizontalType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubHorizontalType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantVerticalType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantVerticalType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipDiagonalType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipDiagonalType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHorizontalSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHorizontalSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHorizontalSize2")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHorizontalSize2 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHorizontalSize3")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHorizontalSize3 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHorizontalSize4")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHorizontalSize4 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantDiagonalSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantDiagonalSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantDiagonalSize2")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantDiagonalSize2 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantDiagonalSize3")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantDiagonalSize3 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantDiagonalSize4")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantDiagonalSize4 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubHorizontalSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubHorizontalSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubDiagonalSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubDiagonalSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaSubDiagLocation")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaSubDiagLocation = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantVerticalSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantVerticalSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipDiagonalSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipDiagonalSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipDiagonalSize2")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipDiagonalSize2 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipDiagonalSize3")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipDiagonalSize3 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipDiagonalSize4")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipDiagonalSize4 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipSize2")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipSize2 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipSize3")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipSize3 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipSize4")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipSize4 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaNumInnerGirts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaNumInnerGirts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaInnerGirtType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerGirtType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaInnerGirtSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerGirtSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaPoleShapeType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaPoleShapeType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaPoleSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaPoleSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaPoleGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaPoleGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaPoleMatlGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaPoleMatlGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaPoleSpliceLength")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaPoleSpliceLength = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaTaperPoleNumSides")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTaperPoleNumSides = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaTaperPoleTopDiameter")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTaperPoleTopDiameter = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaTaperPoleBotDiameter")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTaperPoleBotDiameter = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaTaperPoleWallThickness")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTaperPoleWallThickness = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaTaperPoleBendRadius")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTaperPoleBendRadius = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaTaperPoleGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTaperPoleGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaTaperPoleMatlGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTaperPoleMatlGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaSWMult")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaSWMult = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaWPMult")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaWPMult = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaAutoCalcKSingleAngle")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaAutoCalcKSingleAngle = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaAutoCalcKSolidRound")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaAutoCalcKSolidRound = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaAfGusset")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaAfGusset = Me.settings.USUnits.Length.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaTfGusset")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTfGusset = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaGussetBoltEdgeDistance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaGussetBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaGussetGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaGussetGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaGussetMatlGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaGussetMatlGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaAfMult")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaAfMult = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaArMult")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaArMult = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaFlatIPAPole")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaFlatIPAPole = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRoundIPAPole")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRoundIPAPole = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaFlatIPALeg")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaFlatIPALeg = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRoundIPALeg")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRoundIPALeg = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaFlatIPAHorizontal")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaFlatIPAHorizontal = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRoundIPAHorizontal")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRoundIPAHorizontal = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaFlatIPADiagonal")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaFlatIPADiagonal = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRoundIPADiagonal")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRoundIPADiagonal = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaCSA_S37_SpeedUpFactor")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaCSA_S37_SpeedUpFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaKLegs")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKLegs = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaKXBracedDiags")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKXBracedDiags = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaKKBracedDiags")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKKBracedDiags = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaKZBracedDiags")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKZBracedDiags = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaKHorzs")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKHorzs = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaKSecHorzs")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKSecHorzs = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaKGirts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKGirts = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaKInners")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKInners = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaKXBracedDiagsY")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKXBracedDiagsY = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaKKBracedDiagsY")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKKBracedDiagsY = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaKZBracedDiagsY")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKZBracedDiagsY = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaKHorzsY")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKHorzsY = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaKSecHorzsY")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKSecHorzsY = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaKGirtsY")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKGirtsY = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaKInnersY")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKInnersY = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaKRedHorz")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKRedHorz = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaKRedDiag")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKRedDiag = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaKRedSubDiag")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKRedSubDiag = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaKRedSubHorz")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKRedSubHorz = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaKRedVert")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKRedVert = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaKRedHip")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKRedHip = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaKRedHipDiag")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKRedHipDiag = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaKTLX")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKTLX = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaKTLZ")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKTLZ = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaKTLLeg")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKTLLeg = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaInnerKTLX")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerKTLX = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaInnerKTLZ")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerKTLZ = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaInnerKTLLeg")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerKTLLeg = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaStitchBoltLocationHoriz")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaStitchBoltLocationHoriz = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaStitchBoltLocationDiag")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaStitchBoltLocationDiag = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaStitchSpacing")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaStitchSpacing = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaStitchSpacingHorz")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaStitchSpacingHorz = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaStitchSpacingDiag")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaStitchSpacingDiag = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaStitchSpacingRed")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaStitchSpacingRed = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaLegNetWidthDeduct")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaLegNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaLegUFactor")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaLegUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaDiagonalNetWidthDeduct")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagonalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaTopGirtNetWidthDeduct")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTopGirtNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaBotGirtNetWidthDeduct")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaBotGirtNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaInnerGirtNetWidthDeduct")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerGirtNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaHorizontalNetWidthDeduct")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaHorizontalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaShortHorizontalNetWidthDeduct")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaShortHorizontalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaDiagonalUFactor")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagonalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaTopGirtUFactor")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTopGirtUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaBotGirtUFactor")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaBotGirtUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaInnerGirtUFactor")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerGirtUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaHorizontalUFactor")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaHorizontalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaShortHorizontalUFactor")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaShortHorizontalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaLegConnType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaLegConnType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaLegNumBolts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaLegNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaDiagonalNumBolts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagonalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaTopGirtNumBolts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTopGirtNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaBotGirtNumBolts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaBotGirtNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaInnerGirtNumBolts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerGirtNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaHorizontalNumBolts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaHorizontalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaShortHorizontalNumBolts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaShortHorizontalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaLegBoltGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaLegBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaLegBoltSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaLegBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaDiagonalBoltGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagonalBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaDiagonalBoltSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagonalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaTopGirtBoltGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTopGirtBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaTopGirtBoltSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTopGirtBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaBotGirtBoltGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaBotGirtBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaBotGirtBoltSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaBotGirtBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaInnerGirtBoltGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerGirtBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaInnerGirtBoltSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerGirtBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaHorizontalBoltGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaHorizontalBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaHorizontalBoltSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaHorizontalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaShortHorizontalBoltGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaShortHorizontalBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaShortHorizontalBoltSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaShortHorizontalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaLegBoltEdgeDistance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaLegBoltEdgeDistance = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaDiagonalBoltEdgeDistance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagonalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaTopGirtBoltEdgeDistance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTopGirtBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaBotGirtBoltEdgeDistance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaBotGirtBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaInnerGirtBoltEdgeDistance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerGirtBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaHorizontalBoltEdgeDistance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaHorizontalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaShortHorizontalBoltEdgeDistance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaShortHorizontalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaDiagonalGageG1Distance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagonalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaTopGirtGageG1Distance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTopGirtGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaBotGirtGageG1Distance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaBotGirtGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaInnerGirtGageG1Distance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerGirtGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaHorizontalGageG1Distance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaHorizontalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaShortHorizontalGageG1Distance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaShortHorizontalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHorizontalBoltGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHorizontalBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHorizontalBoltSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHorizontalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHorizontalNumBolts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHorizontalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHorizontalBoltEdgeDistance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHorizontalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHorizontalGageG1Distance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHorizontalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHorizontalNetWidthDeduct")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHorizontalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHorizontalUFactor")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHorizontalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantDiagonalBoltGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantDiagonalBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantDiagonalBoltSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantDiagonalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantDiagonalNumBolts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantDiagonalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantDiagonalBoltEdgeDistance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantDiagonalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantDiagonalGageG1Distance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantDiagonalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantDiagonalNetWidthDeduct")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantDiagonalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantDiagonalUFactor")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantDiagonalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubDiagonalBoltGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubDiagonalBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubDiagonalBoltSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubDiagonalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubDiagonalNumBolts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubDiagonalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubDiagonalBoltEdgeDistance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubDiagonalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubDiagonalGageG1Distance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubDiagonalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubDiagonalNetWidthDeduct")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubDiagonalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubDiagonalUFactor")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubDiagonalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubHorizontalBoltGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubHorizontalBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubHorizontalBoltSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubHorizontalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubHorizontalNumBolts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubHorizontalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubHorizontalBoltEdgeDistance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubHorizontalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubHorizontalGageG1Distance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubHorizontalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubHorizontalNetWidthDeduct")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubHorizontalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubHorizontalUFactor")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubHorizontalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantVerticalBoltGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantVerticalBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantVerticalBoltSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantVerticalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantVerticalNumBolts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantVerticalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantVerticalBoltEdgeDistance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantVerticalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantVerticalGageG1Distance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantVerticalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantVerticalNetWidthDeduct")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantVerticalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantVerticalUFactor")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantVerticalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipBoltGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipBoltSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipNumBolts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipBoltEdgeDistance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipGageG1Distance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipNetWidthDeduct")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipUFactor")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipDiagonalBoltGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipDiagonalBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipDiagonalBoltSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipDiagonalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipDiagonalNumBolts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipDiagonalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipDiagonalBoltEdgeDistance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipDiagonalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipDiagonalGageG1Distance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipDiagonalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipDiagonalNetWidthDeduct")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipDiagonalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipDiagonalUFactor")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipDiagonalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaDiagonalOutOfPlaneRestraint")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagonalOutOfPlaneRestraint = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaTopGirtOutOfPlaneRestraint")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTopGirtOutOfPlaneRestraint = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaBottomGirtOutOfPlaneRestraint")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaBottomGirtOutOfPlaneRestraint = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaMidGirtOutOfPlaneRestraint")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaMidGirtOutOfPlaneRestraint = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaHorizontalOutOfPlaneRestraint")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaHorizontalOutOfPlaneRestraint = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaSecondaryHorizontalOutOfPlaneRestraint")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaSecondaryHorizontalOutOfPlaneRestraint = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaDiagOffsetNEY")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagOffsetNEY = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaDiagOffsetNEX")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagOffsetNEX = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaDiagOffsetPEY")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagOffsetPEY = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaDiagOffsetPEX")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagOffsetPEX = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaKbraceOffsetNEY")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKbraceOffsetNEY = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaKbraceOffsetNEX")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKbraceOffsetNEX = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaKbraceOffsetPEY")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKbraceOffsetPEY = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AntennaKbraceOffsetPEX")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKbraceOffsetPEX = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                    End Select
                Case caseFilter = "Tower"
                    ''''Tower Rec (Base Structure)''''
                    Select Case True
                        Case tnxVar.Equals("TowerRec")
                            Try
                                recIndex = CInt(tnxValue) - 1
                                Me.geometry.baseStructure.Add(New tnxTowerRecord(Me.geometry))
                                Me.geometry.baseStructure(recIndex).Rec = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerDatabase")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDatabase = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerName")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerName = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerHeight")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerHeight = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerFaceWidth")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerFaceWidth = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerNumSections")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerNumSections = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerSectionLength")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerSectionLength = Me.settings.USUnits.Length.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerDiagonalSpacing")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagonalSpacing = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerDiagonalSpacingEx")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagonalSpacingEx = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerBraceType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerBraceType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerFaceBevel")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerFaceBevel = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerTopGirtOffset")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerTopGirtOffset = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerBotGirtOffset")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerBotGirtOffset = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerHasKBraceEndPanels")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerHasKBraceEndPanels = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerHasHorizontals")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerHasHorizontals = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLegType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerLegType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLegSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerLegSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLegGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerLegGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLegMatlGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerLegMatlGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerDiagonalGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagonalGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerDiagonalMatlGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagonalMatlGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerInnerBracingGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerBracingGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerInnerBracingMatlGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerBracingMatlGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerTopGirtGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerTopGirtGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerTopGirtMatlGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerTopGirtMatlGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerBotGirtGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerBotGirtGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerBotGirtMatlGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerBotGirtMatlGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerInnerGirtGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerGirtGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerInnerGirtMatlGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerGirtMatlGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLongHorizontalGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerLongHorizontalGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLongHorizontalMatlGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerLongHorizontalMatlGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerShortHorizontalGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerShortHorizontalGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerShortHorizontalMatlGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerShortHorizontalMatlGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerDiagonalType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagonalType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerDiagonalSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagonalSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerInnerBracingType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerBracingType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerInnerBracingSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerBracingSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerTopGirtType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerTopGirtType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerTopGirtSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerTopGirtSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerBotGirtType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerBotGirtType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerBotGirtSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerBotGirtSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerNumInnerGirts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerNumInnerGirts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerInnerGirtType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerGirtType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerInnerGirtSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerGirtSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLongHorizontalType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerLongHorizontalType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLongHorizontalSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerLongHorizontalSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerShortHorizontalType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerShortHorizontalType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerShortHorizontalSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerShortHorizontalSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantMatlGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantMatlGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantDiagType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantDiagType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubDiagonalType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubDiagonalType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubHorizontalType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubHorizontalType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantVerticalType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantVerticalType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipDiagonalType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipDiagonalType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHorizontalSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHorizontalSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHorizontalSize2")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHorizontalSize2 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHorizontalSize3")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHorizontalSize3 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHorizontalSize4")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHorizontalSize4 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantDiagonalSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantDiagonalSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantDiagonalSize2")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantDiagonalSize2 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantDiagonalSize3")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantDiagonalSize3 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantDiagonalSize4")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantDiagonalSize4 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubHorizontalSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubHorizontalSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubDiagonalSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubDiagonalSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerSubDiagLocation")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerSubDiagLocation = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantVerticalSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantVerticalSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipSize2")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipSize2 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipSize3")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipSize3 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipSize4")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipSize4 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipDiagonalSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipDiagonalSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipDiagonalSize2")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipDiagonalSize2 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipDiagonalSize3")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipDiagonalSize3 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipDiagonalSize4")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipDiagonalSize4 = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerSWMult")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerSWMult = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerWPMult")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerWPMult = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerAutoCalcKSingleAngle")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerAutoCalcKSingleAngle = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerAutoCalcKSolidRound")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerAutoCalcKSolidRound = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerAfGusset")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerAfGusset = Me.settings.USUnits.Length.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerTfGusset")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerTfGusset = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerGussetBoltEdgeDistance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerGussetBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerGussetGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerGussetGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerGussetMatlGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerGussetMatlGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerAfMult")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerAfMult = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerArMult")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerArMult = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerFlatIPAPole")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerFlatIPAPole = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRoundIPAPole")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRoundIPAPole = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerFlatIPALeg")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerFlatIPALeg = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRoundIPALeg")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRoundIPALeg = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerFlatIPAHorizontal")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerFlatIPAHorizontal = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRoundIPAHorizontal")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRoundIPAHorizontal = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerFlatIPADiagonal")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerFlatIPADiagonal = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRoundIPADiagonal")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRoundIPADiagonal = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerCSA_S37_SpeedUpFactor")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerCSA_S37_SpeedUpFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerKLegs")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKLegs = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerKXBracedDiags")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKXBracedDiags = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerKKBracedDiags")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKKBracedDiags = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerKZBracedDiags")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKZBracedDiags = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerKHorzs")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKHorzs = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerKSecHorzs")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKSecHorzs = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerKGirts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKGirts = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerKInners")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKInners = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerKXBracedDiagsY")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKXBracedDiagsY = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerKKBracedDiagsY")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKKBracedDiagsY = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerKZBracedDiagsY")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKZBracedDiagsY = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerKHorzsY")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKHorzsY = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerKSecHorzsY")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKSecHorzsY = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerKGirtsY")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKGirtsY = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerKInnersY")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKInnersY = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerKRedHorz")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKRedHorz = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerKRedDiag")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKRedDiag = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerKRedSubDiag")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKRedSubDiag = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerKRedSubHorz")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKRedSubHorz = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerKRedVert")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKRedVert = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerKRedHip")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKRedHip = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerKRedHipDiag")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKRedHipDiag = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerKTLX")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKTLX = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerKTLZ")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKTLZ = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerKTLLeg")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKTLLeg = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerInnerKTLX")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerKTLX = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerInnerKTLZ")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerKTLZ = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerInnerKTLLeg")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerKTLLeg = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerStitchBoltLocationHoriz")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerStitchBoltLocationHoriz = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerStitchBoltLocationDiag")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerStitchBoltLocationDiag = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerStitchBoltLocationRed")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerStitchBoltLocationRed = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerStitchSpacing")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerStitchSpacing = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerStitchSpacingDiag")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerStitchSpacingDiag = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerStitchSpacingHorz")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerStitchSpacingHorz = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerStitchSpacingRed")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerStitchSpacingRed = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLegNetWidthDeduct")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerLegNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLegUFactor")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerLegUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerDiagonalNetWidthDeduct")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagonalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerTopGirtNetWidthDeduct")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerTopGirtNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerBotGirtNetWidthDeduct")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerBotGirtNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerInnerGirtNetWidthDeduct")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerGirtNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerHorizontalNetWidthDeduct")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerHorizontalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerShortHorizontalNetWidthDeduct")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerShortHorizontalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerDiagonalUFactor")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagonalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerTopGirtUFactor")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerTopGirtUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerBotGirtUFactor")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerBotGirtUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerInnerGirtUFactor")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerGirtUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerHorizontalUFactor")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerHorizontalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerShortHorizontalUFactor")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerShortHorizontalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLegConnType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerLegConnType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLegNumBolts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerLegNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerDiagonalNumBolts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagonalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerTopGirtNumBolts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerTopGirtNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerBotGirtNumBolts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerBotGirtNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerInnerGirtNumBolts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerGirtNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerHorizontalNumBolts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerHorizontalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerShortHorizontalNumBolts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerShortHorizontalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLegBoltGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerLegBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLegBoltSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerLegBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerDiagonalBoltGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagonalBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerDiagonalBoltSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagonalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerTopGirtBoltGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerTopGirtBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerTopGirtBoltSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerTopGirtBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerBotGirtBoltGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerBotGirtBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerBotGirtBoltSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerBotGirtBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerInnerGirtBoltGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerGirtBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerInnerGirtBoltSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerGirtBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerHorizontalBoltGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerHorizontalBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerHorizontalBoltSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerHorizontalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerShortHorizontalBoltGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerShortHorizontalBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerShortHorizontalBoltSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerShortHorizontalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLegBoltEdgeDistance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerLegBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerDiagonalBoltEdgeDistance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagonalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerTopGirtBoltEdgeDistance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerTopGirtBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerBotGirtBoltEdgeDistance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerBotGirtBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerInnerGirtBoltEdgeDistance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerGirtBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerHorizontalBoltEdgeDistance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerHorizontalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerShortHorizontalBoltEdgeDistance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerShortHorizontalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerDiagonalGageG1Distance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagonalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerTopGirtGageG1Distance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerTopGirtGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerBotGirtGageG1Distance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerBotGirtGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerInnerGirtGageG1Distance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerGirtGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerHorizontalGageG1Distance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerHorizontalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerShortHorizontalGageG1Distance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerShortHorizontalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHorizontalBoltGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHorizontalBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHorizontalBoltSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHorizontalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHorizontalNumBolts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHorizontalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHorizontalBoltEdgeDistance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHorizontalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHorizontalGageG1Distance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHorizontalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHorizontalNetWidthDeduct")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHorizontalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHorizontalUFactor")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHorizontalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantDiagonalBoltGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantDiagonalBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantDiagonalBoltSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantDiagonalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantDiagonalNumBolts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantDiagonalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantDiagonalBoltEdgeDistance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantDiagonalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantDiagonalGageG1Distance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantDiagonalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantDiagonalNetWidthDeduct")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantDiagonalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantDiagonalUFactor")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantDiagonalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubDiagonalBoltGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubDiagonalBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubDiagonalBoltSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubDiagonalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubDiagonalNumBolts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubDiagonalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubDiagonalBoltEdgeDistance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubDiagonalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubDiagonalGageG1Distance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubDiagonalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubDiagonalNetWidthDeduct")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubDiagonalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubDiagonalUFactor")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubDiagonalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubHorizontalBoltGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubHorizontalBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubHorizontalBoltSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubHorizontalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubHorizontalNumBolts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubHorizontalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubHorizontalBoltEdgeDistance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubHorizontalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubHorizontalGageG1Distance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubHorizontalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubHorizontalNetWidthDeduct")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubHorizontalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubHorizontalUFactor")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubHorizontalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantVerticalBoltGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantVerticalBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantVerticalBoltSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantVerticalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantVerticalNumBolts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantVerticalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantVerticalBoltEdgeDistance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantVerticalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantVerticalGageG1Distance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantVerticalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantVerticalNetWidthDeduct")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantVerticalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantVerticalUFactor")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantVerticalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipBoltGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipBoltSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipNumBolts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipBoltEdgeDistance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipGageG1Distance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipNetWidthDeduct")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipUFactor")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipDiagonalBoltGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipDiagonalBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipDiagonalBoltSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipDiagonalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipDiagonalNumBolts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipDiagonalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipDiagonalBoltEdgeDistance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipDiagonalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipDiagonalGageG1Distance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipDiagonalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipDiagonalNetWidthDeduct")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipDiagonalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipDiagonalUFactor")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipDiagonalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerDiagonalOutOfPlaneRestraint")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagonalOutOfPlaneRestraint = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerTopGirtOutOfPlaneRestraint")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerTopGirtOutOfPlaneRestraint = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerBottomGirtOutOfPlaneRestraint")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerBottomGirtOutOfPlaneRestraint = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerMidGirtOutOfPlaneRestraint")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerMidGirtOutOfPlaneRestraint = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerHorizontalOutOfPlaneRestraint")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerHorizontalOutOfPlaneRestraint = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerSecondaryHorizontalOutOfPlaneRestraint")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerSecondaryHorizontalOutOfPlaneRestraint = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerUniqueFlag")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerUniqueFlag = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerDiagOffsetNEY")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagOffsetNEY = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerDiagOffsetNEX")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagOffsetNEX = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerDiagOffsetPEY")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagOffsetPEY = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerDiagOffsetPEX")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagOffsetPEX = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerKbraceOffsetNEY")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKbraceOffsetNEY = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerKbraceOffsetNEX")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKbraceOffsetNEX = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerKbraceOffsetPEY")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKbraceOffsetPEY = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerKbraceOffsetPEX")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKbraceOffsetPEX = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                    End Select
                Case caseFilter = "Guy"
                    ''''Guy Rec''''
                    Select Case True
                        Case tnxVar.Equals("GuyRec")
                            Try
                                recIndex = CInt(tnxValue) - 1
                                Me.geometry.guyWires.Add(New tnxGuyRecord(Me.geometry))
                                Me.geometry.guyWires(recIndex).Rec = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyHeight")
                            Try
                                Me.geometry.guyWires(recIndex).GuyHeight = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyAutoCalcKSingleAngle")
                            Try
                                Me.geometry.guyWires(recIndex).GuyAutoCalcKSingleAngle = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyAutoCalcKSolidRound")
                            Try
                                Me.geometry.guyWires(recIndex).GuyAutoCalcKSolidRound = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyMount")
                            Try
                                Me.geometry.guyWires(recIndex).GuyMount = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TorqueArmStyle")
                            Try
                                Me.geometry.guyWires(recIndex).TorqueArmStyle = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyRadius")
                            Try
                                Me.geometry.guyWires(recIndex).GuyRadius = Me.settings.USUnits.Length.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyRadius120")
                            Try
                                Me.geometry.guyWires(recIndex).GuyRadius120 = Me.settings.USUnits.Length.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyRadius240")
                            Try
                                Me.geometry.guyWires(recIndex).GuyRadius240 = Me.settings.USUnits.Length.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyRadius360")
                            Try
                                Me.geometry.guyWires(recIndex).GuyRadius360 = Me.settings.USUnits.Length.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TorqueArmRadius")
                            Try
                                Me.geometry.guyWires(recIndex).TorqueArmRadius = Me.settings.USUnits.Length.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TorqueArmLegAngle")
                            Try
                                Me.geometry.guyWires(recIndex).TorqueArmLegAngle = Me.settings.USUnits.Rotation.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Azimuth0Adjustment")
                            Try
                                Me.geometry.guyWires(recIndex).Azimuth0Adjustment = Me.settings.USUnits.Rotation.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Azimuth120Adjustment")
                            Try
                                Me.geometry.guyWires(recIndex).Azimuth120Adjustment = Me.settings.USUnits.Rotation.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Azimuth240Adjustment")
                            Try
                                Me.geometry.guyWires(recIndex).Azimuth240Adjustment = Me.settings.USUnits.Rotation.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Azimuth360Adjustment")
                            Try
                                Me.geometry.guyWires(recIndex).Azimuth360Adjustment = Me.settings.USUnits.Rotation.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Anchor0Elevation")
                            Try
                                Me.geometry.guyWires(recIndex).Anchor0Elevation = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Anchor120Elevation")
                            Try
                                Me.geometry.guyWires(recIndex).Anchor120Elevation = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Anchor240Elevation")
                            Try
                                Me.geometry.guyWires(recIndex).Anchor240Elevation = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Anchor360Elevation")
                            Try
                                Me.geometry.guyWires(recIndex).Anchor360Elevation = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuySize")
                            Try
                                Me.geometry.guyWires(recIndex).GuySize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Guy120Size")
                            Try
                                Me.geometry.guyWires(recIndex).Guy120Size = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Guy240Size")
                            Try
                                Me.geometry.guyWires(recIndex).Guy240Size = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("Guy360Size")
                            Try
                                Me.geometry.guyWires(recIndex).Guy360Size = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyGrade")
                            Try
                                Me.geometry.guyWires(recIndex).GuyGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TorqueArmSize")
                            Try
                                Me.geometry.guyWires(recIndex).TorqueArmSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TorqueArmSizeBot")
                            Try
                                Me.geometry.guyWires(recIndex).TorqueArmSizeBot = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TorqueArmType")
                            Try
                                Me.geometry.guyWires(recIndex).TorqueArmType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TorqueArmGrade")
                            Try
                                Me.geometry.guyWires(recIndex).TorqueArmGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TorqueArmMatlGrade")
                            Try
                                Me.geometry.guyWires(recIndex).TorqueArmMatlGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TorqueArmKFactor")
                            Try
                                Me.geometry.guyWires(recIndex).TorqueArmKFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TorqueArmKFactorY")
                            Try
                                Me.geometry.guyWires(recIndex).TorqueArmKFactorY = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyPullOffKFactorX")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPullOffKFactorX = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyPullOffKFactorY")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPullOffKFactorY = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyDiagKFactorX")
                            Try
                                Me.geometry.guyWires(recIndex).GuyDiagKFactorX = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyDiagKFactorY")
                            Try
                                Me.geometry.guyWires(recIndex).GuyDiagKFactorY = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyAutoCalc")
                            Try
                                Me.geometry.guyWires(recIndex).GuyAutoCalc = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyAllGuysSame")
                            Try
                                Me.geometry.guyWires(recIndex).GuyAllGuysSame = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyAllGuysAnchorSame")
                            Try
                                Me.geometry.guyWires(recIndex).GuyAllGuysAnchorSame = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyIsStrapping")
                            Try
                                Me.geometry.guyWires(recIndex).GuyIsStrapping = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyPullOffSize")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPullOffSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyPullOffSizeBot")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPullOffSizeBot = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyPullOffType")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPullOffType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyPullOffGrade")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPullOffGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyPullOffMatlGrade")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPullOffMatlGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyUpperDiagSize")
                            Try
                                Me.geometry.guyWires(recIndex).GuyUpperDiagSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyLowerDiagSize")
                            Try
                                Me.geometry.guyWires(recIndex).GuyLowerDiagSize = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyDiagType")
                            Try
                                Me.geometry.guyWires(recIndex).GuyDiagType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyDiagGrade")
                            Try
                                Me.geometry.guyWires(recIndex).GuyDiagGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyDiagMatlGrade")
                            Try
                                Me.geometry.guyWires(recIndex).GuyDiagMatlGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyDiagNetWidthDeduct")
                            Try
                                Me.geometry.guyWires(recIndex).GuyDiagNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyDiagUFactor")
                            Try
                                Me.geometry.guyWires(recIndex).GuyDiagUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyDiagNumBolts")
                            Try
                                Me.geometry.guyWires(recIndex).GuyDiagNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyDiagonalOutOfPlaneRestraint")
                            Try
                                Me.geometry.guyWires(recIndex).GuyDiagonalOutOfPlaneRestraint = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyDiagBoltGrade")
                            Try
                                Me.geometry.guyWires(recIndex).GuyDiagBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyDiagBoltSize")
                            Try
                                Me.geometry.guyWires(recIndex).GuyDiagBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyDiagBoltEdgeDistance")
                            Try
                                Me.geometry.guyWires(recIndex).GuyDiagBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyDiagBoltGageDistance")
                            Try
                                Me.geometry.guyWires(recIndex).GuyDiagBoltGageDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyPullOffNetWidthDeduct")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPullOffNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyPullOffUFactor")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPullOffUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyPullOffNumBolts")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPullOffNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyPullOffOutOfPlaneRestraint")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPullOffOutOfPlaneRestraint = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyPullOffBoltGrade")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPullOffBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyPullOffBoltSize")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPullOffBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyPullOffBoltEdgeDistance")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPullOffBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyPullOffBoltGageDistance")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPullOffBoltGageDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyTorqueArmNetWidthDeduct")
                            Try
                                Me.geometry.guyWires(recIndex).GuyTorqueArmNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyTorqueArmUFactor")
                            Try
                                Me.geometry.guyWires(recIndex).GuyTorqueArmUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyTorqueArmNumBolts")
                            Try
                                Me.geometry.guyWires(recIndex).GuyTorqueArmNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyTorqueArmOutOfPlaneRestraint")
                            Try
                                Me.geometry.guyWires(recIndex).GuyTorqueArmOutOfPlaneRestraint = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyTorqueArmBoltGrade")
                            Try
                                Me.geometry.guyWires(recIndex).GuyTorqueArmBoltGrade = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyTorqueArmBoltSize")
                            Try
                                Me.geometry.guyWires(recIndex).GuyTorqueArmBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyTorqueArmBoltEdgeDistance")
                            Try
                                Me.geometry.guyWires(recIndex).GuyTorqueArmBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyTorqueArmBoltGageDistance")
                            Try
                                Me.geometry.guyWires(recIndex).GuyTorqueArmBoltGageDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyPerCentTension")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPerCentTension = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyPerCentTension120")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPerCentTension120 = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyPerCentTension240")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPerCentTension240 = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyPerCentTension360")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPerCentTension360 = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyEffFactor")
                            Try
                                Me.geometry.guyWires(recIndex).GuyEffFactor = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyEffFactor120")
                            Try
                                Me.geometry.guyWires(recIndex).GuyEffFactor120 = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyEffFactor240")
                            Try
                                Me.geometry.guyWires(recIndex).GuyEffFactor240 = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyEffFactor360")
                            Try
                                Me.geometry.guyWires(recIndex).GuyEffFactor360 = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyNumInsulators")
                            Try
                                Me.geometry.guyWires(recIndex).GuyNumInsulators = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyInsulatorLength")
                            Try
                                Me.geometry.guyWires(recIndex).GuyInsulatorLength = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyInsulatorDia")
                            Try
                                Me.geometry.guyWires(recIndex).GuyInsulatorDia = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("GuyInsulatorWt")
                            Try
                                Me.geometry.guyWires(recIndex).GuyInsulatorWt = Me.settings.USUnits.Force.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                    End Select
                Case caseFilter = "FeedLine"
                    ''''Feed Lines''''
                    Select Case True
                        Case tnxVar.Equals("FeedLineRec")
                            Try
                                recIndex = CInt(tnxValue) - 1
                                Me.feedLines.Add(New tnxFeedLine(Me))
                                Me.feedLines(recIndex).FeedLineRec = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineEnabled")
                            Try
                                Me.feedLines(recIndex).FeedLineEnabled = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineDatabase")
                            Try
                                Me.feedLines(recIndex).FeedLineDatabase = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineDescription")
                            Try
                                Me.feedLines(recIndex).FeedLineDescription = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineClassificationCategory")
                            Try
                                Me.feedLines(recIndex).FeedLineClassificationCategory = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineNote")
                            Try
                                Me.feedLines(recIndex).FeedLineNote = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineNum")
                            Try
                                Me.feedLines(recIndex).FeedLineNum = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineUseShielding")
                            Try
                                Me.feedLines(recIndex).FeedLineUseShielding = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("ExcludeFeedLineFromTorque")
                            Try
                                Me.feedLines(recIndex).ExcludeFeedLineFromTorque = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineNumPerRow")
                            Try
                                Me.feedLines(recIndex).FeedLineNumPerRow = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineFace")
                            Try
                                Me.feedLines(recIndex).FeedLineFace = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineComponentType")
                            Try
                                Me.feedLines(recIndex).FeedLineComponentType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineGroupTreatmentType")
                            Try
                                Me.feedLines(recIndex).FeedLineGroupTreatmentType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineRoundClusterDia")
                            Try
                                Me.feedLines(recIndex).FeedLineRoundClusterDia = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineWidth")
                            Try
                                Me.feedLines(recIndex).FeedLineWidth = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLinePerimeter")
                            Try
                                Me.feedLines(recIndex).FeedLinePerimeter = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FlatAttachmentEffectiveWidthRatio")
                            Try
                                Me.feedLines(recIndex).FlatAttachmentEffectiveWidthRatio = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("AutoCalcFlatAttachmentEffectiveWidthRatio")
                            Try
                                Me.feedLines(recIndex).AutoCalcFlatAttachmentEffectiveWidthRatio = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineShieldingFactorKaNoIce")
                            Try
                                Me.feedLines(recIndex).FeedLineShieldingFactorKaNoIce = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineShieldingFactorKaIce")
                            Try
                                Me.feedLines(recIndex).FeedLineShieldingFactorKaIce = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineAutoCalcKa")
                            Try
                                Me.feedLines(recIndex).FeedLineAutoCalcKa = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineCaAaNoIce")
                            'The units of this input are dependant on the TNX length unit setting but they are an area (in^2 or ft^2)
                            Try
                                Me.feedLines(recIndex).FeedLineCaAaNoIce = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineCaAaIce")
                            'The units of this input are dependant on the TNX length unit setting but they are an area (in^2 or ft^2)
                            Try
                                Me.feedLines(recIndex).FeedLineCaAaIce = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineCaAaIce_1")
                            'The units of this input are dependant on the TNX length unit setting but they are an area (in^2 or ft^2)
                            Try
                                Me.feedLines(recIndex).FeedLineCaAaIce_1 = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineCaAaIce_2")
                            'The units of this input are dependant on the TNX length unit setting but they are an area (in^2 or ft^2)
                            Try
                                Me.feedLines(recIndex).FeedLineCaAaIce_2 = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineCaAaIce_4")
                            'The units of this input are dependant on the TNX length unit setting but they are an area (in^2 or ft^2)
                            Try
                                Me.feedLines(recIndex).FeedLineCaAaIce_4 = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineWtNoIce")
                            Try
                                Me.feedLines(recIndex).FeedLineWtNoIce = Me.settings.USUnits.Load.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineWtIce")
                            Try
                                Me.feedLines(recIndex).FeedLineWtIce = Me.settings.USUnits.Load.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineWtIce_1")
                            Try
                                Me.feedLines(recIndex).FeedLineWtIce_1 = Me.settings.USUnits.Load.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineWtIce_2")
                            Try
                                Me.feedLines(recIndex).FeedLineWtIce_2 = Me.settings.USUnits.Load.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineWtIce_4")
                            Try
                                Me.feedLines(recIndex).FeedLineWtIce_4 = Me.settings.USUnits.Load.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineFaceOffset")
                            Try
                                Me.feedLines(recIndex).FeedLineFaceOffset = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineOffsetFrac")
                            Try
                                Me.feedLines(recIndex).FeedLineOffsetFrac = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLinePerimeterOffsetStartFrac")
                            Try
                                Me.feedLines(recIndex).FeedLinePerimeterOffsetStartFrac = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLinePerimeterOffsetEndFrac")
                            Try
                                Me.feedLines(recIndex).FeedLinePerimeterOffsetEndFrac = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineStartHt")
                            Try
                                Me.feedLines(recIndex).FeedLineStartHt = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineEndHt")
                            Try
                                Me.feedLines(recIndex).FeedLineEndHt = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineClearSpacing")
                            Try
                                Me.feedLines(recIndex).FeedLineClearSpacing = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("FeedLineRowClearSpacing")
                            Try
                                Me.feedLines(recIndex).FeedLineRowClearSpacing = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                    End Select
                Case caseFilter = "Discrete"
                    ''''Discrete''''
                    Select Case True
                        Case tnxVar.Equals("TowerLoadRec")
                            Try
                                recIndex = CInt(tnxValue) - 1
                                Me.discreteLoads.Add(New tnxDiscreteLoad(Me))
                                Me.discreteLoads(recIndex).TowerLoadRec = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLoadEnabled")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadEnabled = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLoadDatabase")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadDatabase = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLoadDescription")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadDescription = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLoadType")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLoadClassificationCategory")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadClassificationCategory = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLoadNote")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadNote = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLoadNum")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadNum = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLoadFace")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadFace = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerOffsetType")
                            Try
                                Me.discreteLoads(recIndex).TowerOffsetType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerOffsetDist")
                            Try
                                Me.discreteLoads(recIndex).TowerOffsetDist = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerVertOffset")
                            Try
                                Me.discreteLoads(recIndex).TowerVertOffset = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLateralOffset")
                            Try
                                Me.discreteLoads(recIndex).TowerLateralOffset = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerAzimuthAdjustment")
                            Try
                                Me.discreteLoads(recIndex).TowerAzimuthAdjustment = Me.settings.USUnits.Rotation.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerAppurtSymbol")
                            Try
                                Me.discreteLoads(recIndex).TowerAppurtSymbol = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLoadShieldingFactorKaNoIce")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadShieldingFactorKaNoIce = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLoadShieldingFactorKaIce")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadShieldingFactorKaIce = CDbl(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLoadAutoCalcKa")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadAutoCalcKa = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLoadCaAaNoIce")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadCaAaNoIce = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLoadCaAaIce")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadCaAaIce = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLoadCaAaIce_1")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadCaAaIce_1 = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLoadCaAaIce_2")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadCaAaIce_2 = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLoadCaAaIce_4")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadCaAaIce_4 = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLoadCaAaNoIce_Side")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadCaAaNoIce_Side = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLoadCaAaIce_Side")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadCaAaIce_Side = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLoadCaAaIce_Side_1")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadCaAaIce_Side_1 = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLoadCaAaIce_Side_2")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadCaAaIce_Side_2 = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLoadCaAaIce_Side_4")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadCaAaIce_Side_4 = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLoadWtNoIce")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadWtNoIce = Me.settings.USUnits.Force.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLoadWtIce")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadWtIce = Me.settings.USUnits.Force.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLoadWtIce_1")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadWtIce_1 = Me.settings.USUnits.Force.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLoadWtIce_2")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadWtIce_2 = Me.settings.USUnits.Force.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLoadWtIce_4")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadWtIce_4 = Me.settings.USUnits.Force.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLoadStartHt")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadStartHt = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("TowerLoadEndHt")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadEndHt = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                    End Select
                Case caseFilter = "Dish"
                    ''''Dishes''''
                    Select Case True
                        Case tnxVar.Equals("DishRec")
                            Try
                                recIndex = CInt(tnxValue) - 1
                                Me.dishes.Add(New tnxDish(Me))
                                Me.dishes(recIndex).DishRec = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DishEnabled")
                            Try
                                Me.dishes(recIndex).DishEnabled = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DishDatabase")
                            Try
                                Me.dishes(recIndex).DishDatabase = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DishDescription")
                            Try
                                Me.dishes(recIndex).DishDescription = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DishClassificationCategory")
                            Try
                                Me.dishes(recIndex).DishClassificationCategory = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DishNote")
                            Try
                                Me.dishes(recIndex).DishNote = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DishNum")
                            Try
                                Me.dishes(recIndex).DishNum = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DishFace")
                            Try
                                Me.dishes(recIndex).DishFace = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DishType")
                            Try
                                Me.dishes(recIndex).DishType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DishOffsetType")
                            Try
                                Me.dishes(recIndex).DishOffsetType = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DishVertOffset")
                            Try
                                Me.dishes(recIndex).DishVertOffset = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DishLateralOffset")
                            Try
                                Me.dishes(recIndex).DishLateralOffset = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DishOffsetDist")
                            Try
                                Me.dishes(recIndex).DishOffsetDist = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DishArea")
                            Try
                                Me.dishes(recIndex).DishArea = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DishAreaIce")
                            Try
                                Me.dishes(recIndex).DishAreaIce = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DishAreaIce_1")
                            Try
                                Me.dishes(recIndex).DishAreaIce_1 = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DishAreaIce_2")
                            Try
                                Me.dishes(recIndex).DishAreaIce_2 = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DishAreaIce_4")
                            Try
                                Me.dishes(recIndex).DishAreaIce_4 = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DishDiameter")
                            Try
                                Me.dishes(recIndex).DishDiameter = Me.settings.USUnits.Length.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DishWtNoIce")
                            Try
                                Me.dishes(recIndex).DishWtNoIce = Me.settings.USUnits.Force.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DishWtIce")
                            Try
                                Me.dishes(recIndex).DishWtIce = Me.settings.USUnits.Force.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DishWtIce_1")
                            Try
                                Me.dishes(recIndex).DishWtIce_1 = Me.settings.USUnits.Force.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DishWtIce_2")
                            Try
                                Me.dishes(recIndex).DishWtIce_2 = Me.settings.USUnits.Force.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DishWtIce_4")
                            Try
                                Me.dishes(recIndex).DishWtIce_4 = Me.settings.USUnits.Force.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DishStartHt")
                            Try
                                Me.dishes(recIndex).DishStartHt = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DishAzimuthAdjustment")
                            Try
                                Me.dishes(recIndex).DishAzimuthAdjustment = Me.settings.USUnits.Rotation.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("DishBeamWidth")
                            Try
                                Me.dishes(recIndex).DishBeamWidth = Me.settings.USUnits.Rotation.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                    End Select
                Case caseFilter = "UserForce"
                    ''''UserForces''''
                    Select Case True
                        Case tnxVar.Equals("UserForceRec")
                            Try
                                recIndex = CInt(tnxValue) - 1
                                Me.userForces.Add(New tnxUserForce(Me))
                                Me.userForces(recIndex).UserForceRec = CInt(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UserForceEnabled")
                            Try
                                Me.userForces(recIndex).UserForceEnabled = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UserForceDescription")
                            Try
                                Me.userForces(recIndex).UserForceDescription = tnxValue
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UserForceStartHt")
                            Try
                                Me.userForces(recIndex).UserForceStartHt = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UserForceOffset")
                            Try
                                Me.userForces(recIndex).UserForceOffset = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UserForceAzimuth")
                            Try
                                Me.userForces(recIndex).UserForceAzimuth = Me.settings.USUnits.Rotation.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UserForceFxNoIce")
                            Try
                                Me.userForces(recIndex).UserForceFxNoIce = Me.settings.USUnits.Force.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UserForceFzNoIce")
                            Try
                                Me.userForces(recIndex).UserForceFzNoIce = Me.settings.USUnits.Force.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UserForceAxialNoIce")
                            Try
                                Me.userForces(recIndex).UserForceAxialNoIce = Me.settings.USUnits.Force.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UserForceShearNoIce")
                            Try
                                Me.userForces(recIndex).UserForceShearNoIce = Me.settings.USUnits.Force.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UserForceCaAcNoIce")
                            Try
                                Me.userForces(recIndex).UserForceCaAcNoIce = Me.settings.USUnits.Length.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UserForceFxIce")
                            Try
                                Me.userForces(recIndex).UserForceFxIce = Me.settings.USUnits.Force.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UserForceFzIce")
                            Try
                                Me.userForces(recIndex).UserForceFzIce = Me.settings.USUnits.Force.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UserForceAxialIce")
                            Try
                                Me.userForces(recIndex).UserForceAxialIce = Me.settings.USUnits.Force.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UserForceShearIce")
                            Try
                                Me.userForces(recIndex).UserForceShearIce = Me.settings.USUnits.Force.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UserForceCaAcIce")
                            Try
                                Me.userForces(recIndex).UserForceCaAcIce = Me.settings.USUnits.Length.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UserForceFxService")
                            Try
                                Me.userForces(recIndex).UserForceFxService = Me.settings.USUnits.Force.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UserForceFzService")
                            Try
                                Me.userForces(recIndex).UserForceFzService = Me.settings.USUnits.Force.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UserForceAxialService")
                            Try
                                Me.userForces(recIndex).UserForceAxialService = Me.settings.USUnits.Force.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UserForceShearService")
                            Try
                                Me.userForces(recIndex).UserForceShearService = Me.settings.USUnits.Force.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UserForceCaAcService")
                            Try
                                Me.userForces(recIndex).UserForceCaAcService = Me.settings.USUnits.Length.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UserForceEhx")
                            Try
                                Me.userForces(recIndex).UserForceEhx = Me.settings.USUnits.Force.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UserForceEhz")
                            Try
                                Me.userForces(recIndex).UserForceEhz = Me.settings.USUnits.Force.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UserForceEv")
                            Try
                                Me.userForces(recIndex).UserForceEv = Me.settings.USUnits.Force.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                        Case tnxVar.Equals("UserForceEh")
                            Try
                                Me.userForces(recIndex).UserForceEh = Me.settings.USUnits.Force.convertToEDSDefaultUnits(tnxValue)
                            Catch ex As Exception
                                Debug.Print("Error parsing TNX variable: " & tnxVar)
                            End Try
                    End Select
            End Select

        Next

        Me.GetResults()

    End Sub

#End Region

#Region "Save to EDS"
    Public Overrides Function EDSQueryBuilder(ItemToCompare As EDSObjectWithQueries, Optional ByRef AllowUpdate As Boolean = True) As String
        Dim TNXToCompare As tnxModel = TryCast(ItemToCompare, tnxModel)
        If ItemToCompare IsNot Nothing And TNXToCompare Is Nothing Then
            Return ""
        Else
            Return Me.EDSQueryBuilder(TNXToCompare, AllowUpdate)
        End If
    End Function
    Public Overloads Function EDSQueryBuilder(TNXToCompare As tnxModel, Optional ByRef AllowUpdate As Boolean = True) As String
        'Compare the ID of the current EDS item to the existing item and determine if the Insert, Update, or Delete query should be used

        EDSQueryBuilder = ""

        'TNX from file won't ever have ID and we're only storing one TNX file per site. If one exists, update it. No need to ever delete.
        If TNXToCompare Is Nothing Then
            EDSQueryBuilder += Me.SQLInsert
            AllowUpdate = False
        Else
            If AllowUpdate Then
                EDSQueryBuilder += Me.SQLSetID(TNXToCompare.ID)
                If Not Me.Equals(TNXToCompare, Nothing, True, True) Then
                    EDSQueryBuilder += Me.SQLUpdate()
                End If
            Else
                EDSQueryBuilder += TNXToCompare.SQLDelete
                EDSQueryBuilder += Me.SQLInsert
                AllowUpdate = False
            End If
        End If

        EDSQueryBuilder += InsertMyLoading()

        'Database
        EDSQueryBuilder += Me.database.members.TNXMemberListQueryBuilder(TNXToCompare?.database.members)
        EDSQueryBuilder += Me.database.materials.TNXMemberListQueryBuilder(TNXToCompare?.database.materials)
        EDSQueryBuilder += Me.database.bolts.TNXMemberListQueryBuilder(TNXToCompare?.database.bolts)

        'Geometry
        EDSQueryBuilder += Me.geometry.upperStructure.TNXGeometryRecListQueryBuilder(TNXToCompare?.geometry.upperStructure, AllowUpdate)
        EDSQueryBuilder += Me.geometry.baseStructure.TNXGeometryRecListQueryBuilder(TNXToCompare?.geometry.baseStructure, AllowUpdate)
        EDSQueryBuilder += Me.geometry.guyWires.TNXGeometryRecListQueryBuilder(TNXToCompare?.geometry.guyWires, AllowUpdate)

        'EDSQueryBuilder += loadQuery
        EDSQueryBuilder += "SET " & EDSStructure.SQLQueryIDVar(Me.EDSTableDepth) & " = NULL" & vbCrLf

        Return EDSQueryBuilder

    End Function

    Private Function InsertMyLoading() As String

        'DELETE all loads associated with the tnx model and insert all new ones. 
        'We may need to look into where this query is added. There are a lot of different pieces to the tnx file. 
        Dim discQuery As String = ""
        Dim lineQuery As String = ""
        Dim userQuery As String = ""
        Dim dishQuery As String = ""
        Dim loadQuery As String = ""

        loadQuery += "DELETE FROM load.discrete_output WHERE tnx_id = @TopLevelID" & vbCrLf
        For Each disc In Me.discreteLoads
            discQuery += disc.SQLInsert & vbCrLf
        Next
        loadQuery += discQuery & vbCrLf

        loadQuery += "DELETE FROM load.linear_output WHERE tnx_id = @TopLevelID" & vbCrLf
        For Each line In Me.feedLines
            lineQuery += line.SQLInsert & vbCrLf
        Next
        loadQuery += lineQuery & vbCrLf

        loadQuery += "DELETE FROM load.user_force_output WHERE tnx_id = @TopLevelID" & vbCrLf
        For Each user In Me.userForces
            userQuery += user.SQLInsert & vbCrLf
        Next
        loadQuery += userQuery & vbCrLf

        loadQuery += "DELETE FROM load.dish_output WHERE tnx_id = @TopLevelID" & vbCrLf
        For Each dish In Me.dishes
            dishQuery += dish.SQLInsert & vbCrLf
        Next
        loadQuery += dishQuery & vbCrLf

        Return loadQuery
    End Function

    'Private _Insert As String
    'Private _Update As String
    'Private _Delete As String
    Public Overrides Function SQLInsert() As String
        SQLInsert = "BEGIN" & vbCrLf &
                 "  INSERT INTO " & Me.EDSTableName & " (" & Me.SQLInsertFields & ")" & vbCrLf &
                 "  OUTPUT INSERTED.ID INTO @TopLevel" & vbCrLf &
                 "  VALUES(" & Me.SQLInsertValues & ")" & vbCrLf &
                 "  SELECT @TopLevelID=ID FROM @TopLevel" & vbCrLf &
                 "  DELETE FROM @TopLevel" & vbCrLf &
                 "END" & vbCrLf
        Return SQLInsert
    End Function

    'Public Overrides Function SQLUpdate() As String
    '    Return Me.SQLUpdate(Nothing)
    'End Function
    Public Overrides Function SQLUpdate() As String
        'For any EDSObject that has sub-objects we will need to overload the update property with a version that excepts the current version being updated.
        SQLUpdate = "BEGIN" & vbCrLf &
                      "  UPDATE " & Me.EDSTableName & vbCrLf &
                      "  SET " & Me.SQLUpdateFieldsandValues & vbCrLf &
                      "  WHERE ID = @TopLevelID" & vbCrLf &
                      "END" & vbCrLf
        Return SQLUpdate
    End Function

    Public Overrides Function SQLDelete() As String
        SQLDelete = "BEGIN" & vbCrLf &
                     "  DELETE FROM load.dish_output WHERE tnx_id = " & Me.ID.ToString.FormatDBValue & vbCrLf &
                     "  DELETE FROM load.discrete_output WHERE tnx_id = " & Me.ID.ToString.FormatDBValue & vbCrLf &
                     "  DELETE FROM load.linear_output WHERE tnx_id = " & Me.ID.ToString.FormatDBValue & vbCrLf &
                     "  DELETE FROM load.user_force_output WHERE tnx_id = " & Me.ID.ToString.FormatDBValue & vbCrLf &
                     "  DELETE FROM tnx.upper_structure_sections WHERE tnx_id = " & Me.ID.ToString.FormatDBValue & vbCrLf &
                     "  DELETE FROM tnx.base_structure_sections WHERE tnx_id = " & Me.ID.ToString.FormatDBValue & vbCrLf &
                     "  DELETE FROM tnx.guys WHERE tnx_id = " & Me.ID.ToString.FormatDBValue & vbCrLf &
                     "  DELETE FROM tnx.members_xref WHERE tnx_id = " & Me.ID.ToString.FormatDBValue & vbCrLf &
                     "  DELETE FROM tnx.materials_xref WHERE tnx_id = " & Me.ID.ToString.FormatDBValue & vbCrLf &
                     "  DELETE FROM tnx.tnx WHERE ID = " & Me.ID.ToString.FormatDBValue & vbCrLf &
                     "END"
        Return SQLDelete
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        SQLInsertFields = SQLInsertFields.AddtoDBString("bus_unit")
        SQLInsertFields = SQLInsertFields.AddtoDBString("structure_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DesignStandardSeries")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UnitsSystem")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ClientName")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ProjectName")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ProjectNumber")
        SQLInsertFields = SQLInsertFields.AddtoDBString("CreatedBy")
        SQLInsertFields = SQLInsertFields.AddtoDBString("CreatedOn")
        SQLInsertFields = SQLInsertFields.AddtoDBString("LastUsedBy")
        SQLInsertFields = SQLInsertFields.AddtoDBString("LastUsedOn")
        SQLInsertFields = SQLInsertFields.AddtoDBString("VersionUsed")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USLength")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USLengthPrec")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USCoordinate")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USCoordinatePrec")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USForce")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USForcePrec")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USLoad")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USLoadPrec")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USMoment")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USMomentPrec")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USProperties")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USPropertiesPrec")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USPressure")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USPressurePrec")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USVelocity")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USVelocityPrec")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USDisplacement")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USDisplacementPrec")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USMass")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USMassPrec")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USAcceleration")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USAccelerationPrec")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USStress")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USStressPrec")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USDensity")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USDensityPrec")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USUnitWt")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USUnitWtPrec")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USStrength")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USStrengthPrec")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USModulus")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USModulusPrec")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USTemperature")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USTemperaturePrec")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USPrinter")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USPrinterPrec")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USRotation")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USRotationPrec")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USSpacing")
        SQLInsertFields = SQLInsertFields.AddtoDBString("USSpacingPrec")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ViewerUserName")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ViewerCompanyName")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ViewerStreetAddress")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ViewerCityState")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ViewerPhone")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ViewerFAX")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ViewerLogo")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ViewerCompanyBitmap")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportProjectNumber")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportJobType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportCarrierName")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportCarrierSiteNumber")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportCarrierSiteName")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportSiteAddress")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportLatitudeDegree")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportLatitudeMinute")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportLatitudeSecond")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportLongitudeDegree")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportLongitudeMinute")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportLongitudeSecond")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportLocalCodeRequirement")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportSiteHistory")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportTowerManufacturer")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportMonthManufactured")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportYearManufactured")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportOriginalSpeed")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportOriginalCode")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportTowerType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportEngrName")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportEngrTitle")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportHQPhoneNumber")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportEmailAddress")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportLogoPath")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportCCiContactName")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportCCiAddress1")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportCCiAddress2")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportCCiBUNumber")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportCCiSiteName")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportCCiJDENumber")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportCCiWONumber")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportCCiPONumber")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportCCiAppNumber")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportCCiRevNumber")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportRecommendations")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportAppurt1Note1")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportAppurt1Note2")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportAppurt1Note3")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportAppurt1Note4")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportAppurt1Note5")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportAppurt1Note6")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportAppurt1Note7")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportAppurt2Note1")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportAppurt2Note2")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportAppurt2Note3")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportAppurt2Note4")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportAppurt2Note5")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportAppurt2Note6")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportAppurt2Note7")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportAddlCapacityNote1")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportAddlCapacityNote2")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportAddlCapacityNote3")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sReportAddlCapacityNote4")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DesignCode")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("OverallHeight")
        SQLInsertFields = SQLInsertFields.AddtoDBString("BaseElevation")
        SQLInsertFields = SQLInsertFields.AddtoDBString("Lambda")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerTopFaceWidth")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerBaseFaceWidth")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindSpeed")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindSpeedIce")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindSpeedService")
        SQLInsertFields = SQLInsertFields.AddtoDBString("IceThickness")
        SQLInsertFields = SQLInsertFields.AddtoDBString("CSA_S37_RefVelPress")
        SQLInsertFields = SQLInsertFields.AddtoDBString("CSA_S37_ReliabilityClass")
        SQLInsertFields = SQLInsertFields.AddtoDBString("CSA_S37_ServiceabilityFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UseModified_TIA_222_IceParameters")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TIA_222_IceThicknessMultiplier")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DoNotUse_TIA_222_IceEscalation")
        SQLInsertFields = SQLInsertFields.AddtoDBString("IceDensity")
        SQLInsertFields = SQLInsertFields.AddtoDBString("SeismicSiteClass")
        SQLInsertFields = SQLInsertFields.AddtoDBString("SeismicSs")
        SQLInsertFields = SQLInsertFields.AddtoDBString("SeismicS1")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TempDrop")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GroutFc")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GirtOffset")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GirtOffsetLatticedPole")
        SQLInsertFields = SQLInsertFields.AddtoDBString("MastVert")
        SQLInsertFields = SQLInsertFields.AddtoDBString("MastHorz")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyVert")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyHorz")
        SQLInsertFields = SQLInsertFields.AddtoDBString("HogRodTakeup")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerTaper")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuyedMonopoleBaseType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TaperHeight")
        SQLInsertFields = SQLInsertFields.AddtoDBString("PivotHeight")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AutoCalcGH")
        SQLInsertFields = SQLInsertFields.AddtoDBString("IncludeCapacityNote")
        SQLInsertFields = SQLInsertFields.AddtoDBString("IncludeAppurtGraphics")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DisplayNotes")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DisplayReactions")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DisplaySchedule")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DisplayAppurtenanceTable")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DisplayMaterialStrengthTable")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AutoCalc_ASCE_GH")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ASCE_ExposureCat")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ASCE_Year")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ASCEGh")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ASCEI")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UseASCEWind")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UserGHElev")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UseCodeGuySF")
        SQLInsertFields = SQLInsertFields.AddtoDBString("GuySF")
        SQLInsertFields = SQLInsertFields.AddtoDBString("CalcWindAt")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerBoltGrade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TowerBoltMinEdgeDist")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AllowStressRatio")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AllowAntStressRatio")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindCalcPoints")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UseIndexPlate")
        SQLInsertFields = SQLInsertFields.AddtoDBString("EnterUserDefinedGhValues")
        SQLInsertFields = SQLInsertFields.AddtoDBString("BaseTowerGhInput")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UpperStructureGhInput")
        SQLInsertFields = SQLInsertFields.AddtoDBString("EnterUserDefinedCgValues")
        SQLInsertFields = SQLInsertFields.AddtoDBString("BaseTowerCgInput")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UpperStructureCgInput")
        SQLInsertFields = SQLInsertFields.AddtoDBString("CheckVonMises")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UseClearSpans")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UseClearSpansKlr")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AntennaFaceWidth")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DoInteraction")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DoHorzInteraction")
        SQLInsertFields = SQLInsertFields.AddtoDBString("DoDiagInteraction")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UseMomentMagnification")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UseFeedlineAsCylinder")
        SQLInsertFields = SQLInsertFields.AddtoDBString("OffsetBotGirt")
        SQLInsertFields = SQLInsertFields.AddtoDBString("PrintBitmaps")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UseTopTakeup")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ConstantSlope")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UseCodeStressRatio")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UseLegLoads")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ERIDesignMode")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindExposure")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindZone")
        SQLInsertFields = SQLInsertFields.AddtoDBString("StructureCategory")
        SQLInsertFields = SQLInsertFields.AddtoDBString("RiskCategory")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TopoCategory")
        SQLInsertFields = SQLInsertFields.AddtoDBString("RSMTopographicFeature")
        SQLInsertFields = SQLInsertFields.AddtoDBString("RSM_L")
        SQLInsertFields = SQLInsertFields.AddtoDBString("RSM_X")
        SQLInsertFields = SQLInsertFields.AddtoDBString("CrestHeight")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TIA_222_H_TopoFeatureDownwind")
        SQLInsertFields = SQLInsertFields.AddtoDBString("BaseElevAboveSeaLevel")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ConsiderRooftopSpeedUp")
        SQLInsertFields = SQLInsertFields.AddtoDBString("RooftopWS")
        SQLInsertFields = SQLInsertFields.AddtoDBString("RooftopHS")
        SQLInsertFields = SQLInsertFields.AddtoDBString("RooftopParapetHt")
        SQLInsertFields = SQLInsertFields.AddtoDBString("RooftopXB")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UseTIA222H_AnnexS")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TIA_222_H_AnnexS_Ratio")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UseASCE7_10_Seismic_LComb")
        SQLInsertFields = SQLInsertFields.AddtoDBString("EIACWindMult")
        SQLInsertFields = SQLInsertFields.AddtoDBString("EIACWindMultIce")
        SQLInsertFields = SQLInsertFields.AddtoDBString("EIACIgnoreCableDrag")
        SQLInsertFields = SQLInsertFields.AddtoDBString("Notes")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ReportInputCosts")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ReportInputGeometry")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ReportInputOptions")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ReportMaxForces")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ReportInputMap")
        SQLInsertFields = SQLInsertFields.AddtoDBString("CostReportOutputType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("CapacityReportOutputType")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ReportPrintForceTotals")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ReportPrintForceDetails")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ReportPrintMastVectors")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ReportPrintAntPoleVectors")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ReportPrintDiscreteVectors")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ReportPrintDishVectors")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ReportPrintFeedTowerVectors")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ReportPrintUserLoadVectors")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ReportPrintPressures")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ReportPrintAppurtForces")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ReportPrintGuyForces")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ReportPrintGuyStressing")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ReportPrintDeflections")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ReportPrintReactions")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ReportPrintStressChecks")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ReportPrintBoltChecks")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ReportPrintInputGVerificationTables")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ReportPrintOutputGVerificationTables")
        SQLInsertFields = SQLInsertFields.AddtoDBString("SocketTopMount")
        SQLInsertFields = SQLInsertFields.AddtoDBString("SRTakeCompression")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AllLegPanelsSame")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UseCombinedBoltCapacity")
        SQLInsertFields = SQLInsertFields.AddtoDBString("SecHorzBracesLeg")
        SQLInsertFields = SQLInsertFields.AddtoDBString("SortByComponent")
        SQLInsertFields = SQLInsertFields.AddtoDBString("SRCutEnds")
        SQLInsertFields = SQLInsertFields.AddtoDBString("SRConcentric")
        SQLInsertFields = SQLInsertFields.AddtoDBString("CalcBlockShear")
        SQLInsertFields = SQLInsertFields.AddtoDBString("Use4SidedDiamondBracing")
        SQLInsertFields = SQLInsertFields.AddtoDBString("TriangulateInnerBracing")
        SQLInsertFields = SQLInsertFields.AddtoDBString("PrintCarrierNotes")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AddIBCWindCase")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UseStateCountyLookup")
        SQLInsertFields = SQLInsertFields.AddtoDBString("State")
        SQLInsertFields = SQLInsertFields.AddtoDBString("County")
        SQLInsertFields = SQLInsertFields.AddtoDBString("LegBoltsAtTop")
        SQLInsertFields = SQLInsertFields.AddtoDBString("PrintMonopoleAtIncrements")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UseTIA222Exemptions_MinBracingResistance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UseTIA222Exemptions_TensionSplice")
        SQLInsertFields = SQLInsertFields.AddtoDBString("IgnoreKLryFor60DegAngleLegs")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ASCE_7_10_WindData")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ASCE_7_10_ConvertWindToASD")
        SQLInsertFields = SQLInsertFields.AddtoDBString("SolutionUsePDelta")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UseFeedlineTorque")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UsePinnedElements")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UseMaxKz")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UseRigidIndex")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UseTrueCable")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UseASCELy")
        SQLInsertFields = SQLInsertFields.AddtoDBString("CalcBracingForces")
        SQLInsertFields = SQLInsertFields.AddtoDBString("IgnoreBracingFEA")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UseSubCriticalFlow")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AssumePoleWithNoAttachments")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AssumePoleWithShroud")
        SQLInsertFields = SQLInsertFields.AddtoDBString("PoleCornerRadiusKnown")
        SQLInsertFields = SQLInsertFields.AddtoDBString("SolutionMinStiffness")
        SQLInsertFields = SQLInsertFields.AddtoDBString("SolutionMaxStiffness")
        SQLInsertFields = SQLInsertFields.AddtoDBString("SolutionMaxCycles")
        SQLInsertFields = SQLInsertFields.AddtoDBString("SolutionPower")
        SQLInsertFields = SQLInsertFields.AddtoDBString("SolutionTolerance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("CantKFactor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("RadiusSampleDist")
        SQLInsertFields = SQLInsertFields.AddtoDBString("BypassStabilityChecks")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UseWindProjection")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UseIceEscalation")
        SQLInsertFields = SQLInsertFields.AddtoDBString("UseDishCoeff")
        SQLInsertFields = SQLInsertFields.AddtoDBString("AutoCalcTorqArmArea")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDirOption")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir0_0")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir0_1")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir0_2")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir0_3")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir0_4")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir0_5")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir0_6")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir0_7")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir0_8")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir0_9")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir0_10")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir0_11")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir0_12")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir0_13")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir0_14")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir0_15")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir1_0")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir1_1")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir1_2")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir1_3")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir1_4")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir1_5")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir1_6")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir1_7")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir1_8")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir1_9")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir1_10")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir1_11")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir1_12")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir1_13")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir1_14")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir1_15")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir2_0")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir2_1")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir2_2")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir2_3")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir2_4")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir2_5")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir2_6")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir2_7")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir2_8")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir2_9")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir2_10")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir2_11")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir2_12")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir2_13")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir2_14")
        SQLInsertFields = SQLInsertFields.AddtoDBString("WindDir2_15")
        SQLInsertFields = SQLInsertFields.AddtoDBString("SuppressWindPatternLoading")
        SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")

        Return SQLInsertFields
    End Function
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""

        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bus_unit.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.structure_id.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.projectInfo.DesignStandardSeries.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.projectInfo.UnitsSystem.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.projectInfo.ClientName.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.projectInfo.ProjectName.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.projectInfo.ProjectNumber.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.projectInfo.CreatedBy.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.projectInfo.CreatedOn.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.projectInfo.LastUsedBy.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.projectInfo.LastUsedOn.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.projectInfo.VersionUsed.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Length.value.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Length.precision.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Coordinate.value.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Coordinate.precision.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Force.value.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Force.precision.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Load.value.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Load.precision.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Moment.value.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Moment.precision.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Properties.value.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Properties.precision.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Pressure.value.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Pressure.precision.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Velocity.value.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Velocity.precision.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Displacement.value.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Displacement.precision.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Mass.value.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Mass.precision.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Acceleration.value.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Acceleration.precision.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Stress.value.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Stress.precision.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Density.value.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Density.precision.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.UnitWt.value.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.UnitWt.precision.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Strength.value.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Strength.precision.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Modulus.value.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Modulus.precision.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Temperature.value.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Temperature.precision.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Printer.value.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Printer.precision.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Rotation.value.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Rotation.precision.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Spacing.value.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.USUnits.Spacing.precision.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.userInfo.ViewerUserName.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.userInfo.ViewerCompanyName.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.userInfo.ViewerStreetAddress.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.userInfo.ViewerCityState.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.userInfo.ViewerPhone.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.userInfo.ViewerFAX.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.userInfo.ViewerLogo.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.settings.userInfo.ViewerCompanyBitmap.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportProjectNumber.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportJobType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportCarrierName.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportCarrierSiteNumber.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportCarrierSiteName.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportSiteAddress.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportLatitudeDegree.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportLatitudeMinute.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportLatitudeSecond.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportLongitudeDegree.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportLongitudeMinute.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportLongitudeSecond.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportLocalCodeRequirement.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportSiteHistory.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportTowerManufacturer.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportMonthManufactured.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportYearManufactured.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportOriginalSpeed.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportOriginalCode.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportTowerType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportEngrName.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportEngrTitle.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportHQPhoneNumber.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportEmailAddress.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportLogoPath.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportCCiContactName.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportCCiAddress1.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportCCiAddress2.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportCCiBUNumber.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportCCiSiteName.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportCCiJDENumber.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportCCiWONumber.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportCCiPONumber.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportCCiAppNumber.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportCCiRevNumber.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportRecommendations.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportAppurt1Note1.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportAppurt1Note2.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportAppurt1Note3.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportAppurt1Note4.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportAppurt1Note5.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportAppurt1Note6.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportAppurt1Note7.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportAppurt2Note1.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportAppurt2Note2.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportAppurt2Note3.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportAppurt2Note4.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportAppurt2Note5.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportAppurt2Note6.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportAppurt2Note7.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportAddlCapacityNote1.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportAddlCapacityNote2.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportAddlCapacityNote3.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CCIReport.sReportAddlCapacityNote4.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.design.DesignCode.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.geometry.TowerType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.geometry.AntennaType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.geometry.OverallHeight.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.geometry.BaseElevation.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.geometry.Lambda.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.geometry.TowerTopFaceWidth.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.geometry.TowerBaseFaceWidth.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.WindSpeed.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.WindSpeedIce.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.WindSpeedService.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.ice.IceThickness.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.CSA_S37_RefVelPress.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.CSA_S37_ReliabilityClass.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.CSA_S37_ServiceabilityFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.ice.UseModified_TIA_222_IceParameters.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.ice.TIA_222_IceThicknessMultiplier.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.ice.DoNotUse_TIA_222_IceEscalation.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.ice.IceDensity.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.seismic.SeismicSiteClass.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.seismic.SeismicSs.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.seismic.SeismicS1.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.thermal.TempDrop.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.misclCode.GroutFc.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.defaultGirtOffsets.GirtOffset.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.defaultGirtOffsets.GirtOffsetLatticedPole.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.foundationStiffness.MastVert.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.foundationStiffness.MastHorz.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.foundationStiffness.GuyVert.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.foundationStiffness.GuyHorz.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.misclOptions.HogRodTakeup.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.geometry.TowerTaper.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.geometry.GuyedMonopoleBaseType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.geometry.TaperHeight.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.geometry.PivotHeight.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.geometry.AutoCalcGH.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.MTOSettings.IncludeCapacityNote.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.MTOSettings.IncludeAppurtGraphics.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.MTOSettings.DisplayNotes.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.MTOSettings.DisplayReactions.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.MTOSettings.DisplaySchedule.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.MTOSettings.DisplayAppurtenanceTable.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.MTOSettings.DisplayMaterialStrengthTable.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.AutoCalc_ASCE_GH.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.ASCE_ExposureCat.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.ASCE_Year.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.ASCEGh.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.ASCEI.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.UseASCEWind.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.geometry.UserGHElev.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.design.UseCodeGuySF.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.design.GuySF.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.CalcWindAt.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.misclCode.TowerBoltGrade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.misclCode.TowerBoltMinEdgeDist.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.design.AllowStressRatio.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.design.AllowAntStressRatio.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.WindCalcPoints.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.geometry.UseIndexPlate.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.geometry.EnterUserDefinedGhValues.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.geometry.BaseTowerGhInput.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.geometry.UpperStructureGhInput.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.geometry.EnterUserDefinedCgValues.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.geometry.BaseTowerCgInput.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.geometry.UpperStructureCgInput.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.cantileverPoles.CheckVonMises.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.UseClearSpans.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.UseClearSpansKlr.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.geometry.AntennaFaceWidth.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.design.DoInteraction.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.design.DoHorzInteraction.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.design.DoDiagInteraction.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.design.UseMomentMagnification.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.UseFeedlineAsCylinder.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.defaultGirtOffsets.OffsetBotGirt.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.design.PrintBitmaps.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.geometry.UseTopTakeup.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.geometry.ConstantSlope.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.design.UseCodeStressRatio.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.UseLegLoads.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.design.ERIDesignMode.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.WindExposure.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.WindZone.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.StructureCategory.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.RiskCategory.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.TopoCategory.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.RSMTopographicFeature.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.RSM_L.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.RSM_X.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.CrestHeight.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.TIA_222_H_TopoFeatureDownwind.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.BaseElevAboveSeaLevel.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.ConsiderRooftopSpeedUp.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.RooftopWS.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.RooftopHS.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.RooftopParapetHt.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.RooftopXB.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.design.UseTIA222H_AnnexS.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.design.TIA_222_H_AnnexS_Ratio.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.seismic.UseASCE7_10_Seismic_Lcomb.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.EIACWindMult.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.EIACWindMultIce.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.EIACIgnoreCableDrag.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.MTOSettings.Notes.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reportSettings.ReportInputCosts.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reportSettings.ReportInputGeometry.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reportSettings.ReportInputOptions.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reportSettings.ReportMaxForces.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reportSettings.ReportInputMap.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reportSettings.CostReportOutputType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reportSettings.CapacityReportOutputType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reportSettings.ReportPrintForceTotals.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reportSettings.ReportPrintForceDetails.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reportSettings.ReportPrintMastVectors.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reportSettings.ReportPrintAntPoleVectors.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reportSettings.ReportPrintDiscreteVectors.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reportSettings.ReportPrintDishVectors.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reportSettings.ReportPrintFeedTowerVectors.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reportSettings.ReportPrintUserLoadVectors.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reportSettings.ReportPrintPressures.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reportSettings.ReportPrintAppurtForces.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reportSettings.ReportPrintGuyForces.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reportSettings.ReportPrintGuyStressing.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reportSettings.ReportPrintDeflections.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reportSettings.ReportPrintReactions.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reportSettings.ReportPrintStressChecks.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reportSettings.ReportPrintBoltChecks.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reportSettings.ReportPrintInputGVerificationTables.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reportSettings.ReportPrintOutputGVerificationTables.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.cantileverPoles.SocketTopMount.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.SRTakeCompression.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.AllLegPanelsSame.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.UseCombinedBoltCapacity.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.SecHorzBracesLeg.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.SortByComponent.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.SRCutEnds.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.SRConcentric.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.CalcBlockShear.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.Use4SidedDiamondBracing.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.TriangulateInnerBracing.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.PrintCarrierNotes.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.AddIBCWindCase.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.UseStateCountyLookup.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.State.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.County.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.LegBoltsAtTop.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.cantileverPoles.PrintMonopoleAtIncrements.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.UseTIA222Exemptions_MinBracingResistance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.UseTIA222Exemptions_TensionSplice.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.IgnoreKLryFor60DegAngleLegs.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.ASCE_7_10_WindData.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.ASCE_7_10_ConvertWindToASD.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.solutionSettings.SolutionUsePDelta.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.UseFeedlineTorque.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.UsePinnedElements.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.wind.UseMaxKz.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.UseRigidIndex.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.UseTrueCable.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.UseASCELy.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.CalcBracingForces.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.IgnoreBracingFEA.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.cantileverPoles.UseSubCriticalFlow.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.cantileverPoles.AssumePoleWithNoAttachments.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.cantileverPoles.AssumePoleWithShroud.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.cantileverPoles.PoleCornerRadiusKnown.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.solutionSettings.SolutionMinStiffness.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.solutionSettings.SolutionMaxStiffness.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.solutionSettings.SolutionMaxCycles.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.solutionSettings.SolutionPower.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.solutionSettings.SolutionTolerance.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.cantileverPoles.CantKFactor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.misclOptions.RadiusSampleDist.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.BypassStabilityChecks.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.UseWindProjection.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.code.ice.UseIceEscalation.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.UseDishCoeff.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.AutoCalcTorqArmArea.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDirOption.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir0_0.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir0_1.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir0_2.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir0_3.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir0_4.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir0_5.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir0_6.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir0_7.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir0_8.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir0_9.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir0_10.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir0_11.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir0_12.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir0_13.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir0_14.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir0_15.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir1_0.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir1_1.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir1_2.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir1_3.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir1_4.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir1_5.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir1_6.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir1_7.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir1_8.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir1_9.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir1_10.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir1_11.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir1_12.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir1_13.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir1_14.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir1_15.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir2_0.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir2_1.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir2_2.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir2_3.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir2_4.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir2_5.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir2_6.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir2_7.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir2_8.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir2_9.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir2_10.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir2_11.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir2_12.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir2_13.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir2_14.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.WindDir2_15.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.options.windDirections.SuppressWindPatternLoading.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function
    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""

        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bus_unit = " & Me.bus_unit.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("structure_id = " & Me.structure_id.NullableToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("DesignStandardSeries = " & Me.settings.projectInfo.DesignStandardSeries.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("UnitsSystem = " & Me.settings.projectInfo.UnitsSystem.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ClientName = " & Me.settings.projectInfo.ClientName.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ProjectName = " & Me.settings.projectInfo.ProjectName.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ProjectNumber = " & Me.settings.projectInfo.ProjectNumber.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("CreatedBy = " & Me.settings.projectInfo.CreatedBy.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("CreatedOn = " & Me.settings.projectInfo.CreatedOn.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("LastUsedBy = " & Me.settings.projectInfo.LastUsedBy.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("LastUsedOn = " & Me.settings.projectInfo.LastUsedOn.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("VersionUsed = " & Me.settings.projectInfo.VersionUsed.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USLength = " & Me.settings.USUnits.Length.value.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USLengthPrec = " & Me.settings.USUnits.Length.precision.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USCoordinate = " & Me.settings.USUnits.Coordinate.value.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USCoordinatePrec = " & Me.settings.USUnits.Coordinate.precision.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USForce = " & Me.settings.USUnits.Force.value.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USForcePrec = " & Me.settings.USUnits.Force.precision.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USLoad = " & Me.settings.USUnits.Load.value.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USLoadPrec = " & Me.settings.USUnits.Load.precision.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USMoment = " & Me.settings.USUnits.Moment.value.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USMomentPrec = " & Me.settings.USUnits.Moment.precision.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USProperties = " & Me.settings.USUnits.Properties.value.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USPropertiesPrec = " & Me.settings.USUnits.Properties.precision.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USPressure = " & Me.settings.USUnits.Pressure.value.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USPressurePrec = " & Me.settings.USUnits.Pressure.precision.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USVelocity = " & Me.settings.USUnits.Velocity.value.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USVelocityPrec = " & Me.settings.USUnits.Velocity.precision.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USDisplacement = " & Me.settings.USUnits.Displacement.value.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USDisplacementPrec = " & Me.settings.USUnits.Displacement.precision.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USMass = " & Me.settings.USUnits.Mass.value.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USMassPrec = " & Me.settings.USUnits.Mass.precision.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USAcceleration = " & Me.settings.USUnits.Acceleration.value.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USAccelerationPrec = " & Me.settings.USUnits.Acceleration.precision.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USStress = " & Me.settings.USUnits.Stress.value.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USStressPrec = " & Me.settings.USUnits.Stress.precision.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USDensity = " & Me.settings.USUnits.Density.value.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USDensityPrec = " & Me.settings.USUnits.Density.precision.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USUnitWt = " & Me.settings.USUnits.UnitWt.value.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USUnitWtPrec = " & Me.settings.USUnits.UnitWt.precision.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USStrength = " & Me.settings.USUnits.Strength.value.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USStrengthPrec = " & Me.settings.USUnits.Strength.precision.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USModulus = " & Me.settings.USUnits.Modulus.value.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USModulusPrec = " & Me.settings.USUnits.Modulus.precision.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USTemperature = " & Me.settings.USUnits.Temperature.value.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USTemperaturePrec = " & Me.settings.USUnits.Temperature.precision.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USPrinter = " & Me.settings.USUnits.Printer.value.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USPrinterPrec = " & Me.settings.USUnits.Printer.precision.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USRotation = " & Me.settings.USUnits.Rotation.value.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USRotationPrec = " & Me.settings.USUnits.Rotation.precision.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USSpacing = " & Me.settings.USUnits.Spacing.value.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("USSpacingPrec = " & Me.settings.USUnits.Spacing.precision.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ViewerUserName = " & Me.settings.userInfo.ViewerUserName.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ViewerCompanyName = " & Me.settings.userInfo.ViewerCompanyName.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ViewerStreetAddress = " & Me.settings.userInfo.ViewerStreetAddress.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ViewerCityState = " & Me.settings.userInfo.ViewerCityState.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ViewerPhone = " & Me.settings.userInfo.ViewerPhone.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ViewerFAX = " & Me.settings.userInfo.ViewerFAX.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ViewerLogo = " & Me.settings.userInfo.ViewerLogo.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ViewerCompanyBitmap = " & Me.settings.userInfo.ViewerCompanyBitmap.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportProjectNumber = " & Me.CCIReport.sReportProjectNumber.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportJobType = " & Me.CCIReport.sReportJobType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportCarrierName = " & Me.CCIReport.sReportCarrierName.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportCarrierSiteNumber = " & Me.CCIReport.sReportCarrierSiteNumber.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportCarrierSiteName = " & Me.CCIReport.sReportCarrierSiteName.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportSiteAddress = " & Me.CCIReport.sReportSiteAddress.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportLatitudeDegree = " & Me.CCIReport.sReportLatitudeDegree.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportLatitudeMinute = " & Me.CCIReport.sReportLatitudeMinute.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportLatitudeSecond = " & Me.CCIReport.sReportLatitudeSecond.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportLongitudeDegree = " & Me.CCIReport.sReportLongitudeDegree.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportLongitudeMinute = " & Me.CCIReport.sReportLongitudeMinute.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportLongitudeSecond = " & Me.CCIReport.sReportLongitudeSecond.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportLocalCodeRequirement = " & Me.CCIReport.sReportLocalCodeRequirement.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportSiteHistory = " & Me.CCIReport.sReportSiteHistory.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportTowerManufacturer = " & Me.CCIReport.sReportTowerManufacturer.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportMonthManufactured = " & Me.CCIReport.sReportMonthManufactured.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportYearManufactured = " & Me.CCIReport.sReportYearManufactured.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportOriginalSpeed = " & Me.CCIReport.sReportOriginalSpeed.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportOriginalCode = " & Me.CCIReport.sReportOriginalCode.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportTowerType = " & Me.CCIReport.sReportTowerType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportEngrName = " & Me.CCIReport.sReportEngrName.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportEngrTitle = " & Me.CCIReport.sReportEngrTitle.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportHQPhoneNumber = " & Me.CCIReport.sReportHQPhoneNumber.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportEmailAddress = " & Me.CCIReport.sReportEmailAddress.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportLogoPath = " & Me.CCIReport.sReportLogoPath.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportCCiContactName = " & Me.CCIReport.sReportCCiContactName.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportCCiAddress1 = " & Me.CCIReport.sReportCCiAddress1.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportCCiAddress2 = " & Me.CCIReport.sReportCCiAddress2.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportCCiBUNumber = " & Me.CCIReport.sReportCCiBUNumber.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportCCiSiteName = " & Me.CCIReport.sReportCCiSiteName.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportCCiJDENumber = " & Me.CCIReport.sReportCCiJDENumber.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportCCiWONumber = " & Me.CCIReport.sReportCCiWONumber.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportCCiPONumber = " & Me.CCIReport.sReportCCiPONumber.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportCCiAppNumber = " & Me.CCIReport.sReportCCiAppNumber.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportCCiRevNumber = " & Me.CCIReport.sReportCCiRevNumber.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportRecommendations = " & Me.CCIReport.sReportRecommendations.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportAppurt1Note1 = " & Me.CCIReport.sReportAppurt1Note1.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportAppurt1Note2 = " & Me.CCIReport.sReportAppurt1Note2.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportAppurt1Note3 = " & Me.CCIReport.sReportAppurt1Note3.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportAppurt1Note4 = " & Me.CCIReport.sReportAppurt1Note4.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportAppurt1Note5 = " & Me.CCIReport.sReportAppurt1Note5.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportAppurt1Note6 = " & Me.CCIReport.sReportAppurt1Note6.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportAppurt1Note7 = " & Me.CCIReport.sReportAppurt1Note7.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportAppurt2Note1 = " & Me.CCIReport.sReportAppurt2Note1.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportAppurt2Note2 = " & Me.CCIReport.sReportAppurt2Note2.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportAppurt2Note3 = " & Me.CCIReport.sReportAppurt2Note3.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportAppurt2Note4 = " & Me.CCIReport.sReportAppurt2Note4.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportAppurt2Note5 = " & Me.CCIReport.sReportAppurt2Note5.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportAppurt2Note6 = " & Me.CCIReport.sReportAppurt2Note6.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportAppurt2Note7 = " & Me.CCIReport.sReportAppurt2Note7.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportAddlCapacityNote1 = " & Me.CCIReport.sReportAddlCapacityNote1.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportAddlCapacityNote2 = " & Me.CCIReport.sReportAddlCapacityNote2.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportAddlCapacityNote3 = " & Me.CCIReport.sReportAddlCapacityNote3.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sReportAddlCapacityNote4 = " & Me.CCIReport.sReportAddlCapacityNote4.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("DesignCode = " & Me.code.design.DesignCode.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerType = " & Me.geometry.TowerType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaType = " & Me.geometry.AntennaType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("OverallHeight = " & Me.geometry.OverallHeight.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("BaseElevation = " & Me.geometry.BaseElevation.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("Lambda = " & Me.geometry.Lambda.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerTopFaceWidth = " & Me.geometry.TowerTopFaceWidth.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerBaseFaceWidth = " & Me.geometry.TowerBaseFaceWidth.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindSpeed = " & Me.code.wind.WindSpeed.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindSpeedIce = " & Me.code.wind.WindSpeedIce.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindSpeedService = " & Me.code.wind.WindSpeedService.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("IceThickness = " & Me.code.ice.IceThickness.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("CSA_S37_RefVelPress = " & Me.code.wind.CSA_S37_RefVelPress.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("CSA_S37_ReliabilityClass = " & Me.code.wind.CSA_S37_ReliabilityClass.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("CSA_S37_ServiceabilityFactor = " & Me.code.wind.CSA_S37_ServiceabilityFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("UseModified_TIA_222_IceParameters = " & Me.code.ice.UseModified_TIA_222_IceParameters.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TIA_222_IceThicknessMultiplier = " & Me.code.ice.TIA_222_IceThicknessMultiplier.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("DoNotUse_TIA_222_IceEscalation = " & Me.code.ice.DoNotUse_TIA_222_IceEscalation.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("IceDensity = " & Me.code.ice.IceDensity.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("SeismicSiteClass = " & Me.code.seismic.SeismicSiteClass.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("SeismicSs = " & Me.code.seismic.SeismicSs.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("SeismicS1 = " & Me.code.seismic.SeismicS1.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TempDrop = " & Me.code.thermal.TempDrop.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GroutFc = " & Me.code.misclCode.GroutFc.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GirtOffset = " & Me.options.defaultGirtOffsets.GirtOffset.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GirtOffsetLatticedPole = " & Me.options.defaultGirtOffsets.GirtOffsetLatticedPole.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("MastVert = " & Me.options.foundationStiffness.MastVert.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("MastHorz = " & Me.options.foundationStiffness.MastHorz.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyVert = " & Me.options.foundationStiffness.GuyVert.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyHorz = " & Me.options.foundationStiffness.GuyHorz.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("HogRodTakeup = " & Me.options.misclOptions.HogRodTakeup.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerTaper = " & Me.geometry.TowerTaper.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuyedMonopoleBaseType = " & Me.geometry.GuyedMonopoleBaseType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TaperHeight = " & Me.geometry.TaperHeight.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("PivotHeight = " & Me.geometry.PivotHeight.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AutoCalcGH = " & Me.geometry.AutoCalcGH.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("IncludeCapacityNote = " & Me.MTOSettings.IncludeCapacityNote.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("IncludeAppurtGraphics = " & Me.MTOSettings.IncludeAppurtGraphics.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("DisplayNotes = " & Me.MTOSettings.DisplayNotes.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("DisplayReactions = " & Me.MTOSettings.DisplayReactions.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("DisplaySchedule = " & Me.MTOSettings.DisplaySchedule.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("DisplayAppurtenanceTable = " & Me.MTOSettings.DisplayAppurtenanceTable.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("DisplayMaterialStrengthTable = " & Me.MTOSettings.DisplayMaterialStrengthTable.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AutoCalc_ASCE_GH = " & Me.code.wind.AutoCalc_ASCE_GH.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ASCE_ExposureCat = " & Me.code.wind.ASCE_ExposureCat.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ASCE_Year = " & Me.code.wind.ASCE_Year.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ASCEGh = " & Me.code.wind.ASCEGh.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ASCEI = " & Me.code.wind.ASCEI.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("UseASCEWind = " & Me.code.wind.UseASCEWind.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("UserGHElev = " & Me.geometry.UserGHElev.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("UseCodeGuySF = " & Me.code.design.UseCodeGuySF.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("GuySF = " & Me.code.design.GuySF.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("CalcWindAt = " & Me.code.wind.CalcWindAt.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerBoltGrade = " & Me.code.misclCode.TowerBoltGrade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TowerBoltMinEdgeDist = " & Me.code.misclCode.TowerBoltMinEdgeDist.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AllowStressRatio = " & Me.code.design.AllowStressRatio.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AllowAntStressRatio = " & Me.code.design.AllowAntStressRatio.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindCalcPoints = " & Me.code.wind.WindCalcPoints.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("UseIndexPlate = " & Me.geometry.UseIndexPlate.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("EnterUserDefinedGhValues = " & Me.geometry.EnterUserDefinedGhValues.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("BaseTowerGhInput = " & Me.geometry.BaseTowerGhInput.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("UpperStructureGhInput = " & Me.geometry.UpperStructureGhInput.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("EnterUserDefinedCgValues = " & Me.geometry.EnterUserDefinedCgValues.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("BaseTowerCgInput = " & Me.geometry.BaseTowerCgInput.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("UpperStructureCgInput = " & Me.geometry.UpperStructureCgInput.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("CheckVonMises = " & Me.options.cantileverPoles.CheckVonMises.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("UseClearSpans = " & Me.options.UseClearSpans.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("UseClearSpansKlr = " & Me.options.UseClearSpansKlr.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AntennaFaceWidth = " & Me.geometry.AntennaFaceWidth.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("DoInteraction = " & Me.code.design.DoInteraction.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("DoHorzInteraction = " & Me.code.design.DoHorzInteraction.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("DoDiagInteraction = " & Me.code.design.DoDiagInteraction.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("UseMomentMagnification = " & Me.code.design.UseMomentMagnification.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("UseFeedlineAsCylinder = " & Me.options.UseFeedlineAsCylinder.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("OffsetBotGirt = " & Me.options.defaultGirtOffsets.OffsetBotGirt.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("PrintBitmaps = " & Me.code.design.PrintBitmaps.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("UseTopTakeup = " & Me.geometry.UseTopTakeup.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ConstantSlope = " & Me.geometry.ConstantSlope.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("UseCodeStressRatio = " & Me.code.design.UseCodeStressRatio.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("UseLegLoads = " & Me.options.UseLegLoads.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ERIDesignMode = " & Me.code.design.ERIDesignMode.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindExposure = " & Me.code.wind.WindExposure.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindZone = " & Me.code.wind.WindZone.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("StructureCategory = " & Me.code.wind.StructureCategory.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("RiskCategory = " & Me.code.wind.RiskCategory.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TopoCategory = " & Me.code.wind.TopoCategory.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("RSMTopographicFeature = " & Me.code.wind.RSMTopographicFeature.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("RSM_L = " & Me.code.wind.RSM_L.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("RSM_X = " & Me.code.wind.RSM_X.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("CrestHeight = " & Me.code.wind.CrestHeight.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TIA_222_H_TopoFeatureDownwind = " & Me.code.wind.TIA_222_H_TopoFeatureDownwind.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("BaseElevAboveSeaLevel = " & Me.code.wind.BaseElevAboveSeaLevel.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ConsiderRooftopSpeedUp = " & Me.code.wind.ConsiderRooftopSpeedUp.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("RooftopWS = " & Me.code.wind.RooftopWS.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("RooftopHS = " & Me.code.wind.RooftopHS.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("RooftopParapetHt = " & Me.code.wind.RooftopParapetHt.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("RooftopXB = " & Me.code.wind.RooftopXB.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("UseTIA222H_AnnexS = " & Me.code.design.UseTIA222H_AnnexS.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TIA_222_H_AnnexS_Ratio = " & Me.code.design.TIA_222_H_AnnexS_Ratio.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("UseASCE7_10_Seismic_LComb = " & Me.code.seismic.UseASCE7_10_Seismic_Lcomb.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("EIACWindMult = " & Me.code.wind.EIACWindMult.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("EIACWindMultIce = " & Me.code.wind.EIACWindMultIce.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("EIACIgnoreCableDrag = " & Me.code.wind.EIACIgnoreCableDrag.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("Notes = " & Me.MTOSettings.Notes.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ReportInputCosts = " & Me.reportSettings.ReportInputCosts.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ReportInputGeometry = " & Me.reportSettings.ReportInputGeometry.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ReportInputOptions = " & Me.reportSettings.ReportInputOptions.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ReportMaxForces = " & Me.reportSettings.ReportMaxForces.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ReportInputMap = " & Me.reportSettings.ReportInputMap.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("CostReportOutputType = " & Me.reportSettings.CostReportOutputType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("CapacityReportOutputType = " & Me.reportSettings.CapacityReportOutputType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ReportPrintForceTotals = " & Me.reportSettings.ReportPrintForceTotals.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ReportPrintForceDetails = " & Me.reportSettings.ReportPrintForceDetails.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ReportPrintMastVectors = " & Me.reportSettings.ReportPrintMastVectors.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ReportPrintAntPoleVectors = " & Me.reportSettings.ReportPrintAntPoleVectors.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ReportPrintDiscreteVectors = " & Me.reportSettings.ReportPrintDiscreteVectors.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ReportPrintDishVectors = " & Me.reportSettings.ReportPrintDishVectors.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ReportPrintFeedTowerVectors = " & Me.reportSettings.ReportPrintFeedTowerVectors.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ReportPrintUserLoadVectors = " & Me.reportSettings.ReportPrintUserLoadVectors.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ReportPrintPressures = " & Me.reportSettings.ReportPrintPressures.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ReportPrintAppurtForces = " & Me.reportSettings.ReportPrintAppurtForces.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ReportPrintGuyForces = " & Me.reportSettings.ReportPrintGuyForces.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ReportPrintGuyStressing = " & Me.reportSettings.ReportPrintGuyStressing.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ReportPrintDeflections = " & Me.reportSettings.ReportPrintDeflections.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ReportPrintReactions = " & Me.reportSettings.ReportPrintReactions.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ReportPrintStressChecks = " & Me.reportSettings.ReportPrintStressChecks.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ReportPrintBoltChecks = " & Me.reportSettings.ReportPrintBoltChecks.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ReportPrintInputGVerificationTables = " & Me.reportSettings.ReportPrintInputGVerificationTables.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ReportPrintOutputGVerificationTables = " & Me.reportSettings.ReportPrintOutputGVerificationTables.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("SocketTopMount = " & Me.options.cantileverPoles.SocketTopMount.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("SRTakeCompression = " & Me.options.SRTakeCompression.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AllLegPanelsSame = " & Me.options.AllLegPanelsSame.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("UseCombinedBoltCapacity = " & Me.options.UseCombinedBoltCapacity.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("SecHorzBracesLeg = " & Me.options.SecHorzBracesLeg.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("SortByComponent = " & Me.options.SortByComponent.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("SRCutEnds = " & Me.options.SRCutEnds.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("SRConcentric = " & Me.options.SRConcentric.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("CalcBlockShear = " & Me.options.CalcBlockShear.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("Use4SidedDiamondBracing = " & Me.options.Use4SidedDiamondBracing.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("TriangulateInnerBracing = " & Me.options.TriangulateInnerBracing.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("PrintCarrierNotes = " & Me.options.PrintCarrierNotes.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AddIBCWindCase = " & Me.options.AddIBCWindCase.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("UseStateCountyLookup = " & Me.code.wind.UseStateCountyLookup.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("State = " & Me.code.wind.State.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("County = " & Me.code.wind.County.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("LegBoltsAtTop = " & Me.options.LegBoltsAtTop.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("PrintMonopoleAtIncrements = " & Me.options.cantileverPoles.PrintMonopoleAtIncrements.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("UseTIA222Exemptions_MinBracingResistance = " & Me.options.UseTIA222Exemptions_MinBracingResistance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("UseTIA222Exemptions_TensionSplice = " & Me.options.UseTIA222Exemptions_TensionSplice.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("IgnoreKLryFor60DegAngleLegs = " & Me.options.IgnoreKLryFor60DegAngleLegs.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ASCE_7_10_WindData = " & Me.code.wind.ASCE_7_10_WindData.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ASCE_7_10_ConvertWindToASD = " & Me.code.wind.ASCE_7_10_ConvertWindToASD.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("SolutionUsePDelta = " & Me.solutionSettings.SolutionUsePDelta.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("UseFeedlineTorque = " & Me.options.UseFeedlineTorque.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("UsePinnedElements = " & Me.options.UsePinnedElements.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("UseMaxKz = " & Me.code.wind.UseMaxKz.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("UseRigidIndex = " & Me.options.UseRigidIndex.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("UseTrueCable = " & Me.options.UseTrueCable.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("UseASCELy = " & Me.options.UseASCELy.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("CalcBracingForces = " & Me.options.CalcBracingForces.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("IgnoreBracingFEA = " & Me.options.IgnoreBracingFEA.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("UseSubCriticalFlow = " & Me.options.cantileverPoles.UseSubCriticalFlow.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AssumePoleWithNoAttachments = " & Me.options.cantileverPoles.AssumePoleWithNoAttachments.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AssumePoleWithShroud = " & Me.options.cantileverPoles.AssumePoleWithShroud.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("PoleCornerRadiusKnown = " & Me.options.cantileverPoles.PoleCornerRadiusKnown.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("SolutionMinStiffness = " & Me.solutionSettings.SolutionMinStiffness.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("SolutionMaxStiffness = " & Me.solutionSettings.SolutionMaxStiffness.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("SolutionMaxCycles = " & Me.solutionSettings.SolutionMaxCycles.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("SolutionPower = " & Me.solutionSettings.SolutionPower.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("SolutionTolerance = " & Me.solutionSettings.SolutionTolerance.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("CantKFactor = " & Me.options.cantileverPoles.CantKFactor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("RadiusSampleDist = " & Me.options.misclOptions.RadiusSampleDist.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("BypassStabilityChecks = " & Me.options.BypassStabilityChecks.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("UseWindProjection = " & Me.options.UseWindProjection.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("UseIceEscalation = " & Me.code.ice.UseIceEscalation.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("UseDishCoeff = " & Me.options.UseDishCoeff.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("AutoCalcTorqArmArea = " & Me.options.AutoCalcTorqArmArea.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDirOption = " & Me.options.windDirections.WindDirOption.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir0_0 = " & Me.options.windDirections.WindDir0_0.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir0_1 = " & Me.options.windDirections.WindDir0_1.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir0_2 = " & Me.options.windDirections.WindDir0_2.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir0_3 = " & Me.options.windDirections.WindDir0_3.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir0_4 = " & Me.options.windDirections.WindDir0_4.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir0_5 = " & Me.options.windDirections.WindDir0_5.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir0_6 = " & Me.options.windDirections.WindDir0_6.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir0_7 = " & Me.options.windDirections.WindDir0_7.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir0_8 = " & Me.options.windDirections.WindDir0_8.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir0_9 = " & Me.options.windDirections.WindDir0_9.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir0_10 = " & Me.options.windDirections.WindDir0_10.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir0_11 = " & Me.options.windDirections.WindDir0_11.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir0_12 = " & Me.options.windDirections.WindDir0_12.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir0_13 = " & Me.options.windDirections.WindDir0_13.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir0_14 = " & Me.options.windDirections.WindDir0_14.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir0_15 = " & Me.options.windDirections.WindDir0_15.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir1_0 = " & Me.options.windDirections.WindDir1_0.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir1_1 = " & Me.options.windDirections.WindDir1_1.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir1_2 = " & Me.options.windDirections.WindDir1_2.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir1_3 = " & Me.options.windDirections.WindDir1_3.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir1_4 = " & Me.options.windDirections.WindDir1_4.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir1_5 = " & Me.options.windDirections.WindDir1_5.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir1_6 = " & Me.options.windDirections.WindDir1_6.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir1_7 = " & Me.options.windDirections.WindDir1_7.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir1_8 = " & Me.options.windDirections.WindDir1_8.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir1_9 = " & Me.options.windDirections.WindDir1_9.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir1_10 = " & Me.options.windDirections.WindDir1_10.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir1_11 = " & Me.options.windDirections.WindDir1_11.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir1_12 = " & Me.options.windDirections.WindDir1_12.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir1_13 = " & Me.options.windDirections.WindDir1_13.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir1_14 = " & Me.options.windDirections.WindDir1_14.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir1_15 = " & Me.options.windDirections.WindDir1_15.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir2_0 = " & Me.options.windDirections.WindDir2_0.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir2_1 = " & Me.options.windDirections.WindDir2_1.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir2_2 = " & Me.options.windDirections.WindDir2_2.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir2_3 = " & Me.options.windDirections.WindDir2_3.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir2_4 = " & Me.options.windDirections.WindDir2_4.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir2_5 = " & Me.options.windDirections.WindDir2_5.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir2_6 = " & Me.options.windDirections.WindDir2_6.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir2_7 = " & Me.options.windDirections.WindDir2_7.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir2_8 = " & Me.options.windDirections.WindDir2_8.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir2_9 = " & Me.options.windDirections.WindDir2_9.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir2_10 = " & Me.options.windDirections.WindDir2_10.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir2_11 = " & Me.options.windDirections.WindDir2_11.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir2_12 = " & Me.options.windDirections.WindDir2_12.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir2_13 = " & Me.options.windDirections.WindDir2_13.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir2_14 = " & Me.options.windDirections.WindDir2_14.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("WindDir2_15 = " & Me.options.windDirections.WindDir2_15.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("SuppressWindPatternLoading = " & Me.options.windDirections.SuppressWindPatternLoading.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_person_id = " & Me.modified_person_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("process_stage = " & Me.process_stage.ToString.FormatDBValue)

        Return SQLUpdateFieldsandValues
    End Function

#End Region

#Region "Save to TNX"

    Public Sub GenerateERI(overwriteFile As OverwriteFile, FilePath As String)
        If File.Exists(FilePath) AndAlso Not overwriteFile(FilePath) Then Exit Sub
        GenerateERI(FilePath)
    End Sub

    Public Sub GenerateERI(FilePath As String, Optional replaceFile As Boolean = True)

        If File.Exists(FilePath) And Not replaceFile Then Exit Sub

        Dim newERIList As New List(Of String)
        Dim i As Integer
        Dim defaults As New DefaultERItxtValues

        'ReportInputCosts is first line after userforces
#Region "Individual fields after User Forces"
        '''''''        ReportInputCosts = No
        '''''''        ReportInputGeometry = Yes
        '''''''        ReportInputOptions = Yes
        '''''''        ReportMaxForces = Yes
        '''''''        ReportInputMap = No
        '''''''        CostReportOutputType = No Cost Output
        '''''''        CapacityReportOutputType = Capacity Summary
        '''''''        ReportPrintForceTotals = No
        '''''''        ReportPrintForceDetails = No
        '''''''        ReportPrintMastVectors = No
        '''''''        ReportPrintAntPoleVectors = No
        '''''''        ReportPrintDiscreteVectors = No
        '''''''        ReportPrintDishVectors = No
        '''''''        ReportPrintFeedTowerVectors = No
        '''''''        ReportPrintUserLoadVectors = No
        '''''''        ReportPrintPressures = No
        '''''''        ReportPrintAppurtForces = No
        '''''''        ReportPrintGuyForces = No
        '''''''        ReportPrintGuyStressing = No
        '''''''        ReportPrintDeflections = Yes
        '''''''        ReportPrintReactions = Yes
        '''''''        ReportPrintStressChecks = Yes
        '''''''        ReportPrintBoltChecks = No
        '''''''        ReportPrintInputGVerificationTables = No
        '''''''        ReportPrintOutputGVerificationTables = No
        '''''''        MastMultiplier = 1.0
        '''''''        SocketTopMount = No
        '''''''        SRTakeCompression = No
        '''''''        AllLegPanelsSame = No
        '''''''        UseCombinedBoltCapacity = No
        '''''''        SecHorzBracesLeg = No
        '''''''        SortByComponent = Yes
        '''''''        SRCutEnds = No
        '''''''        SRConcentric = No
        '''''''        CalcBlockShear = No
        '''''''        Use4SidedDiamondBracing = No
        '''''''        TriangulateInnerBracing = No
        '''''''        PrintCarrierNotes = No
        '''''''        AddIBCWindCase = No
        '''''''        UseStateCountyLookup = Yes
        '''''''        State = Ohio
        '''''''        County = Franklin
        '''''''        EnforceAllOffsetsVer2 = Yes
        '''''''        LegBoltsAtTop = No
        '''''''        PrintMonopoleAtIncrements = No
        '''''''        UseTIA222Exemptions_MinBracingResistance = No
        '''''''        UseTIA222Exemptions_TensionSplice = No
        '''''''        IgnoreKLryFor60DegAngleLegs = No
        '''''''        ASCE_7_10_WindData = No
        '''''''        ASCE_7_10_ConvertWindToASD = No
        '''''''        SolutionUsePDelta = Yes
        '''''''        UseFeedlineTorque = Yes
        '''''''        UsePinnedElements = No
        '''''''        UseMaxKz = No
        '''''''        UseRigidIndex = Yes
        '''''''        UseTrueCable = No
        '''''''        UseASCELy = No
        '''''''        CalcBracingForces = No
        '''''''        IgnoreBracingFEA = No
        '''''''        UseSubCriticalFlow = No
        '''''''        AssumePoleWithNoAttachments = No
        '''''''        AssumePoleWithShroud = No
        '''''''        PoleCornerRadiusKnown = No
        '''''''        SolutionMinStiffness = 0.000000
        '''''''        SolutionMaxStiffness = 0.000000
        '''''''        SolutionMaxCycles = 200
        '''''''        SolutionPower = 0.000000
        '''''''        SolutionTolerance = 0.001
        '''''''        CantKFactor = 1.0
        '''''''        RadiusSampleDist = 5.0
        '''''''        BypassStabilityChecks = Yes
        '''''''        UseWindProjection = Yes
        '''''''        UseIceEscalation = No
        '''''''        UseDishCoeff = Yes
        '''''''        UseMastResponse = No
        '''''''        AutoCalcTorqArmArea = No
        '''''''        Use_TIA_222_G_Addendum2_MP = Yes
        '''''''        WindDirOption = 1
        '''''''        WindDir0_0 = Yes
        '''''''        WindDir0_1 = Yes
        '''''''        WindDir0_2 = No
        '''''''        WindDir0_3 = Yes
        '''''''        WindDir0_4 = Yes
        '''''''        WindDir0_5 = Yes
        '''''''        WindDir0_6 = No
        '''''''        WindDir0_7 = Yes
        '''''''        WindDir0_8 = Yes
        '''''''        WindDir0_9 = Yes
        '''''''        WindDir0_10 = No
        '''''''        WindDir0_11 = Yes
        '''''''        WindDir0_12 = Yes
        '''''''        WindDir0_13 = Yes
        '''''''        WindDir0_14 = No
        '''''''        WindDir0_15 = Yes
        '''''''        WindDir1_0 = Yes
        '''''''        WindDir1_1 = Yes
        '''''''        WindDir1_2 = No
        '''''''        WindDir1_3 = Yes
        '''''''        WindDir1_4 = Yes
        '''''''        WindDir1_5 = Yes
        '''''''        WindDir1_6 = No
        '''''''        WindDir1_7 = Yes
        '''''''        WindDir1_8 = Yes
        '''''''        WindDir1_9 = Yes
        '''''''        WindDir1_10 = No
        '''''''        WindDir1_11 = Yes
        '''''''        WindDir1_12 = Yes
        '''''''        WindDir1_13 = Yes
        '''''''        WindDir1_14 = No
        '''''''        WindDir1_15 = Yes
        '''''''        WindDir2_0 = Yes
        '''''''        WindDir2_1 = Yes
        '''''''        WindDir2_2 = No
        '''''''        WindDir2_3 = Yes
        '''''''        WindDir2_4 = Yes
        '''''''        WindDir2_5 = Yes
        '''''''        WindDir2_6 = No
        '''''''        WindDir2_7 = Yes
        '''''''        WindDir2_8 = Yes
        '''''''        WindDir2_9 = Yes
        '''''''        WindDir2_10 = No
        '''''''        WindDir2_11 = Yes
        '''''''        WindDir2_12 = Yes
        '''''''        WindDir2_13 = Yes
        '''''''        WindDir2_14 = No
        '''''''        WindDir2_15 = Yes
        '''''''        SuppressWindPatternLoading = No
        '''''''        DoELCAnalysis = No
#End Region

        If Me.otherLines.Count < 1 Then
            Me.otherLines.Add(New String() {"[Common]", ""})
            Me.otherLines.Add(New String() {"[US Units]", ""})
            Me.otherLines.Add(New String() {"[SI Units]", ""}) 'Missing
            Me.otherLines.Add(New String() {"[Codes]", ""}) 'Missing
            Me.otherLines.Add(New String() {"[Application]", ""}) 'Missing
            Me.otherLines.Add(New String() {"[Databases]", ""}) 'Missing (Inside of Application
            Me.otherLines.Add(New String() {"[Structure]", ""})
            Me.otherLines.Add(New String() {"NumAntennaRecs", ""})
            Me.otherLines.Add(New String() {"NumTowerRecs", ""})
            Me.otherLines.Add(New String() {"NumGuyRecs", ""})
            Me.otherLines.Add(New String() {"NumFeedLineRecs", ""})
            Me.otherLines.Add(New String() {"NumTowerLoadRecs", ""})
            Me.otherLines.Add(New String() {"NumDishRecs", ""})
            Me.otherLines.Add(New String() {"NumUserForceRecs", ""})
            Me.otherLines.Add(New String() {"[End Structure]", ""}) 'Missing
            Me.otherLines.Add(New String() {"[End Application]", ""}) 'Missing
            Me.otherLines.Add(New String() {"[CHRONOS]", ""}) 'Missing
            Me.otherLines.Add(New String() {"[Nodes]", ""}) 'Missing + comments in section
            Me.otherLines.Add(New String() {"[Supports]", ""}) 'Missing + comments in section
            Me.otherLines.Add(New String() {"[ConnectionArray]", ""}) 'Missing + comments in section
            Me.otherLines.Add(New String() {"[Elements]", ""}) 'Missing + comments in section
            Me.otherLines.Add(New String() {"[Conditions]", ""}) 'Missing + comments in section
            Me.otherLines.Add(New String() {"[Segments]", ""}) 'Missing + comments in section
            Me.otherLines.Add(New String() {"[Material]", ""}) 'Missing + comments in section
            Me.otherLines.Add(New String() {"[Properties]", ""}) 'Missing + comments in section
            Me.otherLines.Add(New String() {"[SelfWeight]", ""}) 'Missing + comments in section
            Me.otherLines.Add(New String() {"[Thermal]", ""}) 'Missing + comments in section
            Me.otherLines.Add(New String() {"[Combinations]", ""}) 'Missing + comments in section
            Me.otherLines.Add(New String() {"[Envelopes]", ""}) 'Missing + comments in section
            Me.otherLines.Add(New String() {"[Cases]", ""}) 'Missing + comments in section
            Me.otherLines.Add(New String() {"[OutputOptions]", ""}) 'Missing + comments in section
            Me.otherLines.Add(New String() {"[ReportOptions]", ""}) 'Missing + comments in section
            Me.otherLines.Add(New String() {"[EndCHRONOS]", ""}) 'Missing
            Me.otherLines.Add(New String() {"[Material]", ""}) 'Missing - This is weird because this is the second one
            Me.otherLines.Add(New String() {"[EndOverwrite]", ""}) 'Missing
        End If

        For Each line In Me.otherLines
            Select Case True
#Region "[Common]"
                Case line(0).Equals("[Common]")
                    newERIList.Add(line(0))
                    'Project Settings
                    newERIList.Add("DesignStandardSeries=" & Me.settings.projectInfo.DesignStandardSeries)
                    newERIList.Add("UnitsSystem=" & Me.settings.projectInfo.UnitsSystem)
                    newERIList.Add("ClientName=" & Me.settings.projectInfo.ClientName)
                    newERIList.Add("ProjectName=" & Me.settings.projectInfo.ProjectName)
                    newERIList.Add("ProjectNumber=" & Me.settings.projectInfo.ProjectNumber)
                    newERIList.Add("CreatedBy=" & Me.settings.projectInfo.CreatedBy)
                    newERIList.Add("CreatedOn=" & Me.settings.projectInfo.CreatedOn)
                    newERIList.Add("LastUsedBy=" & Me.settings.projectInfo.LastUsedBy)
                    newERIList.Add("LastUsedOn=" & Me.settings.projectInfo.LastUsedOn)
                    newERIList.Add("VersionUsed=" & Me.settings.projectInfo.VersionUsed)
#End Region
#Region "US Units"
                Case line(0).Equals("[US Units]")
                    newERIList.Add(line(0))
                    newERIList.Add("Length=" & Me.settings.USUnits.Length.value)
                    newERIList.Add("LengthPrec=" & Me.settings.USUnits.Length.precision)
                    newERIList.Add("Coordinate=" & Me.settings.USUnits.Coordinate.value)
                    newERIList.Add("CoordinatePrec=" & Me.settings.USUnits.Coordinate.precision)
                    newERIList.Add("Force=" & Me.settings.USUnits.Force.value)
                    newERIList.Add("ForcePrec=" & Me.settings.USUnits.Force.precision)
                    newERIList.Add("Load=" & Me.settings.USUnits.Load.value)
                    newERIList.Add("LoadPrec=" & Me.settings.USUnits.Load.precision)
                    newERIList.Add("Moment=" & Me.settings.USUnits.Moment.value)
                    newERIList.Add("MomentPrec=" & Me.settings.USUnits.Moment.precision)
                    newERIList.Add("Properties=" & Me.settings.USUnits.Properties.value)
                    newERIList.Add("PropertiesPrec=" & Me.settings.USUnits.Properties.precision)
                    newERIList.Add("Pressure=" & Me.settings.USUnits.Pressure.value)
                    newERIList.Add("PressurePrec=" & Me.settings.USUnits.Pressure.precision)
                    newERIList.Add("Velocity=" & Me.settings.USUnits.Velocity.value)
                    newERIList.Add("VelocityPrec=" & Me.settings.USUnits.Velocity.precision)
                    newERIList.Add("Displacement=" & Me.settings.USUnits.Displacement.value)
                    newERIList.Add("DisplacementPrec=" & Me.settings.USUnits.Displacement.precision)
                    newERIList.Add("Mass=" & Me.settings.USUnits.Mass.value)
                    newERIList.Add("MassPrec=" & Me.settings.USUnits.Mass.precision)
                    newERIList.Add("Acceleration=" & Me.settings.USUnits.Acceleration.value)
                    newERIList.Add("AccelerationPrec=" & Me.settings.USUnits.Acceleration.precision)
                    newERIList.Add("Stress=" & Me.settings.USUnits.Stress.value)
                    newERIList.Add("StressPrec=" & Me.settings.USUnits.Stress.precision)
                    newERIList.Add("Density=" & Me.settings.USUnits.Density.value)
                    newERIList.Add("DensityPrec=" & Me.settings.USUnits.Density.precision)
                    newERIList.Add("UnitWt=" & Me.settings.USUnits.UnitWt.value)
                    newERIList.Add("UnitWtPrec=" & Me.settings.USUnits.UnitWt.precision)
                    newERIList.Add("Strength=" & Me.settings.USUnits.Strength.value)
                    newERIList.Add("StrengthPrec=" & Me.settings.USUnits.Strength.precision)
                    newERIList.Add("Modulus=" & Me.settings.USUnits.Modulus.value)
                    newERIList.Add("ModulusPrec=" & Me.settings.USUnits.Modulus.precision)
                    newERIList.Add("Temperature=" & Me.settings.USUnits.Temperature.value)
                    newERIList.Add("TemperaturePrec=" & Me.settings.USUnits.Temperature.precision)
                    newERIList.Add("Printer=" & Me.settings.USUnits.Printer.value)
                    newERIList.Add("PrinterPrec=" & Me.settings.USUnits.Printer.precision)
                    newERIList.Add("Rotation=" & Me.settings.USUnits.Rotation.value)
                    newERIList.Add("RotationPrec=" & Me.settings.USUnits.Rotation.precision)
                    newERIList.Add("Spacing=" & Me.settings.USUnits.Spacing.value)
                    newERIList.Add("SpacingPrec=" & Me.settings.USUnits.Spacing.precision)
#End Region
#Region "SI Units"
                Case line(0).Equals("[SI Units]")
                    'SI Units was missing from creating file
                    newERIList.Add(line(0))
                    newERIList.AddRange(defaults.SIUnits)
#End Region
#Region "Codes"
                Case line(0).Equals("[Codes]")
                    newERIList.Add(line(0))
                    newERIList.Add("SteelCode=" & Me.code.design.DesignCode)
                    newERIList.AddRange(defaults.CodeValues)
#End Region
'Code below didn't pull in materials. maybe case sensitive? - added materials under [Databases]
'#Region "Materials"
'                Case line(0).Equals("[APPLICATION]")
'                    newERIList.Add(line(0))
'                    newERIList.Add("[DATABASE]")
'                    For Each mat In database.materials
'                        newERIList.Add(IIf(mat.IsBolt, "BoltMatFile=", "MembermatFile=") & mat.MemberMatFile)
'                        newERIList.Add("MatName=" & mat.MatName)
'                        newERIList.Add("MatValues=" & mat.MatValues)
'                    Next
'#End Region
#Region "[STRUCTURE]"
                Case line(0).Equals("[Structure]")

                    newERIList.Add(line(0))
#Region "User Info Settings"
                    newERIList.Add("ViewerUserName=" & Me.settings.userInfo.ViewerUserName)
                    newERIList.Add("ViewerCompanyName=" & Me.settings.userInfo.ViewerCompanyName)
                    newERIList.Add("ViewerStreetAddress=" & Me.settings.userInfo.ViewerStreetAddress)
                    newERIList.Add("ViewerCityState=" & Me.settings.userInfo.ViewerCityState)
                    newERIList.Add("ViewerPhone=" & Me.settings.userInfo.ViewerPhone)
                    newERIList.Add("ViewerFAX=" & Me.settings.userInfo.ViewerFAX)
                    newERIList.Add("ViewerLogo=" & Me.settings.userInfo.ViewerLogo)
                    newERIList.Add("ViewerCompanyBitmap=" & Me.settings.userInfo.ViewerCompanyBitmap)
#End Region
#Region "Report Settings"
                    newERIList.Add("ReportInputCosts=" & trueFalseYesNo(Me.reportSettings.ReportInputCosts))
                    newERIList.Add("ReportInputGeometry=" & trueFalseYesNo(Me.reportSettings.ReportInputGeometry))
                    newERIList.Add("ReportInputOptions=" & trueFalseYesNo(Me.reportSettings.ReportInputOptions))
                    newERIList.Add("ReportMaxForces=" & trueFalseYesNo(Me.reportSettings.ReportMaxForces))
                    newERIList.Add("ReportInputMap=" & trueFalseYesNo(Me.reportSettings.ReportInputMap))
                    newERIList.Add("CostReportOutputType=" & Me.reportSettings.CostReportOutputType)
                    newERIList.Add("CapacityReportOutputType=" & Me.reportSettings.CapacityReportOutputType)
                    newERIList.Add("ReportPrintForceTotals=" & trueFalseYesNo(Me.reportSettings.ReportPrintForceTotals))
                    newERIList.Add("ReportPrintForceDetails=" & trueFalseYesNo(Me.reportSettings.ReportPrintForceDetails))
                    newERIList.Add("ReportPrintMastVectors=" & trueFalseYesNo(Me.reportSettings.ReportPrintMastVectors))
                    newERIList.Add("ReportPrintAntPoleVectors=" & trueFalseYesNo(Me.reportSettings.ReportPrintAntPoleVectors))
                    newERIList.Add("ReportPrintDiscreteVectors=" & trueFalseYesNo(Me.reportSettings.ReportPrintDiscreteVectors))
                    newERIList.Add("ReportPrintDishVectors=" & trueFalseYesNo(Me.reportSettings.ReportPrintDishVectors))
                    newERIList.Add("ReportPrintFeedTowerVectors=" & trueFalseYesNo(Me.reportSettings.ReportPrintFeedTowerVectors))
                    newERIList.Add("ReportPrintUserLoadVectors=" & trueFalseYesNo(Me.reportSettings.ReportPrintUserLoadVectors))
                    newERIList.Add("ReportPrintPressures=" & trueFalseYesNo(Me.reportSettings.ReportPrintPressures))
                    newERIList.Add("ReportPrintAppurtForces=" & trueFalseYesNo(Me.reportSettings.ReportPrintAppurtForces))
                    newERIList.Add("ReportPrintGuyForces=" & trueFalseYesNo(Me.reportSettings.ReportPrintGuyForces))
                    newERIList.Add("ReportPrintGuyStressing=" & trueFalseYesNo(Me.reportSettings.ReportPrintGuyStressing))
                    newERIList.Add("ReportPrintDeflections=" & trueFalseYesNo(Me.reportSettings.ReportPrintDeflections))
                    newERIList.Add("ReportPrintReactions=" & trueFalseYesNo(Me.reportSettings.ReportPrintReactions))
                    newERIList.Add("ReportPrintStressChecks=" & trueFalseYesNo(Me.reportSettings.ReportPrintStressChecks))
                    newERIList.Add("ReportPrintBoltChecks=" & trueFalseYesNo(Me.reportSettings.ReportPrintBoltChecks))
                    newERIList.Add("ReportPrintInputGVerificationTables=" & trueFalseYesNo(Me.reportSettings.ReportPrintInputGVerificationTables))
                    newERIList.Add("ReportPrintOutputGVerificationTables=" & trueFalseYesNo(Me.reportSettings.ReportPrintOutputGVerificationTables))
#End Region
#Region "Options - Wind"
                    newERIList.Add("WindDirOption=" & Me.options.windDirections.WindDirOption)
                    newERIList.Add("WindDir0_0=" & trueFalseYesNo(Me.options.windDirections.WindDir0_0))
                    newERIList.Add("WindDir0_1=" & trueFalseYesNo(Me.options.windDirections.WindDir0_1))
                    newERIList.Add("WindDir0_2=" & trueFalseYesNo(Me.options.windDirections.WindDir0_2))
                    newERIList.Add("WindDir0_3=" & trueFalseYesNo(Me.options.windDirections.WindDir0_3))
                    newERIList.Add("WindDir0_4=" & trueFalseYesNo(Me.options.windDirections.WindDir0_4))
                    newERIList.Add("WindDir0_5=" & trueFalseYesNo(Me.options.windDirections.WindDir0_5))
                    newERIList.Add("WindDir0_6=" & trueFalseYesNo(Me.options.windDirections.WindDir0_6))
                    newERIList.Add("WindDir0_7=" & trueFalseYesNo(Me.options.windDirections.WindDir0_7))
                    newERIList.Add("WindDir0_8=" & trueFalseYesNo(Me.options.windDirections.WindDir0_8))
                    newERIList.Add("WindDir0_9=" & trueFalseYesNo(Me.options.windDirections.WindDir0_9))
                    newERIList.Add("WindDir0_10=" & trueFalseYesNo(Me.options.windDirections.WindDir0_10))
                    newERIList.Add("WindDir0_11=" & trueFalseYesNo(Me.options.windDirections.WindDir0_11))
                    newERIList.Add("WindDir0_12=" & trueFalseYesNo(Me.options.windDirections.WindDir0_12))
                    newERIList.Add("WindDir0_13=" & trueFalseYesNo(Me.options.windDirections.WindDir0_13))
                    newERIList.Add("WindDir0_14=" & trueFalseYesNo(Me.options.windDirections.WindDir0_14))
                    newERIList.Add("WindDir0_15=" & trueFalseYesNo(Me.options.windDirections.WindDir0_15))
                    newERIList.Add("WindDir1_0=" & trueFalseYesNo(Me.options.windDirections.WindDir1_0))
                    newERIList.Add("WindDir1_1=" & trueFalseYesNo(Me.options.windDirections.WindDir1_1))
                    newERIList.Add("WindDir1_2=" & trueFalseYesNo(Me.options.windDirections.WindDir1_2))
                    newERIList.Add("WindDir1_3=" & trueFalseYesNo(Me.options.windDirections.WindDir1_3))
                    newERIList.Add("WindDir1_4=" & trueFalseYesNo(Me.options.windDirections.WindDir1_4))
                    newERIList.Add("WindDir1_5=" & trueFalseYesNo(Me.options.windDirections.WindDir1_5))
                    newERIList.Add("WindDir1_6=" & trueFalseYesNo(Me.options.windDirections.WindDir1_6))
                    newERIList.Add("WindDir1_7=" & trueFalseYesNo(Me.options.windDirections.WindDir1_7))
                    newERIList.Add("WindDir1_8=" & trueFalseYesNo(Me.options.windDirections.WindDir1_8))
                    newERIList.Add("WindDir1_9=" & trueFalseYesNo(Me.options.windDirections.WindDir1_9))
                    newERIList.Add("WindDir1_10=" & trueFalseYesNo(Me.options.windDirections.WindDir1_10))
                    newERIList.Add("WindDir1_11=" & trueFalseYesNo(Me.options.windDirections.WindDir1_11))
                    newERIList.Add("WindDir1_12=" & trueFalseYesNo(Me.options.windDirections.WindDir1_12))
                    newERIList.Add("WindDir1_13=" & trueFalseYesNo(Me.options.windDirections.WindDir1_13))
                    newERIList.Add("WindDir1_14=" & trueFalseYesNo(Me.options.windDirections.WindDir1_14))
                    newERIList.Add("WindDir1_15=" & trueFalseYesNo(Me.options.windDirections.WindDir1_15))
                    newERIList.Add("WindDir2_0=" & trueFalseYesNo(Me.options.windDirections.WindDir2_0))
                    newERIList.Add("WindDir2_1=" & trueFalseYesNo(Me.options.windDirections.WindDir2_1))
                    newERIList.Add("WindDir2_2=" & trueFalseYesNo(Me.options.windDirections.WindDir2_2))
                    newERIList.Add("WindDir2_3=" & trueFalseYesNo(Me.options.windDirections.WindDir2_3))
                    newERIList.Add("WindDir2_4=" & trueFalseYesNo(Me.options.windDirections.WindDir2_4))
                    newERIList.Add("WindDir2_5=" & trueFalseYesNo(Me.options.windDirections.WindDir2_5))
                    newERIList.Add("WindDir2_6=" & trueFalseYesNo(Me.options.windDirections.WindDir2_6))
                    newERIList.Add("WindDir2_7=" & trueFalseYesNo(Me.options.windDirections.WindDir2_7))
                    newERIList.Add("WindDir2_8=" & trueFalseYesNo(Me.options.windDirections.WindDir2_8))
                    newERIList.Add("WindDir2_9=" & trueFalseYesNo(Me.options.windDirections.WindDir2_9))
                    newERIList.Add("WindDir2_10=" & trueFalseYesNo(Me.options.windDirections.WindDir2_10))
                    newERIList.Add("WindDir2_11=" & trueFalseYesNo(Me.options.windDirections.WindDir2_11))
                    newERIList.Add("WindDir2_12=" & trueFalseYesNo(Me.options.windDirections.WindDir2_12))
                    newERIList.Add("WindDir2_13=" & trueFalseYesNo(Me.options.windDirections.WindDir2_13))
                    newERIList.Add("WindDir2_14=" & trueFalseYesNo(Me.options.windDirections.WindDir2_14))
                    newERIList.Add("WindDir2_15=" & trueFalseYesNo(Me.options.windDirections.WindDir2_15))
                    newERIList.Add("SuppressWindPatternLoading=" & trueFalseYesNo(Me.options.windDirections.SuppressWindPatternLoading))
                    newERIList.Add("DoELCAnalysis=No")
#End Region
#Region "CCIReport"
                    newERIList.Add("sReportProjectNumber=" & Me.CCIReport.sReportProjectNumber)
                    newERIList.Add("sReportJobType=" & Me.CCIReport.sReportJobType)
                    newERIList.Add("sReportCarrierName=" & Me.CCIReport.sReportCarrierName)
                    newERIList.Add("sReportCarrierSiteNumber=" & Me.CCIReport.sReportCarrierSiteNumber)
                    newERIList.Add("sReportCarrierSiteName=" & Me.CCIReport.sReportCarrierSiteName)
                    newERIList.Add("sReportSiteAddress=" & Me.CCIReport.sReportSiteAddress)
                    newERIList.Add("sReportLatitudeDegree=" & Me.CCIReport.sReportLatitudeDegree)
                    newERIList.Add("sReportLatitudeMinute=" & Me.CCIReport.sReportLatitudeMinute)
                    newERIList.Add("sReportLatitudeSecond=" & Me.CCIReport.sReportLatitudeSecond)
                    newERIList.Add("sReportLongitudeDegree=" & Me.CCIReport.sReportLongitudeDegree)
                    newERIList.Add("sReportLongitudeMinute=" & Me.CCIReport.sReportLongitudeMinute)
                    newERIList.Add("sReportLongitudeSecond=" & Me.CCIReport.sReportLongitudeSecond)
                    newERIList.Add("sReportLocalCodeRequirement=" & Me.CCIReport.sReportLocalCodeRequirement)
                    newERIList.Add("sReportSiteHistory=" & Me.CCIReport.sReportSiteHistory)
                    newERIList.Add("sReportTowerManufacturer=" & Me.CCIReport.sReportTowerManufacturer)
                    newERIList.Add("sReportMonthManufactured=" & Me.CCIReport.sReportMonthManufactured)
                    newERIList.Add("sReportYearManufactured=" & Me.CCIReport.sReportYearManufactured)
                    newERIList.Add("sReportOriginalSpeed=" & Me.CCIReport.sReportOriginalSpeed)
                    newERIList.Add("sReportOriginalCode=" & Me.CCIReport.sReportOriginalCode)
                    newERIList.Add("sReportTowerType=" & Me.CCIReport.sReportTowerType)
                    newERIList.Add("sReportEngrName=" & Me.CCIReport.sReportEngrName)
                    newERIList.Add("sReportEngrTitle=" & Me.CCIReport.sReportEngrTitle)
                    newERIList.Add("sReportHQPhoneNumber=" & Me.CCIReport.sReportHQPhoneNumber)
                    newERIList.Add("sReportEmailAddress=" & Me.CCIReport.sReportEmailAddress)
                    newERIList.Add("sReportLogoPath=" & Me.CCIReport.sReportLogoPath)
                    newERIList.Add("sReportCCiContactName=" & Me.CCIReport.sReportCCiContactName)
                    newERIList.Add("sReportCCiAddress1=" & Me.CCIReport.sReportCCiAddress1)
                    newERIList.Add("sReportCCiAddress2=" & Me.CCIReport.sReportCCiAddress2)
                    newERIList.Add("sReportCCiBUNumber=" & Me.CCIReport.sReportCCiBUNumber)
                    newERIList.Add("sReportCCiSiteName=" & Me.CCIReport.sReportCCiSiteName)
                    newERIList.Add("sReportCCiJDENumber=" & Me.CCIReport.sReportCCiJDENumber)
                    newERIList.Add("sReportCCiWONumber=" & Me.CCIReport.sReportCCiWONumber)
                    newERIList.Add("sReportCCiPONumber=" & Me.CCIReport.sReportCCiPONumber)
                    newERIList.Add("sReportCCiAppNumber=" & Me.CCIReport.sReportCCiAppNumber)
                    newERIList.Add("sReportCCiRevNumber=" & Me.CCIReport.sReportCCiRevNumber)
                    For Each row In Me.CCIReport.sReportDocsProvided
                        newERIList.Add("sReportDocsProvided=" & row)
                    Next
                    newERIList.Add("sReportRecommendations=" & Me.CCIReport.sReportRecommendations)
                    For Each row In Me.CCIReport.sReportAppurt1
                        newERIList.Add("sReportAppurt1=" & row)
                    Next
                    For Each row In Me.CCIReport.sReportAppurt2
                        newERIList.Add("sReportAppurt2=" & row)
                    Next
                    For Each row In Me.CCIReport.sReportAppurt3
                        newERIList.Add("sReportAppurt3=" & row)
                    Next
                    For Each row In Me.CCIReport.sReportAddlCapacity
                        newERIList.Add("sReportAddlCapacity=" & row)
                    Next
                    For Each row In Me.CCIReport.sReportAssumption
                        newERIList.Add("sReportAssumption=" & row)
                    Next
                    newERIList.Add("sReportAppurt1Note1=" & Me.CCIReport.sReportAppurt1Note1)
                    newERIList.Add("sReportAppurt1Note2=" & Me.CCIReport.sReportAppurt1Note2)
                    newERIList.Add("sReportAppurt1Note3=" & Me.CCIReport.sReportAppurt1Note3)
                    newERIList.Add("sReportAppurt1Note4=" & Me.CCIReport.sReportAppurt1Note4)
                    newERIList.Add("sReportAppurt1Note5=" & Me.CCIReport.sReportAppurt1Note5)
                    newERIList.Add("sReportAppurt1Note6=" & Me.CCIReport.sReportAppurt1Note6)
                    newERIList.Add("sReportAppurt1Note7=" & Me.CCIReport.sReportAppurt1Note7)
                    newERIList.Add("sReportAppurt2Note1=" & Me.CCIReport.sReportAppurt2Note1)
                    newERIList.Add("sReportAppurt2Note2=" & Me.CCIReport.sReportAppurt2Note2)
                    newERIList.Add("sReportAppurt2Note3=" & Me.CCIReport.sReportAppurt2Note3)
                    newERIList.Add("sReportAppurt2Note4=" & Me.CCIReport.sReportAppurt2Note4)
                    newERIList.Add("sReportAppurt2Note5=" & Me.CCIReport.sReportAppurt2Note5)
                    newERIList.Add("sReportAppurt2Note6=" & Me.CCIReport.sReportAppurt2Note6)
                    newERIList.Add("sReportAppurt2Note7=" & Me.CCIReport.sReportAppurt2Note7)
                    newERIList.Add("sReportAddlCapacityNote1=" & Me.CCIReport.sReportAddlCapacityNote1)
                    newERIList.Add("sReportAddlCapacityNote2=" & Me.CCIReport.sReportAddlCapacityNote2)
                    newERIList.Add("sReportAddlCapacityNote3=" & Me.CCIReport.sReportAddlCapacityNote3)
                    newERIList.Add("sReportAddlCapacityNote4=" & Me.CCIReport.sReportAddlCapacityNote4)
#End Region
#Region "Solution Settings"
                    newERIList.Add("SolutionUsePDelta=" & trueFalseYesNo(Me.solutionSettings.SolutionUsePDelta))
                    newERIList.Add("SolutionMinStiffness=" & Me.solutionSettings.SolutionMinStiffness)
                    newERIList.Add("SolutionMaxStiffness=" & Me.solutionSettings.SolutionMaxStiffness)
                    newERIList.Add("SolutionMaxCycles=" & Me.solutionSettings.SolutionMaxCycles)
                    newERIList.Add("SolutionPower=" & Me.solutionSettings.SolutionPower)
                    newERIList.Add("SolutionTolerance=" & Me.solutionSettings.SolutionTolerance)
#End Region
#Region "MTO Settings"
                    newERIList.Add("IncludeCapacityNote=" & trueFalseYesNo(Me.MTOSettings.IncludeCapacityNote))
                    newERIList.Add("IncludeAppurtGraphics=" & trueFalseYesNo(Me.MTOSettings.IncludeAppurtGraphics))
                    newERIList.Add("DisplayNotes=" & trueFalseYesNo(Me.MTOSettings.DisplayNotes))
                    newERIList.Add("DisplayReactions=" & trueFalseYesNo(Me.MTOSettings.DisplayReactions))
                    newERIList.Add("DisplaySchedule=" & trueFalseYesNo(Me.MTOSettings.DisplaySchedule))
                    newERIList.Add("DisplayAppurtenanceTable=" & trueFalseYesNo(Me.MTOSettings.DisplayAppurtenanceTable))
                    newERIList.Add("DisplayMaterialStrengthTable=" & trueFalseYesNo(Me.MTOSettings.DisplayMaterialStrengthTable))
                    For Each note As String In Split(Me.MTOSettings.Notes, "||")
                        newERIList.Add("Notes=" & note)
                    Next
#End Region
#Region "Code - Design"
                    newERIList.Add("DesignCode=" & Me.code.design.DesignCode)
                    newERIList.Add("ERIDesignMode=" & Me.code.design.ERIDesignMode)
                    newERIList.Add("DoInteraction=" & trueFalseYesNo(Me.code.design.DoInteraction))
                    newERIList.Add("DoHorzInteraction=" & trueFalseYesNo(Me.code.design.DoHorzInteraction))
                    newERIList.Add("DoDiagInteraction=" & trueFalseYesNo(Me.code.design.DoDiagInteraction))
                    newERIList.Add("UseMomentMagnification=" & trueFalseYesNo(Me.code.design.UseMomentMagnification))
                    newERIList.Add("UseCodeStressRatio=" & trueFalseYesNo(Me.code.design.UseCodeStressRatio))
                    newERIList.Add("AllowStressRatio=" & Me.code.design.AllowStressRatio)
                    newERIList.Add("AllowAntStressRatio=" & Me.code.design.AllowAntStressRatio)
                    newERIList.Add("UseCodeGuySF=" & trueFalseYesNo(Me.code.design.UseCodeGuySF))
                    newERIList.Add("GuySF=" & Me.code.design.GuySF)
                    newERIList.Add("UseTIA222H_AnnexS=" & trueFalseYesNo(Me.code.design.UseTIA222H_AnnexS))
                    If Not IsNothing(Me.code.design.TIA_222_H_AnnexS_Ratio) Then 'Sometimes this value stores as null in EDS which causes errors if populated in TNX for subsequent SAs
                        newERIList.Add("TIA_222_H_AnnexS_Ratio=" & Me.code.design.TIA_222_H_AnnexS_Ratio)
                    End If
                    newERIList.Add("PrintBitmaps=" & trueFalseYesNo(Me.code.design.PrintBitmaps))
#End Region
#Region "Code - Wind"
                    newERIList.Add("WindSpeed=" & Me.settings.USUnits.Velocity.convertToERIUnits(Me.code.wind.WindSpeed))
                    newERIList.Add("WindSpeedIce=" & Me.settings.USUnits.Velocity.convertToERIUnits(Me.code.wind.WindSpeedIce))
                    newERIList.Add("WindSpeedService=" & Me.settings.USUnits.Velocity.convertToERIUnits(Me.code.wind.WindSpeedService))
                    newERIList.Add("UseStateCountyLookup=" & trueFalseYesNo(Me.code.wind.UseStateCountyLookup))
                    newERIList.Add("State=" & Me.code.wind.State)
                    newERIList.Add("County=" & Me.code.wind.County)
                    newERIList.Add("UseMaxKz=" & trueFalseYesNo(Me.code.wind.UseMaxKz))
                    newERIList.Add("ASCE_7_10_WindData=" & trueFalseYesNo(Me.code.wind.ASCE_7_10_WindData))
                    newERIList.Add("ASCE_7_10_ConvertWindToASD=" & trueFalseYesNo(Me.code.wind.ASCE_7_10_ConvertWindToASD))
                    newERIList.Add("UseASCEWind=" & trueFalseYesNo(Me.code.wind.UseASCEWind))
                    newERIList.Add("AutoCalc_ASCE_GH=" & trueFalseYesNo(Me.code.wind.AutoCalc_ASCE_GH))
                    newERIList.Add("ASCE_ExposureCat=" & Me.code.wind.ASCE_ExposureCat)
                    newERIList.Add("ASCE_Year=" & Me.code.wind.ASCE_Year)
                    newERIList.Add("ASCEGh=" & Me.code.wind.ASCEGh)
                    newERIList.Add("ASCEI=" & Me.code.wind.ASCEI)
                    newERIList.Add("CalcWindAt=" & Me.code.wind.CalcWindAt)
                    newERIList.Add("WindCalcPoints=" & Me.settings.USUnits.Length.convertToERIUnits(Me.code.wind.WindCalcPoints))
                    newERIList.Add("WindExposure=" & Me.code.wind.WindExposure)
                    newERIList.Add("StructureCategory=" & Me.code.wind.StructureCategory)
                    newERIList.Add("RiskCategory=" & Me.code.wind.RiskCategory)
                    newERIList.Add("TopoCategory=" & Me.code.wind.TopoCategory)
                    newERIList.Add("RSMTopographicFeature=" & Me.code.wind.RSMTopographicFeature)
                    newERIList.Add("RSM_L=" & Me.settings.USUnits.Length.convertToERIUnits(Me.code.wind.RSM_L))
                    newERIList.Add("RSM_X=" & Me.settings.USUnits.Length.convertToERIUnits(Me.code.wind.RSM_X))
                    newERIList.Add("CrestHeight=" & Me.settings.USUnits.Length.convertToERIUnits(Me.code.wind.CrestHeight))
                    newERIList.Add("TIA_222_H_TopoFeatureDownwind=" & trueFalseYesNo(Me.code.wind.TIA_222_H_TopoFeatureDownwind))
                    newERIList.Add("BaseElevAboveSeaLevel=" & Me.settings.USUnits.Length.convertToERIUnits(Me.code.wind.BaseElevAboveSeaLevel))
                    newERIList.Add("ConsiderRooftopSpeedUp=" & trueFalseYesNo(Me.code.wind.ConsiderRooftopSpeedUp))
                    newERIList.Add("RooftopWS=" & Me.settings.USUnits.Length.convertToERIUnits(Me.code.wind.RooftopWS))
                    newERIList.Add("RooftopHS=" & Me.settings.USUnits.Length.convertToERIUnits(Me.code.wind.RooftopHS))
                    newERIList.Add("RooftopParapetHt=" & Me.settings.USUnits.Length.convertToERIUnits(Me.code.wind.RooftopParapetHt))
                    newERIList.Add("RooftopXB=" & Me.settings.USUnits.Length.convertToERIUnits(Me.code.wind.RooftopXB))
                    newERIList.Add("WindZone=" & Me.code.wind.WindZone)
                    newERIList.Add("EIACWindMult=" & Me.code.wind.EIACWindMult)
                    newERIList.Add("EIACWindMultIce=" & Me.code.wind.EIACWindMultIce)
                    newERIList.Add("EIACIgnoreCableDrag=" & trueFalseYesNo(Me.code.wind.EIACIgnoreCableDrag))
                    newERIList.Add("CSA_S37_RefVelPress=" & Me.settings.USUnits.Pressure.convertToERIUnits(Me.code.wind.CSA_S37_RefVelPress))
                    newERIList.Add("CSA_S37_ReliabilityClass=" & Me.code.wind.CSA_S37_ReliabilityClass)
                    newERIList.Add("CSA_S37_ServiceabilityFactor=" & Me.code.wind.CSA_S37_ServiceabilityFactor)
#End Region
#Region "Code - Seismic"
                    newERIList.Add("UseASCE7_10_Seismic_Lcomb=" & trueFalseYesNo(Me.code.seismic.UseASCE7_10_Seismic_Lcomb))
                    newERIList.Add("SeismicSiteClass=" & Me.code.seismic.SeismicSiteClass)
                    newERIList.Add("SeismicSs=" & Me.code.seismic.SeismicSs)
                    newERIList.Add("SeismicS1=" & Me.code.seismic.SeismicS1)
#End Region
#Region "Code - Ice"
                    newERIList.Add("IceThickness=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.code.ice.IceThickness))
                    newERIList.Add("IceDensity=" & Me.settings.USUnits.Density.convertToERIUnits(Me.code.ice.IceDensity))
                    newERIList.Add("UseModified_TIA_222_IceParameters=" & trueFalseYesNo(Me.code.ice.UseModified_TIA_222_IceParameters))
                    newERIList.Add("TIA_222_IceThicknessMultiplier=" & Me.code.ice.TIA_222_IceThicknessMultiplier)
                    newERIList.Add("DoNotUse_TIA_222_IceEscalation=" & trueFalseYesNo(Me.code.ice.DoNotUse_TIA_222_IceEscalation))
                    newERIList.Add("UseIceEscalation=" & trueFalseYesNo(Me.code.ice.UseIceEscalation))
#End Region
#Region "Code - Thermal"
                    newERIList.Add("TempDrop=" & Me.settings.USUnits.Temperature.convertToERIUnits(Me.code.thermal.TempDrop))
#End Region
#Region "Code - Miscl"
                    newERIList.Add("GroutFc=" & Me.settings.USUnits.Strength.convertToERIUnits(Me.code.misclCode.GroutFc))
                    newERIList.Add("TowerBoltGrade=" & Me.code.misclCode.TowerBoltGrade)
                    newERIList.Add("TowerBoltMinEdgeDist=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.code.misclCode.TowerBoltMinEdgeDist))
#End Region
#Region "Options - General"
                    newERIList.Add("UseClearSpans=" & trueFalseYesNo(Me.options.UseClearSpans))
                    newERIList.Add("UseClearSpansKlr=" & trueFalseYesNo(Me.options.UseClearSpansKlr))
                    newERIList.Add("UseFeedlineAsCylinder=" & trueFalseYesNo(Me.options.UseFeedlineAsCylinder))
                    newERIList.Add("UseLegLoads=" & trueFalseYesNo(Me.options.UseLegLoads))
                    newERIList.Add("SRTakeCompression=" & trueFalseYesNo(Me.options.SRTakeCompression))
                    newERIList.Add("AllLegPanelsSame=" & trueFalseYesNo(Me.options.AllLegPanelsSame))
                    newERIList.Add("UseCombinedBoltCapacity=" & trueFalseYesNo(Me.options.UseCombinedBoltCapacity))
                    newERIList.Add("SecHorzBracesLeg=" & trueFalseYesNo(Me.options.SecHorzBracesLeg))
                    newERIList.Add("SortByComponent=" & trueFalseYesNo(Me.options.SortByComponent))
                    newERIList.Add("SRCutEnds=" & trueFalseYesNo(Me.options.SRCutEnds))
                    newERIList.Add("SRConcentric=" & trueFalseYesNo(Me.options.SRConcentric))
                    newERIList.Add("CalcBlockShear=" & trueFalseYesNo(Me.options.CalcBlockShear))
                    newERIList.Add("Use4SidedDiamondBracing=" & trueFalseYesNo(Me.options.Use4SidedDiamondBracing))
                    newERIList.Add("TriangulateInnerBracing=" & trueFalseYesNo(Me.options.TriangulateInnerBracing))
                    newERIList.Add("PrintCarrierNotes=" & trueFalseYesNo(Me.options.PrintCarrierNotes))
                    newERIList.Add("AddIBCWindCase=" & trueFalseYesNo(Me.options.AddIBCWindCase))
                    newERIList.Add("LegBoltsAtTop=" & trueFalseYesNo(Me.options.LegBoltsAtTop))
                    newERIList.Add("UseTIA222Exemptions_MinBracingResistance=" & trueFalseYesNo(Me.options.UseTIA222Exemptions_MinBracingResistance))
                    newERIList.Add("UseTIA222Exemptions_TensionSplice=" & trueFalseYesNo(Me.options.UseTIA222Exemptions_TensionSplice))
                    newERIList.Add("IgnoreKLryFor60DegAngleLegs=" & trueFalseYesNo(Me.options.IgnoreKLryFor60DegAngleLegs))
                    newERIList.Add("UseFeedlineTorque=" & trueFalseYesNo(Me.options.UseFeedlineTorque))
                    newERIList.Add("UsePinnedElements=" & trueFalseYesNo(Me.options.UsePinnedElements))
                    newERIList.Add("UseRigidIndex=" & trueFalseYesNo(Me.options.UseRigidIndex))
                    newERIList.Add("UseTrueCable=" & trueFalseYesNo(Me.options.UseTrueCable))
                    newERIList.Add("UseASCELy=" & trueFalseYesNo(Me.options.UseASCELy))
                    newERIList.Add("CalcBracingForces=" & trueFalseYesNo(Me.options.CalcBracingForces))
                    newERIList.Add("IgnoreBracingFEA=" & trueFalseYesNo(Me.options.IgnoreBracingFEA))
                    newERIList.Add("BypassStabilityChecks=" & trueFalseYesNo(Me.options.BypassStabilityChecks))
                    newERIList.Add("UseWindProjection=" & trueFalseYesNo(Me.options.UseWindProjection))
                    newERIList.Add("UseDishCoeff=" & trueFalseYesNo(Me.options.UseDishCoeff))
                    newERIList.Add("AutoCalcTorqArmArea=" & trueFalseYesNo(Me.options.AutoCalcTorqArmArea))
#End Region
#Region "Options - Foundations"
                    newERIList.Add("MastVert=" & Me.settings.USUnits.convertForcePerUnitLengthtoERISpecified(Me.options.foundationStiffness.MastVert))
                    newERIList.Add("MastHorz=" & Me.settings.USUnits.convertForcePerUnitLengthtoERISpecified(Me.options.foundationStiffness.MastHorz))
                    newERIList.Add("GuyVert=" & Me.settings.USUnits.convertForcePerUnitLengthtoERISpecified(Me.options.foundationStiffness.GuyVert))
                    newERIList.Add("GuyHorz=" & Me.settings.USUnits.convertForcePerUnitLengthtoERISpecified(Me.options.foundationStiffness.GuyHorz))
#End Region
#Region "Options - Poles"
                    newERIList.Add("CheckVonMises=" & trueFalseYesNo(Me.options.cantileverPoles.CheckVonMises))
                    newERIList.Add("SocketTopMount=" & trueFalseYesNo(Me.options.cantileverPoles.SocketTopMount))
                    newERIList.Add("PrintMonopoleAtIncrements=" & trueFalseYesNo(Me.options.cantileverPoles.PrintMonopoleAtIncrements))
                    newERIList.Add("UseSubCriticalFlow=" & trueFalseYesNo(Me.options.cantileverPoles.UseSubCriticalFlow))
                    newERIList.Add("AssumePoleWithNoAttachments=" & trueFalseYesNo(Me.options.cantileverPoles.AssumePoleWithNoAttachments))
                    newERIList.Add("AssumePoleWithShroud=" & trueFalseYesNo(Me.options.cantileverPoles.AssumePoleWithShroud))
                    newERIList.Add("PoleCornerRadiusKnown=" & trueFalseYesNo(Me.options.cantileverPoles.PoleCornerRadiusKnown))
                    newERIList.Add("CantKFactor=" & Me.options.cantileverPoles.CantKFactor)
#End Region
#Region "Options - Girts"
                    newERIList.Add("GirtOffset=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.options.defaultGirtOffsets.GirtOffset))
                    newERIList.Add("GirtOffsetLatticedPole=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.options.defaultGirtOffsets.GirtOffsetLatticedPole))
                    newERIList.Add("OffsetBotGirt=" & trueFalseYesNo(Me.options.defaultGirtOffsets.OffsetBotGirt))
#End Region
#Region "Options - Miscl"
                    newERIList.Add("HogRodTakeup=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.options.misclOptions.HogRodTakeup))
                    newERIList.Add("RadiusSampleDist=" & Me.settings.USUnits.Length.convertToERIUnits(Me.options.misclOptions.RadiusSampleDist))
#End Region
#Region "General Geometry"
                    newERIList.Add("TowerType=" & Me.geometry.TowerType)
                    newERIList.Add("AntennaType=" & Me.geometry.AntennaType)
                    newERIList.Add("OverallHeight=" & Me.settings.USUnits.Length.convertToERIUnits(Me.geometry.OverallHeight))
                    newERIList.Add("BaseElevation=" & Me.settings.USUnits.Length.convertToERIUnits(Me.geometry.BaseElevation))
                    newERIList.Add("Lambda=" & Me.settings.USUnits.Spacing.convertToERIUnits(Me.geometry.Lambda))
                    newERIList.Add("TowerTopFaceWidth=" & Me.settings.USUnits.Spacing.convertToERIUnits(Me.geometry.TowerTopFaceWidth))
                    newERIList.Add("TowerBaseFaceWidth=" & Me.settings.USUnits.Spacing.convertToERIUnits(Me.geometry.TowerBaseFaceWidth))
                    newERIList.Add("TowerTaper=" & Me.geometry.TowerTaper)
                    newERIList.Add("GuyedMonopoleBaseType=" & Me.geometry.GuyedMonopoleBaseType)
                    newERIList.Add("TaperHeight=" & Me.settings.USUnits.Length.convertToERIUnits(Me.geometry.TaperHeight))
                    newERIList.Add("PivotHeight=" & Me.settings.USUnits.Length.convertToERIUnits(Me.geometry.PivotHeight))
                    newERIList.Add("AutoCalcGH=" & trueFalseYesNo(Me.geometry.AutoCalcGH))
                    newERIList.Add("UserGHElev=" & Me.settings.USUnits.Length.convertToERIUnits(Me.geometry.UserGHElev))
                    newERIList.Add("UseIndexPlate=" & trueFalseYesNo(Me.geometry.UseIndexPlate))
                    newERIList.Add("EnterUserDefinedGhValues=" & trueFalseYesNo(Me.geometry.EnterUserDefinedGhValues))
                    newERIList.Add("BaseTowerGhInput=" & Me.geometry.BaseTowerGhInput)
                    newERIList.Add("UpperStructureGhInput=" & Me.geometry.UpperStructureGhInput)
                    newERIList.Add("EnterUserDefinedCgValues=" & trueFalseYesNo(Me.geometry.EnterUserDefinedCgValues))
                    newERIList.Add("BaseTowerCgInput=" & Me.geometry.BaseTowerCgInput)
                    newERIList.Add("UpperStructureCgInput=" & Me.geometry.UpperStructureCgInput)
                    newERIList.Add("AntennaFaceWidth=" & Me.settings.USUnits.Spacing.convertToERIUnits(Me.geometry.AntennaFaceWidth))
                    newERIList.Add("UseTopTakeup=" & trueFalseYesNo(Me.geometry.UseTopTakeup))
                    newERIList.Add("ConstantSlope=" & trueFalseYesNo(Me.geometry.ConstantSlope))

#End Region
#Region "Missing Structure Variables"
                    newERIList.AddRange(defaults.AntennaPole)
                    newERIList.AddRange(defaults.Beacon)
                    newERIList.AddRange(defaults.DesignChoicesList)
                    newERIList.AddRange(defaults.MRSS)
                    newERIList.AddRange(defaults.FTG)
                    newERIList.AddRange(defaults.MissingSettings)
                    newERIList.AddRange(defaults.CostData)
                    newERIList.AddRange(defaults.BasePlate)
#End Region

#End Region
#Region "Geometry"
                Case line(0).Equals("NumAntennaRecs")
                    newERIList.Add(line(0) & "=" & Me.geometry.upperStructure.Count)
                    'For Each upperSection In Me.geometry.upperStructure
                    ''''^^Not entirely sure why this is commented out. I would much rather use For Each here but I assume it is because of the way it is sorted. 
                    For i = 0 To Me.geometry.upperStructure.Count - 1
                        With Me.geometry.upperStructure(i)
                            newERIList.Add("AntennaRec=" & .Rec)
                            newERIList.Add("AntennaBraceType=" & .AntennaBraceType)
                            newERIList.Add("AntennaHeight=" & Me.settings.USUnits.Coordinate.convertToERIUnits(.AntennaHeight))
                            newERIList.Add("AntennaDiagonalSpacing=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaDiagonalSpacing))
                            newERIList.Add("AntennaDiagonalSpacingEx=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaDiagonalSpacingEx))
                            newERIList.Add("AntennaNumSections=" & .AntennaNumSections)
                            newERIList.Add("AntennaNumSesctions=" & .AntennaNumSesctions)
                            newERIList.Add("AntennaSectionLength=" & Me.settings.USUnits.Length.convertToERIUnits(.AntennaSectionLength))
                            newERIList.Add("AntennaLegType=" & .AntennaLegType)
                            newERIList.Add("AntennaLegSize=" & .AntennaLegSize)
                            If Not IsNothing(Me.settings.USUnits.Strength.convertToERIUnits(.AntennaLegGrade)) Then 'This field does not get saved when generated from CCIpole
                                newERIList.Add("AntennaLegGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(.AntennaLegGrade))
                            End If
                            If IsSomethingString(.AntennaLegMatlGrade) Then 'This field does not get saved when generated from CCIpole
                                newERIList.Add("AntennaLegMatlGrade=" & .AntennaLegMatlGrade)
                            End If
                            If Not IsNothing(Me.settings.USUnits.Strength.convertToERIUnits(.AntennaDiagonalGrade)) Then 'This field does not get saved when generated from CCIpole
                                newERIList.Add("AntennaDiagonalGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(.AntennaDiagonalGrade))
                            End If
                            If IsSomethingString(.AntennaDiagonalMatlGrade) Then 'This field does not get saved when generated from CCIpole
                                newERIList.Add("AntennaDiagonalMatlGrade=" & .AntennaDiagonalMatlGrade)
                            End If
                            If Not IsNothing(Me.settings.USUnits.Strength.convertToERIUnits(.AntennaInnerBracingGrade)) Then 'This field does not get saved when generated from CCIpole
                                newERIList.Add("AntennaInnerBracingGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(.AntennaInnerBracingGrade))
                            End If
                            If IsSomethingString(.AntennaInnerBracingMatlGrade) Then 'This field does not get saved when generated from CCIpole
                                newERIList.Add("AntennaInnerBracingMatlGrade=" & .AntennaInnerBracingMatlGrade)
                            End If
                            If Not IsNothing(Me.settings.USUnits.Strength.convertToERIUnits(.AntennaTopGirtGrade)) Then 'This field does not get saved when generated from CCIpole
                                newERIList.Add("AntennaTopGirtGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(.AntennaTopGirtGrade))
                            End If
                            If IsSomethingString(.AntennaTopGirtMatlGrade) Then 'This field does not get saved when generated from CCIpole
                                newERIList.Add("AntennaTopGirtMatlGrade=" & .AntennaTopGirtMatlGrade)
                            End If
                            If Not IsNothing(Me.settings.USUnits.Strength.convertToERIUnits(.AntennaBotGirtGrade)) Then 'This field does not get saved when generated from CCIpole
                                newERIList.Add("AntennaBotGirtGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(.AntennaBotGirtGrade))
                            End If
                            If IsSomethingString(.AntennaBotGirtMatlGrade) Then 'This field does not get saved when generated from CCIpole
                                newERIList.Add("AntennaBotGirtMatlGrade=" & .AntennaBotGirtMatlGrade)
                            End If
                            If Not IsNothing(Me.settings.USUnits.Strength.convertToERIUnits(.AntennaInnerGirtGrade)) Then 'This field does not get saved when generated from CCIpole
                                newERIList.Add("AntennaInnerGirtGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(.AntennaInnerGirtGrade))
                            End If
                            If IsSomethingString(.AntennaInnerGirtMatlGrade) Then 'This field does not get saved when generated from CCIpole
                                newERIList.Add("AntennaInnerGirtMatlGrade=" & .AntennaInnerGirtMatlGrade)
                            End If
                            If Not IsNothing(Me.settings.USUnits.Strength.convertToERIUnits(.AntennaLongHorizontalGrade)) Then 'This field does not get saved when generated from CCIpole
                                newERIList.Add("AntennaLongHorizontalGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(.AntennaLongHorizontalGrade))
                            End If
                            If IsSomethingString(.AntennaLongHorizontalMatlGrade) Then 'This field does not get saved when generated from CCIpole
                                newERIList.Add("AntennaLongHorizontalMatlGrade=" & .AntennaLongHorizontalMatlGrade)
                            End If
                            If Not IsNothing(Me.settings.USUnits.Strength.convertToERIUnits(.AntennaShortHorizontalGrade)) Then 'This field does not get saved when generated from CCIpole
                                newERIList.Add("AntennaShortHorizontalGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(.AntennaShortHorizontalGrade))
                            End If
                            If IsSomethingString(.AntennaShortHorizontalMatlGrade) Then 'This field does not get saved when generated from CCIpole
                                newERIList.Add("AntennaShortHorizontalMatlGrade=" & .AntennaShortHorizontalMatlGrade)
                            End If
                            newERIList.Add("AntennaDiagonalType=" & .AntennaDiagonalType)
                            newERIList.Add("AntennaDiagonalSize=" & .AntennaDiagonalSize)
                            newERIList.Add("AntennaInnerBracingType=" & .AntennaInnerBracingType)
                            newERIList.Add("AntennaInnerBracingSize=" & .AntennaInnerBracingSize)
                            newERIList.Add("AntennaTopGirtType=" & .AntennaTopGirtType)
                            newERIList.Add("AntennaTopGirtSize=" & .AntennaTopGirtSize)
                            newERIList.Add("AntennaBotGirtType=" & .AntennaBotGirtType)
                            newERIList.Add("AntennaBotGirtSize=" & .AntennaBotGirtSize)
                            newERIList.Add("AntennaTopGirtOffset=" & .AntennaTopGirtOffset)
                            newERIList.Add("AntennaBotGirtOffset=" & .AntennaBotGirtOffset)
                            newERIList.Add("AntennaHasKBraceEndPanels=" & trueFalseYesNo(.AntennaHasKBraceEndPanels))
                            newERIList.Add("AntennaHasHorizontals=" & trueFalseYesNo(.AntennaHasHorizontals))
                            newERIList.Add("AntennaLongHorizontalType=" & .AntennaLongHorizontalType)
                            newERIList.Add("AntennaLongHorizontalSize=" & .AntennaLongHorizontalSize)
                            newERIList.Add("AntennaShortHorizontalType=" & .AntennaShortHorizontalType)
                            newERIList.Add("AntennaShortHorizontalSize=" & .AntennaShortHorizontalSize)
                            If Not IsNothing(Me.settings.USUnits.Strength.convertToERIUnits(.AntennaRedundantGrade)) Then 'This field does not get saved when generated from CCIpole
                                newERIList.Add("AntennaRedundantGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(.AntennaRedundantGrade))
                            End If
                            If IsSomethingString(.AntennaRedundantMatlGrade) Then 'This field does not get saved when generated from CCIpole
                                newERIList.Add("AntennaRedundantMatlGrade=" & .AntennaRedundantMatlGrade)
                            End If
                            newERIList.Add("AntennaRedundantType=" & .AntennaRedundantType)
                            newERIList.Add("AntennaRedundantDiagType=" & .AntennaRedundantDiagType)
                            newERIList.Add("AntennaRedundantSubDiagonalType=" & .AntennaRedundantSubDiagonalType)
                            newERIList.Add("AntennaRedundantSubHorizontalType=" & .AntennaRedundantSubHorizontalType)
                            newERIList.Add("AntennaRedundantVerticalType=" & .AntennaRedundantVerticalType)
                            newERIList.Add("AntennaRedundantHipType=" & .AntennaRedundantHipType)
                            newERIList.Add("AntennaRedundantHipDiagonalType=" & .AntennaRedundantHipDiagonalType)
                            newERIList.Add("AntennaRedundantHorizontalSize=" & .AntennaRedundantHorizontalSize)
                            newERIList.Add("AntennaRedundantHorizontalSize2=" & .AntennaRedundantHorizontalSize2)
                            newERIList.Add("AntennaRedundantHorizontalSize3=" & .AntennaRedundantHorizontalSize3)
                            newERIList.Add("AntennaRedundantHorizontalSize4=" & .AntennaRedundantHorizontalSize4)
                            newERIList.Add("AntennaRedundantDiagonalSize=" & .AntennaRedundantDiagonalSize)
                            newERIList.Add("AntennaRedundantDiagonalSize2=" & .AntennaRedundantDiagonalSize2)
                            newERIList.Add("AntennaRedundantDiagonalSize3=" & .AntennaRedundantDiagonalSize3)
                            newERIList.Add("AntennaRedundantDiagonalSize4=" & .AntennaRedundantDiagonalSize4)
                            newERIList.Add("AntennaRedundantSubHorizontalSize=" & .AntennaRedundantSubHorizontalSize)
                            newERIList.Add("AntennaRedundantSubDiagonalSize=" & .AntennaRedundantSubDiagonalSize)
                            newERIList.Add("AntennaSubDiagLocation=" & .AntennaSubDiagLocation)
                            newERIList.Add("AntennaRedundantVerticalSize=" & .AntennaRedundantVerticalSize)
                            newERIList.Add("AntennaRedundantHipDiagonalSize=" & .AntennaRedundantHipDiagonalSize)
                            newERIList.Add("AntennaRedundantHipDiagonalSize2=" & .AntennaRedundantHipDiagonalSize2)
                            newERIList.Add("AntennaRedundantHipDiagonalSize3=" & .AntennaRedundantHipDiagonalSize3)
                            newERIList.Add("AntennaRedundantHipDiagonalSize4=" & .AntennaRedundantHipDiagonalSize4)
                            newERIList.Add("AntennaRedundantHipSize=" & .AntennaRedundantHipSize)
                            newERIList.Add("AntennaRedundantHipSize2=" & .AntennaRedundantHipSize2)
                            newERIList.Add("AntennaRedundantHipSize3=" & .AntennaRedundantHipSize3)
                            newERIList.Add("AntennaRedundantHipSize4=" & .AntennaRedundantHipSize4)
                            newERIList.Add("AntennaNumInnerGirts=" & .AntennaNumInnerGirts)
                            newERIList.Add("AntennaInnerGirtType=" & .AntennaInnerGirtType)
                            newERIList.Add("AntennaInnerGirtSize=" & .AntennaInnerGirtSize)
                            newERIList.Add("AntennaPoleShapeType=" & .AntennaPoleShapeType)
                            newERIList.Add("AntennaPoleSize=" & .AntennaPoleSize)
                            If Not IsNothing(Me.settings.USUnits.Strength.convertToERIUnits(.AntennaPoleGrade)) Then 'This field does not get saved when generated from CCIpole
                                newERIList.Add("AntennaPoleGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(.AntennaPoleGrade))
                            End If
                            If IsSomethingString(.AntennaPoleMatlGrade) Then 'This field does not get saved when generated from CCIpole
                                newERIList.Add("AntennaPoleMatlGrade=" & .AntennaPoleMatlGrade)
                            End If
                            newERIList.Add("AntennaPoleSpliceLength=" & .AntennaPoleSpliceLength)
                            newERIList.Add("AntennaTaperPoleNumSides=" & .AntennaTaperPoleNumSides)
                            newERIList.Add("AntennaTaperPoleTopDiameter=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaTaperPoleTopDiameter))
                            newERIList.Add("AntennaTaperPoleBotDiameter=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaTaperPoleBotDiameter))
                            newERIList.Add("AntennaTaperPoleWallThickness=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaTaperPoleWallThickness))
                            newERIList.Add("AntennaTaperPoleBendRadius=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaTaperPoleBendRadius))
                            If Not IsNothing(Me.settings.USUnits.Strength.convertToERIUnits(.AntennaTaperPoleGrade)) Then 'This field does not get saved when generated from CCIpole
                                newERIList.Add("AntennaTaperPoleGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(.AntennaTaperPoleGrade))
                            End If
                            If IsSomethingString(.AntennaTaperPoleMatlGrade) Then 'This field does not get saved when generated from CCIpole
                                newERIList.Add("AntennaTaperPoleMatlGrade=" & .AntennaTaperPoleMatlGrade)
                            End If
                            newERIList.Add("AntennaSWMult=" & .AntennaSWMult)
                            If Not IsNothing(.AntennaWPMult) Then 'This field does not get saved when generated from CCIpole. Defaults to 1.
                                newERIList.Add("AntennaWPMult=" & .AntennaWPMult)
                            End If
                            newERIList.Add("AntennaAutoCalcKSingleAngle=" & .AntennaAutoCalcKSingleAngle)
                            newERIList.Add("AntennaAutoCalcKSolidRound=" & .AntennaAutoCalcKSolidRound)
                            newERIList.Add("AntennaAfGusset=" & Me.settings.USUnits.Length.convertToERIUnits(.AntennaAfGusset))
                            newERIList.Add("AntennaTfGusset=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaTfGusset))
                            newERIList.Add("AntennaGussetBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaGussetBoltEdgeDistance))
                            If Not IsNothing(Me.settings.USUnits.Strength.convertToERIUnits(.AntennaGussetGrade)) Then 'This field does not get saved when generated from CCIpole
                                newERIList.Add("AntennaGussetGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(.AntennaGussetGrade))
                            End If
                            If IsSomethingString(.AntennaGussetMatlGrade) Then 'This field does not get saved when generated from CCIpole
                                newERIList.Add("AntennaGussetMatlGrade=" & .AntennaGussetMatlGrade)
                            End If
                            If Not IsNothing(.AntennaAfMult) Then 'This field does not get saved when generated from CCIpole. Defaults to 1.
                                newERIList.Add("AntennaAfMult=" & .AntennaAfMult)
                            End If
                            If Not IsNothing(.AntennaArMult) Then 'This field does not get saved when generated from CCIpole. Defaults to 1.
                                newERIList.Add("AntennaArMult=" & .AntennaArMult)
                            End If
                            newERIList.Add("AntennaFlatIPAPole=" & .AntennaFlatIPAPole)
                            newERIList.Add("AntennaRoundIPAPole=" & .AntennaRoundIPAPole)
                            newERIList.Add("AntennaFlatIPALeg=" & .AntennaFlatIPALeg)
                            newERIList.Add("AntennaRoundIPALeg=" & .AntennaRoundIPALeg)
                            newERIList.Add("AntennaFlatIPAHorizontal=" & .AntennaFlatIPAHorizontal)
                            newERIList.Add("AntennaRoundIPAHorizontal=" & .AntennaRoundIPAHorizontal)
                            newERIList.Add("AntennaFlatIPADiagonal=" & .AntennaFlatIPADiagonal)
                            newERIList.Add("AntennaRoundIPADiagonal=" & .AntennaRoundIPADiagonal)
                            newERIList.Add("AntennaCSA_S37_SpeedUpFactor=" & .AntennaCSA_S37_SpeedUpFactor)
                            newERIList.Add("AntennaKLegs=" & .AntennaKLegs)
                            newERIList.Add("AntennaKXBracedDiags=" & .AntennaKXBracedDiags)
                            newERIList.Add("AntennaKKBracedDiags=" & .AntennaKKBracedDiags)
                            newERIList.Add("AntennaKZBracedDiags=" & .AntennaKZBracedDiags)
                            newERIList.Add("AntennaKHorzs=" & .AntennaKHorzs)
                            newERIList.Add("AntennaKSecHorzs=" & .AntennaKSecHorzs)
                            newERIList.Add("AntennaKGirts=" & .AntennaKGirts)
                            newERIList.Add("AntennaKInners=" & .AntennaKInners)
                            newERIList.Add("AntennaKXBracedDiagsY=" & .AntennaKXBracedDiagsY)
                            newERIList.Add("AntennaKKBracedDiagsY=" & .AntennaKKBracedDiagsY)
                            newERIList.Add("AntennaKZBracedDiagsY=" & .AntennaKZBracedDiagsY)
                            newERIList.Add("AntennaKHorzsY=" & .AntennaKHorzsY)
                            newERIList.Add("AntennaKSecHorzsY=" & .AntennaKSecHorzsY)
                            newERIList.Add("AntennaKGirtsY=" & .AntennaKGirtsY)
                            newERIList.Add("AntennaKInnersY=" & .AntennaKInnersY)
                            newERIList.Add("AntennaKRedHorz=" & .AntennaKRedHorz)
                            newERIList.Add("AntennaKRedDiag=" & .AntennaKRedDiag)
                            newERIList.Add("AntennaKRedSubDiag=" & .AntennaKRedSubDiag)
                            newERIList.Add("AntennaKRedSubHorz=" & .AntennaKRedSubHorz)
                            newERIList.Add("AntennaKRedVert=" & .AntennaKRedVert)
                            newERIList.Add("AntennaKRedHip=" & .AntennaKRedHip)
                            newERIList.Add("AntennaKRedHipDiag=" & .AntennaKRedHipDiag)
                            newERIList.Add("AntennaKTLX=" & .AntennaKTLX)
                            newERIList.Add("AntennaKTLZ=" & .AntennaKTLZ)
                            newERIList.Add("AntennaKTLLeg=" & .AntennaKTLLeg)
                            newERIList.Add("AntennaInnerKTLX=" & .AntennaInnerKTLX)
                            newERIList.Add("AntennaInnerKTLZ=" & .AntennaInnerKTLZ)
                            newERIList.Add("AntennaInnerKTLLeg=" & .AntennaInnerKTLLeg)
                            newERIList.Add("AntennaStitchBoltLocationHoriz=" & .AntennaStitchBoltLocationHoriz)
                            newERIList.Add("AntennaStitchBoltLocationDiag=" & .AntennaStitchBoltLocationDiag)
                            newERIList.Add("AntennaStitchSpacing=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaStitchSpacing))
                            newERIList.Add("AntennaStitchSpacingHorz=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaStitchSpacingHorz))
                            newERIList.Add("AntennaStitchSpacingDiag=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaStitchSpacingDiag))
                            newERIList.Add("AntennaStitchSpacingRed=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaStitchSpacingRed))
                            newERIList.Add("AntennaLegNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaLegNetWidthDeduct))
                            newERIList.Add("AntennaLegUFactor=" & .AntennaLegUFactor)
                            newERIList.Add("AntennaDiagonalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaDiagonalNetWidthDeduct))
                            newERIList.Add("AntennaTopGirtNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaTopGirtNetWidthDeduct))
                            newERIList.Add("AntennaBotGirtNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaBotGirtNetWidthDeduct))
                            newERIList.Add("AntennaInnerGirtNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaInnerGirtNetWidthDeduct))
                            newERIList.Add("AntennaHorizontalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaHorizontalNetWidthDeduct))
                            newERIList.Add("AntennaShortHorizontalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaShortHorizontalNetWidthDeduct))
                            newERIList.Add("AntennaDiagonalUFactor=" & .AntennaDiagonalUFactor)
                            newERIList.Add("AntennaTopGirtUFactor=" & .AntennaTopGirtUFactor)
                            newERIList.Add("AntennaBotGirtUFactor=" & .AntennaBotGirtUFactor)
                            newERIList.Add("AntennaInnerGirtUFactor=" & .AntennaInnerGirtUFactor)
                            newERIList.Add("AntennaHorizontalUFactor=" & .AntennaHorizontalUFactor)
                            newERIList.Add("AntennaShortHorizontalUFactor=" & .AntennaShortHorizontalUFactor)
                            newERIList.Add("AntennaLegConnType=" & .AntennaLegConnType)
                            newERIList.Add("AntennaLegNumBolts=" & .AntennaLegNumBolts)
                            newERIList.Add("AntennaDiagonalNumBolts=" & .AntennaDiagonalNumBolts)
                            newERIList.Add("AntennaTopGirtNumBolts=" & .AntennaTopGirtNumBolts)
                            newERIList.Add("AntennaBotGirtNumBolts=" & .AntennaBotGirtNumBolts)
                            newERIList.Add("AntennaInnerGirtNumBolts=" & .AntennaInnerGirtNumBolts)
                            newERIList.Add("AntennaHorizontalNumBolts=" & .AntennaHorizontalNumBolts)
                            newERIList.Add("AntennaShortHorizontalNumBolts=" & .AntennaShortHorizontalNumBolts)
                            If IsSomethingString(.AntennaLegBoltGrade) Then 'This field does not get saved when generated from CCIpole.
                                newERIList.Add("AntennaLegBoltGrade=" & .AntennaLegBoltGrade)
                            End If
                            newERIList.Add("AntennaLegBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaLegBoltSize))
                            If IsSomethingString(.AntennaDiagonalBoltGrade) Then 'This field does not get saved when generated from CCIpole. 
                                newERIList.Add("AntennaDiagonalBoltGrade=" & .AntennaDiagonalBoltGrade)
                            End If
                            newERIList.Add("AntennaDiagonalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaDiagonalBoltSize))
                            If IsSomethingString(.AntennaTopGirtBoltGrade) Then 'This field does not get saved when generated from CCIpole. 
                                newERIList.Add("AntennaTopGirtBoltGrade=" & .AntennaTopGirtBoltGrade)
                            End If
                            newERIList.Add("AntennaTopGirtBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaTopGirtBoltSize))
                            If IsSomethingString(.AntennaBotGirtBoltGrade) Then 'This field does not get saved when generated from CCIpole. 
                                newERIList.Add("AntennaBotGirtBoltGrade=" & .AntennaBotGirtBoltGrade)
                            End If
                            newERIList.Add("AntennaBotGirtBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaBotGirtBoltSize))
                            If IsSomethingString(.AntennaInnerGirtBoltGrade) Then 'This field does not get saved when generated from CCIpole. 
                                newERIList.Add("AntennaInnerGirtBoltGrade=" & .AntennaInnerGirtBoltGrade)
                            End If
                            newERIList.Add("AntennaInnerGirtBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaInnerGirtBoltSize))
                            If IsSomethingString(.AntennaHorizontalBoltGrade) Then 'This field does not get saved when generated from CCIpole. 
                                newERIList.Add("AntennaHorizontalBoltGrade=" & .AntennaHorizontalBoltGrade)
                            End If
                            newERIList.Add("AntennaHorizontalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaHorizontalBoltSize))
                            If IsSomethingString(.AntennaShortHorizontalBoltGrade) Then 'This field does not get saved when generated from CCIpole. 
                                newERIList.Add("AntennaShortHorizontalBoltGrade=" & .AntennaShortHorizontalBoltGrade)
                            End If
                            newERIList.Add("AntennaShortHorizontalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaShortHorizontalBoltSize))
                            newERIList.Add("AntennaLegBoltEdgeDistance=" & .AntennaLegBoltEdgeDistance)
                            newERIList.Add("AntennaDiagonalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaDiagonalBoltEdgeDistance))
                            newERIList.Add("AntennaTopGirtBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaTopGirtBoltEdgeDistance))
                            newERIList.Add("AntennaBotGirtBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaBotGirtBoltEdgeDistance))
                            newERIList.Add("AntennaInnerGirtBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaInnerGirtBoltEdgeDistance))
                            newERIList.Add("AntennaHorizontalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaHorizontalBoltEdgeDistance))
                            newERIList.Add("AntennaShortHorizontalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaShortHorizontalBoltEdgeDistance))
                            newERIList.Add("AntennaDiagonalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaDiagonalGageG1Distance))
                            newERIList.Add("AntennaTopGirtGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaTopGirtGageG1Distance))
                            newERIList.Add("AntennaBotGirtGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaBotGirtGageG1Distance))
                            newERIList.Add("AntennaInnerGirtGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaInnerGirtGageG1Distance))
                            newERIList.Add("AntennaHorizontalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaHorizontalGageG1Distance))
                            newERIList.Add("AntennaShortHorizontalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaShortHorizontalGageG1Distance))
                            If IsSomethingString(.AntennaRedundantHorizontalBoltGrade) Then 'This field does not get saved when generated from CCIpole. 
                                newERIList.Add("AntennaRedundantHorizontalBoltGrade=" & .AntennaRedundantHorizontalBoltGrade)
                            End If
                            newERIList.Add("AntennaRedundantHorizontalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaRedundantHorizontalBoltSize))
                            newERIList.Add("AntennaRedundantHorizontalNumBolts=" & .AntennaRedundantHorizontalNumBolts)
                            newERIList.Add("AntennaRedundantHorizontalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaRedundantHorizontalBoltEdgeDistance))
                            newERIList.Add("AntennaRedundantHorizontalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaRedundantHorizontalGageG1Distance))
                            newERIList.Add("AntennaRedundantHorizontalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaRedundantHorizontalNetWidthDeduct))
                            newERIList.Add("AntennaRedundantHorizontalUFactor=" & .AntennaRedundantHorizontalUFactor)
                            If IsSomethingString(.AntennaRedundantDiagonalBoltGrade) Then 'This field does not get saved when generated from CCIpole. 
                                newERIList.Add("AntennaRedundantDiagonalBoltGrade=" & .AntennaRedundantDiagonalBoltGrade)
                            End If
                            newERIList.Add("AntennaRedundantDiagonalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaRedundantDiagonalBoltSize))
                            newERIList.Add("AntennaRedundantDiagonalNumBolts=" & .AntennaRedundantDiagonalNumBolts)
                            newERIList.Add("AntennaRedundantDiagonalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaRedundantDiagonalBoltEdgeDistance))
                            newERIList.Add("AntennaRedundantDiagonalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaRedundantDiagonalGageG1Distance))
                            newERIList.Add("AntennaRedundantDiagonalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaRedundantDiagonalNetWidthDeduct))
                            newERIList.Add("AntennaRedundantDiagonalUFactor=" & .AntennaRedundantDiagonalUFactor)
                            If IsSomethingString(.AntennaRedundantSubDiagonalBoltGrade) Then 'This field does not get saved when generated from CCIpole. 
                                newERIList.Add("AntennaRedundantSubDiagonalBoltGrade=" & .AntennaRedundantSubDiagonalBoltGrade)
                            End If
                            newERIList.Add("AntennaRedundantSubDiagonalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaRedundantSubDiagonalBoltSize))
                            newERIList.Add("AntennaRedundantSubDiagonalNumBolts=" & .AntennaRedundantSubDiagonalNumBolts)
                            newERIList.Add("AntennaRedundantSubDiagonalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaRedundantSubDiagonalBoltEdgeDistance))
                            newERIList.Add("AntennaRedundantSubDiagonalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaRedundantSubDiagonalGageG1Distance))
                            newERIList.Add("AntennaRedundantSubDiagonalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaRedundantSubDiagonalNetWidthDeduct))
                            newERIList.Add("AntennaRedundantSubDiagonalUFactor=" & .AntennaRedundantSubDiagonalUFactor)
                            If IsSomethingString(.AntennaRedundantSubHorizontalBoltGrade) Then 'This field does not get saved when generated from CCIpole. 
                                newERIList.Add("AntennaRedundantSubHorizontalBoltGrade=" & .AntennaRedundantSubHorizontalBoltGrade)
                            End If
                            newERIList.Add("AntennaRedundantSubHorizontalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaRedundantSubHorizontalBoltSize))
                            newERIList.Add("AntennaRedundantSubHorizontalNumBolts=" & .AntennaRedundantSubHorizontalNumBolts)
                            newERIList.Add("AntennaRedundantSubHorizontalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaRedundantSubHorizontalBoltEdgeDistance))
                            newERIList.Add("AntennaRedundantSubHorizontalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaRedundantSubHorizontalGageG1Distance))
                            newERIList.Add("AntennaRedundantSubHorizontalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaRedundantSubHorizontalNetWidthDeduct))
                            newERIList.Add("AntennaRedundantSubHorizontalUFactor=" & .AntennaRedundantSubHorizontalUFactor)
                            If IsSomethingString(.AntennaRedundantVerticalBoltGrade) Then 'This field does not get saved when generated from CCIpole. 
                                newERIList.Add("AntennaRedundantVerticalBoltGrade=" & .AntennaRedundantVerticalBoltGrade)
                            End If
                            newERIList.Add("AntennaRedundantVerticalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaRedundantVerticalBoltSize))
                            newERIList.Add("AntennaRedundantVerticalNumBolts=" & .AntennaRedundantVerticalNumBolts)
                            newERIList.Add("AntennaRedundantVerticalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaRedundantVerticalBoltEdgeDistance))
                            newERIList.Add("AntennaRedundantVerticalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaRedundantVerticalGageG1Distance))
                            newERIList.Add("AntennaRedundantVerticalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaRedundantVerticalNetWidthDeduct))
                            newERIList.Add("AntennaRedundantVerticalUFactor=" & .AntennaRedundantVerticalUFactor)
                            If IsSomethingString(.AntennaRedundantHipBoltGrade) Then 'This field does not get saved when generated from CCIpole. 
                                newERIList.Add("AntennaRedundantHipBoltGrade=" & .AntennaRedundantHipBoltGrade)
                            End If
                            newERIList.Add("AntennaRedundantHipBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaRedundantHipBoltSize))
                            newERIList.Add("AntennaRedundantHipNumBolts=" & .AntennaRedundantHipNumBolts)
                            newERIList.Add("AntennaRedundantHipBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaRedundantHipBoltEdgeDistance))
                            newERIList.Add("AntennaRedundantHipGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaRedundantHipGageG1Distance))
                            newERIList.Add("AntennaRedundantHipNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaRedundantHipNetWidthDeduct))
                            newERIList.Add("AntennaRedundantHipUFactor=" & .AntennaRedundantHipUFactor)
                            If IsSomethingString(.AntennaShortHorizontalBoltGrade) Then 'This field does not get saved when generated from CCIpole. 
                                newERIList.Add("AntennaRedundantHipDiagonalBoltGrade=" & .AntennaRedundantHipDiagonalBoltGrade)
                            End If
                            newERIList.Add("AntennaRedundantHipDiagonalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaRedundantHipDiagonalBoltSize))
                            newERIList.Add("AntennaRedundantHipDiagonalNumBolts=" & .AntennaRedundantHipDiagonalNumBolts)
                            newERIList.Add("AntennaRedundantHipDiagonalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaRedundantHipDiagonalBoltEdgeDistance))
                            newERIList.Add("AntennaRedundantHipDiagonalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaRedundantHipDiagonalGageG1Distance))
                            newERIList.Add("AntennaRedundantHipDiagonalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaRedundantHipDiagonalNetWidthDeduct))
                            newERIList.Add("AntennaRedundantHipDiagonalUFactor=" & .AntennaRedundantHipDiagonalUFactor)
                            newERIList.Add("AntennaDiagonalOutOfPlaneRestraint=" & trueFalseYesNo(.AntennaDiagonalOutOfPlaneRestraint))
                            newERIList.Add("AntennaTopGirtOutOfPlaneRestraint=" & trueFalseYesNo(.AntennaTopGirtOutOfPlaneRestraint))
                            newERIList.Add("AntennaBottomGirtOutOfPlaneRestraint=" & trueFalseYesNo(.AntennaBottomGirtOutOfPlaneRestraint))
                            newERIList.Add("AntennaMidGirtOutOfPlaneRestraint=" & trueFalseYesNo(.AntennaMidGirtOutOfPlaneRestraint))
                            newERIList.Add("AntennaHorizontalOutOfPlaneRestraint=" & trueFalseYesNo(.AntennaHorizontalOutOfPlaneRestraint))
                            newERIList.Add("AntennaSecondaryHorizontalOutOfPlaneRestraint=" & trueFalseYesNo(.AntennaSecondaryHorizontalOutOfPlaneRestraint))
                            newERIList.Add("AntennaDiagOffsetNEY=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaDiagOffsetNEY))
                            newERIList.Add("AntennaDiagOffsetNEX=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaDiagOffsetNEX))
                            newERIList.Add("AntennaDiagOffsetPEY=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaDiagOffsetPEY))
                            newERIList.Add("AntennaDiagOffsetPEX=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaDiagOffsetPEX))
                            newERIList.Add("AntennaKbraceOffsetNEY=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaKbraceOffsetNEY))
                            newERIList.Add("AntennaKbraceOffsetNEX=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaKbraceOffsetNEX))
                            newERIList.Add("AntennaKbraceOffsetPEY=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaKbraceOffsetPEY))
                            newERIList.Add("AntennaKbraceOffsetPEX=" & Me.settings.USUnits.Properties.convertToERIUnits(.AntennaKbraceOffsetPEX))
                        End With
                    Next i

#Region "Supp Abracadabruh"
                    newERIList.AddRange(defaults.Supp)
#End Region

                Case line(0).Equals("NumTowerRecs")
                    newERIList.Add(line(0) & "=" & Me.geometry.baseStructure.Count)
                    For i = 0 To Me.geometry.baseStructure.Count - 1
                        With Me.geometry.baseStructure(i)
                            newERIList.Add("TowerRec=" & .Rec)
                            newERIList.Add("TowerDatabase=" & .TowerDatabase)
                            newERIList.Add("TowerName=" & .TowerName)
                            newERIList.Add("TowerHeight=" & Me.settings.USUnits.Coordinate.convertToERIUnits(.TowerHeight))
                            newERIList.Add("TowerFaceWidth=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerFaceWidth))
                            newERIList.Add("TowerNumSections=" & .TowerNumSections)
                            newERIList.Add("TowerSectionLength=" & Me.settings.USUnits.Length.convertToERIUnits(.TowerSectionLength))
                            newERIList.Add("TowerDiagonalSpacing=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerDiagonalSpacing))
                            newERIList.Add("TowerDiagonalSpacingEx=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerDiagonalSpacingEx))
                            newERIList.Add("TowerBraceType=" & .TowerBraceType)
                            newERIList.Add("TowerFaceBevel=" & .TowerFaceBevel)
                            newERIList.Add("TowerTopGirtOffset=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerTopGirtOffset))
                            newERIList.Add("TowerBotGirtOffset=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerBotGirtOffset))
                            newERIList.Add("TowerHasKBraceEndPanels=" & trueFalseYesNo(.TowerHasKBraceEndPanels))
                            newERIList.Add("TowerHasHorizontals=" & trueFalseYesNo(.TowerHasHorizontals))
                            newERIList.Add("TowerLegType=" & .TowerLegType)
                            newERIList.Add("TowerLegSize=" & .TowerLegSize)
                            newERIList.Add("TowerLegGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(.TowerLegGrade))
                            newERIList.Add("TowerLegMatlGrade=" & .TowerLegMatlGrade)
                            newERIList.Add("TowerDiagonalGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(.TowerDiagonalGrade))
                            newERIList.Add("TowerDiagonalMatlGrade=" & .TowerDiagonalMatlGrade)
                            newERIList.Add("TowerInnerBracingGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(.TowerInnerBracingGrade))
                            newERIList.Add("TowerInnerBracingMatlGrade=" & .TowerInnerBracingMatlGrade)
                            newERIList.Add("TowerTopGirtGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(.TowerTopGirtGrade))
                            newERIList.Add("TowerTopGirtMatlGrade=" & .TowerTopGirtMatlGrade)
                            newERIList.Add("TowerBotGirtGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(.TowerBotGirtGrade))
                            newERIList.Add("TowerBotGirtMatlGrade=" & .TowerBotGirtMatlGrade)
                            newERIList.Add("TowerInnerGirtGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(.TowerInnerGirtGrade))
                            newERIList.Add("TowerInnerGirtMatlGrade=" & .TowerInnerGirtMatlGrade)
                            newERIList.Add("TowerLongHorizontalGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(.TowerLongHorizontalGrade))
                            newERIList.Add("TowerLongHorizontalMatlGrade=" & .TowerLongHorizontalMatlGrade)
                            newERIList.Add("TowerShortHorizontalGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(.TowerShortHorizontalGrade))
                            newERIList.Add("TowerShortHorizontalMatlGrade=" & .TowerShortHorizontalMatlGrade)
                            newERIList.Add("TowerDiagonalType=" & .TowerDiagonalType)
                            newERIList.Add("TowerDiagonalSize=" & .TowerDiagonalSize)
                            newERIList.Add("TowerInnerBracingType=" & .TowerInnerBracingType)
                            newERIList.Add("TowerInnerBracingSize=" & .TowerInnerBracingSize)
                            newERIList.Add("TowerTopGirtType=" & .TowerTopGirtType)
                            newERIList.Add("TowerTopGirtSize=" & .TowerTopGirtSize)
                            newERIList.Add("TowerBotGirtType=" & .TowerBotGirtType)
                            newERIList.Add("TowerBotGirtSize=" & .TowerBotGirtSize)
                            newERIList.Add("TowerNumInnerGirts=" & .TowerNumInnerGirts)
                            newERIList.Add("TowerInnerGirtType=" & .TowerInnerGirtType)
                            newERIList.Add("TowerInnerGirtSize=" & .TowerInnerGirtSize)
                            newERIList.Add("TowerLongHorizontalType=" & .TowerLongHorizontalType)
                            newERIList.Add("TowerLongHorizontalSize=" & .TowerLongHorizontalSize)
                            newERIList.Add("TowerShortHorizontalType=" & .TowerShortHorizontalType)
                            newERIList.Add("TowerShortHorizontalSize=" & .TowerShortHorizontalSize)
                            newERIList.Add("TowerRedundantGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(.TowerRedundantGrade))
                            newERIList.Add("TowerRedundantMatlGrade=" & .TowerRedundantMatlGrade)
                            newERIList.Add("TowerRedundantType=" & .TowerRedundantType)
                            newERIList.Add("TowerRedundantDiagType=" & .TowerRedundantDiagType)
                            newERIList.Add("TowerRedundantSubDiagonalType=" & .TowerRedundantSubDiagonalType)
                            newERIList.Add("TowerRedundantSubHorizontalType=" & .TowerRedundantSubHorizontalType)
                            newERIList.Add("TowerRedundantVerticalType=" & .TowerRedundantVerticalType)
                            newERIList.Add("TowerRedundantHipType=" & .TowerRedundantHipType)
                            newERIList.Add("TowerRedundantHipDiagonalType=" & .TowerRedundantHipDiagonalType)
                            newERIList.Add("TowerRedundantHorizontalSize=" & .TowerRedundantHorizontalSize)
                            newERIList.Add("TowerRedundantHorizontalSize2=" & .TowerRedundantHorizontalSize2)
                            newERIList.Add("TowerRedundantHorizontalSize3=" & .TowerRedundantHorizontalSize3)
                            newERIList.Add("TowerRedundantHorizontalSize4=" & .TowerRedundantHorizontalSize4)
                            newERIList.Add("TowerRedundantDiagonalSize=" & .TowerRedundantDiagonalSize)
                            newERIList.Add("TowerRedundantDiagonalSize2=" & .TowerRedundantDiagonalSize2)
                            newERIList.Add("TowerRedundantDiagonalSize3=" & .TowerRedundantDiagonalSize3)
                            newERIList.Add("TowerRedundantDiagonalSize4=" & .TowerRedundantDiagonalSize4)
                            newERIList.Add("TowerRedundantSubHorizontalSize=" & .TowerRedundantSubHorizontalSize)
                            newERIList.Add("TowerRedundantSubDiagonalSize=" & .TowerRedundantSubDiagonalSize)
                            newERIList.Add("TowerSubDiagLocation=" & .TowerSubDiagLocation)
                            newERIList.Add("TowerRedundantVerticalSize=" & .TowerRedundantVerticalSize)
                            newERIList.Add("TowerRedundantHipSize=" & .TowerRedundantHipSize)
                            newERIList.Add("TowerRedundantHipSize2=" & .TowerRedundantHipSize2)
                            newERIList.Add("TowerRedundantHipSize3=" & .TowerRedundantHipSize3)
                            newERIList.Add("TowerRedundantHipSize4=" & .TowerRedundantHipSize4)
                            newERIList.Add("TowerRedundantHipDiagonalSize=" & .TowerRedundantHipDiagonalSize)
                            newERIList.Add("TowerRedundantHipDiagonalSize2=" & .TowerRedundantHipDiagonalSize2)
                            newERIList.Add("TowerRedundantHipDiagonalSize3=" & .TowerRedundantHipDiagonalSize3)
                            newERIList.Add("TowerRedundantHipDiagonalSize4=" & .TowerRedundantHipDiagonalSize4)
                            newERIList.Add("TowerSWMult=" & .TowerSWMult)
                            newERIList.Add("TowerWPMult=" & .TowerWPMult)
                            newERIList.Add("TowerAutoCalcKSingleAngle=" & trueFalseYesNo(.TowerAutoCalcKSingleAngle))
                            newERIList.Add("TowerAutoCalcKSolidRound=" & trueFalseYesNo(.TowerAutoCalcKSolidRound))
                            newERIList.Add("TowerAfGusset=" & Me.settings.USUnits.Length.convertToERIUnits(.TowerAfGusset))
                            newERIList.Add("TowerTfGusset=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerTfGusset))
                            newERIList.Add("TowerGussetBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerGussetBoltEdgeDistance))
                            newERIList.Add("TowerGussetGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(.TowerGussetGrade))
                            newERIList.Add("TowerGussetMatlGrade=" & .TowerGussetMatlGrade)
                            newERIList.Add("TowerAfMult=" & .TowerAfMult)
                            newERIList.Add("TowerArMult=" & .TowerArMult)
                            newERIList.Add("TowerFlatIPAPole=" & .TowerFlatIPAPole)
                            newERIList.Add("TowerRoundIPAPole=" & .TowerRoundIPAPole)
                            newERIList.Add("TowerFlatIPALeg=" & .TowerFlatIPALeg)
                            newERIList.Add("TowerRoundIPALeg=" & .TowerRoundIPALeg)
                            newERIList.Add("TowerFlatIPAHorizontal=" & .TowerFlatIPAHorizontal)
                            newERIList.Add("TowerRoundIPAHorizontal=" & .TowerRoundIPAHorizontal)
                            newERIList.Add("TowerFlatIPADiagonal=" & .TowerFlatIPADiagonal)
                            newERIList.Add("TowerRoundIPADiagonal=" & .TowerRoundIPADiagonal)
                            newERIList.Add("TowerCSA_S37_SpeedUpFactor=" & .TowerCSA_S37_SpeedUpFactor)
                            newERIList.Add("TowerKLegs=" & .TowerKLegs)
                            newERIList.Add("TowerKXBracedDiags=" & .TowerKXBracedDiags)
                            newERIList.Add("TowerKKBracedDiags=" & .TowerKKBracedDiags)
                            newERIList.Add("TowerKZBracedDiags=" & .TowerKZBracedDiags)
                            newERIList.Add("TowerKHorzs=" & .TowerKHorzs)
                            newERIList.Add("TowerKSecHorzs=" & .TowerKSecHorzs)
                            newERIList.Add("TowerKGirts=" & .TowerKGirts)
                            newERIList.Add("TowerKInners=" & .TowerKInners)
                            newERIList.Add("TowerKXBracedDiagsY=" & .TowerKXBracedDiagsY)
                            newERIList.Add("TowerKKBracedDiagsY=" & .TowerKKBracedDiagsY)
                            newERIList.Add("TowerKZBracedDiagsY=" & .TowerKZBracedDiagsY)
                            newERIList.Add("TowerKHorzsY=" & .TowerKHorzsY)
                            newERIList.Add("TowerKSecHorzsY=" & .TowerKSecHorzsY)
                            newERIList.Add("TowerKGirtsY=" & .TowerKGirtsY)
                            newERIList.Add("TowerKInnersY=" & .TowerKInnersY)
                            newERIList.Add("TowerKRedHorz=" & .TowerKRedHorz)
                            newERIList.Add("TowerKRedDiag=" & .TowerKRedDiag)
                            newERIList.Add("TowerKRedSubDiag=" & .TowerKRedSubDiag)
                            newERIList.Add("TowerKRedSubHorz=" & .TowerKRedSubHorz)
                            newERIList.Add("TowerKRedVert=" & .TowerKRedVert)
                            newERIList.Add("TowerKRedHip=" & .TowerKRedHip)
                            newERIList.Add("TowerKRedHipDiag=" & .TowerKRedHipDiag)
                            newERIList.Add("TowerKTLX=" & .TowerKTLX)
                            newERIList.Add("TowerKTLZ=" & .TowerKTLZ)
                            newERIList.Add("TowerKTLLeg=" & .TowerKTLLeg)
                            newERIList.Add("TowerInnerKTLX=" & .TowerInnerKTLX)
                            newERIList.Add("TowerInnerKTLZ=" & .TowerInnerKTLZ)
                            newERIList.Add("TowerInnerKTLLeg=" & .TowerInnerKTLLeg)
                            newERIList.Add("TowerStitchBoltLocationHoriz=" & .TowerStitchBoltLocationHoriz)
                            newERIList.Add("TowerStitchBoltLocationDiag=" & .TowerStitchBoltLocationDiag)
                            newERIList.Add("TowerStitchBoltLocationRed=" & .TowerStitchBoltLocationRed)
                            newERIList.Add("TowerStitchSpacing=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerStitchSpacing))
                            newERIList.Add("TowerStitchSpacingDiag=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerStitchSpacingDiag))
                            newERIList.Add("TowerStitchSpacingHorz=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerStitchSpacingHorz))
                            newERIList.Add("TowerStitchSpacingRed=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerStitchSpacingRed))
                            newERIList.Add("TowerLegNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerLegNetWidthDeduct))
                            newERIList.Add("TowerLegUFactor=" & .TowerLegUFactor)
                            newERIList.Add("TowerDiagonalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerDiagonalNetWidthDeduct))
                            newERIList.Add("TowerTopGirtNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerTopGirtNetWidthDeduct))
                            newERIList.Add("TowerBotGirtNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerBotGirtNetWidthDeduct))
                            newERIList.Add("TowerInnerGirtNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerInnerGirtNetWidthDeduct))
                            newERIList.Add("TowerHorizontalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerHorizontalNetWidthDeduct))
                            newERIList.Add("TowerShortHorizontalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerShortHorizontalNetWidthDeduct))
                            newERIList.Add("TowerDiagonalUFactor=" & .TowerDiagonalUFactor)
                            newERIList.Add("TowerTopGirtUFactor=" & .TowerTopGirtUFactor)
                            newERIList.Add("TowerBotGirtUFactor=" & .TowerBotGirtUFactor)
                            newERIList.Add("TowerInnerGirtUFactor=" & .TowerInnerGirtUFactor)
                            newERIList.Add("TowerHorizontalUFactor=" & .TowerHorizontalUFactor)
                            newERIList.Add("TowerShortHorizontalUFactor=" & .TowerShortHorizontalUFactor)
                            newERIList.Add("TowerLegConnType=" & .TowerLegConnType)
                            newERIList.Add("TowerLegNumBolts=" & .TowerLegNumBolts)
                            newERIList.Add("TowerDiagonalNumBolts=" & .TowerDiagonalNumBolts)
                            newERIList.Add("TowerTopGirtNumBolts=" & .TowerTopGirtNumBolts)
                            newERIList.Add("TowerBotGirtNumBolts=" & .TowerBotGirtNumBolts)
                            newERIList.Add("TowerInnerGirtNumBolts=" & .TowerInnerGirtNumBolts)
                            newERIList.Add("TowerHorizontalNumBolts=" & .TowerHorizontalNumBolts)
                            newERIList.Add("TowerShortHorizontalNumBolts=" & .TowerShortHorizontalNumBolts)
                            newERIList.Add("TowerLegBoltGrade=" & .TowerLegBoltGrade)
                            newERIList.Add("TowerLegBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerLegBoltSize))
                            newERIList.Add("TowerDiagonalBoltGrade=" & .TowerDiagonalBoltGrade)
                            newERIList.Add("TowerDiagonalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerDiagonalBoltSize))
                            newERIList.Add("TowerTopGirtBoltGrade=" & .TowerTopGirtBoltGrade)
                            newERIList.Add("TowerTopGirtBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerTopGirtBoltSize))
                            newERIList.Add("TowerBotGirtBoltGrade=" & .TowerBotGirtBoltGrade)
                            newERIList.Add("TowerBotGirtBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerBotGirtBoltSize))
                            newERIList.Add("TowerInnerGirtBoltGrade=" & .TowerInnerGirtBoltGrade)
                            newERIList.Add("TowerInnerGirtBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerInnerGirtBoltSize))
                            newERIList.Add("TowerHorizontalBoltGrade=" & .TowerHorizontalBoltGrade)
                            newERIList.Add("TowerHorizontalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerHorizontalBoltSize))
                            newERIList.Add("TowerShortHorizontalBoltGrade=" & .TowerShortHorizontalBoltGrade)
                            newERIList.Add("TowerShortHorizontalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerShortHorizontalBoltSize))
                            newERIList.Add("TowerLegBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerLegBoltEdgeDistance))
                            newERIList.Add("TowerDiagonalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerDiagonalBoltEdgeDistance))
                            newERIList.Add("TowerTopGirtBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerTopGirtBoltEdgeDistance))
                            newERIList.Add("TowerBotGirtBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerBotGirtBoltEdgeDistance))
                            newERIList.Add("TowerInnerGirtBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerInnerGirtBoltEdgeDistance))
                            newERIList.Add("TowerHorizontalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerHorizontalBoltEdgeDistance))
                            newERIList.Add("TowerShortHorizontalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerShortHorizontalBoltEdgeDistance))
                            newERIList.Add("TowerDiagonalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerDiagonalGageG1Distance))
                            newERIList.Add("TowerTopGirtGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerTopGirtGageG1Distance))
                            newERIList.Add("TowerBotGirtGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerBotGirtGageG1Distance))
                            newERIList.Add("TowerInnerGirtGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerInnerGirtGageG1Distance))
                            newERIList.Add("TowerHorizontalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerHorizontalGageG1Distance))
                            newERIList.Add("TowerShortHorizontalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerShortHorizontalGageG1Distance))
                            newERIList.Add("TowerRedundantHorizontalBoltGrade=" & .TowerRedundantHorizontalBoltGrade)
                            newERIList.Add("TowerRedundantHorizontalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerRedundantHorizontalBoltSize))
                            newERIList.Add("TowerRedundantHorizontalNumBolts=" & .TowerRedundantHorizontalNumBolts)
                            newERIList.Add("TowerRedundantHorizontalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerRedundantHorizontalBoltEdgeDistance))
                            newERIList.Add("TowerRedundantHorizontalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerRedundantHorizontalGageG1Distance))
                            newERIList.Add("TowerRedundantHorizontalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerRedundantHorizontalNetWidthDeduct))
                            newERIList.Add("TowerRedundantHorizontalUFactor=" & .TowerRedundantHorizontalUFactor)
                            newERIList.Add("TowerRedundantDiagonalBoltGrade=" & .TowerRedundantDiagonalBoltGrade)
                            newERIList.Add("TowerRedundantDiagonalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerRedundantDiagonalBoltSize))
                            newERIList.Add("TowerRedundantDiagonalNumBolts=" & .TowerRedundantDiagonalNumBolts)
                            newERIList.Add("TowerRedundantDiagonalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerRedundantDiagonalBoltEdgeDistance))
                            newERIList.Add("TowerRedundantDiagonalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerRedundantDiagonalGageG1Distance))
                            newERIList.Add("TowerRedundantDiagonalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerRedundantDiagonalNetWidthDeduct))
                            newERIList.Add("TowerRedundantDiagonalUFactor=" & .TowerRedundantDiagonalUFactor)
                            newERIList.Add("TowerRedundantSubDiagonalBoltGrade=" & .TowerRedundantSubDiagonalBoltGrade)
                            newERIList.Add("TowerRedundantSubDiagonalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerRedundantSubDiagonalBoltSize))
                            newERIList.Add("TowerRedundantSubDiagonalNumBolts=" & .TowerRedundantSubDiagonalNumBolts)
                            newERIList.Add("TowerRedundantSubDiagonalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerRedundantSubDiagonalBoltEdgeDistance))
                            newERIList.Add("TowerRedundantSubDiagonalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerRedundantSubDiagonalGageG1Distance))
                            newERIList.Add("TowerRedundantSubDiagonalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerRedundantSubDiagonalNetWidthDeduct))
                            newERIList.Add("TowerRedundantSubDiagonalUFactor=" & .TowerRedundantSubDiagonalUFactor)
                            newERIList.Add("TowerRedundantSubHorizontalBoltGrade=" & .TowerRedundantSubHorizontalBoltGrade)
                            newERIList.Add("TowerRedundantSubHorizontalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerRedundantSubHorizontalBoltSize))
                            newERIList.Add("TowerRedundantSubHorizontalNumBolts=" & .TowerRedundantSubHorizontalNumBolts)
                            newERIList.Add("TowerRedundantSubHorizontalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerRedundantSubHorizontalBoltEdgeDistance))
                            newERIList.Add("TowerRedundantSubHorizontalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerRedundantSubHorizontalGageG1Distance))
                            newERIList.Add("TowerRedundantSubHorizontalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerRedundantSubHorizontalNetWidthDeduct))
                            newERIList.Add("TowerRedundantSubHorizontalUFactor=" & .TowerRedundantSubHorizontalUFactor)
                            newERIList.Add("TowerRedundantVerticalBoltGrade=" & .TowerRedundantVerticalBoltGrade)
                            newERIList.Add("TowerRedundantVerticalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerRedundantVerticalBoltSize))
                            newERIList.Add("TowerRedundantVerticalNumBolts=" & .TowerRedundantVerticalNumBolts)
                            newERIList.Add("TowerRedundantVerticalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerRedundantVerticalBoltEdgeDistance))
                            newERIList.Add("TowerRedundantVerticalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerRedundantVerticalGageG1Distance))
                            newERIList.Add("TowerRedundantVerticalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerRedundantVerticalNetWidthDeduct))
                            newERIList.Add("TowerRedundantVerticalUFactor=" & .TowerRedundantVerticalUFactor)
                            newERIList.Add("TowerRedundantHipBoltGrade=" & .TowerRedundantHipBoltGrade)
                            newERIList.Add("TowerRedundantHipBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerRedundantHipBoltSize))
                            newERIList.Add("TowerRedundantHipNumBolts=" & .TowerRedundantHipNumBolts)
                            newERIList.Add("TowerRedundantHipBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerRedundantHipBoltEdgeDistance))
                            newERIList.Add("TowerRedundantHipGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerRedundantHipGageG1Distance))
                            newERIList.Add("TowerRedundantHipNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerRedundantHipNetWidthDeduct))
                            newERIList.Add("TowerRedundantHipUFactor=" & .TowerRedundantHipUFactor)
                            newERIList.Add("TowerRedundantHipDiagonalBoltGrade=" & .TowerRedundantHipDiagonalBoltGrade)
                            newERIList.Add("TowerRedundantHipDiagonalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerRedundantHipDiagonalBoltSize))
                            newERIList.Add("TowerRedundantHipDiagonalNumBolts=" & .TowerRedundantHipDiagonalNumBolts)
                            newERIList.Add("TowerRedundantHipDiagonalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerRedundantHipDiagonalBoltEdgeDistance))
                            newERIList.Add("TowerRedundantHipDiagonalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerRedundantHipDiagonalGageG1Distance))
                            newERIList.Add("TowerRedundantHipDiagonalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerRedundantHipDiagonalNetWidthDeduct))
                            newERIList.Add("TowerRedundantHipDiagonalUFactor=" & .TowerRedundantHipDiagonalUFactor)
                            newERIList.Add("TowerDiagonalOutOfPlaneRestraint=" & trueFalseYesNo(.TowerDiagonalOutOfPlaneRestraint))
                            newERIList.Add("TowerTopGirtOutOfPlaneRestraint=" & trueFalseYesNo(.TowerTopGirtOutOfPlaneRestraint))
                            newERIList.Add("TowerBottomGirtOutOfPlaneRestraint=" & trueFalseYesNo(.TowerBottomGirtOutOfPlaneRestraint))
                            newERIList.Add("TowerMidGirtOutOfPlaneRestraint=" & trueFalseYesNo(.TowerMidGirtOutOfPlaneRestraint))
                            newERIList.Add("TowerHorizontalOutOfPlaneRestraint=" & trueFalseYesNo(.TowerHorizontalOutOfPlaneRestraint))
                            newERIList.Add("TowerSecondaryHorizontalOutOfPlaneRestraint=" & trueFalseYesNo(.TowerSecondaryHorizontalOutOfPlaneRestraint))
                            newERIList.Add("TowerUniqueFlag=" & .TowerUniqueFlag)
                            newERIList.Add("TowerDiagOffsetNEY=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerDiagOffsetNEY))
                            newERIList.Add("TowerDiagOffsetNEX=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerDiagOffsetNEX))
                            newERIList.Add("TowerDiagOffsetPEY=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerDiagOffsetPEY))
                            newERIList.Add("TowerDiagOffsetPEX=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerDiagOffsetPEX))
                            newERIList.Add("TowerKbraceOffsetNEY=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerKbraceOffsetNEY))
                            newERIList.Add("TowerKbraceOffsetNEX=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerKbraceOffsetNEX))
                            newERIList.Add("TowerKbraceOffsetPEY=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerKbraceOffsetPEY))
                            newERIList.Add("TowerKbraceOffsetPEX=" & Me.settings.USUnits.Properties.convertToERIUnits(.TowerKbraceOffsetPEX))
                        End With
                    Next i
                Case line(0).Equals("NumGuyRecs")
                    newERIList.Add(line(0) & "=" & Me.geometry.guyWires.Count)
                    For i = 0 To Me.geometry.guyWires.Count - 1
                        With Me.geometry.guyWires(i)
                            newERIList.Add("GuyRec=" & .Rec)
                            newERIList.Add("GuyHeight=" & Me.settings.USUnits.Coordinate.convertToERIUnits(.GuyHeight))
                            newERIList.Add("GuyAutoCalcKSingleAngle=" & trueFalseYesNo(.GuyAutoCalcKSingleAngle))
                            newERIList.Add("GuyAutoCalcKSolidRound=" & trueFalseYesNo(.GuyAutoCalcKSolidRound))
                            newERIList.Add("GuyMount=" & .GuyMount)
                            newERIList.Add("TorqueArmStyle=" & .TorqueArmStyle)
                            newERIList.Add("GuyRadius=" & Me.settings.USUnits.Length.convertToERIUnits(.GuyRadius))
                            newERIList.Add("GuyRadius120=" & Me.settings.USUnits.Length.convertToERIUnits(.GuyRadius120))
                            newERIList.Add("GuyRadius240=" & Me.settings.USUnits.Length.convertToERIUnits(.GuyRadius240))
                            newERIList.Add("GuyRadius360=" & Me.settings.USUnits.Length.convertToERIUnits(.GuyRadius360))
                            newERIList.Add("TorqueArmRadius=" & Me.settings.USUnits.Length.convertToERIUnits(.TorqueArmRadius))
                            newERIList.Add("TorqueArmLegAngle=" & Me.settings.USUnits.Rotation.convertToERIUnits(.TorqueArmLegAngle))
                            newERIList.Add("Azimuth0Adjustment=" & Me.settings.USUnits.Rotation.convertToERIUnits(.Azimuth0Adjustment))
                            newERIList.Add("Azimuth120Adjustment=" & Me.settings.USUnits.Rotation.convertToERIUnits(.Azimuth120Adjustment))
                            newERIList.Add("Azimuth240Adjustment=" & Me.settings.USUnits.Rotation.convertToERIUnits(.Azimuth240Adjustment))
                            newERIList.Add("Azimuth360Adjustment=" & Me.settings.USUnits.Rotation.convertToERIUnits(.Azimuth360Adjustment))
                            newERIList.Add("Anchor0Elevation=" & Me.settings.USUnits.Coordinate.convertToERIUnits(.Anchor0Elevation))
                            newERIList.Add("Anchor120Elevation=" & Me.settings.USUnits.Coordinate.convertToERIUnits(.Anchor120Elevation))
                            newERIList.Add("Anchor240Elevation=" & Me.settings.USUnits.Coordinate.convertToERIUnits(.Anchor240Elevation))
                            newERIList.Add("Anchor360Elevation=" & Me.settings.USUnits.Coordinate.convertToERIUnits(.Anchor360Elevation))
                            newERIList.Add("GuySize=" & .GuySize)
                            newERIList.Add("Guy120Size=" & .Guy120Size)
                            newERIList.Add("Guy240Size=" & .Guy240Size)
                            newERIList.Add("Guy360Size=" & .Guy360Size)
                            newERIList.Add("GuyGrade=" & .GuyGrade)
                            newERIList.Add("TorqueArmSize=" & .TorqueArmSize)
                            newERIList.Add("TorqueArmSizeBot=" & .TorqueArmSizeBot)
                            newERIList.Add("TorqueArmType=" & .TorqueArmType)
                            newERIList.Add("TorqueArmGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(.TorqueArmGrade))
                            newERIList.Add("TorqueArmMatlGrade=" & .TorqueArmMatlGrade)
                            newERIList.Add("TorqueArmKFactor=" & .TorqueArmKFactor)
                            newERIList.Add("TorqueArmKFactorY=" & .TorqueArmKFactorY)
                            newERIList.Add("GuyPullOffKFactorX=" & .GuyPullOffKFactorX)
                            newERIList.Add("GuyPullOffKFactorY=" & .GuyPullOffKFactorY)
                            newERIList.Add("GuyDiagKFactorX=" & .GuyDiagKFactorX)
                            newERIList.Add("GuyDiagKFactorY=" & .GuyDiagKFactorY)
                            newERIList.Add("GuyAutoCalc=" & trueFalseYesNo(.GuyAutoCalc))
                            newERIList.Add("GuyAllGuysSame=" & trueFalseYesNo(.GuyAllGuysSame))
                            newERIList.Add("GuyAllGuysAnchorSame=" & trueFalseYesNo(.GuyAllGuysAnchorSame))
                            newERIList.Add("GuyIsStrapping=" & trueFalseYesNo(.GuyIsStrapping))
                            newERIList.Add("GuyPullOffSize=" & .GuyPullOffSize)
                            newERIList.Add("GuyPullOffSizeBot=" & .GuyPullOffSizeBot)
                            newERIList.Add("GuyPullOffType=" & .GuyPullOffType)
                            newERIList.Add("GuyPullOffGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(.GuyPullOffGrade))
                            newERIList.Add("GuyPullOffMatlGrade=" & .GuyPullOffMatlGrade)
                            newERIList.Add("GuyUpperDiagSize=" & .GuyUpperDiagSize)
                            newERIList.Add("GuyLowerDiagSize=" & .GuyLowerDiagSize)
                            newERIList.Add("GuyDiagType=" & .GuyDiagType)
                            newERIList.Add("GuyDiagGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(.GuyDiagGrade))
                            newERIList.Add("GuyDiagMatlGrade=" & .GuyDiagMatlGrade)
                            newERIList.Add("GuyDiagNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(.GuyDiagNetWidthDeduct))
                            newERIList.Add("GuyDiagUFactor=" & .GuyDiagUFactor)
                            newERIList.Add("GuyDiagNumBolts=" & .GuyDiagNumBolts)
                            newERIList.Add("GuyDiagonalOutOfPlaneRestraint=" & trueFalseYesNo(.GuyDiagonalOutOfPlaneRestraint))
                            newERIList.Add("GuyDiagBoltGrade=" & .GuyDiagBoltGrade)
                            newERIList.Add("GuyDiagBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(.GuyDiagBoltSize))
                            newERIList.Add("GuyDiagBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.GuyDiagBoltEdgeDistance))
                            newERIList.Add("GuyDiagBoltGageDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.GuyDiagBoltGageDistance))
                            newERIList.Add("GuyPullOffNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(.GuyPullOffNetWidthDeduct))
                            newERIList.Add("GuyPullOffUFactor=" & .GuyPullOffUFactor)
                            newERIList.Add("GuyPullOffNumBolts=" & .GuyPullOffNumBolts)
                            newERIList.Add("GuyPullOffOutOfPlaneRestraint=" & trueFalseYesNo(.GuyPullOffOutOfPlaneRestraint))
                            newERIList.Add("GuyPullOffBoltGrade=" & .GuyPullOffBoltGrade)
                            newERIList.Add("GuyPullOffBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(.GuyPullOffBoltSize))
                            newERIList.Add("GuyPullOffBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.GuyPullOffBoltEdgeDistance))
                            newERIList.Add("GuyPullOffBoltGageDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.GuyPullOffBoltGageDistance))
                            newERIList.Add("GuyTorqueArmNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(.GuyTorqueArmNetWidthDeduct))
                            newERIList.Add("GuyTorqueArmUFactor=" & .GuyTorqueArmUFactor)
                            newERIList.Add("GuyTorqueArmNumBolts=" & .GuyTorqueArmNumBolts)
                            newERIList.Add("GuyTorqueArmOutOfPlaneRestraint=" & trueFalseYesNo(.GuyTorqueArmOutOfPlaneRestraint))
                            newERIList.Add("GuyTorqueArmBoltGrade=" & .GuyTorqueArmBoltGrade)
                            newERIList.Add("GuyTorqueArmBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(.GuyTorqueArmBoltSize))
                            newERIList.Add("GuyTorqueArmBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.GuyTorqueArmBoltEdgeDistance))
                            newERIList.Add("GuyTorqueArmBoltGageDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(.GuyTorqueArmBoltGageDistance))
                            newERIList.Add("GuyPerCentTension=" & .GuyPerCentTension)
                            newERIList.Add("GuyPerCentTension120=" & .GuyPerCentTension120)
                            newERIList.Add("GuyPerCentTension240=" & .GuyPerCentTension240)
                            newERIList.Add("GuyPerCentTension360=" & .GuyPerCentTension360)
                            newERIList.Add("GuyEffFactor=" & .GuyEffFactor)
                            newERIList.Add("GuyEffFactor120=" & .GuyEffFactor120)
                            newERIList.Add("GuyEffFactor240=" & .GuyEffFactor240)
                            newERIList.Add("GuyEffFactor360=" & .GuyEffFactor360)
                            newERIList.Add("GuyNumInsulators=" & .GuyNumInsulators)
                            newERIList.Add("GuyInsulatorLength=" & Me.settings.USUnits.Properties.convertToERIUnits(.GuyInsulatorLength))
                            newERIList.Add("GuyInsulatorDia=" & Me.settings.USUnits.Properties.convertToERIUnits(.GuyInsulatorDia))
                            newERIList.Add("GuyInsulatorWt=" & Me.settings.USUnits.Force.convertToERIUnits(.GuyInsulatorWt))
                        End With
                    Next i
#End Region

#Region "Loading"
                Case line(0).Equals("NumFeedLineRecs")
                    newERIList.Add(line(0) & "=" & Me.feedLines.Count)
                    For i = 0 To Me.feedLines.Count - 1
                        With Me.feedLines(i)
                            newERIList.Add("FeedLineRec=" & .FeedLineRec)
                            newERIList.Add("FeedLineEnabled=" & trueFalseYesNo(.FeedLineEnabled))
                            newERIList.Add("FeedLineDatabase=" & .FeedLineDatabase)
                            newERIList.Add("FeedLineDescription=" & .FeedLineDescription)
                            newERIList.Add("FeedLineClassificationCategory=" & .FeedLineClassificationCategory)
                            newERIList.Add("FeedLineNote=" & .FeedLineNote)
                            newERIList.Add("FeedLineNum=" & .FeedLineNum)
                            newERIList.Add("FeedLineUseShielding=" & trueFalseYesNo(.FeedLineUseShielding))
                            newERIList.Add("ExcludeFeedLineFromTorque=" & trueFalseYesNo(.ExcludeFeedLineFromTorque))
                            newERIList.Add("FeedLineNumPerRow=" & .FeedLineNumPerRow)
                            newERIList.Add("FeedLineFace=" & .FeedLineFace)
                            newERIList.Add("FeedLineComponentType=" & .FeedLineComponentType)
                            newERIList.Add("FeedLineGroupTreatmentType=" & .FeedLineGroupTreatmentType)
                            newERIList.Add("FeedLineRoundClusterDia=" & Me.settings.USUnits.Properties.convertToERIUnits(.FeedLineRoundClusterDia))
                            newERIList.Add("FeedLineWidth=" & Me.settings.USUnits.Properties.convertToERIUnits(.FeedLineWidth))
                            newERIList.Add("FeedLinePerimeter=" & Me.settings.USUnits.Properties.convertToERIUnits(.FeedLinePerimeter))
                            newERIList.Add("FlatAttachmentEffectiveWidthRatio=" & .FlatAttachmentEffectiveWidthRatio)
                            newERIList.Add("AutoCalcFlatAttachmentEffectiveWidthRatio=" & trueFalseYesNo(.AutoCalcFlatAttachmentEffectiveWidthRatio))
                            newERIList.Add("FeedLineShieldingFactorKaNoIce=" & .FeedLineShieldingFactorKaNoIce)
                            newERIList.Add("FeedLineShieldingFactorKaIce=" & .FeedLineShieldingFactorKaIce)
                            newERIList.Add("FeedLineAutoCalcKa=" & trueFalseYesNo(.FeedLineAutoCalcKa))
                            newERIList.Add("FeedLineCaAaNoIce=" & Me.settings.USUnits.Length.convertAreaToERIUnits(.FeedLineCaAaNoIce))
                            newERIList.Add("FeedLineCaAaIce=" & Me.settings.USUnits.Length.convertAreaToERIUnits(.FeedLineCaAaIce))
                            newERIList.Add("FeedLineCaAaIce_1=" & Me.settings.USUnits.Length.convertAreaToERIUnits(.FeedLineCaAaIce_1))
                            newERIList.Add("FeedLineCaAaIce_2=" & Me.settings.USUnits.Length.convertAreaToERIUnits(.FeedLineCaAaIce_2))
                            newERIList.Add("FeedLineCaAaIce_4=" & Me.settings.USUnits.Length.convertAreaToERIUnits(.FeedLineCaAaIce_4))
                            newERIList.Add("FeedLineWtNoIce=" & Me.settings.USUnits.Load.convertToERIUnits(.FeedLineWtNoIce))
                            newERIList.Add("FeedLineWtIce=" & Me.settings.USUnits.Load.convertToERIUnits(.FeedLineWtIce))
                            newERIList.Add("FeedLineWtIce_1=" & Me.settings.USUnits.Load.convertToERIUnits(.FeedLineWtIce_1))
                            newERIList.Add("FeedLineWtIce_2=" & Me.settings.USUnits.Load.convertToERIUnits(.FeedLineWtIce_2))
                            newERIList.Add("FeedLineWtIce_4=" & Me.settings.USUnits.Load.convertToERIUnits(.FeedLineWtIce_4))
                            newERIList.Add("FeedLineFaceOffset=" & .FeedLineFaceOffset)
                            newERIList.Add("FeedLineOffsetFrac=" & .FeedLineOffsetFrac)
                            newERIList.Add("FeedLinePerimeterOffsetStartFrac=" & .FeedLinePerimeterOffsetStartFrac)
                            newERIList.Add("FeedLinePerimeterOffsetEndFrac=" & .FeedLinePerimeterOffsetEndFrac)
                            newERIList.Add("FeedLineStartHt=" & Me.settings.USUnits.Coordinate.convertToERIUnits(.FeedLineStartHt))
                            newERIList.Add("FeedLineEndHt=" & Me.settings.USUnits.Coordinate.convertToERIUnits(.FeedLineEndHt))
                            newERIList.Add("FeedLineClearSpacing=" & Me.settings.USUnits.Properties.convertToERIUnits(.FeedLineClearSpacing))
                            newERIList.Add("FeedLineRowClearSpacing=" & Me.settings.USUnits.Properties.convertToERIUnits(.FeedLineRowClearSpacing))
                        End With
                    Next i
                Case line(0).Equals("NumTowerLoadRecs")
                    newERIList.Add(line(0) & "=" & Me.discreteLoads.Count)
                    For i = 0 To Me.discreteLoads.Count - 1
                        With Me.discreteLoads(i)
                            newERIList.Add("TowerLoadRec=" & .TowerLoadRec)
                            newERIList.Add("TowerLoadEnabled=" & trueFalseYesNo(.TowerLoadEnabled))
                            newERIList.Add("TowerLoadDatabase=" & .TowerLoadDatabase)
                            newERIList.Add("TowerLoadDescription=" & .TowerLoadDescription)
                            newERIList.Add("TowerLoadType=" & .TowerLoadType)
                            newERIList.Add("TowerLoadClassificationCategory=" & .TowerLoadClassificationCategory)
                            newERIList.Add("TowerLoadNote=" & .TowerLoadNote)
                            newERIList.Add("TowerLoadNum=" & .TowerLoadNum)
                            newERIList.Add("TowerLoadFace=" & .TowerLoadFace)
                            newERIList.Add("TowerOffsetType=" & .TowerOffsetType)
                            newERIList.Add("TowerOffsetDist=" & Me.settings.USUnits.Coordinate.convertToERIUnits(.TowerOffsetDist))
                            newERIList.Add("TowerVertOffset=" & Me.settings.USUnits.Coordinate.convertToERIUnits(.TowerVertOffset))
                            newERIList.Add("TowerLateralOffset=" & Me.settings.USUnits.Coordinate.convertToERIUnits(.TowerLateralOffset))
                            newERIList.Add("TowerAzimuthAdjustment=" & Me.settings.USUnits.Rotation.convertToERIUnits(.TowerAzimuthAdjustment))
                            newERIList.Add("TowerAppurtSymbol=" & .TowerAppurtSymbol)
                            newERIList.Add("TowerLoadShieldingFactorKaNoIce=" & .TowerLoadShieldingFactorKaNoIce)
                            newERIList.Add("TowerLoadShieldingFactorKaIce=" & .TowerLoadShieldingFactorKaIce)
                            newERIList.Add("TowerLoadAutoCalcKa=" & trueFalseYesNo(.TowerLoadAutoCalcKa))
                            newERIList.Add("TowerLoadCaAaNoIce=" & Me.settings.USUnits.Length.convertAreaToERIUnits(.TowerLoadCaAaNoIce))
                            newERIList.Add("TowerLoadCaAaIce=" & Me.settings.USUnits.Length.convertAreaToERIUnits(.TowerLoadCaAaIce))
                            newERIList.Add("TowerLoadCaAaIce_1=" & Me.settings.USUnits.Length.convertAreaToERIUnits(.TowerLoadCaAaIce_1))
                            newERIList.Add("TowerLoadCaAaIce_2=" & Me.settings.USUnits.Length.convertAreaToERIUnits(.TowerLoadCaAaIce_2))
                            newERIList.Add("TowerLoadCaAaIce_4=" & Me.settings.USUnits.Length.convertAreaToERIUnits(.TowerLoadCaAaIce_4))
                            newERIList.Add("TowerLoadCaAaNoIce_Side=" & Me.settings.USUnits.Length.convertAreaToERIUnits(.TowerLoadCaAaNoIce_Side))
                            newERIList.Add("TowerLoadCaAaIce_Side=" & Me.settings.USUnits.Length.convertAreaToERIUnits(.TowerLoadCaAaIce_Side))
                            newERIList.Add("TowerLoadCaAaIce_Side_1=" & Me.settings.USUnits.Length.convertAreaToERIUnits(.TowerLoadCaAaIce_Side_1))
                            newERIList.Add("TowerLoadCaAaIce_Side_2=" & Me.settings.USUnits.Length.convertAreaToERIUnits(.TowerLoadCaAaIce_Side_2))
                            newERIList.Add("TowerLoadCaAaIce_Side_4=" & Me.settings.USUnits.Length.convertAreaToERIUnits(.TowerLoadCaAaIce_Side_4))
                            newERIList.Add("TowerLoadWtNoIce=" & Me.settings.USUnits.Force.convertToERIUnits(.TowerLoadWtNoIce))
                            newERIList.Add("TowerLoadWtIce=" & Me.settings.USUnits.Force.convertToERIUnits(.TowerLoadWtIce))
                            newERIList.Add("TowerLoadWtIce_1=" & Me.settings.USUnits.Force.convertToERIUnits(.TowerLoadWtIce_1))
                            newERIList.Add("TowerLoadWtIce_2=" & Me.settings.USUnits.Force.convertToERIUnits(.TowerLoadWtIce_2))
                            newERIList.Add("TowerLoadWtIce_4=" & Me.settings.USUnits.Force.convertToERIUnits(.TowerLoadWtIce_4))
                            newERIList.Add("TowerLoadStartHt=" & Me.settings.USUnits.Coordinate.convertToERIUnits(.TowerLoadStartHt))
                            newERIList.Add("TowerLoadEndHt=" & Me.settings.USUnits.Coordinate.convertToERIUnits(.TowerLoadEndHt))
                        End With
                    Next i
                Case line(0).Equals("NumDishRecs")
                    newERIList.Add(line(0) & "=" & Me.dishes.Count)
                    For i = 0 To Me.dishes.Count - 1
                        With Me.dishes(i)
                            newERIList.Add("DishRec=" & .DishRec)
                            newERIList.Add("DishEnabled=" & trueFalseYesNo(.DishEnabled))
                            newERIList.Add("DishDatabase=" & .DishDatabase)
                            newERIList.Add("DishDescription=" & .DishDescription)
                            newERIList.Add("DishClassificationCategory=" & .DishClassificationCategory)
                            newERIList.Add("DishNote=" & .DishNote)
                            newERIList.Add("DishNum=" & .DishNum)
                            newERIList.Add("DishFace=" & .DishFace)
                            newERIList.Add("DishType=" & .DishType)
                            newERIList.Add("DishOffsetType=" & .DishOffsetType)
                            newERIList.Add("DishVertOffset=" & Me.settings.USUnits.Coordinate.convertToERIUnits(.DishVertOffset))
                            newERIList.Add("DishLateralOffset=" & Me.settings.USUnits.Coordinate.convertToERIUnits(.DishLateralOffset))
                            newERIList.Add("DishOffsetDist=" & Me.settings.USUnits.Properties.convertToERIUnits(.DishOffsetDist))
                            newERIList.Add("DishArea=" & Me.settings.USUnits.Length.convertAreaToERIUnits(.DishArea))
                            newERIList.Add("DishAreaIce=" & Me.settings.USUnits.Length.convertAreaToERIUnits(.DishAreaIce))
                            newERIList.Add("DishAreaIce_1=" & Me.settings.USUnits.Length.convertAreaToERIUnits(.DishAreaIce_1))
                            newERIList.Add("DishAreaIce_2=" & Me.settings.USUnits.Length.convertAreaToERIUnits(.DishAreaIce_2))
                            newERIList.Add("DishAreaIce_4=" & Me.settings.USUnits.Length.convertAreaToERIUnits(.DishAreaIce_4))
                            newERIList.Add("DishDiameter=" & Me.settings.USUnits.Length.convertToERIUnits(.DishDiameter))
                            newERIList.Add("DishWtNoIce=" & Me.settings.USUnits.Force.convertToERIUnits(.DishWtNoIce))
                            newERIList.Add("DishWtIce=" & Me.settings.USUnits.Force.convertToERIUnits(.DishWtIce))
                            newERIList.Add("DishWtIce_1=" & Me.settings.USUnits.Force.convertToERIUnits(.DishWtIce_1))
                            newERIList.Add("DishWtIce_2=" & Me.settings.USUnits.Force.convertToERIUnits(.DishWtIce_2))
                            newERIList.Add("DishWtIce_4=" & Me.settings.USUnits.Force.convertToERIUnits(.DishWtIce_4))
                            newERIList.Add("DishStartHt=" & Me.settings.USUnits.Coordinate.convertToERIUnits(.DishStartHt))
                            newERIList.Add("DishAzimuthAdjustment=" & Me.settings.USUnits.Rotation.convertToERIUnits(.DishAzimuthAdjustment))
                            newERIList.Add("DishBeamWidth=" & Me.settings.USUnits.Rotation.convertToERIUnits(.DishBeamWidth))
                        End With
                    Next i
                Case line(0).Equals("NumUserForceRecs")
                    newERIList.Add(line(0) & "=" & Me.userForces.Count)
                    For i = 0 To Me.userForces.Count - 1
                        With Me.userForces(i)
                            newERIList.Add("UserForceRec=" & .UserForceRec)
                            newERIList.Add("UserForceEnabled=" & trueFalseYesNo(.UserForceEnabled))
                            newERIList.Add("UserForceDescription=" & .UserForceDescription)
                            newERIList.Add("UserForceStartHt=" & Me.settings.USUnits.Coordinate.convertToERIUnits(.UserForceStartHt))
                            newERIList.Add("UserForceOffset=" & Me.settings.USUnits.Coordinate.convertToERIUnits(.UserForceOffset))
                            newERIList.Add("UserForceAzimuth=" & Me.settings.USUnits.Rotation.convertToERIUnits(.UserForceAzimuth))
                            newERIList.Add("UserForceFxNoIce=" & Me.settings.USUnits.Force.convertToERIUnits(.UserForceFxNoIce))
                            newERIList.Add("UserForceFzNoIce=" & Me.settings.USUnits.Force.convertToERIUnits(.UserForceFzNoIce))
                            newERIList.Add("UserForceAxialNoIce=" & Me.settings.USUnits.Force.convertToERIUnits(.UserForceAxialNoIce))
                            newERIList.Add("UserForceShearNoIce=" & Me.settings.USUnits.Force.convertToERIUnits(.UserForceShearNoIce))
                            newERIList.Add("UserForceCaAcNoIce=" & Me.settings.USUnits.Length.convertToERIUnits(.UserForceCaAcNoIce))
                            newERIList.Add("UserForceFxIce=" & Me.settings.USUnits.Force.convertToERIUnits(.UserForceFxIce))
                            newERIList.Add("UserForceFzIce=" & Me.settings.USUnits.Force.convertToERIUnits(.UserForceFzIce))
                            newERIList.Add("UserForceAxialIce=" & Me.settings.USUnits.Force.convertToERIUnits(.UserForceAxialIce))
                            newERIList.Add("UserForceShearIce=" & Me.settings.USUnits.Force.convertToERIUnits(.UserForceShearIce))
                            newERIList.Add("UserForceCaAcIce=" & Me.settings.USUnits.Length.convertToERIUnits(.UserForceCaAcIce))
                            newERIList.Add("UserForceFxService=" & Me.settings.USUnits.Force.convertToERIUnits(.UserForceFxService))
                            newERIList.Add("UserForceFzService=" & Me.settings.USUnits.Force.convertToERIUnits(.UserForceFzService))
                            newERIList.Add("UserForceAxialService=" & Me.settings.USUnits.Force.convertToERIUnits(.UserForceAxialService))
                            newERIList.Add("UserForceShearService=" & Me.settings.USUnits.Force.convertToERIUnits(.UserForceShearService))
                            newERIList.Add("UserForceCaAcService=" & Me.settings.USUnits.Length.convertToERIUnits(.UserForceCaAcService))
                            newERIList.Add("UserForceEhx=" & Me.settings.USUnits.Force.convertToERIUnits(.UserForceEhx))
                            newERIList.Add("UserForceEhz=" & Me.settings.USUnits.Force.convertToERIUnits(.UserForceEhz))
                            newERIList.Add("UserForceEv=" & Me.settings.USUnits.Force.convertToERIUnits(.UserForceEv))
                            newERIList.Add("UserForceEh=" & Me.settings.USUnits.Force.convertToERIUnits(.UserForceEh))
                        End With
                    Next i
#End Region
#Region "Chronos and Data after Structure"
                Case line(0).Equals("[Nodes]")
                    newERIList.Add(line(0))
                    newERIList.Add("#node_number activity x y z")

                Case line(0).Equals("[Supports]")
                    newERIList.Add(line(0))
                    newERIList.Add("#node_number code_dx dy dz ox oy oz stiffness_dx dy dz ox oy oz alpha beta gamma ref_temp")

                Case line(0).Equals("[ConnectionArray]")
                    newERIList.Add(line(0))
                    newERIList.Add("#node_number conn[24]")

                Case line(0).Equals("[Elements]")
                    newERIList.Add(line(0))
                    newERIList.Add("#elem_number activity group mat prop inspect")
                    newERIList.Add("#for groups 1,2,3,4,5,6,7,8,9")
                    newERIList.Add("#   node1 node2")
                    newERIList.Add("#for group 1a (composite beam)")
                    newERIList.Add("#   composite_prop_id node1 node2")

                Case line(0).Equals("[Conditions]")
                    newERIList.Add(line(0))
                    newERIList.Add("#elem_number group temp")
                    newERIList.Add("#if groups 1-9 offsets: dx1 dy1 dz1 dx2 dy2 dz2")
                    newERIList.Add("#if group = 1, 2 (beam or npbeam)")
                    newERIList.Add("#  gamma release1 release2")
                    newERIList.Add("#  node1_stiffness_dx dy dz ox oy oz")
                    newERIList.Add("#  node2_stiffness_dx dy dz ox oy oz")
                    newERIList.Add("#if group = 3, 4, 8 (truss, strut, fuse)")
                    newERIList.Add("#  gamma")
                    newERIList.Add("#if group = 5, 6, 7 (hook, gap, slot)")
                    newERIList.Add("#  gamma gap")
                    newERIList.Add("#if group = 9 (cable)")
                    newERIList.Add("#  initial_tension unstressed_length")

                Case line(0).Equals("[Segments]")
                    newERIList.Add(line(0))
                    newERIList.Add("#master_element number_of_internal_nodes")
                    newERIList.Add("#internal_node1 internal_node2 falseTE: node is - if end of span....")

                Case line(0).Equals("[Material]")
                    newERIList.Add(line(0))
                    newERIList.Add("#mat activity")
                    newERIList.Add("#description")
                    newERIList.Add("#emod poiss thermal mden wden")

                Case line(0).Equals("[Properties]")
                    newERIList.Add(line(0))
                    newERIList.Add("#prop activity group")
                    newERIList.Add("#desigination")
                    newERIList.Add("#section_file")
                    newERIList.Add("#if group = 1 (beam)")
                    newERIList.Add("#  area ix iy iz ixy sfy sfx cw")
                    newERIList.Add("#else if group = 2 (npbeam)")
                    newERIList.Add("#  area ix iy iz r1 r2 r3 r4 n1 n2 n3 n4")
                    newERIList.Add("#else if group = 3,4,5,6,7 (truss,strut,hook,gap,or slot)")
                    newERIList.Add("#  area")
                    newERIList.Add("#else if group = 8 (fuse)")
                    newERIList.Add("#  area comp_capacity ecu/ecy ten_capacity etu/ety")
                    newERIList.Add("#else if group = 9 (cable)")
                    newERIList.Add("#  area diameter")

                Case line(0).Equals("[SelfWeight]")
                    newERIList.Add(line(0))
                    newERIList.Add("#rec activity")
                    newERIList.Add("#description")
                    newERIList.Add("#cases")
                    newERIList.Add("#elem_list")
                    newERIList.Add("#x_mult y_mult z_mult")

                Case line(0).Equals("[Thermal]")
                    newERIList.Add(line(0))
                    newERIList.Add("#rec activity")
                    newERIList.Add("#description")
                    newERIList.Add("#cases")
                    newERIList.Add("#elem_list")
                    newERIList.Add("#multiplier")

                Case line(0).Equals("[Combinations]")
                    newERIList.Add(line(0))
                    newERIList.Add("#comb_number activity")
                    newERIList.Add("#description")
                    newERIList.Add("#list_of_factors")
                    newERIList.AddRange(defaults.LoadCombos)

                Case line(0).Equals("[Envelopes]")
                    newERIList.Add(line(0))
                    newERIList.Add("#envelope_number activity")
                    newERIList.Add("#description")
                    newERIList.Add("#list_of_combinations")
                    newERIList.AddRange(defaults.LoadEnvs)

                Case line(0).Equals("[Cases]")
                    newERIList.Add(line(0))
                    newERIList.Add("#rec activity internal_ref_case")
                    newERIList.Add("#description")
                    newERIList.AddRange(defaults.LoadCases)

                Case line(0).Equals("[OutputOptions]")
                    newERIList.Add(line(0))
                    newERIList.AddRange(defaults.OutputOptions)

                Case line(0).Equals("[ReportOptions]")
                    newERIList.Add(line(0))
                    newERIList.Add("#rec activity group type")
                    newERIList.Add("#elemlist")
                    newERIList.Add("#description")
                    newERIList.Add("#if type = 4,5 (force range SumOfAppliedForcesmary or worst)")
                    newERIList.Add("#  amin amax smin smax mmin mmax tmin tmax")
                    newERIList.Add("#else if type = 6 (axial force past limit)")
                    newERIList.Add("#  compression tension")

                Case line(0).Equals("[Material]")
                    newERIList.Add(line(0))
                    newERIList.Add("#mat activity")
                    newERIList.Add("#description")
                    newERIList.Add("#emod poiss thermal mden wden")
#End Region

#Region "[Databases]"
                Case line(0).Equals("[Databases]")
                    newERIList.Add(line(0))
                    For Each memb In database.members
                        newERIList.Add("File=" & memb.File)
                        newERIList.Add("USName=" & memb.USName)
                        newERIList.Add("SIName=" & memb.SIName)
                        newERIList.Add("Values=" & memb.Values)
                    Next
                    For Each mat In database.materials 'Member materials
                        newERIList.Add(IIf(mat.IsBolt, "BoltMatFile=", "MembermatFile=") & mat.MemberMatFile)
                        newERIList.Add("MatName=" & mat.MatName)
                        newERIList.Add("MatValues=" & mat.MatValues)
                    Next
                    For Each mat In database.bolts 'Bolt materials
                        newERIList.Add(IIf(mat.IsBolt, "BoltMatFile=", "MembermatFile=") & mat.MemberMatFile)
                        newERIList.Add("MatName=" & mat.MatName)
                        newERIList.Add("MatValues=" & mat.MatValues)
                    Next

#End Region

#Region "One liners"
                Case Else '[EndCHRONOS], [EndOverwrite], [CHRONOS], [End Structure], [End Application]
                    If line.Count = 1 Then
                        newERIList.Add(line(0))
                    Else
                        If line(0).Contains("[") Then
                            newERIList.Add(line(0))
                        Else
                            newERIList.Add(line(0) & "=" & line(1))
                        End If
                    End If
#End Region

            End Select
        Next

        File.WriteAllLines(FilePath, newERIList, Text.Encoding.ASCII)

    End Sub
#End Region

#Region "Results"
    Public Sub GetResults()

        Dim tnxResultXMLPath = Me.filePath & ".XMLOUT.xml"

        If Not FileIO.FileSystem.FileExists(tnxResultXMLPath) Then
            Debug.WriteLine("No TNX Results XML found.")
            Exit Sub
        End If

        Using resultsReader As XmlReader = XmlReader.Create(tnxResultXMLPath)
            Dim tnxResultSerializer As New XmlSerializer(GetType(tnxTowerOutput))
            Dim tnxXMLResults As tnxTowerOutput = tnxResultSerializer.Deserialize(resultsReader)
            tnxXMLResults.ConverttoEDSResults(geometry)
        End Using


    End Sub

#End Region

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As tnxModel = TryCast(other, tnxModel)
        If otherToCompare Is Nothing Then Return False
        Equals = If(Me.settings.CheckChange(otherToCompare.settings, changes, categoryName, "Settings"), Equals, False)
        Equals = If(Me.solutionSettings.CheckChange(otherToCompare.solutionSettings, changes, categoryName, "Solution Settings"), Equals, False)
        Equals = If(Me.MTOSettings.CheckChange(otherToCompare.MTOSettings, changes, categoryName, "MTO Settings"), Equals, False)
        Equals = If(Me.reportSettings.CheckChange(otherToCompare.reportSettings, changes, categoryName, "Report Settings"), Equals, False)
        Equals = If(Me.CCIReport.CheckChange(otherToCompare.CCIReport, changes, categoryName, "CCI Report"), Equals, False)
        Equals = If(Me.code.CheckChange(otherToCompare.code, changes, categoryName, "Code"), Equals, False)
        Equals = If(Me.options.CheckChange(otherToCompare.options, changes, categoryName, "Options"), Equals, False)
        'The ConsiderGeometryEquality is passed onto the geometry object because there are a number of geometry options that get saved in the main TNX table and we still need to consider.
        Equals = If(Me.geometry.CheckChange(otherToCompare.geometry, changes, categoryName, "Geometry"), Equals, False)
        If Me.ConsiderGeometryEquality Then
            Equals = If(Me.database.CheckChange(otherToCompare.database, changes, categoryName, "Database"), Equals, False)
        End If
        If Me.ConsiderLoadingEquality Then
            Equals = If(Me.feedLines.CheckChange(otherToCompare.feedLines, changes, categoryName, "Feed Lines"), Equals, False)
            Equals = If(Me.discreteLoads.CheckChange(otherToCompare.discreteLoads, changes, categoryName, "Discrete"), Equals, False)
            Equals = If(Me.dishes.CheckChange(otherToCompare.dishes, changes, categoryName, "Dishes"), Equals, False)
            Equals = If(Me.userForces.CheckChange(otherToCompare.userForces, changes, categoryName, "User Forces"), Equals, False)
        End If
        Return Equals
    End Function

    Public Overloads Function Equals(other As EDSObject, Optional ByRef changes As List(Of AnalysisChange) = Nothing, Optional IgnoreGeometry As Boolean = False, Optional IgnoreLoading As Boolean = False) As Boolean
        'Use this to temporarily set the ConsiderGeometryEquality and ConsiderLoadingEquality for one equality check

        Dim currentGeometry As Boolean = Me.ConsiderGeometryEquality
        Dim currentLoading As Boolean = Me.ConsiderLoadingEquality

        Me.ConsiderGeometryEquality = Not IgnoreGeometry
        Me.ConsiderLoadingEquality = Not IgnoreLoading

        If other Is Nothing Then
            Return False
        Else
            'Call Equals(other As EDSObject, ByRef changes As List(Of AnalysisChanges))
            Return Me.Equals(other, changes)
        End If

        Me.ConsiderGeometryEquality = currentGeometry
        Me.ConsiderLoadingEquality = currentLoading

    End Function

    Public Function TNXEquals(other As tnxModel) As Boolean
        'This function is required because the TNXModel has many subclasses who's properties are all stored in one table.
        'This will determine if any of the properties in that table need to be updated but won't consider changes in the subclasses that have their own tables.
        TNXEquals = True

        If other Is Nothing Then Return False

        TNXEquals = If(Me.settings.projectInfo.DesignStandardSeries = other.settings.projectInfo.DesignStandardSeries, TNXEquals, False)
        TNXEquals = If(Me.settings.projectInfo.UnitsSystem = other.settings.projectInfo.UnitsSystem, TNXEquals, False)
        TNXEquals = If(Me.settings.projectInfo.ClientName = other.settings.projectInfo.ClientName, TNXEquals, False)
        TNXEquals = If(Me.settings.projectInfo.ProjectName = other.settings.projectInfo.ProjectName, TNXEquals, False)
        TNXEquals = If(Me.settings.projectInfo.ProjectNumber = other.settings.projectInfo.ProjectNumber, TNXEquals, False)
        TNXEquals = If(Me.settings.projectInfo.CreatedBy = other.settings.projectInfo.CreatedBy, TNXEquals, False)
        TNXEquals = If(Me.settings.projectInfo.CreatedOn = other.settings.projectInfo.CreatedOn, TNXEquals, False)
        TNXEquals = If(Me.settings.projectInfo.LastUsedBy = other.settings.projectInfo.LastUsedBy, TNXEquals, False)
        TNXEquals = If(Me.settings.projectInfo.LastUsedOn = other.settings.projectInfo.LastUsedOn, TNXEquals, False)
        TNXEquals = If(Me.settings.projectInfo.VersionUsed = other.settings.projectInfo.VersionUsed, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Length.value = other.settings.USUnits.Length.value, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Length.precision = other.settings.USUnits.Length.precision, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Coordinate.value = other.settings.USUnits.Coordinate.value, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Coordinate.precision = other.settings.USUnits.Coordinate.precision, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Force.value = other.settings.USUnits.Force.value, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Force.precision = other.settings.USUnits.Force.precision, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Load.value = other.settings.USUnits.Load.value, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Load.precision = other.settings.USUnits.Load.precision, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Moment.value = other.settings.USUnits.Moment.value, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Moment.precision = other.settings.USUnits.Moment.precision, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Properties.value = other.settings.USUnits.Properties.value, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Properties.precision = other.settings.USUnits.Properties.precision, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Pressure.value = other.settings.USUnits.Pressure.value, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Pressure.precision = other.settings.USUnits.Pressure.precision, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Velocity.value = other.settings.USUnits.Velocity.value, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Velocity.precision = other.settings.USUnits.Velocity.precision, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Displacement.value = other.settings.USUnits.Displacement.value, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Displacement.precision = other.settings.USUnits.Displacement.precision, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Mass.value = other.settings.USUnits.Mass.value, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Mass.precision = other.settings.USUnits.Mass.precision, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Acceleration.value = other.settings.USUnits.Acceleration.value, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Acceleration.precision = other.settings.USUnits.Acceleration.precision, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Stress.value = other.settings.USUnits.Stress.value, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Stress.precision = other.settings.USUnits.Stress.precision, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Density.value = other.settings.USUnits.Density.value, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Density.precision = other.settings.USUnits.Density.precision, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.UnitWt.value = other.settings.USUnits.UnitWt.value, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.UnitWt.precision = other.settings.USUnits.UnitWt.precision, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Strength.value = other.settings.USUnits.Strength.value, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Strength.precision = other.settings.USUnits.Strength.precision, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Modulus.value = other.settings.USUnits.Modulus.value, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Modulus.precision = other.settings.USUnits.Modulus.precision, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Temperature.value = other.settings.USUnits.Temperature.value, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Temperature.precision = other.settings.USUnits.Temperature.precision, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Printer.value = other.settings.USUnits.Printer.value, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Printer.precision = other.settings.USUnits.Printer.precision, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Rotation.value = other.settings.USUnits.Rotation.value, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Rotation.precision = other.settings.USUnits.Rotation.precision, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Spacing.value = other.settings.USUnits.Spacing.value, TNXEquals, False)
        TNXEquals = If(Me.settings.USUnits.Spacing.precision = other.settings.USUnits.Spacing.precision, TNXEquals, False)
        TNXEquals = If(Me.settings.userInfo.ViewerUserName = other.settings.userInfo.ViewerUserName, TNXEquals, False)
        TNXEquals = If(Me.settings.userInfo.ViewerCompanyName = other.settings.userInfo.ViewerCompanyName, TNXEquals, False)
        TNXEquals = If(Me.settings.userInfo.ViewerStreetAddress = other.settings.userInfo.ViewerStreetAddress, TNXEquals, False)
        TNXEquals = If(Me.settings.userInfo.ViewerCityState = other.settings.userInfo.ViewerCityState, TNXEquals, False)
        TNXEquals = If(Me.settings.userInfo.ViewerPhone = other.settings.userInfo.ViewerPhone, TNXEquals, False)
        TNXEquals = If(Me.settings.userInfo.ViewerFAX = other.settings.userInfo.ViewerFAX, TNXEquals, False)
        TNXEquals = If(Me.settings.userInfo.ViewerLogo = other.settings.userInfo.ViewerLogo, TNXEquals, False)
        TNXEquals = If(Me.settings.userInfo.ViewerCompanyBitmap = other.settings.userInfo.ViewerCompanyBitmap, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportProjectNumber = other.CCIReport.sReportProjectNumber, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportJobType = other.CCIReport.sReportJobType, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportCarrierName = other.CCIReport.sReportCarrierName, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportCarrierSiteNumber = other.CCIReport.sReportCarrierSiteNumber, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportCarrierSiteName = other.CCIReport.sReportCarrierSiteName, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportSiteAddress = other.CCIReport.sReportSiteAddress, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportLatitudeDegree = other.CCIReport.sReportLatitudeDegree, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportLatitudeMinute = other.CCIReport.sReportLatitudeMinute, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportLatitudeSecond = other.CCIReport.sReportLatitudeSecond, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportLongitudeDegree = other.CCIReport.sReportLongitudeDegree, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportLongitudeMinute = other.CCIReport.sReportLongitudeMinute, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportLongitudeSecond = other.CCIReport.sReportLongitudeSecond, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportLocalCodeRequirement = other.CCIReport.sReportLocalCodeRequirement, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportSiteHistory = other.CCIReport.sReportSiteHistory, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportTowerManufacturer = other.CCIReport.sReportTowerManufacturer, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportMonthManufactured = other.CCIReport.sReportMonthManufactured, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportYearManufactured = other.CCIReport.sReportYearManufactured, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportOriginalSpeed = other.CCIReport.sReportOriginalSpeed, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportOriginalCode = other.CCIReport.sReportOriginalCode, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportTowerType = other.CCIReport.sReportTowerType, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportEngrName = other.CCIReport.sReportEngrName, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportEngrTitle = other.CCIReport.sReportEngrTitle, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportHQPhoneNumber = other.CCIReport.sReportHQPhoneNumber, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportEmailAddress = other.CCIReport.sReportEmailAddress, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportLogoPath = other.CCIReport.sReportLogoPath, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportCCiContactName = other.CCIReport.sReportCCiContactName, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportCCiAddress1 = other.CCIReport.sReportCCiAddress1, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportCCiAddress2 = other.CCIReport.sReportCCiAddress2, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportCCiBUNumber = other.CCIReport.sReportCCiBUNumber, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportCCiSiteName = other.CCIReport.sReportCCiSiteName, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportCCiJDENumber = other.CCIReport.sReportCCiJDENumber, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportCCiWONumber = other.CCIReport.sReportCCiWONumber, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportCCiPONumber = other.CCIReport.sReportCCiPONumber, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportCCiAppNumber = other.CCIReport.sReportCCiAppNumber, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportCCiRevNumber = other.CCIReport.sReportCCiRevNumber, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportRecommendations = other.CCIReport.sReportRecommendations, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportAppurt1Note1 = other.CCIReport.sReportAppurt1Note1, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportAppurt1Note2 = other.CCIReport.sReportAppurt1Note2, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportAppurt1Note3 = other.CCIReport.sReportAppurt1Note3, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportAppurt1Note4 = other.CCIReport.sReportAppurt1Note4, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportAppurt1Note5 = other.CCIReport.sReportAppurt1Note5, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportAppurt1Note6 = other.CCIReport.sReportAppurt1Note6, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportAppurt1Note7 = other.CCIReport.sReportAppurt1Note7, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportAppurt2Note1 = other.CCIReport.sReportAppurt2Note1, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportAppurt2Note2 = other.CCIReport.sReportAppurt2Note2, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportAppurt2Note3 = other.CCIReport.sReportAppurt2Note3, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportAppurt2Note4 = other.CCIReport.sReportAppurt2Note4, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportAppurt2Note5 = other.CCIReport.sReportAppurt2Note5, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportAppurt2Note6 = other.CCIReport.sReportAppurt2Note6, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportAppurt2Note7 = other.CCIReport.sReportAppurt2Note7, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportAddlCapacityNote1 = other.CCIReport.sReportAddlCapacityNote1, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportAddlCapacityNote2 = other.CCIReport.sReportAddlCapacityNote2, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportAddlCapacityNote3 = other.CCIReport.sReportAddlCapacityNote3, TNXEquals, False)
        TNXEquals = If(Me.CCIReport.sReportAddlCapacityNote4 = other.CCIReport.sReportAddlCapacityNote4, TNXEquals, False)
        TNXEquals = If(Me.code.design.DesignCode = other.code.design.DesignCode, TNXEquals, False)
        TNXEquals = If(Me.geometry.TowerType = other.geometry.TowerType, TNXEquals, False)
        TNXEquals = If(Me.geometry.AntennaType = other.geometry.AntennaType, TNXEquals, False)
        TNXEquals = If(Me.geometry.OverallHeight = other.geometry.OverallHeight, TNXEquals, False)
        TNXEquals = If(Me.geometry.BaseElevation = other.geometry.BaseElevation, TNXEquals, False)
        TNXEquals = If(Me.geometry.Lambda = other.geometry.Lambda, TNXEquals, False)
        TNXEquals = If(Me.geometry.TowerTopFaceWidth = other.geometry.TowerTopFaceWidth, TNXEquals, False)
        TNXEquals = If(Me.geometry.TowerBaseFaceWidth = other.geometry.TowerBaseFaceWidth, TNXEquals, False)
        TNXEquals = If(Me.code.wind.WindSpeed = other.code.wind.WindSpeed, TNXEquals, False)
        TNXEquals = If(Me.code.wind.WindSpeedIce = other.code.wind.WindSpeedIce, TNXEquals, False)
        TNXEquals = If(Me.code.wind.WindSpeedService = other.code.wind.WindSpeedService, TNXEquals, False)
        TNXEquals = If(Me.code.ice.IceThickness = other.code.ice.IceThickness, TNXEquals, False)
        TNXEquals = If(Me.code.wind.CSA_S37_RefVelPress = other.code.wind.CSA_S37_RefVelPress, TNXEquals, False)
        TNXEquals = If(Me.code.wind.CSA_S37_ReliabilityClass = other.code.wind.CSA_S37_ReliabilityClass, TNXEquals, False)
        TNXEquals = If(Me.code.wind.CSA_S37_ServiceabilityFactor = other.code.wind.CSA_S37_ServiceabilityFactor, TNXEquals, False)
        TNXEquals = If(Me.code.ice.UseModified_TIA_222_IceParameters = other.code.ice.UseModified_TIA_222_IceParameters, TNXEquals, False)
        TNXEquals = If(Me.code.ice.TIA_222_IceThicknessMultiplier = other.code.ice.TIA_222_IceThicknessMultiplier, TNXEquals, False)
        TNXEquals = If(Me.code.ice.DoNotUse_TIA_222_IceEscalation = other.code.ice.DoNotUse_TIA_222_IceEscalation, TNXEquals, False)
        TNXEquals = If(Me.code.ice.IceDensity = other.code.ice.IceDensity, TNXEquals, False)
        TNXEquals = If(Me.code.seismic.SeismicSiteClass = other.code.seismic.SeismicSiteClass, TNXEquals, False)
        TNXEquals = If(Me.code.seismic.SeismicSs = other.code.seismic.SeismicSs, TNXEquals, False)
        TNXEquals = If(Me.code.seismic.SeismicS1 = other.code.seismic.SeismicS1, TNXEquals, False)
        TNXEquals = If(Me.code.thermal.TempDrop = other.code.thermal.TempDrop, TNXEquals, False)
        TNXEquals = If(Me.code.misclCode.GroutFc = other.code.misclCode.GroutFc, TNXEquals, False)
        TNXEquals = If(Me.options.defaultGirtOffsets.GirtOffset = other.options.defaultGirtOffsets.GirtOffset, TNXEquals, False)
        TNXEquals = If(Me.options.defaultGirtOffsets.GirtOffsetLatticedPole = other.options.defaultGirtOffsets.GirtOffsetLatticedPole, TNXEquals, False)
        TNXEquals = If(Me.options.foundationStiffness.MastVert = other.options.foundationStiffness.MastVert, TNXEquals, False)
        TNXEquals = If(Me.options.foundationStiffness.MastHorz = other.options.foundationStiffness.MastHorz, TNXEquals, False)
        TNXEquals = If(Me.options.foundationStiffness.GuyVert = other.options.foundationStiffness.GuyVert, TNXEquals, False)
        TNXEquals = If(Me.options.foundationStiffness.GuyHorz = other.options.foundationStiffness.GuyHorz, TNXEquals, False)
        TNXEquals = If(Me.options.misclOptions.HogRodTakeup = other.options.misclOptions.HogRodTakeup, TNXEquals, False)
        TNXEquals = If(Me.geometry.TowerTaper = other.geometry.TowerTaper, TNXEquals, False)
        TNXEquals = If(Me.geometry.GuyedMonopoleBaseType = other.geometry.GuyedMonopoleBaseType, TNXEquals, False)
        TNXEquals = If(Me.geometry.TaperHeight = other.geometry.TaperHeight, TNXEquals, False)
        TNXEquals = If(Me.geometry.PivotHeight = other.geometry.PivotHeight, TNXEquals, False)
        TNXEquals = If(Me.geometry.AutoCalcGH = other.geometry.AutoCalcGH, TNXEquals, False)
        TNXEquals = If(Me.MTOSettings.IncludeCapacityNote = other.MTOSettings.IncludeCapacityNote, TNXEquals, False)
        TNXEquals = If(Me.MTOSettings.IncludeAppurtGraphics = other.MTOSettings.IncludeAppurtGraphics, TNXEquals, False)
        TNXEquals = If(Me.MTOSettings.DisplayNotes = other.MTOSettings.DisplayNotes, TNXEquals, False)
        TNXEquals = If(Me.MTOSettings.DisplayReactions = other.MTOSettings.DisplayReactions, TNXEquals, False)
        TNXEquals = If(Me.MTOSettings.DisplaySchedule = other.MTOSettings.DisplaySchedule, TNXEquals, False)
        TNXEquals = If(Me.MTOSettings.DisplayAppurtenanceTable = other.MTOSettings.DisplayAppurtenanceTable, TNXEquals, False)
        TNXEquals = If(Me.MTOSettings.DisplayMaterialStrengthTable = other.MTOSettings.DisplayMaterialStrengthTable, TNXEquals, False)
        TNXEquals = If(Me.code.wind.AutoCalc_ASCE_GH = other.code.wind.AutoCalc_ASCE_GH, TNXEquals, False)
        TNXEquals = If(Me.code.wind.ASCE_ExposureCat = other.code.wind.ASCE_ExposureCat, TNXEquals, False)
        TNXEquals = If(Me.code.wind.ASCE_Year = other.code.wind.ASCE_Year, TNXEquals, False)
        TNXEquals = If(Me.code.wind.ASCEGh = other.code.wind.ASCEGh, TNXEquals, False)
        TNXEquals = If(Me.code.wind.ASCEI = other.code.wind.ASCEI, TNXEquals, False)
        TNXEquals = If(Me.code.wind.UseASCEWind = other.code.wind.UseASCEWind, TNXEquals, False)
        TNXEquals = If(Me.geometry.UserGHElev = other.geometry.UserGHElev, TNXEquals, False)
        TNXEquals = If(Me.code.design.UseCodeGuySF = other.code.design.UseCodeGuySF, TNXEquals, False)
        TNXEquals = If(Me.code.design.GuySF = other.code.design.GuySF, TNXEquals, False)
        TNXEquals = If(Me.code.wind.CalcWindAt = other.code.wind.CalcWindAt, TNXEquals, False)
        TNXEquals = If(Me.code.misclCode.TowerBoltGrade = other.code.misclCode.TowerBoltGrade, TNXEquals, False)
        TNXEquals = If(Me.code.misclCode.TowerBoltMinEdgeDist = other.code.misclCode.TowerBoltMinEdgeDist, TNXEquals, False)
        TNXEquals = If(Me.code.design.AllowStressRatio = other.code.design.AllowStressRatio, TNXEquals, False)
        TNXEquals = If(Me.code.design.AllowAntStressRatio = other.code.design.AllowAntStressRatio, TNXEquals, False)
        TNXEquals = If(Me.code.wind.WindCalcPoints = other.code.wind.WindCalcPoints, TNXEquals, False)
        TNXEquals = If(Me.geometry.UseIndexPlate = other.geometry.UseIndexPlate, TNXEquals, False)
        TNXEquals = If(Me.geometry.EnterUserDefinedGhValues = other.geometry.EnterUserDefinedGhValues, TNXEquals, False)
        TNXEquals = If(Me.geometry.BaseTowerGhInput = other.geometry.BaseTowerGhInput, TNXEquals, False)
        TNXEquals = If(Me.geometry.UpperStructureGhInput = other.geometry.UpperStructureGhInput, TNXEquals, False)
        TNXEquals = If(Me.geometry.EnterUserDefinedCgValues = other.geometry.EnterUserDefinedCgValues, TNXEquals, False)
        TNXEquals = If(Me.geometry.BaseTowerCgInput = other.geometry.BaseTowerCgInput, TNXEquals, False)
        TNXEquals = If(Me.geometry.UpperStructureCgInput = other.geometry.UpperStructureCgInput, TNXEquals, False)
        TNXEquals = If(Me.options.cantileverPoles.CheckVonMises = other.options.cantileverPoles.CheckVonMises, TNXEquals, False)
        TNXEquals = If(Me.options.UseClearSpans = other.options.UseClearSpans, TNXEquals, False)
        TNXEquals = If(Me.options.UseClearSpansKlr = other.options.UseClearSpansKlr, TNXEquals, False)
        TNXEquals = If(Me.geometry.AntennaFaceWidth = other.geometry.AntennaFaceWidth, TNXEquals, False)
        TNXEquals = If(Me.code.design.DoInteraction = other.code.design.DoInteraction, TNXEquals, False)
        TNXEquals = If(Me.code.design.DoHorzInteraction = other.code.design.DoHorzInteraction, TNXEquals, False)
        TNXEquals = If(Me.code.design.DoDiagInteraction = other.code.design.DoDiagInteraction, TNXEquals, False)
        TNXEquals = If(Me.code.design.UseMomentMagnification = other.code.design.UseMomentMagnification, TNXEquals, False)
        TNXEquals = If(Me.options.UseFeedlineAsCylinder = other.options.UseFeedlineAsCylinder, TNXEquals, False)
        TNXEquals = If(Me.options.defaultGirtOffsets.OffsetBotGirt = other.options.defaultGirtOffsets.OffsetBotGirt, TNXEquals, False)
        TNXEquals = If(Me.code.design.PrintBitmaps = other.code.design.PrintBitmaps, TNXEquals, False)
        TNXEquals = If(Me.geometry.UseTopTakeup = other.geometry.UseTopTakeup, TNXEquals, False)
        TNXEquals = If(Me.geometry.ConstantSlope = other.geometry.ConstantSlope, TNXEquals, False)
        TNXEquals = If(Me.code.design.UseCodeStressRatio = other.code.design.UseCodeStressRatio, TNXEquals, False)
        TNXEquals = If(Me.options.UseLegLoads = other.options.UseLegLoads, TNXEquals, False)
        TNXEquals = If(Me.code.design.ERIDesignMode = other.code.design.ERIDesignMode, TNXEquals, False)
        TNXEquals = If(Me.code.wind.WindExposure = other.code.wind.WindExposure, TNXEquals, False)
        TNXEquals = If(Me.code.wind.WindZone = other.code.wind.WindZone, TNXEquals, False)
        TNXEquals = If(Me.code.wind.StructureCategory = other.code.wind.StructureCategory, TNXEquals, False)
        TNXEquals = If(Me.code.wind.RiskCategory = other.code.wind.RiskCategory, TNXEquals, False)
        TNXEquals = If(Me.code.wind.TopoCategory = other.code.wind.TopoCategory, TNXEquals, False)
        TNXEquals = If(Me.code.wind.RSMTopographicFeature = other.code.wind.RSMTopographicFeature, TNXEquals, False)
        TNXEquals = If(Me.code.wind.RSM_L = other.code.wind.RSM_L, TNXEquals, False)
        TNXEquals = If(Me.code.wind.RSM_X = other.code.wind.RSM_X, TNXEquals, False)
        TNXEquals = If(Me.code.wind.CrestHeight = other.code.wind.CrestHeight, TNXEquals, False)
        TNXEquals = If(Me.code.wind.TIA_222_H_TopoFeatureDownwind = other.code.wind.TIA_222_H_TopoFeatureDownwind, TNXEquals, False)
        TNXEquals = If(Me.code.wind.BaseElevAboveSeaLevel = other.code.wind.BaseElevAboveSeaLevel, TNXEquals, False)
        TNXEquals = If(Me.code.wind.ConsiderRooftopSpeedUp = other.code.wind.ConsiderRooftopSpeedUp, TNXEquals, False)
        TNXEquals = If(Me.code.wind.RooftopWS = other.code.wind.RooftopWS, TNXEquals, False)
        TNXEquals = If(Me.code.wind.RooftopHS = other.code.wind.RooftopHS, TNXEquals, False)
        TNXEquals = If(Me.code.wind.RooftopParapetHt = other.code.wind.RooftopParapetHt, TNXEquals, False)
        TNXEquals = If(Me.code.wind.RooftopXB = other.code.wind.RooftopXB, TNXEquals, False)
        TNXEquals = If(Me.code.design.UseTIA222H_AnnexS = other.code.design.UseTIA222H_AnnexS, TNXEquals, False)
        TNXEquals = If(Me.code.design.TIA_222_H_AnnexS_Ratio = other.code.design.TIA_222_H_AnnexS_Ratio, TNXEquals, False)
        TNXEquals = If(Me.code.wind.EIACWindMult = other.code.wind.EIACWindMult, TNXEquals, False)
        TNXEquals = If(Me.code.wind.EIACWindMultIce = other.code.wind.EIACWindMultIce, TNXEquals, False)
        TNXEquals = If(Me.code.wind.EIACIgnoreCableDrag = other.code.wind.EIACIgnoreCableDrag, TNXEquals, False)
        TNXEquals = If(Me.MTOSettings.Notes = other.MTOSettings.Notes, TNXEquals, False)
        TNXEquals = If(Me.reportSettings.ReportInputCosts = other.reportSettings.ReportInputCosts, TNXEquals, False)
        TNXEquals = If(Me.reportSettings.ReportInputGeometry = other.reportSettings.ReportInputGeometry, TNXEquals, False)
        TNXEquals = If(Me.reportSettings.ReportInputOptions = other.reportSettings.ReportInputOptions, TNXEquals, False)
        TNXEquals = If(Me.reportSettings.ReportMaxForces = other.reportSettings.ReportMaxForces, TNXEquals, False)
        TNXEquals = If(Me.reportSettings.ReportInputMap = other.reportSettings.ReportInputMap, TNXEquals, False)
        TNXEquals = If(Me.reportSettings.CostReportOutputType = other.reportSettings.CostReportOutputType, TNXEquals, False)
        TNXEquals = If(Me.reportSettings.CapacityReportOutputType = other.reportSettings.CapacityReportOutputType, TNXEquals, False)
        TNXEquals = If(Me.reportSettings.ReportPrintForceTotals = other.reportSettings.ReportPrintForceTotals, TNXEquals, False)
        TNXEquals = If(Me.reportSettings.ReportPrintForceDetails = other.reportSettings.ReportPrintForceDetails, TNXEquals, False)
        TNXEquals = If(Me.reportSettings.ReportPrintMastVectors = other.reportSettings.ReportPrintMastVectors, TNXEquals, False)
        TNXEquals = If(Me.reportSettings.ReportPrintAntPoleVectors = other.reportSettings.ReportPrintAntPoleVectors, TNXEquals, False)
        TNXEquals = If(Me.reportSettings.ReportPrintDiscreteVectors = other.reportSettings.ReportPrintDiscreteVectors, TNXEquals, False)
        TNXEquals = If(Me.reportSettings.ReportPrintDishVectors = other.reportSettings.ReportPrintDishVectors, TNXEquals, False)
        TNXEquals = If(Me.reportSettings.ReportPrintFeedTowerVectors = other.reportSettings.ReportPrintFeedTowerVectors, TNXEquals, False)
        TNXEquals = If(Me.reportSettings.ReportPrintUserLoadVectors = other.reportSettings.ReportPrintUserLoadVectors, TNXEquals, False)
        TNXEquals = If(Me.reportSettings.ReportPrintPressures = other.reportSettings.ReportPrintPressures, TNXEquals, False)
        TNXEquals = If(Me.reportSettings.ReportPrintAppurtForces = other.reportSettings.ReportPrintAppurtForces, TNXEquals, False)
        TNXEquals = If(Me.reportSettings.ReportPrintGuyForces = other.reportSettings.ReportPrintGuyForces, TNXEquals, False)
        TNXEquals = If(Me.reportSettings.ReportPrintGuyStressing = other.reportSettings.ReportPrintGuyStressing, TNXEquals, False)
        TNXEquals = If(Me.reportSettings.ReportPrintDeflections = other.reportSettings.ReportPrintDeflections, TNXEquals, False)
        TNXEquals = If(Me.reportSettings.ReportPrintReactions = other.reportSettings.ReportPrintReactions, TNXEquals, False)
        TNXEquals = If(Me.reportSettings.ReportPrintStressChecks = other.reportSettings.ReportPrintStressChecks, TNXEquals, False)
        TNXEquals = If(Me.reportSettings.ReportPrintBoltChecks = other.reportSettings.ReportPrintBoltChecks, TNXEquals, False)
        TNXEquals = If(Me.reportSettings.ReportPrintInputGVerificationTables = other.reportSettings.ReportPrintInputGVerificationTables, TNXEquals, False)
        TNXEquals = If(Me.reportSettings.ReportPrintOutputGVerificationTables = other.reportSettings.ReportPrintOutputGVerificationTables, TNXEquals, False)
        TNXEquals = If(Me.options.cantileverPoles.SocketTopMount = other.options.cantileverPoles.SocketTopMount, TNXEquals, False)
        TNXEquals = If(Me.options.SRTakeCompression = other.options.SRTakeCompression, TNXEquals, False)
        TNXEquals = If(Me.options.AllLegPanelsSame = other.options.AllLegPanelsSame, TNXEquals, False)
        TNXEquals = If(Me.options.UseCombinedBoltCapacity = other.options.UseCombinedBoltCapacity, TNXEquals, False)
        TNXEquals = If(Me.options.SecHorzBracesLeg = other.options.SecHorzBracesLeg, TNXEquals, False)
        TNXEquals = If(Me.options.SortByComponent = other.options.SortByComponent, TNXEquals, False)
        TNXEquals = If(Me.options.SRCutEnds = other.options.SRCutEnds, TNXEquals, False)
        TNXEquals = If(Me.options.SRConcentric = other.options.SRConcentric, TNXEquals, False)
        TNXEquals = If(Me.options.CalcBlockShear = other.options.CalcBlockShear, TNXEquals, False)
        TNXEquals = If(Me.options.Use4SidedDiamondBracing = other.options.Use4SidedDiamondBracing, TNXEquals, False)
        TNXEquals = If(Me.options.TriangulateInnerBracing = other.options.TriangulateInnerBracing, TNXEquals, False)
        TNXEquals = If(Me.options.PrintCarrierNotes = other.options.PrintCarrierNotes, TNXEquals, False)
        TNXEquals = If(Me.options.AddIBCWindCase = other.options.AddIBCWindCase, TNXEquals, False)
        TNXEquals = If(Me.code.wind.UseStateCountyLookup = other.code.wind.UseStateCountyLookup, TNXEquals, False)
        TNXEquals = If(Me.code.wind.State = other.code.wind.State, TNXEquals, False)
        TNXEquals = If(Me.code.wind.County = other.code.wind.County, TNXEquals, False)
        TNXEquals = If(Me.options.LegBoltsAtTop = other.options.LegBoltsAtTop, TNXEquals, False)
        TNXEquals = If(Me.options.cantileverPoles.PrintMonopoleAtIncrements = other.options.cantileverPoles.PrintMonopoleAtIncrements, TNXEquals, False)
        TNXEquals = If(Me.options.UseTIA222Exemptions_MinBracingResistance = other.options.UseTIA222Exemptions_MinBracingResistance, TNXEquals, False)
        TNXEquals = If(Me.options.UseTIA222Exemptions_TensionSplice = other.options.UseTIA222Exemptions_TensionSplice, TNXEquals, False)
        TNXEquals = If(Me.options.IgnoreKLryFor60DegAngleLegs = other.options.IgnoreKLryFor60DegAngleLegs, TNXEquals, False)
        TNXEquals = If(Me.code.wind.ASCE_7_10_WindData = other.code.wind.ASCE_7_10_WindData, TNXEquals, False)
        TNXEquals = If(Me.code.wind.ASCE_7_10_ConvertWindToASD = other.code.wind.ASCE_7_10_ConvertWindToASD, TNXEquals, False)
        TNXEquals = If(Me.solutionSettings.SolutionUsePDelta = other.solutionSettings.SolutionUsePDelta, TNXEquals, False)
        TNXEquals = If(Me.options.UseFeedlineTorque = other.options.UseFeedlineTorque, TNXEquals, False)
        TNXEquals = If(Me.options.UsePinnedElements = other.options.UsePinnedElements, TNXEquals, False)
        TNXEquals = If(Me.code.wind.UseMaxKz = other.code.wind.UseMaxKz, TNXEquals, False)
        TNXEquals = If(Me.options.UseRigidIndex = other.options.UseRigidIndex, TNXEquals, False)
        TNXEquals = If(Me.options.UseTrueCable = other.options.UseTrueCable, TNXEquals, False)
        TNXEquals = If(Me.options.UseASCELy = other.options.UseASCELy, TNXEquals, False)
        TNXEquals = If(Me.options.CalcBracingForces = other.options.CalcBracingForces, TNXEquals, False)
        TNXEquals = If(Me.options.IgnoreBracingFEA = other.options.IgnoreBracingFEA, TNXEquals, False)
        TNXEquals = If(Me.options.cantileverPoles.UseSubCriticalFlow = other.options.cantileverPoles.UseSubCriticalFlow, TNXEquals, False)
        TNXEquals = If(Me.options.cantileverPoles.AssumePoleWithNoAttachments = other.options.cantileverPoles.AssumePoleWithNoAttachments, TNXEquals, False)
        TNXEquals = If(Me.options.cantileverPoles.AssumePoleWithShroud = other.options.cantileverPoles.AssumePoleWithShroud, TNXEquals, False)
        TNXEquals = If(Me.options.cantileverPoles.PoleCornerRadiusKnown = other.options.cantileverPoles.PoleCornerRadiusKnown, TNXEquals, False)
        TNXEquals = If(Me.solutionSettings.SolutionMinStiffness = other.solutionSettings.SolutionMinStiffness, TNXEquals, False)
        TNXEquals = If(Me.solutionSettings.SolutionMaxStiffness = other.solutionSettings.SolutionMaxStiffness, TNXEquals, False)
        TNXEquals = If(Me.solutionSettings.SolutionMaxCycles = other.solutionSettings.SolutionMaxCycles, TNXEquals, False)
        TNXEquals = If(Me.solutionSettings.SolutionPower = other.solutionSettings.SolutionPower, TNXEquals, False)
        TNXEquals = If(Me.solutionSettings.SolutionTolerance = other.solutionSettings.SolutionTolerance, TNXEquals, False)
        TNXEquals = If(Me.options.cantileverPoles.CantKFactor = other.options.cantileverPoles.CantKFactor, TNXEquals, False)
        TNXEquals = If(Me.options.misclOptions.RadiusSampleDist = other.options.misclOptions.RadiusSampleDist, TNXEquals, False)
        TNXEquals = If(Me.options.BypassStabilityChecks = other.options.BypassStabilityChecks, TNXEquals, False)
        TNXEquals = If(Me.options.UseWindProjection = other.options.UseWindProjection, TNXEquals, False)
        TNXEquals = If(Me.code.ice.UseIceEscalation = other.code.ice.UseIceEscalation, TNXEquals, False)
        TNXEquals = If(Me.options.UseDishCoeff = other.options.UseDishCoeff, TNXEquals, False)
        TNXEquals = If(Me.options.AutoCalcTorqArmArea = other.options.AutoCalcTorqArmArea, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDirOption = other.options.windDirections.WindDirOption, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir0_0 = other.options.windDirections.WindDir0_0, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir0_1 = other.options.windDirections.WindDir0_1, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir0_2 = other.options.windDirections.WindDir0_2, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir0_3 = other.options.windDirections.WindDir0_3, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir0_4 = other.options.windDirections.WindDir0_4, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir0_5 = other.options.windDirections.WindDir0_5, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir0_6 = other.options.windDirections.WindDir0_6, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir0_7 = other.options.windDirections.WindDir0_7, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir0_8 = other.options.windDirections.WindDir0_8, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir0_9 = other.options.windDirections.WindDir0_9, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir0_10 = other.options.windDirections.WindDir0_10, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir0_11 = other.options.windDirections.WindDir0_11, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir0_12 = other.options.windDirections.WindDir0_12, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir0_13 = other.options.windDirections.WindDir0_13, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir0_14 = other.options.windDirections.WindDir0_14, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir0_15 = other.options.windDirections.WindDir0_15, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir1_0 = other.options.windDirections.WindDir1_0, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir1_1 = other.options.windDirections.WindDir1_1, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir1_2 = other.options.windDirections.WindDir1_2, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir1_3 = other.options.windDirections.WindDir1_3, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir1_4 = other.options.windDirections.WindDir1_4, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir1_5 = other.options.windDirections.WindDir1_5, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir1_6 = other.options.windDirections.WindDir1_6, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir1_7 = other.options.windDirections.WindDir1_7, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir1_8 = other.options.windDirections.WindDir1_8, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir1_9 = other.options.windDirections.WindDir1_9, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir1_10 = other.options.windDirections.WindDir1_10, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir1_11 = other.options.windDirections.WindDir1_11, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir1_12 = other.options.windDirections.WindDir1_12, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir1_13 = other.options.windDirections.WindDir1_13, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir1_14 = other.options.windDirections.WindDir1_14, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir1_15 = other.options.windDirections.WindDir1_15, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir2_0 = other.options.windDirections.WindDir2_0, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir2_1 = other.options.windDirections.WindDir2_1, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir2_2 = other.options.windDirections.WindDir2_2, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir2_3 = other.options.windDirections.WindDir2_3, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir2_4 = other.options.windDirections.WindDir2_4, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir2_5 = other.options.windDirections.WindDir2_5, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir2_6 = other.options.windDirections.WindDir2_6, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir2_7 = other.options.windDirections.WindDir2_7, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir2_8 = other.options.windDirections.WindDir2_8, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir2_9 = other.options.windDirections.WindDir2_9, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir2_10 = other.options.windDirections.WindDir2_10, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir2_11 = other.options.windDirections.WindDir2_11, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir2_12 = other.options.windDirections.WindDir2_12, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir2_13 = other.options.windDirections.WindDir2_13, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir2_14 = other.options.windDirections.WindDir2_14, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.WindDir2_15 = other.options.windDirections.WindDir2_15, TNXEquals, False)
        TNXEquals = If(Me.options.windDirections.SuppressWindPatternLoading = other.options.windDirections.SuppressWindPatternLoading, TNXEquals, False)

        Return TNXEquals
    End Function

    Public Overrides Sub Clear()
        Me.feedLines.Clear()
        Me.discreteLoads.Clear()
        Me.dishes.Clear()
        Me.userForces.Clear()
        Me.otherLines.Clear()
        Me.geometry.Clear()
    End Sub
End Class

Public Class DefaultERItxtValues
    Private newList As New List(Of String)

#Region "Variables Not part of TNX Class"
    Public ReadOnly Property SIUnits() As List(Of String)
        Get
            newList.Clear()
            newList.Add("Length = m")
            newList.Add("LengthPrec = 4")
            newList.Add("Coordinate = m")
            newList.Add("CoordinatePrec = 4")
            newList.Add("Force = N")
            newList.Add("ForcePrec = 2")
            newList.Add("Load = Nlm")
            newList.Add("LoadPrec = 2")
            newList.Add("Moment = N - m")
            newList.Add("MomentPrec = 0")
            newList.Add("Properties = mm")
            newList.Add("PropertiesPrec = 0")
            newList.Add("Pressure = kPa")
            newList.Add("PressurePrec = 2")
            newList.Add("Velocity = kph")
            newList.Add("VelocityPrec = 0")
            newList.Add("Displacement = mm")
            newList.Add("DisplacementPrec = 2")
            newList.Add("Mass = kg")
            newList.Add("MassPrec = 2")
            newList.Add("Acceleration = G")
            newList.Add("AccelerationPrec = 2")
            newList.Add("Stress = kPa")
            newList.Add("StressPrec = 0")
            newList.Add("Density = N / m3")
            newList.Add("DensityPrec = 2")
            newList.Add("UnitWt = Nlm")
            newList.Add("UnitWtPrec = 2")
            newList.Add("Strength = kPa")
            newList.Add("StrengthPrec = 0")
            newList.Add("Modulus = kPa")
            newList.Add("ModulusPrec = 0")
            newList.Add("Temperature = C")
            newList.Add("TemperaturePrec = 0")
            newList.Add("Printer = mm")
            newList.Add("PrinterPrec = 0")
            newList.Add("Rotation = rad")
            newList.Add("RotationPrec = 4")
            newList.Add("Spacing = m")
            newList.Add("SpacingPrec = 4")
            newList.Add("Architectural = No")

            Return newList
        End Get
    End Property

    Public ReadOnly Property CodeValues As List(Of String)
        Get
            newList.Clear()
            newList.Add("AluminumCode = Aluminum Assoc.")
            newList.Add("TimberCode = AFPA NDS 1993")
            newList.Add("ConcreteCode = ACI USD")
            newList.Add("MasonryCode = UBC Unity")
            newList.Add("BuildingCode = User Defined")
            newList.Add("PatternLive = Yes")
            newList.Add("PostalCode = 53223")
            newList.Add("PatternPerCent = 100")
            newList.Add("LoadsActPositiveDown = Yes")
            newList.Add("IncludeSelfWeight = Yes")
            newList.Add("UseLiveLoadReduction = Yes")
            newList.Add("PreCompositeConstructionLoad = 10")
            newList.Add("ConstructionLoad = 10")
            newList.Add("PreCompositeDeadLoad = 10")
            newList.Add("PostCompositeDeadLoad = 10")
            newList.Add("RoofLiveLoad = 20")
            newList.Add("FloorLiveLoad = 40")
            newList.Add("SnowLoad = 20")
            newList.Add("WindLoad = 20")
            newList.Add("UBCSeismicZone = 4")
            newList.Add("UBCSeismicI = 1")
            newList.Add("UBCSeismicS = 1")
            newList.Add("UBCSeismicRw = 8")
            newList.Add("UBCWindExposureCategory = B")
            newList.Add("UBCWindI = 1")
            newList.Add("UBCWindSpeed = 70")
            newList.Add("UBCWindOpenDoor = No")
            newList.Add("UBCSnowLoad = 20")
            newList.Add("UBCConcentratedForce = 0")
            newList.Add("UBCSnowI = 1")
            newList.Add("UBCSnowCe = 0.7")
            newList.Add("UBCIsGarage = No")
            newList.Add("UBCUsePostal = Yes")
            newList.Add("SBCCISeismicAv = 0.4")
            newList.Add("SBCCISeismicAa = 0.4")
            newList.Add("SBCCISeismicS = 1")
            newList.Add("SBCCISeismicR = 8")
            newList.Add("SBCCISeismicCd = 4")
            newList.Add("SBCCIWindExposureCategory = A")
            newList.Add("SBCCIWindI = 1")
            newList.Add("SBCCIWindSpeed = 70")
            newList.Add("SBCCIWindOpenDoor = No")
            newList.Add("SBCCISnowLoad = 20")
            newList.Add("SBCCIConcentratedForce = 0")
            newList.Add("SBCCISnowI = 1")
            newList.Add("SBCCISnowCe = 0.7")
            newList.Add("SBCCISnowCt = 1")
            newList.Add("SBCCIIsGarage = No")
            newList.Add("SBCCIUsePostal = Yes")
            newList.Add("BOCASeismicAv = 0.4")
            newList.Add("BOCASeismicAa = 0.4")
            newList.Add("BOCASeismicS = 1")
            newList.Add("BOCASeismicR = 8")
            newList.Add("BOCASeismicCd = 4")
            newList.Add("BOCAWindExposureCategory = A")
            newList.Add("BOCAWindI = 1")
            newList.Add("BOCAWindSpeed = 70")
            newList.Add("BOCAWindOpenDoor = No")
            newList.Add("BOCASnowLoad = 20")
            newList.Add("BOCAConcentratedForce = 0")
            newList.Add("BOCASnowI = 1")
            newList.Add("BOCASnowCe = 0.7")
            newList.Add("BOCAIsGarage = No")
            newList.Add("BOCAUsePostal = Yes")

            Return newList
        End Get
    End Property

    Public ReadOnly Property AntennaPole As List(Of String)
        Get
            newList.Clear()
            newList.Add("AntennaPoleType=0")
            newList.Add("AntennaPoleHt=0.000000")
            newList.Add("AntennaPoleIx=0.000000")
            newList.Add("AntennaPoleIy=0.000000")
            newList.Add("AntennaPoleEmod=0.000000")
            newList.Add("AntennaPoleWtNoIce=0.000000")
            newList.Add("AntennaPoleWtIce=0.000000")
            newList.Add("AntennaPoleWindNoIce=0.000000")
            newList.Add("AntennaPoleWindIce=0.000000")
            newList.Add("AntennaPoleWindService=0.000000")
            newList.Add("AntennaPoleCaAaNoIce=0.000000")
            newList.Add("AntennaPoleCaAaIce=0.000000")

            Return newList
        End Get
    End Property

    Public ReadOnly Property Supp As List(Of String)
        Get
            newList.Clear()
            newList.Add("SuppElev1=0.000000")
            newList.Add("SuppNone1=Yes")
            newList.Add("SuppLegA1=No")
            newList.Add("SuppLegB1=No")
            newList.Add("SuppLegC1=No")
            newList.Add("SuppLegD1=No")
            newList.Add("SuppVert1=No")
            newList.Add("SuppElev2=0.000000")
            newList.Add("SuppNone2=Yes")
            newList.Add("SuppLegA2=No")
            newList.Add("SuppLegB2=No")
            newList.Add("SuppLegC2=No")
            newList.Add("SuppLegD2=No")
            newList.Add("SuppVert2=No")
            newList.Add("SuppNone3=Yes")
            newList.Add("SuppLegA3=No")
            newList.Add("SuppLegB3=No")
            newList.Add("SuppLegC3=No")
            newList.Add("SuppLegD3=No")
            newList.Add("SuppVert3=No")

            Return newList
        End Get
    End Property

    Public ReadOnly Property MRSS As List(Of String)
        Get
            newList.Clear()
            newList.Add("MRSSNumSpiderLegs=0")
            newList.Add("MRSSUseMitredLegs=No")
            newList.Add("MRSSSpiderType=Tube")
            newList.Add("MRSSSpiderSize=HSS8x8x3/8")
            newList.Add("MRSSSpiderGradeVal=36")
            newList.Add("MRSSSpiderGrade=A36")
            newList.Add("MRSSSpacerThick=0.250000")
            newList.Add("MRSSSpiderRadius=15.000000")
            newList.Add("MRSSClampRingHt=15.000000")
            newList.Add("MRSSSpiderMaxPoleMoment=0.000000")

            Return newList
        End Get
    End Property

    Public ReadOnly Property Beacon As List(Of String)
        Get
            newList.Clear()
            newList.Add("BeaconDiameter=0.000000")
            newList.Add("BeaconWtNoIce=0.000000")
            newList.Add("BeaconWtIce=0.000000")
            newList.Add("BeaconWindNoIce=0.000000")
            newList.Add("BeaconWindIce=0.000000")
            newList.Add("BeaconWindService=0.000000")
            newList.Add("BeaconCaAaNoIce=0.000000")
            newList.Add("BeaconCaAaIce=0.000000")

            Return newList
        End Get
    End Property

    Public ReadOnly Property BasePlate As List(Of String)
        Get
            newList.Clear()
            newList.Add("BasePlateDataPlateGradeVal=36")
            newList.Add("BasePlateDataPlateGrade=A36")
            newList.Add("BasePlateDataSquareBasePlate=No")
            newList.Add("BasePlateDataGrouted=No")
            newList.Add("BasePlateDataAnchorBoltGrade=A325X")
            newList.Add("BasePlateDataBasePlateType=Plain Plate")
            newList.Add("BasePlateDataNumBolts=12")
            newList.Add("BasePlateDataNumBoltsPerStiffener=1")
            newList.Add("BasePlateDataAnchorBoltSize=1.500000")
            newList.Add("BasePlateDataAnchorBoltCircle=0.000000")
            newList.Add("BasePlateDataEmbedmentLength=24.000000")
            newList.Add("BasePlateDataGroutSpace=2.000000")
            newList.Add("BasePlateDataThick=0.000000")
            newList.Add("BasePlateDataInnerDiameter=0.000000")
            newList.Add("BasePlateDataOuterDiameter=0.000000")
            newList.Add("BasePlateDataClippedCorner=0.000000")
            newList.Add("BasePlateDataStiffenerThick=0.000000")
            newList.Add("BasePlateDataStiffenerHeight=0.000000")
            newList.Add("BasePlateDatafpc=4.000000")

            Return newList
        End Get
    End Property

    Public ReadOnly Property DesignChoicesList As List(Of String)
        Get
            newList.Clear()
            newList.Add("DesignChoicesList=0 1")
            newList.Add("DesignChoicesList=1 1")
            newList.Add("DesignChoicesList=2 1")
            newList.Add("DesignChoicesList=3 1")
            newList.Add("DesignChoicesList=4 1")
            newList.Add("DesignChoicesList=5 1")
            newList.Add("DesignChoicesList=6 1")
            newList.Add("DesignChoicesList=7 1")
            newList.Add("DesignChoicesList=8 1")
            newList.Add("DesignChoicesList=9 1")
            newList.Add("DesignChoicesList=10 1")
            newList.Add("DesignChoicesList=11 1")
            newList.Add("DesignChoicesList=12 1")
            newList.Add("DesignChoicesList=13 1")
            newList.Add("DesignChoicesList=14 1")
            newList.Add("DesignChoicesList=15 1")
            newList.Add("DesignChoicesList=16 1")
            newList.Add("DesignChoicesList=17 1")
            newList.Add("DesignChoicesList=18 1")
            newList.Add("DesignChoicesList=19 1")
            newList.Add("DesignChoicesList=20 1")
            newList.Add("DesignChoicesList=21 1")
            newList.Add("DesignChoicesList=22 1")
            newList.Add("DesignChoicesList=23 1")
            newList.Add("DesignChoicesList=24 1")
            newList.Add("DesignChoicesList=25 1")
            newList.Add("DesignChoicesList=26 1")
            newList.Add("DesignChoicesList=27 1")
            newList.Add("DesignChoicesList=28 1")
            newList.Add("DesignChoicesList=29 1")
            newList.Add("DesignChoicesList=30 1")
            newList.Add("DesignChoicesList=31 1")
            newList.Add("DesignChoicesList=32 1")
            newList.Add("DesignChoicesList=33 1")
            newList.Add("DesignChoicesList=34 1")
            newList.Add("DesignChoicesList=35 1")
            newList.Add("DesignChoicesList=36 1")
            newList.Add("DesignChoicesList=37 1")
            newList.Add("DesignChoicesList=38 1")
            newList.Add("DesignChoicesList=39 1")
            newList.Add("DesignChoicesList=40 1")
            newList.Add("DesignChoicesList=41 1")
            newList.Add("DesignChoicesList=42 1")
            newList.Add("DesignChoicesList=43 1")
            newList.Add("DesignChoicesList=44 1")
            newList.Add("DesignChoicesList=45 1")
            newList.Add("DesignChoicesList=46 1")
            newList.Add("DesignChoicesList=47 1")
            newList.Add("DesignChoicesList=48 1")
            newList.Add("DesignChoicesList=49 1")
            newList.Add("DesignChoicesList=50 1")
            newList.Add("DesignChoicesList=51 1")
            newList.Add("DesignChoicesList=52 1")
            newList.Add("DesignChoicesList=53 1")
            newList.Add("DesignChoicesList=54 1")
            newList.Add("DesignChoicesList=55 1")
            newList.Add("DesignChoicesList=56 1")
            newList.Add("DesignChoicesList=57 1")
            newList.Add("DesignChoicesList=58 1")
            newList.Add("DesignChoicesList=59 1")
            newList.Add("DesignChoicesList=60 1")
            newList.Add("DesignChoicesList=61 1")
            newList.Add("DesignChoicesList=62 1")
            newList.Add("DesignChoicesList=63 1")
            newList.Add("DesignChoicesList=64 1")
            newList.Add("DesignChoicesList=65 1")
            newList.Add("DesignChoicesList=66 1")
            newList.Add("DesignChoicesList=67 1")
            newList.Add("DesignChoicesList=68 1")
            newList.Add("DesignChoicesList=69 1")
            newList.Add("DesignChoicesList=70 1")
            newList.Add("DesignChoicesList=71 1")
            newList.Add("DesignChoicesList=72 1")
            newList.Add("DesignChoicesList=73 1")
            newList.Add("DesignChoicesList=74 1")
            newList.Add("DesignChoicesList=75 1")
            newList.Add("DesignChoicesList=76 1")
            newList.Add("DesignChoicesList=77 1")
            newList.Add("DesignChoicesList=78 1")
            newList.Add("DesignChoicesList=79 1")
            newList.Add("DesignChoicesList=80 1")
            newList.Add("DesignChoicesList=81 1")
            newList.Add("DesignChoicesList=82 1")
            newList.Add("DesignChoicesList=83 1")
            newList.Add("DesignChoicesList=84 1")
            newList.Add("DesignChoicesList=85 1")
            newList.Add("DesignChoicesList=86 1")
            newList.Add("DesignChoicesList=87 1")
            newList.Add("DesignChoicesList=88 1")
            newList.Add("DesignChoicesList=89 1")
            newList.Add("DesignChoicesList=90 1")
            newList.Add("DesignChoicesList=91 1")
            newList.Add("DesignChoicesList=92 1")
            newList.Add("DesignChoicesList=93 1")
            newList.Add("DesignChoicesList=94 1")
            newList.Add("DesignChoicesList=95 1")
            newList.Add("DesignChoicesList=96 1")
            newList.Add("DesignChoicesList=97 1")
            newList.Add("DesignChoicesList=98 1")
            newList.Add("DesignChoicesList=99 1")
            newList.Add("DesignChoicesList=100 1")
            newList.Add("DesignChoicesList=101 1")
            newList.Add("DesignChoicesList=102 1")
            newList.Add("DesignChoicesList=103 1")
            newList.Add("DesignChoicesList=104 1")
            newList.Add("DesignChoicesList=105 1")
            newList.Add("DesignChoicesList=106 1")
            newList.Add("DesignChoicesList=107 1")
            newList.Add("DesignChoicesList=108 1")
            newList.Add("DesignChoicesList=109 1")
            newList.Add("DesignChoicesList=110 1")
            newList.Add("DesignChoicesList=111 1")
            newList.Add("DesignChoicesList=112 1")
            newList.Add("DesignChoicesList=113 1")

            Return newList
        End Get
    End Property

    Public ReadOnly Property FTG As List(Of String)
        Get
            newList.Clear()
            newList.Add("FtgDataFtgType=Pier Broms Eqn")
            newList.Add("FtgDataPierDiameter=60.000000")
            newList.Add("FtgDataPierLength=30.000000")
            newList.Add("FtgDataSpreadPierWidth=24.000000")
            newList.Add("FtgDataSpreadFtgDepth=0.000000")
            newList.Add("FtgDataSpreadFtgWidth=8.000000")
            newList.Add("FtgDataSpreadFtgThickness=24.000000")
            newList.Add("FtgDataTieSpacing=12.000000")
            newList.Add("FtgDataAllowBrgPressure=3000.000000")
            newList.Add("FtgDataConcretefpc=3.500000")
            newList.Add("FtgDataRebarFy=60.000000")
            newList.Add("FtgDataPierBarSize=#10")
            newList.Add("FtgDataTopBarSize=#8")
            newList.Add("FtgDataBotBarSize=#8")
            newList.Add("FtgDataTieSize=#4")
            newList.Add("FtgDataNumBotBars=0")
            newList.Add("FtgDataNumPierBars=0")
            newList.Add("FtgDataNumTopBars=0")
            newList.Add("FtgDataSoilLayers0soilType=Medium Clay")
            newList.Add("FtgDataSoilLayers0layerDepth=50.000000")
            newList.Add("FtgDataSoilLayers0density=120.000000")
            newList.Add("FtgDataSoilLayers0shearStrength=0.000000")
            newList.Add("FtgDataSoilLayers0phiAngle=30.000000")
            newList.Add("FtgDataSoilLayers1soilType=")
            newList.Add("FtgDataSoilLayers1layerDepth=0.000000")
            newList.Add("FtgDataSoilLayers1density=0.000000")
            newList.Add("FtgDataSoilLayers1shearStrength=0.000000")
            newList.Add("FtgDataSoilLayers1phiAngle=0.000000")
            newList.Add("FtgDataSoilLayers2soilType=")
            newList.Add("FtgDataSoilLayers2layerDepth=0.000000")
            newList.Add("FtgDataSoilLayers2density=0.000000")
            newList.Add("FtgDataSoilLayers2shearStrength=0.000000")
            newList.Add("FtgDataSoilLayers2phiAngle=0.000000")
            newList.Add("FtgDataSoilLayers3soilType=")
            newList.Add("FtgDataSoilLayers3layerDepth=0.000000")
            newList.Add("FtgDataSoilLayers3density=0.000000")
            newList.Add("FtgDataSoilLayers3shearStrength=0.000000")
            newList.Add("FtgDataSoilLayers3phiAngle=0.000000")
            newList.Add("FtgDataSoilLayers4soilType=")
            newList.Add("FtgDataSoilLayers4layerDepth=0.000000")
            newList.Add("FtgDataSoilLayers4density=0.000000")
            newList.Add("FtgDataSoilLayers4shearStrength=0.000000")
            newList.Add("FtgDataSoilLayers4phiAngle=0.000000")

            Return newList
        End Get
    End Property

    Public ReadOnly Property CostData As List(Of String)
        Get
            newList.Clear()
            newList.Add("CostData=0.500000 0.500000 0.500000 0.500000 0.500000 0.500000 0.500000 0.500000 0.500000 0.500000 0.500000 0.500000 0.500000 0.500000")
            newList.Add("HandlingCostData=0.500000 0.500000 0.500000 0.500000 0.500000 0.500000 0.500000 0.500000 0.500000 0.500000 0.500000 0.500000 0.500000 0.500000")

            Return newList
        End Get
    End Property

    Public ReadOnly Property MissingSettings As List(Of String)
        Get
            newList.Clear()
            newList.Add("MastMultiplier=1.000000")
            newList.Add("EnforceAllOffsetsVer2=Yes")
            newList.Add("UseMaxKz=No")
            newList.Add("UseMastResponse=No")
            newList.Add("Use_TIA_222_G_Addendum2_MP=Yes")
            newList.Add("UseTopTakeup=Yes")
            newList.Add("EliminateBase=No")

            Return newList
        End Get
    End Property
#End Region

#Region "Extra Chronos Stuff"
    Public ReadOnly Property LoadCombos As List(Of String)
        Get
            newList.Clear()
            newList.Add("1 1")
            newList.Add("Dead Only")
            newList.Add("1.0*1")
            newList.Add("2 1")
            newList.Add("1.2 Dead+1.0 Wind 0 deg - No Ice")
            newList.Add("1.2*1,1.0*2")
            newList.Add("3 1")
            newList.Add("0.9 Dead+1.0 Wind 0 deg - No Ice")
            newList.Add("0.9*1,1.0*2")
            newList.Add("4 1")
            newList.Add("1.2 Dead+1.0 Wind 30 deg - No Ice")
            newList.Add("1.2*1,1.0*3")
            newList.Add("5 1")
            newList.Add("0.9 Dead+1.0 Wind 30 deg - No Ice")
            newList.Add("0.9*1,1.0*3")
            newList.Add("6 1")
            newList.Add("1.2 Dead+1.0 Wind 60 deg - No Ice")
            newList.Add("1.2*1,1.0*4")
            newList.Add("7 1")
            newList.Add("0.9 Dead+1.0 Wind 60 deg - No Ice")
            newList.Add("0.9*1,1.0*4")
            newList.Add("8 1")
            newList.Add("1.2 Dead+1.0 Wind 90 deg - No Ice")
            newList.Add("1.2*1,1.0*5")
            newList.Add("9 1")
            newList.Add("0.9 Dead+1.0 Wind 90 deg - No Ice")
            newList.Add("0.9*1,1.0*5")
            newList.Add("10 1")
            newList.Add("1.2 Dead+1.0 Wind 120 deg - No Ice")
            newList.Add("1.2*1,1.0*6")
            newList.Add("11 1")
            newList.Add("0.9 Dead+1.0 Wind 120 deg - No Ice")
            newList.Add("0.9*1,1.0*6")
            newList.Add("12 1")
            newList.Add("1.2 Dead+1.0 Wind 150 deg - No Ice")
            newList.Add("1.2*1,1.0*7")
            newList.Add("13 1")
            newList.Add("0.9 Dead+1.0 Wind 150 deg - No Ice")
            newList.Add("0.9*1,1.0*7")
            newList.Add("14 1")
            newList.Add("1.2 Dead+1.0 Wind 180 deg - No Ice")
            newList.Add("1.2*1,1.0*8")
            newList.Add("15 1")
            newList.Add("0.9 Dead+1.0 Wind 180 deg - No Ice")
            newList.Add("0.9*1,1.0*8")
            newList.Add("16 1")
            newList.Add("1.2 Dead+1.0 Wind 210 deg - No Ice")
            newList.Add("1.2*1,1.0*9")
            newList.Add("17 1")
            newList.Add("0.9 Dead+1.0 Wind 210 deg - No Ice")
            newList.Add("0.9*1,1.0*9")
            newList.Add("18 1")
            newList.Add("1.2 Dead+1.0 Wind 240 deg - No Ice")
            newList.Add("1.2*1,1.0*10")
            newList.Add("19 1")
            newList.Add("0.9 Dead+1.0 Wind 240 deg - No Ice")
            newList.Add("0.9*1,1.0*10")
            newList.Add("20 1")
            newList.Add("1.2 Dead+1.0 Wind 270 deg - No Ice")
            newList.Add("1.2*1,1.0*11")
            newList.Add("21 1")
            newList.Add("0.9 Dead+1.0 Wind 270 deg - No Ice")
            newList.Add("0.9*1,1.0*11")
            newList.Add("22 1")
            newList.Add("1.2 Dead+1.0 Wind 300 deg - No Ice")
            newList.Add("1.2*1,1.0*12")
            newList.Add("23 1")
            newList.Add("0.9 Dead+1.0 Wind 300 deg - No Ice")
            newList.Add("0.9*1,1.0*12")
            newList.Add("24 1")
            newList.Add("1.2 Dead+1.0 Wind 330 deg - No Ice")
            newList.Add("1.2*1,1.0*13")
            newList.Add("25 1")
            newList.Add("0.9 Dead+1.0 Wind 330 deg - No Ice")
            newList.Add("0.9*1,1.0*13")
            newList.Add("26 1")
            newList.Add("1.2 Dead+1.0 Ice+1.0 Temp")
            newList.Add("1.2*1,1.0*14,1.0*15")
            newList.Add("27 1")
            newList.Add("1.2 Dead+1.0 Wind 0 deg+1.0 Ice+1.0 Temp")
            newList.Add("1.2*1,1.0*16,1.0*14,1.0*15")
            newList.Add("28 1")
            newList.Add("1.2 Dead+1.0 Wind 30 deg+1.0 Ice+1.0 Temp")
            newList.Add("1.2*1,1.0*17,1.0*14,1.0*15")
            newList.Add("29 1")
            newList.Add("1.2 Dead+1.0 Wind 60 deg+1.0 Ice+1.0 Temp")
            newList.Add("1.2*1,1.0*18,1.0*14,1.0*15")
            newList.Add("30 1")
            newList.Add("1.2 Dead+1.0 Wind 90 deg+1.0 Ice+1.0 Temp")
            newList.Add("1.2*1,1.0*19,1.0*14,1.0*15")
            newList.Add("31 1")
            newList.Add("1.2 Dead+1.0 Wind 120 deg+1.0 Ice+1.0 Temp")
            newList.Add("1.2*1,1.0*20,1.0*14,1.0*15")
            newList.Add("32 1")
            newList.Add("1.2 Dead+1.0 Wind 150 deg+1.0 Ice+1.0 Temp")
            newList.Add("1.2*1,1.0*21,1.0*14,1.0*15")
            newList.Add("33 1")
            newList.Add("1.2 Dead+1.0 Wind 180 deg+1.0 Ice+1.0 Temp")
            newList.Add("1.2*1,1.0*22,1.0*14,1.0*15")
            newList.Add("34 1")
            newList.Add("1.2 Dead+1.0 Wind 210 deg+1.0 Ice+1.0 Temp")
            newList.Add("1.2*1,1.0*23,1.0*14,1.0*15")
            newList.Add("35 1")
            newList.Add("1.2 Dead+1.0 Wind 240 deg+1.0 Ice+1.0 Temp")
            newList.Add("1.2*1,1.0*24,1.0*14,1.0*15")
            newList.Add("36 1")
            newList.Add("1.2 Dead+1.0 Wind 270 deg+1.0 Ice+1.0 Temp")
            newList.Add("1.2*1,1.0*25,1.0*14,1.0*15")
            newList.Add("37 1")
            newList.Add("1.2 Dead+1.0 Wind 300 deg+1.0 Ice+1.0 Temp")
            newList.Add("1.2*1,1.0*26,1.0*14,1.0*15")
            newList.Add("38 1")
            newList.Add("1.2 Dead+1.0 Wind 330 deg+1.0 Ice+1.0 Temp")
            newList.Add("1.2*1,1.0*27,1.0*14,1.0*15")
            newList.Add("39 1")
            newList.Add("Dead+Wind 0 deg - Service")
            newList.Add("1.0*1,1.0*28")
            newList.Add("40 1")
            newList.Add("Dead+Wind 30 deg - Service")
            newList.Add("1.0*1,1.0*29")
            newList.Add("41 1")
            newList.Add("Dead+Wind 60 deg - Service")
            newList.Add("1.0*1,1.0*30")
            newList.Add("42 1")
            newList.Add("Dead+Wind 90 deg - Service")
            newList.Add("1.0*1,1.0*31")
            newList.Add("43 1")
            newList.Add("Dead+Wind 120 deg - Service")
            newList.Add("1.0*1,1.0*32")
            newList.Add("44 1")
            newList.Add("Dead+Wind 150 deg - Service")
            newList.Add("1.0*1,1.0*33")
            newList.Add("45 1")
            newList.Add("Dead+Wind 180 deg - Service")
            newList.Add("1.0*1,1.0*34")
            newList.Add("46 1")
            newList.Add("Dead+Wind 210 deg - Service")
            newList.Add("1.0*1,1.0*35")
            newList.Add("47 1")
            newList.Add("Dead+Wind 240 deg - Service")
            newList.Add("1.0*1,1.0*36")
            newList.Add("48 1")
            newList.Add("Dead+Wind 270 deg - Service")
            newList.Add("1.0*1,1.0*37")
            newList.Add("49 1")
            newList.Add("Dead+Wind 300 deg - Service")
            newList.Add("1.0*1,1.0*38")
            newList.Add("50 1")
            newList.Add("Dead+Wind 330 deg - Service")
            newList.Add("1.0*1,1.0*39")

            Return newList
        End Get
    End Property

    Public ReadOnly Property LoadEnvs As List(Of String)
        Get
            newList.Clear()
            newList.Add("1 1")
            newList.Add("Dead Only")
            newList.Add("1")
            newList.Add("2 1")
            newList.Add("1.2 Dead+1.0 Wind 0 deg - No Ice")
            newList.Add("2")
            newList.Add("3 1")
            newList.Add("0.9 Dead+1.0 Wind 0 deg - No Ice")
            newList.Add("3")
            newList.Add("4 1")
            newList.Add("1.2 Dead+1.0 Wind 30 deg - No Ice")
            newList.Add("4")
            newList.Add("5 1")
            newList.Add("0.9 Dead+1.0 Wind 30 deg - No Ice")
            newList.Add("5")
            newList.Add("6 1")
            newList.Add("1.2 Dead+1.0 Wind 60 deg - No Ice")
            newList.Add("6")
            newList.Add("7 1")
            newList.Add("0.9 Dead+1.0 Wind 60 deg - No Ice")
            newList.Add("7")
            newList.Add("8 1")
            newList.Add("1.2 Dead+1.0 Wind 90 deg - No Ice")
            newList.Add("8")
            newList.Add("9 1")
            newList.Add("0.9 Dead+1.0 Wind 90 deg - No Ice")
            newList.Add("9")
            newList.Add("10 1")
            newList.Add("1.2 Dead+1.0 Wind 120 deg - No Ice")
            newList.Add("10")
            newList.Add("11 1")
            newList.Add("0.9 Dead+1.0 Wind 120 deg - No Ice")
            newList.Add("11")
            newList.Add("12 1")
            newList.Add("1.2 Dead+1.0 Wind 150 deg - No Ice")
            newList.Add("12")
            newList.Add("13 1")
            newList.Add("0.9 Dead+1.0 Wind 150 deg - No Ice")
            newList.Add("13")
            newList.Add("14 1")
            newList.Add("1.2 Dead+1.0 Wind 180 deg - No Ice")
            newList.Add("14")
            newList.Add("15 1")
            newList.Add("0.9 Dead+1.0 Wind 180 deg - No Ice")
            newList.Add("15")
            newList.Add("16 1")
            newList.Add("1.2 Dead+1.0 Wind 210 deg - No Ice")
            newList.Add("16")
            newList.Add("17 1")
            newList.Add("0.9 Dead+1.0 Wind 210 deg - No Ice")
            newList.Add("17")
            newList.Add("18 1")
            newList.Add("1.2 Dead+1.0 Wind 240 deg - No Ice")
            newList.Add("18")
            newList.Add("19 1")
            newList.Add("0.9 Dead+1.0 Wind 240 deg - No Ice")
            newList.Add("19")
            newList.Add("20 1")
            newList.Add("1.2 Dead+1.0 Wind 270 deg - No Ice")
            newList.Add("20")
            newList.Add("21 1")
            newList.Add("0.9 Dead+1.0 Wind 270 deg - No Ice")
            newList.Add("21")
            newList.Add("22 1")
            newList.Add("1.2 Dead+1.0 Wind 300 deg - No Ice")
            newList.Add("22")
            newList.Add("23 1")
            newList.Add("0.9 Dead+1.0 Wind 300 deg - No Ice")
            newList.Add("23")
            newList.Add("24 1")
            newList.Add("1.2 Dead+1.0 Wind 330 deg - No Ice")
            newList.Add("24")
            newList.Add("25 1")
            newList.Add("0.9 Dead+1.0 Wind 330 deg - No Ice")
            newList.Add("25")
            newList.Add("26 1")
            newList.Add("1.2 Dead+1.0 Ice+1.0 Temp")
            newList.Add("26")
            newList.Add("27 1")
            newList.Add("1.2 Dead+1.0 Wind 0 deg+1.0 Ice+1.0 Temp")
            newList.Add("27")
            newList.Add("28 1")
            newList.Add("1.2 Dead+1.0 Wind 30 deg+1.0 Ice+1.0 Temp")
            newList.Add("28")
            newList.Add("29 1")
            newList.Add("1.2 Dead+1.0 Wind 60 deg+1.0 Ice+1.0 Temp")
            newList.Add("29")
            newList.Add("30 1")
            newList.Add("1.2 Dead+1.0 Wind 90 deg+1.0 Ice+1.0 Temp")
            newList.Add("30")
            newList.Add("31 1")
            newList.Add("1.2 Dead+1.0 Wind 120 deg+1.0 Ice+1.0 Temp")
            newList.Add("31")
            newList.Add("32 1")
            newList.Add("1.2 Dead+1.0 Wind 150 deg+1.0 Ice+1.0 Temp")
            newList.Add("32")
            newList.Add("33 1")
            newList.Add("1.2 Dead+1.0 Wind 180 deg+1.0 Ice+1.0 Temp")
            newList.Add("33")
            newList.Add("34 1")
            newList.Add("1.2 Dead+1.0 Wind 210 deg+1.0 Ice+1.0 Temp")
            newList.Add("34")
            newList.Add("35 1")
            newList.Add("1.2 Dead+1.0 Wind 240 deg+1.0 Ice+1.0 Temp")
            newList.Add("35")
            newList.Add("36 1")
            newList.Add("1.2 Dead+1.0 Wind 270 deg+1.0 Ice+1.0 Temp")
            newList.Add("36")
            newList.Add("37 1")
            newList.Add("1.2 Dead+1.0 Wind 300 deg+1.0 Ice+1.0 Temp")
            newList.Add("37")
            newList.Add("38 1")
            newList.Add("1.2 Dead+1.0 Wind 330 deg+1.0 Ice+1.0 Temp")
            newList.Add("38")
            newList.Add("39 1")
            newList.Add("Dead+Wind 0 deg - Service")
            newList.Add("39")
            newList.Add("40 1")
            newList.Add("Dead+Wind 30 deg - Service")
            newList.Add("40")
            newList.Add("41 1")
            newList.Add("Dead+Wind 60 deg - Service")
            newList.Add("41")
            newList.Add("42 1")
            newList.Add("Dead+Wind 90 deg - Service")
            newList.Add("42")
            newList.Add("43 1")
            newList.Add("Dead+Wind 120 deg - Service")
            newList.Add("43")
            newList.Add("44 1")
            newList.Add("Dead+Wind 150 deg - Service")
            newList.Add("44")
            newList.Add("45 1")
            newList.Add("Dead+Wind 180 deg - Service")
            newList.Add("45")
            newList.Add("46 1")
            newList.Add("Dead+Wind 210 deg - Service")
            newList.Add("46")
            newList.Add("47 1")
            newList.Add("Dead+Wind 240 deg - Service")
            newList.Add("47")
            newList.Add("48 1")
            newList.Add("Dead+Wind 270 deg - Service")
            newList.Add("48")
            newList.Add("49 1")
            newList.Add("Dead+Wind 300 deg - Service")
            newList.Add("49")
            newList.Add("50 1")
            newList.Add("Dead+Wind 330 deg - Service")
            newList.Add("50")

            Return newList
        End Get
    End Property

    Public ReadOnly Property LoadCases As List(Of String)
        Get
            newList.Clear()
            newList.Add("1 1 3")
            newList.Add("Dead")
            newList.Add("2 1 9")
            newList.Add("No Ice Wind 0 deg")
            newList.Add("3 1 9")
            newList.Add("No Ice Wind 30 deg")
            newList.Add("4 1 9")
            newList.Add("No Ice Wind 60 deg")
            newList.Add("5 1 9")
            newList.Add("No Ice Wind 90 deg")
            newList.Add("6 1 9")
            newList.Add("No Ice Wind 120 deg")
            newList.Add("7 1 9")
            newList.Add("No Ice Wind 150 deg")
            newList.Add("8 1 9")
            newList.Add("No Ice Wind 180 deg")
            newList.Add("9 1 9")
            newList.Add("No Ice Wind 210 deg")
            newList.Add("10 1 9")
            newList.Add("No Ice Wind 240 deg")
            newList.Add("11 1 9")
            newList.Add("No Ice Wind 270 deg")
            newList.Add("12 1 9")
            newList.Add("No Ice Wind 300 deg")
            newList.Add("13 1 9")
            newList.Add("No Ice Wind 330 deg")
            newList.Add("14 1 5")
            newList.Add("Ice")
            newList.Add("15 1 5")
            newList.Add("Temperature Drop")
            newList.Add("16 1 9")
            newList.Add("Ice Wind 0 deg")
            newList.Add("17 1 9")
            newList.Add("Ice Wind 30 deg")
            newList.Add("18 1 9")
            newList.Add("Ice Wind 60 deg")
            newList.Add("19 1 9")
            newList.Add("Ice Wind 90 deg")
            newList.Add("20 1 9")
            newList.Add("Ice Wind 120 deg")
            newList.Add("21 1 9")
            newList.Add("Ice Wind 150 deg")
            newList.Add("22 1 9")
            newList.Add("Ice Wind 180 deg")
            newList.Add("23 1 9")
            newList.Add("Ice Wind 210 deg")
            newList.Add("24 1 9")
            newList.Add("Ice Wind 240 deg")
            newList.Add("25 1 9")
            newList.Add("Ice Wind 270 deg")
            newList.Add("26 1 9")
            newList.Add("Ice Wind 300 deg")
            newList.Add("27 1 9")
            newList.Add("Ice Wind 330 deg")
            newList.Add("28 1 9")
            newList.Add("Service Wind 0 deg")
            newList.Add("29 1 9")
            newList.Add("Service Wind 30 deg")
            newList.Add("30 1 9")
            newList.Add("Service Wind 60 deg")
            newList.Add("31 1 9")
            newList.Add("Service Wind 90 deg")
            newList.Add("32 1 9")
            newList.Add("Service Wind 120 deg")
            newList.Add("33 1 9")
            newList.Add("Service Wind 150 deg")
            newList.Add("34 1 9")
            newList.Add("Service Wind 180 deg")
            newList.Add("35 1 9")
            newList.Add("Service Wind 210 deg")
            newList.Add("36 1 9")
            newList.Add("Service Wind 240 deg")
            newList.Add("37 1 9")
            newList.Add("Service Wind 270 deg")
            newList.Add("38 1 9")
            newList.Add("Service Wind 300 deg")
            newList.Add("39 1 9")
            newList.Add("Service Wind 330 deg")

            Return newList
        End Get
    End Property

    Public ReadOnly Property OutputOptions As List(Of String)
        Get
            newList.Clear()
            newList.Add("#device")
            newList.Add("#echo")
            newList.Add("#case_list")
            newList.Add("#combination_list")
            newList.Add("#envelope_list")
            newList.Add("#prn_report")
            newList.Add("#reband")
            newList.Add("#restart")
            newList.Add("#save_stiff")
            newList.Add("#anatype")
            newList.Add("#if anatype = 2")
            newList.Add("#  pdelta_only")
            newList.Add("#  axial_effect")
            newList.Add("#  tol min_stif power cycle start stop")
            newList.Add("#if anatype = 3 or 4")
            newList.Add("#  dtype (1,2,3,4: mode shapes,modal time history,response spectrum, direct time history)")
            newList.Add("#  if dtype = 1")
            newList.Add("#    no_modes cut_off")
            newList.Add("#    prn_modes")
            newList.Add("#  if dtype = 2")
            newList.Add("#    no_modes cut_off step_size no_steps interval damp_ratio gmo_no")
            newList.Add("#    prn_modes")
            newList.Add("#  if dtype = 3")
            newList.Add("#    no_modes cut_off damp_ratio  gmo_no mode_comb dir_comb")
            newList.Add("#    prn_modes")
            newList.Add("#  if dtype = 4")
            newList.Add("#    step_size no_steps interval alpha_damp beta_damp gmo_no")
            newList.Add("#displ_list")
            newList.Add("#reacts_list")
            newList.Add("#forces_list")
            newList.Add("chronos.txt")
            newList.Add("Yes")
            newList.Add("All")
            newList.Add("All")
            newList.Add("All")
            newList.Add("0")
            newList.Add("N")
            newList.Add("No")
            newList.Add("No")
            newList.Add("1")
            newList.Add("All")
            newList.Add("All")

            Return newList
        End Get
    End Property
#End Region
End Class