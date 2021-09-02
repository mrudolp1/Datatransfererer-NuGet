Option Strict On
Option Compare Binary 'Trying to speed up parsing the TNX file by using Binary Text comparison instead of Text Comparison

Imports System.ComponentModel
Imports System.Data
Imports System.IO

Partial Public Class tnxModel

    Private prop_settings As New tnxSettings()
    Private prop_solutionSettings As New tnxSolutionSettings()
    Private prop_MTOSettings As New tnxMTOSettings()
    Private prop_reportSettings As New tnxReportSettings()
    Private prop_CCIReport As New tnxCCIReport()
    Private prop_code As New tnxCode()
    Private prop_options As New tnxOptions()
    Private prop_geometry As New tnxGeometry()
    Private prop_feedLines As New List(Of tnxFeedLine)
    Private prop_discreteLoads As New List(Of tnxDiscreteLoad)
    Private prop_dishes As New List(Of tnxDish)
    Private prop_userForces As New List(Of tnxUserForce)
    Private prop_otherLines As New List(Of String())




    <Category("TNX"), Description(""), DisplayName("Settings")>
    Public Property settings() As tnxSettings
        Get
            Return Me.prop_settings
        End Get
        Set
            Me.prop_settings = Value
        End Set
    End Property
    <Category("TNX"), Description(""), DisplayName("Solution Settings")>
    Public Property solutionSettings() As tnxSolutionSettings
        Get
            Return Me.prop_solutionSettings
        End Get
        Set
            Me.prop_solutionSettings = Value
        End Set
    End Property
    <Category("TNX"), Description(""), DisplayName("MTO Settings")>
    Public Property MTOSettings() As tnxMTOSettings
        Get
            Return Me.prop_MTOSettings
        End Get
        Set
            Me.prop_MTOSettings = Value
        End Set
    End Property
    <Category("TNX"), Description(""), DisplayName("Report Settings")>
    Public Property reportSettings() As tnxReportSettings
        Get
            Return Me.prop_reportSettings
        End Get
        Set
            Me.prop_reportSettings = Value
        End Set
    End Property
    <Category("TNX"), Description(""), DisplayName("CCI Report")>
    Public Property CCIReport() As tnxCCIReport
        Get
            Return Me.prop_CCIReport
        End Get
        Set
            Me.prop_CCIReport = Value
        End Set
    End Property
    <Category("TNX"), Description(""), DisplayName("Code")>
    Public Property code() As tnxCode
        Get
            Return Me.prop_code
        End Get
        Set
            Me.prop_code = Value
        End Set
    End Property
    <Category("TNX"), Description(""), DisplayName("Options")>
    Public Property options() As tnxOptions
        Get
            Return Me.prop_options
        End Get
        Set
            Me.prop_options = Value
        End Set
    End Property
    <Category("TNX"), Description(""), DisplayName("Geometry")>
    Public Property geometry() As tnxGeometry
        Get
            Return Me.prop_geometry
        End Get
        Set
            Me.prop_geometry = Value
        End Set
    End Property
    <Category("TNX"), Description(""), DisplayName("Feed Lines")>
    Public Property feedLines() As List(Of tnxFeedLine)
        Get
            Return Me.prop_feedLines
        End Get
        Set
            Me.prop_feedLines = Value
        End Set
    End Property
    <Category("TNX"), Description(""), DisplayName("Discrete Loads")>
    Public Property discreteLoads() As List(Of tnxDiscreteLoad)
        Get
            Return Me.prop_discreteLoads
        End Get
        Set
            Me.prop_discreteLoads = Value
        End Set
    End Property
    <Category("TNX"), Description(""), DisplayName("Dishes")>
    Public Property dishes() As List(Of tnxDish)
        Get
            Return Me.prop_dishes
        End Get
        Set
            Me.prop_dishes = Value
        End Set
    End Property
    <Category("TNX"), Description(""), DisplayName("User Forces")>
    Public Property userForces() As List(Of tnxUserForce)
        Get
            Return Me.prop_userForces
        End Get
        Set
            Me.prop_userForces = Value
        End Set
    End Property
    <Category("TNX"), Description(""), DisplayName("All the other stuff")>
    Public Property otherLines() As List(Of String())
        Get
            Return Me.prop_otherLines
        End Get
        Set
            Me.prop_otherLines = Value
        End Set
    End Property

    Public Sub New()
        'Leave method empty
    End Sub

    <Category("Constructor"), Description("Create TNX object from TNX file.")>
    Public Sub New(ByVal tnxPath As String)

        Dim tnxVar As String
        Dim tnxValue As String
        Dim recIndex As Integer
        Dim recordUSUnits As Boolean = False
        Dim sectionFilter As String = ""
        Dim caseFilter As String = ""

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
                Case caseFilter = ""
                    ''''These are all the individual options for the eri file. They are not part of a record which there may be multiple of.'''
                    Select Case True
                            ''''Main Section Filters''''
                        Case tnxVar.Equals("NumAntennaRecs")
                            Try
                                sectionFilter = "Antenna"
                                Me.otherLines.Add(New String() {tnxVar, tnxValue})
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("NumTowerRecs")
                            Try
                                sectionFilter = "Tower"
                                Me.otherLines.Add(New String() {tnxVar, tnxValue})
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("NumGuyRecs")
                            Try
                                sectionFilter = "Guy"
                                Me.otherLines.Add(New String() {tnxVar, tnxValue})
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("NumFeedLineRecs")
                            Try
                                sectionFilter = "FeedLine"
                                Me.otherLines.Add(New String() {tnxVar, tnxValue})
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("NumTowerLoadRecs")
                            Try
                                sectionFilter = "Discrete"
                                Me.otherLines.Add(New String() {tnxVar, tnxValue})
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("NumDishRecs")
                            Try
                                sectionFilter = "Dish"
                                Me.otherLines.Add(New String() {tnxVar, tnxValue})
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("NumUserForceRecs")
                            Try
                                sectionFilter = "UserForce"
                                Me.otherLines.Add(New String() {tnxVar, tnxValue})
                            Catch ex As Exception
                            End Try
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
                            End Try
                        Case tnxVar.Equals("LengthPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Length.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Coordinate")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Coordinate = New tnxCoordinateUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("CoordinatePrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Coordinate.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Force")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Force = New tnxForceUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ForcePrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Force.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Load")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Load = New tnxLoadUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("LoadPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Load.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Moment")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Moment = New tnxMomentUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("MomentPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Moment.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Properties")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Properties = New tnxPropertiesUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("PropertiesPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Properties.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Pressure")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Pressure = New tnxPressureUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("PressurePrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Pressure.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Velocity")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Velocity = New tnxVelocityUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("VelocityPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Velocity.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Displacement")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Displacement = New tnxDisplacementUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DisplacementPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Displacement.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Mass")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Mass = New tnxMassUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("MassPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Mass.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Acceleration")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Acceleration = New tnxAccelerationUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AccelerationPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Acceleration.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Stress")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Stress = New tnxStressUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("StressPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Stress.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Density")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Density = New tnxDensityUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DensityPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Density.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UnitWt")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.UnitWt = New tnxUnitWTUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UnitWtPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.UnitWt.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Strength")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Strength = New tnxStrengthUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("StrengthPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Strength.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Modulus")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Modulus = New tnxModulusUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ModulusPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Modulus.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Temperature")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Temperature = New tnxTempUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TemperaturePrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Temperature.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Printer")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Printer = New tnxPrinterUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("PrinterPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Printer.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Rotation")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Rotation = New tnxRotationUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("RotationPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Rotation.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Spacing")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Spacing = New tnxSpacingUnit(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("SpacingPrec")
                            Try
                                If recordUSUnits Then
                                    Me.settings.USUnits.Spacing.precision = CInt(tnxValue)
                                Else
                                    Me.otherLines.Add(New String() {tnxVar, tnxValue})
                                End If
                            Catch ex As Exception
                            End Try
                    ''''Project Info Settings
                        Case tnxVar.Equals("DesignStandardSeries")
                            Try
                                Me.settings.projectInfo.DesignStandardSeries = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UnitsSystem")
                            Try
                                Me.settings.projectInfo.UnitsSystem = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ClientName")
                            Try
                                Me.settings.projectInfo.ClientName = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ProjectName")
                            Try
                                Me.settings.projectInfo.ProjectName = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ProjectNumber")
                            Try
                                Me.settings.projectInfo.ProjectNumber = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("CreatedBy")
                            Try
                                Me.settings.projectInfo.CreatedBy = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("CreatedOn")
                            Try
                                Me.settings.projectInfo.CreatedOn = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("LastUsedBy")
                            Try
                                Me.settings.projectInfo.LastUsedBy = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("LastUsedOn")
                            Try
                                Me.settings.projectInfo.LastUsedOn = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("VersionUsed")
                            Try
                                Me.settings.projectInfo.VersionUsed = tnxValue
                            Catch ex As Exception
                            End Try
                            '''User Info Settings
                        Case tnxVar.Equals("ViewerUserName")
                            Try
                                Me.settings.userInfo.ViewerUserName = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ViewerCompanyName")
                            Try
                                Me.settings.userInfo.ViewerCompanyName = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ViewerStreetAddress")
                            Try
                                Me.settings.userInfo.ViewerStreetAddress = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ViewerCityState")
                            Try
                                Me.settings.userInfo.ViewerCityState = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ViewerPhone")
                            Try
                                Me.settings.userInfo.ViewerPhone = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ViewerFAX")
                            Try
                                Me.settings.userInfo.ViewerFAX = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ViewerLogo")
                            Try
                                Me.settings.userInfo.ViewerLogo = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ViewerCompanyBitmap")
                            Try
                                Me.settings.userInfo.ViewerCompanyBitmap = tnxValue
                            Catch ex As Exception
                            End Try
                    ''''Code''''
                        Case tnxVar.Equals("DesignCode")
                            Try
                                Me.code.design.DesignCode = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ERIDesignMode")
                            Try
                                Me.code.design.ERIDesignMode = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DoInteraction")
                            Try
                                Me.code.design.DoInteraction = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DoHorzInteraction")
                            Try
                                Me.code.design.DoHorzInteraction = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DoDiagInteraction")
                            Try
                                Me.code.design.DoDiagInteraction = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UseMomentMagnification")
                            Try
                                Me.code.design.UseMomentMagnification = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UseCodeStressRatio")
                            Try
                                Me.code.design.UseCodeStressRatio = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AllowStressRatio")
                            Try
                                Me.code.design.AllowStressRatio = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AllowAntStressRatio")
                            Try
                                Me.code.design.AllowAntStressRatio = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UseCodeGuySF")
                            Try
                                Me.code.design.UseCodeGuySF = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuySF")
                            Try
                                Me.code.design.GuySF = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UseTIA222H_AnnexS")
                            Try
                                Me.code.design.UseTIA222H_AnnexS = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TIA_222_H_AnnexS_Ratio")
                            Try
                                Me.code.design.TIA_222_H_AnnexS_Ratio = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("PrintBitmaps")
                            Try
                                Me.code.design.PrintBitmaps = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("IceThickness")
                            Try
                                Me.code.ice.IceThickness = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("IceDensity")
                            Try
                                Me.code.ice.IceDensity = Me.settings.USUnits.Density.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UseModified_TIA_222_IceParameters")
                            Try
                                Me.code.ice.UseModified_TIA_222_IceParameters = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TIA_222_IceThicknessMultiplier")
                            Try
                                Me.code.ice.TIA_222_IceThicknessMultiplier = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DoNotUse_TIA_222_IceEscalation")
                            Try
                                Me.code.ice.DoNotUse_TIA_222_IceEscalation = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UseIceEscalation")
                            Try
                                Me.code.ice.UseIceEscalation = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TempDrop")
                            Try
                                Me.code.thermal.TempDrop = Me.settings.USUnits.Temperature.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GroutFc")
                            Try
                                Me.code.misclCode.GroutFc = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerBoltGrade")
                            Try
                                Me.code.misclCode.TowerBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerBoltMinEdgeDist")
                            Try
                                Me.code.misclCode.TowerBoltMinEdgeDist = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindSpeed")
                            Try
                                Me.code.wind.WindSpeed = Me.settings.USUnits.Velocity.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindSpeedIce")
                            Try
                                Me.code.wind.WindSpeedIce = Me.settings.USUnits.Velocity.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindSpeedService")
                            Try
                                Me.code.wind.WindSpeedService = Me.settings.USUnits.Velocity.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UseStateCountyLookup")
                            Try
                                Me.code.wind.UseStateCountyLookup = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("State")
                            Try
                                Me.code.wind.State = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("County")
                            Try
                                Me.code.wind.County = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UseMaxKz")
                            Try
                                Me.code.wind.UseMaxKz = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ASCE_7_10_WindData")
                            Try
                                Me.code.wind.ASCE_7_10_WindData = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ASCE_7_10_ConvertWindToASD")
                            Try
                                Me.code.wind.ASCE_7_10_ConvertWindToASD = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UseASCEWind")
                            Try
                                Me.code.wind.UseASCEWind = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AutoCalc_ASCE_GH")
                            Try
                                Me.code.wind.AutoCalc_ASCE_GH = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ASCE_ExposureCat")
                            Try
                                Me.code.wind.ASCE_ExposureCat = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ASCE_Year")
                            Try
                                Me.code.wind.ASCE_Year = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ASCEGh")
                            Try
                                Me.code.wind.ASCEGh = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ASCEI")
                            Try
                                Me.code.wind.ASCEI = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("CalcWindAt")
                            Try
                                Me.code.wind.CalcWindAt = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindCalcPoints")
                            Try
                                Me.code.wind.WindCalcPoints = Me.settings.USUnits.Length.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindExposure")
                            Try
                                Me.code.wind.WindExposure = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("StructureCategory")
                            Try
                                Me.code.wind.StructureCategory = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("RiskCategory")
                            Try
                                Me.code.wind.RiskCategory = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TopoCategory")
                            Try
                                Me.code.wind.TopoCategory = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("RSMTopographicFeature")
                            Try
                                Me.code.wind.RSMTopographicFeature = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("RSM_L")
                            Try
                                Me.code.wind.RSM_L = Me.settings.USUnits.Length.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("RSM_X")
                            Try
                                Me.code.wind.RSM_X = Me.settings.USUnits.Length.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("CrestHeight")
                            Try
                                Me.code.wind.CrestHeight = Me.settings.USUnits.Length.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TIA_222_H_TopoFeatureDownwind")
                            Try
                                Me.code.wind.TIA_222_H_TopoFeatureDownwind = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("BaseElevAboveSeaLevel")
                            Try
                                Me.code.wind.BaseElevAboveSeaLevel = Me.settings.USUnits.Length.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ConsiderRooftopSpeedUp")
                            Try
                                Me.code.wind.ConsiderRooftopSpeedUp = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("RooftopWS")
                            Try
                                Me.code.wind.RooftopWS = Me.settings.USUnits.Length.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("RooftopHS")
                            Try
                                Me.code.wind.RooftopHS = Me.settings.USUnits.Length.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("RooftopParapetHt")
                            Try
                                Me.code.wind.RooftopParapetHt = Me.settings.USUnits.Length.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("RooftopXB")
                            Try
                                Me.code.wind.RooftopXB = Me.settings.USUnits.Length.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindZone")
                            Try
                                Me.code.wind.WindZone = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("EIACWindMult")
                            Try
                                Me.code.wind.EIACWindMult = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("EIACWindMultIce")
                            Try
                                Me.code.wind.EIACWindMultIce = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("EIACIgnoreCableDrag")
                            Try
                                Me.code.wind.EIACIgnoreCableDrag = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("CSA_S37_RefVelPress")
                            Try
                                Me.code.wind.CSA_S37_RefVelPress = Me.settings.USUnits.Pressure.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("CSA_S37_ReliabilityClass")
                            Try
                                Me.code.wind.CSA_S37_ReliabilityClass = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("CSA_S37_ServiceabilityFactor")
                            Try
                                Me.code.wind.CSA_S37_ServiceabilityFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UseASCE7_10_Seismic_Lcomb")
                            Try
                                Me.code.seismic.UseASCE7_10_Seismic_Lcomb = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("SeismicSiteClass")
                            Try
                                Me.code.seismic.SeismicSiteClass = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("SeismicSs")
                            Try
                                Me.code.seismic.SeismicSs = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("SeismicS1")
                            Try
                                Me.code.seismic.SeismicS1 = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                    ''''Options''''
                        Case tnxVar.Equals("UseClearSpans")
                            Try
                                Me.options.UseClearSpans = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UseClearSpansKlr")
                            Try
                                Me.options.UseClearSpansKlr = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UseFeedlineAsCylinder")
                            Try
                                Me.options.UseFeedlineAsCylinder = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UseLegLoads")
                            Try
                                Me.options.UseLegLoads = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("SRTakeCompression")
                            Try
                                Me.options.SRTakeCompression = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AllLegPanelsSame")
                            Try
                                Me.options.AllLegPanelsSame = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UseCombinedBoltCapacity")
                            Try
                                Me.options.UseCombinedBoltCapacity = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("SecHorzBracesLeg")
                            Try
                                Me.options.SecHorzBracesLeg = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("SortByComponent")
                            Try
                                Me.options.SortByComponent = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("SRCutEnds")
                            Try
                                Me.options.SRCutEnds = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("SRConcentric")
                            Try
                                Me.options.SRConcentric = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("CalcBlockShear")
                            Try
                                Me.options.CalcBlockShear = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Use4SidedDiamondBracing")
                            Try
                                Me.options.Use4SidedDiamondBracing = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TriangulateInnerBracing")
                            Try
                                Me.options.TriangulateInnerBracing = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("PrintCarrierNotes")
                            Try
                                Me.options.PrintCarrierNotes = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AddIBCWindCase")
                            Try
                                Me.options.AddIBCWindCase = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("LegBoltsAtTop")
                            Try
                                Me.options.LegBoltsAtTop = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UseTIA222Exemptions_MinBracingResistance")
                            Try
                                Me.options.UseTIA222Exemptions_MinBracingResistance = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UseTIA222Exemptions_TensionSplice")
                            Try
                                Me.options.UseTIA222Exemptions_TensionSplice = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("IgnoreKLryFor60DegAngleLegs")
                            Try
                                Me.options.IgnoreKLryFor60DegAngleLegs = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UseFeedlineTorque")
                            Try
                                Me.options.UseFeedlineTorque = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UsePinnedElements")
                            Try
                                Me.options.UsePinnedElements = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UseRigidIndex")
                            Try
                                Me.options.UseRigidIndex = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UseTrueCable")
                            Try
                                Me.options.UseTrueCable = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UseASCELy")
                            Try
                                Me.options.UseASCELy = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("CalcBracingForces")
                            Try
                                Me.options.CalcBracingForces = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("IgnoreBracingFEA")
                            Try
                                Me.options.IgnoreBracingFEA = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("BypassStabilityChecks")
                            Try
                                Me.options.BypassStabilityChecks = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UseWindProjection")
                            Try
                                Me.options.UseWindProjection = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UseDishCoeff")
                            Try
                                Me.options.UseDishCoeff = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AutoCalcTorqArmArea")
                            Try
                                Me.options.AutoCalcTorqArmArea = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("MastVert")
                            'The units of this input are dependant on the TNX force unit setting and the TNX property unit setting. (ie K/ft or Force/Properties) 
                            'Don't use covertTOEDSDefault, instead multiply Properties mult and divide by Force Mult
                            Try
                                Me.options.foundationStiffness.MastVert = (CDbl(tnxValue)) * (Me.settings.USUnits.Properties.multiplier / Me.settings.USUnits.Force.multiplier)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("MastHorz")
                            'The units of this input are dependant on the TNX force unit setting and the TNX property unit setting. (ie K/ft or Force/Properties) 
                            'Don't use covertTOEDSDefault, instead multiply Properties mult and divide by Force Mult
                            Try
                                Me.options.foundationStiffness.MastHorz = (CDbl(tnxValue)) * (Me.settings.USUnits.Properties.multiplier / Me.settings.USUnits.Force.multiplier)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyVert")
                            'The units of this input are dependant on the TNX force unit setting and the TNX property unit setting. (ie K/ft or Force/Properties) 
                            'Don't use covertTOEDSDefault, instead multiply Properties mult and divide by Force Mult
                            Try
                                Me.options.foundationStiffness.GuyVert = (CDbl(tnxValue)) * (Me.settings.USUnits.Properties.multiplier / Me.settings.USUnits.Force.multiplier)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyHorz")
                            'The units of this input are dependant on the TNX force unit setting and the TNX property unit setting. (ie K/ft or Force/Properties) 
                            'Don't use covertTOEDSDefault, instead multiply Properties mult and divide by Force Mult
                            Try
                                Me.options.foundationStiffness.GuyHorz = (CDbl(tnxValue)) * (Me.settings.USUnits.Properties.multiplier / Me.settings.USUnits.Force.multiplier)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GirtOffset")
                            Try
                                Me.options.defaultGirtOffsets.GirtOffset = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GirtOffsetLatticedPole")
                            Try
                                Me.options.defaultGirtOffsets.GirtOffsetLatticedPole = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("OffsetBotGirt")
                            Try
                                Me.options.defaultGirtOffsets.OffsetBotGirt = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("CheckVonMises")
                            Try
                                Me.options.cantileverPoles.CheckVonMises = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("SocketTopMount")
                            Try
                                Me.options.cantileverPoles.SocketTopMount = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("PrintMonopoleAtIncrements")
                            Try
                                Me.options.cantileverPoles.PrintMonopoleAtIncrements = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UseSubCriticalFlow")
                            Try
                                Me.options.cantileverPoles.UseSubCriticalFlow = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AssumePoleWithNoAttachments")
                            Try
                                Me.options.cantileverPoles.AssumePoleWithNoAttachments = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AssumePoleWithShroud")
                            Try
                                Me.options.cantileverPoles.AssumePoleWithShroud = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("PoleCornerRadiusKnown")
                            Try
                                Me.options.cantileverPoles.PoleCornerRadiusKnown = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("CantKFactor")
                            Try
                                Me.options.cantileverPoles.CantKFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("HogRodTakeup")
                            Try
                                Me.options.misclOptions.HogRodTakeup = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("RadiusSampleDist")
                            Try
                                Me.options.misclOptions.RadiusSampleDist = Me.settings.USUnits.Length.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDirOption")
                            Try
                                Me.options.windDirections.WindDirOption = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir0_0")
                            Try
                                Me.options.windDirections.WindDir0_0 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir0_1")
                            Try
                                Me.options.windDirections.WindDir0_1 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir0_2")
                            Try
                                Me.options.windDirections.WindDir0_2 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir0_3")
                            Try
                                Me.options.windDirections.WindDir0_3 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir0_4")
                            Try
                                Me.options.windDirections.WindDir0_4 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir0_5")
                            Try
                                Me.options.windDirections.WindDir0_5 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir0_6")
                            Try
                                Me.options.windDirections.WindDir0_6 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir0_7")
                            Try
                                Me.options.windDirections.WindDir0_7 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir0_8")
                            Try
                                Me.options.windDirections.WindDir0_8 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir0_9")
                            Try
                                Me.options.windDirections.WindDir0_9 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir0_10")
                            Try
                                Me.options.windDirections.WindDir0_10 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir0_11")
                            Try
                                Me.options.windDirections.WindDir0_11 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir0_12")
                            Try
                                Me.options.windDirections.WindDir0_12 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir0_13")
                            Try
                                Me.options.windDirections.WindDir0_13 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir0_14")
                            Try
                                Me.options.windDirections.WindDir0_14 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir0_15")
                            Try
                                Me.options.windDirections.WindDir0_15 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir1_0")
                            Try
                                Me.options.windDirections.WindDir1_0 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir1_1")
                            Try
                                Me.options.windDirections.WindDir1_1 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir1_2")
                            Try
                                Me.options.windDirections.WindDir1_2 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir1_3")
                            Try
                                Me.options.windDirections.WindDir1_3 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir1_4")
                            Try
                                Me.options.windDirections.WindDir1_4 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir1_5")
                            Try
                                Me.options.windDirections.WindDir1_5 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir1_6")
                            Try
                                Me.options.windDirections.WindDir1_6 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir1_7")
                            Try
                                Me.options.windDirections.WindDir1_7 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir1_8")
                            Try
                                Me.options.windDirections.WindDir1_8 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir1_9")
                            Try
                                Me.options.windDirections.WindDir1_9 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir1_10")
                            Try
                                Me.options.windDirections.WindDir1_10 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir1_11")
                            Try
                                Me.options.windDirections.WindDir1_11 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir1_12")
                            Try
                                Me.options.windDirections.WindDir1_12 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir1_13")
                            Try
                                Me.options.windDirections.WindDir1_13 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir1_14")
                            Try
                                Me.options.windDirections.WindDir1_14 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir1_15")
                            Try
                                Me.options.windDirections.WindDir1_15 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir2_0")
                            Try
                                Me.options.windDirections.WindDir2_0 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir2_1")
                            Try
                                Me.options.windDirections.WindDir2_1 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir2_2")
                            Try
                                Me.options.windDirections.WindDir2_2 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir2_3")
                            Try
                                Me.options.windDirections.WindDir2_3 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir2_4")
                            Try
                                Me.options.windDirections.WindDir2_4 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir2_5")
                            Try
                                Me.options.windDirections.WindDir2_5 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir2_6")
                            Try
                                Me.options.windDirections.WindDir2_6 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir2_7")
                            Try
                                Me.options.windDirections.WindDir2_7 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir2_8")
                            Try
                                Me.options.windDirections.WindDir2_8 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir2_9")
                            Try
                                Me.options.windDirections.WindDir2_9 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir2_10")
                            Try
                                Me.options.windDirections.WindDir2_10 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir2_11")
                            Try
                                Me.options.windDirections.WindDir2_11 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir2_12")
                            Try
                                Me.options.windDirections.WindDir2_12 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir2_13")
                            Try
                                Me.options.windDirections.WindDir2_13 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir2_14")
                            Try
                                Me.options.windDirections.WindDir2_14 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("WindDir2_15")
                            Try
                                Me.options.windDirections.WindDir2_15 = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("SuppressWindPatternLoading")
                            Try
                                Me.options.windDirections.SuppressWindPatternLoading = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try

                    ''''General Geometry''''
                        Case tnxVar.Equals("TowerType")
                            Try
                                Me.geometry.TowerType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaType")
                            Try
                                Me.geometry.AntennaType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("OverallHeight")
                            Try
                                Me.geometry.OverallHeight = Me.settings.USUnits.Length.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("BaseElevation")
                            Try
                                Me.geometry.BaseElevation = Me.settings.USUnits.Length.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Lambda")
                            Try
                                Me.geometry.Lambda = Me.settings.USUnits.Spacing.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerTopFaceWidth")
                            Try
                                Me.geometry.TowerTopFaceWidth = Me.settings.USUnits.Spacing.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerBaseFaceWidth")
                            Try
                                Me.geometry.TowerBaseFaceWidth = Me.settings.USUnits.Spacing.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerTaper")
                            Try
                                Me.geometry.TowerTaper = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyedMonopoleBaseType")
                            Try
                                Me.geometry.GuyedMonopoleBaseType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TaperHeight")
                            Try
                                Me.geometry.TaperHeight = Me.settings.USUnits.Length.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("PivotHeight")
                            Try
                                Me.geometry.PivotHeight = Me.settings.USUnits.Length.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AutoCalcGH")
                            Try
                                Me.geometry.AutoCalcGH = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UserGHElev")
                            Try
                                Me.geometry.UserGHElev = Me.settings.USUnits.Length.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UseIndexPlate")
                            Try
                                Me.geometry.UseIndexPlate = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("EnterUserDefinedGhValues")
                            Try
                                Me.geometry.EnterUserDefinedGhValues = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("BaseTowerGhInput")
                            Try
                                Me.geometry.BaseTowerGhInput = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UpperStructureGhInput")
                            Try
                                Me.geometry.UpperStructureGhInput = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("EnterUserDefinedCgValues")
                            Try
                                Me.geometry.EnterUserDefinedCgValues = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("BaseTowerCgInput")
                            Try
                                Me.geometry.BaseTowerCgInput = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UpperStructureCgInput")
                            Try
                                Me.geometry.UpperStructureCgInput = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaFaceWidth")
                            Try
                                Me.geometry.AntennaFaceWidth = Me.settings.USUnits.Spacing.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UseTopTakeup")
                            Try
                                Me.geometry.UseTopTakeup = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ConstantSlope")
                            Try
                                Me.geometry.ConstantSlope = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("[End Application]")
                            Me.otherLines.Add(New String() {tnxVar})
                            Exit For
                        ''''Solution Options''''
                        Case tnxVar.Equals("SolutionUsePDelta")
                            Try
                                Me.solutionSettings.SolutionUsePDelta = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("SolutionMinStiffness")
                            Try
                                Me.solutionSettings.SolutionMinStiffness = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("SolutionMaxStiffness")
                            Try
                                Me.solutionSettings.SolutionMaxStiffness = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("SolutionMaxCycles")
                            Try
                                Me.solutionSettings.SolutionMaxCycles = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("SolutionPower")
                            Try
                                Me.solutionSettings.SolutionPower = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("SolutionTolerance")
                            Try
                                Me.solutionSettings.SolutionTolerance = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                            '''''MTO Settings''''
                        Case tnxVar.Equals("IncludeCapacityNote")
                            Try
                                Me.mtoSettings.IncludeCapacityNote = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("IncludeAppurtGraphics")
                            Try
                                Me.mtoSettings.IncludeAppurtGraphics = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DisplayNotes")
                            Try
                                Me.mtoSettings.DisplayNotes = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DisplayReactions")
                            Try
                                Me.mtoSettings.DisplayReactions = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DisplaySchedule")
                            Try
                                Me.mtoSettings.DisplaySchedule = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DisplayAppurtenanceTable")
                            Try
                                Me.mtoSettings.DisplayAppurtenanceTable = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DisplayMaterialStrengthTable")
                            Try
                                Me.mtoSettings.DisplayMaterialStrengthTable = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Notes")
                            Try
                                Me.MTOSettings.Notes.Add(tnxValue)
                            Catch ex As Exception
                            End Try
                        ''''Report Settings''''
                        Case tnxVar.Equals("ReportInputCosts")
                            Try
                                Me.reportSettings.ReportInputCosts = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ReportInputGeometry")
                            Try
                                Me.reportSettings.ReportInputGeometry = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ReportInputOptions")
                            Try
                                Me.reportSettings.ReportInputOptions = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ReportMaxForces")
                            Try
                                Me.reportSettings.ReportMaxForces = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ReportInputMap")
                            Try
                                Me.reportSettings.ReportInputMap = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("CostReportOutputType")
                            Try
                                Me.reportSettings.CostReportOutputType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("CapacityReportOutputType")
                            Try
                                Me.reportSettings.CapacityReportOutputType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ReportPrintForceTotals")
                            Try
                                Me.reportSettings.ReportPrintForceTotals = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ReportPrintForceDetails")
                            Try
                                Me.reportSettings.ReportPrintForceDetails = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ReportPrintMastVectors")
                            Try
                                Me.reportSettings.ReportPrintMastVectors = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ReportPrintAntPoleVectors")
                            Try
                                Me.reportSettings.ReportPrintAntPoleVectors = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ReportPrintDiscreteVectors")
                            Try
                                Me.reportSettings.ReportPrintDiscreteVectors = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ReportPrintDishVectors")
                            Try
                                Me.reportSettings.ReportPrintDishVectors = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ReportPrintFeedTowerVectors")
                            Try
                                Me.reportSettings.ReportPrintFeedTowerVectors = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ReportPrintUserLoadVectors")
                            Try
                                Me.reportSettings.ReportPrintUserLoadVectors = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ReportPrintPressures")
                            Try
                                Me.reportSettings.ReportPrintPressures = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ReportPrintAppurtForces")
                            Try
                                Me.reportSettings.ReportPrintAppurtForces = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ReportPrintGuyForces")
                            Try
                                Me.reportSettings.ReportPrintGuyForces = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ReportPrintGuyStressing")
                            Try
                                Me.reportSettings.ReportPrintGuyStressing = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ReportPrintDeflections")
                            Try
                                Me.reportSettings.ReportPrintDeflections = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ReportPrintReactions")
                            Try
                                Me.reportSettings.ReportPrintReactions = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ReportPrintStressChecks")
                            Try
                                Me.reportSettings.ReportPrintStressChecks = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ReportPrintBoltChecks")
                            Try
                                Me.reportSettings.ReportPrintBoltChecks = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ReportPrintInputGVerificationTables")
                            Try
                                Me.reportSettings.ReportPrintInputGVerificationTables = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ReportPrintOutputGVerificationTables")
                            Try
                                Me.reportSettings.ReportPrintOutputGVerificationTables = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        ''''CCI Report''''
                        Case tnxVar.Equals("sReportProjectNumber")
                            Try
                                Me.CCIReport.sReportProjectNumber = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportJobType")
                            Try
                                Me.CCIReport.sReportJobType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportCarrierName")
                            Try
                                Me.CCIReport.sReportCarrierName = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportCarrierSiteNumber")
                            Try
                                Me.CCIReport.sReportCarrierSiteNumber = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportCarrierSiteName")
                            Try
                                Me.CCIReport.sReportCarrierSiteName = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportSiteAddress")
                            Try
                                Me.CCIReport.sReportSiteAddress = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportLatitudeDegree")
                            Try
                                Me.CCIReport.sReportLatitudeDegree = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportLatitudeMinute")
                            Try
                                Me.CCIReport.sReportLatitudeMinute = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportLatitudeSecond")
                            Try
                                Me.CCIReport.sReportLatitudeSecond = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportLongitudeDegree")
                            Try
                                Me.CCIReport.sReportLongitudeDegree = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportLongitudeMinute")
                            Try
                                Me.CCIReport.sReportLongitudeMinute = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportLongitudeSecond")
                            Try
                                Me.CCIReport.sReportLongitudeSecond = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportLocalCodeRequirement")
                            Try
                                Me.CCIReport.sReportLocalCodeRequirement = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportSiteHistory")
                            Try
                                Me.CCIReport.sReportSiteHistory = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportTowerManufacturer")
                            Try
                                Me.CCIReport.sReportTowerManufacturer = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportMonthManufactured")
                            Try
                                Me.CCIReport.sReportMonthManufactured = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportYearManufactured")
                            Try
                                Me.CCIReport.sReportYearManufactured = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportOriginalSpeed")
                            Try
                                Me.CCIReport.sReportOriginalSpeed = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportOriginalCode")
                            Try
                                Me.CCIReport.sReportOriginalCode = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportTowerType")
                            Try
                                Me.CCIReport.sReportTowerType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportEngrName")
                            Try
                                Me.CCIReport.sReportEngrName = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportEngrTitle")
                            Try
                                Me.CCIReport.sReportEngrTitle = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportHQPhoneNumber")
                            Try
                                Me.CCIReport.sReportHQPhoneNumber = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportEmailAddress")
                            Try
                                Me.CCIReport.sReportEmailAddress = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportLogoPath")
                            Try
                                Me.CCIReport.sReportLogoPath = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportCCiContactName")
                            Try
                                Me.CCIReport.sReportCCiContactName = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportCCiAddress1")
                            Try
                                Me.CCIReport.sReportCCiAddress1 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportCCiAddress2")
                            Try
                                Me.CCIReport.sReportCCiAddress2 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportCCiBUNumber")
                            Try
                                Me.CCIReport.sReportCCiBUNumber = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportCCiSiteName")
                            Try
                                Me.CCIReport.sReportCCiSiteName = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportCCiJDENumber")
                            Try
                                Me.CCIReport.sReportCCiJDENumber = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportCCiWONumber")
                            Try
                                Me.CCIReport.sReportCCiWONumber = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportCCiPONumber")
                            Try
                                Me.CCIReport.sReportCCiPONumber = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportCCiAppNumber")
                            Try
                                Me.CCIReport.sReportCCiAppNumber = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportCCiRevNumber")
                            Try
                                Me.CCIReport.sReportCCiRevNumber = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportDocsProvided")
                            Try
                                Me.CCIReport.sReportDocsProvided.Add(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportRecommendations")
                            Try
                                Me.CCIReport.sReportRecommendations = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportAppurt1")
                            Try
                                Me.CCIReport.sReportAppurt1.Add(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportAppurt2")
                            Try
                                Me.CCIReport.sReportAppurt2.Add(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportAppurt3")
                            Try
                                Me.CCIReport.sReportAppurt3.Add(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportAddlCapacity")
                            Try
                                Me.CCIReport.sReportAddlCapacity.Add(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportAssumption")
                            Try
                                Me.CCIReport.sReportAssumption.Add(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportAppurt1Note1")
                            Try
                                Me.CCIReport.sReportAppurt1Note1 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportAppurt1Note2")
                            Try
                                Me.CCIReport.sReportAppurt1Note2 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportAppurt1Note3")
                            Try
                                Me.CCIReport.sReportAppurt1Note3 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportAppurt1Note4")
                            Try
                                Me.CCIReport.sReportAppurt1Note4 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportAppurt1Note5")
                            Try
                                Me.CCIReport.sReportAppurt1Note5 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportAppurt1Note6")
                            Try
                                Me.CCIReport.sReportAppurt1Note6 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportAppurt1Note7")
                            Try
                                Me.CCIReport.sReportAppurt1Note7 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportAppurt2Note1")
                            Try
                                Me.CCIReport.sReportAppurt2Note1 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportAppurt2Note2")
                            Try
                                Me.CCIReport.sReportAppurt2Note2 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportAppurt2Note3")
                            Try
                                Me.CCIReport.sReportAppurt2Note3 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportAppurt2Note4")
                            Try
                                Me.CCIReport.sReportAppurt2Note4 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportAppurt2Note5")
                            Try
                                Me.CCIReport.sReportAppurt2Note5 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportAppurt2Note6")
                            Try
                                Me.CCIReport.sReportAppurt2Note6 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportAppurt2Note7")
                            Try
                                Me.CCIReport.sReportAppurt2Note7 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportAddlCapacityNote1")
                            Try
                                Me.CCIReport.sReportAddlCapacityNote1 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportAddlCapacityNote2")
                            Try
                                Me.CCIReport.sReportAddlCapacityNote2 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportAddlCapacityNote3")
                            Try
                                Me.CCIReport.sReportAddlCapacityNote3 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("sReportAddlCapacityNote4")
                            Try
                                Me.CCIReport.sReportAddlCapacityNote4 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case Else ''''All other lines
                            If line.Contains("=") Then
                                Me.otherLines.Add(New String() {tnxVar, tnxValue})
                            Else
                                Me.otherLines.Add(New String() {tnxVar})
                            End If
                    End Select
                Case caseFilter = "Antenna"
                    ''''Antenna Rec (Upper Structure)''''
                    Select Case True
                        Case tnxVar.Equals("AntennaRec")
                            Try
                                recIndex = CInt(tnxValue) - 1
                                Me.geometry.upperStructure.Add(New tnxAntennaRecord())
                                Me.geometry.upperStructure(recIndex).AntennaRec = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaBraceType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaBraceType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaHeight")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaHeight = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaDiagonalSpacing")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagonalSpacing = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaDiagonalSpacingEx")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagonalSpacingEx = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaNumSections")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaNumSections = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaNumSesctions")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaNumSesctions = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaSectionLength")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaSectionLength = Me.settings.USUnits.Length.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaLegType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaLegType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaLegSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaLegSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaLegGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaLegGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaLegMatlGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaLegMatlGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaDiagonalGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagonalGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaDiagonalMatlGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagonalMatlGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaInnerBracingGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerBracingGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaInnerBracingMatlGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerBracingMatlGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaTopGirtGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTopGirtGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaTopGirtMatlGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTopGirtMatlGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaBotGirtGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaBotGirtGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaBotGirtMatlGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaBotGirtMatlGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaInnerGirtGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerGirtGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaInnerGirtMatlGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerGirtMatlGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaLongHorizontalGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaLongHorizontalGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaLongHorizontalMatlGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaLongHorizontalMatlGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaShortHorizontalGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaShortHorizontalGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaShortHorizontalMatlGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaShortHorizontalMatlGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaDiagonalType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagonalType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaDiagonalSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagonalSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaInnerBracingType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerBracingType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaInnerBracingSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerBracingSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaTopGirtType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTopGirtType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaTopGirtSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTopGirtSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaBotGirtType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaBotGirtType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaBotGirtSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaBotGirtSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaTopGirtOffset")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTopGirtOffset = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaBotGirtOffset")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaBotGirtOffset = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaHasKBraceEndPanels")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaHasKBraceEndPanels = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaHasHorizontals")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaHasHorizontals = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaLongHorizontalType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaLongHorizontalType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaLongHorizontalSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaLongHorizontalSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaShortHorizontalType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaShortHorizontalType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaShortHorizontalSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaShortHorizontalSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantMatlGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantMatlGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantDiagType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantDiagType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubDiagonalType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubDiagonalType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubHorizontalType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubHorizontalType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantVerticalType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantVerticalType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipDiagonalType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipDiagonalType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHorizontalSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHorizontalSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHorizontalSize2")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHorizontalSize2 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHorizontalSize3")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHorizontalSize3 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHorizontalSize4")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHorizontalSize4 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantDiagonalSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantDiagonalSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantDiagonalSize2")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantDiagonalSize2 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantDiagonalSize3")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantDiagonalSize3 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantDiagonalSize4")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantDiagonalSize4 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubHorizontalSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubHorizontalSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubDiagonalSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubDiagonalSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaSubDiagLocation")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaSubDiagLocation = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantVerticalSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantVerticalSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipDiagonalSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipDiagonalSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipDiagonalSize2")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipDiagonalSize2 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipDiagonalSize3")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipDiagonalSize3 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipDiagonalSize4")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipDiagonalSize4 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipSize2")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipSize2 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipSize3")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipSize3 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipSize4")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipSize4 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaNumInnerGirts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaNumInnerGirts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaInnerGirtType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerGirtType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaInnerGirtSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerGirtSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaPoleShapeType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaPoleShapeType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaPoleSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaPoleSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaPoleGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaPoleGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaPoleMatlGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaPoleMatlGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaPoleSpliceLength")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaPoleSpliceLength = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaTaperPoleNumSides")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTaperPoleNumSides = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaTaperPoleTopDiameter")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTaperPoleTopDiameter = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaTaperPoleBotDiameter")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTaperPoleBotDiameter = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaTaperPoleWallThickness")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTaperPoleWallThickness = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaTaperPoleBendRadius")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTaperPoleBendRadius = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaTaperPoleGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTaperPoleGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaTaperPoleMatlGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTaperPoleMatlGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaSWMult")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaSWMult = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaWPMult")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaWPMult = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaAutoCalcKSingleAngle")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaAutoCalcKSingleAngle = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaAutoCalcKSolidRound")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaAutoCalcKSolidRound = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaAfGusset")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaAfGusset = Me.settings.USUnits.Length.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaTfGusset")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTfGusset = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaGussetBoltEdgeDistance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaGussetBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaGussetGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaGussetGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaGussetMatlGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaGussetMatlGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaAfMult")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaAfMult = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaArMult")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaArMult = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaFlatIPAPole")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaFlatIPAPole = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRoundIPAPole")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRoundIPAPole = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaFlatIPALeg")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaFlatIPALeg = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRoundIPALeg")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRoundIPALeg = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaFlatIPAHorizontal")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaFlatIPAHorizontal = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRoundIPAHorizontal")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRoundIPAHorizontal = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaFlatIPADiagonal")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaFlatIPADiagonal = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRoundIPADiagonal")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRoundIPADiagonal = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaCSA_S37_SpeedUpFactor")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaCSA_S37_SpeedUpFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaKLegs")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKLegs = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaKXBracedDiags")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKXBracedDiags = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaKKBracedDiags")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKKBracedDiags = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaKZBracedDiags")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKZBracedDiags = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaKHorzs")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKHorzs = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaKSecHorzs")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKSecHorzs = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaKGirts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKGirts = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaKInners")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKInners = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaKXBracedDiagsY")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKXBracedDiagsY = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaKKBracedDiagsY")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKKBracedDiagsY = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaKZBracedDiagsY")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKZBracedDiagsY = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaKHorzsY")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKHorzsY = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaKSecHorzsY")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKSecHorzsY = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaKGirtsY")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKGirtsY = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaKInnersY")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKInnersY = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaKRedHorz")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKRedHorz = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaKRedDiag")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKRedDiag = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaKRedSubDiag")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKRedSubDiag = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaKRedSubHorz")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKRedSubHorz = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaKRedVert")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKRedVert = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaKRedHip")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKRedHip = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaKRedHipDiag")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKRedHipDiag = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaKTLX")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKTLX = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaKTLZ")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKTLZ = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaKTLLeg")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKTLLeg = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaInnerKTLX")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerKTLX = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaInnerKTLZ")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerKTLZ = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaInnerKTLLeg")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerKTLLeg = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaStitchBoltLocationHoriz")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaStitchBoltLocationHoriz = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaStitchBoltLocationDiag")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaStitchBoltLocationDiag = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaStitchSpacing")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaStitchSpacing = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaStitchSpacingHorz")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaStitchSpacingHorz = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaStitchSpacingDiag")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaStitchSpacingDiag = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaStitchSpacingRed")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaStitchSpacingRed = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaLegNetWidthDeduct")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaLegNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaLegUFactor")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaLegUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaDiagonalNetWidthDeduct")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagonalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaTopGirtNetWidthDeduct")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTopGirtNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaBotGirtNetWidthDeduct")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaBotGirtNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaInnerGirtNetWidthDeduct")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerGirtNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaHorizontalNetWidthDeduct")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaHorizontalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaShortHorizontalNetWidthDeduct")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaShortHorizontalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaDiagonalUFactor")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagonalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaTopGirtUFactor")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTopGirtUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaBotGirtUFactor")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaBotGirtUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaInnerGirtUFactor")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerGirtUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaHorizontalUFactor")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaHorizontalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaShortHorizontalUFactor")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaShortHorizontalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaLegConnType")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaLegConnType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaLegNumBolts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaLegNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaDiagonalNumBolts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagonalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaTopGirtNumBolts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTopGirtNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaBotGirtNumBolts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaBotGirtNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaInnerGirtNumBolts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerGirtNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaHorizontalNumBolts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaHorizontalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaShortHorizontalNumBolts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaShortHorizontalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaLegBoltGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaLegBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaLegBoltSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaLegBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaDiagonalBoltGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagonalBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaDiagonalBoltSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagonalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaTopGirtBoltGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTopGirtBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaTopGirtBoltSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTopGirtBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaBotGirtBoltGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaBotGirtBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaBotGirtBoltSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaBotGirtBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaInnerGirtBoltGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerGirtBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaInnerGirtBoltSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerGirtBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaHorizontalBoltGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaHorizontalBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaHorizontalBoltSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaHorizontalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaShortHorizontalBoltGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaShortHorizontalBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaShortHorizontalBoltSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaShortHorizontalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaLegBoltEdgeDistance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaLegBoltEdgeDistance = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaDiagonalBoltEdgeDistance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagonalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaTopGirtBoltEdgeDistance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTopGirtBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaBotGirtBoltEdgeDistance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaBotGirtBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaInnerGirtBoltEdgeDistance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerGirtBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaHorizontalBoltEdgeDistance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaHorizontalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaShortHorizontalBoltEdgeDistance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaShortHorizontalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaDiagonalGageG1Distance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagonalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaTopGirtGageG1Distance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTopGirtGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaBotGirtGageG1Distance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaBotGirtGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaInnerGirtGageG1Distance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaInnerGirtGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaHorizontalGageG1Distance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaHorizontalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaShortHorizontalGageG1Distance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaShortHorizontalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHorizontalBoltGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHorizontalBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHorizontalBoltSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHorizontalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHorizontalNumBolts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHorizontalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHorizontalBoltEdgeDistance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHorizontalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHorizontalGageG1Distance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHorizontalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHorizontalNetWidthDeduct")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHorizontalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHorizontalUFactor")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHorizontalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantDiagonalBoltGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantDiagonalBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantDiagonalBoltSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantDiagonalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantDiagonalNumBolts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantDiagonalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantDiagonalBoltEdgeDistance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantDiagonalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantDiagonalGageG1Distance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantDiagonalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantDiagonalNetWidthDeduct")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantDiagonalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantDiagonalUFactor")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantDiagonalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubDiagonalBoltGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubDiagonalBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubDiagonalBoltSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubDiagonalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubDiagonalNumBolts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubDiagonalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubDiagonalBoltEdgeDistance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubDiagonalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubDiagonalGageG1Distance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubDiagonalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubDiagonalNetWidthDeduct")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubDiagonalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubDiagonalUFactor")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubDiagonalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubHorizontalBoltGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubHorizontalBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubHorizontalBoltSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubHorizontalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubHorizontalNumBolts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubHorizontalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubHorizontalBoltEdgeDistance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubHorizontalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubHorizontalGageG1Distance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubHorizontalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubHorizontalNetWidthDeduct")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubHorizontalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantSubHorizontalUFactor")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantSubHorizontalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantVerticalBoltGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantVerticalBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantVerticalBoltSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantVerticalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantVerticalNumBolts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantVerticalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantVerticalBoltEdgeDistance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantVerticalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantVerticalGageG1Distance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantVerticalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantVerticalNetWidthDeduct")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantVerticalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantVerticalUFactor")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantVerticalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipBoltGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipBoltSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipNumBolts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipBoltEdgeDistance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipGageG1Distance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipNetWidthDeduct")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipUFactor")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipDiagonalBoltGrade")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipDiagonalBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipDiagonalBoltSize")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipDiagonalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipDiagonalNumBolts")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipDiagonalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipDiagonalBoltEdgeDistance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipDiagonalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipDiagonalGageG1Distance")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipDiagonalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipDiagonalNetWidthDeduct")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipDiagonalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaRedundantHipDiagonalUFactor")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaRedundantHipDiagonalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaDiagonalOutOfPlaneRestraint")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagonalOutOfPlaneRestraint = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaTopGirtOutOfPlaneRestraint")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaTopGirtOutOfPlaneRestraint = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaBottomGirtOutOfPlaneRestraint")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaBottomGirtOutOfPlaneRestraint = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaMidGirtOutOfPlaneRestraint")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaMidGirtOutOfPlaneRestraint = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaHorizontalOutOfPlaneRestraint")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaHorizontalOutOfPlaneRestraint = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaSecondaryHorizontalOutOfPlaneRestraint")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaSecondaryHorizontalOutOfPlaneRestraint = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaDiagOffsetNEY")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagOffsetNEY = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaDiagOffsetNEX")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagOffsetNEX = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaDiagOffsetPEY")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagOffsetPEY = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaDiagOffsetPEX")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaDiagOffsetPEX = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaKbraceOffsetNEY")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKbraceOffsetNEY = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaKbraceOffsetNEX")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKbraceOffsetNEX = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaKbraceOffsetPEY")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKbraceOffsetPEY = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AntennaKbraceOffsetPEX")
                            Try
                                Me.geometry.upperStructure(recIndex).AntennaKbraceOffsetPEX = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                    End Select
                Case caseFilter = "Tower"
                    ''''Tower Rec (Base Structure)''''
                    Select Case True
                        Case tnxVar.Equals("TowerRec")
                            Try
                                recIndex = CInt(tnxValue) - 1
                                Me.geometry.baseStructure.Add(New tnxTowerRecord())
                                Me.geometry.baseStructure(recIndex).TowerRec = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerDatabase")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDatabase = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerName")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerName = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerHeight")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerHeight = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerFaceWidth")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerFaceWidth = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerNumSections")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerNumSections = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerSectionLength")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerSectionLength = Me.settings.USUnits.Length.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerDiagonalSpacing")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagonalSpacing = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerDiagonalSpacingEx")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagonalSpacingEx = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerBraceType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerBraceType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerFaceBevel")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerFaceBevel = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerTopGirtOffset")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerTopGirtOffset = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerBotGirtOffset")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerBotGirtOffset = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerHasKBraceEndPanels")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerHasKBraceEndPanels = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerHasHorizontals")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerHasHorizontals = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLegType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerLegType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLegSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerLegSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLegGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerLegGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLegMatlGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerLegMatlGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerDiagonalGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagonalGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerDiagonalMatlGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagonalMatlGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerInnerBracingGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerBracingGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerInnerBracingMatlGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerBracingMatlGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerTopGirtGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerTopGirtGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerTopGirtMatlGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerTopGirtMatlGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerBotGirtGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerBotGirtGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerBotGirtMatlGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerBotGirtMatlGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerInnerGirtGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerGirtGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerInnerGirtMatlGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerGirtMatlGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLongHorizontalGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerLongHorizontalGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLongHorizontalMatlGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerLongHorizontalMatlGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerShortHorizontalGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerShortHorizontalGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerShortHorizontalMatlGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerShortHorizontalMatlGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerDiagonalType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagonalType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerDiagonalSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagonalSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerInnerBracingType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerBracingType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerInnerBracingSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerBracingSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerTopGirtType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerTopGirtType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerTopGirtSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerTopGirtSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerBotGirtType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerBotGirtType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerBotGirtSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerBotGirtSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerNumInnerGirts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerNumInnerGirts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerInnerGirtType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerGirtType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerInnerGirtSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerGirtSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLongHorizontalType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerLongHorizontalType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLongHorizontalSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerLongHorizontalSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerShortHorizontalType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerShortHorizontalType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerShortHorizontalSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerShortHorizontalSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantMatlGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantMatlGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantDiagType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantDiagType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubDiagonalType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubDiagonalType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubHorizontalType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubHorizontalType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantVerticalType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantVerticalType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipDiagonalType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipDiagonalType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHorizontalSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHorizontalSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHorizontalSize2")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHorizontalSize2 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHorizontalSize3")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHorizontalSize3 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHorizontalSize4")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHorizontalSize4 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantDiagonalSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantDiagonalSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantDiagonalSize2")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantDiagonalSize2 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantDiagonalSize3")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantDiagonalSize3 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantDiagonalSize4")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantDiagonalSize4 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubHorizontalSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubHorizontalSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubDiagonalSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubDiagonalSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerSubDiagLocation")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerSubDiagLocation = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantVerticalSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantVerticalSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipSize2")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipSize2 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipSize3")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipSize3 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipSize4")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipSize4 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipDiagonalSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipDiagonalSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipDiagonalSize2")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipDiagonalSize2 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipDiagonalSize3")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipDiagonalSize3 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipDiagonalSize4")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipDiagonalSize4 = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerSWMult")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerSWMult = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerWPMult")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerWPMult = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerAutoCalcKSingleAngle")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerAutoCalcKSingleAngle = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerAutoCalcKSolidRound")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerAutoCalcKSolidRound = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerAfGusset")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerAfGusset = Me.settings.USUnits.Length.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerTfGusset")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerTfGusset = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerGussetBoltEdgeDistance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerGussetBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerGussetGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerGussetGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerGussetMatlGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerGussetMatlGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerAfMult")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerAfMult = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerArMult")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerArMult = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerFlatIPAPole")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerFlatIPAPole = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRoundIPAPole")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRoundIPAPole = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerFlatIPALeg")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerFlatIPALeg = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRoundIPALeg")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRoundIPALeg = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerFlatIPAHorizontal")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerFlatIPAHorizontal = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRoundIPAHorizontal")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRoundIPAHorizontal = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerFlatIPADiagonal")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerFlatIPADiagonal = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRoundIPADiagonal")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRoundIPADiagonal = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerCSA_S37_SpeedUpFactor")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerCSA_S37_SpeedUpFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerKLegs")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKLegs = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerKXBracedDiags")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKXBracedDiags = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerKKBracedDiags")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKKBracedDiags = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerKZBracedDiags")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKZBracedDiags = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerKHorzs")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKHorzs = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerKSecHorzs")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKSecHorzs = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerKGirts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKGirts = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerKInners")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKInners = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerKXBracedDiagsY")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKXBracedDiagsY = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerKKBracedDiagsY")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKKBracedDiagsY = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerKZBracedDiagsY")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKZBracedDiagsY = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerKHorzsY")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKHorzsY = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerKSecHorzsY")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKSecHorzsY = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerKGirtsY")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKGirtsY = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerKInnersY")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKInnersY = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerKRedHorz")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKRedHorz = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerKRedDiag")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKRedDiag = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerKRedSubDiag")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKRedSubDiag = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerKRedSubHorz")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKRedSubHorz = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerKRedVert")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKRedVert = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerKRedHip")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKRedHip = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerKRedHipDiag")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKRedHipDiag = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerKTLX")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKTLX = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerKTLZ")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKTLZ = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerKTLLeg")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKTLLeg = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerInnerKTLX")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerKTLX = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerInnerKTLZ")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerKTLZ = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerInnerKTLLeg")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerKTLLeg = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerStitchBoltLocationHoriz")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerStitchBoltLocationHoriz = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerStitchBoltLocationDiag")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerStitchBoltLocationDiag = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerStitchBoltLocationRed")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerStitchBoltLocationRed = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerStitchSpacing")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerStitchSpacing = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerStitchSpacingDiag")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerStitchSpacingDiag = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerStitchSpacingHorz")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerStitchSpacingHorz = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerStitchSpacingRed")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerStitchSpacingRed = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLegNetWidthDeduct")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerLegNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLegUFactor")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerLegUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerDiagonalNetWidthDeduct")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagonalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerTopGirtNetWidthDeduct")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerTopGirtNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerBotGirtNetWidthDeduct")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerBotGirtNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerInnerGirtNetWidthDeduct")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerGirtNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerHorizontalNetWidthDeduct")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerHorizontalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerShortHorizontalNetWidthDeduct")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerShortHorizontalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerDiagonalUFactor")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagonalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerTopGirtUFactor")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerTopGirtUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerBotGirtUFactor")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerBotGirtUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerInnerGirtUFactor")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerGirtUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerHorizontalUFactor")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerHorizontalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerShortHorizontalUFactor")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerShortHorizontalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLegConnType")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerLegConnType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLegNumBolts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerLegNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerDiagonalNumBolts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagonalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerTopGirtNumBolts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerTopGirtNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerBotGirtNumBolts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerBotGirtNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerInnerGirtNumBolts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerGirtNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerHorizontalNumBolts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerHorizontalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerShortHorizontalNumBolts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerShortHorizontalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLegBoltGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerLegBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLegBoltSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerLegBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerDiagonalBoltGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagonalBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerDiagonalBoltSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagonalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerTopGirtBoltGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerTopGirtBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerTopGirtBoltSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerTopGirtBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerBotGirtBoltGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerBotGirtBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerBotGirtBoltSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerBotGirtBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerInnerGirtBoltGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerGirtBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerInnerGirtBoltSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerGirtBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerHorizontalBoltGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerHorizontalBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerHorizontalBoltSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerHorizontalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerShortHorizontalBoltGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerShortHorizontalBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerShortHorizontalBoltSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerShortHorizontalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLegBoltEdgeDistance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerLegBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerDiagonalBoltEdgeDistance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagonalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerTopGirtBoltEdgeDistance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerTopGirtBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerBotGirtBoltEdgeDistance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerBotGirtBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerInnerGirtBoltEdgeDistance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerGirtBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerHorizontalBoltEdgeDistance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerHorizontalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerShortHorizontalBoltEdgeDistance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerShortHorizontalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerDiagonalGageG1Distance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagonalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerTopGirtGageG1Distance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerTopGirtGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerBotGirtGageG1Distance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerBotGirtGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerInnerGirtGageG1Distance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerInnerGirtGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerHorizontalGageG1Distance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerHorizontalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerShortHorizontalGageG1Distance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerShortHorizontalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHorizontalBoltGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHorizontalBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHorizontalBoltSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHorizontalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHorizontalNumBolts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHorizontalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHorizontalBoltEdgeDistance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHorizontalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHorizontalGageG1Distance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHorizontalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHorizontalNetWidthDeduct")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHorizontalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHorizontalUFactor")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHorizontalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantDiagonalBoltGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantDiagonalBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantDiagonalBoltSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantDiagonalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantDiagonalNumBolts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantDiagonalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantDiagonalBoltEdgeDistance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantDiagonalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantDiagonalGageG1Distance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantDiagonalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantDiagonalNetWidthDeduct")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantDiagonalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantDiagonalUFactor")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantDiagonalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubDiagonalBoltGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubDiagonalBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubDiagonalBoltSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubDiagonalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubDiagonalNumBolts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubDiagonalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubDiagonalBoltEdgeDistance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubDiagonalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubDiagonalGageG1Distance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubDiagonalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubDiagonalNetWidthDeduct")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubDiagonalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubDiagonalUFactor")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubDiagonalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubHorizontalBoltGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubHorizontalBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubHorizontalBoltSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubHorizontalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubHorizontalNumBolts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubHorizontalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubHorizontalBoltEdgeDistance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubHorizontalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubHorizontalGageG1Distance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubHorizontalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubHorizontalNetWidthDeduct")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubHorizontalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantSubHorizontalUFactor")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantSubHorizontalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantVerticalBoltGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantVerticalBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantVerticalBoltSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantVerticalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantVerticalNumBolts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantVerticalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantVerticalBoltEdgeDistance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantVerticalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantVerticalGageG1Distance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantVerticalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantVerticalNetWidthDeduct")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantVerticalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantVerticalUFactor")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantVerticalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipBoltGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipBoltSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipNumBolts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipBoltEdgeDistance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipGageG1Distance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipNetWidthDeduct")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipUFactor")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipDiagonalBoltGrade")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipDiagonalBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipDiagonalBoltSize")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipDiagonalBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipDiagonalNumBolts")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipDiagonalNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipDiagonalBoltEdgeDistance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipDiagonalBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipDiagonalGageG1Distance")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipDiagonalGageG1Distance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipDiagonalNetWidthDeduct")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipDiagonalNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerRedundantHipDiagonalUFactor")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerRedundantHipDiagonalUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerDiagonalOutOfPlaneRestraint")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagonalOutOfPlaneRestraint = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerTopGirtOutOfPlaneRestraint")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerTopGirtOutOfPlaneRestraint = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerBottomGirtOutOfPlaneRestraint")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerBottomGirtOutOfPlaneRestraint = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerMidGirtOutOfPlaneRestraint")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerMidGirtOutOfPlaneRestraint = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerHorizontalOutOfPlaneRestraint")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerHorizontalOutOfPlaneRestraint = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerSecondaryHorizontalOutOfPlaneRestraint")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerSecondaryHorizontalOutOfPlaneRestraint = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerUniqueFlag")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerUniqueFlag = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerDiagOffsetNEY")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagOffsetNEY = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerDiagOffsetNEX")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagOffsetNEX = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerDiagOffsetPEY")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagOffsetPEY = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerDiagOffsetPEX")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerDiagOffsetPEX = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerKbraceOffsetNEY")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKbraceOffsetNEY = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerKbraceOffsetNEX")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKbraceOffsetNEX = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerKbraceOffsetPEY")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKbraceOffsetPEY = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerKbraceOffsetPEX")
                            Try
                                Me.geometry.baseStructure(recIndex).TowerKbraceOffsetPEX = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                    End Select
                Case caseFilter = "Guy"
                    ''''Guy Rec''''
                    Select Case True
                        Case tnxVar.Equals("GuyRec")
                            Try
                                recIndex = CInt(tnxValue) - 1
                                Me.geometry.guyWires.Add(New tnxGuyRecord())
                                Me.geometry.guyWires(recIndex).GuyRec = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyHeight")
                            Try
                                Me.geometry.guyWires(recIndex).GuyHeight = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyAutoCalcKSingleAngle")
                            Try
                                Me.geometry.guyWires(recIndex).GuyAutoCalcKSingleAngle = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyAutoCalcKSolidRound")
                            Try
                                Me.geometry.guyWires(recIndex).GuyAutoCalcKSolidRound = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyMount")
                            Try
                                Me.geometry.guyWires(recIndex).GuyMount = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TorqueArmStyle")
                            Try
                                Me.geometry.guyWires(recIndex).TorqueArmStyle = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyRadius")
                            Try
                                Me.geometry.guyWires(recIndex).GuyRadius = Me.settings.USUnits.Length.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyRadius120")
                            Try
                                Me.geometry.guyWires(recIndex).GuyRadius120 = Me.settings.USUnits.Length.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyRadius240")
                            Try
                                Me.geometry.guyWires(recIndex).GuyRadius240 = Me.settings.USUnits.Length.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyRadius360")
                            Try
                                Me.geometry.guyWires(recIndex).GuyRadius360 = Me.settings.USUnits.Length.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TorqueArmRadius")
                            Try
                                Me.geometry.guyWires(recIndex).TorqueArmRadius = Me.settings.USUnits.Length.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TorqueArmLegAngle")
                            Try
                                Me.geometry.guyWires(recIndex).TorqueArmLegAngle = Me.settings.USUnits.Rotation.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Azimuth0Adjustment")
                            Try
                                Me.geometry.guyWires(recIndex).Azimuth0Adjustment = Me.settings.USUnits.Rotation.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Azimuth120Adjustment")
                            Try
                                Me.geometry.guyWires(recIndex).Azimuth120Adjustment = Me.settings.USUnits.Rotation.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Azimuth240Adjustment")
                            Try
                                Me.geometry.guyWires(recIndex).Azimuth240Adjustment = Me.settings.USUnits.Rotation.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Azimuth360Adjustment")
                            Try
                                Me.geometry.guyWires(recIndex).Azimuth360Adjustment = Me.settings.USUnits.Rotation.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Anchor0Elevation")
                            Try
                                Me.geometry.guyWires(recIndex).Anchor0Elevation = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Anchor120Elevation")
                            Try
                                Me.geometry.guyWires(recIndex).Anchor120Elevation = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Anchor240Elevation")
                            Try
                                Me.geometry.guyWires(recIndex).Anchor240Elevation = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Anchor360Elevation")
                            Try
                                Me.geometry.guyWires(recIndex).Anchor360Elevation = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuySize")
                            Try
                                Me.geometry.guyWires(recIndex).GuySize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Guy120Size")
                            Try
                                Me.geometry.guyWires(recIndex).Guy120Size = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Guy240Size")
                            Try
                                Me.geometry.guyWires(recIndex).Guy240Size = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("Guy360Size")
                            Try
                                Me.geometry.guyWires(recIndex).Guy360Size = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyGrade")
                            Try
                                Me.geometry.guyWires(recIndex).GuyGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TorqueArmSize")
                            Try
                                Me.geometry.guyWires(recIndex).TorqueArmSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TorqueArmSizeBot")
                            Try
                                Me.geometry.guyWires(recIndex).TorqueArmSizeBot = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TorqueArmType")
                            Try
                                Me.geometry.guyWires(recIndex).TorqueArmType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TorqueArmGrade")
                            Try
                                Me.geometry.guyWires(recIndex).TorqueArmGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TorqueArmMatlGrade")
                            Try
                                Me.geometry.guyWires(recIndex).TorqueArmMatlGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TorqueArmKFactor")
                            Try
                                Me.geometry.guyWires(recIndex).TorqueArmKFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TorqueArmKFactorY")
                            Try
                                Me.geometry.guyWires(recIndex).TorqueArmKFactorY = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyPullOffKFactorX")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPullOffKFactorX = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyPullOffKFactorY")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPullOffKFactorY = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyDiagKFactorX")
                            Try
                                Me.geometry.guyWires(recIndex).GuyDiagKFactorX = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyDiagKFactorY")
                            Try
                                Me.geometry.guyWires(recIndex).GuyDiagKFactorY = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyAutoCalc")
                            Try
                                Me.geometry.guyWires(recIndex).GuyAutoCalc = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyAllGuysSame")
                            Try
                                Me.geometry.guyWires(recIndex).GuyAllGuysSame = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyAllGuysAnchorSame")
                            Try
                                Me.geometry.guyWires(recIndex).GuyAllGuysAnchorSame = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyIsStrapping")
                            Try
                                Me.geometry.guyWires(recIndex).GuyIsStrapping = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyPullOffSize")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPullOffSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyPullOffSizeBot")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPullOffSizeBot = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyPullOffType")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPullOffType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyPullOffGrade")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPullOffGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyPullOffMatlGrade")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPullOffMatlGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyUpperDiagSize")
                            Try
                                Me.geometry.guyWires(recIndex).GuyUpperDiagSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyLowerDiagSize")
                            Try
                                Me.geometry.guyWires(recIndex).GuyLowerDiagSize = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyDiagType")
                            Try
                                Me.geometry.guyWires(recIndex).GuyDiagType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyDiagGrade")
                            Try
                                Me.geometry.guyWires(recIndex).GuyDiagGrade = Me.settings.USUnits.Strength.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyDiagMatlGrade")
                            Try
                                Me.geometry.guyWires(recIndex).GuyDiagMatlGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyDiagNetWidthDeduct")
                            Try
                                Me.geometry.guyWires(recIndex).GuyDiagNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyDiagUFactor")
                            Try
                                Me.geometry.guyWires(recIndex).GuyDiagUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyDiagNumBolts")
                            Try
                                Me.geometry.guyWires(recIndex).GuyDiagNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyDiagonalOutOfPlaneRestraint")
                            Try
                                Me.geometry.guyWires(recIndex).GuyDiagonalOutOfPlaneRestraint = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyDiagBoltGrade")
                            Try
                                Me.geometry.guyWires(recIndex).GuyDiagBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyDiagBoltSize")
                            Try
                                Me.geometry.guyWires(recIndex).GuyDiagBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyDiagBoltEdgeDistance")
                            Try
                                Me.geometry.guyWires(recIndex).GuyDiagBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyDiagBoltGageDistance")
                            Try
                                Me.geometry.guyWires(recIndex).GuyDiagBoltGageDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyPullOffNetWidthDeduct")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPullOffNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyPullOffUFactor")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPullOffUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyPullOffNumBolts")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPullOffNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyPullOffOutOfPlaneRestraint")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPullOffOutOfPlaneRestraint = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyPullOffBoltGrade")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPullOffBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyPullOffBoltSize")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPullOffBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyPullOffBoltEdgeDistance")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPullOffBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyPullOffBoltGageDistance")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPullOffBoltGageDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyTorqueArmNetWidthDeduct")
                            Try
                                Me.geometry.guyWires(recIndex).GuyTorqueArmNetWidthDeduct = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyTorqueArmUFactor")
                            Try
                                Me.geometry.guyWires(recIndex).GuyTorqueArmUFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyTorqueArmNumBolts")
                            Try
                                Me.geometry.guyWires(recIndex).GuyTorqueArmNumBolts = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyTorqueArmOutOfPlaneRestraint")
                            Try
                                Me.geometry.guyWires(recIndex).GuyTorqueArmOutOfPlaneRestraint = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyTorqueArmBoltGrade")
                            Try
                                Me.geometry.guyWires(recIndex).GuyTorqueArmBoltGrade = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyTorqueArmBoltSize")
                            Try
                                Me.geometry.guyWires(recIndex).GuyTorqueArmBoltSize = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyTorqueArmBoltEdgeDistance")
                            Try
                                Me.geometry.guyWires(recIndex).GuyTorqueArmBoltEdgeDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyTorqueArmBoltGageDistance")
                            Try
                                Me.geometry.guyWires(recIndex).GuyTorqueArmBoltGageDistance = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyPerCentTension")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPerCentTension = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyPerCentTension120")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPerCentTension120 = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyPerCentTension240")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPerCentTension240 = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyPerCentTension360")
                            Try
                                Me.geometry.guyWires(recIndex).GuyPerCentTension360 = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyEffFactor")
                            Try
                                Me.geometry.guyWires(recIndex).GuyEffFactor = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyEffFactor120")
                            Try
                                Me.geometry.guyWires(recIndex).GuyEffFactor120 = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyEffFactor240")
                            Try
                                Me.geometry.guyWires(recIndex).GuyEffFactor240 = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyEffFactor360")
                            Try
                                Me.geometry.guyWires(recIndex).GuyEffFactor360 = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyNumInsulators")
                            Try
                                Me.geometry.guyWires(recIndex).GuyNumInsulators = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyInsulatorLength")
                            Try
                                Me.geometry.guyWires(recIndex).GuyInsulatorLength = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyInsulatorDia")
                            Try
                                Me.geometry.guyWires(recIndex).GuyInsulatorDia = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("GuyInsulatorWt")
                            Try
                                Me.geometry.guyWires(recIndex).GuyInsulatorWt = Me.settings.USUnits.Force.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                    End Select
                Case caseFilter = "FeedLine"
                    ''''Feed Lines''''
                    Select Case True
                        Case tnxVar.Equals("FeedLineRec")
                            Try
                                recIndex = CInt(tnxValue) - 1
                                Me.feedLines.Add(New tnxFeedLine())
                                Me.feedLines(recIndex).FeedLineRec = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineEnabled")
                            Try
                                Me.feedLines(recIndex).FeedLineEnabled = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineDatabase")
                            Try
                                Me.feedLines(recIndex).FeedLineDatabase = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineDescription")
                            Try
                                Me.feedLines(recIndex).FeedLineDescription = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineClassificationCategory")
                            Try
                                Me.feedLines(recIndex).FeedLineClassificationCategory = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineNote")
                            Try
                                Me.feedLines(recIndex).FeedLineNote = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineNum")
                            Try
                                Me.feedLines(recIndex).FeedLineNum = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineUseShielding")
                            Try
                                Me.feedLines(recIndex).FeedLineUseShielding = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("ExcludeFeedLineFromTorque")
                            Try
                                Me.feedLines(recIndex).ExcludeFeedLineFromTorque = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineNumPerRow")
                            Try
                                Me.feedLines(recIndex).FeedLineNumPerRow = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineFace")
                            Try
                                Me.feedLines(recIndex).FeedLineFace = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineComponentType")
                            Try
                                Me.feedLines(recIndex).FeedLineComponentType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineGroupTreatmentType")
                            Try
                                Me.feedLines(recIndex).FeedLineGroupTreatmentType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineRoundClusterDia")
                            Try
                                Me.feedLines(recIndex).FeedLineRoundClusterDia = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineWidth")
                            Try
                                Me.feedLines(recIndex).FeedLineWidth = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLinePerimeter")
                            Try
                                Me.feedLines(recIndex).FeedLinePerimeter = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FlatAttachmentEffectiveWidthRatio")
                            Try
                                Me.feedLines(recIndex).FlatAttachmentEffectiveWidthRatio = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("AutoCalcFlatAttachmentEffectiveWidthRatio")
                            Try
                                Me.feedLines(recIndex).AutoCalcFlatAttachmentEffectiveWidthRatio = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineShieldingFactorKaNoIce")
                            Try
                                Me.feedLines(recIndex).FeedLineShieldingFactorKaNoIce = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineShieldingFactorKaIce")
                            Try
                                Me.feedLines(recIndex).FeedLineShieldingFactorKaIce = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineAutoCalcKa")
                            Try
                                Me.feedLines(recIndex).FeedLineAutoCalcKa = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineCaAaNoIce")
                            'The units of this input are dependant on the TNX length unit setting but they are an area (in^2 or ft^2)
                            Try
                                Me.feedLines(recIndex).FeedLineCaAaNoIce = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineCaAaIce")
                            'The units of this input are dependant on the TNX length unit setting but they are an area (in^2 or ft^2)
                            Try
                                Me.feedLines(recIndex).FeedLineCaAaIce = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineCaAaIce_1")
                            'The units of this input are dependant on the TNX length unit setting but they are an area (in^2 or ft^2)
                            Try
                                Me.feedLines(recIndex).FeedLineCaAaIce_1 = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineCaAaIce_2")
                            'The units of this input are dependant on the TNX length unit setting but they are an area (in^2 or ft^2)
                            Try
                                Me.feedLines(recIndex).FeedLineCaAaIce_2 = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineCaAaIce_4")
                            'The units of this input are dependant on the TNX length unit setting but they are an area (in^2 or ft^2)
                            Try
                                Me.feedLines(recIndex).FeedLineCaAaIce_4 = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineWtNoIce")
                            Try
                                Me.feedLines(recIndex).FeedLineWtNoIce = Me.settings.USUnits.Load.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineWtIce")
                            Try
                                Me.feedLines(recIndex).FeedLineWtIce = Me.settings.USUnits.Load.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineWtIce_1")
                            Try
                                Me.feedLines(recIndex).FeedLineWtIce_1 = Me.settings.USUnits.Load.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineWtIce_2")
                            Try
                                Me.feedLines(recIndex).FeedLineWtIce_2 = Me.settings.USUnits.Load.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineWtIce_4")
                            Try
                                Me.feedLines(recIndex).FeedLineWtIce_4 = Me.settings.USUnits.Load.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineFaceOffset")
                            Try
                                Me.feedLines(recIndex).FeedLineFaceOffset = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineOffsetFrac")
                            Try
                                Me.feedLines(recIndex).FeedLineOffsetFrac = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLinePerimeterOffsetStartFrac")
                            Try
                                Me.feedLines(recIndex).FeedLinePerimeterOffsetStartFrac = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLinePerimeterOffsetEndFrac")
                            Try
                                Me.feedLines(recIndex).FeedLinePerimeterOffsetEndFrac = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineStartHt")
                            Try
                                Me.feedLines(recIndex).FeedLineStartHt = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineEndHt")
                            Try
                                Me.feedLines(recIndex).FeedLineEndHt = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineClearSpacing")
                            Try
                                Me.feedLines(recIndex).FeedLineClearSpacing = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("FeedLineRowClearSpacing")
                            Try
                                Me.feedLines(recIndex).FeedLineRowClearSpacing = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                    End Select
                Case caseFilter = "Discrete"
                    ''''Discrete''''
                    Select Case True
                        Case tnxVar.Equals("TowerLoadRec")
                            Try
                                recIndex = CInt(tnxValue) - 1
                                Me.discreteLoads.Add(New tnxDiscreteLoad())
                                Me.discreteLoads(recIndex).TowerLoadRec = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLoadEnabled")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadEnabled = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLoadDatabase")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadDatabase = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLoadDescription")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadDescription = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLoadType")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLoadClassificationCategory")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadClassificationCategory = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLoadNote")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadNote = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLoadNum")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadNum = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLoadFace")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadFace = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerOffsetType")
                            Try
                                Me.discreteLoads(recIndex).TowerOffsetType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerOffsetDist")
                            Try
                                Me.discreteLoads(recIndex).TowerOffsetDist = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerVertOffset")
                            Try
                                Me.discreteLoads(recIndex).TowerVertOffset = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLateralOffset")
                            Try
                                Me.discreteLoads(recIndex).TowerLateralOffset = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerAzimuthAdjustment")
                            Try
                                Me.discreteLoads(recIndex).TowerAzimuthAdjustment = Me.settings.USUnits.Rotation.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerAppurtSymbol")
                            Try
                                Me.discreteLoads(recIndex).TowerAppurtSymbol = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLoadShieldingFactorKaNoIce")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadShieldingFactorKaNoIce = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLoadShieldingFactorKaIce")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadShieldingFactorKaIce = CDbl(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLoadAutoCalcKa")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadAutoCalcKa = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLoadCaAaNoIce")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadCaAaNoIce = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLoadCaAaIce")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadCaAaIce = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLoadCaAaIce_1")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadCaAaIce_1 = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLoadCaAaIce_2")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadCaAaIce_2 = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLoadCaAaIce_4")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadCaAaIce_4 = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLoadCaAaNoIce_Side")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadCaAaNoIce_Side = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLoadCaAaIce_Side")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadCaAaIce_Side = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLoadCaAaIce_Side_1")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadCaAaIce_Side_1 = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLoadCaAaIce_Side_2")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadCaAaIce_Side_2 = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLoadCaAaIce_Side_4")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadCaAaIce_Side_4 = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLoadWtNoIce")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadWtNoIce = Me.settings.USUnits.Force.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLoadWtIce")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadWtIce = Me.settings.USUnits.Force.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLoadWtIce_1")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadWtIce_1 = Me.settings.USUnits.Force.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLoadWtIce_2")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadWtIce_2 = Me.settings.USUnits.Force.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLoadWtIce_4")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadWtIce_4 = Me.settings.USUnits.Force.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLoadStartHt")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadStartHt = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("TowerLoadEndHt")
                            Try
                                Me.discreteLoads(recIndex).TowerLoadEndHt = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                    End Select
                Case caseFilter = "Dish"
                    ''''Dishes''''
                    Select Case True
                        Case tnxVar.Equals("DishRec")
                            Try
                                recIndex = CInt(tnxValue) - 1
                                Me.dishes.Add(New tnxDish())
                                Me.dishes(recIndex).DishRec = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DishEnabled")
                            Try
                                Me.dishes(recIndex).DishEnabled = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DishDatabase")
                            Try
                                Me.dishes(recIndex).DishDatabase = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DishDescription")
                            Try
                                Me.dishes(recIndex).DishDescription = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DishClassificationCategory")
                            Try
                                Me.dishes(recIndex).DishClassificationCategory = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DishNote")
                            Try
                                Me.dishes(recIndex).DishNote = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DishNum")
                            Try
                                Me.dishes(recIndex).DishNum = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DishFace")
                            Try
                                Me.dishes(recIndex).DishFace = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DishType")
                            Try
                                Me.dishes(recIndex).DishType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DishOffsetType")
                            Try
                                Me.dishes(recIndex).DishOffsetType = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DishVertOffset")
                            Try
                                Me.dishes(recIndex).DishVertOffset = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DishLateralOffset")
                            Try
                                Me.dishes(recIndex).DishLateralOffset = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DishOffsetDist")
                            Try
                                Me.dishes(recIndex).DishOffsetDist = Me.settings.USUnits.Properties.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DishArea")
                            Try
                                Me.dishes(recIndex).DishArea = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DishAreaIce")
                            Try
                                Me.dishes(recIndex).DishAreaIce = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DishAreaIce_1")
                            Try
                                Me.dishes(recIndex).DishAreaIce_1 = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DishAreaIce_2")
                            Try
                                Me.dishes(recIndex).DishAreaIce_2 = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DishAreaIce_4")
                            Try
                                Me.dishes(recIndex).DishAreaIce_4 = Me.settings.USUnits.Length.convertAreaToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DishDiameter")
                            Try
                                Me.dishes(recIndex).DishDiameter = Me.settings.USUnits.Length.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DishWtNoIce")
                            Try
                                Me.dishes(recIndex).DishWtNoIce = Me.settings.USUnits.Force.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DishWtIce")
                            Try
                                Me.dishes(recIndex).DishWtIce = Me.settings.USUnits.Force.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DishWtIce_1")
                            Try
                                Me.dishes(recIndex).DishWtIce_1 = Me.settings.USUnits.Force.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DishWtIce_2")
                            Try
                                Me.dishes(recIndex).DishWtIce_2 = Me.settings.USUnits.Force.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DishWtIce_4")
                            Try
                                Me.dishes(recIndex).DishWtIce_4 = Me.settings.USUnits.Force.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DishStartHt")
                            Try
                                Me.dishes(recIndex).DishStartHt = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DishAzimuthAdjustment")
                            Try
                                Me.dishes(recIndex).DishAzimuthAdjustment = Me.settings.USUnits.Rotation.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("DishBeamWidth")
                            Try
                                Me.dishes(recIndex).DishBeamWidth = Me.settings.USUnits.Rotation.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                    End Select
                Case caseFilter = "UserForce"
                    ''''UserForces''''
                    Select Case True
                        Case tnxVar.Equals("UserForceRec")
                            Try
                                recIndex = CInt(tnxValue) - 1
                                Me.userForces.Add(New tnxUserForce())
                                Me.userForces(recIndex).UserForceRec = CInt(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UserForceEnabled")
                            Try
                                Me.userForces(recIndex).UserForceEnabled = trueFalseYesNo(tnxValue)
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UserForceDescription")
                            Try
                                Me.userForces(recIndex).UserForceDescription = tnxValue
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UserForceStartHt")
                            Try
                                Me.userForces(recIndex).UserForceStartHt = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UserForceOffset")
                            Try
                                Me.userForces(recIndex).UserForceOffset = Me.settings.USUnits.Coordinate.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UserForceAzimuth")
                            Try
                                Me.userForces(recIndex).UserForceAzimuth = Me.settings.USUnits.Rotation.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UserForceFxNoIce")
                            Try
                                Me.userForces(recIndex).UserForceFxNoIce = Me.settings.USUnits.Force.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UserForceFzNoIce")
                            Try
                                Me.userForces(recIndex).UserForceFzNoIce = Me.settings.USUnits.Force.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UserForceAxialNoIce")
                            Try
                                Me.userForces(recIndex).UserForceAxialNoIce = Me.settings.USUnits.Force.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UserForceShearNoIce")
                            Try
                                Me.userForces(recIndex).UserForceShearNoIce = Me.settings.USUnits.Force.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UserForceCaAcNoIce")
                            Try
                                Me.userForces(recIndex).UserForceCaAcNoIce = Me.settings.USUnits.Length.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UserForceFxIce")
                            Try
                                Me.userForces(recIndex).UserForceFxIce = Me.settings.USUnits.Force.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UserForceFzIce")
                            Try
                                Me.userForces(recIndex).UserForceFzIce = Me.settings.USUnits.Force.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UserForceAxialIce")
                            Try
                                Me.userForces(recIndex).UserForceAxialIce = Me.settings.USUnits.Force.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UserForceShearIce")
                            Try
                                Me.userForces(recIndex).UserForceShearIce = Me.settings.USUnits.Force.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UserForceCaAcIce")
                            Try
                                Me.userForces(recIndex).UserForceCaAcIce = Me.settings.USUnits.Length.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UserForceFxService")
                            Try
                                Me.userForces(recIndex).UserForceFxService = Me.settings.USUnits.Force.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UserForceFzService")
                            Try
                                Me.userForces(recIndex).UserForceFzService = Me.settings.USUnits.Force.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UserForceAxialService")
                            Try
                                Me.userForces(recIndex).UserForceAxialService = Me.settings.USUnits.Force.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UserForceShearService")
                            Try
                                Me.userForces(recIndex).UserForceShearService = Me.settings.USUnits.Force.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UserForceCaAcService")
                            Try
                                Me.userForces(recIndex).UserForceCaAcService = Me.settings.USUnits.Length.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UserForceEhx")
                            Try
                                Me.userForces(recIndex).UserForceEhx = Me.settings.USUnits.Force.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UserForceEhz")
                            Try
                                Me.userForces(recIndex).UserForceEhz = Me.settings.USUnits.Force.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UserForceEv")
                            Try
                                Me.userForces(recIndex).UserForceEv = Me.settings.USUnits.Force.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                        Case tnxVar.Equals("UserForceEh")
                            Try
                                Me.userForces(recIndex).UserForceEh = Me.settings.USUnits.Force.convertToEDSDefaultUnits(CDbl(tnxValue))
                            Catch ex As Exception
                            End Try
                    End Select
            End Select
        Next

    End Sub

    Public Sub GenerateERI(FilePath As String)

        Dim newERIList As New List(Of String)
        Dim i As Integer

        For Each line In Me.otherLines

            Select Case True
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
                Case line(0).Equals("[Structure]")
                    newERIList.Add(line(0))
                    'User Info Settings
                    newERIList.Add("ViewerUserName=" & Me.settings.userInfo.ViewerUserName)
                    newERIList.Add("ViewerCompanyName=" & Me.settings.userInfo.ViewerCompanyName)
                    newERIList.Add("ViewerStreetAddress=" & Me.settings.userInfo.ViewerStreetAddress)
                    newERIList.Add("ViewerCityState=" & Me.settings.userInfo.ViewerCityState)
                    newERIList.Add("ViewerPhone=" & Me.settings.userInfo.ViewerPhone)
                    newERIList.Add("ViewerFAX=" & Me.settings.userInfo.ViewerFAX)
                    newERIList.Add("ViewerLogo=" & Me.settings.userInfo.ViewerLogo)
                    newERIList.Add("ViewerCompanyBitmap=" & Me.settings.userInfo.ViewerCompanyBitmap)
                    'Solution Settings
                    newERIList.Add("SolutionUsePDelta=" & trueFalseYesNo(Me.solutionSettings.SolutionUsePDelta))
                    newERIList.Add("SolutionMinStiffness=" & Me.solutionSettings.SolutionMinStiffness)
                    newERIList.Add("SolutionMaxStiffness=" & Me.solutionSettings.SolutionMaxStiffness)
                    newERIList.Add("SolutionMaxCycles=" & Me.solutionSettings.SolutionMaxCycles)
                    newERIList.Add("SolutionPower=" & Me.solutionSettings.SolutionPower)
                    newERIList.Add("SolutionTolerance=" & Me.solutionSettings.SolutionTolerance)
                    'MTO Settings
                    newERIList.Add("IncludeCapacityNote=" & trueFalseYesNo(Me.MTOSettings.IncludeCapacityNote))
                    newERIList.Add("IncludeAppurtGraphics=" & trueFalseYesNo(Me.MTOSettings.IncludeAppurtGraphics))
                    newERIList.Add("DisplayNotes=" & trueFalseYesNo(Me.MTOSettings.DisplayNotes))
                    newERIList.Add("DisplayReactions=" & trueFalseYesNo(Me.MTOSettings.DisplayReactions))
                    newERIList.Add("DisplaySchedule=" & trueFalseYesNo(Me.MTOSettings.DisplaySchedule))
                    newERIList.Add("DisplayAppurtenanceTable=" & trueFalseYesNo(Me.MTOSettings.DisplayAppurtenanceTable))
                    newERIList.Add("DisplayMaterialStrengthTable=" & trueFalseYesNo(Me.MTOSettings.DisplayMaterialStrengthTable))
                    For Each note In Me.MTOSettings.Notes
                        newERIList.Add("Notes=" & note)
                    Next
                    'Report Settings
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
                    'CCIReport
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
                    'Code - Design
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
                    newERIList.Add("TIA_222_H_AnnexS_Ratio=" & Me.code.design.TIA_222_H_AnnexS_Ratio)
                    newERIList.Add("PrintBitmaps=" & trueFalseYesNo(Me.code.design.PrintBitmaps))
                    'Code - Wind
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
                    'Code - Seismic
                    newERIList.Add("UseASCE7_10_Seismic_Lcomb=" & trueFalseYesNo(Me.code.seismic.UseASCE7_10_Seismic_Lcomb))
                    newERIList.Add("SeismicSiteClass=" & Me.code.seismic.SeismicSiteClass)
                    newERIList.Add("SeismicSs=" & Me.code.seismic.SeismicSs)
                    newERIList.Add("SeismicS1=" & Me.code.seismic.SeismicS1)
                    'Code - Ice
                    newERIList.Add("IceThickness=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.code.ice.IceThickness))
                    newERIList.Add("IceDensity=" & Me.settings.USUnits.Density.convertToERIUnits(Me.code.ice.IceDensity))
                    newERIList.Add("UseModified_TIA_222_IceParameters=" & trueFalseYesNo(Me.code.ice.UseModified_TIA_222_IceParameters))
                    newERIList.Add("TIA_222_IceThicknessMultiplier=" & Me.code.ice.TIA_222_IceThicknessMultiplier)
                    newERIList.Add("DoNotUse_TIA_222_IceEscalation=" & trueFalseYesNo(Me.code.ice.DoNotUse_TIA_222_IceEscalation))
                    newERIList.Add("UseIceEscalation=" & trueFalseYesNo(Me.code.ice.UseIceEscalation))
                    'Code - Thermal
                    newERIList.Add("TempDrop=" & Me.settings.USUnits.Temperature.convertToERIUnits(Me.code.thermal.TempDrop))
                    'Code - Miscl
                    newERIList.Add("GroutFc=" & Me.settings.USUnits.Strength.convertToERIUnits(Me.code.misclCode.GroutFc))
                    newERIList.Add("TowerBoltGrade=" & Me.code.misclCode.TowerBoltGrade)
                    newERIList.Add("TowerBoltMinEdgeDist=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.code.misclCode.TowerBoltMinEdgeDist))
                    'Options - General
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
                    'Options - Foundations
                    newERIList.Add("MastVert=" & Me.settings.USUnits.Force.convertToERIUnits(Me.options.foundationStiffness.MastVert))
                    newERIList.Add("MastHorz=" & Me.settings.USUnits.Force.convertToERIUnits(Me.options.foundationStiffness.MastHorz))
                    newERIList.Add("GuyVert=" & Me.settings.USUnits.Force.convertToERIUnits(Me.options.foundationStiffness.GuyVert))
                    newERIList.Add("GuyHorz=" & Me.settings.USUnits.Force.convertToERIUnits(Me.options.foundationStiffness.GuyHorz))
                    'Options - Poles
                    newERIList.Add("CheckVonMises=" & trueFalseYesNo(Me.options.cantileverPoles.CheckVonMises))
                    newERIList.Add("SocketTopMount=" & trueFalseYesNo(Me.options.cantileverPoles.SocketTopMount))
                    newERIList.Add("PrintMonopoleAtIncrements=" & trueFalseYesNo(Me.options.cantileverPoles.PrintMonopoleAtIncrements))
                    newERIList.Add("UseSubCriticalFlow=" & trueFalseYesNo(Me.options.cantileverPoles.UseSubCriticalFlow))
                    newERIList.Add("AssumePoleWithNoAttachments=" & trueFalseYesNo(Me.options.cantileverPoles.AssumePoleWithNoAttachments))
                    newERIList.Add("AssumePoleWithShroud=" & trueFalseYesNo(Me.options.cantileverPoles.AssumePoleWithShroud))
                    newERIList.Add("PoleCornerRadiusKnown=" & trueFalseYesNo(Me.options.cantileverPoles.PoleCornerRadiusKnown))
                    newERIList.Add("CantKFactor=" & Me.options.cantileverPoles.CantKFactor)
                    'Options - Girts
                    newERIList.Add("GirtOffset=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.options.defaultGirtOffsets.GirtOffset))
                    newERIList.Add("GirtOffsetLatticedPole=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.options.defaultGirtOffsets.GirtOffsetLatticedPole))
                    newERIList.Add("OffsetBotGirt=" & trueFalseYesNo(Me.options.defaultGirtOffsets.OffsetBotGirt))
                    'Options - Wind
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
                    'Options - Miscl
                    newERIList.Add("HogRodTakeup=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.options.misclOptions.HogRodTakeup))
                    newERIList.Add("RadiusSampleDist=" & Me.settings.USUnits.Length.convertToERIUnits(Me.options.misclOptions.RadiusSampleDist))
                    'General Geometry
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
                Case line(0).Equals("NumAntennaRecs")
                    newERIList.Add(line(0) & "=" & Me.geometry.upperStructure.Count)
                    'For Each upperSection In Me.geometry.upperStructure
                    For i = 0 To Me.geometry.upperStructure.Count - 1
                        newERIList.Add("AntennaRec=" & Me.geometry.upperStructure(i).AntennaRec)
                        newERIList.Add("AntennaBraceType=" & Me.geometry.upperStructure(i).AntennaBraceType)
                        newERIList.Add("AntennaHeight=" & Me.settings.USUnits.Coordinate.convertToERIUnits(Me.geometry.upperStructure(i).AntennaHeight))
                        newERIList.Add("AntennaDiagonalSpacing=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaDiagonalSpacing))
                        newERIList.Add("AntennaDiagonalSpacingEx=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaDiagonalSpacingEx))
                        newERIList.Add("AntennaNumSections=" & Me.geometry.upperStructure(i).AntennaNumSections)
                        newERIList.Add("AntennaNumSesctions=" & Me.geometry.upperStructure(i).AntennaNumSesctions)
                        newERIList.Add("AntennaSectionLength=" & Me.settings.USUnits.Length.convertToERIUnits(Me.geometry.upperStructure(i).AntennaSectionLength))
                        newERIList.Add("AntennaLegType=" & Me.geometry.upperStructure(i).AntennaLegType)
                        newERIList.Add("AntennaLegSize=" & Me.geometry.upperStructure(i).AntennaLegSize)
                        newERIList.Add("AntennaLegGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(Me.geometry.upperStructure(i).AntennaLegGrade))
                        newERIList.Add("AntennaLegMatlGrade=" & Me.geometry.upperStructure(i).AntennaLegMatlGrade)
                        newERIList.Add("AntennaDiagonalGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(Me.geometry.upperStructure(i).AntennaDiagonalGrade))
                        newERIList.Add("AntennaDiagonalMatlGrade=" & Me.geometry.upperStructure(i).AntennaDiagonalMatlGrade)
                        newERIList.Add("AntennaInnerBracingGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(Me.geometry.upperStructure(i).AntennaInnerBracingGrade))
                        newERIList.Add("AntennaInnerBracingMatlGrade=" & Me.geometry.upperStructure(i).AntennaInnerBracingMatlGrade)
                        newERIList.Add("AntennaTopGirtGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(Me.geometry.upperStructure(i).AntennaTopGirtGrade))
                        newERIList.Add("AntennaTopGirtMatlGrade=" & Me.geometry.upperStructure(i).AntennaTopGirtMatlGrade)
                        newERIList.Add("AntennaBotGirtGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(Me.geometry.upperStructure(i).AntennaBotGirtGrade))
                        newERIList.Add("AntennaBotGirtMatlGrade=" & Me.geometry.upperStructure(i).AntennaBotGirtMatlGrade)
                        newERIList.Add("AntennaInnerGirtGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(Me.geometry.upperStructure(i).AntennaInnerGirtGrade))
                        newERIList.Add("AntennaInnerGirtMatlGrade=" & Me.geometry.upperStructure(i).AntennaInnerGirtMatlGrade)
                        newERIList.Add("AntennaLongHorizontalGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(Me.geometry.upperStructure(i).AntennaLongHorizontalGrade))
                        newERIList.Add("AntennaLongHorizontalMatlGrade=" & Me.geometry.upperStructure(i).AntennaLongHorizontalMatlGrade)
                        newERIList.Add("AntennaShortHorizontalGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(Me.geometry.upperStructure(i).AntennaShortHorizontalGrade))
                        newERIList.Add("AntennaShortHorizontalMatlGrade=" & Me.geometry.upperStructure(i).AntennaShortHorizontalMatlGrade)
                        newERIList.Add("AntennaDiagonalType=" & Me.geometry.upperStructure(i).AntennaDiagonalType)
                        newERIList.Add("AntennaDiagonalSize=" & Me.geometry.upperStructure(i).AntennaDiagonalSize)
                        newERIList.Add("AntennaInnerBracingType=" & Me.geometry.upperStructure(i).AntennaInnerBracingType)
                        newERIList.Add("AntennaInnerBracingSize=" & Me.geometry.upperStructure(i).AntennaInnerBracingSize)
                        newERIList.Add("AntennaTopGirtType=" & Me.geometry.upperStructure(i).AntennaTopGirtType)
                        newERIList.Add("AntennaTopGirtSize=" & Me.geometry.upperStructure(i).AntennaTopGirtSize)
                        newERIList.Add("AntennaBotGirtType=" & Me.geometry.upperStructure(i).AntennaBotGirtType)
                        newERIList.Add("AntennaBotGirtSize=" & Me.geometry.upperStructure(i).AntennaBotGirtSize)
                        newERIList.Add("AntennaTopGirtOffset=" & Me.geometry.upperStructure(i).AntennaTopGirtOffset)
                        newERIList.Add("AntennaBotGirtOffset=" & Me.geometry.upperStructure(i).AntennaBotGirtOffset)
                        newERIList.Add("AntennaHasKBraceEndPanels=" & trueFalseYesNo(Me.geometry.upperStructure(i).AntennaHasKBraceEndPanels))
                        newERIList.Add("AntennaHasHorizontals=" & trueFalseYesNo(Me.geometry.upperStructure(i).AntennaHasHorizontals))
                        newERIList.Add("AntennaLongHorizontalType=" & Me.geometry.upperStructure(i).AntennaLongHorizontalType)
                        newERIList.Add("AntennaLongHorizontalSize=" & Me.geometry.upperStructure(i).AntennaLongHorizontalSize)
                        newERIList.Add("AntennaShortHorizontalType=" & Me.geometry.upperStructure(i).AntennaShortHorizontalType)
                        newERIList.Add("AntennaShortHorizontalSize=" & Me.geometry.upperStructure(i).AntennaShortHorizontalSize)
                        newERIList.Add("AntennaRedundantGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(Me.geometry.upperStructure(i).AntennaRedundantGrade))
                        newERIList.Add("AntennaRedundantMatlGrade=" & Me.geometry.upperStructure(i).AntennaRedundantMatlGrade)
                        newERIList.Add("AntennaRedundantType=" & Me.geometry.upperStructure(i).AntennaRedundantType)
                        newERIList.Add("AntennaRedundantDiagType=" & Me.geometry.upperStructure(i).AntennaRedundantDiagType)
                        newERIList.Add("AntennaRedundantSubDiagonalType=" & Me.geometry.upperStructure(i).AntennaRedundantSubDiagonalType)
                        newERIList.Add("AntennaRedundantSubHorizontalType=" & Me.geometry.upperStructure(i).AntennaRedundantSubHorizontalType)
                        newERIList.Add("AntennaRedundantVerticalType=" & Me.geometry.upperStructure(i).AntennaRedundantVerticalType)
                        newERIList.Add("AntennaRedundantHipType=" & Me.geometry.upperStructure(i).AntennaRedundantHipType)
                        newERIList.Add("AntennaRedundantHipDiagonalType=" & Me.geometry.upperStructure(i).AntennaRedundantHipDiagonalType)
                        newERIList.Add("AntennaRedundantHorizontalSize=" & Me.geometry.upperStructure(i).AntennaRedundantHorizontalSize)
                        newERIList.Add("AntennaRedundantHorizontalSize2=" & Me.geometry.upperStructure(i).AntennaRedundantHorizontalSize2)
                        newERIList.Add("AntennaRedundantHorizontalSize3=" & Me.geometry.upperStructure(i).AntennaRedundantHorizontalSize3)
                        newERIList.Add("AntennaRedundantHorizontalSize4=" & Me.geometry.upperStructure(i).AntennaRedundantHorizontalSize4)
                        newERIList.Add("AntennaRedundantDiagonalSize=" & Me.geometry.upperStructure(i).AntennaRedundantDiagonalSize)
                        newERIList.Add("AntennaRedundantDiagonalSize2=" & Me.geometry.upperStructure(i).AntennaRedundantDiagonalSize2)
                        newERIList.Add("AntennaRedundantDiagonalSize3=" & Me.geometry.upperStructure(i).AntennaRedundantDiagonalSize3)
                        newERIList.Add("AntennaRedundantDiagonalSize4=" & Me.geometry.upperStructure(i).AntennaRedundantDiagonalSize4)
                        newERIList.Add("AntennaRedundantSubHorizontalSize=" & Me.geometry.upperStructure(i).AntennaRedundantSubHorizontalSize)
                        newERIList.Add("AntennaRedundantSubDiagonalSize=" & Me.geometry.upperStructure(i).AntennaRedundantSubDiagonalSize)
                        newERIList.Add("AntennaSubDiagLocation=" & Me.geometry.upperStructure(i).AntennaSubDiagLocation)
                        newERIList.Add("AntennaRedundantVerticalSize=" & Me.geometry.upperStructure(i).AntennaRedundantVerticalSize)
                        newERIList.Add("AntennaRedundantHipDiagonalSize=" & Me.geometry.upperStructure(i).AntennaRedundantHipDiagonalSize)
                        newERIList.Add("AntennaRedundantHipDiagonalSize2=" & Me.geometry.upperStructure(i).AntennaRedundantHipDiagonalSize2)
                        newERIList.Add("AntennaRedundantHipDiagonalSize3=" & Me.geometry.upperStructure(i).AntennaRedundantHipDiagonalSize3)
                        newERIList.Add("AntennaRedundantHipDiagonalSize4=" & Me.geometry.upperStructure(i).AntennaRedundantHipDiagonalSize4)
                        newERIList.Add("AntennaRedundantHipSize=" & Me.geometry.upperStructure(i).AntennaRedundantHipSize)
                        newERIList.Add("AntennaRedundantHipSize2=" & Me.geometry.upperStructure(i).AntennaRedundantHipSize2)
                        newERIList.Add("AntennaRedundantHipSize3=" & Me.geometry.upperStructure(i).AntennaRedundantHipSize3)
                        newERIList.Add("AntennaRedundantHipSize4=" & Me.geometry.upperStructure(i).AntennaRedundantHipSize4)
                        newERIList.Add("AntennaNumInnerGirts=" & Me.geometry.upperStructure(i).AntennaNumInnerGirts)
                        newERIList.Add("AntennaInnerGirtType=" & Me.geometry.upperStructure(i).AntennaInnerGirtType)
                        newERIList.Add("AntennaInnerGirtSize=" & Me.geometry.upperStructure(i).AntennaInnerGirtSize)
                        newERIList.Add("AntennaPoleShapeType=" & Me.geometry.upperStructure(i).AntennaPoleShapeType)
                        newERIList.Add("AntennaPoleSize=" & Me.geometry.upperStructure(i).AntennaPoleSize)
                        newERIList.Add("AntennaPoleGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(Me.geometry.upperStructure(i).AntennaPoleGrade))
                        newERIList.Add("AntennaPoleMatlGrade=" & Me.geometry.upperStructure(i).AntennaPoleMatlGrade)
                        newERIList.Add("AntennaPoleSpliceLength=" & Me.geometry.upperStructure(i).AntennaPoleSpliceLength)
                        newERIList.Add("AntennaTaperPoleNumSides=" & Me.geometry.upperStructure(i).AntennaTaperPoleNumSides)
                        newERIList.Add("AntennaTaperPoleTopDiameter=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaTaperPoleTopDiameter))
                        newERIList.Add("AntennaTaperPoleBotDiameter=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaTaperPoleBotDiameter))
                        newERIList.Add("AntennaTaperPoleWallThickness=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaTaperPoleWallThickness))
                        newERIList.Add("AntennaTaperPoleBendRadius=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaTaperPoleBendRadius))
                        newERIList.Add("AntennaTaperPoleGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(Me.geometry.upperStructure(i).AntennaTaperPoleGrade))
                        newERIList.Add("AntennaTaperPoleMatlGrade=" & Me.geometry.upperStructure(i).AntennaTaperPoleMatlGrade)
                        newERIList.Add("AntennaSWMult=" & Me.geometry.upperStructure(i).AntennaSWMult)
                        newERIList.Add("AntennaWPMult=" & Me.geometry.upperStructure(i).AntennaWPMult)
                        newERIList.Add("AntennaAutoCalcKSingleAngle=" & Me.geometry.upperStructure(i).AntennaAutoCalcKSingleAngle)
                        newERIList.Add("AntennaAutoCalcKSolidRound=" & Me.geometry.upperStructure(i).AntennaAutoCalcKSolidRound)
                        newERIList.Add("AntennaAfGusset=" & Me.settings.USUnits.Length.convertToERIUnits(Me.geometry.upperStructure(i).AntennaAfGusset))
                        newERIList.Add("AntennaTfGusset=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaTfGusset))
                        newERIList.Add("AntennaGussetBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaGussetBoltEdgeDistance))
                        newERIList.Add("AntennaGussetGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(Me.geometry.upperStructure(i).AntennaGussetGrade))
                        newERIList.Add("AntennaGussetMatlGrade=" & Me.geometry.upperStructure(i).AntennaGussetMatlGrade)
                        newERIList.Add("AntennaAfMult=" & Me.geometry.upperStructure(i).AntennaAfMult)
                        newERIList.Add("AntennaArMult=" & Me.geometry.upperStructure(i).AntennaArMult)
                        newERIList.Add("AntennaFlatIPAPole=" & Me.geometry.upperStructure(i).AntennaFlatIPAPole)
                        newERIList.Add("AntennaRoundIPAPole=" & Me.geometry.upperStructure(i).AntennaRoundIPAPole)
                        newERIList.Add("AntennaFlatIPALeg=" & Me.geometry.upperStructure(i).AntennaFlatIPALeg)
                        newERIList.Add("AntennaRoundIPALeg=" & Me.geometry.upperStructure(i).AntennaRoundIPALeg)
                        newERIList.Add("AntennaFlatIPAHorizontal=" & Me.geometry.upperStructure(i).AntennaFlatIPAHorizontal)
                        newERIList.Add("AntennaRoundIPAHorizontal=" & Me.geometry.upperStructure(i).AntennaRoundIPAHorizontal)
                        newERIList.Add("AntennaFlatIPADiagonal=" & Me.geometry.upperStructure(i).AntennaFlatIPADiagonal)
                        newERIList.Add("AntennaRoundIPADiagonal=" & Me.geometry.upperStructure(i).AntennaRoundIPADiagonal)
                        newERIList.Add("AntennaCSA_S37_SpeedUpFactor=" & Me.geometry.upperStructure(i).AntennaCSA_S37_SpeedUpFactor)
                        newERIList.Add("AntennaKLegs=" & Me.geometry.upperStructure(i).AntennaKLegs)
                        newERIList.Add("AntennaKXBracedDiags=" & Me.geometry.upperStructure(i).AntennaKXBracedDiags)
                        newERIList.Add("AntennaKKBracedDiags=" & Me.geometry.upperStructure(i).AntennaKKBracedDiags)
                        newERIList.Add("AntennaKZBracedDiags=" & Me.geometry.upperStructure(i).AntennaKZBracedDiags)
                        newERIList.Add("AntennaKHorzs=" & Me.geometry.upperStructure(i).AntennaKHorzs)
                        newERIList.Add("AntennaKSecHorzs=" & Me.geometry.upperStructure(i).AntennaKSecHorzs)
                        newERIList.Add("AntennaKGirts=" & Me.geometry.upperStructure(i).AntennaKGirts)
                        newERIList.Add("AntennaKInners=" & Me.geometry.upperStructure(i).AntennaKInners)
                        newERIList.Add("AntennaKXBracedDiagsY=" & Me.geometry.upperStructure(i).AntennaKXBracedDiagsY)
                        newERIList.Add("AntennaKKBracedDiagsY=" & Me.geometry.upperStructure(i).AntennaKKBracedDiagsY)
                        newERIList.Add("AntennaKZBracedDiagsY=" & Me.geometry.upperStructure(i).AntennaKZBracedDiagsY)
                        newERIList.Add("AntennaKHorzsY=" & Me.geometry.upperStructure(i).AntennaKHorzsY)
                        newERIList.Add("AntennaKSecHorzsY=" & Me.geometry.upperStructure(i).AntennaKSecHorzsY)
                        newERIList.Add("AntennaKGirtsY=" & Me.geometry.upperStructure(i).AntennaKGirtsY)
                        newERIList.Add("AntennaKInnersY=" & Me.geometry.upperStructure(i).AntennaKInnersY)
                        newERIList.Add("AntennaKRedHorz=" & Me.geometry.upperStructure(i).AntennaKRedHorz)
                        newERIList.Add("AntennaKRedDiag=" & Me.geometry.upperStructure(i).AntennaKRedDiag)
                        newERIList.Add("AntennaKRedSubDiag=" & Me.geometry.upperStructure(i).AntennaKRedSubDiag)
                        newERIList.Add("AntennaKRedSubHorz=" & Me.geometry.upperStructure(i).AntennaKRedSubHorz)
                        newERIList.Add("AntennaKRedVert=" & Me.geometry.upperStructure(i).AntennaKRedVert)
                        newERIList.Add("AntennaKRedHip=" & Me.geometry.upperStructure(i).AntennaKRedHip)
                        newERIList.Add("AntennaKRedHipDiag=" & Me.geometry.upperStructure(i).AntennaKRedHipDiag)
                        newERIList.Add("AntennaKTLX=" & Me.geometry.upperStructure(i).AntennaKTLX)
                        newERIList.Add("AntennaKTLZ=" & Me.geometry.upperStructure(i).AntennaKTLZ)
                        newERIList.Add("AntennaKTLLeg=" & Me.geometry.upperStructure(i).AntennaKTLLeg)
                        newERIList.Add("AntennaInnerKTLX=" & Me.geometry.upperStructure(i).AntennaInnerKTLX)
                        newERIList.Add("AntennaInnerKTLZ=" & Me.geometry.upperStructure(i).AntennaInnerKTLZ)
                        newERIList.Add("AntennaInnerKTLLeg=" & Me.geometry.upperStructure(i).AntennaInnerKTLLeg)
                        newERIList.Add("AntennaStitchBoltLocationHoriz=" & Me.geometry.upperStructure(i).AntennaStitchBoltLocationHoriz)
                        newERIList.Add("AntennaStitchBoltLocationDiag=" & Me.geometry.upperStructure(i).AntennaStitchBoltLocationDiag)
                        newERIList.Add("AntennaStitchSpacing=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaStitchSpacing))
                        newERIList.Add("AntennaStitchSpacingHorz=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaStitchSpacingHorz))
                        newERIList.Add("AntennaStitchSpacingDiag=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaStitchSpacingDiag))
                        newERIList.Add("AntennaStitchSpacingRed=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaStitchSpacingRed))
                        newERIList.Add("AntennaLegNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaLegNetWidthDeduct))
                        newERIList.Add("AntennaLegUFactor=" & Me.geometry.upperStructure(i).AntennaLegUFactor)
                        newERIList.Add("AntennaDiagonalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaDiagonalNetWidthDeduct))
                        newERIList.Add("AntennaTopGirtNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaTopGirtNetWidthDeduct))
                        newERIList.Add("AntennaBotGirtNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaBotGirtNetWidthDeduct))
                        newERIList.Add("AntennaInnerGirtNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaInnerGirtNetWidthDeduct))
                        newERIList.Add("AntennaHorizontalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaHorizontalNetWidthDeduct))
                        newERIList.Add("AntennaShortHorizontalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaShortHorizontalNetWidthDeduct))
                        newERIList.Add("AntennaDiagonalUFactor=" & Me.geometry.upperStructure(i).AntennaDiagonalUFactor)
                        newERIList.Add("AntennaTopGirtUFactor=" & Me.geometry.upperStructure(i).AntennaTopGirtUFactor)
                        newERIList.Add("AntennaBotGirtUFactor=" & Me.geometry.upperStructure(i).AntennaBotGirtUFactor)
                        newERIList.Add("AntennaInnerGirtUFactor=" & Me.geometry.upperStructure(i).AntennaInnerGirtUFactor)
                        newERIList.Add("AntennaHorizontalUFactor=" & Me.geometry.upperStructure(i).AntennaHorizontalUFactor)
                        newERIList.Add("AntennaShortHorizontalUFactor=" & Me.geometry.upperStructure(i).AntennaShortHorizontalUFactor)
                        newERIList.Add("AntennaLegConnType=" & Me.geometry.upperStructure(i).AntennaLegConnType)
                        newERIList.Add("AntennaLegNumBolts=" & Me.geometry.upperStructure(i).AntennaLegNumBolts)
                        newERIList.Add("AntennaDiagonalNumBolts=" & Me.geometry.upperStructure(i).AntennaDiagonalNumBolts)
                        newERIList.Add("AntennaTopGirtNumBolts=" & Me.geometry.upperStructure(i).AntennaTopGirtNumBolts)
                        newERIList.Add("AntennaBotGirtNumBolts=" & Me.geometry.upperStructure(i).AntennaBotGirtNumBolts)
                        newERIList.Add("AntennaInnerGirtNumBolts=" & Me.geometry.upperStructure(i).AntennaInnerGirtNumBolts)
                        newERIList.Add("AntennaHorizontalNumBolts=" & Me.geometry.upperStructure(i).AntennaHorizontalNumBolts)
                        newERIList.Add("AntennaShortHorizontalNumBolts=" & Me.geometry.upperStructure(i).AntennaShortHorizontalNumBolts)
                        newERIList.Add("AntennaLegBoltGrade=" & Me.geometry.upperStructure(i).AntennaLegBoltGrade)
                        newERIList.Add("AntennaLegBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaLegBoltSize))
                        newERIList.Add("AntennaDiagonalBoltGrade=" & Me.geometry.upperStructure(i).AntennaDiagonalBoltGrade)
                        newERIList.Add("AntennaDiagonalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaDiagonalBoltSize))
                        newERIList.Add("AntennaTopGirtBoltGrade=" & Me.geometry.upperStructure(i).AntennaTopGirtBoltGrade)
                        newERIList.Add("AntennaTopGirtBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaTopGirtBoltSize))
                        newERIList.Add("AntennaBotGirtBoltGrade=" & Me.geometry.upperStructure(i).AntennaBotGirtBoltGrade)
                        newERIList.Add("AntennaBotGirtBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaBotGirtBoltSize))
                        newERIList.Add("AntennaInnerGirtBoltGrade=" & Me.geometry.upperStructure(i).AntennaInnerGirtBoltGrade)
                        newERIList.Add("AntennaInnerGirtBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaInnerGirtBoltSize))
                        newERIList.Add("AntennaHorizontalBoltGrade=" & Me.geometry.upperStructure(i).AntennaHorizontalBoltGrade)
                        newERIList.Add("AntennaHorizontalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaHorizontalBoltSize))
                        newERIList.Add("AntennaShortHorizontalBoltGrade=" & Me.geometry.upperStructure(i).AntennaShortHorizontalBoltGrade)
                        newERIList.Add("AntennaShortHorizontalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaShortHorizontalBoltSize))
                        newERIList.Add("AntennaLegBoltEdgeDistance=" & Me.geometry.upperStructure(i).AntennaLegBoltEdgeDistance)
                        newERIList.Add("AntennaDiagonalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaDiagonalBoltEdgeDistance))
                        newERIList.Add("AntennaTopGirtBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaTopGirtBoltEdgeDistance))
                        newERIList.Add("AntennaBotGirtBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaBotGirtBoltEdgeDistance))
                        newERIList.Add("AntennaInnerGirtBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaInnerGirtBoltEdgeDistance))
                        newERIList.Add("AntennaHorizontalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaHorizontalBoltEdgeDistance))
                        newERIList.Add("AntennaShortHorizontalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaShortHorizontalBoltEdgeDistance))
                        newERIList.Add("AntennaDiagonalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaDiagonalGageG1Distance))
                        newERIList.Add("AntennaTopGirtGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaTopGirtGageG1Distance))
                        newERIList.Add("AntennaBotGirtGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaBotGirtGageG1Distance))
                        newERIList.Add("AntennaInnerGirtGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaInnerGirtGageG1Distance))
                        newERIList.Add("AntennaHorizontalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaHorizontalGageG1Distance))
                        newERIList.Add("AntennaShortHorizontalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaShortHorizontalGageG1Distance))
                        newERIList.Add("AntennaRedundantHorizontalBoltGrade=" & Me.geometry.upperStructure(i).AntennaRedundantHorizontalBoltGrade)
                        newERIList.Add("AntennaRedundantHorizontalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaRedundantHorizontalBoltSize))
                        newERIList.Add("AntennaRedundantHorizontalNumBolts=" & Me.geometry.upperStructure(i).AntennaRedundantHorizontalNumBolts)
                        newERIList.Add("AntennaRedundantHorizontalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaRedundantHorizontalBoltEdgeDistance))
                        newERIList.Add("AntennaRedundantHorizontalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaRedundantHorizontalGageG1Distance))
                        newERIList.Add("AntennaRedundantHorizontalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaRedundantHorizontalNetWidthDeduct))
                        newERIList.Add("AntennaRedundantHorizontalUFactor=" & Me.geometry.upperStructure(i).AntennaRedundantHorizontalUFactor)
                        newERIList.Add("AntennaRedundantDiagonalBoltGrade=" & Me.geometry.upperStructure(i).AntennaRedundantDiagonalBoltGrade)
                        newERIList.Add("AntennaRedundantDiagonalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaRedundantDiagonalBoltSize))
                        newERIList.Add("AntennaRedundantDiagonalNumBolts=" & Me.geometry.upperStructure(i).AntennaRedundantDiagonalNumBolts)
                        newERIList.Add("AntennaRedundantDiagonalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaRedundantDiagonalBoltEdgeDistance))
                        newERIList.Add("AntennaRedundantDiagonalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaRedundantDiagonalGageG1Distance))
                        newERIList.Add("AntennaRedundantDiagonalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaRedundantDiagonalNetWidthDeduct))
                        newERIList.Add("AntennaRedundantDiagonalUFactor=" & Me.geometry.upperStructure(i).AntennaRedundantDiagonalUFactor)
                        newERIList.Add("AntennaRedundantSubDiagonalBoltGrade=" & Me.geometry.upperStructure(i).AntennaRedundantSubDiagonalBoltGrade)
                        newERIList.Add("AntennaRedundantSubDiagonalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaRedundantSubDiagonalBoltSize))
                        newERIList.Add("AntennaRedundantSubDiagonalNumBolts=" & Me.geometry.upperStructure(i).AntennaRedundantSubDiagonalNumBolts)
                        newERIList.Add("AntennaRedundantSubDiagonalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaRedundantSubDiagonalBoltEdgeDistance))
                        newERIList.Add("AntennaRedundantSubDiagonalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaRedundantSubDiagonalGageG1Distance))
                        newERIList.Add("AntennaRedundantSubDiagonalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaRedundantSubDiagonalNetWidthDeduct))
                        newERIList.Add("AntennaRedundantSubDiagonalUFactor=" & Me.geometry.upperStructure(i).AntennaRedundantSubDiagonalUFactor)
                        newERIList.Add("AntennaRedundantSubHorizontalBoltGrade=" & Me.geometry.upperStructure(i).AntennaRedundantSubHorizontalBoltGrade)
                        newERIList.Add("AntennaRedundantSubHorizontalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaRedundantSubHorizontalBoltSize))
                        newERIList.Add("AntennaRedundantSubHorizontalNumBolts=" & Me.geometry.upperStructure(i).AntennaRedundantSubHorizontalNumBolts)
                        newERIList.Add("AntennaRedundantSubHorizontalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaRedundantSubHorizontalBoltEdgeDistance))
                        newERIList.Add("AntennaRedundantSubHorizontalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaRedundantSubHorizontalGageG1Distance))
                        newERIList.Add("AntennaRedundantSubHorizontalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaRedundantSubHorizontalNetWidthDeduct))
                        newERIList.Add("AntennaRedundantSubHorizontalUFactor=" & Me.geometry.upperStructure(i).AntennaRedundantSubHorizontalUFactor)
                        newERIList.Add("AntennaRedundantVerticalBoltGrade=" & Me.geometry.upperStructure(i).AntennaRedundantVerticalBoltGrade)
                        newERIList.Add("AntennaRedundantVerticalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaRedundantVerticalBoltSize))
                        newERIList.Add("AntennaRedundantVerticalNumBolts=" & Me.geometry.upperStructure(i).AntennaRedundantVerticalNumBolts)
                        newERIList.Add("AntennaRedundantVerticalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaRedundantVerticalBoltEdgeDistance))
                        newERIList.Add("AntennaRedundantVerticalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaRedundantVerticalGageG1Distance))
                        newERIList.Add("AntennaRedundantVerticalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaRedundantVerticalNetWidthDeduct))
                        newERIList.Add("AntennaRedundantVerticalUFactor=" & Me.geometry.upperStructure(i).AntennaRedundantVerticalUFactor)
                        newERIList.Add("AntennaRedundantHipBoltGrade=" & Me.geometry.upperStructure(i).AntennaRedundantHipBoltGrade)
                        newERIList.Add("AntennaRedundantHipBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaRedundantHipBoltSize))
                        newERIList.Add("AntennaRedundantHipNumBolts=" & Me.geometry.upperStructure(i).AntennaRedundantHipNumBolts)
                        newERIList.Add("AntennaRedundantHipBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaRedundantHipBoltEdgeDistance))
                        newERIList.Add("AntennaRedundantHipGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaRedundantHipGageG1Distance))
                        newERIList.Add("AntennaRedundantHipNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaRedundantHipNetWidthDeduct))
                        newERIList.Add("AntennaRedundantHipUFactor=" & Me.geometry.upperStructure(i).AntennaRedundantHipUFactor)
                        newERIList.Add("AntennaRedundantHipDiagonalBoltGrade=" & Me.geometry.upperStructure(i).AntennaRedundantHipDiagonalBoltGrade)
                        newERIList.Add("AntennaRedundantHipDiagonalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaRedundantHipDiagonalBoltSize))
                        newERIList.Add("AntennaRedundantHipDiagonalNumBolts=" & Me.geometry.upperStructure(i).AntennaRedundantHipDiagonalNumBolts)
                        newERIList.Add("AntennaRedundantHipDiagonalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaRedundantHipDiagonalBoltEdgeDistance))
                        newERIList.Add("AntennaRedundantHipDiagonalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaRedundantHipDiagonalGageG1Distance))
                        newERIList.Add("AntennaRedundantHipDiagonalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaRedundantHipDiagonalNetWidthDeduct))
                        newERIList.Add("AntennaRedundantHipDiagonalUFactor=" & Me.geometry.upperStructure(i).AntennaRedundantHipDiagonalUFactor)
                        newERIList.Add("AntennaDiagonalOutOfPlaneRestraint=" & trueFalseYesNo(Me.geometry.upperStructure(i).AntennaDiagonalOutOfPlaneRestraint))
                        newERIList.Add("AntennaTopGirtOutOfPlaneRestraint=" & trueFalseYesNo(Me.geometry.upperStructure(i).AntennaTopGirtOutOfPlaneRestraint))
                        newERIList.Add("AntennaBottomGirtOutOfPlaneRestraint=" & trueFalseYesNo(Me.geometry.upperStructure(i).AntennaBottomGirtOutOfPlaneRestraint))
                        newERIList.Add("AntennaMidGirtOutOfPlaneRestraint=" & trueFalseYesNo(Me.geometry.upperStructure(i).AntennaMidGirtOutOfPlaneRestraint))
                        newERIList.Add("AntennaHorizontalOutOfPlaneRestraint=" & trueFalseYesNo(Me.geometry.upperStructure(i).AntennaHorizontalOutOfPlaneRestraint))
                        newERIList.Add("AntennaSecondaryHorizontalOutOfPlaneRestraint=" & trueFalseYesNo(Me.geometry.upperStructure(i).AntennaSecondaryHorizontalOutOfPlaneRestraint))
                        newERIList.Add("AntennaDiagOffsetNEY=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaDiagOffsetNEY))
                        newERIList.Add("AntennaDiagOffsetNEX=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaDiagOffsetNEX))
                        newERIList.Add("AntennaDiagOffsetPEY=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaDiagOffsetPEY))
                        newERIList.Add("AntennaDiagOffsetPEX=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaDiagOffsetPEX))
                        newERIList.Add("AntennaKbraceOffsetNEY=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaKbraceOffsetNEY))
                        newERIList.Add("AntennaKbraceOffsetNEX=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaKbraceOffsetNEX))
                        newERIList.Add("AntennaKbraceOffsetPEY=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaKbraceOffsetPEY))
                        newERIList.Add("AntennaKbraceOffsetPEX=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.upperStructure(i).AntennaKbraceOffsetPEX))

                    Next i
                Case line(0).Equals("NumTowerRecs")
                    newERIList.Add(line(0) & "=" & Me.geometry.baseStructure.Count)
                    For i = 0 To Me.geometry.baseStructure.Count - 1
                        newERIList.Add("TowerRec=" & Me.geometry.baseStructure(i).TowerRec)
                        newERIList.Add("TowerDatabase=" & Me.geometry.baseStructure(i).TowerDatabase)
                        newERIList.Add("TowerName=" & Me.geometry.baseStructure(i).TowerName)
                        newERIList.Add("TowerHeight=" & Me.settings.USUnits.Coordinate.convertToERIUnits(Me.geometry.baseStructure(i).TowerHeight))
                        newERIList.Add("TowerFaceWidth=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerFaceWidth))
                        newERIList.Add("TowerNumSections=" & Me.geometry.baseStructure(i).TowerNumSections)
                        newERIList.Add("TowerSectionLength=" & Me.settings.USUnits.Length.convertToERIUnits(Me.geometry.baseStructure(i).TowerSectionLength))
                        newERIList.Add("TowerDiagonalSpacing=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerDiagonalSpacing))
                        newERIList.Add("TowerDiagonalSpacingEx=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerDiagonalSpacingEx))
                        newERIList.Add("TowerBraceType=" & Me.geometry.baseStructure(i).TowerBraceType)
                        newERIList.Add("TowerFaceBevel=" & Me.geometry.baseStructure(i).TowerFaceBevel)
                        newERIList.Add("TowerTopGirtOffset=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerTopGirtOffset))
                        newERIList.Add("TowerBotGirtOffset=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerBotGirtOffset))
                        newERIList.Add("TowerHasKBraceEndPanels=" & trueFalseYesNo(Me.geometry.baseStructure(i).TowerHasKBraceEndPanels))
                        newERIList.Add("TowerHasHorizontals=" & trueFalseYesNo(Me.geometry.baseStructure(i).TowerHasHorizontals))
                        newERIList.Add("TowerLegType=" & Me.geometry.baseStructure(i).TowerLegType)
                        newERIList.Add("TowerLegSize=" & Me.geometry.baseStructure(i).TowerLegSize)
                        newERIList.Add("TowerLegGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(Me.geometry.baseStructure(i).TowerLegGrade))
                        newERIList.Add("TowerLegMatlGrade=" & Me.geometry.baseStructure(i).TowerLegMatlGrade)
                        newERIList.Add("TowerDiagonalGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(Me.geometry.baseStructure(i).TowerDiagonalGrade))
                        newERIList.Add("TowerDiagonalMatlGrade=" & Me.geometry.baseStructure(i).TowerDiagonalMatlGrade)
                        newERIList.Add("TowerInnerBracingGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(Me.geometry.baseStructure(i).TowerInnerBracingGrade))
                        newERIList.Add("TowerInnerBracingMatlGrade=" & Me.geometry.baseStructure(i).TowerInnerBracingMatlGrade)
                        newERIList.Add("TowerTopGirtGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(Me.geometry.baseStructure(i).TowerTopGirtGrade))
                        newERIList.Add("TowerTopGirtMatlGrade=" & Me.geometry.baseStructure(i).TowerTopGirtMatlGrade)
                        newERIList.Add("TowerBotGirtGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(Me.geometry.baseStructure(i).TowerBotGirtGrade))
                        newERIList.Add("TowerBotGirtMatlGrade=" & Me.geometry.baseStructure(i).TowerBotGirtMatlGrade)
                        newERIList.Add("TowerInnerGirtGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(Me.geometry.baseStructure(i).TowerInnerGirtGrade))
                        newERIList.Add("TowerInnerGirtMatlGrade=" & Me.geometry.baseStructure(i).TowerInnerGirtMatlGrade)
                        newERIList.Add("TowerLongHorizontalGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(Me.geometry.baseStructure(i).TowerLongHorizontalGrade))
                        newERIList.Add("TowerLongHorizontalMatlGrade=" & Me.geometry.baseStructure(i).TowerLongHorizontalMatlGrade)
                        newERIList.Add("TowerShortHorizontalGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(Me.geometry.baseStructure(i).TowerShortHorizontalGrade))
                        newERIList.Add("TowerShortHorizontalMatlGrade=" & Me.geometry.baseStructure(i).TowerShortHorizontalMatlGrade)
                        newERIList.Add("TowerDiagonalType=" & Me.geometry.baseStructure(i).TowerDiagonalType)
                        newERIList.Add("TowerDiagonalSize=" & Me.geometry.baseStructure(i).TowerDiagonalSize)
                        newERIList.Add("TowerInnerBracingType=" & Me.geometry.baseStructure(i).TowerInnerBracingType)
                        newERIList.Add("TowerInnerBracingSize=" & Me.geometry.baseStructure(i).TowerInnerBracingSize)
                        newERIList.Add("TowerTopGirtType=" & Me.geometry.baseStructure(i).TowerTopGirtType)
                        newERIList.Add("TowerTopGirtSize=" & Me.geometry.baseStructure(i).TowerTopGirtSize)
                        newERIList.Add("TowerBotGirtType=" & Me.geometry.baseStructure(i).TowerBotGirtType)
                        newERIList.Add("TowerBotGirtSize=" & Me.geometry.baseStructure(i).TowerBotGirtSize)
                        newERIList.Add("TowerNumInnerGirts=" & Me.geometry.baseStructure(i).TowerNumInnerGirts)
                        newERIList.Add("TowerInnerGirtType=" & Me.geometry.baseStructure(i).TowerInnerGirtType)
                        newERIList.Add("TowerInnerGirtSize=" & Me.geometry.baseStructure(i).TowerInnerGirtSize)
                        newERIList.Add("TowerLongHorizontalType=" & Me.geometry.baseStructure(i).TowerLongHorizontalType)
                        newERIList.Add("TowerLongHorizontalSize=" & Me.geometry.baseStructure(i).TowerLongHorizontalSize)
                        newERIList.Add("TowerShortHorizontalType=" & Me.geometry.baseStructure(i).TowerShortHorizontalType)
                        newERIList.Add("TowerShortHorizontalSize=" & Me.geometry.baseStructure(i).TowerShortHorizontalSize)
                        newERIList.Add("TowerRedundantGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(Me.geometry.baseStructure(i).TowerRedundantGrade))
                        newERIList.Add("TowerRedundantMatlGrade=" & Me.geometry.baseStructure(i).TowerRedundantMatlGrade)
                        newERIList.Add("TowerRedundantType=" & Me.geometry.baseStructure(i).TowerRedundantType)
                        newERIList.Add("TowerRedundantDiagType=" & Me.geometry.baseStructure(i).TowerRedundantDiagType)
                        newERIList.Add("TowerRedundantSubDiagonalType=" & Me.geometry.baseStructure(i).TowerRedundantSubDiagonalType)
                        newERIList.Add("TowerRedundantSubHorizontalType=" & Me.geometry.baseStructure(i).TowerRedundantSubHorizontalType)
                        newERIList.Add("TowerRedundantVerticalType=" & Me.geometry.baseStructure(i).TowerRedundantVerticalType)
                        newERIList.Add("TowerRedundantHipType=" & Me.geometry.baseStructure(i).TowerRedundantHipType)
                        newERIList.Add("TowerRedundantHipDiagonalType=" & Me.geometry.baseStructure(i).TowerRedundantHipDiagonalType)
                        newERIList.Add("TowerRedundantHorizontalSize=" & Me.geometry.baseStructure(i).TowerRedundantHorizontalSize)
                        newERIList.Add("TowerRedundantHorizontalSize2=" & Me.geometry.baseStructure(i).TowerRedundantHorizontalSize2)
                        newERIList.Add("TowerRedundantHorizontalSize3=" & Me.geometry.baseStructure(i).TowerRedundantHorizontalSize3)
                        newERIList.Add("TowerRedundantHorizontalSize4=" & Me.geometry.baseStructure(i).TowerRedundantHorizontalSize4)
                        newERIList.Add("TowerRedundantDiagonalSize=" & Me.geometry.baseStructure(i).TowerRedundantDiagonalSize)
                        newERIList.Add("TowerRedundantDiagonalSize2=" & Me.geometry.baseStructure(i).TowerRedundantDiagonalSize2)
                        newERIList.Add("TowerRedundantDiagonalSize3=" & Me.geometry.baseStructure(i).TowerRedundantDiagonalSize3)
                        newERIList.Add("TowerRedundantDiagonalSize4=" & Me.geometry.baseStructure(i).TowerRedundantDiagonalSize4)
                        newERIList.Add("TowerRedundantSubHorizontalSize=" & Me.geometry.baseStructure(i).TowerRedundantSubHorizontalSize)
                        newERIList.Add("TowerRedundantSubDiagonalSize=" & Me.geometry.baseStructure(i).TowerRedundantSubDiagonalSize)
                        newERIList.Add("TowerSubDiagLocation=" & Me.geometry.baseStructure(i).TowerSubDiagLocation)
                        newERIList.Add("TowerRedundantVerticalSize=" & Me.geometry.baseStructure(i).TowerRedundantVerticalSize)
                        newERIList.Add("TowerRedundantHipSize=" & Me.geometry.baseStructure(i).TowerRedundantHipSize)
                        newERIList.Add("TowerRedundantHipSize2=" & Me.geometry.baseStructure(i).TowerRedundantHipSize2)
                        newERIList.Add("TowerRedundantHipSize3=" & Me.geometry.baseStructure(i).TowerRedundantHipSize3)
                        newERIList.Add("TowerRedundantHipSize4=" & Me.geometry.baseStructure(i).TowerRedundantHipSize4)
                        newERIList.Add("TowerRedundantHipDiagonalSize=" & Me.geometry.baseStructure(i).TowerRedundantHipDiagonalSize)
                        newERIList.Add("TowerRedundantHipDiagonalSize2=" & Me.geometry.baseStructure(i).TowerRedundantHipDiagonalSize2)
                        newERIList.Add("TowerRedundantHipDiagonalSize3=" & Me.geometry.baseStructure(i).TowerRedundantHipDiagonalSize3)
                        newERIList.Add("TowerRedundantHipDiagonalSize4=" & Me.geometry.baseStructure(i).TowerRedundantHipDiagonalSize4)
                        newERIList.Add("TowerSWMult=" & Me.geometry.baseStructure(i).TowerSWMult)
                        newERIList.Add("TowerWPMult=" & Me.geometry.baseStructure(i).TowerWPMult)
                        newERIList.Add("TowerAutoCalcKSingleAngle=" & trueFalseYesNo(Me.geometry.baseStructure(i).TowerAutoCalcKSingleAngle))
                        newERIList.Add("TowerAutoCalcKSolidRound=" & trueFalseYesNo(Me.geometry.baseStructure(i).TowerAutoCalcKSolidRound))
                        newERIList.Add("TowerAfGusset=" & Me.settings.USUnits.Length.convertToERIUnits(Me.geometry.baseStructure(i).TowerAfGusset))
                        newERIList.Add("TowerTfGusset=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerTfGusset))
                        newERIList.Add("TowerGussetBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerGussetBoltEdgeDistance))
                        newERIList.Add("TowerGussetGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(Me.geometry.baseStructure(i).TowerGussetGrade))
                        newERIList.Add("TowerGussetMatlGrade=" & Me.geometry.baseStructure(i).TowerGussetMatlGrade)
                        newERIList.Add("TowerAfMult=" & Me.geometry.baseStructure(i).TowerAfMult)
                        newERIList.Add("TowerArMult=" & Me.geometry.baseStructure(i).TowerArMult)
                        newERIList.Add("TowerFlatIPAPole=" & Me.geometry.baseStructure(i).TowerFlatIPAPole)
                        newERIList.Add("TowerRoundIPAPole=" & Me.geometry.baseStructure(i).TowerRoundIPAPole)
                        newERIList.Add("TowerFlatIPALeg=" & Me.geometry.baseStructure(i).TowerFlatIPALeg)
                        newERIList.Add("TowerRoundIPALeg=" & Me.geometry.baseStructure(i).TowerRoundIPALeg)
                        newERIList.Add("TowerFlatIPAHorizontal=" & Me.geometry.baseStructure(i).TowerFlatIPAHorizontal)
                        newERIList.Add("TowerRoundIPAHorizontal=" & Me.geometry.baseStructure(i).TowerRoundIPAHorizontal)
                        newERIList.Add("TowerFlatIPADiagonal=" & Me.geometry.baseStructure(i).TowerFlatIPADiagonal)
                        newERIList.Add("TowerRoundIPADiagonal=" & Me.geometry.baseStructure(i).TowerRoundIPADiagonal)
                        newERIList.Add("TowerCSA_S37_SpeedUpFactor=" & Me.geometry.baseStructure(i).TowerCSA_S37_SpeedUpFactor)
                        newERIList.Add("TowerKLegs=" & Me.geometry.baseStructure(i).TowerKLegs)
                        newERIList.Add("TowerKXBracedDiags=" & Me.geometry.baseStructure(i).TowerKXBracedDiags)
                        newERIList.Add("TowerKKBracedDiags=" & Me.geometry.baseStructure(i).TowerKKBracedDiags)
                        newERIList.Add("TowerKZBracedDiags=" & Me.geometry.baseStructure(i).TowerKZBracedDiags)
                        newERIList.Add("TowerKHorzs=" & Me.geometry.baseStructure(i).TowerKHorzs)
                        newERIList.Add("TowerKSecHorzs=" & Me.geometry.baseStructure(i).TowerKSecHorzs)
                        newERIList.Add("TowerKGirts=" & Me.geometry.baseStructure(i).TowerKGirts)
                        newERIList.Add("TowerKInners=" & Me.geometry.baseStructure(i).TowerKInners)
                        newERIList.Add("TowerKXBracedDiagsY=" & Me.geometry.baseStructure(i).TowerKXBracedDiagsY)
                        newERIList.Add("TowerKKBracedDiagsY=" & Me.geometry.baseStructure(i).TowerKKBracedDiagsY)
                        newERIList.Add("TowerKZBracedDiagsY=" & Me.geometry.baseStructure(i).TowerKZBracedDiagsY)
                        newERIList.Add("TowerKHorzsY=" & Me.geometry.baseStructure(i).TowerKHorzsY)
                        newERIList.Add("TowerKSecHorzsY=" & Me.geometry.baseStructure(i).TowerKSecHorzsY)
                        newERIList.Add("TowerKGirtsY=" & Me.geometry.baseStructure(i).TowerKGirtsY)
                        newERIList.Add("TowerKInnersY=" & Me.geometry.baseStructure(i).TowerKInnersY)
                        newERIList.Add("TowerKRedHorz=" & Me.geometry.baseStructure(i).TowerKRedHorz)
                        newERIList.Add("TowerKRedDiag=" & Me.geometry.baseStructure(i).TowerKRedDiag)
                        newERIList.Add("TowerKRedSubDiag=" & Me.geometry.baseStructure(i).TowerKRedSubDiag)
                        newERIList.Add("TowerKRedSubHorz=" & Me.geometry.baseStructure(i).TowerKRedSubHorz)
                        newERIList.Add("TowerKRedVert=" & Me.geometry.baseStructure(i).TowerKRedVert)
                        newERIList.Add("TowerKRedHip=" & Me.geometry.baseStructure(i).TowerKRedHip)
                        newERIList.Add("TowerKRedHipDiag=" & Me.geometry.baseStructure(i).TowerKRedHipDiag)
                        newERIList.Add("TowerKTLX=" & Me.geometry.baseStructure(i).TowerKTLX)
                        newERIList.Add("TowerKTLZ=" & Me.geometry.baseStructure(i).TowerKTLZ)
                        newERIList.Add("TowerKTLLeg=" & Me.geometry.baseStructure(i).TowerKTLLeg)
                        newERIList.Add("TowerInnerKTLX=" & Me.geometry.baseStructure(i).TowerInnerKTLX)
                        newERIList.Add("TowerInnerKTLZ=" & Me.geometry.baseStructure(i).TowerInnerKTLZ)
                        newERIList.Add("TowerInnerKTLLeg=" & Me.geometry.baseStructure(i).TowerInnerKTLLeg)
                        newERIList.Add("TowerStitchBoltLocationHoriz=" & Me.geometry.baseStructure(i).TowerStitchBoltLocationHoriz)
                        newERIList.Add("TowerStitchBoltLocationDiag=" & Me.geometry.baseStructure(i).TowerStitchBoltLocationDiag)
                        newERIList.Add("TowerStitchBoltLocationRed=" & Me.geometry.baseStructure(i).TowerStitchBoltLocationRed)
                        newERIList.Add("TowerStitchSpacing=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerStitchSpacing))
                        newERIList.Add("TowerStitchSpacingDiag=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerStitchSpacingDiag))
                        newERIList.Add("TowerStitchSpacingHorz=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerStitchSpacingHorz))
                        newERIList.Add("TowerStitchSpacingRed=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerStitchSpacingRed))
                        newERIList.Add("TowerLegNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerLegNetWidthDeduct))
                        newERIList.Add("TowerLegUFactor=" & Me.geometry.baseStructure(i).TowerLegUFactor)
                        newERIList.Add("TowerDiagonalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerDiagonalNetWidthDeduct))
                        newERIList.Add("TowerTopGirtNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerTopGirtNetWidthDeduct))
                        newERIList.Add("TowerBotGirtNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerBotGirtNetWidthDeduct))
                        newERIList.Add("TowerInnerGirtNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerInnerGirtNetWidthDeduct))
                        newERIList.Add("TowerHorizontalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerHorizontalNetWidthDeduct))
                        newERIList.Add("TowerShortHorizontalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerShortHorizontalNetWidthDeduct))
                        newERIList.Add("TowerDiagonalUFactor=" & Me.geometry.baseStructure(i).TowerDiagonalUFactor)
                        newERIList.Add("TowerTopGirtUFactor=" & Me.geometry.baseStructure(i).TowerTopGirtUFactor)
                        newERIList.Add("TowerBotGirtUFactor=" & Me.geometry.baseStructure(i).TowerBotGirtUFactor)
                        newERIList.Add("TowerInnerGirtUFactor=" & Me.geometry.baseStructure(i).TowerInnerGirtUFactor)
                        newERIList.Add("TowerHorizontalUFactor=" & Me.geometry.baseStructure(i).TowerHorizontalUFactor)
                        newERIList.Add("TowerShortHorizontalUFactor=" & Me.geometry.baseStructure(i).TowerShortHorizontalUFactor)
                        newERIList.Add("TowerLegConnType=" & Me.geometry.baseStructure(i).TowerLegConnType)
                        newERIList.Add("TowerLegNumBolts=" & Me.geometry.baseStructure(i).TowerLegNumBolts)
                        newERIList.Add("TowerDiagonalNumBolts=" & Me.geometry.baseStructure(i).TowerDiagonalNumBolts)
                        newERIList.Add("TowerTopGirtNumBolts=" & Me.geometry.baseStructure(i).TowerTopGirtNumBolts)
                        newERIList.Add("TowerBotGirtNumBolts=" & Me.geometry.baseStructure(i).TowerBotGirtNumBolts)
                        newERIList.Add("TowerInnerGirtNumBolts=" & Me.geometry.baseStructure(i).TowerInnerGirtNumBolts)
                        newERIList.Add("TowerHorizontalNumBolts=" & Me.geometry.baseStructure(i).TowerHorizontalNumBolts)
                        newERIList.Add("TowerShortHorizontalNumBolts=" & Me.geometry.baseStructure(i).TowerShortHorizontalNumBolts)
                        newERIList.Add("TowerLegBoltGrade=" & Me.geometry.baseStructure(i).TowerLegBoltGrade)
                        newERIList.Add("TowerLegBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerLegBoltSize))
                        newERIList.Add("TowerDiagonalBoltGrade=" & Me.geometry.baseStructure(i).TowerDiagonalBoltGrade)
                        newERIList.Add("TowerDiagonalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerDiagonalBoltSize))
                        newERIList.Add("TowerTopGirtBoltGrade=" & Me.geometry.baseStructure(i).TowerTopGirtBoltGrade)
                        newERIList.Add("TowerTopGirtBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerTopGirtBoltSize))
                        newERIList.Add("TowerBotGirtBoltGrade=" & Me.geometry.baseStructure(i).TowerBotGirtBoltGrade)
                        newERIList.Add("TowerBotGirtBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerBotGirtBoltSize))
                        newERIList.Add("TowerInnerGirtBoltGrade=" & Me.geometry.baseStructure(i).TowerInnerGirtBoltGrade)
                        newERIList.Add("TowerInnerGirtBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerInnerGirtBoltSize))
                        newERIList.Add("TowerHorizontalBoltGrade=" & Me.geometry.baseStructure(i).TowerHorizontalBoltGrade)
                        newERIList.Add("TowerHorizontalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerHorizontalBoltSize))
                        newERIList.Add("TowerShortHorizontalBoltGrade=" & Me.geometry.baseStructure(i).TowerShortHorizontalBoltGrade)
                        newERIList.Add("TowerShortHorizontalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerShortHorizontalBoltSize))
                        newERIList.Add("TowerLegBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerLegBoltEdgeDistance))
                        newERIList.Add("TowerDiagonalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerDiagonalBoltEdgeDistance))
                        newERIList.Add("TowerTopGirtBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerTopGirtBoltEdgeDistance))
                        newERIList.Add("TowerBotGirtBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerBotGirtBoltEdgeDistance))
                        newERIList.Add("TowerInnerGirtBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerInnerGirtBoltEdgeDistance))
                        newERIList.Add("TowerHorizontalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerHorizontalBoltEdgeDistance))
                        newERIList.Add("TowerShortHorizontalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerShortHorizontalBoltEdgeDistance))
                        newERIList.Add("TowerDiagonalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerDiagonalGageG1Distance))
                        newERIList.Add("TowerTopGirtGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerTopGirtGageG1Distance))
                        newERIList.Add("TowerBotGirtGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerBotGirtGageG1Distance))
                        newERIList.Add("TowerInnerGirtGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerInnerGirtGageG1Distance))
                        newERIList.Add("TowerHorizontalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerHorizontalGageG1Distance))
                        newERIList.Add("TowerShortHorizontalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerShortHorizontalGageG1Distance))
                        newERIList.Add("TowerRedundantHorizontalBoltGrade=" & Me.geometry.baseStructure(i).TowerRedundantHorizontalBoltGrade)
                        newERIList.Add("TowerRedundantHorizontalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerRedundantHorizontalBoltSize))
                        newERIList.Add("TowerRedundantHorizontalNumBolts=" & Me.geometry.baseStructure(i).TowerRedundantHorizontalNumBolts)
                        newERIList.Add("TowerRedundantHorizontalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerRedundantHorizontalBoltEdgeDistance))
                        newERIList.Add("TowerRedundantHorizontalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerRedundantHorizontalGageG1Distance))
                        newERIList.Add("TowerRedundantHorizontalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerRedundantHorizontalNetWidthDeduct))
                        newERIList.Add("TowerRedundantHorizontalUFactor=" & Me.geometry.baseStructure(i).TowerRedundantHorizontalUFactor)
                        newERIList.Add("TowerRedundantDiagonalBoltGrade=" & Me.geometry.baseStructure(i).TowerRedundantDiagonalBoltGrade)
                        newERIList.Add("TowerRedundantDiagonalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerRedundantDiagonalBoltSize))
                        newERIList.Add("TowerRedundantDiagonalNumBolts=" & Me.geometry.baseStructure(i).TowerRedundantDiagonalNumBolts)
                        newERIList.Add("TowerRedundantDiagonalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerRedundantDiagonalBoltEdgeDistance))
                        newERIList.Add("TowerRedundantDiagonalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerRedundantDiagonalGageG1Distance))
                        newERIList.Add("TowerRedundantDiagonalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerRedundantDiagonalNetWidthDeduct))
                        newERIList.Add("TowerRedundantDiagonalUFactor=" & Me.geometry.baseStructure(i).TowerRedundantDiagonalUFactor)
                        newERIList.Add("TowerRedundantSubDiagonalBoltGrade=" & Me.geometry.baseStructure(i).TowerRedundantSubDiagonalBoltGrade)
                        newERIList.Add("TowerRedundantSubDiagonalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerRedundantSubDiagonalBoltSize))
                        newERIList.Add("TowerRedundantSubDiagonalNumBolts=" & Me.geometry.baseStructure(i).TowerRedundantSubDiagonalNumBolts)
                        newERIList.Add("TowerRedundantSubDiagonalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerRedundantSubDiagonalBoltEdgeDistance))
                        newERIList.Add("TowerRedundantSubDiagonalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerRedundantSubDiagonalGageG1Distance))
                        newERIList.Add("TowerRedundantSubDiagonalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerRedundantSubDiagonalNetWidthDeduct))
                        newERIList.Add("TowerRedundantSubDiagonalUFactor=" & Me.geometry.baseStructure(i).TowerRedundantSubDiagonalUFactor)
                        newERIList.Add("TowerRedundantSubHorizontalBoltGrade=" & Me.geometry.baseStructure(i).TowerRedundantSubHorizontalBoltGrade)
                        newERIList.Add("TowerRedundantSubHorizontalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerRedundantSubHorizontalBoltSize))
                        newERIList.Add("TowerRedundantSubHorizontalNumBolts=" & Me.geometry.baseStructure(i).TowerRedundantSubHorizontalNumBolts)
                        newERIList.Add("TowerRedundantSubHorizontalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerRedundantSubHorizontalBoltEdgeDistance))
                        newERIList.Add("TowerRedundantSubHorizontalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerRedundantSubHorizontalGageG1Distance))
                        newERIList.Add("TowerRedundantSubHorizontalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerRedundantSubHorizontalNetWidthDeduct))
                        newERIList.Add("TowerRedundantSubHorizontalUFactor=" & Me.geometry.baseStructure(i).TowerRedundantSubHorizontalUFactor)
                        newERIList.Add("TowerRedundantVerticalBoltGrade=" & Me.geometry.baseStructure(i).TowerRedundantVerticalBoltGrade)
                        newERIList.Add("TowerRedundantVerticalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerRedundantVerticalBoltSize))
                        newERIList.Add("TowerRedundantVerticalNumBolts=" & Me.geometry.baseStructure(i).TowerRedundantVerticalNumBolts)
                        newERIList.Add("TowerRedundantVerticalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerRedundantVerticalBoltEdgeDistance))
                        newERIList.Add("TowerRedundantVerticalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerRedundantVerticalGageG1Distance))
                        newERIList.Add("TowerRedundantVerticalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerRedundantVerticalNetWidthDeduct))
                        newERIList.Add("TowerRedundantVerticalUFactor=" & Me.geometry.baseStructure(i).TowerRedundantVerticalUFactor)
                        newERIList.Add("TowerRedundantHipBoltGrade=" & Me.geometry.baseStructure(i).TowerRedundantHipBoltGrade)
                        newERIList.Add("TowerRedundantHipBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerRedundantHipBoltSize))
                        newERIList.Add("TowerRedundantHipNumBolts=" & Me.geometry.baseStructure(i).TowerRedundantHipNumBolts)
                        newERIList.Add("TowerRedundantHipBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerRedundantHipBoltEdgeDistance))
                        newERIList.Add("TowerRedundantHipGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerRedundantHipGageG1Distance))
                        newERIList.Add("TowerRedundantHipNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerRedundantHipNetWidthDeduct))
                        newERIList.Add("TowerRedundantHipUFactor=" & Me.geometry.baseStructure(i).TowerRedundantHipUFactor)
                        newERIList.Add("TowerRedundantHipDiagonalBoltGrade=" & Me.geometry.baseStructure(i).TowerRedundantHipDiagonalBoltGrade)
                        newERIList.Add("TowerRedundantHipDiagonalBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerRedundantHipDiagonalBoltSize))
                        newERIList.Add("TowerRedundantHipDiagonalNumBolts=" & Me.geometry.baseStructure(i).TowerRedundantHipDiagonalNumBolts)
                        newERIList.Add("TowerRedundantHipDiagonalBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerRedundantHipDiagonalBoltEdgeDistance))
                        newERIList.Add("TowerRedundantHipDiagonalGageG1Distance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerRedundantHipDiagonalGageG1Distance))
                        newERIList.Add("TowerRedundantHipDiagonalNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerRedundantHipDiagonalNetWidthDeduct))
                        newERIList.Add("TowerRedundantHipDiagonalUFactor=" & Me.geometry.baseStructure(i).TowerRedundantHipDiagonalUFactor)
                        newERIList.Add("TowerDiagonalOutOfPlaneRestraint=" & trueFalseYesNo(Me.geometry.baseStructure(i).TowerDiagonalOutOfPlaneRestraint))
                        newERIList.Add("TowerTopGirtOutOfPlaneRestraint=" & trueFalseYesNo(Me.geometry.baseStructure(i).TowerTopGirtOutOfPlaneRestraint))
                        newERIList.Add("TowerBottomGirtOutOfPlaneRestraint=" & trueFalseYesNo(Me.geometry.baseStructure(i).TowerBottomGirtOutOfPlaneRestraint))
                        newERIList.Add("TowerMidGirtOutOfPlaneRestraint=" & trueFalseYesNo(Me.geometry.baseStructure(i).TowerMidGirtOutOfPlaneRestraint))
                        newERIList.Add("TowerHorizontalOutOfPlaneRestraint=" & trueFalseYesNo(Me.geometry.baseStructure(i).TowerHorizontalOutOfPlaneRestraint))
                        newERIList.Add("TowerSecondaryHorizontalOutOfPlaneRestraint=" & trueFalseYesNo(Me.geometry.baseStructure(i).TowerSecondaryHorizontalOutOfPlaneRestraint))
                        newERIList.Add("TowerUniqueFlag=" & Me.geometry.baseStructure(i).TowerUniqueFlag)
                        newERIList.Add("TowerDiagOffsetNEY=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerDiagOffsetNEY))
                        newERIList.Add("TowerDiagOffsetNEX=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerDiagOffsetNEX))
                        newERIList.Add("TowerDiagOffsetPEY=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerDiagOffsetPEY))
                        newERIList.Add("TowerDiagOffsetPEX=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerDiagOffsetPEX))
                        newERIList.Add("TowerKbraceOffsetNEY=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerKbraceOffsetNEY))
                        newERIList.Add("TowerKbraceOffsetNEX=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerKbraceOffsetNEX))
                        newERIList.Add("TowerKbraceOffsetPEY=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerKbraceOffsetPEY))
                        newERIList.Add("TowerKbraceOffsetPEX=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.baseStructure(i).TowerKbraceOffsetPEX))

                    Next i
                Case line(0).Equals("NumGuyRecs")
                    newERIList.Add(line(0) & "=" & Me.geometry.guyWires.Count)
                    For i = 0 To Me.geometry.guyWires.Count - 1
                        newERIList.Add("GuyRec=" & Me.geometry.guyWires(i).GuyRec)
                        newERIList.Add("GuyHeight=" & Me.settings.USUnits.Coordinate.convertToERIUnits(Me.geometry.guyWires(i).GuyHeight))
                        newERIList.Add("GuyAutoCalcKSingleAngle=" & trueFalseYesNo(Me.geometry.guyWires(i).GuyAutoCalcKSingleAngle))
                        newERIList.Add("GuyAutoCalcKSolidRound=" & trueFalseYesNo(Me.geometry.guyWires(i).GuyAutoCalcKSolidRound))
                        newERIList.Add("GuyMount=" & Me.geometry.guyWires(i).GuyMount)
                        newERIList.Add("TorqueArmStyle=" & Me.geometry.guyWires(i).TorqueArmStyle)
                        newERIList.Add("GuyRadius=" & Me.settings.USUnits.Length.convertToERIUnits(Me.geometry.guyWires(i).GuyRadius))
                        newERIList.Add("GuyRadius120=" & Me.settings.USUnits.Length.convertToERIUnits(Me.geometry.guyWires(i).GuyRadius120))
                        newERIList.Add("GuyRadius240=" & Me.settings.USUnits.Length.convertToERIUnits(Me.geometry.guyWires(i).GuyRadius240))
                        newERIList.Add("GuyRadius360=" & Me.settings.USUnits.Length.convertToERIUnits(Me.geometry.guyWires(i).GuyRadius360))
                        newERIList.Add("TorqueArmRadius=" & Me.settings.USUnits.Length.convertToERIUnits(Me.geometry.guyWires(i).TorqueArmRadius))
                        newERIList.Add("TorqueArmLegAngle=" & Me.settings.USUnits.Rotation.convertToERIUnits(Me.geometry.guyWires(i).TorqueArmLegAngle))
                        newERIList.Add("Azimuth0Adjustment=" & Me.settings.USUnits.Rotation.convertToERIUnits(Me.geometry.guyWires(i).Azimuth0Adjustment))
                        newERIList.Add("Azimuth120Adjustment=" & Me.settings.USUnits.Rotation.convertToERIUnits(Me.geometry.guyWires(i).Azimuth120Adjustment))
                        newERIList.Add("Azimuth240Adjustment=" & Me.settings.USUnits.Rotation.convertToERIUnits(Me.geometry.guyWires(i).Azimuth240Adjustment))
                        newERIList.Add("Azimuth360Adjustment=" & Me.settings.USUnits.Rotation.convertToERIUnits(Me.geometry.guyWires(i).Azimuth360Adjustment))
                        newERIList.Add("Anchor0Elevation=" & Me.settings.USUnits.Coordinate.convertToERIUnits(Me.geometry.guyWires(i).Anchor0Elevation))
                        newERIList.Add("Anchor120Elevation=" & Me.settings.USUnits.Coordinate.convertToERIUnits(Me.geometry.guyWires(i).Anchor120Elevation))
                        newERIList.Add("Anchor240Elevation=" & Me.settings.USUnits.Coordinate.convertToERIUnits(Me.geometry.guyWires(i).Anchor240Elevation))
                        newERIList.Add("Anchor360Elevation=" & Me.settings.USUnits.Coordinate.convertToERIUnits(Me.geometry.guyWires(i).Anchor360Elevation))
                        newERIList.Add("GuySize=" & Me.geometry.guyWires(i).GuySize)
                        newERIList.Add("Guy120Size=" & Me.geometry.guyWires(i).Guy120Size)
                        newERIList.Add("Guy240Size=" & Me.geometry.guyWires(i).Guy240Size)
                        newERIList.Add("Guy360Size=" & Me.geometry.guyWires(i).Guy360Size)
                        newERIList.Add("GuyGrade=" & Me.geometry.guyWires(i).GuyGrade)
                        newERIList.Add("TorqueArmSize=" & Me.geometry.guyWires(i).TorqueArmSize)
                        newERIList.Add("TorqueArmSizeBot=" & Me.geometry.guyWires(i).TorqueArmSizeBot)
                        newERIList.Add("TorqueArmType=" & Me.geometry.guyWires(i).TorqueArmType)
                        newERIList.Add("TorqueArmGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(Me.geometry.guyWires(i).TorqueArmGrade))
                        newERIList.Add("TorqueArmMatlGrade=" & Me.geometry.guyWires(i).TorqueArmMatlGrade)
                        newERIList.Add("TorqueArmKFactor=" & Me.geometry.guyWires(i).TorqueArmKFactor)
                        newERIList.Add("TorqueArmKFactorY=" & Me.geometry.guyWires(i).TorqueArmKFactorY)
                        newERIList.Add("GuyPullOffKFactorX=" & Me.geometry.guyWires(i).GuyPullOffKFactorX)
                        newERIList.Add("GuyPullOffKFactorY=" & Me.geometry.guyWires(i).GuyPullOffKFactorY)
                        newERIList.Add("GuyDiagKFactorX=" & Me.geometry.guyWires(i).GuyDiagKFactorX)
                        newERIList.Add("GuyDiagKFactorY=" & Me.geometry.guyWires(i).GuyDiagKFactorY)
                        newERIList.Add("GuyAutoCalc=" & trueFalseYesNo(Me.geometry.guyWires(i).GuyAutoCalc))
                        newERIList.Add("GuyAllGuysSame=" & trueFalseYesNo(Me.geometry.guyWires(i).GuyAllGuysSame))
                        newERIList.Add("GuyAllGuysAnchorSame=" & trueFalseYesNo(Me.geometry.guyWires(i).GuyAllGuysAnchorSame))
                        newERIList.Add("GuyIsStrapping=" & trueFalseYesNo(Me.geometry.guyWires(i).GuyIsStrapping))
                        newERIList.Add("GuyPullOffSize=" & Me.geometry.guyWires(i).GuyPullOffSize)
                        newERIList.Add("GuyPullOffSizeBot=" & Me.geometry.guyWires(i).GuyPullOffSizeBot)
                        newERIList.Add("GuyPullOffType=" & Me.geometry.guyWires(i).GuyPullOffType)
                        newERIList.Add("GuyPullOffGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(Me.geometry.guyWires(i).GuyPullOffGrade))
                        newERIList.Add("GuyPullOffMatlGrade=" & Me.geometry.guyWires(i).GuyPullOffMatlGrade)
                        newERIList.Add("GuyUpperDiagSize=" & Me.geometry.guyWires(i).GuyUpperDiagSize)
                        newERIList.Add("GuyLowerDiagSize=" & Me.geometry.guyWires(i).GuyLowerDiagSize)
                        newERIList.Add("GuyDiagType=" & Me.geometry.guyWires(i).GuyDiagType)
                        newERIList.Add("GuyDiagGrade=" & Me.settings.USUnits.Strength.convertToERIUnits(Me.geometry.guyWires(i).GuyDiagGrade))
                        newERIList.Add("GuyDiagMatlGrade=" & Me.geometry.guyWires(i).GuyDiagMatlGrade)
                        newERIList.Add("GuyDiagNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.guyWires(i).GuyDiagNetWidthDeduct))
                        newERIList.Add("GuyDiagUFactor=" & Me.geometry.guyWires(i).GuyDiagUFactor)
                        newERIList.Add("GuyDiagNumBolts=" & Me.geometry.guyWires(i).GuyDiagNumBolts)
                        newERIList.Add("GuyDiagonalOutOfPlaneRestraint=" & trueFalseYesNo(Me.geometry.guyWires(i).GuyDiagonalOutOfPlaneRestraint))
                        newERIList.Add("GuyDiagBoltGrade=" & Me.geometry.guyWires(i).GuyDiagBoltGrade)
                        newERIList.Add("GuyDiagBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.guyWires(i).GuyDiagBoltSize))
                        newERIList.Add("GuyDiagBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.guyWires(i).GuyDiagBoltEdgeDistance))
                        newERIList.Add("GuyDiagBoltGageDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.guyWires(i).GuyDiagBoltGageDistance))
                        newERIList.Add("GuyPullOffNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.guyWires(i).GuyPullOffNetWidthDeduct))
                        newERIList.Add("GuyPullOffUFactor=" & Me.geometry.guyWires(i).GuyPullOffUFactor)
                        newERIList.Add("GuyPullOffNumBolts=" & Me.geometry.guyWires(i).GuyPullOffNumBolts)
                        newERIList.Add("GuyPullOffOutOfPlaneRestraint=" & trueFalseYesNo(Me.geometry.guyWires(i).GuyPullOffOutOfPlaneRestraint))
                        newERIList.Add("GuyPullOffBoltGrade=" & Me.geometry.guyWires(i).GuyPullOffBoltGrade)
                        newERIList.Add("GuyPullOffBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.guyWires(i).GuyPullOffBoltSize))
                        newERIList.Add("GuyPullOffBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.guyWires(i).GuyPullOffBoltEdgeDistance))
                        newERIList.Add("GuyPullOffBoltGageDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.guyWires(i).GuyPullOffBoltGageDistance))
                        newERIList.Add("GuyTorqueArmNetWidthDeduct=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.guyWires(i).GuyTorqueArmNetWidthDeduct))
                        newERIList.Add("GuyTorqueArmUFactor=" & Me.geometry.guyWires(i).GuyTorqueArmUFactor)
                        newERIList.Add("GuyTorqueArmNumBolts=" & Me.geometry.guyWires(i).GuyTorqueArmNumBolts)
                        newERIList.Add("GuyTorqueArmOutOfPlaneRestraint=" & trueFalseYesNo(Me.geometry.guyWires(i).GuyTorqueArmOutOfPlaneRestraint))
                        newERIList.Add("GuyTorqueArmBoltGrade=" & Me.geometry.guyWires(i).GuyTorqueArmBoltGrade)
                        newERIList.Add("GuyTorqueArmBoltSize=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.guyWires(i).GuyTorqueArmBoltSize))
                        newERIList.Add("GuyTorqueArmBoltEdgeDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.guyWires(i).GuyTorqueArmBoltEdgeDistance))
                        newERIList.Add("GuyTorqueArmBoltGageDistance=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.guyWires(i).GuyTorqueArmBoltGageDistance))
                        newERIList.Add("GuyPerCentTension=" & Me.geometry.guyWires(i).GuyPerCentTension)
                        newERIList.Add("GuyPerCentTension120=" & Me.geometry.guyWires(i).GuyPerCentTension120)
                        newERIList.Add("GuyPerCentTension240=" & Me.geometry.guyWires(i).GuyPerCentTension240)
                        newERIList.Add("GuyPerCentTension360=" & Me.geometry.guyWires(i).GuyPerCentTension360)
                        newERIList.Add("GuyEffFactor=" & Me.geometry.guyWires(i).GuyEffFactor)
                        newERIList.Add("GuyEffFactor120=" & Me.geometry.guyWires(i).GuyEffFactor120)
                        newERIList.Add("GuyEffFactor240=" & Me.geometry.guyWires(i).GuyEffFactor240)
                        newERIList.Add("GuyEffFactor360=" & Me.geometry.guyWires(i).GuyEffFactor360)
                        newERIList.Add("GuyNumInsulators=" & Me.geometry.guyWires(i).GuyNumInsulators)
                        newERIList.Add("GuyInsulatorLength=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.guyWires(i).GuyInsulatorLength))
                        newERIList.Add("GuyInsulatorDia=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.geometry.guyWires(i).GuyInsulatorDia))
                        newERIList.Add("GuyInsulatorWt=" & Me.settings.USUnits.Force.convertToERIUnits(Me.geometry.guyWires(i).GuyInsulatorWt))
                    Next i
                Case line(0).Equals("NumFeedLineRecs")
                    newERIList.Add(line(0) & "=" & Me.feedLines.Count)
                    For i = 0 To Me.feedLines.Count - 1
                        newERIList.Add("FeedLineRec=" & Me.feedLines(i).FeedLineRec)
                        newERIList.Add("FeedLineEnabled=" & trueFalseYesNo(Me.feedLines(i).FeedLineEnabled))
                        newERIList.Add("FeedLineDatabase=" & Me.feedLines(i).FeedLineDatabase)
                        newERIList.Add("FeedLineDescription=" & Me.feedLines(i).FeedLineDescription)
                        newERIList.Add("FeedLineClassificationCategory=" & Me.feedLines(i).FeedLineClassificationCategory)
                        newERIList.Add("FeedLineNote=" & Me.feedLines(i).FeedLineNote)
                        newERIList.Add("FeedLineNum=" & Me.feedLines(i).FeedLineNum)
                        newERIList.Add("FeedLineUseShielding=" & trueFalseYesNo(Me.feedLines(i).FeedLineUseShielding))
                        newERIList.Add("ExcludeFeedLineFromTorque=" & trueFalseYesNo(Me.feedLines(i).ExcludeFeedLineFromTorque))
                        newERIList.Add("FeedLineNumPerRow=" & Me.feedLines(i).FeedLineNumPerRow)
                        newERIList.Add("FeedLineFace=" & Me.feedLines(i).FeedLineFace)
                        newERIList.Add("FeedLineComponentType=" & Me.feedLines(i).FeedLineComponentType)
                        newERIList.Add("FeedLineGroupTreatmentType=" & Me.feedLines(i).FeedLineGroupTreatmentType)
                        newERIList.Add("FeedLineRoundClusterDia=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.feedLines(i).FeedLineRoundClusterDia))
                        newERIList.Add("FeedLineWidth=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.feedLines(i).FeedLineWidth))
                        newERIList.Add("FeedLinePerimeter=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.feedLines(i).FeedLinePerimeter))
                        newERIList.Add("FlatAttachmentEffectiveWidthRatio=" & Me.feedLines(i).FlatAttachmentEffectiveWidthRatio)
                        newERIList.Add("AutoCalcFlatAttachmentEffectiveWidthRatio=" & trueFalseYesNo(Me.feedLines(i).AutoCalcFlatAttachmentEffectiveWidthRatio))
                        newERIList.Add("FeedLineShieldingFactorKaNoIce=" & Me.feedLines(i).FeedLineShieldingFactorKaNoIce)
                        newERIList.Add("FeedLineShieldingFactorKaIce=" & Me.feedLines(i).FeedLineShieldingFactorKaIce)
                        newERIList.Add("FeedLineAutoCalcKa=" & trueFalseYesNo(Me.feedLines(i).FeedLineAutoCalcKa))
                        newERIList.Add("FeedLineCaAaNoIce=" & Me.settings.USUnits.Length.convertAreaToERIUnits(Me.feedLines(i).FeedLineCaAaNoIce))
                        newERIList.Add("FeedLineCaAaIce=" & Me.settings.USUnits.Length.convertAreaToERIUnits(Me.feedLines(i).FeedLineCaAaIce))
                        newERIList.Add("FeedLineCaAaIce_1=" & Me.settings.USUnits.Length.convertAreaToERIUnits(Me.feedLines(i).FeedLineCaAaIce_1))
                        newERIList.Add("FeedLineCaAaIce_2=" & Me.settings.USUnits.Length.convertAreaToERIUnits(Me.feedLines(i).FeedLineCaAaIce_2))
                        newERIList.Add("FeedLineCaAaIce_4=" & Me.settings.USUnits.Length.convertAreaToERIUnits(Me.feedLines(i).FeedLineCaAaIce_4))
                        newERIList.Add("FeedLineWtNoIce=" & Me.settings.USUnits.Load.convertToERIUnits(Me.feedLines(i).FeedLineWtNoIce))
                        newERIList.Add("FeedLineWtIce=" & Me.settings.USUnits.Load.convertToERIUnits(Me.feedLines(i).FeedLineWtIce))
                        newERIList.Add("FeedLineWtIce_1=" & Me.settings.USUnits.Load.convertToERIUnits(Me.feedLines(i).FeedLineWtIce_1))
                        newERIList.Add("FeedLineWtIce_2=" & Me.settings.USUnits.Load.convertToERIUnits(Me.feedLines(i).FeedLineWtIce_2))
                        newERIList.Add("FeedLineWtIce_4=" & Me.settings.USUnits.Load.convertToERIUnits(Me.feedLines(i).FeedLineWtIce_4))
                        newERIList.Add("FeedLineFaceOffset=" & Me.feedLines(i).FeedLineFaceOffset)
                        newERIList.Add("FeedLineOffsetFrac=" & Me.feedLines(i).FeedLineOffsetFrac)
                        newERIList.Add("FeedLinePerimeterOffsetStartFrac=" & Me.feedLines(i).FeedLinePerimeterOffsetStartFrac)
                        newERIList.Add("FeedLinePerimeterOffsetEndFrac=" & Me.feedLines(i).FeedLinePerimeterOffsetEndFrac)
                        newERIList.Add("FeedLineStartHt=" & Me.settings.USUnits.Coordinate.convertToERIUnits(Me.feedLines(i).FeedLineStartHt))
                        newERIList.Add("FeedLineEndHt=" & Me.settings.USUnits.Coordinate.convertToERIUnits(Me.feedLines(i).FeedLineEndHt))
                        newERIList.Add("FeedLineClearSpacing=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.feedLines(i).FeedLineClearSpacing))
                        newERIList.Add("FeedLineRowClearSpacing=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.feedLines(i).FeedLineRowClearSpacing))
                    Next i
                Case line(0).Equals("NumTowerLoadRecs")
                    newERIList.Add(line(0) & "=" & Me.discreteLoads.Count)
                    For i = 0 To Me.discreteLoads.Count - 1
                        newERIList.Add("TowerLoadRec=" & Me.discreteLoads(i).TowerLoadRec)
                        newERIList.Add("TowerLoadEnabled=" & trueFalseYesNo(Me.discreteLoads(i).TowerLoadEnabled))
                        newERIList.Add("TowerLoadDatabase=" & Me.discreteLoads(i).TowerLoadDatabase)
                        newERIList.Add("TowerLoadDescription=" & Me.discreteLoads(i).TowerLoadDescription)
                        newERIList.Add("TowerLoadType=" & Me.discreteLoads(i).TowerLoadType)
                        newERIList.Add("TowerLoadClassificationCategory=" & Me.discreteLoads(i).TowerLoadClassificationCategory)
                        newERIList.Add("TowerLoadNote=" & Me.discreteLoads(i).TowerLoadNote)
                        newERIList.Add("TowerLoadNum=" & Me.discreteLoads(i).TowerLoadNum)
                        newERIList.Add("TowerLoadFace=" & Me.discreteLoads(i).TowerLoadFace)
                        newERIList.Add("TowerOffsetType=" & Me.discreteLoads(i).TowerOffsetType)
                        newERIList.Add("TowerOffsetDist=" & Me.settings.USUnits.Coordinate.convertToERIUnits(Me.discreteLoads(i).TowerOffsetDist))
                        newERIList.Add("TowerVertOffset=" & Me.settings.USUnits.Coordinate.convertToERIUnits(Me.discreteLoads(i).TowerVertOffset))
                        newERIList.Add("TowerLateralOffset=" & Me.settings.USUnits.Coordinate.convertToERIUnits(Me.discreteLoads(i).TowerLateralOffset))
                        newERIList.Add("TowerAzimuthAdjustment=" & Me.settings.USUnits.Rotation.convertToERIUnits(Me.discreteLoads(i).TowerAzimuthAdjustment))
                        newERIList.Add("TowerAppurtSymbol=" & Me.discreteLoads(i).TowerAppurtSymbol)
                        newERIList.Add("TowerLoadShieldingFactorKaNoIce=" & Me.discreteLoads(i).TowerLoadShieldingFactorKaNoIce)
                        newERIList.Add("TowerLoadShieldingFactorKaIce=" & Me.discreteLoads(i).TowerLoadShieldingFactorKaIce)
                        newERIList.Add("TowerLoadAutoCalcKa=" & trueFalseYesNo(Me.discreteLoads(i).TowerLoadAutoCalcKa))
                        newERIList.Add("TowerLoadCaAaNoIce=" & Me.settings.USUnits.Length.convertAreaToERIUnits(Me.discreteLoads(i).TowerLoadCaAaNoIce))
                        newERIList.Add("TowerLoadCaAaIce=" & Me.settings.USUnits.Length.convertAreaToERIUnits(Me.discreteLoads(i).TowerLoadCaAaIce))
                        newERIList.Add("TowerLoadCaAaIce_1=" & Me.settings.USUnits.Length.convertAreaToERIUnits(Me.discreteLoads(i).TowerLoadCaAaIce_1))
                        newERIList.Add("TowerLoadCaAaIce_2=" & Me.settings.USUnits.Length.convertAreaToERIUnits(Me.discreteLoads(i).TowerLoadCaAaIce_2))
                        newERIList.Add("TowerLoadCaAaIce_4=" & Me.settings.USUnits.Length.convertAreaToERIUnits(Me.discreteLoads(i).TowerLoadCaAaIce_4))
                        newERIList.Add("TowerLoadCaAaNoIce_Side=" & Me.settings.USUnits.Length.convertAreaToERIUnits(Me.discreteLoads(i).TowerLoadCaAaNoIce_Side))
                        newERIList.Add("TowerLoadCaAaIce_Side=" & Me.settings.USUnits.Length.convertAreaToERIUnits(Me.discreteLoads(i).TowerLoadCaAaIce_Side))
                        newERIList.Add("TowerLoadCaAaIce_Side_1=" & Me.settings.USUnits.Length.convertAreaToERIUnits(Me.discreteLoads(i).TowerLoadCaAaIce_Side_1))
                        newERIList.Add("TowerLoadCaAaIce_Side_2=" & Me.settings.USUnits.Length.convertAreaToERIUnits(Me.discreteLoads(i).TowerLoadCaAaIce_Side_2))
                        newERIList.Add("TowerLoadCaAaIce_Side_4=" & Me.settings.USUnits.Length.convertAreaToERIUnits(Me.discreteLoads(i).TowerLoadCaAaIce_Side_4))
                        newERIList.Add("TowerLoadWtNoIce=" & Me.settings.USUnits.Force.convertToERIUnits(Me.discreteLoads(i).TowerLoadWtNoIce))
                        newERIList.Add("TowerLoadWtIce=" & Me.settings.USUnits.Force.convertToERIUnits(Me.discreteLoads(i).TowerLoadWtIce))
                        newERIList.Add("TowerLoadWtIce_1=" & Me.settings.USUnits.Force.convertToERIUnits(Me.discreteLoads(i).TowerLoadWtIce_1))
                        newERIList.Add("TowerLoadWtIce_2=" & Me.settings.USUnits.Force.convertToERIUnits(Me.discreteLoads(i).TowerLoadWtIce_2))
                        newERIList.Add("TowerLoadWtIce_4=" & Me.settings.USUnits.Force.convertToERIUnits(Me.discreteLoads(i).TowerLoadWtIce_4))
                        newERIList.Add("TowerLoadStartHt=" & Me.settings.USUnits.Coordinate.convertToERIUnits(Me.discreteLoads(i).TowerLoadStartHt))
                        newERIList.Add("TowerLoadEndHt=" & Me.settings.USUnits.Coordinate.convertToERIUnits(Me.discreteLoads(i).TowerLoadEndHt))
                    Next i
                Case line(0).Equals("NumDishRecs")
                    newERIList.Add(line(0) & "=" & Me.dishes.Count)
                    For i = 0 To Me.dishes.Count - 1
                        newERIList.Add("DishRec=" & Me.dishes(i).DishRec)
                        newERIList.Add("DishEnabled=" & trueFalseYesNo(Me.dishes(i).DishEnabled))
                        newERIList.Add("DishDatabase=" & Me.dishes(i).DishDatabase)
                        newERIList.Add("DishDescription=" & Me.dishes(i).DishDescription)
                        newERIList.Add("DishClassificationCategory=" & Me.dishes(i).DishClassificationCategory)
                        newERIList.Add("DishNote=" & Me.dishes(i).DishNote)
                        newERIList.Add("DishNum=" & Me.dishes(i).DishNum)
                        newERIList.Add("DishFace=" & Me.dishes(i).DishFace)
                        newERIList.Add("DishType=" & Me.dishes(i).DishType)
                        newERIList.Add("DishOffsetType=" & Me.dishes(i).DishOffsetType)
                        newERIList.Add("DishVertOffset=" & Me.settings.USUnits.Coordinate.convertToERIUnits(Me.dishes(i).DishVertOffset))
                        newERIList.Add("DishLateralOffset=" & Me.settings.USUnits.Coordinate.convertToERIUnits(Me.dishes(i).DishLateralOffset))
                        newERIList.Add("DishOffsetDist=" & Me.settings.USUnits.Properties.convertToERIUnits(Me.dishes(i).DishOffsetDist))
                        newERIList.Add("DishArea=" & Me.settings.USUnits.Length.convertAreaToERIUnits(Me.dishes(i).DishArea))
                        newERIList.Add("DishAreaIce=" & Me.settings.USUnits.Length.convertAreaToERIUnits(Me.dishes(i).DishAreaIce))
                        newERIList.Add("DishAreaIce_1=" & Me.settings.USUnits.Length.convertAreaToERIUnits(Me.dishes(i).DishAreaIce_1))
                        newERIList.Add("DishAreaIce_2=" & Me.settings.USUnits.Length.convertAreaToERIUnits(Me.dishes(i).DishAreaIce_2))
                        newERIList.Add("DishAreaIce_4=" & Me.settings.USUnits.Length.convertAreaToERIUnits(Me.dishes(i).DishAreaIce_4))
                        newERIList.Add("DishDiameter=" & Me.settings.USUnits.Length.convertToERIUnits(Me.dishes(i).DishDiameter))
                        newERIList.Add("DishWtNoIce=" & Me.settings.USUnits.Force.convertToERIUnits(Me.dishes(i).DishWtNoIce))
                        newERIList.Add("DishWtIce=" & Me.settings.USUnits.Force.convertToERIUnits(Me.dishes(i).DishWtIce))
                        newERIList.Add("DishWtIce_1=" & Me.settings.USUnits.Force.convertToERIUnits(Me.dishes(i).DishWtIce_1))
                        newERIList.Add("DishWtIce_2=" & Me.settings.USUnits.Force.convertToERIUnits(Me.dishes(i).DishWtIce_2))
                        newERIList.Add("DishWtIce_4=" & Me.settings.USUnits.Force.convertToERIUnits(Me.dishes(i).DishWtIce_4))
                        newERIList.Add("DishStartHt=" & Me.settings.USUnits.Coordinate.convertToERIUnits(Me.dishes(i).DishStartHt))
                        newERIList.Add("DishAzimuthAdjustment=" & Me.settings.USUnits.Rotation.convertToERIUnits(Me.dishes(i).DishAzimuthAdjustment))
                        newERIList.Add("DishBeamWidth=" & Me.settings.USUnits.Rotation.convertToERIUnits(Me.dishes(i).DishBeamWidth))
                    Next i
                Case line(0).Equals("NumUserForceRecs")
                    newERIList.Add(line(0) & "=" & Me.userForces.Count)
                    For i = 0 To Me.userForces.Count - 1
                        newERIList.Add("UserForceRec=" & Me.userForces(i).UserForceRec)
                        newERIList.Add("UserForceEnabled=" & trueFalseYesNo(Me.userForces(i).UserForceEnabled))
                        newERIList.Add("UserForceDescription=" & Me.userForces(i).UserForceDescription)
                        newERIList.Add("UserForceStartHt=" & Me.settings.USUnits.Coordinate.convertToERIUnits(Me.userForces(i).UserForceStartHt))
                        newERIList.Add("UserForceOffset=" & Me.settings.USUnits.Coordinate.convertToERIUnits(Me.userForces(i).UserForceOffset))
                        newERIList.Add("UserForceAzimuth=" & Me.settings.USUnits.Rotation.convertToERIUnits(Me.userForces(i).UserForceAzimuth))
                        newERIList.Add("UserForceFxNoIce=" & Me.settings.USUnits.Force.convertToERIUnits(Me.userForces(i).UserForceFxNoIce))
                        newERIList.Add("UserForceFzNoIce=" & Me.settings.USUnits.Force.convertToERIUnits(Me.userForces(i).UserForceFzNoIce))
                        newERIList.Add("UserForceAxialNoIce=" & Me.settings.USUnits.Force.convertToERIUnits(Me.userForces(i).UserForceAxialNoIce))
                        newERIList.Add("UserForceShearNoIce=" & Me.settings.USUnits.Force.convertToERIUnits(Me.userForces(i).UserForceShearNoIce))
                        newERIList.Add("UserForceCaAcNoIce=" & Me.settings.USUnits.Length.convertToERIUnits(Me.userForces(i).UserForceCaAcNoIce))
                        newERIList.Add("UserForceFxIce=" & Me.settings.USUnits.Force.convertToERIUnits(Me.userForces(i).UserForceFxIce))
                        newERIList.Add("UserForceFzIce=" & Me.settings.USUnits.Force.convertToERIUnits(Me.userForces(i).UserForceFzIce))
                        newERIList.Add("UserForceAxialIce=" & Me.settings.USUnits.Force.convertToERIUnits(Me.userForces(i).UserForceAxialIce))
                        newERIList.Add("UserForceShearIce=" & Me.settings.USUnits.Force.convertToERIUnits(Me.userForces(i).UserForceShearIce))
                        newERIList.Add("UserForceCaAcIce=" & Me.settings.USUnits.Length.convertToERIUnits(Me.userForces(i).UserForceCaAcIce))
                        newERIList.Add("UserForceFxService=" & Me.settings.USUnits.Force.convertToERIUnits(Me.userForces(i).UserForceFxService))
                        newERIList.Add("UserForceFzService=" & Me.settings.USUnits.Force.convertToERIUnits(Me.userForces(i).UserForceFzService))
                        newERIList.Add("UserForceAxialService=" & Me.settings.USUnits.Force.convertToERIUnits(Me.userForces(i).UserForceAxialService))
                        newERIList.Add("UserForceShearService=" & Me.settings.USUnits.Force.convertToERIUnits(Me.userForces(i).UserForceShearService))
                        newERIList.Add("UserForceCaAcService=" & Me.settings.USUnits.Length.convertToERIUnits(Me.userForces(i).UserForceCaAcService))
                        newERIList.Add("UserForceEhx=" & Me.settings.USUnits.Force.convertToERIUnits(Me.userForces(i).UserForceEhx))
                        newERIList.Add("UserForceEhz=" & Me.settings.USUnits.Force.convertToERIUnits(Me.userForces(i).UserForceEhz))
                        newERIList.Add("UserForceEv=" & Me.settings.USUnits.Force.convertToERIUnits(Me.userForces(i).UserForceEv))
                        newERIList.Add("UserForceEh=" & Me.settings.USUnits.Force.convertToERIUnits(Me.userForces(i).UserForceEh))
                    Next i
                Case Else
                    If line.Count = 1 Then
                        newERIList.Add(line(0))
                    Else
                        newERIList.Add(line(0) & "=" & line(1))
                    End If
            End Select

        Next

        File.WriteAllLines(FilePath, newERIList, Text.Encoding.ASCII)

    End Sub

    Public Function trueFalseYesNo(input As String) As Boolean

        If input.ToLower = "yes" Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Function trueFalseYesNo(input As Boolean) As String

        If input Then
            Return "Yes"
        Else
            Return "No"
        End If

    End Function

End Class

#Region "Geometry"
Partial Public Class tnxGeometry
    Private prop_TowerType As String
    Private prop_AntennaType As String
    Private prop_OverallHeight As Double
    Private prop_BaseElevation As Double
    Private prop_Lambda As Double
    Private prop_TowerTopFaceWidth As Double
    Private prop_TowerBaseFaceWidth As Double
    Private prop_TowerTaper As String
    Private prop_GuyedMonopoleBaseType As String
    Private prop_TaperHeight As Double
    Private prop_PivotHeight As Double
    Private prop_AutoCalcGH As Boolean
    Private prop_UserGHElev As Double
    Private prop_UseIndexPlate As Boolean
    Private prop_EnterUserDefinedGhValues As Boolean
    Private prop_BaseTowerGhInput As Double
    Private prop_UpperStructureGhInput As Double
    Private prop_EnterUserDefinedCgValues As Boolean
    Private prop_BaseTowerCgInput As Double
    Private prop_UpperStructureCgInput As Double
    Private prop_AntennaFaceWidth As Double
    Private prop_UseTopTakeup As Boolean
    Private prop_ConstantSlope As Boolean
    Private prop_upperStructure As New List(Of tnxAntennaRecord)
    Private prop_baseStructure As New List(Of tnxTowerRecord)
    Private prop_guyWires As New List(Of tnxGuyRecord)

    <Category("TNX Geometry"), Description("Base Tower Type"), DisplayName("TowerType")>
    Public Property TowerType() As String
        Get
            Return Me.prop_TowerType
        End Get
        Set
            Me.prop_TowerType = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("Upper Structure Type"), DisplayName("AntennaType")>
    Public Property AntennaType() As String
        Get
            Return Me.prop_AntennaType
        End Get
        Set
            Me.prop_AntennaType = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description(""), DisplayName("OverallHeight")>
    Public Property OverallHeight() As Double
        Get
            Return Me.prop_OverallHeight
        End Get
        Set
            Me.prop_OverallHeight = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description(""), DisplayName("BaseElevation")>
    Public Property BaseElevation() As Double
        Get
            Return Me.prop_BaseElevation
        End Get
        Set
            Me.prop_BaseElevation = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description(""), DisplayName("Lambda")>
    Public Property Lambda() As Double
        Get
            Return Me.prop_Lambda
        End Get
        Set
            Me.prop_Lambda = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description(""), DisplayName("TowerTopFaceWidth")>
    Public Property TowerTopFaceWidth() As Double
        Get
            Return Me.prop_TowerTopFaceWidth
        End Get
        Set
            Me.prop_TowerTopFaceWidth = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description(""), DisplayName("TowerBaseFaceWidth")>
    Public Property TowerBaseFaceWidth() As Double
        Get
            Return Me.prop_TowerBaseFaceWidth
        End Get
        Set
            Me.prop_TowerBaseFaceWidth = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("Base Type - None, I-Beam, I-Beam Free, Taper, Taper-Free"), DisplayName("TowerTaper")>
    Public Property TowerTaper() As String
        Get
            Return Me.prop_TowerTaper
        End Get
        Set
            Me.prop_TowerTaper = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("Base Type - Fixed Base, Pinned Base (Only active when base tower type is guyed and there are no base tower section in the model)"), DisplayName("GuyedMonopoleBaseType")>
    Public Property GuyedMonopoleBaseType() As String
        Get
            Return Me.prop_GuyedMonopoleBaseType
        End Get
        Set
            Me.prop_GuyedMonopoleBaseType = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("Base Taper Height"), DisplayName("TaperHeight")>
    Public Property TaperHeight() As Double
        Get
            Return Me.prop_TaperHeight
        End Get
        Set
            Me.prop_TaperHeight = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("I-Beam Pivot Dist (replaces base taper height when base type is I-Beam or I-Beam Free)"), DisplayName("PivotHeight")>
    Public Property PivotHeight() As Double
        Get
            Return Me.prop_PivotHeight
        End Get
        Set
            Me.prop_PivotHeight = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description(""), DisplayName("AutoCalcGH")>
    Public Property AutoCalcGH() As Boolean
        Get
            Return Me.prop_AutoCalcGH
        End Get
        Set
            Me.prop_AutoCalcGH = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description(""), DisplayName("UserGHElev")>
    Public Property UserGHElev() As Double
        Get
            Return Me.prop_UserGHElev
        End Get
        Set
            Me.prop_UserGHElev = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("Has Index Plate"), DisplayName("UseIndexPlate")>
    Public Property UseIndexPlate() As Boolean
        Get
            Return Me.prop_UseIndexPlate
        End Get
        Set
            Me.prop_UseIndexPlate = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("Enter Pre-defined Gh values"), DisplayName("EnterUserDefinedGhValues")>
    Public Property EnterUserDefinedGhValues() As Boolean
        Get
            Return Me.prop_EnterUserDefinedGhValues
        End Get
        Set
            Me.prop_EnterUserDefinedGhValues = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("Base Tower - Active when EnterUserDefinedGhValues = Yes"), DisplayName("BaseTowerGhInput")>
    Public Property BaseTowerGhInput() As Double
        Get
            Return Me.prop_BaseTowerGhInput
        End Get
        Set
            Me.prop_BaseTowerGhInput = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("Upper Structure - Active when EnterUserDefinedGhValues = Yes"), DisplayName("UpperStructureGhInput")>
    Public Property UpperStructureGhInput() As Double
        Get
            Return Me.prop_UpperStructureGhInput
        End Get
        Set
            Me.prop_UpperStructureGhInput = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("CSA code only (This controls two inputs in the UI 'Use Default Cg Values' and 'Enter pre-defined Cg Values'. The checked status of the two inputs are always opposite and 'Use Default Cg Values' is opposite of the ERI value."), DisplayName("EnterUserDefinedCgValues")>
    Public Property EnterUserDefinedCgValues() As Boolean
        Get
            Return Me.prop_EnterUserDefinedCgValues
        End Get
        Set
            Me.prop_EnterUserDefinedCgValues = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("CSA code only"), DisplayName("BaseTowerCgInput")>
    Public Property BaseTowerCgInput() As Double
        Get
            Return Me.prop_BaseTowerCgInput
        End Get
        Set
            Me.prop_BaseTowerCgInput = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("CSA code only"), DisplayName("UpperStructureCgInput")>
    Public Property UpperStructureCgInput() As Double
        Get
            Return Me.prop_UpperStructureCgInput
        End Get
        Set
            Me.prop_UpperStructureCgInput = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("Lattice Pole Width - Only applies to lattice upper structures"), DisplayName("AntennaFaceWidth")>
    Public Property AntennaFaceWidth() As Double
        Get
            Return Me.prop_AntennaFaceWidth
        End Get
        Set
            Me.prop_AntennaFaceWidth = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("Top takeup on lambda"), DisplayName("UseTopTakeup")>
    Public Property UseTopTakeup() As Boolean
        Get
            Return Me.prop_UseTopTakeup
        End Get
        Set
            Me.prop_UseTopTakeup = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description("Constant Slope"), DisplayName("ConstantSlope")>
    Public Property ConstantSlope() As Boolean
        Get
            Return Me.prop_ConstantSlope
        End Get
        Set
            Me.prop_ConstantSlope = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description(""), DisplayName("Upper Structure")>
    Public Property upperStructure() As List(Of tnxAntennaRecord)
        Get
            Return Me.prop_upperStructure
        End Get
        Set
            Me.prop_upperStructure = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description(""), DisplayName("Base Structure")>
    Public Property baseStructure() As List(Of tnxTowerRecord)
        Get
            Return Me.prop_baseStructure
        End Get
        Set
            Me.prop_baseStructure = Value
        End Set
    End Property
    <Category("TNX Geometry"), Description(""), DisplayName("Guy Wires")>
    Public Property guyWires() As List(Of tnxGuyRecord)
        Get
            Return Me.prop_guyWires
        End Get
        Set
            Me.prop_guyWires = Value
        End Set
    End Property

End Class

Partial Public Class tnxAntennaRecord
    'upper structure
    Private prop_ID As Integer
    Private prop_tnxID As Integer
    Private prop_AntennaRec As Integer
    Private prop_AntennaBraceType As String
    Private prop_AntennaHeight As Double
    Private prop_AntennaDiagonalSpacing As Double
    Private prop_AntennaDiagonalSpacingEx As Double
    Private prop_AntennaNumSections As Integer
    Private prop_AntennaNumSesctions As Integer
    Private prop_AntennaSectionLength As Double
    Private prop_AntennaLegType As String
    Private prop_AntennaLegSize As String
    Private prop_AntennaLegGrade As Double
    Private prop_AntennaLegMatlGrade As String
    Private prop_AntennaDiagonalGrade As Double
    Private prop_AntennaDiagonalMatlGrade As String
    Private prop_AntennaInnerBracingGrade As Double
    Private prop_AntennaInnerBracingMatlGrade As String
    Private prop_AntennaTopGirtGrade As Double
    Private prop_AntennaTopGirtMatlGrade As String
    Private prop_AntennaBotGirtGrade As Double
    Private prop_AntennaBotGirtMatlGrade As String
    Private prop_AntennaInnerGirtGrade As Double
    Private prop_AntennaInnerGirtMatlGrade As String
    Private prop_AntennaLongHorizontalGrade As Double
    Private prop_AntennaLongHorizontalMatlGrade As String
    Private prop_AntennaShortHorizontalGrade As Double
    Private prop_AntennaShortHorizontalMatlGrade As String
    Private prop_AntennaDiagonalType As String
    Private prop_AntennaDiagonalSize As String
    Private prop_AntennaInnerBracingType As String
    Private prop_AntennaInnerBracingSize As String
    Private prop_AntennaTopGirtType As String
    Private prop_AntennaTopGirtSize As String
    Private prop_AntennaBotGirtType As String
    Private prop_AntennaBotGirtSize As String
    Private prop_AntennaTopGirtOffset As Double
    Private prop_AntennaBotGirtOffset As Double
    Private prop_AntennaHasKBraceEndPanels As Boolean
    Private prop_AntennaHasHorizontals As Boolean
    Private prop_AntennaLongHorizontalType As String
    Private prop_AntennaLongHorizontalSize As String
    Private prop_AntennaShortHorizontalType As String
    Private prop_AntennaShortHorizontalSize As String
    Private prop_AntennaRedundantGrade As Double
    Private prop_AntennaRedundantMatlGrade As String
    Private prop_AntennaRedundantType As String
    Private prop_AntennaRedundantDiagType As String
    Private prop_AntennaRedundantSubDiagonalType As String
    Private prop_AntennaRedundantSubHorizontalType As String
    Private prop_AntennaRedundantVerticalType As String
    Private prop_AntennaRedundantHipType As String
    Private prop_AntennaRedundantHipDiagonalType As String
    Private prop_AntennaRedundantHorizontalSize As String
    Private prop_AntennaRedundantHorizontalSize2 As String
    Private prop_AntennaRedundantHorizontalSize3 As String
    Private prop_AntennaRedundantHorizontalSize4 As String
    Private prop_AntennaRedundantDiagonalSize As String
    Private prop_AntennaRedundantDiagonalSize2 As String
    Private prop_AntennaRedundantDiagonalSize3 As String
    Private prop_AntennaRedundantDiagonalSize4 As String
    Private prop_AntennaRedundantSubHorizontalSize As String
    Private prop_AntennaRedundantSubDiagonalSize As String
    Private prop_AntennaSubDiagLocation As Double
    Private prop_AntennaRedundantVerticalSize As String
    Private prop_AntennaRedundantHipDiagonalSize As String
    Private prop_AntennaRedundantHipDiagonalSize2 As String
    Private prop_AntennaRedundantHipDiagonalSize3 As String
    Private prop_AntennaRedundantHipDiagonalSize4 As String
    Private prop_AntennaRedundantHipSize As String
    Private prop_AntennaRedundantHipSize2 As String
    Private prop_AntennaRedundantHipSize3 As String
    Private prop_AntennaRedundantHipSize4 As String
    Private prop_AntennaNumInnerGirts As Integer
    Private prop_AntennaInnerGirtType As String
    Private prop_AntennaInnerGirtSize As String
    Private prop_AntennaPoleShapeType As String
    Private prop_AntennaPoleSize As String
    Private prop_AntennaPoleGrade As Double
    Private prop_AntennaPoleMatlGrade As String
    Private prop_AntennaPoleSpliceLength As Double
    Private prop_AntennaTaperPoleNumSides As Integer
    Private prop_AntennaTaperPoleTopDiameter As Double
    Private prop_AntennaTaperPoleBotDiameter As Double
    Private prop_AntennaTaperPoleWallThickness As Double
    Private prop_AntennaTaperPoleBendRadius As Double
    Private prop_AntennaTaperPoleGrade As Double
    Private prop_AntennaTaperPoleMatlGrade As String
    Private prop_AntennaSWMult As Double
    Private prop_AntennaWPMult As Double
    Private prop_AntennaAutoCalcKSingleAngle As Double
    Private prop_AntennaAutoCalcKSolidRound As Double
    Private prop_AntennaAfGusset As Double
    Private prop_AntennaTfGusset As Double
    Private prop_AntennaGussetBoltEdgeDistance As Double
    Private prop_AntennaGussetGrade As Double
    Private prop_AntennaGussetMatlGrade As String
    Private prop_AntennaAfMult As Double
    Private prop_AntennaArMult As Double
    Private prop_AntennaFlatIPAPole As Double
    Private prop_AntennaRoundIPAPole As Double
    Private prop_AntennaFlatIPALeg As Double
    Private prop_AntennaRoundIPALeg As Double
    Private prop_AntennaFlatIPAHorizontal As Double
    Private prop_AntennaRoundIPAHorizontal As Double
    Private prop_AntennaFlatIPADiagonal As Double
    Private prop_AntennaRoundIPADiagonal As Double
    Private prop_AntennaCSA_S37_SpeedUpFactor As Double
    Private prop_AntennaKLegs As Double
    Private prop_AntennaKXBracedDiags As Double
    Private prop_AntennaKKBracedDiags As Double
    Private prop_AntennaKZBracedDiags As Double
    Private prop_AntennaKHorzs As Double
    Private prop_AntennaKSecHorzs As Double
    Private prop_AntennaKGirts As Double
    Private prop_AntennaKInners As Double
    Private prop_AntennaKXBracedDiagsY As Double
    Private prop_AntennaKKBracedDiagsY As Double
    Private prop_AntennaKZBracedDiagsY As Double
    Private prop_AntennaKHorzsY As Double
    Private prop_AntennaKSecHorzsY As Double
    Private prop_AntennaKGirtsY As Double
    Private prop_AntennaKInnersY As Double
    Private prop_AntennaKRedHorz As Double
    Private prop_AntennaKRedDiag As Double
    Private prop_AntennaKRedSubDiag As Double
    Private prop_AntennaKRedSubHorz As Double
    Private prop_AntennaKRedVert As Double
    Private prop_AntennaKRedHip As Double
    Private prop_AntennaKRedHipDiag As Double
    Private prop_AntennaKTLX As Double
    Private prop_AntennaKTLZ As Double
    Private prop_AntennaKTLLeg As Double
    Private prop_AntennaInnerKTLX As Double
    Private prop_AntennaInnerKTLZ As Double
    Private prop_AntennaInnerKTLLeg As Double
    Private prop_AntennaStitchBoltLocationHoriz As String
    Private prop_AntennaStitchBoltLocationDiag As String
    Private prop_AntennaStitchSpacing As Double
    Private prop_AntennaStitchSpacingHorz As Double
    Private prop_AntennaStitchSpacingDiag As Double
    Private prop_AntennaStitchSpacingRed As Double
    Private prop_AntennaLegNetWidthDeduct As Double
    Private prop_AntennaLegUFactor As Double
    Private prop_AntennaDiagonalNetWidthDeduct As Double
    Private prop_AntennaTopGirtNetWidthDeduct As Double
    Private prop_AntennaBotGirtNetWidthDeduct As Double
    Private prop_AntennaInnerGirtNetWidthDeduct As Double
    Private prop_AntennaHorizontalNetWidthDeduct As Double
    Private prop_AntennaShortHorizontalNetWidthDeduct As Double
    Private prop_AntennaDiagonalUFactor As Double
    Private prop_AntennaTopGirtUFactor As Double
    Private prop_AntennaBotGirtUFactor As Double
    Private prop_AntennaInnerGirtUFactor As Double
    Private prop_AntennaHorizontalUFactor As Double
    Private prop_AntennaShortHorizontalUFactor As Double
    Private prop_AntennaLegConnType As String
    Private prop_AntennaLegNumBolts As Integer
    Private prop_AntennaDiagonalNumBolts As Integer
    Private prop_AntennaTopGirtNumBolts As Integer
    Private prop_AntennaBotGirtNumBolts As Integer
    Private prop_AntennaInnerGirtNumBolts As Integer
    Private prop_AntennaHorizontalNumBolts As Integer
    Private prop_AntennaShortHorizontalNumBolts As Integer
    Private prop_AntennaLegBoltGrade As String
    Private prop_AntennaLegBoltSize As Double
    Private prop_AntennaDiagonalBoltGrade As String
    Private prop_AntennaDiagonalBoltSize As Double
    Private prop_AntennaTopGirtBoltGrade As String
    Private prop_AntennaTopGirtBoltSize As Double
    Private prop_AntennaBotGirtBoltGrade As String
    Private prop_AntennaBotGirtBoltSize As Double
    Private prop_AntennaInnerGirtBoltGrade As String
    Private prop_AntennaInnerGirtBoltSize As Double
    Private prop_AntennaHorizontalBoltGrade As String
    Private prop_AntennaHorizontalBoltSize As Double
    Private prop_AntennaShortHorizontalBoltGrade As String
    Private prop_AntennaShortHorizontalBoltSize As Double
    Private prop_AntennaLegBoltEdgeDistance As Double
    Private prop_AntennaDiagonalBoltEdgeDistance As Double
    Private prop_AntennaTopGirtBoltEdgeDistance As Double
    Private prop_AntennaBotGirtBoltEdgeDistance As Double
    Private prop_AntennaInnerGirtBoltEdgeDistance As Double
    Private prop_AntennaHorizontalBoltEdgeDistance As Double
    Private prop_AntennaShortHorizontalBoltEdgeDistance As Double
    Private prop_AntennaDiagonalGageG1Distance As Double
    Private prop_AntennaTopGirtGageG1Distance As Double
    Private prop_AntennaBotGirtGageG1Distance As Double
    Private prop_AntennaInnerGirtGageG1Distance As Double
    Private prop_AntennaHorizontalGageG1Distance As Double
    Private prop_AntennaShortHorizontalGageG1Distance As Double
    Private prop_AntennaRedundantHorizontalBoltGrade As String
    Private prop_AntennaRedundantHorizontalBoltSize As Double
    Private prop_AntennaRedundantHorizontalNumBolts As Integer
    Private prop_AntennaRedundantHorizontalBoltEdgeDistance As Double
    Private prop_AntennaRedundantHorizontalGageG1Distance As Double
    Private prop_AntennaRedundantHorizontalNetWidthDeduct As Double
    Private prop_AntennaRedundantHorizontalUFactor As Double
    Private prop_AntennaRedundantDiagonalBoltGrade As String
    Private prop_AntennaRedundantDiagonalBoltSize As Double
    Private prop_AntennaRedundantDiagonalNumBolts As Integer
    Private prop_AntennaRedundantDiagonalBoltEdgeDistance As Double
    Private prop_AntennaRedundantDiagonalGageG1Distance As Double
    Private prop_AntennaRedundantDiagonalNetWidthDeduct As Double
    Private prop_AntennaRedundantDiagonalUFactor As Double
    Private prop_AntennaRedundantSubDiagonalBoltGrade As String
    Private prop_AntennaRedundantSubDiagonalBoltSize As Double
    Private prop_AntennaRedundantSubDiagonalNumBolts As Integer
    Private prop_AntennaRedundantSubDiagonalBoltEdgeDistance As Double
    Private prop_AntennaRedundantSubDiagonalGageG1Distance As Double
    Private prop_AntennaRedundantSubDiagonalNetWidthDeduct As Double
    Private prop_AntennaRedundantSubDiagonalUFactor As Double
    Private prop_AntennaRedundantSubHorizontalBoltGrade As String
    Private prop_AntennaRedundantSubHorizontalBoltSize As Double
    Private prop_AntennaRedundantSubHorizontalNumBolts As Integer
    Private prop_AntennaRedundantSubHorizontalBoltEdgeDistance As Double
    Private prop_AntennaRedundantSubHorizontalGageG1Distance As Double
    Private prop_AntennaRedundantSubHorizontalNetWidthDeduct As Double
    Private prop_AntennaRedundantSubHorizontalUFactor As Double
    Private prop_AntennaRedundantVerticalBoltGrade As String
    Private prop_AntennaRedundantVerticalBoltSize As Double
    Private prop_AntennaRedundantVerticalNumBolts As Integer
    Private prop_AntennaRedundantVerticalBoltEdgeDistance As Double
    Private prop_AntennaRedundantVerticalGageG1Distance As Double
    Private prop_AntennaRedundantVerticalNetWidthDeduct As Double
    Private prop_AntennaRedundantVerticalUFactor As Double
    Private prop_AntennaRedundantHipBoltGrade As String
    Private prop_AntennaRedundantHipBoltSize As Double
    Private prop_AntennaRedundantHipNumBolts As Integer
    Private prop_AntennaRedundantHipBoltEdgeDistance As Double
    Private prop_AntennaRedundantHipGageG1Distance As Double
    Private prop_AntennaRedundantHipNetWidthDeduct As Double
    Private prop_AntennaRedundantHipUFactor As Double
    Private prop_AntennaRedundantHipDiagonalBoltGrade As String
    Private prop_AntennaRedundantHipDiagonalBoltSize As Double
    Private prop_AntennaRedundantHipDiagonalNumBolts As Integer
    Private prop_AntennaRedundantHipDiagonalBoltEdgeDistance As Double
    Private prop_AntennaRedundantHipDiagonalGageG1Distance As Double
    Private prop_AntennaRedundantHipDiagonalNetWidthDeduct As Double
    Private prop_AntennaRedundantHipDiagonalUFactor As Double
    Private prop_AntennaDiagonalOutOfPlaneRestraint As Boolean
    Private prop_AntennaTopGirtOutOfPlaneRestraint As Boolean
    Private prop_AntennaBottomGirtOutOfPlaneRestraint As Boolean
    Private prop_AntennaMidGirtOutOfPlaneRestraint As Boolean
    Private prop_AntennaHorizontalOutOfPlaneRestraint As Boolean
    Private prop_AntennaSecondaryHorizontalOutOfPlaneRestraint As Boolean
    Private prop_AntennaDiagOffsetNEY As Double
    Private prop_AntennaDiagOffsetNEX As Double
    Private prop_AntennaDiagOffsetPEY As Double
    Private prop_AntennaDiagOffsetPEX As Double
    Private prop_AntennaKbraceOffsetNEY As Double
    Private prop_AntennaKbraceOffsetNEX As Double
    Private prop_AntennaKbraceOffsetPEY As Double
    Private prop_AntennaKbraceOffsetPEX As Double

    <Category("TNX Antenna Record"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer
        Get
            Return Me.prop_ID
        End Get
        Set
            Me.prop_ID = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Tnxid")>
    Public Property tnxID() As Integer
        Get
            Return Me.prop_tnxID
        End Get
        Set
            Me.prop_tnxID = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennarec")>
    Public Property AntennaRec() As Integer
        Get
            Return Me.prop_AntennaRec
        End Get
        Set
            Me.prop_AntennaRec = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennabracetype")>
    Public Property AntennaBraceType() As String
        Get
            Return Me.prop_AntennaBraceType
        End Get
        Set
            Me.prop_AntennaBraceType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaheight")>
    Public Property AntennaHeight() As Double
        Get
            Return Me.prop_AntennaHeight
        End Get
        Set
            Me.prop_AntennaHeight = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagonalspacing")>
    Public Property AntennaDiagonalSpacing() As Double
        Get
            Return Me.prop_AntennaDiagonalSpacing
        End Get
        Set
            Me.prop_AntennaDiagonalSpacing = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagonalspacingex")>
    Public Property AntennaDiagonalSpacingEx() As Double
        Get
            Return Me.prop_AntennaDiagonalSpacingEx
        End Get
        Set
            Me.prop_AntennaDiagonalSpacingEx = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennanumsections")>
    Public Property AntennaNumSections() As Integer
        Get
            Return Me.prop_AntennaNumSections
        End Get
        Set
            Me.prop_AntennaNumSections = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennanumsesctions")>
    Public Property AntennaNumSesctions() As Integer
        Get
            Return Me.prop_AntennaNumSesctions
        End Get
        Set
            Me.prop_AntennaNumSesctions = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennasectionlength")>
    Public Property AntennaSectionLength() As Double
        Get
            Return Me.prop_AntennaSectionLength
        End Get
        Set
            Me.prop_AntennaSectionLength = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennalegtype")>
    Public Property AntennaLegType() As String
        Get
            Return Me.prop_AntennaLegType
        End Get
        Set
            Me.prop_AntennaLegType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennalegsize")>
    Public Property AntennaLegSize() As String
        Get
            Return Me.prop_AntennaLegSize
        End Get
        Set
            Me.prop_AntennaLegSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaleggrade")>
    Public Property AntennaLegGrade() As Double
        Get
            Return Me.prop_AntennaLegGrade
        End Get
        Set
            Me.prop_AntennaLegGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennalegmatlgrade")>
    Public Property AntennaLegMatlGrade() As String
        Get
            Return Me.prop_AntennaLegMatlGrade
        End Get
        Set
            Me.prop_AntennaLegMatlGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagonalgrade")>
    Public Property AntennaDiagonalGrade() As Double
        Get
            Return Me.prop_AntennaDiagonalGrade
        End Get
        Set
            Me.prop_AntennaDiagonalGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagonalmatlgrade")>
    Public Property AntennaDiagonalMatlGrade() As String
        Get
            Return Me.prop_AntennaDiagonalMatlGrade
        End Get
        Set
            Me.prop_AntennaDiagonalMatlGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnerbracinggrade")>
    Public Property AntennaInnerBracingGrade() As Double
        Get
            Return Me.prop_AntennaInnerBracingGrade
        End Get
        Set
            Me.prop_AntennaInnerBracingGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnerbracingmatlgrade")>
    Public Property AntennaInnerBracingMatlGrade() As String
        Get
            Return Me.prop_AntennaInnerBracingMatlGrade
        End Get
        Set
            Me.prop_AntennaInnerBracingMatlGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennatopgirtgrade")>
    Public Property AntennaTopGirtGrade() As Double
        Get
            Return Me.prop_AntennaTopGirtGrade
        End Get
        Set
            Me.prop_AntennaTopGirtGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennatopgirtmatlgrade")>
    Public Property AntennaTopGirtMatlGrade() As String
        Get
            Return Me.prop_AntennaTopGirtMatlGrade
        End Get
        Set
            Me.prop_AntennaTopGirtMatlGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennabotgirtgrade")>
    Public Property AntennaBotGirtGrade() As Double
        Get
            Return Me.prop_AntennaBotGirtGrade
        End Get
        Set
            Me.prop_AntennaBotGirtGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennabotgirtmatlgrade")>
    Public Property AntennaBotGirtMatlGrade() As String
        Get
            Return Me.prop_AntennaBotGirtMatlGrade
        End Get
        Set
            Me.prop_AntennaBotGirtMatlGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnergirtgrade")>
    Public Property AntennaInnerGirtGrade() As Double
        Get
            Return Me.prop_AntennaInnerGirtGrade
        End Get
        Set
            Me.prop_AntennaInnerGirtGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnergirtmatlgrade")>
    Public Property AntennaInnerGirtMatlGrade() As String
        Get
            Return Me.prop_AntennaInnerGirtMatlGrade
        End Get
        Set
            Me.prop_AntennaInnerGirtMatlGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennalonghorizontalgrade")>
    Public Property AntennaLongHorizontalGrade() As Double
        Get
            Return Me.prop_AntennaLongHorizontalGrade
        End Get
        Set
            Me.prop_AntennaLongHorizontalGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennalonghorizontalmatlgrade")>
    Public Property AntennaLongHorizontalMatlGrade() As String
        Get
            Return Me.prop_AntennaLongHorizontalMatlGrade
        End Get
        Set
            Me.prop_AntennaLongHorizontalMatlGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennashorthorizontalgrade")>
    Public Property AntennaShortHorizontalGrade() As Double
        Get
            Return Me.prop_AntennaShortHorizontalGrade
        End Get
        Set
            Me.prop_AntennaShortHorizontalGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennashorthorizontalmatlgrade")>
    Public Property AntennaShortHorizontalMatlGrade() As String
        Get
            Return Me.prop_AntennaShortHorizontalMatlGrade
        End Get
        Set
            Me.prop_AntennaShortHorizontalMatlGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagonaltype")>
    Public Property AntennaDiagonalType() As String
        Get
            Return Me.prop_AntennaDiagonalType
        End Get
        Set
            Me.prop_AntennaDiagonalType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagonalsize")>
    Public Property AntennaDiagonalSize() As String
        Get
            Return Me.prop_AntennaDiagonalSize
        End Get
        Set
            Me.prop_AntennaDiagonalSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnerbracingtype")>
    Public Property AntennaInnerBracingType() As String
        Get
            Return Me.prop_AntennaInnerBracingType
        End Get
        Set
            Me.prop_AntennaInnerBracingType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnerbracingsize")>
    Public Property AntennaInnerBracingSize() As String
        Get
            Return Me.prop_AntennaInnerBracingSize
        End Get
        Set
            Me.prop_AntennaInnerBracingSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennatopgirttype")>
    Public Property AntennaTopGirtType() As String
        Get
            Return Me.prop_AntennaTopGirtType
        End Get
        Set
            Me.prop_AntennaTopGirtType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennatopgirtsize")>
    Public Property AntennaTopGirtSize() As String
        Get
            Return Me.prop_AntennaTopGirtSize
        End Get
        Set
            Me.prop_AntennaTopGirtSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennabotgirttype")>
    Public Property AntennaBotGirtType() As String
        Get
            Return Me.prop_AntennaBotGirtType
        End Get
        Set
            Me.prop_AntennaBotGirtType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennabotgirtsize")>
    Public Property AntennaBotGirtSize() As String
        Get
            Return Me.prop_AntennaBotGirtSize
        End Get
        Set
            Me.prop_AntennaBotGirtSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennatopgirtoffset")>
    Public Property AntennaTopGirtOffset() As Double
        Get
            Return Me.prop_AntennaTopGirtOffset
        End Get
        Set
            Me.prop_AntennaTopGirtOffset = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennabotgirtoffset")>
    Public Property AntennaBotGirtOffset() As Double
        Get
            Return Me.prop_AntennaBotGirtOffset
        End Get
        Set
            Me.prop_AntennaBotGirtOffset = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennahaskbraceendpanels")>
    Public Property AntennaHasKBraceEndPanels() As Boolean
        Get
            Return Me.prop_AntennaHasKBraceEndPanels
        End Get
        Set
            Me.prop_AntennaHasKBraceEndPanels = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennahashorizontals")>
    Public Property AntennaHasHorizontals() As Boolean
        Get
            Return Me.prop_AntennaHasHorizontals
        End Get
        Set
            Me.prop_AntennaHasHorizontals = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennalonghorizontaltype")>
    Public Property AntennaLongHorizontalType() As String
        Get
            Return Me.prop_AntennaLongHorizontalType
        End Get
        Set
            Me.prop_AntennaLongHorizontalType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennalonghorizontalsize")>
    Public Property AntennaLongHorizontalSize() As String
        Get
            Return Me.prop_AntennaLongHorizontalSize
        End Get
        Set
            Me.prop_AntennaLongHorizontalSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennashorthorizontaltype")>
    Public Property AntennaShortHorizontalType() As String
        Get
            Return Me.prop_AntennaShortHorizontalType
        End Get
        Set
            Me.prop_AntennaShortHorizontalType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennashorthorizontalsize")>
    Public Property AntennaShortHorizontalSize() As String
        Get
            Return Me.prop_AntennaShortHorizontalSize
        End Get
        Set
            Me.prop_AntennaShortHorizontalSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantgrade")>
    Public Property AntennaRedundantGrade() As Double
        Get
            Return Me.prop_AntennaRedundantGrade
        End Get
        Set
            Me.prop_AntennaRedundantGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantmatlgrade")>
    Public Property AntennaRedundantMatlGrade() As String
        Get
            Return Me.prop_AntennaRedundantMatlGrade
        End Get
        Set
            Me.prop_AntennaRedundantMatlGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanttype")>
    Public Property AntennaRedundantType() As String
        Get
            Return Me.prop_AntennaRedundantType
        End Get
        Set
            Me.prop_AntennaRedundantType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantdiagtype")>
    Public Property AntennaRedundantDiagType() As String
        Get
            Return Me.prop_AntennaRedundantDiagType
        End Get
        Set
            Me.prop_AntennaRedundantDiagType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubdiagonaltype")>
    Public Property AntennaRedundantSubDiagonalType() As String
        Get
            Return Me.prop_AntennaRedundantSubDiagonalType
        End Get
        Set
            Me.prop_AntennaRedundantSubDiagonalType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubhorizontaltype")>
    Public Property AntennaRedundantSubHorizontalType() As String
        Get
            Return Me.prop_AntennaRedundantSubHorizontalType
        End Get
        Set
            Me.prop_AntennaRedundantSubHorizontalType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantverticaltype")>
    Public Property AntennaRedundantVerticalType() As String
        Get
            Return Me.prop_AntennaRedundantVerticalType
        End Get
        Set
            Me.prop_AntennaRedundantVerticalType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthiptype")>
    Public Property AntennaRedundantHipType() As String
        Get
            Return Me.prop_AntennaRedundantHipType
        End Get
        Set
            Me.prop_AntennaRedundantHipType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipdiagonaltype")>
    Public Property AntennaRedundantHipDiagonalType() As String
        Get
            Return Me.prop_AntennaRedundantHipDiagonalType
        End Get
        Set
            Me.prop_AntennaRedundantHipDiagonalType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthorizontalsize")>
    Public Property AntennaRedundantHorizontalSize() As String
        Get
            Return Me.prop_AntennaRedundantHorizontalSize
        End Get
        Set
            Me.prop_AntennaRedundantHorizontalSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthorizontalsize2")>
    Public Property AntennaRedundantHorizontalSize2() As String
        Get
            Return Me.prop_AntennaRedundantHorizontalSize2
        End Get
        Set
            Me.prop_AntennaRedundantHorizontalSize2 = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthorizontalsize3")>
    Public Property AntennaRedundantHorizontalSize3() As String
        Get
            Return Me.prop_AntennaRedundantHorizontalSize3
        End Get
        Set
            Me.prop_AntennaRedundantHorizontalSize3 = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthorizontalsize4")>
    Public Property AntennaRedundantHorizontalSize4() As String
        Get
            Return Me.prop_AntennaRedundantHorizontalSize4
        End Get
        Set
            Me.prop_AntennaRedundantHorizontalSize4 = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantdiagonalsize")>
    Public Property AntennaRedundantDiagonalSize() As String
        Get
            Return Me.prop_AntennaRedundantDiagonalSize
        End Get
        Set
            Me.prop_AntennaRedundantDiagonalSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantdiagonalsize2")>
    Public Property AntennaRedundantDiagonalSize2() As String
        Get
            Return Me.prop_AntennaRedundantDiagonalSize2
        End Get
        Set
            Me.prop_AntennaRedundantDiagonalSize2 = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantdiagonalsize3")>
    Public Property AntennaRedundantDiagonalSize3() As String
        Get
            Return Me.prop_AntennaRedundantDiagonalSize3
        End Get
        Set
            Me.prop_AntennaRedundantDiagonalSize3 = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantdiagonalsize4")>
    Public Property AntennaRedundantDiagonalSize4() As String
        Get
            Return Me.prop_AntennaRedundantDiagonalSize4
        End Get
        Set
            Me.prop_AntennaRedundantDiagonalSize4 = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubhorizontalsize")>
    Public Property AntennaRedundantSubHorizontalSize() As String
        Get
            Return Me.prop_AntennaRedundantSubHorizontalSize
        End Get
        Set
            Me.prop_AntennaRedundantSubHorizontalSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubdiagonalsize")>
    Public Property AntennaRedundantSubDiagonalSize() As String
        Get
            Return Me.prop_AntennaRedundantSubDiagonalSize
        End Get
        Set
            Me.prop_AntennaRedundantSubDiagonalSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennasubdiaglocation")>
    Public Property AntennaSubDiagLocation() As Double
        Get
            Return Me.prop_AntennaSubDiagLocation
        End Get
        Set
            Me.prop_AntennaSubDiagLocation = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantverticalsize")>
    Public Property AntennaRedundantVerticalSize() As String
        Get
            Return Me.prop_AntennaRedundantVerticalSize
        End Get
        Set
            Me.prop_AntennaRedundantVerticalSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipdiagonalsize")>
    Public Property AntennaRedundantHipDiagonalSize() As String
        Get
            Return Me.prop_AntennaRedundantHipDiagonalSize
        End Get
        Set
            Me.prop_AntennaRedundantHipDiagonalSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipdiagonalsize2")>
    Public Property AntennaRedundantHipDiagonalSize2() As String
        Get
            Return Me.prop_AntennaRedundantHipDiagonalSize2
        End Get
        Set
            Me.prop_AntennaRedundantHipDiagonalSize2 = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipdiagonalsize3")>
    Public Property AntennaRedundantHipDiagonalSize3() As String
        Get
            Return Me.prop_AntennaRedundantHipDiagonalSize3
        End Get
        Set
            Me.prop_AntennaRedundantHipDiagonalSize3 = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipdiagonalsize4")>
    Public Property AntennaRedundantHipDiagonalSize4() As String
        Get
            Return Me.prop_AntennaRedundantHipDiagonalSize4
        End Get
        Set
            Me.prop_AntennaRedundantHipDiagonalSize4 = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipsize")>
    Public Property AntennaRedundantHipSize() As String
        Get
            Return Me.prop_AntennaRedundantHipSize
        End Get
        Set
            Me.prop_AntennaRedundantHipSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipsize2")>
    Public Property AntennaRedundantHipSize2() As String
        Get
            Return Me.prop_AntennaRedundantHipSize2
        End Get
        Set
            Me.prop_AntennaRedundantHipSize2 = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipsize3")>
    Public Property AntennaRedundantHipSize3() As String
        Get
            Return Me.prop_AntennaRedundantHipSize3
        End Get
        Set
            Me.prop_AntennaRedundantHipSize3 = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipsize4")>
    Public Property AntennaRedundantHipSize4() As String
        Get
            Return Me.prop_AntennaRedundantHipSize4
        End Get
        Set
            Me.prop_AntennaRedundantHipSize4 = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennanuminnergirts")>
    Public Property AntennaNumInnerGirts() As Integer
        Get
            Return Me.prop_AntennaNumInnerGirts
        End Get
        Set
            Me.prop_AntennaNumInnerGirts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnergirttype")>
    Public Property AntennaInnerGirtType() As String
        Get
            Return Me.prop_AntennaInnerGirtType
        End Get
        Set
            Me.prop_AntennaInnerGirtType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnergirtsize")>
    Public Property AntennaInnerGirtSize() As String
        Get
            Return Me.prop_AntennaInnerGirtSize
        End Get
        Set
            Me.prop_AntennaInnerGirtSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennapoleshapetype")>
    Public Property AntennaPoleShapeType() As String
        Get
            Return Me.prop_AntennaPoleShapeType
        End Get
        Set
            Me.prop_AntennaPoleShapeType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennapolesize")>
    Public Property AntennaPoleSize() As String
        Get
            Return Me.prop_AntennaPoleSize
        End Get
        Set
            Me.prop_AntennaPoleSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennapolegrade")>
    Public Property AntennaPoleGrade() As Double
        Get
            Return Me.prop_AntennaPoleGrade
        End Get
        Set
            Me.prop_AntennaPoleGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennapolematlgrade")>
    Public Property AntennaPoleMatlGrade() As String
        Get
            Return Me.prop_AntennaPoleMatlGrade
        End Get
        Set
            Me.prop_AntennaPoleMatlGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennapolesplicelength")>
    Public Property AntennaPoleSpliceLength() As Double
        Get
            Return Me.prop_AntennaPoleSpliceLength
        End Get
        Set
            Me.prop_AntennaPoleSpliceLength = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennataperpolenumsides")>
    Public Property AntennaTaperPoleNumSides() As Integer
        Get
            Return Me.prop_AntennaTaperPoleNumSides
        End Get
        Set
            Me.prop_AntennaTaperPoleNumSides = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennataperpoletopdiameter")>
    Public Property AntennaTaperPoleTopDiameter() As Double
        Get
            Return Me.prop_AntennaTaperPoleTopDiameter
        End Get
        Set
            Me.prop_AntennaTaperPoleTopDiameter = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennataperpolebotdiameter")>
    Public Property AntennaTaperPoleBotDiameter() As Double
        Get
            Return Me.prop_AntennaTaperPoleBotDiameter
        End Get
        Set
            Me.prop_AntennaTaperPoleBotDiameter = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennataperpolewallthickness")>
    Public Property AntennaTaperPoleWallThickness() As Double
        Get
            Return Me.prop_AntennaTaperPoleWallThickness
        End Get
        Set
            Me.prop_AntennaTaperPoleWallThickness = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennataperpolebendradius")>
    Public Property AntennaTaperPoleBendRadius() As Double
        Get
            Return Me.prop_AntennaTaperPoleBendRadius
        End Get
        Set
            Me.prop_AntennaTaperPoleBendRadius = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennataperpolegrade")>
    Public Property AntennaTaperPoleGrade() As Double
        Get
            Return Me.prop_AntennaTaperPoleGrade
        End Get
        Set
            Me.prop_AntennaTaperPoleGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennataperpolematlgrade")>
    Public Property AntennaTaperPoleMatlGrade() As String
        Get
            Return Me.prop_AntennaTaperPoleMatlGrade
        End Get
        Set
            Me.prop_AntennaTaperPoleMatlGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaswmult")>
    Public Property AntennaSWMult() As Double
        Get
            Return Me.prop_AntennaSWMult
        End Get
        Set
            Me.prop_AntennaSWMult = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennawpmult")>
    Public Property AntennaWPMult() As Double
        Get
            Return Me.prop_AntennaWPMult
        End Get
        Set
            Me.prop_AntennaWPMult = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaautocalcksingleangle")>
    Public Property AntennaAutoCalcKSingleAngle() As Double
        Get
            Return Me.prop_AntennaAutoCalcKSingleAngle
        End Get
        Set
            Me.prop_AntennaAutoCalcKSingleAngle = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaautocalcksolidround")>
    Public Property AntennaAutoCalcKSolidRound() As Double
        Get
            Return Me.prop_AntennaAutoCalcKSolidRound
        End Get
        Set
            Me.prop_AntennaAutoCalcKSolidRound = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaafgusset")>
    Public Property AntennaAfGusset() As Double
        Get
            Return Me.prop_AntennaAfGusset
        End Get
        Set
            Me.prop_AntennaAfGusset = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennatfgusset")>
    Public Property AntennaTfGusset() As Double
        Get
            Return Me.prop_AntennaTfGusset
        End Get
        Set
            Me.prop_AntennaTfGusset = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennagussetboltedgedistance")>
    Public Property AntennaGussetBoltEdgeDistance() As Double
        Get
            Return Me.prop_AntennaGussetBoltEdgeDistance
        End Get
        Set
            Me.prop_AntennaGussetBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennagussetgrade")>
    Public Property AntennaGussetGrade() As Double
        Get
            Return Me.prop_AntennaGussetGrade
        End Get
        Set
            Me.prop_AntennaGussetGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennagussetmatlgrade")>
    Public Property AntennaGussetMatlGrade() As String
        Get
            Return Me.prop_AntennaGussetMatlGrade
        End Get
        Set
            Me.prop_AntennaGussetMatlGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaafmult")>
    Public Property AntennaAfMult() As Double
        Get
            Return Me.prop_AntennaAfMult
        End Get
        Set
            Me.prop_AntennaAfMult = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaarmult")>
    Public Property AntennaArMult() As Double
        Get
            Return Me.prop_AntennaArMult
        End Get
        Set
            Me.prop_AntennaArMult = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaflatipapole")>
    Public Property AntennaFlatIPAPole() As Double
        Get
            Return Me.prop_AntennaFlatIPAPole
        End Get
        Set
            Me.prop_AntennaFlatIPAPole = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaroundipapole")>
    Public Property AntennaRoundIPAPole() As Double
        Get
            Return Me.prop_AntennaRoundIPAPole
        End Get
        Set
            Me.prop_AntennaRoundIPAPole = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaflatipaleg")>
    Public Property AntennaFlatIPALeg() As Double
        Get
            Return Me.prop_AntennaFlatIPALeg
        End Get
        Set
            Me.prop_AntennaFlatIPALeg = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaroundipaleg")>
    Public Property AntennaRoundIPALeg() As Double
        Get
            Return Me.prop_AntennaRoundIPALeg
        End Get
        Set
            Me.prop_AntennaRoundIPALeg = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaflatipahorizontal")>
    Public Property AntennaFlatIPAHorizontal() As Double
        Get
            Return Me.prop_AntennaFlatIPAHorizontal
        End Get
        Set
            Me.prop_AntennaFlatIPAHorizontal = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaroundipahorizontal")>
    Public Property AntennaRoundIPAHorizontal() As Double
        Get
            Return Me.prop_AntennaRoundIPAHorizontal
        End Get
        Set
            Me.prop_AntennaRoundIPAHorizontal = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaflatipadiagonal")>
    Public Property AntennaFlatIPADiagonal() As Double
        Get
            Return Me.prop_AntennaFlatIPADiagonal
        End Get
        Set
            Me.prop_AntennaFlatIPADiagonal = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaroundipadiagonal")>
    Public Property AntennaRoundIPADiagonal() As Double
        Get
            Return Me.prop_AntennaRoundIPADiagonal
        End Get
        Set
            Me.prop_AntennaRoundIPADiagonal = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennacsa_S37_Speedupfactor")>
    Public Property AntennaCSA_S37_SpeedUpFactor() As Double
        Get
            Return Me.prop_AntennaCSA_S37_SpeedUpFactor
        End Get
        Set
            Me.prop_AntennaCSA_S37_SpeedUpFactor = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaklegs")>
    Public Property AntennaKLegs() As Double
        Get
            Return Me.prop_AntennaKLegs
        End Get
        Set
            Me.prop_AntennaKLegs = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakxbraceddiags")>
    Public Property AntennaKXBracedDiags() As Double
        Get
            Return Me.prop_AntennaKXBracedDiags
        End Get
        Set
            Me.prop_AntennaKXBracedDiags = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakkbraceddiags")>
    Public Property AntennaKKBracedDiags() As Double
        Get
            Return Me.prop_AntennaKKBracedDiags
        End Get
        Set
            Me.prop_AntennaKKBracedDiags = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakzbraceddiags")>
    Public Property AntennaKZBracedDiags() As Double
        Get
            Return Me.prop_AntennaKZBracedDiags
        End Get
        Set
            Me.prop_AntennaKZBracedDiags = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakhorzs")>
    Public Property AntennaKHorzs() As Double
        Get
            Return Me.prop_AntennaKHorzs
        End Get
        Set
            Me.prop_AntennaKHorzs = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaksechorzs")>
    Public Property AntennaKSecHorzs() As Double
        Get
            Return Me.prop_AntennaKSecHorzs
        End Get
        Set
            Me.prop_AntennaKSecHorzs = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakgirts")>
    Public Property AntennaKGirts() As Double
        Get
            Return Me.prop_AntennaKGirts
        End Get
        Set
            Me.prop_AntennaKGirts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakinners")>
    Public Property AntennaKInners() As Double
        Get
            Return Me.prop_AntennaKInners
        End Get
        Set
            Me.prop_AntennaKInners = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakxbraceddiagsy")>
    Public Property AntennaKXBracedDiagsY() As Double
        Get
            Return Me.prop_AntennaKXBracedDiagsY
        End Get
        Set
            Me.prop_AntennaKXBracedDiagsY = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakkbraceddiagsy")>
    Public Property AntennaKKBracedDiagsY() As Double
        Get
            Return Me.prop_AntennaKKBracedDiagsY
        End Get
        Set
            Me.prop_AntennaKKBracedDiagsY = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakzbraceddiagsy")>
    Public Property AntennaKZBracedDiagsY() As Double
        Get
            Return Me.prop_AntennaKZBracedDiagsY
        End Get
        Set
            Me.prop_AntennaKZBracedDiagsY = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakhorzsy")>
    Public Property AntennaKHorzsY() As Double
        Get
            Return Me.prop_AntennaKHorzsY
        End Get
        Set
            Me.prop_AntennaKHorzsY = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaksechorzsy")>
    Public Property AntennaKSecHorzsY() As Double
        Get
            Return Me.prop_AntennaKSecHorzsY
        End Get
        Set
            Me.prop_AntennaKSecHorzsY = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakgirtsy")>
    Public Property AntennaKGirtsY() As Double
        Get
            Return Me.prop_AntennaKGirtsY
        End Get
        Set
            Me.prop_AntennaKGirtsY = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakinnersy")>
    Public Property AntennaKInnersY() As Double
        Get
            Return Me.prop_AntennaKInnersY
        End Get
        Set
            Me.prop_AntennaKInnersY = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakredhorz")>
    Public Property AntennaKRedHorz() As Double
        Get
            Return Me.prop_AntennaKRedHorz
        End Get
        Set
            Me.prop_AntennaKRedHorz = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakreddiag")>
    Public Property AntennaKRedDiag() As Double
        Get
            Return Me.prop_AntennaKRedDiag
        End Get
        Set
            Me.prop_AntennaKRedDiag = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakredsubdiag")>
    Public Property AntennaKRedSubDiag() As Double
        Get
            Return Me.prop_AntennaKRedSubDiag
        End Get
        Set
            Me.prop_AntennaKRedSubDiag = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakredsubhorz")>
    Public Property AntennaKRedSubHorz() As Double
        Get
            Return Me.prop_AntennaKRedSubHorz
        End Get
        Set
            Me.prop_AntennaKRedSubHorz = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakredvert")>
    Public Property AntennaKRedVert() As Double
        Get
            Return Me.prop_AntennaKRedVert
        End Get
        Set
            Me.prop_AntennaKRedVert = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakredhip")>
    Public Property AntennaKRedHip() As Double
        Get
            Return Me.prop_AntennaKRedHip
        End Get
        Set
            Me.prop_AntennaKRedHip = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakredhipdiag")>
    Public Property AntennaKRedHipDiag() As Double
        Get
            Return Me.prop_AntennaKRedHipDiag
        End Get
        Set
            Me.prop_AntennaKRedHipDiag = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaktlx")>
    Public Property AntennaKTLX() As Double
        Get
            Return Me.prop_AntennaKTLX
        End Get
        Set
            Me.prop_AntennaKTLX = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaktlz")>
    Public Property AntennaKTLZ() As Double
        Get
            Return Me.prop_AntennaKTLZ
        End Get
        Set
            Me.prop_AntennaKTLZ = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaktlleg")>
    Public Property AntennaKTLLeg() As Double
        Get
            Return Me.prop_AntennaKTLLeg
        End Get
        Set
            Me.prop_AntennaKTLLeg = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnerktlx")>
    Public Property AntennaInnerKTLX() As Double
        Get
            Return Me.prop_AntennaInnerKTLX
        End Get
        Set
            Me.prop_AntennaInnerKTLX = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnerktlz")>
    Public Property AntennaInnerKTLZ() As Double
        Get
            Return Me.prop_AntennaInnerKTLZ
        End Get
        Set
            Me.prop_AntennaInnerKTLZ = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnerktlleg")>
    Public Property AntennaInnerKTLLeg() As Double
        Get
            Return Me.prop_AntennaInnerKTLLeg
        End Get
        Set
            Me.prop_AntennaInnerKTLLeg = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennastitchboltlocationhoriz")>
    Public Property AntennaStitchBoltLocationHoriz() As String
        Get
            Return Me.prop_AntennaStitchBoltLocationHoriz
        End Get
        Set
            Me.prop_AntennaStitchBoltLocationHoriz = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennastitchboltlocationdiag")>
    Public Property AntennaStitchBoltLocationDiag() As String
        Get
            Return Me.prop_AntennaStitchBoltLocationDiag
        End Get
        Set
            Me.prop_AntennaStitchBoltLocationDiag = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennastitchspacing")>
    Public Property AntennaStitchSpacing() As Double
        Get
            Return Me.prop_AntennaStitchSpacing
        End Get
        Set
            Me.prop_AntennaStitchSpacing = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennastitchspacinghorz")>
    Public Property AntennaStitchSpacingHorz() As Double
        Get
            Return Me.prop_AntennaStitchSpacingHorz
        End Get
        Set
            Me.prop_AntennaStitchSpacingHorz = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennastitchspacingdiag")>
    Public Property AntennaStitchSpacingDiag() As Double
        Get
            Return Me.prop_AntennaStitchSpacingDiag
        End Get
        Set
            Me.prop_AntennaStitchSpacingDiag = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennastitchspacingred")>
    Public Property AntennaStitchSpacingRed() As Double
        Get
            Return Me.prop_AntennaStitchSpacingRed
        End Get
        Set
            Me.prop_AntennaStitchSpacingRed = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennalegnetwidthdeduct")>
    Public Property AntennaLegNetWidthDeduct() As Double
        Get
            Return Me.prop_AntennaLegNetWidthDeduct
        End Get
        Set
            Me.prop_AntennaLegNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennalegufactor")>
    Public Property AntennaLegUFactor() As Double
        Get
            Return Me.prop_AntennaLegUFactor
        End Get
        Set
            Me.prop_AntennaLegUFactor = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagonalnetwidthdeduct")>
    Public Property AntennaDiagonalNetWidthDeduct() As Double
        Get
            Return Me.prop_AntennaDiagonalNetWidthDeduct
        End Get
        Set
            Me.prop_AntennaDiagonalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennatopgirtnetwidthdeduct")>
    Public Property AntennaTopGirtNetWidthDeduct() As Double
        Get
            Return Me.prop_AntennaTopGirtNetWidthDeduct
        End Get
        Set
            Me.prop_AntennaTopGirtNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennabotgirtnetwidthdeduct")>
    Public Property AntennaBotGirtNetWidthDeduct() As Double
        Get
            Return Me.prop_AntennaBotGirtNetWidthDeduct
        End Get
        Set
            Me.prop_AntennaBotGirtNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnergirtnetwidthdeduct")>
    Public Property AntennaInnerGirtNetWidthDeduct() As Double
        Get
            Return Me.prop_AntennaInnerGirtNetWidthDeduct
        End Get
        Set
            Me.prop_AntennaInnerGirtNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennahorizontalnetwidthdeduct")>
    Public Property AntennaHorizontalNetWidthDeduct() As Double
        Get
            Return Me.prop_AntennaHorizontalNetWidthDeduct
        End Get
        Set
            Me.prop_AntennaHorizontalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennashorthorizontalnetwidthdeduct")>
    Public Property AntennaShortHorizontalNetWidthDeduct() As Double
        Get
            Return Me.prop_AntennaShortHorizontalNetWidthDeduct
        End Get
        Set
            Me.prop_AntennaShortHorizontalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagonalufactor")>
    Public Property AntennaDiagonalUFactor() As Double
        Get
            Return Me.prop_AntennaDiagonalUFactor
        End Get
        Set
            Me.prop_AntennaDiagonalUFactor = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennatopgirtufactor")>
    Public Property AntennaTopGirtUFactor() As Double
        Get
            Return Me.prop_AntennaTopGirtUFactor
        End Get
        Set
            Me.prop_AntennaTopGirtUFactor = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennabotgirtufactor")>
    Public Property AntennaBotGirtUFactor() As Double
        Get
            Return Me.prop_AntennaBotGirtUFactor
        End Get
        Set
            Me.prop_AntennaBotGirtUFactor = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnergirtufactor")>
    Public Property AntennaInnerGirtUFactor() As Double
        Get
            Return Me.prop_AntennaInnerGirtUFactor
        End Get
        Set
            Me.prop_AntennaInnerGirtUFactor = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennahorizontalufactor")>
    Public Property AntennaHorizontalUFactor() As Double
        Get
            Return Me.prop_AntennaHorizontalUFactor
        End Get
        Set
            Me.prop_AntennaHorizontalUFactor = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennashorthorizontalufactor")>
    Public Property AntennaShortHorizontalUFactor() As Double
        Get
            Return Me.prop_AntennaShortHorizontalUFactor
        End Get
        Set
            Me.prop_AntennaShortHorizontalUFactor = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennalegconntype")>
    Public Property AntennaLegConnType() As String
        Get
            Return Me.prop_AntennaLegConnType
        End Get
        Set
            Me.prop_AntennaLegConnType = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennalegnumbolts")>
    Public Property AntennaLegNumBolts() As Integer
        Get
            Return Me.prop_AntennaLegNumBolts
        End Get
        Set
            Me.prop_AntennaLegNumBolts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagonalnumbolts")>
    Public Property AntennaDiagonalNumBolts() As Integer
        Get
            Return Me.prop_AntennaDiagonalNumBolts
        End Get
        Set
            Me.prop_AntennaDiagonalNumBolts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennatopgirtnumbolts")>
    Public Property AntennaTopGirtNumBolts() As Integer
        Get
            Return Me.prop_AntennaTopGirtNumBolts
        End Get
        Set
            Me.prop_AntennaTopGirtNumBolts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennabotgirtnumbolts")>
    Public Property AntennaBotGirtNumBolts() As Integer
        Get
            Return Me.prop_AntennaBotGirtNumBolts
        End Get
        Set
            Me.prop_AntennaBotGirtNumBolts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnergirtnumbolts")>
    Public Property AntennaInnerGirtNumBolts() As Integer
        Get
            Return Me.prop_AntennaInnerGirtNumBolts
        End Get
        Set
            Me.prop_AntennaInnerGirtNumBolts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennahorizontalnumbolts")>
    Public Property AntennaHorizontalNumBolts() As Integer
        Get
            Return Me.prop_AntennaHorizontalNumBolts
        End Get
        Set
            Me.prop_AntennaHorizontalNumBolts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennashorthorizontalnumbolts")>
    Public Property AntennaShortHorizontalNumBolts() As Integer
        Get
            Return Me.prop_AntennaShortHorizontalNumBolts
        End Get
        Set
            Me.prop_AntennaShortHorizontalNumBolts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennalegboltgrade")>
    Public Property AntennaLegBoltGrade() As String
        Get
            Return Me.prop_AntennaLegBoltGrade
        End Get
        Set
            Me.prop_AntennaLegBoltGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennalegboltsize")>
    Public Property AntennaLegBoltSize() As Double
        Get
            Return Me.prop_AntennaLegBoltSize
        End Get
        Set
            Me.prop_AntennaLegBoltSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagonalboltgrade")>
    Public Property AntennaDiagonalBoltGrade() As String
        Get
            Return Me.prop_AntennaDiagonalBoltGrade
        End Get
        Set
            Me.prop_AntennaDiagonalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagonalboltsize")>
    Public Property AntennaDiagonalBoltSize() As Double
        Get
            Return Me.prop_AntennaDiagonalBoltSize
        End Get
        Set
            Me.prop_AntennaDiagonalBoltSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennatopgirtboltgrade")>
    Public Property AntennaTopGirtBoltGrade() As String
        Get
            Return Me.prop_AntennaTopGirtBoltGrade
        End Get
        Set
            Me.prop_AntennaTopGirtBoltGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennatopgirtboltsize")>
    Public Property AntennaTopGirtBoltSize() As Double
        Get
            Return Me.prop_AntennaTopGirtBoltSize
        End Get
        Set
            Me.prop_AntennaTopGirtBoltSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennabotgirtboltgrade")>
    Public Property AntennaBotGirtBoltGrade() As String
        Get
            Return Me.prop_AntennaBotGirtBoltGrade
        End Get
        Set
            Me.prop_AntennaBotGirtBoltGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennabotgirtboltsize")>
    Public Property AntennaBotGirtBoltSize() As Double
        Get
            Return Me.prop_AntennaBotGirtBoltSize
        End Get
        Set
            Me.prop_AntennaBotGirtBoltSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnergirtboltgrade")>
    Public Property AntennaInnerGirtBoltGrade() As String
        Get
            Return Me.prop_AntennaInnerGirtBoltGrade
        End Get
        Set
            Me.prop_AntennaInnerGirtBoltGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnergirtboltsize")>
    Public Property AntennaInnerGirtBoltSize() As Double
        Get
            Return Me.prop_AntennaInnerGirtBoltSize
        End Get
        Set
            Me.prop_AntennaInnerGirtBoltSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennahorizontalboltgrade")>
    Public Property AntennaHorizontalBoltGrade() As String
        Get
            Return Me.prop_AntennaHorizontalBoltGrade
        End Get
        Set
            Me.prop_AntennaHorizontalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennahorizontalboltsize")>
    Public Property AntennaHorizontalBoltSize() As Double
        Get
            Return Me.prop_AntennaHorizontalBoltSize
        End Get
        Set
            Me.prop_AntennaHorizontalBoltSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennashorthorizontalboltgrade")>
    Public Property AntennaShortHorizontalBoltGrade() As String
        Get
            Return Me.prop_AntennaShortHorizontalBoltGrade
        End Get
        Set
            Me.prop_AntennaShortHorizontalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennashorthorizontalboltsize")>
    Public Property AntennaShortHorizontalBoltSize() As Double
        Get
            Return Me.prop_AntennaShortHorizontalBoltSize
        End Get
        Set
            Me.prop_AntennaShortHorizontalBoltSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennalegboltedgedistance")>
    Public Property AntennaLegBoltEdgeDistance() As Double
        Get
            Return Me.prop_AntennaLegBoltEdgeDistance
        End Get
        Set
            Me.prop_AntennaLegBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagonalboltedgedistance")>
    Public Property AntennaDiagonalBoltEdgeDistance() As Double
        Get
            Return Me.prop_AntennaDiagonalBoltEdgeDistance
        End Get
        Set
            Me.prop_AntennaDiagonalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennatopgirtboltedgedistance")>
    Public Property AntennaTopGirtBoltEdgeDistance() As Double
        Get
            Return Me.prop_AntennaTopGirtBoltEdgeDistance
        End Get
        Set
            Me.prop_AntennaTopGirtBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennabotgirtboltedgedistance")>
    Public Property AntennaBotGirtBoltEdgeDistance() As Double
        Get
            Return Me.prop_AntennaBotGirtBoltEdgeDistance
        End Get
        Set
            Me.prop_AntennaBotGirtBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnergirtboltedgedistance")>
    Public Property AntennaInnerGirtBoltEdgeDistance() As Double
        Get
            Return Me.prop_AntennaInnerGirtBoltEdgeDistance
        End Get
        Set
            Me.prop_AntennaInnerGirtBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennahorizontalboltedgedistance")>
    Public Property AntennaHorizontalBoltEdgeDistance() As Double
        Get
            Return Me.prop_AntennaHorizontalBoltEdgeDistance
        End Get
        Set
            Me.prop_AntennaHorizontalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennashorthorizontalboltedgedistance")>
    Public Property AntennaShortHorizontalBoltEdgeDistance() As Double
        Get
            Return Me.prop_AntennaShortHorizontalBoltEdgeDistance
        End Get
        Set
            Me.prop_AntennaShortHorizontalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagonalgageg1Distance")>
    Public Property AntennaDiagonalGageG1Distance() As Double
        Get
            Return Me.prop_AntennaDiagonalGageG1Distance
        End Get
        Set
            Me.prop_AntennaDiagonalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennatopgirtgageg1Distance")>
    Public Property AntennaTopGirtGageG1Distance() As Double
        Get
            Return Me.prop_AntennaTopGirtGageG1Distance
        End Get
        Set
            Me.prop_AntennaTopGirtGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennabotgirtgageg1Distance")>
    Public Property AntennaBotGirtGageG1Distance() As Double
        Get
            Return Me.prop_AntennaBotGirtGageG1Distance
        End Get
        Set
            Me.prop_AntennaBotGirtGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennainnergirtgageg1Distance")>
    Public Property AntennaInnerGirtGageG1Distance() As Double
        Get
            Return Me.prop_AntennaInnerGirtGageG1Distance
        End Get
        Set
            Me.prop_AntennaInnerGirtGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennahorizontalgageg1Distance")>
    Public Property AntennaHorizontalGageG1Distance() As Double
        Get
            Return Me.prop_AntennaHorizontalGageG1Distance
        End Get
        Set
            Me.prop_AntennaHorizontalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennashorthorizontalgageg1Distance")>
    Public Property AntennaShortHorizontalGageG1Distance() As Double
        Get
            Return Me.prop_AntennaShortHorizontalGageG1Distance
        End Get
        Set
            Me.prop_AntennaShortHorizontalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthorizontalboltgrade")>
    Public Property AntennaRedundantHorizontalBoltGrade() As String
        Get
            Return Me.prop_AntennaRedundantHorizontalBoltGrade
        End Get
        Set
            Me.prop_AntennaRedundantHorizontalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthorizontalboltsize")>
    Public Property AntennaRedundantHorizontalBoltSize() As Double
        Get
            Return Me.prop_AntennaRedundantHorizontalBoltSize
        End Get
        Set
            Me.prop_AntennaRedundantHorizontalBoltSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthorizontalnumbolts")>
    Public Property AntennaRedundantHorizontalNumBolts() As Integer
        Get
            Return Me.prop_AntennaRedundantHorizontalNumBolts
        End Get
        Set
            Me.prop_AntennaRedundantHorizontalNumBolts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthorizontalboltedgedistance")>
    Public Property AntennaRedundantHorizontalBoltEdgeDistance() As Double
        Get
            Return Me.prop_AntennaRedundantHorizontalBoltEdgeDistance
        End Get
        Set
            Me.prop_AntennaRedundantHorizontalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthorizontalgageg1Distance")>
    Public Property AntennaRedundantHorizontalGageG1Distance() As Double
        Get
            Return Me.prop_AntennaRedundantHorizontalGageG1Distance
        End Get
        Set
            Me.prop_AntennaRedundantHorizontalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthorizontalnetwidthdeduct")>
    Public Property AntennaRedundantHorizontalNetWidthDeduct() As Double
        Get
            Return Me.prop_AntennaRedundantHorizontalNetWidthDeduct
        End Get
        Set
            Me.prop_AntennaRedundantHorizontalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthorizontalufactor")>
    Public Property AntennaRedundantHorizontalUFactor() As Double
        Get
            Return Me.prop_AntennaRedundantHorizontalUFactor
        End Get
        Set
            Me.prop_AntennaRedundantHorizontalUFactor = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantdiagonalboltgrade")>
    Public Property AntennaRedundantDiagonalBoltGrade() As String
        Get
            Return Me.prop_AntennaRedundantDiagonalBoltGrade
        End Get
        Set
            Me.prop_AntennaRedundantDiagonalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantdiagonalboltsize")>
    Public Property AntennaRedundantDiagonalBoltSize() As Double
        Get
            Return Me.prop_AntennaRedundantDiagonalBoltSize
        End Get
        Set
            Me.prop_AntennaRedundantDiagonalBoltSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantdiagonalnumbolts")>
    Public Property AntennaRedundantDiagonalNumBolts() As Integer
        Get
            Return Me.prop_AntennaRedundantDiagonalNumBolts
        End Get
        Set
            Me.prop_AntennaRedundantDiagonalNumBolts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantdiagonalboltedgedistance")>
    Public Property AntennaRedundantDiagonalBoltEdgeDistance() As Double
        Get
            Return Me.prop_AntennaRedundantDiagonalBoltEdgeDistance
        End Get
        Set
            Me.prop_AntennaRedundantDiagonalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantdiagonalgageg1Distance")>
    Public Property AntennaRedundantDiagonalGageG1Distance() As Double
        Get
            Return Me.prop_AntennaRedundantDiagonalGageG1Distance
        End Get
        Set
            Me.prop_AntennaRedundantDiagonalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantdiagonalnetwidthdeduct")>
    Public Property AntennaRedundantDiagonalNetWidthDeduct() As Double
        Get
            Return Me.prop_AntennaRedundantDiagonalNetWidthDeduct
        End Get
        Set
            Me.prop_AntennaRedundantDiagonalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantdiagonalufactor")>
    Public Property AntennaRedundantDiagonalUFactor() As Double
        Get
            Return Me.prop_AntennaRedundantDiagonalUFactor
        End Get
        Set
            Me.prop_AntennaRedundantDiagonalUFactor = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubdiagonalboltgrade")>
    Public Property AntennaRedundantSubDiagonalBoltGrade() As String
        Get
            Return Me.prop_AntennaRedundantSubDiagonalBoltGrade
        End Get
        Set
            Me.prop_AntennaRedundantSubDiagonalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubdiagonalboltsize")>
    Public Property AntennaRedundantSubDiagonalBoltSize() As Double
        Get
            Return Me.prop_AntennaRedundantSubDiagonalBoltSize
        End Get
        Set
            Me.prop_AntennaRedundantSubDiagonalBoltSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubdiagonalnumbolts")>
    Public Property AntennaRedundantSubDiagonalNumBolts() As Integer
        Get
            Return Me.prop_AntennaRedundantSubDiagonalNumBolts
        End Get
        Set
            Me.prop_AntennaRedundantSubDiagonalNumBolts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubdiagonalboltedgedistance")>
    Public Property AntennaRedundantSubDiagonalBoltEdgeDistance() As Double
        Get
            Return Me.prop_AntennaRedundantSubDiagonalBoltEdgeDistance
        End Get
        Set
            Me.prop_AntennaRedundantSubDiagonalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubdiagonalgageg1Distance")>
    Public Property AntennaRedundantSubDiagonalGageG1Distance() As Double
        Get
            Return Me.prop_AntennaRedundantSubDiagonalGageG1Distance
        End Get
        Set
            Me.prop_AntennaRedundantSubDiagonalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubdiagonalnetwidthdeduct")>
    Public Property AntennaRedundantSubDiagonalNetWidthDeduct() As Double
        Get
            Return Me.prop_AntennaRedundantSubDiagonalNetWidthDeduct
        End Get
        Set
            Me.prop_AntennaRedundantSubDiagonalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubdiagonalufactor")>
    Public Property AntennaRedundantSubDiagonalUFactor() As Double
        Get
            Return Me.prop_AntennaRedundantSubDiagonalUFactor
        End Get
        Set
            Me.prop_AntennaRedundantSubDiagonalUFactor = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubhorizontalboltgrade")>
    Public Property AntennaRedundantSubHorizontalBoltGrade() As String
        Get
            Return Me.prop_AntennaRedundantSubHorizontalBoltGrade
        End Get
        Set
            Me.prop_AntennaRedundantSubHorizontalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubhorizontalboltsize")>
    Public Property AntennaRedundantSubHorizontalBoltSize() As Double
        Get
            Return Me.prop_AntennaRedundantSubHorizontalBoltSize
        End Get
        Set
            Me.prop_AntennaRedundantSubHorizontalBoltSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubhorizontalnumbolts")>
    Public Property AntennaRedundantSubHorizontalNumBolts() As Integer
        Get
            Return Me.prop_AntennaRedundantSubHorizontalNumBolts
        End Get
        Set
            Me.prop_AntennaRedundantSubHorizontalNumBolts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubhorizontalboltedgedistance")>
    Public Property AntennaRedundantSubHorizontalBoltEdgeDistance() As Double
        Get
            Return Me.prop_AntennaRedundantSubHorizontalBoltEdgeDistance
        End Get
        Set
            Me.prop_AntennaRedundantSubHorizontalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubhorizontalgageg1Distance")>
    Public Property AntennaRedundantSubHorizontalGageG1Distance() As Double
        Get
            Return Me.prop_AntennaRedundantSubHorizontalGageG1Distance
        End Get
        Set
            Me.prop_AntennaRedundantSubHorizontalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubhorizontalnetwidthdeduct")>
    Public Property AntennaRedundantSubHorizontalNetWidthDeduct() As Double
        Get
            Return Me.prop_AntennaRedundantSubHorizontalNetWidthDeduct
        End Get
        Set
            Me.prop_AntennaRedundantSubHorizontalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantsubhorizontalufactor")>
    Public Property AntennaRedundantSubHorizontalUFactor() As Double
        Get
            Return Me.prop_AntennaRedundantSubHorizontalUFactor
        End Get
        Set
            Me.prop_AntennaRedundantSubHorizontalUFactor = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantverticalboltgrade")>
    Public Property AntennaRedundantVerticalBoltGrade() As String
        Get
            Return Me.prop_AntennaRedundantVerticalBoltGrade
        End Get
        Set
            Me.prop_AntennaRedundantVerticalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantverticalboltsize")>
    Public Property AntennaRedundantVerticalBoltSize() As Double
        Get
            Return Me.prop_AntennaRedundantVerticalBoltSize
        End Get
        Set
            Me.prop_AntennaRedundantVerticalBoltSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantverticalnumbolts")>
    Public Property AntennaRedundantVerticalNumBolts() As Integer
        Get
            Return Me.prop_AntennaRedundantVerticalNumBolts
        End Get
        Set
            Me.prop_AntennaRedundantVerticalNumBolts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantverticalboltedgedistance")>
    Public Property AntennaRedundantVerticalBoltEdgeDistance() As Double
        Get
            Return Me.prop_AntennaRedundantVerticalBoltEdgeDistance
        End Get
        Set
            Me.prop_AntennaRedundantVerticalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantverticalgageg1Distance")>
    Public Property AntennaRedundantVerticalGageG1Distance() As Double
        Get
            Return Me.prop_AntennaRedundantVerticalGageG1Distance
        End Get
        Set
            Me.prop_AntennaRedundantVerticalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantverticalnetwidthdeduct")>
    Public Property AntennaRedundantVerticalNetWidthDeduct() As Double
        Get
            Return Me.prop_AntennaRedundantVerticalNetWidthDeduct
        End Get
        Set
            Me.prop_AntennaRedundantVerticalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundantverticalufactor")>
    Public Property AntennaRedundantVerticalUFactor() As Double
        Get
            Return Me.prop_AntennaRedundantVerticalUFactor
        End Get
        Set
            Me.prop_AntennaRedundantVerticalUFactor = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipboltgrade")>
    Public Property AntennaRedundantHipBoltGrade() As String
        Get
            Return Me.prop_AntennaRedundantHipBoltGrade
        End Get
        Set
            Me.prop_AntennaRedundantHipBoltGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipboltsize")>
    Public Property AntennaRedundantHipBoltSize() As Double
        Get
            Return Me.prop_AntennaRedundantHipBoltSize
        End Get
        Set
            Me.prop_AntennaRedundantHipBoltSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipnumbolts")>
    Public Property AntennaRedundantHipNumBolts() As Integer
        Get
            Return Me.prop_AntennaRedundantHipNumBolts
        End Get
        Set
            Me.prop_AntennaRedundantHipNumBolts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipboltedgedistance")>
    Public Property AntennaRedundantHipBoltEdgeDistance() As Double
        Get
            Return Me.prop_AntennaRedundantHipBoltEdgeDistance
        End Get
        Set
            Me.prop_AntennaRedundantHipBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipgageg1Distance")>
    Public Property AntennaRedundantHipGageG1Distance() As Double
        Get
            Return Me.prop_AntennaRedundantHipGageG1Distance
        End Get
        Set
            Me.prop_AntennaRedundantHipGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipnetwidthdeduct")>
    Public Property AntennaRedundantHipNetWidthDeduct() As Double
        Get
            Return Me.prop_AntennaRedundantHipNetWidthDeduct
        End Get
        Set
            Me.prop_AntennaRedundantHipNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipufactor")>
    Public Property AntennaRedundantHipUFactor() As Double
        Get
            Return Me.prop_AntennaRedundantHipUFactor
        End Get
        Set
            Me.prop_AntennaRedundantHipUFactor = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipdiagonalboltgrade")>
    Public Property AntennaRedundantHipDiagonalBoltGrade() As String
        Get
            Return Me.prop_AntennaRedundantHipDiagonalBoltGrade
        End Get
        Set
            Me.prop_AntennaRedundantHipDiagonalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipdiagonalboltsize")>
    Public Property AntennaRedundantHipDiagonalBoltSize() As Double
        Get
            Return Me.prop_AntennaRedundantHipDiagonalBoltSize
        End Get
        Set
            Me.prop_AntennaRedundantHipDiagonalBoltSize = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipdiagonalnumbolts")>
    Public Property AntennaRedundantHipDiagonalNumBolts() As Integer
        Get
            Return Me.prop_AntennaRedundantHipDiagonalNumBolts
        End Get
        Set
            Me.prop_AntennaRedundantHipDiagonalNumBolts = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipdiagonalboltedgedistance")>
    Public Property AntennaRedundantHipDiagonalBoltEdgeDistance() As Double
        Get
            Return Me.prop_AntennaRedundantHipDiagonalBoltEdgeDistance
        End Get
        Set
            Me.prop_AntennaRedundantHipDiagonalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipdiagonalgageg1Distance")>
    Public Property AntennaRedundantHipDiagonalGageG1Distance() As Double
        Get
            Return Me.prop_AntennaRedundantHipDiagonalGageG1Distance
        End Get
        Set
            Me.prop_AntennaRedundantHipDiagonalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipdiagonalnetwidthdeduct")>
    Public Property AntennaRedundantHipDiagonalNetWidthDeduct() As Double
        Get
            Return Me.prop_AntennaRedundantHipDiagonalNetWidthDeduct
        End Get
        Set
            Me.prop_AntennaRedundantHipDiagonalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennaredundanthipdiagonalufactor")>
    Public Property AntennaRedundantHipDiagonalUFactor() As Double
        Get
            Return Me.prop_AntennaRedundantHipDiagonalUFactor
        End Get
        Set
            Me.prop_AntennaRedundantHipDiagonalUFactor = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagonaloutofplanerestraint")>
    Public Property AntennaDiagonalOutOfPlaneRestraint() As Boolean
        Get
            Return Me.prop_AntennaDiagonalOutOfPlaneRestraint
        End Get
        Set
            Me.prop_AntennaDiagonalOutOfPlaneRestraint = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennatopgirtoutofplanerestraint")>
    Public Property AntennaTopGirtOutOfPlaneRestraint() As Boolean
        Get
            Return Me.prop_AntennaTopGirtOutOfPlaneRestraint
        End Get
        Set
            Me.prop_AntennaTopGirtOutOfPlaneRestraint = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennabottomgirtoutofplanerestraint")>
    Public Property AntennaBottomGirtOutOfPlaneRestraint() As Boolean
        Get
            Return Me.prop_AntennaBottomGirtOutOfPlaneRestraint
        End Get
        Set
            Me.prop_AntennaBottomGirtOutOfPlaneRestraint = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennamidgirtoutofplanerestraint")>
    Public Property AntennaMidGirtOutOfPlaneRestraint() As Boolean
        Get
            Return Me.prop_AntennaMidGirtOutOfPlaneRestraint
        End Get
        Set
            Me.prop_AntennaMidGirtOutOfPlaneRestraint = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennahorizontaloutofplanerestraint")>
    Public Property AntennaHorizontalOutOfPlaneRestraint() As Boolean
        Get
            Return Me.prop_AntennaHorizontalOutOfPlaneRestraint
        End Get
        Set
            Me.prop_AntennaHorizontalOutOfPlaneRestraint = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennasecondaryhorizontaloutofplanerestraint")>
    Public Property AntennaSecondaryHorizontalOutOfPlaneRestraint() As Boolean
        Get
            Return Me.prop_AntennaSecondaryHorizontalOutOfPlaneRestraint
        End Get
        Set
            Me.prop_AntennaSecondaryHorizontalOutOfPlaneRestraint = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagoffsetney")>
    Public Property AntennaDiagOffsetNEY() As Double
        Get
            Return Me.prop_AntennaDiagOffsetNEY
        End Get
        Set
            Me.prop_AntennaDiagOffsetNEY = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagoffsetnex")>
    Public Property AntennaDiagOffsetNEX() As Double
        Get
            Return Me.prop_AntennaDiagOffsetNEX
        End Get
        Set
            Me.prop_AntennaDiagOffsetNEX = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagoffsetpey")>
    Public Property AntennaDiagOffsetPEY() As Double
        Get
            Return Me.prop_AntennaDiagOffsetPEY
        End Get
        Set
            Me.prop_AntennaDiagOffsetPEY = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennadiagoffsetpex")>
    Public Property AntennaDiagOffsetPEX() As Double
        Get
            Return Me.prop_AntennaDiagOffsetPEX
        End Get
        Set
            Me.prop_AntennaDiagOffsetPEX = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakbraceoffsetney")>
    Public Property AntennaKbraceOffsetNEY() As Double
        Get
            Return Me.prop_AntennaKbraceOffsetNEY
        End Get
        Set
            Me.prop_AntennaKbraceOffsetNEY = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakbraceoffsetnex")>
    Public Property AntennaKbraceOffsetNEX() As Double
        Get
            Return Me.prop_AntennaKbraceOffsetNEX
        End Get
        Set
            Me.prop_AntennaKbraceOffsetNEX = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakbraceoffsetpey")>
    Public Property AntennaKbraceOffsetPEY() As Double
        Get
            Return Me.prop_AntennaKbraceOffsetPEY
        End Get
        Set
            Me.prop_AntennaKbraceOffsetPEY = Value
        End Set
    End Property
    <Category("TNX Antenna Record"), Description(""), DisplayName("Antennakbraceoffsetpex")>
    Public Property AntennaKbraceOffsetPEX() As Double
        Get
            Return Me.prop_AntennaKbraceOffsetPEX
        End Get
        Set
            Me.prop_AntennaKbraceOffsetPEX = Value
        End Set
    End Property

End Class

Partial Public Class tnxTowerRecord
    'base structure
    Private prop_ID As Integer
    Private prop_tnxID As Integer
    Private prop_TowerRec As Integer
    Private prop_TowerDatabase As String
    Private prop_TowerName As String
    Private prop_TowerHeight As Double
    Private prop_TowerFaceWidth As Double
    Private prop_TowerNumSections As Integer
    Private prop_TowerSectionLength As Double
    Private prop_TowerDiagonalSpacing As Double
    Private prop_TowerDiagonalSpacingEx As Double
    Private prop_TowerBraceType As String
    Private prop_TowerFaceBevel As Double
    Private prop_TowerTopGirtOffset As Double
    Private prop_TowerBotGirtOffset As Double
    Private prop_TowerHasKBraceEndPanels As Boolean
    Private prop_TowerHasHorizontals As Boolean
    Private prop_TowerLegType As String
    Private prop_TowerLegSize As String
    Private prop_TowerLegGrade As Double
    Private prop_TowerLegMatlGrade As String
    Private prop_TowerDiagonalGrade As Double
    Private prop_TowerDiagonalMatlGrade As String
    Private prop_TowerInnerBracingGrade As Double
    Private prop_TowerInnerBracingMatlGrade As String
    Private prop_TowerTopGirtGrade As Double
    Private prop_TowerTopGirtMatlGrade As String
    Private prop_TowerBotGirtGrade As Double
    Private prop_TowerBotGirtMatlGrade As String
    Private prop_TowerInnerGirtGrade As Double
    Private prop_TowerInnerGirtMatlGrade As String
    Private prop_TowerLongHorizontalGrade As Double
    Private prop_TowerLongHorizontalMatlGrade As String
    Private prop_TowerShortHorizontalGrade As Double
    Private prop_TowerShortHorizontalMatlGrade As String
    Private prop_TowerDiagonalType As String
    Private prop_TowerDiagonalSize As String
    Private prop_TowerInnerBracingType As String
    Private prop_TowerInnerBracingSize As String
    Private prop_TowerTopGirtType As String
    Private prop_TowerTopGirtSize As String
    Private prop_TowerBotGirtType As String
    Private prop_TowerBotGirtSize As String
    Private prop_TowerNumInnerGirts As Integer
    Private prop_TowerInnerGirtType As String
    Private prop_TowerInnerGirtSize As String
    Private prop_TowerLongHorizontalType As String
    Private prop_TowerLongHorizontalSize As String
    Private prop_TowerShortHorizontalType As String
    Private prop_TowerShortHorizontalSize As String
    Private prop_TowerRedundantGrade As Double
    Private prop_TowerRedundantMatlGrade As String
    Private prop_TowerRedundantType As String
    Private prop_TowerRedundantDiagType As String
    Private prop_TowerRedundantSubDiagonalType As String
    Private prop_TowerRedundantSubHorizontalType As String
    Private prop_TowerRedundantVerticalType As String
    Private prop_TowerRedundantHipType As String
    Private prop_TowerRedundantHipDiagonalType As String
    Private prop_TowerRedundantHorizontalSize As String
    Private prop_TowerRedundantHorizontalSize2 As String
    Private prop_TowerRedundantHorizontalSize3 As String
    Private prop_TowerRedundantHorizontalSize4 As String
    Private prop_TowerRedundantDiagonalSize As String
    Private prop_TowerRedundantDiagonalSize2 As String
    Private prop_TowerRedundantDiagonalSize3 As String
    Private prop_TowerRedundantDiagonalSize4 As String
    Private prop_TowerRedundantSubHorizontalSize As String
    Private prop_TowerRedundantSubDiagonalSize As String
    Private prop_TowerSubDiagLocation As Double
    Private prop_TowerRedundantVerticalSize As String
    Private prop_TowerRedundantHipSize As String
    Private prop_TowerRedundantHipSize2 As String
    Private prop_TowerRedundantHipSize3 As String
    Private prop_TowerRedundantHipSize4 As String
    Private prop_TowerRedundantHipDiagonalSize As String
    Private prop_TowerRedundantHipDiagonalSize2 As String
    Private prop_TowerRedundantHipDiagonalSize3 As String
    Private prop_TowerRedundantHipDiagonalSize4 As String
    Private prop_TowerSWMult As Double
    Private prop_TowerWPMult As Double
    Private prop_TowerAutoCalcKSingleAngle As Boolean
    Private prop_TowerAutoCalcKSolidRound As Boolean
    Private prop_TowerAfGusset As Double
    Private prop_TowerTfGusset As Double
    Private prop_TowerGussetBoltEdgeDistance As Double
    Private prop_TowerGussetGrade As Double
    Private prop_TowerGussetMatlGrade As String
    Private prop_TowerAfMult As Double
    Private prop_TowerArMult As Double
    Private prop_TowerFlatIPAPole As Double
    Private prop_TowerRoundIPAPole As Double
    Private prop_TowerFlatIPALeg As Double
    Private prop_TowerRoundIPALeg As Double
    Private prop_TowerFlatIPAHorizontal As Double
    Private prop_TowerRoundIPAHorizontal As Double
    Private prop_TowerFlatIPADiagonal As Double
    Private prop_TowerRoundIPADiagonal As Double
    Private prop_TowerCSA_S37_SpeedUpFactor As Double
    Private prop_TowerKLegs As Double
    Private prop_TowerKXBracedDiags As Double
    Private prop_TowerKKBracedDiags As Double
    Private prop_TowerKZBracedDiags As Double
    Private prop_TowerKHorzs As Double
    Private prop_TowerKSecHorzs As Double
    Private prop_TowerKGirts As Double
    Private prop_TowerKInners As Double
    Private prop_TowerKXBracedDiagsY As Double
    Private prop_TowerKKBracedDiagsY As Double
    Private prop_TowerKZBracedDiagsY As Double
    Private prop_TowerKHorzsY As Double
    Private prop_TowerKSecHorzsY As Double
    Private prop_TowerKGirtsY As Double
    Private prop_TowerKInnersY As Double
    Private prop_TowerKRedHorz As Double
    Private prop_TowerKRedDiag As Double
    Private prop_TowerKRedSubDiag As Double
    Private prop_TowerKRedSubHorz As Double
    Private prop_TowerKRedVert As Double
    Private prop_TowerKRedHip As Double
    Private prop_TowerKRedHipDiag As Double
    Private prop_TowerKTLX As Double
    Private prop_TowerKTLZ As Double
    Private prop_TowerKTLLeg As Double
    Private prop_TowerInnerKTLX As Double
    Private prop_TowerInnerKTLZ As Double
    Private prop_TowerInnerKTLLeg As Double
    Private prop_TowerStitchBoltLocationHoriz As String
    Private prop_TowerStitchBoltLocationDiag As String
    Private prop_TowerStitchBoltLocationRed As String
    Private prop_TowerStitchSpacing As Double
    Private prop_TowerStitchSpacingDiag As Double
    Private prop_TowerStitchSpacingHorz As Double
    Private prop_TowerStitchSpacingRed As Double
    Private prop_TowerLegNetWidthDeduct As Double
    Private prop_TowerLegUFactor As Double
    Private prop_TowerDiagonalNetWidthDeduct As Double
    Private prop_TowerTopGirtNetWidthDeduct As Double
    Private prop_TowerBotGirtNetWidthDeduct As Double
    Private prop_TowerInnerGirtNetWidthDeduct As Double
    Private prop_TowerHorizontalNetWidthDeduct As Double
    Private prop_TowerShortHorizontalNetWidthDeduct As Double
    Private prop_TowerDiagonalUFactor As Double
    Private prop_TowerTopGirtUFactor As Double
    Private prop_TowerBotGirtUFactor As Double
    Private prop_TowerInnerGirtUFactor As Double
    Private prop_TowerHorizontalUFactor As Double
    Private prop_TowerShortHorizontalUFactor As Double
    Private prop_TowerLegConnType As String
    Private prop_TowerLegNumBolts As Integer
    Private prop_TowerDiagonalNumBolts As Integer
    Private prop_TowerTopGirtNumBolts As Integer
    Private prop_TowerBotGirtNumBolts As Integer
    Private prop_TowerInnerGirtNumBolts As Integer
    Private prop_TowerHorizontalNumBolts As Integer
    Private prop_TowerShortHorizontalNumBolts As Integer
    Private prop_TowerLegBoltGrade As String
    Private prop_TowerLegBoltSize As Double
    Private prop_TowerDiagonalBoltGrade As String
    Private prop_TowerDiagonalBoltSize As Double
    Private prop_TowerTopGirtBoltGrade As String
    Private prop_TowerTopGirtBoltSize As Double
    Private prop_TowerBotGirtBoltGrade As String
    Private prop_TowerBotGirtBoltSize As Double
    Private prop_TowerInnerGirtBoltGrade As String
    Private prop_TowerInnerGirtBoltSize As Double
    Private prop_TowerHorizontalBoltGrade As String
    Private prop_TowerHorizontalBoltSize As Double
    Private prop_TowerShortHorizontalBoltGrade As String
    Private prop_TowerShortHorizontalBoltSize As Double
    Private prop_TowerLegBoltEdgeDistance As Double
    Private prop_TowerDiagonalBoltEdgeDistance As Double
    Private prop_TowerTopGirtBoltEdgeDistance As Double
    Private prop_TowerBotGirtBoltEdgeDistance As Double
    Private prop_TowerInnerGirtBoltEdgeDistance As Double
    Private prop_TowerHorizontalBoltEdgeDistance As Double
    Private prop_TowerShortHorizontalBoltEdgeDistance As Double
    Private prop_TowerDiagonalGageG1Distance As Double
    Private prop_TowerTopGirtGageG1Distance As Double
    Private prop_TowerBotGirtGageG1Distance As Double
    Private prop_TowerInnerGirtGageG1Distance As Double
    Private prop_TowerHorizontalGageG1Distance As Double
    Private prop_TowerShortHorizontalGageG1Distance As Double
    Private prop_TowerRedundantHorizontalBoltGrade As String
    Private prop_TowerRedundantHorizontalBoltSize As Double
    Private prop_TowerRedundantHorizontalNumBolts As Integer
    Private prop_TowerRedundantHorizontalBoltEdgeDistance As Double
    Private prop_TowerRedundantHorizontalGageG1Distance As Double
    Private prop_TowerRedundantHorizontalNetWidthDeduct As Double
    Private prop_TowerRedundantHorizontalUFactor As Double
    Private prop_TowerRedundantDiagonalBoltGrade As String
    Private prop_TowerRedundantDiagonalBoltSize As Double
    Private prop_TowerRedundantDiagonalNumBolts As Integer
    Private prop_TowerRedundantDiagonalBoltEdgeDistance As Double
    Private prop_TowerRedundantDiagonalGageG1Distance As Double
    Private prop_TowerRedundantDiagonalNetWidthDeduct As Double
    Private prop_TowerRedundantDiagonalUFactor As Double
    Private prop_TowerRedundantSubDiagonalBoltGrade As String
    Private prop_TowerRedundantSubDiagonalBoltSize As Double
    Private prop_TowerRedundantSubDiagonalNumBolts As Integer
    Private prop_TowerRedundantSubDiagonalBoltEdgeDistance As Double
    Private prop_TowerRedundantSubDiagonalGageG1Distance As Double
    Private prop_TowerRedundantSubDiagonalNetWidthDeduct As Double
    Private prop_TowerRedundantSubDiagonalUFactor As Double
    Private prop_TowerRedundantSubHorizontalBoltGrade As String
    Private prop_TowerRedundantSubHorizontalBoltSize As Double
    Private prop_TowerRedundantSubHorizontalNumBolts As Integer
    Private prop_TowerRedundantSubHorizontalBoltEdgeDistance As Double
    Private prop_TowerRedundantSubHorizontalGageG1Distance As Double
    Private prop_TowerRedundantSubHorizontalNetWidthDeduct As Double
    Private prop_TowerRedundantSubHorizontalUFactor As Double
    Private prop_TowerRedundantVerticalBoltGrade As String
    Private prop_TowerRedundantVerticalBoltSize As Double
    Private prop_TowerRedundantVerticalNumBolts As Integer
    Private prop_TowerRedundantVerticalBoltEdgeDistance As Double
    Private prop_TowerRedundantVerticalGageG1Distance As Double
    Private prop_TowerRedundantVerticalNetWidthDeduct As Double
    Private prop_TowerRedundantVerticalUFactor As Double
    Private prop_TowerRedundantHipBoltGrade As String
    Private prop_TowerRedundantHipBoltSize As Double
    Private prop_TowerRedundantHipNumBolts As Integer
    Private prop_TowerRedundantHipBoltEdgeDistance As Double
    Private prop_TowerRedundantHipGageG1Distance As Double
    Private prop_TowerRedundantHipNetWidthDeduct As Double
    Private prop_TowerRedundantHipUFactor As Double
    Private prop_TowerRedundantHipDiagonalBoltGrade As String
    Private prop_TowerRedundantHipDiagonalBoltSize As Double
    Private prop_TowerRedundantHipDiagonalNumBolts As Integer
    Private prop_TowerRedundantHipDiagonalBoltEdgeDistance As Double
    Private prop_TowerRedundantHipDiagonalGageG1Distance As Double
    Private prop_TowerRedundantHipDiagonalNetWidthDeduct As Double
    Private prop_TowerRedundantHipDiagonalUFactor As Double
    Private prop_TowerDiagonalOutOfPlaneRestraint As Boolean
    Private prop_TowerTopGirtOutOfPlaneRestraint As Boolean
    Private prop_TowerBottomGirtOutOfPlaneRestraint As Boolean
    Private prop_TowerMidGirtOutOfPlaneRestraint As Boolean
    Private prop_TowerHorizontalOutOfPlaneRestraint As Boolean
    Private prop_TowerSecondaryHorizontalOutOfPlaneRestraint As Boolean
    Private prop_TowerUniqueFlag As Integer
    Private prop_TowerDiagOffsetNEY As Double
    Private prop_TowerDiagOffsetNEX As Double
    Private prop_TowerDiagOffsetPEY As Double
    Private prop_TowerDiagOffsetPEX As Double
    Private prop_TowerKbraceOffsetNEY As Double
    Private prop_TowerKbraceOffsetNEX As Double
    Private prop_TowerKbraceOffsetPEY As Double
    Private prop_TowerKbraceOffsetPEX As Double

    <Category("TNX Tower Record"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer
        Get
            Return Me.prop_ID
        End Get
        Set
            Me.prop_ID = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Tnxid")>
    Public Property tnxID() As Integer
        Get
            Return Me.prop_tnxID
        End Get
        Set
            Me.prop_tnxID = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerrec")>
    Public Property TowerRec() As Integer
        Get
            Return Me.prop_TowerRec
        End Get
        Set
            Me.prop_TowerRec = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdatabase")>
    Public Property TowerDatabase() As String
        Get
            Return Me.prop_TowerDatabase
        End Get
        Set
            Me.prop_TowerDatabase = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towername")>
    Public Property TowerName() As String
        Get
            Return Me.prop_TowerName
        End Get
        Set
            Me.prop_TowerName = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerheight")>
    Public Property TowerHeight() As Double
        Get
            Return Me.prop_TowerHeight
        End Get
        Set
            Me.prop_TowerHeight = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerfacewidth")>
    Public Property TowerFaceWidth() As Double
        Get
            Return Me.prop_TowerFaceWidth
        End Get
        Set
            Me.prop_TowerFaceWidth = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towernumsections")>
    Public Property TowerNumSections() As Integer
        Get
            Return Me.prop_TowerNumSections
        End Get
        Set
            Me.prop_TowerNumSections = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towersectionlength")>
    Public Property TowerSectionLength() As Double
        Get
            Return Me.prop_TowerSectionLength
        End Get
        Set
            Me.prop_TowerSectionLength = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagonalspacing")>
    Public Property TowerDiagonalSpacing() As Double
        Get
            Return Me.prop_TowerDiagonalSpacing
        End Get
        Set
            Me.prop_TowerDiagonalSpacing = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagonalspacingex")>
    Public Property TowerDiagonalSpacingEx() As Double
        Get
            Return Me.prop_TowerDiagonalSpacingEx
        End Get
        Set
            Me.prop_TowerDiagonalSpacingEx = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerbracetype")>
    Public Property TowerBraceType() As String
        Get
            Return Me.prop_TowerBraceType
        End Get
        Set
            Me.prop_TowerBraceType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerfacebevel")>
    Public Property TowerFaceBevel() As Double
        Get
            Return Me.prop_TowerFaceBevel
        End Get
        Set
            Me.prop_TowerFaceBevel = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towertopgirtoffset")>
    Public Property TowerTopGirtOffset() As Double
        Get
            Return Me.prop_TowerTopGirtOffset
        End Get
        Set
            Me.prop_TowerTopGirtOffset = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerbotgirtoffset")>
    Public Property TowerBotGirtOffset() As Double
        Get
            Return Me.prop_TowerBotGirtOffset
        End Get
        Set
            Me.prop_TowerBotGirtOffset = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerhaskbraceendpanels")>
    Public Property TowerHasKBraceEndPanels() As Boolean
        Get
            Return Me.prop_TowerHasKBraceEndPanels
        End Get
        Set
            Me.prop_TowerHasKBraceEndPanels = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerhashorizontals")>
    Public Property TowerHasHorizontals() As Boolean
        Get
            Return Me.prop_TowerHasHorizontals
        End Get
        Set
            Me.prop_TowerHasHorizontals = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerlegtype")>
    Public Property TowerLegType() As String
        Get
            Return Me.prop_TowerLegType
        End Get
        Set
            Me.prop_TowerLegType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerlegsize")>
    Public Property TowerLegSize() As String
        Get
            Return Me.prop_TowerLegSize
        End Get
        Set
            Me.prop_TowerLegSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerleggrade")>
    Public Property TowerLegGrade() As Double
        Get
            Return Me.prop_TowerLegGrade
        End Get
        Set
            Me.prop_TowerLegGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerlegmatlgrade")>
    Public Property TowerLegMatlGrade() As String
        Get
            Return Me.prop_TowerLegMatlGrade
        End Get
        Set
            Me.prop_TowerLegMatlGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagonalgrade")>
    Public Property TowerDiagonalGrade() As Double
        Get
            Return Me.prop_TowerDiagonalGrade
        End Get
        Set
            Me.prop_TowerDiagonalGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagonalmatlgrade")>
    Public Property TowerDiagonalMatlGrade() As String
        Get
            Return Me.prop_TowerDiagonalMatlGrade
        End Get
        Set
            Me.prop_TowerDiagonalMatlGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnerbracinggrade")>
    Public Property TowerInnerBracingGrade() As Double
        Get
            Return Me.prop_TowerInnerBracingGrade
        End Get
        Set
            Me.prop_TowerInnerBracingGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnerbracingmatlgrade")>
    Public Property TowerInnerBracingMatlGrade() As String
        Get
            Return Me.prop_TowerInnerBracingMatlGrade
        End Get
        Set
            Me.prop_TowerInnerBracingMatlGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towertopgirtgrade")>
    Public Property TowerTopGirtGrade() As Double
        Get
            Return Me.prop_TowerTopGirtGrade
        End Get
        Set
            Me.prop_TowerTopGirtGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towertopgirtmatlgrade")>
    Public Property TowerTopGirtMatlGrade() As String
        Get
            Return Me.prop_TowerTopGirtMatlGrade
        End Get
        Set
            Me.prop_TowerTopGirtMatlGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerbotgirtgrade")>
    Public Property TowerBotGirtGrade() As Double
        Get
            Return Me.prop_TowerBotGirtGrade
        End Get
        Set
            Me.prop_TowerBotGirtGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerbotgirtmatlgrade")>
    Public Property TowerBotGirtMatlGrade() As String
        Get
            Return Me.prop_TowerBotGirtMatlGrade
        End Get
        Set
            Me.prop_TowerBotGirtMatlGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnergirtgrade")>
    Public Property TowerInnerGirtGrade() As Double
        Get
            Return Me.prop_TowerInnerGirtGrade
        End Get
        Set
            Me.prop_TowerInnerGirtGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnergirtmatlgrade")>
    Public Property TowerInnerGirtMatlGrade() As String
        Get
            Return Me.prop_TowerInnerGirtMatlGrade
        End Get
        Set
            Me.prop_TowerInnerGirtMatlGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerlonghorizontalgrade")>
    Public Property TowerLongHorizontalGrade() As Double
        Get
            Return Me.prop_TowerLongHorizontalGrade
        End Get
        Set
            Me.prop_TowerLongHorizontalGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerlonghorizontalmatlgrade")>
    Public Property TowerLongHorizontalMatlGrade() As String
        Get
            Return Me.prop_TowerLongHorizontalMatlGrade
        End Get
        Set
            Me.prop_TowerLongHorizontalMatlGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towershorthorizontalgrade")>
    Public Property TowerShortHorizontalGrade() As Double
        Get
            Return Me.prop_TowerShortHorizontalGrade
        End Get
        Set
            Me.prop_TowerShortHorizontalGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towershorthorizontalmatlgrade")>
    Public Property TowerShortHorizontalMatlGrade() As String
        Get
            Return Me.prop_TowerShortHorizontalMatlGrade
        End Get
        Set
            Me.prop_TowerShortHorizontalMatlGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagonaltype")>
    Public Property TowerDiagonalType() As String
        Get
            Return Me.prop_TowerDiagonalType
        End Get
        Set
            Me.prop_TowerDiagonalType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagonalsize")>
    Public Property TowerDiagonalSize() As String
        Get
            Return Me.prop_TowerDiagonalSize
        End Get
        Set
            Me.prop_TowerDiagonalSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnerbracingtype")>
    Public Property TowerInnerBracingType() As String
        Get
            Return Me.prop_TowerInnerBracingType
        End Get
        Set
            Me.prop_TowerInnerBracingType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnerbracingsize")>
    Public Property TowerInnerBracingSize() As String
        Get
            Return Me.prop_TowerInnerBracingSize
        End Get
        Set
            Me.prop_TowerInnerBracingSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towertopgirttype")>
    Public Property TowerTopGirtType() As String
        Get
            Return Me.prop_TowerTopGirtType
        End Get
        Set
            Me.prop_TowerTopGirtType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towertopgirtsize")>
    Public Property TowerTopGirtSize() As String
        Get
            Return Me.prop_TowerTopGirtSize
        End Get
        Set
            Me.prop_TowerTopGirtSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerbotgirttype")>
    Public Property TowerBotGirtType() As String
        Get
            Return Me.prop_TowerBotGirtType
        End Get
        Set
            Me.prop_TowerBotGirtType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerbotgirtsize")>
    Public Property TowerBotGirtSize() As String
        Get
            Return Me.prop_TowerBotGirtSize
        End Get
        Set
            Me.prop_TowerBotGirtSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towernuminnergirts")>
    Public Property TowerNumInnerGirts() As Integer
        Get
            Return Me.prop_TowerNumInnerGirts
        End Get
        Set
            Me.prop_TowerNumInnerGirts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnergirttype")>
    Public Property TowerInnerGirtType() As String
        Get
            Return Me.prop_TowerInnerGirtType
        End Get
        Set
            Me.prop_TowerInnerGirtType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnergirtsize")>
    Public Property TowerInnerGirtSize() As String
        Get
            Return Me.prop_TowerInnerGirtSize
        End Get
        Set
            Me.prop_TowerInnerGirtSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerlonghorizontaltype")>
    Public Property TowerLongHorizontalType() As String
        Get
            Return Me.prop_TowerLongHorizontalType
        End Get
        Set
            Me.prop_TowerLongHorizontalType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerlonghorizontalsize")>
    Public Property TowerLongHorizontalSize() As String
        Get
            Return Me.prop_TowerLongHorizontalSize
        End Get
        Set
            Me.prop_TowerLongHorizontalSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towershorthorizontaltype")>
    Public Property TowerShortHorizontalType() As String
        Get
            Return Me.prop_TowerShortHorizontalType
        End Get
        Set
            Me.prop_TowerShortHorizontalType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towershorthorizontalsize")>
    Public Property TowerShortHorizontalSize() As String
        Get
            Return Me.prop_TowerShortHorizontalSize
        End Get
        Set
            Me.prop_TowerShortHorizontalSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantgrade")>
    Public Property TowerRedundantGrade() As Double
        Get
            Return Me.prop_TowerRedundantGrade
        End Get
        Set
            Me.prop_TowerRedundantGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantmatlgrade")>
    Public Property TowerRedundantMatlGrade() As String
        Get
            Return Me.prop_TowerRedundantMatlGrade
        End Get
        Set
            Me.prop_TowerRedundantMatlGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanttype")>
    Public Property TowerRedundantType() As String
        Get
            Return Me.prop_TowerRedundantType
        End Get
        Set
            Me.prop_TowerRedundantType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantdiagtype")>
    Public Property TowerRedundantDiagType() As String
        Get
            Return Me.prop_TowerRedundantDiagType
        End Get
        Set
            Me.prop_TowerRedundantDiagType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubdiagonaltype")>
    Public Property TowerRedundantSubDiagonalType() As String
        Get
            Return Me.prop_TowerRedundantSubDiagonalType
        End Get
        Set
            Me.prop_TowerRedundantSubDiagonalType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubhorizontaltype")>
    Public Property TowerRedundantSubHorizontalType() As String
        Get
            Return Me.prop_TowerRedundantSubHorizontalType
        End Get
        Set
            Me.prop_TowerRedundantSubHorizontalType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantverticaltype")>
    Public Property TowerRedundantVerticalType() As String
        Get
            Return Me.prop_TowerRedundantVerticalType
        End Get
        Set
            Me.prop_TowerRedundantVerticalType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthiptype")>
    Public Property TowerRedundantHipType() As String
        Get
            Return Me.prop_TowerRedundantHipType
        End Get
        Set
            Me.prop_TowerRedundantHipType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipdiagonaltype")>
    Public Property TowerRedundantHipDiagonalType() As String
        Get
            Return Me.prop_TowerRedundantHipDiagonalType
        End Get
        Set
            Me.prop_TowerRedundantHipDiagonalType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthorizontalsize")>
    Public Property TowerRedundantHorizontalSize() As String
        Get
            Return Me.prop_TowerRedundantHorizontalSize
        End Get
        Set
            Me.prop_TowerRedundantHorizontalSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthorizontalsize2")>
    Public Property TowerRedundantHorizontalSize2() As String
        Get
            Return Me.prop_TowerRedundantHorizontalSize2
        End Get
        Set
            Me.prop_TowerRedundantHorizontalSize2 = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthorizontalsize3")>
    Public Property TowerRedundantHorizontalSize3() As String
        Get
            Return Me.prop_TowerRedundantHorizontalSize3
        End Get
        Set
            Me.prop_TowerRedundantHorizontalSize3 = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthorizontalsize4")>
    Public Property TowerRedundantHorizontalSize4() As String
        Get
            Return Me.prop_TowerRedundantHorizontalSize4
        End Get
        Set
            Me.prop_TowerRedundantHorizontalSize4 = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantdiagonalsize")>
    Public Property TowerRedundantDiagonalSize() As String
        Get
            Return Me.prop_TowerRedundantDiagonalSize
        End Get
        Set
            Me.prop_TowerRedundantDiagonalSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantdiagonalsize2")>
    Public Property TowerRedundantDiagonalSize2() As String
        Get
            Return Me.prop_TowerRedundantDiagonalSize2
        End Get
        Set
            Me.prop_TowerRedundantDiagonalSize2 = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantdiagonalsize3")>
    Public Property TowerRedundantDiagonalSize3() As String
        Get
            Return Me.prop_TowerRedundantDiagonalSize3
        End Get
        Set
            Me.prop_TowerRedundantDiagonalSize3 = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantdiagonalsize4")>
    Public Property TowerRedundantDiagonalSize4() As String
        Get
            Return Me.prop_TowerRedundantDiagonalSize4
        End Get
        Set
            Me.prop_TowerRedundantDiagonalSize4 = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubhorizontalsize")>
    Public Property TowerRedundantSubHorizontalSize() As String
        Get
            Return Me.prop_TowerRedundantSubHorizontalSize
        End Get
        Set
            Me.prop_TowerRedundantSubHorizontalSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubdiagonalsize")>
    Public Property TowerRedundantSubDiagonalSize() As String
        Get
            Return Me.prop_TowerRedundantSubDiagonalSize
        End Get
        Set
            Me.prop_TowerRedundantSubDiagonalSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towersubdiaglocation")>
    Public Property TowerSubDiagLocation() As Double
        Get
            Return Me.prop_TowerSubDiagLocation
        End Get
        Set
            Me.prop_TowerSubDiagLocation = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantverticalsize")>
    Public Property TowerRedundantVerticalSize() As String
        Get
            Return Me.prop_TowerRedundantVerticalSize
        End Get
        Set
            Me.prop_TowerRedundantVerticalSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipsize")>
    Public Property TowerRedundantHipSize() As String
        Get
            Return Me.prop_TowerRedundantHipSize
        End Get
        Set
            Me.prop_TowerRedundantHipSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipsize2")>
    Public Property TowerRedundantHipSize2() As String
        Get
            Return Me.prop_TowerRedundantHipSize2
        End Get
        Set
            Me.prop_TowerRedundantHipSize2 = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipsize3")>
    Public Property TowerRedundantHipSize3() As String
        Get
            Return Me.prop_TowerRedundantHipSize3
        End Get
        Set
            Me.prop_TowerRedundantHipSize3 = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipsize4")>
    Public Property TowerRedundantHipSize4() As String
        Get
            Return Me.prop_TowerRedundantHipSize4
        End Get
        Set
            Me.prop_TowerRedundantHipSize4 = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipdiagonalsize")>
    Public Property TowerRedundantHipDiagonalSize() As String
        Get
            Return Me.prop_TowerRedundantHipDiagonalSize
        End Get
        Set
            Me.prop_TowerRedundantHipDiagonalSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipdiagonalsize2")>
    Public Property TowerRedundantHipDiagonalSize2() As String
        Get
            Return Me.prop_TowerRedundantHipDiagonalSize2
        End Get
        Set
            Me.prop_TowerRedundantHipDiagonalSize2 = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipdiagonalsize3")>
    Public Property TowerRedundantHipDiagonalSize3() As String
        Get
            Return Me.prop_TowerRedundantHipDiagonalSize3
        End Get
        Set
            Me.prop_TowerRedundantHipDiagonalSize3 = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipdiagonalsize4")>
    Public Property TowerRedundantHipDiagonalSize4() As String
        Get
            Return Me.prop_TowerRedundantHipDiagonalSize4
        End Get
        Set
            Me.prop_TowerRedundantHipDiagonalSize4 = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerswmult")>
    Public Property TowerSWMult() As Double
        Get
            Return Me.prop_TowerSWMult
        End Get
        Set
            Me.prop_TowerSWMult = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerwpmult")>
    Public Property TowerWPMult() As Double
        Get
            Return Me.prop_TowerWPMult
        End Get
        Set
            Me.prop_TowerWPMult = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerautocalcksingleangle")>
    Public Property TowerAutoCalcKSingleAngle() As Boolean
        Get
            Return Me.prop_TowerAutoCalcKSingleAngle
        End Get
        Set
            Me.prop_TowerAutoCalcKSingleAngle = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerautocalcksolidround")>
    Public Property TowerAutoCalcKSolidRound() As Boolean
        Get
            Return Me.prop_TowerAutoCalcKSolidRound
        End Get
        Set
            Me.prop_TowerAutoCalcKSolidRound = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerafgusset")>
    Public Property TowerAfGusset() As Double
        Get
            Return Me.prop_TowerAfGusset
        End Get
        Set
            Me.prop_TowerAfGusset = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towertfgusset")>
    Public Property TowerTfGusset() As Double
        Get
            Return Me.prop_TowerTfGusset
        End Get
        Set
            Me.prop_TowerTfGusset = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towergussetboltedgedistance")>
    Public Property TowerGussetBoltEdgeDistance() As Double
        Get
            Return Me.prop_TowerGussetBoltEdgeDistance
        End Get
        Set
            Me.prop_TowerGussetBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towergussetgrade")>
    Public Property TowerGussetGrade() As Double
        Get
            Return Me.prop_TowerGussetGrade
        End Get
        Set
            Me.prop_TowerGussetGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towergussetmatlgrade")>
    Public Property TowerGussetMatlGrade() As String
        Get
            Return Me.prop_TowerGussetMatlGrade
        End Get
        Set
            Me.prop_TowerGussetMatlGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerafmult")>
    Public Property TowerAfMult() As Double
        Get
            Return Me.prop_TowerAfMult
        End Get
        Set
            Me.prop_TowerAfMult = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerarmult")>
    Public Property TowerArMult() As Double
        Get
            Return Me.prop_TowerArMult
        End Get
        Set
            Me.prop_TowerArMult = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerflatipapole")>
    Public Property TowerFlatIPAPole() As Double
        Get
            Return Me.prop_TowerFlatIPAPole
        End Get
        Set
            Me.prop_TowerFlatIPAPole = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerroundipapole")>
    Public Property TowerRoundIPAPole() As Double
        Get
            Return Me.prop_TowerRoundIPAPole
        End Get
        Set
            Me.prop_TowerRoundIPAPole = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerflatipaleg")>
    Public Property TowerFlatIPALeg() As Double
        Get
            Return Me.prop_TowerFlatIPALeg
        End Get
        Set
            Me.prop_TowerFlatIPALeg = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerroundipaleg")>
    Public Property TowerRoundIPALeg() As Double
        Get
            Return Me.prop_TowerRoundIPALeg
        End Get
        Set
            Me.prop_TowerRoundIPALeg = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerflatipahorizontal")>
    Public Property TowerFlatIPAHorizontal() As Double
        Get
            Return Me.prop_TowerFlatIPAHorizontal
        End Get
        Set
            Me.prop_TowerFlatIPAHorizontal = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerroundipahorizontal")>
    Public Property TowerRoundIPAHorizontal() As Double
        Get
            Return Me.prop_TowerRoundIPAHorizontal
        End Get
        Set
            Me.prop_TowerRoundIPAHorizontal = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerflatipadiagonal")>
    Public Property TowerFlatIPADiagonal() As Double
        Get
            Return Me.prop_TowerFlatIPADiagonal
        End Get
        Set
            Me.prop_TowerFlatIPADiagonal = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerroundipadiagonal")>
    Public Property TowerRoundIPADiagonal() As Double
        Get
            Return Me.prop_TowerRoundIPADiagonal
        End Get
        Set
            Me.prop_TowerRoundIPADiagonal = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towercsa_S37_Speedupfactor")>
    Public Property TowerCSA_S37_SpeedUpFactor() As Double
        Get
            Return Me.prop_TowerCSA_S37_SpeedUpFactor
        End Get
        Set
            Me.prop_TowerCSA_S37_SpeedUpFactor = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerklegs")>
    Public Property TowerKLegs() As Double
        Get
            Return Me.prop_TowerKLegs
        End Get
        Set
            Me.prop_TowerKLegs = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkxbraceddiags")>
    Public Property TowerKXBracedDiags() As Double
        Get
            Return Me.prop_TowerKXBracedDiags
        End Get
        Set
            Me.prop_TowerKXBracedDiags = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkkbraceddiags")>
    Public Property TowerKKBracedDiags() As Double
        Get
            Return Me.prop_TowerKKBracedDiags
        End Get
        Set
            Me.prop_TowerKKBracedDiags = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkzbraceddiags")>
    Public Property TowerKZBracedDiags() As Double
        Get
            Return Me.prop_TowerKZBracedDiags
        End Get
        Set
            Me.prop_TowerKZBracedDiags = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkhorzs")>
    Public Property TowerKHorzs() As Double
        Get
            Return Me.prop_TowerKHorzs
        End Get
        Set
            Me.prop_TowerKHorzs = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerksechorzs")>
    Public Property TowerKSecHorzs() As Double
        Get
            Return Me.prop_TowerKSecHorzs
        End Get
        Set
            Me.prop_TowerKSecHorzs = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkgirts")>
    Public Property TowerKGirts() As Double
        Get
            Return Me.prop_TowerKGirts
        End Get
        Set
            Me.prop_TowerKGirts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkinners")>
    Public Property TowerKInners() As Double
        Get
            Return Me.prop_TowerKInners
        End Get
        Set
            Me.prop_TowerKInners = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkxbraceddiagsy")>
    Public Property TowerKXBracedDiagsY() As Double
        Get
            Return Me.prop_TowerKXBracedDiagsY
        End Get
        Set
            Me.prop_TowerKXBracedDiagsY = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkkbraceddiagsy")>
    Public Property TowerKKBracedDiagsY() As Double
        Get
            Return Me.prop_TowerKKBracedDiagsY
        End Get
        Set
            Me.prop_TowerKKBracedDiagsY = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkzbraceddiagsy")>
    Public Property TowerKZBracedDiagsY() As Double
        Get
            Return Me.prop_TowerKZBracedDiagsY
        End Get
        Set
            Me.prop_TowerKZBracedDiagsY = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkhorzsy")>
    Public Property TowerKHorzsY() As Double
        Get
            Return Me.prop_TowerKHorzsY
        End Get
        Set
            Me.prop_TowerKHorzsY = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerksechorzsy")>
    Public Property TowerKSecHorzsY() As Double
        Get
            Return Me.prop_TowerKSecHorzsY
        End Get
        Set
            Me.prop_TowerKSecHorzsY = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkgirtsy")>
    Public Property TowerKGirtsY() As Double
        Get
            Return Me.prop_TowerKGirtsY
        End Get
        Set
            Me.prop_TowerKGirtsY = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkinnersy")>
    Public Property TowerKInnersY() As Double
        Get
            Return Me.prop_TowerKInnersY
        End Get
        Set
            Me.prop_TowerKInnersY = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkredhorz")>
    Public Property TowerKRedHorz() As Double
        Get
            Return Me.prop_TowerKRedHorz
        End Get
        Set
            Me.prop_TowerKRedHorz = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkreddiag")>
    Public Property TowerKRedDiag() As Double
        Get
            Return Me.prop_TowerKRedDiag
        End Get
        Set
            Me.prop_TowerKRedDiag = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkredsubdiag")>
    Public Property TowerKRedSubDiag() As Double
        Get
            Return Me.prop_TowerKRedSubDiag
        End Get
        Set
            Me.prop_TowerKRedSubDiag = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkredsubhorz")>
    Public Property TowerKRedSubHorz() As Double
        Get
            Return Me.prop_TowerKRedSubHorz
        End Get
        Set
            Me.prop_TowerKRedSubHorz = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkredvert")>
    Public Property TowerKRedVert() As Double
        Get
            Return Me.prop_TowerKRedVert
        End Get
        Set
            Me.prop_TowerKRedVert = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkredhip")>
    Public Property TowerKRedHip() As Double
        Get
            Return Me.prop_TowerKRedHip
        End Get
        Set
            Me.prop_TowerKRedHip = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkredhipdiag")>
    Public Property TowerKRedHipDiag() As Double
        Get
            Return Me.prop_TowerKRedHipDiag
        End Get
        Set
            Me.prop_TowerKRedHipDiag = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerktlx")>
    Public Property TowerKTLX() As Double
        Get
            Return Me.prop_TowerKTLX
        End Get
        Set
            Me.prop_TowerKTLX = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerktlz")>
    Public Property TowerKTLZ() As Double
        Get
            Return Me.prop_TowerKTLZ
        End Get
        Set
            Me.prop_TowerKTLZ = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerktlleg")>
    Public Property TowerKTLLeg() As Double
        Get
            Return Me.prop_TowerKTLLeg
        End Get
        Set
            Me.prop_TowerKTLLeg = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnerktlx")>
    Public Property TowerInnerKTLX() As Double
        Get
            Return Me.prop_TowerInnerKTLX
        End Get
        Set
            Me.prop_TowerInnerKTLX = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnerktlz")>
    Public Property TowerInnerKTLZ() As Double
        Get
            Return Me.prop_TowerInnerKTLZ
        End Get
        Set
            Me.prop_TowerInnerKTLZ = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnerktlleg")>
    Public Property TowerInnerKTLLeg() As Double
        Get
            Return Me.prop_TowerInnerKTLLeg
        End Get
        Set
            Me.prop_TowerInnerKTLLeg = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerstitchboltlocationhoriz")>
    Public Property TowerStitchBoltLocationHoriz() As String
        Get
            Return Me.prop_TowerStitchBoltLocationHoriz
        End Get
        Set
            Me.prop_TowerStitchBoltLocationHoriz = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerstitchboltlocationdiag")>
    Public Property TowerStitchBoltLocationDiag() As String
        Get
            Return Me.prop_TowerStitchBoltLocationDiag
        End Get
        Set
            Me.prop_TowerStitchBoltLocationDiag = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerstitchboltlocationred")>
    Public Property TowerStitchBoltLocationRed() As String
        Get
            Return Me.prop_TowerStitchBoltLocationRed
        End Get
        Set
            Me.prop_TowerStitchBoltLocationRed = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerstitchspacing")>
    Public Property TowerStitchSpacing() As Double
        Get
            Return Me.prop_TowerStitchSpacing
        End Get
        Set
            Me.prop_TowerStitchSpacing = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerstitchspacingdiag")>
    Public Property TowerStitchSpacingDiag() As Double
        Get
            Return Me.prop_TowerStitchSpacingDiag
        End Get
        Set
            Me.prop_TowerStitchSpacingDiag = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerstitchspacinghorz")>
    Public Property TowerStitchSpacingHorz() As Double
        Get
            Return Me.prop_TowerStitchSpacingHorz
        End Get
        Set
            Me.prop_TowerStitchSpacingHorz = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerstitchspacingred")>
    Public Property TowerStitchSpacingRed() As Double
        Get
            Return Me.prop_TowerStitchSpacingRed
        End Get
        Set
            Me.prop_TowerStitchSpacingRed = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerlegnetwidthdeduct")>
    Public Property TowerLegNetWidthDeduct() As Double
        Get
            Return Me.prop_TowerLegNetWidthDeduct
        End Get
        Set
            Me.prop_TowerLegNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerlegufactor")>
    Public Property TowerLegUFactor() As Double
        Get
            Return Me.prop_TowerLegUFactor
        End Get
        Set
            Me.prop_TowerLegUFactor = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagonalnetwidthdeduct")>
    Public Property TowerDiagonalNetWidthDeduct() As Double
        Get
            Return Me.prop_TowerDiagonalNetWidthDeduct
        End Get
        Set
            Me.prop_TowerDiagonalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towertopgirtnetwidthdeduct")>
    Public Property TowerTopGirtNetWidthDeduct() As Double
        Get
            Return Me.prop_TowerTopGirtNetWidthDeduct
        End Get
        Set
            Me.prop_TowerTopGirtNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerbotgirtnetwidthdeduct")>
    Public Property TowerBotGirtNetWidthDeduct() As Double
        Get
            Return Me.prop_TowerBotGirtNetWidthDeduct
        End Get
        Set
            Me.prop_TowerBotGirtNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnergirtnetwidthdeduct")>
    Public Property TowerInnerGirtNetWidthDeduct() As Double
        Get
            Return Me.prop_TowerInnerGirtNetWidthDeduct
        End Get
        Set
            Me.prop_TowerInnerGirtNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerhorizontalnetwidthdeduct")>
    Public Property TowerHorizontalNetWidthDeduct() As Double
        Get
            Return Me.prop_TowerHorizontalNetWidthDeduct
        End Get
        Set
            Me.prop_TowerHorizontalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towershorthorizontalnetwidthdeduct")>
    Public Property TowerShortHorizontalNetWidthDeduct() As Double
        Get
            Return Me.prop_TowerShortHorizontalNetWidthDeduct
        End Get
        Set
            Me.prop_TowerShortHorizontalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagonalufactor")>
    Public Property TowerDiagonalUFactor() As Double
        Get
            Return Me.prop_TowerDiagonalUFactor
        End Get
        Set
            Me.prop_TowerDiagonalUFactor = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towertopgirtufactor")>
    Public Property TowerTopGirtUFactor() As Double
        Get
            Return Me.prop_TowerTopGirtUFactor
        End Get
        Set
            Me.prop_TowerTopGirtUFactor = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerbotgirtufactor")>
    Public Property TowerBotGirtUFactor() As Double
        Get
            Return Me.prop_TowerBotGirtUFactor
        End Get
        Set
            Me.prop_TowerBotGirtUFactor = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnergirtufactor")>
    Public Property TowerInnerGirtUFactor() As Double
        Get
            Return Me.prop_TowerInnerGirtUFactor
        End Get
        Set
            Me.prop_TowerInnerGirtUFactor = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerhorizontalufactor")>
    Public Property TowerHorizontalUFactor() As Double
        Get
            Return Me.prop_TowerHorizontalUFactor
        End Get
        Set
            Me.prop_TowerHorizontalUFactor = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towershorthorizontalufactor")>
    Public Property TowerShortHorizontalUFactor() As Double
        Get
            Return Me.prop_TowerShortHorizontalUFactor
        End Get
        Set
            Me.prop_TowerShortHorizontalUFactor = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerlegconntype")>
    Public Property TowerLegConnType() As String
        Get
            Return Me.prop_TowerLegConnType
        End Get
        Set
            Me.prop_TowerLegConnType = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerlegnumbolts")>
    Public Property TowerLegNumBolts() As Integer
        Get
            Return Me.prop_TowerLegNumBolts
        End Get
        Set
            Me.prop_TowerLegNumBolts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagonalnumbolts")>
    Public Property TowerDiagonalNumBolts() As Integer
        Get
            Return Me.prop_TowerDiagonalNumBolts
        End Get
        Set
            Me.prop_TowerDiagonalNumBolts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towertopgirtnumbolts")>
    Public Property TowerTopGirtNumBolts() As Integer
        Get
            Return Me.prop_TowerTopGirtNumBolts
        End Get
        Set
            Me.prop_TowerTopGirtNumBolts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerbotgirtnumbolts")>
    Public Property TowerBotGirtNumBolts() As Integer
        Get
            Return Me.prop_TowerBotGirtNumBolts
        End Get
        Set
            Me.prop_TowerBotGirtNumBolts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnergirtnumbolts")>
    Public Property TowerInnerGirtNumBolts() As Integer
        Get
            Return Me.prop_TowerInnerGirtNumBolts
        End Get
        Set
            Me.prop_TowerInnerGirtNumBolts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerhorizontalnumbolts")>
    Public Property TowerHorizontalNumBolts() As Integer
        Get
            Return Me.prop_TowerHorizontalNumBolts
        End Get
        Set
            Me.prop_TowerHorizontalNumBolts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towershorthorizontalnumbolts")>
    Public Property TowerShortHorizontalNumBolts() As Integer
        Get
            Return Me.prop_TowerShortHorizontalNumBolts
        End Get
        Set
            Me.prop_TowerShortHorizontalNumBolts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerlegboltgrade")>
    Public Property TowerLegBoltGrade() As String
        Get
            Return Me.prop_TowerLegBoltGrade
        End Get
        Set
            Me.prop_TowerLegBoltGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerlegboltsize")>
    Public Property TowerLegBoltSize() As Double
        Get
            Return Me.prop_TowerLegBoltSize
        End Get
        Set
            Me.prop_TowerLegBoltSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagonalboltgrade")>
    Public Property TowerDiagonalBoltGrade() As String
        Get
            Return Me.prop_TowerDiagonalBoltGrade
        End Get
        Set
            Me.prop_TowerDiagonalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagonalboltsize")>
    Public Property TowerDiagonalBoltSize() As Double
        Get
            Return Me.prop_TowerDiagonalBoltSize
        End Get
        Set
            Me.prop_TowerDiagonalBoltSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towertopgirtboltgrade")>
    Public Property TowerTopGirtBoltGrade() As String
        Get
            Return Me.prop_TowerTopGirtBoltGrade
        End Get
        Set
            Me.prop_TowerTopGirtBoltGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towertopgirtboltsize")>
    Public Property TowerTopGirtBoltSize() As Double
        Get
            Return Me.prop_TowerTopGirtBoltSize
        End Get
        Set
            Me.prop_TowerTopGirtBoltSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerbotgirtboltgrade")>
    Public Property TowerBotGirtBoltGrade() As String
        Get
            Return Me.prop_TowerBotGirtBoltGrade
        End Get
        Set
            Me.prop_TowerBotGirtBoltGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerbotgirtboltsize")>
    Public Property TowerBotGirtBoltSize() As Double
        Get
            Return Me.prop_TowerBotGirtBoltSize
        End Get
        Set
            Me.prop_TowerBotGirtBoltSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnergirtboltgrade")>
    Public Property TowerInnerGirtBoltGrade() As String
        Get
            Return Me.prop_TowerInnerGirtBoltGrade
        End Get
        Set
            Me.prop_TowerInnerGirtBoltGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnergirtboltsize")>
    Public Property TowerInnerGirtBoltSize() As Double
        Get
            Return Me.prop_TowerInnerGirtBoltSize
        End Get
        Set
            Me.prop_TowerInnerGirtBoltSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerhorizontalboltgrade")>
    Public Property TowerHorizontalBoltGrade() As String
        Get
            Return Me.prop_TowerHorizontalBoltGrade
        End Get
        Set
            Me.prop_TowerHorizontalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerhorizontalboltsize")>
    Public Property TowerHorizontalBoltSize() As Double
        Get
            Return Me.prop_TowerHorizontalBoltSize
        End Get
        Set
            Me.prop_TowerHorizontalBoltSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towershorthorizontalboltgrade")>
    Public Property TowerShortHorizontalBoltGrade() As String
        Get
            Return Me.prop_TowerShortHorizontalBoltGrade
        End Get
        Set
            Me.prop_TowerShortHorizontalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towershorthorizontalboltsize")>
    Public Property TowerShortHorizontalBoltSize() As Double
        Get
            Return Me.prop_TowerShortHorizontalBoltSize
        End Get
        Set
            Me.prop_TowerShortHorizontalBoltSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerlegboltedgedistance")>
    Public Property TowerLegBoltEdgeDistance() As Double
        Get
            Return Me.prop_TowerLegBoltEdgeDistance
        End Get
        Set
            Me.prop_TowerLegBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagonalboltedgedistance")>
    Public Property TowerDiagonalBoltEdgeDistance() As Double
        Get
            Return Me.prop_TowerDiagonalBoltEdgeDistance
        End Get
        Set
            Me.prop_TowerDiagonalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towertopgirtboltedgedistance")>
    Public Property TowerTopGirtBoltEdgeDistance() As Double
        Get
            Return Me.prop_TowerTopGirtBoltEdgeDistance
        End Get
        Set
            Me.prop_TowerTopGirtBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerbotgirtboltedgedistance")>
    Public Property TowerBotGirtBoltEdgeDistance() As Double
        Get
            Return Me.prop_TowerBotGirtBoltEdgeDistance
        End Get
        Set
            Me.prop_TowerBotGirtBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnergirtboltedgedistance")>
    Public Property TowerInnerGirtBoltEdgeDistance() As Double
        Get
            Return Me.prop_TowerInnerGirtBoltEdgeDistance
        End Get
        Set
            Me.prop_TowerInnerGirtBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerhorizontalboltedgedistance")>
    Public Property TowerHorizontalBoltEdgeDistance() As Double
        Get
            Return Me.prop_TowerHorizontalBoltEdgeDistance
        End Get
        Set
            Me.prop_TowerHorizontalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towershorthorizontalboltedgedistance")>
    Public Property TowerShortHorizontalBoltEdgeDistance() As Double
        Get
            Return Me.prop_TowerShortHorizontalBoltEdgeDistance
        End Get
        Set
            Me.prop_TowerShortHorizontalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagonalgageg1Distance")>
    Public Property TowerDiagonalGageG1Distance() As Double
        Get
            Return Me.prop_TowerDiagonalGageG1Distance
        End Get
        Set
            Me.prop_TowerDiagonalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towertopgirtgageg1Distance")>
    Public Property TowerTopGirtGageG1Distance() As Double
        Get
            Return Me.prop_TowerTopGirtGageG1Distance
        End Get
        Set
            Me.prop_TowerTopGirtGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerbotgirtgageg1Distance")>
    Public Property TowerBotGirtGageG1Distance() As Double
        Get
            Return Me.prop_TowerBotGirtGageG1Distance
        End Get
        Set
            Me.prop_TowerBotGirtGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerinnergirtgageg1Distance")>
    Public Property TowerInnerGirtGageG1Distance() As Double
        Get
            Return Me.prop_TowerInnerGirtGageG1Distance
        End Get
        Set
            Me.prop_TowerInnerGirtGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerhorizontalgageg1Distance")>
    Public Property TowerHorizontalGageG1Distance() As Double
        Get
            Return Me.prop_TowerHorizontalGageG1Distance
        End Get
        Set
            Me.prop_TowerHorizontalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towershorthorizontalgageg1Distance")>
    Public Property TowerShortHorizontalGageG1Distance() As Double
        Get
            Return Me.prop_TowerShortHorizontalGageG1Distance
        End Get
        Set
            Me.prop_TowerShortHorizontalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthorizontalboltgrade")>
    Public Property TowerRedundantHorizontalBoltGrade() As String
        Get
            Return Me.prop_TowerRedundantHorizontalBoltGrade
        End Get
        Set
            Me.prop_TowerRedundantHorizontalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthorizontalboltsize")>
    Public Property TowerRedundantHorizontalBoltSize() As Double
        Get
            Return Me.prop_TowerRedundantHorizontalBoltSize
        End Get
        Set
            Me.prop_TowerRedundantHorizontalBoltSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthorizontalnumbolts")>
    Public Property TowerRedundantHorizontalNumBolts() As Integer
        Get
            Return Me.prop_TowerRedundantHorizontalNumBolts
        End Get
        Set
            Me.prop_TowerRedundantHorizontalNumBolts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthorizontalboltedgedistance")>
    Public Property TowerRedundantHorizontalBoltEdgeDistance() As Double
        Get
            Return Me.prop_TowerRedundantHorizontalBoltEdgeDistance
        End Get
        Set
            Me.prop_TowerRedundantHorizontalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthorizontalgageg1Distance")>
    Public Property TowerRedundantHorizontalGageG1Distance() As Double
        Get
            Return Me.prop_TowerRedundantHorizontalGageG1Distance
        End Get
        Set
            Me.prop_TowerRedundantHorizontalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthorizontalnetwidthdeduct")>
    Public Property TowerRedundantHorizontalNetWidthDeduct() As Double
        Get
            Return Me.prop_TowerRedundantHorizontalNetWidthDeduct
        End Get
        Set
            Me.prop_TowerRedundantHorizontalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthorizontalufactor")>
    Public Property TowerRedundantHorizontalUFactor() As Double
        Get
            Return Me.prop_TowerRedundantHorizontalUFactor
        End Get
        Set
            Me.prop_TowerRedundantHorizontalUFactor = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantdiagonalboltgrade")>
    Public Property TowerRedundantDiagonalBoltGrade() As String
        Get
            Return Me.prop_TowerRedundantDiagonalBoltGrade
        End Get
        Set
            Me.prop_TowerRedundantDiagonalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantdiagonalboltsize")>
    Public Property TowerRedundantDiagonalBoltSize() As Double
        Get
            Return Me.prop_TowerRedundantDiagonalBoltSize
        End Get
        Set
            Me.prop_TowerRedundantDiagonalBoltSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantdiagonalnumbolts")>
    Public Property TowerRedundantDiagonalNumBolts() As Integer
        Get
            Return Me.prop_TowerRedundantDiagonalNumBolts
        End Get
        Set
            Me.prop_TowerRedundantDiagonalNumBolts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantdiagonalboltedgedistance")>
    Public Property TowerRedundantDiagonalBoltEdgeDistance() As Double
        Get
            Return Me.prop_TowerRedundantDiagonalBoltEdgeDistance
        End Get
        Set
            Me.prop_TowerRedundantDiagonalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantdiagonalgageg1Distance")>
    Public Property TowerRedundantDiagonalGageG1Distance() As Double
        Get
            Return Me.prop_TowerRedundantDiagonalGageG1Distance
        End Get
        Set
            Me.prop_TowerRedundantDiagonalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantdiagonalnetwidthdeduct")>
    Public Property TowerRedundantDiagonalNetWidthDeduct() As Double
        Get
            Return Me.prop_TowerRedundantDiagonalNetWidthDeduct
        End Get
        Set
            Me.prop_TowerRedundantDiagonalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantdiagonalufactor")>
    Public Property TowerRedundantDiagonalUFactor() As Double
        Get
            Return Me.prop_TowerRedundantDiagonalUFactor
        End Get
        Set
            Me.prop_TowerRedundantDiagonalUFactor = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubdiagonalboltgrade")>
    Public Property TowerRedundantSubDiagonalBoltGrade() As String
        Get
            Return Me.prop_TowerRedundantSubDiagonalBoltGrade
        End Get
        Set
            Me.prop_TowerRedundantSubDiagonalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubdiagonalboltsize")>
    Public Property TowerRedundantSubDiagonalBoltSize() As Double
        Get
            Return Me.prop_TowerRedundantSubDiagonalBoltSize
        End Get
        Set
            Me.prop_TowerRedundantSubDiagonalBoltSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubdiagonalnumbolts")>
    Public Property TowerRedundantSubDiagonalNumBolts() As Integer
        Get
            Return Me.prop_TowerRedundantSubDiagonalNumBolts
        End Get
        Set
            Me.prop_TowerRedundantSubDiagonalNumBolts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubdiagonalboltedgedistance")>
    Public Property TowerRedundantSubDiagonalBoltEdgeDistance() As Double
        Get
            Return Me.prop_TowerRedundantSubDiagonalBoltEdgeDistance
        End Get
        Set
            Me.prop_TowerRedundantSubDiagonalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubdiagonalgageg1Distance")>
    Public Property TowerRedundantSubDiagonalGageG1Distance() As Double
        Get
            Return Me.prop_TowerRedundantSubDiagonalGageG1Distance
        End Get
        Set
            Me.prop_TowerRedundantSubDiagonalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubdiagonalnetwidthdeduct")>
    Public Property TowerRedundantSubDiagonalNetWidthDeduct() As Double
        Get
            Return Me.prop_TowerRedundantSubDiagonalNetWidthDeduct
        End Get
        Set
            Me.prop_TowerRedundantSubDiagonalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubdiagonalufactor")>
    Public Property TowerRedundantSubDiagonalUFactor() As Double
        Get
            Return Me.prop_TowerRedundantSubDiagonalUFactor
        End Get
        Set
            Me.prop_TowerRedundantSubDiagonalUFactor = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubhorizontalboltgrade")>
    Public Property TowerRedundantSubHorizontalBoltGrade() As String
        Get
            Return Me.prop_TowerRedundantSubHorizontalBoltGrade
        End Get
        Set
            Me.prop_TowerRedundantSubHorizontalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubhorizontalboltsize")>
    Public Property TowerRedundantSubHorizontalBoltSize() As Double
        Get
            Return Me.prop_TowerRedundantSubHorizontalBoltSize
        End Get
        Set
            Me.prop_TowerRedundantSubHorizontalBoltSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubhorizontalnumbolts")>
    Public Property TowerRedundantSubHorizontalNumBolts() As Integer
        Get
            Return Me.prop_TowerRedundantSubHorizontalNumBolts
        End Get
        Set
            Me.prop_TowerRedundantSubHorizontalNumBolts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubhorizontalboltedgedistance")>
    Public Property TowerRedundantSubHorizontalBoltEdgeDistance() As Double
        Get
            Return Me.prop_TowerRedundantSubHorizontalBoltEdgeDistance
        End Get
        Set
            Me.prop_TowerRedundantSubHorizontalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubhorizontalgageg1Distance")>
    Public Property TowerRedundantSubHorizontalGageG1Distance() As Double
        Get
            Return Me.prop_TowerRedundantSubHorizontalGageG1Distance
        End Get
        Set
            Me.prop_TowerRedundantSubHorizontalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubhorizontalnetwidthdeduct")>
    Public Property TowerRedundantSubHorizontalNetWidthDeduct() As Double
        Get
            Return Me.prop_TowerRedundantSubHorizontalNetWidthDeduct
        End Get
        Set
            Me.prop_TowerRedundantSubHorizontalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantsubhorizontalufactor")>
    Public Property TowerRedundantSubHorizontalUFactor() As Double
        Get
            Return Me.prop_TowerRedundantSubHorizontalUFactor
        End Get
        Set
            Me.prop_TowerRedundantSubHorizontalUFactor = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantverticalboltgrade")>
    Public Property TowerRedundantVerticalBoltGrade() As String
        Get
            Return Me.prop_TowerRedundantVerticalBoltGrade
        End Get
        Set
            Me.prop_TowerRedundantVerticalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantverticalboltsize")>
    Public Property TowerRedundantVerticalBoltSize() As Double
        Get
            Return Me.prop_TowerRedundantVerticalBoltSize
        End Get
        Set
            Me.prop_TowerRedundantVerticalBoltSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantverticalnumbolts")>
    Public Property TowerRedundantVerticalNumBolts() As Integer
        Get
            Return Me.prop_TowerRedundantVerticalNumBolts
        End Get
        Set
            Me.prop_TowerRedundantVerticalNumBolts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantverticalboltedgedistance")>
    Public Property TowerRedundantVerticalBoltEdgeDistance() As Double
        Get
            Return Me.prop_TowerRedundantVerticalBoltEdgeDistance
        End Get
        Set
            Me.prop_TowerRedundantVerticalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantverticalgageg1Distance")>
    Public Property TowerRedundantVerticalGageG1Distance() As Double
        Get
            Return Me.prop_TowerRedundantVerticalGageG1Distance
        End Get
        Set
            Me.prop_TowerRedundantVerticalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantverticalnetwidthdeduct")>
    Public Property TowerRedundantVerticalNetWidthDeduct() As Double
        Get
            Return Me.prop_TowerRedundantVerticalNetWidthDeduct
        End Get
        Set
            Me.prop_TowerRedundantVerticalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundantverticalufactor")>
    Public Property TowerRedundantVerticalUFactor() As Double
        Get
            Return Me.prop_TowerRedundantVerticalUFactor
        End Get
        Set
            Me.prop_TowerRedundantVerticalUFactor = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipboltgrade")>
    Public Property TowerRedundantHipBoltGrade() As String
        Get
            Return Me.prop_TowerRedundantHipBoltGrade
        End Get
        Set
            Me.prop_TowerRedundantHipBoltGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipboltsize")>
    Public Property TowerRedundantHipBoltSize() As Double
        Get
            Return Me.prop_TowerRedundantHipBoltSize
        End Get
        Set
            Me.prop_TowerRedundantHipBoltSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipnumbolts")>
    Public Property TowerRedundantHipNumBolts() As Integer
        Get
            Return Me.prop_TowerRedundantHipNumBolts
        End Get
        Set
            Me.prop_TowerRedundantHipNumBolts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipboltedgedistance")>
    Public Property TowerRedundantHipBoltEdgeDistance() As Double
        Get
            Return Me.prop_TowerRedundantHipBoltEdgeDistance
        End Get
        Set
            Me.prop_TowerRedundantHipBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipgageg1Distance")>
    Public Property TowerRedundantHipGageG1Distance() As Double
        Get
            Return Me.prop_TowerRedundantHipGageG1Distance
        End Get
        Set
            Me.prop_TowerRedundantHipGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipnetwidthdeduct")>
    Public Property TowerRedundantHipNetWidthDeduct() As Double
        Get
            Return Me.prop_TowerRedundantHipNetWidthDeduct
        End Get
        Set
            Me.prop_TowerRedundantHipNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipufactor")>
    Public Property TowerRedundantHipUFactor() As Double
        Get
            Return Me.prop_TowerRedundantHipUFactor
        End Get
        Set
            Me.prop_TowerRedundantHipUFactor = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipdiagonalboltgrade")>
    Public Property TowerRedundantHipDiagonalBoltGrade() As String
        Get
            Return Me.prop_TowerRedundantHipDiagonalBoltGrade
        End Get
        Set
            Me.prop_TowerRedundantHipDiagonalBoltGrade = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipdiagonalboltsize")>
    Public Property TowerRedundantHipDiagonalBoltSize() As Double
        Get
            Return Me.prop_TowerRedundantHipDiagonalBoltSize
        End Get
        Set
            Me.prop_TowerRedundantHipDiagonalBoltSize = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipdiagonalnumbolts")>
    Public Property TowerRedundantHipDiagonalNumBolts() As Integer
        Get
            Return Me.prop_TowerRedundantHipDiagonalNumBolts
        End Get
        Set
            Me.prop_TowerRedundantHipDiagonalNumBolts = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipdiagonalboltedgedistance")>
    Public Property TowerRedundantHipDiagonalBoltEdgeDistance() As Double
        Get
            Return Me.prop_TowerRedundantHipDiagonalBoltEdgeDistance
        End Get
        Set
            Me.prop_TowerRedundantHipDiagonalBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipdiagonalgageg1Distance")>
    Public Property TowerRedundantHipDiagonalGageG1Distance() As Double
        Get
            Return Me.prop_TowerRedundantHipDiagonalGageG1Distance
        End Get
        Set
            Me.prop_TowerRedundantHipDiagonalGageG1Distance = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipdiagonalnetwidthdeduct")>
    Public Property TowerRedundantHipDiagonalNetWidthDeduct() As Double
        Get
            Return Me.prop_TowerRedundantHipDiagonalNetWidthDeduct
        End Get
        Set
            Me.prop_TowerRedundantHipDiagonalNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerredundanthipdiagonalufactor")>
    Public Property TowerRedundantHipDiagonalUFactor() As Double
        Get
            Return Me.prop_TowerRedundantHipDiagonalUFactor
        End Get
        Set
            Me.prop_TowerRedundantHipDiagonalUFactor = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagonaloutofplanerestraint")>
    Public Property TowerDiagonalOutOfPlaneRestraint() As Boolean
        Get
            Return Me.prop_TowerDiagonalOutOfPlaneRestraint
        End Get
        Set
            Me.prop_TowerDiagonalOutOfPlaneRestraint = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towertopgirtoutofplanerestraint")>
    Public Property TowerTopGirtOutOfPlaneRestraint() As Boolean
        Get
            Return Me.prop_TowerTopGirtOutOfPlaneRestraint
        End Get
        Set
            Me.prop_TowerTopGirtOutOfPlaneRestraint = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerbottomgirtoutofplanerestraint")>
    Public Property TowerBottomGirtOutOfPlaneRestraint() As Boolean
        Get
            Return Me.prop_TowerBottomGirtOutOfPlaneRestraint
        End Get
        Set
            Me.prop_TowerBottomGirtOutOfPlaneRestraint = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towermidgirtoutofplanerestraint")>
    Public Property TowerMidGirtOutOfPlaneRestraint() As Boolean
        Get
            Return Me.prop_TowerMidGirtOutOfPlaneRestraint
        End Get
        Set
            Me.prop_TowerMidGirtOutOfPlaneRestraint = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerhorizontaloutofplanerestraint")>
    Public Property TowerHorizontalOutOfPlaneRestraint() As Boolean
        Get
            Return Me.prop_TowerHorizontalOutOfPlaneRestraint
        End Get
        Set
            Me.prop_TowerHorizontalOutOfPlaneRestraint = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towersecondaryhorizontaloutofplanerestraint")>
    Public Property TowerSecondaryHorizontalOutOfPlaneRestraint() As Boolean
        Get
            Return Me.prop_TowerSecondaryHorizontalOutOfPlaneRestraint
        End Get
        Set
            Me.prop_TowerSecondaryHorizontalOutOfPlaneRestraint = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Toweruniqueflag")>
    Public Property TowerUniqueFlag() As Integer
        Get
            Return Me.prop_TowerUniqueFlag
        End Get
        Set
            Me.prop_TowerUniqueFlag = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagoffsetney")>
    Public Property TowerDiagOffsetNEY() As Double
        Get
            Return Me.prop_TowerDiagOffsetNEY
        End Get
        Set
            Me.prop_TowerDiagOffsetNEY = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagoffsetnex")>
    Public Property TowerDiagOffsetNEX() As Double
        Get
            Return Me.prop_TowerDiagOffsetNEX
        End Get
        Set
            Me.prop_TowerDiagOffsetNEX = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagoffsetpey")>
    Public Property TowerDiagOffsetPEY() As Double
        Get
            Return Me.prop_TowerDiagOffsetPEY
        End Get
        Set
            Me.prop_TowerDiagOffsetPEY = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerdiagoffsetpex")>
    Public Property TowerDiagOffsetPEX() As Double
        Get
            Return Me.prop_TowerDiagOffsetPEX
        End Get
        Set
            Me.prop_TowerDiagOffsetPEX = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkbraceoffsetney")>
    Public Property TowerKbraceOffsetNEY() As Double
        Get
            Return Me.prop_TowerKbraceOffsetNEY
        End Get
        Set
            Me.prop_TowerKbraceOffsetNEY = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkbraceoffsetnex")>
    Public Property TowerKbraceOffsetNEX() As Double
        Get
            Return Me.prop_TowerKbraceOffsetNEX
        End Get
        Set
            Me.prop_TowerKbraceOffsetNEX = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkbraceoffsetpey")>
    Public Property TowerKbraceOffsetPEY() As Double
        Get
            Return Me.prop_TowerKbraceOffsetPEY
        End Get
        Set
            Me.prop_TowerKbraceOffsetPEY = Value
        End Set
    End Property
    <Category("TNX Tower Record"), Description(""), DisplayName("Towerkbraceoffsetpex")>
    Public Property TowerKbraceOffsetPEX() As Double
        Get
            Return Me.prop_TowerKbraceOffsetPEX
        End Get
        Set
            Me.prop_TowerKbraceOffsetPEX = Value
        End Set
    End Property

End Class

Partial Public Class tnxGuyRecord
    Private prop_ID As Integer
    Private prop_tnxID As Integer
    Private prop_GuyRec As Integer
    Private prop_GuyHeight As Double
    Private prop_GuyAutoCalcKSingleAngle As Boolean
    Private prop_GuyAutoCalcKSolidRound As Boolean
    Private prop_GuyMount As String
    Private prop_TorqueArmStyle As String
    Private prop_GuyRadius As Double
    Private prop_GuyRadius120 As Double
    Private prop_GuyRadius240 As Double
    Private prop_GuyRadius360 As Double
    Private prop_TorqueArmRadius As Double
    Private prop_TorqueArmLegAngle As Double
    Private prop_Azimuth0Adjustment As Double
    Private prop_Azimuth120Adjustment As Double
    Private prop_Azimuth240Adjustment As Double
    Private prop_Azimuth360Adjustment As Double
    Private prop_Anchor0Elevation As Double
    Private prop_Anchor120Elevation As Double
    Private prop_Anchor240Elevation As Double
    Private prop_Anchor360Elevation As Double
    Private prop_GuySize As String
    Private prop_Guy120Size As String
    Private prop_Guy240Size As String
    Private prop_Guy360Size As String
    Private prop_GuyGrade As String
    Private prop_TorqueArmSize As String
    Private prop_TorqueArmSizeBot As String
    Private prop_TorqueArmType As String
    Private prop_TorqueArmGrade As Double
    Private prop_TorqueArmMatlGrade As String
    Private prop_TorqueArmKFactor As Double
    Private prop_TorqueArmKFactorY As Double
    Private prop_GuyPullOffKFactorX As Double
    Private prop_GuyPullOffKFactorY As Double
    Private prop_GuyDiagKFactorX As Double
    Private prop_GuyDiagKFactorY As Double
    Private prop_GuyAutoCalc As Boolean
    Private prop_GuyAllGuysSame As Boolean
    Private prop_GuyAllGuysAnchorSame As Boolean
    Private prop_GuyIsStrapping As Boolean
    Private prop_GuyPullOffSize As String
    Private prop_GuyPullOffSizeBot As String
    Private prop_GuyPullOffType As String
    Private prop_GuyPullOffGrade As Double
    Private prop_GuyPullOffMatlGrade As String
    Private prop_GuyUpperDiagSize As String
    Private prop_GuyLowerDiagSize As String
    Private prop_GuyDiagType As String
    Private prop_GuyDiagGrade As Double
    Private prop_GuyDiagMatlGrade As String
    Private prop_GuyDiagNetWidthDeduct As Double
    Private prop_GuyDiagUFactor As Double
    Private prop_GuyDiagNumBolts As Integer
    Private prop_GuyDiagonalOutOfPlaneRestraint As Boolean
    Private prop_GuyDiagBoltGrade As String
    Private prop_GuyDiagBoltSize As Double
    Private prop_GuyDiagBoltEdgeDistance As Double
    Private prop_GuyDiagBoltGageDistance As Double
    Private prop_GuyPullOffNetWidthDeduct As Double
    Private prop_GuyPullOffUFactor As Double
    Private prop_GuyPullOffNumBolts As Integer
    Private prop_GuyPullOffOutOfPlaneRestraint As Boolean
    Private prop_GuyPullOffBoltGrade As String
    Private prop_GuyPullOffBoltSize As Double
    Private prop_GuyPullOffBoltEdgeDistance As Double
    Private prop_GuyPullOffBoltGageDistance As Double
    Private prop_GuyTorqueArmNetWidthDeduct As Double
    Private prop_GuyTorqueArmUFactor As Double
    Private prop_GuyTorqueArmNumBolts As Integer
    Private prop_GuyTorqueArmOutOfPlaneRestraint As Boolean
    Private prop_GuyTorqueArmBoltGrade As String
    Private prop_GuyTorqueArmBoltSize As Double
    Private prop_GuyTorqueArmBoltEdgeDistance As Double
    Private prop_GuyTorqueArmBoltGageDistance As Double
    Private prop_GuyPerCentTension As Double
    Private prop_GuyPerCentTension120 As Double
    Private prop_GuyPerCentTension240 As Double
    Private prop_GuyPerCentTension360 As Double
    Private prop_GuyEffFactor As Double
    Private prop_GuyEffFactor120 As Double
    Private prop_GuyEffFactor240 As Double
    Private prop_GuyEffFactor360 As Double
    Private prop_GuyNumInsulators As Integer
    Private prop_GuyInsulatorLength As Double
    Private prop_GuyInsulatorDia As Double
    Private prop_GuyInsulatorWt As Double

    <Category("TNX Guy Record"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer
        Get
            Return Me.prop_ID
        End Get
        Set
            Me.prop_ID = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Tnxid")>
    Public Property tnxID() As Integer
        Get
            Return Me.prop_tnxID
        End Get
        Set
            Me.prop_tnxID = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyrec")>
    Public Property GuyRec() As Integer
        Get
            Return Me.prop_GuyRec
        End Get
        Set
            Me.prop_GuyRec = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyheight")>
    Public Property GuyHeight() As Double
        Get
            Return Me.prop_GuyHeight
        End Get
        Set
            Me.prop_GuyHeight = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyautocalcksingleangle")>
    Public Property GuyAutoCalcKSingleAngle() As Boolean
        Get
            Return Me.prop_GuyAutoCalcKSingleAngle
        End Get
        Set
            Me.prop_GuyAutoCalcKSingleAngle = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyautocalcksolidround")>
    Public Property GuyAutoCalcKSolidRound() As Boolean
        Get
            Return Me.prop_GuyAutoCalcKSolidRound
        End Get
        Set
            Me.prop_GuyAutoCalcKSolidRound = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guymount")>
    Public Property GuyMount() As String
        Get
            Return Me.prop_GuyMount
        End Get
        Set
            Me.prop_GuyMount = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Torquearmstyle")>
    Public Property TorqueArmStyle() As String
        Get
            Return Me.prop_TorqueArmStyle
        End Get
        Set
            Me.prop_TorqueArmStyle = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyradius")>
    Public Property GuyRadius() As Double
        Get
            Return Me.prop_GuyRadius
        End Get
        Set
            Me.prop_GuyRadius = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyradius120")>
    Public Property GuyRadius120() As Double
        Get
            Return Me.prop_GuyRadius120
        End Get
        Set
            Me.prop_GuyRadius120 = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyradius240")>
    Public Property GuyRadius240() As Double
        Get
            Return Me.prop_GuyRadius240
        End Get
        Set
            Me.prop_GuyRadius240 = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyradius360")>
    Public Property GuyRadius360() As Double
        Get
            Return Me.prop_GuyRadius360
        End Get
        Set
            Me.prop_GuyRadius360 = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Torquearmradius")>
    Public Property TorqueArmRadius() As Double
        Get
            Return Me.prop_TorqueArmRadius
        End Get
        Set
            Me.prop_TorqueArmRadius = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Torquearmlegangle")>
    Public Property TorqueArmLegAngle() As Double
        Get
            Return Me.prop_TorqueArmLegAngle
        End Get
        Set
            Me.prop_TorqueArmLegAngle = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Azimuth0Adjustment")>
    Public Property Azimuth0Adjustment() As Double
        Get
            Return Me.prop_Azimuth0Adjustment
        End Get
        Set
            Me.prop_Azimuth0Adjustment = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Azimuth120Adjustment")>
    Public Property Azimuth120Adjustment() As Double
        Get
            Return Me.prop_Azimuth120Adjustment
        End Get
        Set
            Me.prop_Azimuth120Adjustment = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Azimuth240Adjustment")>
    Public Property Azimuth240Adjustment() As Double
        Get
            Return Me.prop_Azimuth240Adjustment
        End Get
        Set
            Me.prop_Azimuth240Adjustment = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Azimuth360Adjustment")>
    Public Property Azimuth360Adjustment() As Double
        Get
            Return Me.prop_Azimuth360Adjustment
        End Get
        Set
            Me.prop_Azimuth360Adjustment = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Anchor0Elevation")>
    Public Property Anchor0Elevation() As Double
        Get
            Return Me.prop_Anchor0Elevation
        End Get
        Set
            Me.prop_Anchor0Elevation = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Anchor120Elevation")>
    Public Property Anchor120Elevation() As Double
        Get
            Return Me.prop_Anchor120Elevation
        End Get
        Set
            Me.prop_Anchor120Elevation = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Anchor240Elevation")>
    Public Property Anchor240Elevation() As Double
        Get
            Return Me.prop_Anchor240Elevation
        End Get
        Set
            Me.prop_Anchor240Elevation = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Anchor360Elevation")>
    Public Property Anchor360Elevation() As Double
        Get
            Return Me.prop_Anchor360Elevation
        End Get
        Set
            Me.prop_Anchor360Elevation = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guysize")>
    Public Property GuySize() As String
        Get
            Return Me.prop_GuySize
        End Get
        Set
            Me.prop_GuySize = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guy120Size")>
    Public Property Guy120Size() As String
        Get
            Return Me.prop_Guy120Size
        End Get
        Set
            Me.prop_Guy120Size = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guy240Size")>
    Public Property Guy240Size() As String
        Get
            Return Me.prop_Guy240Size
        End Get
        Set
            Me.prop_Guy240Size = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guy360Size")>
    Public Property Guy360Size() As String
        Get
            Return Me.prop_Guy360Size
        End Get
        Set
            Me.prop_Guy360Size = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guygrade")>
    Public Property GuyGrade() As String
        Get
            Return Me.prop_GuyGrade
        End Get
        Set
            Me.prop_GuyGrade = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Torquearmsize")>
    Public Property TorqueArmSize() As String
        Get
            Return Me.prop_TorqueArmSize
        End Get
        Set
            Me.prop_TorqueArmSize = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Torquearmsizebot")>
    Public Property TorqueArmSizeBot() As String
        Get
            Return Me.prop_TorqueArmSizeBot
        End Get
        Set
            Me.prop_TorqueArmSizeBot = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Torquearmtype")>
    Public Property TorqueArmType() As String
        Get
            Return Me.prop_TorqueArmType
        End Get
        Set
            Me.prop_TorqueArmType = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Torquearmgrade")>
    Public Property TorqueArmGrade() As Double
        Get
            Return Me.prop_TorqueArmGrade
        End Get
        Set
            Me.prop_TorqueArmGrade = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Torquearmmatlgrade")>
    Public Property TorqueArmMatlGrade() As String
        Get
            Return Me.prop_TorqueArmMatlGrade
        End Get
        Set
            Me.prop_TorqueArmMatlGrade = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Torquearmkfactor")>
    Public Property TorqueArmKFactor() As Double
        Get
            Return Me.prop_TorqueArmKFactor
        End Get
        Set
            Me.prop_TorqueArmKFactor = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Torquearmkfactory")>
    Public Property TorqueArmKFactorY() As Double
        Get
            Return Me.prop_TorqueArmKFactorY
        End Get
        Set
            Me.prop_TorqueArmKFactorY = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypulloffkfactorx")>
    Public Property GuyPullOffKFactorX() As Double
        Get
            Return Me.prop_GuyPullOffKFactorX
        End Get
        Set
            Me.prop_GuyPullOffKFactorX = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypulloffkfactory")>
    Public Property GuyPullOffKFactorY() As Double
        Get
            Return Me.prop_GuyPullOffKFactorY
        End Get
        Set
            Me.prop_GuyPullOffKFactorY = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guydiagkfactorx")>
    Public Property GuyDiagKFactorX() As Double
        Get
            Return Me.prop_GuyDiagKFactorX
        End Get
        Set
            Me.prop_GuyDiagKFactorX = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guydiagkfactory")>
    Public Property GuyDiagKFactorY() As Double
        Get
            Return Me.prop_GuyDiagKFactorY
        End Get
        Set
            Me.prop_GuyDiagKFactorY = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyautocalc")>
    Public Property GuyAutoCalc() As Boolean
        Get
            Return Me.prop_GuyAutoCalc
        End Get
        Set
            Me.prop_GuyAutoCalc = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyallguyssame")>
    Public Property GuyAllGuysSame() As Boolean
        Get
            Return Me.prop_GuyAllGuysSame
        End Get
        Set
            Me.prop_GuyAllGuysSame = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyallguysanchorsame")>
    Public Property GuyAllGuysAnchorSame() As Boolean
        Get
            Return Me.prop_GuyAllGuysAnchorSame
        End Get
        Set
            Me.prop_GuyAllGuysAnchorSame = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyisstrapping")>
    Public Property GuyIsStrapping() As Boolean
        Get
            Return Me.prop_GuyIsStrapping
        End Get
        Set
            Me.prop_GuyIsStrapping = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypulloffsize")>
    Public Property GuyPullOffSize() As String
        Get
            Return Me.prop_GuyPullOffSize
        End Get
        Set
            Me.prop_GuyPullOffSize = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypulloffsizebot")>
    Public Property GuyPullOffSizeBot() As String
        Get
            Return Me.prop_GuyPullOffSizeBot
        End Get
        Set
            Me.prop_GuyPullOffSizeBot = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypullofftype")>
    Public Property GuyPullOffType() As String
        Get
            Return Me.prop_GuyPullOffType
        End Get
        Set
            Me.prop_GuyPullOffType = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypulloffgrade")>
    Public Property GuyPullOffGrade() As Double
        Get
            Return Me.prop_GuyPullOffGrade
        End Get
        Set
            Me.prop_GuyPullOffGrade = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypulloffmatlgrade")>
    Public Property GuyPullOffMatlGrade() As String
        Get
            Return Me.prop_GuyPullOffMatlGrade
        End Get
        Set
            Me.prop_GuyPullOffMatlGrade = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyupperdiagsize")>
    Public Property GuyUpperDiagSize() As String
        Get
            Return Me.prop_GuyUpperDiagSize
        End Get
        Set
            Me.prop_GuyUpperDiagSize = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guylowerdiagsize")>
    Public Property GuyLowerDiagSize() As String
        Get
            Return Me.prop_GuyLowerDiagSize
        End Get
        Set
            Me.prop_GuyLowerDiagSize = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guydiagtype")>
    Public Property GuyDiagType() As String
        Get
            Return Me.prop_GuyDiagType
        End Get
        Set
            Me.prop_GuyDiagType = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guydiaggrade")>
    Public Property GuyDiagGrade() As Double
        Get
            Return Me.prop_GuyDiagGrade
        End Get
        Set
            Me.prop_GuyDiagGrade = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guydiagmatlgrade")>
    Public Property GuyDiagMatlGrade() As String
        Get
            Return Me.prop_GuyDiagMatlGrade
        End Get
        Set
            Me.prop_GuyDiagMatlGrade = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guydiagnetwidthdeduct")>
    Public Property GuyDiagNetWidthDeduct() As Double
        Get
            Return Me.prop_GuyDiagNetWidthDeduct
        End Get
        Set
            Me.prop_GuyDiagNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guydiagufactor")>
    Public Property GuyDiagUFactor() As Double
        Get
            Return Me.prop_GuyDiagUFactor
        End Get
        Set
            Me.prop_GuyDiagUFactor = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guydiagnumbolts")>
    Public Property GuyDiagNumBolts() As Integer
        Get
            Return Me.prop_GuyDiagNumBolts
        End Get
        Set
            Me.prop_GuyDiagNumBolts = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guydiagonaloutofplanerestraint")>
    Public Property GuyDiagonalOutOfPlaneRestraint() As Boolean
        Get
            Return Me.prop_GuyDiagonalOutOfPlaneRestraint
        End Get
        Set
            Me.prop_GuyDiagonalOutOfPlaneRestraint = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guydiagboltgrade")>
    Public Property GuyDiagBoltGrade() As String
        Get
            Return Me.prop_GuyDiagBoltGrade
        End Get
        Set
            Me.prop_GuyDiagBoltGrade = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guydiagboltsize")>
    Public Property GuyDiagBoltSize() As Double
        Get
            Return Me.prop_GuyDiagBoltSize
        End Get
        Set
            Me.prop_GuyDiagBoltSize = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guydiagboltedgedistance")>
    Public Property GuyDiagBoltEdgeDistance() As Double
        Get
            Return Me.prop_GuyDiagBoltEdgeDistance
        End Get
        Set
            Me.prop_GuyDiagBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guydiagboltgagedistance")>
    Public Property GuyDiagBoltGageDistance() As Double
        Get
            Return Me.prop_GuyDiagBoltGageDistance
        End Get
        Set
            Me.prop_GuyDiagBoltGageDistance = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypulloffnetwidthdeduct")>
    Public Property GuyPullOffNetWidthDeduct() As Double
        Get
            Return Me.prop_GuyPullOffNetWidthDeduct
        End Get
        Set
            Me.prop_GuyPullOffNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypulloffufactor")>
    Public Property GuyPullOffUFactor() As Double
        Get
            Return Me.prop_GuyPullOffUFactor
        End Get
        Set
            Me.prop_GuyPullOffUFactor = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypulloffnumbolts")>
    Public Property GuyPullOffNumBolts() As Integer
        Get
            Return Me.prop_GuyPullOffNumBolts
        End Get
        Set
            Me.prop_GuyPullOffNumBolts = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypulloffoutofplanerestraint")>
    Public Property GuyPullOffOutOfPlaneRestraint() As Boolean
        Get
            Return Me.prop_GuyPullOffOutOfPlaneRestraint
        End Get
        Set
            Me.prop_GuyPullOffOutOfPlaneRestraint = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypulloffboltgrade")>
    Public Property GuyPullOffBoltGrade() As String
        Get
            Return Me.prop_GuyPullOffBoltGrade
        End Get
        Set
            Me.prop_GuyPullOffBoltGrade = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypulloffboltsize")>
    Public Property GuyPullOffBoltSize() As Double
        Get
            Return Me.prop_GuyPullOffBoltSize
        End Get
        Set
            Me.prop_GuyPullOffBoltSize = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypulloffboltedgedistance")>
    Public Property GuyPullOffBoltEdgeDistance() As Double
        Get
            Return Me.prop_GuyPullOffBoltEdgeDistance
        End Get
        Set
            Me.prop_GuyPullOffBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypulloffboltgagedistance")>
    Public Property GuyPullOffBoltGageDistance() As Double
        Get
            Return Me.prop_GuyPullOffBoltGageDistance
        End Get
        Set
            Me.prop_GuyPullOffBoltGageDistance = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guytorquearmnetwidthdeduct")>
    Public Property GuyTorqueArmNetWidthDeduct() As Double
        Get
            Return Me.prop_GuyTorqueArmNetWidthDeduct
        End Get
        Set
            Me.prop_GuyTorqueArmNetWidthDeduct = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guytorquearmufactor")>
    Public Property GuyTorqueArmUFactor() As Double
        Get
            Return Me.prop_GuyTorqueArmUFactor
        End Get
        Set
            Me.prop_GuyTorqueArmUFactor = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guytorquearmnumbolts")>
    Public Property GuyTorqueArmNumBolts() As Integer
        Get
            Return Me.prop_GuyTorqueArmNumBolts
        End Get
        Set
            Me.prop_GuyTorqueArmNumBolts = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guytorquearmoutofplanerestraint")>
    Public Property GuyTorqueArmOutOfPlaneRestraint() As Boolean
        Get
            Return Me.prop_GuyTorqueArmOutOfPlaneRestraint
        End Get
        Set
            Me.prop_GuyTorqueArmOutOfPlaneRestraint = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guytorquearmboltgrade")>
    Public Property GuyTorqueArmBoltGrade() As String
        Get
            Return Me.prop_GuyTorqueArmBoltGrade
        End Get
        Set
            Me.prop_GuyTorqueArmBoltGrade = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guytorquearmboltsize")>
    Public Property GuyTorqueArmBoltSize() As Double
        Get
            Return Me.prop_GuyTorqueArmBoltSize
        End Get
        Set
            Me.prop_GuyTorqueArmBoltSize = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guytorquearmboltedgedistance")>
    Public Property GuyTorqueArmBoltEdgeDistance() As Double
        Get
            Return Me.prop_GuyTorqueArmBoltEdgeDistance
        End Get
        Set
            Me.prop_GuyTorqueArmBoltEdgeDistance = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guytorquearmboltgagedistance")>
    Public Property GuyTorqueArmBoltGageDistance() As Double
        Get
            Return Me.prop_GuyTorqueArmBoltGageDistance
        End Get
        Set
            Me.prop_GuyTorqueArmBoltGageDistance = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypercenttension")>
    Public Property GuyPerCentTension() As Double
        Get
            Return Me.prop_GuyPerCentTension
        End Get
        Set
            Me.prop_GuyPerCentTension = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypercenttension120")>
    Public Property GuyPerCentTension120() As Double
        Get
            Return Me.prop_GuyPerCentTension120
        End Get
        Set
            Me.prop_GuyPerCentTension120 = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypercenttension240")>
    Public Property GuyPerCentTension240() As Double
        Get
            Return Me.prop_GuyPerCentTension240
        End Get
        Set
            Me.prop_GuyPerCentTension240 = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guypercenttension360")>
    Public Property GuyPerCentTension360() As Double
        Get
            Return Me.prop_GuyPerCentTension360
        End Get
        Set
            Me.prop_GuyPerCentTension360 = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyefffactor")>
    Public Property GuyEffFactor() As Double
        Get
            Return Me.prop_GuyEffFactor
        End Get
        Set
            Me.prop_GuyEffFactor = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyefffactor120")>
    Public Property GuyEffFactor120() As Double
        Get
            Return Me.prop_GuyEffFactor120
        End Get
        Set
            Me.prop_GuyEffFactor120 = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyefffactor240")>
    Public Property GuyEffFactor240() As Double
        Get
            Return Me.prop_GuyEffFactor240
        End Get
        Set
            Me.prop_GuyEffFactor240 = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyefffactor360")>
    Public Property GuyEffFactor360() As Double
        Get
            Return Me.prop_GuyEffFactor360
        End Get
        Set
            Me.prop_GuyEffFactor360 = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guynuminsulators")>
    Public Property GuyNumInsulators() As Integer
        Get
            Return Me.prop_GuyNumInsulators
        End Get
        Set
            Me.prop_GuyNumInsulators = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyinsulatorlength")>
    Public Property GuyInsulatorLength() As Double
        Get
            Return Me.prop_GuyInsulatorLength
        End Get
        Set
            Me.prop_GuyInsulatorLength = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyinsulatordia")>
    Public Property GuyInsulatorDia() As Double
        Get
            Return Me.prop_GuyInsulatorDia
        End Get
        Set
            Me.prop_GuyInsulatorDia = Value
        End Set
    End Property
    <Category("TNX Guy Record"), Description(""), DisplayName("Guyinsulatorwt")>
    Public Property GuyInsulatorWt() As Double
        Get
            Return Me.prop_GuyInsulatorWt
        End Get
        Set
            Me.prop_GuyInsulatorWt = Value
        End Set
    End Property

End Class
#End Region

#Region "Loading"

Partial Public Class tnxFeedLine
    Private prop_FeedLineRec As Integer
    Private prop_FeedLineEnabled As Boolean
    Private prop_FeedLineDatabase As String
    Private prop_FeedLineDescription As String
    Private prop_FeedLineClassificationCategory As String
    Private prop_FeedLineNote As String
    Private prop_FeedLineNum As Integer
    Private prop_FeedLineUseShielding As Boolean
    Private prop_ExcludeFeedLineFromTorque As Boolean
    Private prop_FeedLineNumPerRow As Integer
    Private prop_FeedLineFace As Integer
    Private prop_FeedLineComponentType As String
    Private prop_FeedLineGroupTreatmentType As String
    Private prop_FeedLineRoundClusterDia As Double
    Private prop_FeedLineWidth As Double
    Private prop_FeedLinePerimeter As Double
    Private prop_FlatAttachmentEffectiveWidthRatio As Double
    Private prop_AutoCalcFlatAttachmentEffectiveWidthRatio As Boolean
    Private prop_FeedLineShieldingFactorKaNoIce As Double
    Private prop_FeedLineShieldingFactorKaIce As Double
    Private prop_FeedLineAutoCalcKa As Boolean
    Private prop_FeedLineCaAaNoIce As Double
    Private prop_FeedLineCaAaIce As Double
    Private prop_FeedLineCaAaIce_1 As Double
    Private prop_FeedLineCaAaIce_2 As Double
    Private prop_FeedLineCaAaIce_4 As Double
    Private prop_FeedLineWtNoIce As Double
    Private prop_FeedLineWtIce As Double
    Private prop_FeedLineWtIce_1 As Double
    Private prop_FeedLineWtIce_2 As Double
    Private prop_FeedLineWtIce_4 As Double
    Private prop_FeedLineFaceOffset As Double
    Private prop_FeedLineOffsetFrac As Double
    Private prop_FeedLinePerimeterOffsetStartFrac As Double
    Private prop_FeedLinePerimeterOffsetEndFrac As Double
    Private prop_FeedLineStartHt As Double
    Private prop_FeedLineEndHt As Double
    Private prop_FeedLineClearSpacing As Double
    Private prop_FeedLineRowClearSpacing As Double

    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineRec")>
    Public Property FeedLineRec() As Integer
        Get
            Return Me.prop_FeedLineRec
        End Get
        Set
            Me.prop_FeedLineRec = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineEnabled")>
    Public Property FeedLineEnabled() As Boolean
        Get
            Return Me.prop_FeedLineEnabled
        End Get
        Set
            Me.prop_FeedLineEnabled = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineDatabase")>
    Public Property FeedLineDatabase() As String
        Get
            Return Me.prop_FeedLineDatabase
        End Get
        Set
            Me.prop_FeedLineDatabase = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineDescription")>
    Public Property FeedLineDescription() As String
        Get
            Return Me.prop_FeedLineDescription
        End Get
        Set
            Me.prop_FeedLineDescription = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineClassificationCategory")>
    Public Property FeedLineClassificationCategory() As String
        Get
            Return Me.prop_FeedLineClassificationCategory
        End Get
        Set
            Me.prop_FeedLineClassificationCategory = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineNote")>
    Public Property FeedLineNote() As String
        Get
            Return Me.prop_FeedLineNote
        End Get
        Set
            Me.prop_FeedLineNote = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineNum")>
    Public Property FeedLineNum() As Integer
        Get
            Return Me.prop_FeedLineNum
        End Get
        Set
            Me.prop_FeedLineNum = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineUseShielding")>
    Public Property FeedLineUseShielding() As Boolean
        Get
            Return Me.prop_FeedLineUseShielding
        End Get
        Set
            Me.prop_FeedLineUseShielding = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("ExcludeFeedLineFromTorque")>
    Public Property ExcludeFeedLineFromTorque() As Boolean
        Get
            Return Me.prop_ExcludeFeedLineFromTorque
        End Get
        Set
            Me.prop_ExcludeFeedLineFromTorque = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineNumPerRow")>
    Public Property FeedLineNumPerRow() As Integer
        Get
            Return Me.prop_FeedLineNumPerRow
        End Get
        Set
            Me.prop_FeedLineNumPerRow = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description("{0 = A, 1 = B,  2 = C, 3 = D}"), DisplayName("FeedLineFace")>
    Public Property FeedLineFace() As Integer
        Get
            Return Me.prop_FeedLineFace
        End Get
        Set
            Me.prop_FeedLineFace = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineComponentType")>
    Public Property FeedLineComponentType() As String
        Get
            Return Me.prop_FeedLineComponentType
        End Get
        Set
            Me.prop_FeedLineComponentType = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineGroupTreatmentType")>
    Public Property FeedLineGroupTreatmentType() As String
        Get
            Return Me.prop_FeedLineGroupTreatmentType
        End Get
        Set
            Me.prop_FeedLineGroupTreatmentType = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineRoundClusterDia")>
    Public Property FeedLineRoundClusterDia() As Double
        Get
            Return Me.prop_FeedLineRoundClusterDia
        End Get
        Set
            Me.prop_FeedLineRoundClusterDia = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineWidth")>
    Public Property FeedLineWidth() As Double
        Get
            Return Me.prop_FeedLineWidth
        End Get
        Set
            Me.prop_FeedLineWidth = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLinePerimeter")>
    Public Property FeedLinePerimeter() As Double
        Get
            Return Me.prop_FeedLinePerimeter
        End Get
        Set
            Me.prop_FeedLinePerimeter = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FlatAttachmentEffectiveWidthRatio")>
    Public Property FlatAttachmentEffectiveWidthRatio() As Double
        Get
            Return Me.prop_FlatAttachmentEffectiveWidthRatio
        End Get
        Set
            Me.prop_FlatAttachmentEffectiveWidthRatio = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("AutoCalcFlatAttachmentEffectiveWidthRatio")>
    Public Property AutoCalcFlatAttachmentEffectiveWidthRatio() As Boolean
        Get
            Return Me.prop_AutoCalcFlatAttachmentEffectiveWidthRatio
        End Get
        Set
            Me.prop_AutoCalcFlatAttachmentEffectiveWidthRatio = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineShieldingFactorKaNoIce")>
    Public Property FeedLineShieldingFactorKaNoIce() As Double
        Get
            Return Me.prop_FeedLineShieldingFactorKaNoIce
        End Get
        Set
            Me.prop_FeedLineShieldingFactorKaNoIce = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineShieldingFactorKaIce")>
    Public Property FeedLineShieldingFactorKaIce() As Double
        Get
            Return Me.prop_FeedLineShieldingFactorKaIce
        End Get
        Set
            Me.prop_FeedLineShieldingFactorKaIce = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineAutoCalcKa")>
    Public Property FeedLineAutoCalcKa() As Boolean
        Get
            Return Me.prop_FeedLineAutoCalcKa
        End Get
        Set
            Me.prop_FeedLineAutoCalcKa = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineCaAaNoIce")>
    Public Property FeedLineCaAaNoIce() As Double
        Get
            Return Me.prop_FeedLineCaAaNoIce
        End Get
        Set
            Me.prop_FeedLineCaAaNoIce = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineCaAaIce")>
    Public Property FeedLineCaAaIce() As Double
        Get
            Return Me.prop_FeedLineCaAaIce
        End Get
        Set
            Me.prop_FeedLineCaAaIce = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineCaAaIce_1")>
    Public Property FeedLineCaAaIce_1() As Double
        Get
            Return Me.prop_FeedLineCaAaIce_1
        End Get
        Set
            Me.prop_FeedLineCaAaIce_1 = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineCaAaIce_2")>
    Public Property FeedLineCaAaIce_2() As Double
        Get
            Return Me.prop_FeedLineCaAaIce_2
        End Get
        Set
            Me.prop_FeedLineCaAaIce_2 = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineCaAaIce_4")>
    Public Property FeedLineCaAaIce_4() As Double
        Get
            Return Me.prop_FeedLineCaAaIce_4
        End Get
        Set
            Me.prop_FeedLineCaAaIce_4 = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineWtNoIce")>
    Public Property FeedLineWtNoIce() As Double
        Get
            Return Me.prop_FeedLineWtNoIce
        End Get
        Set
            Me.prop_FeedLineWtNoIce = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineWtIce")>
    Public Property FeedLineWtIce() As Double
        Get
            Return Me.prop_FeedLineWtIce
        End Get
        Set
            Me.prop_FeedLineWtIce = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineWtIce_1")>
    Public Property FeedLineWtIce_1() As Double
        Get
            Return Me.prop_FeedLineWtIce_1
        End Get
        Set
            Me.prop_FeedLineWtIce_1 = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineWtIce_2")>
    Public Property FeedLineWtIce_2() As Double
        Get
            Return Me.prop_FeedLineWtIce_2
        End Get
        Set
            Me.prop_FeedLineWtIce_2 = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineWtIce_4")>
    Public Property FeedLineWtIce_4() As Double
        Get
            Return Me.prop_FeedLineWtIce_4
        End Get
        Set
            Me.prop_FeedLineWtIce_4 = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineFaceOffset")>
    Public Property FeedLineFaceOffset() As Double
        Get
            Return Me.prop_FeedLineFaceOffset
        End Get
        Set
            Me.prop_FeedLineFaceOffset = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineOffsetFrac")>
    Public Property FeedLineOffsetFrac() As Double
        Get
            Return Me.prop_FeedLineOffsetFrac
        End Get
        Set
            Me.prop_FeedLineOffsetFrac = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLinePerimeterOffsetStartFrac")>
    Public Property FeedLinePerimeterOffsetStartFrac() As Double
        Get
            Return Me.prop_FeedLinePerimeterOffsetStartFrac
        End Get
        Set
            Me.prop_FeedLinePerimeterOffsetStartFrac = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLinePerimeterOffsetEndFrac")>
    Public Property FeedLinePerimeterOffsetEndFrac() As Double
        Get
            Return Me.prop_FeedLinePerimeterOffsetEndFrac
        End Get
        Set
            Me.prop_FeedLinePerimeterOffsetEndFrac = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineStartHt")>
    Public Property FeedLineStartHt() As Double
        Get
            Return Me.prop_FeedLineStartHt
        End Get
        Set
            Me.prop_FeedLineStartHt = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineEndHt")>
    Public Property FeedLineEndHt() As Double
        Get
            Return Me.prop_FeedLineEndHt
        End Get
        Set
            Me.prop_FeedLineEndHt = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineClearSpacing")>
    Public Property FeedLineClearSpacing() As Double
        Get
            Return Me.prop_FeedLineClearSpacing
        End Get
        Set
            Me.prop_FeedLineClearSpacing = Value
        End Set
    End Property
    <Category("TNX Feed Lines"), Description(""), DisplayName("FeedLineRowClearSpacing")>
    Public Property FeedLineRowClearSpacing() As Double
        Get
            Return Me.prop_FeedLineRowClearSpacing
        End Get
        Set
            Me.prop_FeedLineRowClearSpacing = Value
        End Set
    End Property


End Class
Partial Public Class tnxDiscreteLoad
    Private prop_TowerLoadRec As Integer
    Private prop_TowerLoadEnabled As Boolean
    Private prop_TowerLoadDatabase As String
    Private prop_TowerLoadDescription As String
    Private prop_TowerLoadType As String
    Private prop_TowerLoadClassificationCategory As String
    Private prop_TowerLoadNote As String
    Private prop_TowerLoadNum As Integer
    Private prop_TowerLoadFace As Integer
    Private prop_TowerOffsetType As String
    Private prop_TowerOffsetDist As Double
    Private prop_TowerVertOffset As Double
    Private prop_TowerLateralOffset As Double
    Private prop_TowerAzimuthAdjustment As Double
    Private prop_TowerAppurtSymbol As String
    Private prop_TowerLoadShieldingFactorKaNoIce As Double
    Private prop_TowerLoadShieldingFactorKaIce As Double
    Private prop_TowerLoadAutoCalcKa As Boolean
    Private prop_TowerLoadCaAaNoIce As Double
    Private prop_TowerLoadCaAaIce As Double
    Private prop_TowerLoadCaAaIce_1 As Double
    Private prop_TowerLoadCaAaIce_2 As Double
    Private prop_TowerLoadCaAaIce_4 As Double
    Private prop_TowerLoadCaAaNoIce_Side As Double
    Private prop_TowerLoadCaAaIce_Side As Double
    Private prop_TowerLoadCaAaIce_Side_1 As Double
    Private prop_TowerLoadCaAaIce_Side_2 As Double
    Private prop_TowerLoadCaAaIce_Side_4 As Double
    Private prop_TowerLoadWtNoIce As Double
    Private prop_TowerLoadWtIce As Double
    Private prop_TowerLoadWtIce_1 As Double
    Private prop_TowerLoadWtIce_2 As Double
    Private prop_TowerLoadWtIce_4 As Double
    Private prop_TowerLoadStartHt As Double
    Private prop_TowerLoadEndHt As Double


    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadRec")>
    Public Property TowerLoadRec() As Integer
        Get
            Return Me.prop_TowerLoadRec
        End Get
        Set
            Me.prop_TowerLoadRec = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadEnabled")>
    Public Property TowerLoadEnabled() As Boolean
        Get
            Return Me.prop_TowerLoadEnabled
        End Get
        Set
            Me.prop_TowerLoadEnabled = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadDatabase")>
    Public Property TowerLoadDatabase() As String
        Get
            Return Me.prop_TowerLoadDatabase
        End Get
        Set
            Me.prop_TowerLoadDatabase = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadDescription")>
    Public Property TowerLoadDescription() As String
        Get
            Return Me.prop_TowerLoadDescription
        End Get
        Set
            Me.prop_TowerLoadDescription = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadType")>
    Public Property TowerLoadType() As String
        Get
            Return Me.prop_TowerLoadType
        End Get
        Set
            Me.prop_TowerLoadType = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadClassificationCategory")>
    Public Property TowerLoadClassificationCategory() As String
        Get
            Return Me.prop_TowerLoadClassificationCategory
        End Get
        Set
            Me.prop_TowerLoadClassificationCategory = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadNote")>
    Public Property TowerLoadNote() As String
        Get
            Return Me.prop_TowerLoadNote
        End Get
        Set
            Me.prop_TowerLoadNote = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadNum")>
    Public Property TowerLoadNum() As Integer
        Get
            Return Me.prop_TowerLoadNum
        End Get
        Set
            Me.prop_TowerLoadNum = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description("{0 = A, 1 = B,  2 = C, 3 = D}"), DisplayName("TowerLoadFace")>
    Public Property TowerLoadFace() As Integer
        Get
            Return Me.prop_TowerLoadFace
        End Get
        Set
            Me.prop_TowerLoadFace = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerOffsetType")>
    Public Property TowerOffsetType() As String
        Get
            Return Me.prop_TowerOffsetType
        End Get
        Set
            Me.prop_TowerOffsetType = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerOffsetDist")>
    Public Property TowerOffsetDist() As Double
        Get
            Return Me.prop_TowerOffsetDist
        End Get
        Set
            Me.prop_TowerOffsetDist = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerVertOffset")>
    Public Property TowerVertOffset() As Double
        Get
            Return Me.prop_TowerVertOffset
        End Get
        Set
            Me.prop_TowerVertOffset = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLateralOffset")>
    Public Property TowerLateralOffset() As Double
        Get
            Return Me.prop_TowerLateralOffset
        End Get
        Set
            Me.prop_TowerLateralOffset = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerAzimuthAdjustment")>
    Public Property TowerAzimuthAdjustment() As Double
        Get
            Return Me.prop_TowerAzimuthAdjustment
        End Get
        Set
            Me.prop_TowerAzimuthAdjustment = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerAppurtSymbol")>
    Public Property TowerAppurtSymbol() As String
        Get
            Return Me.prop_TowerAppurtSymbol
        End Get
        Set
            Me.prop_TowerAppurtSymbol = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadShieldingFactorKaNoIce")>
    Public Property TowerLoadShieldingFactorKaNoIce() As Double
        Get
            Return Me.prop_TowerLoadShieldingFactorKaNoIce
        End Get
        Set
            Me.prop_TowerLoadShieldingFactorKaNoIce = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadShieldingFactorKaIce")>
    Public Property TowerLoadShieldingFactorKaIce() As Double
        Get
            Return Me.prop_TowerLoadShieldingFactorKaIce
        End Get
        Set
            Me.prop_TowerLoadShieldingFactorKaIce = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadAutoCalcKa")>
    Public Property TowerLoadAutoCalcKa() As Boolean
        Get
            Return Me.prop_TowerLoadAutoCalcKa
        End Get
        Set
            Me.prop_TowerLoadAutoCalcKa = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadCaAaNoIce")>
    Public Property TowerLoadCaAaNoIce() As Double
        Get
            Return Me.prop_TowerLoadCaAaNoIce
        End Get
        Set
            Me.prop_TowerLoadCaAaNoIce = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadCaAaIce")>
    Public Property TowerLoadCaAaIce() As Double
        Get
            Return Me.prop_TowerLoadCaAaIce
        End Get
        Set
            Me.prop_TowerLoadCaAaIce = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadCaAaIce_1")>
    Public Property TowerLoadCaAaIce_1() As Double
        Get
            Return Me.prop_TowerLoadCaAaIce_1
        End Get
        Set
            Me.prop_TowerLoadCaAaIce_1 = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadCaAaIce_2")>
    Public Property TowerLoadCaAaIce_2() As Double
        Get
            Return Me.prop_TowerLoadCaAaIce_2
        End Get
        Set
            Me.prop_TowerLoadCaAaIce_2 = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadCaAaIce_4")>
    Public Property TowerLoadCaAaIce_4() As Double
        Get
            Return Me.prop_TowerLoadCaAaIce_4
        End Get
        Set
            Me.prop_TowerLoadCaAaIce_4 = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadCaAaNoIce_Side")>
    Public Property TowerLoadCaAaNoIce_Side() As Double
        Get
            Return Me.prop_TowerLoadCaAaNoIce_Side
        End Get
        Set
            Me.prop_TowerLoadCaAaNoIce_Side = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadCaAaIce_Side")>
    Public Property TowerLoadCaAaIce_Side() As Double
        Get
            Return Me.prop_TowerLoadCaAaIce_Side
        End Get
        Set
            Me.prop_TowerLoadCaAaIce_Side = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadCaAaIce_Side_1")>
    Public Property TowerLoadCaAaIce_Side_1() As Double
        Get
            Return Me.prop_TowerLoadCaAaIce_Side_1
        End Get
        Set
            Me.prop_TowerLoadCaAaIce_Side_1 = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadCaAaIce_Side_2")>
    Public Property TowerLoadCaAaIce_Side_2() As Double
        Get
            Return Me.prop_TowerLoadCaAaIce_Side_2
        End Get
        Set
            Me.prop_TowerLoadCaAaIce_Side_2 = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadCaAaIce_Side_4")>
    Public Property TowerLoadCaAaIce_Side_4() As Double
        Get
            Return Me.prop_TowerLoadCaAaIce_Side_4
        End Get
        Set
            Me.prop_TowerLoadCaAaIce_Side_4 = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadWtNoIce")>
    Public Property TowerLoadWtNoIce() As Double
        Get
            Return Me.prop_TowerLoadWtNoIce
        End Get
        Set
            Me.prop_TowerLoadWtNoIce = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadWtIce")>
    Public Property TowerLoadWtIce() As Double
        Get
            Return Me.prop_TowerLoadWtIce
        End Get
        Set
            Me.prop_TowerLoadWtIce = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadWtIce_1")>
    Public Property TowerLoadWtIce_1() As Double
        Get
            Return Me.prop_TowerLoadWtIce_1
        End Get
        Set
            Me.prop_TowerLoadWtIce_1 = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadWtIce_2")>
    Public Property TowerLoadWtIce_2() As Double
        Get
            Return Me.prop_TowerLoadWtIce_2
        End Get
        Set
            Me.prop_TowerLoadWtIce_2 = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadWtIce_4")>
    Public Property TowerLoadWtIce_4() As Double
        Get
            Return Me.prop_TowerLoadWtIce_4
        End Get
        Set
            Me.prop_TowerLoadWtIce_4 = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadStartHt")>
    Public Property TowerLoadStartHt() As Double
        Get
            Return Me.prop_TowerLoadStartHt
        End Get
        Set
            Me.prop_TowerLoadStartHt = Value
        End Set
    End Property
    <Category("TNX Discrete Load"), Description(""), DisplayName("TowerLoadEndHt")>
    Public Property TowerLoadEndHt() As Double
        Get
            Return Me.prop_TowerLoadEndHt
        End Get
        Set
            Me.prop_TowerLoadEndHt = Value
        End Set
    End Property

End Class

Partial Public Class tnxDish
    Private prop_DishRec As Integer
    Private prop_DishEnabled As Boolean
    Private prop_DishDatabase As String
    Private prop_DishDescription As String
    Private prop_DishClassificationCategory As String
    Private prop_DishNote As String
    Private prop_DishNum As Integer
    Private prop_DishFace As Integer
    Private prop_DishType As String
    Private prop_DishOffsetType As String
    Private prop_DishVertOffset As Double
    Private prop_DishLateralOffset As Double
    Private prop_DishOffsetDist As Double
    Private prop_DishArea As Double
    Private prop_DishAreaIce As Double
    Private prop_DishAreaIce_1 As Double
    Private prop_DishAreaIce_2 As Double
    Private prop_DishAreaIce_4 As Double
    Private prop_DishDiameter As Double
    Private prop_DishWtNoIce As Double
    Private prop_DishWtIce As Double
    Private prop_DishWtIce_1 As Double
    Private prop_DishWtIce_2 As Double
    Private prop_DishWtIce_4 As Double
    Private prop_DishStartHt As Double
    Private prop_DishAzimuthAdjustment As Double
    Private prop_DishBeamWidth As Double

    <Category("TNX Dish"), Description(""), DisplayName("DishRec")>
    Public Property DishRec() As Integer
        Get
            Return Me.prop_DishRec
        End Get
        Set
            Me.prop_DishRec = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishEnabled")>
    Public Property DishEnabled() As Boolean
        Get
            Return Me.prop_DishEnabled
        End Get
        Set
            Me.prop_DishEnabled = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishDatabase")>
    Public Property DishDatabase() As String
        Get
            Return Me.prop_DishDatabase
        End Get
        Set
            Me.prop_DishDatabase = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishDescription")>
    Public Property DishDescription() As String
        Get
            Return Me.prop_DishDescription
        End Get
        Set
            Me.prop_DishDescription = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishClassificationCategory")>
    Public Property DishClassificationCategory() As String
        Get
            Return Me.prop_DishClassificationCategory
        End Get
        Set
            Me.prop_DishClassificationCategory = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishNote")>
    Public Property DishNote() As String
        Get
            Return Me.prop_DishNote
        End Get
        Set
            Me.prop_DishNote = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishNum")>
    Public Property DishNum() As Integer
        Get
            Return Me.prop_DishNum
        End Get
        Set
            Me.prop_DishNum = Value
        End Set
    End Property
    <Category("TNX Dish"), Description("{0 = A, 1 = B,  2 = C, 3 = D}"), DisplayName("DishFace")>
    Public Property DishFace() As Integer
        Get
            Return Me.prop_DishFace
        End Get
        Set
            Me.prop_DishFace = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishType")>
    Public Property DishType() As String
        Get
            Return Me.prop_DishType
        End Get
        Set
            Me.prop_DishType = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishOffsetType")>
    Public Property DishOffsetType() As String
        Get
            Return Me.prop_DishOffsetType
        End Get
        Set
            Me.prop_DishOffsetType = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishVertOffset")>
    Public Property DishVertOffset() As Double
        Get
            Return Me.prop_DishVertOffset
        End Get
        Set
            Me.prop_DishVertOffset = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishLateralOffset")>
    Public Property DishLateralOffset() As Double
        Get
            Return Me.prop_DishLateralOffset
        End Get
        Set
            Me.prop_DishLateralOffset = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishOffsetDist")>
    Public Property DishOffsetDist() As Double
        Get
            Return Me.prop_DishOffsetDist
        End Get
        Set
            Me.prop_DishOffsetDist = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishArea")>
    Public Property DishArea() As Double
        Get
            Return Me.prop_DishArea
        End Get
        Set
            Me.prop_DishArea = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishAreaIce")>
    Public Property DishAreaIce() As Double
        Get
            Return Me.prop_DishAreaIce
        End Get
        Set
            Me.prop_DishAreaIce = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishAreaIce_1")>
    Public Property DishAreaIce_1() As Double
        Get
            Return Me.prop_DishAreaIce_1
        End Get
        Set
            Me.prop_DishAreaIce_1 = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishAreaIce_2")>
    Public Property DishAreaIce_2() As Double
        Get
            Return Me.prop_DishAreaIce_2
        End Get
        Set
            Me.prop_DishAreaIce_2 = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishAreaIce_4")>
    Public Property DishAreaIce_4() As Double
        Get
            Return Me.prop_DishAreaIce_4
        End Get
        Set
            Me.prop_DishAreaIce_4 = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishDiameter")>
    Public Property DishDiameter() As Double
        Get
            Return Me.prop_DishDiameter
        End Get
        Set
            Me.prop_DishDiameter = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishWtNoIce")>
    Public Property DishWtNoIce() As Double
        Get
            Return Me.prop_DishWtNoIce
        End Get
        Set
            Me.prop_DishWtNoIce = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishWtIce")>
    Public Property DishWtIce() As Double
        Get
            Return Me.prop_DishWtIce
        End Get
        Set
            Me.prop_DishWtIce = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishWtIce_1")>
    Public Property DishWtIce_1() As Double
        Get
            Return Me.prop_DishWtIce_1
        End Get
        Set
            Me.prop_DishWtIce_1 = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishWtIce_2")>
    Public Property DishWtIce_2() As Double
        Get
            Return Me.prop_DishWtIce_2
        End Get
        Set
            Me.prop_DishWtIce_2 = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishWtIce_4")>
    Public Property DishWtIce_4() As Double
        Get
            Return Me.prop_DishWtIce_4
        End Get
        Set
            Me.prop_DishWtIce_4 = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishStartHt")>
    Public Property DishStartHt() As Double
        Get
            Return Me.prop_DishStartHt
        End Get
        Set
            Me.prop_DishStartHt = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishAzimuthAdjustment")>
    Public Property DishAzimuthAdjustment() As Double
        Get
            Return Me.prop_DishAzimuthAdjustment
        End Get
        Set
            Me.prop_DishAzimuthAdjustment = Value
        End Set
    End Property
    <Category("TNX Dish"), Description(""), DisplayName("DishBeamWidth")>
    Public Property DishBeamWidth() As Double
        Get
            Return Me.prop_DishBeamWidth
        End Get
        Set
            Me.prop_DishBeamWidth = Value
        End Set
    End Property

End Class

Partial Public Class tnxUserForce
    Private prop_UserForceRec As Integer
    Private prop_UserForceEnabled As Boolean
    Private prop_UserForceDescription As String
    Private prop_UserForceStartHt As Double
    Private prop_UserForceOffset As Double
    Private prop_UserForceAzimuth As Double
    Private prop_UserForceFxNoIce As Double
    Private prop_UserForceFzNoIce As Double
    Private prop_UserForceAxialNoIce As Double
    Private prop_UserForceShearNoIce As Double
    Private prop_UserForceCaAcNoIce As Double
    Private prop_UserForceFxIce As Double
    Private prop_UserForceFzIce As Double
    Private prop_UserForceAxialIce As Double
    Private prop_UserForceShearIce As Double
    Private prop_UserForceCaAcIce As Double
    Private prop_UserForceFxService As Double
    Private prop_UserForceFzService As Double
    Private prop_UserForceAxialService As Double
    Private prop_UserForceShearService As Double
    Private prop_UserForceCaAcService As Double
    Private prop_UserForceEhx As Double
    Private prop_UserForceEhz As Double
    Private prop_UserForceEv As Double
    Private prop_UserForceEh As Double

    <Category("TNX User Force"), Description(""), DisplayName("UserForceRec")>
    Public Property UserForceRec() As Integer
        Get
            Return Me.prop_UserForceRec
        End Get
        Set
            Me.prop_UserForceRec = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceEnabled")>
    Public Property UserForceEnabled() As Boolean
        Get
            Return Me.prop_UserForceEnabled
        End Get
        Set
            Me.prop_UserForceEnabled = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceDescription")>
    Public Property UserForceDescription() As String
        Get
            Return Me.prop_UserForceDescription
        End Get
        Set
            Me.prop_UserForceDescription = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceStartHt")>
    Public Property UserForceStartHt() As Double
        Get
            Return Me.prop_UserForceStartHt
        End Get
        Set
            Me.prop_UserForceStartHt = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceOffset")>
    Public Property UserForceOffset() As Double
        Get
            Return Me.prop_UserForceOffset
        End Get
        Set
            Me.prop_UserForceOffset = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceAzimuth")>
    Public Property UserForceAzimuth() As Double
        Get
            Return Me.prop_UserForceAzimuth
        End Get
        Set
            Me.prop_UserForceAzimuth = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceFxNoIce")>
    Public Property UserForceFxNoIce() As Double
        Get
            Return Me.prop_UserForceFxNoIce
        End Get
        Set
            Me.prop_UserForceFxNoIce = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceFzNoIce")>
    Public Property UserForceFzNoIce() As Double
        Get
            Return Me.prop_UserForceFzNoIce
        End Get
        Set
            Me.prop_UserForceFzNoIce = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceAxialNoIce")>
    Public Property UserForceAxialNoIce() As Double
        Get
            Return Me.prop_UserForceAxialNoIce
        End Get
        Set
            Me.prop_UserForceAxialNoIce = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceShearNoIce")>
    Public Property UserForceShearNoIce() As Double
        Get
            Return Me.prop_UserForceShearNoIce
        End Get
        Set
            Me.prop_UserForceShearNoIce = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceCaAcNoIce")>
    Public Property UserForceCaAcNoIce() As Double
        Get
            Return Me.prop_UserForceCaAcNoIce
        End Get
        Set
            Me.prop_UserForceCaAcNoIce = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceFxIce")>
    Public Property UserForceFxIce() As Double
        Get
            Return Me.prop_UserForceFxIce
        End Get
        Set
            Me.prop_UserForceFxIce = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceFzIce")>
    Public Property UserForceFzIce() As Double
        Get
            Return Me.prop_UserForceFzIce
        End Get
        Set
            Me.prop_UserForceFzIce = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceAxialIce")>
    Public Property UserForceAxialIce() As Double
        Get
            Return Me.prop_UserForceAxialIce
        End Get
        Set
            Me.prop_UserForceAxialIce = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceShearIce")>
    Public Property UserForceShearIce() As Double
        Get
            Return Me.prop_UserForceShearIce
        End Get
        Set
            Me.prop_UserForceShearIce = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceCaAcIce")>
    Public Property UserForceCaAcIce() As Double
        Get
            Return Me.prop_UserForceCaAcIce
        End Get
        Set
            Me.prop_UserForceCaAcIce = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceFxService")>
    Public Property UserForceFxService() As Double
        Get
            Return Me.prop_UserForceFxService
        End Get
        Set
            Me.prop_UserForceFxService = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceFzService")>
    Public Property UserForceFzService() As Double
        Get
            Return Me.prop_UserForceFzService
        End Get
        Set
            Me.prop_UserForceFzService = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceAxialService")>
    Public Property UserForceAxialService() As Double
        Get
            Return Me.prop_UserForceAxialService
        End Get
        Set
            Me.prop_UserForceAxialService = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceShearService")>
    Public Property UserForceShearService() As Double
        Get
            Return Me.prop_UserForceShearService
        End Get
        Set
            Me.prop_UserForceShearService = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceCaAcService")>
    Public Property UserForceCaAcService() As Double
        Get
            Return Me.prop_UserForceCaAcService
        End Get
        Set
            Me.prop_UserForceCaAcService = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceEhx")>
    Public Property UserForceEhx() As Double
        Get
            Return Me.prop_UserForceEhx
        End Get
        Set
            Me.prop_UserForceEhx = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceEhz")>
    Public Property UserForceEhz() As Double
        Get
            Return Me.prop_UserForceEhz
        End Get
        Set
            Me.prop_UserForceEhz = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceEv")>
    Public Property UserForceEv() As Double
        Get
            Return Me.prop_UserForceEv
        End Get
        Set
            Me.prop_UserForceEv = Value
        End Set
    End Property
    <Category("TNX User Force"), Description(""), DisplayName("UserForceEh")>
    Public Property UserForceEh() As Double
        Get
            Return Me.prop_UserForceEh
        End Get
        Set
            Me.prop_UserForceEh = Value
        End Set
    End Property

End Class

#End Region

#Region "Code"
Partial Public Class tnxCode
    Private prop_design As New tnxDesign()
    Private prop_ice As New tnxIce()
    Private prop_thermal As New tnxThermal()
    Private prop_wind As New tnxWind()
    Private prop_misclCode As New tnxMisclCode()
    Private prop_seismic As New tnxSeismic()

    <Category("TNX Code"), Description(""), DisplayName("Design")>
    Public Property design() As tnxDesign
        Get
            Return Me.prop_design
        End Get
        Set
            Me.prop_design = Value
        End Set
    End Property

    <Category("TNX Code"), Description(""), DisplayName("Ice")>
    Public Property ice() As tnxIce
        Get
            Return Me.prop_ice
        End Get
        Set
            Me.prop_ice = Value
        End Set
    End Property

    <Category("TNX Code"), Description(""), DisplayName("Thermal")>
    Public Property thermal() As tnxThermal
        Get
            Return Me.prop_thermal
        End Get
        Set
            Me.prop_thermal = Value
        End Set
    End Property

    <Category("TNX Code"), Description(""), DisplayName("Wind")>
    Public Property wind() As tnxWind
        Get
            Return Me.prop_wind
        End Get
        Set
            Me.prop_wind = Value
        End Set
    End Property

    <Category("TNX Code"), Description(""), DisplayName("Miscellaneous Code")>
    Public Property misclCode() As tnxMisclCode
        Get
            Return Me.prop_misclCode
        End Get
        Set
            Me.prop_misclCode = Value
        End Set
    End Property

    <Category("TNX Code"), Description(""), DisplayName("Seismic")>
    Public Property seismic() As tnxSeismic
        Get
            Return Me.prop_seismic
        End Get
        Set
            Me.prop_seismic = Value
        End Set
    End Property

End Class

Partial Public Class tnxDesign
    Private prop_DesignCode As String
    Private prop_ERIDesignMode As String
    Private prop_DoInteraction As Boolean
    Private prop_DoHorzInteraction As Boolean
    Private prop_DoDiagInteraction As Boolean
    Private prop_UseMomentMagnification As Boolean
    Private prop_UseCodeStressRatio As Boolean
    Private prop_AllowStressRatio As Double
    Private prop_AllowAntStressRatio As Double
    Private prop_UseCodeGuySF As Boolean
    Private prop_GuySF As Double
    Private prop_UseTIA222H_AnnexS As Boolean
    Private prop_TIA_222_H_AnnexS_Ratio As Double
    Private prop_PrintBitmaps As Boolean

    <Category("TNX Code Design"), Description(""), DisplayName("DesignCode")>
    Public Property DesignCode() As String
        Get
            Return Me.prop_DesignCode
        End Get
        Set
            Me.prop_DesignCode = Value
        End Set
    End Property
    <Category("TNX Code Design"), Description("Analysis Only, Check Sections, Cyclic Design"), DisplayName("ERIDesignMode")>
    Public Property ERIDesignMode() As String
        Get
            Return Me.prop_ERIDesignMode
        End Get
        Set
            Me.prop_ERIDesignMode = Value
        End Set
    End Property
    <Category("TNX Code Design"), Description("consider moments - legs"), DisplayName("DoInteraction")>
    Public Property DoInteraction() As Boolean
        Get
            Return Me.prop_DoInteraction
        End Get
        Set
            Me.prop_DoInteraction = Value
        End Set
    End Property
    <Category("TNX Code Design"), Description("consider moments - horizontals"), DisplayName("DoHorzInteraction")>
    Public Property DoHorzInteraction() As Boolean
        Get
            Return Me.prop_DoHorzInteraction
        End Get
        Set
            Me.prop_DoHorzInteraction = Value
        End Set
    End Property
    <Category("TNX Code Design"), Description("consider moments - diagonals"), DisplayName("DoDiagInteraction")>
    Public Property DoDiagInteraction() As Boolean
        Get
            Return Me.prop_DoDiagInteraction
        End Get
        Set
            Me.prop_DoDiagInteraction = Value
        End Set
    End Property
    <Category("TNX Code Design"), Description(""), DisplayName("UseMomentMagnification")>
    Public Property UseMomentMagnification() As Boolean
        Get
            Return Me.prop_UseMomentMagnification
        End Get
        Set
            Me.prop_UseMomentMagnification = Value
        End Set
    End Property
    <Category("TNX Code Design"), Description(""), DisplayName("UseCodeStressRatio")>
    Public Property UseCodeStressRatio() As Boolean
        Get
            Return Me.prop_UseCodeStressRatio
        End Get
        Set
            Me.prop_UseCodeStressRatio = Value
        End Set
    End Property
    <Category("TNX Code Design"), Description("base structure allowable stress ratio"), DisplayName("AllowStressRatio")>
    Public Property AllowStressRatio() As Double
        Get
            Return Me.prop_AllowStressRatio
        End Get
        Set
            Me.prop_AllowStressRatio = Value
        End Set
    End Property
    <Category("TNX Code Design"), Description("upper structure allowable stress ratio"), DisplayName("AllowAntStressRatio")>
    Public Property AllowAntStressRatio() As Double
        Get
            Return Me.prop_AllowAntStressRatio
        End Get
        Set
            Me.prop_AllowAntStressRatio = Value
        End Set
    End Property
    <Category("TNX Code Design"), Description(""), DisplayName("UseCodeGuySF")>
    Public Property UseCodeGuySF() As Boolean
        Get
            Return Me.prop_UseCodeGuySF
        End Get
        Set
            Me.prop_UseCodeGuySF = Value
        End Set
    End Property
    <Category("TNX Code Design"), Description(""), DisplayName("GuySF")>
    Public Property GuySF() As Double
        Get
            Return Me.prop_GuySF
        End Get
        Set
            Me.prop_GuySF = Value
        End Set
    End Property
    <Category("TNX Code Design"), Description(""), DisplayName("UseTIA222H_AnnexS")>
    Public Property UseTIA222H_AnnexS() As Boolean
        Get
            Return Me.prop_UseTIA222H_AnnexS
        End Get
        Set
            Me.prop_UseTIA222H_AnnexS = Value
        End Set
    End Property
    <Category("TNX Code Design"), Description("TIA-222-H Annex S allowable ratio"), DisplayName("TIA_222_H_AnnexS_Ratio")>
    Public Property TIA_222_H_AnnexS_Ratio() As Double
        Get
            Return Me.prop_TIA_222_H_AnnexS_Ratio
        End Get
        Set
            Me.prop_TIA_222_H_AnnexS_Ratio = Value
        End Set
    End Property
    <Category("TNX Code Design"), Description(""), DisplayName("PrintBitmaps")>
    Public Property PrintBitmaps() As Boolean
        Get
            Return Me.prop_PrintBitmaps
        End Get
        Set
            Me.prop_PrintBitmaps = Value
        End Set
    End Property

End Class
Partial Public Class tnxIce
    Private prop_IceThickness As Double
    Private prop_IceDensity As Double
    Private prop_UseModified_TIA_222_IceParameters As Boolean
    Private prop_TIA_222_IceThicknessMultiplier As Double
    Private prop_DoNotUse_TIA_222_IceEscalation As Boolean
    Private prop_UseIceEscalation As Boolean

    <Category("TNX Code Ice"), Description(""), DisplayName("IceThickness")>
    Public Property IceThickness() As Double
        Get
            Return Me.prop_IceThickness
        End Get
        Set
            Me.prop_IceThickness = Value
        End Set
    End Property
    <Category("TNX Code Ice"), Description(""), DisplayName("IceDensity")>
    Public Property IceDensity() As Double
        Get
            Return Me.prop_IceDensity
        End Get
        Set
            Me.prop_IceDensity = Value
        End Set
    End Property
    <Category("TNX Code Ice"), Description("TIA-222-G/H Custom Ice Options"), DisplayName("UseModified_TIA_222_IceParameters")>
    Public Property UseModified_TIA_222_IceParameters() As Boolean
        Get
            Return Me.prop_UseModified_TIA_222_IceParameters
        End Get
        Set
            Me.prop_UseModified_TIA_222_IceParameters = Value
        End Set
    End Property
    <Category("TNX Code Ice"), Description("TIA-222-G/H Custom Ice Options"), DisplayName("TIA_222_IceThicknessMultiplier")>
    Public Property TIA_222_IceThicknessMultiplier() As Double
        Get
            Return Me.prop_TIA_222_IceThicknessMultiplier
        End Get
        Set
            Me.prop_TIA_222_IceThicknessMultiplier = Value
        End Set
    End Property
    <Category("TNX Code Ice"), Description("TIA-222-G/H Custom Ice Options"), DisplayName("DoNotUse_TIA_222_IceEscalation")>
    Public Property DoNotUse_TIA_222_IceEscalation() As Boolean
        Get
            Return Me.prop_DoNotUse_TIA_222_IceEscalation
        End Get
        Set
            Me.prop_DoNotUse_TIA_222_IceEscalation = Value
        End Set
    End Property
    <Category("TNX Code Ice"), Description("TIA-222-F and earlier"), DisplayName("UseIceEscalation")>
    Public Property UseIceEscalation() As Boolean
        Get
            Return Me.prop_UseIceEscalation
        End Get
        Set
            Me.prop_UseIceEscalation = Value
        End Set
    End Property
End Class
Partial Public Class tnxThermal
    Private prop_TempDrop As Double

    <Category("TNX Code Thermal"), Description(""), DisplayName("TempDrop")>
    Public Property TempDrop() As Double
        Get
            Return Me.prop_TempDrop
        End Get
        Set
            Me.prop_TempDrop = Value
        End Set
    End Property
End Class
Partial Public Class tnxWind
    Private prop_WindSpeed As Double
    Private prop_WindSpeedIce As Double
    Private prop_WindSpeedService As Double
    Private prop_UseStateCountyLookup As Boolean
    Private prop_State As String
    Private prop_County As String
    Private prop_UseMaxKz As Boolean
    Private prop_ASCE_7_10_WindData As Boolean
    Private prop_ASCE_7_10_ConvertWindToASD As Boolean
    Private prop_UseASCEWind As Boolean
    Private prop_AutoCalc_ASCE_GH As Boolean
    Private prop_ASCE_ExposureCat As Integer
    Private prop_ASCE_Year As Integer
    Private prop_ASCEGh As Double
    Private prop_ASCEI As Double
    Private prop_CalcWindAt As Integer
    Private prop_WindCalcPoints As Double
    Private prop_WindExposure As Integer
    Private prop_StructureCategory As Integer
    Private prop_RiskCategory As Integer
    Private prop_TopoCategory As Integer
    Private prop_RSMTopographicFeature As Integer
    Private prop_RSM_L As Double
    Private prop_RSM_X As Double
    Private prop_CrestHeight As Double
    Private prop_TIA_222_H_TopoFeatureDownwind As Boolean
    Private prop_BaseElevAboveSeaLevel As Double
    Private prop_ConsiderRooftopSpeedUp As Boolean
    Private prop_RooftopWS As Double
    Private prop_RooftopHS As Double
    Private prop_RooftopParapetHt As Double
    Private prop_RooftopXB As Double
    Private prop_WindZone As Integer
    Private prop_EIACWindMult As Double
    Private prop_EIACWindMultIce As Double
    Private prop_EIACIgnoreCableDrag As Boolean
    Private prop_CSA_S37_RefVelPress As Double
    Private prop_CSA_S37_ReliabilityClass As Integer
    Private prop_CSA_S37_ServiceabilityFactor As Double

    <Category("TNX Code Wind"), Description(""), DisplayName("WindSpeed")>
    Public Property WindSpeed() As Double
        Get
            Return Me.prop_WindSpeed
        End Get
        Set
            Me.prop_WindSpeed = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("WindSpeedIce")>
    Public Property WindSpeedIce() As Double
        Get
            Return Me.prop_WindSpeedIce
        End Get
        Set
            Me.prop_WindSpeedIce = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("WindSpeedService")>
    Public Property WindSpeedService() As Double
        Get
            Return Me.prop_WindSpeedService
        End Get
        Set
            Me.prop_WindSpeedService = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("UseStateCountyLookup")>
    Public Property UseStateCountyLookup() As Boolean
        Get
            Return Me.prop_UseStateCountyLookup
        End Get
        Set
            Me.prop_UseStateCountyLookup = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("State")>
    Public Property State() As String
        Get
            Return Me.prop_State
        End Get
        Set
            Me.prop_State = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("County")>
    Public Property County() As String
        Get
            Return Me.prop_County
        End Get
        Set
            Me.prop_County = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("UseMaxKz")>
    Public Property UseMaxKz() As Boolean
        Get
            Return Me.prop_UseMaxKz
        End Get
        Set
            Me.prop_UseMaxKz = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("TIA-222-G Only"), DisplayName("ASCE_7_10_WindData")>
    Public Property ASCE_7_10_WindData() As Boolean
        Get
            Return Me.prop_ASCE_7_10_WindData
        End Get
        Set
            Me.prop_ASCE_7_10_WindData = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("TIA-222-G Only"), DisplayName("ASCE_7_10_ConvertWindToASD")>
    Public Property ASCE_7_10_ConvertWindToASD() As Boolean
        Get
            Return Me.prop_ASCE_7_10_ConvertWindToASD
        End Get
        Set
            Me.prop_ASCE_7_10_ConvertWindToASD = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("Use Special Wind Profile"), DisplayName("UseASCEWind")>
    Public Property UseASCEWind() As Boolean
        Get
            Return Me.prop_UseASCEWind
        End Get
        Set
            Me.prop_UseASCEWind = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("Use TIA Gh Value"), DisplayName("AutoCalc_ASCE_GH")>
    Public Property AutoCalc_ASCE_GH() As Boolean
        Get
            Return Me.prop_AutoCalc_ASCE_GH
        End Get
        Set
            Me.prop_AutoCalc_ASCE_GH = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("{0 = B, 1 = C,  2 = D}"), DisplayName("ASCE_ExposureCat")>
    Public Property ASCE_ExposureCat() As Integer
        Get
            Return Me.prop_ASCE_ExposureCat
        End Get
        Set
            Me.prop_ASCE_ExposureCat = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("{0 = ASCE 7-88, 1 = ASCE 7-93, 2= ASCE 7-95, 3 = ASCE 7-98, 4 = ASCE 7-02, 5 = Cook Co., IL, 6 = WIS 53, 7 = Chicago}"), DisplayName("ASCE_Year")>
    Public Property ASCE_Year() As Integer
        Get
            Return Me.prop_ASCE_Year
        End Get
        Set
            Me.prop_ASCE_Year = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("ASCEGh")>
    Public Property ASCEGh() As Double
        Get
            Return Me.prop_ASCEGh
        End Get
        Set
            Me.prop_ASCEGh = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("ASCEI")>
    Public Property ASCEI() As Double
        Get
            Return Me.prop_ASCEI
        End Get
        Set
            Me.prop_ASCEI = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("{0 = Every Section, 1 = between guys, 2 = user specify (WindCalcPoints)}"), DisplayName("CalcWindAt")>
    Public Property CalcWindAt() As Integer
        Get
            Return Me.prop_CalcWindAt
        End Get
        Set
            Me.prop_CalcWindAt = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("WindCalcPoints")>
    Public Property WindCalcPoints() As Double
        Get
            Return Me.prop_WindCalcPoints
        End Get
        Set
            Me.prop_WindCalcPoints = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("{0 = B, 1 = C, 2 = D}"), DisplayName("WindExposure")>
    Public Property WindExposure() As Integer
        Get
            Return Me.prop_WindExposure
        End Get
        Set
            Me.prop_WindExposure = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("Structure Class - TIA-222-G Only {0 = I, 1 = II, 2 = III}"), DisplayName("StructureCategory")>
    Public Property StructureCategory() As Integer
        Get
            Return Me.prop_StructureCategory
        End Get
        Set
            Me.prop_StructureCategory = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("TIA-222-H Only {0 = I, 1 = II,  2 = III,  3 = IV}"), DisplayName("RiskCategory")>
    Public Property RiskCategory() As Integer
        Get
            Return Me.prop_RiskCategory
        End Get
        Set
            Me.prop_RiskCategory = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("{0 = 1, 1 = 2, 2 = 3, 3 = 4, 4 = 5/Rigorous Procedure}"), DisplayName("TopoCategory")>
    Public Property TopoCategory() As Integer
        Get
            Return Me.prop_TopoCategory
        End Get
        Set
            Me.prop_TopoCategory = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("{0 = Continuous Ridge, 1 = Flat Topped Ridge, 2 = Hill, 3 = Flat Topped Hill, 4 = Continuous Escarpment}"), DisplayName("RSMTopographicFeature")>
    Public Property RSMTopographicFeature() As Integer
        Get
            Return Me.prop_RSMTopographicFeature
        End Get
        Set
            Me.prop_RSMTopographicFeature = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("RSM_L")>
    Public Property RSM_L() As Double
        Get
            Return Me.prop_RSM_L
        End Get
        Set
            Me.prop_RSM_L = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("RSM_X")>
    Public Property RSM_X() As Double
        Get
            Return Me.prop_RSM_X
        End Get
        Set
            Me.prop_RSM_X = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("CrestHeight")>
    Public Property CrestHeight() As Double
        Get
            Return Me.prop_CrestHeight
        End Get
        Set
            Me.prop_CrestHeight = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("TIA_222_H_TopoFeatureDownwind")>
    Public Property TIA_222_H_TopoFeatureDownwind() As Boolean
        Get
            Return Me.prop_TIA_222_H_TopoFeatureDownwind
        End Get
        Set
            Me.prop_TIA_222_H_TopoFeatureDownwind = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("BaseElevAboveSeaLevel")>
    Public Property BaseElevAboveSeaLevel() As Double
        Get
            Return Me.prop_BaseElevAboveSeaLevel
        End Get
        Set
            Me.prop_BaseElevAboveSeaLevel = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("ConsiderRooftopSpeedUp")>
    Public Property ConsiderRooftopSpeedUp() As Boolean
        Get
            Return Me.prop_ConsiderRooftopSpeedUp
        End Get
        Set
            Me.prop_ConsiderRooftopSpeedUp = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("RooftopWS")>
    Public Property RooftopWS() As Double
        Get
            Return Me.prop_RooftopWS
        End Get
        Set
            Me.prop_RooftopWS = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("RooftopHS")>
    Public Property RooftopHS() As Double
        Get
            Return Me.prop_RooftopHS
        End Get
        Set
            Me.prop_RooftopHS = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("RooftopParapetHt")>
    Public Property RooftopParapetHt() As Double
        Get
            Return Me.prop_RooftopParapetHt
        End Get
        Set
            Me.prop_RooftopParapetHt = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description(""), DisplayName("RooftopXB")>
    Public Property RooftopXB() As Double
        Get
            Return Me.prop_RooftopXB
        End Get
        Set
            Me.prop_RooftopXB = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("EIA-222-C and earlier {0 = A, 1 = B, 2 = C}"), DisplayName("WindZone")>
    Public Property WindZone() As Integer
        Get
            Return Me.prop_WindZone
        End Get
        Set
            Me.prop_WindZone = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("EIA-222-C and earlier"), DisplayName("EIACWindMult")>
    Public Property EIACWindMult() As Double
        Get
            Return Me.prop_EIACWindMult
        End Get
        Set
            Me.prop_EIACWindMult = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("EIA-222-C and earlier"), DisplayName("EIACWindMultIce")>
    Public Property EIACWindMultIce() As Double
        Get
            Return Me.prop_EIACWindMultIce
        End Get
        Set
            Me.prop_EIACWindMultIce = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("EIA-222-C and earlier - Set Cable Drag Factor to 1.0"), DisplayName("EIACIgnoreCableDrag")>
    Public Property EIACIgnoreCableDrag() As Boolean
        Get
            Return Me.prop_EIACIgnoreCableDrag
        End Get
        Set
            Me.prop_EIACIgnoreCableDrag = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("CSA S37-01 only"), DisplayName("CSA_S37_RefVelPress")>
    Public Property CSA_S37_RefVelPress() As Double
        Get
            Return Me.prop_CSA_S37_RefVelPress
        End Get
        Set
            Me.prop_CSA_S37_RefVelPress = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("CSA S37-01 only {0 = I, 1 = II, 2 = III}"), DisplayName("CSA_S37_ReliabilityClass")>
    Public Property CSA_S37_ReliabilityClass() As Integer
        Get
            Return Me.prop_CSA_S37_ReliabilityClass
        End Get
        Set
            Me.prop_CSA_S37_ReliabilityClass = Value
        End Set
    End Property
    <Category("TNX Code Wind"), Description("CSA S37-01 only"), DisplayName("CSA_S37_ServiceabilityFactor")>
    Public Property CSA_S37_ServiceabilityFactor() As Double
        Get
            Return Me.prop_CSA_S37_ServiceabilityFactor
        End Get
        Set
            Me.prop_CSA_S37_ServiceabilityFactor = Value
        End Set
    End Property

End Class
Partial Public Class tnxMisclCode
    Private prop_GroutFc As Double
    Private prop_TowerBoltGrade As String
    Private prop_TowerBoltMinEdgeDist As Double

    <Category("TNX Code Miscellaneous"), Description(""), DisplayName("GroutFc")>
    Public Property GroutFc() As Double
        Get
            Return Me.prop_GroutFc
        End Get
        Set
            Me.prop_GroutFc = Value
        End Set
    End Property
    <Category("TNX Code Miscellaneous"), Description("Default bolt grade"), DisplayName("TowerBoltGrade")>
    Public Property TowerBoltGrade() As String
        Get
            Return Me.prop_TowerBoltGrade
        End Get
        Set
            Me.prop_TowerBoltGrade = Value
        End Set
    End Property
    <Category("TNX Code Miscellaneous"), Description("Not in UI"), DisplayName("TowerBoltMinEdgeDist")>
    Public Property TowerBoltMinEdgeDist() As Double
        Get
            Return Me.prop_TowerBoltMinEdgeDist
        End Get
        Set
            Me.prop_TowerBoltMinEdgeDist = Value
        End Set
    End Property
End Class

Partial Public Class tnxSeismic
    Private prop_UseASCE7_10_Seismic_Lcomb As Boolean
    Private prop_SeismicSiteClass As Integer
    Private prop_SeismicSs As Double
    Private prop_SeismicS1 As Double

    <Category("TNX Code seismic"), Description(""), DisplayName("UseASCE7_10_Seismic_Lcomb")>
    Public Property UseASCE7_10_Seismic_Lcomb() As Boolean
        Get
            Return Me.prop_UseASCE7_10_Seismic_Lcomb
        End Get
        Set
            Me.prop_UseASCE7_10_Seismic_Lcomb = Value
        End Set
    End Property
    <Category("TNX Code seismic"), Description("not in UI {0 = A, 1 = B, 2 = C, 3 = D, 4 = E} "), DisplayName("SeismicSiteClass")>
    Public Property SeismicSiteClass() As Integer
        Get
            Return Me.prop_SeismicSiteClass
        End Get
        Set
            Me.prop_SeismicSiteClass = Value
        End Set
    End Property
    <Category("TNX Code seismic"), Description("not in UI"), DisplayName("SeismicSs")>
    Public Property SeismicSs() As Double
        Get
            Return Me.prop_SeismicSs
        End Get
        Set
            Me.prop_SeismicSs = Value
        End Set
    End Property
    <Category("TNX Code seismic"), Description("not in UI"), DisplayName("SeismicS1")>
    Public Property SeismicS1() As Double
        Get
            Return Me.prop_SeismicS1
        End Get
        Set
            Me.prop_SeismicS1 = Value
        End Set
    End Property
End Class


#End Region

#Region "Options"
Partial Public Class tnxOptions
    Private prop_UseClearSpans As Boolean
    Private prop_UseClearSpansKlr As Boolean
    Private prop_UseFeedlineAsCylinder As Boolean
    Private prop_UseLegLoads As Boolean
    Private prop_SRTakeCompression As Boolean
    Private prop_AllLegPanelsSame As Boolean
    Private prop_UseCombinedBoltCapacity As Boolean
    Private prop_SecHorzBracesLeg As Boolean
    Private prop_SortByComponent As Boolean
    Private prop_SRCutEnds As Boolean
    Private prop_SRConcentric As Boolean
    Private prop_CalcBlockShear As Boolean
    Private prop_Use4SidedDiamondBracing As Boolean
    Private prop_TriangulateInnerBracing As Boolean
    Private prop_PrintCarrierNotes As Boolean
    Private prop_AddIBCWindCase As Boolean
    Private prop_LegBoltsAtTop As Boolean
    Private prop_UseTIA222Exemptions_MinBracingResistance As Boolean
    Private prop_UseTIA222Exemptions_TensionSplice As Boolean
    Private prop_IgnoreKLryFor60DegAngleLegs As Boolean
    Private prop_UseFeedlineTorque As Boolean
    Private prop_UsePinnedElements As Boolean
    Private prop_UseRigidIndex As Boolean
    Private prop_UseTrueCable As Boolean
    Private prop_UseASCELy As Boolean
    Private prop_CalcBracingForces As Boolean
    Private prop_IgnoreBracingFEA As Boolean
    Private prop_BypassStabilityChecks As Boolean
    Private prop_UseWindProjection As Boolean
    Private prop_UseDishCoeff As Boolean
    Private prop_AutoCalcTorqArmArea As Boolean
    Private prop_foundationStiffness As New tnxFoundaionStiffness()
    Private prop_defaultGirtOffsets As New tnxDefaultGirtOffsets()
    Private prop_cantileverPoles As New tnxCantileverPoles()
    Private prop_windDirections As New tnxWindDirections()
    Private prop_misclOptions As New tnxMisclOptions()

    <Category("TNX Options"), Description(""), DisplayName("UseClearSpans")>
    Public Property UseClearSpans() As Boolean
        Get
            Return Me.prop_UseClearSpans
        End Get
        Set
            Me.prop_UseClearSpans = Value
        End Set
    End Property
    <Category("TNX Options"), Description(""), DisplayName("UseClearSpansKlr")>
    Public Property UseClearSpansKlr() As Boolean
        Get
            Return Me.prop_UseClearSpansKlr
        End Get
        Set
            Me.prop_UseClearSpansKlr = Value
        End Set
    End Property
    <Category("TNX Options"), Description("treat feedline bundles as cylindrical"), DisplayName("UseFeedlineAsCylinder")>
    Public Property UseFeedlineAsCylinder() As Boolean
        Get
            Return Me.prop_UseFeedlineAsCylinder
        End Get
        Set
            Me.prop_UseFeedlineAsCylinder = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Distribute Leg Loads As Uniform"), DisplayName("UseLegLoads")>
    Public Property UseLegLoads() As Boolean
        Get
            Return Me.prop_UseLegLoads
        End Get
        Set
            Me.prop_UseLegLoads = Value
        End Set
    End Property
    <Category("TNX Options"), Description("SR Sleeve Bolts Resist Compression"), DisplayName("SRTakeCompression")>
    Public Property SRTakeCompression() As Boolean
        Get
            Return Me.prop_SRTakeCompression
        End Get
        Set
            Me.prop_SRTakeCompression = Value
        End Set
    End Property
    <Category("TNX Options"), Description("All Leg Panels Have Same Allowable"), DisplayName("AllLegPanelsSame")>
    Public Property AllLegPanelsSame() As Boolean
        Get
            Return Me.prop_AllLegPanelsSame
        End Get
        Set
            Me.prop_AllLegPanelsSame = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Include Bolts In Member Capacity"), DisplayName("UseCombinedBoltCapacity")>
    Public Property UseCombinedBoltCapacity() As Boolean
        Get
            Return Me.prop_UseCombinedBoltCapacity
        End Get
        Set
            Me.prop_UseCombinedBoltCapacity = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Secondary Horizontal Braces Leg"), DisplayName("SecHorzBracesLeg")>
    Public Property SecHorzBracesLeg() As Boolean
        Get
            Return Me.prop_SecHorzBracesLeg
        End Get
        Set
            Me.prop_SecHorzBracesLeg = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Sort Capacity Reports By Component"), DisplayName("SortByComponent")>
    Public Property SortByComponent() As Boolean
        Get
            Return Me.prop_SortByComponent
        End Get
        Set
            Me.prop_SortByComponent = Value
        End Set
    End Property
    <Category("TNX Options"), Description("SR Members Have Cut Ends"), DisplayName("SRCutEnds")>
    Public Property SRCutEnds() As Boolean
        Get
            Return Me.prop_SRCutEnds
        End Get
        Set
            Me.prop_SRCutEnds = Value
        End Set
    End Property
    <Category("TNX Options"), Description("SR Members Are Concentric"), DisplayName("SRConcentric")>
    Public Property SRConcentric() As Boolean
        Get
            Return Me.prop_SRConcentric
        End Get
        Set
            Me.prop_SRConcentric = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Include Angle Block Shear Check"), DisplayName("CalcBlockShear")>
    Public Property CalcBlockShear() As Boolean
        Get
            Return Me.prop_CalcBlockShear
        End Get
        Set
            Me.prop_CalcBlockShear = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Use Diamond Inner Bracing"), DisplayName("Use4SidedDiamondBracing")>
    Public Property Use4SidedDiamondBracing() As Boolean
        Get
            Return Me.prop_Use4SidedDiamondBracing
        End Get
        Set
            Me.prop_Use4SidedDiamondBracing = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Triangulate Diamond Inner Bracing"), DisplayName("TriangulateInnerBracing")>
    Public Property TriangulateInnerBracing() As Boolean
        Get
            Return Me.prop_TriangulateInnerBracing
        End Get
        Set
            Me.prop_TriangulateInnerBracing = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Print Carrier/Notes"), DisplayName("PrintCarrierNotes")>
    Public Property PrintCarrierNotes() As Boolean
        Get
            Return Me.prop_PrintCarrierNotes
        End Get
        Set
            Me.prop_PrintCarrierNotes = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Add IBC .6D+W Combination"), DisplayName("AddIBCWindCase")>
    Public Property AddIBCWindCase() As Boolean
        Get
            Return Me.prop_AddIBCWindCase
        End Get
        Set
            Me.prop_AddIBCWindCase = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Leg Bolts Are At Top Of Section"), DisplayName("LegBoltsAtTop")>
    Public Property LegBoltsAtTop() As Boolean
        Get
            Return Me.prop_LegBoltsAtTop
        End Get
        Set
            Me.prop_LegBoltsAtTop = Value
        End Set
    End Property
    <Category("TNX Options"), Description(""), DisplayName("UseTIA222Exemptions_MinBracingResistance")>
    Public Property UseTIA222Exemptions_MinBracingResistance() As Boolean
        Get
            Return Me.prop_UseTIA222Exemptions_MinBracingResistance
        End Get
        Set
            Me.prop_UseTIA222Exemptions_MinBracingResistance = Value
        End Set
    End Property
    <Category("TNX Options"), Description(""), DisplayName("UseTIA222Exemptions_TensionSplice")>
    Public Property UseTIA222Exemptions_TensionSplice() As Boolean
        Get
            Return Me.prop_UseTIA222Exemptions_TensionSplice
        End Get
        Set
            Me.prop_UseTIA222Exemptions_TensionSplice = Value
        End Set
    End Property
    <Category("TNX Options"), Description(""), DisplayName("IgnoreKLryFor60DegAngleLegs")>
    Public Property IgnoreKLryFor60DegAngleLegs() As Boolean
        Get
            Return Me.prop_IgnoreKLryFor60DegAngleLegs
        End Get
        Set
            Me.prop_IgnoreKLryFor60DegAngleLegs = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Consider Feed Line Torque"), DisplayName("UseFeedlineTorque")>
    Public Property UseFeedlineTorque() As Boolean
        Get
            Return Me.prop_UseFeedlineTorque
        End Get
        Set
            Me.prop_UseFeedlineTorque = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Assume Legs Pinned"), DisplayName("UsePinnedElements")>
    Public Property UsePinnedElements() As Boolean
        Get
            Return Me.prop_UsePinnedElements
        End Get
        Set
            Me.prop_UsePinnedElements = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Assume Rigid Index Plate"), DisplayName("UseRigidIndex")>
    Public Property UseRigidIndex() As Boolean
        Get
            Return Me.prop_UseRigidIndex
        End Get
        Set
            Me.prop_UseRigidIndex = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Retension Guys To Initial Tension"), DisplayName("UseTrueCable")>
    Public Property UseTrueCable() As Boolean
        Get
            Return Me.prop_UseTrueCable
        End Get
        Set
            Me.prop_UseTrueCable = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Use ASCE 10 X-Brace Ly Rules"), DisplayName("UseASCELy")>
    Public Property UseASCELy() As Boolean
        Get
            Return Me.prop_UseASCELy
        End Get
        Set
            Me.prop_UseASCELy = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Calculate Forces in Supporing Bracing Members"), DisplayName("CalcBracingForces")>
    Public Property CalcBracingForces() As Boolean
        Get
            Return Me.prop_CalcBracingForces
        End Get
        Set
            Me.prop_CalcBracingForces = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Ignore Redundant Bracing in FEA"), DisplayName("IgnoreBracingFEA")>
    Public Property IgnoreBracingFEA() As Boolean
        Get
            Return Me.prop_IgnoreBracingFEA
        End Get
        Set
            Me.prop_IgnoreBracingFEA = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Bypass Mast Stability Checks"), DisplayName("BypassStabilityChecks")>
    Public Property BypassStabilityChecks() As Boolean
        Get
            Return Me.prop_BypassStabilityChecks
        End Get
        Set
            Me.prop_BypassStabilityChecks = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Project Wind Area Of Appurtenances"), DisplayName("UseWindProjection")>
    Public Property UseWindProjection() As Boolean
        Get
            Return Me.prop_UseWindProjection
        End Get
        Set
            Me.prop_UseWindProjection = Value
        End Set
    End Property
    <Category("TNX Options"), Description("Use Azimuth Dish Coefficients"), DisplayName("UseDishCoeff")>
    Public Property UseDishCoeff() As Boolean
        Get
            Return Me.prop_UseDishCoeff
        End Get
        Set
            Me.prop_UseDishCoeff = Value
        End Set
    End Property
    <Category("TNX Options"), Description("AutoCalc Torque Arm Area"), DisplayName("AutoCalcTorqArmArea")>
    Public Property AutoCalcTorqArmArea() As Boolean
        Get
            Return Me.prop_AutoCalcTorqArmArea
        End Get
        Set
            Me.prop_AutoCalcTorqArmArea = Value
        End Set
    End Property

    <Category("TNX Options"), Description(""), DisplayName("Foundation Stiffness Options")>
    Public Property foundationStiffness() As tnxFoundaionStiffness
        Get
            Return Me.prop_foundationStiffness
        End Get
        Set
            Me.prop_foundationStiffness = Value
        End Set
    End Property

    <Category("TNX Options"), Description(""), DisplayName("Default Girt Offsets Options")>
    Public Property defaultGirtOffsets() As tnxDefaultGirtOffsets
        Get
            Return Me.prop_defaultGirtOffsets
        End Get
        Set
            Me.prop_defaultGirtOffsets = Value
        End Set
    End Property

    <Category("TNX Options"), Description(""), DisplayName("Cantilever Pole Options")>
    Public Property cantileverPoles() As tnxCantileverPoles
        Get
            Return Me.prop_cantileverPoles
        End Get
        Set
            Me.prop_cantileverPoles = Value
        End Set
    End Property

    <Category("TNX Options"), Description(""), DisplayName("Wind Direction Options")>
    Public Property windDirections() As tnxWindDirections
        Get
            Return Me.prop_windDirections
        End Get
        Set
            Me.prop_windDirections = Value
        End Set
    End Property

    <Category("TNX Options"), Description(""), DisplayName("Miscellaneous Options")>
    Public Property misclOptions() As tnxMisclOptions
        Get
            Return Me.prop_misclOptions
        End Get
        Set
            Me.prop_misclOptions = Value
        End Set
    End Property

End Class

Partial Public Class tnxFoundaionStiffness
    Private prop_MastVert As Double
    Private prop_MastHorz As Double
    Private prop_GuyVert As Double
    Private prop_GuyHorz As Double

    <Category("TNX Foundation Stiffness Options"), Description("foundation stiffness"), DisplayName("MastVert")>
    Public Property MastVert() As Double
        Get
            Return Me.prop_MastVert
        End Get
        Set
            Me.prop_MastVert = Value
        End Set
    End Property
    <Category("TNX Foundation Stiffness Options"), Description("foundation stiffness"), DisplayName("MastHorz")>
    Public Property MastHorz() As Double
        Get
            Return Me.prop_MastHorz
        End Get
        Set
            Me.prop_MastHorz = Value
        End Set
    End Property
    <Category("TNX Foundation Stiffness Options"), Description("foundation stiffness"), DisplayName("GuyVert")>
    Public Property GuyVert() As Double
        Get
            Return Me.prop_GuyVert
        End Get
        Set
            Me.prop_GuyVert = Value
        End Set
    End Property
    <Category("TNX Foundation Stiffness Options"), Description("foundation stiffness"), DisplayName("GuyHorz")>
    Public Property GuyHorz() As Double
        Get
            Return Me.prop_GuyHorz
        End Get
        Set
            Me.prop_GuyHorz = Value
        End Set
    End Property

End Class

Partial Public Class tnxDefaultGirtOffsets
    Private prop_GirtOffset As Double
    Private prop_GirtOffsetLatticedPole As Double
    Private prop_OffsetBotGirt As Boolean

    <Category("TNX Default Girt Offset Options"), Description(""), DisplayName("GirtOffset")>
    Public Property GirtOffset() As Double
        Get
            Return Me.prop_GirtOffset
        End Get
        Set
            Me.prop_GirtOffset = Value
        End Set
    End Property
    <Category("TNX Default Girt Offset Options"), Description(""), DisplayName("GirtOffsetLatticedPole")>
    Public Property GirtOffsetLatticedPole() As Double
        Get
            Return Me.prop_GirtOffsetLatticedPole
        End Get
        Set
            Me.prop_GirtOffsetLatticedPole = Value
        End Set
    End Property
    <Category("TNX Default Girt Offset Options"), Description("offset girt at foundation"), DisplayName("OffsetBotGirt")>
    Public Property OffsetBotGirt() As Boolean
        Get
            Return Me.prop_OffsetBotGirt
        End Get
        Set
            Me.prop_OffsetBotGirt = Value
        End Set
    End Property
End Class

Partial Public Class tnxCantileverPoles
    Private prop_CheckVonMises As Boolean
    Private prop_SocketTopMount As Boolean
    Private prop_PrintMonopoleAtIncrements As Boolean
    Private prop_UseSubCriticalFlow As Boolean
    Private prop_AssumePoleWithNoAttachments As Boolean
    Private prop_AssumePoleWithShroud As Boolean
    Private prop_PoleCornerRadiusKnown As Boolean
    Private prop_CantKFactor As Double

    <Category("TNX Cantilever Pole Options"), Description(")Include Shear-Torsion Interaction"), DisplayName("CheckVonMises")>
    Public Property CheckVonMises() As Boolean
        Get
            Return Me.prop_CheckVonMises
        End Get
        Set
            Me.prop_CheckVonMises = Value
        End Set
    End Property
    <Category("TNX Cantilever Pole Options"), Description("Use Top Mounted Socket"), DisplayName("SocketTopMount")>
    Public Property SocketTopMount() As Boolean
        Get
            Return Me.prop_SocketTopMount
        End Get
        Set
            Me.prop_SocketTopMount = Value
        End Set
    End Property
    <Category("TNX Cantilever Pole Options"), Description("Print Pole Stresses at Increments"), DisplayName("PrintMonopoleAtIncrements")>
    Public Property PrintMonopoleAtIncrements() As Boolean
        Get
            Return Me.prop_PrintMonopoleAtIncrements
        End Get
        Set
            Me.prop_PrintMonopoleAtIncrements = Value
        End Set
    End Property
    <Category("TNX Cantilever Pole Options"), Description("Always Yse Sub-Critical Flow"), DisplayName("UseSubCriticalFlow")>
    Public Property UseSubCriticalFlow() As Boolean
        Get
            Return Me.prop_UseSubCriticalFlow
        End Get
        Set
            Me.prop_UseSubCriticalFlow = Value
        End Set
    End Property
    <Category("TNX Cantilever Pole Options"), Description("Pole Without Linear Attachments"), DisplayName("AssumePoleWithNoAttachments")>
    Public Property AssumePoleWithNoAttachments() As Boolean
        Get
            Return Me.prop_AssumePoleWithNoAttachments
        End Get
        Set
            Me.prop_AssumePoleWithNoAttachments = Value
        End Set
    End Property
    <Category("TNX Cantilever Pole Options"), Description("Pole With Shroud or No Appurtenances"), DisplayName("AssumePoleWithShroud")>
    Public Property AssumePoleWithShroud() As Boolean
        Get
            Return Me.prop_AssumePoleWithShroud
        End Get
        Set
            Me.prop_AssumePoleWithShroud = Value
        End Set
    End Property
    <Category("TNX Cantilever Pole Options"), Description("Outside and Inside Corner Radii Are Known"), DisplayName("PoleCornerRadiusKnown")>
    Public Property PoleCornerRadiusKnown() As Boolean
        Get
            Return Me.prop_PoleCornerRadiusKnown
        End Get
        Set
            Me.prop_PoleCornerRadiusKnown = Value
        End Set
    End Property
    <Category("TNX Cantilever Pole Options"), Description("Cantilevered Poles K Factor"), DisplayName("CantKFactor")>
    Public Property CantKFactor() As Double
        Get
            Return Me.prop_CantKFactor
        End Get
        Set
            Me.prop_CantKFactor = Value
        End Set
    End Property

End Class

Partial Public Class tnxWindDirections
    Private prop_WindDirOption As Integer
    Private prop_WindDir0_0 As Boolean
    Private prop_WindDir0_1 As Boolean
    Private prop_WindDir0_2 As Boolean
    Private prop_WindDir0_3 As Boolean
    Private prop_WindDir0_4 As Boolean
    Private prop_WindDir0_5 As Boolean
    Private prop_WindDir0_6 As Boolean
    Private prop_WindDir0_7 As Boolean
    Private prop_WindDir0_8 As Boolean
    Private prop_WindDir0_9 As Boolean
    Private prop_WindDir0_10 As Boolean
    Private prop_WindDir0_11 As Boolean
    Private prop_WindDir0_12 As Boolean
    Private prop_WindDir0_13 As Boolean
    Private prop_WindDir0_14 As Boolean
    Private prop_WindDir0_15 As Boolean
    Private prop_WindDir1_0 As Boolean
    Private prop_WindDir1_1 As Boolean
    Private prop_WindDir1_2 As Boolean
    Private prop_WindDir1_3 As Boolean
    Private prop_WindDir1_4 As Boolean
    Private prop_WindDir1_5 As Boolean
    Private prop_WindDir1_6 As Boolean
    Private prop_WindDir1_7 As Boolean
    Private prop_WindDir1_8 As Boolean
    Private prop_WindDir1_9 As Boolean
    Private prop_WindDir1_10 As Boolean
    Private prop_WindDir1_11 As Boolean
    Private prop_WindDir1_12 As Boolean
    Private prop_WindDir1_13 As Boolean
    Private prop_WindDir1_14 As Boolean
    Private prop_WindDir1_15 As Boolean
    Private prop_WindDir2_0 As Boolean
    Private prop_WindDir2_1 As Boolean
    Private prop_WindDir2_2 As Boolean
    Private prop_WindDir2_3 As Boolean
    Private prop_WindDir2_4 As Boolean
    Private prop_WindDir2_5 As Boolean
    Private prop_WindDir2_6 As Boolean
    Private prop_WindDir2_7 As Boolean
    Private prop_WindDir2_8 As Boolean
    Private prop_WindDir2_9 As Boolean
    Private prop_WindDir2_10 As Boolean
    Private prop_WindDir2_11 As Boolean
    Private prop_WindDir2_12 As Boolean
    Private prop_WindDir2_13 As Boolean
    Private prop_WindDir2_14 As Boolean
    Private prop_WindDir2_15 As Boolean
    Private prop_SuppressWindPatternLoading As Boolean

    <Category("TNX Wind Direction Options"), Description("Wind Directions - 0 = Basic 3, 1 = All, 2 = Custom"), DisplayName("WindDirOption")>
    Public Property WindDirOption() As Integer
        Get
            Return Me.prop_WindDirOption
        End Get
        Set
            Me.prop_WindDirOption = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind No Ice 0 deg"), DisplayName("WindDir0_0")>
    Public Property WindDir0_0() As Boolean
        Get
            Return Me.prop_WindDir0_0
        End Get
        Set
            Me.prop_WindDir0_0 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind No Ice 30 deg"), DisplayName("WindDir0_1")>
    Public Property WindDir0_1() As Boolean
        Get
            Return Me.prop_WindDir0_1
        End Get
        Set
            Me.prop_WindDir0_1 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind No Ice 45 deg"), DisplayName("WindDir0_2")>
    Public Property WindDir0_2() As Boolean
        Get
            Return Me.prop_WindDir0_2
        End Get
        Set
            Me.prop_WindDir0_2 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind No Ice 60 deg"), DisplayName("WindDir0_3")>
    Public Property WindDir0_3() As Boolean
        Get
            Return Me.prop_WindDir0_3
        End Get
        Set
            Me.prop_WindDir0_3 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind No Ice 90 deg"), DisplayName("WindDir0_4")>
    Public Property WindDir0_4() As Boolean
        Get
            Return Me.prop_WindDir0_4
        End Get
        Set
            Me.prop_WindDir0_4 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind No Ice 120 deg"), DisplayName("WindDir0_5")>
    Public Property WindDir0_5() As Boolean
        Get
            Return Me.prop_WindDir0_5
        End Get
        Set
            Me.prop_WindDir0_5 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind No Ice 135 deg"), DisplayName("WindDir0_6")>
    Public Property WindDir0_6() As Boolean
        Get
            Return Me.prop_WindDir0_6
        End Get
        Set
            Me.prop_WindDir0_6 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind No Ice 150 deg"), DisplayName("WindDir0_7")>
    Public Property WindDir0_7() As Boolean
        Get
            Return Me.prop_WindDir0_7
        End Get
        Set
            Me.prop_WindDir0_7 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind No Ice 180 deg"), DisplayName("WindDir0_8")>
    Public Property WindDir0_8() As Boolean
        Get
            Return Me.prop_WindDir0_8
        End Get
        Set
            Me.prop_WindDir0_8 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind No Ice 210 deg"), DisplayName("WindDir0_9")>
    Public Property WindDir0_9() As Boolean
        Get
            Return Me.prop_WindDir0_9
        End Get
        Set
            Me.prop_WindDir0_9 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind No Ice 225 deg"), DisplayName("WindDir0_10")>
    Public Property WindDir0_10() As Boolean
        Get
            Return Me.prop_WindDir0_10
        End Get
        Set
            Me.prop_WindDir0_10 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind No Ice 240 deg"), DisplayName("WindDir0_11")>
    Public Property WindDir0_11() As Boolean
        Get
            Return Me.prop_WindDir0_11
        End Get
        Set
            Me.prop_WindDir0_11 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind No Ice 270 deg"), DisplayName("WindDir0_12")>
    Public Property WindDir0_12() As Boolean
        Get
            Return Me.prop_WindDir0_12
        End Get
        Set
            Me.prop_WindDir0_12 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind No Ice 300 deg"), DisplayName("WindDir0_13")>
    Public Property WindDir0_13() As Boolean
        Get
            Return Me.prop_WindDir0_13
        End Get
        Set
            Me.prop_WindDir0_13 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind No Ice 315 deg"), DisplayName("WindDir0_14")>
    Public Property WindDir0_14() As Boolean
        Get
            Return Me.prop_WindDir0_14
        End Get
        Set
            Me.prop_WindDir0_14 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind No Ice 330 deg"), DisplayName("WindDir0_15")>
    Public Property WindDir0_15() As Boolean
        Get
            Return Me.prop_WindDir0_15
        End Get
        Set
            Me.prop_WindDir0_15 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Ice 0 deg"), DisplayName("WindDir1_0")>
    Public Property WindDir1_0() As Boolean
        Get
            Return Me.prop_WindDir1_0
        End Get
        Set
            Me.prop_WindDir1_0 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Ice 30 deg"), DisplayName("WindDir1_1")>
    Public Property WindDir1_1() As Boolean
        Get
            Return Me.prop_WindDir1_1
        End Get
        Set
            Me.prop_WindDir1_1 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Ice 45 deg"), DisplayName("WindDir1_2")>
    Public Property WindDir1_2() As Boolean
        Get
            Return Me.prop_WindDir1_2
        End Get
        Set
            Me.prop_WindDir1_2 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Ice 60 deg"), DisplayName("WindDir1_3")>
    Public Property WindDir1_3() As Boolean
        Get
            Return Me.prop_WindDir1_3
        End Get
        Set
            Me.prop_WindDir1_3 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Ice 90 deg"), DisplayName("WindDir1_4")>
    Public Property WindDir1_4() As Boolean
        Get
            Return Me.prop_WindDir1_4
        End Get
        Set
            Me.prop_WindDir1_4 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Ice 120 deg"), DisplayName("WindDir1_5")>
    Public Property WindDir1_5() As Boolean
        Get
            Return Me.prop_WindDir1_5
        End Get
        Set
            Me.prop_WindDir1_5 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Ice 135 deg"), DisplayName("WindDir1_6")>
    Public Property WindDir1_6() As Boolean
        Get
            Return Me.prop_WindDir1_6
        End Get
        Set
            Me.prop_WindDir1_6 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Ice 150 deg"), DisplayName("WindDir1_7")>
    Public Property WindDir1_7() As Boolean
        Get
            Return Me.prop_WindDir1_7
        End Get
        Set
            Me.prop_WindDir1_7 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Ice 180 deg"), DisplayName("WindDir1_8")>
    Public Property WindDir1_8() As Boolean
        Get
            Return Me.prop_WindDir1_8
        End Get
        Set
            Me.prop_WindDir1_8 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Ice 210 deg"), DisplayName("WindDir1_9")>
    Public Property WindDir1_9() As Boolean
        Get
            Return Me.prop_WindDir1_9
        End Get
        Set
            Me.prop_WindDir1_9 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Ice 225 deg"), DisplayName("WindDir1_10")>
    Public Property WindDir1_10() As Boolean
        Get
            Return Me.prop_WindDir1_10
        End Get
        Set
            Me.prop_WindDir1_10 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Ice 240 deg"), DisplayName("WindDir1_11")>
    Public Property WindDir1_11() As Boolean
        Get
            Return Me.prop_WindDir1_11
        End Get
        Set
            Me.prop_WindDir1_11 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Ice 270 deg"), DisplayName("WindDir1_12")>
    Public Property WindDir1_12() As Boolean
        Get
            Return Me.prop_WindDir1_12
        End Get
        Set
            Me.prop_WindDir1_12 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Ice 300 deg"), DisplayName("WindDir1_13")>
    Public Property WindDir1_13() As Boolean
        Get
            Return Me.prop_WindDir1_13
        End Get
        Set
            Me.prop_WindDir1_13 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Ice 315 deg"), DisplayName("WindDir1_14")>
    Public Property WindDir1_14() As Boolean
        Get
            Return Me.prop_WindDir1_14
        End Get
        Set
            Me.prop_WindDir1_14 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Ice 330 deg"), DisplayName("WindDir1_15")>
    Public Property WindDir1_15() As Boolean
        Get
            Return Me.prop_WindDir1_15
        End Get
        Set
            Me.prop_WindDir1_15 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Service 0 deg"), DisplayName("WindDir2_0")>
    Public Property WindDir2_0() As Boolean
        Get
            Return Me.prop_WindDir2_0
        End Get
        Set
            Me.prop_WindDir2_0 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Service 30 deg"), DisplayName("WindDir2_1")>
    Public Property WindDir2_1() As Boolean
        Get
            Return Me.prop_WindDir2_1
        End Get
        Set
            Me.prop_WindDir2_1 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Service 45 deg"), DisplayName("WindDir2_2")>
    Public Property WindDir2_2() As Boolean
        Get
            Return Me.prop_WindDir2_2
        End Get
        Set
            Me.prop_WindDir2_2 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Service 60 deg"), DisplayName("WindDir2_3")>
    Public Property WindDir2_3() As Boolean
        Get
            Return Me.prop_WindDir2_3
        End Get
        Set
            Me.prop_WindDir2_3 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Service 90 deg"), DisplayName("WindDir2_4")>
    Public Property WindDir2_4() As Boolean
        Get
            Return Me.prop_WindDir2_4
        End Get
        Set
            Me.prop_WindDir2_4 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Service 120 deg"), DisplayName("WindDir2_5")>
    Public Property WindDir2_5() As Boolean
        Get
            Return Me.prop_WindDir2_5
        End Get
        Set
            Me.prop_WindDir2_5 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Service 135 deg"), DisplayName("WindDir2_6")>
    Public Property WindDir2_6() As Boolean
        Get
            Return Me.prop_WindDir2_6
        End Get
        Set
            Me.prop_WindDir2_6 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Service 150 deg"), DisplayName("WindDir2_7")>
    Public Property WindDir2_7() As Boolean
        Get
            Return Me.prop_WindDir2_7
        End Get
        Set
            Me.prop_WindDir2_7 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Service 180 deg"), DisplayName("WindDir2_8")>
    Public Property WindDir2_8() As Boolean
        Get
            Return Me.prop_WindDir2_8
        End Get
        Set
            Me.prop_WindDir2_8 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Service 210 deg"), DisplayName("WindDir2_9")>
    Public Property WindDir2_9() As Boolean
        Get
            Return Me.prop_WindDir2_9
        End Get
        Set
            Me.prop_WindDir2_9 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Service 225 deg"), DisplayName("WindDir2_10")>
    Public Property WindDir2_10() As Boolean
        Get
            Return Me.prop_WindDir2_10
        End Get
        Set
            Me.prop_WindDir2_10 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Service 240 deg"), DisplayName("WindDir2_11")>
    Public Property WindDir2_11() As Boolean
        Get
            Return Me.prop_WindDir2_11
        End Get
        Set
            Me.prop_WindDir2_11 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Service 270 deg"), DisplayName("WindDir2_12")>
    Public Property WindDir2_12() As Boolean
        Get
            Return Me.prop_WindDir2_12
        End Get
        Set
            Me.prop_WindDir2_12 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Service 300 deg"), DisplayName("WindDir2_13")>
    Public Property WindDir2_13() As Boolean
        Get
            Return Me.prop_WindDir2_13
        End Get
        Set
            Me.prop_WindDir2_13 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Service 315 deg"), DisplayName("WindDir2_14")>
    Public Property WindDir2_14() As Boolean
        Get
            Return Me.prop_WindDir2_14
        End Get
        Set
            Me.prop_WindDir2_14 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Service 330 deg"), DisplayName("WindDir2_15")>
    Public Property WindDir2_15() As Boolean
        Get
            Return Me.prop_WindDir2_15
        End Get
        Set
            Me.prop_WindDir2_15 = Value
        End Set
    End Property
    <Category("TNX Wind Direction Options"), Description("Wind Directions - Suppress Generation of Pattrn Loading"), DisplayName("SuppressWindPatternLoading")>
    Public Property SuppressWindPatternLoading() As Boolean
        Get
            Return Me.prop_SuppressWindPatternLoading
        End Get
        Set
            Me.prop_SuppressWindPatternLoading = Value
        End Set
    End Property

End Class

Partial Public Class tnxMisclOptions
    Private prop_HogRodTakeup As Double
    Private prop_RadiusSampleDist As Double

    <Category("TNX Miscl Options"), Description("Tension Only Take-Up"), DisplayName("HogRodTakeup")>
    Public Property HogRodTakeup() As Double
        Get
            Return Me.prop_HogRodTakeup
        End Get
        Set
            Me.prop_HogRodTakeup = Value
        End Set
    End Property
    <Category("TNX Miscl Options"), Description("Sampling Distance"), DisplayName("RadiusSampleDist")>
    Public Property RadiusSampleDist() As Double
        Get
            Return Me.prop_RadiusSampleDist
        End Get
        Set
            Me.prop_RadiusSampleDist = Value
        End Set
    End Property

End Class


#End Region

#Region "Settings"

Partial Public Class tnxSettings
    'Other settings are not saved in ERI file
    Private prop_USUnits As New tnxUnits()
    'Private prop_SIunits As tnxSIUnits 
    Private prop_projectInfo As New tnxProjectInfo()
    Private prop_userInfo As New tnxUserInfo()

    <Category("TNX Setings"), Description(""), DisplayName("US Units")>
    Public Property USUnits() As tnxUnits
        Get
            Return Me.prop_USUnits
        End Get
        Set
            Me.prop_USUnits = Value
        End Set
    End Property
    <Category("TNX Setings"), Description(""), DisplayName("Project Info")>
    Public Property projectInfo() As tnxProjectInfo
        Get
            Return Me.prop_projectInfo
        End Get
        Set
            Me.prop_projectInfo = Value
        End Set
    End Property
    <Category("TNX Setings"), Description(""), DisplayName("User Info")>
    Public Property userInfo() As tnxUserInfo
        Get
            Return Me.prop_userInfo
        End Get
        Set
            Me.prop_userInfo = Value
        End Set
    End Property

End Class

Partial Public Class tnxSolutionSettings
    Private prop_SolutionUsePDelta As Boolean
    Private prop_SolutionMinStiffness As Double
    Private prop_SolutionMaxStiffness As Double
    Private prop_SolutionMaxCycles As Integer
    Private prop_SolutionPower As Double
    Private prop_SolutionTolerance As Double

    <Category("TNX Solution Options"), Description(""), DisplayName("SolutionUsePDelta")>
    Public Property SolutionUsePDelta() As Boolean
        Get
            Return Me.prop_SolutionUsePDelta
        End Get
        Set
            Me.prop_SolutionUsePDelta = Value
        End Set
    End Property
    <Category("TNX Solution Options"), Description(""), DisplayName("SolutionMinStiffness")>
    Public Property SolutionMinStiffness() As Double
        Get
            Return Me.prop_SolutionMinStiffness
        End Get
        Set
            Me.prop_SolutionMinStiffness = Value
        End Set
    End Property
    <Category("TNX Solution Options"), Description(""), DisplayName("SolutionMaxStiffness")>
    Public Property SolutionMaxStiffness() As Double
        Get
            Return Me.prop_SolutionMaxStiffness
        End Get
        Set
            Me.prop_SolutionMaxStiffness = Value
        End Set
    End Property
    <Category("TNX Solution Options"), Description(""), DisplayName("SolutionMaxCycles")>
    Public Property SolutionMaxCycles() As Integer
        Get
            Return Me.prop_SolutionMaxCycles
        End Get
        Set
            Me.prop_SolutionMaxCycles = Value
        End Set
    End Property
    <Category("TNX Solution Options"), Description(""), DisplayName("SolutionPower")>
    Public Property SolutionPower() As Double
        Get
            Return Me.prop_SolutionPower
        End Get
        Set
            Me.prop_SolutionPower = Value
        End Set
    End Property
    <Category("TNX Solution Options"), Description(""), DisplayName("SolutionTolerance")>
    Public Property SolutionTolerance() As Double
        Get
            Return Me.prop_SolutionTolerance
        End Get
        Set
            Me.prop_SolutionTolerance = Value
        End Set
    End Property

End Class

Partial Public Class tnxReportSettings
    Private prop_ReportInputCosts As Boolean
    Private prop_ReportInputGeometry As Boolean
    Private prop_ReportInputOptions As Boolean
    Private prop_ReportMaxForces As Boolean
    Private prop_ReportInputMap As Boolean
    Private prop_CostReportOutputType As String
    Private prop_CapacityReportOutputType As String
    Private prop_ReportPrintForceTotals As Boolean
    Private prop_ReportPrintForceDetails As Boolean
    Private prop_ReportPrintMastVectors As Boolean
    Private prop_ReportPrintAntPoleVectors As Boolean
    Private prop_ReportPrintDiscreteVectors As Boolean
    Private prop_ReportPrintDishVectors As Boolean
    Private prop_ReportPrintFeedTowerVectors As Boolean
    Private prop_ReportPrintUserLoadVectors As Boolean
    Private prop_ReportPrintPressures As Boolean
    Private prop_ReportPrintAppurtForces As Boolean
    Private prop_ReportPrintGuyForces As Boolean
    Private prop_ReportPrintGuyStressing As Boolean
    Private prop_ReportPrintDeflections As Boolean
    Private prop_ReportPrintReactions As Boolean
    Private prop_ReportPrintStressChecks As Boolean
    Private prop_ReportPrintBoltChecks As Boolean
    Private prop_ReportPrintInputGVerificationTables As Boolean
    Private prop_ReportPrintOutputGVerificationTables As Boolean

    <Category("TNX Report Settings"), Description(""), DisplayName("ReportInputCosts")>
    Public Property ReportInputCosts() As Boolean
        Get
            Return Me.prop_ReportInputCosts
        End Get
        Set
            Me.prop_ReportInputCosts = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportInputGeometry")>
    Public Property ReportInputGeometry() As Boolean
        Get
            Return Me.prop_ReportInputGeometry
        End Get
        Set
            Me.prop_ReportInputGeometry = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportInputOptions")>
    Public Property ReportInputOptions() As Boolean
        Get
            Return Me.prop_ReportInputOptions
        End Get
        Set
            Me.prop_ReportInputOptions = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportMaxForces")>
    Public Property ReportMaxForces() As Boolean
        Get
            Return Me.prop_ReportMaxForces
        End Get
        Set
            Me.prop_ReportMaxForces = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportInputMap")>
    Public Property ReportInputMap() As Boolean
        Get
            Return Me.prop_ReportInputMap
        End Get
        Set
            Me.prop_ReportInputMap = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description("{No Capacity Output, Capacity Summary, Capacity Details}"), DisplayName("CostReportOutputType")>
    Public Property CostReportOutputType() As String
        Get
            Return Me.prop_CostReportOutputType
        End Get
        Set
            Me.prop_CostReportOutputType = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description("{No Cost Output, Cost Summary, Cost Details}"), DisplayName("CapacityReportOutputType")>
    Public Property CapacityReportOutputType() As String
        Get
            Return Me.prop_CapacityReportOutputType
        End Get
        Set
            Me.prop_CapacityReportOutputType = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintForceTotals")>
    Public Property ReportPrintForceTotals() As Boolean
        Get
            Return Me.prop_ReportPrintForceTotals
        End Get
        Set
            Me.prop_ReportPrintForceTotals = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintForceDetails")>
    Public Property ReportPrintForceDetails() As Boolean
        Get
            Return Me.prop_ReportPrintForceDetails
        End Get
        Set
            Me.prop_ReportPrintForceDetails = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintMastVectors")>
    Public Property ReportPrintMastVectors() As Boolean
        Get
            Return Me.prop_ReportPrintMastVectors
        End Get
        Set
            Me.prop_ReportPrintMastVectors = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintAntPoleVectors")>
    Public Property ReportPrintAntPoleVectors() As Boolean
        Get
            Return Me.prop_ReportPrintAntPoleVectors
        End Get
        Set
            Me.prop_ReportPrintAntPoleVectors = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintDiscreteVectors")>
    Public Property ReportPrintDiscreteVectors() As Boolean
        Get
            Return Me.prop_ReportPrintDiscreteVectors
        End Get
        Set
            Me.prop_ReportPrintDiscreteVectors = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintDishVectors")>
    Public Property ReportPrintDishVectors() As Boolean
        Get
            Return Me.prop_ReportPrintDishVectors
        End Get
        Set
            Me.prop_ReportPrintDishVectors = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintFeedTowerVectors")>
    Public Property ReportPrintFeedTowerVectors() As Boolean
        Get
            Return Me.prop_ReportPrintFeedTowerVectors
        End Get
        Set
            Me.prop_ReportPrintFeedTowerVectors = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintUserLoadVectors")>
    Public Property ReportPrintUserLoadVectors() As Boolean
        Get
            Return Me.prop_ReportPrintUserLoadVectors
        End Get
        Set
            Me.prop_ReportPrintUserLoadVectors = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintPressures")>
    Public Property ReportPrintPressures() As Boolean
        Get
            Return Me.prop_ReportPrintPressures
        End Get
        Set
            Me.prop_ReportPrintPressures = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintAppurtForces")>
    Public Property ReportPrintAppurtForces() As Boolean
        Get
            Return Me.prop_ReportPrintAppurtForces
        End Get
        Set
            Me.prop_ReportPrintAppurtForces = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintGuyForces")>
    Public Property ReportPrintGuyForces() As Boolean
        Get
            Return Me.prop_ReportPrintGuyForces
        End Get
        Set
            Me.prop_ReportPrintGuyForces = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintGuyStressing")>
    Public Property ReportPrintGuyStressing() As Boolean
        Get
            Return Me.prop_ReportPrintGuyStressing
        End Get
        Set
            Me.prop_ReportPrintGuyStressing = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintDeflections")>
    Public Property ReportPrintDeflections() As Boolean
        Get
            Return Me.prop_ReportPrintDeflections
        End Get
        Set
            Me.prop_ReportPrintDeflections = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintReactions")>
    Public Property ReportPrintReactions() As Boolean
        Get
            Return Me.prop_ReportPrintReactions
        End Get
        Set
            Me.prop_ReportPrintReactions = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintStressChecks")>
    Public Property ReportPrintStressChecks() As Boolean
        Get
            Return Me.prop_ReportPrintStressChecks
        End Get
        Set
            Me.prop_ReportPrintStressChecks = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintBoltChecks")>
    Public Property ReportPrintBoltChecks() As Boolean
        Get
            Return Me.prop_ReportPrintBoltChecks
        End Get
        Set
            Me.prop_ReportPrintBoltChecks = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintInputGVerificationTables")>
    Public Property ReportPrintInputGVerificationTables() As Boolean
        Get
            Return Me.prop_ReportPrintInputGVerificationTables
        End Get
        Set
            Me.prop_ReportPrintInputGVerificationTables = Value
        End Set
    End Property
    <Category("TNX Report Settings"), Description(""), DisplayName("ReportPrintOutputGVerificationTables")>
    Public Property ReportPrintOutputGVerificationTables() As Boolean
        Get
            Return Me.prop_ReportPrintOutputGVerificationTables
        End Get
        Set
            Me.prop_ReportPrintOutputGVerificationTables = Value
        End Set
    End Property
End Class

Partial Public Class tnxMTOSettings
    Private prop_IncludeCapacityNote As Boolean
    Private prop_IncludeAppurtGraphics As Boolean
    Private prop_DisplayNotes As Boolean
    Private prop_DisplayReactions As Boolean
    Private prop_DisplaySchedule As Boolean
    Private prop_DisplayAppurtenanceTable As Boolean
    Private prop_DisplayMaterialStrengthTable As Boolean
    Private prop_Notes As New List(Of String)

    <Category("TNX MTO Settings"), Description(""), DisplayName("IncludeCapacityNote")>
    Public Property IncludeCapacityNote() As Boolean
        Get
            Return Me.prop_IncludeCapacityNote
        End Get
        Set
            Me.prop_IncludeCapacityNote = Value
        End Set
    End Property
    <Category("TNX MTO Settings"), Description(""), DisplayName("IncludeAppurtGraphics")>
    Public Property IncludeAppurtGraphics() As Boolean
        Get
            Return Me.prop_IncludeAppurtGraphics
        End Get
        Set
            Me.prop_IncludeAppurtGraphics = Value
        End Set
    End Property
    <Category("TNX MTO Settings"), Description(""), DisplayName("DisplayNotes")>
    Public Property DisplayNotes() As Boolean
        Get
            Return Me.prop_DisplayNotes
        End Get
        Set
            Me.prop_DisplayNotes = Value
        End Set
    End Property
    <Category("TNX MTO Settings"), Description(""), DisplayName("DisplayReactions")>
    Public Property DisplayReactions() As Boolean
        Get
            Return Me.prop_DisplayReactions
        End Get
        Set
            Me.prop_DisplayReactions = Value
        End Set
    End Property
    <Category("TNX MTO Settings"), Description(""), DisplayName("DisplaySchedule")>
    Public Property DisplaySchedule() As Boolean
        Get
            Return Me.prop_DisplaySchedule
        End Get
        Set
            Me.prop_DisplaySchedule = Value
        End Set
    End Property
    <Category("TNX MTO Settings"), Description(""), DisplayName("DisplayAppurtenanceTable")>
    Public Property DisplayAppurtenanceTable() As Boolean
        Get
            Return Me.prop_DisplayAppurtenanceTable
        End Get
        Set
            Me.prop_DisplayAppurtenanceTable = Value
        End Set
    End Property
    <Category("TNX MTO Settings"), Description(""), DisplayName("DisplayMaterialStrengthTable")>
    Public Property DisplayMaterialStrengthTable() As Boolean
        Get
            Return Me.prop_DisplayMaterialStrengthTable
        End Get
        Set
            Me.prop_DisplayMaterialStrengthTable = Value
        End Set
    End Property
    <Category("TNX MTO Settings"), Description(""), DisplayName("Notes")>
    Public Property Notes() As List(Of String)
        Get
            Return Me.prop_Notes
        End Get
        Set
            Me.prop_Notes = Value
        End Set
    End Property

End Class

Partial Public Class tnxProjectInfo

    Private prop_DesignStandardSeries As String
    Private prop_UnitsSystem As String
    Private prop_ClientName As String
    Private prop_ProjectName As String
    Private prop_ProjectNumber As String
    Private prop_CreatedBy As String
    Private prop_CreatedOn As String
    Private prop_LastUsedBy As String
    Private prop_LastUsedOn As String
    Private prop_VersionUsed As String

    <Category("TNX Project Info"), Description("TIA/EIA or CSA-S37"), DisplayName("DesignStandardSeries")>
    Public Property DesignStandardSeries() As String
        Get
            Return Me.prop_DesignStandardSeries
        End Get
        Set
            Me.prop_DesignStandardSeries = Value
        End Set
    End Property
    <Category("TNX Project Info"), Description("US or SI"), DisplayName("UnitsSystem")>
    Public Property UnitsSystem() As String
        Get
            Return Me.prop_UnitsSystem
        End Get
        Set
            Me.prop_UnitsSystem = Value
        End Set
    End Property
    <Category("TNX Project Info"), Description(""), DisplayName("ClientName")>
    Public Property ClientName() As String
        Get
            Return Me.prop_ClientName
        End Get
        Set
            Me.prop_ClientName = Value
        End Set
    End Property
    <Category("TNX Project Info"), Description(""), DisplayName("ProjectName")>
    Public Property ProjectName() As String
        Get
            Return Me.prop_ProjectName
        End Get
        Set
            Me.prop_ProjectName = Value
        End Set
    End Property
    <Category("TNX Project Info"), Description(""), DisplayName("ProjectNumber")>
    Public Property ProjectNumber() As String
        Get
            Return Me.prop_ProjectNumber
        End Get
        Set
            Me.prop_ProjectNumber = Value
        End Set
    End Property
    <Category("TNX Project Info"), Description(""), DisplayName("CreatedBy")>
    Public Property CreatedBy() As String
        Get
            Return Me.prop_CreatedBy
        End Get
        Set
            Me.prop_CreatedBy = Value
        End Set
    End Property
    <Category("TNX Project Info"), Description(""), DisplayName("CreatedOn")>
    Public Property CreatedOn() As String
        Get
            Return Me.prop_CreatedOn
        End Get
        Set
            Me.prop_CreatedOn = Value
        End Set
    End Property
    <Category("TNX Project Info"), Description(""), DisplayName("LastUsedBy")>
    Public Property LastUsedBy() As String
        Get
            Return Me.prop_LastUsedBy
        End Get
        Set
            Me.prop_LastUsedBy = Value
        End Set
    End Property
    <Category("TNX Project Info"), Description(""), DisplayName("LastUsedOn")>
    Public Property LastUsedOn() As String
        Get
            Return Me.prop_LastUsedOn
        End Get
        Set
            Me.prop_LastUsedOn = Value
        End Set
    End Property
    <Category("TNX Project Info"), Description(""), DisplayName("VersionUsed")>
    Public Property VersionUsed() As String
        Get
            Return Me.prop_VersionUsed
        End Get
        Set
            Me.prop_VersionUsed = Value
        End Set
    End Property


End Class

Partial Public Class tnxUserInfo

    Private prop_ViewerUserName As String
    Private prop_ViewerCompanyName As String
    Private prop_ViewerStreetAddress As String
    Private prop_ViewerCityState As String
    Private prop_ViewerPhone As String
    Private prop_ViewerFAX As String
    Private prop_ViewerLogo As String
    Private prop_ViewerCompanyBitmap As String

    <Category("TNX User Info"), Description(""), DisplayName("ViewerUserName")>
    Public Property ViewerUserName() As String
        Get
            Return Me.prop_ViewerUserName
        End Get
        Set
            Me.prop_ViewerUserName = Value
        End Set
    End Property
    <Category("TNX User Info"), Description(""), DisplayName("ViewerCompanyName")>
    Public Property ViewerCompanyName() As String
        Get
            Return Me.prop_ViewerCompanyName
        End Get
        Set
            Me.prop_ViewerCompanyName = Value
        End Set
    End Property
    <Category("TNX User Info"), Description(""), DisplayName("ViewerStreetAddress")>
    Public Property ViewerStreetAddress() As String
        Get
            Return Me.prop_ViewerStreetAddress
        End Get
        Set
            Me.prop_ViewerStreetAddress = Value
        End Set
    End Property
    <Category("TNX User Info"), Description(""), DisplayName("ViewerCityState")>
    Public Property ViewerCityState() As String
        Get
            Return Me.prop_ViewerCityState
        End Get
        Set
            Me.prop_ViewerCityState = Value
        End Set
    End Property
    <Category("TNX User Info"), Description(""), DisplayName("ViewerPhone")>
    Public Property ViewerPhone() As String
        Get
            Return Me.prop_ViewerPhone
        End Get
        Set
            Me.prop_ViewerPhone = Value
        End Set
    End Property
    <Category("TNX User Info"), Description(""), DisplayName("ViewerFAX")>
    Public Property ViewerFAX() As String
        Get
            Return Me.prop_ViewerFAX
        End Get
        Set
            Me.prop_ViewerFAX = Value
        End Set
    End Property
    <Category("TNX User Info"), Description(""), DisplayName("ViewerLogo")>
    Public Property ViewerLogo() As String
        Get
            Return Me.prop_ViewerLogo
        End Get
        Set
            Me.prop_ViewerLogo = Value
        End Set
    End Property
    <Category("TNX User Info"), Description(""), DisplayName("ViewerCompanyBitmap")>
    Public Property ViewerCompanyBitmap() As String
        Get
            Return Me.prop_ViewerCompanyBitmap
        End Get
        Set
            Me.prop_ViewerCompanyBitmap = Value
        End Set
    End Property

End Class

Partial Public Class tnxUnits

    Private prop_Length As New tnxLengthUnit()
    Private prop_Coordinate As New tnxCoordinateUnit()
    Private prop_Force As New tnxForceUnit()
    Private prop_Load As New tnxLoadUnit()
    Private prop_Moment As New tnxMomentUnit()
    Private prop_Properties As New tnxPropertiesUnit()
    Private prop_Pressure As New tnxPressureUnit()
    Private prop_Velocity As New tnxVelocityUnit()
    Private prop_Displacement As New tnxDisplacementUnit()
    Private prop_Mass As New tnxMassUnit()
    Private prop_Acceleration As New tnxAccelerationUnit()
    Private prop_Stress As New tnxStressUnit()
    Private prop_Density As New tnxDensityUnit()
    Private prop_UnitWt As New tnxUnitWTUnit()
    Private prop_Strength As New tnxStrengthUnit()
    Private prop_Modulus As New tnxModulusUnit()
    Private prop_Temperature As New tnxTempUnit()
    Private prop_Printer As New tnxPrinterUnit()
    Private prop_Rotation As New tnxRotationUnit()
    Private prop_Spacing As New tnxSpacingUnit()

    <Category("TNX Units"), Description(""), DisplayName("Length")>
    Public Property Length() As tnxLengthUnit
        Get
            Return Me.prop_Length
        End Get
        Set
            Me.prop_Length = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Coordinate")>
    Public Property Coordinate() As tnxCoordinateUnit
        Get
            Return Me.prop_Coordinate
        End Get
        Set
            Me.prop_Coordinate = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Force")>
    Public Property Force() As tnxForceUnit
        Get
            Return Me.prop_Force
        End Get
        Set
            Me.prop_Force = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Load")>
    Public Property Load() As tnxLoadUnit
        Get
            Return Me.prop_Load
        End Get
        Set
            Me.prop_Load = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Moment")>
    Public Property Moment() As tnxMomentUnit
        Get
            Return Me.prop_Moment
        End Get
        Set
            Me.prop_Moment = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Properties")>
    Public Property Properties() As tnxPropertiesUnit
        Get
            Return Me.prop_Properties
        End Get
        Set
            Me.prop_Properties = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Pressure")>
    Public Property Pressure() As tnxPressureUnit
        Get
            Return Me.prop_Pressure
        End Get
        Set
            Me.prop_Pressure = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Velocity")>
    Public Property Velocity() As tnxVelocityUnit
        Get
            Return Me.prop_Velocity
        End Get
        Set
            Me.prop_Velocity = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Displacement")>
    Public Property Displacement() As tnxDisplacementUnit
        Get
            Return Me.prop_Displacement
        End Get
        Set
            Me.prop_Displacement = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Mass")>
    Public Property Mass() As tnxMassUnit
        Get
            Return Me.prop_Mass
        End Get
        Set
            Me.prop_Mass = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Acceleration")>
    Public Property Acceleration() As tnxAccelerationUnit
        Get
            Return Me.prop_Acceleration
        End Get
        Set
            Me.prop_Acceleration = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Stress")>
    Public Property Stress() As tnxStressUnit
        Get
            Return Me.prop_Stress
        End Get
        Set
            Me.prop_Stress = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Density")>
    Public Property Density() As tnxDensityUnit
        Get
            Return Me.prop_Density
        End Get
        Set
            Me.prop_Density = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Unitwt")>
    Public Property UnitWt() As tnxUnitWTUnit
        Get
            Return Me.prop_UnitWt
        End Get
        Set
            Me.prop_UnitWt = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Strength")>
    Public Property Strength() As tnxStrengthUnit
        Get
            Return Me.prop_Strength
        End Get
        Set
            Me.prop_Strength = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Modulus")>
    Public Property Modulus() As tnxModulusUnit
        Get
            Return Me.prop_Modulus
        End Get
        Set
            Me.prop_Modulus = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Temperature")>
    Public Property Temperature() As tnxTempUnit
        Get
            Return Me.prop_Temperature
        End Get
        Set
            Me.prop_Temperature = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Printer")>
    Public Property Printer() As tnxPrinterUnit
        Get
            Return Me.prop_Printer
        End Get
        Set
            Me.prop_Printer = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Rotation")>
    Public Property Rotation() As tnxRotationUnit
        Get
            Return Me.prop_Rotation
        End Get
        Set
            Me.prop_Rotation = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Spacing")>
    Public Property Spacing() As tnxSpacingUnit
        Get
            Return Me.prop_Spacing
        End Get
        Set
            Me.prop_Spacing = Value
        End Set
    End Property

End Class

Partial Public Class tnxUnitProperty
    'Variables need to be public for inheritance
    Public prop_value As String
    Public prop_precision As Integer
    Public prop_multiplier As Double

    <Category("TNX Unit Property"), Description(""), DisplayName("Value")>
    Public Overridable Property value() As String
        Get
            Return Me.prop_value
        End Get
        Set
            Me.prop_value = Value
        End Set
    End Property
    <Category("TNX Unit Property"), Description(""), DisplayName("Precision")>
    Public Overridable Property precision() As Integer
        Get
            Return Me.prop_precision
        End Get
        Set
            If Value < 0 Then
                Me.prop_precision = 0
            ElseIf Value > 4 Then
                Me.prop_precision = 4
            Else
                Me.prop_precision = Value
            End If
        End Set
    End Property
    <Category("TNX Unit Property"), Description("Used to convert TNX file units to default EDS units during import."), DisplayName("Multiplier")>
    Public Overridable Property multiplier() As Double
        Get
            Return Me.prop_multiplier
        End Get
        Set
            Me.prop_multiplier = Value
        End Set
    End Property

    Public Sub New()
    End Sub
    Public Sub New(new_value As String)

        Me.value = new_value

    End Sub
    Public Overridable Function convertToEDSDefaultUnits(InputValue As Double) As Double

        If Me.prop_value = "" Then
            Throw New System.Exception("Property value not set")
        ElseIf Me.prop_multiplier = 0 Then
            Throw New System.Exception("Property multiplier not set")
        End If

        Return InputValue / Me.multiplier

    End Function

    Public Overridable Function convertToERIUnits(InputValue As Double) As Double

        If Me.prop_value = "" Then
            Throw New System.Exception("Property value not set")
        ElseIf Me.prop_multiplier = 0 Then
            Throw New System.Exception("Property multiplier not set")
        End If

        Return InputValue * Me.multiplier

    End Function

End Class

Partial Public Class tnxLengthUnit
    Inherits tnxUnitProperty

    Public Overrides Property value() As String
        Get
            Return Me.prop_value
        End Get
        Set
            Me.prop_value = Value

            If Me.prop_value = "ft" Then
                Me.prop_multiplier = 1
            ElseIf Me.prop_value = "in" Then
                Me.prop_multiplier = 12
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me.prop_value)
            End If
        End Set
    End Property

    Public Overridable Function convertAreaToEDSDefaultUnits(InputValue As Double) As Double

        If Me.prop_value = "" Then
            Throw New System.Exception("Property value not set")
        ElseIf Me.prop_multiplier = 0 Then
            Throw New System.Exception("Property multiplier not set")
        End If

        Return InputValue / (Me.multiplier * Me.multiplier)

    End Function

    Public Overridable Function convertAreaToERIUnits(InputValue As Double) As Double

        If Me.prop_value = "" Then
            Throw New System.Exception("Property value not set")
        ElseIf Me.prop_multiplier = 0 Then
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
            Return Me.prop_value
        End Get
        Set
            Me.prop_value = Value

            If Me.prop_value = "K" Then
                Me.prop_multiplier = 1
            ElseIf Me.prop_value = "lb" Then
                Me.prop_multiplier = 1000
            ElseIf Me.prop_value = "T" Then
                Me.prop_multiplier = 0.5
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me.prop_value)
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
            Return Me.prop_value
        End Get
        Set
            Me.prop_value = Value

            If Me.prop_value = "klf" Then
                Me.prop_multiplier = 1
            ElseIf Me.prop_value = "plf" Then
                Me.prop_multiplier = 1000
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me.prop_value)
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
            Return Me.prop_value
        End Get
        Set
            Me.prop_value = Value

            If Me.prop_value = "kip-ft" Then
                Me.prop_multiplier = 1
            ElseIf Me.prop_value = "lb-ft" Then
                Me.prop_multiplier = 1000
            ElseIf Me.prop_value = "lb-in" Then
                Me.prop_multiplier = 12000
            ElseIf Me.prop_value = "kip-in" Then
                Me.prop_multiplier = 12
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me.prop_value)
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
            Return Me.prop_value
        End Get
        Set
            Me.prop_value = Value

            If Me.prop_value = "ksf" Then
                Me.prop_multiplier = 1
            ElseIf Me.prop_value = "psf" Then
                Me.prop_multiplier = 1000
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me.prop_value)
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
            Return Me.prop_value
        End Get
        Set
            Me.prop_value = Value

            If Me.prop_value = "mph" Then
                Me.prop_multiplier = 1
            ElseIf Me.prop_value = "fps" Then
                Me.prop_multiplier = 5280 / 3600
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me.prop_value)
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
    Public Overrides Property precision() As Integer
        Get
            Return Me.prop_precision
        End Get
        Set
            If Value < 0 Then
                Me.prop_precision = 0
            ElseIf Value > 6 Then
                Me.prop_precision = 6
            Else
                Me.prop_precision = Value
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
            Return Me.prop_value
        End Get
        Set
            Me.prop_value = Value

            If Me.prop_value = "lb" Then
                Me.prop_multiplier = 1
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me.prop_value)
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
            Return Me.prop_value
        End Get
        Set
            Me.prop_value = Value

            If Me.prop_value = "G" Then
                Me.prop_multiplier = 1
            ElseIf Me.prop_value = "fpss" Then
                Me.prop_multiplier = 32.17405
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me.prop_value)
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
            Return Me.prop_value
        End Get
        Set
            Me.prop_value = Value

            If Me.prop_value = "ksi" Then
                Me.prop_multiplier = 1
            ElseIf Me.prop_value = "psi" Then
                Me.prop_multiplier = 1000
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me.prop_value)
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
            Return Me.prop_value
        End Get
        Set
            Me.prop_value = Value

            If Me.prop_value = "pcf" Then
                Me.prop_multiplier = 1
            ElseIf Me.prop_value = "pci" Then
                Me.prop_multiplier = 1728
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me.prop_value)
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
            Return Me.prop_value
        End Get
        Set
            Me.prop_value = Value

            If Me.prop_value = "plf" Then
                Me.prop_multiplier = 1
            ElseIf Me.prop_value = "klf" Then
                Me.prop_multiplier = 0.001
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me.prop_value)
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
            Return Me.prop_value
        End Get
        Set
            Me.prop_value = Value

            If Me.prop_value = "ksi" Then
                Me.prop_multiplier = 1
            ElseIf Me.prop_value = "psi" Then
                Me.prop_multiplier = 1000
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me.prop_value)
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
            Return Me.prop_value
        End Get
        Set
            Me.prop_value = Value

            If Me.prop_value = "ksi" Then
                Me.prop_multiplier = 1
            ElseIf Me.prop_value = "psi" Then
                Me.prop_multiplier = 1000
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me.prop_value)
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
            Return Me.prop_value
        End Get
        Set
            Me.prop_value = Value

            If Me.prop_value = "F" Then
                Me.prop_multiplier = 1
            ElseIf Me.prop_value = "C" Then
                'This conversion doesn't use a simple multiplier.
                'Override coversion function to get correct results
                Me.prop_multiplier = 1
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me.prop_value)
            End If
        End Set
    End Property

    Public Overrides Function convertToEDSDefaultUnits(InputValue As Double) As Double

        If Me.prop_value = "" Then
            Throw New System.Exception("Property value not set")
        End If

        If Me.prop_value = "C" Then
            Return InputValue * (9 / 5) + 32
        Else
            Return InputValue
        End If

    End Function

    Public Overrides Function convertToERIUnits(InputValue As Double) As Double

        If Me.prop_value = "" Then
            Throw New System.Exception("Property value not set")
        End If

        If Me.prop_value = "C" Then
            Return (InputValue - 32) * (5 / 9)
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
            Return Me.prop_value
        End Get
        Set
            Me.prop_value = Value

            If Me.prop_value = "in" Then
                Me.prop_multiplier = 1
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me.prop_value)
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

    Public Overrides Property precision() As Integer
        Get
            Return Me.prop_precision
        End Get
        Set
            If Value < 0 Then
                Me.prop_precision = 0
            ElseIf Value > 6 Then
                Me.prop_precision = 6
            Else
                Me.prop_precision = Value
            End If
        End Set
    End Property

    Public Overrides Property value() As String
        Get
            Return Me.prop_value
        End Get
        Set
            Me.prop_value = Value

            If Me.prop_value = "deg" Then
                Me.prop_multiplier = 1
            ElseIf Me.prop_value = "rad" Then
                Me.prop_multiplier = 3.14159 / 180
            Else
                Throw New System.Exception("Unrecognized Unit: " & Me.prop_value)
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

Partial Public Class tnxCCIReport
    Private prop_sReportProjectNumber As String
    Private prop_sReportJobType As String
    Private prop_sReportCarrierName As String
    Private prop_sReportCarrierSiteNumber As String
    Private prop_sReportCarrierSiteName As String
    Private prop_sReportSiteAddress As String
    Private prop_sReportLatitudeDegree As Double
    Private prop_sReportLatitudeMinute As Double
    Private prop_sReportLatitudeSecond As Double
    Private prop_sReportLongitudeDegree As Double
    Private prop_sReportLongitudeMinute As Double
    Private prop_sReportLongitudeSecond As Double
    Private prop_sReportLocalCodeRequirement As String
    Private prop_sReportSiteHistory As String
    Private prop_sReportTowerManufacturer As String
    Private prop_sReportMonthManufactured As String
    Private prop_sReportYearManufactured As Integer
    Private prop_sReportOriginalSpeed As Double
    Private prop_sReportOriginalCode As String
    Private prop_sReportTowerType As String
    Private prop_sReportEngrName As String
    Private prop_sReportEngrTitle As String
    Private prop_sReportHQPhoneNumber As String
    Private prop_sReportEmailAddress As String
    Private prop_sReportLogoPath As String
    Private prop_sReportCCiContactName As String
    Private prop_sReportCCiAddress1 As String
    Private prop_sReportCCiAddress2 As String
    Private prop_sReportCCiBUNumber As String
    Private prop_sReportCCiSiteName As String
    Private prop_sReportCCiJDENumber As String
    Private prop_sReportCCiWONumber As String
    Private prop_sReportCCiPONumber As String
    Private prop_sReportCCiAppNumber As String
    Private prop_sReportCCiRevNumber As String
    Private prop_sReportDocsProvided As New List(Of String)
    Private prop_sReportRecommendations As String
    Private prop_sReportAppurt1 As New List(Of String)
    Private prop_sReportAppurt2 As New List(Of String)
    Private prop_sReportAppurt3 As New List(Of String)
    Private prop_sReportAddlCapacity As New List(Of String)
    Private prop_sReportAssumption As New List(Of String)
    Private prop_sReportAppurt1Note1 As String
    Private prop_sReportAppurt1Note2 As String
    Private prop_sReportAppurt1Note3 As String
    Private prop_sReportAppurt1Note4 As String
    Private prop_sReportAppurt1Note5 As String
    Private prop_sReportAppurt1Note6 As String
    Private prop_sReportAppurt1Note7 As String
    Private prop_sReportAppurt2Note1 As String
    Private prop_sReportAppurt2Note2 As String
    Private prop_sReportAppurt2Note3 As String
    Private prop_sReportAppurt2Note4 As String
    Private prop_sReportAppurt2Note5 As String
    Private prop_sReportAppurt2Note6 As String
    Private prop_sReportAppurt2Note7 As String
    Private prop_sReportAddlCapacityNote1 As String
    Private prop_sReportAddlCapacityNote2 As String
    Private prop_sReportAddlCapacityNote3 As String
    Private prop_sReportAddlCapacityNote4 As String

    <Category("TNX CCI Report"), Description(""), DisplayName("sReportProjectNumber")>
    Public Property sReportProjectNumber() As String
        Get
            Return Me.prop_sReportProjectNumber
        End Get
        Set
            Me.prop_sReportProjectNumber = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportJobType")>
    Public Property sReportJobType() As String
        Get
            Return Me.prop_sReportJobType
        End Get
        Set
            Me.prop_sReportJobType = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportCarrierName")>
    Public Property sReportCarrierName() As String
        Get
            Return Me.prop_sReportCarrierName
        End Get
        Set
            Me.prop_sReportCarrierName = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportCarrierSiteNumber")>
    Public Property sReportCarrierSiteNumber() As String
        Get
            Return Me.prop_sReportCarrierSiteNumber
        End Get
        Set
            Me.prop_sReportCarrierSiteNumber = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportCarrierSiteName")>
    Public Property sReportCarrierSiteName() As String
        Get
            Return Me.prop_sReportCarrierSiteName
        End Get
        Set
            Me.prop_sReportCarrierSiteName = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportSiteAddress")>
    Public Property sReportSiteAddress() As String
        Get
            Return Me.prop_sReportSiteAddress
        End Get
        Set
            Me.prop_sReportSiteAddress = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportLatitudeDegree")>
    Public Property sReportLatitudeDegree() As Double
        Get
            Return Me.prop_sReportLatitudeDegree
        End Get
        Set
            Me.prop_sReportLatitudeDegree = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportLatitudeMinute")>
    Public Property sReportLatitudeMinute() As Double
        Get
            Return Me.prop_sReportLatitudeMinute
        End Get
        Set
            Me.prop_sReportLatitudeMinute = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportLatitudeSecond")>
    Public Property sReportLatitudeSecond() As Double
        Get
            Return Me.prop_sReportLatitudeSecond
        End Get
        Set
            Me.prop_sReportLatitudeSecond = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportLongitudeDegree")>
    Public Property sReportLongitudeDegree() As Double
        Get
            Return Me.prop_sReportLongitudeDegree
        End Get
        Set
            Me.prop_sReportLongitudeDegree = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportLongitudeMinute")>
    Public Property sReportLongitudeMinute() As Double
        Get
            Return Me.prop_sReportLongitudeMinute
        End Get
        Set
            Me.prop_sReportLongitudeMinute = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportLongitudeSecond")>
    Public Property sReportLongitudeSecond() As Double
        Get
            Return Me.prop_sReportLongitudeSecond
        End Get
        Set
            Me.prop_sReportLongitudeSecond = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportLocalCodeRequirement")>
    Public Property sReportLocalCodeRequirement() As String
        Get
            Return Me.prop_sReportLocalCodeRequirement
        End Get
        Set
            Me.prop_sReportLocalCodeRequirement = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportSiteHistory")>
    Public Property sReportSiteHistory() As String
        Get
            Return Me.prop_sReportSiteHistory
        End Get
        Set
            Me.prop_sReportSiteHistory = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportTowerManufacturer")>
    Public Property sReportTowerManufacturer() As String
        Get
            Return Me.prop_sReportTowerManufacturer
        End Get
        Set
            Me.prop_sReportTowerManufacturer = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportMonthManufactured")>
    Public Property sReportMonthManufactured() As String
        Get
            Return Me.prop_sReportMonthManufactured
        End Get
        Set
            Me.prop_sReportMonthManufactured = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportYearManufactured")>
    Public Property sReportYearManufactured() As Integer
        Get
            Return Me.prop_sReportYearManufactured
        End Get
        Set
            Me.prop_sReportYearManufactured = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportOriginalSpeed")>
    Public Property sReportOriginalSpeed() As Double
        Get
            Return Me.prop_sReportOriginalSpeed
        End Get
        Set
            Me.prop_sReportOriginalSpeed = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportOriginalCode")>
    Public Property sReportOriginalCode() As String
        Get
            Return Me.prop_sReportOriginalCode
        End Get
        Set
            Me.prop_sReportOriginalCode = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportTowerType")>
    Public Property sReportTowerType() As String
        Get
            Return Me.prop_sReportTowerType
        End Get
        Set
            Me.prop_sReportTowerType = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportEngrName")>
    Public Property sReportEngrName() As String
        Get
            Return Me.prop_sReportEngrName
        End Get
        Set
            Me.prop_sReportEngrName = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportEngrTitle")>
    Public Property sReportEngrTitle() As String
        Get
            Return Me.prop_sReportEngrTitle
        End Get
        Set
            Me.prop_sReportEngrTitle = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportHQPhoneNumber")>
    Public Property sReportHQPhoneNumber() As String
        Get
            Return Me.prop_sReportHQPhoneNumber
        End Get
        Set
            Me.prop_sReportHQPhoneNumber = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportEmailAddress")>
    Public Property sReportEmailAddress() As String
        Get
            Return Me.prop_sReportEmailAddress
        End Get
        Set
            Me.prop_sReportEmailAddress = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportLogoPath")>
    Public Property sReportLogoPath() As String
        Get
            Return Me.prop_sReportLogoPath
        End Get
        Set
            Me.prop_sReportLogoPath = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportCCiContactName")>
    Public Property sReportCCiContactName() As String
        Get
            Return Me.prop_sReportCCiContactName
        End Get
        Set
            Me.prop_sReportCCiContactName = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportCCiAddress1")>
    Public Property sReportCCiAddress1() As String
        Get
            Return Me.prop_sReportCCiAddress1
        End Get
        Set
            Me.prop_sReportCCiAddress1 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportCCiAddress2")>
    Public Property sReportCCiAddress2() As String
        Get
            Return Me.prop_sReportCCiAddress2
        End Get
        Set
            Me.prop_sReportCCiAddress2 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportCCiBUNumber")>
    Public Property sReportCCiBUNumber() As String
        Get
            Return Me.prop_sReportCCiBUNumber
        End Get
        Set
            Me.prop_sReportCCiBUNumber = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportCCiSiteName")>
    Public Property sReportCCiSiteName() As String
        Get
            Return Me.prop_sReportCCiSiteName
        End Get
        Set
            Me.prop_sReportCCiSiteName = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportCCiJDENumber")>
    Public Property sReportCCiJDENumber() As String
        Get
            Return Me.prop_sReportCCiJDENumber
        End Get
        Set
            Me.prop_sReportCCiJDENumber = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportCCiWONumber")>
    Public Property sReportCCiWONumber() As String
        Get
            Return Me.prop_sReportCCiWONumber
        End Get
        Set
            Me.prop_sReportCCiWONumber = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportCCiPONumber")>
    Public Property sReportCCiPONumber() As String
        Get
            Return Me.prop_sReportCCiPONumber
        End Get
        Set
            Me.prop_sReportCCiPONumber = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportCCiAppNumber")>
    Public Property sReportCCiAppNumber() As String
        Get
            Return Me.prop_sReportCCiAppNumber
        End Get
        Set
            Me.prop_sReportCCiAppNumber = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportCCiRevNumber")>
    Public Property sReportCCiRevNumber() As String
        Get
            Return Me.prop_sReportCCiRevNumber
        End Get
        Set
            Me.prop_sReportCCiRevNumber = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description("Reference Document Row. String format: Doc Type<~~>Remarks<~~>Ref No<~~>Source"), DisplayName("sReportDocsProvided")>
    Public Property sReportDocsProvided() As List(Of String)
        Get
            Return Me.prop_sReportDocsProvided
        End Get
        Set
            Me.prop_sReportDocsProvided = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportRecommendations")>
    Public Property sReportRecommendations() As String
        Get
            Return Me.prop_sReportRecommendations
        End Get
        Set
            Me.prop_sReportRecommendations = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description("Proposed Equipment Row. String format: MCL<~~>ECL<~~>qty<~~>manufacturer<~~>model<~~>FL qty<~~>FL Size<~~>Note #<~~>?<~~>Proposed"), DisplayName("sReportAppurt1")>
    Public Property sReportAppurt1() As List(Of String)
        Get
            Return Me.prop_sReportAppurt1
        End Get
        Set
            Me.prop_sReportAppurt1 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description("Existing Equipment Row. String format:MCL<~~>ECL<~~>qty<~~>manufacturer<~~>model<~~>FL qty<~~>FL Size<~~>Note #<~~>?<~~>Existing"), DisplayName("sReportAppurt2")>
    Public Property sReportAppurt2() As List(Of String)
        Get
            Return Me.prop_sReportAppurt2
        End Get
        Set
            Me.prop_sReportAppurt2 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description("Design Equipment Row. String format: MCL<~~>ECL<~~>qty<~~>manufacturer<~~>model<~~>FL qty<~~>FL Size<~~>"), DisplayName("sReportAppurt2")>
    Public Property sReportAppurt3() As List(Of String)
        Get
            Return Me.prop_sReportAppurt3
        End Get
        Set
            Me.prop_sReportAppurt3 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description("Additional Capacity Row. String format: Component<~~>Note #<~~>Elevation<~~>Cap%<~~>Pass/Fail<~~>Include in Report {Yes/No}"), DisplayName("sReportAddlCapacity")>
    Public Property sReportAddlCapacity() As List(Of String)
        Get
            Return Me.prop_sReportAddlCapacity
        End Get
        Set
            Me.prop_sReportAddlCapacity = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAssumption")>
    Public Property sReportAssumption() As List(Of String)
        Get
            Return Me.prop_sReportAssumption
        End Get
        Set
            Me.prop_sReportAssumption = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAppurt1Note1")>
    Public Property sReportAppurt1Note1() As String
        Get
            Return Me.prop_sReportAppurt1Note1
        End Get
        Set
            Me.prop_sReportAppurt1Note1 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAppurt1Note2")>
    Public Property sReportAppurt1Note2() As String
        Get
            Return Me.prop_sReportAppurt1Note2
        End Get
        Set
            Me.prop_sReportAppurt1Note2 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAppurt1Note3")>
    Public Property sReportAppurt1Note3() As String
        Get
            Return Me.prop_sReportAppurt1Note3
        End Get
        Set
            Me.prop_sReportAppurt1Note3 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAppurt1Note4")>
    Public Property sReportAppurt1Note4() As String
        Get
            Return Me.prop_sReportAppurt1Note4
        End Get
        Set
            Me.prop_sReportAppurt1Note4 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAppurt1Note5")>
    Public Property sReportAppurt1Note5() As String
        Get
            Return Me.prop_sReportAppurt1Note5
        End Get
        Set
            Me.prop_sReportAppurt1Note5 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAppurt1Note6")>
    Public Property sReportAppurt1Note6() As String
        Get
            Return Me.prop_sReportAppurt1Note6
        End Get
        Set
            Me.prop_sReportAppurt1Note6 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAppurt1Note7")>
    Public Property sReportAppurt1Note7() As String
        Get
            Return Me.prop_sReportAppurt1Note7
        End Get
        Set
            Me.prop_sReportAppurt1Note7 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAppurt2Note1")>
    Public Property sReportAppurt2Note1() As String
        Get
            Return Me.prop_sReportAppurt2Note1
        End Get
        Set
            Me.prop_sReportAppurt2Note1 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAppurt2Note2")>
    Public Property sReportAppurt2Note2() As String
        Get
            Return Me.prop_sReportAppurt2Note2
        End Get
        Set
            Me.prop_sReportAppurt2Note2 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAppurt2Note3")>
    Public Property sReportAppurt2Note3() As String
        Get
            Return Me.prop_sReportAppurt2Note3
        End Get
        Set
            Me.prop_sReportAppurt2Note3 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAppurt2Note4")>
    Public Property sReportAppurt2Note4() As String
        Get
            Return Me.prop_sReportAppurt2Note4
        End Get
        Set
            Me.prop_sReportAppurt2Note4 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAppurt2Note5")>
    Public Property sReportAppurt2Note5() As String
        Get
            Return Me.prop_sReportAppurt2Note5
        End Get
        Set
            Me.prop_sReportAppurt2Note5 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAppurt2Note6")>
    Public Property sReportAppurt2Note6() As String
        Get
            Return Me.prop_sReportAppurt2Note6
        End Get
        Set
            Me.prop_sReportAppurt2Note6 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAppurt2Note7")>
    Public Property sReportAppurt2Note7() As String
        Get
            Return Me.prop_sReportAppurt2Note7
        End Get
        Set
            Me.prop_sReportAppurt2Note7 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAddlCapacityNote1")>
    Public Property sReportAddlCapacityNote1() As String
        Get
            Return Me.prop_sReportAddlCapacityNote1
        End Get
        Set
            Me.prop_sReportAddlCapacityNote1 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAddlCapacityNote2")>
    Public Property sReportAddlCapacityNote2() As String
        Get
            Return Me.prop_sReportAddlCapacityNote2
        End Get
        Set
            Me.prop_sReportAddlCapacityNote2 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAddlCapacityNote3")>
    Public Property sReportAddlCapacityNote3() As String
        Get
            Return Me.prop_sReportAddlCapacityNote3
        End Get
        Set
            Me.prop_sReportAddlCapacityNote3 = Value
        End Set
    End Property
    <Category("TNX CCI Report"), Description(""), DisplayName("sReportAddlCapacityNote4")>
    Public Property sReportAddlCapacityNote4() As String
        Get
            Return Me.prop_sReportAddlCapacityNote4
        End Get
        Set
            Me.prop_sReportAddlCapacityNote4 = Value
        End Set
    End Property

End Class
