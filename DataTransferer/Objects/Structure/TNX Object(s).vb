Option Strict On

Imports System.ComponentModel
Imports System.Data
Imports System.IO

Partial Public Class TNX_model
    Private prop_units As tnx_units
    'Private prop_code_data As tnx_code
    Private prop_upper_structure As List(Of tnx_antenna_record)
    Private prop_base_structure As List(Of tnx_tower_record)
    Private prop_guy_wires As List(Of tnx_guy_record)
    'Private prop_feed_lines As List(Of tnx_feed_line)
    'Private prop_discrete_loads As List(Of tnx_discrete)
    'Private prop_dishes As List(Of tnx_dish)

    <Category("TNX"), Description(""), DisplayName("prop_units")>
    Public Property units() As tnx_units
        Get
            Return Me.prop_units
        End Get
        Set
            Me.prop_units = Value
        End Set
    End Property
    <Category("TNX"), Description(""), DisplayName("Upper Structure")>
    Public Property upper_structure() As List(Of tnx_antenna_record)
        Get
            Return Me.prop_upper_structure
        End Get
        Set
            Me.prop_upper_structure = Value
        End Set
    End Property
    <Category("TNX"), Description(""), DisplayName("Base Structure")>
    Public Property base_structure() As List(Of tnx_tower_record)
        Get
            Return Me.prop_base_structure
        End Get
        Set
            Me.prop_base_structure = Value
        End Set
    End Property
    <Category("TNX"), Description(""), DisplayName("Guy Wires")>
    Public Property guy_wires() As List(Of tnx_guy_record)
        Get
            Return Me.prop_guy_wires
        End Get
        Set
            Me.prop_guy_wires = Value
        End Set
    End Property

    Public Sub New()
        'Leave method empty
    End Sub

    <Category("Constructor"), Description("Create TNX object from TNX file.")>
    Public Sub New(ByVal tnxPath As String)

        Dim USUnits As Boolean = False
        'instatiate lists
        'Me.upper_structure = New List(Of tnx_antenna_record)
        'Me.base_structure = New List(Of tnx_tower_record)
        'Me.guy_wires = New List(Of tnx_guy_record)

        For Each line In File.ReadLines(tnxPath)

            Dim tnxVar As String
            Dim tnxValue As String
            Dim recIndex As Integer

            If Not line.Contains("=") Then
                tnxVar = line
                tnxValue = ""
            Else
                tnxVar = Left(line, line.IndexOf("="))
                tnxValue = Right(line, Len(line) - line.IndexOf("=") - 1)
            End If

            Select Case tnxVar
                    ''''Units''''
                Case "UnitsSystem"
                    If tnxValue <> "US" Then
                        Throw New System.Exception("TNX file is not in US units.")
                    End If
                Case "[US Units]"
                    USUnits = True
                    Me.units = New tnx_units
                Case "[SI Units]"
                    USUnits = False
                Case "Length"
                    Try
                        If USUnits Then Me.units.Length = New tnx_length_unit(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "LengthPrec"
                    Try
                        If USUnits Then Me.units.Length.precision = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "Coordinate"
                    Try
                        If USUnits Then Me.units.Coordinate = New tnx_coordinate_unit(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "CoordinatePrec"
                    Try
                        If USUnits Then Me.units.Coordinate.precision = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "Force"
                    Try
                        If USUnits Then Me.units.Force = New tnx_force_unit(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "ForcePrec"
                    Try
                        If USUnits Then Me.units.Force.precision = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "Load"
                    Try
                        If USUnits Then Me.units.Load = New tnx_load_unit(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "LoadPrec"
                    Try
                        If USUnits Then Me.units.Load.precision = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "Moment"
                    Try
                        If USUnits Then Me.units.Moment = New tnx_moment_unit(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "MomentPrec"
                    Try
                        If USUnits Then Me.units.Moment.precision = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "Properties"
                    Try
                        If USUnits Then Me.units.Properties = New tnx_properties_unit(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "PropertiesPrec"
                    Try
                        If USUnits Then Me.units.Properties.precision = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "Pressure"
                    Try
                        If USUnits Then Me.units.Pressure = New tnx_pressure_unit(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "PressurePrec"
                    Try
                        If USUnits Then Me.units.Pressure.precision = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "Velocity"
                    Try
                        If USUnits Then Me.units.Velocity = New tnx_velocity_unit(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "VelocityPrec"
                    Try
                        If USUnits Then Me.units.Velocity.precision = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "Displacement"
                    Try
                        If USUnits Then Me.units.Displacement = New tnx_displacement_unit(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "DisplacementPrec"
                    Try
                        If USUnits Then Me.units.Displacement.precision = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "Mass"
                    Try
                        If USUnits Then Me.units.Mass = New tnx_mass_unit(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "MassPrec"
                    Try
                        If USUnits Then Me.units.Mass.precision = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "Acceleration"
                    Try
                        If USUnits Then Me.units.Acceleration = New tnx_acceleration_unit(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AccelerationPrec"
                    Try
                        If USUnits Then Me.units.Acceleration.precision = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "Stress"
                    Try
                        If USUnits Then Me.units.Stress = New tnx_stress_unit(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "StressPrec"
                    Try
                        If USUnits Then Me.units.Stress.precision = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "Density"
                    Try
                        If USUnits Then Me.units.Density = New tnx_density_unit(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "DensityPrec"
                    Try
                        If USUnits Then Me.units.Density.precision = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "UnitWt"
                    Try
                        If USUnits Then Me.units.UnitWt = New tnx_unitwt_unit(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "UnitWtPrec"
                    Try
                        If USUnits Then Me.units.UnitWt.precision = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "Strength"
                    Try
                        If USUnits Then Me.units.Strength = New tnx_strength_unit(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "StrengthPrec"
                    Try
                        If USUnits Then Me.units.Strength.precision = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "Modulus"
                    Try
                        If USUnits Then Me.units.Modulus = New tnx_modulus_unit(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "ModulusPrec"
                    Try
                        If USUnits Then Me.units.Modulus.precision = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "Temperature"
                    Try
                        If USUnits Then Me.units.Temperature = New tnx_temperature_unit(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TemperaturePrec"
                    Try
                        If USUnits Then Me.units.Temperature.precision = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "Printer"
                    Try
                        If USUnits Then Me.units.Printer = New tnx_printer_unit(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "PrinterPrec"
                    Try
                        If USUnits Then Me.units.Printer.precision = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "Rotation"
                    Try
                        If USUnits Then Me.units.Rotation = New tnx_rotation_unit(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "RotationPrec"
                    Try
                        If USUnits Then Me.units.Rotation.precision = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "Spacing"
                    Try
                        If USUnits Then Me.units.Spacing = New tnx_spacing_unit(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "SpacingPrec"
                    Try
                        If USUnits Then Me.units.Spacing.precision = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                    ''''Antenna Rec (Upper Structure)''''
                Case "NumAntennaRecs"
                    Try
                        If CInt(tnxValue) > 0 Then
                            Me.upper_structure = New List(Of tnx_antenna_record)
                        End If
                    Catch ex As Exception
                    End Try
                Case "AntennaRec"
                    Try
                        recIndex = CInt(tnxValue) - 1
                        Me.upper_structure.Add(New tnx_antenna_record())
                        Me.upper_structure(recIndex).AntennaRec = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaBraceType"
                    Try
                        Me.upper_structure(recIndex).AntennaBraceType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaHeight"
                    Try
                        Me.upper_structure(recIndex).AntennaHeight = Me.units.Coordinate.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaDiagonalSpacing"
                    Try
                        Me.upper_structure(recIndex).AntennaDiagonalSpacing = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaDiagonalSpacingEx"
                    Try
                        Me.upper_structure(recIndex).AntennaDiagonalSpacingEx = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaNumSections"
                    Try
                        Me.upper_structure(recIndex).AntennaNumSections = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaNumSesctions"
                    Try
                        Me.upper_structure(recIndex).AntennaNumSesctions = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaSectionLength"
                    Try
                        Me.upper_structure(recIndex).AntennaSectionLength = Me.units.Length.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaLegType"
                    Try
                        Me.upper_structure(recIndex).AntennaLegType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaLegSize"
                    Try
                        Me.upper_structure(recIndex).AntennaLegSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaLegGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaLegGrade = Me.units.Strength.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaLegMatlGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaLegMatlGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaDiagonalGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaDiagonalGrade = Me.units.Strength.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaDiagonalMatlGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaDiagonalMatlGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaInnerBracingGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaInnerBracingGrade = Me.units.Strength.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaInnerBracingMatlGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaInnerBracingMatlGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaTopGirtGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaTopGirtGrade = Me.units.Strength.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaTopGirtMatlGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaTopGirtMatlGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaBotGirtGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaBotGirtGrade = Me.units.Strength.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaBotGirtMatlGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaBotGirtMatlGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaInnerGirtGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaInnerGirtGrade = Me.units.Strength.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaInnerGirtMatlGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaInnerGirtMatlGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaLongHorizontalGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaLongHorizontalGrade = Me.units.Strength.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaLongHorizontalMatlGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaLongHorizontalMatlGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaShortHorizontalGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaShortHorizontalGrade = Me.units.Strength.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaShortHorizontalMatlGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaShortHorizontalMatlGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaDiagonalType"
                    Try
                        Me.upper_structure(recIndex).AntennaDiagonalType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaDiagonalSize"
                    Try
                        Me.upper_structure(recIndex).AntennaDiagonalSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaInnerBracingType"
                    Try
                        Me.upper_structure(recIndex).AntennaInnerBracingType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaInnerBracingSize"
                    Try
                        Me.upper_structure(recIndex).AntennaInnerBracingSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaTopGirtType"
                    Try
                        Me.upper_structure(recIndex).AntennaTopGirtType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaTopGirtSize"
                    Try
                        Me.upper_structure(recIndex).AntennaTopGirtSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaBotGirtType"
                    Try
                        Me.upper_structure(recIndex).AntennaBotGirtType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaBotGirtSize"
                    Try
                        Me.upper_structure(recIndex).AntennaBotGirtSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaTopGirtOffset"
                    Try
                        Me.upper_structure(recIndex).AntennaTopGirtOffset = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaBotGirtOffset"
                    Try
                        Me.upper_structure(recIndex).AntennaBotGirtOffset = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaHasKBraceEndPanels"
                    Try
                        Me.upper_structure(recIndex).AntennaHasKBraceEndPanels = CBool(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaHasHorizontals"
                    Try
                        Me.upper_structure(recIndex).AntennaHasHorizontals = CBool(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaLongHorizontalType"
                    Try
                        Me.upper_structure(recIndex).AntennaLongHorizontalType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaLongHorizontalSize"
                    Try
                        Me.upper_structure(recIndex).AntennaLongHorizontalSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaShortHorizontalType"
                    Try
                        Me.upper_structure(recIndex).AntennaShortHorizontalType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaShortHorizontalSize"
                    Try
                        Me.upper_structure(recIndex).AntennaShortHorizontalSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantGrade = Me.units.Strength.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantMatlGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantMatlGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantType"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantDiagType"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantDiagType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantSubDiagonalType"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantSubDiagonalType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantSubHorizontalType"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantSubHorizontalType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantVerticalType"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantVerticalType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHipType"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHipType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHipDiagonalType"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHipDiagonalType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHorizontalSize"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHorizontalSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHorizontalSize2"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHorizontalSize2 = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHorizontalSize3"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHorizontalSize3 = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHorizontalSize4"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHorizontalSize4 = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantDiagonalSize"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantDiagonalSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantDiagonalSize2"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantDiagonalSize2 = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantDiagonalSize3"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantDiagonalSize3 = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantDiagonalSize4"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantDiagonalSize4 = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantSubHorizontalSize"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantSubHorizontalSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantSubDiagonalSize"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantSubDiagonalSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaSubDiagLocation"
                    Try
                        Me.upper_structure(recIndex).AntennaSubDiagLocation = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantVerticalSize"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantVerticalSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHipDiagonalSize"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHipDiagonalSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHipDiagonalSize2"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHipDiagonalSize2 = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHipDiagonalSize3"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHipDiagonalSize3 = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHipDiagonalSize4"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHipDiagonalSize4 = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHipSize"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHipSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHipSize2"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHipSize2 = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHipSize3"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHipSize3 = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHipSize4"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHipSize4 = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaNumInnerGirts"
                    Try
                        Me.upper_structure(recIndex).AntennaNumInnerGirts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaInnerGirtType"
                    Try
                        Me.upper_structure(recIndex).AntennaInnerGirtType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaInnerGirtSize"
                    Try
                        Me.upper_structure(recIndex).AntennaInnerGirtSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaPoleShapeType"
                    Try
                        Me.upper_structure(recIndex).AntennaPoleShapeType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaPoleSize"
                    Try
                        Me.upper_structure(recIndex).AntennaPoleSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaPoleGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaPoleGrade = Me.units.Strength.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaPoleMatlGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaPoleMatlGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaPoleSpliceLength"
                    Try
                        Me.upper_structure(recIndex).AntennaPoleSpliceLength = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaTaperPoleNumSides"
                    Try
                        Me.upper_structure(recIndex).AntennaTaperPoleNumSides = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaTaperPoleTopDiameter"
                    Try
                        Me.upper_structure(recIndex).AntennaTaperPoleTopDiameter = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaTaperPoleBotDiameter"
                    Try
                        Me.upper_structure(recIndex).AntennaTaperPoleBotDiameter = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaTaperPoleWallThickness"
                    Try
                        Me.upper_structure(recIndex).AntennaTaperPoleWallThickness = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaTaperPoleBendRadius"
                    Try
                        Me.upper_structure(recIndex).AntennaTaperPoleBendRadius = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaTaperPoleGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaTaperPoleGrade = Me.units.Strength.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaTaperPoleMatlGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaTaperPoleMatlGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaSWMult"
                    Try
                        Me.upper_structure(recIndex).AntennaSWMult = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaWPMult"
                    Try
                        Me.upper_structure(recIndex).AntennaWPMult = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaAutoCalcKSingleAngle"
                    Try
                        Me.upper_structure(recIndex).AntennaAutoCalcKSingleAngle = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaAutoCalcKSolidRound"
                    Try
                        Me.upper_structure(recIndex).AntennaAutoCalcKSolidRound = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaAfGusset"
                    Try
                        Me.upper_structure(recIndex).AntennaAfGusset = Me.units.Length.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaTfGusset"
                    Try
                        Me.upper_structure(recIndex).AntennaTfGusset = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaGussetBoltEdgeDistance"
                    Try
                        Me.upper_structure(recIndex).AntennaGussetBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaGussetGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaGussetGrade = Me.units.Strength.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaGussetMatlGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaGussetMatlGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaAfMult"
                    Try
                        Me.upper_structure(recIndex).AntennaAfMult = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaArMult"
                    Try
                        Me.upper_structure(recIndex).AntennaArMult = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaFlatIPAPole"
                    Try
                        Me.upper_structure(recIndex).AntennaFlatIPAPole = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaRoundIPAPole"
                    Try
                        Me.upper_structure(recIndex).AntennaRoundIPAPole = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaFlatIPALeg"
                    Try
                        Me.upper_structure(recIndex).AntennaFlatIPALeg = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaRoundIPALeg"
                    Try
                        Me.upper_structure(recIndex).AntennaRoundIPALeg = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaFlatIPAHorizontal"
                    Try
                        Me.upper_structure(recIndex).AntennaFlatIPAHorizontal = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaRoundIPAHorizontal"
                    Try
                        Me.upper_structure(recIndex).AntennaRoundIPAHorizontal = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaFlatIPADiagonal"
                    Try
                        Me.upper_structure(recIndex).AntennaFlatIPADiagonal = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaRoundIPADiagonal"
                    Try
                        Me.upper_structure(recIndex).AntennaRoundIPADiagonal = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaCSA_S37_SpeedUpFactor"
                    Try
                        Me.upper_structure(recIndex).AntennaCSA_S37_SpeedUpFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaKLegs"
                    Try
                        Me.upper_structure(recIndex).AntennaKLegs = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaKXBracedDiags"
                    Try
                        Me.upper_structure(recIndex).AntennaKXBracedDiags = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaKKBracedDiags"
                    Try
                        Me.upper_structure(recIndex).AntennaKKBracedDiags = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaKZBracedDiags"
                    Try
                        Me.upper_structure(recIndex).AntennaKZBracedDiags = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaKHorzs"
                    Try
                        Me.upper_structure(recIndex).AntennaKHorzs = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaKSecHorzs"
                    Try
                        Me.upper_structure(recIndex).AntennaKSecHorzs = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaKGirts"
                    Try
                        Me.upper_structure(recIndex).AntennaKGirts = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaKInners"
                    Try
                        Me.upper_structure(recIndex).AntennaKInners = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaKXBracedDiagsY"
                    Try
                        Me.upper_structure(recIndex).AntennaKXBracedDiagsY = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaKKBracedDiagsY"
                    Try
                        Me.upper_structure(recIndex).AntennaKKBracedDiagsY = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaKZBracedDiagsY"
                    Try
                        Me.upper_structure(recIndex).AntennaKZBracedDiagsY = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaKHorzsY"
                    Try
                        Me.upper_structure(recIndex).AntennaKHorzsY = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaKSecHorzsY"
                    Try
                        Me.upper_structure(recIndex).AntennaKSecHorzsY = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaKGirtsY"
                    Try
                        Me.upper_structure(recIndex).AntennaKGirtsY = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaKInnersY"
                    Try
                        Me.upper_structure(recIndex).AntennaKInnersY = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaKRedHorz"
                    Try
                        Me.upper_structure(recIndex).AntennaKRedHorz = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaKRedDiag"
                    Try
                        Me.upper_structure(recIndex).AntennaKRedDiag = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaKRedSubDiag"
                    Try
                        Me.upper_structure(recIndex).AntennaKRedSubDiag = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaKRedSubHorz"
                    Try
                        Me.upper_structure(recIndex).AntennaKRedSubHorz = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaKRedVert"
                    Try
                        Me.upper_structure(recIndex).AntennaKRedVert = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaKRedHip"
                    Try
                        Me.upper_structure(recIndex).AntennaKRedHip = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaKRedHipDiag"
                    Try
                        Me.upper_structure(recIndex).AntennaKRedHipDiag = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaKTLX"
                    Try
                        Me.upper_structure(recIndex).AntennaKTLX = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaKTLZ"
                    Try
                        Me.upper_structure(recIndex).AntennaKTLZ = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaKTLLeg"
                    Try
                        Me.upper_structure(recIndex).AntennaKTLLeg = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaInnerKTLX"
                    Try
                        Me.upper_structure(recIndex).AntennaInnerKTLX = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaInnerKTLZ"
                    Try
                        Me.upper_structure(recIndex).AntennaInnerKTLZ = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaInnerKTLLeg"
                    Try
                        Me.upper_structure(recIndex).AntennaInnerKTLLeg = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaStitchBoltLocationHoriz"
                    Try
                        Me.upper_structure(recIndex).AntennaStitchBoltLocationHoriz = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaStitchBoltLocationDiag"
                    Try
                        Me.upper_structure(recIndex).AntennaStitchBoltLocationDiag = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaStitchSpacing"
                    Try
                        Me.upper_structure(recIndex).AntennaStitchSpacing = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaStitchSpacingHorz"
                    Try
                        Me.upper_structure(recIndex).AntennaStitchSpacingHorz = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaStitchSpacingDiag"
                    Try
                        Me.upper_structure(recIndex).AntennaStitchSpacingDiag = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaStitchSpacingRed"
                    Try
                        Me.upper_structure(recIndex).AntennaStitchSpacingRed = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaLegNetWidthDeduct"
                    Try
                        Me.upper_structure(recIndex).AntennaLegNetWidthDeduct = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaLegUFactor"
                    Try
                        Me.upper_structure(recIndex).AntennaLegUFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaDiagonalNetWidthDeduct"
                    Try
                        Me.upper_structure(recIndex).AntennaDiagonalNetWidthDeduct = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaTopGirtNetWidthDeduct"
                    Try
                        Me.upper_structure(recIndex).AntennaTopGirtNetWidthDeduct = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaBotGirtNetWidthDeduct"
                    Try
                        Me.upper_structure(recIndex).AntennaBotGirtNetWidthDeduct = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaInnerGirtNetWidthDeduct"
                    Try
                        Me.upper_structure(recIndex).AntennaInnerGirtNetWidthDeduct = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaHorizontalNetWidthDeduct"
                    Try
                        Me.upper_structure(recIndex).AntennaHorizontalNetWidthDeduct = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaShortHorizontalNetWidthDeduct"
                    Try
                        Me.upper_structure(recIndex).AntennaShortHorizontalNetWidthDeduct = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaDiagonalUFactor"
                    Try
                        Me.upper_structure(recIndex).AntennaDiagonalUFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaTopGirtUFactor"
                    Try
                        Me.upper_structure(recIndex).AntennaTopGirtUFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaBotGirtUFactor"
                    Try
                        Me.upper_structure(recIndex).AntennaBotGirtUFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaInnerGirtUFactor"
                    Try
                        Me.upper_structure(recIndex).AntennaInnerGirtUFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaHorizontalUFactor"
                    Try
                        Me.upper_structure(recIndex).AntennaHorizontalUFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaShortHorizontalUFactor"
                    Try
                        Me.upper_structure(recIndex).AntennaShortHorizontalUFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaLegConnType"
                    Try
                        Me.upper_structure(recIndex).AntennaLegConnType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaLegNumBolts"
                    Try
                        Me.upper_structure(recIndex).AntennaLegNumBolts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaDiagonalNumBolts"
                    Try
                        Me.upper_structure(recIndex).AntennaDiagonalNumBolts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaTopGirtNumBolts"
                    Try
                        Me.upper_structure(recIndex).AntennaTopGirtNumBolts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaBotGirtNumBolts"
                    Try
                        Me.upper_structure(recIndex).AntennaBotGirtNumBolts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaInnerGirtNumBolts"
                    Try
                        Me.upper_structure(recIndex).AntennaInnerGirtNumBolts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaHorizontalNumBolts"
                    Try
                        Me.upper_structure(recIndex).AntennaHorizontalNumBolts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaShortHorizontalNumBolts"
                    Try
                        Me.upper_structure(recIndex).AntennaShortHorizontalNumBolts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaLegBoltGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaLegBoltGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaLegBoltSize"
                    Try
                        Me.upper_structure(recIndex).AntennaLegBoltSize = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaDiagonalBoltGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaDiagonalBoltGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaDiagonalBoltSize"
                    Try
                        Me.upper_structure(recIndex).AntennaDiagonalBoltSize = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaTopGirtBoltGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaTopGirtBoltGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaTopGirtBoltSize"
                    Try
                        Me.upper_structure(recIndex).AntennaTopGirtBoltSize = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaBotGirtBoltGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaBotGirtBoltGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaBotGirtBoltSize"
                    Try
                        Me.upper_structure(recIndex).AntennaBotGirtBoltSize = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaInnerGirtBoltGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaInnerGirtBoltGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaInnerGirtBoltSize"
                    Try
                        Me.upper_structure(recIndex).AntennaInnerGirtBoltSize = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaHorizontalBoltGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaHorizontalBoltGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaHorizontalBoltSize"
                    Try
                        Me.upper_structure(recIndex).AntennaHorizontalBoltSize = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaShortHorizontalBoltGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaShortHorizontalBoltGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaShortHorizontalBoltSize"
                    Try
                        Me.upper_structure(recIndex).AntennaShortHorizontalBoltSize = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaLegBoltEdgeDistance"
                    Try
                        Me.upper_structure(recIndex).AntennaLegBoltEdgeDistance = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaDiagonalBoltEdgeDistance"
                    Try
                        Me.upper_structure(recIndex).AntennaDiagonalBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaTopGirtBoltEdgeDistance"
                    Try
                        Me.upper_structure(recIndex).AntennaTopGirtBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaBotGirtBoltEdgeDistance"
                    Try
                        Me.upper_structure(recIndex).AntennaBotGirtBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaInnerGirtBoltEdgeDistance"
                    Try
                        Me.upper_structure(recIndex).AntennaInnerGirtBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaHorizontalBoltEdgeDistance"
                    Try
                        Me.upper_structure(recIndex).AntennaHorizontalBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaShortHorizontalBoltEdgeDistance"
                    Try
                        Me.upper_structure(recIndex).AntennaShortHorizontalBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaDiagonalGageG1Distance"
                    Try
                        Me.upper_structure(recIndex).AntennaDiagonalGageG1Distance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaTopGirtGageG1Distance"
                    Try
                        Me.upper_structure(recIndex).AntennaTopGirtGageG1Distance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaBotGirtGageG1Distance"
                    Try
                        Me.upper_structure(recIndex).AntennaBotGirtGageG1Distance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaInnerGirtGageG1Distance"
                    Try
                        Me.upper_structure(recIndex).AntennaInnerGirtGageG1Distance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaHorizontalGageG1Distance"
                    Try
                        Me.upper_structure(recIndex).AntennaHorizontalGageG1Distance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaShortHorizontalGageG1Distance"
                    Try
                        Me.upper_structure(recIndex).AntennaShortHorizontalGageG1Distance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHorizontalBoltGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHorizontalBoltGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHorizontalBoltSize"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHorizontalBoltSize = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHorizontalNumBolts"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHorizontalNumBolts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHorizontalBoltEdgeDistance"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHorizontalBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHorizontalGageG1Distance"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHorizontalGageG1Distance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHorizontalNetWidthDeduct"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHorizontalNetWidthDeduct = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHorizontalUFactor"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHorizontalUFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantDiagonalBoltGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantDiagonalBoltGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantDiagonalBoltSize"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantDiagonalBoltSize = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantDiagonalNumBolts"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantDiagonalNumBolts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantDiagonalBoltEdgeDistance"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantDiagonalBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantDiagonalGageG1Distance"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantDiagonalGageG1Distance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantDiagonalNetWidthDeduct"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantDiagonalNetWidthDeduct = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantDiagonalUFactor"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantDiagonalUFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantSubDiagonalBoltGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantSubDiagonalBoltGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantSubDiagonalBoltSize"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantSubDiagonalBoltSize = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantSubDiagonalNumBolts"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantSubDiagonalNumBolts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantSubDiagonalBoltEdgeDistance"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantSubDiagonalBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantSubDiagonalGageG1Distance"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantSubDiagonalGageG1Distance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantSubDiagonalNetWidthDeduct"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantSubDiagonalNetWidthDeduct = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantSubDiagonalUFactor"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantSubDiagonalUFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantSubHorizontalBoltGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantSubHorizontalBoltGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantSubHorizontalBoltSize"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantSubHorizontalBoltSize = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantSubHorizontalNumBolts"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantSubHorizontalNumBolts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantSubHorizontalBoltEdgeDistance"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantSubHorizontalBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantSubHorizontalGageG1Distance"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantSubHorizontalGageG1Distance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantSubHorizontalNetWidthDeduct"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantSubHorizontalNetWidthDeduct = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantSubHorizontalUFactor"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantSubHorizontalUFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantVerticalBoltGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantVerticalBoltGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantVerticalBoltSize"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantVerticalBoltSize = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantVerticalNumBolts"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantVerticalNumBolts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantVerticalBoltEdgeDistance"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantVerticalBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantVerticalGageG1Distance"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantVerticalGageG1Distance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantVerticalNetWidthDeduct"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantVerticalNetWidthDeduct = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantVerticalUFactor"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantVerticalUFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHipBoltGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHipBoltGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHipBoltSize"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHipBoltSize = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHipNumBolts"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHipNumBolts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHipBoltEdgeDistance"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHipBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHipGageG1Distance"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHipGageG1Distance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHipNetWidthDeduct"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHipNetWidthDeduct = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHipUFactor"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHipUFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHipDiagonalBoltGrade"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHipDiagonalBoltGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHipDiagonalBoltSize"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHipDiagonalBoltSize = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHipDiagonalNumBolts"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHipDiagonalNumBolts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHipDiagonalBoltEdgeDistance"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHipDiagonalBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHipDiagonalGageG1Distance"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHipDiagonalGageG1Distance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHipDiagonalNetWidthDeduct"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHipDiagonalNetWidthDeduct = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaRedundantHipDiagonalUFactor"
                    Try
                        Me.upper_structure(recIndex).AntennaRedundantHipDiagonalUFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaDiagonalOutOfPlaneRestraint"
                    Try
                        Me.upper_structure(recIndex).AntennaDiagonalOutOfPlaneRestraint = CBool(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaTopGirtOutOfPlaneRestraint"
                    Try
                        Me.upper_structure(recIndex).AntennaTopGirtOutOfPlaneRestraint = CBool(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaBottomGirtOutOfPlaneRestraint"
                    Try
                        Me.upper_structure(recIndex).AntennaBottomGirtOutOfPlaneRestraint = CBool(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaMidGirtOutOfPlaneRestraint"
                    Try
                        Me.upper_structure(recIndex).AntennaMidGirtOutOfPlaneRestraint = CBool(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaHorizontalOutOfPlaneRestraint"
                    Try
                        Me.upper_structure(recIndex).AntennaHorizontalOutOfPlaneRestraint = CBool(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaSecondaryHorizontalOutOfPlaneRestraint"
                    Try
                        Me.upper_structure(recIndex).AntennaSecondaryHorizontalOutOfPlaneRestraint = CBool(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "AntennaDiagOffsetNEY"
                    Try
                        Me.upper_structure(recIndex).AntennaDiagOffsetNEY = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaDiagOffsetNEX"
                    Try
                        Me.upper_structure(recIndex).AntennaDiagOffsetNEX = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaDiagOffsetPEY"
                    Try
                        Me.upper_structure(recIndex).AntennaDiagOffsetPEY = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaDiagOffsetPEX"
                    Try
                        Me.upper_structure(recIndex).AntennaDiagOffsetPEX = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaKbraceOffsetNEY"
                    Try
                        Me.upper_structure(recIndex).AntennaKbraceOffsetNEY = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaKbraceOffsetNEX"
                    Try
                        Me.upper_structure(recIndex).AntennaKbraceOffsetNEX = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaKbraceOffsetPEY"
                    Try
                        Me.upper_structure(recIndex).AntennaKbraceOffsetPEY = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "AntennaKbraceOffsetPEX"
                    Try
                        Me.upper_structure(recIndex).AntennaKbraceOffsetPEX = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try

                    ''''Tower Rec (Base Structure)''''
                Case "NumTowerRecs"
                    Try
                        If CInt(tnxValue) > 0 Then
                            Me.base_structure = New List(Of tnx_tower_record)
                        End If
                    Catch ex As Exception
                    End Try
                Case "TowerRec"
                    Try
                        recIndex = CInt(tnxValue) - 1
                        Me.base_structure.Add(New tnx_tower_record())
                        Me.base_structure(recIndex).TowerRec = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerDatabase"
                    Try
                        Me.base_structure(recIndex).TowerDatabase = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerName"
                    Try
                        Me.base_structure(recIndex).TowerName = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerHeight"
                    Try
                        Me.base_structure(recIndex).TowerHeight = Me.units.Coordinate.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerFaceWidth"
                    Try
                        Me.base_structure(recIndex).TowerFaceWidth = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerNumSections"
                    Try
                        Me.base_structure(recIndex).TowerNumSections = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerSectionLength"
                    Try
                        Me.base_structure(recIndex).TowerSectionLength = Me.units.Length.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerDiagonalSpacing"
                    Try
                        Me.base_structure(recIndex).TowerDiagonalSpacing = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerDiagonalSpacingEx"
                    Try
                        Me.base_structure(recIndex).TowerDiagonalSpacingEx = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerBraceType"
                    Try
                        Me.base_structure(recIndex).TowerBraceType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerFaceBevel"
                    Try
                        Me.base_structure(recIndex).TowerFaceBevel = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerTopGirtOffset"
                    Try
                        Me.base_structure(recIndex).TowerTopGirtOffset = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerBotGirtOffset"
                    Try
                        Me.base_structure(recIndex).TowerBotGirtOffset = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerHasKBraceEndPanels"
                    Try
                        Me.base_structure(recIndex).TowerHasKBraceEndPanels = CBool(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerHasHorizontals"
                    Try
                        Me.base_structure(recIndex).TowerHasHorizontals = CBool(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerLegType"
                    Try
                        Me.base_structure(recIndex).TowerLegType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerLegSize"
                    Try
                        Me.base_structure(recIndex).TowerLegSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerLegGrade"
                    Try
                        Me.base_structure(recIndex).TowerLegGrade = Me.units.Strength.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerLegMatlGrade"
                    Try
                        Me.base_structure(recIndex).TowerLegMatlGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerDiagonalGrade"
                    Try
                        Me.base_structure(recIndex).TowerDiagonalGrade = Me.units.Strength.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerDiagonalMatlGrade"
                    Try
                        Me.base_structure(recIndex).TowerDiagonalMatlGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerInnerBracingGrade"
                    Try
                        Me.base_structure(recIndex).TowerInnerBracingGrade = Me.units.Strength.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerInnerBracingMatlGrade"
                    Try
                        Me.base_structure(recIndex).TowerInnerBracingMatlGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerTopGirtGrade"
                    Try
                        Me.base_structure(recIndex).TowerTopGirtGrade = Me.units.Strength.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerTopGirtMatlGrade"
                    Try
                        Me.base_structure(recIndex).TowerTopGirtMatlGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerBotGirtGrade"
                    Try
                        Me.base_structure(recIndex).TowerBotGirtGrade = Me.units.Strength.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerBotGirtMatlGrade"
                    Try
                        Me.base_structure(recIndex).TowerBotGirtMatlGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerInnerGirtGrade"
                    Try
                        Me.base_structure(recIndex).TowerInnerGirtGrade = Me.units.Strength.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerInnerGirtMatlGrade"
                    Try
                        Me.base_structure(recIndex).TowerInnerGirtMatlGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerLongHorizontalGrade"
                    Try
                        Me.base_structure(recIndex).TowerLongHorizontalGrade = Me.units.Strength.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerLongHorizontalMatlGrade"
                    Try
                        Me.base_structure(recIndex).TowerLongHorizontalMatlGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerShortHorizontalGrade"
                    Try
                        Me.base_structure(recIndex).TowerShortHorizontalGrade = Me.units.Strength.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerShortHorizontalMatlGrade"
                    Try
                        Me.base_structure(recIndex).TowerShortHorizontalMatlGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerDiagonalType"
                    Try
                        Me.base_structure(recIndex).TowerDiagonalType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerDiagonalSize"
                    Try
                        Me.base_structure(recIndex).TowerDiagonalSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerInnerBracingType"
                    Try
                        Me.base_structure(recIndex).TowerInnerBracingType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerInnerBracingSize"
                    Try
                        Me.base_structure(recIndex).TowerInnerBracingSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerTopGirtType"
                    Try
                        Me.base_structure(recIndex).TowerTopGirtType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerTopGirtSize"
                    Try
                        Me.base_structure(recIndex).TowerTopGirtSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerBotGirtType"
                    Try
                        Me.base_structure(recIndex).TowerBotGirtType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerBotGirtSize"
                    Try
                        Me.base_structure(recIndex).TowerBotGirtSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerNumInnerGirts"
                    Try
                        Me.base_structure(recIndex).TowerNumInnerGirts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerInnerGirtType"
                    Try
                        Me.base_structure(recIndex).TowerInnerGirtType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerInnerGirtSize"
                    Try
                        Me.base_structure(recIndex).TowerInnerGirtSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerLongHorizontalType"
                    Try
                        Me.base_structure(recIndex).TowerLongHorizontalType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerLongHorizontalSize"
                    Try
                        Me.base_structure(recIndex).TowerLongHorizontalSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerShortHorizontalType"
                    Try
                        Me.base_structure(recIndex).TowerShortHorizontalType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerShortHorizontalSize"
                    Try
                        Me.base_structure(recIndex).TowerShortHorizontalSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantGrade"
                    Try
                        Me.base_structure(recIndex).TowerRedundantGrade = Me.units.Strength.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantMatlGrade"
                    Try
                        Me.base_structure(recIndex).TowerRedundantMatlGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantType"
                    Try
                        Me.base_structure(recIndex).TowerRedundantType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantDiagType"
                    Try
                        Me.base_structure(recIndex).TowerRedundantDiagType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantSubDiagonalType"
                    Try
                        Me.base_structure(recIndex).TowerRedundantSubDiagonalType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantSubHorizontalType"
                    Try
                        Me.base_structure(recIndex).TowerRedundantSubHorizontalType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantVerticalType"
                    Try
                        Me.base_structure(recIndex).TowerRedundantVerticalType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHipType"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHipType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHipDiagonalType"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHipDiagonalType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHorizontalSize"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHorizontalSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHorizontalSize2"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHorizontalSize2 = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHorizontalSize3"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHorizontalSize3 = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHorizontalSize4"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHorizontalSize4 = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantDiagonalSize"
                    Try
                        Me.base_structure(recIndex).TowerRedundantDiagonalSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantDiagonalSize2"
                    Try
                        Me.base_structure(recIndex).TowerRedundantDiagonalSize2 = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantDiagonalSize3"
                    Try
                        Me.base_structure(recIndex).TowerRedundantDiagonalSize3 = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantDiagonalSize4"
                    Try
                        Me.base_structure(recIndex).TowerRedundantDiagonalSize4 = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantSubHorizontalSize"
                    Try
                        Me.base_structure(recIndex).TowerRedundantSubHorizontalSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantSubDiagonalSize"
                    Try
                        Me.base_structure(recIndex).TowerRedundantSubDiagonalSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerSubDiagLocation"
                    Try
                        Me.base_structure(recIndex).TowerSubDiagLocation = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantVerticalSize"
                    Try
                        Me.base_structure(recIndex).TowerRedundantVerticalSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHipSize"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHipSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHipSize2"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHipSize2 = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHipSize3"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHipSize3 = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHipSize4"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHipSize4 = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHipDiagonalSize"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHipDiagonalSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHipDiagonalSize2"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHipDiagonalSize2 = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHipDiagonalSize3"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHipDiagonalSize3 = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHipDiagonalSize4"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHipDiagonalSize4 = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerSWMult"
                    Try
                        Me.base_structure(recIndex).TowerSWMult = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerWPMult"
                    Try
                        Me.base_structure(recIndex).TowerWPMult = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerAutoCalcKSingleAngle"
                    Try
                        Me.base_structure(recIndex).TowerAutoCalcKSingleAngle = CBool(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerAutoCalcKSolidRound"
                    Try
                        Me.base_structure(recIndex).TowerAutoCalcKSolidRound = CBool(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerAfGusset"
                    Try
                        Me.base_structure(recIndex).TowerAfGusset = Me.units.Length.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerTfGusset"
                    Try
                        Me.base_structure(recIndex).TowerTfGusset = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerGussetBoltEdgeDistance"
                    Try
                        Me.base_structure(recIndex).TowerGussetBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerGussetGrade"
                    Try
                        Me.base_structure(recIndex).TowerGussetGrade = Me.units.Strength.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerGussetMatlGrade"
                    Try
                        Me.base_structure(recIndex).TowerGussetMatlGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerAfMult"
                    Try
                        Me.base_structure(recIndex).TowerAfMult = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerArMult"
                    Try
                        Me.base_structure(recIndex).TowerArMult = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerFlatIPAPole"
                    Try
                        Me.base_structure(recIndex).TowerFlatIPAPole = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerRoundIPAPole"
                    Try
                        Me.base_structure(recIndex).TowerRoundIPAPole = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerFlatIPALeg"
                    Try
                        Me.base_structure(recIndex).TowerFlatIPALeg = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerRoundIPALeg"
                    Try
                        Me.base_structure(recIndex).TowerRoundIPALeg = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerFlatIPAHorizontal"
                    Try
                        Me.base_structure(recIndex).TowerFlatIPAHorizontal = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerRoundIPAHorizontal"
                    Try
                        Me.base_structure(recIndex).TowerRoundIPAHorizontal = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerFlatIPADiagonal"
                    Try
                        Me.base_structure(recIndex).TowerFlatIPADiagonal = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerRoundIPADiagonal"
                    Try
                        Me.base_structure(recIndex).TowerRoundIPADiagonal = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerCSA_S37_SpeedUpFactor"
                    Try
                        Me.base_structure(recIndex).TowerCSA_S37_SpeedUpFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerKLegs"
                    Try
                        Me.base_structure(recIndex).TowerKLegs = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerKXBracedDiags"
                    Try
                        Me.base_structure(recIndex).TowerKXBracedDiags = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerKKBracedDiags"
                    Try
                        Me.base_structure(recIndex).TowerKKBracedDiags = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerKZBracedDiags"
                    Try
                        Me.base_structure(recIndex).TowerKZBracedDiags = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerKHorzs"
                    Try
                        Me.base_structure(recIndex).TowerKHorzs = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerKSecHorzs"
                    Try
                        Me.base_structure(recIndex).TowerKSecHorzs = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerKGirts"
                    Try
                        Me.base_structure(recIndex).TowerKGirts = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerKInners"
                    Try
                        Me.base_structure(recIndex).TowerKInners = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerKXBracedDiagsY"
                    Try
                        Me.base_structure(recIndex).TowerKXBracedDiagsY = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerKKBracedDiagsY"
                    Try
                        Me.base_structure(recIndex).TowerKKBracedDiagsY = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerKZBracedDiagsY"
                    Try
                        Me.base_structure(recIndex).TowerKZBracedDiagsY = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerKHorzsY"
                    Try
                        Me.base_structure(recIndex).TowerKHorzsY = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerKSecHorzsY"
                    Try
                        Me.base_structure(recIndex).TowerKSecHorzsY = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerKGirtsY"
                    Try
                        Me.base_structure(recIndex).TowerKGirtsY = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerKInnersY"
                    Try
                        Me.base_structure(recIndex).TowerKInnersY = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerKRedHorz"
                    Try
                        Me.base_structure(recIndex).TowerKRedHorz = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerKRedDiag"
                    Try
                        Me.base_structure(recIndex).TowerKRedDiag = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerKRedSubDiag"
                    Try
                        Me.base_structure(recIndex).TowerKRedSubDiag = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerKRedSubHorz"
                    Try
                        Me.base_structure(recIndex).TowerKRedSubHorz = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerKRedVert"
                    Try
                        Me.base_structure(recIndex).TowerKRedVert = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerKRedHip"
                    Try
                        Me.base_structure(recIndex).TowerKRedHip = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerKRedHipDiag"
                    Try
                        Me.base_structure(recIndex).TowerKRedHipDiag = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerKTLX"
                    Try
                        Me.base_structure(recIndex).TowerKTLX = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerKTLZ"
                    Try
                        Me.base_structure(recIndex).TowerKTLZ = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerKTLLeg"
                    Try
                        Me.base_structure(recIndex).TowerKTLLeg = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerInnerKTLX"
                    Try
                        Me.base_structure(recIndex).TowerInnerKTLX = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerInnerKTLZ"
                    Try
                        Me.base_structure(recIndex).TowerInnerKTLZ = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerInnerKTLLeg"
                    Try
                        Me.base_structure(recIndex).TowerInnerKTLLeg = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerStitchBoltLocationHoriz"
                    Try
                        Me.base_structure(recIndex).TowerStitchBoltLocationHoriz = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerStitchBoltLocationDiag"
                    Try
                        Me.base_structure(recIndex).TowerStitchBoltLocationDiag = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerStitchBoltLocationRed"
                    Try
                        Me.base_structure(recIndex).TowerStitchBoltLocationRed = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerStitchSpacing"
                    Try
                        Me.base_structure(recIndex).TowerStitchSpacing = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerStitchSpacingDiag"
                    Try
                        Me.base_structure(recIndex).TowerStitchSpacingDiag = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerStitchSpacingHorz"
                    Try
                        Me.base_structure(recIndex).TowerStitchSpacingHorz = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerStitchSpacingRed"
                    Try
                        Me.base_structure(recIndex).TowerStitchSpacingRed = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerLegNetWidthDeduct"
                    Try
                        Me.base_structure(recIndex).TowerLegNetWidthDeduct = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerLegUFactor"
                    Try
                        Me.base_structure(recIndex).TowerLegUFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerDiagonalNetWidthDeduct"
                    Try
                        Me.base_structure(recIndex).TowerDiagonalNetWidthDeduct = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerTopGirtNetWidthDeduct"
                    Try
                        Me.base_structure(recIndex).TowerTopGirtNetWidthDeduct = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerBotGirtNetWidthDeduct"
                    Try
                        Me.base_structure(recIndex).TowerBotGirtNetWidthDeduct = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerInnerGirtNetWidthDeduct"
                    Try
                        Me.base_structure(recIndex).TowerInnerGirtNetWidthDeduct = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerHorizontalNetWidthDeduct"
                    Try
                        Me.base_structure(recIndex).TowerHorizontalNetWidthDeduct = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerShortHorizontalNetWidthDeduct"
                    Try
                        Me.base_structure(recIndex).TowerShortHorizontalNetWidthDeduct = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerDiagonalUFactor"
                    Try
                        Me.base_structure(recIndex).TowerDiagonalUFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerTopGirtUFactor"
                    Try
                        Me.base_structure(recIndex).TowerTopGirtUFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerBotGirtUFactor"
                    Try
                        Me.base_structure(recIndex).TowerBotGirtUFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerInnerGirtUFactor"
                    Try
                        Me.base_structure(recIndex).TowerInnerGirtUFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerHorizontalUFactor"
                    Try
                        Me.base_structure(recIndex).TowerHorizontalUFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerShortHorizontalUFactor"
                    Try
                        Me.base_structure(recIndex).TowerShortHorizontalUFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerLegConnType"
                    Try
                        Me.base_structure(recIndex).TowerLegConnType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerLegNumBolts"
                    Try
                        Me.base_structure(recIndex).TowerLegNumBolts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerDiagonalNumBolts"
                    Try
                        Me.base_structure(recIndex).TowerDiagonalNumBolts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerTopGirtNumBolts"
                    Try
                        Me.base_structure(recIndex).TowerTopGirtNumBolts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerBotGirtNumBolts"
                    Try
                        Me.base_structure(recIndex).TowerBotGirtNumBolts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerInnerGirtNumBolts"
                    Try
                        Me.base_structure(recIndex).TowerInnerGirtNumBolts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerHorizontalNumBolts"
                    Try
                        Me.base_structure(recIndex).TowerHorizontalNumBolts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerShortHorizontalNumBolts"
                    Try
                        Me.base_structure(recIndex).TowerShortHorizontalNumBolts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerLegBoltGrade"
                    Try
                        Me.base_structure(recIndex).TowerLegBoltGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerLegBoltSize"
                    Try
                        Me.base_structure(recIndex).TowerLegBoltSize = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerDiagonalBoltGrade"
                    Try
                        Me.base_structure(recIndex).TowerDiagonalBoltGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerDiagonalBoltSize"
                    Try
                        Me.base_structure(recIndex).TowerDiagonalBoltSize = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerTopGirtBoltGrade"
                    Try
                        Me.base_structure(recIndex).TowerTopGirtBoltGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerTopGirtBoltSize"
                    Try
                        Me.base_structure(recIndex).TowerTopGirtBoltSize = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerBotGirtBoltGrade"
                    Try
                        Me.base_structure(recIndex).TowerBotGirtBoltGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerBotGirtBoltSize"
                    Try
                        Me.base_structure(recIndex).TowerBotGirtBoltSize = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerInnerGirtBoltGrade"
                    Try
                        Me.base_structure(recIndex).TowerInnerGirtBoltGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerInnerGirtBoltSize"
                    Try
                        Me.base_structure(recIndex).TowerInnerGirtBoltSize = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerHorizontalBoltGrade"
                    Try
                        Me.base_structure(recIndex).TowerHorizontalBoltGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerHorizontalBoltSize"
                    Try
                        Me.base_structure(recIndex).TowerHorizontalBoltSize = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerShortHorizontalBoltGrade"
                    Try
                        Me.base_structure(recIndex).TowerShortHorizontalBoltGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerShortHorizontalBoltSize"
                    Try
                        Me.base_structure(recIndex).TowerShortHorizontalBoltSize = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerLegBoltEdgeDistance"
                    Try
                        Me.base_structure(recIndex).TowerLegBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerDiagonalBoltEdgeDistance"
                    Try
                        Me.base_structure(recIndex).TowerDiagonalBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerTopGirtBoltEdgeDistance"
                    Try
                        Me.base_structure(recIndex).TowerTopGirtBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerBotGirtBoltEdgeDistance"
                    Try
                        Me.base_structure(recIndex).TowerBotGirtBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerInnerGirtBoltEdgeDistance"
                    Try
                        Me.base_structure(recIndex).TowerInnerGirtBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerHorizontalBoltEdgeDistance"
                    Try
                        Me.base_structure(recIndex).TowerHorizontalBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerShortHorizontalBoltEdgeDistance"
                    Try
                        Me.base_structure(recIndex).TowerShortHorizontalBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerDiagonalGageG1Distance"
                    Try
                        Me.base_structure(recIndex).TowerDiagonalGageG1Distance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerTopGirtGageG1Distance"
                    Try
                        Me.base_structure(recIndex).TowerTopGirtGageG1Distance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerBotGirtGageG1Distance"
                    Try
                        Me.base_structure(recIndex).TowerBotGirtGageG1Distance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerInnerGirtGageG1Distance"
                    Try
                        Me.base_structure(recIndex).TowerInnerGirtGageG1Distance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerHorizontalGageG1Distance"
                    Try
                        Me.base_structure(recIndex).TowerHorizontalGageG1Distance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerShortHorizontalGageG1Distance"
                    Try
                        Me.base_structure(recIndex).TowerShortHorizontalGageG1Distance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHorizontalBoltGrade"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHorizontalBoltGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHorizontalBoltSize"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHorizontalBoltSize = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHorizontalNumBolts"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHorizontalNumBolts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHorizontalBoltEdgeDistance"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHorizontalBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHorizontalGageG1Distance"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHorizontalGageG1Distance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHorizontalNetWidthDeduct"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHorizontalNetWidthDeduct = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHorizontalUFactor"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHorizontalUFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantDiagonalBoltGrade"
                    Try
                        Me.base_structure(recIndex).TowerRedundantDiagonalBoltGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantDiagonalBoltSize"
                    Try
                        Me.base_structure(recIndex).TowerRedundantDiagonalBoltSize = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantDiagonalNumBolts"
                    Try
                        Me.base_structure(recIndex).TowerRedundantDiagonalNumBolts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantDiagonalBoltEdgeDistance"
                    Try
                        Me.base_structure(recIndex).TowerRedundantDiagonalBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantDiagonalGageG1Distance"
                    Try
                        Me.base_structure(recIndex).TowerRedundantDiagonalGageG1Distance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantDiagonalNetWidthDeduct"
                    Try
                        Me.base_structure(recIndex).TowerRedundantDiagonalNetWidthDeduct = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantDiagonalUFactor"
                    Try
                        Me.base_structure(recIndex).TowerRedundantDiagonalUFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantSubDiagonalBoltGrade"
                    Try
                        Me.base_structure(recIndex).TowerRedundantSubDiagonalBoltGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantSubDiagonalBoltSize"
                    Try
                        Me.base_structure(recIndex).TowerRedundantSubDiagonalBoltSize = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantSubDiagonalNumBolts"
                    Try
                        Me.base_structure(recIndex).TowerRedundantSubDiagonalNumBolts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantSubDiagonalBoltEdgeDistance"
                    Try
                        Me.base_structure(recIndex).TowerRedundantSubDiagonalBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantSubDiagonalGageG1Distance"
                    Try
                        Me.base_structure(recIndex).TowerRedundantSubDiagonalGageG1Distance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantSubDiagonalNetWidthDeduct"
                    Try
                        Me.base_structure(recIndex).TowerRedundantSubDiagonalNetWidthDeduct = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantSubDiagonalUFactor"
                    Try
                        Me.base_structure(recIndex).TowerRedundantSubDiagonalUFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantSubHorizontalBoltGrade"
                    Try
                        Me.base_structure(recIndex).TowerRedundantSubHorizontalBoltGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantSubHorizontalBoltSize"
                    Try
                        Me.base_structure(recIndex).TowerRedundantSubHorizontalBoltSize = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantSubHorizontalNumBolts"
                    Try
                        Me.base_structure(recIndex).TowerRedundantSubHorizontalNumBolts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantSubHorizontalBoltEdgeDistance"
                    Try
                        Me.base_structure(recIndex).TowerRedundantSubHorizontalBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantSubHorizontalGageG1Distance"
                    Try
                        Me.base_structure(recIndex).TowerRedundantSubHorizontalGageG1Distance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantSubHorizontalNetWidthDeduct"
                    Try
                        Me.base_structure(recIndex).TowerRedundantSubHorizontalNetWidthDeduct = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantSubHorizontalUFactor"
                    Try
                        Me.base_structure(recIndex).TowerRedundantSubHorizontalUFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantVerticalBoltGrade"
                    Try
                        Me.base_structure(recIndex).TowerRedundantVerticalBoltGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantVerticalBoltSize"
                    Try
                        Me.base_structure(recIndex).TowerRedundantVerticalBoltSize = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantVerticalNumBolts"
                    Try
                        Me.base_structure(recIndex).TowerRedundantVerticalNumBolts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantVerticalBoltEdgeDistance"
                    Try
                        Me.base_structure(recIndex).TowerRedundantVerticalBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantVerticalGageG1Distance"
                    Try
                        Me.base_structure(recIndex).TowerRedundantVerticalGageG1Distance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantVerticalNetWidthDeduct"
                    Try
                        Me.base_structure(recIndex).TowerRedundantVerticalNetWidthDeduct = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantVerticalUFactor"
                    Try
                        Me.base_structure(recIndex).TowerRedundantVerticalUFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHipBoltGrade"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHipBoltGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHipBoltSize"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHipBoltSize = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHipNumBolts"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHipNumBolts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHipBoltEdgeDistance"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHipBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHipGageG1Distance"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHipGageG1Distance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHipNetWidthDeduct"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHipNetWidthDeduct = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHipUFactor"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHipUFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHipDiagonalBoltGrade"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHipDiagonalBoltGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHipDiagonalBoltSize"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHipDiagonalBoltSize = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHipDiagonalNumBolts"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHipDiagonalNumBolts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHipDiagonalBoltEdgeDistance"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHipDiagonalBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHipDiagonalGageG1Distance"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHipDiagonalGageG1Distance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHipDiagonalNetWidthDeduct"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHipDiagonalNetWidthDeduct = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerRedundantHipDiagonalUFactor"
                    Try
                        Me.base_structure(recIndex).TowerRedundantHipDiagonalUFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerDiagonalOutOfPlaneRestraint"
                    Try
                        Me.base_structure(recIndex).TowerDiagonalOutOfPlaneRestraint = CBool(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerTopGirtOutOfPlaneRestraint"
                    Try
                        Me.base_structure(recIndex).TowerTopGirtOutOfPlaneRestraint = CBool(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerBottomGirtOutOfPlaneRestraint"
                    Try
                        Me.base_structure(recIndex).TowerBottomGirtOutOfPlaneRestraint = CBool(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerMidGirtOutOfPlaneRestraint"
                    Try
                        Me.base_structure(recIndex).TowerMidGirtOutOfPlaneRestraint = CBool(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerHorizontalOutOfPlaneRestraint"
                    Try
                        Me.base_structure(recIndex).TowerHorizontalOutOfPlaneRestraint = CBool(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerSecondaryHorizontalOutOfPlaneRestraint"
                    Try
                        Me.base_structure(recIndex).TowerSecondaryHorizontalOutOfPlaneRestraint = CBool(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerUniqueFlag"
                    Try
                        Me.base_structure(recIndex).TowerUniqueFlag = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TowerDiagOffsetNEY"
                    Try
                        Me.base_structure(recIndex).TowerDiagOffsetNEY = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerDiagOffsetNEX"
                    Try
                        Me.base_structure(recIndex).TowerDiagOffsetNEX = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerDiagOffsetPEY"
                    Try
                        Me.base_structure(recIndex).TowerDiagOffsetPEY = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerDiagOffsetPEX"
                    Try
                        Me.base_structure(recIndex).TowerDiagOffsetPEX = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerKbraceOffsetNEY"
                    Try
                        Me.base_structure(recIndex).TowerKbraceOffsetNEY = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerKbraceOffsetNEX"
                    Try
                        Me.base_structure(recIndex).TowerKbraceOffsetNEX = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerKbraceOffsetPEY"
                    Try
                        Me.base_structure(recIndex).TowerKbraceOffsetPEY = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TowerKbraceOffsetPEX"
                    Try
                        Me.base_structure(recIndex).TowerKbraceOffsetPEX = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                ''''Guy Rec''''
                Case "NumGuyRecs"
                    Try
                        If CInt(tnxValue) > 0 Then
                            Me.guy_wires = New List(Of tnx_guy_record)
                        End If
                    Catch ex As Exception
                    End Try
                Case "GuyRec"
                    Try
                        recIndex = CInt(tnxValue) - 1
                        Me.guy_wires.Add(New tnx_guy_record())
                        Me.guy_wires(recIndex).GuyRec = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "GuyHeight"
                    Try
                        Me.guy_wires(recIndex).GuyHeight = Me.units.Coordinate.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "GuyAutoCalcKSingleAngle"
                    Try
                        Me.guy_wires(recIndex).GuyAutoCalcKSingleAngle = CBool(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "GuyAutoCalcKSolidRound"
                    Try
                        Me.guy_wires(recIndex).GuyAutoCalcKSolidRound = CBool(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "GuyMount"
                    Try
                        Me.guy_wires(recIndex).GuyMount = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TorqueArmStyle"
                    Try
                        Me.guy_wires(recIndex).TorqueArmStyle = tnxValue
                    Catch ex As Exception
                    End Try
                Case "GuyRadius"
                    Try
                        Me.guy_wires(recIndex).GuyRadius = Me.units.Length.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "GuyRadius120"
                    Try
                        Me.guy_wires(recIndex).GuyRadius120 = Me.units.Length.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "GuyRadius240"
                    Try
                        Me.guy_wires(recIndex).GuyRadius240 = Me.units.Length.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "GuyRadius360"
                    Try
                        Me.guy_wires(recIndex).GuyRadius360 = Me.units.Length.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TorqueArmRadius"
                    Try
                        Me.guy_wires(recIndex).TorqueArmRadius = Me.units.Length.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TorqueArmLegAngle"
                    Try
                        Me.guy_wires(recIndex).TorqueArmLegAngle = Me.units.Rotation.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "Azimuth0Adjustment"
                    Try
                        Me.guy_wires(recIndex).Azimuth0Adjustment = Me.units.Rotation.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "Azimuth120Adjustment"
                    Try
                        Me.guy_wires(recIndex).Azimuth120Adjustment = Me.units.Rotation.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "Azimuth240Adjustment"
                    Try
                        Me.guy_wires(recIndex).Azimuth240Adjustment = Me.units.Rotation.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "Azimuth360Adjustment"
                    Try
                        Me.guy_wires(recIndex).Azimuth360Adjustment = Me.units.Rotation.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "Anchor0Elevation"
                    Try
                        Me.guy_wires(recIndex).Anchor0Elevation = Me.units.Coordinate.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "Anchor120Elevation"
                    Try
                        Me.guy_wires(recIndex).Anchor120Elevation = Me.units.Coordinate.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "Anchor240Elevation"
                    Try
                        Me.guy_wires(recIndex).Anchor240Elevation = Me.units.Coordinate.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "Anchor360Elevation"
                    Try
                        Me.guy_wires(recIndex).Anchor360Elevation = Me.units.Coordinate.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "GuySize"
                    Try
                        Me.guy_wires(recIndex).GuySize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "Guy120Size"
                    Try
                        Me.guy_wires(recIndex).Guy120Size = tnxValue
                    Catch ex As Exception
                    End Try
                Case "Guy240Size"
                    Try
                        Me.guy_wires(recIndex).Guy240Size = tnxValue
                    Catch ex As Exception
                    End Try
                Case "Guy360Size"
                    Try
                        Me.guy_wires(recIndex).Guy360Size = tnxValue
                    Catch ex As Exception
                    End Try
                Case "GuyGrade"
                    Try
                        Me.guy_wires(recIndex).GuyGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TorqueArmSize"
                    Try
                        Me.guy_wires(recIndex).TorqueArmSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TorqueArmSizeBot"
                    Try
                        Me.guy_wires(recIndex).TorqueArmSizeBot = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TorqueArmType"
                    Try
                        Me.guy_wires(recIndex).TorqueArmType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TorqueArmGrade"
                    Try
                        Me.guy_wires(recIndex).TorqueArmGrade = Me.units.Strength.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "TorqueArmMatlGrade"
                    Try
                        Me.guy_wires(recIndex).TorqueArmMatlGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "TorqueArmKFactor"
                    Try
                        Me.guy_wires(recIndex).TorqueArmKFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "TorqueArmKFactorY"
                    Try
                        Me.guy_wires(recIndex).TorqueArmKFactorY = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "GuyPullOffKFactorX"
                    Try
                        Me.guy_wires(recIndex).GuyPullOffKFactorX = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "GuyPullOffKFactorY"
                    Try
                        Me.guy_wires(recIndex).GuyPullOffKFactorY = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "GuyDiagKFactorX"
                    Try
                        Me.guy_wires(recIndex).GuyDiagKFactorX = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "GuyDiagKFactorY"
                    Try
                        Me.guy_wires(recIndex).GuyDiagKFactorY = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "GuyAutoCalc"
                    Try
                        Me.guy_wires(recIndex).GuyAutoCalc = CBool(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "GuyAllGuysSame"
                    Try
                        Me.guy_wires(recIndex).GuyAllGuysSame = CBool(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "GuyAllGuysAnchorSame"
                    Try
                        Me.guy_wires(recIndex).GuyAllGuysAnchorSame = CBool(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "GuyIsStrapping"
                    Try
                        Me.guy_wires(recIndex).GuyIsStrapping = CBool(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "GuyPullOffSize"
                    Try
                        Me.guy_wires(recIndex).GuyPullOffSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "GuyPullOffSizeBot"
                    Try
                        Me.guy_wires(recIndex).GuyPullOffSizeBot = tnxValue
                    Catch ex As Exception
                    End Try
                Case "GuyPullOffType"
                    Try
                        Me.guy_wires(recIndex).GuyPullOffType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "GuyPullOffGrade"
                    Try
                        Me.guy_wires(recIndex).GuyPullOffGrade = Me.units.Strength.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "GuyPullOffMatlGrade"
                    Try
                        Me.guy_wires(recIndex).GuyPullOffMatlGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "GuyUpperDiagSize"
                    Try
                        Me.guy_wires(recIndex).GuyUpperDiagSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "GuyLowerDiagSize"
                    Try
                        Me.guy_wires(recIndex).GuyLowerDiagSize = tnxValue
                    Catch ex As Exception
                    End Try
                Case "GuyDiagType"
                    Try
                        Me.guy_wires(recIndex).GuyDiagType = tnxValue
                    Catch ex As Exception
                    End Try
                Case "GuyDiagGrade"
                    Try
                        Me.guy_wires(recIndex).GuyDiagGrade = Me.units.Strength.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "GuyDiagMatlGrade"
                    Try
                        Me.guy_wires(recIndex).GuyDiagMatlGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "GuyDiagNetWidthDeduct"
                    Try
                        Me.guy_wires(recIndex).GuyDiagNetWidthDeduct = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "GuyDiagUFactor"
                    Try
                        Me.guy_wires(recIndex).GuyDiagUFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "GuyDiagNumBolts"
                    Try
                        Me.guy_wires(recIndex).GuyDiagNumBolts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "GuyDiagonalOutOfPlaneRestraint"
                    Try
                        Me.guy_wires(recIndex).GuyDiagonalOutOfPlaneRestraint = CBool(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "GuyDiagBoltGrade"
                    Try
                        Me.guy_wires(recIndex).GuyDiagBoltGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "GuyDiagBoltSize"
                    Try
                        Me.guy_wires(recIndex).GuyDiagBoltSize = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "GuyDiagBoltEdgeDistance"
                    Try
                        Me.guy_wires(recIndex).GuyDiagBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "GuyDiagBoltGageDistance"
                    Try
                        Me.guy_wires(recIndex).GuyDiagBoltGageDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "GuyPullOffNetWidthDeduct"
                    Try
                        Me.guy_wires(recIndex).GuyPullOffNetWidthDeduct = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "GuyPullOffUFactor"
                    Try
                        Me.guy_wires(recIndex).GuyPullOffUFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "GuyPullOffNumBolts"
                    Try
                        Me.guy_wires(recIndex).GuyPullOffNumBolts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "GuyPullOffOutOfPlaneRestraint"
                    Try
                        Me.guy_wires(recIndex).GuyPullOffOutOfPlaneRestraint = CBool(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "GuyPullOffBoltGrade"
                    Try
                        Me.guy_wires(recIndex).GuyPullOffBoltGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "GuyPullOffBoltSize"
                    Try
                        Me.guy_wires(recIndex).GuyPullOffBoltSize = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "GuyPullOffBoltEdgeDistance"
                    Try
                        Me.guy_wires(recIndex).GuyPullOffBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "GuyPullOffBoltGageDistance"
                    Try
                        Me.guy_wires(recIndex).GuyPullOffBoltGageDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "GuyTorqueArmNetWidthDeduct"
                    Try
                        Me.guy_wires(recIndex).GuyTorqueArmNetWidthDeduct = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "GuyTorqueArmUFactor"
                    Try
                        Me.guy_wires(recIndex).GuyTorqueArmUFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "GuyTorqueArmNumBolts"
                    Try
                        Me.guy_wires(recIndex).GuyTorqueArmNumBolts = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "GuyTorqueArmOutOfPlaneRestraint"
                    Try
                        Me.guy_wires(recIndex).GuyTorqueArmOutOfPlaneRestraint = CBool(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "GuyTorqueArmBoltGrade"
                    Try
                        Me.guy_wires(recIndex).GuyTorqueArmBoltGrade = tnxValue
                    Catch ex As Exception
                    End Try
                Case "GuyTorqueArmBoltSize"
                    Try
                        Me.guy_wires(recIndex).GuyTorqueArmBoltSize = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "GuyTorqueArmBoltEdgeDistance"
                    Try
                        Me.guy_wires(recIndex).GuyTorqueArmBoltEdgeDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "GuyTorqueArmBoltGageDistance"
                    Try
                        Me.guy_wires(recIndex).GuyTorqueArmBoltGageDistance = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "GuyPerCentTension"
                    Try
                        Me.guy_wires(recIndex).GuyPerCentTension = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "GuyPerCentTension120"
                    Try
                        Me.guy_wires(recIndex).GuyPerCentTension120 = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "GuyPerCentTension240"
                    Try
                        Me.guy_wires(recIndex).GuyPerCentTension240 = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "GuyPerCentTension360"
                    Try
                        Me.guy_wires(recIndex).GuyPerCentTension360 = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "GuyEffFactor"
                    Try
                        Me.guy_wires(recIndex).GuyEffFactor = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "GuyEffFactor120"
                    Try
                        Me.guy_wires(recIndex).GuyEffFactor120 = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "GuyEffFactor240"
                    Try
                        Me.guy_wires(recIndex).GuyEffFactor240 = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "GuyEffFactor360"
                    Try
                        Me.guy_wires(recIndex).GuyEffFactor360 = CDbl(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "GuyNumInsulators"
                    Try
                        Me.guy_wires(recIndex).GuyNumInsulators = CInt(tnxValue)
                    Catch ex As Exception
                    End Try
                Case "GuyInsulatorLength"
                    Try
                        Me.guy_wires(recIndex).GuyInsulatorLength = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "GuyInsulatorDia"
                    Try
                        Me.guy_wires(recIndex).GuyInsulatorDia = Me.units.Properties.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try
                Case "GuyInsulatorWt"
                    Try
                        Me.guy_wires(recIndex).GuyInsulatorWt = Me.units.Force.Convert_to_EDS_Default(CDbl(tnxValue))
                    Catch ex As Exception
                    End Try


                    ''''End of the input data''''
                Case "[End Application]"
                    Exit For
            End Select

        Next

    End Sub
End Class

Partial Public Class tnx_antenna_record
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

Partial Public Class tnx_tower_record
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

Partial Public Class tnx_guy_record
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

#Region "Units"
Partial Public Class tnx_units

    Private propLength As tnx_length_unit
    Private propCoordinate As tnx_coordinate_unit
    Private propForce As tnx_force_unit
    Private propLoad As tnx_load_unit
    Private propMoment As tnx_moment_unit
    Private propProperties As tnx_properties_unit
    Private propPressure As tnx_pressure_unit
    Private propVelocity As tnx_velocity_unit
    Private propDisplacement As tnx_displacement_unit
    Private propMass As tnx_mass_unit
    Private propAcceleration As tnx_acceleration_unit
    Private propStress As tnx_stress_unit
    Private propDensity As tnx_density_unit
    Private propUnitWt As tnx_unitwt_unit
    Private propStrength As tnx_strength_unit
    Private propModulus As tnx_modulus_unit
    Private propTemperature As tnx_temperature_unit
    Private propPrinter As tnx_printer_unit
    Private propRotation As tnx_rotation_unit
    Private propSpacing As tnx_spacing_unit

    <Category("TNX Units"), Description(""), DisplayName("Length")>
    Public Property Length() As tnx_length_unit
        Get
            Return Me.propLength
        End Get
        Set
            Me.propLength = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Coordinate")>
    Public Property Coordinate() As tnx_coordinate_unit
        Get
            Return Me.propCoordinate
        End Get
        Set
            Me.propCoordinate = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Force")>
    Public Property Force() As tnx_force_unit
        Get
            Return Me.propForce
        End Get
        Set
            Me.propForce = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Load")>
    Public Property Load() As tnx_load_unit
        Get
            Return Me.propLoad
        End Get
        Set
            Me.propLoad = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Moment")>
    Public Property Moment() As tnx_moment_unit
        Get
            Return Me.propMoment
        End Get
        Set
            Me.propMoment = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Properties")>
    Public Property Properties() As tnx_properties_unit
        Get
            Return Me.propProperties
        End Get
        Set
            Me.propProperties = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Pressure")>
    Public Property Pressure() As tnx_pressure_unit
        Get
            Return Me.propPressure
        End Get
        Set
            Me.propPressure = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Velocity")>
    Public Property Velocity() As tnx_velocity_unit
        Get
            Return Me.propVelocity
        End Get
        Set
            Me.propVelocity = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Displacement")>
    Public Property Displacement() As tnx_displacement_unit
        Get
            Return Me.propDisplacement
        End Get
        Set
            Me.propDisplacement = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Mass")>
    Public Property Mass() As tnx_mass_unit
        Get
            Return Me.propMass
        End Get
        Set
            Me.propMass = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Acceleration")>
    Public Property Acceleration() As tnx_acceleration_unit
        Get
            Return Me.propAcceleration
        End Get
        Set
            Me.propAcceleration = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Stress")>
    Public Property Stress() As tnx_stress_unit
        Get
            Return Me.propStress
        End Get
        Set
            Me.propStress = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Density")>
    Public Property Density() As tnx_density_unit
        Get
            Return Me.propDensity
        End Get
        Set
            Me.propDensity = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Unitwt")>
    Public Property UnitWt() As tnx_unitwt_unit
        Get
            Return Me.propUnitWt
        End Get
        Set
            Me.propUnitWt = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Strength")>
    Public Property Strength() As tnx_strength_unit
        Get
            Return Me.propStrength
        End Get
        Set
            Me.propStrength = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Modulus")>
    Public Property Modulus() As tnx_modulus_unit
        Get
            Return Me.propModulus
        End Get
        Set
            Me.propModulus = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Temperature")>
    Public Property Temperature() As tnx_temperature_unit
        Get
            Return Me.propTemperature
        End Get
        Set
            Me.propTemperature = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Printer")>
    Public Property Printer() As tnx_printer_unit
        Get
            Return Me.propPrinter
        End Get
        Set
            Me.propPrinter = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Rotation")>
    Public Property Rotation() As tnx_rotation_unit
        Get
            Return Me.propRotation
        End Get
        Set
            Me.propRotation = Value
        End Set
    End Property
    <Category("TNX Units"), Description(""), DisplayName("Spacing")>
    Public Property Spacing() As tnx_spacing_unit
        Get
            Return Me.propSpacing
        End Get
        Set
            Me.propSpacing = Value
        End Set
    End Property

End Class

Partial Public Class tnx_Unit_Property
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
    Public Overridable Function Convert_to_EDS_Default(InputValue As Double) As Double

        If Me.prop_value = "" Then
            Throw New System.Exception("Property value not set")
        ElseIf Me.prop_multiplier = 0 Then
            Throw New System.Exception("Property multiplier not set")
        End If

        Return InputValue / Me.multiplier

    End Function

End Class

Partial Public Class tnx_length_unit
    Inherits tnx_Unit_Property

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

    Public Sub New()
    End Sub
    Public Sub New(new_value As String)

        Me.value = new_value

    End Sub

End Class

Partial Public Class tnx_coordinate_unit
    Inherits tnx_length_unit
    Public Sub New()
    End Sub
    Public Sub New(new_value As String)

        Me.value = new_value

    End Sub
End Class

Partial Public Class tnx_force_unit
    Inherits tnx_Unit_Property

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
Partial Public Class tnx_load_unit
    Inherits tnx_Unit_Property

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
Partial Public Class tnx_moment_unit
    Inherits tnx_Unit_Property

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
Partial Public Class tnx_properties_unit
    Inherits tnx_length_unit
    Public Sub New()
    End Sub
    Public Sub New(new_value As String)

        Me.value = new_value

    End Sub
End Class
Partial Public Class tnx_pressure_unit
    Inherits tnx_Unit_Property

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
Partial Public Class tnx_velocity_unit
    Inherits tnx_Unit_Property

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
Partial Public Class tnx_displacement_unit
    'Note: This is called deflection in the TNX UI
    Inherits tnx_length_unit
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
Partial Public Class tnx_mass_unit
    'This property isn't accessible in the TNX UI
    Inherits tnx_Unit_Property

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
Partial Public Class tnx_acceleration_unit
    Inherits tnx_Unit_Property

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
Partial Public Class tnx_stress_unit
    Inherits tnx_Unit_Property

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
Partial Public Class tnx_density_unit
    Inherits tnx_Unit_Property

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
Partial Public Class tnx_unitwt_unit
    Inherits tnx_Unit_Property
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
Partial Public Class tnx_strength_unit
    Inherits tnx_Unit_Property

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
Partial Public Class tnx_modulus_unit
    Inherits tnx_Unit_Property

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
Partial Public Class tnx_temperature_unit
    Inherits tnx_Unit_Property

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

    Public Overrides Function Convert_to_EDS_Default(InputValue As Double) As Double

        If Me.prop_value = "" Then
            Throw New System.Exception("Property value not set")
        End If

        If Me.prop_value = "C" Then
            Return InputValue * (9 / 5) + 32
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
Partial Public Class tnx_printer_unit
    'This property isn't accessible in the TNX UI
    Inherits tnx_Unit_Property

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
Partial Public Class tnx_rotation_unit
    Inherits tnx_Unit_Property

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
Partial Public Class tnx_spacing_unit
    Inherits tnx_length_unit

    Public Sub New()
    End Sub
    Public Sub New(new_value As String)

        Me.value = new_value

    End Sub
End Class

#End Region

