Imports System.ComponentModel
Imports DevExpress.Spreadsheet
Imports System.Runtime.Serialization

<DataContractAttribute()>
<TypeConverterAttribute(GetType(ExpandableObjectConverter))>
<KnownType(GetType(DrilledPierFoundation))>
Partial Public Class DrilledPierFoundation
    Inherits EDSExcelObject

    <DataMember()> Public Property DrilledPiers As New List(Of DrilledPier)

    ' <DataMember()> Public Property ParentFile As New FileUpload

    'Origin row in the driled pier database. Basically just where the profile numbers are in the database worksheet.
    'This is actually 58 but due to the 0,0 origin in excel, it is 1 less
    Private pierProfileRow As Integer = 57

#Region "Inheritted"
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Drilled Pier Foundation"
        End Get
    End Property
    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "fnd.drilled_pier_tool"
        End Get
    End Property

    Public Overrides ReadOnly Property TemplatePath As String
        Get
            Return IO.Path.Combine(My.Application.Info.DirectoryPath, "Templates", "Drilled Pier Foundation.xlsm")
        End Get
    End Property

    Public Overrides ReadOnly Property Template As Byte() = CCI_Engineering_Templates.My.Resources.Drilled_Pier_Foundation

    Private _Insert As String
    Private _Update As String
    Private _Delete As String

    Public Overrides Function SQLInsert() As String
        SQLInsert = ""
        SQLInsert = CCI_Engineering_Templates.My.Resources.Drilled_Pier_Tool__INSERT  ' QueryBuilderFromFile(queryPath & "Drilled Pier\Drilled Pier Tool (INSERT).sql")
        SQLInsert = SQLInsert.Replace("[TABLE NAME]", Me.EDSTableName)
        SQLInsert = SQLInsert.Replace("[INSERT FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.Replace("[INSERT VALUES]", Me.SQLInsertValues)

        SQLInsert = SQLInsert.Replace("[SUBLEVEL ID]", "@SubLevel4ID")
        SQLInsert = SQLInsert.Replace("[SUBLEVEL TABLE]", "@SubLevel4")

        Dim _dpInsert As String
        For Each dp In Me.DrilledPiers
            _dpInsert += dp.SQLInsert + vbCrLf
        Next

        SQLInsert = SQLInsert.Replace("--[REQUIRED CHILDREN]", _dpInsert)
        SQLInsert = SQLInsert.TrimEnd

        Return SQLInsert + vbCrLf
    End Function

    Public Overrides Function SQLUpdate() As String
        'This section not only needs to call update commands but also needs to call insert and delete commands since subtables may involve adding or deleting records
        SQLUpdate = CCI_Engineering_Templates.My.Resources.General__UPDATE ' queryPath(queryPath & "Drilled Pier\General (UPDATE).sql")
        SQLUpdate = SQLUpdate.Replace("[TABLE NAME]", Me.EDSTableName)
        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)

        Dim dpUp As String = ""
        For Each dp In Me.DrilledPiers
            If dp.ID IsNot Nothing And dp?.ID > 0 Then
                If dp.local_drilled_pier_id IsNot Nothing Then
                    dpUp += dp.SQLUpdate
                Else
                    dpUp += dp.SQLDelete
                End If
            Else
                dpUp += dp.SQLInsert
            End If
        Next

        Return dpUp
    End Function

    Public Overrides Function SQLDelete() As String
        SQLDelete = CCI_Engineering_Templates.My.Resources.General__DELETE
        SQLDelete = SQLDelete.Replace("[TABLE]", Me.EDSTableName)
        SQLDelete = SQLDelete.Replace("[ID]", Me.ID)

        Return SQLDelete & vbCrLf
    End Function

    Public Overrides ReadOnly Property ExcelDTParams As List(Of EXCELDTParameter)
        Get
            Dim dp As New DrilledPier
            Dim dpProf As New DrilledPierProfile
            Dim dpSec As New DrilledPierSection
            Dim dpReb As New DrilledPierRebar
            Dim dpSProf As New DrilledPierSoilProfile
            Dim dpSlay As New DrilledPierSoilLayer
            Dim dpEmbed As New EmbeddedPole
            Dim dpBell As New BelledPier
            Dim dpRes As New DrilledPierResult
            Dim dpTool As New DrilledPierFoundation

            Return New List(Of EXCELDTParameter) From {
                                                        New EXCELDTParameter(dp.EDSObjectName, "A2:I52", "Profiles"),  'It is slightly confusing but to keep naming issues consistent in the tool a drilled pier = profile and a drilled pier profile = drilled pier details
                                                        New EXCELDTParameter(dpProf.EDSObjectName, "A2:X52", "Details"),
                                                        New EXCELDTParameter(dpSec.EDSObjectName, "A2:M252", "Section"),
                                                        New EXCELDTParameter(dpReb.EDSObjectName, "A2:K702", "Rebar"),
                                                        New EXCELDTParameter(dpSProf.EDSObjectName, "A2:F52", "Soil Profile"),
                                                        New EXCELDTParameter(dpSlay.EDSObjectName, "A2:N1502", "Soil Layer"),
                                                        New EXCELDTParameter(dpEmbed.EDSObjectName, "A2:Q52", "Embedded"),
                                                        New EXCELDTParameter(dpRes.EDSObjectName, "BC8:CF58", "Foundation Input"),
                                                        New EXCELDTParameter(dpBell.EDSObjectName, "A2:P52", "Belled"),
                                                        New EXCELDTParameter(dpTool.EDSObjectName, "A2:E3", "Tool")
                                                                                        }
            '***Add additional table references here****
        End Get
    End Property
#End Region

#Region "Constructors"
    Public Sub New()
    End Sub

    Private Sub ConstructMe(ByVal dr As DataRow)
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        'Me.file_lkup_code = DBtoStr(dr.Item("file_lkup_code"))
        'Me.file_file_name = DBtoStr(dr.Item("file_file_name"))
        'Me.file_file_ext = DBtoStr(dr.Item("file_file_ext"))
        'Me.file_file_ver = DBtoStr(dr.Item("file_file_ver"))
    End Sub

    Public Sub New(ByVal strDS As DataSet, Optional ByVal Parent As EDSObject = Nothing, Optional ByVal dr As DataRow = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Dim dp As New DrilledPier

        ConstructMe(dr)

        For Each dpDr As DataRow In strDS.Tables(dp.EDSObjectName).Rows

            dp = New DrilledPier(dpDr)
            If Me.ID = dp.drilled_pier_tool_id Then
                Me.DrilledPiers.Add(New DrilledPier(dpDr, strDS, False, Me))
            End If
        Next
    End Sub

    Public Sub New(ByVal ExcelFilePath As String, Optional ByVal Parent As EDSObject = Nothing)
        Me.WorkBookPath = ExcelFilePath
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)
        LoadFromExcel()

    End Sub
#End Region

#Region "Load From Excel"
    Public Overrides Sub LoadFromExcel()
        ''''''Customize for each foundation type'''''
        Dim excelDS As New DataSet

        For Each item As EXCELDTParameter In ExcelDTParams
            Try
                excelDS.Tables.Add(ExcelDatasourceToDataTable(GetExcelDataSource(Me.WorkBookPath, item.xlsSheet, item.xlsRange), item.xlsDatatable))
            Catch ex As Exception
                Debug.Print(String.Format("Failed to create datatable for: {0}, {1}, {2}", IO.Path.GetFileName(Me.WorkBookPath), item.xlsSheet, item.xlsRange))
            End Try
        Next

        If excelDS.Tables.Contains("Drilled Pier Foundation") Then

            If excelDS.Tables("Drilled Pier").Rows.Count > 0 Then

                Dim dr As DataRow = excelDS.Tables("Drilled Pier Foundation").Rows(0)

                ConstructMe(dr)

                Dim myDP As New DrilledPier
                For Each dprow As DataRow In excelDS.Tables(myDP.EDSObjectName).Rows
                    If IsSomething(dprow.Item("local_drilled_pier_id")) Or (IsSomething(dprow.Item("ID")) And IsNothing(dprow.Item("local_drilled_pier_id"))) Then
                        Me.DrilledPiers.Add(New DrilledPier(dprow, excelDS, True, Me))
                    End If
                Next
            End If
        End If
    End Sub
#End Region

#Region "Save to Excel"
    Public Overrides Sub workBookFiller(ByRef wb As Workbook)
        '''''Customize for each excel tool'''''

        'Site Code Criteria
        Dim tia_current, site_name, structure_type As String
        Dim rev_h_section_15_5 As Boolean?
        Dim returnRow = 2
        Dim sumRow = 10 'Actually starts on row 11 but excel uses a (0, 0) origin

        For Each drilledPier In Me.DrilledPiers

            'Starting with column G
            'This will need updated if the structure of the database in the drilled pier spreadsheet ever changes
            Dim myCol As Integer = drilledPier.local_pier_profile_id + 5 'Column A is an origin of 0
            'Quantity of inputs associated with sections
            Dim bump15 As Integer = 15
            Dim bump3 As Integer = 3

            wb.Worksheets("Tool").Range("A3").Value = Me.ID.ToString

            With wb.Worksheets("Foundation Input")
                .Range("D3").Value = Me.bus_unit
                .Range("D4").Value = Me.ParentStructure?.structureCodeCriteria?.site_name
                .Range("D5").Value = MyOrder()
                .Range("EDSLOADED").Value = True
                .Range("FIRSTEDSOPEN").Value = True
                .Range("D6").Value = MyTIA()

                If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.structure_type) Then
                    If Me.ParentStructure?.structureCodeCriteria?.structure_type = "SELF SUPPORT" Then
                        structure_type = "Self Support"
                        .Range("D8").Value = "Self Support"
                    ElseIf Me.ParentStructure?.structureCodeCriteria?.structure_type = "MONOPOLE" Then
                        structure_type = "Monopole"
                        .Range("D8").Value = "Monopole"
                    Else
                        If Me.DrilledPiers.Count > 1 Then
                            structure_type = "Guyed (Base)"
                            .Range("D8").Value = "Base"
                        Else
                            structure_type = "Guyed (Anchor)"
                        End If
                    End If
                    .Range("D7").Value = CType(structure_type, String)
                End If
                .Range("CurrentLocation").Value = 1
                If structure_type?.Contains("Guyed") Then
                    .Range("CurrentLocation").Value = 1
                End If
            End With

            With wb.Worksheets("SUMMARY")
                If Not IsNothing(drilledPier.local_pier_profile_id) Then .Cells(sumRow, 3).Value = CType(drilledPier.local_pier_profile_id, Integer)
                If Not IsNothing(drilledPier.reaction_location) Then .Cells(sumRow, 2).Value = CType(drilledPier.reaction_location, String)
                If Not IsNothing(drilledPier.local_soil_profile) Then .Cells(sumRow, 4).Value = CType(drilledPier.local_soil_profile, Integer)
                .Cells(sumRow, 21).Value = CType(drilledPier.ID, Integer)
                .Cells(sumRow, 22).Value = CType(drilledPier.PierProfile.ID, Integer)
                If Not IsNothing(drilledPier.PierProfile.BelledPier.ID) Then .Cells(sumRow, 73).Value = CType(drilledPier.PierProfile.BelledPier.ID, Integer)
                If Not IsNothing(drilledPier.PierProfile.EmbeddedPole.ID) Then .Cells(sumRow, 74).Value = CType(drilledPier.PierProfile.EmbeddedPole.ID, Integer)
                .Cells(sumRow, 75).Value = CType(drilledPier.SoilProfile.ID, Integer)
                For Each sec As DrilledPierSection In drilledPier.PierProfile.Sections
                    Dim secAdj As Integer = 4 * (sec.local_section_id - 1) + 1
                    .Cells(sumRow, 22 + secAdj).Value = CType(sec.ID, Integer)
                    For Each reb As DrilledPierRebar In sec.Rebar
                        .Cells(sumRow, 22 + secAdj + reb.local_rebar_id).Value = CType(reb.ID, Integer)
                    Next
                Next
                Dim slRowStart As Integer = 43
                For Each layer In drilledPier.SoilProfile.DPSoilLayers
                    .Cells(sumRow, slRowStart).Value = CType(layer.ID, Integer)
                    slRowStart += 1
                Next
                sumRow += 1
            End With

            With wb.Worksheets("Database")
                'Profile
                If Not IsNothing(drilledPier.PierProfile.foundation_depth) Then .Cells(pierProfileRow + 10, myCol).Value = CType(drilledPier.PierProfile.foundation_depth, Double)
                If Not IsNothing(drilledPier.PierProfile.extension_above_grade) Then .Cells(pierProfileRow + 11, myCol).Value = CType(drilledPier.PierProfile.extension_above_grade, Double)
                If Not IsNothing(drilledPier.PierProfile.assume_min_steel) Then .Cells(pierProfileRow + 4389, myCol).Value = CType(drilledPier.PierProfile.assume_min_steel, Boolean)
                If Not IsNothing(drilledPier.PierProfile.check_shear_along_depth) Then .Cells(pierProfileRow + 97, myCol).Value = CType(drilledPier.PierProfile.check_shear_along_depth, Boolean)
                If Not IsNothing(drilledPier.PierProfile.utilize_shear_friction_methodology) Then .Cells(pierProfileRow + 98, myCol).Value = CType(drilledPier.PierProfile.utilize_shear_friction_methodology, Boolean)
                If Not IsNothing(drilledPier.PierProfile.embedded_pole) Then .Cells(pierProfileRow + 100, myCol).Value = CType(drilledPier.PierProfile.embedded_pole, Boolean)
                If Not IsNothing(drilledPier.PierProfile.belled_pier) Then .Cells(pierProfileRow + 112, myCol).Value = CType(drilledPier.PierProfile.belled_pier, Boolean)
                If Not IsNothing(drilledPier.PierProfile.concrete_compressive_strength) Then .Cells(pierProfileRow + 7, myCol).Value = CType(drilledPier.PierProfile.concrete_compressive_strength, Double)
                If Not IsNothing(drilledPier.PierProfile.longitudinal_rebar_yield_strength) Then .Cells(pierProfileRow + 8, myCol).Value = CType(drilledPier.PierProfile.longitudinal_rebar_yield_strength, Double)
                If Not IsNothing(drilledPier.PierProfile.rebar_cage_2_fy_override) Then .Cells(pierProfileRow + 4391, myCol).Value = CType(drilledPier.PierProfile.rebar_cage_2_fy_override, Double)
                If Not IsNothing(drilledPier.PierProfile.rebar_cage_3_fy_override) Then .Cells(pierProfileRow + 4392, myCol).Value = CType(drilledPier.PierProfile.rebar_cage_3_fy_override, Double)
                If Not IsNothing(drilledPier.PierProfile.rebar_effective_depths) Then .Cells(pierProfileRow + 4390, myCol).Value = CType(drilledPier.PierProfile.rebar_effective_depths, Boolean)
                If Not IsNothing(drilledPier.PierProfile.shear_crit_depth_override_comp) Then .Cells(pierProfileRow + 374, myCol).Value = CType(drilledPier.PierProfile.shear_crit_depth_override_comp, Double)
                If Not IsNothing(drilledPier.PierProfile.shear_crit_depth_override_uplift) Then .Cells(pierProfileRow + 376, myCol).Value = CType(drilledPier.PierProfile.shear_crit_depth_override_uplift, Double)
                If Not IsNothing(drilledPier.PierProfile.shear_override_crit_depth) Then .Cells(pierProfileRow + 99, myCol).Value = CType(drilledPier.PierProfile.shear_override_crit_depth, Boolean)
                If Not IsNothing(drilledPier.PierProfile.tie_yield_strength) Then .Cells(pierProfileRow + 9, myCol).Value = CType(drilledPier.PierProfile.tie_yield_strength, Double)
                .Cells(pierProfileRow + 18, myCol).Value = drilledPier.SoilProfile.DPSoilLayers.Count
                If drilledPier.PierProfile.ultimate_gross_bearing Then
                    .Cells(pierProfileRow + 19, myCol).Value = "Ult. Gross Bearing Capacity (ksf)"
                Else
                    .Cells(pierProfileRow + 19, myCol).Value = "Ult. Net Bearing Capacity (ksf)"
                End If

                'H Section 15.5
                If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.rev_h_section_15_5) Then
                    rev_h_section_15_5 = Me.ParentStructure?.structureCodeCriteria?.rev_h_section_15_5
                    .Cells(pierProfileRow + 95, myCol).Value = CType(rev_h_section_15_5, Boolean)
                End If

                'Embedded pole
                If drilledPier.PierProfile.embedded_pole Then
                    If Not IsNothing(drilledPier.PierProfile.EmbeddedPole.encased_in_concrete) Then .Cells(pierProfileRow + 101, myCol).Value = CType(drilledPier.PierProfile.EmbeddedPole.encased_in_concrete, Boolean)
                    If Not IsNothing(drilledPier.PierProfile.EmbeddedPole.pole_side_quantity) Then .Cells(pierProfileRow + 102, myCol).Value = CType(drilledPier.PierProfile.EmbeddedPole.pole_side_quantity, Integer)
                    If Not IsNothing(drilledPier.PierProfile.EmbeddedPole.pole_yield_strength) Then .Cells(pierProfileRow + 103, myCol).Value = CType(drilledPier.PierProfile.EmbeddedPole.pole_yield_strength, Double)
                    If Not IsNothing(drilledPier.PierProfile.EmbeddedPole.pole_thickness) Then .Cells(pierProfileRow + 104, myCol).Value = CType(drilledPier.PierProfile.EmbeddedPole.pole_thickness, Double)
                    If Not IsNothing(drilledPier.PierProfile.EmbeddedPole.embedded_pole_input_type) Then .Cells(pierProfileRow + 105, myCol).Value = CType(drilledPier.PierProfile.EmbeddedPole.embedded_pole_input_type, String)
                    If Not IsNothing(drilledPier.PierProfile.EmbeddedPole.pole_diameter_toc) Then .Cells(pierProfileRow + 106, myCol).Value = CType(drilledPier.PierProfile.EmbeddedPole.pole_diameter_toc, Double)
                    If Not IsNothing(drilledPier.PierProfile.EmbeddedPole.pole_top_diameter) Then .Cells(pierProfileRow + 107, myCol).Value = CType(drilledPier.PierProfile.EmbeddedPole.pole_top_diameter, Double)
                    If Not IsNothing(drilledPier.PierProfile.EmbeddedPole.pole_bottom_diameter) Then .Cells(pierProfileRow + 108, myCol).Value = CType(drilledPier.PierProfile.EmbeddedPole.pole_bottom_diameter, Double)
                    If Not IsNothing(drilledPier.PierProfile.EmbeddedPole.pole_section_length) Then .Cells(pierProfileRow + 109, myCol).Value = CType(drilledPier.PierProfile.EmbeddedPole.pole_section_length, Double)
                    If Not IsNothing(drilledPier.PierProfile.EmbeddedPole.pole_taper_factor) Then .Cells(pierProfileRow + 110, myCol).Value = CType(drilledPier.PierProfile.EmbeddedPole.pole_taper_factor, Double)
                    If Not IsNothing(drilledPier.PierProfile.EmbeddedPole.pole_bend_radius_override) Then .Cells(pierProfileRow + 111, myCol).Value = CType(drilledPier.PierProfile.EmbeddedPole.pole_bend_radius_override, Double)
                End If

                'Belled Pier
                If drilledPier.PierProfile.belled_pier Then
                    If Not IsNothing(drilledPier.PierProfile.BelledPier.bottom_diameter_of_bell) Then .Cells(pierProfileRow + 113, myCol).Value = CType(drilledPier.PierProfile.BelledPier.bottom_diameter_of_bell, Double)
                    If Not IsNothing(drilledPier.PierProfile.BelledPier.bell_input_type) Then .Cells(pierProfileRow + 114, myCol).Value = CType(drilledPier.PierProfile.BelledPier.bell_input_type, String)
                    If Not IsNothing(drilledPier.PierProfile.BelledPier.bell_angle) Then .Cells(pierProfileRow + 115, myCol).Value = CType(drilledPier.PierProfile.BelledPier.bell_angle, Double)
                    If Not IsNothing(drilledPier.PierProfile.BelledPier.bell_height) Then .Cells(pierProfileRow + 116, myCol).Value = CType(drilledPier.PierProfile.BelledPier.bell_height, Double)
                    If Not IsNothing(drilledPier.PierProfile.BelledPier.bell_toe_height) Then .Cells(pierProfileRow + 120, myCol).Value = CType(drilledPier.PierProfile.BelledPier.bell_toe_height, Double)
                    If Not IsNothing(drilledPier.PierProfile.BelledPier.neglect_top_soil_layer) Then .Cells(pierProfileRow + 122, myCol).Value = CType(drilledPier.PierProfile.BelledPier.neglect_top_soil_layer, Boolean)
                    If Not IsNothing(drilledPier.PierProfile.BelledPier.swelling_expansive_soil) Then .Cells(pierProfileRow + 123, myCol).Value = CType(drilledPier.PierProfile.BelledPier.swelling_expansive_soil, Boolean)
                    If Not IsNothing(drilledPier.PierProfile.BelledPier.depth_of_expansive_soil) Then .Cells(pierProfileRow + 124, myCol).Value = CType(drilledPier.PierProfile.BelledPier.depth_of_expansive_soil, Double)
                    If Not IsNothing(drilledPier.PierProfile.BelledPier.expansive_soil_force) Then .Cells(pierProfileRow + 125, myCol).Value = CType(drilledPier.PierProfile.BelledPier.expansive_soil_force, Double)
                End If

                'Sections
                Dim rebCountAdj As Integer = 0
                For Each section In drilledPier.PierProfile.Sections
                    Dim secAdj As Integer = section.local_section_id - 5 * (1 - 1) - 1 ' '1-1 --> drilledPier.local_pier_profile_id - 1

                    If Not IsNothing(section.bottom_elevation) Then .Cells(pierProfileRow + 12 + secAdj, myCol).Value = CType(section.bottom_elevation, Integer)
                    If Not IsNothing(section.pier_diameter) Then .Cells(pierProfileRow + 20 + bump15 * secAdj, myCol).Value = CType(section.pier_diameter, Double)
                    If Not IsNothing(section.clear_cover) Then .Cells(pierProfileRow + 23 + bump15 * secAdj, myCol).Value = CType(section.clear_cover, Double)
                    If Not IsNothing(section.tie_size) Then .Cells(pierProfileRow + 24 + bump15 * secAdj, myCol).Value = CType(section.tie_size, Integer)
                    If Not IsNothing(section.tie_spacing) Then .Cells(pierProfileRow + 25 + bump15 * secAdj, myCol).Value = CType(section.tie_spacing, Double)
                    With .Cells(pierProfileRow + 34 + bump15 * secAdj, myCol)
                        If section.clear_cover_rebar_cage_option Then
                            .Value = "Clear Cover to Ties"
                        Else
                            .Value = "Rebar Cage Diameter"
                        End If
                    End With

                    'Rebar
                    For Each rebar In section.Rebar
                        Select Case rebar.local_rebar_id - 15 * (1 - 1) '1-1 --> drilledPier.local_pier_profile_id - 1
                            Case 1
                                If Not IsNothing(rebar.longitudinal_rebar_quantity) Then .Cells(pierProfileRow + 21 + bump15 * secAdj, myCol).Value = CType(rebar.longitudinal_rebar_quantity, Integer)
                                If Not IsNothing(rebar.longitudinal_rebar_size) Then .Cells(pierProfileRow + 22 + bump15 * secAdj, myCol).Value = CType(rebar.longitudinal_rebar_size, Integer)
                            Case 2
                                If Not IsNothing(rebar.longitudinal_rebar_quantity) Then .Cells(pierProfileRow + 26 + bump15 * secAdj, myCol).Value = CType(rebar.longitudinal_rebar_quantity, Integer)
                                If Not IsNothing(rebar.longitudinal_rebar_size) Then .Cells(pierProfileRow + 27 + bump15 * secAdj, myCol).Value = CType(rebar.longitudinal_rebar_size, Integer)
                                If Not IsNothing(rebar.longitudinal_rebar_cage_diameter) Then .Cells(pierProfileRow + 28 + bump15 * secAdj, myCol).Value = CType(rebar.longitudinal_rebar_cage_diameter, Integer)
                                If Not IsNothing(section.tie_size) Then .Cells(pierProfileRow + 29 + bump15 * secAdj, myCol).Value = CType(section.tie_size, Integer)
                            Case 3
                                If Not IsNothing(rebar.longitudinal_rebar_quantity) Then .Cells(pierProfileRow + 30 + bump15 * secAdj, myCol).Value = CType(rebar.longitudinal_rebar_quantity, Integer)
                                If Not IsNothing(rebar.longitudinal_rebar_size) Then .Cells(pierProfileRow + 31 + bump15 * secAdj, myCol).Value = CType(rebar.longitudinal_rebar_size, Integer)
                                If Not IsNothing(rebar.longitudinal_rebar_cage_diameter) Then .Cells(pierProfileRow + 32 + bump15 * secAdj, myCol).Value = CType(rebar.longitudinal_rebar_cage_diameter, Integer)
                                If Not IsNothing(section.tie_size) Then .Cells(pierProfileRow + 33 + bump15 * secAdj, myCol).Value = CType(section.tie_size, Integer)
                        End Select
                    Next

                    If Not structure_type?.Contains("Guyed") Then
                        'Row R of the foundation input tab. 
                        'This only works for Monopoles and SSTs
                        'Guyed towers with a drilled pier with sections with multiple rebar is a known issue with the tool.
                        wb.Worksheets("Foundation Input").Range("R" & 109 + rebCountAdj).Value = CType(section.Rebar.Count.ToString, Integer)
                    End If
                    rebCountAdj += 1
                Next

                'Soil Profile
                If Not IsNothing(drilledPier.SoilProfile.groundwater_depth) Then
                    If drilledPier.SoilProfile.groundwater_depth <= 0 Then
                        .Cells(pierProfileRow + 17, myCol).Value = "N/A"
                    Else
                        .Cells(pierProfileRow + 17, myCol).Value = CType(drilledPier.SoilProfile.groundwater_depth, Double)
                    End If

                End If

                'Soil Layers
                Dim layAdj As Integer = 0
                For Each layer In drilledPier.SoilProfile.DPSoilLayers
                    If Not IsNothing(layer.bottom_depth) Then .Cells(pierProfileRow + layAdj + 127, myCol).Value = CType(layer.bottom_depth, Double)
                    If Not IsNothing(layer.effective_soil_density) Then .Cells(pierProfileRow + layAdj + 158, myCol).Value = CType(layer.effective_soil_density, Double)
                    If Not IsNothing(layer.cohesion) Then .Cells(pierProfileRow + layAdj + 189, myCol).Value = CType(layer.cohesion, Double)
                    If Not IsNothing(layer.friction_angle) Then .Cells(pierProfileRow + layAdj + 220, myCol).Value = CType(layer.friction_angle, Double)
                    If Not IsNothing(layer.skin_friction_override_comp) Then .Cells(pierProfileRow + layAdj + 251, myCol).Value = CType(layer.skin_friction_override_comp, Double)
                    If Not IsNothing(layer.skin_friction_override_uplift) Then .Cells(pierProfileRow + layAdj + 282, myCol).Value = CType(layer.skin_friction_override_uplift, Double)
                    If Not IsNothing(layer.nominal_bearing_capacity) Then .Cells(pierProfileRow + layAdj + 313, myCol).Value = CType(layer.nominal_bearing_capacity, Double)
                    If Not IsNothing(layer.spt_blow_count) Then .Cells(pierProfileRow + layAdj + 344, myCol).Value = CType(layer.spt_blow_count, Double)
                    layAdj += 1
                Next
            End With
        Next

    End Sub
#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bus_unit.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.structure_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.work_order_seq_num.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.file_lkup_code.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.file_file_name.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.file_file_ext.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.file_file_ver.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Now.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = SQLInsertFields.AddtoDBString("bus_unit")
        SQLInsertFields = SQLInsertFields.AddtoDBString("structure_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("work_order_seq_num")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("file_lkup_code")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("file_file_name")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("file_file_ext")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("file_file_ver")
        SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")
        SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("upload_date")

        Return SQLInsertFields

    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("work_order_seq_num = " & Me.work_order_seq_num.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("file_lkup_code = " & Me.file_lkup_code.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("file_file_name = " & Me.file_file_name.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("file_file_ext = " & Me.file_file_ext.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("file_file_ver = " & Me.file_file_ver.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("process_stage = " & Me.process_stage.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_person_id = " & Me.modified_person_id.ToString.FormatDBValue)
        Return SQLUpdateFieldsandValues
    End Function
#End Region

#Region "Equals"
    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        Dim otherToCompare As DrilledPierFoundation = TryCast(other, DrilledPierFoundation)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)

        'Drilled Piers
        Equals = If(Me.DrilledPiers.CheckChange(otherToCompare.DrilledPiers, changes, categoryName, "Drilled Piers"), Equals, False)

        Return Equals
    End Function
#End Region

    Public Overrides Sub Clear()
        Me.DrilledPiers.Clear()
        Me.Results.Clear()
    End Sub
End Class

<DataContractAttribute()>
<KnownType(GetType(DrilledPier))>
Partial Public Class DrilledPier
    Inherits EDSObjectWithQueries

#Region "Inheritted"
    '''Must override these inherited properties
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Drilled Pier"
        End Get
    End Property
    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "fnd.drilled_pier"
        End Get
    End Property

    Private _Insert As String
    Private _Update As String
    Private _Delete As String

    Public Overrides Function SQLInsert() As String
        SQLInsert = ""
        SQLInsert = CCI_Engineering_Templates.My.Resources.Drilled_Pier__INSERT  'QueryBuilderFromFile(queryPath & "Drilled Pier\Drilled Pier (INSERT).sql")
        SQLInsert = SQLInsert.Replace("[TABLE NAME]", Me.EDSTableName)
        SQLInsert = SQLInsert.Replace("[INSERT FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.Replace("[INSERT VALUES]", Me.SQLInsertValues)

        SQLInsert = SQLInsert.Replace("--[PIER PROFILE INSERT]", Me.PierProfile.SQLInsert)
        Dim tempProf As New SoilProfile
        tempProf = TryCast(Me.SoilProfile, SoilProfile)

        For Each layer As DrilledPierSoilLayer In Me.SoilProfile.DPSoilLayers
            Dim newlay As New SoilLayer
            newlay = TryCast(layer, SoilLayer)
            tempProf.SoilLayers.Add(newlay)
        Next

        SQLInsert = SQLInsert.Replace("--[SOIL PROFILE INSERT]", tempProf.SQLInsert)

        Dim resInsert As String = ""
        For Each res As EDSResult In Me.Results
            resInsert += res.Insert & vbCrLf
        Next

        SQLInsert = SQLInsert.Replace("--[RESULTS]", resInsert)


        SQLInsert = SQLInsert.TrimEnd

        Return SQLInsert
    End Function

    Public Overrides Function SQLUpdate() As String
        SQLUpdate = ""
        SQLUpdate = CCI_Engineering_Templates.My.Resources.General__UPDATE ' QueryBuilderFromFile(queryPath & "Drilled Pier\General (UPDATE).sql")
        SQLUpdate = SQLUpdate.Replace("[TABLE NAME]", Me.EDSTableName)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID)

        Dim _pierInsert As String
        Dim _soilInsert As String

        If Me.PierProfile?.ID IsNot Nothing And Me.PierProfile?.ID > 0 Then
            _pierInsert = Me.PierProfile.SQLUpdate
        Else
            _pierInsert = Me.PierProfile.SQLInsert
        End If

        If Me.SoilProfile?.ID IsNot Nothing And Me.SoilProfile?.ID > 0 Then
            _soilInsert = Me.SoilProfile.SQLUpdate
        Else
            _soilInsert = Me.SoilProfile.SQLInsert
        End If

        SQLUpdate = SQLUpdate.Replace("--[OPTIONAL]", _pierInsert + vbCrLf + _soilInsert + vbCrLf + Me.ResultQuery(True))
        SQLUpdate = SQLUpdate.TrimEnd

        Return SQLUpdate
    End Function

    Public Overrides Function SQLDelete() As String

        SQLDelete = CCI_Engineering_Templates.My.Resources.General__DELETE
        SQLDelete = SQLDelete.Replace("[TABLE]", Me.EDSTableName)
        SQLDelete = SQLDelete.Replace("[ID]", Me.ID)

        SQLDelete += Me.PierProfile?.SQLDelete + vbCrLf
        SQLDelete += Me.SoilProfile?.SQLDelete + vbCrLf

        Return SQLDelete
    End Function

#End Region

#Region "Define"
    <DataMember()> Public Property PierProfile As DrilledPierProfile
    <DataMember()> Public Property SoilProfile As DrilledPierSoilProfile


    Private _local_drilled_pier_id
    Private _pier_profile_id As Integer?
    Private _soil_profile_id As Integer?
    Private _reaction_position As Integer?
    Private _reaction_location As String
    Private _local_soil_profile As Integer?
    Private _local_pier_profile_id As Integer?
    Private _drilled_pier_tool_id As Integer?


    <Category("Drilled Pier"), Description(""), DisplayName("Drilled Pier Profile Id")>
    <DataMember()> Public Property pier_profile_id() As Integer?
        Get
            Return Me._pier_profile_id
        End Get
        Set
            Me._pier_profile_id = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Soil Profile Id")>
    <DataMember()> Public Property soil_profile_id() As Integer?
        Get
            Return Me._soil_profile_id
        End Get
        Set
            Me._soil_profile_id = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Reaction Position")>
    <DataMember()> Public Property reaction_position() As Integer?
        Get
            Return Me._reaction_position
        End Get
        Set
            Me._reaction_position = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Reaction Location")>
    <DataMember()> Public Property reaction_location() As String
        Get
            Return Me._reaction_location
        End Get
        Set
            Me._reaction_location = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Local Soil Profile")>
    <DataMember()> Public Property local_soil_profile() As Integer?
        Get
            Return Me._local_soil_profile
        End Get
        Set
            Me._local_soil_profile = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Local Drilled Pier Profile")>
    <DataMember()> Public Property local_pier_profile_id() As Integer?
        Get
            Return Me._local_pier_profile_id
        End Get
        Set
            Me._local_pier_profile_id = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Local Drilled Pier ID")>
    <DataMember()> Public Property local_drilled_pier_id() As Integer?
        Get
            Return Me._local_drilled_pier_id
        End Get
        Set
            Me._local_drilled_pier_id = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Drilled Pier Tool ID")>
    <DataMember()> Public Property drilled_pier_tool_id() As Integer?
        Get
            Return Me._drilled_pier_tool_id
        End Get
        Set
            Me._drilled_pier_tool_id = Value
        End Set
    End Property


#End Region

#Region "Constructors"
    Public Sub New()

    End Sub

    Public Sub New(ByVal dr As DataRow)
        ConstructMe(dr)
    End Sub

    Public Sub ConstructMe(ByVal dr As DataRow)
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        Me.pier_profile_id = DBtoNullableInt(dr.Item("pier_profile_id"))
        Me.soil_profile_id = DBtoNullableInt(dr.Item("soil_profile_id"))
        Me.reaction_position = DBtoNullableInt(dr.Item("reaction_position"))
        Me.reaction_location = DBtoStr(dr.Item("reaction_location"))
        Me.local_soil_profile = DBtoNullableInt(dr.Item("local_soil_profile_id"))
        Me.local_pier_profile_id = DBtoNullableInt(dr.Item("local_pier_profile_id"))
        Me.local_drilled_pier_id = DBtoNullableInt(dr.Item("local_drilled_pier_id"))
        Try
            Me.drilled_pier_tool_id = DBtoNullableInt(dr.Item("drilled_pier_tool_id"))
        Catch
            Me.drilled_pier_tool_id = Nothing
        End Try
    End Sub

    Public Sub New(ByVal dr As DataRow, ByVal strDS As DataSet, ByVal isExcel As Boolean, Optional ByVal Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Dim dpProfile As New DrilledPierProfile
        Dim dpSection As New DrilledPierSection
        Dim dpRebar As New DrilledPierRebar
        Dim dpBelled As New BelledPier
        Dim dpEmbedded As New EmbeddedPole
        Dim dpSProfile As New DrilledPierSoilProfile
        Dim dpLayer As New DrilledPierSoilLayer

        ConstructMe(dr)

        For Each profileRow As DataRow In strDS.Tables(dpProfile.EDSObjectName).Rows
            dpProfile = New DrilledPierProfile(profileRow, Me)
            If IsSomething(dpProfile.local_drilled_pier_id) Or (IsSomething(dpProfile.ID) And IsNothing(dpProfile.local_drilled_pier_id)) Then
                If If(isExcel, Me.local_pier_profile_id = dpProfile.local_pier_profile_id And Me.local_drilled_pier_id = dpProfile.local_drilled_pier_id, Me.pier_profile_id = dpProfile.ID) Then
                    Me.PierProfile = dpProfile
                    For Each sectionrow As DataRow In strDS.Tables(dpSection.EDSObjectName).Rows
                        dpSection = (New DrilledPierSection(sectionrow, dpProfile))
                        If IsSomething(dpSection.local_section_id) Or (IsSomething(dpSection.ID) And IsNothing(dpSection.local_section_id)) Then
                            If If(isExcel, dpProfile.local_drilled_pier_id = dpSection.local_drilled_pier_id, dpProfile.ID = dpSection.pier_profile_id) Then
                                dpProfile.Sections.Add(dpSection)

                                For Each rebarRow As DataRow In strDS.Tables(dpRebar.EDSObjectName).Rows
                                    dpRebar = (New DrilledPierRebar(rebarRow, dpSection))
                                    If IsSomething(dpRebar.local_rebar_id) Or (IsSomething(dpRebar.ID) And IsNothing(dpRebar.local_rebar_id)) Then
                                        If If(isExcel, dpSection.local_section_id = dpRebar.local_section_id And dpRebar.local_drilled_pier_id = dpSection.local_drilled_pier_id, dpSection.ID = dpRebar.section_id) Then
                                            dpSection.Rebar.Add(dpRebar)
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    Next

                    For Each bellRow As DataRow In strDS.Tables(dpBelled.EDSObjectName).Rows
                        dpBelled = (New BelledPier(bellRow, dpProfile))
                        If If(isExcel, dpProfile.local_drilled_pier_id = dpBelled.local_drilled_pier_id, dpProfile.ID = dpBelled.pier_profile_id) Then
                            Me.PierProfile.BelledPier = dpBelled
                        End If
                    Next

                    For Each emedRow As DataRow In strDS.Tables(dpEmbedded.EDSObjectName).Rows
                        dpEmbedded = (New EmbeddedPole(emedRow, dpProfile))
                        If If(isExcel, dpProfile.local_drilled_pier_id = dpEmbedded.local_drilled_pier_id, dpProfile.ID = dpEmbedded.pier_profile_id) Then
                            Me.PierProfile.EmbeddedPole = dpEmbedded
                        End If
                    Next
                End If
            End If
        Next

        For Each soilRow As DataRow In strDS.Tables(dpSProfile.EDSObjectName).Rows
            dpSProfile = (New DrilledPierSoilProfile(soilRow, Me))
            If If(isExcel, Me.local_soil_profile = dpSProfile.local_soil_profile_id And Me.local_drilled_pier_id = dpSProfile.local_drilled_pier_id, Me.soil_profile_id = dpSProfile.ID) Then
                Me.SoilProfile = dpSProfile

                For Each layerRow As DataRow In strDS.Tables(dpLayer.EDSObjectName).Rows
                    dpLayer = (New DrilledPierSoilLayer(layerRow, dpSProfile))
                    If If(isExcel, dpSProfile.local_drilled_pier_id = dpLayer.local_drilled_pier_id Or Me.soil_profile_id = dpLayer.soil_profile_id, Me.soil_profile_id = dpLayer.soil_profile_id) Then
                        dpSProfile.DPSoilLayers.Add(dpLayer)
                    End If
                Next
            End If
        Next

        If isExcel Then
            Dim res As New EDSResult
            Dim dt As New DataTable
            dt.Columns.Add("result_lkup", GetType(String))
            dt.Columns.Add("rating", GetType(Double))
            Dim rowNum As Integer = 0
            'Workflow bug 218 
            'Added wording to separate anchor and base foundations for guyed towers
            Dim extra As String = ""
            For Each resRow In strDS.Tables("Drilled Pier Result").Rows
                rowNum += 1
                If rowNum > 1 Then extra = "_ANCHOR"
                If DBtoNullableInt(resRow.item("local_drilled_pier_id")) = Me.local_drilled_pier_id Then

                    Try
                        IIf(resRow.item("Soil Later (Comp)") IsNot Nothing, dt.Rows.Add("FND_SOIL_LAT_COMP" + extra, Math.Round(CType(resRow.item("Soil Later (Comp)"), Double), 3) * 100), Nothing)
                    Catch
                    End Try
                    Try
                        IIf(resRow.item("Soil Later (Uplift)") IsNot Nothing, dt.Rows.Add("FND_SOIL_LAT_UP" + extra, Math.Round(CType(resRow.item("Soil Later (Uplift)"), Double), 3) * 100), Nothing)
                    Catch
                    End Try
                    Try
                        IIf(resRow.item("Soil Vertical (Comp)") IsNot Nothing, dt.Rows.Add("FND_SOIL_VERT_COMP" + extra, Math.Round(CType(resRow.item("Soil Vertical (Comp)"), Double), 3) * 100), Nothing)
                    Catch
                    End Try
                    Try
                        IIf(resRow.item("Soil Vertical (Uplift)") IsNot Nothing, dt.Rows.Add("FND_SOIL_VERT_UP" + extra, Math.Round(CType(resRow.item("Soil Vertical (Uplift)"), Double), 3) * 100), Nothing)
                    Catch
                    End Try
                    Try
                        IIf(resRow.item("Concrete Flexure (Uplift)") IsNot Nothing, dt.Rows.Add("FND_CON_FLEX_UP" + extra, Math.Round(CType(resRow.item("Concrete Flexure (Uplift)"), Double), 3) * 100), Nothing)
                    Catch
                    End Try
                    Try
                        IIf(resRow.item("Concrete Shear (Comp)") IsNot Nothing, dt.Rows.Add("FND_CON_SHEAR_COMP" + extra, Math.Round(CType(resRow.item("Concrete Shear (Comp)"), Double), 3) * 100), Nothing)
                    Catch
                    End Try
                    Try
                        IIf(resRow.item("Concrete Flexure (Comp)") IsNot Nothing, dt.Rows.Add("FND_CON_FLEX_COMP" + extra, Math.Round(CType(resRow.item("Concrete Flexure (Comp)"), Double), 3) * 100), Nothing)
                    Catch
                    End Try
                    Try
                        IIf(resRow.item("Concrete Shear (Uplift)") IsNot Nothing, dt.Rows.Add("FND_CON_SHEAR_UP" + extra, Math.Round(CType(resRow.item("Concrete Shear (Uplift)"), Double), 3) * 100), Nothing)
                    Catch
                    End Try
                    Try
                        IIf(resRow.item("Soil Max") IsNot Nothing, dt.Rows.Add("FND_SOIL_MAX" + extra, Math.Round(CType(resRow.item("Soil Max"), Double), 3) * 100), Nothing)
                    Catch
                    End Try
                    Try
                        IIf(resRow.item("Structure Max") IsNot Nothing, dt.Rows.Add("FND_STRC_MAX" + extra, Math.Round(CType(resRow.item("Structure Max"), Double), 3) * 100), Nothing)
                    Catch
                    End Try
                    'dt.Rows.Add("DPSOIL", Math.Round(CType(resRow.item("Soil Rating"), Double), 3))
                    'dt.Rows.Add("DPSTRUC", Math.Round(CType(resRow.item("Structural Rating"), Double), 3))
                    'Exit For
                End If
            Next

            For Each resRow As DataRow In dt.Rows
                resRow.Item("rating") = resRow.Item("rating") / 100
                res = New EDSResult(resRow, Me.Parent)
                res.EDSTableName = "fnd.drilled_pier_results"
                res.ForeignKeyName = "drilled_pier_id"
                res.foreign_key = Me.ID
                res.EDSTableDepth = 1
                res.modified_person_id = Me.modified_person_id
                Dim x As DrilledPierFoundation = Me.Parent
                x.Results.Add(res)
                Me.Results.Add(res)
            Next
        End If
    End Sub
#End Region

#Region "Save to EDS"

    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bus_unit.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.structure_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel2ID")
        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel1ID")
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reaction_position.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reaction_location.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_soil_profile.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_pier_profile_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel4ID")
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
        Return SQLInsertValues

    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = SQLInsertFields.AddtoDBString("bus_unit")
        SQLInsertFields = SQLInsertFields.AddtoDBString("structure_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pier_profile_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("soil_profile_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("reaction_position")
        SQLInsertFields = SQLInsertFields.AddtoDBString("reaction_location")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_soil_profile_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_pier_profile_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("drilled_pier_tool_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")
        SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
        Return SQLInsertFields

    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pier_profile_id = " & Me.pier_profile_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("soil_profile_id = " & Me.soil_profile_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("reaction_position = " & Me.reaction_position.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("reaction_location = " & Me.reaction_location.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_soil_profile_id = " & Me.local_soil_profile.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_pier_profile_id = " & Me.local_pier_profile_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("process_stage = " & Me.process_stage.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_person_id = " & Me.modified_person_id.ToString.FormatDBValue)
        Return SQLUpdateFieldsandValues

    End Function

#End Region

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As DrilledPier = TryCast(other, DrilledPier)
        If otherToCompare Is Nothing Then Return False
        Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
        Equals = If(Me.pier_profile_id.CheckChange(otherToCompare.pier_profile_id, changes, categoryName, "Drilled Pier Profile Id"), Equals, False)
        Equals = If(Me.soil_profile_id.CheckChange(otherToCompare.soil_profile_id, changes, categoryName, "Soil Profile Id"), Equals, False)
        Equals = If(Me.reaction_position.CheckChange(otherToCompare.reaction_position, changes, categoryName, "Reaction Position"), Equals, False)
        Equals = If(Me.reaction_location.CheckChange(otherToCompare.reaction_location, changes, categoryName, "Reaction Location"), Equals, False)
        Equals = If(Me.local_soil_profile.CheckChange(otherToCompare.local_soil_profile, changes, categoryName, "Local Soil Profile"), Equals, False)
        Equals = If(Me.local_pier_profile_id.CheckChange(otherToCompare.local_pier_profile_id, changes, categoryName, "Local Drilled Pier Profile"), Equals, False)

        'Drilled Pier Profile
        Equals = If(Me.PierProfile.CheckChange(otherToCompare.PierProfile, changes, categoryName, "Pier Profile"), Equals, False)

        'Soil Profile
        Equals = If(Me.SoilProfile.CheckChange(otherToCompare.SoilProfile, changes, categoryName, "Soil Profile"), Equals, False)

        Return Equals

    End Function

End Class
<DataContractAttribute()>
<KnownType(GetType(DrilledPierProfile))>
Partial Public Class DrilledPierProfile
    Inherits EDSObjectWithQueries

#Region "Inheritted"
    '''Must override these inherited properties
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Drilled Pier Profile"
        End Get
    End Property

    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "fnd.drilled_pier_profile"
        End Get
    End Property

    Private _Insert As String
    Private _Update As String
    Private _Delete As String

    Public Overrides Function SQLInsert() As String
        SQLInsert = ""
        SQLInsert = CCI_Engineering_Templates.My.Resources.Drilled_Pier_Profile__INSERT ' QueryBuilderFromFile(queryPath & "Drilled Pier\Drilled Pier Profile (INSERT).sql")
        SQLInsert = SQLInsert.Replace("[TABLE NAME]", Me.EDSTableName)
        SQLInsert = SQLInsert.Replace("[INSERT FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.Replace("[INSERT VALUES]", Me.SQLInsertValues)

        Dim _sectionInsert As String
        Dim _belledInsert As String
        Dim _embedInsert As String

        For Each sec In Me.Sections
            'If sec.ID IsNot Nothing And sec?.ID > 0 Then
            '    If sec.local_section_id IsNot Nothing Then
            '        _sectionInsert += sec.SQLUpdate + vbCrLf
            '    Else
            '        _sectionInsert += sec.SQLDelete + vbCrLf
            '    End If
            'Else
            _sectionInsert += sec.SQLInsert + vbCrLf
            'End If
        Next

        If Me.belled_pier Then
            'If Me.BelledPier?.ID IsNot Nothing And Me.BelledPier?.ID > 0 Then
            '    If Me.belled_pier = False Then
            '        _belledInsert = Me.BelledPier.SQLDelete + vbCrLf
            '    Else
            '        _belledInsert = Me.BelledPier.SQLUpdate + vbCrLf
            '    End If
            'Else
            _belledInsert = Me.BelledPier.SQLInsert + vbCrLf
            'End If
        End If

        If Me.embedded_pole Then
            'If Me.EmbeddedPole?.ID IsNot Nothing And Me.EmbeddedPole?.ID > 0 Then
            '    If Me.embedded_pole = False Then
            '        _embedInsert = Me.EmbeddedPole.SQLDelete + vbCrLf
            '    Else
            '        _embedInsert = Me.EmbeddedPole.SQLUpdate + vbCrLf
            '    End If
            'Else
            _embedInsert = Me.EmbeddedPole.SQLInsert + vbCrLf
            'End If
        End If

        SQLInsert = SQLInsert.Replace("--[SECTION INSERT]", _sectionInsert + vbCrLf)
        SQLInsert = SQLInsert.Replace("--[BELLED INSERT]", _belledInsert + vbCrLf)
        SQLInsert = SQLInsert.Replace("--[EMBEDDED INSERT]", _embedInsert + vbCrLf)
        SQLInsert = SQLInsert.TrimEnd()

        Return SQLInsert
    End Function

    Public Overrides Function SQLUpdate() As String
        SQLUpdate = ""
        SQLUpdate = CCI_Engineering_Templates.My.Resources.General__UPDATE ' QueryBuilderFromFile(queryPath & "Drilled Pier\General (UPDATE).sql")
        SQLUpdate = SQLUpdate.Replace("[TABLE NAME]", Me.EDSTableName)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID & vbCrLf & "SET @SubLevel2ID = " & Me.ID)

        Dim _sectionInsert As String
        Dim _belledInsert As String
        Dim _embedInsert As String

        For Each sec In Me.Sections
            If sec.ID IsNot Nothing And sec?.ID > 0 Then
                If sec.local_section_id IsNot Nothing Then
                    _sectionInsert += sec.SQLUpdate + vbCrLf
                Else
                    _sectionInsert += sec.SQLDelete + vbCrLf
                End If
            Else
                _sectionInsert += sec.SQLInsert + vbCrLf
            End If
        Next

        If Me.belled_pier Then
            If Me.BelledPier?.ID IsNot Nothing And Me.BelledPier?.ID > 0 Then
                If Me.belled_pier = False Then
                    _belledInsert = Me.BelledPier.SQLDelete + vbCrLf
                Else
                    _belledInsert = Me.BelledPier.SQLUpdate + vbCrLf
                End If
            Else
                _belledInsert = Me.BelledPier.SQLInsert + vbCrLf
            End If
        End If

        If Me.embedded_pole Then
            If Me.EmbeddedPole?.ID IsNot Nothing And Me.EmbeddedPole?.ID > 0 Then
                If Me.embedded_pole = False Then
                    _embedInsert = Me.EmbeddedPole.SQLDelete + vbCrLf
                Else
                    _embedInsert = Me.EmbeddedPole.SQLUpdate + vbCrLf
                End If
            Else
                _embedInsert = Me.EmbeddedPole.SQLInsert + vbCrLf
            End If
        End If

        _Insert = ""
        _Insert = _sectionInsert + _belledInsert + _embedInsert
        SQLUpdate = SQLUpdate.Replace("--[OPTIONAL]", _Insert + vbCrLf)
        SQLUpdate.TrimEnd()

        Return SQLUpdate
    End Function

    Public Overrides Function SQLDelete() As String
        SQLDelete = CCI_Engineering_Templates.My.Resources.General__DELETE
        SQLDelete = SQLDelete.Replace("[TABLE]", Me.EDSTableName)
        SQLDelete = SQLDelete.Replace("[ID]", Me.ID)

        Return SQLDelete
    End Function

#End Region

#Region "Define"
    <DataMember()> Public Property Sections As New List(Of DrilledPierSection)
    <DataMember()> Public Property EmbeddedPole As New EmbeddedPole
    <DataMember()> Public Property BelledPier As New BelledPier


    Private _foundation_depth As Double?
    Private _extension_above_grade As Double?
    Private _assume_min_steel As Boolean?
    Private _check_shear_along_depth As Boolean?
    Private _utilize_shear_friction_methodology As Boolean?
    Private _embedded_pole As Boolean?
    Private _belled_pier As Boolean?
    Private _concrete_compressive_strength As Double?
    Private _longitudinal_rebar_yield_strength As Double?
    Private _rebar_cage_2_fy_override As Double?
    Private _rebar_cage_3_fy_override As Double?
    Private _rebar_effective_depths As Boolean?
    Private _shear_crit_depth_override_comp As Double?
    Private _shear_crit_depth_override_uplift As Double?
    Private _shear_override_crit_depth As Boolean?
    Private _tie_yield_strength As Double?
    Private _tool_version As String
    Private _ultimate_gross_bearing As Boolean?
    Private _local_pier_profile_id As Integer?
    Private _local_drilled_pier_id As Integer?
    Private _drilled_pier_id As Integer?


    <Category("Drilled Pier"), Description(""), DisplayName("Foundation Depth")>
    <DataMember()> Public Property foundation_depth() As Double?
        Get
            Return Me._foundation_depth
        End Get
        Set
            Me._foundation_depth = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Extension Above Grade")>
    <DataMember()> Public Property extension_above_grade() As Double?
        Get
            Return Me._extension_above_grade
        End Get
        Set
            Me._extension_above_grade = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Assume Min Steel")>
    <DataMember()> Public Property assume_min_steel() As Boolean?
        Get
            Return Me._assume_min_steel
        End Get
        Set
            Me._assume_min_steel = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Check Shear Along Depth")>
    <DataMember()> Public Property check_shear_along_depth() As Boolean?
        Get
            Return Me._check_shear_along_depth
        End Get
        Set
            Me._check_shear_along_depth = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Utilize Shear Friction Methodology")>
    <DataMember()> Public Property utilize_shear_friction_methodology() As Boolean?
        Get
            Return Me._utilize_shear_friction_methodology
        End Get
        Set
            Me._utilize_shear_friction_methodology = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Embedded Pole")>
    <DataMember()> Public Property embedded_pole() As Boolean?
        Get
            Return Me._embedded_pole
        End Get
        Set
            Me._embedded_pole = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Belled Pier")>
    <DataMember()> Public Property belled_pier() As Boolean?
        Get
            Return Me._belled_pier
        End Get
        Set
            Me._belled_pier = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Concrete Compressive Strength")>
    <DataMember()> Public Property concrete_compressive_strength() As Double?
        Get
            Return Me._concrete_compressive_strength
        End Get
        Set
            Me._concrete_compressive_strength = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Longitudinal Rebar Yield Strength")>
    <DataMember()> Public Property longitudinal_rebar_yield_strength() As Double?
        Get
            Return Me._longitudinal_rebar_yield_strength
        End Get
        Set
            Me._longitudinal_rebar_yield_strength = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Rebar Cage 2 Fy Override")>
    <DataMember()> Public Property rebar_cage_2_fy_override() As Double?
        Get
            Return Me._rebar_cage_2_fy_override
        End Get
        Set
            Me._rebar_cage_2_fy_override = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Rebar Cage 3 Fy Override")>
    <DataMember()> Public Property rebar_cage_3_fy_override() As Double?
        Get
            Return Me._rebar_cage_3_fy_override
        End Get
        Set
            Me._rebar_cage_3_fy_override = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Rebar Effective Depths")>
    <DataMember()> Public Property rebar_effective_depths() As Boolean?
        Get
            Return Me._rebar_effective_depths
        End Get
        Set
            Me._rebar_effective_depths = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Shear Crit Depth Override Comp")>
    <DataMember()> Public Property shear_crit_depth_override_comp() As Double?
        Get
            Return Me._shear_crit_depth_override_comp
        End Get
        Set
            Me._shear_crit_depth_override_comp = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Shear Crit Depth Override Uplift")>
    <DataMember()> Public Property shear_crit_depth_override_uplift() As Double?
        Get
            Return Me._shear_crit_depth_override_uplift
        End Get
        Set
            Me._shear_crit_depth_override_uplift = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Shear Override Crit Depth")>
    <DataMember()> Public Property shear_override_crit_depth() As Boolean?
        Get
            Return Me._shear_override_crit_depth
        End Get
        Set
            Me._shear_override_crit_depth = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Tie Yield Strength")>
    <DataMember()> Public Property tie_yield_strength() As Double?
        Get
            Return Me._tie_yield_strength
        End Get
        Set
            Me._tie_yield_strength = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Tool Version")>
    <DataMember()> Public Property tool_version() As String
        Get
            Return Me._tool_version
        End Get
        Set
            Me._tool_version = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Ultimate Bearing")>
    <DataMember()> Public Property ultimate_gross_bearing() As Boolean?
        Get
            Return Me._ultimate_gross_bearing
        End Get
        Set
            Me._ultimate_gross_bearing = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Local Pier Profile ID")>
    <DataMember()> Public Property local_pier_profile_id As Integer?
        Get
            Return _local_pier_profile_id
        End Get
        Set()
            Me._local_pier_profile_id = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Local Drilled Pier ID")>
    <DataMember()> Public Property local_drilled_pier_id As Integer?
        Get
            Return _local_drilled_pier_id
        End Get
        Set()
            Me._local_drilled_pier_id = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Drilled Pier ID")>
    <DataMember()> Public Property drilled_pier_id As Integer?
        Get
            Return _drilled_pier_id
        End Get
        Set()
            Me._drilled_pier_id = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()

    End Sub

    Public Sub New(ByVal dr As DataRow, Optional ByVal Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        Me.foundation_depth = DBtoNullableDbl(dr.Item("foundation_depth"))
        Me.extension_above_grade = DBtoNullableDbl(dr.Item("extension_above_grade"))
        Me.assume_min_steel = DBtoNullableBool(dr.Item("assume_min_steel"))
        Me.check_shear_along_depth = DBtoNullableBool(dr.Item("check_shear_along_depth"))
        Me.utilize_shear_friction_methodology = DBtoNullableBool(dr.Item("utilize_shear_friction_methodology"))
        Me.embedded_pole = DBtoNullableBool(dr.Item("embedded_pole"))
        Me.belled_pier = DBtoNullableBool(dr.Item("belled_pier"))
        Me.concrete_compressive_strength = DBtoNullableDbl(dr.Item("concrete_compressive_strength"))
        Me.longitudinal_rebar_yield_strength = DBtoNullableDbl(dr.Item("longitudinal_rebar_yield_strength"))
        Me.rebar_cage_2_fy_override = DBtoNullableDbl(dr.Item("rebar_cage_2_fy_override"))
        Me.rebar_cage_3_fy_override = DBtoNullableDbl(dr.Item("rebar_cage_3_fy_override"))
        Me.rebar_effective_depths = DBtoNullableBool(dr.Item("rebar_effective_depths"))
        Me.shear_crit_depth_override_comp = DBtoNullableDbl(dr.Item("shear_crit_depth_override_comp"))
        Me.shear_crit_depth_override_uplift = DBtoNullableDbl(dr.Item("shear_crit_depth_override_uplift"))
        Me.shear_override_crit_depth = DBtoNullableBool(dr.Item("shear_override_crit_depth"))
        Me.tie_yield_strength = DBtoNullableDbl(dr.Item("tie_yield_strength"))
        Me.tool_version = DBtoStr(dr.Item("tool_version"))

        Dim bearing_option As String = ""
        Try
            bearing_option = dr.Item("ultimate_gross_bearing").ToString
        Catch ex As Exception
        End Try

        If bearing_option = "Ult. Gross Bearing Capacity (ksf)" Then
            Me.ultimate_gross_bearing = True
        ElseIf bearing_option.ToUpper = "TRUE" Or bearing_option.ToUpper = "FALSE" Or bearing_option = "" Then
            Me.ultimate_gross_bearing = DBtoNullableBool(dr.Item("ultimate_gross_bearing"))
        ElseIf bearing_option = "Ult. Net Bearing Capacity (ksf)" Then
            Me.ultimate_gross_bearing = False
        End If

        Try
            Me.local_pier_profile_id = DBtoNullableInt(dr.Item("local_pier_profile_id"))
        Catch
        End Try
        Try
            Me.local_drilled_pier_id = DBtoNullableInt(dr.Item("local_drilled_pier_id"))
        Catch
        End Try
        Try
            Me.drilled_pier_id = DBtoNullableInt(dr.Item("drilled_pier_id"))
        Catch
        End Try
    End Sub
#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.foundation_depth.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.extension_above_grade.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.assume_min_steel.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.check_shear_along_depth.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.utilize_shear_friction_methodology.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.embedded_pole.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.belled_pier.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.concrete_compressive_strength.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.longitudinal_rebar_yield_strength.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rebar_cage_2_fy_override.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rebar_cage_3_fy_override.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rebar_effective_depths.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.shear_crit_depth_override_comp.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.shear_crit_depth_override_uplift.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.shear_override_crit_depth.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.tie_yield_strength.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.tool_version.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ultimate_gross_bearing.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.drilled_pier_id.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString("@TopLevelID")
        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = SQLInsertFields.AddtoDBString("foundation_depth")
        SQLInsertFields = SQLInsertFields.AddtoDBString("extension_above_grade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("assume_min_steel")
        SQLInsertFields = SQLInsertFields.AddtoDBString("check_shear_along_depth")
        SQLInsertFields = SQLInsertFields.AddtoDBString("utilize_shear_friction_methodology")
        SQLInsertFields = SQLInsertFields.AddtoDBString("embedded_pole")
        SQLInsertFields = SQLInsertFields.AddtoDBString("belled_pier")
        SQLInsertFields = SQLInsertFields.AddtoDBString("concrete_compressive_strength")
        SQLInsertFields = SQLInsertFields.AddtoDBString("longitudinal_rebar_yield_strength")
        SQLInsertFields = SQLInsertFields.AddtoDBString("rebar_cage_2_fy_override")
        SQLInsertFields = SQLInsertFields.AddtoDBString("rebar_cage_3_fy_override")
        SQLInsertFields = SQLInsertFields.AddtoDBString("rebar_effective_depths")
        SQLInsertFields = SQLInsertFields.AddtoDBString("shear_crit_depth_override_comp")
        SQLInsertFields = SQLInsertFields.AddtoDBString("shear_crit_depth_override_uplift")
        SQLInsertFields = SQLInsertFields.AddtoDBString("shear_override_crit_depth")
        SQLInsertFields = SQLInsertFields.AddtoDBString("tie_yield_strength")
        SQLInsertFields = SQLInsertFields.AddtoDBString("tool_version")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ultimate_gross_bearing")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("drilled_pier_id")
        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("foundation_depth = " & Me.foundation_depth.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("extension_above_grade = " & Me.extension_above_grade.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("assume_min_steel = " & Me.assume_min_steel.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("check_shear_along_depth = " & Me.check_shear_along_depth.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("utilize_shear_friction_methodology = " & Me.utilize_shear_friction_methodology.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("embedded_pole = " & Me.embedded_pole.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("belled_pier = " & Me.belled_pier.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("concrete_compressive_strength = " & Me.concrete_compressive_strength.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("longitudinal_rebar_yield_strength = " & Me.longitudinal_rebar_yield_strength.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("rebar_cage_2_fy_override = " & Me.rebar_cage_2_fy_override.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("rebar_cage_3_fy_override = " & Me.rebar_cage_3_fy_override.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("rebar_effective_depths = " & Me.rebar_effective_depths.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("shear_crit_depth_override_comp = " & Me.shear_crit_depth_override_comp.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("shear_crit_depth_override_uplift = " & Me.shear_crit_depth_override_uplift.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("shear_override_crit_depth = " & Me.shear_override_crit_depth.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("tie_yield_strength = " & Me.tie_yield_strength.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("tool_version = " & Me.tool_version.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ultimate_gross_bearing = " & Me.ultimate_gross_bearing.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("drilled_pier_id = " & Me.drilled_pier_id.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("drilled_pier_id = " & "@TopLevelID")
        Return SQLUpdateFieldsandValues
    End Function
#End Region

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As DrilledPierProfile = TryCast(other, DrilledPierProfile)
        If otherToCompare Is Nothing Then Return False
        Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
        Equals = If(Me.foundation_depth.CheckChange(otherToCompare.foundation_depth, changes, categoryName, "Foundation Depth"), Equals, False)
        Equals = If(Me.extension_above_grade.CheckChange(otherToCompare.extension_above_grade, changes, categoryName, "Extension Above Grade"), Equals, False)
        Equals = If(Me.assume_min_steel.CheckChange(otherToCompare.assume_min_steel, changes, categoryName, "Assume Min Steel"), Equals, False)
        Equals = If(Me.check_shear_along_depth.CheckChange(otherToCompare.check_shear_along_depth, changes, categoryName, "Check Shear Along Depth"), Equals, False)
        Equals = If(Me.utilize_shear_friction_methodology.CheckChange(otherToCompare.utilize_shear_friction_methodology, changes, categoryName, "Utilize Shear Friction Methodology"), Equals, False)
        Equals = If(Me.embedded_pole.CheckChange(otherToCompare.embedded_pole, changes, categoryName, "Embedded Pole"), Equals, False)
        Equals = If(Me.belled_pier.CheckChange(otherToCompare.belled_pier, changes, categoryName, "Belled Pier"), Equals, False)
        Equals = If(Me.concrete_compressive_strength.CheckChange(otherToCompare.concrete_compressive_strength, changes, categoryName, "Concrete Compressive Strength"), Equals, False)
        Equals = If(Me.longitudinal_rebar_yield_strength.CheckChange(otherToCompare.longitudinal_rebar_yield_strength, changes, categoryName, "Longitudinal Rebar Yield Strength"), Equals, False)
        Equals = If(Me.rebar_cage_2_fy_override.CheckChange(otherToCompare.rebar_cage_2_fy_override, changes, categoryName, "Rebar Cage 2 Fy Override"), Equals, False)
        Equals = If(Me.rebar_cage_3_fy_override.CheckChange(otherToCompare.rebar_cage_3_fy_override, changes, categoryName, "Rebar Cage 3 Fy Override"), Equals, False)
        Equals = If(Me.rebar_effective_depths.CheckChange(otherToCompare.rebar_effective_depths, changes, categoryName, "Rebar Effective Depths"), Equals, False)
        Equals = If(Me.shear_crit_depth_override_comp.CheckChange(otherToCompare.shear_crit_depth_override_comp, changes, categoryName, "Shear Crit Depth Override Comp"), Equals, False)
        Equals = If(Me.shear_crit_depth_override_uplift.CheckChange(otherToCompare.shear_crit_depth_override_uplift, changes, categoryName, "Shear Crit Depth Override Uplift"), Equals, False)
        Equals = If(Me.shear_override_crit_depth.CheckChange(otherToCompare.shear_override_crit_depth, changes, categoryName, "Shear Override Crit Depth"), Equals, False)
        Equals = If(Me.tie_yield_strength.CheckChange(otherToCompare.tie_yield_strength, changes, categoryName, "Tie Yield Strength"), Equals, False)
        Equals = If(Me.tool_version.CheckChange(otherToCompare.tool_version, changes, categoryName, "Tool Version"), Equals, False)
        Equals = If(Me.ultimate_gross_bearing.CheckChange(otherToCompare.ultimate_gross_bearing, changes, categoryName, "Ultimate Bearing"), Equals, False)
        'Equals = If(Me.drilled_pier_id.CheckChange(otherToCompare.drilled_pier_id, changes, categoryName, "Drilled Pier Foreign Key"), Equals, False)

        'Sections
        Equals = If(Me.Sections.CheckChange(otherToCompare.Sections, changes, categoryName, "Drilled Pier Sections"), Equals, False)

        Return Equals

    End Function
End Class
<DataContractAttribute()>
<KnownType(GetType(DrilledPierSection))>
Partial Public Class DrilledPierSection
    Inherits EDSObjectWithQueries

#Region "Inheritted"
    '''Must override these inherited properties
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Drilled Pier Section"
        End Get
    End Property

    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "fnd.drilled_pier_section"
        End Get
    End Property

    Private _Insert As String
    Private _Update As String
    Private _Delete As String

    Public Overrides Function SQLInsert() As String
        SQLInsert = ""
        SQLInsert = CCI_Engineering_Templates.My.Resources.Drilled_Pier_Section__INSERT ' QueryBuilderFromFile(queryPath & "Drilled Pier\Drilled Pier Section (INSERT).sql")
        SQLInsert = SQLInsert.Replace("[TABLE NAME]", Me.EDSTableName)
        SQLInsert = SQLInsert.Replace("[INSERT FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.Replace("[INSERT VALUES]", Me.SQLInsertValues)

        Dim _rebarInsert As String

        For Each bar In Me.Rebar
            'If bar.ID IsNot Nothing And bar.ID > 0 Then
            '    If bar.local_rebar_id IsNot Nothing Then
            '        _rebarInsert += bar.SQLUpdate + vbCrLf
            '    Else
            '        _rebarInsert += bar.SQLDelete + vbCrLf
            '    End If
            'Else
            _rebarInsert += bar.SQLInsert + vbCrLf
            'End If
        Next

        SQLInsert = SQLInsert.Replace("--[REBAR INSERT]", _rebarInsert)
        SQLInsert = SQLInsert.TrimEnd

        Return SQLInsert
    End Function

    Public Overrides Function SQLUpdate() As String
        SQLUpdate = ""
        SQLUpdate = CCI_Engineering_Templates.My.Resources.General__UPDATE ' QueryBuilderFromFile(queryPath & "Drilled Pier\General (UPDATE).sql")
        SQLUpdate = SQLUpdate.Replace("[TABLE NAME]", Me.EDSTableName)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID & vbCrLf & "SET @SubLevel3ID = " & Me.ID)

        Dim _rebarInsert As String

        For Each bar In Me.Rebar
            If bar.ID IsNot Nothing And bar.ID > 0 Then
                If bar.local_rebar_id IsNot Nothing Then
                    _rebarInsert += bar.SQLUpdate + vbCrLf
                Else
                    _rebarInsert += bar.SQLDelete + vbCrLf
                End If
            Else
                _rebarInsert += bar.SQLInsert + vbCrLf
            End If
        Next

        SQLUpdate = SQLUpdate.Replace("--[OPTIONAL]", _rebarInsert)
        SQLUpdate = SQLUpdate.TrimEnd

        Return SQLUpdate
    End Function

    Public Overrides Function SQLDelete() As String
        SQLDelete = CCI_Engineering_Templates.My.Resources.General__DELETE
        SQLDelete = SQLDelete.Replace("[TABLE]", Me.EDSTableName)
        SQLDelete = SQLDelete.Replace("[ID]", Me.ID)

        Return SQLDelete
    End Function

#End Region

#Region "Define"
    <DataMember()> Public Property Rebar As New List(Of DrilledPierRebar)


    Private _pier_diameter As Double?
    Private _clear_cover As Double?
    Private _clear_cover_rebar_cage_option As Boolean?
    Private _tie_size As Integer?
    Private _tie_spacing As Double?
    Private _bottom_elevation As Double?
    Private _local_section_id As Integer?
    Private _rho_override As Double?
    Private _drilled_pier_id As Integer?
    Private _local_drilled_pier_id As Integer?
    Private _local_pier_profile_id As Integer?
    Private _pier_profile_id As Integer?


    <Category("Drilled Pier"), Description(""), DisplayName("Pier Diameter")>
    <DataMember()> Public Property pier_diameter() As Double?
        Get
            Return Me._pier_diameter
        End Get
        Set
            Me._pier_diameter = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Clear Cover")>
    <DataMember()> Public Property clear_cover() As Double?
        Get
            Return Me._clear_cover
        End Get
        Set
            Me._clear_cover = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Clear Cover Rebar Cage Option")>
    <DataMember()> Public Property clear_cover_rebar_cage_option() As Boolean?
        Get
            Return Me._clear_cover_rebar_cage_option
        End Get
        Set
            Me._clear_cover_rebar_cage_option = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Tie Size")>
    <DataMember()> Public Property tie_size() As Integer?
        Get
            Return Me._tie_size
        End Get
        Set
            Me._tie_size = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Tie Spacing")>
    <DataMember()> Public Property tie_spacing() As Double?
        Get
            Return Me._tie_spacing
        End Get
        Set
            Me._tie_spacing = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Bottom Elevation")>
    <DataMember()> Public Property bottom_elevation() As Double?
        Get
            Return Me._bottom_elevation
        End Get
        Set
            Me._bottom_elevation = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Local Section Id")>
    <DataMember()> Public Property local_section_id() As Integer?
        Get
            Return Me._local_section_id
        End Get
        Set
            Me._local_section_id = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Rho Override")>
    <DataMember()> Public Property rho_override() As Double?
        Get
            Return Me._rho_override
        End Get
        Set
            Me._rho_override = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Local Pier Profile ID")>
    <DataMember()> Public Property local_pier_profile_id As Integer?
        Get
            Return _local_pier_profile_id
        End Get
        Set()
            Me._local_pier_profile_id = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Local Drilled Pier ID")>
    <DataMember()> Public Property local_drilled_pier_id As Integer?
        Get
            Return _local_drilled_pier_id
        End Get
        Set()
            Me._local_drilled_pier_id = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Drilled Pier ID")>
    <DataMember()> Public Property drilled_pier_id As Integer?
        Get
            Return _drilled_pier_id
        End Get
        Set()
            Me._drilled_pier_id = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Pier Profile ID")>
    <DataMember()> Public Property pier_profile_id As Integer?
        Get
            Return _pier_profile_id
        End Get
        Set()
            Me._pier_profile_id = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()

    End Sub

    Public Sub New(ByVal dr As DataRow, Optional ByVal Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        Me.pier_diameter = DBtoNullableDbl(dr.Item("pier_diameter"))
        Me.clear_cover = DBtoNullableDbl(dr.Item("clear_cover"))

        Dim clear_cover_option As String
        Try
            clear_cover_option = dr.Item("clear_cover_rebar_cage_option").ToString
        Catch ex As Exception
            clear_cover_option = ""
        End Try

        If clear_cover_option = "Clear Cover to Ties" Then
            Me.clear_cover_rebar_cage_option = True
        ElseIf clear_cover_option.ToUpper = "TRUE" Or clear_cover_option.ToUpper = "FALSE" Or clear_cover_option = "" Then
            Me.clear_cover_rebar_cage_option = DBtoNullableBool(dr.Item("clear_cover_rebar_cage_option"))
        ElseIf clear_cover_option = "Rebar Cage Diameter" Then
            Me.clear_cover_rebar_cage_option = False
        End If

        Me.tie_size = DBtoNullableInt(dr.Item("tie_size"))
        Me.tie_spacing = DBtoNullableDbl(dr.Item("tie_spacing"))
        Me.bottom_elevation = DBtoNullableDbl(dr.Item("bottom_elevation"))
        Me.local_section_id = DBtoNullableInt(dr.Item("local_section_id"))
        Me.rho_override = DBtoNullableDbl(dr.Item("rho_override"))

        Try
            Me.local_pier_profile_id = DBtoNullableInt(dr.Item("local_pier_profile_id"))
        Catch
        End Try
        Try
            Me.local_drilled_pier_id = DBtoNullableInt(dr.Item("local_drilled_pier_id"))
        Catch
        End Try
        Try
            Me.drilled_pier_id = DBtoNullableInt(dr.Item("drilled_pier_id"))
        Catch
        End Try
        Try
            Me.pier_profile_id = DBtoNullableInt(dr.Item("pier_profile_id"))
        Catch
        End Try
    End Sub
#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel2ID") 'Pier_profile_id
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pier_diameter.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.clear_cover.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.clear_cover_rebar_cage_option.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.tie_size.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.tie_spacing.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bottom_elevation.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_section_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rho_override.ToString.FormatDBValue)
        Return SQLInsertValues

    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = SQLInsertFields.AddtoDBString("pier_profile_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pier_diameter")
        SQLInsertFields = SQLInsertFields.AddtoDBString("clear_cover")
        SQLInsertFields = SQLInsertFields.AddtoDBString("clear_cover_rebar_cage_option")
        SQLInsertFields = SQLInsertFields.AddtoDBString("tie_size")
        SQLInsertFields = SQLInsertFields.AddtoDBString("tie_spacing")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bottom_elevation")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_section_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("rho_override")
        Return SQLInsertFields

    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pier_profile_id = " & "@SubLevel2ID")
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pier_diameter = " & Me.pier_diameter.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("clear_cover = " & Me.clear_cover.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("clear_cover_rebar_cage_option = " & Me.clear_cover_rebar_cage_option.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("tie_size = " & Me.tie_size.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("tie_spacing = " & Me.tie_spacing.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bottom_elevation = " & Me.bottom_elevation.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_section_id = " & Me.local_section_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("rho_override = " & Me.rho_override.ToString.FormatDBValue)
        Return SQLUpdateFieldsandValues

    End Function
#End Region

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As DrilledPierSection = TryCast(other, DrilledPierSection)
        If otherToCompare Is Nothing Then Return False
        Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
        Equals = If(Me.pier_profile_id.CheckChange(otherToCompare.pier_profile_id, changes, categoryName, "Drilled Pier Profile Id"), Equals, False)
        Equals = If(Me.pier_diameter.CheckChange(otherToCompare.pier_diameter, changes, categoryName, "Pier Diameter"), Equals, False)
        Equals = If(Me.clear_cover.CheckChange(otherToCompare.clear_cover, changes, categoryName, "Clear Cover"), Equals, False)
        Equals = If(Me.clear_cover_rebar_cage_option.CheckChange(otherToCompare.clear_cover_rebar_cage_option, changes, categoryName, "Clear Cover Rebar Cage Option"), Equals, False)
        Equals = If(Me.tie_size.CheckChange(otherToCompare.tie_size, changes, categoryName, "Tie Size"), Equals, False)
        Equals = If(Me.tie_spacing.CheckChange(otherToCompare.tie_spacing, changes, categoryName, "Tie Spacing"), Equals, False)
        Equals = If(Me.bottom_elevation.CheckChange(otherToCompare.bottom_elevation, changes, categoryName, "Bottom Elevation"), Equals, False)
        Equals = If(Me.local_section_id.CheckChange(otherToCompare.local_section_id, changes, categoryName, "Local Section Id"), Equals, False)
        Equals = If(Me.rho_override.CheckChange(otherToCompare.rho_override, changes, categoryName, "Rho Override"), Equals, False)

        'Drilled Pier Rebar
        Equals = If(Me.Rebar.CheckChange(otherToCompare.Rebar, changes, categoryName, "Drilled Pier Rebar"), Equals, False)


        Return Equals

    End Function
End Class
<DataContractAttribute()>
<KnownType(GetType(DrilledPierRebar))>
Partial Public Class DrilledPierRebar
    Inherits EDSObjectWithQueries

#Region "Inheritted"
    '''Must override these inherited properties
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Drilled Pier Rebar"
        End Get
    End Property

    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "fnd.drilled_pier_rebar"
        End Get
    End Property

    Private _Insert As String
    Private _Update As String
    Private _Delete As String

    Public Overrides Function SQLInsert() As String
        SQLInsert = ""
        SQLInsert = CCI_Engineering_Templates.My.Resources.General__INSERT ' QueryBuilderFromFile(queryPath & "Drilled Pier\General (INSERT).sql")
        SQLInsert = SQLInsert.Replace("[TABLE NAME]", Me.EDSTableName)
        SQLInsert = SQLInsert.Replace("[INSERT FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.Replace("[INSERT VALUES]", Me.SQLInsertValues)

        Return SQLInsert
    End Function

    Public Overrides Function SQLUpdate() As String
        SQLUpdate = ""
        SQLUpdate = CCI_Engineering_Templates.My.Resources.General__UPDATE ' QueryBuilderFromFile(queryPath & "Drilled Pier\General (UPDATE).sql")
        SQLUpdate = SQLUpdate.Replace("[TABLE NAME]", Me.EDSTableName)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID)

        Return SQLUpdate
    End Function

    Public Overrides Function SQLDelete() As String
        SQLDelete = CCI_Engineering_Templates.My.Resources.General__DELETE
        SQLDelete = SQLDelete.Replace("[TABLE]", Me.EDSTableName)
        SQLDelete = SQLDelete.Replace("[ID]", Me.ID)

        Return SQLDelete
    End Function

#End Region

#Region "Define"
    Public local_section_id As Integer?


    Private _section_id As Integer?
    Private _longitudinal_rebar_quantity As Integer?
    Private _longitudinal_rebar_size As Integer?
    Private _longitudinal_rebar_cage_diameter As Double?
    Private _local_rebar_id As Integer?
    Private _drilled_pier_id As Integer?
    Private _local_drilled_pier_id As Integer?
    Private _local_pier_profile_id As Integer?
    Private _pier_profile_id As Integer?


    <Category("Drilled Pier"), Description(""), DisplayName("Section Id")>
    <DataMember()> Public Property section_id() As Integer?
        Get
            Return Me._section_id
        End Get
        Set
            Me._section_id = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Longitudinal Rebar Quantity")>
    <DataMember()> Public Property longitudinal_rebar_quantity() As Integer?
        Get
            Return Me._longitudinal_rebar_quantity
        End Get
        Set
            Me._longitudinal_rebar_quantity = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Longitudinal Rebar Size")>
    <DataMember()> Public Property longitudinal_rebar_size() As Integer?
        Get
            Return Me._longitudinal_rebar_size
        End Get
        Set
            Me._longitudinal_rebar_size = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Longitudinal Rebar Cage Diameter")>
    <DataMember()> Public Property longitudinal_rebar_cage_diameter() As Double?
        Get
            Return Me._longitudinal_rebar_cage_diameter
        End Get
        Set
            Me._longitudinal_rebar_cage_diameter = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Local Rebar Id")>
    <DataMember()> Public Property local_rebar_id() As Integer?
        Get
            Return Me._local_rebar_id
        End Get
        Set
            Me._local_rebar_id = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Local Pier Profile ID")>
    <DataMember()> Public Property local_pier_profile_id As Integer?
        Get
            Return _local_pier_profile_id
        End Get
        Set()
            Me._local_pier_profile_id = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Local Drilled Pier ID")>
    <DataMember()> Public Property local_drilled_pier_id As Integer?
        Get
            Return _local_drilled_pier_id
        End Get
        Set()
            Me._local_drilled_pier_id = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Drilled Pier ID")>
    <DataMember()> Public Property drilled_pier_id As Integer?
        Get
            Return _drilled_pier_id
        End Get
        Set()
            Me._drilled_pier_id = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Pier Profile ID")>
    <DataMember()> Public Property pier_profile_id As Integer?
        Get
            Return _pier_profile_id
        End Get
        Set()
            Me._pier_profile_id = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()

    End Sub

    Public Sub New(ByVal dr As DataRow, Optional ByVal Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        Me.section_id = DBtoNullableInt(dr.Item("section_id"))
        Me.longitudinal_rebar_quantity = DBtoNullableInt(dr.Item("longitudinal_rebar_quantity"))
        Me.longitudinal_rebar_size = DBtoNullableInt(dr.Item("longitudinal_rebar_size"))
        Me.longitudinal_rebar_cage_diameter = DBtoNullableDbl(dr.Item("longitudinal_rebar_cage_diameter"))
        Me.local_rebar_id = DBtoNullableInt(dr.Item("local_rebar_id"))

        Try
            Me.local_section_id = DBtoNullableInt(dr.Item("local_section_id"))
        Catch
        End Try
        Try
            Me.local_pier_profile_id = DBtoNullableInt(dr.Item("local_pier_profile_id"))
        Catch
        End Try
        Try
            Me.local_drilled_pier_id = DBtoNullableInt(dr.Item("local_drilled_pier_id"))
        Catch
        End Try
        Try
            Me.drilled_pier_id = DBtoNullableInt(dr.Item("drilled_pier_id"))
        Catch
        End Try
        Try
            Me.pier_profile_id = DBtoNullableInt(dr.Item("pier_profile_id"))
        Catch
        End Try

    End Sub
#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel3ID")
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.longitudinal_rebar_quantity.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.longitudinal_rebar_size.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.longitudinal_rebar_cage_diameter.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_rebar_id.ToString.FormatDBValue)
        Return SQLInsertValues

    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = SQLInsertFields.AddtoDBString("section_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("longitudinal_rebar_quantity")
        SQLInsertFields = SQLInsertFields.AddtoDBString("longitudinal_rebar_size")
        SQLInsertFields = SQLInsertFields.AddtoDBString("longitudinal_rebar_cage_diameter")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_rebar_id")
        Return SQLInsertFields

    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("section_id = " & "@SubLevel3ID")
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("longitudinal_rebar_quantity = " & Me.longitudinal_rebar_quantity.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("longitudinal_rebar_size = " & Me.longitudinal_rebar_size.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("longitudinal_rebar_cage_diameter = " & Me.longitudinal_rebar_cage_diameter.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_rebar_id = " & Me.local_rebar_id.ToString.FormatDBValue)
        Return SQLUpdateFieldsandValues

    End Function
#End Region

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As DrilledPierRebar = TryCast(other, DrilledPierRebar)
        If otherToCompare Is Nothing Then Return False
        Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
        Equals = If(Me.section_id.CheckChange(otherToCompare.section_id, changes, categoryName, "Section Id"), Equals, False)
        Equals = If(Me.longitudinal_rebar_quantity.CheckChange(otherToCompare.longitudinal_rebar_quantity, changes, categoryName, "Longitudinal Rebar Quantity"), Equals, False)
        Equals = If(Me.longitudinal_rebar_size.CheckChange(otherToCompare.longitudinal_rebar_size, changes, categoryName, "Longitudinal Rebar Size"), Equals, False)
        Equals = If(Me.longitudinal_rebar_cage_diameter.CheckChange(otherToCompare.longitudinal_rebar_cage_diameter, changes, categoryName, "Longitudinal Rebar Cage Diameter"), Equals, False)
        Equals = If(Me.local_rebar_id.CheckChange(otherToCompare.local_rebar_id, changes, categoryName, "Local Rebar Id"), Equals, False)
        Return Equals

    End Function
End Class
<DataContractAttribute()>
<KnownType(GetType(EmbeddedPole))>
Partial Public Class EmbeddedPole
    Inherits EDSObjectWithQueries

#Region "Inheritted"
    '''Must override these inherited properties
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Embedded Pole"
        End Get
    End Property

    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "fnd.embedded_pole"
        End Get
    End Property

    Private _Insert As String
    Private _Update As String
    Private _Delete As String

    Public Overrides Function SQLInsert() As String
        SQLInsert = ""
        SQLInsert = CCI_Engineering_Templates.My.Resources.General__INSERT 'QueryBuilderFromFile(queryPath & "Drilled Pier\General (INSERT).sql")
        SQLInsert = SQLInsert.Replace("[TABLE NAME]", Me.EDSTableName)
        SQLInsert = SQLInsert.Replace("[INSERT FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.Replace("[INSERT VALUES]", Me.SQLInsertValues)

        Return SQLInsert
    End Function

    Public Overrides Function SQLUpdate() As String
        SQLUpdate = ""
        SQLUpdate = CCI_Engineering_Templates.My.Resources.General__UPDATE ' QueryBuilderFromFile(queryPath & "Drilled Pier\General (UPDATE).sql")
        SQLUpdate = SQLUpdate.Replace("[TABLE NAME]", Me.EDSTableName)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID)

        Return SQLUpdate
    End Function

    Public Overrides Function SQLDelete() As String
        SQLDelete = CCI_Engineering_Templates.My.Resources.General__DELETE
        SQLDelete = SQLDelete.Replace("[TABLE]", Me.EDSTableName)
        SQLDelete = SQLDelete.Replace("[ID]", Me.ID)

        Return SQLDelete
    End Function

#End Region

#Region "Define"
    Private _embedded_pole_option As Boolean?
    Private _encased_in_concrete As Boolean?
    Private _pole_side_quantity As Integer?
    Private _pole_yield_strength As Double?
    Private _pole_thickness As Double?
    Private _embedded_pole_input_type As String
    Private _pole_diameter_toc As Double?
    Private _pole_top_diameter As Double?
    Private _pole_bottom_diameter As Double?
    Private _pole_section_length As Double?
    Private _pole_taper_factor As Double?
    Private _pole_bend_radius_override As Double?
    Private _drilled_pier_id As Integer?
    Private _local_drilled_pier_id As Integer?
    Private _local_pier_profile_id As Integer?
    Private _pier_profile_id As Integer?

    <Category("Drilled Pier"), Description(""), DisplayName("Embedded Pole Option")>
    <DataMember()> Public Property embedded_pole_option() As Boolean?
        Get
            Return Me._embedded_pole_option
        End Get
        Set
            Me._embedded_pole_option = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Encased In Concrete")>
    <DataMember()> Public Property encased_in_concrete() As Boolean?
        Get
            Return Me._encased_in_concrete
        End Get
        Set
            Me._encased_in_concrete = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Pole Side Quantity")>
    <DataMember()> Public Property pole_side_quantity() As Integer?
        Get
            Return Me._pole_side_quantity
        End Get
        Set
            Me._pole_side_quantity = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Pole Yield Strength")>
    <DataMember()> Public Property pole_yield_strength() As Double?
        Get
            Return Me._pole_yield_strength
        End Get
        Set
            Me._pole_yield_strength = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Pole Thickness")>
    <DataMember()> Public Property pole_thickness() As Double?
        Get
            Return Me._pole_thickness
        End Get
        Set
            Me._pole_thickness = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Embedded Pole Input Type")>
    <DataMember()> Public Property embedded_pole_input_type() As String
        Get
            Return Me._embedded_pole_input_type
        End Get
        Set
            Me._embedded_pole_input_type = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Pole Diameter Toc")>
    <DataMember()> Public Property pole_diameter_toc() As Double?
        Get
            Return Me._pole_diameter_toc
        End Get
        Set
            Me._pole_diameter_toc = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Pole Top Diameter")>
    <DataMember()> Public Property pole_top_diameter() As Double?
        Get
            Return Me._pole_top_diameter
        End Get
        Set
            Me._pole_top_diameter = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Pole Bottom Diameter")>
    <DataMember()> Public Property pole_bottom_diameter() As Double?
        Get
            Return Me._pole_bottom_diameter
        End Get
        Set
            Me._pole_bottom_diameter = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Pole Section Length")>
    <DataMember()> Public Property pole_section_length() As Double?
        Get
            Return Me._pole_section_length
        End Get
        Set
            Me._pole_section_length = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Pole Taper Factor")>
    <DataMember()> Public Property pole_taper_factor() As Double?
        Get
            Return Me._pole_taper_factor
        End Get
        Set
            Me._pole_taper_factor = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Pole Bend Radius Override")>
    <DataMember()> Public Property pole_bend_radius_override() As Double?
        Get
            Return Me._pole_bend_radius_override
        End Get
        Set
            Me._pole_bend_radius_override = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Local Pier Profile ID")>
    <DataMember()> Public Property local_pier_profile_id As Integer?
        Get
            Return _local_pier_profile_id
        End Get
        Set()
            Me._local_pier_profile_id = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Local Drilled Pier ID")>
    <DataMember()> Public Property local_drilled_pier_id As Integer?
        Get
            Return _local_drilled_pier_id
        End Get
        Set()
            Me._local_drilled_pier_id = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Drilled Pier ID")>
    <DataMember()> Public Property drilled_pier_id As Integer?
        Get
            Return _drilled_pier_id
        End Get
        Set()
            Me._drilled_pier_id = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Pier Profile ID")>
    <DataMember()> Public Property pier_profile_id As Integer?
        Get
            Return _pier_profile_id
        End Get
        Set()
            Me._pier_profile_id = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()

    End Sub

    Public Sub New(ByVal dr As DataRow, Optional ByVal Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        Me.embedded_pole_option = DBtoNullableBool(dr.Item("embedded_pole_option"))
        Me.encased_in_concrete = DBtoNullableBool(dr.Item("encased_in_concrete"))
        Me.pole_side_quantity = DBtoNullableInt(dr.Item("pole_side_quantity"))
        Me.pole_yield_strength = DBtoNullableDbl(dr.Item("pole_yield_strength"))
        Me.pole_thickness = DBtoNullableDbl(dr.Item("pole_thickness"))
        Me.embedded_pole_input_type = DBtoStr(dr.Item("embedded_pole_input_type"))
        Me.pole_diameter_toc = DBtoNullableDbl(dr.Item("pole_diameter_toc"))
        Me.pole_top_diameter = DBtoNullableDbl(dr.Item("pole_top_diameter"))
        Me.pole_bottom_diameter = DBtoNullableDbl(dr.Item("pole_bottom_diameter"))
        Me.pole_section_length = DBtoNullableDbl(dr.Item("pole_section_length"))
        Me.pole_taper_factor = DBtoNullableDbl(dr.Item("pole_taper_factor"))
        Me.pole_bend_radius_override = DBtoNullableDbl(dr.Item("pole_bend_radius_override"))
        Try
            Me.local_pier_profile_id = DBtoNullableInt(dr.Item("local_pier_profile_id"))
        Catch
        End Try
        Try
            Me.local_drilled_pier_id = DBtoNullableInt(dr.Item("local_drilled_pier_id"))
        Catch
        End Try
        Try
            Me.drilled_pier_id = DBtoNullableInt(dr.Item("drilled_pier_id"))
        Catch
        End Try
        Try
            Me.pier_profile_id = DBtoNullableInt(dr.Item("pier_profile_id"))
        Catch
        End Try
    End Sub
#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel2ID")
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.embedded_pole_option.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.encased_in_concrete.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pole_side_quantity.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pole_yield_strength.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pole_thickness.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.embedded_pole_input_type.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pole_diameter_toc.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pole_top_diameter.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pole_bottom_diameter.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pole_section_length.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pole_taper_factor.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pole_bend_radius_override.ToString.FormatDBValue)
        Return SQLInsertValues

    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = SQLInsertFields.AddtoDBString("pier_profile_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("embedded_pole_option")
        SQLInsertFields = SQLInsertFields.AddtoDBString("encased_in_concrete")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pole_side_quantity")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pole_yield_strength")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pole_thickness")
        SQLInsertFields = SQLInsertFields.AddtoDBString("embedded_pole_input_type")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pole_diameter_toc")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pole_top_diameter")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pole_bottom_diameter")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pole_section_length")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pole_taper_factor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pole_bend_radius_override")
        Return SQLInsertFields

    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pier_profile_id = " & "@SubLevel2ID")
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("embedded_pole_option = " & Me.embedded_pole_option.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("encased_in_concrete = " & Me.encased_in_concrete.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pole_side_quantity = " & Me.pole_side_quantity.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pole_yield_strength = " & Me.pole_yield_strength.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pole_thickness = " & Me.pole_thickness.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("embedded_pole_input_type = " & Me.embedded_pole_input_type.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pole_diameter_toc = " & Me.pole_diameter_toc.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pole_top_diameter = " & Me.pole_top_diameter.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pole_bottom_diameter = " & Me.pole_bottom_diameter.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pole_section_length = " & Me.pole_section_length.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pole_taper_factor = " & Me.pole_taper_factor.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pole_bend_radius_override = " & Me.pole_bend_radius_override.ToString.FormatDBValue)
        Return SQLUpdateFieldsandValues

    End Function
#End Region

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As EmbeddedPole = TryCast(other, EmbeddedPole)
        If otherToCompare Is Nothing Then Return False
        Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
        Equals = If(Me.pier_profile_id.CheckChange(otherToCompare.pier_profile_id, changes, categoryName, "Drilled Pier Profile Id"), Equals, False)
        Equals = If(Me.embedded_pole_option.CheckChange(otherToCompare.embedded_pole_option, changes, categoryName, "Embedded Pole Option"), Equals, False)
        Equals = If(Me.encased_in_concrete.CheckChange(otherToCompare.encased_in_concrete, changes, categoryName, "Encased In Concrete"), Equals, False)
        Equals = If(Me.pole_side_quantity.CheckChange(otherToCompare.pole_side_quantity, changes, categoryName, "Pole Side Quantity"), Equals, False)
        Equals = If(Me.pole_yield_strength.CheckChange(otherToCompare.pole_yield_strength, changes, categoryName, "Pole Yield Strength"), Equals, False)
        Equals = If(Me.pole_thickness.CheckChange(otherToCompare.pole_thickness, changes, categoryName, "Pole Thickness"), Equals, False)
        Equals = If(Me.embedded_pole_input_type.CheckChange(otherToCompare.embedded_pole_input_type, changes, categoryName, "Embedded Pole Input Type"), Equals, False)
        Equals = If(Me.pole_diameter_toc.CheckChange(otherToCompare.pole_diameter_toc, changes, categoryName, "Pole Diameter Toc"), Equals, False)
        Equals = If(Me.pole_top_diameter.CheckChange(otherToCompare.pole_top_diameter, changes, categoryName, "Pole Top Diameter"), Equals, False)
        Equals = If(Me.pole_bottom_diameter.CheckChange(otherToCompare.pole_bottom_diameter, changes, categoryName, "Pole Bottom Diameter"), Equals, False)
        Equals = If(Me.pole_section_length.CheckChange(otherToCompare.pole_section_length, changes, categoryName, "Pole Section Length"), Equals, False)
        Equals = If(Me.pole_taper_factor.CheckChange(otherToCompare.pole_taper_factor, changes, categoryName, "Pole Taper Factor"), Equals, False)
        Equals = If(Me.pole_bend_radius_override.CheckChange(otherToCompare.pole_bend_radius_override, changes, categoryName, "Pole Bend Radius Override"), Equals, False)
        Return Equals

    End Function
End Class
<DataContractAttribute()>
<KnownType(GetType(BelledPier))>
Partial Public Class BelledPier
    Inherits EDSObjectWithQueries

#Region "Inheritted"
    '''Must override these inherited properties
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Belled Pier"
        End Get
    End Property

    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "fnd.belled_pier"
        End Get
    End Property

    Private _Insert As String
    Private _Update As String
    Private _Delete As String

    Public Overrides Function SQLInsert() As String
        SQLInsert = ""
        SQLInsert = CCI_Engineering_Templates.My.Resources.General__INSERT ' QueryBuilderFromFile(queryPath & "Drilled Pier\General (INSERT).sql")
        SQLInsert = SQLInsert.Replace("[TABLE NAME]", Me.EDSTableName)
        SQLInsert = SQLInsert.Replace("[INSERT FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.Replace("[INSERT VALUES]", Me.SQLInsertValues)

        Return SQLInsert
    End Function

    Public Overrides Function SQLUpdate() As String
        SQLUpdate = ""
        SQLUpdate = CCI_Engineering_Templates.My.Resources.General__UPDATE ' QueryBuilderFromFile(queryPath & "Drilled Pier\General (UPDATE).sql")
        SQLUpdate = SQLUpdate.Replace("[TABLE NAME]", Me.EDSTableName)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID)

        Return SQLUpdate
    End Function

    Public Overrides Function SQLDelete() As String
        SQLDelete = CCI_Engineering_Templates.My.Resources.General__DELETE
        SQLDelete = SQLDelete.Replace("[TABLE]", Me.EDSTableName)
        SQLDelete = SQLDelete.Replace("[ID]", Me.ID)

        Return SQLDelete
    End Function

#End Region

#Region "Define"
    Private _belled_pier_option As Boolean?
    Private _bottom_diameter_of_bell As Double?
    Private _bell_input_type As String
    Private _bell_angle As Double?
    Private _bell_height As Double?
    Private _bell_toe_height As Double?
    Private _neglect_top_soil_layer As Boolean?
    Private _swelling_expansive_soil As Boolean?
    Private _depth_of_expansive_soil As Double?
    Private _expansive_soil_force As Double?
    Private _drilled_pier_id As Integer?
    Private _local_drilled_pier_id As Integer?
    Private _local_pier_profile_id As Integer?
    Private _pier_profile_id As Integer?

    <Category("Drilled Pier"), Description(""), DisplayName("Belled Pier Option")>
    <DataMember()> Public Property belled_pier_option() As Boolean?
        Get
            Return Me._belled_pier_option
        End Get
        Set
            Me._belled_pier_option = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Bottom Diameter Of Bell")>
    <DataMember()> Public Property bottom_diameter_of_bell() As Double?
        Get
            Return Me._bottom_diameter_of_bell
        End Get
        Set
            Me._bottom_diameter_of_bell = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Bell Input Type")>
    <DataMember()> Public Property bell_input_type() As String
        Get
            Return Me._bell_input_type
        End Get
        Set
            Me._bell_input_type = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Bell Angle")>
    <DataMember()> Public Property bell_angle() As Double?
        Get
            Return Me._bell_angle
        End Get
        Set
            Me._bell_angle = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Bell Height")>
    <DataMember()> Public Property bell_height() As Double?
        Get
            Return Me._bell_height
        End Get
        Set
            Me._bell_height = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Bell Toe Height")>
    <DataMember()> Public Property bell_toe_height() As Double?
        Get
            Return Me._bell_toe_height
        End Get
        Set
            Me._bell_toe_height = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Neglect Top Soil Layer")>
    <DataMember()> Public Property neglect_top_soil_layer() As Boolean?
        Get
            Return Me._neglect_top_soil_layer
        End Get
        Set
            Me._neglect_top_soil_layer = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Swelling Expansive Soil")>
    <DataMember()> Public Property swelling_expansive_soil() As Boolean?
        Get
            Return Me._swelling_expansive_soil
        End Get
        Set
            Me._swelling_expansive_soil = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Depth Of Expansive Soil")>
    <DataMember()> Public Property depth_of_expansive_soil() As Double?
        Get
            Return Me._depth_of_expansive_soil
        End Get
        Set
            Me._depth_of_expansive_soil = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Expansive Soil Force")>
    <DataMember()> Public Property expansive_soil_force() As Double?
        Get
            Return Me._expansive_soil_force
        End Get
        Set
            Me._expansive_soil_force = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Local Pier Profile ID")>
    <DataMember()> Public Property local_pier_profile_id As Integer?
        Get
            Return _local_pier_profile_id
        End Get
        Set()
            Me._local_pier_profile_id = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Local Drilled Pier ID")>
    <DataMember()> Public Property local_drilled_pier_id As Integer?
        Get
            Return _local_drilled_pier_id
        End Get
        Set()
            Me._local_drilled_pier_id = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Drilled Pier ID")>
    <DataMember()> Public Property drilled_pier_id As Integer?
        Get
            Return _drilled_pier_id
        End Get
        Set()
            Me._drilled_pier_id = Value
        End Set
    End Property
    <Category("Drilled Pier"), Description(""), DisplayName("Pier Profile ID")>
    <DataMember()> Public Property pier_profile_id As Integer?
        Get
            Return _pier_profile_id
        End Get
        Set()
            Me._pier_profile_id = Value
        End Set
    End Property
#End Region

#Region "Constructors"
    Public Sub New()

    End Sub

    Public Sub New(ByVal dr As DataRow, Optional ByVal Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        Me.belled_pier_option = DBtoNullableBool(dr.Item("belled_pier_option"))
        Me.bottom_diameter_of_bell = DBtoNullableDbl(dr.Item("bottom_diameter_of_bell"))
        Me.bell_input_type = DBtoStr(dr.Item("bell_input_type"))
        Me.bell_angle = DBtoNullableDbl(dr.Item("bell_angle"))
        Me.bell_height = DBtoNullableDbl(dr.Item("bell_height"))
        Me.bell_toe_height = DBtoNullableDbl(dr.Item("bell_toe_height"))
        Me.neglect_top_soil_layer = DBtoNullableBool(dr.Item("neglect_top_soil_layer"))
        Me.swelling_expansive_soil = DBtoNullableBool(dr.Item("swelling_expansive_soil"))
        Me.depth_of_expansive_soil = DBtoNullableDbl(dr.Item("depth_of_expansive_soil"))
        Me.expansive_soil_force = DBtoNullableDbl(dr.Item("expansive_soil_force"))
        Try
            Me.local_pier_profile_id = DBtoNullableInt(dr.Item("local_pier_profile_id"))
        Catch
        End Try
        Try
            Me.local_drilled_pier_id = DBtoNullableInt(dr.Item("local_drilled_pier_id"))
        Catch
        End Try
        Try
            Me.drilled_pier_id = DBtoNullableInt(dr.Item("drilled_pier_id"))
        Catch
        End Try
        Try
            Me.pier_profile_id = DBtoNullableInt(dr.Item("pier_profile_id"))
        Catch
        End Try
    End Sub
#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel2ID")
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.belled_pier_option.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bottom_diameter_of_bell.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bell_input_type.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bell_angle.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bell_height.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bell_toe_height.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.neglect_top_soil_layer.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.swelling_expansive_soil.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.depth_of_expansive_soil.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.expansive_soil_force.ToString.FormatDBValue)
        Return SQLInsertValues

    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = SQLInsertFields.AddtoDBString("pier_profile_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("belled_pier_option")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bottom_diameter_of_bell")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bell_input_type")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bell_angle")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bell_height")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bell_toe_height")
        SQLInsertFields = SQLInsertFields.AddtoDBString("neglect_top_soil_layer")
        SQLInsertFields = SQLInsertFields.AddtoDBString("swelling_expansive_soil")
        SQLInsertFields = SQLInsertFields.AddtoDBString("depth_of_expansive_soil")
        SQLInsertFields = SQLInsertFields.AddtoDBString("expansive_soil_force")
        Return SQLInsertFields

    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pier_profile_id = " & "@SubLevel2ID")
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("belled_pier_option = " & Me.belled_pier_option.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bottom_diameter_of_bell = " & Me.bottom_diameter_of_bell.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bell_input_type = " & Me.bell_input_type.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bell_angle = " & Me.bell_angle.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bell_height = " & Me.bell_height.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bell_toe_height = " & Me.bell_toe_height.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("neglect_top_soil_layer = " & Me.neglect_top_soil_layer.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("swelling_expansive_soil = " & Me.swelling_expansive_soil.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("depth_of_expansive_soil = " & Me.depth_of_expansive_soil.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("expansive_soil_force = " & Me.expansive_soil_force.ToString.FormatDBValue)
        Return SQLUpdateFieldsandValues

    End Function
#End Region

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As BelledPier = TryCast(other, BelledPier)
        If otherToCompare Is Nothing Then Return False
        Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
        Equals = If(Me.pier_profile_id.CheckChange(otherToCompare.pier_profile_id, changes, categoryName, "Drilled Pier Profile Id"), Equals, False)
        Equals = If(Me.belled_pier_option.CheckChange(otherToCompare.belled_pier_option, changes, categoryName, "Belled Pier Option"), Equals, False)
        Equals = If(Me.bottom_diameter_of_bell.CheckChange(otherToCompare.bottom_diameter_of_bell, changes, categoryName, "Bottom Diameter Of Bell"), Equals, False)
        Equals = If(Me.bell_input_type.CheckChange(otherToCompare.bell_input_type, changes, categoryName, "Bell Input Type"), Equals, False)
        Equals = If(Me.bell_angle.CheckChange(otherToCompare.bell_angle, changes, categoryName, "Bell Angle"), Equals, False)
        Equals = If(Me.bell_height.CheckChange(otherToCompare.bell_height, changes, categoryName, "Bell Height"), Equals, False)
        Equals = If(Me.bell_toe_height.CheckChange(otherToCompare.bell_toe_height, changes, categoryName, "Bell Toe Height"), Equals, False)
        Equals = If(Me.neglect_top_soil_layer.CheckChange(otherToCompare.neglect_top_soil_layer, changes, categoryName, "Neglect Top Soil Layer"), Equals, False)
        Equals = If(Me.swelling_expansive_soil.CheckChange(otherToCompare.swelling_expansive_soil, changes, categoryName, "Swelling Expansive Soil"), Equals, False)
        Equals = If(Me.depth_of_expansive_soil.CheckChange(otherToCompare.depth_of_expansive_soil, changes, categoryName, "Depth Of Expansive Soil"), Equals, False)
        Equals = If(Me.expansive_soil_force.CheckChange(otherToCompare.expansive_soil_force, changes, categoryName, "Expansive Soil Force"), Equals, False)
        Return Equals

    End Function
End Class
<DataContractAttribute()>
<KnownType(GetType(DrilledPierSoilProfile))>
Public Class DrilledPierSoilProfile
    Inherits SoilProfile
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Soil Profiles"
        End Get
    End Property
    <DataMember()> Public Property local_soil_profile_id As Integer?
    <DataMember()> Public Property local_drilled_pier_id As Integer?
    <DataMember()> Public Property drilled_pier_id As Integer?
    <DataMember()> Public Property DPSoilLayers As New List(Of DrilledPierSoilLayer)

    Public Sub New()

    End Sub

    Public Sub New(ByVal dr As DataRow, Optional ByVal Parent As EDSObject = Nothing)
        ConstructMe(dr, Parent)

        Try
            Me.local_soil_profile_id = DBtoNullableDbl(dr.Item("local_soil_profile_id"))
            Me.local_drilled_pier_id = DBtoNullableDbl(dr.Item("local_drilled_pier_id"))
            Me.drilled_pier_id = DBtoNullableDbl(dr.Item("drilled_pier_id"))
        Catch
        End Try
    End Sub

    Public Overrides Function SQLUpdate() As String

        'SQLUpdate = QueryBuilderFromFile(queryPath & "Soil Profile\Soil Profile (UPDATE).sql")
        SQLUpdate = CCI_Engineering_Templates.My.Resources.Soil_Profile_UPDATE
        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
        SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

        Dim slUp As String

        For Each sl In Me.DPSoilLayers
            Dim dpsl As DrilledPierSoilLayer = TryCast(sl, DrilledPierSoilLayer)

            If dpsl.ID IsNot Nothing And dpsl?.ID > 0 Then
                If dpsl.local_soil_layer_id IsNot Nothing Then
                    slUp += dpsl.SQLUpdate
                Else
                    slUp += dpsl.SQLDelete
                End If
            Else
                slUp += dpsl.SQLInsert
            End If

            slUp += vbCrLf
        Next

        SQLUpdate += vbCrLf + slUp

        Return SQLUpdate
    End Function

End Class
<DataContractAttribute()>
<KnownType(GetType(DrilledPierSoilLayer))>
Public Class DrilledPierSoilLayer
    Inherits SoilLayer

    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Soil Layers"
        End Get
    End Property

    <DataMember()> Public Property local_soil_profile_id As Integer?
    <DataMember()> Public Property soil_profile_id As Integer?
    <DataMember()> Public Property local_soil_layer_id As Integer?
    <DataMember()> Public Property drilled_pier_id As Integer?
    <DataMember()> Public Property local_drilled_pier_id As Integer?

    Public Sub New()

    End Sub

    Public Sub New(ByVal dr As DataRow, Optional ByVal Parent As EDSObject = Nothing)
        ConstructMe(dr, Parent)

        Try
            Me.local_soil_profile_id = DBtoNullableDbl(dr.Item("local_soil_profile"))
        Catch ex As Exception
        End Try
        Try
            Me.local_soil_layer_id = DBtoNullableDbl(dr.Item("local_soil_layer_id"))
        Catch ex As Exception
        End Try
        Try
            Me.drilled_pier_id = DBtoNullableDbl(dr.Item("drilled_pier_id"))
        Catch ex As Exception
        End Try
        Try
            Me.local_drilled_pier_id = DBtoNullableDbl(dr.Item("local_drilled_pier_id"))
        Catch ex As Exception
        End Try
        Try
            Me.soil_profile_id = DBtoNullableDbl(dr.Item("soil_profile_id"))
        Catch
        End Try
    End Sub
End Class
<DataContractAttribute()>
<KnownType(GetType(DrilledPierResult))>
Public Class DrilledPierResult
    Inherits EDSResult

    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Drilled Pier Result"
        End Get
    End Property

    Public Overloads ReadOnly Property EDSTableName As String
        Get
            Return "fnd.drilled_pier_results"
        End Get
    End Property

    Public Function InsertCustom() As String
        InsertCustom = "BEGIN" & vbCrLf &
                "  INSERT INTO " & Me.EDSTableName & "(" & Me.SQLInsertFields & ")" & vbCrLf &
                "  VALUES([VALUES])" & vbCrLf &
                "END" & vbCrLf
        InsertCustom = InsertCustom.Replace("[VALUES]", Me.SQLInsertValuesExtended(Me.Parent.ID))
        InsertCustom = InsertCustom.Replace("[FIELDS]", Me.SQLInsertFieldsExtended)
        InsertCustom = InsertCustom.TrimEnd()
        Return InsertCustom()
    End Function

    Public Sub New()

    End Sub

    Public Sub New(ByVal resultDr As DataRow, ByVal Parent As EDSObjectWithQueries)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then
            Me.Absorb(Parent)
        End If

        Me.foreign_key = Parent?.ID
        Me.result_lkup = DBtoStr(resultDr.Item("result_lkup"))
        Me.rating = DBtoNullableDbl(resultDr.Item("rating"), 10)
        Me.ForeignKeyName = "drilled_pier_id"
        Me.EDSTableDepth = 0
    End Sub

    Public Function SQLInsertValuesExtended(Optional ByVal parentID As Integer? = Nothing) As String
        Dim sqlValues As String = SQLInsertValues(parentID)

        sqlValues = sqlValues.AddtoDBString(Me.work_order_seq_num.FormatDBValue)
        sqlValues = sqlValues.AddtoDBString(If(parentID Is Nothing, "@TopLevel", Me.foreign_key.ToString.FormatDBValue))
        sqlValues = sqlValues.AddtoDBString(Me.result_lkup.FormatDBValue)
        sqlValues = sqlValues.AddtoDBString(Math.Round(CDbl(Me.rating), 4).ToString.FormatDBValue)
        sqlValues = sqlValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
        sqlValues = sqlValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)
        sqlValues = sqlValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
        sqlValues = sqlValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)

        Return SQLInsertValuesExtended
    End Function

    Public Function SQLInsertFieldsExtended(Optional ByVal parentID As Integer? = Nothing) As String
        Dim sqlFields As String = SQLInsertFields()

        sqlFields = sqlFields.AddtoDBString("work_order_seq_num")
        sqlFields = sqlFields.AddtoDBString("drilled_pier")
        sqlFields = sqlFields.AddtoDBString("result_lkup")
        sqlFields = sqlFields.AddtoDBString("rating")
        sqlFields = sqlFields.AddtoDBString("modified_person_id")
        sqlFields = sqlFields.AddtoDBString("process_stage")
        sqlFields = sqlFields.AddtoDBString("modified_person_id")
        sqlFields = sqlFields.AddtoDBString("process_stage")

        Return SQLInsertFieldsExtended
    End Function
End Class