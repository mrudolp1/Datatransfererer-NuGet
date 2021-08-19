Option Strict Off

Imports DevExpress.Spreadsheet
Imports System.Security.Principal
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop

Partial Public Class DataTransfererDrilledPier

#Region "Define"
    Private NewDrilledPierWb As New Workbook
    Private prop_ExcelFilePath As String

    Public Property DrilledPiers As New List(Of DrilledPier)
    Private Property DrilledPierTemplatePath As String = "C:\Users\" & Environment.UserName & "\Desktop\Drilled Pier Foundation (5.1.0) - TEMPLATE - 8-19-2021.xlsm"
    Private Property DrilledPierFileType As DocumentFormat = DocumentFormat.Xlsm

    'Public Property dpDS As New DataSet
    'Public Property ds As New DataSet
    Public Property dpDB As String
    Public Property dpID As WindowsIdentity

    Public Property ExcelFilePath() As String
        Get
            Return Me.prop_ExcelFilePath
        End Get
        Set
            Me.prop_ExcelFilePath = Value
        End Set
    End Property

    Public Property xlApp As Object
#End Region

#Region "Constructors"
    Public Sub New()
        'Leave method empty
    End Sub

    Public Sub New(ByVal MyDataSet As DataSet, ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String, ByVal BU As String, ByVal Strucutre_ID As String)
        'dpDS = MyDataSet
        ds = MyDataSet
        dpID = LogOnUser
        dpDB = ActiveDatabase
        'BUNumber = BU 'Need to turn back on when connecting to dashboard. Turned off for testing.
        'STR_ID = Strucutre_ID 'Need to turn back on when connecting to dashboard. Turned off for testing.
    End Sub
#End Region

#Region "Load Data"
    Public Function LoadFromEDS() As Boolean
        Dim refid As Integer

        Dim DrilledPierLoader As String

        'Load data to get pier and pad details data for the existing structure model
        For Each item As SQLParameter In DrilledPierSQLDataTables()
            DrilledPierLoader = QueryBuilderFromFile(queryPath & "Drilled Pier\" & item.sqlQuery).Replace("[EXISTING MODEL]", GetExistingModelQuery())
            'DoDaSQL.sqlLoader(DrilledPierLoader, item.sqlDatatable, dpDS, dpDB, dpID, "0")
            DoDaSQL.sqlLoader(DrilledPierLoader, item.sqlDatatable, ds, dpDB, dpID, "0")
            'If dpDS.Tables(item.sqlDatatable).Rows.Count = 0 Then Return False 'This may need adjusted since some tables can be empty
            'If ds.Tables(item.sqlDatatable).Rows.Count = 0 Then Return False 'This may need adjusted since some tables can be empty
        Next

        'Custom Section to transfer data for the drilled pier tool. Needs to be adjusted for each tool.
        'For Each DrilledPierDataRow As DataRow In dpDS.Tables("Drilled Pier General Details SQL").Rows
        For Each DrilledPierDataRow As DataRow In ds.Tables("Drilled Pier General Details SQL").Rows
            refid = CType(DrilledPierDataRow.Item("drilled_pier_id"), Integer)

            DrilledPiers.Add(New DrilledPier(DrilledPierDataRow, refid))
        Next

        Return True
    End Function 'Create Drilled Pier objects based on what is saved in EDS

    Public Sub LoadFromExcel()
        Dim refID As Integer
        Dim refCol As String

        For Each item As EXCELDTParameter In DrilledPierExcelDTParameters()
            'Get tables from excel file 
            ds.Tables.Add(ExcelDatasourceToDataTable(GetExcelDataSource(ExcelFilePath, item.xlsSheet, item.xlsRange), item.xlsDatatable))
        Next

        'Custom Section to transfer data for the drilled pier tool. Needs to be adjusted for each tool.
        For Each DrilledPierDataRow As DataRow In ds.Tables("Drilled Pier General Details EXCEL").Rows
            'If DrilledPierDataRow.Item("foudation_id").ToString = "" Then
            '    refCol = "local_drilled_pier_id"
            '    refID = CType(DrilledPierDataRow.Item(refCol), Integer)
            'Else
            '    refCol = "drilled_pier_id"
            '    refID = CType(DrilledPierDataRow.Item(refCol), Integer)
            'End If
            ''commented out in case drilled pier id and local drilled pier id matched, prevents possible overriding of data
            refCol = "local_drilled_pier_id"
            refID = CType(DrilledPierDataRow.Item(refCol), Integer)

            DrilledPiers.Add(New DrilledPier(DrilledPierDataRow, refID, refCol))
        Next
    End Sub 'Create Drilled Pier objects based on what is coming from the excel file
#End Region

#Region "Save Data"
    Public Sub SaveToEDS()
        Dim firstOne As Boolean = True
        Dim mySoils As String = ""
        Dim mySections As String = ""
        Dim myRebar As String = ""
        Dim myProfiles As String = ""

        For Each dp As DrilledPier In DrilledPiers
            Dim DrilledPierSaver As String = QueryBuilderFromFile(queryPath & "Drilled Pier\Drilled Piers (IN_UP).sql")
            Dim dpSectionQuery As String = QueryBuilderFromFile(queryPath & "Drilled Pier\Drilled Piers Section (IN_UP).txt")

            DrilledPierSaver = DrilledPierSaver.Replace("[BU NUMBER]", BUNumber)
            DrilledPierSaver = DrilledPierSaver.Replace("[STRUCTURE ID]", STR_ID)
            DrilledPierSaver = DrilledPierSaver.Replace("[FOUNDATION TYPE]", "Drilled Pier")
            If dp.pier_id = 0 Or IsDBNull(dp.pier_id) Then
                DrilledPierSaver = DrilledPierSaver.Replace("'[DRILLED PIER ID]'", "NULL")
            Else
                DrilledPierSaver = DrilledPierSaver.Replace("'[DRILLED PIER ID]'", dp.pier_id.ToString)
            End If
            DrilledPierSaver = DrilledPierSaver.Replace("[EMBED BOOLEAN]", dp.embedded_pole.ToString)
            DrilledPierSaver = DrilledPierSaver.Replace("[BELL BOOLEAN]", dp.belled_pier.ToString)
            DrilledPierSaver = DrilledPierSaver.Replace("[INSERT ALL PIER DETAILS]", InsertDrilledPierDetail(dp))

            If dp.pier_id = 0 Or IsDBNull(dp.pier_id) Then
                For Each dpsl As DrilledPierSoilLayer In dp.soil_layers
                    Dim tempSoilLayer As String = InsertDrilledPierSoilLayer(dpsl)

                    If Not firstOne Then
                        mySoils += ",(" & tempSoilLayer & ")"
                    Else
                        mySoils += "(" & tempSoilLayer & ")"
                    End If

                    firstOne = False
                Next 'Add Soil Layer INSERT statments
                DrilledPierSaver = DrilledPierSaver.Replace("([INSERT ALL SOIL LAYERS])", mySoils)
                firstOne = True

                For Each dpsec As DrilledPierSection In dp.sections
                    Dim tempSection As String = dpSectionQuery.Replace("[DRILLED PIER SECTION]", InsertDrilledPierSection(dpsec))

                    For Each dpreb In dpsec.rebar
                        Dim temprebar As String = InsertDrilledPierRebar(dpreb)

                        If Not firstOne Then
                            myRebar += ",(" & temprebar & ")"
                        Else
                            myRebar += "(" & temprebar & ")"
                        End If

                        firstOne = False
                    Next 'Add Rebar INSERT Statements

                    tempSection = tempSection.Replace("([DRILLED PIER SECTION REBAR])", myRebar)
                    firstOne = True
                    myRebar = ""
                    mySections += tempSection + vbNewLine
                Next 'Add Section INSERT Statements
                DrilledPierSaver = DrilledPierSaver.Replace("--*[DRILLED PIER SECTIONS]*--", mySections)

                If dp.belled_pier Then
                    DrilledPierSaver = DrilledPierSaver.Replace("[INSERT ALL BELLED PIER DETAILS]", InsertDrilledPierBell(dp.belled_details))
                Else
                    DrilledPierSaver = DrilledPierSaver.Replace("BEGIN --Belled Pier", "--BEGIN --Belled Pier")
                    DrilledPierSaver = DrilledPierSaver.Replace("IF @IsBelled = 'True'", "--IF @IsBelled = 'True'")
                    DrilledPierSaver = DrilledPierSaver.Replace("INSERT INTO belled_pier_details VALUES ([INSERT ALL BELLED PIER DETAILS])", "")
                    DrilledPierSaver = DrilledPierSaver.Replace("END --INSERT Belled Pier information if required", "--END --INSERT Belled Pier information if required")
                End If 'Add Belled Pier INSERT Statment

                If dp.embedded_pole Then
                    DrilledPierSaver = DrilledPierSaver.Replace("[INSERT ALL EMBEDDED POLE DETAILS]", InsertDrilledPierEmbed(dp.embed_details))
                Else
                    DrilledPierSaver = DrilledPierSaver.Replace("BEGIN --Embedded Pole", "--BEGIN --Embedded Pole")
                    DrilledPierSaver = DrilledPierSaver.Replace("IF @IsEmbed = 'True'", "--IF @IsEmbed = 'True'")
                    DrilledPierSaver = DrilledPierSaver.Replace("INSERT INTO embedded_pole_details OUTPUT INSERTED.ID INTO @EmbeddedPole VALUES ([INSERT ALL EMBEDDED POLE DETAILS])", "")
                    DrilledPierSaver = DrilledPierSaver.Replace("SELECT @EmbedID=EmbedID FROM @EmbeddedPole", "--SELECT @EmbedID=EmbedID FROM @EmbeddedPole")
                    DrilledPierSaver = DrilledPierSaver.Replace("END --INSERT Embedded Pole information if required", "--END --INSERT Embedded Pole information if required")
                End If 'Add Embedded Pole INSERT Statment

                For Each dpp As DrilledPierProfile In dp.drilled_pier_profiles
                    Dim tempDrilledPierProfile As String = InsertDrilledPierProfile(dpp)

                    If Not firstOne Then
                        myProfiles += ",(" & tempDrilledPierProfile & ")"
                    Else
                        myProfiles += "(" & tempDrilledPierProfile & ")"
                    End If

                    firstOne = False
                Next 'Add Pier Profile INSERT statements
                DrilledPierSaver = DrilledPierSaver.Replace("([INSERT ALL PIER PROFILES])", myProfiles)
                firstOne = True

                mySoils = ""
                mySections = ""
                myProfiles = ""
            Else
                Dim tempUpdater As String = ""
                tempUpdater += UpdateDrilledPierDetail(dp)

                'comment out soil layer insertion. Added in next step if a layer does not have an ID
                DrilledPierSaver = DrilledPierSaver.Replace("INSERT INTO drilled_pier_soil_layer VALUES ([INSERT ALL SOIL LAYERS])", "--INSERT INTO drilled_pier_soil_layer VALUES ([INSERT ALL SOIL LAYERS])")

                For Each dpsl As DrilledPierSoilLayer In dp.soil_layers
                    If dpsl.soil_layer_id = 0 Or IsDBNull(dpsl.soil_layer_id) Then
                        tempUpdater += "INSERT INTO drilled_pier_soil_layers VALUES (" & InsertDrilledPierSoilLayer(dpsl) & ") " & vbNewLine
                    Else
                        tempUpdater += UpdateDrilledPierSoilLayer(dpsl)
                    End If
                Next

                If dp.belled_pier Then
                    If dp.belled_details.belled_pier_id = 0 Or IsDBNull(dp.belled_details.belled_pier_id) Then
                        tempUpdater += "INSERT INTO belled_pier_details VALUES (" & InsertDrilledPierBell(dp.belled_details) & ") " & vbNewLine
                    Else
                        tempUpdater += UpdateDrilledPierBell(dp.belled_details)
                    End If
                Else
                    DrilledPierSaver = DrilledPierSaver.Replace("BEGIN --Belled Pier", "--BEGIN --Belled Pier")
                    DrilledPierSaver = DrilledPierSaver.Replace("IF @IsBelled = 'True'", "--IF @IsBelled = 'True'")
                    DrilledPierSaver = DrilledPierSaver.Replace("INSERT INTO belled_pier_details VALUES ([INSERT ALL BELLED PIER DETAILS])", "")
                    DrilledPierSaver = DrilledPierSaver.Replace("END --INSERT Belled Pier information if required", "--END --INSERT Belled Pier information if required")
                End If

                If dp.embedded_pole Then
                    If dp.embed_details.embedded_id = 0 Or IsDBNull(dp.embed_details.embedded_id) Then
                        tempUpdater += "BEGIN INSERT INTO embedded_pole_details OUTPUT INSERTED.ID INTO @EmbeddedPole VALUES (" & InsertDrilledPierEmbed(dp.embed_details) & ") " & vbNewLine & " SELECT @EmbedID=EmbedID FROM @EmbeddedPole"
                        tempUpdater += " END " & vbNewLine
                    Else
                        tempUpdater += UpdateDrilledPierEmbed(dp.embed_details)
                    End If
                Else
                    DrilledPierSaver = DrilledPierSaver.Replace("BEGIN --Embedded Pole", "--BEGIN --Embedded Pole")
                    DrilledPierSaver = DrilledPierSaver.Replace("IF @IsEmbed = 'True'", "--IF @IsEmbed = 'True'")
                    DrilledPierSaver = DrilledPierSaver.Replace("INSERT INTO embedded_pole_details OUTPUT INSERTED.ID INTO @EmbeddedPole VALUES ([INSERT ALL EMBEDDED POLE DETAILS])", "")
                    DrilledPierSaver = DrilledPierSaver.Replace("SELECT @EmbedID=EmbedID FROM @EmbeddedPole", "--SELECT @EmbedID=EmbedID FROM @EmbeddedPole")
                    DrilledPierSaver = DrilledPierSaver.Replace("END --INSERT Embedded Pole information if required", "--END --INSERT Embedded Pole information if required")
                End If

                For Each dpSec As DrilledPierSection In dp.sections
                    If dpSec.section_id = 0 Or IsDBNull(dpSec.section_id) Then
                        tempUpdater += "BEGIN INSERT INTO drilled_pier_section OUTPUT INSERTED.ID INTO @DrilledPierSection VALUES (" & InsertDrilledPierSection(dpSec) & ") " & vbNewLine & " SELECT @SecID=SecID FROM @DrilledPierSection"
                        For Each dpreb As DrilledPierRebar In dpSec.rebar
                            tempUpdater += "INSERT INTO drilled_pier_rebar VALUES (" & InsertDrilledPierRebar(dpreb) & ") " & vbNewLine
                        Next
                        tempUpdater += " END " & vbNewLine
                    Else
                        tempUpdater += UpdateDrilledPierSection(dpSec)
                        For Each dpreb As DrilledPierRebar In dpSec.rebar
                            If dpreb.rebar_id = 0 Or IsDBNull(dpreb.rebar_id) Then
                                tempUpdater += "INSERT INTO drilled_pier_rebar VALUES (" & InsertDrilledPierRebar(dpreb).Replace("@SecID", dpSec.section_id.ToString) & ") " & vbNewLine
                            Else
                                tempUpdater += UpdateDrilledPierRebar(dpreb)
                            End If
                        Next
                    End If
                Next

                DrilledPierSaver = DrilledPierSaver.Replace("INSERT INTO drilled_pier_profile VALUES ([INSERT ALL PIER PROFILES])", "--INSERT INTO drilled_pier_profile VALUES ([INSERT ALL PIER PROFILES])")
                For Each dpp As DrilledPierProfile In dp.drilled_pier_profiles
                    If dpp.profile_id = 0 Or IsDBNull(dpp.profile_id) Then
                        tempUpdater += "INSERT INTO drilled_pier_profile VALUES (" & InsertDrilledPierProfile(dpp) & ") " & vbNewLine
                    Else
                        tempUpdater += UpdateDrilledPierProfile(dpp)
                    End If
                Next

                DrilledPierSaver = DrilledPierSaver.Replace("SELECT * FROM TEMPORARY", tempUpdater)
            End If

            DrilledPierSaver = DrilledPierSaver.Replace("[INSERT ALL PIER DETAILS DETAILS]", InsertDrilledPierDetail(dp))

            sqlSender(DrilledPierSaver, dpDB, dpID, "0")
        Next


    End Sub

    Public Sub SaveToExcel()
        Dim dpRow As Integer = 3
        Dim secRow As Integer = 3
        Dim rebRow As Integer = 3
        Dim soilRow As Integer = 3
        Dim profileRow As Integer = 3

        LoadNewDrilledPier()

        With NewDrilledPierWb
            'For Each dp As DrilledPier In DrilledPiers
            '    If Not IsNothing(dp.local_drilled_pier_id) Then
            '        .Worksheets("Details (RETURN)").Range("A" & dpRow).Value = CType(dp.local_drilled_pier_id, Integer)
            '    Else .Worksheets("Details (RETURN)").Range("A" & dpRow).ClearContents
            '    End If
            '    .Worksheets("Details (RETURN)").Range("B" & dpRow).Value = dp.pier_id
            '    If Not IsNothing(dp.foundation_depth) Then
            '        .Worksheets("Details (RETURN)").Range("C" & dpRow).Value = CType(dp.foundation_depth, Double)
            '    Else .Worksheets("Details (RETURN)").Range("C" & dpRow).ClearContents
            '    End If
            '    If Not IsNothing(dp.extension_above_grade) Then
            '        .Worksheets("Details (RETURN)").Range("D" & dpRow).Value = CType(dp.extension_above_grade, Double)
            '    Else .Worksheets("Details (RETURN)").Range("D" & dpRow).ClearContents
            '    End If
            '    If Not IsNothing(dp.groundwater_depth) Then
            '        .Worksheets("Details (RETURN)").Range("E" & dpRow).Value = CType(dp.groundwater_depth, Double)
            '    Else .Worksheets("Details (RETURN)").Range("E" & dpRow).ClearContents
            '    End If
            '    .Worksheets("Details (RETURN)").Range("F" & dpRow).Value = dp.assume_min_steel
            '    .Worksheets("Details (RETURN)").Range("G" & dpRow).Value = dp.check_shear_along_depth
            '    .Worksheets("Details (RETURN)").Range("H" & dpRow).Value = dp.utilize_shear_friction_methodology
            '    .Worksheets("Details (RETURN)").Range("I" & dpRow).Value = dp.embedded_pole
            '    .Worksheets("Details (RETURN)").Range("J" & dpRow).Value = dp.belled_pier
            '    If Not IsNothing(dp.soil_layer_quantity) Then
            '        .Worksheets("Details (RETURN)").Range("K" & dpRow).Value = CType(dp.soil_layer_quantity, Double)
            '    Else .Worksheets("Details (RETURN)").Range("K" & dpRow).ClearContents
            '    End If
            '    If Not IsNothing(dp.concrete_compressive_strength) Then
            '        .Worksheets("Details (RETURN)").Range("L" & dpRow).Value = CType(dp.concrete_compressive_strength, Double)
            '    Else .Worksheets("Details (RETURN)").Range("L" & dpRow).ClearContents
            '    End If
            '    If Not IsNothing(dp.tie_yield_strength) Then
            '        .Worksheets("Details (RETURN)").Range("M" & dpRow).Value = CType(dp.tie_yield_strength, Double)
            '    Else .Worksheets("Details (RETURN)").Range("M" & dpRow).ClearContents
            '    End If
            '    If Not IsNothing(dp.longitudinal_rebar_yield_strength) Then
            '        .Worksheets("Details (RETURN)").Range("N" & dpRow).Value = CType(dp.longitudinal_rebar_yield_strength, Double)
            '    Else .Worksheets("Details (RETURN)").Range("N" & dpRow).ClearContents
            '    End If
            '    .Worksheets("Details (RETURN)").Range("O" & dpRow).Value = dp.rebar_effective_depths
            '    If Not IsNothing(dp.rebar_cage_2_fy_override) Then
            '        .Worksheets("Details (RETURN)").Range("P" & dpRow).Value = CType(dp.rebar_cage_2_fy_override, Double)
            '    Else .Worksheets("Details (RETURN)").Range("P" & dpRow).ClearContents
            '    End If
            '    If Not IsNothing(dp.rebar_cage_3_fy_override) Then
            '        .Worksheets("Details (RETURN)").Range("Q" & dpRow).Value = CType(dp.rebar_cage_3_fy_override, Double)
            '    Else .Worksheets("Details (RETURN)").Range("Q" & dpRow).ClearContents
            '    End If
            '    .Worksheets("Details (RETURN)").Range("R" & dpRow).Value = dp.shear_override_crit_depth
            '    If Not IsNothing(dp.shear_crit_depth_override_comp) Then
            '        .Worksheets("Details (RETURN)").Range("S" & dpRow).Value = CType(dp.shear_crit_depth_override_comp, Double)
            '    Else .Worksheets("Details (RETURN)").Range("S" & dpRow).ClearContents
            '    End If
            '    If Not IsNothing(dp.shear_crit_depth_override_uplift) Then
            '        .Worksheets("Details (RETURN)").Range("T" & dpRow).Value = CType(dp.shear_crit_depth_override_uplift, Double)
            '    Else .Worksheets("Details (RETURN)").Range("T" & dpRow).ClearContents
            '    End If
            '    .Worksheets("Details (RETURN)").Range("V" & dpRow).Value = dp.foundation_id
            '    If Not IsNothing(dp.drilled_pier_profile_qty) Then
            '        .Worksheets("Details (RETURN)").Range("W" & dpRow).Value = CType(dp.drilled_pier_profile_qty, Integer)
            '    Else .Worksheets("Details (RETURN)").Range("W" & dpRow).ClearContents
            '    End If
            '    If Not IsNothing(dp.soil_profiles) Then
            '        .Worksheets("Details (RETURN)").Range("X" & dpRow).Value = CType(dp.soil_profiles, Integer)
            '    Else .Worksheets("Details (RETURN)").Range("X" & dpRow).ClearContents
            '    End If
            '    If Not IsNothing(dp.rho_override_1) Then
            '        .Worksheets("Details (RETURN)").Range("Y" & dpRow).Value = CType(dp.rho_override_1, Double)
            '    Else .Worksheets("Details (RETURN)").Range("Y" & dpRow).ClearContents
            '    End If
            '    If Not IsNothing(dp.rho_override_2) Then
            '        .Worksheets("Details (RETURN)").Range("Z" & dpRow).Value = CType(dp.rho_override_2, Double)
            '    Else .Worksheets("Details (RETURN)").Range("Z" & dpRow).ClearContents
            '    End If
            '    If Not IsNothing(dp.rho_override_3) Then
            '        .Worksheets("Details (RETURN)").Range("AA" & dpRow).Value = CType(dp.rho_override_3, Double)
            '    Else .Worksheets("Details (RETURN)").Range("AA" & dpRow).ClearContents
            '    End If
            '    If Not IsNothing(dp.rho_override_4) Then
            '        .Worksheets("Details (RETURN)").Range("AB" & dpRow).Value = CType(dp.rho_override_4, Double)
            '    Else .Worksheets("Details (RETURN)").Range("AB" & dpRow).ClearContents
            '    End If
            '    If Not IsNothing(dp.rho_override_5) Then
            '        .Worksheets("Details (RETURN)").Range("AC" & dpRow).Value = CType(dp.rho_override_5, Double)
            '    Else .Worksheets("Details (RETURN)").Range("AC" & dpRow).ClearContents
            '    End If
            '    'If dp.bearing_type_toggle = True Then
            '    '    .Worksheets("Details (RETURN)").Range("U" & dpRow).Value = "Ult. Net Bearing Capacity (ksf)"
            '    'Else
            '    '    .Worksheets("Details (RETURN)").Range("U" & dpRow).Value = "Ult. Gross Bearing Capacity (ksf)"
            '    'End If
            '    If Not IsNothing(dp.bearing_type_toggle) Then
            '        .Worksheets("Details (RETURN)").Range("U" & dpRow).Value = dp.bearing_type_toggle
            '    End If

            '    For Each dpSec As DrilledPierSection In dp.sections
            '        If Not IsNothing(dpSec.local_drilled_pier_id) Then
            '            .Worksheets("Sections (RETURN)").Range("A" & secRow).Value = CType(dpSec.local_drilled_pier_id, Integer)
            '        Else .Worksheets("Sections (RETURN)").Range("A" & secRow).ClearContents
            '        End If
            '        If Not IsNothing(dpSec.local_section_id) Then
            '            .Worksheets("Sections (RETURN)").Range("B" & secRow).Value = CType(dpSec.local_section_id, Integer)
            '        Else .Worksheets("Sections (RETURN)").Range("B" & secRow).ClearContents
            '        End If
            '        .Worksheets("Sections (RETURN)").Range("C" & secRow).Value = dp.pier_id
            '        .Worksheets("Sections (RETURN)").Range("D" & secRow).Value = dpSec.section_id
            '        If Not IsNothing(dpSec.pier_diameter) Then
            '            .Worksheets("Sections (RETURN)").Range("E" & secRow).Value = CType(dpSec.pier_diameter, Double)
            '        Else .Worksheets("Sections (RETURN)").Range("E" & secRow).ClearContents
            '        End If
            '        If Not IsNothing(dpSec.clear_cover) Then
            '            .Worksheets("Sections (RETURN)").Range("F" & secRow).Value = CType(dpSec.clear_cover, Double)
            '        Else .Worksheets("Sections (RETURN)").Range("F" & secRow).ClearContents
            '        End If
            '        .Worksheets("Sections (RETURN)").Range("G" & secRow).Value = dpSec.clear_cover_rebar_cage_option
            '        If Not IsNothing(dpSec.tie_size) Then
            '            .Worksheets("Sections (RETURN)").Range("H" & secRow).Value = CType(dpSec.tie_size, Integer)
            '        Else .Worksheets("Sections (RETURN)").Range("H" & secRow).ClearContents
            '        End If
            '        If Not IsNothing(dpSec.tie_spacing) Then
            '            .Worksheets("Sections (RETURN)").Range("I" & secRow).Value = CType(dpSec.tie_spacing, Double)
            '        Else .Worksheets("Sections (RETURN)").Range("I" & secRow).ClearContents
            '        End If
            '        If Not IsNothing(dpSec.bottom_elevation) Then
            '            .Worksheets("Sections (RETURN)").Range("J" & secRow).Value = CType(dpSec.bottom_elevation, Double)
            '        Else .Worksheets("Sections (RETURN)").Range("J" & secRow).ClearContents
            '        End If

            '        For Each dpReb As DrilledPierRebar In dpSec.rebar
            '            If Not IsNothing(dpReb.local_drilled_pier_id) Then
            '                .Worksheets("Rebar (RETURN)").Range("A" & rebRow).Value = CType(dpReb.local_drilled_pier_id, Integer)
            '            Else .Worksheets("Rebar (RETURN)").Range("A" & rebRow).ClearContents
            '            End If
            '            If Not IsNothing(dpReb.local_section_id) Then
            '                .Worksheets("Rebar (RETURN)").Range("B" & rebRow).Value = CType(dpReb.local_section_id, Integer)
            '            Else .Worksheets("Rebar (RETURN)").Range("B" & rebRow).ClearContents
            '            End If
            '            If Not IsNothing(dpReb.local_rebar_id) Then
            '                .Worksheets("Rebar (RETURN)").Range("C" & rebRow).Value = CType(dpReb.local_rebar_id, Integer)
            '            Else .Worksheets("Rebar (RETURN)").Range("C" & rebRow).ClearContents
            '            End If
            '            .Worksheets("Rebar (RETURN)").Range("D" & rebRow).Value = dp.pier_id
            '            .Worksheets("Rebar (RETURN)").Range("E" & rebRow).Value = dpSec.section_id
            '            .Worksheets("Rebar (RETURN)").Range("F" & rebRow).Value = dpReb.rebar_id
            '            If Not IsNothing(dpReb.longitudinal_rebar_quantity) Then
            '                .Worksheets("Rebar (RETURN)").Range("G" & rebRow).Value = CType(dpReb.longitudinal_rebar_quantity, Integer)
            '            Else .Worksheets("Rebar (RETURN)").Range("G" & rebRow).ClearContents
            '            End If
            '            If Not IsNothing(dpReb.longitudinal_rebar_size) Then
            '                .Worksheets("Rebar (RETURN)").Range("H" & rebRow).Value = CType(dpReb.longitudinal_rebar_size, Integer)
            '            Else .Worksheets("Rebar (RETURN)").Range("H" & rebRow).ClearContents
            '            End If
            '            If Not IsNothing(dpReb.longitudinal_rebar_cage_diameter) Then
            '                .Worksheets("Rebar (RETURN)").Range("I" & rebRow).Value = CType(dpReb.longitudinal_rebar_cage_diameter, Double)
            '            Else .Worksheets("Rebar (RETURN)").Range("I" & rebRow).ClearContents
            '            End If

            '            rebRow += 1
            '        Next

            '        secRow += 1
            '    Next

            '    For Each dpSL As DrilledPierSoilLayer In dp.soil_layers
            '        If Not IsNothing(dpSL.local_drilled_pier_id) Then
            '            .Worksheets("Soil Layers (RETURN)").Range("A" & soilRow).Value = CType(dpSL.local_drilled_pier_id, Integer)
            '        Else .Worksheets("Soil Layers (RETURN)").Range("A" & soilRow).ClearContents
            '        End If
            '        If Not IsNothing(dpSL.local_soil_layer_id) Then
            '            .Worksheets("Soil Layers (RETURN)").Range("B" & soilRow).Value = CType(dpSL.local_soil_layer_id, Integer)
            '        Else .Worksheets("Soil Layers (RETURN)").Range("B" & soilRow).ClearContents
            '        End If
            '        .Worksheets("Soil Layers (RETURN)").Range("C" & soilRow).Value = dp.pier_id
            '        .Worksheets("Soil Layers (RETURN)").Range("D" & soilRow).Value = dpSL.soil_layer_id
            '        If Not IsNothing(dpSL.bottom_depth) Then
            '            .Worksheets("Soil Layers (RETURN)").Range("E" & soilRow).Value = CType(dpSL.bottom_depth, Double)
            '        Else .Worksheets("Soil Layers (RETURN)").Range("E" & soilRow).ClearContents
            '        End If
            '        If Not IsNothing(dpSL.effective_soil_density) Then
            '            .Worksheets("Soil Layers (RETURN)").Range("F" & soilRow).Value = CType(dpSL.effective_soil_density, Double)
            '        Else .Worksheets("Soil Layers (RETURN)").Range("F" & soilRow).ClearContents
            '        End If
            '        If Not IsNothing(dpSL.cohesion) Then
            '            .Worksheets("Soil Layers (RETURN)").Range("G" & soilRow).Value = CType(dpSL.cohesion, Double)
            '        Else .Worksheets("Soil Layers (RETURN)").Range("G" & soilRow).ClearContents
            '        End If
            '        If Not IsNothing(dpSL.friction_angle) Then
            '            .Worksheets("Soil Layers (RETURN)").Range("H" & soilRow).Value = CType(dpSL.friction_angle, Double)
            '        Else .Worksheets("Soil Layers (RETURN)").Range("H" & soilRow).ClearContents
            '        End If
            '        If Not IsNothing(dpSL.skin_friction_override_comp) Then
            '            .Worksheets("Soil Layers (RETURN)").Range("I" & soilRow).Value = CType(dpSL.skin_friction_override_comp, Double)
            '        Else .Worksheets("Soil Layers (RETURN)").Range("I" & soilRow).ClearContents
            '        End If
            '        If Not IsNothing(dpSL.skin_friction_override_uplift) Then
            '            .Worksheets("Soil Layers (RETURN)").Range("J" & soilRow).Value = CType(dpSL.skin_friction_override_uplift, Double)
            '        Else .Worksheets("Soil Layers (RETURN)").Range("J" & soilRow).ClearContents
            '        End If
            '        If Not IsNothing(dpSL.nominal_bearing_capacity) Then
            '            .Worksheets("Soil Layers (RETURN)").Range("K" & soilRow).Value = CType(dpSL.nominal_bearing_capacity, Double)
            '        Else .Worksheets("Soil Layers (RETURN)").Range("K" & soilRow).ClearContents
            '        End If
            '        If Not IsNothing(dpSL.spt_blow_count) Then
            '            .Worksheets("Soil Layers (RETURN)").Range("L" & soilRow).Value = CType(dpSL.spt_blow_count, Double)
            '        Else .Worksheets("Soil Layers (RETURN)").Range("L" & soilRow).ClearContents
            '        End If

            '        soilRow += 1
            '    Next

            '    If ds.Tables("Belled Details SQL").Rows.Count > 0 Then
            '        If dp.belled_pier = True Then
            '            If Not IsNothing(dp.belled_details.local_drilled_pier_id) Then
            '                .Worksheets("Belled (RETURN)").Range("A" & dpRow).Value = CType(dp.belled_details.local_drilled_pier_id, Integer)
            '            Else .Worksheets("Belled (RETURN)").Range("A" & dpRow).ClearContents
            '            End If
            '            .Worksheets("Belled (RETURN)").Range("B" & dpRow).Value = dp.pier_id
            '            .Worksheets("Belled (RETURN)").Range("C" & dpRow).Value = dp.belled_details.belled_pier_id
            '            .Worksheets("Belled (RETURN)").Range("D" & dpRow).Value = dp.belled_details.belled_pier_option
            '            If Not IsNothing(dp.belled_details.bottom_diameter_of_bell) Then
            '                .Worksheets("Belled (RETURN)").Range("E" & dpRow).Value = CType(dp.belled_details.bottom_diameter_of_bell, Double)
            '            Else .Worksheets("Belled (RETURN)").Range("E" & dpRow).ClearContents
            '            End If
            '            If Not IsNothing(dp.belled_details.bell_input_type) Then
            '                .Worksheets("Belled (RETURN)").Range("F" & dpRow).Value = CType(dp.belled_details.bell_input_type, Double)
            '            Else .Worksheets("Belled (RETURN)").Range("F" & dpRow).ClearContents
            '            End If
            '            If Not IsNothing(dp.belled_details.bell_angle) Then
            '                .Worksheets("Belled (RETURN)").Range("G" & dpRow).Value = CType(dp.belled_details.bell_angle, Double)
            '            Else .Worksheets("Belled (RETURN)").Range("G" & dpRow).ClearContents
            '            End If
            '            If Not IsNothing(dp.belled_details.bell_height) Then
            '                .Worksheets("Belled (RETURN)").Range("H" & dpRow).Value = CType(dp.belled_details.bell_height, Double)
            '            Else .Worksheets("Belled (RETURN)").Range("H" & dpRow).ClearContents
            '            End If
            '            If Not IsNothing(dp.belled_details.bell_toe_height) Then
            '                .Worksheets("Belled (RETURN)").Range("I" & dpRow).Value = CType(dp.belled_details.bell_toe_height, Double)
            '            Else .Worksheets("Belled (RETURN)").Range("I" & dpRow).ClearContents
            '            End If
            '            If Not IsNothing(dp.belled_details.neglect_top_soil_layer) Then
            '                .Worksheets("Belled (RETURN)").Range("J" & dpRow).Value = CType(dp.belled_details.neglect_top_soil_layer, Double)
            '            Else .Worksheets("Belled (RETURN)").Range("J" & dpRow).ClearContents
            '            End If
            '            If Not IsNothing(dp.belled_details.swelling_expansive_soil) Then
            '                .Worksheets("Belled (RETURN)").Range("K" & dpRow).Value = CType(dp.belled_details.swelling_expansive_soil, Double)
            '            Else .Worksheets("Belled (RETURN)").Range("K" & dpRow).ClearContents
            '            End If
            '            If Not IsNothing(dp.belled_details.depth_of_expansive_soil) Then
            '                .Worksheets("Belled (RETURN)").Range("L" & dpRow).Value = CType(dp.belled_details.depth_of_expansive_soil, Double)
            '            Else .Worksheets("Belled (RETURN)").Range("L" & dpRow).ClearContents
            '            End If
            '            If Not IsNothing(dp.belled_details.expansive_soil_force) Then
            '                .Worksheets("Belled (RETURN)").Range("M" & dpRow).Value = CType(dp.belled_details.expansive_soil_force, Double)
            '            Else .Worksheets("Belled (RETURN)").Range("M" & dpRow).ClearContents
            '            End If
            '        End If
            '    End If

            '    If ds.Tables("Embedded Details SQL").Rows.Count > 0 Then
            '        If dp.embedded_pole Then
            '            If Not IsNothing(dp.embed_details.local_drilled_pier_id) Then
            '                .Worksheets("Embedded (RETURN)").Range("A" & dpRow).Value = CType(dp.embed_details.local_drilled_pier_id, Integer)
            '            Else .Worksheets("Embedded (RETURN)").Range("A" & dpRow).ClearContents
            '            End If
            '            .Worksheets("Embedded (RETURN)").Range("B" & dpRow).Value = dp.pier_id
            '            .Worksheets("Embedded (RETURN)").Range("C" & dpRow).Value = dp.embed_details.embedded_id
            '            .Worksheets("Embedded (RETURN)").Range("D" & dpRow).Value = dp.embed_details.embedded_pole_option
            '            .Worksheets("Embedded (RETURN)").Range("E" & dpRow).Value = dp.embed_details.encased_in_concrete
            '            If Not IsNothing(dp.embed_details.pole_side_quantity) Then
            '                .Worksheets("Embedded (RETURN)").Range("F" & dpRow).Value = CType(dp.embed_details.pole_side_quantity, Integer)
            '            Else .Worksheets("Embedded (RETURN)").Range("F" & dpRow).ClearContents
            '            End If
            '            If Not IsNothing(dp.embed_details.pole_yield_strength) Then
            '                .Worksheets("Embedded (RETURN)").Range("G" & dpRow).Value = CType(dp.embed_details.pole_yield_strength, Double)
            '            Else .Worksheets("Embedded (RETURN)").Range("G" & dpRow).ClearContents
            '            End If
            '            If Not IsNothing(dp.embed_details.pole_thickness) Then
            '                .Worksheets("Embedded (RETURN)").Range("H" & dpRow).Value = CType(dp.embed_details.pole_thickness, Double)
            '            Else .Worksheets("Embedded (RETURN)").Range("H" & dpRow).ClearContents
            '            End If
            '            If Not IsNothing(dp.embed_details.embedded_pole_input_type) Then
            '                .Worksheets("Embedded (RETURN)").Range("I" & dpRow).Value = CType(dp.embed_details.embedded_pole_input_type, Double)
            '            Else .Worksheets("Embedded (RETURN)").Range("I" & dpRow).ClearContents
            '            End If
            '            If Not IsNothing(dp.embed_details.pole_diameter_toc) Then
            '                .Worksheets("Embedded (RETURN)").Range("J" & dpRow).Value = CType(dp.embed_details.pole_diameter_toc, Double)
            '            Else .Worksheets("Embedded (RETURN)").Range("J" & dpRow).ClearContents
            '            End If
            '            If Not IsNothing(dp.embed_details.pole_top_diameter) Then
            '                .Worksheets("Embedded (RETURN)").Range("K" & dpRow).Value = CType(dp.embed_details.pole_top_diameter, Double)
            '            Else .Worksheets("Embedded (RETURN)").Range("K" & dpRow).ClearContents
            '            End If
            '            If Not IsNothing(dp.embed_details.pole_bottom_diameter) Then
            '                .Worksheets("Embedded (RETURN)").Range("L" & dpRow).Value = CType(dp.embed_details.pole_bottom_diameter, Double)
            '            Else .Worksheets("Embedded (RETURN)").Range("L" & dpRow).ClearContents
            '            End If
            '            If Not IsNothing(dp.embed_details.pole_section_length) Then
            '                .Worksheets("Embedded (RETURN)").Range("M" & dpRow).Value = CType(dp.embed_details.pole_section_length, Double)
            '            Else .Worksheets("Embedded (RETURN)").Range("M" & dpRow).ClearContents
            '            End If
            '            If Not IsNothing(dp.embed_details.pole_taper_factor) Then
            '                .Worksheets("Embedded (RETURN)").Range("N" & dpRow).Value = CType(dp.embed_details.pole_taper_factor, Double)
            '            Else .Worksheets("Embedded (RETURN)").Range("N" & dpRow).ClearContents
            '            End If
            '            If Not IsNothing(dp.embed_details.pole_bend_radius_override) Then
            '                .Worksheets("Embedded (RETURN)").Range("O" & dpRow).Value = CType(dp.embed_details.pole_bend_radius_override, Double)
            '            Else .Worksheets("Embedded (RETURN)").Range("O" & dpRow).ClearContents
            '            End If
            '        End If
            '    End If

            '    For Each dpp As DrilledPierProfile In dp.drilled_pier_profiles
            '        If Not IsNothing(dpp.local_drilled_pier_id) Then
            '            .Worksheets("Profiles (RETURN)").Range("A" & profileRow).Value = CType(dpp.local_drilled_pier_id, Integer)
            '        Else .Worksheets("Profiles (RETURN)").Range("A" & profileRow).ClearContents
            '        End If
            '        If Not IsNothing(dpp.reaction_position) Then
            '            .Worksheets("Profiles (RETURN)").Range("B" & profileRow).Value = CType(dpp.reaction_position, Integer)
            '        Else .Worksheets("Profiles (RETURN)").Range("B" & profileRow).ClearContents
            '        End If
            '        If Not IsNothing(dpp.drilled_pier_id) Then
            '            .Worksheets("Profiles (RETURN)").Range("C" & profileRow).Value = CType(dpp.drilled_pier_id, Integer)
            '        Else .Worksheets("Profiles (RETURN)").Range("C" & profileRow).ClearContents
            '        End If
            '        .Worksheets("Profiles (RETURN)").Range("D" & profileRow).Value = CType(dpp.profile_id, Integer)
            '        If Not IsNothing(dpp.reaction_location) Then
            '            .Worksheets("Profiles (RETURN)").Range("E" & profileRow).Value = CType(dpp.reaction_location, String)
            '        Else .Worksheets("Profiles (RETURN)").Range("E" & profileRow).ClearContents
            '        End If
            '        If Not IsNothing(dpp.drilled_pier_profile) Then
            '            .Worksheets("Profiles (RETURN)").Range("F" & profileRow).Value = CType(dpp.drilled_pier_profile, String)
            '        Else .Worksheets("Profiles (RETURN)").Range("F" & profileRow).ClearContents
            '        End If
            '        If Not IsNothing(dpp.soil_profile) Then
            '            .Worksheets("Profiles (RETURN)").Range("G" & profileRow).Value = CType(dpp.soil_profile, String)
            '        Else .Worksheets("Profiles (RETURN)").Range("G" & profileRow).ClearContents
            '        End If

            '        profileRow += 1

            '    Next

            '    dpRow += 1

            'Next

            'set booleans for workbook open event
            '.Worksheets("SUMMARY").Range("EDSReturn").Value = True
            '.Worksheets("SUMMARY").Range("EDSReactions").Value = True

            ''add code here rather than workbook open. Populates internal tool database
            'Dim ws1 As Worksheet = .Worksheets("Database")
            'Dim ws2 As Worksheet = .Worksheets("Database (RETURN)")
            'Dim ws3 As Worksheet = .Worksheets("SUMMARY")
            'Dim ws4 As Worksheet = .Worksheets("Profiles (RETURN)")

            'Dim ProfileStartRow1, ProfileStartRow2, ProfileStartCol1, ProfileStartCol2, CopyProfiles As Integer

            'ProfileStartRow1 = 53 'Database start row
            'ProfileStartRow2 = 4 'Import from EDS data start row
            'ProfileStartCol1 = 7 'Database start column
            'ProfileStartCol2 = 6 'Import from EDS data start column
            'CopyProfiles = 50 'number of profiles in database

            ''General through Pier Section 5 (Part 1 - Depth 1 input is formula, handled below)
            'ws1.Range("G" & ProfileStartRow1 + 7 & ":" & "BD" & ProfileStartRow1 + 11).Value = ws2.Range("F" & ProfileStartRow2 + 7 & ":" & "BC" & ProfileStartRow2 + 11).Value

            ''General through Pier Section 5 (Part 2 - Depth 1 input is formula, handled below)
            'ws1.Range("G" & ProfileStartRow1 + 13 & ":" & "BD" & ProfileStartRow1 + 94).Value = ws2.Range("F" & ProfileStartRow2 + 13 & ":" & "BC" & ProfileStartRow2 + 94).Value

            ''Checks, Embedded Pole, and Belled Pier (Part 1)
            'ws1.Range("G" & ProfileStartRow1 + 97 & ":" & "BD" & ProfileStartRow1 + 116).Value = ws2.Range("F" & ProfileStartRow2 + 97 & ":" & "BC" & ProfileStartRow2 + 116).Value

            ''Belled Pier (Part 2)
            'ws1.Range("G" & ProfileStartRow1 + 120 & ":" & "BD" & ProfileStartRow1 + 120).Value = ws2.Range("F" & ProfileStartRow2 + 120 & ":" & "BC" & ProfileStartRow2 + 120).Value

            ''Belled Pier (Part 3)
            'ws1.Range("G" & ProfileStartRow1 + 122 & ":" & "BD" & ProfileStartRow1 + 125).Value = ws2.Range("F" & ProfileStartRow2 + 122 & ":" & "BC" & ProfileStartRow2 + 125).Value

            ''Soil
            'ws1.Range("G" & ProfileStartRow1 + 127 & ":" & "BD" & ProfileStartRow1 + 373).Value = ws2.Range("F" & ProfileStartRow2 + 127 & ":" & "BC" & ProfileStartRow2 + 373).Value

            ''Additional Rebar Info
            'ws1.Range("G" & ProfileStartRow1 + 4389 & ":" & "BD" & ProfileStartRow1 + 4397).Value = ws2.Range("F" & ProfileStartRow2 + 378 & ":" & "BC" & ProfileStartRow2 + 386).Value

            'Dim i, j, MaxPierID, CountID As Integer

            ''maximum pier ID returned from EDS
            'MaxPierID = Math.Max(.Worksheets("Details (RETURN)").Range("A1:A1000"))
            ''Count of Returned EDS
            'CountID = Math.Count(.Worksheets("Details (RETURN)").Range("A1:A1000"))

            'MaxPierID = (From DrilledPier In DrilledPiers Order By DrilledPier.local_drilled_pier_id Descending).First.local_drilled_pier_id
            'MsgBox(MaxPierID)

            'CountID = 0
            'For Each DrilledPier In DrilledPiers
            '    CountID = CountID + 1
            'Next

            'MsgBox(CountID)



            'Dim list As List(Of Integer) = New List(Of Integer)({19, 23, 29})

            '' Find value greater than 20.
            'Dim val As Integer = list.FindLast(Function(value As Integer)
            '                                       Return value > 20
            '                                   End Function)
            ''Console.WriteLine(val)
            'MsgBox(val)


            'Dim val As String

            '' Find last 5-letter string.
            'val = List.FindLast(Function(value As String)
            '                        Return value.Length = 5
            '                    End Function)
            'Console.WriteLine("FINDLAST: {0}", val)



            ''8-18-2021 TESTING START~~~~~~~~~~~~~~~~
            'For Each dp As DrilledPier In DrilledPiers

            '    Dim currentIdxVal As Integer
            '    Dim lastIdxVal As Integer

            '    currentIdxVal = DrilledPiers.IndexOf(dp)
            '    lastIdxVal = DrilledPiers.LastIndexOf(dp, CType(dp.local_drilled_pier_id, Integer)) 'error here not recognizing the last matching index
            '    'this line is is only looking at indeces up to the second value (local drilled pier)


            '    'run for only the last index of local pier ids in list of drilled piers
            '    If currentIdxVal = lastIdxVal Then
            '        Console.WriteLine(currentIdxVal)
            '    End If

            'Next
            ''8-18-2021 TESTING END~~~~~~~~~~~~~~~~




            Dim colCounter As Integer = 6
            Dim myCol As String
            Dim rowStart As Integer = 53

            For Each dp As DrilledPier In DrilledPiers

                colCounter = 6 + dp.local_drilled_pier_id
                myCol = GetExcelColumnName(colCounter)

                'DRILLED PIER DETAILS
                If Not IsNothing(dp.pier_id) Then
                    .Worksheets("Database").Range(myCol & rowStart - 1).Value = CType(dp.pier_id, Integer)
                Else .Worksheets("Database").Range(myCol & rowStart - 1).ClearContents
                End If
                If Not IsNothing(dp.concrete_compressive_strength) Then
                    .Worksheets("Database").Range(myCol & rowStart + 7).Value = CType(dp.concrete_compressive_strength, Double)
                Else .Worksheets("Database").Range(myCol & rowStart + 7).ClearContents
                End If
                If Not IsNothing(dp.longitudinal_rebar_yield_strength) Then
                    .Worksheets("Database").Range(myCol & rowStart + 8).Value = CType(dp.longitudinal_rebar_yield_strength, Double)
                Else .Worksheets("Database").Range(myCol & rowStart + 8).ClearContents
                End If
                If Not IsNothing(dp.tie_yield_strength) Then
                    .Worksheets("Database").Range(myCol & rowStart + 9).Value = CType(dp.tie_yield_strength, Double)
                Else .Worksheets("Database").Range(myCol & rowStart + 9).ClearContents
                End If
                If Not IsNothing(dp.foundation_depth) Then
                    .Worksheets("Database").Range(myCol & rowStart + 10).Value = CType(dp.foundation_depth, Double)
                Else .Worksheets("Database").Range(myCol & rowStart + 10).ClearContents
                End If
                If Not IsNothing(dp.extension_above_grade) Then
                    .Worksheets("Database").Range(myCol & rowStart + 11).Value = CType(dp.extension_above_grade, Double)
                Else .Worksheets("Database").Range(myCol & rowStart + 11).ClearContents
                End If
                If CType(dp.groundwater_depth, String) = "N/A" Then
                    .Worksheets("Database").Range(myCol & rowStart + 17).Value = CType(dp.groundwater_depth, String)
                ElseIf Not IsNothing(dp.groundwater_depth) Then
                    .Worksheets("Database").Range(myCol & rowStart + 17).Value = CType(dp.groundwater_depth, Double)
                Else .Worksheets("Database").Range(myCol & rowStart + 17).ClearContents
                End If
                If Not IsNothing(dp.soil_layer_quantity) Then
                    .Worksheets("Database").Range(myCol & rowStart + 18).Value = CType(dp.soil_layer_quantity, Double)
                Else .Worksheets("Database").Range(myCol & rowStart + 18).ClearContents
                End If
                If Not IsNothing(dp.bearing_type_toggle) Then
                    .Worksheets("Database").Range(myCol & rowStart + 19).Value = CType(dp.bearing_type_toggle, String)
                Else .Worksheets("Database").Range(myCol & rowStart + 19).ClearContents
                End If
                .Worksheets("Database").Range(myCol & rowStart + 97).Value = CType(dp.check_shear_along_depth, Boolean)
                .Worksheets("Database").Range(myCol & rowStart + 98).Value = CType(dp.utilize_shear_friction_methodology, Boolean)
                .Worksheets("Database").Range(myCol & rowStart + 100).Value = CType(dp.embedded_pole, Boolean)
                .Worksheets("Database").Range(myCol & rowStart + 112).Value = CType(dp.belled_pier, Boolean)
                .Worksheets("Database").Range(myCol & rowStart + 4389).Value = CType(dp.assume_min_steel, String)
                .Worksheets("Database").Range(myCol & rowStart + 4390).Value = CType(dp.rebar_effective_depths, Boolean)
                If Not IsNothing(dp.rebar_cage_2_fy_override) Then
                    .Worksheets("Database").Range(myCol & rowStart + 4391).Value = CType(dp.rebar_cage_2_fy_override, Double)
                Else .Worksheets("Database").Range(myCol & rowStart + 4391).ClearContents
                End If
                If Not IsNothing(dp.rebar_cage_3_fy_override) Then
                    .Worksheets("Database").Range(myCol & rowStart + 4392).Value = CType(dp.rebar_cage_3_fy_override, Double)
                Else .Worksheets("Database").Range(myCol & rowStart + 4392).ClearContents
                End If
                If Not IsNothing(dp.rho_override_1) Then
                    .Worksheets("Database").Range(myCol & rowStart + 4393).Value = CType(dp.rho_override_1, Double)
                Else .Worksheets("Database").Range(myCol & rowStart + 4393).ClearContents
                End If
                If Not IsNothing(dp.rho_override_2) Then
                    .Worksheets("Database").Range(myCol & rowStart + 4394).Value = CType(dp.rho_override_2, Double)
                Else .Worksheets("Database").Range(myCol & rowStart + 4394).ClearContents
                End If
                If Not IsNothing(dp.rho_override_3) Then
                    .Worksheets("Database").Range(myCol & rowStart + 4395).Value = CType(dp.rho_override_3, Double)
                Else .Worksheets("Database").Range(myCol & rowStart + 4395).ClearContents
                End If
                If Not IsNothing(dp.rho_override_4) Then
                    .Worksheets("Database").Range(myCol & rowStart + 4396).Value = CType(dp.rho_override_4, Double)
                Else .Worksheets("Database").Range(myCol & rowStart + 4396).ClearContents
                End If
                If Not IsNothing(dp.rho_override_5) Then
                    .Worksheets("Database").Range(myCol & rowStart + 4397).Value = CType(dp.rho_override_5, Double)
                Else .Worksheets("Database").Range(myCol & rowStart + 4397).ClearContents
                End If
                .Worksheets("Database").Range(myCol & rowStart + 99).Value = CType(dp.shear_override_crit_depth, Boolean)
                If Not IsNothing(dp.shear_crit_depth_override_comp) Then
                    .Worksheets("Database").Range(myCol & rowStart + 374).Value = CType(dp.shear_crit_depth_override_comp, Double)
                Else .Worksheets("Database").Range(myCol & rowStart + 374).Formula = .Worksheets("Database").Range(GetExcelColumnName(colCounter + 51) & rowStart + 374).Formula
                End If
                If Not IsNothing(dp.shear_crit_depth_override_uplift) Then
                    .Worksheets("Database").Range(myCol & rowStart + 376).Value = CType(dp.shear_crit_depth_override_uplift, Double)
                Else .Worksheets("Database").Range(myCol & rowStart + 376).Formula = .Worksheets("Database").Range(GetExcelColumnName(colCounter + 51) & rowStart + 376).Formula
                End If

                Dim depth As Integer = 0
                Dim secBump As Integer = 0
                Dim secStart As Integer = 20
                Dim secCount As Integer = 1

                'DRILLED PIER SECTION
                For Each dpSec As DrilledPierSection In dp.sections

                    If Not IsNothing(dpSec.pier_diameter) Then
                        .Worksheets("Database").Range(myCol & rowStart + secStart + 0).Value = CType(dpSec.pier_diameter, Double)
                    Else .Worksheets("Database").Range(myCol & rowStart + secStart + 0).ClearContents
                    End If
                    If Not IsNothing(dpSec.clear_cover) Then
                        .Worksheets("Database").Range(myCol & rowStart + secStart + 3).Value = CType(dpSec.clear_cover, Double)
                    Else .Worksheets("Database").Range(myCol & rowStart + secStart + 3).ClearContents
                    End If
                    If Not IsNothing(dpSec.tie_size) Then
                        .Worksheets("Database").Range(myCol & rowStart + secStart + 4).Value = CType(dpSec.tie_size, Integer)
                    Else .Worksheets("Database").Range(myCol & rowStart + secStart + 4).ClearContents
                    End If
                    If Not IsNothing(dpSec.tie_spacing) Then
                        .Worksheets("Database").Range(myCol & rowStart + secStart + 5).Value = CType(dpSec.tie_spacing, Double)
                    Else .Worksheets("Database").Range(myCol & rowStart + secStart + 5).ClearContents
                    End If
                    If Not IsNothing(dpSec.clear_cover_rebar_cage_option) Then
                        .Worksheets("Database").Range(myCol & rowStart + secStart + 14).Value = CType(dpSec.clear_cover_rebar_cage_option, String)
                    Else .Worksheets("Database").Range(myCol & rowStart + secStart + 14).ClearContents
                    End If

                    If secCount > 1 Then depth += 1
                    If Not IsNothing(dpSec.bottom_elevation) Then
                        .Worksheets("Database").Range(myCol & rowStart + 12 + depth).Value = CType(dpSec.bottom_elevation, Double)
                    Else .Worksheets("Database").Range(myCol & rowStart + 12 + depth).ClearContents
                    End If

                    'DRILLED PIER REBAR
                    Dim rebCount As Integer = 1

                    For Each dpReb As DrilledPierRebar In dpSec.rebar

                        If rebCount = 1 Then
                            If Not IsNothing(dpReb.longitudinal_rebar_quantity) Then
                                .Worksheets("Database").Range(myCol & rowStart + secStart + 1).Value = CType(dpReb.longitudinal_rebar_quantity, Double)
                            Else .Worksheets("Database").Range(myCol & rowStart + secStart + 1).ClearContents
                            End If
                            If Not IsNothing(dpReb.longitudinal_rebar_size) Then
                                .Worksheets("Database").Range(myCol & rowStart + secStart + 2).Value = CType(dpReb.longitudinal_rebar_size, Double)
                            Else .Worksheets("Database").Range(myCol & rowStart + secStart + 2).ClearContents
                            End If
                        ElseIf rebCount = 2 Then
                            If Not IsNothing(dpReb.longitudinal_rebar_quantity) Then
                                .Worksheets("Database").Range(myCol & rowStart + secStart + 6).Value = CType(dpReb.longitudinal_rebar_quantity, Double)
                            Else .Worksheets("Database").Range(myCol & rowStart + secStart + 6).ClearContents
                            End If
                            If Not IsNothing(dpReb.longitudinal_rebar_size) Then
                                .Worksheets("Database").Range(myCol & rowStart + secStart + 7).Value = CType(dpReb.longitudinal_rebar_size, Double)
                            Else .Worksheets("Database").Range(myCol & rowStart + secStart + 7).ClearContents
                            End If
                            If Not IsNothing(dpReb.longitudinal_rebar_cage_diameter) Then
                                .Worksheets("Database").Range(myCol & rowStart + secStart + 8).Value = CType(dpReb.longitudinal_rebar_cage_diameter, Double)
                            Else .Worksheets("Database").Range(myCol & rowStart + secStart + 8).ClearContents
                            End If
                        Else
                            If Not IsNothing(dpReb.longitudinal_rebar_quantity) Then
                                .Worksheets("Database").Range(myCol & rowStart + secStart + 10).Value = CType(dpReb.longitudinal_rebar_quantity, Double)
                            Else .Worksheets("Database").Range(myCol & rowStart + secStart + 10).ClearContents
                            End If
                            If Not IsNothing(dpReb.longitudinal_rebar_size) Then
                                .Worksheets("Database").Range(myCol & rowStart + secStart + 11).Value = CType(dpReb.longitudinal_rebar_size, Double)
                            Else .Worksheets("Database").Range(myCol & rowStart + secStart + 11).ClearContents
                            End If
                            If Not IsNothing(dpReb.longitudinal_rebar_cage_diameter) Then
                                .Worksheets("Database").Range(myCol & rowStart + secStart + 12).Value = CType(dpReb.longitudinal_rebar_cage_diameter, Double)
                            Else .Worksheets("Database").Range(myCol & rowStart + secStart + 12).ClearContents
                            End If
                        End If

                        rebCount += 1
                    Next

                    secCount += 1
                    secBump += 15
                    secStart += secBump

                Next

                'BELLED PIER
                If dp.belled_pier = True Then

                    .Worksheets("Database").Range(myCol & 112).Value = CType(dp.belled_pier, Boolean)
                    If Not IsNothing(dp.belled_details.bottom_diameter_of_bell) Then
                        .Worksheets("Database").Range(myCol & 113).Value = CType(dp.belled_details.bottom_diameter_of_bell, Double)
                    Else .Worksheets("Database").Range(myCol & 113).ClearContents
                    End If
                    If Not IsNothing(dp.belled_details.bell_angle) Then
                        .Worksheets("Database").Range(myCol & 114).Value = CType(dp.belled_details.bell_angle, Double)
                    Else .Worksheets("Database").Range(myCol & 114).ClearContents
                    End If
                    .Worksheets("Database").Range(myCol & 115).Value = CType(dp.belled_details.bell_input_type, String)
                    If Not IsNothing(dp.belled_details.bell_height) Then
                        .Worksheets("Database").Range(myCol & 116).Value = CType(dp.belled_details.bell_height, Double)
                    Else .Worksheets("Database").Range(myCol & 116).ClearContents
                    End If
                    If Not IsNothing(dp.belled_details.bell_toe_height) Then
                        .Worksheets("Database").Range(myCol & 120).Value = CType(dp.belled_details.bell_toe_height, Double)
                    Else .Worksheets("Database").Range(myCol & 120).ClearContents
                    End If
                    .Worksheets("Database").Range(myCol & 122).Value = CType(dp.belled_details.neglect_top_soil_layer, Boolean)
                    .Worksheets("Database").Range(myCol & 123).Value = CType(dp.belled_details.swelling_expansive_soil, Boolean)
                    If Not IsNothing(dp.belled_details.depth_of_expansive_soil) Then
                        .Worksheets("Database").Range(myCol & 124).Value = CType(dp.belled_details.depth_of_expansive_soil, Double)
                    Else .Worksheets("Database").Range(myCol & 124).ClearContents
                    End If
                    If Not IsNothing(dp.belled_details.expansive_soil_force) Then
                        .Worksheets("Database").Range(myCol & 125).Value = CType(dp.belled_details.expansive_soil_force, Double)
                    Else .Worksheets("Database").Range(myCol & 125).ClearContents
                    End If

                End If

                'EMBEDDED PIER
                If dp.embedded_pole = True Then

                    .Worksheets("Database").Range(myCol & 100).Value = CType(dp.embedded_pole, Boolean)
                    .Worksheets("Database").Range(myCol & 101).Value = CType(dp.embed_details.encased_in_concrete, Boolean)
                    If Not IsNothing(dp.embed_details.pole_side_quantity) Then
                        .Worksheets("Database").Range(myCol & 102).Value = CType(dp.embed_details.pole_side_quantity, Integer)
                    Else .Worksheets("Database").Range(myCol & 102).ClearContents
                    End If
                    If Not IsNothing(dp.embed_details.pole_yield_strength) Then
                        .Worksheets("Database").Range(myCol & 103).Value = CType(dp.embed_details.pole_yield_strength, Double)
                    Else .Worksheets("Database").Range(myCol & 103).ClearContents
                    End If
                    If Not IsNothing(dp.embed_details.pole_thickness) Then
                        .Worksheets("Database").Range(myCol & 104).Value = CType(dp.embed_details.pole_thickness, Double)
                    Else .Worksheets("Database").Range(myCol & 104).ClearContents
                    End If
                    .Worksheets("Database").Range(myCol & 105).Value = CType(dp.embed_details.embedded_pole_input_type, String)
                    If Not IsNothing(dp.embed_details.pole_diameter_toc) Then
                        .Worksheets("Database").Range(myCol & 106).Value = CType(dp.embed_details.pole_diameter_toc, Double)
                    Else .Worksheets("Database").Range(myCol & 106).ClearContents
                    End If
                    If Not IsNothing(dp.embed_details.pole_top_diameter) Then
                        .Worksheets("Database").Range(myCol & 107).Value = CType(dp.embed_details.pole_top_diameter, Double)
                    Else .Worksheets("Database").Range(myCol & 107).ClearContents
                    End If
                    If Not IsNothing(dp.embed_details.pole_bottom_diameter) Then
                        .Worksheets("Database").Range(myCol & 108).Value = CType(dp.embed_details.pole_bottom_diameter, Double)
                    Else .Worksheets("Database").Range(myCol & 108).ClearContents
                    End If
                    If Not IsNothing(dp.embed_details.pole_section_length) Then
                        .Worksheets("Database").Range(myCol & 109).Value = CType(dp.embed_details.pole_section_length, Double)
                    Else .Worksheets("Database").Range(myCol & 109).ClearContents
                    End If
                    If Not IsNothing(dp.embed_details.pole_taper_factor) Then
                        .Worksheets("Database").Range(myCol & 110).Value = CType(dp.embed_details.pole_taper_factor, Double)
                    Else .Worksheets("Database").Range(myCol & 110).ClearContents
                    End If
                    If Not IsNothing(dp.embed_details.pole_bend_radius_override) Then
                        .Worksheets("Database").Range(myCol & 111).Value = CType(dp.embed_details.pole_bend_radius_override, Double)
                    Else .Worksheets("Database").Range(myCol & 111).ClearContents
                    End If

                End If

                'DRILLED PIER PROFILES
                Dim summaryRowStart As Integer = 10

                For Each dpp As DrilledPierProfile In dp.drilled_pier_profiles
                    'Profile Return
                    If Not IsNothing(dpp.local_drilled_pier_id) Then
                        .Worksheets("Profiles (RETURN)").Range("A" & profileRow).Value = CType(dpp.local_drilled_pier_id, Integer)
                    Else .Worksheets("Profiles (RETURN)").Range("A" & profileRow).ClearContents
                    End If
                    If Not IsNothing(dpp.reaction_position) Then
                        .Worksheets("Profiles (RETURN)").Range("B" & profileRow).Value = CType(dpp.reaction_position, Integer)
                    Else .Worksheets("Profiles (RETURN)").Range("B" & profileRow).ClearContents
                    End If
                    If Not IsNothing(dpp.drilled_pier_id) Then
                        .Worksheets("Profiles (RETURN)").Range("C" & profileRow).Value = CType(dpp.drilled_pier_id, Integer)
                    Else .Worksheets("Profiles (RETURN)").Range("C" & profileRow).ClearContents
                    End If
                    .Worksheets("Profiles (RETURN)").Range("D" & profileRow).Value = CType(dpp.profile_id, Integer)
                    If Not IsNothing(dpp.reaction_location) Then
                        .Worksheets("Profiles (RETURN)").Range("E" & profileRow).Value = CType(dpp.reaction_location, String)
                    Else .Worksheets("Profiles (RETURN)").Range("E" & profileRow).ClearContents
                    End If
                    If Not IsNothing(dpp.drilled_pier_profile) Then
                        .Worksheets("Profiles (RETURN)").Range("F" & profileRow).Value = CType(dpp.drilled_pier_profile, String)
                    Else .Worksheets("Profiles (RETURN)").Range("F" & profileRow).ClearContents
                    End If
                    If Not IsNothing(dpp.soil_profile) Then
                        .Worksheets("Profiles (RETURN)").Range("G" & profileRow).Value = CType(dpp.soil_profile, String)
                    Else .Worksheets("Profiles (RETURN)").Range("G" & profileRow).ClearContents
                    End If

                    'SUMMARY
                    If Not IsNothing(dpp.reaction_position) Then
                        .Worksheets("SUMMARY").Range("D" & summaryRowStart + CType(dpp.reaction_position, Integer)).Value = CType(dpp.drilled_pier_profile, String)
                        If dpp.drilled_pier_profile = dpp.reaction_position Then
                            .Worksheets("SUMMARY").Range("G" & summaryRowStart + CType(dpp.reaction_position, Integer)).Value = False
                        Else
                            .Worksheets("SUMMARY").Range("G" & summaryRowStart + CType(dpp.reaction_position, Integer)).Value = True
                        End If
                    End If
                    If Not IsNothing(dpp.reaction_position) Then
                        .Worksheets("SUMMARY").Range("E" & summaryRowStart + CType(dpp.reaction_position, Integer)).Value = CType(dpp.soil_profile, String)
                        If dpp.soil_profile = dpp.reaction_position Then
                            .Worksheets("SUMMARY").Range("H" & summaryRowStart + CType(dpp.reaction_position, Integer)).Value = False
                        Else
                            .Worksheets("SUMMARY").Range("H" & summaryRowStart + CType(dpp.reaction_position, Integer)).Value = True
                        End If
                    End If
                    .Worksheets("SUMMARY").Range("I" & summaryRowStart + CType(dpp.reaction_position, Integer)).Value = False

                    profileRow += 1

                Next

                .Worksheets("SUMMARY").Range("EDSReactions").Value = True

                'DRILLED PIER SOIL LAYER
                Dim soilCount As Integer
                Dim soilStart As Integer = 127
                Dim soilColCounter As Integer
                Dim mySoilCol As String

                For Each dpp As DrilledPierProfile In dp.drilled_pier_profiles

                    soilCount = 1

                    soilColCounter = 6 + dpp.soil_profile
                    mySoilCol = GetExcelColumnName(soilColCounter)

                    For Each dpSL As DrilledPierSoilLayer In dp.soil_layers

                        If Not IsNothing(dpSL.bottom_depth) Then
                            .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 0 + (soilCount - 1)).Value = CType(dpSL.bottom_depth, Double)
                        Else .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 0 + (soilCount - 1)).ClearContents
                        End If
                        If Not IsNothing(dpSL.effective_soil_density) Then
                            .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 1 + (soilCount - 1)).Value = CType(dpSL.effective_soil_density, Double)
                        Else .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 1 + (soilCount - 1)).ClearContents
                        End If
                        If Not IsNothing(dpSL.cohesion) Then
                            .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 2 + (soilCount - 1)).Value = CType(dpSL.cohesion, Double)
                        Else .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 2 + (soilCount - 1)).ClearContents
                        End If
                        If Not IsNothing(dpSL.friction_angle) Then
                            .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 3 + (soilCount - 1)).Value = CType(dpSL.friction_angle, Double)
                        Else .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 3 + (soilCount - 1)).ClearContents
                        End If
                        If Not IsNothing(dpSL.skin_friction_override_comp) Then
                            .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 4 + (soilCount - 1)).Value = CType(dpSL.skin_friction_override_comp, Double)
                        Else .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 4 + (soilCount - 1)).ClearContents
                        End If
                        If Not IsNothing(dpSL.skin_friction_override_uplift) Then
                            .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 5 + (soilCount - 1)).Value = CType(dpSL.skin_friction_override_uplift, Double)
                        Else .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 5 + (soilCount - 1)).ClearContents
                        End If
                        If Not IsNothing(dpSL.nominal_bearing_capacity) Then
                            .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 6 + (soilCount - 1)).Value = CType(dpSL.nominal_bearing_capacity, Double)
                        Else .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 6 + (soilCount - 1)).ClearContents
                        End If
                        If Not IsNothing(dpSL.spt_blow_count) Then
                            .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 7 + (soilCount - 1)).Value = CType(dpSL.spt_blow_count, Integer)
                        Else .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 7 + (soilCount - 1)).ClearContents
                        End If

                        soilCount += 1
                    Next

                Next

                dpRow += 1
                colCounter += 1

            Next


        End With

        SaveAndCloseDrilledPier()
    End Sub

    Private Function GetExcelColumnName(columnNumber As Integer) As String
        Dim dividend As Integer = columnNumber
        Dim columnName As String = String.Empty
        Dim modulo As Integer

        While dividend > 0
            modulo = (dividend - 1) Mod 26
            columnName = Convert.ToChar(65 + modulo).ToString() & columnName
            dividend = CInt((dividend - modulo) / 26)
        End While

        Return columnName
    End Function

    Private Sub LoadNewDrilledPier()
        NewDrilledPierWb.LoadDocument(DrilledPierTemplatePath, DrilledPierFileType)
        NewDrilledPierWb.BeginUpdate()
    End Sub

    Private Sub SaveAndCloseDrilledPier()
        NewDrilledPierWb.EndUpdate()
        NewDrilledPierWb.SaveDocument(ExcelFilePath, DrilledPierFileType)
    End Sub
#End Region

#Region "SQL Insert Statements"
    Private Function InsertDrilledPierDetail(ByVal dp As DrilledPier) As String
        Dim insertString As String = ""

        insertString += "@FndID"
        insertString += "," & IIf(IsNothing(dp.foundation_depth), "Null", "'" & dp.foundation_depth.ToString & "'")
        insertString += "," & IIf(IsNothing(dp.extension_above_grade), "Null", "'" & dp.extension_above_grade.ToString & "'")
        insertString += "," & IIf(IsNothing(dp.groundwater_depth), "Null", "'" & dp.groundwater_depth.ToString & "'")
        insertString += "," & IIf(IsNothing(dp.assume_min_steel), "Null", "'" & dp.assume_min_steel.ToString & "'")
        insertString += "," & IIf(IsNothing(dp.check_shear_along_depth), "Null", "'" & dp.check_shear_along_depth.ToString & "'")
        insertString += "," & IIf(IsNothing(dp.utilize_shear_friction_methodology), "Null", "'" & dp.utilize_shear_friction_methodology.ToString & "'")
        insertString += "," & IIf(IsNothing(dp.embedded_pole), "Null", "'" & dp.embedded_pole.ToString & "'")
        insertString += "," & IIf(IsNothing(dp.belled_pier), "Null", "'" & dp.belled_pier.ToString & "'")
        insertString += "," & IIf(IsNothing(dp.soil_layer_quantity), "Null", "'" & dp.soil_layer_quantity.ToString & "'")
        insertString += "," & IIf(IsNothing(dp.concrete_compressive_strength), "Null", "'" & dp.concrete_compressive_strength.ToString & "'")
        insertString += "," & IIf(IsNothing(dp.tie_yield_strength), "Null", "'" & dp.tie_yield_strength.ToString & "'")
        insertString += "," & IIf(IsNothing(dp.longitudinal_rebar_yield_strength), "Null", "'" & dp.longitudinal_rebar_yield_strength.ToString & "'")
        insertString += "," & IIf(IsNothing(dp.rebar_effective_depths), "Null", "'" & dp.rebar_effective_depths.ToString & "'")
        insertString += "," & IIf(IsNothing(dp.rebar_cage_2_fy_override), "Null", "'" & dp.rebar_cage_2_fy_override.ToString & "'")
        insertString += "," & IIf(IsNothing(dp.rebar_cage_3_fy_override), "Null", "'" & dp.rebar_cage_3_fy_override.ToString & "'")
        insertString += "," & IIf(IsNothing(dp.shear_override_crit_depth), "Null", "'" & dp.shear_override_crit_depth.ToString & "'")
        insertString += "," & IIf(IsNothing(dp.shear_crit_depth_override_comp), "Null", "'" & dp.shear_crit_depth_override_comp.ToString & "'")
        insertString += "," & IIf(IsNothing(dp.shear_crit_depth_override_uplift), "Null", "'" & dp.shear_crit_depth_override_uplift.ToString & "'")
        insertString += "," & IIf(IsNothing(dp.drilled_pier_profile_qty), "Null", "'" & dp.drilled_pier_profile_qty.ToString & "'")
        insertString += "," & IIf(IsNothing(dp.soil_profiles), "Null", "'" & dp.soil_profiles.ToString & "'")
        insertString += "," & IIf(IsNothing(dp.local_drilled_pier_id), "Null", "'" & dp.local_drilled_pier_id.ToString & "'")
        insertString += "," & IIf(IsNothing(dp.rho_override_1), "Null", "'" & dp.rho_override_1.ToString & "'")
        insertString += "," & IIf(IsNothing(dp.rho_override_2), "Null", "'" & dp.rho_override_2.ToString & "'")
        insertString += "," & IIf(IsNothing(dp.rho_override_3), "Null", "'" & dp.rho_override_3.ToString & "'")
        insertString += "," & IIf(IsNothing(dp.rho_override_4), "Null", "'" & dp.rho_override_4.ToString & "'")
        insertString += "," & IIf(IsNothing(dp.rho_override_5), "Null", "'" & dp.rho_override_5.ToString & "'")
        insertString += "," & IIf(IsNothing(dp.bearing_type_toggle), "Null", "'" & dp.bearing_type_toggle.ToString & "'")

        Return insertString
    End Function

    Private Function InsertDrilledPierBell(ByVal bp As DrilledPierBelledPier) As String
        Dim insertString As String = ""

        insertString += "@DpID"
        insertString += "," & IIf(IsNothing(bp.belled_pier_option), "Null", "'" & bp.belled_pier_option.ToString & "'")
        insertString += "," & IIf(IsNothing(bp.bottom_diameter_of_bell), "Null", "'" & bp.bottom_diameter_of_bell.ToString & "'")
        insertString += "," & IIf(IsNothing(bp.bell_input_type), "Null", "'" & bp.bell_input_type.ToString & "'")
        insertString += "," & IIf(IsNothing(bp.bell_angle), "Null", "'" & bp.bell_angle.ToString & "'")
        insertString += "," & IIf(IsNothing(bp.bell_height), "Null", "'" & bp.bell_height.ToString & "'")
        insertString += "," & IIf(IsNothing(bp.bell_toe_height), "Null", "'" & bp.bell_toe_height.ToString & "'")
        insertString += "," & IIf(IsNothing(bp.neglect_top_soil_layer), "Null", "'" & bp.neglect_top_soil_layer.ToString & "'")
        insertString += "," & IIf(IsNothing(bp.swelling_expansive_soil), "Null", "'" & bp.swelling_expansive_soil.ToString & "'")
        insertString += "," & IIf(IsNothing(bp.depth_of_expansive_soil), "Null", "'" & bp.depth_of_expansive_soil.ToString & "'")
        insertString += "," & IIf(IsNothing(bp.expansive_soil_force), "Null", "'" & bp.expansive_soil_force.ToString & "'")
        insertString += "," & IIf(IsNothing(bp.local_drilled_pier_id), "Null", "'" & bp.local_drilled_pier_id.ToString & "'")

        Return insertString
    End Function

    Private Function InsertDrilledPierEmbed(ByVal ep As DrilledPierEmbeddedPier) As String
        Dim insertString As String = ""

        insertString += "@DpID"
        insertString += "," & IIf(IsNothing(ep.embedded_pole_option), "Null", "'" & ep.embedded_pole_option.ToString & "'")
        insertString += "," & IIf(IsNothing(ep.encased_in_concrete), "Null", "'" & ep.encased_in_concrete.ToString & "'")
        insertString += "," & IIf(IsNothing(ep.pole_side_quantity), "Null", "'" & ep.pole_side_quantity.ToString & "'")
        insertString += "," & IIf(IsNothing(ep.pole_yield_strength), "Null", "'" & ep.pole_yield_strength.ToString & "'")
        insertString += "," & IIf(IsNothing(ep.pole_thickness), "Null", "'" & ep.pole_thickness.ToString & "'")
        insertString += "," & IIf(IsNothing(ep.embedded_pole_input_type), "Null", "'" & ep.embedded_pole_input_type.ToString & "'")
        insertString += "," & IIf(IsNothing(ep.pole_diameter_toc), "Null", "'" & ep.pole_diameter_toc.ToString & "'")
        insertString += "," & IIf(IsNothing(ep.pole_top_diameter), "Null", "'" & ep.pole_top_diameter.ToString & "'")
        insertString += "," & IIf(IsNothing(ep.pole_bottom_diameter), "Null", "'" & ep.pole_bottom_diameter.ToString & "'")
        insertString += "," & IIf(IsNothing(ep.pole_section_length), "Null", "'" & ep.pole_section_length.ToString & "'")
        insertString += "," & IIf(IsNothing(ep.pole_taper_factor), "Null", "'" & ep.pole_taper_factor.ToString & "'")
        insertString += "," & IIf(IsNothing(ep.pole_bend_radius_override), "Null", "'" & ep.pole_bend_radius_override.ToString & "'")
        insertString += "," & IIf(IsNothing(ep.local_drilled_pier_id), "Null", "'" & ep.local_drilled_pier_id.ToString & "'")

        Return insertString
    End Function

    Private Function InsertDrilledPierSoilLayer(ByVal dpsl As DrilledPierSoilLayer) As String
        Dim insertString As String = ""

        insertString += "@DpID"
        insertString += "," & IIf(IsNothing(dpsl.bottom_depth), "Null", "'" & dpsl.bottom_depth.ToString & "'")
        insertString += "," & IIf(IsNothing(dpsl.effective_soil_density), "Null", "'" & dpsl.effective_soil_density.ToString & "'")
        insertString += "," & IIf(IsNothing(dpsl.cohesion), "Null", "'" & dpsl.cohesion.ToString & "'")
        insertString += "," & IIf(IsNothing(dpsl.friction_angle), "Null", "'" & dpsl.friction_angle.ToString & "'")
        insertString += "," & IIf(IsNothing(dpsl.skin_friction_override_comp), "Null", "'" & dpsl.skin_friction_override_comp.ToString & "'")
        insertString += "," & IIf(IsNothing(dpsl.skin_friction_override_uplift), "Null", "'" & dpsl.skin_friction_override_uplift.ToString & "'")
        insertString += "," & IIf(IsNothing(dpsl.nominal_bearing_capacity), "Null", "'" & dpsl.nominal_bearing_capacity.ToString & "'")
        insertString += "," & IIf(IsNothing(dpsl.spt_blow_count), "Null", "'" & dpsl.spt_blow_count.ToString & "'")
        insertString += "," & IIf(IsNothing(dpsl.local_soil_layer_id), "Null", "'" & dpsl.local_soil_layer_id.ToString & "'")
        insertString += "," & IIf(IsNothing(dpsl.local_drilled_pier_id), "Null", "'" & dpsl.local_drilled_pier_id.ToString & "'")

        Return insertString
    End Function

    Private Function InsertDrilledPierSection(ByVal dpsec As DrilledPierSection) As String
        Dim insertString As String = ""

        insertString += "@DpID"
        insertString += "," & IIf(IsNothing(dpsec.pier_diameter), "Null", "'" & dpsec.pier_diameter.ToString & "'")
        insertString += "," & IIf(IsNothing(dpsec.clear_cover), "Null", "'" & dpsec.clear_cover.ToString & "'")
        insertString += "," & IIf(IsNothing(dpsec.clear_cover_rebar_cage_option), "Null", "'" & dpsec.clear_cover_rebar_cage_option.ToString & "'")
        insertString += "," & IIf(IsNothing(dpsec.tie_size), "Null", "'" & dpsec.tie_size.ToString & "'")
        insertString += "," & IIf(IsNothing(dpsec.tie_spacing), "Null", "'" & dpsec.tie_spacing.ToString & "'")
        insertString += "," & IIf(IsNothing(dpsec.bottom_elevation), "Null", "'" & dpsec.bottom_elevation.ToString & "'")
        'insertString += "," & IIf(IsNothing(dpsec.assume_min_steel_rho_override), "Null", "'" & dpsec.assume_min_steel_rho_override.ToString & "'")
        insertString += "," & IIf(IsNothing(dpsec.local_section_id), "Null", "'" & dpsec.local_section_id.ToString & "'")
        insertString += "," & IIf(IsNothing(dpsec.local_drilled_pier_id), "Null", "'" & dpsec.local_drilled_pier_id.ToString & "'")

        Return insertString
    End Function

    Private Function InsertDrilledPierRebar(ByVal dpreb As DrilledPierRebar) As String
        Dim insertString As String = ""

        insertString += "@SecID"
        insertString += "," & IIf(IsNothing(dpreb.longitudinal_rebar_quantity), "Null", "'" & dpreb.longitudinal_rebar_quantity.ToString & "'")
        insertString += "," & IIf(IsNothing(dpreb.longitudinal_rebar_size), "Null", "'" & dpreb.longitudinal_rebar_size.ToString & "'")
        insertString += "," & IIf(IsNothing(dpreb.longitudinal_rebar_cage_diameter), "Null", "'" & dpreb.longitudinal_rebar_cage_diameter.ToString & "'")
        insertString += "," & IIf(IsNothing(dpreb.local_rebar_id), "Null", "'" & dpreb.local_rebar_id.ToString & "'")
        insertString += "," & IIf(IsNothing(dpreb.local_drilled_pier_id), "Null", "'" & dpreb.local_drilled_pier_id.ToString & "'")
        insertString += "," & IIf(IsNothing(dpreb.local_section_id), "Null", "'" & dpreb.local_section_id.ToString & "'")

        Return insertString
    End Function
    Private Function InsertDrilledPierProfile(ByVal dpp As DrilledPierProfile) As String
        Dim insertString As String = ""

        insertString += "@DpID"
        insertString += "," & IIf(IsNothing(dpp.local_drilled_pier_id), "Null", "'" & dpp.local_drilled_pier_id.ToString & "'")
        insertString += "," & IIf(IsNothing(dpp.reaction_position), "Null", "'" & dpp.reaction_position.ToString & "'")
        insertString += "," & IIf(IsNothing(dpp.reaction_location), "Null", "'" & dpp.reaction_location.ToString & "'")
        insertString += "," & IIf(IsNothing(dpp.drilled_pier_profile), "Null", "'" & dpp.drilled_pier_profile.ToString & "'")
        insertString += "," & IIf(IsNothing(dpp.soil_profile), "Null", "'" & dpp.soil_profile.ToString & "'")

        Return insertString
    End Function
#End Region

#Region "SQL Update Statements"
    Private Function UpdateDrilledPierDetail(ByVal dp As DrilledPier) As String
        Dim updateString As String = ""

        updateString += "UPDATE drilled_pier_details SET "
        updateString += "foundation_depth=" & IIf(IsNothing(dp.foundation_depth), "Null", "'" & dp.foundation_depth.ToString & "'")
        updateString += ", extension_above_grade=" & IIf(IsNothing(dp.extension_above_grade), "Null", "'" & dp.extension_above_grade.ToString & "'")
        updateString += ", groundwater_depth=" & IIf(IsNothing(dp.groundwater_depth), "Null", "'" & dp.groundwater_depth.ToString & "'")
        updateString += ", assume_min_steel=" & IIf(IsNothing(dp.assume_min_steel), "Null", "'" & dp.assume_min_steel.ToString & "'")
        updateString += ", check_shear_along_depth=" & IIf(IsNothing(dp.check_shear_along_depth), "Null", "'" & dp.check_shear_along_depth.ToString & "'")
        updateString += ", utilize_shear_friction_methodology=" & IIf(IsNothing(dp.utilize_shear_friction_methodology), "Null", "'" & dp.utilize_shear_friction_methodology.ToString & "'")
        updateString += ", embedded_pole=" & IIf(IsNothing(dp.embedded_pole), "Null", "'" & dp.embedded_pole.ToString & "'")
        updateString += ", belled_pier=" & IIf(IsNothing(dp.belled_pier), "Null", "'" & dp.belled_pier.ToString & "'")
        updateString += ", soil_layer_quantity=" & IIf(IsNothing(dp.soil_layer_quantity), "Null", "'" & dp.soil_layer_quantity.ToString & "'")
        updateString += ", concrete_compressive_strength=" & IIf(IsNothing(dp.concrete_compressive_strength), "Null", "'" & dp.concrete_compressive_strength.ToString & "'")
        updateString += ", tie_yield_strength=" & IIf(IsNothing("'" & dp.tie_yield_strength), "Null", "'" & dp.tie_yield_strength.ToString & "'")
        updateString += ", longitudinal_rebar_yield_strength=" & IIf(IsNothing(dp.longitudinal_rebar_yield_strength), "Null", "'" & dp.longitudinal_rebar_yield_strength.ToString & "'")
        updateString += ", rebar_effective_depths=" & IIf(IsNothing(dp.rebar_effective_depths), "Null", "'" & dp.rebar_effective_depths.ToString & "'")
        updateString += ", rebar_cage_2_fy_override=" & IIf(IsNothing(dp.rebar_cage_2_fy_override), "Null", "'" & dp.rebar_cage_2_fy_override.ToString & "'")
        updateString += ", rebar_cage_3_fy_override=" & IIf(IsNothing(dp.rebar_cage_3_fy_override), "Null", "'" & dp.rebar_cage_3_fy_override.ToString & "'")
        updateString += ", shear_override_crit_depth=" & IIf(IsNothing(dp.shear_override_crit_depth), "Null", "'" & dp.shear_override_crit_depth.ToString & "'")
        updateString += ", shear_crit_depth_override_comp=" & IIf(IsNothing(dp.shear_crit_depth_override_comp), "Null", "'" & dp.shear_crit_depth_override_comp.ToString & "'")
        updateString += ", shear_crit_depth_override_uplift=" & IIf(IsNothing(dp.shear_crit_depth_override_uplift), "Null", "'" & dp.shear_crit_depth_override_uplift.ToString & "'")
        updateString += ", drilled_pier_profile_qty=" & IIf(IsNothing(dp.drilled_pier_profile_qty), "Null", "'" & dp.drilled_pier_profile_qty.ToString & "'")
        updateString += ", soil_profiles=" & IIf(IsNothing(dp.soil_profiles), "Null", "'" & dp.soil_profiles.ToString & "'")
        updateString += ", local_drilled_pier_id=" & IIf(IsNothing(dp.local_drilled_pier_id), "Null", "'" & dp.local_drilled_pier_id.ToString & "'")
        updateString += ", rho_override_1=" & IIf(IsNothing(dp.rho_override_1), "Null", "'" & dp.rho_override_1.ToString & "'")
        updateString += ", rho_override_2=" & IIf(IsNothing(dp.rho_override_2), "Null", "'" & dp.rho_override_2.ToString & "'")
        updateString += ", rho_override_3=" & IIf(IsNothing(dp.rho_override_3), "Null", "'" & dp.rho_override_3.ToString & "'")
        updateString += ", rho_override_4=" & IIf(IsNothing(dp.rho_override_4), "Null", "'" & dp.rho_override_4.ToString & "'")
        updateString += ", rho_override_5=" & IIf(IsNothing(dp.rho_override_5), "Null", "'" & dp.rho_override_5.ToString & "'")
        updateString += ", bearing_type_toggle=" & IIf(IsNothing(dp.bearing_type_toggle), "Null", "'" & dp.bearing_type_toggle.ToString & "'")
        updateString += " WHERE ID=" & dp.pier_id & vbNewLine

        Return updateString
    End Function

    Private Function UpdateDrilledPierBell(ByVal bp As DrilledPierBelledPier) As String
        Dim updateString As String = ""

        updateString += "UPDATE belled_pier_details SET "
        updateString += "belled_pier_option=" & IIf(IsNothing(bp.belled_pier_option), "Null", "'" & bp.belled_pier_option.ToString & "'")
        updateString += ", bottom_diameter_of_bell=" & IIf(IsNothing(bp.bottom_diameter_of_bell), "Null", "'" & bp.bottom_diameter_of_bell.ToString & "'")
        updateString += ", bell_input_type=" & IIf(IsNothing(bp.bell_input_type), "Null", "'" & bp.bell_input_type.ToString & "'")
        updateString += ", bell_angle=" & IIf(IsNothing(bp.bell_angle), "Null", "'" & bp.bell_angle.ToString & "'")
        updateString += ", bell_height=" & IIf(IsNothing(bp.bell_height), "Null", "'" & bp.bell_height.ToString & "'")
        updateString += ", bell_toe_height=" & IIf(IsNothing(bp.bell_toe_height), "Null", "'" & bp.bell_toe_height.ToString & "'")
        updateString += ", neglect_top_soil_layer=" & IIf(IsNothing(bp.neglect_top_soil_layer), "Null", "'" & bp.neglect_top_soil_layer.ToString & "'")
        updateString += ", swelling_expansive_soil=" & IIf(IsNothing(bp.swelling_expansive_soil), "Null", "'" & bp.swelling_expansive_soil.ToString & "'")
        updateString += ", depth_of_expansive_soil=" & IIf(IsNothing(bp.depth_of_expansive_soil), "Null", "'" & bp.depth_of_expansive_soil.ToString & "'")
        updateString += ", expansive_soil_force=" & IIf(IsNothing(bp.expansive_soil_force), "Null", "'" & bp.expansive_soil_force.ToString & "'")
        updateString += ", local_drilled_pier_id=" & IIf(IsNothing(bp.local_drilled_pier_id), "Null", "'" & bp.local_drilled_pier_id.ToString & "'")
        updateString += " WHERE ID=" & bp.belled_pier_id & vbNewLine

        Return updateString
    End Function

    Private Function UpdateDrilledPierEmbed(ByVal ep As DrilledPierEmbeddedPier) As String
        Dim updateString As String = ""

        updateString += "UPDATE embedded_pole_details SET "
        updateString += "embedded_pole_option=" & IIf(IsNothing(ep.embedded_pole_option), "Null", "'" & ep.embedded_pole_option.ToString & "'")
        updateString += ", encased_in_concrete=" & IIf(IsNothing(ep.encased_in_concrete), "Null", "'" & ep.encased_in_concrete.ToString & "'")
        updateString += ", pole_side_quantity=" & IIf(IsNothing(ep.pole_side_quantity), "Null", "'" & ep.pole_side_quantity.ToString & "'")
        updateString += ", pole_yield_strength=" & IIf(IsNothing(ep.pole_yield_strength), "Null", "'" & ep.pole_yield_strength.ToString & "'")
        updateString += ", pole_thickness=" & IIf(IsNothing(ep.pole_thickness), "Null", "'" & ep.pole_thickness.ToString & "'")
        updateString += ", embedded_pole_input_type=" & IIf(IsNothing(ep.embedded_pole_input_type), "Null", "'" & ep.embedded_pole_input_type.ToString & "'")
        updateString += ", pole_diameter_toc=" & IIf(IsNothing(ep.pole_diameter_toc), "Null", "'" & ep.pole_diameter_toc.ToString & "'")
        updateString += ", pole_top_diameter=" & IIf(IsNothing(ep.pole_top_diameter), "Null", "'" & ep.pole_top_diameter.ToString & "'")
        updateString += ", pole_bottom_diameter=" & IIf(IsNothing(ep.pole_bottom_diameter), "Null", "'" & ep.pole_bottom_diameter.ToString & "'")
        updateString += ", pole_section_length=" & IIf(IsNothing(ep.pole_section_length), "Null", "'" & ep.pole_section_length.ToString & "'")
        updateString += ", pole_taper_factor=" & IIf(IsNothing(ep.pole_taper_factor), "Null", "'" & ep.pole_taper_factor.ToString & "'")
        updateString += ", pole_bend_radius_override=" & IIf(IsNothing(ep.pole_bend_radius_override), "Null", "'" & ep.pole_bend_radius_override.ToString & "'")
        updateString += ", local_drilled_pier_id=" & IIf(IsNothing(ep.local_drilled_pier_id), "Null", "'" & ep.local_drilled_pier_id.ToString & "'")
        updateString += " WHERE ID=" & ep.embedded_id & vbNewLine

        Return updateString
    End Function

    Private Function UpdateDrilledPierSoilLayer(ByVal dpsl As DrilledPierSoilLayer) As String
        Dim updateString As String = ""

        updateString += "UPDATE drilled_pier_soil_layer SET "
        updateString += "bottom_depth=" & IIf(IsNothing(dpsl.bottom_depth), "Null", "'" & dpsl.bottom_depth.ToString & "'")
        updateString += ", effective_soil_density=" & IIf(IsNothing(dpsl.effective_soil_density), "Null", "'" & dpsl.effective_soil_density.ToString & "'")
        updateString += ", cohesion=" & IIf(IsNothing(dpsl.cohesion), "Null", "'" & dpsl.cohesion.ToString & "'")
        updateString += ", friction_angle=" & IIf(IsNothing(dpsl.friction_angle), "Null", "'" & dpsl.friction_angle.ToString & "'")
        updateString += ", skin_friction_override_comp=" & IIf(IsNothing(dpsl.skin_friction_override_comp), "Null", "'" & dpsl.skin_friction_override_comp.ToString & "'")
        updateString += ", skin_friction_override_uplift=" & IIf(IsNothing(dpsl.skin_friction_override_uplift), "Null", "'" & dpsl.skin_friction_override_uplift.ToString & "'")
        updateString += ", nominal_bearing_capacity=" & IIf(IsNothing(dpsl.nominal_bearing_capacity), "Null", "'" & dpsl.nominal_bearing_capacity.ToString & "'")
        updateString += ", spt_blow_count=" & IIf(IsNothing(dpsl.spt_blow_count), "Null", "'" & dpsl.spt_blow_count.ToString & "'")
        updateString += ", local_soil_layer_id=" & IIf(IsNothing(dpsl.local_soil_layer_id), "Null", "'" & dpsl.local_soil_layer_id.ToString & "'")
        updateString += ", local_drilled_pier_id=" & IIf(IsNothing(dpsl.local_drilled_pier_id), "Null", "'" & dpsl.local_drilled_pier_id.ToString & "'")
        updateString += " WHERE ID=" & dpsl.soil_layer_id & vbNewLine

        Return updateString
    End Function

    Private Function UpdateDrilledPierSection(ByVal dpsec As DrilledPierSection) As String
        Dim updateString As String = ""

        updateString += "UPDATE drilled_pier_section SET "
        updateString += "pier_diameter=" & IIf(IsNothing(dpsec.pier_diameter), "Null", "'" & dpsec.pier_diameter.ToString & "'")
        updateString += ", clear_cover=" & IIf(IsNothing(dpsec.clear_cover), "Null", "'" & dpsec.clear_cover.ToString & "'")
        updateString += ", clear_cover_rebar_cage_option=" & IIf(IsNothing(dpsec.clear_cover_rebar_cage_option), "Null", "'" & dpsec.clear_cover_rebar_cage_option.ToString & "'")
        updateString += ", tie_size=" & IIf(IsNothing(dpsec.tie_size), "Null", "'" & dpsec.tie_size.ToString & "'")
        updateString += ", tie_spacing=" & IIf(IsNothing(dpsec.tie_spacing), "Null", "'" & dpsec.tie_spacing.ToString & "'")
        updateString += ", bottom_elevation=" & IIf(IsNothing(dpsec.bottom_elevation), "Null", "'" & dpsec.bottom_elevation.ToString & "'")
        'updateString += ", assume_min_steel_rho_override=" & IIf(IsNothing(dpsec.assume_min_steel_rho_override), "Null", "'" & dpsec.assume_min_steel_rho_override.ToString & "'")
        updateString += ", local_section_id=" & IIf(IsNothing(dpsec.local_section_id), "Null", "'" & dpsec.local_section_id.ToString & "'")
        updateString += ", local_drilled_pier_id=" & IIf(IsNothing(dpsec.local_drilled_pier_id), "Null", "'" & dpsec.local_drilled_pier_id.ToString & "'")
        updateString += " WHERE ID=" & dpsec.section_id & vbNewLine

        Return updateString
    End Function

    Private Function UpdateDrilledPierRebar(ByVal dpreb As DrilledPierRebar) As String
        Dim updateString As String = ""

        updateString += "UPDATE drilled_pier_rebar SET "
        updateString += "longitudinal_rebar_quantity=" & IIf(IsNothing(dpreb.longitudinal_rebar_quantity), "Null", "'" & dpreb.longitudinal_rebar_quantity.ToString & "'")
        updateString += ", longitudinal_rebar_size=" & IIf(IsNothing(dpreb.longitudinal_rebar_size), "Null", "'" & dpreb.longitudinal_rebar_size.ToString & "'")
        updateString += ", longitudinal_rebar_cage_diameter=" & IIf(IsNothing(dpreb.longitudinal_rebar_cage_diameter), "Null", "'" & dpreb.longitudinal_rebar_cage_diameter.ToString & "'")
        updateString += ", local_rebar_id=" & IIf(IsNothing(dpreb.local_rebar_id), "Null", "'" & dpreb.local_rebar_id.ToString & "'")
        updateString += ", local_drilled_pier_id=" & IIf(IsNothing(dpreb.local_drilled_pier_id), "Null", "'" & dpreb.local_drilled_pier_id.ToString & "'")
        updateString += ", local_section_id=" & IIf(IsNothing(dpreb.local_section_id), "Null", "'" & dpreb.local_section_id.ToString & "'")
        updateString += " WHERE ID=" & dpreb.rebar_id & vbNewLine

        Return updateString
    End Function

    Private Function UpdateDrilledPierProfile(ByVal dpp As DrilledPierProfile) As String
        Dim updateString As String = ""

        updateString += "UPDATE drilled_pier_profile SET "
        updateString += "local_drilled_pier_id=" & IIf(IsNothing(dpp.local_drilled_pier_id), "Null", "'" & dpp.local_drilled_pier_id.ToString & "'")
        updateString += ", reaction_position=" & IIf(IsNothing(dpp.reaction_position), "Null", "'" & dpp.reaction_position.ToString & "'")
        updateString += ", reaction_location=" & IIf(IsNothing(dpp.reaction_location), "Null", "'" & dpp.reaction_location.ToString & "'")
        updateString += ", drilled_pier_profile=" & IIf(IsNothing(dpp.drilled_pier_profile), "Null", "'" & dpp.drilled_pier_profile.ToString & "'")
        updateString += ", soil_profile=" & IIf(IsNothing(dpp.soil_profile), "Null", "'" & dpp.soil_profile.ToString & "'")
        updateString += " WHERE ID=" & dpp.profile_id & vbNewLine

        Return updateString
    End Function
#End Region

#Region "General"
    Public Sub Clear()
        ExcelFilePath = ""
        DrilledPiers.Clear()
    End Sub

    Private Function DrilledPierSQLDataTables() As List(Of SQLParameter)
        Dim MyParameters As New List(Of SQLParameter)

        MyParameters.Add(New SQLParameter("Drilled Pier General Details SQL", "Drilled Piers (SELECT Details).sql"))
        MyParameters.Add(New SQLParameter("Drilled Pier Section SQL", "Drilled Piers (SELECT Section).sql"))
        MyParameters.Add(New SQLParameter("Drilled Pier Rebar SQL", "Drilled Piers (SELECT Rebar).sql"))
        MyParameters.Add(New SQLParameter("Drilled Pier Soil SQL", "Drilled Piers (SELECT Soil Layers).sql"))
        MyParameters.Add(New SQLParameter("Belled Details SQL", "Drilled Piers (SELECT Belled).sql"))
        MyParameters.Add(New SQLParameter("Embedded Details SQL", "Drilled Piers (SELECT Embedded).sql"))
        MyParameters.Add(New SQLParameter("Drilled Pier Profiles SQL", "Drilled Piers (SELECT Profile).sql"))

        Return MyParameters
    End Function

    Private Function DrilledPierExcelDTParameters() As List(Of EXCELDTParameter)
        Dim MyParameters As New List(Of EXCELDTParameter)

        MyParameters.Add(New EXCELDTParameter("Drilled Pier General Details EXCEL", "A2:AC1000", "Details (ENTER)"))
        MyParameters.Add(New EXCELDTParameter("Drilled Pier Section EXCEL", "A2:J1000", "Sections (ENTER)"))
        MyParameters.Add(New EXCELDTParameter("Drilled Pier Rebar EXCEL", "A2:I1000", "Rebar (ENTER)"))
        MyParameters.Add(New EXCELDTParameter("Drilled Pier Soil EXCEL", "A2:L1502", "Soil Layers (ENTER)")) 'use range of 1000 to be safe that multiple generations of EDS values are brought in. This range need to go to 1500 values to match the tool's limit
        MyParameters.Add(New EXCELDTParameter("Belled Details EXCEL", "A2:M1000", "Belled (ENTER)"))
        MyParameters.Add(New EXCELDTParameter("Embedded Details EXCEL", "A2:O1000", "Embedded (ENTER)"))
        MyParameters.Add(New EXCELDTParameter("Drilled Pier Profiles EXCEL", "A2:G1000", "Profiles (ENTER)"))

        Return MyParameters
    End Function
#End Region

End Class