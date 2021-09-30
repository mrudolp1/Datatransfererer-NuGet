Option Strict Off

Imports DevExpress.Spreadsheet
Imports System.Security.Principal

Partial Public Class DataTransfererCCIpole

    '#Region "Define"
    '    Private NewCCIpoleWb As New Workbook
    '    Private prop_ExcelFilePath As String

    '    Public Property Poles As New List(Of CCIpole)
    '    Private Property CCIpoleTemplatePath As String = "C:\Users\" & Environment.UserName & "\source\repos\Datatransferer NuGet\Reference\CCIpole (4.6.0) - TEMPLATE.xlsm"
    '    Private Property CCIpoleFileType As DocumentFormat = DocumentFormat.Xlsm

    '    Public Property poleDB As String
    '    Public Property poleID As WindowsIdentity

    '    Public Property ExcelFilePath() As String
    '        Get
    '            Return Me.prop_ExcelFilePath
    '        End Get
    '        Set
    '            Me.prop_ExcelFilePath = Value
    '        End Set
    '    End Property

    '    Public Property xlApp As Object
    '#End Region

    '#Region "Constructors"
    '    Public Sub New()
    '        'Leave method empty
    '    End Sub

    '    Public Sub New(ByVal MyDataSet As DataSet, ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String, ByVal BU As String, ByVal Strucutre_ID As String)
    '        'dpDS = MyDataSet
    '        ds = MyDataSet
    '        poleID = LogOnUser
    '        poleDB = ActiveDatabase
    '        'BUNumber = BU 'Need to turn back on when connecting to dashboard. Turned off for testing.
    '        'STR_ID = Strucutre_ID 'Need to turn back on when connecting to dashboard. Turned off for testing.
    '    End Sub
    '#End Region

    '#Region "Load Data"
    '    Public Function LoadFromEDS() As Boolean
    '        Dim refid As Integer

    '        Dim CCIpoleLoader As String

    '        'Load data to get pier and pad details data for the existing structure model
    '        For Each item As SQLParameter In CCIpoleSQLDataTables()
    '            CCIpoleLoader = QueryBuilderFromFile(queryPath & "CCIpole\" & item.sqlQuery).Replace("[EXISTING MODEL]", GetExistingModelQuery())
    '            DoDaSQL.sqlLoader(CCIpoleLoader, item.sqlDatatable, ds, poleDB, poleID, "0")
    '            'If ds.Tables(item.sqlDatatable).Rows.Count = 0 Then Return False 'This may need adjusted since some tables can be empty
    '        Next

    '        'Custom Section to transfer data for the drilled pier tool. Needs to be adjusted for each tool.
    '        For Each CCIpoleDataRow As DataRow In ds.Tables("CCIpole General Details SQL").Rows
    '            refid = CType(CCIpoleDataRow.Item("pole_structure_id"), Integer)

    '            Poles.Add(New CCIpole(CCIpoleDataRow, refid))
    '        Next

    '        Return True
    '    End Function 'Create Drilled Pier objects based on what is saved in EDS

    '    Public Sub LoadFromExcel()
    '        Dim refID As Integer
    '        Dim refCol As String

    '        For Each item As EXCELDTParameter In CCIpoleExcelDTParameters()
    '            'Get tables from excel file 
    '            ds.Tables.Add(ExcelDatasourceToDataTable(GetExcelDataSource(ExcelFilePath, item.xlsSheet, item.xlsRange), item.xlsDatatable))
    '        Next

    '        'Custom Section to transfer data for the drilled pier tool. Needs to be adjusted for each tool.
    '        For Each CCIpoleDataRow As DataRow In ds.Tables("CCIpole General Details EXCEL").Rows
    '            'If DrilledPierDataRow.Item("foudation_id").ToString = "" Then
    '            '    refCol = "local_drilled_pier_id"
    '            '    refID = CType(DrilledPierDataRow.Item(refCol), Integer)
    '            'Else
    '            '    refCol = "drilled_pier_id"
    '            '    refID = CType(DrilledPierDataRow.Item(refCol), Integer)
    '            'End If
    '            ''commented out in case drilled pier id and local drilled pier id matched, prevents possible overriding of data
    '            refCol = "pole_structure_id"
    '            refID = CType(CCIpoleDataRow.Item(refCol), Integer)

    '            Poles.Add(New CCIpole(CCIpoleDataRow, refID))
    '        Next
    '    End Sub 'Create Drilled Pier objects based on what is coming from the excel file
    '#End Region

    '#Region "Save Data"
    '    Public Sub SaveToEDS()
    '        Dim firstOne As Boolean = True
    '        Dim mySoils As String = ""
    '        Dim mySections As String = ""
    '        Dim myRebar As String = ""
    '        Dim myProfiles As String = ""

    '        For Each pole As CCIpole In Poles
    '            Dim DrilledPierSaver As String = QueryBuilderFromFile(queryPath & "Drilled Pier\Drilled Piers (IN_UP).sql")
    '            Dim dpSectionQuery As String = QueryBuilderFromFile(queryPath & "Drilled Pier\Drilled Piers Section (IN_UP).txt")

    '            DrilledPierSaver = DrilledPierSaver.Replace("[BU NUMBER]", BUNumber)
    '            DrilledPierSaver = DrilledPierSaver.Replace("[STRUCTURE ID]", STR_ID)
    '            DrilledPierSaver = DrilledPierSaver.Replace("[FOUNDATION TYPE]", "Drilled Pier")
    '            If dp.pier_id = 0 Or IsDBNull(dp.pier_id) Then
    '                DrilledPierSaver = DrilledPierSaver.Replace("'[DRILLED PIER ID]'", "NULL")
    '            Else
    '                DrilledPierSaver = DrilledPierSaver.Replace("'[DRILLED PIER ID]'", dp.pier_id.ToString)
    '            End If
    '            DrilledPierSaver = DrilledPierSaver.Replace("[EMBED BOOLEAN]", dp.embedded_pole.ToString)
    '            DrilledPierSaver = DrilledPierSaver.Replace("[BELL BOOLEAN]", dp.belled_pier.ToString)
    '            DrilledPierSaver = DrilledPierSaver.Replace("[INSERT ALL PIER DETAILS]", InsertDrilledPierDetail(dp))

    '            If dp.pier_id = 0 Or IsDBNull(dp.pier_id) Then
    '                For Each dpsl As DrilledPierSoilLayer In dp.soil_layers
    '                    Dim tempSoilLayer As String = InsertDrilledPierSoilLayer(dpsl)

    '                    If Not firstOne Then
    '                        mySoils += ",(" & tempSoilLayer & ")"
    '                    Else
    '                        mySoils += "(" & tempSoilLayer & ")"
    '                    End If

    '                    firstOne = False
    '                Next 'Add Soil Layer INSERT statments
    '                DrilledPierSaver = DrilledPierSaver.Replace("([INSERT ALL SOIL LAYERS])", mySoils)
    '                firstOne = True

    '                For Each dpsec As DrilledPierSection In dp.sections
    '                    Dim tempSection As String = dpSectionQuery.Replace("[DRILLED PIER SECTION]", InsertDrilledPierSection(dpsec))

    '                    For Each dpreb In dpsec.rebar
    '                        Dim temprebar As String = InsertDrilledPierRebar(dpreb)

    '                        If Not firstOne Then
    '                            myRebar += ",(" & temprebar & ")"
    '                        Else
    '                            myRebar += "(" & temprebar & ")"
    '                        End If

    '                        firstOne = False
    '                    Next 'Add Rebar INSERT Statements

    '                    tempSection = tempSection.Replace("([DRILLED PIER SECTION REBAR])", myRebar)
    '                    firstOne = True
    '                    myRebar = ""
    '                    mySections += tempSection + vbNewLine
    '                Next 'Add Section INSERT Statements
    '                DrilledPierSaver = DrilledPierSaver.Replace("--*[DRILLED PIER SECTIONS]*--", mySections)

    '                If dp.belled_pier Then
    '                    DrilledPierSaver = DrilledPierSaver.Replace("[INSERT ALL BELLED PIER DETAILS]", InsertDrilledPierBell(dp.belled_details))
    '                Else
    '                    DrilledPierSaver = DrilledPierSaver.Replace("BEGIN --Belled Pier", "--BEGIN --Belled Pier")
    '                    DrilledPierSaver = DrilledPierSaver.Replace("IF @IsBelled = 'True'", "--IF @IsBelled = 'True'")
    '                    DrilledPierSaver = DrilledPierSaver.Replace("INSERT INTO belled_pier_details VALUES ([INSERT ALL BELLED PIER DETAILS])", "")
    '                    DrilledPierSaver = DrilledPierSaver.Replace("END --INSERT Belled Pier information if required", "--END --INSERT Belled Pier information if required")
    '                End If 'Add Belled Pier INSERT Statment

    '                If dp.embedded_pole Then
    '                    DrilledPierSaver = DrilledPierSaver.Replace("[INSERT ALL EMBEDDED POLE DETAILS]", InsertDrilledPierEmbed(dp.embed_details))
    '                Else
    '                    DrilledPierSaver = DrilledPierSaver.Replace("BEGIN --Embedded Pole", "--BEGIN --Embedded Pole")
    '                    DrilledPierSaver = DrilledPierSaver.Replace("IF @IsEmbed = 'True'", "--IF @IsEmbed = 'True'")
    '                    DrilledPierSaver = DrilledPierSaver.Replace("INSERT INTO embedded_pole_details OUTPUT INSERTED.ID INTO @EmbeddedPole VALUES ([INSERT ALL EMBEDDED POLE DETAILS])", "")
    '                    DrilledPierSaver = DrilledPierSaver.Replace("SELECT @EmbedID=EmbedID FROM @EmbeddedPole", "--SELECT @EmbedID=EmbedID FROM @EmbeddedPole")
    '                    DrilledPierSaver = DrilledPierSaver.Replace("END --INSERT Embedded Pole information if required", "--END --INSERT Embedded Pole information if required")
    '                End If 'Add Embedded Pole INSERT Statment

    '                For Each dpp As DrilledPierProfile In dp.drilled_pier_profiles
    '                    Dim tempDrilledPierProfile As String = InsertDrilledPierProfile(dpp)

    '                    If Not firstOne Then
    '                        myProfiles += ",(" & tempDrilledPierProfile & ")"
    '                    Else
    '                        myProfiles += "(" & tempDrilledPierProfile & ")"
    '                    End If

    '                    firstOne = False
    '                Next 'Add Pier Profile INSERT statements
    '                DrilledPierSaver = DrilledPierSaver.Replace("([INSERT ALL PIER PROFILES])", myProfiles)
    '                firstOne = True

    '                mySoils = ""
    '                mySections = ""
    '                myProfiles = ""
    '            Else
    '                Dim tempUpdater As String = ""
    '                tempUpdater += UpdateDrilledPierDetail(dp)

    '                'comment out soil layer insertion. Added in next step if a layer does not have an ID
    '                DrilledPierSaver = DrilledPierSaver.Replace("INSERT INTO drilled_pier_soil_layer VALUES ([INSERT ALL SOIL LAYERS])", "--INSERT INTO drilled_pier_soil_layer VALUES ([INSERT ALL SOIL LAYERS])")

    '                For Each dpsl As DrilledPierSoilLayer In dp.soil_layers
    '                    If dpsl.soil_layer_id = 0 Or IsDBNull(dpsl.soil_layer_id) Then
    '                        tempUpdater += "INSERT INTO drilled_pier_soil_layers VALUES (" & InsertDrilledPierSoilLayer(dpsl) & ") " & vbNewLine
    '                    Else
    '                        tempUpdater += UpdateDrilledPierSoilLayer(dpsl)
    '                    End If
    '                Next

    '                If dp.belled_pier Then
    '                    If dp.belled_details.belled_pier_id = 0 Or IsDBNull(dp.belled_details.belled_pier_id) Then
    '                        tempUpdater += "INSERT INTO belled_pier_details VALUES (" & InsertDrilledPierBell(dp.belled_details) & ") " & vbNewLine
    '                    Else
    '                        tempUpdater += UpdateDrilledPierBell(dp.belled_details)
    '                    End If
    '                Else
    '                    DrilledPierSaver = DrilledPierSaver.Replace("BEGIN --Belled Pier", "--BEGIN --Belled Pier")
    '                    DrilledPierSaver = DrilledPierSaver.Replace("IF @IsBelled = 'True'", "--IF @IsBelled = 'True'")
    '                    DrilledPierSaver = DrilledPierSaver.Replace("INSERT INTO belled_pier_details VALUES ([INSERT ALL BELLED PIER DETAILS])", "")
    '                    DrilledPierSaver = DrilledPierSaver.Replace("END --INSERT Belled Pier information if required", "--END --INSERT Belled Pier information if required")
    '                End If

    '                If dp.embedded_pole Then
    '                    If dp.embed_details.embedded_id = 0 Or IsDBNull(dp.embed_details.embedded_id) Then
    '                        tempUpdater += "BEGIN INSERT INTO embedded_pole_details OUTPUT INSERTED.ID INTO @EmbeddedPole VALUES (" & InsertDrilledPierEmbed(dp.embed_details) & ") " & vbNewLine & " SELECT @EmbedID=EmbedID FROM @EmbeddedPole"
    '                        tempUpdater += " END " & vbNewLine
    '                    Else
    '                        tempUpdater += UpdateDrilledPierEmbed(dp.embed_details)
    '                    End If
    '                Else
    '                    DrilledPierSaver = DrilledPierSaver.Replace("BEGIN --Embedded Pole", "--BEGIN --Embedded Pole")
    '                    DrilledPierSaver = DrilledPierSaver.Replace("IF @IsEmbed = 'True'", "--IF @IsEmbed = 'True'")
    '                    DrilledPierSaver = DrilledPierSaver.Replace("INSERT INTO embedded_pole_details OUTPUT INSERTED.ID INTO @EmbeddedPole VALUES ([INSERT ALL EMBEDDED POLE DETAILS])", "")
    '                    DrilledPierSaver = DrilledPierSaver.Replace("SELECT @EmbedID=EmbedID FROM @EmbeddedPole", "--SELECT @EmbedID=EmbedID FROM @EmbeddedPole")
    '                    DrilledPierSaver = DrilledPierSaver.Replace("END --INSERT Embedded Pole information if required", "--END --INSERT Embedded Pole information if required")
    '                End If

    '                For Each dpSec As DrilledPierSection In dp.sections
    '                    If dpSec.section_id = 0 Or IsDBNull(dpSec.section_id) Then
    '                        tempUpdater += "BEGIN INSERT INTO drilled_pier_section OUTPUT INSERTED.ID INTO @DrilledPierSection VALUES (" & InsertDrilledPierSection(dpSec) & ") " & vbNewLine & " SELECT @SecID=SecID FROM @DrilledPierSection"
    '                        For Each dpreb As DrilledPierRebar In dpSec.rebar
    '                            tempUpdater += "INSERT INTO drilled_pier_rebar VALUES (" & InsertDrilledPierRebar(dpreb) & ") " & vbNewLine
    '                        Next
    '                        tempUpdater += " END " & vbNewLine
    '                    Else
    '                        tempUpdater += UpdateDrilledPierSection(dpSec)
    '                        For Each dpreb As DrilledPierRebar In dpSec.rebar
    '                            If dpreb.rebar_id = 0 Or IsDBNull(dpreb.rebar_id) Then
    '                                tempUpdater += "INSERT INTO drilled_pier_rebar VALUES (" & InsertDrilledPierRebar(dpreb).Replace("@SecID", dpSec.section_id.ToString) & ") " & vbNewLine
    '                            Else
    '                                tempUpdater += UpdateDrilledPierRebar(dpreb)
    '                            End If
    '                        Next
    '                    End If
    '                Next

    '                DrilledPierSaver = DrilledPierSaver.Replace("INSERT INTO drilled_pier_profile VALUES ([INSERT ALL PIER PROFILES])", "--INSERT INTO drilled_pier_profile VALUES ([INSERT ALL PIER PROFILES])")
    '                For Each dpp As DrilledPierProfile In dp.drilled_pier_profiles
    '                    If dpp.profile_id = 0 Or IsDBNull(dpp.profile_id) Then
    '                        tempUpdater += "INSERT INTO drilled_pier_profile VALUES (" & InsertDrilledPierProfile(dpp) & ") " & vbNewLine
    '                    Else
    '                        tempUpdater += UpdateDrilledPierProfile(dpp)
    '                    End If
    '                Next

    '                DrilledPierSaver = DrilledPierSaver.Replace("SELECT * FROM TEMPORARY", tempUpdater)
    '            End If

    '            DrilledPierSaver = DrilledPierSaver.Replace("[INSERT ALL PIER DETAILS DETAILS]", InsertDrilledPierDetail(dp))

    '            sqlSender(DrilledPierSaver, dpDB, dpID, "0")
    '        Next


    '    End Sub

    '    Public Sub SaveToExcel()
    '        Dim dpRow As Integer = 3
    '        Dim secRow As Integer = 3
    '        Dim rebRow As Integer = 3
    '        Dim soilRow As Integer = 3
    '        Dim profileRow As Integer = 3

    '        LoadNewDrilledPier()

    '        With NewDrilledPierWb

    '            Dim colCounter As Integer = 6
    '            Dim myCol As String
    '            Dim rowStart As Integer = 56

    '            For Each dp As DrilledPier In DrilledPiers

    '                colCounter = 6 + dp.local_drilled_pier_id
    '                myCol = GetExcelColumnName(colCounter)

    '                'DRILLED PIER DETAILS
    '                If Not IsNothing(dp.pier_id) Then
    '                    .Worksheets("Database").Range(myCol & rowStart - 54).Value = CType(dp.pier_id, Integer)
    '                Else .Worksheets("Database").Range(myCol & rowStart - 54).ClearContents
    '                End If
    '                If Not IsNothing(dp.concrete_compressive_strength) Then
    '                    .Worksheets("Database").Range(myCol & rowStart + 7).Value = CType(dp.concrete_compressive_strength, Double)
    '                Else .Worksheets("Database").Range(myCol & rowStart + 7).ClearContents
    '                End If
    '                If Not IsNothing(dp.longitudinal_rebar_yield_strength) Then
    '                    .Worksheets("Database").Range(myCol & rowStart + 8).Value = CType(dp.longitudinal_rebar_yield_strength, Double)
    '                Else .Worksheets("Database").Range(myCol & rowStart + 8).ClearContents
    '                End If
    '                If Not IsNothing(dp.tie_yield_strength) Then
    '                    .Worksheets("Database").Range(myCol & rowStart + 9).Value = CType(dp.tie_yield_strength, Double)
    '                Else .Worksheets("Database").Range(myCol & rowStart + 9).ClearContents
    '                End If
    '                If Not IsNothing(dp.foundation_depth) Then
    '                    .Worksheets("Database").Range(myCol & rowStart + 10).Value = CType(dp.foundation_depth, Double)
    '                Else .Worksheets("Database").Range(myCol & rowStart + 10).ClearContents
    '                End If
    '                If Not IsNothing(dp.extension_above_grade) Then
    '                    .Worksheets("Database").Range(myCol & rowStart + 11).Value = CType(dp.extension_above_grade, Double)
    '                Else .Worksheets("Database").Range(myCol & rowStart + 11).ClearContents
    '                End If
    '                If CType(dp.groundwater_depth, Double) = -1 Then
    '                    .Worksheets("Database").Range(myCol & rowStart + 17).Value = "N/A"
    '                ElseIf Not IsNothing(dp.groundwater_depth) Then
    '                    .Worksheets("Database").Range(myCol & rowStart + 17).Value = CType(dp.groundwater_depth, Double)
    '                Else .Worksheets("Database").Range(myCol & rowStart + 17).ClearContents
    '                End If
    '                If Not IsNothing(dp.soil_layer_quantity) Then
    '                    .Worksheets("Database").Range(myCol & rowStart + 18).Value = CType(dp.soil_layer_quantity, Double)
    '                Else .Worksheets("Database").Range(myCol & rowStart + 18).ClearContents
    '                End If
    '                If Not IsNothing(dp.bearing_type_toggle) Then
    '                    .Worksheets("Database").Range(myCol & rowStart + 19).Value = CType(dp.bearing_type_toggle, String)
    '                Else .Worksheets("Database").Range(myCol & rowStart + 19).ClearContents
    '                End If
    '                .Worksheets("Database").Range(myCol & rowStart + 97).Value = CType(dp.check_shear_along_depth, Boolean)
    '                .Worksheets("Database").Range(myCol & rowStart + 98).Value = CType(dp.utilize_shear_friction_methodology, Boolean)
    '                .Worksheets("Database").Range(myCol & rowStart + 100).Value = CType(dp.embedded_pole, Boolean)
    '                .Worksheets("Database").Range(myCol & rowStart + 112).Value = CType(dp.belled_pier, Boolean)
    '                .Worksheets("Database").Range(myCol & rowStart + 4389).Value = CType(dp.assume_min_steel, String)
    '                .Worksheets("Database").Range(myCol & rowStart + 4390).Value = CType(dp.rebar_effective_depths, Boolean)
    '                If Not IsNothing(dp.rebar_cage_2_fy_override) Then
    '                    .Worksheets("Database").Range(myCol & rowStart + 4391).Value = CType(dp.rebar_cage_2_fy_override, Double)
    '                Else .Worksheets("Database").Range(myCol & rowStart + 4391).ClearContents
    '                End If
    '                If Not IsNothing(dp.rebar_cage_3_fy_override) Then
    '                    .Worksheets("Database").Range(myCol & rowStart + 4392).Value = CType(dp.rebar_cage_3_fy_override, Double)
    '                Else .Worksheets("Database").Range(myCol & rowStart + 4392).ClearContents
    '                End If
    '                .Worksheets("Database").Range(myCol & rowStart + 99).Value = CType(dp.shear_override_crit_depth, Boolean)
    '                If Not IsNothing(dp.shear_crit_depth_override_comp) Then
    '                    .Worksheets("Database").Range(myCol & rowStart + 374).Value = CType(dp.shear_crit_depth_override_comp, Double)
    '                Else .Worksheets("Database").Range(myCol & rowStart + 374).Formula = .Worksheets("Database").Range(GetExcelColumnName(colCounter + 51) & rowStart + 374).Formula
    '                End If
    '                If Not IsNothing(dp.shear_crit_depth_override_uplift) Then
    '                    .Worksheets("Database").Range(myCol & rowStart + 376).Value = CType(dp.shear_crit_depth_override_uplift, Double)
    '                Else .Worksheets("Database").Range(myCol & rowStart + 376).Formula = .Worksheets("Database").Range(GetExcelColumnName(colCounter + 51) & rowStart + 376).Formula
    '                End If

    '                Dim depth As Integer = 0
    '                Dim secBump As Integer = 0
    '                Dim secStart As Integer = 20
    '                Dim secCount As Integer = 1

    '                'DRILLED PIER SECTION
    '                For Each dpSec As DrilledPierSection In dp.sections

    '                    If Not IsNothing(dpSec.section_id) Then
    '                        .Worksheets("Database").Range(myCol & rowStart - 54 + secCount).Value = CType(dpSec.section_id, Integer)
    '                    Else .Worksheets("Database").Range(myCol & rowStart - 54 + secCount).ClearContents
    '                    End If

    '                    If Not IsNothing(dpSec.pier_diameter) Then
    '                        .Worksheets("Database").Range(myCol & rowStart + secStart + 0).Value = CType(dpSec.pier_diameter, Double)
    '                    Else .Worksheets("Database").Range(myCol & rowStart + secStart + 0).ClearContents
    '                    End If
    '                    If Not IsNothing(dpSec.clear_cover) Then
    '                        .Worksheets("Database").Range(myCol & rowStart + secStart + 3).Value = CType(dpSec.clear_cover, Double)
    '                    Else .Worksheets("Database").Range(myCol & rowStart + secStart + 3).ClearContents
    '                    End If
    '                    If Not IsNothing(dpSec.tie_size) Then
    '                        .Worksheets("Database").Range(myCol & rowStart + secStart + 4).Value = CType(dpSec.tie_size, Integer)
    '                    Else .Worksheets("Database").Range(myCol & rowStart + secStart + 4).ClearContents
    '                    End If
    '                    If Not IsNothing(dpSec.tie_spacing) Then
    '                        .Worksheets("Database").Range(myCol & rowStart + secStart + 5).Value = CType(dpSec.tie_spacing, Double)
    '                    Else .Worksheets("Database").Range(myCol & rowStart + secStart + 5).ClearContents
    '                    End If
    '                    If Not IsNothing(dpSec.clear_cover_rebar_cage_option) Then
    '                        .Worksheets("Database").Range(myCol & rowStart + secStart + 14).Value = CType(dpSec.clear_cover_rebar_cage_option, String)
    '                    Else .Worksheets("Database").Range(myCol & rowStart + secStart + 14).ClearContents
    '                    End If
    '                    If Not IsNothing(dpSec.rho_override) Then
    '                        .Worksheets("Database").Range(myCol & rowStart + 4392 + secCount).Value = CType(dpSec.rho_override, Double)
    '                    Else .Worksheets("Database").Range(myCol & rowStart + 4392 + secCount).ClearContents
    '                    End If

    '                    If secCount > 1 Then depth += 1
    '                    If Not IsNothing(dpSec.bottom_elevation) Then
    '                        .Worksheets("Database").Range(myCol & rowStart + 12 + depth).Value = CType(dpSec.bottom_elevation, Double)
    '                    Else .Worksheets("Database").Range(myCol & rowStart + 12 + depth).ClearContents
    '                    End If

    '                    'DRILLED PIER REBAR
    '                    Dim rebCount As Integer = 1

    '                    For Each dpReb As DrilledPierRebar In dpSec.rebar

    '                        If Not IsNothing(dpReb.rebar_id) Then
    '                            .Worksheets("Database").Range(myCol & rowStart - 48 + 3 * (secCount - 1) + (rebCount - 1)).Value = CType(dpReb.rebar_id, Integer)
    '                        Else .Worksheets("Database").Range(myCol & rowStart - 48 + 3 * (secCount - 1) + (rebCount - 1)).ClearContents
    '                        End If

    '                        If rebCount = 1 Then
    '                            If Not IsNothing(dpReb.longitudinal_rebar_quantity) Then
    '                                .Worksheets("Database").Range(myCol & rowStart + secStart + 1).Value = CType(dpReb.longitudinal_rebar_quantity, Double)
    '                            Else .Worksheets("Database").Range(myCol & rowStart + secStart + 1).ClearContents
    '                            End If
    '                            If Not IsNothing(dpReb.longitudinal_rebar_size) Then
    '                                .Worksheets("Database").Range(myCol & rowStart + secStart + 2).Value = CType(dpReb.longitudinal_rebar_size, Double)
    '                            Else .Worksheets("Database").Range(myCol & rowStart + secStart + 2).ClearContents
    '                            End If
    '                        ElseIf rebCount = 2 Then
    '                            If Not IsNothing(dpReb.longitudinal_rebar_quantity) Then
    '                                .Worksheets("Database").Range(myCol & rowStart + secStart + 6).Value = CType(dpReb.longitudinal_rebar_quantity, Double)
    '                            Else .Worksheets("Database").Range(myCol & rowStart + secStart + 6).ClearContents
    '                            End If
    '                            If Not IsNothing(dpReb.longitudinal_rebar_size) Then
    '                                .Worksheets("Database").Range(myCol & rowStart + secStart + 7).Value = CType(dpReb.longitudinal_rebar_size, Double)
    '                            Else .Worksheets("Database").Range(myCol & rowStart + secStart + 7).ClearContents
    '                            End If
    '                            If Not IsNothing(dpReb.longitudinal_rebar_cage_diameter) Then
    '                                .Worksheets("Database").Range(myCol & rowStart + secStart + 8).Value = CType(dpReb.longitudinal_rebar_cage_diameter, Double)
    '                            Else .Worksheets("Database").Range(myCol & rowStart + secStart + 8).ClearContents
    '                            End If
    '                        ElseIf rebCount = 3 Then
    '                            If Not IsNothing(dpReb.longitudinal_rebar_quantity) Then
    '                                .Worksheets("Database").Range(myCol & rowStart + secStart + 10).Value = CType(dpReb.longitudinal_rebar_quantity, Double)
    '                            Else .Worksheets("Database").Range(myCol & rowStart + secStart + 10).ClearContents
    '                            End If
    '                            If Not IsNothing(dpReb.longitudinal_rebar_size) Then
    '                                .Worksheets("Database").Range(myCol & rowStart + secStart + 11).Value = CType(dpReb.longitudinal_rebar_size, Double)
    '                            Else .Worksheets("Database").Range(myCol & rowStart + secStart + 11).ClearContents
    '                            End If
    '                            If Not IsNothing(dpReb.longitudinal_rebar_cage_diameter) Then
    '                                .Worksheets("Database").Range(myCol & rowStart + secStart + 12).Value = CType(dpReb.longitudinal_rebar_cage_diameter, Double)
    '                            Else .Worksheets("Database").Range(myCol & rowStart + secStart + 12).ClearContents
    '                            End If
    '                        End If

    '                        rebCount += 1
    '                    Next

    '                    secCount += 1
    '                    secBump += 15
    '                    'secStart += secBump
    '                    secStart += 15

    '                Next

    '                'BELLED PIER
    '                If dp.belled_pier = True Then

    '                    If Not IsNothing(dp.belled_details.belled_pier_id) Then
    '                        .Worksheets("Database").Range(myCol & rowStart - 3).Value = CType(dp.belled_details.belled_pier_id, Integer)
    '                    Else .Worksheets("Database").Range(myCol & rowStart - 3).ClearContents
    '                    End If

    '                    .Worksheets("Database").Range(myCol & rowStart + 112).Value = CType(dp.belled_pier, Boolean)
    '                    If Not IsNothing(dp.belled_details.bottom_diameter_of_bell) Then
    '                        .Worksheets("Database").Range(myCol & rowStart + 113).Value = CType(dp.belled_details.bottom_diameter_of_bell, Double)
    '                    Else .Worksheets("Database").Range(myCol & rowStart + 113).ClearContents
    '                    End If
    '                    If Not IsNothing(dp.belled_details.bell_angle) Then
    '                        .Worksheets("Database").Range(myCol & rowStart + 114).Value = CType(dp.belled_details.bell_angle, Double)
    '                    Else .Worksheets("Database").Range(myCol & rowStart + 114).ClearContents
    '                    End If
    '                    .Worksheets("Database").Range(myCol & rowStart + 115).Value = CType(dp.belled_details.bell_input_type, String)
    '                    If Not IsNothing(dp.belled_details.bell_height) Then
    '                        .Worksheets("Database").Range(myCol & rowStart + 116).Value = CType(dp.belled_details.bell_height, Double)
    '                    Else .Worksheets("Database").Range(myCol & rowStart + 116).ClearContents
    '                    End If
    '                    If Not IsNothing(dp.belled_details.bell_toe_height) Then
    '                        .Worksheets("Database").Range(myCol & rowStart + 120).Value = CType(dp.belled_details.bell_toe_height, Double)
    '                    Else .Worksheets("Database").Range(myCol & rowStart + 120).ClearContents
    '                    End If
    '                    .Worksheets("Database").Range(myCol & rowStart + 122).Value = CType(dp.belled_details.neglect_top_soil_layer, Boolean)
    '                    .Worksheets("Database").Range(myCol & rowStart + 123).Value = CType(dp.belled_details.swelling_expansive_soil, Boolean)
    '                    If Not IsNothing(dp.belled_details.depth_of_expansive_soil) Then
    '                        .Worksheets("Database").Range(myCol & rowStart + 124).Value = CType(dp.belled_details.depth_of_expansive_soil, Double)
    '                    Else .Worksheets("Database").Range(myCol & rowStart + 124).ClearContents
    '                    End If
    '                    If Not IsNothing(dp.belled_details.expansive_soil_force) Then
    '                        .Worksheets("Database").Range(myCol & rowStart + 125).Value = CType(dp.belled_details.expansive_soil_force, Double)
    '                    Else .Worksheets("Database").Range(myCol & rowStart + 125).ClearContents
    '                    End If

    '                End If

    '                'EMBEDDED PIER
    '                If dp.embedded_pole = True Then

    '                    If Not IsNothing(dp.embed_details.embedded_id) Then
    '                        .Worksheets("Database").Range(myCol & rowStart - 2).Value = CType(dp.embed_details.embedded_id, Integer)
    '                    Else .Worksheets("Database").Range(myCol & rowStart - 2).ClearContents
    '                    End If

    '                    .Worksheets("Database").Range(myCol & rowStart + 100).Value = CType(dp.embedded_pole, Boolean)
    '                    .Worksheets("Database").Range(myCol & rowStart + 101).Value = CType(dp.embed_details.encased_in_concrete, Boolean)
    '                    If Not IsNothing(dp.embed_details.pole_side_quantity) Then
    '                        .Worksheets("Database").Range(myCol & rowStart + 102).Value = CType(dp.embed_details.pole_side_quantity, Integer)
    '                    Else .Worksheets("Database").Range(myCol & rowStart + 102).ClearContents
    '                    End If
    '                    If Not IsNothing(dp.embed_details.pole_yield_strength) Then
    '                        .Worksheets("Database").Range(myCol & rowStart + 103).Value = CType(dp.embed_details.pole_yield_strength, Double)
    '                    Else .Worksheets("Database").Range(myCol & rowStart + 103).ClearContents
    '                    End If
    '                    If Not IsNothing(dp.embed_details.pole_thickness) Then
    '                        .Worksheets("Database").Range(myCol & rowStart + 104).Value = CType(dp.embed_details.pole_thickness, Double)
    '                    Else .Worksheets("Database").Range(myCol & rowStart + 104).ClearContents
    '                    End If
    '                    .Worksheets("Database").Range(myCol & rowStart + 105).Value = CType(dp.embed_details.embedded_pole_input_type, String)
    '                    If Not IsNothing(dp.embed_details.pole_diameter_toc) Then
    '                        .Worksheets("Database").Range(myCol & rowStart + 106).Value = CType(dp.embed_details.pole_diameter_toc, Double)
    '                    Else .Worksheets("Database").Range(myCol & rowStart + 106).ClearContents
    '                    End If
    '                    If Not IsNothing(dp.embed_details.pole_top_diameter) Then
    '                        .Worksheets("Database").Range(myCol & rowStart + 107).Value = CType(dp.embed_details.pole_top_diameter, Double)
    '                    Else .Worksheets("Database").Range(myCol & rowStart + 107).ClearContents
    '                    End If
    '                    If Not IsNothing(dp.embed_details.pole_bottom_diameter) Then
    '                        .Worksheets("Database").Range(myCol & rowStart + 108).Value = CType(dp.embed_details.pole_bottom_diameter, Double)
    '                    Else .Worksheets("Database").Range(myCol & rowStart + 108).ClearContents
    '                    End If
    '                    If Not IsNothing(dp.embed_details.pole_section_length) Then
    '                        .Worksheets("Database").Range(myCol & rowStart + 109).Value = CType(dp.embed_details.pole_section_length, Double)
    '                    Else .Worksheets("Database").Range(myCol & rowStart + 109).ClearContents
    '                    End If
    '                    If Not IsNothing(dp.embed_details.pole_taper_factor) Then
    '                        .Worksheets("Database").Range(myCol & rowStart + 110).Value = CType(dp.embed_details.pole_taper_factor, Double)
    '                    Else .Worksheets("Database").Range(myCol & rowStart + 110).ClearContents
    '                    End If
    '                    If Not IsNothing(dp.embed_details.pole_bend_radius_override) Then
    '                        .Worksheets("Database").Range(myCol & rowStart + 111).Value = CType(dp.embed_details.pole_bend_radius_override, Double)
    '                    Else .Worksheets("Database").Range(myCol & rowStart + 111).ClearContents
    '                    End If

    '                End If

    '                'DRILLED PIER PROFILES
    '                Dim summaryRowStart As Integer = 10

    '                For Each dpp As DrilledPierProfile In dp.drilled_pier_profiles
    '                    'Profile Return
    '                    If Not IsNothing(dp.local_drilled_pier_id) Then
    '                        .Worksheets("Profiles (RETURN)").Range("A" & profileRow).Value = CType(dp.local_drilled_pier_id, Integer)
    '                    Else .Worksheets("Profiles (RETURN)").Range("A" & profileRow).ClearContents
    '                    End If
    '                    If Not IsNothing(dpp.reaction_position) Then
    '                        .Worksheets("Profiles (RETURN)").Range("B" & profileRow).Value = CType(dpp.reaction_position, Integer)
    '                    Else .Worksheets("Profiles (RETURN)").Range("B" & profileRow).ClearContents
    '                    End If
    '                    If Not IsNothing(dpp.drilled_pier_id) Then
    '                        .Worksheets("Profiles (RETURN)").Range("C" & profileRow).Value = CType(dpp.drilled_pier_id, Integer)
    '                    Else .Worksheets("Profiles (RETURN)").Range("C" & profileRow).ClearContents
    '                    End If
    '                    .Worksheets("Profiles (RETURN)").Range("D" & profileRow).Value = CType(dpp.profile_id, Integer)
    '                    If Not IsNothing(dpp.reaction_location) Then
    '                        .Worksheets("Profiles (RETURN)").Range("E" & profileRow).Value = CType(dpp.reaction_location, String)
    '                    Else .Worksheets("Profiles (RETURN)").Range("E" & profileRow).ClearContents
    '                    End If
    '                    If Not IsNothing(dpp.drilled_pier_profile) Then
    '                        .Worksheets("Profiles (RETURN)").Range("F" & profileRow).Value = CType(dpp.drilled_pier_profile, String)
    '                    Else .Worksheets("Profiles (RETURN)").Range("F" & profileRow).ClearContents
    '                    End If
    '                    If Not IsNothing(dpp.soil_profile) Then
    '                        .Worksheets("Profiles (RETURN)").Range("G" & profileRow).Value = CType(dpp.soil_profile, String)
    '                    Else .Worksheets("Profiles (RETURN)").Range("G" & profileRow).ClearContents
    '                    End If

    '                    'SUMMARY
    '                    If Not IsNothing(dpp.reaction_position) Then
    '                        .Worksheets("SUMMARY").Range("D" & summaryRowStart + CType(dpp.reaction_position, Integer)).Value = CType(dpp.drilled_pier_profile, Integer)
    '                        If dpp.drilled_pier_profile = dpp.reaction_position Then
    '                            .Worksheets("SUMMARY").Range("G" & summaryRowStart + CType(dpp.reaction_position, Integer)).Value = False
    '                        Else
    '                            .Worksheets("SUMMARY").Range("G" & summaryRowStart + CType(dpp.reaction_position, Integer)).Value = True
    '                        End If
    '                    End If
    '                    If Not IsNothing(dpp.reaction_position) Then
    '                        .Worksheets("SUMMARY").Range("E" & summaryRowStart + CType(dpp.reaction_position, Integer)).Value = CType(dpp.soil_profile, Integer)
    '                        If dpp.soil_profile = dpp.reaction_position Then
    '                            .Worksheets("SUMMARY").Range("H" & summaryRowStart + CType(dpp.reaction_position, Integer)).Value = False
    '                        Else
    '                            .Worksheets("SUMMARY").Range("H" & summaryRowStart + CType(dpp.reaction_position, Integer)).Value = True
    '                        End If
    '                    End If
    '                    .Worksheets("SUMMARY").Range("I" & summaryRowStart + CType(dpp.reaction_position, Integer)).Value = False
    '                    .Worksheets("SUMMARY").Range("J" & summaryRowStart + CType(dpp.reaction_position, Integer)).Value = CType(dpp.profile_id, Integer)

    '                    profileRow += 1

    '                Next

    '                .Worksheets("SUMMARY").Range("EDSReactions").Value = True

    '                'DRILLED PIER SOIL LAYER
    '                Dim soilCount As Integer
    '                Dim soilStart As Integer = 127
    '                Dim soilColCounter As Integer
    '                Dim mySoilCol As String

    '                For Each dpp As DrilledPierProfile In dp.drilled_pier_profiles

    '                    soilCount = 1

    '                    soilColCounter = 6 + dpp.soil_profile
    '                    mySoilCol = GetExcelColumnName(soilColCounter)

    '                    For Each dpSL As DrilledPierSoilLayer In dp.soil_layers

    '                        If Not IsNothing(dpSL.soil_layer_id) Then
    '                            .Worksheets("Database").Range(mySoilCol & rowStart - 33 + (soilCount - 1)).Value = CType(dpSL.soil_layer_id, Integer)
    '                        Else .Worksheets("Database").Range(mySoilCol & rowStart - 33 + (soilCount - 1)).ClearContents
    '                        End If

    '                        If Not IsNothing(dpSL.bottom_depth) Then
    '                            .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 0 + (soilCount - 1)).Value = CType(dpSL.bottom_depth, Double)
    '                        Else .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 0 + (soilCount - 1)).ClearContents
    '                        End If
    '                        If Not IsNothing(dpSL.effective_soil_density) Then
    '                            .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 1 + (soilCount - 1)).Value = CType(dpSL.effective_soil_density, Double)
    '                        Else .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 1 + (soilCount - 1)).ClearContents
    '                        End If
    '                        If Not IsNothing(dpSL.cohesion) Then
    '                            .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 2 + (soilCount - 1)).Value = CType(dpSL.cohesion, Double)
    '                        Else .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 2 + (soilCount - 1)).ClearContents
    '                        End If
    '                        If Not IsNothing(dpSL.friction_angle) Then
    '                            .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 3 + (soilCount - 1)).Value = CType(dpSL.friction_angle, Double)
    '                        Else .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 3 + (soilCount - 1)).ClearContents
    '                        End If
    '                        If Not IsNothing(dpSL.skin_friction_override_comp) Then
    '                            .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 4 + (soilCount - 1)).Value = CType(dpSL.skin_friction_override_comp, Double)
    '                        Else .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 4 + (soilCount - 1)).ClearContents
    '                        End If
    '                        If Not IsNothing(dpSL.skin_friction_override_uplift) Then
    '                            .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 5 + (soilCount - 1)).Value = CType(dpSL.skin_friction_override_uplift, Double)
    '                        Else .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 5 + (soilCount - 1)).ClearContents
    '                        End If
    '                        If Not IsNothing(dpSL.nominal_bearing_capacity) Then
    '                            .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 6 + (soilCount - 1)).Value = CType(dpSL.nominal_bearing_capacity, Double)
    '                        Else .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 6 + (soilCount - 1)).ClearContents
    '                        End If
    '                        If Not IsNothing(dpSL.spt_blow_count) Then
    '                            .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 7 + (soilCount - 1)).Value = CType(dpSL.spt_blow_count, Integer)
    '                        Else .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + 31 * 7 + (soilCount - 1)).ClearContents
    '                        End If

    '                        soilCount += 1
    '                    Next

    '                Next

    '                dpRow += 1
    '                colCounter += 1

    '            Next




    '            '~~~~~~~~POPULATE TOOL INPUTS WITH THE FIRST INSTANCE IN TOOL'S LOCAL DATABASE

    '            Dim firstReaction As String = DrilledPiers(0).drilled_pier_profiles(0).reaction_location

    '            If firstReaction = "Monopole" Then
    '                .Worksheets("Foundation Input").Range("TowerType").Value = "Monopole"
    '            ElseIf firstReaction = "Self Support" Then
    '                .Worksheets("Foundation Input").Range("TowerType").Value = "Self Support"
    '            ElseIf firstReaction = "Base" Then
    '                .Worksheets("Foundation Input").Range("TowerType").Value = "Guyed (Base)"
    '                .Worksheets("Foundation Input").Range("Location").Value = "Base"
    '            End If

    '            Dim firstPierProfile As Integer = DrilledPiers(0).drilled_pier_profiles(0).drilled_pier_profile
    '            Dim firstSoilProfile As Integer = DrilledPiers(0).drilled_pier_profiles(0).soil_profile

    '            colCounter = 7

    '            myCol = GetExcelColumnName(colCounter)

    '            If firstReaction <> "" Then

    '                'MATERIAL PROPERTIES
    '                If DrilledPiers(0).concrete_compressive_strength.HasValue Then
    '                    .Worksheets("Foundation Input").Range("f\c").Value = CType(DrilledPiers(0).concrete_compressive_strength, Double)
    '                Else .Worksheets("Foundation Input").Range("f\c").ClearContents
    '                End If
    '                If DrilledPiers(0).longitudinal_rebar_yield_strength.HasValue Then
    '                    .Worksheets("Foundation Input").Range("Fy_rebar").Value = CType(DrilledPiers(0).longitudinal_rebar_yield_strength, Double)
    '                Else .Worksheets("Foundation Input").Range("Fy_rebar").ClearContents
    '                End If
    '                If DrilledPiers(0).tie_yield_strength.HasValue Then
    '                    .Worksheets("Foundation Input").Range("yield_tie").Value = CType(DrilledPiers(0).tie_yield_strength, Double)
    '                Else .Worksheets("Foundation Input").Range("yield_tie").ClearContents
    '                End If
    '                If DrilledPiers(0).rebar_cage_2_fy_override.HasValue Then
    '                    .Worksheets("Foundation Input").Range("RebarCage2FyOverride").Value = CType(DrilledPiers(0).rebar_cage_2_fy_override, Double)
    '                Else .Worksheets("Foundation Input").Range("RebarCage2FyOverride").ClearContents
    '                End If
    '                If DrilledPiers(0).rebar_cage_3_fy_override.HasValue Then
    '                    .Worksheets("Foundation Input").Range("RebarCage3FyOverride").Value = CType(DrilledPiers(0).rebar_cage_3_fy_override, Double)
    '                Else .Worksheets("Foundation Input").Range("RebarCage3FyOverride").ClearContents
    '                End If

    '                'PIER DESIGN DATA (GENERAL)
    '                If DrilledPiers(0).foundation_depth.HasValue Then
    '                    .Worksheets("Foundation Input").Range("depth").Value = CType(DrilledPiers(0).foundation_depth, Double)
    '                Else .Worksheets("Foundation Input").Range("depth").ClearContents
    '                End If
    '                If DrilledPiers(0).extension_above_grade.HasValue Then
    '                    .Worksheets("Foundation Input").Range("ConcreteAboveGrade").Value = CType(DrilledPiers(0).extension_above_grade, Double)
    '                Else .Worksheets("Foundation Input").Range("ConcreteAboveGrade").ClearContents
    '                End If
    '                'groundwater
    '                If CType(DrilledPiers(0).groundwater_depth, Double) = -1 Then
    '                    .Worksheets("Foundation Input").Range("GW").Value = "N/A"
    '                Else .Worksheets("Foundation Input").Range("GW").Value = CType(DrilledPiers(0).groundwater_depth, Double)
    '                End If
    '                'soil layers
    '                If DrilledPiers(0).soil_layer_quantity.HasValue Then
    '                    .Worksheets("Foundation Input").Range("SoilLayerQty").Value = CType(DrilledPiers(0).soil_layer_quantity, Integer)
    '                Else .Worksheets("Foundation Input").Range("SoilLayerQty").ClearContents
    '                End If
    '                'min steel
    '                If Not IsNothing(CType(DrilledPiers(0).assume_min_steel, String)) Then
    '                    .Worksheets("Foundation Input").Range("AssumeMinSteel").Value = CType(DrilledPiers(0).assume_min_steel, String)
    '                Else .Worksheets("Foundation Input").Range("AssumeMinSteel").ClearContents
    '                End If

    '                'PIER DESIGN DATA (SECTIONS)
    '                Dim secCount As Integer = 1
    '                Dim secRowStart As Integer = 26

    '                For Each dpSec As DrilledPierSection In DrilledPiers(0).sections

    '                    If dpSec.pier_diameter.HasValue Then
    '                        .Worksheets("Foundation Input").Range("D" & secRowStart).Value = CType(dpSec.pier_diameter, Double)
    '                    Else .Worksheets("Foundation Input").Range("D" & secRowStart).ClearContents
    '                    End If
    '                    If dpSec.clear_cover.HasValue Then
    '                        .Worksheets("Foundation Input").Range("D" & secRowStart + 3).Value = CType(dpSec.clear_cover, Double)
    '                    Else .Worksheets("Foundation Input").Range("D" & secRowStart + 3).ClearContents
    '                    End If
    '                    If Not IsNothing(CType(dpSec.clear_cover_rebar_cage_option, String)) Then
    '                        .Worksheets("Foundation Input").Range("B" & secRowStart + 3).Value = CType(dpSec.clear_cover_rebar_cage_option, String)
    '                    Else .Worksheets("Foundation Input").Range("B" & secRowStart + 3).ClearContents
    '                    End If
    '                    If dpSec.tie_size.HasValue Then
    '                        .Worksheets("Foundation Input").Range("D" & secRowStart + 4).Value = CType(dpSec.tie_size, Integer)
    '                    Else .Worksheets("Foundation Input").Range("D" & secRowStart + 4).ClearContents
    '                    End If
    '                    If dpSec.tie_spacing.HasValue Then
    '                        .Worksheets("Foundation Input").Range("D" & secRowStart + 5).Value = CType(dpSec.tie_spacing, Double)
    '                    Else .Worksheets("Foundation Input").Range("D" & secRowStart + 5).ClearContents
    '                    End If
    '                    If dpSec.bottom_elevation.HasValue Then
    '                        .Worksheets("Foundation Input").Range("Depth" & secCount).Value = CType(dpSec.bottom_elevation, Double)
    '                    Else .Worksheets("Foundation Input").Range("Depth" & secCount).ClearContents
    '                    End If
    '                    If dpSec.rho_override.HasValue Then
    '                        .Worksheets("Foundation Input").Range("rhoOverride" & secCount).Value = CType(dpSec.rho_override, Double)
    '                    Else .Worksheets("Foundation Input").Range("rhoOverride" & secCount).ClearContents
    '                    End If

    '                    'PIER DESIGN DATA (REBAR)
    '                    Dim rebCount As Integer = 1

    '                    For Each dpReb As DrilledPierRebar In dpSec.rebar

    '                        If rebCount = 1 Then
    '                            If dpReb.longitudinal_rebar_quantity.HasValue Then
    '                                .Worksheets("Foundation Input").Range("D" & secRowStart + 1).Value = CType(dpReb.longitudinal_rebar_quantity, Integer)
    '                            Else .Worksheets("Foundation Input").Range("D" & secRowStart + 1).ClearContents
    '                            End If
    '                            If dpReb.longitudinal_rebar_size.HasValue Then
    '                                .Worksheets("Foundation Input").Range("D" & secRowStart + 2).Value = CType(dpReb.longitudinal_rebar_size, Integer)
    '                            Else .Worksheets("Foundation Input").Range("D" & secRowStart + 2).ClearContents
    '                            End If
    '                        End If

    '                        If rebCount = 2 Then
    '                            If dpReb.longitudinal_rebar_quantity.HasValue Then
    '                                .Worksheets("Foundation Input").Range("D" & secRowStart + 6).Value = CType(dpReb.longitudinal_rebar_quantity, Integer)
    '                            Else .Worksheets("Foundation Input").Range("D" & secRowStart + 6).ClearContents
    '                            End If
    '                            If dpReb.longitudinal_rebar_size.HasValue Then
    '                                .Worksheets("Foundation Input").Range("D" & secRowStart + 7).Value = CType(dpReb.longitudinal_rebar_size, Integer)
    '                            Else .Worksheets("Foundation Input").Range("D" & secRowStart + 7).ClearContents
    '                            End If
    '                            If dpReb.longitudinal_rebar_cage_diameter.HasValue Then
    '                                .Worksheets("Foundation Input").Range("D" & secRowStart + 8).Value = CType(dpReb.longitudinal_rebar_cage_diameter, Integer)
    '                            Else .Worksheets("Foundation Input").Range("D" & secRowStart + 8).ClearContents
    '                            End If
    '                        End If

    '                        If rebCount = 3 Then
    '                            If dpReb.longitudinal_rebar_quantity.HasValue Then
    '                                .Worksheets("Foundation Input").Range("D" & secRowStart + 10).Value = CType(dpReb.longitudinal_rebar_quantity, Integer)
    '                            Else .Worksheets("Foundation Input").Range("D" & secRowStart + 10).ClearContents
    '                            End If
    '                            If dpReb.longitudinal_rebar_size.HasValue Then
    '                                .Worksheets("Foundation Input").Range("D" & secRowStart + 11).Value = CType(dpReb.longitudinal_rebar_size, Integer)
    '                            Else .Worksheets("Foundation Input").Range("D" & secRowStart + 11).ClearContents
    '                            End If
    '                            If dpReb.longitudinal_rebar_cage_diameter.HasValue Then
    '                                .Worksheets("Foundation Input").Range("D" & secRowStart + 12).Value = CType(dpReb.longitudinal_rebar_cage_diameter, Integer)
    '                            Else .Worksheets("Foundation Input").Range("D" & secRowStart + 12).ClearContents
    '                            End If
    '                        End If

    '                        rebCount += 1

    '                    Next

    '                    'populate rebar cage qty (hidden input in tool, typically populated by the Pier Options)
    '                    .Worksheets("Foundation Input").Range("Rebar" & secCount).Value = rebCount - 1

    '                    secCount += 1

    '                    secRowStart += 16

    '                Next


    '                'SOIL
    '                Dim soilRowStart As Integer = 121
    '                Dim soilCount As Integer = 1

    '                For Each dpSL As DrilledPierSoilLayer In DrilledPiers(0).soil_layers

    '                    If dpSL.bottom_depth.HasValue Then
    '                        .Worksheets("Foundation Input").Range("D" & soilRowStart + soilCount).Value = CType(dpSL.bottom_depth, Double)
    '                    Else .Worksheets("Foundation Input").Range("D" & soilRowStart + soilCount).ClearContents
    '                    End If
    '                    If dpSL.effective_soil_density.HasValue Then
    '                        .Worksheets("Foundation Input").Range("F" & soilRowStart + soilCount).Value = CType(dpSL.effective_soil_density, Double)
    '                    Else .Worksheets("Foundation Input").Range("F" & soilRowStart + soilCount).ClearContents
    '                    End If
    '                    If dpSL.cohesion.HasValue Then
    '                        .Worksheets("Foundation Input").Range("H" & soilRowStart + soilCount).Value = CType(dpSL.cohesion, Double)
    '                    Else .Worksheets("Foundation Input").Range("H" & soilRowStart + soilCount).ClearContents
    '                    End If
    '                    If dpSL.friction_angle.HasValue Then
    '                        .Worksheets("Foundation Input").Range("I" & soilRowStart + soilCount).Value = CType(dpSL.friction_angle, Double)
    '                    Else .Worksheets("Foundation Input").Range("I" & soilRowStart + soilCount).ClearContents
    '                    End If
    '                    If dpSL.skin_friction_override_comp.HasValue Then
    '                        .Worksheets("Foundation Input").Range("M" & soilRowStart + soilCount).Value = CType(dpSL.skin_friction_override_comp, Double)
    '                    Else .Worksheets("Foundation Input").Range("M" & soilRowStart + soilCount).ClearContents
    '                    End If
    '                    If dpSL.skin_friction_override_uplift.HasValue Then
    '                        .Worksheets("Foundation Input").Range("N" & soilRowStart + soilCount).Value = CType(dpSL.skin_friction_override_uplift, Double)
    '                    Else .Worksheets("Foundation Input").Range("N" & soilRowStart + soilCount).ClearContents
    '                    End If
    '                    If dpSL.nominal_bearing_capacity.HasValue Then
    '                        .Worksheets("Foundation Input").Range("O" & soilRowStart + soilCount).Value = CType(dpSL.nominal_bearing_capacity, Double)
    '                    Else .Worksheets("Foundation Input").Range("O" & soilRowStart + soilCount).ClearContents
    '                    End If
    '                    If dpSL.spt_blow_count.HasValue Then
    '                        .Worksheets("Foundation Input").Range("P" & soilRowStart + soilCount).Value = CType(dpSL.spt_blow_count, Integer)
    '                    Else .Worksheets("Foundation Input").Range("P" & soilRowStart + soilCount).ClearContents
    '                    End If

    '                    soilCount += 1

    '                Next


    '                'OPTIONS
    '                .Worksheets("Foundation Input").Range("EffectiveDepthInput").Value = CType(DrilledPiers(0).rebar_effective_depths, Boolean)
    '                .Worksheets("Foundation Input").Range("ShearAlongDepth").Value = CType(DrilledPiers(0).check_shear_along_depth, Boolean)
    '                .Worksheets("Foundation Input").Range("ShearFriction").Value = CType(DrilledPiers(0).utilize_shear_friction_methodology, Boolean)
    '                .Worksheets("Foundation Input").Range("ShearInputOverride").Value = CType(DrilledPiers(0).shear_override_crit_depth, Boolean)
    '                If .Worksheets("Foundation Input").Range("ShearInputOverride").Value = CType(DrilledPiers(0).shear_override_crit_depth, Boolean) = True Then
    '                    If DrilledPiers(0).shear_crit_depth_override_comp.HasValue Then
    '                        .Worksheets("Foundation Input").Range("ShearCritDepthComp").Value = CType(DrilledPiers(0).shear_crit_depth_override_comp, Double)
    '                    End If
    '                    If DrilledPiers(0).shear_crit_depth_override_uplift.HasValue Then
    '                        .Worksheets("Foundation Input").Range("ShearCritDepthUplift").Value = CType(DrilledPiers(0).shear_crit_depth_override_uplift, Double)
    '                    End If
    '                End If


    '                'BELLED PIER
    '                .Worksheets("Belled Pier").Range("Belled").Value = CType(DrilledPiers(0).belled_pier, Boolean)
    '                If DrilledPiers(0).belled_pier = True Then
    '                    If DrilledPiers(0).belled_details.bottom_diameter_of_bell.HasValue Then
    '                        .Worksheets("Belled Pier").Range("Dia_Bell").Value = CType(DrilledPiers(0).belled_details.bottom_diameter_of_bell, Double)
    '                    Else .Worksheets("Belled Pier").Range("Dia_Bell").ClearContents
    '                    End If
    '                    If Not IsNothing(CType(DrilledPiers(0).belled_details.bell_input_type, String)) Then
    '                        .Worksheets("Belled Pier").Range("BellInputType").Value = CType(DrilledPiers(0).belled_details.bell_input_type, String)
    '                    Else .Worksheets("Belled Pier").Range("BellInputType").ClearContents
    '                    End If
    '                    If DrilledPiers(0).belled_details.bell_angle.HasValue Then
    '                        .Worksheets("Belled Pier").Range("BellAngle").Value = CType(DrilledPiers(0).belled_details.bell_angle, Double)
    '                    Else .Worksheets("Belled Pier").Range("BellAngle").ClearContents
    '                    End If
    '                    If DrilledPiers(0).belled_details.bell_height.HasValue Then
    '                        .Worksheets("Belled Pier").Range("hbell").Value = CType(DrilledPiers(0).belled_details.bell_height, Double)
    '                    Else .Worksheets("Belled Pier").Range("hbell").ClearContents
    '                    End If
    '                    If DrilledPiers(0).belled_details.bell_toe_height.HasValue Then
    '                        .Worksheets("Belled Pier").Range("t_bell").Value = CType(DrilledPiers(0).belled_details.bell_toe_height, Double)
    '                    Else .Worksheets("Belled Pier").Range("t_bell").ClearContents
    '                    End If
    '                    .Worksheets("Belled Pier").Range("Neglect_Top").Value = CType(DrilledPiers(0).belled_details.neglect_top_soil_layer, Boolean)
    '                    .Worksheets("Belled Pier").Range("expansive").Value = CType(DrilledPiers(0).belled_details.expansive_soil_force, Boolean)
    '                    If DrilledPiers(0).belled_details.depth_of_expansive_soil.HasValue Then
    '                        .Worksheets("Belled Pier").Range("D_expansive").Value = CType(DrilledPiers(0).belled_details.depth_of_expansive_soil, Double)
    '                    Else .Worksheets("Belled Pier").Range("D_expansive").ClearContents
    '                    End If
    '                    If DrilledPiers(0).belled_details.expansive_soil_force.HasValue Then
    '                        .Worksheets("Belled Pier").Range("Force_Expansive").Value = CType(DrilledPiers(0).belled_details.expansive_soil_force, Double)
    '                    Else .Worksheets("Belled Pier").Range("Force_Expansive").ClearContents
    '                    End If
    '                End If


    '                'EMBEDDED POLE
    '                .Worksheets("Soil Calculations").Range("Embedded").Value = CType(DrilledPiers(0).embedded_pole, Boolean)
    '                If DrilledPiers(0).embedded_pole = True Then
    '                    .Worksheets("Soil Calculations").Range("Encased").Value = CType(DrilledPiers(0).embed_details.encased_in_concrete, Boolean)
    '                    If DrilledPiers(0).embed_details.pole_side_quantity.HasValue Then
    '                        .Worksheets("Soil Calculations").Range("Sides").Value = CType(DrilledPiers(0).embed_details.pole_side_quantity, Integer)
    '                    Else .Worksheets("Soil Calculations").Range("Sides").ClearContents
    '                    End If
    '                    If DrilledPiers(0).embed_details.pole_yield_strength.HasValue Then
    '                        .Worksheets("Soil Calculations").Range("Fy").Value = CType(DrilledPiers(0).embed_details.pole_yield_strength, Double)
    '                    Else .Worksheets("Soil Calculations").Range("Fy").ClearContents
    '                    End If
    '                    If DrilledPiers(0).embed_details.pole_thickness.HasValue Then
    '                        .Worksheets("Soil Calculations").Range("t").Value = CType(DrilledPiers(0).embed_details.pole_thickness, Double)
    '                    Else .Worksheets("Soil Calculations").Range("t").ClearContents
    '                    End If
    '                    If Not IsNothing(CType(DrilledPiers(0).embed_details.embedded_pole_input_type, String)) Then
    '                        .Worksheets("Soil Calculations").Range("EmbeddedPoleInputType").Value = CType(DrilledPiers(0).embed_details.embedded_pole_input_type, String)
    '                    Else .Worksheets("Soil Calculations").Range("EmbeddedPoleInputType").ClearContents
    '                    End If
    '                    If DrilledPiers(0).embed_details.pole_diameter_toc.HasValue Then
    '                        .Worksheets("Soil Calculations").Range("dia_grade").Value = CType(DrilledPiers(0).embed_details.pole_diameter_toc, Double)
    '                    Else .Worksheets("Soil Calculations").Range("dia_grade").ClearContents
    '                    End If
    '                    If DrilledPiers(0).embed_details.pole_top_diameter.HasValue Then
    '                        .Worksheets("Soil Calculations").Range("TopDiameter").Value = CType(DrilledPiers(0).embed_details.pole_top_diameter, Double)
    '                    Else .Worksheets("Soil Calculations").Range("TopDiameter").ClearContents
    '                    End If
    '                    If DrilledPiers(0).embed_details.pole_bottom_diameter.HasValue Then
    '                        .Worksheets("Soil Calculations").Range("BottomDiameter").Value = CType(DrilledPiers(0).embed_details.pole_bottom_diameter, Double)
    '                    Else .Worksheets("Soil Calculations").Range("BottomDiameter").ClearContents
    '                    End If
    '                    If DrilledPiers(0).embed_details.pole_section_length.HasValue Then
    '                        .Worksheets("Soil Calculations").Range("LengthOfSection").Value = CType(DrilledPiers(0).embed_details.pole_section_length, Double)
    '                    Else .Worksheets("Soil Calculations").Range("LengthOfSection").ClearContents
    '                    End If
    '                    If DrilledPiers(0).embed_details.pole_taper_factor.HasValue Then
    '                        .Worksheets("Soil Calculations").Range("taper").Value = CType(DrilledPiers(0).embed_details.pole_taper_factor, Double)
    '                    Else .Worksheets("Soil Calculations").Range("taper").ClearContents
    '                    End If
    '                    If DrilledPiers(0).embed_details.pole_bend_radius_override.HasValue Then
    '                        .Worksheets("Soil Calculations").Range("bend_user").Value = CType(DrilledPiers(0).embed_details.pole_bend_radius_override, Double)
    '                    Else .Worksheets("Soil Calculations").Range("bend_user").ClearContents
    '                    End If
    '                End If

    '            End If


    '        End With


    '        SaveAndCloseDrilledPier()
    '    End Sub

    '    Private Function GetExcelColumnName(columnNumber As Integer) As String
    '        Dim dividend As Integer = columnNumber
    '        Dim columnName As String = String.Empty
    '        Dim modulo As Integer

    '        While dividend > 0
    '            modulo = (dividend - 1) Mod 26
    '            columnName = Convert.ToChar(65 + modulo).ToString() & columnName
    '            dividend = CInt((dividend - modulo) / 26)
    '        End While

    '        Return columnName
    '    End Function

    '    Private Sub LoadNewDrilledPier()
    '        NewDrilledPierWb.LoadDocument(DrilledPierTemplatePath, DrilledPierFileType)
    '        NewDrilledPierWb.BeginUpdate()
    '    End Sub

    '    Private Sub SaveAndCloseDrilledPier()
    '        NewDrilledPierWb.EndUpdate()
    '        NewDrilledPierWb.SaveDocument(ExcelFilePath, DrilledPierFileType)
    '    End Sub
    '#End Region

    '#Region "SQL Insert Statements"
    '    Private Function InsertPoleCriteria(ByVal pc As PoleCriteria) As String
    '        Dim insertString As String = ""

    '        insertString += "@PoleCriteriaID"
    '        insertString += "," & IIf(IsNothing(pc.criteria_id), "Null", pc.criteria_id.ToString)
    '        insertString += "," & IIf(IsNothing(pc.upper_structure_type), "Null", "'" & pc.upper_structure_type.ToString & "'")
    '        insertString += "," & IIf(IsNothing(pc.analysis_deg), "Null", pc.analysis_deg.ToString)
    '        insertString += "," & IIf(IsNothing(pc.geom_increment_length), "Null", pc.geom_increment_length.ToString)
    '        insertString += "," & IIf(IsNothing(pc.vnum), "Null", "'" & pc.vnum.ToString & "'")
    '        insertString += "," & IIf(IsNothing(pc.check_connections), "Null", "'" & pc.check_connections.ToString & "'")
    '        insertString += "," & IIf(IsNothing(pc.hole_deformation), "Null", "'" & pc.hole_deformation.ToString & "'")
    '        insertString += "," & IIf(IsNothing(pc.ineff_mod_check), "Null", "'" & pc.ineff_mod_check.ToString & "'")
    '        insertString += "," & IIf(IsNothing(pc.modified), "Null", "'" & pc.modified.ToString & "'")

    '        Return insertString
    '    End Function

    '    Private Function InsertPoleSection(ByVal ps As PoleSection) As String
    '        Dim insertString As String = ""

    '        insertString += "@PoleSectionID"
    '        insertString += "," & IIf(IsNothing(ps.section_id), "Null", ps.section_id.ToString)
    '        insertString += "," & IIf(IsNothing(ps.analysis_section_id), "Null", ps.analysis_section_id.ToString)
    '        insertString += "," & IIf(IsNothing(ps.elev_bot), "Null", ps.elev_bot.ToString)
    '        insertString += "," & IIf(IsNothing(ps.elev_top), "Null", ps.elev_top.ToString)
    '        insertString += "," & IIf(IsNothing(ps.length_section), "Null", ps.length_section.ToString)
    '        insertString += "," & IIf(IsNothing(ps.length_splice), "Null", ps.length_splice.ToString)
    '        insertString += "," & IIf(IsNothing(ps.num_sides), "Null", ps.num_sides.ToString)
    '        insertString += "," & IIf(IsNothing(ps.diam_bot), "Null", ps.diam_bot.ToString)
    '        insertString += "," & IIf(IsNothing(ps.diam_top), "Null", ps.diam_top.ToString)
    '        insertString += "," & IIf(IsNothing(ps.wall_thickness), "Null", ps.wall_thickness.ToString)
    '        insertString += "," & IIf(IsNothing(ps.bend_radius), "Null", ps.bend_radius.ToString)
    '        insertString += "," & IIf(IsNothing(ps.steel_grade_id), "Null", ps.steel_grade_id.ToString)
    '        insertString += "," & IIf(IsNothing(ps.pole_type), "Null", "'" & ps.pole_type.ToString & "'")
    '        insertString += "," & IIf(IsNothing(ps.section_name), "Null", "'" & ps.section_name.ToString & "'")
    '        insertString += "," & IIf(IsNothing(ps.socket_length), "Null", ps.socket_length.ToString)
    '        insertString += "," & IIf(IsNothing(ps.weight_mult), "Null", ps.weight_mult.ToString)
    '        insertString += "," & IIf(IsNothing(ps.wp_mult), "Null", ps.wp_mult.ToString)
    '        insertString += "," & IIf(IsNothing(ps.af_factor), "Null", ps.af_factor.ToString)
    '        insertString += "," & IIf(IsNothing(ps.ar_factor), "Null", ps.ar_factor.ToString)
    '        insertString += "," & IIf(IsNothing(ps.round_area_ratio), "Null", ps.round_area_ratio.ToString)
    '        insertString += "," & IIf(IsNothing(ps.flat_area_ratio), "Null", ps.flat_area_ratio.ToString)

    '        Return insertString
    '    End Function

    '    Private Function InsertPoleReinfSection(ByVal prs As PoleReinfSection) As String
    '        Dim insertString As String = ""

    '        insertString += "@PoleReinfSectionID"
    '        insertString += "," & IIf(IsNothing(prs.section_ID), "Null", prs.section_ID.ToString)
    '        insertString += "," & IIf(IsNothing(prs.analysis_section_ID), "Null", prs.analysis_section_ID.ToString)
    '        insertString += "," & IIf(IsNothing(prs.elev_bot), "Null", prs.elev_bot.ToString)
    '        insertString += "," & IIf(IsNothing(prs.elev_top), "Null", prs.elev_top.ToString)
    '        insertString += "," & IIf(IsNothing(prs.length_section), "Null", prs.length_section.ToString)
    '        insertString += "," & IIf(IsNothing(prs.length_splice), "Null", prs.length_splice.ToString)
    '        insertString += "," & IIf(IsNothing(prs.num_sides), "Null", prs.num_sides.ToString)
    '        insertString += "," & IIf(IsNothing(prs.diam_bot), "Null", prs.diam_bot.ToString)
    '        insertString += "," & IIf(IsNothing(prs.diam_top), "Null", prs.diam_top.ToString)
    '        insertString += "," & IIf(IsNothing(prs.wall_thickness), "Null", prs.wall_thickness.ToString)
    '        insertString += "," & IIf(IsNothing(prs.bend_radius), "Null", prs.bend_radius.ToString)
    '        insertString += "," & IIf(IsNothing(prs.steel_grade_id), "Null", prs.steel_grade_id.ToString)
    '        insertString += "," & IIf(IsNothing(prs.pole_type), "Null", "'" & prs.pole_type.ToString & "'")
    '        insertString += "," & IIf(IsNothing(prs.weight_mult), "Null", prs.weight_mult.ToString)
    '        insertString += "," & IIf(IsNothing(prs.section_name), "Null", "'" & prs.section_name.ToString & "'")
    '        insertString += "," & IIf(IsNothing(prs.socket_length), "Null", prs.socket_length.ToString)
    '        insertString += "," & IIf(IsNothing(prs.wp_mult), "Null", prs.wp_mult.ToString)
    '        insertString += "," & IIf(IsNothing(prs.af_factor), "Null", prs.af_factor.ToString)
    '        insertString += "," & IIf(IsNothing(prs.ar_factor), "Null", prs.ar_factor.ToString)
    '        insertString += "," & IIf(IsNothing(prs.round_area_ratio), "Null", prs.round_area_ratio.ToString)
    '        insertString += "," & IIf(IsNothing(prs.flat_area_ratio), "Null", prs.flat_area_ratio.ToString)

    '        Return insertString
    '    End Function

    '    Private Function InsertPoleReinfGroup(ByVal prg As PoleReinfGroup) As String
    '        Dim insertString As String = ""

    '        insertString += "@PoleReinfGroupID"
    '        insertString += "," & IIf(IsNothing(prg.reinf_group_id), "Null", prg.reinf_group_id.ToString)
    '        insertString += "," & IIf(IsNothing(prg.elev_bot_actual), "Null", prg.elev_bot_actual.ToString)
    '        insertString += "," & IIf(IsNothing(prg.elev_bot_eff), "Null", prg.elev_bot_eff.ToString)
    '        insertString += "," & IIf(IsNothing(prg.elev_top_actual), "Null", prg.elev_top_actual.ToString)
    '        insertString += "," & IIf(IsNothing(prg.elev_top_eff), "Null", prg.elev_top_eff.ToString)
    '        insertString += "," & IIf(IsNothing(prg.reinf_db_id), "Null", prg.reinf_db_id.ToString)

    '        Return insertString
    '    End Function

    '    Private Function InsertPoleReinfDetail(ByVal prd As PoleReinfDetail) As String
    '        Dim insertString As String = ""

    '        insertString += "@PoleReinfDetailID"
    '        insertString += "," & IIf(IsNothing(prd.reinf_id), "Null", prd.reinf_id.ToString)
    '        insertString += "," & IIf(IsNothing(prd.pole_flat), "Null", prd.pole_flat.ToString)
    '        insertString += "," & IIf(IsNothing(prd.horizontal_offset), "Null", prd.horizontal_offset.ToString)
    '        insertString += "," & IIf(IsNothing(prd.rotation), "Null", prd.rotation.ToString)
    '        insertString += "," & IIf(IsNothing(prd.note), "Null", "'" & prd.note.ToString & "'")

    '        Return insertString
    '    End Function

    '    Private Function InsertPoleIntGroup(ByVal pig As PoleIntGroup) As String
    '        Dim insertString As String = ""

    '        insertString += "@PoleIntGroupID"
    '        insertString += "," & IIf(IsNothing(pig.interference_group_id), "Null", pig.interference_group_id.ToString)
    '        insertString += "," & IIf(IsNothing(pig.elev_bot), "Null", pig.elev_bot.ToString)
    '        insertString += "," & IIf(IsNothing(pig.elev_top), "Null", pig.elev_top.ToString)
    '        insertString += "," & IIf(IsNothing(pig.width), "Null", pig.width.ToString)
    '        insertString += "," & IIf(IsNothing(pig.description), "Null", "'" & pig.description.ToString & "'")

    '        Return insertString
    '    End Function

    '    Private Function InsertPoleIntDetail(ByVal pid As PoleIntDetail) As String
    '        Dim insertString As String = ""

    '        insertString += "@IntDetailID"
    '        insertString += "," & IIf(IsNothing(pid.interference_id), "Null", pid.interference_id.ToString)
    '        insertString += "," & IIf(IsNothing(pid.pole_flat), "Null", pid.pole_flat.ToString)
    '        insertString += "," & IIf(IsNothing(pid.horizontal_offset), "Null", pid.horizontal_offset.ToString)
    '        insertString += "," & IIf(IsNothing(pid.rotation), "Null", pid.rotation.ToString)
    '        insertString += "," & IIf(IsNothing(pid.note), "Null", "'" & pid.note.ToString & "'")

    '        Return insertString
    '    End Function

    '    Private Function InsertPoleReinfResults(ByVal prr As PoleReinfResults) As String
    '        Dim insertString As String = ""

    '        insertString += "@PoleReinfResultID"
    '        insertString += "," & IIf(IsNothing(prr.section_id), "Null", prr.section_id.ToString)
    '        insertString += "," & IIf(IsNothing(prr.work_order_seq_num), "Null", prr.work_order_seq_num.ToString)
    '        insertString += "," & IIf(IsNothing(prr.reinf_group_id), "Null", prr.reinf_group_id.ToString)
    '        insertString += "," & IIf(IsNothing(prr.result_lkup_value), "Null", prr.result_lkup_value.ToString)
    '        insertString += "," & IIf(IsNothing(prr.rating), "Null", prr.rating.ToString)

    '        Return insertString
    '    End Function

    '    Private Function InsertPropReinf(ByVal pr As PropReinf) As String
    '        Dim insertString As String = ""

    '        insertString += "@ReinfID"
    '        insertString += "," & IIf(IsNothing(pr.reinf_db_id), "Null", pr.reinf_db_id.ToString)
    '        insertString += "," & IIf(IsNothing(pr.name), "Null", "'" & pr.name.ToString & "'")
    '        insertString += "," & IIf(IsNothing(pr.type), "Null", "'" & pr.type.ToString & "'")
    '        insertString += "," & IIf(IsNothing(pr.b), "Null", pr.b.ToString)
    '        insertString += "," & IIf(IsNothing(pr.h), "Null", pr.h.ToString)
    '        insertString += "," & IIf(IsNothing(pr.sr_diam), "Null", pr.sr_diam.ToString)
    '        insertString += "," & IIf(IsNothing(pr.channel_thkns_web), "Null", pr.channel_thkns_web.ToString)
    '        insertString += "," & IIf(IsNothing(pr.channel_thkns_flange), "Null", pr.channel_thkns_flange.ToString)
    '        insertString += "," & IIf(IsNothing(pr.channel_eo), "Null", pr.channel_eo.ToString)
    '        insertString += "," & IIf(IsNothing(pr.channel_J), "Null", pr.channel_J.ToString)
    '        insertString += "," & IIf(IsNothing(pr.channel_Cw), "Null", pr.channel_Cw.ToString)
    '        insertString += "," & IIf(IsNothing(pr.area_gross), "Null", pr.area_gross.ToString)
    '        insertString += "," & IIf(IsNothing(pr.centroid), "Null", pr.centroid.ToString)
    '        insertString += "," & IIf(IsNothing(pr.istension), "Null", "'" & pr.istension.ToString & "'")
    '        insertString += "," & IIf(IsNothing(pr.matl_id), "Null", pr.matl_id.ToString)
    '        insertString += "," & IIf(IsNothing(pr.Ix), "Null", pr.Ix.ToString)
    '        insertString += "," & IIf(IsNothing(pr.Iy), "Null", pr.Iy.ToString)
    '        insertString += "," & IIf(IsNothing(pr.Lu), "Null", pr.Lu.ToString)
    '        insertString += "," & IIf(IsNothing(pr.Kx), "Null", pr.Kx.ToString)
    '        insertString += "," & IIf(IsNothing(pr.Ky), "Null", pr.Ky.ToString)
    '        insertString += "," & IIf(IsNothing(pr.bolt_hole_size), "Null", pr.bolt_hole_size.ToString)
    '        insertString += "," & IIf(IsNothing(pr.area_net), "Null", pr.area_net.ToString)
    '        insertString += "," & IIf(IsNothing(pr.shear_lag), "Null", pr.shear_lag.ToString)
    '        insertString += "," & IIf(IsNothing(pr.connection_type_bot), "Null", "'" & pr.connection_type_bot.ToString & "'")
    '        insertString += "," & IIf(IsNothing(pr.connection_cap_revF_bot), "Null", pr.connection_cap_revF_bot.ToString)
    '        insertString += "," & IIf(IsNothing(pr.connection_cap_revG_bot), "Null", pr.connection_cap_revG_bot.ToString)
    '        insertString += "," & IIf(IsNothing(pr.connection_cap_revH_bot), "Null", pr.connection_cap_revH_bot.ToString)
    '        insertString += "," & IIf(IsNothing(pr.bolt_type_id_bot), "Null", pr.bolt_type_id_bot.ToString)
    '        insertString += "," & IIf(IsNothing(pr.bolt_N_or_X_bot), "Null", "'" & pr.bolt_N_or_X_bot.ToString & "'")
    '        insertString += "," & IIf(IsNothing(pr.bolt_num_bot), "Null", pr.bolt_num_bot.ToString)
    '        insertString += "," & IIf(IsNothing(pr.bolt_spacing_bot), "Null", pr.bolt_spacing_bot.ToString)
    '        insertString += "," & IIf(IsNothing(pr.bolt_edge_dist_bot), "Null", pr.bolt_edge_dist_bot.ToString)
    '        insertString += "," & IIf(IsNothing(pr.FlangeOrBP_connected_bot), "Null", "'" & pr.FlangeOrBP_connected_bot.ToString & "'")
    '        insertString += "," & IIf(IsNothing(pr.weld_grade_bot), "Null", pr.weld_grade_bot.ToString)
    '        insertString += "," & IIf(IsNothing(pr.weld_trans_type_bot), "Null", "'" & pr.weld_trans_type_bot.ToString & "'")
    '        insertString += "," & IIf(IsNothing(pr.weld_trans_length_bot), "Null", pr.weld_trans_length_bot.ToString)
    '        insertString += "," & IIf(IsNothing(pr.weld_groove_depth_bot), "Null", pr.weld_groove_depth_bot.ToString)
    '        insertString += "," & IIf(IsNothing(pr.weld_groove_angle_bot), "Null", pr.weld_groove_angle_bot.ToString)
    '        insertString += "," & IIf(IsNothing(pr.weld_trans_fillet_size_bot), "Null", pr.weld_trans_fillet_size_bot.ToString)
    '        insertString += "," & IIf(IsNothing(pr.weld_trans_eff_throat_bot), "Null", pr.weld_trans_eff_throat_bot.ToString)
    '        insertString += "," & IIf(IsNothing(pr.weld_long_type_bot), "Null", "'" & pr.weld_long_type_bot.ToString & "'")
    '        insertString += "," & IIf(IsNothing(pr.weld_long_length_bot), "Null", pr.weld_long_length_bot.ToString)
    '        insertString += "," & IIf(IsNothing(pr.weld_long_fillet_size_bot), "Null", pr.weld_long_fillet_size_bot.ToString)
    '        insertString += "," & IIf(IsNothing(pr.weld_long_eff_throat_bot), "Null", pr.weld_long_eff_throat_bot.ToString)
    '        insertString += "," & IIf(IsNothing(pr.top_bot_connections_symmetrical), "Null", "'" & pr.top_bot_connections_symmetrical.ToString & "'")
    '        insertString += "," & IIf(IsNothing(pr.connection_type_top), "Null", "'" & pr.connection_type_top.ToString & "'")
    '        insertString += "," & IIf(IsNothing(pr.connection_cap_revF_top), "Null", pr.connection_cap_revF_top.ToString)
    '        insertString += "," & IIf(IsNothing(pr.connection_cap_revG_top), "Null", pr.connection_cap_revG_top.ToString)
    '        insertString += "," & IIf(IsNothing(pr.connection_cap_revH_top), "Null", pr.connection_cap_revH_top.ToString)
    '        insertString += "," & IIf(IsNothing(pr.bolt_type_id_top), "Null", pr.bolt_type_id_top.ToString)
    '        insertString += "," & IIf(IsNothing(pr.bolt_N_or_X_top), "Null", "'" & pr.bolt_N_or_X_top.ToString & "'")
    '        insertString += "," & IIf(IsNothing(pr.bolt_num_top), "Null", pr.bolt_num_top.ToString)
    '        insertString += "," & IIf(IsNothing(pr.bolt_spacing_top), "Null", pr.bolt_spacing_top.ToString)
    '        insertString += "," & IIf(IsNothing(pr.bolt_edge_dist_top), "Null", pr.bolt_edge_dist_top.ToString)
    '        insertString += "," & IIf(IsNothing(pr.FlangeOrBP_connected_top), "Null", "'" & pr.FlangeOrBP_connected_top.ToString & "'")
    '        insertString += "," & IIf(IsNothing(pr.weld_grade_top), "Null", pr.weld_grade_top.ToString)
    '        insertString += "," & IIf(IsNothing(pr.weld_trans_type_top), "Null", "'" & pr.weld_trans_type_top.ToString & "'")
    '        insertString += "," & IIf(IsNothing(pr.weld_trans_length_top), "Null", pr.weld_trans_length_top.ToString)
    '        insertString += "," & IIf(IsNothing(pr.weld_groove_depth_top), "Null", pr.weld_groove_depth_top.ToString)
    '        insertString += "," & IIf(IsNothing(pr.weld_groove_angle_top), "Null", pr.weld_groove_angle_top.ToString)
    '        insertString += "," & IIf(IsNothing(pr.weld_trans_fillet_size_top), "Null", pr.weld_trans_fillet_size_top.ToString)
    '        insertString += "," & IIf(IsNothing(pr.weld_trans_eff_throat_top), "Null", pr.weld_trans_eff_throat_top.ToString)
    '        insertString += "," & IIf(IsNothing(pr.weld_long_type_top), "Null", "'" & pr.weld_long_type_top.ToString & "'")
    '        insertString += "," & IIf(IsNothing(pr.weld_long_length_top), "Null", pr.weld_long_length_top.ToString)
    '        insertString += "," & IIf(IsNothing(pr.weld_long_fillet_size_top), "Null", pr.weld_long_fillet_size_top.ToString)
    '        insertString += "," & IIf(IsNothing(pr.weld_long_eff_throat_top), "Null", pr.weld_long_eff_throat_top.ToString)
    '        insertString += "," & IIf(IsNothing(pr.conn_length_bot), "Null", pr.conn_length_bot.ToString)
    '        insertString += "," & IIf(IsNothing(pr.conn_length_top), "Null", pr.conn_length_top.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_comp_xx_f), "Null", pr.cap_comp_xx_f.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_comp_yy_f), "Null", pr.cap_comp_yy_f.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_tens_yield_f), "Null", pr.cap_tens_yield_f.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_tens_rupture_f), "Null", pr.cap_tens_rupture_f.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_shear_f), "Null", pr.cap_shear_f.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_bolt_shear_bot_f), "Null", pr.cap_bolt_shear_bot_f.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_bolt_shear_top_f), "Null", pr.cap_bolt_shear_top_f.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_boltshaft_bearing_nodeform_bot_f), "Null", pr.cap_boltshaft_bearing_nodeform_bot_f.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_boltshaft_bearing_deform_bot_f), "Null", pr.cap_boltshaft_bearing_deform_bot_f.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_boltshaft_bearing_nodeform_top_f), "Null", pr.cap_boltshaft_bearing_nodeform_top_f.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_boltshaft_bearing_deform_top_f), "Null", pr.cap_boltshaft_bearing_deform_top_f.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_boltreinf_bearing_nodeform_bot_f), "Null", pr.cap_boltreinf_bearing_nodeform_bot_f.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_boltreinf_bearing_deform_bot_f), "Null", pr.cap_boltreinf_bearing_deform_bot_f.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_boltreinf_bearing_nodeform_top_f), "Null", pr.cap_boltreinf_bearing_nodeform_top_f.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_boltreinf_bearing_deform_top_f), "Null", pr.cap_boltreinf_bearing_deform_top_f.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_weld_trans_bot_f), "Null", pr.cap_weld_trans_bot_f.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_weld_long_bot_f), "Null", pr.cap_weld_long_bot_f.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_weld_trans_top_f), "Null", pr.cap_weld_trans_top_f.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_weld_long_top_f), "Null", pr.cap_weld_long_top_f.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_comp_xx_g), "Null", pr.cap_comp_xx_g.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_comp_yy_g), "Null", pr.cap_comp_yy_g.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_tens_yield_g), "Null", pr.cap_tens_yield_g.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_tens_rupture_g), "Null", pr.cap_tens_rupture_g.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_shear_g), "Null", pr.cap_shear_g.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_bolt_shear_bot_g), "Null", pr.cap_bolt_shear_bot_g.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_bolt_shear_top_g), "Null", pr.cap_bolt_shear_top_g.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_boltshaft_bearing_nodeform_bot_g), "Null", pr.cap_boltshaft_bearing_nodeform_bot_g.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_boltshaft_bearing_deform_bot_g), "Null", pr.cap_boltshaft_bearing_deform_bot_g.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_boltshaft_bearing_nodeform_top_g), "Null", pr.cap_boltshaft_bearing_nodeform_top_g.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_boltshaft_bearing_deform_top_g), "Null", pr.cap_boltshaft_bearing_deform_top_g.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_boltreinf_bearing_nodeform_bot_g), "Null", pr.cap_boltreinf_bearing_nodeform_bot_g.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_boltreinf_bearing_deform_bot_g), "Null", pr.cap_boltreinf_bearing_deform_bot_g.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_boltreinf_bearing_nodeform_top_g), "Null", pr.cap_boltreinf_bearing_nodeform_top_g.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_boltreinf_bearing_deform_top_g), "Null", pr.cap_boltreinf_bearing_deform_top_g.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_weld_trans_bot_g), "Null", pr.cap_weld_trans_bot_g.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_weld_long_bot_g), "Null", pr.cap_weld_long_bot_g.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_weld_trans_top_g), "Null", pr.cap_weld_trans_top_g.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_weld_long_top_g), "Null", pr.cap_weld_long_top_g.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_comp_xx_h), "Null", pr.cap_comp_xx_h.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_comp_yy_h), "Null", pr.cap_comp_yy_h.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_tens_yield_h), "Null", pr.cap_tens_yield_h.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_tens_rupture_h), "Null", pr.cap_tens_rupture_h.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_shear_h), "Null", pr.cap_shear_h.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_bolt_shear_bot_h), "Null", pr.cap_bolt_shear_bot_h.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_bolt_shear_top_h), "Null", pr.cap_bolt_shear_top_h.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_boltshaft_bearing_nodeform_bot_h), "Null", pr.cap_boltshaft_bearing_nodeform_bot_h.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_boltshaft_bearing_deform_bot_h), "Null", pr.cap_boltshaft_bearing_deform_bot_h.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_boltshaft_bearing_nodeform_top_h), "Null", pr.cap_boltshaft_bearing_nodeform_top_h.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_boltshaft_bearing_deform_top_h), "Null", pr.cap_boltshaft_bearing_deform_top_h.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_boltreinf_bearing_nodeform_bot_h), "Null", pr.cap_boltreinf_bearing_nodeform_bot_h.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_boltreinf_bearing_deform_bot_h), "Null", pr.cap_boltreinf_bearing_deform_bot_h.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_boltreinf_bearing_nodeform_top_h), "Null", pr.cap_boltreinf_bearing_nodeform_top_h.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_boltreinf_bearing_deform_top_h), "Null", pr.cap_boltreinf_bearing_deform_top_h.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_weld_trans_bot_h), "Null", pr.cap_weld_trans_bot_h.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_weld_long_bot_h), "Null", pr.cap_weld_long_bot_h.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_weld_trans_top_h), "Null", pr.cap_weld_trans_top_h.ToString)
    '        insertString += "," & IIf(IsNothing(pr.cap_weld_long_top_h), "Null", pr.cap_weld_long_top_h.ToString)

    '        Return insertString
    '    End Function

    '    Private Function InsertPropBolt(ByVal pb As PropBolt) As String
    '        Dim insertString As String = ""

    '        insertString += "@BoltID"
    '        insertString += "," & IIf(IsNothing(pb.bolt_db_id), "Null", pb.bolt_db_id.ToString)
    '        insertString += "," & IIf(IsNothing(pb.name), "Null", "'" & pb.name.ToString & "'")
    '        insertString += "," & IIf(IsNothing(pb.description), "Null", "'" & pb.description.ToString & "'")
    '        insertString += "," & IIf(IsNothing(pb.diam), "Null", pb.diam.ToString)
    '        insertString += "," & IIf(IsNothing(pb.area), "Null", pb.area.ToString)
    '        insertString += "," & IIf(IsNothing(pb.fu_bolt), "Null", pb.fu_bolt.ToString)
    '        insertString += "," & IIf(IsNothing(pb.sleeve_diam_out), "Null", pb.sleeve_diam_out.ToString)
    '        insertString += "," & IIf(IsNothing(pb.sleeve_diam_in), "Null", pb.sleeve_diam_in.ToString)
    '        insertString += "," & IIf(IsNothing(pb.fu_sleeve), "Null", pb.fu_sleeve.ToString)
    '        insertString += "," & IIf(IsNothing(pb.bolt_n_sleeve_shear_revF), "Null", pb.bolt_n_sleeve_shear_revF.ToString)
    '        insertString += "," & IIf(IsNothing(pb.bolt_x_sleeve_shear_revF), "Null", pb.bolt_x_sleeve_shear_revF.ToString)
    '        insertString += "," & IIf(IsNothing(pb.bolt_n_sleeve_shear_revG), "Null", pb.bolt_n_sleeve_shear_revG.ToString)
    '        insertString += "," & IIf(IsNothing(pb.bolt_x_sleeve_shear_revG), "Null", pb.bolt_x_sleeve_shear_revG.ToString)
    '        insertString += "," & IIf(IsNothing(pb.bolt_n_sleeve_shear_revH), "Null", pb.bolt_n_sleeve_shear_revH.ToString)
    '        insertString += "," & IIf(IsNothing(pb.bolt_x_sleeve_shear_revH), "Null", pb.bolt_x_sleeve_shear_revH.ToString)
    '        insertString += "," & IIf(IsNothing(pb.rb_applied_revH), "Null", "'" & pb.rb_applied_revH.ToString & "'")

    '        Return insertString
    '    End Function

    '    Private Function InsertPropMatl(ByVal pm As PropMatl) As String
    '        Dim insertString As String = ""

    '        insertString += "@MatlID"
    '        insertString += "," & IIf(IsNothing(pm.matl_db_id), "Null", pm.matl_db_id.ToString)
    '        insertString += "," & IIf(IsNothing(pm.name), "Null", "'" & pm.name.ToString & "'")
    '        insertString += "," & IIf(IsNothing(pm.fy), "Null", pm.fy.ToString)
    '        insertString += "," & IIf(IsNothing(pm.fu), "Null", pm.fu.ToString)

    '        Return insertString
    '    End Function

    '#End Region

    '#Region "SQL Update Statements"
    '    Private Function UpdatePoleCriteria(ByVal pc As PoleCriteria) As String
    '        Dim updateString As String = ""

    '        updateString += "UPDATE pole_analysis_criteria SET "
    '        updateString += " criteria_id=" & IIf(IsNothing(pc.criteria_id), "Null", pc.criteria_id.ToString)
    '        updateString += ", upper_structure_type=" & IIf(IsNothing(pc.upper_structure_type), "Null", "'" & pc.upper_structure_type.ToString & "'")
    '        updateString += ", analysis_deg=" & IIf(IsNothing(pc.analysis_deg), "Null", pc.analysis_deg.ToString)
    '        updateString += ", geom_increment_length=" & IIf(IsNothing(pc.geom_increment_length), "Null", pc.geom_increment_length.ToString)
    '        updateString += ", vnum=" & IIf(IsNothing(pc.vnum), "Null", "'" & pc.vnum.ToString & "'")
    '        updateString += ", check_connections=" & IIf(IsNothing(pc.check_connections), "Null", "'" & pc.check_connections.ToString & "'")
    '        updateString += ", hole_deformation=" & IIf(IsNothing(pc.hole_deformation), "Null", "'" & pc.hole_deformation.ToString & "'")
    '        updateString += ", ineff_mod_check=" & IIf(IsNothing(pc.ineff_mod_check), "Null", "'" & pc.ineff_mod_check.ToString & "'")
    '        updateString += ", modified=" & IIf(IsNothing(pc.modified), "Null", "'" & pc.modified.ToString & "'")
    '        updateString += " WHERE ID = " & pc.criteria_id.ToString

    '        Return updateString
    '    End Function

    '    Private Function UpdatePoleSection(ByVal ps As PoleSection) As String
    '        Dim updateString As String = ""

    '        updateString += "UPDATE pole_section SET "
    '        updateString += " section_id=" & IIf(IsNothing(ps.section_id), "Null", ps.section_id.ToString)
    '        updateString += ", analysis_section_id=" & IIf(IsNothing(ps.analysis_section_id), "Null", ps.analysis_section_id.ToString)
    '        updateString += ", elev_bot=" & IIf(IsNothing(ps.elev_bot), "Null", ps.elev_bot.ToString)
    '        updateString += ", elev_top=" & IIf(IsNothing(ps.elev_top), "Null", ps.elev_top.ToString)
    '        updateString += ", length_section=" & IIf(IsNothing(ps.length_section), "Null", ps.length_section.ToString)
    '        updateString += ", length_splice=" & IIf(IsNothing(ps.length_splice), "Null", ps.length_splice.ToString)
    '        updateString += ", num_sides=" & IIf(IsNothing(ps.num_sides), "Null", ps.num_sides.ToString)
    '        updateString += ", diam_bot=" & IIf(IsNothing(ps.diam_bot), "Null", ps.diam_bot.ToString)
    '        updateString += ", diam_top=" & IIf(IsNothing(ps.diam_top), "Null", ps.diam_top.ToString)
    '        updateString += ", wall_thickness=" & IIf(IsNothing(ps.wall_thickness), "Null", ps.wall_thickness.ToString)
    '        updateString += ", bend_radius=" & IIf(IsNothing(ps.bend_radius), "Null", ps.bend_radius.ToString)
    '        updateString += ", steel_grade=" & IIf(IsNothing(ps.steel_grade_id), "Null", ps.steel_grade_id.ToString)
    '        updateString += ", pole_type=" & IIf(IsNothing(ps.pole_type), "Null", "'" & ps.pole_type.ToString & "'")
    '        updateString += ", section_name=" & IIf(IsNothing(ps.section_name), "Null", "'" & ps.section_name.ToString & "'")
    '        updateString += ", socket_length=" & IIf(IsNothing(ps.socket_length), "Null", ps.socket_length.ToString)
    '        updateString += ", weight_mult=" & IIf(IsNothing(ps.weight_mult), "Null", ps.weight_mult.ToString)
    '        updateString += ", wp_mult=" & IIf(IsNothing(ps.wp_mult), "Null", ps.wp_mult.ToString)
    '        updateString += ", af_factor=" & IIf(IsNothing(ps.af_factor), "Null", ps.af_factor.ToString)
    '        updateString += ", ar_factor=" & IIf(IsNothing(ps.ar_factor), "Null", ps.ar_factor.ToString)
    '        updateString += ", round_area_ratio=" & IIf(IsNothing(ps.round_area_ratio), "Null", ps.round_area_ratio.ToString)
    '        updateString += ", flat_area_ratio=" & IIf(IsNothing(ps.flat_area_ratio), "Null", ps.flat_area_ratio.ToString)
    '        updateString += " WHERE ID = " & ps.section_id.ToString

    '        Return updateString
    '    End Function

    '    Private Function UpdatePoleReinfSection(ByVal prs As PoleReinfSection) As String
    '        Dim updateString As String = ""

    '        updateString += "UPDATE pole_reinf_section SET "
    '        updateString += " section_ID=" & IIf(IsNothing(prs.section_ID), "Null", prs.section_ID.ToString)
    '        updateString += ", analysis_section_ID=" & IIf(IsNothing(prs.analysis_section_ID), "Null", prs.analysis_section_ID.ToString)
    '        updateString += ", elev_bot=" & IIf(IsNothing(prs.elev_bot), "Null", prs.elev_bot.ToString)
    '        updateString += ", elev_top=" & IIf(IsNothing(prs.elev_top), "Null", prs.elev_top.ToString)
    '        updateString += ", length_section=" & IIf(IsNothing(prs.length_section), "Null", prs.length_section.ToString)
    '        updateString += ", length_splice=" & IIf(IsNothing(prs.length_splice), "Null", prs.length_splice.ToString)
    '        updateString += ", num_sides=" & IIf(IsNothing(prs.num_sides), "Null", prs.num_sides.ToString)
    '        updateString += ", diam_bot=" & IIf(IsNothing(prs.diam_bot), "Null", prs.diam_bot.ToString)
    '        updateString += ", diam_top=" & IIf(IsNothing(prs.diam_top), "Null", prs.diam_top.ToString)
    '        updateString += ", wall_thickness=" & IIf(IsNothing(prs.wall_thickness), "Null", prs.wall_thickness.ToString)
    '        updateString += ", bend_radius=" & IIf(IsNothing(prs.bend_radius), "Null", prs.bend_radius.ToString)
    '        updateString += ", steel_grade_id=" & IIf(IsNothing(prs.steel_grade_id), "Null", prs.steel_grade_id.ToString)
    '        updateString += ", pole_type=" & IIf(IsNothing(prs.pole_type), "Null", "'" & prs.pole_type.ToString & "'")
    '        updateString += ", weight_mult=" & IIf(IsNothing(prs.weight_mult), "Null", prs.weight_mult.ToString)
    '        updateString += ", section_name=" & IIf(IsNothing(prs.section_name), "Null", "'" & prs.section_name.ToString & "'")
    '        updateString += ", socket_length=" & IIf(IsNothing(prs.socket_length), "Null", prs.socket_length.ToString)
    '        updateString += ", wp_mult=" & IIf(IsNothing(prs.wp_mult), "Null", prs.wp_mult.ToString)
    '        updateString += ", af_factor=" & IIf(IsNothing(prs.af_factor), "Null", prs.af_factor.ToString)
    '        updateString += ", ar_factor=" & IIf(IsNothing(prs.ar_factor), "Null", prs.ar_factor.ToString)
    '        updateString += ", round_area_ratio=" & IIf(IsNothing(prs.round_area_ratio), "Null", prs.round_area_ratio.ToString)
    '        updateString += ", flat_area_ratio=" & IIf(IsNothing(prs.flat_area_ratio), "Null", prs.flat_area_ratio.ToString)
    '        updateString += " WHERE ID = " & prs.section_ID.ToString

    '        Return updateString
    '    End Function

    '    Private Function UpdatePoleReinfGroup(ByVal prg As PoleReinfGroup) As String
    '        Dim updateString As String = ""

    '        updateString += "UPDATE pole_reinf_group SET "
    '        updateString += " reinf_group_id=" & IIf(IsNothing(prg.reinf_group_id), "Null", prg.reinf_group_id.ToString)
    '        updateString += ", elev_bot_actual=" & IIf(IsNothing(prg.elev_bot_actual), "Null", prg.elev_bot_actual.ToString)
    '        updateString += ", elev_bot_eff=" & IIf(IsNothing(prg.elev_bot_eff), "Null", prg.elev_bot_eff.ToString)
    '        updateString += ", elev_top_actual=" & IIf(IsNothing(prg.elev_top_actual), "Null", prg.elev_top_actual.ToString)
    '        updateString += ", elev_top_eff=" & IIf(IsNothing(prg.elev_top_eff), "Null", prg.elev_top_eff.ToString)
    '        updateString += ", reinf_db_id=" & IIf(IsNothing(prg.reinf_db_id), "Null", prg.reinf_db_id.ToString)
    '        updateString += " WHERE ID = " & prg.reinf_group_id.ToString

    '        Return updateString
    '    End Function

    '    Private Function UpdatePoleReinfDetail(ByVal prd As PoleReinfDetail) As String
    '        Dim updateString As String = ""

    '        updateString += "UPDATE pole_reinf_details SET "
    '        updateString += " reinf_id=" & IIf(IsNothing(prd.reinf_id), "Null", prd.reinf_id.ToString)
    '        updateString += ", pole_flat=" & IIf(IsNothing(prd.pole_flat), "Null", prd.pole_flat.ToString)
    '        updateString += ", horizontal_offset=" & IIf(IsNothing(prd.horizontal_offset), "Null", prd.horizontal_offset.ToString)
    '        updateString += ", rotation=" & IIf(IsNothing(prd.rotation), "Null", prd.rotation.ToString)
    '        updateString += ", note=" & IIf(IsNothing(prd.note), "Null", "'" & prd.note.ToString & "'")
    '        updateString += " WHERE ID = " & prd.reinf_id.ToString

    '        Return updateString
    '    End Function

    '    Private Function UpdatePoleIntGroup(ByVal pig As PoleIntGroup) As String
    '        Dim updateString As String = ""

    '        updateString += "UPDATE pole_interference_group SET "
    '        updateString += " interference_group_id=" & IIf(IsNothing(pig.interference_group_id), "Null", pig.interference_group_id.ToString)
    '        updateString += ", elev_bot=" & IIf(IsNothing(pig.elev_bot), "Null", pig.elev_bot.ToString)
    '        updateString += ", elev_top=" & IIf(IsNothing(pig.elev_top), "Null", pig.elev_top.ToString)
    '        updateString += ", width=" & IIf(IsNothing(pig.width), "Null", pig.width.ToString)
    '        updateString += ", description=" & IIf(IsNothing(pig.description), "Null", "'" & pig.description.ToString & "'")
    '        updateString += " WHERE ID = " & pig.interference_group_id.ToString

    '        Return updateString
    '    End Function

    '    Private Function UpdatePoleIntDetail(ByVal pid As PoleIntDetail) As String
    '        Dim updateString As String = ""

    '        updateString += "UPDATE pole_interference_details SET "
    '        updateString += " interference_id=" & IIf(IsNothing(pid.interference_id), "Null", pid.interference_id.ToString)
    '        updateString += ", pole_flat=" & IIf(IsNothing(pid.pole_flat), "Null", pid.pole_flat.ToString)
    '        updateString += ", horizontal_offset=" & IIf(IsNothing(pid.horizontal_offset), "Null", pid.horizontal_offset.ToString)
    '        updateString += ", rotation=" & IIf(IsNothing(pid.rotation), "Null", pid.rotation.ToString)
    '        updateString += ", note=" & IIf(IsNothing(pid.note), "Null", "'" & pid.note.ToString & "'")
    '        updateString += " WHERE ID = " & pid.interference_id.ToString

    '        Return updateString
    '    End Function

    '    Private Function UpdatePoleReinfResults(ByVal prr As PoleReinfResults) As String
    '        Dim updateString As String = ""

    '        updateString += "UPDATE pole_reinf_results SET "
    '        updateString += " section_id=" & IIf(IsNothing(prr.section_id), "Null", prr.section_id.ToString)
    '        updateString += ", work_order_seq_num=" & IIf(IsNothing(prr.work_order_seq_num), "Null", prr.work_order_seq_num.ToString)
    '        updateString += ", reinf_group_id=" & IIf(IsNothing(prr.reinf_group_id), "Null", prr.reinf_group_id.ToString)
    '        updateString += ", result_lkup_value=" & IIf(IsNothing(prr.result_lkup_value), "Null", prr.result_lkup_value.ToString)
    '        updateString += ", rating=" & IIf(IsNothing(prr.rating), "Null", prr.rating.ToString)
    '        updateString += " WHERE ID = " & prr.section_id.ToString

    '        Return updateString
    '    End Function

    '    Private Function UpdatePropReinf(ByVal pr As PropReinf) As String
    '        Dim updateString As String = ""

    '        updateString += "UPDATE memb_prop_flat_plate SET "
    '        updateString += " reinf_db_id=" & IIf(IsNothing(pr.reinf_db_id), "Null", pr.reinf_db_id.ToString)
    '        updateString += ", name=" & IIf(IsNothing(pr.name), "Null", "'" & pr.name.ToString & "'")
    '        updateString += ", type=" & IIf(IsNothing(pr.type), "Null", "'" & pr.type.ToString & "'")
    '        updateString += ", b=" & IIf(IsNothing(pr.b), "Null", pr.b.ToString)
    '        updateString += ", h=" & IIf(IsNothing(pr.h), "Null", pr.h.ToString)
    '        updateString += ", sr_diam=" & IIf(IsNothing(pr.sr_diam), "Null", pr.sr_diam.ToString)
    '        updateString += ", channel_thkns_web=" & IIf(IsNothing(pr.channel_thkns_web), "Null", pr.channel_thkns_web.ToString)
    '        updateString += ", channel_thkns_flange=" & IIf(IsNothing(pr.channel_thkns_flange), "Null", pr.channel_thkns_flange.ToString)
    '        updateString += ", channel_eo=" & IIf(IsNothing(pr.channel_eo), "Null", pr.channel_eo.ToString)
    '        updateString += ", channel_J=" & IIf(IsNothing(pr.channel_J), "Null", pr.channel_J.ToString)
    '        updateString += ", channel_Cw=" & IIf(IsNothing(pr.channel_Cw), "Null", pr.channel_Cw.ToString)
    '        updateString += ", area_gross=" & IIf(IsNothing(pr.area_gross), "Null", pr.area_gross.ToString)
    '        updateString += ", centroid=" & IIf(IsNothing(pr.centroid), "Null", pr.centroid.ToString)
    '        updateString += ", istension=" & IIf(IsNothing(pr.istension), "Null", "'" & pr.istension.ToString & "'")
    '        updateString += ", matl_id=" & IIf(IsNothing(pr.matl_id), "Null", pr.matl_id.ToString)
    '        updateString += ", Ix=" & IIf(IsNothing(pr.Ix), "Null", pr.Ix.ToString)
    '        updateString += ", Iy=" & IIf(IsNothing(pr.Iy), "Null", pr.Iy.ToString)
    '        updateString += ", Lu=" & IIf(IsNothing(pr.Lu), "Null", pr.Lu.ToString)
    '        updateString += ", Kx=" & IIf(IsNothing(pr.Kx), "Null", pr.Kx.ToString)
    '        updateString += ", Ky=" & IIf(IsNothing(pr.Ky), "Null", pr.Ky.ToString)
    '        updateString += ", bolt_hole_size=" & IIf(IsNothing(pr.bolt_hole_size), "Null", pr.bolt_hole_size.ToString)
    '        updateString += ", area_net=" & IIf(IsNothing(pr.area_net), "Null", pr.area_net.ToString)
    '        updateString += ", shear_lag=" & IIf(IsNothing(pr.shear_lag), "Null", pr.shear_lag.ToString)
    '        updateString += ", connection_type_bot=" & IIf(IsNothing(pr.connection_type_bot), "Null", "'" & pr.connection_type_bot.ToString & "'")
    '        updateString += ", connection_cap_revF_bot=" & IIf(IsNothing(pr.connection_cap_revF_bot), "Null", pr.connection_cap_revF_bot.ToString)
    '        updateString += ", connection_cap_revG_bot=" & IIf(IsNothing(pr.connection_cap_revG_bot), "Null", pr.connection_cap_revG_bot.ToString)
    '        updateString += ", connection_cap_revH_bot=" & IIf(IsNothing(pr.connection_cap_revH_bot), "Null", pr.connection_cap_revH_bot.ToString)
    '        updateString += ", bolt_type_id_bot=" & IIf(IsNothing(pr.bolt_type_id_bot), "Null", pr.bolt_type_id_bot.ToString)
    '        updateString += ", bolt_N_or_X_bot=" & IIf(IsNothing(pr.bolt_N_or_X_bot), "Null", "'" & pr.bolt_N_or_X_bot.ToString & "'")
    '        updateString += ", bolt_num_bot=" & IIf(IsNothing(pr.bolt_num_bot), "Null", pr.bolt_num_bot.ToString)
    '        updateString += ", bolt_spacing_bot=" & IIf(IsNothing(pr.bolt_spacing_bot), "Null", pr.bolt_spacing_bot.ToString)
    '        updateString += ", bolt_edge_dist_bot=" & IIf(IsNothing(pr.bolt_edge_dist_bot), "Null", pr.bolt_edge_dist_bot.ToString)
    '        updateString += ", FlangeOrBP_connected_bot=" & IIf(IsNothing(pr.FlangeOrBP_connected_bot), "Null", "'" & pr.FlangeOrBP_connected_bot.ToString & "'")
    '        updateString += ", weld_grade_bot=" & IIf(IsNothing(pr.weld_grade_bot), "Null", pr.weld_grade_bot.ToString)
    '        updateString += ", weld_trans_type_bot=" & IIf(IsNothing(pr.weld_trans_type_bot), "Null", "'" & pr.weld_trans_type_bot.ToString & "'")
    '        updateString += ", weld_trans_length_bot=" & IIf(IsNothing(pr.weld_trans_length_bot), "Null", pr.weld_trans_length_bot.ToString)
    '        updateString += ", weld_groove_depth_bot=" & IIf(IsNothing(pr.weld_groove_depth_bot), "Null", pr.weld_groove_depth_bot.ToString)
    '        updateString += ", weld_groove_angle_bot=" & IIf(IsNothing(pr.weld_groove_angle_bot), "Null", pr.weld_groove_angle_bot.ToString)
    '        updateString += ", weld_trans_fillet_size_bot=" & IIf(IsNothing(pr.weld_trans_fillet_size_bot), "Null", pr.weld_trans_fillet_size_bot.ToString)
    '        updateString += ", weld_trans_eff_throat_bot=" & IIf(IsNothing(pr.weld_trans_eff_throat_bot), "Null", pr.weld_trans_eff_throat_bot.ToString)
    '        updateString += ", weld_long_type_bot=" & IIf(IsNothing(pr.weld_long_type_bot), "Null", "'" & pr.weld_long_type_bot.ToString & "'")
    '        updateString += ", weld_long_length_bot=" & IIf(IsNothing(pr.weld_long_length_bot), "Null", pr.weld_long_length_bot.ToString)
    '        updateString += ", weld_long_fillet_size_bot=" & IIf(IsNothing(pr.weld_long_fillet_size_bot), "Null", pr.weld_long_fillet_size_bot.ToString)
    '        updateString += ", weld_long_eff_throat_bot=" & IIf(IsNothing(pr.weld_long_eff_throat_bot), "Null", pr.weld_long_eff_throat_bot.ToString)
    '        updateString += ", top_bot_connections_symmetrical=" & IIf(IsNothing(pr.top_bot_connections_symmetrical), "Null", "'" & pr.top_bot_connections_symmetrical.ToString & "'")
    '        updateString += ", connection_type_top=" & IIf(IsNothing(pr.connection_type_top), "Null", "'" & pr.connection_type_top.ToString & "'")
    '        updateString += ", connection_cap_revF_top=" & IIf(IsNothing(pr.connection_cap_revF_top), "Null", pr.connection_cap_revF_top.ToString)
    '        updateString += ", connection_cap_revG_top=" & IIf(IsNothing(pr.connection_cap_revG_top), "Null", pr.connection_cap_revG_top.ToString)
    '        updateString += ", connection_cap_revH_top=" & IIf(IsNothing(pr.connection_cap_revH_top), "Null", pr.connection_cap_revH_top.ToString)
    '        updateString += ", bolt_type_id_top=" & IIf(IsNothing(pr.bolt_type_id_top), "Null", pr.bolt_type_id_top.ToString)
    '        updateString += ", bolt_N_or_X_top=" & IIf(IsNothing(pr.bolt_N_or_X_top), "Null", "'" & pr.bolt_N_or_X_top.ToString & "'")
    '        updateString += ", bolt_num_top=" & IIf(IsNothing(pr.bolt_num_top), "Null", pr.bolt_num_top.ToString)
    '        updateString += ", bolt_spacing_top=" & IIf(IsNothing(pr.bolt_spacing_top), "Null", pr.bolt_spacing_top.ToString)
    '        updateString += ", bolt_edge_dist_top=" & IIf(IsNothing(pr.bolt_edge_dist_top), "Null", pr.bolt_edge_dist_top.ToString)
    '        updateString += ", FlangeOrBP_connected_top=" & IIf(IsNothing(pr.FlangeOrBP_connected_top), "Null", "'" & pr.FlangeOrBP_connected_top.ToString & "'")
    '        updateString += ", weld_grade_top=" & IIf(IsNothing(pr.weld_grade_top), "Null", pr.weld_grade_top.ToString)
    '        updateString += ", weld_trans_type_top=" & IIf(IsNothing(pr.weld_trans_type_top), "Null", "'" & pr.weld_trans_type_top.ToString & "'")
    '        updateString += ", weld_trans_length_top=" & IIf(IsNothing(pr.weld_trans_length_top), "Null", pr.weld_trans_length_top.ToString)
    '        updateString += ", weld_groove_depth_top=" & IIf(IsNothing(pr.weld_groove_depth_top), "Null", pr.weld_groove_depth_top.ToString)
    '        updateString += ", weld_groove_angle_top=" & IIf(IsNothing(pr.weld_groove_angle_top), "Null", pr.weld_groove_angle_top.ToString)
    '        updateString += ", weld_trans_fillet_size_top=" & IIf(IsNothing(pr.weld_trans_fillet_size_top), "Null", pr.weld_trans_fillet_size_top.ToString)
    '        updateString += ", weld_trans_eff_throat_top=" & IIf(IsNothing(pr.weld_trans_eff_throat_top), "Null", pr.weld_trans_eff_throat_top.ToString)
    '        updateString += ", weld_long_type_top=" & IIf(IsNothing(pr.weld_long_type_top), "Null", "'" & pr.weld_long_type_top.ToString & "'")
    '        updateString += ", weld_long_length_top=" & IIf(IsNothing(pr.weld_long_length_top), "Null", pr.weld_long_length_top.ToString)
    '        updateString += ", weld_long_fillet_size_top=" & IIf(IsNothing(pr.weld_long_fillet_size_top), "Null", pr.weld_long_fillet_size_top.ToString)
    '        updateString += ", weld_long_eff_throat_top=" & IIf(IsNothing(pr.weld_long_eff_throat_top), "Null", pr.weld_long_eff_throat_top.ToString)
    '        updateString += ", conn_length_bot=" & IIf(IsNothing(pr.conn_length_bot), "Null", pr.conn_length_bot.ToString)
    '        updateString += ", conn_length_top=" & IIf(IsNothing(pr.conn_length_top), "Null", pr.conn_length_top.ToString)
    '        updateString += ", cap_comp_xx_f=" & IIf(IsNothing(pr.cap_comp_xx_f), "Null", pr.cap_comp_xx_f.ToString)
    '        updateString += ", cap_comp_yy_f=" & IIf(IsNothing(pr.cap_comp_yy_f), "Null", pr.cap_comp_yy_f.ToString)
    '        updateString += ", cap_tens_yield_f=" & IIf(IsNothing(pr.cap_tens_yield_f), "Null", pr.cap_tens_yield_f.ToString)
    '        updateString += ", cap_tens_rupture_f=" & IIf(IsNothing(pr.cap_tens_rupture_f), "Null", pr.cap_tens_rupture_f.ToString)
    '        updateString += ", cap_shear_f=" & IIf(IsNothing(pr.cap_shear_f), "Null", pr.cap_shear_f.ToString)
    '        updateString += ", cap_bolt_shear_bot_f=" & IIf(IsNothing(pr.cap_bolt_shear_bot_f), "Null", pr.cap_bolt_shear_bot_f.ToString)
    '        updateString += ", cap_bolt_shear_top_f=" & IIf(IsNothing(pr.cap_bolt_shear_top_f), "Null", pr.cap_bolt_shear_top_f.ToString)
    '        updateString += ", cap_boltshaft_bearing_nodeform_bot_f=" & IIf(IsNothing(pr.cap_boltshaft_bearing_nodeform_bot_f), "Null", pr.cap_boltshaft_bearing_nodeform_bot_f.ToString)
    '        updateString += ", cap_boltshaft_bearing_deform_bot_f=" & IIf(IsNothing(pr.cap_boltshaft_bearing_deform_bot_f), "Null", pr.cap_boltshaft_bearing_deform_bot_f.ToString)
    '        updateString += ", cap_boltshaft_bearing_nodeform_top_f=" & IIf(IsNothing(pr.cap_boltshaft_bearing_nodeform_top_f), "Null", pr.cap_boltshaft_bearing_nodeform_top_f.ToString)
    '        updateString += ", cap_boltshaft_bearing_deform_top_f=" & IIf(IsNothing(pr.cap_boltshaft_bearing_deform_top_f), "Null", pr.cap_boltshaft_bearing_deform_top_f.ToString)
    '        updateString += ", cap_boltreinf_bearing_nodeform_bot_f=" & IIf(IsNothing(pr.cap_boltreinf_bearing_nodeform_bot_f), "Null", pr.cap_boltreinf_bearing_nodeform_bot_f.ToString)
    '        updateString += ", cap_boltreinf_bearing_deform_bot_f=" & IIf(IsNothing(pr.cap_boltreinf_bearing_deform_bot_f), "Null", pr.cap_boltreinf_bearing_deform_bot_f.ToString)
    '        updateString += ", cap_boltreinf_bearing_nodeform_top_f=" & IIf(IsNothing(pr.cap_boltreinf_bearing_nodeform_top_f), "Null", pr.cap_boltreinf_bearing_nodeform_top_f.ToString)
    '        updateString += ", cap_boltreinf_bearing_deform_top_f=" & IIf(IsNothing(pr.cap_boltreinf_bearing_deform_top_f), "Null", pr.cap_boltreinf_bearing_deform_top_f.ToString)
    '        updateString += ", cap_weld_trans_bot_f=" & IIf(IsNothing(pr.cap_weld_trans_bot_f), "Null", pr.cap_weld_trans_bot_f.ToString)
    '        updateString += ", cap_weld_long_bot_f=" & IIf(IsNothing(pr.cap_weld_long_bot_f), "Null", pr.cap_weld_long_bot_f.ToString)
    '        updateString += ", cap_weld_trans_top_f=" & IIf(IsNothing(pr.cap_weld_trans_top_f), "Null", pr.cap_weld_trans_top_f.ToString)
    '        updateString += ", cap_weld_long_top_f=" & IIf(IsNothing(pr.cap_weld_long_top_f), "Null", pr.cap_weld_long_top_f.ToString)
    '        updateString += ", cap_comp_xx_g=" & IIf(IsNothing(pr.cap_comp_xx_g), "Null", pr.cap_comp_xx_g.ToString)
    '        updateString += ", cap_comp_yy_g=" & IIf(IsNothing(pr.cap_comp_yy_g), "Null", pr.cap_comp_yy_g.ToString)
    '        updateString += ", cap_tens_yield_g=" & IIf(IsNothing(pr.cap_tens_yield_g), "Null", pr.cap_tens_yield_g.ToString)
    '        updateString += ", cap_tens_rupture_g=" & IIf(IsNothing(pr.cap_tens_rupture_g), "Null", pr.cap_tens_rupture_g.ToString)
    '        updateString += ", cap_shear_g=" & IIf(IsNothing(pr.cap_shear_g), "Null", pr.cap_shear_g.ToString)
    '        updateString += ", cap_bolt_shear_bot_g=" & IIf(IsNothing(pr.cap_bolt_shear_bot_g), "Null", pr.cap_bolt_shear_bot_g.ToString)
    '        updateString += ", cap_bolt_shear_top_g=" & IIf(IsNothing(pr.cap_bolt_shear_top_g), "Null", pr.cap_bolt_shear_top_g.ToString)
    '        updateString += ", cap_boltshaft_bearing_nodeform_bot_g=" & IIf(IsNothing(pr.cap_boltshaft_bearing_nodeform_bot_g), "Null", pr.cap_boltshaft_bearing_nodeform_bot_g.ToString)
    '        updateString += ", cap_boltshaft_bearing_deform_bot_g=" & IIf(IsNothing(pr.cap_boltshaft_bearing_deform_bot_g), "Null", pr.cap_boltshaft_bearing_deform_bot_g.ToString)
    '        updateString += ", cap_boltshaft_bearing_nodeform_top_g=" & IIf(IsNothing(pr.cap_boltshaft_bearing_nodeform_top_g), "Null", pr.cap_boltshaft_bearing_nodeform_top_g.ToString)
    '        updateString += ", cap_boltshaft_bearing_deform_top_g=" & IIf(IsNothing(pr.cap_boltshaft_bearing_deform_top_g), "Null", pr.cap_boltshaft_bearing_deform_top_g.ToString)
    '        updateString += ", cap_boltreinf_bearing_nodeform_bot_g=" & IIf(IsNothing(pr.cap_boltreinf_bearing_nodeform_bot_g), "Null", pr.cap_boltreinf_bearing_nodeform_bot_g.ToString)
    '        updateString += ", cap_boltreinf_bearing_deform_bot_g=" & IIf(IsNothing(pr.cap_boltreinf_bearing_deform_bot_g), "Null", pr.cap_boltreinf_bearing_deform_bot_g.ToString)
    '        updateString += ", cap_boltreinf_bearing_nodeform_top_g=" & IIf(IsNothing(pr.cap_boltreinf_bearing_nodeform_top_g), "Null", pr.cap_boltreinf_bearing_nodeform_top_g.ToString)
    '        updateString += ", cap_boltreinf_bearing_deform_top_g=" & IIf(IsNothing(pr.cap_boltreinf_bearing_deform_top_g), "Null", pr.cap_boltreinf_bearing_deform_top_g.ToString)
    '        updateString += ", cap_weld_trans_bot_g=" & IIf(IsNothing(pr.cap_weld_trans_bot_g), "Null", pr.cap_weld_trans_bot_g.ToString)
    '        updateString += ", cap_weld_long_bot_g=" & IIf(IsNothing(pr.cap_weld_long_bot_g), "Null", pr.cap_weld_long_bot_g.ToString)
    '        updateString += ", cap_weld_trans_top_g=" & IIf(IsNothing(pr.cap_weld_trans_top_g), "Null", pr.cap_weld_trans_top_g.ToString)
    '        updateString += ", cap_weld_long_top_g=" & IIf(IsNothing(pr.cap_weld_long_top_g), "Null", pr.cap_weld_long_top_g.ToString)
    '        updateString += ", cap_comp_xx_h=" & IIf(IsNothing(pr.cap_comp_xx_h), "Null", pr.cap_comp_xx_h.ToString)
    '        updateString += ", cap_comp_yy_h=" & IIf(IsNothing(pr.cap_comp_yy_h), "Null", pr.cap_comp_yy_h.ToString)
    '        updateString += ", cap_tens_yield_h=" & IIf(IsNothing(pr.cap_tens_yield_h), "Null", pr.cap_tens_yield_h.ToString)
    '        updateString += ", cap_tens_rupture_h=" & IIf(IsNothing(pr.cap_tens_rupture_h), "Null", pr.cap_tens_rupture_h.ToString)
    '        updateString += ", cap_shear_h=" & IIf(IsNothing(pr.cap_shear_h), "Null", pr.cap_shear_h.ToString)
    '        updateString += ", cap_bolt_shear_bot_h=" & IIf(IsNothing(pr.cap_bolt_shear_bot_h), "Null", pr.cap_bolt_shear_bot_h.ToString)
    '        updateString += ", cap_bolt_shear_top_h=" & IIf(IsNothing(pr.cap_bolt_shear_top_h), "Null", pr.cap_bolt_shear_top_h.ToString)
    '        updateString += ", cap_boltshaft_bearing_nodeform_bot_h=" & IIf(IsNothing(pr.cap_boltshaft_bearing_nodeform_bot_h), "Null", pr.cap_boltshaft_bearing_nodeform_bot_h.ToString)
    '        updateString += ", cap_boltshaft_bearing_deform_bot_h=" & IIf(IsNothing(pr.cap_boltshaft_bearing_deform_bot_h), "Null", pr.cap_boltshaft_bearing_deform_bot_h.ToString)
    '        updateString += ", cap_boltshaft_bearing_nodeform_top_h=" & IIf(IsNothing(pr.cap_boltshaft_bearing_nodeform_top_h), "Null", pr.cap_boltshaft_bearing_nodeform_top_h.ToString)
    '        updateString += ", cap_boltshaft_bearing_deform_top_h=" & IIf(IsNothing(pr.cap_boltshaft_bearing_deform_top_h), "Null", pr.cap_boltshaft_bearing_deform_top_h.ToString)
    '        updateString += ", cap_boltreinf_bearing_nodeform_bot_h=" & IIf(IsNothing(pr.cap_boltreinf_bearing_nodeform_bot_h), "Null", pr.cap_boltreinf_bearing_nodeform_bot_h.ToString)
    '        updateString += ", cap_boltreinf_bearing_deform_bot_h=" & IIf(IsNothing(pr.cap_boltreinf_bearing_deform_bot_h), "Null", pr.cap_boltreinf_bearing_deform_bot_h.ToString)
    '        updateString += ", cap_boltreinf_bearing_nodeform_top_h=" & IIf(IsNothing(pr.cap_boltreinf_bearing_nodeform_top_h), "Null", pr.cap_boltreinf_bearing_nodeform_top_h.ToString)
    '        updateString += ", cap_boltreinf_bearing_deform_top_h=" & IIf(IsNothing(pr.cap_boltreinf_bearing_deform_top_h), "Null", pr.cap_boltreinf_bearing_deform_top_h.ToString)
    '        updateString += ", cap_weld_trans_bot_h=" & IIf(IsNothing(pr.cap_weld_trans_bot_h), "Null", pr.cap_weld_trans_bot_h.ToString)
    '        updateString += ", cap_weld_long_bot_h=" & IIf(IsNothing(pr.cap_weld_long_bot_h), "Null", pr.cap_weld_long_bot_h.ToString)
    '        updateString += ", cap_weld_trans_top_h=" & IIf(IsNothing(pr.cap_weld_trans_top_h), "Null", pr.cap_weld_trans_top_h.ToString)
    '        updateString += ", cap_weld_long_top_h=" & IIf(IsNothing(pr.cap_weld_long_top_h), "Null", pr.cap_weld_long_top_h.ToString)
    '        updateString += " WHERE ID = " & pr.reinf_db_id.ToString

    '        Return updateString
    '    End Function

    '    Private Function UpdatePropBolt(ByVal pb As PropBolt) As String
    '        Dim updateString As String = ""

    '        updateString += "UPDATE bolt_prop_flat_plate SET "
    '        updateString += " bolt_db_id=" & IIf(IsNothing(pb.bolt_db_id), "Null", pb.bolt_db_id.ToString)
    '        updateString += ", name=" & IIf(IsNothing(pb.name), "Null", "'" & pb.name.ToString & "'")
    '        updateString += ", description=" & IIf(IsNothing(pb.description), "Null", "'" & pb.description.ToString & "'")
    '        updateString += ", diam=" & IIf(IsNothing(pb.diam), "Null", pb.diam.ToString)
    '        updateString += ", area=" & IIf(IsNothing(pb.area), "Null", pb.area.ToString)
    '        updateString += ", fu_bolt=" & IIf(IsNothing(pb.fu_bolt), "Null", pb.fu_bolt.ToString)
    '        updateString += ", sleeve_diam_out=" & IIf(IsNothing(pb.sleeve_diam_out), "Null", pb.sleeve_diam_out.ToString)
    '        updateString += ", sleeve_diam_in=" & IIf(IsNothing(pb.sleeve_diam_in), "Null", pb.sleeve_diam_in.ToString)
    '        updateString += ", fu_sleeve=" & IIf(IsNothing(pb.fu_sleeve), "Null", pb.fu_sleeve.ToString)
    '        updateString += ", bolt_n_sleeve_shear_revF=" & IIf(IsNothing(pb.bolt_n_sleeve_shear_revF), "Null", pb.bolt_n_sleeve_shear_revF.ToString)
    '        updateString += ", bolt_x_sleeve_shear_revF=" & IIf(IsNothing(pb.bolt_x_sleeve_shear_revF), "Null", pb.bolt_x_sleeve_shear_revF.ToString)
    '        updateString += ", bolt_n_sleeve_shear_revG=" & IIf(IsNothing(pb.bolt_n_sleeve_shear_revG), "Null", pb.bolt_n_sleeve_shear_revG.ToString)
    '        updateString += ", bolt_x_sleeve_shear_revG=" & IIf(IsNothing(pb.bolt_x_sleeve_shear_revG), "Null", pb.bolt_x_sleeve_shear_revG.ToString)
    '        updateString += ", bolt_n_sleeve_shear_revH=" & IIf(IsNothing(pb.bolt_n_sleeve_shear_revH), "Null", pb.bolt_n_sleeve_shear_revH.ToString)
    '        updateString += ", bolt_x_sleeve_shear_revH=" & IIf(IsNothing(pb.bolt_x_sleeve_shear_revH), "Null", pb.bolt_x_sleeve_shear_revH.ToString)
    '        updateString += ", rb_applied_revH=" & IIf(IsNothing(pb.rb_applied_revH), "Null", "'" & pb.rb_applied_revH.ToString & "'")
    '        updateString += " WHERE ID = " & pb.bolt_db_id.ToString

    '        Return updateString
    '    End Function

    '    Private Function UpdatePropMatl(ByVal pm As PropMatl) As String
    '        Dim updateString As String = ""

    '        updateString += "UPDATE matl_prop_flat_plate SET "
    '        updateString += " matl_db_id=" & IIf(IsNothing(pm.matl_db_id), "Null", pm.matl_db_id.ToString)
    '        updateString += ", name=" & IIf(IsNothing(pm.name), "Null", "'" & pm.name.ToString & "'")
    '        updateString += ", fy=" & IIf(IsNothing(pm.fy), "Null", pm.fy.ToString)
    '        updateString += ", fu=" & IIf(IsNothing(pm.fu), "Null", pm.fu.ToString)
    '        updateString += " WHERE ID = " & pm.matl_db_id.ToString

    '        Return updateString
    '    End Function

    '#End Region

    '#Region "General"
    '    Public Sub Clear()
    '        ExcelFilePath = ""
    '        Poles.Clear()
    '    End Sub

    '    Private Function CCIpoleSQLDataTables() As List(Of SQLParameter)
    '        Dim MyParameters As New List(Of SQLParameter)

    '        MyParameters.Add(New SQLParameter("CCIpole Criteria SQL", "CCIpole (SELECT Criteria).sql"))
    '        MyParameters.Add(New SQLParameter("CCIpole Pole Sections SQL", "CCIpole (SELECT Pole Sections).sql"))
    '        MyParameters.Add(New SQLParameter("CCIpole Pole Reinf Sections SQL", "CCIpole (SELECT Pole Reinf Sections).sql"))
    '        MyParameters.Add(New SQLParameter("CCIpole Reinf Groups SQL", "CCIpole (SELECT Reinf Groups).sql"))
    '        MyParameters.Add(New SQLParameter("CCIpole Reinf Details SQL", "CCIpole (SELECT Reinf Details).sql"))
    '        MyParameters.Add(New SQLParameter("CCIpole Int Groups SQL", "CCIpole (SELECT Int Groups).sql"))
    '        MyParameters.Add(New SQLParameter("CCIpole Int Details SQL", "CCIpole (SELECT Int Details).sql"))
    '        MyParameters.Add(New SQLParameter("CCIpole Pole Reinf Results SQL", "CCIpole (SELECT Pole Reinf Results).sql"))
    '        MyParameters.Add(New SQLParameter("CCIpole Reinf Property Details SQL", "CCIpole (SELECT Prop Reinfs).sql"))
    '        MyParameters.Add(New SQLParameter("CCIpole Bolt Property Details SQL", "CCIpole (SELECT Prop Bolts).sql"))
    '        MyParameters.Add(New SQLParameter("CCIpole Matl Property Details SQL", "CCIpole (SELECT Prop Matls).sql"))

    '        Return MyParameters
    '    End Function

    '    Private Function CCIpoleExcelDTParameters() As List(Of EXCELDTParameter)
    '        Dim MyParameters As New List(Of EXCELDTParameter)

    '        MyParameters.Add(New EXCELDTParameter("CCIpole Criteria EXCEL", "A2:L3", "Analysis Criteria (SAPI)"))
    '        MyParameters.Add(New EXCELDTParameter("CCIpole Pole Sections EXCEL", "A2:U20", "Unreinf Pole (SAPI)"))
    '        MyParameters.Add(New EXCELDTParameter("CCIpole Pole Reinf Sections EXCEL", "A2:S200", "Reinf Pole (SAPI)"))
    '        MyParameters.Add(New EXCELDTParameter("CCIpole Reinf Groups EXCEL", "A2:F50", "Reinf Groups (SAPI)"))
    '        MyParameters.Add(New EXCELDTParameter("CCIpole Reinf Details EXCEL", "A2:F200", "Reinf ID (SAPI)"))
    '        MyParameters.Add(New EXCELDTParameter("CCIpole Int Groups EXCEL", "A2:E50", "Interference Groups (SAPI)"))
    '        MyParameters.Add(New EXCELDTParameter("CCIpole Int Details EXCEL", "A2:F200", "Interference ID (SAPI)"))
    '        MyParameters.Add(New EXCELDTParameter("CCIpole Pole Reinf Results EXCEL", "A2:E200", "Reinf Results (SAPI)"))
    '        MyParameters.Add(New EXCELDTParameter("CCIpole Reinf Property Details EXCEL", "A2:DV50", "Reinforcements (SAPI)"))
    '        MyParameters.Add(New EXCELDTParameter("CCIpole Bolt Property Details EXCEL", "A2:P10", "Bolts (SAPI)"))
    '        MyParameters.Add(New EXCELDTParameter("CCIpole Matl Property Details EXCEL", "A2:D10", "Materials (SAPI)"))

    '        Return MyParameters
    '    End Function
    '#End Region

End Class