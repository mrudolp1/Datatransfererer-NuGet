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
    Public Property sqlDrilledPiers As New List(Of DrilledPier)
    Private Property DrilledPierTemplatePath As String = "C:\Users\" & Environment.UserName & "\Desktop\WIP - Drilled Pier Foundation (5.1.0) - 10-14-21 - EDIT.xlsm"
    Private Property DrilledPierFileType As DocumentFormat = DocumentFormat.Xlsm

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
    Sub CreateSQLDrilledPiers(ByRef DrilledPiers As List(Of DrilledPier))
        Dim refid As Integer

        Dim DrilledPierLoader As String

        'Load data to get pier and pad details data for the existing structure model
        For Each item As SQLParameter In DrilledPierSQLDataTables()
            DrilledPierLoader = QueryBuilderFromFile(queryPath & "Drilled Pier\" & item.sqlQuery).Replace("[EXISTING MODEL]", GetExistingModelQuery())
            DoDaSQL.sqlLoader(DrilledPierLoader, item.sqlDatatable, ds, dpDB, dpID, "0")
            'If ds.Tables(item.sqlDatatable).Rows.Count = 0 Then Return False 'This may need adjusted since some tables can be empty
        Next

        'Custom Section to transfer data for the drilled pier tool. Needs to be adjusted for each tool.
        For Each DrilledPierDataRow As DataRow In ds.Tables("Drilled Pier General Details SQL").Rows
            refid = CType(DrilledPierDataRow.Item("ID"), Integer)
            DrilledPiers.Add(New DrilledPier(DrilledPierDataRow, refid))
        Next

    End Sub

    Public Function LoadFromEDS() As Boolean
        CreateSQLDrilledPiers(DrilledPiers)
        Return True
    End Function 'Create Drilled Pier objects based on what is saved in EDS

    Public Sub LoadFromExcel()

        For Each item As EXCELDTParameter In DrilledPierExcelDTParameters()
            'Get tables from excel file 
            ds.Tables.Add(ExcelDatasourceToDataTable(GetExcelDataSource(ExcelFilePath, item.xlsSheet, item.xlsRange), item.xlsDatatable))
        Next

        Dim refID As Integer
        Dim refCol As String

        'Custom Section to transfer data for the drilled pier tool. Needs to be adjusted for each tool.
        For Each DrilledPierDataRow As DataRow In ds.Tables("Drilled Pier General Details EXCEL").Rows

            refCol = "local_drilled_pier_id"
            refID = CType(DrilledPierDataRow.Item(refCol), Integer)

            DrilledPiers.Add(New DrilledPier(DrilledPierDataRow, refID, refCol))
        Next

        'Pull SQL data, if applicable, to compare with excel data
        CreateSQLDrilledPiers(sqlDrilledPiers)

        'If sqlGuyedAnchorBlocks.Count > 0 Then 'same as if checking for id in tool, if ID greater than 0.
        For Each fnd As DrilledPier In DrilledPiers
            Dim IDmatch As Boolean = False
            'If fnd.ID > 0 Then 'can skip loading SQL data if id = 0 (first time adding to EDS)
            If fnd.pier_id > 0 Then 'can skip loading SQL data if id = 0 (first time adding to EDS)
                For Each sqlfnd As DrilledPier In sqlDrilledPiers 'MRP - UPDATES NEEDED!!! Chackchanges needs updated to apply to multiple objects within the same tool. Only want one foundation group per tool
                    'If fnd.ID = sqlfnd.ID Then
                    If fnd.pier_id = sqlfnd.pier_id Then
                        IDmatch = True
                        If CheckChanges(fnd, sqlfnd) Then
                            isModelNeeded = True
                            isfndGroupNeeded = True
                            isDrilledPierNeeded = True
                        End If
                        Exit For
                    End If
                Next
                'IF ID match = False, Save the data because nothing exists in sql (could have copied tool from a different BU)
                If IDmatch = False Then
                    isModelNeeded = True
                    isfndGroupNeeded = True
                    isDrilledPierNeeded = True
                End If
            Else
                'Save the data because nothing exists in sql
                isModelNeeded = True
                isfndGroupNeeded = True
                isDrilledPierNeeded = True
            End If
        Next

    End Sub 'Create Drilled Pier objects based on what is coming from the excel file
#End Region

#Region "Save Data"
    Sub Save1DrilledPier(ByVal dp As DrilledPier)

        'Dim firstOne As Boolean = True
        'Dim mySoils As String = ""
        'Dim mySections As String = ""
        'Dim myRebar As String = ""
        'Dim myProfiles As String = ""

        'For Each dp As DrilledPier In DrilledPiers

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

        'create new information only once per tool, rather than each instance of the foundation from the tool
        If firstDrilledPier Then

            'Determine if new model ID needs created. Shouldn't be added to all individual tools (only needs to be referenced once)
            If isModelNeeded Then
                DrilledPierSaver = DrilledPierSaver.Replace("'[Model ID Needed]'", 1)
            Else
                DrilledPierSaver = DrilledPierSaver.Replace("'[Model ID Needed]'", 0)
            End If

            'Determine if new foundation group ID needs created. 
            If isfndGroupNeeded Then
                DrilledPierSaver = DrilledPierSaver.Replace("'[Fnd GRP ID Needed]'", 1)
            Else
                DrilledPierSaver = DrilledPierSaver.Replace("'[Fnd GRP ID Needed]'", 0)
            End If

        Else

            DrilledPierSaver = DrilledPierSaver.Replace("'[Model ID Needed]'", 0)
            DrilledPierSaver = DrilledPierSaver.Replace("'[Fnd GRP ID Needed]'", 0)

        End If

        'Determine if new Drilled Pier ID needs created
        If isDrilledPierNeeded Then
            DrilledPierSaver = DrilledPierSaver.Replace("'[DRILLED PIER ID Needed]'", 1)
        Else
            DrilledPierSaver = DrilledPierSaver.Replace("'[DRILLED PIER ID Needed]'", 0)
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
                DrilledPierSaver = DrilledPierSaver.Replace("INSERT INTO fnd.belled_pier_details VALUES ([INSERT ALL BELLED PIER DETAILS])", "")
                DrilledPierSaver = DrilledPierSaver.Replace("END --INSERT Belled Pier information if required", "--END --INSERT Belled Pier information if required")
            End If 'Add Belled Pier INSERT Statment

            If dp.embedded_pole Then
                DrilledPierSaver = DrilledPierSaver.Replace("[INSERT ALL EMBEDDED POLE DETAILS]", InsertDrilledPierEmbed(dp.embed_details))
            Else
                DrilledPierSaver = DrilledPierSaver.Replace("BEGIN --Embedded Pole", "--BEGIN --Embedded Pole")
                DrilledPierSaver = DrilledPierSaver.Replace("IF @IsEmbed = 'True'", "--IF @IsEmbed = 'True'")
                DrilledPierSaver = DrilledPierSaver.Replace("INSERT INTO fnd.embedded_pole_details OUTPUT INSERTED.ID INTO @EmbeddedPole VALUES ([INSERT ALL EMBEDDED POLE DETAILS])", "")
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
            DrilledPierSaver = DrilledPierSaver.Replace("([INSERT ALL DRILLED PIER PROFILES])", myProfiles)
            firstOne = True

            mySoils = ""
            mySections = ""
            myProfiles = ""
            'Else
            '    Dim tempUpdater As String = ""
            '    tempUpdater += UpdateDrilledPierDetail(dp)

            '    'comment out soil layer insertion. Added in next step if a layer does not have an ID
            '    DrilledPierSaver = DrilledPierSaver.Replace("INSERT INTO drilled_pier_soil_layer VALUES ([INSERT ALL SOIL LAYERS])", "--INSERT INTO drilled_pier_soil_layer VALUES ([INSERT ALL SOIL LAYERS])")

            '    For Each dpsl As DrilledPierSoilLayer In dp.soil_layers
            '        If dpsl.soil_layer_id = 0 Or IsDBNull(dpsl.soil_layer_id) Then
            '            tempUpdater += "INSERT INTO drilled_pier_soil_layers VALUES (" & InsertDrilledPierSoilLayer(dpsl) & ") " & vbNewLine
            '        Else
            '            tempUpdater += UpdateDrilledPierSoilLayer(dpsl)
            '        End If
            '    Next

            '    If dp.belled_pier Then
            '        If dp.belled_details.belled_pier_id = 0 Or IsDBNull(dp.belled_details.belled_pier_id) Then
            '            tempUpdater += "INSERT INTO belled_pier_details VALUES (" & InsertDrilledPierBell(dp.belled_details) & ") " & vbNewLine
            '        Else
            '            tempUpdater += UpdateDrilledPierBell(dp.belled_details)
            '        End If
            '    Else
            '        DrilledPierSaver = DrilledPierSaver.Replace("BEGIN --Belled Pier", "--BEGIN --Belled Pier")
            '        DrilledPierSaver = DrilledPierSaver.Replace("IF @IsBelled = 'True'", "--IF @IsBelled = 'True'")
            '        DrilledPierSaver = DrilledPierSaver.Replace("INSERT INTO belled_pier_details VALUES ([INSERT ALL BELLED PIER DETAILS])", "")
            '        DrilledPierSaver = DrilledPierSaver.Replace("END --INSERT Belled Pier information if required", "--END --INSERT Belled Pier information if required")
            '    End If

            '    If dp.embedded_pole Then
            '        If dp.embed_details.embedded_id = 0 Or IsDBNull(dp.embed_details.embedded_id) Then
            '            tempUpdater += "BEGIN INSERT INTO embedded_pole_details OUTPUT INSERTED.ID INTO @EmbeddedPole VALUES (" & InsertDrilledPierEmbed(dp.embed_details) & ") " & vbNewLine & " SELECT @EmbedID=EmbedID FROM @EmbeddedPole"
            '            tempUpdater += " END " & vbNewLine
            '        Else
            '            tempUpdater += UpdateDrilledPierEmbed(dp.embed_details)
            '        End If
            '    Else
            '        DrilledPierSaver = DrilledPierSaver.Replace("BEGIN --Embedded Pole", "--BEGIN --Embedded Pole")
            '        DrilledPierSaver = DrilledPierSaver.Replace("IF @IsEmbed = 'True'", "--IF @IsEmbed = 'True'")
            '        DrilledPierSaver = DrilledPierSaver.Replace("INSERT INTO embedded_pole_details OUTPUT INSERTED.ID INTO @EmbeddedPole VALUES ([INSERT ALL EMBEDDED POLE DETAILS])", "")
            '        DrilledPierSaver = DrilledPierSaver.Replace("SELECT @EmbedID=EmbedID FROM @EmbeddedPole", "--SELECT @EmbedID=EmbedID FROM @EmbeddedPole")
            '        DrilledPierSaver = DrilledPierSaver.Replace("END --INSERT Embedded Pole information if required", "--END --INSERT Embedded Pole information if required")
            '    End If

            '    For Each dpSec As DrilledPierSection In dp.sections
            '        If dpSec.section_id = 0 Or IsDBNull(dpSec.section_id) Then
            '            tempUpdater += "BEGIN INSERT INTO drilled_pier_section OUTPUT INSERTED.ID INTO @DrilledPierSection VALUES (" & InsertDrilledPierSection(dpSec) & ") " & vbNewLine & " SELECT @SecID=SecID FROM @DrilledPierSection"
            '            For Each dpreb As DrilledPierRebar In dpSec.rebar
            '                tempUpdater += "INSERT INTO drilled_pier_rebar VALUES (" & InsertDrilledPierRebar(dpreb) & ") " & vbNewLine
            '            Next
            '            tempUpdater += " END " & vbNewLine
            '        Else
            '            tempUpdater += UpdateDrilledPierSection(dpSec)
            '            For Each dpreb As DrilledPierRebar In dpSec.rebar
            '                If dpreb.rebar_id = 0 Or IsDBNull(dpreb.rebar_id) Then
            '                    tempUpdater += "INSERT INTO drilled_pier_rebar VALUES (" & InsertDrilledPierRebar(dpreb).Replace("@SecID", dpSec.section_id.ToString) & ") " & vbNewLine
            '                Else
            '                    tempUpdater += UpdateDrilledPierRebar(dpreb)
            '                End If
            '            Next
            '        End If
            '    Next

            '    DrilledPierSaver = DrilledPierSaver.Replace("INSERT INTO drilled_pier_profile VALUES ([INSERT ALL PIER PROFILES])", "--INSERT INTO drilled_pier_profile VALUES ([INSERT ALL PIER PROFILES])")
            '    For Each dpp As DrilledPierProfile In dp.drilled_pier_profiles
            '        If dpp.profile_id = 0 Or IsDBNull(dpp.profile_id) Then
            '            tempUpdater += "INSERT INTO drilled_pier_profile VALUES (" & InsertDrilledPierProfile(dpp) & ") " & vbNewLine
            '        Else
            '            tempUpdater += UpdateDrilledPierProfile(dpp)
            '        End If
            '    Next

            '    DrilledPierSaver = DrilledPierSaver.Replace("SELECT * FROM TEMPORARY", tempUpdater)
        End If

        DrilledPierSaver = DrilledPierSaver.Replace("[INSERT ALL DRILLED PIER DETAILS]", InsertDrilledPierDetail(dp))

        sqlSender(DrilledPierSaver, dpDB, dpID, "0")

        'Next

    End Sub

    Dim firstDrilledPier As Boolean = True
    Dim firstOne As Boolean = True
    Dim mySoils As String = ""
    Dim mySections As String = ""
    Dim myRebar As String = ""
    Dim myProfiles As String = ""

    Public Sub SaveToEDS()
        For Each dp As DrilledPier In DrilledPiers
            Save1DrilledPier(dp)
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

            Dim colCounter As Integer = 6
            Dim myCol As String
            Dim rowStart As Integer = 56

            For Each dp As DrilledPier In DrilledPiers

                colCounter = 6 + dp.local_drilled_pier_id
                myCol = GetExcelColumnName(colCounter)

                'DRILLED PIER DETAILS
                If Not IsNothing(dp.pier_id) Then
                    .Worksheets("Database").Range(myCol & rowStart - 54).Value = CType(dp.pier_id, Integer)
                Else .Worksheets("Database").Range(myCol & rowStart - 54).ClearContents
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
                If CType(dp.groundwater_depth, Double) = -1 Then
                    .Worksheets("Database").Range(myCol & rowStart + 17).Value = "N/A"
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

                    If Not IsNothing(dpSec.section_id) Then
                        .Worksheets("Database").Range(myCol & rowStart - 54 + secCount).Value = CType(dpSec.section_id, Integer)
                    Else .Worksheets("Database").Range(myCol & rowStart - 54 + secCount).ClearContents
                    End If

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
                    If Not IsNothing(dpSec.rho_override) Then
                        .Worksheets("Database").Range(myCol & rowStart + 4392 + secCount).Value = CType(dpSec.rho_override, Double)
                    Else .Worksheets("Database").Range(myCol & rowStart + 4392 + secCount).ClearContents
                    End If

                    If secCount > 1 Then depth += 1
                    If Not IsNothing(dpSec.bottom_elevation) Then
                        .Worksheets("Database").Range(myCol & rowStart + 12 + depth).Value = CType(dpSec.bottom_elevation, Double)
                    Else .Worksheets("Database").Range(myCol & rowStart + 12 + depth).ClearContents
                    End If

                    'DRILLED PIER REBAR
                    Dim rebCount As Integer = 1

                    For Each dpReb As DrilledPierRebar In dpSec.rebar

                        If Not IsNothing(dpReb.rebar_id) Then
                            .Worksheets("Database").Range(myCol & rowStart - 48 + 3 * (secCount - 1) + (rebCount - 1)).Value = CType(dpReb.rebar_id, Integer)
                        Else .Worksheets("Database").Range(myCol & rowStart - 48 + 3 * (secCount - 1) + (rebCount - 1)).ClearContents
                        End If

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
                        ElseIf rebCount = 3 Then
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
                    'secStart += secBump
                    secStart += 15

                Next

                'BELLED PIER
                If dp.belled_pier = True Then

                    If Not IsNothing(dp.belled_details.belled_pier_id) Then
                        .Worksheets("Database").Range(myCol & rowStart - 3).Value = CType(dp.belled_details.belled_pier_id, Integer)
                    Else .Worksheets("Database").Range(myCol & rowStart - 3).ClearContents
                    End If

                    .Worksheets("Database").Range(myCol & rowStart + 112).Value = CType(dp.belled_pier, Boolean)
                    If Not IsNothing(dp.belled_details.bottom_diameter_of_bell) Then
                        .Worksheets("Database").Range(myCol & rowStart + 113).Value = CType(dp.belled_details.bottom_diameter_of_bell, Double)
                    Else .Worksheets("Database").Range(myCol & rowStart + 113).ClearContents
                    End If
                    If Not IsNothing(dp.belled_details.bell_angle) Then
                        .Worksheets("Database").Range(myCol & rowStart + 114).Value = CType(dp.belled_details.bell_angle, Double)
                    Else .Worksheets("Database").Range(myCol & rowStart + 114).ClearContents
                    End If
                    .Worksheets("Database").Range(myCol & rowStart + 115).Value = CType(dp.belled_details.bell_input_type, String)
                    If Not IsNothing(dp.belled_details.bell_height) Then
                        .Worksheets("Database").Range(myCol & rowStart + 116).Value = CType(dp.belled_details.bell_height, Double)
                    Else .Worksheets("Database").Range(myCol & rowStart + 116).ClearContents
                    End If
                    If Not IsNothing(dp.belled_details.bell_toe_height) Then
                        .Worksheets("Database").Range(myCol & rowStart + 120).Value = CType(dp.belled_details.bell_toe_height, Double)
                    Else .Worksheets("Database").Range(myCol & rowStart + 120).ClearContents
                    End If
                    .Worksheets("Database").Range(myCol & rowStart + 122).Value = CType(dp.belled_details.neglect_top_soil_layer, Boolean)
                    .Worksheets("Database").Range(myCol & rowStart + 123).Value = CType(dp.belled_details.swelling_expansive_soil, Boolean)
                    If Not IsNothing(dp.belled_details.depth_of_expansive_soil) Then
                        .Worksheets("Database").Range(myCol & rowStart + 124).Value = CType(dp.belled_details.depth_of_expansive_soil, Double)
                    Else .Worksheets("Database").Range(myCol & rowStart + 124).ClearContents
                    End If
                    If Not IsNothing(dp.belled_details.expansive_soil_force) Then
                        .Worksheets("Database").Range(myCol & rowStart + 125).Value = CType(dp.belled_details.expansive_soil_force, Double)
                    Else .Worksheets("Database").Range(myCol & rowStart + 125).ClearContents
                    End If

                End If

                'EMBEDDED PIER
                If dp.embedded_pole = True Then

                    If Not IsNothing(dp.embed_details.embedded_id) Then
                        .Worksheets("Database").Range(myCol & rowStart - 2).Value = CType(dp.embed_details.embedded_id, Integer)
                    Else .Worksheets("Database").Range(myCol & rowStart - 2).ClearContents
                    End If

                    .Worksheets("Database").Range(myCol & rowStart + 100).Value = CType(dp.embedded_pole, Boolean)
                    .Worksheets("Database").Range(myCol & rowStart + 101).Value = CType(dp.embed_details.encased_in_concrete, Boolean)
                    If Not IsNothing(dp.embed_details.pole_side_quantity) Then
                        .Worksheets("Database").Range(myCol & rowStart + 102).Value = CType(dp.embed_details.pole_side_quantity, Integer)
                    Else .Worksheets("Database").Range(myCol & rowStart + 102).ClearContents
                    End If
                    If Not IsNothing(dp.embed_details.pole_yield_strength) Then
                        .Worksheets("Database").Range(myCol & rowStart + 103).Value = CType(dp.embed_details.pole_yield_strength, Double)
                    Else .Worksheets("Database").Range(myCol & rowStart + 103).ClearContents
                    End If
                    If Not IsNothing(dp.embed_details.pole_thickness) Then
                        .Worksheets("Database").Range(myCol & rowStart + 104).Value = CType(dp.embed_details.pole_thickness, Double)
                    Else .Worksheets("Database").Range(myCol & rowStart + 104).ClearContents
                    End If
                    .Worksheets("Database").Range(myCol & rowStart + 105).Value = CType(dp.embed_details.embedded_pole_input_type, String)
                    If Not IsNothing(dp.embed_details.pole_diameter_toc) Then
                        .Worksheets("Database").Range(myCol & rowStart + 106).Value = CType(dp.embed_details.pole_diameter_toc, Double)
                    Else .Worksheets("Database").Range(myCol & rowStart + 106).ClearContents
                    End If
                    If Not IsNothing(dp.embed_details.pole_top_diameter) Then
                        .Worksheets("Database").Range(myCol & rowStart + 107).Value = CType(dp.embed_details.pole_top_diameter, Double)
                    Else .Worksheets("Database").Range(myCol & rowStart + 107).ClearContents
                    End If
                    If Not IsNothing(dp.embed_details.pole_bottom_diameter) Then
                        .Worksheets("Database").Range(myCol & rowStart + 108).Value = CType(dp.embed_details.pole_bottom_diameter, Double)
                    Else .Worksheets("Database").Range(myCol & rowStart + 108).ClearContents
                    End If
                    If Not IsNothing(dp.embed_details.pole_section_length) Then
                        .Worksheets("Database").Range(myCol & rowStart + 109).Value = CType(dp.embed_details.pole_section_length, Double)
                    Else .Worksheets("Database").Range(myCol & rowStart + 109).ClearContents
                    End If
                    If Not IsNothing(dp.embed_details.pole_taper_factor) Then
                        .Worksheets("Database").Range(myCol & rowStart + 110).Value = CType(dp.embed_details.pole_taper_factor, Double)
                    Else .Worksheets("Database").Range(myCol & rowStart + 110).ClearContents
                    End If
                    If Not IsNothing(dp.embed_details.pole_bend_radius_override) Then
                        .Worksheets("Database").Range(myCol & rowStart + 111).Value = CType(dp.embed_details.pole_bend_radius_override, Double)
                    Else .Worksheets("Database").Range(myCol & rowStart + 111).ClearContents
                    End If

                End If

                'DRILLED PIER PROFILES
                Dim summaryRowStart As Integer = 10

                For Each dpp As DrilledPierProfile In dp.drilled_pier_profiles
                    'Profile Return
                    If Not IsNothing(dp.local_drilled_pier_id) Then
                        .Worksheets("Profiles (RETURN)").Range("A" & profileRow).Value = CType(dp.local_drilled_pier_id, Integer)
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
                        .Worksheets("SUMMARY").Range("D" & summaryRowStart + CType(dpp.reaction_position, Integer)).Value = CType(dpp.drilled_pier_profile, Integer)
                        If dpp.drilled_pier_profile = dpp.reaction_position Then
                            .Worksheets("SUMMARY").Range("G" & summaryRowStart + CType(dpp.reaction_position, Integer)).Value = False
                        Else
                            .Worksheets("SUMMARY").Range("G" & summaryRowStart + CType(dpp.reaction_position, Integer)).Value = True
                        End If
                    End If
                    If Not IsNothing(dpp.reaction_position) Then
                        .Worksheets("SUMMARY").Range("E" & summaryRowStart + CType(dpp.reaction_position, Integer)).Value = CType(dpp.soil_profile, Integer)
                        If dpp.soil_profile = dpp.reaction_position Then
                            .Worksheets("SUMMARY").Range("H" & summaryRowStart + CType(dpp.reaction_position, Integer)).Value = False
                        Else
                            .Worksheets("SUMMARY").Range("H" & summaryRowStart + CType(dpp.reaction_position, Integer)).Value = True
                        End If
                    End If
                    .Worksheets("SUMMARY").Range("I" & summaryRowStart + CType(dpp.reaction_position, Integer)).Value = False
                    .Worksheets("SUMMARY").Range("J" & summaryRowStart + CType(dpp.reaction_position, Integer)).Value = CType(dpp.profile_id, Integer)

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

                        If Not IsNothing(dpSL.soil_layer_id) Then
                            .Worksheets("Database").Range(mySoilCol & rowStart - 33 + (soilCount - 1)).Value = CType(dpSL.soil_layer_id, Integer)
                        Else .Worksheets("Database").Range(mySoilCol & rowStart - 33 + (soilCount - 1)).ClearContents
                        End If

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




            '~~~~~~~~POPULATE TOOL INPUTS WITH THE FIRST INSTANCE IN TOOL'S LOCAL DATABASE

            Dim firstReaction As String = DrilledPiers(0).drilled_pier_profiles(0).reaction_location 'MRP - error here due to change in profiles. Must reassociate profile with pier. Same with sections

            If firstReaction = "Monopole" Then
                .Worksheets("Foundation Input").Range("TowerType").Value = "Monopole"
            ElseIf firstReaction = "Self Support" Then
                .Worksheets("Foundation Input").Range("TowerType").Value = "Self Support"
            ElseIf firstReaction = "Base" Then
                .Worksheets("Foundation Input").Range("TowerType").Value = "Guyed (Base)"
                .Worksheets("Foundation Input").Range("Location").Value = "Base"
            End If

            Dim firstPierProfile As Integer = DrilledPiers(0).drilled_pier_profiles(0).drilled_pier_profile
            Dim firstSoilProfile As Integer = DrilledPiers(0).drilled_pier_profiles(0).soil_profile

            colCounter = 7

            myCol = GetExcelColumnName(colCounter)

            If firstReaction <> "" Then

                'MATERIAL PROPERTIES
                If DrilledPiers(0).concrete_compressive_strength.HasValue Then
                    .Worksheets("Foundation Input").Range("f\c").Value = CType(DrilledPiers(0).concrete_compressive_strength, Double)
                Else .Worksheets("Foundation Input").Range("f\c").clearcontents
                End If
                If DrilledPiers(0).longitudinal_rebar_yield_strength.HasValue Then
                    .Worksheets("Foundation Input").Range("Fy_rebar").Value = CType(DrilledPiers(0).longitudinal_rebar_yield_strength, Double)
                Else .Worksheets("Foundation Input").Range("Fy_rebar").ClearContents
                End If
                If DrilledPiers(0).tie_yield_strength.HasValue Then
                    .Worksheets("Foundation Input").Range("yield_tie").Value = CType(DrilledPiers(0).tie_yield_strength, Double)
                Else .Worksheets("Foundation Input").Range("yield_tie").ClearContents
                End If
                If DrilledPiers(0).rebar_cage_2_fy_override.HasValue Then
                    .Worksheets("Foundation Input").Range("RebarCage2FyOverride").Value = CType(DrilledPiers(0).rebar_cage_2_fy_override, Double)
                Else .Worksheets("Foundation Input").Range("RebarCage2FyOverride").ClearContents
                End If
                If DrilledPiers(0).rebar_cage_3_fy_override.HasValue Then
                    .Worksheets("Foundation Input").Range("RebarCage3FyOverride").Value = CType(DrilledPiers(0).rebar_cage_3_fy_override, Double)
                Else .Worksheets("Foundation Input").Range("RebarCage3FyOverride").ClearContents
                End If

                'PIER DESIGN DATA (GENERAL)
                If DrilledPiers(0).foundation_depth.HasValue Then
                    .Worksheets("Foundation Input").Range("depth").Value = CType(DrilledPiers(0).foundation_depth, Double)
                Else .Worksheets("Foundation Input").Range("depth").ClearContents
                End If
                If DrilledPiers(0).extension_above_grade.HasValue Then
                    .Worksheets("Foundation Input").Range("ConcreteAboveGrade").Value = CType(DrilledPiers(0).extension_above_grade, Double)
                Else .Worksheets("Foundation Input").Range("ConcreteAboveGrade").ClearContents
                End If
                'groundwater
                If CType(DrilledPiers(0).groundwater_depth, Double) = -1 Then
                    .Worksheets("Foundation Input").Range("GW").Value = "N/A"
                Else .Worksheets("Foundation Input").Range("GW").Value = CType(DrilledPiers(0).groundwater_depth, Double)
                End If
                'soil layers
                If DrilledPiers(0).soil_layer_quantity.HasValue Then
                    .Worksheets("Foundation Input").Range("SoilLayerQty").Value = CType(DrilledPiers(0).soil_layer_quantity, Integer)
                Else .Worksheets("Foundation Input").Range("SoilLayerQty").ClearContents
                End If
                'min steel
                If Not IsNothing(CType(DrilledPiers(0).assume_min_steel, String)) Then
                    .Worksheets("Foundation Input").Range("AssumeMinSteel").Value = CType(DrilledPiers(0).assume_min_steel, String)
                Else .Worksheets("Foundation Input").Range("AssumeMinSteel").ClearContents
                End If

                'PIER DESIGN DATA (SECTIONS)
                Dim secCount As Integer = 1
                Dim secRowStart As Integer = 26

                For Each dpSec As DrilledPierSection In DrilledPiers(0).sections

                    If dpSec.pier_diameter.HasValue Then
                        .Worksheets("Foundation Input").Range("D" & secRowStart).Value = CType(dpSec.pier_diameter, Double)
                    Else .Worksheets("Foundation Input").Range("D" & secRowStart).ClearContents
                    End If
                    If dpSec.clear_cover.HasValue Then
                        .Worksheets("Foundation Input").Range("D" & secRowStart + 3).Value = CType(dpSec.clear_cover, Double)
                    Else .Worksheets("Foundation Input").Range("D" & secRowStart + 3).ClearContents
                    End If
                    If Not IsNothing(CType(dpSec.clear_cover_rebar_cage_option, String)) Then
                        .Worksheets("Foundation Input").Range("B" & secRowStart + 3).Value = CType(dpSec.clear_cover_rebar_cage_option, String)
                    Else .Worksheets("Foundation Input").Range("B" & secRowStart + 3).ClearContents
                    End If
                    If dpSec.tie_size.HasValue Then
                        .Worksheets("Foundation Input").Range("D" & secRowStart + 4).Value = CType(dpSec.tie_size, Integer)
                    Else .Worksheets("Foundation Input").Range("D" & secRowStart + 4).ClearContents
                    End If
                    If dpSec.tie_spacing.HasValue Then
                        .Worksheets("Foundation Input").Range("D" & secRowStart + 5).Value = CType(dpSec.tie_spacing, Double)
                    Else .Worksheets("Foundation Input").Range("D" & secRowStart + 5).ClearContents
                    End If
                    If dpSec.bottom_elevation.HasValue Then
                        .Worksheets("Foundation Input").Range("Depth" & secCount).Value = CType(dpSec.bottom_elevation, Double)
                    Else .Worksheets("Foundation Input").Range("Depth" & secCount).ClearContents
                    End If
                    If dpSec.rho_override.HasValue Then
                        .Worksheets("Foundation Input").Range("rhoOverride" & secCount).Value = CType(dpSec.rho_override, Double)
                    Else .Worksheets("Foundation Input").Range("rhoOverride" & secCount).ClearContents
                    End If

                    'PIER DESIGN DATA (REBAR)
                    Dim rebCount As Integer = 1

                    For Each dpReb As DrilledPierRebar In dpSec.rebar

                        If rebCount = 1 Then
                            If dpReb.longitudinal_rebar_quantity.HasValue Then
                                .Worksheets("Foundation Input").Range("D" & secRowStart + 1).Value = CType(dpReb.longitudinal_rebar_quantity, Integer)
                            Else .Worksheets("Foundation Input").Range("D" & secRowStart + 1).ClearContents
                            End If
                            If dpReb.longitudinal_rebar_size.HasValue Then
                                .Worksheets("Foundation Input").Range("D" & secRowStart + 2).Value = CType(dpReb.longitudinal_rebar_size, Integer)
                            Else .Worksheets("Foundation Input").Range("D" & secRowStart + 2).ClearContents
                            End If
                        End If

                        If rebCount = 2 Then
                            If dpReb.longitudinal_rebar_quantity.HasValue Then
                                .Worksheets("Foundation Input").Range("D" & secRowStart + 6).Value = CType(dpReb.longitudinal_rebar_quantity, Integer)
                            Else .Worksheets("Foundation Input").Range("D" & secRowStart + 6).ClearContents
                            End If
                            If dpReb.longitudinal_rebar_size.HasValue Then
                                .Worksheets("Foundation Input").Range("D" & secRowStart + 7).Value = CType(dpReb.longitudinal_rebar_size, Integer)
                            Else .Worksheets("Foundation Input").Range("D" & secRowStart + 7).ClearContents
                            End If
                            If dpReb.longitudinal_rebar_cage_diameter.HasValue Then
                                .Worksheets("Foundation Input").Range("D" & secRowStart + 8).Value = CType(dpReb.longitudinal_rebar_cage_diameter, Integer)
                            Else .Worksheets("Foundation Input").Range("D" & secRowStart + 8).ClearContents
                            End If
                        End If

                        If rebCount = 3 Then
                            If dpReb.longitudinal_rebar_quantity.HasValue Then
                                .Worksheets("Foundation Input").Range("D" & secRowStart + 10).Value = CType(dpReb.longitudinal_rebar_quantity, Integer)
                            Else .Worksheets("Foundation Input").Range("D" & secRowStart + 10).ClearContents
                            End If
                            If dpReb.longitudinal_rebar_size.HasValue Then
                                .Worksheets("Foundation Input").Range("D" & secRowStart + 11).Value = CType(dpReb.longitudinal_rebar_size, Integer)
                            Else .Worksheets("Foundation Input").Range("D" & secRowStart + 11).ClearContents
                            End If
                            If dpReb.longitudinal_rebar_cage_diameter.HasValue Then
                                .Worksheets("Foundation Input").Range("D" & secRowStart + 12).Value = CType(dpReb.longitudinal_rebar_cage_diameter, Integer)
                            Else .Worksheets("Foundation Input").Range("D" & secRowStart + 12).ClearContents
                            End If
                        End If

                        rebCount += 1

                    Next

                    'populate rebar cage qty (hidden input in tool, typically populated by the Pier Options)
                    .Worksheets("Foundation Input").Range("Rebar" & secCount).Value = rebCount - 1

                    secCount += 1

                    secRowStart += 16

                Next


                'SOIL
                Dim soilRowStart As Integer = 121
                Dim soilCount As Integer = 1

                For Each dpSL As DrilledPierSoilLayer In DrilledPiers(0).soil_layers

                    If dpSL.bottom_depth.HasValue Then
                        .Worksheets("Foundation Input").Range("D" & soilRowStart + soilCount).Value = CType(dpSL.bottom_depth, Double)
                    Else .Worksheets("Foundation Input").Range("D" & soilRowStart + soilCount).ClearContents
                    End If
                    If dpSL.effective_soil_density.HasValue Then
                        .Worksheets("Foundation Input").Range("F" & soilRowStart + soilCount).Value = CType(dpSL.effective_soil_density, Double)
                    Else .Worksheets("Foundation Input").Range("F" & soilRowStart + soilCount).ClearContents
                    End If
                    If dpSL.cohesion.HasValue Then
                        .Worksheets("Foundation Input").Range("H" & soilRowStart + soilCount).Value = CType(dpSL.cohesion, Double)
                    Else .Worksheets("Foundation Input").Range("H" & soilRowStart + soilCount).ClearContents
                    End If
                    If dpSL.friction_angle.HasValue Then
                        .Worksheets("Foundation Input").Range("I" & soilRowStart + soilCount).Value = CType(dpSL.friction_angle, Double)
                    Else .Worksheets("Foundation Input").Range("I" & soilRowStart + soilCount).ClearContents
                    End If
                    If dpSL.skin_friction_override_comp.HasValue Then
                        .Worksheets("Foundation Input").Range("M" & soilRowStart + soilCount).Value = CType(dpSL.skin_friction_override_comp, Double)
                    Else .Worksheets("Foundation Input").Range("M" & soilRowStart + soilCount).ClearContents
                    End If
                    If dpSL.skin_friction_override_uplift.HasValue Then
                        .Worksheets("Foundation Input").Range("N" & soilRowStart + soilCount).Value = CType(dpSL.skin_friction_override_uplift, Double)
                    Else .Worksheets("Foundation Input").Range("N" & soilRowStart + soilCount).ClearContents
                    End If
                    If dpSL.nominal_bearing_capacity.HasValue Then
                        .Worksheets("Foundation Input").Range("O" & soilRowStart + soilCount).Value = CType(dpSL.nominal_bearing_capacity, Double)
                    Else .Worksheets("Foundation Input").Range("O" & soilRowStart + soilCount).ClearContents
                    End If
                    If dpSL.spt_blow_count.HasValue Then
                        .Worksheets("Foundation Input").Range("P" & soilRowStart + soilCount).Value = CType(dpSL.spt_blow_count, Integer)
                    Else .Worksheets("Foundation Input").Range("P" & soilRowStart + soilCount).ClearContents
                    End If

                    soilCount += 1

                Next


                'OPTIONS
                .Worksheets("Foundation Input").Range("EffectiveDepthInput").Value = CType(DrilledPiers(0).rebar_effective_depths, Boolean)
                .Worksheets("Foundation Input").Range("ShearAlongDepth").Value = CType(DrilledPiers(0).check_shear_along_depth, Boolean)
                .Worksheets("Foundation Input").Range("ShearFriction").Value = CType(DrilledPiers(0).utilize_shear_friction_methodology, Boolean)
                .Worksheets("Foundation Input").Range("ShearInputOverride").Value = CType(DrilledPiers(0).shear_override_crit_depth, Boolean)
                If .Worksheets("Foundation Input").Range("ShearInputOverride").Value = CType(DrilledPiers(0).shear_override_crit_depth, Boolean) = True Then
                    If DrilledPiers(0).shear_crit_depth_override_comp.HasValue Then
                        .Worksheets("Foundation Input").Range("ShearCritDepthComp").Value = CType(DrilledPiers(0).shear_crit_depth_override_comp, Double)
                    End If
                    If DrilledPiers(0).shear_crit_depth_override_uplift.HasValue Then
                        .Worksheets("Foundation Input").Range("ShearCritDepthUplift").Value = CType(DrilledPiers(0).shear_crit_depth_override_uplift, Double)
                    End If
                End If


                'BELLED PIER
                .Worksheets("Belled Pier").Range("Belled").Value = CType(DrilledPiers(0).belled_pier, Boolean)
                If DrilledPiers(0).belled_pier = True Then
                    If DrilledPiers(0).belled_details.bottom_diameter_of_bell.HasValue Then
                        .Worksheets("Belled Pier").Range("Dia_Bell").Value = CType(DrilledPiers(0).belled_details.bottom_diameter_of_bell, Double)
                    Else .Worksheets("Belled Pier").Range("Dia_Bell").ClearContents
                    End If
                    If Not IsNothing(CType(DrilledPiers(0).belled_details.bell_input_type, String)) Then
                        .Worksheets("Belled Pier").Range("BellInputType").Value = CType(DrilledPiers(0).belled_details.bell_input_type, String)
                    Else .Worksheets("Belled Pier").Range("BellInputType").ClearContents
                    End If
                    If DrilledPiers(0).belled_details.bell_angle.HasValue Then
                        .Worksheets("Belled Pier").Range("BellAngle").Value = CType(DrilledPiers(0).belled_details.bell_angle, Double)
                    Else .Worksheets("Belled Pier").Range("BellAngle").ClearContents
                    End If
                    If DrilledPiers(0).belled_details.bell_height.HasValue Then
                        .Worksheets("Belled Pier").Range("hbell").Value = CType(DrilledPiers(0).belled_details.bell_height, Double)
                    Else .Worksheets("Belled Pier").Range("hbell").ClearContents
                    End If
                    If DrilledPiers(0).belled_details.bell_toe_height.HasValue Then
                        .Worksheets("Belled Pier").Range("t_bell").Value = CType(DrilledPiers(0).belled_details.bell_toe_height, Double)
                    Else .Worksheets("Belled Pier").Range("t_bell").ClearContents
                    End If
                    .Worksheets("Belled Pier").Range("Neglect_Top").Value = CType(DrilledPiers(0).belled_details.neglect_top_soil_layer, Boolean)
                    .Worksheets("Belled Pier").Range("expansive").Value = CType(DrilledPiers(0).belled_details.expansive_soil_force, Boolean)
                    If DrilledPiers(0).belled_details.depth_of_expansive_soil.HasValue Then
                        .Worksheets("Belled Pier").Range("D_expansive").Value = CType(DrilledPiers(0).belled_details.depth_of_expansive_soil, Double)
                    Else .Worksheets("Belled Pier").Range("D_expansive").ClearContents
                    End If
                    If DrilledPiers(0).belled_details.expansive_soil_force.HasValue Then
                        .Worksheets("Belled Pier").Range("Force_Expansive").Value = CType(DrilledPiers(0).belled_details.expansive_soil_force, Double)
                    Else .Worksheets("Belled Pier").Range("Force_Expansive").ClearContents
                    End If
                End If


                'EMBEDDED POLE
                .Worksheets("Soil Calculations").Range("Embedded").Value = CType(DrilledPiers(0).embedded_pole, Boolean)
                If DrilledPiers(0).embedded_pole = True Then
                    .Worksheets("Soil Calculations").Range("Encased").Value = CType(DrilledPiers(0).embed_details.encased_in_concrete, Boolean)
                    If DrilledPiers(0).embed_details.pole_side_quantity.HasValue Then
                        .Worksheets("Soil Calculations").Range("Sides").Value = CType(DrilledPiers(0).embed_details.pole_side_quantity, Integer)
                    Else .Worksheets("Soil Calculations").Range("Sides").ClearContents
                    End If
                    If DrilledPiers(0).embed_details.pole_yield_strength.HasValue Then
                        .Worksheets("Soil Calculations").Range("Fy").Value = CType(DrilledPiers(0).embed_details.pole_yield_strength, Double)
                    Else .Worksheets("Soil Calculations").Range("Fy").ClearContents
                    End If
                    If DrilledPiers(0).embed_details.pole_thickness.HasValue Then
                        .Worksheets("Soil Calculations").Range("t").Value = CType(DrilledPiers(0).embed_details.pole_thickness, Double)
                    Else .Worksheets("Soil Calculations").Range("t").ClearContents
                    End If
                    If Not IsNothing(CType(DrilledPiers(0).embed_details.embedded_pole_input_type, String)) Then
                        .Worksheets("Soil Calculations").Range("EmbeddedPoleInputType").Value = CType(DrilledPiers(0).embed_details.embedded_pole_input_type, String)
                    Else .Worksheets("Soil Calculations").Range("EmbeddedPoleInputType").ClearContents
                    End If
                    If DrilledPiers(0).embed_details.pole_diameter_toc.HasValue Then
                        .Worksheets("Soil Calculations").Range("dia_grade").Value = CType(DrilledPiers(0).embed_details.pole_diameter_toc, Double)
                    Else .Worksheets("Soil Calculations").Range("dia_grade").ClearContents
                    End If
                    If DrilledPiers(0).embed_details.pole_top_diameter.HasValue Then
                        .Worksheets("Soil Calculations").Range("TopDiameter").Value = CType(DrilledPiers(0).embed_details.pole_top_diameter, Double)
                    Else .Worksheets("Soil Calculations").Range("TopDiameter").ClearContents
                    End If
                    If DrilledPiers(0).embed_details.pole_bottom_diameter.HasValue Then
                        .Worksheets("Soil Calculations").Range("BottomDiameter").Value = CType(DrilledPiers(0).embed_details.pole_bottom_diameter, Double)
                    Else .Worksheets("Soil Calculations").Range("BottomDiameter").ClearContents
                    End If
                    If DrilledPiers(0).embed_details.pole_section_length.HasValue Then
                        .Worksheets("Soil Calculations").Range("LengthOfSection").Value = CType(DrilledPiers(0).embed_details.pole_section_length, Double)
                    Else .Worksheets("Soil Calculations").Range("LengthOfSection").ClearContents
                    End If
                    If DrilledPiers(0).embed_details.pole_taper_factor.HasValue Then
                        .Worksheets("Soil Calculations").Range("taper").Value = CType(DrilledPiers(0).embed_details.pole_taper_factor, Double)
                    Else .Worksheets("Soil Calculations").Range("taper").ClearContents
                    End If
                    If DrilledPiers(0).embed_details.pole_bend_radius_override.HasValue Then
                        .Worksheets("Soil Calculations").Range("bend_user").Value = CType(DrilledPiers(0).embed_details.pole_bend_radius_override, Double)
                    Else .Worksheets("Soil Calculations").Range("bend_user").ClearContents
                    End If
                End If

                'LOCATION
                .Worksheets("Foundation Input").Range("Location").Value = firstReaction

            End If


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
        NewDrilledPierWb.Calculate()
        NewDrilledPierWb.EndUpdate()
        NewDrilledPierWb.SaveDocument(ExcelFilePath, DrilledPierFileType)
    End Sub
#End Region

#Region "SQL Insert Statements"
    Private Function InsertDrilledPierDetail(ByVal dp As DrilledPier) As String
        Dim insertString As String = ""

        'insertString += "@FndID"
        insertString += IIf(IsNothing(dp.local_drilled_pier_id), "Null", "'" & dp.local_drilled_pier_id.ToString & "'")
        insertString += "," & IIf(IsNothing(dp.local_drilled_pier_profile), "Null", "'" & dp.local_drilled_pier_profile.ToString & "'")
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
        insertString += "," & IIf(IsNothing(dp.bearing_type_toggle), "Null", "'" & dp.bearing_type_toggle.ToString & "'")
        insertString += "," & IIf(IsNothing(dp.tool_version), "Null", "'" & dp.tool_version.ToString & "'")
        insertString += "," & IIf(IsNothing(dp.modified), "Null", "'" & dp.modified.ToString & "'")

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
        insertString += "," & IIf(IsNothing(dpsec.local_section_id), "Null", "'" & dpsec.local_section_id.ToString & "'")
        insertString += "," & IIf(IsNothing(dpsec.rho_override), "Null", "'" & dpsec.rho_override.ToString & "'")

        Return insertString
    End Function

    Private Function InsertDrilledPierRebar(ByVal dpreb As DrilledPierRebar) As String
        Dim insertString As String = ""

        insertString += "@SecID"
        insertString += "," & IIf(IsNothing(dpreb.longitudinal_rebar_quantity), "Null", "'" & dpreb.longitudinal_rebar_quantity.ToString & "'")
        insertString += "," & IIf(IsNothing(dpreb.longitudinal_rebar_size), "Null", "'" & dpreb.longitudinal_rebar_size.ToString & "'")
        insertString += "," & IIf(IsNothing(dpreb.longitudinal_rebar_cage_diameter), "Null", "'" & dpreb.longitudinal_rebar_cage_diameter.ToString & "'")
        insertString += "," & IIf(IsNothing(dpreb.local_rebar_id), "Null", "'" & dpreb.local_rebar_id.ToString & "'")

        Return insertString
    End Function
    Private Function InsertDrilledPierProfile(ByVal dpp As DrilledPierProfile) As String
        Dim insertString As String = ""

        insertString += "@DpID"
        insertString += "," & IIf(IsNothing(dpp.reaction_position), "Null", "'" & dpp.reaction_position.ToString & "'")
        insertString += "," & IIf(IsNothing(dpp.reaction_location), "Null", "'" & dpp.reaction_location.ToString & "'")
        insertString += "," & IIf(IsNothing(dpp.drilled_pier_profile), "Null", "'" & dpp.drilled_pier_profile.ToString & "'")
        insertString += "," & IIf(IsNothing(dpp.soil_profile), "Null", "'" & dpp.soil_profile.ToString & "'")

        Return insertString
    End Function
#End Region

    '#Region "SQL Update Statements"
    '    Private Function UpdateDrilledPierDetail(ByVal dp As DrilledPier) As String
    '        Dim updateString As String = ""

    '        updateString += "UPDATE drilled_pier_details SET "
    '        updateString += "foundation_depth=" & IIf(IsNothing(dp.foundation_depth), "Null", "'" & dp.foundation_depth.ToString & "'")
    '        updateString += ", extension_above_grade=" & IIf(IsNothing(dp.extension_above_grade), "Null", "'" & dp.extension_above_grade.ToString & "'")
    '        updateString += ", groundwater_depth=" & IIf(IsNothing(dp.groundwater_depth), "Null", "'" & dp.groundwater_depth.ToString & "'")
    '        updateString += ", assume_min_steel=" & IIf(IsNothing(dp.assume_min_steel), "Null", "'" & dp.assume_min_steel.ToString & "'")
    '        updateString += ", check_shear_along_depth=" & IIf(IsNothing(dp.check_shear_along_depth), "Null", "'" & dp.check_shear_along_depth.ToString & "'")
    '        updateString += ", utilize_shear_friction_methodology=" & IIf(IsNothing(dp.utilize_shear_friction_methodology), "Null", "'" & dp.utilize_shear_friction_methodology.ToString & "'")
    '        updateString += ", embedded_pole=" & IIf(IsNothing(dp.embedded_pole), "Null", "'" & dp.embedded_pole.ToString & "'")
    '        updateString += ", belled_pier=" & IIf(IsNothing(dp.belled_pier), "Null", "'" & dp.belled_pier.ToString & "'")
    '        updateString += ", soil_layer_quantity=" & IIf(IsNothing(dp.soil_layer_quantity), "Null", "'" & dp.soil_layer_quantity.ToString & "'")
    '        updateString += ", concrete_compressive_strength=" & IIf(IsNothing(dp.concrete_compressive_strength), "Null", "'" & dp.concrete_compressive_strength.ToString & "'")
    '        updateString += ", tie_yield_strength=" & IIf(IsNothing("'" & dp.tie_yield_strength), "Null", "'" & dp.tie_yield_strength.ToString & "'")
    '        updateString += ", longitudinal_rebar_yield_strength=" & IIf(IsNothing(dp.longitudinal_rebar_yield_strength), "Null", "'" & dp.longitudinal_rebar_yield_strength.ToString & "'")
    '        updateString += ", rebar_effective_depths=" & IIf(IsNothing(dp.rebar_effective_depths), "Null", "'" & dp.rebar_effective_depths.ToString & "'")
    '        updateString += ", rebar_cage_2_fy_override=" & IIf(IsNothing(dp.rebar_cage_2_fy_override), "Null", "'" & dp.rebar_cage_2_fy_override.ToString & "'")
    '        updateString += ", rebar_cage_3_fy_override=" & IIf(IsNothing(dp.rebar_cage_3_fy_override), "Null", "'" & dp.rebar_cage_3_fy_override.ToString & "'")
    '        updateString += ", shear_override_crit_depth=" & IIf(IsNothing(dp.shear_override_crit_depth), "Null", "'" & dp.shear_override_crit_depth.ToString & "'")
    '        updateString += ", shear_crit_depth_override_comp=" & IIf(IsNothing(dp.shear_crit_depth_override_comp), "Null", "'" & dp.shear_crit_depth_override_comp.ToString & "'")
    '        updateString += ", shear_crit_depth_override_uplift=" & IIf(IsNothing(dp.shear_crit_depth_override_uplift), "Null", "'" & dp.shear_crit_depth_override_uplift.ToString & "'")
    '        updateString += ", local_drilled_pier_id=" & IIf(IsNothing(dp.local_drilled_pier_id), "Null", "'" & dp.local_drilled_pier_id.ToString & "'")
    '        updateString += ", bearing_type_toggle=" & IIf(IsNothing(dp.bearing_type_toggle), "Null", "'" & dp.bearing_type_toggle.ToString & "'")
    '        updateString += ", modified=" & IIf(IsNothing(dp.modified), "Null", "'" & dp.modified.ToString & "'")
    '        updateString += ", local_drilled_pier_profile=" & IIf(IsNothing(dp.local_drilled_pier_profile), "Null", "'" & dp.local_drilled_pier_profile.ToString & "'")
    '        updateString += " WHERE ID=" & dp.pier_id & vbNewLine

    '        Return updateString
    '    End Function

    '    Private Function UpdateDrilledPierBell(ByVal bp As DrilledPierBelledPier) As String
    '        Dim updateString As String = ""

    '        updateString += "UPDATE belled_pier_details SET "
    '        updateString += "belled_pier_option=" & IIf(IsNothing(bp.belled_pier_option), "Null", "'" & bp.belled_pier_option.ToString & "'")
    '        updateString += ", bottom_diameter_of_bell=" & IIf(IsNothing(bp.bottom_diameter_of_bell), "Null", "'" & bp.bottom_diameter_of_bell.ToString & "'")
    '        updateString += ", bell_input_type=" & IIf(IsNothing(bp.bell_input_type), "Null", "'" & bp.bell_input_type.ToString & "'")
    '        updateString += ", bell_angle=" & IIf(IsNothing(bp.bell_angle), "Null", "'" & bp.bell_angle.ToString & "'")
    '        updateString += ", bell_height=" & IIf(IsNothing(bp.bell_height), "Null", "'" & bp.bell_height.ToString & "'")
    '        updateString += ", bell_toe_height=" & IIf(IsNothing(bp.bell_toe_height), "Null", "'" & bp.bell_toe_height.ToString & "'")
    '        updateString += ", neglect_top_soil_layer=" & IIf(IsNothing(bp.neglect_top_soil_layer), "Null", "'" & bp.neglect_top_soil_layer.ToString & "'")
    '        updateString += ", swelling_expansive_soil=" & IIf(IsNothing(bp.swelling_expansive_soil), "Null", "'" & bp.swelling_expansive_soil.ToString & "'")
    '        updateString += ", depth_of_expansive_soil=" & IIf(IsNothing(bp.depth_of_expansive_soil), "Null", "'" & bp.depth_of_expansive_soil.ToString & "'")
    '        updateString += ", expansive_soil_force=" & IIf(IsNothing(bp.expansive_soil_force), "Null", "'" & bp.expansive_soil_force.ToString & "'")
    '        updateString += " WHERE ID=" & bp.belled_pier_id & vbNewLine

    '        Return updateString
    '    End Function

    '    Private Function UpdateDrilledPierEmbed(ByVal ep As DrilledPierEmbeddedPier) As String
    '        Dim updateString As String = ""

    '        updateString += "UPDATE embedded_pole_details SET "
    '        updateString += "embedded_pole_option=" & IIf(IsNothing(ep.embedded_pole_option), "Null", "'" & ep.embedded_pole_option.ToString & "'")
    '        updateString += ", encased_in_concrete=" & IIf(IsNothing(ep.encased_in_concrete), "Null", "'" & ep.encased_in_concrete.ToString & "'")
    '        updateString += ", pole_side_quantity=" & IIf(IsNothing(ep.pole_side_quantity), "Null", "'" & ep.pole_side_quantity.ToString & "'")
    '        updateString += ", pole_yield_strength=" & IIf(IsNothing(ep.pole_yield_strength), "Null", "'" & ep.pole_yield_strength.ToString & "'")
    '        updateString += ", pole_thickness=" & IIf(IsNothing(ep.pole_thickness), "Null", "'" & ep.pole_thickness.ToString & "'")
    '        updateString += ", embedded_pole_input_type=" & IIf(IsNothing(ep.embedded_pole_input_type), "Null", "'" & ep.embedded_pole_input_type.ToString & "'")
    '        updateString += ", pole_diameter_toc=" & IIf(IsNothing(ep.pole_diameter_toc), "Null", "'" & ep.pole_diameter_toc.ToString & "'")
    '        updateString += ", pole_top_diameter=" & IIf(IsNothing(ep.pole_top_diameter), "Null", "'" & ep.pole_top_diameter.ToString & "'")
    '        updateString += ", pole_bottom_diameter=" & IIf(IsNothing(ep.pole_bottom_diameter), "Null", "'" & ep.pole_bottom_diameter.ToString & "'")
    '        updateString += ", pole_section_length=" & IIf(IsNothing(ep.pole_section_length), "Null", "'" & ep.pole_section_length.ToString & "'")
    '        updateString += ", pole_taper_factor=" & IIf(IsNothing(ep.pole_taper_factor), "Null", "'" & ep.pole_taper_factor.ToString & "'")
    '        updateString += ", pole_bend_radius_override=" & IIf(IsNothing(ep.pole_bend_radius_override), "Null", "'" & ep.pole_bend_radius_override.ToString & "'")
    '        updateString += " WHERE ID=" & ep.embedded_id & vbNewLine

    '        Return updateString
    '    End Function

    '    Private Function UpdateDrilledPierSoilLayer(ByVal dpsl As DrilledPierSoilLayer) As String
    '        Dim updateString As String = ""

    '        updateString += "UPDATE drilled_pier_soil_layer SET "
    '        updateString += "bottom_depth=" & IIf(IsNothing(dpsl.bottom_depth), "Null", "'" & dpsl.bottom_depth.ToString & "'")
    '        updateString += ", effective_soil_density=" & IIf(IsNothing(dpsl.effective_soil_density), "Null", "'" & dpsl.effective_soil_density.ToString & "'")
    '        updateString += ", cohesion=" & IIf(IsNothing(dpsl.cohesion), "Null", "'" & dpsl.cohesion.ToString & "'")
    '        updateString += ", friction_angle=" & IIf(IsNothing(dpsl.friction_angle), "Null", "'" & dpsl.friction_angle.ToString & "'")
    '        updateString += ", skin_friction_override_comp=" & IIf(IsNothing(dpsl.skin_friction_override_comp), "Null", "'" & dpsl.skin_friction_override_comp.ToString & "'")
    '        updateString += ", skin_friction_override_uplift=" & IIf(IsNothing(dpsl.skin_friction_override_uplift), "Null", "'" & dpsl.skin_friction_override_uplift.ToString & "'")
    '        updateString += ", nominal_bearing_capacity=" & IIf(IsNothing(dpsl.nominal_bearing_capacity), "Null", "'" & dpsl.nominal_bearing_capacity.ToString & "'")
    '        updateString += ", spt_blow_count=" & IIf(IsNothing(dpsl.spt_blow_count), "Null", "'" & dpsl.spt_blow_count.ToString & "'")
    '        updateString += ", local_soil_layer_id=" & IIf(IsNothing(dpsl.local_soil_layer_id), "Null", "'" & dpsl.local_soil_layer_id.ToString & "'")
    '        updateString += " WHERE ID=" & dpsl.soil_layer_id & vbNewLine

    '        Return updateString
    '    End Function

    '    Private Function UpdateDrilledPierSection(ByVal dpsec As DrilledPierSection) As String
    '        Dim updateString As String = ""

    '        updateString += "UPDATE drilled_pier_section SET "
    '        updateString += "pier_diameter=" & IIf(IsNothing(dpsec.pier_diameter), "Null", "'" & dpsec.pier_diameter.ToString & "'")
    '        updateString += ", clear_cover=" & IIf(IsNothing(dpsec.clear_cover), "Null", "'" & dpsec.clear_cover.ToString & "'")
    '        updateString += ", clear_cover_rebar_cage_option=" & IIf(IsNothing(dpsec.clear_cover_rebar_cage_option), "Null", "'" & dpsec.clear_cover_rebar_cage_option.ToString & "'")
    '        updateString += ", tie_size=" & IIf(IsNothing(dpsec.tie_size), "Null", "'" & dpsec.tie_size.ToString & "'")
    '        updateString += ", tie_spacing=" & IIf(IsNothing(dpsec.tie_spacing), "Null", "'" & dpsec.tie_spacing.ToString & "'")
    '        updateString += ", bottom_elevation=" & IIf(IsNothing(dpsec.bottom_elevation), "Null", "'" & dpsec.bottom_elevation.ToString & "'")
    '        updateString += ", local_section_id=" & IIf(IsNothing(dpsec.local_section_id), "Null", "'" & dpsec.local_section_id.ToString & "'")
    '        updateString += ", local_drilled_pier_id=" & IIf(IsNothing(dpsec.rho_override), "Null", "'" & dpsec.rho_override.ToString & "'")
    '        updateString += " WHERE ID=" & dpsec.section_id & vbNewLine

    '        Return updateString
    '    End Function

    '    Private Function UpdateDrilledPierRebar(ByVal dpreb As DrilledPierRebar) As String
    '        Dim updateString As String = ""

    '        updateString += "UPDATE drilled_pier_rebar SET "
    '        updateString += "longitudinal_rebar_quantity=" & IIf(IsNothing(dpreb.longitudinal_rebar_quantity), "Null", "'" & dpreb.longitudinal_rebar_quantity.ToString & "'")
    '        updateString += ", longitudinal_rebar_size=" & IIf(IsNothing(dpreb.longitudinal_rebar_size), "Null", "'" & dpreb.longitudinal_rebar_size.ToString & "'")
    '        updateString += ", longitudinal_rebar_cage_diameter=" & IIf(IsNothing(dpreb.longitudinal_rebar_cage_diameter), "Null", "'" & dpreb.longitudinal_rebar_cage_diameter.ToString & "'")
    '        updateString += ", local_rebar_id=" & IIf(IsNothing(dpreb.local_rebar_id), "Null", "'" & dpreb.local_rebar_id.ToString & "'")
    '        updateString += " WHERE ID=" & dpreb.rebar_id & vbNewLine

    '        Return updateString
    '    End Function

    '    Private Function UpdateDrilledPierProfile(ByVal dpp As DrilledPierProfile) As String
    '        Dim updateString As String = ""

    '        updateString += "UPDATE drilled_pier_profile SET "
    '        updateString += ", reaction_position=" & IIf(IsNothing(dpp.reaction_position), "Null", "'" & dpp.reaction_position.ToString & "'")
    '        updateString += ", reaction_location=" & IIf(IsNothing(dpp.reaction_location), "Null", "'" & dpp.reaction_location.ToString & "'")
    '        updateString += ", drilled_pier_profile=" & IIf(IsNothing(dpp.drilled_pier_profile), "Null", "'" & dpp.drilled_pier_profile.ToString & "'")
    '        updateString += ", soil_profile=" & IIf(IsNothing(dpp.soil_profile), "Null", "'" & dpp.soil_profile.ToString & "'")
    '        updateString += " WHERE ID=" & dpp.profile_id & vbNewLine

    '        Return updateString
    '    End Function
    '#End Region

#Region "General"
    Public Sub Clear()
        ExcelFilePath = ""
        DrilledPiers.Clear()

        'Remove all datatables from the main dataset
        For Each item As EXCELDTParameter In DrilledPierExcelDTParameters()
            Try
                ds.Tables.Remove(item.xlsDatatable)
            Catch ex As Exception
            End Try
        Next

        For Each item As SQLParameter In DrilledPierSQLDataTables()
            Try
                ds.Tables.Remove(item.sqlDatatable)
            Catch ex As Exception
            End Try
        Next
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

        MyParameters.Add(New EXCELDTParameter("Drilled Pier General Details EXCEL", "A2:V1000", "Details (ENTER)"))
        MyParameters.Add(New EXCELDTParameter("Drilled Pier Section EXCEL", "A2:K1000", "Sections (ENTER)"))
        MyParameters.Add(New EXCELDTParameter("Drilled Pier Rebar EXCEL", "A2:I1000", "Rebar (ENTER)"))
        MyParameters.Add(New EXCELDTParameter("Drilled Pier Soil EXCEL", "A2:L1502", "Soil Layers (ENTER)")) 'use range of 1000 to be safe that multiple generations of EDS values are brought in. This range need to go to 1500 values to match the tool's limit
        MyParameters.Add(New EXCELDTParameter("Belled Details EXCEL", "A2:M1000", "Belled (ENTER)"))
        MyParameters.Add(New EXCELDTParameter("Embedded Details EXCEL", "A2:O1000", "Embedded (ENTER)"))
        MyParameters.Add(New EXCELDTParameter("Drilled Pier Profiles EXCEL", "A2:G1000", "Profiles (ENTER)"))

        Return MyParameters
    End Function
#End Region

#Region "Check Changes"
    Private changeDt As New DataTable
    Private changeList As New List(Of AnalysisChanges)
    Function CheckChanges(ByVal xlDrilledPier As DrilledPier, ByVal sqlDrilledPier As DrilledPier) As Boolean
        Dim changesMade As Boolean = False

        changeDt.Columns.Add("Variable", Type.GetType("System.String"))
        changeDt.Columns.Add("New Value", Type.GetType("System.String"))
        changeDt.Columns.Add("Previuos Value", Type.GetType("System.String"))
        changeDt.Columns.Add("WO", Type.GetType("System.String"))

        ''Check Details
        'If Check1Change(xlGuyedAnchorBlock.anchor_depth, sqlGuyedAnchorBlock.anchor_depth, 1, "Anchor_Depth") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.anchor_width, sqlGuyedAnchorBlock.anchor_width, 1, "Anchor_Width") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.anchor_thickness, sqlGuyedAnchorBlock.anchor_thickness, 1, "Anchor_Thickness") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.anchor_length, sqlGuyedAnchorBlock.anchor_length, 1, "Anchor_Length") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.anchor_toe_width, sqlGuyedAnchorBlock.anchor_toe_width, 1, "Anchor_Toe_Width") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.anchor_top_rebar_size, sqlGuyedAnchorBlock.anchor_top_rebar_size, 1, "Anchor_Top_Rebar_Size") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.anchor_top_rebar_quantity, sqlGuyedAnchorBlock.anchor_top_rebar_quantity, 1, "Anchor_Top_Rebar_Quantity") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.anchor_front_rebar_size, sqlGuyedAnchorBlock.anchor_front_rebar_size, 1, "Anchor_Front_Rebar_Size") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.anchor_front_rebar_quantity, sqlGuyedAnchorBlock.anchor_front_rebar_quantity, 1, "Anchor_Front_Rebar_Quantity") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.anchor_stirrup_size, sqlGuyedAnchorBlock.anchor_stirrup_size, 1, "Anchor_Stirrup_Size") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.anchor_shaft_diameter, sqlGuyedAnchorBlock.anchor_shaft_diameter, 1, "Anchor_Shaft_Diameter") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.anchor_shaft_quantity, sqlGuyedAnchorBlock.anchor_shaft_quantity, 1, "Anchor_Shaft_Quantity") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.anchor_shaft_area_override, sqlGuyedAnchorBlock.anchor_shaft_area_override, 1, "Anchor_Shaft_Area_Override") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.anchor_shaft_shear_lag_factor, sqlGuyedAnchorBlock.anchor_shaft_shear_lag_factor, 1, "Anchor_Shaft_Shear_Lag_Factor") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.concrete_compressive_strength, sqlGuyedAnchorBlock.concrete_compressive_strength, 1, "Concrete_Compressive_Strength") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.clear_cover, sqlGuyedAnchorBlock.clear_cover, 1, "Clear_Cover") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.anchor_shaft_yield_strength, sqlGuyedAnchorBlock.anchor_shaft_yield_strength, 1, "Anchor_Shaft_Yield_Strength") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.anchor_shaft_ultimate_strength, sqlGuyedAnchorBlock.anchor_shaft_ultimate_strength, 1, "Anchor_Shaft_Ultimate_Strength") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.neglect_depth, sqlGuyedAnchorBlock.neglect_depth, 1, "Neglect_Depth") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.groundwater_depth, sqlGuyedAnchorBlock.groundwater_depth, 1, "Groundwater_Depth") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.soil_layer_quantity, sqlGuyedAnchorBlock.soil_layer_quantity, 1, "Soil_Layer_Quantity") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.tool_version, sqlGuyedAnchorBlock.tool_version, 1, "Tool_Version") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.anchor_shaft_section, sqlGuyedAnchorBlock.anchor_shaft_section, 1, "Anchor_Shaft_Section") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.anchor_rebar_grade, sqlGuyedAnchorBlock.anchor_rebar_grade, 1, "Anchor_Rebar_Grade") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.anchor_shaft_known, sqlGuyedAnchorBlock.anchor_shaft_known, 1, "Anchor_Shaft_Known") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.basic_soil_check, sqlGuyedAnchorBlock.basic_soil_check, 1, "Basic_Soil_Check") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.structural_check, sqlGuyedAnchorBlock.structural_check, 1, "Structural_Check") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.rebar_known, sqlGuyedAnchorBlock.rebar_known, 1, "Rebar_Known") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.local_anchor_id, sqlGuyedAnchorBlock.local_anchor_id, 1, "Local_Anchor_Id") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.local_anchor_profile, sqlGuyedAnchorBlock.local_anchor_profile, 1, "Local_Anchor_Profile") Then changesMade = True

        ''Check Soil Layer
        ''If xlGuyedAnchorBlock.soil_layers.Count <> sqlGuyedAnchorBlock.soil_layers.Count Then changesMade = True 'If want to bypass all the checks below

        'For Each gabsl As GuyedAnchorBlockSoilLayer In xlGuyedAnchorBlock.soil_layers
        '    For Each sqlgabsl As GuyedAnchorBlockSoilLayer In sqlGuyedAnchorBlock.soil_layers

        '        If gabsl.soil_layer_id = sqlgabsl.soil_layer_id Then
        '            If Check1Change(gabsl.bottom_depth, sqlgabsl.bottom_depth, 1, "Bottom_Depth" & gabsl.soil_layer_id.ToString) Then changesMade = True
        '            If Check1Change(gabsl.effective_soil_density, sqlgabsl.effective_soil_density, 1, "Effective_Soil_Density" & gabsl.soil_layer_id.ToString) Then changesMade = True
        '            If Check1Change(gabsl.cohesion, sqlgabsl.cohesion, 1, "Cohesion" & gabsl.soil_layer_id.ToString) Then changesMade = True
        '            If Check1Change(gabsl.friction_angle, sqlgabsl.friction_angle, 1, "Friction_Angle" & gabsl.soil_layer_id.ToString) Then changesMade = True
        '            If Check1Change(gabsl.skin_friction_override_uplift, sqlgabsl.skin_friction_override_uplift, 1, "Ultimate_Skin_Friction_Override_Uplift" & gabsl.soil_layer_id.ToString) Then changesMade = True
        '            If Check1Change(gabsl.spt_blow_count, sqlgabsl.spt_blow_count, 1, "spt_blow_count" & gabsl.soil_layer_id.ToString) Then changesMade = True
        '            If Check1Change(gabsl.local_soil_layer_id, sqlgabsl.local_soil_layer_id, 1, "local_soil_layer_id" & gabsl.soil_layer_id.ToString) Then changesMade = True
        '            If Check1Change(gabsl.local_soil_profile, sqlgabsl.local_soil_profile, 1, "local_soil_profile" & gabsl.soil_layer_id.ToString) Then changesMade = True
        '            Exit For
        '        End If

        '        If gabsl.soil_layer_id = 0 Then 'accounts for inserting new rows. additional rows won't have an ID associated to them. 
        '            If Check1Change(gabsl.bottom_depth, Nothing, 1, "Bottom_Depth" & gabsl.soil_layer_id.ToString) Then changesMade = True
        '            If Check1Change(gabsl.effective_soil_density, Nothing, 1, "Effective_Soil_Density" & gabsl.soil_layer_id.ToString) Then changesMade = True
        '            If Check1Change(gabsl.cohesion, Nothing, 1, "Cohesion" & gabsl.soil_layer_id.ToString) Then changesMade = True
        '            If Check1Change(gabsl.friction_angle, Nothing, 1, "Friction_Angle" & gabsl.soil_layer_id.ToString) Then changesMade = True
        '            If Check1Change(gabsl.skin_friction_override_uplift, Nothing, 1, "Ultimate_Skin_Friction_Override_Uplift" & gabsl.soil_layer_id.ToString) Then changesMade = True
        '            If Check1Change(gabsl.spt_blow_count, Nothing, 1, "spt_blow_count" & gabsl.soil_layer_id.ToString) Then changesMade = True
        '            If Check1Change(gabsl.local_soil_layer_id, Nothing, 1, "local_soil_layer_id" & gabsl.soil_layer_id.ToString) Then changesMade = True
        '            If Check1Change(gabsl.local_soil_profile, Nothing, 1, "local_soil_profile" & gabsl.soil_layer_id.ToString) Then changesMade = True
        '            Exit For
        '        End If

        '    Next
        'Next

        ''Guyed Anchor Block Profiles
        'For Each gabp As GuyedAnchorBlockProfile In xlGuyedAnchorBlock.anchor_profiles
        '    For Each sqlgabp As GuyedAnchorBlockProfile In sqlGuyedAnchorBlock.anchor_profiles

        '        If gabp.ID = sqlgabp.ID Then
        '            If Check1Change(gabp.reaction_location, sqlgabp.reaction_location, 1, "reaction_location" & gabp.ID.ToString) Then changesMade = True
        '            If Check1Change(gabp.anchor_profile, sqlgabp.anchor_profile, 1, "anchor_profile" & gabp.ID.ToString) Then changesMade = True
        '            If Check1Change(gabp.soil_profile, sqlgabp.soil_profile, 1, "soil_profile" & gabp.ID.ToString) Then changesMade = True
        '            If Check1Change(gabp.local_anchor_id, sqlgabp.local_anchor_id, 1, "local_anchor_id" & gabp.ID.ToString) Then changesMade = True
        '            Exit For
        '        End If

        '        If gabp.ID = 0 Then 'accounts for inserting new rows. additional rows won't have an ID associated to them.
        '            If Check1Change(gabp.reaction_location, Nothing, 1, "reaction_location" & gabp.ID.ToString) Then changesMade = True
        '            If Check1Change(gabp.anchor_profile, Nothing, 1, "anchor_profile" & gabp.ID.ToString) Then changesMade = True
        '            If Check1Change(gabp.soil_profile, Nothing, 1, "soil_profile" & gabp.ID.ToString) Then changesMade = True
        '            If Check1Change(gabp.local_anchor_id, Nothing, 1, "local_anchor_id" & gabp.ID.ToString) Then changesMade = True
        '            Exit For
        '        End If

        '    Next
        'Next

        CreateChangeSummary(changeDt) 'possible alternative to listing change summary
        Return changesMade

    End Function

    Function CreateChangeSummary(ByVal changeDt As DataTable) As String
        'Sub CreateChangeSummary(ByVal changeDt As DataTable)
        'Create your string based on data in the datatable
        Dim summary As String
        Dim counter As Integer = 0

        For Each chng As AnalysisChanges In changeList
            If counter = 0 Then
                summary += chng.Name & " = " & chng.NewValue & " | Previously: " & chng.PreviousValue
            Else
                summary += vbNewLine & chng.Name & " = " & chng.NewValue & " | Previously: " & chng.PreviousValue
            End If

            counter += 1
        Next

        'write to text file
        'End Sub
    End Function

    Function Check1Change(ByVal newValue As Object, ByVal oldvalue As Object, ByVal tolerance As Double, ByVal variable As String) As Boolean
        If newValue <> oldvalue Then
            changeDt.Rows.Add(variable, newValue, oldvalue, CurWO) 'Need to determine what we want to store in this datatable or list (Foundation Type, Foundation ID)?
            changeList.Add(New AnalysisChanges(oldvalue, newValue, variable, "Drilled Pier Foundations"))
            Return True
        ElseIf Not IsNothing(newValue) And IsNothing(oldvalue) Then 'accounts for when new rows are added. New rows from excel=0 where sql=nothing
            changeDt.Rows.Add(variable, newValue, oldvalue, CurWO) 'Need to determine what we want to store in this datatable or list (Foundation Type, Foundation ID)?
            changeList.Add(New AnalysisChanges(oldvalue, newValue, variable, "Drilled Pier Foundations"))
            Return True
        ElseIf IsNothing(newValue) And Not IsNothing(oldvalue) Then 'accounts for when rows are removed. Rows from excel=nothing where sql=value
            changeDt.Rows.Add(variable, newValue, oldvalue, CurWO) 'Need to determine what we want to store in this datatable or list (Foundation Type, Foundation ID)?
            changeList.Add(New AnalysisChanges(oldvalue, newValue, variable, "Drilled Pier Foundations"))
            Return True
        End If
    End Function
#End Region

End Class