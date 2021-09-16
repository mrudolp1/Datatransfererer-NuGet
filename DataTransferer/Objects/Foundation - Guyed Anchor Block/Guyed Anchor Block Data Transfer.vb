Option Strict Off

Imports DevExpress.Spreadsheet
Imports System.Security.Principal
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop

Partial Public Class DataTransfererGuyedAnchorBlock

#Region "Define"
    Private NewGuyedAnchorBlockWb As New Workbook
    Private prop_ExcelFilePath As String

    Public Property GuyedAnchorBlocks As New List(Of GuyedAnchorBlock)
    Private Property GuyedAnchorBlockTemplatePath As String = "C:\Users\" & Environment.UserName & "\Desktop\Guyed Anchor Block Foundation (4.1.0) - TEMPLATE - 9-9-2021.xlsm"
    Private Property GuyedAnchorBlockFileType As DocumentFormat = DocumentFormat.Xlsm

    Public Property gabDB As String
    Public Property gabID As WindowsIdentity

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
        ds = MyDataSet
        gabID = LogOnUser
        gabDB = ActiveDatabase
        'BUNumber = BU 'Need to turn back on when connecting to dashboard. Turned off for testing.
        'STR_ID = Strucutre_ID 'Need to turn back on when connecting to dashboard. Turned off for testing.
    End Sub
#End Region

#Region "Load Data"
    Public Function LoadFromEDS() As Boolean
        Dim refid As Integer

        Dim GuyedAnchorBlockLoader As String

        'Load data to get pier and pad details data for the existing structure model
        For Each item As SQLParameter In GuyedAnchorBlockSQLDataTables()
            GuyedAnchorBlockLoader = QueryBuilderFromFile(queryPath & "Guyed Anchor Block\" & item.sqlQuery).Replace("[EXISTING MODEL]", GetExistingModelQuery())
            DoDaSQL.sqlLoader(GuyedAnchorBlockLoader, item.sqlDatatable, ds, gabDB, gabID, "0")
            'If ds.Tables(item.sqlDatatable).Rows.Count = 0 Then Return False 'This may need adjusted since some tables can be empty
        Next

        'Custom Section to transfer data for the tool. Needs to be adjusted for each tool.
        For Each GuyedAnchorBlockDataRow As DataRow In ds.Tables("Guyed Anchor Block General Details SQL").Rows
            refid = CType(GuyedAnchorBlockDataRow.Item("ID"), Integer)

            GuyedAnchorBlocks.Add(New GuyedAnchorBlock(GuyedAnchorBlockDataRow, refid))
        Next

        Return True
    End Function 'Create Guyed Anchor Block objects based on what is saved in EDS

    Public Sub LoadFromExcel()
        Dim refID As Integer
        Dim refCol As String

        For Each item As EXCELDTParameter In GuyedAnchorBlockExcelDTParameters()
            'Get tables from excel file 
            ds.Tables.Add(ExcelDatasourceToDataTable(GetExcelDataSource(ExcelFilePath, item.xlsSheet, item.xlsRange), item.xlsDatatable))
        Next

        'Custom Section to transfer data for the tool. Needs to be adjusted for each tool.
        For Each GuyedAnchorBlockDataRow As DataRow In ds.Tables("Guyed Anchor Block General Details EXCEL").Rows

            refCol = "local_anchor_id"
            refID = CType(GuyedAnchorBlockDataRow.Item(refCol), Integer)

            GuyedAnchorBlocks.Add(New GuyedAnchorBlock(GuyedAnchorBlockDataRow, refID, refCol))
        Next
    End Sub 'Create Guyed Anchor Block  objects based on what is coming from the excel file
#End Region

#Region "Save Data"
    Public Sub SaveToEDS()
        Dim firstOne As Boolean = True
        Dim mySoils As String = ""
        Dim myProfiles As String = ""

        For Each gab As GuyedAnchorBlock In GuyedAnchorBlocks
            Dim GuyedAnchorBlockSaver As String = QueryBuilderFromFile(queryPath & "Guyed Anchor Block\Guyed Anchor Block (IN_UP).sql")

            GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("[BU NUMBER]", BUNumber)
            GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("[STRUCTURE ID]", STR_ID)
            GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("[FOUNDATION TYPE]", "Guyed Anchor Block")
            If gab.anchor_id = 0 Or IsDBNull(gab.anchor_id) Then
                GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("'[GUYED ANCHOR BLOCK ID]'", "NULL")
            Else
                GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("'[GUYED ANCHOR BLOCK ID]'", gab.anchor_id.ToString)
            End If
            GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("[INSERT ALL GUYED ANCHOR BLOCK DETAILS]", InsertGuyedAnchorBlockDetail(gab))

            If gab.anchor_id = 0 Or IsDBNull(gab.anchor_id) Then
                For Each gabsl As GuyedAnchorBlockSoilLayer In gab.soil_layers
                    Dim tempSoilLayer As String = InsertGuyedAnchorBlockSoilLayer(gabsl)

                    If Not firstOne Then
                        mySoils += ",(" & tempSoilLayer & ")"
                    Else
                        mySoils += "(" & tempSoilLayer & ")"
                    End If

                    firstOne = False
                Next 'Add Soil Layer INSERT statments
                GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("([INSERT ALL SOIL LAYERS])", mySoils)
                firstOne = True

                For Each gabp As GuyedAnchorBlockProfile In gab.anchor_profiles
                    Dim tempGuyedAnchorBlockProfile As String = InsertGuyedAnchorBlockProfile(gabp)

                    If Not firstOne Then
                        myProfiles += ",(" & tempGuyedAnchorBlockProfile & ")"
                    Else
                        myProfiles += "(" & tempGuyedAnchorBlockProfile & ")"
                    End If

                    firstOne = False
                Next 'Add Pier Profile INSERT statements
                GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("([INSERT ALL GUYED ANCHOR BLOCK PROFILES])", myProfiles)
                firstOne = True

                mySoils = ""
                myProfiles = ""
            Else
                Dim tempUpdater As String = ""
                tempUpdater += UpdateGuyedAnchorBlockDetail(gab)

                'comment out soil layer insertion. Added in next step if a layer does not have an ID
                GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("INSERT INTO anchor_soil_layer VALUES ([INSERT ALL SOIL LAYERS])", "--INSERT INTO anchor_soil_layer VALUES ([INSERT ALL SOIL LAYERS])")

                For Each gabsl As GuyedAnchorBlockSoilLayer In gab.soil_layers
                    If gabsl.soil_layer_id = 0 Or IsDBNull(gabsl.soil_layer_id) Then
                        tempUpdater += "INSERT INTO anchor_soil_layer VALUES (" & InsertGuyedAnchorBlockSoilLayer(gabsl) & ") " & vbNewLine
                    Else
                        tempUpdater += UpdateGuyedAnchorBlockSoilLayer(gabsl)
                    End If
                Next

                GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("INSERT INTO anchor_profile VALUES ([INSERT ALL GUYED ANCHOR BLOCK PROFILES])", "--INSERT INTO anchor_profile VALUES ([INSERT ALL GUYED ANCHOR BLOCK PROFILES])")
                For Each gabp As GuyedAnchorBlockProfile In gab.anchor_profiles
                    If gabp.profile_id = 0 Or IsDBNull(gabp.profile_id) Then
                        tempUpdater += "INSERT INTO anchor_profile VALUES (" & InsertGuyedAnchorBlockProfile(gabp) & ") " & vbNewLine
                    Else
                        tempUpdater += UpdateGuyedAnchorBlockProfile(gabp)
                    End If
                Next

                GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("SELECT * FROM TEMPORARY", tempUpdater)
            End If

            GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("[INSERT ALL GUYED ANCHOR BLOCK DETAILS]", InsertGuyedAnchorBlockDetail(gab))

            sqlSender(GuyedAnchorBlockSaver, gabDB, gabID, "0")
        Next


    End Sub

    Public Sub SaveToExcel()
        Dim gabRow As Integer = 3
        Dim soilRow As Integer = 3
        Dim profileRow As Integer = 3

        LoadNewGuyedAnchorBlock()

        With NewGuyedAnchorBlockWb

            Dim colCounter As Integer = 10
            Dim myCol As String
            Dim rowStart As Integer = 71

            For Each gab As GuyedAnchorBlock In GuyedAnchorBlocks

                'define column for anchor ID
                colCounter = 11 + 6 * (gab.local_anchor_id - 1)
                myCol = GetExcelColumnName(colCounter)

                'GUYED ANCHOR BLOCK DETAILS
                If Not IsNothing(gab.ID) Then 'FIX NEEDED. RESULTS IN 0
                    .Worksheets("Database").Range(myCol & rowStart - 12).Value = CType(gab.ID, Integer)
                Else .Worksheets("Database").Range(myCol & rowStart - 12).ClearContents
                End If

                'define column for details based on anchor profile
                colCounter = 11 + 6 * (gab.local_anchor_profile - 1)
                myCol = GetExcelColumnName(colCounter)

                If Not IsNothing(gab.anchor_depth) Then
                    .Worksheets("Database").Range(myCol & rowStart + 1).Value = CType(gab.anchor_depth, Double)
                Else .Worksheets("Database").Range(myCol & rowStart + 1).ClearContents
                End If
                If Not IsNothing(gab.anchor_width) Then
                    .Worksheets("Database").Range(myCol & rowStart + 2).Value = CType(gab.anchor_width, Double)
                Else .Worksheets("Database").Range(myCol & rowStart + 2).ClearContents
                End If
                If Not IsNothing(gab.anchor_thickness) Then
                    .Worksheets("Database").Range(myCol & rowStart + 3).Value = CType(gab.anchor_thickness, Double)
                Else .Worksheets("Database").Range(myCol & rowStart + 3).ClearContents
                End If
                If Not IsNothing(gab.anchor_length) Then
                    .Worksheets("Database").Range(myCol & rowStart + 4).Value = CType(gab.anchor_length, Double)
                Else .Worksheets("Database").Range(myCol & rowStart + 4).ClearContents
                End If
                If Not IsNothing(gab.anchor_toe_width) Then
                    .Worksheets("Database").Range(myCol & rowStart + 5).Value = CType(gab.anchor_toe_width, Double)
                Else .Worksheets("Database").Range(myCol & rowStart + 5).ClearContents
                End If
                If Not IsNothing(gab.anchor_top_rebar_size) Then
                    .Worksheets("Database").Range(myCol & rowStart + 6).Value = CType(gab.anchor_top_rebar_size, Integer)
                Else .Worksheets("Database").Range(myCol & rowStart + 6).ClearContents
                End If
                If Not IsNothing(gab.anchor_top_rebar_quantity) Then
                    .Worksheets("Database").Range(myCol & rowStart + 7).Value = CType(gab.anchor_top_rebar_quantity, Integer)
                Else .Worksheets("Database").Range(myCol & rowStart + 7).ClearContents
                End If
                If Not IsNothing(gab.anchor_front_rebar_size) Then
                    .Worksheets("Database").Range(myCol & rowStart + 8).Value = CType(gab.anchor_front_rebar_size, Integer)
                Else .Worksheets("Database").Range(myCol & rowStart + 8).ClearContents
                End If
                If Not IsNothing(gab.anchor_front_rebar_quantity) Then
                    .Worksheets("Database").Range(myCol & rowStart + 9).Value = CType(gab.anchor_front_rebar_quantity, Integer)
                Else .Worksheets("Database").Range(myCol & rowStart + 9).ClearContents
                End If
                If Not IsNothing(gab.anchor_stirrup_size) Then
                    .Worksheets("Database").Range(myCol & rowStart + 10).Value = CType(gab.anchor_stirrup_size, Integer)
                Else .Worksheets("Database").Range(myCol & rowStart + 10).ClearContents
                End If
                If Not IsNothing(gab.anchor_shaft_diameter) Then
                    .Worksheets("Database").Range(myCol & rowStart + 11).Value = CType(gab.anchor_shaft_diameter, Double)
                Else .Worksheets("Database").Range(myCol & rowStart + 11).ClearContents
                End If
                If Not IsNothing(gab.anchor_shaft_quantity) Then
                    .Worksheets("Database").Range(myCol & rowStart + 12).Value = CType(gab.anchor_shaft_quantity, Integer)
                Else .Worksheets("Database").Range(myCol & rowStart + 12).ClearContents
                End If
                If Not IsNothing(gab.anchor_shaft_area_override) Then
                    .Worksheets("Database").Range(myCol & rowStart + 13).Value = CType(gab.anchor_shaft_area_override, Double)
                Else .Worksheets("Database").Range(myCol & rowStart + 13).ClearContents
                End If
                If Not IsNothing(gab.anchor_shaft_shear_lag_factor) Then
                    .Worksheets("Database").Range(myCol & rowStart + 14).Value = CType(gab.anchor_shaft_shear_lag_factor, Double)
                Else .Worksheets("Database").Range(myCol & rowStart + 14).ClearContents
                End If
                If Not IsNothing(gab.anchor_shaft_section_type) Then
                    .Worksheets("Database").Range(myCol & rowStart + 15).Value = CType(gab.anchor_shaft_section_type, String)
                Else .Worksheets("Database").Range(myCol & rowStart + 15).ClearContents
                End If
                If Not IsNothing(gab.rebar_known) Then
                    .Worksheets("Database").Range(myCol & rowStart + 16).Value = CType(gab.rebar_known, Boolean)
                Else .Worksheets("Database").Range(myCol & rowStart + 16).ClearContents
                End If
                If Not IsNothing(gab.anchor_shaft_known) Then
                    .Worksheets("Database").Range(myCol & rowStart + 17).Value = CType(gab.anchor_shaft_known, Boolean)
                Else .Worksheets("Database").Range(myCol & rowStart + 17).ClearContents
                End If
                'If Not IsNothing(gab.neglect_depth) Then
                '    .Worksheets("Database").Range(myCol & rowStart + 20).Value = CType(gab.neglect_depth, Double)
                'Else .Worksheets("Database").Range(myCol & rowStart + 20).ClearContents
                'End If
                'If Not IsNothing(gab.groundwater_depth) Then
                '    If CType(gab.groundwater_depth, Double) = -1 Then
                '        .Worksheets("Database").Range(myCol & rowStart + 21).Value = "N/A"
                '    Else
                '        .Worksheets("Database").Range(myCol & rowStart + 21).Value = CType(gab.groundwater_depth, Double)
                '    End If
                'End If
                'If Not IsNothing(gab.soil_layer_quantity) Then
                '    .Worksheets("Database").Range(myCol & rowStart + 22).Value = CType(gab.soil_layer_quantity, Double)
                'Else .Worksheets("Database").Range(myCol & rowStart + 22).ClearContents
                'End If
                If Not IsNothing(gab.anchor_rebar_grade) Then
                    .Worksheets("Database").Range(myCol & rowStart + 35).Value = CType(gab.anchor_rebar_grade, Double)
                Else .Worksheets("Database").Range(myCol & rowStart + 35).ClearContents
                End If
                If Not IsNothing(gab.concrete_compressive_strength) Then
                    .Worksheets("Database").Range(myCol & rowStart + 36).Value = CType(gab.concrete_compressive_strength, Double)
                Else .Worksheets("Database").Range(myCol & rowStart + 36).ClearContents
                End If
                If Not IsNothing(gab.clear_cover) Then
                    .Worksheets("Database").Range(myCol & rowStart + 37).Value = CType(gab.clear_cover, Double)
                Else .Worksheets("Database").Range(myCol & rowStart + 37).ClearContents
                End If
                If Not IsNothing(gab.anchor_shaft_yield_strength) Then
                    .Worksheets("Database").Range(myCol & rowStart + 38).Value = CType(gab.anchor_shaft_yield_strength, Double)
                Else .Worksheets("Database").Range(myCol & rowStart + 38).ClearContents
                End If
                If Not IsNothing(gab.anchor_shaft_ultimate_strength) Then
                    .Worksheets("Database").Range(myCol & rowStart + 39).Value = CType(gab.anchor_shaft_ultimate_strength, Double)
                Else .Worksheets("Database").Range(myCol & rowStart + 39).ClearContents
                End If
                If Not IsNothing(gab.basic_soil_check) Then
                    .Worksheets("Database").Range(myCol & rowStart + 42).Value = CType(gab.basic_soil_check, Boolean)
                Else .Worksheets("Database").Range(myCol & rowStart + 42).ClearContents
                End If
                If Not IsNothing(gab.structural_check) Then
                    .Worksheets("Database").Range(myCol & rowStart + 43).Value = CType(gab.structural_check, Boolean)
                Else .Worksheets("Database").Range(myCol & rowStart + 43).ClearContents
                End If
                'If Not IsNothing(gab.local_anchor_id) Then
                '    .Worksheets("Database").Range(myCol & rowStart + 43).Value = CType(gab.local_anchor_id, Integer)
                'Else .Worksheets("Database").Range(myCol & rowStart + 43).ClearContents
                'End If

                'GUYED ANCHOR BLOCK PROFILES
                Dim summaryRowStart As Integer = 10

                For Each gabp As GuyedAnchorBlockProfile In gab.anchor_profiles
                    'Profile Return
                    If Not IsNothing(gabp.local_anchor_id) Then
                        .Worksheets("Profiles (RETURN)").Range("A" & profileRow).Value = CType(gabp.local_anchor_id, Integer)
                    Else .Worksheets("Profiles (RETURN)").Range("A" & profileRow).ClearContents
                    End If
                    'If Not IsNothing(gabp.reaction_position) Then
                    '    .Worksheets("Profiles (RETURN)").Range("B" & profileRow).Value = CType(gabp.reaction_position, Integer)
                    'Else .Worksheets("Profiles (RETURN)").Range("B" & profileRow).ClearContents
                    'End If
                    If Not IsNothing(gabp.anchor_id) Then
                        .Worksheets("Profiles (RETURN)").Range("C" & profileRow).Value = CType(gabp.anchor_id, Integer)
                    Else .Worksheets("Profiles (RETURN)").Range("C" & profileRow).ClearContents
                    End If
                    .Worksheets("Profiles (RETURN)").Range("D" & profileRow).Value = CType(gabp.profile_id, Integer)
                    If Not IsNothing(gabp.reaction_location) Then
                        .Worksheets("Profiles (RETURN)").Range("E" & profileRow).Value = CType(gabp.reaction_location, String)
                    Else .Worksheets("Profiles (RETURN)").Range("E" & profileRow).ClearContents
                    End If
                    If Not IsNothing(gabp.anchor_profile) Then
                        .Worksheets("Profiles (RETURN)").Range("F" & profileRow).Value = CType(gabp.anchor_profile, String)
                    Else .Worksheets("Profiles (RETURN)").Range("F" & profileRow).ClearContents
                    End If
                    If Not IsNothing(gabp.soil_profile) Then
                        .Worksheets("Profiles (RETURN)").Range("G" & profileRow).Value = CType(gabp.soil_profile, String)
                    Else .Worksheets("Profiles (RETURN)").Range("G" & profileRow).ClearContents
                    End If

                    'SUMMARY
                    If Not IsNothing(gabp.local_anchor_id) Then
                        .Worksheets("SUMMARY").Range("D" & summaryRowStart + CType(gabp.local_anchor_id, Integer)).Value = CType(gabp.anchor_profile, Integer)
                        If gabp.anchor_profile = gabp.local_anchor_id Then
                            .Worksheets("SUMMARY").Range("G" & summaryRowStart + CType(gabp.local_anchor_id, Integer)).Value = False
                        Else
                            .Worksheets("SUMMARY").Range("G" & summaryRowStart + CType(gabp.local_anchor_id, Integer)).Value = True
                        End If
                    End If
                    If Not IsNothing(gabp.local_anchor_id) Then
                        .Worksheets("SUMMARY").Range("E" & summaryRowStart + CType(gabp.local_anchor_id, Integer)).Value = CType(gabp.soil_profile, Integer)
                        If gabp.soil_profile = gabp.local_anchor_id Then
                            .Worksheets("SUMMARY").Range("H" & summaryRowStart + CType(gabp.local_anchor_id, Integer)).Value = False
                        Else
                            .Worksheets("SUMMARY").Range("H" & summaryRowStart + CType(gabp.local_anchor_id, Integer)).Value = True
                        End If
                    End If
                    '        .Worksheets("SUMMARY").Range("I" & summaryRowStart + CType(dpp.reaction_position, Integer)).Value = False
                    .Worksheets("SUMMARY").Range("I" & summaryRowStart + CType(gabp.local_anchor_id, Integer)).Value = CType(gabp.ID, Integer)

                    profileRow += 1

                Next


                'GUYED ANCHOR SOIL LAYER
                Dim soilCount As Integer
                Dim soilStart As Integer = 23
                Dim soilColCounter As Integer
                Dim mySoilCol As String

                For Each gabp As GuyedAnchorBlockProfile In gab.anchor_profiles

                    soilCount = 1

                    For Each gabSL As GuyedAnchorBlockSoilLayer In gab.soil_layers

                        soilColCounter = 11 + 6 * (gabp.soil_profile - 1)
                        mySoilCol = GetExcelColumnName(soilColCounter)
                        If Not IsNothing(gabSL.ID) Then
                            .Worksheets("Database").Range(mySoilCol & rowStart - 11 + (soilCount - 1)).Value = CType(gabSL.ID, Integer)
                        Else .Worksheets("Database").Range(mySoilCol & rowStart - 11 + (soilCount - 1)).ClearContents
                        End If
                        If Not IsNothing(gabSL.friction_angle) Then
                            .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + soilCount).Value = CType(gabSL.friction_angle, Double)
                        Else .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + soilCount).ClearContents
                        End If
                        If Not IsNothing(gab.soil_layer_quantity) Then
                            .Worksheets("Database").Range(mySoilCol & rowStart + 22).Value = CType(gab.soil_layer_quantity, Double)
                        Else .Worksheets("Database").Range(mySoilCol & rowStart + 22).ClearContents
                        End If
                        If Not IsNothing(gab.neglect_depth) Then
                            .Worksheets("Database").Range(mySoilCol & rowStart + 20).Value = CType(gab.neglect_depth, Double)
                        Else .Worksheets("Database").Range(mySoilCol & rowStart + 20).ClearContents
                        End If
                        If Not IsNothing(gab.groundwater_depth) Then
                            If CType(gab.groundwater_depth, Double) = -1 Then
                                .Worksheets("Database").Range(mySoilCol & rowStart + 21).Value = "N/A"
                            Else
                                .Worksheets("Database").Range(mySoilCol & rowStart + 21).Value = CType(gab.groundwater_depth, Double)
                            End If
                        End If

                        soilColCounter = 11 + 6 * (gabp.soil_profile - 1) + 1
                        mySoilCol = GetExcelColumnName(soilColCounter)
                        If Not IsNothing(gabSL.cohesion) Then
                            .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + soilCount).Value = CType(gabSL.cohesion, Double)
                        Else .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + soilCount).ClearContents
                        End If

                        soilColCounter = 11 + 6 * (gabp.soil_profile - 1) + 2
                        mySoilCol = GetExcelColumnName(soilColCounter)
                        If Not IsNothing(gabSL.effective_soil_density) Then
                            .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + soilCount).Value = CType(gabSL.effective_soil_density, Double)
                        Else .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + soilCount).ClearContents
                        End If

                        soilColCounter = 11 + 6 * (gabp.soil_profile - 1) + 3
                        mySoilCol = GetExcelColumnName(soilColCounter)
                        If Not IsNothing(gabSL.bottom_depth) Then
                            .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + soilCount).Value = CType(gabSL.bottom_depth, Double)
                        Else .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + soilCount).ClearContents
                        End If

                        soilColCounter = 11 + 6 * (gabp.soil_profile - 1) + 4
                        mySoilCol = GetExcelColumnName(soilColCounter)
                        If Not IsNothing(gabSL.skin_friction_override_uplift) Then
                            .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + soilCount).Value = CType(gabSL.skin_friction_override_uplift, Double)
                        Else .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + soilCount).ClearContents
                        End If

                        soilColCounter = 11 + 6 * (gabp.soil_profile - 1) + 5
                        mySoilCol = GetExcelColumnName(soilColCounter)
                        If Not IsNothing(gabSL.spt_blow_count) Then
                            .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + soilCount).Value = CType(gabSL.spt_blow_count, Double)
                        Else .Worksheets("Database").Range(mySoilCol & rowStart + soilStart + soilCount).ClearContents
                        End If

                        soilCount += 1
                    Next

                Next

                gabRow += 1
                colCounter += 1

            Next

            .Worksheets("SUMMARY").Range("EDSReactions").Value = True


            '~~~~~~~~POPULATE TOOL INPUTS WITH THE FIRST INSTANCE IN TOOL'S LOCAL DATABASE

            Dim firstReaction As String = GuyedAnchorBlocks(0).anchor_profiles(0).reaction_location
            Dim firstAnchorProfile As Integer = GuyedAnchorBlocks(0).anchor_profiles(0).anchor_profile
            Dim firstSoilProfile As Integer = GuyedAnchorBlocks(0).anchor_profiles(0).soil_profile

            colCounter = 7

            myCol = GetExcelColumnName(colCounter)

            If firstReaction <> "" Then

                'MATERIAL PROPERTIES
                If GuyedAnchorBlocks(0).anchor_rebar_grade.HasValue Then
                    .Worksheets("Input").Range("Fy").Value = CType(GuyedAnchorBlocks(0).anchor_rebar_grade, Double)
                Else .Worksheets("Inpit").Range("Fy").ClearContents
                End If
                If GuyedAnchorBlocks(0).concrete_compressive_strength.HasValue Then
                    .Worksheets("Input").Range("F\c").Value = CType(GuyedAnchorBlocks(0).concrete_compressive_strength, Double)
                Else .Worksheets("Inpit").Range("F\c").ClearContents
                End If
                If GuyedAnchorBlocks(0).clear_cover.HasValue Then
                    .Worksheets("Input").Range("cc").Value = CType(GuyedAnchorBlocks(0).clear_cover, Double)
                Else .Worksheets("Inpit").Range("cc").ClearContents
                End If
                If GuyedAnchorBlocks(0).anchor_shaft_yield_strength.HasValue Then
                    .Worksheets("Input").Range("Fy\").Value = CType(GuyedAnchorBlocks(0).anchor_shaft_yield_strength, Double)
                Else .Worksheets("Inpit").Range("Fy\").ClearContents
                End If
                If GuyedAnchorBlocks(0).anchor_shaft_ultimate_strength.HasValue Then
                    .Worksheets("Input").Range("Fu\").Value = CType(GuyedAnchorBlocks(0).anchor_shaft_ultimate_strength, Double)
                Else .Worksheets("Inpit").Range("Fu\").ClearContents
                End If

                'GUY ANCHOR PROPERTIES
                If GuyedAnchorBlocks(0).anchor_depth.HasValue Then
                    .Worksheets("Input").Range("Da").Value = CType(GuyedAnchorBlocks(0).anchor_depth, Double)
                Else .Worksheets("Input").Range("Da").ClearContents
                End If
                If GuyedAnchorBlocks(0).anchor_width.HasValue Then
                    .Worksheets("Input").Range("Wa").Value = CType(GuyedAnchorBlocks(0).anchor_width, Double)
                Else .Worksheets("Input").Range("Wa").ClearContents
                End If
                If GuyedAnchorBlocks(0).anchor_thickness.HasValue Then
                    .Worksheets("Input").Range("Ta").Value = CType(GuyedAnchorBlocks(0).anchor_thickness, Double)
                Else .Worksheets("Input").Range("Ta").ClearContents
                End If
                If GuyedAnchorBlocks(0).anchor_length.HasValue Then
                    .Worksheets("Input").Range("La").Value = CType(GuyedAnchorBlocks(0).anchor_length, Double)
                Else .Worksheets("Input").Range("La").ClearContents
                End If
                If GuyedAnchorBlocks(0).anchor_toe_width.HasValue Then
                    .Worksheets("Input").Range("toe").Value = CType(GuyedAnchorBlocks(0).anchor_toe_width, Double)
                Else .Worksheets("Input").Range("toe").ClearContents
                End If
                If GuyedAnchorBlocks(0).anchor_top_rebar_size.HasValue Then
                    .Worksheets("Input").Range("Sat").Value = CType(GuyedAnchorBlocks(0).anchor_top_rebar_size, Integer)
                Else .Worksheets("Input").Range("Sat").ClearContents
                End If
                If GuyedAnchorBlocks(0).anchor_top_rebar_quantity.HasValue Then
                    .Worksheets("Input").Range("mu").Value = CType(GuyedAnchorBlocks(0).anchor_top_rebar_quantity, Integer)
                Else .Worksheets("Input").Range("mu").ClearContents
                End If
                If GuyedAnchorBlocks(0).anchor_front_rebar_size.HasValue Then
                    .Worksheets("Input").Range("Saf").Value = CType(GuyedAnchorBlocks(0).anchor_front_rebar_size, Integer)
                Else .Worksheets("Input").Range("Saf").ClearContents
                End If
                If GuyedAnchorBlocks(0).anchor_front_rebar_quantity.HasValue Then
                    .Worksheets("Input").Range("ms").Value = CType(GuyedAnchorBlocks(0).anchor_front_rebar_quantity, Integer)
                Else .Worksheets("Input").Range("ms").ClearContents
                End If
                If GuyedAnchorBlocks(0).anchor_stirrup_size.HasValue Then
                    .Worksheets("Input").Range("stirrup").Value = CType(GuyedAnchorBlocks(0).anchor_stirrup_size, Integer)
                Else .Worksheets("Input").Range("stirrup").ClearContents
                End If
                If GuyedAnchorBlocks(0).anchor_shaft_diameter.HasValue Then
                    .Worksheets("Input").Range("ds").Value = CType(GuyedAnchorBlocks(0).anchor_shaft_diameter, Double)
                Else .Worksheets("Input").Range("ds").ClearContents
                End If
                If GuyedAnchorBlocks(0).anchor_shaft_quantity.HasValue Then
                    .Worksheets("Input").Range("n").Value = CType(GuyedAnchorBlocks(0).anchor_shaft_quantity, Integer)
                Else .Worksheets("Input").Range("n").ClearContents
                End If
                If GuyedAnchorBlocks(0).anchor_shaft_area_override.HasValue Then
                    .Worksheets("Input").Range("anchor_area").Value = CType(GuyedAnchorBlocks(0).anchor_shaft_area_override, Double)
                Else .Worksheets("Input").Range("anchor_area").ClearContents
                End If
                If GuyedAnchorBlocks(0).anchor_shaft_shear_lag_factor.HasValue Then
                    .Worksheets("Input").Range("u").Value = CType(GuyedAnchorBlocks(0).anchor_shaft_shear_lag_factor, Double)
                Else .Worksheets("Input").Range("u").ClearContents
                End If
                If GuyedAnchorBlocks(0).neglect_depth.HasValue Then
                    .Worksheets("Input").Range("Fd").Value = CType(GuyedAnchorBlocks(0).neglect_depth, Double)
                Else .Worksheets("Input").Range("Fd").ClearContents
                End If
                If GuyedAnchorBlocks(0).groundwater_depth.HasValue Then
                    If CType(GuyedAnchorBlocks(0).groundwater_depth, Double) = -1 Then
                        .Worksheets("Input").Range("gw").Value = "N/A"
                    Else
                        .Worksheets("Input").Range("gw").Value = CType(GuyedAnchorBlocks(0).groundwater_depth, Double)
                    End If
                Else .Worksheets("Input").Range("gw").ClearContents
                End If
                If GuyedAnchorBlocks(0).soil_layer_quantity.HasValue Then
                    .Worksheets("Input").Range("Layer_Qty").Value = CType(GuyedAnchorBlocks(0).soil_layer_quantity, Integer)
                Else .Worksheets("Input").Range("Layer_Qty").ClearContents
                End If

                'OPTIONS
                .Worksheets("Input").Range("S12").Value = CType(GuyedAnchorBlocks(0).rebar_known, Boolean)
                .Worksheets("Input").Range("S13").Value = CType(GuyedAnchorBlocks(0).anchor_shaft_known, Boolean)
                .Worksheets("Input").Range("S14").Value = CType(GuyedAnchorBlocks(0).basic_soil_check, Boolean)
                .Worksheets("Input").Range("S15").Value = CType(GuyedAnchorBlocks(0).structural_check, Boolean)

                'SOIL
                Dim soilRowStart As Integer = 28
                Dim soilCount As Integer = 1

                For Each gabSL As GuyedAnchorBlockSoilLayer In GuyedAnchorBlocks(0).soil_layers

                    If gabSL.friction_angle.HasValue Then
                        .Worksheets("Input").Range("G" & soilRowStart + soilCount).Value = CType(gabSL.friction_angle, Double)
                    Else .Worksheets("Input").Range("G" & soilRowStart + soilCount).ClearContents
                    End If
                    If gabSL.cohesion.HasValue Then
                        .Worksheets("Input").Range("H" & soilRowStart + soilCount).Value = CType(gabSL.cohesion, Double)
                    Else .Worksheets("Input").Range("H" & soilRowStart + soilCount).ClearContents
                    End If
                    If gabSL.effective_soil_density.HasValue Then
                        .Worksheets("Input").Range("I" & soilRowStart + soilCount).Value = CType(gabSL.effective_soil_density, Double)
                    Else .Worksheets("Input").Range("I" & soilRowStart + soilCount).ClearContents
                    End If
                    If gabSL.bottom_depth.HasValue Then
                        .Worksheets("Input").Range("J" & soilRowStart + soilCount).Value = CType(gabSL.bottom_depth, Double)
                    Else .Worksheets("Input").Range("J" & soilRowStart + soilCount).ClearContents
                    End If
                    If gabSL.skin_friction_override_uplift.HasValue Then
                        .Worksheets("Input").Range("K" & soilRowStart + soilCount).Value = CType(gabSL.skin_friction_override_uplift, Double)
                    Else .Worksheets("Input").Range("K" & soilRowStart + soilCount).ClearContents
                    End If
                    If gabSL.spt_blow_count.HasValue Then
                        .Worksheets("Input").Range("L" & soilRowStart + soilCount).Value = CType(gabSL.spt_blow_count, Integer)
                    Else .Worksheets("Input").Range("L" & soilRowStart + soilCount).ClearContents
                    End If

                    soilCount += 1

                Next

                'LOCATION
                .Worksheets("Input").Range("Location").Value = firstReaction

            End If


        End With


        SaveAndCloseGuyedAnchorBlock()
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

    Private Sub LoadNewGuyedAnchorBlock()
        NewGuyedAnchorBlockWb.LoadDocument(GuyedAnchorBlockTemplatePath, GuyedAnchorBlockFileType)
        NewGuyedAnchorBlockWb.BeginUpdate()
    End Sub

    Private Sub SaveAndCloseGuyedAnchorBlock()
        NewGuyedAnchorBlockWb.EndUpdate()
        NewGuyedAnchorBlockWb.SaveDocument(ExcelFilePath, GuyedAnchorBlockFileType)
    End Sub
#End Region

#Region "SQL Insert Statements"
    Private Function InsertGuyedAnchorBlockDetail(ByVal gab As GuyedAnchorBlock) As String
        Dim insertString As String = ""

        insertString += "@FndID"
        'insertString += "," & IIf(IsNothing(dp.foundation_depth), "Null", "'" & dp.foundation_depth.ToString & "'")
        '    insertString += "," & IIf(IsNothing(dp.extension_above_grade), "Null", "'" & dp.extension_above_grade.ToString & "'")
        '    insertString += "," & IIf(IsNothing(dp.groundwater_depth), "Null", "'" & dp.groundwater_depth.ToString & "'")
        '    insertString += "," & IIf(IsNothing(dp.assume_min_steel), "Null", "'" & dp.assume_min_steel.ToString & "'")
        '    insertString += "," & IIf(IsNothing(dp.check_shear_along_depth), "Null", "'" & dp.check_shear_along_depth.ToString & "'")
        '    insertString += "," & IIf(IsNothing(dp.utilize_shear_friction_methodology), "Null", "'" & dp.utilize_shear_friction_methodology.ToString & "'")
        '    insertString += "," & IIf(IsNothing(dp.embedded_pole), "Null", "'" & dp.embedded_pole.ToString & "'")
        '    insertString += "," & IIf(IsNothing(dp.belled_pier), "Null", "'" & dp.belled_pier.ToString & "'")
        '    insertString += "," & IIf(IsNothing(dp.soil_layer_quantity), "Null", "'" & dp.soil_layer_quantity.ToString & "'")
        '    insertString += "," & IIf(IsNothing(dp.concrete_compressive_strength), "Null", "'" & dp.concrete_compressive_strength.ToString & "'")
        '    insertString += "," & IIf(IsNothing(dp.tie_yield_strength), "Null", "'" & dp.tie_yield_strength.ToString & "'")
        '    insertString += "," & IIf(IsNothing(dp.longitudinal_rebar_yield_strength), "Null", "'" & dp.longitudinal_rebar_yield_strength.ToString & "'")
        '    insertString += "," & IIf(IsNothing(dp.rebar_effective_depths), "Null", "'" & dp.rebar_effective_depths.ToString & "'")
        '    insertString += "," & IIf(IsNothing(dp.rebar_cage_2_fy_override), "Null", "'" & dp.rebar_cage_2_fy_override.ToString & "'")
        '    insertString += "," & IIf(IsNothing(dp.rebar_cage_3_fy_override), "Null", "'" & dp.rebar_cage_3_fy_override.ToString & "'")
        '    insertString += "," & IIf(IsNothing(dp.shear_override_crit_depth), "Null", "'" & dp.shear_override_crit_depth.ToString & "'")
        '    insertString += "," & IIf(IsNothing(dp.shear_crit_depth_override_comp), "Null", "'" & dp.shear_crit_depth_override_comp.ToString & "'")
        '    insertString += "," & IIf(IsNothing(dp.shear_crit_depth_override_uplift), "Null", "'" & dp.shear_crit_depth_override_uplift.ToString & "'")
        '    insertString += "," & IIf(IsNothing(dp.local_drilled_pier_id), "Null", "'" & dp.local_drilled_pier_id.ToString & "'")
        '    insertString += "," & IIf(IsNothing(dp.bearing_type_toggle), "Null", "'" & dp.bearing_type_toggle.ToString & "'")

        Return insertString
    End Function
    Private Function InsertGuyedAnchorBlockSoilLayer(ByVal dpsl As GuyedAnchorBlockSoilLayer) As String
        Dim insertString As String = ""

        'insertString += "@DpID"
        'insertString += "," & IIf(IsNothing(dpsl.bottom_depth), "Null", "'" & dpsl.bottom_depth.ToString & "'")
        'insertString += "," & IIf(IsNothing(dpsl.effective_soil_density), "Null", "'" & dpsl.effective_soil_density.ToString & "'")
        'insertString += "," & IIf(IsNothing(dpsl.cohesion), "Null", "'" & dpsl.cohesion.ToString & "'")
        'insertString += "," & IIf(IsNothing(dpsl.friction_angle), "Null", "'" & dpsl.friction_angle.ToString & "'")
        'insertString += "," & IIf(IsNothing(dpsl.skin_friction_override_comp), "Null", "'" & dpsl.skin_friction_override_comp.ToString & "'")
        'insertString += "," & IIf(IsNothing(dpsl.skin_friction_override_uplift), "Null", "'" & dpsl.skin_friction_override_uplift.ToString & "'")
        'insertString += "," & IIf(IsNothing(dpsl.nominal_bearing_capacity), "Null", "'" & dpsl.nominal_bearing_capacity.ToString & "'")
        'insertString += "," & IIf(IsNothing(dpsl.spt_blow_count), "Null", "'" & dpsl.spt_blow_count.ToString & "'")
        'insertString += "," & IIf(IsNothing(dpsl.local_soil_layer_id), "Null", "'" & dpsl.local_soil_layer_id.ToString & "'")

        Return insertString
    End Function
    Private Function InsertGuyedAnchorBlockProfile(ByVal gabp As GuyedAnchorBlockProfile) As String
        Dim insertString As String = ""

        'insertString += "@DpID"
        'insertString += "," & IIf(IsNothing(dpp.reaction_position), "Null", "'" & dpp.reaction_position.ToString & "'")
        'insertString += "," & IIf(IsNothing(dpp.reaction_location), "Null", "'" & dpp.reaction_location.ToString & "'")
        'insertString += "," & IIf(IsNothing(dpp.drilled_pier_profile), "Null", "'" & dpp.drilled_pier_profile.ToString & "'")
        'insertString += "," & IIf(IsNothing(dpp.soil_profile), "Null", "'" & dpp.soil_profile.ToString & "'")

        Return insertString
    End Function
    'Private Function InsertDrilledPierDetail(ByVal dp As DrilledPier) As String
    '    Dim insertString As String = ""

    '    insertString += "@FndID"
    '    insertString += "," & IIf(IsNothing(dp.foundation_depth), "Null", "'" & dp.foundation_depth.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dp.extension_above_grade), "Null", "'" & dp.extension_above_grade.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dp.groundwater_depth), "Null", "'" & dp.groundwater_depth.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dp.assume_min_steel), "Null", "'" & dp.assume_min_steel.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dp.check_shear_along_depth), "Null", "'" & dp.check_shear_along_depth.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dp.utilize_shear_friction_methodology), "Null", "'" & dp.utilize_shear_friction_methodology.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dp.embedded_pole), "Null", "'" & dp.embedded_pole.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dp.belled_pier), "Null", "'" & dp.belled_pier.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dp.soil_layer_quantity), "Null", "'" & dp.soil_layer_quantity.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dp.concrete_compressive_strength), "Null", "'" & dp.concrete_compressive_strength.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dp.tie_yield_strength), "Null", "'" & dp.tie_yield_strength.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dp.longitudinal_rebar_yield_strength), "Null", "'" & dp.longitudinal_rebar_yield_strength.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dp.rebar_effective_depths), "Null", "'" & dp.rebar_effective_depths.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dp.rebar_cage_2_fy_override), "Null", "'" & dp.rebar_cage_2_fy_override.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dp.rebar_cage_3_fy_override), "Null", "'" & dp.rebar_cage_3_fy_override.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dp.shear_override_crit_depth), "Null", "'" & dp.shear_override_crit_depth.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dp.shear_crit_depth_override_comp), "Null", "'" & dp.shear_crit_depth_override_comp.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dp.shear_crit_depth_override_uplift), "Null", "'" & dp.shear_crit_depth_override_uplift.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dp.local_drilled_pier_id), "Null", "'" & dp.local_drilled_pier_id.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dp.bearing_type_toggle), "Null", "'" & dp.bearing_type_toggle.ToString & "'")

    '    Return insertString
    'End Function

    'Private Function InsertDrilledPierBell(ByVal bp As DrilledPierBelledPier) As String
    '    Dim insertString As String = ""

    '    insertString += "@DpID"
    '    insertString += "," & IIf(IsNothing(bp.belled_pier_option), "Null", "'" & bp.belled_pier_option.ToString & "'")
    '    insertString += "," & IIf(IsNothing(bp.bottom_diameter_of_bell), "Null", "'" & bp.bottom_diameter_of_bell.ToString & "'")
    '    insertString += "," & IIf(IsNothing(bp.bell_input_type), "Null", "'" & bp.bell_input_type.ToString & "'")
    '    insertString += "," & IIf(IsNothing(bp.bell_angle), "Null", "'" & bp.bell_angle.ToString & "'")
    '    insertString += "," & IIf(IsNothing(bp.bell_height), "Null", "'" & bp.bell_height.ToString & "'")
    '    insertString += "," & IIf(IsNothing(bp.bell_toe_height), "Null", "'" & bp.bell_toe_height.ToString & "'")
    '    insertString += "," & IIf(IsNothing(bp.neglect_top_soil_layer), "Null", "'" & bp.neglect_top_soil_layer.ToString & "'")
    '    insertString += "," & IIf(IsNothing(bp.swelling_expansive_soil), "Null", "'" & bp.swelling_expansive_soil.ToString & "'")
    '    insertString += "," & IIf(IsNothing(bp.depth_of_expansive_soil), "Null", "'" & bp.depth_of_expansive_soil.ToString & "'")
    '    insertString += "," & IIf(IsNothing(bp.expansive_soil_force), "Null", "'" & bp.expansive_soil_force.ToString & "'")

    '    Return insertString
    'End Function

    'Private Function InsertDrilledPierEmbed(ByVal ep As DrilledPierEmbeddedPier) As String
    '    Dim insertString As String = ""

    '    insertString += "@DpID"
    '    insertString += "," & IIf(IsNothing(ep.embedded_pole_option), "Null", "'" & ep.embedded_pole_option.ToString & "'")
    '    insertString += "," & IIf(IsNothing(ep.encased_in_concrete), "Null", "'" & ep.encased_in_concrete.ToString & "'")
    '    insertString += "," & IIf(IsNothing(ep.pole_side_quantity), "Null", "'" & ep.pole_side_quantity.ToString & "'")
    '    insertString += "," & IIf(IsNothing(ep.pole_yield_strength), "Null", "'" & ep.pole_yield_strength.ToString & "'")
    '    insertString += "," & IIf(IsNothing(ep.pole_thickness), "Null", "'" & ep.pole_thickness.ToString & "'")
    '    insertString += "," & IIf(IsNothing(ep.embedded_pole_input_type), "Null", "'" & ep.embedded_pole_input_type.ToString & "'")
    '    insertString += "," & IIf(IsNothing(ep.pole_diameter_toc), "Null", "'" & ep.pole_diameter_toc.ToString & "'")
    '    insertString += "," & IIf(IsNothing(ep.pole_top_diameter), "Null", "'" & ep.pole_top_diameter.ToString & "'")
    '    insertString += "," & IIf(IsNothing(ep.pole_bottom_diameter), "Null", "'" & ep.pole_bottom_diameter.ToString & "'")
    '    insertString += "," & IIf(IsNothing(ep.pole_section_length), "Null", "'" & ep.pole_section_length.ToString & "'")
    '    insertString += "," & IIf(IsNothing(ep.pole_taper_factor), "Null", "'" & ep.pole_taper_factor.ToString & "'")
    '    insertString += "," & IIf(IsNothing(ep.pole_bend_radius_override), "Null", "'" & ep.pole_bend_radius_override.ToString & "'")

    '    Return insertString
    'End Function

    'Private Function InsertDrilledPierSoilLayer(ByVal dpsl As DrilledPierSoilLayer) As String
    '    Dim insertString As String = ""

    '    insertString += "@DpID"
    '    insertString += "," & IIf(IsNothing(dpsl.bottom_depth), "Null", "'" & dpsl.bottom_depth.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dpsl.effective_soil_density), "Null", "'" & dpsl.effective_soil_density.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dpsl.cohesion), "Null", "'" & dpsl.cohesion.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dpsl.friction_angle), "Null", "'" & dpsl.friction_angle.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dpsl.skin_friction_override_comp), "Null", "'" & dpsl.skin_friction_override_comp.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dpsl.skin_friction_override_uplift), "Null", "'" & dpsl.skin_friction_override_uplift.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dpsl.nominal_bearing_capacity), "Null", "'" & dpsl.nominal_bearing_capacity.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dpsl.spt_blow_count), "Null", "'" & dpsl.spt_blow_count.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dpsl.local_soil_layer_id), "Null", "'" & dpsl.local_soil_layer_id.ToString & "'")

    '    Return insertString
    'End Function

    'Private Function InsertDrilledPierSection(ByVal dpsec As DrilledPierSection) As String
    '    Dim insertString As String = ""

    '    insertString += "@DpID"
    '    insertString += "," & IIf(IsNothing(dpsec.pier_diameter), "Null", "'" & dpsec.pier_diameter.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dpsec.clear_cover), "Null", "'" & dpsec.clear_cover.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dpsec.clear_cover_rebar_cage_option), "Null", "'" & dpsec.clear_cover_rebar_cage_option.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dpsec.tie_size), "Null", "'" & dpsec.tie_size.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dpsec.tie_spacing), "Null", "'" & dpsec.tie_spacing.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dpsec.bottom_elevation), "Null", "'" & dpsec.bottom_elevation.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dpsec.local_section_id), "Null", "'" & dpsec.local_section_id.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dpsec.rho_override), "Null", "'" & dpsec.rho_override.ToString & "'")

    '    Return insertString
    'End Function

    'Private Function InsertDrilledPierRebar(ByVal dpreb As DrilledPierRebar) As String
    '    Dim insertString As String = ""

    '    insertString += "@SecID"
    '    insertString += "," & IIf(IsNothing(dpreb.longitudinal_rebar_quantity), "Null", "'" & dpreb.longitudinal_rebar_quantity.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dpreb.longitudinal_rebar_size), "Null", "'" & dpreb.longitudinal_rebar_size.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dpreb.longitudinal_rebar_cage_diameter), "Null", "'" & dpreb.longitudinal_rebar_cage_diameter.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dpreb.local_rebar_id), "Null", "'" & dpreb.local_rebar_id.ToString & "'")

    '    Return insertString
    'End Function
    'Private Function InsertDrilledPierProfile(ByVal dpp As DrilledPierProfile) As String
    '    Dim insertString As String = ""

    '    insertString += "@DpID"
    '    insertString += "," & IIf(IsNothing(dpp.reaction_position), "Null", "'" & dpp.reaction_position.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dpp.reaction_location), "Null", "'" & dpp.reaction_location.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dpp.drilled_pier_profile), "Null", "'" & dpp.drilled_pier_profile.ToString & "'")
    '    insertString += "," & IIf(IsNothing(dpp.soil_profile), "Null", "'" & dpp.soil_profile.ToString & "'")

    '    Return insertString
    'End Function
#End Region

#Region "SQL Update Statements"
    Private Function UpdateGuyedAnchorBlockDetail(ByVal gab As GuyedAnchorBlock) As String
        Dim updateString As String = ""

        'updateString += "UPDATE drilled_pier_details SET "
        'updateString += "foundation_depth=" & IIf(IsNothing(dp.foundation_depth), "Null", "'" & dp.foundation_depth.ToString & "'")
        'updateString += ", extension_above_grade=" & IIf(IsNothing(dp.extension_above_grade), "Null", "'" & dp.extension_above_grade.ToString & "'")
        'updateString += ", groundwater_depth=" & IIf(IsNothing(dp.groundwater_depth), "Null", "'" & dp.groundwater_depth.ToString & "'")
        'updateString += ", assume_min_steel=" & IIf(IsNothing(dp.assume_min_steel), "Null", "'" & dp.assume_min_steel.ToString & "'")
        'updateString += ", check_shear_along_depth=" & IIf(IsNothing(dp.check_shear_along_depth), "Null", "'" & dp.check_shear_along_depth.ToString & "'")
        'updateString += ", utilize_shear_friction_methodology=" & IIf(IsNothing(dp.utilize_shear_friction_methodology), "Null", "'" & dp.utilize_shear_friction_methodology.ToString & "'")
        'updateString += ", embedded_pole=" & IIf(IsNothing(dp.embedded_pole), "Null", "'" & dp.embedded_pole.ToString & "'")
        'updateString += ", belled_pier=" & IIf(IsNothing(dp.belled_pier), "Null", "'" & dp.belled_pier.ToString & "'")
        'updateString += ", soil_layer_quantity=" & IIf(IsNothing(dp.soil_layer_quantity), "Null", "'" & dp.soil_layer_quantity.ToString & "'")
        'updateString += ", concrete_compressive_strength=" & IIf(IsNothing(dp.concrete_compressive_strength), "Null", "'" & dp.concrete_compressive_strength.ToString & "'")
        'updateString += ", tie_yield_strength=" & IIf(IsNothing("'" & dp.tie_yield_strength), "Null", "'" & dp.tie_yield_strength.ToString & "'")
        'updateString += ", longitudinal_rebar_yield_strength=" & IIf(IsNothing(dp.longitudinal_rebar_yield_strength), "Null", "'" & dp.longitudinal_rebar_yield_strength.ToString & "'")
        'updateString += ", rebar_effective_depths=" & IIf(IsNothing(dp.rebar_effective_depths), "Null", "'" & dp.rebar_effective_depths.ToString & "'")
        'updateString += ", rebar_cage_2_fy_override=" & IIf(IsNothing(dp.rebar_cage_2_fy_override), "Null", "'" & dp.rebar_cage_2_fy_override.ToString & "'")
        'updateString += ", rebar_cage_3_fy_override=" & IIf(IsNothing(dp.rebar_cage_3_fy_override), "Null", "'" & dp.rebar_cage_3_fy_override.ToString & "'")
        'updateString += ", shear_override_crit_depth=" & IIf(IsNothing(dp.shear_override_crit_depth), "Null", "'" & dp.shear_override_crit_depth.ToString & "'")
        'updateString += ", shear_crit_depth_override_comp=" & IIf(IsNothing(dp.shear_crit_depth_override_comp), "Null", "'" & dp.shear_crit_depth_override_comp.ToString & "'")
        'updateString += ", shear_crit_depth_override_uplift=" & IIf(IsNothing(dp.shear_crit_depth_override_uplift), "Null", "'" & dp.shear_crit_depth_override_uplift.ToString & "'")
        'updateString += ", local_drilled_pier_id=" & IIf(IsNothing(dp.local_drilled_pier_id), "Null", "'" & dp.local_drilled_pier_id.ToString & "'")
        'updateString += ", bearing_type_toggle=" & IIf(IsNothing(dp.bearing_type_toggle), "Null", "'" & dp.bearing_type_toggle.ToString & "'")
        'updateString += " WHERE ID=" & dp.pier_id & vbNewLine

        Return updateString
    End Function
    Private Function UpdateGuyedAnchorBlockSoilLayer(ByVal gabsl As GuyedAnchorBlockSoilLayer) As String
        Dim updateString As String = ""

        'updateString += "UPDATE drilled_pier_soil_layer SET "
        'updateString += "bottom_depth=" & IIf(IsNothing(dpsl.bottom_depth), "Null", "'" & dpsl.bottom_depth.ToString & "'")
        'updateString += ", effective_soil_density=" & IIf(IsNothing(dpsl.effective_soil_density), "Null", "'" & dpsl.effective_soil_density.ToString & "'")
        'updateString += ", cohesion=" & IIf(IsNothing(dpsl.cohesion), "Null", "'" & dpsl.cohesion.ToString & "'")
        'updateString += ", friction_angle=" & IIf(IsNothing(dpsl.friction_angle), "Null", "'" & dpsl.friction_angle.ToString & "'")
        'updateString += ", skin_friction_override_comp=" & IIf(IsNothing(dpsl.skin_friction_override_comp), "Null", "'" & dpsl.skin_friction_override_comp.ToString & "'")
        'updateString += ", skin_friction_override_uplift=" & IIf(IsNothing(dpsl.skin_friction_override_uplift), "Null", "'" & dpsl.skin_friction_override_uplift.ToString & "'")
        'updateString += ", nominal_bearing_capacity=" & IIf(IsNothing(dpsl.nominal_bearing_capacity), "Null", "'" & dpsl.nominal_bearing_capacity.ToString & "'")
        'updateString += ", spt_blow_count=" & IIf(IsNothing(dpsl.spt_blow_count), "Null", "'" & dpsl.spt_blow_count.ToString & "'")
        'updateString += ", local_soil_layer_id=" & IIf(IsNothing(dpsl.local_soil_layer_id), "Null", "'" & dpsl.local_soil_layer_id.ToString & "'")
        'updateString += " WHERE ID=" & dpsl.soil_layer_id & vbNewLine

        Return updateString
    End Function

    Private Function UpdateGuyedAnchorBlockProfile(ByVal gabp As GuyedAnchorBlockProfile) As String
        Dim updateString As String = ""

        'updateString += "UPDATE drilled_pier_profile SET "
        'updateString += ", reaction_position=" & IIf(IsNothing(dpp.reaction_position), "Null", "'" & dpp.reaction_position.ToString & "'")
        'updateString += ", reaction_location=" & IIf(IsNothing(dpp.reaction_location), "Null", "'" & dpp.reaction_location.ToString & "'")
        'updateString += ", drilled_pier_profile=" & IIf(IsNothing(dpp.drilled_pier_profile), "Null", "'" & dpp.drilled_pier_profile.ToString & "'")
        'updateString += ", soil_profile=" & IIf(IsNothing(dpp.soil_profile), "Null", "'" & dpp.soil_profile.ToString & "'")
        'updateString += " WHERE ID=" & dpp.profile_id & vbNewLine

        Return updateString
    End Function
    'Private Function UpdateDrilledPierDetail(ByVal dp As DrilledPier) As String
    '    Dim updateString As String = ""

    '    updateString += "UPDATE drilled_pier_details SET "
    '    updateString += "foundation_depth=" & IIf(IsNothing(dp.foundation_depth), "Null", "'" & dp.foundation_depth.ToString & "'")
    '    updateString += ", extension_above_grade=" & IIf(IsNothing(dp.extension_above_grade), "Null", "'" & dp.extension_above_grade.ToString & "'")
    '    updateString += ", groundwater_depth=" & IIf(IsNothing(dp.groundwater_depth), "Null", "'" & dp.groundwater_depth.ToString & "'")
    '    updateString += ", assume_min_steel=" & IIf(IsNothing(dp.assume_min_steel), "Null", "'" & dp.assume_min_steel.ToString & "'")
    '    updateString += ", check_shear_along_depth=" & IIf(IsNothing(dp.check_shear_along_depth), "Null", "'" & dp.check_shear_along_depth.ToString & "'")
    '    updateString += ", utilize_shear_friction_methodology=" & IIf(IsNothing(dp.utilize_shear_friction_methodology), "Null", "'" & dp.utilize_shear_friction_methodology.ToString & "'")
    '    updateString += ", embedded_pole=" & IIf(IsNothing(dp.embedded_pole), "Null", "'" & dp.embedded_pole.ToString & "'")
    '    updateString += ", belled_pier=" & IIf(IsNothing(dp.belled_pier), "Null", "'" & dp.belled_pier.ToString & "'")
    '    updateString += ", soil_layer_quantity=" & IIf(IsNothing(dp.soil_layer_quantity), "Null", "'" & dp.soil_layer_quantity.ToString & "'")
    '    updateString += ", concrete_compressive_strength=" & IIf(IsNothing(dp.concrete_compressive_strength), "Null", "'" & dp.concrete_compressive_strength.ToString & "'")
    '    updateString += ", tie_yield_strength=" & IIf(IsNothing("'" & dp.tie_yield_strength), "Null", "'" & dp.tie_yield_strength.ToString & "'")
    '    updateString += ", longitudinal_rebar_yield_strength=" & IIf(IsNothing(dp.longitudinal_rebar_yield_strength), "Null", "'" & dp.longitudinal_rebar_yield_strength.ToString & "'")
    '    updateString += ", rebar_effective_depths=" & IIf(IsNothing(dp.rebar_effective_depths), "Null", "'" & dp.rebar_effective_depths.ToString & "'")
    '    updateString += ", rebar_cage_2_fy_override=" & IIf(IsNothing(dp.rebar_cage_2_fy_override), "Null", "'" & dp.rebar_cage_2_fy_override.ToString & "'")
    '    updateString += ", rebar_cage_3_fy_override=" & IIf(IsNothing(dp.rebar_cage_3_fy_override), "Null", "'" & dp.rebar_cage_3_fy_override.ToString & "'")
    '    updateString += ", shear_override_crit_depth=" & IIf(IsNothing(dp.shear_override_crit_depth), "Null", "'" & dp.shear_override_crit_depth.ToString & "'")
    '    updateString += ", shear_crit_depth_override_comp=" & IIf(IsNothing(dp.shear_crit_depth_override_comp), "Null", "'" & dp.shear_crit_depth_override_comp.ToString & "'")
    '    updateString += ", shear_crit_depth_override_uplift=" & IIf(IsNothing(dp.shear_crit_depth_override_uplift), "Null", "'" & dp.shear_crit_depth_override_uplift.ToString & "'")
    '    updateString += ", local_drilled_pier_id=" & IIf(IsNothing(dp.local_drilled_pier_id), "Null", "'" & dp.local_drilled_pier_id.ToString & "'")
    '    updateString += ", bearing_type_toggle=" & IIf(IsNothing(dp.bearing_type_toggle), "Null", "'" & dp.bearing_type_toggle.ToString & "'")
    '    updateString += " WHERE ID=" & dp.pier_id & vbNewLine

    '    Return updateString
    'End Function

    'Private Function UpdateDrilledPierBell(ByVal bp As DrilledPierBelledPier) As String
    '    Dim updateString As String = ""

    '    updateString += "UPDATE belled_pier_details SET "
    '    updateString += "belled_pier_option=" & IIf(IsNothing(bp.belled_pier_option), "Null", "'" & bp.belled_pier_option.ToString & "'")
    '    updateString += ", bottom_diameter_of_bell=" & IIf(IsNothing(bp.bottom_diameter_of_bell), "Null", "'" & bp.bottom_diameter_of_bell.ToString & "'")
    '    updateString += ", bell_input_type=" & IIf(IsNothing(bp.bell_input_type), "Null", "'" & bp.bell_input_type.ToString & "'")
    '    updateString += ", bell_angle=" & IIf(IsNothing(bp.bell_angle), "Null", "'" & bp.bell_angle.ToString & "'")
    '    updateString += ", bell_height=" & IIf(IsNothing(bp.bell_height), "Null", "'" & bp.bell_height.ToString & "'")
    '    updateString += ", bell_toe_height=" & IIf(IsNothing(bp.bell_toe_height), "Null", "'" & bp.bell_toe_height.ToString & "'")
    '    updateString += ", neglect_top_soil_layer=" & IIf(IsNothing(bp.neglect_top_soil_layer), "Null", "'" & bp.neglect_top_soil_layer.ToString & "'")
    '    updateString += ", swelling_expansive_soil=" & IIf(IsNothing(bp.swelling_expansive_soil), "Null", "'" & bp.swelling_expansive_soil.ToString & "'")
    '    updateString += ", depth_of_expansive_soil=" & IIf(IsNothing(bp.depth_of_expansive_soil), "Null", "'" & bp.depth_of_expansive_soil.ToString & "'")
    '    updateString += ", expansive_soil_force=" & IIf(IsNothing(bp.expansive_soil_force), "Null", "'" & bp.expansive_soil_force.ToString & "'")
    '    updateString += " WHERE ID=" & bp.belled_pier_id & vbNewLine

    '    Return updateString
    'End Function

    'Private Function UpdateDrilledPierEmbed(ByVal ep As DrilledPierEmbeddedPier) As String
    '    Dim updateString As String = ""

    '    updateString += "UPDATE embedded_pole_details SET "
    '    updateString += "embedded_pole_option=" & IIf(IsNothing(ep.embedded_pole_option), "Null", "'" & ep.embedded_pole_option.ToString & "'")
    '    updateString += ", encased_in_concrete=" & IIf(IsNothing(ep.encased_in_concrete), "Null", "'" & ep.encased_in_concrete.ToString & "'")
    '    updateString += ", pole_side_quantity=" & IIf(IsNothing(ep.pole_side_quantity), "Null", "'" & ep.pole_side_quantity.ToString & "'")
    '    updateString += ", pole_yield_strength=" & IIf(IsNothing(ep.pole_yield_strength), "Null", "'" & ep.pole_yield_strength.ToString & "'")
    '    updateString += ", pole_thickness=" & IIf(IsNothing(ep.pole_thickness), "Null", "'" & ep.pole_thickness.ToString & "'")
    '    updateString += ", embedded_pole_input_type=" & IIf(IsNothing(ep.embedded_pole_input_type), "Null", "'" & ep.embedded_pole_input_type.ToString & "'")
    '    updateString += ", pole_diameter_toc=" & IIf(IsNothing(ep.pole_diameter_toc), "Null", "'" & ep.pole_diameter_toc.ToString & "'")
    '    updateString += ", pole_top_diameter=" & IIf(IsNothing(ep.pole_top_diameter), "Null", "'" & ep.pole_top_diameter.ToString & "'")
    '    updateString += ", pole_bottom_diameter=" & IIf(IsNothing(ep.pole_bottom_diameter), "Null", "'" & ep.pole_bottom_diameter.ToString & "'")
    '    updateString += ", pole_section_length=" & IIf(IsNothing(ep.pole_section_length), "Null", "'" & ep.pole_section_length.ToString & "'")
    '    updateString += ", pole_taper_factor=" & IIf(IsNothing(ep.pole_taper_factor), "Null", "'" & ep.pole_taper_factor.ToString & "'")
    '    updateString += ", pole_bend_radius_override=" & IIf(IsNothing(ep.pole_bend_radius_override), "Null", "'" & ep.pole_bend_radius_override.ToString & "'")
    '    updateString += " WHERE ID=" & ep.embedded_id & vbNewLine

    '    Return updateString
    'End Function

    'Private Function UpdateDrilledPierSoilLayer(ByVal dpsl As DrilledPierSoilLayer) As String
    '    Dim updateString As String = ""

    '    updateString += "UPDATE drilled_pier_soil_layer SET "
    '    updateString += "bottom_depth=" & IIf(IsNothing(dpsl.bottom_depth), "Null", "'" & dpsl.bottom_depth.ToString & "'")
    '    updateString += ", effective_soil_density=" & IIf(IsNothing(dpsl.effective_soil_density), "Null", "'" & dpsl.effective_soil_density.ToString & "'")
    '    updateString += ", cohesion=" & IIf(IsNothing(dpsl.cohesion), "Null", "'" & dpsl.cohesion.ToString & "'")
    '    updateString += ", friction_angle=" & IIf(IsNothing(dpsl.friction_angle), "Null", "'" & dpsl.friction_angle.ToString & "'")
    '    updateString += ", skin_friction_override_comp=" & IIf(IsNothing(dpsl.skin_friction_override_comp), "Null", "'" & dpsl.skin_friction_override_comp.ToString & "'")
    '    updateString += ", skin_friction_override_uplift=" & IIf(IsNothing(dpsl.skin_friction_override_uplift), "Null", "'" & dpsl.skin_friction_override_uplift.ToString & "'")
    '    updateString += ", nominal_bearing_capacity=" & IIf(IsNothing(dpsl.nominal_bearing_capacity), "Null", "'" & dpsl.nominal_bearing_capacity.ToString & "'")
    '    updateString += ", spt_blow_count=" & IIf(IsNothing(dpsl.spt_blow_count), "Null", "'" & dpsl.spt_blow_count.ToString & "'")
    '    updateString += ", local_soil_layer_id=" & IIf(IsNothing(dpsl.local_soil_layer_id), "Null", "'" & dpsl.local_soil_layer_id.ToString & "'")
    '    updateString += " WHERE ID=" & dpsl.soil_layer_id & vbNewLine

    '    Return updateString
    'End Function

    'Private Function UpdateDrilledPierSection(ByVal dpsec As DrilledPierSection) As String
    '    Dim updateString As String = ""

    '    updateString += "UPDATE drilled_pier_section SET "
    '    updateString += "pier_diameter=" & IIf(IsNothing(dpsec.pier_diameter), "Null", "'" & dpsec.pier_diameter.ToString & "'")
    '    updateString += ", clear_cover=" & IIf(IsNothing(dpsec.clear_cover), "Null", "'" & dpsec.clear_cover.ToString & "'")
    '    updateString += ", clear_cover_rebar_cage_option=" & IIf(IsNothing(dpsec.clear_cover_rebar_cage_option), "Null", "'" & dpsec.clear_cover_rebar_cage_option.ToString & "'")
    '    updateString += ", tie_size=" & IIf(IsNothing(dpsec.tie_size), "Null", "'" & dpsec.tie_size.ToString & "'")
    '    updateString += ", tie_spacing=" & IIf(IsNothing(dpsec.tie_spacing), "Null", "'" & dpsec.tie_spacing.ToString & "'")
    '    updateString += ", bottom_elevation=" & IIf(IsNothing(dpsec.bottom_elevation), "Null", "'" & dpsec.bottom_elevation.ToString & "'")
    '    updateString += ", local_section_id=" & IIf(IsNothing(dpsec.local_section_id), "Null", "'" & dpsec.local_section_id.ToString & "'")
    '    updateString += ", local_drilled_pier_id=" & IIf(IsNothing(dpsec.rho_override), "Null", "'" & dpsec.rho_override.ToString & "'")
    '    updateString += " WHERE ID=" & dpsec.section_id & vbNewLine

    '    Return updateString
    'End Function

    'Private Function UpdateDrilledPierRebar(ByVal dpreb As DrilledPierRebar) As String
    '    Dim updateString As String = ""

    '    updateString += "UPDATE drilled_pier_rebar SET "
    '    updateString += "longitudinal_rebar_quantity=" & IIf(IsNothing(dpreb.longitudinal_rebar_quantity), "Null", "'" & dpreb.longitudinal_rebar_quantity.ToString & "'")
    '    updateString += ", longitudinal_rebar_size=" & IIf(IsNothing(dpreb.longitudinal_rebar_size), "Null", "'" & dpreb.longitudinal_rebar_size.ToString & "'")
    '    updateString += ", longitudinal_rebar_cage_diameter=" & IIf(IsNothing(dpreb.longitudinal_rebar_cage_diameter), "Null", "'" & dpreb.longitudinal_rebar_cage_diameter.ToString & "'")
    '    updateString += ", local_rebar_id=" & IIf(IsNothing(dpreb.local_rebar_id), "Null", "'" & dpreb.local_rebar_id.ToString & "'")
    '    updateString += " WHERE ID=" & dpreb.rebar_id & vbNewLine

    '    Return updateString
    'End Function

    'Private Function UpdateDrilledPierProfile(ByVal dpp As DrilledPierProfile) As String
    '    Dim updateString As String = ""

    '    updateString += "UPDATE drilled_pier_profile SET "
    '    updateString += ", reaction_position=" & IIf(IsNothing(dpp.reaction_position), "Null", "'" & dpp.reaction_position.ToString & "'")
    '    updateString += ", reaction_location=" & IIf(IsNothing(dpp.reaction_location), "Null", "'" & dpp.reaction_location.ToString & "'")
    '    updateString += ", drilled_pier_profile=" & IIf(IsNothing(dpp.drilled_pier_profile), "Null", "'" & dpp.drilled_pier_profile.ToString & "'")
    '    updateString += ", soil_profile=" & IIf(IsNothing(dpp.soil_profile), "Null", "'" & dpp.soil_profile.ToString & "'")
    '    updateString += " WHERE ID=" & dpp.profile_id & vbNewLine

    '    Return updateString
    'End Function
#End Region

#Region "General"
    Public Sub Clear()
        ExcelFilePath = ""
        GuyedAnchorBlocks.Clear()
    End Sub

    Private Function GuyedAnchorBlockSQLDataTables() As List(Of SQLParameter)
        Dim MyParameters As New List(Of SQLParameter)

        MyParameters.Add(New SQLParameter("Guyed Anchor Block General Details SQL", "Guyed Anchor Block (SELECT Details).sql"))
        MyParameters.Add(New SQLParameter("Guyed Anchor Block Soil SQL", "Guyed Anchor Block (SELECT Soil Layers).sql"))
        MyParameters.Add(New SQLParameter("Guyed Anchor Block Profiles SQL", "Guyed Anchor Block (SELECT Profile).sql"))

        Return MyParameters
    End Function

    Private Function GuyedAnchorBlockExcelDTParameters() As List(Of EXCELDTParameter)
        Dim MyParameters As New List(Of EXCELDTParameter)

        MyParameters.Add(New EXCELDTParameter("Guyed Anchor Block General Details EXCEL", "A2:AF52", "Details (ENTER)"))
        MyParameters.Add(New EXCELDTParameter("Guyed Anchor Block Soil EXCEL", "A2:L452", "Soil Layers (ENTER)"))
        MyParameters.Add(New EXCELDTParameter("Guyed Anchor Block Profiles EXCEL", "A2:G52", "Profiles (ENTER)"))

        Return MyParameters
    End Function
#End Region

End Class