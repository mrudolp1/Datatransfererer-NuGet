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
    Public Property sqlGuyedAnchorBlocks As New List(Of GuyedAnchorBlock)
    Private Property GuyedAnchorBlockTemplatePath As String = "C:\Users\" & Environment.UserName & "\Crown Castle USA Inc\ECS - Tools\Tools\Foundations\Guy Anchor Block\SAPI\Guyed Anchor Block Foundation (4.1.0) - TEMPLATE - 11-2-2021.xlsm"
    'Private Property GuyedAnchorBlockTemplatePath As String = "C:\Users\" & Environment.UserName & "\Documents\.NET Testing\Foundations\Guyed Anchor Block\Template\Guyed Anchor Block Foundation (4.1.0) - TEMPLATE - 11-2-2021.xlsm"
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

    Sub CreateSQLGuyedAnchorBlocks(ByRef GuyedAnchorBlocks As List(Of GuyedAnchorBlock))
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

    End Sub

    Public Function LoadFromEDS() As Boolean
        CreateSQLGuyedAnchorBlocks(GuyedAnchorBlocks)
        Return True
    End Function 'Create Guyed Anchor Block objects based on what is saved in EDS

    Public Sub LoadFromExcel()


        For Each item As EXCELDTParameter In GuyedAnchorBlockExcelDTParameters()
            'Get tables from excel file 
            ds.Tables.Add(ExcelDatasourceToDataTable(GetExcelDataSource(ExcelFilePath, item.xlsSheet, item.xlsRange), item.xlsDatatable))
        Next


        Dim refID As Integer
        Dim refCol As String

        'Custom Section to transfer data for the tool. Needs to be adjusted for each tool.
        For Each GuyedAnchorBlockDataRow As DataRow In ds.Tables("Guyed Anchor Block General Details EXCEL").Rows

            refCol = "local_anchor_id"
            refID = CType(GuyedAnchorBlockDataRow.Item(refCol), Integer)

            GuyedAnchorBlocks.Add(New GuyedAnchorBlock(GuyedAnchorBlockDataRow, refID, refCol))
        Next

        'GuyedAnchorBlocks.Add(New GuyedAnchorBlock(ExcelFilePath))


        'Pull SQL data, if applicable, to compare with excel data
        CreateSQLGuyedAnchorBlocks(sqlGuyedAnchorBlocks)

        'If sqlGuyedAnchorBlocks.Count > 0 Then 'same as if checking for id in tool, if ID greater than 0.
        'For Each fnd As GuyedAnchorBlock In GuyedAnchorBlocks
        '    'If fnd.ID > 0 Then 'can skip loading SQL data if id = 0 (first time adding to EDS)
        '    If fnd.anchor_id > 0 Then 'can skip loading SQL data if id = 0 (first time adding to EDS)
        '            For Each sqlfnd As GuyedAnchorBlock In sqlGuyedAnchorBlocks 'MRP - UPDATES NEEDED!!! Chackchanges needs updated to apply to multiple objects within the same tool. Only want one foundation group per tool
        '                'If fnd.ID = sqlfnd.ID Then
        '                If fnd.anchor_id = sqlfnd.ID Then
        '                    If CheckChanges(fnd, sqlfnd) Then
        '                        isModelNeeded = True
        '                        isfndGroupNeeded = True
        '                        isGuyedAnchorBlockNeeded = True
        '                    End If
        '                    Exit For
        '                End If
        '            Next

        '        Else
        '            'Save the data because nothing exists in sql
        '            isModelNeeded = True
        '        isfndGroupNeeded = True
        '        isGuyedAnchorBlockNeeded = True
        '    End If
        'Next

        For Each fnd As GuyedAnchorBlock In GuyedAnchorBlocks
            Dim IDmatch As Boolean = False
            If fnd.anchor_id > 0 Then 'can skip loading SQL data if id = 0 (Either first time adding to EDS or guy location has been redefined with new profile ID)
                For Each sqlfnd As GuyedAnchorBlock In sqlGuyedAnchorBlocks
                    'If fnd.ID = sqlfnd.ID Then
                    If fnd.anchor_id = sqlfnd.ID Then
                        IDmatch = True
                        If CheckChanges(fnd, sqlfnd) Then
                            isModelNeeded = True
                            isfndGroupNeeded = True
                            isGuyedAnchorBlockNeeded = True
                        End If
                        Exit For
                    End If
                Next
                'IF ID match = False, Save the data because nothing exists in sql (could have copied tool from a different BU)
                If IDmatch = False Then
                    isModelNeeded = True
                    isfndGroupNeeded = True
                    isGuyedAnchorBlockNeeded = True
                End If

            Else
                For Each gabp As GuyedAnchorBlockProfile In fnd.anchor_profiles
                    If gabp.profile_id > 0 Then
                        'This portion checks to see if the new guy anchor ID was related to a change based on existing profiles (e.g. updated 1 guy anchor location with new inputs). If so, will report the changes made. 
                        For Each sqlfnd As GuyedAnchorBlock In sqlGuyedAnchorBlocks
                            For Each sqlgabp As GuyedAnchorBlockProfile In sqlfnd.anchor_profiles
                                If gabp.profile_id = sqlgabp.ID Then
                                    If CheckChanges(fnd, sqlfnd) Then
                                        isModelNeeded = True
                                        isfndGroupNeeded = True
                                        isGuyedAnchorBlockNeeded = True
                                    End If
                                    Exit For
                                End If
                            Next
                        Next
                    Else
                        'Save the data because nothing exists in sql
                        isModelNeeded = True
                        isfndGroupNeeded = True
                        isGuyedAnchorBlockNeeded = True
                    End If
                Next

            End If
        Next

        'Dim refID As Integer
        'Dim refCol As String


        ''Custom Section to transfer data for the tool. Needs to be adjusted for each tool.
        'For Each GuyedAnchorBlockDataRow As DataRow In ds.Tables("Guyed Anchor Block General Details EXCEL").Rows

        '    refCol = "local_anchor_id"
        '    refID = CType(GuyedAnchorBlockDataRow.Item(refCol), Integer)

        '    GuyedAnchorBlocks.Add(New GuyedAnchorBlock(GuyedAnchorBlockDataRow, refID, refCol))
        'Next
    End Sub 'Create Guyed Anchor Block  objects based on what is coming from the excel file
#End Region

#Region "Save Data"
    Sub Save1GuyedAnchorBlock(ByVal gab As GuyedAnchorBlock)

        'Dim firstOne As Boolean = True
        'Dim mySoils As String = ""
        'Dim myProfiles As String = ""

        'For Each fnd As GuyedAnchorBlock In GuyedAnchorBlocks

        Dim GuyedAnchorBlockSaver As String = QueryBuilderFromFile(queryPath & "Guyed Anchor Block\Guyed Anchor Block (IN_UP).sql")

            GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("[BU NUMBER]", BUNumber)
            GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("[STRUCTURE ID]", STR_ID)
            GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("[FOUNDATION TYPE]", "Guyed Anchor Block")
            If gab.anchor_id = 0 Or IsDBNull(gab.anchor_id) Then
                GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("'[GUYED ANCHOR BLOCK ID]'", "NULL")
            Else
                GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("'[GUYED ANCHOR BLOCK ID]'", gab.anchor_id.ToString)
            End If

        'create new information only once per tool, rather than each instance of the foundation from the tool
        If firstGuyedAnchorBlock Then

            'Determine if new model ID needs created. Shouldn't be added to all individual tools (only needs to be referenced once)
            If isModelNeeded Then
                GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("'[Model ID Needed]'", 1)
            Else
                GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("'[Model ID Needed]'", 0)
            End If

            'Determine if new foundation group ID needs created. 
            If isfndGroupNeeded Then
                GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("'[Fnd GRP ID Needed]'", 1)
            Else
                GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("'[Fnd GRP ID Needed]'", 0)
            End If

        Else

            GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("'[Model ID Needed]'", 0)
            GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("'[Fnd GRP ID Needed]'", 0)

        End If

        'Determine if new Guyed Anchor Block ID needs created
        If isGuyedAnchorBlockNeeded Then
            GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("'[GUYED ANCHOR BLOCK ID Needed]'", 1)
        Else
            GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("'[GUYED ANCHOR BLOCK ID Needed]'", 0)
        End If

        GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("[INSERT ALL GUYED ANCHOR BLOCK DETAILS]", InsertGuyedAnchorBlockDetail(gab))

        'If gab.anchor_id = 0 Or IsDBNull(gab.anchor_id) Then
        For Each gabsl As GuyedAnchorBlockSoilLayer In gab.soil_layers 'Might need to add a line of code similar to Piles
            Dim tempSoilLayer As String = InsertGuyedAnchorBlockSoilLayer(gabsl)

            If Not firstOne Then
                mySoils += ",(" & tempSoilLayer & ")"
            Else
                mySoils += "(" & tempSoilLayer & ")"
            End If

            firstOne = False
        Next 'Add Soil Layer INSERT statments
        If firstOne = False Then 'If soil layers exist, store and save soil layers.
            GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("--INSERT INTO fnd.anchor_block_soil_layer VALUES ([INSERT ALL SOIL LAYERS])", "INSERT INTO fnd.anchor_block_soil_layer VALUES ([INSERT ALL SOIL LAYERS])")
            GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("([INSERT ALL SOIL LAYERS])", mySoils)
        End If
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
        'End If

        mySoils = ""
        myProfiles = ""
        firstGuyedAnchorBlock = False

        'Else
        '    Dim tempUpdater As String = ""
        '    tempUpdater += UpdateGuyedAnchorBlockDetail(gab)

        '    'comment out soil layer insertion. Added in next step if a layer does not have an ID
        '    GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("INSERT INTO anchor_soil_layer VALUES ([INSERT ALL SOIL LAYERS])", "--INSERT INTO anchor_soil_layer VALUES ([INSERT ALL SOIL LAYERS])")

        '    For Each gabsl As GuyedAnchorBlockSoilLayer In gab.soil_layers
        '        If gabsl.soil_layer_id = 0 Or IsDBNull(gabsl.soil_layer_id) Then
        '            tempUpdater += "INSERT INTO anchor_soil_layer VALUES (" & InsertGuyedAnchorBlockSoilLayer(gabsl) & ") " & vbNewLine
        '        Else
        '            tempUpdater += UpdateGuyedAnchorBlockSoilLayer(gabsl)
        '        End If
        '    Next

        '    GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("INSERT INTO anchor_profile VALUES ([INSERT ALL GUYED ANCHOR BLOCK PROFILES])", "--INSERT INTO anchor_profile VALUES ([INSERT ALL GUYED ANCHOR BLOCK PROFILES])")
        '    For Each gabp As GuyedAnchorBlockProfile In gab.anchor_profiles
        '        If gabp.profile_id = 0 Or IsDBNull(gabp.profile_id) Then
        '            tempUpdater += "INSERT INTO anchor_profile VALUES (" & InsertGuyedAnchorBlockProfile(gabp) & ") " & vbNewLine
        '        Else
        '            tempUpdater += UpdateGuyedAnchorBlockProfile(gabp)
        '        End If
        '    Next

        '    GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("SELECT * FROM TEMPORARY", tempUpdater)
        'End If

        'GuyedAnchorBlockSaver = GuyedAnchorBlockSaver.Replace("[INSERT ALL GUYED ANCHOR BLOCK DETAILS]", InsertGuyedAnchorBlockDetail(gab))

        sqlSender(GuyedAnchorBlockSaver, gabDB, gabID, "0")

        'Next

    End Sub

    Dim firstGuyedAnchorBlock As Boolean = True
    Dim firstOne As Boolean = True
    Dim mySoils As String = ""
    Dim myProfiles As String = ""

    Public Sub SaveToEDS()
        For Each gab As GuyedAnchorBlock In GuyedAnchorBlocks
            Save1GuyedAnchorBlock(gab)
        Next
    End Sub

    Public Sub SaveToExcel()
        Dim gabRow As Integer = 3
        Dim soilRow As Integer = 3
        Dim profileRow As Integer = 3
        Dim summaryRowStart As Integer = 11

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
                If Not IsNothing(gab.ID) Then
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
                If Not IsNothing(gab.anchor_shaft_section) Then
                    .Worksheets("Database").Range(myCol & rowStart + 15).Value = CType(gab.anchor_shaft_section, String)
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
                'Dim summaryRowStart As Integer = 10 'commented out per updating summary below. previously required when local id varied

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
                    'If Not IsNothing(gabp.local_anchor_id) Then
                    '    .Worksheets("SUMMARY").Range("D" & summaryRowStart + CType(gabp.local_anchor_id, Integer)).Value = CType(gabp.anchor_profile, Integer)
                    '    If gabp.anchor_profile = gabp.local_anchor_id Then
                    '        .Worksheets("SUMMARY").Range("G" & summaryRowStart + CType(gabp.local_anchor_id, Integer)).Value = True
                    '    Else
                    '        .Worksheets("SUMMARY").Range("G" & summaryRowStart + CType(gabp.local_anchor_id, Integer)).Value = False
                    '    End If
                    'End If
                    'If Not IsNothing(gabp.local_anchor_id) Then
                    '    .Worksheets("SUMMARY").Range("E" & summaryRowStart + CType(gabp.local_anchor_id, Integer)).Value = CType(gabp.soil_profile, Integer)
                    '    If gabp.soil_profile = gabp.local_anchor_id Then
                    '        .Worksheets("SUMMARY").Range("H" & summaryRowStart + CType(gabp.local_anchor_id, Integer)).Value = True
                    '    Else
                    '        .Worksheets("SUMMARY").Range("H" & summaryRowStart + CType(gabp.local_anchor_id, Integer)).Value = False
                    '    End If
                    'End If
                    ''        .Worksheets("SUMMARY").Range("I" & summaryRowStart + CType(dpp.reaction_position, Integer)).Value = False
                    '.Worksheets("SUMMARY").Range("I" & summaryRowStart + CType(gabp.local_anchor_id, Integer)).Value = CType(gabp.ID, Integer)

                    'SUMMARY
                    If Not IsNothing(gabp.local_anchor_id) Then
                        .Worksheets("SUMMARY").Range("D" & summaryRowStart).Value = CType(gabp.anchor_profile, Integer)
                        .Worksheets("SUMMARY").Range("E" & summaryRowStart).Value = CType(gabp.soil_profile, Integer)
                        .Worksheets("SUMMARY").Range("G" & summaryRowStart).Value = True
                        .Worksheets("SUMMARY").Range("H" & summaryRowStart).Value = True

                        'If summaryRowStart - 10 = gabp.local_anchor_id Then
                        '    If gabp.anchor_profile = gabp.local_anchor_id Then
                        '        .Worksheets("SUMMARY").Range("G" & summaryRowStart).Value = False
                        '    Else
                        '        .Worksheets("SUMMARY").Range("G" & summaryRowStart).Value = True
                        '    End If
                        '    .Worksheets("SUMMARY").Range("E" & summaryRowStart).Value = CType(gabp.soil_profile, Integer)
                        '    If gabp.soil_profile = gabp.local_anchor_id Then
                        '        .Worksheets("SUMMARY").Range("H" & summaryRowStart).Value = False
                        '    Else
                        '        .Worksheets("SUMMARY").Range("H" & summaryRowStart).Value = True
                        '    End If
                        'Else
                        '    If gabp.anchor_profile = gabp.local_anchor_id Then
                        '        .Worksheets("SUMMARY").Range("G" & summaryRowStart).Value = True
                        '        'Else
                        '        '    .Worksheets("SUMMARY").Range("G" & summaryRowStart).Value = False
                        '    End If
                        '    .Worksheets("SUMMARY").Range("E" & summaryRowStart).Value = CType(gabp.soil_profile, Integer)
                        '    If gabp.soil_profile = gabp.local_anchor_id Then
                        '        .Worksheets("SUMMARY").Range("H" & summaryRowStart).Value = True
                        '        'Else
                        '        '    .Worksheets("SUMMARY").Range("H" & summaryRowStart).Value = False
                        '    End If
                        'End If
                    End If
                    '        .Worksheets("SUMMARY").Range("I" & summaryRowStart + CType(dpp.reaction_position, Integer)).Value = False
                    .Worksheets("SUMMARY").Range("I" & summaryRowStart).Value = CType(gabp.ID, Integer)

                    profileRow += 1
                    summaryRowStart += 1

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
                Else .Worksheets("Input").Range("Fy").ClearContents
                End If
                If GuyedAnchorBlocks(0).concrete_compressive_strength.HasValue Then
                    .Worksheets("Input").Range("F\c").Value = CType(GuyedAnchorBlocks(0).concrete_compressive_strength, Double)
                Else .Worksheets("Input").Range("F\c").ClearContents
                End If
                If GuyedAnchorBlocks(0).clear_cover.HasValue Then
                    .Worksheets("Input").Range("cc").Value = CType(GuyedAnchorBlocks(0).clear_cover, Double)
                Else .Worksheets("Input").Range("cc").ClearContents
                End If
                If GuyedAnchorBlocks(0).anchor_shaft_yield_strength.HasValue Then
                    .Worksheets("Input").Range("Fy\").Value = CType(GuyedAnchorBlocks(0).anchor_shaft_yield_strength, Double)
                Else .Worksheets("Input").Range("Fy\").ClearContents
                End If
                If GuyedAnchorBlocks(0).anchor_shaft_ultimate_strength.HasValue Then
                    .Worksheets("Input").Range("Fu\").Value = CType(GuyedAnchorBlocks(0).anchor_shaft_ultimate_strength, Double)
                Else .Worksheets("Input").Range("Fu\").ClearContents
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
                If Not IsNothing(GuyedAnchorBlocks(0).anchor_shaft_section) Then .Worksheets("Input").Range("C35").Value = GuyedAnchorBlocks(0).anchor_shaft_section
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
                .Worksheets("Input").Range("CurrentLocation").Value = firstReaction

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
        NewGuyedAnchorBlockWb.Calculate()
        NewGuyedAnchorBlockWb.EndUpdate()
        NewGuyedAnchorBlockWb.SaveDocument(ExcelFilePath, GuyedAnchorBlockFileType)
    End Sub
#End Region

#Region "SQL Insert Statements"
    Private Function InsertGuyedAnchorBlockDetail(ByVal gab As GuyedAnchorBlock) As String
        Dim insertString As String = ""

        'insertString += "@FndID"
        insertString += IIf(IsNothing(gab.anchor_depth), "Null", gab.anchor_depth.ToString)
        insertString += "," & IIf(IsNothing(gab.anchor_width), "Null", gab.anchor_width.ToString)
        insertString += "," & IIf(IsNothing(gab.anchor_thickness), "Null", gab.anchor_thickness.ToString)
        insertString += "," & IIf(IsNothing(gab.anchor_length), "Null", gab.anchor_length.ToString)
        insertString += "," & IIf(IsNothing(gab.anchor_toe_width), "Null", gab.anchor_toe_width.ToString)
        insertString += "," & IIf(IsNothing(gab.anchor_top_rebar_size), "Null", gab.anchor_top_rebar_size.ToString)
        insertString += "," & IIf(IsNothing(gab.anchor_top_rebar_quantity), "Null", gab.anchor_top_rebar_quantity.ToString)
        insertString += "," & IIf(IsNothing(gab.anchor_front_rebar_size), "Null", gab.anchor_front_rebar_size.ToString)
        insertString += "," & IIf(IsNothing(gab.anchor_front_rebar_quantity), "Null", gab.anchor_front_rebar_quantity.ToString)
        insertString += "," & IIf(IsNothing(gab.anchor_stirrup_size), "Null", gab.anchor_stirrup_size.ToString)
        insertString += "," & IIf(IsNothing(gab.anchor_shaft_diameter), "Null", gab.anchor_shaft_diameter.ToString)
        insertString += "," & IIf(IsNothing(gab.anchor_shaft_quantity), "Null", gab.anchor_shaft_quantity.ToString)
        insertString += "," & IIf(IsNothing(gab.anchor_shaft_area_override), "Null", gab.anchor_shaft_area_override.ToString)
        insertString += "," & IIf(IsNothing(gab.anchor_shaft_shear_lag_factor), "Null", gab.anchor_shaft_shear_lag_factor.ToString)
        insertString += "," & IIf(IsNothing(gab.concrete_compressive_strength), "Null", gab.concrete_compressive_strength.ToString)
        insertString += "," & IIf(IsNothing(gab.clear_cover), "Null", gab.clear_cover.ToString)
        insertString += "," & IIf(IsNothing(gab.anchor_shaft_yield_strength), "Null", gab.anchor_shaft_yield_strength.ToString)
        insertString += "," & IIf(IsNothing(gab.anchor_shaft_ultimate_strength), "Null", gab.anchor_shaft_ultimate_strength.ToString)
        insertString += "," & IIf(IsNothing(gab.neglect_depth), "Null", gab.neglect_depth.ToString)
        insertString += "," & IIf(IsNothing(gab.groundwater_depth), "Null", gab.groundwater_depth.ToString)
        insertString += "," & IIf(IsNothing(gab.soil_layer_quantity), "Null", gab.soil_layer_quantity.ToString)
        insertString += "," & IIf(IsNothing(gab.tool_version), "Null", "'" & gab.tool_version.ToString & "'")
        insertString += "," & IIf(IsNothing(gab.anchor_shaft_section), "Null", "'" & gab.anchor_shaft_section.ToString & "'")
        insertString += "," & IIf(IsNothing(gab.anchor_rebar_grade), "Null", gab.anchor_rebar_grade.ToString)
        insertString += "," & IIf(IsNothing(gab.anchor_shaft_known), "Null", "'" & gab.anchor_shaft_known.ToString & "'")
        insertString += "," & IIf(IsNothing(gab.basic_soil_check), "Null", "'" & gab.basic_soil_check.ToString & "'")
        insertString += "," & IIf(IsNothing(gab.structural_check), "Null", "'" & gab.structural_check.ToString & "'")
        insertString += "," & IIf(IsNothing(gab.rebar_known), "Null", "'" & gab.rebar_known.ToString & "'")
        insertString += "," & IIf(IsNothing(gab.local_anchor_id), "Null", gab.local_anchor_id.ToString)
        insertString += "," & IIf(IsNothing(gab.local_anchor_profile), "Null", gab.local_anchor_profile.ToString)

        Return insertString
    End Function
    Private Function InsertGuyedAnchorBlockSoilLayer(ByVal gabsl As GuyedAnchorBlockSoilLayer) As String
        Dim insertString As String = ""

        insertString += "@GABID"
        insertString += "," & IIf(IsNothing(gabsl.bottom_depth), "Null", "'" & gabsl.bottom_depth.ToString & "'")
        insertString += "," & IIf(IsNothing(gabsl.effective_soil_density), "Null", "'" & gabsl.effective_soil_density.ToString & "'")
        insertString += "," & IIf(IsNothing(gabsl.cohesion), "Null", "'" & gabsl.cohesion.ToString & "'")
        insertString += "," & IIf(IsNothing(gabsl.friction_angle), "Null", "'" & gabsl.friction_angle.ToString & "'")
        insertString += "," & IIf(IsNothing(gabsl.skin_friction_override_uplift), "Null", "'" & gabsl.skin_friction_override_uplift.ToString & "'")
        insertString += "," & IIf(IsNothing(gabsl.spt_blow_count), "Null", "'" & gabsl.spt_blow_count.ToString & "'")
        insertString += "," & IIf(IsNothing(gabsl.local_soil_layer_id), "Null", "'" & gabsl.local_soil_layer_id.ToString & "'")
        insertString += "," & IIf(IsNothing(gabsl.local_soil_profile), "Null", "'" & gabsl.local_soil_profile.ToString & "'")

        Return insertString
    End Function
    Private Function InsertGuyedAnchorBlockProfile(ByVal gabp As GuyedAnchorBlockProfile) As String
        Dim insertString As String = ""

        insertString += "@GABID"
        insertString += "," & IIf(IsNothing(gabp.reaction_location), "Null", "'" & gabp.reaction_location.ToString & "'")
        insertString += "," & IIf(IsNothing(gabp.anchor_profile), "Null", "'" & gabp.anchor_profile.ToString & "'")
        insertString += "," & IIf(IsNothing(gabp.soil_profile), "Null", "'" & gabp.soil_profile.ToString & "'")
        insertString += "," & IIf(IsNothing(gabp.local_anchor_id), "Null", "'" & gabp.local_anchor_id.ToString & "'")

        Return insertString
    End Function

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

        'Remove all datatables from the main dataset
        For Each item As EXCELDTParameter In GuyedAnchorBlockExcelDTParameters()
            Try
                ds.Tables.Remove(item.xlsDatatable)
            Catch ex As Exception
            End Try
        Next

        For Each item As SQLParameter In GuyedAnchorBlockSQLDataTables()
            Try
                ds.Tables.Remove(item.sqlDatatable)
            Catch ex As Exception
            End Try
        Next
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

#Region "Check Changes"
    'Private changeDt As New DataTable
    'Private changeList As New List(Of AnalysisChanges)
    Function CheckChanges(ByVal xlGuyedAnchorBlock As GuyedAnchorBlock, ByVal sqlGuyedAnchorBlock As GuyedAnchorBlock) As Boolean
        Dim changesMade As Boolean = False

        'changeDt.Columns.Add("Variable", Type.GetType("System.String"))
        'changeDt.Columns.Add("New Value", Type.GetType("System.String"))
        'changeDt.Columns.Add("Previuos Value", Type.GetType("System.String"))
        'changeDt.Columns.Add("WO", Type.GetType("System.String"))

        'Check Details
        If Check1Change(xlGuyedAnchorBlock.anchor_depth, sqlGuyedAnchorBlock.anchor_depth, "Guyed Anchor Block", "Anchor_Depth") Then changesMade = True
        If Check1Change(xlGuyedAnchorBlock.anchor_width, sqlGuyedAnchorBlock.anchor_width, "Guyed Anchor Block", "Anchor_Width") Then changesMade = True
        If Check1Change(xlGuyedAnchorBlock.anchor_thickness, sqlGuyedAnchorBlock.anchor_thickness, "Guyed Anchor Block", "Anchor_Thickness") Then changesMade = True
        If Check1Change(xlGuyedAnchorBlock.anchor_length, sqlGuyedAnchorBlock.anchor_length, "Guyed Anchor Block", "Anchor_Length") Then changesMade = True
        If Check1Change(xlGuyedAnchorBlock.anchor_toe_width, sqlGuyedAnchorBlock.anchor_toe_width, "Guyed Anchor Block", "Anchor_Toe_Width") Then changesMade = True
        If Check1Change(xlGuyedAnchorBlock.anchor_top_rebar_size, sqlGuyedAnchorBlock.anchor_top_rebar_size, "Guyed Anchor Block", "Anchor_Top_Rebar_Size") Then changesMade = True
        If Check1Change(xlGuyedAnchorBlock.anchor_top_rebar_quantity, sqlGuyedAnchorBlock.anchor_top_rebar_quantity, "Guyed Anchor Block", "Anchor_Top_Rebar_Quantity") Then changesMade = True
        If Check1Change(xlGuyedAnchorBlock.anchor_front_rebar_size, sqlGuyedAnchorBlock.anchor_front_rebar_size, "Guyed Anchor Block", "Anchor_Front_Rebar_Size") Then changesMade = True
        If Check1Change(xlGuyedAnchorBlock.anchor_front_rebar_quantity, sqlGuyedAnchorBlock.anchor_front_rebar_quantity, "Guyed Anchor Block", "Anchor_Front_Rebar_Quantity") Then changesMade = True
        If Check1Change(xlGuyedAnchorBlock.anchor_stirrup_size, sqlGuyedAnchorBlock.anchor_stirrup_size, "Guyed Anchor Block", "Anchor_Stirrup_Size") Then changesMade = True
        If Check1Change(xlGuyedAnchorBlock.anchor_shaft_diameter, sqlGuyedAnchorBlock.anchor_shaft_diameter, "Guyed Anchor Block", "Anchor_Shaft_Diameter") Then changesMade = True
        If Check1Change(xlGuyedAnchorBlock.anchor_shaft_quantity, sqlGuyedAnchorBlock.anchor_shaft_quantity, "Guyed Anchor Block", "Anchor_Shaft_Quantity") Then changesMade = True
        If Check1Change(xlGuyedAnchorBlock.anchor_shaft_area_override, sqlGuyedAnchorBlock.anchor_shaft_area_override, "Guyed Anchor Block", "Anchor_Shaft_Area_Override") Then changesMade = True
        If Check1Change(xlGuyedAnchorBlock.anchor_shaft_shear_lag_factor, sqlGuyedAnchorBlock.anchor_shaft_shear_lag_factor, "Guyed Anchor Block", "Anchor_Shaft_Shear_Lag_Factor") Then changesMade = True
        If Check1Change(xlGuyedAnchorBlock.concrete_compressive_strength, sqlGuyedAnchorBlock.concrete_compressive_strength, "Guyed Anchor Block", "Concrete_Compressive_Strength") Then changesMade = True
        If Check1Change(xlGuyedAnchorBlock.clear_cover, sqlGuyedAnchorBlock.clear_cover, "Guyed Anchor Block", "Clear_Cover") Then changesMade = True
        If Check1Change(xlGuyedAnchorBlock.anchor_shaft_yield_strength, sqlGuyedAnchorBlock.anchor_shaft_yield_strength, "Guyed Anchor Block", "Anchor_Shaft_Yield_Strength") Then changesMade = True
        If Check1Change(xlGuyedAnchorBlock.anchor_shaft_ultimate_strength, sqlGuyedAnchorBlock.anchor_shaft_ultimate_strength, "Guyed Anchor Block", "Anchor_Shaft_Ultimate_Strength") Then changesMade = True
        If Check1Change(xlGuyedAnchorBlock.neglect_depth, sqlGuyedAnchorBlock.neglect_depth, "Guyed Anchor Block", "Neglect_Depth") Then changesMade = True
        If Check1Change(xlGuyedAnchorBlock.groundwater_depth, sqlGuyedAnchorBlock.groundwater_depth, "Guyed Anchor Block", "Groundwater_Depth") Then changesMade = True
        If Check1Change(xlGuyedAnchorBlock.soil_layer_quantity, sqlGuyedAnchorBlock.soil_layer_quantity, "Guyed Anchor Block", "Soil_Layer_Quantity") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.tool_version, sqlGuyedAnchorBlock.tool_version, 1, "Tool_Version") Then changesMade = True
        If Check1Change(xlGuyedAnchorBlock.anchor_shaft_section, sqlGuyedAnchorBlock.anchor_shaft_section, "Guyed Anchor Block", "Anchor_Shaft_Section") Then changesMade = True
        If Check1Change(xlGuyedAnchorBlock.anchor_rebar_grade, sqlGuyedAnchorBlock.anchor_rebar_grade, "Guyed Anchor Block", "Anchor_Rebar_Grade") Then changesMade = True
        If Check1Change(xlGuyedAnchorBlock.anchor_shaft_known, sqlGuyedAnchorBlock.anchor_shaft_known, "Guyed Anchor Block", "Anchor_Shaft_Known") Then changesMade = True
        If Check1Change(xlGuyedAnchorBlock.basic_soil_check, sqlGuyedAnchorBlock.basic_soil_check, "Guyed Anchor Block", "Basic_Soil_Check") Then changesMade = True
        If Check1Change(xlGuyedAnchorBlock.structural_check, sqlGuyedAnchorBlock.structural_check, "Guyed Anchor Block", "Structural_Check") Then changesMade = True
        If Check1Change(xlGuyedAnchorBlock.rebar_known, sqlGuyedAnchorBlock.rebar_known, "Guyed Anchor Block", "Rebar_Known") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.local_anchor_id, sqlGuyedAnchorBlock.local_anchor_id, 1, "Local_Anchor_Id") Then changesMade = True
        'If Check1Change(xlGuyedAnchorBlock.local_anchor_profile, sqlGuyedAnchorBlock.local_anchor_profile, 1, "Local_Anchor_Profile") Then changesMade = True

        'Check Soil Layer
        'If xlGuyedAnchorBlock.soil_layers.Count <> sqlGuyedAnchorBlock.soil_layers.Count Then changesMade = True 'If want to bypass all the checks below

        For Each gabsl As GuyedAnchorBlockSoilLayer In xlGuyedAnchorBlock.soil_layers
            For Each sqlgabsl As GuyedAnchorBlockSoilLayer In sqlGuyedAnchorBlock.soil_layers

                If gabsl.soil_layer_id = sqlgabsl.ID Then
                    If Check1Change(gabsl.bottom_depth, sqlgabsl.bottom_depth, "Guyed Anchor Block", "Bottom_Depth" & gabsl.soil_layer_id.ToString) Then changesMade = True
                    If Check1Change(gabsl.effective_soil_density, sqlgabsl.effective_soil_density, "Guyed Anchor Block", "Effective_Soil_Density" & gabsl.soil_layer_id.ToString) Then changesMade = True
                    If Check1Change(gabsl.cohesion, sqlgabsl.cohesion, "Guyed Anchor Block", "Cohesion" & gabsl.soil_layer_id.ToString) Then changesMade = True
                    If Check1Change(gabsl.friction_angle, sqlgabsl.friction_angle, "Guyed Anchor Block", "Friction_Angle" & gabsl.soil_layer_id.ToString) Then changesMade = True
                    If Check1Change(gabsl.skin_friction_override_uplift, sqlgabsl.skin_friction_override_uplift, "Guyed Anchor Block", "Ultimate_Skin_Friction_Override_Uplift" & gabsl.soil_layer_id.ToString) Then changesMade = True
                    If Check1Change(gabsl.spt_blow_count, sqlgabsl.spt_blow_count, "Guyed Anchor Block", "spt_blow_count" & gabsl.soil_layer_id.ToString) Then changesMade = True
                    'If Check1Change(gabsl.local_soil_layer_id, sqlgabsl.local_soil_layer_id, 1, "local_soil_layer_id" & gabsl.soil_layer_id.ToString) Then changesMade = True
                    'If Check1Change(gabsl.local_soil_profile, sqlgabsl.local_soil_profile, 1, "local_soil_profile" & gabsl.soil_layer_id.ToString) Then changesMade = True
                    Exit For
                End If

                If gabsl.soil_layer_id = 0 Then 'accounts for inserting new rows. additional rows won't have an ID associated to them. 
                    If Check1Change(gabsl.bottom_depth, Nothing, "Guyed Anchor Block", "Bottom_Depth" & gabsl.soil_layer_id.ToString) Then changesMade = True
                    If Check1Change(gabsl.effective_soil_density, Nothing, "Guyed Anchor Block", "Effective_Soil_Density" & gabsl.soil_layer_id.ToString) Then changesMade = True
                    If Check1Change(gabsl.cohesion, Nothing, "Guyed Anchor Block", "Cohesion" & gabsl.soil_layer_id.ToString) Then changesMade = True
                    If Check1Change(gabsl.friction_angle, Nothing, "Guyed Anchor Block", "Friction_Angle" & gabsl.soil_layer_id.ToString) Then changesMade = True
                    If Check1Change(gabsl.skin_friction_override_uplift, Nothing, "Guyed Anchor Block", "Ultimate_Skin_Friction_Override_Uplift" & gabsl.soil_layer_id.ToString) Then changesMade = True
                    If Check1Change(gabsl.spt_blow_count, Nothing, "Guyed Anchor Block", "spt_blow_count" & gabsl.soil_layer_id.ToString) Then changesMade = True
                    'If Check1Change(gabsl.local_soil_layer_id, Nothing, 1, "local_soil_layer_id" & gabsl.soil_layer_id.ToString) Then changesMade = True
                    'If Check1Change(gabsl.local_soil_profile, Nothing, 1, "local_soil_profile" & gabsl.soil_layer_id.ToString) Then changesMade = True
                    Exit For
                End If

            Next
        Next

        'Guyed Anchor Block Profiles
        For Each gabp As GuyedAnchorBlockProfile In xlGuyedAnchorBlock.anchor_profiles
            Dim profilechecked As Boolean = False 'new field added based on different iterations
            For Each sqlgabp As GuyedAnchorBlockProfile In sqlGuyedAnchorBlock.anchor_profiles

                If gabp.profile_id = sqlgabp.ID Then
                    If Check1Change(gabp.reaction_location, sqlgabp.reaction_location, "Guyed Anchor Block", "reaction_location" & gabp.profile_id.ToString) Then changesMade = True
                    If Check1Change(gabp.anchor_profile, sqlgabp.anchor_profile, "Guyed Anchor Block", "anchor_profile" & gabp.profile_id.ToString) Then changesMade = True
                    If Check1Change(gabp.soil_profile, sqlgabp.soil_profile, "Guyed Anchor Block", "soil_profile" & gabp.profile_id.ToString) Then changesMade = True
                    If Check1Change(gabp.local_anchor_id, sqlgabp.local_anchor_id, "Guyed Anchor Block", "local_anchor_id" & gabp.profile_id.ToString) Then changesMade = True
                    profilechecked = True
                    Exit For
                End If

                If gabp.profile_id = 0 Then 'accounts for inserting new rows. additional rows won't have an ID associated to them.
                    If Check1Change(gabp.reaction_location, Nothing, "Guyed Anchor Block", "reaction_location" & gabp.profile_id.ToString) Then changesMade = True
                    If Check1Change(gabp.anchor_profile, Nothing, "Guyed Anchor Block", "anchor_profile" & gabp.profile_id.ToString) Then changesMade = True
                    If Check1Change(gabp.soil_profile, Nothing, "Guyed Anchor Block", "soil_profile" & gabp.profile_id.ToString) Then changesMade = True
                    If Check1Change(gabp.local_anchor_id, Nothing, "Guyed Anchor Block", "local_anchor_id" & gabp.profile_id.ToString) Then changesMade = True
                    profilechecked = True
                    Exit For
                End If

            Next

            If gabp.profile_id > 0 And profilechecked = False Then 'User copied an existing AR id and overrode another existing location with it. Same logic as if =0 
                For Each sqlfnd As GuyedAnchorBlock In sqlGuyedAnchorBlocks
                    For Each sqlgabp As GuyedAnchorBlockProfile In sqlfnd.anchor_profiles
                        If gabp.profile_id = sqlgabp.ID Then
                            If Check1Change(gabp.reaction_location, sqlgabp.reaction_location, "Guyed Anchor Block", "reaction_location" & gabp.profile_id.ToString) Then changesMade = True
                            If Check1Change(gabp.anchor_profile, sqlgabp.anchor_profile, "Guyed Anchor Block", "anchor_profile" & gabp.profile_id.ToString) Then changesMade = True
                            If Check1Change(gabp.soil_profile, sqlgabp.soil_profile, "Guyed Anchor Block", "soil_profile" & gabp.profile_id.ToString) Then changesMade = True
                            If Check1Change(gabp.local_anchor_id, sqlgabp.local_anchor_id, "Guyed Anchor Block", "local_anchor_id" & gabp.profile_id.ToString) Then changesMade = True
                            Exit For
                        End If
                    Next
                Next
            End If

        Next

        CreateChangeSummary(changeDt) 'possible alternative to listing change summary
        Return changesMade

    End Function

    'Function CreateChangeSummary(ByVal changeDt As DataTable) As String
    '    'Sub CreateChangeSummary(ByVal changeDt As DataTable)
    '    'Create your string based on data in the datatable
    '    Dim summary As String
    '    Dim counter As Integer = 0

    '    For Each chng As AnalysisChanges In changeList
    '        If counter = 0 Then
    '            summary += chng.Name & " = " & chng.NewValue & " | Previously: " & chng.PreviousValue
    '        Else
    '            summary += vbNewLine & chng.Name & " = " & chng.NewValue & " | Previously: " & chng.PreviousValue
    '        End If

    '        counter += 1
    '    Next

    '    'write to text file
    '    'End Sub
    'End Function

    'Function Check1Change(ByVal newValue As Object, ByVal oldvalue As Object, ByVal tolerance As Double, ByVal variable As String) As Boolean
    '    If newValue <> oldvalue Then
    '        changeDt.Rows.Add(variable, newValue, oldvalue, CurWO) 'Need to determine what we want to store in this datatable or list (Foundation Type, Foundation ID)?
    '        changeList.Add(New AnalysisChanges(oldvalue, newValue, variable, "Guyed Anchor Block Foundations"))
    '        Return True
    '    ElseIf Not IsNothing(newValue) And IsNothing(oldvalue) Then 'accounts for when new rows are added. New rows from excel=0 where sql=nothing
    '        changeDt.Rows.Add(variable, newValue, oldvalue, CurWO) 'Need to determine what we want to store in this datatable or list (Foundation Type, Foundation ID)?
    '        changeList.Add(New AnalysisChanges(oldvalue, newValue, variable, "Guyed Anchor Block Foundations"))
    '        Return True
    '    ElseIf IsNothing(newValue) And Not IsNothing(oldvalue) Then 'accounts for when rows are removed. Rows from excel=nothing where sql=value
    '        changeDt.Rows.Add(variable, newValue, oldvalue, CurWO) 'Need to determine what we want to store in this datatable or list (Foundation Type, Foundation ID)?
    '        changeList.Add(New AnalysisChanges(oldvalue, newValue, variable, "Guyed Anchor Block Foundations"))
    '        Return True
    '    End If
    'End Function
#End Region

End Class