Option Strict Off

Imports DevExpress.Spreadsheet
Imports System.Security.Principal

Partial Public Class DataTransfererPile

#Region "Define"
    Private NewPileWb As New Workbook
    Private prop_ExcelFilePath As String

    Public Property Piles As New List(Of Pile)
    Public Property sqlPiles As New List(Of Pile)
    Private Property PileTemplatePath As String = "C:\Users\" & Environment.UserName & "\Documents\.NET Testing\Foundations\Pile\Template\Pile Foundation (2.2.1.5).xlsm"
    Private Property PileFileType As DocumentFormat = DocumentFormat.Xlsm

    'Public Property pileDS As New DataSet
    Public Property pileDB As String
    Public Property pileID As WindowsIdentity
    Public Property ExcelFilePath() As String
        Get
            Return Me.prop_ExcelFilePath
        End Get
        Set
            Me.prop_ExcelFilePath = Value
        End Set
    End Property
#End Region

#Region "Constructors"
    Sub New()
        'Leave method empty
    End Sub

    Public Sub New(ByVal MyDataSet As DataSet, ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String, ByVal BU As String, ByVal Strucutre_ID As String)
        ds = MyDataSet
        pileID = LogOnUser
        pileDB = ActiveDatabase
        'BUNumber = BU 'Need to turn back on when connecting to dashboard. Turned off for testing. 
        'STR_ID = Strucutre_ID 'Need to turn back on when connecting to dashboard. Turned off for testing. 
    End Sub
#End Region

#Region "Load Data"

    Sub CreateSQLPiles(ByRef pileList As List(Of Pile))
        Dim refid As Integer
        Dim PileLoader As String

        'Load data to get Pile details for the existing structure model
        For Each item As SQLParameter In PileSQLDataTables()
            PileLoader = QueryBuilderFromFile(queryPath & "Pile\" & item.sqlQuery).Replace("[EXISTING MODEL]", GetExistingModelQuery())
            DoDaSQL.sqlLoader(PileLoader, item.sqlDatatable, ds, pileDB, pileID, "0")
            'If pileDS.Tables(item.sqlDatatable).Rows.Count = 0 Then Return False 'This may need adjusted since some tables can be empty
        Next

        'Custom Section to transfer data for the pile tool. Needs to be adjusted for each tool.

        For Each PileDataRow As DataRow In ds.Tables("Pile General Details SQL").Rows
            refid = CType(PileDataRow.Item("pile_id"), Integer)

            pileList.Add(New Pile(PileDataRow, refid))
        Next
    End Sub

    Public Function LoadFromEDS() As Boolean
        CreateSQLPiles(Piles)
        'Moved code to separate method, above (CreateSQLPiles) No changes were made to the code copied over
        Return True
    End Function 'Create Pile objects based on what is saved in EDS

    Public Sub LoadFromExcel()


        For Each item As EXCELDTParameter In PileExcelDTParameters()
            'Get tables from excel file 
            ds.Tables.Add(ExcelDatasourceToDataTable(GetExcelDataSource(ExcelFilePath, item.xlsSheet, item.xlsRange), item.xlsDatatable))
        Next

        'Piles.Add(New Pile(ExcelFilePath))

        '****Test Comparing Excel to EDS****
        Dim refid As Integer
        Dim PileLoader As String
        'Load data to get Pile details for the existing structure model
        For Each item As SQLParameter In PileSQLDataTables()
            PileLoader = QueryBuilderFromFile(queryPath & "Pile\" & item.sqlQuery).Replace("[EXISTING MODEL]", GetExistingModelQuery())
            DoDaSQL.sqlLoader(PileLoader, item.sqlDatatable, ds, pileDB, pileID, "0")
            'If pileDS.Tables(item.sqlDatatable).Rows.Count = 0 Then Return False 'This may need adjusted since some tables can be empty
        Next

        'Custom Section to transfer data for the pile tool. Needs to be adjusted for each tool.

        For Each PileDataRow As DataRow In ds.Tables("Pile General Details SQL").Rows
            refid = CType(PileDataRow.Item("pile_id"), Integer)

            'Piles.Add(New Pile(PileDataRow, refid))
            Piles.Add(New Pile(ExcelFilePath, PileDataRow, refid))
        Next



        'Return True
        '****Test Comparing Excel to EDS****



        'IEM 9/8/2021
        CreateSQLPiles(sqlPiles)
        If sqlPiles.Count > 0 Then
            For Each fnd As Pile In Piles
                For Each sqlfnd As Pile In sqlPiles
                    If fnd.pile_id = sqlfnd.pile_id Then
                        If CheckChanges(fnd, sqlfnd) Then
                            isModelNeeded = True
                            fndGroupNeeded = True
                            Save1Pile(fnd)
                        End If
                        Exit For
                    End If
                Next
            Next
        Else
            'Save the data because nothing exists in sql
        End If

    End Sub 'Create Pile objects based on what is coming from the excel file


#End Region

#Region "Save Data"

    Sub Save1Pile(ByVal pf As Pile)
        Dim firstOne As Boolean = True
        Dim mySoils As String = ""
        Dim myLocations As String = ""

        Dim PileSaver As String = QueryBuilderFromFile(queryPath & "Pile\Pile (IN_UP).sql")

        PileSaver = PileSaver.Replace("[BU NUMBER]", BUNumber)
        PileSaver = PileSaver.Replace("[STRUCTURE ID]", STR_ID)
        PileSaver = PileSaver.Replace("[FOUNDATION TYPE]", "Pile")
        If pf.pile_id = 0 Or IsDBNull(pf.pile_id) Then
            PileSaver = PileSaver.Replace("'[Pile ID]'", "NULL")
        Else
            PileSaver = PileSaver.Replace("[Pile ID]", pf.pile_id.ToString)
            'PileSaver = PileSaver.Replace("(SELECT * FROM TEMPORARY)", UpdatePileDetail(pf))
        End If
        PileSaver = PileSaver.Replace("[INSERT ALL PILE DETAILS]", InsertPileDetail(pf))
        PileSaver = PileSaver.Replace("[CONFIGURATION]", pf.pile_group_config.ToString)

        If pf.pile_id = 0 Or IsDBNull(pf.pile_id) Then
            If pf.pile_soil_capacity_given = False And pf.pile_shape <> "H-Pile" Then
                For Each pfsl As PileSoilLayer In pf.soil_layers
                    Dim tempSoilLayer As String = InsertPileSoilLayer(pfsl)

                    If Not firstOne Then
                        mySoils += ",(" & tempSoilLayer & ")"
                    Else
                        mySoils += "(" & tempSoilLayer & ")"
                    End If

<<<<<<< HEAD
                    For Each pf As Pile In Piles
                        'If pf.change_flag Then

                        'End If
                        Dim PileSaver As String = QueryBuilderFromFile(queryPath & "Pile\Pile (IN_UP).sql")

                        PileSaver = PileSaver.Replace("[BU NUMBER]", BUNumber)
                        PileSaver = PileSaver.Replace("[STRUCTURE ID]", STR_ID)
                        PileSaver = PileSaver.Replace("[FOUNDATION TYPE]", "Pile")
                        If pf.pile_id = 0 Or IsDBNull(pf.pile_id) Then
                            PileSaver = PileSaver.Replace("'[Pile ID]'", "NULL")
=======
                    firstOne = False
                Next 'Add Soil Layer INSERT statments
                PileSaver = PileSaver.Replace("([INSERT ALL SOIL LAYERS])", mySoils)
                firstOne = True
>>>>>>> b5d900c7dabf63871016a46826cd402d18aaf2ce
                        Else
                            PileSaver = PileSaver.Replace("INSERT INTO pile_soil_layer VALUES ([INSERT ALL SOIL LAYERS])", "--INSERT INTO pile_soil_layer VALUES ([INSERT ALL SOIL LAYERS])")
                        End If

                        If pf.pile_group_config = "Asymmetric" Then
                            'PileSaver = PileSaver.Replace("[INSERT ALL PILE LOCATIONS]", InsertPileLocation(dp.embed_details))

<<<<<<< HEAD
                            If Not firstOne Then
                                mySoils += ",(" & tempSoilLayer & ")"
                            Else
                                mySoils += "(" & tempSoilLayer & ")"
                            End If

                            firstOne = False
                    Next 'Add Soil Layer INSERT statments
                    PileSaver = PileSaver.Replace("([INSERT ALL SOIL LAYERS])", mySoils)
                    firstOne = True
                    Else
                    PileSaver = PileSaver.Replace("INSERT INTO pile_soil_layer VALUES ([INSERT ALL SOIL LAYERS])", "--INSERT INTO pile_soil_layer VALUES ([INSERT ALL SOIL LAYERS])")
                End If

            If pf.pile_group_config = "Asymmetric" Then
                'PileSaver = PileSaver.Replace("[INSERT ALL PILE LOCATIONS]", InsertPileLocation(dp.embed_details))

                For Each pfpl As PileLocation In pf.pile_locations
                    Dim tempLocation As String = InsertPileLocation(pfpl)

                    If Not firstOne Then
                        myLocations += ",(" & tempLocation & ")"
                    Else
                        myLocations += "(" & tempLocation & ")"
                    End If

                    firstOne = False
                Next
                PileSaver = PileSaver.Replace("([INSERT ALL PILE LOCATIONS])", myLocations)
            Else
                PileSaver = PileSaver.Replace("BEGIN IF @IsCONFIG = 'Asymmetric'", "--BEGIN IF @IsCONFIG = 'Asymmetric'")
                PileSaver = PileSaver.Replace("INSERT INTO pile_location VALUES ([INSERT ALL PILE LOCATIONS]) End", "--INSERT INTO pile_location VALUES ([INSERT ALL PILE LOCATIONS]) End")
            End If 'Add Embedded Pole INSERT Statment

            mySoils = ""
            myLocations = ""

        Else

            PileSaver = PileSaver.Replace("BEGIN IF @IsCONFIG = 'Asymmetric'", "--BEGIN IF @IsCONFIG = 'Asymmetric'")
            PileSaver = PileSaver.Replace("INSERT INTO pile_soil_layer VALUES ([INSERT ALL SOIL LAYERS])", "--INSERT INTO pile_soil_layer VALUES ([INSERT ALL SOIL LAYERS])")
            PileSaver = PileSaver.Replace("INSERT INTO pile_location VALUES ([INSERT ALL PILE LOCATIONS]) End", "--INSERT INTO pile_location VALUES ([INSERT ALL PILE LOCATIONS]) End")
=======
                For Each pfpl As PileLocation In pf.pile_locations
                    Dim tempLocation As String = InsertPileLocation(pfpl)

                    If Not firstOne Then
                        myLocations += ",(" & tempLocation & ")"
                    Else
                        myLocations += "(" & tempLocation & ")"
                    End If

                    firstOne = False
                Next
                PileSaver = PileSaver.Replace("([INSERT ALL PILE LOCATIONS])", myLocations)
            Else
                PileSaver = PileSaver.Replace("BEGIN IF @IsCONFIG = 'Asymmetric'", "--BEGIN IF @IsCONFIG = 'Asymmetric'")
                PileSaver = PileSaver.Replace("INSERT INTO pile_location VALUES ([INSERT ALL PILE LOCATIONS]) End", "--INSERT INTO pile_location VALUES ([INSERT ALL PILE LOCATIONS]) End")
            End If 'Add Embedded Pole INSERT Statment

            mySoils = ""
            myLocations = ""

        Else
>>>>>>> b5d900c7dabf63871016a46826cd402d18aaf2ce

            PileSaver = PileSaver.Replace("BEGIN IF @IsCONFIG = 'Asymmetric'", "--BEGIN IF @IsCONFIG = 'Asymmetric'")
            PileSaver = PileSaver.Replace("INSERT INTO pile_soil_layer VALUES ([INSERT ALL SOIL LAYERS])", "--INSERT INTO pile_soil_layer VALUES ([INSERT ALL SOIL LAYERS])")
            PileSaver = PileSaver.Replace("INSERT INTO pile_location VALUES ([INSERT ALL PILE LOCATIONS]) End", "--INSERT INTO pile_location VALUES ([INSERT ALL PILE LOCATIONS]) End")

            Dim tempUpdater As String = ""
            tempUpdater += UpdatePileDetail(pf)

            If pf.pile_soil_capacity_given = False And pf.pile_shape <> "H-Pile" Then
                For Each pfsl As PileSoilLayer In pf.soil_layers
                    If pfsl.soil_layer_id = 0 Or IsDBNull(pfsl.soil_layer_id) Then
                        tempUpdater += "INSERT INTO pile_soil_layer VALUES (" & InsertPileSoilLayer(pfsl) & ") " & vbNewLine
                    Else
                        tempUpdater += UpdatePileSoilLayer(pfsl)
                    End If
                Next
            End If

            'PileSaver = PileSaver.Replace("(SELECT * FROM TEMPORARY)", tempUpdater)

            'End If

<<<<<<< HEAD
            For Each pfpl As PileLocation In pf.pile_locations
                If pfpl.location_id = 0 Or IsDBNull(pfpl.location_id) Then
                    tempUpdater += "INSERT INTO pile_location VALUES (" & InsertPileLocation(pfpl) & ") " & vbNewLine
                Else
                    tempUpdater += UpdatePileLocation(pfpl)
                End If
            Next




            '    'If pfpl.location_id = 0 Or IsDBNull(dp.embed_details.embedded_id) Then
            '    'tempUpdater += "BEGIN INSERT INTO embedded_pole_details OUTPUT INSERTED.ID INTO @EmbeddedPole VALUES (" & InsertDrilledPierEmbed(dp.embed_details) & ") " & vbNewLine & " SELECT @EmbedID=EmbedID FROM @EmbeddedPole"
            '    For Each pfpl As PileLocation In pf.pile_locations
            '        tempUpdater += "INSERT INTO pile_location VALUES (" & InsertPileLocation(pfpl) & ") " & vbNewLine
            '    Next
            '    tempUpdater += " END " & vbNewLine
            'Else
            '    tempUpdater += UpdateDrilledPierEmbed(dp.embed_details)
            '    For Each esec As DrilledPierEmbedSection In dp.embed_details.sections
            '        If esec.section_id = 0 Or IsDBNull(esec.section_id) Then
            '            tempUpdater += "INSERT INTO embedded_pole_section VALUES (" & InsertDrilledPierEmbedSection(esec).Replace("@EmbedID", dp.embed_details.embedded_id.ToString) & ") " & vbNewLine
            '        Else
            '            tempUpdater += UpdateDrilledPierEmbedSection(esec)
            '        End If
            '    Next
            '    'End If
        End If

        PileSaver = PileSaver.Replace("(SELECT * FROM TEMPORARY)", tempUpdater)

        End If

        sqlSender(PileSaver, pileDB, pileID, "0")
        Next
=======
            If pf.pile_group_config = "Asymmetric" Then

                For Each pfpl As PileLocation In pf.pile_locations
                    If pfpl.location_id = 0 Or IsDBNull(pfpl.location_id) Then
                        tempUpdater += "INSERT INTO pile_location VALUES (" & InsertPileLocation(pfpl) & ") " & vbNewLine
                    Else
                        tempUpdater += UpdatePileLocation(pfpl)
                    End If
                Next




                '    'If pfpl.location_id = 0 Or IsDBNull(dp.embed_details.embedded_id) Then
                '    'tempUpdater += "BEGIN INSERT INTO embedded_pole_details OUTPUT INSERTED.ID INTO @EmbeddedPole VALUES (" & InsertDrilledPierEmbed(dp.embed_details) & ") " & vbNewLine & " SELECT @EmbedID=EmbedID FROM @EmbeddedPole"
                '    For Each pfpl As PileLocation In pf.pile_locations
                '        tempUpdater += "INSERT INTO pile_location VALUES (" & InsertPileLocation(pfpl) & ") " & vbNewLine
                '    Next
                '    tempUpdater += " END " & vbNewLine
                'Else
                '    tempUpdater += UpdateDrilledPierEmbed(dp.embed_details)
                '    For Each esec As DrilledPierEmbedSection In dp.embed_details.sections
                '        If esec.section_id = 0 Or IsDBNull(esec.section_id) Then
                '            tempUpdater += "INSERT INTO embedded_pole_section VALUES (" & InsertDrilledPierEmbedSection(esec).Replace("@EmbedID", dp.embed_details.embedded_id.ToString) & ") " & vbNewLine
                '        Else
                '            tempUpdater += UpdateDrilledPierEmbedSection(esec)
                '        End If
                '    Next
                '    'End If
            End If

            PileSaver = PileSaver.Replace("(SELECT * FROM TEMPORARY)", tempUpdater)

        End If
>>>>>>> b5d900c7dabf63871016a46826cd402d18aaf2ce

        sqlSender(PileSaver, pileDB, pileID, "0")
    End Sub

    Public Sub SaveToEDS()
        For Each pf As Pile In Piles
            Save1Pile(pf)
            'Moved code to separate method, above (Save1Pile) No changes were made to the code copied over
        Next
    End Sub

    Public Sub SaveToExcel()
        'Dim pfRow As Integer = 3
        'Dim soilRow As Integer = 4 'identify first row to copy data into Excel Sheet
        'Dim soilRow As Integer = 57 'identify first row to copy data into Excel Sheet
        'Dim locRow As Integer = 4
        'Dim locRow As Integer = 5
        'LoadNewPile() 'follows drilled pier format

        'With NewPileWb 'follows drilled pier format
        For Each pf As Pile In Piles
            Dim soilRow As Integer = 57 'identify first row to copy data into Excel Sheet
            Dim locRow As Integer = 5
            LoadNewPile() 'follows p&p format
            With NewPileWb 'follows p&p format

                If Not IsNothing(pf.pile_id) Then
                    .Worksheets("Input").Range("ID").Value = CType(pf.pile_id, Integer)
                Else .Worksheets("Input").Range("ID").ClearContents
                End If
                If Not IsNothing(pf.load_eccentricity) Then
                    .Worksheets("Input").Range("Ecc").Value = CType(pf.load_eccentricity, Double)
                Else .Worksheets("Input").Range("Ecc").ClearContents
                End If
                If Not IsNothing(pf.bolt_circle_bearing_plate_width) Then
                    .Worksheets("Input").Range("BC").Value = CType(pf.bolt_circle_bearing_plate_width, Double)
                Else .Worksheets("Input").Range("BC").ClearContents
                End If
                If Not IsNothing(pf.pile_shape) Then .Worksheets("Input").Range("D23").Value = pf.pile_shape
                If Not IsNothing(pf.pile_material) Then .Worksheets("Input").Range("D24").Value = pf.pile_material
                If Not IsNothing(pf.pile_length) Then
                    .Worksheets("Input").Range("Lpile").Value = CType(pf.pile_length, Double)
                Else .Worksheets("Input").Range("Lpile").ClearContents
                End If
                If Not IsNothing(pf.pile_diameter_width) Then
                    .Worksheets("Input").Range("D26").Value = CType(pf.pile_diameter_width, Double)
                Else .Worksheets("Input").Range("D26").ClearContents
                End If
                If Not IsNothing(pf.pile_pipe_thickness) Then
                    .Worksheets("Input").Range("D27").Value = CType(pf.pile_pipe_thickness, Double)
                Else .Worksheets("Input").Range("D27").ClearContents
                End If

                If pf.pile_soil_capacity_given = True Then
                    .Worksheets("Input").Range("D29").Value = "Yes"
                Else
                    .Worksheets("Input").Range("D29").Value = "No"
                End If

                If Not IsNothing(pf.steel_yield_strength) Then
                    .Worksheets("Input").Range("D30").Value = CType(pf.steel_yield_strength, Double)
                Else .Worksheets("Input").Range("D30").ClearContents
                End If
                If Not IsNothing(pf.pile_type_option) Then .Worksheets("Input").Range("Psize").Value = pf.pile_type_option
                If Not IsNothing(pf.rebar_quantity) Then
                    .Worksheets("Input").Range("Pquan").Value = CType(pf.rebar_quantity, Integer)
                Else .Worksheets("Input").Range("Pquan").ClearContents
                End If
                If Not IsNothing(pf.pile_group_config) Then .Worksheets("Input").Range("Config").Value = pf.pile_group_config
                If Not IsNothing(pf.foundation_depth) Then
                    .Worksheets("Input").Range("D").Value = CType(pf.foundation_depth, Double)
                Else .Worksheets("Input").Range("D").ClearContents
                End If
                If Not IsNothing(pf.pad_thickness) Then
                    .Worksheets("Input").Range("T").Value = CType(pf.pad_thickness, Double)
                Else .Worksheets("Input").Range("T").ClearContents
                End If
                If Not IsNothing(pf.pad_width_dir1) Then
                    .Worksheets("Input").Range("Wx").Value = CType(pf.pad_width_dir1, Double)
                Else .Worksheets("Input").Range("Wx").ClearContents
                End If
                If Not IsNothing(pf.pad_width_dir2) Then
                    .Worksheets("Input").Range("Wy").Value = CType(pf.pad_width_dir2, Double)
                Else .Worksheets("Input").Range("Wy").ClearContents
                End If
                If Not IsNothing(pf.pad_rebar_size_bottom) Then
                    .Worksheets("Input").Range("Spad").Value = CType(pf.pad_rebar_size_bottom, Integer)
                Else .Worksheets("Input").Range("Spad").ClearContents
                End If
                If Not IsNothing(pf.pad_rebar_size_top) Then
                    .Worksheets("Input").Range("Spad_top").Value = CType(pf.pad_rebar_size_top, Integer)
                Else .Worksheets("Input").Range("Spad_top").ClearContents
                End If
                If Not IsNothing(pf.pad_rebar_quantity_bottom_dir1) Then
                    .Worksheets("Input").Range("Mpad").Value = CType(pf.pad_rebar_quantity_bottom_dir1, Integer)
                Else .Worksheets("Input").Range("Mpad").ClearContents
                End If
                If Not IsNothing(pf.pad_rebar_quantity_top_dir1) Then
                    .Worksheets("Input").Range("Mpad_top").Value = CType(pf.pad_rebar_quantity_top_dir1, Integer)
                Else .Worksheets("Input").Range("Mpad_top").ClearContents
                End If
                If Not IsNothing(pf.pad_rebar_quantity_bottom_dir2) Then
                    .Worksheets("Input").Range("Mpad_y").Value = CType(pf.pad_rebar_quantity_bottom_dir2, Integer)
                Else .Worksheets("Input").Range("Mpad_y").ClearContents
                End If
                If Not IsNothing(pf.pad_rebar_quantity_top_dir2) Then
                    .Worksheets("Input").Range("Mpad_y_top").Value = CType(pf.pad_rebar_quantity_top_dir2, Integer)
                Else .Worksheets("Input").Range("Mpad_y_top").ClearContents
                End If
                If Not IsNothing(pf.pier_shape) Then .Worksheets("Input").Range("D57").Value = pf.pier_shape
                If Not IsNothing(pf.pier_diameter) Then
                    .Worksheets("Input").Range("di").Value = CType(pf.pier_diameter, Integer)
                Else .Worksheets("Input").Range("di").ClearContents
                End If
                If Not IsNothing(pf.extension_above_grade) Then
                    .Worksheets("Input").Range("E").Value = CType(pf.extension_above_grade, Double)
                Else .Worksheets("Input").Range("E").ClearContents
                End If
                If Not IsNothing(pf.pier_rebar_size) Then
                    .Worksheets("Input").Range("Rs").Value = CType(pf.pier_rebar_size, Integer)
                Else .Worksheets("Input").Range("Rs").ClearContents
                End If
                If Not IsNothing(pf.pier_rebar_quantity) Then
                    .Worksheets("Input").Range("mc").Value = CType(pf.pier_rebar_quantity, Integer)
                Else .Worksheets("Input").Range("mc").ClearContents
                End If
                If Not IsNothing(pf.pier_tie_size) Then
                    .Worksheets("Input").Range("St").Value = CType(pf.pier_tie_size, Integer)
                Else .Worksheets("Input").Range("St").ClearContents
                End If
                'If Not IsNothing(pf.pier_tie_quantity) Then
                '    .Worksheets("").Range("").Value = CType(pf.pier_tie_quantity, Integer)
                'Else .Worksheets("").Range("").ClearContents
                'End If
                If Not IsNothing(pf.rebar_grade) Then
                    .Worksheets("Input").Range("Fy").Value = CType(pf.rebar_grade, Double)
                Else .Worksheets("Input").Range("Fy").ClearContents
                End If
                If Not IsNothing(pf.concrete_compressive_strength) Then
                    .Worksheets("Input").Range("Fc").Value = CType(pf.concrete_compressive_strength, Double)
                Else .Worksheets("Input").Range("Fc").ClearContents
                End If
                If Not IsNothing(pf.groundwater_depth) Then
                    .Worksheets("Input").Range("D69").Value = CType(pf.groundwater_depth, Double)
                Else .Worksheets("Input").Range("D69").ClearContents
                End If
                If Not IsNothing(pf.total_soil_unit_weight) Then
                    .Worksheets("Input").Range("γsoil_dry").Value = CType(pf.total_soil_unit_weight, Double)
                Else .Worksheets("Input").Range("γsoil_dry").ClearContents
                End If
                If Not IsNothing(pf.cohesion) Then
                    .Worksheets("Input").Range("Co").Value = CType(pf.cohesion, Double)
                Else .Worksheets("Input").Range("Co").ClearContents
                End If
                If Not IsNothing(pf.friction_angle) Then
                    .Worksheets("Input").Range("ɸ").Value = CType(pf.friction_angle, Double)
                Else .Worksheets("Input").Range("ɸ").ClearContents
                End If
                If Not IsNothing(pf.neglect_depth) Then
                    .Worksheets("Input").Range("ND").Value = CType(pf.neglect_depth, Double)
                Else .Worksheets("Input").Range("ND").ClearContents
                End If
                If Not IsNothing(pf.spt_blow_count) Then
                    .Worksheets("Input").Range("N_blows").Value = CType(pf.spt_blow_count, Integer)
                Else .Worksheets("Input").Range("N_blows").ClearContents
                End If
                If Not IsNothing(pf.pile_negative_friction_force) Then
                    .Worksheets("Input").Range("Sw").Value = CType(pf.pile_negative_friction_force, Double)
                Else .Worksheets("Input").Range("Sw").ClearContents
                End If
                If Not IsNothing(pf.pile_ultimate_compression) Then
                    .Worksheets("Input").Range("K45").Value = CType(pf.pile_ultimate_compression, Double)
                Else .Worksheets("Input").Range("K45").ClearContents
                End If
                If Not IsNothing(pf.pile_ultimate_tension) Then
                    .Worksheets("Input").Range("K46").Value = CType(pf.pile_ultimate_tension, Double)
                Else .Worksheets("Input").Range("K46").ClearContents
                End If
                If Not IsNothing(pf.top_and_bottom_rebar_different) Then .Worksheets("Input").Range("Z10").Value = pf.top_and_bottom_rebar_different
                If Not IsNothing(pf.ultimate_gross_end_bearing) Then
                    .Worksheets("Input").Range("M71").Value = CType(pf.ultimate_gross_end_bearing, Double)
                Else .Worksheets("Input").Range("M71").ClearContents
                End If

                If pf.skin_friction_given = True Then
                    .Worksheets("Input").Range("N54").Value = "Yes"
                Else
                    .Worksheets("Input").Range("N54").Value = "No"
                End If

                If pf.pile_group_config = "Circular" Then
                    If Not IsNothing(pf.pile_quantity_circular) Then
                        .Worksheets("Input").Range("D36").Value = CType(pf.pile_quantity_circular, Integer)
                    Else .Worksheets("Input").Range("D36").ClearContents
                    End If
                    If Not IsNothing(pf.group_diameter_circular) Then
                        .Worksheets("Input").Range("D37").Value = CType(pf.group_diameter_circular, Double)
                    Else .Worksheets("Input").Range("D37").ClearContents
                    End If
                End If

                If pf.pile_group_config = "Rectangular" Then
                    If Not IsNothing(pf.pile_column_quantity) Then
                        .Worksheets("Input").Range("D36").Value = CType(pf.pile_column_quantity, Integer)
                    Else .Worksheets("Input").Range("D36").ClearContents
                    End If
                    If Not IsNothing(pf.pile_row_quantity) Then
                        .Worksheets("Input").Range("D37").Value = CType(pf.pile_row_quantity, Integer)
                    Else .Worksheets("Input").Range("D37").ClearContents
                    End If
                End If

                If Not IsNothing(pf.pile_columns_spacing) Then
                    .Worksheets("Input").Range("D38").Value = CType(pf.pile_columns_spacing, Double)
                Else .Worksheets("Input").Range("D38").ClearContents
                End If
                If Not IsNothing(pf.pile_row_spacing) Then
                    .Worksheets("Input").Range("D39").Value = CType(pf.pile_row_spacing, Double)
                Else .Worksheets("Input").Range("D39").ClearContents
                End If

                If pf.group_efficiency_factor_given = True Then
                    .Worksheets("Input").Range("D41").Value = "Yes"
                Else
                    .Worksheets("Input").Range("D41").Value = "No"
                End If

                If Not IsNothing(pf.group_efficiency_factor) Then
                    .Worksheets("Input").Range("D42").Value = CType(pf.group_efficiency_factor, Double)
                Else .Worksheets("Input").Range("D42").ClearContents
                End If
                If Not IsNothing(pf.cap_type) Then .Worksheets("Input").Range("D45").Value = pf.cap_type
                If Not IsNothing(pf.pile_quantity_asymmetric) Then
                    .Worksheets("Moment of Inertia").Range("D10").Value = CType(pf.pile_quantity_asymmetric, Integer)
                Else .Worksheets("Moment of Inertia").Range("D10").ClearContents
                End If
                If Not IsNothing(pf.pile_spacing_min_asymmetric) Then
                    .Worksheets("Moment of Inertia").Range("D11").Value = CType(pf.pile_spacing_min_asymmetric, Double)
                Else .Worksheets("Moment of Inertia").Range("D11").ClearContents
                End If
                If Not IsNothing(pf.quantity_piles_surrounding) Then
                    .Worksheets("Moment of Inertia").Range("D12").Value = CType(pf.quantity_piles_surrounding, Integer)
                Else .Worksheets("Moment of Inertia").Range("D12").ClearContents
                End If
                If Not IsNothing(pf.pile_cap_reference) Then .Worksheets("Input").Range("G47").Value = pf.pile_cap_reference

                If pf.pile_soil_capacity_given = False And pf.pile_shape <> "H-Pile" Then
                    For Each pfSL As PileSoilLayer In pf.soil_layers

                        'If Not IsNothing(pfSL.soil_layer_id) Then
                        '    .Worksheets("SAPI").Range("J" & soilRow).Value = CType(pfSL.soil_layer_id, Integer)
                        'Else .Worksheets("SAPI").Range("J" & soilRow).ClearContents
                        'End If
                        'If Not IsNothing(pfSL.bottom_depth) Then
                        '    .Worksheets("SAPI").Range("K" & soilRow).Value = CType(pfSL.bottom_depth, Double)
                        'Else .Worksheets("SAPI").Range("K" & soilRow).ClearContents
                        'End If
                        'If Not IsNothing(pfSL.effective_soil_density) Then
                        '    .Worksheets("SAPI").Range("L" & soilRow).Value = CType(pfSL.effective_soil_density, Double)
                        'Else .Worksheets("SAPI").Range("L" & soilRow).ClearContents
                        'End If
                        'If Not IsNothing(pfSL.cohesion) Then
                        '    .Worksheets("SAPI").Range("M" & soilRow).Value = CType(pfSL.cohesion, Double)
                        'Else .Worksheets("SAPI").Range("M" & soilRow).ClearContents
                        'End If
                        'If Not IsNothing(pfSL.friction_angle) Then
                        '    .Worksheets("SAPI").Range("N" & soilRow).Value = CType(pfSL.friction_angle, Double)
                        'Else .Worksheets("SAPI").Range("N" & soilRow).ClearContents
                        'End If
                        ''If Not IsNothing(pfSL.skin_friction_override_uplift) Then
                        ''    .Worksheets("SAPI").Range("N54").Value = CType(pfSL.skin_friction_override_uplift, Double)
                        ''Else .Worksheets("SAPI").Range("N54").ClearContents
                        ''End If
                        'If Not IsNothing(pfSL.spt_blow_count) Then
                        '    .Worksheets("SAPI").Range("O" & soilRow).Value = CType(pfSL.spt_blow_count, Integer)
                        'Else .Worksheets("SAPI").Range("O" & soilRow).ClearContents
                        'End If
                        'If Not IsNothing(pfSL.ultimate_skin_friction_comp) Then
                        '    .Worksheets("SAPI").Range("P" & soilRow).Value = CType(pfSL.ultimate_skin_friction_comp, Double)
                        'Else .Worksheets("SAPI").Range("P" & soilRow).ClearContents
                        'End If
                        'If Not IsNothing(pfSL.ultimate_skin_friction_uplift) Then
                        '    .Worksheets("SAPI").Range("Q" & soilRow).Value = CType(pfSL.ultimate_skin_friction_uplift, Double)
                        'Else .Worksheets("SAPI").Range("Q" & soilRow).ClearContents
                        'End If

                        '**** Section Below eliminates workbook open events ****

                        If Not IsNothing(pfSL.soil_layer_id) Then
                            .Worksheets("SAPI").Range("J" & soilRow - 53).Value = CType(pfSL.soil_layer_id, Integer)
                            'Else .Worksheets("SAPI").Range("J" & soilRow - 53).ClearContents
                        End If
                        If Not IsNothing(pfSL.bottom_depth) Then
                            .Worksheets("Input").Range("H" & soilRow).Value = CType(pfSL.bottom_depth, Double)
                            'Else .Worksheets("Input").Range("H" & soilRow).ClearContents
                        End If
                        If Not IsNothing(pfSL.effective_soil_density) Then
                            .Worksheets("Input").Range("K" & soilRow).Value = CType(pfSL.effective_soil_density, Double)
                            'Else .Worksheets("Input").Range("K" & soilRow).ClearContents
                        End If
                        If Not IsNothing(pfSL.cohesion) Then
                            .Worksheets("Input").Range("I" & soilRow).Value = CType(pfSL.cohesion, Double)
                            'Else .Worksheets("Input").Range("I" & soilRow).ClearContents
                        End If
                        If Not IsNothing(pfSL.friction_angle) Then
                            .Worksheets("Input").Range("J" & soilRow).Value = CType(pfSL.friction_angle, Double)
                            'Else .Worksheets("Input").Range("J" & soilRow).ClearContents
                        End If
                        'If Not IsNothing(pfSL.skin_friction_override_uplift) Then
                        '    .Worksheets("Input").Range("N54").Value = CType(pfSL.skin_friction_override_uplift, Double)
                        'Else .Worksheets("Input").Range("N54").ClearContents
                        'End If
                        If Not IsNothing(pfSL.spt_blow_count) Then
                            .Worksheets("Input").Range("L" & soilRow).Value = CType(pfSL.spt_blow_count, Integer)
                            'Else .Worksheets("Input").Range("L" & soilRow).ClearContents
                        End If
                        If Not IsNothing(pfSL.ultimate_skin_friction_comp) Then
                            .Worksheets("Input").Range("M" & soilRow).Value = CType(pfSL.ultimate_skin_friction_comp, Double)
                            'Else .Worksheets("Input").Range("M" & soilRow).ClearContents
                        End If
                        If Not IsNothing(pfSL.ultimate_skin_friction_uplift) Then
                            .Worksheets("Input").Range("N" & soilRow).Value = CType(pfSL.ultimate_skin_friction_uplift, Double)
                            'Else .Worksheets("Input").Range("N" & soilRow).ClearContents
                        End If
                        '******

                        soilRow += 1
                    Next
                End If

                If pf.pile_group_config = "Asymmetric" Then

                    For Each pfPL As PileLocation In pf.pile_locations
                        'If Not IsNothing(pfPL.location_id) Then
                        '    .Worksheets("SAPI").Range("W" & locRow).Value = CType(pfPL.location_id, Integer)
                        'Else .Worksheets("SAPI").Range("W" & locRow).ClearContents
                        'End If
                        'If Not IsNothing(pfPL.pile_x_coordinate) Then
                        '    .Worksheets("SAPI").Range("X" & locRow).Value = CType(pfPL.pile_x_coordinate, Double)
                        'Else .Worksheets("SAPI").Range("X" & locRow).ClearContents
                        'End If
                        'If Not IsNothing(pfPL.pile_y_coordinate) Then
                        '    .Worksheets("SAPI").Range("Y" & locRow).Value = CType(pfPL.pile_y_coordinate, Double)
                        'Else .Worksheets("SAPI").Range("Y" & locRow).ClearContents
                        'End If

                        '**** Section Below eliminates workbook open events ****
                        If Not IsNothing(pfPL.location_id) Then
                            .Worksheets("SAPI").Range("W" & locRow - 1).Value = CType(pfPL.location_id, Integer)
                        Else .Worksheets("SAPI").Range("W" & locRow - 1).ClearContents
                        End If
                        If Not IsNothing(pfPL.pile_x_coordinate) Then
                            .Worksheets("Moment of Inertia").Range("K" & locRow).Value = CType(pfPL.pile_x_coordinate, Double)
                            'Else .Worksheets("Moment of Inertia").Range("K" & locRow).ClearContents
                        End If
                        If Not IsNothing(pfPL.pile_y_coordinate) Then
                            .Worksheets("Moment of Inertia").Range("L" & locRow).Value = CType(pfPL.pile_y_coordinate, Double)
                            'Else .Worksheets("Moment of Inertia").Range("L" & locRow).ClearContents
                        End If
                        '*****

                        locRow += 1
                    Next
                End If

                'Worksheet Change Events
                'Hiding/unhiding specific tabs
                If pf.pile_group_config = "Circular" Then
                    .Worksheets("Moment of Inertia").Visible = False
                    .Worksheets("Moment of Inertia (Circle)").Visible = True
                Else
                    .Worksheets("Moment of Inertia").Visible = True
                    .Worksheets("Moment of Inertia (Circle)").Visible = False
                End If

                'Resizing Image
                Try
                    With .Worksheets("Input").Charts(0)
                        .Width = (300 / Math.Max(CType(pf.pad_width_dir1, Double), CType(pf.pad_width_dir2, Double))) * CType(pf.pad_width_dir1, Double) * 4.19 '4.19 multiplier determined through trial and error. 
                        .Height = (300 / Math.Max(CType(pf.pad_width_dir1, Double), CType(pf.pad_width_dir2, Double))) * CType(pf.pad_width_dir2, Double) * 4.19
                    End With
                Catch
                    'error handling to avoid dividing by zero
                End Try
            End With 'follows P&P format

            SaveAndClosePile() 'follows P&P format

        Next 'follows drilled pier format

        'End With 'follows drilled pier format

        'SaveAndClosePile() 'follows drilled pier format
    End Sub

    Private Sub LoadNewPile()
        NewPileWb.LoadDocument(PileTemplatePath, PileFileType)
        NewPileWb.BeginUpdate()
    End Sub

    Private Sub SaveAndClosePile()
        NewPileWb.EndUpdate()
        NewPileWb.SaveDocument(ExcelFilePath, PileFileType)
    End Sub
#End Region

#Region "SQL Insert Statements"
    Private Function InsertPileDetail(ByVal pf As Pile) As String
        Dim insertString As String = ""

        insertString += "@FndID"
        'insertString += "," & IIf(IsNothing(pf.pile_id), "Null", pf.pile_id.ToString)
        insertString += "," & IIf(IsNothing(pf.load_eccentricity), "Null", pf.load_eccentricity.ToString)
        insertString += "," & IIf(IsNothing(pf.bolt_circle_bearing_plate_width), "Null", pf.bolt_circle_bearing_plate_width.ToString)
        insertString += "," & IIf(IsNothing(pf.pile_shape), "Null", "'" & pf.pile_shape.ToString & "'")
        insertString += "," & IIf(IsNothing(pf.pile_material), "Null", "'" & pf.pile_material.ToString & "'")
        insertString += "," & IIf(IsNothing(pf.pile_length), "Null", pf.pile_length.ToString)
        insertString += "," & IIf(IsNothing(pf.pile_diameter_width), "Null", pf.pile_diameter_width.ToString)
        insertString += "," & IIf(IsNothing(pf.pile_pipe_thickness), "Null", pf.pile_pipe_thickness.ToString)
        insertString += "," & IIf(IsNothing(pf.pile_soil_capacity_given), "Null", "'" & pf.pile_soil_capacity_given.ToString & "'")
        insertString += "," & IIf(IsNothing(pf.steel_yield_strength), "Null", pf.steel_yield_strength.ToString)
        insertString += "," & IIf(IsNothing(pf.pile_type_option), "Null", "'" & pf.pile_type_option.ToString & "'")
        insertString += "," & IIf(IsNothing(pf.rebar_quantity), "Null", pf.rebar_quantity.ToString)
        insertString += "," & IIf(IsNothing(pf.pile_group_config), "Null", "'" & pf.pile_group_config.ToString & "'")
        insertString += "," & IIf(IsNothing(pf.foundation_depth), "Null", pf.foundation_depth.ToString)
        insertString += "," & IIf(IsNothing(pf.pad_thickness), "Null", pf.pad_thickness.ToString)
        insertString += "," & IIf(IsNothing(pf.pad_width_dir1), "Null", pf.pad_width_dir1.ToString)
        insertString += "," & IIf(IsNothing(pf.pad_width_dir2), "Null", pf.pad_width_dir2.ToString)
        insertString += "," & IIf(IsNothing(pf.pad_rebar_size_bottom), "Null", pf.pad_rebar_size_bottom.ToString)
        insertString += "," & IIf(IsNothing(pf.pad_rebar_size_top), "Null", pf.pad_rebar_size_top.ToString)
        insertString += "," & IIf(IsNothing(pf.pad_rebar_quantity_bottom_dir1), "Null", pf.pad_rebar_quantity_bottom_dir1.ToString)
        insertString += "," & IIf(IsNothing(pf.pad_rebar_quantity_top_dir1), "Null", pf.pad_rebar_quantity_top_dir1.ToString)
        insertString += "," & IIf(IsNothing(pf.pad_rebar_quantity_bottom_dir2), "Null", pf.pad_rebar_quantity_bottom_dir2.ToString)
        insertString += "," & IIf(IsNothing(pf.pad_rebar_quantity_top_dir2), "Null", pf.pad_rebar_quantity_top_dir2.ToString)
        insertString += "," & IIf(IsNothing(pf.pier_shape), "Null", "'" & pf.pier_shape.ToString & "'")
        insertString += "," & IIf(IsNothing(pf.pier_diameter), "Null", pf.pier_diameter.ToString)
        insertString += "," & IIf(IsNothing(pf.extension_above_grade), "Null", pf.extension_above_grade.ToString)
        insertString += "," & IIf(IsNothing(pf.pier_rebar_size), "Null", pf.pier_rebar_size.ToString)
        insertString += "," & IIf(IsNothing(pf.pier_rebar_quantity), "Null", pf.pier_rebar_quantity.ToString)
        insertString += "," & IIf(IsNothing(pf.pier_tie_size), "Null", pf.pier_tie_size.ToString)
        'insertString += "," & IIf(IsNothing(pf.pier_tie_quantity), "Null", pf.pier_tie_quantity.ToString)
        insertString += "," & IIf(IsNothing(pf.rebar_grade), "Null", pf.rebar_grade.ToString)
        insertString += "," & IIf(IsNothing(pf.concrete_compressive_strength), "Null", pf.concrete_compressive_strength.ToString)
        insertString += "," & IIf(IsNothing(pf.groundwater_depth), "Null", pf.groundwater_depth.ToString)
        insertString += "," & IIf(IsNothing(pf.total_soil_unit_weight), "Null", pf.total_soil_unit_weight.ToString)
        insertString += "," & IIf(IsNothing(pf.cohesion), "Null", pf.cohesion.ToString)
        insertString += "," & IIf(IsNothing(pf.friction_angle), "Null", pf.friction_angle.ToString)
        insertString += "," & IIf(IsNothing(pf.neglect_depth), "Null", pf.neglect_depth.ToString)
        insertString += "," & IIf(IsNothing(pf.spt_blow_count), "Null", pf.spt_blow_count.ToString)
        insertString += "," & IIf(IsNothing(pf.pile_negative_friction_force), "Null", pf.pile_negative_friction_force.ToString)
        insertString += "," & IIf(IsNothing(pf.pile_ultimate_compression), "Null", pf.pile_ultimate_compression.ToString)
        insertString += "," & IIf(IsNothing(pf.pile_ultimate_tension), "Null", pf.pile_ultimate_tension.ToString)
        insertString += "," & IIf(IsNothing(pf.top_and_bottom_rebar_different), "Null", "'" & pf.top_and_bottom_rebar_different.ToString & "'")
        insertString += "," & IIf(IsNothing(pf.ultimate_gross_end_bearing), "Null", pf.ultimate_gross_end_bearing.ToString)
        insertString += "," & IIf(IsNothing(pf.skin_friction_given), "Null", "'" & pf.skin_friction_given.ToString & "'")
        insertString += "," & IIf(IsNothing(pf.pile_quantity_circular), "Null", pf.pile_quantity_circular.ToString)
        insertString += "," & IIf(IsNothing(pf.group_diameter_circular), "Null", pf.group_diameter_circular.ToString)
        insertString += "," & IIf(IsNothing(pf.pile_column_quantity), "Null", pf.pile_column_quantity.ToString)
        insertString += "," & IIf(IsNothing(pf.pile_row_quantity), "Null", pf.pile_row_quantity.ToString)
        insertString += "," & IIf(IsNothing(pf.pile_columns_spacing), "Null", pf.pile_columns_spacing.ToString)
        insertString += "," & IIf(IsNothing(pf.pile_row_spacing), "Null", pf.pile_row_spacing.ToString)
        insertString += "," & IIf(IsNothing(pf.group_efficiency_factor_given), "Null", "'" & pf.group_efficiency_factor_given.ToString & "'")
        insertString += "," & IIf(IsNothing(pf.group_efficiency_factor), "Null", pf.group_efficiency_factor.ToString)
        insertString += "," & IIf(IsNothing(pf.cap_type), "Null", "'" & pf.cap_type.ToString & "'")
        insertString += "," & IIf(IsNothing(pf.pile_quantity_asymmetric), "Null", pf.pile_quantity_asymmetric.ToString)
        insertString += "," & IIf(IsNothing(pf.pile_spacing_min_asymmetric), "Null", pf.pile_spacing_min_asymmetric.ToString)
        insertString += "," & IIf(IsNothing(pf.quantity_piles_surrounding), "Null", pf.quantity_piles_surrounding.ToString)
        insertString += "," & IIf(IsNothing(pf.pile_cap_reference), "Null", "'" & pf.pile_cap_reference.ToString & "'")

        Return insertString
    End Function

    Private Function InsertPileSoilLayer(ByVal pfsl As PileSoilLayer) As String
        Dim insertString As String = ""

        insertString += "@PID"
        'insertString += "," & IIf(IsNothing(pfsl.soil_layer_id), "Null", pfsl.soil_layer_id.ToString)
        insertString += "," & IIf(IsNothing(pfsl.bottom_depth), "Null", pfsl.bottom_depth.ToString)
        insertString += "," & IIf(IsNothing(pfsl.effective_soil_density), "Null", pfsl.effective_soil_density.ToString)
        insertString += "," & IIf(IsNothing(pfsl.cohesion), "Null", pfsl.cohesion.ToString)
        insertString += "," & IIf(IsNothing(pfsl.friction_angle), "Null", pfsl.friction_angle.ToString)
        'insertString += "," & IIf(IsNothing(pfsl.skin_friction_override_uplift), "Null", pfsl.skin_friction_override_uplift.ToString)
        insertString += "," & IIf(IsNothing(pfsl.spt_blow_count), "Null", pfsl.spt_blow_count.ToString)
        insertString += "," & IIf(IsNothing(pfsl.ultimate_skin_friction_comp), "Null", pfsl.ultimate_skin_friction_comp.ToString)
        insertString += "," & IIf(IsNothing(pfsl.ultimate_skin_friction_uplift), "Null", pfsl.ultimate_skin_friction_uplift.ToString)

        Return insertString
    End Function

    Private Function InsertPileLocation(ByVal pfpl As PileLocation) As String
        Dim insertString As String = ""

        insertString += "@PID"
        'insertString += "," & IIf(IsNothing(pfpl.location_id), "Null", pfpl.location_id.ToString)
        insertString += "," & IIf(IsNothing(pfpl.pile_x_coordinate), "Null", pfpl.pile_x_coordinate.ToString)
        insertString += "," & IIf(IsNothing(pfpl.pile_y_coordinate), "Null", pfpl.pile_y_coordinate.ToString)

        Return insertString
    End Function


#End Region

#Region "SQL Update Statements"
    Private Function UpdatePileDetail(ByVal pf As Pile) As String
        Dim updateString As String = ""

        updateString += "UPDATE pile_details SET "
        'updateString += ", pile_id=" & IIf(IsNothing(pf.pile_id), "Null", pf.pile_id.ToString)
        updateString += " load_eccentricity=" & IIf(IsNothing(pf.load_eccentricity), "Null", pf.load_eccentricity.ToString)
        updateString += ", bolt_circle_bearing_plate_width=" & IIf(IsNothing(pf.bolt_circle_bearing_plate_width), "Null", pf.bolt_circle_bearing_plate_width.ToString)
        updateString += ", pile_shape=" & IIf(IsNothing(pf.pile_shape), "Null", "'" & pf.pile_shape.ToString & "'")
        updateString += ", pile_material=" & IIf(IsNothing(pf.pile_material), "Null", "'" & pf.pile_material.ToString & "'")
        updateString += ", pile_length=" & IIf(IsNothing(pf.pile_length), "Null", pf.pile_length.ToString)
        updateString += ", pile_diameter_width=" & IIf(IsNothing(pf.pile_diameter_width), "Null", pf.pile_diameter_width.ToString)
        updateString += ", pile_pipe_thickness=" & IIf(IsNothing(pf.pile_pipe_thickness), "Null", pf.pile_pipe_thickness.ToString)
        updateString += ", pile_soil_capacity_given=" & IIf(IsNothing(pf.pile_soil_capacity_given), "Null", "'" & pf.pile_soil_capacity_given.ToString & "'")
        updateString += ", steel_yield_strength=" & IIf(IsNothing(pf.steel_yield_strength), "Null", pf.steel_yield_strength.ToString)
        updateString += ", pile_type_option=" & IIf(IsNothing(pf.pile_type_option), "Null", "'" & pf.pile_type_option.ToString & "'")
        updateString += ", rebar_quantity=" & IIf(IsNothing(pf.rebar_quantity), "Null", pf.rebar_quantity.ToString)
        updateString += ", pile_group_config=" & IIf(IsNothing(pf.pile_group_config), "Null", "'" & pf.pile_group_config.ToString & "'")
        updateString += ", foundation_depth=" & IIf(IsNothing(pf.foundation_depth), "Null", pf.foundation_depth.ToString)
        updateString += ", pad_thickness=" & IIf(IsNothing(pf.pad_thickness), "Null", pf.pad_thickness.ToString)
        updateString += ", pad_width_dir1=" & IIf(IsNothing(pf.pad_width_dir1), "Null", pf.pad_width_dir1.ToString)
        updateString += ", pad_width_dir2=" & IIf(IsNothing(pf.pad_width_dir2), "Null", pf.pad_width_dir2.ToString)
        updateString += ", pad_rebar_size_bottom=" & IIf(IsNothing(pf.pad_rebar_size_bottom), "Null", pf.pad_rebar_size_bottom.ToString)
        updateString += ", pad_rebar_size_top=" & IIf(IsNothing(pf.pad_rebar_size_top), "Null", pf.pad_rebar_size_top.ToString)
        updateString += ", pad_rebar_quantity_bottom_dir1=" & IIf(IsNothing(pf.pad_rebar_quantity_bottom_dir1), "Null", pf.pad_rebar_quantity_bottom_dir1.ToString)
        updateString += ", pad_rebar_quantity_top_dir1=" & IIf(IsNothing(pf.pad_rebar_quantity_top_dir1), "Null", pf.pad_rebar_quantity_top_dir1.ToString)
        updateString += ", pad_rebar_quantity_bottom_dir2=" & IIf(IsNothing(pf.pad_rebar_quantity_bottom_dir2), "Null", pf.pad_rebar_quantity_bottom_dir2.ToString)
        updateString += ", pad_rebar_quantity_top_dir2=" & IIf(IsNothing(pf.pad_rebar_quantity_top_dir2), "Null", pf.pad_rebar_quantity_top_dir2.ToString)
        updateString += ", pier_shape=" & IIf(IsNothing(pf.pier_shape), "Null", "'" & pf.pier_shape.ToString & "'")
        updateString += ", pier_diameter=" & IIf(IsNothing(pf.pier_diameter), "Null", pf.pier_diameter.ToString)
        updateString += ", extension_above_grade=" & IIf(IsNothing(pf.extension_above_grade), "Null", pf.extension_above_grade.ToString)
        updateString += ", pier_rebar_size=" & IIf(IsNothing(pf.pier_rebar_size), "Null", pf.pier_rebar_size.ToString)
        updateString += ", pier_rebar_quantity=" & IIf(IsNothing(pf.pier_rebar_quantity), "Null", pf.pier_rebar_quantity.ToString)
        updateString += ", pier_tie_size=" & IIf(IsNothing(pf.pier_tie_size), "Null", pf.pier_tie_size.ToString)
        'updateString += ", pier_tie_quantity=" & IIf(IsNothing(pf.pier_tie_quantity), "Null", pf.pier_tie_quantity.ToString)
        updateString += ", rebar_grade=" & IIf(IsNothing(pf.rebar_grade), "Null", pf.rebar_grade.ToString)
        updateString += ", concrete_compressive_strength=" & IIf(IsNothing(pf.concrete_compressive_strength), "Null", pf.concrete_compressive_strength.ToString)
        updateString += ", groundwater_depth=" & IIf(IsNothing(pf.groundwater_depth), "Null", pf.groundwater_depth.ToString)
        updateString += ", total_soil_unit_weight=" & IIf(IsNothing(pf.total_soil_unit_weight), "Null", pf.total_soil_unit_weight.ToString)
        updateString += ", cohesion=" & IIf(IsNothing(pf.cohesion), "Null", pf.cohesion.ToString)
        updateString += ", friction_angle=" & IIf(IsNothing(pf.friction_angle), "Null", pf.friction_angle.ToString)
        updateString += ", neglect_depth=" & IIf(IsNothing(pf.neglect_depth), "Null", pf.neglect_depth.ToString)
        updateString += ", spt_blow_count=" & IIf(IsNothing(pf.spt_blow_count), "Null", pf.spt_blow_count.ToString)
        updateString += ", pile_negative_friction_force=" & IIf(IsNothing(pf.pile_negative_friction_force), "Null", pf.pile_negative_friction_force.ToString)
        updateString += ", pile_ultimate_compression=" & IIf(IsNothing(pf.pile_ultimate_compression), "Null", pf.pile_ultimate_compression.ToString)
        updateString += ", pile_ultimate_tension=" & IIf(IsNothing(pf.pile_ultimate_tension), "Null", pf.pile_ultimate_tension.ToString)
        updateString += ", top_and_bottom_rebar_different=" & IIf(IsNothing(pf.top_and_bottom_rebar_different), "Null", "'" & pf.top_and_bottom_rebar_different.ToString & "'")
        updateString += ", ultimate_gross_end_bearing=" & IIf(IsNothing(pf.ultimate_gross_end_bearing), "Null", pf.ultimate_gross_end_bearing.ToString)
        updateString += ", skin_friction_given=" & IIf(IsNothing(pf.skin_friction_given), "Null", "'" & pf.skin_friction_given.ToString & "'")
        updateString += ", pile_quantity_circular=" & IIf(IsNothing(pf.pile_quantity_circular), "Null", pf.pile_quantity_circular.ToString)
        updateString += ", group_diameter_circular=" & IIf(IsNothing(pf.group_diameter_circular), "Null", pf.group_diameter_circular.ToString)
        updateString += ", pile_column_quantity=" & IIf(IsNothing(pf.pile_column_quantity), "Null", pf.pile_column_quantity.ToString)
        updateString += ", pile_row_quantity=" & IIf(IsNothing(pf.pile_row_quantity), "Null", pf.pile_row_quantity.ToString)
        updateString += ", pile_columns_spacing=" & IIf(IsNothing(pf.pile_columns_spacing), "Null", pf.pile_columns_spacing.ToString)
        updateString += ", pile_row_spacing=" & IIf(IsNothing(pf.pile_row_spacing), "Null", pf.pile_row_spacing.ToString)
        updateString += ", group_efficiency_factor_given=" & IIf(IsNothing(pf.group_efficiency_factor_given), "Null", "'" & pf.group_efficiency_factor_given.ToString & "'")
        updateString += ", group_efficiency_factor=" & IIf(IsNothing(pf.group_efficiency_factor), "Null", pf.group_efficiency_factor.ToString)
        updateString += ", cap_type=" & IIf(IsNothing(pf.cap_type), "Null", "'" & pf.cap_type.ToString & "'")
        updateString += ", pile_quantity_asymmetric=" & IIf(IsNothing(pf.pile_quantity_asymmetric), "Null", pf.pile_quantity_asymmetric.ToString)
        updateString += ", pile_spacing_min_asymmetric=" & IIf(IsNothing(pf.pile_spacing_min_asymmetric), "Null", pf.pile_spacing_min_asymmetric.ToString)
        updateString += ", quantity_piles_surrounding=" & IIf(IsNothing(pf.quantity_piles_surrounding), "Null", pf.quantity_piles_surrounding.ToString)
        updateString += ", pile_cap_reference=" & IIf(IsNothing(pf.pile_cap_reference), "Null", "'" & pf.pile_cap_reference.ToString & "'")
        updateString += " WHERE ID = " & pf.pile_id.ToString

        Return updateString

    End Function

    Private Function UpdatePileSoilLayer(ByVal pfsl As PileSoilLayer) As String
        Dim updateString As String = ""

        updateString += "UPDATE pile_soil_layer SET "
        'updateString += " soil_layer_id=" & IIf(IsNothing(pfsl.soil_layer_id), "Null", pfsl.soil_layer_id.ToString)
        updateString += " bottom_depth=" & IIf(IsNothing(pfsl.bottom_depth), "Null", pfsl.bottom_depth.ToString)
        updateString += ", effective_soil_density=" & IIf(IsNothing(pfsl.effective_soil_density), "Null", pfsl.effective_soil_density.ToString)
        updateString += ", cohesion=" & IIf(IsNothing(pfsl.cohesion), "Null", pfsl.cohesion.ToString)
        updateString += ", friction_angle=" & IIf(IsNothing(pfsl.friction_angle), "Null", pfsl.friction_angle.ToString)
        'updateString += ", skin_friction_override_uplift=" & IIf(IsNothing(pfsl.skin_friction_override_uplift), "Null", pfsl.skin_friction_override_uplift.ToString)
        updateString += ", spt_blow_count=" & IIf(IsNothing(pfsl.spt_blow_count), "Null", pfsl.spt_blow_count.ToString)
        updateString += ", ultimate_skin_friction_comp=" & IIf(IsNothing(pfsl.ultimate_skin_friction_comp), "Null", pfsl.ultimate_skin_friction_comp.ToString)
        updateString += ", ultimate_skin_friction_uplift=" & IIf(IsNothing(pfsl.ultimate_skin_friction_uplift), "Null", pfsl.ultimate_skin_friction_uplift.ToString)
        updateString += " WHERE ID = " & pfsl.soil_layer_id.ToString

        Return updateString
    End Function

    Private Function UpdatePileLocation(ByVal pfpl As PileLocation) As String
        Dim updateString As String = ""

        updateString += "UPDATE pile_location SET "
        'updateString += " location_id=" & IIf(IsNothing(pfpl.location_id), "Null", pfpl.location_id.ToString)
        updateString += " pile_x_coordinate=" & IIf(IsNothing(pfpl.pile_x_coordinate), "Null", pfpl.pile_x_coordinate.ToString)
        updateString += ", pile_y_coordinate=" & IIf(IsNothing(pfpl.pile_y_coordinate), "Null", pfpl.pile_y_coordinate.ToString)
        updateString += " WHERE ID = " & pfpl.location_id.ToString

        Return updateString
    End Function

#End Region

#Region "General"
    Public Sub Clear()
        ExcelFilePath = ""
        Piles.Clear()
    End Sub

    Private Function PileSQLDataTables() As List(Of SQLParameter)
        Dim MyParameters As New List(Of SQLParameter)

        MyParameters.Add(New SQLParameter("Pile General Details SQL", "Pile (SELECT Details).sql"))
        MyParameters.Add(New SQLParameter("Pile Soil SQL", "Pile (SELECT Soil Layers).sql"))
        MyParameters.Add(New SQLParameter("Pile Location SQL", "Pile (SELECT Location).sql"))

        Return MyParameters
    End Function

    Private Function PileExcelDTParameters() As List(Of EXCELDTParameter)
        Dim MyParameters As New List(Of EXCELDTParameter)

        MyParameters.Add(New EXCELDTParameter("Pile Soil EXCEL", "A3:H17", "SAPI"))
        MyParameters.Add(New EXCELDTParameter("Pile Location EXCEL", "S3:U103", "SAPI"))

        'MyParameters.Add(New EXCELDTParameter("Drilled Pier General Details EXCEL", "A2:K1000", "Details (SAPI)"))
        'MyParameters.Add(New EXCELDTParameter("Drilled Pier Section EXCEL", "A2:N1000", "Sections (SAPI)"))
        'MyParameters.Add(New EXCELDTParameter("Drilled Pier Rebar EXCEL", "A2:I1000", "Rebar (SAPI)"))
        'MyParameters.Add(New EXCELDTParameter("Drilled Pier Soil EXCEL", "A2:L1000", "Soil Layers (SAPI)"))
        'MyParameters.Add(New EXCELDTParameter("Belled Details EXCEL", "A2:M1000", "Belled (SAPI)"))
        'MyParameters.Add(New EXCELDTParameter("Embedded Details EXCEL", "A2:O1000", "Embedded (SAPI)"))
        'MyParameters.Add(New EXCELDTParameter("Embedded Section EXCEL", "A2:D1000", "Embedded Section (SAPI)"))

        Return MyParameters
    End Function

#End Region

#Region "IEM"
    Private changeDt As New DataTable
    Private changeList As New List(Of AnalysisChanges)
    Function CheckChanges(ByVal xlPile As Pile, ByVal sqlPile As Pile) As Boolean
        Dim changesMade As Boolean = False


        changeDt.Columns.Add("Variable", Type.GetType("System.String"))
        changeDt.Columns.Add("New Value", Type.GetType("System.String"))
        changeDt.Columns.Add("Previuos Value", Type.GetType("System.String"))
        changeDt.Columns.Add("WO", Type.GetType("System.String"))

        'If xlPile.neglect_depth <> sqlPile.neglect_depth Then
        '    changesMade = True
        '    changeDt.Rows.Add("neglect_depth", xlPile.neglect_depth.ToString, sqlPile.neglect_depth.ToString, CurWO)
        'End If

        If Check1Change(xlPile.neglect_depth, sqlPile.neglect_depth, 1, "Neglect Depth") Then changesMade = True
        If Check1Change(xlPile.neglect_depth, sqlPile.neglect_depth, 1, "Neglect Depth") Then changesMade = True
        If Check1Change(xlPile.neglect_depth, sqlPile.neglect_depth, 1, "Neglect Depth") Then changesMade = True
        If Check1Change(xlPile.neglect_depth, sqlPile.neglect_depth, 1, "Neglect Depth") Then changesMade = True
        If Check1Change(xlPile.neglect_depth, sqlPile.neglect_depth, 1, "Neglect Depth") Then changesMade = True
        If Check1Change(xlPile.neglect_depth, sqlPile.neglect_depth, 1, "Neglect Depth") Then changesMade = True
        If Check1Change(xlPile.neglect_depth, sqlPile.neglect_depth, 1, "Neglect Depth") Then changesMade = True
        If Check1Change(xlPile.neglect_depth, sqlPile.neglect_depth, 1, "Neglect Depth") Then changesMade = True
        Return changesMade
    End Function

    Sub CreateChangeSummary(ByVal changedt As DataTable)
        'Create your string based on data in the datatable
        Dim summary As String
        Dim counter As Integer = 0

        For Each chng As AnalysisChanges In changeList
            If counter = 0 Then
                summary += chng.Name & " = " & chng.NewValue & " | Previously " & chng.PreviousValue
            Else
                summary += vbNewLine & chng.Name & " = " & chng.NewValue & " | Previously " & chng.PreviousValue
            End If

            counter += 1
        Next
    End Sub

    Function Check1Change(ByVal newValue As Object, ByVal oldvalue As Object, ByVal tolerance As Double, ByVal variable As String) As Boolean
        If newValue <> oldvalue Then
            changeDt.Rows.Add(variable, newValue, oldvalue, CurWO) 'Need to determine what we want to store in this datatable or list (Foundation Type, Foundation ID)?
            changeList.Add(New AnalysisChanges(oldvalue, newValue, variable, "Pile Foundations"))
            Return True
        End If
    End Function
#End Region

End Class


Class AnalysisChanges
    Property PreviousValue As String
    Property NewValue As String
    Property Name As String
    Property PartofDatabase As String

    Public Sub New(prev As String, Newval As String, name As String, db As String)

    End Sub
End Class