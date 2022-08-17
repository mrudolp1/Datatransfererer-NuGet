'Option Strict Off

'Imports DevExpress.Spreadsheet
'Imports System.Security.Principal

'Partial Public Class DataTransfererPile

'#Region "Define"
'    Private NewPileWb As New Workbook
'    Private prop_ExcelFilePath As String

'    Public Property Piles As New List(Of Pile)
'    Public Property sqlPiles As New List(Of Pile)
'    'Private Property PileTemplatePath As String = "C:\Users\" & Environment.UserName & "\Desktop\Pile Foundation\VB.Net Test Cases\Pile Foundation (2.2.1.6).xlsm"
'    Private Property PileTemplatePath As String = "C:\Users\" & Environment.UserName & "\Documents\.NET Testing\Foundations\Pile\Template\Pile Foundation (2.2.1.6).xlsm"
'    Private Property PileFileType As DocumentFormat = DocumentFormat.Xlsm

'    'Public Property pileDS As New DataSet
'    Public Property pileDB As String
'    Public Property pileID As WindowsIdentity
'    Public Property ExcelFilePath() As String
'        Get
'            Return Me.prop_ExcelFilePath
'        End Get
'        Set
'            Me.prop_ExcelFilePath = Value
'        End Set
'    End Property
'#End Region

'#Region "Constructors"
'    Sub New()
'        'Leave method empty
'    End Sub

'    Public Sub New(ByVal MyDataSet As DataSet, ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String, ByVal BU As String, ByVal Strucutre_ID As String)
'        ds = MyDataSet
'        pileID = LogOnUser
'        pileDB = ActiveDatabase
'        'BUNumber = BU 'Need to turn back on when connecting to dashboard. Turned off for testing. 
'        'STR_ID = Strucutre_ID 'Need to turn back on when connecting to dashboard. Turned off for testing. 
'    End Sub
'#End Region

'#Region "Load Data"

'    Sub CreateSQLPiles(ByRef pileList As List(Of Pile))
'        Dim refid As Integer
'        Dim PileLoader As String

'        'Load data to get Pile details for the existing structure model
'        For Each item As SQLParameter In PileSQLDataTables()
'            PileLoader = QueryBuilderFromFile(queryPath & "Pile\" & item.sqlQuery).Replace("[EXISTING MODEL]", GetExistingModelQuery())
'            DoDaSQL.sqlLoader(PileLoader, item.sqlDatatable, ds, pileDB, pileID, "0")
'        Next

'        'Custom Section to transfer data for the pile tool. Needs to be adjusted for each tool.
'        For Each PileDataRow As DataRow In ds.Tables("Pile General Details SQL").Rows
'            refid = CType(PileDataRow.Item("pile_id"), Integer)
'            pileList.Add(New Pile(PileDataRow, refid))
'        Next

'    End Sub

'    Public Function LoadFromEDS() As Boolean
'        CreateSQLPiles(Piles)
'        'Moved code to separate method, above (CreateSQLPiles) No changes were made to the code copied over
'        Return True
'    End Function 'Create Pile objects based on what is saved in EDS

'    Public Sub LoadFromExcel()


'        For Each item As EXCELDTParameter In PileExcelDTParameters()
'            'Get additional tables from excel file 
'            ds.Tables.Add(ExcelDatasourceToDataTable(GetExcelDataSource(ExcelFilePath, item.xlsSheet, item.xlsRange), item.xlsDatatable))
'        Next

'        Piles.Add(New Pile(ExcelFilePath))


'        'Pull SQL data, if applicable, to compare with excel data
'        CreateSQLPiles(sqlPiles)

'        'If sqlPiles.Count > 0 Then 'same as if checking for id in tool, if ID greater than 0.
'        For Each fnd As Pile In Piles
'            Dim IDmatch As Boolean = False
'            If fnd.pile_id > 0 Then 'can skip loading SQL data if id = 0 (first time adding to EDS)
'                For Each sqlfnd As Pile In sqlPiles
'                    If fnd.pile_id = sqlfnd.pile_id Then
'                        IDmatch = True
'                        If CheckChanges(fnd, sqlfnd) Then
'                            isModelNeeded = True
'                            isfndGroupNeeded = True
'                            isPileNeeded = True
'                        End If
'                        Exit For
'                    End If
'                Next
'                'IF ID match = False, Save the data because nothing exists in sql (could have copied tool from a different BU)
'                If IDmatch = False Then
'                    isModelNeeded = True
'                    isfndGroupNeeded = True
'                    isPileNeeded = True
'                End If
'            Else
'                'Save the data because nothing exists in sql
'                isModelNeeded = True
'                isfndGroupNeeded = True
'                isPileNeeded = True
'            End If
'        Next

'        'Else
'        '    'Save the data because nothing exists in sql
'        '    isModelNeeded = True
'        '    isfndGroupNeeded = True
'        '    isPileNeeded = True
'        '    End If

'        'End If

'    End Sub 'Create Pile objects based on what is coming from the excel file


'#End Region

'#Region "Save Data"

'    Sub Save1Pile(ByVal pf As Pile)

'        Dim firstOne As Boolean = True
'        Dim mySoils As String = ""
'        Dim myLocations As String = ""

'        Dim PileSaver As String = QueryBuilderFromFile(queryPath & "Pile\Pile (IN_UP).sql")
'        PileSaver = PileSaver.Replace("[BU NUMBER]", BUNumber)
'        PileSaver = PileSaver.Replace("[STRUCTURE ID]", STR_ID)
'        PileSaver = PileSaver.Replace("[FOUNDATION TYPE]", "Pile")
'        If pf.pile_id = 0 Or IsDBNull(pf.pile_id) Then
'            PileSaver = PileSaver.Replace("'[Pile ID]'", "NULL")
'        Else
'            PileSaver = PileSaver.Replace("[Pile ID]", pf.pile_id.ToString)
'        End If

'        'Determine if new model ID needs created. Shouldn't be added to all individual tools (only needs to be referenced once)
'        If isModelNeeded Then
'            PileSaver = PileSaver.Replace("'[Model ID Needed]'", 1)
'        Else
'            PileSaver = PileSaver.Replace("'[Model ID Needed]'", 0)
'        End If

'        'Determine if new foundation group ID needs created. 
'        If isfndGroupNeeded Then
'            PileSaver = PileSaver.Replace("'[Fnd GRP ID Needed]'", 1)
'        Else
'            PileSaver = PileSaver.Replace("'[Fnd GRP ID Needed]'", 0)
'        End If

'        'Determine if new Pile ID needs created
'        If isPileNeeded Then
'            PileSaver = PileSaver.Replace("'[Pile ID Needed]'", 1)
'        Else
'            PileSaver = PileSaver.Replace("'[Pile ID Needed]'", 0)
'        End If

'        PileSaver = PileSaver.Replace("[INSERT ALL PILE DETAILS]", InsertPileDetail(pf))
'        PileSaver = PileSaver.Replace("[CONFIGURATION]", pf.pile_group_config.ToString)

'        'If pf.pile_id = 0 Or IsDBNull(pf.pile_id) Then 'Don't need since only performing insert commands. 
'        If pf.pile_soil_capacity_given = False And pf.pile_shape <> "H-Pile" Then
'            For Each pfsl As PileSoilLayer In pf.soil_layers
'                'line added below to avoid adding blank rows to tables when rows are removed. 
'                If Not IsNothing(pfsl.bottom_depth) Or Not IsNothing(pfsl.effective_soil_density) Or Not IsNothing(pfsl.cohesion) Or Not IsNothing(pfsl.friction_angle) Or Not IsNothing(pfsl.spt_blow_count) Or Not IsNothing(pfsl.ultimate_skin_friction_comp) Or Not IsNothing(pfsl.ultimate_skin_friction_uplift) Then
'                    Dim tempSoilLayer As String = InsertPileSoilLayer(pfsl)

'                    If Not firstOne Then
'                        mySoils += ",(" & tempSoilLayer & ")"
'                    Else
'                        mySoils += "(" & tempSoilLayer & ")"
'                    End If

'                    firstOne = False
'                End If
'            Next 'Add Soil Layer INSERT statments
'            PileSaver = PileSaver.Replace("([INSERT ALL SOIL LAYERS])", mySoils)
'            firstOne = True
'        Else
'            PileSaver = PileSaver.Replace("INSERT INTO fnd.pile_soil_layer VALUES ([INSERT ALL SOIL LAYERS])", "--INSERT INTO fnd.pile_soil_layer VALUES ([INSERT ALL SOIL LAYERS])")
'        End If

'        If pf.pile_group_config = "Asymmetric" Then
'            'PileSaver = PileSaver.Replace("[INSERT ALL PILE LOCATIONS]", InsertPileLocation(dp.embed_details))

'            For Each pfpl As PileLocation In pf.pile_locations
'                'line added below to avoid adding blank rows to tables when rows are removed. 
'                If Not IsNothing(pfpl.pile_x_coordinate) Or Not IsNothing(pfpl.pile_y_coordinate) Then
'                    Dim tempLocation As String = InsertPileLocation(pfpl)

'                    If Not firstOne Then
'                        myLocations += ",(" & tempLocation & ")"
'                    Else
'                        myLocations += "(" & tempLocation & ")"
'                    End If
'                End If
'                firstOne = False
'            Next
'            PileSaver = PileSaver.Replace("([INSERT ALL PILE LOCATIONS])", myLocations)
'        Else
'            PileSaver = PileSaver.Replace("BEGIN IF @IsCONFIG = 'Asymmetric'", "--BEGIN IF @IsCONFIG = 'Asymmetric'")
'            PileSaver = PileSaver.Replace("INSERT INTO fnd.pile_location VALUES ([INSERT ALL PILE LOCATIONS]) End", "--INSERT INTO fnd.pile_location VALUES ([INSERT ALL PILE LOCATIONS]) End")
'        End If 'Add Embedded Pole INSERT Statment

'        mySoils = ""
'        myLocations = ""

'        'Else 'No longer need to perform update commands

'        '    PileSaver = PileSaver.Replace("BEGIN IF @IsCONFIG = 'Asymmetric'", "--BEGIN IF @IsCONFIG = 'Asymmetric'")
'        '    PileSaver = PileSaver.Replace("INSERT INTO pile_soil_layer VALUES ([INSERT ALL SOIL LAYERS])", "--INSERT INTO pile_soil_layer VALUES ([INSERT ALL SOIL LAYERS])")
'        '    PileSaver = PileSaver.Replace("INSERT INTO pile_location VALUES ([INSERT ALL PILE LOCATIONS]) End", "--INSERT INTO pile_location VALUES ([INSERT ALL PILE LOCATIONS]) End")

'        '    Dim tempUpdater As String = ""
'        '    tempUpdater += UpdatePileDetail(pf)

'        '    If pf.pile_soil_capacity_given = False And pf.pile_shape <> "H-Pile" Then
'        '        For Each pfsl As PileSoilLayer In pf.soil_layers
'        '            If pfsl.soil_layer_id = 0 Or IsDBNull(pfsl.soil_layer_id) Then
'        '                tempUpdater += "INSERT INTO pile_soil_layer VALUES (" & InsertPileSoilLayer(pfsl) & ") " & vbNewLine
'        '            Else
'        '                tempUpdater += UpdatePileSoilLayer(pfsl)
'        '            End If
'        '        Next
'        '    End If

'        '    'PileSaver = PileSaver.Replace("(SELECT * FROM TEMPORARY)", tempUpdater)

'        '    'End If

'        '    If pf.pile_group_config = "Asymmetric" Then

'        '        For Each pfpl As PileLocation In pf.pile_locations
'        '            If pfpl.location_id = 0 Or IsDBNull(pfpl.location_id) Then
'        '                tempUpdater += "INSERT INTO pile_location VALUES (" & InsertPileLocation(pfpl) & ") " & vbNewLine
'        '            Else
'        '                tempUpdater += UpdatePileLocation(pfpl)
'        '            End If
'        '        Next




'        '        '    'If pfpl.location_id = 0 Or IsDBNull(dp.embed_details.embedded_id) Then
'        '        '    'tempUpdater += "BEGIN INSERT INTO embedded_pole_details OUTPUT INSERTED.ID INTO @EmbeddedPole VALUES (" & InsertDrilledPierEmbed(dp.embed_details) & ") " & vbNewLine & " SELECT @EmbedID=EmbedID FROM @EmbeddedPole"
'        '        '    For Each pfpl As PileLocation In pf.pile_locations
'        '        '        tempUpdater += "INSERT INTO pile_location VALUES (" & InsertPileLocation(pfpl) & ") " & vbNewLine
'        '        '    Next
'        '        '    tempUpdater += " END " & vbNewLine
'        '        'Else
'        '        '    tempUpdater += UpdateDrilledPierEmbed(dp.embed_details)
'        '        '    For Each esec As DrilledPierEmbedSection In dp.embed_details.sections
'        '        '        If esec.section_id = 0 Or IsDBNull(esec.section_id) Then
'        '        '            tempUpdater += "INSERT INTO embedded_pole_section VALUES (" & InsertDrilledPierEmbedSection(esec).Replace("@EmbedID", dp.embed_details.embedded_id.ToString) & ") " & vbNewLine
'        '        '        Else
'        '        '            tempUpdater += UpdateDrilledPierEmbedSection(esec)
'        '        '        End If
'        '        '    Next
'        '        '    'End If
'        '    End If

'        '    PileSaver = PileSaver.Replace("(SELECT * FROM TEMPORARY)", tempUpdater)

'        'End If

'        sqlSender(PileSaver, pileDB, pileID, "0")
'    End Sub

'    Public Sub SaveToEDS()
'        For Each pf As Pile In Piles
'            Save1Pile(pf)
'            'Moved code to separate method, above (Save1Pile) No changes were made to the code copied over
'        Next
'    End Sub

'    Public Sub SaveToExcel()
'        'Dim pfRow As Integer = 3
'        'Dim soilRow As Integer = 4 'identify first row to copy data into Excel Sheet
'        'Dim soilRow As Integer = 57 'identify first row to copy data into Excel Sheet
'        'Dim locRow As Integer = 4
'        'Dim locRow As Integer = 5
'        'LoadNewPile() 'follows drilled pier format

'        'With NewPileWb 'follows drilled pier format
'        For Each pf As Pile In Piles
'            Dim soilRow As Integer = 57 'identify first row to copy data into Excel Sheet
'            Dim locRow As Integer = 5
'            LoadNewPile() 'follows p&p format
'            With NewPileWb 'follows p&p format

'                If Not IsNothing(pf.pile_id) Then
'                    .Worksheets("Input").Range("ID").Value = CType(pf.pile_id, Integer)
'                Else .Worksheets("Input").Range("ID").ClearContents
'                End If
'                If Not IsNothing(pf.load_eccentricity) Then
'                    .Worksheets("Input").Range("Ecc").Value = CType(pf.load_eccentricity, Double)
'                Else .Worksheets("Input").Range("Ecc").ClearContents
'                End If
'                If Not IsNothing(pf.bolt_circle_bearing_plate_width) Then
'                    .Worksheets("Input").Range("BC").Value = CType(pf.bolt_circle_bearing_plate_width, Double)
'                Else .Worksheets("Input").Range("BC").ClearContents
'                End If
'                If Not IsNothing(pf.pile_shape) Then .Worksheets("Input").Range("D23").Value = pf.pile_shape
'                If Not IsNothing(pf.pile_material) Then .Worksheets("Input").Range("D24").Value = pf.pile_material
'                If Not IsNothing(pf.pile_length) Then
'                    .Worksheets("Input").Range("Lpile").Value = CType(pf.pile_length, Double)
'                Else .Worksheets("Input").Range("Lpile").ClearContents
'                End If
'                If Not IsNothing(pf.pile_diameter_width) Then
'                    .Worksheets("Input").Range("D26").Value = CType(pf.pile_diameter_width, Double)
'                Else .Worksheets("Input").Range("D26").ClearContents
'                End If
'                If Not IsNothing(pf.pile_pipe_thickness) Then
'                    .Worksheets("Input").Range("D27").Value = CType(pf.pile_pipe_thickness, Double)
'                Else .Worksheets("Input").Range("D27").ClearContents
'                End If

'                If pf.pile_soil_capacity_given = True Then
'                    .Worksheets("Input").Range("D29").Value = "Yes"
'                Else
'                    .Worksheets("Input").Range("D29").Value = "No"
'                End If

'                If Not IsNothing(pf.steel_yield_strength) Then
'                    .Worksheets("Input").Range("D30").Value = CType(pf.steel_yield_strength, Double)
'                Else .Worksheets("Input").Range("D30").ClearContents
'                End If
'                'If Not IsNothing(pf.pile_type_option) Then .Worksheets("Input").Range("Psize").Value = pf.pile_type_option
'                If Not IsNothing(pf.pile_type_option) Then
'                    If pf.pile_material = "Concrete" Then
'                        .Worksheets("Input").Range("Psize").Value = CType(pf.pile_type_option, Integer)
'                    Else
'                        .Worksheets("Input").Range("Psize").Value = CType(pf.pile_type_option, String)
'                    End If
'                End If
'                If Not IsNothing(pf.rebar_quantity) Then
'                    .Worksheets("Input").Range("Pquan").Value = CType(pf.rebar_quantity, Double)
'                Else .Worksheets("Input").Range("Pquan").ClearContents
'                End If
'                If Not IsNothing(pf.pile_group_config) Then .Worksheets("Input").Range("Config").Value = pf.pile_group_config
'                If Not IsNothing(pf.foundation_depth) Then
'                    .Worksheets("Input").Range("D").Value = CType(pf.foundation_depth, Double)
'                Else .Worksheets("Input").Range("D").ClearContents
'                End If
'                If Not IsNothing(pf.pad_thickness) Then
'                    .Worksheets("Input").Range("T").Value = CType(pf.pad_thickness, Double)
'                Else .Worksheets("Input").Range("T").ClearContents
'                End If
'                If Not IsNothing(pf.pad_width_dir1) Then
'                    .Worksheets("Input").Range("Wx").Value = CType(pf.pad_width_dir1, Double)
'                Else .Worksheets("Input").Range("Wx").ClearContents
'                End If
'                If Not IsNothing(pf.pad_width_dir2) Then
'                    .Worksheets("Input").Range("Wy").Value = CType(pf.pad_width_dir2, Double)
'                Else .Worksheets("Input").Range("Wy").ClearContents
'                End If
'                If Not IsNothing(pf.pad_rebar_size_bottom) Then
'                    .Worksheets("Input").Range("Spad").Value = CType(pf.pad_rebar_size_bottom, Integer)
'                Else .Worksheets("Input").Range("Spad").ClearContents
'                End If
'                If Not IsNothing(pf.pad_rebar_size_top) Then
'                    .Worksheets("Input").Range("Spad_top").Value = CType(pf.pad_rebar_size_top, Integer)
'                Else .Worksheets("Input").Range("Spad_top").ClearContents
'                End If
'                If Not IsNothing(pf.pad_rebar_quantity_bottom_dir1) Then
'                    .Worksheets("Input").Range("Mpad").Value = CType(pf.pad_rebar_quantity_bottom_dir1, Double)
'                Else .Worksheets("Input").Range("Mpad").ClearContents
'                End If
'                If Not IsNothing(pf.pad_rebar_quantity_top_dir1) Then
'                    .Worksheets("Input").Range("Mpad_top").Value = CType(pf.pad_rebar_quantity_top_dir1, Double)
'                Else .Worksheets("Input").Range("Mpad_top").ClearContents
'                End If
'                If Not IsNothing(pf.pad_rebar_quantity_bottom_dir2) Then
'                    .Worksheets("Input").Range("Mpad_y").Value = CType(pf.pad_rebar_quantity_bottom_dir2, Double)
'                Else .Worksheets("Input").Range("Mpad_y").ClearContents
'                End If
'                If Not IsNothing(pf.pad_rebar_quantity_top_dir2) Then
'                    .Worksheets("Input").Range("Mpad_y_top").Value = CType(pf.pad_rebar_quantity_top_dir2, Double)
'                Else .Worksheets("Input").Range("Mpad_y_top").ClearContents
'                End If
'                If Not IsNothing(pf.pier_shape) Then .Worksheets("Input").Range("D57").Value = pf.pier_shape
'                If Not IsNothing(pf.pier_diameter) Then
'                    .Worksheets("Input").Range("di").Value = CType(pf.pier_diameter, Double)
'                Else .Worksheets("Input").Range("di").ClearContents
'                End If
'                If Not IsNothing(pf.extension_above_grade) Then
'                    .Worksheets("Input").Range("E").Value = CType(pf.extension_above_grade, Double)
'                Else .Worksheets("Input").Range("E").ClearContents
'                End If
'                If Not IsNothing(pf.pier_rebar_size) Then
'                    .Worksheets("Input").Range("Rs").Value = CType(pf.pier_rebar_size, Integer)
'                Else .Worksheets("Input").Range("Rs").ClearContents
'                End If
'                If Not IsNothing(pf.pier_rebar_quantity) Then
'                    .Worksheets("Input").Range("mc").Value = CType(pf.pier_rebar_quantity, Double)
'                Else .Worksheets("Input").Range("mc").ClearContents
'                End If
'                If Not IsNothing(pf.pier_tie_size) Then
'                    .Worksheets("Input").Range("St").Value = CType(pf.pier_tie_size, Integer)
'                Else .Worksheets("Input").Range("St").ClearContents
'                End If
'                'If Not IsNothing(pf.pier_tie_quantity) Then
'                '    .Worksheets("").Range("").Value = CType(pf.pier_tie_quantity, Integer)
'                'Else .Worksheets("").Range("").ClearContents
'                'End If
'                If Not IsNothing(pf.rebar_grade) Then
'                    .Worksheets("Input").Range("Fy").Value = CType(pf.rebar_grade, Double)
'                Else .Worksheets("Input").Range("Fy").ClearContents
'                End If
'                If Not IsNothing(pf.concrete_compressive_strength) Then
'                    .Worksheets("Input").Range("Fc").Value = CType(pf.concrete_compressive_strength, Double)
'                Else .Worksheets("Input").Range("Fc").ClearContents
'                End If
'                If Not IsNothing(pf.groundwater_depth) Then
'                    .Worksheets("Input").Range("D69").Value = CType(pf.groundwater_depth, Double)
'                    'Else .Worksheets("Input").Range("D69").ClearContents 'adjusted so will always report N/A if null
'                Else .Worksheets("Input").Range("D69").Value = "N/A"
'                End If
'                If Not IsNothing(pf.total_soil_unit_weight) Then
'                    .Worksheets("Input").Range("γsoil_dry").Value = CType(pf.total_soil_unit_weight, Double)
'                Else .Worksheets("Input").Range("γsoil_dry").ClearContents
'                End If
'                If Not IsNothing(pf.cohesion) Then
'                    .Worksheets("Input").Range("Co").Value = CType(pf.cohesion, Double)
'                Else .Worksheets("Input").Range("Co").ClearContents
'                End If
'                If Not IsNothing(pf.friction_angle) Then
'                    .Worksheets("Input").Range("ɸ").Value = CType(pf.friction_angle, Double)
'                Else .Worksheets("Input").Range("ɸ").ClearContents
'                End If
'                If Not IsNothing(pf.neglect_depth) Then
'                    .Worksheets("Input").Range("ND").Value = CType(pf.neglect_depth, Double)
'                Else .Worksheets("Input").Range("ND").ClearContents
'                End If
'                If Not IsNothing(pf.spt_blow_count) Then
'                    .Worksheets("Input").Range("N_blows").Value = CType(pf.spt_blow_count, Double)
'                Else .Worksheets("Input").Range("N_blows").ClearContents
'                End If
'                If Not IsNothing(pf.pile_negative_friction_force) Then
'                    .Worksheets("Input").Range("Sw").Value = CType(pf.pile_negative_friction_force, Double)
'                Else .Worksheets("Input").Range("Sw").ClearContents
'                End If
'                If Not IsNothing(pf.pile_ultimate_compression) Then
'                    .Worksheets("Input").Range("K45").Value = CType(pf.pile_ultimate_compression, Double)
'                Else .Worksheets("Input").Range("K45").ClearContents
'                End If
'                If Not IsNothing(pf.pile_ultimate_tension) Then
'                    .Worksheets("Input").Range("K46").Value = CType(pf.pile_ultimate_tension, Double)
'                Else .Worksheets("Input").Range("K46").ClearContents
'                End If
'                If Not IsNothing(pf.top_and_bottom_rebar_different) Then .Worksheets("Input").Range("Z10").Value = pf.top_and_bottom_rebar_different
'                If Not IsNothing(pf.ultimate_gross_end_bearing) Then
'                    .Worksheets("Input").Range("M71").Value = CType(pf.ultimate_gross_end_bearing, Double)
'                Else .Worksheets("Input").Range("M71").ClearContents
'                End If

'                If pf.skin_friction_given = True Then
'                    .Worksheets("Input").Range("N54").Value = "Yes"
'                Else
'                    .Worksheets("Input").Range("N54").Value = "No"
'                End If

'                If pf.pile_group_config = "Circular" Then
'                    If Not IsNothing(pf.pile_quantity_circular) Then
'                        .Worksheets("Input").Range("D36").Value = CType(pf.pile_quantity_circular, Double)
'                    Else .Worksheets("Input").Range("D36").ClearContents
'                    End If
'                    If Not IsNothing(pf.group_diameter_circular) Then
'                        .Worksheets("Input").Range("D37").Value = CType(pf.group_diameter_circular, Double)
'                    Else .Worksheets("Input").Range("D37").ClearContents
'                    End If
'                End If

'                If pf.pile_group_config = "Rectangular" Then
'                    If Not IsNothing(pf.pile_column_quantity) Then
'                        .Worksheets("Input").Range("D36").Value = CType(pf.pile_column_quantity, Double)
'                    Else .Worksheets("Input").Range("D36").ClearContents
'                    End If
'                    If Not IsNothing(pf.pile_row_quantity) Then
'                        .Worksheets("Input").Range("D37").Value = CType(pf.pile_row_quantity, Double)
'                    Else .Worksheets("Input").Range("D37").ClearContents
'                    End If
'                End If

'                If Not IsNothing(pf.pile_columns_spacing) Then
'                    .Worksheets("Input").Range("D38").Value = CType(pf.pile_columns_spacing, Double)
'                Else .Worksheets("Input").Range("D38").ClearContents
'                End If
'                If Not IsNothing(pf.pile_row_spacing) Then
'                    .Worksheets("Input").Range("D39").Value = CType(pf.pile_row_spacing, Double)
'                Else .Worksheets("Input").Range("D39").ClearContents
'                End If

'                If pf.group_efficiency_factor_given = True Then
'                    .Worksheets("Input").Range("D41").Value = "Yes"
'                Else
'                    .Worksheets("Input").Range("D41").Value = "No"
'                End If

'                If Not IsNothing(pf.group_efficiency_factor) Then
'                    .Worksheets("Input").Range("D42").Value = CType(pf.group_efficiency_factor, Double)
'                Else .Worksheets("Input").Range("D42").ClearContents
'                End If
'                If Not IsNothing(pf.cap_type) Then .Worksheets("Input").Range("D45").Value = pf.cap_type
'                If Not IsNothing(pf.pile_quantity_asymmetric) Then
'                    .Worksheets("Moment of Inertia").Range("D10").Value = CType(pf.pile_quantity_asymmetric, Double)
'                Else .Worksheets("Moment of Inertia").Range("D10").ClearContents
'                End If
'                If Not IsNothing(pf.pile_spacing_min_asymmetric) Then
'                    .Worksheets("Moment of Inertia").Range("D11").Value = CType(pf.pile_spacing_min_asymmetric, Double)
'                Else .Worksheets("Moment of Inertia").Range("D11").ClearContents
'                End If
'                If Not IsNothing(pf.quantity_piles_surrounding) Then
'                    .Worksheets("Moment of Inertia").Range("D12").Value = CType(pf.quantity_piles_surrounding, Double)
'                Else .Worksheets("Moment of Inertia").Range("D12").ClearContents
'                End If
'                If Not IsNothing(pf.pile_cap_reference) Then .Worksheets("Input").Range("G47").Value = pf.pile_cap_reference
'                'If Not IsNothing(pf.tool_version) Then .Worksheets("Revision History").Range("Revision").Value = pf.tool_version
'                If Not IsNothing(pf.Soil_110) Then .Worksheets("Input").Range("Z13").Value = pf.Soil_110
'                If Not IsNothing(pf.Structural_105) Then .Worksheets("Input").Range("Z14").Value = pf.Structural_105

'                If pf.pile_soil_capacity_given = False And pf.pile_shape <> "H-Pile" Then
'                    For Each pfSL As PileSoilLayer In pf.soil_layers

'                        'If Not IsNothing(pfSL.soil_layer_id) Then
'                        '    .Worksheets("SAPI").Range("J" & soilRow).Value = CType(pfSL.soil_layer_id, Integer)
'                        'Else .Worksheets("SAPI").Range("J" & soilRow).ClearContents
'                        'End If
'                        'If Not IsNothing(pfSL.bottom_depth) Then
'                        '    .Worksheets("SAPI").Range("K" & soilRow).Value = CType(pfSL.bottom_depth, Double)
'                        'Else .Worksheets("SAPI").Range("K" & soilRow).ClearContents
'                        'End If
'                        'If Not IsNothing(pfSL.effective_soil_density) Then
'                        '    .Worksheets("SAPI").Range("L" & soilRow).Value = CType(pfSL.effective_soil_density, Double)
'                        'Else .Worksheets("SAPI").Range("L" & soilRow).ClearContents
'                        'End If
'                        'If Not IsNothing(pfSL.cohesion) Then
'                        '    .Worksheets("SAPI").Range("M" & soilRow).Value = CType(pfSL.cohesion, Double)
'                        'Else .Worksheets("SAPI").Range("M" & soilRow).ClearContents
'                        'End If
'                        'If Not IsNothing(pfSL.friction_angle) Then
'                        '    .Worksheets("SAPI").Range("N" & soilRow).Value = CType(pfSL.friction_angle, Double)
'                        'Else .Worksheets("SAPI").Range("N" & soilRow).ClearContents
'                        'End If
'                        ''If Not IsNothing(pfSL.skin_friction_override_uplift) Then
'                        ''    .Worksheets("SAPI").Range("N54").Value = CType(pfSL.skin_friction_override_uplift, Double)
'                        ''Else .Worksheets("SAPI").Range("N54").ClearContents
'                        ''End If
'                        'If Not IsNothing(pfSL.spt_blow_count) Then
'                        '    .Worksheets("SAPI").Range("O" & soilRow).Value = CType(pfSL.spt_blow_count, Integer)
'                        'Else .Worksheets("SAPI").Range("O" & soilRow).ClearContents
'                        'End If
'                        'If Not IsNothing(pfSL.ultimate_skin_friction_comp) Then
'                        '    .Worksheets("SAPI").Range("P" & soilRow).Value = CType(pfSL.ultimate_skin_friction_comp, Double)
'                        'Else .Worksheets("SAPI").Range("P" & soilRow).ClearContents
'                        'End If
'                        'If Not IsNothing(pfSL.ultimate_skin_friction_uplift) Then
'                        '    .Worksheets("SAPI").Range("Q" & soilRow).Value = CType(pfSL.ultimate_skin_friction_uplift, Double)
'                        'Else .Worksheets("SAPI").Range("Q" & soilRow).ClearContents
'                        'End If

'                        '**** Section Below eliminates workbook open events ****

'                        If Not IsNothing(pfSL.soil_layer_id) Then
'                            .Worksheets("SAPI").Range("J" & soilRow - 53).Value = CType(pfSL.soil_layer_id, Integer)
'                            'Else .Worksheets("SAPI").Range("J" & soilRow - 53).ClearContents
'                        End If
'                        If Not IsNothing(pfSL.bottom_depth) Then
'                            .Worksheets("Input").Range("H" & soilRow).Value = CType(pfSL.bottom_depth, Double)
'                            'Else .Worksheets("Input").Range("H" & soilRow).ClearContents
'                        End If
'                        If Not IsNothing(pfSL.effective_soil_density) Then
'                            .Worksheets("Input").Range("K" & soilRow).Value = CType(pfSL.effective_soil_density, Double)
'                            'Else .Worksheets("Input").Range("K" & soilRow).ClearContents
'                        End If
'                        If Not IsNothing(pfSL.cohesion) Then
'                            .Worksheets("Input").Range("I" & soilRow).Value = CType(pfSL.cohesion, Double)
'                            'Else .Worksheets("Input").Range("I" & soilRow).ClearContents
'                        End If
'                        If Not IsNothing(pfSL.friction_angle) Then
'                            .Worksheets("Input").Range("J" & soilRow).Value = CType(pfSL.friction_angle, Double)
'                            'Else .Worksheets("Input").Range("J" & soilRow).ClearContents
'                        End If
'                        'If Not IsNothing(pfSL.skin_friction_override_uplift) Then
'                        '    .Worksheets("Input").Range("N54").Value = CType(pfSL.skin_friction_override_uplift, Double)
'                        'Else .Worksheets("Input").Range("N54").ClearContents
'                        'End If
'                        If Not IsNothing(pfSL.spt_blow_count) Then
'                            .Worksheets("Input").Range("L" & soilRow).Value = CType(pfSL.spt_blow_count, Double)
'                            'Else .Worksheets("Input").Range("L" & soilRow).ClearContents
'                        End If
'                        If Not IsNothing(pfSL.ultimate_skin_friction_comp) Then
'                            .Worksheets("Input").Range("M" & soilRow).Value = CType(pfSL.ultimate_skin_friction_comp, Double)
'                            'Else .Worksheets("Input").Range("M" & soilRow).ClearContents
'                        End If
'                        If Not IsNothing(pfSL.ultimate_skin_friction_uplift) Then
'                            .Worksheets("Input").Range("N" & soilRow).Value = CType(pfSL.ultimate_skin_friction_uplift, Double)
'                            'Else .Worksheets("Input").Range("N" & soilRow).ClearContents
'                        End If
'                        '******

'                        soilRow += 1
'                    Next
'                End If

'                If pf.pile_group_config = "Asymmetric" Then

'                    For Each pfPL As PileLocation In pf.pile_locations
'                        'If Not IsNothing(pfPL.location_id) Then
'                        '    .Worksheets("SAPI").Range("W" & locRow).Value = CType(pfPL.location_id, Integer)
'                        'Else .Worksheets("SAPI").Range("W" & locRow).ClearContents
'                        'End If
'                        'If Not IsNothing(pfPL.pile_x_coordinate) Then
'                        '    .Worksheets("SAPI").Range("X" & locRow).Value = CType(pfPL.pile_x_coordinate, Double)
'                        'Else .Worksheets("SAPI").Range("X" & locRow).ClearContents
'                        'End If
'                        'If Not IsNothing(pfPL.pile_y_coordinate) Then
'                        '    .Worksheets("SAPI").Range("Y" & locRow).Value = CType(pfPL.pile_y_coordinate, Double)
'                        'Else .Worksheets("SAPI").Range("Y" & locRow).ClearContents
'                        'End If

'                        '**** Section Below eliminates workbook open events ****
'                        If Not IsNothing(pfPL.location_id) Then
'                            .Worksheets("SAPI").Range("W" & locRow - 1).Value = CType(pfPL.location_id, Integer)
'                        Else .Worksheets("SAPI").Range("W" & locRow - 1).ClearContents
'                        End If
'                        If Not IsNothing(pfPL.pile_x_coordinate) Then
'                            .Worksheets("Moment of Inertia").Range("K" & locRow).Value = CType(pfPL.pile_x_coordinate, Double)
'                            'Else .Worksheets("Moment of Inertia").Range("K" & locRow).ClearContents
'                        End If
'                        If Not IsNothing(pfPL.pile_y_coordinate) Then
'                            .Worksheets("Moment of Inertia").Range("L" & locRow).Value = CType(pfPL.pile_y_coordinate, Double)
'                            'Else .Worksheets("Moment of Inertia").Range("L" & locRow).ClearContents
'                        End If
'                        '*****

'                        locRow += 1
'                    Next
'                End If

'                'Worksheet Change Events
'                'Hiding/unhiding specific tabs
'                If pf.pile_group_config = "Circular" Then
'                    .Worksheets("Moment of Inertia").Visible = False
'                    .Worksheets("Moment of Inertia (Circle)").Visible = True
'                Else
'                    .Worksheets("Moment of Inertia").Visible = True
'                    .Worksheets("Moment of Inertia (Circle)").Visible = False
'                End If

'                'Resizing Image
'                'Try
'                '    With .Worksheets("Input").Charts(0)
'                '        .Width = (300 / Math.Max(CType(pf.pad_width_dir1, Double), CType(pf.pad_width_dir2, Double))) * CType(pf.pad_width_dir1, Double) * 4.19 '4.19 multiplier determined through trial and error. 
'                '        .Height = (300 / Math.Max(CType(pf.pad_width_dir1, Double), CType(pf.pad_width_dir2, Double))) * CType(pf.pad_width_dir2, Double) * 4.19
'                '    End With
'                'Catch
'                '    'error handling to avoid dividing by zero
'                'End Try
'            End With 'follows P&P format

'            SaveAndClosePile() 'follows P&P format

'        Next 'follows drilled pier format

'        'End With 'follows drilled pier format

'        'SaveAndClosePile() 'follows drilled pier format
'    End Sub

'    Private Sub LoadNewPile()
'        NewPileWb.LoadDocument(PileTemplatePath, PileFileType)
'        NewPileWb.BeginUpdate()
'    End Sub

'    Private Sub SaveAndClosePile()
'        NewPileWb.Calculate()
'        NewPileWb.EndUpdate()
'        NewPileWb.SaveDocument(ExcelFilePath, PileFileType)
'    End Sub
'#End Region

'#Region "SQL Insert Statements"
'    Private Function InsertPileDetail(ByVal pf As Pile) As String
'        Dim insertString As String = ""

'        'insertString += "@FndID"
'        'insertString += "@FndgrpID"
'        'insertString += "," & IIf(IsNothing(pf.pile_id), "Null", pf.pile_id.ToString)
'        insertString += "" & IIf(IsNothing(pf.load_eccentricity), "Null", pf.load_eccentricity.ToString)
'        insertString += "," & IIf(IsNothing(pf.bolt_circle_bearing_plate_width), "Null", pf.bolt_circle_bearing_plate_width.ToString)
'        insertString += "," & IIf(IsNothing(pf.pile_shape), "Null", "'" & pf.pile_shape.ToString & "'")
'        insertString += "," & IIf(IsNothing(pf.pile_material), "Null", "'" & pf.pile_material.ToString & "'")
'        insertString += "," & IIf(IsNothing(pf.pile_length), "Null", pf.pile_length.ToString)
'        insertString += "," & IIf(IsNothing(pf.pile_diameter_width), "Null", pf.pile_diameter_width.ToString)
'        insertString += "," & IIf(IsNothing(pf.pile_pipe_thickness), "Null", pf.pile_pipe_thickness.ToString)
'        insertString += "," & IIf(IsNothing(pf.pile_soil_capacity_given), "Null", "'" & pf.pile_soil_capacity_given.ToString & "'")
'        insertString += "," & IIf(IsNothing(pf.steel_yield_strength), "Null", pf.steel_yield_strength.ToString)
'        insertString += "," & IIf(IsNothing(pf.pile_type_option), "Null", "'" & pf.pile_type_option.ToString & "'")
'        insertString += "," & IIf(IsNothing(pf.rebar_quantity), "Null", pf.rebar_quantity.ToString)
'        insertString += "," & IIf(IsNothing(pf.pile_group_config), "Null", "'" & pf.pile_group_config.ToString & "'")
'        insertString += "," & IIf(IsNothing(pf.foundation_depth), "Null", pf.foundation_depth.ToString)
'        insertString += "," & IIf(IsNothing(pf.pad_thickness), "Null", pf.pad_thickness.ToString)
'        insertString += "," & IIf(IsNothing(pf.pad_width_dir1), "Null", pf.pad_width_dir1.ToString)
'        insertString += "," & IIf(IsNothing(pf.pad_width_dir2), "Null", pf.pad_width_dir2.ToString)
'        insertString += "," & IIf(IsNothing(pf.pad_rebar_size_bottom), "Null", pf.pad_rebar_size_bottom.ToString)
'        insertString += "," & IIf(IsNothing(pf.pad_rebar_size_top), "Null", pf.pad_rebar_size_top.ToString)
'        insertString += "," & IIf(IsNothing(pf.pad_rebar_quantity_bottom_dir1), "Null", pf.pad_rebar_quantity_bottom_dir1.ToString)
'        insertString += "," & IIf(IsNothing(pf.pad_rebar_quantity_top_dir1), "Null", pf.pad_rebar_quantity_top_dir1.ToString)
'        insertString += "," & IIf(IsNothing(pf.pad_rebar_quantity_bottom_dir2), "Null", pf.pad_rebar_quantity_bottom_dir2.ToString)
'        insertString += "," & IIf(IsNothing(pf.pad_rebar_quantity_top_dir2), "Null", pf.pad_rebar_quantity_top_dir2.ToString)
'        insertString += "," & IIf(IsNothing(pf.pier_shape), "Null", "'" & pf.pier_shape.ToString & "'")
'        insertString += "," & IIf(IsNothing(pf.pier_diameter), "Null", pf.pier_diameter.ToString)
'        insertString += "," & IIf(IsNothing(pf.extension_above_grade), "Null", pf.extension_above_grade.ToString)
'        insertString += "," & IIf(IsNothing(pf.pier_rebar_size), "Null", pf.pier_rebar_size.ToString)
'        insertString += "," & IIf(IsNothing(pf.pier_rebar_quantity), "Null", pf.pier_rebar_quantity.ToString)
'        insertString += "," & IIf(IsNothing(pf.pier_tie_size), "Null", pf.pier_tie_size.ToString)
'        insertString += "," & IIf(IsNothing(pf.rebar_grade), "Null", pf.rebar_grade.ToString)
'        insertString += "," & IIf(IsNothing(pf.concrete_compressive_strength), "Null", pf.concrete_compressive_strength.ToString)
'        insertString += "," & IIf(IsNothing(pf.groundwater_depth), "Null", pf.groundwater_depth.ToString)
'        insertString += "," & IIf(IsNothing(pf.total_soil_unit_weight), "Null", pf.total_soil_unit_weight.ToString)
'        insertString += "," & IIf(IsNothing(pf.cohesion), "Null", pf.cohesion.ToString)
'        insertString += "," & IIf(IsNothing(pf.friction_angle), "Null", pf.friction_angle.ToString)
'        insertString += "," & IIf(IsNothing(pf.neglect_depth), "Null", pf.neglect_depth.ToString)
'        insertString += "," & IIf(IsNothing(pf.spt_blow_count), "Null", pf.spt_blow_count.ToString)
'        insertString += "," & IIf(IsNothing(pf.pile_negative_friction_force), "Null", pf.pile_negative_friction_force.ToString)
'        insertString += "," & IIf(IsNothing(pf.pile_ultimate_compression), "Null", pf.pile_ultimate_compression.ToString)
'        insertString += "," & IIf(IsNothing(pf.pile_ultimate_tension), "Null", pf.pile_ultimate_tension.ToString)
'        insertString += "," & IIf(IsNothing(pf.top_and_bottom_rebar_different), "Null", "'" & pf.top_and_bottom_rebar_different.ToString & "'")
'        insertString += "," & IIf(IsNothing(pf.ultimate_gross_end_bearing), "Null", pf.ultimate_gross_end_bearing.ToString)
'        insertString += "," & IIf(IsNothing(pf.skin_friction_given), "Null", "'" & pf.skin_friction_given.ToString & "'")
'        insertString += "," & IIf(IsNothing(pf.pile_quantity_circular), "Null", pf.pile_quantity_circular.ToString)
'        insertString += "," & IIf(IsNothing(pf.group_diameter_circular), "Null", pf.group_diameter_circular.ToString)
'        insertString += "," & IIf(IsNothing(pf.pile_column_quantity), "Null", pf.pile_column_quantity.ToString)
'        insertString += "," & IIf(IsNothing(pf.pile_row_quantity), "Null", pf.pile_row_quantity.ToString)
'        insertString += "," & IIf(IsNothing(pf.pile_columns_spacing), "Null", pf.pile_columns_spacing.ToString)
'        insertString += "," & IIf(IsNothing(pf.pile_row_spacing), "Null", pf.pile_row_spacing.ToString)
'        insertString += "," & IIf(IsNothing(pf.group_efficiency_factor_given), "Null", "'" & pf.group_efficiency_factor_given.ToString & "'")
'        insertString += "," & IIf(IsNothing(pf.group_efficiency_factor), "Null", pf.group_efficiency_factor.ToString)
'        insertString += "," & IIf(IsNothing(pf.cap_type), "Null", "'" & pf.cap_type.ToString & "'")
'        insertString += "," & IIf(IsNothing(pf.pile_quantity_asymmetric), "Null", pf.pile_quantity_asymmetric.ToString)
'        insertString += "," & IIf(IsNothing(pf.pile_spacing_min_asymmetric), "Null", pf.pile_spacing_min_asymmetric.ToString)
'        insertString += "," & IIf(IsNothing(pf.quantity_piles_surrounding), "Null", pf.quantity_piles_surrounding.ToString)
'        insertString += "," & IIf(IsNothing(pf.pile_cap_reference), "Null", "'" & pf.pile_cap_reference.ToString & "'")
'        insertString += "," & IIf(IsNothing(pf.tool_version), "Null", "'" & pf.tool_version.ToString & "'")
'        insertString += "," & IIf(IsNothing(pf.Soil_110), "Null", "'" & pf.Soil_110.ToString & "'")
'        insertString += "," & IIf(IsNothing(pf.Structural_105), "Null", "'" & pf.Structural_105.ToString & "'")

'        Return insertString
'    End Function

'    Private Function InsertPileSoilLayer(ByVal pfsl As PileSoilLayer) As String
'        Dim insertString As String = ""

'        insertString += "@PID"
'        'insertString += "," & IIf(IsNothing(pfsl.soil_layer_id), "Null", pfsl.soil_layer_id.ToString)
'        insertString += "," & IIf(IsNothing(pfsl.bottom_depth), "Null", pfsl.bottom_depth.ToString)
'        insertString += "," & IIf(IsNothing(pfsl.effective_soil_density), "Null", pfsl.effective_soil_density.ToString)
'        insertString += "," & IIf(IsNothing(pfsl.cohesion), "Null", pfsl.cohesion.ToString)
'        insertString += "," & IIf(IsNothing(pfsl.friction_angle), "Null", pfsl.friction_angle.ToString)
'        'insertString += "," & IIf(IsNothing(pfsl.skin_friction_override_uplift), "Null", pfsl.skin_friction_override_uplift.ToString)
'        insertString += "," & IIf(IsNothing(pfsl.spt_blow_count), "Null", pfsl.spt_blow_count.ToString)
'        insertString += "," & IIf(IsNothing(pfsl.ultimate_skin_friction_comp), "Null", pfsl.ultimate_skin_friction_comp.ToString)
'        insertString += "," & IIf(IsNothing(pfsl.ultimate_skin_friction_uplift), "Null", pfsl.ultimate_skin_friction_uplift.ToString)

'        Return insertString
'    End Function

'    Private Function InsertPileLocation(ByVal pfpl As PileLocation) As String
'        Dim insertString As String = ""

'        insertString += "@PID"
'        'insertString += "," & IIf(IsNothing(pfpl.location_id), "Null", pfpl.location_id.ToString)
'        insertString += "," & IIf(IsNothing(pfpl.pile_x_coordinate), "Null", pfpl.pile_x_coordinate.ToString)
'        insertString += "," & IIf(IsNothing(pfpl.pile_y_coordinate), "Null", pfpl.pile_y_coordinate.ToString)

'        Return insertString
'    End Function


'#End Region

'#Region "SQL Update Statements"
'    Private Function UpdatePileDetail(ByVal pf As Pile) As String
'        Dim updateString As String = ""

'        updateString += "UPDATE pile_details SET "
'        'updateString += ", pile_id=" & IIf(IsNothing(pf.pile_id), "Null", pf.pile_id.ToString)
'        updateString += " load_eccentricity=" & IIf(IsNothing(pf.load_eccentricity), "Null", pf.load_eccentricity.ToString)
'        updateString += ", bolt_circle_bearing_plate_width=" & IIf(IsNothing(pf.bolt_circle_bearing_plate_width), "Null", pf.bolt_circle_bearing_plate_width.ToString)
'        updateString += ", pile_shape=" & IIf(IsNothing(pf.pile_shape), "Null", "'" & pf.pile_shape.ToString & "'")
'        updateString += ", pile_material=" & IIf(IsNothing(pf.pile_material), "Null", "'" & pf.pile_material.ToString & "'")
'        updateString += ", pile_length=" & IIf(IsNothing(pf.pile_length), "Null", pf.pile_length.ToString)
'        updateString += ", pile_diameter_width=" & IIf(IsNothing(pf.pile_diameter_width), "Null", pf.pile_diameter_width.ToString)
'        updateString += ", pile_pipe_thickness=" & IIf(IsNothing(pf.pile_pipe_thickness), "Null", pf.pile_pipe_thickness.ToString)
'        updateString += ", pile_soil_capacity_given=" & IIf(IsNothing(pf.pile_soil_capacity_given), "Null", "'" & pf.pile_soil_capacity_given.ToString & "'")
'        updateString += ", steel_yield_strength=" & IIf(IsNothing(pf.steel_yield_strength), "Null", pf.steel_yield_strength.ToString)
'        updateString += ", pile_type_option=" & IIf(IsNothing(pf.pile_type_option), "Null", "'" & pf.pile_type_option.ToString & "'")
'        updateString += ", rebar_quantity=" & IIf(IsNothing(pf.rebar_quantity), "Null", pf.rebar_quantity.ToString)
'        updateString += ", pile_group_config=" & IIf(IsNothing(pf.pile_group_config), "Null", "'" & pf.pile_group_config.ToString & "'")
'        updateString += ", foundation_depth=" & IIf(IsNothing(pf.foundation_depth), "Null", pf.foundation_depth.ToString)
'        updateString += ", pad_thickness=" & IIf(IsNothing(pf.pad_thickness), "Null", pf.pad_thickness.ToString)
'        updateString += ", pad_width_dir1=" & IIf(IsNothing(pf.pad_width_dir1), "Null", pf.pad_width_dir1.ToString)
'        updateString += ", pad_width_dir2=" & IIf(IsNothing(pf.pad_width_dir2), "Null", pf.pad_width_dir2.ToString)
'        updateString += ", pad_rebar_size_bottom=" & IIf(IsNothing(pf.pad_rebar_size_bottom), "Null", pf.pad_rebar_size_bottom.ToString)
'        updateString += ", pad_rebar_size_top=" & IIf(IsNothing(pf.pad_rebar_size_top), "Null", pf.pad_rebar_size_top.ToString)
'        updateString += ", pad_rebar_quantity_bottom_dir1=" & IIf(IsNothing(pf.pad_rebar_quantity_bottom_dir1), "Null", pf.pad_rebar_quantity_bottom_dir1.ToString)
'        updateString += ", pad_rebar_quantity_top_dir1=" & IIf(IsNothing(pf.pad_rebar_quantity_top_dir1), "Null", pf.pad_rebar_quantity_top_dir1.ToString)
'        updateString += ", pad_rebar_quantity_bottom_dir2=" & IIf(IsNothing(pf.pad_rebar_quantity_bottom_dir2), "Null", pf.pad_rebar_quantity_bottom_dir2.ToString)
'        updateString += ", pad_rebar_quantity_top_dir2=" & IIf(IsNothing(pf.pad_rebar_quantity_top_dir2), "Null", pf.pad_rebar_quantity_top_dir2.ToString)
'        updateString += ", pier_shape=" & IIf(IsNothing(pf.pier_shape), "Null", "'" & pf.pier_shape.ToString & "'")
'        updateString += ", pier_diameter=" & IIf(IsNothing(pf.pier_diameter), "Null", pf.pier_diameter.ToString)
'        updateString += ", extension_above_grade=" & IIf(IsNothing(pf.extension_above_grade), "Null", pf.extension_above_grade.ToString)
'        updateString += ", pier_rebar_size=" & IIf(IsNothing(pf.pier_rebar_size), "Null", pf.pier_rebar_size.ToString)
'        updateString += ", pier_rebar_quantity=" & IIf(IsNothing(pf.pier_rebar_quantity), "Null", pf.pier_rebar_quantity.ToString)
'        updateString += ", pier_tie_size=" & IIf(IsNothing(pf.pier_tie_size), "Null", pf.pier_tie_size.ToString)
'        'updateString += ", pier_tie_quantity=" & IIf(IsNothing(pf.pier_tie_quantity), "Null", pf.pier_tie_quantity.ToString)
'        updateString += ", rebar_grade=" & IIf(IsNothing(pf.rebar_grade), "Null", pf.rebar_grade.ToString)
'        updateString += ", concrete_compressive_strength=" & IIf(IsNothing(pf.concrete_compressive_strength), "Null", pf.concrete_compressive_strength.ToString)
'        updateString += ", groundwater_depth=" & IIf(IsNothing(pf.groundwater_depth), "Null", pf.groundwater_depth.ToString)
'        updateString += ", total_soil_unit_weight=" & IIf(IsNothing(pf.total_soil_unit_weight), "Null", pf.total_soil_unit_weight.ToString)
'        updateString += ", cohesion=" & IIf(IsNothing(pf.cohesion), "Null", pf.cohesion.ToString)
'        updateString += ", friction_angle=" & IIf(IsNothing(pf.friction_angle), "Null", pf.friction_angle.ToString)
'        updateString += ", neglect_depth=" & IIf(IsNothing(pf.neglect_depth), "Null", pf.neglect_depth.ToString)
'        updateString += ", spt_blow_count=" & IIf(IsNothing(pf.spt_blow_count), "Null", pf.spt_blow_count.ToString)
'        updateString += ", pile_negative_friction_force=" & IIf(IsNothing(pf.pile_negative_friction_force), "Null", pf.pile_negative_friction_force.ToString)
'        updateString += ", pile_ultimate_compression=" & IIf(IsNothing(pf.pile_ultimate_compression), "Null", pf.pile_ultimate_compression.ToString)
'        updateString += ", pile_ultimate_tension=" & IIf(IsNothing(pf.pile_ultimate_tension), "Null", pf.pile_ultimate_tension.ToString)
'        updateString += ", top_and_bottom_rebar_different=" & IIf(IsNothing(pf.top_and_bottom_rebar_different), "Null", "'" & pf.top_and_bottom_rebar_different.ToString & "'")
'        updateString += ", ultimate_gross_end_bearing=" & IIf(IsNothing(pf.ultimate_gross_end_bearing), "Null", pf.ultimate_gross_end_bearing.ToString)
'        updateString += ", skin_friction_given=" & IIf(IsNothing(pf.skin_friction_given), "Null", "'" & pf.skin_friction_given.ToString & "'")
'        updateString += ", pile_quantity_circular=" & IIf(IsNothing(pf.pile_quantity_circular), "Null", pf.pile_quantity_circular.ToString)
'        updateString += ", group_diameter_circular=" & IIf(IsNothing(pf.group_diameter_circular), "Null", pf.group_diameter_circular.ToString)
'        updateString += ", pile_column_quantity=" & IIf(IsNothing(pf.pile_column_quantity), "Null", pf.pile_column_quantity.ToString)
'        updateString += ", pile_row_quantity=" & IIf(IsNothing(pf.pile_row_quantity), "Null", pf.pile_row_quantity.ToString)
'        updateString += ", pile_columns_spacing=" & IIf(IsNothing(pf.pile_columns_spacing), "Null", pf.pile_columns_spacing.ToString)
'        updateString += ", pile_row_spacing=" & IIf(IsNothing(pf.pile_row_spacing), "Null", pf.pile_row_spacing.ToString)
'        updateString += ", group_efficiency_factor_given=" & IIf(IsNothing(pf.group_efficiency_factor_given), "Null", "'" & pf.group_efficiency_factor_given.ToString & "'")
'        updateString += ", group_efficiency_factor=" & IIf(IsNothing(pf.group_efficiency_factor), "Null", pf.group_efficiency_factor.ToString)
'        updateString += ", cap_type=" & IIf(IsNothing(pf.cap_type), "Null", "'" & pf.cap_type.ToString & "'")
'        updateString += ", pile_quantity_asymmetric=" & IIf(IsNothing(pf.pile_quantity_asymmetric), "Null", pf.pile_quantity_asymmetric.ToString)
'        updateString += ", pile_spacing_min_asymmetric=" & IIf(IsNothing(pf.pile_spacing_min_asymmetric), "Null", pf.pile_spacing_min_asymmetric.ToString)
'        updateString += ", quantity_piles_surrounding=" & IIf(IsNothing(pf.quantity_piles_surrounding), "Null", pf.quantity_piles_surrounding.ToString)
'        updateString += ", pile_cap_reference=" & IIf(IsNothing(pf.pile_cap_reference), "Null", "'" & pf.pile_cap_reference.ToString & "'")
'        updateString += " WHERE ID = " & pf.pile_id.ToString

'        Return updateString

'    End Function

'    Private Function UpdatePileSoilLayer(ByVal pfsl As PileSoilLayer) As String
'        Dim updateString As String = ""

'        updateString += "UPDATE pile_soil_layer SET "
'        'updateString += " soil_layer_id=" & IIf(IsNothing(pfsl.soil_layer_id), "Null", pfsl.soil_layer_id.ToString)
'        updateString += " bottom_depth=" & IIf(IsNothing(pfsl.bottom_depth), "Null", pfsl.bottom_depth.ToString)
'        updateString += ", effective_soil_density=" & IIf(IsNothing(pfsl.effective_soil_density), "Null", pfsl.effective_soil_density.ToString)
'        updateString += ", cohesion=" & IIf(IsNothing(pfsl.cohesion), "Null", pfsl.cohesion.ToString)
'        updateString += ", friction_angle=" & IIf(IsNothing(pfsl.friction_angle), "Null", pfsl.friction_angle.ToString)
'        'updateString += ", skin_friction_override_uplift=" & IIf(IsNothing(pfsl.skin_friction_override_uplift), "Null", pfsl.skin_friction_override_uplift.ToString)
'        updateString += ", spt_blow_count=" & IIf(IsNothing(pfsl.spt_blow_count), "Null", pfsl.spt_blow_count.ToString)
'        updateString += ", ultimate_skin_friction_comp=" & IIf(IsNothing(pfsl.ultimate_skin_friction_comp), "Null", pfsl.ultimate_skin_friction_comp.ToString)
'        updateString += ", ultimate_skin_friction_uplift=" & IIf(IsNothing(pfsl.ultimate_skin_friction_uplift), "Null", pfsl.ultimate_skin_friction_uplift.ToString)
'        updateString += " WHERE ID = " & pfsl.soil_layer_id.ToString

'        Return updateString
'    End Function

'    Private Function UpdatePileLocation(ByVal pfpl As PileLocation) As String
'        Dim updateString As String = ""

'        updateString += "UPDATE pile_location SET "
'        'updateString += " location_id=" & IIf(IsNothing(pfpl.location_id), "Null", pfpl.location_id.ToString)
'        updateString += " pile_x_coordinate=" & IIf(IsNothing(pfpl.pile_x_coordinate), "Null", pfpl.pile_x_coordinate.ToString)
'        updateString += ", pile_y_coordinate=" & IIf(IsNothing(pfpl.pile_y_coordinate), "Null", pfpl.pile_y_coordinate.ToString)
'        updateString += " WHERE ID = " & pfpl.location_id.ToString

'        Return updateString
'    End Function

'#End Region

'#Region "General"
'    Public Sub Clear()
'        ExcelFilePath = ""
'        Piles.Clear()

'        'Remove all datatables from the main dataset
'        For Each item As EXCELDTParameter In PileExcelDTParameters()
'            Try
'                ds.Tables.Remove(item.xlsDatatable)
'            Catch ex As Exception
'            End Try
'        Next

'        For Each item As SQLParameter In PileSQLDataTables()
'            Try
'                ds.Tables.Remove(item.sqlDatatable)
'            Catch ex As Exception
'            End Try
'        Next
'    End Sub

'    Private Function PileSQLDataTables() As List(Of SQLParameter)
'        Dim MyParameters As New List(Of SQLParameter)

'        MyParameters.Add(New SQLParameter("Pile General Details SQL", "Pile (SELECT Details).sql"))
'        MyParameters.Add(New SQLParameter("Pile Soil SQL", "Pile (SELECT Soil Layers).sql"))
'        MyParameters.Add(New SQLParameter("Pile Location SQL", "Pile (SELECT Location).sql"))

'        Return MyParameters
'    End Function

'    Private Function PileExcelDTParameters() As List(Of EXCELDTParameter)
'        Dim MyParameters As New List(Of EXCELDTParameter)

'        MyParameters.Add(New EXCELDTParameter("Pile Soil EXCEL", "A3:H17", "SAPI"))
'        MyParameters.Add(New EXCELDTParameter("Pile Location EXCEL", "S3:U103", "SAPI"))

'        Return MyParameters
'    End Function

'#End Region

'#Region "Check Changes"
'    'Private changeDt As New DataTable
'    'Private changeList As New List(Of AnalysisChanges)
'    Function CheckChanges(ByVal xlPile As Pile, ByVal sqlPile As Pile) As Boolean
'        Dim changesMade As Boolean = False

'        'changeDt.Columns.Add("Variable", Type.GetType("System.String"))
'        'changeDt.Columns.Add("New Value", Type.GetType("System.String"))
'        'changeDt.Columns.Add("Previuos Value", Type.GetType("System.String"))
'        'changeDt.Columns.Add("WO", Type.GetType("System.String"))

'        'Check Details
'        If Check1Change(xlPile.load_eccentricity, sqlPile.load_eccentricity, "Pile", "Load_Eccentricity") Then changesMade = True
'        If Check1Change(xlPile.bolt_circle_bearing_plate_width, sqlPile.bolt_circle_bearing_plate_width, "Pile", "Bolt_Circle_Bearing_Plate_Width") Then changesMade = True
'        If Check1Change(xlPile.pile_shape, sqlPile.pile_shape, "Pile", "Pile_Shape") Then changesMade = True
'        If Check1Change(xlPile.pile_material, sqlPile.pile_material, "Pile", "Pile_Material") Then changesMade = True
'        If Check1Change(xlPile.pile_length, sqlPile.pile_length, "Pile", "Pile_Length") Then changesMade = True
'        If Check1Change(xlPile.pile_diameter_width, sqlPile.pile_diameter_width, "Pile", "Pile_Diameter_Width") Then changesMade = True
'        If Check1Change(xlPile.pile_pipe_thickness, sqlPile.pile_pipe_thickness, "Pile", "Pile_Pipe_Thickness") Then changesMade = True
'        If Check1Change(xlPile.pile_soil_capacity_given, sqlPile.pile_soil_capacity_given, "Pile", "Pile_Soil_Capacity_Given") Then changesMade = True
'        If Check1Change(xlPile.steel_yield_strength, sqlPile.steel_yield_strength, "Pile", "Steel_Yield_Strength") Then changesMade = True
'        If Check1Change(xlPile.pile_type_option, sqlPile.pile_type_option, "Pile", "Pile_Type_Option") Then changesMade = True
'        If Check1Change(xlPile.rebar_quantity, sqlPile.rebar_quantity, "Pile", "Rebar_Quantity") Then changesMade = True
'        If Check1Change(xlPile.pile_group_config, sqlPile.pile_group_config, "Pile", "Pile_Group_Config") Then changesMade = True
'        If Check1Change(xlPile.foundation_depth, sqlPile.foundation_depth, "Pile", "Foundation_Depth") Then changesMade = True
'        If Check1Change(xlPile.pad_thickness, sqlPile.pad_thickness, "Pile", "Pad_Thickness") Then changesMade = True
'        If Check1Change(xlPile.pad_width_dir1, sqlPile.pad_width_dir1, "Pile", "Pad_Width_Dir1") Then changesMade = True
'        If Check1Change(xlPile.pad_width_dir2, sqlPile.pad_width_dir2, "Pile", "Pad_Width_Dir2") Then changesMade = True
'        If Check1Change(xlPile.pad_rebar_size_bottom, sqlPile.pad_rebar_size_bottom, "Pile", "Pad_Rebar_Size_Bottom") Then changesMade = True
'        If Check1Change(xlPile.pad_rebar_size_top, sqlPile.pad_rebar_size_top, "Pile", "Pad_Rebar_Size_Top") Then changesMade = True
'        If Check1Change(xlPile.pad_rebar_quantity_bottom_dir1, sqlPile.pad_rebar_quantity_bottom_dir1, "Pile", "Pad_Rebar_Quantity_Bottom_Dir1") Then changesMade = True
'        If Check1Change(xlPile.pad_rebar_quantity_top_dir1, sqlPile.pad_rebar_quantity_top_dir1, "Pile", "Pad_Rebar_Quantity_Top_Dir1") Then changesMade = True
'        If Check1Change(xlPile.pad_rebar_quantity_bottom_dir2, sqlPile.pad_rebar_quantity_bottom_dir2, "Pile", "Pad_Rebar_Quantity_Bottom_Dir2") Then changesMade = True
'        If Check1Change(xlPile.pad_rebar_quantity_top_dir2, sqlPile.pad_rebar_quantity_top_dir2, "Pile", "Pad_Rebar_Quantity_Top_Dir2") Then changesMade = True
'        If Check1Change(xlPile.pier_shape, sqlPile.pier_shape, "Pile", "Pier_Shape") Then changesMade = True
'        If Check1Change(xlPile.pier_diameter, sqlPile.pier_diameter, "Pile", "Pier_Diameter") Then changesMade = True
'        If Check1Change(xlPile.extension_above_grade, sqlPile.extension_above_grade, "Pile", "Extension_Above_Grade") Then changesMade = True
'        If Check1Change(xlPile.pier_rebar_size, sqlPile.pier_rebar_size, "Pile", "Pier_Rebar_Size") Then changesMade = True
'        If Check1Change(xlPile.pier_rebar_quantity, sqlPile.pier_rebar_quantity, "Pile", "Pier_Rebar_Quantity") Then changesMade = True
'        If Check1Change(xlPile.pier_tie_size, sqlPile.pier_tie_size, "Pile", "Pier_Tie_Size") Then changesMade = True
'        If Check1Change(xlPile.rebar_grade, sqlPile.rebar_grade, "Pile", "Rebar_Grade") Then changesMade = True
'        If Check1Change(xlPile.concrete_compressive_strength, sqlPile.concrete_compressive_strength, "Pile", "Concrete_Compressive_Strength") Then changesMade = True
'        If Check1Change(xlPile.groundwater_depth, sqlPile.groundwater_depth, "Pile", "Groundwater_Depth") Then changesMade = True
'        If Check1Change(xlPile.total_soil_unit_weight, sqlPile.total_soil_unit_weight, "Pile", "Total_Soil_Unit_Weight") Then changesMade = True
'        If Check1Change(xlPile.cohesion, sqlPile.cohesion, "Pile", "Cohesion") Then changesMade = True
'        If Check1Change(xlPile.friction_angle, sqlPile.friction_angle, "Pile", "Friction_Angle") Then changesMade = True
'        If Check1Change(xlPile.neglect_depth, sqlPile.neglect_depth, "Pile", "Neglect_Depth") Then changesMade = True
'        If Check1Change(xlPile.spt_blow_count, sqlPile.spt_blow_count, "Pile", "Spt_Blow_Count") Then changesMade = True
'        If Check1Change(xlPile.pile_negative_friction_force, sqlPile.pile_negative_friction_force, "Pile", "Pile_Negative_Friction_Force") Then changesMade = True
'        If Check1Change(xlPile.pile_ultimate_compression, sqlPile.pile_ultimate_compression, "Pile", "Pile_Ultimate_Compression") Then changesMade = True
'        If Check1Change(xlPile.pile_ultimate_tension, sqlPile.pile_ultimate_tension, "Pile", "Pile_Ultimate_Tension") Then changesMade = True
'        If Check1Change(xlPile.top_and_bottom_rebar_different, sqlPile.top_and_bottom_rebar_different, "Pile", "Top_And_Bottom_Rebar_Different") Then changesMade = True
'        If Check1Change(xlPile.ultimate_gross_end_bearing, sqlPile.ultimate_gross_end_bearing, "Pile", "Ultimate_Gross_End_Bearing") Then changesMade = True
'        If Check1Change(xlPile.skin_friction_given, sqlPile.skin_friction_given, "Pile", "Skin_Friction_Given") Then changesMade = True
'        If Check1Change(xlPile.pile_quantity_circular, sqlPile.pile_quantity_circular, "Pile", "Pile_Quantity_Circular") Then changesMade = True
'        If Check1Change(xlPile.group_diameter_circular, sqlPile.group_diameter_circular, "Pile", "Group_Diameter_Circular") Then changesMade = True
'        If Check1Change(xlPile.pile_column_quantity, sqlPile.pile_column_quantity, "Pile", "Pile_Column_Quantity") Then changesMade = True
'        If Check1Change(xlPile.pile_row_quantity, sqlPile.pile_row_quantity, "Pile", "Pile_Row_Quantity") Then changesMade = True
'        If Check1Change(xlPile.pile_columns_spacing, sqlPile.pile_columns_spacing, "Pile", "Pile_Columns_Spacing") Then changesMade = True
'        If Check1Change(xlPile.pile_row_spacing, sqlPile.pile_row_spacing, "Pile", "Pile_Row_Spacing") Then changesMade = True
'        If Check1Change(xlPile.group_efficiency_factor_given, sqlPile.group_efficiency_factor_given, "Pile", "Group_Efficiency_Factor_Given") Then changesMade = True
'        If Check1Change(xlPile.group_efficiency_factor, sqlPile.group_efficiency_factor, "Pile", "Group_Efficiency_Factor") Then changesMade = True
'        If Check1Change(xlPile.cap_type, sqlPile.cap_type, "Pile", "Cap_Type") Then changesMade = True
'        If Check1Change(xlPile.pile_quantity_asymmetric, sqlPile.pile_quantity_asymmetric, "Pile", "Pile_Quantity_Asymmetric") Then changesMade = True
'        If Check1Change(xlPile.pile_spacing_min_asymmetric, sqlPile.pile_spacing_min_asymmetric, "Pile", "Pile_Spacing_Min_Asymmetric") Then changesMade = True
'        If Check1Change(xlPile.quantity_piles_surrounding, sqlPile.quantity_piles_surrounding, "Pile", "Quantity_Piles_Surrounding") Then changesMade = True
'        If Check1Change(xlPile.pile_cap_reference, sqlPile.pile_cap_reference, "Pile", "Pile_Cap_Reference") Then changesMade = True
'        'If Check1Change(xlPile.tool_version, sqlPile.tool_version, "Pile",  "Tool_Version") Then changesMade = True
'        If Check1Change(xlPile.Soil_110, sqlPile.Soil_110, "Pile", "Soil_110") Then changesMade = True
'        If Check1Change(xlPile.Structural_105, sqlPile.Structural_105, "Pile", "Structural_105") Then changesMade = True


'        'Check Soil Layer
'        'If xlPile.soil_layers.Count <> sqlPile.soil_layers.Count Then changesMade = True 'If want to bypass all the checks below

'        If xlPile.pile_soil_capacity_given = False And xlPile.pile_shape <> "H-Pile" Then
'            For Each psl As PileSoilLayer In xlPile.soil_layers
'                For Each sqlpsl As PileSoilLayer In sqlPile.soil_layers
'                    If psl.soil_layer_id = sqlpsl.soil_layer_id Then

'                        If Check1Change(psl.bottom_depth, sqlpsl.bottom_depth, "Pile", "Bottom_Depth" & psl.soil_layer_id.ToString) Then changesMade = True
'                        If Check1Change(psl.effective_soil_density, sqlpsl.effective_soil_density, "Pile", "Effective_Soil_Density" & psl.soil_layer_id.ToString) Then changesMade = True
'                        If Check1Change(psl.cohesion, sqlpsl.cohesion, "Pile", "Cohesion" & psl.soil_layer_id.ToString) Then changesMade = True
'                        If Check1Change(psl.friction_angle, sqlpsl.friction_angle, "Pile", "Friction_Angle" & psl.soil_layer_id.ToString) Then changesMade = True
'                        If Check1Change(psl.spt_blow_count, sqlpsl.spt_blow_count, "Pile", "Spt_Blow_Count" & psl.soil_layer_id.ToString) Then changesMade = True
'                        If Check1Change(psl.ultimate_skin_friction_comp, sqlpsl.ultimate_skin_friction_comp, "Pile", "Ultimate_Skin_Friction_Comp" & psl.soil_layer_id.ToString) Then changesMade = True
'                        If Check1Change(psl.ultimate_skin_friction_uplift, sqlpsl.ultimate_skin_friction_uplift, "Pile", "Ultimate_Skin_Friction_Uplift" & psl.soil_layer_id.ToString) Then changesMade = True

'                        Exit For
'                    End If
'                    If psl.soil_layer_id = 0 Then 'accounts for inserting new rows. additional rows won't have an ID associated to them. 

'                        If Check1Change(psl.bottom_depth, Nothing, "Pile", "Bottom_Depth" & psl.soil_layer_id.ToString) Then changesMade = True
'                        If Check1Change(psl.effective_soil_density, Nothing, "Pile", "Effective_Soil_Density" & psl.soil_layer_id.ToString) Then changesMade = True
'                        If Check1Change(psl.cohesion, Nothing, "Pile", "Cohesion" & psl.soil_layer_id.ToString) Then changesMade = True
'                        If Check1Change(psl.friction_angle, Nothing, "Pile", "Friction_Angle" & psl.soil_layer_id.ToString) Then changesMade = True
'                        If Check1Change(psl.spt_blow_count, Nothing, "Pile", "Spt_Blow_Count" & psl.soil_layer_id.ToString) Then changesMade = True
'                        If Check1Change(psl.ultimate_skin_friction_comp, Nothing, "Pile", "Ultimate_Skin_Friction_Comp" & psl.soil_layer_id.ToString) Then changesMade = True
'                        If Check1Change(psl.ultimate_skin_friction_uplift, Nothing, "Pile", "Ultimate_Skin_Friction_Uplift" & psl.soil_layer_id.ToString) Then changesMade = True

'                        Exit For
'                    End If
'                Next
'            Next

'        End If

'        'Pile Location
'        If xlPile.pile_group_config = "Asymmetric" Then
'            For Each pfpl As PileLocation In xlPile.pile_locations
'                For Each sqlpfpl As PileLocation In sqlPile.pile_locations
'                    If pfpl.location_id = sqlpfpl.location_id Then

'                        If Check1Change(pfpl.pile_x_coordinate, sqlpfpl.pile_x_coordinate, "Pile", "Pile_X_Coordinate" & pfpl.location_id.ToString) Then changesMade = True
'                        If Check1Change(pfpl.pile_y_coordinate, sqlpfpl.pile_y_coordinate, "Pile", "Pile_Y_Coordinate" & pfpl.location_id.ToString) Then changesMade = True

'                        Exit For
'                    End If
'                    If pfpl.location_id = 0 Then 'accounts for inserting new rows. additional rows won't have an ID associated to them.

'                        If Check1Change(pfpl.pile_x_coordinate, Nothing, "Pile", "Pile_X_Coordinate" & pfpl.location_id.ToString) Then changesMade = True
'                        If Check1Change(pfpl.pile_y_coordinate, Nothing, "Pile", "Pile_Y_Coordinate" & pfpl.location_id.ToString) Then changesMade = True

'                        Exit For
'                    End If
'                Next
'            Next

'        End If
'        CreateChangeSummary(changeDt) 'possible alternative to listing change summary
'        Return changesMade
'    End Function

'    'Function CreateChangeSummary(ByVal changeDt As DataTable) As String
'    '    'Sub CreateChangeSummary(ByVal changeDt As DataTable)
'    '    'Create your string based on data in the datatable
'    '    Dim summary As String
'    '    Dim counter As Integer = 0

'    '    For Each chng As AnalysisChanges In changeList
'    '        If counter = 0 Then
'    '            summary += chng.Name & " = " & chng.NewValue & " | Previously: " & chng.PreviousValue
'    '        Else
'    '            summary += vbNewLine & chng.Name & " = " & chng.NewValue & " | Previously: " & chng.PreviousValue
'    '        End If

'    '        counter += 1
'    '    Next

'    '    'write to text file
'    '    'End Sub
'    'End Function

'    'Function Check1Change(ByVal newValue As Object, ByVal oldvalue As Object, ByVal tolerance As Double, ByVal variable As String) As Boolean
'    '    If newValue <> oldvalue Then
'    '        changeDt.Rows.Add(variable, newValue, oldvalue, CurWO) 'Need to determine what we want to store in this datatable or list (Foundation Type, Foundation ID)?
'    '        changeList.Add(New AnalysisChanges(oldvalue, newValue, variable, "Pile Foundations"))
'    '        Return True
'    '    ElseIf Not IsNothing(newValue) And IsNothing(oldvalue) Then 'accounts for when new rows are added. New rows from excel=0 where sql=nothing
'    '        changeDt.Rows.Add(variable, newValue, oldvalue, CurWO) 'Need to determine what we want to store in this datatable or list (Foundation Type, Foundation ID)?
'    '        changeList.Add(New AnalysisChanges(oldvalue, newValue, variable, "Pile Foundations"))
'    '        Return True
'    '    ElseIf IsNothing(newValue) And Not IsNothing(oldvalue) Then 'accounts for when rows are removed. Rows from excel=nothing where sql=value
'    '        changeDt.Rows.Add(variable, newValue, oldvalue, CurWO) 'Need to determine what we want to store in this datatable or list (Foundation Type, Foundation ID)?
'    '        changeList.Add(New AnalysisChanges(oldvalue, newValue, variable, "Pile Foundations"))
'    '        Return True
'    '    End If
'    'End Function
'#End Region

'End Class


''Class AnalysisChanges
''    Property PreviousValue As String
''    Property NewValue As String
''    Property Name As String
''    Property PartofDatabase As String

''    Public Sub New(prev As String, Newval As String, name As String, db As String)
''        Me.PreviousValue = prev
''        Me.NewValue = Newval
''        Me.Name = name
''        Me.PartofDatabase = db
''    End Sub
''End Class