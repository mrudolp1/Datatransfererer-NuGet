Option Strict Off

Imports DevExpress.Spreadsheet
Imports CCI_Engineering_Templates
Imports System.Security.Principal

Partial Public Class DataTransfererDrilledPier

#Region "Define"
    Private NewDrilledPierWb As New Workbook
    Private prop_ExcelFilePath As String

    Public Property DrilledPiers As New List(Of DrilledPier)
    Private Property DrilledPierTemplatePath As String = "C:\Users\imiller\source\repos\DevExpress Objects\Drilled Pier Foundation (4.2.3).xlsm"
    Private Property DrilledPierFileType As DocumentFormat = DocumentFormat.Xlsm

    Public Property dpDS As DataSet
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
#End Region

#Region "Constructors"
    Public Sub New()
        'Leave method empty
    End Sub

    Public Sub New(ByVal MyDataSet As DataSet, ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String, ByVal BU As String, ByVal Strucutre_ID As String)
        dpDS = MyDataSet
        dpID = LogOnUser
        dpDB = ActiveDatabase
        BUNumber = BU
        STR_ID = Strucutre_ID
    End Sub
#End Region

#Region "Load Data"
    Public Sub LoadFromEDS()
        Dim refid As Integer

        Dim DrilledPierLoader As String

        'Load data to get pier and pad details data for the existing structure model
        For Each item As SQLParameter In DrilledPierSQLDataTables()
            DrilledPierLoader = QueryBuilderFromFile(queryPath & "Drilled Pier\" & item.sqlQuery).Replace("[EXISTING MODEL]", GetExistingModelQuery())
            DoDaSQL.sqlLoader(DrilledPierLoader, item.sqlDatatable, dpDS, dpDB, dpID, "0")
        Next

        'Custom Section to transfer data for the drilled pier tool. Needs to be adjusted for each tool.
        For Each DrilledPierDataRow As DataRow In dpDS.Tables("Drilled Pier General Details SQL").Rows
            refid = CType(DrilledPierDataRow.Item("drilled_pier_id"), Integer)

            DrilledPiers.Add(New DrilledPier(DrilledPierDataRow, refid))
        Next

    End Sub 'Create Drilled Pier objects based on what is saved in EDS

    Public Sub LoadFromExcel()
        Dim refID As Integer
        Dim refCol As String

        For Each item As EXCELDTParameter In DrilledPierExcelDTParameters()
            'Get tables from excel file 
            ds.Tables.Add(ExcelDatasourceToDataTable(GetExcelDataSource(ExcelFilePath, item.xlsSheet, item.xlsRange), item.xlsDatatable))
        Next

        'Custom Section to transfer data for the drilled pier tool. Needs to be adjusted for each tool.
        For Each DrilledPierDataRow As DataRow In ds.Tables("Drilled Pier General Details EXCEL").Rows
            If DrilledPierDataRow.Item("foudation_id").ToString = "" Then
                refCol = "local_drilled_pier_id"
                refID = CType(DrilledPierDataRow.Item(refCol), Integer)
            Else
                refCol = "drilled_pier_id"
                refID = CType(DrilledPierDataRow.Item(refCol), Integer)
            End If

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
        Dim myEmbedSections As String = ""

        For Each dp As DrilledPier In DrilledPiers
            Dim DrilledPierSaver As String = QueryBuilderFromFile(queryPath & "Drilled Pier\Drilled Piers (IN_UP).sql")
            Dim dpSectionQuery As String = QueryBuilderFromFile(queryPath & "Drilled Pier\Drilled Piers Sections (IN_UP).txt")

            DrilledPierSaver = DrilledPierSaver.Replace("[BU NUMBER]", BUNumber)
            DrilledPierSaver = DrilledPierSaver.Replace("[STRUCTURE ID]", STR_ID)
            DrilledPierSaver = DrilledPierSaver.Replace("[FOUNDATION TYPE]", "Drilled Pier")
            If dp.pier_id = 0 Or IsDBNull(dp.pier_id) Then
                DrilledPierSaver = DrilledPierSaver.Replace("[DRILLED PIER ID]", "NULL")
            Else
                DrilledPierSaver = DrilledPierSaver.Replace("[DRILLED PIER ID]", dp.pier_id.ToString)
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
                End If 'Add Belled Pier INSERT Statment

                If dp.embedded_pole Then
                    DrilledPierSaver = DrilledPierSaver.Replace("[INSERT ALL EMBEDDED POLE DETAILS]", InsertDrilledPierEmbed(dp.embed_details))

                    For Each eSec As DrilledPierEmbedSection In dp.embed_details.sections
                        Dim tempEmbedSection As String = InsertDrilledPierEmbedSection(eSec)

                        If Not firstOne Then
                            myEmbedSections += ",(" & tempEmbedSection & ")"
                        Else
                            myEmbedSections += "(" & tempEmbedSection & ")"
                        End If

                        firstOne = False
                    Next
                    DrilledPierSaver = DrilledPierSaver.Replace("[INSERT ALL EMBEDDED SECTIONS]", myEmbedSections)
                End If 'Add Embedded Pole INSERT Statment

                mySoils = ""
                mySections = ""
                myEmbedSections = ""
            Else
                Dim tempUpdater As String = ""
                tempUpdater += UpdateDrilledPierDetail(dp)

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
                End If

                If dp.embedded_pole Then
                    If dp.embed_details.embedded_id = 0 Or IsDBNull(dp.embed_details.embedded_id) Then
                        tempUpdater += "BEGIN INSERT INTO embedded_pole_details OUTPUT INSERTED.ID INTO @EmbeddedPole VALUES (" & InsertDrilledPierEmbed(dp.embed_details) & ") " & vbNewLine & " SELECT @EmbedID=EmbedID FROM @EmbeddedPole"
                        For Each eSec As DrilledPierEmbedSection In dp.embed_details.sections
                            tempUpdater += "INSERT INTO embedded_pole_section VALUES (" & InsertDrilledPierEmbedSection(eSec) & ") " & vbNewLine
                        Next
                        tempUpdater += " END " & vbNewLine
                    Else
                        tempUpdater += UpdateDrilledPierEmbed(dp.embed_details)
                        For Each esec As DrilledPierEmbedSection In dp.embed_details.sections
                            If esec.section_id = 0 Or IsDBNull(esec.section_id) Then
                                tempUpdater += "INSERT INTO embedded_pole_section VALUES (" & InsertDrilledPierEmbedSection(esec).Replace("@EmbedID", dp.embed_details.embedded_id.ToString) & ") " & vbNewLine
                            Else
                                tempUpdater += UpdateDrilledPierEmbedSection(esec)
                            End If
                        Next
                    End If
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

                DrilledPierSaver = DrilledPierSaver.Replace("SELECT * FROM TEMPORARY", tempUpdater)
            End If

            'sqlSender(DrilledPierSaver, 0)
        Next


    End Sub

    Public Sub SaveToExcel()
        Dim dpRow As Integer = 3
        Dim secRow As Integer = 3
        Dim rebRow As Integer = 3
        Dim soilRow As Integer = 3
        Dim embedSecRow As Integer = 3

        LoadNewDrilledPier()

        With NewDrilledPierWb
            For Each dp As DrilledPier In DrilledPiers
                .Worksheets("Details (SAPI)").Range("B" & dpRow).Value = dp.pier_id
                .Worksheets("Details (SAPI)").Range("C" & dpRow).Value = dp.foundation_depth
                .Worksheets("Details (SAPI)").Range("D" & dpRow).Value = dp.extension_above_grade
                .Worksheets("Details (SAPI)").Range("E" & dpRow).Value = dp.groundwater_depth
                .Worksheets("Details (SAPI)").Range("F" & dpRow).Value = dp.assume_min_steel
                .Worksheets("Details (SAPI)").Range("G" & dpRow).Value = dp.check_shear_along_depth
                .Worksheets("Details (SAPI)").Range("H" & dpRow).Value = dp.utilize_skin_friction_methodology
                .Worksheets("Details (SAPI)").Range("I" & dpRow).Value = dp.embedded_pole
                .Worksheets("Details (SAPI)").Range("J" & dpRow).Value = dp.belled_pier
                .Worksheets("Details (SAPI)").Range("K" & dpRow).Value = dp.soil_layer_quantity

                For Each dpSec As DrilledPierSection In dp.sections
                    .Worksheets("Sections (SAPI)").Range("C" & secRow).Value = dp.pier_id
                    .Worksheets("Sections (SAPI)").Range("D" & secRow).Value = dpSec.section_id
                    .Worksheets("Sections (SAPI)").Range("E" & secRow).Value = dpSec.pier_diameter
                    .Worksheets("Sections (SAPI)").Range("F" & secRow).Value = dpSec.clear_cover
                    .Worksheets("Sections (SAPI)").Range("G" & secRow).Value = dpSec.clear_cover_rebar_cage_option
                    .Worksheets("Sections (SAPI)").Range("H" & secRow).Value = dpSec.tie_size
                    .Worksheets("Sections (SAPI)").Range("I" & secRow).Value = dpSec.tie_spacing
                    .Worksheets("Sections (SAPI)").Range("J" & secRow).Value = dpSec.top_elevation
                    .Worksheets("Sections (SAPI)").Range("K" & secRow).Value = dpSec.bottom_elevation
                    .Worksheets("Sections (SAPI)").Range("L" & secRow).Value = dpSec.tie_yield_strength
                    .Worksheets("Sections (SAPI)").Range("M" & secRow).Value = dpSec.concrete_compressive_strength
                    .Worksheets("Sections (SAPI)").Range("N" & secRow).Value = dpSec.assum_min_steel_rho_override

                    For Each dpReb As DrilledPierRebar In dpSec.rebar
                        .Worksheets("Rebar (SAPI)").Range("C" & rebRow).Value = dp.pier_id
                        .Worksheets("Rebar (SAPI)").Range("D" & rebRow).Value = dpSec.section_id
                        .Worksheets("Rebar (SAPI)").Range("E" & rebRow).Value = dpReb.rebar_id
                        .Worksheets("Rebar (SAPI)").Range("F" & rebRow).Value = dpReb.longitudinal_rebar_quantity
                        .Worksheets("Rebar (SAPI)").Range("G" & rebRow).Value = dpReb.longitudinal_rebar_size
                        .Worksheets("Rebar (SAPI)").Range("H" & rebRow).Value = dpReb.longitudinal_rebar_cage_diameter
                        .Worksheets("Rebar (SAPI)").Range("I" & rebRow).Value = dpReb.longitudinal_rebar_yield_strength

                        rebRow += 1
                    Next

                    secRow += 1
                Next

                For Each dpSL As DrilledPierSoilLayer In dp.soil_layers
                    .Worksheets("Soil Layers (SAPI)").Range("B" & soilRow).Value = dp.pier_id
                    .Worksheets("Soil Layers (SAPI)").Range("C" & soilRow).Value = dpSL.soil_layer_id
                    .Worksheets("Soil Layers (SAPI)").Range("D" & soilRow).Value = dpSL.bottom_depth
                    .Worksheets("Soil Layers (SAPI)").Range("E" & soilRow).Value = dpSL.effective_soil_density
                    .Worksheets("Soil Layers (SAPI)").Range("F" & soilRow).Value = dpSL.cohesion
                    .Worksheets("Soil Layers (SAPI)").Range("G" & soilRow).Value = dpSL.friction_angle
                    .Worksheets("Soil Layers (SAPI)").Range("H" & soilRow).Value = dpSL.skin_friction_override_comp
                    .Worksheets("Soil Layers (SAPI)").Range("I" & soilRow).Value = dpSL.skin_friction_override_uplift
                    .Worksheets("Soil Layers (SAPI)").Range("J" & soilRow).Value = dpSL.bearing_type_toggle
                    .Worksheets("Soil Layers (SAPI)").Range("K" & soilRow).Value = dpSL.nominal_bearing_capacity
                    .Worksheets("Soil Layers (SAPI)").Range("L" & soilRow).Value = dpSL.spt_blow_count

                    soilRow += 1
                Next

                If ds.Tables("Belled Details SQL").Rows.Count > 0 Then
                    .Worksheets("Belled (SAPI)").Range("B" & dpRow).Value = dp.pier_id
                    .Worksheets("Belled (SAPI)").Range("C" & dpRow).Value = dp.belled_details.belled_pier_id
                    .Worksheets("Belled (SAPI)").Range("D" & dpRow).Value = dp.belled_details.belled_pier_option
                    .Worksheets("Belled (SAPI)").Range("E" & dpRow).Value = dp.belled_details.bottom_diameter_of_bell
                    .Worksheets("Belled (SAPI)").Range("F" & dpRow).Value = dp.belled_details.bell_input_type
                    .Worksheets("Belled (SAPI)").Range("G" & dpRow).Value = dp.belled_details.bell_angle
                    .Worksheets("Belled (SAPI)").Range("H" & dpRow).Value = dp.belled_details.bell_height
                    .Worksheets("Belled (SAPI)").Range("I" & dpRow).Value = dp.belled_details.bell_toe_height
                    .Worksheets("Belled (SAPI)").Range("J" & dpRow).Value = dp.belled_details.neglect_top_soil_layer
                    .Worksheets("Belled (SAPI)").Range("K" & dpRow).Value = dp.belled_details.swelling_expansive_soil
                    .Worksheets("Belled (SAPI)").Range("L" & dpRow).Value = dp.belled_details.depth_of_expansive_soil
                    .Worksheets("Belled (SAPI)").Range("M" & dpRow).Value = dp.belled_details.expansive_soil_force
                End If

                If ds.Tables("Embedded Details SQL").Rows.Count > 0 Then
                    .Worksheets("Embedded (SAPI)").Range("B" & dpRow).Value = dp.pier_id
                    .Worksheets("Embedded (SAPI)").Range("C" & dpRow).Value = dp.embed_details.embedded_id
                    .Worksheets("Embedded (SAPI)").Range("D" & dpRow).Value = dp.embed_details.embedded_pole_option
                    .Worksheets("Embedded (SAPI)").Range("E" & dpRow).Value = dp.embed_details.encased_in_concrete
                    .Worksheets("Embedded (SAPI)").Range("F" & dpRow).Value = dp.embed_details.pole_side_quantity
                    .Worksheets("Embedded (SAPI)").Range("G" & dpRow).Value = dp.embed_details.pole_yield_strength
                    .Worksheets("Embedded (SAPI)").Range("H" & dpRow).Value = dp.embed_details.pole_thickness
                    .Worksheets("Embedded (SAPI)").Range("I" & dpRow).Value = dp.embed_details.embedded_pole_input_type
                    .Worksheets("Embedded (SAPI)").Range("J" & dpRow).Value = dp.embed_details.pole_diameter_toc
                    .Worksheets("Embedded (SAPI)").Range("K" & dpRow).Value = dp.embed_details.pole_top_diameter
                    .Worksheets("Embedded (SAPI)").Range("L" & dpRow).Value = dp.embed_details.pole_bottom_diameter
                    .Worksheets("Embedded (SAPI)").Range("M" & dpRow).Value = dp.embed_details.pole_section_length
                    .Worksheets("Embedded (SAPI)").Range("N" & dpRow).Value = dp.embed_details.pole_taper_factor
                    .Worksheets("Embedded (SAPI)").Range("O" & dpRow).Value = dp.embed_details.pole_bend_radius_override

                    For Each eSec As DrilledPierEmbedSection In dp.embed_details.sections
                        .Worksheets("Embedded (SAPI)").Range("B" & dpRow).Value = dp.pier_id
                        'Section ID
                        .Worksheets("Embedded (SAPI)").Range("D" & embedSecRow).Value = eSec.pier_diameter

                        embedSecRow += 1
                    Next
                End If

                dpRow += 1
            Next
        End With

        SaveAndCloseDrilledPier()
    End Sub

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
        insertString += "," & dp.foundation_depth.ToString
        insertString += "," & dp.extension_above_grade.ToString
        insertString += "," & dp.groundwater_depth.ToString
        insertString += "," & "'" & dp.assume_min_steel.ToString & "'"
        insertString += "," & "'" & dp.check_shear_along_depth.ToString & "'"
        insertString += "," & "'" & dp.utilize_skin_friction_methodology.ToString & "'"
        insertString += "," & "'" & dp.embedded_pole.ToString & "'"
        insertString += "," & "'" & dp.belled_pier.ToString & "'"
        insertString += "," & dp.soil_layer_quantity

        Return insertString
    End Function

    Private Function InsertDrilledPierBell(ByVal bp As DrilledPierBelledPier) As String
        Dim insertString As String = ""

        insertString += "@DpID"
        insertString += "," & "'" & bp.belled_pier_option.ToString & "'"
        insertString += "," & bp.bottom_diameter_of_bell.ToString
        insertString += "," & "'" & bp.bell_input_type.ToString & "'"
        insertString += "," & bp.bell_angle.ToString
        insertString += "," & bp.bell_height.ToString
        insertString += "," & bp.bell_toe_height.ToString
        insertString += "," & "'" & bp.neglect_top_soil_layer.ToString & "'"
        insertString += "," & "'" & bp.swelling_expansive_soil.ToString & "'"
        insertString += "," & bp.depth_of_expansive_soil.ToString
        insertString += "," & bp.expansive_soil_force.ToString

        Return insertString
    End Function

    Private Function InsertDrilledPierEmbed(ByVal ep As DrilledPierEmbeddedPier) As String
        Dim insertString As String = ""

        insertString += "@DpID"
        insertString += "," & "'" & ep.embedded_pole_option.ToString & "'"
        insertString += "," & "'" & ep.encased_in_concrete.ToString & "'"
        insertString += "," & ep.pole_side_quantity.ToString
        insertString += "," & ep.pole_yield_strength.ToString
        insertString += "," & ep.pole_thickness.ToString
        insertString += "," & "'" & ep.embedded_pole_input_type.ToString & "'"
        insertString += "," & ep.pole_diameter_toc.ToString
        insertString += "," & ep.pole_top_diameter.ToString
        insertString += "," & ep.pole_bottom_diameter.ToString
        insertString += "," & ep.pole_section_length.ToString
        insertString += "," & ep.pole_taper_factor.ToString
        insertString += "," & ep.pole_bend_radius_override.ToString

        Return insertString
    End Function

    Private Function InsertDrilledPierSoilLayer(ByVal dpsl As DrilledPierSoilLayer) As String
        Dim insertString As String = ""

        insertString += "@DpID"
        insertString += "," & dpsl.bottom_depth.ToString
        insertString += "," & dpsl.effective_soil_density.ToString
        insertString += "," & dpsl.cohesion.ToString
        insertString += "," & dpsl.friction_angle.ToString
        insertString += "," & dpsl.skin_friction_override_comp.ToString
        insertString += "," & dpsl.skin_friction_override_uplift.ToString
        insertString += "," & "'" & dpsl.bearing_type_toggle.ToString & "'"
        insertString += "," & dpsl.nominal_bearing_capacity.ToString
        insertString += "," & dpsl.spt_blow_count.ToString

        Return insertString
    End Function

    Private Function InsertDrilledPierSection(ByVal dpsec As DrilledPierSection) As String
        Dim insertString As String = ""

        insertString += "@DpID"
        insertString += "," & dpsec.pier_diameter.ToString
        insertString += "," & dpsec.clear_cover.ToString
        insertString += "," & "'" & dpsec.clear_cover_rebar_cage_option.ToString & "'"
        insertString += "," & dpsec.tie_size.ToString
        insertString += "," & dpsec.tie_spacing.ToString
        insertString += "," & dpsec.top_elevation.ToString
        insertString += "," & dpsec.bottom_elevation.ToString
        insertString += "," & dpsec.tie_yield_strength.ToString
        insertString += "," & dpsec.concrete_compressive_strength.ToString
        insertString += "," & dpsec.assum_min_steel_rho_override.ToString

        Return insertString
    End Function

    Private Function InsertDrilledPierRebar(ByVal dpreb As DrilledPierRebar) As String
        Dim insertString As String = ""

        insertString += "@SecID"
        insertString += "," & dpreb.longitudinal_rebar_quantity.ToString
        insertString += "," & dpreb.longitudinal_rebar_size.ToString
        insertString += "," & dpreb.longitudinal_rebar_cage_diameter.ToString
        insertString += "," & dpreb.longitudinal_rebar_yield_strength.ToString

        Return insertString
    End Function

    Private Function InsertDrilledPierEmbedSection(ByVal eSec As DrilledPierEmbedSection) As String
        Dim insertString As String = ""

        insertString += "@EmbedID"
        insertString += "," & "'" & eSec.pier_diameter.ToString & "'"

        Return insertString
    End Function
#End Region

#Region "SQL Update Statements"
    Private Function UpdateDrilledPierDetail(ByVal dp As DrilledPier) As String
        Dim updateString As String = ""

        updateString += "UPDATE drilled_pier_details SET "
        updateString += "foundation_depth=" & dp.foundation_depth.ToString
        updateString += ",extension_above_grade=" & dp.extension_above_grade.ToString
        updateString += ",groundwater_depth=" & dp.groundwater_depth.ToString
        updateString += ",assume_min_steel=" & "'" & dp.assume_min_steel.ToString & "'"
        updateString += ",check_shear_along_depth=" & "'" & dp.check_shear_along_depth.ToString & "'"
        updateString += ",utilize_skin_friction_methodology=" & "'" & dp.utilize_skin_friction_methodology.ToString & "'"
        updateString += ",embedded_pole=" & "'" & dp.embedded_pole.ToString & "'"
        updateString += ",belled_pier=" & "'" & dp.belled_pier.ToString & "'"
        updateString += ",soil_layer_quantity=" & dp.soil_layer_quantity.ToString
        updateString += " WHERE ID=" & dp.pier_id & vbNewLine

        Return updateString
    End Function

    Private Function UpdateDrilledPierBell(ByVal bp As DrilledPierBelledPier) As String
        Dim updateString As String = ""

        updateString += "UPDATE belled_pier_details SET "
        updateString += "belled_pier_option=" & "'" & bp.belled_pier_option.ToString & "'"
        updateString += ",bottom_diameter_of_bell=" & bp.bottom_diameter_of_bell.ToString
        updateString += ",bell_input_type=" & "'" & bp.bell_input_type.ToString & "'"
        updateString += ",bell_angle=" & bp.bell_angle.ToString
        updateString += ",bell_height=" & bp.bell_height.ToString
        updateString += ",bell_toe_height=" & bp.bell_toe_height.ToString
        updateString += ",neglect_top_soil_layer=" & "'" & bp.neglect_top_soil_layer.ToString & "'"
        updateString += ",swelling_expansive_soil=" & "'" & bp.swelling_expansive_soil.ToString & "'"
        updateString += ",depth_of_expansive_soil=" & bp.depth_of_expansive_soil.ToString
        updateString += ",expansive_soil_force=" & bp.expansive_soil_force.ToString
        updateString += " WHERE ID=" & bp.belled_pier_id & vbNewLine

        Return updateString
    End Function

    Private Function UpdateDrilledPierEmbed(ByVal ep As DrilledPierEmbeddedPier) As String
        Dim updateString As String = ""

        updateString += "UPDATE embedded_pole_details SET "
        updateString += "embedded_pole_option=" & "'" & ep.embedded_pole_option.ToString & "'"
        updateString += ",encased_in_concrete=" & "'" & ep.encased_in_concrete.ToString & "'"
        updateString += ",pole_side_quantity=" & ep.pole_side_quantity.ToString
        updateString += ",pole_yield_strength=" & ep.pole_yield_strength.ToString
        updateString += ",pole_thickness=" & ep.pole_thickness.ToString
        updateString += ",embedded_pole_input_type=" & "'" & ep.embedded_pole_input_type.ToString & "'"
        updateString += ",pole_diameter_toc=" & ep.pole_diameter_toc.ToString
        updateString += ",pole_top_diameter=" & ep.pole_top_diameter.ToString
        updateString += ",pole_bottom_diameter=" & ep.pole_bottom_diameter.ToString
        updateString += ",pole_section_length=" & ep.pole_section_length.ToString
        updateString += ",pole_taper_factor=" & ep.pole_taper_factor.ToString
        updateString += ",pole_bend_radius_override=" & ep.pole_bend_radius_override.ToString
        updateString += " WHERE ID=" & ep.embedded_id & vbNewLine

        Return updateString
    End Function

    Private Function UpdateDrilledPierSoilLayer(ByVal dpsl As DrilledPierSoilLayer) As String
        Dim updateString As String = ""

        updateString += "UPDATE drilled_pier_soil_layer SET "
        updateString += ",bottom_depth=" & dpsl.bottom_depth.ToString
        updateString += ",effective_soil_density=" & dpsl.effective_soil_density.ToString
        updateString += ",cohesion=" & dpsl.cohesion.ToString
        updateString += ",friction_angle=" & dpsl.friction_angle.ToString
        updateString += ",skin_friction_override_comp=" & dpsl.skin_friction_override_comp.ToString
        updateString += ",skin_friction_override_uplift=" & dpsl.skin_friction_override_uplift.ToString
        updateString += ",bearing_type_toggle=" & "'" & dpsl.bearing_type_toggle.ToString & "'"
        updateString += ",nominal_bearing_capacity=" & dpsl.nominal_bearing_capacity.ToString
        updateString += ",spt_blow_count=" & dpsl.spt_blow_count.ToString
        updateString += " WHERE ID=" & dpsl.soil_layer_id & vbNewLine

        Return updateString
    End Function

    Private Function UpdateDrilledPierSection(ByVal dpsec As DrilledPierSection) As String
        Dim updateString As String = ""

        updateString += "UPDATE drilled_pier_section SET "
        updateString += "pier_diameter=" & dpsec.pier_diameter.ToString
        updateString += ",clear_cover=" & dpsec.clear_cover.ToString
        updateString += ",clear_cover_rebar_cage_option=" & "'" & dpsec.clear_cover_rebar_cage_option.ToString & "'"
        updateString += ",tie_size=" & dpsec.tie_size.ToString
        updateString += ",tie_spacing=" & dpsec.tie_spacing.ToString
        updateString += ",top_elevation=" & dpsec.top_elevation.ToString
        updateString += ",bottom_elevation=" & dpsec.bottom_elevation.ToString
        updateString += ",tie_yield_strength=" & dpsec.tie_yield_strength.ToString
        updateString += ",concrete_compressive_strength=" & dpsec.concrete_compressive_strength.ToString
        updateString += ",assum_min_steel_rho_override=" & dpsec.assum_min_steel_rho_override.ToString
        updateString += " WHERE ID=" & dpsec.section_id & vbNewLine

        Return updateString
    End Function

    Private Function UpdateDrilledPierRebar(ByVal dpreb As DrilledPierRebar) As String
        Dim updateString As String = ""

        updateString += "UPDATE drilled_pier_rebar SET "
        updateString += "longitudinal_rebar_quantity=" & dpreb.longitudinal_rebar_quantity.ToString
        updateString += ",longitudinal_rebar_size=" & dpreb.longitudinal_rebar_size.ToString
        updateString += ",longitudinal_rebar_cage_diameter=" & dpreb.longitudinal_rebar_cage_diameter.ToString
        updateString += ",longitudinal_rebar_yield_strength=" & dpreb.longitudinal_rebar_yield_strength.ToString
        updateString += " WHERE ID=" & dpreb.rebar_id & vbNewLine

        Return updateString
    End Function

    Private Function UpdateDrilledPierEmbedSection(ByVal eSec As DrilledPierEmbedSection) As String
        Dim updateString As String = ""

        updateString += "UPDATE embedded_pole_section SET "
        updateString += "pier_diameter=" & "'" & eSec.pier_diameter.ToString & "'"
        updateString += " WHERE ID=" & eSec.section_id & vbNewLine

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
        MyParameters.Add(New SQLParameter("Embedded Section SQL", "Drilled Piers (SELECT Embedded Section).sql"))

        Return MyParameters
    End Function

    Private Function DrilledPierExcelDTParameters() As List(Of EXCELDTParameter)
        Dim MyParameters As New List(Of EXCELDTParameter)

        MyParameters.Add(New EXCELDTParameter("Drilled Pier General Details EXCEL", "A2:K1000", "Details (SAPI)"))
        MyParameters.Add(New EXCELDTParameter("Drilled Pier Section EXCEL", "A2:N1000", "Sections (SAPI)"))
        MyParameters.Add(New EXCELDTParameter("Drilled Pier Rebar EXCEL", "A2:I1000", "Rebar (SAPI)"))
        MyParameters.Add(New EXCELDTParameter("Drilled Pier Soil EXCEL", "A2:L1000", "Soil Layers (SAPI)"))
        MyParameters.Add(New EXCELDTParameter("Belled Details EXCEL", "A2:M1000", "Belled (SAPI)"))
        MyParameters.Add(New EXCELDTParameter("Embedded Details EXCEL", "A2:O1000", "Embedded (SAPI)"))
        MyParameters.Add(New EXCELDTParameter("Embedded Section EXCEL", "A2:D1000", "Embedded Section (SAPI)"))

        Return MyParameters
    End Function
#End Region

End Class